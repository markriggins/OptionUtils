/**
 * ImportEtrade.js
 * Imports E*Trade transaction history into the Legs table.
 *
 * Features:
 * - Parses E*Trade transaction CSV
 * - Pairs consecutive opens on same date into spreads
 * - Merges into existing Legs positions (weighted avg price)
 * - Adds new spreads as new groups
 * - Tracks LastImportDate to avoid duplicates
 */

const ETRADE_CONFIG_RANGE = "ImportConfig"; // Named range for config (LastImportDate, etc.)

/**
 * Imports E*Trade transactions from a CSV file in Google Drive.
 * Call from menu or script.
 *
 * @param {string} [fileName] - CSV filename in Drive (default: "DownloadTxnHistory.csv")
 * @param {string} [folderPath] - Folder path in Drive (default: "Investing/Data")
 */
function importEtradeTransactions(fileName, folderPath) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get config
  const config = getImportConfig_(ss);
  const lastImportDate = config.lastImportDate;

  // Find CSV file in Drive folder
  const csvName = fileName || "DownloadTxnHistory.csv";
  const path = folderPath || "Investing/Data";

  // Navigate to folder
  const root = DriveApp.getRootFolder();
  let folder = root;
  const parts = path.split('/').filter(p => p.trim());
  for (const part of parts) {
    folder = getFolder_(folder, part);
  }

  // Find file in folder
  const files = folder.getFilesByName(csvName);
  if (!files.hasNext()) {
    SpreadsheetApp.getUi().alert(`File not found: ${csvName}\n\nUpload the E*Trade transaction CSV to Google Drive under /${path}/`);
    return;
  }
  const file = files.next();
  const csvContent = file.getBlob().getDataAsString();

  // Parse CSV
  const { transactions, stockTxns } = parseEtradeCsv_(csvContent, lastImportDate);
  Logger.log(`Parsed ${transactions.length} new transactions after ${lastImportDate || 'beginning'}`);

  if (transactions.length === 0) {
    SpreadsheetApp.getUi().alert("No new transactions to import.");
    return;
  }

  // Pair into spreads (opens only)
  const spreads = pairTransactionsIntoSpreads_(transactions);
  Logger.log(`Paired into ${spreads.length} spread orders`);

  // Build map of closing prices by leg key
  const closingPrices = buildClosingPricesMap_(transactions, stockTxns);
  Logger.log(`Found closing prices for ${closingPrices.size} legs`);

  // Read existing Legs table
  const legsRange = getNamedRangeWithTableFallback_(ss, "Legs");
  let existingLegs = new Map();
  let headers = [];
  if (legsRange) {
    const rows = legsRange.getValues();
    headers = rows[0];
    existingLegs = parseLegsTable_(rows);
  }

  // Merge spreads into existing positions
  const { updatedLegs, newLegs } = mergeSpreads_(existingLegs, spreads);

  // Write back to Legs table
  writeLegsTable_(ss, headers, updatedLegs, newLegs, closingPrices);

  // Update LastImportDate
  const maxDate = transactions.reduce((max, t) => t.date > max ? t.date : max, lastImportDate || "");
  setImportConfig_(ss, { lastImportDate: maxDate });

  // Report
  SpreadsheetApp.getUi().alert(
    `Import Complete\n\n` +
    `Transactions: ${transactions.length}\n` +
    `Spread orders: ${spreads.length}\n` +
    `Updated positions: ${updatedLegs.length}\n` +
    `New positions: ${newLegs.length}`
  );
}

/**
 * Gets import config from named range.
 */
function getImportConfig_(ss) {
  const range = ss.getRangeByName(ETRADE_CONFIG_RANGE);
  if (!range) return { lastImportDate: null };

  const values = range.getValues();
  const config = {};
  for (const row of values) {
    const key = String(row[0] || "").trim();
    const val = row[1];
    if (key === "LastImportDate") {
      config.lastImportDate = val instanceof Date
        ? Utilities.formatDate(val, Session.getScriptTimeZone(), "MM/dd/yy")
        : String(val || "").trim();
    }
  }
  return config;
}

/**
 * Sets import config in named range.
 */
function setImportConfig_(ss, config) {
  let range = ss.getRangeByName(ETRADE_CONFIG_RANGE);
  if (!range) {
    // Create config range on a Config sheet
    let sheet = ss.getSheetByName("Config");
    if (!sheet) sheet = ss.insertSheet("Config");
    range = sheet.getRange("A1:B2");
    ss.setNamedRange(ETRADE_CONFIG_RANGE, range);
    range.setValues([["Setting", "Value"], ["LastImportDate", ""]]);
  }

  const values = range.getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === "LastImportDate") {
      values[i][1] = config.lastImportDate || "";
    }
  }
  range.setValues(values);
}

/**
 * Parses E*Trade CSV content into transaction objects.
 * Returns { transactions, stockTxns }.
 * transactions: option opens, closes, exercises, assignments after lastImportDate
 * stockTxns: stock Bought/Sold for matching exercise/assignment to market price
 */
function parseEtradeCsv_(csvContent, lastImportDate) {
  const lines = csvContent.split(/\r?\n/);
  const transactions = [];
  const stockTxns = [];

  // Find header row
  let headerIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].startsWith("TransactionDate,")) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx < 0) return { transactions, stockTxns };

  const optionTypes = [
    "Bought To Open", "Sold Short",
    "Sold To Close", "Bought To Cover",
    "Option Assigned", "Option Exercised",
  ];

  // Parse rows
  for (let i = headerIdx + 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    const cols = parseCsvLine_(line);
    if (cols.length < 8) continue;

    const [dateStr, txnType, secType, symbol, qtyStr, amountStr, priceStr, commStr] = cols;

    // Filter by date
    if (lastImportDate && dateStr <= lastImportDate) continue;

    // Stock transactions (for exercise/assignment matching)
    if (secType === "EQ" && (txnType === "Bought" || txnType === "Sold")) {
      stockTxns.push({
        date: dateStr,
        ticker: symbol.trim().toUpperCase(),
        qty: parseFloat(qtyStr) || 0,
        price: parseFloat(priceStr) || 0,
      });
      continue;
    }

    // Option transactions
    if (secType !== "OPTN") continue;
    if (!optionTypes.includes(txnType)) continue;

    // Parse option symbol: "TSLA Dec 15 '28 $400 Call"
    const parsed = parseEtradeOptionSymbol_(symbol);
    if (!parsed) continue;

    transactions.push({
      date: dateStr,
      txnType,
      ticker: parsed.ticker,
      expiration: parsed.expiration,
      strike: parsed.strike,
      optionType: parsed.type,
      qty: parseFloat(qtyStr) || 0,
      price: parseFloat(priceStr) || 0,
      amount: parseFloat(amountStr) || 0,
      isOpen: txnType === "Bought To Open" || txnType === "Sold Short",
      isClosed: txnType === "Sold To Close" || txnType === "Bought To Cover",
      isExercised: txnType === "Option Exercised",
      isAssigned: txnType === "Option Assigned",
    });
  }

  return { transactions, stockTxns };
}

/**
 * Parses a CSV line handling quoted fields.
 */
function parseCsvLine_(line) {
  const result = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      inQuotes = !inQuotes;
    } else if (ch === ',' && !inQuotes) {
      result.push(current.trim());
      current = "";
    } else {
      current += ch;
    }
  }
  result.push(current.trim());
  return result;
}

/**
 * Parses E*Trade option symbol like "TSLA Dec 15 '28 $400 Call"
 * Returns { ticker, expiration, strike, type }
 */
function parseEtradeOptionSymbol_(symbol) {
  // Pattern: TICKER Mon DD 'YY $Strike Type
  const match = symbol.match(/^(\w+)\s+(\w+)\s+(\d+)\s+'(\d+)\s+\$(\d+(?:\.\d+)?)\s+(Call|Put)$/i);
  if (!match) return null;

  const [, ticker, month, day, year, strike, type] = match;

  const months = {
    Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6,
    Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12
  };
  const monthNum = months[month];
  if (!monthNum) return null;

  const fullYear = 2000 + parseInt(year, 10);
  const expiration = `${monthNum}/${day}/${fullYear}`;

  return {
    ticker: ticker.toUpperCase(),
    expiration,
    strike: parseFloat(strike),
    type: type.charAt(0).toUpperCase() + type.slice(1).toLowerCase(), // "Call" or "Put"
  };
}

/**
 * Pairs consecutive opens on same date into spread orders.
 * Detects iron condors (2 puts + 2 calls with matching qty).
 */
function pairTransactionsIntoSpreads_(transactions) {
  const spreads = [];

  // Group by date + ticker + expiration
  const groups = new Map();
  for (const txn of transactions) {
    if (!txn.isOpen) continue; // Only process opens for now

    const key = `${txn.date}|${txn.ticker}|${txn.expiration}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(txn);
  }

  // Pair within each group
  for (const [key, txns] of groups) {
    // Separate by option type first
    const calls = txns.filter(t => t.optionType === "Call");
    const puts = txns.filter(t => t.optionType === "Put");

    // Check for iron condor: 2 calls (1 long, 1 short) + 2 puts (1 long, 1 short)
    const longCalls = calls.filter(t => t.qty > 0);
    const shortCalls = calls.filter(t => t.qty < 0);
    const longPuts = puts.filter(t => t.qty > 0);
    const shortPuts = puts.filter(t => t.qty < 0);

    if (longCalls.length === 1 && shortCalls.length === 1 &&
        longPuts.length === 1 && shortPuts.length === 1) {
      // Check if quantities match (all same absolute qty)
      const qty = longCalls[0].qty;
      if (Math.abs(shortCalls[0].qty) === qty &&
          longPuts[0].qty === qty &&
          Math.abs(shortPuts[0].qty) === qty) {
        // Iron condor detected - create as 4 legs
        spreads.push({
          type: "iron-condor",
          ticker: longCalls[0].ticker,
          expiration: longCalls[0].expiration,
          qty: qty,
          date: longCalls[0].date,
          legs: [
            { strike: longPuts[0].strike, optionType: "Put", qty: qty, price: longPuts[0].price },
            { strike: shortPuts[0].strike, optionType: "Put", qty: -qty, price: shortPuts[0].price },
            { strike: shortCalls[0].strike, optionType: "Call", qty: -qty, price: shortCalls[0].price },
            { strike: longCalls[0].strike, optionType: "Call", qty: qty, price: longCalls[0].price },
          ].sort((a, b) => a.strike - b.strike),
        });
        continue; // Skip normal pairing for this group
      }
    }

    // Normal pairing: pair calls with calls, puts with puts
    for (const optionType of ["Call", "Put"]) {
      const legsOfType = txns.filter(t => t.optionType === optionType);
      const longs = legsOfType.filter(t => t.qty > 0).sort((a, b) => a.strike - b.strike);
      const shorts = legsOfType.filter(t => t.qty < 0).sort((a, b) => a.strike - b.strike);

      // Pair by matching quantities
      let li = 0, si = 0;
      while (li < longs.length && si < shorts.length) {
        const long = longs[li];
        const short = shorts[si];

        const pairQty = Math.min(long.qty, Math.abs(short.qty));

        spreads.push({
          ticker: long.ticker,
          expiration: long.expiration,
          lowerStrike: long.strike,
          upperStrike: short.strike,
          optionType: long.optionType,
          qty: pairQty,
          lowerPrice: long.price,
          upperPrice: short.price,
          date: long.date,
        });

        long.qty -= pairQty;
        short.qty += pairQty; // short.qty is negative, so adding makes it less negative

        if (long.qty === 0) li++;
        if (short.qty === 0) si++;
      }

      // Handle unmatched legs (naked positions)
      while (li < longs.length) {
        const long = longs[li];
        if (long.qty > 0) {
          spreads.push({
            ticker: long.ticker,
            expiration: long.expiration,
            lowerStrike: long.strike,
            upperStrike: null, // Naked long
            optionType: long.optionType,
            qty: long.qty,
            lowerPrice: long.price,
            upperPrice: 0,
            date: long.date,
          });
        }
        li++;
      }
      while (si < shorts.length) {
        const short = shorts[si];
        if (short.qty < 0) {
          spreads.push({
            ticker: short.ticker,
            expiration: short.expiration,
            lowerStrike: null, // Naked short
            upperStrike: short.strike,
            optionType: short.optionType,
            qty: short.qty,
            lowerPrice: 0,
            upperPrice: short.price,
            date: short.date,
          });
        }
        si++;
      }
    }
  }

  return spreads;
}

/**
 * Builds a map of closing prices from close transactions.
 * Key: "TICKER|EXPIRATION|STRIKE|TYPE" -> price
 *
 * Handles:
 * 1. Sold To Close / Bought To Cover → use transaction price
 * 2. Option Exercised / Option Assigned → compute intrinsic from stock transactions
 * 3. Expired worthless (expiration < today, no close) → set to 0
 *
 * For multiple closes of same leg, uses weighted average.
 */
function buildClosingPricesMap_(transactions, stockTxns) {
  const result = new Map();
  const closes = new Map(); // key -> { totalQty, totalValue }

  // 1. Normal closes (Sold To Close, Bought To Cover)
  for (const txn of transactions) {
    if (!txn.isClosed) continue;

    const key = `${txn.ticker}|${txn.expiration}|${txn.strike}|${txn.optionType}`;
    const qty = Math.abs(txn.qty);
    const value = qty * txn.price;

    if (!closes.has(key)) closes.set(key, { totalQty: 0, totalValue: 0 });
    const entry = closes.get(key);
    entry.totalQty += qty;
    entry.totalValue += value;
  }

  for (const [key, { totalQty, totalValue }] of closes) {
    if (totalQty > 0) {
      result.set(key, roundTo_(totalValue / totalQty, 2));
    }
  }

  // 2. Exercise/Assignment → compute intrinsic from stock transactions
  // Group stock txns by date + ticker to find market price reference
  const stockByDateTicker = new Map(); // "date|ticker" -> [prices]
  for (const stk of (stockTxns || [])) {
    const key = `${stk.date}|${stk.ticker}`;
    if (!stockByDateTicker.has(key)) stockByDateTicker.set(key, []);
    stockByDateTicker.get(key).push(stk.price);
  }

  for (const txn of transactions) {
    if (!txn.isExercised && !txn.isAssigned) continue;

    const key = `${txn.ticker}|${txn.expiration}|${txn.strike}|${txn.optionType}`;
    if (result.has(key)) continue; // Already have a closing price

    // Find stock transactions on same date for same ticker
    const stkKey = `${txn.date}|${txn.ticker}`;
    const stockPrices = stockByDateTicker.get(stkKey) || [];

    if (stockPrices.length > 0) {
      // Use highest stock price as market price proxy
      // (for paired exercise/assignment, this gives correct spread P&L)
      const marketPrice = Math.max(...stockPrices);

      let intrinsic;
      if (txn.optionType === "Call") {
        intrinsic = Math.max(0, marketPrice - txn.strike);
      } else {
        intrinsic = Math.max(0, txn.strike - marketPrice);
      }
      result.set(key, roundTo_(intrinsic, 2));
    }
    // If no stock transactions found, leave blank (user fills manually)
  }

  // 3. Expired worthless: if expiration < today and no close, set to 0
  const today = new Date();
  const openLegs = new Set();
  for (const txn of transactions) {
    if (!txn.isOpen) continue;
    const key = `${txn.ticker}|${txn.expiration}|${txn.strike}|${txn.optionType}`;
    openLegs.add(key);
  }

  for (const legKey of openLegs) {
    if (result.has(legKey)) continue; // Already have closing price

    // Parse expiration from key
    const parts = legKey.split("|");
    const expStr = parts[1]; // e.g., "12/19/2025"
    const expDate = new Date(expStr);
    if (!isNaN(expDate) && expDate < today) {
      result.set(legKey, 0); // Expired worthless
    }
  }

  return result;
}

/**
 * Parses existing Legs table into position objects.
 */
function parseLegsTable_(rows) {
  if (rows.length < 2) return [];

  const headers = rows[0];
  const idxSym = findColumn_(headers, ["symbol", "ticker"]);
  const idxGroup = findColumn_(headers, ["group", "grp"]);
  const idxStrike = findColumn_(headers, ["strike", "strikeprice"]);
  const idxType = findColumn_(headers, ["type", "optiontype", "callput"]);
  const idxExp = findColumn_(headers, ["expiration", "exp"]);
  const idxQty = findColumn_(headers, ["qty", "quantity"]);
  const idxPrice = findColumn_(headers, ["price", "cost"]);

  const positions = new Map(); // key -> { legs, groupNum }

  let lastSym = "";
  let lastGroup = "";
  let currentLegs = [];

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];

    const rawSym = idxSym >= 0 ? String(row[idxSym] || "").trim().toUpperCase() : "";
    if (rawSym) lastSym = rawSym;

    const rawGroup = idxGroup >= 0 ? String(row[idxGroup] || "").trim() : "";
    if (rawGroup && rawGroup !== lastGroup) {
      // Save previous group
      if (currentLegs.length > 0) {
        const key = makeSpreadKey_(currentLegs);
        if (key) positions.set(key, { legs: currentLegs, groupNum: lastGroup });
      }
      lastGroup = rawGroup;
      currentLegs = [];
    }

    const strike = idxStrike >= 0 ? parseNumber_(row[idxStrike]) : NaN;
    const type = idxType >= 0 ? parseOptionType_(row[idxType]) : null;
    const exp = idxExp >= 0 ? row[idxExp] : "";
    const qty = idxQty >= 0 ? parseNumber_(row[idxQty]) : NaN;
    const price = idxPrice >= 0 ? parseNumber_(row[idxPrice]) : NaN;

    if (Number.isFinite(strike) && Number.isFinite(qty)) {
      currentLegs.push({
        symbol: lastSym,
        strike,
        type,
        expiration: exp,
        qty,
        price,
        row: r,
      });
    }
  }

  // Save last group
  if (currentLegs.length > 0) {
    const key = makeSpreadKey_(currentLegs);
    if (key) positions.set(key, { legs: currentLegs, groupNum: lastGroup });
  }

  return positions;
}

/**
 * Creates a unique key for a spread from its legs.
 */
function makeSpreadKey_(legs) {
  if (legs.length === 0) return null;

  const ticker = legs[0].symbol;
  const exp = normalizeExpiration_(legs[0].expiration) || legs[0].expiration;
  const strikes = legs.map(l => l.strike).sort((a, b) => a - b);
  const type = legs[0].type || "Call";

  return `${ticker}|${exp}|${strikes.join("/")}|${type}`;
}

/**
 * Creates spread key from a spread order.
 */
function makeSpreadKeyFromOrder_(spread) {
  const exp = normalizeExpiration_(spread.expiration) || spread.expiration;

  if (spread.type === "iron-condor" && spread.legs) {
    const strikes = spread.legs.map(l => l.strike).sort((a, b) => a - b);
    return `${spread.ticker}|${exp}|${strikes.join("/")}|IC`;
  }

  const strikes = [spread.lowerStrike, spread.upperStrike].filter(s => s != null).sort((a, b) => a - b);
  return `${spread.ticker}|${exp}|${strikes.join("/")}|${spread.optionType}`;
}

/**
 * Merges new spreads into existing positions.
 */
function mergeSpreads_(existingPositions, newSpreads) {
  const updatedLegs = [];
  const newLegs = [];
  const processedKeys = new Set();

  for (const spread of newSpreads) {
    const key = makeSpreadKeyFromOrder_(spread);

    if (existingPositions.has(key)) {
      // Merge into existing
      const existing = existingPositions.get(key);
      const longLeg = existing.legs.find(l => l.qty > 0);
      const shortLeg = existing.legs.find(l => l.qty < 0);

      if (longLeg && spread.lowerStrike) {
        const oldQty = longLeg.qty;
        const newQty = spread.qty;
        const totalQty = oldQty + newQty;
        longLeg.price = (oldQty * longLeg.price + newQty * spread.lowerPrice) / totalQty;
        longLeg.qty = totalQty;
      }

      if (shortLeg && spread.upperStrike) {
        const oldQty = Math.abs(shortLeg.qty);
        const newQty = spread.qty;
        const totalQty = oldQty + newQty;
        shortLeg.price = (oldQty * shortLeg.price + newQty * spread.upperPrice) / totalQty;
        shortLeg.qty = -(totalQty);
      }

      if (!processedKeys.has(key)) {
        updatedLegs.push(existing);
        processedKeys.add(key);
      }
    } else {
      // New spread — preserve full structure (including iron condor legs)
      newLegs.push(spread);
    }
  }

  return { updatedLegs, newLegs };
}

/**
 * Writes the Legs table back to the sheet.
 * @param {Map} [closingPrices] - Map of leg keys to closing prices
 */
function writeLegsTable_(ss, headers, updatedLegs, newLegs, closingPrices) {
  closingPrices = closingPrices || new Map();

  const legsRange = getNamedRangeWithTableFallback_(ss, "Legs");
  if (!legsRange || headers.length === 0) {
    Logger.log("Legs table not found, creating new one");
    // Create new Legs sheet
    let sheet = ss.getSheetByName("Legs");
    if (!sheet) sheet = ss.insertSheet("Legs");

    headers = ["Symbol", "Group", "Strategy", "Strike", "Type", "Expiration", "Qty", "Price", "Investment", "Rec Close", "Closed", "Gain", "Link"];
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground("#93c47d"); // Green header
    headerRange.setFontWeight("bold");
    ss.setNamedRange("LegsTable", sheet.getRange("A:M"));

    // Add filter to the data range
    const filterRange = sheet.getRange("A:M");
    filterRange.createFilter();

    // Set wrap to clip
    sheet.getRange("A:M").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  const range = getNamedRangeWithTableFallback_(ss, "Legs");
  const sheet = range.getSheet();
  const startRow = range.getRow();
  const startCol = range.getColumn();

  // Find column indexes
  const idxSym = findColumn_(headers, ["symbol", "ticker"]);
  const idxGroup = findColumn_(headers, ["group", "grp"]);
  const idxStrategy = findColumn_(headers, ["strategy"]);
  const idxStrike = findColumn_(headers, ["strike"]);
  const idxType = findColumn_(headers, ["type"]);
  const idxExp = findColumn_(headers, ["expiration", "exp"]);
  const idxQty = findColumn_(headers, ["qty", "quantity"]);
  const idxPrice = findColumn_(headers, ["price"]);
  const idxInvestment = findColumn_(headers, ["investment"]);
  const idxRecClose = findColumn_(headers, ["recclose", "rec close"]);
  const idxClosed = findColumn_(headers, ["closed", "actualclose", "closedat"]);
  const idxGain = findColumn_(headers, ["gain"]);
  const idxLink = findColumn_(headers, ["link"]);

  // Column letters for formulas
  const colLetter = (idx) => String.fromCharCode(65 + idx);
  const symCol = idxSym >= 0 ? colLetter(idxSym) : "A";
  const strikeCol = idxStrike >= 0 ? colLetter(idxStrike) : "D";
  const typeCol = idxType >= 0 ? colLetter(idxType) : "E";
  const expCol = idxExp >= 0 ? colLetter(idxExp) : "F";
  const qtyCol = idxQty >= 0 ? colLetter(idxQty) : "G";
  const priceCol = idxPrice >= 0 ? colLetter(idxPrice) : "H";
  const recCloseCol = idxRecClose >= 0 ? colLetter(idxRecClose) : "J";
  const closedCol = idxClosed >= 0 ? colLetter(idxClosed) : "K";

  // Update existing rows
  for (const pos of updatedLegs) {
    for (const leg of pos.legs) {
      if (leg.row != null) {
        const rowNum = startRow + leg.row;
        if (idxQty >= 0) sheet.getRange(rowNum, startCol + idxQty).setValue(leg.qty);
        if (idxPrice >= 0) sheet.getRange(rowNum, startCol + idxPrice).setValue(roundTo_(leg.price, 2));
      }
    }
  }

  // Append new spreads
  if (newLegs.length > 0) {
    let lastRow = sheet.getLastRow();
    let nextGroup = 1;

    // Find max group number
    if (idxGroup >= 0 && lastRow > startRow) {
      const groupData = sheet.getRange(startRow + 1, startCol + idxGroup, lastRow - startRow, 1).getValues();
      for (const row of groupData) {
        const g = parseInt(row[0], 10);
        if (Number.isFinite(g) && g >= nextGroup) nextGroup = g + 1;
      }
    }

    for (const spread of newLegs) {
      const rows = [];

      // Helper to get closing price for a leg
      const getClosingPrice = (ticker, expiration, strike, optionType) => {
        const key = `${ticker}|${expiration}|${strike}|${optionType}`;
        const val = closingPrices.get(key);
        return val != null ? val : "";
      };

      // Handle iron condor (4 legs)
      if (spread.type === "iron-condor" && spread.legs) {
        for (let i = 0; i < spread.legs.length; i++) {
          const leg = spread.legs[i];
          const row = new Array(headers.length).fill("");
          if (i === 0) {
            // First row gets symbol and group
            if (idxSym >= 0) row[idxSym] = spread.ticker;
            if (idxGroup >= 0) row[idxGroup] = nextGroup;
          }
          if (idxStrike >= 0) row[idxStrike] = leg.strike;
          if (idxType >= 0) row[idxType] = leg.optionType;
          if (idxExp >= 0) row[idxExp] = spread.expiration;
          if (idxQty >= 0) row[idxQty] = leg.qty;
          if (idxPrice >= 0) row[idxPrice] = roundTo_(leg.price, 2);
          if (idxClosed >= 0) row[idxClosed] = getClosingPrice(spread.ticker, spread.expiration, leg.strike, leg.optionType);
          rows.push(row);
        }
      } else {
        // Handle 2-leg spread or single leg
        const numLegs = (spread.lowerStrike != null ? 1 : 0) + (spread.upperStrike != null ? 1 : 0);
        if (numLegs === 0) continue;

        // Long leg (first row)
        if (spread.lowerStrike != null) {
          const row = new Array(headers.length).fill("");
          if (idxSym >= 0) row[idxSym] = spread.ticker;
          if (idxGroup >= 0) row[idxGroup] = nextGroup;
          if (idxStrike >= 0) row[idxStrike] = spread.lowerStrike;
          if (idxType >= 0) row[idxType] = spread.optionType;
          if (idxExp >= 0) row[idxExp] = spread.expiration;
          if (idxQty >= 0) row[idxQty] = spread.qty;
          if (idxPrice >= 0) row[idxPrice] = roundTo_(spread.lowerPrice, 2);
          if (idxClosed >= 0) row[idxClosed] = getClosingPrice(spread.ticker, spread.expiration, spread.lowerStrike, spread.optionType);
          rows.push(row);
        }

        // Short leg (second row)
        if (spread.upperStrike != null) {
          const row = new Array(headers.length).fill("");
          // Symbol and Group only on first row (blank for carry-forward)
          if (idxStrike >= 0) row[idxStrike] = spread.upperStrike;
          if (idxType >= 0) row[idxType] = spread.optionType;
          if (idxExp >= 0) row[idxExp] = spread.expiration;
          if (idxQty >= 0) row[idxQty] = -spread.qty;
          if (idxPrice >= 0) row[idxPrice] = roundTo_(spread.upperPrice, 2);
          if (idxClosed >= 0) row[idxClosed] = getClosingPrice(spread.ticker, spread.expiration, spread.upperStrike, spread.optionType);
          rows.push(row);
        }
      }

      if (rows.length === 0) continue;

      const firstRow = lastRow + 1;
      const lastLegRow = firstRow + rows.length - 1;

      // Write data rows
      sheet.getRange(firstRow, startCol, rows.length, headers.length).setValues(rows);

      // Write formulas on first row of group
      const rangeStr = `${firstRow}:${lastLegRow}`;

      // Strategy formula
      if (idxStrategy >= 0) {
        const formula = `=detectStrategy($${strikeCol}${firstRow}:$${strikeCol}${lastLegRow}, $${typeCol}${firstRow}:$${typeCol}${lastLegRow}, $${qtyCol}${firstRow}:$${qtyCol}${lastLegRow})`;
        sheet.getRange(firstRow, startCol + idxStrategy).setFormula(formula);
      }

      // Investment formula
      if (idxInvestment >= 0) {
        const formula = `=SUMPRODUCT($${qtyCol}${firstRow}:$${qtyCol}${lastLegRow}, $${priceCol}${firstRow}:$${priceCol}${lastLegRow}) * 100`;
        sheet.getRange(firstRow, startCol + idxInvestment).setFormula(formula);
      }

      // Gain formula: use Closed if available, otherwise Rec Close
      if (idxGain >= 0 && idxRecClose >= 0) {
        let closeRef;
        if (idxClosed >= 0) {
          // IF(Closed<>"", Closed, RecClose) per leg
          closeRef = `IF($${closedCol}${firstRow}:$${closedCol}${lastLegRow}<>"", $${closedCol}${firstRow}:$${closedCol}${lastLegRow}, $${recCloseCol}${firstRow}:$${recCloseCol}${lastLegRow})`;
        } else {
          closeRef = `$${recCloseCol}${firstRow}:$${recCloseCol}${lastLegRow}`;
        }
        const formula = `=SUMPRODUCT($${qtyCol}${firstRow}:$${qtyCol}${lastLegRow}, ${closeRef} - $${priceCol}${firstRow}:$${priceCol}${lastLegRow}) * 100`;
        sheet.getRange(firstRow, startCol + idxGain).setFormula(formula);
      }

      // Link formula: HYPERLINK with "OptionStrat" display text
      if (idxLink >= 0) {
        const urlFormula = `buildOptionStratUrlFromLegs($${symCol}$1:$${symCol}${firstRow}, $${strikeCol}${firstRow}:$${strikeCol}${lastLegRow}, $${typeCol}${firstRow}:$${typeCol}${lastLegRow}, $${expCol}${firstRow}:$${expCol}${lastLegRow}, $${qtyCol}${firstRow}:$${qtyCol}${lastLegRow})`;
        const formula = `=HYPERLINK(${urlFormula}, "OptionStrat")`;
        sheet.getRange(firstRow, startCol + idxLink).setFormula(formula);
      }

      // Rec Close formula for each leg (skip if Closed is already populated)
      if (idxRecClose >= 0) {
        for (let i = 0; i < rows.length; i++) {
          const legRow = firstRow + i;
          const hasClosed = idxClosed >= 0 && rows[i][idxClosed] !== "";
          if (!hasClosed) {
            const formula = `=recommendClose($${symCol}$1:$${symCol}${legRow}, $${expCol}${legRow}, $${strikeCol}${legRow}, $${typeCol}${legRow}, $${qtyCol}${legRow}, 60)`;
            sheet.getRange(legRow, startCol + idxRecClose).setFormula(formula);
          }
        }
      }

      // Alternate group colors: odd = pale yellow, even = white
      const bgColor = (nextGroup % 2 === 1) ? "#fff2cc" : "#ffffff";
      sheet.getRange(firstRow, startCol, rows.length, headers.length).setBackground(bgColor);

      lastRow = lastLegRow;
      nextGroup++;
    }
  }
}
