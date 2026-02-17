// @ts-check
/**
 * ImportEtrade.js
 * Imports E*Trade portfolio and transactions into the Portfolio sheet.
 *
 * Features:
 * - Parses E*Trade transaction CSV and portfolio CSV
 * - Pairs consecutive opens on same date into spreads
 * - Merges into existing Portfolio positions (weighted avg price)
 * - Adds new spreads as new groups
 * - Per-group LastTxnDate deduplication
 */

/**
 * Import Latest Transactions - adds new transactions, skips duplicates.
 * Menu action for incremental imports.
 */
function importLatestTransactions() {
  importEtradePortfolio_("update");
}

/**
 * Clear & Rebuild Portfolio - deletes existing portfolio and imports all transactions fresh.
 * Menu action for full rebuild.
 */
function rebuildPortfolio() {
  importEtradePortfolio_("rebuild");
}

/**
 * Imports E*Trade portfolio and transactions from CSV files in Google Drive.
 *
 * @param {string} importMode - "fresh" (no existing), "update" (merge), or "rebuild" (delete and recreate)
 * @param {string} [fileName] - Transaction CSV filename (default: all "DownloadTxnHistory*.csv" files)
 * @param {string} [folderPath] - Folder path in Drive (default: "<DataFolder>/Etrade")
 */
function importEtradePortfolio_(importMode, fileName, folderPath) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const path = folderPath || getConfigValue_(ss, "DataFolder", "SpreadFinder/DATA") + "/Etrade";

  // Navigate to folder
  const root = DriveApp.getRootFolder();
  let folder = root;
  const parts = path.split('/').filter(p => p.trim());
  try {
    for (const part of parts) {
      folder = getFolder_(folder, part);
    }
  } catch (e) {
    ui.alert(
      "Folder Not Found",
      `Folder not found: ${path}\n\n` +
      `To set up your E*Trade import folder:\n` +
      `1. Create the folder in Google Drive: ${path}\n` +
      `2. Upload your E*Trade CSV files there\n\n` +
      `Or change the DataFolder setting on the Config sheet.`,
      ui.ButtonSet.OK
    );
    return;
  }

  // Handle rebuild mode - check for custom groups and warn before deleting
  if (importMode === "rebuild") {
    const existingSheet = ss.getSheetByName("Portfolio");
    if (existingSheet) {
      // Check for custom groups before deleting
      const customGroups = findCustomGroups_(existingSheet);
      if (customGroups.length > 0) {
        const groupList = customGroups.map(g => `  • Group ${g.group}: ${g.description}`).join("\n");
        const resp = ui.alert(
          "Custom Groups Will Be Lost",
          `The following custom groups cannot be auto-recreated from transactions:\n\n${groupList}\n\n` +
          `These would need to be manually re-created after rebuild.\n\n` +
          `Continue with rebuild?`,
          ui.ButtonSet.OK_CANCEL
        );
        if (resp !== ui.Button.OK) return;
      }
      ss.deleteSheet(existingSheet);
      const nr = ss.getNamedRanges().find(r => r.getName() === "PortfolioTable");
      if (nr) nr.remove();
    }
  }

  // Read existing Portfolio table (if kept) to merge with imported data
  const portfolioRange = getNamedRangeWithTableFallback_(ss, "Portfolio");
  let existingPositions = new Map();
  let headers = [];
  if (portfolioRange) {
    const rows = portfolioRange.getValues();
    headers = rows[0];
    existingPositions = parsePortfolioTable_(rows);
  }

  // Find CSV files
  const txnFiles = fileName
    ? findFilesByName_(folder, fileName)
    : findFilesByPrefix_(folder, "DownloadTxnHistory");
  const portfolioFiles = findFilesByPrefix_(folder, "PortfolioDownload");

  if (txnFiles.length === 0 && portfolioFiles.length === 0) {
    SpreadsheetApp.getUi().alert(
      `No E*Trade CSV files found.\n\n` +
      `To import your portfolio:\n` +
      `1. Log into E*Trade\n` +
      `2. Download "Portfolio" CSV (Accounts > Portfolio > Download)\n` +
      `3. Download "Transaction History" CSV (Accounts > Transactions > Download)\n` +
      `4. Upload both files to Google Drive:\n` +
      `   ${path}/\n\n` +
      `Expected filenames:\n` +
      `  • PortfolioDownload*.csv\n` +
      `  • DownloadTxnHistory*.csv`
    );
    return;
  }

  if (txnFiles.length === 0) {
    SpreadsheetApp.getUi().alert(
      `No transaction history CSV found.\n\n` +
      `Found ${portfolioFiles.length} PortfolioDownload file(s), but no DownloadTxnHistory files.\n\n` +
      `To download transaction history from E*Trade:\n` +
      `1. Go to Accounts > Transactions\n` +
      `2. Select date range and click Download\n` +
      `3. Upload the CSV to: ${path}/`
    );
    return;
  }

  // Process all transaction CSVs, dedup across file boundaries (overlapping date ranges)
  log.info("import", `Processing ${txnFiles.length} transaction CSV(s)`);

  let transactions = [];
  let stockTxns = [];
  const seenTxns = new Set();
  const seenStockTxns = new Set();

  for (const file of txnFiles) {
    const csvContent = file.getBlob().getDataAsString();
    const result = parseEtradeTransactionsFromCsv_(csvContent);
    let txnAdded = 0, txnDupes = 0;
    for (const txn of result.transactions) {
      const key = `${txn.date}|${txn.txnType}|${txn.ticker}|${txn.expiration}|${txn.strike}|${txn.optionType}|${txn.qty}|${txn.price}|${txn.amount}`;
      if (seenTxns.has(key)) { txnDupes++; continue; }
      seenTxns.add(key);
      transactions.push(txn);
      txnAdded++;
    }
    for (const stk of result.stockTxns) {
      const key = `${stk.date}|${stk.ticker}|${stk.qty}|${stk.price}|${stk.amount}`;
      if (seenStockTxns.has(key)) continue;
      seenStockTxns.add(key);
      stockTxns.push(stk);
    }
    log.debug("import", `${file.getName()}: ${txnAdded} transactions (${txnDupes} dupes skipped)`);
  }

  if (transactions.length === 0) {
    const fileNames = txnFiles.map(f => f.getName()).join(", ");
    SpreadsheetApp.getUi().alert(
      `No option transactions found in CSV file(s).\n\n` +
      `Files checked: ${fileNames}\n\n` +
      `Make sure you downloaded the Transaction History CSV from E*Trade ` +
      `(Accounts > Transactions > Download), not a different report type.`
    );
    return;
  }

  // Get stock positions based on import mode
  let stockPositions = [];
  let portfolioCash = 0;
  if (importMode === "rebuild" || importMode === "fresh") {
    // For rebuild/fresh: use PortfolioDownload CSV for quantities (source of truth)
    // Transaction dates come from stockTxns (or NOW if none found)
    if (portfolioFiles.length > 0) {
      const portfolioResult = parsePortfolioStocksAndCashFromFile_(portfolioFiles[0], stockTxns);
      stockPositions = portfolioResult.stocks;
      portfolioCash = portfolioResult.cash || 0;
      log.info("import", `Found ${stockPositions.length} stock positions and $${portfolioCash} cash from portfolio CSV`);
    }
  } else {
    // For update mode: aggregate only NEW stock transactions (after existing LastTxnDate)
    const stockCutoffDates = new Map();
    for (const [key, pos] of existingPositions) {
      if (key.endsWith("|STOCK")) {
        const ticker = key.split("|")[0];
        const cutoff = parseDateAtMidnight_(pos.lastTxnDate);
        if (cutoff) {
          stockCutoffDates.set(ticker, cutoff);
        }
      }
    }
    stockPositions = aggregateStockTransactions_(stockTxns, stockCutoffDates);
    log.info("import", `Found ${stockPositions.length} stock positions with new transactions`);
  }

  // Pair into spreads (opens only)
  const rawSpreads = [...stockPositions, ...pairTransactionsIntoSpreads_(transactions)];

  // Add cash as a position if present
  if (portfolioCash > 0) {
    rawSpreads.push({
      type: "cash",
      ticker: "CASH",
      qty: 1,
      price: portfolioCash,
      date: new Date(),
      expiration: null,
      lowerStrike: null,
      upperStrike: null,
      optionType: "Cash",
    });
    log.debug("import", `Added cash position: $${portfolioCash}`);
  }

  // Pre-merge spreads with the same key, keeping latest date and summing quantities
  const spreads = preMergeSpreads_(rawSpreads);
  log.info("import", `Paired into ${spreads.length} spread orders (including stocks)`);

  // Build map of closing prices by leg key
  const closingPrices = buildClosingPricesMap_(transactions, stockTxns);
  log.debug("import", `Found closing prices for ${closingPrices.size} legs`);

  // Validate option quantities against portfolio CSV (for rebuild/fresh mode)
  let validationWarnings = "";
  if ((importMode === "rebuild" || importMode === "fresh") && portfolioFiles.length > 0) {
    const portfolioOptionData = parsePortfolioOptionsWithPricesFromFile_(portfolioFiles[0]);
    const validation = validateOptionQuantities_(spreads, portfolioOptionData.quantities);

    // Add orphaned options (in portfolio but not in transactions) as single-leg positions
    if (validation.missing.length > 0) {
      validationWarnings += "\n\n⚠️ ORPHANED OPTIONS (added as single legs):\n";
      for (const m of validation.missing) {
        const [ticker, exp, strike, type] = m.key.split("|");
        const priceInfo = portfolioOptionData.prices.get(m.key) || { pricePaid: 0 };

        // Create single-leg position
        // Long options: positive qty, use lowerStrike
        // Short options: negative qty, use upperStrike
        const isLong = m.portfolio > 0;
        const singleLeg = {
          type: "single-option",
          ticker: ticker,
          qty: isLong ? m.portfolio : Math.abs(m.portfolio), // Qty is positive, short indicated by upperStrike
          expiration: exp,
          optionType: type,
          lowerStrike: isLong ? parseFloat(strike) : null,
          upperStrike: isLong ? null : parseFloat(strike),
          lowerPrice: isLong ? priceInfo.pricePaid : 0,
          upperPrice: isLong ? 0 : priceInfo.pricePaid,
          date: new Date(),
        };
        spreads.push(singleLeg);

        const direction = isLong ? "long" : "short";
        validationWarnings += `  • Added ${ticker} ${exp} $${strike} ${type}: ${m.portfolio} contracts (${direction})\n`;
        log.debug("import", `Added orphaned option: ${m.key} qty=${m.portfolio}`);
      }
    }

    if (validation.mismatches.length > 0) {
      validationWarnings += "\n⚠️ QUANTITY MISMATCHES (check transactions):\n";
      for (const m of validation.mismatches) {
        const [ticker, exp, strike, type] = m.key.split("|");
        validationWarnings += `  • ${ticker} ${exp} $${strike} ${type}: expected ${m.expected}, portfolio has ${m.portfolio}\n`;
      }
    }

    // Note: validation.extra (in transactions but not in portfolio) is expected for closed positions
    // Only log for debugging, don't show as warning
    if (validation.extra.length > 0) {
      log.debug("import", "Closed positions (in transactions but not in portfolio): " +
        validation.extra.map(m => m.key).join(", "));
    }

    if (validationWarnings) {
      log.warn("import", "Validation warnings: " + validationWarnings);
    }
  }

  // Merge spreads into existing positions (per-group dedup via LastTxnDate)
  const { updatedLegs, newLegs, skippedCount } = mergeSpreads_(existingPositions, spreads);

  // Write back to Portfolio table
  writePortfolioTable_(ss, headers, updatedLegs, newLegs, closingPrices);

  // Report
  const txnFileNames = txnFiles.map(f => f.getName()).join("\n  ");
  const portfolioFileNames = portfolioFiles.map(f => f.getName()).join("\n  ");
  const modeLabel = importMode === "update" ? "Import Latest Transactions" : "Portfolio Import";
  let summary = `Transaction files:\n  ${txnFileNames || "(none)"}\n`;
  if (portfolioFiles.length > 0) {
    summary += `Portfolio files:\n  ${portfolioFileNames}\n`;
  }
  summary += `\nTransactions parsed: ${transactions.length}\n`;
  summary += `Spread orders: ${spreads.length}\n`;

  if (importMode === "update") {
    summary += `\nNew positions added: ${newLegs.length}`;
    if (newLegs.length > 0) {
      summary += "\n  " + newLegs.map(s => formatSpreadLabel_(s)).join("\n  ");
    }
    summary += `\n\nExisting positions updated: ${updatedLegs.length}`;
    if (updatedLegs.length > 0) {
      summary += "\n  " + updatedLegs.map(p => formatPositionLabel_(p)).join("\n  ");
    }
    summary += `\n\nSkipped (already imported): ${skippedCount}`;
  } else {
    summary += `Positions imported: ${newLegs.length}`;
    if (portfolioCash > 0) {
      summary += `\nCash: $${portfolioCash.toLocaleString()}`;
    }
    if (stockPositions.length > 0) {
      summary += `\nStocks: ${stockPositions.map(s => s.ticker).join(", ")}`;
    }
  }

  summary += validationWarnings;

  ui.alert(modeLabel + " Complete", summary, ui.ButtonSet.OK);
}

/**
 * Formats a spread order for display in the report.
 */
function formatSpreadLabel_(spread) {
  if (spread.type === "cash") {
    return `Cash $${spread.price}`;
  }
  if (spread.type === "stock") {
    return `${spread.ticker} Stock`;
  }
  if (spread.type === "iron-condor") {
    const strikes = spread.legs.map(l => l.strike).join("/");
    return `${spread.ticker} ${formatExpirationShort_(spread.expiration)} ${strikes} iron-condor`;
  }
  const strikes = [spread.lowerStrike, spread.upperStrike].filter(s => s).join("/");
  const strategyType = spread.lowerStrike && spread.upperStrike ? "bull-call-spread" :
                       spread.lowerStrike ? "long-call" : "short-call";
  return `${spread.ticker} ${formatExpirationShort_(spread.expiration)} ${strikes} ${strategyType}`;
}

/**
 * Formats an existing position for display in the report.
 */
function formatPositionLabel_(pos) {
  if (!pos.legs || pos.legs.length === 0) return "Unknown position";
  const leg = pos.legs[0];
  const strikes = pos.legs.map(l => l.strike).filter(s => s).sort((a, b) => a - b).join("/");
  const exp = formatExpirationShort_(leg.expiration);
  const debug = pos.debugReason || "";
  return `${leg.symbol} ${exp} ${strikes} ${leg.type || "Call"}${debug}`;
}

/**
 * Formats expiration as "Mon YYYY" for display.
 */
function formatExpirationShort_(exp) {
  if (!exp) return "";
  const d = parseDateAtMidnight_(exp);
  if (!d) return String(exp);
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return `${months[d.getMonth()]} ${d.getFullYear()}`;
}

/**
 * Formats a date as MM/DD/YYYY (e.g., 2/6/2026).
 */
function formatDateLong_(dateVal) {
  if (!dateVal) return "";
  const d = parseDateAtMidnight_(dateVal);
  if (!d) return String(dateVal);
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}

/**
 * Generates a description for a spread like "500/600 bull-call-spread".
 */
function generateSpreadDescription_(spread) {
  if (spread.type === "stock") {
    return "Stock";
  }

  if (spread.type === "cash") {
    return "Cash";
  }

  if (spread.type === "iron-condor" && spread.legs) {
    const strikes = spread.legs.map(l => l.strike).sort((a, b) => a - b).join("/");
    // Detect iron-butterfly (middle strikes are equal) vs iron-condor
    const sortedStrikes = spread.legs.map(l => l.strike).sort((a, b) => a - b);
    const isButterfly = sortedStrikes[1] === sortedStrikes[2];
    return `${strikes} ${isButterfly ? "iron-butterfly" : "iron-condor"}`;
  }

  // Custom multi-leg position (imbalanced qty or more than 2 strikes of same type)
  if (spread.type === "custom" && spread.legs) {
    // Format strikes with - prefix for shorts: 500/-600/740/-900
    const formattedStrikes = spread.legs
      .sort((a, b) => a.strike - b.strike)
      .map(l => l.qty < 0 ? `-${l.strike}` : `${l.strike}`)
      .join("/");
    return `${formattedStrikes} custom`;
  }

  // Straddle/strangle (2 legs: call + put)
  if (spread.legs && spread.legs.length === 2) {
    const strikes = spread.legs.map(l => l.strike).sort((a, b) => a - b);
    const strikeStr = strikes[0] === strikes[1] ? String(strikes[0]) : strikes.join("/");
    return `${strikeStr} ${spread.type}`;
  }

  // Regular spread
  const strikes = [spread.lowerStrike, spread.upperStrike].filter(s => s != null);
  const strikeStr = strikes.sort((a, b) => a - b).join("/");

  // Determine strategy type
  let strategy;
  if (spread.lowerStrike && spread.upperStrike) {
    // Two legs
    if (spread.optionType === "Call") {
      strategy = "bull-call-spread";
    } else {
      strategy = "bull-put-spread";
    }
  } else if (spread.lowerStrike) {
    // Long only
    strategy = spread.optionType === "Call" ? "long-call" : "long-put";
  } else {
    // Short only
    strategy = spread.optionType === "Call" ? "short-call" : "short-put";
  }

  return `${strikeStr} ${strategy}`;
}

/**
 * Imports E*Trade portfolio from a specific folder and filenames.
 * Used by loadSamplePortfolio to import sample data without UI prompts.
 *
 * @param {Folder} folder - Google Drive folder containing the CSV files
 * @param {string} txnFileName - Transaction history CSV filename
 * @param {string} portfolioFileName - Portfolio CSV filename
 */
function importEtradePortfolioFromFolder_(folder, txnFileName, portfolioFileName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Find the specific files
  const txnFiles = findFilesByName_(folder, txnFileName);
  const portfolioFiles = findFilesByName_(folder, portfolioFileName);

  if (txnFiles.length === 0) {
    throw new Error(`Transaction file not found: ${txnFileName}`);
  }

  // Process transaction CSV
  let transactions = [];
  let stockTxns = [];

  for (const file of txnFiles) {
    const csvContent = file.getBlob().getDataAsString();
    const result = parseEtradeTransactionsFromCsv_(csvContent);
    transactions.push(...result.transactions);
    stockTxns.push(...result.stockTxns);
  }

  if (transactions.length === 0) {
    throw new Error("No transactions found in " + txnFileName);
  }

  // Parse stock positions from portfolio CSV (qty and price are source of truth)
  // Then add latest transaction date from stockTxns (or NOW if no transactions found)
  let stockPositions = [];
  let portfolioCash = 0;
  if (portfolioFiles.length > 0) {
    const portfolioResult = parsePortfolioStocksAndCashFromFile_(portfolioFiles[0], stockTxns);
    stockPositions = portfolioResult.stocks;
    portfolioCash = portfolioResult.cash || 0;
  }

  // Pair into spreads
  const spreads = [...stockPositions, ...pairTransactionsIntoSpreads_(transactions)];

  // Add cash as a position if present
  if (portfolioCash > 0) {
    spreads.push({
      type: "cash",
      ticker: "CASH",
      qty: 1,
      price: portfolioCash,
      date: new Date(),
      expiration: null,
      lowerStrike: null,
      upperStrike: null,
      optionType: "Cash",
    });
  }

  // Build closing prices map
  const closingPrices = buildClosingPricesMap_(transactions, stockTxns);

  // Write to Portfolio table (fresh, no merge)
  writePortfolioTable_(ss, [], [], spreads, closingPrices);

  log.info("import", `Sample portfolio imported: ${spreads.length} positions`);
}

/**
 * Parses stock positions from a PortfolioDownload CSV file.
 * Quantities and prices come from the CSV (source of truth).
 * Transaction dates come from stockTxns (latest date per ticker, or NOW if none found).
 *
 * @param {File} file - The PortfolioDownload CSV file
 * @param {Object[]} stockTxns - Array of stock transactions from parseEtradeTransactionsFromCsv_
 * @returns {{ stocks: Object[], cash: number }} Object with stocks array and cash amount
 */
function parsePortfolioStocksAndCashFromFile_(file, stockTxns) {
  const csv = file.getBlob().getDataAsString();
  const lines = csv.split(/\r?\n/);

  // Find the data header row
  let headerIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].startsWith("Symbol,Last Price")) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx < 0) return { stocks: [], cash: 0 };

  const headers = parseCsvLine_(lines[headerIdx]);
  const idxSym = headers.findIndex(h => h === "Symbol");
  const idxQty = headers.findIndex(h => h === "Quantity");
  const idxPricePaid = headers.findIndex(h => h.startsWith("Price Paid"));
  const idxMarketValue = headers.findIndex(h => h.startsWith("Market Value") || h.startsWith("Value"));

  if (idxSym < 0 || idxQty < 0 || idxPricePaid < 0) return { stocks: [], cash: 0 };

  // Build map of latest transaction date per ticker
  const latestDateByTicker = buildLatestStockDates_(stockTxns);

  const stocks = [];
  let cash = 0;

  for (let i = headerIdx + 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    const cols = parseCsvLine_(line);
    const symbol = (cols[idxSym] || "").trim().toUpperCase();

    if (!symbol || symbol === "TOTAL") continue;

    // Extract cash from CASH row - value is in the "Value $" column
    if (symbol === "CASH") {
      if (idxMarketValue >= 0) {
        cash = parseFloat((cols[idxMarketValue] || "").replace(/[$,]/g, "")) || 0;
      }
      log.debug("import", `Found cash in portfolio: $${cash} (column ${idxMarketValue})`);
      continue;
    }

    if (symbol.includes(" ")) continue; // Skip options
    if (/\d{3,}/.test(symbol)) continue; // Skip CUSIPs

    const qty = parseFloat(cols[idxQty]) || 0;
    const pricePaid = parseFloat(cols[idxPricePaid]) || 0;
    if (qty === 0) continue;

    // Use latest transaction date, or NOW if no transactions found
    const lastDate = latestDateByTicker.get(symbol) || new Date();

    stocks.push({
      type: "stock",
      ticker: symbol,
      qty: qty,
      price: roundTo_(pricePaid, 2),
      date: lastDate,
      expiration: null,
      lowerStrike: null,
      upperStrike: null,
      optionType: "Stock",
    });
  }

  return { stocks, cash };
}

/**
 * Parses option positions from a PortfolioDownload CSV file.
 * Options have format like "TSLA Jun 16 '28 $350 Call"
 *
 * @param {File} file - The PortfolioDownload CSV file
 * @returns {{ quantities: Map<string, number>, prices: Map<string, {pricePaid: number}> }}
 */
function parsePortfolioOptionsWithPricesFromFile_(file) {
  const quantities = new Map();
  const prices = new Map();
  const csv = file.getBlob().getDataAsString();
  const lines = csv.split(/\r?\n/);

  // Find the data header row
  let headerIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].startsWith("Symbol,Last Price")) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx < 0) return { quantities, prices };

  const headers = parseCsvLine_(lines[headerIdx]);
  const idxSym = headers.findIndex(h => h === "Symbol");
  const idxQty = headers.findIndex(h => h === "Quantity");
  const idxPricePaid = headers.findIndex(h => h.startsWith("Price Paid"));

  if (idxSym < 0 || idxQty < 0) return { quantities, prices };

  // Month name to number mapping
  const monthMap = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
  };

  for (let i = headerIdx + 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    const cols = parseCsvLine_(line);
    const symbolStr = (cols[idxSym] || "").trim();
    const qty = parseFloat(cols[idxQty]) || 0;
    const pricePaid = idxPricePaid >= 0 ? (parseFloat(cols[idxPricePaid]) || 0) : 0;

    // Parse option format: "TSLA Jun 16 '28 $350 Call"
    // Pattern: TICKER MONTH DD 'YY $STRIKE TYPE
    const match = symbolStr.match(/^(\w+)\s+(\w{3})\s+(\d{1,2})\s+'(\d{2})\s+\$(\d+(?:\.\d+)?)\s+(Call|Put)$/i);
    if (!match) continue;

    const ticker = match[1].toUpperCase();
    const month = monthMap[match[2].toLowerCase()];
    const day = parseInt(match[3], 10);
    const year = 2000 + parseInt(match[4], 10);
    const strike = parseFloat(match[5]);
    const optionType = match[6].charAt(0).toUpperCase() + match[6].slice(1).toLowerCase();

    if (!month || !Number.isFinite(strike)) continue;

    const expiration = `${month}/${day}/${year}`;
    const key = `${ticker}|${expiration}|${strike}|${optionType}`;

    quantities.set(key, (quantities.get(key) || 0) + qty);
    prices.set(key, { pricePaid: Math.abs(pricePaid) });
  }

  return { quantities, prices };
}

/**
 * Validates spread positions against E*Trade portfolio option quantities.
 * Returns a list of discrepancies.
 *
 * @param {Object[]} spreads - Array of spread positions from transaction parsing
 * @param {Map<string, number>} portfolioOptions - Map of option key to quantity from portfolio CSV
 * @returns {{ mismatches: Array, missing: Array, extra: Array }}
 */
function validateOptionQuantities_(spreads, portfolioOptions) {
  // Build expected quantities from spreads
  const expected = new Map();

  for (const spread of spreads) {
    if (spread.type === "stock" || spread.type === "cash") continue;

    const ticker = spread.ticker;
    const expiration = formatExpirationForKey_(spread.expiration);

    if (spread.type === "iron-condor" && spread.legs) {
      for (const leg of spread.legs) {
        const key = `${ticker}|${expiration}|${leg.strike}|${leg.optionType}`;
        expected.set(key, (expected.get(key) || 0) + leg.qty);
      }
    } else {
      // Regular spread (2 legs or single leg)
      if (spread.lowerStrike != null) {
        const key = `${ticker}|${expiration}|${spread.lowerStrike}|${spread.optionType}`;
        expected.set(key, (expected.get(key) || 0) + spread.qty);
      }
      if (spread.upperStrike != null) {
        const key = `${ticker}|${expiration}|${spread.upperStrike}|${spread.optionType}`;
        expected.set(key, (expected.get(key) || 0) - spread.qty); // Short leg
      }
    }
  }

  const mismatches = [];
  const missing = [];
  const extra = [];

  // Compare expected vs portfolio
  for (const [key, expectedQty] of expected) {
    const portfolioQty = portfolioOptions.get(key) || 0;
    if (expectedQty !== portfolioQty) {
      if (portfolioQty === 0) {
        extra.push({ key, expected: expectedQty, portfolio: 0 });
      } else {
        mismatches.push({ key, expected: expectedQty, portfolio: portfolioQty });
      }
    }
  }

  // Check for options in portfolio that we don't have
  for (const [key, portfolioQty] of portfolioOptions) {
    if (!expected.has(key) && portfolioQty !== 0) {
      missing.push({ key, expected: 0, portfolio: portfolioQty });
    }
  }

  return { mismatches, missing, extra };
}

/**
 * Builds a map of ticker -> latest transaction date from stock transactions.
 *
 * @param {Object[]} stockTxns - Array of stock transactions
 * @returns {Map<string, Date>} Map of ticker to latest transaction date
 */
function buildLatestStockDates_(stockTxns) {
  const latestByTicker = new Map();

  if (!stockTxns) return latestByTicker;

  for (const txn of stockTxns) {
    const ticker = txn.ticker;
    if (!ticker) continue;

    const txnDate = parseDateAtMidnight_(txn.date);
    if (!txnDate) continue;

    const existing = latestByTicker.get(ticker);
    if (!existing || txnDate > existing) {
      latestByTicker.set(ticker, txnDate);
    }
  }

  return latestByTicker;
}

/**
 * Finds all files in a folder whose name starts with prefix, sorted newest first by name.
 */
function findFilesByPrefix_(folder, prefix) {
  log.debug("files", `Looking for '${prefix}' in folder '${folder.getName()}'`);

  // Try CSV MIME type first
  let iter = folder.searchFiles(`title contains '${prefix}' and mimeType = 'text/csv'`);
  let files = [];
  while (iter.hasNext()) files.push(iter.next());
  log.debug("files", `CSV MIME type search found: ${files.length} files`);

  // If no CSV files found, try any file with the prefix (MIME type might be wrong)
  if (files.length === 0) {
    iter = folder.searchFiles(`title contains '${prefix}'`);
    while (iter.hasNext()) {
      const f = iter.next();
      log.debug("files", `Found file: ${f.getName()} (MIME: ${f.getMimeType()})`);
      // Include CSV files and other text formats (e.g., .cindy)
      const name = f.getName().toLowerCase();
      if (name.endsWith('.csv') || name.endsWith('.cindy') || name.endsWith('.txt')) {
        files.push(f);
      }
    }
    log.debug("files", `After extension filter: ${files.length} files`);
  }

  // If still nothing, list all files in folder for debugging
  if (files.length === 0) {
    log.warn("files", `No files found. Listing all files in folder:`);
    const allFiles = folder.getFiles();
    while (allFiles.hasNext()) {
      const f = allFiles.next();
      log.debug("files", `  - ${f.getName()} (MIME: ${f.getMimeType()})`);
    }
  }

  // If we have real files, exclude -sample files
  const realFiles = files.filter(f => !f.getName().toLowerCase().includes("-sample"));
  const result = realFiles.length > 0 ? realFiles : files;

  result.sort((a, b) => b.getName().localeCompare(a.getName()));
  log.debug("files", `Returning ${result.length} files`);
  return result;
}

/**
 * Finds all files in a folder with an exact name.
 */
function findFilesByName_(folder, name) {
  const iter = folder.getFilesByName(name);
  const files = [];
  while (iter.hasNext()) files.push(iter.next());
  return files;
}

/**
 * Parses E*Trade CSV content into transaction objects.
 * Returns { transactions, stockTxns }.
 * transactions: option opens, closes, exercises, assignments
 * stockTxns: stock Bought/Sold for matching exercise/assignment to market price
 */
function parseEtradeTransactionsFromCsv_(csvContent) {
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

  // New-format types: "Bought"/"Sold" with OPENING/CLOSING in Description
  const newFormatTypes = ["Bought", "Sold"];

  // Find Description column index from header
  const headerCols = parseCsvLine_(lines[headerIdx]);
  const descIdx = headerCols.findIndex(h => h.trim().toLowerCase() === "description");

  // Parse rows
  for (let i = headerIdx + 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    const cols = parseCsvLine_(line);
    if (cols.length < 8) continue;

    const [dateStr, txnType, secType, symbol, qtyStr, amountStr, priceStr, commStr] = cols;
    const desc = descIdx >= 0 && cols.length > descIdx ? cols[descIdx] : "";

    // Stock transactions (for exercise/assignment matching and portfolio aggregation)
    if (secType === "EQ" && (txnType === "Bought" || txnType === "Sold")) {
      stockTxns.push({
        date: dateStr,
        txnType: txnType,
        ticker: symbol.trim().toUpperCase(),
        qty: parseFloat(qtyStr) || 0,
        price: parseFloat(priceStr) || 0,
      });
      continue;
    }

    // Option transactions
    if (secType !== "OPTN") continue;

    // Determine if old format or new format
    const isOldFormat = optionTypes.includes(txnType);
    const isNewFormat = newFormatTypes.includes(txnType);
    if (!isOldFormat && !isNewFormat) continue;

    // Parse option symbol: try human-readable first, then OCC format
    const parsed = parseEtradeOptionSymbol_(symbol) || parseOccOptionSymbol_(symbol);
    if (!parsed) continue;

    // Determine open/close/exercise/assigned
    let isOpen, isClosed, isExercised, isAssigned;
    if (isOldFormat) {
      isOpen = txnType === "Bought To Open" || txnType === "Sold Short";
      isClosed = txnType === "Sold To Close" || txnType === "Bought To Cover";
      isExercised = txnType === "Option Exercised";
      isAssigned = txnType === "Option Assigned";
    } else {
      // New format: check Description for OPENING/CLOSING
      const descUpper = desc.toUpperCase();
      isOpen = descUpper.includes("OPENING");
      isClosed = descUpper.includes("CLOSING");
      isExercised = descUpper.includes("EXERCIS");
      isAssigned = descUpper.includes("ASSIGN");
    }

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
      isOpen,
      isClosed,
      isExercised,
      isAssigned,
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
 * Converts an expiration (Date or string) to M/D/YYYY format for consistent key matching.
 */
function formatExpirationForKey_(exp) {
  // Normalize all dates to M/D/YYYY format for consistent key matching
  if (exp instanceof Date) {
    return `${exp.getMonth() + 1}/${exp.getDate()}/${exp.getFullYear()}`;
  }
  const s = String(exp || "").trim();
  // Already in M/D/YYYY format
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) return s;
  // YYYY-MM-DD → M/D/YYYY
  const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    return `${parseInt(isoMatch[2], 10)}/${parseInt(isoMatch[3], 10)}/${isoMatch[1]}`;
  }
  return s;
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
  const expiration = `${monthNum}/${parseInt(day, 10)}/${fullYear}`;

  return {
    ticker: ticker.toUpperCase(),
    expiration,
    strike: parseFloat(strike),
    type: type.charAt(0).toUpperCase() + type.slice(1).toLowerCase(), // "Call" or "Put"
  };
}

/**
 * Parses OCC-format option symbol like "TSLA--281215C00500000"
 * Format: TICKER + padding + YYMMDD + C/P + 8-digit strike (price * 1000)
 * Returns { ticker, expiration, strike, type } or null
 */
function parseOccOptionSymbol_(symbol) {
  const match = symbol.match(/^([A-Z]+)\W*(\d{6})([CP])(\d{8})$/i);
  if (!match) return null;

  const [, ticker, dateStr, typeChar, strikeStr] = match;

  const yy = parseInt(dateStr.slice(0, 2), 10);
  const mm = parseInt(dateStr.slice(2, 4), 10);
  const dd = parseInt(dateStr.slice(4, 6), 10);
  const fullYear = 2000 + yy;

  const strike = parseInt(strikeStr, 10) / 1000;
  const type = typeChar.toUpperCase() === "C" ? "Call" : "Put";

  return {
    ticker: ticker.toUpperCase(),
    expiration: `${mm}/${dd}/${fullYear}`,
    strike,
    type,
  };
}

/**
 * Pairs consecutive opens on same date into spread orders.
 * Detects iron condors (2 puts + 2 calls with matching qty).
 * Applies closes to reduce open quantities (FIFO by date).
 */
function pairTransactionsIntoSpreads_(transactions) {
  const spreads = [];

  // Group opens by date + ticker + expiration
  // Keep original quantities - closed positions will be marked via the Closed column
  const groups = new Map();
  for (const txn of transactions) {
    if (!txn.isOpen) continue;

    const key = `${txn.date}|${txn.ticker}|${txn.expiration}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ ...txn }); // Clone to avoid mutating original
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

    // Check for long straddle/strangle: 1 long call + 1 long put (no shorts)
    if (longCalls.length >= 1 && longPuts.length >= 1 &&
        shortCalls.length === 0 && shortPuts.length === 0) {
      const lc = longCalls[0];
      const lp = longPuts[0];
      const pairQty = Math.min(lc.qty, lp.qty);
      const isStraddle = lc.strike === lp.strike;

      spreads.push({
        type: isStraddle ? "long-straddle" : "long-strangle",
        ticker: lc.ticker,
        expiration: lc.expiration,
        qty: pairQty,
        date: lc.date,
        legs: [
          { strike: lp.strike, optionType: "Put", qty: pairQty, price: lp.price },
          { strike: lc.strike, optionType: "Call", qty: pairQty, price: lc.price },
        ].sort((a, b) => a.strike - b.strike),
      });

      // Reduce quantities for any remainder
      lc.qty -= pairQty;
      lp.qty -= pairQty;

      // If fully consumed, skip to next group
      if (lc.qty === 0 && lp.qty === 0 && longCalls.length === 1 && longPuts.length === 1) {
        continue;
      }
    }

    // Check for short straddle/strangle: 1 short call + 1 short put (no longs)
    if (shortCalls.length >= 1 && shortPuts.length >= 1 &&
        longCalls.length === 0 && longPuts.length === 0) {
      const sc = shortCalls[0];
      const sp = shortPuts[0];
      const pairQty = Math.min(Math.abs(sc.qty), Math.abs(sp.qty));
      const isStraddle = sc.strike === sp.strike;

      spreads.push({
        type: isStraddle ? "short-straddle" : "short-strangle",
        ticker: sc.ticker,
        expiration: sc.expiration,
        qty: pairQty,
        date: sc.date,
        legs: [
          { strike: sp.strike, optionType: "Put", qty: -pairQty, price: sp.price },
          { strike: sc.strike, optionType: "Call", qty: -pairQty, price: sc.price },
        ].sort((a, b) => a.strike - b.strike),
      });

      // Reduce quantities for any remainder
      sc.qty += pairQty; // negative + positive = less negative
      sp.qty += pairQty;

      // If fully consumed, skip to next group
      if (sc.qty === 0 && sp.qty === 0 && shortCalls.length === 1 && shortPuts.length === 1) {
        continue;
      }
    }

    // Normal pairing: pair calls with calls, puts with puts
    // Pairs what it can into 2-leg spreads, leaves remainder as single legs
    // Custom position detection happens in Parsing.js when reading the Positions sheet
    for (const optionType of ["Call", "Put"]) {
      const legsOfType = txns.filter(t => t.optionType === optionType);
      if (legsOfType.length === 0) continue;

      const longs = legsOfType.filter(t => t.qty > 0);
      const shorts = legsOfType.filter(t => t.qty < 0);

      // Clone and sort for pairing
      const longsToProcess = longs.map(t => ({ ...t })).sort((a, b) => a.strike - b.strike);
      const shortsToProcess = shorts.map(t => ({ ...t })).sort((a, b) => a.strike - b.strike);

      // Pair by matching quantities
      let li = 0, si = 0;
      while (li < longsToProcess.length && si < shortsToProcess.length) {
        const long = longsToProcess[li];
        const short = shortsToProcess[si];

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
      while (li < longsToProcess.length) {
        const long = longsToProcess[li];
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
      while (si < shortsToProcess.length) {
        const short = shortsToProcess[si];
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
 * Finds custom groups in a Portfolio sheet (groups that don't match standard strategies).
 * Returns array of { group, description } for each custom group found.
 */
function findCustomGroups_(sheet) {
  const customGroups = [];
  const range = sheet.getDataRange();
  const rows = range.getValues();
  if (rows.length < 2) return customGroups;

  const headers = rows[0];
  const idxSym = findColumn_(headers, ["symbol", "ticker"]);
  const idxGroup = findColumn_(headers, ["group", "grp"]);
  const idxDesc = findColumn_(headers, ["description", "desc"]);
  const idxStrike = findColumn_(headers, ["strike", "strikeprice"]);
  const idxType = findColumn_(headers, ["type", "optiontype", "callput"]);
  const idxQty = findColumn_(headers, ["qty", "quantity"]);

  // Group rows by (symbol, group)
  const groups = new Map();
  let lastSym = "", lastGroup = "", lastDesc = "";

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const rawSym = idxSym >= 0 ? String(row[idxSym] || "").trim().toUpperCase() : "";
    if (rawSym) lastSym = rawSym;
    const rawGroup = idxGroup >= 0 ? String(row[idxGroup] || "").trim() : "";
    if (rawGroup) lastGroup = rawGroup;
    const rawDesc = idxDesc >= 0 ? String(row[idxDesc] || "").trim() : "";
    if (rawDesc) lastDesc = rawDesc;

    const groupKey = `${lastSym}|${lastGroup}`;
    if (!groups.has(groupKey)) {
      groups.set(groupKey, { legs: [], description: lastDesc, group: lastGroup });
    }

    const strike = idxStrike >= 0 ? parseNumber_(row[idxStrike]) : NaN;
    const type = idxType >= 0 ? parseOptionType_(row[idxType]) : null;
    const qty = idxQty >= 0 ? parseNumber_(row[idxQty]) : 0;

    if (Number.isFinite(qty) && qty !== 0) {
      groups.get(groupKey).legs.push({ strike, type, qty });
    }
  }

  // Check each group - if detectPositionType_ returns null, it's custom
  for (const [key, g] of groups) {
    if (g.legs.length === 0) continue;
    const posType = detectPositionType_(g.legs);
    if (posType === null) {
      // Build description from strikes if not provided
      let desc = g.description;
      if (!desc) {
        const strikes = g.legs
          .filter(l => Number.isFinite(l.strike))
          .sort((a, b) => a.strike - b.strike)
          .map(l => l.qty < 0 ? `-${l.strike}` : `${l.strike}`)
          .join('/');
        desc = `${strikes} custom`;
      }
      customGroups.push({ group: g.group, description: desc });
    }
  }

  return customGroups;
}

/**
 * Parses existing Portfolio table into position objects.
 */
function parsePortfolioTable_(rows) {
  if (rows.length < 2) return [];

  const headers = rows[0];
  const idxSym = findColumn_(headers, ["symbol", "ticker"]);
  const idxGroup = findColumn_(headers, ["group", "grp"]);
  const idxStrike = findColumn_(headers, ["strike", "strikeprice"]);
  const idxType = findColumn_(headers, ["type", "optiontype", "callput"]);
  const idxExp = findColumn_(headers, ["expiration", "exp"]);
  const idxQty = findColumn_(headers, ["qty", "quantity"]);
  const idxPrice = findColumn_(headers, ["price", "cost"]);
  const idxLastTxnDate = findColumn_(headers, ["lasttxndate", "last txn date"]);

  const positions = new Map(); // key -> { legs, groupNum, lastTxnDate }

  let lastSym = "";
  let lastGroup = "";
  let currentLegs = [];
  let currentLastTxnDate = "";

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];

    const rawSym = idxSym >= 0 ? String(row[idxSym] || "").trim().toUpperCase() : "";
    if (rawSym) lastSym = rawSym;

    const rawGroup = idxGroup >= 0 ? String(row[idxGroup] || "").trim() : "";
    if (rawGroup && rawGroup !== lastGroup) {
      // Save previous group
      if (currentLegs.length > 0) {
        const key = makeSpreadKey_(currentLegs);
        if (key) positions.set(key, { legs: currentLegs, groupNum: lastGroup, lastTxnDate: currentLastTxnDate });
      }
      lastGroup = rawGroup;
      currentLegs = [];
      // Read LastTxnDate from the first row of each new group
      if (idxLastTxnDate >= 0) {
        const rawDate = row[idxLastTxnDate];
        currentLastTxnDate = rawDate instanceof Date
          ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "MM/dd/yy")
          : String(rawDate || "").trim();
      } else {
        currentLastTxnDate = "";
      }
    }

    const strike = idxStrike >= 0 ? parseNumber_(row[idxStrike]) : NaN;
    const type = idxType >= 0 ? parseOptionType_(row[idxType]) : null;
    const exp = idxExp >= 0 ? row[idxExp] : "";
    const qty = idxQty >= 0 ? parseNumber_(row[idxQty]) : NaN;
    const price = idxPrice >= 0 ? parseNumber_(row[idxPrice]) : NaN;

    const isStock = type === "Stock" || (!Number.isFinite(strike) && Number.isFinite(qty) && !type);
    if ((Number.isFinite(strike) || isStock) && Number.isFinite(qty)) {
      currentLegs.push({
        symbol: lastSym,
        strike: Number.isFinite(strike) ? strike : null,
        type: isStock ? "Stock" : type,
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
    if (key) positions.set(key, { legs: currentLegs, groupNum: lastGroup, lastTxnDate: currentLastTxnDate });
  }

  return positions;
}

/**
 * Creates a unique key for a spread from its legs.
 */
function makeSpreadKey_(legs) {
  if (legs.length === 0) return null;

  const ticker = legs[0].symbol;

  // Cash positions
  if (legs.length === 1 && (legs[0].type === "Cash" || ticker === "CASH")) {
    return "CASH|CASH";
  }

  // Stock positions
  if (legs.length === 1 && (legs[0].type === "Stock" || !Number.isFinite(legs[0].strike))) {
    return `${ticker}|STOCK`;
  }

  const exp = normalizeExpiration_(legs[0].expiration) || legs[0].expiration;
  const strikes = legs.map(l => l.strike).sort((a, b) => a - b);

  // Detect iron-condor/iron-butterfly: 4 legs with both puts and calls
  const types = new Set(legs.map(l => l.type));
  if (legs.length === 4 && types.has("Put") && types.has("Call")) {
    return `${ticker}|${exp}|${strikes.join("/")}|IC`;
  }

  const type = legs[0].type || "Call";
  return `${ticker}|${exp}|${strikes.join("/")}|${type}`;
}

/**
 * Creates spread key from a spread order.
 */
function makeSpreadKeyFromOrder_(spread) {
  if (spread.type === "stock") {
    return `${spread.ticker}|STOCK`;
  }

  if (spread.type === "cash") {
    return "CASH|CASH";
  }

  const exp = normalizeExpiration_(spread.expiration) || spread.expiration;

  // Multi-leg strategies with spread.legs
  if (spread.legs && spread.legs.length > 0) {
    const strikes = spread.legs.map(l => l.strike).sort((a, b) => a - b);
    // Use type abbreviation for the key
    let typeKey;
    switch (spread.type) {
      case "iron-condor": typeKey = "IC"; break;
      case "long-straddle": typeKey = "LS"; break;
      case "short-straddle": typeKey = "SS"; break;
      case "long-strangle": typeKey = "LSg"; break;
      case "short-strangle": typeKey = "SSg"; break;
      default: typeKey = spread.type || "?";
    }
    return `${spread.ticker}|${exp}|${strikes.join("/")}|${typeKey}`;
  }

  const strikes = [spread.lowerStrike, spread.upperStrike].filter(s => s != null).sort((a, b) => a - b);
  return `${spread.ticker}|${exp}|${strikes.join("/")}|${spread.optionType}`;
}

/**
 * Pre-merges spreads with the same key, keeping the latest date and summing quantities.
 * This ensures that when multiple transactions create the same spread on different dates,
 * only one spread is created with the combined quantity and the latest date.
 */
function preMergeSpreads_(spreads) {
  const merged = new Map(); // key -> spread

  for (const spread of spreads) {
    const key = makeSpreadKeyFromOrder_(spread);

    if (!merged.has(key)) {
      // First occurrence - clone the spread
      merged.set(key, { ...spread });
    } else {
      // Merge into existing
      const existing = merged.get(key);

      // Keep the later date
      if (spread.date && (!existing.date || spread.date > existing.date)) {
        existing.date = spread.date;
      }

      // Sum quantities and compute weighted average prices
      if (spread.type === "stock") {
        const oldQty = existing.qty || 0;
        const newQty = spread.qty || 0;
        const totalQty = oldQty + newQty;
        if (totalQty !== 0) {
          existing.price = ((oldQty * (existing.price || 0)) + (newQty * (spread.price || 0))) / totalQty;
        }
        existing.qty = totalQty;
      } else if (spread.type === "iron-condor" && spread.legs) {
        // For iron condors, sum quantities (legs stay the same strikes)
        existing.qty = (existing.qty || 0) + (spread.qty || 0);
        // Could also merge leg prices but keeping it simple
      } else {
        // Regular spread - sum quantities and merge prices
        const oldQty = existing.qty || 0;
        const newQty = spread.qty || 0;
        const totalQty = oldQty + newQty;

        if (totalQty !== 0 && existing.lowerPrice !== undefined && spread.lowerPrice !== undefined) {
          existing.lowerPrice = ((oldQty * existing.lowerPrice) + (newQty * spread.lowerPrice)) / totalQty;
        }
        if (totalQty !== 0 && existing.upperPrice !== undefined && spread.upperPrice !== undefined) {
          existing.upperPrice = ((oldQty * existing.upperPrice) + (newQty * spread.upperPrice)) / totalQty;
        }
        existing.qty = totalQty;
      }
    }
  }

  return Array.from(merged.values());
}

/**
 * Merges new spreads into existing positions.
 * Skips spreads whose date is not newer than the group's LastTxnDate.
 * Returns { updatedLegs, newLegs, skippedCount }.
 */
function mergeSpreads_(existingPositions, newSpreads) {
  const updatedLegs = [];
  const newLegs = [];
  const processedKeys = new Set();
  let skippedCount = 0;

  for (const spread of newSpreads) {
    const key = makeSpreadKeyFromOrder_(spread);

    if (existingPositions.has(key)) {
      const existing = existingPositions.get(key);

      // Stock positions: add delta qty and update lastTxnDate
      if (spread.type === "stock") {
        // spread.qty is the DELTA (bought - sold) from new transactions
        if (spread.qty === 0 && !spread.date) {
          // No new transactions for this stock
          skippedCount++;
          continue;
        }

        // Update quantity by adding the delta
        const stockLeg = existing.legs[0];
        if (stockLeg) {
          stockLeg.qty += spread.qty;
          // Update price to most recent transaction price
          if (spread.price) stockLeg.price = spread.price;
        }

        // Update lastTxnDate
        if (spread.date) {
          existing.lastTxnDate = spread.date;
        }

        existing.updated = true;
        updatedCount++;
        for (const leg of existing.legs) {
          leg.updated = true;
          updatedLegs.push(leg);
        }
        continue;
      }

      // Cash positions: just update the amount
      if (spread.type === "cash") {
        const cashLeg = existing.legs[0];
        if (cashLeg) {
          cashLeg.price = spread.price;
          cashLeg.updated = true;
          updatedLegs.push(cashLeg);
        }
        continue;
      }

      // Per-group dedup: skip if spread's date is not newer than LastTxnDate
      // Parse dates carefully to handle MM/DD/YY vs MM/DD/YYYY formats
      const spreadDate = parseDateAtMidnight_(spread.date);
      const lastTxnDate = parseDateAtMidnight_(existing.lastTxnDate);

      if (spreadDate && lastTxnDate && spreadDate <= lastTxnDate) {
        skippedCount++;
        continue;
      }

      // Debug: capture why this wasn't skipped
      existing.debugReason = ` [txn:${spread.date} vs last:${existing.lastTxnDate}]`;

      // Merge into existing
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

      // Track the newest date for this group
      if (spread.date > (existing.lastTxnDate || "")) {
        existing.lastTxnDate = spread.date;
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

  return { updatedLegs, newLegs, skippedCount };
}

/**
 * Writes the Portfolio table back to the sheet.
 * @param {Map} [closingPrices] - Map of leg keys to closing prices
 */
function writePortfolioTable_(ss, headers, updatedLegs, newLegs, closingPrices) {
  closingPrices = closingPrices || new Map();

  const legsRange = getNamedRangeWithTableFallback_(ss, "Portfolio");
  if (!legsRange || headers.length === 0) {
    log.info("import", "Portfolio table not found, creating new one");
    // Create new Portfolio sheet
    let sheet = ss.getSheetByName("Portfolio");
    if (!sheet) sheet = ss.insertSheet("Portfolio");

    headers = ["Symbol", "Group", "Description", "Strategy", "Strike", "Type", "Expiration", "Qty", "Price", "Investment", "Rec Close", "Closed", "Gain", "Current Value", "LastTxnDate", "Link"];
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground("#93c47d"); // Green header
    headerRange.setFontWeight("bold");
    ss.setNamedRange("PortfolioTable", sheet.getRange("A:P"));

    // Add filter to the data range
    const filterRange = sheet.getRange("A:P");
    filterRange.createFilter();

    // Set wrap to clip
    sheet.getRange("A:P").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  const range = getNamedRangeWithTableFallback_(ss, "Portfolio");
  const sheet = range.getSheet();
  const startRow = range.getRow();
  const startCol = range.getColumn();

  // Find column indexes
  const idxSym = findColumn_(headers, ["symbol", "ticker"]);
  const idxGroup = findColumn_(headers, ["group", "grp"]);
  const idxDescription = findColumn_(headers, ["description", "desc"]);
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
  const idxCurrentValue = findColumn_(headers, ["currentvalue", "current value", "currvalue"]);
  const idxLastTxnDate = findColumn_(headers, ["lasttxndate", "last txn date"]);
  const idxLink = findColumn_(headers, ["link"]);

  // Column letters for formulas
  const colLetter = (idx) => String.fromCharCode(65 + idx);
  const symCol = idxSym >= 0 ? colLetter(idxSym) : "A";
  const descCol = idxDescription >= 0 ? colLetter(idxDescription) : "C";
  const stratCol = idxStrategy >= 0 ? colLetter(idxStrategy) : "D";
  const strikeCol = idxStrike >= 0 ? colLetter(idxStrike) : "E";
  const typeCol = idxType >= 0 ? colLetter(idxType) : "F";
  const expCol = idxExp >= 0 ? colLetter(idxExp) : "G";
  const qtyCol = idxQty >= 0 ? colLetter(idxQty) : "H";
  const priceCol = idxPrice >= 0 ? colLetter(idxPrice) : "I";
  const recCloseCol = idxRecClose >= 0 ? colLetter(idxRecClose) : "K";
  const closedCol = idxClosed >= 0 ? colLetter(idxClosed) : "L";

  // Update existing rows
  for (const pos of updatedLegs) {
    for (const leg of pos.legs) {
      if (leg.row != null) {
        const rowNum = startRow + leg.row;
        if (idxQty >= 0) sheet.getRange(rowNum, startCol + idxQty).setValue(leg.qty);
        if (idxPrice >= 0) sheet.getRange(rowNum, startCol + idxPrice).setValue(roundTo_(leg.price, 2));

        // Write closing price if available and not already filled
        if (idxClosed >= 0 && leg.symbol && leg.expiration && leg.strike && leg.type && leg.type !== "Stock") {
          const existingVal = sheet.getRange(rowNum, startCol + idxClosed).getValue();
          if (existingVal === "" || existingVal == null) {
            const expStr = formatExpirationForKey_(leg.expiration);
            const key = `${leg.symbol}|${expStr}|${leg.strike}|${leg.type}`;
            const closePrice = closingPrices.get(key);
            if (closePrice != null) {
              sheet.getRange(rowNum, startCol + idxClosed).setValue(closePrice);
            }
          }
        }
      }
    }
    // Update LastTxnDate on the first row of the group (format as MM/DD/YYYY)
    if (idxLastTxnDate >= 0 && pos.lastTxnDate && pos.legs.length > 0 && pos.legs[0].row != null) {
      sheet.getRange(startRow + pos.legs[0].row, startCol + idxLastTxnDate).setValue(formatDateLong_(pos.lastTxnDate));
    }
  }

  // Append new spreads
  if (newLegs.length > 0) {
    let lastRow = sheet.getLastRow();
    let nextGroup = 1;

    // Find max group number AND last data row (before summary rows)
    let lastDataRow = startRow; // header row
    if (idxGroup >= 0 && lastRow > startRow) {
      const groupData = sheet.getRange(startRow + 1, startCol + idxGroup, lastRow - startRow, 1).getValues();
      for (let i = 0; i < groupData.length; i++) {
        const g = parseInt(groupData[i][0], 10);
        if (Number.isFinite(g)) {
          if (g >= nextGroup) nextGroup = g + 1;
          lastDataRow = startRow + 1 + i; // This row has valid data
        }
      }
    }

    // Delete any existing summary rows (everything after last data row)
    if (lastRow > lastDataRow) {
      sheet.deleteRows(lastDataRow + 1, lastRow - lastDataRow);
    }

    // Reset lastRow to the actual last data row
    lastRow = lastDataRow;

    for (const spread of newLegs) {
      const rows = [];

      // Helper to get closing price for a leg
      const getClosingPrice = (ticker, expiration, strike, optionType) => {
        const key = `${ticker}|${expiration}|${strike}|${optionType}`;
        const val = closingPrices.get(key);
        return val != null ? val : "";
      };

      // Generate description for this spread (shown on every row for filtering)
      const spreadDescription = generateSpreadDescription_(spread);

      // Handle stock position (single row, no strike/expiration)
      if (spread.type === "stock") {
        const row = new Array(headers.length).fill("");
        if (idxSym >= 0) row[idxSym] = spread.ticker;
        if (idxGroup >= 0) row[idxGroup] = nextGroup;
        if (idxDescription >= 0) row[idxDescription] = spreadDescription;
        if (idxType >= 0) row[idxType] = "Stock";
        if (idxQty >= 0) row[idxQty] = spread.qty;
        if (idxPrice >= 0) row[idxPrice] = roundTo_(spread.price, 2);
        rows.push(row);
      }
      // Handle cash position (single row, no strike/expiration)
      else if (spread.type === "cash") {
        const row = new Array(headers.length).fill("");
        if (idxSym >= 0) row[idxSym] = "CASH";
        if (idxGroup >= 0) row[idxGroup] = nextGroup;
        if (idxDescription >= 0) row[idxDescription] = "Cash";
        if (idxStrategy >= 0) row[idxStrategy] = "Cash";
        if (idxType >= 0) row[idxType] = "Cash";
        if (idxQty >= 0) row[idxQty] = 1;
        if (idxPrice >= 0) row[idxPrice] = roundTo_(spread.price, 2);
        rows.push(row);
      }
      // Handle iron condor (4 legs), straddle/strangle (2 legs), or custom multi-leg
      else if (spread.legs && spread.legs.length > 0) {
        for (let i = 0; i < spread.legs.length; i++) {
          const leg = spread.legs[i];
          const row = new Array(headers.length).fill("");
          if (i === 0) {
            // First row gets symbol, group, and description (formula set later)
            if (idxSym >= 0) row[idxSym] = spread.ticker;
            if (idxGroup >= 0) row[idxGroup] = nextGroup;
          }
          // Description only on first row (will be replaced by formula later)
          // Leave other rows blank so they don't show stale data
          if (idxStrike >= 0) row[idxStrike] = leg.strike;
          if (idxType >= 0) row[idxType] = leg.optionType;
          if (idxExp >= 0) row[idxExp] = formatExpirationForKey_(spread.expiration);
          if (idxQty >= 0) row[idxQty] = leg.qty;
          if (idxPrice >= 0) row[idxPrice] = roundTo_(leg.price, 2);
          if (idxClosed >= 0) row[idxClosed] = getClosingPrice(spread.ticker, spread.expiration, leg.strike, leg.optionType);
          rows.push(row);
        }
      } else {
        // Handle 2-leg spread or single leg
        const hasLong = spread.lowerStrike != null;
        const hasShort = spread.upperStrike != null;
        if (!hasLong && !hasShort) continue;
        const isFirstRow = [true]; // track whether next row gets symbol/group
        const normalizedExp = formatExpirationForKey_(spread.expiration);

        // Long leg
        if (hasLong) {
          const row = new Array(headers.length).fill("");
          if (idxSym >= 0) row[idxSym] = spread.ticker;
          if (idxGroup >= 0) row[idxGroup] = nextGroup;
          if (idxDescription >= 0) row[idxDescription] = spreadDescription;
          if (idxStrike >= 0) row[idxStrike] = spread.lowerStrike;
          if (idxType >= 0) row[idxType] = spread.optionType;
          if (idxExp >= 0) row[idxExp] = normalizedExp;
          if (idxQty >= 0) row[idxQty] = spread.qty;
          if (idxPrice >= 0) row[idxPrice] = roundTo_(spread.lowerPrice, 2);
          if (idxClosed >= 0) row[idxClosed] = getClosingPrice(spread.ticker, spread.expiration, spread.lowerStrike, spread.optionType);
          rows.push(row);
          isFirstRow[0] = false;
        }

        // Short leg
        if (hasShort) {
          const row = new Array(headers.length).fill("");
          if (isFirstRow[0]) {
            // Naked short — this is the first (only) row, needs symbol/group
            if (idxSym >= 0) row[idxSym] = spread.ticker;
            if (idxGroup >= 0) row[idxGroup] = nextGroup;
            // Description only on first row (formula set later for multi-leg)
            if (idxDescription >= 0) row[idxDescription] = spreadDescription;
          }
          // Leave description blank on second row of 2-leg spreads
          if (idxStrike >= 0) row[idxStrike] = spread.upperStrike;
          if (idxType >= 0) row[idxType] = spread.optionType;
          if (idxExp >= 0) row[idxExp] = normalizedExp;
          // For 2-leg spread, qty is positive (long side) so negate for short.
          // For naked short (no long leg), qty is already negative — use directly.
          if (idxQty >= 0) row[idxQty] = hasLong ? -spread.qty : spread.qty;
          if (idxPrice >= 0) row[idxPrice] = roundTo_(spread.upperPrice, 2);
          if (idxClosed >= 0) row[idxClosed] = getClosingPrice(spread.ticker, spread.expiration, spread.upperStrike, spread.optionType);
          rows.push(row);
        }
      }

      if (rows.length === 0) continue;

      // Set LastTxnDate on the first row of the group (format as MM/DD/YYYY)
      if (idxLastTxnDate >= 0 && spread.date) {
        rows[0][idxLastTxnDate] = formatDateLong_(spread.date);
      }

      const firstRow = lastRow + 1;
      const lastLegRow = firstRow + rows.length - 1;

      // Write data rows
      sheet.getRange(firstRow, startCol, rows.length, headers.length).setValues(rows);

      // Write formulas on first row of group
      const isStock = spread.type === "stock";
      const isCash = spread.type === "cash";
      const rangeStr = `${firstRow}:${lastLegRow}`;

      // Strategy formula
      if (idxStrategy >= 0) {
        if (isStock) {
          sheet.getRange(firstRow, startCol + idxStrategy).setValue("Stock");
        } else if (isCash) {
          sheet.getRange(firstRow, startCol + idxStrategy).setValue("Cash");
        } else {
          const formula = `=detectStrategy($${strikeCol}${firstRow}:$${strikeCol}${lastLegRow}, $${typeCol}${firstRow}:$${typeCol}${lastLegRow}, $${qtyCol}${firstRow}:$${qtyCol}${lastLegRow})`;
          sheet.getRange(firstRow, startCol + idxStrategy).setFormula(formula);
        }
      }

      // Description formula (strikes with -prefix for shorts + strategy name)
      if (idxDescription >= 0) {
        if (isStock) {
          sheet.getRange(firstRow, startCol + idxDescription).setValue("Stock");
        } else if (isCash) {
          sheet.getRange(firstRow, startCol + idxDescription).setValue("Cash");
        } else {
          // Include Strategy column as suffix: "500/-600 bull-call-spread"
          const formula = `=formatLegsDescription($${strikeCol}${firstRow}:$${strikeCol}${lastLegRow}, $${qtyCol}${firstRow}:$${qtyCol}${lastLegRow}, $${stratCol}${firstRow})`;
          sheet.getRange(firstRow, startCol + idxDescription).setFormula(formula);
        }
      }

      // Investment formula (stocks and cash: no *100 multiplier)
      if (idxInvestment >= 0) {
        if (isCash) {
          // Cash: investment = price (the cash amount)
          sheet.getRange(firstRow, startCol + idxInvestment).setFormula(`=$${priceCol}${firstRow}`);
        } else {
          const multiplier = isStock ? "" : " * 100";
          const formula = `=SUMPRODUCT($${qtyCol}${firstRow}:$${qtyCol}${lastLegRow}, $${priceCol}${firstRow}:$${priceCol}${lastLegRow})${multiplier}`;
          sheet.getRange(firstRow, startCol + idxInvestment).setFormula(formula);
        }
      }

      // Gain formula: use Closed if available, otherwise Rec Close
      // Cash: gain is always 0, Stocks: no *100 multiplier
      if (idxGain >= 0 && idxRecClose >= 0) {
        if (isCash) {
          // Cash doesn't gain or lose value
          sheet.getRange(firstRow, startCol + idxGain).setValue(0);
        } else {
          let closeRef;
          if (idxClosed >= 0) {
            closeRef = `IF($${closedCol}${firstRow}:$${closedCol}${lastLegRow}<>"", $${closedCol}${firstRow}:$${closedCol}${lastLegRow}, $${recCloseCol}${firstRow}:$${recCloseCol}${lastLegRow})`;
          } else {
            closeRef = `$${recCloseCol}${firstRow}:$${recCloseCol}${lastLegRow}`;
          }
          const multiplier = isStock ? "" : " * 100";
          const formula = `=SUMPRODUCT($${qtyCol}${firstRow}:$${qtyCol}${lastLegRow}, ${closeRef} - $${priceCol}${firstRow}:$${priceCol}${lastLegRow})${multiplier}`;
          sheet.getRange(firstRow, startCol + idxGain).setFormula(formula);
        }
      }

      // Current Value formula: blank if Gain is blank, 0 if closed, otherwise Investment + Gain
      // Only the first row of a spread has Gain, so only first row gets Current Value
      // Cash: current value = investment (no gain/loss)
      if (idxCurrentValue >= 0 && idxInvestment >= 0 && idxGain >= 0) {
        if (isCash) {
          // Cash current value = investment
          const invCell = colLetter(idxInvestment) + firstRow;
          sheet.getRange(firstRow, startCol + idxCurrentValue).setFormula(`=$${invCell}`);
        } else {
          const invCell = colLetter(idxInvestment) + firstRow;
          const gainCell = colLetter(idxGain) + firstRow;
          let formula;
          if (idxClosed >= 0) {
            formula = `=IF($${gainCell}="", "", IF(COUNTBLANK($${closedCol}${firstRow}:$${closedCol}${lastLegRow})=0, 0, $${invCell}+$${gainCell}))`;
          } else {
            formula = `=IF($${gainCell}="", "", $${invCell}+$${gainCell})`;
          }
          sheet.getRange(firstRow, startCol + idxCurrentValue).setFormula(formula);
        }
      }

      // Link formula: HYPERLINK with "OptionStrat" display text (skip for stocks and cash)
      if (idxLink >= 0 && !isStock && !isCash) {
        const urlFormula = `buildOptionStratUrlFromLegs($${symCol}$1:$${symCol}${firstRow}, $${strikeCol}${firstRow}:$${strikeCol}${lastLegRow}, $${typeCol}${firstRow}:$${typeCol}${lastLegRow}, $${expCol}${firstRow}:$${expCol}${lastLegRow}, $${qtyCol}${firstRow}:$${qtyCol}${lastLegRow}, $${priceCol}${firstRow}:$${priceCol}${lastLegRow})`;
        const formula = `=HYPERLINK(${urlFormula}, "OptionStrat")`;
        sheet.getRange(firstRow, startCol + idxLink).setFormula(formula);
      }

      // Rec Close formula for each leg
      if (idxRecClose >= 0) {
        for (let i = 0; i < rows.length; i++) {
          const legRow = firstRow + i;
          const hasClosed = idxClosed >= 0 && rows[i][idxClosed] !== "";
          if (!hasClosed) {
            if (isCash) {
              // Cash: Rec Close = Price (cash value doesn't change)
              sheet.getRange(legRow, startCol + idxRecClose).setFormula(`=$${priceCol}${legRow}`);
            } else if (isStock) {
              // Use GOOGLEFINANCE for stock current price
              const formula = `=GOOGLEFINANCE("${spread.ticker}")`;
              sheet.getRange(legRow, startCol + idxRecClose).setFormula(formula);
            } else {
              const formula = `=recommendClose($${symCol}$1:$${symCol}${legRow}, $${expCol}${legRow}, $${strikeCol}${legRow}, $${typeCol}${legRow}, $${qtyCol}${legRow}, 60)`;
              sheet.getRange(legRow, startCol + idxRecClose).setFormula(formula);
            }
          }
        }
      }

      // Check if all legs are closed
      const allClosed = idxClosed >= 0 && rows.every(r => r[idxClosed] !== "");

      // Alternate group colors: odd = pale yellow, even = white
      const bgColor = (nextGroup % 2 === 1) ? "#fff2cc" : "#ffffff";
      const groupRange = sheet.getRange(firstRow, startCol, rows.length, headers.length);
      groupRange.setBackground(bgColor);

      // Dim closed positions: light gray text
      if (allClosed) {
        groupRange.setFontColor("#999999");
      }

      lastRow = lastLegRow;
      nextGroup++;
    }

    // Write summary rows after all data
    const summaryStart = lastRow + 2; // blank row then summary
    const invCol = idxInvestment >= 0 ? colLetter(idxInvestment) : "I";
    const gainCol = idxGain >= 0 ? colLetter(idxGain) : "L";
    const dr = (col) => `$${col}$2:$${col}$${lastRow}`; // proper range reference

    // Realized row: gain for closed positions
    sheet.getRange(summaryStart, startCol).setValue("Realized").setFontWeight("bold");
    sheet.getRange(summaryStart, startCol + idxGain)
      .setFormula(`=SUMPRODUCT((${dr(closedCol)}<>"")*${dr(gainCol)})`)
      .setFontWeight("bold");

    // Unrealized row: gain for open positions
    sheet.getRange(summaryStart + 1, startCol).setValue("Unrealized").setFontWeight("bold");
    sheet.getRange(summaryStart + 1, startCol + idxGain)
      .setFormula(`=SUMPRODUCT((${dr(closedCol)}="")*${dr(gainCol)})`)
      .setFontWeight("bold");

    // Total row: open investment + total gain
    sheet.getRange(summaryStart + 2, startCol).setValue("Total").setFontWeight("bold");
    sheet.getRange(summaryStart + 2, startCol + idxInvestment)
      .setFormula(`=SUMPRODUCT((${dr(closedCol)}="")*${dr(invCol)})`)
      .setFontWeight("bold");
    sheet.getRange(summaryStart + 2, startCol + idxGain)
      .setFormula(`=$${gainCol}$${summaryStart}+$${gainCol}$${summaryStart + 1}`)
      .setFontWeight("bold");

    // Current Value summary (only for open positions)
    if (idxCurrentValue >= 0) {
      const currValCol = colLetter(idxCurrentValue);
      sheet.getRange(summaryStart + 2, startCol + idxCurrentValue)
        .setFormula(`=SUM(${dr(currValCol)})`)
        .setFontWeight("bold");
    }

    // Format summary area
    const summaryRange = sheet.getRange(summaryStart, startCol, 3, headers.length);
    summaryRange.setBackground("#d9ead3"); // light green
  }

  // Apply number formats to all data rows (always, even with no new legs)
  const lastDataRow = sheet.getLastRow();
  if (lastDataRow >= 2) {
    const dataRowCount = lastDataRow - 1; // rows 2 through lastDataRow
    const fmtCols = [
      { idx: idxQty, fmt: "#,##0" },
      { idx: idxPrice, fmt: "#,##0.00" },
      { idx: idxInvestment, fmt: "#,##0.00" },
      { idx: idxRecClose, fmt: "#,##0.00" },
      { idx: idxClosed, fmt: "#,##0.00" },
      { idx: idxGain, fmt: "#,##0.00" },
      { idx: idxCurrentValue, fmt: "#,##0.00" },
      { idx: idxLastTxnDate, fmt: "mm/dd/yy" },
    ];
    for (const { idx, fmt } of fmtCols) {
      if (idx >= 0) {
        sheet.getRange(2, startCol + idx, dataRowCount, 1).setNumberFormat(fmt);
      }
    }

    // Auto-fit column widths to data (not headers) and wrap headers
    const headerRange = sheet.getRange(1, startCol, 1, headers.length);
    const headerValues = headerRange.getValues()[0];

    // Temporarily clear headers, auto-resize, then restore
    headerRange.clearContent();
    for (let c = 0; c < headers.length; c++) {
      sheet.autoResizeColumn(startCol + c);
    }
    headerRange.setValues([headerValues]);

    // Wrap headers and set vertical alignment to bottom
    headerRange.setWrap(true).setVerticalAlignment("bottom");

    // Set a minimum width for very narrow columns
    for (let c = 0; c < headers.length; c++) {
      const width = sheet.getColumnWidth(startCol + c);
      if (width < 50) {
        sheet.setColumnWidth(startCol + c, 50);
      }
    }
  }
}

/**
 * Aggregates stock transactions into net positions by symbol.
 * Bought transactions add to quantity, Sold transactions subtract.
 * Tracks the latest transaction date for each symbol.
 *
 * @param {Object[]} stockTxns - Array of stock transaction objects from parseEtradeTransactionsFromCsv_
 * @returns {Object[]} Array of stock position objects (spread-like with type="stock")
 */
/**
 * Aggregates stock transactions into delta positions by symbol.
 * Bought transactions add to quantity, Sold transactions subtract.
 *
 * @param {Object[]} stockTxns - Array of stock transaction objects from parseEtradeTransactionsFromCsv_
 * @param {Map<string, Date>} [sinceByTicker] - Optional map of ticker -> cutoff date.
 *        Only transactions AFTER this date are included for each ticker.
 * @returns {Object[]} Array of stock position objects (spread-like with type="stock")
 */
function aggregateStockTransactions_(stockTxns, sinceByTicker) {
  if (!stockTxns || stockTxns.length === 0) return [];

  // Group by ticker: { ticker: { qty: number, lastDate: Date, lastPrice: number } }
  const byTicker = new Map();

  for (const txn of stockTxns) {
    const ticker = txn.ticker;
    if (!ticker) continue;

    // Parse transaction date
    const txnDate = parseDateAtMidnight_(txn.date);

    // If sinceByTicker provided, skip transactions on or before the cutoff date
    if (sinceByTicker && sinceByTicker.has(ticker)) {
      const cutoff = sinceByTicker.get(ticker);
      if (txnDate && cutoff && txnDate <= cutoff) {
        continue; // Skip this transaction - it's already been processed
      }
    }

    // Get or initialize ticker entry
    if (!byTicker.has(ticker)) {
      byTicker.set(ticker, {
        qty: 0,
        lastDate: null,
        lastPrice: 0,
      });
    }

    const entry = byTicker.get(ticker);

    // Accumulate quantity: Bought adds, Sold subtracts
    const qtyChange = txn.txnType === "Bought" ? txn.qty : -txn.qty;
    entry.qty += qtyChange;

    // Track latest transaction date
    if (txnDate && (!entry.lastDate || txnDate > entry.lastDate)) {
      entry.lastDate = txnDate;
      entry.lastPrice = txn.price; // Use price from most recent transaction
    }
  }

  // Convert to spread-like objects
  const stocks = [];
  for (const [ticker, entry] of byTicker) {
    // Skip if no new transactions after cutoff (qty would be 0 and no lastDate)
    if (entry.qty === 0 && !entry.lastDate) continue;

    stocks.push({
      type: "stock",
      ticker: ticker,
      qty: entry.qty,  // This is the DELTA qty (change from new transactions)
      price: entry.lastPrice,
      date: entry.lastDate,
      expiration: null,
      lowerStrike: null,
      upperStrike: null,
      optionType: "Stock",
    });
  }

  return stocks;
}

/* =========================================================
   File Upload Dialogs
   ========================================================= */

/**
 * Shows file upload dialog for full portfolio rebuild.
 */
function showUploadRebuildDialog() {
  const html = HtmlService.createHtmlOutputFromFile("FileUpload")
    .setWidth(500)
    .setHeight(450);
  const content = html.getContent().replace(
    "if (mode) init(mode);",
    "init('rebuildPortfolio');"
  );
  const output = HtmlService.createHtmlOutput(content)
    .setWidth(500)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(output, "Upload E*Trade Files");
}

/**
 * Strips browser-added disambiguators like "(1)", " (1)", "(2)" from filenames.
 * E.g., "PortfolioDownload (1).csv" -> "PortfolioDownload.csv"
 */
function stripBrowserDisambiguator_(fileName) {
  // Remove patterns like " (1)", "(1)", " (2)", etc. before the extension
  return fileName.replace(/\s*\(\d+\)(?=\.[^.]+$|$)/, "");
}

/**
 * Handles uploaded files for portfolio rebuild.
 * Saves files to Drive and rebuilds portfolio.
 * @param {{name: string, content: string}|null} portfolio - Portfolio CSV (optional)
 * @param {Array<{name: string, content: string}>} transactions - Transaction CSV(s) (optional)
 * @returns {string} Status message
 */
function uploadAndRebuildPortfolio(portfolio, transactions) {
  if (!portfolio && (!transactions || transactions.length === 0)) {
    throw new Error("At least one file (Portfolio or Transactions) is required");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folder = getOrCreateEtradeFolder_(ss);

  const uploadedParts = [];

  // Save portfolio file with canonical name (always replace - it's current state)
  if (portfolio) {
    saveFileToFolder_(folder, "PortfolioDownload.csv", portfolio.content);
    uploadedParts.push("portfolio");
  }

  // Save transaction files with unique names (preserves history)
  if (transactions && transactions.length > 0) {
    for (const txn of transactions) {
      const cleanName = stripBrowserDisambiguator_(txn.name);
      saveFileWithUniqueName_(folder, cleanName, txn.content);
    }
    uploadedParts.push(`${transactions.length} transaction file(s)`);
  }

  // Run rebuild
  rebuildPortfolio();

  return `Uploaded ${uploadedParts.join(" and ")}. Rebuilt portfolio.`;
}

/**
 * Gets or creates the E*Trade data folder.
 */
function getOrCreateEtradeFolder_(ss) {
  const dataFolderPath = getConfigValue_(ss, "DataFolder", "SpreadFinder/DATA") + "/Etrade";
  let folder = DriveApp.getRootFolder();
  for (const part of dataFolderPath.split("/").filter(p => p.trim())) {
    const it = folder.getFoldersByName(part);
    if (it.hasNext()) {
      folder = it.next();
    } else {
      folder = folder.createFolder(part);
    }
  }
  return folder;
}

/**
 * Saves a file to a folder, replacing existing file with same name.
 */
function saveFileToFolder_(folder, fileName, content) {
  const existing = folder.getFilesByName(fileName);
  if (existing.hasNext()) {
    existing.next().setTrashed(true);
  }
  folder.createFile(fileName, content, MimeType.CSV);
}

/**
 * Saves a file to a folder with a unique name.
 * If a file with the same name exists, appends a timestamp to make it unique.
 * Returns the actual filename used.
 */
function saveFileWithUniqueName_(folder, fileName, content) {
  // Check if file with same name exists
  const existing = folder.getFilesByName(fileName);
  let finalName = fileName;

  if (existing.hasNext()) {
    // File exists - create unique name by inserting timestamp before extension
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
    const lastDot = fileName.lastIndexOf(".");
    if (lastDot > 0) {
      finalName = fileName.substring(0, lastDot) + "-" + timestamp + fileName.substring(lastDot);
    } else {
      finalName = fileName + "-" + timestamp;
    }
  }

  folder.createFile(finalName, content, MimeType.CSV);
  return finalName;
}
