/**
 * EtradeCsvParser.js
 * E*Trade-specific CSV parsing functions.
 *
 * Parses E*Trade transaction history and portfolio download CSVs
 * into normalized transaction objects that can be processed by PositionBuilder.
 */

/**
 * Parses E*Trade transaction CSV content.
 * Supports both old format (TransactionDate header) and new format (Activity/Trade Date header).
 *
 * @param {string} csvContent - Raw CSV content
 * @returns {{ transactions: Object[], stockTxns: Object[] }}
 */
function parseEtradeTransactionsFromCsv_(csvContent) {
  const lines = csvContent.split(/\r?\n/);
  const transactions = [];
  const stockTxns = [];

  // Find header row - support both old and new formats
  let headerIdx = -1;
  let isNewCsvFormat = false;
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].startsWith("TransactionDate,")) {
      headerIdx = i;
      isNewCsvFormat = false;
      break;
    }
    if (lines[i].startsWith("Activity/Trade Date,")) {
      headerIdx = i;
      isNewCsvFormat = true;
      break;
    }
  }
  if (headerIdx < 0) {
    throw new Error(
      "E*Trade Transaction CSV: Could not find header row.\n" +
      "Expected row starting with 'TransactionDate,' or 'Activity/Trade Date,'\n" +
      "Tip: Make sure you're uploading the Transaction History CSV from E*Trade."
    );
  }

  const optionTxnTypes = [
    "Bought To Open", "Sold Short",
    "Sold To Close", "Bought To Cover",
    "Option Assigned", "Option Exercised",
  ];

  // Old-format types: "Bought"/"Sold" with OPENING/CLOSING in Description
  const oldFormatSimpleTypes = ["Bought", "Sold"];

  // Find column indices from header
  const headerCols = parseCsvLine_(lines[headerIdx]);
  const colIndex = (name) => headerCols.findIndex(h => h.trim().toLowerCase() === name.toLowerCase());

  // Column mappings differ between formats
  let dateIdx, txnTypeIdx, secTypeIdx, symbolIdx, qtyIdx, amountIdx, priceIdx, descIdx;
  if (isNewCsvFormat) {
    // New format: Activity/Trade Date, Transaction Date, Settlement Date, Activity Type, Description, Symbol, Cusip, Quantity #, Price $, Amount $, Commission, Category, Note
    dateIdx = colIndex("Activity/Trade Date");
    txnTypeIdx = colIndex("Activity Type");
    secTypeIdx = -1;  // No SecurityType in new format
    symbolIdx = colIndex("Symbol");
    qtyIdx = colIndex("Quantity #");
    priceIdx = colIndex("Price $");
    amountIdx = colIndex("Amount $");
    descIdx = colIndex("Description");
  } else {
    // Old format: TransactionDate, TransactionType, SecurityType, Symbol, Quantity, Amount, Price, Commission, Description
    dateIdx = 0;
    txnTypeIdx = 1;
    secTypeIdx = 2;
    symbolIdx = 3;
    qtyIdx = 4;
    amountIdx = 5;
    priceIdx = 6;
    descIdx = colIndex("Description");
  }

  // Parse rows
  for (let i = headerIdx + 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    const cols = parseCsvLine_(line);
    if (cols.length < 6) continue;

    const dateStr = cols[dateIdx] || "";
    const txnType = cols[txnTypeIdx] || "";
    const secType = secTypeIdx >= 0 ? cols[secTypeIdx] || "" : "";
    const symbol = cols[symbolIdx] || "";
    const qtyStr = cols[qtyIdx] || "";
    const amountStr = cols[amountIdx] || "";
    const priceStr = cols[priceIdx] || "";
    const desc = descIdx >= 0 && cols.length > descIdx ? cols[descIdx] : "";

    // Stock transactions (for exercise/assignment matching and portfolio aggregation)
    const isStockTxn = isNewCsvFormat
      ? (txnType === "Bought" || txnType === "Sold") && !desc.match(/^(CALL|PUT)\s/)
      : secType === "EQ" && (txnType === "Bought" || txnType === "Sold");

    if (isStockTxn) {
      stockTxns.push({
        date: dateStr,
        txnType: txnType,
        ticker: symbol.trim().toUpperCase().replace(/^--$/, ""),
        qty: parseFloat(qtyStr) || 0,
        price: parseFloat(priceStr) || 0,
      });
      continue;
    }

    // Option transactions - check if this is an option trade
    const isOptionTxn = isNewCsvFormat
      ? optionTxnTypes.includes(txnType) || ((txnType === "Bought" || txnType === "Sold") && desc.match(/^(CALL|PUT)\s/))
      : secType === "OPTN";

    if (!isOptionTxn) continue;

    // Determine transaction type category
    const isStandardOptionType = optionTxnTypes.includes(txnType);
    const isSimpleType = oldFormatSimpleTypes.includes(txnType);
    if (!isStandardOptionType && !isSimpleType) continue;

    // Parse option details - in new format, option info is in Description; in old format, it's in Symbol
    let parsed;
    if (isNewCsvFormat) {
      // New format: parse from Description like "PUT  TSLA   02/27/26   400.000"
      parsed = parseEtradeOptionDescription_(desc);
    } else {
      // Old format: parse from Symbol (OCC format or human-readable)
      parsed = parseEtradeOptionSymbol_(symbol) || parseOccOptionSymbol_(symbol);
    }
    if (!parsed) continue;

    // Determine open/close/exercise/assigned
    let isOpen, isClosed, isExercised, isAssigned;
    if (isStandardOptionType) {
      isOpen = txnType === "Bought To Open" || txnType === "Sold Short";
      isClosed = txnType === "Sold To Close" || txnType === "Bought To Cover";
      isExercised = txnType === "Option Exercised";
      isAssigned = txnType === "Option Assigned";
    } else {
      // Simple "Bought"/"Sold" types: check Description for OPENING/CLOSING
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
 * Parses option details from new E*Trade CSV Description format.
 * Example: "PUT  TSLA   02/27/26   400.000" or "CALL TSLA   12/15/28   600.000"
 * @returns {{ ticker, expiration, strike, type }} or null
 */
function parseEtradeOptionDescription_(desc) {
  if (!desc) return null;

  // Match: TYPE TICKER DATE STRIKE
  // Examples: "PUT  TSLA   02/27/26   400.000"
  //           "CALL TSLA   12/15/28   600.000"
  const match = desc.match(/^(CALL|PUT)\s+([A-Z]+)\s+(\d{2}\/\d{2}\/\d{2})\s+([\d.]+)/i);
  if (!match) return null;

  const [, type, ticker, dateStr, strikeStr] = match;

  // Parse date MM/DD/YY to YYYY-MM-DD
  const [mm, dd, yy] = dateStr.split("/");
  const year = parseInt(yy) < 50 ? 2000 + parseInt(yy) : 1900 + parseInt(yy);
  const expiration = `${year}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}`;

  return {
    ticker: ticker.toUpperCase(),
    expiration,
    strike: parseFloat(strikeStr),
    type: type.toUpperCase() === "CALL" ? "Call" : "Put",
  };
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
 * Parses stock positions and cash from E*Trade PortfolioDownload CSV.
 *
 * @param {File} file - Google Drive file object
 * @param {Object[]} stockTxns - Stock transactions for date lookup
 * @returns {{ stocks: Object[], cash: number }}
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
  if (headerIdx < 0) {
    throw new Error(
      "E*Trade Portfolio CSV: Could not find header row.\n" +
      "Expected row starting with 'Symbol,Last Price'\n" +
      "Tip: Make sure you're uploading the Portfolio Download CSV from E*Trade."
    );
  }

  const headers = parseCsvLine_(lines[headerIdx]);

  // Validate required columns
  const required = validateRequiredColumns_(headers, [
    { name: "Symbol", aliases: ["symbol"] },
    { name: "Quantity", aliases: ["quantity", "qty #", "qty"] },
    { name: "PricePaid", aliases: ["price paid", "price paid $", "cost basis", "avg price"] },
  ], "E*Trade Portfolio CSV");

  const idxSym = required.Symbol;
  const idxQty = required.Quantity;
  const idxPricePaid = required.PricePaid;

  // Optional columns
  const optional = findOptionalColumns_(headers, [
    { name: "MarketValue", aliases: ["market value", "market value $", "value", "value $"] },
  ]);
  const idxMarketValue = optional.MarketValue;

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
  if (headerIdx < 0) {
    throw new Error(
      "E*Trade Portfolio CSV: Could not find header row.\n" +
      "Expected row starting with 'Symbol,Last Price'\n" +
      "Tip: Make sure you're uploading the Portfolio Download CSV from E*Trade."
    );
  }

  const headers = parseCsvLine_(lines[headerIdx]);

  // Validate required columns
  const required = validateRequiredColumns_(headers, [
    { name: "Symbol", aliases: ["symbol"] },
    { name: "Quantity", aliases: ["quantity", "qty #", "qty"] },
  ], "E*Trade Portfolio Options CSV");

  const idxSym = required.Symbol;
  const idxQty = required.Quantity;

  // Optional columns
  const optional = findOptionalColumns_(headers, [
    { name: "PricePaid", aliases: ["price paid", "price paid $", "cost basis", "avg price"] },
  ]);
  const idxPricePaid = optional.PricePaid;

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
