/**
 * ImportEtrade.js
 * Orchestrates E*Trade portfolio import via file upload (no Drive storage).
 *
 * Parsing, spread building, and sheet output are delegated to:
 *   - EtradeCsvParser.js - CSV parsing
 *   - PositionBuilder.js - spread pairing/aggregation
 *   - PortfolioWriter.js - sheet output
 */

/* =========================================================
   File Upload Dialog
   ========================================================= */

/**
 * Checks if the Portfolio sheet has any data.
 * @returns {boolean} True if there is an existing portfolio.
 */
function hasExistingPortfolio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Portfolio");
  if (!sheet) return false;
  return sheet.getLastRow() > 1; // More than just header row
}

/**
 * Shows file upload dialog for portfolio import.
 */
function showUploadRebuildDialog() {
  const html = HtmlService.createHtmlOutputFromFile("ui/FileUpload")
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
 */
function stripBrowserDisambiguator_(fileName) {
  return fileName.replace(/\s*\(\d+\)(?=\.[^.]+$|$)/, "");
}

/* =========================================================
   Portfolio Upload Handler
   ========================================================= */

/**
 * Uploads portfolio/transaction files and processes them directly (no Drive storage).
 * @param {{name: string, content: string}} portfolio - Portfolio CSV file (optional)
 * @param {Array<{name: string, content: string}>} transactions - Transaction CSV files
 * @param {string} importMode - "addTransactions" or "rebuild"
 * @returns {string} Status message including any missing option prices warning
 */
function uploadAndRebuildPortfolio(portfolio, transactions, importMode) {
  importMode = importMode || "rebuild";

  if (importMode === "addTransactions") {
    if (!transactions || transactions.length === 0) {
      throw new Error("Transaction file(s) required for 'Add transactions' mode");
    }
  } else {
    if (!portfolio && (!transactions || transactions.length === 0)) {
      throw new Error("At least one file (Portfolio or Transactions) is required");
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Handle rebuild mode - delete existing sheet
  if (importMode === "rebuild") {
    const existingSheet = ss.getSheetByName("Portfolio");
    if (existingSheet) {
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
    log.info("import", `Existing positions: ${existingPositions.size} groups`);
    for (const [key, pos] of existingPositions) {
      if (key.includes("STOCK")) {
        log.info("import", `  Existing stock: ${key} qty=${pos.legs?.[0]?.qty}`);
      }
    }
  } else {
    log.info("import", `No existing Portfolio table found`);
  }

  // Parse transaction CSVs directly
  let allTransactions = [];
  let stockTxns = [];
  const seenTxns = new Set();
  const seenStockTxns = new Set();
  const uploadedParts = [];

  if (transactions && transactions.length > 0) {
    for (const txnFile of transactions) {
      let result;
      try {
        result = parseEtradeTransactionsFromCsv_(txnFile.content);
      } catch (e) {
        // Check if this looks like a portfolio file uploaded in wrong slot
        if (txnFile.content.includes("Symbol,Last Price")) {
          throw new Error(
            `"${txnFile.name}" appears to be a Portfolio CSV, not a Transaction CSV.\n` +
            "Use the Portfolio file chooser for portfolio downloads."
          );
        }
        throw e;
      }
      let txnAdded = 0;
      for (const txn of result.transactions) {
        const key = `${txn.date}|${txn.txnType}|${txn.ticker}|${txn.expiration}|${txn.strike}|${txn.optionType}|${txn.qty}|${txn.price}|${txn.amount}`;
        if (seenTxns.has(key)) continue;
        seenTxns.add(key);
        allTransactions.push(txn);
        txnAdded++;
      }
      for (const stk of result.stockTxns) {
        const key = `${stk.date}|${stk.ticker}|${stk.qty}|${stk.price}|${stk.amount}`;
        if (seenStockTxns.has(key)) continue;
        seenStockTxns.add(key);
        stockTxns.push(stk);
      }
      log.debug("import", `${txnFile.name}: ${txnAdded} transactions`);
    }
    uploadedParts.push(`${transactions.length} transaction file(s)`);
  }

  if (allTransactions.length === 0 && importMode === "addTransactions") {
    throw new Error("No option transactions found in the uploaded file(s)");
  }

  // Parse portfolio CSV directly
  let stockPositions = [];
  let portfolioCash = 0;
  let portfolioOptionData = null;

  if (portfolio && importMode !== "addTransactions") {
    let portfolioResult;
    try {
      portfolioResult = parsePortfolioStocksAndCash_(portfolio.content, stockTxns);
    } catch (e) {
      // Check if this looks like a transaction file uploaded in wrong slot
      if (portfolio.content.includes("TransactionDate,") || portfolio.content.includes("Activity/Trade Date,")) {
        throw new Error(
          `"${portfolio.name}" appears to be a Transaction CSV, not a Portfolio CSV.\n` +
          "Use the Transaction file chooser for transaction history."
        );
      }
      throw e;
    }
    stockPositions = portfolioResult.stocks;
    portfolioCash = portfolioResult.cash || 0;
    portfolioOptionData = parsePortfolioOptionsWithPrices_(portfolio.content);
    uploadedParts.push("portfolio");
    log.info("import", `Found ${stockPositions.length} stock positions and $${portfolioCash} cash from portfolio CSV`);
  } else if (importMode === "addTransactions") {
    // For update mode, aggregate stock transactions
    const stockCutoffDates = new Map();
    for (const [key, pos] of existingPositions) {
      if (key.endsWith("|STOCK")) {
        const ticker = key.split("|")[0];
        const cutoff = parseDateAtMidnight_(pos.lastTxnDate);
        if (cutoff) stockCutoffDates.set(ticker, cutoff);
      }
    }
    stockPositions = aggregateStockTransactions_(stockTxns, stockCutoffDates);
  }

  // Pair into spreads
  const rawSpreads = [...stockPositions, ...pairTransactionsIntoSpreads_(allTransactions)];

  // Debug: log stock positions
  for (const sp of rawSpreads) {
    if (sp.type === "stock") {
      log.info("import", `Stock position: ${sp.ticker} qty=${sp.qty} from ${sp.date}`);
    }
  }

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
  }

  // Pre-merge spreads with the same key
  const spreads = preMergeSpreads_(rawSpreads);
  log.info("import", `Paired into ${spreads.length} spread orders`);

  // Debug: log stock positions after merge
  for (const sp of spreads) {
    if (sp.type === "stock") {
      log.info("import", `After merge: ${sp.ticker} qty=${sp.qty}`);
    }
  }

  // Build closing prices map
  const closingPrices = buildClosingPricesMap_(allTransactions, stockTxns);

  // Validate quantities if we have portfolio data
  if (portfolioOptionData && importMode !== "addTransactions") {
    const validation = validateOptionQuantities_(spreads, portfolioOptionData.quantities, allTransactions);

    // Add orphaned options as single-leg positions
    if (validation.extra.length > 0) {
      for (const m of validation.extra) {
        const [ticker, exp, strike, type] = m.key.split("|");
        const priceInfo = portfolioOptionData.prices.get(m.key) || { pricePaid: 0 };
        const isLong = m.qty > 0;
        spreads.push({
          type: "single-option",
          ticker: ticker,
          qty: isLong ? m.qty : Math.abs(m.qty),
          expiration: exp,
          optionType: type,
          lowerStrike: isLong ? parseFloat(strike) : null,
          upperStrike: isLong ? null : parseFloat(strike),
          lowerPrice: isLong ? priceInfo.pricePaid : 0,
          upperPrice: isLong ? 0 : priceInfo.pricePaid,
          date: new Date(),
        });
      }
    }
  }

  // Merge spreads into existing positions
  const { updatedLegs, newLegs, skippedCount } = mergeSpreads_(existingPositions, spreads);

  // Write to Portfolio table
  writePortfolioTable_(ss, headers, updatedLegs, newLegs, closingPrices);

  // Build summary
  let summary = "";
  if (importMode === "addTransactions") {
    summary = `Added ${newLegs.length} new positions, updated ${updatedLegs.length} existing.`;
    if (skippedCount > 0) summary += ` Skipped ${skippedCount} (already imported).`;
  } else {
    summary = `Rebuilt portfolio with ${newLegs.length} positions.`;
    if (portfolioCash > 0) summary += ` Cash: $${portfolioCash.toLocaleString()}.`;
  }

  // Check for missing option prices
  const missingPrices = findMissingOptionPrices_(ss);
  if (missingPrices.length > 0) {
    summary += "\n\n⚠️ Missing option prices for:\n• " + missingPrices.join("\n• ");
    summary += "\n\nUpload option prices for these expirations to see current values.";
  }

  return summary;
}

/* =========================================================
   Missing Prices Check
   ========================================================= */

/**
 * Finds open portfolio positions that don't have option prices uploaded.
 * Skips closed positions (qty = 0).
 * @returns {string[]} List of "SYMBOL EXP_DATE" strings for positions missing prices
 */
function findMissingOptionPrices_(ss) {
  const missing = [];

  // Get portfolio positions
  const portfolioRange = getNamedRangeWithTableFallback_(ss, "Portfolio");
  if (!portfolioRange) return missing;

  const rows = portfolioRange.getValues();
  if (rows.length < 2) return missing;

  const headers = rows[0];
  const symIdx = findColumn_(headers, ["symbol", "ticker"]);
  const expIdx = findColumn_(headers, ["expiration", "exp", "expiry"]);
  const qtyIdx = findColumn_(headers, ["qty", "quantity", "contracts"]);

  if (symIdx < 0 || expIdx < 0) return missing;

  // Collect unique symbol+expiration combos from open positions only
  const portfolioCombos = new Set();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 1; i < rows.length; i++) {
    const sym = String(rows[i][symIdx] || "").trim().toUpperCase();
    const expRaw = rows[i][expIdx];
    if (!sym || !expRaw) continue;

    // Skip closed positions
    if (qtyIdx >= 0) {
      const qty = parseFloat(rows[i][qtyIdx]) || 0;
      if (qty === 0) continue;
    }

    const expDate = parseDateAtMidnight_(expRaw);
    if (!expDate) continue;

    // Skip expired positions
    if (expDate < today) continue;

    const expStr = expDate.getFullYear() + "-" +
      String(expDate.getMonth() + 1).padStart(2, "0") + "-" +
      String(expDate.getDate()).padStart(2, "0");

    portfolioCombos.add(`${sym}|${expStr}`);
  }

  if (portfolioCombos.size === 0) return missing;

  // Get uploaded option prices
  const optionSheet = ss.getSheetByName("OptionPricesUploaded");
  const uploadedCombos = new Set();

  if (optionSheet && optionSheet.getLastRow() > 1) {
    const optionData = optionSheet.getRange(2, 1, optionSheet.getLastRow() - 1, 2).getValues();
    for (const row of optionData) {
      const sym = String(row[0] || "").trim().toUpperCase();
      const expRaw = row[1];
      if (!sym || !expRaw) continue;

      const expDate = parseDateAtMidnight_(expRaw);
      if (!expDate) continue;

      const expStr = expDate.getFullYear() + "-" +
        String(expDate.getMonth() + 1).padStart(2, "0") + "-" +
        String(expDate.getDate()).padStart(2, "0");

      uploadedCombos.add(`${sym}|${expStr}`);
    }
  }

  // Find portfolio combos not in uploaded prices
  for (const combo of portfolioCombos) {
    if (!uploadedCombos.has(combo)) {
      const [sym, exp] = combo.split("|");
      missing.push(`${sym} ${exp}`);
    }
  }

  // Sort by symbol, then date
  missing.sort();

  return missing;
}
