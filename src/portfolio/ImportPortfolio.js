/**
 * ImportPortfolio.js
 * Orchestrates portfolio import via file upload (no Drive storage).
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
      SpreadsheetApp.flush();
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
        // TODO: support other brokerages
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
      portfolioResult = parseEtradePortfolioStocksAndCash_(portfolio.content, stockTxns);
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
    portfolioOptionData = parseEtradePortfolioOptionsWithPrices_(portfolio.content);
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

  // Pair into spreads, then combine naked legs across dates
  const pairedSpreads = pairTransactionsIntoSpreads_(allTransactions);

  // In addTransactions mode, include existing naked short puts for combining
  // so new long puts can be matched with them
  let existingNakedShortPuts = [];
  const combinedExistingKeys = new Set();
  const rowsToDelete = []; // Track spreadsheet rows to delete for combined positions

  if (importMode === "addTransactions" && existingPositions.size > 0) {
    for (const [key, pos] of existingPositions) {
      // Look for single-leg naked short puts: key like "TSLA|12/20/2024|400|Put"
      // Single strike means it's a naked position (not a spread)
      if (pos.legs && pos.legs.length === 1) {
        const leg = pos.legs[0];
        if (leg.type === "Put" && leg.qty < 0 && leg.strike) {
          existingNakedShortPuts.push({
            ticker: leg.symbol,
            expiration: leg.expiration,
            lowerStrike: null,
            upperStrike: leg.strike,
            optionType: "Put",
            qty: leg.qty,
            lowerPrice: 0,
            upperPrice: leg.price || 0,
            date: pos.lastTxnDate,
            _existingKey: key, // Track which existing position this came from
            _existingRow: leg.row, // Track the spreadsheet row number
          });
        }
      }
    }
    log.info("import", `Found ${existingNakedShortPuts.length} existing naked short puts to check for combining`);
  }

  // Build a set of existing naked short put keys to filter duplicates from transactions
  const existingShortPutKeys = new Set();
  for (const esp of existingNakedShortPuts) {
    const exp = formatExpirationForKey_(esp.expiration);
    const key = `${esp.ticker}|${exp}|${esp.upperStrike}`;
    existingShortPutKeys.add(key);
  }

  // Filter out naked short puts from transactions that duplicate existing portfolio positions
  // This ensures the EXISTING shorts (with _existingKey) get matched with new longs
  const filteredPairedSpreads = pairedSpreads.filter(sp => {
    if (sp.optionType === "Put" && sp.lowerStrike === null && sp.upperStrike !== null) {
      const exp = formatExpirationForKey_(sp.expiration);
      const key = `${sp.ticker}|${exp}|${sp.upperStrike}`;
      if (existingShortPutKeys.has(key)) {
        log.info("import", `Filtering duplicate naked short from transactions: ${key}`);
        return false; // Skip - use existing portfolio position instead
      }
    }
    return true;
  });


  // Combine naked legs - include filtered transactions AND existing naked puts
  // Combine matching bull-put-spreads and bear-call-spreads into iron condors/butterflies
  const combinedSpreads = combineRelatedSpreadsIntoIronCondorsAndButterflies_(
    combineNakedLegsIntoSpreads_([...filteredPairedSpreads, ...existingNakedShortPuts]));

  // Identify which existing naked puts were combined vs still naked
  // - Unmatched existing puts remain in result with _existingKey
  // - Matched existing puts are NOT in result (replaced by combined spread without _existingKey)
  const unmatchedExistingKeys = new Set();
  for (const sp of combinedSpreads) {
    if (sp._existingKey) {
      unmatchedExistingKeys.add(sp._existingKey);
    }
  }

  // Existing positions that were matched (combined) need to be removed from existingPositions
  for (const existing of existingNakedShortPuts) {
    if (!unmatchedExistingKeys.has(existing._existingKey)) {
      // This existing naked put was combined with a new long put
      combinedExistingKeys.add(existing._existingKey);
      existingPositions.delete(existing._existingKey);
      // Track the row number to delete from the spreadsheet
      if (existing._existingRow != null) {
        rowsToDelete.push(existing._existingRow);
      }
      log.info("import", `Removed existing naked put ${existing._existingKey} row ${existing._existingRow} (now combined into spread)`);
    }
  }

  // Filter out unmatched existing naked puts (they already exist in portfolio)
  // and clean up tracking field
  const newSpreadsOnly = combinedSpreads.filter(sp => {
    if (sp._existingKey) {
      delete sp._existingKey; // Clean up tracking field
      return false; // Filter out - already in portfolio
    }
    return true;
  });

  const rawSpreads = [...stockPositions, ...newSpreadsOnly];

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

  // Write to Portfolio table (pass rows to delete for combined positions)
  writePortfolioTable_(ss, headers, updatedLegs, newLegs, closingPrices, rowsToDelete);

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
