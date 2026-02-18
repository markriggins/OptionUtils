/**
 * ImportEtrade.js
 * Orchestrates E*Trade portfolio import.
 *
 * Entry points and coordination logic. Parsing, spread building, and
 * sheet output are delegated to:
 *   - EtradeCsvParser.js - CSV parsing
 *   - PositionBuilder.js - spread pairing/aggregation
 *   - PortfolioWriter.js - sheet output
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

  // Process all transaction CSVs, dedup across file boundaries
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
    if (portfolioFiles.length > 0) {
      const portfolioResult = parsePortfolioStocksAndCashFromFile_(portfolioFiles[0], stockTxns);
      stockPositions = portfolioResult.stocks;
      portfolioCash = portfolioResult.cash || 0;
      log.info("import", `Found ${stockPositions.length} stock positions and $${portfolioCash} cash from portfolio CSV`);
    }
  } else {
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

  // Pre-merge spreads with the same key
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

    // Add orphaned options as single-leg positions
    if (validation.missing.length > 0) {
      validationWarnings += "\n\n⚠️ ORPHANED OPTIONS (added as single legs):\n";
      for (const m of validation.missing) {
        const [ticker, exp, strike, type] = m.key.split("|");
        const priceInfo = portfolioOptionData.prices.get(m.key) || { pricePaid: 0 };

        const isLong = m.portfolio > 0;
        const singleLeg = {
          type: "single-option",
          ticker: ticker,
          qty: isLong ? m.portfolio : Math.abs(m.portfolio),
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

    if (validation.extra.length > 0) {
      log.debug("import", "Closed positions (in transactions but not in portfolio): " +
        validation.extra.map(m => m.key).join(", "));
    }

    if (validationWarnings) {
      log.warn("import", "Validation warnings: " + validationWarnings);
    }
  }

  // Merge spreads into existing positions
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
 * Imports E*Trade portfolio from a specific folder and filenames.
 * Used by loadSamplePortfolio to import sample data without UI prompts.
 *
 * @param {Folder} folder - Google Drive folder containing the CSV files
 * @param {string} txnFileName - Transaction history CSV filename
 * @param {string} portfolioFileName - Portfolio CSV filename
 */
function importEtradePortfolioFromFolder_(folder, txnFileName, portfolioFileName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const txnFiles = findFilesByName_(folder, txnFileName);
  const portfolioFiles = findFilesByName_(folder, portfolioFileName);

  if (txnFiles.length === 0) {
    throw new Error(`Transaction file not found: ${txnFileName}`);
  }

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

  let stockPositions = [];
  let portfolioCash = 0;
  if (portfolioFiles.length > 0) {
    const portfolioResult = parsePortfolioStocksAndCashFromFile_(portfolioFiles[0], stockTxns);
    stockPositions = portfolioResult.stocks;
    portfolioCash = portfolioResult.cash || 0;
  }

  const spreads = [...stockPositions, ...pairTransactionsIntoSpreads_(transactions)];

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

  const closingPrices = buildClosingPricesMap_(transactions, stockTxns);

  writePortfolioTable_(ss, [], [], spreads, closingPrices);

  log.info("import", `Sample portfolio imported: ${spreads.length} positions`);
}

/* =========================================================
   File Finding Utilities
   ========================================================= */

/**
 * Finds all files in a folder whose name starts with prefix, sorted newest first by name.
 */
function findFilesByPrefix_(folder, prefix) {
  log.debug("files", `Looking for '${prefix}' in folder '${folder.getName()}'`);

  let iter = folder.searchFiles(`title contains '${prefix}' and mimeType = 'text/csv'`);
  let files = [];
  while (iter.hasNext()) files.push(iter.next());
  log.debug("files", `CSV MIME type search found: ${files.length} files`);

  if (files.length === 0) {
    iter = folder.searchFiles(`title contains '${prefix}'`);
    while (iter.hasNext()) {
      const f = iter.next();
      log.debug("files", `Found file: ${f.getName()} (MIME: ${f.getMimeType()})`);
      const name = f.getName().toLowerCase();
      if (name.endsWith('.csv') || name.endsWith('.cindy') || name.endsWith('.txt')) {
        files.push(f);
      }
    }
    log.debug("files", `After extension filter: ${files.length} files`);
  }

  if (files.length === 0) {
    log.warn("files", `No files found. Listing all files in folder:`);
    const allFiles = folder.getFiles();
    while (allFiles.hasNext()) {
      const f = allFiles.next();
      log.debug("files", `  - ${f.getName()} (MIME: ${f.getMimeType()})`);
    }
  }

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
 */
function stripBrowserDisambiguator_(fileName) {
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

  if (portfolio) {
    saveFileToFolder_(folder, "PortfolioDownload.csv", portfolio.content);
    uploadedParts.push("portfolio");
  }

  if (transactions && transactions.length > 0) {
    for (const txn of transactions) {
      const cleanName = stripBrowserDisambiguator_(txn.name);
      saveFileWithUniqueName_(folder, cleanName, txn.content);
    }
    uploadedParts.push(`${transactions.length} transaction file(s)`);
  }

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
 */
function saveFileWithUniqueName_(folder, fileName, content) {
  const existing = folder.getFilesByName(fileName);
  let finalName = fileName;

  if (existing.hasNext()) {
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
