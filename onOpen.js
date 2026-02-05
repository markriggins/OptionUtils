/**
 *  Add AppScript menu items to a google sheet
 *  OptionTools/
 *    RefreshOptionPrices
 *    Warm XLookup Cache
 **/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('OptionTools')
    .addItem('Refresh Option Prices', 'refreshOptionPrices')
    .addItem('PlotPortfolioValueByPrice', 'PlotPortfolioValueByPrice')
    .addItem('Run SpreadFinder', 'runSpreadFinder')
    .addItem('View SpreadFinder Graphs', 'showSpreadFinderGraphs')
    .addSeparator()
    .addItem('Import Transactions from E*Trade', 'importEtradeTransactions')
    .addSeparator()
    .addItem('Initialize / Clear Project', 'initializeProject')
    .addToUi();
}

/**
 * Creates a fresh Legs sheet with headers and a sample TSLA bull call spread.
 * If a Legs sheet already exists, prompts for confirmation before clearing.
 */
function initializeProject() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const existing = ss.getSheetByName("Legs");
  if (existing) {
    const resp = ui.alert(
      "Initialize / Clear Project",
      "This will delete the existing Legs sheet and all position data.\n\nContinue?",
      ui.ButtonSet.OK_CANCEL
    );
    if (resp !== ui.Button.OK) return;
    ss.deleteSheet(existing);
  }

  // Remove any existing Config named ranges
  for (const nr of ss.getNamedRanges()) {
    if (nr.getName().startsWith("Config_")) nr.remove();
  }

  // Remove any existing PortfolioValueByPrice sheets
  for (const sheet of ss.getSheets()) {
    if (sheet.getName().endsWith("PortfolioValueByPrice")) {
      ss.deleteSheet(sheet);
    }
  }

  // Create Legs sheet
  const sheet = ss.insertSheet("Legs");
  const headers = ["Symbol", "Group", "Strategy", "Strike", "Type", "Expiration", "Qty", "Price", "Investment", "Rec Close", "Closed", "Gain", "LastTxnDate", "Link"];

  // Header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground("#93c47d");
  headerRange.setFontWeight("bold");

  // Sample data: TSLA bull call spread
  const sampleData = [
    ["TSLA", 1, "bull-call-spread", 500, "Call", "12/15/2028", 10, 127.90, "", "", "", "", "", ""],
    ["",      "", "",                 600, "Call", "12/15/2028", -10, 105.90, "", "", "", "", "", ""],
  ];
  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);

  // Named range
  ss.setNamedRange("LegsTable", sheet.getRange("A:N"));

  // Filter
  sheet.getRange("A:N").createFilter();

  // Clip wrap
  sheet.getRange("A:N").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // Auto-resize
  sheet.autoResizeColumns(1, headers.length);

  // Create OptionPricesUploaded sheet if needed
  let pricesSheet = ss.getSheetByName("OptionPricesUploaded");
  if (pricesSheet) {
    pricesSheet.clearContents();
  } else {
    pricesSheet = ss.insertSheet("OptionPricesUploaded");
  }
  const priceHeaders = ["symbol", "expiration", "strike", "type", "bid", "mid", "ask", "iv", "delta", "volume", "openint", "moneyness"];
  const samplePrices = [
    ["TSLA", new Date(2028, 11, 15), 500, "Call", 127.2, 128.23, 129.25, "54.18%", 0.6313, 206, 1345, "-23.89%"],
    ["TSLA", new Date(2028, 11, 15), 600, "Call", 105.4, 106.18, 106.95, "54.45%", 0.5558, 99, 508, "-48.67%"],
  ];
  pricesSheet.getRange(1, 1, 1, priceHeaders.length).setValues([priceHeaders]);
  pricesSheet.getRange(1, 1, 1, priceHeaders.length).setFontWeight("bold");
  pricesSheet.getRange(2, 1, samplePrices.length, priceHeaders.length).setValues(samplePrices);
  pricesSheet.setFrozenRows(1);

  // Create README sheet
  let readmeSheet = ss.getSheetByName("README");
  if (readmeSheet) {
    readmeSheet.clearContents();
  } else {
    readmeSheet = ss.insertSheet("README", 0); // First tab
  }

  const readmeRows = [
    ["OptionUtils - Portfolio Modeling for Option Spreads"],
    ["This spreadsheet models option portfolios using Google Sheets. It tracks stock positions and option spreads (bull call spreads, bull put spreads, iron condors, iron butterflies), calculates value at every price point, and generates charts showing profit/loss scenarios at expiration."],
    ["The OptionTools menu provides all available actions. After editing your positions on the Legs sheet, use PlotPortfolioValueByPrice to generate per-symbol charts. The system reads your positions, computes portfolio value across a range of stock prices, and creates four charts: strategy-level dollar value, strategy-level % ROI, individual spread dollar value, and individual spread ROI."],
    ["--- Getting Started ---"],
    ["A sample TSLA bull call spread is included on the Legs sheet. Run OptionTools > PlotPortfolioValueByPrice and select TSLA to see it in action. Edit the Legs sheet to add your own positions."],
    ["--- The Legs Sheet ---"],
    ["Each position group occupies one or more rows. The first row of a group has the Symbol and Group number; subsequent legs in the same group leave those columns blank. Columns: Symbol, Group, Strategy, Strike, Type (Call/Put/Stock), Expiration, Qty (positive=long, negative=short), Price (entry price per share). The Strategy column is auto-detected. The Closed column tracks closing prices; when all legs in a group have a closing price, the position is excluded from charts."],
    ["--- Importing Option Prices from Barchart.com ---"],
    ["Go to barchart.com and navigate to the options page for your symbol (e.g. barchart.com/stocks/quotes/TSLA/options). Select the expiration date you want, choose \"Stacked\" view to see calls and puts together, then click the download/export button to get a CSV file."],
    ["Save the CSV to your Google Drive under: <DataFolder>/OptionPrices/ (default: OptionUtils/DATA/OptionPrices/). The filename must contain the symbol and expiration date in the format <symbol>-options-exp-YYYY-MM-DD. Example: tsla-options-exp-2028-12-15-monthly-show-all-stacked-01-15-2026.csv"],
    ["Then run OptionTools > Refresh Option Prices. The script scans all symbol folders, picks the most recent CSV per expiration, and loads the data into the OptionPricesUploaded sheet. This data is used by SpreadFinder and the Rec Close column on the Legs sheet."],
    ["--- Importing Transactions from E*Trade ---"],
    ["Log into E*Trade, go to Accounts > Transaction History, and download the transaction CSV. Save it to your Google Drive under: <DataFolder>/Etrade/ (default: OptionUtils/DATA/Etrade/). All transaction CSVs go in one folder."],
    ["Run OptionTools > Import Transactions from E*Trade. The script reads all CSV files in that folder, deduplicates transactions across overlapping date ranges, pairs opening trades into spreads (including iron condors and iron butterflies), and writes them to the Legs sheet. Closing transactions (Sold To Close, Bought To Cover) automatically fill the Closed column with closing prices."],
    ["--- Visualizing Your Portfolio ---"],
    ["Run OptionTools > PlotPortfolioValueByPrice, then select which symbols to chart. For each symbol, the script creates a tab with four charts: (1) Portfolio Value by Price ($) showing aggregated strategy curves and total, (2) % Return by Price showing ROI for shares and each strategy type, (3) Individual Spreads ($) showing each spread separately, and (4) Individual Spreads ROI. A config table in columns K-L lets you adjust the price range, step size, and chart title."],
    ["--- Finding New Spreads with SpreadFinder ---"],
    ["Run OptionTools > Run SpreadFinder to scan the OptionPricesUploaded data for attractive bull call spread opportunities. Configure filters (max spread width, min open interest, min ROI, max debit) on the SpreadFinderConfig sheet. Results are written to the Spreads sheet, ranked by ROI. Use OptionTools > View SpreadFinder Graphs for visual analysis."],
    ["--- Config Sheet ---"],
    ["The Config sheet controls where the scripts look for data files on Google Drive. The DataFolder setting (default: OptionUtils/DATA) is the base path. E*Trade CSVs go in <DataFolder>/Etrade/ and option price CSVs go in <DataFolder>/OptionPrices/. Initialize / Clear Project creates these folders and sample CSV files automatically."],
    ["--- Tips ---"],
    ["Keep your option price CSVs up to date by downloading fresh ones from barchart.com regularly. The Rec Close column on the Legs sheet shows recommended closing prices based on current market data, helping you decide when to take profits or cut losses. Positions with all legs closed are automatically dimmed on the Legs sheet and excluded from portfolio charts."],
  ];

  readmeSheet.getRange(1, 1, readmeRows.length, 1).setValues(readmeRows);

  // Title row bold and larger
  readmeSheet.getRange(1, 1).setFontWeight("bold").setFontSize(14);

  // Section headers bold
  for (let r = 0; r < readmeRows.length; r++) {
    if (readmeRows[r][0].startsWith("---")) {
      readmeSheet.getRange(r + 1, 1).setFontWeight("bold").setFontSize(11)
        .setValue(readmeRows[r][0].replace(/^--- ?| ?---$/g, ""));
    }
  }

  // Column A: 60 characters wide (~420 pixels), wrap text
  readmeSheet.setColumnWidth(1, 420);
  readmeSheet.getRange("A:A").setWrap(true);

  // Hide other columns
  readmeSheet.hideColumns(2, readmeSheet.getMaxColumns() - 1);

  // ---- Create Config sheet ----
  let configSheet = ss.getSheetByName("Config");
  if (configSheet) {
    configSheet.clearContents();
  } else {
    configSheet = ss.insertSheet("Config");
  }
  const configHeaders = ["Setting", "Value"];
  configSheet.getRange(1, 1, 1, 2).setValues([configHeaders]).setFontWeight("bold");
  configSheet.getRange(2, 1, 1, 2).setValues([["DataFolder", "OptionUtils/DATA"]]);
  configSheet.autoResizeColumns(1, 2);

  // ---- Create Google Drive folders and sample files from GitHub ----
  const dataFolderPath = getConfigValue_(ss, "DataFolder", "OptionUtils/DATA");
  try {
    let driveFolder = DriveApp.getRootFolder();
    for (const part of dataFolderPath.split("/")) {
      driveFolder = getOrCreateFolder_(driveFolder, part);
    }
    const etradeFolder = getOrCreateFolder_(driveFolder, "Etrade");
    const optionPricesFolder = getOrCreateFolder_(driveFolder, "OptionPrices");

    const sampleFiles = [
      { folder: etradeFolder, name: "PortfolioDownload-sample.csv", path: "DATA/Etrade/PortfolioDownload-sample.csv" },
      { folder: etradeFolder, name: "DownloadTxnHistory-sample.csv", path: "DATA/Etrade/DownloadTxnHistory-sample.csv" },
      { folder: optionPricesFolder, name: "tsla-options-exp-2028-12-15-monthly-show-all-stacked-sample.csv", path: "DATA/OptionPrices/tsla-options-exp-2028-12-15-monthly-show-all-stacked-sample.csv" },
    ];

    for (const sf of sampleFiles) {
      if (!findFileByName_(sf.folder, sf.name)) {
        const csv = fetchGitHubFile_(sf.path);
        if (csv) sf.folder.createFile(sf.name, csv, MimeType.CSV);
      }
    }
  } catch (e) {
    Logger.log("Drive sample file setup skipped: " + e.message);
  }

  // Activate the Legs sheet
  ss.setActiveSheet(sheet);

  ui.alert("Project initialized with a sample TSLA bull call spread.\n\nSee the README tab for instructions.\n\nEdit the Legs table with your positions, then use:\n  OptionTools > PlotPortfolioValueByPrice\n  OptionTools > Refresh Option Prices");
}

/**
 * Reads a value from the Config sheet by key.
 * Looks for a row where column A matches `key` and returns the column B value.
 *
 * @param {Spreadsheet} ss - The active spreadsheet.
 * @param {string} key - The setting name to look up (e.g. "DataFolder").
 * @param {string} defaultValue - Value to return if key not found or no Config sheet.
 * @returns {string} The config value or defaultValue.
 */
function getConfigValue_(ss, key, defaultValue) {
  const sheet = ss.getSheetByName("Config");
  if (!sheet) return defaultValue;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { // skip header row
    if (String(data[i][0]).trim() === key) {
      const val = String(data[i][1] || "").trim();
      return val || defaultValue;
    }
  }
  return defaultValue;
}

/**
 * Navigates to a subfolder by name, creating it if it doesn't exist.
 *
 * @param {Folder} parent - Parent folder.
 * @param {string} name - Subfolder name.
 * @returns {Folder} The existing or newly created subfolder.
 */
function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

/**
 * Fetches a file from the OptionUtils GitHub repo.
 * @param {string} path - Path relative to repo root (e.g. "DATA/Etrade/PortfolioDownload-sample.csv").
 * @returns {string|null} File content, or null on failure.
 */
function fetchGitHubFile_(path) {
  const url = "https://raw.githubusercontent.com/markriggins/OptionUtils/main/" + path;
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() === 200) return resp.getContentText();
    Logger.log("GitHub fetch " + resp.getResponseCode() + " for " + path);
  } catch (e) {
    Logger.log("GitHub fetch failed for " + path + ": " + e.message);
  }
  return null;
}

/**
 * Checks if a file with the given name exists in a folder.
 * @param {Folder} folder - The folder to search.
 * @param {string} name - The file name.
 * @returns {File|null} The file if found, null otherwise.
 */
function findFileByName_(folder, name) {
  const it = folder.getFilesByName(name);
  return it.hasNext() ? it.next() : null;
}

function warmXLookupCache() {
  XLookupByKeys_WarmCache("OptionPricesUploaded", ["Symbol", "Expiration", "Strike", "Type"], ["Bid", "Ask"]);
  SpreadsheetApp.getActiveSpreadsheet().toast("Cache warmed", "Done", 3);
}
