/**
 *  Add AppScript menu items to a google sheet
 *  OptionTools/
 *    RefreshOptionPrices
 *    Warm XLookup Cache
 **/

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const spreadFinderMenu = ui.createMenu('SpreadFinder')
    .addItem('Refresh Option Prices', 'refreshOptionPrices')
    .addSeparator()
    .addItem('Run SpreadFinder', 'runSpreadFinder')
    .addItem('View Graphs', 'showSpreadFinderGraphs');

  const portfolioMenu = ui.createMenu('Portfolio')
    .addItem('Import Latest Transactions', 'importLatestTransactions')
    .addItem('Clear & Rebuild from E*Trade', 'rebuildPortfolio')
    .addSeparator()
    .addItem('Load Sample Portfolio', 'loadSamplePortfolio')
    .addSeparator()
    .addItem('View Performance Graphs', 'PlotPortfolioValueByPrice');

  ui.createMenu('OptionTools')
    .addItem('Initialize / Clear Project', 'initializeProject')
    .addSeparator()
    .addSubMenu(spreadFinderMenu)
    .addSubMenu(portfolioMenu)
    .addToUi();
}

/**
 * Initializes the project with README, Config, and option price data.
 * Deletes all sheets except README and Config.
 */
function initializeProject() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  // Find sheets that will be deleted (everything except README and Config)
  const keepSheets = ["README", "Config"];
  const sheetsToDelete = ss.getSheets().filter(s => !keepSheets.includes(s.getName()));

  if (sheetsToDelete.length > 0) {
    const sheetNames = sheetsToDelete.map(s => "  • " + s.getName()).join("\n");
    const resp = ui.alert(
      "Initialize / Clear Project",
      "This will delete the following sheets:\n" + sheetNames + "\n\nContinue?",
      ui.ButtonSet.OK_CANCEL
    );
    if (resp !== ui.Button.OK) return;
  }

  // Delete all sheets except README and Config
  for (const sheet of sheetsToDelete) {
    ss.deleteSheet(sheet);
  }

  // Remove any existing named ranges (except Config-related)
  for (const nr of ss.getNamedRanges()) {
    nr.remove();
  }

  // Create OptionPricesUploaded sheet (will be populated by refreshOptionPrices)
  const pricesSheet = ss.insertSheet("OptionPricesUploaded");

  // Create README sheet
  let readmeSheet = ss.getSheetByName("README");
  if (readmeSheet) {
    readmeSheet.clearContents();
  } else {
    readmeSheet = ss.insertSheet("README", 0); // First tab
  }

  const readmeRows = [
    ["SpreadFinder - Find and Analyze Bull Call Spread Opportunities"],
    ["SpreadFinder scans option prices to find attractive bull call spread opportunities. It ranks spreads by expected ROI using a probability-of-touch model, filters by liquidity, and displays results in interactive charts. This spreadsheet includes real TSLA option prices for June 2028 and December 2028 LEAP expirations."],
    [""],
    ["--- Quick Start ---"],
    ["  1. Run OptionTools > Initialize / Clear Project to set up sheets and load TSLA option data"],
    ["  2. Run OptionTools > Run SpreadFinder to analyze spreads (results on TSLASpreads sheet)"],
    ["  3. Run OptionTools > View SpreadFinder Graphs for visual analysis of Delta vs ROI and Strike vs ROI"],
    ["--- SpreadFinder Configuration ---"],
    ["The SpreadFinderConfig sheet controls the analysis parameters: symbol filter, min/max spread width, min open interest, min volume, max debit, min ROI, strike range, and expiration range. Adjust these to narrow down the results to spreads that match your criteria. After changing settings, re-run OptionTools > Run SpreadFinder to see updated results."],
    ["--- Adding More Option Prices ---"],
    ["Download option prices from barchart.com: navigate to the options page for your symbol (e.g. barchart.com/stocks/quotes/TSLA/options), select an expiration, choose \"Stacked\" view, and download the CSV. Save it to Google Drive under SpreadFinder/DATA/OptionPrices/ with the filename format: <symbol>-options-exp-YYYY-MM-DD-....csv. Then run OptionTools > Refresh Option Prices to load the data, and re-run OptionTools > Run SpreadFinder to analyze the new prices."],
    [""],
    ["--- Portfolio Modeling (Additional Feature) ---"],
    ["Track your actual positions and visualize profit/loss scenarios. Bull-Call-Spreads, Long Calls and other strategies will be automatically detected and analyzed."],
    [""],
    ["To try it with sample data:"],
    ["  • Run OptionTools > Load Sample Portfolio"],
    ["  • Run OptionTools > View Portfolio Performance Graphs"],
    [""],
    ["To import your real E*Trade positions:"],
    ["  1. Download your Portfolio CSV and Transaction History CSV from E*Trade"],
    ["  2. Save both files to Google Drive under SpreadFinder/DATA/Etrade/"],
    ["  3. Run OptionTools > Import Portfolio from E*Trade"],
    ["  4. Run OptionTools > View Portfolio Performance Graphs"],
    [""],
    ["--- Apps Script Library ---"],
    ["Available as a Google Apps Script library named SpreadFinder. Script ID: 1qvAlZ99zluKSr3ws4NsxH8xXo1FncbWzu6Yq6raBumHdCLpLDaKveM0T. Source: github.com/markriggins/OptionUtils"],
  ];

  readmeSheet.getRange(1, 1, readmeRows.length, 1).setValues(readmeRows);

  // Set all content to normal weight, font size 18, wrapped
  const contentRange = readmeSheet.getRange(1, 1, readmeRows.length, 1);
  contentRange.setFontWeight("normal").setFontSize(18).setWrap(true);

  // Title row bold
  readmeSheet.getRange(1, 1).setFontWeight("bold");

  // Section headers bold (rows starting with "---")
  for (let r = 0; r < readmeRows.length; r++) {
    if (readmeRows[r][0].startsWith("---")) {
      readmeSheet.getRange(r + 1, 1).setFontWeight("bold")
        .setValue(readmeRows[r][0].replace(/^--- ?| ?---$/g, ""));
    }
  }

  // Column A width
  readmeSheet.setColumnWidth(1, 840);

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
  configSheet.getRange(2, 1, 1, 2).setValues([["DataFolder", "SpreadFinder/DATA"]]);
  configSheet.autoResizeColumns(1, 2);

  // ---- Create Google Drive folders and sample files from GitHub ----
  const dataFolderPath = getConfigValue_(ss, "DataFolder", "SpreadFinder/DATA");
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
      { folder: optionPricesFolder, name: "tsla-options-exp-2028-06-16-monthly-show-all-stacked-02-04-2026.csv", path: "DATA/OptionPrices/tsla-options-exp-2028-06-16-monthly-show-all-stacked-02-04-2026.csv" },
      { folder: optionPricesFolder, name: "tsla-options-exp-2028-12-15-monthly-show-all-stacked-02-04-2026.csv", path: "DATA/OptionPrices/tsla-options-exp-2028-12-15-monthly-show-all-stacked-02-04-2026.csv" },
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

  // Load option prices from the downloaded CSV files
  try {
    refreshOptionPrices();
  } catch (e) {
    Logger.log("Option price refresh skipped: " + e.message);
  }

  // Activate the README sheet
  ss.setActiveSheet(readmeSheet);

  ui.alert("Project initialized with real TSLA option prices for June 2028 and December 2028 LEAPs.\n\nTry it now:\n  OptionTools > Run SpreadFinder\n  OptionTools > View SpreadFinder Graphs\n\nTo track your own positions:\n  OptionTools > Import Portfolio from E*Trade\n  OptionTools > Load Sample Portfolio");
}

/**
 * Loads sample portfolio data from the sample CSV files in GitHub.
 * Downloads the sample files if not present, then imports them.
 */
function loadSamplePortfolio() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const existing = ss.getSheetByName("Portfolio");
  if (existing) {
    const resp = ui.alert(
      "Load Sample Portfolio",
      "A Portfolio sheet already exists. Replace it with sample data?",
      ui.ButtonSet.OK_CANCEL
    );
    if (resp !== ui.Button.OK) return;
    ss.deleteSheet(existing);
    // Also remove the named range
    const nr = ss.getNamedRanges().find(r => r.getName() === "PortfolioTable");
    if (nr) nr.remove();
  }

  // Ensure sample CSV files are downloaded to Drive
  const dataFolderPath = getConfigValue_(ss, "DataFolder", "SpreadFinder/DATA");
  let etradeFolder;
  try {
    let driveFolder = DriveApp.getRootFolder();
    for (const part of dataFolderPath.split("/")) {
      driveFolder = getOrCreateFolder_(driveFolder, part);
    }
    etradeFolder = getOrCreateFolder_(driveFolder, "Etrade");

    const sampleFiles = [
      { name: "PortfolioDownload-sample.csv", path: "DATA/Etrade/PortfolioDownload-sample.csv" },
      { name: "DownloadTxnHistory-sample.csv", path: "DATA/Etrade/DownloadTxnHistory-sample.csv" },
    ];

    for (const sf of sampleFiles) {
      // Always refresh sample files to get latest from GitHub
      const existingFile = findFileByName_(etradeFolder, sf.name);
      if (existingFile) existingFile.setTrashed(true);
      const csv = fetchGitHubFile_(sf.path);
      if (csv) {
        etradeFolder.createFile(sf.name, csv, MimeType.CSV);
      } else {
        throw new Error("Could not download " + sf.name + " from GitHub");
      }
    }
  } catch (e) {
    ui.alert("Error setting up sample files:\n" + e.message);
    return;
  }

  // Import the sample portfolio using the standard import function
  // Pass the specific sample filenames to import
  try {
    importEtradePortfolioFromFolder_(etradeFolder, "DownloadTxnHistory-sample.csv", "PortfolioDownload-sample.csv");
  } catch (e) {
    ui.alert("Error importing sample portfolio:\n" + e.message);
    return;
  }

  // Activate the Portfolio sheet
  const portfolioSheet = ss.getSheetByName("Portfolio");
  if (portfolioSheet) {
    ss.setActiveSheet(portfolioSheet);
  }

  ui.alert("Sample portfolio loaded with multiple TSLA positions.\n\nTry:\n  OptionTools > View Portfolio Performance Graphs\n\nTo import your real positions:\n  OptionTools > Import Portfolio from E*Trade");
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
