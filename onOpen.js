/**
 *  Add AppScript menu items to a google sheet
 *  OptionTools/
 *    RefreshOptionPrices
 *    Warm XLookup Cache
 **/

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const spreadFinderMenu = ui.createMenu('SpreadFinder')
    .addItem('Upload & Refresh...', 'showUploadOptionPricesDialog')
    .addItem('Refresh from Drive', 'refreshOptionPrices')
    .addSeparator()
    .addItem('Run SpreadFinder', 'runSpreadFinder')
    .addItem('View Graphs', 'showSpreadFinderGraphs');

  const portfolioMenu = ui.createMenu('Portfolio')
    .addItem('Upload & Rebuild...', 'showUploadRebuildDialog')
    .addItem('Rebuild from Drive', 'rebuildPortfolio')
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
 * Shows the initialization dialog with options to load existing data.
 */
function initializeProject() {
  const ss = SpreadsheetApp.getActive();

  // Find sheets that will be deleted (everything except README and Config)
  const keepSheets = ["README", "Config"];
  const sheetsToDelete = ss.getSheets().filter(s => !keepSheets.includes(s.getName())).map(s => s.getName());

  // Check for existing data in Drive folders
  const dataFolderPath = getConfigValue_(ss, "DataFolder", "SpreadFinder/DATA");
  let optionPricesCount = 0;
  let portfolioCount = 0;

  try {
    let driveFolder = DriveApp.getRootFolder();
    for (const part of dataFolderPath.split("/")) {
      const it = driveFolder.getFoldersByName(part);
      if (!it.hasNext()) break;
      driveFolder = it.next();
    }

    // Check OptionPrices folder
    const opIt = driveFolder.getFoldersByName("OptionPrices");
    if (opIt.hasNext()) {
      const opFolder = opIt.next();
      const files = opFolder.getFilesByType(MimeType.CSV);
      while (files.hasNext()) { files.next(); optionPricesCount++; }
    }

    // Check Etrade folder
    const etIt = driveFolder.getFoldersByName("Etrade");
    if (etIt.hasNext()) {
      const etFolder = etIt.next();
      const files = etFolder.getFilesByType(MimeType.CSV);
      while (files.hasNext()) { files.next(); portfolioCount++; }
    }
  } catch (e) {
    log.warn("init", "Error checking for existing data: " + e.message);
  }

  // Show dialog with options
  const html = HtmlService.createHtmlOutputFromFile("InitializeDialog")
    .setWidth(450)
    .setHeight(400);

  // Inject the initialization data
  const initData = JSON.stringify({
    sheetsToDelete: sheetsToDelete,
    optionPricesCount: optionPricesCount,
    portfolioCount: portfolioCount
  });

  const content = html.getContent().replace(
    '</body>',
    `<script>init(${initData});</script></body>`
  );

  const output = HtmlService.createHtmlOutput(content)
    .setWidth(450)
    .setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(output, "Initialize / Clear Project");
}

/**
 * Completes the initialization process (called from dialog).
 * @param {boolean} loadOptionPrices - Whether to load option prices after init
 * @param {boolean} loadPortfolio - Whether to load portfolio after init
 * @returns {string} Status message
 */
function completeInitialization(loadOptionPrices, loadPortfolio) {
  const ss = SpreadsheetApp.getActive();

  // Find and delete sheets (everything except README and Config)
  const keepSheets = ["README", "Config"];
  const sheetsToDelete = ss.getSheets().filter(s => !keepSheets.includes(s.getName()));
  for (const sheet of sheetsToDelete) {
    ss.deleteSheet(sheet);
  }

  // Remove any existing named ranges
  for (const nr of ss.getNamedRanges()) {
    nr.remove();
  }

  // Create README sheet
  let readmeSheet = ss.getSheetByName("README");
  if (readmeSheet) {
    readmeSheet.clearContents();
  } else {
    readmeSheet = ss.insertSheet("README", 0); // First tab
  }

  const readmeRows = [
    ["SpreadFinder - Find and Analyze Bull Call Spread Opportunities"],
    ["SpreadFinder scans option prices to find attractive bull call spread opportunities. It ranks spreads by expected ROI using a probability-of-touch model, filters by liquidity, and displays results in interactive charts."],
    [""],
    ["--- Quick Start ---"],
    ["  1. Download option prices from barchart.com: Options page > Select expiration > Stacked view > Download CSV"],
    ["  2. Run OptionTools > SpreadFinder > Upload & Refresh and select your CSV file(s)"],
    ["  3. Run OptionTools > SpreadFinder > Run SpreadFinder to analyze spreads"],
    ["  4. Run OptionTools > SpreadFinder > View Graphs for visual analysis"],
    [""],
    ["--- SpreadFinder Configuration ---"],
    ["The SpreadFinderConfig sheet controls the analysis parameters: symbol filter, min/max spread width, min open interest, min volume, max debit, min ROI, strike range, and expiration range. Adjust these to narrow down the results."],
    [""],
    ["--- Portfolio Modeling ---"],
    ["Track your actual positions and visualize profit/loss scenarios. Bull-Call-Spreads, Long Calls, and other strategies will be automatically detected and analyzed."],
    [""],
    ["To try with sample data:"],
    ["  1. Run OptionTools > Portfolio > Load Sample Portfolio"],
    ["  2. Run OptionTools > Portfolio > View Performance Graphs"],
    [""],
    ["To import your real positions:"],
    ["  1. Download Portfolio CSV and Transaction History CSV from your brokerage"],
    ["  2. Run OptionTools > Portfolio > Upload & Rebuild"],
    ["  3. Select your files and click Upload"],
    ["  4. Run OptionTools > Portfolio > View Performance Graphs"],
    [""],
    ["--- Apps Script Library ---"],
    ["Available as a Google Apps Script library. Script ID: 1qvAlZ99zluKSr3ws4NsxH8xXo1FncbWzu6Yq6raBumHdCLpLDaKveM0T"],
    ["Source: github.com/markriggins/OptionUtils"],
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

  // ---- Create Google Drive folder structure ----
  const dataFolderPath = getConfigValue_(ss, "DataFolder", "SpreadFinder/DATA");
  try {
    let driveFolder = DriveApp.getRootFolder();
    for (const part of dataFolderPath.split("/")) {
      driveFolder = getOrCreateFolder_(driveFolder, part);
    }
    getOrCreateFolder_(driveFolder, "Etrade");
    getOrCreateFolder_(driveFolder, "OptionPrices");
  } catch (e) {
    log.warn("init", "Drive folder setup skipped: " + e.message);
  }

  // Activate the README sheet
  ss.setActiveSheet(readmeSheet);

  // Load existing data if requested
  const loaded = [];

  if (loadOptionPrices) {
    try {
      refreshOptionPrices();
      loaded.push("Option Prices");
    } catch (e) {
      log.error("init", "Error loading option prices: " + e.message);
    }
  }

  if (loadPortfolio) {
    try {
      rebuildPortfolio();
      loaded.push("Portfolio");
    } catch (e) {
      log.error("init", "Error loading portfolio: " + e.message);
    }
  }

  // Return status message
  if (loaded.length > 0) {
    return "Project initialized. Loaded: " + loaded.join(", ");
  } else {
    return "Project initialized.\n\nTo get started, run:\n  OptionTools > SpreadFinder > Upload & Refresh";
  }
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

  ui.alert("Sample portfolio loaded with multiple TSLA positions.\n\nTry:\n  OptionTools > Portfolio > View Performance Graphs\n\nTo import your real positions:\n  OptionTools > Portfolio > Upload & Rebuild");
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
    log.warn("github", "GitHub fetch " + resp.getResponseCode() + " for " + path);
  } catch (e) {
    log.error("github", "GitHub fetch failed for " + path + ": " + e.message);
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
