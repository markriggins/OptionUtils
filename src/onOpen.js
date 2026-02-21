/**
 * Add AppScript menu items to a google sheet.
 * OptionTools menu provides access to SpreadFinder and Portfolio features.
 */
function setupSpreadFinderMenu() {
  const ui = SpreadsheetApp.getUi();

  const spreadFinderMenu = ui.createMenu('SpreadFinder')
    .addItem('Upload Option Prices...', 'showUploadOptionPricesDialog')
    .addSeparator()
    .addItem('Run SpreadFinder', 'runSpreadFinder')
    .addItem('View Graphs', 'showSpreadFinderGraphs');

  const portfolioMenu = ui.createMenu('Portfolio')
    .addItem('Upload Portfolio/Transactions...', 'showUploadRebuildDialog')
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
 * Shows the initialization dialog with options to clear the project.
 */
function initializeProject() {
  const ss = SpreadsheetApp.getActive();

  // Find sheets that will be deleted (everything except README)
  const keepSheets = ["README"];
  const sheetsToDelete = ss.getSheets().filter(s => !keepSheets.includes(s.getName())).map(s => s.getName());

  // Show dialog
  const html = HtmlService.createHtmlOutputFromFile("ui/InitializeDialog")
    .setWidth(450)
    .setHeight(300);

  const initData = JSON.stringify({
    sheetsToDelete: sheetsToDelete
  });

  const content = html.getContent().replace(
    '</body>',
    `<script>init(${initData});</script></body>`
  );

  const output = HtmlService.createHtmlOutput(content)
    .setWidth(450)
    .setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(output, "Initialize / Clear Project");
}

/**
 * Completes the initialization process (called from dialog).
 * Clears all sheets except README and removes named ranges.
 * @returns {string} Status message
 */
function completeInitialization() {
  const ss = SpreadsheetApp.getActive();

  // Find and delete sheets (everything except README)
  const keepSheets = ["README"];
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
    readmeSheet = ss.insertSheet("README", 0);
  }

  const readmeRows = [
    ["OptionUtils - Option Portfolio Analysis Tools"],
    ["Analyze option spreads, track portfolios, and visualize profit/loss scenarios."],
    [""],
    ["--- Quick Start: SpreadFinder ---"],
    ["  1. Download option prices from barchart.com: Options > Select expiration > Stacked view > Download CSV"],
    ["  2. Run OptionTools > SpreadFinder > Upload Option Prices"],
    ["  3. Run OptionTools > SpreadFinder > Run SpreadFinder"],
    ["  4. Run OptionTools > SpreadFinder > View Graphs"],
    [""],
    ["--- Quick Start: Portfolio ---"],
    ["  1. Download Portfolio CSV and Transaction History CSV from your brokerage"],
    ["  2. Run OptionTools > Portfolio > Upload Portfolio/Transactions"],
    ["  3. Run OptionTools > Portfolio > View Performance Graphs"],
    [""],
    ["--- Apps Script Library ---"],
    ["Script ID: 1qvAlZ99zluKSr3ws4NsxH8xXo1FncbWzu6Yq6raBumHdCLpLDaKveM0T"],
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

  // Activate the README sheet
  ss.setActiveSheet(readmeSheet);

  return "Project initialized.\n\nTo get started:\n  OptionTools > SpreadFinder > Upload Option Prices";
}

