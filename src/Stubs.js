/**
 * Stubs.js - Library wrapper stubs for SpreadFinder
 * Updated: 2026-02-18
 *
 * This file contains thin wrapper functions that delegate to the SpreadFinder library.
 * Copy this into the Apps Script editor of any spreadsheet that uses the SpreadFinder library.
 *
 * Library setup:
 *   1. Open Extensions > Apps Script
 *   2. Click "+" next to Libraries in the left sidebar
 *   3. Enter Script ID: 1qvAlZ99zluKSr3ws4NsxH8xXo1FncbWzu6Yq6raBumHdCLpLDaKveM0T
 *   4. Click "Look up" and select the latest version
 *   5. Set identifier to "SpreadFinder"
 *   6. Click "Add"
 *
 * Why stubs are needed:
 *   - Custom spreadsheet functions (used in cell formulas) must be defined locally
 *   - Triggers like onOpen() and onEdit() must be defined locally
 *   - Dialog callbacks (google.script.run) must be defined locally
 *   - Menu action functions must be callable by name from the menu
 */

// ============================================================
// RUNNER (ensures logs are flushed after every action)
// ============================================================

/**
 * Executes a library function with standardized error handling and log flushing.
 * Ensures buffered logs are written to the Logs sheet even on errors.
 * @param {Function} fn - The library function to execute.
 * @param {Array} [args] - Arguments to pass to the function.
 * @returns {*} The return value of the function.
 */
function runner_(fn, args) {
  try {
    return fn.apply(null, args || []);
  } catch (e) {
    // Log error if log is available
    if (SpreadFinder && SpreadFinder.log && typeof SpreadFinder.log.error === "function") {
      SpreadFinder.log.error("Runner", "Error in " + (fn.name || "anonymous") + ": " + e.message);
    }
    SpreadsheetApp.getUi().alert("An error occurred: " + e.message + "\n\nCheck the 'Logs' sheet for details.");
    throw e;
  } finally {
    if (SpreadFinder && typeof SpreadFinder.flush === "function") {
      SpreadFinder.flush();
    }
  }
}

/**
 * Lightweight wrapper for custom functions (called from cells).
 * Flushes logs but no UI alerts (would be disruptive for cell formulas).
 * @param {Function} fn - The library function to execute.
 * @param {Array} [args] - Arguments to pass to the function.
 * @returns {*} The return value of the function.
 */
function customFn_(fn, args) {
  try {
    return fn.apply(null, args || []);
  } finally {
    if (SpreadFinder && typeof SpreadFinder.flush === "function") {
      SpreadFinder.flush();
    }
  }
}

// ============================================================
// TRIGGERS (must be local)
// ============================================================

/**
 * Trigger that runs when the spreadsheet is opened.
 * Sets up the OptionTools menu.
 * @param {Object} e - The onOpen event object.
 */
function onOpen(e) {
  SpreadFinder.setupSpreadFinderMenu(e);
}

// ============================================================
// MENU ACTIONS (called by name from OptionTools menu)
// ============================================================

/**
 * Initializes the project, creating required sheets and Drive folders.
 */
function initializeProject() {
  runner_(SpreadFinder.initializeProject);
}

/**
 * Configures logging settings for this document.
 * Edit this function to customize log level and filtered functions.
 */
function setupLogging() {
  configureLogging({
    level: "INFO",
    mode: "blacklist",
    functions: ["quotes", "positions"]
  });
  showLogConfig();
}

/**
 * Warms the XLookupByKeys cache for faster lookups.
 */
function warmXLookupCache() {
  runner_(SpreadFinder.warmXLookupCache);
}

/**
 * Generates portfolio value vs price charts for all symbols.
 */
function PlotPortfolioValueByPrice() {
  runner_(SpreadFinder.PlotPortfolioValueByPrice);
}

/**
 * Imports the latest transactions from Drive.
 */
function importLatestTransactions() {
  runner_(SpreadFinder.importLatestTransactions);
}

/**
 * Rebuilds the portfolio from E*Trade CSV files in Drive.
 */
function rebuildPortfolio() {
  runner_(SpreadFinder.rebuildPortfolio);
}

/**
 * Loads sample portfolio data for demonstration.
 */
function loadSamplePortfolio() {
  runner_(SpreadFinder.loadSamplePortfolio);
}

/**
 * Refreshes option prices from CSV files in Drive.
 */
function refreshOptionPrices() {
  runner_(SpreadFinder.refreshOptionPrices);
}

/**
 * Runs SpreadFinder analysis to find attractive spreads.
 */
function runSpreadFinder() {
  runner_(SpreadFinder.runSpreadFinder);
}

/**
 * Shows the SpreadFinder results in interactive charts.
 */
function showSpreadFinderGraphs() {
  runner_(SpreadFinder.showSpreadFinderGraphs);
}

/**
 * Shows the file upload dialog for option prices.
 */
function showUploadOptionPricesDialog() {
  runner_(SpreadFinder.showUploadOptionPricesDialog);
}

/**
 * Shows the file upload dialog for portfolio rebuild.
 */
function showUploadRebuildDialog() {
  runner_(SpreadFinder.showUploadRebuildDialog);
}

// ============================================================
// DIALOG CALLBACKS (called via google.script.run from HTML)
// ============================================================

/**
 * Gets list of symbols available in the portfolio.
 * @returns {string[]} Array of ticker symbols.
 */
function getAvailableSymbols() {
  return runner_(SpreadFinder.getAvailableSymbols);
}

/**
 * Plots performance charts for selected symbols.
 * @param {string[]} symbols - Array of ticker symbols to plot.
 */
function plotSelectedSymbols(symbols) {
  return runner_(SpreadFinder.plotSelectedSymbols, [symbols]);
}

/**
 * Gets data for SpreadFinder graphs.
 * @returns {Object} Graph data for rendering charts.
 */
function getSpreadFinderGraphData() {
  return runner_(SpreadFinder.getSpreadFinderGraphData);
}

/**
 * Gets available options for SpreadFinder selection dialog.
 * @returns {Object} Available symbols and expirations.
 */
function getSpreadFinderOptions() {
  return runner_(SpreadFinder.getSpreadFinderOptions);
}

/**
 * Runs SpreadFinder with user-selected symbols and expirations.
 * @param {string[]} symbols - Selected ticker symbols.
 * @param {string[]} expirations - Selected expiration dates.
 * @returns {Object} Analysis results.
 */
function runSpreadFinderWithSelection(symbols, expirations) {
  return runner_(SpreadFinder.runSpreadFinderWithSelection, [symbols, expirations]);
}

/**
 * Gets data for portfolio performance graphs.
 * @returns {Object} Graph data for rendering portfolio charts.
 */
function getPortfolioGraphData() {
  return runner_(SpreadFinder.getPortfolioGraphData);
}

/**
 * Uploads option price CSV files and refreshes prices.
 * @param {Array<{name: string, content: string}>} files - Array of file objects.
 * @returns {string} Status message.
 */
function uploadOptionPrices(files) {
  return runner_(SpreadFinder.uploadOptionPrices, [files]);
}

/**
 * Uploads portfolio and transaction files, then rebuilds the portfolio.
 * @param {{name: string, content: string}} portfolio - Portfolio CSV file.
 * @param {{name: string, content: string}} transactions - Transaction history CSV file.
 * @returns {string} Status message.
 */
function uploadAndRebuildPortfolio(portfolio, transactions) {
  return runner_(SpreadFinder.uploadAndRebuildPortfolio, [portfolio, transactions]);
}

/**
 * Completes project initialization with optional data loading.
 * @param {boolean} loadOptionPrices - Whether to load option prices from Drive.
 * @param {boolean} loadPortfolio - Whether to load portfolio from Drive.
 * @returns {string} Status message.
 */
function completeInitialization(loadOptionPrices, loadPortfolio) {
  return runner_(SpreadFinder.completeInitialization, [loadOptionPrices, loadPortfolio]);
}

/**
 * Refreshes portfolio formulas after option prices are updated.
 * @returns {string} Status message.
 */
function refreshPortfolioPrices() {
  return runner_(SpreadFinder.refreshPortfolioPrices);
}

// ============================================================
// CUSTOM SPREADSHEET FUNCTIONS (used in cell formulas)
// ============================================================

/**
 * Multi-key lookup with caching. Looks up values in a sheet using multiple key columns.
 * @param {Array} keyValues - Values to match against key columns.
 * @param {Array} keyHeaders - Column headers for the key columns.
 * @param {Array} returnHeaders - Column headers for the values to return.
 * @param {string} sheetName - Name of the sheet to search.
 * @returns {Array} The matching values from the return columns.
 * @customfunction
 */
function XLookupByKeys(keyValues, keyHeaders, returnHeaders, sheetName) {
  return customFn_(SpreadFinder.XLookupByKeys, [keyValues, keyHeaders, returnHeaders, sheetName]);
}

/**
 * Two-key lookup. Finds a row matching two key values and returns a value from another column.
 * @param {*} key1 - First key value to match.
 * @param {*} key2 - Second key value to match.
 * @param {Range} col1 - Range containing first key column.
 * @param {Range} col2 - Range containing second key column.
 * @param {Range} returnCol - Range containing values to return.
 * @returns {*} The matching value from returnCol.
 * @customfunction
 */
function X2LOOKUP(key1, key2, col1, col2, returnCol) {
  return customFn_(SpreadFinder.X2LOOKUP, [key1, key2, col1, col2, returnCol]);
}

/**
 * Three-key lookup. Finds a row matching three key values and returns a value from another column.
 * @param {*} key1 - First key value to match.
 * @param {*} key2 - Second key value to match.
 * @param {*} key3 - Third key value to match.
 * @param {Range} key1Col - Range containing first key column.
 * @param {Range} key2Col - Range containing second key column.
 * @param {Range} key3Col - Range containing third key column.
 * @param {Range} returnCol - Range containing values to return.
 * @returns {*} The matching value from returnCol.
 * @customfunction
 */
function X3LOOKUP(key1, key2, key3, key1Col, key2Col, key3Col, returnCol) {
  return customFn_(SpreadFinder.X3LOOKUP, [key1, key2, key3, key1Col, key2Col, key3Col, returnCol]);
}

/**
 * Detects option strategy from legs (e.g., bull-call-spread, iron-condor).
 * @param {Range} strikeRange - Range of strike prices.
 * @param {Range} typeRange - Range of option types (Call/Put).
 * @param {Range} qtyRange - Range of quantities (positive=long, negative=short).
 * @param {Range} [_labels] - Optional labels for cache busting.
 * @returns {string} The detected strategy name.
 * @customfunction
 */
function detectStrategy(strikeRange, typeRange, qtyRange, _labels) {
  return customFn_(SpreadFinder.detectStrategy, [strikeRange, typeRange, qtyRange, _labels]);
}

/**
 * Builds OptionStrat URL from leg ranges for visualization.
 * @param {Range} symbolRange - Range containing ticker symbol.
 * @param {Range} strikeRange - Range of strike prices.
 * @param {Range} typeRange - Range of option types.
 * @param {Range} expirationRange - Range of expiration dates.
 * @param {Range} qtyRange - Range of quantities.
 * @param {Range} priceRange - Range of prices.
 * @returns {string} URL to optionstrat.com.
 * @customfunction
 */
function buildOptionStratUrlFromLegs(symbolRange, strikeRange, typeRange, expirationRange, qtyRange, priceRange) {
  return customFn_(SpreadFinder.buildOptionStratUrlFromLegs, [symbolRange, strikeRange, typeRange, expirationRange, qtyRange, priceRange]);
}

/**
 * Builds OptionStrat URL from parameters.
 * @param {Array} strikes - Array of strike prices.
 * @param {string} ticker - Ticker symbol.
 * @param {string} strategy - Strategy name.
 * @param {Date|string} expiration - Expiration date.
 * @returns {string} URL to optionstrat.com.
 * @customfunction
 */
function buildOptionStratUrl(strikes, ticker, strategy, expiration) {
  return customFn_(SpreadFinder.buildOptionStratUrl, [strikes, ticker, strategy, expiration]);
}

/**
 * Builds OptionStrat URL for custom multi-leg positions.
 * @param {string} symbol - Ticker symbol.
 * @param {Array} legs - Array of leg objects with strike, type, qty, expiration.
 * @returns {string} URL to optionstrat.com.
 * @customfunction
 */
function buildCustomOptionStratUrl(symbol, legs) {
  return customFn_(SpreadFinder.buildCustomOptionStratUrl, [symbol, legs]);
}

/**
 * Recommends debit to open a bull call spread based on bid/ask prices.
 * @param {string} symbol - Ticker symbol.
 * @param {Date|string} expiration - Expiration date.
 * @param {number} lowerStrike - Lower strike price (long call).
 * @param {number} upperStrike - Upper strike price (short call).
 * @param {number} [avgMinutesToExecute] - Expected minutes to fill order.
 * @param {Range} [_labels] - Optional labels for cache busting.
 * @returns {number} Recommended debit per contract.
 * @customfunction
 */
function recommendBullCallSpreadOpenDebit(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute, _labels) {
  return customFn_(SpreadFinder.recommendBullCallSpreadOpenDebit, [symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute, _labels]);
}

/**
 * Recommends credit to close a bull call spread based on bid/ask prices.
 * @param {string} symbol - Ticker symbol.
 * @param {Date|string} expiration - Expiration date.
 * @param {number} lowerStrike - Lower strike price (long call).
 * @param {number} upperStrike - Upper strike price (short call).
 * @param {number} [avgMinutesToExecute] - Expected minutes to fill order.
 * @param {Range} [_labels] - Optional labels for cache busting.
 * @returns {number} Recommended credit per contract.
 * @customfunction
 */
function recommendBullCallSpreadCloseCredit(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute, _labels) {
  return customFn_(SpreadFinder.recommendBullCallSpreadCloseCredit, [symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute, _labels]);
}

/**
 * Recommends credit to open an iron condor based on bid/ask prices.
 * @param {string} symbol - Ticker symbol.
 * @param {Date|string} expiration - Expiration date.
 * @param {number} putLower - Lower put strike (long put).
 * @param {number} putUpper - Upper put strike (short put).
 * @param {number} callLower - Lower call strike (short call).
 * @param {number} callUpper - Upper call strike (long call).
 * @param {number} [avgMinutesToExecute] - Expected minutes to fill order.
 * @param {Range} [_labels] - Optional labels for cache busting.
 * @returns {number} Recommended credit per contract.
 * @customfunction
 */
function recommendIronCondorOpenCredit(symbol, expiration, putLower, putUpper, callLower, callUpper, avgMinutesToExecute, _labels) {
  return customFn_(SpreadFinder.recommendIronCondorOpenCredit, [symbol, expiration, putLower, putUpper, callLower, callUpper, avgMinutesToExecute, _labels]);
}

/**
 * Recommends debit to close an iron condor based on bid/ask prices.
 * @param {string} symbol - Ticker symbol.
 * @param {Date|string} expiration - Expiration date.
 * @param {number} putLower - Lower put strike (long put).
 * @param {number} putUpper - Upper put strike (short put).
 * @param {number} callLower - Lower call strike (short call).
 * @param {number} callUpper - Upper call strike (long call).
 * @param {number} [avgMinutesToExecute] - Expected minutes to fill order.
 * @param {Range} [_labels] - Optional labels for cache busting.
 * @returns {number} Recommended debit per contract.
 * @customfunction
 */
function recommendIronCondorCloseDebit(symbol, expiration, putLower, putUpper, callLower, callUpper, avgMinutesToExecute, _labels) {
  return customFn_(SpreadFinder.recommendIronCondorCloseDebit, [symbol, expiration, putLower, putUpper, callLower, callUpper, avgMinutesToExecute, _labels]);
}

/**
 * Recommends opening price for a single option leg.
 * @param {string} symbol - Ticker symbol.
 * @param {Date|string} expiration - Expiration date.
 * @param {number} strike - Strike price.
 * @param {string} type - Option type ("Call" or "Put").
 * @param {number} qty - Quantity (positive=buy, negative=sell).
 * @param {number} [avgMinutesToExecute] - Expected minutes to fill order.
 * @returns {number} Recommended price per contract.
 * @customfunction
 */
function recommendOpen(symbol, expiration, strike, type, qty, avgMinutesToExecute) {
  return customFn_(SpreadFinder.recommendOpen, [symbol, expiration, strike, type, qty, avgMinutesToExecute]);
}

/**
 * Recommends closing price for a single option leg.
 * @param {string} symbol - Ticker symbol.
 * @param {Date|string} expiration - Expiration date.
 * @param {number} strike - Strike price.
 * @param {string} type - Option type ("Call" or "Put").
 * @param {number} qty - Quantity (positive=buy, negative=sell).
 * @param {number} [avgMinutesToExecute] - Expected minutes to fill order.
 * @returns {number} Recommended price per contract.
 * @customfunction
 */
function recommendClose(symbol, expiration, strike, type, qty, avgMinutesToExecute) {
  return customFn_(SpreadFinder.recommendClose, [symbol, expiration, strike, type, qty, avgMinutesToExecute]);
}

/**
 * Returns first non-empty value from a range.
 * @param {Range} range - Range to search for non-empty value.
 * @returns {*} The first non-empty value, or empty string if none found.
 * @customfunction
 */
function coalesce(range) {
  return customFn_(SpreadFinder.coalesce, [range]);
}

/**
 * Formats option legs as description with negative prefixes for shorts.
 * @param {Range} strikeRange - Range of strike prices.
 * @param {Range} qtyRange - Range of quantities.
 * @param {string} [suffix] - Optional suffix (e.g., "BCS", "BPS").
 * @returns {string} Formatted description like "300/-400 BCS".
 * @customfunction
 */
function formatLegsDescription(strikeRange, qtyRange, suffix) {
  return customFn_(SpreadFinder.formatLegsDescription, [strikeRange, qtyRange, suffix]);
}

/**
 * Formats descriptions for all position groups in one call.
 * Place formula in first data row of Description column - it fills down automatically.
 * @param {Range} groups - Range of group identifiers.
 * @param {Range} strikes - Range of strike prices.
 * @param {Range} qtys - Range of quantities.
 * @param {Range} strategies - Range of strategy names.
 * @returns {Array} 2D array of descriptions (one per row, first row of each group has value).
 * @customfunction
 */
function formatAllDescriptions(groups, strikes, qtys, strategies) {
  return customFn_(SpreadFinder.formatAllDescriptions, [groups, strikes, qtys, strategies]);
}
