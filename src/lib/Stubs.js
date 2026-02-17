// @ts-check
/**
 * Stubs.js - Library wrapper stubs for SpreadFinder
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
// TRIGGERS (must be local)
// ============================================================

function onOpen(e) {
  SpreadFinder.onOpen(e);
}

// ============================================================
// MENU ACTIONS (called by name from OptionTools menu)
// ============================================================

function initializeProject() {
  SpreadFinder.initializeProject();
}

// Setup logging
function setupLogging() {
  configureLogging({
    level: "INFO",
    mode: "blacklist",
    functions: ["quotes", "positions"]
  });
  showLogConfig();
}

function warmXLookupCache() {
  SpreadFinder.warmXLookupCache();
}

function PlotPortfolioValueByPrice() {
  SpreadFinder.PlotPortfolioValueByPrice();
}

function importLatestTransactions() {
  SpreadFinder.importLatestTransactions();
}

function rebuildPortfolio() {
  SpreadFinder.rebuildPortfolio();
}

function loadSamplePortfolio() {
  SpreadFinder.loadSamplePortfolio();
}

function refreshOptionPrices() {
  SpreadFinder.refreshOptionPrices();
}

function runSpreadFinder() {
  SpreadFinder.runSpreadFinder();
}

function showSpreadFinderGraphs() {
  SpreadFinder.showSpreadFinderGraphs();
}

function showUploadOptionPricesDialog() {
  SpreadFinder.showUploadOptionPricesDialog();
}

function showUploadRebuildDialog() {
  SpreadFinder.showUploadRebuildDialog();
}

// ============================================================
// DIALOG CALLBACKS (called via google.script.run from HTML)
// ============================================================

function getAvailableSymbols() {
  return SpreadFinder.getAvailableSymbols();
}

function plotSelectedSymbols(symbols) {
  return SpreadFinder.plotSelectedSymbols(symbols);
}

function getSpreadFinderGraphData() {
  return SpreadFinder.getSpreadFinderGraphData();
}

function getSpreadFinderOptions() {
  return SpreadFinder.getSpreadFinderOptions();
}

function runSpreadFinderWithSelection(symbols, expirations) {
  return SpreadFinder.runSpreadFinderWithSelection(symbols, expirations);
}

function getPortfolioGraphData() {
  return SpreadFinder.getPortfolioGraphData();
}

function uploadOptionPrices(files) {
  return SpreadFinder.uploadOptionPrices(files);
}

function uploadAndRebuildPortfolio(portfolio, transactions) {
  return SpreadFinder.uploadAndRebuildPortfolio(portfolio, transactions);
}

function completeInitialization(loadOptionPrices, loadPortfolio) {
  return SpreadFinder.completeInitialization(loadOptionPrices, loadPortfolio);
}

function refreshPortfolioPrices() {
  return SpreadFinder.refreshPortfolioPrices();
}

// ============================================================
// CUSTOM SPREADSHEET FUNCTIONS (used in cell formulas)
// ============================================================

/**
 * Multi-key lookup with caching.
 * @customfunction
 */
function XLookupByKeys(keyValues, keyHeaders, returnHeaders, sheetName) {
  return SpreadFinder.XLookupByKeys(keyValues, keyHeaders, returnHeaders, sheetName);
}

/**
 * Two-key lookup.
 * @customfunction
 */
function X2LOOKUP(key1, key2, col1, col2, returnCol) {
  return SpreadFinder.X2LOOKUP(key1, key2, col1, col2, returnCol);
}

/**
 * Three-key lookup.
 * @customfunction
 */
function X3LOOKUP(key1, key2, key3, key1Col, key2Col, key3Col, returnCol) {
  return SpreadFinder.X3LOOKUP(key1, key2, key3, key1Col, key2Col, key3Col, returnCol);
}

/**
 * Detects option strategy from legs.
 * @customfunction
 */
function detectStrategy(strikeRange, typeRange, qtyRange, _labels) {
  return SpreadFinder.detectStrategy(strikeRange, typeRange, qtyRange, _labels);
}

/**
 * Builds OptionStrat URL from leg ranges.
 * @customfunction
 */
function buildOptionStratUrlFromLegs(symbolRange, strikeRange, typeRange, expirationRange, qtyRange, priceRange) {
  return SpreadFinder.buildOptionStratUrlFromLegs(symbolRange, strikeRange, typeRange, expirationRange, qtyRange, priceRange);
}

/**
 * Builds OptionStrat URL from parameters.
 * @customfunction
 */
function buildOptionStratUrl(strikes, ticker, strategy, expiration) {
  return SpreadFinder.buildOptionStratUrl(strikes, ticker, strategy, expiration);
}

/**
 * Builds OptionStrat URL for custom multi-leg positions.
 */
function buildCustomOptionStratUrl(symbol, legs) {
  return SpreadFinder.buildCustomOptionStratUrl(symbol, legs);
}

/**
 * Recommends debit to open a bull call spread.
 * @customfunction
 */
function recommendBullCallSpreadOpenDebit(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute, _labels) {
  return SpreadFinder.recommendBullCallSpreadOpenDebit(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute, _labels);
}

/**
 * Recommends credit to close a bull call spread.
 * @customfunction
 */
function recommendBullCallSpreadCloseCredit(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute, _labels) {
  return SpreadFinder.recommendBullCallSpreadCloseCredit(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute, _labels);
}

/**
 * Recommends credit to open an iron condor.
 * @customfunction
 */
function recommendIronCondorOpenCredit(symbol, expiration, putLower, putUpper, callLower, callUpper, avgMinutesToExecute, _labels) {
  return SpreadFinder.recommendIronCondorOpenCredit(symbol, expiration, putLower, putUpper, callLower, callUpper, avgMinutesToExecute, _labels);
}

/**
 * Recommends debit to close an iron condor.
 * @customfunction
 */
function recommendIronCondorCloseDebit(symbol, expiration, putLower, putUpper, callLower, callUpper, avgMinutesToExecute, _labels) {
  return SpreadFinder.recommendIronCondorCloseDebit(symbol, expiration, putLower, putUpper, callLower, callUpper, avgMinutesToExecute, _labels);
}

/**
 * Recommends opening price for a single option leg.
 * @customfunction
 */
function recommendOpen(symbol, expiration, strike, type, qty, avgMinutesToExecute) {
  return SpreadFinder.recommendOpen(symbol, expiration, strike, type, qty, avgMinutesToExecute);
}

/**
 * Recommends closing price for a single option leg.
 * @customfunction
 */
function recommendClose(symbol, expiration, strike, type, qty, avgMinutesToExecute) {
  return SpreadFinder.recommendClose(symbol, expiration, strike, type, qty, avgMinutesToExecute);
}

/**
 * Returns first non-empty value from a range.
 * @customfunction
 */
function coalesce(range) {
  return SpreadFinder.coalesce(range);
}

/**
 * Formats option legs as description with negative prefixes for shorts.
 * @customfunction
 */
function formatLegsDescription(strikeRange, qtyRange, suffix) {
  return SpreadFinder.formatLegsDescription(strikeRange, qtyRange, suffix);
}

/**
 * Formats descriptions for all position groups in one call.
 * Place formula in first data row of Description column - it fills down automatically.
 * @customfunction
 */
function formatAllDescriptions(groups, strikes, qtys, strategies) {
  return SpreadFinder.formatAllDescriptions(groups, strikes, qtys, strategies);
}
