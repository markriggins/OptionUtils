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

function onEdit(e) {
  SpreadFinder.onEdit(e);
}

// ============================================================
// MENU ACTIONS (called by name from OptionTools menu)
// ============================================================

function initializeProject() {
  SpreadFinder.initializeProject();
}

function warmXLookupCache() {
  SpreadFinder.warmXLookupCache();
}

function PlotPortfolioValueByPrice() {
  SpreadFinder.PlotPortfolioValueByPrice();
}

function importEtradePortfolio() {
  SpreadFinder.importEtradePortfolio();
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

function getPortfolioGraphData() {
  return SpreadFinder.getPortfolioGraphData();
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
function buildOptionStratUrlFromLegs(symbolRange, strikeRange, typeRange, expirationRange, qtyRange) {
  return SpreadFinder.buildOptionStratUrlFromLegs(symbolRange, strikeRange, typeRange, expirationRange, qtyRange);
}

/**
 * Builds OptionStrat URL from parameters.
 * @customfunction
 */
function buildOptionStratUrl(strikes, ticker, strategy, expiration) {
  return SpreadFinder.buildOptionStratUrl(strikes, ticker, strategy, expiration);
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
