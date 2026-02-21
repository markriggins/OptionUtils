/**
 * SpreadFinder.js
 * Analyzes OptionPricesUploaded to find and rank bull call spread opportunities.
 *
 * Config lives on the SpreadFinderConfig sheet.
 * Results are written to the Spreads sheet.
 *
 * Related files:
 * - SpreadFinderInit.js: Initialization, loading, and output functions
 * - SpreadFinderCalc.js: Calculation functions
 *
 * Version: 2.1
 */

/* =========================================================
   Entry points (called from menu/UI)
   ========================================================= */

/**
 * Runs SpreadFinder using config symbols or all available symbols/expirations.
 */
function runSpreadFinder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check if config sheet already exists (to show first-run message)
  const isFirstRun = !ss.getSheetByName(SPREAD_FINDER_CONFIG_SHEET);

  // Get available symbols and expirations from option prices
  const options = getSpreadFinderOptions();

  // Use config symbols if set, otherwise use all available
  const configSheet = ss.getSheetByName(SPREAD_FINDER_CONFIG_SHEET);
  const config = configSheet ? loadSpreadFinderConfig_(configSheet) : {};

  const symbols = (config.symbols && config.symbols.length > 0)
    ? config.symbols
    : options.symbols;

  // Use all expirations
  const expirations = options.expirations.map(e => e.value);

  // Run with these selections
  runSpreadFinderWithSelection(symbols, expirations);

  // Show first-run message
  if (isFirstRun) {
    SpreadsheetApp.getUi().alert(
      "SpreadFinder Initialized",
      "SpreadFinderConfig sheet created with default settings.\n\n" +
      "You can modify it to customize:\n" +
      "• Target ROI and max debit\n" +
      "• Strike range and spread width\n" +
      "• Months to expiration\n" +
      "• Minimum volume/open interest\n\n" +
      "Run SpreadFinder again after making changes.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Runs SpreadFinder with the selected symbols and expirations.
 * @param {string[]} symbols - Selected symbols
 * @param {string[]} expirations - Selected expiration dates (YYYY-MM-DD format)
 */
function runSpreadFinderWithSelection(symbols, expirations) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ensure config sheet exists, load config
  const configSheet = ensureSpreadFinderConfigSheet_(ss);
  const config = loadSpreadFinderConfig_(configSheet);

  // Override config with selection
  config.symbols = symbols;
  config.selectedExpirations = new Set(expirations);

  // Name results sheet after symbol(s)
  const spreadsSheetName = symbols.length > 0
    ? symbols.join(",") + "Spreads"
    : SPREADS_SHEET;
  const sheet = ensureSpreadsSheet_(ss, spreadsSheetName);
  log.debug("spreadFinder", "Config: " + JSON.stringify(config));

  // Load option data (filtered by selection)
  const options = loadOptionData_(ss, config.symbols, config.selectedExpirations);
  log.info("spreadFinder", "Loaded " + options.length + " options");

  // Filter to calls only
  const calls = options.filter(o => o.type === "Call");
  log.debug("spreadFinder", "Filtered to " + calls.length + " calls");

  // Group by symbol+expiration
  const grouped = groupBySymbolExpiration_(calls);

  // Derive current price from ATM calls (delta closest to 0.5)
  const currentPrice = estimateCurrentPrice_(calls);
  log.debug("spreadFinder", "Estimated current price: " + currentPrice);

  // Default outlook if not set by user
  if (!config.outlookFuturePrice) {
    config.outlookFuturePrice = roundTo_(currentPrice * 1.25, 2);
    log.debug("spreadFinder", "Defaulting outlookFuturePrice to " + config.outlookFuturePrice);
  }
  if (!config.outlookConfidence) {
    config.outlookConfidence = 0.5;
  }
  if (!config.outlookDate) {
    // Default to 18 months from now
    const d = new Date();
    d.setMonth(d.getMonth() + 18);
    config.outlookDate = d;
  }

  // Generate and score all spreads
  const spreads = [];
  for (const key of Object.keys(grouped)) {
    const chain = grouped[key];
    const chainSpreads = generateSpreads_(chain, config);
    spreads.push(...chainSpreads);
  }
  log.info("spreadFinder", "Generated " + spreads.length + " spreads");

  // Load held positions from Positions sheet
  const conflicts = loadHeldPositions_(ss);
  log.debug("spreadFinder", "Loaded " + conflicts.size + " held positions");

  // Filter by config constraints, mark conflicts instead of removing
  // Skip expiration date range filter if user selected specific expirations
  const minExpDate = config.minExpirationDate;
  const maxExpDate = config.maxExpirationDate;
  const skipExpDateFilter = !!config.selectedExpirations;
  const filtered = spreads.filter(s => {
    const expDate = parseDateAtMidnight_(s.expiration);
    // Mark conflicts as held (but keep them in results)
    s.held = conflicts.has(`${s.symbol}|${s.lowerStrike}|${s.expiration}`);
    if (config.symbols && !config.symbols.includes(s.symbol)) return false;
    return s.debit > 0 &&
      s.roi >= config.minROI &&
      s.liquidityScore >= config.minLiquidityScore &&
      s.lowerStrike >= config.minStrike &&
      s.upperStrike <= config.maxStrike &&
      (skipExpDateFilter || (expDate >= minExpDate && expDate <= maxExpDate));
  });
  log.info("spreadFinder", "Filtered to " + filtered.length + " spreads meeting criteria");

  // Sort by fitness (descending)
  filtered.sort((a, b) => b.fitness - a.fitness);

  // Output results to same sheet below config
  outputSpreadResults_(sheet, filtered, config);

  // Debug info
  const debugMsg = `Options loaded: ${options.length}\n` +
    `Calls found: ${calls.length}\n` +
    `Spreads generated: ${spreads.length}\n` +
    `After filtering: ${filtered.length}`;

  SpreadsheetApp.getUi().alert(
    "SpreadFinder Complete",
    debugMsg,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Opens a large dashboard window with Delta vs ROI and Strike vs ROI.
 * If no spreads data exists or config has changed, runs SpreadFinder first.
 */
function showSpreadFinderGraphs() {
  SpreadsheetApp.flush();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SPREAD_FINDER_CONFIG_SHEET);
  const config = configSheet ? loadSpreadFinderConfig_(configSheet) : {};
  const spreadsSheetName = config.symbols && config.symbols.length > 0
    ? config.symbols.join(",") + "Spreads"
    : SPREADS_SHEET;
  let sheet = ss.getSheetByName(spreadsSheetName);

  // Check if we need to run SpreadFinder
  let needsRun = false;
  if (!sheet || sheet.getLastRow() < 3) {
    // No spreads data
    needsRun = true;
  } else {
    // Check if config has changed since last run
    const currentHash = computeConfigHash_(config);
    const storedHash = getStoredConfigHash_(sheet);
    if (currentHash !== storedHash) {
      needsRun = true;
    }
  }

  if (needsRun) {
    runSpreadFinder();
    SpreadsheetApp.flush();
  }

  // Creates the SpreadFinderGraphs modal dialog
  const html = HtmlService.createHtmlOutputFromFile('ui/SpreadFinderGraphs')
      .setWidth(1050) // Wide enough for side-by-side or large stacked charts
      .setHeight(850);

  SpreadsheetApp.getUi().showModalDialog(html, 'Spread Finder Graphs');
}

/**
 * Fetches spread data for SpreadFinderGraphs.
 * Orders by Fitness so the best points are drawn last (on top).
 */
function getSpreadFinderGraphData() {
  log.debug("spreadFinder", "getSpreadFinderGraphData called");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SPREAD_FINDER_CONFIG_SHEET);
  const config = configSheet ? loadSpreadFinderConfig_(configSheet) : {};
  const spreadsSheetName = config.symbols && config.symbols.length > 0
    ? config.symbols.join(",") + "Spreads"
    : SPREADS_SHEET;
  const sheet = ss.getSheetByName(spreadsSheetName);
  const lastRow = sheet.getLastRow();
  const headerRow = 2; // Row 1=timestamp, Row 2=headers
  const startRow = 3;  // Row 3+=data
  if (lastRow < startRow) return [];

  // Build column index from header row
  const hdrs = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const c = {};
  hdrs.forEach((h, i) => c[h.toString().trim()] = i);

  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, hdrs.length).getValues();
  const today = new Date();
  today.setHours(0,0,0,0);

  return data.map(row => {
    const sym = row[c.Symbol];
    // Parse expiration to local midnight to avoid timezone shifts
    const expDate = parseDateAtMidnight_(row[c.Expiration]);
    const lowStrike = row[c.Lower];
    const highStrike = row[c.Upper];

    const osUrl = buildOptionStratUrl(`${lowStrike}/${highStrike}`, sym, "bull-call-spread", expDate);

    const diffTime = expDate.getTime() - today.getTime();
    const dte = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

    return {
      delta: parseFloat(row[c.LowerDelta]) || 0,
      roi: parseFloat(row[c.ROI]) || 0,
      strike: parseFloat(row[c.Lower]) || 0,
      fitness: parseFloat(row[c.Fitness]) || 0,
      label: String(row[c.Label] || ""),
      osUrl: osUrl,

      width: row[c.Width],
      debit: row[c.Debit],
      maxProfit: row[c.MaxProfit],
      expectedGain: row[c.ExpGain],
      expectedROI: row[c.ExpROI],
      lowerDelta: row[c.LowerDelta],
      upperDelta: row[c.UpperDelta],
      lowerOI: row[c.LowerOI],
      upperOI: row[c.UpperOI],
      liquidity: row[c.Liquidity],
      dte: dte > 0 ? dte : 0,
      held: (row[c.Held] || "").toString().trim() === "HELD",
      iv: parseFloat(row[c.IV]) || 0
    };
  }).sort((a, b) => a.fitness - b.fitness);
}
