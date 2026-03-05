/**
 * SpreadFinder.js
 * Analyzes OptionPricesUploaded to find and rank bull call spread opportunities.
 *
 * Config is stored in hidden per-symbol sheets: _<Symbol>CallSpreadFinderConfig
 * Results are written to <Symbol>CallSpreads sheets.
 * Outlook data lives in the Outlook sheet.
 *
 * Related files:
 * - SpreadFinderInit.js: Initialization, loading, and output functions
 * - SpreadFinderCalc.js: Calculation functions
 *
 * Version: 3.0
 */

/* =========================================================
   Entry points (called from menu/UI)
   ========================================================= */

/**
 * Shows the Call Spread Finder modal dialog.
 */
function showCallSpreadFinderDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ui/CallSpreadFinderDialog')
    .setWidth(450)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Call Spread Finder');
}

/**
 * Gets data for the Call Spread Finder dialog.
 * Returns symbols, expirations per symbol, and saved config per symbol.
 */
function getCallSpreadFinderDialogData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(OPTION_PRICES_SHEET);

  if (!sheet) {
    throw new Error("No option prices loaded. Run 'Upload Option Prices' first.");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("No option data found in " + OPTION_PRICES_SHEET);
  }

  // Read header row to find column indices
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => h.toString().trim().toLowerCase());
  const symIdx = headers.indexOf("symbol");
  const expIdx = headers.indexOf("expiration");

  if (symIdx < 0 || expIdx < 0) {
    throw new Error("Required columns 'symbol' and 'expiration' not found");
  }

  // Read symbol and expiration columns
  const symCol = sheet.getRange(2, symIdx + 1, lastRow - 1, 1).getValues();
  const expCol = sheet.getRange(2, expIdx + 1, lastRow - 1, 1).getValues();

  // Build symbols set and expirations per symbol
  const symbols = new Set();
  const expirationsBySymbol = {}; // { TSLA: Map(key -> Date) }

  for (let i = 0; i < symCol.length; i++) {
    const sym = (symCol[i][0] || "").toString().trim().toUpperCase();
    if (!sym) continue;
    symbols.add(sym);

    const exp = expCol[i][0];
    if (exp) {
      const expDate = parseDateAtMidnight_(exp);
      if (expDate) {
        const key = expDate.getFullYear() + "-" +
          String(expDate.getMonth() + 1).padStart(2, "0") + "-" +
          String(expDate.getDate()).padStart(2, "0");
        if (!expirationsBySymbol[sym]) {
          expirationsBySymbol[sym] = new Map();
        }
        expirationsBySymbol[sym].set(key, expDate);
      }
    }
  }

  // Sort symbols
  const sortedSymbols = Array.from(symbols).sort();

  // Format expirations per symbol
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const formattedExpirationsBySymbol = {};
  for (const sym of sortedSymbols) {
    const expMap = expirationsBySymbol[sym] || new Map();
    formattedExpirationsBySymbol[sym] = Array.from(expMap.entries())
      .sort((a, b) => a[1] - b[1])
      .map(([key, date]) => ({
        value: key,
        label: months[date.getMonth()] + " " + date.getDate() + ", " + date.getFullYear()
      }));
  }

  // Load saved config for each symbol
  const configBySymbol = {};
  for (const sym of sortedSymbols) {
    configBySymbol[sym] = loadCallSpreadConfig_(ss, sym);
  }

  return {
    symbols: sortedSymbols,
    expirationsBySymbol: formattedExpirationsBySymbol,
    configBySymbol: configBySymbol
  };
}

/**
 * Runs Call Spread Finder for a single symbol with the given config.
 * Called from the dialog.
 * @param {string} symbol - Stock symbol
 * @param {Object} config - Config from dialog
 */
function runCallSpreadFinder(symbol, config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Save config to hidden sheet
  saveCallSpreadConfig_(ss, symbol, config);

  // Ensure Outlook sheet exists
  ensureOutlookSheet_(ss);

  // Parse selected expirations
  const selectedExpirations = new Set(
    (config.selectedExpirations || "").split(",").filter(Boolean)
  );

  if (selectedExpirations.size === 0) {
    throw new Error("No expirations selected");
  }

  // Load option data for this symbol and selected expirations
  const options = loadOptionData_(ss, [symbol], selectedExpirations);
  log.info("spreadFinder", "Loaded " + options.length + " options for " + symbol);

  // Filter to calls only
  const calls = options.filter(o => o.type === "Call");
  log.debug("spreadFinder", "Filtered to " + calls.length + " calls");

  if (calls.length === 0) {
    throw new Error("No call options found for " + symbol + " with selected expirations");
  }

  // Group by expiration
  const grouped = groupBySymbolExpiration_(calls);

  // Derive current price from ATM calls
  const currentPrice = estimateCurrentPrice_(calls);
  log.debug("spreadFinder", "Estimated current price: " + currentPrice);

  // Generate and score all spreads
  const spreads = [];
  for (const key of Object.keys(grouped)) {
    const chain = grouped[key];
    const expDate = parseDateAtMidnight_(chain[0].expiration);

    // Get outlook for this expiration
    const outlook = getOutlookForExpiration_(ss, symbol, expDate, currentPrice);

    const chainConfig = {
      ...config,
      currentPrice: currentPrice,
      outlook: outlook
    };

    const chainSpreads = generateCallSpreads_(chain, chainConfig);
    spreads.push(...chainSpreads);
  }
  log.info("spreadFinder", "Generated " + spreads.length + " spreads");

  // Load held positions
  const conflicts = loadHeldPositions_(ss);

  // Filter by config constraints
  const filtered = spreads.filter(s => {
    s.held = conflicts.has(`${s.symbol}|${s.lowerStrike}|${s.expiration}`);
    return s.debit > 0 &&
      s.roi >= config.minROI &&
      s.liquidityScore >= config.minLiquidityScore &&
      s.lowerStrike >= config.minStrike &&
      s.upperStrike <= config.maxStrike;
  });
  log.info("spreadFinder", "Filtered to " + filtered.length + " spreads meeting criteria");

  // Sort by fitness (descending)
  filtered.sort((a, b) => b.fitness - a.fitness);

  // Output to <Symbol>CallSpreads sheet
  const sheetName = symbol + "CallSpreads";
  const outputSheet = ensureSpreadsSheet_(ss, sheetName);
  outputSpreadResults_(outputSheet, filtered, config);

  // Show summary
  SpreadsheetApp.getUi().alert(
    "Call Spread Finder Complete",
    `Symbol: ${symbol}\n` +
    `Options loaded: ${options.length}\n` +
    `Calls found: ${calls.length}\n` +
    `Spreads generated: ${spreads.length}\n` +
    `After filtering: ${filtered.length}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Finds the most recently modified CallSpreads sheet.
 * @param {Spreadsheet} ss - The active spreadsheet
 * @returns {Sheet|null} The CallSpreads sheet or null
 */
function findCallSpreadsSheet_(ss) {
  const sheets = ss.getSheets();
  // Look for sheets ending in "CallSpreads"
  const callSpreadsSheets = sheets.filter(s => s.getName().endsWith("CallSpreads"));
  if (callSpreadsSheets.length === 0) return null;
  // Return first one found (could enhance to track most recent)
  return callSpreadsSheets[0];
}

/**
 * Opens a large dashboard window with Delta vs ROI and Strike vs ROI.
 * If no spreads data exists, prompts user to run Call Spread Finder first.
 */
function showSpreadFinderGraphs() {
  SpreadsheetApp.flush();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = findCallSpreadsSheet_(ss);

  if (!sheet || sheet.getLastRow() < 3) {
    SpreadsheetApp.getUi().alert(
      "No Spread Data",
      "No call spread data found. Please run 'Call Spread Finder' first.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Creates the SpreadFinderGraphs modal dialog
  const html = HtmlService.createHtmlOutputFromFile('ui/SpreadFinderGraphs')
      .setWidth(1050)
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
  const sheet = findCallSpreadsSheet_(ss);

  if (!sheet) return [];

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
