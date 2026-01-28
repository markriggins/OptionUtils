/**
 * SpreadFinder.js
 * Analyzes OptionPricesUploaded to find and rank bull call spread opportunities.
 *
 * Uses a config table at top of SpreadFinder sheet for tunable parameters.
 * Results appear below the config on the same sheet.
 *
 * Config table format (auto-created):
 *   | Setting            | Value | Description                                    |
 *   | maxSpreadWidth     | 150   | Maximum spread width in dollars                |
 *   | minOpenInterest    | 10    | Minimum open interest for both legs            |
 *   | minVolume          | 0     | Minimum volume for both legs                   |
 *   | patience           | 60    | Minutes for price calculation (0=aggressive)   |
 *   | maxDebit           | 50    | Maximum debit per share                        |
 *   | minROI             | 0.5   | Minimum ROI (0.5 = 50%)                        |
 *
 * Version: 1.0
 */

const SPREAD_FINDER_SHEET = "SpreadFinder";
const OPTION_PRICES_SHEET = "OptionPricesUploaded";
const CONFIG_COL = 1; // Column A
const CONFIG_START_ROW = 1;

// Column letters for formulas
const COLS = "ABCDEFGHIJKLMNOP";
// A=Symbol, B=Expiration, C=Lower, D=Upper, E=Width, F=Debit, G=MaxProfit, H=ROI,
// I=LowerDelta, J=UpperDelta, K=LowerOI, L=UpperOI, M=Liquidity, N=Tightness, O=Fitness, P=OptionStrat

/**
 * Runs the spread finder analysis. Call from menu or script.
 * Ensures config exists, reads it, scans options, ranks spreads, outputs results.
 */
function runSpreadFinder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ensure sheet and config exist
  const sheet = ensureSpreadFinderSheet_(ss);

  // Load config from sheet
  const config = loadSpreadFinderConfig_(sheet);
  Logger.log("SpreadFinder config: " + JSON.stringify(config));

  // Load option data
  const options = loadOptionData_(ss);
  Logger.log("Loaded " + options.length + " options");

  // Filter to calls only
  const calls = options.filter(o => o.type === "Call");
  Logger.log("Filtered to " + calls.length + " calls");

  // Group by symbol+expiration
  const grouped = groupBySymbolExpiration_(calls);

  // Generate and score all spreads
  const spreads = [];
  for (const key of Object.keys(grouped)) {
    const chain = grouped[key];
    const chainSpreads = generateSpreads_(chain, config);
    spreads.push(...chainSpreads);
  }
  Logger.log("Generated " + spreads.length + " spreads");

  // Load existing short strikes from BullCallSpreads to avoid conflicts (by symbol)
  const heldShortStrikesBySymbol = loadHeldShortStrikes_(ss);
  for (const [sym, strikes] of heldShortStrikesBySymbol) {
    Logger.log(`Held short strikes for ${sym}: ${JSON.stringify([...strikes])}`);
  }

  // Filter by config constraints
  const minExpDate = config.minExpirationDate;
  const filtered = spreads.filter(s => {
    // Parse expiration date
    const expDate = new Date(s.expiration);
    // Can't use a strike as long if we're already short it (for same symbol)
    const symbolShorts = heldShortStrikesBySymbol.get(s.symbol);
    const hasConflict = symbolShorts ? symbolShorts.has(s.lowerStrike) : false;
    return s.debit > 0 &&
      s.debit <= config.maxDebit &&
      s.roi >= config.minROI &&
      s.lowerOI >= config.minOpenInterest &&
      s.upperOI >= config.minOpenInterest &&
      s.lowerVol >= config.minVolume &&
      s.upperVol >= config.minVolume &&
      s.lowerStrike <= config.maxLowerStrike &&
      expDate >= minExpDate &&
      !hasConflict;
  });
  Logger.log("Filtered to " + filtered.length + " spreads meeting criteria");

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
 * Ensures SpreadFinder sheet exists with config table.
 * Creates sheet and config if needed, returns sheet.
 */
function ensureSpreadFinderSheet_(ss) {
  let sheet = ss.getSheetByName(SPREAD_FINDER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SPREAD_FINDER_SHEET);
  }

  // Always recreate config to ensure latest settings
  const configData = [
    ["Setting", "Value", "Description"],
    ["maxSpreadWidth", 150, "Maximum spread width in dollars"],
    ["minOpenInterest", 10, "Minimum open interest for both legs"],
    ["minVolume", 0, "Minimum volume for both legs"],
    ["patience", 60, "Minutes for price calculation (0=aggressive, 60=patient)"],
    ["maxDebit", 50, "Maximum debit per share"],
    ["minROI", 0.5, "Minimum ROI (0.5 = 50% return)"],
    ["maxLowerStrike", 9999, "Maximum lower strike price"],
    ["minExpirationMonths", 6, "Minimum months until expiration"],
    ["startTableOnRow", 20, "Row number where results table starts"]
  ];
  // Read existing values to preserve user edits
  const existingValues = {};
  try {
    const existing = sheet.getRange(CONFIG_START_ROW + 1, CONFIG_COL, 10, 2).getValues();
    for (const row of existing) {
      if (row[0] && row[1] !== "") existingValues[row[0]] = row[1];
    }
  } catch (e) { /* ignore */ }
  // Merge existing values
  for (let i = 1; i < configData.length; i++) {
    const key = configData[i][0];
    if (key in existingValues) configData[i][1] = existingValues[key];
  }
  const configRange = sheet.getRange(CONFIG_START_ROW, CONFIG_COL, configData.length, 3);
  // Remove existing banding before applying new
  const existingBanding = configRange.getBandings();
  existingBanding.forEach(b => b.remove());

  configRange.setValues(configData);
  // Style config as table
  sheet.getRange(CONFIG_START_ROW, CONFIG_COL, 1, 3).setBackground("#34a853").setFontColor("white").setFontWeight("bold");
  const configDataRange = sheet.getRange(CONFIG_START_ROW + 1, CONFIG_COL, configData.length - 1, 3);
  configDataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREEN, false, false);
  configRange.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  sheet.autoResizeColumn(CONFIG_COL);
  sheet.autoResizeColumn(CONFIG_COL + 1);
  sheet.autoResizeColumn(CONFIG_COL + 2);

  return sheet;
}

/**
 * Loads config from SpreadFinder sheet config table.
 * Returns object with settings and defaults.
 */
function loadSpreadFinderConfig_(sheet) {
  const defaults = {
    maxSpreadWidth: 150,
    minOpenInterest: 10,
    minVolume: 0,
    patience: 60,
    maxDebit: 50,
    minROI: 0.5,
    maxLowerStrike: 9999,
    minExpirationMonths: 6,
    startTableOnRow: 20
  };

  const config = { ...defaults };

  // Read config rows (rows 2-11, columns A-B)
  const data = sheet.getRange(CONFIG_START_ROW + 1, CONFIG_COL, 10, 2).getValues();
  for (const row of data) {
    const setting = (row[0] || "").toString().trim();
    const value = row[1];
    if (setting && value !== "" && value != null && setting in defaults) {
      config[setting] = +value;
    }
  }

  // Calculate min expiration date
  const now = new Date();
  config.minExpirationDate = new Date(now.getFullYear(), now.getMonth() + config.minExpirationMonths, now.getDate());

  return config;
}

/**
 * Loads all options from OptionPricesUploaded.
 * Returns array of option objects.
 */
function loadOptionData_(ss) {
  const sheet = ss.getSheetByName(OPTION_PRICES_SHEET);
  if (!sheet) throw new Error(`Sheet '${OPTION_PRICES_SHEET}' not found`);

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  // Build header index
  const headers = data[0].map(h => h.toString().trim().toLowerCase());
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const required = ["symbol", "expiration", "strike", "type", "bid", "ask"];
  for (const r of required) {
    if (!(r in idx)) throw new Error(`Required column '${r}' not found`);
  }

  const options = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const symbol = (row[idx.symbol] || "").toString().trim().toUpperCase();
    if (!symbol) continue;

    let exp = row[idx.expiration];
    if (exp instanceof Date) {
      exp = Utilities.formatDate(exp, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      exp = (exp || "").toString().trim();
    }
    if (!exp) continue;

    const strike = +row[idx.strike];
    if (!Number.isFinite(strike)) continue;

    const type = normalizeType_(row[idx.type]);
    if (!type) continue;

    const bid = +row[idx.bid] || 0;
    const ask = +row[idx.ask] || 0;
    const mid = idx.mid !== undefined ? (+row[idx.mid] || 0) : (bid + ask) / 2;
    const iv = idx.iv !== undefined ? (+row[idx.iv] || 0) : 0;
    const delta = idx.delta !== undefined ? (+row[idx.delta] || 0) : 0;
    const volume = idx.volume !== undefined ? (+row[idx.volume] || 0) : 0;
    const openint = idx.openint !== undefined ? (+row[idx.openint] || 0) : 0;
    const moneyness = idx.moneyness !== undefined ? (row[idx.moneyness] || "") : "";

    options.push({
      symbol, expiration: exp, strike, type,
      bid, mid, ask, iv, delta, volume, openint, moneyness
    });
  }

  return options;
}

/**
 * Groups options by symbol+expiration key.
 * Returns { "TSLA|2028-06-16": [options...], ... }
 */
function groupBySymbolExpiration_(options) {
  const groups = {};
  for (const o of options) {
    const key = `${o.symbol}|${o.expiration}`;
    if (!groups[key]) groups[key] = [];
    groups[key].push(o);
  }
  // Sort each group by strike
  for (const key of Object.keys(groups)) {
    groups[key].sort((a, b) => a.strike - b.strike);
  }
  return groups;
}

/**
 * Generates all valid spreads from a sorted chain of calls.
 * Returns array of spread objects with metrics.
 */
function generateSpreads_(chain, config) {
  const spreads = [];
  const n = chain.length;

  for (let i = 0; i < n; i++) {
    const lower = chain[i];
    for (let j = i + 1; j < n; j++) {
      const upper = chain[j];
      const width = upper.strike - lower.strike;

      // Skip if too wide
      if (width > config.maxSpreadWidth) continue;

      // Skip if no valid bid/ask
      if (lower.ask <= 0 || upper.bid < 0) continue;

      // Calculate debit using mid pricing (validated by executed trades with GT 60 patience)
      // For patient orders, mid is achievable; below mid may not fill
      const lowerMid = (lower.bid + lower.ask) / 2;
      const upperMid = (upper.bid + upper.ask) / 2;

      let debit = lowerMid - upperMid;
      if (debit < 0) debit = 0;
      debit = round2_(debit);

      // Calculate metrics
      const maxProfit = width - debit;
      const maxLoss = debit;
      const roi = debit > 0 ? maxProfit / debit : 0;

      // Liquidity score: geometric mean of OI, scaled
      const minOI = Math.min(lower.openint, upper.openint);
      const liquidityScore = Math.sqrt(lower.openint * upper.openint) / 100;

      // Bid-ask tightness (lower is better, so invert)
      const lowerSpread = lower.ask - lower.bid;
      const upperSpread = upper.ask - upper.bid;
      const avgBidAskSpread = (lowerSpread + upperSpread) / 2;
      const tightness = avgBidAskSpread > 0 ? 1 / avgBidAskSpread : 10;

      // Delta-based probability estimate (rough)
      // Lower delta = more OTM = lower prob but higher reward
      const probITM = Math.abs(lower.delta); // P(stock > lower strike)

      // Time factor: longer expirations are better (more time to reach target)
      // Use sqrt for diminishing returns - going from 1yr to 2yr helps more than 2yr to 3yr
      const expDate = new Date(lower.expiration);
      const now = new Date();
      const daysToExp = Math.max(1, (expDate - now) / (1000 * 60 * 60 * 24));
      const yearsToExp = daysToExp / 365;
      const timeFactor = Math.sqrt(yearsToExp);

      // Fitness = ROI * liquidity * tightness * time (tunable)
      const fitness = round2_(roi * Math.sqrt(liquidityScore) * Math.sqrt(tightness) * timeFactor);

      spreads.push({
        symbol: lower.symbol,
        expiration: lower.expiration,
        lowerStrike: lower.strike,
        upperStrike: upper.strike,
        width,
        debit,
        maxProfit: round2_(maxProfit),
        maxLoss: round2_(maxLoss),
        roi: round2_(roi),
        lowerDelta: round2_(lower.delta),
        upperDelta: round2_(upper.delta),
        lowerOI: lower.openint,
        upperOI: upper.openint,
        lowerVol: lower.volume,
        upperVol: upper.volume,
        liquidityScore: round2_(liquidityScore),
        tightness: round2_(tightness),
        fitness: round2_(fitness)
      });
    }
  }

  return spreads;
}

/**
 * Loads existing short strikes from BullCallSpreads table on Positions sheet.
 * Returns a Map of symbol -> Set of short strike numbers to avoid conflicts.
 */
function loadHeldShortStrikes_(ss) {
  const shortStrikesBySymbol = new Map();

  const sheet = ss.getSheetByName("Positions");
  if (!sheet) {
    Logger.log("Positions sheet not found, skipping conflict check");
    return shortStrikesBySymbol;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return shortStrikesBySymbol;

  // Find header row with "Symbol", "Short Strike", "Contracts"
  let headerRow = -1;
  let symbolCol = -1;
  let shortStrikeCol = -1;
  let contractsCol = -1;

  for (let r = 0; r < Math.min(10, data.length); r++) {
    for (let c = 0; c < data[r].length; c++) {
      const val = (data[r][c] || "").toString().toLowerCase().trim();
      if (val === "symbol") symbolCol = c;
      if (val === "short strike") {
        headerRow = r;
        shortStrikeCol = c;
      }
      if (val === "contracts") contractsCol = c;
    }
    if (headerRow >= 0) break;
  }

  if (headerRow < 0 || shortStrikeCol < 0) {
    Logger.log("BullCallSpreads table not found on Positions sheet");
    return shortStrikesBySymbol;
  }

  // Read short strikes from rows below header (where contracts > 0)
  for (let r = headerRow + 1; r < data.length; r++) {
    const symbol = symbolCol >= 0 ? (data[r][symbolCol] || "").toString().trim().toUpperCase() : "";
    const shortStrike = +data[r][shortStrikeCol];
    const contracts = contractsCol >= 0 ? +data[r][contractsCol] : 1;

    // Only count if we have contracts
    if (symbol && Number.isFinite(shortStrike) && shortStrike > 0 && contracts > 0) {
      if (!shortStrikesBySymbol.has(symbol)) {
        shortStrikesBySymbol.set(symbol, new Set());
      }
      shortStrikesBySymbol.get(symbol).add(shortStrike);
    }
  }

  return shortStrikesBySymbol;
}

/**
 * Outputs spread results to SpreadFinder sheet below config.
 * Uses formulas for MaxProfit, ROI, Fitness so user can edit Debit.
 */
function outputSpreadResults_(sheet, spreads, config) {
  const RESULTS_START_ROW = config.startTableOnRow;

  // Clear results area (from RESULTS_START_ROW - 1 down for timestamp)
  const lastRow = Math.max(sheet.getLastRow(), RESULTS_START_ROW);
  if (lastRow >= RESULTS_START_ROW - 1) {
    const clearRange = sheet.getRange(RESULTS_START_ROW - 1, 1, lastRow - RESULTS_START_ROW + 2, 20);
    // Remove any existing banding
    const bandings = clearRange.getBandings();
    bandings.forEach(b => b.remove());
    // Clear content, formatting, borders
    clearRange.clear();
  }

  // Remove existing filter if any
  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();

  // Timestamp above table
  sheet.getRange(RESULTS_START_ROW - 1, 1).setValue("Results - " + new Date().toLocaleString());

  // Headers (A-Q)
  const headers = [
    "Symbol", "Expiration", "Lower", "Upper", "Width",
    "Debit", "MaxProfit", "ROI", "LowerDelta", "UpperDelta",
    "LowerOI", "UpperOI", "Liquidity", "Tightness", "Fitness", "OptionStrat", "Label"
  ];
  const headerNotes = [
    "Stock ticker symbol",
    "Option expiration date",
    "Lower (long) strike - you BUY this call",
    "Upper (short) strike - you SELL this call",
    "Spread width = Upper - Lower (max profit potential)",
    "Net debit to open (editable - formulas recalculate)",
    "Max profit = Width - Debit (if stock > Upper at expiry)",
    "Return on Investment = MaxProfit / Debit",
    "Delta of lower call (0-1). Higher = more ITM, higher prob of profit",
    "Delta of upper call. Lower than LowerDelta since further OTM",
    "Open Interest on lower strike. Higher = better liquidity",
    "Open Interest on upper strike. Want both legs liquid",
    "Liquidity score = sqrt(LowerOI × UpperOI) / 100",
    "Bid-ask tightness. Higher = tighter spreads, better fills",
    "Fitness = ROI × sqrt(Liquidity) × sqrt(Tightness)",
    "Link to OptionStrat visualization",
    "Label for chart identification"
  ];
  const hdrRange = sheet.getRange(RESULTS_START_ROW, 1, 1, headers.length);
  hdrRange.setValues([headers]).setFontWeight("bold");
  hdrRange.setNotes([headerNotes]);

  if (spreads.length === 0) return;

  const dataStartRow = RESULTS_START_ROW + 1;

  // Data columns A-F (values), I-N (values)
  // G, H, O, P, Q will be formulas
  const dataRows = spreads.map(s => [
    s.symbol, s.expiration, s.lowerStrike, s.upperStrike, s.width,
    s.debit, "", "", // G=MaxProfit, H=ROI (formulas)
    s.lowerDelta, s.upperDelta, s.lowerOI, s.upperOI,
    s.liquidityScore, s.tightness, "", "", "" // O=Fitness, P=OptionStrat, Q=Label (formulas)
  ]);
  sheet.getRange(dataStartRow, 1, dataRows.length, headers.length).setValues(dataRows);

  // Build formula arrays for columns G, H, O, P, Q
  const maxProfitFormulas = [];
  const roiFormulas = [];
  const fitnessFormulas = [];
  const optionStratFormulas = [];
  const labelFormulas = [];

  for (let i = 0; i < spreads.length; i++) {
    const row = dataStartRow + i;
    // MaxProfit = Width - Debit (E - F)
    maxProfitFormulas.push([`=E${row}-F${row}`]);
    // ROI = MaxProfit / Debit (G / F)
    roiFormulas.push([`=IF(F${row}>0,G${row}/F${row},0)`]);
    // Fitness = ROI * SQRT(Liquidity) * SQRT(Tightness)
    fitnessFormulas.push([`=ROUND(H${row}*SQRT(M${row})*SQRT(N${row}),2)`]);
    // OptionStrat URL as hyperlink
    const osUrl = `buildOptionStratUrl(C${row}&"/"&D${row},A${row},"bull-call-spread",B${row})`;
    optionStratFormulas.push([`=HYPERLINK(${osUrl},"OptionStrat")`]);

    // Format labels as: "TSLA 350/450 Jan 28"
    const sym = sheet.getRange(row, 1).getValue();
    const lowStrike = sheet.getRange(row, 3).getValue();
    const highStrike = sheet.getRange(row, 4).getValue();
    const expDate = new Date(sheet.getRange(row, 2).getValue());
    const dateStr = Utilities.formatDate(expDate, Session.getScriptTimeZone(), "MMM yy");
    labelValues.push([`${sym} ${lowStrike}/${highStrike} ${dateStr}`]);  }

  // Set formulas
  sheet.getRange(dataStartRow, 7, spreads.length, 1).setFormulas(maxProfitFormulas); // G
  sheet.getRange(dataStartRow, 8, spreads.length, 1).setFormulas(roiFormulas); // H
  sheet.getRange(dataStartRow, 15, spreads.length, 1).setFormulas(fitnessFormulas); // O
  sheet.getRange(dataStartRow, 16, spreads.length, 1).setFormulas(optionStratFormulas); // P
  sheet.getRange(dataStartRow, 17, spreads.length, 1).setValues(labelFormulas); // Q

  // Format
  sheet.getRange(dataStartRow, 6, spreads.length, 1).setNumberFormat("$#,##0.00"); // Debit
  sheet.getRange(dataStartRow, 7, spreads.length, 1).setNumberFormat("$#,##0.00"); // MaxProfit
  sheet.getRange(dataStartRow, 8, spreads.length, 1).setNumberFormat("0.00"); // ROI
  sheet.getRange(dataStartRow, 9, spreads.length, 2).setNumberFormat("0.00"); // Deltas
  sheet.getRange(dataStartRow, 15, spreads.length, 1).setNumberFormat("0.00"); // Fitness

  // Add filter
  const tableRange = sheet.getRange(RESULTS_START_ROW, 1, spreads.length + 1, headers.length);
  tableRange.createFilter();

  // Style as table: header row + banded rows
  const headerRange = sheet.getRange(RESULTS_START_ROW, 1, 1, headers.length);
  headerRange.setBackground("#4285f4").setFontColor("white").setFontWeight("bold");

  // Banded rows
  const dataRange = sheet.getRange(dataStartRow, 1, spreads.length, headers.length);
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  // Borders
  tableRange.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

  // Auto-resize columns (except OptionStrat)
  for (let i = 1; i < headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  // OptionStrat column: clip text, fixed width
  const optionStratCol = sheet.getRange(RESULTS_START_ROW, 16, spreads.length + 1, 1);
  optionStratCol.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.setColumnWidth(16, 100);

  showSpreadFinderGraphs();
}

/**
 * Opens a large dashboard window with Delta vs ROI and Strike vs ROI.
 */
function showSpreadFinderGraphs() {
  SpreadsheetApp.flush();

  // Creates the HTML interface from the SidebarChart.html file
  const html = HtmlService.createHtmlOutputFromFile('SpreadFinderGraphs')
      .setWidth(1050) // Wide enough for side-by-side or large stacked charts
      .setHeight(850);

  SpreadsheetApp.getUi().showModalDialog(html, 'Spread Finder Graphs');
}

/**
 * Fetches and formats data for the Sidebar chart.
 * Orders by Fitness so the best points are drawn last (on top).
 */
/**
 * Enhanced Data Fetch for Graphs
 * Integrates IV and uses the column mapping confirmed from your screenshots.
 */
function getSidebarData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SpreadFinder");
  const lastRow = sheet.getLastRow();
  const startRow = 21;
  if (lastRow < startRow) return [];

  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 17).getValues();
  const today = new Date();
  today.setHours(0,0,0,0);

  return data.map(row => {
    const expDate = new Date(row[1]);
    const diffTime = expDate.getTime() - today.getTime();
    const dte = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

    return {
      // Basic Charting
      delta: parseFloat(row[8]) || 0,     // Col I
      roi: parseFloat(row[7]) || 0,       // Col H
      strike: parseFloat(row[2]) || 0,    // Col C
      fitness: parseFloat(row[14]) || 0,  // Col O
      label: String(row[16] || ""),       // Col Q

      // Full Analytics Suite
      width: row[4],         // Col E
      debit: row[5],         // Col F
      maxProfit: row[6],     // Col G
      lowerDelta: row[8],    // Col I
      upperDelta: row[9],    // Col J
      lowerOI: row[10],      // Col K
      upperOI: row[11],      // Col L
      liquidity: row[12],    // Col M
      tightness: row[13],    // Col N
      iv: row[15],           // Col P (OptionStrat/IV field)
      dte: dte > 0 ? dte : 0
    };
  }).sort((a, b) => a.fitness - b.fitness);
}