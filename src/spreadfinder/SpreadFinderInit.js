/**
 * SpreadFinderInit.js
 * Initialization, loading, and output functions for SpreadFinder.
 */

const SPREAD_FINDER_CONFIG_SHEET = "SpreadFinderConfig";
const SPREADS_SHEET = "Spreads";
const OPTION_PRICES_SHEET = "OptionPricesUploaded";
const CONFIG_COL = 1; // Column A
const CONFIG_START_ROW = 1;

/** Test function to verify file loads */
function testSpreadFinderLoaded() {
  return "SpreadFinder.js loaded OK";
}

/**
 * Gets distinct symbols and expirations from OptionPricesUploaded for the selection dialog.
 * @returns {Object} { symbols: string[], expirations: {value: string, label: string}[] }
 */
function getSpreadFinderOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(OPTION_PRICES_SHEET);
  if (!sheet) {
    throw new Error("No option prices loaded. Run 'Upload Option Prices & Refresh' first.");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("No option data found in " + OPTION_PRICES_SHEET);
  }

  // Read only header row first to find column indices
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => h.toString().trim().toLowerCase());
  const symIdx = headers.indexOf("symbol");
  const expIdx = headers.indexOf("expiration");

  if (symIdx < 0 || expIdx < 0) {
    throw new Error("Required columns 'symbol' and 'expiration' not found");
  }

  // Read only symbol and expiration columns (much faster than reading entire sheet)
  const symCol = sheet.getRange(2, symIdx + 1, lastRow - 1, 1).getValues();
  const expCol = sheet.getRange(2, expIdx + 1, lastRow - 1, 1).getValues();

  const symbols = new Set();
  const expirations = new Map(); // key: normalized date string, value: Date

  for (let i = 0; i < symCol.length; i++) {
    const sym = (symCol[i][0] || "").toString().trim().toUpperCase();
    if (sym) symbols.add(sym);

    const exp = expCol[i][0];
    if (exp) {
      const expDate = parseDateAtMidnight_(exp);
      if (expDate) {
        // Normalize to YYYY-MM-DD key
        const key = expDate.getFullYear() + "-" +
          String(expDate.getMonth() + 1).padStart(2, "0") + "-" +
          String(expDate.getDate()).padStart(2, "0");
        expirations.set(key, expDate);
      }
    }
  }

  // Sort symbols alphabetically
  const sortedSymbols = Array.from(symbols).sort();

  // Sort expirations by date and format labels
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const sortedExpirations = Array.from(expirations.entries())
    .sort((a, b) => a[1] - b[1])
    .map(([key, date]) => ({
      value: key,
      label: months[date.getMonth()] + " " + date.getDate() + ", " + date.getFullYear()
    }));

  return { symbols: sortedSymbols, expirations: sortedExpirations };
}

/**
 * Ensures SpreadFinderConfig sheet exists with config table.
 * Creates sheet and config if needed, returns sheet.
 */
function ensureSpreadFinderConfigSheet_(ss) {
  let sheet = ss.getSheetByName(SPREAD_FINDER_CONFIG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SPREAD_FINDER_CONFIG_SHEET);
  }

  // Always recreate config to ensure latest settings
  const configData = [
    ["Setting", "Value", "Description"],
    ["symbol", "TSLA", "Comma-separated symbols to analyze (blank=all)"],
    ["minSpreadWidth", 20, "Minimum spread width in dollars"],
    ["maxSpreadWidth", 150, "Maximum spread width in dollars"],
    ["minOpenInterest", 10, "Minimum open interest for both legs"],
    ["minVolume", 5, "Minimum volume for both legs"],
    ["patience", 60, "Minutes for price calculation (0=aggressive, 60=patient)"],
    ["maxDebit", 50, "Maximum debit per share"],
    ["minROI", 2.0, "Minimum ROI (0.5 = 50% return)"],
    ["minStrike", 300, "Minimum lower strike price"],
    ["maxStrike", 700, "Maximum upper strike price"],
    ["minExpirationMonths", 6, "Minimum months until expiration"],
    ["maxExpirationMonths", 36, "Maximum months until expiration"],
    ["", "", ""],
    ["Outlook", "", "Price outlook for boosting fitness"],
    ["outlookFuturePrice", "500", "Target future price (e.g. 700)"],
    ["outlookDate", "3/1/2027", "Target date (e.g. 3/1/2027)"],
    ["outlookConfidence", "0.6", "Confidence 0-1 (e.g. 0.7 = 70%)"]
  ];
  // Read existing values to preserve user edits
  const existingValues = {};
  try {
    const existing = sheet.getRange(CONFIG_START_ROW + 1, CONFIG_COL, configData.length - 1, 2).getValues();
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
 * Ensures the Spreads results sheet exists.
 */
function ensureSpreadsSheet_(ss, name) {
  name = name || SPREADS_SHEET;
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * Loads config from SpreadFinderConfig sheet config table.
 * Returns object with settings and defaults.
 */
function loadSpreadFinderConfig_(sheet) {
  const defaults = {
    minSpreadWidth: 20,
    maxSpreadWidth: 150,
    minOpenInterest: 10,
    minVolume: 0,
    patience: 60,
    maxDebit: 50,
    minROI: 0.5,
    minStrike: 300,
    maxStrike: 700,
    minExpirationMonths: 6,
    maxExpirationMonths: 36
  };

  const config = { ...defaults };

  // Read config rows (rows 2+, columns A-B)
  const lastRow = Math.max(CONFIG_START_ROW + 1, sheet.getLastRow());
  const numRows = lastRow - CONFIG_START_ROW;
  const data = sheet.getRange(CONFIG_START_ROW + 1, CONFIG_COL, numRows, 2).getValues();
  for (const row of data) {
    const setting = (row[0] || "").toString().trim();
    const value = row[1];
    if (setting && value !== "" && value != null && setting in defaults) {
      config[setting] = +value;
    }
  }

  // Read outlook settings (not in defaults, handled separately)
  const outlookKeys = ["outlookFuturePrice", "outlookDate", "outlookConfidence"];
  for (const row of data) {
    const setting = (row[0] || "").toString().trim();
    const value = row[1];
    if (setting && value !== "" && value != null && outlookKeys.includes(setting)) {
      if (setting === "outlookDate") {
        config[setting] = parseDateAtMidnight_(value);
      } else {
        config[setting] = +value;
      }
    }
  }
  // Default outlook to disabled
  if (!config.outlookFuturePrice || !config.outlookConfidence) {
    config.outlookFuturePrice = 0;
    config.outlookConfidence = 0;
  }

  // Read symbol filter (string, not in numeric defaults)
  for (const row of data) {
    const setting = (row[0] || "").toString().trim();
    const value = (row[1] || "").toString().trim();
    if (setting === "symbol" && value) {
      config.symbols = value.split(",").map(s => s.trim().toUpperCase()).filter(s => s);
    }
  }

  // Calculate min/max expiration dates
  const now = new Date();
  config.minExpirationDate = new Date(now.getFullYear(), now.getMonth() + config.minExpirationMonths, now.getDate());
  config.maxExpirationDate = new Date(now.getFullYear(), now.getMonth() + config.maxExpirationMonths, now.getDate());

  return config;
}

/**
 * Loads options from OptionPricesUploaded, optionally filtered by symbols and expirations.
 * @param {Spreadsheet} ss - The active spreadsheet
 * @param {string[]} [filterSymbols] - Optional array of symbols to include
 * @param {Set<string>} [filterExpirations] - Optional set of expiration dates (YYYY-MM-DD) to include
 * @returns {Array} Array of option objects
 */
function loadOptionData_(ss, filterSymbols, filterExpirations) {
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

  // Convert filterSymbols to a Set for O(1) lookup
  const symbolSet = filterSymbols ? new Set(filterSymbols.map(s => s.toUpperCase())) : null;

  const options = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const symbol = (row[idx.symbol] || "").toString().trim().toUpperCase();
    if (!symbol) continue;

    // Filter by symbol if specified
    if (symbolSet && !symbolSet.has(symbol)) continue;

    const rawExp = row[idx.expiration];
    const expDate = parseDateAtMidnight_(rawExp);
    if (!expDate) continue;

    // Normalized key for filtering (YYYY-MM-DD)
    const expNormalized = expDate.getFullYear() + "-" +
      String(expDate.getMonth() + 1).padStart(2, "0") + "-" +
      String(expDate.getDate()).padStart(2, "0");
    // Display format (M/D/YYYY)
    const exp = formatDateMDYYYY_(expDate);

    // Filter by expiration if specified
    if (filterExpirations && expNormalized && !filterExpirations.has(expNormalized)) continue;

    const strike = +row[idx.strike];
    if (!Number.isFinite(strike)) continue;

    const type = parseOptionType_(row[idx.type]);
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
 * Loads held short positions from the Positions sheet BullCallSpreads table.
 * Returns a Set of "SYMBOL|STRIKE|EXPIRATION" keys for fast lookup.
 */
function loadHeldPositions_(ss) {
  const held = new Set();

  // Try Legs table first
  const legsRange = getNamedRangeWithTableFallback_(ss, "Portfolio");
  if (legsRange) {
    const rows = legsRange.getValues();
    log.debug("spreadFinder", "Found Portfolio table with rows: " + rows.length);
    if (rows.length >= 2) {
      const headers = rows[0];
      const idxSym = findColumn_(headers, ["symbol", "ticker"]);
      const idxStrike = findColumn_(headers, ["strike", "strikeprice"]);
      const idxQty = findColumn_(headers, ["qty", "quantity", "contracts", "contract", "count", "shares"]);
      const idxExp = findColumn_(headers, ["expiration", "exp", "expiry", "expirationdate", "expdate"]);

      if (idxStrike >= 0 && idxQty >= 0) {
        let lastSym = "";
        let lastExp = "";
        for (let r = 1; r < rows.length; r++) {
          const rawSym = idxSym >= 0 ? String(rows[r][idxSym] ?? "").trim().toUpperCase() : "";
          if (rawSym) lastSym = rawSym;

          if (idxExp >= 0) {
            const parsed = parseDateAtMidnight_(rows[r][idxExp]);
            if (parsed) {
              lastExp = formatDateMDYYYY_(parsed);
            }
          }

          const qty = parseNumber_(rows[r][idxQty]);
          const strike = parseNumber_(rows[r][idxStrike]);

          // Short legs have negative qty
          if (lastSym && Number.isFinite(strike) && strike > 0 && Number.isFinite(qty) && qty < 0) {
            held.add(`${lastSym}|${strike}|${lastExp}`);
          }
        }
        return held;
      }
    }
  }

  // Fall back to old Positions sheet logic
  const sheet = ss.getSheetByName("Positions");
  if (!sheet) {
    log.debug("spreadFinder", "Positions sheet not found, skipping held position check");
    return held;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return held;

  // Find the header row containing "Short Strike", then read all columns from that row
  let headerRow = -1;
  let symbolCol = -1, shortStrikeCol = -1, expirationCol = -1, contractsCol = -1;

  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const val = (data[r][c] || "").toString().toLowerCase().trim();
      if (val === "short strike") { headerRow = r; break; }
    }
    if (headerRow >= 0) break;
  }

  if (headerRow >= 0) {
    for (let c = 0; c < data[headerRow].length; c++) {
      const val = (data[headerRow][c] || "").toString().toLowerCase().trim();
      if (val === "symbol" && symbolCol < 0) symbolCol = c;
      if (val === "short strike" && shortStrikeCol < 0) shortStrikeCol = c;
      if (val === "expiration" && expirationCol < 0) expirationCol = c;
      if (val === "contracts" && contractsCol < 0) contractsCol = c;
    }
  }

  if (headerRow < 0 || shortStrikeCol < 0) {
    log.debug("spreadFinder", "BullCallSpreads table not found on Positions sheet");
    return held;
  }

  let lastSymbol = "";
  let lastExp = "";
  for (let r = headerRow + 1; r < data.length; r++) {
    const rowSymbol = symbolCol >= 0 ? (data[r][symbolCol] || "").toString().trim().toUpperCase() : "";
    const strike = +data[r][shortStrikeCol];
    const contracts = contractsCol >= 0 ? +data[r][contractsCol] : 1;

    // Carry forward symbol and expiration from group header rows
    if (rowSymbol) lastSymbol = rowSymbol;

    if (expirationCol >= 0) {
      const parsed = parseDateAtMidnight_(data[r][expirationCol]);
      if (parsed) {
        lastExp = formatDateMDYYYY_(parsed);
      }
    }

    const symbol = lastSymbol;
    const exp = lastExp;

    if (symbol && Number.isFinite(strike) && strike > 0 && contracts > 0) {
      held.add(`${symbol}|${strike}|${exp}`);
    }
  }

  return held;
}

/**
 * Outputs spread results to SpreadFinder sheet below config.
 * Uses formulas for MaxProfit, ROI, Fitness so user can edit Debit.
 */
function outputSpreadResults_(sheet, spreads, config) {
  const RESULTS_START_ROW = 2; // Row 1 = timestamp, Row 2 = headers

  // Clear entire sheet
  const lastRow = Math.max(sheet.getLastRow(), RESULTS_START_ROW);
  if (lastRow >= 1) {
    const clearRange = sheet.getRange(1, 1, lastRow, Math.max(sheet.getLastColumn(), 1));
    // Remove any existing banding
    const bandings = clearRange.getBandings();
    bandings.forEach(b => b.remove());
    // Clear content, formatting, borders
    clearRange.clear();
  }

  // Remove existing filter if any
  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();

  // Timestamp
  sheet.getRange(1, 1).setValue("Results - " + new Date().toLocaleString());

  // Headers (A-T)
  const headers = [
    "Symbol", "Expiration", "Lower", "Upper", "Width",
    "Debit", "MaxProfit", "ROI", "ExpGain", "ExpROI",
    "LowerDelta", "UpperDelta",
    "LowerOI", "UpperOI", "Liquidity", "Tightness", "Fitness", "OptionStrat", "Label", "Held", "IV"
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
    "Expected dollar gain using prob-of-touch model (80% target)",
    "Expected ROI = ExpGain / Debit",
    "Delta of lower call (0-1). Higher = more ITM, higher prob of profit",
    "Delta of upper call. Lower than LowerDelta since further OTM",
    "Open Interest on lower strike. Higher = better liquidity",
    "Open Interest on upper strike. Want both legs liquid",
    "Liquidity score = sqrt(LowerOI × UpperOI) / 100",
    "Bid-ask tightness. Higher = tighter spreads, better fills",
    "Fitness = ExpROI × Liquidity^0.1 × Tightness^0.1",
    "Link to OptionStrat visualization",
    "Label for chart identification",
    "HELD = you already have a conflicting short position",
    "Implied volatility of the lower (long) leg"
  ];
  const hdrRange = sheet.getRange(RESULTS_START_ROW, 1, 1, headers.length);
  hdrRange.setValues([headers]).setFontWeight("bold");
  hdrRange.setNotes([headerNotes]);

  if (spreads.length === 0) return;

  // Build column index: col("Name") returns 1-based sheet column number
  const colIdx = {};
  headers.forEach((h, i) => colIdx[h] = i);
  const col = name => colIdx[name] + 1;   // 1-based for sheet APIs
  const colLetter = name => String.fromCharCode(65 + colIdx[name]); // A-Z

  // Per-column number formats and widths, keyed by header name
  const formatMap = {
    Symbol: "@", Expiration: "@", Lower: "#,##0", Upper: "#,##0", Width: "#,##0",
    Debit: "$#,##0.00", MaxProfit: "$#,##0.00", ROI: "0.00",
    ExpGain: "$#,##0.00", ExpROI: "0.00",
    LowerDelta: "0.00", UpperDelta: "0.00", LowerOI: "#,##0", UpperOI: "#,##0",
    Liquidity: "0.00", Tightness: "0.00", Fitness: "0.00",
    OptionStrat: "@", Label: "@", Held: "@", IV: "0.00%"
  };
  const widthMap = {
    Symbol: 60, Expiration: 90, Lower: 60, Upper: 60, Width: 50,
    Debit: 70, MaxProfit: 70, ROI: 50, ExpGain: 70, ExpROI: 55,
    LowerDelta: 55, UpperDelta: 55, LowerOI: 55, UpperOI: 55,
    Liquidity: 55, Tightness: 55, Fitness: 55, OptionStrat: 100, Label: 150, Held: 50, IV: 55
  };

  const dataStartRow = RESULTS_START_ROW + 1;
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  // Build a row-value map per spread, placed into correct column positions
  const allRows = [];
  for (let i = 0; i < spreads.length; i++) {
    const s = spreads[i];
    // Parse expiration carefully to avoid timezone issues with ISO strings
    const expDate = parseDateAtMidnight_(s.expiration);
    if (!expDate) continue; // Skip spreads with invalid expiration
    const dateStr = months[expDate.getMonth()] + " " + String(expDate.getFullYear()).slice(2);
    const label = `${s.symbol} ${s.lowerStrike}/${s.upperStrike} ${dateStr}`;

    const rowData = new Array(headers.length).fill("");
    rowData[colIdx.Symbol] = s.symbol;
    rowData[colIdx.Expiration] = s.expiration;
    rowData[colIdx.Lower] = s.lowerStrike;
    rowData[colIdx.Upper] = s.upperStrike;
    rowData[colIdx.Width] = s.width;
    rowData[colIdx.Debit] = s.debit;
    rowData[colIdx.MaxProfit] = s.maxProfit;
    rowData[colIdx.ROI] = s.roi;
    rowData[colIdx.ExpGain] = s.expectedGain;
    rowData[colIdx.ExpROI] = s.expectedROI;
    rowData[colIdx.LowerDelta] = s.lowerDelta;
    rowData[colIdx.UpperDelta] = s.upperDelta;
    rowData[colIdx.LowerOI] = s.lowerOI;
    rowData[colIdx.UpperOI] = s.upperOI;
    rowData[colIdx.Liquidity] = s.liquidityScore;
    rowData[colIdx.Tightness] = s.tightness;
    rowData[colIdx.Fitness] = s.fitness;
    // Compute OptionStrat URL directly (avoids popup warning from custom function in HYPERLINK)
    const osUrl = buildOptionStratUrl(
      `${s.lowerStrike}/${s.upperStrike}`,
      s.symbol,
      "bull-call-spread",
      s.expiration
    );
    rowData[colIdx.OptionStrat] = osUrl;
    rowData[colIdx.Label] = label;
    rowData[colIdx.Held] = s.held ? "HELD" : "";
    rowData[colIdx.IV] = s.lowerIV;
    allRows.push(rowData);
  }
  sheet.getRange(dataStartRow, 1, allRows.length, headers.length).setValues(allRows);

  // Number formats from the map
  const formats = spreads.map(() => headers.map(h => formatMap[h] || "@"));
  sheet.getRange(dataStartRow, 1, spreads.length, headers.length).setNumberFormats(formats);

  // Style: header, banding, borders, filter in minimal API calls
  const tableRange = sheet.getRange(RESULTS_START_ROW, 1, spreads.length + 1, headers.length);
  tableRange.createFilter();
  tableRange.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(RESULTS_START_ROW, 1, 1, headers.length)
    .setBackground("#4285f4").setFontColor("white").setFontWeight("bold");

  const dataRange = sheet.getRange(dataStartRow, 1, spreads.length, headers.length);
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  // Column widths from the map
  headers.forEach((h, i) => sheet.setColumnWidth(i + 1, widthMap[h] || 55));

  // Clip OptionStrat column
  sheet.getRange(RESULTS_START_ROW, col("OptionStrat"), spreads.length + 1, 1)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  showSpreadFinderGraphs();
}
