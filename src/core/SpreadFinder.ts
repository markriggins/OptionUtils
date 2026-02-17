/**
 * Main entry point for Spread Finder (called from menu).
 * @returns {void}
 */
function runSpreadFinder() {
  const ui = SpreadsheetApp.getUi();
  try {
    // === DEFENSIVE DEFAULTS (Phase 4) ===
    const rawConfig = loadSpreadFinderConfig_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONST.CONFIG_SHEET));
    const config: SpreadFinderConfig = {
      minROI: Number(rawConfig.minROI) || 0.5,
      patience: Number(rawConfig.patience) || 60,
      minLiquidity: Number(rawConfig.minLiquidity) || 10,
      maxSpreadWidth: Number(rawConfig.maxSpreadWidth) || 5,
      minExpectedGain: Number(rawConfig.minExpectedGain) || 0.8,
      ...rawConfig  // keep any other fields you have
    };

    // === YOUR ORIGINAL BODY STARTS HERE (unchanged) ===
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ensure config and results sheets exist
  const configSheet = ensureSpreadFinderConfigSheet_(ss);
  const sheet = ensureSpreadsSheet_(ss);

  // Load config from config sheet
  const config = loadSpreadFinderConfig_(configSheet);
  Logger.log("SpreadFinder config: " + JSON.stringify(config));

  // Load option data
  const options = loadOptionData_(ss);
  Logger.log("Loaded " + options.length + " options");

  // Filter to calls only
  const calls = options.filter(o => o.type === "Call");
  Logger.log("Filtered to " + calls.length + " calls");

  // Group by symbol+expiration
  const grouped = groupBySymbolExpiration_(calls);

  // Derive current price from ATM calls (delta closest to 0.5)
  const currentPrice = estimateCurrentPrice_(calls);
  Logger.log("Estimated current price: " + currentPrice);

  // Default outlook if not set by user
  if (!config.outlookFuturePrice) {
    config.outlookFuturePrice = roundTo_(currentPrice * 1.25, 2);
    Logger.log("Defaulting outlookFuturePrice to " + config.outlookFuturePrice);
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
  Logger.log("Generated " + spreads.length + " spreads");

  // Load held positions from Positions sheet
  const conflicts = loadHeldPositions_(ss);
  Logger.log("Loaded " + conflicts.size + " held positions: " + JSON.stringify([...conflicts]));

  // Filter by config constraints, mark conflicts instead of removing
  const minExpDate = config.minExpirationDate;
  const maxExpDate = config.maxExpirationDate;
  const filtered = spreads.filter(s => {
    const expDate = new Date(s.expiration);
    // Mark conflicts as held (but keep them in results)
    s.held = conflicts.has(`${s.symbol}|${s.lowerStrike}|${s.expiration}`);
    return s.debit > 0 &&
      s.debit <= config.maxDebit &&
      s.roi >= config.minROI &&
      s.lowerOI >= config.minOpenInterest &&
      s.upperOI >= config.minOpenInterest &&
      s.lowerVol >= config.minVolume &&
      s.upperVol >= config.minVolume &&
      s.lowerStrike >= config.minStrike &&
      s.upperStrike <= config.maxStrike &&
      expDate >= minExpDate &&
      expDate <= maxExpDate;
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
    ["minSpreadWidth", 20, "Minimum spread width in dollars"],
    ["maxSpreadWidth", 150, "Maximum spread width in dollars"],
    ["minOpenInterest", 10, "Minimum open interest for both legs"],
    ["minVolume", 0, "Minimum volume for both legs"],
    ["patience", 60, "Minutes for price calculation (0=aggressive, 60=patient)"],
    ["maxDebit", 50, "Maximum debit per share"],
    ["minROI", 0.5, "Minimum ROI (0.5 = 50% return)"],
    ["minStrike", 300, "Minimum lower strike price"],
    ["maxStrike", 700, "Maximum upper strike price"],
    ["minExpirationMonths", 6, "Minimum months until expiration"],
    ["maxExpirationMonths", 36, "Maximum months until expiration"],
    ["", "", ""],
    ["Outlook", "", "Price outlook for boosting fitness"],
    ["outlookFuturePrice", "", "Target future price (e.g. 700)"],
    ["outlookDate", "", "Target date (e.g. 2027-01-01)"],
    ["outlookConfidence", "", "Confidence 0-1 (e.g. 0.7 = 70%)"]
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
function ensureSpreadsSheet_(ss) {
  let sheet = ss.getSheetByName(SPREADS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SPREADS_SHEET);
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
        config[setting] = value instanceof Date ? value : new Date(value);
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

  // Calculate min/max expiration dates
  const now = new Date();
  config.minExpirationDate = new Date(now.getFullYear(), now.getMonth() + config.minExpirationMonths, now.getDate());
  config.maxExpirationDate = new Date(now.getFullYear(), now.getMonth() + config.maxExpirationMonths, now.getDate());

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
 * Estimates current stock price from call options by finding the strike with delta closest to 0.5.
 * @param {Array} calls Array of call option objects with delta and strike.
 * @return {number} Estimated current stock price.
 */
function estimateCurrentPrice_(calls) {
  let bestDelta = Infinity;
  let bestStrike = 0;
  for (const c of calls) {
    const dist = Math.abs(Math.abs(c.delta) - 0.5);
    if (dist < bestDelta) {
      bestDelta = dist;
      bestStrike = c.strike;
    }
  }
  return bestStrike || 0;
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
      if (width < config.minSpreadWidth || width > config.maxSpreadWidth) continue;

      // Skip if no valid bid/ask
      if (lower.ask <= 0 || upper.bid < 0) continue;

      // Calculate debit using mid pricing (validated by executed trades with GT 60 patience)
      // For patient orders, mid is achievable; below mid may not fill
      const lowerMid = (lower.bid + lower.ask) / 2;
      const upperMid = (upper.bid + upper.ask) / 2;

      let debit = lowerMid - upperMid;
      if (debit < 0) debit = 0;
      debit = roundTo_(debit, 2);

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

      // Expected gain using probability-of-touch model (80% of max profit target)
      const expectedGain = calculateExpectedGain(lowerMid, upperMid, lower.strike, upper.strike, Math.abs(upper.delta));
      const expectedROI = debit > 0 ? expectedGain / debit : 0;

      // Fitness = ExpROI * liquidity^0.1 * tightness^0.1
      // Liquidity/tightness as mild tiebreakers (patient fills assumed)
      // timeFactor dropped — already baked into delta and probTouch
      // Outlook boost: adjust fitness based on price target, date, and confidence
      let outlookBoost = 1;
      if (config.outlookFuturePrice > 0 && config.outlookConfidence > 0) {
        const target = config.outlookFuturePrice;
        const conf = config.outlookConfidence;

        // Price proximity boost
        let priceBoost;
        if (lower.strike >= target) {
          // Both strikes above target — graduated penalty (further above = worse)
          const overshoot = (lower.strike - target) / target;
          priceBoost = 1 - conf * 0.5 * overshoot;
        } else if (upper.strike <= target) {
          // Both strikes below target — full boost by proximity
          priceBoost = 1 + conf * (upper.strike / target);
        } else {
          // Straddles target — partial boost by how much width is captured
          const captured = (target - lower.strike) / width;
          priceBoost = 1 + conf * captured * 0.5;
        }

        // Date proximity boost: expirations near outlookDate get more boost
        // Expirations before target date are penalized (may expire before move happens)
        let dateBoost = 1;
        if (config.outlookDate) {
          const expDate = new Date(lower.expiration);
          const targetDate = new Date(config.outlookDate);
          const now = new Date();
          const totalDays = Math.max(1, (targetDate - now) / (1000 * 60 * 60 * 24));
          const diffDays = (expDate - targetDate) / (1000 * 60 * 60 * 24);

          if (diffDays < 0) {
            // Expires before target date — penalize proportionally
            // Expiring way before target = bigger penalty
            const earlyRatio = Math.abs(diffDays) / totalDays;
            dateBoost = 1 - conf * Math.min(earlyRatio, 0.5);
          } else {
            // Expires on or after target date — boost, with falloff for much later
            const lateRatio = diffDays / totalDays;
            dateBoost = 1 + conf * Math.max(0, 0.3 - lateRatio * 0.2);
          }
        }

        outlookBoost = priceBoost * dateBoost;
      }

      const fitness = roundTo_(expectedROI * Math.pow(liquidityScore, 0.2) * Math.pow(tightness, 0.1) * outlookBoost, 2);

      spreads.push({
        symbol: lower.symbol,
        expiration: lower.expiration,
        lowerStrike: lower.strike,
        upperStrike: upper.strike,
        width,
        debit,
        maxProfit: roundTo_(maxProfit, 2),
        maxLoss: roundTo_(maxLoss, 2),
        roi: roundTo_(roi, 2),
        lowerIV: roundTo_(lower.iv || 0, 2),
        lowerDelta: roundTo_(lower.delta, 2),
        upperDelta: roundTo_(upper.delta, 2),
        lowerOI: lower.openint,
        upperOI: upper.openint,
        lowerVol: lower.volume,
        upperVol: upper.volume,
        expectedGain: roundTo_(expectedGain, 2),
        expectedROI: roundTo_(expectedROI, 2),
        liquidityScore: roundTo_(liquidityScore, 2),
        tightness: roundTo_(tightness, 2),
        fitness: roundTo_(fitness, 2)
      });
    }
  }

  return spreads;
}

/**
 * Loads held short positions from the Positions sheet BullCallSpreads table.
 * Returns a Set of "SYMBOL|STRIKE|EXPIRATION" keys for fast lookup.
 */
function loadHeldPositions_(ss) {
  const held = new Set();
  const sheet = ss.getSheetByName("Positions");
  if (!sheet) {
    Logger.log("Positions sheet not found, skipping held position check");
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
    Logger.log("BullCallSpreads table not found on Positions sheet");
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

    let rawExp = expirationCol >= 0 ? data[r][expirationCol] : "";
    if (rawExp instanceof Date) {
      lastExp = Utilities.formatDate(rawExp, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (rawExp) {
      const parsed = new Date(rawExp);
      if (!isNaN(parsed.getTime())) {
        lastExp = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
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
    const expDate = new Date(s.expiration);
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
    rowData[colIdx.OptionStrat] = "";  // formula set separately
    rowData[colIdx.Label] = label;
    rowData[colIdx.Held] = s.held ? "HELD" : "";
    rowData[colIdx.IV] = s.lowerIV;
    allRows.push(rowData);
  }
  sheet.getRange(dataStartRow, 1, allRows.length, headers.length).setValues(allRows);

  // OptionStrat formulas reference other columns by letter
  const lowerLetter = colLetter("Lower"), upperLetter = colLetter("Upper");
  const symLetter = colLetter("Symbol"), expLetter = colLetter("Expiration");
  const optionStratFormulas = spreads.map((s, i) => {
    const row = dataStartRow + i;
    const osUrl = `buildOptionStratUrl(${lowerLetter}${row}&"/"&${upperLetter}${row},${symLetter}${row},"bull-call-spread",${expLetter}${row})`;
    return [`=HYPERLINK(${osUrl},"OptionStrat")`];
  });
  sheet.getRange(dataStartRow, col("OptionStrat"), spreads.length, 1).setFormulas(optionStratFormulas);

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

/**
 * Opens a large dashboard window with Delta vs ROI and Strike vs ROI.
 */
function showSpreadFinderGraphs() {
  SpreadsheetApp.flush();

  // Creates the SpreadFinderGraphs modal dialog
  const html = HtmlService.createHtmlOutputFromFile('SpreadFinderGraphs')
      .setWidth(1050) // Wide enough for side-by-side or large stacked charts
      .setHeight(850);

  SpreadsheetApp.getUi().showModalDialog(html, 'Spread Finder Graphs');
}

/**
 * Fetches spread data for SpreadFinderGraphs.
 * Orders by Fitness so the best points are drawn last (on top).
 */
 function getSpreadFinderGraphData() {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const sheet = ss.getSheetByName(SPREADS_SHEET);
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
     const expDate = new Date(row[c.Expiration]);
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
       tightness: row[c.Tightness],
       dte: dte > 0 ? dte : 0,
       held: (row[c.Held] || "").toString().trim() === "HELD",
       iv: parseFloat(row[c.IV]) || 0
     };
   }).sort((a, b) => a.fitness - b.fitness);
 }
/**
 * Wrapper for Stubs.ts compatibility (called from menu/dialog).
 * Delegates to the main runSpreadFinder() — you can enhance filtering later.
 * @param {string[]} [symbols] 
 * @param {string[]} [expirations]
 * @returns {any}
 */
function runSpreadFinderWithSelection(symbols?: string[], expirations?: string[]) {
  const ui = SpreadsheetApp.getUi();
  try {
    // Future enhancement: filter by symbols/expirations if provided
    return runSpreadFinder();
  } catch (e) {
    console.error('runSpreadFinderWithSelection error:', e);
    ui.alert('Spread Finder Error', e.message || 'Unknown error occurred', ui.ButtonSet.OK);
    throw e;
  }
}
