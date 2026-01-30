/**
 * SpreadFinder.js
 * Analyzes OptionPricesUploaded to find and rank bull call spread opportunities.
 *
 * Config lives on the SpreadFinderConfig sheet.
 * Results are written to the Spreads sheet.
 *
 * Config table format (auto-created on SpreadFinderConfig):
 *   | Setting            | Value | Description                                    |
 *   | maxSpreadWidth     | 150   | Maximum spread width in dollars                |
 *   | minOpenInterest    | 10    | Minimum open interest for both legs            |
 *   | minVolume          | 0     | Minimum volume for both legs                   |
 *   | patience           | 60    | Minutes for price calculation (0=aggressive)   |
 *   | maxDebit           | 50    | Maximum debit per share                        |
 *   | minROI             | 0.5   | Minimum ROI (0.5 = 50%)                        |
 *
 * Version: 2.0
 */

const SPREAD_FINDER_CONFIG_SHEET = "SpreadFinderConfig";
const SPREADS_SHEET = "Spreads";
const OPTION_PRICES_SHEET = "OptionPricesUploaded";
const CONFIG_COL = 1; // Column A
const CONFIG_START_ROW = 1;

/**
 * Calculates the Expected Gain for a Bull Call Spread based on an 80%-of-max-profit early exit.
 * Uses the "Rule of Touch" (probTouch ≈ 1.6x delta) to estimate probability of reaching target.
 * @param {number} longMid The mid price of the lower (long) leg.
 * @param {number} shortMid The mid price of the upper (short) leg.
 * @param {number} longStrike The strike price of the lower leg.
 * @param {number} shortStrike The strike price of the upper leg.
 * @param {number} shortDelta The delta of the upper (short) leg.
 * @return {number} The expected dollar gain per spread.
 */
function calculateExpectedGain(longMid, shortMid, longStrike, shortStrike, shortDelta) {
  var netDebit = longMid - shortMid;
  var spreadWidth = shortStrike - longStrike;
  var maxProfit = spreadWidth - netDebit;

  var targetProfit = maxProfit * 0.80;

  // Prob(Touch) ≈ 1.6x short delta, capped at 95%
  var probTouch = Math.min(shortDelta * 1.6, 0.95);
  var probLoss = 1 - probTouch;

  // EV = (Prob of Win * Win Amount) + (Prob of Loss * Loss Amount)
  var expectedValue = (probTouch * targetProfit) + (probLoss * -netDebit);

  return expectedValue;
}

/**
 * Runs the spread finder analysis. Call from menu or script.
 * Ensures config exists, reads it, scans options, ranks spreads, outputs results.
 */
function runSpreadFinder() {
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
    config.outlookFuturePrice = round2_(currentPrice * 1.25);
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

  // Load conflicts from config sheet
  const conflicts = loadConflicts_(configSheet);
  Logger.log("Loaded " + conflicts.size + " conflicts");

  // Filter by config constraints
  const minExpDate = config.minExpirationDate;
  const maxExpDate = config.maxExpirationDate;
  const filtered = spreads.filter(s => {
    // Parse expiration date
    const expDate = new Date(s.expiration);
    // Check if this spread's lower strike conflicts with a held short position
    const hasConflict = conflicts.has(`${s.symbol}|${s.lowerStrike}|${s.expiration}`);
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
      expDate <= maxExpDate &&
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

  // --- Conflicts table ---
  const conflictsStartRow = CONFIG_START_ROW + configData.length + 1; // blank row separator

  // Read existing conflicts before overwriting
  const existingConflicts = [];
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow > conflictsStartRow + 1) {
      const conflictData = sheet.getRange(conflictsStartRow + 2, CONFIG_COL, lastRow - conflictsStartRow - 1, 3).getValues();
      for (const row of conflictData) {
        const sym = (row[0] || "").toString().trim();
        if (sym) existingConflicts.push(row);
      }
    }
  } catch (e) { /* no existing conflicts */ }

  // If no existing conflicts, seed from Positions sheet
  if (existingConflicts.length === 0) {
    const posConflicts = seedConflictsFromPositions_(ss);
    existingConflicts.push(...posConflicts);
  }

  // Write conflicts header + description
  sheet.getRange(conflictsStartRow, CONFIG_COL, 1, 3).setValues([
    ["Enter specific symbol, strike and expiration dates to ignore (presuming you already have conflicting short positions in them)", "", ""]
  ]).setFontStyle("italic").setFontColor("#5f6368").setFontSize(9);
  sheet.getRange(conflictsStartRow, CONFIG_COL).merge();

  const conflictHeader = [["Symbol", "Strike", "Expiration"]];
  const conflictHeaderRow = conflictsStartRow + 1;
  sheet.getRange(conflictHeaderRow, CONFIG_COL, 1, 3).setValues(conflictHeader)
    .setBackground("#e8710a").setFontColor("white").setFontWeight("bold");

  // Write conflict data rows (or one empty row for user to fill in)
  if (existingConflicts.length > 0) {
    sheet.getRange(conflictHeaderRow + 1, CONFIG_COL, existingConflicts.length, 3).setValues(existingConflicts);
  }

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

      const fitness = round2_(expectedROI * Math.pow(liquidityScore, 0.2) * Math.pow(tightness, 0.1) * outlookBoost);

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
        expectedGain: round2_(expectedGain),
        expectedROI: round2_(expectedROI),
        liquidityScore: round2_(liquidityScore),
        tightness: round2_(tightness),
        fitness: round2_(fitness)
      });
    }
  }

  return spreads;
}

/**
 * Seeds initial conflicts from the Positions sheet BullCallSpreads table.
 * Returns array of [symbol, strike, expiration] rows.
 */
function seedConflictsFromPositions_(ss) {
  const conflicts = [];
  const sheet = ss.getSheetByName("Positions");
  if (!sheet) return conflicts;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return conflicts;

  // Find header row with "Symbol", "Short Strike", "Expiration"
  let headerRow = -1;
  let symbolCol = -1, shortStrikeCol = -1, expirationCol = -1, contractsCol = -1;

  for (let r = 0; r < Math.min(10, data.length); r++) {
    for (let c = 0; c < data[r].length; c++) {
      const val = (data[r][c] || "").toString().toLowerCase().trim();
      if (val === "symbol") symbolCol = c;
      if (val === "short strike") { headerRow = r; shortStrikeCol = c; }
      if (val === "expiration") expirationCol = c;
      if (val === "contracts") contractsCol = c;
    }
    if (headerRow >= 0) break;
  }

  if (headerRow < 0 || shortStrikeCol < 0) return conflicts;

  for (let r = headerRow + 1; r < data.length; r++) {
    const symbol = symbolCol >= 0 ? (data[r][symbolCol] || "").toString().trim().toUpperCase() : "";
    const strike = +data[r][shortStrikeCol];
    const contracts = contractsCol >= 0 ? +data[r][contractsCol] : 1;
    let exp = expirationCol >= 0 ? data[r][expirationCol] : "";
    if (exp instanceof Date) {
      exp = Utilities.formatDate(exp, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    if (symbol && Number.isFinite(strike) && strike > 0 && contracts > 0) {
      conflicts.push([symbol, strike, exp || ""]);
    }
  }

  return conflicts;
}

/**
 * Loads conflicts from the SpreadFinderConfig sheet.
 * Returns a Set of "SYMBOL|STRIKE|EXPIRATION" keys for fast lookup.
 */
function loadConflicts_(configSheet) {
  const conflicts = new Set();
  const data = configSheet.getDataRange().getValues();

  // Find the conflicts header row
  let headerRow = -1;
  for (let r = 0; r < data.length; r++) {
    const a = (data[r][0] || "").toString().trim();
    const b = (data[r][1] || "").toString().trim();
    const c = (data[r][2] || "").toString().trim();
    if (a === "Symbol" && b === "Strike" && c === "Expiration") {
      headerRow = r;
      break;
    }
  }

  if (headerRow < 0) return conflicts;

  for (let r = headerRow + 1; r < data.length; r++) {
    const symbol = (data[r][0] || "").toString().trim().toUpperCase();
    const strike = +data[r][1];
    let exp = data[r][2];
    if (exp instanceof Date) {
      exp = Utilities.formatDate(exp, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      exp = (exp || "").toString().trim();
    }

    if (symbol && Number.isFinite(strike) && strike > 0) {
      conflicts.add(`${symbol}|${strike}|${exp}`);
    }
  }

  return conflicts;
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
    const clearRange = sheet.getRange(1, 1, lastRow, 20);
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

  // Headers (A-S)
  const headers = [
    "Symbol", "Expiration", "Lower", "Upper", "Width",
    "Debit", "MaxProfit", "ROI", "ExpGain", "ExpROI",
    "LowerDelta", "UpperDelta",
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
    "Label for chart identification"
  ];
  const hdrRange = sheet.getRange(RESULTS_START_ROW, 1, 1, headers.length);
  hdrRange.setValues([headers]).setFontWeight("bold");
  hdrRange.setNotes([headerNotes]);

  if (spreads.length === 0) return;

  const dataStartRow = RESULTS_START_ROW + 1;

  // Build all data, formulas, labels in one pass (no per-row sheet reads)
  const tz = Session.getScriptTimeZone();
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  // Build all formulas and values in one pass, write as single 2D array
  // Columns: A-F data, G formula, H formula, I-P data, Q formula, R formula, S value
  const allRows = [];
  for (let i = 0; i < spreads.length; i++) {
    const s = spreads[i];
    const row = dataStartRow + i;
    const expDate = new Date(s.expiration);
    const dateStr = months[expDate.getMonth()] + " " + String(expDate.getFullYear()).slice(2);
    const label = `${s.symbol} ${s.lowerStrike}/${s.upperStrike} ${dateStr}`;
    const osUrl = `buildOptionStratUrl(C${row}&"/"&D${row},A${row},"bull-call-spread",B${row})`;

    allRows.push([
      s.symbol, s.expiration, s.lowerStrike, s.upperStrike, s.width,
      s.debit,
      s.maxProfit,   // G - pre-computed instead of formula
      s.roi,         // H - pre-computed instead of formula
      s.expectedGain, s.expectedROI,
      s.lowerDelta, s.upperDelta, s.lowerOI, s.upperOI,
      s.liquidityScore, s.tightness,
      s.fitness,     // Q - pre-computed instead of formula
      "",            // R - OptionStrat (formula, set separately)
      label          // S
    ]);
  }
  sheet.getRange(dataStartRow, 1, allRows.length, headers.length).setValues(allRows);

  // Only OptionStrat needs formulas (HYPERLINK with custom function)
  const optionStratFormulas = spreads.map((s, i) => {
    const row = dataStartRow + i;
    const osUrl = `buildOptionStratUrl(C${row}&"/"&D${row},A${row},"bull-call-spread",B${row})`;
    return [`=HYPERLINK(${osUrl},"OptionStrat")`];
  });
  sheet.getRange(dataStartRow, 18, spreads.length, 1).setFormulas(optionStratFormulas);

  // Apply all number formats in one batch using a format array
  const formats = spreads.map(() => [
    "@", "@", "#,##0", "#,##0", "#,##0",           // A-E
    "$#,##0.00", "$#,##0.00", "0.00",               // F-H
    "$#,##0.00", "0.00",                             // I-J
    "0.00", "0.00", "#,##0", "#,##0",               // K-N
    "0.00", "0.00", "0.00",                          // O-Q
    "@", "@"                                          // R-S
  ]);
  sheet.getRange(dataStartRow, 1, spreads.length, headers.length).setNumberFormats(formats);

  // Style: header, banding, borders, filter in minimal API calls
  const tableRange = sheet.getRange(RESULTS_START_ROW, 1, spreads.length + 1, headers.length);
  tableRange.createFilter();
  tableRange.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(RESULTS_START_ROW, 1, 1, headers.length)
    .setBackground("#4285f4").setFontColor("white").setFontWeight("bold");

  const dataRange = sheet.getRange(dataStartRow, 1, spreads.length, headers.length);
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  // Set column widths in batch instead of autoResizeColumn loop
  const colWidths = [60, 90, 60, 60, 50, 70, 70, 50, 70, 55, 55, 55, 55, 55, 55, 55, 55, 100, 150];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  // Clip OptionStrat column
  sheet.getRange(RESULTS_START_ROW, 18, spreads.length + 1, 1)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

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
   const sheet = ss.getSheetByName(SPREADS_SHEET);
   const lastRow = sheet.getLastRow();
   const startRow = 3; // Row 1=timestamp, Row 2=headers, Row 3+=data
   if (lastRow < startRow) return [];

   const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 19).getValues();
   const today = new Date();
   today.setHours(0,0,0,0);

   return data.map(row => {
     const sym = row[0];
     const expDate = new Date(row[1]);
     const lowStrike = row[2];
     const highStrike = row[3];

     // Re-build the URL string directly instead of scraping the formula
     // Matches: optionstrat.com/build/bull-call-spread/TSLA/280119C350,280119C450 (example format)
     // Note: Adjust the buildOptionStratUrl logic if your function uses a different format
     const osUrl = buildOptionStratUrl(`${lowStrike}/${highStrike}`, sym, "bull-call-spread", expDate);

     const diffTime = expDate.getTime() - today.getTime();
     const dte = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

     return {
       delta: parseFloat(row[10]) || 0,
       roi: parseFloat(row[7]) || 0,
       strike: parseFloat(row[2]) || 0,
       fitness: parseFloat(row[16]) || 0,
       label: String(row[18] || ""),
       osUrl: osUrl, // Now a clean string

       width: row[4],
       debit: row[5],
       maxProfit: row[6],
       expectedGain: row[8],
       expectedROI: row[9],
       lowerDelta: row[10],
       upperDelta: row[11],
       lowerOI: row[12],
       upperOI: row[13],
       liquidity: row[14],
       tightness: row[15],
       dte: dte > 0 ? dte : 0
     };
   }).sort((a, b) => a.fitness - b.fitness);
 }