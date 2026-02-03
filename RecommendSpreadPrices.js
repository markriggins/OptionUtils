/**
 * Bull Call Spread pricing helpers (OPEN debit / CLOSE credit)
 *
 * These functions provide recommended limit prices for opening and closing bull call spreads using option quotes from the "OptionPricesUploaded" sheet.
 * They leverage cached lookups via XLookupByKeys for efficiency.
 *
 * Expected sheet headers (case-insensitive): symbol, expiration, strike, type, bid, mid, ask, volume, openint.
 *
 * Expiration inputs are normalized to "YYYY-MM-DD" format (accepts strings or Date objects).
 * Strikes are converted to numbers.
 *
 * The avgMinutesToExecute parameter models execution time/patience:
 * - 0: Aggressive (worst-case fills at current ask/bid).
 * - Higher values: More patient, improving toward mid price.
 *
 * Liquidity-aware pricing:
 * - Uses volume and open interest to calculate a liquidity score (0-1).
 * - High liquidity (volume+OI ~1000+): patient orders can approach mid price.
 * - Low liquidity: stuck near aggressive prices even with patience.
 * - This prevents unrealistically optimistic prices for illiquid options.
 *
 * Returns a single rounded number (to 2 decimals) or an error string starting with "#".
 *
 * Dependencies: Requires XLookupByKeys and its cache setup.
 *
 * Version: 2.0 - Added liquidity-aware pricing
 */

/**
 * Recommends debit to OPEN a bull call spread (BUY lower, SELL upper).
 *
 * @param {string} symbol - Ticker (e.g. "TSLA")
 * @param {Date|string} expiration - Expiration date
 * @param {number} lowerStrike - Strike to BUY
 * @param {number} upperStrike - Strike to SELL
 * @param {number} avgMinutesToExecute - Patience: 0=aggressive, 60=patient
 * @param {Array} [_labels] - Optional; ignored, for readability
 * @return {number} Debit per share or error
 * @customfunction
 */
function recommendBullCallSpreadOpenDebit(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute, _labels) {
  /*
   * Examples:
   *   Spreadsheet: =recommendBullCallSpreadOpenDebit("TSLA", "2028-06-16", 450, 500, 30)
   *   Script:      recommendBullCallSpreadOpenDebit("TSLA", new Date("2028-06-16"), 450, 500, 60)
   *
   * At alpha=0: debit = lower.ask - upper.bid (aggressive)
   * As alpha increases: moves toward lower.bid - upper.ask (patient, may not fill)
   */
  const parsed = normalizeSpreadInputs_(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute);
  if (!parsed) return "#Could not parse inputs";
  const { sym, exp, lo, hi, alpha } = parsed;
  // OPEN a bull call spread:
  // BUY the lower strike call
  // SELL the upper strike call
  const lower = getOptionQuote_(sym, exp, lo, "Call"); // you BUY this to open
  const upper = getOptionQuote_(sym, exp, hi, "Call"); // you SELL this to open
  if (!hasBidAsk_(lower)) return "#No Data for Lower:" + lowerStrike;
  if (!hasBidAsk_(upper)) return "#No Data for Upper:" + upperStrike;
  // Use liquidity-aware pricing:
  // - High liquidity + patience → can approach mid
  // - Low liquidity → stuck near aggressive prices
  const buyLimit = getRealisticBuyPrice_(lower, alpha);
  const sellLimit = getRealisticSellPrice_(upper, alpha);
  let debit = buyLimit - sellLimit;
  if (debit < 0) debit = 0;
  return roundTo_(debit, 2);
}

/**
 * Recommends credit to CLOSE a bull call spread (SELL lower, BUY upper).
 *
 * @param {string} symbol - Ticker (e.g. "TSLA")
 * @param {Date|string} expiration - Expiration date
 * @param {number} lowerStrike - Strike to SELL (close long)
 * @param {number} upperStrike - Strike to BUY (close short)
 * @param {number} avgMinutesToExecute - Patience: 0=aggressive, 60=patient
 * @param {Array} [_labels] - Optional; ignored, for readability
 * @return {number} Credit per share or error
 * @customfunction
 */
function recommendBullCallSpreadCloseCredit(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute, _labels) {
  /*
   * Examples:
   *   Spreadsheet: =recommendBullCallSpreadCloseCredit("TSLA", "2028-06-16", 450, 500, 30)
   *   Script:      recommendBullCallSpreadCloseCredit("TSLA", new Date("2028-06-16"), 450, 500, 60)
   *
   * At alpha=0: credit = lower.bid - upper.ask (aggressive)
   * As alpha increases: moves toward lower.ask - upper.bid (patient)
   */
  const parsed = normalizeSpreadInputs_(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute);
  if (!parsed) return "#Could not parse inputs";
  const { sym, exp, lo, hi, alpha } = parsed;
  const lower = getOptionQuote_(sym, exp, lo, "Call"); // you SELL this to close
  const upper = getOptionQuote_(sym, exp, hi, "Call"); // you BUY this to close
  if (!hasBidAsk_(lower)) return "#No Data for Lower:" + lowerStrike;
  if (!hasBidAsk_(upper)) return "#No Data for Upper:" + upperStrike;
  // Use liquidity-aware pricing
  const sellLimit = getRealisticSellPrice_(lower, alpha);
  const buyLimit = getRealisticBuyPrice_(upper, alpha);
  let credit = sellLimit - buyLimit;
  if (credit < 0) credit = 0;
  return roundTo_(credit, 2);
}

/**
 * normalizeSpreadInputs_ - Parses and validates inputs for bull call spread functions.
 *
 * Internal helper: Normalizes symbol, expiration, strikes, and computes alpha from minutes.
 * Alpha uses exponential decay: 1 - exp(-mins / 60), saturating at 1 for large mins.
 *
 * Returns null on invalid inputs.
 *
 * @param {string} symbol - Stock symbol.
 * @param {string|Date} expiration - Expiration date.
 * @param {number} lowerStrike - Lower strike.
 * @param {number} upperStrike - Upper strike.
 * @param {number} avgMinutesToExecute - Patience in minutes.
 * @returns {Object|null} Parsed {sym, exp, lo, hi, alpha} or null if invalid.
 */

function normalizeSpreadInputs_(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute) {
  const sym = (symbol || "").toString().trim().toUpperCase(); // Faster toString
  const exp = normalizeExpiration_(expiration);
  const lo = +lowerStrike;
  const hi = +upperStrike;
  let mins = +avgMinutesToExecute;
  if (!sym || !exp || !Number.isFinite(lo) || !Number.isFinite(hi) || lo >= hi) return null;
  if (!Number.isFinite(mins) || mins < 0) mins = 0;
  // Aggressiveness curve (saturating)
  // mins=0 => alpha=0 (cross)
  // mins=60 => alpha≈0.63
  // mins=1440 => alpha≈1
  const HALF_LIFE_MIN = 60;
  const alpha = 1 - Math.exp(-mins / HALF_LIFE_MIN);
  return { sym, exp, lo, hi, alpha };
}

/**
 * normalizeExpiration_ - Normalizes expiration to "YYYY-MM-DD" string.
 *
 * Accepts Date objects or valid "YYYY-MM-DD" strings.
 * Returns null if invalid.
 *
 * @param {string|Date} expiration - Input expiration.
 * @returns {string|null} Normalized "YYYY-MM-DD" or null.
 */

function normalizeExpiration_(expiration) {
  if (expiration instanceof Date) {
    return Utilities.formatDate(expiration, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  if (typeof expiration === "string") {
    const s = expiration.trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  }
  return null;
}

/**
 * hasBidAsk_ - Checks if a quote object has valid finite bid and ask.
 *
 * @param {Object} q - Quote {bid, ask}.
 * @returns {boolean} True if valid bid/ask present.
 */

function hasBidAsk_(q) {
  return q && q.bid != null && q.ask != null && Number.isFinite(q.bid) && Number.isFinite(q.ask);
}

/**
 * getOptionQuote_ - Fetches option quote (bid, mid, ask) using cached XLookupByKeys.
 *
 * Caches results in ScriptCache for 5 minutes to reduce repeated lookups.
 * Sheet: "OptionPricesUploaded".
 *
 * Returns null if no data found.
 *
 * @param {string} symbol - Stock symbol (uppercased).
 * @param {string} expiration - "YYYY-MM-DD".
 * @param {number} strike - Strike price.
 * @param {string} type - "Call" or "Put".
 * @returns {Object|null} {bid, mid, ask} or null.
 */

function getOptionQuote_(symbol, expiration, strike, type) {
  const cache = CacheService.getScriptCache();
  const cacheKey = [symbol, expiration, strike, type].join('|');
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const SHEET_NAME = "OptionPricesUploaded";
  const res = XLookupByKeys(
    [symbol, expiration, +strike, type],
    ["symbol", "expiration", "strike", "type"],
    ["bid", "mid", "ask", "volume", "openint"],
    SHEET_NAME
  );
  const row = res && res[0];
  if (!row) return null;
  const quote = {
    bid: row[0] === "" ? null : +row[0],
    mid: row[1] === "" ? null : +row[1],
    ask: row[2] === "" ? null : +row[2],
    volume: row[3] === "" ? 0 : +row[3],
    openint: row[4] === "" ? 0 : +row[4]
  };
  cache.put(cacheKey, JSON.stringify(quote), 300); // Cache for 5 min
  return quote;
}

/**
 * Calculate liquidity score from volume and open interest.
 * Returns 0 (illiquid) to 1 (very liquid).
 *
 * @param {Object} quote - Quote with volume and openint
 * @returns {number} Liquidity score 0-1
 */
function getLiquidityScore_(quote) {
  const volume = quote.volume || 0;
  const openint = quote.openint || 0;
  // Log scale: need ~1000 combined activity for full liquidity
  // volume counts more than OI (OI/10)
  return Math.min(1, Math.log10(1 + volume + openint / 10) / 3);
}

/**
 * Get realistic buy price accounting for liquidity.
 *
 * - alpha=0: pay ask (aggressive)
 * - alpha=1, high liquidity: approaches mid
 * - alpha=1, low liquidity: stuck near ask
 *
 * @param {Object} quote - Quote with bid, mid, ask, volume, openint
 * @param {number} alpha - Patience factor 0-1
 * @returns {number} Realistic buy limit price
 */
function getRealisticBuyPrice_(quote, alpha) {
  const liquidity = getLiquidityScore_(quote);
  // Effective alpha: patience scaled by liquidity
  const effectiveAlpha = alpha * liquidity;
  // Interpolate from ask toward mid based on effective alpha
  return quote.ask - effectiveAlpha * (quote.ask - quote.mid);
}

/**
 * Get realistic sell price accounting for liquidity.
 *
 * - alpha=0: receive bid (aggressive)
 * - alpha=1, high liquidity: approaches mid
 * - alpha=1, low liquidity: stuck near bid
 *
 * @param {Object} quote - Quote with bid, mid, ask, volume, openint
 * @param {number} alpha - Patience factor 0-1
 * @returns {number} Realistic sell limit price
 */
function getRealisticSellPrice_(quote, alpha) {
  const liquidity = getLiquidityScore_(quote);
  // Effective alpha: patience scaled by liquidity
  const effectiveAlpha = alpha * liquidity;
  // Interpolate from bid toward mid based on effective alpha
  return quote.bid + effectiveAlpha * (quote.mid - quote.bid);
}

/**
 * Recommends credit to OPEN a short iron condor.
 *
 * @param {string} symbol - Ticker (e.g. "TSLA")
 * @param {Date|string} expiration - Expiration date
 * @param {number} buyPut - Lowest put (buy protective)
 * @param {number} sellPut - Higher put (sell short)
 * @param {number} sellCall - Lower call (sell short)
 * @param {number} buyCall - Highest call (buy protective)
 * @param {number} avgMinutesToExecute - Patience: 0=aggressive, 60=patient
 * @param {Array} [_labels] - Optional; ignored, for readability
 * @return {number} Credit per share or error
 * @customfunction
 */
function recommendIronCondorOpenCredit(
  symbol,
  expiration,
  buyPut,
  sellPut,
  sellCall,
  buyCall,
  avgMinutesToExecute,
  _labels
) {
  /*
   * Examples:
   *   Spreadsheet: =recommendIronCondorOpenCredit("TSLA", "2028-06-16", 400, 420, 480, 500, 30)
   *   Script:      recommendIronCondorOpenCredit("TSLA", new Date("2028-06-16"), 400, 420, 480, 500, 60)
   *
   * Strikes must satisfy: buyPut < sellPut and sellCall < buyCall
   */
  const parsed = normalizeIronCondorInputs_(
    symbol,
    expiration,
    buyPut,
    sellPut,
    sellCall,
    buyCall,
    avgMinutesToExecute
  );
  if (parsed.error) return parsed.error;
  const { sym, exp, bp, sp, sc, bc, alpha } = parsed;
  // Quotes
  const qBuyPut = getOptionQuote_(sym, exp, bp, "Put"); // BUY
  const qSellPut = getOptionQuote_(sym, exp, sp, "Put"); // SELL
  const qBuyCall = getOptionQuote_(sym, exp, bc, "Call"); // BUY
  const qSellCall = getOptionQuote_(sym, exp, sc, "Call"); // SELL
  if (!hasBidAsk_(qBuyPut)) return "#No Data for Buy Put:" + sym + " " + exp + " @" + buyPut;
  if (!hasBidAsk_(qSellPut)) return "#No Data for Sell Put:" + sym + " " + exp + " @" + sellPut;
  if (!hasBidAsk_(qSellCall)) return "#No Data for Sell Call:" + sym + " " + exp + " @" + sellCall;
  if (!hasBidAsk_(qBuyCall)) return "#No Data for Buy Call:" + sym + " " + exp + " @" + buyCall;
  // Use liquidity-aware pricing
  const buyPutLimit = getRealisticBuyPrice_(qBuyPut, alpha);
  const buyCallLimit = getRealisticBuyPrice_(qBuyCall, alpha);
  const sellPutLimit = getRealisticSellPrice_(qSellPut, alpha);
  const sellCallLimit = getRealisticSellPrice_(qSellCall, alpha);
  // Net credit to open
  let credit = (sellPutLimit + sellCallLimit) - (buyPutLimit + buyCallLimit);
  if (credit < 0) credit = 0;
  return roundTo_(credit, 2);
}

/**
 * Recommends debit to CLOSE a short iron condor.
 *
 * @param {string} symbol - Ticker (e.g. "TSLA")
 * @param {Date|string} expiration - Expiration date
 * @param {number} buyPut - Original buy put (sell to close)
 * @param {number} sellPut - Original sell put (buy to close)
 * @param {number} buyCall - Original buy call (sell to close)
 * @param {number} sellCall - Original sell call (buy to close)
 * @param {number} avgMinutesToExecute - Patience: 0=aggressive, 60=patient
 * @param {Array} [_labels] - Optional; ignored, for readability
 * @return {number} Debit per share or error
 * @customfunction
 */
function recommendIronCondorCloseDebit(
  symbol,
  expiration,
  buyPut,
  sellPut,
  buyCall,
  sellCall,
  avgMinutesToExecute,
  _labels
) {
  /*
   * Examples:
   *   Spreadsheet: =recommendIronCondorCloseDebit("TSLA", "2028-06-16", 400, 420, 480, 500, 30)
   *   Script:      recommendIronCondorCloseDebit("TSLA", new Date("2028-06-16"), 400, 420, 480, 500, 60)
   */
  const parsed = normalizeIronCondorInputs_(
    symbol,
    expiration,
    buyPut,
    sellPut,
    buyCall,
    sellCall,
    avgMinutesToExecute
  );
  if (parsed.error) return parsed.error;
  const { sym, exp, bp, sp, bc, sc, alpha } = parsed;
  // Quotes for each leg
  const qLongPut = getOptionQuote_(sym, exp, bp, "Put"); // was BUY to open -> SELL to close
  const qShortPut = getOptionQuote_(sym, exp, sp, "Put"); // was SELL to open -> BUY to close
  const qLongCall = getOptionQuote_(sym, exp, bc, "Call"); // was BUY to open -> SELL to close
  const qShortCall = getOptionQuote_(sym, exp, sc, "Call"); // was SELL to open -> BUY to close
  if (!hasBidAsk_(qLongPut)) return "#No Data for Buy Put:" + buyPut;
  if (!hasBidAsk_(qShortPut)) return "#No Data for Sell Put:" + sellPut;
  if (!hasBidAsk_(qLongCall)) return "#No Data for Buy Call:" + buyCall;
  if (!hasBidAsk_(qShortCall)) return "#No Data for Sell Call:" + sellCall;
  // Use liquidity-aware pricing
  const sellLongPutLimit = getRealisticSellPrice_(qLongPut, alpha);
  const sellLongCallLimit = getRealisticSellPrice_(qLongCall, alpha);
  const buyShortPutLimit = getRealisticBuyPrice_(qShortPut, alpha);
  const buyShortCallLimit = getRealisticBuyPrice_(qShortCall, alpha);
  // Net debit to close
  let debit = (buyShortPutLimit + buyShortCallLimit) - (sellLongPutLimit + sellLongCallLimit);
  if (debit < 0) debit = 0;
  return roundTo_(debit, 2);
}

/**
 * Recommends price to close a single option leg.
 *
 * @param {string|Range} symbol - Ticker (e.g. "TSLA"), or a range like $A$1:$A3 to find first non-blank looking upward
 * @param {Date|string} expiration - Expiration date
 * @param {number} strike - Strike price
 * @param {string} type - "Call" or "Put"
 * @param {number} qty - Position quantity (positive=long, negative=short)
 * @param {number} avgMinutesToExecute - Patience: 0=aggressive, 60=patient
 * @param {Array} [_labels] - Optional; ignored, for spreadsheet readability
 * @return {number} Credit (positive) if closing long, Debit (negative) if closing short, or error string
 * @customfunction
 */
function recommendClose(symbol, expiration, strike, type, qty, avgMinutesToExecute, _labels) {
  /*
   * Examples:
   *   Spreadsheet: =recommendClose("TSLA", "2028-06-16", 450, "Call", 7, 60)
   *   Spreadsheet: =recommendClose($A$1:$A3, "2028-06-16", 450, "Call", -7, 60)  // range for merged/grouped symbols
   *
   * Closing a LONG position (qty > 0): SELL to close → returns credit (positive)
   * Closing a SHORT position (qty < 0): BUY to close → returns debit (negative)
   */
  const parsed = normalizeLegInputs_(symbol, expiration, strike, type, qty, avgMinutesToExecute);
  if (parsed.error) return parsed.error;

  const { sym, exp, k, optType, position, alpha } = parsed;

  const quote = getOptionQuote_(sym, exp, k, optType);
  if (!hasBidAsk_(quote)) return "#No Data:" + sym + " " + exp + " " + k + " " + optType;

  if (position > 0) {
    // Long position: SELL to close → credit (positive)
    const sellLimit = getRealisticSellPrice_(quote, alpha);
    return roundTo_(sellLimit, 2);
  } else {
    // Short position: BUY to close → debit (negative)
    const buyLimit = getRealisticBuyPrice_(quote, alpha);
    return roundTo_(-buyLimit, 2);
  }
}

/**
 * normalizeLegInputs_ - Parses and validates inputs for single leg functions.
 *
 * @param {string|Array} symbol - Stock symbol, or a vertical range to search upward for first non-blank.
 * @param {string|Date} expiration - Expiration date.
 * @param {number} strike - Strike price.
 * @param {string} type - "Call" or "Put".
 * @param {number} qty - Position quantity (positive=long, negative=short).
 * @param {number} avgMinutesToExecute - Patience in minutes.
 * @returns {Object} {sym, exp, k, optType, position, alpha, error} (error null if valid).
 */
function normalizeLegInputs_(symbol, expiration, strike, type, qty, avgMinutesToExecute) {
  // Handle range input: find last non-blank value (bottom-up, for merged cells / grouped rows)
  let symRaw = symbol;
  if (Array.isArray(symbol)) {
    const flat = symbol.flat ? symbol.flat() : [].concat(...symbol);
    symRaw = "";
    for (let i = flat.length - 1; i >= 0; i--) {
      const v = (flat[i] ?? "").toString().trim();
      if (v) { symRaw = v; break; }
    }
  }
  const sym = (symRaw || "").toString().trim().toUpperCase();
  if (!sym) return { error: "#Symbol required" };

  const exp = normalizeExpiration_(expiration);
  if (!exp) return { error: "#Bad expiration" };

  const k = +strike;
  if (!Number.isFinite(k) || k <= 0) return { error: "#Bad strike" };

  const optType = parseOptionType_(type);
  if (!optType) return { error: "#Bad type (use Call or Put)" };

  const position = +qty;
  if (!Number.isFinite(position) || position === 0) return { error: "#Bad qty (non-zero required)" };

  let mins = +avgMinutesToExecute;
  if (!Number.isFinite(mins) || mins < 0) mins = 0;

  const HALF_LIFE_MIN = 60;
  const alpha = 1 - Math.exp(-mins / HALF_LIFE_MIN);

  return { sym, exp, k, optType, position, alpha, error: null };
}

/**
 * normalizeIronCondorInputs_ - Parses and validates inputs for iron condor functions.
 *
 * Validates symbol, expiration, finite strikes, and order (buyPut < sellPut, sellCall < buyCall).
 * Computes alpha similarly to spreads.
 *
 * Returns {error: string} if invalid, else parsed object.
 *
 * @param {string} symbol - Stock symbol.
 * @param {string|Date} expiration - Expiration.
 * @param {number} buyPut - Buy put strike.
 * @param {number} sellPut - Sell put strike.
 * @param {number} sellCall - Sell call strike.
 * @param {number} buyCall - Buy call strike.
 * @param {number} avgMinutesToExecute - Patience in minutes.
 * @returns {Object} {sym, exp, bp, sp, sc, bc, alpha, error} (error null if valid).
 */

function normalizeIronCondorInputs_(
  symbol,
  expiration,
  buyPut,
  sellPut,
  sellCall,
  buyCall,
  avgMinutesToExecute
) {
  const sym = (symbol || "").toString().trim().toUpperCase();
  if (!sym) return { error: "#Symbol required" };
  const exp = normalizeExpiration_(expiration);
  if (!exp) return { error: "#Bad expiration" };
  const bp = +buyPut;
  const sp = +sellPut;
  const sc = +sellCall;
  const bc = +buyCall;
  if (![bp, sp, sc, bc].every(Number.isFinite)) return { error: "#Bad strikes" };
  // Basic sanity for a typical short IC:
  // buyPut < sellPut and sellCall < buyCall
  if (!(bp < sp)) return { error: "#Put strikes must be buyPut < sellPut: " + bp + " !< " + sp };
  if (!(sc < bc)) return { error: "#Call strikes must be sellCall < buyCall: " + sc + " !< " + bc };
  let mins = +avgMinutesToExecute;
  if (!Number.isFinite(mins) || mins < 0) mins = 0;
  // Same alpha curve concept as spreads (tune if desired)
  const HALF_LIFE_MIN = 60;
  const alpha = 1 - Math.exp(-mins / HALF_LIFE_MIN);
  return { sym, exp, bp, sp, sc, bc, alpha, error: null };
}
