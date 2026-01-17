/**
 * Bull Call Spread pricing helpers (OPEN debit / CLOSE credit)
 *
 * These functions provide recommended limit prices for opening and closing bull call spreads using option quotes from the "OptionPricesUploaded" sheet.
 * They leverage cached lookups via XLookupByKeys for efficiency.
 *
 * Expected sheet headers (case-insensitive): symbol, expiration, strike, type, bid, mid, ask.
 *
 * Expiration inputs are normalized to "YYYY-MM-DD" format (accepts strings or Date objects).
 * Strikes are converted to numbers.
 *
 * The avgMinutesToExecute parameter models execution time/patience:
 * - 0: Aggressive (worst-case fills at current ask/bid).
 * - Higher values: More patient, improving toward better prices using an exponential decay curve (half-life 60 minutes).
 *
 * Returns a single rounded number (to 2 decimals) or an error string starting with "#".
 *
 * Dependencies: Requires XLookupByKeys and its cache setup.
 *
 * Version: 1.0
 */

/**
 * recommendBullCallSpreadOpenDebit - Recommends a debit limit price to OPEN a bull call spread.
 *
 * Strategy: BUY call at lowerStrike (long), SELL call at upperStrike (short).
 *
 * The price is the net debit to pay, adjusted for execution patience.
 * At alpha=0: debit = lower.ask - upper.bid (aggressive).
 * As alpha increases: Moves toward lower.bid - upper.ask (patient, but may not fill).
 * Minimum debit is 0.
 *
 * Usage Examples:
 *
 * 1. Spreadsheet Formula:
 *    =recommendBullCallSpreadOpenDebit("TSLA", "2028-06-16", 450, 500, 30)
 *    // Returns e.g., 5.25 (net debit per share)
 *
 * 2. From Script:
 *    const debit = recommendBullCallSpreadOpenDebit("TSLA", new Date("2028-06-16"), 450, 500, 60);
 *    // debit is a number or error string
 *
 * @param {string} symbol - The stock symbol (e.g., "TSLA"). Case-insensitive, trimmed.
 * @param {string|Date} expiration - Expiration date as "YYYY-MM-DD" string or Date object.
 * @param {number} lowerStrike - Lower strike price for the long call.
 * @param {number} upperStrike - Upper strike price for the short call (must be > lowerStrike).
 * @param {number} avgMinutesToExecute - Minutes of patience (0+); higher improves price but delays fill.
 * @returns {number|string} Recommended debit (rounded to 2 decimals) or "#Error message".
 */

function recommendBullCallSpreadOpenDebit(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute) {
  const parsed = parseSpreadInputs_(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute);
  if (!parsed) return "#Could not parse inputs";
  const { sym, exp, lo, hi, alpha } = parsed;
  // OPEN a bull call spread:
  // BUY the lower strike call
  // SELL the upper strike call
  const lower = getOptionQuote_(sym, exp, lo, "Call"); // you BUY this to open
  const upper = getOptionQuote_(sym, exp, hi, "Call"); // you SELL this to open
  if (!hasBidAsk_(lower)) return "#No Data for Lower:" + lowerStrike;
  if (!hasBidAsk_(upper)) return "#No Data for Upper:" + upperStrike;
  // At alpha=0:
  // buyLimit = lower.ask
  // sellLimit = upper.bid
  //
  // As alpha increases (more patience):
  // buyLimit moves Ask -> Bid (you try to pay less)
  // sellLimit moves Bid -> Ask (you try to receive more)
  const buyLimit = lower.ask - alpha * (lower.ask - lower.bid); // Ask -> Bid
  const sellLimit = upper.bid + alpha * (upper.ask - upper.bid); // Bid -> Ask
  let debit = buyLimit - sellLimit;
  if (debit < 0) debit = 0;
  return round2_(debit);
}

/**
 * recommendBullCallSpreadCloseCredit - Recommends a credit limit price to CLOSE a bull call spread.
 *
 * Strategy: SELL call at lowerStrike (close long), BUY call at upperStrike (close short).
 *
 * The price is the net credit to receive, adjusted for execution patience.
 * At alpha=0: credit = lower.bid - upper.ask (aggressive).
 * As alpha increases: Moves toward lower.ask - upper.bid (patient, better credit).
 * Minimum credit is 0.
 *
 * Usage Examples:
 *
 * 1. Spreadsheet Formula:
 *    =recommendBullCallSpreadCloseCredit("TSLA", "2028-06-16", 450, 500, 30)
 *    // Returns e.g., 2.75 (net credit per share)
 *
 * 2. From Script:
 *    const credit = recommendBullCallSpreadCloseCredit("TSLA", new Date("2028-06-16"), 450, 500, 60);
 *    // credit is a number or error string
 *
 * @param {string} symbol - The stock symbol (e.g., "TSLA"). Case-insensitive, trimmed.
 * @param {string|Date} expiration - Expiration date as "YYYY-MM-DD" string or Date object.
 * @param {number} lowerStrike - Lower strike price (long call to close).
 * @param {number} upperStrike - Upper strike price (short call to close, must be > lowerStrike).
 * @param {number} avgMinutesToExecute - Minutes of patience (0+); higher improves price but delays fill.
 * @returns {number|string} Recommended credit (rounded to 2 decimals) or "#Error message".
 */

function recommendBullCallSpreadCloseCredit(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute) {
  const parsed = parseSpreadInputs_(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute);
  if (!parsed) return "#Could not parse inputs";
  const { sym, exp, lo, hi, alpha } = parsed;
  const lower = getOptionQuote_(sym, exp, lo, "Call"); // you SELL this to close
  const upper = getOptionQuote_(sym, exp, hi, "Call"); // you BUY this to close
  if (!hasBidAsk_(lower)) return "#No Data for Lower:" + lowerStrike;
  if (!hasBidAsk_(upper)) return "#No Data for Upper:" + upperStrike;
  // At alpha=0:
  // sellLimit = lower.bid
  // buyLimit = upper.ask
  const sellLimit = lower.bid + alpha * (lower.ask - lower.bid); // Bid -> Ask
  const buyLimit = upper.ask - alpha * (upper.ask - upper.bid); // Ask -> Bid
  let credit = sellLimit - buyLimit;
  if (credit < 0) credit = 0;
  return round2_(credit);
}

/**
 * parseSpreadInputs_ - Parses and validates inputs for bull call spread functions.
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

function parseSpreadInputs_(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute) {
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
    ["bid", "mid", "ask"],
    SHEET_NAME
  );
  const row = res && res[0];
  if (!row) return null;
  const quote = {
    bid: row[0] === "" ? null : row[0],
    mid: row[1] === "" ? null : row[1],
    ask: row[2] === "" ? null : row[2]
  };
  cache.put(cacheKey, JSON.stringify(quote), 300); // Cache for 5 min
  return quote;
}

/**
 * round2_ - Rounds a number to 2 decimal places.
 *
 * Uses Math.round for banker’s rounding.
 *
 * @param {number} n - Number to round.
 * @returns {number} Rounded to 2 decimals.
 */

function round2_(n) {
  return Math.round(n * 100) / 100;
}

/**
 * BCS_OPEN_DEBIT_XX - Alias for recommendBullCallSpreadOpenDebit.
 *
 * For backward compatibility or shorthand in formulas.
 *
 * @param {string} symbol - Stock symbol.
 * @param {string|Date} expiration - Expiration.
 * @param {number} lowerStrike - Lower strike.
 * @param {number} upperStrike - Upper strike.
 * @param {number} avgMinutes - Patience in minutes.
 * @returns {number|string} Debit or error.
 */

function BCS_OPEN_DEBIT_XX(symbol, expiration, lowerStrike, upperStrike, avgMinutes) {
  return recommendBullCallSpreadOpenDebit(symbol, expiration, lowerStrike, upperStrike, avgMinutes);
}

/**
 * DEBUG_BCS_OPEN_DEBIT_XX - Debug version of open debit for bull call spread.
 *
 * Returns detailed string with alpha, quotes, and computed debit.
 * Useful for troubleshooting.
 *
 * @param {string} symbol - Stock symbol.
 * @param {string|Date} expiration - Expiration.
 * @param {number} lowerStrike - Lower strike.
 * @param {number} upperStrike - Upper strike.
 * @param {number} avgMinutesToExecute - Patience in minutes.
 * @returns {string} Debug info or error message.
 */

function DEBUG_BCS_OPEN_DEBIT_XX(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute) {
  const parsed = parseSpreadInputs_(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute);
  if (!parsed) return "parseSpreadInputs_ failed (symbol/expiration/strikes/minutes)";
  const { sym, exp, lo, hi, alpha } = parsed;
  const longQ = getOptionQuote_(sym, exp, lo, "Call");
  const shortQ = getOptionQuote_(sym, exp, hi, "Call");
  if (!longQ) return `Missing long quote: ${sym} ${exp} ${lo} Call`;
  if (!shortQ) return `Missing short quote: ${sym} ${exp} ${hi} Call`;
  if (!hasBidAsk_(longQ)) return `Long missing bid/ask: bid=${longQ.bid} ask=${longQ.ask}`;
  if (!hasBidAsk_(shortQ)) return `Short missing bid/ask: bid=${shortQ.bid} ask=${shortQ.ask}`;
  const buyLimit = longQ.ask - alpha * (longQ.ask - longQ.bid);
  const sellLimit = shortQ.bid + alpha * (shortQ.ask - shortQ.bid);
  let debit = buyLimit - sellLimit;
  if (debit < 0) debit = 0;
  return `OK debit=${round2_(debit)} (alpha=${alpha.toFixed(2)}) long[${longQ.bid}/${longQ.ask}] short[${shortQ.bid}/${shortQ.ask}]`;
}

/**
 * recommendIronCondorOpenCredit - Recommends a credit limit price to OPEN a short iron condor.
 *
 * Strategy (short IC for credit):
 * - BUY put at buyPut (protective, lowest strike).
 * - SELL put at sellPut (short, higher put strike).
 * - SELL call at sellCall (short, lower call strike).
 * - BUY call at buyCall (protective, highest strike).
 *
 * Validates strikes: buyPut < sellPut and sellCall < buyCall.
 * Net credit = (sellPut + sellCall) - (buyPut + buyCall), adjusted for alpha.
 * At alpha=0: Smaller credit (aggressive fills).
 * Higher alpha: Larger credit (patient).
 * Minimum credit is 0.
 *
 * Usage Examples:
 *
 * 1. Spreadsheet Formula:
 *    =recommendIronCondorOpenCredit("TSLA", "2028-06-16", 400, 420, 480, 500, 30)
 *    // Returns e.g., 3.50 (net credit per share)
 *
 * 2. From Script:
 *    const credit = recommendIronCondorOpenCredit("TSLA", new Date("2028-06-16"), 400, 420, 480, 500, 60);
 *
 * @param {string} symbol - Stock symbol (e.g., "TSLA").
 * @param {string|Date} expiration - Expiration date.
 * @param {number} buyPut - Lowest put strike (buy protective).
 * @param {number} sellPut - Higher put strike (sell short).
 * @param {number} sellCall - Lower call strike (sell short).
 * @param {number} buyCall - Highest call strike (buy protective).
 * @param {number} avgMinutesToExecute - Patience in minutes (0+).
 * @returns {number|string} Recommended credit or "#Error message".
 */

function recommendIronCondorOpenCredit(
  symbol,
  expiration,
  buyPut,
  sellPut,
  sellCall,
  buyCall,
  avgMinutesToExecute
) {
  const parsed = parseIronCondorInputs_(
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
  // BUY legs: Ask -> Bid as alpha increases (try to pay less)
  const buyPutLimit = qBuyPut.ask - alpha * (qBuyPut.ask - qBuyPut.bid);
  const buyCallLimit = qBuyCall.ask - alpha * (qBuyCall.ask - qBuyCall.bid);
  // SELL legs: Bid -> Ask as alpha increases (try to receive more)
  const sellPutLimit = qSellPut.bid + alpha * (qSellPut.ask - qSellPut.bid);
  const sellCallLimit = qSellCall.bid + alpha * (qSellCall.ask - qSellCall.bid);
  // Net credit to open
  let credit = (sellPutLimit + sellCallLimit) - (buyPutLimit + buyCallLimit);
  if (credit < 0) credit = 0;
  return round2_(credit);
}

/**
 * recommendIronCondorCloseDebit - Recommends a debit limit price to CLOSE a short iron condor.
 *
 * Assumes original open: BUY put@buyPut, SELL put@sellPut, SELL call@sellCall, BUY call@buyCall.
 * Close: SELL put@buyPut (close long), BUY put@sellPut (close short), BUY call@sellCall (close short), SELL call@buyCall (close long).
 *
 * Net debit = (buy shorts) - (sell longs), adjusted for alpha.
 * At alpha=0: Larger debit (aggressive).
 * Higher alpha: Smaller debit (patient).
 * Minimum debit is 0.
 *
 * Usage Examples:
 *
 * 1. Spreadsheet Formula:
 *    =recommendIronCondorCloseDebit("TSLA", "2028-06-16", 400, 420, 480, 500, 30)
 *    // Returns e.g., 1.25 (net debit per share)
 *
 * 2. From Script:
 *    const debit = recommendIronCondorCloseDebit("TSLA", new Date("2028-06-16"), 400, 420, 480, 500, 60);
 *
 * @param {string} symbol - Stock symbol.
 * @param {string|Date} expiration - Expiration date.
 * @param {number} buyPut - Original buy put strike (now sell to close).
 * @param {number} sellPut - Original sell put strike (now buy to close).
 * @param {number} buyCall - Original buy call strike (now sell to close).
 * @param {number} sellCall - Original sell call strike (now buy to close).
 * @param {number} avgMinutesToExecute - Patience in minutes (0+).
 * @returns {number|string} Recommended debit or "#Error message".
 */

function recommendIronCondorCloseDebit(
  symbol,
  expiration,
  buyPut,
  sellPut,
  buyCall,
  sellCall,
  avgMinutesToExecute
) {
  const parsed = parseIronCondorInputs_(
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
  // SELL-to-close legs: Bid -> Ask as alpha increases (try to receive more)
  const sellLongPutLimit = qLongPut.bid + alpha * (qLongPut.ask - qLongPut.bid);
  const sellLongCallLimit = qLongCall.bid + alpha * (qLongCall.ask - qLongCall.bid);
  // BUY-to-close legs: Ask -> Bid as alpha increases (try to pay less)
  const buyShortPutLimit = qShortPut.ask - alpha * (qShortPut.ask - qShortPut.bid);
  const buyShortCallLimit = qShortCall.ask - alpha * (qShortCall.ask - qShortCall.bid);
  // Net debit to close
  let debit = (buyShortPutLimit + buyShortCallLimit) - (sellLongPutLimit + sellLongCallLimit);
  if (debit < 0) debit = 0;
  return round2_(debit);
}

/**
 * parseIronCondorInputs_ - Parses and validates inputs for iron condor functions.
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

function parseIronCondorInputs_(
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
