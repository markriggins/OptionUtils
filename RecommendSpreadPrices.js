/**
 * Bull Call Spread pricing helpers (OPEN debit / CLOSE credit)
 *
 * Uses XLookupByKeys (generic + cached) against sheet "OptionPricesUploaded"
 * Expected headers (lowercase recommended):
 *   symbol | expiration | strike | type | bid | mid | ask
 *
 * Expiration accepted as:
 *   - "YYYY-MM-DD" string
 *   - Date object (from Sheets cell)  -> normalized to "YYYY-MM-DD"
 */

/**
 * recommendBullCallSpreadOpenDebit
 *
 * Debit limit to OPEN bull call spread:
 *   BUY  Call @ lowerStrike
 *   SELL Call @ upperStrike
 *
 * avgMinutesToExecute:
 *   0      -> aggressive (Ask(long) - Bid(short))
 *   larger -> more patient (improve toward Bid(long) and Ask(short))
 *
 * Returns ONE number suitable for a single cell.
 */
function recommendBullCallSpreadOpenDebit(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute) {
  const parsed = parseSpreadInputs_(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute);
  if (!parsed) return "#Could not parse inputs";
  const { sym, exp, lo, hi, alpha } = parsed;

  // OPEN a bull call spread:
  //   BUY the lower strike call
  //   SELL the upper strike call
  const lower = getOptionQuote_(sym, exp, lo, "Call"); // you BUY this to open
  const upper = getOptionQuote_(sym, exp, hi, "Call"); // you SELL this to open

  if (!hasBidAsk_(lower)) return "#No Data for Lower:" + lowerStrike;
  if (!hasBidAsk_(upper)) return "#No Data for Upper:" + upperStrike;

  // At alpha=0:
  //   buyLimit  = lower.ask
  //   sellLimit = upper.bid
  //
  // As alpha increases (more patience):
  //   buyLimit moves Ask -> Bid (you try to pay less)
  //   sellLimit moves Bid -> Ask (you try to receive more)
  const buyLimit  = lower.ask - alpha * (lower.ask - lower.bid);     // Ask -> Bid
  const sellLimit = upper.bid + alpha * (upper.ask - upper.bid);     // Bid -> Ask

  let debit = buyLimit - sellLimit;
  if (debit < 0) debit = 0;
  return round2_(debit);
}


/**
 * recommendBullCallSpreadCloseCredit
 *
 * Credit limit to CLOSE bull call spread:
 *   SELL Call @ lowerStrike  (close the long)
 *   BUY  Call @ upperStrike  (close the short)
 *
 * avgMinutesToExecute:
 *   0      -> aggressive (Bid(sell long) - Ask(buy short))
 *   larger -> more patient (improve toward Ask(sell long) and Bid(buy short))
 *
 * Returns ONE number suitable for a single cell.
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
  //   sellLimit = lower.bid
  //   buyLimit  = upper.ask
  const sellLimit = lower.bid + alpha * (lower.ask - lower.bid);     // Bid -> Ask
  const buyLimit  = upper.ask - alpha * (upper.ask - upper.bid);     // Ask -> Bid

  let credit = sellLimit - buyLimit;
  if (credit < 0) credit = 0;
  return round2_(credit);
}

/** ===========================
 * Helpers (shared)
 * =========================== */

function parseSpreadInputs_(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute) {
  const sym = String(symbol).toUpperCase().trim();
  const exp = normalizeExpiration_(expiration);
  const lo = Number(lowerStrike);
  const hi = Number(upperStrike);
  let mins = Number(avgMinutesToExecute);

  if (!sym) return null;
  if (!exp) return null;
  if (!Number.isFinite(lo) || !Number.isFinite(hi)) return null;
  if (lo >= hi) return null;
  if (!Number.isFinite(mins) || mins < 0) mins = 0;

  // Aggressiveness curve (saturating)
  // mins=0    => alpha=0 (cross)
  // mins=60   => alpha≈0.63
  // mins=1440 => alpha≈1
  const HALF_LIFE_MIN = 60;
  const alpha = 1 - Math.exp(-mins / HALF_LIFE_MIN);

  return { sym, exp, lo, hi, alpha };
}

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

function hasBidAsk_(q) {
  return q && q.bid != null && q.ask != null;
}

/**
 * Uses your cached XLookupByKeys under the hood.
 * Requires OptionPricesUploaded headers: symbol|expiration|strike|type|bid|mid|ask
 */
function getOptionQuote_(symbol, expiration, strike, type) {
  const SHEET_NAME = "OptionPricesUploaded";

  const res = XLookupByKeys(
    [symbol, expiration, Number(strike), type],
    ["symbol", "expiration", "strike", "type"],
    ["bid", "mid", "ask"],
    SHEET_NAME
  );

  const row = res && res[0];
  if (!row) return null;

  return {
    bid: row[0] === "" ? null : row[0],
    mid: row[1] === "" ? null : row[1],
    ask: row[2] === "" ? null : row[2]
  };
}

function round2_(n) {
  return Math.round(n * 100) / 100;
}


function BCS_OPEN_DEBIT_XX(symbol, expiration, lowerStrike, upperStrike, avgMinutes) {
  return recommendBullCallSpreadOpenDebit(symbol, expiration, lowerStrike, upperStrike, avgMinutes);
}

function DEBUG_BCS_OPEN_DEBIT_XX(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute) {
  const parsed = parseSpreadInputs_(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute);
  if (!parsed) return "parseSpreadInputs_ failed (symbol/expiration/strikes/minutes)";

  const { sym, exp, lo, hi, alpha } = parsed;

  const longQ  = getOptionQuote_(sym, exp, lo, "Call");
  const shortQ = getOptionQuote_(sym, exp, hi, "Call");

  if (!longQ) return `Missing long quote: ${sym} ${exp} ${lo} Call`;
  if (!shortQ) return `Missing short quote: ${sym} ${exp} ${hi} Call`;

  if (!hasBidAsk_(longQ)) return `Long missing bid/ask: bid=${longQ.bid} ask=${longQ.ask}`;
  if (!hasBidAsk_(shortQ)) return `Short missing bid/ask: bid=${shortQ.bid} ask=${shortQ.ask}`;

  const buyLimit  = longQ.ask  - alpha * (longQ.ask  - longQ.bid);
  const sellLimit = shortQ.bid + alpha * (shortQ.ask - shortQ.bid);
  let debit = buyLimit - sellLimit;
  if (debit < 0) debit = 0;

  return `OK debit=${round2_(debit)} (alpha=${alpha.toFixed(2)}) long[${longQ.bid}/${longQ.ask}] short[${shortQ.bid}/${shortQ.ask}]`;
}

/**
 * recommendIronCondorOpenCredit
 *
 * Opens a standard SHORT iron condor for a NET CREDIT.
 * Legs assumed (typical short IC):
 *   BUY  put  @ buyPut     (protective put, lower strike)
 *   SELL put  @ sellPut    (short put, higher strike)
 *   SELL call @ sellCall   (short call, lower strike)
 *   BUY  call @ buyCall    (protective call, higher strike)
 *
 * avgMinutesToExecute controls aggressiveness:
 *   0   => worst-case immediate fills (buy at ask / sell at bid)  -> smaller credit
 *   >0  => moves toward better prices (buy cheaper / sell richer) -> larger credit
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
  const qBuyPut   = getOptionQuote_(sym, exp, bp, "Put");   // BUY
  const qSellPut  = getOptionQuote_(sym, exp, sp, "Put");   // SELL
  const qBuyCall  = getOptionQuote_(sym, exp, bc, "Call");  // BUY
  const qSellCall = getOptionQuote_(sym, exp, sc, "Call");  // SELL

  if (!hasBidAsk_(qBuyPut))   return "#No Data for Buy Put:" + sym + " " + exp + " @" + buyPut;
  if (!hasBidAsk_(qSellPut))  return "#No Data for Sell Put:" + sym + " " + exp + " @" + sellPut;
  if (!hasBidAsk_(qSellCall)) return "#No Data for Sell Call:" + sym + " " + exp + " @" + sellCall;
  if (!hasBidAsk_(qBuyCall))  return "#No Data for Buy Call:" + sym + " " + exp + " @" + buyCall;

  // BUY legs:  Ask -> Bid as alpha increases (try to pay less)
  const buyPutLimit  = qBuyPut.ask  - alpha * (qBuyPut.ask  - qBuyPut.bid);
  const buyCallLimit = qBuyCall.ask - alpha * (qBuyCall.ask - qBuyCall.bid);

  // SELL legs: Bid -> Ask as alpha increases (try to receive more)
  const sellPutLimit  = qSellPut.bid  + alpha * (qSellPut.ask  - qSellPut.bid);
  const sellCallLimit = qSellCall.bid + alpha * (qSellCall.ask - qSellCall.bid);

  // Net credit to open
  let credit = (sellPutLimit + sellCallLimit) - (buyPutLimit + buyCallLimit);
  if (credit < 0) credit = 0;

  return round2_(credit);
}

/**
 * recommendIronCondorCloseDebit
 *
 * Closes a standard SHORT iron condor by paying a NET DEBIT.
 * (This is the typical close for a short IC: you buy back the short legs and sell the longs.)
 *
 * Assumes the iron condor was opened as:
 *   BUY  put  @ buyPut     (protective put, lower strike)
 *   SELL put  @ sellPut    (short put, higher strike)
 *   SELL call @ sellCall   (short call, lower strike)
 *   BUY  call @ buyCall    (protective call, higher strike)
 *
 * To CLOSE (buy-to-close shorts, sell-to-close longs):
 *   SELL put  @ buyPut     (close long put)
 *   BUY  put  @ sellPut    (close short put)
 *   BUY  call @ sellCall   (close short call)
 *   SELL call @ buyCall    (close long call)
 *
 * avgMinutesToExecute controls aggressiveness:
 *   0   => worst-case immediate fills (buy at ask / sell at bid)  -> larger debit
 *   >0  => moves toward better prices (buy cheaper / sell richer) -> smaller debit
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
  const qLongPut   = getOptionQuote_(sym, exp, bp, "Put");   // was BUY to open -> SELL to close
  const qShortPut  = getOptionQuote_(sym, exp, sp, "Put");   // was SELL to open -> BUY to close
  const qLongCall  = getOptionQuote_(sym, exp, bc, "Call");  // was BUY to open -> SELL to close
  const qShortCall = getOptionQuote_(sym, exp, sc, "Call");  // was SELL to open -> BUY to close

  if (!hasBidAsk_(qLongPut))   return "#No Data for Buy Put:" + buyPut;
  if (!hasBidAsk_(qShortPut))  return "#No Data for Sell Put:" + sellPut;
  if (!hasBidAsk_(qLongCall))  return "#No Data for Buy Call:" + buyCall;
  if (!hasBidAsk_(qShortCall)) return "#No Data for Sell Call:" + sellCall;

  // SELL-to-close legs: Bid -> Ask as alpha increases (try to receive more)
  const sellLongPutLimit  = qLongPut.bid  + alpha * (qLongPut.ask  - qLongPut.bid);
  const sellLongCallLimit = qLongCall.bid + alpha * (qLongCall.ask - qLongCall.bid);

  // BUY-to-close legs: Ask -> Bid as alpha increases (try to pay less)
  const buyShortPutLimit  = qShortPut.ask  - alpha * (qShortPut.ask  - qShortPut.bid);
  const buyShortCallLimit = qShortCall.ask - alpha * (qShortCall.ask - qShortCall.bid);

  // Net debit to close
  let debit = (buyShortPutLimit + buyShortCallLimit) - (sellLongPutLimit + sellLongCallLimit);
  if (debit < 0) debit = 0;

  return round2_(debit);
}

/** -------- helper (self-contained) -------- */

function parseIronCondorInputs_(
  symbol,
  expiration,
  buyPut,
  sellPut,
  sellCall,
  buyCall,
  avgMinutesToExecute
) {
  const sym = String(symbol || "").trim().toUpperCase();
  if (!sym) return { error: "#Symbol required" };

  const exp = normalizeExpiration_(expiration);
  if (!exp) return { error: "#Bad expiration" };

  const bp = Number(buyPut);
  const sp = Number(sellPut);
  const bc = Number(buyCall);
  const sc = Number(sellCall);

  if (![bp, sp, bc, sc].every(Number.isFinite)) return { error: "#Bad strikes" };

  // Basic sanity for a typical short IC:
  // buyPut < sellPut and sellCall < buyCall
  if (!(bp < sp)) return { error: "#Put strikes must be buyPut < sellPut"+ pb + " !< " + sp };
  if (!(sc < bc)) return { error: "#Call strikes must be sellCall < buyCall" + sc + " !< " + bc };

  let mins = Number(avgMinutesToExecute);
  if (!Number.isFinite(mins) || mins < 0) mins = 0;

  // Same alpha curve concept as spreads (tune if desired)
  const HALF_LIFE_MIN = 60;
  const alpha = 1 - Math.exp(-mins / HALF_LIFE_MIN);

  return { sym, exp, bp, sp, sc, bc, alpha, error: null };
}

