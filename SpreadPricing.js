/*******************************************************
 * EXPECTED-FILL PRICING MODEL (BULL CALL SPREAD)
 *
 * Public custom functions:
 *  - recommendedBullCallDebitPrice(...)  // BUY bull call spread (net debit)
 *  - recommendedBullCallCreditPrice(...) // SELL bull call spread (net credit)
 *
 * Defaults (if omitted in Sheets):
 *  - probabilityOfFilling = 0.20
 *  - timeHorizonMinutes  = 60
 *  - debug               = false
 *******************************************************/

/**
 * @customfunction
 */
function recommendedBullCallDebitPrice(
  lowerBid, lowerAsk,
  upperBid, upperAsk,
  probabilityOfFilling,
  timeHorizonMinutes,
  debug
) {
  const args = [
    lowerBid, lowerAsk,
    upperBid, upperAsk,
    probabilityOfFilling,
    timeHorizonMinutes,
    debug
  ];

  const d = withDefaults_(probabilityOfFilling, timeHorizonMinutes, debug);
  if (d.debug) return "DEBUG: called with " + JSON.stringify(args);

  const q = normalizeQuotes_(lowerBid, lowerAsk, upperBid, upperAsk, args);
  if (!q.ok) return q.value;

  if (!Number.isFinite(d.p) || !Number.isFinite(d.minutes)) return "#VALUE!";

  return recommendBullCallSpreadPrice_(
    "BUY",
    q.lb, q.la,
    q.ub, q.ua,
    clamp_(d.p, 0, 1),
    Math.max(0, d.minutes),
    0.01
  );
}

/**
 * @customfunction
 */
function recommendedBullCallCreditPriceX(
  lowerBid, lowerAsk,
  upperBid, upperAsk,
  probabilityOfFilling,
  timeHorizonMinutes,
  debug
) {
  const args = [
    lowerBid, lowerAsk,
    upperBid, upperAsk,
    probabilityOfFilling,
    timeHorizonMinutes,
    debug
  ];

  const d = withDefaults_(probabilityOfFilling, timeHorizonMinutes, debug);
  if (d.debug) return "DEBUG: called with " + JSON.stringify(args);

  const q = normalizeQuotes_(lowerBid, lowerAsk, upperBid, upperAsk, args);
  if (!q.ok) return q.value;

  if (!Number.isFinite(d.p) || !Number.isFinite(d.minutes)) return "#VALUE!";

  return recommendBullCallSpreadPrice_(
    "SELL",
    q.lb, q.la,
    q.ub, q.ua,
    clamp_(d.p, 0, 1),
    Math.max(0, d.minutes),
    0.01
  );
}

/*******************************************************
 * CORE PRICING LOGIC
 *******************************************************/

/**
 * Common core for bull call spread expected-fill pricing.
 *
 * side = "BUY"  => recommend net DEBIT (what you pay)
 * side = "SELL" => recommend net CREDIT (what you receive)
 */
function recommendBullCallSpreadPrice_(side, lb, la, ub, ua, p, minutes, tick) {
  const w = concessionWeight_(p, minutes);

  if (side === "BUY") {
    // Best-for-you if patient: buy lower at bid, sell upper at ask
    // Worst-for-you if urgent: buy lower at ask, sell upper at bid
    const buyLower  = lerpBidAsk_(lb, la, w);   // bid -> ask
    const sellUpper = lerpAskBid_(ua, ub, w);   // ask -> bid
    const debit = buyLower - sellUpper;
    return roundToTick_(debit, tick);
  }

  if (side === "SELL") {
    // Best-for-you if patient: sell lower at ask, buy upper at bid
    // Worst-for-you if urgent: sell lower at bid, buy upper at ask
    const sellLower = lerpAskBid_(la, lb, w);   // ask -> bid
    const buyUpper  = lerpBidAsk_(ub, ua, w);   // bid -> ask
    const credit = sellLower - buyUpper;
    return roundToTick_(credit, tick);
  }

  return "#VALUE!";
}

/**
 * Concession weight w in [0,1]
 *  w=0 => best-for-you pricing (harder fill)
 *  w=1 => worst-for-you pricing (easier fill)
 */
function concessionWeight_(p, minutes) {
  const g = gammaFromMinutes_(minutes); // 1.35 .. 3.0
  return clamp_(Math.pow(clamp_(p, 0, 1), g), 0, 1);
}

/**
 * w=0 -> bid, w=1 -> ask
 */
function lerpBidAsk_(bid, ask, w) {
  return bid + w * (ask - bid);
}

/**
 * w=0 -> ask, w=1 -> bid
 */
function lerpAskBid_(ask, bid, w) {
  return ask - w * (ask - bid);
}

/*******************************************************
 * INPUT NORMALIZATION / DEFAULTS
 *******************************************************/

function withDefaults_(pFill, minutes, debug) {
  const pRaw = (pFill === "" || pFill == null) ? 0.20 : Number(pFill);
  const mRaw = (minutes === "" || minutes == null) ? 60 : Number(minutes);
  return {
    p: pRaw,
    minutes: mRaw,
    debug: debug === true
  };
}

/**
 * Validate/coerce the four quote inputs and propagate sheet errors.
 * Returns either:
 *   { ok: true, lb, la, ub, ua }
 * or
 *   { ok: false, value: "" | "#VALUE!" | "#N/A" | "#REF!" | ... }
 */
function normalizeQuotes_(lowerBid, lowerAsk, upperBid, upperAsk, args) {
  // 1) Propagate Sheets-style errors from X3LOOKUP or elsewhere
  const err = args.find(v => typeof v === "string" && v.startsWith("#"));
  if (err) return { ok: false, value: err };

  // 2) Missing inputs → blank
  if ([lowerBid, lowerAsk, upperBid, upperAsk].some(v => v === "" || v == null)) {
    return { ok: false, value: "" };
  }

  // 3) Coerce to numbers
  const lb = Number(lowerBid);
  const la = Number(lowerAsk);
  const ub = Number(upperBid);
  const ua = Number(upperAsk);

  // 4) Invalid numeric inputs → VALUE error
  if (![lb, la, ub, ua].every(Number.isFinite)) {
    return { ok: false, value: "#VALUE!" };
  }

  // 5) Basic sanity: bid <= ask
  if (lb > la || ub > ua) {
    return { ok: false, value: "#VALUE!" };
  }

  return { ok: true, lb, la, ub, ua };
}

/*******************************************************
 * YOUR EXISTING HELPERS (kept)
 *******************************************************/

function gammaFromMinutes_(minutes) {
  const gMin = 1.35;
  const gMax = 3.0;
  const x = Math.log(1 + Math.max(0, minutes)) / Math.log(31);
  return gMin + (gMax - gMin) * clamp_(x, 0, 1);
}

function clamp_(x, lo, hi) {
  return Math.min(hi, Math.max(lo, x));
}

function roundToTick_(x, tick) {
  return Math.round(x / tick) * tick;
}


/*******************************************************
 * TEST CASES
 *******************************************************/

function test_BullCallSpreadPricing() {

  // --- Jun 16 2028 (350 / 550) ---
  const junLowerBid = 203.15;
  const junLowerAsk = 207.05;
  const junUpperBid = 133.10;
  const junUpperAsk = 136.15;

  // --- Dec 15 2028 (350 / 550) ---
  const decLowerBid = 214.00;
  const decLowerAsk = 221.90;
  const decUpperBid = 148.00;
  const decUpperAsk = 157.00;

  const pPatient = 0.25;
  const pUrgent  = 0.85;
  const minutes  = 30;

  console.log("=== JUN 2028 350/550 ===");
  console.log("BUY (patient): ",
    recommendedDebitPrice(
      junLowerBid, junLowerAsk,
      junUpperBid, junUpperAsk,
      pPatient, minutes, false
    )
  );
  console.log("BUY (urgent): ",
    recommendedDebitPrice(
      junLowerBid, junLowerAsk,
      junUpperBid, junUpperAsk,
      pUrgent, minutes, false
    )
  );
  console.log("SELL (patient): ",
    recommendedCreditPrice(
      junLowerBid, junLowerAsk,
      junUpperBid, junUpperAsk,
      pPatient, minutes, false
    )
  );
  console.log("SELL (urgent): ",
    recommendedCreditPrice(
      junLowerBid, junLowerAsk,
      junUpperBid, junUpperAsk,
      pUrgent, minutes, false
    )
  );

  console.log("=== DEC 2028 350/550 ===");
  console.log("BUY (patient): ",
    recommendedDebitPrice(
      decLowerBid, decLowerAsk,
      decUpperBid, decUpperAsk,
      pPatient, minutes, false
    )
  );
  console.log("BUY (urgent): ",
    recommendedDebitPrice(
      decLowerBid, decLowerAsk,
      decUpperBid, decUpperAsk,
      pUrgent, minutes, false
    )
  );
  console.log("SELL (patient): ",
    recommendedCreditPrice(
      decLowerBid, decLowerAsk,
      decUpperBid, decUpperAsk,
      pPatient, minutes, false
    )
  );
  console.log("SELL (urgent): ",
    recommendedCreditPrice(
      decLowerBid, decLowerAsk,
      decUpperBid, decUpperAsk,
      pUrgent, minutes, false
    )
  );

  console.log("✅ test_BullCallSpreadPricing PASSED");

}


// MORE

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
function suggestDebit(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute) {
  const parsed = parseSpreadInputs_(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute);
  if (!parsed) return "";
  const { sym, exp, lo, hi, alpha } = parsed;

  const longQ  = getOptionQuote_(sym, exp, lo, "Call"); // buy
  const shortQ = getOptionQuote_(sym, exp, hi, "Call"); // sell
  if (!hasBidAsk_(longQ) || !hasBidAsk_(shortQ)) return "";

  const buyLimit  = longQ.ask  - alpha * (longQ.ask  - longQ.bid); // Ask -> Bid
  const sellLimit = shortQ.bid + alpha * (shortQ.ask - shortQ.bid); // Bid -> Ask

  let debitLimit = buyLimit - sellLimit;
  if (debitLimit < 0) debitLimit = 0;

  return round2_(debitLimit);
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
function suggestCredit(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute) {
  const parsed = parseSpreadInputs_(symbol, expiration, lowerStrike, upperStrike, avgMinutesToExecute);
  if (!parsed) return "";
  const { sym, exp, lo, hi, alpha } = parsed;

  const longQ  = getOptionQuote_(sym, exp, lo, "Call"); // you SELL to close
  const shortQ = getOptionQuote_(sym, exp, hi, "Call"); // you BUY to close
  if (!hasBidAsk_(longQ) || !hasBidAsk_(shortQ)) return "";

  const sellLimit = longQ.bid + alpha * (longQ.ask - longQ.bid);   // Bid -> Ask
  const buyLimit  = shortQ.ask - alpha * (shortQ.ask - shortQ.bid); // Ask -> Bid

  let creditLimit = sellLimit - buyLimit;
  if (creditLimit < 0) creditLimit = 0;

  return round2_(creditLimit);
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

