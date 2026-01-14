/*******************************************************
 * OPTIONSTRAT SPREAD URL BUILDER
 *
 * Usage in Google Sheets:
 *   =buildOptionStratSpreadUrl(I7, "TSLA", "bull-call-spread", "expiration")
 *
 * Strike cell may contain:
 *   "450/460"
 *   "450/460 and other stuff"
 *
 * Strategy must be supported (see allow-list).
 * Option type (C/P) is inferred from strategy.
 *******************************************************/

// =buildOptionStratUrl(B4,"TSLA","bull-call-spread", C4)

function testbuildOptionStratUrl () {
  
  buildOptionStratUrl("350/450","TSLA", "bull-call-spread", new Date(2028, 5, 16));
  buildOptionStratUrl("350/450","TSLA", "bull-call-spread", 'Jun \'28');
  buildOptionStratUrl("350/450","TSLA", "bull-call-spread", 'Jun 17 2028');
  buildOptionStratUrl("350/450","TSLA", "bull-call-spread", 'Jun 2028');
  buildOptionStratUrl("350/450","TSLA", "bull-call-spread", 'Jan 16 2026');
  //buildOptionStratUrl(undefined,"TSLA", "bull-call-spread", 'Jan 16 2026');
}

function buildOptionStratUrlX(strikes, ticker, strategy, expiration) {

  const months = {
    Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6,
    Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12
  };


  function parseExpirationToDate(exp) {

    // Already a Date → use it
    if (exp instanceof Date) {
      if (isNaN(exp.getTime())) throw new Error("Invalid Date expiration");
      return exp;
    }

    let s = String(exp).trim();

    // Normalize "'28" → "2028"
    s = s.replace(/'(\d{2})\b/, "20$1");

    // Let JS parse it
    let d = new Date(s);
    if (isNaN(d.getTime())) {
      throw new Error("Invalid expiration format: " + exp);
    }

    // If no day was specified → assume monthly → 3rd Friday
    if (!/\b\d{1,2}\b/.test(s)) {
      d.setDate(getThirdFriday(d.getFullYear(), d.getMonth()));
    }

    // Normalize time (avoid DST edge cases)
    d.setHours(12, 0, 0, 0);

    return d;
  }

  function getThirdFriday(year, monthIndex0) {
    const d = new Date(year, monthIndex0, 1);

    // Move forward to first Friday
    while (d.getDay() !== 5) { // 5 = Friday
      d.setDate(d.getDate() + 1);
    }

    // Add two weeks → third Friday
    d.setDate(d.getDate() + 14);

    return d.getDate();
  }

  function formatDateCode(date) {
    return (
      String(date.getFullYear() % 100).padStart(2, "0") +
      String(date.getMonth() + 1).padStart(2, "0") +
      String(date.getDate()).padStart(2, "0")
    );
  }

  function validateSupportedStrategy(strategy) {
    const supported = new Set([
      "bull-call-spread",
      "bear-call-spread",
      "bull-put-spread",
      "bear-put-spread"
    ]);

    if (!supported.has(strategy)) {
      throw new Error(
        `Unsupported strategy "${strategy}". Supported: ${Array.from(supported).join(", ")}`
      );
    }
  }

  /*******************************************************
   * STRIKE PARSING (STRICT ORDER, FLEXIBLE TEXT)
   *******************************************************/

  function parseStrikePairStrict(strikePair) {
    const text = String(strikePair);

    // Extract first number/number anywhere in string
    const match = text.match(/(\d+(?:\.\d+)?)\s*\/\s*(\d+(?:\.\d+)?)/);
    if (!match) {
      throw new Error(`Strike must contain a pair like "450/460". Got: "${strikePair}"`);
    }

    const lower = Number(match[1]);
    const upper = Number(match[2]);

    if (!Number.isFinite(lower) || !Number.isFinite(upper)) {
      throw new Error(`Strikes must be numeric: "${strikePair}"`);
    }

    if (lower >= upper) {
      throw new Error(`Invalid strike order: ${lower} must be < ${upper}`);
    }
    return [match[1], match[2]];
  }

  if (!(strikes && ticker && strategy && expiration) ) {
    throw new Error("undefined parameter");
  }
  
  validateSupportedStrategy(strategy);
  

  // const { month, day: specifiedDay, year } = parseExpiration(expiration);
  // const day = specifiedDay !== null ? specifiedDay : getThirdFriday(year, month);

  const expDate = parseExpirationToDate(expiration);
  const dateCode = formatDateCode(expDate);

  //const dateCode = formatDateCode(year, month, day);

  const callOrPutChar = strategy.toLowerCase().includes("call") ? "C" : "P";

  //const [lowStrike, highStrike] = strikes.split('/');
  const [lowStrike, highStrike] = parseStrikePairStrict(strikes);
  const symbolLow = `.${ticker}${dateCode}${callOrPutChar}${lowStrike}`;
  const symbolHigh = `-.${ticker}${dateCode}${callOrPutChar}${highStrike}`;
 
  const url = `https://optionstrat.com/build/${strategy}/${ticker}/${symbolLow},${symbolHigh}`;
  console.log("optionstrat URL:" + url);
  return url;
}



/**
 * X3LOOKUP(expiration, strike, ExpCol, StrikeCol, ReturnCol)
 * Two-key lookup for Google Sheets custom functions.
 *
 * Returns:
 *  - value from ReturnCol when found
 *  - "NA" if not found
 *  - "REF" if ranges are mismatched/missing
 *  - "VALUE" if wrong arg count
 *
 * @customfunction
 */
function X3LOOKUP(key1, key2, col1, col2, returnCol) {
  // Helper: convert Sheets range (2D) or scalar into 1D array
  function to1D(x) {
    if (!Array.isArray(x)) return [x];
    const out = new Array(x.length);
    for (let i = 0; i < x.length; i++) {
      const row = x[i];
      out[i] = Array.isArray(row) ? row[0] : row;
    }
    return out;
  }

  // Helper: normalize Date -> midnight local time (ignore time portion)
  function dateSerial(d) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  }

  // Argument / range checks
  if (arguments.length !== 5) return "#VALUE";
  if (col1 == null || col2 == null || returnCol == null) return "#REF";

  const c1 = to1D(col1);
  const c2 = to1D(col2);
  const r  = to1D(returnCol);

  if (c1.length !== c2.length || c1.length !== r.length) return "#REF";

  const k1IsDate = key1 instanceof Date;
  const k1 = k1IsDate ? dateSerial(key1) : key1;
  const k2 = Number(key2);

  for (let i = 0; i < c1.length; i++) {
    let v1 = c1[i];
    let v2 = c2[i];

    if (k1IsDate && v1 instanceof Date) v1 = dateSerial(v1);

    // Strike compare as number (handles "350" vs 350)
    if (v1 === k1 && Number(v2) === k2) {
      return r[i];
    }
  }

  return "#NA";
}


/*******************************************************
 * EXPECTED-FILL PRICING MODEL (BULL CALL SPREAD)
 *
 * Two public custom functions:
 *  - recommendedDebitPrice(...)  // BUY bull call spread (net debit)
 *  - recommendedCreditPrice(...) // SELL bull call spread (net credit)
 *
 * Uses probabilityOfFilling + timeHorizonMinutes to bias
 * toward bid-side (buy) / ask-side (sell) expected fills.
 *******************************************************/

function normalizeQuotes_(lowerBid, lowerAsk, upperBid, upperAsk, args) {
  // Propagate sheet-style errors
  const err = args.find(v => typeof v === "string" && v.startsWith("#"));
  if (err) return { ok: false, value: err };

  // Missing inputs => blank
  if ([lowerBid, lowerAsk, upperBid, upperAsk].some(v => v === "" || v == null)) {
    return { ok: false, value: "" };
  }

  const lb = Number(lowerBid);
  const la = Number(lowerAsk);
  const ub = Number(upperBid);
  const ua = Number(upperAsk);

  if (![lb, la, ub, ua].every(Number.isFinite)) return { ok: false, value: "#VALUE!" };

  // Basic sanity: bid <= ask
  if (lb > la || ub > ua) return { ok: false, value: "#VALUE!" };

  return { ok: true, lb, la, ub, ua };
}

function normalizeFillParams_(probabilityOfFilling, timeHorizonMinutes) {
  // Allow blanks to propagate blank (common in sheets)
  if (probabilityOfFilling === "" || probabilityOfFilling == null ||
      timeHorizonMinutes === "" || timeHorizonMinutes == null) {
    return { ok: false, value: "" };
  }

  const p = Number(probabilityOfFilling);
  const m = Number(timeHorizonMinutes);

  if (!Number.isFinite(p) || !Number.isFinite(m)) return { ok: false, value: "#VALUE!" };

  // clamp p to [0,1], clamp minutes to >=0
  const pc = clamp_(p, 0, 1);
  const mc = Math.max(0, m);

  return { ok: true, p: pc, minutes: mc };
}

/**
 * Core: compute concession weight w in [0,1]
 * w=0 => best-for-you pricing (harder fill)
 * w=1 => worst-for-you pricing (easier fill)
 *
 * We shape it with gammaFromMinutes_ so longer horizon => smaller w
 * for the same target probability.
 */
function concessionWeight_(p, minutes) {
  const g = gammaFromMinutes_(minutes);        // 1.35 .. 3.0
  return clamp_(Math.pow(p, g), 0, 1);
}

/**
 * Linearly interpolate between bid and ask by weight w.
 * w=0 -> bid, w=1 -> ask
 */
function lerpBidAsk_(bid, ask, w) {
  return bid + w * (ask - bid);
}

/**
 * Linearly interpolate between ask and bid by weight w.
 * w=0 -> ask, w=1 -> bid
 */
function lerpAskBid_(ask, bid, w) {
  return ask - w * (ask - bid);
}

/**
 * Shared pricing core for bull call spread.
 *
 * side = "BUY"  => recommend net DEBIT (what you pay)
 * side = "SELL" => recommend net CREDIT (what you receive)
 */
function recommendBullCallSpreadPrice_(side, lb, la, ub, ua, p, minutes, tick) {
  const w = concessionWeight_(p, minutes);

  // BUY spread: buy lower, sell upper
  if (side === "BUY") {
    // buy lower: bid -> ask as w increases
    const buyLower = lerpBidAsk_(lb, la, w);
    // sell upper: ask -> bid as w increases
    const sellUpper = lerpAskBid_(ua, ub, w);
    const debit = buyLower - sellUpper;
    return roundToTick_(debit, tick);
  }

  // SELL spread: sell lower, buy upper (closing / writing spread)
  if (side === "SELL") {
    // sell lower: ask -> bid as w increases
    const sellLower = lerpAskBid_(la, lb, w);
    // buy upper: bid -> ask as w increases
    const buyUpper = lerpBidAsk_(ub, ua, w);
    const credit = sellLower - buyUpper;
    return roundToTick_(credit, tick);
  }

  return "#VALUE!";
}

/**
 * @customfunction
 * Recommend a DEBIT to BUY a bull call spread.
 */
function recommendedDebitPrice(
  lowerBid, lowerAsk,
  upperBid, upperAsk,
  probabilityOfFilling,
  timeHorizonMinutes,
  debug
) {
  const args = [lowerBid, lowerAsk, upperBid, upperAsk, probabilityOfFilling, timeHorizonMinutes, debug];
  if (debug === true) return "DEBUG: called with " + JSON.stringify(args);

  const q = normalizeQuotes_(lowerBid, lowerAsk, upperBid, upperAsk, args);
  if (!q.ok) return q.value;

  const fp = normalizeFillParams_(probabilityOfFilling, timeHorizonMinutes);
  if (!fp.ok) return fp.value;

  // tick size: 0.01 for options in most cases
  return recommendBullCallSpreadPrice_("BUY", q.lb, q.la, q.ub, q.ua, fp.p, fp.minutes, 0.01);
}

/**
 * @customfunction
 * Recommend a CREDIT to SELL a bull call spread.
 */
function recommendedCreditPrice(
  lowerBid, lowerAsk,
  upperBid, upperAsk,
  probabilityOfFilling,
  timeHorizonMinutes,
  debug
) {
  const args = [lowerBid, lowerAsk, upperBid, upperAsk, probabilityOfFilling, timeHorizonMinutes, debug];
  if (debug === true) return "DEBUG: called with " + JSON.stringify(args);

  const q = normalizeQuotes_(lowerBid, lowerAsk, upperBid, upperAsk, args);
  if (!q.ok) return q.value;

  const fp = normalizeFillParams_(probabilityOfFilling, timeHorizonMinutes);
  if (!fp.ok) return fp.value;

  return recommendBullCallSpreadPrice_("SELL", q.lb, q.la, q.ub, q.ua, fp.p, fp.minutes, 0.01);
}

/*******************************************************
 * YOUR EXISTING HELPERS (unchanged)
 *******************************************************/

function gammaFromMinutes_(minutes) {
  const gMin = 1.35;
  const gMax = 3.0;
  const x = Math.log(1 + minutes) / Math.log(31);
  return gMin + (gMax - gMin) * clamp_(x, 0, 1);
}

function clamp_(x, lo, hi) {
  return Math.min(hi, Math.max(lo, x));
}

function roundToTick_(x, tick) {
  return Math.round(x / tick) * tick;
}

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
}

