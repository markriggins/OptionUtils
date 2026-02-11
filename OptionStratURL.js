/*******************************************************
 * Optionstrat.com  URL BUILDER
 *
 * Usage in Google Sheets:
 *   =buildOptionStratUrl("350/450","TSLA", "bull-call-spread", 'Jun 17 2028');
 *
 * Strike cell may contain:
 *   "450/460"
 *   "450/460 and other stuff"
 *
 * Strategy must be supported (see allow-list).
 * Option type (C/P) is inferred from strategy.
 *******************************************************/


function testbuildOptionStratUrl () {

  buildOptionStratUrl("350/450","TSLA", "bull-call-spread", new Date(2028, 5, 16));
  buildOptionStratUrl("350/450","TSLA", "bull-call-spread", 'Jun \'28');
  buildOptionStratUrl("350/450","TSLA", "bull-put-spread", 'Jun 17 2028');
  buildOptionStratUrl("350/450","TSLA", "bull-call-spread", 'Jun 2028');
  buildOptionStratUrl("350/450","TSLA", "bull-call-spread", 'Jan 16 2026');
  //buildOptionStratUrl(undefined,"TSLA", "bull-call-spread", 'Jan 16 2026');
  console.log("✅ testbuildOptionStratUrl PASSED");
}

/**
 * Builds an OptionStrat URL from multi-leg position data.
 *
 * @param {Range|string} symbolRange - Symbol or range to search upward for first non-blank
 * @param {Range} strikeRange - Strike column for the group
 * @param {Range} typeRange - Type column (Call/Put) for the group
 * @param {Range} expirationRange - Expiration column for the group
 * @param {Range} qtyRange - Qty column for the group (positive=long, negative=short)
 * @return {string} OptionStrat URL
 * @customfunction
 */
function buildOptionStratUrlFromLegs(symbolRange, strikeRange, typeRange, expirationRange, qtyRange) {
  // Find symbol (with upward lookup if range)
  let symbol = symbolRange;
  if (Array.isArray(symbolRange)) {
    const flat = symbolRange.flat ? symbolRange.flat() : [].concat(...symbolRange);
    symbol = "";
    for (let i = flat.length - 1; i >= 0; i--) {
      const v = (flat[i] ?? "").toString().trim();
      if (v) { symbol = v; break; }
    }
  }
  symbol = (symbol || "").toString().trim().toUpperCase();
  if (!symbol) return "#Symbol required";

  // Flatten inputs
  const strikes = Array.isArray(strikeRange) ? strikeRange.flat() : [strikeRange];
  const types = Array.isArray(typeRange) ? typeRange.flat() : [typeRange];
  const expirations = Array.isArray(expirationRange) ? expirationRange.flat() : [expirationRange];
  const qtys = Array.isArray(qtyRange) ? qtyRange.flat() : [qtyRange];

  // Build legs array
  const legs = [];
  const n = Math.max(strikes.length, types.length, expirations.length, qtys.length);
  for (let i = 0; i < n; i++) {
    const strike = parseNumber_(strikes[i] ?? "");
    const type = parseOptionType_(types[i] ?? "");
    const qty = parseNumber_(qtys[i] ?? "");
    const exp = expirations[i];

    if (!Number.isFinite(qty) || qty === 0) continue;
    if (!Number.isFinite(strike)) continue;
    if (!type || type === "Stock") continue;

    legs.push({ strike, type, qty, expiration: exp });
  }

  if (legs.length === 0) return "#No valid option legs";

  // Detect strategy
  const posType = detectPositionType_(legs);
  const strategyMap = {
    "bull-call-spread": "bull-call-spread",
    "bull-put-spread": "bull-put-spread",
    "iron-condor": "iron-condor",
    "iron-butterfly": "iron-butterfly",
    "bear-call-spread": "bear-call-spread",
    "long-call": "long-call",
    "short-call": "short-call",
    "long-put": "long-put",
    "short-put": "short-put",
  };
  const strategy = strategyMap[posType] || "custom";

  // Format date code (YYMMDD)
  function formatDateCode(exp) {
    let d = exp;
    if (!(d instanceof Date)) {
      d = new Date(exp);
    }
    if (isNaN(d.getTime())) return "000000";
    // Normalize to noon to avoid timezone/DST edge cases
    d = new Date(d);
    d.setHours(12, 0, 0, 0);
    return (
      String(d.getFullYear() % 100).padStart(2, "0") +
      String(d.getMonth() + 1).padStart(2, "0") +
      String(d.getDate()).padStart(2, "0")
    );
  }

  // Build leg strings: [sign].SYMBOL[YYMMDD][C/P][STRIKE]
  const legStrings = legs.map(leg => {
    const sign = leg.qty < 0 ? "-" : "";
    const dateCode = formatDateCode(leg.expiration);
    const typeChar = leg.type === "Call" ? "C" : "P";
    return `${sign}.${symbol}${dateCode}${typeChar}${leg.strike}`;
  });

  return `https://optionstrat.com/build/${strategy}/${symbol}/${legStrings.join(",")}`;
}

/**
 * build a URL for optionstrat.com
 * @param strikes -- a string containing strikes separated by '/' characters such as 400/450
 * @param ticker -- an uppercase stock symbol such as TSLA
 * @param strategy -- bull-call-spread, bull-put-spread literals as they appear in Optionstrat
 *                    URLs such as https://optionstrat.com/build/bull-call-spread/TSLA/.TSLA281215C440,-.TSLA281215C490
 * @param expiration -- an expiration date in many forms such as "Jun '28", or a Date.  If
 *                      no day-of-month is specified, then the 3rd Friday of the month is
 *                      computed (typical for long LEAPS)
 */
function buildOptionStratUrl(strikes, ticker, strategy, expiration) {

  const months = {
    Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6,
    Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12
  };


  function parseExpirationToDate(exp) {

    // Already a Date → use it (but create a copy at noon to avoid timezone issues)
    if (exp instanceof Date) {
      if (isNaN(exp.getTime())) throw new Error("Invalid Date expiration");
      // Create new date at noon local time to avoid timezone shifts
      return new Date(exp.getFullYear(), exp.getMonth(), exp.getDate(), 12, 0, 0);
    }

    let s = String(exp).trim();

    // Normalize "'28" → "2028"
    s = s.replace(/'(\d{2})\b/, "20$1");

    // Try to parse ISO format YYYY-MM-DD directly (avoids UTC timezone issues)
    const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (isoMatch) {
      const year = parseInt(isoMatch[1], 10);
      const month = parseInt(isoMatch[2], 10) - 1; // 0-indexed
      const day = parseInt(isoMatch[3], 10);
      return new Date(year, month, day, 12, 0, 0);
    }

    // Try M/D/YYYY format
    const mdyMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (mdyMatch) {
      const month = parseInt(mdyMatch[1], 10) - 1; // 0-indexed
      const day = parseInt(mdyMatch[2], 10);
      const year = parseInt(mdyMatch[3], 10);
      return new Date(year, month, day, 12, 0, 0);
    }

    // Fallback: Let JS parse it
    let d = new Date(s);
    if (isNaN(d.getTime())) {
      throw new Error("Invalid expiration format: " + exp);
    }

    // If no day was specified → assume monthly → 3rd Friday
    if (!/\b\d{1,2}\b/.test(s)) {
      d.setDate(getThirdFriday(d.getFullYear(), d.getMonth()));
    }

    // Normalize time (avoid DST edge cases)
    d = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0);

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
    const d = new Date(date);
    return (
      String(d.getFullYear() % 100).padStart(2, "0") +
      String(d.getMonth() + 1).padStart(2, "0") +
      String(d.getDate()).padStart(2, "0")
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
  const [lowStrike, highStrike] = parseStrikePairStrict_(strikes);
  const symbolLow = `.${ticker}${dateCode}${callOrPutChar}${lowStrike}`;
  const symbolHigh = `-.${ticker}${dateCode}${callOrPutChar}${highStrike}`;
 
  const url = `https://optionstrat.com/build/${strategy}/${ticker}/${symbolLow},${symbolHigh}`;
  //console.log("optionstrat URL:" + url);
  return url;
}
