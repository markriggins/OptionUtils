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
  console.log("✅ testbuildOptionStratUrl PASSED");
}

function buildOptionStratUrl(strikes, ticker, strategy, expiration) {

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
