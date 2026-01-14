/**
 * lookupOptionQuote
 *
 * Returns an object:
 *   { bid: Number|null, mid: Number|null, ask: Number|null }
 *
 * Uses XLookupByKeys under the hood (generic).
 *
 * @param {string} symbol      e.g. "TSLA"
 * @param {string} expiration e.g. "2028-06-16" (YYYY-MM-DD)
 * @param {number} strike     e.g. 450
 * @param {string} [type]     optional: "Call" or "Put" (default: "Call")
 */
function lookupOptionQuote(symbol, expiration, strike, type = "Call", sheetName = "OptionPricesUploaded") {
  const result = XLookupByKeys(
    [String(symbol).toUpperCase(), expiration, Number(strike), type],
    ["symbol", "expiration", "strike", "type"],
    ["bid", "mid", "ask"],
    sheetName
  );

  // XLookupByKeys always returns a 2D array [ [ ... ] ]
  const row = result && result[0];
  if (!row) {
    return { bid: null, mid: null, ask: null };
  }

  const [bid, mid, ask] = row;

  return {
    bid: bid === "" ? null : bid,
    mid: mid === "" ? null : mid,
    ask: ask === "" ? null : ask
  };
}
