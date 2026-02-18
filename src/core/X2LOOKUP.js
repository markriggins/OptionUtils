
/**
 * X2LOOKUP(expiration, strike, ExpCol, StrikeCol, ReturnCol)
 * Two-key lookup for Google Sheets custom functions.
 *
 *  USAGE: X2LOOKUP(jun, 350, Exp, Strike, Bid)
 *
 * looks for a row with 'jun' in the Exp range and 350 in the Strike range
 * and returns the corresponding value from the Bid range
 *
 * Returns:
 *  - value from ReturnCol when found
 *  - "NA" if not found
 *  - "REF" if ranges are mismatched/missing
 *  - "VALUE" if wrong arg count
 *
 * @customfunction
 */
function X2LOOKUP(key1, key2, col1, col2, returnCol) {
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


/**
 * Run via: Apps Script → Run → test_X2LOOKUP
 * If nothing throws, the test PASSES.
 */
function test_X2LOOKUP() {

  // Simulated OptionPrices columns (exactly how Sheets passes them)
  const Exp = [
    [new Date(2028, 5, 16)], // Jun 16 2028
    [new Date(2028,11,15)], // Dec 15 2028
    [new Date(2028, 5, 16)],
    [new Date(2028,11,15)]
  ];

  const Strike = [
    [350],
    [350],
    [550],
    [550]
  ];

  const Bid = [
    [203.15],
    [214.00],
    [133.10],
    [148.00]
  ];

  const Ask = [
    [207.05],
    [221.90],
    [136.15],
    [157.00]
  ];

  const jun = new Date(2028, 5, 16);
  const dec = new Date(2028,11,15);

  // ---- HAPPY PATHS ----
  assertEqual(X2LOOKUP(jun, 350, Exp, Strike, Bid), 203.15);
  assertEqual(X2LOOKUP(jun, 550, Exp, Strike, Ask), 136.15);
  assertEqual(X2LOOKUP(dec, 350, Exp, Strike, Ask), 221.90);
  assertEqual(X2LOOKUP(dec, 550, Exp, Strike, Bid), 148.00);

  // ---- TYPE COERCION ----
  assertEqual(X2LOOKUP(jun, "350", Exp, Strike, Bid), 203.15);

  // ---- NOT FOUND ----
  assertEqual(X2LOOKUP(new Date(2028, 11, 16), 350, Exp, Strike, Bid), "#NA");

  // ---- RANGE MISMATCH ----
  const BadBid = [[1], [2], [3]];
  assertEqual(X2LOOKUP(jun, 350, Exp, Strike, BadBid), "#REF");
  console.log("✅ test_X2LOOKUP PASSED");

}

