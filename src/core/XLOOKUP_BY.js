/**
 * XLOOKUP_BY - Multi-key lookup with variable number of keys.
 *
 * Usage:
 *   XLOOKUP_BY(key1, col1, returnCol)                    // 1 key
 *   XLOOKUP_BY(key1, col1, key2, col2, returnCol)        // 2 keys
 *   XLOOKUP_BY(key1, col1, key2, col2, key3, col3, returnCol)  // 3 keys
 *   ... and so on for any number of keys
 *
 * @param {...*} args - Pairs of (key, column) followed by return column.
 * @returns {*} The matching value from returnCol, or error string.
 * @customfunction
 */
function XLOOKUP_BY(...args) {
  // Must have odd number of args >= 3: (key, col) pairs + returnCol
  if (args.length < 3 || args.length % 2 === 0) {
    return "#VALUE! Expected: key1, col1, [key2, col2, ...], returnCol";
  }

  const numKeys = (args.length - 1) / 2;
  const returnCol = args[args.length - 1];

  // Extract keys and columns
  const keys = [];
  const cols = [];
  for (let i = 0; i < numKeys; i++) {
    keys.push(args[i * 2]);
    cols.push(args[i * 2 + 1]);
  }

  // Validate columns
  if (cols.some(c => c == null) || returnCol == null) {
    return "#REF! Missing column range";
  }

  // Convert ranges to 1D arrays
  const colArrays = cols.map(to1D_);
  const returnArray = to1D_(returnCol);

  // Validate all columns have same length
  const len = returnArray.length;
  if (colArrays.some(c => c.length !== len)) {
    return "#REF! Column ranges must have same length";
  }

  // Normalize keys for comparison
  const normalizedKeys = keys.map(normalizeValue_);

  // Search for matching row
  for (let i = 0; i < len; i++) {
    let match = true;
    for (let k = 0; k < numKeys; k++) {
      const cellValue = normalizeValue_(colArrays[k][i]);
      if (cellValue !== normalizedKeys[k]) {
        match = false;
        break;
      }
    }
    if (match) {
      return returnArray[i];
    }
  }

  return "#N/A";
}

/**
 * Convert Sheets range (2D array) or scalar to 1D array.
 * @private
 */
function to1D_(x) {
  if (!Array.isArray(x)) return [x];
  const out = new Array(x.length);
  for (let i = 0; i < x.length; i++) {
    const row = x[i];
    out[i] = Array.isArray(row) ? row[0] : row;
  }
  return out;
}

/**
 * Normalize value for comparison.
 * - Dates -> midnight timestamp
 * - Strings -> trimmed uppercase
 * - Numbers -> as-is
 * @private
 */
function normalizeValue_(v) {
  if (v instanceof Date) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate()).getTime();
  }
  if (typeof v === "string") {
    return v.trim().toUpperCase();
  }
  if (typeof v === "number" || !isNaN(Number(v))) {
    return Number(v);
  }
  return v;
}

// ============================================================
// DEPRECATED aliases - use XLOOKUP_BY instead
// ============================================================

/**
 * DEPRECATED: Use XLOOKUP_BY instead.
 * Two-key lookup with new signature: (key1, col1, key2, col2, returnCol).
 * @customfunction
 */
function X2LOOKUP(key1, col1, key2, col2, returnCol) {
  return XLOOKUP_BY(key1, col1, key2, col2, returnCol);
}

/**
 * DEPRECATED: Use XLOOKUP_BY instead.
 * Three-key lookup with new signature: (key1, col1, key2, col2, key3, col3, returnCol).
 * @customfunction
 */
function X3LOOKUP(key1, col1, key2, col2, key3, col3, returnCol) {
  return XLOOKUP_BY(key1, col1, key2, col2, key3, col3, returnCol);
}

// ============================================================
// Tests
// ============================================================

function test_XLOOKUP_BY() {
  // Simulated data columns (as Sheets passes them - 2D arrays)
  const Symbol = [["TSLA"], ["TSLA"], ["TSLA"], ["AMZN"], ["AMZN"]];
  const Exp = [
    [new Date(2028, 5, 16)],  // Jun 16 2028
    [new Date(2028, 11, 15)], // Dec 15 2028
    [new Date(2028, 5, 16)],
    [new Date(2028, 11, 15)],
    [new Date(2028, 11, 15)]
  ];
  const Strike = [[350], [350], [550], [180], [200]];
  const Bid = [[203.15], [214.00], [133.10], [22.30], [18.50]];
  const Ask = [[207.05], [221.90], [136.15], [24.10], [20.25]];

  const jun = new Date(2028, 5, 16);
  const dec = new Date(2028, 11, 15);

  // ---- 1 KEY ----
  assertEqual(XLOOKUP_BY("AMZN", Symbol, Bid), 22.30, "1 key: AMZN -> first match");

  // ---- 2 KEYS ----
  assertEqual(XLOOKUP_BY(jun, Exp, 350, Strike, Bid), 203.15, "2 keys: jun/350");
  assertEqual(XLOOKUP_BY(jun, Exp, 550, Strike, Ask), 136.15, "2 keys: jun/550");
  assertEqual(XLOOKUP_BY(dec, Exp, 350, Strike, Ask), 221.90, "2 keys: dec/350");

  // ---- 3 KEYS ----
  assertEqual(XLOOKUP_BY("TSLA", Symbol, jun, Exp, 350, Strike, Bid), 203.15, "3 keys: TSLA/jun/350");
  assertEqual(XLOOKUP_BY("TSLA", Symbol, dec, Exp, 350, Strike, Bid), 214.00, "3 keys: TSLA/dec/350");
  assertEqual(XLOOKUP_BY("AMZN", Symbol, dec, Exp, 180, Strike, Bid), 22.30, "3 keys: AMZN/dec/180");

  // ---- CASE INSENSITIVE ----
  assertEqual(XLOOKUP_BY("tsla", Symbol, jun, Exp, 550, Strike, Bid), 133.10, "case insensitive");

  // ---- NOT FOUND ----
  assertEqual(XLOOKUP_BY("GOOG", Symbol, Bid), "#N/A", "not found");
  assertEqual(XLOOKUP_BY(jun, Exp, 999, Strike, Bid), "#N/A", "strike not found");

  // ---- ERRORS ----
  assertEqual(XLOOKUP_BY(jun, Exp), "#VALUE! Expected: key1, col1, [key2, col2, ...], returnCol", "too few args");
  assertEqual(XLOOKUP_BY(jun, Exp, 350, Strike), "#VALUE! Expected: key1, col1, [key2, col2, ...], returnCol", "even args");

  // ---- RANGE MISMATCH ----
  const BadCol = [[1], [2], [3]];
  assertEqual(XLOOKUP_BY(jun, Exp, 350, BadCol, Bid), "#REF! Column ranges must have same length", "range mismatch");

  // ---- X2LOOKUP ALIAS (new signature) ----
  assertEqual(X2LOOKUP(jun, Exp, 350, Strike, Bid), 203.15, "X2LOOKUP alias");

  // ---- X3LOOKUP ALIAS (new signature) ----
  assertEqual(X3LOOKUP("TSLA", Symbol, jun, Exp, 350, Strike, Bid), 203.15, "X3LOOKUP alias");

  log.info("test", "All XLOOKUP_BY tests passed");
}
