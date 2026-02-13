/** ========= CONFIG ========= */
const CACHE_TTL_SECONDS = 6 * 60 * 1; // 1 hour
const CACHE_KEY_PREFIX = "XLookupByKeys:v3:";
/** In-memory memo (per execution) */
let __MEMO = Object.create(null);

/**
 * XLookupByKeys - Performs a multi-key lookup in a Google Spreadsheet sheet, returning values from specified return columns.
 *
 * This function builds or retrieves a cached map of return values keyed by a composite key (joined by '|') from the provided key values.
 * Lookups are case-insensitive (headers are lowercased for matching), and values are normalized (trimmed strings, dates formatted as YYYY-MM-DD, numbers as strings).
 *
 * Caching Strategy:
 * - In-memory memoization (per execution) for fastest access.
 * - Document Cache (chunked, gzipped, base64-encoded) with TTL of 1 hour for persistence across executions.
 * - Cache key incorporates sheet signature (ID, dimensions, header hash) and spec hash (keys/returns) to invalidate on changes.
 * - If cache miss, builds the map from sheet data, skipping empty rows.
 *
 * If no match is found, returns an array of empty strings matching the returnHeaders length.
 *
 * Related Functions:
 * - XLookupByKeys_WarmCache(sheetName, keyHeaders, returnHeaders): Proactively builds/rebuilds the cache for the given spec (useful after data refreshes).
 * - XLookupByKeys_clearMemo(): Clears the in-memory memo (e.g., for testing or after edits).
 * - XLookupByKeys_onEdit(e): Clears memo on sheet edits (install as onEdit trigger).
 *
 * Version: 3.5
 * Changes:
 * - Added generic WarmCache helper and OptionTools menu integration (with refreshOptionPrices caller).
 * - Optimizations: Batch cache.getAll/putAll for chunks; early empty row filtering; faster header processing.
 *
 * Usage Examples:
 *
 * 1. Spreadsheet Formula (Custom Function):
 *    Assume a sheet named "Options" with headers: symbol, expiration, strike, type, bid, mid, ask.
 *    In cell E2 (to get bid/mid/ask for keys in A2:D2):
 *    =XLookupByKeys(A2:D2, {"symbol", "expiration", "strike", "type"}, {"bid", "mid", "ask"}, "Options")
 *
 *    This returns a horizontal array: [bid, mid, ask] or ["", "", ""] if no match.
 *
 * 2. From Another Script (e.g., to fetch values programmatically):
 *    const result = XLookupByKeys(
 *      ["TSLA", "2028-06-16", 450, "Call"],
 *      ["symbol", "expiration", "strike", "type"],
 *      ["bid", "mid", "ask"],
 *      "Options"
 *    );
 *    // result = [[203.15, 206.00, 208.85]] or [["", "", ""]]
 *
 * 3. Warm Cache (e.g., after refreshing data in "Options" sheet):
 *    XLookupByKeys_WarmCache(
 *      "Options",
 *      ["symbol", "expiration", "strike", "type"],
 *      ["bid", "mid", "ask"]
 *    );
 *    // Returns true; cache is built for future lookups.
 *
 * 4. Handling Missing Data:
 *    =XLookupByKeys(A2:D2, {"symbol", "expiration", "strike", "type"}, {"bid", "mid", "ask"}, "Options")
 *    // If no row matches keys in A2:D2, returns ["", "", ""]
 *
 * Note: For vertical output, wrap in TRANSPOSE(). Ensure DEFAULT_SHEET is defined if using default sheetName.
 *
 * @param {Array|string} keyValues - The value(s) for the key columns (flattened if array).
 * @param {Array|string} keyHeaders - The header name(s) of the key columns (flattened if array, case-insensitive).
 * @param {Array|string} returnHeaders - The header name(s) of the return columns (flattened if array, case-insensitive).
 * @param {string} [sheetName=DEFAULT_SHEET] - The name of the sheet to query (defaults to DEFAULT_SHEET if not provided).
 * @returns {Array<Array<*>}} - A 2D array with one row of return values (or empties if no match).
 * @throws {Error} - If sheet not found, or required headers missing.
 */
function XLookupByKeys(keyValues, keyHeaders, returnHeaders, sheetName = DEFAULT_SHEET) {
  keyValues = flatten_(keyValues);
  keyHeaders = flatten_(keyHeaders);
  returnHeaders = flatten_(returnHeaders);
  const { map } = getOrBuildReturnMap_(sheetName, keyHeaders, returnHeaders);
  const compositeKey = makeCompositeKey_(keyValues);
  const hit = map[compositeKey];
  if (!hit) return [returnHeaders.map(() => "")];
  return [hit];
}

/**
 * Generic helper: proactively build the cache for a given (sheet, keys, returns).
 * Call from scripts (e.g., after refreshOptionPrices).
 */
function XLookupByKeys_WarmCache(sheetName, keyHeaders, returnHeaders) {
  keyHeaders = flatten_(keyHeaders);
  returnHeaders = flatten_(returnHeaders);
  getOrBuildReturnMap_(sheetName, keyHeaders, returnHeaders);
  return true;
}

function getOrBuildReturnMap_(sheetName, keyHeaders, returnHeaders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet '${sheetName}' not found`);
  const sig = getSheetSignature_(sheet);
  const specHash = hashSpec_(keyHeaders, returnHeaders);
  const cacheKeyBase = `${CACHE_KEY_PREFIX}${sheetName}:${sig}:${specHash}`;
  // 1) In-memory memo
  if (__MEMO[cacheKeyBase]) return { map: __MEMO[cacheKeyBase], sig, cacheKeyBase };
  // 2) DocumentCache (chunked) — robust load
  const cache = CacheService.getDocumentCache();
  const loaded = cacheLoadChunked_(cache, cacheKeyBase);
  if (loaded) {
    __MEMO[cacheKeyBase] = loaded;
    return { map: loaded, sig, cacheKeyBase };
  }
  // 3) Build from sheet
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    const empty = Object.create(null);
    __MEMO[cacheKeyBase] = empty;
    return { map: empty, sig, cacheKeyBase };
  }
  // Match headers case-insensitively by lowercasing them
  const headers = values[0].map(h => h.toString().trim().toLowerCase());
  let rows = values.slice(1);
  // Skip empty rows early
  rows = rows.filter(row => row.some(v => v !== "" && v != null));
  const colIndex = {};
  headers.forEach((h, i) => colIndex[h] = i);
  const lowerKeyHeaders = keyHeaders.map(h => h.toString().trim().toLowerCase());
  const lowerRetHeaders = returnHeaders.map(h => h.toString().trim().toLowerCase());
  lowerKeyHeaders.forEach(k => {
    if (!(k in colIndex)) throw new Error(`Key column '${k}' not found in sheet '${sheetName}'`);
  });
  lowerRetHeaders.forEach(r => {
    if (!(r in colIndex)) throw new Error(`Return column '${r}' not found in sheet '${sheetName}'`);
  });
  const keyIdx = lowerKeyHeaders.map(k => colIndex[k]);
  const retIdx = lowerRetHeaders.map(r => colIndex[r]);
  const map = {};
  for (const row of rows) {
    const k = keyIdx.map(i => normalize_(row[i])).join("|");
    const ret = retIdx.map(i => row[i]);
    map[k] = ret;
  }
  cacheSaveChunked_(cache, cacheKeyBase, map, CACHE_TTL_SECONDS);
  __MEMO[cacheKeyBase] = map;
  return { map, sig, cacheKeyBase };
}

function makeCompositeKey_(keyValues) {
  return keyValues.map(normalize_).join("|");
}

function normalize_(v) {
  if (v == null) return "";
  // Normalize dates to M/D/YYYY format (add 12 hours to avoid timezone boundary issues)
  if (v instanceof Date) {
    const d = new Date(v.getTime() + 12 * 60 * 60 * 1000);
    return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
  }
  if (typeof v === "number") return v.toString();

  const s = v.toString().trim();

  // Convert ISO dates (YYYY-MM-DD) to M/D/YYYY
  const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    const [, y, m, d] = isoMatch;
    return `${parseInt(m, 10)}/${parseInt(d, 10)}/${y}`;
  }

  return s;
}

function flatten_(v) {
  return Array.isArray(v) ? v.flat() : [v];
}

function getSheetSignature_(sheet) {
  const r = sheet.getLastRow();
  const c = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, c).getValues()[0].join(",");
  const headerHash = Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, header)
  );
  return `${sheet.getSheetId()}:${r}x${c}:${headerHash}`;
}

function hashSpec_(keyHeaders, returnHeaders) {
  const spec = JSON.stringify({ k: keyHeaders, r: returnHeaders });
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, spec)
  );
}

/** ---- Chunked cache helpers ---- */
function cacheSaveChunked_(cache, baseKey, obj, ttlSeconds) {
  const json = JSON.stringify(obj);
  const jsonBlob = Utilities.newBlob(json, "application/json", "map.json");
  const gzBlob = Utilities.gzip(jsonBlob);
  const b64 = Utilities.base64Encode(gzBlob.getBytes());
  const CHUNK_SIZE = 80000; // Close to 100KB limit post-overhead
  const chunks = [];
  for (let i = 0; i < b64.length; i += CHUNK_SIZE) {
    chunks.push(b64.slice(i, i + CHUNK_SIZE));
  }
  const putDict = {};
  putDict[`${baseKey}:n`] = chunks.length.toString();
  putDict[`${baseKey}:len`] = b64.length.toString(); // integrity hint
  chunks.forEach((chunk, i) => putDict[`${baseKey}:${i}`] = chunk);
  cache.putAll(putDict, ttlSeconds);
}

/**
 * Robust cache load:
 * - if any chunk missing => miss
 * - if base64Decode/ungzip/JSON.parse throws => miss
 */
function cacheLoadChunked_(cache, baseKey) {
  try {
    // Batch get metadata (:n, :len)
    const metaKeys = [`${baseKey}:n`, `${baseKey}:len`];
    const meta = cache.getAll(metaKeys);
    const nStr = meta[metaKeys[0]];
    if (!nStr) return null;
    const n = +nStr;
    if (!Number.isFinite(n) || n <= 0 || n > 5000) return null;
    // Batch get chunks
    const chunkKeys = [];
    for (let i = 0; i < n; i++) chunkKeys.push(`${baseKey}:${i}`);
    const parts = cache.getAll(chunkKeys);
    if (Object.keys(parts).length !== n) return null; // Missing chunks => miss
    let b64 = "";
    for (let i = 0; i < n; i++) {
      const part = parts[chunkKeys[i]];
      if (!part) return null; // Shouldn't happen after length check, but safety
      b64 += part;
    }
    // Optional length check (helps detect truncation)
    const lenStr = meta[metaKeys[1]];
    if (lenStr) {
      const expectedLen = +lenStr;
      if (Number.isFinite(expectedLen) && expectedLen > 0 && b64.length !== expectedLen) {
        return null;
      }
    }
    const gzBytes = Utilities.base64Decode(b64); // can throw "Invalid argument"
    const gzBlob = Utilities.newBlob(gzBytes, "application/gzip", "map.json.gz");
    const jsonBlob = Utilities.ungzip(gzBlob); // can throw
    const json = jsonBlob.getDataAsString("UTF-8");
    return JSON.parse(json);
  } catch (e) {
    return null; // treat ANY error as miss
  }
}

function XLookupByKeys_clearMemo() {
  __MEMO = Object.create(null);
}

function XLookupByKeys_onEdit(e) {
  XLookupByKeys_clearMemo();
}

/** ===========================
 * TESTS (self-contained)
 * =========================== */
function assertEqual_XLookupByKeys_(actual, expected, msg = "") {
  if (actual !== expected) {
    throw new Error(
      `ASSERT FAILED${msg ? " – " + msg : ""}\nExpected: ${expected}\nActual: ${actual}`
    );
  }
}

function assertArrayEqual(actual, expected, msg = "") {
  if (actual.length !== expected.length) {
    throw new Error(
      `ASSERT FAILED${msg ? " – " + msg : ""}\nLength mismatch\nExpected: ${expected.length}\nActual: ${actual.length}`
    );
  }
  actual.forEach((v, i) => {
    if (v !== expected[i]) {
      throw new Error(
        `ASSERT FAILED${msg ? " – " + msg : ""}\nIndex ${i}\nExpected: ${expected[i]}\nActual: ${v}`
      );
    }
  });
}

function ensureTestSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = "Options__TEST";
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();
  const headers = ["symbol", "expiration", "strike", "type", "bid", "mid", "ask"];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  const rows = [
    ["TSLA", "2028-06-16", 450, "Call", 203.15, 206.00, 208.85],
    ["TSLA", "2028-06-16", 350, "Call", 250.00, 255.00, 260.00],
    ["AMZN", "2026-02-06", 260, "Call", 6.10, 6.30, 6.50],
  ];
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  return sh.getName();
}

function test_XLookupByKeys_returns_expected_values() {
  XLookupByKeys_clearMemo();
  const sheetName = ensureTestSheet_();
  const result = XLookupByKeys(
    ["TSLA", "2028-06-16", 450, "Call"],
    ["symbol", "expiration", "strike", "type"],
    ["bid", "mid", "ask"],
    sheetName
  );
  assertArrayEqual(result[0], [203.15, 206.00, 208.85], "lookup returned wrong bid/mid/ask");
  Logger.log("✅ test_XLookupByKeys_returns_expected_values PASSED");
}

function test_XLookupByKeys_cache_reused_in_memory() {
  XLookupByKeys_clearMemo();
  const sheetName = ensureTestSheet_();
  const beforeKeys = Object.keys(__MEMO).length;
  const r1 = XLookupByKeys(
    ["AMZN", "2026-02-06", 260, "Call"],
    ["symbol", "expiration", "strike", "type"],
    ["bid", "mid", "ask"],
    sheetName
  );
  assertArrayEqual(r1[0], [6.10, 6.30, 6.50], "first call wrong");
  const afterFirstKeys = Object.keys(__MEMO).length;
  if (afterFirstKeys <= beforeKeys) throw new Error("Expected memo to gain at least one entry");
  const r2 = XLookupByKeys(
    ["AMZN", "2026-02-06", 260, "Call"],
    ["symbol", "expiration", "strike", "type"],
    ["bid", "mid", "ask"],
    sheetName
  );
  assertArrayEqual(r2[0], [6.10, 6.30, 6.50], "second call wrong");
  const afterSecondKeys = Object.keys(__MEMO).length;
  assertEqual_XLookupByKeys_(afterSecondKeys, afterFirstKeys, "Memo should be reused without creating new entries");
  Logger.log("✅ test_XLookupByKeys_cache_reused_in_memory PASSED");
}

function test_XLookupByKeys_missing_row_returns_blanks() {
  XLookupByKeys_clearMemo();
  const sheetName = ensureTestSheet_();
  const result = XLookupByKeys(
    ["TSLA", "2028-06-16", 999, "Call"],
    ["symbol", "expiration", "strike", "type"],
    ["bid", "mid", "ask"],
    sheetName
  );
  assertArrayEqual(result[0], ["", "", ""], "missing row should return blanks");
  Logger.log("✅ test_XLookupByKeys_missing_row_returns_blanks PASSED");
}

function test_XLookupByKeys_all() {
  const globalObj = globalThis;
  const tests = Object.keys(globalObj)
    .filter(k => typeof globalObj[k] === "function" && k.startsWith("test_XLookupByKeys") && k !== "test_XLookupByKeys_all")
    .sort();
  Logger.log(`Found ${tests.length} tests`);
  tests.forEach(fn => {
    Logger.log(`Running ${fn}`);
    globalObj[fn]();
  });
  Logger.log("✅ All tests completed");
}
