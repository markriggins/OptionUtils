/** 
 * XLookupByKeys (v3.4) — cached return-map builder (robust cache load)
 *
 * Fix: CacheService can return partial/expired chunks while the count key still exists.
 *      We now treat ANY decode/ungzip/parse error as a cache miss (rebuild).
 *
 * Added:
 * - XLookupByKeys_WarmCache(sheetName, keyHeaders, returnHeaders) (generic)
 * - OptionTools menu + refreshOptionPrices (option-specific caller)
 */

/** ========= CONFIG ========= */
const CACHE_TTL_SECONDS = 6 * 60 * 60; // 6 hours
const CACHE_KEY_PREFIX = "XLookupByKeys:v3:";

/** In-memory memo (per execution) */
let __MEMO = Object.create(null);

/**
 * @customfunction
 */
function XLookupByKeys(keyValues, keyHeaders, returnHeaders, sheetName = DEFAULT_SHEET) {
  keyValues     = flatten_(keyValues);
  keyHeaders    = flatten_(keyHeaders);
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
  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const rows = values.slice(1);

  const colIndex = Object.create(null);
  headers.forEach((h, i) => (colIndex[h] = i));

  keyHeaders.forEach(h => {
    const k = String(h).trim().toLowerCase();
    if (!(k in colIndex)) throw new Error(`Key column '${h}' not found in sheet '${sheetName}'`);
  });
  returnHeaders.forEach(h => {
    const r = String(h).trim().toLowerCase();
    if (!(r in colIndex)) throw new Error(`Return column '${h}' not found in sheet '${sheetName}'`);
  });

  const keyIdx = keyHeaders.map(h => colIndex[String(h).trim().toLowerCase()]);
  const retIdx = returnHeaders.map(h => colIndex[String(h).trim().toLowerCase()]);

  const map = Object.create(null);

  for (const row of rows) {
    if (!row || row.every(v => v === "" || v === null)) continue;

    const k = keyIdx.map(i => normalize_(row[i])).join("|");
    const ret = retIdx.map(i => row[i]);
    map[k] = ret;
  }

  cacheSaveChunked_(cache, cacheKeyBase, map, CACHE_TTL_SECONDS);
  __MEMO[cacheKeyBase] = map;

  return { map, sig, cacheKeyBase };
}

function makeCompositeKey_(keyValues) {
  return keyValues.map(v => normalize_(v)).join("|");
}

function normalize_(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  if (typeof v === "number") return String(v);
  return String(v).trim();
}

function flatten_(v) {
  if (!Array.isArray(v)) return [v];
  return v.flat();
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

  const CHUNK_SIZE = 80000;

  const chunks = [];
  for (let i = 0; i < b64.length; i += CHUNK_SIZE) {
    chunks.push(b64.slice(i, i + CHUNK_SIZE));
  }

  cache.put(`${baseKey}:n`, String(chunks.length), ttlSeconds);
  for (let i = 0; i < chunks.length; i++) {
    cache.put(`${baseKey}:${i}`, chunks[i], ttlSeconds);
  }

  // optional integrity hint
  cache.put(`${baseKey}:len`, String(b64.length), ttlSeconds);
}

/**
 * Robust cache load:
 * - if any chunk missing => miss
 * - if base64Decode/ungzip/JSON.parse throws => miss
 */
function cacheLoadChunked_(cache, baseKey) {
  try {
    const nStr = cache.get(`${baseKey}:n`);
    if (!nStr) return null;

    const n = Number(nStr);
    if (!Number.isFinite(n) || n <= 0 || n > 5000) return null;

    let b64 = "";
    for (let i = 0; i < n; i++) {
      const part = cache.get(`${baseKey}:${i}`);
      if (!part) return null; // partial/expired => miss
      b64 += part;
    }

    // Optional length check (helps detect truncation)
    const lenStr = cache.get(`${baseKey}:len`);
    if (lenStr) {
      const expectedLen = Number(lenStr);
      if (Number.isFinite(expectedLen) && expectedLen > 0 && b64.length !== expectedLen) {
        return null;
      }
    }

    const gzBytes = Utilities.base64Decode(b64); // can throw "Invalid argument"
    const gzBlob = Utilities.newBlob(gzBytes, "application/gzip", "map.json.gz");
    const jsonBlob = Utilities.ungzip(gzBlob);   // can throw
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
 *  TESTS (self-contained) — PRESERVED
 * =========================== */

function assertEqual(actual, expected, msg = "") {
  if (actual !== expected) {
    throw new Error(
      `ASSERT FAILED${msg ? " – " + msg : ""}\nExpected: ${expected}\nActual:   ${actual}`
    );
  }
}

function assertArrayEqual(actual, expected, msg = "") {
  if (actual.length !== expected.length) {
    throw new Error(
      `ASSERT FAILED${msg ? " – " + msg : ""}\nLength mismatch\nExpected: ${expected.length}\nActual:   ${actual.length}`
    );
  }
  actual.forEach((v, i) => {
    if (v !== expected[i]) {
      throw new Error(
        `ASSERT FAILED${msg ? " – " + msg : ""}\nIndex ${i}\nExpected: ${expected[i]}\nActual:   ${v}`
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
  assertEqual(afterSecondKeys, afterFirstKeys, "Memo should be reused without creating new entries");

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
    .filter(k => typeof globalObj[k] === "function" && k.startsWith("test_XLookupByKeys"))
    .sort();

  Logger.log(`Found ${tests.length} tests`);
  tests.forEach(fn => {
    Logger.log(`Running ${fn}`);
    globalObj[fn]();
  });

  Logger.log("✅ All tests completed");
}

