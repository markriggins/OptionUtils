/**
 * refreshOptionPrices
 * Menu action: Refresh OptionPricesUploaded CSV files that have been uploaded into
 *      your Google drive under /Investing/Data/OptionPrices/<symbol>
 *
 * Currently supports the barchart.com format of option prices
 * For EACH expiration (exp-YYYY-MM-DD found in filename), load the MOST RECENT file
 * (by Drive "last updated") and ingest its rows.
 * Example File:
 *    amzn-options-exp-2028-12-15-monthly-show-all-stacked-01-15-2026.csv
 *       Strike,Moneyness,Bid,Mid,Ask,Latest,Change,%Change,Volume,"Open Int","OI Chg",IV,Delta,Type,Time
 *       115.00,+51.72%,140.00,142.50,145.00,141.60,-0.75,-0.53%,37,129,+43,43.04%,0.9330,Call,01/15/26
 *       120.00,+49.62%,136.50,138.75,141.00,144.00,unch,unch,0,156,unch,44.10%,0.9228,Call,01/13/26
 *
 * Output sheet columns (lowercase headers):
 *   symbol | expiration | strike | type | bid | mid | ask | iv | delta | volume | openint | moneyness
 *
 * Notes:
 * - expiration is stored as a REAL Date (midnight) for proper sorting/date math
 * - getOptionQuote_/XLookupByKeys normalizes Dates to day-strings for cache keys
 */
function refreshOptionPrices() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SHEET_NAME = "OptionPricesUploaded";
  let targetSheet = ss.getSheetByName(SHEET_NAME);
  if (!targetSheet) targetSheet = ss.insertSheet(SHEET_NAME);
  targetSheet.clearContents();

  const allRows = [];
  let symbolCount = 0;
  let expGroupsLoaded = 0;
  let filesScanned = 0;
  let filesSkippedNoExp = 0;

  // ---- Iterate symbol folders and find input files ----
  const fileMap = findInputFiles_();
  filesScanned = fileMap.filesScanned;
  filesSkippedNoExp = fileMap.filesSkippedNoExp;

  // ---- Process selected files ----
  for (const symbol in fileMap.bestBySymbol) {
    const bestByExp = fileMap.bestBySymbol[symbol];
    const expStrs = Object.keys(bestByExp);
    if (expStrs.length === 0) continue;

    for (const expStr of expStrs) {
      const entry = bestByExp[expStr];
      const file = entry.file;

      // Parse expStr into a Date (midnight) for sheet storage
      const expDate = parseYyyyMmDdToDate_(expStr);
      if (!expDate) {
        Logger.log(`Warning: Could not parse expiration '${expStr}' for ${symbol}`);
        continue;
      }

      const csvContent = file.getBlob().getDataAsString();
      const csvData = Utilities.parseCsv(csvContent);
      if (csvData.length < 2) continue;

      const parsedRows = parseCsvData_(csvData, symbol, expDate);
      if (parsedRows.length > 0) {
        allRows.push(...parsedRows);
        expGroupsLoaded++;
      }
    }

    symbolCount++;
  }

  if (allRows.length === 0) {
    let msg = "No valid data found.";
    msg += `\n\nScanned CSV files: ${filesScanned}`;
    msg += `\nSkipped (no exp-YYYY-MM-DD in filename): ${filesSkippedNoExp}`;
    ui.alert(msg);
    return;
  }

  // ---- Write output ----
  const headersOut = ["symbol", "expiration", "strike", "type", "bid", "mid", "ask", "iv", "delta", "volume", "openint", "moneyness"];
  targetSheet.getRange(1, 1, 1, headersOut.length).setValues([headersOut]);
  targetSheet.getRange(2, 1, allRows.length, headersOut.length).setValues(allRows);
  SpreadsheetApp.flush(); // Force commit to avoid timing issues

  targetSheet.setFrozenRows(1);

  // Sort: symbol, expiration, type, strike
  targetSheet.getRange(2, 1, allRows.length, headersOut.length).sort([
    { column: 1, ascending: true },
    { column: 2, ascending: true },
    { column: 4, ascending: true },
    { column: 3, ascending: true }
  ]);

  // Format IV and moneyness columns as percent
  const ivCol = headersOut.indexOf("iv") + 1;
  const moneynessCol = headersOut.indexOf("moneyness") + 1;
  if (allRows.length > 0) {
    targetSheet.getRange(2, ivCol, allRows.length, 1).setNumberFormat("0.00%");
    targetSheet.getRange(2, moneynessCol, allRows.length, 1).setNumberFormat("0.00%");
  }

  // Filter + banding
  const fullRange = targetSheet.getRange(1, 1, allRows.length + 1, headersOut.length);
  if (targetSheet.getFilter()) targetSheet.getFilter().remove();
  fullRange.createFilter();
  try { fullRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY); } catch (e) {}

  // Clear memo cache so lookups get fresh data
  try {
    XLookupByKeys_clearMemo();
  } catch (e) {}

  // Warm caches (optional but recommended)
  try {
    XLookupByKeys_WarmCache(SHEET_NAME, ["symbol", "expiration", "strike", "type"], ["bid", "mid", "ask", "iv", "delta", "volume", "openint", "moneyness"]);
  } catch (e) {}

  // Force recalculation of BullCallSpreads formulas
  try {
    forceRecalcBullCallSpreads_(ss);
  } catch (e) {
    Logger.log("Could not force recalc BullCallSpreads: " + e);
  }

  ss.toast(
    `Refreshed ${allRows.length} rows from ${symbolCount} symbols\n` +
      `Loaded latest files for ${expGroupsLoaded} expirations`,
    "OptionPrices",
    5
  );
}

/**
 * Finds and organizes input CSV files from the OptionPrices folder.
 *
 * Scans CSV files directly in the folder (no subfolders needed).
 * Extracts symbol and expiration from filename pattern: <symbol>-options-exp-<YYYY-MM-DD>-...csv
 * Selects the most recent file per symbol/expiration based on last updated time.
 *
 * @param {string} [path="Investing/Data/OptionPrices"] - The path to the OptionPrices folder (from root).
 * @returns {Object} {
 *   bestBySymbol: { [symbol]: { [expStr]: { file: File, updated: number } } },
 *   filesScanned: number,
 *   filesSkippedNoExp: number
 * }
 */
function findInputFiles_(path = "Investing/Data/OptionPrices") {
  // ---- Locate folder ----
  const root = DriveApp.getRootFolder();
  let opFolder = root;
  const parts = path.split('/').filter(p => p.trim());
  for (const part of parts) {
    opFolder = getFolder_(opFolder, part);
  }

  const bestBySymbol = Object.create(null);
  let filesScanned = 0;
  let filesSkippedNoExp = 0;

  // Scan all CSV files directly in the folder
  // Force fresh query by accessing folder properties first (workaround for Drive API caching)
  opFolder.getName();
  Utilities.sleep(100);

  const files = opFolder.getFilesByType(MimeType.CSV);
  while (files.hasNext()) {
    const file = files.next();
    filesScanned++;

    const fname = String(file.getName()).toLowerCase();

    // Extract symbol and expiration from filename
    // Pattern: <symbol>-options-exp-<YYYY-MM-DD>-...csv
    // Examples: amzn-options-exp-2028-12-15-monthly-show-all-stacked-01-15-2026.csv
    //           tsla-options-exp-2028-06-16-monthly-show-all-stacked-01-15-2026.csv
    const m = fname.match(/^([a-z]+)-.*exp-(\d{4}-\d{2}-\d{2})(?:\D|$)/i);
    if (!m || !m[1] || !m[2]) {
      filesSkippedNoExp++;
      continue;
    }

    const symbol = m[1].toUpperCase();
    const expStr = m[2];
    const updated = file.getLastUpdated().getTime();

    // Initialize symbol entry if needed
    if (!bestBySymbol[symbol]) {
      bestBySymbol[symbol] = Object.create(null);
    }

    const prev = bestBySymbol[symbol][expStr];
    if (!prev || updated > prev.updated) {
      bestBySymbol[symbol][expStr] = { file, updated };
    }
  }

  Logger.log(`findInputFiles_: scanned ${filesScanned} files, skipped ${filesSkippedNoExp}, found ${Object.keys(bestBySymbol).length} symbols`);
  for (const sym in bestBySymbol) {
    Logger.log(`  ${sym}: ${Object.keys(bestBySymbol[sym]).length} expirations`);
  }

  return { bestBySymbol, filesScanned, filesSkippedNoExp };
}

/**
 * Parses CSV data into output rows for the sheet.
 *
 * Processes the CSV array (from Utilities.parseCsv), finds column indexes (case-insensitive),
 * and extracts relevant fields for each data row.
 *
 * Skips rows with invalid strike or type.
 * Optional columns (mid, iv, delta, volume, openInt, moneyness) set to null if not found.
 *
 * @param {Array<Array<string>>} csvData - Parsed CSV array (headers in row 0).
 * @param {string} symbol - Uppercase symbol.
 * @param {Date} expDate - Expiration as Date object (midnight).
 * @returns {Array<Array<*>}} Array of [symbol, expDate, strike, type, bid, mid, ask, iv, delta, volume, openInt, moneyness] rows.
 */
function parseCsvData_(csvData, symbol, expDate) {
  const rows = [];

  const headers = csvData[0].map(h => String(h).trim().toLowerCase());
  const { strikeIdx, bidIdx, midIdx, askIdx, typeIdx, ivIdx, deltaIdx, volumeIdx, openIntIdx, moneynessIdx } = findColumnIndexes_(headers);

  if (strikeIdx === -1 || bidIdx === -1 || askIdx === -1 || typeIdx === -1) {
    Logger.log(`Skipping ${symbol} for exp ${expDate}: missing required columns`);
    return rows;
  }

  if (midIdx === -1) {
    Logger.log(`Note ${symbol} for exp ${expDate}: no mid column found (mid will be null)`);
  }

  for (let i = 1; i < csvData.length; i++) {
    const r = csvData[i];
    if (!r || r.length === 0) continue;

    const strike = safeNumber_(r[strikeIdx]);
    if (!Number.isFinite(strike)) continue;

    const optionType = normalizeOptionType_(r[typeIdx]);
    if (!optionType) continue;

    const bid = safeNumber_(r[bidIdx]);
    const ask = safeNumber_(r[askIdx]);
    const mid = midIdx === -1 ? null : safeNumber_(r[midIdx]);

    // Parse IV (stored as decimal, e.g., 0.5532 for 55.32%)
    const ivRaw = ivIdx === -1 ? null : safePercent_(r[ivIdx]);

    // Parse delta (already a decimal like 0.6873)
    const delta = deltaIdx === -1 ? null : safeNumber_(r[deltaIdx]);

    // Parse volume (integer)
    const volume = volumeIdx === -1 ? null : safeInteger_(r[volumeIdx]);

    // Parse open interest (integer)
    const openInt = openIntIdx === -1 ? null : safeInteger_(r[openIntIdx]);

    // Parse moneyness (stored as decimal, e.g., -0.0912 for -9.12%)
    const moneyness = moneynessIdx === -1 ? null : safePercent_(r[moneynessIdx]);

    rows.push([
      symbol,
      expDate,
      strike,
      optionType,
      Number.isFinite(bid) ? bid : null,
      Number.isFinite(mid) ? mid : null,
      Number.isFinite(ask) ? ask : null,
      Number.isFinite(ivRaw) ? ivRaw : null,
      Number.isFinite(delta) ? delta : null,
      Number.isFinite(volume) ? volume : null,
      Number.isFinite(openInt) ? openInt : null,
      Number.isFinite(moneyness) ? moneyness : null,
    ]);
  }

  return rows;
}

/**
 * Finds indexes of required columns in headers (all in lowercase).
 *
 * Supports variations in column names:
 * - strike: "strike"
 * - bid: "bid"
 * - mid: "mid" (optional)
 * - ask: "ask"
 * - type: "type", "option type", "call/put", "cp", "put/call"
 * - iv: "iv", "implied volatility", "impliedvolatility"
 * - delta: "delta"
 * - volume: "volume", "vol"
 * - openInt: "open int", "openint", "open interest", "openinterest", "oi"
 * - moneyness: "moneyness", "money", "itm/otm"
 *
 * @param {Array<string>} headers - Lowercased, trimmed headers.
 * @returns {Object} { strikeIdx, bidIdx, midIdx, askIdx, typeIdx, ivIdx, deltaIdx, volumeIdx, openIntIdx, moneynessIdx } (-1 if not found).
 */
function findColumnIndexes_(headers) {
  const strikeIdx = headers.indexOf("strike");
  const bidIdx = headers.indexOf("bid");
  const midIdx = headers.indexOf("mid");
  const askIdx = headers.indexOf("ask");

  let typeIdx = headers.indexOf("type");
  if (typeIdx === -1) typeIdx = headers.indexOf("option type");
  if (typeIdx === -1) typeIdx = headers.indexOf("call/put");
  if (typeIdx === -1) typeIdx = headers.indexOf("cp");
  if (typeIdx === -1) typeIdx = headers.indexOf("put/call");

  let ivIdx = headers.indexOf("iv");
  if (ivIdx === -1) ivIdx = headers.indexOf("implied volatility");
  if (ivIdx === -1) ivIdx = headers.indexOf("impliedvolatility");

  const deltaIdx = headers.indexOf("delta");

  let volumeIdx = headers.indexOf("volume");
  if (volumeIdx === -1) volumeIdx = headers.indexOf("vol");

  let openIntIdx = headers.indexOf("open int");
  if (openIntIdx === -1) openIntIdx = headers.indexOf("openint");
  if (openIntIdx === -1) openIntIdx = headers.indexOf("open interest");
  if (openIntIdx === -1) openIntIdx = headers.indexOf("openinterest");
  if (openIntIdx === -1) openIntIdx = headers.indexOf("oi");

  let moneynessIdx = headers.indexOf("moneyness");
  if (moneynessIdx === -1) moneynessIdx = headers.indexOf("money");
  if (moneynessIdx === -1) moneynessIdx = headers.indexOf("itm/otm");

  return { strikeIdx, bidIdx, midIdx, askIdx, typeIdx, ivIdx, deltaIdx, volumeIdx, openIntIdx, moneynessIdx };
}

/** ---------- helpers ---------- */

function getFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  if (!it.hasNext()) throw new Error(`Required folder not found: ${name}`);
  return it.next();
}

/**
 * Forces recalculation of a named table by adding then deleting a column on the left.
 */
function forceRecalcTable_(ss, tableName) {
  const range = ss.getRangeByName(tableName);
  if (!range) {
    Logger.log(tableName + " not found");
    return;
  }

  const sheet = range.getSheet();
  const firstCol = range.getColumn();

  // Insert column before table
  sheet.insertColumnBefore(firstCol);
  sheet.getRange(range.getRow(), firstCol).setValue("Recalc");
  SpreadsheetApp.flush();

  // Delete it
  sheet.deleteColumn(firstCol);
  SpreadsheetApp.flush();

  Logger.log("Forced recalc of " + tableName + " via column add/delete");
}

/**
 * Forces recalculation of position tables.
 */
function forceRecalcBullCallSpreads_(ss) {
  forceRecalcTable_(ss, "BullCallSpreadsTable");
  forceRecalcTable_(ss, "IronCondorsTable");
}

function parseYyyyMmDdToDate_(s) {
  // Create a Date at midnight local time for deterministic day comparisons
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const y = Number(m[1]), mo = Number(m[2]) - 1, d = Number(m[3]);
  const dt = new Date(y, mo, d);
  if (isNaN(dt.getTime())) return null;
  dt.setHours(0, 0, 0, 0);
  return dt;
}

function safeNumber_(v) {
  if (v === null || v === undefined) return NaN;
  const s = String(v).trim();
  if (!s || s === "--" || s.toLowerCase() === "na" || s.toLowerCase() === "n/a" || s.toLowerCase() === "unch") return NaN;
  const n = Number(s.replace(/,/g, ""));
  return Number.isFinite(n) ? n : NaN;
}

/**
 * Parse a percentage string like "55.32%" or "-9.12%" to decimal (0.5532 or -0.0912)
 */
function safePercent_(v) {
  if (v === null || v === undefined) return NaN;
  const s = String(v).trim();
  if (!s || s === "--" || s.toLowerCase() === "na" || s.toLowerCase() === "n/a") return NaN;

  // Handle +/- prefix and % suffix
  const cleaned = s.replace(/,/g, "").replace(/%$/, "");
  const n = Number(cleaned);
  if (!Number.isFinite(n)) return NaN;

  // Convert from percentage to decimal (55.32 -> 0.5532)
  return n / 100;
}

/**
 * Parse an integer, handling commas
 */
function safeInteger_(v) {
  if (v === null || v === undefined) return NaN;
  const s = String(v).trim();
  if (!s || s === "--" || s.toLowerCase() === "na" || s.toLowerCase() === "n/a" || s.toLowerCase() === "unch") return NaN;

  const cleaned = s.replace(/,/g, "").replace(/^\+/, "");
  const n = parseInt(cleaned, 10);
  return Number.isFinite(n) ? n : NaN;
}

function normalizeOptionType_(v) {
  const s = String(v || "").trim().toLowerCase();
  if (!s) return null;
  if (["call", "calls", "c"].includes(s)) return "Call";
  if (["put", "puts", "p"].includes(s)) return "Put";
  return null;
}

/** ---------- test cases ---------- */

/**
 * Tests parseCsvData_ with different column orders.
 *
 * Mocks CSV data with headers in varying orders (e.g., Ask,Bid,Mid or Bid,Mid,Ask).
 * Ensures output rows are correct regardless of order, as long as names match (case-insensitive).
 *
 * Run: test_parseCsvData_columnOrders
 */
function test_parseCsvData_columnOrders() {
  const symbol = "TSLA";
  const expDate = new Date(2028, 5, 16); // 2028-06-16 (month 5=June)

  // Test 1: Standard order Bid,Mid,Ask + Type at end
  const csv1 = [
    ["Strike", "Bid", "Mid", "Ask", "Type"],
    ["450", "203.15", "206.00", "208.85", "Call"],
    ["350", "250.00", "255.00", "260.00", "Put"]
  ];
  const rows1 = parseCsvData_(csv1, symbol, expDate);
  assertArrayDeepEqual(rows1, [
    [symbol, expDate, 450, "Call", 203.15, 206.00, 208.85],
    [symbol, expDate, 350, "Put", 250.00, 255.00, 260.00]
  ], "Test 1: Standard order");

  // Test 2: Reversed order Ask,Bid,Mid + Type in middle
  const csv2 = [
    ["Ask", "Bid", "Mid", "Type", "Strike"],
    ["208.85", "203.15", "206.00", "Call", "450"],
    ["260.00", "250.00", "255.00", "Put", "350"]
  ];
  const rows2 = parseCsvData_(csv2, symbol, expDate);
  assertArrayDeepEqual(rows2, [
    [symbol, expDate, 450, "Call", 203.15, 206.00, 208.85],
    [symbol, expDate, 350, "Put", 250.00, 255.00, 260.00]
  ], "Test 2: Reversed order");

  // Test 3: Mixed case headers, no Mid, different Type name "Option Type"
  const csv3 = [
    ["sTriKe", "BID", "ASK", "Option Type"],
    ["450", "203.15", "208.85", "Call"],
    ["350", "250.00", "260.00", "Put"]
  ];
  const rows3 = parseCsvData_(csv3, symbol, expDate);
  assertArrayDeepEqual(rows3, [
    [symbol, expDate, 450, "Call", 203.15, null, 208.85],
    [symbol, expDate, 350, "Put", 250.00, null, 260.00]
  ], "Test 3: Mixed case, no Mid, alt Type");

  // Test 4: Invalid (missing Bid) -> empty rows
  const csv4 = [
    ["Strike", "Mid", "Ask", "Type"],
    ["450", "206.00", "208.85", "Call"]
  ];
  const rows4 = parseCsvData_(csv4, symbol, expDate);
  assertArrayDeepEqual(rows4, [], "Test 4: Missing Bid -> empty");

  Logger.log("✅ All parseCsvData_ column order tests passed");
}

/**
 * Helper: Assert two arrays are deeply equal.
 *
 * @param {Array<*>} actual
 * @param {Array<*>} expected
 * @param {string} msg
 */
function assertArrayDeepEqual(actual, expected, msg = "") {
  if (actual.length !== expected.length) {
    throw new Error(`ASSERT FAILED${msg ? " – " + msg : ""}\nLength mismatch: ${actual.length} != ${expected.length}`);
  }
  actual.forEach((row, i) => {
    const expRow = expected[i];
    if (row.length !== expRow.length) {
      throw new Error(`ASSERT FAILED${msg ? " – " + msg : ""}\nRow ${i} length mismatch`);
    }
    row.forEach((v, j) => {
      if (v instanceof Date && expRow[j] instanceof Date) {
        if (v.getTime() !== expRow[j].getTime()) {
          throw new Error(`ASSERT FAILED${msg ? " – " + msg : ""}\nRow ${i}, Col ${j}: Dates differ`);
        }
      } else if (v !== expRow[j]) {
        throw new Error(`ASSERT FAILED${msg ? " – " + msg : ""}\nRow ${i}, Col ${j}: ${v} != ${expRow[j]}`);
      }
    });
  });
}