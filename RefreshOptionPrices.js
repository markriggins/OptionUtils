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
      const expDate = parseYyyyMmDdToDateAtMidnight_(expStr);
      if (!expDate || isNaN(expDate.getTime())) {
        throw new Error(`Cannot parse expiration '${expStr}' from filename for ${symbol}. Expected format: exp-YYYY-MM-DD`);
      }

      const csvContent = file.getBlob().getDataAsString();
      const csvData = Utilities.parseCsv(csvContent);
      if (csvData.length < 2) continue;

      const parsedRows = loadCsvData_(csvData, symbol, expDate);
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

  // ---- Validate and write output ----
  const headersOut = ["symbol", "expiration", "strike", "type", "bid", "mid", "ask", "iv", "delta", "volume", "openint", "moneyness"];

  // Verify no rows have missing expirations
  const badRows = allRows.filter(r => !r[1] || (r[1] instanceof Date && isNaN(r[1].getTime())));
  if (badRows.length > 0) {
    throw new Error(`${badRows.length} rows have missing or invalid expiration dates. First bad row: ${JSON.stringify(badRows[0])}`);
  }

  targetSheet.getRange(1, 1, 1, headersOut.length).setValues([headersOut]);
  targetSheet.getRange(2, 1, allRows.length, headersOut.length).setValues(allRows);
  SpreadsheetApp.flush(); // Force commit to avoid timing issues

  targetSheet.setFrozenRows(1);

  // Sort: symbol, expiration, type, strike
  targetSheet.getRange(2, 1, allRows.length, headersOut.length).sort(
    ["symbol", "expiration", "type", "strike"].map(name => ({
      column: headersOut.indexOf(name) + 1, ascending: true
    }))
  );

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

  // Force recalculate of BullCallSpreads formulas
  try {
    forceRecalculateBullCallSpreads_(ss);
  } catch (e) {
    Logger.log("Could not force recalculate BullCallSpreads: " + e);
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
function loadCsvData_(csvData, symbol, expDate) {
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

    const strike = parseNumber_(r[strikeIdx]);
    if (!Number.isFinite(strike)) continue;

    const optionType = parseOptionType_(r[typeIdx]);
    if (!optionType) continue;

    const bid = parseNumber_(r[bidIdx]);
    const ask = parseNumber_(r[askIdx]);
    const mid = midIdx === -1 ? null : parseNumber_(r[midIdx]);

    // Parse IV (stored as decimal, e.g., 0.5532 for 55.32%)
    const ivRaw = ivIdx === -1 ? null : parsePercent_(r[ivIdx]);

    // Parse delta (already a decimal like 0.6873)
    const delta = deltaIdx === -1 ? null : parseNumber_(r[deltaIdx]);

    // Parse volume (integer)
    const volume = volumeIdx === -1 ? null : parseInteger_(r[volumeIdx]);

    // Parse open interest (integer)
    const openInt = openIntIdx === -1 ? null : parseInteger_(r[openIntIdx]);

    // Parse moneyness (stored as decimal, e.g., -0.0912 for -9.12%)
    const moneyness = moneynessIdx === -1 ? null : parsePercent_(r[moneynessIdx]);

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

/** ---------- helpers ---------- */

function getFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  if (!it.hasNext()) throw new Error(`Required folder not found: ${name}`);
  return it.next();
}

/**
 * Forces recalculation of a named table by adding then deleting a column on the left.
 */
function forceRecalculateTable_(ss, tableName) {
  const range = ss.getRangeByName(tableName);
  if (!range) {
    Logger.log(tableName + " not found");
    return;
  }

  const sheet = range.getSheet();
  const firstCol = range.getColumn();

  // Insert column before table to force recalculateulation
  sheet.insertColumnBefore(firstCol);
  sheet.getRange(range.getRow(), firstCol).setValue("Recalculate");
  SpreadsheetApp.flush();

  // Delete it
  sheet.deleteColumn(firstCol);
  SpreadsheetApp.flush();
  Logger.log("Forced recalculate of " + tableName + " via column add/delete");
}

/**
 * Forces recalculation of position tables.
 */
function forceRecalculateBullCallSpreads_(ss) {
  forceRecalculateTable_(ss, "BullCallSpreadsTable");
  forceRecalculateTable_(ss, "IronCondorsTable");
}

