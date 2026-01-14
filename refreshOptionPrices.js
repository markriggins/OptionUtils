/**
 * refreshOptionPrices
 * Menu action: Refresh OptionPricesUploaded from all symbol CSVs
 *
 * Folder path:
 *   Drive root / Investing / Data / OptionPrices / [SYMBOL] / *.csv
 *
 * For EACH expiration (exp-YYYY-MM-DD found in filename), load the MOST RECENT file
 * (by Drive "last updated") and ingest its rows.
 *
 * Output sheet columns (lowercase headers):
 *   symbol | expiration | strike | type | bid | mid | ask
 *
 * Notes:
 * - expiration is stored as a REAL Date (midnight) for proper sorting/date math
 * - getOptionQuote_/XLookupByKeys normalize Dates to day-strings for cache keys
 */
function refreshOptionPrices() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SHEET_NAME = "OptionPricesUploaded";
  let targetSheet = ss.getSheetByName(SHEET_NAME);
  if (!targetSheet) targetSheet = ss.insertSheet(SHEET_NAME);
  targetSheet.clearContents();

  // ---- Locate folders ----
  const root = DriveApp.getRootFolder();
  const investing = getFolder_(root, "Investing");
  const dataFolder = getFolder_(investing, "Data");
  const opParent = getFolder_(dataFolder, "OptionPrices");

  const allRows = [];
  let symbolCount = 0;
  let expGroupsLoaded = 0;
  let filesScanned = 0;
  let filesSkippedNoExp = 0;

  // ---- Iterate symbol folders ----
  const symbolFolders = opParent.getFolders();
  while (symbolFolders.hasNext()) {
    const symFolder = symbolFolders.next();
    const symbol = symFolder.getName().trim().toUpperCase();
    if (!symbol) continue;

    // Build: expiration -> { file, updated }
    const bestByExp = Object.create(null);

    const files = symFolder.getFilesByType(MimeType.CSV);
    while (files.hasNext()) {
      const file = files.next();
      filesScanned++;

      const fname = String(file.getName()).toLowerCase();
      const m = fname.match(/exp-(\d{4}-\d{2}-\d{2})(?:\D|$)/i);
      if (!m || !m[1]) {
        filesSkippedNoExp++;
        continue;
      }

      const expStr = m[1];
      const updated = file.getLastUpdated().getTime();

      const prev = bestByExp[expStr];
      if (!prev || updated > prev.updated) {
        bestByExp[expStr] = { file, updated };
      }
    }

    const expStrs = Object.keys(bestByExp);
    if (expStrs.length === 0) continue;

    // Ingest each expiration's most-recent file
    for (const expStr of expStrs) {
      const entry = bestByExp[expStr];
      const file = entry.file;

      // Parse expStr into a Date (midnight) for sheet storage
      const expDate = parseYyyyMmDdToDate_(expStr);
      if (!expDate) continue;

      const csvContent = file.getBlob().getDataAsString();
      const csvData = Utilities.parseCsv(csvContent);
      if (csvData.length < 2) continue;

      const headers = csvData[0].map(h => String(h).trim().toLowerCase());

      const strikeIdx = headers.findIndex(h => h.includes("strike"));
      const bidIdx = headers.findIndex(h => h === "bid" || h.includes("bid"));
      const midIdx = headers.findIndex(h => h === "mid" || h.includes("mid"));
      const askIdx = headers.findIndex(h => h === "ask" || h.includes("ask"));

      // Type column can vary by source: "type", "option type", "call/put", etc.
      let typeIdx = headers.findIndex(h => h === "type");
      if (typeIdx === -1) typeIdx = headers.findIndex(h => h.includes("option type"));
      if (typeIdx === -1) typeIdx = headers.findIndex(h => h.includes("call/put"));
      if (typeIdx === -1) typeIdx = headers.findIndex(h => h.includes("cp"));
      if (typeIdx === -1) typeIdx = headers.findIndex(h => h.includes("put/call"));

      if (strikeIdx === -1 || bidIdx === -1 || askIdx === -1) {
        Logger.log(`Skipping ${symbol} ${expStr}: missing strike/bid/ask columns in ${file.getName()}`);
        continue;
      }
      if (midIdx === -1) {
        // mid is optional; we'll compute later if you want, but keep as null for now
        Logger.log(`Note ${symbol} ${expStr}: no mid column found in ${file.getName()} (mid will be null)`);
      }
      if (typeIdx === -1) {
        Logger.log(`Skipping ${symbol} ${expStr}: missing type column in ${file.getName()}`);
        continue;
      }

      for (let i = 1; i < csvData.length; i++) {
        const r = csvData[i];
        if (!r || r.length === 0) continue;

        const strike = safeNumber_(r[strikeIdx]);
        if (!Number.isFinite(strike)) continue;

        const type = normalizeType_(r[typeIdx]);
        if (!type) continue;

        const bid = safeNumber_(r[bidIdx]);
        const ask = safeNumber_(r[askIdx]);
        const mid = midIdx === -1 ? null : safeNumber_(r[midIdx]);

        allRows.push([
          symbol,
          expDate,     // Date (midnight)
          strike,
          type,        // "Call" / "Put"
          Number.isFinite(bid) ? bid : null,
          Number.isFinite(mid) ? mid : null,
          Number.isFinite(ask) ? ask : null
        ]);
      }

      expGroupsLoaded++;
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
  const headersOut = ["symbol", "expiration", "strike", "type", "bid", "mid", "ask"];
  targetSheet.getRange(1, 1, 1, headersOut.length).setValues([headersOut]);
  targetSheet.getRange(2, 1, allRows.length, headersOut.length).setValues(allRows);

  targetSheet.setFrozenRows(1);

  // Sort: symbol, expiration, type, strike
  targetSheet.getRange(2, 1, allRows.length, headersOut.length).sort([
    { column: 1, ascending: true },
    { column: 2, ascending: true },
    { column: 4, ascending: true },
    { column: 3, ascending: true }
  ]);

  // Filter + banding
  const fullRange = targetSheet.getRange(1, 1, allRows.length + 1, headersOut.length);
  if (targetSheet.getFilter()) targetSheet.getFilter().remove();
  fullRange.createFilter();
  try { fullRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY); } catch (e) {}

  // Warm caches (optional but recommended)
  // (Assumes XLookupByKeys_WarmCache exists; comment out if you don't have it.)
  try {
    XLookupByKeys_WarmCache(SHEET_NAME, ["symbol", "expiration", "strike", "type"], ["bid", "mid", "ask"]);
  } catch (e) {
    // ignore if warm function not present
  }

  ss.toast(
    `Refreshed ${allRows.length} rows from ${symbolCount} symbols\n` +
      `Loaded latest files for ${expGroupsLoaded} expirations`,
    "OptionPrices",
    5
  );
}

/** ---------- helpers ---------- */

function getFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  if (!it.hasNext()) throw new Error(`Required folder not found: ${name}`);
  return it.next();
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
  if (!s || s === "--" || s.toLowerCase() === "na" || s.toLowerCase() === "n/a") return NaN;
  const n = Number(s.replace(/,/g, ""));
  return Number.isFinite(n) ? n : NaN;
}

function normalizeType_(v) {
  const s = String(v || "").trim().toLowerCase();
  if (!s) return null;
  if (s === "call" || s === "c") return "Call";
  if (s === "put" || s === "p") return "Put";
  // Sometimes data has "Calls"/"Puts"
  if (s.startsWith("call")) return "Call";
  if (s.startsWith("put")) return "Put";
  return null;
}
