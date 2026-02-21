/**
 * RefreshOptionPrices.js
 * Handles option price uploads from CSV files (direct upload, no Drive storage).
 *
 * Currently supports the barchart.com format of option prices.
 * Example File:
 *    amzn-options-exp-2028-12-15-monthly-show-all-stacked-01-15-2026.csv
 *       Strike,Moneyness,Bid,Mid,Ask,Latest,Change,%Change,Volume,"Open Int","OI Chg",IV,Delta,Type,Time
 *       115.00,+51.72%,140.00,142.50,145.00,141.60,-0.75,-0.53%,37,129,+43,43.04%,0.9330,Call,01/15/26
 *       120.00,+49.62%,136.50,138.75,141.00,144.00,unch,unch,0,156,unch,44.10%,0.9228,Call,01/13/26
 *
 * Output sheet columns (lowercase headers):
 *   symbol | expiration | strike | type | bid | mid | ask | iv | delta | volume | openint | moneyness | dataDate
 *
 * Notes:
 * - expiration is stored as a REAL Date (midnight) for proper sorting/date math
 * - dataDate is extracted from filename (MM-DD-YYYY.csv pattern) or defaults to upload date
 * - getOptionQuote_/XLookupByKeys normalizes Dates to day-strings for cache keys
 */

/**
 * Forces Portfolio sheet to recalculate custom functions.
 * Called from dialog after refreshing option prices.
 * @returns {string} Status message
 */
function refreshPortfolioPrices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const portfolioSheet = ss.getSheetByName("Portfolio");

  if (!portfolioSheet) {
    return "No Portfolio sheet found.";
  }

  // Insert and delete a column after col 1 to force formula recalculation
  portfolioSheet.insertColumnAfter(1);
  SpreadsheetApp.flush();
  portfolioSheet.deleteColumn(2);
  SpreadsheetApp.flush();

  return "Portfolio formulas refreshed with new option prices.";
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

  // Validate required columns
  const required = validateRequiredColumns_(headers, [
    { name: "Strike", aliases: ["strike"] },
    { name: "Bid", aliases: ["bid"] },
    { name: "Ask", aliases: ["ask"] },
    { name: "Type", aliases: ["type", "option type", "call/put", "cp", "put/call"] },
  ], `Option Prices CSV (${symbol} exp ${expDate})`);

  const strikeIdx = required.Strike;
  const bidIdx = required.Bid;
  const askIdx = required.Ask;
  const typeIdx = required.Type;

  // Find optional columns
  const optional = findOptionalColumns_(headers, [
    { name: "Mid", aliases: ["mid"] },
    { name: "IV", aliases: ["iv", "implied volatility"] },
    { name: "Delta", aliases: ["delta"] },
    { name: "Volume", aliases: ["volume", "vol"] },
    { name: "OpenInt", aliases: ["open int", "open interest", "oi"] },
    { name: "Moneyness", aliases: ["moneyness", "money", "itm/otm"] },
  ]);

  const midIdx = optional.Mid;
  const ivIdx = optional.IV;
  const deltaIdx = optional.Delta;
  const volumeIdx = optional.Volume;
  const openIntIdx = optional.OpenInt;
  const moneynessIdx = optional.Moneyness;

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

/* =========================================================
   File Upload Dialog
   ========================================================= */

/**
 * Checks if the OptionPricesUploaded sheet has any data.
 * @returns {boolean} True if there are existing option prices.
 */
function hasExistingOptionPrices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("OptionPricesUploaded");
  if (!sheet) return false;
  return sheet.getLastRow() > 1; // More than just header row
}

/**
 * Shows the file upload dialog for option prices.
 */
function showUploadOptionPricesDialog() {
  const html = HtmlService.createHtmlOutputFromFile("ui/FileUpload")
    .setWidth(500)
    .setHeight(400);
  // Inject the mode
  const content = html.getContent().replace(
    "if (mode) init(mode);",
    "init('optionPrices');"
  );
  const output = HtmlService.createHtmlOutput(content)
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(output, "Upload Option Prices & Refresh");
}

/**
 * Handles uploaded option price files from the file chooser.
 * Parses CSV directly to sheet (no Drive storage).
 * @param {Array<{name: string, content: string}>} files - Array of file objects
 * @param {boolean} replaceAll - If true, clear all existing prices first
 * @param {boolean} confirmed - If true, skip orphan warning check
 * @returns {string|Object} Status message or confirmation request object
 */
function uploadOptionPrices(files, replaceAll, confirmed) {
  if (!files || files.length === 0) {
    throw new Error("No files provided");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SHEET_NAME = "OptionPricesUploaded";

  // Parse all files to extract symbol/expiration and rows
  const parsedFiles = [];
  const uploadedCombos = new Set(); // "SYMBOL|YYYY-MM-DD" keys

  for (const file of files) {
    const cleanName = stripBrowserDisambiguator_(file.name);
    const parsed = parseOptionPriceFile_(cleanName, file.content);
    if (parsed) {
      parsedFiles.push(parsed);
      uploadedCombos.add(`${parsed.symbol}|${parsed.expStr}`);
    }
  }

  if (parsedFiles.length === 0) {
    throw new Error("No valid option price files found. Expected filename pattern: <symbol>-options-exp-YYYY-MM-DD-....csv");
  }

  // If replaceAll and not confirmed, check for orphaned portfolio expirations
  if (replaceAll && !confirmed) {
    const orphaned = findOrphanedPortfolioExpirations_(ss, uploadedCombos);
    if (orphaned.length > 0) {
      return {
        needsConfirmation: true,
        orphaned: orphaned,
        message: "These portfolio positions will lose prices:\n• " + orphaned.join("\n• ")
      };
    }
  }

  // Get or create sheet
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Headers now include dataDate
  const headersOut = ["symbol", "expiration", "strike", "type", "bid", "mid", "ask", "iv", "delta", "volume", "openint", "moneyness", "dataDate"];

  // Track skipped files (older than existing data)
  const skippedFiles = [];

  if (replaceAll) {
    // Clear entire sheet
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headersOut.length).setValues([headersOut]);
  } else {
    // Merge mode: ensure headers exist
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      sheet.getRange(1, 1, 1, headersOut.length).setValues([headersOut]);
    }

    // Get existing dataDates to compare
    const existingDates = getExistingDataDates_(sheet);

    // Filter out files that are older than existing data
    const filesToUpload = [];
    const combosToDelete = new Set();

    for (const pf of parsedFiles) {
      const key = `${pf.symbol}|${pf.expStr}`;
      const existingDate = existingDates.get(key);

      if (existingDate && pf.dataDate <= existingDate) {
        // Uploaded file is same age or older - skip it
        skippedFiles.push({
          symbol: pf.symbol,
          expStr: pf.expStr,
          uploadedDate: pf.dataDateStr,
          existingDate: formatDateMDYYYY_(existingDate)
        });
      } else {
        // Uploaded file is newer - include it
        filesToUpload.push(pf);
        combosToDelete.add(key);
      }
    }

    // Update parsedFiles to only include newer files
    parsedFiles.length = 0;
    parsedFiles.push(...filesToUpload);

    // Delete existing rows only for combos we're updating
    if (combosToDelete.size > 0) {
      deleteRowsForCombos_(sheet, combosToDelete);
    }
  }

  // Collect all rows from parsed files (only newer ones in merge mode)
  const allRows = [];
  for (const pf of parsedFiles) {
    allRows.push(...pf.rows);
  }

  // Append new rows
  if (allRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, allRows.length, headersOut.length).setValues(allRows);

    // Sort and format
    const totalRows = sheet.getLastRow() - 1;
    if (totalRows > 0) {
      sheet.getRange(2, 1, totalRows, headersOut.length).sort([
        { column: 1, ascending: true },  // symbol
        { column: 2, ascending: true },  // expiration
        { column: 4, ascending: true },  // type
        { column: 3, ascending: true }   // strike
      ]);
    }

    // Format columns
    const ivCol = headersOut.indexOf("iv") + 1;
    const moneynessCol = headersOut.indexOf("moneyness") + 1;
    const dataDateCol = headersOut.indexOf("dataDate") + 1;
    if (totalRows > 0) {
      sheet.getRange(2, ivCol, totalRows, 1).setNumberFormat("0.00%");
      sheet.getRange(2, moneynessCol, totalRows, 1).setNumberFormat("0.00%");
      sheet.getRange(2, dataDateCol, totalRows, 1).setNumberFormat("m/d/yyyy");
    }

    // Filter + banding
    const fullRange = sheet.getRange(1, 1, totalRows + 1, headersOut.length);
    if (sheet.getFilter()) sheet.getFilter().remove();
    fullRange.createFilter();
    try { fullRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY); } catch (e) {}
  }

  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();

  // Clear lookup caches
  try { XLookupByKeys_clearMemo(); } catch (e) {}
  try {
    XLookupByKeys_WarmCache(SHEET_NAME, ["symbol", "expiration", "strike", "type"], ["bid", "mid", "ask", "iv", "delta", "volume", "openint", "moneyness"]);
  } catch (e) {}

  // Build summary
  let summary = "";

  if (parsedFiles.length > 0) {
    const mode = replaceAll ? "Replaced all" : "Merged";
    summary += `${mode}: ${allRows.length} option prices from ${parsedFiles.length} file(s).\n\n`;
    summary += "Files processed:\n";
    for (const pf of parsedFiles) {
      summary += `  • ${pf.symbol} exp ${pf.expStr} (data from ${pf.dataDateStr})\n`;
    }
  }

  if (skippedFiles.length > 0) {
    if (summary) summary += "\n";
    summary += `Skipped ${skippedFiles.length} file(s) with older data:\n`;
    for (const sf of skippedFiles) {
      summary += `  • ${sf.symbol} exp ${sf.expStr} (uploaded ${sf.uploadedDate}, existing ${sf.existingDate})\n`;
    }
  }

  if (!summary) {
    summary = "No files were uploaded (all skipped as older than existing data).";
  }

  return summary;
}

/**
 * Parses an option price CSV file.
 * @param {string} filename - The filename
 * @param {string} content - The CSV content
 * @returns {Object|null} { symbol, expStr, dataDate, dataDateStr, rows } or null if invalid
 */
function parseOptionPriceFile_(filename, content) {
  // Extract symbol and expiration from filename
  // Pattern: <symbol>-options-exp-<YYYY-MM-DD>-...csv
  const match = filename.toLowerCase().match(/^([a-z]+)-.*exp-(\d{4}-\d{2}-\d{2})/i);
  if (!match) {
    log.warn("upload", `Skipping file (no exp-YYYY-MM-DD): ${filename}`);
    return null;
  }

  const symbol = match[1].toUpperCase();
  const expStr = match[2];
  const expDate = parseDateAtMidnight_(expStr);
  if (!expDate) {
    log.warn("upload", `Skipping file (invalid expiration): ${filename}`);
    return null;
  }

  // Try to extract data date from filename (the date the prices were captured)
  // Pattern: ...-MM-DD-YYYY.csv at the end
  const dataDateMatch = filename.match(/(\d{2})-(\d{2})-(\d{4})\.csv$/i);
  let dataDate;
  let dataDateStr;
  if (dataDateMatch) {
    dataDate = parseDateAtMidnight_(`${dataDateMatch[3]}-${dataDateMatch[1]}-${dataDateMatch[2]}`);
    dataDateStr = `${dataDateMatch[1]}/${dataDateMatch[2]}/${dataDateMatch[3]}`;
  } else {
    // Fallback to today
    dataDate = new Date();
    dataDate.setHours(0, 0, 0, 0);
    dataDateStr = formatDateMDYYYY_(dataDate);
  }

  // Parse CSV
  const csvData = Utilities.parseCsv(content);
  if (csvData.length < 2) {
    log.warn("upload", `Skipping file (no data rows): ${filename}`);
    return null;
  }

  const rows = loadCsvDataWithDate_(csvData, symbol, expDate, dataDate);
  if (rows.length === 0) {
    log.warn("upload", `Skipping file (no valid rows): ${filename}`);
    return null;
  }

  return { symbol, expStr, expDate, dataDate, dataDateStr, rows };
}

/**
 * Parses CSV data into output rows, including dataDate column.
 */
function loadCsvDataWithDate_(csvData, symbol, expDate, dataDate) {
  const rows = [];
  const headers = csvData[0].map(h => String(h).trim().toLowerCase());

  // Validate required columns
  const required = validateRequiredColumns_(headers, [
    { name: "Strike", aliases: ["strike"] },
    { name: "Bid", aliases: ["bid"] },
    { name: "Ask", aliases: ["ask"] },
    { name: "Type", aliases: ["type", "option type", "call/put", "cp", "put/call"] },
  ], `Option Prices CSV (${symbol} exp ${expDate})`);

  const strikeIdx = required.Strike;
  const bidIdx = required.Bid;
  const askIdx = required.Ask;
  const typeIdx = required.Type;

  // Find optional columns
  const optional = findOptionalColumns_(headers, [
    { name: "Mid", aliases: ["mid"] },
    { name: "IV", aliases: ["iv", "implied volatility"] },
    { name: "Delta", aliases: ["delta"] },
    { name: "Volume", aliases: ["volume", "vol"] },
    { name: "OpenInt", aliases: ["open int", "open interest", "oi"] },
    { name: "Moneyness", aliases: ["moneyness", "money", "itm/otm"] },
  ]);

  for (let i = 1; i < csvData.length; i++) {
    const r = csvData[i];
    if (!r || r.length === 0) continue;

    const strike = parseNumber_(r[strikeIdx]);
    if (!Number.isFinite(strike)) continue;

    const optionType = parseOptionType_(r[typeIdx]);
    if (!optionType) continue;

    const bid = parseNumber_(r[bidIdx]);
    const ask = parseNumber_(r[askIdx]);
    const mid = optional.Mid === -1 ? null : parseNumber_(r[optional.Mid]);
    const ivRaw = optional.IV === -1 ? null : parsePercent_(r[optional.IV]);
    const delta = optional.Delta === -1 ? null : parseNumber_(r[optional.Delta]);
    const volume = optional.Volume === -1 ? null : parseInteger_(r[optional.Volume]);
    const openInt = optional.OpenInt === -1 ? null : parseInteger_(r[optional.OpenInt]);
    const moneyness = optional.Moneyness === -1 ? null : parsePercent_(r[optional.Moneyness]);

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
      dataDate
    ]);
  }

  return rows;
}

/**
 * Finds portfolio expirations that won't have prices after replace-all.
 * @param {Spreadsheet} ss
 * @param {Set<string>} uploadedCombos - Set of "SYMBOL|YYYY-MM-DD" keys being uploaded
 * @returns {string[]} List of "SYMBOL EXP_DATE" strings for orphaned positions
 */
function findOrphanedPortfolioExpirations_(ss, uploadedCombos) {
  const orphaned = [];

  // Check Portfolio sheet
  const portfolioRange = getNamedRangeWithTableFallback_(ss, "Portfolio");
  if (!portfolioRange) return orphaned;

  const rows = portfolioRange.getValues();
  if (rows.length < 2) return orphaned;

  const headers = rows[0];
  const symIdx = findColumn_(headers, ["symbol", "ticker"]);
  const expIdx = findColumn_(headers, ["expiration", "exp", "expiry"]);

  if (symIdx < 0 || expIdx < 0) return orphaned;

  const seen = new Set();
  for (let i = 1; i < rows.length; i++) {
    const sym = String(rows[i][symIdx] || "").trim().toUpperCase();
    const expRaw = rows[i][expIdx];
    if (!sym || !expRaw) continue;

    const expDate = parseDateAtMidnight_(expRaw);
    if (!expDate) continue;

    const expStr = expDate.getFullYear() + "-" +
      String(expDate.getMonth() + 1).padStart(2, "0") + "-" +
      String(expDate.getDate()).padStart(2, "0");

    const key = `${sym}|${expStr}`;
    if (!uploadedCombos.has(key) && !seen.has(key)) {
      seen.add(key);
      orphaned.push(`${sym} ${expStr}`);
    }
  }

  return orphaned;
}

/**
 * Gets the most recent dataDate for each symbol/expiration combo in the sheet.
 * @param {Sheet} sheet
 * @returns {Map<string, Date>} Map of "SYMBOL|YYYY-MM-DD" to dataDate
 */
function getExistingDataDates_(sheet) {
  const result = new Map();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dataDateIdx = headers.findIndex(h => String(h).toLowerCase() === "datadate");
  if (dataDateIdx < 0) return result;

  // Get symbol, expiration, and dataDate columns
  const data = sheet.getRange(2, 1, lastRow - 1, Math.max(3, dataDateIdx + 1)).getValues();

  for (const row of data) {
    const sym = String(row[0] || "").toUpperCase();
    const expRaw = row[1];
    const dataDateRaw = row[dataDateIdx];

    const expDate = parseDateAtMidnight_(expRaw);
    const dataDate = parseDateAtMidnight_(dataDateRaw);
    if (!expDate || !dataDate) continue;

    const expStr = expDate.getFullYear() + "-" +
      String(expDate.getMonth() + 1).padStart(2, "0") + "-" +
      String(expDate.getDate()).padStart(2, "0");

    const key = `${sym}|${expStr}`;
    const existing = result.get(key);
    if (!existing || dataDate > existing) {
      result.set(key, dataDate);
    }
  }

  return result;
}

/**
 * Deletes rows matching any of the given symbol/expiration combos.
 * @param {Sheet} sheet
 * @param {Set<string>} combos - Set of "SYMBOL|YYYY-MM-DD" keys
 */
function deleteRowsForCombos_(sheet, combos) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // symbol, expiration columns
  const rowsToDelete = [];

  for (let i = 0; i < data.length; i++) {
    const sym = String(data[i][0] || "").toUpperCase();
    const expRaw = data[i][1];
    const expDate = parseDateAtMidnight_(expRaw);
    if (!expDate) continue;

    const expStr = expDate.getFullYear() + "-" +
      String(expDate.getMonth() + 1).padStart(2, "0") + "-" +
      String(expDate.getDate()).padStart(2, "0");

    if (combos.has(`${sym}|${expStr}`)) {
      rowsToDelete.push(i + 2); // 1-based row number, +1 for header
    }
  }

  // Delete from bottom to top to preserve row numbers
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }
}
