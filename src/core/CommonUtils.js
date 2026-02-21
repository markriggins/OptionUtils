
/**
 * CommonUtils.js
 * Shared small helpers used across OptionUtils.
 *
 * IMPORTANT: Apps Script has a single global namespace across all .gs/.js files in a project.
 * Keep shared helpers here to avoid duplicate function definitions across files.
 */

/** Clamp a number between lo and hi. */
function clamp_(x, lo, hi) {
  return Math.max(lo, Math.min(hi, x));
}

/** Round to N decimal places. */
function roundTo_(n, digits) {
  const f = Math.pow(10, digits);
  return Math.round(Number(n) * f) / f;
}

// ---- TimeZone and Date Utilities ----
// Standard: Store dates as Date objects at midnight.
// Display: Format as M/D/YYYY using formatDateMDYYYY_().

/**
 * Gets the timezone for date formatting.
 * Uses the spreadsheet's timezone setting, defaulting to NYSE timezone.
 * @returns {string} IANA timezone identifier
 */
function getTimeZone_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return "America/New_York";
    return ss.getSpreadsheetTimeZone() || "America/New_York";
  } catch (e) {
    return "America/New_York";
  }
}

/**
 * Creates a Date at midnight local time.
 * Use this instead of new Date() constructor to avoid timezone issues.
 *
 * @param {number} year - Full year (e.g., 2026)
 * @param {number} month - Month 1-12 (NOT 0-indexed like JS Date)
 * @param {number} day - Day of month 1-31
 * @returns {Date} Date at midnight local time
 */
function createDate_(year, month, day) {
  return new Date(year, month - 1, day, 0, 0, 0, 0);
}

/**
 * Formats a Date as M/D/YYYY (standard display format).
 * Returns empty string for null/invalid dates.
 *
 * @param {Date|null} date - Date to format
 * @returns {string} Formatted date string or ""
 */
function formatDateMDYYYY_(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) return "";
  return (date.getMonth() + 1) + "/" + date.getDate() + "/" + date.getFullYear();
}

/**
 * Formats a Date using Utilities.formatDate with configured timezone.
 * For custom format patterns.
 *
 * @param {Date} date - Date to format
 * @param {string} pattern - Format pattern (e.g., "MMM d, yyyy")
 * @returns {string} Formatted date string or ""
 */
function formatDateWithTZ_(date, pattern) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) return "";
  return Utilities.formatDate(date, getTimeZone_(), pattern);
}

/**
 * Returns the first non-blank value from a range, scanning from bottom to top.
 * Useful for finding values in merged cells when called from a lower row.
 *
 * @param {Array} range - A vertical range (e.g., A$1:A3)
 * @return {*} First non-blank value found (bottom-up), or empty string if all blank
 * @customfunction
 */
function coalesce(range) {
  if (!Array.isArray(range)) {
    const v = (range ?? "").toString().trim();
    return v || "";
  }
  // Flatten 2D array (vertical range comes as [[a],[b],[c]])
  const flat = range.flat ? range.flat() : [].concat(...range);
  // Scan from end (bottom of range) toward start (top)
  for (let i = flat.length - 1; i >= 0; i--) {
    const v = flat[i];
    if (v != null && v !== "") return v;
  }
  return "";
}

/**
 * Formats option legs as a description string with negative prefixes for shorts.
 * Example: =formatLegsDescription(D2:D5, E2:E5) → "500/-600/740/-900"
 *
 * @param {Range} strikeRange - Range containing strike prices
 * @param {Range} qtyRange - Range containing quantities (negative = short)
 * @param {string} [suffix] - Optional suffix to append (e.g., "custom")
 * @return {string} Formatted description like "500/-600/740/-900 custom"
 * @customfunction
 */
function formatLegsDescription(strikeRange, qtyRange, suffix) {
  // Flatten inputs
  const strikes = Array.isArray(strikeRange) ? strikeRange.flat() : [strikeRange];
  const qtys = Array.isArray(qtyRange) ? qtyRange.flat() : [qtyRange];

  // Build legs array with strike and qty
  const legs = [];
  const n = Math.min(strikes.length, qtys.length);
  for (let i = 0; i < n; i++) {
    const strike = parseFloat(strikes[i]);
    const qty = parseFloat(qtys[i]);
    if (Number.isFinite(strike) && Number.isFinite(qty) && qty !== 0) {
      legs.push({ strike, qty });
    }
  }

  if (legs.length === 0) return "";

  // Sort by strike and format with negative prefix for shorts
  const formatted = legs
    .sort((a, b) => a.strike - b.strike)
    .map(l => l.qty < 0 ? `-${l.strike}` : `${l.strike}`)
    .join('/');

  return suffix ? `${formatted} ${suffix}` : formatted;
}

/**
 * Formats descriptions for all position groups in one call.
 * Place formula in first data row of Description column - it fills down automatically.
 *
 * @param {Range} groups - Group column (identifies group boundaries)
 * @param {Range} strikes - Strike column
 * @param {Range} qtys - Qty column
 * @param {Range} strategies - Strategy column
 * @return {string[][]} Description for first row of each group, blank for others
 * @customfunction
 */
function formatAllDescriptions(groups, strikes, qtys, strategies) {
  // Ensure 2D arrays
  const groupArr = Array.isArray(groups) ? groups : [[groups]];
  const strikeArr = Array.isArray(strikes) ? strikes : [[strikes]];
  const qtyArr = Array.isArray(qtys) ? qtys : [[qtys]];
  const stratArr = Array.isArray(strategies) ? strategies : [[strategies]];

  const numRows = groupArr.length;
  const result = [];

  let currentGroup = null;
  let groupStrikes = [];
  let groupQtys = [];
  let groupStrategy = "";
  let groupStartIdx = 0;

  for (let i = 0; i < numRows; i++) {
    const group = groupArr[i][0];
    const strike = strikeArr[i][0];
    const qty = qtyArr[i][0];
    const strategy = stratArr[i][0];

    // New group starts when group value changes (and is non-empty)
    if (group && group !== currentGroup) {
      // Output previous group's description
      if (currentGroup !== null && groupStrikes.length > 0) {
        result[groupStartIdx] = [formatLegsDescriptionCore_(groupStrikes, groupQtys, groupStrategy)];
      }
      // Start new group
      currentGroup = group;
      groupStrikes = [];
      groupQtys = [];
      groupStrategy = strategy || "";
      groupStartIdx = i;
    }

    // Accumulate strikes/qtys for current group
    const strikeNum = parseFloat(strike);
    const qtyNum = parseFloat(qty);
    if (Number.isFinite(strikeNum) && Number.isFinite(qtyNum) && qtyNum !== 0) {
      groupStrikes.push(strikeNum);
      groupQtys.push(qtyNum);
    }

    // Default to blank
    if (!result[i]) result[i] = [""];
  }

  // Don't forget last group
  if (currentGroup !== null && groupStrikes.length > 0) {
    result[groupStartIdx] = [formatLegsDescriptionCore_(groupStrikes, groupQtys, groupStrategy)];
  }

  return result;
}

/**
 * Core logic for formatting legs description.
 * @private
 */
function formatLegsDescriptionCore_(strikes, qtys, suffix) {
  const legs = [];
  const n = Math.min(strikes.length, qtys.length);
  for (let i = 0; i < n; i++) {
    const strike = strikes[i];
    const qty = qtys[i];
    if (Number.isFinite(strike) && Number.isFinite(qty) && qty !== 0) {
      legs.push({ strike, qty });
    }
  }

  if (legs.length === 0) return "";

  // Sort by strike and format with negative prefix for shorts
  const formatted = legs
    .sort((a, b) => a.strike - b.strike)
    .map(l => l.qty < 0 ? `-${l.strike}` : `${l.strike}`)
    .join('/');

  return suffix ? `${formatted} ${suffix}` : formatted;
}
