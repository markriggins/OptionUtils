
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

// ---- Date Utilities ----
// Standard: Store dates as Date objects at midnight.
// Display: Format as M/D/YYYY using formatDateMDYYYY_().

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
