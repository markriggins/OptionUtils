/**
 * CommonUtils.js
 * Shared small helpers used across OptionUtils.
 *
 * IMPORTANT: Apps Script has a single global namespace across all .gs/.js files in a project.
 * Keep shared helpers here to avoid duplicate function definitions across files.
 */

/**
 * Returns the Unix timestamp (ms) for 00:00:00 on the same calendar day.
 * @param {Date} d - Input date (any time of day)
 * @returns {number|null} Milliseconds since epoch, or null if invalid
 */
function dateToMidnightTimestamp(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}

/** Check if two Google Sheets ranges refer to the same cells. */
function rangesEqual_(a, b) {
  return (
    a.getSheet().getSheetId() === b.getSheet().getSheetId() &&
    a.getRow() === b.getRow() &&
    a.getColumn() === b.getColumn() &&
    a.getNumRows() === b.getNumRows() &&
    a.getNumColumns() === b.getNumColumns()
  );
}

/** Check if two Google Sheets ranges overlap. */
function rangesIntersect_(a, b) {
  if (a.getSheet().getSheetId() !== b.getSheet().getSheetId()) return false;

  const aR1 = a.getRow(), aC1 = a.getColumn();
  const aR2 = aR1 + a.getNumRows() - 1;
  const aC2 = aC1 + a.getNumColumns() - 1;

  const bR1 = b.getRow(), bC1 = b.getColumn();
  const bR2 = bR1 + b.getNumRows() - 1;
  const bC2 = bC1 + b.getNumColumns() - 1;

  return aR1 <= bR2 && aR2 >= bR1 && aC1 <= bC2 && aC2 >= bC1;
}

/** Clamp a number between lo and hi. */
function clamp_(x, lo, hi) {
  return Math.max(lo, Math.min(hi, x));
}

/** Round to N decimal places. */
function roundTo_(n, digits) {
  const f = Math.pow(10, digits);
  return Math.round(Number(n) * f) / f;
}
