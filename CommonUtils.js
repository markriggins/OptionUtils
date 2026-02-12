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
