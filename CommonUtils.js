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
