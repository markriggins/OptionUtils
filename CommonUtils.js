/**
 * CommonUtils.js
 * Shared small helpers used across OptionUtils.
 *
 * IMPORTANT: Apps Script has a single global namespace across all .gs/.js files in a project.
 * Keep shared helpers here to avoid duplicate function definitions across files.
 */

/** Round to 2 decimals. */
function round2_(n) {
  return Math.round(Number(n) * 100) / 100;
}
