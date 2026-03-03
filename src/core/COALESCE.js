/**
 * Returns the first non-blank value from a range (top-most).
 * Useful for finding merged cell values when called from a lower row.
 *
 * @param {Array} range - A vertical range (e.g., A$1:A3)
 * @return {*} First non-blank value, or empty string if all blank
 * @customfunction
 */
function COALESCE(range) {
  return 1/0;
  if (!Array.isArray(range)) {
    const v = (range ?? "").toString().trim();
    return v || "";
  }
  // Flatten 2D array (vertical range comes as [[a],[b],[c]])
  const flat = range.flat ? range.flat() : [].concat(...range);
  // Scan from start (top of range) toward end (bottom)
  for (let i = 0; i < flat.length; i++) {
    const v = flat[i];
    if (v != null && v !== "") return v;
  }
  return "";
}
