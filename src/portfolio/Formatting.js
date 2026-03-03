/**
 * Formatting.js
 * Portfolio description formatting functions.
 */

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
