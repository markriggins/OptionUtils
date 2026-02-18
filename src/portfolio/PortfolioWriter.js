/**
 * PortfolioWriter.js
 * Sheet output functions for portfolio data.
 *
 * Writes spread positions to the Portfolio sheet with formulas,
 * formatting, and summary calculations.
 */

/**
 * Formats a spread for display label.
 */
function formatSpreadLabel_(spread) {
  if (spread.type === "cash") {
    return `Cash $${spread.price}`;
  }
  if (spread.type === "stock") {
    return `${spread.ticker} Stock`;
  }
  if (spread.type === "iron-condor") {
    const strikes = spread.legs.map(l => l.strike).join("/");
    return `${spread.ticker} ${formatExpirationShort_(spread.expiration)} ${strikes} iron-condor`;
  }
  const strikes = [spread.lowerStrike, spread.upperStrike].filter(s => s).join("/");
  const strategyType = spread.lowerStrike && spread.upperStrike ? "bull-call-spread" :
                       spread.lowerStrike ? "long-call" : "short-call";
  return `${spread.ticker} ${formatExpirationShort_(spread.expiration)} ${strikes} ${strategyType}`;
}

/**
 * Formats an existing position for display in the report.
 */
function formatPositionLabel_(pos) {
  if (!pos.legs || pos.legs.length === 0) return "Unknown position";
  const leg = pos.legs[0];
  const strikes = pos.legs.map(l => l.strike).filter(s => s).sort((a, b) => a - b).join("/");
  const exp = formatExpirationShort_(leg.expiration);
  const debug = pos.debugReason || "";
  return `${leg.symbol} ${exp} ${strikes} ${leg.type || "Call"}${debug}`;
}

/**
 * Formats expiration as "Mon YYYY" for display.
 */
function formatExpirationShort_(exp) {
  if (!exp) return "";
  const d = parseDateAtMidnight_(exp);
  if (!d) return String(exp);
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return `${months[d.getMonth()]} ${d.getFullYear()}`;
}

/**
 * Formats a date as MM/DD/YYYY (e.g., 2/6/2026).
 */
function formatDateLong_(dateVal) {
  if (!dateVal) return "";
  const d = parseDateAtMidnight_(dateVal);
  if (!d) return String(dateVal);
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}

/**
 * Generates a description for a spread like "500/600 bull-call-spread".
 */
function generateSpreadDescription_(spread) {
  if (spread.type === "stock") {
    return "Stock";
  }

  if (spread.type === "cash") {
    return "Cash";
  }

  if (spread.type === "iron-condor" && spread.legs) {
    const strikes = spread.legs.map(l => l.strike).sort((a, b) => a - b).join("/");
    const sortedStrikes = spread.legs.map(l => l.strike).sort((a, b) => a - b);
    const isButterfly = sortedStrikes[1] === sortedStrikes[2];
    return `${strikes} ${isButterfly ? "iron-butterfly" : "iron-condor"}`;
  }

  // Custom multi-leg position
  if (spread.type === "custom" && spread.legs) {
    const formattedStrikes = spread.legs
      .sort((a, b) => a.strike - b.strike)
      .map(l => l.qty < 0 ? `-${l.strike}` : `${l.strike}`)
      .join("/");
    return `${formattedStrikes} custom`;
  }

  // Straddle/strangle (2 legs: call + put)
  if (spread.legs && spread.legs.length === 2) {
    const strikes = spread.legs.map(l => l.strike).sort((a, b) => a - b);
    const strikeStr = strikes[0] === strikes[1] ? String(strikes[0]) : strikes.join("/");
    return `${strikeStr} ${spread.type}`;
  }

  // Regular spread
  const strikes = [spread.lowerStrike, spread.upperStrike].filter(s => s != null);
  const strikeStr = strikes.sort((a, b) => a - b).join("/");

  let strategy;
  if (spread.lowerStrike && spread.upperStrike) {
    if (spread.optionType === "Call") {
      strategy = "bull-call-spread";
    } else {
      strategy = "bull-put-spread";
    }
  } else if (spread.lowerStrike) {
    strategy = spread.optionType === "Call" ? "long-call" : "long-put";
  } else {
    strategy = spread.optionType === "Call" ? "short-call" : "short-put";
  }

  return `${strikeStr} ${strategy}`;
}

/**
 * Finds custom groups in a Portfolio sheet (groups that don't match standard strategies).
 * Returns array of { group, description } for each custom group found.
 */
function findCustomGroups_(sheet) {
  const customGroups = [];
  const range = sheet.getDataRange();
  const rows = range.getValues();
  if (rows.length < 2) return customGroups;

  const headers = rows[0];
  const idxSym = findColumn_(headers, ["symbol", "ticker"]);
  const idxGroup = findColumn_(headers, ["group", "grp"]);
  const idxDesc = findColumn_(headers, ["description", "desc"]);
  const idxStrike = findColumn_(headers, ["strike", "strikeprice"]);
  const idxType = findColumn_(headers, ["type", "optiontype", "callput"]);
  const idxQty = findColumn_(headers, ["qty", "quantity"]);

  const groups = new Map();
  let lastSym = "", lastGroup = "", lastDesc = "";

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const rawSym = idxSym >= 0 ? String(row[idxSym] || "").trim().toUpperCase() : "";
    if (rawSym) lastSym = rawSym;
    const rawGroup = idxGroup >= 0 ? String(row[idxGroup] || "").trim() : "";
    if (rawGroup) lastGroup = rawGroup;
    const rawDesc = idxDesc >= 0 ? String(row[idxDesc] || "").trim() : "";
    if (rawDesc) lastDesc = rawDesc;

    const groupKey = `${lastSym}|${lastGroup}`;
    if (!groups.has(groupKey)) {
      groups.set(groupKey, { legs: [], description: lastDesc, group: lastGroup });
    }

    const strike = idxStrike >= 0 ? parseNumber_(row[idxStrike]) : NaN;
    const type = idxType >= 0 ? parseOptionType_(row[idxType]) : null;
    const qty = idxQty >= 0 ? parseNumber_(row[idxQty]) : 0;

    if (Number.isFinite(qty) && qty !== 0) {
      groups.get(groupKey).legs.push({ strike, type, qty });
    }
  }

  for (const [key, g] of groups) {
    if (g.legs.length === 0) continue;
    const posType = detectPositionType_(g.legs);
    if (posType === null) {
      let desc = g.description;
      if (!desc) {
        const strikes = g.legs
          .filter(l => Number.isFinite(l.strike))
          .sort((a, b) => a.strike - b.strike)
          .map(l => l.qty < 0 ? `-${l.strike}` : `${l.strike}`)
          .join('/');
        desc = `${strikes} custom`;
      }
      customGroups.push({ group: g.group, description: desc });
    }
  }

  return customGroups;
}

/**
 * Parses existing Portfolio table into position objects.
 */
function parsePortfolioTable_(rows) {
  if (rows.length < 2) return [];

  const headers = rows[0];
  const idxSym = findColumn_(headers, ["symbol", "ticker"]);
  const idxGroup = findColumn_(headers, ["group", "grp"]);
  const idxStrike = findColumn_(headers, ["strike", "strikeprice"]);
  const idxType = findColumn_(headers, ["type", "optiontype", "callput"]);
  const idxExp = findColumn_(headers, ["expiration", "exp"]);
  const idxQty = findColumn_(headers, ["qty", "quantity"]);
  const idxPrice = findColumn_(headers, ["price", "cost"]);
  const idxLastTxnDate = findColumn_(headers, ["lasttxndate", "last txn date"]);

  const positions = new Map();

  let lastSym = "";
  let lastGroup = "";
  let currentLegs = [];
  let currentLastTxnDate = "";

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];

    const rawSym = idxSym >= 0 ? String(row[idxSym] || "").trim().toUpperCase() : "";
    if (rawSym) lastSym = rawSym;

    const rawGroup = idxGroup >= 0 ? String(row[idxGroup] || "").trim() : "";
    if (rawGroup && rawGroup !== lastGroup) {
      if (currentLegs.length > 0) {
        const key = makeSpreadKey_(currentLegs);
        if (key) positions.set(key, { legs: currentLegs, groupNum: lastGroup, lastTxnDate: currentLastTxnDate });
      }
      lastGroup = rawGroup;
      currentLegs = [];
      if (idxLastTxnDate >= 0) {
        const rawDate = row[idxLastTxnDate];
        currentLastTxnDate = rawDate instanceof Date
          ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "MM/dd/yy")
          : String(rawDate || "").trim();
      } else {
        currentLastTxnDate = "";
      }
    }

    const strike = idxStrike >= 0 ? parseNumber_(row[idxStrike]) : NaN;
    const type = idxType >= 0 ? parseOptionType_(row[idxType]) : null;
    const exp = idxExp >= 0 ? row[idxExp] : "";
    const qty = idxQty >= 0 ? parseNumber_(row[idxQty]) : NaN;
    const price = idxPrice >= 0 ? parseNumber_(row[idxPrice]) : NaN;

    const isStock = type === "Stock" || (!Number.isFinite(strike) && Number.isFinite(qty) && !type);
    if ((Number.isFinite(strike) || isStock) && Number.isFinite(qty)) {
      currentLegs.push({
        symbol: lastSym,
        strike: Number.isFinite(strike) ? strike : null,
        type: isStock ? "Stock" : type,
        expiration: exp,
        qty,
        price,
        row: r,
      });
    }
  }

  if (currentLegs.length > 0) {
    const key = makeSpreadKey_(currentLegs);
    if (key) positions.set(key, { legs: currentLegs, groupNum: lastGroup, lastTxnDate: currentLastTxnDate });
  }

  return positions;
}

/**
 * Validates spread positions against E*Trade portfolio option quantities.
 * Returns a list of discrepancies.
 */
function validateOptionQuantities_(spreads, portfolioOptions) {
  const expected = new Map();

  for (const spread of spreads) {
    if (spread.type === "stock" || spread.type === "cash") continue;

    const ticker = spread.ticker;
    const expiration = formatExpirationForKey_(spread.expiration);

    if (spread.type === "iron-condor" && spread.legs) {
      for (const leg of spread.legs) {
        const key = `${ticker}|${expiration}|${leg.strike}|${leg.optionType}`;
        expected.set(key, (expected.get(key) || 0) + leg.qty);
      }
    } else {
      if (spread.lowerStrike != null) {
        const key = `${ticker}|${expiration}|${spread.lowerStrike}|${spread.optionType}`;
        expected.set(key, (expected.get(key) || 0) + spread.qty);
      }
      if (spread.upperStrike != null) {
        const key = `${ticker}|${expiration}|${spread.upperStrike}|${spread.optionType}`;
        const shortQty = spread.lowerStrike != null ? -spread.qty : spread.qty;
        expected.set(key, (expected.get(key) || 0) + shortQty);
      }
    }
  }

  const mismatches = [];
  const missing = [];
  const extra = [];

  for (const [key, expectedQty] of expected) {
    const actualQty = portfolioOptions.get(key) || 0;
    if (actualQty !== expectedQty) {
      const [ticker, exp, strike, type] = key.split("|");
      mismatches.push({
        key,
        ticker,
        expiration: exp,
        strike: parseFloat(strike),
        type,
        expected: expectedQty,
        actual: actualQty,
      });
    }
  }

  for (const [key, actualQty] of portfolioOptions) {
    if (!expected.has(key)) {
      const [ticker, exp, strike, type] = key.split("|");
      extra.push({
        key,
        ticker,
        expiration: exp,
        strike: parseFloat(strike),
        type,
        qty: actualQty,
      });
    }
  }

  return { mismatches, missing, extra };
}

/**
 * Writes the Portfolio table back to the sheet.
 * @param {Spreadsheet} ss - The active spreadsheet
 * @param {string[]} headers - Column headers
 * @param {Object[]} updatedLegs - Existing positions to update
 * @param {Object[]} newLegs - New positions to append
 * @param {Map} [closingPrices] - Map of leg keys to closing prices
 */
function writePortfolioTable_(ss, headers, updatedLegs, newLegs, closingPrices) {
  closingPrices = closingPrices || new Map();

  const legsRange = getNamedRangeWithTableFallback_(ss, "Portfolio");
  if (!legsRange || headers.length === 0) {
    log.info("import", "Portfolio table not found, creating new one");
    let sheet = ss.getSheetByName("Portfolio");
    if (!sheet) sheet = ss.insertSheet("Portfolio");

    headers = ["Symbol", "Group", "Description", "Strategy", "Strike", "Type", "Expiration", "Qty", "Price", "Investment", "Rec Close", "Closed", "Gain", "Current Value", "LastTxnDate", "Link"];
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground("#93c47d");
    headerRange.setFontWeight("bold");
    ss.setNamedRange("PortfolioTable", sheet.getRange("A:P"));

    const filterRange = sheet.getRange("A:P");
    filterRange.createFilter();

    sheet.getRange("A:P").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  const range = getNamedRangeWithTableFallback_(ss, "Portfolio");
  const sheet = range.getSheet();
  const startRow = range.getRow();
  const startCol = range.getColumn();

  // Find column indexes
  const idxSym = findColumn_(headers, ["symbol", "ticker"]);
  const idxGroup = findColumn_(headers, ["group", "grp"]);
  const idxDescription = findColumn_(headers, ["description", "desc"]);
  const idxStrategy = findColumn_(headers, ["strategy"]);
  const idxStrike = findColumn_(headers, ["strike"]);
  const idxType = findColumn_(headers, ["type"]);
  const idxExp = findColumn_(headers, ["expiration", "exp"]);
  const idxQty = findColumn_(headers, ["qty", "quantity"]);
  const idxPrice = findColumn_(headers, ["price"]);
  const idxInvestment = findColumn_(headers, ["investment"]);
  const idxRecClose = findColumn_(headers, ["recclose", "rec close"]);
  const idxClosed = findColumn_(headers, ["closed", "actualclose", "closedat"]);
  const idxGain = findColumn_(headers, ["gain"]);
  const idxCurrentValue = findColumn_(headers, ["currentvalue", "current value", "currvalue"]);
  const idxLastTxnDate = findColumn_(headers, ["lasttxndate", "last txn date"]);
  const idxLink = findColumn_(headers, ["link"]);

  // Column letters for formulas
  const colLetter = (idx) => String.fromCharCode(65 + idx);
  const symCol = idxSym >= 0 ? colLetter(idxSym) : "A";
  const stratCol = idxStrategy >= 0 ? colLetter(idxStrategy) : "D";
  const strikeCol = idxStrike >= 0 ? colLetter(idxStrike) : "E";
  const typeCol = idxType >= 0 ? colLetter(idxType) : "F";
  const expCol = idxExp >= 0 ? colLetter(idxExp) : "G";
  const qtyCol = idxQty >= 0 ? colLetter(idxQty) : "H";
  const priceCol = idxPrice >= 0 ? colLetter(idxPrice) : "I";
  const recCloseCol = idxRecClose >= 0 ? colLetter(idxRecClose) : "K";
  const closedCol = idxClosed >= 0 ? colLetter(idxClosed) : "L";

  // Update existing rows
  for (const pos of updatedLegs) {
    for (const leg of pos.legs) {
      if (leg.row != null) {
        const rowNum = startRow + leg.row;
        if (idxQty >= 0) sheet.getRange(rowNum, startCol + idxQty).setValue(leg.qty);
        if (idxPrice >= 0) sheet.getRange(rowNum, startCol + idxPrice).setValue(roundTo_(leg.price, 2));

        if (idxClosed >= 0 && leg.symbol && leg.expiration && leg.strike && leg.type && leg.type !== "Stock") {
          const existingVal = sheet.getRange(rowNum, startCol + idxClosed).getValue();
          if (existingVal === "" || existingVal == null) {
            const expStr = formatExpirationForKey_(leg.expiration);
            const key = `${leg.symbol}|${expStr}|${leg.strike}|${leg.type}`;
            const closePrice = closingPrices.get(key);
            if (closePrice != null) {
              sheet.getRange(rowNum, startCol + idxClosed).setValue(closePrice);
            }
          }
        }
      }
    }
    if (idxLastTxnDate >= 0 && pos.lastTxnDate && pos.legs.length > 0 && pos.legs[0].row != null) {
      sheet.getRange(startRow + pos.legs[0].row, startCol + idxLastTxnDate).setValue(formatDateLong_(pos.lastTxnDate));
    }
  }

  // Append new spreads
  if (newLegs.length > 0) {
    let lastRow = sheet.getLastRow();
    let nextGroup = 1;

    let lastDataRow = startRow;
    if (idxGroup >= 0 && lastRow > startRow) {
      const groupData = sheet.getRange(startRow + 1, startCol + idxGroup, lastRow - startRow, 1).getValues();
      for (let i = 0; i < groupData.length; i++) {
        const g = parseInt(groupData[i][0], 10);
        if (Number.isFinite(g)) {
          if (g >= nextGroup) nextGroup = g + 1;
          lastDataRow = startRow + 1 + i;
        }
      }
    }

    if (lastRow > lastDataRow) {
      sheet.deleteRows(lastDataRow + 1, lastRow - lastDataRow);
    }

    lastRow = lastDataRow;

    for (const spread of newLegs) {
      const rows = [];

      const getClosingPrice = (ticker, expiration, strike, optionType) => {
        const key = `${ticker}|${expiration}|${strike}|${optionType}`;
        const val = closingPrices.get(key);
        return val != null ? val : "";
      };

      const spreadDescription = generateSpreadDescription_(spread);

      if (spread.type === "stock") {
        const row = new Array(headers.length).fill("");
        if (idxSym >= 0) row[idxSym] = spread.ticker;
        if (idxGroup >= 0) row[idxGroup] = nextGroup;
        if (idxDescription >= 0) row[idxDescription] = spreadDescription;
        if (idxType >= 0) row[idxType] = "Stock";
        if (idxQty >= 0) row[idxQty] = spread.qty;
        if (idxPrice >= 0) row[idxPrice] = roundTo_(spread.price, 2);
        rows.push(row);
      }
      else if (spread.type === "cash") {
        const row = new Array(headers.length).fill("");
        if (idxSym >= 0) row[idxSym] = "CASH";
        if (idxGroup >= 0) row[idxGroup] = nextGroup;
        if (idxDescription >= 0) row[idxDescription] = "Cash";
        if (idxStrategy >= 0) row[idxStrategy] = "Cash";
        if (idxType >= 0) row[idxType] = "Cash";
        if (idxQty >= 0) row[idxQty] = 1;
        if (idxPrice >= 0) row[idxPrice] = roundTo_(spread.price, 2);
        rows.push(row);
      }
      else if (spread.legs && spread.legs.length > 0) {
        for (let i = 0; i < spread.legs.length; i++) {
          const leg = spread.legs[i];
          const row = new Array(headers.length).fill("");
          if (i === 0) {
            if (idxSym >= 0) row[idxSym] = spread.ticker;
            if (idxGroup >= 0) row[idxGroup] = nextGroup;
          }
          if (idxStrike >= 0) row[idxStrike] = leg.strike;
          if (idxType >= 0) row[idxType] = leg.optionType;
          if (idxExp >= 0) row[idxExp] = formatExpirationForKey_(spread.expiration);
          if (idxQty >= 0) row[idxQty] = leg.qty;
          if (idxPrice >= 0) row[idxPrice] = roundTo_(leg.price, 2);
          if (idxClosed >= 0) row[idxClosed] = getClosingPrice(spread.ticker, spread.expiration, leg.strike, leg.optionType);
          rows.push(row);
        }
      } else {
        const hasLong = spread.lowerStrike != null;
        const hasShort = spread.upperStrike != null;
        if (!hasLong && !hasShort) continue;
        const isFirstRow = [true];
        const normalizedExp = formatExpirationForKey_(spread.expiration);

        if (hasLong) {
          const row = new Array(headers.length).fill("");
          if (idxSym >= 0) row[idxSym] = spread.ticker;
          if (idxGroup >= 0) row[idxGroup] = nextGroup;
          if (idxDescription >= 0) row[idxDescription] = spreadDescription;
          if (idxStrike >= 0) row[idxStrike] = spread.lowerStrike;
          if (idxType >= 0) row[idxType] = spread.optionType;
          if (idxExp >= 0) row[idxExp] = normalizedExp;
          if (idxQty >= 0) row[idxQty] = spread.qty;
          if (idxPrice >= 0) row[idxPrice] = roundTo_(spread.lowerPrice, 2);
          if (idxClosed >= 0) row[idxClosed] = getClosingPrice(spread.ticker, spread.expiration, spread.lowerStrike, spread.optionType);
          rows.push(row);
          isFirstRow[0] = false;
        }

        if (hasShort) {
          const row = new Array(headers.length).fill("");
          if (isFirstRow[0]) {
            if (idxSym >= 0) row[idxSym] = spread.ticker;
            if (idxGroup >= 0) row[idxGroup] = nextGroup;
            if (idxDescription >= 0) row[idxDescription] = spreadDescription;
          }
          if (idxStrike >= 0) row[idxStrike] = spread.upperStrike;
          if (idxType >= 0) row[idxType] = spread.optionType;
          if (idxExp >= 0) row[idxExp] = normalizedExp;
          if (idxQty >= 0) row[idxQty] = hasLong ? -spread.qty : spread.qty;
          if (idxPrice >= 0) row[idxPrice] = roundTo_(spread.upperPrice, 2);
          if (idxClosed >= 0) row[idxClosed] = getClosingPrice(spread.ticker, spread.expiration, spread.upperStrike, spread.optionType);
          rows.push(row);
        }
      }

      if (rows.length === 0) continue;

      if (idxLastTxnDate >= 0 && spread.date) {
        rows[0][idxLastTxnDate] = formatDateLong_(spread.date);
      }

      const firstRow = lastRow + 1;
      const lastLegRow = firstRow + rows.length - 1;

      sheet.getRange(firstRow, startCol, rows.length, headers.length).setValues(rows);

      const isStock = spread.type === "stock";
      const isCash = spread.type === "cash";

      // Strategy formula
      if (idxStrategy >= 0) {
        if (isStock) {
          sheet.getRange(firstRow, startCol + idxStrategy).setValue("Stock");
        } else if (isCash) {
          sheet.getRange(firstRow, startCol + idxStrategy).setValue("Cash");
        } else {
          const formula = `=detectStrategy($${strikeCol}${firstRow}:$${strikeCol}${lastLegRow}, $${typeCol}${firstRow}:$${typeCol}${lastLegRow}, $${qtyCol}${firstRow}:$${qtyCol}${lastLegRow})`;
          sheet.getRange(firstRow, startCol + idxStrategy).setFormula(formula);
        }
      }

      // Description formula
      if (idxDescription >= 0) {
        if (isStock) {
          sheet.getRange(firstRow, startCol + idxDescription).setValue("Stock");
        } else if (isCash) {
          sheet.getRange(firstRow, startCol + idxDescription).setValue("Cash");
        } else {
          const formula = `=formatLegsDescription($${strikeCol}${firstRow}:$${strikeCol}${lastLegRow}, $${qtyCol}${firstRow}:$${qtyCol}${lastLegRow}, $${stratCol}${firstRow})`;
          sheet.getRange(firstRow, startCol + idxDescription).setFormula(formula);
        }
      }

      // Investment formula
      if (idxInvestment >= 0) {
        if (isCash) {
          sheet.getRange(firstRow, startCol + idxInvestment).setFormula(`=$${priceCol}${firstRow}`);
        } else {
          const multiplier = isStock ? "" : " * 100";
          const formula = `=SUMPRODUCT($${qtyCol}${firstRow}:$${qtyCol}${lastLegRow}, $${priceCol}${firstRow}:$${priceCol}${lastLegRow})${multiplier}`;
          sheet.getRange(firstRow, startCol + idxInvestment).setFormula(formula);
        }
      }

      // Gain formula
      if (idxGain >= 0 && idxRecClose >= 0) {
        if (isCash) {
          sheet.getRange(firstRow, startCol + idxGain).setValue(0);
        } else {
          let closeRef;
          if (idxClosed >= 0) {
            closeRef = `IF($${closedCol}${firstRow}:$${closedCol}${lastLegRow}<>"", $${closedCol}${firstRow}:$${closedCol}${lastLegRow}, $${recCloseCol}${firstRow}:$${recCloseCol}${lastLegRow})`;
          } else {
            closeRef = `$${recCloseCol}${firstRow}:$${recCloseCol}${lastLegRow}`;
          }
          const multiplier = isStock ? "" : " * 100";
          const formula = `=SUMPRODUCT($${qtyCol}${firstRow}:$${qtyCol}${lastLegRow}, ${closeRef} - $${priceCol}${firstRow}:$${priceCol}${lastLegRow})${multiplier}`;
          sheet.getRange(firstRow, startCol + idxGain).setFormula(formula);
        }
      }

      // Current Value formula
      if (idxCurrentValue >= 0 && idxInvestment >= 0 && idxGain >= 0) {
        if (isCash) {
          const invCell = colLetter(idxInvestment) + firstRow;
          sheet.getRange(firstRow, startCol + idxCurrentValue).setFormula(`=$${invCell}`);
        } else {
          const invCell = colLetter(idxInvestment) + firstRow;
          const gainCell = colLetter(idxGain) + firstRow;
          let formula;
          if (idxClosed >= 0) {
            formula = `=IF($${gainCell}="", "", IF(COUNTBLANK($${closedCol}${firstRow}:$${closedCol}${lastLegRow})=0, 0, $${invCell}+$${gainCell}))`;
          } else {
            formula = `=IF($${gainCell}="", "", $${invCell}+$${gainCell})`;
          }
          sheet.getRange(firstRow, startCol + idxCurrentValue).setFormula(formula);
        }
      }

      // Link formula
      if (idxLink >= 0 && !isStock && !isCash) {
        const urlFormula = `buildOptionStratUrlFromLegs($${symCol}$1:$${symCol}${firstRow}, $${strikeCol}${firstRow}:$${strikeCol}${lastLegRow}, $${typeCol}${firstRow}:$${typeCol}${lastLegRow}, $${expCol}${firstRow}:$${expCol}${lastLegRow}, $${qtyCol}${firstRow}:$${qtyCol}${lastLegRow}, $${priceCol}${firstRow}:$${priceCol}${lastLegRow})`;
        const formula = `=HYPERLINK(${urlFormula}, "OptionStrat")`;
        sheet.getRange(firstRow, startCol + idxLink).setFormula(formula);
      }

      // Rec Close formula for each leg
      if (idxRecClose >= 0) {
        for (let i = 0; i < rows.length; i++) {
          const legRow = firstRow + i;
          const hasClosed = idxClosed >= 0 && rows[i][idxClosed] !== "";
          if (!hasClosed) {
            if (isCash) {
              sheet.getRange(legRow, startCol + idxRecClose).setFormula(`=$${priceCol}${legRow}`);
            } else if (isStock) {
              const formula = `=GOOGLEFINANCE("${spread.ticker}")`;
              sheet.getRange(legRow, startCol + idxRecClose).setFormula(formula);
            } else {
              const formula = `=recommendClose($${symCol}$1:$${symCol}${legRow}, $${expCol}${legRow}, $${strikeCol}${legRow}, $${typeCol}${legRow}, $${qtyCol}${legRow}, 60)`;
              sheet.getRange(legRow, startCol + idxRecClose).setFormula(formula);
            }
          }
        }
      }

      const allClosed = idxClosed >= 0 && rows.every(r => r[idxClosed] !== "");

      const bgColor = (nextGroup % 2 === 1) ? "#fff2cc" : "#ffffff";
      const groupRange = sheet.getRange(firstRow, startCol, rows.length, headers.length);
      groupRange.setBackground(bgColor);

      if (allClosed) {
        groupRange.setFontColor("#999999");
      }

      lastRow = lastLegRow;
      nextGroup++;
    }

    // Write summary rows
    const summaryStart = lastRow + 2;
    const invCol = idxInvestment >= 0 ? colLetter(idxInvestment) : "I";
    const gainCol = idxGain >= 0 ? colLetter(idxGain) : "L";
    const dr = (col) => `$${col}$2:$${col}$${lastRow}`;

    sheet.getRange(summaryStart, startCol).setValue("Realized").setFontWeight("bold");
    sheet.getRange(summaryStart, startCol + idxGain)
      .setFormula(`=SUMPRODUCT((${dr(closedCol)}<>"")*${dr(gainCol)})`)
      .setFontWeight("bold");

    sheet.getRange(summaryStart + 1, startCol).setValue("Unrealized").setFontWeight("bold");
    sheet.getRange(summaryStart + 1, startCol + idxGain)
      .setFormula(`=SUMPRODUCT((${dr(closedCol)}="")*${dr(gainCol)})`)
      .setFontWeight("bold");

    sheet.getRange(summaryStart + 2, startCol).setValue("Total").setFontWeight("bold");
    sheet.getRange(summaryStart + 2, startCol + idxInvestment)
      .setFormula(`=SUMPRODUCT((${dr(closedCol)}="")*${dr(invCol)})`)
      .setFontWeight("bold");
    sheet.getRange(summaryStart + 2, startCol + idxGain)
      .setFormula(`=$${gainCol}$${summaryStart}+$${gainCol}$${summaryStart + 1}`)
      .setFontWeight("bold");

    if (idxCurrentValue >= 0) {
      const currValCol = colLetter(idxCurrentValue);
      sheet.getRange(summaryStart + 2, startCol + idxCurrentValue)
        .setFormula(`=SUM(${dr(currValCol)})`)
        .setFontWeight("bold");
    }

    const summaryRange = sheet.getRange(summaryStart, startCol, 3, headers.length);
    summaryRange.setBackground("#d9ead3");
  }

  // Apply number formats
  const lastDataRow = sheet.getLastRow();
  if (lastDataRow >= 2) {
    const dataRowCount = lastDataRow - 1;
    const fmtCols = [
      { idx: idxQty, fmt: "#,##0" },
      { idx: idxPrice, fmt: "#,##0.00" },
      { idx: idxInvestment, fmt: "#,##0.00" },
      { idx: idxRecClose, fmt: "#,##0.00" },
      { idx: idxClosed, fmt: "#,##0.00" },
      { idx: idxGain, fmt: "#,##0.00" },
      { idx: idxCurrentValue, fmt: "#,##0.00" },
      { idx: idxLastTxnDate, fmt: "mm/dd/yy" },
    ];
    for (const { idx, fmt } of fmtCols) {
      if (idx >= 0) {
        sheet.getRange(2, startCol + idx, dataRowCount, 1).setNumberFormat(fmt);
      }
    }

    const headerRange = sheet.getRange(1, startCol, 1, headers.length);
    const headerValues = headerRange.getValues()[0];

    headerRange.clearContent();
    for (let c = 0; c < headers.length; c++) {
      sheet.autoResizeColumn(startCol + c);
    }
    headerRange.setValues([headerValues]);

    headerRange.setWrap(true).setVerticalAlignment("bottom");

    for (let c = 0; c < headers.length; c++) {
      const width = sheet.getColumnWidth(startCol + c);
      if (width < 50) {
        sheet.setColumnWidth(startCol + c, 50);
      }
    }
  }
}
