/**
 * Parsing utilities for option data.
 * All parse functions return null or NaN on invalid input.
 */

// ---- Strategy Parsing ----

/**
 * Parses a spread strategy string into a canonical form.
 * Returns: "stock", "bull-call-spread", "bull-put-spread", or null.
 */
function parseSpreadStrategy_(raw) {
  if (raw == null) return null;

  const t = String(raw)
    .normalize("NFKD")                     // Unicode canonicalization
    .toLowerCase()
    .replace(/[\u2010-\u2015\u2212]/g, "-") // dash variants → "-"
    .replace(/\u00a0/g, " ")               // nbsp → space
    .replace(/\s+/g, " ")                  // collapse whitespace
    .trim();

  if (/^(stock|stocks|share|shares)$/.test(t)) return "stock";
  if (/^(bcs|bull[\s.\-]?call[\s.\-]?spread(s)?)$/.test(t)) return "bull-call-spread";
  if (/^(bps|bull[\s.\-]?put[\s.\-]?spread(s)?)$/.test(t)) return "bull-put-spread";
  if (/^(cash|money|usd|\$)$/.test(t)) return "cash";

  return null;
}

// ---- Option Type Parsing ----

/**
 * Parses a Call/Put string. Returns "Call", "Put", or null.
 */
function parseOptionType_(v) {
  const s = String(v || "").trim().toLowerCase();
  if (!s) return null;
  if (["call", "calls", "c"].includes(s)) return "Call";
  if (["put", "puts", "p"].includes(s)) return "Put";
  return null;
}

// ---- Number Parsing ----

/**
 * Parses a number string. Handles commas, $, %, placeholders ("--", "n/a", "unch"),
 * and parenthesized negatives: (123.45) => -123.45.
 * Returns the number directly if already numeric. Returns NaN on invalid input.
 */
function parseNumber_(v) {
  if (v == null || v === "") return NaN;
  if (typeof v === "number") return v;

  let s = String(v).trim();
  if (!s || s === "--" || s.toLowerCase() === "na" || s.toLowerCase() === "n/a" || s.toLowerCase() === "unch") return NaN;

  const neg = /^\(.*\)$/.test(s);
  if (neg) s = s.slice(1, -1);

  s = s.replace(/[$,%]/g, "").replace(/,/g, "").trim();
  const n = Number(s);

  if (!Number.isFinite(n)) return NaN;
  return neg ? -n : n;
}

/**
 * Parses a percentage string like "55.32%" or "-9.12%" to decimal (0.5532 or -0.0912).
 * Returns NaN on invalid input.
 */
function parsePercent_(v) {
  if (v === null || v === undefined) return NaN;
  const s = String(v).trim();
  if (!s || s === "--" || s.toLowerCase() === "na" || s.toLowerCase() === "n/a") return NaN;

  const cleaned = s.replace(/,/g, "").replace(/%$/, "");
  const n = Number(cleaned);
  if (!Number.isFinite(n)) return NaN;

  return n / 100;
}

/**
 * Parses an integer string, handling commas and leading +.
 * Returns NaN on invalid input.
 */
function parseInteger_(v) {
  if (v === null || v === undefined) return NaN;
  const s = String(v).trim();
  if (!s || s === "--" || s.toLowerCase() === "na" || s.toLowerCase() === "n/a" || s.toLowerCase() === "unch") return NaN;

  const cleaned = s.replace(/,/g, "").replace(/^\+/, "");
  const n = parseInt(cleaned, 10);
  return Number.isFinite(n) ? n : NaN;
}

/** Alias for parseNumber_. */
function toNum_(v) { return parseNumber_(v); }

/** Parse a number, returning fallback if not finite. */
function numOr_(v, fallback) {
  const n = parseNumber_(v);
  return Number.isFinite(n) ? n : fallback;
}

/**
 * Format expiration for chart label: "Dec 28" from a Date or string.
 */
function formatExpirationLabel_(exp) {
  if (!exp) return null;

  let d = exp;
  if (!(d instanceof Date)) {
    d = new Date(exp);
  }
  if (isNaN(d.getTime())) return null;

  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const mon = months[d.getMonth()];
  const yr = String(d.getFullYear()).slice(-2);
  return `${mon} ${yr}`;
}

// ---- Date Parsing ----

/**
 * Parses various date formats into a Date at midnight local time.
 * Handles: Date objects, YYYY-MM-DD, M/D/YY, M/D/YYYY, and native Date parsing.
 * Always returns a midnight-normalized date or null on failure.
 *
 * @param {Date|string|number} dateVal - Date value to parse
 * @returns {Date|null} Date at midnight local time, or null if invalid
 */
function parseDateAtMidnight_(dateVal) {
  if (!dateVal) return null;

  // Already a Date object - normalize to midnight
  if (dateVal instanceof Date) {
    if (isNaN(dateVal.getTime())) return null;
    return new Date(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate());
  }

  const s = String(dateVal).trim();
  if (!s) return null;

  // Try YYYY-MM-DD (ISO format)
  const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    const dt = new Date(
      parseInt(isoMatch[1], 10),
      parseInt(isoMatch[2], 10) - 1,
      parseInt(isoMatch[3], 10)
    );
    return isNaN(dt.getTime()) ? null : dt;
  }

  // Try M/D/YY or M/D/YYYY
  const mdyMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (mdyMatch) {
    let year = parseInt(mdyMatch[3], 10);
    // Handle 2-digit year: assume 20xx for years 00-49, 19xx for 50-99
    if (year < 100) {
      year = year < 50 ? 2000 + year : 1900 + year;
    }
    const dt = new Date(year, parseInt(mdyMatch[1], 10) - 1, parseInt(mdyMatch[2], 10));
    return isNaN(dt.getTime()) ? null : dt;
  }

  // Fallback to native parsing, then normalize to midnight
  const parsed = new Date(s);
  if (isNaN(parsed.getTime())) return null;
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}


// ---- Column Index Helpers ----

/**
 * Normalizes a header string: lowercase, strip all non-alphanumeric chars.
 * "Open Int" → "openint", "Call/Put" → "callput", "Implied Volatility" → "impliedvolatility"
 */
function normKey_(v) {
  return String(v ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * Finds the index of the first matching alias in headers. Returns -1 if none found.
 * Both headers and aliases are normalized via normKey_ before comparison.
 */
function findColumn_(headers, aliases) {
  const normHeaders = headers.map(normKey_);
  const normAliases = aliases.map(normKey_);
  for (let i = 0; i < normHeaders.length; i++) {
    if (normAliases.includes(normHeaders[i])) return i;
  }
  return -1;
}

/**
 * Finds indexes of option price columns in headers using alias lists.
 * Headers are normalized via normKey_, so whitespace/punctuation/case don't matter.
 * @param {Array<string>} headers - Raw headers from the sheet.
 */
function findColumnIndexes_(headers) {
  const norm = headers.map(normKey_);
  const find = (aliases) => {
    const na = aliases.map(normKey_);
    for (let i = 0; i < norm.length; i++) {
      if (na.includes(norm[i])) return i;
    }
    return -1;
  };
  return {
    strikeIdx:    find(["strike"]),
    bidIdx:       find(["bid"]),
    midIdx:       find(["mid"]),
    askIdx:       find(["ask"]),
    typeIdx:      find(["type", "option type", "call/put", "cp", "put/call"]),
    ivIdx:        find(["iv", "implied volatility"]),
    deltaIdx:     find(["delta"]),
    volumeIdx:    find(["volume", "vol"]),
    openIntIdx:   find(["open int", "open interest", "oi"]),
    moneynessIdx: find(["moneyness", "money", "itm/otm"]),
  };
}

// ---- Strike Parsing ----

/**
 * Parses a strike pair string like "450/460" into [lower, upper] strings.
 * Extracts the first "number/number" pattern found anywhere in the input.
 * Throws if not found, non-numeric, or lower >= upper.
 */
function parseStrikePairStrict_(strikePair) {
  const text = String(strikePair);
  const match = text.match(/(\d+(?:\.\d+)?)\s*\/\s*(\d+(?:\.\d+)?)/);
  if (!match) {
    throw new Error(`Strike must contain a pair like "450/460". Got: "${strikePair}"`);
  }

  const lower = Number(match[1]);
  const upper = Number(match[2]);

  if (!Number.isFinite(lower) || !Number.isFinite(upper)) {
    throw new Error(`Strikes must be numeric: "${strikePair}"`);
  }

  if (lower >= upper) {
    throw new Error(`Invalid strike order: ${lower} must be < ${upper}`);
  }
  return [match[1], match[2]];
}

// ---- Unified Multi-Leg Position Parsing ----

/**
 * Detects strategy type from leg data. For use in Legs table Strategy column.
 *
 * @param {Range} strikeRange - Strike column for the group (e.g., C3:C4)
 * @param {Range} typeRange - Type column for the group (e.g., D3:D4)
 * @param {Range} qtyRange - Qty column for the group (e.g., F3:F4)
 * @param {Array} [_labels] - Optional; ignored, for spreadsheet readability
 * @return {string} Strategy: "bull-call-spread", "bull-put-spread", "iron-condor", "stock", or "?"
 * @customfunction
 */
function detectStrategy(strikeRange, typeRange, qtyRange, _labels) {
  // Flatten inputs (ranges come as 2D arrays)
  const strikes = Array.isArray(strikeRange) ? strikeRange.flat() : [strikeRange];
  const types = Array.isArray(typeRange) ? typeRange.flat() : [typeRange];
  const qtys = Array.isArray(qtyRange) ? qtyRange.flat() : [qtyRange];

  // Build legs array
  const legs = [];
  const n = Math.max(strikes.length, types.length, qtys.length);
  for (let i = 0; i < n; i++) {
    const strike = parseNumber_(strikes[i] ?? "");
    const type = parseOptionType_(types[i] ?? "");
    const qty = parseNumber_(qtys[i] ?? "");
    if (!Number.isFinite(qty) || qty === 0) continue;
    legs.push({ strike, type: type || (Number.isFinite(strike) ? null : "Stock"), qty });
  }

  const posType = detectPositionType_(legs);
  return posType || "?";
}

/**
 * Detects position type from an array of leg objects.
 * Each leg: { strike, type, qty }
 * Returns: "stock", "bull-call-spread", "bull-put-spread", "iron-condor", "custom", or null.
 * Multi-leg positions with imbalanced quantities return "custom".
 */
function detectPositionType_(legs) {
  if (!legs || legs.length === 0) return null;

  if (legs.length === 1) {
    const leg = legs[0];
    if (leg.type === "Cash") return "cash";
    if (!leg.type || leg.type === "Stock" || !Number.isFinite(leg.strike)) return "stock";
    const direction = leg.qty >= 0 ? "long" : "short";
    if (leg.type === "Call") return direction + "-call";
    if (leg.type === "Put") return direction + "-put";
    return null;
  }

  // Multi-leg positions: check for known patterns with balanced quantities
  if (legs.length === 2) {
    const [a, b] = legs;
    // Both must have option types
    if (!a.type || !b.type) return "custom";
    if (a.type === "Stock" || b.type === "Stock") return "custom";

    const sameType = a.type === b.type;
    const bothLong = a.qty > 0 && b.qty > 0;
    const bothShort = a.qty < 0 && b.qty < 0;
    const oppositeSigns = (a.qty > 0 && b.qty < 0) || (a.qty < 0 && b.qty > 0);
    const equalQty = Math.abs(a.qty) === Math.abs(b.qty);

    // Straddle/strangle: call + put, same direction, equal qty
    if (!sameType && (bothLong || bothShort) && equalQty) {
      const sameStrike = a.strike === b.strike;
      if (bothLong) {
        return sameStrike ? "long-straddle" : "long-strangle";
      } else {
        return sameStrike ? "short-straddle" : "short-strangle";
      }
    }

    // Vertical spreads: same type, opposite signs, AND equal absolute quantities
    if (sameType && oppositeSigns && equalQty) {
      if (a.type === "Call") return "bull-call-spread";
      if (a.type === "Put") return "bull-put-spread";
    }

    // 2 option legs that don't match known patterns = custom
    return "custom";
  }

  if (legs.length === 4) {
    const calls = legs.filter(l => l.type === "Call");
    const puts = legs.filter(l => l.type === "Put");
    if (calls.length === 2 && puts.length === 2) {
      // All legs must have equal absolute quantities for balanced iron condor/butterfly
      const absQtys = legs.map(l => Math.abs(l.qty));
      const allEqualQty = absQtys.every(q => q === absQtys[0]);
      if (!allEqualQty) return "custom"; // Unequal qty = custom

      // Distinguish iron-butterfly vs iron-condor
      // Iron butterfly: short put and short call have the same strike
      const shortCall = calls.find(l => l.qty < 0);
      const shortPut = puts.find(l => l.qty < 0);
      if (shortCall && shortPut && shortCall.strike === shortPut.strike) {
        return "iron-butterfly";
      }
      return "iron-condor";
    }
    // 4 legs that don't match known patterns = custom
    return "custom";
  }

  // 3+ legs that don't match known patterns = custom
  if (legs.length >= 2) return "custom";

  return null;
}

/**
 * Parses positions from a unified Legs table for a specific symbol.
 * Handles merged-cell carry-forward for Symbol, Group, Strategy columns.
 *
 * @param {Array[]} rows - 2D array with header row first
 * @param {string} symbol - Uppercase ticker to filter by
 * @returns {{ shares: Array<{qty, basis}>, bullCallSpreads: Array, bullPutSpreads: Array, bearCallSpreads: Array }}
 */
function parsePositionsForSymbol_(rows, symbol) {
  const result = {
    shares: [],
    bullCallSpreads: [],
    bullPutSpreads: [],
    bearCallSpreads: [],
    longCalls: [],
    shortCalls: [],
    longPuts: [],
    shortPuts: [],
    customPositions: [],
    cash: 0
  };

  if (!rows || rows.length < 2) return result;

  const headers = rows[0];
  const idxSym = findColumn_(headers, ["symbol", "ticker"]);
  const idxGroup = findColumn_(headers, ["group", "grp"]);
  const idxStrategy = findColumn_(headers, ["strategy", "strat", "type"]);
  const idxStrike = findColumn_(headers, ["strike", "strikeprice"]);
  const idxType = findColumn_(headers, ["type", "optiontype", "callput", "cp", "putcall", "legtype"]);
  const idxExp = findColumn_(headers, ["expiration", "exp", "expiry", "expirationdate", "expdate"]);
  const idxQty = findColumn_(headers, ["qty", "quantity", "contracts", "contract", "count", "shares"]);
  const idxPrice = findColumn_(headers, ["price", "cost", "entry", "premium", "basis", "costbasis", "avgprice", "pricepaid"]);
  const idxClosed = findColumn_(headers, ["closed", "actualclose", "closedat"]);
  const idxDesc = findColumn_(headers, ["description", "desc", "label"]);

  // If Strategy and Type resolve to same column, disambiguate: Strategy needs its own column
  // Strategy aliases shouldn't include bare "type" if Type column also uses "type"
  // Re-find Strategy without "type" alias if they collided
  let idxStrat = idxStrategy;
  if (idxStrat >= 0 && idxStrat === idxType) {
    idxStrat = findColumn_(headers, ["strategy", "strat"]);
  }

  if (idxQty < 0 || idxPrice < 0) return result;

  // Group rows by (symbol, group) with carry-forward
  let lastSym = "";
  let lastGroup = "";

  // Collect groups: Map<groupKey, { legs: [...], firstRow: number, closed: boolean, closedCount: number, legCount: number }>
  const groups = new Map();

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];

    // Carry-forward symbol (skip summary rows)
    const rawSym = String(row[idxSym] ?? "").trim().toUpperCase();
    const skipSymbols = new Set(["TOTAL", "REALIZED", "UNREALIZED", "CASH", "SUMMARY"]);
    if (rawSym && !skipSymbols.has(rawSym)) lastSym = rawSym;
    if (lastSym !== symbol) continue;

    // Carry-forward group
    const rawGroup = String(row[idxGroup >= 0 ? idxGroup : -1] ?? "").trim();
    if (rawGroup) lastGroup = rawGroup;

    // Track closed status per group
    const groupKey = `${lastSym}|${lastGroup || r}`;
    if (!groups.has(groupKey)) {
      const desc = idxDesc >= 0 ? String(row[idxDesc] ?? "").trim() : "";
      groups.set(groupKey, { legs: [], firstRow: r, closedCount: 0, legCount: 0, groupName: lastGroup, description: desc });
    }
    const g = groups.get(groupKey);
    g.legCount++;

    // Closed column (closing price): non-empty means this leg is closed
    if (idxClosed >= 0) {
      const closedVal = String(row[idxClosed] ?? "").trim();
      if (closedVal !== "") g.closedCount++;
    }

    const qty = parseNumber_(row[idxQty]);
    const price = parseNumber_(row[idxPrice]);
    if (!Number.isFinite(qty) || qty === 0) continue;
    if (!Number.isFinite(price)) continue;

    const strike = idxStrike >= 0 ? parseNumber_(row[idxStrike]) : NaN;
    const optType = idxType >= 0 ? parseOptionType_(row[idxType]) : null;

    // For stock legs, parseOptionType_ may return null but the Strategy column or
    // lack of strike indicates stock
    let legType = optType;
    if (!legType && !Number.isFinite(strike)) {
      legType = "Stock";
    }
    // Check if strategy column says Stock or Cash
    if (!legType && idxStrat >= 0) {
      const strat = parseSpreadStrategy_(row[idxStrat]);
      if (strat === "stock") legType = "Stock";
      if (strat === "cash") legType = "Cash";
    }
    // Also check type column for Cash
    if (!legType && idxType >= 0) {
      const typeStr = String(row[idxType] ?? "").trim().toLowerCase();
      if (typeStr === "cash" || typeStr === "$" || typeStr === "usd") legType = "Cash";
    }

    const leg = { qty, price, strike, type: legType };
    if (idxExp >= 0 && row[idxExp]) leg.expiration = row[idxExp];
    groups.get(groupKey).legs.push(leg);
  }

  // Process each group (skip closed)
  for (const [key, group] of groups) {
    const legs = group.legs;
    const posType = detectPositionType_(legs);
    // Skip closed groups: all legs have a value in the Closed column
    const allLegsClosed = idxClosed >= 0 && group.closedCount >= group.legCount && group.legCount > 0;
    if (allLegsClosed) continue;

    if (posType === "stock") {
      const leg = legs[0];
      result.shares.push({ qty: leg.qty, basis: leg.price });
    } else if (posType === "cash") {
      // Cash: just add the amount (qty * price or just price if qty=1)
      const leg = legs[0];
      const amount = leg.qty * leg.price;
      result.cash += amount;
    } else if (posType === "bull-call-spread") {
      const longLeg = legs.find(l => l.qty > 0);
      const shortLeg = legs.find(l => l.qty < 0);
      if (!longLeg || !shortLeg) continue;

      const debit = longLeg.price - shortLeg.price;
      let label = `${longLeg.strike}/-${shortLeg.strike} BCS`;
      if (longLeg.expiration) {
        const expLabel = formatExpirationLabel_(longLeg.expiration);
        if (expLabel) label = `${expLabel} ${label}`;
      }

      result.bullCallSpreads.push({
        qty: Math.abs(longLeg.qty),
        kLong: longLeg.strike,
        kShort: shortLeg.strike,
        priceLong: longLeg.price,
        priceShort: shortLeg.price,
        debit,
        flavor: "CALL",
        label,
        expiration: longLeg.expiration,
        symbol: lastSym,
      });
    } else if (posType === "bull-put-spread") {
      const longLeg = legs.find(l => l.qty > 0);
      const shortLeg = legs.find(l => l.qty < 0);
      if (!longLeg || !shortLeg) continue;

      const debit = longLeg.price - shortLeg.price;
      let label = `${longLeg.strike}/-${shortLeg.strike} BPS`;
      if (longLeg.expiration) {
        const expLabel = formatExpirationLabel_(longLeg.expiration);
        if (expLabel) label = `${expLabel} ${label}`;
      }

      result.bullPutSpreads.push({
        qty: Math.abs(longLeg.qty),
        kLong: longLeg.strike,
        kShort: shortLeg.strike,
        priceLong: longLeg.price,
        priceShort: shortLeg.price,
        debit,
        flavor: "PUT",
        label,
        expiration: longLeg.expiration,
        symbol: lastSym,
      });
    } else if (posType === "iron-condor") {
      const puts = legs.filter(l => l.type === "Put").sort((a, b) => a.strike - b.strike);
      const calls = legs.filter(l => l.type === "Call").sort((a, b) => a.strike - b.strike);

      // Bull put spread: long lower put, short higher put
      const longPut = puts.find(l => l.qty > 0) || puts[0];
      const shortPut = puts.find(l => l.qty < 0) || puts[1];
      // Bear call spread: short lower call, long higher call
      const shortCall = calls.find(l => l.qty < 0) || calls[0];
      const longCall = calls.find(l => l.qty > 0) || calls[1];

      if (longPut && shortPut) {
        const putDebit = longPut.price - shortPut.price;
        let label = `IC ${longPut.strike}/-${shortPut.strike}/-${shortCall.strike}/${longCall.strike}`;
        if (longPut.expiration) {
          const expLabel = formatExpirationLabel_(longPut.expiration);
          if (expLabel) label = `${expLabel} ${label}`;
        }
        result.bullPutSpreads.push({
          qty: Math.abs(longPut.qty),
          kLong: longPut.strike,
          kShort: shortPut.strike,
          priceLong: longPut.price,
          priceShort: shortPut.price,
          debit: putDebit,
          flavor: "PUT",
          label: label + " (put)",
          expiration: longPut.expiration,
          symbol: lastSym,
        });
      }

      if (shortCall && longCall) {
        const callDebit = longCall.price - shortCall.price;
        let label = `IC ${longPut.strike}/-${shortPut.strike}/-${shortCall.strike}/${longCall.strike}`;
        if (shortCall.expiration) {
          const expLabel = formatExpirationLabel_(shortCall.expiration);
          if (expLabel) label = `${expLabel} ${label}`;
        }
        result.bearCallSpreads.push({
          qty: Math.abs(shortCall.qty),
          kLong: shortCall.strike,
          kShort: longCall.strike,
          priceLong: shortCall.price,
          priceShort: longCall.price,
          debit: callDebit,
          flavor: "BEAR_CALL",
          label: label + " (call)",
          expiration: shortCall.expiration,
          symbol: lastSym,
        });
      }
    } else if (posType === "long-call" || posType === "short-call" ||
               posType === "long-put" || posType === "short-put") {
      // Single-leg option position
      const leg = legs[0];
      if (!leg || !Number.isFinite(leg.strike)) continue;

      // Short positions show negative strike prefix
      const isShort = posType === "short-call" || posType === "short-put";
      let label = isShort ? `-${leg.strike}` : `${leg.strike}`;
      if (leg.expiration) {
        const expLabel = formatExpirationLabel_(leg.expiration);
        if (expLabel) label = `${expLabel} ${label}`;
      }

      const optionPos = {
        qty: Math.abs(leg.qty),
        strike: leg.strike,
        price: leg.price,
        type: leg.type,
        label,
        expiration: leg.expiration,
        symbol: lastSym,
        isLong: leg.qty > 0,
      };

      if (posType === "long-call") {
        result.longCalls.push(optionPos);
      } else if (posType === "short-call") {
        result.shortCalls.push(optionPos);
      } else if (posType === "long-put") {
        result.longPuts.push(optionPos);
      } else if (posType === "short-put") {
        result.shortPuts.push(optionPos);
      }
    } else if (posType === "custom") {
      // Custom multi-leg position (imbalanced qty or unrecognized pattern)
      // Build label: user description, or "max(qty) - expiration strike1/strike2/..."
      const maxQty = Math.max(...legs.map(l => Math.abs(l.qty)));
      const expLeg = legs.find(l => l.expiration);
      const expLabel = expLeg ? formatExpirationLabel_(expLeg.expiration) : null;
      // Show strikes with - prefix for short positions: 500/-600/740/-900
      const strikes = legs
        .filter(l => Number.isFinite(l.strike))
        .sort((a, b) => a.strike - b.strike)
        .map(l => l.qty < 0 ? `-${l.strike}` : `${l.strike}`)
        .join('/');

      // Label: use Description from sheet, or auto-generate "Jun 27 500/-600/740/-900 custom"
      // User can set Description with: =formatLegsDescription(strikeRange, qtyRange, "custom")
      let label = group.description;
      if (!label) {
        label = `${expLabel || '?'} ${strikes || '?'} custom`;
      }

      // Build custom position with all legs as individual options
      const customLegs = legs.filter(l => l.type === "Call" || l.type === "Put").map(l => ({
        qty: Math.abs(l.qty),
        strike: l.strike,
        price: l.price,
        type: l.type,
        expiration: l.expiration,
        symbol: lastSym,
        isLong: l.qty > 0,
      }));

      if (customLegs.length > 0) {
        result.customPositions.push({
          legs: customLegs,
          label,
          qty: maxQty,
          symbol: lastSym,
          expiration: expLeg?.expiration,
          groupName: group.groupName,
        });
      }
    }
  }

  return result;
}

/**
 * Returns unique symbols from a Legs table, with carry-forward for merged cells.
 * @param {Array[]} rows - 2D array with header row first
 * @returns {string[]} Sorted unique symbols
 */
function getSymbolsFromLegsTable_(rows) {
  if (!rows || rows.length < 2) return [];

  const headers = rows[0];
  const idxSym = findColumn_(headers, ["symbol", "ticker"]);
  if (idxSym < 0) return [];

  const symbols = new Set();
  let lastSym = "";

  for (let r = 1; r < rows.length; r++) {
    const rawSym = String(rows[r][idxSym] ?? "").trim().toUpperCase();
    if (rawSym) lastSym = rawSym;
    if (lastSym) symbols.add(lastSym);
  }

  const skip = new Set(["REALIZED", "UNREALIZED", "TOTAL", "CASH"]);
  return Array.from(symbols).filter(s => !skip.has(s)).sort();
}

// ---- Tests ----

function test_parseSpreadStrategy() {
  // Null/Empty
  assertEqual(parseSpreadStrategy_(null), null, "null input");
  assertEqual(parseSpreadStrategy_(undefined), null, "undefined input");
  assertEqual(parseSpreadStrategy_(""), null, "empty string");
  assertEqual(parseSpreadStrategy_("   "), null, "whitespace only");

  // Stock
  assertEqual(parseSpreadStrategy_("stock"), "stock", "stock basic");
  assertEqual(parseSpreadStrategy_("STOCK"), "stock", "stock uppercase");
  assertEqual(parseSpreadStrategy_("Stock"), "stock", "stock mixed case");
  assertEqual(parseSpreadStrategy_("stocks"), "stock", "stocks plural");
  assertEqual(parseSpreadStrategy_("SHARES"), "stock", "shares uppercase");
  assertEqual(parseSpreadStrategy_("share"), "stock", "share singular");
  assertEqual(parseSpreadStrategy_(" stock "), "stock", "stock with spaces");
  assertEqual(parseSpreadStrategy_("st ock"), null, "invalid with space inside");
  assertEqual(parseSpreadStrategy_("stock\u00a0"), "stock", "stock with nbsp");

  // Call/Put should NOT match spread strategy
  assertEqual(parseSpreadStrategy_("call"), null, "call is not a strategy");
  assertEqual(parseSpreadStrategy_("put"), null, "put is not a strategy");
  assertEqual(parseSpreadStrategy_("c"), null, "c is not a strategy");
  assertEqual(parseSpreadStrategy_("p"), null, "p is not a strategy");

  // Bull Call Spread
  assertEqual(parseSpreadStrategy_("bcs"), "bull-call-spread", "bcs abbr");
  assertEqual(parseSpreadStrategy_("BCS"), "bull-call-spread", "BCS upper");
  assertEqual(parseSpreadStrategy_("bull call spread"), "bull-call-spread", "bull call spread spaces");
  assertEqual(parseSpreadStrategy_("Bull-Call-Spread"), "bull-call-spread", "bull-call-spread dashes");
  assertEqual(parseSpreadStrategy_("bull.call.spread"), "bull-call-spread", "bull.call.spread dots");
  assertEqual(parseSpreadStrategy_("bull call spreads"), "bull-call-spread", "bull call spreads plural");
  assertEqual(parseSpreadStrategy_("bull\u2013call\u2014spread"), "bull-call-spread", "bull en/em dash spread");
  assertEqual(parseSpreadStrategy_(" bull  call   spread "), "bull-call-spread", "extra spaces");
  assertEqual(parseSpreadStrategy_("BuLl CaLl SpReAd"), "bull-call-spread", "mixed case");
  assertEqual(parseSpreadStrategy_("bull calls spread"), null, "invalid plural mismatch");
  assertEqual(parseSpreadStrategy_("bcs extra"), null, "bcs with extra");

  // Bull Put Spread
  assertEqual(parseSpreadStrategy_("bps"), "bull-put-spread", "bps abbr");
  assertEqual(parseSpreadStrategy_("BPS"), "bull-put-spread", "BPS upper");
  assertEqual(parseSpreadStrategy_("bull put spread"), "bull-put-spread", "bull put spread spaces");
  assertEqual(parseSpreadStrategy_("Bull-Put-Spread"), "bull-put-spread", "bull-put-spread dashes");
  assertEqual(parseSpreadStrategy_("bull.put.spread"), "bull-put-spread", "bull.put.spread dots");
  assertEqual(parseSpreadStrategy_("bull put spreads"), "bull-put-spread", "bull put spreads plural");
  assertEqual(parseSpreadStrategy_("bull\u2013put\u2014spread"), "bull-put-spread", "bull en/em dash spread");
  assertEqual(parseSpreadStrategy_(" bull  put   spread "), "bull-put-spread", "extra spaces");
  assertEqual(parseSpreadStrategy_("BuLl PuT SpReAd"), "bull-put-spread", "mixed case");
  assertEqual(parseSpreadStrategy_("bull puts spread"), null, "invalid plural mismatch");
  assertEqual(parseSpreadStrategy_("bps extra"), null, "bps with extra");

  // Invalid
  assertEqual(parseSpreadStrategy_("bull spread"), null, "missing type");
  assertEqual(parseSpreadStrategy_("bear call spread"), null, "wrong direction");
  assertEqual(parseSpreadStrategy_("123"), null, "numbers");
  assertEqual(parseSpreadStrategy_("stock spread"), null, "mixed invalid");
  assertEqual(parseSpreadStrategy_("sto\u0301ck"), null, "accented stock");

  Logger.log("All parseSpreadStrategy tests passed");
}

function test_parseOptionType() {
  assertEqual(parseOptionType_("call"), "Call", "call");
  assertEqual(parseOptionType_("Call"), "Call", "Call mixed");
  assertEqual(parseOptionType_("CALL"), "Call", "CALL upper");
  assertEqual(parseOptionType_("calls"), "Call", "calls plural");
  assertEqual(parseOptionType_("c"), "Call", "c shorthand");
  assertEqual(parseOptionType_("C"), "Call", "C shorthand upper");
  assertEqual(parseOptionType_("put"), "Put", "put");
  assertEqual(parseOptionType_("Put"), "Put", "Put mixed");
  assertEqual(parseOptionType_("PUT"), "Put", "PUT upper");
  assertEqual(parseOptionType_("puts"), "Put", "puts plural");
  assertEqual(parseOptionType_("p"), "Put", "p shorthand");
  assertEqual(parseOptionType_("P"), "Put", "P shorthand upper");
  assertEqual(parseOptionType_(null), null, "null");
  assertEqual(parseOptionType_(""), null, "empty");
  assertEqual(parseOptionType_("stock"), null, "stock is not an option type");
  assertEqual(parseOptionType_("bcs"), null, "bcs is not an option type");

  Logger.log("All parseOptionType tests passed");
}

function test_loadCsvData_columnOrders() {
  const symbol = "TSLA";
  const expDate = new Date(2028, 5, 16);

  // loadCsvData_ returns 12 columns: [symbol, expDate, strike, type, bid, mid, ask, iv, delta, volume, openInt, moneyness]
  // When optional columns are missing, they are null

  // Test 1: Standard order (no iv, delta, volume, openInt, moneyness columns)
  const csv1 = [
    ["Strike", "Bid", "Mid", "Ask", "Type"],
    ["450", "203.15", "206.00", "208.85", "Call"],
    ["350", "250.00", "255.00", "260.00", "Put"]
  ];
  const rows1 = loadCsvData_(csv1, symbol, expDate);
  assertArrayDeepEqual(rows1, [
    [symbol, expDate, 450, "Call", 203.15, 206.00, 208.85, null, null, null, null, null],
    [symbol, expDate, 350, "Put", 250.00, 255.00, 260.00, null, null, null, null, null]
  ], "Test 1: Standard order");

  // Test 2: Reversed order
  const csv2 = [
    ["Ask", "Bid", "Mid", "Type", "Strike"],
    ["208.85", "203.15", "206.00", "Call", "450"],
    ["260.00", "250.00", "255.00", "Put", "350"]
  ];
  const rows2 = loadCsvData_(csv2, symbol, expDate);
  assertArrayDeepEqual(rows2, [
    [symbol, expDate, 450, "Call", 203.15, 206.00, 208.85, null, null, null, null, null],
    [symbol, expDate, 350, "Put", 250.00, 255.00, 260.00, null, null, null, null, null]
  ], "Test 2: Reversed order");

  // Test 3: Mixed case headers, no Mid, alt Type name
  const csv3 = [
    ["sTriKe", "BID", "ASK", "Option Type"],
    ["450", "203.15", "208.85", "Call"],
    ["350", "250.00", "260.00", "Put"]
  ];
  const rows3 = loadCsvData_(csv3, symbol, expDate);
  assertArrayDeepEqual(rows3, [
    [symbol, expDate, 450, "Call", 203.15, null, 208.85, null, null, null, null, null],
    [symbol, expDate, 350, "Put", 250.00, null, 260.00, null, null, null, null, null]
  ], "Test 3: Mixed case, no Mid, alt Type");

  // Test 4: Missing required column -> empty
  const csv4 = [
    ["Strike", "Mid", "Ask", "Type"],
    ["450", "206.00", "208.85", "Call"]
  ];
  const rows4 = loadCsvData_(csv4, symbol, expDate);
  assertArrayDeepEqual(rows4, [], "Test 4: Missing Bid -> empty");

  Logger.log("All parseCsvData column order tests passed");
}

function test_detectPositionType_stock() {
  assertEqual(detectPositionType_([{ strike: NaN, type: "Stock", qty: 100 }]), "stock", "single stock leg");
  assertEqual(detectPositionType_([{ strike: NaN, type: null, qty: 100 }]), "stock", "no type no strike = stock");
  assertEqual(detectPositionType_([]), null, "empty legs");
  assertEqual(detectPositionType_(null), null, "null");
  Logger.log("All detectPositionType stock tests passed");
}

function test_detectPositionType_spreads() {
  assertEqual(
    detectPositionType_([
      { strike: 300, type: "Call", qty: 7 },
      { strike: 440, type: "Call", qty: -7 },
    ]),
    "bull-call-spread",
    "BCS: 2 calls opposite signs"
  );
  assertEqual(
    detectPositionType_([
      { strike: 200, type: "Put", qty: 5 },
      { strike: 250, type: "Put", qty: -5 },
    ]),
    "bull-put-spread",
    "BPS: 2 puts opposite signs"
  );
  assertEqual(
    detectPositionType_([
      { strike: 200, type: "Call", qty: 5 },
      { strike: 250, type: "Call", qty: 5 },
    ]),
    "custom",
    "same sign = custom"
  );
  // Ratio spreads (unequal quantities) should be detected as custom
  assertEqual(
    detectPositionType_([
      { strike: 740, type: "Call", qty: 2 },
      { strike: 900, type: "Call", qty: -1 },
    ]),
    "custom",
    "unequal qty = custom (ratio spread)"
  );
  assertEqual(
    detectPositionType_([
      { strike: 200, type: "Put", qty: 3 },
      { strike: 250, type: "Put", qty: -1 },
    ]),
    "custom",
    "unequal qty puts = custom (ratio spread)"
  );
  Logger.log("All detectPositionType spread tests passed");
}

function test_detectPositionType_ironCondor() {
  assertEqual(
    detectPositionType_([
      { strike: 200, type: "Put", qty: 3 },
      { strike: 250, type: "Put", qty: -3 },
      { strike: 400, type: "Call", qty: -3 },
      { strike: 450, type: "Call", qty: 3 },
    ]),
    "iron-condor",
    "IC: 2 puts + 2 calls balanced"
  );
  // Imbalanced iron condor = custom
  assertEqual(
    detectPositionType_([
      { strike: 200, type: "Put", qty: 2 },
      { strike: 250, type: "Put", qty: -3 },
      { strike: 400, type: "Call", qty: -3 },
      { strike: 450, type: "Call", qty: 3 },
    ]),
    "custom",
    "IC with imbalanced qty = custom"
  );
  Logger.log("All detectPositionType iron condor tests passed");
}

function test_detectPositionType_ironButterfly() {
  assertEqual(
    detectPositionType_([
      { strike: 350, type: "Put", qty: 5 },
      { strike: 400, type: "Put", qty: -5 },
      { strike: 400, type: "Call", qty: -5 },
      { strike: 450, type: "Call", qty: 5 },
    ]),
    "iron-butterfly",
    "IB: short put and short call at same strike"
  );
  Logger.log("All detectPositionType iron butterfly tests passed");
}

function test_parsePositionsForSymbol_stock() {
  const rows = [
    ["Symbol", "Group", "Strategy", "Strike", "Type", "Expiration", "Qty", "Price"],
    ["TSLA",   "4",     "",         "",       "Stock", "",          "600", "333"],
  ];
  const result = parsePositionsForSymbol_(rows, "TSLA");
  assertEqual(result.shares.length, 1, "one stock position");
  assertEqual(result.shares[0].qty, 600, "stock qty");
  assertEqual(result.shares[0].basis, 333, "stock basis");
  Logger.log("All parsePositionsForSymbol stock tests passed");
}

function test_parsePositionsForSymbol_bullCallSpread() {
  const rows = [
    ["Symbol", "Group", "Strategy", "Strike", "Type", "Expiration",  "Qty", "Price"],
    ["TSLA",   "1",     "",         "300",    "Call", "12/15/2028",  "7",   "223.50"],
    ["",       "",      "",         "440",    "Call", "12/15/2028",  "-7",  "165.50"],
  ];
  const result = parsePositionsForSymbol_(rows, "TSLA");
  assertEqual(result.bullCallSpreads.length, 1, "one BCS");
  const bcs = result.bullCallSpreads[0];
  assertEqual(bcs.qty, 7, "BCS qty");
  assertEqual(bcs.kLong, 300, "BCS long strike");
  assertEqual(bcs.kShort, 440, "BCS short strike");
  assertEqual(bcs.debit, 58, "BCS debit = 223.50 - 165.50");
  assertEqual(bcs.flavor, "CALL", "BCS flavor");
  Logger.log("All parsePositionsForSymbol BCS tests passed");
}

function test_parsePositionsForSymbol_ironCondor() {
  const rows = [
    ["Symbol", "Group", "Strategy", "Strike", "Type", "Expiration",  "Qty", "Price"],
    ["TSLA",   "5",     "",         "200",    "Put",  "12/15/2028",  "3",   "10.00"],
    ["",       "",      "",         "250",    "Put",  "12/15/2028",  "-3",  "15.00"],
    ["",       "",      "",         "400",    "Call", "12/15/2028",  "-3",  "20.00"],
    ["",       "",      "",         "450",    "Call", "12/15/2028",  "3",   "12.00"],
  ];
  const result = parsePositionsForSymbol_(rows, "TSLA");
  assertEqual(result.bullPutSpreads.length, 1, "IC produces one BPS");
  assertEqual(result.bearCallSpreads.length, 1, "IC produces one bear call");
  assertEqual(result.bullPutSpreads[0].kLong, 200, "BPS long strike");
  assertEqual(result.bullPutSpreads[0].kShort, 250, "BPS short strike");
  assertEqual(result.bearCallSpreads[0].kLong, 400, "bear call long (short) strike");
  assertEqual(result.bearCallSpreads[0].kShort, 450, "bear call short (long) strike");
  Logger.log("All parsePositionsForSymbol iron condor tests passed");
}
