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
 * Parses "YYYY-MM-DD" into a Date at midnight local time.
 * Returns null on invalid input.
 */
function parseYyyyMmDdToDateAtMidnight_(s) {
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const y = Number(m[1]), mo = Number(m[2]) - 1, d = Number(m[3]);
  const dt = new Date(y, mo, d);
  if (isNaN(dt.getTime())) return null;
  dt.setHours(0, 0, 0, 0);
  return dt;
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

// ---- Position Table Parsing ----

/**
 * Parse iron condors from table for a specific symbol.
 * Iron condor = bull put spread + bear call spread
 *
 * Expected columns:
 *   Symbol, Expiration, Buy Put, Sell Put, Sell Call, Buy Call, Credit, Qty
 *
 * Returns { bullPutSpreads: [...], bearCallSpreads: [...] }
 */
function parseIronCondorsFromTableForSymbol_(rows, symbol) {
  const result = { bullPutSpreads: [], bearCallSpreads: [] };

  if (!rows || rows.length < 2) return result;

  const headerNorm = rows[0].map(normKey_);

  const idxSym = findColumn_(headerNorm, ["symbol", "ticker"]);
  const idxExp = findColumn_(headerNorm, ["expiration", "exp", "expiry", "expirationdate", "expdate"]);
  const idxStatus = findColumn_(headerNorm, ["status"]);
  const idxBuyPut = findColumn_(headerNorm, ["buyput", "longput", "putlong", "putbuy"]);
  const idxSellPut = findColumn_(headerNorm, ["sellput", "shortput", "putshort", "putsell"]);
  const idxSellCall = findColumn_(headerNorm, ["sellcall", "shortcall", "callshort", "callsell"]);
  const idxBuyCall = findColumn_(headerNorm, ["buycall", "longcall", "calllong", "callbuy"]);
  const idxCredit = findColumn_(headerNorm, ["credit", "netcredit", "premium"]);
  const idxQty = findColumn_(headerNorm, ["qty", "quantity", "contracts", "contract", "count"]);

  // Must have all four strikes
  if (idxBuyPut < 0 || idxSellPut < 0 || idxSellCall < 0 || idxBuyCall < 0) {
    return result;
  }

  for (let r = 1; r < rows.length; r++) {
    // Filter by symbol
    if (idxSym >= 0) {
      const rowSym = String(rows[r][idxSym] ?? "").trim().toUpperCase();
      if (rowSym && rowSym !== symbol) continue;
    }

    // Skip closed positions
    if (idxStatus >= 0) {
      const status = String(rows[r][idxStatus] ?? "").trim().toLowerCase();
      if (status === "closed") continue;
    }

    const buyPut = toNum_(rows[r][idxBuyPut]);
    const sellPut = toNum_(rows[r][idxSellPut]);
    const sellCall = toNum_(rows[r][idxSellCall]);
    const buyCall = toNum_(rows[r][idxBuyCall]);

    // All four strikes required
    if (!Number.isFinite(buyPut) || !Number.isFinite(sellPut) ||
        !Number.isFinite(sellCall) || !Number.isFinite(buyCall)) {
      continue;
    }

    const qty = idxQty >= 0 ? toNum_(rows[r][idxQty]) : 1;
    if (!Number.isFinite(qty) || qty === 0) continue;

    // Get credit (stored as negative debit for credit spreads)
    let credit = 0;
    if (idxCredit >= 0) {
      credit = toNum_(rows[r][idxCredit]);
      if (!Number.isFinite(credit)) credit = 0;
    }

    // Build label
    let label = `IC ${buyPut}/${sellPut}/${sellCall}/${buyCall}`;
    if (idxExp >= 0) {
      const expLabel = formatExpirationLabel_(rows[r][idxExp]);
      if (expLabel) label = `${expLabel} ${label}`;
    }

    // Split credit between the two spreads (approximation)
    const putWidth = sellPut - buyPut;
    const callWidth = buyCall - sellCall;
    const totalWidth = putWidth + callWidth;
    const putCredit = totalWidth > 0 ? credit * (putWidth / totalWidth) : credit / 2;
    const callCredit = totalWidth > 0 ? credit * (callWidth / totalWidth) : credit / 2;

    // Bull put spread: long lower put, short higher put
    result.bullPutSpreads.push({
      qty,
      kLong: buyPut,
      kShort: sellPut,
      debit: -putCredit,
      flavor: "PUT",
      label: label + " (put)",
    });

    // Bear call spread: short lower call, long higher call
    result.bearCallSpreads.push({
      qty,
      kLong: sellCall,
      kShort: buyCall,
      debit: -callCredit,
      flavor: "BEAR_CALL",
      label: label + " (call)",
    });
  }

  return result;
}

/**
 * Parse shares/stock positions from a table for a specific symbol.
 * Returns array of { qty, basis } objects.
 */
function parseSharesFromTableForSymbol_(rows, symbol, outMeta) {
  if (!rows || rows.length < 2) return [];

  const headerNorm = rows[0].map(normKey_);

  const idxSym = findColumn_(headerNorm, ["symbol", "ticker"]);
  const idxQty = findColumn_(headerNorm, ["shares", "share", "qty", "quantity", "units", "position"]);
  const idxBasis = findColumn_(headerNorm, [
    "costbasis", "basis", "avgcost", "averagecost", "avgprice", "averageprice",
    "aveprice", "avepricepaid", "pricepaid", "entry", "entryprice", "cost", "purchaseprice",
  ]);

  const out = [];

  for (let r = 1; r < rows.length; r++) {
    if (idxSym >= 0) {
      const rowSym = String(rows[r][idxSym] ?? "").trim().toUpperCase();
      if (rowSym && rowSym !== symbol) continue;
    }

    if (idxQty < 0 || idxBasis < 0) continue;

    const qty = parseNumber_(rows[r][idxQty]);
    const basis = parseNumber_(rows[r][idxBasis]);

    if (!Number.isFinite(qty) || qty === 0) continue;
    if (!Number.isFinite(basis)) continue;

    out.push({ qty, basis });
  }

  return out;
}

/**
 * Parse spreads from a table for a specific symbol.
 * - Counts ONLY "definition rows" that contain BOTH Long Strike and Short Strike
 * - Ignores fill-detail rows automatically
 * - Uses debit cost preference: Ave Debit > Rec Debit > Net Debit/Debit/Cost/Entry/Price
 * - If contracts column missing, defaults qty=1
 * - Captures expiration date to generate labels like "Dec 28 350/450"
 *
 * @param {Array[]} rows - 2D array with header row first
 * @param {string} symbol - Uppercase ticker to filter by
 * @param {string} flavor - "CALL" or "PUT"
 * @returns {Array<{qty, kLong, kShort, debit, flavor, label}>}
 */
function parseSpreadsFromTableForSymbol_(rows, symbol, flavor) {
  if (!rows || rows.length < 2) return [];

  const headerNorm = rows[0].map(normKey_);

  const idxSym = findColumn_(headerNorm, ["symbol", "ticker"]);
  const idxQty = findColumn_(headerNorm, [
    "contracts", "contract", "qty", "quantity", "count", "numcontracts", "spreads", "spreadqty",
  ]);
  const idxLong = findColumn_(headerNorm, [
    "lower", "lowerstrike", "long", "longstrike", "buystrike", "strikebuy", "strikelong",
  ]);
  const idxShort = findColumn_(headerNorm, [
    "upper", "upperstrike", "short", "shortstrike", "sellstrike", "strikesell", "strikeshort",
  ]);
  const idxExp = findColumn_(headerNorm, [
    "expiration", "exp", "expiry", "expirationdate", "expdate",
  ]);
  const idxAveDebit = findColumn_(headerNorm, ["avedebit", "avgdebit", "averagedebit"]);
  const idxRecDebit = findColumn_(headerNorm, ["recdebit", "recommendeddebit"]);
  const idxDebitFallback = findColumn_(headerNorm, ["netdebit", "debit", "cost", "price", "entry", "premium"]);

  if (idxLong < 0 || idxShort < 0) return [];

  const assumeQty = idxQty < 0;
  const out = [];

  for (let r = 1; r < rows.length; r++) {
    if (idxSym >= 0) {
      const rowSym = String(rows[r][idxSym] ?? "").trim().toUpperCase();
      if (rowSym && rowSym !== symbol) continue;
    }

    const kLong = parseNumber_(rows[r][idxLong]);
    const kShort = parseNumber_(rows[r][idxShort]);

    if (!Number.isFinite(kLong) || !Number.isFinite(kShort)) continue;

    const qty = assumeQty ? 1 : parseNumber_(rows[r][idxQty]);
    if (!Number.isFinite(qty) || qty === 0) continue;

    let debit = NaN;
    if (idxAveDebit >= 0) debit = parseNumber_(rows[r][idxAveDebit]);
    if (!Number.isFinite(debit) && idxRecDebit >= 0) debit = parseNumber_(rows[r][idxRecDebit]);
    if (!Number.isFinite(debit) && idxDebitFallback >= 0) debit = parseNumber_(rows[r][idxDebitFallback]);
    if (!Number.isFinite(debit)) continue;

    let label = `${kLong}/${kShort}`;
    if (idxExp >= 0) {
      const expLabel = formatExpirationLabel_(rows[r][idxExp]);
      if (expLabel) label = `${expLabel} ${label}`;
    }

    out.push({ qty, kLong, kShort, debit, flavor, label });
  }

  return out;
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
 * Returns: "stock", "bull-call-spread", "bull-put-spread", "iron-condor", or null.
 */
function detectPositionType_(legs) {
  if (!legs || legs.length === 0) return null;

  if (legs.length === 1) {
    const leg = legs[0];
    if (!leg.type || leg.type === "Stock" || !Number.isFinite(leg.strike)) return "stock";
    const direction = leg.qty >= 0 ? "long" : "short";
    if (leg.type === "Call") return direction + "-call";
    if (leg.type === "Put") return direction + "-put";
    return null;
  }

  if (legs.length === 2) {
    const [a, b] = legs;
    // Both must have types and opposite qty signs
    if (!a.type || !b.type) return null;
    if (a.type === "Stock" || b.type === "Stock") return null;
    const sameType = a.type === b.type;
    const oppositeSigns = (a.qty > 0 && b.qty < 0) || (a.qty < 0 && b.qty > 0);
    if (!sameType || !oppositeSigns) return null;

    if (a.type === "Call") return "bull-call-spread";
    if (a.type === "Put") return "bull-put-spread";
    return null;
  }

  if (legs.length === 4) {
    const calls = legs.filter(l => l.type === "Call");
    const puts = legs.filter(l => l.type === "Put");
    if (calls.length === 2 && puts.length === 2) {
      // Distinguish iron-butterfly vs iron-condor
      // Iron butterfly: short put and short call have the same strike
      const shortCall = calls.find(l => l.qty < 0);
      const shortPut = puts.find(l => l.qty < 0);
      if (shortCall && shortPut && shortCall.strike === shortPut.strike) {
        return "iron-butterfly";
      }
      return "iron-condor";
    }
    return null;
  }

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
  const result = { shares: [], bullCallSpreads: [], bullPutSpreads: [], bearCallSpreads: [] };

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

    // Carry-forward symbol
    const rawSym = String(row[idxSym] ?? "").trim().toUpperCase();
    if (rawSym) lastSym = rawSym;
    if (lastSym !== symbol) continue;

    // Carry-forward group
    const rawGroup = String(row[idxGroup >= 0 ? idxGroup : -1] ?? "").trim();
    if (rawGroup) lastGroup = rawGroup;

    // Track closed status per group
    const groupKey = `${lastSym}|${lastGroup || r}`;
    if (!groups.has(groupKey)) {
      groups.set(groupKey, { legs: [], firstRow: r, closedCount: 0, legCount: 0 });
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
    // Check if strategy column says Stock
    if (!legType && idxStrat >= 0) {
      const strat = parseSpreadStrategy_(row[idxStrat]);
      if (strat === "stock") legType = "Stock";
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
    } else if (posType === "bull-call-spread") {
      const longLeg = legs.find(l => l.qty > 0);
      const shortLeg = legs.find(l => l.qty < 0);
      if (!longLeg || !shortLeg) continue;

      const debit = longLeg.price - shortLeg.price;
      let label = `${longLeg.strike}/${shortLeg.strike}`;
      if (longLeg.expiration) {
        const expLabel = formatExpirationLabel_(longLeg.expiration);
        if (expLabel) label = `${expLabel} ${label}`;
      }

      result.bullCallSpreads.push({
        qty: Math.abs(longLeg.qty),
        kLong: longLeg.strike,
        kShort: shortLeg.strike,
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
      let label = `${longLeg.strike}/${shortLeg.strike}`;
      if (longLeg.expiration) {
        const expLabel = formatExpirationLabel_(longLeg.expiration);
        if (expLabel) label = `${expLabel} ${label}`;
      }

      result.bullPutSpreads.push({
        qty: Math.abs(longLeg.qty),
        kLong: longLeg.strike,
        kShort: shortLeg.strike,
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
        let label = `IC ${longPut.strike}/${shortPut.strike}/${shortCall.strike}/${longCall.strike}`;
        if (longPut.expiration) {
          const expLabel = formatExpirationLabel_(longPut.expiration);
          if (expLabel) label = `${expLabel} ${label}`;
        }
        result.bullPutSpreads.push({
          qty: Math.abs(longPut.qty),
          kLong: longPut.strike,
          kShort: shortPut.strike,
          debit: putDebit,
          flavor: "PUT",
          label: label + " (put)",
          expiration: longPut.expiration,
          symbol: lastSym,
        });
      }

      if (shortCall && longCall) {
        const callDebit = longCall.price - shortCall.price;
        let label = `IC ${longPut.strike}/${shortPut.strike}/${shortCall.strike}/${longCall.strike}`;
        if (shortCall.expiration) {
          const expLabel = formatExpirationLabel_(shortCall.expiration);
          if (expLabel) label = `${expLabel} ${label}`;
        }
        result.bearCallSpreads.push({
          qty: Math.abs(shortCall.qty),
          kLong: shortCall.strike,
          kShort: longCall.strike,
          debit: callDebit,
          flavor: "BEAR_CALL",
          label: label + " (call)",
          expiration: shortCall.expiration,
          symbol: lastSym,
        });
      }
    }
    // null/unknown types are silently skipped
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

  const skip = new Set(["REALIZED", "UNREALIZED", "TOTAL"]);
  return Array.from(symbols).filter(s => !skip.has(s)).sort();
}

/**
 * Returns the strategy abbreviation for writing back to the sheet.
 */
function strategyAbbrev_(type) {
  switch (type) {
    case "bull-call-spread": return "BCS";
    case "bull-put-spread": return "BPS";
    case "iron-condor": return "IC";
    case "stock": return "Stock";
    default: return "?";
  }
}

/**
 * Auto-fills Strategy column on the Legs sheet for each group.
 * @param {Sheet} sheet - The Legs sheet
 * @param {Range} range - The named range for the Legs table
 * @param {Array[]} rows - 2D values from the range
 */
function updateLegsSheetStrategy_(sheet, range, rows) {
  if (!rows || rows.length < 2) return;

  const headers = rows[0];
  const idxSym = findColumn_(headers, ["symbol", "ticker"]);
  const idxGroup = findColumn_(headers, ["group", "grp"]);
  const idxStrat = findColumn_(headers, ["strategy", "strat"]);
  const idxStrike = findColumn_(headers, ["strike", "strikeprice"]);
  const idxType = findColumn_(headers, ["type", "optiontype", "callput", "cp", "putcall", "legtype"]);
  const idxQty = findColumn_(headers, ["qty", "quantity", "contracts", "contract", "count", "shares"]);
  const idxPrice = findColumn_(headers, ["price", "cost", "entry", "premium", "basis", "costbasis", "avgprice", "pricepaid"]);

  // Disambiguate strategy vs type if they resolve to same column
  let stratCol = idxStrat;
  if (stratCol >= 0 && stratCol === idxType) {
    stratCol = findColumn_(headers, ["strategy", "strat"]);
  }
  if (stratCol < 0) return; // no strategy column to write

  // Group rows with carry-forward
  let lastSym = "";
  let lastGroup = "";
  const groups = []; // { legs: [{strike, type, qty}], firstRow: number }
  let currentGroup = null;

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const rawSym = String(row[idxSym] ?? "").trim().toUpperCase();
    if (rawSym) lastSym = rawSym;

    const rawGroup = String(row[idxGroup >= 0 ? idxGroup : -1] ?? "").trim();
    if (rawGroup) lastGroup = rawGroup;

    const qty = parseNumber_(row[idxQty]);
    if (!Number.isFinite(qty) || qty === 0) continue;

    const strike = idxStrike >= 0 ? parseNumber_(row[idxStrike]) : NaN;
    const optType = idxType >= 0 ? parseOptionType_(row[idxType]) : null;
    let legType = optType;
    if (!legType && !Number.isFinite(strike)) legType = "Stock";

    const groupKey = `${lastSym}|${lastGroup || r}`;
    if (!currentGroup || currentGroup.key !== groupKey) {
      currentGroup = { key: groupKey, legs: [], firstRow: r };
      groups.push(currentGroup);
    }
    currentGroup.legs.push({ strike, type: legType, qty });
  }

  // Write strategy for each group
  const rangeStartRow = range.getRow();
  const rangeStartCol = range.getColumn();

  for (const group of groups) {
    const posType = detectPositionType_(group.legs);
    const abbrev = strategyAbbrev_(posType);
    const sheetRow = rangeStartRow + group.firstRow; // firstRow is 1-based from rows array
    const sheetCol = rangeStartCol + stratCol;
    sheet.getRange(sheetRow, sheetCol).setValue(abbrev);
  }
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
    null,
    "same sign = null"
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
    "IC: 2 puts + 2 calls"
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
