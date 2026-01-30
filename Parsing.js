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

  // Test 1: Standard order
  const csv1 = [
    ["Strike", "Bid", "Mid", "Ask", "Type"],
    ["450", "203.15", "206.00", "208.85", "Call"],
    ["350", "250.00", "255.00", "260.00", "Put"]
  ];
  const rows1 = loadCsvData_(csv1, symbol, expDate);
  assertArrayDeepEqual(rows1, [
    [symbol, expDate, 450, "Call", 203.15, 206.00, 208.85],
    [symbol, expDate, 350, "Put", 250.00, 255.00, 260.00]
  ], "Test 1: Standard order");

  // Test 2: Reversed order
  const csv2 = [
    ["Ask", "Bid", "Mid", "Type", "Strike"],
    ["208.85", "203.15", "206.00", "Call", "450"],
    ["260.00", "250.00", "255.00", "Put", "350"]
  ];
  const rows2 = loadCsvData_(csv2, symbol, expDate);
  assertArrayDeepEqual(rows2, [
    [symbol, expDate, 450, "Call", 203.15, 206.00, 208.85],
    [symbol, expDate, 350, "Put", 250.00, 255.00, 260.00]
  ], "Test 2: Reversed order");

  // Test 3: Mixed case headers, no Mid, alt Type name
  const csv3 = [
    ["sTriKe", "BID", "ASK", "Option Type"],
    ["450", "203.15", "208.85", "Call"],
    ["350", "250.00", "260.00", "Put"]
  ];
  const rows3 = loadCsvData_(csv3, symbol, expDate);
  assertArrayDeepEqual(rows3, [
    [symbol, expDate, 450, "Call", 203.15, null, 208.85],
    [symbol, expDate, 350, "Put", 250.00, null, 260.00]
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
