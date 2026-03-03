/**
 * PositionBuilder.js
 * Broker-agnostic spread pairing and position aggregation.
 *
 * Takes normalized transactions and builds spread positions.
 * Can be used with any brokerage that provides transaction data.
 */

/**
 * Converts an expiration (Date or string) to M/D/YYYY format for consistent key matching.
 */
function formatExpirationForKey_(exp) {
  // Normalize all dates to M/D/YYYY format for consistent key matching
  if (exp instanceof Date) {
    return `${exp.getMonth() + 1}/${exp.getDate()}/${exp.getFullYear()}`;
  }
  const s = String(exp || "").trim();
  // Already in M/D/YYYY format
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) return s;
  // YYYY-MM-DD → M/D/YYYY
  const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    return `${parseInt(isoMatch[2], 10)}/${parseInt(isoMatch[3], 10)}/${isoMatch[1]}`;
  }
  return s;
}

/**
 * Builds map of latest transaction date per ticker from stock transactions.
 */
function buildLatestStockDates_(stockTxns) {
  const latestByTicker = new Map();

  if (!stockTxns) return latestByTicker;

  for (const txn of stockTxns) {
    const ticker = txn.ticker;
    if (!ticker) continue;

    const txnDate = parseDateAtMidnight_(txn.date);
    if (!txnDate) continue;

    const existing = latestByTicker.get(ticker);
    if (!existing || txnDate > existing) {
      latestByTicker.set(ticker, txnDate);
    }
  }

  return latestByTicker;
}

/**
 * Pairs consecutive opens on same date into spread orders.
 * Detects iron condors (2 puts + 2 calls with matching qty).
 * Applies closes to reduce open quantities (FIFO by date).
 *
 * @param {Object[]} transactions - Normalized option transactions
 * @returns {Object[]} Array of spread positions
 */
function pairTransactionsIntoSpreads_(transactions) {
  const spreads = [];

  // Group opens by date + ticker + expiration
  // Keep original quantities - closed positions will be marked via the Closed column
  const groups = new Map();
  for (const txn of transactions) {
    if (!txn.isOpen) continue;

    const key = `${txn.date}|${txn.ticker}|${txn.expiration}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ ...txn }); // Clone to avoid mutating original
  }

  // Pair within each group
  for (const [key, txns] of groups) {
    // Separate by option type first
    const calls = txns.filter(t => t.optionType === "Call");
    const puts = txns.filter(t => t.optionType === "Put");

    // Check for iron condor: 2 calls (1 long, 1 short) + 2 puts (1 long, 1 short)
    const longCalls = calls.filter(t => t.qty > 0);
    const shortCalls = calls.filter(t => t.qty < 0);
    const longPuts = puts.filter(t => t.qty > 0);
    const shortPuts = puts.filter(t => t.qty < 0);

    if (longCalls.length === 1 && shortCalls.length === 1 &&
        longPuts.length === 1 && shortPuts.length === 1) {
      // Check if quantities match (all same absolute qty)
      const qty = longCalls[0].qty;
      if (Math.abs(shortCalls[0].qty) === qty &&
          longPuts[0].qty === qty &&
          Math.abs(shortPuts[0].qty) === qty) {
        // Iron condor detected - create as 4 legs
        spreads.push({
          type: "iron-condor",
          ticker: longCalls[0].ticker,
          expiration: longCalls[0].expiration,
          qty: qty,
          date: longCalls[0].date,
          legs: [
            { strike: longPuts[0].strike, optionType: "Put", qty: qty, price: longPuts[0].price },
            { strike: shortPuts[0].strike, optionType: "Put", qty: -qty, price: shortPuts[0].price },
            { strike: shortCalls[0].strike, optionType: "Call", qty: -qty, price: shortCalls[0].price },
            { strike: longCalls[0].strike, optionType: "Call", qty: qty, price: longCalls[0].price },
          ].sort((a, b) => a.strike - b.strike),
        });
        continue; // Skip normal pairing for this group
      }
    }

    // Check for long straddle/strangle: 1 long call + 1 long put (no shorts)
    if (longCalls.length >= 1 && longPuts.length >= 1 &&
        shortCalls.length === 0 && shortPuts.length === 0) {
      const lc = longCalls[0];
      const lp = longPuts[0];
      const pairQty = Math.min(lc.qty, lp.qty);
      const isStraddle = lc.strike === lp.strike;

      spreads.push({
        type: isStraddle ? "long-straddle" : "long-strangle",
        ticker: lc.ticker,
        expiration: lc.expiration,
        qty: pairQty,
        date: lc.date,
        legs: [
          { strike: lp.strike, optionType: "Put", qty: pairQty, price: lp.price },
          { strike: lc.strike, optionType: "Call", qty: pairQty, price: lc.price },
        ].sort((a, b) => a.strike - b.strike),
      });

      // Reduce quantities for any remainder
      lc.qty -= pairQty;
      lp.qty -= pairQty;

      // If fully consumed, skip to next group
      if (lc.qty === 0 && lp.qty === 0 && longCalls.length === 1 && longPuts.length === 1) {
        continue;
      }
    }

    // Check for short straddle/strangle: 1 short call + 1 short put (no longs)
    if (shortCalls.length >= 1 && shortPuts.length >= 1 &&
        longCalls.length === 0 && longPuts.length === 0) {
      const sc = shortCalls[0];
      const sp = shortPuts[0];
      const pairQty = Math.min(Math.abs(sc.qty), Math.abs(sp.qty));
      const isStraddle = sc.strike === sp.strike;

      spreads.push({
        type: isStraddle ? "short-straddle" : "short-strangle",
        ticker: sc.ticker,
        expiration: sc.expiration,
        qty: pairQty,
        date: sc.date,
        legs: [
          { strike: sp.strike, optionType: "Put", qty: -pairQty, price: sp.price },
          { strike: sc.strike, optionType: "Call", qty: -pairQty, price: sc.price },
        ].sort((a, b) => a.strike - b.strike),
      });

      // Reduce quantities for any remainder
      sc.qty += pairQty; // negative + positive = less negative
      sp.qty += pairQty;

      // If fully consumed, skip to next group
      if (sc.qty === 0 && sp.qty === 0 && shortCalls.length === 1 && shortPuts.length === 1) {
        continue;
      }
    }

    // Normal pairing: pair calls with calls, puts with puts
    for (const optionType of ["Call", "Put"]) {
      const legsOfType = txns.filter(t => t.optionType === optionType);
      if (legsOfType.length === 0) continue;

      const longs = legsOfType.filter(t => t.qty > 0);
      const shorts = legsOfType.filter(t => t.qty < 0);

      // Clone and sort for pairing
      const longsToProcess = longs.map(t => ({ ...t })).sort((a, b) => a.strike - b.strike);
      const shortsToProcess = shorts.map(t => ({ ...t })).sort((a, b) => a.strike - b.strike);

      // Pair by matching quantities
      let li = 0, si = 0;
      while (li < longsToProcess.length && si < shortsToProcess.length) {
        const long = longsToProcess[li];
        const short = shortsToProcess[si];

        const pairQty = Math.min(long.qty, Math.abs(short.qty));

        spreads.push({
          ticker: long.ticker,
          expiration: long.expiration,
          lowerStrike: long.strike,
          upperStrike: short.strike,
          optionType: long.optionType,
          qty: pairQty,
          lowerPrice: long.price,
          upperPrice: short.price,
          date: long.date,
        });

        long.qty -= pairQty;
        short.qty += pairQty; // short.qty is negative, so adding makes it less negative

        if (long.qty === 0) li++;
        if (short.qty === 0) si++;
      }

      // Handle unmatched legs (naked positions)
      while (li < longsToProcess.length) {
        const long = longsToProcess[li];
        if (long.qty > 0) {
          spreads.push({
            ticker: long.ticker,
            expiration: long.expiration,
            lowerStrike: long.strike,
            upperStrike: null, // Naked long
            optionType: long.optionType,
            qty: long.qty,
            lowerPrice: long.price,
            upperPrice: 0,
            date: long.date,
          });
        }
        li++;
      }
      while (si < shortsToProcess.length) {
        const short = shortsToProcess[si];
        if (short.qty < 0) {
          spreads.push({
            ticker: short.ticker,
            expiration: short.expiration,
            lowerStrike: null, // Naked short
            upperStrike: short.strike,
            optionType: short.optionType,
            qty: short.qty,
            lowerPrice: 0,
            upperPrice: short.price,
            date: short.date,
          });
        }
        si++;
      }
    }
  }

  return spreads;
}

/**
 * Combines naked single-leg options across different dates into spreads.
 *
 * When a user sells a put on one date and buys a protective put on another date,
 * they end up as separate "naked" positions. This function combines them into
 * proper vertical spreads if they have matching ticker, expiration, and quantity.
 *
 * @param {Object[]} spreads - Array of spread positions from pairTransactionsIntoSpreads_
 * @returns {Object[]} Array with naked positions combined into spreads where possible
 */
function combineNakedLegsIntoSpreads_(spreads) {
  const result = [];

  // Separate naked puts and other positions
  const rawShortPuts = []; // lowerStrike: null
  const rawLongPuts = [];  // upperStrike: null
  const other = [];

  for (const sp of spreads) {
    if (sp.optionType === "Put" && sp.lowerStrike === null && sp.upperStrike !== null) {
      rawShortPuts.push(sp);
    } else if (sp.optionType === "Put" && sp.upperStrike === null && sp.lowerStrike !== null) {
      rawLongPuts.push(sp);
    } else {
      other.push(sp);
    }
  }

  // Merge duplicate naked legs (same ticker/expiration/strike) before combining
  // This handles cases where multiple transactions created separate entries for the same position
  const mergeNakedLegs = function(legs, isShort) {
    const merged = new Map();
    for (const leg of legs) {
      const exp = formatExpirationForKey_(leg.expiration);
      const strike = isShort ? leg.upperStrike : leg.lowerStrike;
      const key = `${leg.ticker}|${exp}|${strike}`;

      if (!merged.has(key)) {
        merged.set(key, { ...leg });
      } else {
        const existing = merged.get(key);
        const oldQty = Math.abs(existing.qty);
        const newQty = Math.abs(leg.qty);
        const totalQty = oldQty + newQty;

        // Weighted average price
        const priceField = isShort ? 'upperPrice' : 'lowerPrice';
        existing[priceField] = (oldQty * existing[priceField] + newQty * leg[priceField]) / totalQty;
        existing.qty = isShort ? -totalQty : totalQty;

        // Keep latest date
        if (leg.date && (!existing.date || leg.date > existing.date)) {
          existing.date = leg.date;
        }
      }
    }
    return Array.from(merged.values());
  };

  const nakedShortPuts = mergeNakedLegs(rawShortPuts, true);
  const nakedLongPuts = mergeNakedLegs(rawLongPuts, false);

  // Try to match naked short puts with naked long puts (supports partial matching)
  // Clone arrays since we'll modify quantities
  const shortsToProcess = nakedShortPuts.map(sp => ({ ...sp, remainingQty: Math.abs(sp.qty) }));
  const longsToProcess = nakedLongPuts.map(lp => ({ ...lp, remainingQty: lp.qty }));

  for (const shortPut of shortsToProcess) {
    const exp = formatExpirationForKey_(shortPut.expiration);

    // Find matching long puts: same ticker, same expiration, lower strike
    for (const longPut of longsToProcess) {
      if (longPut.remainingQty <= 0) continue;

      const longExp = formatExpirationForKey_(longPut.expiration);

      if (longPut.ticker === shortPut.ticker &&
          longExp === exp &&
          longPut.lowerStrike < shortPut.upperStrike &&
          shortPut.remainingQty > 0) {

        // Match as many contracts as possible
        const matchQty = Math.min(shortPut.remainingQty, longPut.remainingQty);

        // Combine into bull put spread
        result.push({
          ticker: shortPut.ticker,
          expiration: shortPut.expiration,
          lowerStrike: longPut.lowerStrike,
          upperStrike: shortPut.upperStrike,
          optionType: "Put",
          qty: matchQty,
          lowerPrice: longPut.lowerPrice,
          upperPrice: shortPut.upperPrice,
          date: shortPut.date,
        });

        shortPut.remainingQty -= matchQty;
        longPut.remainingQty -= matchQty;
        log.info("combine", `Combined naked puts into spread: ${shortPut.ticker} ${exp} ${longPut.lowerStrike}/${shortPut.upperStrike} x${matchQty}`);
      }
    }

    // Add any remaining short quantity as naked short put
    if (shortPut.remainingQty > 0) {
      result.push({
        ...shortPut,
        qty: -shortPut.remainingQty, // Restore negative sign for short
      });
      delete result[result.length - 1].remainingQty;
    }
  }

  // Add remaining long puts
  for (const longPut of longsToProcess) {
    if (longPut.remainingQty > 0) {
      result.push({
        ...longPut,
        qty: longPut.remainingQty,
      });
      delete result[result.length - 1].remainingQty;
    }
  }

  // Add all other positions
  result.push(...other);

  return result;
}

/**
 * Combines matching bull-put-spreads and bear-call-spreads into iron condors/butterflies.
 *
 * When a user has separate put and call credit spreads on the same underlying with
 * matching expiration and quantity, they form an iron condor (or iron butterfly if
 * the short strikes match).
 *
 * Spread structure from pairTransactionsIntoSpreads_:
 * - lowerStrike = LONG leg's strike
 * - upperStrike = SHORT leg's strike
 * - qty = always positive (contract count)
 *
 * So:
 * - Bull put spread: Put, lowerStrike < upperStrike (long lower, short higher)
 * - Bear call spread: Call, lowerStrike > upperStrike (long higher, short lower)
 *
 * @param {Object[]} spreads - Array of spread positions
 * @returns {Object[]} Array with matching spreads combined into iron condors/butterflies
 */
function combineRelatedSpreadsIntoIronCondorsAndButterflies_(spreads) {
  const result = [];

  // Identify bull-put-spreads and bear-call-spreads
  const bullPutSpreads = [];
  const bearCallSpreads = [];
  const other = [];

  for (const sp of spreads) {
    // Bull put spread: Put, long lower strike, short higher strike
    // lowerStrike (long) < upperStrike (short)
    if (sp.optionType === "Put" && sp.lowerStrike != null && sp.upperStrike != null &&
        sp.lowerStrike < sp.upperStrike && sp.qty > 0) {
      bullPutSpreads.push(sp);
    }
    // Bear call spread: Call, long higher strike, short lower strike
    // lowerStrike (long) > upperStrike (short)
    else if (sp.optionType === "Call" && sp.lowerStrike != null && sp.upperStrike != null &&
             sp.lowerStrike > sp.upperStrike && sp.qty > 0) {
      bearCallSpreads.push(sp);
    }
    else {
      other.push(sp);
    }
  }

  // Try to match bull-put-spreads with bear-call-spreads
  const usedBullPuts = new Set();
  const usedBearCalls = new Set();

  for (let i = 0; i < bullPutSpreads.length; i++) {
    const bps = bullPutSpreads[i];
    const bpsExp = formatExpirationForKey_(bps.expiration);
    const bpsQty = bps.qty;

    for (let j = 0; j < bearCallSpreads.length; j++) {
      if (usedBearCalls.has(j)) continue;

      const bcs = bearCallSpreads[j];
      const bcsExp = formatExpirationForKey_(bcs.expiration);
      const bcsQty = bcs.qty;

      // Match criteria: same ticker, same expiration, same quantity
      if (bps.ticker === bcs.ticker && bpsExp === bcsExp && bpsQty === bcsQty) {
        // For bull put spread: upperStrike is the SHORT put
        // For bear call spread: upperStrike is the SHORT call
        const shortPutStrike = bps.upperStrike;
        const shortCallStrike = bcs.upperStrike;

        // Determine if iron condor or iron butterfly
        let strategyType;
        if (shortPutStrike === shortCallStrike) {
          strategyType = "iron-butterfly";
        } else if (shortPutStrike < shortCallStrike) {
          strategyType = "iron-condor";
        } else {
          // Short strikes overlap (put > call), not a valid IC/IB, skip
          continue;
        }

        // Create the combined position with all 4 legs
        // Bull put: long bps.lowerStrike put, short bps.upperStrike put
        // Bear call: short bcs.upperStrike call, long bcs.lowerStrike call
        result.push({
          type: strategyType,
          ticker: bps.ticker,
          expiration: bps.expiration,
          qty: bpsQty,
          date: bps.date > bcs.date ? bps.date : bcs.date,
          legs: [
            { strike: bps.lowerStrike, optionType: "Put", qty: bpsQty, price: bps.lowerPrice },
            { strike: bps.upperStrike, optionType: "Put", qty: -bpsQty, price: bps.upperPrice },
            { strike: bcs.upperStrike, optionType: "Call", qty: -bcsQty, price: bcs.upperPrice },
            { strike: bcs.lowerStrike, optionType: "Call", qty: bcsQty, price: bcs.lowerPrice },
          ].sort((a, b) => a.strike - b.strike),
        });

        usedBullPuts.add(i);
        usedBearCalls.add(j);
        log.info("combine", `Combined ${bps.ticker} ${bpsExp} into ${strategyType}: ` +
          `${bps.lowerStrike}/${bps.upperStrike}/${bcs.upperStrike}/${bcs.lowerStrike} x${bpsQty}`);
        break;
      }
    }
  }

  // Add unmatched spreads
  for (let i = 0; i < bullPutSpreads.length; i++) {
    if (!usedBullPuts.has(i)) {
      result.push(bullPutSpreads[i]);
    }
  }
  for (let j = 0; j < bearCallSpreads.length; j++) {
    if (!usedBearCalls.has(j)) {
      result.push(bearCallSpreads[j]);
    }
  }

  // Add all other positions
  result.push(...other);

  return result;
}

/**
 * Builds a map of closing prices from close transactions.
 * Key: "TICKER|EXPIRATION|STRIKE|TYPE" -> price
 *
 * Handles:
 * 1. Sold To Close / Bought To Cover → use transaction price
 * 2. Option Exercised / Option Assigned → compute intrinsic from stock transactions
 * 3. Expired worthless (expiration < today, no close) → set to 0
 *
 * For multiple closes of same leg, uses weighted average.
 */
function buildClosingPricesMap_(transactions, stockTxns) {
  const result = new Map();
  const closes = new Map(); // key -> { totalQty, totalValue }

  // 1. Normal closes (Sold To Close, Bought To Cover)
  for (const txn of transactions) {
    if (!txn.isClosed) continue;

    // Normalize expiration for consistent key matching
    const exp = formatExpirationForKey_(txn.expiration);
    const key = `${txn.ticker}|${exp}|${txn.strike}|${txn.optionType}`;
    const qty = Math.abs(txn.qty);
    const value = qty * txn.price;

    if (!closes.has(key)) closes.set(key, { totalQty: 0, totalValue: 0 });
    const entry = closes.get(key);
    entry.totalQty += qty;
    entry.totalValue += value;
  }

  for (const [key, { totalQty, totalValue }] of closes) {
    if (totalQty > 0) {
      result.set(key, roundTo_(totalValue / totalQty, 2));
    }
  }

  // 2. Exercise/Assignment → compute intrinsic from stock transactions
  const stockByDateTicker = new Map(); // "date|ticker" -> [prices]
  for (const stk of (stockTxns || [])) {
    const key = `${stk.date}|${stk.ticker}`;
    if (!stockByDateTicker.has(key)) stockByDateTicker.set(key, []);
    stockByDateTicker.get(key).push(stk.price);
  }

  for (const txn of transactions) {
    if (!txn.isExercised && !txn.isAssigned) continue;

    const exp = formatExpirationForKey_(txn.expiration);
    const key = `${txn.ticker}|${exp}|${txn.strike}|${txn.optionType}`;
    if (result.has(key)) continue; // Already have a closing price

    const stkKey = `${txn.date}|${txn.ticker}`;
    const stockPrices = stockByDateTicker.get(stkKey) || [];

    if (stockPrices.length > 0) {
      const marketPrice = Math.max(...stockPrices);
      let intrinsic;
      if (txn.optionType === "Call") {
        intrinsic = Math.max(0, marketPrice - txn.strike);
      } else {
        intrinsic = Math.max(0, txn.strike - marketPrice);
      }
      result.set(key, roundTo_(intrinsic, 2));
    }
  }

  // 3. Expired worthless: if expiration < today and no close, set to 0
  const today = new Date();
  const openLegs = new Set();
  for (const txn of transactions) {
    if (!txn.isOpen) continue;
    const exp = formatExpirationForKey_(txn.expiration);
    const key = `${txn.ticker}|${exp}|${txn.strike}|${txn.optionType}`;
    openLegs.add(key);
  }

  for (const legKey of openLegs) {
    if (result.has(legKey)) continue;

    const parts = legKey.split("|");
    const expStr = parts[1];
    const expDate = parseDateAtMidnight_(expStr);
    if (expDate && expDate < today) {
      result.set(legKey, 0); // Expired worthless
    }
  }

  return result;
}

/**
 * Creates a unique key for a spread from its legs.
 */
function makeSpreadKey_(legs) {
  if (legs.length === 0) return null;

  const ticker = legs[0].symbol;

  // Cash positions
  if (legs.length === 1 && (legs[0].type === "Cash" || ticker === "CASH")) {
    return "CASH|CASH";
  }

  // Stock positions
  if (legs.length === 1 && (legs[0].type === "Stock" || !Number.isFinite(legs[0].strike))) {
    return `${ticker}|STOCK`;
  }

  const exp = normalizeExpiration_(legs[0].expiration) || legs[0].expiration;
  const strikes = legs.map(l => l.strike).sort((a, b) => a - b);

  // Detect iron-condor/iron-butterfly: 4 legs with both puts and calls
  const types = new Set(legs.map(l => l.type));
  if (legs.length === 4 && types.has("Put") && types.has("Call")) {
    return `${ticker}|${exp}|${strikes.join("/")}|IC`;
  }

  const type = legs[0].type || "Call";
  return `${ticker}|${exp}|${strikes.join("/")}|${type}`;
}

/**
 * Creates spread key from a spread order.
 */
function makeSpreadKeyFromOrder_(spread) {
  if (spread.type === "stock") {
    return `${spread.ticker}|STOCK`;
  }

  if (spread.type === "cash") {
    return "CASH|CASH";
  }

  const exp = normalizeExpiration_(spread.expiration) || spread.expiration;

  // Multi-leg strategies with spread.legs
  if (spread.legs && spread.legs.length > 0) {
    const strikes = spread.legs.map(l => l.strike).sort((a, b) => a - b);
    let typeKey;
    switch (spread.type) {
      case "iron-condor": typeKey = "IC"; break;
      case "long-straddle": typeKey = "LS"; break;
      case "short-straddle": typeKey = "SS"; break;
      case "long-strangle": typeKey = "LSg"; break;
      case "short-strangle": typeKey = "SSg"; break;
      default: typeKey = spread.type || "?";
    }
    return `${spread.ticker}|${exp}|${strikes.join("/")}|${typeKey}`;
  }

  const strikes = [spread.lowerStrike, spread.upperStrike].filter(s => s != null).sort((a, b) => a - b);
  return `${spread.ticker}|${exp}|${strikes.join("/")}|${spread.optionType}`;
}

/**
 * Pre-merges spreads with the same key, keeping the latest date and summing quantities.
 */
function preMergeSpreads_(spreads) {
  const merged = new Map();

  for (const spread of spreads) {
    const key = makeSpreadKeyFromOrder_(spread);

    if (!merged.has(key)) {
      merged.set(key, { ...spread });
    } else {
      const existing = merged.get(key);

      // Keep the later date
      if (spread.date && (!existing.date || spread.date > existing.date)) {
        existing.date = spread.date;
      }

      // Sum quantities and compute weighted average prices
      if (spread.type === "stock") {
        const oldQty = existing.qty || 0;
        const newQty = spread.qty || 0;
        const totalQty = oldQty + newQty;
        if (totalQty !== 0) {
          existing.price = ((oldQty * (existing.price || 0)) + (newQty * (spread.price || 0))) / totalQty;
        }
        existing.qty = totalQty;
      } else if (spread.type === "iron-condor" && spread.legs) {
        existing.qty = (existing.qty || 0) + (spread.qty || 0);
      } else {
        const oldQty = existing.qty || 0;
        const newQty = spread.qty || 0;
        const totalQty = oldQty + newQty;

        if (totalQty !== 0 && existing.lowerPrice !== undefined && spread.lowerPrice !== undefined) {
          existing.lowerPrice = ((oldQty * existing.lowerPrice) + (newQty * spread.lowerPrice)) / totalQty;
        }
        if (totalQty !== 0 && existing.upperPrice !== undefined && spread.upperPrice !== undefined) {
          existing.upperPrice = ((oldQty * existing.upperPrice) + (newQty * spread.upperPrice)) / totalQty;
        }
        existing.qty = totalQty;
      }
    }
  }

  return Array.from(merged.values());
}

/**
 * Merges new spreads into existing positions.
 * Skips spreads whose date is not newer than the group's LastTxnDate.
 * Returns { updatedLegs, newLegs, skippedCount }.
 */
function mergeSpreads_(existingPositions, newSpreads) {
  const updatedLegs = [];
  const newLegs = [];
  const processedKeys = new Set();
  let skippedCount = 0;
  let updatedCount = 0;

  for (const spread of newSpreads) {
    const key = makeSpreadKeyFromOrder_(spread);

    if (existingPositions.has(key)) {
      const existing = existingPositions.get(key);

      // Stock positions: add delta qty and update lastTxnDate
      if (spread.type === "stock") {
        if (spread.qty === 0 && !spread.date) {
          skippedCount++;
          continue;
        }

        const stockLeg = existing.legs[0];
        if (stockLeg) {
          stockLeg.qty += spread.qty;
          if (spread.price) stockLeg.price = spread.price;
        }

        if (spread.date) {
          existing.lastTxnDate = spread.date;
        }

        existing.updated = true;
        updatedCount++;
        for (const leg of existing.legs) {
          leg.updated = true;
          updatedLegs.push(leg);
        }
        continue;
      }

      // Cash positions: just update the amount
      if (spread.type === "cash") {
        const cashLeg = existing.legs[0];
        if (cashLeg) {
          cashLeg.price = spread.price;
          cashLeg.updated = true;
          updatedLegs.push(cashLeg);
        }
        continue;
      }

      // Per-group dedup: skip if spread's date is not newer than LastTxnDate
      const spreadDate = parseDateAtMidnight_(spread.date);
      const lastTxnDate = parseDateAtMidnight_(existing.lastTxnDate);

      if (spreadDate && lastTxnDate && spreadDate <= lastTxnDate) {
        skippedCount++;
        continue;
      }

      existing.debugReason = ` [txn:${spread.date} vs last:${existing.lastTxnDate}]`;

      // Merge into existing
      const longLeg = existing.legs.find(l => l.qty > 0);
      const shortLeg = existing.legs.find(l => l.qty < 0);

      if (longLeg && spread.lowerStrike) {
        const oldQty = longLeg.qty;
        const newQty = spread.qty;
        const totalQty = oldQty + newQty;
        longLeg.price = (oldQty * longLeg.price + newQty * spread.lowerPrice) / totalQty;
        longLeg.qty = totalQty;
      }

      if (shortLeg && spread.upperStrike) {
        const oldQty = Math.abs(shortLeg.qty);
        const newQty = spread.qty;
        const totalQty = oldQty + newQty;
        shortLeg.price = (oldQty * shortLeg.price + newQty * spread.upperPrice) / totalQty;
        shortLeg.qty = -(totalQty);
      }

      if (spread.date > (existing.lastTxnDate || "")) {
        existing.lastTxnDate = spread.date;
      }

      if (!processedKeys.has(key)) {
        updatedLegs.push(existing);
        processedKeys.add(key);
      }
    } else {
      newLegs.push(spread);
    }
  }

  return { updatedLegs, newLegs, skippedCount };
}

/**
 * Aggregates stock transactions into net positions.
 * Used when no portfolio CSV is available or for incremental updates.
 *
 * @param {Object[]} stockTxns - Stock transaction records
 * @param {Map} [sinceByTicker] - Optional map of ticker -> cutoff date to skip older transactions
 * @returns {Object[]} Array of stock position objects
 */
function aggregateStockTransactions_(stockTxns, sinceByTicker) {
  if (!stockTxns || stockTxns.length === 0) return [];

  const byTicker = new Map();

  for (const txn of stockTxns) {
    const ticker = txn.ticker;
    if (!ticker) continue;

    const txnDate = parseDateAtMidnight_(txn.date);

    // If sinceByTicker provided, skip transactions on or before the cutoff date
    if (sinceByTicker && sinceByTicker.has(ticker)) {
      const cutoff = sinceByTicker.get(ticker);
      if (txnDate && cutoff && txnDate <= cutoff) {
        continue;
      }
    }

    if (!byTicker.has(ticker)) {
      byTicker.set(ticker, {
        qty: 0,
        lastDate: null,
        lastPrice: 0,
      });
    }

    const entry = byTicker.get(ticker);

    // Accumulate quantity: Bought adds, Sold subtracts
    const qtyChange = txn.txnType === "Bought" ? txn.qty : -txn.qty;
    entry.qty += qtyChange;

    // Track latest transaction date
    if (txnDate && (!entry.lastDate || txnDate > entry.lastDate)) {
      entry.lastDate = txnDate;
      entry.lastPrice = txn.price;
    }
  }

  // Convert to spread-like objects
  const stocks = [];
  for (const [ticker, entry] of byTicker) {
    if (entry.qty === 0 && !entry.lastDate) continue;

    stocks.push({
      type: "stock",
      ticker: ticker,
      qty: entry.qty,
      price: entry.lastPrice,
      date: entry.lastDate,
      expiration: null,
      lowerStrike: null,
      upperStrike: null,
      optionType: "Stock",
    });
  }

  return stocks;
}
