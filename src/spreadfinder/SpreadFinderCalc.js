/**
 * SpreadFinderCalc.js
 * Calculation functions for SpreadFinder.
 */

/**
 * Calculates the Expected Gain for a Bull Call Spread based on an 80%-of-max-profit early exit.
 * Uses the "Rule of Touch" (probTouch ≈ 1.6x delta) to estimate probability of reaching target.
 * @param {number} longMid The mid price of the lower (long) leg.
 * @param {number} shortMid The mid price of the upper (short) leg.
 * @param {number} longStrike The strike price of the lower leg.
 * @param {number} shortStrike The strike price of the upper leg.
 * @param {number} shortDelta The delta of the upper (short) leg.
 * @return {number} The expected dollar gain per spread.
 */
function calculateExpectedGain(longMid, shortMid, longStrike, shortStrike, shortDelta) {
  const netDebit = longMid - shortMid;
  const spreadWidth = shortStrike - longStrike;
  const maxProfit = spreadWidth - netDebit;
  const targetProfit = maxProfit * 0.80;

  // Prob(Touch) ≈ 1.6x short delta, capped at 95%
  const probTouch = Math.min(shortDelta * 1.6, 0.95);
  const probLoss = 1 - probTouch;

  // EV = (Prob of Win * Win Amount) + (Prob of Loss * Loss Amount)
  return (probTouch * targetProfit) + (probLoss * -netDebit);
}

/**
 * Estimates current stock price from call options by finding the strike with delta closest to 0.5.
 * @param {Array} calls Array of call option objects with delta and strike.
 * @return {number} Estimated current stock price.
 */
function estimateCurrentPrice_(calls) {
  let bestDelta = Infinity;
  let bestStrike = 0;
  for (const c of calls) {
    const dist = Math.abs(Math.abs(c.delta) - 0.5);
    if (dist < bestDelta) {
      bestDelta = dist;
      bestStrike = c.strike;
    }
  }
  return bestStrike || 0;
}

/**
 * Generates all valid spreads from a sorted chain of calls.
 * Returns array of spread objects with metrics.
 */
function generateSpreads_(chain, config) {
  const spreads = [];
  const n = chain.length;

  for (let i = 0; i < n; i++) {
    const lower = chain[i];
    for (let j = i + 1; j < n; j++) {
      const upper = chain[j];
      const width = upper.strike - lower.strike;

      // Skip if too wide
      if (width < config.minSpreadWidth || width > config.maxSpreadWidth) continue;

      // Skip if no valid bid/ask
      if (lower.ask <= 0 || upper.bid < 0) continue;

      // Calculate debit using mid pricing (validated by executed trades with GT 60 patience)
      // For patient orders, mid is achievable; below mid may not fill
      const lowerMid = (lower.bid + lower.ask) / 2;
      const upperMid = (upper.bid + upper.ask) / 2;

      let debit = lowerMid - upperMid;
      if (debit < 0) debit = 0;
      debit = roundTo_(debit, 2);

      // Calculate metrics
      const maxProfit = width - debit;
      const maxLoss = debit;
      const roi = debit > 0 ? maxProfit / debit : 0;

      // Liquidity score: geometric mean of OI, scaled
      const minOI = Math.min(lower.openint, upper.openint);
      const liquidityScore = Math.sqrt(lower.openint * upper.openint) / 100;

      // Bid-ask tightness (lower is better, so invert)
      const lowerSpread = lower.ask - lower.bid;
      const upperSpread = upper.ask - upper.bid;
      const avgBidAskSpread = (lowerSpread + upperSpread) / 2;
      const tightness = avgBidAskSpread > 0 ? 1 / avgBidAskSpread : 10;

      // Expected gain using probability-of-touch model (80% of max profit target)
      const expectedGain = calculateExpectedGain(lowerMid, upperMid, lower.strike, upper.strike, Math.abs(upper.delta));
      const expectedROI = debit > 0 ? expectedGain / debit : 0;

      // Fitness = ExpROI * liquidity^0.1 * tightness^0.1
      // Liquidity/tightness as mild tiebreakers (patient fills assumed)
      // timeFactor dropped — already baked into delta and probTouch
      // Outlook boost: adjust fitness based on price target, date, and confidence
      let outlookBoost = 1;
      if (config.outlookFuturePrice > 0 && config.outlookConfidence > 0) {
        const target = config.outlookFuturePrice;
        const conf = config.outlookConfidence;

        // Price proximity boost
        let priceBoost;
        if (lower.strike >= target) {
          // Both strikes above target — graduated penalty (further above = worse)
          const overshoot = (lower.strike - target) / target;
          priceBoost = 1 - conf * 0.5 * overshoot;
        } else if (upper.strike <= target) {
          // Both strikes below target — full boost by proximity
          priceBoost = 1 + conf * (upper.strike / target);
        } else {
          // Straddles target — partial boost by how much width is captured
          const captured = (target - lower.strike) / width;
          priceBoost = 1 + conf * captured * 0.5;
        }

        // Date proximity boost: expirations near outlookDate get more boost
        // Expirations before target date are penalized (may expire before move happens)
        let dateBoost = 1;
        if (config.outlookDate) {
          const expDate = parseDateAtMidnight_(lower.expiration);
          const targetDate = parseDateAtMidnight_(config.outlookDate);
          const now = new Date();
          if (!expDate || !targetDate) continue;
          const totalDays = Math.max(1, (targetDate - now) / (1000 * 60 * 60 * 24));
          const diffDays = (expDate - targetDate) / (1000 * 60 * 60 * 24);

          if (diffDays < 0) {
            // Expires before target date — penalize proportionally
            // Expiring way before target = bigger penalty
            const earlyRatio = Math.abs(diffDays) / totalDays;
            dateBoost = 1 - conf * Math.min(earlyRatio, 0.5);
          } else {
            // Expires on or after target date — boost, with falloff for much later
            const lateRatio = diffDays / totalDays;
            dateBoost = 1 + conf * Math.max(0, 0.3 - lateRatio * 0.2);
          }
        }

        outlookBoost = priceBoost * dateBoost;
      }

      const fitness = roundTo_(expectedROI * Math.pow(liquidityScore, 0.2) * Math.pow(tightness, 0.1) * outlookBoost, 2);

      spreads.push({
        symbol: lower.symbol,
        expiration: lower.expiration,
        lowerStrike: lower.strike,
        upperStrike: upper.strike,
        width,
        debit,
        maxProfit: roundTo_(maxProfit, 2),
        maxLoss: roundTo_(maxLoss, 2),
        roi: roundTo_(roi, 2),
        lowerIV: roundTo_(lower.iv || 0, 2),
        lowerDelta: roundTo_(lower.delta, 2),
        upperDelta: roundTo_(upper.delta, 2),
        lowerOI: lower.openint,
        upperOI: upper.openint,
        lowerVol: lower.volume,
        upperVol: upper.volume,
        expectedGain: roundTo_(expectedGain, 2),
        expectedROI: roundTo_(expectedROI, 2),
        liquidityScore: roundTo_(liquidityScore, 2),
        tightness: roundTo_(tightness, 2),
        fitness: roundTo_(fitness, 2)
      });
    }
  }

  return spreads;
}
