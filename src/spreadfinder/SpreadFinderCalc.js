/**
 * SpreadFinderCalc.js
 * Calculation functions for SpreadFinder.
 */

/**
 * Calculates a liquidity score from 0 (illiquid) to 1 (highly liquid).
 * Weighted composite: 60% bid-ask spread, 25% volume, 15% open interest.
 * @param {number} bid The current bid price.
 * @param {number} ask The current ask price.
 * @param {number} volume Daily trading volume.
 * @param {number} openInterest Total open interest.
 * @return {number} Liquidity score between 0 and 1.
 */
function calcLiquidityScore(bid, ask, volume, openInterest) {
  if (bid <= 0 || ask <= 0) return 0;

  const mid = (bid + ask) / 2;
  const relSpread = (ask - bid) / mid;

  // 1. Spread Score: Penalty starts at 0.5% spread, hits zero at 5% spread
  const spreadScore = Math.max(0, 1 - (relSpread / 0.05));

  // 2. Volume Score: Logarithmic scale, 1000 is "good" liquidity
  const volScore = Math.min(1, Math.log10((volume || 0) + 1) / 3);

  // 3. OI Score: Ensures the market has depth
  const oiScore = Math.min(1, Math.log10((openInterest || 0) + 1) / 4);

  // Weighted Average: Spread is most important for immediate entry/exit
  const totalScore = (spreadScore * 0.6) + (volScore * 0.25) + (oiScore * 0.15);

  return Math.round(totalScore * 100) / 100;
}

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
 * @deprecated Use generateCallSpreads_ instead
 */
function generateSpreads_(chain, config) {
  return generateCallSpreads_(chain, config);
}

/**
 * Generates all valid call spreads from a sorted chain of calls.
 * Uses the new Outlook system with pro-rated targets.
 * @param {Array} chain - Sorted array of call options
 * @param {Object} config - Config including outlook data
 * @returns {Array} Array of spread objects with metrics
 */
function generateCallSpreads_(chain, config) {
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

      // Calculate debit using mid pricing
      const lowerMid = (lower.bid + lower.ask) / 2;
      const upperMid = (upper.bid + upper.ask) / 2;

      let debit = lowerMid - upperMid;
      if (debit < 0) debit = 0;
      debit = roundTo_(debit, 2);

      // Calculate metrics
      const maxProfit = width - debit;
      const maxLoss = debit;
      const roi = debit > 0 ? maxProfit / debit : 0;

      // Liquidity score: minimum of both legs (weakest link)
      const lowerLiquidity = calcLiquidityScore(lower.bid, lower.ask, lower.volume, lower.openint);
      const upperLiquidity = calcLiquidityScore(upper.bid, upper.ask, upper.volume, upper.openint);
      const liquidityScore = Math.min(lowerLiquidity, upperLiquidity);

      // Expected gain using probability-of-touch model (80% of max profit target)
      const expectedGain = calculateExpectedGain(lowerMid, upperMid, lower.strike, upper.strike, Math.abs(upper.delta));
      const expectedROI = debit > 0 ? expectedGain / debit : 0;

      // Outlook boost using pro-rated target from Outlook sheet
      let outlookBoost = 1;
      const outlook = config.outlook;
      if (outlook && outlook.proRatedTarget > 0 && outlook.confidence > 0) {
        const target = outlook.proRatedTarget;
        const conf = outlook.confidence;

        // Price proximity boost based on pro-rated target
        let priceBoost;
        if (lower.strike >= target) {
          // Both strikes above target — graduated penalty
          const overshoot = (lower.strike - target) / target;
          priceBoost = 1 - conf * 0.5 * overshoot;
        } else if (upper.strike <= target) {
          // Both strikes below target — full boost by proximity
          priceBoost = 1 + conf * (upper.strike / target);
        } else {
          // Straddles target — partial boost
          const captured = (target - lower.strike) / width;
          priceBoost = 1 + conf * captured * 0.5;
        }

        outlookBoost = priceBoost;
      }

      const fitness = roundTo_(expectedROI * Math.pow(liquidityScore, 0.2) * outlookBoost, 2);

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
        fitness: roundTo_(fitness, 2)
      });
    }
  }

  return spreads;
}
