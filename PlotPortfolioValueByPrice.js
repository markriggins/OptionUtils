/**
 * PlotPortfolioValueByPrice.js
 * ------------------------------------------------------------
 * Displays portfolio performance graphs in a modal dialog.
 *
 * Reads positions from the Portfolio (or PortfolioTable) named range,
 * extracts unique symbols, and shows interactive charts.
 *
 * Supported position types:
 *   - Stock/shares
 *   - Bull call spreads, bull put spreads, bear call spreads
 *   - Iron condors (split into put + call spreads)
 *   - Single-leg options: long calls, short calls, long puts, short puts
 *
 * Charts displayed:
 *   1. Portfolio Value ($) - Shares, strategy groups, total
 *   2. Portfolio ROI (%) - Return on investment by category
 *   3. Individual Spreads and Options ($) - Each position's value
 *   4. Individual Spreads and Options ROI (%) - Each position's return
 *
 * Features:
 *   - Toggle between "At Expiration" and "Current" value modes
 *   - Double-click spreads/options to open in OptionStrat
 *   - Current price line overlay from GOOGLEFINANCE
 */

/* =========================================================
   Entry point
   ========================================================= */

// Store selected symbol for the graph dialog
var selectedSymbolForGraph_ = null;

function PlotPortfolioValueByPrice() {
  Logger.log("PlotPortfolioValueByPrice Started");

  const ss = SpreadsheetApp.getActive();

  // Get unique symbols from position tables
  const symbols = getUniqueSymbolsFromPositions_(ss);

  if (symbols.length === 0) {
    SpreadsheetApp.getUi().alert(
      "No symbols found in position tables.\n\n" +
        "Add positions to tables on the Positions sheet:\n" +
        "  - BullCallSpreadsTable (with Symbol column)\n" +
        "  - Stocks (with Symbol column)\n" +
        "  - IronCondorsTable (with Symbol column)\n"
    );
    return;
  }

  // If only one symbol, show graphs directly
  if (symbols.length === 1) {
    plotSelectedSymbols(symbols);
    return;
  }

  // Multiple symbols - show symbol selection dialog first
  const html = HtmlService.createHtmlOutputFromFile("SelectSymbols")
    .setWidth(350)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Select Symbol for Portfolio Graphs");
}

/**
 * Returns the list of available symbols. Called by the dialog on load.
 */
function getAvailableSymbols() {
  const ss = SpreadsheetApp.getActive();
  return getUniqueSymbolsFromPositions_(ss);
}

/**
 * Shows the portfolio graphs modal for selected symbols.
 */
function plotSelectedSymbols(symbols) {
  if (!symbols || symbols.length === 0) return;

  // Store the first symbol for the graph data function
  const props = PropertiesService.getDocumentProperties();
  props.setProperty("portfolioGraphSymbol", symbols[0]);

  // Show the portfolio graphs modal (25% larger)
  const html = HtmlService.createHtmlOutputFromFile("PortfolioGraphs")
    .setWidth(1500)
    .setHeight(1125);
  SpreadsheetApp.getUi().showModalDialog(html, symbols[0] + " Portfolio Performance");
}

/**
 * Returns portfolio graph data for the selected symbol.
 * Called by PortfolioGraphs.html via google.script.run.
 */
function getPortfolioGraphData() {
  try {
    Logger.log("getPortfolioGraphData: Starting...");
    const ss = SpreadsheetApp.getActive();
    const props = PropertiesService.getDocumentProperties();
    const symbol = props.getProperty("portfolioGraphSymbol") || "";

    Logger.log("getPortfolioGraphData: symbol = " + symbol);

    if (!symbol) {
      throw new Error("No symbol selected for portfolio graphs");
    }

    const result = computePortfolioGraphData_(ss, symbol);
    Logger.log("getPortfolioGraphData: Computed data, prices count = " + (result.prices ? result.prices.length : 0));
    return result;
  } catch (e) {
    Logger.log("getPortfolioGraphData ERROR: " + e.message + "\n" + e.stack);
    throw e;
  }
}

/**
 * Computes portfolio graph data for a symbol.
 */
function computePortfolioGraphData_(ss, symbol) {
  // Parse positions
  let shares = [];
  let bullCallSpreads = [];
  let bullPutSpreads = [];
  let bearCallSpreads = [];
  let longCalls = [];
  let shortCalls = [];
  let longPuts = [];
  let shortPuts = [];

  let cash = 0;
  let customPositions = [];

  const legsRange = getNamedRangeWithTableFallback_(ss, "Portfolio");
  if (legsRange) {
    const legsRows = legsRange.getValues();
    const parsed = parsePositionsForSymbol_(legsRows, symbol);
    shares = parsed.shares;
    bullCallSpreads = parsed.bullCallSpreads;
    bullPutSpreads = parsed.bullPutSpreads;
    bearCallSpreads = parsed.bearCallSpreads;
    longCalls = parsed.longCalls || [];
    shortCalls = parsed.shortCalls || [];
    longPuts = parsed.longPuts || [];
    shortPuts = parsed.shortPuts || [];
    customPositions = parsed.customPositions || [];
    cash = parsed.cash || 0;
  }

  const allSpreads = [...bullCallSpreads, ...bullPutSpreads, ...bearCallSpreads];
  const allSingleLegs = [...longCalls, ...shortCalls, ...longPuts, ...shortPuts];

  // Flatten custom position legs for quote fetching and DTE calculation
  const allCustomLegs = customPositions.flatMap(cp => cp.legs);

  // Calculate DTE (days to expiration) for each spread and single leg
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  for (const sp of allSpreads) {
    if (sp.expiration) {
      let expDate = sp.expiration;
      if (!(expDate instanceof Date)) {
        expDate = new Date(sp.expiration);
      }
      if (!isNaN(expDate.getTime())) {
        expDate.setHours(0, 0, 0, 0);
        sp.dte = Math.max(1, Math.round((expDate - today) / (1000 * 60 * 60 * 24)));
      }
    }
    if (!sp.dte) sp.dte = 365; // Default fallback
    Logger.log(`Spread ${sp.label}: expiration=${sp.expiration}, dte=${sp.dte}`);
  }
  for (const leg of allSingleLegs) {
    if (leg.expiration) {
      let expDate = leg.expiration;
      if (!(expDate instanceof Date)) {
        expDate = new Date(leg.expiration);
      }
      if (!isNaN(expDate.getTime())) {
        expDate.setHours(0, 0, 0, 0);
        leg.dte = Math.max(1, Math.round((expDate - today) / (1000 * 60 * 60 * 24)));
      }
    }
    if (!leg.dte) leg.dte = 365;
    Logger.log(`Single leg ${leg.label}: expiration=${leg.expiration}, dte=${leg.dte}`);
  }
  // Calculate DTE for custom position legs
  for (const leg of allCustomLegs) {
    if (leg.expiration) {
      let expDate = leg.expiration;
      if (!(expDate instanceof Date)) {
        expDate = new Date(leg.expiration);
      }
      if (!isNaN(expDate.getTime())) {
        expDate.setHours(0, 0, 0, 0);
        leg.dte = Math.max(1, Math.round((expDate - today) / (1000 * 60 * 60 * 24)));
      }
    }
    if (!leg.dte) leg.dte = 365;
  }

  // Pre-fetch actual option prices for all spreads (for "current" mode)
  // Each spread gets { longBid, longMid, longAsk, shortBid, shortMid, shortAsk }
  const spreadQuotes = allSpreads.map(sp => {
    const quotes = fetchSpreadQuotes_(symbol, sp.expiration, sp.kLong, sp.kShort, sp.flavor === "CALL" ? "Call" : "Put");
    // Log for debugging
    const spreadValue = quotes.longMid != null && quotes.shortMid != null ? quotes.longMid - quotes.shortMid : null;
    Logger.log(`Spread ${sp.label}: longMid=${quotes.longMid}, shortMid=${quotes.shortMid}, spreadValue=${spreadValue}, debit=${sp.debit}`);
    return quotes;
  });

  // Pre-fetch quotes for single-leg options
  const singleLegQuotes = allSingleLegs.map(leg => {
    const quote = getOptionQuote_(symbol, leg.expiration, leg.strike, leg.type);
    Logger.log(`Single leg ${leg.label}: mid=${quote?.mid}, price=${leg.price}`);
    return quote;
  });

  // Pre-fetch quotes for custom position legs
  const customLegQuotes = allCustomLegs.map(leg => {
    const quote = getOptionQuote_(symbol, leg.expiration, leg.strike, leg.type);
    return quote;
  });
  Logger.log(`Custom positions: ${customPositions.length}, total legs: ${allCustomLegs.length}`);

  // Compute smart price range
  const smart = computeSmartDefaults_(ss, symbol);
  const minPrice = smart.minPrice;
  const maxPrice = smart.maxPrice;
  const step = smart.step;

  // Compute denominators for ROI
  const sharesCost = shares.reduce((sum, sh) => sum + sh.qty * sh.basis, 0);
  const totalShares = shares.reduce((sum, sh) => sum + sh.qty, 0);

  // Summary logging for debugging
  Logger.log(`=== Portfolio Summary for ${symbol} ===`);
  Logger.log(`Shares: ${totalShares} shares, cost basis = $${sharesCost.toFixed(2)}`);
  Logger.log(`Bull Call Spreads: ${bullCallSpreads.length} positions`);
  Logger.log(`Bull Put Spreads: ${bullPutSpreads.length} positions`);
  Logger.log(`Bear Call Spreads: ${bearCallSpreads.length} positions`);
  Logger.log(`Long Calls: ${longCalls.length} positions`);
  Logger.log(`Short Calls: ${shortCalls.length} positions`);
  Logger.log(`Long Puts: ${longPuts.length} positions`);
  Logger.log(`Short Puts: ${shortPuts.length} positions`);
  Logger.log(`Cash: $${cash.toFixed(2)}`);

  const spreadInvestments = allSpreads.map((sp) => {
    if (sp.flavor === "CALL") {
      return sp.debit * 100 * sp.qty;
    } else {
      const width = sp.kShort - sp.kLong;
      const credit = -sp.debit;
      return (width - credit) * 100 * sp.qty;
    }
  });
  const totalSpreadInvestment = spreadInvestments.reduce((sum, v) => sum + v, 0);

  // Single-leg investments: long = cost, short = premium received (negative cost)
  const singleLegInvestments = allSingleLegs.map(leg => {
    return leg.isLong ? leg.price * 100 * leg.qty : -leg.price * 100 * leg.qty;
  });

  // Custom position investments: sum of all leg investments
  const customPositionInvestments = customPositions.map(cp => {
    return cp.legs.reduce((sum, leg) => {
      return sum + (leg.isLong ? leg.price * 100 * leg.qty : -leg.price * 100 * leg.qty);
    }, 0);
  });
  const totalSingleLegInvestment = singleLegInvestments.reduce((sum, v) => sum + v, 0);

  // Build strategy groups (spreads, single legs, and custom positions)
  const strategyGroupDefs = [
    { name: "Bull Call Spreads", spreads: bullCallSpreads, flavor: "CALL", isSingleLeg: false },
    { name: "Bull Put Spreads", spreads: bullPutSpreads, flavor: "PUT", isSingleLeg: false },
    { name: "Bear Call Spreads", spreads: bearCallSpreads, flavor: "BEAR_CALL", isSingleLeg: false },
    { name: "Long Calls", spreads: longCalls, flavor: "LONG_CALL", isSingleLeg: true },
    { name: "Short Calls", spreads: shortCalls, flavor: "SHORT_CALL", isSingleLeg: true },
    { name: "Long Puts", spreads: longPuts, flavor: "LONG_PUT", isSingleLeg: true },
    { name: "Short Puts", spreads: shortPuts, flavor: "SHORT_PUT", isSingleLeg: true },
  ];
  // Add each custom position as its own strategy group
  for (let i = 0; i < customPositions.length; i++) {
    strategyGroupDefs.push({
      name: customPositions[i].label,
      spreads: [customPositions[i]],
      isCustom: true,
      customIndex: i
    });
  }
  const strategyGroups = strategyGroupDefs.filter(g => g.spreads.length > 0);
  const strategyLabels = strategyGroups.map(g => g.name);

  // Compute strategy investments
  let spreadIdx = 0;
  let singleLegIdx = 0;
  const strategyInvestments = strategyGroups.map(g => {
    let sum = 0;
    if (g.isCustom) {
      sum = customPositionInvestments[g.customIndex];
    } else if (g.isSingleLeg) {
      for (let i = 0; i < g.spreads.length; i++) {
        sum += singleLegInvestments[singleLegIdx++];
      }
    } else {
      for (let i = 0; i < g.spreads.length; i++) {
        sum += spreadInvestments[spreadIdx++];
      }
    }
    return sum;
  });

  // Get current price via GOOGLEFINANCE
  let currentPrice = null;
  try {
    const tempSheet = ss.insertSheet("__temp_price__");
    tempSheet.getRange("A1").setFormula(`=GOOGLEFINANCE("${symbol}")`);
    SpreadsheetApp.flush();

    // Try up to 3 times with increasing wait, as GOOGLEFINANCE can be slow
    for (let attempt = 0; attempt < 3; attempt++) {
      Utilities.sleep(500 + attempt * 500); // 500ms, 1000ms, 1500ms
      const val = tempSheet.getRange("A1").getValue();
      if (typeof val === "number" && isFinite(val) && val > 0) {
        currentPrice = val;
        break;
      }
    }

    ss.deleteSheet(tempSheet);
    Logger.log("GOOGLEFINANCE returned currentPrice: " + currentPrice + " for " + symbol);
  } catch (e) {
    Logger.log("Could not get current price: " + e);
    // Clean up temp sheet if it exists
    try {
      const tempSheet = ss.getSheetByName("__temp_price__");
      if (tempSheet) ss.deleteSheet(tempSheet);
    } catch (e2) {}
  }

  // Build data arrays
  const prices = [];
  const sharesValues = [];
  const sharesRoi = [];
  const totalValues = [];
  const totalValuesCurrent = [];

  // Strategy values: [strategyIdx][priceIdx]
  const strategyValuesExp = strategyGroups.map(() => []);
  const strategyValuesCurrent = strategyGroups.map(() => []);
  const strategyRoisExp = strategyGroups.map(() => []);
  const strategyRoisCurrent = strategyGroups.map(() => []);

  // Individual spread and option values: [idx][priceIdx]
  // Include qty in label: "5 - Dec 28 480/490 BCS" or "3 - Dec 28 500 Long Call"
  const spreadLabels = [
    ...allSpreads.map(sp => `${sp.qty} - ${sp.label}`),
    ...allSingleLegs.map(leg => `${leg.qty} - ${leg.label} ${leg.isLong ? 'Long' : 'Short'} ${leg.type}`),
    ...customPositions.map(cp => `${cp.qty} - ${cp.label}`)
  ];
  const allPositions = allSpreads.length + allSingleLegs.length + customPositions.length;
  const spreadValuesExp = Array.from({ length: allPositions }, () => []);
  const spreadValuesCurrent = Array.from({ length: allPositions }, () => []);
  const spreadRoisExp = Array.from({ length: allPositions }, () => []);
  const spreadRoisCurrent = Array.from({ length: allPositions }, () => []);

  // Build OptionStrat URLs for spreads and single-leg options
  // Use custom URL format to include qty
  const spreadUrls = [
    ...allSpreads.map(sp => {
      try {
        const optionType = sp.flavor === "PUT" ? "Put" : "Call";
        const legs = [
          { strike: sp.kLong, type: optionType, qty: sp.qty, expiration: sp.expiration, price: sp.priceLong },
          { strike: sp.kShort, type: optionType, qty: -sp.qty, expiration: sp.expiration, price: sp.priceShort }
        ];
        return buildCustomOptionStratUrl(symbol, legs);
      } catch (e) {
        return null;
      }
    }),
    ...allSingleLegs.map(leg => {
      try {
        const legs = [{
          strike: leg.strike,
          type: leg.type,
          qty: leg.isLong ? leg.qty : -leg.qty,
          expiration: leg.expiration,
          price: leg.price
        }];
        return buildCustomOptionStratUrl(symbol, legs);
      } catch (e) {
        return null;
      }
    }),
    ...customPositions.map(cp => {
      try {
        // Convert legs to format expected by buildCustomOptionStratUrl
        // (signed qty: positive=long, negative=short)
        const legs = cp.legs.map(l => ({
          strike: l.strike,
          type: l.type,
          qty: l.isLong ? l.qty : -l.qty,
          expiration: l.expiration,
          price: l.price
        }));
        return buildCustomOptionStratUrl(cp.symbol, legs);
      } catch (e) {
        return null;
      }
    })
  ];

  for (let S = minPrice; S <= maxPrice + 1e-9; S += step) {
    prices.push(roundTo_(S, 2));

    // Shares value
    const sharesValue = shares.reduce((sum, sh) => sum + S * sh.qty, 0);
    sharesValues.push(roundTo_(sharesValue, 2));
    const sharesPL = sharesValue - sharesCost;
    sharesRoi.push(sharesCost > 0 ? roundTo_(sharesPL / sharesCost, 4) : 0);

    // Compute individual spread values (at expiration)
    const individualExp = [];
    const individualCurrent = [];

    for (let spIdx = 0; spIdx < allSpreads.length; spIdx++) {
      const sp = allSpreads[spIdx];
      const width = sp.kShort - sp.kLong;
      let valueExp = 0;
      let valueCurrent = 0;

      // Get actual current value at current stock price (from quotes fetched earlier)
      const quotes = spreadQuotes[spIdx];
      const currentSpreadValue = quotes && quotes.longMid != null && quotes.shortMid != null
        ? quotes.longMid - quotes.shortMid
        : sp.debit; // Fall back to debit paid

      if (sp.flavor === "CALL") {
        // Bull call spread (debit spread)
        // VALUE = long call intrinsic - short call intrinsic = clamped spread
        const intrinsic = clamp_(S - sp.kLong, 0, width);
        valueExp = intrinsic * 100 * sp.qty;
        // Current: estimate value at stock price S, anchored at actual current value
        valueCurrent = estimateSpreadValueAtPrice_(S, sp.kLong, sp.kShort, currentSpreadValue, sp.dte || 365, currentPrice) * 100 * sp.qty;
      } else if (sp.flavor === "PUT") {
        // Bull put spread (credit spread)
        // VALUE = long put intrinsic - short put intrinsic = -loss (negative when losing)
        // At max profit (S > kShort): value = 0
        // At max loss (S < kLong): value = -width
        const loss = clamp_(sp.kShort - S, 0, width);
        valueExp = -loss * 100 * sp.qty;
        // For current value, estimate uses recovery semantics, then convert to VALUE
        const putRecovery = quotes && quotes.longMid != null && quotes.shortMid != null
          ? width - (quotes.shortMid - quotes.longMid)
          : width + sp.debit; // debit is negative for credit spreads
        const recoveryEstimate = estimatePutSpreadValueAtPrice_(S, sp.kLong, sp.kShort, putRecovery, sp.dte || 365, currentPrice);
        // Convert from recovery (0 to width) to VALUE (-width to 0)
        valueCurrent = (recoveryEstimate - width) * 100 * sp.qty;
      } else {
        // Bear call spread (credit spread)
        // VALUE = long call intrinsic - short call intrinsic = -loss (negative when losing)
        const loss = clamp_(S - sp.kLong, 0, width);
        valueExp = -loss * 100 * sp.qty;
        valueCurrent = valueExp; // Simplified for bear call
      }

      individualExp.push(roundTo_(valueExp, 2));
      individualCurrent.push(roundTo_(valueCurrent, 2));
    }

    // Compute individual single-leg option values
    const singleLegExp = [];
    const singleLegCurrent = [];

    for (let legIdx = 0; legIdx < allSingleLegs.length; legIdx++) {
      const leg = allSingleLegs[legIdx];
      const quote = singleLegQuotes[legIdx];
      const currentMid = quote?.mid ?? leg.price;
      let valueExp = 0;
      let valueCurrent = 0;

      if (leg.type === "Call") {
        // Call option: intrinsic = max(0, S - strike) at expiration
        const intrinsic = Math.max(0, S - leg.strike);
        if (leg.isLong) {
          // Long call: value = intrinsic at expiration
          valueExp = intrinsic * 100 * leg.qty;
          valueCurrent = estimateSingleOptionValueAtPrice_(S, leg.strike, currentMid, leg.dte || 365, currentPrice, "Call") * 100 * leg.qty;
        } else {
          // Short call: value = credit - liability = premium - intrinsic
          // At max profit (OTM): value = premium (credit kept)
          // At max loss (ITM): value = premium - intrinsic (could be negative)
          valueExp = (leg.price - intrinsic) * 100 * leg.qty;
          const optionValue = estimateSingleOptionValueAtPrice_(S, leg.strike, currentMid, leg.dte || 365, currentPrice, "Call");
          valueCurrent = (leg.price - optionValue) * 100 * leg.qty;
        }
      } else {
        // Put option: intrinsic = max(0, strike - S) at expiration
        const intrinsic = Math.max(0, leg.strike - S);
        if (leg.isLong) {
          // Long put: value = intrinsic at expiration
          valueExp = intrinsic * 100 * leg.qty;
          valueCurrent = estimateSingleOptionValueAtPrice_(S, leg.strike, currentMid, leg.dte || 365, currentPrice, "Put") * 100 * leg.qty;
        } else {
          // Short put: value = credit - liability = premium - intrinsic
          // At max profit (OTM): value = premium (credit kept)
          // At max loss (ITM): value = premium - intrinsic (could be negative)
          valueExp = (leg.price - intrinsic) * 100 * leg.qty;
          const optionValue = estimateSingleOptionValueAtPrice_(S, leg.strike, currentMid, leg.dte || 365, currentPrice, "Put");
          valueCurrent = (leg.price - optionValue) * 100 * leg.qty;
        }
      }

      singleLegExp.push(roundTo_(valueExp, 2));
      singleLegCurrent.push(roundTo_(valueCurrent, 2));
    }

    // Compute custom position values (sum of all legs' P&L)
    const customPosExp = [];
    const customPosCurrent = [];
    let customLegIdx = 0;
    for (let cpIdx = 0; cpIdx < customPositions.length; cpIdx++) {
      const cp = customPositions[cpIdx];
      let sumExp = 0, sumCurrent = 0;
      for (const leg of cp.legs) {
        const quote = customLegQuotes[customLegIdx++];
        const currentMid = quote?.mid ?? leg.price;
        let legValueExp = 0, legValueCurrent = 0;

        if (leg.type === "Call") {
          const intrinsic = Math.max(0, S - leg.strike);
          if (leg.isLong) {
            legValueExp = intrinsic * 100 * leg.qty;
            legValueCurrent = estimateSingleOptionValueAtPrice_(S, leg.strike, currentMid, leg.dte || 365, currentPrice, "Call") * 100 * leg.qty;
          } else {
            // Short: value = credit - liability = premium - intrinsic
            legValueExp = (leg.price - intrinsic) * 100 * leg.qty;
            const optionValue = estimateSingleOptionValueAtPrice_(S, leg.strike, currentMid, leg.dte || 365, currentPrice, "Call");
            legValueCurrent = (leg.price - optionValue) * 100 * leg.qty;
          }
        } else {
          const intrinsic = Math.max(0, leg.strike - S);
          if (leg.isLong) {
            legValueExp = intrinsic * 100 * leg.qty;
            legValueCurrent = estimateSingleOptionValueAtPrice_(S, leg.strike, currentMid, leg.dte || 365, currentPrice, "Put") * 100 * leg.qty;
          } else {
            // Short: value = credit - liability = premium - intrinsic
            legValueExp = (leg.price - intrinsic) * 100 * leg.qty;
            const optionValue = estimateSingleOptionValueAtPrice_(S, leg.strike, currentMid, leg.dte || 365, currentPrice, "Put");
            legValueCurrent = (leg.price - optionValue) * 100 * leg.qty;
          }
        }
        sumExp += legValueExp;
        sumCurrent += legValueCurrent;
      }
      customPosExp.push(roundTo_(sumExp, 2));
      customPosCurrent.push(roundTo_(sumCurrent, 2));
    }

    // Store individual spread values
    for (let i = 0; i < allSpreads.length; i++) {
      spreadValuesExp[i].push(individualExp[i]);
      spreadValuesCurrent[i].push(individualCurrent[i]);

      const inv = spreadInvestments[i];
      const sp = allSpreads[i];
      // ROI = P&L / investment = (value - debit) / investment
      // For debit spreads: debit > 0, inv = debit, so this equals (value - inv) / inv
      // For credit spreads: debit < 0 (credit), inv = width - credit
      //   P&L = value + credit = value - debit
      const debit100 = sp.debit * 100 * sp.qty;
      spreadRoisExp[i].push(inv > 0 ? roundTo_((individualExp[i] - debit100) / inv, 4) : 0);
      spreadRoisCurrent[i].push(inv > 0 ? roundTo_((individualCurrent[i] - debit100) / inv, 4) : 0);
    }

    // Store single-leg option values (appended after spreads)
    for (let i = 0; i < allSingleLegs.length; i++) {
      const idx = allSpreads.length + i;
      spreadValuesExp[idx].push(singleLegExp[i]);
      spreadValuesCurrent[idx].push(singleLegCurrent[i]);

      const inv = singleLegInvestments[i];
      // For long positions (inv > 0): ROI = (value - cost) / cost
      // For short positions (inv < 0): value already includes credit (value = P&L = premium - intrinsic)
      //   ROI = value / |inv| = (premium - intrinsic) / premium
      //   At max profit (OTM): ROI = premium / premium = 100%
      //   At max loss: ROI = (premium - intrinsic) / premium (negative when intrinsic > premium)
      if (inv > 0) {
        spreadRoisExp[idx].push(roundTo_((singleLegExp[i] - inv) / inv, 4));
        spreadRoisCurrent[idx].push(roundTo_((singleLegCurrent[i] - inv) / inv, 4));
      } else if (inv < 0) {
        // Value already includes credit, so ROI = value / |inv|
        spreadRoisExp[idx].push(roundTo_(singleLegExp[i] / Math.abs(inv), 4));
        spreadRoisCurrent[idx].push(roundTo_(singleLegCurrent[i] / Math.abs(inv), 4));
      } else {
        spreadRoisExp[idx].push(0);
        spreadRoisCurrent[idx].push(0);
      }
    }

    // Store custom position values (appended after single legs)
    for (let i = 0; i < customPositions.length; i++) {
      const idx = allSpreads.length + allSingleLegs.length + i;
      spreadValuesExp[idx].push(customPosExp[i]);
      spreadValuesCurrent[idx].push(customPosCurrent[i]);

      const inv = customPositionInvestments[i];
      if (inv > 0) {
        spreadRoisExp[idx].push(roundTo_((customPosExp[i] - inv) / inv, 4));
        spreadRoisCurrent[idx].push(roundTo_((customPosCurrent[i] - inv) / inv, 4));
      } else if (inv < 0) {
        // Net credit position: value already includes credits, ROI = value / |inv|
        spreadRoisExp[idx].push(roundTo_(customPosExp[i] / Math.abs(inv), 4));
        spreadRoisCurrent[idx].push(roundTo_(customPosCurrent[i] / Math.abs(inv), 4));
      } else {
        spreadRoisExp[idx].push(0);
        spreadRoisCurrent[idx].push(0);
      }
    }

    // Aggregate by strategy (handles spreads, single legs, and custom positions)
    let sIdx = 0;
    let slIdx = 0;
    for (let g = 0; g < strategyGroups.length; g++) {
      let sumExp = 0, sumCurrent = 0;
      let totalDebit = 0; // Track total debit for ROI calculation (spreads only)
      const isCustomOrSingleLeg = strategyGroups[g].isCustom || strategyGroups[g].isSingleLeg;

      if (strategyGroups[g].isCustom) {
        const cpIdx = strategyGroups[g].customIndex;
        sumExp = customPosExp[cpIdx];
        sumCurrent = customPosCurrent[cpIdx];
        // Don't set totalDebit - value already includes credits for custom positions
      } else if (strategyGroups[g].isSingleLeg) {
        for (let i = 0; i < strategyGroups[g].spreads.length; i++) {
          sumExp += singleLegExp[slIdx];
          sumCurrent += singleLegCurrent[slIdx];
          // Don't add to totalDebit - value already includes credits for single legs
          slIdx++;
        }
      } else {
        // For spreads, track the actual debit (which is negative for credit spreads)
        const startIdx = sIdx;
        for (let i = 0; i < strategyGroups[g].spreads.length; i++) {
          sumExp += individualExp[sIdx];
          sumCurrent += individualCurrent[sIdx];
          sIdx++;
        }
        // Calculate total debit for this strategy group's spreads
        for (let i = 0; i < strategyGroups[g].spreads.length; i++) {
          const sp = strategyGroups[g].spreads[i];
          totalDebit += sp.debit * 100 * sp.qty;
        }
      }
      strategyValuesExp[g].push(roundTo_(sumExp, 2));
      strategyValuesCurrent[g].push(roundTo_(sumCurrent, 2));

      const inv = strategyInvestments[g];

      if (isCustomOrSingleLeg) {
        // For single legs and custom positions, value already includes credits
        // ROI = value / |inv| (long: (value - cost)/cost, short: value/credit)
        if (inv > 0) {
          // Long position: value is intrinsic, need to subtract cost
          strategyRoisExp[g].push(roundTo_((sumExp - inv) / inv, 4));
          strategyRoisCurrent[g].push(roundTo_((sumCurrent - inv) / inv, 4));
        } else if (inv < 0) {
          // Short position: value already includes credit
          strategyRoisExp[g].push(roundTo_(sumExp / Math.abs(inv), 4));
          strategyRoisCurrent[g].push(roundTo_(sumCurrent / Math.abs(inv), 4));
        } else {
          strategyRoisExp[g].push(0);
          strategyRoisCurrent[g].push(0);
        }
      } else {
        // For spreads, value is just intrinsic, need to account for debit/credit
        // ROI = (value - debit) / investment
        if (inv > 0) {
          strategyRoisExp[g].push(roundTo_((sumExp - totalDebit) / inv, 4));
          strategyRoisCurrent[g].push(roundTo_((sumCurrent - totalDebit) / inv, 4));
        } else if (inv < 0) {
          // Credit spread: P&L = value + credit = value - debit
          strategyRoisExp[g].push(roundTo_((sumExp - totalDebit) / Math.abs(inv), 4));
          strategyRoisCurrent[g].push(roundTo_((sumCurrent - totalDebit) / Math.abs(inv), 4));
        } else {
          strategyRoisExp[g].push(0);
          strategyRoisCurrent[g].push(0);
        }
      }
    }

    // Total values (excluding cash - cash doesn't change with price)
    const customTotalExp = customPosExp.reduce((sum, v) => sum + v, 0);
    const customTotalCurrent = customPosCurrent.reduce((sum, v) => sum + v, 0);
    const optionsExp = individualExp.reduce((sum, v) => sum + v, 0) + singleLegExp.reduce((sum, v) => sum + v, 0) + customTotalExp;
    const optionsCurrent = individualCurrent.reduce((sum, v) => sum + v, 0) + singleLegCurrent.reduce((sum, v) => sum + v, 0) + customTotalCurrent;
    totalValues.push(roundTo_(sharesValue + optionsExp, 2));
    totalValuesCurrent.push(roundTo_(sharesValue + optionsCurrent, 2));
  }

  return {
    symbol: symbol,
    prices: prices,
    sharesValues: sharesValues,
    sharesRoi: sharesRoi,
    totalValues: totalValues,
    totalValuesCurrent: totalValuesCurrent,
    cash: cash,
    strategyLabels: strategyLabels,
    strategyValuesExp: strategyValuesExp,
    strategyValuesCurrent: strategyValuesCurrent,
    strategyRoisExp: strategyRoisExp,
    strategyRoisCurrent: strategyRoisCurrent,
    spreadLabels: spreadLabels,
    spreadUrls: spreadUrls,
    spreadValuesExp: spreadValuesExp,
    spreadValuesCurrent: spreadValuesCurrent,
    spreadRoisExp: spreadRoisExp,
    spreadRoisCurrent: spreadRoisCurrent,
    totalShares: totalShares,
    sharesCost: sharesCost,
    spreadCount: allSpreads.length,
    spreadInvestment: totalSpreadInvestment,
    currentPrice: currentPrice,
  };
}

/**
 * Fetches actual option quotes for a spread from OptionPricesUploaded.
 * Returns { longBid, longMid, longAsk, shortBid, shortMid, shortAsk } or null values if not found.
 */
function fetchSpreadQuotes_(symbol, expiration, kLong, kShort, optionType) {
  const result = {
    longBid: null, longMid: null, longAsk: null,
    shortBid: null, shortMid: null, shortAsk: null
  };

  try {
    // Format expiration for lookup
    const expStr = formatExpirationForLookup_(expiration);

    // Look up long leg
    const longRes = XLookupByKeys(
      [symbol, expStr, kLong, optionType],
      ["symbol", "expiration", "strike", "type"],
      ["bid", "mid", "ask"],
      "OptionPricesUploaded"
    );
    if (longRes && longRes[0]) {
      result.longBid = parseFloat(longRes[0][0]) || null;
      result.longMid = parseFloat(longRes[0][1]) || null;
      result.longAsk = parseFloat(longRes[0][2]) || null;
    }

    // Look up short leg
    const shortRes = XLookupByKeys(
      [symbol, expStr, kShort, optionType],
      ["symbol", "expiration", "strike", "type"],
      ["bid", "mid", "ask"],
      "OptionPricesUploaded"
    );
    if (shortRes && shortRes[0]) {
      result.shortBid = parseFloat(shortRes[0][0]) || null;
      result.shortMid = parseFloat(shortRes[0][1]) || null;
      result.shortAsk = parseFloat(shortRes[0][2]) || null;
    }
  } catch (e) {
    Logger.log("fetchSpreadQuotes_ error: " + e.message);
  }

  return result;
}

/**
 * Formats expiration for XLookupByKeys lookup (M/D/YYYY format).
 */
function formatExpirationForLookup_(exp) {
  if (!exp) return "";
  if (exp instanceof Date) {
    return `${exp.getMonth() + 1}/${exp.getDate()}/${exp.getFullYear()}`;
  }
  const s = String(exp).trim();
  // If already M/D/YYYY format, return as is
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) return s;
  // Convert yyyy-MM-dd (ISO) to M/D/YYYY
  const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    const [, y, m, d] = isoMatch;
    return `${parseInt(m, 10)}/${parseInt(d, 10)}/${y}`;
  }
  return s;
}


/**
 * Estimates bull call spread value at a given stock price S using a
 * calibrated Black-Scholes model.
 *
 * @param {number} S - Stock price to evaluate
 * @param {number} kLong - Long call strike
 * @param {number} kShort - Short call strike
 * @param {number} currentValue - Current spread value (per share) at current stock price
 * @param {number} dte - Days to expiration
 * @param {number} currentStockPrice - Actual current stock price
 * @returns {number} Estimated spread value per share at price S
 */
function estimateSpreadValueAtPrice_(S, kLong, kShort, currentValue, dte, currentStockPrice) {
    const r = 0.04; // Assumed risk-free rate
    const t = Math.max(dte, 1) / 365;

    // Standard Abramowitz & Stegun 5-term CND approximation
    function cnd(x) {
        const neg = (x < 0);
        const z = Math.abs(x);
        const k = 1.0 / (1.0 + 0.2316419 * z);
        const pdf = Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);
        let v = 1.0 - pdf * (0.319381530 * k - 0.356563782 * Math.pow(k, 2) + 1.781477937 * Math.pow(k, 3) - 1.821255978 * Math.pow(k, 4) + 1.330274429 * Math.pow(k, 5));
        return neg ? 1.0 - v : v;
    }

    function bsCall(stock, strike, time, rate, sigma) {
        if (time <= 0) return Math.max(0, stock - strike);
        if (sigma <= 0.0001) return Math.max(0, stock - strike * Math.exp(-rate * time));
        const d1 = (Math.log(stock / strike) + (rate + (sigma * sigma) / 2) * time) / (sigma * Math.sqrt(time));
        const d2 = d1 - sigma * Math.sqrt(time);
        return stock * cnd(d1) - strike * Math.exp(-rate * time) * cnd(d2);
    }

    const getSpreadPrice = (price, sigma) => bsCall(price, kLong, t, r, sigma) - bsCall(price, kShort, t, r, sigma);

    // Solve for Implied Volatility (sigma) using reference price/value
    let sigma = 0.5;
    for (let i = 0; i < 40; i++) {
        let p = getSpreadPrice(currentStockPrice, sigma);
        let diff = p - currentValue;
        if (Math.abs(diff) < 0.00001) break;

        let vega = (getSpreadPrice(currentStockPrice, sigma + 0.01) - p) / 0.01;
        sigma -= diff / (vega || 0.01);
        if (sigma <= 0) sigma = 0.001;
    }

    return getSpreadPrice(S, sigma);
}

/**
 * Data-driven test runner for the spread estimation logic.
 */
function testEstimateSpreadValueAtPrice() {
    const refS = 424.00;
    const refP = 23.60;
    const minK = 500;
    const maxK = 600;
    const days = 1040;

    const testTable = [
        { label: "Calibration",   s: 424.00,  expected: 23.60 },
        { label: "Moderate Rise",  s: 500.00,  expected: 28.87 },
        { label: "Strike Breach",  s: 600.00,  expected: 35.19 },
        { label: "Deep ITM",       s: 1200.00, expected: 60.23 },
        { label: "Extreme ITM",    s: 2000.00, expected: 74.76 },
        { label: "Downside",       s: 300.00,  expected: 14.38 }
    ];

    console.log("Scenario".padEnd(18) + " | " + "Target S".padEnd(10) + " | " + "Actual".padEnd(10) + " | " + "Expected".padEnd(10) + " | " + "Result");
    console.log("-".repeat(70));

    testTable.forEach(row => {
        // Parameter order: S, kLong, kShort, currentValue, dte, currentStockPrice
        const actual = estimateSpreadValueAtPrice_(row.s, minK, maxK, refP, days, refS);
        const diff = Math.abs(actual - row.expected);
        const passed = diff < 0.15;

        console.log(
            row.label.padEnd(18) + " | " +
            row.s.toFixed(2).padEnd(10) + " | " +
            actual.toFixed(2).padEnd(10) + " | " +
            row.expected.toFixed(2).padEnd(10) + " | " +
            (passed ? "PASS ✅" : "FAIL ❌")
        );
    });
}


/**
 * Estimates bull put spread value at a given stock price S.
 * Bull put spread (credit spread): Sell higher strike put, buy lower strike put.
 * Profits when stock stays above short strike.
 *
 * @param {number} S - Stock price to evaluate
 * @param {number} kLong - Long put strike (lower)
 * @param {number} kShort - Short put strike (higher)
 * @param {number} currentValue - Current spread value at current stock price
 * @param {number} dte - Days to expiration
 * @param {number} currentStockPrice - Actual current stock price
 * @returns {number} Estimated spread value per share at price S
 */
function estimatePutSpreadValueAtPrice_(S, kLong, kShort, currentValue, dte, currentStockPrice) {
  const width = kShort - kLong;

  // At stock price 0, loss is max (value = 0)
  if (S <= 0) return 0;

  // Intrinsic value: what we keep of the width
  const loss = clamp_(kShort - S, 0, width);
  const intrinsic = width - loss;

  // If no valid current data, just return intrinsic
  if (!currentValue || currentValue <= 0 || !currentStockPrice || currentStockPrice <= 0) {
    return intrinsic;
  }

  // Above short strike: full profit (value = width)
  if (S >= kShort) return width;

  // At current stock price: return actual current value (anchor point)
  if (Math.abs(S - currentStockPrice) < 1) {
    return currentValue;
  }

  // Compute time value from actual market prices
  const currentLoss = clamp_(kShort - currentStockPrice, 0, width);
  const currentIntrinsic = width - currentLoss;
  const timeValue = Math.max(0, currentValue - currentIntrinsic);

  if (S <= kLong) {
    // Below long strike: max loss on intrinsic, but some time value remains
    const otmFactor = S / kLong;
    return timeValue * otmFactor;
  } else {
    // Between strikes: intrinsic + decaying time value
    const progress = (S - kLong) / width;
    const adjustedTimeValue = timeValue * (1 - progress);
    return intrinsic + adjustedTimeValue;
  }
}

/**
 * Estimates single option value at a given stock price S.
 * Uses Black-Scholes with implied volatility backed out from current market price.
 *
 * @param {number} S - Stock price to evaluate
 * @param {number} strike - Option strike price
 * @param {number} currentValue - Current option value (mid price)
 * @param {number} dte - Days to expiration
 * @param {number} currentStockPrice - Actual current stock price
 * @param {string} optionType - "Call" or "Put"
 * @returns {number} Estimated option value per share at price S
 */
function estimateSingleOptionValueAtPrice_(S, strike, currentValue, dte, currentStockPrice, optionType) {
  const r = 0.04; // Assumed risk-free rate
  const t = Math.max(dte, 1) / 365;

  // Standard CND approximation
  function cnd(x) {
    const neg = (x < 0);
    const z = Math.abs(x);
    const k = 1.0 / (1.0 + 0.2316419 * z);
    const pdf = Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);
    let v = 1.0 - pdf * (0.319381530 * k - 0.356563782 * Math.pow(k, 2) + 1.781477937 * Math.pow(k, 3) - 1.821255978 * Math.pow(k, 4) + 1.330274429 * Math.pow(k, 5));
    return neg ? 1.0 - v : v;
  }

  function bsCall(stock, k, time, rate, sigma) {
    if (time <= 0) return Math.max(0, stock - k);
    if (sigma <= 0.0001) return Math.max(0, stock - k * Math.exp(-rate * time));
    const d1 = (Math.log(stock / k) + (rate + (sigma * sigma) / 2) * time) / (sigma * Math.sqrt(time));
    const d2 = d1 - sigma * Math.sqrt(time);
    return stock * cnd(d1) - k * Math.exp(-rate * time) * cnd(d2);
  }

  function bsPut(stock, k, time, rate, sigma) {
    if (time <= 0) return Math.max(0, k - stock);
    if (sigma <= 0.0001) return Math.max(0, k * Math.exp(-rate * time) - stock);
    const d1 = (Math.log(stock / k) + (rate + (sigma * sigma) / 2) * time) / (sigma * Math.sqrt(time));
    const d2 = d1 - sigma * Math.sqrt(time);
    return k * Math.exp(-rate * time) * cnd(-d2) - stock * cnd(-d1);
  }

  const getOptionPrice = optionType === "Call"
    ? (price, sigma) => bsCall(price, strike, t, r, sigma)
    : (price, sigma) => bsPut(price, strike, t, r, sigma);

  // If no valid current data, use intrinsic value
  if (!currentValue || currentValue <= 0 || !currentStockPrice || currentStockPrice <= 0) {
    return optionType === "Call" ? Math.max(0, S - strike) : Math.max(0, strike - S);
  }

  // Solve for Implied Volatility (sigma) using reference price/value
  let sigma = 0.5;
  for (let i = 0; i < 40; i++) {
    let p = getOptionPrice(currentStockPrice, sigma);
    let diff = p - currentValue;
    if (Math.abs(diff) < 0.00001) break;

    let vega = (getOptionPrice(currentStockPrice, sigma + 0.01) - p) / 0.01;
    sigma -= diff / (vega || 0.01);
    if (sigma <= 0) sigma = 0.001;
    if (sigma > 5) sigma = 5; // Cap volatility
  }

  return getOptionPrice(S, sigma);
}


/* =========================================================
   Get unique symbols from position tables
   ========================================================= */

function getUniqueSymbolsFromPositions_(ss) {
  const legsRange = getNamedRangeWithTableFallback_(ss, "Portfolio");
  if (!legsRange) return [];

  const legsRows = legsRange.getValues();
  return getSymbolsFromLegsTable_(legsRows);
}

/* =========================================================
   Smart price-range defaults based on symbol's positions
   ========================================================= */

/**
 * Computes sensible min/max/step defaults from Legs table data.
 * Falls back to 350/900/5 when no position data is available.
 */
function computeSmartDefaults_(ss, symbol) {
  const fallback = { minPrice: 350, maxPrice: 900, step: 5 };

  const legsRange = getNamedRangeWithTableFallback_(ss, "Portfolio");
  if (!legsRange) return fallback;

  const rows = legsRange.getValues();
  if (!rows || rows.length < 2) return fallback;

  const headers = rows[0];
  const idxSym = findColumn_(headers, ["symbol", "ticker"]);
  const idxStrike = findColumn_(headers, ["strike", "strikeprice"]);
  const idxPrice = findColumn_(headers, ["price", "cost", "entry", "premium", "basis", "costbasis", "avgprice", "pricepaid"]);
  const idxType = findColumn_(headers, ["type", "optiontype", "callput", "cp", "putcall", "legtype"]);

  if (idxSym < 0 || idxPrice < 0) return fallback;

  // Collect reference prices: stock basis and option strikes
  const refPrices = [];
  let lastSym = "";

  for (let r = 1; r < rows.length; r++) {
    const rawSym = String(rows[r][idxSym] ?? "").trim().toUpperCase();
    if (rawSym) lastSym = rawSym;
    if (lastSym !== symbol) continue;

    const strike = idxStrike >= 0 ? parseNumber_(rows[r][idxStrike]) : NaN;
    const price = parseNumber_(rows[r][idxPrice]);
    const type = idxType >= 0 ? parseOptionType_(rows[r][idxType]) : null;
    const isStock = type === null && !Number.isFinite(strike);

    if (isStock && Number.isFinite(price) && price > 0) {
      // Stock basis — primary reference
      refPrices.push(price);
    }
    if (Number.isFinite(strike) && strike > 0) {
      refPrices.push(strike);
    }
  }

  if (refPrices.length === 0) return fallback;

  const minRef = Math.min(...refPrices);
  const maxRef = Math.max(...refPrices);

  // Range: 20% below min to 200% of max, giving room to see outcomes
  let minPrice = Math.floor(minRef * 0.2);
  let maxPrice = Math.ceil(maxRef * 2.0);

  // Pick step based on range width
  const range = maxPrice - minPrice;
  let step;
  if (range < 50) step = 1;
  else if (range < 200) step = 2;
  else step = 5;

  // Round min/max to nice multiples of step
  minPrice = Math.floor(minPrice / step) * step;
  maxPrice = Math.ceil(maxPrice / step) * step;

  // Ensure we have at least some range
  if (minPrice >= maxPrice) {
    minPrice = Math.floor(minRef * 0.5 / step) * step;
    maxPrice = Math.ceil(maxRef * 1.5 / step) * step;
  }

  return { minPrice, maxPrice, step };
}

/* =========================================================
   Named range resolution
   ========================================================= */

function getNamedRangeWithTableFallback_(ss, rangeNameRaw) {
  const name = String(rangeNameRaw || "").trim();
  if (!name) return null;

  let r = ss.getRangeByName(name);
  if (r) return r;

  r = ss.getRangeByName(name + "Table");
  if (r) return r;

  return null;
}




