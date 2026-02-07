/**
 * PlotPortfolioValueByPrice.js
 * ------------------------------------------------------------
 * PlotPortfolioValueByPrice
 *
 * Reads positions from named ranges on the Positions sheet, extracts unique symbols,
 * then creates/refreshes a "<SYMBOL>PortfolioValueByPrice" tab for EACH symbol.
 *
 * Position tables (on Positions sheet):
 *   - Stocks or StocksTable
 *   - BullCallSpreads or BullCallSpreadsTable
 *   - IronCondors or IronCondorsTable
 *
 * For each symbol tab:
 *   - Config table at K1:L10 (column K = 11) with labels/descriptions
 *   - Config columns are NEVER hidden
 *   - Default tableStartRow = 85
 *   - Generated data table starts at row tableStartRow (headers) / tableStartRow+1 (data)
 *   - FOUR charts:
 *       Chart 1: Portfolio Value vs Price (Shares $, Options $, Total $)
 *       Chart 2: % Return vs Price (Shares %, Options %) (ROI style, not contribution)
 *       Chart 3: Individual Spreads Value by Price (one series per spread, labeled like "Dec 28 350/450")
 *       Chart 4: Individual Spreads ROI by Price (% return for each spread)
 *   - Chart series labels come from the first row of the data table (headers)
 *
 * Named range resolution:
 *   - RangeName must be a Named Range; if not found, script tries RangeName + "Table"
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

  // Show the portfolio graphs modal
  const html = HtmlService.createHtmlOutputFromFile("PortfolioGraphs")
    .setWidth(1200)
    .setHeight(900);
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

  const legsRange = getNamedRangeWithTableFallback_(ss, "Portfolio");
  if (legsRange) {
    const legsRows = legsRange.getValues();
    const parsed = parsePositionsForSymbol_(legsRows, symbol);
    shares = parsed.shares;
    bullCallSpreads = parsed.bullCallSpreads;
    bullPutSpreads = parsed.bullPutSpreads;
    bearCallSpreads = parsed.bearCallSpreads;
  }

  const allSpreads = [...bullCallSpreads, ...bullPutSpreads, ...bearCallSpreads];

  // Compute smart price range
  const smart = computeSmartDefaults_(ss, symbol);
  const minPrice = smart.minPrice;
  const maxPrice = smart.maxPrice;
  const step = smart.step;

  // Compute denominators for ROI
  const sharesCost = shares.reduce((sum, sh) => sum + sh.qty * sh.basis, 0);
  const totalShares = shares.reduce((sum, sh) => sum + sh.qty, 0);

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

  // Build strategy groups
  const strategyGroupDefs = [
    { name: "Bull Call Spreads", spreads: bullCallSpreads, flavor: "CALL" },
    { name: "Bull Put Spreads", spreads: bullPutSpreads, flavor: "PUT" },
    { name: "Bear Call Spreads", spreads: bearCallSpreads, flavor: "BEAR_CALL" },
  ];
  const strategyGroups = strategyGroupDefs.filter(g => g.spreads.length > 0);
  const strategyLabels = strategyGroups.map(g => g.name);

  // Compute strategy investments
  let spreadIdx = 0;
  const strategyInvestments = strategyGroups.map(g => {
    let sum = 0;
    for (let i = 0; i < g.spreads.length; i++) {
      sum += spreadInvestments[spreadIdx++];
    }
    return sum;
  });

  // Get current price via GOOGLEFINANCE
  let currentPrice = null;
  try {
    const tempSheet = ss.insertSheet("__temp_price__");
    tempSheet.getRange("A1").setFormula(`=GOOGLEFINANCE("${symbol}")`);
    SpreadsheetApp.flush();
    Utilities.sleep(500); // Give GOOGLEFINANCE time to evaluate
    currentPrice = tempSheet.getRange("A1").getValue();
    ss.deleteSheet(tempSheet);
    if (typeof currentPrice !== "number" || !isFinite(currentPrice)) {
      currentPrice = null;
    }
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

  // Individual spread values: [spreadIdx][priceIdx]
  const spreadLabels = allSpreads.map(sp => sp.label);
  const spreadValuesExp = allSpreads.map(() => []);
  const spreadValuesCurrent = allSpreads.map(() => []);
  const spreadRoisExp = allSpreads.map(() => []);
  const spreadRoisCurrent = allSpreads.map(() => []);

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

    for (const sp of allSpreads) {
      const width = sp.kShort - sp.kLong;
      let valueExp = 0;
      let valueCurrent = 0;

      if (sp.flavor === "CALL") {
        // Bull call spread
        const intrinsic = clamp_(S - sp.kLong, 0, width);
        valueExp = intrinsic * 100 * sp.qty;
        // Current: add time value estimate (simplified: linear decay based on moneyness)
        const timeValue = estimateTimeValue_(S, sp.kLong, sp.kShort, sp.dte || 365, width);
        valueCurrent = (intrinsic + timeValue) * 100 * sp.qty;
      } else if (sp.flavor === "PUT") {
        // Bull put spread
        const loss = clamp_(sp.kShort - S, 0, width);
        valueExp = (width - loss) * 100 * sp.qty;
        valueCurrent = valueExp; // Simplified for puts
      } else {
        // Bear call spread
        const loss = clamp_(S - sp.kLong, 0, width);
        valueExp = (width - loss) * 100 * sp.qty;
        valueCurrent = valueExp; // Simplified
      }

      individualExp.push(roundTo_(valueExp, 2));
      individualCurrent.push(roundTo_(valueCurrent, 2));
    }

    // Store individual spread values
    for (let i = 0; i < allSpreads.length; i++) {
      spreadValuesExp[i].push(individualExp[i]);
      spreadValuesCurrent[i].push(individualCurrent[i]);

      const inv = spreadInvestments[i];
      spreadRoisExp[i].push(inv > 0 ? roundTo_((individualExp[i] - inv) / inv, 4) : 0);
      spreadRoisCurrent[i].push(inv > 0 ? roundTo_((individualCurrent[i] - inv) / inv, 4) : 0);
    }

    // Aggregate by strategy
    let sIdx = 0;
    for (let g = 0; g < strategyGroups.length; g++) {
      let sumExp = 0, sumCurrent = 0;
      for (let i = 0; i < strategyGroups[g].spreads.length; i++) {
        sumExp += individualExp[sIdx];
        sumCurrent += individualCurrent[sIdx];
        sIdx++;
      }
      strategyValuesExp[g].push(roundTo_(sumExp, 2));
      strategyValuesCurrent[g].push(roundTo_(sumCurrent, 2));

      const inv = strategyInvestments[g];
      strategyRoisExp[g].push(inv > 0 ? roundTo_((sumExp - inv) / inv, 4) : 0);
      strategyRoisCurrent[g].push(inv > 0 ? roundTo_((sumCurrent - inv) / inv, 4) : 0);
    }

    // Total values
    const optionsExp = individualExp.reduce((sum, v) => sum + v, 0);
    const optionsCurrent = individualCurrent.reduce((sum, v) => sum + v, 0);
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
    strategyLabels: strategyLabels,
    strategyValuesExp: strategyValuesExp,
    strategyValuesCurrent: strategyValuesCurrent,
    strategyRoisExp: strategyRoisExp,
    strategyRoisCurrent: strategyRoisCurrent,
    spreadLabels: spreadLabels,
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
 * Estimates time value for a call spread (simplified model).
 * Returns per-contract time value estimate.
 */
function estimateTimeValue_(S, kLong, kShort, dte, width) {
  // Simplified time value: decays linearly with time, higher when near the money
  const midStrike = (kLong + kShort) / 2;
  const moneyness = S / midStrike;

  // Time factor: sqrt(dte/365) gives time decay curve
  const timeFactor = Math.sqrt(Math.max(dte, 1) / 365);

  // ATM gets max time value, decreases as you move away
  const atmFactor = Math.exp(-Math.pow(moneyness - 1, 2) * 4);

  // Max time value is about 20% of width for ATM with 1 year to expiry
  const maxTimeValue = width * 0.2;

  return maxTimeValue * timeFactor * atmFactor;
}

/* =========================================================
   onEdit trigger - rebuild when config edited
   ========================================================= */

function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const ss = e.range.getSheet().getParent();
    const named = ss.getNamedRanges().filter(nr => nr.getName().startsWith("Config_"));
    if (named.length === 0) return;

    for (const nr of named) {
      const cfgRange = nr.getRange();
      if (rangesIntersect_(e.range, cfgRange)) {
        const symbol = nr.getName().slice("Config_".length);
        plotForSymbol_(ss, symbol);
        return;
      }
    }
  } catch (err) {
    // silent
  }
}

/* =========================================================
   Get unique symbols from position tables
   ========================================================= */

function getUniqueSymbolsFromPositions_(ss) {
  // Try Legs table first
  const legsRange = getNamedRangeWithTableFallback_(ss, "Portfolio");
  if (legsRange) {
    const legsRows = legsRange.getValues();
    const legsSymbols = getSymbolsFromLegsTable_(legsRows);
    if (legsSymbols.length > 0) return legsSymbols;
  }

  // Fall back to old 3-table logic
  const symbols = new Set();

  const tableNames = ["Stocks", "BullCallSpreads", "IronCondors"];

  for (const tableName of tableNames) {
    const range = getNamedRangeWithTableFallback_(ss, tableName);
    if (!range) continue;

    const rows = range.getValues();
    if (rows.length < 2) continue;

    const headerNorm = rows[0].map(normKey_);
    const idxSym = findColumn_(headerNorm, ["symbol", "ticker"]);
    if (idxSym < 0) continue;

    // Check for Status column (used by IronCondors to mark closed positions)
    const idxStatus = findColumn_(headerNorm, ["status"]);

    for (let r = 1; r < rows.length; r++) {
      // Skip closed positions
      if (idxStatus >= 0) {
        const status = String(rows[r][idxStatus] ?? "").trim().toLowerCase();
        if (status === "closed") continue;
      }

      const sym = String(rows[r][idxSym] ?? "").trim().toUpperCase();
      if (sym) symbols.add(sym);
    }
  }

  return Array.from(symbols).sort();
}

/* =========================================================
   Main per-symbol logic
   ========================================================= */

function plotForSymbol_(ss, symbolRaw) {
  const symbol = String(symbolRaw || "").trim().toUpperCase();
  if (!symbol) return;

  const sheetName = `${symbol}PortfolioValueByPrice`;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.showSheet();

  const cfg = ensureAndReadConfig_(ss, sheet, symbol);

  // ── Parse positions ────────────────────────────────────────────────
  let shares = [];
  let bullCallSpreads = [];
  let bullPutSpreads = [];
  let bearCallSpreads = [];

  const legsRange = getNamedRangeWithTableFallback_(ss, "Portfolio");
  if (legsRange) {
    // Unified Legs table path
    const legsRows = legsRange.getValues();
    const parsed = parsePositionsForSymbol_(legsRows, symbol);
    shares = parsed.shares;
    bullCallSpreads = parsed.bullCallSpreads;
    bullPutSpreads = parsed.bullPutSpreads;
    bearCallSpreads = parsed.bearCallSpreads;

    clearStatus_(sheet);
  } else {
    // Fall back to old 3-table logic
    const stockRanges = [];
    const callSpreadRanges = [];
    const putSpreadRanges = [];

    const stocksRange = getNamedRangeWithTableFallback_(ss, "Stocks");
    if (stocksRange) stockRanges.push(stocksRange);

    const bcsRange = getNamedRangeWithTableFallback_(ss, "BullCallSpreads");
    if (bcsRange) callSpreadRanges.push(bcsRange);

    const icRange = getNamedRangeWithTableFallback_(ss, "IronCondors");

    if (stockRanges.length === 0 && callSpreadRanges.length === 0 && !icRange) {
      writeStatus_(sheet, `No position tables found. Create named ranges:\n  - Portfolio or PortfolioTable\n  - Stocks or StocksTable\n  - BullCallSpreads or BullCallSpreadsTable\n  - IronCondors or IronCondorsTable`);
      return;
    }

    clearStatus_(sheet);

    const meta = { currentPrice: null };
    for (const rng of stockRanges) {
      shares.push(...parseSharesFromTableForSymbol_(rng.getValues(), symbol, meta));
    }
    for (const rng of callSpreadRanges) {
      bullCallSpreads.push(...parseSpreadsFromTableForSymbol_(rng.getValues(), symbol, "CALL"));
    }
    for (const rng of putSpreadRanges) {
      bullPutSpreads.push(...parseSpreadsFromTableForSymbol_(rng.getValues(), symbol, "PUT"));
    }
    if (icRange) {
      const icPositions = parseIronCondorsFromTableForSymbol_(icRange.getValues(), symbol);
      bearCallSpreads.push(...icPositions.bearCallSpreads);
      bullPutSpreads.push(...icPositions.bullPutSpreads);
    }
  }

  const totalSpreads = bullCallSpreads.length + bullPutSpreads.length + bearCallSpreads.length;
  if (shares.length === 0 && totalSpreads === 0) {
    writeStatus_(sheet, `No open positions for ${symbol}.`);
    return;
  }

  // ── ROI denominators ────────────────────────────────────────────────
  // Shares denominator: total cost basis
  const sharesCost = shares.reduce((sum, sh) => sum + sh.qty * sh.basis, 0);

  // Bull call spread: investment = debit paid
  const bullCallInvest = bullCallSpreads.reduce((sum, sp) => sum + (sp.debit * 100 * sp.qty), 0);

  // Bull put spread (credit): max risk = width - credit
  const bullPutRisk = bullPutSpreads.reduce((sum, sp) => {
    const width = sp.kShort - sp.kLong;
    if (width <= 0) return sum;
    const credit = -sp.debit; // debit stored negative for credit spreads
    const maxLossPer = (width - credit) * 100;
    return sum + maxLossPer * sp.qty;
  }, 0);

  // Bear call spread (credit): max risk = width - credit
  const bearCallRisk = bearCallSpreads.reduce((sum, sp) => {
    const width = sp.kShort - sp.kLong;
    if (width <= 0) return sum;
    const credit = -sp.debit; // debit stored negative for credit spreads
    const maxLossPer = (width - credit) * 100;
    return sum + maxLossPer * sp.qty;
  }, 0);

  const optionsDenom = bullCallInvest + bullPutRisk + bearCallRisk;

  // ── Collect all spreads with labels for individual charting ────────────────
  const allSpreads = [...bullCallSpreads, ...bullPutSpreads, ...bearCallSpreads];

  // ── Calculate individual spread investments (for ROI) ────────────────
  const spreadInvestments = allSpreads.map((sp) => {
    if (sp.flavor === "CALL") {
      // Bull call spread: investment = debit * 100 * qty
      return sp.debit * 100 * sp.qty;
    } else {
      // Credit spreads (PUT or BEAR_CALL): investment = max risk = (width - credit) * 100 * qty
      const width = sp.kShort - sp.kLong;
      const credit = -sp.debit;
      return (width - credit) * 100 * sp.qty;
    }
  });

  // ── Build strategy groups (only non-empty types) ────────────────────
  const strategyGroupDefs = [
    { name: "Bull Call Spreads", spreads: bullCallSpreads },
    { name: "Bull Put Spreads", spreads: bullPutSpreads },
    { name: "Bear Call Spreads", spreads: bearCallSpreads },
  ];
  const strategyGroups = strategyGroupDefs.filter(g => g.spreads.length > 0);
  const strategyLabels = strategyGroups.map(g => g.name);

  // Strategy-level investment denominators (sum of individual investments within each group)
  let spreadIdx = 0;
  const strategyInvestments = strategyGroups.map(g => {
    let sum = 0;
    for (let i = 0; i < g.spreads.length; i++) {
      sum += spreadInvestments[spreadIdx++];
    }
    return sum;
  });

  // ── Build data table ────────────────────────────────────────────────
  // Structure: Price | Shares $ | [strategy $] | Total $ | Shares % ROI | [strategy % ROI] | [individual spread $] | [individual spread % ROI]
  const spreadLabels = allSpreads.map(sp => sp.label);
  const spreadRoiLabels = allSpreads.map(sp => sp.label + " %");
  const strategyRoiLabels = strategyLabels.map(l => l + " %");
  const headerRow = [
    "Price ($)", "Shares $",
    ...strategyLabels.map(l => l + " $"),
    "Total $",
    "Shares % ROI",
    ...strategyRoiLabels,
    ...spreadLabels,
    ...spreadRoiLabels,
  ];
  const table = [headerRow];

  // Column indexes for chart building
  const colPrice = 0;
  const colShares = 1;
  const colFirstStrategy = 2;
  const colTotal = 2 + strategyLabels.length;
  const colSharesRoi = colTotal + 1;
  const colFirstStrategyRoi = colSharesRoi + 1;
  const colFirstSpread = colFirstStrategyRoi + strategyLabels.length;
  const colFirstSpreadRoi = colFirstSpread + spreadLabels.length;

  for (let S = cfg.minPrice; S <= cfg.maxPrice + 1e-9; S += cfg.step) {
    // Calculate portfolio VALUE at price S (not P/L)

    // Shares value = current market value
    let sharesValue = 0;
    for (const sh of shares) {
      sharesValue += S * sh.qty;
    }

    // Track individual spread values
    const individualSpreadValues = [];

    // Bull call (debit) spread value at expiration
    for (const sp of bullCallSpreads) {
      const width = sp.kShort - sp.kLong;
      if (width <= 0) {
        individualSpreadValues.push(0);
        continue;
      }
      const intrinsic = clamp_(S - sp.kLong, 0, width);
      const valuePerSpread = intrinsic * 100; // intrinsic value in dollars
      const totalValue = valuePerSpread * sp.qty;
      individualSpreadValues.push(roundTo_(totalValue, 2));
    }

    // Bull put (credit) spread value at expiration
    for (const sp of bullPutSpreads) {
      const width = sp.kShort - sp.kLong;
      if (width <= 0) {
        individualSpreadValues.push(0);
        continue;
      }
      const loss = clamp_(sp.kShort - S, 0, width);
      const valuePerSpread = (width - loss) * 100; // what you keep
      const totalValue = valuePerSpread * sp.qty;
      individualSpreadValues.push(roundTo_(totalValue, 2));
    }

    // Bear call (credit) spread value at expiration
    for (const sp of bearCallSpreads) {
      const width = sp.kShort - sp.kLong;
      if (width <= 0) {
        individualSpreadValues.push(0);
        continue;
      }
      // Loss is how much ITM the short call is
      const loss = clamp_(S - sp.kLong, 0, width);
      const valuePerSpread = (width - loss) * 100; // what you keep
      const totalValue = valuePerSpread * sp.qty;
      individualSpreadValues.push(roundTo_(totalValue, 2));
    }

    const optionsValue = individualSpreadValues.reduce((sum, v) => sum + v, 0);
    const totalValue = sharesValue + optionsValue;

    // ROI based on P/L
    const sharesPL = sharesValue - sharesCost;
    const sharesROI = sharesCost > 0 ? (sharesPL / sharesCost) : 0;

    // Calculate strategy-level values by summing individual spreads per group
    const strategyValues = [];
    let sIdx = 0;
    for (const g of strategyGroups) {
      let groupVal = 0;
      for (let i = 0; i < g.spreads.length; i++) {
        groupVal += individualSpreadValues[sIdx++];
      }
      strategyValues.push(roundTo_(groupVal, 2));
    }

    // Calculate strategy-level ROIs
    const strategyROIs = strategyValues.map((val, idx) => {
      const inv = strategyInvestments[idx];
      if (inv <= 0) return 0;
      return roundTo_((val - inv) / inv, 4);
    });

    // Calculate individual spread ROIs
    const individualSpreadROIs = individualSpreadValues.map((val, idx) => {
      const investment = spreadInvestments[idx];
      if (investment <= 0) return 0;
      const pl = val - investment;
      return roundTo_(pl / investment, 4);
    });

    table.push([
      roundTo_(S, 2),
      roundTo_(sharesValue, 2),
      ...strategyValues,
      roundTo_(totalValue, 2),
      roundTo_(sharesROI, 4),
      ...strategyROIs,
      ...individualSpreadValues,
      ...individualSpreadROIs,
    ]);
  }

  // ── Write to sheet ──────────────────────────────────────────────────
  const startRow = cfg.tableStartRow; // default 50
  const startCol = cfg.tableStartCol;

  // Clear only the data output area (avoid nuking config)
  sheet.getRange(startRow - 1, startCol, 3000, 50).clearContent();

  sheet.getRange(startRow - 1, startCol).setValue("Data (generated)").setFontWeight("bold");
  sheet.getRange(startRow, startCol, table.length, table[0].length).setValues(table);
  sheet.autoResizeColumns(startCol, table[0].length);

  // Format data columns
  if (table.length > 1) {
    const dataRows = table.length - 1;
    const commaFmt = "#,##0";

    // Price column
    sheet.getRange(startRow + 1, startCol + colPrice, dataRows, 1).setNumberFormat(commaFmt);

    // Shares $ column
    sheet.getRange(startRow + 1, startCol + colShares, dataRows, 1).setNumberFormat(commaFmt);

    // Strategy $ columns
    if (strategyLabels.length > 0) {
      sheet.getRange(startRow + 1, startCol + colFirstStrategy, dataRows, strategyLabels.length).setNumberFormat(commaFmt);
    }

    // Total $ column
    sheet.getRange(startRow + 1, startCol + colTotal, dataRows, 1).setNumberFormat(commaFmt);

    // Individual spread $ columns
    if (spreadLabels.length > 0) {
      sheet.getRange(startRow + 1, startCol + colFirstSpread, dataRows, spreadLabels.length).setNumberFormat(commaFmt);
    }

    // Shares % ROI column
    sheet.getRange(startRow + 1, startCol + colSharesRoi, dataRows, 1).setNumberFormat("0.00%");

    // Strategy ROI columns
    if (strategyLabels.length > 0) {
      sheet.getRange(startRow + 1, startCol + colFirstStrategyRoi, dataRows, strategyLabels.length).setNumberFormat("0.00%");
    }

    // Individual spread ROI columns
    if (spreadLabels.length > 0) {
      sheet.getRange(startRow + 1, startCol + colFirstSpreadRoi, dataRows, spreadLabels.length).setNumberFormat("0.00%");
    }
  }

  // ── Charts: ensure charts exist ──────────────────────────────────────
  ensureFourCharts_(sheet, symbol, cfg, {
    startRow,
    startCol,
    tableRows: table.length,
    headerRow,
    spreadLabels,
    spreadRoiLabels,
    strategyLabels,
    colPrice,
    colShares,
    colFirstStrategy,
    colTotal,
    colSharesRoi,
    colFirstStrategyRoi,
    colFirstSpread,
    colFirstSpreadRoi,
  });
}

/* =========================================================
   Charts
   ========================================================= */

function ensureFourCharts_(sheet, symbol, cfg, args) {
  const { startRow, startCol, tableRows, headerRow, spreadLabels, spreadRoiLabels,
          strategyLabels, colPrice, colShares, colFirstStrategy, colTotal, colSharesRoi,
          colFirstStrategyRoi, colFirstSpread, colFirstSpreadRoi } = args;

  const dollarTitle = (cfg.chartTitle || `${symbol} Portfolio Value by Price`) + " ($)";
  const pctTitle = (cfg.chartTitle || `${symbol} Portfolio Value by Price`) + " (%)";
  const spreadsTitle = `${symbol} Individual Spreads ($)`;
  const spreadsRoiTitle = `${symbol} Individual Spreads ROI`;

  const existingCharts = sheet.getCharts();

  // Identify charts by title substring (best-effort)
  let dollarChart = null;
  let pctChart = null;
  let spreadsChart = null;
  let spreadsRoiChart = null;

  for (const ch of existingCharts) {
    const t = extractChartTitle_(ch);
    if (!t) continue;
    if (t === dollarTitle || (t.indexOf(" ($)") >= 0 && t.indexOf("Spreads") < 0)) {
      dollarChart = dollarChart || ch;
    } else if (t === pctTitle || (t.indexOf(" (%)") >= 0 && t.indexOf("Spreads") < 0)) {
      pctChart = pctChart || ch;
    } else if (t === spreadsTitle || (t.indexOf("Spreads") >= 0 && t.indexOf("($)") >= 0)) {
      spreadsChart = spreadsChart || ch;
    } else if (t === spreadsRoiTitle || (t.indexOf("Spreads") >= 0 && t.indexOf("ROI") >= 0)) {
      spreadsRoiChart = spreadsRoiChart || ch;
    }
  }

  // Chart size: wide enough for legend, tall enough to cover 25 rows
  const chartWidth = 900;
  const chartHeight = 500;

  function rebuildChartPreserveBox_(oldChart, newBuilder, defaultRow, defaultCol) {
    newBuilder = newBuilder.setOption("width", chartWidth).setOption("height", chartHeight);
    if (!oldChart) {
      sheet.insertChart(newBuilder.setPosition(defaultRow, defaultCol, 0, 0).build());
      return;
    }
    const ci = oldChart.getContainerInfo();
    const anchorRow = ci.getAnchorRow();
    const anchorCol = ci.getAnchorColumn();
    const offsetX = ci.getOffsetX();
    const offsetY = ci.getOffsetY();

    sheet.removeChart(oldChart);
    sheet.insertChart(newBuilder.setPosition(anchorRow, anchorCol, offsetX, offsetY).build());
  }

  // --- Build $ chart: Price, Shares $, [strategy $], Total $ ---
  const priceColRange = sheet.getRange(startRow, startCol + colPrice, tableRows, 1);
  const sharesColRange = sheet.getRange(startRow, startCol + colShares, tableRows, 1);
  const totalColRange = sheet.getRange(startRow, startCol + colTotal, tableRows, 1);

  let dollarBuilder = sheet
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(priceColRange)
    .addRange(sharesColRange);

  // Add strategy $ columns
  if (strategyLabels && strategyLabels.length > 0) {
    const strategyRange = sheet.getRange(startRow, startCol + colFirstStrategy, tableRows, strategyLabels.length);
    dollarBuilder = dollarBuilder.addRange(strategyRange);
  }

  dollarBuilder = dollarBuilder
    .addRange(totalColRange)
    .setOption("title", dollarTitle)
    .setOption("hAxis", { title: `${symbol} Price ($)` })
    .setOption("hAxis.format", "#,##0")
    .setOption("vAxis", { title: "Portfolio Value ($)" })
    .setOption("vAxis.format", "#,##0")
    .setOption("legend", { position: "right" })
    .setOption("curveType", "none");

  // Series labels: Shares $, [strategy labels $], Total $
  const dollarSeries = { 0: { labelInLegend: "Shares $" } };
  for (let i = 0; i < strategyLabels.length; i++) {
    dollarSeries[i + 1] = { labelInLegend: strategyLabels[i] + " $" };
  }
  dollarSeries[strategyLabels.length + 1] = { labelInLegend: "Total $" };
  dollarBuilder = dollarBuilder.setOption("series", dollarSeries);

  // --- Build % chart: Price, Shares % ROI, [strategy % ROI] ---
  const sharesRoiRange = sheet.getRange(startRow, startCol + colSharesRoi, tableRows, 1);

  let pctBuilder = sheet
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(priceColRange)
    .addRange(sharesRoiRange);

  // Add strategy ROI columns
  if (strategyLabels && strategyLabels.length > 0) {
    const strategyRoiRange = sheet.getRange(startRow, startCol + colFirstStrategyRoi, tableRows, strategyLabels.length);
    pctBuilder = pctBuilder.addRange(strategyRoiRange);
  }

  pctBuilder = pctBuilder
    .setOption("title", pctTitle)
    .setOption("hAxis", { title: `${symbol} Price ($)` })
    .setOption("hAxis.format", "#,##0")
    .setOption("vAxis", { title: "% Return (ROI)" })
    .setOption("vAxis.format", "percent")
    .setOption("legend", { position: "right" })
    .setOption("curveType", "none");

  // Series labels: Shares % ROI, [strategy labels %]
  const pctSeries = { 0: { labelInLegend: "Shares % ROI" } };
  for (let i = 0; i < strategyLabels.length; i++) {
    pctSeries[i + 1] = { labelInLegend: strategyLabels[i] + " %" };
  }
  pctBuilder = pctBuilder.setOption("series", pctSeries);

  // --- Build individual spreads value chart (unchanged) ---
  let spreadsBuilder = null;
  let spreadsRoiBuilder = null;

  if (spreadLabels && spreadLabels.length > 0) {
    const spreadsRange = sheet.getRange(startRow, startCol + colFirstSpread, tableRows, spreadLabels.length);

    spreadsBuilder = sheet
      .newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(priceColRange)
      .addRange(spreadsRange)
      .setOption("title", spreadsTitle)
      .setOption("hAxis", { title: `${symbol} Price ($)` })
      .setOption("hAxis.format", "#,##0")
      .setOption("vAxis", { title: "Spread Value ($)" })
      .setOption("vAxis.format", "#,##0")
      .setOption("legend", { position: "right" })
      .setOption("curveType", "none");

    const spreadsSeries = {};
    for (let i = 0; i < spreadLabels.length; i++) {
      spreadsSeries[i] = { labelInLegend: spreadLabels[i] };
    }
    spreadsBuilder = spreadsBuilder.setOption("series", spreadsSeries);

    // --- Build individual spreads ROI chart (unchanged) ---
    const spreadsRoiRange = sheet.getRange(startRow, startCol + colFirstSpreadRoi, tableRows, spreadLabels.length);

    spreadsRoiBuilder = sheet
      .newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(priceColRange)
      .addRange(spreadsRoiRange)
      .setOption("title", spreadsRoiTitle)
      .setOption("hAxis", { title: `${symbol} Price ($)` })
      .setOption("hAxis.format", "#,##0")
      .setOption("vAxis", { title: "% Return (ROI)" })
      .setOption("vAxis.format", "percent")
      .setOption("legend", { position: "right" })
      .setOption("curveType", "none");

    const spreadsRoiSeries = {};
    for (let i = 0; i < spreadLabels.length; i++) {
      spreadsRoiSeries[i] = { labelInLegend: spreadLabels[i] };
    }
    spreadsRoiBuilder = spreadsRoiBuilder.setOption("series", spreadsRoiSeries);
  }

  // Create / rebuild while preserving chart placement
  // Each chart spans 25 rows, with 1 gap row between them
  rebuildChartPreserveBox_(dollarChart, dollarBuilder, 1, 1);
  rebuildChartPreserveBox_(pctChart, pctBuilder, 27, 1);
  if (spreadsBuilder) {
    rebuildChartPreserveBox_(spreadsChart, spreadsBuilder, 53, 1);
  }
  if (spreadsRoiBuilder) {
    rebuildChartPreserveBox_(spreadsRoiChart, spreadsRoiBuilder, 79, 1);
  }
}

function extractChartTitle_(chart) {
  try {
    const opts = chart.getOptions && chart.getOptions();
    if (!opts) return "";
    // opts can be a plain object in some contexts; use both patterns
    if (typeof opts.get === "function") {
      const t = opts.get("title");
      return t ? String(t) : "";
    }
    if (opts.title) return String(opts.title);
    return "";
  } catch (e) {
    return "";
  }
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
   Config table – now in K:L, default tableStartRow = 85
   Includes descriptions/labels in-column K.
   ========================================================= */

function ensureAndReadConfig_(ss, sheet, symbol) {
  const smart = computeSmartDefaults_(ss, symbol);
  const defaults = {
    minPrice: smart.minPrice,
    maxPrice: smart.maxPrice,
    step: smart.step,
    tableStartRow: 200, // default is 200
    tableStartCol: 1,
    chartTitle: `${symbol} Portfolio Value by Price`,
  };

  // Config location
  const cfgRow = 1;
  const cfgCol = 11; // K
  const cfgNumRows = 10;
  const cfgNumCols = 2;

  // Column K contains human-readable description; column L contains value.
  // Row 1 is header; remaining rows are settings.
  const values = [
    ["Config Setting (Description)", "Value"],
    ["Min price on x-axis (inclusive)", defaults.minPrice],
    ["Max price on x-axis (inclusive)", defaults.maxPrice],
    ["Price step size", defaults.step],
    ["Data table start row (header row)", defaults.tableStartRow],
    ["Data table start column", defaults.tableStartCol],
    ["Chart title base (used for both charts)", defaults.chartTitle],
    ["(Edit values in column L; config is not hidden)", ""],
    ["", ""],
    ["", ""],
  ];

  // If config not present, write it
  const header = sheet.getRange(cfgRow, cfgCol).getValue();
  if (String(header).trim() !== "Config Setting (Description)") {
    sheet.getRange(cfgRow, cfgCol, cfgNumRows, cfgNumCols).setValues(values);
    sheet.getRange(cfgRow, cfgCol, 1, cfgNumCols).setFontWeight("bold");
    sheet.autoResizeColumns(cfgCol, cfgNumCols);
  }

  const cfgRange = sheet.getRange(cfgRow, cfgCol, cfgNumRows, cfgNumCols);

  // Named range per symbol (global namespace)
  const cfgName = `Config_${symbol}`;
  const existing = ss.getNamedRanges().find(nr => nr.getName() === cfgName);
  if (!existing) {
    ss.setNamedRange(cfgName, cfgRange);
  } else if (!rangesEqual_(existing.getRange(), cfgRange)) {
    existing.remove();
    ss.setNamedRange(cfgName, cfgRange);
  }

  // Read values from column L by row number (stable; avoids needing "keys")
  // Row mapping:
  // 2=minPrice, 3=maxPrice, 4=step, 5=tableStartRow, 6=tableStartCol, 7=chartTitle
  const raw = sheet.getRange(cfgRow + 1, cfgCol + 1, 6, 1).getValues().map(r => r[0]);

  const cfg = {
    minPrice: numOr_(raw[0], defaults.minPrice),
    maxPrice: numOr_(raw[1], defaults.maxPrice),
    step: numOr_(raw[2], defaults.step),
    tableStartRow: Math.max(5, Math.floor(numOr_(raw[3], defaults.tableStartRow))),
    tableStartCol: Math.max(1, Math.floor(numOr_(raw[4], defaults.tableStartCol))),
    chartTitle: String(raw[5] ?? defaults.chartTitle),
  };

  if (!(cfg.minPrice < cfg.maxPrice)) {
    cfg.minPrice = defaults.minPrice;
    cfg.maxPrice = defaults.maxPrice;
  }
  if (!(cfg.step > 0)) cfg.step = defaults.step;

  return cfg;
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

/* =========================================================
   Parsing with optional Symbol column filtering
   ========================================================= */

// parseSharesFromTableForSymbol_ is in Parsing.js

// parseSpreadsFromTableForSymbol_, formatExpirationLabel_, parseIronCondorsFromTableForSymbol_ are in Parsing.js

/* =========================================================
   Status messaging
   ========================================================= */

function writeStatus_(sheet, message) {
  sheet.getRange("D1").setValue("Status").setFontWeight("bold");
  sheet.getRange("D2").setValue(message).setWrap(true);
}

function clearStatus_(sheet) {
  sheet.getRange("D1:D2").clearContent();
}

/* =========================================================
   Range helpers
   ========================================================= */

// rangesIntersect_ is in CommonUtils.js

