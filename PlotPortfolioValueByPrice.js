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

  for (const symbol of symbols) {
    plotForSymbol_(ss, symbol);
  }
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
  const symbols = new Set();

  const tableNames = ["Stocks", "BullCallSpreads", "IronCondors"];

  for (const tableName of tableNames) {
    const range = getNamedRangeWithTableFallback_(ss, tableName);
    if (!range) continue;

    const rows = range.getValues();
    if (rows.length < 2) continue;

    const headerNorm = rows[0].map(normKey_);
    const idxSym = findCol_(headerNorm, ["symbol", "ticker"]);
    if (idxSym < 0) continue;

    // Check for Status column (used by IronCondors to mark closed positions)
    const idxStatus = findCol_(headerNorm, ["status"]);

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

  // Get ranges from known position tables
  const stockRanges = [];
  const callSpreadRanges = [];
  const putSpreadRanges = [];

  const stocksRange = getNamedRangeWithTableFallback_(ss, "Stocks");
  if (stocksRange) stockRanges.push(stocksRange);

  const bcsRange = getNamedRangeWithTableFallback_(ss, "BullCallSpreads");
  if (bcsRange) callSpreadRanges.push(bcsRange);

  const icRange = getNamedRangeWithTableFallback_(ss, "IronCondors");

  if (stockRanges.length === 0 && callSpreadRanges.length === 0 && !icRange) {
    writeStatus_(sheet, `No position tables found. Create named ranges:\n  - Stocks or StocksTable\n  - BullCallSpreads or BullCallSpreadsTable\n  - IronCondors or IronCondorsTable`);
    return;
  }

  clearStatus_(sheet);

  // ── Parse positions ────────────────────────────────────────────────
  const shares = [];
  const meta = { currentPrice: null };

  for (const rng of stockRanges) {
    shares.push(...parseSharesFromTableForSymbol_(rng.getValues(), symbol, meta));
  }

  const bullCallSpreads = [];
  for (const rng of callSpreadRanges) {
    bullCallSpreads.push(...parseSpreadsFromTableForSymbol_(rng.getValues(), symbol, "CALL"));
  }

  const bullPutSpreads = [];
  for (const rng of putSpreadRanges) {
    bullPutSpreads.push(...parseSpreadsFromTableForSymbol_(rng.getValues(), symbol, "PUT"));
  }

  // Bear call spreads (from iron condors)
  const bearCallSpreads = [];

  // Parse iron condors
  if (icRange) {
    const icPositions = parseIronCondorsFromTableForSymbol_(icRange.getValues(), symbol);
    bearCallSpreads.push(...icPositions.bearCallSpreads);
    bullPutSpreads.push(...icPositions.bullPutSpreads);
  }

  const totalSpreads = bullCallSpreads.length + bullPutSpreads.length + bearCallSpreads.length;
  if (shares.length === 0 && totalSpreads === 0) {
    writeStatus_(sheet, `Parsed 0 valid rows for ${symbol}.\nCheck headers and numeric values.`);
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

  // ── Build data table ────────────────────────────────────────────────
  // Base headers + individual spread value columns + individual spread ROI columns
  const baseHeaders = ["Price ($)", "Shares $", "Options $", "Total $", "Shares % ROI", "Options % ROI"];
  const spreadLabels = allSpreads.map(sp => sp.label);
  const spreadRoiLabels = allSpreads.map(sp => sp.label + " %");
  const headerRow = [...baseHeaders, ...spreadLabels, ...spreadRoiLabels];
  const table = [headerRow];

  for (let S = cfg.minPrice; S <= cfg.maxPrice + 1e-9; S += cfg.step) {
    // Calculate portfolio VALUE at price S (not P/L)

    // Shares value = current market value
    let sharesValue = 0;
    for (const sh of shares) {
      sharesValue += S * sh.qty;
    }

    // Options value = intrinsic value at expiration
    let optionsValue = 0;

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
      optionsValue += totalValue;
      individualSpreadValues.push(round2_(totalValue));
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
      optionsValue += totalValue;
      individualSpreadValues.push(round2_(totalValue));
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
      optionsValue += totalValue;
      individualSpreadValues.push(round2_(totalValue));
    }

    const totalValue = sharesValue + optionsValue;

    // ROI still based on P/L
    const sharesPL = sharesValue - sharesCost;
    const optionsPL = optionsValue - bullCallInvest - bullPutRisk - bearCallRisk;

    const sharesROI = sharesCost > 0 ? (sharesPL / sharesCost) : 0;
    const optionsROI = optionsDenom > 0 ? (optionsPL / optionsDenom) : 0;

    // Calculate individual spread ROIs
    const individualSpreadROIs = individualSpreadValues.map((val, idx) => {
      const investment = spreadInvestments[idx];
      if (investment <= 0) return 0;
      const pl = val - investment;
      return round4_(pl / investment);
    });

    table.push([
      round2_(S),
      round2_(sharesValue),
      round2_(optionsValue),
      round2_(totalValue),
      round4_(sharesROI),
      round4_(optionsROI),
      ...individualSpreadValues,
      ...individualSpreadROIs,
    ]);
  }

  // For vline series bounds on $ chart (use Total $)
  const totalYs = table.slice(1).map(r => toNum_(r[3])).filter(n => Number.isFinite(n));
  const minY = totalYs.length ? Math.min(...totalYs) : 0;
  const maxY = totalYs.length ? Math.max(...totalYs) : 0;

  // ── Write to sheet ──────────────────────────────────────────────────
  const startRow = cfg.tableStartRow; // default 50
  const startCol = cfg.tableStartCol;

  // Clear only the data output area (avoid nuking config)
  sheet.getRange(startRow - 1, startCol, 3000, 30).clearContent();

  sheet.getRange(startRow - 1, startCol).setValue("Data (generated)").setFontWeight("bold");
  sheet.getRange(startRow, startCol, table.length, table[0].length).setValues(table);
  sheet.autoResizeColumns(startCol, table[0].length);

  // Format ROI columns as percent
  if (table.length > 1) {
    // columns: Price=0, Shares$=1, Options$=2, Total$=3, Shares%=4, Options%=5
    sheet.getRange(startRow + 1, startCol + 4, table.length - 1, 2).setNumberFormat("0.00%");

    // Format individual spread ROI columns as percent (after spread value columns)
    if (spreadLabels.length > 0) {
      const spreadRoiStartCol = startCol + baseHeaders.length + spreadLabels.length;
      sheet.getRange(startRow + 1, spreadRoiStartCol, table.length - 1, spreadLabels.length).setNumberFormat("0.00%");
    }
  }

  // ── Charts: ensure FOUR charts exist (create missing; refresh existing preserving box) ──────────────────
  ensureFourCharts_(sheet, symbol, cfg, {
    startRow,
    startCol,
    tableRows: table.length,
    headerRow,
    spreadLabels,
    spreadRoiLabels,
    baseHeaderCount: baseHeaders.length,
  });
}

/* =========================================================
   Charts
   ========================================================= */

function ensureFourCharts_(sheet, symbol, cfg, args) {
  const { startRow, startCol, tableRows, headerRow, spreadLabels, spreadRoiLabels, baseHeaderCount } = args;

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

  // Ranges:
  // Dollars chart uses contiguous 4 columns: Price, Shares$, Options$, Total$
  const dollarRange = sheet.getRange(startRow, startCol, tableRows, 4);

  // Percent chart uses: Price + (Shares% ROI) + (Options% ROI)
  const priceColRange = sheet.getRange(startRow, startCol, tableRows, 1);
  const sharesPctRange = sheet.getRange(startRow, startCol + 4, tableRows, 1);
  const optionsPctRange = sheet.getRange(startRow, startCol + 5, tableRows, 1);

  function rebuildChartPreserveBox_(oldChart, newBuilder, defaultRow, defaultCol) {
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

  // --- Build $ chart builder ---
  let dollarBuilder = sheet
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dollarRange)
    .setOption("title", dollarTitle)
    .setOption("hAxis", { title: `${symbol} Price ($)` })
    .setOption("vAxis", { title: "Portfolio Value ($)" })
    .setOption("legend", { position: "right" })
    .setOption("curveType", "none");

  // Ensure series labels come from header row:
  const dollarSeries = {
    0: { labelInLegend: headerRow[1] },
    1: { labelInLegend: headerRow[2] },
    2: { labelInLegend: headerRow[3] },
  };

  dollarBuilder = dollarBuilder.setOption("series", dollarSeries);

  // --- Build % chart builder ---
  let pctBuilder = sheet
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(priceColRange)
    .addRange(sharesPctRange)
    .addRange(optionsPctRange)
    .setOption("title", pctTitle)
    .setOption("hAxis", { title: `${symbol} Price ($)` })
    .setOption("vAxis", { title: "% Return (ROI)", format: "percent" })
    .setOption("legend", { position: "right" })
    .setOption("curveType", "none")
    .setOption("series", {
      0: { labelInLegend: headerRow[4] },
      1: { labelInLegend: headerRow[5] },
    });

  // --- Build individual spreads value chart builder ---
  let spreadsBuilder = null;
  let spreadsRoiBuilder = null;

  if (spreadLabels && spreadLabels.length > 0) {
    // Individual spread value columns start after base headers
    const spreadsRange = sheet.getRange(startRow, startCol + baseHeaderCount, tableRows, spreadLabels.length);

    spreadsBuilder = sheet
      .newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(priceColRange)
      .addRange(spreadsRange)
      .setOption("title", spreadsTitle)
      .setOption("hAxis", { title: `${symbol} Price ($)` })
      .setOption("vAxis", { title: "Spread Value ($)" })
      .setOption("legend", { position: "right" })
      .setOption("curveType", "none");

    // Set series labels from spread labels
    const spreadsSeries = {};
    for (let i = 0; i < spreadLabels.length; i++) {
      spreadsSeries[i] = { labelInLegend: spreadLabels[i] };
    }
    spreadsBuilder = spreadsBuilder.setOption("series", spreadsSeries);

    // --- Build individual spreads ROI chart builder ---
    // ROI columns start after spread value columns
    const spreadsRoiRange = sheet.getRange(startRow, startCol + baseHeaderCount + spreadLabels.length, tableRows, spreadLabels.length);

    spreadsRoiBuilder = sheet
      .newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(priceColRange)
      .addRange(spreadsRoiRange)
      .setOption("title", spreadsRoiTitle)
      .setOption("hAxis", { title: `${symbol} Price ($)` })
      .setOption("vAxis", { title: "% Return (ROI)", format: "percent" })
      .setOption("legend", { position: "right" })
      .setOption("curveType", "none");

    // Set series labels (without the " %" suffix for cleaner legend)
    const spreadsRoiSeries = {};
    for (let i = 0; i < spreadLabels.length; i++) {
      spreadsRoiSeries[i] = { labelInLegend: spreadLabels[i] };
    }
    spreadsRoiBuilder = spreadsRoiBuilder.setOption("series", spreadsRoiSeries);
  }

  // Create / rebuild while preserving chart placement
  // Default positions if missing:
  //   $ chart: top-left
  //   % chart: below it
  //   spreads $ chart: below % chart
  //   spreads ROI chart: below spreads $ chart
  rebuildChartPreserveBox_(dollarChart, dollarBuilder, 1, 1);
  rebuildChartPreserveBox_(pctChart, pctBuilder, 15, 1);
  if (spreadsBuilder) {
    rebuildChartPreserveBox_(spreadsChart, spreadsBuilder, 29, 1);
  }
  if (spreadsRoiBuilder) {
    rebuildChartPreserveBox_(spreadsRoiChart, spreadsRoiBuilder, 43, 1);
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
   Config table – now in K:L, default tableStartRow = 85
   Includes descriptions/labels in-column K.
   ========================================================= */

function ensureAndReadConfig_(ss, sheet, symbol) {
  const defaults = {
    minPrice: 350,
    maxPrice: 900,
    step: 5,
    tableStartRow: 85, // default is 85
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

function parseSharesFromTableForSymbol_(rows, symbol, outMeta) {
  if (!rows || rows.length < 2) return [];

  const headerNorm = rows[0].map(normKey_);

  const idxSym = findCol_(headerNorm, ["symbol", "ticker"]);
  const idxQty = findCol_(headerNorm, ["shares", "share", "qty", "quantity", "units", "position"]);
  const idxBasis = findCol_(headerNorm, [
    "costbasis",
    "basis",
    "avgcost",
    "averagecost",
    "avgprice",
    "averageprice",
    "aveprice",
    "avepricepaid",
    "pricepaid",
    "entry",
    "entryprice",
    "cost",
    "purchaseprice",
  ]);

  const out = [];

  for (let r = 1; r < rows.length; r++) {
    if (idxSym >= 0) {
      const rowSym = String(rows[r][idxSym] ?? "").trim().toUpperCase();
      if (rowSym && rowSym !== symbol) continue;
    }

    if (idxQty < 0 || idxBasis < 0) continue;

    const qty = toNum_(rows[r][idxQty]);
    const basis = toNum_(rows[r][idxBasis]);

    if (!Number.isFinite(qty) || qty === 0) continue;
    if (!Number.isFinite(basis)) continue;

    out.push({ qty, basis });
  }

  return out;
}

/**
 * Spread parsing:
 * - Counts ONLY "definition rows" that contain BOTH Long Strike and Short Strike
 * - Ignores fill-detail rows automatically
 * - Uses debit cost preference:
 *     Ave Debit / Avg Debit / Average Debit
 *     else Rec Debit / Recommended Debit
 *     else Net Debit / Debit / Cost / Entry / Price
 * - If contracts column missing, defaults qty=1
 * - Captures expiration date to generate labels like "Dec 28 350/450"
 */
function parseSpreadsFromTableForSymbol_(rows, symbol, flavor) {
  if (!rows || rows.length < 2) return [];

  const headerNorm = rows[0].map(normKey_);

  const idxSym = findCol_(headerNorm, ["symbol", "ticker"]);

  const idxQty = findCol_(headerNorm, [
    "contracts",
    "contract",
    "qty",
    "quantity",
    "count",
    "numcontracts",
    "spreads",
    "spreadqty",
  ]);

  const idxLong = findCol_(headerNorm, [
    "lower",
    "lowerstrike",
    "long",
    "longstrike",
    "buystrike",
    "strikebuy",
    "strikelong",
  ]);
  const idxShort = findCol_(headerNorm, [
    "upper",
    "upperstrike",
    "short",
    "shortstrike",
    "sellstrike",
    "strikesell",
    "strikeshort",
  ]);

  const idxExp = findCol_(headerNorm, [
    "expiration",
    "exp",
    "expiry",
    "expirationdate",
    "expdate",
  ]);

  const idxAveDebit = findCol_(headerNorm, ["avedebit", "avgdebit", "averagedebit"]);
  const idxRecDebit = findCol_(headerNorm, ["recdebit", "recommendeddebit"]);
  const idxDebitFallback = findCol_(headerNorm, ["netdebit", "debit", "cost", "price", "entry", "premium"]);

  if (idxLong < 0 || idxShort < 0) return [];

  const assumeQty = idxQty < 0;
  const out = [];

  for (let r = 1; r < rows.length; r++) {
    if (idxSym >= 0) {
      const rowSym = String(rows[r][idxSym] ?? "").trim().toUpperCase();
      if (rowSym && rowSym !== symbol) continue;
    }

    const kLong = toNum_(rows[r][idxLong]);
    const kShort = toNum_(rows[r][idxShort]);

    // Only definition rows have both strikes
    if (!Number.isFinite(kLong) || !Number.isFinite(kShort)) continue;

    const qty = assumeQty ? 1 : toNum_(rows[r][idxQty]);

    if (!Number.isFinite(qty) || qty === 0) continue;

    let debit = NaN;
    if (idxAveDebit >= 0) debit = toNum_(rows[r][idxAveDebit]);
    if (!Number.isFinite(debit) && idxRecDebit >= 0) debit = toNum_(rows[r][idxRecDebit]);
    if (!Number.isFinite(debit) && idxDebitFallback >= 0) debit = toNum_(rows[r][idxDebitFallback]);
    if (!Number.isFinite(debit)) continue;

    // Build label like "Dec 28 350/450"
    let label = `${kLong}/${kShort}`;
    if (idxExp >= 0) {
      const expLabel = formatExpLabel_(rows[r][idxExp]);
      if (expLabel) label = `${expLabel} ${label}`;
    }

    out.push({ qty, kLong, kShort, debit, flavor, label });
  }

  return out;
}

/**
 * Format expiration for chart label: "Dec 28" from a Date or string
 */
function formatExpLabel_(exp) {
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

  const idxSym = findCol_(headerNorm, ["symbol", "ticker"]);
  const idxExp = findCol_(headerNorm, ["expiration", "exp", "expiry", "expirationdate", "expdate"]);
  const idxStatus = findCol_(headerNorm, ["status"]);
  const idxBuyPut = findCol_(headerNorm, ["buyput", "longput", "putlong", "putbuy"]);
  const idxSellPut = findCol_(headerNorm, ["sellput", "shortput", "putshort", "putsell"]);
  const idxSellCall = findCol_(headerNorm, ["sellcall", "shortcall", "callshort", "callsell"]);
  const idxBuyCall = findCol_(headerNorm, ["buycall", "longcall", "calllong", "callbuy"]);
  const idxCredit = findCol_(headerNorm, ["credit", "netcredit", "premium"]);
  const idxQty = findCol_(headerNorm, ["qty", "quantity", "contracts", "contract", "count"]);

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
      const expLabel = formatExpLabel_(rows[r][idxExp]);
      if (expLabel) label = `${expLabel} ${label}`;
    }

    // Split credit between the two spreads (approximation)
    const putWidth = sellPut - buyPut;
    const callWidth = buyCall - sellCall;
    const totalWidth = putWidth + callWidth;
    const putCredit = totalWidth > 0 ? credit * (putWidth / totalWidth) : credit / 2;
    const callCredit = totalWidth > 0 ? credit * (callWidth / totalWidth) : credit / 2;

    // Bull put spread: long lower put, short higher put
    // debit is negative for credit spreads
    result.bullPutSpreads.push({
      qty,
      kLong: buyPut,
      kShort: sellPut,
      debit: -putCredit,
      flavor: "PUT",
      label: label + " (put)",
    });

    // Bear call spread: short lower call, long higher call
    // For bear call spreads, we use kLong=sellCall, kShort=buyCall
    // and mark it as a bear call so the value calc handles it correctly
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

function rangesIntersect_(a, b) {
  if (a.getSheet().getSheetId() !== b.getSheet().getSheetId()) return false;

  const aR1 = a.getRow(),
    aC1 = a.getColumn();
  const aR2 = aR1 + a.getNumRows() - 1;
  const aC2 = aC1 + a.getNumColumns() - 1;

  const bR1 = b.getRow(),
    bC1 = b.getColumn();
  const bR2 = bR1 + b.getNumRows() - 1;
  const bC2 = bC1 + b.getNumColumns() - 1;

  return aR1 <= bR2 && aR2 >= bR1 && aC1 <= bC2 && aC2 >= bC1;
}

function rangesEqual_(a, b) {
  return (
    a.getSheet().getSheetId() === b.getSheet().getSheetId() &&
    a.getRow() === b.getRow() &&
    a.getColumn() === b.getColumn() &&
    a.getNumRows() === b.getNumRows() &&
    a.getNumColumns() === b.getNumColumns()
  );
}

/* =========================================================
   Generic helpers
   ========================================================= */

function normKey_(v) {
  return String(v ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function findCol_(normHeaders, synonyms) {
  const syn = synonyms.map(normKey_);
  for (let i = 0; i < normHeaders.length; i++) {
    if (syn.includes(normHeaders[i])) return i;
  }
  return -1;
}

/**
 * Robust numeric parse:
 * - strips $, %, commas
 * - supports parentheses negatives: (123.45) => -123.45
 */
function toNum_(v) {
  if (v == null || v === "") return NaN;
  if (typeof v === "number") return v;

  let s = String(v).trim();
  if (!s) return NaN;

  const neg = /^\(.*\)$/.test(s);
  if (neg) s = s.slice(1, -1);

  s = s.replace(/[$,%]/g, "").replace(/,/g, "").trim();
  const n = Number(s);

  if (!Number.isFinite(n)) return NaN;
  return neg ? -n : n;
}

function numOr_(v, fallback) {
  const n = toNum_(v);
  return Number.isFinite(n) ? n : fallback;
}

function clamp_(x, lo, hi) {
  return Math.max(lo, Math.min(hi, x));
}

function round4_(x) {
  return Math.round(x * 10000) / 10000;
}