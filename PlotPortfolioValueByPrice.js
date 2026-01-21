/**
 * PlotPortfolioValueByPrice.js
 * ------------------------------------------------------------
 * PlotPortfolioValueByPrice
 *
 * Reads a "Portfolios" table describing which inputs belong to which symbol,
 * then creates/refreshes a "<SYMBOL>PortfolioValueByPrice" tab for EACH symbol.
 *
 * For each symbol tab:
 *   - Config table at K1:L10 (column K = 11) with labels/descriptions
 *   - Config columns are NEVER hidden
 *   - Default tableStartRow = 50
 *   - Generated data table starts at row tableStartRow (headers) / tableStartRow+1 (data)
 *   - TWO charts:
 *       Chart 1: $ P/L vs Price (Shares $, Options $, Total $) + vertical dashed line at current price (if available)
 *       Chart 2: % Return vs Price (Shares %, Options %) (ROI style, not contribution)
 *   - Chart series labels come from the first row of the data table (headers)
 *
 * Portfolios table columns: Symbol | Type | RangeName
 * Supported types (canonical singular): stock, bull-call-spread, bull-put-spread
 * Accepted aliases:
 *   - shares, share, stocks -> stock
 *   - bull-call-spreads, BCS -> bull-call-spread
 *   - bull-put-spreads,  BPS -> bull-put-spread
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
  const portfoliosBySymbol = ensureAndReadPortfolios_(ss);

  const symbols = Object.keys(portfoliosBySymbol);
  if (symbols.length === 0) {
    SpreadsheetApp.getUi().alert(
      "No symbols found in Portfolios.\n\n" +
        "Add rows to the 'Portfolios' sheet with columns:\n" +
        "  Symbol | Type | RangeName\n\n" +
        "Type must be one of:\n" +
        "  stock\n" +
        "  bull-call-spread\n" +
        "  bull-put-spread\n"
    );
    return;
  }

  for (const symbol of symbols) {
    plotForSymbol_(ss, symbol, portfoliosBySymbol[symbol]);
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
        const portfoliosBySymbol = ensureAndReadPortfolios_(ss);
        const entries = portfoliosBySymbol[String(symbol || "").trim().toUpperCase()] || [];
        plotForSymbol_(ss, symbol, entries);
        return;
      }
    }
  } catch (err) {
    // silent
  }
}

/* =========================================================
   Main per-symbol logic
   ========================================================= */

function plotForSymbol_(ss, symbolRaw, entries) {
  const symbol = String(symbolRaw || "").trim().toUpperCase();
  if (!symbol) return;

  const sheetName = `${symbol}PortfolioValueByPrice`;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.showSheet();

  const cfg = ensureAndReadConfig_(ss, sheet, symbol);

  if (!entries || entries.length === 0) {
    writeStatus_(sheet, `No portfolio inputs found for ${symbol} in the Portfolios table.`);
    return;
  }

  const stockRanges = [];
  const callSpreadRanges = [];
  const putSpreadRanges = [];
  const missingItems = [];

  for (const e of entries) {
    const rangeNameRaw = String(e.rangeName ?? "").trim();
    const type = String(e.type ?? "").trim();

    if (!rangeNameRaw || !type) continue;

    const rng = getNamedRangeWithTableFallback_(ss, rangeNameRaw);
    if (!rng) {
      missingItems.push({ type, rangeName: rangeNameRaw });
      continue;
    }

    if (type === "stock") stockRanges.push(rng);
    else if (type === "bull-call-spread") callSpreadRanges.push(rng);
    else if (type === "bull-put-spread") putSpreadRanges.push(rng);
  }

  if (
    missingItems.length > 0 ||
    (stockRanges.length === 0 && callSpreadRanges.length === 0 && putSpreadRanges.length === 0)
  ) {
    const msg = buildMissingRangeMessage_(
      ss,
      symbol,
      missingItems,
      stockRanges.length,
      callSpreadRanges.length + putSpreadRanges.length
    );
    writeStatus_(sheet, msg);
    return;
  }

  clearStatus_(sheet);

  // ── Parse positions ─────────────────────────────────────
  const shares = [];
  const meta = { currentPrice: null };

  for (const rng of stockRanges) {
    shares.push(...parseSharesFromTableForSymbol_(rng.getValues(), symbol, meta));
  }

  const callSpreads = [];
  for (const rng of callSpreadRanges) {
    callSpreads.push(...parseSpreadsFromTableForSymbol_(rng.getValues(), symbol, "CALL"));
  }

  const putSpreads = [];
  for (const rng of putSpreadRanges) {
    putSpreads.push(...parseSpreadsFromTableForSymbol_(rng.getValues(), symbol, "PUT"));
  }

  if (shares.length === 0 && callSpreads.length === 0 && putSpreads.length === 0) {
    writeStatus_(sheet, `Parsed 0 valid rows for ${symbol}.\nCheck headers and numeric values.`);
    return;
  }

  // ── ROI denominators ────────────────────────────────────
  // Shares denominator: total cost basis
  const sharesCost = shares.reduce((sum, sh) => sum + sh.qty * sh.basis, 0);

  // Options denominator: debit paid for call spreads + max-risk for put spreads
  const callInvest = callSpreads.reduce((sum, sp) => sum + (sp.debit * 100 * sp.qty), 0);
  const putRisk = putSpreads.reduce((sum, sp) => {
    const width = sp.kShort - sp.kLong;
    if (width <= 0) return sum;
    const credit = -sp.debit; // debit stored negative for credit spreads
    const maxLossPer = (width - credit) * 100;
    return sum + maxLossPer * sp.qty;
  }, 0);
  const optionsDenom = callInvest + putRisk;

  // ── Build data table ────────────────────────────────────
  // NOTE: % columns are ROI-style: PL / capital (shares cost basis, options denom)
  const headerRow = ["Price ($)", "Shares $ P/L", "Options $ P/L", "Total $ P/L", "Shares % ROI", "Options % ROI"];
  const table = [headerRow];

  for (let S = cfg.minPrice; S <= cfg.maxPrice + 1e-9; S += cfg.step) {
    let sharesPL = 0;
    let spreadsPL = 0;

    // Shares P/L at scenario price S
    for (const sh of shares) sharesPL += (S - sh.basis) * sh.qty;

    // Bull call (debit) spread payoff at expiration
    for (const sp of callSpreads) {
      const width = sp.kShort - sp.kLong;
      if (width <= 0) continue;
      const intrinsic = clamp_(S - sp.kLong, 0, width);
      spreadsPL += (intrinsic - sp.debit) * 100 * sp.qty;
    }

    // Bull put (credit) spread payoff at expiration
    for (const sp of putSpreads) {
      const width = sp.kShort - sp.kLong;
      if (width <= 0) continue;
      const loss = clamp_(sp.kShort - S, 0, width);
      spreadsPL += (-sp.debit - loss) * 100 * sp.qty;
    }

    const totalPL = sharesPL + spreadsPL;
    const sharesROI = sharesCost > 0 ? (sharesPL / sharesCost) : 0;
    const optionsROI = optionsDenom > 0 ? (spreadsPL / optionsDenom) : 0;

    table.push([
      round2_(S),
      round2_(sharesPL),
      round2_(spreadsPL),
      round2_(totalPL),
      round4_(sharesROI),
      round4_(optionsROI),
    ]);
  }

  // For vline series bounds on $ chart (use Total $)
  const totalYs = table.slice(1).map(r => toNum_(r[3])).filter(n => isFinite(n));
  const minY = totalYs.length ? Math.min(...totalYs) : 0;
  const maxY = totalYs.length ? Math.max(...totalYs) : 0;

  // ── Write to sheet ──────────────────────────────────────
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
  }

  // Vertical line helper table to the right of data table
  const vlineCol = startCol + table[0].length + 2;
  const vlineRow = startRow;

  sheet.getRange(vlineRow, vlineCol, 3, 2).clearContent();

  const hasCurrentPrice = isFinite(meta.currentPrice);
  if (hasCurrentPrice) {
    sheet.getRange(vlineRow, vlineCol, 3, 2).setValues([
      ["CurrentPrice", "Y"],
      [meta.currentPrice, minY],
      [meta.currentPrice, maxY],
    ]);
  }

  // ── Charts: ensure TWO charts exist (create missing; refresh existing preserving box) ───────────────
  ensureTwoCharts_(sheet, symbol, cfg, {
    startRow,
    startCol,
    tableRows: table.length,
    hasCurrentPrice,
    vlineRow,
    vlineCol,
    headerRow,
  });
}

/* =========================================================
   Charts
   ========================================================= */

function ensureTwoCharts_(sheet, symbol, cfg, args) {
  const { startRow, startCol, tableRows, hasCurrentPrice, vlineRow, vlineCol, headerRow } = args;

  const dollarTitle = (cfg.chartTitle || `${symbol} Portfolio Value by Price`) + " ($)";
  const pctTitle = (cfg.chartTitle || `${symbol} Portfolio Value by Price`) + " (%)";

  const existingCharts = sheet.getCharts();

  // Identify charts by title substring (best-effort)
  let dollarChart = null;
  let pctChart = null;

  for (const ch of existingCharts) {
    const t = extractChartTitle_(ch);
    if (!t) continue;
    if (t === dollarTitle || t.indexOf(" ($)") >= 0) dollarChart = dollarChart || ch;
    if (t === pctTitle || t.indexOf(" (%)") >= 0) pctChart = pctChart || ch;
  }

  // Ranges:
  // Dollars chart uses contiguous 4 columns: Price, Shares$, Options$, Total$
  const dollarRange = sheet.getRange(startRow, startCol, tableRows, 4);

  // Percent chart uses: Price + (Shares% ROI) + (Options% ROI)
  const priceColRange = sheet.getRange(startRow, startCol, tableRows, 1);
  const sharesPctRange = sheet.getRange(startRow, startCol + 4, tableRows, 1);
  const optionsPctRange = sheet.getRange(startRow, startCol + 5, tableRows, 1);

  const vlineRange = hasCurrentPrice ? sheet.getRange(vlineRow + 1, vlineCol, 2, 2) : null;

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
    .setOption("vAxis", { title: "P/L ($)" })
    .setOption("legend", { position: "right" })
    .setOption("curveType", "none");

  // Ensure series labels come from header row:
  // With addRange(contiguous), Sheets uses the first row as headers automatically if present.
  // We additionally set labelInLegend explicitly to headerRow names (best-effort).
  const dollarSeries = {
    0: { labelInLegend: headerRow[1] },
    1: { labelInLegend: headerRow[2] },
    2: { labelInLegend: headerRow[3] },
  };

  if (vlineRange) {
    dollarBuilder = dollarBuilder.addRange(vlineRange);
    dollarSeries[3] = { labelInLegend: "Current Price", lineDashStyle: [6, 4], lineWidth: 2 };
  }

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

  // Create / rebuild while preserving chart placement
  // Default positions if missing:
  //   $ chart: top-left
  //   % chart: below it
  rebuildChartPreserveBox_(dollarChart, dollarBuilder, 1, 1);
  rebuildChartPreserveBox_(pctChart, pctBuilder, 15, 1);
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
   Config table – now in K:L, default tableStartRow = 50
   Includes descriptions/labels in-column K.
   ========================================================= */

function ensureAndReadConfig_(ss, sheet, symbol) {
  const defaults = {
    minPrice: 350,
    maxPrice: 900,
    step: 5,
    tableStartRow: 50, // default is 50
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
   Portfolios sheet & validation
   ========================================================= */

function ensureAndReadPortfolios_(ss) {
  let sh = ss.getSheetByName("Portfolios");
  if (!sh) sh = ss.insertSheet("Portfolios");

  const want = ["Symbol", "Type", "RangeName"];
  const existingHeader = sh.getRange(1, 1, 1, 3).getValues()[0];
  const headerOk = want.every((h, i) => normKey_(existingHeader[i]) === normKey_(h));
  if (!headerOk) {
    sh.getRange(1, 1, 1, 3).setValues([want]).setFontWeight("bold");
    sh.autoResizeColumns(1, 3);
  }

  const lastRow = Math.max(1, sh.getLastRow());
  const rng = sh.getRange(1, 1, lastRow, 3);
  const nr = ss.getNamedRanges().find(x => x.getName() === "Portfolios");
  if (!nr) {
    ss.setNamedRange("Portfolios", rng);
  } else if (!rangesEqual_(nr.getRange(), rng)) {
    nr.remove();
    ss.setNamedRange("Portfolios", rng);
  }

  const values = rng.getValues();
  if (values.length < 2) return {};

  validatePortfoliosTableOrThrow_(values);

  const h = values[0].map(normKey_);
  const iSym = findCol_(h, ["symbol"]);
  const iType = findCol_(h, ["type"]);
  const iRange = findCol_(h, ["rangename"]);
  if ([iSym, iType, iRange].some(i => i < 0)) return {};

  const out = {};
  for (let r = 1; r < values.length; r++) {
    const sym = String(values[r][iSym] ?? "").trim().toUpperCase();
    const typ = normalizePortfolioType_(values[r][iType]);
    const rangeName = String(values[r][iRange] ?? "").trim();

    if (!sym || !typ || !rangeName) continue;

    out[sym] = out[sym] || [];
    out[sym].push({ type: typ, rangeName });
  }

  return out;
}

function normalizePortfolioType_(raw) {
  if (raw == null) return null;
  const t = String(raw)
    .normalize("NFKD")
    .toLowerCase()
    .replace(/[\u2010-\u2015\u2212]/g, "-")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (/^(stock|stocks|share|shares)$/.test(t)) return "stock";
  if (/^(bcs|bull[\s.\-]?call[\s.\-]?spread(s)?)$/.test(t)) return "bull-call-spread";
  if (/^(bps|bull[\s.\-]?put[\s.\-]?spread(s)?)$/.test(t)) return "bull-put-spread";
  return null;
}

function validatePortfoliosTableOrThrow_(rows) {
  if (!rows || rows.length < 2) return;

  const header = rows[0].map(normKey_);
  const iSym = findCol_(header, ["symbol"]);
  const iType = findCol_(header, ["type"]);
  const iRange = findCol_(header, ["rangename"]);
  if (iSym < 0 || iType < 0 || iRange < 0) return;

  const errors = [];
  for (let r = 1; r < rows.length; r++) {
    const rowNum = r + 1;
    const sym = String(rows[r][iSym] || "").trim().toUpperCase();
    const typ = normalizePortfolioType_(rows[r][iType]);
    const rangeName = String(rows[r][iRange] || "").trim();

    if (!sym) errors.push(`Row ${rowNum}: Symbol is blank`);
    else if (!/^[A-Z0-9.-]+$/.test(sym)) errors.push(`Row ${rowNum}: Invalid symbol "${rows[r][iSym]}"`);

    if (!typ) errors.push(`Row ${rowNum}: Invalid Type "${rows[r][iType]}"`);

    if (!rangeName) errors.push(`Row ${rowNum}: RangeName is blank`);
  }

  if (errors.length > 0) {
    showPortfolioValidationDialog_(errors);
    throw new Error("Invalid Portfolios table");
  }
}

function showPortfolioValidationDialog_(errors) {
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial; font-size:13px">
       <h3>Invalid Portfolios Table</h3>
       <p>Please fix the following issues:</p>
       <ul>${errors.map(e => `<li><pre style="margin:0;white-space:pre-wrap">${escapeHtml_(e)}</pre></li>`).join("")}</ul>
       <p><b>Allowed Types:</b><br>
          stock (or shares/stocks)<br>
          bull-call-spread (or bull-call-spreads / BCS)<br>
          bull-put-spread (or bull-put-spreads / BPS)
       </p>
     </div>`
  )
    .setWidth(680)
    .setHeight(520);

  SpreadsheetApp.getUi().showModalDialog(html, "Portfolio Validation Error");
}

function escapeHtml_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/* =========================================================
   Named range resolution + UX messaging
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

function buildMissingRangeMessage_(ss, symbol, missingItems, stockCount, spreadCount) {
  const lines = [`Missing inputs for ${symbol}:`, `- Missing named ranges:`];
  const tableHints = [];

  for (const item of missingItems || []) {
    lines.push(`  • ${item.type} -> ${item.rangeName}`);
    if (tableExistsByStructuredRefProbe_(ss, item.rangeName)) tableHints.push(item.rangeName);
  }

  if (stockCount === 0 && spreadCount === 0) lines.push(`- No valid input ranges found.`);
  lines.push("");
  lines.push("Fix Portfolios sheet and create named ranges, then rerun PlotPortfolioValueByPrice.");

  if (tableHints.length > 0) {
    lines.push("");
    for (const t of tableHints) {
      lines.push(`I can see a Sheets TABLE named "${t}", but Apps Script can’t read tables directly.`);
      lines.push(`Please create a named range "${t}Table" that contains the table’s data.`);
      lines.push("");
    }
  }

  return lines.join("\n");
}

function tableExistsByStructuredRefProbe_(ss, tableNameRaw) {
  const name = String(tableNameRaw || "").trim();
  if (!name) return false;

  const scratch = ss.getSheetByName("Portfolios") || ss.getActiveSheet();
  const cell = scratch.getRange("ZZ1");

  const ref = tableRefName_(name);
  cell.setFormula(`=IFERROR(ROWS(${ref}[#ALL]),"")`);
  SpreadsheetApp.flush();

  const v = String(cell.getDisplayValue() || "").trim();
  cell.clearContent();

  return v !== "";
}

function tableRefName_(name) {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  const escaped = name.replace(/'/g, "''");
  return `'${escaped}'`;
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
  const idxCurrentPrice = findCol_(headerNorm, ["currentprice", "marketprice", "mark", "last", "price"]);

  const out = [];

  for (let r = 1; r < rows.length; r++) {
    if (idxSym >= 0) {
      const rowSym = String(rows[r][idxSym] ?? "").trim().toUpperCase();
      if (rowSym && rowSym !== symbol) continue;
    }

    if (idxCurrentPrice >= 0 && outMeta && !isFinite(outMeta.currentPrice)) {
      const cp = toNum_(rows[r][idxCurrentPrice]);
      if (isFinite(cp)) outMeta.currentPrice = cp;
    }

    if (idxQty < 0 || idxBasis < 0) continue;

    const qty = toNum_(rows[r][idxQty]);
    const basis = toNum_(rows[r][idxBasis]);

    if (!isFinite(qty) || qty === 0) continue;
    if (!isFinite(basis)) continue;

    out.push({ qty, basis });
  }

  return out;
}

/**
 * Spread parsing:
 * - Counts ONLY “definition rows” that contain BOTH Long Strike and Short Strike
 * - Ignores fill-detail rows automatically
 * - Uses debit cost preference:
 *     Ave Debit / Avg Debit / Average Debit
 *     else Rec Debit / Recommended Debit
 *     else Net Debit / Debit / Cost / Entry / Price
 * - If contracts column missing, defaults qty=1
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
    if (!isFinite(kLong) || !isFinite(kShort)) continue;

    const qty = assumeQty ? 1 : toNum_(rows[r][idxQty]);
    if (!isFinite(qty) || qty === 0) continue;

    let debit = NaN;
    if (idxAveDebit >= 0) debit = toNum_(rows[r][idxAveDebit]);
    if (!isFinite(debit) && idxRecDebit >= 0) debit = toNum_(rows[r][idxRecDebit]);
    if (!isFinite(debit) && idxDebitFallback >= 0) debit = toNum_(rows[r][idxDebitFallback]);
    if (!isFinite(debit)) continue;

    out.push({ qty, kLong, kShort, debit, flavor });
  }

  return out;
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

  if (!isFinite(n)) return NaN;
  return neg ? -n : n;
}

function numOr_(v, fallback) {
  const n = toNum_(v);
  return isFinite(n) ? n : fallback;
}

function clamp_(x, lo, hi) {
  return Math.max(lo, Math.min(hi, x));
}

function round4_(x) {
  return Math.round(x * 10000) / 10000;
}
