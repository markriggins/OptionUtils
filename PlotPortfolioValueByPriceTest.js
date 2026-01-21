/**
 * PlotPortfolioValueByPrice.js
 * ------------------------------------------------------------
 * PlotPortfolioValueByPrice
 *
 * Reads a "Portfolios" table describing which inputs belong to which symbol, then creates/refreshes
 * a "<SYMBOL>PortfolioValueByPrice" tab for EACH symbol.
 *
 * For each symbol tab:
 *   - Creates/maintains a per-symbol Config table at H1:I9
 *   - Writes a generated data table with both $ and % series:
 *       Price | Shares $ | Options $ | Total $ | Shares % | Options %
 *   - Ensures TWO charts exist (and refreshes them on every run while preserving position best-effort):
 *       Chart 1: $ P/L vs Price (Shares $, Options $, Total $) + vertical dashed line at Current Price (if available)
 *       Chart 2: % Return vs Price (Shares %, Options %)
 *   - If the stock table contains a "Current Price" column, draws a VERTICAL DASHED LINE at that price on the $ chart.
 *
 * Portfolios table:
 *   - Sheet: "Portfolios" (created if missing)
 *   - Columns: Symbol | Type | RangeName
 *   - Type canonical singular:
 *       stock
 *       bull-call-spread  (aliases: bull-call-spreads, BCS)
 *       bull-put-spread   (aliases: bull-put-spreads, BPS)
 *
 * RangeName:
 *   - Must be a Named Range.
 *   - Convenience convention: if RangeName isn't found, the script tries RangeName + "Table"
 *     (e.g., Sheets table "Stocks" paired with named range "StocksTable").
 *
 * Input tables:
 *   - Header matching is case/space-insensitive.
 *   - If an input table includes a "Symbol" or "Ticker" column, rows not matching the desired symbol are ignored.
 *
 * Spread tables special handling:
 *   - Only rows with BOTH Long Strike and Short Strike are treated as spread definition rows
 *   - Fill-detail rows (missing strikes) are ignored automatically
 *   - Debit column preference:
 *       Ave Debit / Avg Debit / Average Debit
 *       else Rec Debit / Recommended Debit
 *       else Net Debit / Debit / Cost / Entry / Price
 *
 * ROI (% Return) definitions:
 *   - Shares % = Shares $ / Total Shares Cost Basis
 *   - Options % = Options $ / (Sum(Call debits paid) + Sum(Put spread max risk))
 *     For bull put spreads: credit is stored as NEGATIVE debit. Max risk per contract = (width - credit) * 100.
 *
 * Config tables:
 *   - Named ranges are global, so each symbol has a unique config named range: Config_<SYMBOL>
 *   - Config columns H:I are hidden AFTER a successful build, visible if inputs are missing or parsing yields nothing.
 *
 * Auto-refresh:
 *   - onEdit(e) rebuilds only when edits intersect any Config_<SYMBOL> named range.
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

/**
 * Simple trigger: rebuild only when edits intersect any Config_<SYMBOL> named range.
 */
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
    // swallow to avoid noisy errors on edit
  }
}

/* =========================================================
   Per-symbol builder
   ========================================================= */

function plotForSymbol_(ss, symbolRaw, entries) {
  const symbol = String(symbolRaw || "").trim().toUpperCase();
  if (!symbol) return;

  const sheetName = `${symbol}PortfolioValueByPrice`;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.showSheet();

  // Ensure config exists and read it (kept visible unless we succeed)
  const cfg = ensureAndReadConfig_(ss, sheet, symbol);

  if (!entries || entries.length === 0) {
    writeStatus_(sheet, `No portfolio inputs found for ${symbol} in the Portfolios table.`);
    showConfig_(sheet);
    return;
  }

  const stockRanges = [];
  const callSpreadRanges = [];
  const putSpreadRanges = [];
  const missingItems = []; // [{type, rangeName}]

  for (const e of entries) {
    const rangeNameRaw = String(e.rangeName ?? "").trim();
    const type = String(e.type ?? "").trim(); // canonical singular values

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
    showConfig_(sheet);
    return;
  }

  clearStatus_(sheet);

  // Parse inputs
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

  // If we parsed nothing, say so clearly (helps debugging)
  if (shares.length === 0 && callSpreads.length === 0 && putSpreads.length === 0) {
    writeStatus_(
      sheet,
      `Parsed 0 rows for ${symbol}.\n` +
        `This usually means your headers don't match expected names, or numeric cells are non-numeric.\n\n` +
        `Try:\n` +
        `- Ensure stock table has Shares/Qty + AvgCost/Basis/Ave Price (Paid) columns\n` +
        `- Ensure spread tables have Long Strike + Short Strike + Contracts + Ave Debit (or Rec Debit)\n` +
        `- Ensure numbers aren't stored as text with commas/extra characters`
    );
    showConfig_(sheet);
    return;
  }

  // --- ROI denominators ---
  const sharesCost = shares.reduce((sum, sh) => sum + sh.qty * sh.basis, 0);
  const callInvest = callSpreads.reduce((sum, sp) => sum + (sp.debit * 100 * sp.qty), 0);
  const putRisk = putSpreads.reduce((sum, sp) => {
    const width = sp.kShort - sp.kLong;
    if (width <= 0) return sum;
    const credit = -sp.debit; // debit stored negative for credit spreads
    const maxLossPer = (width - credit) * 100;
    return sum + maxLossPer * sp.qty;
  }, 0);
  const optionsDenom = callInvest + putRisk;

  // ---------- Build output table ----------
  const table = [["Price", "Shares $", "Options $", "Total $", "Shares %", "Options %"]];

  for (let S = cfg.minPrice; S <= cfg.maxPrice + 1e-9; S += cfg.step) {
    let sharesPL = 0;
    let spreadsPL = 0;

    // Shares P/L at scenario price S
    for (const sh of shares) sharesPL += (S - sh.basis) * sh.qty;

    // CALL spreads: payoff at expiration (intrinsic)
    for (const sp of callSpreads) {
      const width = sp.kShort - sp.kLong;
      if (width <= 0) continue;
      const intrinsic = clamp_(S - sp.kLong, 0, width);
      spreadsPL += (intrinsic - sp.debit) * 100 * sp.qty;
    }

    // PUT spreads (bull put spread credit convention: debit is negative credit)
    for (const sp of putSpreads) {
      const width = sp.kShort - sp.kLong;
      if (width <= 0) continue;
      const loss = clamp_(sp.kShort - S, 0, width);
      spreadsPL += (-sp.debit - loss) * 100 * sp.qty;
    }

    const totalPL = sharesPL + spreadsPL;
    const sharesPct = sharesCost > 0 ? (sharesPL / sharesCost) : 0;
    const optionsPct = optionsDenom > 0 ? (spreadsPL / optionsDenom) : 0;

    table.push([
      round2_(S),
      round2_(sharesPL),
      round2_(spreadsPL),
      round2_(totalPL),
      round4_(sharesPct),
      round4_(optionsPct)
    ]);
  }

  // Compute Y bounds for vertical current-price line (use Total $)
  const totalYs = table
    .slice(1)
    .map(r => toNum_(r[3]))
    .filter(n => isFinite(n));
  const minY = totalYs.length ? Math.min(...totalYs) : 0;
  const maxY = totalYs.length ? Math.max(...totalYs) : 0;

  // ---------- Layout (table below charts) ----------
  const startRow = cfg.tableStartRow;
  const startCol = cfg.tableStartCol;

  // Clear data area only (preserve charts + config)
  sheet.getRange(startRow - 1, startCol, 2000, 30).clearContent();

  sheet.getRange(startRow - 1, startCol).setValue("Data (generated)").setFontWeight("bold");
  sheet.getRange(startRow, startCol, table.length, table[0].length).setValues(table);
  sheet.autoResizeColumns(startCol, table[0].length);

  // Format % columns as percent
  if (table.length > 1) {
    sheet.getRange(startRow + 1, startCol + 4, table.length - 1, 2).setNumberFormat("0.00%");
  }

  // Write vertical-line helper table (off to the right)
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

  // ---------- Charts (ensure both exist; refresh each run preserving placement best-effort) ----------
  ensureOrRefreshCharts_(
    sheet,
    symbol,
    cfg,
    startRow,
    startCol,
    table.length,
    hasCurrentPrice,
    vlineRow,
    vlineCol
  );

  // Hide config after successful build
  hideConfig_(sheet);
}

/**
 * Ensures BOTH charts exist and refreshes them on every run.
 * Preserves position (anchor row/col and offsets) best-effort by rebuilding charts at the same container.
 * NOTE: Apps Script does not expose explicit width/height, but anchor+offset is preserved.
 */
function ensureOrRefreshCharts_(sheet, symbol, cfg, startRow, startCol, nRows, hasCurrentPrice, vlineRow, vlineCol) {
  const charts = sheet.getCharts();

  // Identify by title convention
  const dollarTitle = (cfg.chartTitle || `${symbol} Portfolio Value by Price`) + " ($)";
  const pctTitle = (cfg.chartTitle || `${symbol} Portfolio Value by Price`) + " (%)";

  let dollarChart = null;
  let pctChart = null;

  for (const ch of charts) {
    const t = getChartTitle_(ch);
    if (t === dollarTitle) dollarChart = ch;
    if (t === pctTitle) pctChart = ch;
  }

  // Build ranges
  const dollarRange = sheet.getRange(startRow, startCol, nRows, 4); // Price + 3 $ series
  const priceColRange = sheet.getRange(startRow, startCol, nRows, 1);
  const sharesPctRange = sheet.getRange(startRow, startCol + 4, nRows, 1);
  const optionsPctRange = sheet.getRange(startRow, startCol + 5, nRows, 1);

  const vlineRange = hasCurrentPrice ? sheet.getRange(vlineRow + 1, vlineCol, 2, 2) : null;

  // Helper: remove+insert chart preserving anchor/offset
  function rebuildChartPreserveBox_(oldChart, builder) {
    const ci = oldChart.getContainerInfo();
    const anchorRow = ci.getAnchorRow();
    const anchorCol = ci.getAnchorColumn();
    const offsetX = ci.getOffsetX();
    const offsetY = ci.getOffsetY();

    sheet.removeChart(oldChart);
    const built = builder.setPosition(anchorRow, anchorCol, offsetX, offsetY).build();
    sheet.insertChart(built);
  }

  // ----- Chart 1: $ P/L -----
  let dollarBuilder = sheet
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dollarRange)
    .setOption("title", dollarTitle)
    .setOption("hAxis", { title: `${symbol} Price` })
    .setOption("vAxis", { title: "P/L ($)" })
    .setOption("legend", { position: "right" })
    .setOption("curveType", "none");

  if (vlineRange) {
    dollarBuilder = dollarBuilder.addRange(vlineRange);
    // Series indices: 0 Shares$, 1 Options$, 2 Total$, 3 vline
    dollarBuilder = dollarBuilder.setOption("series", {
      3: { lineDashStyle: [6, 4], lineWidth: 2 }
    });
  }

  if (!dollarChart) {
    // Default top placement
    sheet.insertChart(dollarBuilder.setPosition(1, 1, 0, 0).build());
  } else {
    rebuildChartPreserveBox_(dollarChart, dollarBuilder);
  }

  // ----- Chart 2: % Return -----
  let pctBuilder = sheet
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(priceColRange)
    .addRange(sharesPctRange)
    .addRange(optionsPctRange)
    .setOption("title", pctTitle)
    .setOption("hAxis", { title: `${symbol} Price` })
    .setOption("vAxis", { title: "% Return", format: "percent" })
    .setOption("legend", { position: "right" })
    .setOption("curveType", "none");

  if (!pctChart) {
    // Default below the first chart
    sheet.insertChart(pctBuilder.setPosition(15, 1, 0, 0).build());
  } else {
    rebuildChartPreserveBox_(pctChart, pctBuilder);
  }
}

/**
 * Best-effort chart title getter. The EmbeddedChart API does not provide a reliable direct title accessor,
 * so we read options when available.
 */
function getChartTitle_(chart) {
  try {
    const opts = chart.getOptions && chart.getOptions();
    if (opts && typeof opts.get === "function") {
      const t = opts.get("title");
      if (t != null) return String(t);
    }
    if (opts && opts.title != null) return String(opts.title);
  } catch (e) {}
  return "";
}

/* =========================================================
   Portfolios table (creates if missing) + validation
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

  // Maintain named range "Portfolios" for the A:C range we use
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

  const out = {}; // symbol -> [{type, rangeName}]
  for (let r = 1; r < values.length; r++) {
    const sym = String(values[r][iSym] ?? "").trim().toUpperCase();
    const typ = normalizePortfolioType_(values[r][iType]); // canonical singular
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
    const rawSym = rows[r][iSym];
    const rawType = rows[r][iType];
    const rawRange = rows[r][iRange];

    const sym = String(rawSym || "").trim().toUpperCase();
    const typ = normalizePortfolioType_(rawType);
    const rangeName = String(rawRange || "").trim();

    if (!sym) errors.push(`Row ${rowNum}: Symbol is blank`);
    else if (!/^[A-Z0-9.-]+$/.test(sym)) errors.push(`Row ${rowNum}: Invalid symbol "${rawSym}"`);

    if (!typ) errors.push(`Row ${rowNum}: Invalid Type "${rawType}"`);

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
       <p>Please fix the following issues before running the tool:</p>
       <ul>${errors.map(e => `<li><pre style="margin:0;white-space:pre-wrap">${escapeHtml_(e)}</pre></li>`).join("")}</ul>
       <p><b>Allowed Types (singular):</b><br>
          stock (or stocks/shares)<br>
          bull-call-spread (or bull-call-spreads / BCS)<br>
          bull-put-spread (or bull-put-spreads / BPS)
       </p>
     </div>`
  ).setWidth(700).setHeight(520);

  SpreadsheetApp.getUi().showModalDialog(html, "Portfolio Validation Error");
}

function escapeHtml_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/* =========================================================
   Config per symbol sheet: table at H1:I9, named range Config_<SYMBOL>
   ========================================================= */

function ensureAndReadConfig_(ss, sheet, symbol) {
  const defaults = {
    minPrice: 350,
    maxPrice: 900,
    step: 5,
    tableStartRow: 25,
    tableStartCol: 1,
    includeComponentSeries: true,
    chartTitle: `${symbol} Portfolio Value by Price`,
  };

  const cfgRow = 1;
  const cfgCol = 8; // H
  const cfgNumRows = 9;
  const cfgNumCols = 2;

  const values = [
    ["ConfigKey", "ConfigValue"],
    ["minPrice", defaults.minPrice],
    ["maxPrice", defaults.maxPrice],
    ["step", defaults.step],
    ["tableStartRow", defaults.tableStartRow],
    ["tableStartCol", defaults.tableStartCol],
    ["includeComponentSeries", defaults.includeComponentSeries],
    ["chartTitle", defaults.chartTitle],
    ["(unhide columns H:I to edit)", ""],
  ];

  const header = sheet.getRange(cfgRow, cfgCol).getValue();
  if (String(header).trim() !== "ConfigKey") {
    sheet.getRange(cfgRow, cfgCol, cfgNumRows, cfgNumCols).setValues(values);
    sheet.getRange(cfgRow, cfgCol, 1, cfgNumCols).setFontWeight("bold");
  }

  const cfgRange = sheet.getRange(cfgRow, cfgCol, cfgNumRows, cfgNumCols);

  const cfgName = `Config_${symbol}`;
  const existing = ss.getNamedRanges().find(nr => nr.getName() === cfgName);
  if (!existing) {
    ss.setNamedRange(cfgName, cfgRange);
  } else if (!rangesEqual_(existing.getRange(), cfgRange)) {
    existing.remove();
    ss.setNamedRange(cfgName, cfgRange);
  }

  const kv = sheet.getRange(cfgRow + 1, cfgCol, cfgNumRows - 1, 2).getValues();
  const map = {};
  for (const r of kv) {
    const k = String(r[0] ?? "").trim();
    if (!k || k.startsWith("(")) continue;
    map[k] = r[1];
  }

  const cfg = {
    minPrice: numOr_(map.minPrice, defaults.minPrice),
    maxPrice: numOr_(map.maxPrice, defaults.maxPrice),
    step: numOr_(map.step, defaults.step),
    tableStartRow: Math.max(5, Math.floor(numOr_(map.tableStartRow, defaults.tableStartRow))),
    tableStartCol: Math.max(1, Math.floor(numOr_(map.tableStartCol, defaults.tableStartCol))),
    includeComponentSeries: boolOr_(map.includeComponentSeries, defaults.includeComponentSeries),
    chartTitle: String(map.chartTitle ?? defaults.chartTitle),
  };

  if (!(cfg.minPrice < cfg.maxPrice)) {
    cfg.minPrice = defaults.minPrice;
    cfg.maxPrice = defaults.maxPrice;
  }
  if (!(cfg.step > 0)) cfg.step = defaults.step;

  return cfg;
}

function hideConfig_(sheet) {
  sheet.hideColumns(8, 2); // H:I
}
function showConfig_(sheet) {
  sheet.showColumns(8, 2); // H:I
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

function buildMissingRangeMessage_(ss, symbol, missingItems, stockRangeCount, spreadRangeCount) {
  const lines = [];
  lines.push(`Missing inputs for ${symbol}:`);
  lines.push(`- Missing named ranges:`);

  for (const item of (missingItems || [])) {
    lines.push(`  • ${item.type} -> ${item.rangeName}`);
  }

  if (stockRangeCount === 0 && spreadRangeCount === 0) {
    lines.push(`- No valid input ranges found for this symbol.`);
  }

  lines.push("");
  lines.push("Fix Portfolios (Symbol | Type | RangeName) and create the named ranges, then rerun PlotPortfolioValueByPrice.");

  return lines.join("\n");
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

  // For meta.currentPrice line:
  // Prefer "Current Price" explicitly; avoid accidentally using a generic "Price Paid" column.
  const idxCurrentPrice = findCol_(headerNorm, ["currentprice", "marketprice", "mark", "last"]);

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

function parseSpreadsFromTableForSymbol_(rows, symbol, flavor /* "CALL" or "PUT" */) {
  if (!rows || rows.length < 2) return [];

  const headerNorm = rows[0].map(normKey_);
  const idxSym = findCol_(headerNorm, ["symbol", "ticker"]);

  const idxQty = findCol_(headerNorm, ["contracts", "contract", "qty", "quantity", "count", "numcontracts", "spreads", "spreadqty"]);

  const idxLong = findCol_(headerNorm, ["lower", "lowerstrike", "long", "longstrike", "buystrike", "strikebuy", "strikelong"]);
  const idxShort = findCol_(headerNorm, ["upper", "upperstrike", "short", "shortstrike", "sellstrike", "strikesell", "strikeshort"]);

  const idxAveDebit = findCol_(headerNorm, ["avedebit", "avgdebit", "averagedebit"]);
  const idxRecDebit = findCol_(headerNorm, ["recdebit", "recommendeddebit"]);
  const idxDebitFallback = findCol_(headerNorm, ["netdebit", "debit", "cost", "entry", "price"]);

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

    // Only definition rows have both strikes (auto-ignores fill detail rows)
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

  const aR1 = a.getRow(), aC1 = a.getColumn();
  const aR2 = aR1 + a.getNumRows() - 1;
  const aC2 = aC1 + a.getNumColumns() - 1;

  const bR1 = b.getRow(), bC1 = b.getColumn();
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

function findCol_(normalizedHeaderRow, synonyms) {
  const syn = synonyms.map(normKey_);
  for (let i = 0; i < normalizedHeaderRow.length; i++) {
    if (syn.includes(normalizedHeaderRow[i])) return i;
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

  const isParenNeg = /^\(.*\)$/.test(s);
  if (isParenNeg) s = s.slice(1, -1);

  s = s.replace(/[$,%]/g, "").replace(/,/g, "").trim();

  const n = Number(s);
  if (!isFinite(n)) return NaN;
  return isParenNeg ? -n : n;
}

function numOr_(v, fallback) {
  const n = toNum_(v);
  return isFinite(n) ? n : fallback;
}

function boolOr_(v, fallback) {
  if (typeof v === "boolean") return v;
  const s = String(v ?? "").trim().toLowerCase();
  if (["true", "yes", "y", "1"].includes(s)) return true;
  if (["false", "no", "n", "0"].includes(s)) return false;
  return fallback;
}

function clamp_(x, lo, hi) {
  return Math.max(lo, Math.min(hi, x));
}

function round2_(x) {
  return Math.round(x * 100) / 100;
}

function round4_(x) {
  return Math.round(x * 10000) / 10000;
}
