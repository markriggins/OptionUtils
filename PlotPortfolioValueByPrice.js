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
 *   - Writes a generated data table (Price vs Shares P/L vs Spreads P/L vs Total P/L)
 *   - Creates a chart at the top (created once; later runs preserve user resizing/position)
 *   - If the stock table contains a "Current Price" column, draws a VERTICAL DASHED LINE at that price.
 *     (Implemented as an added series of two points at X=currentPrice spanning [minY,maxY].)
 *
 * Portfolios table:
 *   - Sheet: "Portfolios" (created if missing)
 *   - Columns: Symbol | Type | RangeName
 *   - Type is FORCED to singular canonical values (case-insensitive, dash-normalized):
 *       stock
 *       bull-call-spread
 *       bull-put-spread
 *     Accepted aliases:
 *       shares -> stock
 *       bull-call-spreads -> bull-call-spread
 *       bull-put-spreads  -> bull-put-spread
 *       BCS -> bull-call-spread
 *       BPS -> bull-put-spread
 *
 * RangeName:
 *   - Must be a Named Range.
 *   - Convenience convention: if RangeName isn't found, the script tries RangeName + "Table"
 *     (e.g., Sheets Table "Stocks" paired with Named Range "StocksTable").
 *
 * Input tables:
 *   - Header matching is case/space-insensitive.
 *   - If an input table includes a "Symbol" or "Ticker" column, rows not matching the desired symbol
 *     are ignored.
 *
 * Spread tables special handling:
 *   Many users store one “summary” row per spread (has Long Strike, Short Strike, Contracts, Ave Debit),
 *   followed by “fill detail” rows that may have Contracts/Prices but do NOT repeat strikes.
 *   This script:
 *     - Treats ONLY rows that have BOTH Long Strike and Short Strike as actual spread definitions
 *     - Ignores fill-detail rows automatically
 *     - Uses "Ave Debit" (or "Avg Debit"/"Average Debit") as the cost to enter when present
 *       otherwise falls back to "Rec Debit", otherwise "Debit/Net Debit/Cost/Entry"
 *
 * Config tables:
 *   - Named ranges are global (not scoped per sheet), so each symbol has a unique config named range:
 *       Config_<SYMBOL> (e.g., Config_TSLA)
 *   - Config columns H:I are hidden AFTER a successful build, but kept visible if inputs are missing.
 *
 * Auto-refresh:
 *   - onEdit(e) rebuilds only when the edited cell intersects any Config_<SYMBOL> range.
 *
 * Chart sizing:
 *   - The chart is created only if none exists yet to preserve manual resizing/position.
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
 * (It is OK that onEdit fires for all user edits — we return immediately unless config was edited.)
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

  // Parse
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
        `- Ensure stock table has Shares/Qty + AvgCost/Basis/Ave Price columns\n` +
        `- Ensure spread tables have Long Strike + Short Strike + Contracts + Ave Debit (or Rec Debit)\n` +
        `- Ensure numbers aren't stored as text with commas/extra characters`
    );
    showConfig_(sheet);
    return;
  }

  // ---------- Build output table ----------
  const table = [["Price", "Shares", "Options", "Total"]]; // Updated headers for simpler labels

  for (let S = cfg.minPrice; S <= cfg.maxPrice + 1e-9; S += cfg.step) {
    let sharesPL = 0;
    let spreadsPL = 0;

    for (const sh of shares) sharesPL += (S - sh.basis) * sh.qty;

    // CALL spreads: payoff at expiration (intrinsic)
    for (const sp of callSpreads) {
      const width = sp.kShort - sp.kLong;
      if (width <= 0) continue;
      const intrinsic = clamp_(S - sp.kLong, 0, width);
      spreadsPL += (intrinsic - sp.debit) * 100 * sp.qty;
    }

    // PUT spreads (bull put spread credit convention):
    // Store 'Ave Debit' as NEGATIVE credit (e.g., -2.50) for credit spreads.
    // Then -debit = credit. Terminal loss component is clamp_(kShort - S, 0, width).
    for (const sp of putSpreads) {
      const width = sp.kShort - sp.kLong;
      if (width <= 0) continue;
      const loss = clamp_(sp.kShort - S, 0, width);
      spreadsPL += (-sp.debit - loss) * 100 * sp.qty;
    }

    table.push([round2_(S), round2_(sharesPL), round2_(spreadsPL), round2_(sharesPL + spreadsPL)]);
  }

  // Compute Y bounds for vertical current-price line
  const totalYs = table
    .slice(1)
    .map(r => toNum_(r[3]))
    .filter(n => isFinite(n));
  const minY = totalYs.length ? Math.min(...totalYs) : 0;
  const maxY = totalYs.length ? Math.max(...totalYs) : 0;

  // ---------- Layout ----------
  const startRow = cfg.tableStartRow;
  const startCol = cfg.tableStartCol;

  // Clear data area only (preserve chart + config)
  sheet.getRange(startRow - 1, startCol, 2000, 30).clearContent();

  sheet.getRange(startRow - 1, startCol).setValue("Data (generated)").setFontWeight("bold");
  sheet.getRange(startRow, startCol, table.length, table[0].length).setValues(table);
  sheet.autoResizeColumns(startCol, table[0].length);

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

  // ---------- Chart (create once only to preserve user sizing) ----------
  const charts = sheet.getCharts();
  if (charts.length === 0) {
    const dataRange = sheet.getRange(startRow, startCol, table.length, table[0].length);

    let chartBuilder = sheet
      .newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataRange)
      .setPosition(1, 1, 0, 0)
      .setOption("title", cfg.chartTitle || `${symbol} Portfolio Value by Price`)
      .setOption("hAxis", { title: `${symbol} Price` })
      .setOption("vAxis", { title: "Unrealized P/L" })
      .setOption("legend", { position: "right" });

    if (!cfg.includeComponentSeries) {
      const priceColRange = sheet.getRange(startRow, startCol, table.length, 1);
      const totalColRange = sheet.getRange(startRow, startCol + 3, table.length, 1);

      chartBuilder = sheet
        .newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(priceColRange) // X axis
        .addRange(totalColRange) // Total P/L series
        .setPosition(1, 1, 0, 0)
        .setOption("title", cfg.chartTitle || `${symbol} Portfolio Value by Price`)
        .setOption("hAxis", { title: `${symbol} Price` })
        .setOption("vAxis", { title: "Unrealized P/L" })
        .setOption("legend", { position: "right" });
    }

    // Add vertical line as an extra series if we have current price
    if (hasCurrentPrice) {
      const vlineRange = sheet.getRange(vlineRow + 1, vlineCol, 2, 2);
      chartBuilder = chartBuilder.addRange(vlineRange);
    }

    // Style the series with colors and labels (best-effort; Google Charts defaults may vary, but this sets them explicitly)
    const seriesOpt = {
      0: { color: 'blue', labelInLegend: 'Shares' }, // First series: Shares
      1: { color: 'red', labelInLegend: 'Options' }, // Second: Options (Spreads)
      2: { color: 'yellow', labelInLegend: 'Total' } // Third: Total
    };

    // Style the vertical line series as dashed (last series index). Best-effort.
    if (hasCurrentPrice) {
      // includeComponentSeries=true => 3 series => vline is 4th (index 3)
      // includeComponentSeries=false => Total => 1 series => vline is 2nd (index 1)
      const vlineSeriesIndex = cfg.includeComponentSeries ? 3 : 1;
      seriesOpt[vlineSeriesIndex] = { lineDashStyle: [6, 4], lineWidth: 2 };
    }

    chartBuilder = chartBuilder.setOption("series", seriesOpt).setOption("curveType", "none");

    sheet.insertChart(chartBuilder.build());
  }

  // Hide config after successful build
  hideConfig_(sheet);
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
    .normalize("NFKD")                     // Unicode canonicalization
    .toLowerCase()
    .replace(/[\u2010-\u2015\u2212]/g, "-") // dash variants → "-"
    .replace(/\u00a0/g, " ")               // nbsp → space
    .replace(/\s+/g, " ")                  // collapse whitespace
    .trim();

  // ---- STOCK ----
  if (/^(stock|stocks|share|shares)$/.test(t)) {
    return "stock";
  }

  // ---- BULL CALL SPREAD ----
  // Matches:
  //   bcs
  //   bull-call-spread
  //   bull call spreads
  //   bull.call.spread
  if (/^(bcs|bull[\s.\-]?call[\s.\-]?spread(s)?)$/.test(t)) {
    return "bull-call-spread";
  }

  // ---- BULL PUT SPREAD ----
  // Matches:
  //   bps
  //   bull-put-spread
  //   bull put spreads
  //   bull.put.spread
  if (/^(bps|bull[\s.\-]?put[\s.\-]?spread(s)?)$/.test(t)) {
    return "bull-put-spread";
  }

  return null;
}

function validatePortfoliosTableOrThrow_(rows) {
  if (!rows || rows.length < 2) return;

  const header = rows[0].map(normKey_);
  const iSym = findCol_(header, ["symbol"]);
  const iType = findCol_(header, ["type"]);

  if (iSym < 0 || iType < 0) return;

  const errors = [];

  for (let r = 1; r < rows.length; r++) {
    const rowNum = r + 1;

    const rawSym = rows[r][iSym];
    const rawType = rows[r][iType];

    Logger.log(debugChars_("Raw Type Before Norm:", rawType));

    const sym = String(rawSym || "").trim().toUpperCase();
    const typ = normalizePortfolioType_(rawType);

    if (!sym) {
      errors.push(`Row ${rowNum}: Symbol is blank`);
    } else if (!/^[A-Z0-9.-]+$/.test(sym)) {
      errors.push(`Row ${rowNum}: Invalid symbol "${rawSym}"`);
    }

    if (!typ) {
      errors.push(`Row ${rowNum}: Invalid Type ${debugChars_(rawType)}`);
      errors.push(`Row ${rowNum}: Invalid Type "${rawType}"`);
    }
  }

  if (errors.length > 0) {
    showPortfolioValidationDialog_(errors);
    throw new Error("Invalid Portfolios table");
  }
}

/**
 * Prints a string char-by-char with Unicode code points.
 * Useful to detect NBSP, zero-width, Unicode hyphens, etc.
 */
function debugChars_(label, s) {
  s = (s == null) ? "" : String(s);
  const parts = [];
  for (let i = 0; i < s.length; i++) {
    const ch = s.charAt(i);
    const cp = s.codePointAt(i);
    // If it's a surrogate pair, advance i an extra step
    if (cp > 0xffff) i++;
    parts.push(`${i}:${JSON.stringify(ch)} U+${cp.toString(16).toUpperCase().padStart(4, "0")}`);
  }
  return `${label} len=${s.length} -> ${parts.join(" | ")}`;
}


function showPortfolioValidationDialog_(errors) {
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial; font-size:13px">
       <h3>Invalid Portfolios Table</h3>
       <p>Please fix the following issues before running the tool:</p>
       <ul>${errors.map(e => `<li>${e}</li>`).join("")}</ul>
       <p><b>Allowed Types (singular):</b><br>
          stock<br>
          bull-call-spread (BCS)<br>
          bull-put-spread (BPS)
       </p>
     </div>`
  )
    .setWidth(520)
    .setHeight(380);

  SpreadsheetApp.getUi().showModalDialog(html, "Portfolio Validation Error");
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
    chartTitle: `${symbol} Portfolio Value (P/L) vs ${symbol} Price`,
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

  const tableHints = [];
  for (const item of (missingItems || [])) {
    lines.push(`  • ${item.type} -> ${item.rangeName}`);
    if (tableExistsByStructuredRefProbe_(ss, item.rangeName)) {
      tableHints.push(item.rangeName);
    }
  }

  if (stockRangeCount === 0 && spreadRangeCount === 0) {
    lines.push(`- No valid input ranges found for this symbol.`);
  }

  lines.push("");
  lines.push("Fix Portfolios (Symbol | Type | RangeName) and create the named ranges, then rerun PlotPortfolioValueByPrice.");

  if (tableHints.length > 0) {
    lines.push("");
    for (const t of tableHints) {
      lines.push(`I can see a Sheets TABLE named ${t},`);
      lines.push(`but Apps Script can’t read tables directly.`);
      lines.push(`Please create a range named "${t}Table"`);
      lines.push(`that contains the table’s data`);
      lines.push("");
    }
  }

  return lines.join("\n");
}

function tableExistsByStructuredRefProbe_(ss, tableNameRaw) {
  const name = String(tableNameRaw || "").trim();
  if (!name) return false;

  const scratchSheet = ss.getSheetByName("Portfolios") || ss.getActiveSheet();
  const cell = scratchSheet.getRange("ZZ1");

  const ref = tableRefName_(name);
  const formula = `=IFERROR(ROWS(${ref}[#ALL]),"")`;

  cell.setFormula(formula);
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
  if (!rows || rows.length < 2) return []; // No data rows

  const headers = rows[0].map(h => String(h || '').trim());

  // Flexible column lookup: qtyCol for Shares/Qty, basisCol for AvgCost/Basis/Ave Price
  const symbolCol = headers.findIndex(h => h.toUpperCase() === 'SYMBOL' || h.toUpperCase() === 'TICKER');
  const qtyCol = headers.findIndex(h => ['SHARES', 'QTY', 'QUANTITY'].includes(h.toUpperCase()));
  const basisCol = headers.findIndex(h => ['AVGCOST', 'BASIS', 'AVE PRICE', 'AVERAGE PRICE', 'COST BASIS'].includes(h.toUpperCase()));
  const priceCol = headers.findIndex(h => ['CURRENT PRICE', 'MARKET PRICE', 'PRICE'].includes(h.toUpperCase())); // For meta.currentPrice

  if (qtyCol === -1 || basisCol === -1) return []; // Missing required columns

  const shares = [];
  for (let i = 1; i < rows.length; i++) { // Skip header
    const row = rows[i];
    const rowSymbol = symbolCol !== -1 ? String(row[symbolCol] || '').trim().toUpperCase() : symbol; // Default to input symbol if no column

    if (rowSymbol !== symbol) continue; // Filter for this symbol

    const qty = parseFloat(row[qtyCol]);
    const basis = parseFloat(row[basisCol]);

    if (isNaN(qty) || qty <= 0 || isNaN(basis)) continue; // Invalid row

    // Update meta if current price available
    if (priceCol !== -1) {
      const currentPrice = parseFloat(row[priceCol]);
      if (!isNaN(currentPrice)) {
        outMeta.currentPrice = currentPrice;
        shares.push({ qty, currentPrice });
      } else {
        shares.push({ qty, basis });
      }
    } else {
     shares.push({ qty, basis });
    }
  }

  return shares;
}

/**
 * Spread parsing:
 * - Counts ONLY “definition rows” that contain BOTH Long Strike and Short Strike
 * - Ignores fill-detail rows automatically
 * - Uses debit cost preference:
 *     1) "Ave Debit" / "Avg Debit" / "Average Debit"
 *     2) "Rec Debit" / "Recommended Debit"
 *     3) "Net Debit" / "Debit" / "Cost" / "Entry"
 * - If contracts column missing, defaults qty=1
 */
function parseSpreadsFromTableForSymbol_(rows, symbol, flavor /* "CALL" or "PUT" */) {
  if (!rows || rows.length < 2) return [];

  const headerNorm = rows[0].map(normKey_);
  const idxSym = findCol_(headerNorm, ["symbol", "ticker"]);

  const idxQty = findCol_(headerNorm, ["contracts", "contract", "qty", "quantity", "count", "numcontracts", "spreads", "spreadqty"]);

  const idxLong = findCol_(headerNorm, ["lower", "lowerstrike", "long", "longstrike", "buystrike", "strikebuy", "strikelong"]);
  const idxShort = findCol_(headerNorm, ["upper", "upperstrike", "short", "shortstrike", "sellstrike", "strikesell", "strikeshort"]);

  const idxAveDebit = findCol_(headerNorm, ["avedebit", "avgdebit", "averagedebit", "avedevit"]);
  const idxRecDebit = findCol_(headerNorm, ["recdebit", "recommendeddebit"]);
  const idxDebitFallback = findCol_(headerNorm, ["netdebit", "debit", "cost", "price", "entry", "premiumpaid", "premium"]);

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

function findCol_(normalizedHeaderRow, synonyms) {
  const syn = synonyms.map(normKey_);
  for (let i = 0; i < normalizedHeaderRow.length; i++) {
    if (syn.includes(normalizedHeaderRow[i])) return i;
  }
  return -1;
}

function test_normKey() {
  let normKey1 = normKey_("average price");
  let normKey2 = normKey_("ave price");
  Logger.log("normKey1=" + normKey1)
  Logger.log("normKey2=" + normKey2)
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