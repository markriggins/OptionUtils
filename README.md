# OptionUtils

Option portfolio analysis tools for Google Sheets. SpreadFinder scans option prices to find attractive bull call spread opportunities. Portfolio tools let you track positions and visualize profit/loss scenarios.

**[Try it now — make a copy of the template spreadsheet](https://docs.google.com/spreadsheets/d/1MYMzTpoGlKuAXFyN8eGePmAsZ8R6379zhDoe-JHm7D0/copy)**

The template includes real TSLA option prices for June 2028 and December 2028 LEAP expirations.

## Understanding the Charts
![Strike vs ExpROI Chart](images/strike-vs-exproi-chart.png)
Each bubble represents a potential **bull call spread** you could purchase. The bubbles are color-coded by fitness score percentile to help you quickly identify the best opportunities:

- **Dark Green (large)** — Top 10% — Best opportunities with favorable liquidity and tight bid/ask spreads
- **Light Green** — Top 10–40% — Good opportunities
- **Yellow** — Top 40–75% — Moderate opportunities
- **Red** — Bottom 25% — Lower quality (poor liquidity or wide spreads)
- **Gray** — Conflicts with existing holdings in your portfolio (you cannot be long and short the same option)



Expected ROI is based on an estimated probability that the stock price will reach the upper strike before expiration. This accounts for: (1) the chance the spread expires fully in-the-money at max profit, or (2) the price touches the upper strike earlier, allowing you to exit early yet still capture ~80% of gains

Selecting a Spread

Click any bubble to see its details in the panel below, including strike prices, debit cost, ROI, expected ROI, and liquidity metrics.

![Spread Details Panel](images/spread-details-panel.png)


Double-click a bubble to open it in OptionStrat for visual profit/loss analysis.

![OptionStrat Profit/Loss](images/optionstrat-profit-loss.png)

You can also upload your portfolio and transactions from Etrade, or play around with a Sample portfolio that is built into the app

![Portfolio Spreadsheet](images/portfolio-spreadsheet.png)

If you upload transactions, it automatically detects bull-call-spread, iron-condors and iron-butterflies.   Once your portfolio has been built, you can visualize potential results as TSLA rises (we hope!)

![Portfolio Value Chart](images/portfolio-value-chart.png)
You can upload fresh option prices as the market changes and redraw these graph with the latest data.

This is beta code, so please keep your expectations low.  But I wrote it for myself to help me chose the best options and monitor my spreads, found it useful and wanted to share with others.



---

## Installation

1. Open the [OptionUtils Template Spreadsheet](https://docs.google.com/spreadsheets/d/1MYMzTpoGlKuAXFyN8eGePmAsZ8R6379zhDoe-JHm7D0/copy)
2. Click "File > Make a copy"
3. The Apps Script will be automatically installed
4. Wait for the **OptionTools** menu to appear
5. Run **OptionTools > Initialize / Clear Project** to set up the README sheet

You'll be prompted to authorize the app. The app only accesses your spreadsheet — no Drive access is required. Source code: https://github.com/markriggins/OptionUtils

---

## How It Works

The library architecture means:
- All logic lives in the SpreadFinder library
- Your spreadsheet has thin wrapper stubs that delegate to the library
- When the library is updated, you automatically get the latest version (after reload)

### Custom Functions

The following functions are defined for use in Google Sheets:

| Function | Description |
|----------|-------------|
| `XLookupByKeys(keys, keyHeaders, returnHeaders, sheet)` | Multi-key lookup with caching |
| `X2LOOKUP(key1, key2, col1, col2, returnCol)` | Two-key lookup |
| `X3LOOKUP(key1, key2, key3, col1, col2, col3, returnCol)` | Three-key lookup |
| `detectStrategy(strikes, types, qtys)` | Detects option strategy from legs |
| `recommendClose(symbol, exp, strike, type, qty, patience)` | Recommended closing price |
| `coalesce(range)` | First non-empty value |

---

## Data Sources

### Option Prices from Barchart.com

1. Go to barchart.com, navigate to Options for your symbol
2. Select expiration, choose "Stacked" view
3. Download CSV (filename format: `symbol-options-exp-YYYY-MM-DD-....csv`)
4. Run **OptionTools > SpreadFinder > Upload Option Prices**
5. Select your CSV file(s) and click Upload

### Portfolio and Transactions from E*Trade

1. Download transaction history CSV from E*Trade (Accounts > Transactions > Download)
2. Optionally download portfolio CSV (Accounts > Portfolio > Download)
3. Run **OptionTools > Portfolio > Upload Portfolio/Transactions**
4. Select your files and choose "Add transactions" or "Clear and rebuild"

## Supported Positions
- Stock positions
- Bull call spreads
- Bull put spreads
- Iron condors
- Iron butterflies
- Single-leg options

---

## License

MIT License.

Use it, fork it, adapt it, improve it.

---

## Author

**Mark Riggins**

Built to support real-world option portfolios where Google Sheets can be a convenient tool for modeling.
