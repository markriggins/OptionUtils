# OptionUtils / SpreadFinder

A Google Apps Script toolkit for modeling option portfolios in Google Sheets. Track stock positions and option spreads, calculate value at every price point, and generate profit/loss charts.

---

## Installation

### Option A: Copy the Template Spreadsheet (Recommended)

1. Open the [SpreadFinder Template Spreadsheet](https://docs.google.com/spreadsheets/d/YOUR_TEMPLATE_ID/copy)
2. Click "Make a copy"
3. The library and stubs are already configured — you're ready to go

### Option B: Manual Setup

If you have an existing spreadsheet or want to set up from scratch:

1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. Delete any existing `.gs` files
4. Create a new file called `Code.gs`
5. Copy the contents of [`Stubs.js`](Stubs.js) into `Code.gs`
6. In the left sidebar, click **+** next to **Libraries**
7. Enter the Script ID: `1qvAlZ99zluKSr3ws4NsxH8xXo1FncbWzu6Yq6raBumHdCLpLDaKveM0T`
8. Click **Look up**, select the latest version
9. Set the identifier to `SpreadFinder`
10. Click **Add**
11. Save and reload the spreadsheet

The OptionTools menu will appear after reload.

---

## How It Works

The library architecture means:
- All logic lives in the SpreadFinder library
- Your spreadsheet has thin wrapper stubs that delegate to the library
- When the library is updated, you automatically get the latest version (after reload)

### Why Stubs Are Needed

Google Apps Script has limitations that require local wrapper functions:
- **Custom functions** (cell formulas like `=XLookupByKeys(...)`) must be defined locally
- **Triggers** (`onOpen`, `onEdit`) must be defined locally
- **Dialog callbacks** (`google.script.run`) must call local functions

The stubs in `Code.gs` simply forward calls to `SpreadFinder.functionName()`.

---

## Features

### Portfolio Value Charts

For each symbol, generates a `<SYMBOL>PortfolioValueByPrice` tab with:
- Config table for price range settings
- Data table with value calculations
- Four charts: Portfolio $ value, Portfolio % ROI, Individual spreads $, Individual spreads %

### SpreadFinder

Scans option prices to find attractive bull call spread opportunities. Configure filters on the SpreadFinderConfig sheet, then run to see ranked results.

### Custom Functions

Use these in cell formulas:

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
3. Download CSV
4. Save to Google Drive under `<DataFolder>/OptionPrices/`
5. Run **OptionTools > Refresh Option Prices**

### Transactions from E*Trade

1. Download transaction CSV from E*Trade
2. Save to Google Drive under `<DataFolder>/Etrade/`
3. Run **OptionTools > Import Transactions from E*Trade**

---

## Supported Positions

That tab contains:
- A per-symbol Config table
- A generated data table
- A line chart of **Price vs Portfolio Value**

### Supported positions
- Common stock
- Bull call spreads

### Tables vs Named Ranges (important)

Google Sheets **Tables** are not readable by Apps Script.

**Convention used:**
- Table name: `BullCallSpreads`
- Named range: `BullCallSpreadsTable`

The script:
1. Tries `BullCallSpreads`
2. Falls back to `BullCallSpreadsTable`

If a Table exists but the Named Range does not, the script tells you exactly what to create.

---

## Menus and Triggers

`onOpen.js` wires menu items into the spreadsheet UI.

`PlotPortfolioValueByPrice` also installs an `onEdit(e)` trigger that:
- Rebuilds only when a Config table is edited
- Ignores all other edits

---

## Design Philosophy

- **Spreadsheets are the source of truth**
- **Scripts accelerate, not obscure**
- **Explicit > magical**
- **Performance matters**
- **Fill realism matters**

This repo is opinionated — deliberately.

---

## Apps Script Library

Available as a Google Apps Script library:

- **Library name:** SpreadFinder
- **Script ID:** `1qvAlZ99zluKSr3ws4NsxH8xXo1FncbWzu6Yq6raBumHdCLpLDaKveM0T`
- **URL:** https://script.google.com/macros/library/d/1qvAlZ99zluKSr3ws4NsxH8xXo1FncbWzu6Yq6raBumHdCLpLDaKveM0T/58

---

## Development

### Pushing Changes to the Library

```bash
# Push local changes to Apps Script
clasp push

# Pull changes from Apps Script
clasp pull
```

### Testing

Run `runAllTests()` or individual `test_*` functions in the Apps Script editor.

---

## License

MIT License.

Use it, fork it, adapt it, improve it.

---

## Author

**Mark Riggins**

Built to support real-world option portfolios where Google Sheets remains the fastest modeling surface.
