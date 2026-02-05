# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

SpreadFinder is a Google Apps Script project for finding and analyzing bull call spread opportunities. It scans option prices, ranks spreads by expected ROI, and provides interactive charts. It also supports portfolio modeling for tracking actual positions imported from E*Trade.

## Development Commands

```bash
# Push code to Google Apps Script
clasp push

# Pull code from Google Apps Script
clasp pull

# Run tests (in Apps Script editor)
# Execute runAllTests() or individual test_* functions
```

## Architecture

**Deployment**: Uses clasp to sync JavaScript files with Google Apps Script. All `.js` files in the root directory are pushed to the script project.

**Runtime Environment**: Google Apps Script V8 runtime with access to SpreadsheetApp, DriveApp, and CacheService APIs.

**Key Components**:

- `SpreadFinder.js` - Scans option prices to find bull call spread opportunities. Ranks by expected ROI using probability-of-touch model. Results written to `<SYMBOL>Spreads` sheet.

- `ImportEtrade.js` - Imports E*Trade portfolio and transaction CSVs into the Portfolio sheet. Pairs transactions into spreads, detects iron condors, and tracks closing prices.

- `PlotPortfolioValueByPrice.js` - Reads positions from the Portfolio sheet (or legacy named ranges), generates per-symbol `<SYMBOL>PortfolioValueByPrice` tabs with charts showing $ value and % ROI at expiration.

- `RefreshOptionPrices.js` - Imports option price CSVs from Google Drive (`SpreadFinder/DATA/OptionPrices/`) into the `OptionPricesUploaded` sheet. Selects most recent file per expiration.

- `XLookupByKeys.js` - Multi-key lookup function with three-tier caching. Exposes `XLookupByKeys()` as a custom spreadsheet function.

- `onOpen.js` - Registers the "OptionTools" menu with available actions.

**Data Flow**:
1. User imports positions via "Import Portfolio from E*Trade" or manually edits the Portfolio sheet
2. SpreadFinder scans OptionPricesUploaded data to find attractive spreads
3. PlotPortfolioValueByPrice generates charts showing portfolio value across price scenarios

**Named Range Convention**: Google Sheets Tables are not readable by Apps Script. The convention is:
- Table name: `Portfolio`
- Named range: `PortfolioTable`

The script tries the base name first, then appends "Table" as fallback.

## Testing

Tests use a `test_*` naming convention discovered by `runAllTests()`. The `assertEqual()` helper in `TestUtils.js` provides floating-point tolerant comparisons.

## Supported Position Types

- `stock` (aliases: shares, share, stocks)
- `bull-call-spread` (aliases: bull-call-spreads, BCS)
- `bull-put-spread` (aliases: bull-put-spreads, BPS)

## Design Philosophy

- Spreadsheets are the source of truth
- Scripts accelerate, not obscure
- Explicit > magical
