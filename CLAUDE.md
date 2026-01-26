# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

OptionUtils is a Google Apps Script project for modeling option portfolios in Google Sheets. It provides custom functions, menu items, and automated charting for analyzing stock positions and option spreads.

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

- `PlotPortfolioValueByPrice.js` - Main entry point. Reads a "Portfolios" table defining which named ranges belong to which symbol, then generates per-symbol tabs with config tables, data tables, and charts ($ value and % ROI).

- `XLookupByKeys.js` - Multi-key lookup function with three-tier caching (in-memory memo → chunked/gzipped DocumentCache → sheet rebuild). Exposes `XLookupByKeys()` as a custom spreadsheet function.

- `refreshOptionPrices.js` - Imports option price CSVs from Google Drive (`Investing/Data/OptionPrices/<symbol>/`) into an `OptionPricesUploaded` sheet. Selects most recent file per expiration.

- `onOpen.js` - Registers the "OptionTools" menu with available actions.

**Data Flow**:
1. User defines positions in named ranges (stocks table, spreads table)
2. User creates "Portfolios" table mapping: Symbol | Type | RangeName
3. `PlotPortfolioValueByPrice` reads portfolios, parses positions, generates price-vs-value data and charts

**Named Range Convention**: Google Sheets Tables are not readable by Apps Script. The convention is:
- Table name: `TslaBullCallSpreads`
- Named range: `TslaBullCallSpreadsTable`

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
