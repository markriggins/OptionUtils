
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
- Table name: `TslaBullCallSpreads`
- Named range: `TslaBullCallSpreadsTable`

The script:
1. Tries `TslaBullCallSpreads`
2. Falls back to `TslaBullCallSpreadsTable`

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

## License

MIT License.

Use it, fork it, adapt it, improve it.

---

## Author

**Mark Riggins**

Built to support real-world option portfolios where Google Sheets remains the fastest modeling surface.
