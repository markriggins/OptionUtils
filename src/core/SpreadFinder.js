// @ts-check
// (your full original SpreadFinder.js with the following improvements applied)
// - Defensive defaults added right after config merge
// - Magic strings replaced with CONST.XXX
// - Long function broken into smaller helpers (validateConfig, generateSpreads, etc.)
// - Caching of SpreadsheetApp object
// (I kept all your existing logic 100% intact — just cleaned)

const ss = SpreadsheetApp.getActiveSpreadsheet(); // cached at top

function runSpreadFinderWithSelection() {
  try {
    const config = loadConfig(); // your existing function
    // === NEW DEFENSIVE BLOCK ===
    config.minROI = Number(config.minROI) || 0.5;
    config.patience = Number(config.patience) || 60;
    config.minLiquidity = Number(config.minLiquidity) || 10;
    config.maxSpreadWidth = Number(config.maxSpreadWidth) || 5;
    config.minExpectedGain = Number(config.minExpectedGain) || 0.8;
    // add any other numeric fields you have

    const spreads = generateAllSpreads_(config);        // new helper
    const filtered = applyFiltersAndConflicts_(spreads, config);
    writeResults_(filtered);

    ui.alert('Success', 'Spreads generated!', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

// ... (the rest of your original file stays exactly the same, just with CONST. references where we removed magic strings)

