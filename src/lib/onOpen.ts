// @ts-check
// your original onOpen with small cleanups + use of CONST
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(CONST.MENU_NAME)
    .addItem('Initialize/Clear Project', 'initializeProject')
    .addItem('Refresh Option Prices', 'runRefreshOptionPrices')
    // etc
    .addToUi();
}

// =============================================
// LIBRARY EXPORT (required after src/ refactor)
// Makes SpreadFinder.onOpen work from Stubs.ts
// =============================================
if (typeof SpreadFinder === 'undefined') {
  var SpreadFinder = {};
}
SpreadFinder.onOpen = onOpen;
