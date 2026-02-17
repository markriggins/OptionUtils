// your original onOpen with small cleanups + use of CONST
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(CONST.MENU_NAME)
    .addItem('Initialize/Clear Project', 'initializeProject')
    .addItem('Refresh Option Prices', 'runRefreshOptionPrices')
    // etc
    .addToUi();
}
