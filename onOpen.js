/**
 *  Add AppScript menu items to a google sheet
 *  OptionTools/
 *    RefreshOptionPrices
 *    Warm XLookup Cache
 **/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('OptionTools')
    .addItem('Refresh Option Prices', 'refreshOptionPrices')
    .addItem("PlotPortfolioValueByPrice", "PlotPortfolioValueByPrice")
    .addToUi();
}

function warmXLookupCache() {
  XLookupByKeys_WarmCache("OptionPricesUploaded", ["Symbol", "Expiration", "Strike", "Type"], ["Bid", "Ask"]);
  SpreadsheetApp.getActiveSpreadsheet().toast("Cache warmed", "Done", 3);
}
