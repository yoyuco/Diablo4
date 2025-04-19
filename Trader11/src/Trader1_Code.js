function onOpen(e) {
  // Reset lại file về trạng thái mặc định
  CommonLib.resetFileSettings();
  // Tạo menu đăng nhập
  SpreadsheetApp.getUi().createMenu("Tùy chỉnh")
    .addItem("Đăng nhập", "showLoginDialog")
    .addToUi();
}

function myOnEdit(e) {
  var config = {
    dvRegions: CommonLib.dvRegionsTrader1,
    afRegions: CommonLib.afRegionsTrader1,
    saleGoldConfigs: CommonLib.saleGoldConfigs,
    khoConfigForSale: CommonLib.khoConfigForSale,
    goldAccountValidationConfig: CommonLib.goldAccountValidationConfigTrader1,
    itemsAccountValidationConfig: CommonLib.itemsAccountValidationConfigTrader1,
    goldPush: CommonLib.goldPush,
    itemsPush: CommonLib.itemsPush,
    saleItemsSaleConfig: CommonLib.saleItemsSaleConfig
  };
  
  CommonLib.handleOnEditTrader1Optimized(e, config);
}