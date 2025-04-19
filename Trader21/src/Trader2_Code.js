function myOnOpen(e) {
  // Dựa vào tên file Trader21, tự cập nhật trạng thái online
  CommonLib.manualUpdateTraderStatus();
  SpreadsheetApp.getUi().createMenu("Tùy chỉnh")
    .addItem("Đăng nhập", "showLoginDialog")
    .addToUi();
}

function myOnEdit(e) {
  var config = {
    conversionRegions: CommonLib.conversionRegions,
    dvRegions: CommonLib.dvRegionsTrader2,
    afRegions: CommonLib.afRegionsTrader2,
    traderUnifiedConfig: CommonLib.traderUnifiedConfig, // sử dụng cấu hình unified
    itemsUpdateConfig: CommonLib.itemsUpdateConfig,
    goldAccountValidationConfig: CommonLib.goldAccountValidationConfigTrader2,
    itemsAccountValidationConfig: CommonLib.itemsAccountValidationConfigTrader2
  };
  CommonLib.handleOnEditTrader2(e, config);
}