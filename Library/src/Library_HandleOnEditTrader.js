/**
 * handleOnEditTrader1: Xử lý sự kiện onEdit cho Trader1.
 * Các xử lý bao gồm:
 *   - Data Validation & Auto Fill
 *   - Gold Sale Update (Trader1)
 *   - Items Update
 *   - Items Sale Update (lấy dữ liệu từ Kho và cập nhật giá bán cho items)
 *   - Account Dropdown Update cho bảng Gold & Items
 *   - PUSH DATA: gọi pushGoldDataFromTrader1 và pushItemsDataFromTrader1
 *
 * @param {Object} e - Sự kiện onEdit.
 * @param {Object} config - Cấu hình dành cho Trader1, bao gồm các trường:
 *        dvRegions, afRegions, saleGoldConfigs, khoConfigForSale,
 *        goldAccountValidationConfig, itemsAccountValidationConfig,
 *        itemsUpdateConfig, goldPush, itemsPush,
 *        saleItemsSaleConfig (mới) – dùng cho cập nhật giá bán items.
 */
function handleOnEditTrader1Optimized(e, config) {
  // Cập nhật trạng thái online theo file hiện hành
  manualUpdateTraderStatus();
  
  var sheet = e.range.getSheet();
  if (!isMainSheet(sheet)) return;
  
  var startRow = e.range.getRow();
  var numRows = e.range.getNumRows();
  if (startRow < 11) return;
  
  // Xử lý từng dòng thay đổi
  for (var r = startRow; r < startRow + numRows; r++) {
    // Data Validation & Auto Fill
    if (config.dvRegions) {
      config.dvRegions.forEach(function(dvReg) {
        if (isConditionColumn(e.range, dvReg.conditionCol)) {
          processRowGenericOptimized(sheet, r, {
            conditionCol: dvReg.conditionCol,
            targetStart: dvReg.targetStart,
            targetCount: dvReg.targetCount,
            dependentConfig: dvReg.dependentConfig,
            commonConfig: dvReg.commonConfig,
            finalClearCol: dvReg.finalClearCol
          });
        }
      });
    }
    if (config.afRegions) {
      config.afRegions.forEach(function(afReg) {
        autoFillGroup(sheet, r, afReg);
      });
    }
    
    // Cập nhật Gold Sale Record nếu có (giữ nguyên hàm updateGoldSaleRecord nếu không tối ưu thêm)
    if (config.saleGoldConfigs && config.khoConfigForSale) {
      config.saleGoldConfigs.forEach(function(saleConfig) {
        if (e.range.getColumn() >= saleConfig.saleQtyCol && e.range.getColumn() <= saleConfig.profitCol2) {
          updateGoldSaleRecord(sheet, r, saleConfig, config.khoConfigForSale);
        }
      });
    }
    
    // Cập nhật Items Sale Record nếu có
    if (config.saleItemsSaleConfig) {
  if (Array.isArray(config.saleItemsSaleConfig)) {
    config.saleItemsSaleConfig.forEach(function(saleConfigObj) {
      if (e.range.getColumn() >= saleConfigObj.saleConfig.salePriceCol &&
          e.range.getColumn() <= saleConfigObj.saleConfig.saleCurrencyCol) {
        updateItemsSaleRecordSale(e, saleConfigObj);
      }
    });
  } else {
    var saleItemsSale = config.saleItemsSaleConfig.saleConfig;
    if (e.range.getColumn() >= saleItemsSale.salePriceCol && e.range.getColumn() <= saleItemsSale.saleCurrencyCol) {
      updateItemsSaleRecordSale(e, config.saleItemsSaleConfig);
    }
  }
}
    
    // Cập nhật Account Dropdown cho bảng Gold và Items
    var rangeStartCol = e.range.getColumn();
    var rangeEndCol = rangeStartCol + e.range.getNumColumns() - 1;
    if (config.goldAccountValidationConfig) {
      if (rangeStartCol <= config.goldAccountValidationConfig.checkRange.endCol &&
          rangeEndCol >= config.goldAccountValidationConfig.checkRange.startCol) {
        applyGoldAccountDropdown(sheet, r, config.goldAccountValidationConfig);
      }
    }
    if (config.itemsAccountValidationConfig) {
      if (Array.isArray(config.itemsAccountValidationConfig)) {
        config.itemsAccountValidationConfig.forEach(function(itemConfig) {
          if (rangeStartCol <= itemConfig.checkRange.endCol &&
          rangeEndCol >= itemConfig.checkRange.startCol) {
            applyItemsAccountDropdown(sheet, r, itemConfig);
          }
        });
      } else {
        if (rangeStartCol <= config.itemsAccountValidationConfig.checkRange.endCol &&
        rangeEndCol >= config.itemsAccountValidationConfig.checkRange.startCol) {
          applyItemsAccountDropdown(sheet, r, config.itemsAccountValidationConfig);
        }
      }
    }
    
    // Khởi tạo biến nextInsertionRow cho các push (cho bảng Gold và Items riêng biệt)
  var goldNextInsertionRow;   // cho bảng Gold
  var itemsNextInsertionRow;  // cho bảng Items
  
  // Xử lý push dữ liệu cho từng dòng
  for (var r = startRow; r < startRow + numRows; r++) {
    // Nếu cột chỉnh sửa thuộc vùng push của bảng Gold
    if (e.range.getColumn() >= config.goldPush.checkRange.startCol && e.range.getColumn() <= config.goldPush.checkRange.endCol) {
      goldNextInsertionRow = pushGoldDataFromTrader1Optimized(r, config.goldPush, goldNextInsertionRow);
    }
    // Nếu cột chỉnh sửa thuộc vùng push của bảng Items
    if (config.itemsPush) {
    if (Array.isArray(config.itemsPush)) {
      config.itemsPush.forEach(function(pushConfig) {
        if (e.range.getColumn() >= pushConfig.checkRange.startCol && e.range.getColumn() <= pushConfig.checkRange.endCol) {
            itemsNextInsertionRow = pushItemsDataFromTrader1Optimized(r, pushConfig, itemsNextInsertionRow);
          }
        });
    } else {
      if (e.range.getColumn() >= config.itemsPush.checkRange.startCol && e.range.getColumn() <= config.itemsPush.checkRange.endCol) {
          itemsNextInsertionRow = pushItemsDataFromTrader1Optimized(r, config.itemsPush, itemsNextInsertionRow);
        }
      }
    }
  }
}
}


/**
 * handleOnEditTrader2: Xử lý sự kiện onEdit cho Trader2.
 * Các xử lý bao gồm:
 *   - Conversion Update (Trader2)
 *   - Gold Record Update (Trader2)
 *   - Data Validation & Auto Fill (nếu cấu hình có sẵn)
 *   - Items Update (nếu cấu hình có sẵn)
 *   - Account Dropdown Update cho bảng Gold & Items (nếu cấu hình có sẵn)
 *
 * @param {Object} e - Sự kiện onEdit.
 * @param {Object} config - Cấu hình dành cho Trader2, bao gồm:
 *        conversionRegions, dvRegions, afRegions, traderConfigs, khoConfigForBuy,
 *        itemsUpdateConfig, goldAccountValidationConfig, itemsAccountValidationConfig, v.v.
 */
function handleOnEditTrader2(e, config) {
  // Cập nhật trạng thái online dựa trên tên file hiện hành
  manualUpdateTraderStatus();
  
  var sheet = e.range.getSheet();
  if (!isMainSheet(sheet)) return;
  
  var startRow = e.range.getRow();
  var numRows = e.range.getNumRows();
  if (startRow < 11) return;
  
  var editedRow = startRow;
  var editedCol = e.range.getColumn();
  
  // --- PHẦN 1: Conversion Update (Trader2) ---
  if (config.conversionRegions) {
    for (var i = 0; i < config.conversionRegions.length; i++) {
      var region = config.conversionRegions[i];
      var priceCol = region.priceCol;
      var currencyCol = priceCol + 1;
      if (editedCol === priceCol || editedCol === currencyCol) {
        for (var r = startRow; r < startRow + numRows; r++) {
          updateConversionRowOptimized(sheet, r, {
            priceCol: priceCol,
            headerRow: 10,
            targetStartCol: region.targetStartCol,
            targetEndCol: region.targetEndCol
          });
        }
      }
    }
  }
  
  // --- PHẦN 2: Gold Record Update (Trader2) ---
if (config.traderUnifiedConfig && Array.isArray(config.traderUnifiedConfig)) {
  config.traderUnifiedConfig.forEach(function(goldConfig) {
    var goldRegion = goldConfig.trader.dataRegion;
    if (editedCol >= goldRegion.startCol && editedCol <= goldRegion.endCol) {
      for (var r = startRow; r < startRow + numRows; r++) {
        try {
          updateGoldRecordOptimized(r, goldConfig);
        } catch (error) {
          Logger.log("Lỗi cập nhật gold ở dòng " + r + ": " + error.message);
        }
      }
    }
  });
} else if (config.traderUnifiedConfig) { // Nếu vẫn là đối tượng đơn
  var goldRegion = config.traderUnifiedConfig.trader.dataRegion;
  if (editedCol >= goldRegion.startCol && editedCol <= goldRegion.endCol) {
    for (var r = startRow; r < startRow + numRows; r++) {
      try {
        updateGoldRecordOptimized(r, config.traderUnifiedConfig);
      } catch (error) {
        Logger.log("Lỗi cập nhật gold ở dòng " + r + ": " + error.message);
      }
    }
  }
}
  
  // --- PHẦN 3: Data Validation & Auto Fill ---
  var rangeList = sheet.getActiveRangeList();
  if (rangeList) {
    var ranges = rangeList.getRanges();
    var processedDV = {};
    for (var i = 0; i < ranges.length; i++) {
      var rng = ranges[i];
      var sRow = rng.getRow();
      var nRows = rng.getNumRows();
      var eRow = sRow + nRows - 1;
      
      if (config.dvRegions) {
        for (var d = 0; d < config.dvRegions.length; d++) {
          var dvReg = config.dvRegions[d];
          if (isConditionColumn(rng, dvReg.conditionCol)) {
            for (var r = sRow; r <= eRow; r++) {
              if (r >= 11 && !processedDV[r]) {
                processRow(sheet, r, {
                  conditionCol: dvReg.conditionCol,
                  targetStart: dvReg.targetStart,
                  targetCount: dvReg.targetCount,
                  dependentConfig: dvReg.dependentConfig,
                  commonConfig: dvReg.commonConfig,
                  finalClearCol: dvReg.finalClearCol
                });
                processedDV[r] = true;
              }
            }
          }
        }
      }
      
      if (config.afRegions) {
        for (var a = 0; a < config.afRegions.length; a++) {
          var afReg = config.afRegions[a];
          for (var r = sRow; r <= eRow; r++) {
            if (r >= 11 && r <= 1000) {
              autoFillGroup(sheet, r, afReg);
            }
          }
        }
      }
    }
  } else {
    var r = e.range.getRow();
    if (r >= 11) {
      if (config.dvRegions) {
        for (var d = 0; d < config.dvRegions.length; d++) {
          var dvReg = config.dvRegions[d];
          if (isConditionColumn(e.range, dvReg.conditionCol)) {
            processRow(sheet, r, {
              conditionCol: dvReg.conditionCol,
              targetStart: dvReg.targetStart,
              targetCount: dvReg.targetCount,
              dependentConfig: dvReg.dependentConfig,
              commonConfig: dvReg.commonConfig,
              finalClearCol: dvReg.finalClearCol
            });
          }
        }
      }
      if (config.afRegions) {
        for (var a = 0; a < config.afRegions.length; a++) {
          autoFillGroup(sheet, r, config.afRegions[a]);
        }
      }
    }
  }
  
  // --- PHẦN 5: Items Update ---
  if (editedCol >= config.itemsUpdateConfig.dataRegion.startCol && editedCol <= config.itemsUpdateConfig.dataRegion.endCol) {
    var itemsDataCheck = sheet.getRange(startRow, config.itemsUpdateConfig.dataRegion.startCol, 1, 
                          config.itemsUpdateConfig.dataRegion.endCol - config.itemsUpdateConfig.dataRegion.startCol + 1).getValues()[0];
    var isDataComplete = itemsDataCheck.every(function(cell) {
      return cell.toString().trim() !== "";
    });
    if (isDataComplete) {
      updateItemsRecord(e, config.itemsUpdateConfig);
    }
  }
  
  // --- PHẦN 6: Account Dropdown Update for Gold & Items ---
var rangeStartRow = e.range.getRow();
var rangeNumRows = e.range.getNumRows();
var rangeStartCol = e.range.getColumn();
var rangeEndCol = rangeStartCol + e.range.getNumColumns() - 1;

// Xử lý Gold Account Dropdown
if (config.goldAccountValidationConfig) {
  if (Array.isArray(config.goldAccountValidationConfig)) {
    config.goldAccountValidationConfig.forEach(function(goldConf) {
      if (rangeStartCol <= goldConf.checkRange.endCol && rangeEndCol >= goldConf.checkRange.startCol) {
        for (var r = rangeStartRow; r < rangeStartRow + rangeNumRows; r++) {
          if (r >= 11) {
            applyGoldAccountDropdown(sheet, r, goldConf);
          }
        }
      }
    });
  } else {
    if (rangeStartCol <= config.goldAccountValidationConfig.checkRange.endCol &&
        rangeEndCol >= config.goldAccountValidationConfig.checkRange.startCol) {
      for (var r = rangeStartRow; r < rangeStartRow + rangeNumRows; r++) {
        if (r >= 11) {
          applyGoldAccountDropdown(sheet, r, config.goldAccountValidationConfig);
        }
      }
    }
  }
}

// Xử lý Items Account Dropdown
if (config.itemsAccountValidationConfig) {
  if (Array.isArray(config.itemsAccountValidationConfig)) {
    config.itemsAccountValidationConfig.forEach(function(itemsConf) {
      if (rangeStartCol <= itemsConf.checkRange.endCol && rangeEndCol >= itemsConf.checkRange.startCol) {
        for (var r = rangeStartRow; r < rangeStartRow + rangeNumRows; r++) {
          if (r >= 11) {
            applyItemsAccountDropdown(sheet, r, itemsConf);
          }
        }
      }
    });
  } else {
    if (rangeStartCol <= config.itemsAccountValidationConfig.checkRange.endCol &&
        rangeEndCol >= config.itemsAccountValidationConfig.checkRange.startCol) {
      for (var r = rangeStartRow; r < rangeStartRow + rangeNumRows; r++) {
        if (r >= 11) {
          applyItemsAccountDropdown(sheet, r, config.itemsAccountValidationConfig);
        }
      }
    }
  }
}
  
  // Trader2 không thực hiện PUSH DATA (chức năng này của Trader1)
}