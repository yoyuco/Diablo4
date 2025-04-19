/**
 * 1.2 manualUpdateTraderStatus: Cập nhật trạng thái online dựa trên role.
 * Nếu không truyền role, lấy role từ tên file hiện hành.
 * @param {string} userRole (tùy chọn) Role của người dùng.
 */
function manualUpdateTraderStatus(userRole) {
  if (!userRole) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var fileName = ss.getName(); // Ví dụ: "Trader11"
    userRole = fileName.charAt(0).toLowerCase() + fileName.slice(1);
  } else {
    userRole = userRole.toString().trim().toLowerCase();
  }
  
  var ssDatabase = SpreadsheetApp.openById(DATABASE_ID);
  var sheet = ssDatabase.getSheetByName("OnlineStatus");
  if (!sheet) {
    Logger.log("Không tìm thấy sheet 'OnlineStatus' trong file Database");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toLowerCase() === userRole) {
      sheet.getRange(i + 1, 2).setValue("Online");
      sheet.getRange(i + 1, 3).setValue(new Date());
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow([userRole, "Online", new Date()]);
  }
  Logger.log("manualUpdateTraderStatus: Cập nhật thành công cho role: " + userRole);
}

/**
 * 1.3 updateOfflineStatus: Cập nhật trạng thái Offline cho các role sau ngưỡng thời gian.
 */
function updateOfflineStatus() { 
  var ssDatabase = SpreadsheetApp.openById(DATABASE_ID);
  var sheet = ssDatabase.getSheetByName("OnlineStatus");
  if (!sheet) {
    Logger.log("Không tìm thấy sheet 'OnlineStatus' trong file Database");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var now = new Date().getTime();
  var threshold = 15 * 60 * 1000; // 15 phút
  
  for (var i = 1; i < data.length; i++) {
    var currentRole = data[i][0];
    var status = data[i][1];
    var lastTime = data[i][2];
    
    if (status === "Online" && lastTime) {
      var lastTimestamp = new Date(lastTime).getTime();
      if (now - lastTimestamp > threshold) {
        sheet.getRange(i + 1, 2).setValue("Offline");
        sheet.getRange(i + 1, 3).setValue(new Date());
        Logger.log("Role " + currentRole + " chuyển sang Offline và timestamp cập nhật.");
      }
    } else if (status === "Offline") {
      Logger.log("Role " + currentRole + " đã ở Offline, giữ nguyên timestamp.");
    }
  }
}

/**
 * 1.4 updateTrader2List: Cập nhật danh sách Trader2 từ file Accounts trong Database.
 */
function updateTrader2List() {
  var ssDatabase = SpreadsheetApp.openById(DATABASE_ID);
  var sheetAcc = ssDatabase.getSheetByName("Accounts");
  if (!sheetAcc) {
    Logger.log("Không tìm thấy sheet 'Accounts' trong Database");
    return;
  }
  
  var lastRowAcc = sheetAcc.getLastRow();
  var data = sheetAcc.getRange(1, 12, lastRowAcc, 3).getValues();
  
  var userList = [];
  for (var i = 1; i < data.length; i++) {
    var role = data[i][2];
    if (role) {
      userList.push([role, "Offline", ""]);
    }
  }
  
  var sheetOnline = ssDatabase.getSheetByName("OnlineStatus");
  if (!sheetOnline) {
    sheetOnline = ssDatabase.insertSheet("OnlineStatus");
  } else {
    sheetOnline.clear();
  }
  
  sheetOnline.appendRow(["Role", "Status", "Timestamp"]);
  if (userList.length > 0) {
    sheetOnline.getRange(2, 1, userList.length, 3).setValues(userList);
  }
}

/* =============================================================================
 * [3] Data Validation & Auto Fill Generic Functions
 * =============================================================================
 */

function autoFillGeneric(sheet, row, config) {
  // Nếu timestampCol chưa được định nghĩa, đặt mặc định là (conditionStart - 2)
  if (config.timestampCol === undefined) {
    config.timestampCol = config.conditionStart - 2;
  }
  // Nếu không bỏ qua sequence và sequenceCol chưa định nghĩa, đặt mặc định là (conditionStart - 1)
  if (!config.skipSequence && config.sequenceCol === undefined) {
    config.sequenceCol = config.conditionStart - 1;
  }
  
  // Lấy toàn bộ giá trị trong vùng điều kiện
  var condRange = sheet.getRange(row, config.conditionStart, 1, config.conditionEnd - config.conditionStart + 1);
  var condValues = condRange.getValues()[0];
  var allFilled = condValues.every(function(val) {
    return val !== "" && val != null && String(val).trim() !== "";
  });
  if (!allFilled) return;
  
  var tsCell = sheet.getRange(row, config.timestampCol);
  // Nếu cấu hình diffTime = true, chỉ tính hiệu số thời gian nếu ô trống
  if (config.diffTime) {
    if (!tsCell.getValue()) {
      // Lấy giá trị ban đầu từ ô tại cột conditionStart (giả sử ô đó chứa thời gian bắt đầu)
      var baseTime = sheet.getRange(row, config.conditionStart).getValue();
      if (!baseTime || baseTime == "") return;
      var now = new Date();
      var diffMs = now - new Date(baseTime);
      tsCell.setValue(formatTimeDifference(diffMs));
    }
  } else {
    // Nếu không diffTime, chỉ set giá trị thời gian hiện tại nếu ô chưa có giá trị
    if (!tsCell.getValue()) {
      tsCell.setValue(new Date());
    }
  }
  
  // Xử lý Auto Fill Sequence nếu không bỏ qua sequence
  if (!config.skipSequence) {
    var seqCell = sheet.getRange(row, config.sequenceCol);
    if (!seqCell.getValue()) {
      var lastRowSheet = sheet.getLastRow();
      var seqRange = sheet.getRange(11, config.sequenceCol, lastRowSheet - 10, 1);
      var seqValues = seqRange.getValues();
      var maxNum = 0;
      for (var i = 0; i < seqValues.length; i++) {
        var num = parseInt(seqValues[i][0]);
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
      seqCell.setValue(maxNum + 1);
    }
  }
}

/**
 * 3.2 applyDropdownsGeneric: Áp dụng Data Validation cho một nhóm ô theo danh sách Named Range.
 * @param {Sheet} sheet
 * @param {number} row
 * @param {Object} config - { targetStartCol, listNames }
 */
function applyDropdownsGeneric(sheet, row, config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  config.listNames.forEach(function(name, i) {
    var namedRange = ss.getRangeByName(name);
    if (namedRange) {
      var rule = SpreadsheetApp.newDataValidation().requireValueInRange(namedRange, true).build();
      sheet.getRange(row, config.targetStartCol + i).setDataValidation(rule);
    } else {
      Logger.log("Không tìm thấy Named Range: " + name);
    }
  });
}

/**
 * 3.3 applyDependentDropdownsGeneric: Áp dụng Data Validation phụ thuộc dựa trên conditionValue.
 * @param {Sheet} sheet
 * @param {number} row
 * @param {Object} config - { targetStartCol, suffixes, conditionValue }
 */
function applyDependentDropdownsGeneric(sheet, row, config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  config.suffixes.forEach(function(suffix, i) {
    var nrName = config.conditionValue + suffix;
    var namedRange = ss.getRangeByName(nrName);
    if (namedRange) {
      var rule = SpreadsheetApp.newDataValidation().requireValueInRange(namedRange, true).build();
      sheet.getRange(row, config.targetStartCol + i).setDataValidation(rule);
    } else {
      Logger.log("Không tìm thấy Named Range: " + nrName);
    }
  });
}

/**
 * 3.4 processRowGeneric: Xử lý Data Validation cho một dòng dựa trên giá trị ô condition.
 * @param {Sheet} sheet
 * @param {number} row
 * @param {Object} config - { conditionCol, targetStart, targetCount, dependentConfig, commonConfig, finalClearCol }
 */
function processRowGenericOptimized(sheet, row, config) {
  // Lấy toàn bộ dữ liệu của dòng để xử lý (giảm số lần gọi getRange)
  var lastCol = sheet.getLastColumn();
  var rowData = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  var conditionValue = rowData[config.conditionCol - 1];
  
  // Xóa vùng cần thiết (vùng dropdown) một lần
  sheet.getRange(row, config.targetStart, 1, config.targetCount)
       .clearContent()
       .clearDataValidations();
       
  if (conditionValue === "" || conditionValue == null) return;
  
  // Lấy dữ liệu từ sheet "CpItems" để xác định dropdown nào sẽ áp dụng
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cpItemsSheet = ss.getSheetByName("CpItems");
  if (!cpItemsSheet) {
    Logger.log("Sheet 'CpItems' không tồn tại.");
    return;
  }
  var lastRowCp = cpItemsSheet.getLastRow();
  if (lastRowCp < 2) return;
  var specialCategories = cpItemsSheet.getRange(2, 2, lastRowCp - 1, 1).getValues().flat();
  
  // Áp dụng dropdown phụ thuộc hoặc chung tùy theo điều kiện
  if (specialCategories.indexOf(conditionValue) >= 0) {
    if (config.dependentConfig) {
      config.dependentConfig.conditionValue = conditionValue;
      applyDependentDropdownsGeneric(sheet, row, config.dependentConfig);
      if (config.finalClearCol) {
        sheet.getRange(row, config.finalClearCol).clearContent().clearDataValidations();
      }
    }
  } else {
    if (config.commonConfig) {
      applyDropdownsGeneric(sheet, row, config.commonConfig);
    }
  }
}

/**
 * 4.1 autoFillGroup: Wrapper gọi autoFillGeneric.
 * @param {Sheet} sheet
 * @param {number} row
 * @param {Object} config
 */
function autoFillGroup(sheet, row, config) {
  autoFillGeneric(sheet, row, config);
}

/**
 * 4.2 applyDependentDropdownsWrapper: Wrapper cho dropdown phụ thuộc.
 * @param {Sheet} sheet
 * @param {number} row
 * @param {string} conditionValue
 * @param {Object} config
 */
function applyDependentDropdownsWrapper(sheet, row, conditionValue, config) {
  var baseConfig = { targetStartCol: 22, suffixes: ["_List1", "_List2"] };
  var finalConfig = Object.assign({}, baseConfig, config, { conditionValue: conditionValue });
  applyDependentDropdownsGeneric(sheet, row, finalConfig);
}

/**
 * 4.3 applyCommonDropdownsWrapper: Wrapper cho dropdown chung.
 * @param {Sheet} sheet
 * @param {number} row
 * @param {Object} config
 */
function applyCommonDropdownsWrapper(sheet, row, config) {
  var baseConfig = { targetStartCol: 22, listNames: ["Common_List1", "Common_List2", "Common_List3"] };
  var finalConfig = Object.assign({}, baseConfig, config);
  applyDropdownsGeneric(sheet, row, finalConfig);
}

/**
 * 4.4 processRow: Wrapper gọi processRowGeneric.
 * @param {Sheet} sheet
 * @param {number} row
 * @param {Object} config
 */
function processRow(sheet, row, config) {
  processRowGenericOptimized(sheet, row, config);
}


/**
 * 5.3 updateConversionRow: Cập nhật giá quy đổi cho một dòng trên sheet Trader2.
 * @param {Sheet} sheet
 * @param {Number} row
 * @param {Object} options - { priceCol, currencyCol, headerRow, targetStartCol, targetEndCol }.
 */
function updateConversionRowOptimized(sheet, row, options) {
  var currencyCol = options.currencyCol || (options.priceCol + 1);
  var headerRow = options.headerRow || 10;
  var priceCell = sheet.getRange(row, options.priceCol).getValue();
  var currencyValue = sheet.getRange(row, currencyCol).getValue();
  
  if (priceCell === "" || priceCell == null || currencyValue === "" || currencyValue == null) {
    var clearWidth = options.targetEndCol - options.targetStartCol + 1;
    sheet.getRange(row, options.targetStartCol, 1, clearWidth).clearContent();
    return;
  }
  
  var price = parseFloat(priceCell);
  var sourceCurrency = currencyValue.toString().trim().toUpperCase();
  if (isNaN(price) || sourceCurrency === "") return;
  
  // Sử dụng rateMatrix được truyền vào (nếu có) hoặc tự gọi getConversionMatrix()
  var rateMatrix = options.rateMatrix || getConversionMatrix();
  if (!rateMatrix) {
    Logger.log("Không lấy được bảng tỷ giá");
    return;
  }
  
  var headerRange = sheet.getRange(headerRow, options.targetStartCol, 1, options.targetEndCol - options.targetStartCol + 1);
  var headerValues = headerRange.getValues()[0];
  
  var matrixCurrencies = rateMatrix[0].slice(1).map(function(item) {
    return item.toString().trim().toUpperCase();
  });
  
  var targetMapping = {};
  for (var col = options.targetStartCol; col <= options.targetEndCol; col++){
    var cellValue = sheet.getRange(headerRow, col).getValue().toString().trim().toUpperCase();
    if (matrixCurrencies.indexOf(cellValue) !== -1){
      targetMapping[cellValue] = col;
    }
  }
  
  for (var i = 0; i < matrixCurrencies.length; i++){
    var targetCurrency = matrixCurrencies[i];
    var converted;
    if (sourceCurrency === targetCurrency) {
      converted = price;
    } else {
      converted = convertPrice(price, sourceCurrency, targetCurrency, rateMatrix);
    }
    if (targetMapping[targetCurrency]) {
      sheet.getRange(row, targetMapping[targetCurrency]).setValue(converted);
    }
  }
}

/* =============================================================================
 * [6] Update Functions for Gold & Items
 * =============================================================================
 */

/**
 * Cấu hình mẫu:
 * config = {
 *   row: số thứ tự dòng cần cập nhật,
 *   trader: {
 *     fileId: "Trader21" (lấy từ mapping, ví dụ getFileIdOptimized("Trader21")),
 *     // Nếu đã có đối tượng sheet (được lưu khi add Library) thì truyền vào, nếu không, hàm sẽ tự mở lại từ fileId
 *     sheet: optional,
 *     dataRegion: { startCol: 1, endCol: 14 },
 *     qtyCol: 5,
 *     currencyCols: { usd: 8, vnd: 9, fg: 10 },
 *     accountCol: 11,
 *     rowHeader: 10
 *   },
 *   kho: {
 *     sheetName: "Items",
 *     rowData: 2,
 *     rowHeader: 1,
 *     qtyCol: 3,
 *     currencyAvgCols: { usd: 5, vnd: 6, fg: 7 }
 *   }
 * }
 */
function updateGoldRecordOptimized(row, config, ssKho, khoSheet) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    // Lấy sheet Trader: sử dụng config.trader.sheet nếu có, nếu không mở lại từ fileId
    var traderSheet = config.trader.sheet;
    if (!traderSheet) {
      if (!config.trader.fileId) throw new Error("Trader config fileId bị thiếu.");
      var ssTrader = SpreadsheetApp.openById(config.trader.fileId);
      traderSheet = ssTrader.getActiveSheet();
      Logger.log("Active Trader sheet: " + traderSheet.getName());
      if (!isMainSheet(traderSheet)) {
        var sheets = ssTrader.getSheets();
        for (var i = 0; i < sheets.length; i++) {
          if (isMainSheet(sheets[i])) {
            traderSheet = sheets[i];
            break;
          }
        }
      }
      if (!traderSheet) traderSheet = ssTrader.getSheets()[0];
    }
    Logger.log("Chosen Trader sheet: " + traderSheet.getName());
    
    // Kiểm tra nếu dòng đã được xử lý ("Pushed") để tránh cập nhật nhiều lần
    var pushedCol = config.trader.dataRegion.endCol + 1;
    var pushedVal = traderSheet.getRange(row, pushedCol).getValue().toString().trim();
    if (pushedVal === "Pushed") {
      Logger.log("updateGoldRecordOptimized: Row " + row + " đã được xử lý, bỏ qua.");
      return;
    }
    
    // Đọc dữ liệu dòng từ Trader sheet
    var lastCol = traderSheet.getLastColumn();
    var rowData = traderSheet.getRange(row, 1, 1, lastCol).getValues()[0];
    Logger.log("Row " + row + " data: " + rowData.join(", "));
    
    // Kiểm tra dữ liệu trong vùng trader.dataRegion
    for (var i = config.trader.dataRegion.startCol; i <= config.trader.dataRegion.endCol; i++) {
      if (rowData[i - 1] === "" || rowData[i - 1] == null) {
        Logger.log("Dữ liệu không đầy đủ tại row " + row + " cột " + i);
        return;
      }
    }
    
    var newQty = Number(rowData[config.trader.qtyCol - 1]);
    if (isNaN(newQty)) {
      Logger.log("Số lượng không hợp lệ tại row " + row);
      return;
    }
    
    // Lấy giá trị từ các cột cấu hình (nếu có) trong config.trader.currencyCols
    var newCurrencyTotals = {};
    if (config.trader.currencyCols) {
      for (var key in config.trader.currencyCols) {
        newCurrencyTotals[key.toLowerCase()] = Number(rowData[config.trader.currencyCols[key] - 1]);
        if (isNaN(newCurrencyTotals[key.toLowerCase()])) newCurrencyTotals[key.toLowerCase()] = 0;
      }
    }
    // Nếu không có giá trị USD được cung cấp, cố gắng lấy từ cột purchasePriceCol (nếu đã định nghĩa)
    if (newCurrencyTotals["usd"] === undefined && config.trader.purchasePriceCol) {
      newCurrencyTotals["usd"] = Number(rowData[config.trader.purchasePriceCol - 1]);
      if (isNaN(newCurrencyTotals["usd"])) {
         Logger.log("Không có giá mua USD hợp lệ tại row " + row);
         return;
      }
    }
    
    // Lấy thông tin account từ Trader sheet (ví dụ "Acc05-Trader21")
    var traderAccount = rowData[config.trader.accountCol - 1].toString().trim();
    if (!traderAccount) {
      Logger.log("Thiếu thông tin account tại row " + row);
      return;
    }
    var prefix = traderAccount.indexOf("-") > -1 ? traderAccount.split("-")[0] : traderAccount;
    
    // Xác định header cần tìm: nếu không tìm thấy "AccXX-Helm" thì thử "AccXX-Spirit1"
    var desiredHeader1 = prefix + "-Helm";
    var desiredHeader2 = prefix + "-Spirit1";
    
    // Mở file Kho nếu chưa truyền ssKho và khoSheet
    if (!ssKho || !khoSheet) {
      var khoId = getFileIdOptimized("Kho");
      if (!khoId) throw new Error("Không lấy được ID của file Kho từ mapping.");
      ssKho = SpreadsheetApp.openById(khoId);
      khoSheet = ssKho.getSheetByName(config.kho.sheetName);
      if (!khoSheet) throw new Error("Không tìm thấy sheet '" + config.kho.sheetName + "' trong file Kho.");
      Logger.log("Kho sheet: " + khoSheet.getName());
    }
    
    var khoRow = config.kho.rowData || 2;
    var khoHeaderRow = config.kho.rowHeader || 1;
    var headerValues = khoSheet.getRange(khoHeaderRow, 1, 1, khoSheet.getLastColumn()).getValues()[0];
    var targetCol = null;
    var desiredHeader = desiredHeader1;
    for (var j = 0; j < headerValues.length; j++) {
      if (headerValues[j].toString().trim().toUpperCase() === desiredHeader1.toUpperCase()) {
        targetCol = j + 1;
        break;
      }
    }
    if (!targetCol) {
      for (var j = 0; j < headerValues.length; j++) {
        if (headerValues[j].toString().trim().toUpperCase() === desiredHeader2.toUpperCase()) {
          targetCol = j + 1;
          desiredHeader = desiredHeader2;
          break;
        }
      }
    }
    if (!targetCol) {
      throw new Error("Không tìm thấy cột cho account " + desiredHeader1 + " hoặc " + desiredHeader2 + " trong sheet Kho");
    }
    
    // *** KIỂM TRA GIỚI HẠN SỐ LƯỢNG CHO ACCOUNT TRƯỚC UPDATE ***
    var accCell = khoSheet.getRange(khoRow, targetCol);
    var oldAccQty = Number(accCell.getValue());
    if (isNaN(oldAccQty)) oldAccQty = 0;
    if (oldAccQty + newQty > 99999) {
      SpreadsheetApp.getUi().alert("Tổng số lượng gold cho account " + desiredHeader + " vượt quá giới hạn (99999). Vui lòng chọn account khác.");
      return;
    }
    
    // Tính lại tổng số lượng ở kho bằng cách lấy giá trị cũ của ô tổng (config.kho.qtyCol)
    var oldTotalQty = Number(khoSheet.getRange(khoRow, config.kho.qtyCol).getValue());
    if (isNaN(oldTotalQty)) oldTotalQty = 0;
    var updatedTotalQty = oldTotalQty + newQty;
    
    // Tính trung bình giá cho mỗi đơn vị theo từng loại tiền (sử dụng công thức trọng số)
    var rateMatrix = getConversionMatrix();
    if (!rateMatrix) {
      Logger.log("updateGoldRecordOptimized: Không lấy được bảng tỷ giá");
      return;
    }
    
    var currenciesToUpdate = Object.keys(config.kho.currencyAvgCols);
    for (var i = 0; i < currenciesToUpdate.length; i++) {
      var currKey = currenciesToUpdate[i].toLowerCase();
      var colIndex = config.kho.currencyAvgCols[currenciesToUpdate[i]];
      var oldAvg = Number(khoSheet.getRange(khoRow, colIndex).getValue());
      if (isNaN(oldAvg)) oldAvg = 0;
      var newPrice;
      if (newCurrencyTotals[currKey] !== undefined) {
        newPrice = newCurrencyTotals[currKey];
      } else {
        newPrice = convertPrice(newCurrencyTotals["usd"], "USD", currKey.toUpperCase(), rateMatrix);
        if (newPrice === null) {
          Logger.log("updateGoldRecordOptimized: Không chuyển đổi được giá từ USD sang " + currKey.toUpperCase());
          continue;
        }
      }
      // Tính trung bình giá theo công thức trọng số:
      // updatedAvg = ((oldAvg * oldTotalQty) + (newPrice * newQty)) / (oldTotalQty + newQty)
      var updatedAvg = ((oldAvg * oldTotalQty) + newPrice) / (oldTotalQty + newQty);
      khoSheet.getRange(khoRow, colIndex).setValue(updatedAvg);
      Logger.log("Cập nhật " + currKey.toUpperCase() + " - cột " + colIndex + ": Giá trung bình mới = " + updatedAvg);
    }
    
    // Cập nhật tổng số lượng ở cột tổng
    khoSheet.getRange(khoRow, config.kho.qtyCol).setValue(updatedTotalQty);
    Logger.log("updateGoldRecordOptimized: Tổng số lượng mới tại kho = " + updatedTotalQty);
    
    // *** CẬP NHẬT SỐ LƯỢNG CHO ACCOUNT RIÊNG ***
    var updatedAccQty = oldAccQty + newQty;
    accCell.setValue(updatedAccQty);
    Logger.log("updateGoldRecordOptimized: Cập nhật số lượng cho " + desiredHeader + " = " + updatedAccQty);
    
    SpreadsheetApp.flush();
    
    // Đánh dấu dòng trên Trader là "Pushed" và khóa vùng dữ liệu đó
    traderSheet.getRange(row, pushedCol).setValue("Pushed");
    applyProtection(traderSheet, row, config.trader.dataRegion.startCol, pushedCol, "Gold data updated and locked by admin");
    Logger.log("updateGoldRecordOptimized: Cập nhật gold thành công tại dòng " + row);
    
  } catch (e) {
    Logger.log("updateGoldRecordOptimized error: " + e);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/**
 * updateGoldSaleRecord: Cập nhật thông tin bán gold cho Trader1.
 * Nếu saleCurrencyCol là VND:
 *    - purchasePrice = (giá trung bình VND ở Kho * saleQty)
 *    - profitCol1 = 0; profitCol2 = salePrice - purchasePrice
 * Nếu saleCurrencyCol là đơn vị khác:
 *    - purchasePrice = (giá trung bình USD ở Kho * saleQty), nếu saleCurrency là USD thì giữ nguyên,
 *      nếu không, chuyển đổi sang USD qua convertPrice.
 *    - profitCol1 = salePrice - purchasePrice; profitCol2 = 0.
 *
 * Nếu ô saleQty, salePrice hoặc saleCurrency bị xóa, thì sẽ xóa luôn 3 ô purchasePriceCol, profitCol1, profitCol2.
 *
 * @param {Sheet} sheet - Sheet của Trader1.
 * @param {number} row - Dòng cần cập nhật.
 * @param {Object} saleConfig - Cấu hình bán gold, gồm:
 *         saleQtyCol, salePriceCol, saleCurrencyCol, purchasePriceCol, profitCol1, profitCol2.
 * @param {Object} khoConfig - Cấu hình file Kho, gồm:
 *         sheetName, rowData, ...
 */
function updateGoldSaleRecord(sheet, row, saleConfig, khoConfig) {
  // Lấy giá trị từ các ô saleQty, salePrice, saleCurrency
  var saleQtyCell = sheet.getRange(row, saleConfig.saleQtyCol);
  var salePriceCell = sheet.getRange(row, saleConfig.salePriceCol);
  var saleCurrencyCell = sheet.getRange(row, saleConfig.saleCurrencyCol);
  
  var saleQty = saleQtyCell.getValue();
  var salePrice = salePriceCell.getValue();
  var saleCurrency = saleCurrencyCell.getValue();
  
  // Nếu một trong các ô này bị xóa (rỗng), xóa luôn 3 cột purchasePrice, profitCol1, profitCol2
  if (!saleQty || !salePrice || saleCurrency.toString().trim() === "") {
    sheet.getRange(row, saleConfig.purchasePriceCol).clearContent();
    sheet.getRange(row, saleConfig.profitCol1).clearContent();
    sheet.getRange(row, saleConfig.profitCol2).clearContent();
    Logger.log("updateGoldSaleRecord: Đã xóa purchasePrice, profitCol1, profitCol2 vì saleQty/salePrice/saleCurrency bị xóa tại dòng " + row);
    return;
  }
  
  // Chuyển đổi saleQty và salePrice sang số
  saleQty = Number(saleQty);
  salePrice = Number(salePrice);
  if (isNaN(saleQty) || isNaN(salePrice)) {
    Logger.log("updateGoldSaleRecord: Số liệu không hợp lệ ở dòng " + row);
    return;
  }
  
  // Lấy sale currency và ép về chữ thường
  saleCurrency = saleCurrency.toString().trim().toLowerCase();
  
  // Mở file Kho và lấy sheet
  var khoId = getFileIdOptimized("Kho");
  if (!khoId) throw new Error("updateGoldSaleRecord: Không lấy được ID của file Kho từ mapping.");
  var ssKho = SpreadsheetApp.openById(khoId);
  var khoSheet = ssKho.getSheetByName(khoConfig.sheetName);
  if (!khoSheet) throw new Error("updateGoldSaleRecord: Không tìm thấy sheet '" + khoConfig.sheetName + "' trong file Kho.");
  
  // Giả sử record Gold trong Kho nằm ở dòng khoConfig.rowData (ví dụ dòng 2)
  var khoRow = khoConfig.rowData;
  
  // Lấy header của file Kho (dòng 1)
  var headerValues = khoSheet.getRange(1, 1, 1, khoSheet.getLastColumn()).getValues()[0];
  var targetCol = null;
  
  // Nếu saleCurrency là VND, tìm cột có header "VND", ngược lại dùng "USD"
  if (saleCurrency === "vnd") {
    for (var i = 0; i < headerValues.length; i++) {
      if (headerValues[i].toString().trim().toUpperCase() === "VND") {
        targetCol = i + 1;
        break;
      }
    }
  } else {
    for (var i = 0; i < headerValues.length; i++) {
      if (headerValues[i].toString().trim().toUpperCase() === "USD") {
        targetCol = i + 1;
        break;
      }
    }
  }
  
  if (!targetCol) {
    throw new Error("updateGoldSaleRecord: Không tìm thấy cột trung bình giá cho " + (saleCurrency === "vnd" ? "VND" : "USD") + " trong file Kho.");
  }
  
  // Lấy giá trung bình mua từ file Kho
  var avgPrice = Number(khoSheet.getRange(khoRow, targetCol).getValue());
  if (isNaN(avgPrice) || avgPrice <= 0) {
    throw new Error("updateGoldSaleRecord: Giá trung bình mua ở Kho không hợp lệ ở dòng " + khoRow);
  }
  
  var purchasePrice, profitCol1, profitCol2;
  if (saleCurrency === "vnd") {
    // Nếu bán bằng VND
    purchasePrice = avgPrice * saleQty;
    profitCol1 = 0;
    profitCol2 = salePrice - purchasePrice;
  } else {
    // Nếu bán bằng đơn vị khác
    var totalAvgPrice = avgPrice * saleQty;
    if (saleCurrency === "usd") {
      purchasePrice = totalAvgPrice;
    } else {
      var rateMatrix = getConversionMatrix();
      if (!rateMatrix) throw new Error("updateGoldSaleRecord: Không lấy được bảng tỷ giá");
      var converted = convertPrice(totalAvgPrice, saleCurrency.toUpperCase(), "USD", rateMatrix);
      if (converted === null) throw new Error("updateGoldSaleRecord: Không chuyển đổi được giá từ " + saleCurrency + " sang USD");
      purchasePrice = converted;
    }
    profitCol1 = salePrice - purchasePrice;
    profitCol2 = 0;
  }
  
  // Cập nhật giá trị vào sheet Trader1: đặt purchasePrice, profitCol1 và profitCol2
  sheet.getRange(row, saleConfig.purchasePriceCol).setValue(purchasePrice);
  sheet.getRange(row, saleConfig.profitCol1).setValue(profitCol1);
  sheet.getRange(row, saleConfig.profitCol2).setValue(profitCol2);
  
  Logger.log("updateGoldSaleRecord: Ở dòng " + row + ", purchasePrice=" + purchasePrice + ", profitCol1=" + profitCol1 + ", profitCol2=" + profitCol2);
}

/**
 * updateItemsRecord:
 * Cập nhật (hoặc chèn mới) record Items từ Trader2 vào file Kho.
 * Cải tiến: Sử dụng batch read và DocumentLock để tăng tốc và tránh trùng lặp trên file Kho.
 *
 * @param {Object} e - Sự kiện onEdit từ Trader2.
 * @param {Object} config - Cấu hình itemsUpdateConfig, bao gồm:
 *        { dataRegion: { startCol, endCol },
 *          itemRegion: { startCol, endCol },
 *          accountCol, pushedCol, currencyCols }
 */
function updateItemsRecord(e, config) {
  var khoLock = LockService.getScriptLock();
  try {
    khoLock.waitLock(30000); // Chờ tối đa 30 giây để lấy khóa

    var row = e.range.getRow();
    var traderSheet = e.range.getSheet();
    
    // Kiểm tra nếu row đã được xử lý (đã được đánh dấu "Pushed")
    var pushedVal = traderSheet.getRange(row, config.pushedCol).getValue().toString().trim();
    if (pushedVal === "Pushed") {
      Logger.log("Row " + row + " đã được xử lý, bỏ qua updateItemsRecord.");
      return;
    }
    
    var lastCol = traderSheet.getLastColumn();
    // Đọc toàn bộ dữ liệu của dòng từ Trader sheet một lần
    var rowData = traderSheet.getRange(row, 1, 1, lastCol).getValues()[0];
    
    // Kiểm tra vùng dữ liệu chính đã đầy đủ chưa
    var dataValues = rowData.slice(config.dataRegion.startCol - 1, config.dataRegion.endCol);
    var complete = dataValues.every(function(cell) {
      return cell.toString().trim() !== "";
    });
    if (!complete) return;
    
    // Lấy dữ liệu của vùng items
    var itemData = rowData.slice(config.itemRegion.startCol - 1, config.itemRegion.endCol);
    if (itemData[0].toString().trim() === "") return;
    
    var itemString = combineItemString(itemData);
    Logger.log("Combined Item String: " + itemString);
    
    // Mở file Kho và lấy sheet theo khoConfigForItems (giả sử khoConfigForItems được cấu hình trong thư viện)
    var khoId = getFileIdOptimized("Kho");
    if (!khoId) throw new Error("Không lấy được ID của file Kho");
    var ssKho = SpreadsheetApp.openById(khoId);
    var khoConfig = khoConfigForItems; // cấu hình kho đã được định nghĩa trong thư viện
    if (!khoConfig) throw new Error("Cấu hình khoConfigForItems không tồn tại");
    var khoSheet = ssKho.getSheetByName(khoConfig.sheetName);
    if (!khoSheet) throw new Error("Không tìm thấy sheet " + khoConfig.sheetName + " trong file Kho");
    
    var startRowKho = khoConfig.rowData;
    var itemsCol = khoConfig.itemsCol;
    var lastRowKho = khoSheet.getLastRow();
    var foundRow = null;
    var insertPosition = startRowKho;
    
    // Đọc toàn bộ cột items từ khoSheet một lần để tìm record phù hợp
    if (lastRowKho >= startRowKho) {
      var itemsRange = khoSheet.getRange(startRowKho, itemsCol, lastRowKho - startRowKho + 1, 1);
      var itemsValues = itemsRange.getValues();
      for (var i = 0; i < itemsValues.length; i++) {
        var currentItem = itemsValues[i][0].toString().trim();
        var cmp = itemString.localeCompare(currentItem, 'vi', { sensitivity: 'base' });
        if (cmp === 0) {
          foundRow = startRowKho + i;
          break;
        } else if (cmp < 0) {
          insertPosition = startRowKho + i;
          break;
        } else {
          insertPosition = startRowKho + i + 1;
        }
      }
    } else {
      insertPosition = startRowKho;
    }
    
    // Lấy thông tin Account từ Trader sheet
    var traderAccount = rowData[config.accountCol - 1].toString().trim();
    if (!traderAccount) {
      Logger.log("Không có thông tin account tại dòng " + row);
      return;
    }
    var desiredHeader = traderAccount;
    Logger.log("Desired account header: " + desiredHeader);
    
    // Đọc header của khoSheet để xác định cột account tương ứng
    var headerValues = khoSheet.getRange(1, 1, 1, khoSheet.getLastColumn()).getValues()[0];
    var accountColIndex = null;
    for (var j = 0; j < headerValues.length; j++) {
      if (headerValues[j].toString().trim().toLowerCase() === desiredHeader.toLowerCase()) {
        accountColIndex = j + 1;
        break;
      }
    }
    if (!accountColIndex) {
      throw new Error("Không tìm thấy cột account cho header " + desiredHeader + " trong file Kho");
    }
    
    if (foundRow) {
      // Nếu tìm thấy record, cập nhật số lượng và trung bình giá
      var qtyCol = khoConfig.qtyCol;
      var oldQty = Number(khoSheet.getRange(foundRow, qtyCol).getValue());
      if (isNaN(oldQty)) oldQty = 0;
      var newQty = oldQty + 1;
      khoSheet.getRange(foundRow, qtyCol).setValue(newQty);
      
      if (config.currencyCols && Object.keys(config.currencyCols).length > 0) {
        for (var key in config.currencyCols) {
          var traderCol = config.currencyCols[key];
          var traderPrice = Number(rowData[traderCol - 1]);
          if (isNaN(traderPrice)) traderPrice = 0;
          
          var khoCol = khoConfig.currencyAvgCols[key];
          if (!khoCol) continue;
          var oldAvg = Number(khoSheet.getRange(foundRow, khoCol).getValue());
          if (isNaN(oldAvg)) oldAvg = 0;
          var updatedAvg = ((oldAvg * oldQty) + traderPrice) / newQty;
          khoSheet.getRange(foundRow, khoCol).setValue(updatedAvg);
        }
      }
      
      // Cập nhật số lượng cho account
      var accCell = khoSheet.getRange(foundRow, accountColIndex);
      var oldAccQty = Number(accCell.getValue());
      if (isNaN(oldAccQty)) oldAccQty = 0;
      accCell.setValue(oldAccQty + 1);
      Logger.log("Updated existing item record at row " + foundRow);
    } else {
      // Nếu không tìm thấy, chèn dòng mới tại insertPosition
      khoSheet.insertRowBefore(insertPosition);
      khoSheet.getRange(insertPosition, itemsCol).setValue(itemString);
      khoSheet.getRange(insertPosition, khoConfig.qtyCol).setValue(1);
      
      if (config.currencyCols && Object.keys(config.currencyCols).length > 0) {
        for (var key in config.currencyCols) {
          var traderCol = config.currencyCols[key];
          var traderPrice = Number(rowData[traderCol - 1]);
          if (isNaN(traderPrice)) traderPrice = 0;
          var khoCol = khoConfig.currencyAvgCols[key];
          if (!khoCol) continue;
          khoSheet.getRange(insertPosition, khoCol).setValue(traderPrice);
        }
      }
      
      // Cập nhật số lượng cho account trong dòng mới
      var accountColIndexNew = null;
      for (var j = 0; j < headerValues.length; j++) {
        if (headerValues[j].toString().trim().toLowerCase() === desiredHeader.toLowerCase()) {
          accountColIndexNew = j + 1;
          break;
        }
      }
      if (!accountColIndexNew) {
        throw new Error("Không tìm thấy cột account cho header " + desiredHeader + " trong file Kho");
      }
      khoSheet.getRange(insertPosition, accountColIndexNew).setValue(1);
      Logger.log("Inserted new item record at row " + insertPosition);
    }
    
    // Đánh dấu dòng là "Pushed" và khóa phạm vi dữ liệu đó để tránh xử lý nhiều lần
    traderSheet.getRange(row, config.pushedCol).setValue("Pushed");
    applyProtection(traderSheet, row, config.checkRange.startCol, config.checkRange.endCol, "Data pushed and locked by admin");
    
  } finally {
    khoLock.releaseLock();
  }
}

/**
 * 6.4 updateItemsSaleRecordSale: Cập nhật bảng giá bán cho items từ đơn hàng Trader1 vào file Kho.
 *
 * Đặc điểm:
 * - Qty luôn bằng 1.
 * - Dữ liệu item được xác định từ 4 cột liên tiếp được chỉ định trong saleItemsConfig.sourceDataRange, ghép thành chuỗi bằng combineItemString.
 * - Hàm tìm dòng (rowData) của item trong Kho (trong sheet khoConfig.sheetName, cột khoConfig.itemsCol, bắt đầu từ khoConfig.rowData).
 *   Nếu không tìm thấy, hàm sẽ không thực hiện cập nhật.
 * - Nếu saleCurrency là "vnd":
 *      • purchasePrice = 1 * (giá trung bình mua VND lấy từ kho ở cột khoConfig.currencyAvgCols.vnd)
 *      • profitVND = salePrice - purchasePrice, profitUSD = 0.
 * - Nếu saleCurrency không phải "vnd":
 *      • purchasePrice = 1 * (giá trung bình mua USD lấy từ kho ở cột khoConfig.currencyAvgCols.usd)
 *      • Nếu saleCurrency là "usd": profitUSD = salePrice - purchasePrice.
 *        Nếu không, chuyển đổi salePrice sang USD (dùng convertPrice) và tính profitUSD = convertedSalePrice - purchasePrice.
 *      • profitVND = 0.
 *
 * Các cột cập nhật trong file Kho được xác định qua khoConfig.purchasePriceCol, profitCol1 và profitCol2.
 *
 * @param {Object} e - Sự kiện onEdit từ Trader sheet.
 * @param {Object} config - Đối tượng chứa:
 *      saleItemsConfig: { sourceDataRange: { startCol, endCol } },
 *      saleConfig: { salePriceCol, saleCurrencyCol, purchasePriceCol, profitCol1, profitCol2 },
 *      khoConfig: {
 *         sheetName, 
 *         rowData, 
 *         itemsCol, 
 *         qtyCol,
 *         currencyAvgCols: { usd: colUSD, vnd: colVND },
 *         purchasePriceCol, 
 *         profitCol1, 
 *         profitCol2
 *      }
 */
function updateItemsSaleRecordSale(e, config) {
  var row = e.range.getRow();
  var traderSheet = e.range.getSheet();
  
  // Lấy giá bán và đơn vị tiền từ Trader1 (saleConfig)
  var salePriceCell = traderSheet.getRange(row, config.saleConfig.salePriceCol);
  var salePrice = Number(salePriceCell.getValue());
  var currencyCell = config.saleConfig.saleCurrencyCol ?
                     traderSheet.getRange(row, config.saleConfig.saleCurrencyCol) :
                     traderSheet.getRange(row, config.saleConfig.salePriceCol + 1);
  var saleCurrency = currencyCell.getValue().toString().trim().toLowerCase();
  
  // Nếu ô giá bán hoặc đơn vị tiền bị xóa, xóa luôn các ô giá mua và lợi nhuận
  if (!salePrice || saleCurrency === "") {
    traderSheet.getRange(row, config.saleConfig.purchasePriceCol).clearContent();
    traderSheet.getRange(row, config.saleConfig.profitCol1).clearContent();
    traderSheet.getRange(row, config.saleConfig.profitCol2).clearContent();
    Logger.log("updateItemsSaleRecordSale: Đã xóa các giá trị do thiếu salePrice hoặc saleCurrency ở dòng " + row);
    return;
  }
  
  // 2. Lấy chuỗi định danh của item từ 4 cột (saleItemsConfig)
  var itemRange = traderSheet.getRange(
    row, 
    config.saleItemsConfig.sourceDataRange.startCol, 
    1,
    config.saleItemsConfig.sourceDataRange.endCol - config.saleItemsConfig.sourceDataRange.startCol + 1
  );
  var itemData = itemRange.getValues()[0];
  var itemString = combineItemString(itemData);
  if (itemString === "") {
    Logger.log("updateItemsSaleRecordSale: Chuỗi item rỗng.");
    return;
  }
  
  // 3. Mở file Kho và lấy sheet theo khoConfig.sheetName
  var khoId = getFileIdOptimized("Kho");
  if (!khoId) throw new Error("Không lấy được ID của file Kho");
  var ssKho = SpreadsheetApp.openById(khoId);
  var khoSheet = ssKho.getSheetByName(config.khoConfig.sheetName);
  if (!khoSheet) throw new Error("Không tìm thấy sheet " + config.khoConfig.sheetName + " trong file Kho");
  
  // 4. Tìm dòng record của item trong Kho (dựa trên cột khoConfig.itemsCol, bắt đầu từ rowData)
  var startRowKho = config.khoConfig.rowData;
  var itemsCol = config.khoConfig.itemsCol;
  var lastRowKho = khoSheet.getLastRow();
  var foundRow = null;
  if (lastRowKho >= startRowKho) {
    var itemsRangeKho = khoSheet.getRange(startRowKho, itemsCol, lastRowKho - startRowKho + 1, 1);
    var itemsValues = itemsRangeKho.getValues();
    for (var i = 0; i < itemsValues.length; i++) {
      var currentItem = itemsValues[i][0].toString().trim();
      if (itemString.localeCompare(currentItem, 'vi', { sensitivity: 'base' }) === 0) {
        foundRow = startRowKho + i;
        break;
      }
    }
  }
  
  if (!foundRow) {
    Logger.log("updateItemsSaleRecordSale: Không tìm thấy record cho item: " + itemString);
    return;
  }
  
  // 5. Lấy bảng tỷ giá để chuyển đổi nếu cần
  var rateMatrix = getConversionMatrix();
  if (!rateMatrix) {
    Logger.log("updateItemsSaleRecordSale: Không lấy được bảng tỷ giá.");
    return;
  }
  
  // 6. Tính toán purchasePrice và lợi nhuận (qty luôn = 1)
  var purchasePrice, profitUSD, profitVND, avgPrice;
  if (saleCurrency === "vnd") {
    // Lấy giá trung bình mua VND từ Kho tại dòng foundRow, cột được xác định bởi khoConfig.currencyAvgCols.vnd
    avgPrice = Number(khoSheet.getRange(foundRow, config.khoConfig.currencyAvgCols.vnd).getValue());
    if (isNaN(avgPrice)) avgPrice = 6; // fallback
    purchasePrice = 1 * avgPrice;
    profitVND = salePrice - purchasePrice;
    profitUSD = 0;
  } else {
    // Lấy giá trung bình mua USD từ Kho tại dòng foundRow, cột khoConfig.currencyAvgCols.usd
    avgPrice = Number(khoSheet.getRange(foundRow, config.khoConfig.currencyAvgCols.usd).getValue());
    if (isNaN(avgPrice)) avgPrice = 5; // fallback
    purchasePrice = 1 * avgPrice;
    if (saleCurrency === "usd") {
      profitUSD = salePrice - purchasePrice;
    } else {
      var convertedSalePrice = convertPrice(salePrice, saleCurrency.toUpperCase(), "USD", rateMatrix);
      if (convertedSalePrice === null) {
        Logger.log("updateItemsSaleRecordSale: Không chuyển đổi được giá bán từ " + saleCurrency + " sang USD.");
        profitUSD = salePrice - purchasePrice;
      } else {
        profitUSD = convertedSalePrice - purchasePrice;
      }
    }
    profitVND = 0;
  }
  
  // 7. Cập nhật các giá trị vào file Trader1 tại các cột được cấu hình trong saleConfig
  traderSheet.getRange(row, config.saleConfig.purchasePriceCol).setValue(purchasePrice);
  traderSheet.getRange(row, config.saleConfig.profitCol1).setValue(profitUSD);
  traderSheet.getRange(row, config.saleConfig.profitCol2).setValue(profitVND);
  Logger.log("updateItemsSaleRecordSale: Đã cập nhật dữ liệu cho item " + itemString + " tại Trader1 row " + row);
}

/* =============================================================================
 * [7] Authentication & Trader File Setup Functions
 * =============================================================================
 */

/**
 * 7.1 authenticateUser: Xác thực đăng nhập.
 * @param {string} userID
 * @param {string} userPassword
 * @return {string|null} Role nếu hợp lệ, null nếu không.
 */
function getAccountsMap() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("accountsMap");
  if (cached) return JSON.parse(cached);
  
  var ssDatabase = SpreadsheetApp.openById(DATABASE_ID);
  var sheetAcc = ssDatabase.getSheetByName("Accounts");
  if (!sheetAcc) return {};
  
  var data = sheetAcc.getRange(2, 12, sheetAcc.getLastRow() - 1, 3).getValues();
  var accountsMap = {};
  data.forEach(function(row) {
    var id = row[0].toString().trim();
    if (id) {
      accountsMap[id] = {
        password: row[1].toString().trim(),
        role: row[2].toString().trim()
      };
    }
  });
  cache.put("accountsMap", JSON.stringify(accountsMap), 300); // cache 5 phút
  return accountsMap;
}

function authenticateUserOptimized(userID, userPassword) {
  var accountsMap = getAccountsMap();
  var account = accountsMap[userID.trim()];
  if (account && userPassword.trim() === account.password) {
    return account.role;
  }
  return null;
}

/**
 * 7.2 setupTraderFileByRole: Cấu hình file Trader theo role.
 * @param {string} fileId
 * @param {Object} config - { visibleStart, visibleEnd }
 */
function setupTraderFileByRole(fileId, config) {
  if (!fileId) {
    throw new Error("File ID không hợp lệ. Vui lòng kiểm tra lại trong Database.");
  }
  var ss = SpreadsheetApp.openById(fileId);
  if (!ss) {
    throw new Error("Không thể mở Spreadsheet với File ID: " + fileId);
  }
  var sheets = ss.getSheets();
  if (!sheets || sheets.length === 0) {
    throw new Error("Không tìm thấy bất kỳ sheet nào trong Spreadsheet với File ID: " + fileId);
  }
  var sheet = ss.getActiveSheet() || sheets[0];
  if (!sheet) {
    throw new Error("Không thể lấy được active sheet hoặc sheet đầu tiên từ Spreadsheet với File ID: " + fileId);
  }
  Logger.log("Đã mở sheet: " + sheet.getName() + " từ file ID: " + fileId);
  
  if (!config.visibleStart || !config.visibleEnd) {
    throw new Error("Cấu hình visibleStart và visibleEnd không hợp lệ.");
  }
  
  sheet.showColumns(config.visibleStart, config.visibleEnd - config.visibleStart + 1);
}

/**
 * 7.3 resetFileSettings: Đưa file Spreadsheet về trạng thái mặc định.
 */
function resetFileSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Không thể lấy được Spreadsheet hiện hành.");
  }
  
  var sheet = ss.getActiveSheet();
  if (!sheet) {
    throw new Error("Không tìm thấy active sheet trong Spreadsheet.");
  }
  
  var maxColumns = sheet.getMaxColumns();
  sheet.showColumns(1, maxColumns);
  
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(function(prot) {
    try {
      prot.remove();
    } catch (e) {
      Logger.log("Lỗi khi xóa bảo vệ: " + e.message);
    }
  });
  
  Logger.log("resetFileSettings: File đã được đưa về trạng thái mặc định.");
}


/* =============================================================================
 * [8] Utility Functions for Items
 * =============================================================================
 */

/**
 * 8.1 combineItemString: Ghép 4 cột Items thành chuỗi theo định dạng "ItemType|attr1|attr2|attr3".
 * @param {Array} itemArray - [ItemType, attr1, attr2, attr3]
 * @return {string} Chuỗi kết hợp.
 */
function combineItemString(itemArray) {
  if (!itemArray || itemArray.length < 4) return "";
  
  var itemType = itemArray[0].toString().trim();
  var attrs = [];
  for (var i = 1; i < 4; i++) {
    attrs.push(itemArray[i].toString().trim());
  }
  
  if (!itemType.toUpperCase().startsWith("U")) {
    attrs.sort(function(a, b) {
      return a.localeCompare(b, 'vi', { sensitivity: 'base' });
    });
  }
  
  return itemType + "|" + attrs.join("|");
}


/* =============================================================================
 * [9] Login & Process Functions
 * =============================================================================
 */

/**
 * 9.1 getTraderFileIdByRole: Lấy FileID của file Trader dựa trên role.
 * @param {string} role
 * @return {string} FileID.
 */
function getTraderFileIdByRole(role) {
  var fileName = role.charAt(0).toUpperCase() + role.slice(1);
  return getFileIdOptimized(fileName);
}

/**
 * 9.2 processLogin: Xác thực đăng nhập, cập nhật trạng thái online và cấu hình file Trader.
 * @param {string} userID
 * @param {string} userPassword
 * @return {Object|string} { role, fileId } nếu thành công, thông báo lỗi nếu không.
 */
function processLoginOptimized(userID, userPassword) {
  Logger.log("processLoginOptimized bắt đầu với userID: " + userID);
  var role = authenticateUserOptimized(userID, userPassword);
  Logger.log("authenticateUserOptimized trả về role: " + role);
  
  if (!role) {
    Logger.log("Đăng nhập thất bại: ID hoặc mật khẩu không đúng");
    return "ID hoặc mật khẩu không đúng!";
  }
  
  manualUpdateTraderStatus(role);
  Logger.log("Cập nhật trạng thái online cho role: " + role);
  
  var traderFileId = getFileIdOptimized(role.charAt(0).toUpperCase() + role.slice(1));
  Logger.log("Lấy được traderFileId: " + traderFileId);
  if (!traderFileId) throw new Error("Không tìm thấy file làm việc cho role " + role);
  
  var traderSpreadsheet = SpreadsheetApp.openById(traderFileId);
  var roleConfig = ROLE_CONFIG[role.toLowerCase()];
  if (!roleConfig) {
    throw new Error("Không tìm thấy cấu hình cho role " + role);
  }
  setupTraderFileByRole(traderFileId, roleConfig);
  Logger.log("Đã cấu hình file Trader cho role " + role);
  
  PropertiesService.getUserProperties().setProperty("userRole", role);
  Logger.log("processLoginOptimized hoàn tất. Trả về đối tượng với role và fileId.");
  
  return { role: role, fileId: traderFileId };
}


/* =============================================================================
 * [10] Trader2 Online & Accounts Dropdown Functions
 * =============================================================================
 */

/**
 * 10.1 getOnlineTrader2Roles: Lấy danh sách các role của Trader2 đang online.
 * @return {Array} Mảng role.
 */
function getOnlineTrader2Roles() {
  var ssDatabase = SpreadsheetApp.openById(DATABASE_ID);
  var onlineStatusSheet = ssDatabase.getSheetByName("OnlineStatus");
  if (!onlineStatusSheet) {
    Logger.log("Không tìm thấy sheet OnlineStatus");
    return [];
  }
  
  var data = onlineStatusSheet.getDataRange().getValues();
  var onlineTrader2Roles = [];
  for (var i = 1; i < data.length; i++) {
    var role = data[i][0].toString().trim().toLowerCase();
    var status = data[i][1].toString().trim().toLowerCase();
    if (role.indexOf("trader2") === 0 && status === "online") {
      onlineTrader2Roles.push(role);
    }
  }
  return onlineTrader2Roles;
}

/**
 * 10.2 getTrader2AccountsForDropdown: Lấy danh sách tài khoản từ sheet "CpAccounts" theo role Trader2 online.
 * @param {Array} onlineRoles
 * @return {Array} Mảng tài khoản.
 */
function getTrader2AccountsForDropdown(onlineRoles) {
  if (!onlineRoles || !Array.isArray(onlineRoles)) {
    onlineRoles = getOnlineTrader2Roles();
  }
  if (onlineRoles.length === 0) {
    SpreadsheetApp.getUi().alert("Không có Trader2 nào online.");
    return [];
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cpSheet = ss.getSheetByName("CpAccounts");
  if (!cpSheet) {
    Logger.log("Không tìm thấy sheet CpAccounts");
    return [];
  }
  
  var lastRow = cpSheet.getLastRow();
  var data = cpSheet.getRange(2, 3, lastRow - 1, 2).getValues();
  var accounts = [];
  
  for (var i = 0; i < data.length; i++) {
    var traderRole = data[i][0].toString().trim().toLowerCase();
    var accountValue = data[i][1].toString().trim();
    if ((onlineRoles.indexOf(traderRole) !== -1 || traderRole === "all") && accountValue !== "") {
      accounts.push(accountValue);
    }
  }
  return accounts;
}

/**
 * 10.3 applyGoldAccountDropdown: Áp dụng Data Validation cho ô account của bảng Gold.
 * @param {Sheet} sheet
 * @param {number} row
 * @param {Object} config - { checkRange: {startCol, endCol}, accountCol }
 */
function applyGoldAccountDropdown(sheet, row, config) {
  if (row < 11) return;
  
  var condRange = sheet.getRange(row, config.checkRange.startCol, 1, config.checkRange.endCol - config.checkRange.startCol + 1);
  var condValues = condRange.getValues()[0];
  
  var isDataComplete = condValues.every(function(cell) {
    return cell.toString().trim() !== "";
  });
  
  var accountRange = sheet.getRange(row, config.accountCol);
  if (!isDataComplete) {
    Logger.log("applyGoldAccountDropdown: Vùng điều kiện không đầy đủ ở row " + row + ". Xóa nội dung và dropdown.");
    accountRange.clearContent().clearDataValidations();
    return;
  } else {
    if (accountRange.getDataValidation()) {
      Logger.log("applyGoldAccountDropdown: Ô account tại row " + row + " đã có dropdown, giữ nguyên.");
      return;
    }
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fileName = ss.getName().toLowerCase();
  var accountList;
  
  if (fileName.indexOf("trader1") === 0) {
    var onlineRoles = getOnlineTrader2Roles();
    if (onlineRoles.length === 0) {
      SpreadsheetApp.getUi().alert("Không có Trader2 nào online.");
      return;
    }
    accountList = getTrader2AccountsForDropdown(onlineRoles);
  } else if (fileName.indexOf("trader2") === 0) {
    // Với file Trader2 (ví dụ "trader21"), lấy danh sách tài khoản dựa theo tên file
    accountList = getTrader2AccountsForDropdown([fileName]);
  } else {
    accountList = [];
  }
  
  if (accountList && accountList.length > 0) {
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(accountList, true).build();
    accountRange.setDataValidation(rule);
    Logger.log("applyGoldAccountDropdown: Đã đặt dropdown với dữ liệu: " + accountList.join(", "));
  } else {
    Logger.log("applyGoldAccountDropdown: Không có dữ liệu hợp lệ.");
  }
}

/**
 * 10.4 applyItemsAccountDropdown: Áp dụng Data Validation cho ô account của bảng Items.
 * @param {Sheet} sheet
 * @param {number} row
 * @param {Object} config - { checkRange: {startCol, endCol}, accountCol }
 */
function applyItemsAccountDropdown(sheet, row, config) {
  if (row < 11) return;
  
  // Lấy vùng dữ liệu kiểm tra (checkRange)
  var condRange = sheet.getRange(row, config.checkRange.startCol, 1, config.checkRange.endCol - config.checkRange.startCol + 1);
  var condValues = condRange.getValues()[0];
  
  // Kiểm tra tất cả các ô trong vùng đã có dữ liệu chưa
  var isDataComplete = condValues.every(function(cell) {
    return cell.toString().trim() !== "";
  });
  
  var accountRange = sheet.getRange(row, config.accountCol);
  if (!isDataComplete) {
    Logger.log("applyItemsAccountDropdown: Vùng điều kiện không đầy đủ ở row " + row + ". Xóa dropdown account.");
    accountRange.clearContent().clearDataValidations();
    return;
  }
  if (accountRange.getDataValidation()) {
    Logger.log("applyItemsAccountDropdown: Ô account tại row " + row + " đã có dropdown, giữ nguyên.");
    return;
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fileName = ss.getName().toLowerCase();
  var accountList = [];
  
  // Nếu file là Trader1: sử dụng role của Trader2 từ OnlineStatus
  if (fileName.indexOf("trader1") === 0) {
    var onlineRoles = getOnlineTrader2Roles();
    if (onlineRoles.length === 0) {
      SpreadsheetApp.getUi().alert("Không có Trader2 nào online.");
      return;
    }
    // Lọc danh sách account từ CpMatrixAccounts dựa vào các role trong onlineRoles
    var cpSheet = ss.getSheetByName("CpMatrixAccounts");
    if (!cpSheet) {
      Logger.log("applyItemsAccountDropdown: Không tìm thấy sheet CpMatrixAccounts");
      return;
    }
    var lastRowCP = cpSheet.getLastRow();
    if (lastRowCP < 3) {
      Logger.log("applyItemsAccountDropdown: Không có dữ liệu từ hàng 3 trở đi trong CpMatrixAccounts.");
      return;
    }
    var lastCol = cpSheet.getLastColumn();
    var data = cpSheet.getRange(3, 1, lastRowCP - 2, lastCol).getValues();
    for (var i = 0; i < data.length; i++) {
      var rowData = data[i];
      var rowRole = rowData[0].toString().trim().toLowerCase();
      // Nếu role trong dòng là một trong số các role online hoặc "all"
      if (onlineRoles.indexOf(rowRole) !== -1 || rowRole === "all") {
        var effectiveLastIndex = -1;
        for (var j = rowData.length - 1; j >= 2; j--) {
          if (rowData[j].toString().trim() !== "") {
            effectiveLastIndex = j;
            break;
          }
        }
        if (effectiveLastIndex >= 2) {
          var rowValues = rowData.slice(2, effectiveLastIndex + 1).filter(function(cell) {
            return cell.toString().trim() !== "";
          });
          accountList = accountList.concat(rowValues);
        }
      }
    }
    accountList = accountList.filter(function(item, index, self) {
      return self.indexOf(item) === index;
    });
  }
  // Nếu file là Trader2: sử dụng tên file (role) để lọc danh sách account từ CpMatrixAccounts
  else if (fileName.indexOf("trader2") === 0) {
    var role = fileName; // Sử dụng chính tên file làm role
    var cpSheet = ss.getSheetByName("CpMatrixAccounts");
    if (!cpSheet) {
      Logger.log("applyItemsAccountDropdown: Không tìm thấy sheet CpMatrixAccounts");
      return;
    }
    var lastRowCP = cpSheet.getLastRow();
    if (lastRowCP < 3) {
      Logger.log("applyItemsAccountDropdown: Không có dữ liệu từ hàng 3 trở đi trong CpMatrixAccounts.");
      return;
    }
    var lastCol = cpSheet.getLastColumn();
    var data = cpSheet.getRange(3, 1, lastRowCP - 2, lastCol).getValues();
    for (var i = 0; i < data.length; i++) {
      var rowData = data[i];
      var rowRole = rowData[0].toString().trim().toLowerCase();
      if (rowRole === role || rowRole === "all") {
        var effectiveLastIndex = -1;
        for (var j = rowData.length - 1; j >= 2; j--) {
          if (rowData[j].toString().trim() !== "") {
            effectiveLastIndex = j;
            break;
          }
        }
        if (effectiveLastIndex >= 2) {
          var rowValues = rowData.slice(2, effectiveLastIndex + 1).filter(function(cell) {
            return cell.toString().trim() !== "";
          });
          accountList = accountList.concat(rowValues);
        }
      }
    }
    accountList = accountList.filter(function(item, index, self) {
      return self.indexOf(item) === index;
    });
  } else {
    accountList = [];
  }
  
  if (!accountList || accountList.length === 0) {
    Logger.log("applyItemsAccountDropdown: Không có dữ liệu hợp lệ. Xóa dropdown account.");
    accountRange.clearContent().clearDataValidations();
    return;
  }
  
  accountRange.clearContent();
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(accountList, true).build();
  accountRange.setDataValidation(rule);
  Logger.log("applyItemsAccountDropdown: Đã đặt dropdown với dữ liệu: " + accountList.join(", "));
}

/**
 * 10.5 getTrader2FileIdFromAccountPrefix: Lấy file ID của Trader2 từ account prefix.
 * Nếu trader identifier là "all", chọn trader2 online gần nhất.
 * @param {string} accountPrefix
 * @return {string|null}
 */
function getTrader2FileIdFromAccountPrefix(accountPrefix) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cpSheet = ss.getSheetByName("MatrixAccounts");
  if (!cpSheet) {
    Logger.log("Không tìm thấy sheet CpMatrixAccounts");
    return null;
  }
  
  var lastRow = cpSheet.getLastRow();
  var range = cpSheet.getRange(3, 2, lastRow - 2, 1);
  var values = range.getValues();
  var traderIdentifier = null;
  
  for (var i = 0; i < values.length; i++) {
    if (values[i][0].toString().trim().toLowerCase() === accountPrefix.toLowerCase()) {
      traderIdentifier = cpSheet.getRange(i + 3, 1).getValue().toString().trim();
      break;
    }
  }
  
  if (!traderIdentifier) {
    Logger.log("Không tìm thấy trader identifier cho account prefix: " + accountPrefix);
    return null;
  }
  
  if (traderIdentifier.toLowerCase() === "all") {
    var closestRole = getClosestOnlineTrader2Role();
    if (!closestRole) {
      Logger.log("Không tìm thấy Trader2 online cho account thuộc 'all'");
      return null;
    }
    traderIdentifier = closestRole;
  }
  
  var traderFileId = getFileIdOptimized(traderIdentifier);
  return traderFileId;
}

/**
 * 10.6 getClosestOnlineTrader2Role: Lấy role của Trader2 online gần nhất (theo thời gian cập nhật mới nhất).
 * @return {string|null}
 */
function getClosestOnlineTrader2Role() {
  var ssDatabase = SpreadsheetApp.openById(DATABASE_ID);
  var sheet = ssDatabase.getSheetByName("OnlineStatus");
  if (!sheet) {
    Logger.log("Không tìm thấy sheet OnlineStatus");
    return null;
  }
  
  var data = sheet.getDataRange().getValues();
  var closestRole = null;
  var latestTimestamp = 0;
  for (var i = 1; i < data.length; i++) {
    var role = data[i][0].toString().trim().toLowerCase();
    var status = data[i][1].toString().trim().toLowerCase();
    var timestamp = new Date(data[i][2]).getTime();
    if (role.indexOf("trader2") === 0 && status === "online" && timestamp > latestTimestamp) {
      latestTimestamp = timestamp;
      closestRole = role;
    }
  }
  return closestRole;
}


/* =============================================================================
 * [11] Push Data Functions
 * =============================================================================
 */

/**
 * deductGoldSale: Trừ số lượng gold bán khỏi file Kho.
 * @param {Sheet} traderSheet - Sheet của Trader1 chứa thông tin bán gold.
 * @param {number} row - Dòng trên sheet Trader chứa thông tin bán gold.
 * @param {number} saleQtyCol - Cột chứa số lượng gold bán.
 * @param {number} accountCol - Cột chứa thông tin account (ví dụ "Acc05-Trader21").
 * @param {Object} khoConfig - Cấu hình cho file Kho gồm:
 *        sheetName, rowData (dòng chứa dữ liệu Gold).
 */
function deductGoldSale(traderSheet, row, saleQtyCol, accountCol, khoConfig) {
  var ui = SpreadsheetApp.getUi();
  
  // 1. Lấy số lượng bán từ Trader sheet
  var saleQty = Number(traderSheet.getRange(row, saleQtyCol).getValue());
  if (isNaN(saleQty) || saleQty <= 0) {
    ui.alert("deductGoldSale: Số lượng bán không hợp lệ tại dòng " + row);
    throw new Error("deductGoldSale: Số lượng bán không hợp lệ tại dòng " + row);
  }
  
  // 2. Lấy thông tin account từ Trader sheet
  var traderAccount = traderSheet.getRange(row, accountCol).getValue().toString().trim();
  if (!traderAccount) {
    ui.alert("deductGoldSale: Thiếu thông tin account tại dòng " + row);
    throw new Error("deductGoldSale: Thiếu thông tin account tại dòng " + row);
  }
  var prefix = traderAccount.indexOf("-") > -1 ? traderAccount.split("-")[0] : traderAccount;
  var desiredHeader1 = prefix + "-Helm";
  var desiredHeader2 = prefix + "-Spirit1";
  
  // 3. Mở file Kho và lấy sheet theo khoConfig
  var khoId = getFileIdOptimized("Kho");
  if (!khoId) {
    ui.alert("deductGoldSale: Không lấy được ID của file Kho từ mapping.");
    throw new Error("deductGoldSale: Không lấy được ID của file Kho từ mapping.");
  }
  var ssKho = SpreadsheetApp.openById(khoId);
  var khoSheet = ssKho.getSheetByName(khoConfig.sheetName);
  if (!khoSheet) {
    ui.alert("deductGoldSale: Không tìm thấy sheet '" + khoConfig.sheetName + "' trong file Kho.");
    throw new Error("deductGoldSale: Không tìm thấy sheet '" + khoConfig.sheetName + "' trong file Kho.");
  }
  
  var khoRow = khoConfig.rowData; // Dòng chứa dữ liệu gold trong Kho
  
  // 4. Kiểm tra tổng số lượng trong Kho
  var totalCell = khoSheet.getRange(khoRow, khoConfig.qtyCol);
  var currentTotal = Number(totalCell.getValue());
  if (isNaN(currentTotal)) currentTotal = 0;
  var newTotal = currentTotal - saleQty;
  if (newTotal < 0) {
    ui.alert("deductGoldSale: Số lượng bán vượt quá tổng số gold hiện có.");
    throw new Error("deductGoldSale: Số lượng bán vượt quá tổng số gold hiện có.");
  }
  
  // 5. Tìm cột của account trong Kho dựa trên header
  var headerValues = khoSheet.getRange(1, 1, 1, khoSheet.getLastColumn()).getValues()[0];
  var targetCol = null;
  var desiredHeader = desiredHeader1;
  for (var i = 0; i < headerValues.length; i++) {
    if (headerValues[i].toString().trim().toUpperCase() === desiredHeader1.toUpperCase()) {
      targetCol = i + 1;
      break;
    }
  }
  if (!targetCol) {
    for (var i = 0; i < headerValues.length; i++) {
      if (headerValues[i].toString().trim().toUpperCase() === desiredHeader2.toUpperCase()) {
        targetCol = i + 1;
        desiredHeader = desiredHeader2;
        break;
      }
    }
  }
  if (!targetCol) {
    ui.alert("deductGoldSale: Không tìm thấy cột cho account " + desiredHeader1 + " hoặc " + desiredHeader2 + " trong file Kho.");
    throw new Error("deductGoldSale: Không tìm thấy cột cho account " + desiredHeader1 + " hoặc " + desiredHeader2 + " trong file Kho.");
  }
  
  // 6. Kiểm tra số lượng gold trên account
  var accCell = khoSheet.getRange(khoRow, targetCol);
  var currentAccQty = Number(accCell.getValue());
  if (isNaN(currentAccQty)) currentAccQty = 0;
  var newAccQty = currentAccQty - saleQty;
  if (newAccQty < 0) {
    ui.alert("deductGoldSale: Số lượng bán vượt quá số lượng gold trên account " + desiredHeader + ".");
    throw new Error("deductGoldSale: Số lượng bán vượt quá số lượng gold trên account " + desiredHeader + ".");
  }
  
  // 7. Nếu tất cả các kiểm tra đều qua, cập nhật dữ liệu
  totalCell.setValue(newTotal);
  accCell.setValue(newAccQty);
  
  Logger.log("deductGoldSale: Đã trừ " + saleQty + " gold. Tổng gold mới = " + newTotal + ", gold trên " + desiredHeader + " mới = " + newAccQty);
}

/**
 * pushGoldDataFromTrader1Optimized:
 * Đẩy dữ liệu bảng Gold từ Trader1 sang Trader2.
 * Sau khi đẩy thành công, đánh dấu dòng là "Pushed" và khóa vùng dữ liệu.
 * Sử dụng DocumentLock để đảm bảo các lần ghi dữ liệu được thực hiện tuần tự.
 *
 * Tham số:
 *   row: số thứ tự dòng dữ liệu trên sheet Trader1.
 *   config: đối tượng cấu hình chứa các trường:
 *           {
 *             checkRange: { startCol, endCol },
 *             accountCol: cột chứa thông tin Account (ví dụ: "AccXX-TraderYY"),
 *             sourceDataRange: { startCol, endCol },  // vùng dữ liệu cần đẩy từ Trader1
 *             destStartRow: hàng bắt đầu của bảng trong file Trader2,
 *             destDataStartCol: cột bắt đầu ghi dữ liệu trong file Trader2,
 *             headerRow (tùy chọn): nếu file Trader2 có tiêu đề,
 *             trader2FileId: (nếu có) file ID của Trader2; nếu không có thì hàm sẽ tự xác định qua getTrader2FileIdFromAccountPrefix.
 *           }
 *   nextInsertionRow: (tùy chọn) vị trí chèn hiện tại đã được tính trước (để xử lý nhiều dòng đẩy cùng lúc).
 *
 * Trả về:
 *   nextInsertionRow mới (giả sử mỗi lần push thêm 1 dòng)
 */
function pushGoldDataFromTrader1Optimized(row, config, nextInsertionRow) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Chờ tối đa 30 giây để lấy khóa

    var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var pushedCol = config.checkRange.endCol + 1;
    if (sourceSheet.getRange(row, pushedCol).getValue().toString().trim() === "Pushed") {
      Logger.log("pushGoldData: Row " + row + " đã được đẩy (đã có 'Pushed').");
      return nextInsertionRow;
    }
    
    // Kiểm tra dữ liệu đầy đủ sử dụng hàm tiện ích
    if (!isDataComplete(sourceSheet, row, config.checkRange.startCol, config.checkRange.endCol)) {
      Logger.log("pushGoldData: Dữ liệu không đầy đủ ở dòng " + row);
      return nextInsertionRow;
    }
    
    // Lấy thông tin account để xác định accountPrefix
    var accountPrefix = getAccountPrefix(sourceSheet, row, config.accountCol);
    if (!accountPrefix) {
      Logger.log("pushGoldData: Không xác định được account prefix ở dòng " + row);
      return nextInsertionRow;
    }
    
    // Xác định file Trader2: Nếu chưa được cung cấp thì tìm qua mapping dựa trên accountPrefix
    // Lấy tên file Trader2 (ví dụ "Trader21")
    var trader2FileName = config.trader2FileId;
    if (!trader2FileName) {
      trader2FileName = getTargetFileId(accountPrefix, "MatrixAccounts");
      if (!trader2FileName) {
        Logger.log("pushGoldData: Không tìm thấy file Trader2 cho account prefix " + accountPrefix);
        return nextInsertionRow;
      }
    }
    var trader2FileId = getFileIdOptimized(trader2FileName);
    if (!trader2FileId) {
      Logger.log("pushGoldData: Không lấy được file ID từ tên file " + trader2FileName);
      return nextInsertionRow;
    }
    
    Logger.log("pushGoldData: trader2FileId = '" + trader2FileId + "'");
    var destSpreadsheet = SpreadsheetApp.openById(trader2FileId);
    var destSheetName = sourceSheet.getName();
    var destSheet = destSpreadsheet.getSheetByName(destSheetName);
    if (!destSheet) {
      Logger.log("pushGoldData: Không tìm thấy sheet " + destSheetName + " trong file Trader2");
      return nextInsertionRow;
    }
    
    // Xác định vị trí chèn (insertionRow)
    var headerRow = config.headerRow || config.destStartRow;
    var insertionRow;
    if (typeof nextInsertionRow === 'number') {
      insertionRow = nextInsertionRow;
    } else {
      var lastRow = destSheet.getLastRow();
      if (lastRow <= headerRow) {
        insertionRow = headerRow + 1;
      } else {
        var numRowsToCheck = lastRow - headerRow;
        var colValues = destSheet.getRange(headerRow + 1, config.destDataStartCol, numRowsToCheck, 1).getValues();
        var foundBlank = false;
        for (var i = 0; i < colValues.length; i++) {
          if (colValues[i][0] === "" || colValues[i][0] == null) {
            insertionRow = headerRow + 1 + i;
            foundBlank = true;
            break;
          }
        }
        if (!foundBlank) {
          insertionRow = lastRow + 1;
        }
      }
    }
    
    // Lấy dữ liệu cần đẩy từ Trader1
    var sourceDataRange = sourceSheet.getRange(row, config.sourceDataRange.startCol, 1, config.sourceDataRange.endCol - config.sourceDataRange.startCol + 1);
    var dataToPush = sourceDataRange.getValues();
    
    // Sau khi đẩy dữ liệu, gọi deductGoldSale để trừ số lượng gold bán khỏi Kho.
    // Các tham số: saleQtyCol, accountCol từ goldPush, và kho config từ goldPush.kho.
    try {
      if (config.saleQtyCol && config.accountCol && config.kho) {
        deductGoldSale(sourceSheet, row, config.saleQtyCol, config.accountCol, config.kho);
      }
    } catch(e) {
      throw new Error("pushGoldDataFromTrader1Optimized: deductGoldSale failed: " + e.message);
    }

    // Ghi dữ liệu vào file Trader2 tại vị trí chèn tính được
    destSheet.getRange(insertionRow, config.destDataStartCol, 1, dataToPush[0].length).setValues(dataToPush);
    SpreadsheetApp.flush();
    Logger.log("pushGoldData: Đã đẩy dữ liệu từ Trader1 (row " + row + ") sang Trader2 tại dòng " + insertionRow);
    
    
    // Đánh dấu dòng đã đẩy và khóa vùng dữ liệu đã đẩy
    sourceSheet.getRange(row, pushedCol).setValue("Pushed");
    applyProtection(sourceSheet, row, config.checkRange.startCol, config.checkRange.endCol + 1, "Data pushed and locked by admin");
    
    return insertionRow + 1;
  } catch (e) {
    Logger.log("pushGoldData error: " + e);
    return nextInsertionRow;
  } finally {
    lock.releaseLock();
  }
}

/**
 * deductItemsSale: Trừ số lượng bán cho record Items trong file Kho.
 * Đối với Items, số lượng bán luôn = 1.
 * 
 * Yêu cầu: itemsPushConfig phải chứa:
 *    - sourceDataRange: { startCol, endCol } (phạm vi chứa đúng 4 cột thông tin item)
 *    - kho: { sheetName, rowData, itemsCol, qtyCol }
 * 
 * @param {Sheet} traderSheet - Sheet của Trader chứa thông tin đơn hàng Items.
 * @param {number} row - Dòng trên traderSheet cần xử lý.
 * @param {number} accountCol - Cột chứa thông tin Account (ví dụ "AccXX-TraderYY") dùng để xác định cột trong Kho.
 * @param {Object} itemsPushConfig - Cấu hình cho Items push, bao gồm:
 *        sourceDataRange: { startCol, endCol },
 *        kho: { sheetName, rowData, itemsCol, qtyCol }
 */
function deductItemsSale(traderSheet, row, accountCol, itemsPushConfig) {
  var ui = SpreadsheetApp.getUi();
  
  // Với Items, saleQty mặc định = 1
  var saleQty = 1;
  
  // 1. Kiểm tra sourceDataRange phải chứa đúng 4 cột
  var numCols = itemsPushConfig.sourceDataRange.endCol - itemsPushConfig.sourceDataRange.startCol + 1;
  if (numCols !== 4) {
    ui.alert("deductItemsSale: Phạm vi dữ liệu item phải chứa đúng 4 cột, hiện tại: " + numCols);
    throw new Error("deductItemsSale: Phạm vi dữ liệu item phải chứa đúng 4 cột, hiện tại: " + numCols);
  }
  
  // 2. Lấy chuỗi item từ Trader sheet dựa trên sourceDataRange
  var sourceDataRange = traderSheet.getRange(row, itemsPushConfig.sourceDataRange.startCol, 1, numCols);
  var itemData = sourceDataRange.getValues()[0];
  var itemString = combineItemString(itemData);
  if (!itemString) {
    ui.alert("deductItemsSale: Chuỗi item rỗng ở dòng " + row);
    throw new Error("deductItemsSale: Chuỗi item rỗng ở dòng " + row);
  }
  
  // 3. Mở file Kho và lấy sheet theo itemsPushConfig.kho
  var khoId = getFileIdOptimized("Kho");
  if (!khoId) {
    ui.alert("deductItemsSale: Không lấy được ID của file Kho từ mapping.");
    throw new Error("deductItemsSale: Không lấy được ID của file Kho từ mapping.");
  }
  var ssKho = SpreadsheetApp.openById(khoId);
  var khoSheet = ssKho.getSheetByName(itemsPushConfig.kho.sheetName);
  if (!khoSheet) {
    ui.alert("deductItemsSale: Không tìm thấy sheet '" + itemsPushConfig.kho.sheetName + "' trong file Kho.");
    throw new Error("deductItemsSale: Không tìm thấy sheet '" + itemsPushConfig.kho.sheetName + "' trong file Kho.");
  }
  
  // 4. Tìm dòng record của item trong Kho bằng cách quét từ itemsPushConfig.kho.rowData đến cuối file
  var startRowKho = itemsPushConfig.kho.rowData;
  var itemsCol = itemsPushConfig.kho.itemsCol; // Cột chứa chuỗi item trong Kho
  var lastRowKho = khoSheet.getLastRow();
  var foundRow = null;
  if (lastRowKho >= startRowKho) {
    var itemsRangeKho = khoSheet.getRange(startRowKho, itemsCol, lastRowKho - startRowKho + 1, 1);
    var itemsValues = itemsRangeKho.getValues();
    for (var i = 0; i < itemsValues.length; i++) {
      var currentItem = itemsValues[i][0].toString().trim();
      if (itemString === currentItem) {
        foundRow = startRowKho + i;
        break;
      }
    }
  }
  if (!foundRow) {
    ui.alert("deductItemsSale: Không tìm thấy record cho item: " + itemString);
    throw new Error("deductItemsSale: Không tìm thấy record cho item: " + itemString);
  }
  
  // 5. Cập nhật tổng số lượng trong Kho (ở cột qty)
  var totalCell = khoSheet.getRange(foundRow, itemsPushConfig.kho.qtyCol);
  var currentTotal = Number(totalCell.getValue());
  if (isNaN(currentTotal)) currentTotal = 0;
  var newTotal = currentTotal - saleQty;
  if (newTotal < 0) {
    ui.alert("deductItemsSale: Số lượng bán vượt quá tổng số items hiện có.");
    throw new Error("deductItemsSale: Số lượng bán vượt quá tổng số items hiện có.");
  }
  
  // 6. Xác định cột account trong Kho dựa trên header (so sánh trực tiếp với giá trị account)
  var traderAccount = traderSheet.getRange(row, accountCol).getValue().toString().trim();
  if (!traderAccount) {
    ui.alert("deductItemsSale: Thiếu thông tin account tại dòng " + row);
    throw new Error("deductItemsSale: Thiếu thông tin account tại dòng " + row);
  }
  var desiredHeader = traderAccount; // Với Items, dùng account trực tiếp
  var headerValues = khoSheet.getRange(1, 1, 1, khoSheet.getLastColumn()).getValues()[0];
  var targetCol = null;
  for (var j = 0; j < headerValues.length; j++) {
    if (headerValues[j].toString().trim().toLowerCase() === desiredHeader.toLowerCase()) {
      targetCol = j + 1;
      break;
    }
  }
  if (!targetCol) {
    ui.alert("deductItemsSale: Không tìm thấy cột account cho '" + desiredHeader + "' trong file Kho.");
    throw new Error("deductItemsSale: Không tìm thấy cột account cho '" + desiredHeader + "' trong file Kho.");
  }
  
  // 7. Trừ số lượng trong cột account
  var accCell = khoSheet.getRange(foundRow, targetCol);
  var currentAccQty = Number(accCell.getValue());
  if (isNaN(currentAccQty)) currentAccQty = 0;
  var newAccQty = currentAccQty - saleQty;
  if (newAccQty < 0) {
    ui.alert("deductItemsSale: Số lượng bán vượt quá số lượng items trên account '" + desiredHeader + "'.");
    throw new Error("deductItemsSale: Số lượng bán vượt quá số lượng items trên account '" + desiredHeader + "'.");
  }
  
  // 8. Nếu tất cả các kiểm tra đều qua, cập nhật dữ liệu
  totalCell.setValue(newTotal);
  accCell.setValue(newAccQty);
  
  Logger.log("deductItemsSale: Đã trừ 1 item. Tổng items mới = " + newTotal + ", items trên account '" + desiredHeader + "' mới = " + newAccQty);
}

/**
 * pushItemsDataFromTrader1Optimized:
 * Đẩy dữ liệu bảng Items từ Trader1 sang Trader2.
 * Sau khi đẩy thành công, đánh dấu dòng là "Pushed" và khóa vùng dữ liệu.
 * Sử dụng DocumentLock của file Trader2 để đảm bảo rằng các lần ghi dữ liệu được thực hiện tuần tự.
 *
 * Tham số:
 *   row: số thứ tự dòng dữ liệu trên sheet Trader1.
 *   config: đối tượng cấu hình chứa các trường:
 *           {
 *             checkRange: { startCol, endCol },
 *             accountCol: cột chứa Account (ví dụ: "Acc01-TraderXX"),
 *             sourceDataRange: { startCol, endCol },  // vùng dữ liệu cần đẩy từ Trader1
 *             destStartRow: hàng bắt đầu của bảng trong file Trader2,
 *             destDataStartCol: cột bắt đầu ghi dữ liệu trong file Trader2,
 *             pushedCol: (nếu không truyền, sẽ dùng checkRange.endCol + 1),
 *             trader2FileId: (tùy chọn) nếu có, file ID của file Trader2; nếu không có, sẽ tự xác định từ accountPrefix.
 *           }
 *   nextInsertionRow: (tùy chọn) vị trí chèn hiện tại đã được tính trước (để xử lý nhiều dòng đẩy cùng lúc).
 *
 * Trả về:
 *   nextInsertionRow mới (giả sử mỗi lần push thêm 1 dòng)
 */
function pushItemsDataFromTrader1Optimized(row, config, nextInsertionRow) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Chờ tối đa 30 giây để lấy khóa

    var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var pushedCol = config.pushedCol || (config.checkRange.endCol + 1);
    if (sourceSheet.getRange(row, pushedCol).getValue().toString().trim() === "Pushed") {
      Logger.log("pushItemsData: Row " + row + " đã được đẩy (đã có 'Pushed').");
      return nextInsertionRow;
    }
    
    // Kiểm tra dữ liệu đầy đủ
    if (!isDataComplete(sourceSheet, row, config.checkRange.startCol, config.checkRange.endCol)) {
      Logger.log("pushItemsData: Dữ liệu không đầy đủ ở dòng " + row);
      return nextInsertionRow;
    }
    
    // Lấy thông tin account để xác định accountPrefix
    var accountPrefix = getAccountPrefix(sourceSheet, row, config.accountCol);
    if (!accountPrefix) {
      Logger.log("pushItemsData: Không xác định được account prefix ở dòng " + row);
      return nextInsertionRow;
    }
    
    // Lấy tên file Trader2 và chuyển đổi sang file ID
    var trader2FileName = config.trader2FileId;
    if (!trader2FileName) {
      trader2FileName = getTargetFileId(accountPrefix, "MatrixAccounts");
      if (!trader2FileName) {
        Logger.log("pushItemsData: Không tìm thấy file Trader2 cho account prefix " + accountPrefix);
        return nextInsertionRow;
      }
    }
    var trader2FileId = getFileIdOptimized(trader2FileName);
    if (!trader2FileId) {
      Logger.log("pushItemsData: Không lấy được file ID từ tên file " + trader2FileName);
      return nextInsertionRow;
    }
    
    Logger.log("pushItemsData: trader2FileId = '" + trader2FileId + "'");
    var destSpreadsheet = SpreadsheetApp.openById(trader2FileId);
    var destSheetName = sourceSheet.getName();
    var destSheet = destSpreadsheet.getSheetByName(destSheetName);
    if (!destSheet) {
      Logger.log("pushItemsData: Không tìm thấy sheet " + destSheetName + " trong file Trader2");
      return nextInsertionRow;
    }
    
    // Xác định vị trí chèn (insertionRow)
    var headerRow = config.headerRow || config.destStartRow;
    var insertionRow;
    if (typeof nextInsertionRow === 'number') {
      insertionRow = nextInsertionRow;
    } else {
      var lastRow = destSheet.getLastRow();
      if (lastRow <= headerRow) {
        insertionRow = headerRow + 1;
      } else {
        var numRowsToCheck = lastRow - headerRow;
        var colValues = destSheet.getRange(headerRow + 1, config.destDataStartCol, numRowsToCheck, 1).getValues();
        var foundBlank = false;
        for (var i = 0; i < colValues.length; i++) {
          if (colValues[i][0] === "" || colValues[i][0] == null) {
            insertionRow = headerRow + 1 + i;
            foundBlank = true;
            break;
          }
        }
        if (!foundBlank) {
          insertionRow = lastRow + 1;
        }
      }
    }
    
    // Sau khi đẩy dữ liệu, gọi deductItemsSale cho items
    try {
    if (config.accountCol && config.kho) {
      deductItemsSale(sourceSheet, row, config.accountCol, config);
    }
    } catch(e) {
      throw new Error("pushItemsDataFromTrader1Optimized: deductItemsSale failed: " + e.message);
    }

    // Lấy dữ liệu cần đẩy từ Trader1
    var sourceDataRange = sourceSheet.getRange(row, config.checkRange.startCol, 1, config.checkRange.endCol - config.checkRange.startCol + 1);
    var dataToPush = sourceDataRange.getValues();
    
    // Ghi dữ liệu vào file Trader2 tại vị trí chèn tính được
    destSheet.getRange(insertionRow, config.destDataStartCol, 1, dataToPush[0].length).setValues(dataToPush);
    SpreadsheetApp.flush();
    Logger.log("pushItemsData: Đã đẩy dữ liệu từ Trader1 (row " + row + ") sang Trader2 tại dòng " + insertionRow);
    
    // Đánh dấu dòng đã đẩy và khóa vùng dữ liệu đã đẩy
    sourceSheet.getRange(row, pushedCol).setValue("Pushed");
    applyProtection(sourceSheet, row, config.checkRange.startCol, config.checkRange.endCol + 1, "Data pushed and locked by admin");
    
    return insertionRow + 1;
  } catch (e) {
    Logger.log("pushItemsData error: " + e);
    return nextInsertionRow;
  } finally {
    lock.releaseLock();
  }
}