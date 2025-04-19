/**
 * 1.1 getFileId: Lấy file ID từ file Database dựa trên tên file.
 * @param {string} fileName - Tên file cần lấy ID.
 * @return {string|null} File ID nếu tìm thấy, null nếu không.
 */
function getFileMapping() {
  var cache = CacheService.getScriptCache();
  var cachedMapping = cache.get("fileMapping");
  if (cachedMapping) {
    return JSON.parse(cachedMapping);
  }
  
  var ssDatabase = SpreadsheetApp.openById("1MOEie6MQS3P7tzKYqpbX-tOacN2u0qc2B7hlEshMItc");
  var sheetFileIDs = ssDatabase.getSheetByName("FileIDs");
  if (!sheetFileIDs) {
    Logger.log("Không tìm thấy sheet 'FileIDs'");
    return {};
  }
  
  var data = sheetFileIDs.getDataRange().getValues();
  var mapping = {};
  for (var i = 1; i < data.length; i++) {
    var name = data[i][0].toString().trim().toLowerCase();
    var id = data[i][1].toString().trim();
    if (name) mapping[name] = id;
  }
  cache.put("fileMapping", JSON.stringify(mapping), 300); // cache 5 phút
  return mapping;
}

/**
 * 2.1 isMainSheet: Kiểm tra xem tên sheet có nằm trong khoảng "01" đến "30" hay không.
 * @param {Sheet} sheet
 * @return {boolean} true nếu là main sheet.
 */
function isMainSheet(sheet) {
  var name = sheet.getName();
  return /^(0[1-9]|[12]\d|30)$/.test(name);
}

/**
 * 2.2 isConditionColumn: Kiểm tra xem vùng được chọn có giao với cột conditionIndex không.
 * @param {Range} rng - Vùng được chọn.
 * @param {number} conditionIndex - Số thứ tự cột điều kiện.
 * @return {boolean} true nếu giao, false nếu không.
 */
function isConditionColumn(rng, conditionIndex) {
  var startCol = rng.getColumn();
  var endCol = startCol + rng.getNumColumns() - 1;
  return (startCol <= conditionIndex && endCol >= conditionIndex);
}

function getFileIdOptimized(fileName) {
  var mapping = getFileMapping();
  return mapping[fileName.toString().trim().toLowerCase()] || null;
}

function isDataComplete(sheet, row, startCol, endCol) {
  var values = sheet.getRange(row, startCol, 1, endCol - startCol + 1).getValues()[0];
  return values.every(function(cell) {
    return cell.toString().trim() !== "";
  });
}

function getAccountPrefix(sheet, row, accountCol) {
  var accountVal = sheet.getRange(row, accountCol).getValue().toString().trim();
  if (!accountVal) return null;
  var parts = accountVal.split("-");
  return parts[0] || null;
}

function getTargetFileId(accountPrefix, cpMatrixSheetName) {
  var ss = SpreadsheetApp.openById(DATABASE_ID);  
  var cpSheet = ss.getSheetByName(cpMatrixSheetName);
  if (!cpSheet) {
    Logger.log("Không tìm thấy sheet " + cpMatrixSheetName + " trong Database");
    return null;
  }
  
  var lastRow = cpSheet.getLastRow();
  // Giả sử dữ liệu bắt đầu từ hàng 3, file ID nằm ở cột 1 và account prefix ở cột 2
  for (var i = 3; i <= lastRow; i++) {
    var prefix = cpSheet.getRange(i, 2).getValue().toString().trim();
    if (prefix.toLowerCase() === accountPrefix.toLowerCase()) {
      return cpSheet.getRange(i, 1).getValue().toString().trim();
    }
  }
  return null;
}

function openFileAndSheet(fileKey, sheetName) {
  // Lấy file ID từ mapping bằng getFileIdOptimized
  var fileId = getFileIdOptimized(fileKey);
  if (!fileId) {
    throw new Error("Không tìm thấy file '" + fileKey + "' trong mapping.");
  }
  
  // Mở Spreadsheet dựa trên fileId
  var ss = SpreadsheetApp.openById(fileId);
  if (!ss) {
    throw new Error("Không thể mở Spreadsheet với File ID: " + fileId);
  }
  
  // Lấy sheet theo tên được chỉ định
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error("Không tìm thấy sheet '" + sheetName + "' trong file '" + fileKey + "'.");
  }
  
  // Trả về đối tượng chứa cả Spreadsheet và sheet
  return { spreadsheet: ss, sheet: sheet };
}

function applyProtection(sheet, row, startCol, endCol, description) {
  var range = sheet.getRange(row, startCol, 1, endCol - startCol + 1);
  var protection = range.protect().setDescription(description);
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  return protection;
}

/**
 * formatTimeDifference: Chuyển đổi hiệu số thời gian (ms) sang định dạng:
 * "x ngày: x giờ: x phút: x giây" (nếu ≥ 1 ngày),
 * "x giờ: x phút: x giây" (nếu ≥ 1 giờ nhưng < 1 ngày),
 * "x phút: x giây" (nếu ≥ 1 phút nhưng < 1 giờ),
 * "x giây" (nếu < 1 phút).
 *
 * @param {number} diffMs - Hiệu số thời gian tính theo milliseconds.
 * @return {string} Chuỗi định dạng thời gian.
 */
function formatTimeDifference(diffMs) {
  var totalSeconds = Math.floor(diffMs / 1000);
  var days = Math.floor(totalSeconds / 86400);
  var hours = Math.floor((totalSeconds % 86400) / 3600);
  var minutes = Math.floor((totalSeconds % 3600) / 60);
  var seconds = totalSeconds % 60;
  
  if (days > 0) {
    return days + " ngày: " + hours + " giờ: " + minutes + " phút: " + seconds + " giây";
  } else if (hours > 0) {
    return hours + " giờ: " + minutes + " phút: " + seconds + " giây";
  } else if (minutes > 0) {
    return minutes + " phút: " + seconds + " giây";
  } else {
    return seconds + " giây";
  }
}

/**
 * getConversionMatrix:
 * Lấy ma trận tỷ giá chuyển đổi từ file Database.
 *
 * Quy ước: Ma trận tỷ giá được lưu trên sheet "D4" của file Database, bắt đầu từ hàng 1, cột 9 (cột I) đến hàng, cột cuối cùng.
 * Ví dụ: 
 *   - Ô tại hàng 1, cột 9 cho biết header của cột (ví dụ "USD", "FG", "GOLD", …)
 *   - Hàng đầu tiên sau header chứa các giá trị tương ứng.
 *
 * @return {Array|null} Ma trận tỷ giá ở dạng mảng 2 chiều, hoặc null nếu xảy ra lỗi.
 */
function getConversionMatrix() {
  try {
    var dbId = getFileIdOptimized("Database");
    if (!dbId) {
      Logger.log("Không lấy được ID của Database");
      return null;
    }
    var dbSpreadsheet = SpreadsheetApp.openById(dbId);
    var dbSheet = dbSpreadsheet.getSheetByName("D4");
    if (!dbSheet) {
      Logger.log("Không tìm thấy sheet 'D4' trong Database");
      return null;
    }
    
    var startRow = 1;
    var startCol = 9; // Cột I chứa dữ liệu ma trận
    var lastRow = dbSheet.getLastRow();
    var lastCol = dbSheet.getLastColumn();
    var numRows = lastRow - startRow + 1;
    var numCols = lastCol - startCol + 1;
    var matrix = dbSheet.getRange(startRow, startCol, numRows, numCols).getValues();
    return matrix;
  } catch (error) {
    Logger.log("Error in getConversionMatrix: " + error);
    return null;
  }
}

/**
 * convertPrice: Chuyển đổi giá từ sourceCurrency sang targetCurrency dựa trên ma trận tỷ giá.
 *
 * Quy ước ma trận: 
 * - Hàng đầu tiên (index 0) chứa các đơn vị mục tiêu (với các giá trị bắt đầu từ cột 2).
 * - Cột đầu tiên (index 0) chứa các đơn vị nguồn (với các giá trị bắt đầu từ hàng 2).
 * Công thức chuyển đổi: Giá chuyển đổi = price * (tỷ số tương ứng)
 *
 * @param {Number} price - Giá gốc cần chuyển đổi.
 * @param {String} sourceCurrency - Đơn vị của giá gốc.
 * @param {String} targetCurrency - Đơn vị cần chuyển sang.
 * @param {Array} rateMatrix - Ma trận tỷ giá được lấy từ getConversionMatrix.
 * @return {Number|null} Giá sau khi chuyển đổi, hoặc null nếu không tìm thấy tỷ giá.
 */
function convertPrice(price, sourceCurrency, targetCurrency, rateMatrix) {
  var sourceCurr = sourceCurrency.toString().trim().toUpperCase();
  var targetCurr = targetCurrency.toString().trim().toUpperCase();
  
  if (sourceCurr === targetCurr) {
    return price;
  }
  
  // Duyệt ma trận tỷ giá, giả định đầu tiên chứa các đơn vị ở hàng 0 và cột 0
  for (var i = 1; i < rateMatrix.length; i++) {
    var rowCurr = rateMatrix[i][0].toString().trim().toUpperCase();
    if (rowCurr === sourceCurr) {
      for (var j = 1; j < rateMatrix[0].length; j++) {
        var colCurr = rateMatrix[0][j].toString().trim().toUpperCase();
        if (colCurr === targetCurr) {
          var factor = parseFloat(rateMatrix[i][j]);
          if (isNaN(factor)) {
            Logger.log("convertPrice: Factor không hợp lệ cho chuyển đổi từ " + sourceCurr + " sang " + targetCurr);
            return null;
          }
          return price * factor;
        }
      }
    }
  }
  Logger.log("convertPrice: Không tìm thấy tỷ giá chuyển đổi từ " + sourceCurr + " sang " + targetCurr);
  return null;
}