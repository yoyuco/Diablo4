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