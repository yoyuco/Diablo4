//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Xử lý Chung
//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────

/**
 * Hàm autoFillTime:
 * - Nhận đầu vào: fileId, sheetName, row (dòng đang chỉnh sửa), startCol, endCol, fillCol.
 * - Kiểm tra:
 *    + sheetName (lấy từ sheet hiện tại) phải là main sheet (chuỗi 2 ký tự số từ "01" đến "31").
 *    + Dữ liệu trong khoảng từ startCol đến endCol ở dòng (điều chỉnh) đã đầy đủ.
 * - Nếu thỏa điều kiện, điền ngày giờ hiện tại vào ô tại (row, fillCol).
 */
function autoFillTime(fileId, sheetName, row, startCol, endCol, fillCol) {
  // Kiểm tra sheet hiện tại có phải là main sheet không (theo định nghĩa: tên 2 ký tự số "01" đến "31")
  if (!isMainSheet(sheetName)) {
    return;
  }
  
  // Kiểm tra dữ liệu ở dòng 'row' trong khoảng từ startCol đến endCol đã đầy đủ hay chưa
  if (!checkDataComplete(row, row, startCol, endCol)) {
    return;
  }
  
  // Mở file và lấy sheet theo sheetName
  var ss = SpreadsheetApp.openById(fileId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("Không tìm thấy sheet: " + sheetName);
    return;
  }
  
  // Điền ngày giờ hiện tại vào ô tại dòng row, cột fillCol
  sheet.getRange(row, fillCol).setValue(new Date());
}

//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Xử lý Trader1
//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────




//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Xử lý Trader2
//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────




//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Xử lý Farmer
//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────

