/**
 * Lấy dữ liệu từ một sheet của Google Sheets dựa vào phạm vi cho trước.
 *
 * @param {string} sheetId - ID của file Google Sheets.
 * @param {string} sheetName - Tên của sheet cần lấy dữ liệu. Nếu rỗng, mặc định lấy sheet đầu tiên.
 * @param {string} startCell - Ô bắt đầu của phạm vi (ví dụ: "A1").
 * @param {string} endCell - Ô kết thúc của phạm vi (ví dụ: "D10"). Nếu chỉ truyền "D" hoặc "10" hoặc không truyền, tự xác định.
 * @return {Array} - Mảng chứa dữ liệu trong phạm vi đã chỉ định.
 */
function getSheetData(sheetId, sheetName, startCell, endCell) {
  // Mở file Google Sheets theo ID
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var sheet;
  // Lấy sheet theo tên nếu có, nếu không lấy sheet đầu tiên
  if (sheetName && sheetName.trim() !== "") {
    sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Không tìm thấy sheet có tên: " + sheetName);
    }
  } else {
    sheet = spreadsheet.getSheets()[0];
  }
  
  // Xác định tọa độ ô bắt đầu: phân tách chữ (cột) và số (dòng)
  var startMatch = startCell.match(/([A-Za-z]+)([0-9]+)/);
  if (!startMatch) {
    throw new Error("startCell không hợp lệ: " + startCell);
  }
  var startColLetter = startMatch[1];
  var startRow = parseInt(startMatch[2], 10);
  var startCol = letterToColumn(startColLetter);
  
  var finalEndRow, finalEndCol;
  
  // Nếu không truyền endCell hoặc truyền chuỗi rỗng, tự xác định hết dữ liệu
  if (!endCell || endCell.trim() === "") {
    finalEndRow = sheet.getLastRow();
    finalEndCol = sheet.getLastColumn();
  } else {
    // Kiểm tra nếu endCell chỉ là chữ (cột) hoặc chỉ là số (dòng)
    if (endCell.match(/^[A-Za-z]+$/)) {
      // endCell chỉ chứa ký tự cột
      finalEndCol = letterToColumn(endCell);
      finalEndRow = sheet.getLastRow();
    } else if (endCell.match(/^[0-9]+$/)) {
      // endCell chỉ chứa số dòng
      finalEndRow = parseInt(endCell, 10);
      finalEndCol = sheet.getLastColumn();
    } else {
      // Giả sử endCell là địa chỉ ô đầy đủ, ví dụ "D10"
      var endMatch = endCell.match(/([A-Za-z]+)([0-9]+)/);
      if (!endMatch) {
        throw new Error("endCell không hợp lệ: " + endCell);
      }
      var endColLetter = endMatch[1];
      finalEndCol = letterToColumn(endColLetter);
      finalEndRow = parseInt(endMatch[2], 10);
    }
  }
  
  // Xây dựng lại địa chỉ ô kết thúc đầy đủ
  var finalEndCell = columnToLetter(finalEndCol) + finalEndRow;
  
  // Lấy phạm vi từ ô bắt đầu đến ô kết thúc
  var range = sheet.getRange(startCell + ":" + finalEndCell);
  var data = range.getValues();
  return data;
}

/**
 * Điền các phần tử của 1 mảng vào file Google Sheets.
 *
 * @param {string} sheetId - ID của file Google Sheets.
 * @param {string} sheetName - Tên của sheet cần điền dữ liệu. Nếu rỗng, mặc định sử dụng sheet đầu tiên.
 * @param {string} startCell - Ô bắt đầu điền dữ liệu (ví dụ: "A1").
 * @param {Array} arr - Mảng dữ liệu cần điền.
 * @param {number} fillType - Kiểu điền: 
 *                            0 - Điền theo cấu trúc ban đầu của mảng (nếu mảng 1 chiều sẽ điền theo hàng ngang),
 *                            2 - Điền theo hàng ngang,
 *                            3 - Điền theo cột dọc.
 */
function fillSheetData(sheetId, sheetName, startCell, arr, fillType) {
  // Mở file Google Sheets theo ID
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var sheet;
  if (sheetName && sheetName.trim() !== "") {
    sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Không tìm thấy sheet có tên: " + sheetName);
    }
  } else {
    sheet = spreadsheet.getSheets()[0];
  }
  
  // Tách thông tin ô bắt đầu (ví dụ "B2")
  var startMatch = startCell.match(/([A-Za-z]+)([0-9]+)/);
  if (!startMatch) {
    throw new Error("startCell không hợp lệ: " + startCell);
  }
  var startColLetter = startMatch[1],
      startRow = parseInt(startMatch[2], 10);
  var startCol = letterToColumn(startColLetter);
  
  // Xử lý mảng dữ liệu theo kiểu điền
  var data;
  if (fillType === 0) {
    // Kiểu 0: nếu mảng là 2 chiều thì dùng nguyên, nếu 1 chiều thì chuyển thành 1 hàng ngang.
    if (Array.isArray(arr) && arr.length > 0 && Array.isArray(arr[0])) {
      data = arr;
    } else {
      data = [arr];
    }
  } else if (fillType === 2) {
    // Kiểu 2: điền theo hàng ngang, đảm bảo dữ liệu là 1 hàng.
    if (Array.isArray(arr) && arr.length > 0 && !Array.isArray(arr[0])) {
      data = [arr];
    } else {
      data = [arr.flat()];
    }
  } else if (fillType === 3) {
    // Kiểu 3: điền theo cột dọc, chuyển mảng thành dạng cột.
    var flatArr = (Array.isArray(arr) && arr.length > 0 && !Array.isArray(arr[0])) ? arr : arr.flat();
    data = flatArr.map(function(item) {
      return [item];
    });
  } else {
    throw new Error("Kiểu điền không hợp lệ: " + fillType);
  }
  
  // Chuẩn hóa dữ liệu: làm cho tất cả các hàng có cùng số cột bằng cách bổ sung giá trị rỗng nếu cần.
  data = normalizeData(data);

  // Xác định số hàng và số cột sau khi chuẩn hóa
  var numRows = data.length;
  var numCols = data[0].length;
  
  // Tính ô kết thúc dựa trên ô bắt đầu và kích thước dữ liệu
  var endCol = startCol + numCols - 1;
  var endRow = startRow + numRows - 1;
  var endCell = columnToLetter(endCol) + endRow;
  
  // Lấy vùng cần điền dữ liệu và tiến hành ghi
  var range = sheet.getRange(startCell + ":" + endCell);
  range.setValues(data);
}

/**
 * Hàm này chuẩn hóa mảng 2 chiều sao cho mỗi hàng có cùng số cột. 
 * Nếu một hàng có ít phần tử hơn, sẽ bổ sung giá trị rỗng "" cho đến khi đủ.
 *
 * @param {Array} data - Mảng dữ liệu 2 chiều cần chuẩn hóa.
 * @return {Array} - Mảng chuẩn hóa.
 */
function normalizeData(data) {
  // Tìm số cột tối đa trong các hàng
  var maxCols = Math.max.apply(null, data.map(function(row) {
    return row.length;
  }));
  
  // Với mỗi hàng, nếu số phần tử < maxCols thì bổ sung giá trị rỗng
  for (var i = 0; i < data.length; i++) {
    while (data[i].length < maxCols) {
      data[i].push("");
    }
  }
  return data;
}

/**
 * Chuyển đổi ký tự cột (ví dụ "A", "AB") thành số thứ tự cột.
 *
 * @param {string} letter - Ký tự cột cần chuyển.
 * @return {number} - Số thứ tự cột.
 */
function letterToColumn(letter) {
  var column = 0;
  var length = letter.length;
  for (var i = 0; i < length; i++) {
    column *= 26;
    column += letter.toUpperCase().charCodeAt(i) - 64;
  }
  return column;
}
