//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Các hàm xử lý liên quan đến Data (Get/fill/Normalize,....)
//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
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

function isDataComplete(sheet, row, startCol, endCol) {
  var values = sheet.getRange(row, startCol, 1, endCol - startCol + 1).getValues()[0];
  return values.every(function(cell) {
    return cell.toString().trim() !== "";
  });
}

/**
 * Quét và lấy mảng các dòng pending từ sheet nguồn theo config.
 * @param {Object} src  Config nguồn, gồm các trường:
 *   - fileID, sheetName, dataStartRow, flashCol,
 *   - idCol, noteCol, dateTimeCol, handlingTimeCol
 * @return {Array<{row:number, data:any[]}>}
 */
function fetchPendingData(src) {
  var ss  = SpreadsheetApp.openById(src.fileID);
  var sh  = ss.getSheetByName(src.sheetName);
  var last = sh.getLastRow();
  var now  = new Date();
  var out  = [];

  for (var r = src.dataStartRow; r <= last; r++) {
    if (sh.getRange(r, src.flashCol).getValue() === 'Pushed') continue;
    if (isDataComplete(sh, r, src.idCol, src.noteCol)) {
      // đánh dấu timestamp
      sh.getRange(r, src.dateTimeCol).setValue(now);

      // lấy mảng data
      var num = src.handlingTimeCol - src.dateTimeCol + 1;
      var row = sh.getRange(r, src.dateTimeCol, 1, num).getValues()[0];
      out.push({row: r, data: row});
    }
  }
  SpreadsheetApp.flush();
  return out;
}

/**
 * Ghi mảng pending xuống sheet Orders và đánh dấu & khóa dòng nguồn.
 */
function writeData(pending, src, dst) {
  var ssDst = SpreadsheetApp.openById(dst.fileID);
  var shDst = ssDst.getSheetByName(dst.sheetName);
  var ssSrc = SpreadsheetApp.openById(src.fileID);
  var shSrc = ssSrc.getSheetByName(src.sheetName);

  // 1) Tìm dòng trống đầu tiên
  var nextDst = CommonLib.findFirstEmptyRow(
    shDst,
    dst.dataStartRow,
    dst.dateTimeCol
  );

  // 2) Lặp qua từng item trong pending
  pending.forEach(function(item) {
    // nếu cần, insert thêm dòng
    if (nextDst > shDst.getLastRow()) {
      shDst.insertRowAfter(shDst.getLastRow());
    }
    // ghi dữ liệu
    shDst.getRange(nextDst, dst.dateTimeCol,
                  1, item.data.length)
         .setValues([item.data]);
    SpreadsheetApp.flush();

    // 3) đánh dấu & khóa dòng nguồn
    CommonLib.markRowAsPushedAndProtect(
      shSrc,
      item.row,
      src.flashCol
    );

    nextDst++;
  });
}

/**
 * Hàm chung để push data từ bất kỳ src→dst nào.
 * Chỉ cần gọi CommonLib.pushData(srcConfig, dstConfig);
 */
function pushData(srcConfig, dstConfig) {
  // lock để tránh chạy chồng
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var pending = fetchPendingData(srcConfig);
    if (pending.length) {
      writeData(pending, srcConfig, dstConfig);
      Logger.log('Push ' + pending.length + ' dòng thành công.');
    } else {
      Logger.log('Không có dữ liệu mới để push.');
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Tìm dòng trống đầu tiên trong sheet, dựa vào cột kiểm tra.
 * @param {Sheet} sheet       – sheet đích
 * @param {number} startRow   – dòng bắt đầu tìm
 * @param {number} checkCol   – cột dùng làm tiêu chí (ô trống)
 * @return {number}           – số dòng đầu tiên có ô checkCol trống
 */
function findFirstEmptyRow(sheet, startRow, checkCol) {
  var last = sheet.getLastRow();
  for (var r = startRow; r <= last; r++) {
    if (sheet.getRange(r, checkCol).getValue() === "") {
      return r;
    }
  }
  // nếu hết rồi, trả về dòng kế tiếp để insert
  return last + 1;
}


/**
 * Đánh dấu đã Push và khóa dòng nguồn không cho chỉnh sửa nữa.
 * @param {Sheet} sheet     – sheet nguồn
 * @param {number} row      – dòng cần đánh dấu
 * @param {number} flashCol – cột flag “Pushed”
 */
function markRowAsPushedAndProtect(sheet, row, flashCol) {
  // 1) Đánh dấu
  sheet.getRange(row, flashCol).setValue('Pushed');
  // 2) Khóa toàn bộ row
  var range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
  var protection = range.protect().setDescription('Lock after push');
  // remove all editors except owner
  try {
    protection.removeEditors(protection.getEditors());
  } catch(e) {
    // nếu không có editor nào để remove thì bỏ qua
  }
}
//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Các hàm xử lý liên quan đến Sheet/File
//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Liên quan đến column/Row****************************************************************************************************
/**
 * Chuyển đổi số thứ tự cột thành ký tự cột (ví dụ 1 thành "A", 28 thành "AB").
 *
 * @param {number} column - Số thứ tự cột.
 * @return {string} - Ký tự cột tương ứng.
 */
function columnToLetter(column) {
  var letter = '';
  while (column > 0) {
    var temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
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

// Liên quan đến File/Sheet ***************************************************************************************************
/**
 * Kiểm tra xem tên sheet có phải là main sheet hay không.
 * Điều kiện: Tên sheet phải là chuỗi gồm 2 ký tự số từ "01" đến "31".
 *
 * @param {string} sheetName - Tên của sheet cần kiểm tra.
 * @return {boolean} - Trả về true nếu tên sheet nằm trong khoảng từ "01" đến "31", ngược lại trả về false.
 */
function isMainSheet(sheetName) {
  // Sử dụng biểu thức chính quy để kiểm tra:
  // ^ : bắt đầu chuỗi
  // (0[1-9] : số 01 đến 09
  // |[12][0-9] : số 10 đến 29
  // |3[01]) : số 30 hoặc 31
  // $ : kết thúc chuỗi
  var regex = /^(0[1-9]|[12][0-9]|3[01])$/;
  return regex.test(sheetName);
}

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

function getFileIdOptimized(fileName) {
  var mapping = getFileMapping();
  return mapping[fileName.toString().trim().toLowerCase()] || null;
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
//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Các hàm tiện ích khác
//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Xử lý chuỗi thời gian ******************************************************************************************************

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