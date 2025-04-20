// Cấu hình phân quyền dựa trên role (các role có thể được mở rộng)
var ROLE_CONFIG = {
  "trader11": { fileId: TRADER11_ID, visibleStart: 1,  visibleEnd: 45 },
  "trader21": { fileId: TRADER21_ID, visibleStart: 1,  visibleEnd: 70 }
  // Bạn có thể mở rộng thêm: trader12, trader22, trader13, trader23, v.v.
};

// ------------------------------
// Cấu hình cho Trader1
// ------------------------------
var dvRegionsTrader1 = [
  {
    conditionCol: 21,
    targetStart: 22,
    targetCount: 3,
    finalClearCol: 24,
    dependentConfig: { targetStartCol: 22, suffixes: ["_List1", "_List2"] },
    commonConfig: { targetStartCol: 22, listNames: ["Common_List1", "Common_List2", "Common_List3"] }
  },
  {
    conditionCol: 39,
    targetStart: 40,
    targetCount: 3,
    finalClearCol: 42,
    dependentConfig: { targetStartCol: 40, suffixes: ["_List1", "_List2"] },
    commonConfig: { targetStartCol: 40, listNames: ["Common_List1", "Common_List2", "Common_List3"] }
  }
];

var afRegionsTrader1 = [
  { conditionStart: 3, conditionEnd: 13 },
  { conditionStart: 18, conditionEnd: 31 },
  { conditionStart: 36, conditionEnd: 49 },
  { conditionStart: 54, conditionEnd: 63 }
];

var saleGoldConfigs = [
  {
    saleQtyCol: 6,
    salePriceCol: 7,
    saleCurrencyCol: 8,
    purchasePriceCol: 9,
    profitCol1: 10,
    profitCol2: 11
  }
];

var khoConfigForSale = {
  sheetName: "Items",
  rowData: 2,
  currencyAvgCols: { usd: 5, vnd: 6, fg: 7, gold: 8 }
};

// Cấu hình cho bảng Gold: Kiểm tra từ cột 3 đến 13, nếu có dữ liệu thì tạo dropdown cho ô ở cột 11.
var goldAccountValidationConfigTrader1 = {
  checkRange: { startCol: 3, endCol: 11 },
  accountCol: 12
  };

// Cấu hình cho bảng Items: Ví dụ kiểm tra từ cột 17 đến 34, tạo dropdown cho ô ở cột 31.
var itemsAccountValidationConfigTrader1 = [
  {
    checkRange: { startCol: 18, endCol: 29 },
    accountCol: 30
  },
  {
    checkRange: { startCol: 36, endCol: 47 },
    accountCol: 48
  }
];

var goldPush = {
  checkRange: { startCol: 1, endCol: 13 },      // Vùng kiểm tra dữ liệu (cột 3 đến 10)
  accountCol: 12,                                // Cột chứa ô Account (định dạng "AccXX-TraderYY")
  sourceDataRange: { startCol: 1, endCol: 13 },    // Phạm vi dữ liệu cần đẩy từ Trader1
  destStartRow: 10,                              // Hàng bắt đầu của bảng trong file Trader2
  destDataStartCol: 37,                           // Cột bắt đầu ghi dữ liệu trong file Trader2
  saleQtyCol: 6,  // Cột chứa số lượng bán gold trên Trader1
  kho: {
    sheetName: khoConfigForSale.sheetName,
    rowData: khoConfigForSale.rowData,
    qtyCol: 3  // Nếu bạn có cột tổng số lượng trong khoConfigForSale (nếu chưa có, thêm vào)
  }
};

var itemsPush = [
  {
    checkRange: { startCol: 16, endCol: 31 },
    accountCol: 30,
    sourceDataRange: { startCol: 21, endCol: 24 },
    destStartRow: 10,
    destDataStartCol: 54,
    pushedCol: 32,
    // Đối với items, saleQty mặc định = 1, nên không cần chỉ định
    kho: {
      sheetName: "Items", // ví dụ: "Items"
      rowData: 3,       // ví dụ: 3
      itemsCol: 2,       // cột chứa chuỗi items trong Kho
      qtyCol: 3            // cột tổng số lượng của items
    }
  },
  {
    checkRange: { startCol: 34, endCol: 49 },
    accountCol: 48,
    sourceDataRange: { startCol: 39, endCol: 42 },
    destStartRow: 10,
    destDataStartCol: 74,
    pushedCol: 50,
    kho: {
      sheetName: "Items", // ví dụ: "Items"
      rowData: 3,       // ví dụ: 3
      itemsCol: 2,       // cột chứa chuỗi items trong Kho
      qtyCol: 3 
    }
  }
];

var saleItemsSaleConfig = [
  {
  // Xác định phạm vi 4 cột chứa dữ liệu item trong đơn hàng của Trader1
  saleItemsConfig: {
    sourceDataRange: { startCol: 21, endCol: 24 }
    },
  // Các cột trên file Trader1 chứa giá bán, đơn vị tiền và cột ghi kết quả
    saleConfig: {
      salePriceCol: 25,       // Cột chứa giá bán của items trong Trader1
      saleCurrencyCol: 26,    // Cột chứa đơn vị tiền tệ bán (ví dụ: "usd" hoặc "vnd")
      purchasePriceCol: 27,   // Cột ghi giá mua (sau khi tính toán) trong file Trader1
      profitCol1: 28,         // Cột ghi lợi nhuận USD (nếu đơn vị bán không phải VND)
      profitCol2: 29          // Cột ghi lợi nhuận VND (nếu đơn vị bán là VND)
    },
  // Cấu hình cho file Kho để lấy giá trung bình mua và ghi kết quả
    khoConfig: {
      sheetName: "Items",     // Tên sheet chứa record items trong file Kho
      rowData: 3,             // Dữ liệu bắt đầu từ hàng 3 (hàng 1 tiêu đề, hàng 2 dành cho gold)
      itemsCol: 2,            // Cột chứa chuỗi Items (sau khi ghép bằng combineItemString)
      qtyCol: 3,              // Cột chứa số lượng của items trong Kho (dùng cho việc tính toán, mặc định qty = 1)
      currencyAvgCols: {      // Cột trung bình giá mua theo đơn vị tiền
        usd: 5,               // Cột trung bình giá mua USD trong Kho
        vnd: 6                // Cột trung bình giá mua VND trong Kho
      },
      purchasePriceCol: 27,   // Cột ghi giá mua (sau khi tính toán) trong file Trader1
      profitCol1: 28,         // Cột ghi lợi nhuận USD (nếu đơn vị bán không phải VND)
      profitCol2: 29
    }
  },
  {
  // Xác định phạm vi 4 cột chứa dữ liệu item trong đơn hàng của Trader1
  saleItemsConfig: {
    sourceDataRange: { startCol: 39, endCol: 42 }
    },
  // Các cột trên file Trader1 chứa giá bán, đơn vị tiền và cột ghi kết quả
    saleConfig: {
      salePriceCol: 43,       // Cột chứa giá bán của items trong Trader1
      saleCurrencyCol: 44,    // Cột chứa đơn vị tiền tệ bán (ví dụ: "usd" hoặc "vnd")
      purchasePriceCol: 45,   // Cột ghi giá mua (sau khi tính toán) trong file Trader1
      profitCol1: 46,         // Cột ghi lợi nhuận USD (nếu đơn vị bán không phải VND)
      profitCol2: 47          // Cột ghi lợi nhuận VND (nếu đơn vị bán là VND)
    },
  // Cấu hình cho file Kho để lấy giá trung bình mua và ghi kết quả
    khoConfig: {
      sheetName: "Items",     // Tên sheet chứa record items trong file Kho
      rowData: 3,             // Dữ liệu bắt đầu từ hàng 3 (hàng 1 tiêu đề, hàng 2 dành cho gold)
      itemsCol: 2,            // Cột chứa chuỗi Items (sau khi ghép bằng combineItemString)
      qtyCol: 3,              // Cột chứa số lượng của items trong Kho (dùng cho việc tính toán, mặc định qty = 1)
      currencyAvgCols: {      // Cột trung bình giá mua theo đơn vị tiền
        usd: 5,               // Cột trung bình giá mua USD trong Kho
        vnd: 6                // Cột trung bình giá mua VND trong Kho
      },
      purchasePriceCol: 45,   // Cột ghi giá mua (sau khi tính toán) trong file Trader1
      profitCol1: 46,         // Cột ghi lợi nhuận USD (nếu đơn vị bán không phải VND)
      profitCol2: 47
    }
  },
];

// ------------------------------
// Cấu hình cho Trader2
// ------------------------------
var conversionRegions = [
  { priceCol: 6, targetStartCol: 8, targetEndCol: 10 },
  { priceCol: 25, targetStartCol: 27, targetEndCol: 30 }
];

var dvRegionsTrader2 = [
  {
    conditionCol: 21,
    targetStart: 22,
    targetCount: 3,
    finalClearCol: 24,
    dependentConfig: { targetStartCol: 22, suffixes: ["_List1", "_List2"] },
    commonConfig: { targetStartCol: 22, listNames: ["Common_List1", "Common_List2", "Common_List3"] }
  }
];

var afRegionsTrader2 = [
  { conditionStart: 3, conditionEnd: 10 },
  { conditionStart: 19, conditionEnd: 30 },
  { conditionStart: 37, conditionEnd: 51, timestampCol: 52, skipSequence: true, diffTime: true },
  { conditionStart: 54, conditionEnd: 71, timestampCol: 72, skipSequence: true, diffTime: true },
  { conditionStart: 74, conditionEnd: 92, timestampCol: 93, skipSequence: true, diffTime: true }
];

// Ví dụ cấu hình cho Trader2 (updategold vào kho)
// Cấu hình cho 2 bảng Gold (trong Trader2)
var traderUnifiedConfig = [
  {
    trader: {
      fileId: getFileIdOptimized("Trader21"),
      sheet: undefined,  // Nếu chưa có, hàm sẽ tự mở lại từ fileId
      dataRegion: { startCol: 1, endCol: 14 },
      qtyCol: 5,
      // Chỉ có cột giá mua USD được cung cấp (không có VND, FG)
      currencyCols: { usd: 8, vnd: 9, fg: 10 },
      accountCol: 11,
      rowHeader: 10
    },
    kho: {
      sheetName: "Items",
      rowData: 2,
      rowHeader: 1,
      qtyCol: 3,
      // Các cột trung bình giá theo đơn vị trong Kho cho bảng Gold thứ nhất
      currencyAvgCols: { USD: 5, VND: 6, FG: 7 }
    }
  },
  {
    trader: {
      fileId: getFileIdOptimized("Trader21"),
      sheet: undefined,
      // Ví dụ bảng Gold thứ hai nằm ở cột khác (chỉnh theo thực tế)
      dataRegion: { startCol: 74, endCol: 93 },
      qtyCol: 83,
      currencyCols: { usd: 85 },
      accountCol: 90,
      rowHeader: 10
    },
    kho: {
      sheetName: "Items",
      rowData: 2,
      rowHeader: 1,
      qtyCol: 3,
      // Các cột trung bình giá theo đơn vị trong Kho cho bảng Gold thứ nhất
      currencyAvgCols: { USD: 5, VND: 6, FG: 7 }
    }
  }
];

var itemsUpdateConfig = {
  dataRegion: { startCol: 17, endCol: 34 },
  itemRegion: { startCol: 21, endCol: 24 },
  purchasePriceCol: 0,
  currencyCols: { usd: 27, vnd: 28, fg: 29, gold: 30 },
  accountCol: 31,
  pushedCol: 35
};

var khoConfigForItems = {
  sheetName: "Items",           // Tên sheet chứa record items trong file Kho
  rowData: 3,                   // Dữ liệu bắt đầu từ hàng 3 (hàng 1 tiêu đề, hàng 2 dành cho gold)
  itemsCol: 2,                  // Cột chứa chuỗi Items (sau khi ghép)
  qtyCol: 3,                    // Cột chứa tổng số lượng của Items
  currencyAvgCols: {            // Cột trung bình giá mua cho các đơn vị tiền (nếu cần)
    usd: 5,
    vnd: 6,
    fg: 7,
    gold: 8
  }
};

// Cấu hình cho bảng Gold: Kiểm tra từ cột 3 đến 13, nếu có dữ liệu thì tạo dropdown cho ô ở cột 11.
var goldAccountValidationConfigTrader2 = [
  {
    checkRange: { startCol: 3, endCol: 10 },
    accountCol: 11
  },
  {
    checkRange: { startCol: 74, endCol: 89 },
    accountCol: 90
  }
];

// Cấu hình cho bảng Items: Ví dụ kiểm tra từ cột 17 đến 34, tạo dropdown cho ô ở cột 31.
var itemsAccountValidationConfigTrader2 = {
  checkRange: { startCol: 19, endCol: 30 },
  accountCol: 31
};