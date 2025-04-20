//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Khai báo biến/biến cấu hình toàn cục
//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────

var DATABASE_ID = "1MOEie6MQS3P7tzKYqpbX-tOacN2u0qc2B7hlEshMItc";
var KHO_ID = "1ca9InAhsmgdjtRHDiLrSIEjciLHa05DVN2GiVChyUR0";
var BOOSTING_ID = "1YDOU6soJsgjeNix8COLfpCCM6S1y9x_fbgmDd8YGaB8";
var TRADER11_ID = "1bwAecdS_qavOJtJlSjja4AI_HPHqADt5UtdhTLJtF0o";
var TRADER21_ID = "1Ofbz2zlL4sqOWhZnY17DKAcuDQ6DAmCVBRGRgUVe8Fw";
var TRADER12_ID = "1bwAecdS_qavOJtJlSjja4AI_HPHqADt5UtdhTLJtF0o";
var TRADER22_ID = "1Ofbz2zlL4sqOWhZnY17DKAcuDQ6DAmCVBRGRgUVe8Fw";
var FARMER1_ID = "1m5akI-hi_dCht8hIYDSNuFDT7ZjYmEpPm_a0GBUJzO4";
var FARMER2_ID = "1m5akI-hi_dCht8hIYDSNuFDT7ZjYmEpPm_a0GBUJzO4";
var FARMER3_ID = "1m5akI-hi_dCht8hIYDSNuFDT7ZjYmEpPm_a0GBUJzO4";
var FARMER4_ID = "1m5akI-hi_dCht8hIYDSNuFDT7ZjYmEpPm_a0GBUJzO4";

//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Config Trader1
//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Config Trader11*************************************************************************************************************
var trader11Boosting = {
  fileID: TRADER11_ID,           // ID của file Google Sheets chứa dữ liệu
  sheetName: "Boosting",         // Tên của sheet lưu các đơn hàng boosting

  headerRow: 10,                 // Dòng Header
  dataStartRow: 11,              // Dòng bắt đầu chứa dữ liệu thực tế (bỏ qua header)

  // Các cột dữ liệu (đánh số theo thứ tự cột trong sheet)
  dateTimeCol: 1,                // Cột ngày giờ tạo đơn
  idCol: 2,                      // Cột mã đơn hàng
  suorceSellCol: 3,              // Cột nguồn bán (ví dụ: Discord, G2G, Website...)
  customerNameCol: 4,            // Cột tên khách hàng
  playModeCol: 5,                // Cột hình thức chơi (Pilot / Selfplay)
  btagCol: 6,                    // Cột BattleTag hoặc thông tin tài khoản game
  serviceNameCol: 7,             // Cột tên dịch vụ (ví dụ: level 1–90, campaign, v.v.)

  handlingTimeCol: 8,            // Cột thời gian thực hiện đơn hàng

  priceSellCol: 9,               // Cột thông giá bán
  priceSellTypeCol: 10,          // Cột đơn vị giá bán
  lowestPriceCol: 11,            // Cột giá thấp nhất G2G

  noteCol: 12,                   // Cột note/cờ để push

  flashCol: 13,                  // Cột cờ đã pushed
};

// Config Trader12*************************************************************************************************************


//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Config Trader2
//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────

// Config Trader21*************************************************************************************************************

// Config Trader22*************************************************************************************************************


//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Config Farmer
//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────




//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Config Chung
//────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
// Cấu hình cho sheet Boosting Orders
var boostingOrders = {
  fileID: BOOSTING_ID,           // ID của file Google Sheets chứa dữ liệu
  sheetName: "Orders",           // Tên của sheet lưu các đơn hàng boosting

  dataStartRow: 6,               // Dòng bắt đầu chứa dữ liệu thực tế (bỏ qua header)

  // Các cột dữ liệu (đánh số theo thứ tự cột trong sheet)
  dateTimeCol: 1,                // Cột ngày giờ tạo đơn
  idCol: 2,                      // Cột mã đơn hàng
  suorceSellCol: 3,              // Cột nguồn bán (ví dụ: Discord, G2G, Website...)
  customerNameCol: 4,            // Cột tên khách hàng
  playModeCol: 5,                // Cột hình thức chơi (Pilot / Selfplay)
  btagCol: 6,                    // Cột BattleTag hoặc thông tin tài khoản game
  serviceNameCol: 7,            // Cột tên dịch vụ (ví dụ: level 1–90, campaign, v.v.)

  handlingTimeCol: 8,           // Cột thời gian thực hiện đơn hàng
  pcCol: 9,                      // Cột thông tin PC / Trạng thái khi nhận đơn

  deadLineCol: 10,              // Cột thời hạn hoàn thành đơn
  orderStatusCol: 11,           // Cột trạng thái đơn hàng (mới, đang làm, hoàn thành...)
  handlerCol: 12,               // Cột người thực hiện (booster)

  timeStampCol: 13,             // Cột timestamp các thao tác hệ thống (ghi log)
  preStepsCol: 14,              // Cột các bước trước khi thực hiện
  postStepsCol: 15,             // Cột các bước sau khi thực hiện
  stepProofCol: 16,             // Cột bằng chứng thực hiện (ảnh/video)

  dateTimeCompletionsCol: 17    // Cột ngày giờ hoàn thành đơn hàng
};