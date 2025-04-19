function doGet(e) {
  var template = HtmlService.createTemplateFromFile("Login");
  // Không cần truyền trader1FileId, trader2FileId nữa vì bây giờ lấy tự động trong Library
  template.pageTitle = "Đăng nhập";
  return template.evaluate().setTitle(template.pageTitle);
}

function webProcessLogin(userID, userPassword) {
  return CommonLib.processLoginOptimized(userID, userPassword);
}