function removeAdminLockProtections() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  sheets.forEach(function(sheet) {
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(function(protection) {
      if (protection.getDescription() === "Gold data updated and locked by admin") {
        protection.remove();
        Logger.log("Removed protection on sheet: " + sheet.getName());
      }
    });
  });
  
  Logger.log("Finished removing all admin lock protections.");
}

