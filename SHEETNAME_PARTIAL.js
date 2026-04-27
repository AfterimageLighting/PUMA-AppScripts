function SHEETNAME_PARTIAL() {
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  var parts = sheetName.split(" - ");
  return parts[0];
}