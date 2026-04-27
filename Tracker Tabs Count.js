function countTrackerTabs() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();

  var count = 0;

  sheets.forEach(function(sh) {
    var name = sh.getName().toLowerCase().trim();

    if (name.includes("project tracker")) {
      count++;
      Logger.log(name);
    }
  });

  Logger.log("TOTAL PROJECT TRACKERS: " + count);
}