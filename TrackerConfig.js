// TODO: put your real "Open Projects" folder ID here
var OPEN_PROJECTS_FOLDER_ID = '1acRZOrQUIzhoIav1Rw8d2GPosaDvNWx5';

/**
 * Add any missing project folders into the "Tracker Config" sheet.
 * - Keeps existing rows and Quote IDs intact.
 * - New rows get:
 *   A: unchecked checkbox
 *   B: HYPERLINK to the project folder, label = short project name
 *   C: "<short project name> - Project Tracker"
 *   D/E: blank (you fill in Quote Sheet ID + Tab later)
 */
function updateTrackerConfigFromFolders() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Tracker Config');
  if (!sheet) {
    sheet = ss.insertSheet('Tracker Config');
    sheet.getRange(1, 1, 1, 5).setValues([[
      'Enable?', 'Project', 'Tracker Sheet Name', 'Quote Sheet ID', 'Quote Tab'
    ]]);
  }

  // Read existing Tracker Sheet Names (column C) so we don't duplicate rows
  var lastRow = sheet.getLastRow();
  var existingNames = new Set();
  if (lastRow > 1) {
    var existing = sheet.getRange(2, 3, lastRow - 1, 1).getValues(); // col C
    existing.forEach(function (row) {
      var name = row[0];
      if (name) existingNames.add(String(name));
    });
  }

  var parent = DriveApp.getFolderById(OPEN_PROJECTS_FOLDER_ID);
  var folders = parent.getFolders();

  var newRows = [];
  while (folders.hasNext()) {
    var folder = folders.next();
    var folderName = folder.getName();        // e.g. "Boyd Natatorium- Hickory Construction"

    // Take everything before the first "-" and trim spaces: "Boyd Natatorium"
    var shortName = folderName.split('-')[0].trim();

    var trackerSheetName = shortName + ' - Project Tracker';

    // Skip if we already have this tracker in column C
    if (existingNames.has(trackerSheetName)) {
      continue;
    }

    var url = folder.getUrl();

    // Escape any double quotes in the label for the HYPERLINK formula
    var safeLabel = shortName.replace(/"/g, '""');

    var projectCellFormula = '=HYPERLINK("' + url + '","' + safeLabel + '")';

    newRows.push([
      false,              // Enable? (unchecked)
      projectCellFormula, // Project (hyperlink)
      trackerSheetName,   // Tracker Sheet Name
      '',                 // Quote Sheet ID
      ''                  // Quote Tab
    ]);
  }

  if (newRows.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, 5).setValues(newRows);

    // Add checkboxes to the new "Enable?" cells
    sheet.getRange(startRow, 1, newRows.length, 1).insertCheckboxes();
  }

  Logger.log('Added ' + newRows.length + ' new rows to Tracker Config.');
}
