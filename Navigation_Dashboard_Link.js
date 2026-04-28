/**
 * Adds a "⬅ Back to Dashboard" link to every Project Tracker sheet.
 * Safe to re-run.
 */
function addBackToDashboardLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('Dashboard');

  if (!dashboard) {
    throw new Error('Dashboard sheet not found.');
  }

  const dashboardGid = dashboard.getSheetId();
  const dashboardUrl = ss.getUrl() + '#gid=' + dashboardGid;

  const sheets = ss.getSheets();
  let updated = 0;

  sheets.forEach(sheet => {
    const name = sheet.getName();

    // Only touch tracker sheets
    if (!isProjectTrackerSheet_(name)) return;

    const cell = sheet.getRange('A1');

    const richText = SpreadsheetApp.newRichTextValue()
      .setText('⬅ Back to Dashboard')
      .setLinkUrl(dashboardUrl)
      .build();

    cell.setRichTextValue(richText);
    cell.setFontWeight('bold');
    cell.setFontSize(12);
    cell.setBackground('#d9ead3');
    cell.setHorizontalAlignment('center');

    sheet.setColumnWidth(1, Math.max(sheet.getColumnWidth(1), 170));

    updated++;
  });

  SpreadsheetApp.getUi().alert(`Back to Dashboard links added to ${updated} tracker sheets.`);
}

/**
 * Decides whether a sheet is a project tracker tab.
 */
function isProjectTrackerSheet_(sheetName) {
  const name = String(sheetName || '').trim();

  if (!name) return false;

  // Main tracker naming patterns we have used
  if (/project tracker$/i.test(name)) return true;
  if (/project track$/i.test(name)) return true;

  return false;
}