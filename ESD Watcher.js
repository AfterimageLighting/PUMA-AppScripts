function onEdit(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  // --- CONFIG ---
  const ESD_COL = 10; // Column J
  const TRACKER_SUFFIX = ' - Project Tracker';
  // --------------

  // Only run on sheets whose name ends with " - Project Tracker"
  const sheetName = sheet.getName();
  if (!sheetName.endsWith(TRACKER_SUFFIX)) return;

  // Only run on edits in column J
  if (range.getColumn() !== ESD_COL) return;

  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  // Get the spreadsheet's time zone
  const timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const nowString = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm');

  // Handle single-cell and multi-cell edits
  if (numRows === 1 && numCols === 1) {
    const value = e.value;

    // If cleared, remove note
    if (!value) {
      range.setNote('');
      return;
    }

    // Set/update the timestamp note
    const noteText = 'ESD entered: ' + nowString;
    range.setNote(noteText);
  } else {
    // Multi-cell edit in column J (e.g., paste)
    const values = range.getValues();
    for (let r = 0; r < numRows; r++) {
      const cell = range.getCell(r + 1, 1); // still column J
      const value = values[r][0];

      if (!value) {
        cell.setNote('');
      } else {
        const noteText = 'ESD entered: ' + nowString;
        cell.setNote(noteText);
      }
    }
  }
}
