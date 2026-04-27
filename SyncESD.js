/************************************
 * ESD central tracking
 * - syncESD(): rebuilds the ESD sheet from all "* - Project Tracker" sheets
 * - onEdit(e): when ESD is edited, push changes back to trackers
 ************************************/

function getSheetGidByName(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  return sheet ? sheet.getSheetId() : null;
}

function syncESD() {
  const ss = SpreadsheetApp.getActive();
  const esdSheetName = 'ESD';
  let esdSheet = ss.getSheetByName(esdSheetName);
  if (!esdSheet) {
    esdSheet = ss.insertSheet(esdSheetName);
  }

  // --- Read existing ESD to preserve Date Entered and detect manual overrides ---
  const lastEsdRow = esdSheet.getLastRow();
  const existingDataRange = lastEsdRow > 1
    ? esdSheet.getRange(2, 1, lastEsdRow - 1, 6) // A–F
    : null;
  const existingData = existingDataRange ? existingDataRange.getValues() : [];

  // Maps keyed by Project|PO|Type|Item
  const esdMap = {};   // ESD date currently on ESD sheet
  const dateMap = {};  // "Date Entered" from col F (manual override marker)

  existingData.forEach(row => {
    const projectName = row[0]; // col A (display value, but we only used it previously)
    const po          = row[1]; // col B
    const type        = row[2]; // col C
    const item        = row[3]; // col D
    const esdVal      = row[4]; // col E
    const dateEntered = row[5]; // col F

    const key = [projectName, po, type, item].join('||');
    if (!projectName || !po || !type || !item) return;

    if (esdVal !== '') {
      esdMap[key] = esdVal;
    }
    if (dateEntered) {
      dateMap[key] = dateEntered; // presence of a date means "user touched this on ESD"
    }
  });

  // --- Walk all Project Tracker sheets ---
  const trackerSheets = ss.getSheets().filter(s =>
    / - Project Tracker$/.test(s.getName())
  );

  const outputRows = [];

  trackerSheets.forEach(sheet => {
    const sheetProjectName = sheet.getName().replace(/ - Project Tracker$/, '');
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;

    const numRows = lastRow - 2;
    const dataRange = sheet.getRange(3, 1, numRows, 10); // A–J
    const data = dataRange.getValues();
    let changed = false;

    for (let i = 0; i < data.length; i++) {
      const row = data[i];

      const status = row[7];  // col H = Status
      const po     = row[8];  // col I = PO
      let   esd    = row[9];  // col J = ESD

      if (!po) continue; // nothing to track

      // Skip rows that are done or scheduled
      const normalizedStatus = String(status).trim().toLowerCase();
      const skipStatuses = ['delivered', 'received', 'scheduled'];
      if (skipStatuses.includes(normalizedStatus)) continue;

      const type = row[2]; // col C
      const item = row[3]; // col D

      const key = [sheetProjectName, po, type, item].join('||');

      // If this row has a manual override on the ESD sheet (i.e., timestamp in Date Entered),
      // then the ESD sheet's value is the source of truth and should be pushed back.
      if (dateMap[key] && esdMap[key] !== undefined && esdMap[key] !== esd) {
        esd = esdMap[key];
        row[9] = esd;
        changed = true;
      }

      // Build ESD row (Project hyperlink formula, etc.)
      const trackerGid = sheet.getSheetId();
      const projectHyperlinkFormula =
        `=HYPERLINK("#gid=${trackerGid}", "${sheetProjectName}")`;

      const dateEntered = dateMap[key] || '';

      outputRows.push([
        projectHyperlinkFormula, // Project (hyperlink)
        po,                      // Purchase Order
        type,                    // Type
        item,                    // Item #
        esd,                     // ESD
        dateEntered              // Date Entered
      ]);
    }

    // If we changed any ESDs on this tracker based on ESD sheet overrides, write back
    if (changed) {
      dataRange.setValues(data);
    }
  });

  // Sort by Project, then PO, then Item #
  outputRows.sort((a, b) => {
    if (a[0] === b[0]) {
      if (a[1] === b[1]) return String(a[3]).localeCompare(String(b[3]));
      return String(a[1]).localeCompare(String(b[1]));
    }
    return String(a[0]).localeCompare(String(b[0]));
  });

  // --- Rebuild ESD sheet ---
  esdSheet.clearContents();
  esdSheet.getRange(1, 1, 1, 6).setValues([[
    'Project',
    'Purchase Order',
    'Type',
    'Item #',
    'ESD',
    'Date Entered'
  ]]);

  if (outputRows.length > 0) {
    // Column A contains formulas (hyperlinks), B–F are plain values
    const formulasColA = outputRows.map(r => [r[0]]);
    const valuesColsBF = outputRows.map(r => r.slice(1));

    esdSheet.getRange(2, 1, outputRows.length, 1).setFormulas(formulasColA);
    esdSheet.getRange(2, 2, outputRows.length, 5).setValues(valuesColsBF);
  }
}

/**
 * When you edit the ESD sheet, push changes back into the
 * corresponding "* - Project Tracker" row(s).
 *
 * NOTE: If you already have an onEdit(e) in this project,
 * merge this logic into that function instead of having two.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  // Only react to edits on the ESD sheet, column E (ESD), data rows
  if (sheetName !== 'ESD') return;
  const row = range.getRow();
  const col = range.getColumn();
  if (row < 2 || col !== 5) return; // ignore header / other columns

  const ss = sheet.getParent();
  const newEsd = range.getValue();

  // Project name is a HYPERLINK formula, so use display value
  const projectName = sheet.getRange(row, 1).getDisplayValue(); // col A
  const po          = sheet.getRange(row, 2).getValue();        // col B
  const type        = sheet.getRange(row, 3).getValue();        // col C
  const item        = sheet.getRange(row, 4).getValue();        // col D

  if (!projectName || !po || !type || !item) return;

  // Stamp Date Entered in col F
  const dateCell = sheet.getRange(row, 6);
  if (newEsd) {
    dateCell.setValue(new Date());
  } else {
    dateCell.clearContent();
  }

  // Push the new ESD value back into the matching tracker sheet rows immediately
  const trackerSheets = ss.getSheets().filter(s =>
    / - Project Tracker$/.test(s.getName())
  );

  trackerSheets.forEach(trackerSheet => {
    const trackerProjectName = trackerSheet.getName().replace(/ - Project Tracker$/, '');
    if (trackerProjectName !== projectName) return;

    const lastRow = trackerSheet.getLastRow();
    if (lastRow < 3) return;

    const numRows = lastRow - 2;
    const dataRange = trackerSheet.getRange(3, 1, numRows, 10); // A–J
    const data = dataRange.getValues();
    let changed = false;

    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const rType = r[2]; // C
      const rItem = r[3]; // D
      const rPo   = r[8]; // I

      if (rPo === po && rType === type && rItem === item) {
        r[9] = newEsd; // J = ESD
        changed = true;
      }
    }

    if (changed) {
      dataRange.setValues(data);
    }
  });
}
