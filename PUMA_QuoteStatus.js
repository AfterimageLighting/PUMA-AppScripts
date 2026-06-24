/************************************************************
 * PUMA Quote Status
 *
 * Adds quote/project closing status to Tracker config.
 *
 * Writes:
 * G = Closing %
 * H = Closing Status Note
 *
 * Rules:
 * 100%  = quote has an Approved tab AND that tab is locked/protected
 * >85%  = quote exists but is not Approved + locked
 * <85%  = project folder has no quotation yet
 ************************************************************/

function updateTrackerConfigQuoteStatuses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName('Tracker config');
  const openProjects = ss.getSheetByName('Open Projects');

  if (!config) throw new Error('Missing sheet: Tracker config');
  if (!openProjects) throw new Error('Missing sheet: Open Projects');

  ensureQuoteStatusHeaders_(config);

  const lastRow = config.getLastRow();
  if (lastRow < 2) return;

  const values = config.getRange(2, 1, lastRow - 1, 8).getValues();

  const output = values.map(row => {
     const project = String(row[1] || '').trim();   // B
     const quoteName = String(row[4] || '').trim(); // E
     const quoteId = String(row[5] || '').trim();   // F

     if (!project) return ['', ''];

     const isNotYetQuoted =
      !quoteId ||
       quoteName.toLowerCase() === 'not yet quoted' ||
       quoteName.toLowerCase().includes('not yet quoted');

     if (isNotYetQuoted) {
       return ['<85%', 'Not yet quoted'];
     }

     const quoteStatus = getQuoteApprovedLockStatus_(quoteId);

     if (quoteStatus.approved && quoteStatus.locked) {
       return ['100%', 'Approved quotation is locked'];
     }

     if (quoteStatus.approved && !quoteStatus.locked) {
       return ['>85%', 'Approved quotation is NOT locked'];
     }

     return ['>85%', 'Quoted, not approved'];
  });

  config.getRange(2, 7, output.length, 2).setValues(output);
  formatQuoteStatusColumns_(config);
}


/**
 * Checks whether a quotation should count as 100%.
 */
function getQuoteApprovedLockStatus_(quoteSpreadsheetId) {
  const result = {
    approved: false,
    locked: false
  };

  if (!quoteSpreadsheetId) return result;

  try {
    const qss = SpreadsheetApp.openById(quoteSpreadsheetId);

    const approvedSheets = qss.getSheets().filter(sheet =>
      String(sheet.getName() || '').toLowerCase().includes('approved')
    );

    if (!approvedSheets.length) return result;

    result.approved = true;

    for (const sheet of approvedSheets) {
      const sheetProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      const rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

      const hasLockedSheet = sheetProtections.some(p => !p.isWarningOnly());
      const hasLockedRange = rangeProtections.some(p => !p.isWarningOnly());

      if (hasLockedSheet || hasLockedRange) {
        result.locked = true;
        return result;
      }
    }

    return result;

  } catch (err) {
    return result;
  }
}


function ensureQuoteStatusHeaders_(sheet) {
  sheet.getRange(1, 7).setValue('Closing %');
  sheet.getRange(1, 8).setValue('Closing Status Note');

  sheet.getRange(1, 7, 1, 2)
    .setFontWeight('bold')
    .setBackground('#d9e2f3')
    .setHorizontalAlignment('center');
}


function formatQuoteStatusColumns_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 7, lastRow - 1, 1);
  const values = range.getDisplayValues();

  const backgrounds = values.map(row => {
    const status = String(row[0] || '').trim();

    if (status === '100%') return ['#d9ead3'];  // green
    if (status === '>85%') return ['#fff2cc'];  // yellow
    if (status === '<85%') return ['#f4cccc'];  // red

    return ['#ffffff'];
  });

  range
    .setBackgrounds(backgrounds)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');

  sheet.autoResizeColumns(7, 2);
}