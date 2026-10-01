/************************************************************
 * PUMA_QuoteAudit.js
 *
 * Purpose:
 * 1) REPORT ONLY:
 *    Scan Open Projects folders for quotation Google Sheets.
 *    Compare found quote spreadsheet IDs against Tracker config.
 *    Write missing quote actions to PUMA_QUOTE_AUDIT_ACTIONS.
 *
 * 2) AUTO ADD:
 *    Read PUMA_QUOTE_AUDIT_ACTIONS rows created by this report.
 *    Add missing quote spreadsheet IDs to Tracker config.
 *
 * Requires:
 * - Advanced Google Service enabled: Drive API
 * - Sheet: Open Projects
 *   Col A = Project / Folder Name
 *   Col B = Folder ID
 * - Sheet: Tracker config
 *   A Enable?
 *   B Project
 *   C Tracker Sheet Name
 *   D Date Updated
 *   E Quote Name
 *   F Quote Sheet ID
 ************************************************************/

const PUMA_QUOTE_AUDIT = {
  OPEN_PROJECTS_SHEET: 'Open Projects',
  TRACKER_CONFIG_SHEET: 'Tracker config',
  ACTIONS_SHEET: 'PUMA_QUOTE_AUDIT_ACTIONS',
  OPEN_PROJECTS_PARENT_FOLDER_ID: '1acRZOrQUIzhoIav1Rw8d2GPosaDvNWx5',

  ACTION_CLASSIFICATION: 'QUOTE_SYNC',
  ACTION_ADD_MISSING: 'ADD MISSING QUOTE TO CONFIG',
  ACTION_OK: 'QUOTE SYNC OK',
  ACTION_PENDING: 'PENDING QUOTES',

  ACTION_HEADERS: [
    'Priority',
    'Recommended Action',
    'Sheet Name',
    'Classification',
    'Config Match',
    'Expected Paired Sheet',
    'Pair Exists',
    'Risk Level',
    'Reason',
    'Owner Notes',
    'Source Row',
    'Folder ID',
    'Scan Result',
    'Tracker Exists?',
    'Tracker Tab Name',
    'Tasks Exists?',
    'Done?'
  ]
};

/**
 * REPORT ONLY.
 *
 * Scans all Open Projects folders, finds quote Google Sheets,
 * compares them to Tracker config, and writes action rows.
 *
 * Does NOT modify Tracker config.
 */
function reportPumaMissingQuotes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const openProjects = readPumaOpenProjects_(ss);
  const config = readPumaTrackerConfig_(ss);

  const rows = [PUMA_QUOTE_AUDIT.ACTION_HEADERS];
  const now = new Date();
  const addedActionKeys = new Set();

  const projectsToScan = openProjects;
  const scannedFolderIds = new Set();

  projectsToScan.forEach(project => {
    const sheetStatus = getPumaProjectSheetStatus_(ss, project.project);

    let foundQuotes = [];

    try {
      foundQuotes = listPumaQuoteSheetsInFolder_(project.folderId);
      scannedFolderIds.add(project.folderId);
    } catch (err) {
      rows.push([
        1,
        'FOLDER SCAN ERROR',
        project.project,
        PUMA_QUOTE_AUDIT.ACTION_CLASSIFICATION,
        '',
        '',
        project.folderId,
        'HIGH',
        `Could not scan project folder. Folder ID may be invalid, deleted, moved, or inaccessible. Error: ${err.message}`,
        '',
        project.sourceRow,
        project.folderId,
        'ERROR',
        sheetStatus.trackerExists ? 'YES' : 'NO',
        sheetStatus.trackerName,
        sheetStatus.tasksExists ? 'YES' : 'NO',
        false
      ]);
      return;
    }

    const configuredQuoteIds =
      config.byProject.get(normalizePumaKey_(project.project)) || new Set();

    const foundCount = foundQuotes.length;
    let configuredCount = 0;
    const missing = [];

    foundQuotes.forEach(q => {
      if (configuredQuoteIds.has(String(q.id))) {
        configuredCount++;
      } else {
        missing.push(q);
      }
    });

    // PROJECT-LEVEL SCAN MANIFEST ROW
    let scanStatus = '';
    let riskLevel = 'OK';
    let reason = '';

    if (foundCount === 0) {
      scanStatus = 'PENDING QUOTES';
      riskLevel = 'REVIEW';
      reason = `Folder scanned successfully. No quotation Google Sheets found as of ${formatPumaDateTime_(now)}.`;
    } else if (missing.length === 0) {
      scanStatus = PUMA_QUOTE_AUDIT.ACTION_OK;
      riskLevel = 'OK';
      reason = `Folder scanned successfully. All detected quote spreadsheets are already present in Tracker config.`;
    } else {
      scanStatus = `OK - ${configuredCount}/${foundCount} Quotes Syncing`;
      riskLevel = 'HIGH';
      reason = `Folder scanned successfully. ${missing.length} quote spreadsheet(s) found in Drive but missing from Tracker config.`;
    }

    rows.push([
      missing.length ? 3 : 5,
      scanStatus,
      project.project,
      PUMA_QUOTE_AUDIT.ACTION_CLASSIFICATION,
      `${configuredCount}/${foundCount} Quotes Syncing`,
      '',
      '',
      riskLevel,
      reason,
      '',
      project.sourceRow,
      project.folderId,
      'SCANNED',
      sheetStatus.trackerExists ? 'YES' : 'NO',
      sheetStatus.trackerName,
      sheetStatus.tasksExists ? 'YES' : 'NO',
      false
    ]);

    // ACTION ROWS FOR EACH MISSING QUOTE
    missing.forEach(q => {
      const actionKey = `${normalizePumaKey_(project.project)}|${q.id}`;
      if (addedActionKeys.has(actionKey)) return;
      addedActionKeys.add(actionKey);

      rows.push([
        2,
        PUMA_QUOTE_AUDIT.ACTION_ADD_MISSING,
        project.project,
        PUMA_QUOTE_AUDIT.ACTION_CLASSIFICATION,
        `${configuredCount}/${foundCount} Quotes Syncing`,
        q.name,
        q.id,
        'HIGH',
        `Found in Drive folder but missing from Tracker config. Missing Spreadsheet ID = ${q.id}`,
        '',
        project.sourceRow,
        project.folderId,
        'SCANNED',
        sheetStatus.trackerExists ? 'YES' : 'NO',
        sheetStatus.trackerName,
        sheetStatus.tasksExists ? 'YES' : 'NO',
        false
      ]);
    });
  });

  writePumaQuoteAuditActions_(ss, rows);

  const message =
  `PUMA Quote Audit complete.\n\n` +
  `Open Projects rows read: ${openProjects.length}\n` +
  `Unique folders scanned: ${scannedFolderIds.size}\n` +
  `Output rows created: ${Math.max(0, rows.length - 1)}`;

  Logger.log(message);

  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (err) {
    // Running from Apps Script editor, not the Sheet UI.}
}
}


/**
 * AUTO ADD.
 *
 * Reads PUMA_QUOTE_AUDIT_ACTIONS and adds rows to Tracker config
 * for missing quote IDs detected by reportPumaMissingQuotes().
 *
 * Only touches rows where:
 * - Classification = QUOTE_SYNC
 * - Recommended Action = ADD MISSING QUOTE TO CONFIG
 * - Done? is not checked
 */
function autoAddPumaMissingQuotesToConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const actionsSheet = ss.getSheetByName(PUMA_QUOTE_AUDIT.ACTIONS_SHEET);
  if (!actionsSheet) {
    throw new Error(`Missing sheet: ${PUMA_QUOTE_AUDIT.ACTIONS_SHEET}`);
  }

  const configSheet = getOrCreatePumaTrackerConfigSheet_(ss);
  ensurePumaTrackerConfigHeaders_(configSheet);

  const actionValues = actionsSheet.getDataRange().getValues();
  if (actionValues.length < 2) {
    SpreadsheetApp.getUi().alert('No action rows found.');
    return;
  }

  const headers = actionValues[0].map(String);
  const idx = indexPumaHeaders_(headers, [
    'Priority',
    'Recommended Action',
    'Sheet Name',
    'Classification',
    'Config Match',
    'Expected Paired Sheet',
    'Pair Exists',
    'Risk Level',
    'Reason',
    'Owner Notes',
    'Source Row',
    'Folder ID',
    'Scan Result',
    'Tracker Exists?',
    'Tracker Tab Name',
    'Tasks Exists?',
    'Done?'
  ]);

  const existingConfigIds = readPumaTrackerConfig_(ss).allQuoteIds;

  let added = 0;
  let skipped = 0;
  const rowsToMarkDone = [];

  actionValues.slice(1).forEach((row, offset) => {
    const sheetRowNumber = offset + 2;

    const action = String(row[idx['Recommended Action']] || '').trim();
    const classification = String(row[idx['Classification']] || '').trim();
    const project = String(row[idx['Sheet Name']] || '').trim();
    const quoteName = String(row[idx['Expected Paired Sheet']] || '').trim();
    const quoteId = String(row[idx['Pair Exists']] || '').trim();
    const done = row[idx['Done?']] === true;

    if (done) return;
    if (classification !== PUMA_QUOTE_AUDIT.ACTION_CLASSIFICATION) return;
    if (action !== PUMA_QUOTE_AUDIT.ACTION_ADD_MISSING) return;

    if (!project || !quoteName || !quoteId) {
      skipped++;
      return;
    }

    if (existingConfigIds.has(quoteId)) {
      skipped++;
      rowsToMarkDone.push(sheetRowNumber);
      return;
    }

    configSheet.appendRow([
      true,
      project,
      `${project} - Project Tracker`,
      new Date(),
      quoteName,
      quoteId
    ]);

    existingConfigIds.add(quoteId);
    added++;
    rowsToMarkDone.push(sheetRowNumber);
  });

  // Mark completed rows as Done.
  rowsToMarkDone.forEach(rowNum => {
    actionsSheet.getRange(rowNum, idx['Done?'] + 1).setValue(true);
  });

  configSheet.autoResizeColumns(1, 6);

  const message =
  `Auto-add complete.\n\n` +
  `Added to Tracker config: ${added}\n` +
  `Skipped/already present: ${skipped}`;

  Logger.log(message);

  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (err) {
    // Running from Apps Script editor, not the Sheet UI.
  }
}

/**
 * Reads Open Projects.
 * Expected:
 * Col A = Project name
 * Col B = Folder ID
 */
function readPumaOpenProjects_(ss) {
  const sheet = ss.getSheetByName(PUMA_QUOTE_AUDIT.OPEN_PROJECTS_SHEET);
  if (!sheet) {
    throw new Error(`Missing sheet: ${PUMA_QUOTE_AUDIT.OPEN_PROJECTS_SHEET}`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  const seen = new Set();

  return values
    .map((row, index) => ({
      project: String(row[0] || '').trim(),
      folderId: String(row[1] || '').trim(),
      sourceRow: index + 2
    }))

    .filter(r => {
      if (!r.project || !r.folderId) return false;

      const projectLower = r.project.toLowerCase();
      const folderLower = r.folderId.toLowerCase();

      if (projectLower === 'folder name') return false;
      if (folderLower === 'folder id') return false;
      if (projectLower.includes('folder name')) return false;
      if (folderLower.includes('folder id')) return false;

      const key = `${normalizePumaKey_(r.project)}|${r.folderId}`;
      if (seen.has(key)) return false;

      seen.add(key);
      return true;
    });
}

/**
 * Reads Tracker config.
 * Expected columns:
 * A Enable?
 * B Project
 * C Tracker Sheet Name
 * D Date Updated
 * E Quote Name
 * F Quote Sheet ID
 */
function readPumaTrackerConfig_(ss) {
  const sheet = ss.getSheetByName(PUMA_QUOTE_AUDIT.TRACKER_CONFIG_SHEET);

  const byProject = new Map();
  const allQuoteIds = new Set();

  if (!sheet || sheet.getLastRow() < 2) {
    return { byProject, allQuoteIds };
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();

  values.forEach(row => {
    const project = String(row[1] || '').trim();
    const quoteId = String(row[5] || '').trim();
    if (!project || !quoteId) return;

    const key = normalizePumaKey_(project);
    if (!byProject.has(key)) byProject.set(key, new Set());

    byProject.get(key).add(quoteId);
    allQuoteIds.add(quoteId);
  });

  return { byProject, allQuoteIds };
}

/**
 * Shared Drive-safe listing of Google Sheets in a folder.
 * Then filters to likely quotation sheets.
 */
function listPumaQuoteSheetsInFolder_(folderId) {
  const results = [];
  let pageToken;

  const q = [
    `'${folderId}' in parents`,
    `mimeType='application/vnd.google-apps.spreadsheet'`,
    `trashed=false`
  ].join(' and ');

  do {
    const resp = Drive.Files.list({
      q,
      fields: 'nextPageToken, files(id, name)',
      pageToken: pageToken,
      pageSize: 1000,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    });

    (resp.files || []).forEach(file => {
      const classification = classifyPumaQuotationFile_(file);
      if (classification.isQuote) {
        results.push({
          id: file.id,
          name: file.name,
          confidence: classification.confidence,
          reason: classification.reason
        });
      }
    });

    pageToken = resp.nextPageToken;
  } while (pageToken);

  results.sort((a, b) => a.name.localeCompare(b.name));
  return results;
}


/**
 * Determine whether a Google Sheet is a PUMA quotation.
 *
 * Name alone is not authoritative. We inspect the workbook structure so
 * non-standard but legitimate names (e.g. "Sayler Residence Deco") can pass,
 * while obvious billing/inventory sheets do not get imported just because
 * they contain the word "lighting".
 */
function classifyPumaQuotationFile_(file) {
  const name = String(file && file.name || '').trim();
  const lower = name.toLowerCase();

  const obviousNegative =
    /\b(billing|invoice|inventory|packing slip|receiving|delivery report|tracker|tasks|dashboard|summary)\b/i.test(lower) ||
    /^alig\b/i.test(lower);

  const strongName =
    /\b(quote|quotation|adder|architectural|decorative|heater|landscape lighting|lighting quotation|downlighting|reconcilliation|reconciliation)\b/i.test(lower);

  try {
    const qss = SpreadsheetApp.openById(file.id);
    const sheets = qss.getSheets();
    if (!sheets.length) {
      return {isQuote: false, confidence: 'NONE', reason: 'Spreadsheet has no tabs.'};
    }

    // Inspect up to the first 5 tabs and first 15 rows. Quote workbooks vary
    // between Master Quotation, Approved, and dated/working tabs.
    for (let s = 0; s < Math.min(sheets.length, 5); s++) {
      const sheet = sheets[s];
      const lastCol = Math.max(1, Math.min(sheet.getLastColumn(), 16));
      const lastRow = Math.max(1, Math.min(sheet.getLastRow(), 15));
      const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

      const flattened = values
        .reduce((acc, row) => acc.concat(row), [])
        .map(v => String(v || '').toLowerCase());

      const joined = flattened.join(' | ');
      const hasQuotationTitle = joined.includes('quotation');
      const hasQty = flattened.some(v => /^(qty|quantity|qnty)$/i.test(v.trim()));
      const hasDescription = flattened.some(v => v.trim() === 'description');
      const hasManufacturer = flattened.some(v => /^(manufacturer|mfg|vendor)$/i.test(v.trim()));
      const hasPart = flattened.some(v => /^(part number|part #|part no\.?|product number)$/i.test(v.trim()));
      const hasPricing = flattened.some(v =>
        /cost per unit|unit price|total cost|margin|extended|sell price/i.test(v)
      );

      const structuralHeaders = hasQty && hasDescription && hasManufacturer && hasPart;
      if (hasQuotationTitle || structuralHeaders || (hasQty && hasPart && hasPricing)) {
        return {
          isQuote: true,
          confidence: strongName ? 'HIGH' : 'STRUCTURAL',
          reason: strongName
            ? 'Quote-like name and quotation structure detected.'
            : 'Quotation structure detected despite non-standard file name.'
        };
      }
    }

    if (obviousNegative) {
      return {isQuote: false, confidence: 'HIGH', reason: 'Non-quotation business sheet by name/structure.'};
    }

    return {
      isQuote: false,
      confidence: strongName ? 'REVIEW' : 'NONE',
      reason: strongName
        ? 'Name looks like a quote but quotation structure was not detected.'
        : 'No quotation structure detected.'
    };

  } catch (err) {
    // Failure to inspect is never enough for an automatic add. Strong names
    // remain audit candidates, but are not treated as confirmed quotes.
    return {
      isQuote: false,
      confidence: 'REVIEW',
      reason: 'Could not inspect workbook: ' + err.message
    };
  }
}


function isLikelyPumaQuotation_(file) {
  return classifyPumaQuotationFile_(file).isQuote;
}

/**
 * Writes quote audit rows to PUMA_QUOTE_AUDIT_ACTIONS.
 *
 * This intentionally clears PUMA_QUOTE_AUDIT_ACTIONS because this function is
 * a focused quote-sync report. If you want to combine workbook audit actions
 * and quote audit actions later, we can merge this with buildPumaAuditActions().
 */
function writePumaQuoteAuditActions_(ss, rows) {
  let sheet = ss.getSheetByName(PUMA_QUOTE_AUDIT.ACTIONS_SHEET);
  if (!sheet) sheet = ss.insertSheet(PUMA_QUOTE_AUDIT.ACTIONS_SHEET);

  sheet.clearContents();
  sheet.clearFormats();

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);

  const header = sheet.getRange(1, 1, 1, rows[0].length);
  header.setFontWeight('bold');
  header.setBackground('#d9ead3');

  if (rows.length > 1) {
    sheet
      .getRange(2, 1, rows.length - 1, rows[0].length)
      .sort([
        { column: 1, ascending: true },
        { column: 2, ascending: true },
        { column: 3, ascending: true }
      ]);

    // Done? checkbox column.
    sheet.getRange(2, 17, rows.length - 1, 1).insertCheckboxes();

    const actionRange = sheet.getRange(2, 2, rows.length - 1, 1);
    const actions = actionRange.getValues();

    const backgrounds = actions.map(row => {
      const action = String(row[0] || '').toUpperCase();
      if (action.includes('ADD MISSING')) return ['#f4cccc'];
      if (action.includes('OK')) return ['#d9ead3'];
      if (action.includes('PENDING')) return ['#fff2cc'];
      return ['#cfe2f3'];
    });

    actionRange.setBackgrounds(backgrounds);
  }

  sheet.autoResizeColumns(1, rows[0].length);
}

/**
 * Creates Tracker config if missing.
 */
function getOrCreatePumaTrackerConfigSheet_(ss) {
  let sheet = ss.getSheetByName(PUMA_QUOTE_AUDIT.TRACKER_CONFIG_SHEET);
  if (!sheet) sheet = ss.insertSheet(PUMA_QUOTE_AUDIT.TRACKER_CONFIG_SHEET);
  return sheet;
}

/**
 * Ensures Tracker config has the expected headers.
 */
function ensurePumaTrackerConfigHeaders_(sheet) {
  const headers = [
    'Enable?',
    'Project',
    'Tracker Sheet Name',
    'Date Updated',
    'Quote Name',
    'Quote Sheet ID'
  ];

  const existing = sheet
    .getRange(1, 1, 1, headers.length)
    .getValues()[0]
    .map(String);

  const matches = headers.every((h, i) => existing[i] === h);

  if (!matches) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).insertCheckboxes();
  }
}

/**
 * Helper: required header index lookup.
 */
function indexPumaHeaders_(headers, requiredHeaders) {
  const idx = {};

  requiredHeaders.forEach(h => {
    const i = headers.indexOf(h);
    if (i === -1) {
      throw new Error(`Missing required column in PUMA_QUOTE_AUDIT_ACTIONS: ${h}`);
    }
    idx[h] = i;
  });

  return idx;
}

/**
 * Helper: normalize project names for matching.
 */
function normalizePumaKey_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+-\s+project tracker$/i, '')
    .replace(/\s+-\s+tasks$/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}


/**
 * Helper: Build a list of folders currently inside the Open Projects Drive folder, then only scan IDs that are in that list.
 */
function getCurrentOpenProjectFolderIds_() {
  const folderIds = new Set();

  const parent = DriveApp.getFolderById(
    PUMA_QUOTE_AUDIT.OPEN_PROJECTS_PARENT_FOLDER_ID
  );

  const folders = parent.getFolders();

  while (folders.hasNext()) {
    const folder = folders.next();
    folderIds.add(folder.getId());
  }

  return folderIds;
}

/**
 * Helper: Check if expected Project Tracker and Tasks tabs exist for a project, to help prioritize which quote sync issues to review first.
 */
function getPumaProjectSheetStatus_(ss, projectName) {
  const trackerName = `${projectName} - Project Tracker`;
  const tasksName = `${projectName} - Tasks`;

  return {
    trackerName,
    trackerExists: !!ss.getSheetByName(trackerName),
    tasksExists: !!ss.getSheetByName(tasksName)
  };
}

/**
 * Helper: readable timestamp.
 */
function formatPumaDateTime_(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'M/d/yyyy h:mm a'
  );
}

function testOpenProjectsParent() {
  const folderId = '1S6m5hsxpkyt1GtcGWo1Bat-dXMCCsiIR'; // Willoughby Residence

  const folder = Drive.Files.get(folderId, {
    supportsAllDrives: true
  });

  Logger.log(JSON.stringify(folder, null, 2));
}