// ============================================================
// PUMA - Open Projects Manifest Audit / Controlled Reconciliation
// ============================================================
//
// Open Projects is a manifest of the folders currently beneath the real
// Open Projects Drive parent. The Project label in col A is intentionally
// allowed to be a business/canonical name that differs from the raw Drive
// folder name. Folder ID (col B) is the stable identity.
//
// This replaces the old "new projects only" checker, which:
// - looked for the wrong Tracker Config capitalization,
// - compared against Tracker Config rather than the actual Open Projects sheet,
// - never identified stale/deleted/moved folders,
// - could not repair replaced folder IDs,
// - and used unsafe first-hyphen project parsing.

const PUMA_OPEN_PROJECTS = {
  SHEET_NAME: 'Open Projects',
  AUDIT_SHEET_NAME: 'PUMA_OPEN_PROJECTS_AUDIT',
  // Verified 2026-10-01. Keep this consistent with TrackerConfig/PUMA Quote Audit.
  PARENT_FOLDER_ID: '1acRZOrQUIzhoIav1Rw8d2GPosaDvNWx5'
};


/**
 * READ ONLY with respect to Open Projects.
 * Writes only the dedicated audit report sheet.
 */
function pumaAuditOpenProjectsManifest() {
  const ss = SpreadsheetApp.getActive();
  const current = pumaReadOpenProjectsManifest_(ss);
  const actual = pumaListOpenProjectFolders_();

  const analysis = pumaAnalyzeOpenProjectsManifest_(current, actual);
  pumaWriteOpenProjectsAudit_(ss, analysis.auditRows);

  const summary =
    'Open Projects audit complete.\n\n' +
    'Current sheet rows: ' + current.length + '\n' +
    'Actual Drive folders: ' + actual.length + '\n' +
    'Missing from sheet: ' + analysis.summary.missing + '\n' +
    'Stale sheet rows: ' + analysis.summary.stale + '\n' +
    'Safe folder-ID replacements: ' + analysis.summary.replacements + '\n' +
    'Name-review required: ' + analysis.summary.review;

  Logger.log(summary);
  try { SpreadsheetApp.getUi().alert(summary); } catch (err) {}

  return analysis;
}


/**
 * Controlled write. Rebuilds Open Projects from the actual Drive parent while:
 * - preserving the existing canonical project name when the same folder ID exists,
 * - safely replacing a stale folder ID when one new folder resolves to the same
 *   canonical identity,
 * - adding new folders only when the project name can be derived conservatively,
 * - excluding stale rows whose folders no longer live under Open Projects.
 *
 * Any unresolved new folder remains in the audit report and is NOT written.
 */
function pumaSyncOpenProjectsManifest() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Reconcile Open Projects',
    'This will rebuild the Open Projects sheet from the actual Drive parent. ' +
      'Canonical names are preserved by folder ID. Unresolved new names are skipped. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActive();
  const current = pumaReadOpenProjectsManifest_(ss);
  const actual = pumaListOpenProjectFolders_();
  const analysis = pumaAnalyzeOpenProjectsManifest_(current, actual);

  const currentById = new Map(current.map(r => [r.folderId, r]));
  const staleByKey = new Map();
  const actualIds = new Set(actual.map(r => r.folderId));

  current.forEach(row => {
    if (actualIds.has(row.folderId)) return;
    const key = normalizeKey_(row.project);
    if (!key) return;
    if (!staleByKey.has(key)) staleByKey.set(key, []);
    staleByKey.get(key).push(row);
  });

  const output = [];
  const unresolved = [];

  actual.forEach(folder => {
    const existing = currentById.get(folder.folderId);
    if (existing) {
      output.push([existing.project, folder.folderId]);
      return;
    }

    const parsed = pumaExtractProjectNameFromFolderDetailed_(folder.rawFolderName);
    const key = normalizeKey_(parsed.projectName);
    const replacements = key ? (staleByKey.get(key) || []) : [];

    if (replacements.length === 1) {
      // Folder was likely replaced/moved but the business identity stayed the same.
      output.push([replacements[0].project, folder.folderId]);
      return;
    }

    if (parsed.safe && parsed.projectName) {
      output.push([parsed.projectName, folder.folderId]);
      return;
    }

    unresolved.push(folder);
  });

  output.sort((a, b) => String(a[0]).localeCompare(String(b[0])));

  let sheet = ss.getSheetByName(PUMA_OPEN_PROJECTS.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(PUMA_OPEN_PROJECTS.SHEET_NAME);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 2).setValues([['Folder Name', 'Folder ID']]);
  if (output.length) {
    sheet.getRange(2, 1, output.length, 2).setValues(output);
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);

  // Re-run the audit after the controlled write.
  const post = pumaAuditOpenProjectsManifest();

  ui.alert(
    'Open Projects reconciliation complete.\n\n' +
    'Rows written: ' + output.length + '\n' +
    'Unresolved folders skipped: ' + unresolved.length + '\n' +
    'Review: ' + PUMA_OPEN_PROJECTS.AUDIT_SHEET_NAME
  );

  return {written: output.length, unresolved: unresolved, audit: post};
}


/**
 * Backward-compatible entry point from the previous toolbar/workflow.
 */
function pumaFindNewOpenProjectFolders() {
  return pumaAuditOpenProjectsManifest();
}


function pumaReadOpenProjectsManifest_(ss) {
  const sheet = ss.getSheetByName(PUMA_OPEN_PROJECTS.SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 2)
    .getValues()
    .map((row, i) => ({
      project: String(row[0] || '').trim(),
      folderId: String(row[1] || '').trim(),
      sourceRow: i + 2
    }))
    .filter(r => r.project && r.folderId);
}


function pumaListOpenProjectFolders_() {
  const parent = DriveApp.getFolderById(PUMA_OPEN_PROJECTS.PARENT_FOLDER_ID);
  const iter = parent.getFolders();
  const rows = [];

  while (iter.hasNext()) {
    const folder = iter.next();
    rows.push({
      rawFolderName: String(folder.getName() || '').trim(),
      folderId: folder.getId(),
      url: folder.getUrl()
    });
  }

  rows.sort((a, b) => a.rawFolderName.localeCompare(b.rawFolderName));
  return rows;
}


function pumaAnalyzeOpenProjectsManifest_(current, actual) {
  const actualById = new Map(actual.map(r => [r.folderId, r]));
  const currentById = new Map(current.map(r => [r.folderId, r]));
  const actualIds = new Set(actual.map(r => r.folderId));

  const stale = current.filter(r => !actualIds.has(r.folderId));
  const staleByKey = new Map();

  stale.forEach(row => {
    const key = normalizeKey_(row.project);
    if (!key) return;
    if (!staleByKey.has(key)) staleByKey.set(key, []);
    staleByKey.get(key).push(row);
  });

  const auditRows = [[
    'Status',
    'Canonical Project',
    'Sheet Folder ID',
    'Actual Drive Folder Name',
    'Actual Folder ID',
    'Proposed Project',
    'Safe?',
    'Reason'
  ]];

  let missing = 0;
  let replacements = 0;
  let review = 0;

  actual.forEach(folder => {
    const existing = currentById.get(folder.folderId);

    if (existing) {
      auditRows.push([
        'CURRENT',
        existing.project,
        existing.folderId,
        folder.rawFolderName,
        folder.folderId,
        existing.project,
        'YES',
        existing.project === folder.rawFolderName
          ? 'Folder ID and name are current.'
          : 'Folder ID is current; canonical project label intentionally differs from raw folder name.'
      ]);
      return;
    }

    missing++;
    const parsed = pumaExtractProjectNameFromFolderDetailed_(folder.rawFolderName);
    const key = normalizeKey_(parsed.projectName);
    const candidateStale = key ? (staleByKey.get(key) || []) : [];

    if (candidateStale.length === 1) {
      replacements++;
      auditRows.push([
        'REPLACED_FOLDER_ID',
        candidateStale[0].project,
        candidateStale[0].folderId,
        folder.rawFolderName,
        folder.folderId,
        candidateStale[0].project,
        'YES',
        'A stale manifest row has the same conservative project identity; preserve its canonical name and replace the folder ID.'
      ]);
      return;
    }

    if (parsed.safe && parsed.projectName) {
      auditRows.push([
        'NEW_SAFE',
        '',
        '',
        folder.rawFolderName,
        folder.folderId,
        parsed.projectName,
        'YES',
        parsed.reason
      ]);
      return;
    }

    review++;
    auditRows.push([
      'NEW_REVIEW',
      '',
      '',
      folder.rawFolderName,
      folder.folderId,
      parsed.projectName || '',
      'NO',
      parsed.reason || 'Project identity could not be derived conservatively.'
    ]);
  });

  stale.forEach(row => {
    // If this stale identity has exactly one actual replacement candidate,
    // the replacement row above already explains it. Keep the stale row visible
    // as well for a complete manifest diff.
    auditRows.push([
      'STALE_IN_SHEET',
      row.project,
      row.folderId,
      '',
      '',
      '',
      'NO',
      'Folder ID is no longer a direct child of the real Open Projects parent.'
    ]);
  });

  return {
    auditRows: auditRows,
    summary: {
      missing: missing,
      stale: stale.length,
      replacements: replacements,
      review: review
    }
  };
}


/**
 * Conservative folder-name parser.
 *
 * We NEVER split arbitrary hyphenated names such as "Bryson-Kleine Residence".
 * A suffix is removed only when:
 * - it is a known workflow/client suffix; OR
 * - the left side ends in a strong project-identity word such as Residence,
 *   House, Property, Point, etc.
 */
function pumaExtractProjectNameFromFolderDetailed_(folderName) {
  let name = String(folderName || '').trim();
  if (!name) return {projectName: '', safe: false, reason: 'Blank folder name.'};

  const knownSuffix = /\s*-\s*(owner|night\s+design|yoakum|ncid|fe\s+trainer|hickory(?:\s+construction)?|parallel|knight\s*vision|knightvision|arcadia(?:\s+schaad)?|purvis(?:\s+builders|\s+construction)?|schmi?dt?\s*&\s*rhodes|copper\s+inspired|lady\s+katie\s+creations|norris\s+studio|bazan\s+construction|spaces\s+in\s+the\s+city|flatrock\s+knightvison|lamon\s*&\s*mcdaniel\s+builders|johnson\s*&\s+gaylen|avs-wood\s+point|energy\s+electric)\s*$/i;

  if (knownSuffix.test(name)) {
    return {
      projectName: name.replace(knownSuffix, '').trim(),
      safe: true,
      reason: 'Removed a recognized contractor/client suffix.'
    };
  }

  // Safe structural split: only use the LAST hyphen and only when the left
  // side itself clearly looks like a complete project identity.
  const lastDash = name.lastIndexOf('-');
  if (lastDash > 0) {
    const left = name.slice(0, lastDash).trim();
    const projectEnding = /\b(residence|house|court|villa|property|restaurant|hotel|project|natatorium|view|shore|farmhouse|point|cottage|pool\s+house|mountain\s+house|mtn\.?\s+house)\s*$/i;
    if (projectEnding.test(left)) {
      return {
        projectName: left,
        safe: true,
        reason: 'Removed a suffix after a complete project-identity phrase.'
      };
    }
  }

  // No delimiter means the whole folder name is a safe project candidate.
  if (!name.includes('-')) {
    return {
      projectName: name,
      safe: true,
      reason: 'Folder name contains no project/client delimiter.'
    };
  }

  return {
    projectName: name,
    safe: false,
    reason: 'Hyphenated folder name could be a real project name; manual review required.'
  };
}


function extractProjectNameFromFolder_(folderName) {
  return pumaExtractProjectNameFromFolderDetailed_(folderName).projectName;
}


function buildTrackerNameSet_(sheet, headerText) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return new Set();

  const headers = values[0].map(h => String(h || '').trim().toLowerCase());
  const colIndex = headers.findIndex(h => h === String(headerText || '').toLowerCase());
  if (colIndex === -1) {
    throw new Error(`Header "${headerText}" not found in ${sheet.getName()}`);
  }

  const set = new Set();
  for (let r = 1; r < values.length; r++) {
    const name = values[r][colIndex];
    if (name) set.add(normalizeKey_(name));
  }
  return set;
}


function normalizeKey_(s) {
  if (typeof pumaNormalizeProjectKey_ === 'function') {
    return pumaNormalizeProjectKey_(s);
  }

  return String(s || '')
    .toLowerCase()
    .replace(/\b(residence|project)\b/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


function pumaWriteOpenProjectsAudit_(ss, rows) {
  let sheet = ss.getSheetByName(PUMA_OPEN_PROJECTS.AUDIT_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(PUMA_OPEN_PROJECTS.AUDIT_SHEET_NAME);

  sheet.clearContents();
  sheet.clearFormats();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, rows[0].length)
    .setFontWeight('bold')
    .setBackground('#cfe2f3');

  if (rows.length > 1) {
    const statusRange = sheet.getRange(2, 1, rows.length - 1, 1);
    const vals = statusRange.getValues();
    statusRange.setBackgrounds(vals.map(row => {
      const s = String(row[0] || '').toUpperCase();
      if (s === 'CURRENT') return ['#d9ead3'];
      if (s === 'REPLACED_FOLDER_ID' || s === 'NEW_SAFE') return ['#fff2cc'];
      if (s === 'NEW_REVIEW' || s === 'STALE_IN_SHEET') return ['#f4cccc'];
      return ['#ffffff'];
    }));
  }

  sheet.autoResizeColumns(1, rows[0].length);
}


function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
