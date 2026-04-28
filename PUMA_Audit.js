/**
 * PUMA AUDIT
 * Builds a workbook inventory and risk report.
 *
 * Output sheet:
 *   PUMA_AUDIT
 */

const PUMA_AUDIT = {
  OUTPUT_SHEET: 'PUMA_AUDIT',
  DASHBOARD: 'Dashboard',
  TRACKER_CONFIG: 'Tracker Config',
  OPEN_PROJECTS: 'Open Projects',
  HEADER_ROW: 1
};

function runPumaAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditRows = [];

  auditRows.push([
    'Timestamp',
    'Sheet Name',
    'Classification',
    'Rows',
    'Columns',
    'Frozen Rows',
    'Has A1 Dashboard Link',
    'Config Match',
    'Risk Level',
    'Notes'
  ]);

  const timestamp = new Date();
  const configProjects = getAuditTrackerConfigProjects_(ss);

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name === PUMA_AUDIT.OUTPUT_SHEET) return;

    const classification = classifyPumaSheet_(name);
    const hasDashboardLink = hasBackToDashboardLink_(sheet);
    const normalizedProject = normalizeAuditProjectName_(name);
    const configMatch = configProjects.has(normalizedProject) ? 'YES' : 'NO';

    const risk = getAuditRiskLevel_(name, classification, hasDashboardLink, configMatch);

    auditRows.push([
      timestamp,
      name,
      classification,
      sheet.getMaxRows(),
      sheet.getMaxColumns(),
      sheet.getFrozenRows(),
      hasDashboardLink ? 'YES' : 'NO',
      shouldCheckConfigMatch_(classification) ? configMatch : '',
      risk.level,
      risk.notes
    ]);
  });

  writePumaAudit_(ss, auditRows);
  SpreadsheetApp.getUi().alert(`PUMA audit complete.\nSheets scanned: ${auditRows.length - 1}`);
}

function getAuditTrackerConfigProjects_(ss) {
  const set = new Set();
  const sh = ss.getSheetByName(PUMA_AUDIT.TRACKER_CONFIG);
  if (!sh) return set;

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return set;

  const headers = values[0].map(h => String(h || '').trim().toLowerCase());

  const projectIdx = findHeaderIndex_(headers, [
    'project',
    'project name',
    'projectname'
  ]);

  const trackerIdx = findHeaderIndex_(headers, [
    'tracker sheet name',
    'tracker',
    'tracker tab',
    'tracker sheet'
  ]);

  for (let r = 1; r < values.length; r++) {
    if (projectIdx !== -1 && values[r][projectIdx]) {
      set.add(normalizeAuditProjectName_(values[r][projectIdx]));
    }

    if (trackerIdx !== -1 && values[r][trackerIdx]) {
      set.add(normalizeAuditProjectName_(values[r][trackerIdx]));
    }
  }

  return set;
}

function classifyPumaSheet_(sheetName) {
  const name = String(sheetName || '').trim();

  if (name === PUMA_AUDIT.DASHBOARD) return 'CORE_DASHBOARD';
  if (name === PUMA_AUDIT.TRACKER_CONFIG) return 'CORE_CONFIG';
  if (name === PUMA_AUDIT.OPEN_PROJECTS) return 'CORE_PROJECT_LIST';

  if (/project tracker$/i.test(name)) return 'PROJECT_TRACKER';
  if (/project track$/i.test(name)) return 'PROJECT_TRACKER_SHORT_NAME';
  if (/tasks$/i.test(name)) return 'PROJECT_TASKS';

  if (/raw_po_import/i.test(name)) return 'PO_PIPELINE';
  if (/esd/i.test(name)) return 'ESD_PIPELINE';
  if (/quote summary/i.test(name)) return 'QUOTE_SUMMARY';
  if (/rfp|rfps|email/i.test(name)) return 'LEGACY_EMAIL_RFPS';

  if (/config/i.test(name)) return 'CONFIG_OR_HELPER';
  if (/template/i.test(name)) return 'TEMPLATE_OR_HELPER';

  return 'UNKNOWN_REVIEW';
}

function hasBackToDashboardLink_(sheet) {
  const cell = sheet.getRange('A1');
  const rich = cell.getRichTextValue();

  if (rich && rich.getLinkUrl()) {
    return String(cell.getDisplayValue() || '').toLowerCase().indexOf('dashboard') !== -1;
  }

  return String(cell.getDisplayValue() || '').toLowerCase().indexOf('dashboard') !== -1;
}

function shouldCheckConfigMatch_(classification) {
  return classification === 'PROJECT_TRACKER' ||
         classification === 'PROJECT_TRACKER_SHORT_NAME' ||
         classification === 'PROJECT_TASKS';
}

function getAuditRiskLevel_(sheetName, classification, hasDashboardLink, configMatch) {
  if (classification === 'PROJECT_TRACKER' || classification === 'PROJECT_TRACKER_SHORT_NAME') {
    if (!hasDashboardLink) {
      return {
        level: 'MEDIUM',
        notes: 'Tracker sheet is missing Back to Dashboard link.'
      };
    }

    if (configMatch === 'NO') {
      return {
        level: 'HIGH',
        notes: 'Tracker-like sheet does not appear to match Tracker Config.'
      };
    }

    return {
      level: 'LOW',
      notes: 'Tracker sheet detected and linked.'
    };
  }

  if (classification === 'PROJECT_TRACKER_SHORT_NAME') {
    return {
      level: 'MEDIUM',
      notes: 'Uses shortened Project Track naming; should standardize later.'
    };
  }

  if (classification === 'LEGACY_EMAIL_RFPS') {
    return {
      level: 'LEGACY',
      notes: 'Candidate for quarantine after confirmation.'
    };
  }

  if (classification === 'UNKNOWN_REVIEW') {
    return {
      level: 'REVIEW',
      notes: 'Could not classify automatically.'
    };
  }

  return {
    level: 'LOW',
    notes: ''
  };
}

function normalizeAuditProjectName_(value) {
  let name = String(value || '').trim();

  name = name
    .replace(/\s+–\s+Project Tracker$/i, '')
    .replace(/\s+-\s+Project Tracker$/i, '')
    .replace(/\s+Project Tracker$/i, '')
    .replace(/\s+Project Track$/i, '')
    .replace(/\s+Tasks$/i, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

  return name;
}

function findHeaderIndex_(headers, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const idx = headers.indexOf(candidates[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

function writePumaAudit_(ss, rows) {
  let sh = ss.getSheetByName(PUMA_AUDIT.OUTPUT_SHEET);
  if (!sh) sh = ss.insertSheet(PUMA_AUDIT.OUTPUT_SHEET);

  sh.clearContents();
  sh.clearFormats();

  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, rows[0].length);

  const header = sh.getRange(1, 1, 1, rows[0].length);
  header.setFontWeight('bold');
  header.setBackground('#d9ead3');

  const riskCol = 9;
  const lastRow = sh.getLastRow();

  if (lastRow > 1) {
    const riskRange = sh.getRange(2, riskCol, lastRow - 1, 1);
    const risks = riskRange.getValues();

    const backgrounds = risks.map(row => {
      const risk = String(row[0] || '').toUpperCase();

      if (risk === 'HIGH') return ['#f4cccc'];
      if (risk === 'MEDIUM') return ['#fff2cc'];
      if (risk === 'LEGACY') return ['#d9d2e9'];
      if (risk === 'REVIEW') return ['#cfe2f3'];
      return ['#d9ead3'];
    });

    riskRange.setBackgrounds(backgrounds);
  }
}