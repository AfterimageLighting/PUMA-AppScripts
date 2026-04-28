/**
 * PUMA AUDIT v2
 * Builds a workbook inventory and risk report.
 *
 * Output sheet:
 *   PUMA_AUDIT
 */

const PUMA_AUDIT = {
  OUTPUT_SHEET: 'PUMA_AUDIT',
  DASHBOARD: 'Dashboard',
  TRACKER_CONFIG_NAMES: ['Tracker Config', 'Tracker config'],
  OPEN_PROJECTS: 'Open Projects',
  HEADER_ROW: 1
};

function runPumaAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheetNames = ss.getSheets().map(sh => sh.getName());
  const allSheetNameSet = new Set(allSheetNames.map(n => normalizeAuditLoose_(n)));

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
    'Expected Paired Sheet',
    'Pair Exists',
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

    const pairInfo = getExpectedPairInfo_(name, classification, allSheetNameSet);
    const risk = getAuditRiskLevel_(
      name,
      classification,
      hasDashboardLink,
      shouldCheckConfigMatch_(classification) ? configMatch : '',
      pairInfo.exists
    );

    auditRows.push([
      timestamp,
      name,
      classification,
      sheet.getMaxRows(),
      sheet.getMaxColumns(),
      sheet.getFrozenRows(),
      hasDashboardLink ? 'YES' : 'NO',
      shouldCheckConfigMatch_(classification) ? configMatch : '',
      pairInfo.expectedName,
      pairInfo.required ? (pairInfo.exists ? 'YES' : 'NO') : '',
      risk.level,
      risk.notes
    ]);
  });

  writePumaAudit_(ss, auditRows);
  SpreadsheetApp.getUi().alert(`PUMA audit complete.\nSheets scanned: ${auditRows.length - 1}`);
}

function getAuditTrackerConfigProjects_(ss) {
  const set = new Set();
  const sh = getSheetByAnyName_(ss, PUMA_AUDIT.TRACKER_CONFIG_NAMES);
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
    'trackername',
    'tracker name'
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
  if (PUMA_AUDIT.TRACKER_CONFIG_NAMES.some(n => normalizeAuditLoose_(n) === normalizeAuditLoose_(name))) return 'CORE_CONFIG';
  if (name === PUMA_AUDIT.OPEN_PROJECTS) return 'CORE_PROJECT_LIST';

  if (name === '- Project Tracker' || name === ' - Project Tracker') return 'TEMPLATE_TRACKER';
  if (name === '- Tasks' || name === ' - Tasks') return 'TEMPLATE_TASKS';

  if (/^pipeline$/i.test(name)) return 'CORE_DASHBOARD_SUPPORT';
  if (/^<\s*85%$/i.test(name)) return 'CORE_PIPELINE_BUCKET';
  if (/^85%$/i.test(name)) return 'CORE_PIPELINE_BUCKET';
  if (/^100%$/i.test(name)) return 'CORE_PIPELINE_BUCKET';
  if (/^new projects$/i.test(name)) return 'PROJECT_INTAKE';
  if (/^po unmatched$/i.test(name)) return 'PO_PIPELINE_EXCEPTION';

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

function getExpectedPairInfo_(sheetName, classification, allSheetNameSet) {
  const name = String(sheetName || '').trim();

  if (classification === 'PROJECT_TRACKER' || classification === 'PROJECT_TRACKER_SHORT_NAME') {
    const base = normalizeAuditProjectDisplayName_(name);

    const expectedTaskA = `${base} - Tasks`;
    const expectedTaskB = `${base} – Tasks`;
    const expectedTaskC = `${base} Tasks`;

    const exists =
      allSheetNameSet.has(normalizeAuditLoose_(expectedTaskA)) ||
      allSheetNameSet.has(normalizeAuditLoose_(expectedTaskB)) ||
      allSheetNameSet.has(normalizeAuditLoose_(expectedTaskC));

    return {
      required: true,
      expectedName: expectedTaskA,
      exists
    };
  }

  if (classification === 'PROJECT_TASKS') {
    const base = normalizeAuditProjectDisplayName_(name);

    const expectedTrackerA = `${base} - Project Tracker`;
    const expectedTrackerB = `${base} – Project Tracker`;
    const expectedTrackerC = `${base} Project Tracker`;
    const expectedTrackerD = `${base} Project Track`;

    const exists =
      allSheetNameSet.has(normalizeAuditLoose_(expectedTrackerA)) ||
      allSheetNameSet.has(normalizeAuditLoose_(expectedTrackerB)) ||
      allSheetNameSet.has(normalizeAuditLoose_(expectedTrackerC)) ||
      allSheetNameSet.has(normalizeAuditLoose_(expectedTrackerD));

    return {
      required: true,
      expectedName: expectedTrackerA,
      exists
    };
  }

  return {
    required: false,
    expectedName: '',
    exists: true
  };
}

function getAuditRiskLevel_(sheetName, classification, hasDashboardLink, configMatch, pairExists) {
  if (classification === 'TEMPLATE_TRACKER' || classification === 'TEMPLATE_TASKS') {
    return {
      level: 'LOW',
      notes: 'Intentional hanging template sheet.'
    };
  }

  if (classification === 'PROJECT_TRACKER' || classification === 'PROJECT_TRACKER_SHORT_NAME') {
    if (pairExists === false) {
      return {
        level: 'HIGH',
        notes: 'Tracker sheet is missing its paired Tasks sheet.'
      };
    }

    if (!hasDashboardLink) {
      return {
        level: 'MEDIUM',
        notes: 'Tracker sheet is missing Back to Dashboard link.'
      };
    }

    if (configMatch === 'NO') {
      return {
        level: 'HIGH',
        notes: 'Tracker-like sheet does not appear to match Tracker Config. May be historical/orphaned.'
      };
    }

    if (classification === 'PROJECT_TRACKER_SHORT_NAME') {
      return {
        level: 'MEDIUM',
        notes: 'Uses shortened Project Track naming; should standardize later.'
      };
    }

    return {
      level: 'LOW',
      notes: 'Tracker sheet detected, linked, and paired.'
    };
  }

  if (classification === 'PROJECT_TASKS') {
    if (pairExists === false) {
      return {
        level: 'HIGH',
        notes: 'Tasks sheet is missing its paired Project Tracker sheet.'
      };
    }

    return {
      level: 'LOW',
      notes: 'Tasks sheet detected and paired.'
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

function normalizeAuditProjectName_(value) {
  return normalizeAuditProjectDisplayName_(value).toLowerCase();
}

function normalizeAuditProjectDisplayName_(value) {
  let name = String(value || '').trim();

  name = name
    .replace(/\s*[--]\s*Project Tracker$/i, '')
    .replace(/\s+Project Tracker$/i, '')
    .replace(/\s+Project Track$/i, '')
    .replace(/\s*[--]\s*Tasks$/i, '')
    .replace(/\s+Tasks$/i, '')
    .replace(/\s+/g, ' ')
    .trim();

  return name;
}

function normalizeAuditLoose_(value) {
  return String(value || '')
    .replace(/[–—]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function getSheetByAnyName_(ss, names) {
  for (let i = 0; i < names.length; i++) {
    const sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }

  const looseMap = {};
  ss.getSheets().forEach(sh => {
    looseMap[normalizeAuditLoose_(sh.getName())] = sh;
  });

  for (let i = 0; i < names.length; i++) {
    const match = looseMap[normalizeAuditLoose_(names[i])];
    if (match) return match;
  }

  return null;
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

  const riskCol = 11;
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