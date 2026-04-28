/**
 * Builds PUMA_AUDIT_ACTIONS from the current PUMA_AUDIT sheet.
 * This does NOT change project tabs. It only creates a cleanup planning sheet.
 */
function buildPumaAuditActions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const audit = ss.getSheetByName('PUMA_AUDIT');

  if (!audit) {
    throw new Error('PUMA_AUDIT sheet not found. Run runPumaAudit() first.');
  }

  const values = audit.getDataRange().getValues();
  if (values.length < 2) {
    throw new Error('PUMA_AUDIT has no data.');
  }

  const headers = values[0].map(String);
  const idx = {
    sheetName: headers.indexOf('Sheet Name'),
    classification: headers.indexOf('Classification'),
    configMatch: headers.indexOf('Config Match'),
    expectedPair: headers.indexOf('Expected Paired Sheet'),
    pairExists: headers.indexOf('Pair Exists'),
    risk: headers.indexOf('Risk Level'),
    notes: headers.indexOf('Notes')
  };

  const required = Object.keys(idx).filter(k => idx[k] === -1);
  if (required.length) {
    throw new Error('Missing audit columns: ' + required.join(', '));
  }

  const out = [[
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
    'Done?'
  ]];

  values.slice(1).forEach(row => {
    const risk = String(row[idx.risk] || '').toUpperCase();
    if (risk !== 'HIGH' && risk !== 'LEGACY' && risk !== 'REVIEW') return;

    const sheetName = String(row[idx.sheetName] || '');
    const classification = String(row[idx.classification] || '');
    const configMatch = String(row[idx.configMatch] || '');
    const expectedPair = String(row[idx.expectedPair] || '');
    const pairExists = String(row[idx.pairExists] || '');
    const notes = String(row[idx.notes] || '');

    const rec = recommendAuditAction_({
      sheetName,
      classification,
      configMatch,
      expectedPair,
      pairExists,
      risk,
      notes
    });

    out.push([
      rec.priority,
      rec.action,
      sheetName,
      classification,
      configMatch,
      expectedPair,
      pairExists,
      risk,
      rec.reason,
      '',
      false
    ]);
  });

  writePumaAuditActions_(ss, out);
  SpreadsheetApp.getUi().alert(`PUMA_AUDIT_ACTIONS created.\nItems: ${out.length - 1}`);
}

function recommendAuditAction_(item) {
  if (item.risk === 'LEGACY') {
    return {
      priority: 4,
      action: 'QUARANTINE LEGACY',
      reason: 'Legacy/RFPS/email workflow candidate. Do not delete yet.'
    };
  }

  if (item.classification === 'PROJECT_TRACKER' || item.classification === 'PROJECT_TRACKER_SHORT_NAME') {
    if (item.pairExists === 'NO' && item.configMatch === 'YES') {
      return {
        priority: 1,
        action: 'CREATE MISSING TASK TAB',
        reason: 'Active configured tracker is missing its task companion.'
      };
    }

    if (item.configMatch === 'NO' && item.pairExists === 'YES') {
      return {
        priority: 3,
        action: 'ARCHIVE HISTORICAL / REMOVE FROM ACTIVE WORKBOOK',
        reason: 'Tracker is paired but not configured; likely completed or historical.'
      };
    }

    if (item.configMatch === 'NO' && item.pairExists === 'NO') {
      return {
        priority: 2,
        action: 'REVIEW OR ARCHIVE',
        reason: 'Tracker is not configured and has no matching task tab.'
      };
    }
  }

  if (item.classification === 'PROJECT_TASKS') {
    if (item.pairExists === 'NO') {
      return {
        priority: 1,
        action: 'FIX NAME OR CREATE MATCHING TRACKER',
        reason: 'Task tab has no matching tracker. Likely typo or orphan.'
      };
    }
  }

  if (looksLikeTypoPair_(item.sheetName, item.expectedPair)) {
    return {
      priority: 1,
      action: 'FIX NAME TYPO',
      reason: 'Expected pair is very close to sheet name; likely spelling mismatch.'
    };
  }

  return {
    priority: 5,
    action: 'REVIEW',
    reason: item.notes || 'Needs manual review.'
  };
}

function looksLikeTypoPair_(a, b) {
  if (!a || !b) return false;

  const aa = normalizeAuditActionName_(a);
  const bb = normalizeAuditActionName_(b);

  if (!aa || !bb) return false;
  if (aa === bb) return false;

  return levenshteinDistance_(aa, bb) <= 3;
}

function normalizeAuditActionName_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/project tracker/g, '')
    .replace(/project track/g, '')
    .replace(/tasks/g, '')
    .replace(/[^\w\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function levenshteinDistance_(a, b) {
  const matrix = [];

  for (let i = 0; i <= b.length; i++) matrix[i] = [i];
  for (let j = 0; j <= a.length; j++) matrix[0][j] = j;

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }

  return matrix[b.length][a.length];
}

function writePumaAuditActions_(ss, rows) {
  let sh = ss.getSheetByName('PUMA_AUDIT_ACTIONS');
  if (!sh) sh = ss.insertSheet('PUMA_AUDIT_ACTIONS');

  sh.clearContents();
  sh.clearFormats();

  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, rows[0].length);

  const header = sh.getRange(1, 1, 1, rows[0].length);
  header.setFontWeight('bold');
  header.setBackground('#d9ead3');

  if (rows.length > 1) {
    sh.getRange(2, 1, rows.length - 1, rows[0].length)
      .sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);

    sh.getRange(2, 11, rows.length - 1, 1).insertCheckboxes();

    const actionRange = sh.getRange(2, 2, rows.length - 1, 1);
    const actions = actionRange.getValues();

    const backgrounds = actions.map(row => {
      const action = String(row[0] || '').toUpperCase();

      if (action.indexOf('CREATE') !== -1) return ['#f4cccc'];
      if (action.indexOf('FIX') !== -1) return ['#fff2cc'];
      if (action.indexOf('ARCHIVE') !== -1) return ['#d9d2e9'];
      if (action.indexOf('QUARANTINE') !== -1) return ['#d9d2e9'];
      return ['#cfe2f3'];
    });

    actionRange.setBackgrounds(backgrounds);
  }
}