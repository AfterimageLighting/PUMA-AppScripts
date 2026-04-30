/**
 * PUMA SAFE SYNC v2
 *
 * Adds:
 * - PUMA_SOURCE_TYPE
 * - Safer zeroing rules
 * - Review flag color coding
 *
 * Source types:
 * QUOTE   = came from quote sync
 * PO_ONLY = exists in tracker but not quote; likely PO/import field add
 * MANUAL  = manually added/internal row
 * UNKNOWN = legacy row not yet classified
 */

function safeSyncActiveTracker() {
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    'PUMA Safe Sync v2',
    'This will:\n\n' +
      '1. Create a backup of this tracker\n' +
      '2. Update quote-driven fields only\n' +
      '3. Preserve Status / PO / ESD / received dates\n' +
      '4. Add PUMA_SOURCE_TYPE\n' +
      '5. Zero dollar fields ONLY for quote-origin rows removed from quote\n' +
      '6. Append new quote lines at the bottom\n' +
      '7. Color-code review flags\n\n' +
      'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tracker = ss.getActiveSheet();

  if (!/project track/i.test(tracker.getName())) {
    ui.alert('Open a Project Tracker tab first.');
    return;
  }

  createTrackerBackup_(ss, tracker);

  const existingRows = readTrackerRowsForSoftMatch_(tracker);
  const incomingRows = readLiveQuoteRowsForTracker_(ss, tracker.getName());

  if (!incomingRows.length) {
    ui.alert('No live quote rows found for this tracker from Tracker config.');
    return;
  }

  const results = softMatchRows_(existingRows, incomingRows);
  const col = mapTrackerColumnsForSafeSync_(tracker);

  validateSafeSyncColumns_(col);

  let updated = 0;
  let zeroed = 0;
  let flaggedOnly = 0;

  results.forEach(result => {
    const existing = result.existing;

    if (!existing || !existing.sourceRow) return;

    const rowNum = existing.sourceRow;
    const currentSourceType = getSourceTypeForRow_(tracker, rowNum, col);

    tracker.getRange(rowNum, col.pumaLineId)
      .setValue(result.pumaLineId || existing.pumaLineId || generatePumaLineId_());

    if (result.tier === 'EXACT_MATCH' || result.tier === 'SOFT_MATCH_TYPE') {
      const inc = result.incoming || {};

      tracker.getRange(rowNum, col.source).setValue(buildQuoteSourceLabel_(inc));
      tracker.getRange(rowNum, col.type).setValue(inc.type || '');
      tracker.getRange(rowNum, col.partNumber).setValue(inc.partNumber || '');
      tracker.getRange(rowNum, col.description).setValue(inc.description || '');
      tracker.getRange(rowNum, col.manufacturer).setValue(inc.manufacturer || '');
      tracker.getRange(rowNum, col.qty).setValue(inc.qty || '');

      if (col.costPerUnit > 0) tracker.getRange(rowNum, col.costPerUnit).setValue(inc.costPerUnit || 0);
      if (col.costWithMargin > 0) tracker.getRange(rowNum, col.costWithMargin).setValue(inc.costWithMargin || 0);
      if (col.totalCost > 0) tracker.getRange(rowNum, col.totalCost).setValue(inc.totalCost || 0);
      if (col.total > 0) tracker.getRange(rowNum, col.total).setValue(inc.total || 0);

      tracker.getRange(rowNum, col.sourceType).setValue('QUOTE');
      tracker.getRange(rowNum, col.reviewFlag).setValue(result.reviewFlag || '');

      updated++;
      return;
    }

    if (result.tier === 'REMOVED_OR_SUPERSEDED') {
      const sourceType = currentSourceType || inferSourceTypeFromRow_(tracker, rowNum, col);

      if (sourceType === 'QUOTE') {
        zeroDollarFields_(tracker, rowNum, col);
        tracker.getRange(rowNum, col.reviewFlag).setValue('REMOVED FROM QUOTE - dollar fields zeroed');
        tracker.getRange(rowNum, col.sourceType).setValue('QUOTE');
        zeroed++;
      } else if (sourceType === 'PO_ONLY') {
        tracker.getRange(rowNum, col.reviewFlag).setValue('PO_ONLY - not in quote; kept without zeroing');
        tracker.getRange(rowNum, col.sourceType).setValue('PO_ONLY');
        flaggedOnly++;
      } else if (sourceType === 'MANUAL') {
        tracker.getRange(rowNum, col.reviewFlag).setValue('MANUAL - not in quote; kept without zeroing');
        tracker.getRange(rowNum, col.sourceType).setValue('MANUAL');
        flaggedOnly++;
      } else {
        tracker.getRange(rowNum, col.reviewFlag).setValue('UNKNOWN SOURCE - review before zeroing');
        tracker.getRange(rowNum, col.sourceType).setValue('UNKNOWN');
        flaggedOnly++;
      }
    }
  });

  const appended = appendNewQuoteLines_(tracker, results, col);

  applyReviewFlagColors_(tracker, col);

  writeSoftMatchTestReport_(ss, tracker.getName() + ' SAFE SYNC', results);

  ui.alert(
    'Safe Sync v2 Complete\n\n' +
      `Tracker: ${tracker.getName()}\n` +
      `Updated quote rows: ${updated}\n` +
      `Zeroed removed quote rows: ${zeroed}\n` +
      `Flagged non-quote rows: ${flaggedOnly}\n` +
      `Appended new quote rows: ${appended}\n\n` +
      'Backup created.\nReport updated.'
  );
}

function createTrackerBackup_(ss, tracker) {
  const name = tracker.getName();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH.mm');

  let backupName = `BACKUP - ${name} - ${timestamp}`;
  if (backupName.length > 100) backupName = backupName.substring(0, 100);

  let finalName = backupName;
  let i = 2;

  while (ss.getSheetByName(finalName)) {
    finalName = `${backupName.substring(0, 95)} ${i}`;
    i++;
  }

  const copy = tracker.copyTo(ss);
  copy.setName(finalName);
  ss.setActiveSheet(tracker);
}

function mapTrackerColumnsForSafeSync_(sheet) {
  const headerRow = 1;
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h || '').trim());

  const find = names => {
    const normalizedHeaders = headers.map(h => normalizeSafeSyncHeader_(h));

    for (let n of names) {
      const idx = normalizedHeaders.indexOf(normalizeSafeSyncHeader_(n));
      if (idx !== -1) return idx + 1;
    }

    return -1;
  };

  const col = {
    source: find(['Source']),
    type: find(['Type']),
    partNumber: find(['Part Number']),
    description: find(['Description']),
    manufacturer: find(['Manufacturer']),
    qty: find(['Quantity', 'Qty']),
    status: find(['Status']),
    poNumber: find(['PO Number', 'PO #', 'PO']),
    esd: find(['Estimated Ship Date (ESD)', 'Estimated Ship Date', 'ESD']),
    dateReceived: find(['Date Received']),
    dateScheduled: find(['Date Scheduled']),
    dateDelivered: find(['Date Delivered']),
    costPerUnit: find(['Cost Per Unit']),
    costWithMargin: find(['Cost Per Unit with Margin']),
    totalCost: find(['Total Cost']),
    total: find(['Total']),
    pumaLineId: find(['PUMA_LINE_ID']),
    reviewFlag: find(['PUMA_REVIEW_FLAG']),
    sourceType: find(['PUMA_SOURCE_TYPE'])
  };

  const ensured = ensureSafeSyncWriteColumns_(sheet, col);
  col.pumaLineId = ensured.pumaLineId;
  col.reviewFlag = ensured.reviewFlag;
  col.sourceType = ensured.sourceType;

  return col;
}

function ensureSafeSyncWriteColumns_(sheet, col) {
  let lastCol = sheet.getLastColumn();
  const headerRow = 1;

  if (col.pumaLineId === -1) {
    lastCol++;
    sheet.getRange(headerRow, lastCol).setValue('PUMA_LINE_ID')
      .setFontWeight('bold')
      .setBackground('#cfe2f3');
    col.pumaLineId = lastCol;
  }

  if (col.reviewFlag === -1) {
    lastCol++;
    sheet.getRange(headerRow, lastCol).setValue('PUMA_REVIEW_FLAG')
      .setFontWeight('bold')
      .setBackground('#fff2cc');
    col.reviewFlag = lastCol;
  }

  if (col.sourceType === -1) {
    lastCol++;
    sheet.getRange(headerRow, lastCol).setValue('PUMA_SOURCE_TYPE')
      .setFontWeight('bold')
      .setBackground('#d9ead3');
    col.sourceType = lastCol;
  }

  try {
    sheet.hideColumns(col.pumaLineId);
  } catch (err) {}

  return {
    pumaLineId: col.pumaLineId,
    reviewFlag: col.reviewFlag,
    sourceType: col.sourceType
  };
}

function validateSafeSyncColumns_(col) {
  const required = [
    'source',
    'type',
    'partNumber',
    'description',
    'manufacturer',
    'qty',
    'pumaLineId',
    'reviewFlag',
    'sourceType'
  ];

  const missing = required.filter(key => !col[key] || col[key] < 1);
  if (missing.length) {
    throw new Error('Safe Sync missing required tracker columns: ' + missing.join(', '));
  }
}

function appendNewQuoteLines_(sheet, results, col) {
  let appendAt = sheet.getLastRow() + 1;
  let appended = 0;

  const existing = readTrackerRowsForSoftMatch_(sheet);
  const existingFingerprints = new Set(existing.map(r => buildQuoteFingerprint_(r)).filter(Boolean));

  results.forEach(result => {
    if (result.tier !== 'NEW_LINE') return;

    const inc = result.incoming || {};
    const fingerprint = buildQuoteFingerprint_(inc);

    if (fingerprint && existingFingerprints.has(fingerprint)) return;

    const rowNum = appendAt++;

    if (col.source > 0) sheet.getRange(rowNum, col.source).setValue(buildQuoteSourceLabel_(inc));
    if (col.type > 0) sheet.getRange(rowNum, col.type).setValue(inc.type || '');
    if (col.partNumber > 0) sheet.getRange(rowNum, col.partNumber).setValue(inc.partNumber || '');
    if (col.description > 0) sheet.getRange(rowNum, col.description).setValue(inc.description || '');
    if (col.manufacturer > 0) sheet.getRange(rowNum, col.manufacturer).setValue(inc.manufacturer || '');
    if (col.qty > 0) sheet.getRange(rowNum, col.qty).setValue(inc.qty || '');

    if (col.costPerUnit > 0) sheet.getRange(rowNum, col.costPerUnit).setValue(inc.costPerUnit || 0);
    if (col.costWithMargin > 0) sheet.getRange(rowNum, col.costWithMargin).setValue(inc.costWithMargin || 0);
    if (col.totalCost > 0) sheet.getRange(rowNum, col.totalCost).setValue(inc.totalCost || 0);
    if (col.total > 0) sheet.getRange(rowNum, col.total).setValue(inc.total || 0);

    sheet.getRange(rowNum, col.pumaLineId).setValue(result.pumaLineId || generatePumaLineId_());
    sheet.getRange(rowNum, col.reviewFlag).setValue('NEW FROM QUOTE - review');
    sheet.getRange(rowNum, col.sourceType).setValue('QUOTE');

    existingFingerprints.add(fingerprint);
    appended++;
  });

  return appended;
}

function getSourceTypeForRow_(sheet, rowNum, col) {
  if (!col.sourceType || col.sourceType < 1) return '';
  return String(sheet.getRange(rowNum, col.sourceType).getValue() || '').trim().toUpperCase();
}

function inferSourceTypeFromRow_(sheet, rowNum, col) {
  const source = col.source > 0 ? String(sheet.getRange(rowNum, col.source).getDisplayValue() || '') : '';
  const po = col.poNumber > 0 ? String(sheet.getRange(rowNum, col.poNumber).getDisplayValue() || '') : '';
  const type = col.type > 0 ? String(sheet.getRange(rowNum, col.type).getDisplayValue() || '') : '';

  if (/quote/i.test(source)) return 'QUOTE';
  if (po && !/quote/i.test(source)) return 'PO_ONLY';
  if (!source && !po && type) return 'MANUAL';

  return 'UNKNOWN';
}

function zeroDollarFields_(sheet, rowNum, col) {
  if (col.costPerUnit > 0) sheet.getRange(rowNum, col.costPerUnit).setValue(0);
  if (col.costWithMargin > 0) sheet.getRange(rowNum, col.costWithMargin).setValue(0);
  if (col.totalCost > 0) sheet.getRange(rowNum, col.totalCost).setValue(0);
  if (col.total > 0) sheet.getRange(rowNum, col.total).setValue(0);
}

function applyReviewFlagColors_(sheet, col) {
  if (!col.reviewFlag || col.reviewFlag < 1) return;

  const startRow = 4;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const range = sheet.getRange(startRow, col.reviewFlag, lastRow - startRow + 1, 1);
  const values = range.getDisplayValues();

  const backgrounds = values.map(row => {
    const flag = String(row[0] || '').toUpperCase();

    if (!flag) return ['#ffffff'];
    if (flag.indexOf('NEW FROM QUOTE') !== -1) return ['#cfe2f3'];
    if (flag.indexOf('REMOVED FROM QUOTE') !== -1) return ['#d9d2e9'];
    if (flag.indexOf('CHANGED') !== -1 || flag.indexOf('VERIFY') !== -1) return ['#fff2cc'];
    if (flag.indexOf('PO_ONLY') !== -1) return ['#fce5cd'];
    if (flag.indexOf('MANUAL') !== -1) return ['#eadcf8'];
    if (flag.indexOf('UNKNOWN') !== -1 || flag.indexOf('REVIEW') !== -1) return ['#f4cccc'];

    return ['#ffffff'];
  });

  range.setBackgrounds(backgrounds);
}

function buildQuoteSourceLabel_(incoming) {
  if (incoming && incoming.quoteName) return incoming.quoteName;
  if (incoming && incoming.quoteTabName) return incoming.quoteTabName;
  return 'Quote Import';
}

function normalizeSafeSyncHeader_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^\w]+/g, '')
    .trim();
}