/**
 * PUMA SAFE SYNC
 *
 * Uses the soft match engine from PUMA_SoftMatch.js.
 *
 * Safe Sync does:
 * 1. Creates a backup of the active tracker
 * 2. Reads existing tracker rows
 * 3. Reads live quote rows from Tracker config
 * 4. Updates quote-driven fields only
 * 5. Preserves operational fields
 * 6. Zeroes dollar fields for quote rows removed from quote
 * 7. Appends brand-new quote lines at bottom
 *
 * Requires helper functions from PUMA_SoftMatch.js:
 * - readTrackerRowsForSoftMatch_
 * - readLiveQuoteRowsForTracker_
 * - softMatchRows_
 * - writeSoftMatchTestReport_
 * - generatePumaLineId_
 * - buildQuoteFingerprint_
 */

function safeSyncActiveTracker() {
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    'PUMA Safe Sync',
    'This will:\n\n' +
      '1. Create a backup of this tracker\n' +
      '2. Update quote-driven fields only\n' +
      '3. Preserve Status / PO / ESD / received dates\n' +
      '4. Zero dollar fields for rows removed from quote\n' +
      '5. Append new quote lines at the bottom\n\n' +
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

  results.forEach(result => {
    const existing = result.existing;

    // NEW_LINE has no existing tracker row yet.
    // Those are handled later by appendNewQuoteLines_().
    if (!existing || !existing.sourceRow) return;

    const rowNum = existing.sourceRow;

    // Always write identity and review flag.
    tracker
      .getRange(rowNum, col.pumaLineId)
      .setValue(result.pumaLineId || existing.pumaLineId || generatePumaLineId_());

    tracker
      .getRange(rowNum, col.reviewFlag)
      .setValue(result.reviewFlag || '');

    // Refresh quote-driven fields only.
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

      updated++;
    }

    // Keep removed quote-origin rows, but zero pipeline dollars.
    if (result.tier === 'REMOVED_OR_SUPERSEDED') {
      if (col.costPerUnit > 0) tracker.getRange(rowNum, col.costPerUnit).setValue(0);
      if (col.costWithMargin > 0) tracker.getRange(rowNum, col.costWithMargin).setValue(0);
      if (col.totalCost > 0) tracker.getRange(rowNum, col.totalCost).setValue(0);
      if (col.total > 0) tracker.getRange(rowNum, col.total).setValue(0);

      tracker
        .getRange(rowNum, col.reviewFlag)
        .setValue('REMOVED FROM QUOTE - dollar fields zeroed');

      zeroed++;
    }
  });

  const appended = appendNewQuoteLines_(tracker, results, col);

  writeSoftMatchTestReport_(ss, tracker.getName() + ' SAFE SYNC', results);

  ui.alert(
    'Safe Sync Complete\n\n' +
      `Tracker: ${tracker.getName()}\n` +
      `Updated existing rows: ${updated}\n` +
      `Zeroed removed quote rows: ${zeroed}\n` +
      `Appended new quote rows: ${appended}\n\n` +
      'Backup created.\nReport updated.'
  );
}

/**
 * Creates a backup copy of the active tracker before writing.
 */
function createTrackerBackup_(ss, tracker) {
  const name = tracker.getName();
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH.mm'
  );

  let backupName = `BACKUP - ${name} - ${timestamp}`;

  // Sheet names max at 100 chars.
  if (backupName.length > 100) {
    backupName = backupName.substring(0, 100);
  }

  // Avoid rare duplicate backup name collision.
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

/**
 * Maps tracker columns by header names.
 * Returns 1-based column numbers.
 */
function mapTrackerColumnsForSafeSync_(sheet) {
  const headerRow = 1;
  const headers = sheet
    .getRange(headerRow, 1, 1, sheet.getLastColumn())
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
    reviewFlag: find(['PUMA_REVIEW_FLAG'])
  };

  const ensured = ensureSafeSyncWriteColumns_(sheet, col);
  col.pumaLineId = ensured.pumaLineId;
  col.reviewFlag = ensured.reviewFlag;

  return col;
}

/**
 * Ensures PUMA_LINE_ID and PUMA_REVIEW_FLAG exist.
 */
function ensureSafeSyncWriteColumns_(sheet, col) {
  let lastCol = sheet.getLastColumn();
  const headerRow = 1;

  if (col.pumaLineId === -1) {
    lastCol++;
    sheet.getRange(headerRow, lastCol).setValue('PUMA_LINE_ID');
    sheet.getRange(headerRow, lastCol).setFontWeight('bold').setBackground('#cfe2f3');
    col.pumaLineId = lastCol;
  }

  if (col.reviewFlag === -1) {
    lastCol++;
    sheet.getRange(headerRow, lastCol).setValue('PUMA_REVIEW_FLAG');
    sheet.getRange(headerRow, lastCol).setFontWeight('bold').setBackground('#fff2cc');
    col.reviewFlag = lastCol;
  }

  try {
    sheet.hideColumns(col.pumaLineId);
  } catch (err) {
    // Ignore if already hidden or protected.
  }

  return {
    pumaLineId: col.pumaLineId,
    reviewFlag: col.reviewFlag
  };
}

/**
 * Validates required columns before writing.
 */
function validateSafeSyncColumns_(col) {
  const required = [
    'source',
    'type',
    'partNumber',
    'description',
    'manufacturer',
    'qty',
    'pumaLineId',
    'reviewFlag'
  ];

  const missing = required.filter(key => !col[key] || col[key] < 1);

  if (missing.length) {
    throw new Error('Safe Sync missing required tracker columns: ' + missing.join(', '));
  }
}

/**
 * Appends NEW_LINE incoming quote rows to the bottom.
 * Designed to avoid duplicate appends on repeat runs.
 */
function appendNewQuoteLines_(sheet, results, col) {
  let appendAt = sheet.getLastRow() + 1;
  let appended = 0;

  const existing = readTrackerRowsForSoftMatch_(sheet);
  const existingFingerprints = new Set(
    existing
      .map(r => buildQuoteFingerprint_(r))
      .filter(Boolean)
  );

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

    if (col.pumaLineId > 0) {
      sheet.getRange(rowNum, col.pumaLineId).setValue(result.pumaLineId || generatePumaLineId_());
    }

    if (col.reviewFlag > 0) {
      sheet.getRange(rowNum, col.reviewFlag).setValue('NEW FROM QUOTE - review');
    }

    existingFingerprints.add(fingerprint);
    appended++;
  });

  return appended;
}

/**
 * Label for Source column.
 */
function buildQuoteSourceLabel_(incoming) {
  if (incoming && incoming.quoteName) return incoming.quoteName;
  if (incoming && incoming.quoteTabName) return incoming.quoteTabName;
  return 'Quote Import';
}

/**
 * Header normalizer.
 */
function normalizeSafeSyncHeader_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^\w]+/g, '')
    .trim();
}