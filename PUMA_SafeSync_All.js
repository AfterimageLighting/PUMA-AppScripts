/**
 * PUMA SAFE SYNC ALL
 *
 * Runs Safe Sync across all enabled trackers in Tracker config.
 *
 * Requires:
 * - PUMA_SafeSync.js
 * - PUMA_SoftMatch.js
 */

function safeSyncALLConfiguredTrackers() {
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    'PUMA: Safe Sync ALL',
    'This will run Safe Sync on ALL enabled trackers in Tracker config.\n\n' +
      'Each tracker will get a backup before changes are made.\n\n' +
      'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName('Tracker config') || ss.getSheetByName('Tracker Config');

  if (!config) {
    ui.alert('Tracker config sheet not found.');
    return;
  }

  const values = config.getDataRange().getValues();
  if (values.length < 2) {
    ui.alert('Tracker config has no data rows.');
    return;
  }

  const headers = values[0].map(h => String(h || '').trim());

  const enableIdx = findSafeSyncAllHeader_(headers, ['Enable?', 'Enable', 'Enabled']);
  const trackerIdx = findSafeSyncAllHeader_(headers, ['Tracker Sheet Name', 'trackerName', 'Tracker Name', 'Tracker']);

  if (enableIdx === -1 || trackerIdx === -1) {
    ui.alert('Tracker config missing Enable? or Tracker Sheet Name column.');
    return;
  }

  const logRows = [];
  let success = 0;
  let skipped = 0;
  let failed = 0;

  for (let r = 1; r < values.length; r++) {
    const enabled = isSafeSyncAllEnabled_(values[r][enableIdx]);
    const trackerName = String(values[r][trackerIdx] || '').trim();

    if (!trackerName) continue;

    if (!enabled) {
      skipped++;
      logRows.push([trackerName, 'SKIPPED', 'Not enabled']);
      continue;
    }

    const tracker = ss.getSheetByName(trackerName);

    if (!tracker) {
      failed++;
      logRows.push([trackerName, 'ERROR', 'Tracker sheet not found']);
      continue;
    }

    try {
      ss.setActiveSheet(tracker);

      const result = safeSyncTrackerNoPrompt_(ss, tracker);

      success++;
      logRows.push([
        trackerName,
        'SUCCESS',
        `Updated: ${result.updated}, Zeroed: ${result.zeroed}, Flagged: ${result.flaggedOnly}, Appended: ${result.appended}`
      ]);

    } catch (err) {
      failed++;
      logRows.push([trackerName, 'ERROR', err.message || String(err)]);
    }
  }

  writeSafeSyncAllLog_(ss, logRows);

  ui.alert(
    'Safe Sync ALL Complete\n\n' +
      `Success: ${success}\n` +
      `Skipped: ${skipped}\n` +
      `Failed: ${failed}\n\n` +
      'See PUMA_SYNC_LOG for details.'
  );
}

/**
 * Runs the same safe sync logic as safeSyncActiveTracker(),
 * but without prompts/popups so ALL can run cleanly.
 */
function safeSyncTrackerNoPrompt_(ss, tracker) {
  if (!/project track/i.test(tracker.getName())) {
    throw new Error('Not a project tracker tab.');
  }

  createTrackerBackup_(ss, tracker);

  const existingRows = readTrackerRowsForSoftMatch_(tracker);
  const incomingRows = readLiveQuoteRowsForTracker_(ss, tracker.getName());

  if (!incomingRows.length) {
    return {
      updated: 0,
      zeroed: 0,
      flaggedOnly: 0,
      appended: 0,
      note: 'No incoming quote rows found'
    };
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

  return {
    updated,
    zeroed,
    flaggedOnly,
    appended
  };
}

function findSafeSyncAllHeader_(headers, candidates) {
  const normalizedHeaders = headers.map(h => normalizeSafeSyncAllHeader_(h));

  for (let i = 0; i < candidates.length; i++) {
    const idx = normalizedHeaders.indexOf(normalizeSafeSyncAllHeader_(candidates[i]));
    if (idx !== -1) return idx;
  }

  return -1;
}

function normalizeSafeSyncAllHeader_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^\w]+/g, '')
    .trim();
}

function isSafeSyncAllEnabled_(value) {
  if (value === true) return true;

  const text = String(value || '').trim().toLowerCase();

  return text === 'true' ||
         text === 'yes' ||
         text === 'y' ||
         text === '1' ||
         text === 'enabled';
}

function writeSafeSyncAllLog_(ss, rows) {
  let sh = ss.getSheetByName('PUMA_SYNC_LOG');
  if (!sh) sh = ss.insertSheet('PUMA_SYNC_LOG');

  const timestamp = new Date();

  const output = [
    ['Timestamp', 'Tracker', 'Result', 'Details'],
    ...rows.map(r => [timestamp, r[0], r[1], r[2]])
  ];

  sh.clearContents();
  sh.clearFormats();

  sh.getRange(1, 1, output.length, output[0].length).setValues(output);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, output[0].length);

  sh.getRange(1, 1, 1, output[0].length)
    .setFontWeight('bold')
    .setBackground('#d9ead3');
}