function safeSyncActiveTracker() {
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    'PUMA Safe Sync',
    'This will:\n\n1. Create a backup of this tracker\n2. Update quote-driven fields only\n3. Preserve all status/PO/ESD\n4. Zero removed quote dollar values\n\nContinue?',
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

  const results = softMatchRows_(existingRows, incomingRows);

  const col = mapTrackerColumnsForSafeSync_(tracker);

  let updated = 0;
  let zeroed = 0;

  results.forEach(result => {
    const existing = result.existing;
    if (!existing || !existing.sourceRow) return;

    const r = existing.sourceRow;

    // Always write PUMA ID + flag
    tracker.getRange(r, col.pumaLineId).setValue(result.pumaLineId || existing.pumaLineId || generatePumaLineId_());
    tracker.getRange(r, col.reviewFlag).setValue(result.reviewFlag || '');

    if (result.tier === 'EXACT_MATCH' || result.tier === 'SOFT_MATCH_TYPE') {
      const inc = result.incoming;

      tracker.getRange(r, col.type).setValue(inc.type);
      tracker.getRange(r, col.manufacturer).setValue(inc.manufacturer);
      tracker.getRange(r, col.partNumber).setValue(inc.partNumber);
      tracker.getRange(r, col.description).setValue(inc.description);
      tracker.getRange(r, col.qty).setValue(inc.qty);

      updated++;
    }

    if (result.tier === 'REMOVED_OR_SUPERSEDED') {
      tracker.getRange(r, col.costPerUnit).setValue(0);
      tracker.getRange(r, col.costWithMargin).setValue(0);
      tracker.getRange(r, col.totalCost).setValue(0);
      tracker.getRange(r, col.total).setValue(0);

      tracker.getRange(r, col.reviewFlag).setValue('REMOVED FROM QUOTE - dollar fields zeroed');

      zeroed++;
    }
  });

  writeSoftMatchTestReport_(ss, tracker.getName() + ' SAFE SYNC', results);

  ui.alert(
    `Safe Sync Complete\n\n` +
    `Updated rows: ${updated}\n` +
    `Zeroed rows: ${zeroed}\n\n` +
    `Backup created.\nReport updated.`
  );
}

function createTrackerBackup_(ss, tracker) {
  const name = tracker.getName();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  const copy = tracker.copyTo(ss);
  copy.setName(`BACKUP - ${name} - ${timestamp}`);

  ss.setActiveSheet(tracker);
}

function mapTrackerColumnsForSafeSync_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const find = (names) => {
    const norm = headers.map(h => String(h).toLowerCase());
    for (let n of names) {
      const i = norm.indexOf(n.toLowerCase());
      if (i !== -1) return i + 1;
    }
    return -1;
  };

  return {
    type: find(['Type']),
    manufacturer: find(['Manufacturer']),
    partNumber: find(['Part Number']),
    description: find(['Description']),
    qty: find(['Quantity']),
    costPerUnit: find(['Cost Per Unit']),
    costWithMargin: find(['Cost Per Unit with Margin']),
    totalCost: find(['Total Cost']),
    total: find(['Total']),
    pumaLineId: find(['PUMA_LINE_ID']),
    reviewFlag: find(['PUMA_REVIEW_FLAG'])
  };
}
