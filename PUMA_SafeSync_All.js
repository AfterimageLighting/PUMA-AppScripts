function safeSyncALLConfiguredTrackers() {
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    'PUMA: Safe Sync ALL',
    'This will run Safe Sync on ALL ENABLED trackers.\n\nContinue?',
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
  const headers = values[0].map(h => String(h || '').trim());

  const trackerIdx = headers.findIndex(h => /tracker/i.test(h));
  const enableIdx = headers.findIndex(h => /enable/i.test(h));

  if (trackerIdx === -1 || enableIdx === -1) {
    ui.alert('Tracker Config missing required columns.');
    return;
  }

  let results = [];
  let success = 0;
  let failed = 0;

  for (let i = 1; i < values.length; i++) {
    const trackerName = values[i][trackerIdx];
    const enabled = String(values[i][enableIdx]).toLowerCase();

    if (enabled !== 'true' && enabled !== 'yes') continue;
    if (!trackerName) continue;

    try {
      const sheet = ss.getSheetByName(trackerName);

      if (!sheet) {
        results.push([trackerName, 'NOT FOUND']);
        failed++;
        continue;
      }

      ss.setActiveSheet(sheet);

      safeSyncActiveTracker_INTERNAL_();

      results.push([trackerName, 'SUCCESS']);
      success++;

    } catch (err) {
      results.push([trackerName, 'ERROR: ' + err.message]);
      failed++;
    }
  }

  writeSafeSyncLog_(ss, results);

  ui.alert(
    `Safe Sync ALL Complete\n\n` +
    `Success: ${success}\n` +
    `Failed: ${failed}\n\n` +
    `See log: PUMA_SYNC_LOG`
  );
}

function writeSafeSyncLog_(ss, rows) {
  let sh = ss.getSheetByName('PUMA_SYNC_LOG');
  if (!sh) sh = ss.insertSheet('PUMA_SYNC_LOG');

  const timestamp = new Date();

  const output = [
    ['Timestamp', 'Tracker', 'Result'],
    ...rows.map(r => [timestamp, r[0], r[1]])
  ];

  sh.clearContents();
  sh.getRange(1, 1, output.length, output[0].length).setValues(output);
}