/************************************************************
 * Dashboard.gs
 *
 * Source of truth:
 * - Tracker config
 *
 * Layout:
 * - Keeps old Dashboard style
 * - One parent project row
 * - Expandable quote child rows underneath
 ************************************************************/

function refreshDashboardProjectTrackers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = getOrCreateDashboardSheet_(ss);
  const configSheet = ss.getSheetByName('Tracker config');

  if (!configSheet) throw new Error('Sheet "Tracker config" not found.');

  clearDashboardForRebuild_(dashboard);
  buildDashboardHeaderV4_(dashboard);

  const projectMap = readDashboardProjectsFromTrackerConfig_(configSheet);

  const sheetMap = {};
  ss.getSheets().forEach(sheet => {
    sheetMap[sheet.getName()] = sheet;
  });

  Object.keys(projectMap).forEach(projectName => {
    const p = projectMap[projectName];

    const trackerSheet = sheetMap[p.trackerName] || null;
    const tasksSheet = sheetMap[`${projectName} - Tasks`] || null;

    p.trackerExists = !!trackerSheet;
    p.tasksExists = !!tasksSheet;
    p.openTasks = countOpenTasksOnly_(tasksSheet);
    p.hyperlink = trackerSheet ? `${ss.getUrl()}#gid=${trackerSheet.getSheetId()}` : '';

    const trackerMetrics = readTrackerMetricsBySource_(trackerSheet);

    copyDashboardMetricsToProject_(p, trackerMetrics.projectTotals);

    p.quotes.forEach(q => {
      const metrics = trackerMetrics.bySource[q.quoteName] || createEmptyDashboardMetrics_();
      copyDashboardMetricsToQuote_(q, metrics);
    });

    p.projectClosing = getHighestProjectClosingPct_(p.quotes);
  });

  const projects = Object.keys(projectMap)
    .map(k => projectMap[k])
    .sort((a, b) => a.projectName.localeCompare(b.projectName));

  const values = [];
  const richProjectCells = [];
  const closingBackgrounds = [];
  const snapshotBackgrounds = [];
  const rowTypes = [];
  const remainingFormulaRows = [];

  projects.forEach(p => {
    values.push([
      p.projectName,
      '',
      p.projectClosing,
      p.openTasks,
      p.totalLineItems,
      p.totalSold,
      '',
      '',
      '',
      '',
      '',
      '',
      p.stageSnapshot
    ]);

    richProjectCells.push([buildProjectRichText_(p.projectName, p.hyperlink)]);
    closingBackgrounds.push([p.trackerExists ? '#d9ead3' : '#f4cccc']);
    snapshotBackgrounds.push([p.trackerExists ? '#d9ead3' : '#f4cccc']);
    rowTypes.push('project');

    p.quotes.forEach(q => {
      const closing = String(q.closingPct || '').trim();
      const quoteTotal = toNumber_(q.quoteTotal);

      values.push([
        '',
        '↳ ' + q.quoteName,
        closing,
        '',
        q.totalLineItems,
        q.totalSold,
        quoteTotal,
        closing === '<85%' ? quoteTotal : '',
        closing === '>85%' ? quoteTotal : '',
        closing === '100%' ? quoteTotal : '',
        '',
        '',
        q.enabled ? 'Enabled' : 'Disabled'
      ]);

      const sheetRow = values.length + 1;
      if (closing === '100%') remainingFormulaRows.push(sheetRow);

      richProjectCells.push([SpreadsheetApp.newRichTextValue().setText('').build()]);
      closingBackgrounds.push(['#ffffff']);
      snapshotBackgrounds.push([q.enabled ? '#d9ead3' : '#fff2cc']);
      rowTypes.push('quote');
    });
  });

  const startRow = 2;
  const numRows = values.length;
  const totalCols = 13;

  if (numRows > 0) {
    dashboard.getRange(startRow, 1, numRows, totalCols).setValues(values);
    dashboard.getRange(startRow, 1, numRows, 1).setRichTextValues(richProjectCells);
    dashboard.getRange(startRow, 3, numRows, 1).setBackgrounds(closingBackgrounds);
    dashboard.getRange(startRow, 13, numRows, 1).setBackgrounds(snapshotBackgrounds);

    remainingFormulaRows.forEach(row => {
      dashboard.getRange(row, 12)
        .setFormula(`=IF(J${row}="","",MAX(0,J${row}-N(K${row})))`);
    });
  }

  const lastDataRow = numRows > 0 ? startRow + numRows - 1 : 1;
  const totalsRow = lastDataRow + 1;

  dashboard.getRange(totalsRow, 1, 1, totalCols).setValues([[
    'PIPELINE TOTALS',
    new Date(),
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    ''
  ]]);

  dashboard.getRange(totalsRow, 2).setNumberFormat('MM/dd/yyyy');

  if (numRows > 0) {
    dashboard.getRange(totalsRow, 8).setFormula(`=SUM(H${startRow}:H${lastDataRow})`);
    dashboard.getRange(totalsRow, 9).setFormula(`=SUM(I${startRow}:I${lastDataRow})`);
    dashboard.getRange(totalsRow, 10).setFormula(`=SUM(J${startRow}:J${lastDataRow})`);
    dashboard.getRange(totalsRow, 11).setFormula(`=SUM(K${startRow}:K${lastDataRow})`);
    dashboard.getRange(totalsRow, 12).setFormula(`=SUM(L${startRow}:L${lastDataRow})`);
  } else {
    dashboard.getRange(totalsRow, 8, 1, 5).setValues([[0, 0, 0, 0, 0]]);
  }

  formatDashboardSheetV4_(dashboard, numRows, totalsRow);
  applyProjectQuoteRowStylingV4_(dashboard, rowTypes, startRow);
  addRowGroupsForQuotesV4_(dashboard, rowTypes, startRow);

  dashboard.getRange(totalsRow, 1, 1, totalCols)
    .setFontWeight('bold')
    .setBackground('#d9d9d9')
    .setBorder(true, true, true, true, true, true, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  dashboard.setFrozenRows(1);

  colorDashboardClosingColumn_();
}

function readDashboardProjectsFromTrackerConfig_(configSheet) {
  const lastRow = configSheet.getLastRow();
  const projectMap = {};

  if (lastRow < 2) return projectMap;

  const values = configSheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();

  values.forEach(row => {
    const enabled = row[0] === true || String(row[0]).toLowerCase() === 'true';
    const projectName = String(row[1] || '').trim();
    const trackerName = String(row[2] || '').trim() || `${projectName} - Project Tracker`;
    const dateUpdated = row[3];
    const quoteName = String(row[4] || '').trim();
    const quoteId = String(row[5] || '').trim();
    const closingPct = String(row[6] || '').trim();   // G
    const closingNote = String(row[7] || '').trim();  // H

    if (!projectName) return;

    if (!projectMap[projectName]) {
      projectMap[projectName] = createEmptyProjectRecordV4_(projectName);
      projectMap[projectName].trackerName = trackerName;
      projectMap[projectName].dateUpdated = dateUpdated || '';
    }

    const p = projectMap[projectName];

    if (trackerName) p.trackerName = trackerName;
    if (dateUpdated) p.dateUpdated = dateUpdated;

    if (quoteName || quoteId) {
      p.quotes.push({
        quoteName: quoteName || '(Unnamed Quote)',
        quoteId: quoteId,
        enabled: enabled,
        closingPct: closingPct,
        closingNote: closingNote
      });
    }
  });

  return projectMap;
}


function createEmptyProjectRecordV4_(projectName) {
  return {
    projectName: projectName,
    trackerName: `${projectName} - Project Tracker`,
    hyperlink: null,
    openTasks: 0,
    totalLineItems: 0,
    totalSold: 0,
    quoteTotal: 0,
    toBeOrdered: 0,
    onOrder: 0,
    received: 0,
    delivered: 0,
    quotes: [],
    projectClosing: '',
    toBeOrderedPct: 0,
    onOrderPct: 0,
    receivedPct: 0,
    deliveredPct: 0,
    stageSnapshot: '',
    trackerExists: false,
    tasksExists: false,
    dateUpdated: ''
  };
}


function buildDashboardHeaderV4_(dashboard) {
  const headers = [[
    'Project',
    'Quote',
    'Project Closing',
    'Open Tasks',
    'Total Line Items',
    'Total Sold',
    'Quote Total',
    '<85%',
    '>85%',
    '100%',
    'Billed',
    'Remaining Revenue',
    'Stage Snapshot'
  ]];

  dashboard.getRange(1, 1, 1, headers[0].length).setValues(headers);
  dashboard.getRange(1, 1, 1, headers[0].length)
    .setFontWeight('bold')
    .setBackground('#d9e2f3')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  dashboard.setRowHeight(1, 34);
}


function clearDashboardForRebuild_(dashboard) {
  const maxRows = dashboard.getMaxRows();
  const maxCols = Math.max(dashboard.getMaxColumns(), 16);

  try {
    if (dashboard.getFilter()) dashboard.getFilter().remove();
  } catch (e) {}

  try {
    for (let r = maxRows; r >= 2; r--) {
      let grp = dashboard.getRowGroup(r, 1);
      while (grp) {
        grp.remove();
        grp = dashboard.getRowGroup(r, 1);
      }
    }
  } catch (e) {}

  if (maxRows > 1) {
    dashboard.getRange(2, 1, maxRows - 1, maxCols)
      .clearContent()
      .clearFormat()
      .clearNote();
  }
}


function countOpenTasksOnly_(tasksSheet) {
  if (!tasksSheet) return 0;

  const lastRow = tasksSheet.getLastRow();
  if (lastRow < 2) return 0;

  const lastCol = tasksSheet.getLastColumn();
  const headers = tasksSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const statusCol = findHeaderIndexLoose_(headers, ['status']);

  if (statusCol === -1) return 0;

  const values = tasksSheet.getRange(2, statusCol + 1, lastRow - 1, 1).getValues();

  let count = 0;
  values.forEach(row => {
    const v = String(row[0] || '').trim().toLowerCase();
    if (v === 'open') count++;
  });

  return count;
}


function findHeaderIndexLoose_(headers, candidates) {
  const normalized = headers.map(h =>
    String(h || '').toLowerCase().replace(/\s+/g, ' ').trim()
  );

  for (let i = 0; i < normalized.length; i++) {
    for (let j = 0; j < candidates.length; j++) {
      const c = String(candidates[j] || '').toLowerCase().trim();
      if (normalized[i] === c || normalized[i].indexOf(c) !== -1) {
        return i;
      }
    }
  }

  return -1;
}


function safePercent_(numerator, denominator) {
  numerator = toNumber_(numerator);
  denominator = toNumber_(denominator);

  if (!denominator || denominator <= 0) return 0;

  return Math.min(1, numerator / denominator);
}


function formatPctText_(pct) {
  return Math.round((pct || 0) * 100) + '%';
}


function applyProjectQuoteRowStylingV4_(dashboard, rowTypes, startRow) {
  const totalCols = 13;

  for (let i = 0; i < rowTypes.length; i++) {
    const row = startRow + i;

    if (rowTypes[i] === 'project') {
      dashboard.getRange(row, 1, 1, totalCols).setFontWeight('bold');
      dashboard.getRange(row, 2).setBackground('#f7f9fc');
      dashboard.getRange(row, 4, 1, 4).setBackground('#f7f9fc');
    } else {
      dashboard.getRange(row, 2).setFontStyle('italic');
      dashboard.getRange(row, 2).setHorizontalAlignment('left');
      dashboard.getRange(row, 1, 1, totalCols).setFontSize(9);
    }
  }
}


function addRowGroupsForQuotesV4_(dashboard, rowTypes, startRow) {
  let i = 0;

  while (i < rowTypes.length) {
    if (rowTypes[i] !== 'project') {
      i++;
      continue;
    }

    const firstQuoteIndex = i + 1;
    let quoteCount = 0;
    let j = i + 1;

    while (j < rowTypes.length && rowTypes[j] === 'quote') {
      quoteCount++;
      j++;
    }

    if (quoteCount > 0) {
      try {
        dashboard.getRange(startRow + firstQuoteIndex, 1, quoteCount, 1).shiftRowGroupDepth(1);
        const group = dashboard.getRowGroup(startRow + firstQuoteIndex, 1);
        if (group) group.collapse();
      } catch (e) {}
    }

    i = j;
  }
}


function formatDashboardSheetV4_(dashboard, numRows, totalsRow) {
  const totalCols = 13;
  const bodyRows = Math.max(1, numRows);
  const lastDataRow = Math.max(1, numRows + 1);

  dashboard.getRange(2, 1, Math.max(1, totalsRow - 1), totalCols)
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setWrap(false);

  if (numRows > 0) {
    dashboard.getRange(2, 3, bodyRows, 2).setHorizontalAlignment('center');
    dashboard.getRange(2, 8, bodyRows, 5).setHorizontalAlignment('right');

    dashboard.getRange(2, 5, bodyRows, 1).setNumberFormat('#,##0');
    dashboard.getRange(2, 6, bodyRows, 7).setNumberFormat('$#,##0.00');
  }

  dashboard.getRange(1, 1, totalsRow, totalCols)
    .setBorder(true, true, true, true, true, true, '#d0d0d0', SpreadsheetApp.BorderStyle.SOLID);

  dashboard.setColumnWidth(1, 220);
  dashboard.setColumnWidth(2, 400);
  dashboard.setColumnWidth(3, 110);
  dashboard.setColumnWidth(4, 85);
  dashboard.setColumnWidth(5, 105);
  dashboard.setColumnWidth(6, 110);
  dashboard.setColumnWidth(7, 110);
  dashboard.setColumnWidth(8, 105);
  dashboard.setColumnWidth(9, 105);
  dashboard.setColumnWidth(10, 105);
  dashboard.setColumnWidth(11, 110);
  dashboard.setColumnWidth(12, 120);
  dashboard.setColumnWidth(13, 300);

  try {
    const bandings = dashboard.getBandings();
    bandings.forEach(b => b.remove());
  } catch (e) {}

  if (numRows > 0) {
    dashboard.getRange(1, 1, lastDataRow, totalCols)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

    dashboard.getRange(1, 8, lastDataRow, 1).setBackground('#f4cccc');
    dashboard.getRange(1, 9, lastDataRow, 1).setBackground('#fff2cc');
    dashboard.getRange(1, 10, lastDataRow, 1).setBackground('#d9ead3');
  }

  dashboard.getRange(1, 1, 1, totalCols)
    .setFontWeight('bold')
    .setBackground('#d9e2f3')
    .setHorizontalAlignment('center');

  dashboard.getRange(1, 8).setBackground('#f4cccc');
  dashboard.getRange(1, 9).setBackground('#fff2cc');
  dashboard.getRange(1, 10).setBackground('#d9ead3');

  dashboard.getRange(totalsRow, 6, 1, 7).setNumberFormat('$#,##0.00');

  try {
    if (dashboard.getFilter()) dashboard.getFilter().remove();
  } catch (e) {}

  if (numRows > 0) {
    try {
      dashboard.getRange(1, 1, lastDataRow, totalCols).createFilter();
    } catch (e) {}
  }
}


function buildProjectRichText_(text, linkUrl) {
  const builder = SpreadsheetApp.newRichTextValue().setText(text || '');

  if (linkUrl) {
    builder.setLinkUrl(linkUrl);
    builder.setTextStyle(
      0,
      String(text || '').length,
      SpreadsheetApp.newTextStyle()
        .setForegroundColor('#1155cc')
        .setUnderline(true)
        .build()
    );
  }

  return builder.build();
}


function getOrCreateDashboardSheet_(ss) {
  let sheet = ss.getSheetByName('Dashboard');
  if (!sheet) sheet = ss.insertSheet('Dashboard');
  return sheet;
}


function toNumber_(value) {
  if (typeof value === 'number') return value;

  const s = String(value || '').replace(/[$,]/g, '').trim();
  const n = parseFloat(s);

  return isNaN(n) ? 0 : n;
}

function readTrackerMetricsBySource_(trackerSheet) {
  const result = {
    bySource: {},
    projectTotals: createEmptyDashboardMetrics_()
  };

  if (!trackerSheet) return result;

  const lastRow = trackerSheet.getLastRow();
  if (lastRow < 4) return result;

  // Read A:Q
  const values = trackerSheet.getRange(1, 1, lastRow, 17).getValues();

  // Hard-coded from tracker layout:
  // B Source
  // H Status
  // Q Total
  const SOURCE_COL = 1; // B, zero-based
  const STATUS_COL = 7; // H
  const TOTAL_COL = 16; // Q

  for (let i = 3; i < values.length; i++) {
    const row = values[i];

    const source = String(row[SOURCE_COL] || '').trim();
    const status = String(row[STATUS_COL] || '').trim();
    const total = toNumber_(row[TOTAL_COL]);

    if (!source) continue;
    if (isDashboardSeparatorRow_(row)) continue;
    if (status.toLowerCase() === 'omitted') continue;

    if (!result.bySource[source]) {
      result.bySource[source] = createEmptyDashboardMetrics_();
    }

    addTrackerRowToMetrics_(result.bySource[source], status, total);
    addTrackerRowToMetrics_(result.projectTotals, status, total);
  }

  return result;
}


function createEmptyDashboardMetrics_() {
  return {
    totalLineItems: 0,
    totalSold: 0,
    quoteTotal: 0,
    toBeOrdered: 0,
    onOrder: 0,
    received: 0,
    delivered: 0
  };
}


function addTrackerRowToMetrics_(metrics, status, total) {
  const normalizedStatus = String(status || '').trim().toLowerCase();

  metrics.totalLineItems++;
  metrics.totalSold += toNumber_(total);
  metrics.quoteTotal += toNumber_(total);

  if (
    !normalizedStatus ||
    normalizedStatus === 'unapproved' ||
    normalizedStatus === 'approved'
  ) {
    metrics.toBeOrdered++;
    return;
  }

  if (normalizedStatus === 'ordered') {
    metrics.onOrder++;
    return;
  }

  if (normalizedStatus === 'received') {
    metrics.received++;
    return;
  }

  if (normalizedStatus === 'scheduled') {
    metrics.received++;
    return;
  }

  if (normalizedStatus === 'delivered') {
    metrics.delivered++;
    return;
  }

  // Unknown statuses stay in total line items and total sold,
  // but do not count toward a stage bucket.
}


function isDashboardSeparatorRow_(row) {
  const meaningful = row.filter(cell => String(cell || '').trim() !== '').length;

  // Gray spacer rows usually have very little meaningful text.
  return meaningful === 0;
}


function copyDashboardMetricsToProject_(projectRecord, metrics) {
  projectRecord.totalLineItems = metrics.totalLineItems;
  projectRecord.totalSold = metrics.totalSold;
  projectRecord.quoteTotal = metrics.quoteTotal;
  projectRecord.toBeOrdered = metrics.toBeOrdered;
  projectRecord.onOrder = metrics.onOrder;
  projectRecord.received = metrics.received;
  projectRecord.delivered = metrics.delivered;

  projectRecord.toBeOrderedPct = safePercent_(metrics.toBeOrdered, metrics.totalLineItems);
  projectRecord.onOrderPct = safePercent_(metrics.onOrder, metrics.totalLineItems);
  projectRecord.receivedPct = safePercent_(metrics.received, metrics.totalLineItems);
  projectRecord.deliveredPct = safePercent_(metrics.delivered, metrics.totalLineItems);

  projectRecord.stageSnapshot =
    'TB ' + formatPctText_(projectRecord.toBeOrderedPct) +
    ' | OO ' + formatPctText_(projectRecord.onOrderPct) +
    ' | RC ' + formatPctText_(projectRecord.receivedPct) +
    ' | DL ' + formatPctText_(projectRecord.deliveredPct);
}


function copyDashboardMetricsToQuote_(quoteRecord, metrics) {
  quoteRecord.totalLineItems = metrics.totalLineItems;
  quoteRecord.totalSold = metrics.totalSold;
  quoteRecord.quoteTotal = metrics.quoteTotal;
  quoteRecord.toBeOrdered = metrics.toBeOrdered;
  quoteRecord.onOrder = metrics.onOrder;
  quoteRecord.received = metrics.received;
  quoteRecord.delivered = metrics.delivered;
}

function getHighestProjectClosingPct_(quotes) {
  let best = '<85%';

  quotes.forEach(q => {
    const pct = String(q.closingPct || '').trim();

    if (pct === '100%') {
      best = '100%';
      return;
    }

    if (pct === '>85%' && best !== '100%') {
      best = '>85%';
    }
  });

  return best;
}

function colorDashboardClosingColumn_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Dashboard');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 3, lastRow - 1, 1); // Column C
  const values = range.getDisplayValues();

  const backgrounds = values.map(row => {
    const status = String(row[0] || '').trim();

    if (status === '100%') return ['#d9ead3'];  // green
    if (status === '>85%') return ['#fff2cc'];  // yellow
    if (status === '<85%') return ['#f4cccc'];  // red

    return ['#ffffff'];
  });

  range
    .setBackgrounds(backgrounds)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');
}

/**
 * Expands every grouped row on the Dashboard.
 */
function expandDashboard() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Dashboard');

  if (!sheet) throw new Error('Sheet "Dashboard" not found.');
  sheet.expandAllRowGroups();
}

/**
 * Collapses every grouped row on the Dashboard.
 */
function collapseDashboard() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Dashboard');

  if (!sheet) throw new Error('Sheet "Dashboard" not found.');
  sheet.collapseAllRowGroups();
}

/**
 * Prepares the Dashboard for a pipeline meeting.
 */
function switchToPipelineView() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Dashboard');

  if (!sheet) throw new Error('Sheet "Dashboard" not found.');

  // Start from a known visibility state.
  sheet.showColumns(1, 13);

  // Pipeline view shows A:C and H:M.
  sheet.hideColumns(4, 4); // Hide D:G

  setPipelineViewHeaders_(sheet);

  // Pipeline-friendly widths.
  sheet.setColumnWidth(1, 230);
  sheet.setColumnWidth(2, 400);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(8, 105);
  sheet.setColumnWidth(9, 105);
  sheet.setColumnWidth(10, 105);
  sheet.setColumnWidth(11, 110);
  sheet.setColumnWidth(12, 125);
  sheet.setColumnWidth(13, 300);

  sheet.expandAllRowGroups();
  sheet.setActiveSelection('A1');
}

/**
 * Restores the operations-focused Dashboard view.
 */
function switchToDashboardView() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Dashboard');

  if (!sheet) throw new Error('Sheet "Dashboard" not found.');

  // Start from a known visibility state.
  sheet.showColumns(1, 13);

  // Dashboard view shows A:G and M.
  sheet.hideColumns(8, 5); // Hide H:L

  setDashboardViewHeaders_(sheet);

  // Dashboard-friendly widths.
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 290);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 85);
  sheet.setColumnWidth(5, 105);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(13, 250);

  sheet.collapseAllRowGroups();
  sheet.setActiveSelection('A1');
}

/**
 * Applies the concise headers used during pipeline meetings.
 */
function setPipelineViewHeaders_(sheet) {
  sheet.getRange(1, 1, 1, 13).setValues([[
    'Project',
    'Quote',
    'Closing',
    'Open Tasks',
    'Items',
    'Sold',
    'Quote Total',
    '<85%',
    '>85%',
    '100%',
    'Billed',
    'Remaining',
    'Stage Snapshot'
  ]]);
}

/**
 * Applies the concise headers used for the operations Dashboard.
 */
function setDashboardViewHeaders_(sheet) {
  sheet.getRange(1, 1, 1, 13).setValues([[
    'Project',
    'Quote',
    'Closing',
    'Open Tasks',
    'Items',
    'Sold',
    'Quote Total',
    '<85%',
    '>85%',
    '100%',
    'Billed',
    'Remaining',
    'Stage Snapshot'
  ]]);
}
