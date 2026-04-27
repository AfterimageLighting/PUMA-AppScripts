function refreshDashboardProjectTrackers() {
  var ss = SpreadsheetApp.getActive();
  var dashboard = ss.getSheetByName('Dashboard');
  var quoteSummary = ss.getSheetByName('Quote Summary');

  if (!dashboard) throw new Error('Sheet "Dashboard" not found.');
  if (!quoteSummary) throw new Error('Sheet "Quote Summary" not found.');

  clearDashboardForRebuild_(dashboard);
  buildDashboardHeaderV4_(dashboard);

  var qsLastRow = quoteSummary.getLastRow();
  if (qsLastRow < 2) {
    formatDashboardSheetV4_(dashboard, 0);
    return;
  }

  // Quote Summary columns:
  // A Project
  // B Quote
  // C Project Closing
  // E Total Line Items
  // F Total Sold
  // G Quote Total
  // H To Be Ordered  (line count)
  // I On Order       (line count)
  // J Received       (line count; still stored in old "On Hand" col)
  // K Delivered      (line count)
  var qsValues = quoteSummary.getRange(2, 1, qsLastRow - 1, 11).getValues();
  var qsRich = quoteSummary.getRange(2, 1, qsLastRow - 1, 1).getRichTextValues();

  var projectMap = {};

  qsValues.forEach(function(row, i) {
    var projectCell = row[0];
    var quoteName = String(row[1] || '').trim();
    var closing = normalizeClosing_(row[2]);
    var totalLineItems = toNumber_(row[4]);
    var totalSold = toNumber_(row[5]);
    var quoteTotal = toNumber_(row[6]);
    var toBeOrdered = toNumber_(row[7]);
    var onOrder = toNumber_(row[8]);
    var received = toNumber_(row[9]);
    var delivered = toNumber_(row[10]);

    var richProject = qsRich[i][0];
    var hyperlink = richProject ? richProject.getLinkUrl() : null;
    var displayText = richProject ? richProject.getText() : String(projectCell || '').trim();

    var isProjectRow = !quoteName;

    if (isProjectRow) {
      var projectName = displayText || String(projectCell || '').trim();
      if (!projectName) return;

      if (!projectMap[projectName]) {
        projectMap[projectName] = createEmptyProjectRecordV4_(projectName);
      }

      var p = projectMap[projectName];
      p.projectName = projectName;
      p.hyperlink = hyperlink;
      p.totalLineItems = totalLineItems;
      p.totalSold = totalSold;
      p.toBeOrdered = toBeOrdered;
      p.onOrder = onOrder;
      p.received = received;
      p.delivered = delivered;
      p.projectClosing = closing;
      return;
    }

    var projectNameFromRow = String(projectCell || '').trim();
    if (!projectNameFromRow) return;

    if (!projectMap[projectNameFromRow]) {
      projectMap[projectNameFromRow] = createEmptyProjectRecordV4_(projectNameFromRow);
    }

    projectMap[projectNameFromRow].quotes.push({
      quoteName: quoteName,
      closing: closing,
      totalLineItems: totalLineItems,
      totalSold: totalSold,
      quoteTotal: quoteTotal,
      toBeOrdered: toBeOrdered,
      onOrder: onOrder,
      received: received,
      delivered: delivered
    });
  });

  var sheets = ss.getSheets();
  var sheetMap = {};
  sheets.forEach(function(sh) {
    sheetMap[sh.getName()] = sh;
  });

  Object.keys(projectMap).forEach(function(projectName) {
    var p = projectMap[projectName];
    p.openTasks = countOpenTasksOnly_(sheetMap[projectName + ' - Tasks']);

    // Re-derive closing from quotes if quote rows exist
    if (p.quotes.length) {
      p.projectClosing = deriveProjectClosingFromQuotesV4_(p.quotes);
    }

    p.toBeOrderedPct = safePercent_(p.toBeOrdered, p.totalLineItems);
    p.onOrderPct = safePercent_(p.onOrder, p.totalLineItems);
    p.receivedPct = safePercent_(p.received, p.totalLineItems);
    p.deliveredPct = safePercent_(p.delivered, p.totalLineItems);

    p.stageSnapshot = buildStageSnapshot_(
      p.toBeOrderedPct,
      p.onOrderPct,
      p.receivedPct,
      p.deliveredPct
    );
  });

  var projects = Object.keys(projectMap).map(function(k) {
    return projectMap[k];
  });

  projects.sort(function(a, b) {
    var rankA = closingRank_(a.projectClosing);
    var rankB = closingRank_(b.projectClosing);

    if (rankA !== rankB) return rankA - rankB;
    if (b.openTasks !== a.openTasks) return b.openTasks - a.openTasks;
    return a.projectName.localeCompare(b.projectName);
  });

  var values = [];
  var richProjectCells = [];
  var closingBackgrounds = [];
  var snapshotBackgrounds = [];
  var rowTypes = [];

  projects.forEach(function(p) {
    values.push([
      p.projectName,        // A Project
      '',                   // B Quote
      p.projectClosing,     // C Project Closing
      p.openTasks,          // D Open Tasks
      p.totalLineItems,     // E Total Line Items
      p.totalSold,          // F Total Sold
      '',                   // G Quote Total
      p.toBeOrdered,        // H To Be Ordered
      p.onOrder,            // I On Order
      p.received,           // J Received
      p.delivered,          // K Delivered
      p.toBeOrderedPct,     // L TB Ordered %
      p.onOrderPct,         // M On Order %
      p.receivedPct,        // N Received %
      p.deliveredPct,       // O Delivered %
      p.stageSnapshot       // P Stage Snapshot
    ]);

    richProjectCells.push([buildProjectRichText_(p.projectName, p.hyperlink)]);
    closingBackgrounds.push([closingColor_(p.projectClosing)]);
    snapshotBackgrounds.push([snapshotColor_(p.deliveredPct)]);
    rowTypes.push('project');

    p.quotes.forEach(function(q) {
      var tbPct = safePercent_(q.toBeOrdered, q.totalLineItems);
      var ooPct = safePercent_(q.onOrder, q.totalLineItems);
      var rcPct = safePercent_(q.received, q.totalLineItems);
      var dlPct = safePercent_(q.delivered, q.totalLineItems);

      values.push([
        '',
        '↳ ' + q.quoteName,
        q.closing,
        '',
        q.totalLineItems,
        q.totalSold,
        q.quoteTotal,
        q.toBeOrdered,
        q.onOrder,
        q.received,
        q.delivered,
        tbPct,
        ooPct,
        rcPct,
        dlPct,
        buildStageSnapshot_(tbPct, ooPct, rcPct, dlPct)
      ]);

      richProjectCells.push([SpreadsheetApp.newRichTextValue().setText('').build()]);
      closingBackgrounds.push([closingColor_(q.closing)]);
      snapshotBackgrounds.push([snapshotColor_(dlPct)]);
      rowTypes.push('quote');
    });
  });

  if (!values.length) {
    formatDashboardSheetV4_(dashboard, 0);
    return;
  }

  var startRow = 2;
  var numRows = values.length;
  var numCols = values[0].length;

  dashboard.getRange(startRow, 1, numRows, numCols).setValues(values);
  dashboard.getRange(startRow, 1, numRows, 1).setRichTextValues(richProjectCells);
  dashboard.getRange(startRow, 3, numRows, 1).setBackgrounds(closingBackgrounds);
  dashboard.getRange(startRow, 16, numRows, 1).setBackgrounds(snapshotBackgrounds);

  formatDashboardSheetV4_(dashboard, numRows);
  applyProjectQuoteRowStylingV4_(dashboard, rowTypes, startRow);
  addRowGroupsForQuotesV4_(dashboard, rowTypes, startRow);

  dashboard.setFrozenRows(1);
}


function createEmptyProjectRecordV4_(projectName) {
  return {
    projectName: projectName,
    hyperlink: null,
    openTasks: 0,
    totalLineItems: 0,
    totalSold: 0,
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
    stageSnapshot: ''
  };
}


function buildDashboardHeaderV4_(dashboard) {
  var headers = [[
    'Project',
    'Quote',
    'Project Closing',
    'Open Tasks',
    'Total Line Items',
    'Total Sold',
    'Quote Total',
    'To Be Ordered',
    'On Order',
    'Received',
    'Delivered',
    'TB Ordered %',
    'On Order %',
    'Received %',
    'Delivered %',
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
  var maxRows = dashboard.getMaxRows();
  var maxCols = Math.max(dashboard.getMaxColumns(), 16);

  try {
    if (dashboard.getFilter()) dashboard.getFilter().remove();
  } catch (e) {}

  try {
    for (var r = maxRows; r >= 2; r--) {
      var grp = dashboard.getRowGroup(r, 1);
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

  var lastRow = tasksSheet.getLastRow();
  if (lastRow < 2) return 0;

  var lastCol = tasksSheet.getLastColumn();
  var headers = tasksSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var statusCol = findHeaderIndexLoose_(headers, ['status']);
  if (statusCol === -1) return 0;

  var values = tasksSheet.getRange(2, statusCol + 1, lastRow - 1, 1).getValues();
  var count = 0;

  values.forEach(function(r) {
    var v = String(r[0] || '').trim().toLowerCase();
    if (v === 'open') count++;
  });

  return count;
}


function findHeaderIndexLoose_(headers, candidates) {
  var normalized = headers.map(function(h) {
    return String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
  });

  for (var i = 0; i < normalized.length; i++) {
    for (var j = 0; j < candidates.length; j++) {
      var c = String(candidates[j] || '').toLowerCase().trim();
      if (normalized[i] === c || normalized[i].indexOf(c) !== -1) {
        return i;
      }
    }
  }
  return -1;
}


function deriveProjectClosingFromQuotesV4_(quotes) {
  if (!quotes || !quotes.length) return '';

  var has100 = false;
  var hasOpen = false;
  var has85 = false;

  quotes.forEach(function(q) {
    var c = normalizeClosing_(q.closing);

    if (c === '100%') {
      has100 = true;
    } else if (c === '< 85%') {
      hasOpen = true;
    } else if (c === '85%') {
      has85 = true;
    }
  });

  // New rule:
  // If ANY quote is 100%, show project as 100%
  if (has100) return '100%';

  // Otherwise preserve the existing hierarchy
  if (hasOpen) return '< 85%';
  if (has85) return '85%';

  return '';
}


function sumQuoteTotals_(projectRecord) {
  var total = 0;
  projectRecord.quotes.forEach(function(q) {
    total += toNumber_(q.quoteTotal);
  });
  return total;
}


function safePercent_(numerator, denominator) {
  numerator = toNumber_(numerator);
  denominator = toNumber_(denominator);
  if (!denominator || denominator <= 0) return 0;
  return Math.min(1, numerator / denominator);
}


function buildStageSnapshot_(tbPct, ooPct, rcPct, dlPct) {
  return 'TB ' + formatPctText_(tbPct) +
         ' | OO ' + formatPctText_(ooPct) +
         ' | RC ' + formatPctText_(rcPct) +
         ' | DL ' + formatPctText_(dlPct);
}


function formatPctText_(pct) {
  return Math.round((pct || 0) * 100) + '%';
}


function applyProjectQuoteRowStylingV4_(dashboard, rowTypes, startRow) {
  for (var i = 0; i < rowTypes.length; i++) {
    var row = startRow + i;

    if (rowTypes[i] === 'project') {
      dashboard.getRange(row, 1, 1, 16).setFontWeight('bold');
      dashboard.getRange(row, 2).setBackground('#f7f9fc');
      dashboard.getRange(row, 4, 1, 13).setBackground('#f7f9fc');
    } else {
      dashboard.getRange(row, 2).setFontStyle('italic');
      dashboard.getRange(row, 2).setHorizontalAlignment('left');
      dashboard.getRange(row, 1, 1, 16).setFontSize(9);
    }
  }
}


function addRowGroupsForQuotesV4_(dashboard, rowTypes, startRow) {
  var i = 0;

  while (i < rowTypes.length) {
    if (rowTypes[i] !== 'project') {
      i++;
      continue;
    }

    var firstQuoteIndex = i + 1;
    var quoteCount = 0;
    var j = i + 1;

    while (j < rowTypes.length && rowTypes[j] === 'quote') {
      quoteCount++;
      j++;
    }

    if (quoteCount > 0) {
      try {
        dashboard.getRange(startRow + firstQuoteIndex, 1, quoteCount, 1).shiftRowGroupDepth(1);
        var group = dashboard.getRowGroup(startRow + firstQuoteIndex, 1);
        if (group) group.collapse();
      } catch (e) {}
    }

    i = j;
  }
}


function formatDashboardSheetV4_(dashboard, numRows) {
  var totalCols = 16;
  var lastRow = Math.max(2, numRows + 1);
  var bodyRows = Math.max(1, numRows);

  dashboard.getRange(2, 1, bodyRows, totalCols)
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setWrap(false);

  dashboard.getRange(2, 3, bodyRows, 2).setHorizontalAlignment('center');  // C:D
  dashboard.getRange(2, 8, bodyRows, 4).setHorizontalAlignment('center');  // H:K
  dashboard.getRange(2, 12, bodyRows, 4).setHorizontalAlignment('center'); // L:O
  dashboard.getRange(2, 6, bodyRows, 2).setHorizontalAlignment('right');   // F:G

  dashboard.getRange(2, 5, bodyRows, 1).setNumberFormat('#,##0');
  dashboard.getRange(2, 6, bodyRows, 2).setNumberFormat('$#,##0.00');
  dashboard.getRange(2, 8, bodyRows, 4).setNumberFormat('#,##0');
  dashboard.getRange(2, 12, bodyRows, 4).setNumberFormat('0%');

  dashboard.getRange(1, 1, lastRow, totalCols)
    .setBorder(true, true, true, true, true, true, '#d0d0d0', SpreadsheetApp.BorderStyle.SOLID);

  dashboard.setColumnWidth(1, 220); // Project
  dashboard.setColumnWidth(2, 290); // Quote
  dashboard.setColumnWidth(3, 110); // Project Closing
  dashboard.setColumnWidth(4, 85);  // Open Tasks
  dashboard.setColumnWidth(5, 105); // Total Line Items
  dashboard.setColumnWidth(6, 110); // Total Sold
  dashboard.setColumnWidth(7, 110); // Quote Total
  dashboard.setColumnWidth(8, 95);  // To Be Ordered
  dashboard.setColumnWidth(9, 85);  // On Order
  dashboard.setColumnWidth(10, 85); // Received
  dashboard.setColumnWidth(11, 85); // Delivered
  dashboard.setColumnWidth(12, 90); // TB Ordered %
  dashboard.setColumnWidth(13, 85); // On Order %
  dashboard.setColumnWidth(14, 85); // Received %
  dashboard.setColumnWidth(15, 85); // Delivered %
  dashboard.setColumnWidth(16, 250); // Snapshot

  try {
    var bandings = dashboard.getBandings();
    bandings.forEach(function(b) { b.remove(); });
  } catch (e) {}

  if (numRows > 0) {
    dashboard.getRange(1, 1, lastRow, totalCols)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

    dashboard.getRange(1, 1, 1, totalCols)
      .setFontWeight('bold')
      .setBackground('#d9e2f3')
      .setHorizontalAlignment('center');
  }

  try {
    if (dashboard.getFilter()) dashboard.getFilter().remove();
  } catch (e) {}

  try {
    dashboard.getRange(1, 1, lastRow, totalCols).createFilter();
  } catch (e) {}
}


function buildProjectRichText_(text, linkUrl) {
  var builder = SpreadsheetApp.newRichTextValue().setText(text || '');

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


function snapshotColor_(deliveredPct) {
  if (deliveredPct >= 0.85) return '#c6e0b4';
  if (deliveredPct >= 0.40) return '#fce5cd';
  return '#f4cccc';
}


function normalizeClosing_(value) {
  if (value === null || value === '') return '';

  // Handle actual numbers or percent-formatted numeric values
  if (typeof value === 'number') {
    if (value >= 0.9999) return '100%';
    if (value >= 0.8499 && value < 0.9999) return '85%';
    return '< 85%';
  }

  var s = String(value).trim().toLowerCase().replace(/\s+/g, '');

  if (s === '100%' || s === '100') return '100%';
  if (s === '85%' || s === '85' || s === '0.85') return '85%';
  if (s === '<85%' || s === '<85') return '< 85%';

  // Handle strings that are numeric
  var n = parseFloat(s.replace('%', ''));
  if (!isNaN(n)) {
    if (n >= 99.99) return '100%';
    if (n >= 0.9999) return '100%';   // catches "1"
    if (n >= 84.99 && n < 99.99) return '85%';
    if (n >= 0.8499 && n < 0.9999) return '85%'; // catches "0.85"
    return '< 85%';
  }

  return '';
}


function closingRank_(closing) {
  if (closing === '< 85%') return 1;
  if (closing === '85%') return 2;
  if (closing === '100%') return 3;
  return 4;
}


function closingColor_(closing) {
  if (closing === '100%') return '#c6e0b4';
  if (closing === '85%') return '#f1dd96';
  if (closing === '< 85%') return '#f4c7b5';
  return '#ffffff';
}


function toNumber_(value) {
  if (typeof value === 'number') return value;
  var s = String(value || '').replace(/[$,]/g, '').trim();
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}