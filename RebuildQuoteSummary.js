function rebuildQuoteSummary() {
  var ss = SpreadsheetApp.getActive();
  var summary = ss.getSheetByName('Quote Summary');
  var config = ss.getSheetByName('Tracker config') || ss.getSheetByName('Tracker Config');

  if (!summary) {
    throw new Error('Quote Summary sheet not found.');
  }
  if (!config) {
    throw new Error('Tracker Config sheet not found.');
  }

  // Clear old body
  var maxRows = summary.getMaxRows();
  if (maxRows > 1) {
    summary.getRange(2, 1, maxRows - 1, 11).clearContent().setBackground(null);
  }

  // Read Tracker Config
  // A Enable?
  // B Project
  // C trackerName
  // D Date Updated
  // E Quote Name
  // F Quote Sheet ID
  var lastConfigRow = config.getLastRow();
  if (lastConfigRow < 2) return;

  var configData = config.getRange(2, 1, lastConfigRow - 1, 6).getValues();

  // Group configured quotes by tracker
  var trackerMap = {}; // { trackerName: { projectName, quotes: [{quoteName, quoteId}] } }

  configData.forEach(function(row) {
    var enabled = row[0];
    var projectName = row[1];
    var trackerName = row[2];
    var quoteName = row[4];
    var quoteId = row[5];

    if (enabled !== true) return;
    if (!trackerName || !projectName || !quoteName) return;
    if (!/\bquotation\b/i.test(String(quoteName))) return;

    if (!trackerMap[trackerName]) {
      trackerMap[trackerName] = {
        projectName: projectName,
        quotes: []
      };
    }

    trackerMap[trackerName].quotes.push({
      quoteName: String(quoteName).trim(),
      quoteId: String(quoteId || '').trim()
    });
  });

  var output = [];

  Object.keys(trackerMap).forEach(function(trackerName) {
    var trackerSheet = ss.getSheetByName(trackerName);
    if (!trackerSheet) return;

    var projectName = trackerMap[trackerName].projectName;
    var configuredQuotes = trackerMap[trackerName].quotes;
    var openTasks = getOpenTaskCount_(ss, projectName);

    var lastRow = trackerSheet.getLastRow();
    var projectTotals = {
      lineItems: 0,
      totalSold: 0,
      toBeOrdered: 0,
      onOrder: 0,
      onHand: 0,
      delivered: 0
    };

    var quoteSummaries = [];

    if (lastRow >= 4) {
      var numRows = lastRow - 3;

      // A:Q
      var values = trackerSheet.getRange(4, 1, numRows, 17).getValues();
      var displayValues = trackerSheet.getRange(4, 1, numRows, 17).getDisplayValues();
      var formulasB = trackerSheet.getRange(4, 2, numRows, 1).getFormulas();

      configuredQuotes.forEach(function(q) {
        var quoteSummary = summarizeQuoteFromTracker_(values, displayValues, formulasB, q.quoteName);

        quoteSummary.quoteName = q.quoteName;
        quoteSummary.closing = getQuoteClosingFromSpreadsheetId_(q.quoteId);

        projectTotals.lineItems += quoteSummary.lineItems;
        projectTotals.totalSold += quoteSummary.total;
        projectTotals.toBeOrdered += quoteSummary.toBeOrdered;
        projectTotals.onOrder += quoteSummary.onOrder;
        projectTotals.onHand += quoteSummary.onHand;
        projectTotals.delivered += quoteSummary.delivered;

        quoteSummaries.push(quoteSummary);
      });
    }

    var projectClosing = rollupProjectClosingFromQuotes_(quoteSummaries);

    var trackerId = trackerSheet.getSheetId();
    var safeProjectName = String(projectName).replace(/"/g, '""');

    // Project row
    output.push([
      '=HYPERLINK("#gid=' + trackerId + '&range=A1","' + safeProjectName + '")',
      '',
      projectClosing,
      openTasks,
      projectTotals.lineItems,
      projectTotals.totalSold,
      '',
      projectTotals.toBeOrdered,
      projectTotals.onOrder,
      projectTotals.onHand,
      projectTotals.delivered
    ]);

    // Quote rows
    quoteSummaries.forEach(function(qs) {
      output.push([
        projectName,
        qs.quoteName,
        qs.closing,
        '',
        qs.lineItems,
        '',
        qs.total,
        qs.toBeOrdered,
        qs.onOrder,
        qs.onHand,
        qs.delivered
      ]);
    });

    // Spacer row
    output.push(['', '', '', '', '', '', '', '', '', '', '']);
  });

  if (output.length > 0) {
    summary.getRange(2, 1, output.length, 11).setValues(output);
    summary.getRange(2, 6, output.length, 2).setNumberFormat('$#,##0.00');
    applyStatusColors_(summary);
  }
}


/**
 * Summarize one configured quote by matching Quote Name against tracker col B hyperlink label.
 */
function summarizeQuoteFromTracker_(values, displayValues, formulasB, targetQuoteName) {
  var summary = {
    quoteName: targetQuoteName,
    lineItems: 0,
    total: 0,
    toBeOrdered: 0,
    onOrder: 0,
    onHand: 0,
    delivered: 0,
    closing: '85%'
  };

  var target = normalizeQuoteName_(targetQuoteName);

  for (var i = 0; i < formulasB.length; i++) {
    var formulaB = formulasB[i][0];
    if (!formulaB) continue;

    var parsed = parseHyperlinkFormula_(formulaB);
    var rowQuoteName = normalizeQuoteName_(parsed.label || displayValues[i][1] || '');

    if (rowQuoteName !== target) continue;

    var qty = numericOrZero_(values[i][6]);          // G
    var status = String(displayValues[i][7] || '');  // H
    var total = numericOrZero_(values[i][16]);       // Q

    summary.lineItems += 1;
    summary.total += total;

    var bucket = mapFulfillmentBucket_(status);
    summary[bucket] += qty;
  }

  return summary;
}


/**
 * Roll up project closing from child quotes.
 * Rule:
 * - if any quote is 100% => project is 100%
 * - else if any quote is < 85% => project is < 85%
 * - else => 85%
 */
function rollupProjectClosingFromQuotes_(quoteSummaries) {
  var has100 = false;
  var hasUnder85 = false;
  var has85 = false;

  quoteSummaries.forEach(function(q) {
    var closing = String(q.closing || '').trim();

    if (closing === '100%') has100 = true;
    else if (closing === '< 85%') hasUnder85 = true;
    else if (closing === '85%') has85 = true;
  });

  if (has100) return '100%';
  if (hasUnder85) return '< 85%';
  if (has85) return '85%';

  return '85%';
}


function normalizeQuoteName_(name) {
  return String(name || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}


function getOpenTaskCount_(ss, projectName) {
  var tasksSheet = ss.getSheetByName(projectName + ' - Tasks');
  if (!tasksSheet) return 0;

  var lastRow = tasksSheet.getLastRow();
  if (lastRow < 2) return 0;

  var statuses = tasksSheet.getRange(2, 5, lastRow - 1, 1).getDisplayValues();
  var count = 0;

  statuses.forEach(function(r) {
    if (String(r[0]).trim().toLowerCase() === 'open') count++;
  });

  return count;
}


function getQuoteClosingFromSpreadsheetId_(spreadsheetId) {
  if (!spreadsheetId) return '85%';

  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheets = ss.getSheets();
    if (!sheets || !sheets.length) return '85%';

    var firstSheet = sheets[0];
    var firstTabName = String(firstSheet.getName() || '');
    var isApproved = /approved/i.test(firstTabName);

    var protections = firstSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    var isProtected = protections && protections.length > 0;

    return (isApproved && isProtected) ? '100%' : '85%';
  } catch (e) {
    Logger.log('Could not inspect quote spreadsheet ' + spreadsheetId + ': ' + e);
    return '85%';
  }
}


function parseHyperlinkFormula_(formula) {
  var m = String(formula).match(/=HYPERLINK\("([^"]+)"\s*,\s*"([^"]*)"\)/i);
  if (m) {
    return { url: m[1], label: m[2] };
  }
  return { url: '', label: '' };
}


function mapFulfillmentBucket_(status) {
  var s = String(status || '').trim().toLowerCase();

  if (!s) return 'toBeOrdered';
  if (s.indexOf('delivered') !== -1) return 'delivered';
  if (s.indexOf('received') !== -1 || s.indexOf('on hand') !== -1) return 'onHand';
  if (s.indexOf('order') !== -1) return 'onOrder';

  return 'toBeOrdered';
}


function numericOrZero_(v) {
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v).replace(/[$,]/g, '').trim());
  return isNaN(n) ? 0 : n;
}


function applyStatusColors_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var range = sheet.getRange(2, 3, lastRow - 1, 1);
  var vals = range.getValues();

  var colors = vals.map(function(r) {
    var s = String(r[0]).trim();

    if (s === '100%') return ['#c6e0b4'];
    if (s === '85%') return ['#f1dd96'];
    if (s === '< 85%') return ['#f4c7b5'];

    return [null];
  });

  range.setBackgrounds(colors);
}