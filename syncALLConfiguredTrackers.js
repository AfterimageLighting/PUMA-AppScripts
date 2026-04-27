// ===== Generic quote → tracker sync utilities =====

const QUOTE_START_ROW = 10;
const QUOTE_BG_CHECK_COLUMN = 2; // use column B to detect grey area

function syncAllConfiguredTrackers() {
  var ss = SpreadsheetApp.getActive();
  var configSheet = getTrackerConfigSheet_(ss);

  if (!configSheet) {
    throw new Error('Config sheet "Tracker config" not found.');
  }

  var lastRow = configSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No config rows found on Tracker config.');
    return;
  }

  // Tracker config layout:
  // A = Enable?
  // B = Project
  // C = trackerName
  // D = Date Updated
  // E = Quote Name   (for logging only)
  // F = Quote Sheet ID
  var data = configSheet.getRange(2, 1, lastRow - 1, 6).getValues();

  // Tracks next available write row per tracker sheet during this run
  var nextRowByTracker = {};

  // Tracks project-level closing result per tracker
  // default will become 85% once tracker is encountered
  var closingByTracker = {}; // { trackerName: 'Probability of close 85%' or 'Probability of close 100%' }

  data.forEach(function(row, idx) {
    var rowNum      = idx + 2;
    var enabled     = row[0]; // A
    var projectName = row[1]; // B
    var trackerName = row[2]; // C
    var quoteLabel  = row[4]; // E
    var quoteId     = row[5]; // F

    if (enabled !== true) {
      Logger.log('Row ' + rowNum + ' (' + projectName + ') disabled; skipping.');
      return;
    }

    if (!trackerName) {
      Logger.log('Row ' + rowNum + ' (' + projectName + ') missing trackerName; skipping.');
      return;
    }

    try {
      // First time we encounter this tracker in the run
      if (nextRowByTracker[trackerName] === undefined) {
        ensureTopProbabilitySpacer_(trackerName);
        clearQuoteSpacerRows_(trackerName);
        nextRowByTracker[trackerName] = 4;

        // Default project closing to 85% unless a qualifying approved quote upgrades it
        closingByTracker[trackerName] = 'Probability of close 85%';
      }

      // If no quote ID, leave this tracker at default 85%
      if (!quoteId) {
        Logger.log(
          'Row ' + rowNum +
          ' (' + projectName + ')' +
          ' | Tracker: ' + trackerName +
          ' | Quote: ' + (quoteLabel || '[Unnamed Quote]') +
          ' | No Quote Sheet ID; leaving project closing at 85%.'
        );
        return;
      }

      var startRow = nextRowByTracker[trackerName];

      Logger.log(
        'Syncing row ' + rowNum +
        ' | Project: ' + projectName +
        ' | Tracker: ' + trackerName +
        ' | Quote: ' + (quoteLabel || '[Unnamed Quote]') +
        ' | Start Row: ' + startRow
      );

      var result = syncQuoteToTracker_(quoteId, trackerName, startRow, quoteLabel);

      // Only insert quote spacer if rows were actually written
      if (result.numRows > 0) {
        insertQuoteSpacerRows_(trackerName, result.nextRow);
        nextRowByTracker[trackerName] = result.nextRow + 2;
      }

      // Upgrade tracker/project closing to 100% if:
      // - quote parsed
      // - first tab is protected
      // - first tab name contains "-Approved"
      if (result.numRows > 0 && result.isApprovedLocked === true) {
        closingByTracker[trackerName] = 'Probability of close 100%';
      }

      // Stamp Date Updated on success
      configSheet.getRange(rowNum, 4).setValue(new Date());

      Logger.log(
        'Finished row ' + rowNum +
        ' | Project: ' + projectName +
        ' | Tracker: ' + trackerName +
        ' | Quote: ' + (quoteLabel || '[Unnamed Quote]') +
        ' | Source Tab Used: ' + result.sourceTabName +
        ' | Parsed Rows: ' + result.numRows +
        ' | Approved+Locked: ' + result.isApprovedLocked +
        ' | Next Row: ' + result.nextRow
      );

    } catch (e) {
      Logger.log(
        'Error syncing row ' + rowNum +
        ' (' + projectName + ')' +
        ' | Tracker: ' + trackerName +
        ' | Quote: ' + (quoteLabel || '[Unnamed Quote]') +
        ' | Error: ' + e
      );

      // If a quote errors, tracker remains at default 85%
      if (closingByTracker[trackerName] === undefined) {
        closingByTracker[trackerName] = 'Probability of close 85%';
      }
    }
  });

  // After all quotes are processed, set the project-level closing dropdown in A3
  Object.keys(closingByTracker).forEach(function(trackerName) {
    try {
      setTrackerProjectClosing_(trackerName, closingByTracker[trackerName]);
      Logger.log(
        'Set project closing for tracker "' + trackerName +
        '" to "' + closingByTracker[trackerName] + '".'
      );
    } catch (e) {
      Logger.log(
        'Failed setting project closing for tracker "' + trackerName +
        '": ' + e
      );
    }
  });
}


/**
 * Returns the Tracker config sheet regardless of capitalization style.
 */
function getTrackerConfigSheet_(ss) {
  return ss.getSheetByName('Tracker config') ||
         ss.getSheetByName('Tracker Config');
}


/**
 * Sync one quote workbook into one tracker sheet,
 * always using the FIRST tab in the quote workbook.
 *
 * Returns:
 * {
 *   nextRow: number,
 *   sourceTabName: string,
 *   numRows: number,
 *   isApprovedLocked: boolean
 * }
 */
function syncQuoteToTracker_(quoteSpreadsheetId, trackerSheetName, startRow, quoteLabel) {
  var pumaSS = SpreadsheetApp.getActive();

  var trackerSheet = pumaSS.getSheetByName(trackerSheetName);
  if (!trackerSheet) {
    throw new Error('Tracker sheet "' + trackerSheetName + '" not found in PUMA workbook.');
  }

  if (!startRow || startRow < 4) {
    startRow = 4;
  }

  // --- Open quote workbook + ALWAYS use first tab ---
  var quoteSS = SpreadsheetApp.openById(quoteSpreadsheetId);
  var quoteSheets = quoteSS.getSheets();

  if (!quoteSheets || quoteSheets.length === 0) {
    throw new Error('No sheets found in quote spreadsheet ' + quoteSpreadsheetId);
  }

  var quoteSheet = quoteSheets[0];
  var sourceTabName = quoteSheet.getName();

  // Check approval/lock status of FIRST tab
  var isApprovedLocked = isApprovedLockedSheet_(quoteSheet);

  Logger.log(
    'Using first tab "' + sourceTabName +
    '" for quote "' + (quoteLabel || '[Unnamed Quote]') +
    '" from spreadsheet ' + quoteSpreadsheetId +
    ' | Approved+Locked=' + isApprovedLocked
  );

  var lastRow = quoteSheet.getLastRow();
  if (lastRow < QUOTE_START_ROW) {
    Logger.log(
      'Quote "' + (quoteLabel || '[Unnamed Quote]') +
      '" has no data at or below row ' + QUOTE_START_ROW
    );
    return {
      nextRow: startRow,
      sourceTabName: sourceTabName,
      numRows: 0,
      isApprovedLocked: false
    };
  }

  // --- Determine end row by:
  //  (1) grey area (background color change),
  //  (2) "SubT Check" row,
  //  (3) 3 consecutive blank rows in columns A & B.
  var endByColor  = findEndRowByBackground_(quoteSheet, QUOTE_START_ROW, lastRow, QUOTE_BG_CHECK_COLUMN);
  var endBySubt   = findEndRowBySubtCheck_(quoteSheet, QUOTE_START_ROW, lastRow);
  var endByBlanks = findEndRowByBlankStreak_(quoteSheet, QUOTE_START_ROW, lastRow, 3);

  var candidates = [];
  if (endByColor  >= QUOTE_START_ROW) candidates.push(endByColor);
  if (endBySubt   >= QUOTE_START_ROW) candidates.push(endBySubt);
  if (endByBlanks >= QUOTE_START_ROW) candidates.push(endByBlanks);

  var endRow = candidates.length ? Math.min.apply(null, candidates) : lastRow;

  var numRows = endRow - QUOTE_START_ROW + 1;
  if (numRows <= 0) {
    Logger.log(
      'No data rows found for quote "' + (quoteLabel || '[Unnamed Quote]') +
      '" between row ' + QUOTE_START_ROW + ' and detected end row.'
    );
    return {
      nextRow: startRow,
      sourceTabName: sourceTabName,
      numRows: 0,
      isApprovedLocked: false
    };
  }

  // --- Build "Source" hyperlink block (column B) ---
  var sourceDisplayName = quoteLabel || quoteSS.getName();
  var quoteUrl = 'https://docs.google.com/spreadsheets/d/' +
                 quoteSpreadsheetId +
                 '/edit#gid=' + quoteSheet.getSheetId();
  var safeLabel = String(sourceDisplayName).replace(/"/g, '""');
  var sourceFormula = '=HYPERLINK("' + quoteUrl + '","' + safeLabel + '")';

  var sourceBlock = [];
  for (var r = 0; r < numRows; r++) {
    sourceBlock.push([sourceFormula]);
  }

  // --- Read data from quote ---
  var typeValues               = quoteSheet.getRange(QUOTE_START_ROW, 2, numRows, 1).getValues();
  var partNumberValues         = quoteSheet.getRange(QUOTE_START_ROW, 5, numRows, 1).getValues();
  var descriptionValues        = quoteSheet.getRange(QUOTE_START_ROW, 3, numRows, 1).getValues();
  var manufacturerValues       = quoteSheet.getRange(QUOTE_START_ROW, 4, numRows, 1).getValues();
  var quantityValues           = quoteSheet.getRange(QUOTE_START_ROW, 1, numRows, 1).getValues();
  var costPerUnitValues        = quoteSheet.getRange(QUOTE_START_ROW, 8, numRows, 1).getValues();
  var costPerUnitMarginValues  = quoteSheet.getRange(QUOTE_START_ROW, 7, numRows, 1).getValues();
  var totalCostValues          = quoteSheet.getRange(QUOTE_START_ROW, 11, numRows, 1).getValues();
  var totalValues              = quoteSheet.getRange(QUOTE_START_ROW, 6, numRows, 1).getValues();

  var leftBlock = [];
  var rightBlock = [];

  for (var i = 0; i < numRows; i++) {
    leftBlock.push([
      typeValues[i][0],
      partNumberValues[i][0],
      descriptionValues[i][0],
      manufacturerValues[i][0],
      quantityValues[i][0]
    ]);

    rightBlock.push([
      costPerUnitValues[i][0],
      costPerUnitMarginValues[i][0],
      totalCostValues[i][0],
      totalValues[i][0]
    ]);
  }

  // --- Clear old data ONLY on first quote for this tracker (startRow == 4) ---
  if (startRow === 4) {
    var maxRows = trackerSheet.getMaxRows();
    var clearNumRows = Math.max(0, maxRows - 3);

    if (clearNumRows > 0) {
      trackerSheet.getRange(4, 2, clearNumRows, 1).clearContent();  // B
      trackerSheet.getRange(4, 3, clearNumRows, 5).clearContent();  // C:G
      trackerSheet.getRange(4, 14, clearNumRows, 4).clearContent(); // N:Q
    }
  }

  // --- Write new data starting at startRow ---
  trackerSheet.getRange(startRow, 2, sourceBlock.length, 1).setFormulas(sourceBlock); // B
  trackerSheet.getRange(startRow, 3, leftBlock.length, 5).setValues(leftBlock);        // C:G
  trackerSheet.getRange(startRow, 14, rightBlock.length, 4).setValues(rightBlock);     // N:Q

  var nextRow = startRow + numRows;

  Logger.log(
    'Synced ' + numRows + ' rows from quote "' + (quoteLabel || quoteSS.getName()) +
    '" using source tab "' + sourceTabName +
    '" into tracker "' + trackerSheetName +
    '" (rows ' + startRow + '–' + (nextRow - 1) + ').'
  );

  return {
    nextRow: nextRow,
    sourceTabName: sourceTabName,
    numRows: numRows,
    isApprovedLocked: isApprovedLocked
  };
}


/**
 * Returns true only when:
 * - the first tab name contains "-Approved"
 * - and the sheet is protected/locked
 */
function isApprovedLockedSheet_(sheet) {
  var name = String(sheet.getName() || '');
  var hasApprovedInName = /\-approved\b/i.test(name) || /approved/i.test(name);

  if (!hasApprovedInName) {
    return false;
  }

  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  var isProtected = protections && protections.length > 0;

  return isProtected;
}


/**
 * Sets tracker A3 to the project-level closing value.
 */
function setTrackerProjectClosing_(trackerSheetName, closingValue) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(trackerSheetName);
  if (!sheet) {
    throw new Error('Tracker sheet "' + trackerSheetName + '" not found.');
  }

  var cell = sheet.getRange('A3');
  var rule = getProbabilityDropdownRule_();

  cell.setDataValidation(rule);
  cell.setValue(closingValue || 'Probability of close 85%');
}


/**
 * Find last "normal" row based on background color in one column.
 * Stops before the first row where the color changes.
 */
function findEndRowByBackground_(sheet, startRow, lastRow, column) {
  var numRows = lastRow - startRow + 1;
  var bgRange = sheet.getRange(startRow, column, numRows, 1);
  var backgrounds = bgRange.getBackgrounds();

  var baseColor = backgrounds[0][0];
  var endRow = lastRow;

  for (var i = 1; i < backgrounds.length; i++) {
    var color = backgrounds[i][0];
    if (color !== baseColor) {
      endRow = startRow + i - 1;
      break;
    }
  }
  return endRow;
}


/**
 * Find the row just before the one containing "SubT Check".
 */
function findEndRowBySubtCheck_(sheet, startRow, lastRow) {
  var finder = sheet.createTextFinder('SubT Check').matchCase(false);
  var matches = finder.findAll();
  var bestRow = null;

  matches.forEach(function(range) {
    var r = range.getRow();
    if (r >= startRow && r <= lastRow) {
      if (bestRow === null || r < bestRow) {
        bestRow = r;
      }
    }
  });

  if (bestRow === null) {
    return startRow - 1;
  }
  return bestRow - 1;
}


/**
 * Find the last data row before a run of N blank rows in BOTH A and B.
 */
function findEndRowByBlankStreak_(sheet, startRow, lastRow, blankStreakThreshold) {
  var numRows = lastRow - startRow + 1;
  if (numRows <= 0) {
    return startRow - 1;
  }

  var range = sheet.getRange(startRow, 1, numRows, 2);
  var values = range.getValues();

  var consecutiveBlanks = 0;
  var lastDataRow = startRow - 1;

  for (var i = 0; i < values.length; i++) {
    var a = values[i][0];
    var b = values[i][1];

    var isBlank =
      String(a).trim() === '' &&
      String(b).trim() === '';

    if (isBlank) {
      consecutiveBlanks++;
      if (consecutiveBlanks >= blankStreakThreshold) {
        return lastDataRow;
      }
    } else {
      consecutiveBlanks = 0;
      lastDataRow = startRow + i;
    }
  }

  return lastDataRow;
}


// Central helper: get the "Probability of close" dropdown rule.
function getProbabilityDropdownRule_() {
  var ss = SpreadsheetApp.getActive();
  var configSheet = ss.getSheetByName('Tracker Config') || ss.getSheetByName('Tracker config');
  if (configSheet) {
    var tplRule = configSheet.getRange('F1').getDataValidation();
    if (tplRule) {
      return tplRule;
    }
  }

  return SpreadsheetApp.newDataValidation()
    .requireValueInList(
      [
        'Probability of close 100%',
        'Probability of close 85%',
        'Probability of close < 85%'
      ],
      true
    )
    .setAllowInvalid(false)
    .build();
}


/**
 * Ensure rows 2–3 on a tracker are set up.
 */
function ensureTopProbabilitySpacer_(trackerSheetName) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(trackerSheetName);
  if (!sheet) {
    Logger.log('ensureTopProbabilitySpacer_: tracker sheet "' + trackerSheetName + '" not found.');
    return;
  }

  var bandRange = sheet.getRange(2, 1, 1, 17); // A2:Q2
  bandRange.setBackground('#e0e0e0');

  var labelCell = bandRange.getCell(1, 1); // A2
  if (!String(labelCell.getValue()).trim()) {
    labelCell.setValue('Probability of close');
  }

  var dropdownCell = sheet.getRange(3, 1);
  dropdownCell.setBackground(null);

  var rule = getProbabilityDropdownRule_();
  dropdownCell.setDataValidation(rule);

  // Default blank trackers to 85%
  if (!String(dropdownCell.getValue()).trim()) {
    dropdownCell.setValue('Probability of close 85%');
  }
}


/**
 * Insert spacer rows after a quote block on a tracker.
 */
function insertQuoteSpacerRows_(trackerSheetName, firstEmptyRow) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(trackerSheetName);
  if (!sheet) {
    Logger.log('insertQuoteSpacerRows_: tracker sheet "' + trackerSheetName + '" not found.');
    return;
  }

  var bandRow = firstEmptyRow;
  var dropdownRow = firstEmptyRow + 1;

  var bandRange = sheet.getRange(bandRow, 1, 1, 17);
  bandRange.clearContent();
  bandRange.setBackground('#e0e0e0');

  var cell = sheet.getRange(dropdownRow, 1);
  cell.clearContent();
  cell.setBackground(null);

  var rule = getProbabilityDropdownRule_();
  cell.setDataValidation(rule);
}


/**
 * Remove old quote-level spacer rows, but leave top project-level dropdown in A3 alone.
 */
function clearQuoteSpacerRows_(trackerSheetName) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(trackerSheetName);
  if (!sheet) {
    Logger.log('clearQuoteSpacerRows_: tracker sheet "' + trackerSheetName + '" not found.');
    return;
  }

  var maxRows = sheet.getMaxRows();
  if (maxRows <= 4) return;

  var startRow = 4;
  var numRows = maxRows - startRow + 1;
  if (numRows <= 0) return;

  var colARange = sheet.getRange(startRow, 1, numRows, 1);
  var validations = colARange.getDataValidations();

  var PROB_OPTIONS = [
    'Probability of close 100%',
    'Probability of close 85%',
    'Probability of close < 85%'
  ];

  for (var i = 0; i < numRows; i++) {
    var rule = validations[i][0];
    if (!rule) continue;

    if (rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      var critVals = rule.getCriteriaValues();
      if (!critVals || !critVals[0]) continue;
      var list = critVals[0];

      var matchesProb = PROB_OPTIONS.every(function(opt) {
        return list.indexOf(opt) !== -1;
      });
      if (!matchesProb) continue;

      var dropdownRow = startRow + i;
      Logger.log('Clearing old probability spacer at row ' + dropdownRow + ' on "' + trackerSheetName + '"');

      sheet.getRange(dropdownRow, 1, 1, 17)
        .clearContent()
        .setDataValidation(null)
        .setBackground(null);

      if (dropdownRow > 1) {
        sheet.getRange(dropdownRow - 1, 1, 1, 17)
          .setBackground(null)
          .clearContent();
      }
    }
  }
}