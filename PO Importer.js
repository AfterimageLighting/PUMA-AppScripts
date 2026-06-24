/************************************************************
 * RAW PO IMPORT -> Project Tracker Updater
 ************************************************************/

var PO_IMPORT_CONFIG = {
  RAW_SHEET_NAME: 'RAW_PO_IMPORT',
  UNMATCHED_SHEET_NAME: 'PO Unmatched',
  TRACKER_SUFFIX: ' - Project Tracker',

  // Limit PO PDF lookup to this folder tree.
  PO_ROOT_FOLDER_ID: '1VlCypDA_iF5dEUmA9c3E7ABYyS4-m6W2',

  RAW_HEADERS: {
    project: 'project_raw',
    poNumber: 'po_number',
    itemName: 'item_name',
    description: 'description',
    qty: 'qty',
    unitCost: 'unit_cost',
    vendor: 'vendor'
  },

  TRACKER_HEADERS: {
    project: 'Project',
    source: 'Source',
    type: 'Type',
    partNumber: 'Part Number',
    description: 'Description',
    manufacturer: 'Manufacturer',
    quantity: 'Quantity',
    status: 'Status',
    poNumber: 'PO Number',
    costPerUnit: 'Cost Per Unit'
  },

  TRACKER_OPTIONAL_HEADERS: ['Project', 'Source', 'Description', 'Manufacturer'],

  LOCKED_STATUSES: {
    'Received': true,
    'Delivered': true
  },

  RAW_PROCESSED_COLUMN: 17, // Q
  RAW_ACTION_COLUMN: 18,    // R
  RAW_TRACKER_OVERRIDE_COLUMN: 19, // S
  TIMESTAMP_FORMAT: 'M/d/yyyy h:mm:ss a'
};


/**
 * Main runner
 */
function applyRawPoImportToProjectTrackers() {
  var ss = SpreadsheetApp.getActive();
  var rawSheet = ss.getSheetByName(PO_IMPORT_CONFIG.RAW_SHEET_NAME);
  if (!rawSheet) {
    throw new Error('RAW sheet not found: ' + PO_IMPORT_CONFIG.RAW_SHEET_NAME);
  }

  var rawData = rawSheet.getDataRange().getValues();
  if (rawData.length < 2) {
    Logger.log('No RAW PO data found.');
    return;
  }

  var rawHeaderMap = makeHeaderMap_(rawData[0]);
  validateHeaders_(rawHeaderMap, PO_IMPORT_CONFIG.RAW_HEADERS, 'RAW_PO_IMPORT');

  var poPdfMap = buildPoPdfMap_();
  var trackerCache = {};
  var usedRowsBySheet = {};
  var unmatched = [];

  var matchedCount = 0;
  var appendedCount = 0;
  var skippedStampedCount = 0;

  for (var i = 1; i < rawData.length; i++) {
    var row = rawData[i];
    if (isBlankRow_(row)) continue;

    var processedVal = rawSheet.getRange(i + 1, PO_IMPORT_CONFIG.RAW_PROCESSED_COLUMN).getValue();
    if (processedVal !== '' && processedVal != null) {
      skippedStampedCount++;
      continue;
    }

    var rawRecord = {
      sheetRow: i + 1,
      projectRaw: getCellByHeader_(row, rawHeaderMap, PO_IMPORT_CONFIG.RAW_HEADERS.project),
      poNumber: String(getCellByHeader_(row, rawHeaderMap, PO_IMPORT_CONFIG.RAW_HEADERS.poNumber) || '').trim(),
      itemName: String(getCellByHeader_(row, rawHeaderMap, PO_IMPORT_CONFIG.RAW_HEADERS.itemName) || '').trim(),
      description: String(getCellByHeader_(row, rawHeaderMap, PO_IMPORT_CONFIG.RAW_HEADERS.description) || '').trim(),
      qty: getCellByHeader_(row, rawHeaderMap, PO_IMPORT_CONFIG.RAW_HEADERS.qty),
      unitCost: getCellByHeader_(row, rawHeaderMap, PO_IMPORT_CONFIG.RAW_HEADERS.unitCost),
      vendor: String(getCellByHeader_(row, rawHeaderMap, PO_IMPORT_CONFIG.RAW_HEADERS.vendor) || '').trim()
    };

    var normalizedProject = normalizeProjectName_(rawRecord.projectRaw);

    if (!normalizedProject) {
      var reason = 'Could not normalize project name';
      unmatched.push(makeUnmatchedRow_(rawRecord, reason));
      stampProcessedRawRow_(rawSheet, rawRecord.sheetRow, 'Unmatched - ' + reason);
      continue;
    }

    var trackerOverride = String(
      rawSheet.getRange(
      rawRecord.sheetRow,
      PO_IMPORT_CONFIG.RAW_TRACKER_OVERRIDE_COLUMN
      ).getDisplayValue() || ''
    ).trim();

    var trackerOverride = String(
     rawSheet.getRange(
       rawRecord.sheetRow,
       PO_IMPORT_CONFIG.RAW_TRACKER_OVERRIDE_COLUMN
      ).getDisplayValue() || ''
    ).trim();

    var trackerSheetName = trackerOverride || (normalizedProject + PO_IMPORT_CONFIG.TRACKER_SUFFIX);
    var trackerSheet = ss.getSheetByName(trackerSheetName);

    if (!trackerSheet) {
     var reason = 'Tracker sheet not found: ' + trackerSheetName;
     unmatched.push(makeUnmatchedRow_(rawRecord, reason));
     stampProcessedRawRow_(rawSheet, rawRecord.sheetRow, 'Unmatched - ' + reason);
     continue;
    }

    if (!usedRowsBySheet[trackerSheetName]) {
      usedRowsBySheet[trackerSheetName] = {};
    }

    var trackerInfo = trackerCache[trackerSheetName];
    if (!trackerInfo || !trackerInfo.rows) {
      try {
        trackerInfo = loadTrackerSheet_(trackerSheet);
        trackerCache[trackerSheetName] = trackerInfo;
      } catch (e) {
        var reason = 'Could not load tracker sheet: ' + e.message;
        unmatched.push(makeUnmatchedRow_(rawRecord, reason));
        stampProcessedRawRow_(rawSheet, rawRecord.sheetRow, 'Unmatched - ' + reason);
        continue;
      }
    }

    var existingPo = trackerAlreadyHasPo_(trackerSheet, trackerInfo, rawRecord.poNumber);
    if (existingPo && existingPo.found) {
      stampProcessedRawRow_(
       rawSheet,
       rawRecord.sheetRow,
       'Already in tracker - Row ' + existingPo.row
      );
      continue;
    }

    var match = findBestTrackerMatch_(trackerInfo, rawRecord, usedRowsBySheet[trackerSheetName]);

    if (match) {
      try {
        var matchResult = applyMatchToTrackerRow_(trackerSheet, trackerInfo, match, rawRecord, poPdfMap);
        usedRowsBySheet[trackerSheetName][match.trackerRow.sheetRow] = true;

        var matchAction = (matchResult && matchResult.poPdfFound === false)
          ? 'Matched - PO PDF Missing'
          : 'Matched';

        stampProcessedRawRow_(rawSheet, rawRecord.sheetRow, matchAction);
        matchedCount++;
        continue;
      } catch (e1) {
        unmatched.push(makeUnmatchedRow_(rawRecord, 'Match update failed: ' + e1.message));
        continue;
      }
    }

    try {
      var appendResult = appendRawRecordToTracker_(trackerSheet, trackerInfo, rawRecord, poPdfMap);
      if (appendResult && appendResult.success) {
        var appendAction = (appendResult.poPdfFound === false)
          ? 'Appended - PO PDF Missing'
          : 'Appended';

        stampProcessedRawRow_(rawSheet, rawRecord.sheetRow, appendAction);
        appendedCount++;
        continue;
      }
    } catch (e2) {
      unmatched.push(makeUnmatchedRow_(rawRecord, 'Append failed: ' + e2.message));
      continue;
    }

    var finalReason = 'Project found, but no tracker match and append failed';
    unmatched.push(makeUnmatchedRow_(rawRecord, finalReason));
    stampProcessedRawRow_(rawSheet, rawRecord.sheetRow, 'Unmatched - ' + finalReason);
  }

  writeUnmatchedLog_(ss, unmatched);

  Logger.log('PO Import update complete.');
  Logger.log('Matched existing rows: ' + matchedCount);
  Logger.log('Appended new rows: ' + appendedCount);
  Logger.log('Skipped already stamped rows: ' + skippedStampedCount);
  Logger.log('Unmatched rows: ' + unmatched.length);
}


/**
 * Menu
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PUMA')
    .addItem('Apply RAW PO Import', 'applyRawPoImportToProjectTrackers')
    .addToUi();
}


/**
 * Load tracker sheet into memory
 */
function loadTrackerSheet_(sheet) {
  var data = sheet.getDataRange().getValues();
  if (!data.length) {
    throw new Error('Tracker sheet is empty: ' + sheet.getName());
  }

  var headerRowIndex = findTrackerHeaderRowIndex_(data);
  if (headerRowIndex === -1) {
    throw new Error('Could not find tracker header row in sheet: ' + sheet.getName());
  }

  var headers = data[headerRowIndex];
  var headerMap = makeHeaderMap_(headers);

  validateHeaders_(
    headerMap,
    PO_IMPORT_CONFIG.TRACKER_HEADERS,
    sheet.getName(),
    PO_IMPORT_CONFIG.TRACKER_OPTIONAL_HEADERS
  );

  var rows = [];
  for (var r = headerRowIndex + 1; r < data.length; r++) {
    var row = data[r];
    if (isBlankRow_(row)) continue;

    var typeVal = String(getCellByHeader_(row, headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.type) || '').trim();
    var partVal = String(getCellByHeader_(row, headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.partNumber) || '').trim();
    var descVal = String(getCellByHeader_(row, headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.description) || '').trim();
    var statusVal = String(getCellByHeader_(row, headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.status) || '').trim();
    var qtyVal = getCellByHeader_(row, headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.quantity);

    if (!typeVal && !partVal && !descVal) continue;

    rows.push({
      sheetRow: r + 1,
      type: typeVal,
      partNumber: partVal,
      description: descVal,
      status: statusVal,
      quantity: qtyVal
    });
  }

  return {
    headerRowIndex: headerRowIndex,
    headerMap: headerMap,
    rows: rows
  };
}


/**
 * Find header row in first 10 rows
 */
function findTrackerHeaderRowIndex_(data) {
  var required = [
    PO_IMPORT_CONFIG.TRACKER_HEADERS.source,
    PO_IMPORT_CONFIG.TRACKER_HEADERS.type,
    PO_IMPORT_CONFIG.TRACKER_HEADERS.partNumber,
    PO_IMPORT_CONFIG.TRACKER_HEADERS.status,
    PO_IMPORT_CONFIG.TRACKER_HEADERS.poNumber
  ];

  for (var r = 0; r < Math.min(data.length, 10); r++) {
    var headerMap = makeHeaderMap_(data[r]);
    var ok = true;

    for (var i = 0; i < required.length; i++) {
      if (headerMap[normalizeHeader_(required[i])] == null) {
        ok = false;
        break;
      }
    }

    if (ok) return r;
  }

  return -1;
}


/**
 * Match priority:
 * 100 = Type exact match to RAW description
 *  90 = exact part number
 *  80 = wildcard/family part match
 *  70 = tracker description exact
 */
function findBestTrackerMatch_(trackerInfo, rawRecord, usedRowsMap) {
  if (!trackerInfo || !trackerInfo.rows || !trackerInfo.rows.length) {
    return null;
  }

  var best = null;
  usedRowsMap = usedRowsMap || {};

  for (var i = 0; i < trackerInfo.rows.length; i++) {
    var tr = trackerInfo.rows[i];
    if (usedRowsMap[tr.sheetRow]) continue;

    var score = 0;

    var trackerType = normalizeToken_(tr.type);
    var rawDesc = normalizeToken_(rawRecord.description);

    var trackerPart = normalizePartNumber_(tr.partNumber);
    var rawPart = normalizePartNumber_(rawRecord.itemName);

    var trackerDesc = normalizeToken_(tr.description);

    if (trackerType && rawDesc && trackerType === rawDesc) {
      score = Math.max(score, 100);
    }

    if (trackerPart && rawPart && trackerPart === rawPart) {
      score = Math.max(score, 90);
    }

    if (trackerPart && rawPart && partNumbersAreFamilyMatch_(tr.partNumber, rawRecord.itemName)) {
      score = Math.max(score, 80);
    }

    if (trackerDesc && rawDesc && trackerDesc === rawDesc) {
      score = Math.max(score, 70);
    }

    if (score > 0) {
      var trackerQty = Number(tr.quantity || 0);
      var rawQty = Number(rawRecord.qty || 0);
      if (trackerQty && rawQty && trackerQty === rawQty) {
        score += 1;
      }
    }

    if (score > 0 && (!best || score > best.score)) {
      best = {
        trackerRow: tr,
        score: score
      };
    }
  }

  return best;
}


/**
 * Update an existing tracker row
 */
function applyMatchToTrackerRow_(sheet, trackerInfo, match, rawRecord, poPdfMap) {
  var headerMap = trackerInfo.headerMap;
  var rowNum = match.trackerRow.sheetRow;
  var poPdfFound = true;

  var partCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.partNumber);
  var descCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.description);
  var qtyCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.quantity);
  var statusCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.status);
  var poCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.poNumber);
  var costCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.costPerUnit);

  if (partCol && rawRecord.itemName) {
    clearCellValidationIfNeeded_(sheet, rowNum, partCol);
    sheet.getRange(rowNum, partCol).setValue(rawRecord.itemName);
  }

  if (descCol && rawRecord.description) {
    var existingDesc = String(sheet.getRange(rowNum, descCol).getDisplayValue() || '').trim();
    if (!existingDesc) {
      clearCellValidationIfNeeded_(sheet, rowNum, descCol);
      sheet.getRange(rowNum, descCol).setValue(rawRecord.description);
    }
  }

  if (qtyCol && rawRecord.qty !== '' && rawRecord.qty != null) {
    var existingQty = sheet.getRange(rowNum, qtyCol).getValue();
    if (existingQty === '' || existingQty == null) {
      clearCellValidationIfNeeded_(sheet, rowNum, qtyCol);
      sheet.getRange(rowNum, qtyCol).setValue(rawRecord.qty);
    }
  }

  if (costCol && rawRecord.unitCost !== '' && rawRecord.unitCost != null) {
    clearCellValidationIfNeeded_(sheet, rowNum, costCol);
    sheet.getRange(rowNum, costCol).setValue(rawRecord.unitCost);
  }

  if (statusCol) {
    var currentStatus = String(sheet.getRange(rowNum, statusCol).getDisplayValue() || '').trim();
    if (!PO_IMPORT_CONFIG.LOCKED_STATUSES[currentStatus]) {
      try {
        sheet.getRange(rowNum, statusCol).setValue('Ordered');
      } catch (e) {
        clearCellValidationIfNeeded_(sheet, rowNum, statusCol);
        sheet.getRange(rowNum, statusCol).setValue('Ordered');
      }
    }
  }

  if (poCol && rawRecord.poNumber) {
    var poUrl = poPdfMap[normalizePoKey_(rawRecord.poNumber)] || '';
    if (!poUrl) poPdfFound = false;
    setPoNumberRichLink_(sheet.getRange(rowNum, poCol), rawRecord.poNumber, poUrl);
  }

  return {
    poPdfFound: poPdfFound
  };
}


/**
 * Append a new tracker row
 */
function appendRawRecordToTracker_(sheet, trackerInfo, rawRecord, poPdfMap) {
  var headerMap = trackerInfo.headerMap;
  var poPdfFound = true;

  var projectCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.project);
  var sourceCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.source);
  var typeCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.type);
  var partCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.partNumber);
  var descCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.description);
  var manufacturerCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.manufacturer);
  var qtyCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.quantity);
  var statusCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.status);
  var poCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.poNumber);
  var costCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.costPerUnit);

  var insertAfterRow = findLastRealTrackerDataRow_(sheet, trackerInfo);
  sheet.insertRowAfter(insertAfterRow);

  var targetRow = insertAfterRow + 1;

  // Clear inherited validations across the new row
  sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).clearDataValidations();

  if (projectCol) {
    sheet.getRange(targetRow, projectCol).setValue(normalizeProjectName_(rawRecord.projectRaw));
  }

  if (sourceCol) {
    sheet.getRange(targetRow, sourceCol).setValue('PO Import');
  }

  if (typeCol && rawRecord.description) {
    sheet.getRange(targetRow, typeCol).setValue(rawRecord.description);
  }

  if (partCol && rawRecord.itemName) {
    sheet.getRange(targetRow, partCol).setValue(rawRecord.itemName);
  }

  if (descCol) {
    var longDesc = buildTrackerDescriptionFromRaw_(rawRecord);
    if (longDesc) {
      sheet.getRange(targetRow, descCol).setValue(longDesc);
    }
  }

  if (manufacturerCol && rawRecord.vendor) {
    sheet.getRange(targetRow, manufacturerCol).setValue(rawRecord.vendor);
  }

  if (qtyCol && rawRecord.qty !== '' && rawRecord.qty != null) {
    sheet.getRange(targetRow, qtyCol).setValue(rawRecord.qty);
  }

  if (statusCol) {
    sheet.getRange(targetRow, statusCol).setValue('Ordered');
  }

  if (costCol && rawRecord.unitCost !== '' && rawRecord.unitCost != null) {
    sheet.getRange(targetRow, costCol).setValue(rawRecord.unitCost);
  }

  if (poCol && rawRecord.poNumber) {
    var poUrl = poPdfMap[normalizePoKey_(rawRecord.poNumber)] || '';
    if (!poUrl) poPdfFound = false;
    setPoNumberRichLink_(sheet.getRange(targetRow, poCol), rawRecord.poNumber, poUrl);
  }

  return {
    success: true,
    poPdfFound: poPdfFound
  };
}


function buildTrackerDescriptionFromRaw_(rawRecord) {
  var parts = [];
  if (rawRecord.description) parts.push(rawRecord.description);
  if (rawRecord.itemName) parts.push(rawRecord.itemName);
  return parts.join(' | ');
}


/**
 * Stamp RAW row
 */
function stampProcessedRawRow_(rawSheet, rowNum, action) {
  var now = new Date();

  var tsCell = rawSheet.getRange(rowNum, PO_IMPORT_CONFIG.RAW_PROCESSED_COLUMN);
  tsCell.setValue(now);
  tsCell.setNumberFormat(PO_IMPORT_CONFIG.TIMESTAMP_FORMAT);

  if (PO_IMPORT_CONFIG.RAW_ACTION_COLUMN) {
    rawSheet.getRange(rowNum, PO_IMPORT_CONFIG.RAW_ACTION_COLUMN).setValue(action || '');
  }
}


/**
 * Write PO hyperlink safely
 */
function setPoNumberRichLink_(range, displayText, url) {
  displayText = String(displayText || '').trim();
  range.clearDataValidations();

  if (!displayText) {
    range.clearContent();
    return;
  }

  if (!url) {
    range.setValue(displayText);
    return;
  }

  var richText = SpreadsheetApp.newRichTextValue()
    .setText(displayText)
    .setLinkUrl(url)
    .build();

  range.setRichTextValue(richText);
}


/**
 * Build PO PDF lookup map
 */
function buildPoPdfMap_() {
  var map = {};

  if (PO_IMPORT_CONFIG.PO_ROOT_FOLDER_ID) {
    var root = DriveApp.getFolderById(PO_IMPORT_CONFIG.PO_ROOT_FOLDER_ID);
    indexPoFilesInFolderRecursive_(root, map);
    return map;
  }

  var files = DriveApp.searchFiles('mimeType = "application/pdf" and trashed = false');
  while (files.hasNext()) {
    var file = files.next();
    var key = extractPoNumberFromFileName_(file.getName());
    if (key && !map[key]) {
      map[key] = file.getUrl();
    }
  }

  return map;
}


function indexPoFilesInFolderRecursive_(folder, map) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    if (String(file.getMimeType()) === 'application/pdf') {
      var key = extractPoNumberFromFileName_(file.getName());
      if (key && !map[key]) {
        map[key] = file.getUrl();
      }
    }
  }

  var subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    indexPoFilesInFolderRecursive_(subfolders.next(), map);
  }
}


function extractPoNumberFromFileName_(name) {
  name = String(name || '').trim();

  var m = name.match(/PO[\s\-_]?(\d+)/i);
  if (m) return normalizePoKey_(m[1]);

  var bare = name.match(/^(\d+)\.pdf$/i);
  if (bare) return normalizePoKey_(bare[1]);

  return '';
}


function normalizePoKey_(poNumber) {
  return String(poNumber || '').replace(/[^\d]/g, '');
}


/**
 * Write unmatched log
 */
function writeUnmatchedLog_(ss, unmatched) {
  var sheet = ss.getSheetByName(PO_IMPORT_CONFIG.UNMATCHED_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PO_IMPORT_CONFIG.UNMATCHED_SHEET_NAME);
  }

  sheet.clearContents();

  var output = [[
    'raw_row',
    'project_raw',
    'normalized_project',
    'po_number',
    'item_name',
    'description',
    'qty',
    'unit_cost',
    'vendor',
    'reason'
  ]];

  for (var i = 0; i < unmatched.length; i++) {
    output.push(unmatched[i]);
  }

  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}


function makeUnmatchedRow_(rawRecord, reason) {
  return [
    rawRecord.sheetRow,
    rawRecord.projectRaw,
    normalizeProjectName_(rawRecord.projectRaw),
    rawRecord.poNumber,
    rawRecord.itemName,
    rawRecord.description,
    rawRecord.qty,
    rawRecord.unitCost,
    rawRecord.vendor,
    reason
  ];
}


/* =========================
 * Helpers
 * ========================= */

function makeHeaderMap_(headerRow) {
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    var key = normalizeHeader_(headerRow[i]);
    if (key) map[key] = i;
  }
  return map;
}

function validateHeaders_(headerMap, expectedObj, contextName, optionalHeaders) {
  optionalHeaders = optionalHeaders || [];

  var optionalLookup = {};
  for (var i = 0; i < optionalHeaders.length; i++) {
    optionalLookup[normalizeHeader_(optionalHeaders[i])] = true;
  }

  var missing = [];
  for (var key in expectedObj) {
    var headerName = expectedObj[key];
    var normalized = normalizeHeader_(headerName);

    if (optionalLookup[normalized]) continue;
    if (headerMap[normalized] == null) {
      missing.push(headerName);
    }
  }

  if (missing.length) {
    throw new Error('Missing required headers in "' + contextName + '": ' + missing.join(', '));
  }
}

function getCellByHeader_(row, headerMap, headerName) {
  var idx = headerMap[normalizeHeader_(headerName)];
  return idx == null ? '' : row[idx];
}

function getColNum_(headerMap, headerName) {
  var idx = headerMap[normalizeHeader_(headerName)];
  return idx == null ? 0 : idx + 1;
}

function normalizeHeader_(value) {
  return String(value || '')
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ' ');
}

function normalizeProjectName_(projectRaw) {
  var text = String(projectRaw || '').trim();
  if (!text) return '';

  if (text.indexOf(':') !== -1) {
    var parts = text.split(':');
    text = parts[parts.length - 1];
  }

  return text.replace(/\s+/g, ' ').trim();
}

function normalizeToken_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[^\w]/g, '');
}

function normalizePartNumber_(value) {
  return String(value || '')
    .toUpperCase()
    .replace(/\s+/g, '')
    .replace(/[^A-Z0-9\-]/g, '');
}

function partNumbersAreFamilyMatch_(trackerPart, actualPart) {
  trackerPart = String(trackerPart || '').toUpperCase().trim();
  actualPart = String(actualPart || '').toUpperCase().trim();

  if (!trackerPart || !actualPart) return false;

  var trackerSeg = trackerPart.split('-');
  var actualSeg = actualPart.split('-');

  if (trackerSeg.length !== actualSeg.length) return false;

  for (var i = 0; i < trackerSeg.length; i++) {
    var t = trackerSeg[i];
    var a = actualSeg[i];

    if (t === 'XX' || t === 'X') continue;
    if (t !== a) return false;
  }

  return true;
}

function isBlankRow_(row) {
  for (var i = 0; i < row.length; i++) {
    if (row[i] !== '' && row[i] != null) return false;
  }
  return true;
}

function findLastRealTrackerDataRow_(sheet, trackerInfo) {
  var data = sheet.getDataRange().getValues();
  var headerMap = trackerInfo.headerMap;
  var startRow = trackerInfo.headerRowIndex + 2;

  var projectIdx = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.project) - 1;
  var sourceIdx = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.source) - 1;
  var typeIdx = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.type) - 1;
  var partIdx = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.partNumber) - 1;
  var descIdx = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.description) - 1;
  var qtyIdx = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.quantity) - 1;
  var statusIdx = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.status) - 1;
  var poIdx = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.poNumber) - 1;
  var costIdx = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.costPerUnit) - 1;

  var lastRealRow = startRow - 1;

  for (var r = startRow - 1; r < data.length; r++) {
    var row = data[r];

    var hasRealContent =
      hasCellValue_(row, projectIdx) ||
      hasCellValue_(row, sourceIdx) ||
      hasCellValue_(row, typeIdx) ||
      hasCellValue_(row, partIdx) ||
      hasCellValue_(row, descIdx) ||
      hasCellValue_(row, qtyIdx) ||
      hasCellValue_(row, statusIdx) ||
      hasCellValue_(row, poIdx) ||
      hasCellValue_(row, costIdx);

    if (hasRealContent) {
      lastRealRow = r + 1;
    }
  }

  return Math.max(lastRealRow, startRow - 1);
}

function hasCellValue_(row, idx) {
  if (idx < 0 || idx >= row.length) return false;
  return row[idx] !== '' && row[idx] != null;
}

function clearCellValidationIfNeeded_(sheet, row, col) {
  if (!col) return;
  sheet.getRange(row, col).clearDataValidations();
}

function trackerAlreadyHasPo_(sheet, trackerInfo, poNumber) {
  if (!poNumber) return false;

  var poCol = getColNum_(trackerInfo.headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.poNumber);
  if (!poCol) return false;

  var targetKey = normalizePoKey_(poNumber);
  var data = sheet.getDataRange().getDisplayValues();

  for (var r = trackerInfo.headerRowIndex + 1; r < data.length; r++) {
    var existingPo = data[r][poCol - 1];
    if (normalizePoKey_(existingPo) === targetKey) {
      return {
        found: true,
        row: r + 1
      };
    }
  }

  return false;
}
