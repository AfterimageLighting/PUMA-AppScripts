/************************************************************
 * RAW PO IMPORT -> Project Tracker Updater
 ************************************************************/

var PO_IMPORT_CONFIG = {
  RAW_SHEET_NAME: 'RAW_PO_IMPORT',
  UNMATCHED_SHEET_NAME: 'PO Unmatched',
  TRACKER_SUFFIX: ' - Project Tracker',

  // Limit PO PDF lookup to an environment-appropriate folder tree.
  LIVE_PO_ROOT_FOLDER_ID: '1VlCypDA_iF5dEUmA9c3E7ABYyS4-m6W2',
  TEST_PO_ROOT_FOLDER_ID: '1EeexxItqTX896tq64lf1-8lBz2bqCBmH',

  RAW_HEADERS: {
    project: 'project_raw',
    poNumber: 'po_number',
    itemName: 'item_name',
    itemType: 'item_type',
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

  // The PO importer owns the transition INTO Ordered. It must never move a
  // later/special workflow state backward to Ordered.
  ORDERABLE_STATUSES: {
    '': true,
    'Unapproved': true,
    'Approved': true,
    'To Be Ordered': true,
    'Ordered': true
  },

  NON_PRODUCT_KEYWORDS: [
    'freight',
    'shipping',
    'delivery',
    'tax',
    'handling',
    'labor',
    'service charge',
    'discount'
  ],

  PUMA_LINE_ID_HEADER: 'PUMA_LINE_ID',
  PUMA_REVIEW_FLAG_HEADER: 'PUMA_REVIEW_FLAG',
  PUMA_SOURCE_TYPE_HEADER: 'PUMA_SOURCE_TYPE',

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

  // Central project resolver + independent tracker evidence. These are built
  // once per run so every PO line uses the same authoritative project registry.
  var projectRegistry = pumaBuildProjectRegistry_(ss);
  var trackerEvidenceIndex = pumaBuildTrackerEvidenceIndex_(ss, projectRegistry);

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
      itemType: String(getCellByHeader_(row, rawHeaderMap, PO_IMPORT_CONFIG.RAW_HEADERS.itemType) || '').trim(),
      description: String(getCellByHeader_(row, rawHeaderMap, PO_IMPORT_CONFIG.RAW_HEADERS.description) || '').trim(),
      qty: getCellByHeader_(row, rawHeaderMap, PO_IMPORT_CONFIG.RAW_HEADERS.qty),
      unitCost: getCellByHeader_(row, rawHeaderMap, PO_IMPORT_CONFIG.RAW_HEADERS.unitCost),
      vendor: String(getCellByHeader_(row, rawHeaderMap, PO_IMPORT_CONFIG.RAW_HEADERS.vendor) || '').trim()
    };

    var trackerOverride = String(
      rawSheet.getRange(
        rawRecord.sheetRow,
        PO_IMPORT_CONFIG.RAW_TRACKER_OVERRIDE_COLUMN
      ).getDisplayValue() || ''
    ).trim();

    rawRecord.rawTrackerOverride = trackerOverride;

    // Ignore charges that are not physical tracker items. Keep them in RAW for
    // audit/history, but never allow them to match a fixture row.
    if (isNonProductPoLine_(rawRecord)) {
      stampProcessedRawRow_(rawSheet, rawRecord.sheetRow, 'Ignored - Non-product charge');
      continue;
    }

    // Manual override is evaluated FIRST, even when project_raw is blank.
    var projectEvidence = pumaGatherRawPoEvidence_(
      rawRecord,
      projectRegistry,
      trackerEvidenceIndex
    );

    var projectResolution = pumaResolveProjectWithRegistry_(
      projectRegistry,
      rawRecord.projectRaw,
      projectEvidence
    );

    if (projectResolution.status !== 'CONFIRMED') {
      var resolutionReason =
        'Project resolution ' + projectResolution.status +
        (projectResolution.method ? ' (' + projectResolution.method + ')' : '');

      unmatched.push(makeUnmatchedRow_(rawRecord, resolutionReason));
      stampProcessedRawRow_(rawSheet, rawRecord.sheetRow, 'Unmatched - ' + resolutionReason);
      continue;
    }

    rawRecord.resolvedProject = projectResolution.canonicalProject;
    rawRecord.resolvedTrackerSheetName = projectResolution.trackerSheetName;

    var trackerSheetName = projectResolution.trackerSheetName;
    var trackerSheet = ss.getSheetByName(trackerSheetName);

    if (!trackerSheet) {
      var reason = 'Resolved tracker sheet not found: ' + trackerSheetName;
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

    var existingPoLine = trackerAlreadyHasPoLine_(trackerSheet, trackerInfo, rawRecord);
    if (existingPoLine && existingPoLine.found) {
      stampProcessedRawRow_(
        rawSheet,
        rawRecord.sheetRow,
        'Already in tracker - Row ' + existingPoLine.row
      );
      continue;
    }

    var match = findBestTrackerMatch_(trackerInfo, rawRecord, usedRowsBySheet[trackerSheetName]);

    if (match && match.ambiguous) {
      var ambiguousReason = 'Ambiguous tracker item match: ' + match.candidateRows.join(', ');
      unmatched.push(makeUnmatchedRow_(rawRecord, ambiguousReason));
      stampProcessedRawRow_(rawSheet, rawRecord.sheetRow, 'Unmatched - ' + ambiguousReason);
      continue;
    }

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
 * Deterministic tracker-item matching.
 *
 * Strongest:
 *   exact Type + exact Part Number
 *   exact Part Number
 *   exact Type + family Part Number
 *   family Part Number
 *   exact Type + exact Description
 *   exact Type
 *   exact Description
 *
 * Quantity is a tie-breaker only. If multiple rows remain tied for the best
 * score, this function returns ambiguous instead of guessing.
 */
function findBestTrackerMatch_(trackerInfo, rawRecord, usedRowsMap) {
  if (!trackerInfo || !trackerInfo.rows || !trackerInfo.rows.length) {
    return null;
  }

  usedRowsMap = usedRowsMap || {};
  var candidates = [];

  var rawType = normalizeToken_(rawRecord.itemType);
  var rawPart = normalizePartNumber_(rawRecord.itemName);
  var rawDesc = normalizeToken_(rawRecord.description);
  var rawQty = normalizeComparableNumber_(rawRecord.qty);

  for (var i = 0; i < trackerInfo.rows.length; i++) {
    var tr = trackerInfo.rows[i];
    if (usedRowsMap[tr.sheetRow]) continue;

    var trackerType = normalizeToken_(tr.type);
    var trackerPart = normalizePartNumber_(tr.partNumber);
    var trackerDesc = normalizeToken_(tr.description);
    var trackerQty = normalizeComparableNumber_(tr.quantity);

    var exactType = !!(trackerType && rawType && trackerType === rawType);
    var exactPart = !!(trackerPart && rawPart && trackerPart === rawPart);
    var familyPart = !!(
      trackerPart && rawPart &&
      !exactPart &&
      partNumbersAreFamilyMatch_(tr.partNumber, rawRecord.itemName)
    );
    var exactDesc = !!(trackerDesc && rawDesc && trackerDesc === rawDesc);

    var score = 0;
    var method = '';

    if (exactType && exactPart) {
      score = 130;
      method = 'TYPE+PART';
    } else if (exactPart) {
      score = 110;
      method = 'PART';
    } else if (exactType && familyPart) {
      score = 100;
      method = 'TYPE+PART_FAMILY';
    } else if (familyPart) {
      score = 90;
      method = 'PART_FAMILY';
    } else if (exactType && exactDesc) {
      score = 80;
      method = 'TYPE+DESCRIPTION';
    } else if (exactType) {
      score = 75;
      method = 'TYPE';
    } else if (exactDesc) {
      score = 60;
      method = 'DESCRIPTION';
    }

    if (!score) continue;

    // Quantity only breaks an otherwise equivalent match. Partial ordering is
    // common, so quantity mismatch must not invalidate a strong Type/Part match.
    if (rawQty !== '' && trackerQty !== '' && rawQty === trackerQty) score += 2;

    candidates.push({
      trackerRow: tr,
      score: score,
      method: method
    });
  }

  if (!candidates.length) return null;

  candidates.sort(function(a, b) {
    if (b.score !== a.score) return b.score - a.score;
    return a.trackerRow.sheetRow - b.trackerRow.sheetRow;
  });

  var bestScore = candidates[0].score;
  var tied = candidates.filter(function(x) { return x.score === bestScore; });

  if (tied.length > 1) {
    return {
      ambiguous: true,
      score: bestScore,
      candidateRows: tied.map(function(x) { return x.trackerRow.sheetRow; }),
      candidates: tied
    };
  }

  return candidates[0];
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

    // PO import may advance early purchasing states to Ordered, but may never
    // regress Received/Scheduled/Delivered/Omitted/Note or any unknown state.
    if (canPoImporterSetOrdered_(currentStatus) && currentStatus !== 'Ordered') {
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
  var lineIdCol = getColNum_(headerMap, PO_IMPORT_CONFIG.PUMA_LINE_ID_HEADER);
  var reviewFlagCol = getColNum_(headerMap, PO_IMPORT_CONFIG.PUMA_REVIEW_FLAG_HEADER);
  var sourceTypeCol = getColNum_(headerMap, PO_IMPORT_CONFIG.PUMA_SOURCE_TYPE_HEADER);

  var insertAfterRow = findLastRealTrackerDataRow_(sheet, trackerInfo);
  sheet.insertRowAfter(insertAfterRow);
  var targetRow = insertAfterRow + 1;
  var lastCol = Math.max(sheet.getLastColumn(), 1);

  if (insertAfterRow >= 1) {
    sheet.getRange(insertAfterRow, 1, 1, lastCol).copyTo(
      sheet.getRange(targetRow, 1, 1, lastCol),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );
  }

  var rowValues = new Array(lastCol).fill('');
  setRowValueByCol_(
    rowValues,
    projectCol,
    rawRecord.resolvedProject || normalizeProjectName_(rawRecord.projectRaw)
  );
  setRowValueByCol_(rowValues, sourceCol, 'PO Import - Unmatched');
  setRowValueByCol_(rowValues, typeCol, rawRecord.itemType || 'PO Import Item');
  setRowValueByCol_(rowValues, partCol, rawRecord.itemName);
  setRowValueByCol_(rowValues, descCol, rawRecord.description);
  setRowValueByCol_(rowValues, manufacturerCol, rawRecord.vendor);
  setRowValueByCol_(rowValues, qtyCol, rawRecord.qty);
  setRowValueByCol_(rowValues, statusCol, 'Ordered');
  setRowValueByCol_(rowValues, costCol, rawRecord.unitCost);
  setRowValueByCol_(rowValues, lineIdCol, generatePumaLineId_());
  setRowValueByCol_(rowValues, reviewFlagCol, '');
  setRowValueByCol_(rowValues, sourceTypeCol, 'PO_ONLY');

  var targetRange = sheet.getRange(targetRow, 1, 1, lastCol);
  targetRange.clearDataValidations();
  targetRange.setValues([rowValues]);

  if (poCol && rawRecord.poNumber) {
    var poUrl = poPdfMap[normalizePoKey_(rawRecord.poNumber)] || '';
    if (!poUrl) poPdfFound = false;
    setPoNumberRichLink_(sheet.getRange(targetRow, poCol), rawRecord.poNumber, poUrl);
  }

  if (sourceCol) sheet.getRange(targetRow, sourceCol).setNote('This line came from a PO but did not match an existing quoted tracker line.');

  trackerInfo.rows.push({
    sheetRow: targetRow,
    type: String(rawRecord.itemType || 'PO Import Item'),
    partNumber: String(rawRecord.itemName || ''),
    description: String(rawRecord.description || ''),
    status: 'Ordered',
    quantity: rawRecord.qty
  });

  return { success: true, poPdfFound: poPdfFound, row: targetRow };
}


function canPoImporterSetOrdered_(status) {
  status = String(status || '').trim();
  return !!PO_IMPORT_CONFIG.ORDERABLE_STATUSES[status];
}


function isNonProductPoLine_(rawRecord) {
  var haystack = [
    rawRecord && rawRecord.itemName,
    rawRecord && rawRecord.itemType,
    rawRecord && rawRecord.description
  ].map(function(v) {
    return String(v || '').toLowerCase();
  }).join(' ');

  if (!haystack.trim()) return false;

  for (var i = 0; i < PO_IMPORT_CONFIG.NON_PRODUCT_KEYWORDS.length; i++) {
    var term = PO_IMPORT_CONFIG.NON_PRODUCT_KEYWORDS[i];
    if (haystack.indexOf(term) !== -1) return true;
  }

  return false;
}


function generatePumaLineId_() {
  // Stable-enough unique ID for a newly appended PO-only tracker row.
  return 'PUMA-LINE-' + Utilities.getUuid().replace(/-/g, '').substring(0, 8);
}


function setRowValueByCol_(rowValues, colNum, value) {
  if (!colNum) return;
  if (value === '' || value == null) return;
  rowValues[colNum - 1] = value;
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
function getPoPdfRootFolderId_() {
  if (typeof pumaIsTestWorkbook_ === 'function' && pumaIsTestWorkbook_()) {
    return PO_IMPORT_CONFIG.TEST_PO_ROOT_FOLDER_ID;
  }
  return PO_IMPORT_CONFIG.LIVE_PO_ROOT_FOLDER_ID;
}

function buildPoPdfMap_() {
  var map = {};
  var rootFolderId = getPoPdfRootFolderId_();

  if (rootFolderId) {
    var root = DriveApp.getFolderById(rootFolderId);
    indexPoFilesInFolderRecursive_(root, map);
    return map;
  }

  // Fail closed: unrestricted Drive-wide PDF search is intentionally disabled.
  throw new Error('PO PDF root folder is not configured.');
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
    'item_type',
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
    rawRecord.itemType,
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

function trackerAlreadyHasPoLine_(sheet, trackerInfo, rawRecord) {
  if (!rawRecord || !rawRecord.poNumber) return false;
  var headerMap = trackerInfo.headerMap;
  var poCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.poNumber);
  if (!poCol) return false;
  var partCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.partNumber);
  var descCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.description);
  var qtyCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.quantity);
  var costCol = getColNum_(headerMap, PO_IMPORT_CONFIG.TRACKER_HEADERS.costPerUnit);
  var targetPo = normalizePoKey_(rawRecord.poNumber);
  var targetPart = normalizePartNumber_(rawRecord.itemName);
  var targetDesc = normalizeToken_(rawRecord.description);
  var targetQty = normalizeComparableNumber_(rawRecord.qty);
  var targetCost = normalizeComparableNumber_(rawRecord.unitCost);
  var data = sheet.getDataRange().getDisplayValues();
  for (var r = trackerInfo.headerRowIndex + 1; r < data.length; r++) {
    var row = data[r];
    if (normalizePoKey_(row[poCol - 1]) !== targetPo) continue;
    var existingPart = partCol ? normalizePartNumber_(row[partCol - 1]) : '';
    var existingDesc = descCol ? normalizeToken_(row[descCol - 1]) : '';
    var existingQty = qtyCol ? normalizeComparableNumber_(row[qtyCol - 1]) : '';
    var existingCost = costCol ? normalizeComparableNumber_(row[costCol - 1]) : '';
    var partMatches = targetPart && existingPart && targetPart === existingPart;
    var descMatches = targetDesc && existingDesc && (existingDesc === targetDesc || existingDesc.indexOf(targetDesc) !== -1);
    var qtyMatches = targetQty === '' || existingQty === '' || targetQty === existingQty;
    var costMatches = targetCost === '' || existingCost === '' || targetCost === existingCost;
    if ((partMatches || descMatches) && qtyMatches && costMatches) return {found:true,row:r+1};
  }
  return false;
}


function normalizeComparableNumber_(value) {
  if (value === '' || value == null) return '';
  var normalized = String(value).replace(/[$,\s]/g, '').trim();
  if (!normalized) return '';
  var numberValue = Number(normalized);
  return isNaN(numberValue) ? normalized : String(Math.round(numberValue * 10000) / 10000);
}

function installDailyPoImportTrigger() {
  removeDailyPoImportTriggers_();
  ScriptApp.newTrigger('applyRawPoImportToProjectTrackers').timeBased().everyDays(1).atHour(6).create();
  SpreadsheetApp.getUi().alert('Daily PO Import trigger installed. It will run once per day around 6:00 AM.');
}

function removeDailyPoImportTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'applyRawPoImportToProjectTrackers') ScriptApp.deleteTrigger(triggers[i]);
  }
}

function removeDailyPoImportTrigger() {
  removeDailyPoImportTriggers_();
  SpreadsheetApp.getUi().alert('Daily PO Import trigger removed.');
}
