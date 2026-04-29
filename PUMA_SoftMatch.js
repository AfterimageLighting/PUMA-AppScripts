/**
 * PUMA SOFT MATCH ENGINE v1
 *
 * Purpose:
 * Compare existing tracker rows against newly imported quote rows using tiered matching.
 *
 * This does NOT write changes to project trackers yet.
 * It only builds a test report sheet: PUMA_SOFT_MATCH_TEST
 */

const PUMA_SOFT_MATCH = {
  OUTPUT_SHEET: 'PUMA_SOFT_MATCH_TEST',

  // We will tune these once we test against a real tracker.
  HEADER_ROW: 1,
  DATA_START_ROW: 4,

  // Expected tracker headers / aliases.
  HEADER_ALIASES: {
    type: ['Type', 'Location', 'Type/Location'],
    manufacturer: ['Manufacturer', 'MFG', 'Vendor'],
    partNumber: ['Part Number', 'Part #', 'Part'],
    qty: ['Qty', 'QTY', 'Quantity'],
    status: ['Status'],
    po: ['PO', 'PO #', 'P.O.', 'PO Number'],
    esd: ['ESD', 'Estimated Ship Date', 'Estimated Ship Date (ESD)'],
    notes: ['Notes', 'Internal Notes'],
    pumaLineId: ['PUMA_LINE_ID'],
    fingerprint: ['PUMA_FINGERPRINT'],
    reviewFlag: ['PUMA_REVIEW_FLAG']
  }
};

/**
 * Test the soft matching engine on the ACTIVE tracker sheet.
 *
 * Usage:
 * 1. Open a project tracker tab.
 * 2. Run testSoftMatchOnActiveTracker().
 * 3. Review PUMA_SOFT_MATCH_TEST.
 */
function testSoftMatchOnActiveTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tracker = ss.getActiveSheet();

  if (!/project track/i.test(tracker.getName())) {
    SpreadsheetApp.getUi().alert('Open a Project Tracker tab first, then rerun.');
    return;
  }

  const trackerRows = readTrackerRowsForSoftMatch_(tracker);

  // For v1 test, we compare tracker rows to themselves.
  // This confirms header detection, fingerprinting, and tier logic without touching quote imports yet.
  const incomingQuoteRows = trackerRows.map(r => ({
    type: r.type,
    manufacturer: r.manufacturer,
    partNumber: r.partNumber,
    qty: r.qty,
    sourceRow: r.sourceRow
  }));

  const results = softMatchRows_(trackerRows, incomingQuoteRows);

  writeSoftMatchTestReport_(ss, tracker.getName(), results);

  SpreadsheetApp.getUi().alert(
    `Soft match test complete.\nTracker rows read: ${trackerRows.length}\nReport: ${PUMA_SOFT_MATCH.OUTPUT_SHEET}`
  );
}

/**
 * Core matching engine.
 *
 * Match tiers:
 * 1. EXACT_MATCH: Type + Manufacturer + Part Number + Qty
 * 2. SOFT_MATCH_TYPE: Same unique Type, but details changed
 * 3. NEW_LINE: Incoming line not matched
 * 4. REMOVED_OR_SUPERSEDED: Existing tracker line not found in incoming quote
 */
function softMatchRows_(existingRows, incomingRows) {
  const results = [];
  const usedExistingIds = new Set();

  const exactMap = buildExactMatchMap_(existingRows);
  const typeMap = buildTypeMatchMap_(existingRows);

  incomingRows.forEach(incoming => {
    const exactKey = buildQuoteFingerprint_(incoming);
    const exactCandidates = exactMap[exactKey] || [];

    const availableExact = exactCandidates.find(r => !usedExistingIds.has(r.matchId));

    if (availableExact) {
      usedExistingIds.add(availableExact.matchId);

      results.push({
        tier: 'EXACT_MATCH',
        action: 'AUTO_KEEP',
        confidence: 'HIGH',
        incoming,
        existing: availableExact,
        pumaLineId: availableExact.pumaLineId || generatePumaLineId_(),
        reviewFlag: '',
        reason: 'Type, manufacturer, part number, and quantity match exactly.'
      });

      return;
    }

    const typeKey = normalizeSoftMatchPart_(incoming.type);
    const typeCandidates = typeMap[typeKey] || [];
    const availableTypeCandidates = typeCandidates.filter(r => !usedExistingIds.has(r.matchId));

    if (availableTypeCandidates.length === 1) {
      const soft = availableTypeCandidates[0];
      usedExistingIds.add(soft.matchId);

      results.push({
        tier: 'SOFT_MATCH_TYPE',
        action: 'AUTO_LINK_WITH_REVIEW',
        confidence: 'MEDIUM',
        incoming,
        existing: soft,
        pumaLineId: soft.pumaLineId || generatePumaLineId_(),
        reviewFlag: buildReviewFlag_(soft, incoming),
        reason: 'Type matched uniquely, but manufacturer, part number, or quantity changed.'
      });

      return;
    }

    if (availableTypeCandidates.length > 1) {
      results.push({
        tier: 'AMBIGUOUS_TYPE_MATCH',
        action: 'REVIEW_REQUIRED',
        confidence: 'LOW',
        incoming,
        existing: null,
        pumaLineId: generatePumaLineId_(),
        reviewFlag: 'Multiple existing rows share this Type. Manual review required.',
        reason: 'Type is not unique in the existing tracker.'
      });

      return;
    }

    results.push({
      tier: 'NEW_LINE',
      action: 'CREATE_NEW',
      confidence: 'HIGH',
      incoming,
      existing: null,
      pumaLineId: generatePumaLineId_(),
      reviewFlag: '',
      reason: 'No exact or type match found.'
    });
  });

  existingRows.forEach(existing => {
    if (usedExistingIds.has(existing.matchId)) return;

    results.push({
      tier: 'REMOVED_OR_SUPERSEDED',
      action: 'MARK_REVIEW',
      confidence: 'HIGH',
      incoming: null,
      existing,
      pumaLineId: existing.pumaLineId || generatePumaLineId_(),
      reviewFlag: 'Existing tracker row not found in latest quote import.',
      reason: 'Existing row has no incoming exact or soft match.'
    });
  });

  return results;
}

function readTrackerRowsForSoftMatch_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < PUMA_SOFT_MATCH.DATA_START_ROW) return [];

  const headerRowIndex = PUMA_SOFT_MATCH.HEADER_ROW - 1;
  const headers = values[headerRowIndex].map(h => String(h || '').trim());

  const col = buildSoftMatchColumnMap_(headers);

  const rows = [];

  for (let r = PUMA_SOFT_MATCH.DATA_START_ROW - 1; r < values.length; r++) {
    const row = values[r];

    const type = getCellByIndex_(row, col.type);
    const manufacturer = getCellByIndex_(row, col.manufacturer);
    const partNumber = getCellByIndex_(row, col.partNumber);
    const qty = getCellByIndex_(row, col.qty);

    if (!type && !manufacturer && !partNumber && !qty) continue;

    rows.push({
      matchId: `ROW_${r + 1}`,
      sourceRow: r + 1,
      type,
      manufacturer,
      partNumber,
      qty,
      status: getCellByIndex_(row, col.status),
      po: getCellByIndex_(row, col.po),
      esd: getCellByIndex_(row, col.esd),
      notes: getCellByIndex_(row, col.notes),
      pumaLineId: getCellByIndex_(row, col.pumaLineId),
      fingerprint: getCellByIndex_(row, col.fingerprint),
      reviewFlag: getCellByIndex_(row, col.reviewFlag)
    });
  }

  return rows;
}

function buildSoftMatchColumnMap_(headers) {
  const map = {};

  Object.keys(PUMA_SOFT_MATCH.HEADER_ALIASES).forEach(key => {
    map[key] = findHeaderByAliases_(headers, PUMA_SOFT_MATCH.HEADER_ALIASES[key]);
  });

  return map;
}

function findHeaderByAliases_(headers, aliases) {
  const normalizedHeaders = headers.map(h => normalizeHeader_(h));

  for (let i = 0; i < aliases.length; i++) {
    const target = normalizeHeader_(aliases[i]);
    const idx = normalizedHeaders.indexOf(target);
    if (idx !== -1) return idx;
  }

  return -1;
}

function normalizeHeader_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^\w]+/g, '')
    .trim();
}

function getCellByIndex_(row, idx) {
  if (idx === -1 || idx == null) return '';
  return row[idx];
}

function buildExactMatchMap_(rows) {
  const map = {};

  rows.forEach(row => {
    const key = buildQuoteFingerprint_(row);
    if (!map[key]) map[key] = [];
    map[key].push(row);
  });

  return map;
}

function buildTypeMatchMap_(rows) {
  const map = {};

  rows.forEach(row => {
    const key = normalizeSoftMatchPart_(row.type);
    if (!key) return;
    if (!map[key]) map[key] = [];
    map[key].push(row);
  });

  return map;
}

function buildQuoteFingerprint_(row) {
  return [
    normalizeSoftMatchPart_(row.type),
    normalizeSoftMatchPart_(row.manufacturer),
    normalizeSoftMatchPart_(row.partNumber),
    normalizeQtyForMatch_(row.qty)
  ].join('|');
}

function normalizeSoftMatchPart_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[^\w.\-/ ]+/g, '')
    .trim();
}

function normalizeQtyForMatch_(value) {
  const n = Number(value);
  if (!isNaN(n)) return String(n);
  return String(value || '').trim().toLowerCase();
}

function buildReviewFlag_(existing, incoming) {
  const changes = [];

  if (normalizeSoftMatchPart_(existing.manufacturer) !== normalizeSoftMatchPart_(incoming.manufacturer)) {
    changes.push(`Manufacturer changed: "${existing.manufacturer}" → "${incoming.manufacturer}"`);
  }

  if (normalizeSoftMatchPart_(existing.partNumber) !== normalizeSoftMatchPart_(incoming.partNumber)) {
    changes.push(`Part changed: "${existing.partNumber}" → "${incoming.partNumber}"`);
  }

  if (normalizeQtyForMatch_(existing.qty) !== normalizeQtyForMatch_(incoming.qty)) {
    changes.push(`Qty changed: "${existing.qty}" → "${incoming.qty}"`);
  }

  return changes.join(' | ');
}

function generatePumaLineId_() {
  return 'PUMA-LINE-' + Utilities.getUuid().slice(0, 8);
}

function debugSoftMatchHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const values = sheet.getDataRange().getValues();

  for (let r = 0; r < Math.min(values.length, 10); r++) {
    Logger.log(`ROW ${r + 1}: ` + JSON.stringify(values[r]));
  }
}

function writeSoftMatchTestReport_(ss, trackerName, results) {
  let sh = ss.getSheetByName(PUMA_SOFT_MATCH.OUTPUT_SHEET);
  if (!sh) sh = ss.insertSheet(PUMA_SOFT_MATCH.OUTPUT_SHEET);

  const rows = [[
    'Tracker',
    'Tier',
    'Action',
    'Confidence',
    'PUMA_LINE_ID',
    'Existing Row',
    'Existing Type',
    'Existing Manufacturer',
    'Existing Part',
    'Existing Qty',
    'Incoming Source Row',
    'Incoming Type',
    'Incoming Manufacturer',
    'Incoming Part',
    'Incoming Qty',
    'Review Flag',
    'Reason'
  ]];

  results.forEach(result => {
    const existing = result.existing || {};
    const incoming = result.incoming || {};

    rows.push([
      trackerName,
      result.tier,
      result.action,
      result.confidence,
      result.pumaLineId,
      existing.sourceRow || '',
      existing.type || '',
      existing.manufacturer || '',
      existing.partNumber || '',
      existing.qty || '',
      incoming.sourceRow || '',
      incoming.type || '',
      incoming.manufacturer || '',
      incoming.partNumber || '',
      incoming.qty || '',
      result.reviewFlag || '',
      result.reason || ''
    ]);
  });

  sh.clearContents();
  sh.clearFormats();

  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, rows[0].length);

  const header = sh.getRange(1, 1, 1, rows[0].length);
  header.setFontWeight('bold');
  header.setBackground('#d9ead3');

  if (rows.length > 1) {
    const tierRange = sh.getRange(2, 2, rows.length - 1, 1);
    const tiers = tierRange.getValues();

    tierRange.setBackgrounds(tiers.map(row => {
      const tier = String(row[0] || '').toUpperCase();

      if (tier === 'EXACT_MATCH') return ['#d9ead3'];
      if (tier === 'SOFT_MATCH_TYPE') return ['#fff2cc'];
      if (tier === 'AMBIGUOUS_TYPE_MATCH') return ['#f4cccc'];
      if (tier === 'NEW_LINE') return ['#cfe2f3'];
      if (tier === 'REMOVED_OR_SUPERSEDED') return ['#d9d2e9'];

      return ['#ffffff'];
    }));
  }
}