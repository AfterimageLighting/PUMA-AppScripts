/**
 * Builds "Tracker config" by scanning each folder listed in "Open Projects"
 * and recording every Google Sheet inside each folder.
 *
 * INPUT (tab):  "Open Projects"
 *   Col A: Folder Name (Project)
 *   Col B: Folder ID
 *
 * OUTPUT (tab): "Tracker config"
 *   Project | Folder ID | Quote Name | Quote ID
 *
 * Requires Advanced Google Service: Drive API (Drive.Files.list)
 */

/**
 * Full updated buildTrackerConfig with fail-safes:
 *  - preserves prior Enable choices
 *  - treats inaccessible ("unknown") quotes as "do not touch" (Enable = false)
 *  - only auto-disable (uncheck) when a sheet/tab name contains "approved" AND that tab is protected (not warning-only)
 *
 * Requires Advanced Drive Service: Drive (Drive.Files.list)
 */

function buildTrackerConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheetName = "Open Projects";
  const outputSheetName = "Tracker config";

  const inputSheet = ss.getSheetByName(inputSheetName);
  if (!inputSheet) throw new Error(`Missing sheet: "${inputSheetName}"`);

  // Capture the current config BEFORE rebuilding it. Existing quote->tracker
  // mappings are valuable human/system evidence and are preserved when valid.
  const existingConfigSheet = ss.getSheetByName(outputSheetName);
  const priorMap = getPriorTrackerConfigMap_(existingConfigSheet);
  const registry = (typeof pumaBuildProjectRegistry_ === 'function')
    ? pumaBuildProjectRegistry_(ss)
    : null;

  let out = existingConfigSheet || ss.insertSheet(outputSheetName);

  const lastRow = inputSheet.getLastRow();
  const sourceRows = lastRow >= 2
    ? inputSheet.getRange(2, 1, lastRow - 1, 2).getValues()
    : [];

  const output = [];
  const seenQuoteIds = new Set();

  sourceRows.forEach(([projectValue, folderIdValue]) => {
    const project = String(projectValue || "").trim();
    const folderID = String(folderIdValue || "").trim();
    if (!project || !folderID) return;

    let resolution = null;
    if (registry && typeof pumaResolveProjectWithRegistry_ === 'function') {
      resolution = pumaResolveProjectWithRegistry_(registry, project, {});
    }

    let files = [];
    try {
      files = (typeof listPumaQuoteSheetsInFolder_ === 'function')
        ? listPumaQuoteSheetsInFolder_(folderID)
        : listSheetsInFolder_(folderID).filter(file => {
            return typeof isLikelyPumaQuotation_ === 'function'
              ? isLikelyPumaQuotation_(file)
              : true;
          });
    } catch (err) {
      output.push([
        false,
        project,
        '',
        '',
        '[FOLDER SCAN ERROR]',
        '',
        'ERROR',
        err.message || String(err)
      ]);
      return;
    }

    const familyCounts = {};
    files.forEach(file => {
      const family = pumaQuoteFamilyKey_(file.name || '', project);
      if (family) familyCounts[family] = (familyCounts[family] || 0) + 1;
    });

    files.forEach(file => {
      const quoteId = String(file.id || '').trim();
      const quoteName = String(file.name || '').trim();
      const familyKey = pumaQuoteFamilyKey_(quoteName, project);
      const familyCollision = familyKey && (familyCounts[familyKey] || 0) > 1;
      if (!quoteId || seenQuoteIds.has(quoteId)) return;
      seenQuoteIds.add(quoteId);

      const prior = priorMap.get(quoteId);
      let trackerName = '';
      let resolutionMethod = '';
      let resolutionNote = '';

      // Existing explicit quote->tracker mapping wins if the tracker still exists.
      if (prior && prior.trackerName && ss.getSheetByName(prior.trackerName)) {
        trackerName = prior.trackerName;
        resolutionMethod = 'PRIOR_CONFIG';
        resolutionNote = 'Preserved existing quote-to-tracker mapping.';
      } else if (resolution && resolution.status === 'CONFIRMED' && resolution.trackerSheetName) {
        trackerName = resolution.trackerSheetName;
        resolutionMethod = resolution.method || 'PROJECT_RESOLVER';
        resolutionNote = (resolution.reasons || []).join(' ');
      } else {
        resolutionMethod = resolution && resolution.method
          ? resolution.method
          : 'TRACKER_SETUP_REQUIRED';
        resolutionNote = resolution && resolution.reasons && resolution.reasons.length
          ? resolution.reasons.join(' ')
          : 'No confirmed existing tracker. Row is disabled until project setup is resolved.';
      }

      const approvedAndLocked = isApprovedAndLocked_(quoteId, quoteName);
      const canEnable = !!trackerName && !approvedAndLocked;
      const priorEnable = prior ? prior.enabled : undefined;

      // Existing configured quotes preserve their prior enable choice.
      // Newly discovered quotes are NEVER auto-enabled; they require review.
      const enableValue = canEnable && prior
        ? Boolean(priorEnable)
        : false;

      if (!prior && familyCollision) {
        resolutionMethod = 'POSSIBLE_DUPLICATE_QUOTE_FAMILY';
        resolutionNote =
          'Multiple quote files in this project folder normalize to the same quote family. ' +
          'New file is disabled until the correct version/scope is confirmed.';
      } else if (!prior) {
        resolutionMethod = resolutionMethod || 'NEW_QUOTE_DISCOVERED';
        resolutionNote =
          (resolutionNote ? resolutionNote + ' ' : '') +
          'Newly discovered quote is disabled until reviewed.';
      }

      output.push([
        enableValue,
        project,
        trackerName,
        prior && prior.dateUpdated ? prior.dateUpdated : '',
        quoteName,
        quoteId,
        resolutionMethod,
        approvedAndLocked
          ? 'APPROVED/LOCKED - disabled'
          : resolutionNote
      ]);
    });
  });

  // Preserve prior config references that were not rediscovered in the current
  // direct-folder scan. Never silently drop them: keep them disabled for review.
  priorMap.forEach((prior, quoteId) => {
    if (seenQuoteIds.has(quoteId)) return;
    output.push([
      false,
      prior.project || '',
      prior.trackerName && ss.getSheetByName(prior.trackerName) ? prior.trackerName : '',
      prior.dateUpdated || '',
      prior.quoteName || '[Prior configured quote]',
      quoteId,
      'STALE_CONFIG_REFERENCE',
      'Quote ID existed in prior Tracker config but was not rediscovered in the current Open Projects direct-folder scan. Preserved disabled for review.'
    ]);
  });

  out.clearContents();

  const headers = [
    "Enable?",
    "Project",
    "Tracker Sheet Name",
    "Date Updated",
    "Quote Name",
    "Quote Sheet ID",
    "Project Resolution",
    "Resolution Note"
  ];
  out.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (output.length) {
    out.getRange(2, 1, output.length, headers.length).setValues(output);
  }

  out.setFrozenRows(1);
  out.autoResizeColumns(1, headers.length);

  if (out.getLastRow() > 1) {
    out.getRange(2, 1, out.getLastRow() - 1, 1).insertCheckboxes();
    out.getRange(2, 4, out.getLastRow() - 1, 1).setNumberFormat("m/d/yyyy");
  }

  out.getRange(1, headers.length + 2).setValue("Last built:");
  out.getRange(1, headers.length + 3).setValue(new Date());
}

/**
 * Read prior Tracker config to preserve Enable states.
 * Returns Map of quoteSheetId -> boolean
 */
function getPriorTrackerConfigMap_(sheet) {
  const map = new Map();
  if (!sheet || sheet.getLastRow() < 2) return map;

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(v => String(v || '').trim().toLowerCase());
  const idx = name => headers.indexOf(String(name).toLowerCase());

  let enableIdx = idx('enable?');
  const projectIdx = idx('project');
  const trackerIdx = idx('tracker sheet name');

  // Legacy PUMA has a blank A1 while column A still contains checkbox values.
  // Treat column A as Enable? only when B/C prove this is the known config layout.
  if (enableIdx === -1 && projectIdx === 1 && trackerIdx === 2) {
    enableIdx = 0;
  }
  const dateIdx = idx('date updated');
  const quoteNameIdx = idx('quote name');
  const quoteIdIdx = idx('quote sheet id');

  if (quoteIdIdx === -1) return map;

  values.slice(1).forEach(row => {
    const quoteId = String(row[quoteIdIdx] || '').trim();
    if (!quoteId) return;
    map.set(quoteId, {
      enabled: enableIdx === -1 ? undefined : Boolean(row[enableIdx]),
      project: projectIdx === -1 ? '' : String(row[projectIdx] || '').trim(),
      trackerName: trackerIdx === -1 ? '' : String(row[trackerIdx] || '').trim(),
      dateUpdated: dateIdx === -1 ? '' : row[dateIdx],
      quoteName: quoteNameIdx === -1 ? '' : String(row[quoteNameIdx] || '').trim()
    });
  });

  return map;
}


function getPriorEnableMap_(sheet) {
  const rich = getPriorTrackerConfigMap_(sheet);
  const simple = new Map();
  rich.forEach((value, key) => simple.set(key, value.enabled));
  return simple;
}

function pumaQuoteFamilyKey_(quoteName, projectName) {
  let name = String(quoteName || '').toLowerCase();
  const projectKey = String(projectName || '')
    .toLowerCase()
    .replace(/\b(residence|project)\b/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  // Remove obvious copy/version/date decorations, but KEEP scope words such as
  // adder, decorative, bulb, heater, landscape, cove, etc.
  name = name
    .replace(/^copy of\s+/i, '')
    .replace(/\bapproved\b/gi, ' ')
    .replace(/\binternal use only\b/gi, ' ')
    .replace(/\brev(?:ision)?\.?\s*\d+\b/gi, ' ')
    .replace(/\b\d{6,8}\b/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (projectKey) {
    const tokens = new Set(projectKey.split(' ').filter(Boolean));
    name = name
      .split(' ')
      .filter(token => !tokens.has(token))
      .join(' ')
      .trim();
  }

  return name;
}


/**
 * Returns true if spreadsheet should be treated as "approved + locked" OR is unknown/inaccessible (fail-safe),
 * meaning we should NOT touch it (Enable = false).
 *
 * Logic:
 *  - Try open the spreadsheet by ID.
 *  - Find any tab whose name contains "approved" (case-insensitive).
 *  - If found, confirm that specific approved tab has at least one protection (sheet or range) that is NOT a warning-only protection.
 *  - If both name contains "approved" AND there is a real protection on that tab => return true.
 *  - If no approved tab name found => return false.
 *  - If any error occurs opening the spreadsheet (permissions/timeouts) => return true (fail-safe: don't touch).
 */
function isApprovedAndLocked_(quoteSpreadsheetId, quoteFileName) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `apprLocked:${quoteSpreadsheetId}`;
  const cached = cache.get(cacheKey);
  if (cached !== null) return cached === "true";

  try {
    const qss = SpreadsheetApp.openById(quoteSpreadsheetId);
    const approvedSheets = qss.getSheets().filter(s =>
      /approved/i.test(String(s.getName() || ''))
    );

    if (approvedSheets.length === 0) {
      cache.put(cacheKey, "false", 21600);
      return false;
    }

    const sheetProtections = qss.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    const rangeProtections = qss.getProtections(SpreadsheetApp.ProtectionType.RANGE);

    for (const sheet of approvedSheets) {
      const sName = sheet.getName();

      const hasRealSheetProtection = sheetProtections.some(p => {
        try {
          const rng = p.getRange();
          return rng && rng.getSheet().getName() === sName && !p.isWarningOnly();
        } catch (e) {
          return false;
        }
      });

      const hasRealRangeProtection = rangeProtections.some(p => {
        try {
          const rng = p.getRange();
          return rng && rng.getSheet().getName() === sName && !p.isWarningOnly();
        } catch (e) {
          return false;
        }
      });

      if (hasRealSheetProtection || hasRealRangeProtection) {
        cache.put(cacheKey, "true", 21600);
        return true;
      }
    }

    cache.put(cacheKey, "false", 21600);
    return false;
  } catch (err) {
    // Fail safe: inaccessible quote workbooks must never be auto-enabled.
    cache.put(cacheKey, "true", 21600);
    return true;
  }
}

/**
 * Drive v3 listing for Google Sheets in a folder (Shared Drive safe).
 * Requires Advanced Drive Service: Drive
 */
function listSheetsInFolder_(folderId) {
  const results = [];
  let pageToken;

  const q = [
    `'${folderId}' in parents`,
    `mimeType='application/vnd.google-apps.spreadsheet'`,
    `trashed=false`
  ].join(" and ");

  do {
    const resp = Drive.Files.list({
      q,
      fields: "nextPageToken, files(id, name)",
      pageToken: pageToken,
      pageSize: 1000,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    });

    const files = resp.files || [];
    files.forEach(f => results.push({ id: f.id, name: f.name }));

    pageToken = resp.nextPageToken;
  } while (pageToken);

  results.sort((a, b) => a.name.localeCompare(b.name));
  return results;
}
