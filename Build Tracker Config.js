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

    files.forEach(file => {
      const quoteId = String(file.id || '').trim();
      const quoteName = String(file.name || '').trim();
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
      const enableValue = canEnable
        ? (priorEnable !== undefined ? priorEnable : true)
        : false;

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

  const enableIdx = idx('enable?');
  const projectIdx = idx('project');
  const trackerIdx = idx('tracker sheet name');
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
