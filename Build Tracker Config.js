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

  // read previous enable map to preserve user choices
  const priorEnableMap = getPriorEnableMap_(ss.getSheetByName(outputSheetName));

  let out = ss.getSheetByName(outputSheetName);
  if (!out) out = ss.insertSheet(outputSheetName);
  out.clearContents();

  // Columns: A Enable? | B Project | C Tracker Sheet Name | D Date Updated | E Quote Name | F Quote Sheet ID
  const headers = ["Enable?", "Project", "Tracker Sheet Name", "Date Updated", "Quote Name", "Quote Sheet ID"];
  out.getRange(1, 1, 1, headers.length).setValues([headers]);

  const lastRow = inputSheet.getLastRow();
  if (lastRow < 2) return;

  const rows = inputSheet
    .getRange(2, 1, lastRow - 1, 2) // A:B
    .getValues()
    .filter(r => String(r[0] || "").trim() || String(r[1] || "").trim());

  const output = [];

  rows.forEach(([folderName, folderId]) => {
    const project = String(folderName || "").trim();
    const folderID = String(folderId || "").trim();
    if (!folderID) return;

    const files = listSheetsInFolder_(folderID); // [{id,name}, ...]

    files.forEach(file => {
      const quoteId = file.id;
      const quoteName = file.name || "";

      // Determine if this quote is "approved + locked"
      const approvedAndLocked = isApprovedAndLocked_(quoteId, quoteName);

      // If unknown/inaccessible, isApprovedAndLocked_ will return true because we treat unknown as fail-safe (do not touch)
      // Logic for Enable:
      // - If approvedAndLocked => ALWAYS false (unchecked)
      // - Else if prior exists => keep the user's prior choice
      // - Else default to true (enabled)
      const prior = priorEnableMap.get(String(quoteId));
      const enableValue = approvedAndLocked ? false : (prior !== undefined ? prior : true);

      output.push([
        enableValue,
        project,
        `${project} - Project Tracker`,
        "",           // Date Updated left blank for now
        quoteName,
        quoteId
      ]);
    });
  });

  if (output.length) {
    out.getRange(2, 1, output.length, headers.length).setValues(output);
  }

  // Formatting & UX
  out.setFrozenRows(1);
  out.autoResizeColumns(1, headers.length);

  if (out.getLastRow() > 1) {
    out.getRange(2, 1, out.getLastRow() - 1, 1).insertCheckboxes();
  }

  out.getRange(2, 4, Math.max(1, out.getLastRow() - 1), 1).setNumberFormat("m/d/yyyy");

  // friendly timestamp
  out.getRange(1, headers.length + 2).setValue("Last built:");
  out.getRange(1, headers.length + 3).setValue(new Date());
}


/**
 * Read prior Tracker config to preserve Enable states.
 * Returns Map of quoteSheetId -> boolean
 */
function getPriorEnableMap_(sheet) {
  const map = new Map();
  if (!sheet) return map;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;

  // Expect columns A..F where F is Quote Sheet ID
  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  values.forEach(r => {
    const enable = r[0];
    const quoteId = r[5];
    if (quoteId) map.set(String(quoteId), Boolean(enable));
  });
  return map;
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

  // Fail-safe: if we can't confirm, do NOT touch it.
  // But we only do the expensive open/check if it looks like it might be approved.
  const nameLooksApproved = String(quoteFileName || "").toLowerCase().includes("approved");
  if (!nameLooksApproved) {
    // not likely approved → treat as NOT approved+locked
    cache.put(cacheKey, "false", 21600); // 6 hours
    return false;
  }

  try {
    const qss = SpreadsheetApp.openById(quoteSpreadsheetId);
    const sheets = qss.getSheets();

    const approvedSheets = sheets.filter(s =>
      String(s.getName() || "").toLowerCase().includes("approved")
    );

    if (approvedSheets.length === 0) {
      cache.put(cacheKey, "false", 21600);
      return false;
    }

    // Pull protections ONCE each (faster)
    const sheetProtections = qss.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    const rangeProtections = qss.getProtections(SpreadsheetApp.ProtectionType.RANGE);

    for (const sheet of approvedSheets) {
      const sName = sheet.getName();

      const hasRealSheetProtection = sheetProtections.some(p => {
        try {
          const rng = p.getRange();
          return rng && rng.getSheet().getName() === sName && !p.isWarningOnly();
        } catch (e) { return false; }
      });

      if (hasRealSheetProtection) {
        cache.put(cacheKey, "true", 21600);
        return true;
      }

      const hasRealRangeProtection = rangeProtections.some(p => {
        try {
          const rng = p.getRange();
          return rng && rng.getSheet().getName() === sName && !p.isWarningOnly();
        } catch (e) { return false; }
      });

      if (hasRealRangeProtection) {
        cache.put(cacheKey, "true", 21600);
        return true;
      }
    }

    cache.put(cacheKey, "false", 21600);
    return false;

  } catch (err) {
    // unknown -> do not touch
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

/**
 * Returns [{id, name}, ...] for Google Sheets directly in a folder.
 * Shared Drive compatible.
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

