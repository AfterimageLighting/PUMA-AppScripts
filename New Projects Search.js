// =============================
// PUMA - New Project Folder Audit
// =============================

const TRACKER_CONFIG_SHEET_NAME = "Tracker Config";
const TRACKER_NAME_HEADER = "Project";
const OUTPUT_SHEET_NAME = "New Projects";

// Uses global OPEN_PROJECTS_FOLDER_ID already defined elsewhere in PUMA


function pumaFindNewOpenProjectFolders() {

  const ss = SpreadsheetApp.getActive();
  const trackerSheet = ss.getSheetByName(TRACKER_CONFIG_SHEET_NAME);
  if (!trackerSheet) throw new Error(`Sheet not found: ${TRACKER_CONFIG_SHEET_NAME}`);

  const trackerSet = buildTrackerNameSet_(trackerSheet, TRACKER_NAME_HEADER);

  const openProjectsFolder = DriveApp.getFolderById(OPEN_PROJECTS_FOLDER_ID);
  const folderIter = openProjectsFolder.getFolders();

  const missing = [];

  while (folderIter.hasNext()) {
    const f = folderIter.next();
    const rawFolderName = f.getName();

    // 🔥 STEP 1 — Extract clean project name
    const extractedProjectName = extractProjectNameFromFolder_(rawFolderName);

    // 🔥 STEP 2 — Normalize
    const normalized = normalizeKey_(extractedProjectName);

    // 🔥 STEP 3 — Compare
    if (!trackerSet.has(normalized)) {
      missing.push([
        rawFolderName,
        f.getId(),
        f.getDateCreated ? f.getDateCreated() : "",
        f.getLastUpdated ? f.getLastUpdated() : ""
      ]);
    }
  }

  const out = getOrCreateSheet_(ss, OUTPUT_SHEET_NAME);
  out.clear();

  out.getRange(1,1,1,4).setValues([[
    "Folder Name",
    "Folder ID",
    "Created",
    "Last Updated"
  ]]);

  if (missing.length) {
    out.getRange(2,1,missing.length,4).setValues(missing);
  }

  out.autoResizeColumns(1,4);

  SpreadsheetApp.getActive().toast(
    `Audit complete: ${missing.length} folder(s) appear NEW.`,
    "PUMA",
    6
  );
}


// ========================================
// Extract Project Name From Folder Name
// ========================================

function extractProjectNameFromFolder_(folderName) {
  const name = String(folderName || "").trim();
  if (!name) return name;

  const firstDash = name.indexOf("-");
  if (firstDash === -1) return name;

  // everything before the first dash is the project name
  const left = name.substring(0, firstDash).trim();
  return left || name;
}


// ========================================
// Build Tracker Set
// ========================================

function buildTrackerNameSet_(sheet, headerText) {

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return new Set();

  const headers = values[0].map(h => String(h || "").trim().toLowerCase());
  const colIndex = headers.findIndex(h => h === headerText.toLowerCase());

  if (colIndex === -1) {
    throw new Error(`Header "${headerText}" not found in ${sheet.getName()}`);
  }

  const set = new Set();

  for (let r = 1; r < values.length; r++) {
    const name = values[r][colIndex];
    if (!name) continue;
    set.add(normalizeKey_(name));
  }

  return set;
}


// ========================================
// Normalizer
// ========================================

function normalizeKey_(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^\w\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}


// ========================================
// Sheet Helper
// ========================================

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
