/**
 * PUMA System Constants
 * Centralized config to avoid hardcoding everywhere
 */

const PUMA = {
  SHEETS: {
    TRACKER_CONFIG: 'Tracker config',
    DASHBOARD: 'Dashboard',
    RAW_PO_IMPORT: 'RAW_PO_IMPORT',
    ESD: 'ESD'
  },

  STATUS: {
    TO_BE_ORDERED: 'To Be Ordered',
    ORDERED: 'Ordered',
    RECEIVED: 'Received',
    SCHEDULED: 'Scheduled',
    DELIVERED: 'Delivered'
  },

  SETTINGS: {
    MAX_RUNTIME_MS: 5 * 60 * 1000,
    TEST_SPREADSHEET_ID: '1wjGB4dUTbBUiWVC7Kh0huVw8wM_ppmjTgktLw6ElVf4'
  }
};

function pumaIsTestWorkbook_() {
  try {
    return SpreadsheetApp.getActive().getId() === PUMA.SETTINGS.TEST_SPREADSHEET_ID;
  } catch (err) {
    return false;
  }
}