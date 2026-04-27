/**
 * PUMA System Constants
 * Centralized config to avoid hardcoding everywhere
 */

const PUMA = {
  SHEETS: {
    TRACKER_CONFIG: 'Tracker Config',
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
    MAX_RUNTIME_MS: 5 * 60 * 1000
  }
};