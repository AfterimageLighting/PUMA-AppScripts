// CLASP sync test - safe comment only
/**
 * PUMA Master Menu
 * Centralized toolbar definition
 * VERSION: 2026-01-13
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('PUMA');

  /* =========================
   * Dashboard
   * ========================= */
  const dashboardMenu = ui.createMenu('Dashboard')
    .addItem('Add Back to Dashboard Links', 'addBackToDashboardLinks')
    .addItem('Refresh Dashboard', 'refreshDashboardProjectTrackers')
    .addItem('Update Pipeline', 'applyProbabilityFormattingToAllTrackers');

  /* =========================
   * Tracker Configuration
   * ========================= */
  const trackerConfigMenu = ui.createMenu('Tracker Configuration')
    .addItem('Update Open Projects', 'TrackerConfig')
    .addItem(
      'Sync ALL Configured Quotes → Trackers',
      'syncAllConfiguredTrackers'
    )
    .addItem('Test Soft Match: Active Tracker vs Live Quotes', 'testSoftMatchActiveTrackerVsLiveQuotes')
    .addItem('Controlled Write: Soft Match IDs + Flags', 'controlledWriteSoftMatchActiveTrackerVsLiveQuotes')
    .addItem('Safe Sync Active Tracker', 'safeSyncActiveTracker');

  /* =========================
   * ESD Sheet
   * ========================= */
  const esdMenu = ui.createMenu('ESD Sheet')
    .addItem('Sync ESD', 'syncESD');

  /* =========================
   * Purchase Orders
   * ========================= */
  const poMenu = ui.createMenu('Purchase Orders')
    .addItem('Import POs (Only New PDFs)', 'importPOs_onlyNew')
    .addItem('Import POs (Scan Recent PDFs)', 'importPOs_scanAll')
    .addItem('Reprocess Selected Log Rows', 'importPOs_reprocessSelectedLogRows');

  /* =========================
   * Assemble Menu
   * ========================= */
  menu
    .addSubMenu(dashboardMenu)
    .addSubMenu(trackerConfigMenu)
    .addSubMenu(esdMenu)
    .addSubMenu(poMenu)
    .addToUi();
}