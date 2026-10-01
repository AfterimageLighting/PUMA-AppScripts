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
   * Dashboard & Pipeline
   * ========================= */
  const dashboardMenu = ui.createMenu('Dashboard & Pipeline')
    .addItem('Switch to Pipeline View', 'switchToPipelineView')
    .addItem('Switch to Dashboard View', 'switchToDashboardView')
    .addSeparator()
    .addItem('Expand ALL Projects', 'expandDashboard')
    .addItem('Collapse All Projects', 'collapseDashboard')
    .addSeparator()
    .addItem('Refresh Dashboard', 'refreshDashboardProjectTrackers')
    .addItem('Update Quote Closing Statuses', 'updateTrackerConfigQuoteStatuses');

  /* =========================
   * Tracker Configuration
   * ========================= */
  const trackerConfigMenu = ui.createMenu('Tracker Configuration')
    .addItem('Audit Open Projects Manifest (Read Only)', 'pumaAuditOpenProjectsManifest')
    .addItem('Reconcile Open Projects Manifest', 'pumaSyncOpenProjectsManifest')
    .addSeparator()
    .addItem('Audit Missing Quotes (Read Only)', 'reportPumaMissingQuotes')
    .addItem('Add Audited Missing Quotes to Config', 'autoAddPumaMissingQuotesToConfig')
    .addSeparator()
    .addItem('Rebuild Tracker Config from Open Projects', 'buildTrackerConfig')
    .addSeparator()
    .addItem('Test Soft Match: Active Tracker vs Live Quotes', 'testSoftMatchActiveTrackerVsLiveQuotes')
    .addItem('Controlled Write: Soft Match IDs + Flags', 'controlledWriteSoftMatchActiveTrackerVsLiveQuotes')
    .addItem('Safe Sync Active Tracker', 'safeSyncActiveTracker')
    .addItem('Safe Sync ALL Trackers', 'safeSyncALLConfiguredTrackers');

  /* =========================
   * ESD Sheet
   * ========================= */
  const esdMenu = ui.createMenu('ESD Sheet')
    .addItem('Sync ESD', 'syncESD');

  /* =========================
   * Purchase Orders
   * ========================= */
  const poMenu = ui.createMenu('Purchase Orders')
    .addItem('Audit RAW PO Project Resolution (Read Only)', 'pumaAuditRawPoProjectResolution')
    .addItem('Audit Project Registry (Read Only)', 'pumaAuditProjectRegistryReadOnly')
    .addSeparator()
    .addItem('Apply RAW PO Import', 'applyRawPoImportToProjectTrackers')
    .addSeparator()
    .addItem('Install Daily PO Import Trigger', 'installDailyPoImportTrigger')
    .addItem('Remove Daily PO Import Trigger', 'removeDailyPoImportTrigger');

  /* =========================
   * Assemble Menu
   * ========================= */
  menu
    .addSubMenu(dashboardMenu)
    .addSubMenu(trackerConfigMenu)
    .addSubMenu(esdMenu)
    .addSubMenu(poMenu)
    .addToUi();

  // Separate navigation menu, dispatched from the single project onOpen.
  if (typeof pumaAddNavigationMenu_ === 'function') {
    pumaAddNavigationMenu_();
  }
}