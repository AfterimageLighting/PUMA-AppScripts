/** Adds a custom menu on open */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Navigate')
    .addItem('Open Sidebar', 'openSidebar')
    .addToUi();
}

/** Builds and shows the sidebar */
function openSidebar() {
  const html = HtmlService
    .createTemplateFromFile('Sidebar')   // loads Sidebar.html
    .evaluate()
    .setTitle('Sheet Navigator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Return visible sheet list (filtered) to the client */
function getVisibleSheets() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheets()
    .filter(sh => !sh.isSheetHidden())
    .filter(sh => !/unmatched/i.test(sh.getName()))
    .map(sh => ({ name: sh.getName(), id: sh.getSheetId() }));
}

/** Jump to a sheet by name */
function goToSheetByName(name) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name);
  if (sh) ss.setActiveSheet(sh);
}
