function debugOnEdit(e) {
  // prove trigger fired even before any logic
  console.log('🔥 debugOnEdit fired', new Date().toISOString());

  // show a toast in the sheet so you see it locally
  try { e?.source?.toast('debugOnEdit fired at ' + new Date().toLocaleTimeString(), 'PUMA', 3); } catch (_) {}

  // dump key event fields
  try {
    const sh = e?.source?.getActiveSheet();
    const info = {
      a1: e?.range?.getA1Notation(),
      row: e?.range?.getRow(),
      col: e?.range?.getColumn(),
      sheetName: sh ? sh.getName() : '(no sheet)',
      user: Session.getActiveUser().getEmail(),
    };
    console.log('ℹ️ event', JSON.stringify(info));
  } catch (err) {
    console.error('event parse error', err);
  }
}
