/**
 * PUMA simple-trigger dispatcher.
 *
 * Apps Script exposes one global onEdit(e). Individual feature handlers live
 * under unique names so they cannot overwrite each other in the global project
 * namespace.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const handlers = [
    pumaHandleTrackerEsdEdit_,
    pumaHandleEsdSheetEdit_
  ];

  handlers.forEach(function(handler) {
    try {
      handler(e);
    } catch (err) {
      console.error('[PUMA onEdit] ' + handler.name + ': ' + err);
    }
  });
}
