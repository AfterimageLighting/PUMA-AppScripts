function createGoogleTaskInChat(e) {
  console.log('🚀 createGoogleTaskInChat start', new Date().toISOString());
  if (!e || !e.source || !e.range) { console.log('bad event'); return; }

  const ss = e.source;
  const sh = ss.getActiveSheet();
  const sheetName = sh.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const a1  = e.range.getA1Notation();

  console.log('✍️ edit', JSON.stringify({ sheetName, row, col, a1, value: e.value }));
  try { ss.toast(`PUMA heard ${a1} on "${sheetName}"`, 'PUMA', 3); } catch (_) {}

  // Only act on “… Tasks” tabs and edits in B..E (Task..Status)
  const isTasksTab = sheetName.trim().toLowerCase().endsWith('tasks');
  if (!isTasksTab) { console.log('Not a Tasks tab.'); return; }
  if (col < 2 || col > 5) { console.log('Not B..E; ignore'); return; }

  // Read values
  const getDisp = (c) => sh.getRange(row, c).getDisplayValue();
  const project  = getDisp(1); // A
  const task     = getDisp(2); // B
  const assignee = getDisp(3); // C
  const dueDisp  = getDisp(4); // D
  const status   = getDisp(5); // E

  // Require only Status (and a non-empty Task title)
  if (!task)   { console.log('No task yet');   return; }
  if (!status) { console.log('No status yet'); return; }

  // Duplicate guard: Column G “Sent to Chat”
  const postedCol = 7; // G
  const already = sh.getRange(row, postedCol).getValue();
  if (already) { console.log('Already posted at', already); return; }

  // Build message text:
  //  - line 1 = Task (clean title for Chat’s Create task)
  //  - line 2 = Project | Assignee | Due | Status (omit empty fields)
  const parts = [];
  if (project)  parts.push(`Project: ${project}`);
  if (assignee) parts.push(`Assignee: ${assignee}`);
  if (dueDisp)  parts.push(`Due: ${dueDisp}`);
  parts.push(`Status: ${status}`); // status exists by requirement

  const messageText = parts.length ? `${task}\n${parts.join(' | ')}` : task;
  const payload = { text: messageText };

  const webhookUrl = 'https://chat.googleapis.com/v1/spaces/AAQANZ7wX1U/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=0ApBRP_gsoG_aj7K3MPjdxQ0opy5VWN0YeivN4esxMg'; // keep key=... & token=...
  try {
    const res  = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    console.log('🌐 Chat POST', code, res.getContentText());

    if (code >= 200 && code < 300) {
      sh.getRange(row, postedCol).setValue(new Date()); // stamp “Sent to Chat”
      ss.toast('✅ Posted to Chat: ' + task, 'PUMA', 3);
    } else {
      ss.toast('⚠️ Chat returned ' + code + ' (see logs)', 'PUMA', 5);
    }
  } catch (err) {
    console.error('❌ UrlFetch error:', err);
    ss.toast('❌ Failed to send to Chat (see logs)', 'PUMA', 5);
  }

  console.log('✅ createGoogleTaskInChat end');
}
