function testChatWebhook() {
  const webhookUrl = 'https://chat.googleapis.com/v1/spaces/AAQANZ7wX1U/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=0ApBRP_gsoG_aj7K3MPjdxQ0opy5VWN0YeivN4esxMg'; // must have BOTH key=... & token=...
  const payload = { text: 'Ping from Apps Script ✅' };
  const res = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true, // ← so we can read error bodies
  });
  console.log('HTTP', res.getResponseCode(), res.getContentText());
}