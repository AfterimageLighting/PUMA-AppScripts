function testChatWebhook() {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('PUMA_CHAT_WEBHOOK_URL') || '';
  if (!webhookUrl) throw new Error('PUMA_CHAT_WEBHOOK_URL is not configured.');
  if (typeof pumaIsTestWorkbook_ === 'function' && pumaIsTestWorkbook_()) {
    throw new Error('PUMA TEST: Chat webhook calls are disabled.');
  }
  const payload = { text: 'Ping from Apps Script ✅' };
  const res = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true, // ← so we can read error bodies
  });
  console.log('HTTP', res.getResponseCode(), res.getContentText());
}