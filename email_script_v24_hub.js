// OEM EXCESS Automation — Apps Script v25
// John Fluman / Intransit Technologies
// worker.js = brain; this file = thin I/O adapter (Gmail + Sheets only)

var SPREADSHEET_ID    = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
var MAIN_SHEET_NAME   = 'sheet1';
var DELETED_SHEET_NAME = 'Deleted Rows';
var NOTIFY_EMAIL      = 'john.fluman@intransittech.com';
var JOHN_EMAIL        = 'john.fluman@intransittech.com';
var DAVID_EMAIL       = 'david@fortetechno.com';
var BILL_EMAIL        = 'bill.pratt@intransittech.com';
var DEB_EMAIL         = 'deb@intransittech.com';
var BLOCKED_DOMAINS   = ['sourceschip.com', 'bulechip.com', 'feelchips.com', 'chip-wintrading.com', 'qizhongsmart.com'];
var INCOMING_LABEL    = 'oem-nostock-seen';
var FORTE_SHEET_ID    = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
var IN_STOCK_ID       = '1iOFHUBiWRgA6EjtO2ujoGpz-8v1qTRkgCXSvCa2Gf54';
var STAN_SHEET_ID     = '1pGRDpkqftQNoEYna53MxRJfUY8jEf5_w32FNa56OUIM';
var FORTE_HISTORY_COL = 9;
var FORTE_STATUS_COL  = 10;

var MSG_NEED_TP_500  = 'We need a target price to proceed. Please note there is a $500 minimum line requirement. Once we have your target we will get back to you right away.';
var MSG_NEED_TP_2000 = 'We need a target price to proceed. Please note there is a $2,000 minimum line requirement. Once we have your target we will get back to you right away.';
var MSG_CHECKING     = 'We are checking on it now. If we get a response from the OEM, I will respond to you right away. If we do not respond back to you, please consider this a no bid. Thank you very much for the opportunity.';
var MSG_BILL         = 'Bill will help with this request';

var HUB_URL    = 'https://intransit-hub.intransit-sales.workers.dev';
var HUB_SECRET = 'InTransit!Hub#2026';

function getBlockedDomains() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('blocked_domains');
  if (cached) return JSON.parse(cached);
  try {
    var resp = UrlFetchApp.fetch(HUB_URL + '/api/rules?type=blocked_domain', {
      headers: { Authorization: 'Bearer ' + HUB_SECRET }, muteHttpExceptions: true
    });
    var data = JSON.parse(resp.getContentText());
    var domains = (data.rules || []).map(function(r) { return r.key; });
    if (domains.length) {
      cache.put('blocked_domains', JSON.stringify(domains), 300);
      return domains;
    }
  } catch(e) { Logger.log('getBlockedDomains error: ' + e); }
  return BLOCKED_DOMAINS; // fallback to hardcoded
}

// Dispatcher — routes sidebar action to the correct executor
function executeAction(action, threadId, subject, fromH) {
  var type = action.type || 'create_draft';
  if (type === 'create_draft')      return executeCreateDraft(action, threadId, subject, fromH);
  if (type === 'add_forte')         return executeAddForte(action);
  if (type === 'remove_oem_excess') return executeRemoveOemExcess(action);
  if (type === 'apply_label')       return executeApplyLabel(action, threadId);
  if (type === 'update_rule')       return executeUpdateRule(action);
  if (type === 'multi')             return executeMulti(action, threadId, subject, fromH);
  return { ok: false, message: 'Unknown action type: ' + type };
}

function executeCreateDraft(action, threadId, subject, fromH) {
  var thread = GmailApp.getThreadById(threadId);
  if (!thread) return { ok: false, message: 'Thread not found.' };
  var msgs = thread.getMessages();
  var lastMsg = msgs[msgs.length - 1];
  var htmlBody = '<div dir="ltr">' + action.body.replace(/\n/g, '<br>') + getSignatureHTML() + '</div>';
  var draft = lastMsg.createDraftReply('', { htmlBody: htmlBody });
  if (!draft) return { ok: false, message: 'Draft creation failed.' };
  hubPostDraft(threadId, null, fromH || lastMsg.getFrom(), subject, action.body, draft.getId(), action.advice || '');
  return { ok: true, message: '✅ Draft created', draftId: draft.getId() };
}

function executeCreateAndSendDraft(action, threadId, subject, fromH) {
  var thread = GmailApp.getThreadById(threadId);
  if (!thread) return { ok: false, message: 'Thread not found.' };
  var msgs = thread.getMessages();
  var lastMsg = msgs[msgs.length - 1];
  // No advice block in outbound email
  var htmlBody = '<div dir="ltr">' + action.body.replace(/\n/g, '<br>') + getSignatureHTML() + '</div>';
  var draft = lastMsg.createDraftReply('', { htmlBody: htmlBody });
  if (!draft) return { ok: false, message: 'Draft creation failed.' };
  var draftId = draft.getId();
  var token = ScriptApp.getOAuthToken();
  var sendResp = UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/me/drafts/send', {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ id: draftId }),
    muteHttpExceptions: true
  });
  var sendData = JSON.parse(sendResp.getContentText());
  if (sendData.error) return { ok: false, message: 'Send failed: ' + (sendData.error.message || '') };
  hubLog('run', 'executeCreateAndSendDraft: sent — ' + subject, {});
  return { ok: true, message: '✅ Sent to ' + (fromH || lastMsg.getFrom()) };
}

function executeAddForte(action) {
  if (!action.mpn) return { ok: false, message: 'add_forte requires mpn.' };
  if (!action.qty) return { ok: false, message: 'add_forte requires qty — cardinal rule: never add without QTY.' };
  try {
    addToForteSheet(action.mpn, action.qty, action.tp || '', action.country || '', '');
    hubLog('run', 'executeAddForte: ' + action.mpn, { qty: action.qty, tp: action.tp });
    return { ok: true, message: '✅ Added ' + action.mpn + ' to Forte (QTY: ' + action.qty + ')' };
  } catch(e) { return { ok: false, message: 'Forte error: ' + e }; }
}

function executeRemoveOemExcess(action) {
  if (!action.mpn) return { ok: false, message: 'remove_oem_excess requires mpn.' };
  try {
    deletePart(action.mpn, 'sidebar-remove');
    hubLog('run', 'executeRemoveOemExcess: ' + action.mpn, {});
    return { ok: true, message: '✅ Removed ' + action.mpn + ' from OEM EXCESS' };
  } catch(e) { return { ok: false, message: 'Remove error: ' + e }; }
}

function executeApplyLabel(action, threadId) {
  if (!action.label) return { ok: false, message: 'apply_label requires label.' };
  try {
    var thread = GmailApp.getThreadById(threadId);
    if (!thread) return { ok: false, message: 'Thread not found.' };
    var lbl = GmailApp.getUserLabelByName(action.label) || GmailApp.createLabel(action.label);
    thread.addLabel(lbl);
    return { ok: true, message: '✅ Applied label: ' + action.label };
  } catch(e) { return { ok: false, message: 'Label error: ' + e }; }
}

function executeUpdateRule(action) {
  if (!action.rule_type || !action.key) return { ok: false, message: 'update_rule requires rule_type and key.' };
  try {
    var method = action.delete ? 'DELETE' : 'POST';
    UrlFetchApp.fetch(HUB_URL + '/api/rules', {
      method: method, contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({ type: action.rule_type, key: action.key, value: action.value || 'true', notes: action.notes || '' }),
      muteHttpExceptions: true
    });
    if (action.rule_type === 'blocked_domain') CacheService.getScriptCache().remove('blocked_domains');
    return { ok: true, message: (action.delete ? '✅ Deleted rule: ' : '✅ Updated rule: ') + action.rule_type + '/' + action.key };
  } catch(e) { return { ok: false, message: 'Rule update error: ' + e }; }
}

function executeMulti(action, threadId, subject, fromH) {
  var actions = action.actions || [];
  var messages = [];
  for (var i = 0; i < actions.length; i++) {
    var result = executeAction(actions[i], threadId, subject, fromH);
    messages.push(result.message);
    if (!result.ok) {
      messages.push('⛔ Stopped at step ' + (i + 1));
      return { ok: false, message: messages.join('\n') };
    }
  }
  return { ok: true, message: messages.join('\n') };
}

function hubLog(eventType, summary, details) {
  try {
    UrlFetchApp.fetch(HUB_URL + '/api/logs', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({ app_name: 'email_automation', event_type: eventType,
                                summary: summary, details: details || null }),
      muteHttpExceptions: true,
    });
  } catch(e) { Logger.log('hubLog error: ' + e); }
}

function hubPostDraft(threadId, mpn, sender, subject, draftContent, gmailDraftId, adviceText) {
  var content = draftContent || '';
  if (adviceText) content += '\n\n[ADVICE_STORED]:' + adviceText;
  if (gmailDraftId) content += '\n\n[GMAIL_DRAFT:' + gmailDraftId + ']';
  try {
    UrlFetchApp.fetch(HUB_URL + '/api/drafts', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({ thread_id: threadId, mpn: mpn, sender: sender,
                                subject: subject, draft_content: content }),
      muteHttpExceptions: true,
    });
  } catch(e) { Logger.log('hubPostDraft error: ' + e); }
}

function hubLearn(feedback, draftBody, correctedBody, threadId, subject, sender, mpn, action) {
  try {
    UrlFetchApp.fetch(HUB_URL + '/api/learn', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({
        feedback: feedback || '',
        draft_body: draftBody || '',
        corrected_body: correctedBody || '',
        thread_id: threadId || '',
        subject: subject || '',
        sender: sender || '',
        mpn: mpn || '',
        action: action || '',
      }),
      muteHttpExceptions: true,
    });
  } catch(e) { Logger.log('hubLearn error: ' + e); }
}

function hubPatchEntry(id, payload) {
  try {
    UrlFetchApp.fetch(HUB_URL + '/api/drafts/' + id, {
      method: 'PATCH', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
  } catch(e) { Logger.log('hubPatchEntry error: ' + e); }
}

function archiveBlockedDomains() {
  var BLOCKED_DOMAINS = getBlockedDomains();
  BLOCKED_DOMAINS.forEach(function(domain) {
    var threads = GmailApp.search('in:inbox from:' + domain, 0, 50);
    if (!threads.length) return;
    threads.forEach(function(thread) {
      thread.moveToArchive();
      Logger.log('Blocked & archived: ' + domain + ' | ' + thread.getFirstMessageSubject());
    });
    hubLog('run', 'Blocked & archived ' + threads.length + ' email(s) from ' + domain);
  });
}

function getRemoteConfig() {
  try {
    var resp = UrlFetchApp.fetch(HUB_URL + '/api/configs/email_automation', {
      method: 'get',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      muteHttpExceptions: true,
    });
    if (resp.getResponseCode() !== 200) return {};
    var row = JSON.parse(resp.getContentText());
    return row.config ? JSON.parse(row.config) : {};
  } catch(e) { Logger.log('getRemoteConfig error: ' + e); return {}; }
}

function applyRemoteConfig(cfg) {
  if (!cfg || typeof cfg !== 'object') return;
  if (cfg.MSG_NEED_TP_500)  MSG_NEED_TP_500  = cfg.MSG_NEED_TP_500;
  if (cfg.MSG_NEED_TP_2000) MSG_NEED_TP_2000 = cfg.MSG_NEED_TP_2000;
  if (cfg.MSG_CHECKING)     MSG_CHECKING     = cfg.MSG_CHECKING;
  if (cfg.MSG_BILL)         MSG_BILL         = cfg.MSG_BILL;
  if (cfg.DAVID_EMAIL)      DAVID_EMAIL      = cfg.DAVID_EMAIL;
  if (cfg.BILL_EMAIL)       BILL_EMAIL       = cfg.BILL_EMAIL;
  if (cfg.DEB_EMAIL)        DEB_EMAIL        = cfg.DEB_EMAIL;
}

// ── Gmail REST API — threaded draft creation ──────────────────

function getRFC2822MessageId(gmailMsgId) {
  var url = 'https://gmail.googleapis.com/gmail/v1/users/me/messages/' + gmailMsgId + '?format=metadata&metadataHeaders=Message-ID';
  var resp = UrlFetchApp.fetch(url, { method: 'GET', headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  var data = JSON.parse(resp.getContentText());
  if (data.payload && data.payload.headers) {
    for (var i = 0; i < data.payload.headers.length; i++) {
      if (data.payload.headers[i].name === 'Message-ID') return data.payload.headers[i].value;
    }
  }
  return null;
}

function createThreadedDraft(toEmail, subject, htmlBody, replyToGmailMsgId, threadId, ccEmail) {
  var rfcId = getRFC2822MessageId(replyToGmailMsgId);
  var lines = ['From: John Fluman <' + JOHN_EMAIL + '>', 'To: ' + toEmail];
  if (ccEmail) lines.push('Cc: ' + ccEmail);
  lines.push('Subject: ' + subject);
  if (rfcId) { lines.push('In-Reply-To: ' + rfcId); lines.push('References: ' + rfcId); }
  lines.push('MIME-Version: 1.0');
  lines.push('Content-Type: text/html; charset=UTF-8');
  lines.push('');
  lines.push(htmlBody);
  var encoded = Utilities.base64EncodeWebSafe(lines.join('\r\n'));
  var url = 'https://gmail.googleapis.com/gmail/v1/users/me/drafts';
  var resp = UrlFetchApp.fetch(url, {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(), 'Content-Type': 'application/json' },
    payload: JSON.stringify({ message: { raw: encoded, threadId: threadId } }),
    muteHttpExceptions: true
  });
  var result = JSON.parse(resp.getContentText());
  if (result.error) { Logger.log('API draft error: ' + JSON.stringify(result.error)); return null; }
  Logger.log('Draft created | To: ' + toEmail + ' | ' + subject);
  return result.id;
}

function sendThreadedReply(toEmail, subject, htmlBody, replyToGmailMsgId, threadId) {
  var rfcId = getRFC2822MessageId(replyToGmailMsgId);
  var lines = ['From: John Fluman <' + JOHN_EMAIL + '>', 'To: ' + toEmail, 'Subject: ' + subject];
  if (rfcId) { lines.push('In-Reply-To: ' + rfcId); lines.push('References: ' + rfcId); }
  lines.push('MIME-Version: 1.0', 'Content-Type: text/html; charset=UTF-8', '', htmlBody);
  var encoded = Utilities.base64EncodeWebSafe(lines.join('\r\n'));
  var url = 'https://gmail.googleapis.com/gmail/v1/users/me/messages/send';
  var resp = UrlFetchApp.fetch(url, {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(), 'Content-Type': 'application/json' },
    payload: JSON.stringify({ raw: encoded, threadId: threadId }),
    muteHttpExceptions: true
  });
  var result = JSON.parse(resp.getContentText());
  if (result.error) { Logger.log('API send error: ' + JSON.stringify(result.error)); return null; }
  Logger.log('Auto-sent reply | To: ' + toEmail + ' | ' + subject);
  return result.id;
}

// ── Signature + HTML builder ──────────────────────────────────

function getSignatureHTML() {
  return '<br><br><div><b><span style="color:rgb(31,73,125);font-family:Tahoma,sans-serif;font-size:10pt">Regards,</span></b></div>'
    + '<div><b><span style="color:rgb(31,73,125);font-family:Tahoma,sans-serif;font-size:10pt">John Fluman</span></b></div>'
    + '<div><b><span style="color:rgb(31,73,125);font-family:Arial,sans-serif;font-size:8pt">Intransit Technologies</span></b></div>'
    + '<div><a href="mailto:john.fluman@intransittech.com" style="font-family:Calibri;font-size:8pt">john.fluman@intransittech.com</a></div>'
    + '<div><i><span style="color:gray;font-family:Arial,sans-serif;font-size:7.5pt">An ISO 9001 Certified Company</span></i></div>'
    + '<div><span style="color:rgb(31,73,125);font-family:Tahoma,sans-serif;font-size:8pt">Toll (877) 677-5868 x101 - Local (949) 481-7935 x101</span></div>'
    + '<br><div><span style="color:rgb(166,166,166);font-family:Calibri,sans-serif;font-size:8pt">The information contained in this communication and its attachment(s) is intended only for the use of the individual to whom it is addressed and may contain information that is privileged, confidential, or exempt from disclosure. If the reader of this message is not the intended recipient, you are hereby notified that any dissemination, distribution, or copying of this communication is strictly prohibited. If you have received this communication in error, please notify <a href="mailto:john.fluman@intransittech.com" style="font-family:Calibri;font-size:8pt">john.fluman@intransittech.com</a> and delete the communication without retaining any copies. Thank you.</span></div>';
}

function buildSimpleHTML(bodyText) {
  return '<div dir="ltr">' + bodyText + getSignatureHTML() + '</div>';
}

function buildDraftHTML(replyText, originalMessage) {
  var sig = getSignatureHTML();
  var origDate = Utilities.formatDate(originalMessage.getDate(), Session.getScriptTimeZone(), 'EEE, MMM d, yyyy, h:mm a');
  var origFrom = originalMessage.getFrom();
  var origBody = originalMessage.getBody() || originalMessage.getPlainBody().replace(/\n/g, '<br>');
  var quoted = '<br><div class="gmail_quote"><div dir="ltr" class="gmail_attr">On ' + origDate + ', ' + origFrom + ' wrote:<br></div>'
    + '<blockquote class="gmail_quote" style="margin:0px 0px 0px 0.8ex;border-left:1px solid rgb(204,204,204);padding-left:1ex">'
    + origBody + '</blockquote></div>';
  return '<div dir="ltr">' + replyText + sig + quoted + '</div>';
}

// ── Utility ───────────────────────────────────────────────────

function normalize(pn) {
  return String(pn).trim().replace(/[-.:\/()\\s*+\\#_,]/g, '').toLowerCase();
}

function getOrCreateDeletedSheet(ss) {
  var sheet = ss.getSheetByName(DELETED_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DELETED_SHEET_NAME);
    sheet.appendRow(['Date Deleted','FullPartNumber','Man','DC','QTY','Notes','Source Email Subject']);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,7).setFontWeight('bold');
  }
  return sheet;
}

function logDeletion(deletedSheet, rowData, emailSubject) {
  deletedSheet.appendRow([new Date(), rowData[0], rowData[1], rowData[2], rowData[3], rowData[4], emailSubject||'']);
}

function sendReviewEmail(partNumber, emailSubject, matches) {
  var body = 'Notice for part: ' + partNumber + '\nSubject: "' + emailSubject + '"\n\n';
  if (matches.length > 0) {
    matches.forEach(function(m,i) { body += (i+1) + '. MPN: ' + m.data[0] + ' | QTY: ' + m.data[3] + '\n'; });
    body += '\nhttps://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID;
  } else { body += 'No match.\nhttps://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID; }
  GmailApp.sendEmail(NOTIFY_EMAIL, 'OEM EXCESS: Review needed for MPN ' + partNumber, body);
}


// ── Sheet searches ────────────────────────────────────────────

function searchOEMExcess(mpn) {
  if (!mpn) return [];
  var searchNorm = normalize(mpn);
  var data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME).getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var cellNorm = normalize(String(data[i][0]));
    var reverseOk = searchNorm.startsWith(cellNorm) && cellNorm.length >= Math.ceil(searchNorm.length * 0.75);
    if (cellNorm.length >= 3 && (cellNorm === searchNorm || cellNorm.startsWith(searchNorm) || reverseOk)) {
      results.push({ row: i+1, mpn: data[i][0], man: data[i][1], dc: data[i][2], qty: data[i][3], notes: data[i][4] });
    }
  }
  if (results.length) { Logger.log('OEM FOUND ' + results.length + ': ' + mpn); results.forEach(function(r){Logger.log('  Row '+r.row+' | QTY:'+r.qty+' | '+r.notes);}); }
  else Logger.log('OEM NOT FOUND: ' + mpn);
  return results;
}

function searchInStock(mpn) {
  if (!mpn) return [];
  var searchNorm = normalize(mpn);
  var data = SpreadsheetApp.openById(IN_STOCK_ID).getSheets()[0].getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var cellNorm = normalize(String(data[i][0]));
    if (cellNorm.length >= 3 && (cellNorm === searchNorm || cellNorm.startsWith(searchNorm) || searchNorm.startsWith(cellNorm))) {
      results.push({ row: i+1, mpn: data[i][0], man: data[i][1], dc: data[i][2], qty: data[i][3], notes: data[i][4] });
    }
  }
  if (results.length) Logger.log('IN STOCK FOUND ' + results.length + ': ' + mpn);
  else Logger.log('IN STOCK NOT FOUND: ' + mpn);
  return results;
}

function searchStanSheet(mpn) {
  if (!mpn) return [];
  var searchNorm = normalize(mpn);
  var data = SpreadsheetApp.openById(STAN_SHEET_ID).getSheets()[0].getDataRange().getValues();
  var results = [];
  for (var i = 2; i < data.length; i++) {
    var cellNorm = normalize(String(data[i][4]));
    var reverseOkStan = searchNorm.startsWith(cellNorm) && cellNorm.length >= Math.ceil(searchNorm.length * 0.75);
    if (cellNorm.length >= 3 && (cellNorm === searchNorm || cellNorm.startsWith(searchNorm) || reverseOkStan)) {
      results.push({ row: i+1, status: data[i][0], colB: data[i][1], colC: data[i][2], date: data[i][3], mpn: data[i][4], country: data[i][5] });
    }
  }
  if (results.length) { Logger.log('STAN FOUND ' + results.length + ': ' + mpn); results.forEach(function(r){Logger.log('  Row '+r.row+' | Status:'+r.status+' | ColB:'+r.colB);}); }
  else Logger.log('STAN NOT FOUND: ' + mpn);
  return results;
}

function checkForteForMPN(mpn, days) {
  if (!mpn) return [];
  var data = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0].getDataRange().getValues();
  var cutoff = new Date(); cutoff.setDate(cutoff.getDate() - (days||60));
  var matches = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === mpn.trim().toLowerCase()) {
      var status = String(data[i][FORTE_STATUS_COL]).trim();
      var recent = new Date(data[i][0]) >= cutoff;
      matches.push({ row: i+1, date: data[i][0], status: status, recent: recent, colH: String(data[i][7]).trim(), colI: String(data[i][8]).trim() });
    }
  }
  if (!matches.length) Logger.log('FORTE NOT FOUND: ' + mpn);
  return matches;
}

function buildForteHistory(mpn) {
  if (!mpn) return '';
  try {
    var data = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0].getDataRange().getValues();
    var entries = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() !== mpn.trim().toLowerCase()) continue;
      var rawDate = data[i][0];
      var qty     = String(data[i][2] || '').trim();
      var tp      = String(data[i][3] || '').trim();
      var quoted  = String(data[i][7] || '').trim();   // col H: John Quoted
      var notes   = String(data[i][8] || '').trim();   // col I: Notes
      var status  = String(data[i][10] || '').trim();  // col K: Status
      var dateStr = rawDate ? Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), 'M/d/yyyy') : '?';
      var line = dateStr;
      if (qty)    line += ' | Qty: ' + qty;
      if (tp)     line += ' | TP: ' + tp;
      if (quoted) line += ' | Quoted: ' + quoted;
      if (status && status.toLowerCase() !== 'open') line += ' | ' + status;
      if (notes)  line += ' | ' + notes;
      entries.push({ date: rawDate ? new Date(rawDate) : new Date(0), text: line });
    }
    // Most recent first
    entries.sort(function(a, b) { return b.date - a.date; });
    return entries.map(function(e) { return e.text; }).join('\n');
  } catch(e) {
    Logger.log('buildForteHistory error: ' + e);
    return '';
  }
}

function addToForteSheet(mpn, qty, targetPrice, country, historyNote) {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  // Build prior-entry history BEFORE appending (so the new row sees all existing rows)
  var priorHistory = buildForteHistory(mpn);
  var finalHistory = '';
  if (priorHistory) finalHistory = priorHistory;
  if (historyNote) finalHistory = finalHistory ? finalHistory + '\n---\n' + historyNote : historyNote;
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
  var nextRow = sheet.getLastRow() + 1;
  var potentialFormula = '=C' + nextRow + '*D' + nextRow;
  sheet.appendRow([today, mpn, qty||'', targetPrice||'', '', country||'', potentialFormula, '', '', finalHistory, 'Open']);
  Logger.log('Added to Forte: ' + mpn + (priorHistory ? ' [history populated]' : ''));
}

function updateForteSheet(mpn, customDate) {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var today = customDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
  var newStatus = 'NO STK - ' + today;
  var updated = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === mpn.trim().toLowerCase()
      && String(data[i][FORTE_STATUS_COL]).trim().toUpperCase() !== 'CLOSED') {
      var cell = sheet.getRange(i+1, FORTE_STATUS_COL+1);
      cell.clearDataValidations(); cell.setValue(newStatus);
      cell.setBackground('#000000'); cell.setFontColor('#FFFFFF'); cell.setFontWeight('bold');
      updated++;
    }
  }
  Logger.log('Forte NO STK ' + mpn + ': updated=' + updated);
}

function addToStanSheet(mpn, country, qty, tp) {
  var existing = searchStanSheet(mpn);
  if (existing.length > 0) {
    Logger.log('Stan sheet skip — already exists: ' + mpn);
    return;
  }
  var sheet = SpreadsheetApp.openById(STAN_SHEET_ID).getSheets()[0];
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
  sheet.appendRow(['', '', '', today, mpn, country||'USA', qty||'', tp||'']);
  Logger.log('Stan sheet row added: '+mpn+' | '+country+' | QTY:'+qty+' | TP:'+tp);
}

// ── OEM EXCESS delete ─────────────────────────────────────────

function findMatches(data, partNumber) {
  var exact = [], fuzzy = [], sn = normalize(partNumber);
  for (var i = 1; i < data.length; i++) {
    var cr = String(data[i][0]).trim(), cn = normalize(cr);
    if (cr.toLowerCase() === partNumber.trim().toLowerCase()) exact.push({row:i+1,data:data[i]});
    else if (cn === sn) fuzzy.push({row:i+1,data:data[i],type:'stripped'});
    else if (cn.length >= 3 && (cn.startsWith(sn) || sn.startsWith(cn))) fuzzy.push({row:i+1,data:data[i],type:'prefix'});
  }
  return {exact:exact,fuzzy:fuzzy};
}

function deletePart(partNumber, emailSubject) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var deletedSheet = getOrCreateDeletedSheet(ss);
  var data = mainSheet.getDataRange().getValues();
  var r = findMatches(data, partNumber), exact = r.exact, fuzzy = r.fuzzy;
  var noStkStamp = 'NO STK ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
  if (exact.length===1){mainSheet.getRange(exact[0].row,5).setValue(noStkStamp);logDeletion(deletedSheet,exact[0].data,emailSubject);mainSheet.deleteRow(exact[0].row);return 'DELETED';}
  if (exact.length>1){sendReviewEmail(partNumber,emailSubject,exact);return 'MULTIPLE';}
  if (!exact.length&&fuzzy.length===1&&fuzzy[0].type==='stripped'){mainSheet.getRange(fuzzy[0].row,5).setValue(noStkStamp);logDeletion(deletedSheet,fuzzy[0].data,emailSubject);mainSheet.deleteRow(fuzzy[0].row);return 'FUZZY';}
  // Single prefix fuzzy match with ≤3-char suffix diff (e.g. W25Q256JWEIM → W25Q256JWEIMS) — safe to auto-delete
  if (!exact.length&&fuzzy.length===1&&fuzzy[0].type==='prefix'){var _pdiff=Math.abs(String(fuzzy[0].data[0]).trim().length-partNumber.trim().length);if(_pdiff<=3){mainSheet.getRange(fuzzy[0].row,5).setValue(noStkStamp);logDeletion(deletedSheet,fuzzy[0].data,emailSubject);mainSheet.deleteRow(fuzzy[0].row);return 'FUZZY';}}
  if (fuzzy.length){sendReviewEmail(partNumber,emailSubject,fuzzy);return 'FUZZY_REVIEW';}
  sendReviewEmail(partNumber,emailSubject,[]);return 'NOT_FOUND';
}

// ── Trigger 1 — David no-stock emails ────────────────────────
function checkDavidNoStockEmails() {
  var _cfg = getRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) return;
  // Query relies on in:inbox only — do NOT filter by INCOMING_LABEL here.
  // John's Gmail filter auto-labels ALL David emails with oem-nostock-seen, so
  // -label:oem-nostock-seen would exclude every David email and return 0 forever.
  var query = 'from:' + DAVID_EMAIL + ' in:inbox';
  var threads = GmailApp.search(query, 0, 20);
  hubLog('run', 'checkDavidNoStockEmails: ' + threads.length + ' thread(s)');
  if (!threads.length) return;
  var noStkKeywords = ['no stk', 'no stock', 'stk sold', 'stock sold', 'cant find', 'cant share', 'cannot find', 'removed', 'no inventory'];
  var label = GmailApp.getUserLabelByName(INCOMING_LABEL) || GmailApp.createLabel(INCOMING_LABEL);
  threads.forEach(function(thread) {
    var msg = thread.getMessages()[thread.getMessageCount() - 1];
    var subjectLower = msg.getSubject().toLowerCase();
    var bodySnippet = msg.getPlainBody().toLowerCase().substring(0, 300);
    var isNoStk = noStkKeywords.some(function(kw) { return subjectLower.indexOf(kw) >= 0 || bodySnippet.indexOf(kw) >= 0; });
    if (!isNoStk) return;
    try { processThread(thread); } catch(e) { hubLog('error', 'checkDavidNoStockEmails processThread error: ' + e, {}); }
    thread.addLabel(label);
    thread.moveToArchive();
  });
}


function sendDailyCostReport() {
  try {
    var resp = UrlFetchApp.fetch(HUB_URL + '/api/cost-report?days=1', {
      method: 'GET', headers: { Authorization: 'Bearer ' + HUB_SECRET },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) { Logger.log('cost-report HTTP ' + resp.getResponseCode()); return; }
    var data = JSON.parse(resp.getContentText());
    var rows = data.rows || [];
    var total = data.total_cost_usd || 0;

    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy');
    var lines = ['Intransit Hub — Daily API Cost Report', 'Date: ' + today, ''];

    if (!rows.length) {
      lines.push('No API calls recorded in the last 24 hours.');
    } else {
      lines.push('TOTAL COST: $' + total.toFixed(4));
      lines.push('');
      lines.push('Breakdown:');
      rows.forEach(function(r) {
        var modelShort = r.model.indexOf('haiku') >= 0 ? 'Haiku' : 'Sonnet';
        lines.push('  ' + r.endpoint + ' [' + modelShort + ']');
        lines.push('    Calls:  ' + r.calls);
        lines.push('    Tokens: ' + Number(r.total_input).toLocaleString() + ' in / ' + Number(r.total_output).toLocaleString() + ' out');
        lines.push('    Cost:   $' + (r.total_cost || 0).toFixed(4));
        lines.push('');
      });
    }
    lines.push('—');
    lines.push('Intransit Hub Automation');

    var subject = 'Intransit Hub — Daily Cost ($' + total.toFixed(4) + ') ' + today;
    GmailApp.sendEmail('john.fluman@intransittech.com', subject, lines.join('\n'));
    Logger.log('Daily cost report sent: $' + total.toFixed(4));
  } catch(e) {
    Logger.log('sendDailyCostReport error: ' + e);
  }
}

// ── Slim worker bridge ───────────────────────────────────────

function extractBuyerEmail(fromRaw) {
  if (!fromRaw) return '';
  var m = fromRaw.match(/<([^>]+)>/);
  return m ? m[1].trim() : fromRaw.trim();
}

function callWorker(payload) {
  try {
    var resp = UrlFetchApp.fetch(HUB_URL + '/api/email-agent', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    return resp.getResponseCode() === 200 ? JSON.parse(resp.getContentText()) : null;
  } catch(e) { Logger.log('callWorker error: ' + e); return null; }
}

function deleteOemRow(row) {
  if (!row) return;
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
    var deletedSheet = getOrCreateDeletedSheet(ss);
    var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    logDeletion(deletedSheet, rowData, 'worker remove_oem');
    sheet.deleteRow(row);
    Logger.log('deleteOemRow: row ' + row);
  } catch(e) { Logger.log('deleteOemRow error: ' + e); }
}

function executeDecision(decision, thread) {
  if (!decision || !decision.action) return;
  var action = decision.action;
  var messages = thread.getMessages();
  var lastMsg = messages[messages.length - 1];
  var threadId = thread.getId();
  var subject = thread.getFirstMessageSubject();
  if (action === 'no_action') { thread.markRead(); return; }
  if (action === 'no_bid' && !decision.draft_body) { thread.moveToArchive(); return; }
  if (action === 'remove_oem') {
    if (decision.oem_delete_row) deleteOemRow(decision.oem_delete_row);
    if (decision.mpn) updateForteSheet(decision.mpn);
  }
  if (action === 'add_to_stan' && decision.forte_entry) {
    var fe0 = decision.forte_entry;
    addToStanSheet(fe0.mpn, fe0.country || 'USA', fe0.qty || '', fe0.target_price || '');
  }
  if (action === 'david_nostock') {
    if (decision.mpn) {
      var dRes = deletePart(decision.mpn, subject);
      hubLog('run', 'david_nostock: deletePart ' + decision.mpn + ' → ' + dRes, {});
      if (dRes === 'DELETED' || dRes === 'FUZZY') updateForteSheet(decision.mpn);
    }
  }
  if (decision.forte_entry) {
    var fe = decision.forte_entry;
    if (fe.mpn && fe.qty) {
      var existing = checkForteForMPN(fe.mpn, 60);
      var hasRecent = existing.some(function(r){ return r.recent && r.status.toLowerCase() !== 'closed'; });
      if (!hasRecent) addToForteSheet(fe.mpn, fe.qty, fe.target_price || '', fe.country || '', '');
      else hubLog('run', 'Forte 60-day skip: ' + fe.mpn);
    }
  }
  if (decision.draft_body) {
    var replyTo = decision.buyer_email || extractBuyerEmail(lastMsg.getFrom());
    if (!replyTo || replyTo.indexOf('intransittech.com') >= 0) {
      for (var i = messages.length - 1; i >= 0; i--) {
        var sf = messages[i].getFrom();
        if (sf.indexOf('intransittech.com') < 0) { replyTo = extractBuyerEmail(sf); break; }
      }
    }
    if (!replyTo || replyTo.indexOf('intransittech.com') >= 0) {
      hubLog('error', 'SAFETY ABORT: no external replyTo for ' + (decision.mpn || '?'));
      return;
    }
    var bodyText = decision.draft_body.replace(/\s*(Best regards?,?|Regards?,?|Sincerely,?)\s*$/i, '').trim();
    var origMsg = null;
    for (var j = 0; j < messages.length; j++) {
      if (messages[j].getFrom().indexOf(JOHN_EMAIL) < 0 && messages[j].getFrom().indexOf('intransittech') < 0) {
        origMsg = messages[j]; break;
      }
    }
    var ccEmail = (action === 'bill_handle') ? BILL_EMAIL : null;
    var htmlBody = origMsg ? buildDraftHTML(bodyText, origMsg) : buildSimpleHTML(bodyText);
    var draftId = createThreadedDraft(replyTo, 'Re: ' + subject, htmlBody, lastMsg.getId(), threadId, ccEmail);
    hubPostDraft(threadId, decision.mpn || '', replyTo, 'Re: ' + subject, bodyText, draftId, decision.reasoning || action);
    hubLog('draft_created', 'Worker draft (' + action + '): ' + (decision.mpn || '?'), {mpn: decision.mpn, type: action});
    if (draftId && decision.id) {
      try {
        UrlFetchApp.fetch(HUB_URL + '/api/agent-decisions/' + decision.id, {
          method: 'PATCH', contentType: 'application/json',
          headers: { Authorization: 'Bearer ' + HUB_SECRET },
          payload: JSON.stringify({ status: 'drafted', gmail_draft_id: draftId }),
          muteHttpExceptions: true
        });
      } catch(e) {}
    }
  }
  if (action === 'no_bid') thread.moveToArchive();
  if (decision._corrected_from) {
    try {
      GmailApp.sendEmail(NOTIFY_EMAIL,
        'Bug Auto-Corrected: [' + decision._corrected_from + '->' + action + '] ' + subject,
        'MPN: ' + (decision.mpn || '?') + '\nOriginal: ' + decision._corrected_from + '\nCorrected to: ' + action + '\nReason: ' + (decision._correction_reason || '?')
      );
    } catch(e) {}
  }
}

// Parses a netCOMPONENTS HTML table to extract QtyReq and TgtPrice.
// Returns { qtyReq, tgtPrice } or null if not found / not a netCOMPONENTS email.
function extractNetcompRFQ(messages) {
  var msg = messages[0];
  var html = msg.getBody() || '';
  var htmlLower = html.toLowerCase();
  var fromLower = msg.getFrom().toLowerCase();
  if (fromLower.indexOf('netcomponents') < 0 && htmlLower.indexOf('netcomponents') < 0) return null;
  if (htmlLower.indexOf('qty') < 0 && htmlLower.indexOf('quantity') < 0) return null;

  var rows = html.split(/<tr[^>]*>/i);
  var qtyCol = -1, tpCol = -1, foundHeader = false;
  for (var r = 0; r < rows.length; r++) {
    var rowHtml = rows[r];
    var cells = rowHtml.match(/<t[hd][^>]*>([\s\S]*?)<\/t[hd]>/gi) || [];
    if (!cells.length) continue;
    var vals = cells.map(function(c) {
      return c.replace(/<[^>]+>/g, '').replace(/&nbsp;/gi, ' ').replace(/&#\d+;/g, '').trim();
    });
    if (!foundHeader && rowHtml.indexOf('<th') >= 0) {
      vals.forEach(function(v, i) {
        if (/^qty/i.test(v)) { if (qtyCol < 0 || /qtyreq/i.test(v)) qtyCol = i; }
        if (/target\s*price|tgt\s*price|tgtprice/i.test(v)) tpCol = i;
      });
      if (qtyCol >= 0 && tpCol >= 0) foundHeader = true;
      continue;
    }
    if (foundHeader && vals.length > Math.max(qtyCol, tpCol)) {
      var qty = parseInt((vals[qtyCol] || '').replace(/,/g, ''), 10);
      var tp = parseFloat((vals[tpCol] || '').replace(/[$,\s]/g, ''));
      if (!isNaN(qty) && qty > 0 && !isNaN(tp) && tp > 0) return { qtyReq: qty, tgtPrice: tp };
    }
  }
  return null;
}

function processThread(thread) {
  var messages = thread.getMessages();
  var lastMsg = messages[messages.length - 1];
  var subject = thread.getFirstMessageSubject();
  var parts = ['Subject: ' + subject, ''];
  messages.forEach(function(m, i) {
    var body = (m.getPlainBody() || '').split('\n').filter(function(ln){ return ln.charAt(0) !== '>'; }).join('\n').trim();
    parts.push('--- Msg ' + (i+1) + ' | From: ' + m.getFrom() + ' ---');
    parts.push(body.substring(0, 2000));
  });
  var content = parts.join('\n');
  if (content.length > 8000) content = content.substring(0, 8000) + '\n[truncated]';

  var parsedRFQ = extractNetcompRFQ(messages);
  if (parsedRFQ) {
    content = '[PARSED_RFQ: QtyReq=' + parsedRFQ.qtyReq + ', TgtPrice=' + parsedRFQ.tgtPrice + ']\n' + content;
  }

  var mpnHint = extractMPNFromSubject(subject) || extractMPN(subject);
  var priorQuotes = mpnHint ? getRecentSentQuotesFull(mpnHint, 5) : 'None found';

  var payload = {
    thread_id:       thread.getId(),
    last_message_id: lastMsg.getId(),
    subject:         subject,
    sender:          extractBuyerEmail(lastMsg.getFrom()),
    thread_content:  content,
    current_labels:  thread.getLabels().map(function(l){ return l.getName(); }),
    prior_quotes:    priorQuotes
  };
  var decision = callWorker(payload);
  if (decision) executeDecision(decision, thread);
  return decision;
}

function runEmailScan() {
  var _cfg = getRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) { hubLog('run', 'runEmailScan: disabled'); return; }
  try { archiveBlockedDomains(); } catch(e) {}
  var BLOCKED_DOMAINS = getBlockedDomains();
  var blockFilter = BLOCKED_DOMAINS.map(function(d){ return '-from:' + d; }).join(' ');

  try { checkDavidNoStockEmails(); } catch(e) { hubLog('error', 'checkDavidNoStockEmails crashed: ' + e, {}); }

  var rfqLabel = GmailApp.getUserLabelByName('oem-rfq-incoming-processed') || GmailApp.createLabel('oem-rfq-incoming-processed');
  var rfqQ = 'in:inbox (to:rfq@intransittech.com OR deliveredto:rfq@intransittech.com OR subject:rfq OR subject:"please quote" OR subject:"request for quote" OR subject:"request for quotation" OR ((to:john.fluman@intransittech.com OR deliveredto:john.fluman@intransittech.com) ("quotation" OR "best price" OR "netcomponents" OR "looking for" OR "quote your stock" OR "can you quote"))) -from:intransittech.com -from:david@fortetechno.com -from:steve@fortetechno.com -label:oem-rfq-incoming-processed ' + blockFilter;
  GmailApp.search(rfqQ, 0, 10).forEach(function(t) {
    t.addLabel(rfqLabel);
    var msgs = t.getMessages();
    if (msgs.some(function(m){ return m.getFrom().indexOf(JOHN_EMAIL) >= 0; })) return;
    if (msgs[msgs.length-1].getFrom().indexOf('intransittech.com') >= 0) return;
    var allTo = (msgs[0].getTo() + ',' + (msgs[0].getCc() || '')).toLowerCase();
    var validT = ['rfq@intransittech.com','john.fluman@intransittech.com','sales@intransittech.com','websiterfq@intransittech.com'];
    if (allTo.indexOf('intransittech.com') >= 0 && !validT.some(function(v){ return allTo.indexOf(v) >= 0; })) return;
    try { processThread(t); } catch(e) { hubLog('error', 'runEmailScan rfqQ processThread error: ' + e, {}); }
  });

  var tpLabel = GmailApp.getUserLabelByName('oem-tp-processed') || GmailApp.createLabel('oem-tp-processed');
  var tpQ = 'in:inbox label:oem-rfq-incoming-processed -label:oem-tp-processed newer_than:30d ' + blockFilter;
  GmailApp.search(tpQ, 0, 20).forEach(function(t) {
    var msgs = t.getMessages();
    if (msgs[msgs.length-1].getFrom().indexOf(JOHN_EMAIL) >= 0) return;
    var buyerCount = 0;
    msgs.forEach(function(m){ if (m.getFrom().indexOf(JOHN_EMAIL) < 0 && m.getFrom().indexOf('intransittech') < 0) buyerCount++; });
    if (buyerCount < 2) return;
    t.addLabel(tpLabel);
    try { processThread(t); } catch(e) { hubLog('error', 'runEmailScan tpQ processThread error: ' + e, {}); }
  });

  var agentLabel = GmailApp.getUserLabelByName(AGENT_LABEL) || GmailApp.createLabel(AGENT_LABEL);
  var agentQ = 'in:inbox -label:' + AGENT_LABEL + ' -label:oem-rfq-incoming-processed newer_than:2d -from:' + DAVID_EMAIL + ' ' + blockFilter;
  GmailApp.search(agentQ, 0, 15).forEach(function(t) {
    var from = t.getMessages()[0].getFrom().toLowerCase();
    if (from.indexOf('intransittech.com') >= 0 || from.indexOf('fortetechno.com') >= 0) { t.addLabel(agentLabel); return; }
    var liveLabels = t.getLabels().map(function(l){ return l.getName(); });
    if (liveLabels.indexOf('oem-rfq-incoming-processed') >= 0) { t.addLabel(agentLabel); return; }
    t.addLabel(agentLabel);
    t.addLabel(rfqLabel);
    try { processThread(t); } catch(e) { hubLog('error', 'runEmailScan agentQ processThread error: ' + e, {}); }
  });

  hubLog('run', 'runEmailScan complete');
}


// ── Trigger 6 — Payment Advice → forward to Deb ──────────────
function checkInboxForPaymentAdvice() {
  var _cfg = getRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) { hubLog('run', 'checkInboxForPaymentAdvice: disabled via hub config'); return; }
  var query = 'in:inbox subject:"payment advice" -label:oem-payment-forwarded';
  var threads = GmailApp.search(query,0,10);
  hubLog('run', 'checkInboxForPaymentAdvice: ' + threads.length + ' thread(s)');
  if (!threads.length) return;
  var label = GmailApp.getUserLabelByName('oem-payment-forwarded')||GmailApp.createLabel('oem-payment-forwarded');
  threads.forEach(function(thread) {
    var threadLabelNames = thread.getLabels().map(function(l){return l.getName();});
    if (threadLabelNames.indexOf('oem-payment-forwarded') >= 0) {
      Logger.log('Payment advice already forwarded (label present): ' + thread.getFirstMessageSubject());
      return;
    }
    var messages = thread.getMessages();
    var firstMsg = messages[0];
    if (firstMsg.getFrom().toLowerCase().indexOf('intransittech.com') >= 0) {
      thread.addLabel(label);
      Logger.log('Payment advice skipped — internal sender: ' + firstMsg.getFrom());
      return;
    }
    var recipients = (firstMsg.getTo() + ',' + (firstMsg.getCc() || '')).toLowerCase();
    if (recipients.indexOf(DEB_EMAIL.toLowerCase()) >= 0) {
      thread.addLabel(label);
      Logger.log('Payment advice skipped — Deb already a recipient: ' + firstMsg.getSubject());
      return;
    }
    try {
      firstMsg.forward(DEB_EMAIL);
      hubLog('run', 'Payment advice forwarded: ' + firstMsg.getSubject());
      Logger.log('Payment advice forwarded to Deb: ' + firstMsg.getSubject());
    } catch(e) {
      Logger.log('Forward error: ' + e.toString());
    }
    thread.addLabel(label);
  });
}

// ── Web App ───────────────────────────────────────────────────
function searchForteSheet(mpn) {
  if (!mpn) return [];
  var data = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0].getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === mpn.trim().toLowerCase()) {
      results.push({
        row: i+1, date: data[i][0], mpn: data[i][1], qty: data[i][2],
        buyerTP: data[i][3], johnBuy: data[i][4], country: data[i][5],
        potential: data[i][6], johnQuoted: data[i][7], notes: data[i][8],
        history: data[i][9], status: data[i][10]
      });
    }
  }
  return results;
}

function doGet(e) {
  var SECRET = 'baSDJ#444FE&8';
  if (e.parameter.key!==SECRET) return ContentService.createTextOutput(JSON.stringify({error:'Unauthorized'})).setMimeType(ContentService.MimeType.JSON);
  var mpn=(e.parameter.mpn||'').trim();
  if (!mpn) return ContentService.createTextOutput(JSON.stringify({error:'No MPN'})).setMimeType(ContentService.MimeType.JSON);
  return ContentService.createTextOutput(JSON.stringify({
    mpn: mpn,
    oem_excess: searchOEMExcess(mpn),
    in_stock: searchInStock(mpn),
    stan_sheet: searchStanSheet(mpn),
    forte_sheet: searchForteSheet(mpn)
  })).setMimeType(ContentService.MimeType.JSON);
}

// ── Fix queue — execute draft fixes queued remotely ───────────
function processFixQueue() {
  try {
    var resp = UrlFetchApp.fetch(HUB_URL + '/api/fix-queue?status=pending', {
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      muteHttpExceptions: true
    });
    var fixes = (JSON.parse(resp.getContentText()).fixes) || [];
    if (!fixes.length) return;

    fixes.forEach(function(fix) {
      try {
        if (fix.type === 'replace_draft') {
          var thread = GmailApp.getThreadById(fix.thread_id);
          if (!thread) throw new Error('Thread not found: ' + fix.thread_id);

          // Delete all existing drafts for this thread
          var allDrafts = GmailApp.getDrafts();
          for (var d = 0; d < allDrafts.length; d++) {
            try {
              if (allDrafts[d].getMessage().getThread().getId() === fix.thread_id) {
                allDrafts[d].deleteDraft();
                Logger.log('Fix queue: deleted draft for ' + fix.thread_id);
              }
            } catch(e2) {}
          }

          // Create the replacement draft
          var firstMsg = thread.getMessages()[0];
          var htmlBody = buildDraftHTML(fix.draft_body, firstMsg);
          var draftId = createThreadedDraft(
            fix.to_email, fix.subject, htmlBody, firstMsg.getId(), fix.thread_id, null
          );
          if (!draftId) throw new Error('createThreadedDraft returned null');

          UrlFetchApp.fetch(HUB_URL + '/api/fix-queue/' + fix.id, {
            method: 'PATCH', contentType: 'application/json',
            headers: { Authorization: 'Bearer ' + HUB_SECRET },
            payload: JSON.stringify({ status: 'done' }),
            muteHttpExceptions: true
          });
          Logger.log('Fix queue done #' + fix.id + ' | thread ' + fix.thread_id);
        }
      } catch(e) {
        Logger.log('Fix queue error #' + fix.id + ': ' + e.toString());
        UrlFetchApp.fetch(HUB_URL + '/api/fix-queue/' + fix.id, {
          method: 'PATCH', contentType: 'application/json',
          headers: { Authorization: 'Bearer ' + HUB_SECRET },
          payload: JSON.stringify({ status: 'failed', error: e.toString() }),
          muteHttpExceptions: true
        });
      }
    });
  } catch(e) {
    Logger.log('processFixQueue error: ' + e.toString());
  }
}

// ── Command queue — inventory actions queued remotely ─────────
function processCommandQueue() {
  try {
    var resp = UrlFetchApp.fetch(HUB_URL + '/api/command-queue?status=pending', {
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      muteHttpExceptions: true
    });
    var commands = (JSON.parse(resp.getContentText()).commands) || [];
    if (!commands.length) return;

    commands.forEach(function(cmd) {
      try {
        var data = {};
        try { data = JSON.parse(cmd.data || '{}'); } catch(e) {}

        if (cmd.type === 'remove_instock_mpn') {
          var mpn = (data.mpn || '').trim();
          if (!mpn) throw new Error('No MPN provided');
          var sheet = SpreadsheetApp.openById(IN_STOCK_ID).getSheets()[0];
          var sheetData = sheet.getDataRange().getValues();
          var searchNorm = normalize(mpn);
          var rowsToDelete = [];
          for (var i = 1; i < sheetData.length; i++) {
            if (normalize(String(sheetData[i][0])) === searchNorm) rowsToDelete.push(i + 1);
          }
          if (!rowsToDelete.length) throw new Error('MPN not found in InStock: ' + mpn);
          rowsToDelete.sort(function(a, b) { return b - a; });
          rowsToDelete.forEach(function(row) { sheet.deleteRow(row); });
          hubLog('inventory', 'Removed ' + rowsToDelete.length + ' row(s) for MPN ' + mpn + ' from InStock', { mpn: mpn, rows_deleted: rowsToDelete.length });

        } else if (cmd.type === 'remove_oem_mpn') {
          var mpn = (data.mpn || '').trim();
          if (!mpn) throw new Error('No MPN provided');
          var result = deletePart(mpn, 'Hub command: remove_oem_mpn');
          if (result === 'NOT_FOUND') throw new Error('MPN not found in OEM EXCESS: ' + mpn);
          if (result === 'MULTIPLE') throw new Error('Multiple exact matches for ' + mpn + ' — review email sent to John');
          if (result === 'FUZZY_REVIEW') throw new Error('Ambiguous match for ' + mpn + ' — review email sent to John');
          updateForteSheet(mpn);
          hubLog('inventory', 'Hub command: removed ' + mpn + ' from OEM EXCESS (result: ' + result + ')', { mpn: mpn, result: result });

        } else if (cmd.type === 'send_datamaster_email') {
          var token = ScriptApp.getOAuthToken();
          var fetchOpts = { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true };
          var DATAMASTER_BCC = [
            '5BDFA5@stkdst.com',
            'datamaster@netcomponents.com',
            'post@icsource.com',
            'bill@intransittech.com',
            'david@fortetechno.com',
            'Stan@amorelectronics.com'
          ].join(',');

          // Build a clean filtered OEM EXCESS XLSX via a temp sheet:
          // - Keeps header row (ICS needs it for column mapping)
          // - Skips rows with blank MPN or blank/non-numeric QTY
          // - Converts comma-text quantities ("1,254") to plain numbers
          // - Removes trailing whitespace from MPN
          // This prevents ICS rejections for null qty and phantom blank rows.
          var oemBlob;
          var oemSS = SpreadsheetApp.openById(SPREADSHEET_ID);
          var srcData = oemSS.getSheets()[0].getDataRange().getValues();
          var tempSheet = oemSS.insertSheet('_ICS_UPLOAD_TEMP');
          try {
            tempSheet.appendRow(srcData[0]); // header row
            var skipped = 0;
            for (var di = 1; di < srcData.length; di++) {
              var mpn = String(srcData[di][0]).trim();
              if (!mpn) { skipped++; continue; }
              var qtyRaw = srcData[di][3];
              var qtyNum = (typeof qtyRaw === 'number') ? qtyRaw
                         : parseFloat(String(qtyRaw).replace(/,/g, ''));
              if (isNaN(qtyNum) || qtyNum <= 0) { skipped++; continue; }
              var cleanRow = srcData[di].slice();
              cleanRow[0] = mpn;
              cleanRow[3] = qtyNum;
              tempSheet.appendRow(cleanRow);
            }
            SpreadsheetApp.flush();
            var gid = tempSheet.getSheetId();
            oemBlob = UrlFetchApp.fetch(
              'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/export?format=xlsx&gid=' + gid,
              fetchOpts
            ).getBlob().setName('OEM_EXCESS.xlsx');
            hubLog('inventory', 'OEM EXCESS clean export: ' + (srcData.length - 1 - skipped) + ' rows sent, ' + skipped + ' blank/null-qty rows skipped', {});
          } finally {
            try { oemSS.deleteSheet(tempSheet); } catch(e) {}
          }

          var inBlob = UrlFetchApp.fetch(
            'https://docs.google.com/spreadsheets/d/' + IN_STOCK_ID + '/export?format=xlsx',
            fetchOpts
          ).getBlob().setName('IN STOCK.xlsx');

          GmailApp.sendEmail(NOTIFY_EMAIL, 'Please post', '', {
            attachments: [oemBlob, inBlob],
            bcc: DATAMASTER_BCC,
            name: 'John Fluman'
          });
          hubLog('inventory', 'Sent NetCOMPONENTS report (OEM_EXCESS + IN STOCK) to ' + DATAMASTER_BCC, {});
        }

        UrlFetchApp.fetch(HUB_URL + '/api/command-queue/' + cmd.id, {
          method: 'PATCH', contentType: 'application/json',
          headers: { Authorization: 'Bearer ' + HUB_SECRET },
          payload: JSON.stringify({ status: 'done' }),
          muteHttpExceptions: true
        });
        Logger.log('Command queue done #' + cmd.id + ' | type: ' + cmd.type);

      } catch(e) {
        Logger.log('Command queue error #' + cmd.id + ': ' + e.toString());
        UrlFetchApp.fetch(HUB_URL + '/api/command-queue/' + cmd.id, {
          method: 'PATCH', contentType: 'application/json',
          headers: { Authorization: 'Bearer ' + HUB_SECRET },
          payload: JSON.stringify({ status: 'failed', error: e.toString() }),
          muteHttpExceptions: true
        });
      }
    });
  } catch(e) {
    Logger.log('processCommandQueue error: ' + e.toString());
  }
}

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger('runEmailScan').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('checkBillNetcompRemovals').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('checkInboxForPaymentAdvice').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('processFixQueue').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('processCommandQueue').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('sendDailyCostReport').timeBased().atHour(8).everyDays(1).create();
  Logger.log('6 triggers installed.');
}

// ── Gmail Add-on — Draft Review Sidebar ──────────────────────
// ── Draft HTML helpers ────────────────────────────────────────

function extractDraftHtmlBody(payload) {
  if (!payload) return null;
  if (payload.mimeType === 'text/html' && payload.body && payload.body.data) {
    try { return Utilities.newBlob(Utilities.base64Decode(payload.body.data.replace(/-/g,'+').replace(/_/g,'/'))).getDataAsString(); } catch(e) { return null; }
  }
  var parts = payload.parts || [];
  for (var i = 0; i < parts.length; i++) { var r = extractDraftHtmlBody(parts[i]); if (r) return r; }
  return null;
}

function extractAdviceText(htmlBody) {
  if (!htmlBody) return null;
  var m = htmlBody.match(/Note for John \(remove before sending\):<\/b><br>([\s\S]*?)<\/div>/);
  return m ? m[1].replace(/<[^>]+>/g, '').trim() : null;
}

function stripAdviceFromHtml(htmlBody) {
  if (!htmlBody) return htmlBody;
  return htmlBody.replace(/<div style="background:#fff3cd[\s\S]*?<\/div>/, '').trim();
}

function rebuildRawMessage(draft, newHtmlBody) {
  var headers = (draft.message && draft.message.payload && draft.message.payload.headers) || [];
  var toH = '', subjectH = '', inReplyTo = '', references = '', ccH = '';
  headers.forEach(function(h) {
    if (h.name === 'To')          toH        = h.value;
    if (h.name === 'Subject')     subjectH   = h.value;
    if (h.name === 'In-Reply-To') inReplyTo  = h.value;
    if (h.name === 'References')  references = h.value;
    if (h.name === 'Cc')          ccH        = h.value;
  });
  var lines = ['From: John Fluman <' + JOHN_EMAIL + '>', 'To: ' + toH];
  if (ccH) lines.push('Cc: ' + ccH);
  lines.push('Subject: ' + subjectH);
  if (inReplyTo) lines.push('In-Reply-To: ' + inReplyTo);
  if (references) lines.push('References: ' + references);
  lines.push('MIME-Version: 1.0', 'Content-Type: text/html; charset=UTF-8', '');
  lines.push(newHtmlBody);
  return { raw: Utilities.base64EncodeWebSafe(lines.join('\r\n')), to: toH, subject: subjectH };
}

// Homepage card — lists all current drafts with Send/Wrong/Fix buttons
function buildHomepageCard() {
  try {
    var token = ScriptApp.getOAuthToken();
    var listResp = UrlFetchApp.fetch(
      'https://gmail.googleapis.com/gmail/v1/users/me/drafts?maxResults=15',
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    var drafts = JSON.parse(listResp.getContentText()).drafts || [];

    var builder = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Intransit Assistant')
        .setSubtitle(drafts.length === 0 ? 'No drafts' : drafts.length + ' draft(s) ready to review'));

    if (drafts.length === 0) {
      builder.addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText('No drafts in your mailbox. Close and reopen this panel to refresh.')));
      return builder.build();
    }

    // Fetch D1 pending drafts to get advice (stored separately from email body)
    var d1Map = {};
    try {
      var d1Resp = UrlFetchApp.fetch(HUB_URL + '/api/drafts?status=pending&limit=100', {
        headers: { Authorization: 'Bearer ' + HUB_SECRET }, muteHttpExceptions: true
      });
      var d1Rows = JSON.parse(d1Resp.getContentText()).rows || [];
      d1Rows.forEach(function(row) {
        if (!row.thread_id) return;
        var content = row.draft_content || '';
        var advIdx = content.indexOf('[ADVICE_STORED]:');
        if (advIdx >= 0) {
          var afterAdv = content.substring(advIdx + '[ADVICE_STORED]:'.length);
          var gmailIdx = afterAdv.indexOf('\n\n[GMAIL_DRAFT:');
          var advice = gmailIdx >= 0 ? afterAdv.substring(0, gmailIdx).trim() : afterAdv.trim();
          if (advice && !d1Map[row.thread_id]) d1Map[row.thread_id] = advice;
        }
      });
    } catch(e2) { Logger.log('buildHomepageCard D1 fetch error: ' + e2); }

    // Fetch all Gmail draft details in parallel
    var batch = drafts.slice(0, 12);
    var requests = batch.map(function(stub) {
      return {
        url: 'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + stub.id + '?format=full',
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      };
    });
    var responses = UrlFetchApp.fetchAll(requests);

    responses.forEach(function(resp, i) {
      if (resp.getResponseCode() !== 200) return;
      var draft = JSON.parse(resp.getContentText());
      var stub = batch[i];
      var headers = (draft.message && draft.message.payload && draft.message.payload.headers) || [];
      var threadId = (draft.message && draft.message.threadId) || '';
      var toH = '', subjectH = '';
      headers.forEach(function(h) {
        if (h.name === 'To') toH = h.value;
        if (h.name === 'Subject') subjectH = h.value;
      });

      // Get advice from D1 (never from draft body — body is always clean now)
      var adviceText = d1Map[threadId] || null;
      var label = (subjectH || 'Draft').replace(/^Re:\s*/i, '');
      if (label.length > 55) label = label.substring(0, 52) + '...';

      var section = CardService.newCardSection().setHeader('📧 ' + label);
      section.addWidget(CardService.newTextParagraph().setText('To: ' + (toH || 'unknown')));

      if (adviceText) {
        var short = adviceText.length > 220 ? adviceText.substring(0, 217) + '...' : adviceText;
        section.addWidget(CardService.newTextParagraph().setText('💡 ' + short));
      } else {
        section.addWidget(CardService.newTextParagraph().setText('Ready to send.'));
      }

      section.addWidget(CardService.newTextButton()
        .setText('✅ Send')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#1a7340')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('addonSendDraft')
          .setParameters({ draftId: stub.id, threadId: threadId, hasAdvice: '0' })));

      var fbField = 'fb_' + stub.id.replace(/[^a-zA-Z0-9]/g, '_');
      section.addWidget(CardService.newTextInput()
        .setFieldName(fbField)
        .setTitle('What was wrong / how to fix?')
        .setHint('e.g. "Should ask for TP not check" or "Wrong MPN"')
        .setMultiline(false));

      section.addWidget(CardService.newTextButton()
        .setText('🔧 Fix this draft')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#1565c0')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('addonFixDraft')
          .setParameters({ draftId: stub.id, threadId: threadId, subject: subjectH, toEmail: toH })));

      section.addWidget(CardService.newTextButton()
        .setText('❌ Wrong — delete & retrain')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#c0392b')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('addonSubmitFeedback')
          .setParameters({ draftId: stub.id, threadId: threadId, fbField: fbField })));

      builder.addSection(section);
    });

    // ── Inventory section ────────────────────────────────────────────────────
    var invSectionHome = CardService.newCardSection().setHeader('📦 Inventory');
    invSectionHome.addWidget(CardService.newTextInput()
      .setFieldName('invMpn')
      .setTitle('Part Number (MPN)')
      .setHint('Used by all buttons below')
      .setMultiline(false));
    invSectionHome.addWidget(CardService.newTextButton()
      .setText('🗑 Remove from InStock')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#7b1fa2')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonRemoveStock')
        .setParameters({ threadId: '' })));
    invSectionHome.addWidget(CardService.newTextButton()
      .setText('🗑 Remove from OEM EXCESS')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#6a1b9a')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonRemoveOEM')
        .setParameters({})));
    invSectionHome.addWidget(CardService.newTextButton()
      .setText('✓ Verify OEM Removed')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#1b5e20')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonVerifyOEMRemoved')
        .setParameters({})));
    invSectionHome.addWidget(CardService.newTextButton()
      .setText('✓ Verify InStock Removed')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#1b5e20')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonVerifyInStockRemoved')
        .setParameters({})));
    invSectionHome.addWidget(CardService.newTextButton()
      .setText('📋 Quote History')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#37474f')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonQuoteHistory')
        .setParameters({})));
    invSectionHome.addWidget(CardService.newTextButton()
      .setText('📤 Send to NetCOMPONENTS')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#1565c0')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonSendNetCom')
        .setParameters({})));
    builder.addSection(invSectionHome);

    var blockSectionHome = CardService.newCardSection().setHeader('🚫 Block Domain');
    blockSectionHome.addWidget(CardService.newTextInput()
      .setFieldName('blockDomain')
      .setTitle('Domain to block')
      .setHint('e.g. spamco.com'));
    blockSectionHome.addWidget(CardService.newTextButton()
      .setText('🚫 Block This Domain')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#b71c1c')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonBlockDomain')
        .setParameters({})));
    builder.addSection(blockSectionHome);

    return builder.build();
  } catch(err) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Intransit Assistant').setSubtitle('Error'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText(err.toString())))
      .build();
  }
}

// Contextual card — fires when any email is opened; shows thread-relevant info.
function buildContextualCard(e) {
  try {
    var gmailThreadId = e.gmail && e.gmail.threadId;
    if (!gmailThreadId) return buildHomepageCard();

    var token = ScriptApp.getOAuthToken();

    // Get thread subject + sender
    var threadResp = UrlFetchApp.fetch(
      'https://gmail.googleapis.com/gmail/v1/users/me/threads/' + gmailThreadId +
      '?format=metadata&metadataHeaders=Subject&metadataHeaders=From',
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    var threadData = JSON.parse(threadResp.getContentText());
    var msgs = threadData.messages || [];
    var subject = '', fromH = '';
    if (msgs.length > 0) {
      (msgs[0].payload && msgs[0].payload.headers || []).forEach(function(h) {
        if (h.name === 'Subject') subject = h.value;
        if (h.name === 'From') fromH = h.value;
      });
    }

    // Look up D1 advice for this thread
    var d1Advice = null;
    try {
      var d1Resp = UrlFetchApp.fetch(HUB_URL + '/api/drafts?status=pending&limit=100', {
        headers: { Authorization: 'Bearer ' + HUB_SECRET }, muteHttpExceptions: true
      });
      var d1Rows = JSON.parse(d1Resp.getContentText()).rows || [];
      var d1Match = d1Rows.filter(function(r) { return r.thread_id === gmailThreadId; })[0];
      if (d1Match) {
        var content = d1Match.draft_content || '';
        var advIdx = content.indexOf('[ADVICE_STORED]:');
        if (advIdx >= 0) {
          var after = content.substring(advIdx + '[ADVICE_STORED]:'.length);
          var gIdx = after.indexOf('\n\n[GMAIL_DRAFT:');
          d1Advice = (gIdx >= 0 ? after.substring(0, gIdx) : after).trim();
        }
      }
    } catch(e2) { Logger.log('buildContextualCard D1 error: ' + e2); }

    // Find a Gmail draft for this thread
    var matchDraftId = null, matchToH = '';
    try {
      var listResp = UrlFetchApp.fetch(
        'https://gmail.googleapis.com/gmail/v1/users/me/drafts?maxResults=50',
        { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
      );
      var stubs = JSON.parse(listResp.getContentText()).drafts || [];
      if (stubs.length > 0) {
        var reqs = stubs.map(function(s) {
          return {
            url: 'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + s.id + '?format=metadata&metadataHeaders=To',
            headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
          };
        });
        UrlFetchApp.fetchAll(reqs).forEach(function(r, i) {
          if (r.getResponseCode() !== 200 || matchDraftId) return;
          var d = JSON.parse(r.getContentText());
          if (d.message && d.message.threadId === gmailThreadId) {
            matchDraftId = stubs[i].id;
            (d.message.payload && d.message.payload.headers || []).forEach(function(h) {
              if (h.name === 'To') matchToH = h.value;
            });
          }
        });
      }
    } catch(e3) { Logger.log('buildContextualCard draft search error: ' + e3); }

    var label = (subject || 'Email').replace(/^Re:\s*/i, '').substring(0, 55);
    var builder = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Intransit Assistant')
        .setSubtitle(label));

    if (matchDraftId) {
      var fbField = 'fb_' + matchDraftId.replace(/[^a-zA-Z0-9]/g, '_');

      var infoSection = CardService.newCardSection().setHeader('📝 Draft ready');
      infoSection.addWidget(CardService.newTextParagraph().setText('To: ' + (matchToH || 'unknown')));
      if (d1Advice) {
        var short = d1Advice.length > 250 ? d1Advice.substring(0, 247) + '...' : d1Advice;
        infoSection.addWidget(CardService.newTextParagraph().setText('💡 ' + short));
      }
      infoSection.addWidget(CardService.newTextButton()
        .setText('✅ Send')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#1a7340')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('addonSendDraft')
          .setParameters({ draftId: matchDraftId, threadId: gmailThreadId, hasAdvice: '0' })));

      var fixSection = CardService.newCardSection().setHeader('Something wrong?');
      fixSection.addWidget(CardService.newTextButton()
        .setText('🤔 Figure Out What\'s Wrong')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#7b1fa2')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('addonDiagnoseDraft')
          .setParameters({ draftId: matchDraftId, threadId: gmailThreadId, subject: subject, fromH: fromH, toEmail: matchToH })));
      fixSection.addWidget(CardService.newTextInput()
        .setFieldName(fbField)
        .setTitle('Or type what was wrong:')
        .setHint('e.g. "Should ask for TP" or "Wrong MPN extracted"')
        .setMultiline(false));
      fixSection.addWidget(CardService.newTextButton()
        .setText('🔧 Fix this draft')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#1565c0')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('addonFixDraft')
          .setParameters({ draftId: matchDraftId, threadId: gmailThreadId, subject: subject, toEmail: matchToH })));
      fixSection.addWidget(CardService.newTextButton()
        .setText('❌ Wrong — delete & retrain')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#c0392b')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('addonSubmitFeedback')
          .setParameters({ draftId: matchDraftId, threadId: gmailThreadId, fbField: fbField })));

      builder.addSection(infoSection).addSection(fixSection);

    } else {
      var noSection = CardService.newCardSection().setHeader('📥 No draft — create one now');
      noSection.addWidget(CardService.newTextParagraph().setText('From: ' + (fromH || 'unknown')));
      if (d1Advice) {
        noSection.addWidget(CardService.newTextParagraph()
          .setText('⚠️ Previously processed — draft may have been sent or deleted.'));
      }
      noSection.addWidget(CardService.newTextInput()
        .setFieldName('draftInstruction')
        .setTitle('Tell me what to draft')
        .setHint('e.g. "MSG_CHECKING", "ask for TP $500 min", "David no stock removal for [MPN]"')
        .setMultiline(false));
      noSection.addWidget(CardService.newTextButton()
        .setText('📤 Create draft now')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#6a1b9a')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('addonCreateDraft')
          .setParameters({ threadId: gmailThreadId, subject: subject, fromEmail: fromH })));
      builder.addSection(noSection);
    }

    // ── Chat section — always shown ──────────────────────────────────────────
    var chatSection = CardService.newCardSection().setHeader('💬 Chat with assistant');
    chatSection.addWidget(CardService.newTextInput()
      .setFieldName('chatMessage')
      .setTitle('Message')
      .setHint('Ask anything or describe what you need — align first, then act')
      .setMultiline(false));
    chatSection.addWidget(CardService.newTextButton()
      .setText('Send')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#37474f')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonChat')
        .setParameters({
          threadId: gmailThreadId,
          subject: subject,
          fromH: fromH,
          draftId: matchDraftId || '',
          draftBody: ''
        })));
    builder.addSection(chatSection);

    // ── Inventory section — always shown ──────────────────────────────────────
    var invMpnHint = extractMPNFromSubject(subject) || '';
    var invSection = CardService.newCardSection().setHeader('📦 Inventory');
    invSection.addWidget(CardService.newTextInput()
      .setFieldName('invMpn')
      .setTitle('Part Number (MPN)')
      .setHint('Used by all buttons below')
      .setValue(invMpnHint)
      .setMultiline(false));
    invSection.addWidget(CardService.newTextButton()
      .setText('🗑 Remove from InStock')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#7b1fa2')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonRemoveStock')
        .setParameters({ threadId: gmailThreadId })));
    invSection.addWidget(CardService.newTextButton()
      .setText('🗑 Remove from OEM EXCESS')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#6a1b9a')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonRemoveOEM')
        .setParameters({})));
    invSection.addWidget(CardService.newTextButton()
      .setText('✓ Verify OEM Removed')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#1b5e20')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonVerifyOEMRemoved')
        .setParameters({})));
    invSection.addWidget(CardService.newTextButton()
      .setText('✓ Verify InStock Removed')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#1b5e20')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonVerifyInStockRemoved')
        .setParameters({})));
    invSection.addWidget(CardService.newTextButton()
      .setText('📋 Quote History')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#37474f')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonQuoteHistory')
        .setParameters({})));
    invSection.addWidget(CardService.newTextButton()
      .setText('📤 Send to NetCOMPONENTS')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#1565c0')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonSendNetCom')
        .setParameters({})));
    builder.addSection(invSection);

    // ── Jiggle My Mind — diagnose why this email was missed ────────────────
    var jiggleSection = CardService.newCardSection().setHeader('🧠 Jiggle My Mind');
    jiggleSection.addWidget(CardService.newTextParagraph()
      .setText('Think this email should have been drafted? Let me diagnose what went wrong.'));
    jiggleSection.addWidget(CardService.newTextButton()
      .setText('🧠 Diagnose This Email')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#263238')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonDiagnoseEmail')
        .setParameters({ threadId: gmailThreadId, subject: subject, fromH: fromH })));
    builder.addSection(jiggleSection);

    // ── Smart Reply ─────────────────────────────────────────────────────────
    var smartSection = CardService.newCardSection().setHeader('🤖 Smart Reply');
    smartSection.addWidget(CardService.newTextParagraph()
      .setText('AI reads the full thread + your inventory and drafts the best reply. Review it, then copy or create as draft.'));
    smartSection.addWidget(CardService.newTextButton()
      .setText('🤖 Generate Smart Reply')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#0d47a1')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonSmartReply')
        .setParameters({ threadId: gmailThreadId, subject: subject, fromH: fromH })));
    builder.addSection(smartSection);

    // ── Block Domain ────────────────────────────────────────────────────────
    var blockSection = CardService.newCardSection().setHeader('🚫 Block Domain');
    // Auto-fill sender domain — but never for passthrough relays like netcomponents.com / icsource.com
    var senderDomain = (fromH.match(/@([\w.-]+)/) || ['',''])[1].toLowerCase();
    var passthroughDomains = ['netcomponents.com','icsource.com','messagesend.com','autosend.com'];
    var prefill = passthroughDomains.indexOf(senderDomain) >= 0 ? '' : senderDomain;
    blockSection.addWidget(CardService.newTextInput()
      .setFieldName('blockDomain')
      .setTitle('Domain to block')
      .setValue(prefill)
      .setHint('e.g. spamco.com'));
    blockSection.addWidget(CardService.newTextButton()
      .setText('🚫 Block This Domain')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#b71c1c')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonBlockDomain')
        .setParameters({})));
    builder.addSection(blockSection);

    return [builder.build()];
  } catch(err) {
    return [buildAddonError(err.toString())];
  }
}

// Main compose trigger card — fires when user clicks add-on icon in compose toolbar
function buildComposeCard(e) {
  try {
    var draftId = e && e.gmail && e.gmail.draftId;

    if (!draftId) {
      return [CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle('Intransit Assistant'))
        .addSection(CardService.newCardSection()
          .addWidget(CardService.newTextParagraph()
            .setText('Open an existing draft (from your Drafts folder) to see advice and send options.')))
        .build()];
    }

    // Fetch full draft via REST API
    var token = ScriptApp.getOAuthToken();
    var resp = UrlFetchApp.fetch(
      'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + draftId + '?format=full',
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200) {
      return [buildAddonError('Could not load draft (HTTP ' + resp.getResponseCode() + ').')];
    }

    var draft = JSON.parse(resp.getContentText());
    var htmlBody  = extractDraftHtmlBody(draft.message && draft.message.payload);
    var threadId  = (draft.message && draft.message.threadId) || '';
    var headers   = (draft.message && draft.message.payload && draft.message.payload.headers) || [];
    var toH = '', subjectH = '';
    headers.forEach(function(h) {
      if (h.name === 'To')      toH      = h.value;
      if (h.name === 'Subject') subjectH = h.value;
    });

    var adviceText = extractAdviceText(htmlBody);
    var hasAdvice  = !!adviceText;

    // ── Advice display section ──
    var adviceSection = CardService.newCardSection()
      .setHeader(subjectH ? ('📧 ' + subjectH.replace(/^Re:\s*/i,'')) : 'Draft');

    if (hasAdvice) {
      adviceSection.addWidget(CardService.newTextParagraph().setText('💡 ' + adviceText));
    } else {
      adviceSection.addWidget(CardService.newTextParagraph()
        .setText('No advice block found — draft looks clean.'));
    }

    // ── Send button ──
    var sendSection = CardService.newCardSection();
    sendSection.addWidget(
      CardService.newTextButton()
        .setText(hasAdvice ? '✅  Send  (advice stripped automatically)' : '✅  Send as-is')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#1a7340')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('addonSendDraft')
          .setParameters({ draftId: draftId, threadId: threadId, hasAdvice: hasAdvice ? '1' : '0' }))
    );

    // ── Feedback / retrain section (collapsed) ──
    var feedbackSection = CardService.newCardSection()
      .setHeader('❌  Wrong draft — retrain')
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(0);
    feedbackSection.addWidget(
      CardService.newTextInput()
        .setFieldName('feedbackText')
        .setTitle('What was wrong with this draft?')
        .setHint('e.g. "Should be need TP not checking" or "Wrong MPN extracted"')
        .setMultiline(true)
    );
    feedbackSection.addWidget(
      CardService.newTextButton()
        .setText('Submit feedback & delete draft')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#c0392b')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('addonSubmitFeedback')
          .setParameters({ draftId: draftId, threadId: threadId }))
    );

    return [CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Intransit Assistant')
        .setSubtitle('To: ' + (toH || 'unknown')))
      .addSection(adviceSection)
      .addSection(sendSection)
      .addSection(feedbackSection)
      .build()];

  } catch(err) {
    return [buildAddonError(err.toString())];
  }
}

function buildAddonError(msg) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Intransit Assistant').setSubtitle('Error'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText(msg)))
    .build();
}

// ── Add-on action handlers ────────────────────────────────────

function addonSendDraft(e) {
  try {
    var params   = e.commonEventObject.parameters;
    var draftId  = params.draftId;
    var threadId = params.threadId;
    var hasAdvice = params.hasAdvice === '1';
    var token    = ScriptApp.getOAuthToken();

    if (hasAdvice) {
      // Fetch full draft, strip advice, update draft, then send
      var fetchResp = UrlFetchApp.fetch(
        'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + draftId + '?format=full',
        { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
      );
      if (fetchResp.getResponseCode() !== 200) {
        return notify('Error fetching draft: HTTP ' + fetchResp.getResponseCode());
      }
      var draft     = JSON.parse(fetchResp.getContentText());
      var htmlBody  = extractDraftHtmlBody(draft.message && draft.message.payload);
      var cleanHtml = stripAdviceFromHtml(htmlBody);
      var rebuilt   = rebuildRawMessage(draft, cleanHtml);

      // Update draft with clean HTML
      var putResp = UrlFetchApp.fetch(
        'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + draftId,
        {
          method: 'PUT',
          headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
          payload: JSON.stringify({ message: { raw: rebuilt.raw, threadId: threadId || undefined } }),
          muteHttpExceptions: true,
        }
      );
      if (JSON.parse(putResp.getContentText()).error) {
        return notify('Error updating draft before send.');
      }
      hubLog('run', 'Add-on: stripped advice from draft — ' + rebuilt.subject + ' → ' + rebuilt.to);
    }

    // Send the draft
    var sendResp = UrlFetchApp.fetch(
      'https://gmail.googleapis.com/gmail/v1/users/me/drafts/send',
      {
        method: 'POST',
        headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify({ id: draftId }),
        muteHttpExceptions: true,
      }
    );
    var sendData = JSON.parse(sendResp.getContentText());
    if (sendData.error) {
      return notify('Send failed: ' + (sendData.error.message || JSON.stringify(sendData.error)));
    }

    hubLog('run', 'Add-on: sent draft ' + draftId);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('✅ Email sent (advice stripped)'))
      .build();

  } catch(err) {
    return notify('Error: ' + err.toString());
  }
}

function addonFixDraft(e) {
  try {
    var params     = e.commonEventObject.parameters;
    var draftId    = params.draftId;
    var threadId   = params.threadId;
    var subject    = params.subject || '';
    var toEmail    = params.toEmail || '';
    var formInputs = e.commonEventObject.formInputs || {};
    var fbField    = 'fb_' + draftId.replace(/[^a-zA-Z0-9]/g, '_');
    var feedback   = '';
    if (formInputs[fbField] && formInputs[fbField].stringInputs) {
      feedback = (formInputs[fbField].stringInputs.value || [])[0] || '';
    }
    if (!feedback.trim()) return notify('Please type what was wrong before clicking Fix.');

    var token = ScriptApp.getOAuthToken();

    // Fetch current draft body (clean — no advice in body now)
    var fetchResp = UrlFetchApp.fetch(
      'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + draftId + '?format=full',
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (fetchResp.getResponseCode() !== 200) return notify('Could not load draft.');
    var draft = JSON.parse(fetchResp.getContentText());
    var htmlBody = extractDraftHtmlBody(draft.message && draft.message.payload);
    var currentBody = htmlBody ? htmlBody.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim() : '';

    // Call hub fix-draft endpoint
    var fixResp = UrlFetchApp.fetch(HUB_URL + '/api/fix-draft', {
      method: 'POST',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({
        draft_body: currentBody,
        feedback: feedback,
        subject: subject,
        to_email: toEmail,
        thread_id: threadId
      }),
      muteHttpExceptions: true
    });
    if (fixResp.getResponseCode() !== 200) {
      return notify('Fix failed: ' + fixResp.getContentText().substring(0, 100));
    }
    var fixResult = JSON.parse(fixResp.getContentText());
    var correctedBody = fixResult.corrected_body || '';
    var newAdvice = fixResult.advice || ('Fixed per feedback: ' + feedback);

    // Build clean HTML (no advice in body)
    var newHtml = buildSimpleHTML(correctedBody.replace(/\n/g, '<br>'));
    var rebuilt = rebuildRawMessage(draft, newHtml);

    // Update the draft in Gmail
    var putResp = UrlFetchApp.fetch(
      'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + draftId,
      {
        method: 'PUT',
        headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify({ message: { raw: rebuilt.raw, threadId: threadId || undefined } }),
        muteHttpExceptions: true,
      }
    );
    if (JSON.parse(putResp.getContentText()).error) return notify('Could not update draft.');

    // Store new advice in D1
    hubPostDraft(threadId, null, toEmail, subject, correctedBody, draftId, newAdvice);
    hubLearn(feedback, currentBody, correctedBody, threadId, subject, toEmail, null, null);
    hubLog('run', 'addonFixDraft: draft fixed + lesson saved — ' + subject, { feedback: feedback });

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('✅ Draft fixed! 🧠 Lesson saved — agent will remember this correction.'))
      .build();

  } catch(err) {
    return notify('Error: ' + err.toString());
  }
}

function addonCreateDraft(e) {
  try {
    var params     = e.commonEventObject.parameters;
    var threadId   = params.threadId;
    var subject    = params.subject || '';
    var fromEmail  = params.fromEmail || '';
    var formInputs = e.commonEventObject.formInputs || {};
    var instruction = '';
    if (formInputs.draftInstruction && formInputs.draftInstruction.stringInputs) {
      instruction = (formInputs.draftInstruction.stringInputs.value || [])[0] || '';
    }
    if (!instruction.trim()) return notify('Please type what to draft before clicking Create.');

    // Get thread to find the last message and reply target
    var thread = GmailApp.getThreadById(threadId);
    if (!thread) return notify('Thread not found.');
    var messages = thread.getMessages();
    var lastMsg = messages[messages.length - 1];
    var threadSnippet = lastMsg.getPlainBody().substring(0, 600);
    var replyTo = fromEmail || lastMsg.getReplyTo() || lastMsg.getFrom();

    // Call fix-draft endpoint (also handles create-from-scratch when draft_body is empty)
    var fixResp = UrlFetchApp.fetch(HUB_URL + '/api/fix-draft', {
      method: 'POST',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({
        draft_body: '',
        feedback: instruction,
        subject: subject,
        to_email: replyTo,
        thread_id: threadId,
        thread_context: threadSnippet
      }),
      muteHttpExceptions: true
    });

    if (fixResp.getResponseCode() !== 200) {
      return notify('Failed: ' + fixResp.getContentText().substring(0, 120));
    }
    var result = JSON.parse(fixResp.getContentText());
    var bodyText = result.corrected_body || '';
    var advice = result.advice || ('Created per: ' + instruction);
    if (!bodyText.trim()) return notify('Received empty draft body from AI. Try rephrasing the instruction.');

    var htmlBody = buildSimpleHTML(bodyText.replace(/\n/g, '<br>'));
    var draft = lastMsg.createDraftReply('', { htmlBody: htmlBody });
    if (!draft) return notify('Failed to save draft to Gmail.');

    var draftId = draft.getId();
    hubPostDraft(threadId, null, replyTo, subject, bodyText, draftId, advice);
    hubLog('run', 'addonCreateDraft: sidebar-created draft — ' + subject, { instruction: instruction });

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('✅ Draft created! Close and reopen the sidebar to review it.'))
      .build();

  } catch(err) {
    return notify('Error: ' + err.toString());
  }
}

function addonSubmitFeedback(e) {
  try {
    var params    = e.commonEventObject.parameters;
    var draftId   = params.draftId;
    var threadId  = params.threadId;
    var fbField   = params.fbField || params.feedbackField || 'feedbackText';
    var formInputs = e.commonEventObject.formInputs || {};
    var feedback  = '';
    if (formInputs[fbField] && formInputs[fbField].stringInputs) {
      feedback = (formInputs[fbField].stringInputs.value || [])[0] || '';
    }

    var token = ScriptApp.getOAuthToken();

    // Fetch draft body BEFORE deleting — needed for lesson extraction
    var draftBody = '';
    var draftSubject = '', draftSender = '', draftMpn = '', draftAction = '';
    try {
      var fetchResp = UrlFetchApp.fetch(
        'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + draftId + '?format=full',
        { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
      );
      if (fetchResp.getResponseCode() === 200) {
        var draftData = JSON.parse(fetchResp.getContentText());
        var rawHtml = extractDraftHtmlBody(draftData.message && draftData.message.payload);
        draftBody = rawHtml ? rawHtml.replace(/<[^>]+>/g,' ').replace(/\s+/g,' ').trim().substring(0,400) : '';
        (draftData.message && draftData.message.payload && draftData.message.payload.headers || []).forEach(function(h) {
          if (h.name === 'To') draftSender = h.value;
          if (h.name === 'Subject') draftSubject = h.value;
        });
      }
    } catch(eF) {}

    // Look up D1 for MPN and action context
    try {
      var d1Resp = UrlFetchApp.fetch(HUB_URL + '/api/drafts?status=pending&limit=50', {
        headers: { Authorization: 'Bearer ' + HUB_SECRET }, muteHttpExceptions: true,
      });
      var rows = JSON.parse(d1Resp.getContentText()).rows || [];
      var match = rows.filter(function(r) { return r.thread_id === threadId; })[0];
      if (match) {
        hubPatchEntry(match.id, { action: 'wrong', sent_content: feedback });
        draftMpn = match.mpn || '';
        draftSender = draftSender || match.sender || '';
      }
    } catch(e2) {}

    // Extract and store a lesson (async — don't block the UI)
    if (feedback.trim()) {
      hubLearn(feedback, draftBody, '', threadId, draftSubject, draftSender, draftMpn, draftAction);
    }

    hubLog('feedback', 'Draft marked wrong: ' + (feedback || '(no reason)'), { draft_id: draftId, thread_id: threadId });

    // Delete the draft
    UrlFetchApp.fetch(
      'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + draftId,
      { method: 'DELETE', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(feedback
          ? '🧠 Lesson saved + draft deleted. The agent will remember this.'
          : 'Draft deleted. (Add a reason next time so the agent can learn.)'))
      .build();

  } catch(err) {
    return notify('Error: ' + err.toString());
  }
}

function getRecentSentQuotesFull(mpn, maxThreads) {
  if (!mpn) return 'No MPN provided.';
  try {
    var threads = GmailApp.search('in:sent "' + mpn + '"', 0, maxThreads || 5);
    if (!threads.length) return 'No prior sent emails found for ' + mpn + '.';
    var out = [];
    threads.forEach(function(thread) {
      var msgs = thread.getMessages();
      for (var i = msgs.length - 1; i >= 0; i--) {
        var msg = msgs[i];
        if (msg.getFrom().indexOf(JOHN_EMAIL) >= 0) {
          var body = stripQuotedLines(msg.getPlainBody()).substring(0, 350);
          var dateStr = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), 'MMM d, yyyy');
          out.push('[Sent ' + dateStr + ' to ' + msg.getTo() + ']\n' + body);
          break;
        }
      }
    });
    return out.length ? out.join('\n\n') : 'No sent messages from John found for ' + mpn + '.';
  } catch(e) {
    return 'Email search error: ' + e.toString();
  }
}

function addonChat(e) {
  try {
    var params     = e.commonEventObject.parameters;
    var threadId   = params.threadId  || '';
    var subject    = params.subject   || '';
    var fromH      = params.fromH     || '';
    var draftId    = params.draftId   || '';
    var formInputs = e.commonEventObject.formInputs || {};
    var message    = '';
    if (formInputs.chatMessage && formInputs.chatMessage.stringInputs) {
      message = (formInputs.chatMessage.stringInputs.value || [])[0] || '';
    }
    if (!message.trim()) return notify('Please type a message before clicking Send.');

    // ── Gather full context (Apps Script has Gmail + Sheets access) ──────────
    var mpn = extractMPN(subject);

    // Full thread text (up to 5 messages, 600 chars each)
    var fullThread = '';
    try {
      var thread = GmailApp.getThreadById(threadId);
      if (thread) {
        var threadMsgs = thread.getMessages();
        fullThread = threadMsgs.map(function(m, i) {
          return 'Message ' + (i+1) + ' | From: ' + m.getFrom() + ' | ' + Utilities.formatDate(m.getDate(), Session.getScriptTimeZone(), 'MMM d, h:mm a') + '\n'
            + m.getPlainBody().substring(0, 600);
        }).join('\n\n---\n\n');
      }
    } catch(e2) { fullThread = '(thread fetch error: ' + e2 + ')'; }

    // Prior sent quotes for this MPN
    var priorQuotes = mpn ? getRecentSentQuotesFull(mpn, 5) : 'No MPN extracted from subject.';

    // OEM EXCESS + Forte data via web app
    var oemResults = [], forteResults = [];
    if (mpn) {
      try {
        var webUrl = 'https://script.google.com/macros/s/AKfycbyuuBmiYVW5mKI82D5YQGPh1nNGLJZzlLKoxuOdtmOUwUe75VlhhakqgwKooZu5LHFK/exec'
          + '?key=baSDJ%23444FE%268&mpn=' + encodeURIComponent(mpn);
        var webResp = UrlFetchApp.fetch(webUrl, { followRedirects: true, muteHttpExceptions: true });
        var webData = JSON.parse(webResp.getContentText());
        oemResults   = webData.oem_excess   || [];
        forteResults = webData.forte_sheet  || [];
      } catch(e3) { Logger.log('addonChat web app error: ' + e3); }
    }

    // Inbox summary — other threads needing attention (quick scan, not full content)
    var inboxSummary = '';
    try {
      var inboxThreads = GmailApp.search('in:inbox -from:intransittech.com -label:oem-rfq-incoming-processed newer_than:3d', 0, 8);
      if (inboxThreads.length) {
        inboxSummary = inboxThreads.map(function(t) {
          var m = t.getMessages();
          var last = m[m.length - 1];
          return Utilities.formatDate(last.getDate(), Session.getScriptTimeZone(), 'MMM d h:mm a')
            + ' | From: ' + last.getFrom().replace(/<.*>/, '').trim()
            + ' | ' + t.getFirstMessageSubject().substring(0, 80);
        }).join('\n');
      }
    } catch(e4) { Logger.log('addonChat inbox scan error: ' + e4); }

    // Last agent decision for this thread — tells Claude what draft was created and why
    var agentDraftBody = '', agentReasoning = '', agentAction = '';
    try {
      var decResp = UrlFetchApp.fetch(HUB_URL + '/api/agent-decisions?thread_id=' + encodeURIComponent(threadId), {
        headers: { Authorization: 'Bearer ' + HUB_SECRET }, muteHttpExceptions: true
      });
      var decData = JSON.parse(decResp.getContentText());
      var decisions = decData.decisions || [];
      if (decisions.length) {
        agentDraftBody  = decisions[0].draft_body  || '';
        agentReasoning  = decisions[0].reasoning   || '';
        agentAction     = decisions[0].action      || '';
      }
    } catch(e5) { Logger.log('addonChat decision fetch error: ' + e5); }

    var chatResp = UrlFetchApp.fetch(HUB_URL + '/api/chat', {
      method: 'POST', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({
        thread_id:        threadId,
        message:          message,
        subject:          subject,
        from_email:       fromH,
        mpn:              mpn || '',
        full_thread:      fullThread,
        prior_quotes:     priorQuotes,
        oem_results:      oemResults,
        forte_results:    forteResults,
        inbox_summary:    inboxSummary || '(no other inbox threads)',
        draft_body:       agentDraftBody,
        agent_action:     agentAction,
        agent_reasoning:  agentReasoning
      }),
      muteHttpExceptions: true
    });

    if (chatResp.getResponseCode() !== 200) {
      return notify('Chat error: ' + chatResp.getContentText().substring(0, 120));
    }
    var result  = JSON.parse(chatResp.getContentText());
    var reply   = result.response || '(no response)';
    var action  = result.action   || null;

    // Build response card (pushed on top, user can press back)
    var builder = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Intransit Assistant').setSubtitle('Chat — ' + (subject || '').substring(0, 40)));

    var convSection = CardService.newCardSection().setHeader('💬');
    convSection.addWidget(CardService.newTextParagraph().setText('You: ' + message));
    convSection.addWidget(CardService.newTextParagraph().setText('Claude: ' + reply));
    builder.addSection(convSection);

    // Confirmed action → show action buttons
    if (action && action.type) {
      var actionType = action.type;
      var actSection = CardService.newCardSection().setHeader('✅ Ready to execute');
      // Preview text
      var preview = action.advice || action.body || '';
      if (actionType === 'add_forte') preview = 'Add ' + action.mpn + ' to Forte (QTY: ' + action.qty + ', TP: $' + (action.tp || '?') + ')';
      if (actionType === 'remove_oem_excess') preview = 'Remove ' + action.mpn + ' from OEM EXCESS';
      if (actionType === 'update_rule') preview = (action.delete ? 'Delete' : 'Update') + ' rule: ' + action.rule_type + '/' + action.key;
      if (actionType === 'apply_label') preview = 'Apply label: ' + action.label;
      if (actionType === 'multi') preview = (action.advice || 'Execute ' + (action.actions || []).length + ' actions');
      actSection.addWidget(CardService.newTextParagraph().setText(preview.substring(0, 200)));
      var actionParams = { threadId: threadId, subject: subject, fromH: fromH, actionJson: JSON.stringify(action) };
      if (actionType === 'create_draft') {
        actSection.addWidget(CardService.newTextButton()
          .setText('📝 Create Draft')
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setBackgroundColor('#1a7340')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('addonExecuteAction')
            .setParameters(actionParams)));
        actSection.addWidget(CardService.newTextButton()
          .setText('🚀 Create & Send Now')
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setBackgroundColor('#1565c0')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('addonExecuteAndSend')
            .setParameters(actionParams)));
      } else {
        actSection.addWidget(CardService.newTextButton()
          .setText('✅ Execute')
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setBackgroundColor('#1a7340')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('addonExecuteAction')
            .setParameters(actionParams)));
      }
      builder.addSection(actSection);
    }

    // Continue chat
    var contSection = CardService.newCardSection().setHeader('Continue');
    contSection.addWidget(CardService.newTextInput()
      .setFieldName('chatMessage').setTitle('Reply').setMultiline(false));
    contSection.addWidget(CardService.newTextButton()
      .setText('Send')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#37474f')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonChat')
        .setParameters({ threadId: threadId, subject: subject, fromH: fromH, draftId: draftId })));
    builder.addSection(contSection);

    // Report Issue section — always visible at bottom
    var issueSection = CardService.newCardSection().setHeader('🐛 Something wrong?');
    issueSection.addWidget(CardService.newTextInput()
      .setFieldName('issueDescription')
      .setTitle('Describe the issue')
      .setHint('e.g. "Wrong routing — should have asked for TP not sent to Bill"')
      .setMultiline(true));
    issueSection.addWidget(CardService.newTextButton()
      .setText('Report Issue & Fix')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#b71c1c')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonReportIssue')
        .setParameters({ threadId: threadId, subject: subject, mpn: (e.commonEventObject.parameters.mpn || '') })));
    builder.addSection(issueSection);

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(builder.build()))
      .build();

  } catch(err) {
    return notify('Error in chat: ' + err.toString());
  }
}

function addonRemoveStock(e) {
  try {
    var formInputs = (e.commonEventObject && e.commonEventObject.formInputs) || {};
    var mpnField = formInputs.invMpn;
    var mpn = (mpnField && mpnField.stringInputs && mpnField.stringInputs.value && mpnField.stringInputs.value[0] || '').trim().toUpperCase();
    if (!mpn) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('⚠️ Enter an MPN first.'))
        .build();
    }
    UrlFetchApp.fetch(HUB_URL + '/api/command-queue', {
      method: 'POST', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({ type: 'remove_instock_mpn', data: { mpn: mpn } }),
      muteHttpExceptions: true
    });
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('✅ Queued: removing ' + mpn + ' from InStock (~5 min)'))
      .build();
  } catch(err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('❌ Error: ' + err.toString()))
      .build();
  }
}

function addonBlockDomain(e) {
  try {
    var domain = ((e.formInput && e.formInput.blockDomain) || '').toLowerCase().trim()
                  .replace(/^@/, '').replace(/\/.*$/, ''); // strip leading @ or paths
    if (!domain || domain.indexOf('.') < 0) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('⚠️ Enter a valid domain (e.g. spamco.com).'))
        .build();
    }
    var result = executeUpdateRule({ rule_type: 'blocked_domain', key: domain, value: 'true', notes: 'Blocked from sidebar' });
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(result.ok ? '🚫 Blocked: ' + domain : '❌ ' + result.message))
      .build();
  } catch(err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('❌ ' + err.toString()))
      .build();
  }
}

function addonSendNetCom(e) {
  try {
    UrlFetchApp.fetch(HUB_URL + '/api/command-queue', {
      method: 'POST', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({ type: 'send_datamaster_email' }),
      muteHttpExceptions: true
    });
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('✅ Queued: NetCOMPONENTS email will send in ~5 min'))
      .build();
  } catch(err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('❌ Error: ' + err.toString()))
      .build();
  }
}

function addonRemoveOEM(e) {
  try {
    var formInputs = (e.commonEventObject && e.commonEventObject.formInputs) || {};
    var mpn = (formInputs.invMpn && formInputs.invMpn.stringInputs && formInputs.invMpn.stringInputs.value && formInputs.invMpn.stringInputs.value[0] || '').trim().toUpperCase();
    if (!mpn) return notify('Enter an MPN first.');
    UrlFetchApp.fetch(HUB_URL + '/api/command-queue', {
      method: 'POST', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({ type: 'remove_oem_mpn', data: { mpn: mpn } }),
      muteHttpExceptions: true
    });
    return notify('✅ Queued: removing ' + mpn + ' from OEM EXCESS (~5 min)');
  } catch(err) { return notify('❌ Error: ' + err.toString()); }
}

function addonSheetLookup(mpn) {
  var resp = UrlFetchApp.fetch(HUB_URL + '/api/sheet-lookup?mpn=' + encodeURIComponent(mpn), {
    headers: { Authorization: 'Bearer ' + HUB_SECRET },
    muteHttpExceptions: true, followRedirects: true
  });
  return JSON.parse(resp.getContentText());
}

function addonVerifyOEMRemoved(e) {
  try {
    var formInputs = (e.commonEventObject && e.commonEventObject.formInputs) || {};
    var mpn = (formInputs.invMpn && formInputs.invMpn.stringInputs && formInputs.invMpn.stringInputs.value && formInputs.invMpn.stringInputs.value[0] || '').trim().toUpperCase();
    if (!mpn) return notify('Enter an MPN first.');
    var data = addonSheetLookup(mpn);
    var rows = (data && data.oem_excess) || [];
    var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('OEM EXCESS — ' + mpn).setSubtitle(rows.length === 0 ? '✓ Removed' : '⚠ Still present'));
    var section = CardService.newCardSection();
    if (rows.length === 0) {
      section.addWidget(CardService.newTextParagraph().setText('✅ ' + mpn + ' is NOT in OEM EXCESS. Successfully removed.'));
    } else {
      section.addWidget(CardService.newTextParagraph().setText('⚠️ ' + mpn + ' still found — ' + rows.length + ' row(s):'));
      rows.forEach(function(r) {
        section.addWidget(CardService.newTextParagraph().setText('Row ' + r.row + ': QTY=' + r.qty + ' | Notes=' + (r.notes || '—')));
      });
    }
    card.addSection(section);
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card.build())).build();
  } catch(err) { return notify('❌ Error: ' + err.toString()); }
}

function addonVerifyInStockRemoved(e) {
  try {
    var formInputs = (e.commonEventObject && e.commonEventObject.formInputs) || {};
    var mpn = (formInputs.invMpn && formInputs.invMpn.stringInputs && formInputs.invMpn.stringInputs.value && formInputs.invMpn.stringInputs.value[0] || '').trim().toUpperCase();
    if (!mpn) return notify('Enter an MPN first.');
    var data = addonSheetLookup(mpn);
    var rows = (data && data.in_stock) || [];
    var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('IN STOCK — ' + mpn).setSubtitle(rows.length === 0 ? '✓ Removed' : '⚠ Still present'));
    var section = CardService.newCardSection();
    if (rows.length === 0) {
      section.addWidget(CardService.newTextParagraph().setText('✅ ' + mpn + ' is NOT in IN STOCK. Successfully removed.'));
    } else {
      section.addWidget(CardService.newTextParagraph().setText('⚠️ ' + mpn + ' still found — ' + rows.length + ' row(s):'));
      rows.forEach(function(r) {
        section.addWidget(CardService.newTextParagraph().setText('Row ' + r.row + ': QTY=' + r.qty + (r.notes ? ' | Notes=' + r.notes : '')));
      });
    }
    card.addSection(section);
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card.build())).build();
  } catch(err) { return notify('❌ Error: ' + err.toString()); }
}

function addonQuoteHistory(e) {
  try {
    var formInputs = (e.commonEventObject && e.commonEventObject.formInputs) || {};
    var mpn = (formInputs.invMpn && formInputs.invMpn.stringInputs && formInputs.invMpn.stringInputs.value && formInputs.invMpn.stringInputs.value[0] || '').trim().toUpperCase();
    if (!mpn) return notify('Enter an MPN first.');
    var data = addonSheetLookup(mpn);
    var forte = (data && data.forte_sheet) || [];
    var oem   = (data && data.oem_excess)  || [];
    var stock = (data && data.in_stock)    || [];
    var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Quote History — ' + mpn)
        .setSubtitle(forte.length + ' prior quote' + (forte.length !== 1 ? 's' : '') + ' in Forte'));
    var stockParts = [];
    if (oem.length)   stockParts.push('OEM: ' + oem.reduce(function(s,r){return s+(parseInt(r.qty)||0);},0).toLocaleString() + ' pcs');
    if (stock.length) stockParts.push('InStock: ' + stock.reduce(function(s,r){return s+(parseInt(r.qty)||0);},0).toLocaleString() + ' pcs');
    var stockSection = CardService.newCardSection().setHeader('Current Stock');
    stockSection.addWidget(CardService.newTextParagraph().setText(stockParts.length ? stockParts.join(' · ') : 'Not in any sheet'));
    card.addSection(stockSection);
    if (!forte.length) {
      var noSection = CardService.newCardSection();
      noSection.addWidget(CardService.newTextParagraph().setText('No prior Forte entries found.'));
      card.addSection(noSection);
    } else {
      var sixMonthsAgo = new Date(); sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
      forte.forEach(function(r) {
        var d = new Date(r.date); var stale = isNaN(d.getTime()) || d < sixMonthsAgo;
        var pot = (r.qty && r.buyerTP) ? '$' + (parseFloat(r.qty) * parseFloat(r.buyerTP)).toFixed(2) : '—';
        var entrySection = CardService.newCardSection().setHeader((stale ? '⚠ STALE — ' : '') + (r.date || '—') + ' · ' + (r.status || '—'));
        entrySection.addWidget(CardService.newTextParagraph()
          .setText('QTY: ' + (r.qty||'—') + '  TP: ' + (r.buyerTP?'$'+r.buyerTP:'—') + '  Potential: ' + pot + '  Country: ' + (r.country||'—')));
        card.addSection(entrySection);
      });
    }
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card.build())).build();
  } catch(err) { return notify('❌ Error: ' + err.toString()); }
}

function addonDiagnoseEmail(e) {
  try {
    var params    = e.commonEventObject.parameters;
    var threadId  = params.threadId  || '';
    var subject   = params.subject   || '';
    var fromH     = params.fromH     || '';

    // Read email body from thread (full body, not stripped — stripping can hide context)
    var content = '';
    if (threadId) {
      try {
        var thread = GmailApp.getThreadById(threadId);
        var msgs = thread ? thread.getMessages() : [];
        if (msgs.length) {
          var raw = msgs[msgs.length - 1].getPlainBody();
          content = raw.substring(0, 1000);
        }
      } catch(e2) {}
    }

    // Extract MPN — handle netCOMPONENTS "Company | MPN)" format first
    var mpn = '';
    var ncPipe = subject.match(/\|\s*([A-Z0-9][A-Z0-9\-\.\/]{3,})\s*\)/i);
    if (ncPipe) mpn = ncPipe[1];
    else mpn = extractMPNFromSubject(subject) || extractMPN(subject) || '';

    var oem_results = [], in_stock_results = [], forte_results = [];
    if (mpn) {
      try {
        var inv = addonSheetLookup(mpn);
        oem_results      = (inv && inv.oem_excess)  || [];
        in_stock_results = (inv && inv.in_stock)    || [];
        forte_results    = (inv && inv.forte_sheet) || [];
      } catch(e3) {}
    }

    // Call /api/diagnose
    var diagResp = UrlFetchApp.fetch(HUB_URL + '/api/diagnose', {
      method: 'POST', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({ subject: subject, sender: fromH, content: content, oem_results: oem_results, in_stock_results: in_stock_results, forte_results: forte_results }),
      muteHttpExceptions: true
    });
    var result = JSON.parse(diagResp.getContentText());

    var confLabel = result.confidence === 'high' ? '🟢 High' : result.confidence === 'low' ? '🔴 Low' : '🟡 Medium';
    var options = result.reply_options || [];
    var needsScript = result.needs_script_change === true;
    var scriptNote = result.script_change_note || '';

    var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('🧠 Jiggle My Mind').setSubtitle(mpn || subject.substring(0,40)));

    // Analysis — brief
    var diagSection = CardService.newCardSection().setHeader('Analysis — ' + confLabel);
    diagSection.addWidget(CardService.newTextParagraph().setText('Should be: ' + (result.action_should_have_been || '—')));
    diagSection.addWidget(CardService.newTextParagraph().setText('Why missed: ' + (result.reason_missed || '—')));
    if (needsScript) {
      diagSection.addWidget(CardService.newTextParagraph().setText('⚠️ Code fix needed: ' + scriptNote));
    }
    card.addSection(diagSection);

    // Reply options — one section per option with preview + button
    if (options.length) {
      var optSection = CardService.newCardSection().setHeader('Choose the correct reply:');
      options.forEach(function(opt, i) {
        var preview = (opt.draft || '(No reply sent)').substring(0, 130);
        if ((opt.draft || '').length > 130) preview += '…';
        optSection.addWidget(CardService.newTextParagraph().setText((i === 0 ? '▶ ' : '  ') + (opt.label || opt.action) + '\n"' + preview + '"'));
        if (opt.action !== 'no_bid' && opt.action !== 'decline' && opt.draft && opt.draft !== '(No reply sent)') {
          optSection.addWidget(CardService.newTextButton()
            .setText('✓ Use: ' + (opt.label || opt.action))
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setBackgroundColor('#1a7340')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('addonUseReplyOption')
              .setParameters({
                threadId: threadId, subject: subject, fromH: fromH,
                action: opt.action || '',
                optLabel: (opt.label || opt.action).substring(0, 80),
                draft: (opt.draft || '').substring(0, 500),
                diagAction: result.action_should_have_been || '',
                diagReason: (result.reason_missed || '').substring(0, 200),
                needsScript: needsScript ? 'true' : 'false',
                scriptNote: scriptNote.substring(0, 200)
              })));
        } else {
          optSection.addWidget(CardService.newTextButton()
            .setText('✓ Log: No reply sent')
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setBackgroundColor('#37474f')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('addonLogDiagnosis')
              .setParameters({
                subject: subject, fromH: fromH,
                action: opt.action || result.action_should_have_been || '',
                trigger: result.trigger_responsible || '',
                reason: (result.reason_missed || '').substring(0, 200),
                agreed: 'true',
                needsScript: needsScript ? 'true' : 'false',
                scriptNote: scriptNote.substring(0, 200)
              })));
        }
      });
      card.addSection(optSection);
    }

    // Correction fallback
    var corrSection = CardService.newCardSection().setHeader('Or — enter your own correction:');
    corrSection.addWidget(CardService.newTextInput()
      .setFieldName('diagCorrection')
      .setTitle('What should have happened?')
      .setHint('e.g. "Should have been request_tp_500 — BCM5338 is in OEM EXCESS"')
      .setMultiline(false));
    corrSection.addWidget(CardService.newTextButton()
      .setText('💬 Log Correction as Training')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#37474f')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonLogDiagnosis')
        .setParameters({
          subject: subject, fromH: fromH,
          action: result.action_should_have_been || '',
          trigger: result.trigger_responsible || '',
          reason: (result.reason_missed || '').substring(0, 200),
          agreed: 'false',
          needsScript: needsScript ? 'true' : 'false',
          scriptNote: scriptNote.substring(0, 200)
        })));
    card.addSection(corrSection);

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card.build())).build();
  } catch(err) { return notify('❌ Diagnose error: ' + err.toString()); }
}

function addonUseReplyOption(e) {
  try {
    var params   = e.commonEventObject.parameters;
    var threadId = params.threadId || '';
    var subject  = params.subject  || '';
    var fromH    = params.fromH    || '';
    var draft    = params.draft    || '';
    var action   = params.action   || '';
    var optLabel = params.optLabel || action;

    // Show preview card so John can edit before sending
    var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('📝 Preview Draft').setSubtitle(optLabel));

    var previewSection = CardService.newCardSection().setHeader('Edit before sending:');
    previewSection.addWidget(CardService.newTextInput()
      .setFieldName('draftContent')
      .setTitle('Draft body')
      .setValue(draft)
      .setMultiline(true));
    previewSection.addWidget(CardService.newTextInput()
      .setFieldName('editNote')
      .setTitle('Why did you change it? (optional — saved as training)')
      .setMultiline(false));
    previewSection.addWidget(CardService.newTextButton()
      .setText('✅ Create This Draft')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#1a7340')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonConfirmDraft')
        .setParameters({
          threadId: threadId, subject: subject, fromH: fromH,
          action: action, optLabel: optLabel,
          diagAction: params.diagAction || '',
          diagReason: params.diagReason || '',
          needsScript: params.needsScript || 'false',
          scriptNote: params.scriptNote || ''
        })));
    card.addSection(previewSection);

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card.build())).build();
  } catch(err) { return notify('Error showing preview: ' + err.toString()); }
}

function addonConfirmDraft(e) {
  try {
    var params     = e.commonEventObject.parameters;
    var formInputs = (e.commonEventObject && e.commonEventObject.formInputs) || {};
    var threadId   = params.threadId || '';
    var subject    = params.subject  || '';
    var fromH      = params.fromH    || '';
    var action     = params.action   || '';
    var optLabel   = params.optLabel || action;
    var draftContent = (formInputs.draftContent && formInputs.draftContent.stringInputs && formInputs.draftContent.stringInputs.value[0]) || '';
    var editNote     = (formInputs.editNote     && formInputs.editNote.stringInputs     && formInputs.editNote.stringInputs.value[0])     || '';

    var thread = GmailApp.getThreadById(threadId);
    if (!thread) return notify('Thread not found.');
    var msgs = thread.getMessages();
    var lastMsg = msgs[msgs.length - 1];
    var toEmail = extractBuyerEmail(fromH || lastMsg.getFrom());

    // For remove_oem: delete from OEM sheet + stamp Forte before drafting
    if (action === 'remove_oem') {
      var rmMpn = extractMPN(subject);
      if (rmMpn) {
        deletePart(rmMpn, subject);
        updateForteSheet(rmMpn);
        hubLog('run', 'Sidebar remove_oem: ' + rmMpn, {mpn: rmMpn, source: 'sidebar'});
      }
    }

    var html = buildSimpleHTML(draftContent.replace(/\n/g, '<br>'));
    createThreadedDraft(toEmail, 'Re: ' + subject, html, lastMsg.getId(), threadId, null);

    // Show training confirmation card
    var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('✅ Draft Created').setSubtitle(optLabel));
    var confirmSection = CardService.newCardSection().setHeader('Log as training?');
    var trainingNote = 'Draft created using: ' + optLabel;
    if (editNote) trainingNote += '\n\nYour note: "' + editNote + '"';
    confirmSection.addWidget(CardService.newTextParagraph().setText(trainingNote + '\n\nSave this as training so the worker learns from it?'));
    confirmSection.addWidget(CardService.newTextButton()
      .setText('✓ Yes — Save to Worker Memory')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#1a7340')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonLogDiagnosis')
        .setParameters({
          subject: subject, fromH: fromH,
          action: action,
          trigger: '',
          reason: (params.diagReason || '') + (editNote ? ' | John note: ' + editNote : ''),
          agreed: 'true',
          needsScript: params.needsScript || 'false',
          scriptNote: params.scriptNote || ''
        })));
    confirmSection.addWidget(CardService.newTextButton()
      .setText('✗ Skip — Draft only')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonDismissCard')
        .setParameters({})));
    card.addSection(confirmSection);

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card.build())).build();
  } catch(err) { return notify('Error creating draft: ' + err.toString()); }
}

function addonLogDiagnosis(e) {
  try {
    var params = e.commonEventObject.parameters;
    var formInputs = (e.commonEventObject && e.commonEventObject.formInputs) || {};
    var correction = (formInputs.diagCorrection && formInputs.diagCorrection.stringInputs && formInputs.diagCorrection.stringInputs.value && formInputs.diagCorrection.stringInputs.value[0] || '').trim();
    var agreed = params.agreed === 'true';
    var finalAction = agreed ? (params.action || '') : (correction || params.action || '');
    var needsScript = params.needsScript === 'true';
    var scriptNote = params.scriptNote || '';

    var summary = agreed
      ? 'Jiggle My Mind — AGREED: ' + params.action + ' missed'
      : 'Jiggle My Mind — CORRECTION: ' + (correction || '(no text)');

    // 1. Log to hub
    UrlFetchApp.fetch(HUB_URL + '/api/logs', {
      method: 'POST', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({
        app_name: 'email_automation', event_type: 'training',
        summary: summary,
        details: JSON.stringify({ subject: params.subject, from: params.fromH, action: finalAction, trigger: params.trigger, reason: params.reason, john_correction: correction || null, agreed: agreed, needs_script_change: needsScript })
      }),
      muteHttpExceptions: true
    });

    // 2. Save to worker AI memory so it applies on the next email decision
    var memSlug = 'training-' + (finalAction || 'correction').replace(/[^a-z0-9]/gi, '-').toLowerCase() + '-' + new Date().getTime();
    var memBody = 'John corrected the automation:\n'
      + 'Subject: ' + (params.subject || '') + '\n'
      + 'From: ' + (params.fromH || '') + '\n'
      + 'Correct action: ' + finalAction + '\n'
      + 'Why missed: ' + (params.reason || '') + '\n'
      + (correction ? 'John\'s note: ' + correction : '');
    UrlFetchApp.fetch(HUB_URL + '/api/memory', {
      method: 'POST', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({
        slug: memSlug,
        description: 'Training: ' + (params.subject || '').substring(0, 80) + ' → ' + finalAction,
        type: 'training',
        body: memBody
      }),
      muteHttpExceptions: true
    });

    var msg = '✅ Saved to worker memory — agent will use this on next email.';
    if (needsScript && scriptNote) {
      msg += '\n\n⚠️ Apps Script update also needed:\n' + scriptNote;
    }
    return notify(msg);
  } catch(err) { return notify('❌ Error: ' + err.toString()); }
}

function addonDismissCard(e) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().popCard()).build();
}

function addonSmartReply(e) {
  try {
    var params   = e.commonEventObject.parameters;
    var threadId = params.threadId || '';
    var subject  = params.subject  || '';
    var fromH    = params.fromH    || '';

    // Read ALL messages in thread for full context
    var threadContext = '';
    if (threadId) {
      try {
        var thread = GmailApp.getThreadById(threadId);
        var msgs = thread ? thread.getMessages() : [];
        var parts = [];
        msgs.forEach(function(msg) {
          var from = msg.getFrom();
          var body = msg.getPlainBody().substring(0, 600);
          parts.push('FROM: ' + from + '\n' + body);
        });
        threadContext = parts.join('\n\n---\n\n').substring(0, 3000);
      } catch(e2) {}
    }

    // Extract MPN — handle netCOMPONENTS pipe format too
    var mpn = '';
    var ncPipe = subject.match(/\|\s*([A-Z0-9][A-Z0-9\-\.\/]{3,})\s*\)/i);
    if (ncPipe) mpn = ncPipe[1];
    else mpn = extractMPNFromSubject(subject) || extractMPN(subject) || '';

    var oem_results = [], in_stock_results = [], forte_results = [];
    if (mpn) {
      try {
        var inv = addonSheetLookup(mpn);
        oem_results      = (inv && inv.oem_excess)  || [];
        in_stock_results = (inv && inv.in_stock)    || [];
        forte_results    = (inv && inv.forte_sheet) || [];
      } catch(e3) {}
    }

    // Call /api/smart-reply
    var resp = UrlFetchApp.fetch(HUB_URL + '/api/smart-reply', {
      method: 'POST', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({
        subject: subject, sender: fromH,
        thread_context: threadContext,
        oem_results: oem_results, in_stock_results: in_stock_results, forte_results: forte_results
      }),
      muteHttpExceptions: true
    });
    var result = JSON.parse(resp.getContentText());
    if (result.error) return notify('Smart Reply error: ' + result.error);

    var replyText = result.reply_text || '';
    var reasoning = result.reasoning || '';

    var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('🤖 Smart Reply')
        .setSubtitle((result.action || '') + (mpn ? ' — ' + mpn : '')));

    if (reasoning) {
      var reasonSection = CardService.newCardSection();
      reasonSection.addWidget(CardService.newTextParagraph().setText('💡 ' + reasoning));
      card.addSection(reasonSection);
    }

    var replySection = CardService.newCardSection().setHeader('Suggested reply — edit and copy:');
    replySection.addWidget(CardService.newTextInput()
      .setFieldName('smart_reply_body')
      .setTitle('Reply text')
      .setMultiline(true)
      .setValue(replyText));
    card.addSection(replySection);

    var actSection = CardService.newCardSection();
    actSection.addWidget(CardService.newTextButton()
      .setText('✉ Create as Draft in Gmail')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#1a7340')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonSmartReplyCreateDraft')
        .setParameters({ threadId: threadId, subject: subject, toEmail: extractBuyerEmail(fromH) })));
    card.addSection(actSection);

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card.build())).build();
  } catch(err) { return notify('❌ Smart Reply error: ' + err.toString()); }
}

function addonSmartReplyCreateDraft(e) {
  try {
    var params     = e.commonEventObject.parameters;
    var threadId   = params.threadId   || '';
    var subject    = params.subject    || '';
    var toEmail    = params.toEmail    || '';
    var formInputs = (e.commonEventObject && e.commonEventObject.formInputs) || {};
    var body = (formInputs.smart_reply_body && formInputs.smart_reply_body.stringInputs && formInputs.smart_reply_body.stringInputs.value && formInputs.smart_reply_body.stringInputs.value[0] || '').trim();
    if (!body) return notify('No reply text — please type or generate a reply first.');

    var thread = GmailApp.getThreadById(threadId);
    if (!thread) return notify('Thread not found.');
    var msgs = thread.getMessages();
    var lastMsg = msgs[msgs.length - 1];
    if (!toEmail) toEmail = extractBuyerEmail(lastMsg.getReplyTo() || lastMsg.getFrom());

    var html = buildSimpleHTML(body.replace(/\n/g, '<br>'));
    createThreadedDraft(toEmail, 'Re: ' + subject, html, threadId, lastMsg.getId());
    hubLog('run', 'addonSmartReplyCreateDraft: draft created for ' + subject, { to: toEmail });

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().popCard())
      .setNotification(CardService.newNotification().setText('✅ Draft created — check Gmail!'))
      .build();
  } catch(err) { return notify('Error creating draft: ' + err.toString()); }
}

function addonDiagnoseDraft(e) {
  try {
    var params    = e.commonEventObject.parameters;
    var draftId   = params.draftId   || '';
    var threadId  = params.threadId  || '';
    var subject   = params.subject   || '';
    var fromH     = params.fromH     || '';
    var toEmail   = params.toEmail   || '';

    if (!draftId) return notify('No draft ID found.');

    // Fetch current draft body
    var token = ScriptApp.getOAuthToken();
    var fetchResp = UrlFetchApp.fetch(
      'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + draftId + '?format=full',
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (fetchResp.getResponseCode() !== 200) return notify('Could not load draft.');
    var draft = JSON.parse(fetchResp.getContentText());
    var htmlBody = extractDraftHtmlBody(draft.message && draft.message.payload);
    var draftBody = htmlBody ? htmlBody.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim() : '';

    // Read buyer's last message from thread
    var buyerContent = '';
    if (threadId) {
      try {
        var thread = GmailApp.getThreadById(threadId);
        var msgs = thread ? thread.getMessages() : [];
        if (msgs.length) buyerContent = stripQuotedLines(msgs[msgs.length - 1].getPlainBody()).substring(0, 800);
      } catch(e2) {}
    }

    // Look up inventory
    var mpn = extractMPNFromSubject(subject) || extractMPN(subject);
    var oem_results = [], in_stock_results = [], forte_results = [];
    if (mpn) {
      try {
        var inv = addonSheetLookup(mpn);
        oem_results      = (inv && inv.oem_excess)  || [];
        in_stock_results = (inv && inv.in_stock)    || [];
        forte_results    = (inv && inv.forte_sheet) || [];
      } catch(e3) {}
    }

    // Call /api/diagnose in draft mode
    var diagResp = UrlFetchApp.fetch(HUB_URL + '/api/diagnose', {
      method: 'POST', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({
        mode: 'draft',
        draft_body: draftBody,
        subject: subject, sender: fromH, content: buyerContent,
        oem_results: oem_results, in_stock_results: in_stock_results, forte_results: forte_results
      }),
      muteHttpExceptions: true
    });
    var result = JSON.parse(diagResp.getContentText());

    var confLabel = result.confidence === 'high' ? '🟢 High' : result.confidence === 'low' ? '🔴 Low' : '🟡 Medium';
    var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('🤔 Draft Diagnosis').setSubtitle(subject ? subject.substring(0,50) : 'Draft'));

    var diagSection = CardService.newCardSection().setHeader('What I found');
    diagSection.addWidget(CardService.newTextParagraph().setText('❌ Wrong: ' + (result.what_is_wrong || '—')));
    diagSection.addWidget(CardService.newTextParagraph().setText('✅ Should be: ' + (result.what_it_should_say || '—')));
    diagSection.addWidget(CardService.newTextParagraph().setText('Fix: ' + (result.corrected_instruction || '—')));
    diagSection.addWidget(CardService.newTextParagraph().setText('Confidence: ' + confLabel));
    card.addSection(diagSection);

    var agreeSection = CardService.newCardSection().setHeader('Your verdict');
    agreeSection.addWidget(CardService.newTextButton()
      .setText('✓ Agree — Fix the Draft')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#1a7340')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonAgreeAndFixDraft')
        .setParameters({
          draftId: draftId, threadId: threadId, subject: subject, toEmail: toEmail,
          correction: (result.corrected_instruction || '').substring(0, 250)
        })));
    agreeSection.addWidget(CardService.newTextInput()
      .setFieldName('draftDiagCorrection')
      .setTitle('Or type your own correction')
      .setHint('e.g. "It should be a decline — qty×TP < $500"')
      .setMultiline(false));
    agreeSection.addWidget(CardService.newTextButton()
      .setText('💬 Use My Correction')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#37474f')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonManualFixDraftFromDiag')
        .setParameters({ draftId: draftId, threadId: threadId, subject: subject, toEmail: toEmail })));
    card.addSection(agreeSection);

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card.build())).build();
  } catch(err) { return notify('❌ Diagnose draft error: ' + err.toString()); }
}

function addonAgreeAndFixDraft(e) {
  try {
    var params     = e.commonEventObject.parameters;
    var draftId    = params.draftId    || '';
    var threadId   = params.threadId   || '';
    var subject    = params.subject    || '';
    var toEmail    = params.toEmail    || '';
    var correction = params.correction || '';
    if (!correction.trim()) return notify('No correction found from diagnosis.');

    var token = ScriptApp.getOAuthToken();
    var fetchResp = UrlFetchApp.fetch(
      'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + draftId + '?format=full',
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (fetchResp.getResponseCode() !== 200) return notify('Could not load draft to fix.');
    var draft = JSON.parse(fetchResp.getContentText());
    var htmlBody = extractDraftHtmlBody(draft.message && draft.message.payload);
    var currentBody = htmlBody ? htmlBody.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim() : '';

    var fixResp = UrlFetchApp.fetch(HUB_URL + '/api/fix-draft', {
      method: 'POST', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({ draft_body: currentBody, feedback: correction, subject: subject, to_email: toEmail, thread_id: threadId }),
      muteHttpExceptions: true
    });
    if (fixResp.getResponseCode() !== 200) return notify('Fix failed: ' + fixResp.getContentText().substring(0, 100));
    var fixResult = JSON.parse(fixResp.getContentText());
    var correctedBody = fixResult.corrected_body || '';

    var newHtml = buildSimpleHTML(correctedBody.replace(/\n/g, '<br>'));
    var rebuilt = rebuildRawMessage(draft, newHtml);
    var putResp = UrlFetchApp.fetch(
      'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + draftId,
      { method: 'PUT', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify({ message: { raw: rebuilt.raw, threadId: threadId || undefined } }),
        muteHttpExceptions: true }
    );
    if (JSON.parse(putResp.getContentText()).error) return notify('Could not update draft.');

    hubPostDraft(threadId, null, toEmail, subject, correctedBody, draftId, 'Auto-diagnosis fix: ' + correction);
    hubLearn('auto-diagnosis: ' + correction, currentBody, correctedBody, threadId, subject, toEmail, null, null);
    hubLog('run', 'addonAgreeAndFixDraft: draft corrected via auto-diagnosis — ' + subject, { correction: correction });

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('✅ Draft fixed! 🧠 Mistake logged for training.'))
      .build();
  } catch(err) { return notify('Error: ' + err.toString()); }
}

function addonManualFixDraftFromDiag(e) {
  var formInputs = (e.commonEventObject && e.commonEventObject.formInputs) || {};
  var correction = (formInputs.draftDiagCorrection && formInputs.draftDiagCorrection.stringInputs && formInputs.draftDiagCorrection.stringInputs.value && formInputs.draftDiagCorrection.stringInputs.value[0] || '').trim();
  if (!correction) return notify('Please type your correction first.');
  var params = e.commonEventObject.parameters;
  // Build a synthetic params with correction injected, reuse agreeAndFix logic
  e.commonEventObject.parameters.correction = correction;
  return addonAgreeAndFixDraft(e);
}

function addonReportIssue(e) {
  try {
    var params      = e.commonEventObject.parameters;
    var formInputs  = e.commonEventObject.formInputs || {};
    var description = (formInputs.issueDescription || {}).stringInputs
      ? formInputs.issueDescription.stringInputs.value[0] : '';
    if (!description || description.trim().length < 5) {
      return notify('Please describe the issue before reporting.');
    }
    var threadId = params.threadId || '';
    var subject  = params.subject  || '';
    var mpn      = params.mpn      || '';

    // Gather context: last agent decision for this thread
    var context = null;
    try {
      var decResp = UrlFetchApp.fetch(HUB_URL + '/api/agent-decisions?thread_id=' + encodeURIComponent(threadId), {
        headers: { Authorization: 'Bearer ' + HUB_SECRET }, muteHttpExceptions: true
      });
      var decData = JSON.parse(decResp.getContentText());
      var decisions = (decData.decisions || []);
      if (decisions.length) context = { last_decision: decisions[0] };
    } catch(ce) {}

    // Post the issue
    var resp = UrlFetchApp.fetch(HUB_URL + '/api/issues', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({
        thread_id: threadId, mpn: mpn,
        description: 'Subject: ' + subject + '\nMPN: ' + mpn + '\n\n' + description,
        context: context
      }),
      muteHttpExceptions: true
    });
    var data = JSON.parse(resp.getContentText());
    if (!data.ok) return notify('Failed to log issue: ' + (data.error || 'unknown'));
    var issueId = data.id;

    // Trigger self-heal immediately
    var healResp = UrlFetchApp.fetch(HUB_URL + '/api/self-heal', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({ issue_id: issueId }),
      muteHttpExceptions: true
    });
    var healData = JSON.parse(healResp.getContentText());

    var builder = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Self-Heal').setSubtitle('Issue #' + issueId));

    var section = CardService.newCardSection();
    if (healData.ok) {
      section.addWidget(CardService.newTextParagraph().setText(
        '✅ Fix pushed to GitHub\n\n' +
        '📝 ' + (healData.explanation || '') + '\n\n' +
        '⏳ GitHub Actions is deploying now — takes about 60 seconds.\n\n' +
        'Commit: ' + (healData.commit || 'pending').substring(0, 8)
      ));
    } else {
      section.addWidget(CardService.newTextParagraph().setText(
        '⚠️ Fix attempt failed:\n\n' + (healData.error || JSON.stringify(healData))
      ));
    }
    section.addWidget(CardService.newTextButton()
      .setText('← Back')
      .setOnClickAction(CardService.newAction().setFunctionName('addonDismissCard')));
    builder.addSection(section);

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(builder.build()))
      .build();

  } catch(err) {
    return notify('Report issue error: ' + err.toString());
  }
}

function addonExecuteAction(e) {
  try {
    var params   = e.commonEventObject.parameters;
    var threadId = params.threadId || '';
    var subject  = params.subject  || '';
    var fromH    = params.fromH    || '';
    var action   = JSON.parse(params.actionJson || '{}');

    if (!action.type) return notify('No action to execute.');
    var result = executeAction(action, threadId, subject, fromH);
    var msg = result.message || (result.ok ? '✅ Done' : '⛔ Failed');
    if (action.type === 'create_draft') msg += ' — Press back then reopen to review it.';
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(msg))
      .build();

  } catch(err) {
    return notify('Error: ' + err.toString());
  }
}

function addonExecuteAndSend(e) {
  try {
    var params   = e.commonEventObject.parameters;
    var threadId = params.threadId || '';
    var subject  = params.subject  || '';
    var fromH    = params.fromH    || '';
    var action   = JSON.parse(params.actionJson || '{}');

    if (action.type !== 'create_draft' && action.type !== undefined) {
      return notify('Create & Send is only for email drafts.');
    }
    var result = executeCreateAndSendDraft(action, threadId, subject, fromH);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(result.message))
      .build();

  } catch(err) {
    return notify('Error sending: ' + err.toString());
  }
}

function notify(text) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(text))
    .build();
}

// ── Test / utility functions ──────────────────────────────────
function testSearch(mpn) {
  Logger.log('--- OEM EXCESS ---'); searchOEMExcess(mpn);
  Logger.log('--- IN STOCK ---'); searchInStock(mpn);
  Logger.log('--- STAN SHEET ---'); searchStanSheet(mpn);
}
function testNewRFQs() { runEmailScan(); }
function testTPReplies() { runEmailScan(); }
function testSentChecking() { Logger.log('Replaced by executeDecision — Forte added inline.'); }

function unlabelUnprocessedRFQs() {
  var label = GmailApp.getUserLabelByName('oem-rfq-incoming-processed');
  if (!label) return;
  var threads = GmailApp.search('label:oem-rfq-incoming-processed in:inbox -in:sent newer_than:7d', 0, 50);
  threads.forEach(function(t) {
    var msgs = t.getMessages();
    var hasJohnReply = msgs.some(function(m){ return m.getFrom().indexOf('john.fluman@intransittech.com') >= 0; });
    if (!hasJohnReply) {
      t.removeLabel(label);
      Logger.log('Unlabeled: ' + t.getFirstMessageSubject());
    }
  });
}

// ── Trigger 8 — Bill netcomp removals → delete from OEM EXCESS
function checkBillNetcompRemovals() {
  var _cfg = getRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) { hubLog('run', 'checkBillNetcompRemovals: disabled'); return; }
  var BILL_EMAIL = 'bill.pratt@intransittech.com';
  var DONE_LABEL = 'oem-bill-removal-processed';
  var query = 'from:' + BILL_EMAIL +
    ' (netcomp OR netcomponents) (remove OR removing OR removed)' +
    ' -label:' + DONE_LABEL + ' newer_than:14d';
  var threads = GmailApp.search(query, 0, 20);
  hubLog('run', 'checkBillNetcompRemovals: ' + threads.length + ' thread(s)');
  if (!threads.length) return;
  var doneLabel = GmailApp.getUserLabelByName(DONE_LABEL) || GmailApp.createLabel(DONE_LABEL);

  threads.forEach(function(thread) {
    var msgs    = thread.getMessages();
    var lastMsg = msgs[msgs.length - 1];
    var subject = thread.getFirstMessageSubject();

    // Primary: look for @John Fluman -MPN tag in the message body (most reliable)
    var mpn = null;
    try {
      var bodyText = lastMsg.getPlainBody();
      var tagMatch = bodyText.match(/@John\s+Fluman\s*[-–]\s*([A-Z0-9][A-Z0-9\-\.\/]{2,})/i);
      if (tagMatch) mpn = tagMatch[1].trim();
    } catch(e) {}

    // Fallback 1: extract from subject using standard MPN parser
    if (!mpn) mpn = extractMPNFromSubject(subject);
    // Fallback 2: raw subject for all-digit/dash MPNs (e.g. 900-13448-0020-000).
    // extractMPN() requires letters; Bill subjects are often just the bare part number.
    // Spaces prevent descriptive subjects from matching; length cap rules out long garbage.
    if (!mpn) {
      var rawSubj = subject.replace(/^(RE|FW|FWD):\s*/gi, '').trim();
      if (rawSubj.length >= 4 && rawSubj.length <= 40 && /^[A-Z0-9][A-Z0-9\-\.\/]{3,}$/i.test(rawSubj)) {
        mpn = rawSubj;
      }
    }

    if (mpn) {
      var result = deletePart(mpn, subject);
      var replyBody = buildSimpleHTML('Got it - removing ' + mpn + ' from OEM EXCESS now.');
      // Reply in-thread to Bill's message
      createThreadedDraft(BILL_EMAIL, 'Re: ' + subject, replyBody, lastMsg.getId(), thread.getId(), null);
      hubLog('run', 'Bill netcomp removal [' + result + ']: ' + mpn, { mpn: mpn });
      Logger.log('Bill removal [' + result + ']: ' + mpn);
      thread.moveToArchive();
    } else {
      Logger.log('Bill removal: could not extract MPN from body or subject: ' + subject);
      GmailApp.sendEmail(NOTIFY_EMAIL, 'OEM EXCESS: Bill removal — could not extract MPN',
        'Subject: "' + subject + '"\nhttps://mail.google.com/mail/u/0/#inbox/' + thread.getId());
    }
    thread.addLabel(doneLabel);
  });
}

// ── Trigger 7 — AI email agent ───────────────────────────────

var AGENT_LABEL = 'oem-agent-processed';


// Slim subject-only MPN extractor (sidebar/addon use)
function extractMPN(subject) {
  return extractMPNFromSubject(subject) || (function() {
    if (!subject) return null;
    var clean = subject.replace(/^(Re:|Fwd:|FW:|RE:|FWD:|\[EXTERNAL\]|Subject:)\s*/gi, '').replace(/^RFQ#?\s*/i, '').trim();
    var stopwords = ['no','stk','stock','removed','remove','out','of','the','a','an','is','has','for','from','please','and','or','not','new','update','cant','share','rfq','quote','quotation','request','inquiry','inquire','netcomponents','member','price','target','pcs','qty','quantity','external','ics','source','on','standard','subject','requirements'];
    var tokens = clean.split(/\s+/);
    for (var i = 0; i < tokens.length; i++) {
      var token = tokens[i].replace(/[,;:?()[\]]/g, '');
      if (!token || token.length < 3 || stopwords.indexOf(token.toLowerCase()) >= 0 || /^#\d+$/.test(token)) continue;
      return token;
    }
    return null;
  })();
}

function stripQuotedLines(text) {
  if (!text) return '';
  var lines = text.split('\n'), result = [];
  for (var i = 0; i < lines.length; i++) {
    var tr = lines[i].trim();
    if (tr.charAt(0) === '>') continue;
    if (/^From:\s/i.test(tr) || /^-{3,}\s*Original Message/i.test(tr) || /^On .+ wrote:/i.test(tr)) break;
    result.push(lines[i]);
  }
  return result.join('\n');
}

function extractMPNFromSubject(subject) {
  if (!subject) return null;
  var nc = subject.match(/\|\s*([A-Z0-9][A-Z0-9\-\/\.\s]{2,40})\s*\)?$/i);
  if (nc) return nc[1].trim().replace(/\s{2,}/g, ' ');
  var ic = subject.match(/RFQ[:\-\s]+([A-Z0-9][A-Z0-9\-\/\.]{4,})/i);
  if (ic) return ic[1].trim();
  return null;
}
