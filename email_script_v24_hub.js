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
var PENDING_LABEL     = 'oem-pending-process';

var MSG_NEED_TP_500  = 'We need a target price to proceed. Please note there is a $500 minimum line requirement. Once we have your target we will get back to you right away.';
var MSG_NEED_TP_2000 = 'We need a target price to proceed. Please note there is a $2,000 minimum line requirement. Once we have your target we will get back to you right away.';
var MSG_CHECKING     = 'We are checking on it now. If we get a response from the OEM, I will respond to you right away. If we do not respond back to you, please consider this a no bid. Thank you very much for the opportunity.';
var MSG_BILL         = 'Bill will help with this request';

var HUB_URL    = 'https://intransit-hub.intransit-sales.workers.dev';
var HUB_SECRET = 'InTransit!Hub#2026';



// ── Gmail Add-on: Sidebar Shell (Phase 7) ──────────────────────────────────────
// All sidebar UI is now served by the Cloudflare Worker web app.
// This shell just opens that URL — zero quota risk, no CardService logic.

function buildGmailHomepage(e) {
  return buildContextualCard(e);
}


function buildContextualCard(e) {
  var threadId = (e && e.gmail && e.gmail.threadId) || "";
  try {
    var resp = UrlFetchApp.fetch(HUB_URL + "/api/sidebar/token", {
      method: "post", contentType: "application/json",
      headers: { Authorization: "Bearer " + HUB_SECRET },
      payload: JSON.stringify({ thread_id: threadId }),
      muteHttpExceptions: true
    });
    var data = {};
    try { data = JSON.parse(resp.getContentText()); } catch(pe) {}
    var workerUrl = data.url || (HUB_URL + "/sidebar");
    var section = CardService.newCardSection();
    section.addWidget(
      CardService.newTextButton()
        .setText("Open Intransit Assistant")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor("#1a3c6d")
        .setOpenLink(CardService.newOpenLink()
          .setUrl(workerUrl)
          .setOpenAs(CardService.OpenAs.FULL_SIZE)
          .setOnClose(CardService.OnClose.NOTHING))
    );
    if (threadId) {
      section.addWidget(CardService.newTextParagraph().setText("Thread loaded — click above to open the assistant."));
    }
    section.addWidget(
      CardService.newTextButton()
        .setText("Fix Claude Drafts")
        .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
        .setOnClickAction(CardService.newAction().setFunctionName("addonFixClaudeDrafts"))
    );
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("Intransit Hub"))
      .addSection(section)
      .build();
  } catch(err) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("Intransit Hub"))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText("Error: " + err.toString())))
      .build();
  }
}


function buildHomepageCard() {
  return buildContextualCard(null);
}


function buildComposeCard(e) {
  return buildContextualCard(null);
}


// Card action: runs synchronously — no trigger creation, no trigger limit issues ever.
// Processes up to 22s worth of "claude" drafts and shows results inline.
// If there are more drafts than fit in 22s, click again to process the next batch.
function addonFixClaudeDrafts() {
  try {
    var startMs = new Date().getTime();
    var results = findAndFixClaudeDrafts(startMs);
    // Write every result to the persistent log sheet
    results.forEach(function(r) { appendFixLog_(r.subject, r.action, r.draft_created); });
    var msg;
    if (results.length === 0) {
      msg = 'No "claude" drafts found.';
    } else {
      msg = 'Processed ' + results.length + ' draft(s):\n' +
        results.map(function(r) {
          var icon = r.draft_created ? '✅' : (r.action === 'error' ? '❌' : '⚪');
          return icon + ' ' + (r.subject || '').substring(0, 40) + '\n   → ' + (r.action || 'done');
        }).join('\n');
      if (results._partial) msg += '\n\nMore remain — click again to continue.';
    }
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(msg))
      .build();
  } catch(err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Error: ' + err.toString()))
      .build();
  }
}

// Appends one row to "FixClaudeLog" sheet in the OEM EXCESS spreadsheet.
// Creates the sheet + header on first use. Keeps a rolling 30-day history.
function appendFixLog_(subject, action, draftCreated) {
  try {
    var ss    = SpreadsheetApp.openById('1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4');
    var sheet = ss.getSheetByName('FixClaudeLog');
    if (!sheet) {
      sheet = ss.insertSheet('FixClaudeLog');
      sheet.appendRow(['Timestamp', 'Subject', 'Action', 'Draft Created?']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    }
    sheet.appendRow([new Date(), subject || '', action || '', draftCreated ? 'YES' : 'NO']);
    // Prune rows older than 30 days (keep header row 1)
    var now = new Date().getTime();
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      var ts = data[i][0];
      if (ts instanceof Date && (now - ts.getTime()) > 30 * 24 * 3600 * 1000) {
        sheet.deleteRow(i + 1);
      }
    }
  } catch(e) {
    Logger.log('appendFixLog_ error: ' + e);
  }
}

// Standalone runner — call from Apps Script editor for manual/debug runs (no time limit)
function runClaudeDraftFix() {
  try {
    var results = findAndFixClaudeDrafts(null);
    hubLog('info', 'Fix Claude Drafts: processed ' + results.length + ' draft(s)', {});
    results.forEach(function(r) {
      hubLog('info', 'Fixed draft: ' + r.subject + ' → action=' + r.action, {});
    });
  } catch(e) {
    hubLog('error', 'runClaudeDraftFix error: ' + e, {});
  }
}


function buildHubCard_(statusText, isError) {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Intransit Hub'));
  var section = CardService.newCardSection();
  if (statusText) section.addWidget(CardService.newTextParagraph().setText(statusText));
  card.addSection(section);
  return card.build();
}


function buildSimpleHTML(bodyText) {
  return '<div dir="ltr">' + bodyText.replace(/\n/g, '<br>') + getSignatureHTML() + '</div>';
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






function createThreadedDraft(toEmail, subject, htmlBody, replyToGmailMsgId, threadId, ccEmail) {
  try {
    var thread = GmailApp.getThreadById(threadId);
    if (!thread) { Logger.log('createThreadedDraft: thread not found ' + threadId); return null; }
    var msgs = thread.getMessages();
    // Find the specific message to reply to, default to last message
    var replyMsg = msgs[msgs.length - 1];
    if (replyToGmailMsgId) {
      for (var i = msgs.length - 1; i >= 0; i--) {
        if (msgs[i].getId() === replyToGmailMsgId) { replyMsg = msgs[i]; break; }
      }
    }
    var opts = { htmlBody: htmlBody, name: 'John Fluman' };
    if (toEmail) opts.to = toEmail;
    if (ccEmail) opts.cc = ccEmail;
    var draft = replyMsg.createDraftReply('', opts);
    Logger.log('Draft created | To: ' + toEmail + ' | ' + subject);
    return draft.getId();
  } catch(e) {
    Logger.log('createThreadedDraft error: ' + e);
    return null;
  }
}


// ONE-TIME: Jul 17 2026 — full audit of Forte rows 4010+ with David no-stk in col L but col K still Open


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


function doGet(e) {
  // Sidebar now served by Cloudflare Worker — redirect if old URL is hit
  if ((e.parameter.page || '') === 'sidebar') {
    return HtmlService.createHtmlOutput('<meta http-equiv="refresh" content="0;url=https://intransit-hub.intransit-sales.workers.dev/sidebar">')
      .setTitle('Intransit Hub')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
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


function executeAddForte(action) {
  if (!action.mpn) return { ok: false, message: 'add_forte requires mpn.' };
  if (!action.qty) return { ok: false, message: 'add_forte requires qty — cardinal rule: never add without QTY.' };
  try {
    addToForteSheet(action.mpn, action.qty, action.tp || '', action.country || '', '');
    hubLog('run', 'executeAddForte: ' + action.mpn, { qty: action.qty, tp: action.tp });
    return { ok: true, message: '✅ Added ' + action.mpn + ' to Forte (QTY: ' + action.qty + ')' };
  } catch(e) { return { ok: false, message: 'Forte error: ' + e }; }
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


function executeDecision(decision, thread) {
  if (!decision || !decision.action) return;
  var action = decision.action;
  var messages = thread.getMessages();
  var lastMsg = messages[messages.length - 1];
  var threadId = thread.getId();
  var subject = thread.getFirstMessageSubject();
  if (action === 'no_action') { thread.markRead(); return; }
  if (action === 'no_bid' && !decision.draft_body) { thread.markRead(); return; }
  if (action === 'remove_oem') {
    if (decision.oem_delete_row) {
      deleteOemRow(decision.oem_delete_row);
    } else if (decision.mpn) {
      deletePart(decision.mpn, subject);
    }
    if (decision.mpn) updateForteSheet(decision.mpn);
  }
  if (action === 'add_to_stan') {
    var fe0 = decision.forte_entry || {};
    var stanMpn = fe0.mpn || decision.mpn || '';
    var stanCountry = fe0.country || decision.country || 'USA';
    var stanQty = fe0.qty || decision.qty || '';
    var stanTp = fe0.target_price || decision.target_price || '';
    if (stanMpn) addToStanSheet(stanMpn, stanCountry, stanQty, stanTp);
  }
  if (action === 'david_nostock') {
    if (decision.mpn) {
      var dRes = deletePart(decision.mpn, subject);
      hubLog('run', 'david_nostock: deletePart ' + decision.mpn + ' → ' + dRes, {});
      // Always stamp Forte col K regardless of whether part was found in OEM EXCESS
      updateForteSheet(decision.mpn);
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
    var isRelayAddr = function(addr) {
      return !addr || addr.indexOf('intransittech.com') >= 0 ||
             addr.indexOf('messagesend@netcomponents') >= 0 ||
             addr.indexOf('autosend@icsource') >= 0;
    };
    if (isRelayAddr(replyTo)) {
      for (var i = messages.length - 1; i >= 0; i--) {
        var candidate = extractBuyerEmail(messages[i].getFrom());
        if (!isRelayAddr(candidate)) { replyTo = candidate; break; }
        // For ICS/relay emails, try Reply-To header
        var rt = extractBuyerEmail(messages[i].getReplyTo() || '');
        if (!isRelayAddr(rt)) { replyTo = rt; break; }
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
    var htmlBody = buildSimpleHTML(bodyText);
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
  if (action === 'no_bid') thread.markRead();
  if (decision._corrected_from) {
    try {
      GmailApp.sendEmail(NOTIFY_EMAIL,
        'Bug Auto-Corrected: [' + decision._corrected_from + '->' + action + '] ' + subject,
        'MPN: ' + (decision.mpn || '?') + '\nOriginal: ' + decision._corrected_from + '\nCorrected to: ' + action + '\nReason: ' + (decision._correction_reason || '?')
      );
    } catch(e) {}
  }
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


function executeRemoveOemExcess(action) {
  if (!action.mpn) return { ok: false, message: 'remove_oem_excess requires mpn.' };
  try {
    deletePart(action.mpn, 'sidebar-remove');
    hubLog('run', 'executeRemoveOemExcess: ' + action.mpn, {});
    return { ok: true, message: '✅ Removed ' + action.mpn + ' from OEM EXCESS' };
  } catch(e) { return { ok: false, message: 'Remove error: ' + e }; }
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


function extractAdviceText(htmlBody) {
  if (!htmlBody) return null;
  var m = htmlBody.match(/Note for John \(remove before sending\):<\/b><br>([\s\S]*?)<\/div>/);
  return m ? m[1].replace(/<[^>]+>/g, '').trim() : null;
}


// ── Slim worker bridge ───────────────────────────────────────

function extractBuyerEmail(fromRaw) {
  if (!fromRaw) return '';
  var m = fromRaw.match(/<([^>]+)>/);
  return m ? m[1].trim() : fromRaw.trim();
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
      if (/^\d+(?:pcs?|pc|k|m|units?)?$/i.test(token)) continue; // skip quantity tokens like "15pcs", "100k"
      return token;
    }
    return null;
  })();
}


function extractMPNFromSubject(subject) {
  if (!subject) return null;
  var nc = subject.match(/\|\s*([A-Z0-9][A-Z0-9\-\/\.#\+\s]{2,40})\s*\)?$/i);
  if (nc) return nc[1].trim().replace(/\s{2,}/g, ' ');
  var ic = subject.match(/RFQ[:\-\s]+([A-Z0-9][A-Z0-9\-\/\.#\+]{4,})/i);
  if (ic) return ic[1].trim();
  // "RFQ for Xpcs of MPN" — e.g. "RFQ for 15pcs of XCF08PVOG48C"
  var ofMpn = subject.match(/\bof\s+([A-Z][A-Z0-9\-\/\.#\+]{4,})\s*$/i);
  if (ofMpn && /[0-9]/.test(ofMpn[1])) return ofMpn[1].trim();
  return null;
}


// Extract an MPN-like token from free-form chat message text
function extractMPNFromText(text) {
  if (!text) return null;
  var SKIP = ['quote', 'history', 'price', 'stock', 'check', 'please', 'have', 'what',
              'this', 'that', 'with', 'from', 'send', 'need', 'order', 'about', 'more',
              'the', 'for', 'and', 'get', 'can', 'look', 'into', 'any', 'how', 'much'];
  var tokens = text.split(/\s+/);
  for (var i = 0; i < tokens.length; i++) {
    var tok = tokens[i].replace(/[^A-Za-z0-9\-\/\.#]/g, '').toUpperCase();
    if (tok.length < 5) continue;
    if (SKIP.indexOf(tok.toLowerCase()) >= 0) continue;
    if (/[A-Z]/.test(tok) && /[0-9]/.test(tok)) return tok;
  }
  return null;
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
      if (qtyCol >= 0) foundHeader = true;
      continue;
    }
    if (foundHeader && vals.length > qtyCol) {
      var qty = parseInt((vals[qtyCol] || '').replace(/,/g, ''), 10);
      if (!isNaN(qty) && qty > 0) {
        var tp = (tpCol >= 0 && vals.length > tpCol) ? parseFloat((vals[tpCol] || '').replace(/[$,\s]/g, '')) : NaN;
        return { qtyReq: qty, tgtPrice: (!isNaN(tp) && tp > 0) ? tp : null };
      }
    }
  }
  return null;
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


// ─────────────────────────────────────────────────────────────────────────────

function getRecentSentQuotesFull(mpn, maxThreads) {
  if (!mpn) return 'No MPN provided.';
  try {
    var resp = UrlFetchApp.fetch(HUB_URL + '/api/gmail/sent-quotes?mpn=' + encodeURIComponent(mpn) + '&max=' + (maxThreads || 5), {
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return 'Email search error: HTTP ' + resp.getResponseCode();
    var data = JSON.parse(resp.getContentText());
    var quotes = data.quotes || [];
    return quotes.length ? quotes.join('\n\n') : 'No prior sent emails found for ' + mpn + '.';
  } catch(e) {
    return 'Email search error: ' + e.toString();
  }
}




// ── Sidebar: "Fix Claude Drafts" ─────────────────────────────


// ── Sidebar: "Fix Claude Drafts" — finds drafts where body starts with "claude",
//    deletes the placeholder, then reprocesses the thread through the full email agent
//    (same path as the automation: inventory lookup + Claude decision + correct draft).
// startMs: Date.getTime() from the card callback — used to stop before the 30s timeout.
// Pass null for manual/background runs (no time limit).
function findAndFixClaudeDrafts(startMs) {
  var TIME_LIMIT_MS = 22000;
  var results = [];
  var data = gmailREST_('/drafts?maxResults=50');
  var drafts = (data.drafts || []);

  for (var i = 0; i < drafts.length; i++) {
    // Stop before hitting the 30s card callback timeout
    if (startMs && (new Date().getTime() - startMs) > TIME_LIMIT_MS) {
      results._partial = true;
      break;
    }
    var d = drafts[i];
    var msgId = d.message && d.message.id;
    if (!msgId) continue;
    try {
      var msg = GmailApp.getMessageById(msgId);
      if (!msg) continue;
      var body = (msg.getPlainBody() || '').trim();
      if (!/^claude\b/i.test(body)) continue;

      var thread  = msg.getThread();
      var subject = thread.getFirstMessageSubject();

      hubLog('info', 'FixClaudeDraft START: ' + subject, {draft_id: d.id});

      // CRITICAL: run processThread FIRST — it creates the new draft if appropriate.
      // Only delete the placeholder AFTER we know what happened.
      // If processThread throws, leave the placeholder intact so John can retry.
      var decision = null;
      try {
        decision = processThread(thread);
      } catch(procErr) {
        hubLog('error', 'FixClaudeDraft ERROR: "' + subject + '" — ' + procErr, {draft_id: d.id});
        results.push({ subject: subject, action: 'error', draft_created: false });
        continue; // Leave placeholder intact
      }

      var action       = decision ? (decision.action || 'unknown') : 'null';
      var draftCreated = !!(decision && decision.draft_body &&
                           action !== 'no_bid' && action !== 'no_action' && action !== 'no_longer_available');

      hubLog('info', 'FixClaudeDraft DONE: "' + subject + '" → ' + action + (draftCreated ? ' ✓ draft' : ' (no draft)'), {draft_id: d.id, action: action});

      // Now safe to delete the placeholder — the new draft (if any) is already created above
      try { gmailREST_('/drafts/' + d.id, 'DELETE'); } catch(delErr) {
        hubLog('warn', 'FixClaudeDraft: could not delete placeholder ' + d.id + ': ' + delErr, {});
      }

      results.push({ subject: subject, action: action, draft_created: draftCreated });
    } catch(e) {
      hubLog('error', 'findAndFixClaudeDrafts outer: ' + e, {});
    }
  }
  return results;
}


function applyClaudeDraftFix(draftId, threadId, toEmail, subject, correctedBody) {
  if (!correctedBody || correctedBody === '(could not determine)') {
    throw new Error('No valid corrected body — draft left unchanged');
  }
  if (!toEmail) throw new Error('No recipient — draft left unchanged');

  // Create new draft FIRST, then delete old one so we never lose the draft
  var resp = UrlFetchApp.fetch(HUB_URL + '/api/gmail/draft', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      to:        toEmail,
      subject:   subject,
      body:      correctedBody,
      thread_id: threadId
    }),
    headers: { Authorization: 'Bearer ' + HUB_SECRET },
    muteHttpExceptions: true
  });
  var result = {};
  try { result = JSON.parse(resp.getContentText()); } catch(e) {}
  if (!result.ok) throw new Error('Draft creation failed: ' + JSON.stringify(result));

  // Only delete the old draft once the new one is confirmed created
  gmailREST_('/drafts/' + draftId, 'DELETE');

  return { ok: true, draft_id: result.draft_id };
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


function gmailArchiveThread_(threadId) {
  try {
    var thread = GmailApp.getThreadById(threadId);
    if (thread) thread.moveToArchive();
  } catch(e) { Logger.log('gmailArchiveThread_ error: ' + e); }
}




function gmailModifyThread_(threadId, addLabels, removeLabels) {
  // Use GmailApp label methods — reliable, no REST scope required
  try {
    var thread = GmailApp.getThreadById(threadId);
    if (!thread) return;
    (addLabels || []).forEach(function(name) {
      try {
        var lbl = GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
        thread.addLabel(lbl);
      } catch(e) { Logger.log('gmailModifyThread_ addLabel error ' + name + ': ' + e); }
    });
    (removeLabels || []).forEach(function(name) {
      try {
        var lbl = GmailApp.getUserLabelByName(name);
        if (lbl) thread.removeLabel(lbl);
      } catch(e) { Logger.log('gmailModifyThread_ removeLabel error ' + name + ': ' + e); }
    });
  } catch(e) { Logger.log('gmailModifyThread_ error: ' + e); }
}


// ── Gmail REST helpers — bypass premium GmailApp.search() quota ──
// GmailApp.search() exhausts the Apps Script "premium gmail" daily quota.
// These helpers use UrlFetchApp (100k calls/day quota) instead.

function gmailREST_(path, method, body) {
  var opts = {
    method: method || 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  if (body) { opts.contentType = 'application/json'; opts.payload = JSON.stringify(body); }
  var resp = UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/me' + path, opts);
  var text = resp.getContentText();
  if (!text || !text.trim()) return {};   // 204 No Content (e.g. DELETE) returns empty body
  var data = JSON.parse(text);
  if (data.error) { hubLog('error', 'gmailREST_ ' + (method||'GET') + ' ' + path + ': ' + data.error.message, {}); }
  return data;
}


function gmailSearchREST(query, maxResults) {
  var data = gmailREST_('/threads?q=' + encodeURIComponent(query) + '&maxResults=' + (maxResults || 50));
  return (data.threads || []).map(function(t) { return t.id; });
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


function logDeletion(deletedSheet, rowData, emailSubject) {
  deletedSheet.appendRow([new Date(), rowData[0], rowData[1], rowData[2], rowData[3], rowData[4], emailSubject||'']);
}


// ── Utility ───────────────────────────────────────────────────

function normalize(pn) {
  return String(pn).trim().replace(/[-.:\/()\\s*+\\#_,]/g, '').toLowerCase();
}


function notify(text) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(text))
    .build();
}


// ── Command queue — inventory actions queued remotely ─────────
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
          var threadMsgs = thread.getMessages();
          var firstMsg = threadMsgs[0];
          var lastMsg  = threadMsgs[threadMsgs.length - 1];
          var replyMsg = fix.reply_to_msg_id
            ? (function() {
                for (var mi = threadMsgs.length - 1; mi >= 0; mi--) {
                  if (threadMsgs[mi].getId() === fix.reply_to_msg_id) return threadMsgs[mi];
                }
                return lastMsg;
              })()
            : lastMsg;
          // Use pre-built HTML if provided; otherwise build from plain draft_body
          var htmlBody = fix.html || (fix.draft_body != null
            ? buildSimpleHTML(String(fix.draft_body))
            : null);
          if (!htmlBody) throw new Error('fix-queue #' + fix.id + ': no html or draft_body');
          var draftId = createThreadedDraft(
            fix.to || fix.to_email, fix.subject, htmlBody, replyMsg.getId(), fix.thread_id, null
          );
          if (!draftId) throw new Error('createThreadedDraft returned null');

          // Ensure oem-tp-processed is on the thread so tpQ won't re-pick it
          gmailModifyThread_(fix.thread_id, ['oem-tp-processed'], []);

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


// ── Sidebar: "Process Next Email" button (called via google.script.run) ──
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

  // netCOMPONENTS: parse table in Apps Script (text-based, no raw HTML needed)
  var parsedRFQ = extractNetcompRFQ(messages);
  if (parsedRFQ) {
    // Only include TgtPrice when the table actually had one.
    var rLine = '[PARSED_RFQ: QtyReq=' + parsedRFQ.qtyReq;
    if (parsedRFQ.tgtPrice !== null) rLine += ', TgtPrice=' + parsedRFQ.tgtPrice;
    rLine += ']';
    content = rLine + '\n' + content;
  }

  var mpnHint = extractMPNFromSubject(subject) || extractMPN(subject) || null;
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
  // Send MPN hint to worker as fallback if Haiku extraction fails.
  // Only send if it looks like a real MPN (letters + digits, no pure-quantity tokens).
  if (mpnHint && /[A-Za-z]/.test(mpnHint) && /[0-9]/.test(mpnHint) && mpnHint.length >= 5
      && !/^\d+(?:pcs?|pc|k|m|units?)?$/i.test(mpnHint)) {
    payload.mpn = mpnHint;
  }
  // IC Source: send raw HTML to worker — parsing logic lives in worker.js
  var lastFromLC = (lastMsg.getFrom() || '').toLowerCase();
  if (lastFromLC.indexOf('icsource') >= 0 || lastFromLC.indexOf('autosend') >= 0) {
    payload.icsource_html = lastMsg.getBody() || '';
  }
  var decision = callWorker(payload);
  if (decision) executeDecision(decision, thread);
  return decision;
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


function searchInStock(mpn) {
  if (!mpn) return [];
  var searchNorm = normalize(mpn);
  var data = SpreadsheetApp.openById(IN_STOCK_ID).getSheets()[0].getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var cellNorm = normalize(String(data[i][0]));
    var reverseOk = searchNorm.startsWith(cellNorm) && cellNorm.length >= Math.ceil(searchNorm.length * 0.75);
    if (cellNorm.length >= 3 && (cellNorm === searchNorm || cellNorm.startsWith(searchNorm) || reverseOk)) {
      results.push({ row: i+1, mpn: data[i][0], man: data[i][1], dc: data[i][2], qty: data[i][3], notes: data[i][4], price_to_quote: data[i][5] || '', price_history: data[i][9] || '' });
    }
  }
  if (results.length) Logger.log('IN STOCK FOUND ' + results.length + ': ' + mpn);
  else Logger.log('IN STOCK NOT FOUND: ' + mpn);
  return results;
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





// Run this NOW to send the Please Post email immediately (bypasses command queue).
function sendPleasePostNow() {
  var DATAMASTER_BCC = [
    '5BDFA5@stkdst.com',
    'datamaster@netcomponents.com',
    'post@icsource.com',
    'bill@intransittech.com',
    'david@fortetechno.com',
    'Stan@amorelectronics.com'
  ].join(',');

  var oemBlob = buildFilteredOemBlob_();
  var token = ScriptApp.getOAuthToken();
  var inGid = SpreadsheetApp.openById(IN_STOCK_ID).getSheets()[0].getSheetId();
  var inResp = UrlFetchApp.fetch(
    'https://docs.google.com/spreadsheets/d/' + IN_STOCK_ID + '/export?format=xlsx&gid=' + inGid,
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  if (inResp.getResponseCode() !== 200) throw new Error('IN STOCK export failed (' + inResp.getResponseCode() + '): ' + inResp.getContentText().substring(0, 200));
  var inBlob = inResp.getBlob();
  inBlob.setName('IN STOCK.xlsx');
  sendPleasePostViaREST(token, oemBlob, inBlob, DATAMASTER_BCC);
  Logger.log('sendPleasePostNow: DONE');
}

// Exports OEM EXCESS as XLSX directly via the Sheets export URL (no DriveApp, no Drive API needed).
// Uses docs.google.com/spreadsheets/export which works with the OAuth token and Sheets scope alone.
function buildFilteredOemBlob_() {
  var gid = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0].getSheetId();
  var token = ScriptApp.getOAuthToken();
  var url = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID +
            '/export?format=xlsx&gid=' + gid;
  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('OEM EXCESS export failed (' + resp.getResponseCode() + '): ' + resp.getContentText().substring(0, 300));
  }
  var blob = resp.getBlob();
  blob.setName('OEM_EXCESS.xlsx');
  Logger.log('OEM EXCESS export: OK (' + blob.getBytes().length + ' bytes)');
  return blob;
}


// ── Please Post — REST API send (no GmailApp quota) ──────────────────────────

function sendPleasePostViaREST(token, oemBlob, inBlob, bccList) {
  var boundary = 'bnd' + Math.random().toString(36).slice(2, 18);
  var oemB64   = Utilities.base64Encode(oemBlob.getBytes());
  var inB64    = Utilities.base64Encode(inBlob.getBytes());

  var rawParts = [
    'MIME-Version: 1.0',
    'From: John Fluman <' + JOHN_EMAIL + '>',
    'To: ' + NOTIFY_EMAIL,
    'Bcc: ' + bccList,
    'Subject: Please post',
    'Content-Type: multipart/mixed; boundary="' + boundary + '"',
    '',
    '--' + boundary,
    'Content-Type: text/plain; charset=UTF-8',
    '',
    '',
    '--' + boundary,
    'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Transfer-Encoding: base64',
    'Content-Disposition: attachment; filename="OEM_EXCESS.xlsx"',
    '',
    oemB64,
    '--' + boundary,
    'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Transfer-Encoding: base64',
    'Content-Disposition: attachment; filename="IN STOCK.xlsx"',
    '',
    inB64,
    '--' + boundary + '--'
  ];

  var resp = UrlFetchApp.fetch(
    'https://gmail.googleapis.com/gmail/v1/users/me/messages/send',
    {
      method: 'POST',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ raw: Utilities.base64EncodeWebSafe(rawParts.join('\r\n')) }),
      muteHttpExceptions: true
    }
  );
  var data = JSON.parse(resp.getContentText());
  if (data.error) throw new Error('Gmail REST send failed: ' + JSON.stringify(data.error));
  Logger.log('Please post sent OK — message id: ' + data.id);
}


function sendReviewEmail(partNumber, emailSubject, matches) {
  var body = 'Notice for part: ' + partNumber + '\nSubject: "' + emailSubject + '"\n\n';
  if (matches.length > 0) {
    matches.forEach(function(m,i) { body += (i+1) + '. MPN: ' + m.data[0] + ' | QTY: ' + m.data[3] + '\n'; });
    body += '\nhttps://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID;
  } else { body += 'No match.\nhttps://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID; }
  GmailApp.sendEmail(NOTIFY_EMAIL, 'OEM EXCESS: Review needed for MPN ' + partNumber, body);
}





// ── Forte bulk import ─────────────────────────────────────────────────────────
// Run manually after process_forte_david_v3.py has written forte_processed_output.csv.
// Reads the CSV from Google Drive, preserves BILL EXT rows, replaces all other OEM EXCESS rows.
function importForteBulk() {
  var START = new Date();
  Logger.log('importForteBulk: starting...');

  // 1. Find the processed CSV in Drive
  var files = DriveApp.getFilesByName('forte_processed_output.csv');
  if (!files.hasNext()) {
    Logger.log('importForteBulk ERROR: forte_processed_output.csv not found in Drive');
    return;
  }
  var csvText = files.next().getBlob().getDataAsString('UTF-8');
  var csvRows = Utilities.parseCsv(csvText);
  var dataRows = csvRows.slice(1);  // skip header row
  Logger.log('importForteBulk: CSV parsed — ' + dataRows.length + ' data rows');

  // 2. Open OEM EXCESS sheet and snapshot current data
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var ws = ss.getSheetByName(MAIN_SHEET_NAME);
  var currentData = ws.getDataRange().getValues();
  var headerRow = currentData[0];

  // 3. Extract BILL EXT rows to preserve (col E = index 4)
  var billRows = [];
  for (var i = 1; i < currentData.length; i++) {
    if (String(currentData[i][4] || '').toUpperCase().indexOf('BILL EXT') >= 0) {
      billRows.push(currentData[i]);
    }
  }
  Logger.log('importForteBulk: preserving ' + billRows.length + ' BILL EXT rows');

  // 4. Convert CSV rows to sheet format [MPN, Man, DC, QTY, Notes]
  var newRows = dataRows.map(function(r) {
    return [r[0] || '', r[1] || '', '', parseInt(r[3], 10) || 0, r[4] || ''];
  });

  // 5. Build final data: header + BILL EXT + new Forte rows
  var allRows = [headerRow].concat(billRows).concat(newRows);

  // 6. Clear and rewrite in chunks of 5000 rows to avoid timeout
  ws.clearContents();
  var CHUNK = 5000;
  var numCols = allRows[0].length;
  for (var start = 0; start < allRows.length; start += CHUNK) {
    var chunk = allRows.slice(start, start + CHUNK);
    ws.getRange(start + 1, 1, chunk.length, numCols).setValues(chunk);
  }

  var elapsed = Math.round((new Date() - START) / 1000);
  Logger.log('importForteBulk: DONE in ' + elapsed + 's | ' +
    'Total rows: ' + allRows.length + ' | BILL EXT kept: ' + billRows.length +
    ' | New Forte rows: ' + newRows.length);
}


function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){ScriptApp.deleteTrigger(t);});
  // Phase 3: inbox scanning + thread processing → worker cron
  // Phase 4: David/Bill/PaymentAdvice → worker cron
  // Phase 5: forte_add/stan_add/oem_remove/daily cost report → worker cron (Sheets API)
  // Phase 6: processCommandQueue (all 8 inventory/draft command types) → worker cron
  // Apps Script: processFixQueue only (replace_draft backup via Gmail REST)
  ScriptApp.newTrigger('processFixQueue').timeBased().everyMinutes(5).create();
  Logger.log('1 trigger installed. All automation runs in worker cron.');
}


function stripAdviceFromHtml(htmlBody) {
  if (!htmlBody) return htmlBody;
  return htmlBody.replace(/<div style="background:#fff3cd[\s\S]*?<\/div>/, '').trim();
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


