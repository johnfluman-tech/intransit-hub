// ============================================================
// ALL ONE-TIME FIXES — add this as a NEW FILE in your existing
// Apps Script project (the one with email_script_v24_hub.js).
// Run each function ONE AT A TIME from the editor.
//
// 0. draftDavidRemovals()      — STANDALONE — draft replies to David for
//                                3 no-stock threads that were never replied to.
// FUNCTION LIST (run in this order):
//
// 1. removeNoStockParts()      — STANDALONE — delete 3 David no-stock parts
//                                from OEM EXCESS sheet. Run this first.
//
// 2. fixWrongNetcompDrafts()   — STANDALONE — fix BCM5461SA1KPFG and
//                                BCM5221A4KPTG wrong "need TP" drafts → own-stock quotes.
//
// 3. fixWrongOemDrafts()       — REQUIRES MAIN SCRIPT — fix 5 wrong OEM drafts
//                                (wrong TP amount, missing MSG_BILL for NRF52832-QFAA-R).
//
// 4. fixAdviceBlockDrafts()    — REQUIRES MAIN SCRIPT — remove old [ADVICE:] blocks
//                                from K3KL8L80CM-MGCT and MT25QL256ABA8ESF-0SIT drafts.
//
// 5. inboxFullAudit()          — REQUIRES MAIN SCRIPT — scan every inbox thread +
//                                every drafted thread. Fix BILL EXT routing, wrong
//                                TP amount, advice blocks, stale drafts. Full log output.
//
// ============================================================


// ============================================================
// FUNCTION 1: removeNoStockParts() — STANDALONE
// Deletes 3 parts David confirmed no-stock from OEM EXCESS sheet.
// Archives deleted rows to a "Deleted Rows" tab.
// ============================================================

// ============================================================
// FUNCTION 0: draftDavidRemovals() — STANDALONE
// Drafts replies to David for the 3 no-stock threads that were
// removed from OEM EXCESS but never replied to.
// ============================================================

function draftDavidRemovals() {
  var DAVID_EMAIL = 'david@fortetechno.com';
  var sig = '<br><br>Regards,<br>John Fluman<br>Intransit Technologies<br>'
    + 'john.fluman@intransittech.com<br>An ISO 9001 Certified Company<br>'
    + 'Toll (877) 677-5868 x101 - Local (949) 481-7935 x101';

  var threads = [
    { mpn: 'MCIMX535DVP1C2', subject: '3900' },
    { mpn: 'BD429B',          subject: '3904' },
    { mpn: '10M08SAM153I7G',  subject: '3896' },
  ];

  threads.forEach(function(item) {
    try {
      // Search by MPN or ticket number
      var results = GmailApp.search('from:' + DAVID_EMAIL + ' subject:' + item.mpn, 0, 5);
      if (!results.length) results = GmailApp.search('from:' + DAVID_EMAIL + ' subject:' + item.subject, 0, 5);
      if (!results.length) { Logger.log('Thread not found for ' + item.mpn); return; }

      var thread = results[0];
      var lastMsg = thread.getMessages()[thread.getMessages().length - 1];
      var replyText = 'Removing ' + item.mpn + ' from OEM EXCESS.';
      var replyHtml = '<div dir="ltr">' + replyText + sig + '</div>';
      lastMsg.createDraftReply(replyText, { htmlBody: replyHtml });
      Logger.log('Drafted reply to David for: ' + item.mpn);
    } catch(e) {
      Logger.log('draftDavidRemovals error for ' + item.mpn + ': ' + e);
    }
  });

  Logger.log('draftDavidRemovals: done. Check Drafts to review before sending.');
}


var _OEM_SHEET_ID   = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
var _OEM_SHEET_NAME = 'sheet1';

function removeNoStockParts() {
  var parts = [
    { mpn: 'MCIMX535DVP1C2', note: '#3900 David — No stock' },
    { mpn: 'BD429B',          note: '#3904 David — Cant find' },
    { mpn: '10M08SAM153I7G',  note: '#3896 David — NO STOCK'  },
  ];

  var ss       = SpreadsheetApp.openById(_OEM_SHEET_ID);
  var main     = ss.getSheetByName(_OEM_SHEET_NAME);
  var delSheet = _getOrMakeDeletedSheet(ss);
  var data     = main.getDataRange().getValues();

  parts.forEach(function(item) {
    var found = false;
    for (var i = data.length - 1; i >= 1; i--) {
      var cell = String(data[i][0]).trim();
      if (cell.toLowerCase() === item.mpn.toLowerCase()) {
        var archived = data[i].concat([new Date(), item.note]);
        delSheet.appendRow(archived);
        main.deleteRow(i + 1);
        Logger.log('DELETED row ' + (i + 1) + ': ' + item.mpn);
        found = true;
        break;
      }
    }
    if (!found) Logger.log('NOT FOUND (already removed?): ' + item.mpn);
  });

  Logger.log('removeNoStockParts: done. Check View > Logs.');
}

function _getOrMakeDeletedSheet(ss) {
  var s = ss.getSheetByName('Deleted Rows');
  if (!s) {
    s = ss.insertSheet('Deleted Rows');
    s.appendRow(['MPN', 'QTY', 'Notes', 'Deleted At', 'Reason']);
  }
  return s;
}


// ============================================================
// FUNCTION 2: fixWrongNetcompDrafts() — STANDALONE
// Fixes wrong "need TP" drafts for BCM5461SA1KPFG and BCM5221A4KPTG.
// These are our own stock — replaces drafts with correct own-stock quotes.
// ============================================================

function fixWrongNetcompDrafts() {
  var fixes = [
    {
      mpn:   'BCM5461SA1KPFG',
      dc:    '2010+',
      qty:   94,
      man:   'BROADCOM',
      notes: 'Also have OEM EXCESS (QTY 24,665) and Warehouse#3 stock available if needed.'
    },
    {
      mpn:   'BCM5221A4KPTG',
      dc:    '2130',
      qty:   115,
      man:   'BROADCOM',
      notes: 'Also have OEM EXCESS and Warehouse#3 stock available if needed.'
    }
  ];

  var sig = '<br><br>Regards,<br>John Fluman<br>Intransit Technologies<br>'
    + 'john.fluman@intransittech.com<br>An ISO 9001 Certified Company<br>'
    + 'Toll (877) 677-5868 x101 - Local (949) 481-7935 x101';

  var token = ScriptApp.getOAuthToken();

  fixes.forEach(function(fix) {
    try {
      var threads = GmailApp.search('in:inbox subject:' + fix.mpn, 0, 5);
      if (!threads.length) { Logger.log('Thread not found: ' + fix.mpn); return; }
      var thread  = threads[0];
      var lastMsg = thread.getMessages()[thread.getMessages().length - 1];

      // Delete existing drafts on this thread
      var deletedCount = 0;
      var draftsResp = UrlFetchApp.fetch(
        'https://gmail.googleapis.com/gmail/v1/users/me/drafts?maxResults=50',
        { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
      );
      var allDrafts = (JSON.parse(draftsResp.getContentText()).drafts) || [];
      allDrafts.forEach(function(d) {
        try {
          var detail = JSON.parse(UrlFetchApp.fetch(
            'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + d.id + '?format=metadata',
            { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
          ).getContentText());
          if (detail.message && detail.message.threadId === thread.getId()) {
            UrlFetchApp.fetch(
              'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + d.id,
              { method: 'DELETE', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
            );
            deletedCount++;
          }
        } catch(e) { Logger.log('Draft delete error: ' + e); }
      });

      // Create correct own-stock draft
      var bodyText = 'This is our stock\n\n'
        + 'MPN: ' + fix.mpn + '\n'
        + 'Manufacturer: ' + fix.man + '\n'
        + 'DC: ' + fix.dc + '\n'
        + 'QTY Available: ' + fix.qty + '\n'
        + 'Price: $[FILL IN]\n\n'
        + 'There is a $100 minimum on stock items.';
      var htmlBody = '<div dir="ltr">' + bodyText.replace(/\n/g, '<br>') + sig + '</div>';

      lastMsg.createDraftReply(bodyText, { htmlBody: htmlBody });
      Logger.log('Created own-stock draft for: ' + fix.mpn + ' (deleted ' + deletedCount + ' wrong draft(s))');
    } catch(e) {
      Logger.log('fixWrongNetcompDrafts error for ' + fix.mpn + ': ' + e);
    }
  });

  Logger.log('fixWrongNetcompDrafts: done');
}


// ============================================================
// FUNCTION 3: fixWrongOemDrafts() — REQUIRES MAIN SCRIPT
// Fix 5 wrong OEM drafts:
//   MT25QL256ABA8ESF-0SIT  — wrong $2000 draft → correct need_tp_500
//   5962-8687503XA          — wrong draft → correct need_tp_500
//   116119AN1699            — wrong draft → correct need_tp_500
//   MT47H128M16RT-25E:CTR  — wrong draft → correct need_tp_500
//   NRF52832-QFAA-R        — missing MSG_BILL (alice gave TP $0.80, BILL EXT) → create MSG_BILL
// ============================================================

function fixWrongOemDrafts() {
  var BILL_EMAIL      = 'bill.pratt@intransittech.com';
  var MSG_NEED_TP_500 = 'We need a target price to proceed. Please note there is a $500 minimum line requirement. Once we have your target we will get back to you right away.';
  var MSG_BILL_TEXT   = 'Bill will help with this request';
  var token           = ScriptApp.getOAuthToken();

  var sig = (typeof getSignatureHTML === 'function') ? getSignatureHTML() :
    '<br><br>Regards,<br>John Fluman<br>Intransit Technologies<br>'
    + 'john.fluman@intransittech.com<br>An ISO 9001 Certified Company<br>'
    + 'Toll (877) 677-5868 x101 - Local (949) 481-7935 x101';

  function deleteDraftsForThread(threadId) {
    var count = 0;
    try {
      var drafts = (JSON.parse(UrlFetchApp.fetch(
        'https://gmail.googleapis.com/gmail/v1/users/me/drafts?maxResults=50',
        { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
      ).getContentText()).drafts) || [];
      drafts.forEach(function(d) {
        try {
          var detail = JSON.parse(UrlFetchApp.fetch(
            'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + d.id + '?format=metadata',
            { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
          ).getContentText());
          if (detail.message && detail.message.threadId === threadId) {
            UrlFetchApp.fetch(
              'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + d.id,
              { method: 'DELETE', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
            );
            count++;
          }
        } catch(e) {}
      });
    } catch(e) { Logger.log('deleteDraftsForThread error: ' + e); }
    return count;
  }

  function findThread(mpn) {
    var clean = mpn.replace(/[^A-Za-z0-9\-]/g, ' ').trim();
    var threads = GmailApp.search('in:inbox subject:' + clean, 0, 5);
    if (!threads.length) threads = GmailApp.search('in:inbox ' + clean, 0, 5);
    return threads.length ? threads[0] : null;
  }

  // Parts that need need_tp_500 (delete wrong draft, create correct one)
  ['MT25QL256ABA8ESF-0SIT', '5962-8687503XA', '116119AN1699', 'MT47H128M16RT-25E:CTR'].forEach(function(mpn) {
    try {
      var thread = findThread(mpn);
      if (!thread) { Logger.log('Thread not found: ' + mpn); return; }
      var deleted = deleteDraftsForThread(thread.getId());
      var lastMsg = thread.getMessages()[thread.getMessages().length - 1];
      var html    = '<div dir="ltr">' + MSG_NEED_TP_500.replace(/\n/g, '<br>') + sig + '</div>';
      lastMsg.createDraftReply(MSG_NEED_TP_500, { htmlBody: html });
      Logger.log('Fixed need_tp_500 for: ' + mpn + ' (deleted ' + deleted + ' wrong draft(s))');
    } catch(e) { Logger.log('fixWrongOemDrafts error for ' + mpn + ': ' + e); }
  });

  // NRF52832-QFAA-R — BILL EXT, alice gave TP $0.80 → create MSG_BILL
  try {
    var nrfThread = findThread('NRF52832-QFAA-R');
    if (!nrfThread) {
      Logger.log('Thread not found: NRF52832-QFAA-R');
    } else {
      var deleted2 = deleteDraftsForThread(nrfThread.getId());
      var lastMsg2 = nrfThread.getMessages()[nrfThread.getMessages().length - 1];
      var billHtml = '<div dir="ltr">' + MSG_BILL_TEXT + sig + '</div>';
      lastMsg2.createDraftReply(MSG_BILL_TEXT, { htmlBody: billHtml, to: lastMsg2.getTo(), cc: BILL_EMAIL });
      Logger.log('Created MSG_BILL for NRF52832-QFAA-R (deleted ' + deleted2 + ' draft(s))');
    }
  } catch(e) { Logger.log('fixWrongOemDrafts error for NRF52832-QFAA-R: ' + e); }

  Logger.log('fixWrongOemDrafts: done');
}


// ============================================================
// FUNCTION 4: fixAdviceBlockDrafts() — REQUIRES MAIN SCRIPT
// Removes old [ADVICE:] yellow boxes from:
//   K3KL8L80CM-MGCT       → clean need_tp_500
//   MT25QL256ABA8ESF-0SIT → clean need_tp_500
// ============================================================

function fixAdviceBlockDrafts() {
  var MSG_NEED_TP_500 = 'We need a target price to proceed. Please note there is a $500 minimum line requirement. Once we have your target we will get back to you right away.';
  var token = ScriptApp.getOAuthToken();

  var sig = (typeof getSignatureHTML === 'function') ? getSignatureHTML() :
    '<br><br>Regards,<br>John Fluman<br>Intransit Technologies<br>'
    + 'john.fluman@intransittech.com<br>An ISO 9001 Certified Company<br>'
    + 'Toll (877) 677-5868 x101 - Local (949) 481-7935 x101';

  ['K3KL8L80CM-MGCT', 'MT25QL256ABA8ESF-0SIT'].forEach(function(mpn) {
    try {
      var threads = GmailApp.search('subject:' + mpn.replace(/[^A-Za-z0-9\-]/g, ' '), 0, 5);
      if (!threads.length) { Logger.log('Thread not found: ' + mpn); return; }
      var thread  = threads[0];
      var lastMsg = thread.getMessages()[thread.getMessages().length - 1];

      var deleted = 0;
      var drafts = (JSON.parse(UrlFetchApp.fetch(
        'https://gmail.googleapis.com/gmail/v1/users/me/drafts?maxResults=50',
        { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
      ).getContentText()).drafts) || [];
      drafts.forEach(function(d) {
        try {
          var detail = JSON.parse(UrlFetchApp.fetch(
            'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + d.id + '?format=metadata',
            { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
          ).getContentText());
          if (detail.message && detail.message.threadId === thread.getId()) {
            UrlFetchApp.fetch(
              'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + d.id,
              { method: 'DELETE', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
            );
            deleted++;
          }
        } catch(e) {}
      });

      var html = '<div dir="ltr">' + MSG_NEED_TP_500 + sig + '</div>';
      lastMsg.createDraftReply(MSG_NEED_TP_500, { htmlBody: html });
      Logger.log('Fixed advice block for ' + mpn + ' (deleted ' + deleted + ' old draft(s))');
    } catch(e) { Logger.log('fixAdviceBlockDrafts error for ' + mpn + ': ' + e); }
  });

  Logger.log('fixAdviceBlockDrafts: done');
}


// ============================================================
// FUNCTION 5: inboxFullAudit() — REQUIRES MAIN SCRIPT
// Full audit of every inbox thread + every thread with a draft.
// Fixes: BILL EXT → MSG_BILL, wrong $2,000 → $500, stale drafts deleted,
//        [ADVICE:] blocks removed. Detailed log output at end.
// ============================================================

function inboxFullAudit() {
  var BILL_EMAIL = 'bill.pratt@intransittech.com';
  var LOG = [], FIXED = [], FLAGGED = [];

  function log(level, subj, mpn, msg) {
    var line = '[' + level + '] ' + (subj || '').substring(0, 55) + (mpn ? ' | ' + mpn : '') + '\n     ' + msg;
    LOG.push(line);
    Logger.log(line);
  }

  // Prefetch all agent decisions (one call)
  var decisionMap = {};
  try {
    var dr = UrlFetchApp.fetch(HUB_URL + '/api/agent-decisions?limit=500', {
      headers: { Authorization: 'Bearer ' + HUB_SECRET }, muteHttpExceptions: true
    });
    (JSON.parse(dr.getContentText()).decisions || []).forEach(function(d) {
      if (!decisionMap[d.thread_id]) decisionMap[d.thread_id] = [];
      decisionMap[d.thread_id].push(d);
    });
  } catch(e) { Logger.log('Could not prefetch decisions: ' + e); }

  // Map all current Gmail drafts by thread_id
  var draftByThread = {};
  GmailApp.getDrafts().forEach(function(d) {
    try {
      var tId = d.getMessage().getThread().getId();
      if (!draftByThread[tId]) draftByThread[tId] = [];
      draftByThread[tId].push(d);
    } catch(e) {}
  });

  // Collect threads: inbox (last 30d) + any thread with a draft
  var threadMap = {};
  var BLOCKED_DOMAINS_LOCAL = (typeof getBlockedDomains === 'function') ? getBlockedDomains() : [];

  GmailApp.search('in:inbox newer_than:30d -from:intransittech.com', 0, 60).forEach(function(t) {
    threadMap[t.getId()] = t;
  });
  Object.keys(draftByThread).forEach(function(tId) {
    if (!threadMap[tId]) {
      try { var t = GmailApp.getThreadById(tId); if (t) threadMap[tId] = t; } catch(e) {}
    }
  });

  var allThreadIds = Object.keys(threadMap);
  Logger.log('=== INBOX AUDIT: ' + allThreadIds.length + ' threads ===\n');

  // Audit each thread
  allThreadIds.forEach(function(tId) {
    var thread = threadMap[tId];
    try {
      var subject  = thread.getFirstMessageSubject();
      var messages = thread.getMessages();
      var firstMsg = messages[0];
      var lastMsg  = messages[messages.length - 1];
      var sender   = firstMsg.getFrom();

      if (sender.toLowerCase().indexOf('intransittech.com') >= 0) return;
      if (BLOCKED_DOMAINS_LOCAL.some(function(d) { return sender.toLowerCase().indexOf(d) >= 0; })) {
        log('SKIP', subject, null, 'Blocked domain'); return;
      }

      var mpn       = (typeof extractMPN === 'function') ? (extractMPN(subject) || null) : null;
      var decisions = decisionMap[tId] || [];
      var latest    = decisions.length ? decisions[0] : null;
      var drafts    = draftByThread[tId] || [];

      // OEM EXCESS check
      var oemResults = [], inOEM = false, hasBillExt = false, has2k = false;
      if (mpn) {
        try {
          var webUrl = 'https://script.google.com/macros/s/AKfycbyuuBmiYVW5mKI82D5YQGPh1nNGLJZzlLKoxuOdtmOUwUe75VlhhakqgwKooZu5LHFK/exec'
            + '?key=baSDJ%23444FE%268&mpn=' + encodeURIComponent(mpn);
          var wr = UrlFetchApp.fetch(webUrl, { followRedirects: true, muteHttpExceptions: true });
          oemResults  = (JSON.parse(wr.getContentText()).oem_excess) || [];
          inOEM       = oemResults.length > 0;
          hasBillExt  = inOEM && oemResults.some(function(r) { return r.notes.indexOf('BILL EXT') >= 0; });
          has2k       = inOEM && oemResults.some(function(r) { return r.notes.indexOf('$2000') >= 0 || r.notes.indexOf('$2,000') >= 0; });
        } catch(we) { Logger.log('OEM check error (' + mpn + '): ' + we); }
      }

      // No drafts path
      if (!drafts.length) {
        var act = latest ? latest.action : '(none)';
        if (act === 'no_bid' || act === 'no_action' || act === 'forward_deb') {
          var warn = (act === 'no_bid' && inOEM) ? ' ⚠️ STILL IN OEM EXCESS (qty=' + oemResults[0].qty + ')' : '';
          log('OK', subject, mpn, 'Agent: ' + act + warn);
        } else if (inOEM) {
          log('INFO', subject, mpn, 'In OEM EXCESS — agent: ' + act + '. No draft (possibly sent or pending).');
        } else {
          log('OK', subject, mpn || '(no MPN)', 'Not in OEM. Agent: ' + act + '. No draft. Correct.');
        }
        return;
      }

      // Has drafts — check each one
      drafts.forEach(function(draft) {
        try {
          var draftMsg   = draft.getMessage();
          var draftHtml  = draftMsg.getBody();
          var draftPlain = draftMsg.getPlainBody();
          var draftCc    = draftMsg.getCc() || '';
          var draftTo    = draftMsg.getTo();

          var isMsgChecking = draftPlain.indexOf('checking on it now') >= 0;
          var isNeedTp500   = draftPlain.indexOf('$500 minimum') >= 0 && draftPlain.indexOf('$2,000 minimum') < 0;
          var isNeedTp2000  = draftPlain.indexOf('$2,000 minimum') >= 0;
          var isMsgBill     = draftPlain.indexOf('Bill will help') >= 0;
          var hasAdvice     = draftPlain.indexOf('[ADVICE:') >= 0 || draftHtml.indexOf('Note for John (remove') >= 0;
          var isBillCCed    = draftCc.indexOf('bill.pratt') >= 0;

          var draftType = isMsgChecking ? 'MSG_CHECKING'
                        : isNeedTp500   ? 'NEED_TP_500'
                        : isNeedTp2000  ? 'NEED_TP_2000'
                        : isMsgBill     ? 'MSG_BILL'
                        : '(custom/unknown)';

          // Determine buyer email
          var buyerEmail = (typeof extractBuyerEmail === 'function') ? extractBuyerEmail(sender) : sender.match(/<([^>]+)>/)?.[1] || sender;
          if (typeof isNetcompEmail === 'function' && isNetcompEmail(sender, subject)) {
            buyerEmail = (typeof extractNetcompsBuyerEmail === 'function' ? extractNetcompsBuyerEmail(lastMsg.getPlainBody()) : null) || buyerEmail;
          } else if (typeof isICSouceEmail === 'function' && isICSouceEmail(sender, subject)) {
            buyerEmail = (typeof extractICSourcBuyerEmail === 'function' ? extractICSourcBuyerEmail(lastMsg.getBody()) : null) || buyerEmail;
          }

          var issues = [];
          if (hasAdvice)                       issues.push('[ADVICE:] block present');
          if (isNeedTp2000 && !has2k && inOEM) issues.push('WRONG TP MIN — notes say $500 not $2,000');
          if (hasBillExt && isMsgChecking)      issues.push('BILL EXT WRONG — should be MSG_BILL to Bill');
          if (!inOEM && mpn && (isMsgChecking || isNeedTp500 || isNeedTp2000)) {
            issues.push('STALE — ' + mpn + ' no longer in OEM EXCESS');
          }

          if (!issues.length) {
            log('OK', subject, mpn, 'Draft [' + draftType + '] to: ' + draftTo + ' — correct.');
            return;
          }

          log('FIX', subject, mpn, 'Draft [' + draftType + '] issues: ' + issues.join(' | '));

          var sig = (typeof getSignatureHTML === 'function') ? getSignatureHTML() :
            '<br><br>Regards,<br>John Fluman<br>Intransit Technologies<br>'
            + 'john.fluman@intransittech.com<br>An ISO 9001 Certified Company<br>'
            + 'Toll (877) 677-5868 x101 - Local (949) 481-7935 x101';

          var MSG_CHECK   = (typeof MSG_CHECKING    === 'string') ? MSG_CHECKING    : 'We are checking on it now and will get back to you right away.';
          var MSG_TP500   = (typeof MSG_NEED_TP_500 === 'string') ? MSG_NEED_TP_500 : 'We need a target price to proceed. Please note there is a $500 minimum line requirement. Once we have your target we will get back to you right away.';
          var MSG_TP2000  = (typeof MSG_NEED_TP_2000 === 'string') ? MSG_NEED_TP_2000 : 'We need a target price to proceed. Please note there is a $2,000 minimum line requirement. Once we have your target we will get back to you right away.';
          var MSG_BILL_TXT = (typeof MSG_BILL === 'string') ? MSG_BILL : 'Bill will help with this request';

          var fixNote;

          // FIX: BILL EXT + MSG_CHECKING → MSG_BILL
          if (hasBillExt && isMsgChecking) {
            draft.deleteDraft();
            lastMsg.createDraftReply('', { htmlBody: '<div dir="ltr">' + MSG_BILL_TXT + sig + '</div>', to: buyerEmail, cc: BILL_EMAIL });
            fixNote = 'BILL EXT FIX: replaced MSG_CHECKING with MSG_BILL (CC: bill.pratt)';
            log('FIXED', subject, mpn, fixNote);
            FIXED.push({ subject: subject, mpn: mpn, fix: fixNote });
            return;
          }

          // FIX: Wrong TP min ($2000 → $500)
          if (isNeedTp2000 && !has2k && inOEM) {
            draft.deleteDraft();
            lastMsg.createDraftReply('', { htmlBody: '<div dir="ltr">' + MSG_TP500 + sig + '</div>', to: buyerEmail });
            fixNote = 'TP FIX: replaced $2,000 message with $500 message — OEM notes say $500 min';
            log('FIXED', subject, mpn, fixNote);
            FIXED.push({ subject: subject, mpn: mpn, fix: fixNote });
            return;
          }

          // FIX: Stale draft — part no longer in OEM EXCESS
          if (!inOEM && mpn && (isMsgChecking || isNeedTp500 || isNeedTp2000)) {
            draft.deleteDraft();
            fixNote = 'STALE FIX: deleted draft — ' + mpn + ' no longer in OEM EXCESS';
            log('FIXED', subject, mpn, fixNote);
            FIXED.push({ subject: subject, mpn: mpn, fix: fixNote });
            return;
          }

          // FIX: [ADVICE:] block — recreate clean draft
          if (hasAdvice) {
            var cleanMsg = isMsgChecking ? MSG_CHECK
                         : isNeedTp2000  ? MSG_TP2000
                         : isMsgBill     ? MSG_BILL_TXT
                         : MSG_TP500;
            draft.deleteDraft();
            var opts = { htmlBody: '<div dir="ltr">' + cleanMsg + sig + '</div>', to: buyerEmail };
            if (isMsgBill || isBillCCed) opts.cc = BILL_EMAIL;
            lastMsg.createDraftReply('', opts);
            fixNote = 'ADVICE FIX: removed [ADVICE:] block, recreated clean ' + draftType;
            log('FIXED', subject, mpn, fixNote);
            FIXED.push({ subject: subject, mpn: mpn, fix: fixNote });
          }

        } catch(draftErr) {
          log('ERROR', subject, mpn, 'Draft fix error: ' + draftErr.toString());
        }
      });

    } catch(threadErr) {
      Logger.log('Thread error: ' + threadErr.toString());
    }
  });

  // Final report
  Logger.log('\n========================================');
  Logger.log('INBOX AUDIT COMPLETE');
  Logger.log('Threads audited: ' + allThreadIds.length);
  Logger.log('Fixes applied  : ' + FIXED.length);
  Logger.log('========================================');
  if (FIXED.length) {
    Logger.log('\n--- FIXES APPLIED ---');
    FIXED.forEach(function(f) { Logger.log('  [' + (f.mpn || 'n/a') + '] ' + f.subject.substring(0, 50) + '\n    ' + f.fix); });
  }
  Logger.log('\n--- FULL LOG ---');
  LOG.forEach(function(l) { Logger.log(l); });
}


// ============================================================
// addXGL4020ToForte() — 2026-06-23
// Ronda Strand (Abstract Electronics) gave TP $0.67 for XGL4020-472MEC,
// 774 pcs. Script created wrong "need TP" draft (she already gave TP).
// Manual fix: delete the wrong draft, send the MSG_CHECKING draft.
// Then run this to add the Forte entry.
// ============================================================
function addXGL4020ToForte() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var today = '6/23/2026';
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow([today, 'XGL4020-472MEC', 774, 0.67, '', 'US',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added: XGL4020-472MEC | 774 pcs | $0.67 TP');
}

// ============================================================
// addSD1446ToStan() — 2026-06-23
// Rafael Pacas (1-Source) RFQ for SD1446, TP $90, "as many as you have"
// Warehouse#3 IN STOCK: 25+38+4 = 67 pcs. Routes to Stan sheet.
// ============================================================
function addSD1446ToStan() {
  var STAN_SHEET_ID = '1pGRDpkqftQNoEYna53MxRJfUY8jEf5_w32FNa56OUIM';
  var sheet = SpreadsheetApp.openById(STAN_SHEET_ID).getSheets()[0];
  sheet.appendRow(['', '', '', '6/23/2026', 'SD1446', 'US', 67, 90]);
  Logger.log('Stan sheet: SD1446 | US | 67 pcs | TP $90');
}

// ============================================================
// extractCBPFormText() — 2026-06-23
// Reads the CBP Form 7501 PDF (entry 1FX55523609) directly from
// Gmail message 19eef4ad45a1ad98, extracts readable ASCII text
// from the PDF bytes, and logs HTS codes, duty amounts, and
// all key field values. No Drive API needed.
// ============================================================
function extractCBPFormText() {
  var message = GmailApp.getMessageById('19eef4ad45a1ad98');
  if (!message) { Logger.log('ERROR: message not found'); return; }

  var attachments = message.getAttachments();
  var found = false;

  attachments.forEach(function(att) {
    if (!(att.getContentType() === 'application/pdf' || att.getName().match(/\.pdf$/i))) return;
    found = true;
    Logger.log('Reading: ' + att.getName() + ' (' + att.getSize() + ' bytes)');

    // Convert raw bytes to ASCII string (readable PDF text is stored as plain ASCII)
    var bytes = att.getBytes();
    var raw = '';
    for (var i = 0; i < bytes.length; i++) {
      var b = bytes[i] & 0xFF;
      raw += (b >= 32 && b < 127) ? String.fromCharCode(b) : ' ';
    }

    // Collapse whitespace runs for cleaner output
    var text = raw.replace(/ {3,}/g, '   ');

    // --- Extract HTS codes (format: NNNN.NN.NNNN or NNNN.NN.NN) ---
    var htsCodes = text.match(/\b\d{4}\.\d{2}\.\d{2,4}\b/g) || [];
    Logger.log('HTS CODES FOUND: ' + (htsCodes.length ? htsCodes.join(', ') : 'none'));

    // --- Extract dollar amounts ---
    var dollars = text.match(/\$\s*[\d,]+\.?\d*/g) || [];
    Logger.log('DOLLAR AMOUNTS: ' + (dollars.length ? dollars.join(', ') : 'none'));

    // --- Log surrounding context for key CBP fields ---
    var keywords = ['HTS', 'DUTY', 'RATE', 'ENTERED VALUE', 'DUTIABLE', 'COUNTRY OF ORIGIN',
                    'TARIFF', 'ASSESS', 'TOTAL', 'VALUE', 'DESCRIPTION', 'UNIT'];
    keywords.forEach(function(kw) {
      var idx = text.indexOf(kw);
      var count = 0;
      while (idx >= 0 && count < 3) {
        Logger.log('[' + kw + '] ...' + text.substring(Math.max(0, idx - 15), Math.min(text.length, idx + 120)).trim() + '...');
        idx = text.indexOf(kw, idx + kw.length + 1);
        count++;
      }
    });

    // --- Also dump first 3000 chars of readable text for full context ---
    Logger.log('\n=== RAW TEXT (first 3000 chars) ===\n' + text.substring(0, 3000));
    Logger.log('\n=== RAW TEXT (chars 3000-6000) ===\n' + text.substring(3000, 6000));
    Logger.log('\n=== RAW TEXT (chars 6000-9000) ===\n' + text.substring(6000, 9000));
  });

  if (!found) Logger.log('No PDF attachments found on message 19eef4ad45a1ad98');
}
