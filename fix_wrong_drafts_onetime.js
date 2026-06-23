// ============================================================
// ONE-TIME FIX — Delete wrong "need TP" drafts and replace with
// correct own-stock quotes for BCM5461SA1KPFG and BCM5221A4KPTG
// Run once from Apps Script editor, then delete this file.
// ============================================================

function fixWrongNetcompDrafts() {
  var fixes = [
    {
      subjectFragment: 'BCM5461SA1KPFG',
      mpn:    'BCM5461SA1KPFG',
      dc:     '2010+',
      qty:    94,
      man:    'BROADCOM',
      notes:  'Also have OEM EXCESS (QTY 24,665) and Warehouse#3 stock available if needed.'
    },
    {
      subjectFragment: 'BCM5221A4KPTG',
      mpn:    'BCM5221A4KPTG',
      dc:     '2130',
      qty:    115,
      man:    'BROADCOM',
      notes:  'Also have OEM EXCESS and Warehouse#3 stock available if needed.'
    }
  ];

  var JOHN_EMAIL = 'john.fluman@intransittech.com';
  var sig = '<br><br>Regards,<br>John Fluman<br>Intransit Technologies<br>'
    + 'john.fluman@intransittech.com<br>An ISO 9001 Certified Company<br>'
    + 'Toll (877) 677-5868 x101 - Local (949) 481-7935 x101';

  fixes.forEach(function(fix) {
    try {
      // ── Find the thread ───────────────────────────────────
      var threads = GmailApp.search('in:inbox subject:' + fix.mpn, 0, 5);
      if (!threads.length) {
        Logger.log('Thread not found for: ' + fix.mpn);
        return;
      }
      var thread = threads[0];
      var threadId = thread.getId();
      var messages = thread.getMessages();
      var lastMsg  = messages[messages.length - 1];
      var subject  = thread.getFirstMessageSubject();

      // ── Delete existing wrong drafts for this thread ──────
      var token = ScriptApp.getOAuthToken();
      var draftsResp = UrlFetchApp.fetch(
        'https://gmail.googleapis.com/gmail/v1/users/me/drafts?q=in:draft',
        { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
      );
      var draftsData = JSON.parse(draftsResp.getContentText());
      var allDrafts  = draftsData.drafts || [];

      var deletedCount = 0;
      allDrafts.forEach(function(d) {
        try {
          var draftDetail = JSON.parse(UrlFetchApp.fetch(
            'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + d.id + '?format=metadata',
            { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
          ).getContentText());
          if (draftDetail.message && draftDetail.message.threadId === threadId) {
            UrlFetchApp.fetch(
              'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + d.id,
              { method: 'DELETE', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
            );
            deletedCount++;
            Logger.log('Deleted wrong draft for: ' + fix.mpn);
          }
        } catch(e) { Logger.log('Draft delete error: ' + e); }
      });

      // ── Build correct own-stock draft ────────────────────
      var bodyText = 'This is our stock\n\n'
        + 'MPN: ' + fix.mpn + '\n'
        + 'Manufacturer: ' + fix.man + '\n'
        + 'DC: ' + fix.dc + '\n'
        + 'QTY Available: ' + fix.qty + '\n'
        + 'Price: $[FILL IN]\n\n'
        + 'There is a $100 minimum on stock items.';

      var htmlBody = '<div dir="ltr">'
        + bodyText.replace(/\n/g, '<br>')
        + sig + '</div>';

      var newDraft = lastMsg.createDraftReply(bodyText, { htmlBody: htmlBody });
      Logger.log('Created correct own-stock draft for: ' + fix.mpn + ' (deleted ' + deletedCount + ' wrong draft(s))');

    } catch(e) {
      Logger.log('fixWrongNetcompDrafts error for ' + fix.mpn + ': ' + e);
    }
  });

  Logger.log('fixWrongNetcompDrafts: done');
}

// ============================================================
// Fix 5 wrong OEM drafts found in inbox:
//   MT25QL256ABA8ESF-0SIT  — wrong $2000 draft → correct need_tp_500
//   5962-8687503XA          — wrong MSG_BILL (no TP) → correct need_tp_500
//   116119AN1699            — wrong MSG_CHECKING (no TP) → correct need_tp_500
//   MT47H128M16RT-25E:CTR  — wrong MSG_CHECKING (no TP) → correct need_tp_500
//   NRF52832-QFAA-R        — no draft despite alice giving TP $0.80 → create MSG_BILL
// ============================================================

function fixWrongOemDrafts() {
  var BILL_EMAIL = 'bill.pratt@intransittech.com';
  var MSG_NEED_TP_500 = 'We need a target price to proceed. Please note there is a $500 minimum line requirement. Once we have your target we will get back to you right away.';
  var MSG_BILL = 'Bill will help with this request';
  var token = ScriptApp.getOAuthToken();

  var sig = getSignatureHTML ? getSignatureHTML() :
    '<br><br>Regards,<br>John Fluman<br>Intransit Technologies<br>'
    + 'john.fluman@intransittech.com<br>An ISO 9001 Certified Company<br>'
    + 'Toll (877) 677-5868 x101 - Local (949) 481-7935 x101';

  function deleteDraftsForThread(threadId) {
    var count = 0;
    try {
      var resp = UrlFetchApp.fetch(
        'https://gmail.googleapis.com/gmail/v1/users/me/drafts?maxResults=50',
        { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
      );
      var drafts = (JSON.parse(resp.getContentText()).drafts) || [];
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
    // Try exact subject search first; fall back to body search for MPNs with special chars
    var sanitized = mpn.replace(/[^A-Za-z0-9\-]/g, ' ').trim();
    var threads = GmailApp.search('in:inbox subject:' + sanitized, 0, 5);
    if (!threads.length) threads = GmailApp.search('in:inbox ' + sanitized, 0, 5);
    return threads.length ? threads[0] : null;
  }

  // ── Parts that need a need_tp_500 draft (delete wrong draft, create correct one) ──
  var needTpFixes = [
    { mpn: 'MT25QL256ABA8ESF-0SIT' },
    { mpn: '5962-8687503XA' },
    { mpn: '116119AN1699' },
    { mpn: 'MT47H128M16RT-25E:CTR' }
  ];

  needTpFixes.forEach(function(fix) {
    try {
      var thread = findThread(fix.mpn);
      if (!thread) { Logger.log('Thread not found: ' + fix.mpn); return; }
      var deleted = deleteDraftsForThread(thread.getId());
      var lastMsg = thread.getMessages()[thread.getMessages().length - 1];
      var htmlBody = '<div dir="ltr">' + MSG_NEED_TP_500.replace(/\n/g, '<br>') + sig + '</div>';
      lastMsg.createDraftReply(MSG_NEED_TP_500, { htmlBody: htmlBody });
      Logger.log('Fixed need_tp_500 draft for: ' + fix.mpn + ' (deleted ' + deleted + ' wrong draft(s))');
    } catch(e) {
      Logger.log('fixWrongOemDrafts error for ' + fix.mpn + ': ' + e);
    }
  });

  // ── NRF52832-QFAA-R — alice gave TP $0.80, BILL EXT part → create MSG_BILL ──
  try {
    var nrfThread = findThread('NRF52832-QFAA-R');
    if (!nrfThread) {
      Logger.log('Thread not found: NRF52832-QFAA-R');
    } else {
      var nrfMessages = nrfThread.getMessages();
      var nrfLastMsg  = nrfMessages[nrfMessages.length - 1];
      // Check if MSG_BILL draft already exists
      var existingDeleted = deleteDraftsForThread(nrfThread.getId()); // clear any bad drafts
      var htmlBody = '<div dir="ltr">' + MSG_BILL + sig + '</div>';
      nrfLastMsg.createDraftReply(MSG_BILL, { htmlBody: htmlBody, to: BILL_EMAIL });
      Logger.log('Created MSG_BILL draft for NRF52832-QFAA-R (deleted ' + existingDeleted + ' existing draft(s))');
    }
  } catch(e) {
    Logger.log('fixWrongOemDrafts error for NRF52832-QFAA-R: ' + e);
  }

  Logger.log('fixWrongOemDrafts: done');
}

// ============================================================
// Fix drafts that still have advice blocks or wrong text:
//   K3KL8L80CM-MGCT        — wrong custom text → correct need_tp_500
//   MT25QL256ABA8ESF-0SIT  — correct text but has old advice block → clean version
// Run once from Apps Script editor.
// ============================================================

function fixAdviceBlockDrafts() {
  var MSG_NEED_TP_500 = 'We need a target price to proceed. Please note there is a $500 minimum line requirement. Once we have your target we will get back to you right away.';
  var token = ScriptApp.getOAuthToken();

  var sig = getSignatureHTML ? getSignatureHTML() :
    '<br><br>Regards,<br>John Fluman<br>Intransit Technologies<br>'
    + 'john.fluman@intransittech.com<br>An ISO 9001 Certified Company<br>'
    + 'Toll (877) 677-5868 x101 - Local (949) 481-7935 x101';

  var mpns = ['K3KL8L80CM-MGCT', 'MT25QL256ABA8ESF-0SIT'];

  mpns.forEach(function(mpn) {
    try {
      // Search inbox and drafts — some threads may not have inbox label if already processed
      var threads = GmailApp.search('subject:' + mpn.replace(/[^A-Za-z0-9\-]/g, ' '), 0, 5);
      if (!threads.length) { Logger.log('Thread not found: ' + mpn); return; }
      var thread  = threads[0];
      var threadId = thread.getId();
      var lastMsg = thread.getMessages()[thread.getMessages().length - 1];

      // Delete all existing drafts on this thread
      var deleted = 0;
      try {
        var resp = UrlFetchApp.fetch(
          'https://gmail.googleapis.com/gmail/v1/users/me/drafts?maxResults=50',
          { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
        );
        var drafts = JSON.parse(resp.getContentText()).drafts || [];
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
              deleted++;
            }
          } catch(e) {}
        });
      } catch(e) { Logger.log('Draft delete error for ' + mpn + ': ' + e); }

      // Create clean need_tp_500 draft — no advice block
      var htmlBody = '<div dir="ltr">' + MSG_NEED_TP_500 + sig + '</div>';
      lastMsg.createDraftReply(MSG_NEED_TP_500, { htmlBody: htmlBody });
      Logger.log('Fixed draft for ' + mpn + ' (deleted ' + deleted + ' old draft(s))');

    } catch(e) {
      Logger.log('fixAdviceBlockDrafts error for ' + mpn + ': ' + e);
    }
  });

  Logger.log('fixAdviceBlockDrafts: done');
}
