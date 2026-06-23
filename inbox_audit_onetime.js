// ============================================================
// ONE-TIME INBOX AUDIT + FIX — inboxFullAudit()
// Scans every inbox thread + every thread with a draft.
// Explains what the agent did. Fixes: BILL EXT routing, wrong TP amount,
// advice blocks, stale drafts. Run once from Apps Script editor.
//
// Also run fixMissedDavidRemovals() separately to delete the 3 parts
// David said were no-stock but the agent failed to remove.
// ============================================================

// ONE-TIME: Delete parts David confirmed as no-stock that the agent missed
function fixMissedDavidRemovals() {
  var missed = [
    { mpn: 'MCIMX535DVP1C2', subject: '#3900 MCIMX535DVP1C2 No stock' },
    { mpn: 'BD429B',          subject: '#3904 BD429B Cant find'         },
    { mpn: '10M08SAM153I7G',  subject: '#3896 10M08SAM153I7G NO STOCK'  },
  ];
  missed.forEach(function(item) {
    try {
      var result = deletePart(item.mpn, item.subject);
      Logger.log('Removed ' + item.mpn + ' from OEM EXCESS: ' + JSON.stringify(result));
      hubLog('run', 'fixMissedDavidRemovals: removed ' + item.mpn, { mpn: item.mpn });
    } catch(e) {
      Logger.log('ERROR removing ' + item.mpn + ': ' + e.toString());
    }
  });
  Logger.log('fixMissedDavidRemovals complete — 3 parts removed.');
}

function inboxFullAudit() {
  var BILL_EMAIL_LOCAL  = 'bill.pratt@intransittech.com';
  var JOHN_EMAIL_LOCAL  = 'john.fluman@intransittech.com';

  var LOG      = [];  // full detail log
  var FIXED    = [];  // summary of fixes
  var FLAGGED  = [];  // needs John's manual review

  function log(level, subj, mpn, msg) {
    var line = '[' + level + '] ' + (subj || '').substring(0, 55) + (mpn ? ' | ' + mpn : '') + '\n     ' + msg;
    LOG.push(line);
    Logger.log(line);
  }

  // ── 1. Prefetch all agent decisions (one call) ──────────────
  var decisionMap = {};
  try {
    var dr = UrlFetchApp.fetch(HUB_URL + '/api/agent-decisions?limit=500', {
      headers: { Authorization: 'Bearer ' + HUB_SECRET }, muteHttpExceptions: true
    });
    var dd = JSON.parse(dr.getContentText());
    (dd.decisions || []).forEach(function(d) {
      if (!decisionMap[d.thread_id]) decisionMap[d.thread_id] = [];
      decisionMap[d.thread_id].push(d);  // newest first per API default
    });
  } catch(e) { Logger.log('Could not prefetch decisions: ' + e); }

  // ── 2. Map all current Gmail drafts by thread_id ────────────
  var draftByThread = {};
  var allGmailDrafts = GmailApp.getDrafts();
  allGmailDrafts.forEach(function(d) {
    try {
      var tId = d.getMessage().getThread().getId();
      if (!draftByThread[tId]) draftByThread[tId] = [];
      draftByThread[tId].push(d);
    } catch(e) {}
  });

  // ── 3. Collect all threads to audit ─────────────────────────
  var threadMap = {};
  var BLOCKED_DOMAINS_LOCAL = getBlockedDomains();

  // Inbox (last 30 days, skip internal)
  var inboxThreads = GmailApp.search('in:inbox newer_than:30d -from:intransittech.com', 0, 60);
  inboxThreads.forEach(function(t) { threadMap[t.getId()] = t; });

  // Any thread that has a current Gmail draft (even if not in inbox)
  Object.keys(draftByThread).forEach(function(tId) {
    if (!threadMap[tId]) {
      try { var t = GmailApp.getThreadById(tId); if (t) threadMap[tId] = t; } catch(e) {}
    }
  });

  var allThreadIds = Object.keys(threadMap);
  Logger.log('=== INBOX AUDIT: ' + allThreadIds.length + ' threads (' + inboxThreads.length + ' inbox + ' + (allThreadIds.length - inboxThreads.length) + ' draft-only) ===\n');

  // ── 4. Audit each thread ────────────────────────────────────
  allThreadIds.forEach(function(tId) {
    var thread;
    try { thread = threadMap[tId]; } catch(e) { return; }

    try {
      var subject  = thread.getFirstMessageSubject();
      var messages = thread.getMessages();
      var firstMsg = messages[0];
      var lastMsg  = messages[messages.length - 1];
      var sender   = firstMsg.getFrom();

      // Skip internal
      if (sender.toLowerCase().indexOf('intransittech.com') >= 0) return;

      // Skip blocked domains
      var blocked = BLOCKED_DOMAINS_LOCAL.some(function(d) { return sender.toLowerCase().indexOf(d) >= 0; });
      if (blocked) { log('SKIP', subject, null, 'Blocked domain'); return; }

      var mpn       = extractMPN(subject) || extractMPNFromRFQBody(lastMsg.getPlainBody().substring(0, 600));
      var decisions = decisionMap[tId] || [];
      var latest    = decisions.length ? decisions[0] : null;
      var drafts    = draftByThread[tId] || [];

      // ── OEM EXCESS check ──────────────────────────────────
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

      // ── No drafts path ────────────────────────────────────
      if (!drafts.length) {
        var act = latest ? latest.action : '(none)';
        if (act === 'no_bid' || act === 'no_action' || act === 'forward_deb') {
          var inOEMWarning = (act === 'no_bid' && inOEM) ? ' ⚠️ STILL IN OEM EXCESS (qty=' + oemResults[0].qty + ')' : '';
          log('OK', subject, mpn, 'Agent: ' + act + inOEMWarning);
        } else if (inOEM) {
          log('INFO', subject, mpn, 'In OEM EXCESS — agent: ' + act + '. No draft in Gmail (possibly already sent or pending).');
        } else {
          log('OK', subject, mpn || '(no MPN)', 'Not in OEM EXCESS. Agent: ' + act + '. No draft. Correct.');
        }
        return;
      }

      // ── Has drafts — check each one ──────────────────────
      drafts.forEach(function(draft) {
        try {
          var draftMsg     = draft.getMessage();
          var draftHtml    = draftMsg.getBody();
          var draftPlain   = draftMsg.getPlainBody();
          var draftTo      = draftMsg.getTo();
          var draftCc      = draftMsg.getCc() || '';
          var draftSubject = draftMsg.getSubject();

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

          // Determine correct buyer email for any replacement drafts
          var buyerEmail = extractBuyerEmail(firstMsg.getFrom());
          var isNetcomp  = isNetcompEmail(firstMsg.getFrom(), firstMsg.getSubject());
          var isICS      = isICSouceEmail(firstMsg.getFrom(), firstMsg.getSubject());
          if (isNetcomp) {
            buyerEmail = extractNetcompsBuyerEmail(lastMsg.getPlainBody()) || buyerEmail;
          } else if (isICS) {
            buyerEmail = extractICSourcBuyerEmail(lastMsg.getBody()) || buyerEmail;
          }

          var issues = [];
          if (hasAdvice)                         issues.push('[ADVICE:] block present');
          if (isNeedTp2000 && !has2k && inOEM)   issues.push('WRONG TP MIN — notes say $500 not $2,000');
          if (hasBillExt && isMsgChecking)        issues.push('BILL EXT WRONG — should be MSG_BILL to Bill, not MSG_CHECKING');
          if (!inOEM && mpn && (isMsgChecking || isNeedTp500 || isNeedTp2000)) {
            issues.push('STALE — ' + mpn + ' no longer in OEM EXCESS');
          }

          if (!issues.length) {
            log('OK', subject, mpn, 'Draft [' + draftType + '] to: ' + draftTo + ' — correct.');
            return;
          }

          log('FIX', subject, mpn, 'Draft [' + draftType + '] issues: ' + issues.join(' | '));

          var sig = getSignatureHTML();

          // ── FIX: BILL EXT + MSG_CHECKING → MSG_BILL ──────────
          if (hasBillExt && isMsgChecking) {
            draft.deleteDraft();
            var billHtml = '<div dir="ltr">' + MSG_BILL + sig + '</div>';
            lastMsg.createDraftReply('', { htmlBody: billHtml, to: buyerEmail, cc: BILL_EMAIL_LOCAL });
            var fixNote = 'BILL EXT FIX: replaced MSG_CHECKING with MSG_BILL (CC: bill.pratt) — BILL EXT part, buyer had TP';
            log('FIXED', subject, mpn, fixNote);
            FIXED.push({ subject: subject, mpn: mpn, fix: fixNote });
            return;
          }

          // ── FIX: Wrong TP min ($2000 → $500) ─────────────────
          if (isNeedTp2000 && !has2k && inOEM) {
            draft.deleteDraft();
            var tp500Html = '<div dir="ltr">' + MSG_NEED_TP_500 + sig + '</div>';
            lastMsg.createDraftReply('', { htmlBody: tp500Html, to: buyerEmail });
            var fixNote2 = 'TP FIX: replaced $2,000 message with $500 message — OEM notes say $500 min';
            log('FIXED', subject, mpn, fixNote2);
            FIXED.push({ subject: subject, mpn: mpn, fix: fixNote2 });
            return;
          }

          // ── FIX: Stale draft — part no longer in OEM EXCESS ──
          if (!inOEM && mpn && (isMsgChecking || isNeedTp500 || isNeedTp2000)) {
            draft.deleteDraft();
            var fixNote3 = 'STALE FIX: deleted draft — ' + mpn + ' is no longer in OEM EXCESS';
            log('FIXED', subject, mpn, fixNote3);
            FIXED.push({ subject: subject, mpn: mpn, fix: fixNote3 });
            return;
          }

          // ── FIX: [ADVICE:] block — recreate clean draft ───────
          if (hasAdvice) {
            var cleanMsg, cleanHtml;
            if (isMsgChecking) {
              cleanMsg = MSG_CHECKING;
            } else if (isNeedTp2000) {
              cleanMsg = MSG_NEED_TP_2000;
            } else if (isMsgBill) {
              cleanMsg = MSG_BILL;
            } else {
              cleanMsg = MSG_NEED_TP_500;
            }

            draft.deleteDraft();
            cleanHtml = '<div dir="ltr">' + cleanMsg + sig + '</div>';

            if (isMsgBill || isBillCCed) {
              lastMsg.createDraftReply('', { htmlBody: cleanHtml, to: buyerEmail, cc: BILL_EMAIL_LOCAL });
            } else {
              lastMsg.createDraftReply('', { htmlBody: cleanHtml, to: buyerEmail });
            }

            var fixNote4 = 'ADVICE FIX: removed [ADVICE:] block, recreated clean ' + draftType + ' draft';
            log('FIXED', subject, mpn, fixNote4);
            FIXED.push({ subject: subject, mpn: mpn, fix: fixNote4 });
            return;
          }

        } catch(draftErr) {
          log('ERROR', subject, mpn, 'Draft fix error: ' + draftErr.toString());
        }
      });

    } catch(threadErr) {
      Logger.log('Thread error: ' + threadErr.toString());
    }
  });

  // ── 5. Final report ─────────────────────────────────────────
  Logger.log('\n========================================');
  Logger.log('INBOX AUDIT COMPLETE');
  Logger.log('Threads audited : ' + allThreadIds.length);
  Logger.log('Fixes applied   : ' + FIXED.length);
  Logger.log('Flagged manual  : ' + FLAGGED.length);
  Logger.log('========================================');

  if (FIXED.length) {
    Logger.log('\n--- FIXES APPLIED ---');
    FIXED.forEach(function(f) { Logger.log('  • [' + (f.mpn || 'n/a') + '] ' + f.subject.substring(0, 50) + '\n    ' + f.fix); });
  }

  if (FLAGGED.length) {
    Logger.log('\n--- NEEDS MANUAL REVIEW ---');
    FLAGGED.forEach(function(f) { Logger.log('  • ' + f); });
  }

  Logger.log('\n--- FULL LOG ---');
  LOG.forEach(function(l) { Logger.log(l); });
}
