// ============================================================
// INBOX REVIEW — Run anytime from Apps Script editor
// Scans ALL inbox threads, calls Claude on each, creates drafts
// where appropriate, then sends John a summary email.
// Safe to run repeatedly — skips threads that already have a draft.
// ============================================================

function reviewInboxOneTime() {
  var HUB_URL    = 'https://intransit-hub.intransit-sales.workers.dev';
  var HUB_SECRET = 'InTransit!Hub#2026';
  var JOHN_EMAIL = 'john.fluman@intransittech.com';
  var DAVID_EMAIL = 'david@fortetechno.com';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';

  // ── 1. Fetch blocked domains ──────────────────────────────
  var blockedDomains = ['sourceschip.com', 'bulechip.com', 'feelchips.com'];
  try {
    var bResp = UrlFetchApp.fetch(HUB_URL + '/api/rules?type=blocked_domain', {
      headers: { Authorization: 'Bearer ' + HUB_SECRET }, muteHttpExceptions: true
    });
    var bData = JSON.parse(bResp.getContentText());
    if ((bData.rules || []).length) blockedDomains = bData.rules.map(function(r) { return r.key; });
  } catch(e) { Logger.log('Blocked domain fetch error: ' + e); }

  // ── 2. Scan inbox — all non-internal threads ──────────────
  var query = 'in:inbox -from:intransittech.com -from:' + DAVID_EMAIL;
  var threads = GmailApp.search(query, 0, 50);
  Logger.log('reviewInboxOneTime: found ' + threads.length + ' threads');

  var summary = [];   // rows for the report email
  var draftCount = 0;
  var skipCount = 0;

  threads.forEach(function(thread) {
    try {
      var result = _reviewThread(thread, blockedDomains, HUB_URL, HUB_SECRET, FORTE_SHEET_ID);
      summary.push(result);
      if (result.drafted) draftCount++;
      else skipCount++;
    } catch(e) {
      summary.push({
        subject: thread.getFirstMessageSubject(),
        from: '',
        action: 'ERROR',
        drafted: false,
        note: e.toString()
      });
    }
    Utilities.sleep(1200); // pace Claude calls
  });

  // ── 3. Build summary email ────────────────────────────────
  var html = '<div style="font-family:Arial,sans-serif;font-size:11pt;">'
    + '<h2 style="color:#1a5276;">📬 Inbox Review — ' + new Date().toDateString() + '</h2>'
    + '<p><b>' + threads.length + ' threads reviewed</b> · '
    + '<b style="color:#1a7340;">' + draftCount + ' drafts created</b> · '
    + '<b>' + skipCount + ' no action needed</b></p>'
    + '<hr style="border:1px solid #ccc;">';

  summary.forEach(function(r, i) {
    var color = r.drafted ? '#e8f5e9' : (r.action === 'ERROR' ? '#fdecea' : '#f5f5f5');
    var badge = r.drafted ? '📝 DRAFT CREATED' : (r.action === 'no_bid' ? '🚫 NO BID' : (r.action === 'ERROR' ? '⛔ ERROR' : '⏭ SKIP'));
    html += '<div style="background:' + color + ';border:1px solid #ddd;border-radius:4px;padding:10px 14px;margin:8px 0;">'
      + '<b>' + (i + 1) + '. ' + _escHtml(r.subject) + '</b>'
      + ' <span style="color:#666;font-size:10pt;">— ' + _escHtml(r.from) + '</span>'
      + '<br><span style="font-size:10pt;font-weight:bold;">' + badge + '</span>'
      + (r.note ? ' — <span style="font-size:10pt;color:#555;">' + _escHtml(r.note) + '</span>' : '')
      + '</div>';
  });

  html += '</div>';

  GmailApp.sendEmail(JOHN_EMAIL, '📬 Inbox Review — ' + threads.length + ' threads (' + draftCount + ' drafts)', '', { htmlBody: html });
  Logger.log('reviewInboxOneTime: done. Drafts: ' + draftCount + ', Skipped: ' + skipCount);
}

// ── Process one thread ──────────────────────────────────────
function _reviewThread(thread, blockedDomains, HUB_URL, HUB_SECRET, FORTE_SHEET_ID) {
  var messages  = thread.getMessages();
  var firstMsg  = messages[0];
  var lastMsg   = messages[messages.length - 1];
  var subject   = thread.getFirstMessageSubject();
  var threadId  = thread.getId();
  var sender    = firstMsg.getFrom();

  var result = { subject: subject, from: sender, action: '', drafted: false, note: '' };

  // Skip blocked domains
  for (var b = 0; b < blockedDomains.length; b++) {
    if (sender.toLowerCase().indexOf(blockedDomains[b]) >= 0) {
      result.action = 'blocked'; result.note = 'Blocked domain: ' + blockedDomains[b];
      return result;
    }
  }

  // Skip internal
  if (sender.toLowerCase().indexOf('intransittech.com') >= 0) {
    result.action = 'internal'; result.note = 'Internal email — skipped';
    return result;
  }

  // Build thread text
  var threadText = messages.map(function(m, i) {
    var body = (m.getPlainBody() || '').split('\n').filter(function(l) {
      return l.trim() && !l.match(/^(>|On .* wrote:|--)/);
    }).slice(0, 60).join('\n');
    return 'MSG ' + (i + 1) + ' | From: ' + m.getFrom() + '\n' + body.substring(0, 800);
  }).join('\n\n---\n\n').substring(0, 6000);

  // Extract MPN
  var mpn = _extractMPN(subject + ' ' + (lastMsg.getPlainBody() || '').substring(0, 500));

  // OEM + Forte lookup
  var oemResults = [], forteResults = [];
  if (mpn) {
    try {
      var webUrl = 'https://script.google.com/macros/s/AKfycbyuuBmiYVW5mKI82D5YQGPh1nNGLJZzlLKoxuOdtmOUwUe75VlhhakqgwKooZu5LHFK/exec'
        + '?key=baSDJ%23444FE%268&mpn=' + encodeURIComponent(mpn);
      var webResp = UrlFetchApp.fetch(webUrl, { followRedirects: true, muteHttpExceptions: true });
      var webData = JSON.parse(webResp.getContentText());
      oemResults   = webData.oem_excess  || [];
      forteResults = webData.forte_sheet || [];
    } catch(e) {}
  }

  // Call email agent
  var agentResp = UrlFetchApp.fetch(HUB_URL + '/api/email-agent', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + HUB_SECRET },
    payload: JSON.stringify({
      thread_id:       threadId,
      last_message_id: lastMsg.getId(),
      subject:         subject,
      sender:          sender,
      thread_content:  threadText,
      oem_results:     oemResults,
      forte_results:   forteResults,
      current_labels:  thread.getLabels().map(function(l) { return l.getName(); })
    }),
    muteHttpExceptions: true
  });

  if (agentResp.getResponseCode() !== 200) {
    result.action = 'api_error';
    result.note   = 'Worker error ' + agentResp.getResponseCode();
    return result;
  }

  var decision = JSON.parse(agentResp.getContentText());
  var action   = decision.action || 'no_action';
  result.action = action;
  result.note   = decision.reasoning ? decision.reasoning.substring(0, 120) : '';

  // No draft needed
  if (action === 'no_action' || action === 'no_bid') return result;

  // Skip if a draft already exists for this thread (avoid duplicates on repeat runs)
  try {
    var existingDrafts = Gmail.Users.Drafts.list('me', { q: 'in:draft' }).drafts || [];
    var hasDraft = existingDrafts.some(function(d) {
      try {
        var msg = Gmail.Users.Drafts.get('me', d.id, { format: 'metadata', metadataHeaders: ['Subject'] });
        return msg.message && msg.message.threadId === threadId;
      } catch(e) { return false; }
    });
    if (hasDraft) {
      result.action = 'draft_exists';
      result.note   = 'Draft already pending — skipped to avoid duplicate';
      return result;
    }
  } catch(e) { /* Gmail Advanced Service not enabled — skip this check */ }

  // Create draft
  if (decision.draft_body) {
    var plainBody = decision.draft_body;
    var adviceMatch = plainBody.match(/\[ADVICE:([\s\S]*?)\]\s*$/);
    var bodyText  = adviceMatch ? plainBody.replace(/\s*\[ADVICE:[\s\S]*?\]\s*$/, '').trim() : plainBody.trim();
    var agentAdvice = adviceMatch ? adviceMatch[1].trim() : (decision.reasoning || 'AI review draft. Verify before sending.');

    // Lock standard texts
    if (action === 'msg_checking')     bodyText = 'We are checking on it now. If we get a response from the OEM, I will respond to you right away. If we do not respond back to you, please consider this a no bid. Thank you very much for the opportunity.';
    if (action === 'request_tp_500')   bodyText = 'We need a target price to proceed. Please note there is a $500 minimum line requirement. Once we have your target we will get back to you right away.';
    if (action === 'request_tp_2000')  bodyText = 'We need a target price to proceed. Please note there is a $2,000 minimum line requirement. Once we have your target we will get back to you right away.';
    if (action === 'bill_handle')      bodyText = 'Bill will help with this request';

    var sig = '<br><br>Regards,<br>John Fluman<br>Intransit Technologies<br>john.fluman@intransittech.com<br>An ISO 9001 Certified Company<br>Toll (877) 677-5868 x101 - Local (949) 481-7935 x101';
    var htmlBody = '<div dir="ltr">' + bodyText.replace(/\n/g, '<br>') + sig + '</div>';

    var draft = lastMsg.createDraftReply(bodyText, { htmlBody: htmlBody });
    if (draft) {
      result.drafted = true;
      result.note    = agentAdvice.substring(0, 120);

      // Forte entry if agent recommends it
      var fe = decision.forte_entry;
      if (fe && fe.mpn && fe.qty && fe.target_price) {
        try {
          var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
          var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
          var nextRow = sheet.getLastRow() + 1;
          sheet.appendRow([today, fe.mpn, fe.qty, fe.target_price, '', fe.country || '', '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
          result.note += ' | Forte: ' + fe.mpn;
        } catch(e) { Logger.log('Forte error: ' + e); }
      }
    }
  }

  return result;
}

// ── Helpers ──────────────────────────────────────────────────
function _extractMPN(text) {
  var m = text.match(/\b([A-Z]{1,4}[\-]?[0-9]{3,}[A-Z0-9\-]*|[0-9]{3,}[A-Z]{2,}[A-Z0-9\-]*)\b/);
  return m ? m[1] : null;
}

function _escHtml(str) {
  return String(str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
