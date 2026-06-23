// ============================================================
// AI EMAIL AGENT — Trigger 7
// Paste this entire block at the BOTTOM of email_script_v24_hub.js
// Then run setupTriggers() once to install the new trigger.
// ============================================================

var AGENT_LABEL = 'oem-agent-processed';

var AGENT_HTML_SIG =
  '<br><br>' +
  '<div><b><span style="color:rgb(31,73,125);font-family:Tahoma,sans-serif;font-size:10pt">Regards,</span></b></div>' +
  '<div><b><span style="color:rgb(31,73,125);font-family:Tahoma,sans-serif;font-size:10pt">John Fluman</span></b></div>' +
  '<div><b><span style="color:rgb(31,73,125);font-family:Arial,sans-serif;font-size:8pt">Intransit Technologies</span></b></div>' +
  '<div><a href="mailto:john.fluman@intransittech.com" style="font-family:Calibri;font-size:8pt">john.fluman@intransittech.com</a></div>' +
  '<div><i><span style="color:gray;font-family:Arial,sans-serif;font-size:7.5pt">An ISO 9001 Certified Company</span></i></div>' +
  '<div><span style="color:rgb(31,73,125);font-family:Tahoma,sans-serif;font-size:8pt">Toll (877) 677-5868 x101 - Local (949) 481-7935 x101</span></div>' +
  '<br>' +
  '<div><span style="color:rgb(166,166,166);font-family:Calibri,sans-serif;font-size:8pt">' +
  'The information contained in this communication and its attachment(s) is intended only for the use of the individual ' +
  'to whom it is addressed and may contain information that is privileged, confidential, or exempt from disclosure. ' +
  'If the reader of this message is not the intended recipient, you are hereby notified that any dissemination, ' +
  'distribution, or copying of this communication is strictly prohibited. If you have received this communication in ' +
  'error, please notify <a href="mailto:john.fluman@intransittech.com" style="font-family:Calibri;font-size:8pt">john.fluman@intransittech.com</a> ' +
  'and delete the communication without retaining any copies. Thank you.</span></div>';

// ── Main trigger function ────────────────────────────────────
function runEmailAgent() {
  var _cfg = getRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) { hubLog('run', 'runEmailAgent: disabled via hub config'); return; }

  var label   = GmailApp.getUserLabelByName(AGENT_LABEL) || GmailApp.createLabel(AGENT_LABEL);
  var query   = 'in:inbox -label:' + AGENT_LABEL + ' newer_than:2d';
  var threads = GmailApp.search(query, 0, 15);

  hubLog('run', 'runEmailAgent: ' + threads.length + ' thread(s) to evaluate');
  if (!threads.length) return;

  threads.forEach(function(thread) {
    try {
      processThreadWithAgent(thread, label);
    } catch(e) {
      hubLog('error', 'runEmailAgent error on thread: ' + e.toString());
    }
    Utilities.sleep(1500); // avoid Gmail rate limits
  });
}

// ── Process one thread ───────────────────────────────────────
function processThreadWithAgent(thread, agentLabel) {
  var messages = thread.getMessages();
  var firstMsg = messages[0];
  var lastMsg  = messages[messages.length - 1];
  var subject  = thread.getFirstMessageSubject();
  var threadId = thread.getId();
  var sender   = firstMsg.getFrom();

  // Skip blocked domains
  var BLOCKED = ['sourceschip.com', 'bulechip.com'];
  for (var b = 0; b < BLOCKED.length; b++) {
    if (sender.toLowerCase().indexOf(BLOCKED[b]) >= 0) {
      thread.addLabel(agentLabel);
      return;
    }
  }

  // Skip internal emails
  if (sender.toLowerCase().indexOf('intransittech.com') >= 0) {
    thread.addLabel(agentLabel);
    return;
  }

  // Build thread text (capped to keep token cost low)
  var threadText = buildAgentThreadText(messages, 6000);

  // Extract MPN for OEM EXCESS check
  var mpn = extractMPNFromSubject(subject);

  // Check OEM EXCESS + Forte 60-day
  var oemResults   = [];
  var forteResults = [];
  if (mpn) {
    try {
      var webUrl = 'https://script.google.com/macros/s/AKfycbyuuBmiYVW5mKI82D5YQGPh1nNGLJZzlLKoxuOdtmOUwUe75VlhhakqgwKooZu5LHFK/exec'
        + '?key=baSDJ%23444FE%268&mpn=' + encodeURIComponent(mpn);
      var resp = UrlFetchApp.fetch(webUrl, { followRedirects: true, muteHttpExceptions: true });
      var data = JSON.parse(resp.getContentText());
      oemResults   = data.oem_excess  || [];
      forteResults = data.forte_sheet || [];
    } catch(e) {
      hubLog('error', 'runEmailAgent OEM check error: ' + e.toString());
    }
  }

  // Call Worker /api/email-agent
  var payload = JSON.stringify({
    thread_id:       threadId,
    last_message_id: lastMsg.getId(),
    subject:         subject,
    sender:          sender,
    thread_content:  threadText,
    oem_results:     oemResults,
    forte_results:   forteResults,
    current_labels:  thread.getLabels().map(function(l){ return l.getName(); }),
  });

  var workerResp = UrlFetchApp.fetch(HUB_URL + '/api/email-agent', {
    method:          'post',
    contentType:     'application/json',
    headers:         { Authorization: 'Bearer ' + HUB_SECRET },
    payload:         payload,
    muteHttpExceptions: true,
  });

  // Mark thread as seen by agent — always, regardless of outcome
  thread.addLabel(agentLabel);

  if (workerResp.getResponseCode() !== 200) {
    hubLog('error', 'runEmailAgent Worker ' + workerResp.getResponseCode() + ': ' + workerResp.getContentText().substring(0, 200));
    return;
  }

  var decision = JSON.parse(workerResp.getContentText());
  var action   = decision.action;

  hubLog('run', 'runEmailAgent [' + action + ']: ' + subject + ' — ' + (decision.reasoning || '(no reasoning)'));

  // Silent actions — nothing to do
  if (action === 'no_action' || action === 'no_bid') return;

  // Create Gmail draft reply with full HTML signature
  if (decision.draft_body) {
    var plainBody = decision.draft_body;
    var htmlBody  = '<div dir="ltr">' + plainBody.replace(/\n/g, '<br>') + '</div>' + AGENT_HTML_SIG;
    var draft     = lastMsg.createDraftReply(plainBody, { htmlBody: htmlBody });
    var draftId   = draft ? draft.getId() : null;

    // Log to Hub dashboard
    hubPostDraft(
      threadId,
      decision.mpn || mpn || '(unknown)',
      sender,
      subject,
      '[AGENT:' + action + '] ' + decision.draft_body,
      draftId
    );

    // Update agent_decisions row with draft ID
    if (draftId && decision.id) {
      try {
        UrlFetchApp.fetch(HUB_URL + '/api/agent-decisions/' + decision.id, {
          method:      'PATCH',
          contentType: 'application/json',
          headers:     { Authorization: 'Bearer ' + HUB_SECRET },
          payload:     JSON.stringify({ status: 'drafted', gmail_draft_id: draftId }),
          muteHttpExceptions: true,
        });
      } catch(e) {}
    }

    Logger.log('Agent draft created [' + action + ']: ' + subject);
  }

  // Add to Forte if agent provided a forte_entry
  var fe = decision.forte_entry;
  if (fe && fe.mpn && fe.qty && fe.target_price) {
    // Double-check 60-day duplicate (agent already checked, but verify locally)
    var alreadyThere = checkForteForMPN(fe.mpn, 60);
    if (!alreadyThere) {
      addToForteSheet(fe.mpn, fe.qty, fe.target_price, fe.country || '', '');
      hubLog('run', 'Agent added to Forte: ' + fe.mpn + ' qty=' + fe.qty + ' tp=' + fe.target_price);
    } else {
      hubLog('run', 'Agent skipped Forte (60-day duplicate): ' + fe.mpn);
    }
  }
}

// ── Helper: build thread text for Claude ────────────────────
function buildAgentThreadText(messages, maxChars) {
  var parts = [];
  messages.forEach(function(m, i) {
    var body  = m.getPlainBody() || '';
    // Strip deep quote chains to save tokens
    var lines   = body.split('\n');
    var cleaned = [];
    var quoteDepth = 0;
    for (var j = 0; j < lines.length; j++) {
      if (lines[j].match(/^[>\|]/)) {
        quoteDepth++;
        if (quoteDepth <= 4) cleaned.push(lines[j]); // keep shallow quotes for context
      } else {
        quoteDepth = 0;
        cleaned.push(lines[j]);
      }
    }
    parts.push(
      '--- Message ' + (i + 1) + ' | From: ' + m.getFrom() + ' | ' + m.getDate() + ' ---\n' +
      cleaned.join('\n').trim()
    );
  });
  var full = parts.join('\n\n');
  return full.length > maxChars ? full.substring(0, maxChars) + '\n[truncated]' : full;
}

// ── Helper: extract MPN from subject line ────────────────────
function extractMPNFromSubject(subject) {
  if (!subject) return null;
  // netCOMPONENTS: "RFQ from netCOMPONENTS Member (CompanyName | MPN)"
  var nc = subject.match(/\|\s*([A-Z0-9][A-Z0-9\-\/\.\s]{2,40})\s*\)?$/i);
  if (nc) return nc[1].trim().replace(/\s{2,}/g, ' ');
  // IC Source / generic: contains "RFQ" then MPN
  var ic = subject.match(/RFQ[:\-\s]+([A-Z0-9][A-Z0-9\-\/\.]{4,})/i);
  if (ic) return ic[1].trim();
  // Website RFQ: just the subject may be the MPN
  return null;
}

// ============================================================
// ADD TO setupTriggers() — insert this line:
//   ScriptApp.newTrigger('runEmailAgent').timeBased().everyMinutes(5).create();
// Then run setupTriggers() once.
// ============================================================
