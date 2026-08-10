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


// ONE-TIME: Add IDT7202LA50P* to Stan's RFQ sheet (was misrouted as own_stock on Jul 16 2026)
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
    // If user's message contains an MPN not found in the subject, use that instead
    var msgMpn = extractMPNFromText(message);
    if (msgMpn && (!mpn || msgMpn.toUpperCase() !== mpn.toUpperCase())) {
      mpn = msgMpn;
    }

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


function addonDismissCard(e) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().popCard()).build();
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


function addonManualFixDraftFromDiag(e) {
  var formInputs = (e.commonEventObject && e.commonEventObject.formInputs) || {};
  var correction = (formInputs.draftDiagCorrection && formInputs.draftDiagCorrection.stringInputs && formInputs.draftDiagCorrection.stringInputs.value && formInputs.draftDiagCorrection.stringInputs.value[0] || '').trim();
  if (!correction) return notify('Please type your correction first.');
  var params = e.commonEventObject.parameters;
  // Build a synthetic params with correction injected, reuse agreeAndFix logic
  e.commonEventObject.parameters.correction = correction;
  return addonAgreeAndFixDraft(e);
}


function addonProcessNext(e) {
  try {
    var threadId = e && e.gmail && e.gmail.threadId;
    var statusCard;
    if (threadId) {
      // Contextual: label THIS thread PENDING so the 1-min processPendingThreads trigger picks it up
      gmailModifyThread_(threadId, [PENDING_LABEL], []);
      statusCard = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle('Intransit Assistant').setSubtitle('Queued for AI processing'))
        .addSection(CardService.newCardSection()
          .addWidget(CardService.newTextButton()
            .setText('Process Next Email')
            .setBackgroundColor('#1a3c6d')
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(CardService.newAction().setFunctionName('addonProcessNext')))
          .addWidget(CardService.newTextParagraph()
            .setText('This email is queued. Draft will appear in Drafts within ~1 minute.\nClose and reopen this panel to see results.')))
        .build();
    } else {
      // Homepage: scan inbox for any new threads
      fastScanInbox();
      statusCard = buildHomepageCard();
    }
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(statusCard))
      .build();
  } catch(err) {
    var errCard = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Intransit Assistant').setSubtitle('Error'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText('Error: ' + err.toString()))).build();
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(errCard))
      .build();
  }
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


function addonSaveStockPrice(e) {
  try {
    var formInputs = (e.commonEventObject && e.commonEventObject.formInputs) || {};
    var mpn = ((formInputs.stockPriceMpn || {}).stringInputs || {}).value;
    mpn = (mpn && mpn[0] ? mpn[0].trim().toUpperCase() : '');
    var priceStr = ((formInputs.stockPriceValue || {}).stringInputs || {}).value;
    priceStr = (priceStr && priceStr[0] ? priceStr[0].trim() : '');
    if (!mpn || !priceStr || isNaN(parseFloat(priceStr))) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('⚠️ Enter a valid MPN and price.'))
        .build();
    }
    var price = parseFloat(priceStr);
    UrlFetchApp.fetch(HUB_URL + '/api/stock-prices', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({ mpn: mpn, price: price }),
      muteHttpExceptions: true
    });
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('✅ Price saved: ' + mpn + ' = $' + price.toFixed(2) + ' each'))
      .build();
  } catch(err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('❌ ' + err.toString()))
      .build();
  }
}


function addonClearStockPrice(e) {
  try {
    var params = (e.commonEventObject && e.commonEventObject.parameters) || {};
    var mpn = (params.mpn || '').toUpperCase();
    if (!mpn) throw new Error('No MPN');
    UrlFetchApp.fetch(HUB_URL + '/api/stock-prices?mpn=' + encodeURIComponent(mpn), {
      method: 'delete',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      muteHttpExceptions: true
    });
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('✅ Price cleared for ' + mpn))
      .build();
  } catch(err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('❌ ' + err.toString()))
      .build();
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


function addonSendNetCom(e) {
  try {
    // Queue via command queue — avoids trigger limit entirely.
    // processCommandQueue (runs every 5 min) picks this up and calls sendPleasePostViaREST.
    UrlFetchApp.fetch(HUB_URL + '/api/command-queue', {
      method: 'POST', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({ type: 'send_datamaster_email', data: {} }),
      muteHttpExceptions: true
    });
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('✅ Queued — NetCOMPONENTS email sending within ~5 min'))
      .build();
  } catch(err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('❌ Error: ' + err.toString()))
      .build();
  }
}


function addonSheetLookup(mpn) {
  var resp = UrlFetchApp.fetch(HUB_URL + '/api/sheet-lookup?mpn=' + encodeURIComponent(mpn), {
    headers: { Authorization: 'Bearer ' + HUB_SECRET },
    muteHttpExceptions: true, followRedirects: true
  });
  return JSON.parse(resp.getContentText());
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
    createThreadedDraft(toEmail, 'Re: ' + subject, html, lastMsg.getId(), threadId);
    hubLog('run', 'addonSmartReplyCreateDraft: draft created for ' + subject, { to: toEmail });

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().popCard())
      .setNotification(CardService.newNotification().setText('✅ Draft created — check Gmail!'))
      .build();
  } catch(err) { return notify('Error creating draft: ' + err.toString()); }
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


function addonWrongDraftApply(e) {
  try {
    var params     = e.commonEventObject.parameters;
    var draftId    = params.draftId  || '';
    var threadId   = params.threadId || '';
    var subject    = params.subject  || '';
    var toEmail    = params.toEmail  || '';
    var formInputs = e.commonEventObject.formInputs || {};

    var missed = params.missed === '1';
    var correctAction = '';
    if (formInputs.correct_action && formInputs.correct_action.stringInputs) {
      correctAction = (formInputs.correct_action.stringInputs.value || [])[0] || '';
    }
    var wrongReason = '';
    if (formInputs.wrong_reason && formInputs.wrong_reason.stringInputs) {
      wrongReason = (formInputs.wrong_reason.stringInputs.value || [])[0] || '';
    }

    if (!correctAction) return notify('Please select the correct action first.');

    var token = ScriptApp.getOAuthToken();

    // Capture wrong draft body before deleting (skip if no draft)
    var wrongDraftBody = '';
    if (draftId) {
      try {
        var fetchResp = UrlFetchApp.fetch(
          'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + draftId + '?format=full',
          { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
        );
        if (fetchResp.getResponseCode() === 200) {
          var draftData = JSON.parse(fetchResp.getContentText());
          var rawHtml = extractDraftHtmlBody(draftData.message && draftData.message.payload);
          wrongDraftBody = rawHtml ? rawHtml.replace(/<[^>]+>/g,' ').replace(/\s+/g,' ').trim().substring(0,400) : '';
        }
      } catch(eF) {}
    }

    // Delete the wrong draft (only if one existed)
    if (draftId) {
      UrlFetchApp.fetch(
        'https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + draftId,
        { method: 'DELETE', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
      );
    }

    // Find the correct action definition
    var actionDef = null;
    WRONG_DRAFT_ACTIONS.forEach(function(a) { if (a.key === correctAction) actionDef = a; });

    var newDraftBody = actionDef ? actionDef.body : null;
    var notificationText = '';

    if (newDraftBody) {
      // Fixed-template action — create the correct draft now
      try {
        var thread = GmailApp.getThreadById(threadId);
        if (thread) {
          var messages = thread.getMessages();
          var lastMsg  = messages[messages.length - 1];
          var htmlBody = buildSimpleHTML(newDraftBody);
          lastMsg.createDraftReply('', { htmlBody: htmlBody });
          notificationText = missed
            ? '✅ Draft created: ' + actionDef.label + '. Bug logged.'
            : '✅ Wrong draft deleted. Correct draft created: ' + actionDef.label;
        } else {
          notificationText = '⚠️ Could not find thread to create draft.';
        }
      } catch(eD) {
        notificationText = '⚠️ Error creating draft: ' + eD.toString();
      }
    } else {
      // Dynamic action (own_stock, stan_quoted, no_bid, no_action) — can't auto-generate
      notificationText = missed
        ? '✅ Bug logged. Action "' + correctAction + '" requires manual draft — please create it yourself.'
        : '✅ Wrong draft deleted. Action "' + correctAction + '" requires manual draft — please create it yourself.';
    }

    // Log the lesson via hubLearn
    if (wrongReason.trim()) {
      var lessonPrefix = missed ? '[MISSED DRAFT] ' : '[WRONG DRAFT] ';
      hubLearn(
        lessonPrefix + wrongReason + ' [Correct action should have been: ' + correctAction + ']',
        wrongDraftBody,
        newDraftBody || '[' + correctAction + ' — dynamic, not auto-generated]',
        threadId, subject, toEmail, '', correctAction
      );
    }

    hubLog('feedback', 'Wrong draft corrected: ' + correctAction + ' | Reason: ' + (wrongReason || '(none)'), {
      draft_id: draftId, thread_id: threadId, correct_action: correctAction
    });

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().popCard())
      .setNotification(CardService.newNotification().setText(notificationText))
      .build();

  } catch(err) {
    return notify('addonWrongDraftApply error: ' + err.toString());
  }
}


// ─── WRONG DRAFT PICKER ──────────────────────────────────────────────────────

var WRONG_DRAFT_ACTIONS = [
  { key: 'request_tp_500',  label: 'Ask for TP ($500 min)',
    body: 'We need a target price to proceed. Please note there is a $500 minimum line requirement. Once we have your target we will get back to you right away.' },
  { key: 'request_tp_2000', label: 'Ask for TP ($2,000 min)',
    body: 'We need a target price to proceed. Please note there is a $2,000 minimum line requirement. Once we have your target we will get back to you right away.' },
  { key: 'msg_checking',    label: 'Checking on it (MSG_CHECKING)',
    body: 'We are checking on it now. If we get a response from the OEM, I will respond to you right away. If we do not respond back to you, please consider this a no bid. Thank you very much for the opportunity.' },
  { key: 'add_to_stan',     label: 'Warehouse checking (add to Stan)',
    body: 'Warehouse is checking details and I will update ASAP' },
  { key: 'bill_handle',     label: 'Bill will help',
    body: 'Bill will help with this request' },
  { key: 'remove_oem',      label: 'Remove from listing (David no-stk)',
    body: 'Ok, removed from listing.' },
  { key: 'own_stock',       label: 'Own stock — quote directly (dynamic)',
    body: null },
  { key: 'stan_quoted',     label: 'Stan quoted — send pricing (dynamic)',
    body: null },
  { key: 'no_bid',          label: 'No bid (silent — no reply)',
    body: null },
  { key: 'no_action',       label: 'No action needed (skip)',
    body: null },
];

function addonWrongDraftPicker(e) {
  try {
    var params = e.commonEventObject.parameters;
    var draftId  = params.draftId  || '';
    var threadId = params.threadId || '';
    var subject  = params.subject  || '';
    var toEmail  = params.toEmail  || '';
    var missed   = params.missed   === '1';

    var card  = CardService.newCardBuilder().setHeader(
      CardService.newCardHeader().setTitle(missed ? '📋 Missed Draft — Pick Action' : '📋 Pick Correct Response')
    );
    var sect = CardService.newCardSection()
      .setHeader(missed ? 'What should have been sent?' : 'What should the draft have said?');

    // Dropdown of all standard actions
    var dropdown = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setFieldName('correct_action')
      .setTitle('Correct action');
    WRONG_DRAFT_ACTIONS.forEach(function(a) {
      dropdown.addItem(a.label, a.key, false);
    });
    sect.addWidget(dropdown);

    // Reason text box — label differs for missed vs wrong
    sect.addWidget(CardService.newTextInput()
      .setFieldName('wrong_reason')
      .setTitle(missed ? 'Why should this have triggered?' : 'Why was the draft wrong?')
      .setHint('Used to train the AI — be specific')
      .setMultiline(true));

    // Submit button
    sect.addWidget(CardService.newTextButton()
      .setText(missed ? 'Create Correct Draft & Log Bug' : 'Fix Draft & Log Bug')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#1565c0')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonWrongDraftApply')
        .setParameters({ draftId: draftId, threadId: threadId, subject: subject, toEmail: toEmail, missed: missed ? '1' : '0' })));

    card.addSection(sect);
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card.build()))
      .build();
  } catch(err) {
    return notify('addonWrongDraftPicker error: ' + err.toString());
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


function archiveBlockedDomains() {
  var BLOCKED_DOMAINS = getBlockedDomains();
  BLOCKED_DOMAINS.forEach(function(domain) {
    var ids = gmailSearchREST('in:inbox from:' + domain, 50);
    if (!ids.length) return;
    ids.forEach(function(tid) { gmailArchiveThread_(tid); });
    hubLog('run', 'archiveBlockedDomains: archived ' + ids.length + ' from ' + domain);
  });
}


function buildAddonError(msg) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Intransit Assistant').setSubtitle('Error'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText(msg)))
    .build();
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


// Contextual card — fires when any email is opened; shows thread-relevant info.
function buildContextualCard(e) {
  try {
    var gmailThreadId = e.gmail && e.gmail.threadId;
    if (!gmailThreadId) return buildHomepageCard();

    // Use GmailApp service instead of REST — avoids premium UrlFetch quota entirely
    var gmailThread = GmailApp.getThreadById(gmailThreadId);
    var subject = '', fromH = '', matchToH = '', matchDraftId = null;
    if (gmailThread) {
      subject = gmailThread.getFirstMessageSubject();
      var threadMsgs = gmailThread.getMessages();
      if (threadMsgs.length > 0) fromH = threadMsgs[0].getFrom();
    }

    // Find draft for this thread — GmailApp.getDrafts() uses Gmail service quota, not UrlFetch
    var allDrafts = GmailApp.getDrafts();
    for (var di = 0; di < allDrafts.length; di++) {
      try {
        var dm = allDrafts[di].getMessage();
        if (dm.getThread().getId() === gmailThreadId) {
          matchDraftId = allDrafts[di].getId();
          matchToH = dm.getTo();
          break;
        }
      } catch(de) { Logger.log('Draft scan error: ' + de); }
    }

    var label = (subject || 'Email').replace(/^Re:\s*/i, '').substring(0, 55);
    var builder = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Intransit Assistant')
        .setSubtitle(label));

    if (matchDraftId) {
      var fbField = 'fb_' + matchDraftId.replace(/[^a-zA-Z0-9]/g, '_');

      var infoSection = CardService.newCardSection().setHeader('📝 Draft ready');
      infoSection.addWidget(CardService.newTextParagraph().setText('To: ' + (matchToH || 'unknown')));
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
      fixSection.addWidget(CardService.newTextButton()
        .setText('📋 Wrong Draft — Pick Correct Response')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#e65100')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('addonWrongDraftPicker')
          .setParameters({ draftId: matchDraftId, threadId: gmailThreadId, subject: subject, toEmail: matchToH })));

      builder.addSection(infoSection).addSection(fixSection);

    } else {
      var noSection = CardService.newCardSection().setHeader('📥 No draft — create one now');
      noSection.addWidget(CardService.newTextParagraph().setText('From: ' + (fromH || 'unknown')));
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
      noSection.addWidget(CardService.newTextButton()
        .setText('📋 Missed Draft — Pick What Should Have Been Sent')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#e65100')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('addonWrongDraftPicker')
          .setParameters({ draftId: '', threadId: gmailThreadId, subject: subject, toEmail: fromH, missed: '1' })));
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

    // ── Stock Price ─────────────────────────────────────────────────────────
    // Remembers a per-MPN sell price so own_stock drafts auto-fill the price.
    var spMpn = invMpnHint;
    var spCurrent = null;
    if (spMpn) {
      try {
        var spResp = UrlFetchApp.fetch(HUB_URL + '/api/stock-prices?mpn=' + encodeURIComponent(spMpn), {
          headers: { Authorization: 'Bearer ' + HUB_SECRET }, muteHttpExceptions: true
        });
        spCurrent = JSON.parse(spResp.getContentText()).price;
      } catch(esp) {}
    }
    var priceSection = CardService.newCardSection()
      .setHeader('💰 Stock Price' + (spCurrent != null ? ' — $' + parseFloat(spCurrent).toFixed(2) + ' each' : ' — not set'));
    priceSection.addWidget(CardService.newTextInput()
      .setFieldName('stockPriceMpn')
      .setTitle('MPN')
      .setValue(spMpn)
      .setHint('Pre-filled from subject — change if needed'));
    priceSection.addWidget(CardService.newTextInput()
      .setFieldName('stockPriceValue')
      .setTitle('Price per unit ($)')
      .setValue(spCurrent != null ? String(parseFloat(spCurrent).toFixed(2)) : '')
      .setHint('e.g. 4.50 — saved until you clear it'));
    priceSection.addWidget(CardService.newTextButton()
      .setText('💾 Save Price')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#1a7340')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('addonSaveStockPrice')
        .setParameters({})));
    if (spCurrent != null) {
      priceSection.addWidget(CardService.newTextButton()
        .setText('✕ Clear Saved Price')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#7b1fa2')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('addonClearStockPrice')
          .setParameters({ mpn: spMpn })));
    }
    builder.addSection(priceSection);

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

    builder.setFixedFooter(CardService.newFixedFooter()
      .setPrimaryButton(CardService.newTextButton()
        .setText('Process Next Email')
        .setBackgroundColor('#1a3c6d')
        .setOnClickAction(CardService.newAction().setFunctionName('addonProcessNext'))));

    return [builder.build()];
  } catch(err) {
    return [buildAddonError(err.toString())];
  }
}


function buildDraftHTML(replyText, originalMessage) {
  var sig = getSignatureHTML();
  var origDate = Utilities.formatDate(originalMessage.getDate(), Session.getScriptTimeZone(), 'EEE, MMM d, yyyy, h:mm a');
  var origFrom = originalMessage.getFrom();
  var origBody = originalMessage.getBody() || originalMessage.getPlainBody().replace(/\n/g, '<br>');
  var quoted = '<br><div class="gmail_quote"><div dir="ltr" class="gmail_attr">On ' + origDate + ', ' + origFrom + ' wrote:<br></div>'
    + '<blockquote class="gmail_quote" style="margin:0px 0px 0px 0.8ex;border-left:1px solid rgb(204,204,204);padding-left:1ex">'
    + origBody + '</blockquote></div>';
  var htmlText = replyText.replace(/\n/g, '<br>');
  return '<div dir="ltr">' + htmlText + sig + quoted + '</div>';
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


// ── Gmail Add-on: sidebar card ────────────────────────────────
function buildGmailHomepage(e) {
  return buildHubCard_('Ready — tap the button to process the next email.');
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

    // Process Next Email button — always at top
    var processSection = CardService.newCardSection();
    processSection.addWidget(
      CardService.newTextButton()
        .setText('Process Next Email')
        .setBackgroundColor('#1a3c6d')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('addonProcessNext'))
    );
    processSection.addWidget(
      CardService.newTextButton()
        .setText('Fix Claude Drafts')
        .setBackgroundColor('#1565c0')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('addonFixClaudeDrafts'))
    );
    builder.addSection(processSection);

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


// Card action: schedules background fix via trigger, returns immediately to avoid 30s timeout
function addonFixClaudeDrafts() {
  try {
    // Delete any existing pending fix triggers first (no trailing underscore)
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'runClaudeDraftFix') ScriptApp.deleteTrigger(t);
    });
    // Schedule background run
    ScriptApp.newTrigger('runClaudeDraftFix').timeBased().after(3000).create();

    var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Fix Claude Drafts')
        .setSubtitle('Processing in background...'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph()
          .setText('Scanning and fixing "claude" drafts now. Check your Gmail drafts in ~30 seconds. Results are logged in Apps Script execution log.')));

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card.build()))
      .build();
  } catch(err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Error: ' + err.toString()))
      .build();
  }
}

// Background trigger handler — runs findAndFixClaudeDrafts with full 6-min Apps Script limit
function runClaudeDraftFix() {
  // Clean up this trigger
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runClaudeDraftFix') ScriptApp.deleteTrigger(t);
  });
  try {
    var results = findAndFixClaudeDrafts();
    Logger.log('Fix Claude Drafts: found ' + results.length + ' draft(s)');
    results.forEach(function(r) {
      try {
        applyClaudeDraftFix(r.draft_id, r.thread_id, r.to_email, r.subject, r.corrected_body);
        Logger.log('Fixed: ' + r.subject + ' → ' + r.corrected_body.substring(0, 80));
      } catch(e) {
        Logger.log('Error fixing "' + r.subject + '": ' + e.toString());
      }
    });
    Logger.log('Fix Claude Drafts: complete');
  } catch(e) {
    Logger.log('runClaudeDraftFix_ error: ' + e.toString());
  }
}


function buildHubCard_(statusText, isError) {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Intransit Hub'));
  var section = CardService.newCardSection();
  var btn = CardService.newTextButton()
    .setText('Process Next Email')
    .setBackgroundColor('#1a3c6d')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(CardService.newAction().setFunctionName('addonProcessNext'));
  section.addWidget(btn);
  if (statusText) {
    section.addWidget(CardService.newTextParagraph().setText(statusText));
  }
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


// ── Trigger 8 — Bill netcomp removals → delete from OEM EXCESS
function checkBillNetcompRemovals() {
  var _cfg = getRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) { hubLog('run', 'checkBillNetcompRemovals: disabled'); return; }
  var BILL_EMAIL = 'bill.pratt@intransittech.com';
  var DONE_LABEL = 'oem-bill-removal-processed';
  var query = 'from:' + BILL_EMAIL +
    ' (netcomp OR netcomponents) (remove OR removing OR removed)' +
    ' -label:' + DONE_LABEL + ' newer_than:14d';
  var threadIds = gmailSearchREST(query, 20);
  hubLog('run', 'checkBillNetcompRemovals: ' + threadIds.length + ' thread(s)');
  if (!threadIds.length) return;

  threadIds.forEach(function(tid) {
    var thread = GmailApp.getThreadById(tid);
    if (!thread) return;
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
      gmailArchiveThread_(tid);
    } else {
      Logger.log('Bill removal: could not extract MPN from body or subject: ' + subject);
      GmailApp.sendEmail(NOTIFY_EMAIL, 'OEM EXCESS: Bill removal — could not extract MPN',
        'Subject: "' + subject + '"\nhttps://mail.google.com/mail/u/0/#inbox/' + thread.getId());
    }
    gmailModifyThread_(tid, [DONE_LABEL], []);
  });
}


// ── Trigger 1 — David no-stock emails ────────────────────────
function checkDavidNoStockEmails() {
  var _cfg = getRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) return;
  // Dual query: all-mail search (for Gmail-filtered/archived emails) + explicit inbox search
  // (Gmail filter occasionally misses David emails, leaving them in inbox only).
  // oem-rfq-incoming-processed is applied to ALL processed David threads (no-stk or not)
  // so they're properly excluded on re-runs. oem-nostock-seen can't be used — Gmail filter
  // auto-applies it to every David email regardless of processing status.
  var seenIds = {};
  var allIds = [];
  var davidQ = '(from:' + DAVID_EMAIL + ' OR from:david@fortecomp.com)';
  var q1 = davidQ + ' -label:oem-rfq-incoming-processed newer_than:14d';
  var q2 = 'in:inbox ' + davidQ + ' -label:oem-rfq-incoming-processed';
  // Use GmailApp.search (reliable, properly authorized) instead of REST API which silently fails
  GmailApp.search(q1, 0, 50).concat(GmailApp.search(q2, 0, 50)).forEach(function(t) {
    var tid = t.getId();
    if (!seenIds[tid]) { seenIds[tid] = true; allIds.push(tid); }
  });
  hubLog('run', 'checkDavidNoStockEmails: ' + allIds.length + ' thread(s)');
  if (!allIds.length) return;
  var noStkKeywords = ['no stk', 'no stock', 'stk sold', 'stock sold', 'cant find', 'cant share', 'cannot find', 'removed', 'no inventory',
                       'sold lying commie', 'soly lying commie', 'lying commie', 'sold out', 'all sold', 'no longer have'];
  allIds.forEach(function(tid) {
    try {
      var thread = GmailApp.getThreadById(tid);
      if (!thread) { gmailModifyThread_(tid, ['oem-rfq-incoming-processed'], []); return; }
      var msg = thread.getMessages()[thread.getMessageCount() - 1];
      var subject = msg.getSubject();
      var subjectLower = subject.toLowerCase();
      var bodySnippet = msg.getPlainBody().toLowerCase().substring(0, 300);
      var isNoStk = noStkKeywords.some(function(kw) { return subjectLower.indexOf(kw) >= 0 || bodySnippet.indexOf(kw) >= 0; });
      if (!isNoStk) {
        // Not a no-stk — mark processed so it stops appearing in search results
        gmailModifyThread_(tid, ['oem-rfq-incoming-processed'], []);
        return;
      }
      // Stamp Forte col K directly by row number extracted from subject (#XXXX)
      // This runs before processThread so it always fires even if worker decision fails
      var rowMatch = subject.match(/#(\d{3,5})/);
      if (rowMatch) {
        var forteRow = parseInt(rowMatch[1], 10);
        try {
          var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
          var statusCell = forteSheet.getRange(forteRow, FORTE_STATUS_COL + 1);
          var currentVal = String(statusCell.getValue()).trim();
          if (currentVal.toUpperCase().indexOf('NO STK') === -1 && currentVal.toUpperCase() !== 'CLOSED') {
            var stamp = 'NO STK - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
            statusCell.clearDataValidations();
            statusCell.setValue(stamp);
            statusCell.setBackground('#000000'); statusCell.setFontColor('#FFFFFF'); statusCell.setFontWeight('bold');
            hubLog('run', 'checkDavidNoStockEmails: stamped Forte row ' + forteRow + ' directly from subject', {});
          }
        } catch(e) {
          hubLog('error', 'checkDavidNoStockEmails: Forte row stamp error row ' + forteRow + ': ' + e, {});
        }
      }
      try { processThread(thread); } catch(e) { hubLog('error', 'checkDavidNoStockEmails processThread error: ' + e, {}); }
      gmailModifyThread_(tid, [INCOMING_LABEL, 'oem-rfq-incoming-processed'], []);
      gmailArchiveThread_(tid);
    } catch(e) {
      hubLog('error', 'checkDavidNoStockEmails thread ' + tid + ' error: ' + e, {});
    }
  });
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


// ── Forte no-stock reply handler ──────────────────────────────



// ── Trigger 6 — Payment Advice → forward to Deb ──────────────
function checkInboxForPaymentAdvice() {
  var _cfg = getRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) { hubLog('run', 'checkInboxForPaymentAdvice: disabled via hub config'); return; }
  var query = 'in:inbox subject:"payment advice" -label:oem-payment-forwarded';
  var threadIds = gmailSearchREST(query, 10);
  hubLog('run', 'checkInboxForPaymentAdvice: ' + threadIds.length + ' thread(s)');
  if (!threadIds.length) return;
  threadIds.forEach(function(tid) {
    var thread = GmailApp.getThreadById(tid);
    if (!thread) return;
    var messages = thread.getMessages();
    var firstMsg = messages[0];
    if (firstMsg.getFrom().toLowerCase().indexOf('intransittech.com') >= 0) {
      gmailModifyThread_(tid, ['oem-payment-forwarded'], []);
      Logger.log('Payment advice skipped — internal sender: ' + firstMsg.getFrom());
      return;
    }
    var recipients = (firstMsg.getTo() + ',' + (firstMsg.getCc() || '')).toLowerCase();
    if (recipients.indexOf(DEB_EMAIL.toLowerCase()) >= 0) {
      gmailModifyThread_(tid, ['oem-payment-forwarded'], []);
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
    gmailModifyThread_(tid, ['oem-payment-forwarded'], []);
  });
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
  // Sidebar page — served as HtmlService so google.script.run works
  if ((e.parameter.page || '') === 'sidebar') {
    return HtmlService.createHtmlOutput(getSidebarHTML_())
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


// ── Two-phase scanning: fastScanInbox (1 min, labels only) ────
// ── + processPendingThreads (5 min, Claude calls) ─────────────
// Root cause of quota exhaustion: runEmailScan calls Claude API
// synchronously every minute. Each processThread() takes 30-60s,
// so 10 threads/run = up to 10 min, blowing the 6-min exec limit
// and consuming the 6-hr daily quota within hours. Fix: label-only
// fast scan every minute, Claude calls on a slower 5-min trigger.

function fastScanInbox() {
  var _cfg = getCachedRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) { hubLog('run', 'fastScanInbox: disabled'); return; }
  try { archiveBlockedDomains(); } catch(e) {}
  var BLOCKED_DOMAINS = getBlockedDomains();
  var blockFilter = BLOCKED_DOMAINS.map(function(d){ return '-from:' + d; }).join(' ');

  // Pre-warm PENDING_LABEL via GmailApp so getLabelId_ always hits the in-memory cache.
  // Without this: if gmailREST_('/labels') fails, getLabelId_(PENDING_LABEL) returns null,
  // filter(Boolean) in gmailModifyThread_ silently drops it, and threads get
  // oem-rfq-incoming-processed but NOT PENDING_LABEL — permanently orphaned.
  var _plObj = GmailApp.getUserLabelByName(PENDING_LABEL) || GmailApp.createLabel(PENDING_LABEL);
  _gmailLabelCache_[PENDING_LABEL] = _plObj.getId();

  // All label lookups via REST (cached in Script Properties — zero premium gmail quota cost)
  var rfqLabelId     = getLabelId_('oem-rfq-incoming-processed');
  var tpLabelId      = getLabelId_('oem-tp-processed');
  var agentLabelId   = getLabelId_(AGENT_LABEL);
  var pendingLabelId = getLabelId_(PENDING_LABEL);

  var rfqQ = 'in:inbox (to:rfq@intransittech.com OR deliveredto:rfq@intransittech.com OR subject:rfq OR from:autosend@icsource.com OR subject:"please quote" OR subject:"request for quote" OR subject:"request for quotation" OR subject:"looking for" OR ((to:john.fluman@intransittech.com OR deliveredto:john.fluman@intransittech.com OR to:rfq@intransittech.com) ("quotation" OR "best price" OR "netcomponents" OR "looking for" OR "quote your stock" OR "can you quote" OR "is it in stock" OR "availability"))) -from:intransittech.com -from:fortetechno.com -from:fortecomp.com -label:oem-rfq-incoming-processed ' + blockFilter;
  var rfqCount = 0;
  gmailSearchREST(rfqQ, 50).forEach(function(tid) {
    var meta = gmailGetThreadMeta_(tid);
    if (meta.senders.some(function(s){ return s.indexOf(JOHN_EMAIL) >= 0; })) return;
    if (meta.senders.length && meta.senders[meta.senders.length-1].indexOf('intransittech.com') >= 0) return;
    gmailModifyThread_(tid, ['oem-rfq-incoming-processed', PENDING_LABEL], []);
    rfqCount++;
  });

  var tpQ = 'in:inbox label:oem-rfq-incoming-processed -label:oem-tp-processed -label:' + PENDING_LABEL + ' newer_than:60d ' + blockFilter;
  var tpCount = 0;
  gmailSearchREST(tpQ, 50).forEach(function(tid) {
    var meta = gmailGetThreadMeta_(tid);
    if (meta.senders.length && meta.senders[meta.senders.length-1].indexOf(JOHN_EMAIL) >= 0) return;
    var buyerCount = meta.senders.filter(function(s){ return s.indexOf(JOHN_EMAIL) < 0 && s.indexOf('intransittech') < 0; }).length;
    if (buyerCount < 2) return;
    // Must have at least one John message — confirms we already sent a TP request.
    // Prevents threads with only the original buyer RFQ from being mis-routed as TP replies.
    var hasJohnMsg = meta.senders.some(function(s){ return s.indexOf(JOHN_EMAIL) >= 0 || s.indexOf('intransittech') >= 0; });
    if (!hasJohnMsg) return;
    gmailModifyThread_(tid, ['oem-tp-processed', PENDING_LABEL], []);
    tpCount++;
  });

  // Bug 24 fix: require subject/sender keyword so consumer notification emails
  // (receipts, order confirmations, restaurant alerts) are not sent to the AI worker.
  var agentQ = 'in:inbox -label:' + AGENT_LABEL + ' -label:oem-rfq-incoming-processed -label:' + PENDING_LABEL +
    ' newer_than:3d -from:fortetechno.com -from:fortecomp.com ' + blockFilter +
    ' (subject:rfq OR subject:quot OR subject:offer OR subject:"best price" OR subject:"looking for"' +
    ' OR subject:availability OR subject:qty OR subject:inquiry OR subject:sourcing OR subject:parts' +
    ' OR subject:"request for" OR from:netcomponents.com OR from:icsource.com OR from:messagesend' +
    ' OR quotation OR "looking for" OR "please quote" OR "please check" OR "can you quote" OR "provide the price")';
  var agentCount = 0;
  gmailSearchREST(agentQ, 50).forEach(function(tid) {
    var meta = gmailGetThreadMeta_(tid);
    var firstSender = (meta.senders[0] || '').toLowerCase();
    if (firstSender.indexOf('intransittech.com') >= 0 || firstSender.indexOf('fortetechno.com') >= 0 || firstSender.indexOf('fortecomp.com') >= 0) {
      gmailModifyThread_(tid, [AGENT_LABEL], []);
      return;
    }
    if (rfqLabelId && meta.labelIds.indexOf(rfqLabelId) >= 0) {
      gmailModifyThread_(tid, [AGENT_LABEL], []);
      return;
    }
    gmailModifyThread_(tid, [AGENT_LABEL, 'oem-rfq-incoming-processed', PENDING_LABEL], []);
    agentCount++;
  });

  if (rfqCount + tpCount + agentCount > 0) {
    hubLog('run', 'fastScanInbox: queued ' + (rfqCount+tpCount+agentCount) + ' thread(s) (rfq:' + rfqCount + ' tp:' + tpCount + ' agent:' + agentCount + ')');
  }
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


function getCachedRemoteConfig() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('remote_config_v2');
  if (cached) try { return JSON.parse(cached); } catch(e) {}
  var cfg = getRemoteConfig();
  try { cache.put('remote_config_v2', JSON.stringify(cfg || {}), 300); } catch(e) {}
  return cfg || {};
}


function getLabelId_(labelName) {
  if (_gmailLabelCache_[labelName]) return _gmailLabelCache_[labelName];
  var props = PropertiesService.getScriptProperties();
  var key = 'lbl_' + labelName.replace(/[^a-zA-Z0-9]/g, '_');
  var cached = props.getProperty(key);
  if (cached) { _gmailLabelCache_[labelName] = cached; return cached; }
  var all = gmailREST_('/labels');
  (all.labels || []).forEach(function(l) {
    var k = 'lbl_' + l.name.replace(/[^a-zA-Z0-9]/g, '_');
    props.setProperty(k, l.id);
    _gmailLabelCache_[l.name] = l.id;
  });
  if (_gmailLabelCache_[labelName]) return _gmailLabelCache_[labelName];
  var created = gmailREST_('/labels', 'post', { name: labelName });
  if (created.id) {
    props.setProperty(key, created.id);
    _gmailLabelCache_[labelName] = created.id;
    return created.id;
  }
  return null;
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
    var max = maxThreads || 5;
    // Multi-strategy: subject search first (more precise), then full-text phrase, then loose (Gmail tokenizes hyphens)
    var threads = GmailApp.search('in:sent subject:"' + mpn + '"', 0, max);
    if (!threads.length) threads = GmailApp.search('in:sent "' + mpn + '"', 0, max);
    if (!threads.length) {
      var loose = mpn.replace(/-/g, ' ');
      threads = GmailApp.search('in:sent subject:(' + loose + ')', 0, max);
    }
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


// ── Sidebar: "Fix Claude Drafts" ─────────────────────────────


// ── Sidebar: "Fix Claude Drafts" — finds drafts where body starts with "claude" ──
function findAndFixClaudeDrafts() {
  var results = [];
  var data = gmailREST_('/drafts?maxResults=50');
  var drafts = (data.drafts || []);

  for (var i = 0; i < drafts.length; i++) {
    var d = drafts[i];
    var msgId = d.message && d.message.id;
    if (!msgId) continue;
    try {
      var msg = GmailApp.getMessageById(msgId);
      if (!msg) continue;
      var body = (msg.getPlainBody() || '').trim();
      if (!/^claude\b/i.test(body)) continue;

      var thread    = msg.getThread();
      var threadId  = thread.getId();
      var messages  = thread.getMessages();
      var rawSubj   = thread.getFirstMessageSubject();
      var subject   = 'Re: ' + rawSubj.replace(/^Re:\s*/i, '');
      var toEmail   = msg.getTo();

      // If To header is empty fall back to last non-John sender
      if (!toEmail) {
        for (var mi = messages.length - 1; mi >= 0; mi--) {
          var fe = extractBuyerEmail(messages[mi].getFrom());
          if (fe && fe.indexOf('intransittech.com') === -1) { toEmail = fe; break; }
        }
      }

      // Build thread content for context
      var parts = ['Subject: ' + rawSubj];
      messages.forEach(function(m, idx) {
        var mb = (m.getPlainBody() || '').split('\n')
          .filter(function(ln) { return ln.charAt(0) !== '>'; })
          .join('\n').trim().substring(0, 1500);
        parts.push('--- Msg ' + (idx + 1) + ' | From: ' + m.getFrom() + ' ---');
        parts.push(mb);
      });
      var threadContent = parts.join('\n').substring(0, 5000);

      var fixResp = UrlFetchApp.fetch(HUB_URL + '/api/fix-draft', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          draft_body: 'claude',
          feedback: 'Draft body says "claude" as a placeholder. Determine the correct reply based on the thread context.',
          subject: subject,
          to_email: toEmail,
          thread_content: threadContent
        }),
        headers: { Authorization: 'Bearer ' + HUB_SECRET },
        muteHttpExceptions: true
      });
      var fix = {};
      try { fix = JSON.parse(fixResp.getContentText()); } catch(e) {}

      results.push({
        draft_id:       d.id,
        thread_id:      threadId,
        subject:        subject,
        to_email:       toEmail,
        corrected_body: fix.corrected_body || '(could not determine)',
        advice:         fix.advice || ''
      });
    } catch(e) {
      hubLog('error', 'findAndFixClaudeDrafts: ' + e, {});
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


function getSidebarHTML_() {
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>' +
    '*{box-sizing:border-box}body{font-family:Arial,sans-serif;padding:16px;background:#f5f5f5;margin:0}' +
    'h2{color:#1a3c6d;font-size:16px;margin:0 0 14px;border-bottom:2px solid #1a3c6d;padding-bottom:8px}' +
    '.btn{background:#1a3c6d;color:#fff;border:none;padding:12px;border-radius:5px;cursor:pointer;width:100%;font-size:14px;font-weight:bold;transition:all .2s}' +
    '.btn:hover:not(:disabled){background:#2255a0}.btn:disabled{background:#aaa;cursor:not-allowed}' +
    '.card{margin-top:14px;padding:12px 14px;background:#fff;border-radius:6px;border-left:4px solid #1a3c6d;display:none;box-shadow:0 1px 3px rgba(0,0,0,.1)}' +
    '.card.err{border-left-color:#c00}' +
    '.badge{display:inline-block;padding:2px 8px;border-radius:3px;font-size:11px;font-weight:bold;text-transform:uppercase;margin-bottom:8px;background:#1a3c6d;color:#fff}' +
    '.badge.david_nostock{background:#7b1f1f}.badge.error{background:#c00}.badge.no_action{background:#888}.badge.nothing{background:#888}' +
    '.subj{font-size:13px;font-weight:bold;margin:4px 0;word-break:break-word}' +
    '.frm{color:#777;font-size:11px;margin-bottom:10px;word-break:break-all}' +
    '.draft{background:#eef2ff;padding:9px;border-radius:4px;font-size:12px;white-space:pre-wrap;word-break:break-word;border-left:3px solid #4a6fc7;margin:8px 0}' +
    '.forte{background:#efffef;padding:7px 9px;border-radius:4px;font-size:12px;border-left:3px solid #2a8a2a;margin:6px 0}' +
    '.msg{font-size:11px;color:#555;margin-top:8px;padding-top:8px;border-top:1px solid #eee}' +
    '.spin{display:inline-block;width:14px;height:14px;border:2px solid rgba(255,255,255,.5);border-top-color:#fff;border-radius:50%;animation:sp .7s linear infinite;vertical-align:middle;margin-right:6px}' +
    '@keyframes sp{to{transform:rotate(360deg)}}</style></head><body>' +
    '<h2>Intransit Hub</h2>' +
    '<button class="btn" id="btn" onclick="go()">Process Next Email</button>' +
    '<div class="card" id="card"></div>' +

    '<div style="margin-top:20px;padding-top:14px;border-top:2px solid #1a3c6d">' +
    '<h2 style="margin:0 0 10px">Fix Claude Drafts</h2>' +
    '<button class="btn" id="clbtn" onclick="checkDrafts()">Check Claude Drafts</button>' +
    '<div class="card" id="clcard"></div>' +
    '</div>' +

    '<script>' +
    'var btn=document.getElementById("btn"),card=document.getElementById("card");' +
    'var clbtn=document.getElementById("clbtn"),clcard=document.getElementById("clcard");' +

    // ── Process Next Email ──
    'function go(){' +
    'btn.disabled=true;btn.innerHTML=\'<span class="spin"></span> Processing...\';card.style.display="none";' +
    'google.script.run.withSuccessHandler(onDone).withFailureHandler(onErr).processNextEmailManual();}' +
    'function onDone(d){try{' +
    'var a=d.action||(d.nothing?"nothing":"?");' +
    'var html=\'<span class="badge \'+a+\'">\'+esc(d.nothing?"Inbox clear":a)+\'</span>\';' +
    'if(d.subject)html+=\'<div class="subj">\'+esc(d.subject)+\'</div>\';' +
    'if(d.from_email)html+=\'<div class="frm">\'+esc(d.from_email)+\'</div>\';' +
    'if(d.draft_preview)html+=\'<div class="draft">\'+esc(d.draft_preview)+\'</div>\';' +
    'if(d.forte_entry){var fe=d.forte_entry;html+=\'<div class="forte">Forte: \'+esc(fe.mpn||"")+\' | \'+esc(String(fe.qty||""))+\' pcs | $\'+esc(String(fe.target_price||""))+\' | \'+esc(fe.country||"")+\'</div>\';}' +
    'if(d.message)html+=\'<div class="msg">\'+esc(d.message)+\'</div>\';' +
    'card.className="card";card.innerHTML=html;card.style.display="block";' +
    '}catch(ex){card.className="card err";card.innerHTML=\'<span class="badge error">Display error</span><div class="msg">\'+esc(ex.message)+\'</div>\';card.style.display="block";}' +
    'finally{btn.disabled=false;btn.innerHTML="Process Next Email";}}' +
    'function onErr(ex){card.className="card err";card.innerHTML=\'<span class="badge error">Error</span><div class="msg">\'+esc(ex.message||String(ex))+\'</div>\';card.style.display="block";btn.disabled=false;btn.innerHTML="Process Next Email";}' +

    // ── Fix Claude Drafts ──
    'var _cdResults=[];' +
    'function checkDrafts(){' +
    'clbtn.disabled=true;clbtn.innerHTML=\'<span class="spin"></span> Checking...\';clcard.style.display="none";' +
    'google.script.run.withSuccessHandler(onDraftsDone).withFailureHandler(onDraftsErr).findAndFixClaudeDrafts();}' +

    'function onDraftsDone(results){' +
    'clbtn.disabled=false;clbtn.innerHTML="Check Claude Drafts";' +
    '_cdResults=results||[];' +
    'if(!_cdResults.length){' +
    'clcard.className="card";clcard.innerHTML=\'<span class="badge no_action">No Claude Drafts</span><div class="msg">All drafts look good.</div>\';clcard.style.display="block";return;}' +
    'var html=\'<span class="badge">\'+_cdResults.length+\' Draft(s) Found</span>\';' +
    '_cdResults.forEach(function(r,i){' +
    'html+=\'<div style="margin-top:12px;padding:10px;background:#f8f8f8;border-radius:4px;border-left:3px solid #4a6fc7">\';' +
    'html+=\'<div class="subj">\'+esc(r.subject)+\'</div>\';' +
    'html+=\'<div class="frm">To: \'+esc(r.to_email)+\'</div>\';' +
    'html+=\'<div class="draft">→ \'+esc(r.corrected_body)+\'</div>\';' +
    'if(r.advice)html+=\'<div class="msg">\'+esc(r.advice)+\'</div>\';' +
    'html+=\'<button class="btn" style="margin-top:8px" id="fixbtn\'+i+\'" onclick="applyFix(\'+i+\')">Apply Fix</button>\';' +
    'html+=\'</div>\';});' +
    'clcard.className="card";clcard.innerHTML=html;clcard.style.display="block";}' +

    'function onDraftsErr(ex){clcard.className="card err";clcard.innerHTML=\'<span class="badge error">Error</span><div class="msg">\'+esc(ex.message||String(ex))+\'</div>\';clcard.style.display="block";clbtn.disabled=false;clbtn.innerHTML="Check Claude Drafts";}' +

    'function applyFix(i){' +
    'var r=_cdResults[i];var fb=document.getElementById("fixbtn"+i);' +
    'fb.disabled=true;fb.innerHTML=\'<span class="spin"></span> Applying...\';' +
    'google.script.run' +
    '.withSuccessHandler(function(){fb.innerHTML="✓ Fixed";fb.style.background="#2a8a2a";})' +
    '.withFailureHandler(function(e){fb.disabled=false;fb.innerHTML="Retry";alert("Error: "+(e.message||String(e)));})' +
    '.applyClaudeDraftFix(r.draft_id,r.thread_id,r.to_email,r.subject,r.corrected_body);}' +

    'function esc(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}' +
    '<\/script></body></html>';
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


function gmailGetThreadMeta_(threadId) {
  var data = gmailREST_('/threads/' + threadId + '?format=METADATA&metadataHeaders=From');
  var msgs = data.messages || [];
  var senders = msgs.map(function(m) {
    var h = ((m.payload || {}).headers || []).filter(function(h) { return h.name === 'From'; });
    return h.length ? h[0].value : '';
  });
  var lastLabelIds = msgs.length ? (msgs[msgs.length - 1].labelIds || []) : [];
  return { senders: senders, labelIds: lastLabelIds };
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
var _gmailLabelCache_ = {};

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
          // Delete ALL matching rows directly — no review emails, handles multiple lots
          var oemSS = SpreadsheetApp.openById(SPREADSHEET_ID);
          var oemSheet = oemSS.getSheetByName(MAIN_SHEET_NAME);
          var oemData = oemSheet.getDataRange().getValues();
          var noStkStamp = 'NO STK ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
          var normTarget = normalize(mpn);
          var toDelete = [];
          for (var oi = oemData.length - 1; oi >= 1; oi--) {
            if (normalize(String(oemData[oi][0] || '')) === normTarget) toDelete.push(oi + 1);
          }
          if (!toDelete.length) {
            hubLog('inventory', 'remove_oem_mpn: not found (already removed?) ' + mpn, { mpn: mpn });
          } else {
            toDelete.forEach(function(row) {
              oemSheet.getRange(row, 5).setValue(noStkStamp);
              oemSheet.deleteRow(row);
            });
            SpreadsheetApp.flush();
            updateForteSheet(mpn);
            hubLog('inventory', 'remove_oem_mpn: deleted ' + toDelete.length + ' row(s) for ' + mpn, { mpn: mpn, rows: toDelete.length });
          }

        } else if (cmd.type === 'add_forte_entry') {
          var mpn = (data.mpn || '').trim();
          var qty = data.qty;
          var tp  = data.tp || '';
          var country = (data.country || '').trim();
          if (!mpn) throw new Error('add_forte_entry: mpn required');
          if (!qty)  throw new Error('add_forte_entry: qty required — cardinal rule');
          var existing = checkForteForMPN(mpn, 60);
          var hasRecent = existing.some(function(r){ return r.recent && r.status.toLowerCase() !== 'closed'; });
          if (hasRecent) {
            hubLog('run', 'add_forte_entry: 60-day skip for ' + mpn, {});
          } else {
            addToForteSheet(mpn, qty, tp, country, '');
            hubLog('inventory', 'add_forte_entry: ' + mpn + ' qty=' + qty + ' tp=' + tp + ' ' + country, { mpn: mpn, qty: qty, tp: tp, country: country });
          }

        } else if (cmd.type === 'delete_draft') {
          var draftId = (data.draft_id || '').trim();
          if (!draftId) throw new Error('No draft_id provided');
          var allDrafts = GmailApp.getDrafts();
          var found = false;
          for (var d = 0; d < allDrafts.length; d++) {
            if (allDrafts[d].getId() === draftId) {
              allDrafts[d].deleteDraft();
              found = true;
              break;
            }
          }
          if (!found) throw new Error('Draft not found: ' + draftId);
          hubLog('drafts', 'Deleted draft ' + draftId, { draft_id: draftId });

        } else if (cmd.type === 'delete_thread_drafts') {
          var threadId = (data.thread_id || '').trim();
          if (!threadId) throw new Error('No thread_id provided');
          var allDrafts = GmailApp.getDrafts();
          var delCount = 0;
          allDrafts.forEach(function(draft) {
            try {
              if (draft.getMessage().getThread().getId() === threadId) {
                draft.deleteDraft();
                delCount++;
              }
            } catch(e) {}
          });
          hubLog('drafts', 'delete_thread_drafts: deleted ' + delCount + ' draft(s) in thread ' + threadId, { thread_id: threadId, count: delCount });

        } else if (cmd.type === 'delete_forte_row') {
          var rowNum = parseInt(data.row, 10);
          var expectedMpn = (data.mpn || '').trim();
          if (!rowNum) throw new Error('No row number provided');
          var fSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
          var actualMpn = String(fSheet.getRange(rowNum, 2).getValue()).trim();
          if (expectedMpn && actualMpn.toUpperCase() !== expectedMpn.toUpperCase()) {
            throw new Error('delete_forte_row safety check failed: row ' + rowNum + ' MPN is "' + actualMpn + '" not "' + expectedMpn + '"');
          }
          fSheet.deleteRow(rowNum);
          hubLog('inventory', 'delete_forte_row: deleted row ' + rowNum + ' (' + actualMpn + ')', { row: rowNum, mpn: actualMpn });

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
          var _ex = oemSS.getSheetByName('_ICS_UPLOAD_TEMP');
          if (_ex) oemSS.deleteSheet(_ex);
          var tempSheet = oemSS.insertSheet('_ICS_UPLOAD_TEMP');
          try {
            var filteredRows = [srcData[0]];
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
              filteredRows.push(cleanRow);
            }
            tempSheet.getRange(1, 1, filteredRows.length, filteredRows[0].length).setValues(filteredRows);
            SpreadsheetApp.flush();
            var gid = tempSheet.getSheetId();
            oemBlob = UrlFetchApp.fetch(
              'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/export?format=xlsx&gid=' + gid,
              fetchOpts
            ).getBlob().setName('OEM_EXCESS.xlsx');
            hubLog('inventory', 'OEM EXCESS clean export: ' + (filteredRows.length - 1) + ' rows sent, ' + skipped + ' blank/null-qty rows skipped', {});
          } finally {
            try { oemSS.deleteSheet(tempSheet); } catch(e) {}
          }

          var inBlob = UrlFetchApp.fetch(
            'https://docs.google.com/spreadsheets/d/' + IN_STOCK_ID + '/export?format=xlsx',
            fetchOpts
          ).getBlob().setName('IN STOCK.xlsx');

          sendPleasePostViaREST(token, oemBlob, inBlob, DATAMASTER_BCC);
          hubLog('inventory', 'Sent NetCOMPONENTS report (OEM_EXCESS + IN STOCK) to ' + DATAMASTER_BCC, {});

        } else if (cmd.type === 'read_sheet_rows') {
          var sheetId   = (data.sheet_id   || '').trim();
          var rangeName = (data.range      || '').trim();
          var sheetName = (data.sheet_name || '').trim();
          if (!sheetId)   throw new Error('read_sheet_rows: no sheet_id');
          if (!rangeName) throw new Error('read_sheet_rows: no range');
          var rss   = SpreadsheetApp.openById(sheetId);
          var rsheet = sheetName ? rss.getSheetByName(sheetName) : rss.getSheets()[0];
          if (!rsheet)    throw new Error('read_sheet_rows: sheet "' + sheetName + '" not found');
          var values = rsheet.getRange(rangeName).getValues();
          hubLog('sheet_data', 'read_sheet_rows: ' + rangeName + ' (' + sheetId + ')', { sheet_id: sheetId, range: rangeName, rows: values });
          Logger.log('read_sheet_rows: ' + values.length + ' rows returned for ' + rangeName);

        } else {
          throw new Error('Unknown command type: ' + cmd.type);
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
            ? buildDraftHTML(String(fix.draft_body), firstMsg)
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
function processNextEmailManual() {
  var out = {success:true,nothing:false,action:null,subject:null,from_email:null,draft_preview:null,forte_entry:null,message:''};
  var noStkKw = ['no stk','no stock','stk sold','stock sold','cant find','cant share','cannot find',
                 'removed','no inventory','sold lying commie','soly lying commie','lying commie',
                 'sold out','all sold','no longer have'];

  // ── 1. David no-stk first ──
  var davidIds = gmailSearchREST('in:inbox from:' + DAVID_EMAIL + ' -label:oem-rfq-incoming-processed', 5);
  for (var di = 0; di < davidIds.length; di++) {
    try {
      var dt = GmailApp.getThreadById(davidIds[di]);
      if (!dt) continue;
      var dm = dt.getMessages()[dt.getMessageCount() - 1];
      var dsub = dm.getSubject();
      var dbody = dm.getPlainBody().toLowerCase().substring(0, 300);
      var isNoStk = noStkKw.some(function(kw){ return dsub.toLowerCase().indexOf(kw) >= 0 || dbody.indexOf(kw) >= 0; });
      if (!isNoStk) { gmailModifyThread_(davidIds[di], ['oem-rfq-incoming-processed'], []); continue; }
      var decision = processThread(dt);
      gmailModifyThread_(davidIds[di], [INCOMING_LABEL, 'oem-rfq-incoming-processed'], []);
      gmailArchiveThread_(davidIds[di]);
      out.action = (decision && decision.action) || 'david_nostock';
      out.subject = dsub;
      out.from_email = DAVID_EMAIL;
      out.draft_preview = (decision && decision.draft_body) || 'Ok, removed from listing.';
      out.message = 'David no-stk processed — archived';
      return out;
    } catch(e2) { continue; }
  }

  // ── 2. Process threads queued by fastScanInbox ──
  var pending = gmailSearchREST('label:' + PENDING_LABEL, 5);
  if (!pending.length) { fastScanInbox(); pending = gmailSearchREST('label:' + PENDING_LABEL, 5); }
  if (pending.length) {
    var tid = pending[0];
    try {
      var t2 = GmailApp.getThreadById(tid);
      if (!t2) { gmailModifyThread_(tid, [], [PENDING_LABEL]); }
      else {
        var msgs2 = t2.getMessages();
        var lastFrom = (msgs2[msgs2.length - 1].getFrom() || '').toLowerCase();
        if (lastFrom.indexOf(JOHN_EMAIL) >= 0 || lastFrom.indexOf('intransittech.com') >= 0) {
          gmailModifyThread_(tid, [], [PENDING_LABEL]);
          out.nothing = true; out.message = 'Skipped (John was last sender) — tap again.'; return out;
        }
        var decision2 = processThread(t2);
        gmailModifyThread_(tid, [], [PENDING_LABEL]);
        out.action = (decision2 && decision2.action) || 'no_action';
        out.subject = t2.getFirstMessageSubject();
        out.from_email = msgs2[0].getFrom();
        out.draft_preview = (decision2 && decision2.draft_body) ? decision2.draft_body.substring(0, 280) : null;
        out.forte_entry = (decision2 && decision2.forte_entry) ? decision2.forte_entry : null;
        out.message = out.action + (decision2 && decision2.mpn ? ' — ' + decision2.mpn : '');
        return out;
      }
    } catch(pe) {
      gmailModifyThread_(tid, [], [PENDING_LABEL]);
      out.action = 'error'; out.message = pe.toString().substring(0, 200); return out;
    }
  }
  out.nothing = true; out.message = 'Inbox clear — nothing to process'; return out;
}


function processPendingThreads() {
  var _cfg = getCachedRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) return;
  var threadIds = gmailSearchREST('label:' + PENDING_LABEL, 10);
  if (!threadIds.length) return;
  hubLog('run', 'processPendingThreads: processing ' + threadIds.length + ' thread(s)');
  threadIds.forEach(function(tid) {
    var t = GmailApp.getThreadById(tid);
    if (!t) { gmailModifyThread_(tid, [], [PENDING_LABEL]); return; }
    var msgs = t.getMessages();
    var lastMsg = msgs[msgs.length - 1];
    if (lastMsg.getFrom().indexOf(JOHN_EMAIL) >= 0) { gmailModifyThread_(tid, [], [PENDING_LABEL]); return; }
    if (lastMsg.getFrom().indexOf('intransittech.com') >= 0) { gmailModifyThread_(tid, [], [PENDING_LABEL]); return; }
    try {
      processThread(t);
      gmailModifyThread_(tid, [], [PENDING_LABEL]); // only remove label on success — keeps retry on failure
    } catch(e) {
      hubLog('error', 'processPendingThreads error (will retry next run): ' + e);
      // PENDING_LABEL stays → processPendingThreads will retry on next trigger fire
    }
  });
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
    // Only include TgtPrice when the netcomp table actually had one.
    // If TgtPrice is null (blank in table), omit it so the worker reads
    // the buyer's TP from later messages (e.g. their reply to our TP request)
    // rather than treating a missing table value as "no TP given."
    var rLine = '[PARSED_RFQ: QtyReq=' + parsedRFQ.qtyReq;
    if (parsedRFQ.tgtPrice !== null) rLine += ', TgtPrice=' + parsedRFQ.tgtPrice;
    rLine += ']';
    content = rLine + '\n' + content;
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
  // Send MPN hint to worker as fallback if Haiku extraction fails.
  // Only send if it looks like a real MPN (letters + digits, no pure-quantity tokens).
  if (mpnHint && /[A-Za-z]/.test(mpnHint) && /[0-9]/.test(mpnHint) && mpnHint.length >= 5
      && !/^\d+(?:pcs?|pc|k|m|units?)?$/i.test(mpnHint)) {
    payload.mpn = mpnHint;
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
    if (cellNorm.length >= 3 && (cellNorm === searchNorm || cellNorm.startsWith(searchNorm) || searchNorm.startsWith(cellNorm))) {
      results.push({ row: i+1, mpn: data[i][0], man: data[i][1], dc: data[i][2], qty: data[i][3], notes: data[i][4] });
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


// Run this NOW to send the Please Post email immediately (bypasses command queue).
function sendPleasePostNow() {
  var token    = ScriptApp.getOAuthToken();
  var fetchOpts = { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true };

  var DATAMASTER_BCC = [
    '5BDFA5@stkdst.com',
    'datamaster@netcomponents.com',
    'post@icsource.com',
    'bill@intransittech.com',
    'david@fortetechno.com',
    'Stan@amorelectronics.com'
  ].join(',');

  var oemBlob;
  var oemSS   = SpreadsheetApp.openById(SPREADSHEET_ID);
  var srcData = oemSS.getSheets()[0].getDataRange().getValues();
  var _existing = oemSS.getSheetByName('_ICS_UPLOAD_TEMP');
  if (_existing) oemSS.deleteSheet(_existing);
  var tempSheet = oemSS.insertSheet('_ICS_UPLOAD_TEMP');
  try {
    var filteredRows = [srcData[0]];
    var skipped = 0;
    for (var i = 1; i < srcData.length; i++) {
      var mpn    = String(srcData[i][0]).trim();
      if (!mpn) { skipped++; continue; }
      var qtyRaw = srcData[i][3];
      var qtyNum = (typeof qtyRaw === 'number') ? qtyRaw : parseFloat(String(qtyRaw).replace(/,/g, ''));
      if (isNaN(qtyNum) || qtyNum <= 0) { skipped++; continue; }
      var row = srcData[i].slice(); row[0] = mpn; row[3] = qtyNum;
      filteredRows.push(row);
    }
    tempSheet.getRange(1, 1, filteredRows.length, filteredRows[0].length).setValues(filteredRows);
    SpreadsheetApp.flush();
    oemBlob = UrlFetchApp.fetch(
      'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/export?format=xlsx&gid=' + tempSheet.getSheetId(),
      fetchOpts
    ).getBlob().setName('OEM_EXCESS.xlsx');
    Logger.log('OEM EXCESS: ' + (filteredRows.length - 1) + ' rows exported, ' + skipped + ' skipped');
  } finally {
    try { oemSS.deleteSheet(tempSheet); } catch(e) {}
  }

  var inBlob = UrlFetchApp.fetch(
    'https://docs.google.com/spreadsheets/d/' + IN_STOCK_ID + '/export?format=xlsx',
    fetchOpts
  ).getBlob().setName('IN STOCK.xlsx');

  sendPleasePostViaREST(token, oemBlob, inBlob, DATAMASTER_BCC);
  Logger.log('sendPleasePostNow: DONE');
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




function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){ScriptApp.deleteTrigger(t);});
  // Two-phase architecture: fastScanInbox labels threads only (< 10s/run),
  // processPendingThreads calls Claude API on a slower 5-min schedule.
  // This prevents runEmailScan from exhausting the 6-hr daily execution quota.
  ScriptApp.newTrigger('fastScanInbox').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('processPendingThreads').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('checkDavidNoStockEmails').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('checkBillNetcompRemovals').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('checkInboxForPaymentAdvice').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('processFixQueue').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('processCommandQueue').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('sendDailyCostReport').timeBased().atHour(8).everyDays(1).create();
  Logger.log('8 triggers installed.');
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


