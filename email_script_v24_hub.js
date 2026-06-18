// ============================================================
// COMPLETE SCRIPT v24 + HUB LOGGING — Replace ALL existing code with this
// OEM EXCESS Automation — John Fluman / Intransit Technologies
// Hub logging added: hubLog() tracks all runs, hubPostDraft() feeds Email Triage tab
// ============================================================

var SPREADSHEET_ID    = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
var MAIN_SHEET_NAME   = 'sheet1';
var DELETED_SHEET_NAME = 'Deleted Rows';
var NOTIFY_EMAIL      = 'john.fluman@intransittech.com';
var JOHN_EMAIL        = 'john.fluman@intransittech.com';
var DAVID_EMAIL       = 'david@fortetechno.com';
var BILL_EMAIL        = 'bill.pratt@intransittech.com';
var DEB_EMAIL         = 'deb@intransittech.com';
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

// ============================================================
// INTRANSIT HUB — Logging & Draft Sync
// ============================================================
var HUB_URL    = 'https://intransit-hub.intransit-sales.workers.dev';
var HUB_SECRET = 'InTransit!Hub#2026';

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

function hubPostDraft(threadId, mpn, sender, subject, draftContent) {
  try {
    UrlFetchApp.fetch(HUB_URL + '/api/drafts', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + HUB_SECRET },
      payload: JSON.stringify({ thread_id: threadId, mpn: mpn, sender: sender,
                                subject: subject, draft_content: draftContent }),
      muteHttpExceptions: true,
    });
  } catch(e) { Logger.log('hubPostDraft error: ' + e); }
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

// ============================================================
// GMAIL REST API — PROPER THREADED DRAFT CREATION
// ============================================================

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

// ============================================================
// SIGNATURE + HTML BUILDER
// ============================================================

function getSignatureHTML() {
  return '<br><br><div><b><span style="color:rgb(31,73,125);font-family:Tahoma,sans-serif;font-size:10pt">Regards,</span></b></div>'
    + '<div><b><span style="color:rgb(31,73,125);font-family:Tahoma,sans-serif;font-size:10pt">John Fluman</span></b></div>'
    + '<div><b><span style="color:rgb(31,73,125);font-family:Arial,sans-serif;font-size:8pt">Intransit Technologies</span></b></div>'
    + '<div><a href="mailto:john.fluman@intransittech.com" style="font-family:Calibri;font-size:8pt">john.fluman@intransittech.com</a></div>'
    + '<div><i><span style="color:gray;font-family:Arial,sans-serif;font-size:7.5pt">An ISO 9001 Certified Company</span></i></div>'
    + '<div><span style="color:rgb(31,73,125);font-family:Tahoma,sans-serif;font-size:8pt">Toll (877) 677-5868 x101 - Local (949) 481-7935 x101</span></div>'
    + '<br><div><span style="color:rgb(166,166,166);font-family:Calibri,sans-serif;font-size:8pt">The information contained in this communication and its attachment(s) is intended only for the use of the individual to whom it is addressed and may contain information that is privileged, confidential, or exempt from disclosure. If the reader of this message is not the intended recipient, you are hereby notified that any dissemination, distribution, or copying of this communication is strictly prohibited. If you have received this communication in error, please notify <a href="mailto:john.fluman@intransittech.com" style="font-family:Calibri;font-size:8pt">john.fluman@intransittech.com</a> and delete the communication without retaining any copies. Thank you.</span></div>';
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

// ============================================================
// UTILITY
// ============================================================

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

// ============================================================
// EXTRACT HELPERS
// ============================================================

function extractMPN(subject) {
  if (!subject) return null;
  var clean = subject.replace(/^(Re:|Fwd:|FW:|RE:|FWD:|\[EXTERNAL\]|Subject:)\s*/gi, '').trim();
  var netcompMatch = clean.match(/\|\s*([^|)]+)\)\s*$/);
  if (netcompMatch) return netcompMatch[1].trim();
  var icsMatch = clean.match(/::\s*(\S+)/);
  if (icsMatch) return icsMatch[1].trim();
  clean = clean.replace(/^RFQ#?\s*/i, '').trim();
  var wordColonMpnMatch = clean.match(/(?:^|\s)\w+:([A-Z][A-Z0-9\-\.\/]{3,})/i);
  if (wordColonMpnMatch && /[A-Za-z]/.test(wordColonMpnMatch[1]) && /[0-9]/.test(wordColonMpnMatch[1])) {
    return wordColonMpnMatch[1].trim();
  }
  var stopwords = ['no','stk','stock','removed','remove','out','of','npi','the','a','an','is','has','for','from','please','and','or','not','new','update','cant','share','rfq','quote','quotation','request','inquiry','inquire','netcomponents','member','price','target','pcs','qty','quantity','external','ics','source','on','standard','subject','requirements'];
  var tokens = clean.split(/\s+/);
  for (var i = 0; i < tokens.length; i++) {
    var token = tokens[i].replace(/[,;:?()[\]]/g, '');
    token = token.replace(/:[A-Z0-9]{1,5}$/, '');
    if (stopwords.indexOf(token.toLowerCase()) >= 0) continue;
    if (/^#\d+$/.test(token)) continue;
    if (/^[A-Za-z]{1,4}\d{5,}$/.test(token)) continue;
    if (token.length < 3) continue;
    return token;
  }
  var icsOnMatch = clean.match(/\bRFQ\s+on\s+(\S+)\s+from\b/i);
  if (icsOnMatch) return icsOnMatch[1].trim();
  return null;
}

function extractMPNFromBody(body) {
  if (!body) return null;
  var match = body.match(/Removed\s*-\s*MPN:\s*([^\s\n\r]+)/i);
  return match ? match[1].trim() : null;
}

function extractMPNFromRFQBody(body) {
  if (!body) return null;
  var labelPatterns = [
    /(?:part\s*#|part\s+number|p\/n|mpn)\s*:?\s*([A-Z0-9][A-Z0-9\-\/\.]{3,})/i
  ];
  for (var i = 0; i < labelPatterns.length; i++) {
    var m = body.match(labelPatterns[i]);
    if (m && m[1]) return m[1].trim();
  }
  var commonWords = ['please','quote','quotation','stock','thank','regards','best','hello',
    'john','need','your','following','item','molex','intel','texas','instruments','mfr',
    'dear','attached','regards','sincerely','each','pieces','quantity'];
  var lines = body.split('\n');
  for (var l = 0; l < lines.length; l++) {
    var line = lines[l].trim();
    if (line.charAt(0) === '>') continue;
    var tokens = line.split(/\s+/);
    for (var t = 0; t < tokens.length; t++) {
      var token = tokens[t].replace(/[,;:?()[\]\.\*：，]/g, '').trim();
      if (token.length < 5) continue;
      if (commonWords.indexOf(token.toLowerCase()) >= 0) continue;
      if (/^\d+(?:pcs?|pieces?|units?|ea|each)$/i.test(token)) continue;
      if (/[A-Za-z]/.test(token) && /[0-9]/.test(token)) return token;
    }
  }
  return null;
}

function extractNetcompsQtyReq(htmlBody) {
  if (!htmlBody) return null;
  var tableMatch = htmlBody.match(/class="partlist"[\s\S]*?<\/tr>([\s\S]*?)<\/table>/i);
  if (!tableMatch) return null;
  var cells = tableMatch[1].match(/<td[^>]*>([\s\S]*?)<\/td>/gi);
  if (!cells || cells.length < 4) return null;
  var qtyCell = cells[3].replace(/<[^>]+>/g, '').trim();
  if (!qtyCell) return null;
  var qty = parseInt(qtyCell.replace(/[^0-9]/g, ''), 10);
  return qty > 0 ? String(qty) : null;
}

function extractNetcompsTgtPrice(plainBody, htmlBody) {
  if (htmlBody) {
    var tableMatch = htmlBody.match(/class="partlist"[\s\S]*?<\/tr>([\s\S]*?)<\/table>/i);
    if (tableMatch) {
      var cells = tableMatch[1].match(/<td[^>]*>([\s\S]*?)<\/td>/gi);
      if (cells && cells.length >= 5) {
        var tpCell = cells[4].replace(/<[^>]+>/g, '').trim();
        if (tpCell) {
          var price = parseFloat(tpCell.replace(/[^0-9.]/g, ''));
          var qtyReqVal = parseFloat((cells[3] || '').replace(/<[^>]+>/g, '').replace(/[^0-9]/g, ''));
          if (price > 0 && price < 100000 && price !== qtyReqVal) return '$' + price;
        }
        return null;
      }
    }
  }
  if (plainBody) {
    var m2 = plainBody.match(/\bTP[:\s]+\$?(\d+(?:\.\d+)?)\s*(?:usd|USD)?/i);
    if (m2 && parseFloat(m2[1]) > 0) return '$' + parseFloat(m2[1]);
  }
  return null;
}

function extractTargetPrice(text) {
  if (!text) return null;
  var rangeMatch = text.match(/\$?\s*(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)\s*(?:\$|usd|\/each)?/i);
  if (rangeMatch) {
    var lo = parseFloat(rangeMatch[1]), hi = parseFloat(rangeMatch[2]);
    var hasCurrency = rangeMatch[0].indexOf('$') >= 0 || /usd|each/i.test(rangeMatch[0]);
    var hasDecimal = rangeMatch[1].indexOf('.') >= 0 || rangeMatch[2].indexOf('.') >= 0;
    if (hi > lo && hi < 100000 && (hasCurrency || hasDecimal)) return '$' + hi;
  }
  var patterns = [
    /\btp\s*([\d,.]+)/i,
    /\btarget[:\s]+\$?\s*([\d,.]+)/i,
    /USD\s*([\d,.]+)/i,
    /([\d,.]+)\s*USD/i,
    /\b(\d+(?:\.\d+)?)U\b/i,
    /([\d,.]+)\s*\/?\s*(?:each|ea)[.,?!]?(?:\s|$)/i,
    /around\s+\$?\s*([\d,.]+)/i,
    /\bprice:\s*\$?\s*([\d,.]+)/i,
    /\$\s*([\d,]*\.\d+|[\d,]+)(?!\s*(?:min|minimum|500|2000))/i,
    /^\s*(\.\d+)\b/,
    /^\s*(\d+\.\d+)\s*(?:\r?\n|$)/
  ];
  for (var i = 0; i < patterns.length; i++) {
    var m = text.match(patterns[i]);
    if (m && m[1]) {
      var val = parseFloat(m[1].replace(/,/g, ''));
      if (val > 0) return '$' + val;
    }
  }
  return null;
}

function parseQtyValue(rawQty, country) {
  var s = String(rawQty || '').trim().replace(/\s+/g, '');
  var euCountries = ['Germany', 'Netherlands', 'France', 'Denmark', 'Sweden', 'Norway', 'Finland', 'Belgium', 'Austria'];
  if (euCountries.indexOf(country) >= 0 && /^\d{1,3}(\.\d{3})+$/.test(s)) {
    return parseInt(s.replace(/\./g, ''), 10);
  }
  return parseFloat(s.replace(/,/g, '')) || 0;
}

function extractCountryFromEmail(from) {
  if (!from) return 'USA';
  var lower = from.toLowerCase();
  if (lower.indexOf('.cn') >= 0 || lower.indexOf('shenzhen') >= 0 || lower.indexOf('china') >= 0) return 'China';
  if (lower.indexOf('.hk') >= 0 || lower.indexOf('hong kong') >= 0) return 'Hong Kong';
  if (lower.indexOf('.il') >= 0 || lower.indexOf('israel') >= 0) return 'Israel';
  if (lower.indexOf('.de') >= 0 || lower.indexOf('germany') >= 0) return 'Germany';
  if (lower.indexOf('.fr') >= 0 || lower.indexOf('france') >= 0) return 'France';
  if (lower.indexOf('.ca') >= 0) return 'Canada';
  if (lower.indexOf('.tw') >= 0 || lower.indexOf('taiwan') >= 0) return 'Taiwan';
  if (lower.indexOf('.kr') >= 0 || lower.indexOf('korea') >= 0) return 'Korea';
  if (lower.indexOf('.dk') >= 0 || lower.indexOf('denmark') >= 0) return 'Denmark';
  if (lower.indexOf('.nl') >= 0) return 'Netherlands';
  if (lower.indexOf('.co.uk') >= 0 || lower.indexOf('.uk') >= 0) return 'UK';
  if (lower.indexOf('.in') >= 0 || lower.indexOf('india') >= 0) return 'India';
  if (lower.indexOf('.sg') >= 0 || lower.indexOf('singapore') >= 0) return 'Singapore';
  if (lower.indexOf('.sk') >= 0 || lower.indexOf('slovakia') >= 0) return 'Slovakia';
  if (lower.indexOf('.cz') >= 0 || lower.indexOf('czech') >= 0) return 'Czech Republic';
  if (lower.indexOf('.pl') >= 0 || lower.indexOf('poland') >= 0) return 'Poland';
  if (lower.indexOf('.eu') >= 0) return 'EU';
  return 'USA';
}

function extractBuyerEmail(fromRaw) {
  if (!fromRaw) return '';
  var m = fromRaw.match(/<([^>]+)>/);
  return m ? m[1].trim() : fromRaw.trim();
}

function getPriorQuoteHistory(mpn) {
  if (!mpn) return '';
  try {
    var threads = GmailApp.search('in:sent "' + mpn + '"', 0, 20);
    var quotes = [];
    for (var i = 0; i < threads.length && quotes.length < 3; i++) {
      var messages = threads[i].getMessages();
      for (var j = messages.length - 1; j >= 0 && quotes.length < 3; j--) {
        var msg = messages[j];
        if (msg.getFrom().indexOf(JOHN_EMAIL) < 0) continue;
        var body = stripQuotedLines(msg.getPlainBody());
        var priceMatch = body.match(/\$\s*(\d+(?:\.\d+)?)\s*(?:each|\/ea)/i)
          || body.match(/(\d+(?:\.\d+)?)\s*(?:each|per\s*unit)/i)
          || body.match(/price[:\s]+\$?\s*(\d+(?:\.\d+)?)/i);
        if (priceMatch && parseFloat(priceMatch[1]) > 0) {
          var dateStr = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), 'MMM d, yyyy');
          quotes.push('- ' + dateStr + ': $' + parseFloat(priceMatch[1]).toFixed(2) + ' each');
        }
      }
    }
    if (!quotes.length) return '';
    return '\n\n---\n[JOHN - PRIOR QUOTES - DELETE BEFORE SENDING:\n' + quotes.join('\n') + ']';
  } catch(e) {
    Logger.log('getPriorQuoteHistory error: ' + e.toString());
    return '';
  }
}

function stripQuotedLines(text) {
  if (!text) return '';
  var lines = text.split('\n');
  var result = [];
  for (var i = 0; i < lines.length; i++) {
    var trimmed = lines[i].trim();
    if (trimmed.charAt(0) === '>') continue;
    if (/^From:\s/i.test(trimmed) || /^De\s*:\s/i.test(trimmed) || /^-{3,}\s*Original Message/i.test(trimmed) || /^On .+ wrote:/i.test(trimmed)) break;
    result.push(lines[i]);
  }
  return result.join('\n');
}

function isNetcompEmail(from, subject) {
  if (from && from.toLowerCase().indexOf('netcomponents') >= 0) return true;
  if (subject && subject.indexOf('netCOMPONENTS') >= 0) return true;
  return false;
}

function isICSouceEmail(from, subject) {
  if (from && from.toLowerCase().indexOf('icsource.com') >= 0) return true;
  if (subject && subject.indexOf('IC Source RFQ') >= 0) return true;
  return false;
}

function extractICSourcBuyerEmail(htmlBody) {
  if (!htmlBody) return null;
  // Primary: mailto: links, excluding icsource.com and intransittech.com
  var matches = htmlBody.match(/href=["']mailto:([^"']+)["']/gi);
  if (matches) {
    for (var i = 0; i < matches.length; i++) {
      var m = matches[i].match(/href=["']mailto:([^"']+)["']/i);
      if (m && m[1]) {
        var addr = m[1].trim().toLowerCase();
        if (addr.indexOf('intransittech') < 0 && addr.indexOf('icsource.com') < 0) return m[1].trim();
      }
    }
  }
  // Fallback: scan plain-text email addresses in the body
  var emailMatches = htmlBody.match(/\b([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})\b/g);
  if (emailMatches) {
    for (var j = 0; j < emailMatches.length; j++) {
      var e = emailMatches[j].trim().toLowerCase();
      if (e.indexOf('intransittech') < 0 && e.indexOf('icsource.com') < 0) return emailMatches[j].trim();
    }
  }
  return null;
}

function extractNetcompsBuyerEmail(body) {
  if (!body) return null;
  var m = body.match(/RFQ From:.*?\(([^)]+@[^)]+)\)/i);
  if (!m) return null;
  var email = m[1].replace(/<[^>]*>/g, '').trim();
  return email || null;
}

// ============================================================
// SEARCHES
// ============================================================

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

function addToForteSheet(mpn, qty, targetPrice, country, historyNote) {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
  var nextRow = sheet.getLastRow() + 1;
  var potentialFormula = '=C' + nextRow + '*D' + nextRow;
  sheet.appendRow([today, mpn, qty||'', targetPrice||'', '', country||'', potentialFormula, '', '', historyNote||'', 'Open']);
  Logger.log('Added to Forte: ' + mpn);
}

function updateForteSheet(mpn) {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
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

// ============================================================
// ADD TO STAN'S SHEET
// ============================================================
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

// ============================================================
// OEM EXCESS DELETE
// ============================================================

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
  if (exact.length===1){logDeletion(deletedSheet,exact[0].data,emailSubject);mainSheet.deleteRow(exact[0].row);return 'DELETED';}
  if (exact.length>1){sendReviewEmail(partNumber,emailSubject,exact);return 'MULTIPLE';}
  if (!exact.length&&fuzzy.length===1&&fuzzy[0].type==='stripped'){logDeletion(deletedSheet,fuzzy[0].data,emailSubject);mainSheet.deleteRow(fuzzy[0].row);return 'FUZZY';}
  if (fuzzy.length){sendReviewEmail(partNumber,emailSubject,fuzzy);return 'FUZZY_REVIEW';}
  sendReviewEmail(partNumber,emailSubject,[]);return 'NOT_FOUND';
}

// ============================================================
// TRIGGER 1 — David no-stock emails
// ============================================================
function checkDavidNoStockEmails() {
  var _cfg = getRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) { hubLog('run', 'checkDavidNoStockEmails: disabled via hub config'); return; }
  var query = 'from:'+DAVID_EMAIL+' (subject:"no stk" OR subject:"no stock" OR subject:"removed" OR subject:"stock sold" OR "removed" OR "cant share") -label:'+INCOMING_LABEL+' in:inbox';
  var threads = GmailApp.search(query,0,20);
  hubLog('run', 'checkDavidNoStockEmails: ' + threads.length + ' thread(s)');
  if (!threads.length) return;
  var label = GmailApp.getUserLabelByName(INCOMING_LABEL)||GmailApp.createLabel(INCOMING_LABEL);
  threads.forEach(function(thread) {
    var msg = thread.getMessages()[thread.getMessageCount()-1];
    var subject = msg.getSubject();
    var mpn = extractMPN(subject);
    if (mpn) {
      var davidBody = '<div dir="ltr">Removed - MPN: '+mpn+getSignatureHTML()+'</div>';
      createThreadedDraft(DAVID_EMAIL,'Re: '+subject,davidBody,msg.getId(),thread.getId(),null);
      hubLog('draft_created', 'David removal draft: ' + mpn, {mpn: mpn});
      Logger.log('David draft created: ' + mpn);
    } else {
      GmailApp.sendEmail(NOTIFY_EMAIL,'OEM EXCESS: Could not identify MPN',
        'Subject: "'+subject+'"\nDate: '+msg.getDate()+'\nhttps://mail.google.com/mail/u/0/#inbox/'+thread.getId()+'\n\nReply to David: "Removed - MPN: [part number]"');
    }
    thread.addLabel(label);
  });
}

// ============================================================
// TRIGGER 2 — Sent "Removed - MPN:" → delete from OEM EXCESS
// ============================================================
function checkSentRemovals() {
  var _cfg = getRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) { hubLog('run', 'checkSentRemovals: disabled via hub config'); return; }
  var query = 'in:sent to:'+DAVID_EMAIL+' "Removed - MPN:" -label:oem-removal-processed';
  var threads = GmailApp.search(query,0,20);
  hubLog('run', 'checkSentRemovals: ' + threads.length + ' thread(s)');
  if (!threads.length) return;
  var doneLabel = GmailApp.getUserLabelByName('oem-removal-processed')||GmailApp.createLabel('oem-removal-processed');
  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      if (msg.getFrom().indexOf(JOHN_EMAIL)>=0) {
        var mpn = extractMPNFromBody(msg.getPlainBody());
        if (mpn) {
          deletePart(mpn,msg.getSubject());
          updateForteSheet(mpn);
          hubLog('run', 'Deleted from OEM EXCESS: ' + mpn, {mpn: mpn});
        }
      }
    });
    thread.addLabel(doneLabel);
  });
}

// ============================================================
// TRIGGER 3 — Inbox TP replies → threaded draft
// ============================================================
function checkInboxForTPReplies() {
  var _cfg = getRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) { hubLog('run', 'checkInboxForTPReplies: disabled via hub config'); return; }
  var query = 'in:inbox "minimum line requirement" -label:oem-tp-processed';
  var threads = GmailApp.search(query,0,20);
  hubLog('run', 'checkInboxForTPReplies: ' + threads.length + ' thread(s)');
  if (!threads.length) return;
  var label = GmailApp.getUserLabelByName('oem-tp-processed')||GmailApp.createLabel('oem-tp-processed');
  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    var lastMsg = messages[messages.length-1];
    if (lastMsg.getFrom().indexOf(JOHN_EMAIL)>=0){return;}
    var lastBuyerDate = null;
    var johnAlreadyReplied = false;
    for (var mi = 0; mi < messages.length; mi++) {
      var mi_msg = messages[mi];
      if (mi_msg.getFrom().indexOf(JOHN_EMAIL) < 0 && mi_msg.getFrom().indexOf('intransittech') < 0) {
        lastBuyerDate = mi_msg.getDate(); johnAlreadyReplied = false;
      } else if (mi_msg.getFrom().indexOf(JOHN_EMAIL) >= 0 && lastBuyerDate) {
        johnAlreadyReplied = true;
      }
    }
    if (johnAlreadyReplied){thread.addLabel(label);return;}
    var body = lastMsg.getPlainBody();
    var tp = extractTargetPrice(stripQuotedLines(body));
    if (!tp){thread.addLabel(label);return;}
    var mpn = extractMPN(thread.getFirstMessageSubject());
    if (!mpn || !/[0-9\-]/.test(mpn)) {
      var tpMsgs = thread.getMessages();
      for (var ti = 0; ti < tpMsgs.length; ti++) {
        if (tpMsgs[ti].getFrom().indexOf(JOHN_EMAIL) < 0 && tpMsgs[ti].getFrom().indexOf('intransittech') < 0) {
          var bodyMpnTP = extractMPNFromRFQBody(tpMsgs[ti].getPlainBody());
          if (bodyMpnTP) { mpn = bodyMpnTP.replace(/#\w+$/, ''); break; }
        }
      }
    }
    if (!mpn){thread.addLabel(label);return;}
    var buyerEmail = extractBuyerEmail(lastMsg.getFrom());
    var oemResults = searchOEMExcess(mpn);
    if (!oemResults.length){thread.addLabel(label);return;}
    var hasBillExt = oemResults.some(function(r){return r.notes.indexOf('BILL EXT 117')>=0;})
      || body.indexOf('BILL EXT 117') >= 0;
    var firstMsg = messages[0];
    var netcompTP = isNetcompEmail(firstMsg.getFrom(), firstMsg.getSubject());
    var icsourceTP = isICSouceEmail(firstMsg.getFrom(), firstMsg.getSubject());
    var replyTo = buyerEmail;
    if (netcompTP) replyTo = extractNetcompsBuyerEmail(body) || buyerEmail;
    else if (icsourceTP) replyTo = extractICSourcBuyerEmail(lastMsg.getBody()) || buyerEmail;
    var origMsg = null;
    for (var i=0;i<messages.length;i++){
      if (messages[i].getFrom().indexOf(JOHN_EMAIL)<0&&messages[i].getFrom().indexOf('intransittech')<0){origMsg=messages[i];break;}
    }
    var replyText = hasBillExt ? MSG_BILL : MSG_CHECKING;
    var htmlBody = origMsg ? buildDraftHTML(replyText,origMsg) : '<div dir="ltr">'+replyText+getSignatureHTML()+'</div>';
    var subj = 'Re: '+thread.getFirstMessageSubject();
    var draftId = hasBillExt
      ? createThreadedDraft(replyTo,subj,htmlBody,lastMsg.getId(),thread.getId(),BILL_EMAIL)
      : createThreadedDraft(replyTo,subj,htmlBody,lastMsg.getId(),thread.getId(),null);
    if (!draftId) { Logger.log('TP draft FAILED (no draftId) — skipping label: '+mpn+' | replyTo: '+replyTo); return; }
    hubPostDraft(thread.getId(), mpn, replyTo, subj, replyText);
    hubLog('draft_created', 'TP reply: ' + mpn + ' | TP: ' + tp, {mpn: mpn, tp: tp, replyTo: replyTo});
    Logger.log('TP reply draft: '+mpn+' | TP: '+tp);
    thread.addLabel(label);
  });
}

// ============================================================
// TRIGGER 4 — New inbox RFQs → threaded draft
// ============================================================
function checkInboxForNewRFQs() {
  var _cfg = getRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) { hubLog('run', 'checkInboxForNewRFQs: disabled via hub config'); return; }
  var query = 'in:inbox (to:rfq@intransittech.com OR deliveredto:rfq@intransittech.com OR subject:rfq OR subject:"please quote" OR subject:"request for quote" OR subject:"request for quotation" OR ((to:john.fluman@intransittech.com OR deliveredto:john.fluman@intransittech.com) ("quotation" OR "best price" OR "net components" OR "netcomponents" OR "netcomp" OR "looking for" OR "quote your stock"))) -from:intransittech.com -from:partalert@netcomponents.com -label:oem-rfq-incoming-processed';
  var threads = GmailApp.search(query,0,10);
  hubLog('run', 'checkInboxForNewRFQs: ' + threads.length + ' thread(s)');
  if (!threads.length) return;
  var label = GmailApp.getUserLabelByName('oem-rfq-incoming-processed')||GmailApp.createLabel('oem-rfq-incoming-processed');
  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    var johnReplied = messages.some(function(m){return m.getFrom().indexOf(JOHN_EMAIL)>=0;});
    if (johnReplied){thread.addLabel(label);return;}
    var lastMsg = messages[messages.length-1];
    if (lastMsg.getFrom().indexOf('intransittech.com')>=0){thread.addLabel(label);return;}
    var subject = thread.getFirstMessageSubject();
    var mpn = extractMPN(subject);
    var fullBody = lastMsg.getPlainBody();
    if (!fullBody || fullBody.trim().length < 30) {
      fullBody = lastMsg.getBody().replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
    }
    if (!mpn || !/[A-Za-z]/.test(mpn) || !/[0-9]/.test(mpn)) {
      var bodyMpn = extractMPNFromRFQBody(fullBody);
      if (bodyMpn) mpn = bodyMpn.replace(/#\w+$/, '');
    }
    if (!mpn){thread.addLabel(label);return;}
    var htmlBody = lastMsg.getBody();

    var buyerEmail = extractBuyerEmail(lastMsg.getFrom());
    var firstMsg = messages[0];
    var netcomp = isNetcompEmail(firstMsg.getFrom(), firstMsg.getSubject());
    var icsource = isICSouceEmail(firstMsg.getFrom(), firstMsg.getSubject());
    var replyTo = buyerEmail;
    if (netcomp) replyTo = extractNetcompsBuyerEmail(fullBody) || buyerEmail;
    else if (icsource) replyTo = extractICSourcBuyerEmail(htmlBody) || buyerEmail;
    var country = extractCountryFromEmail(replyTo);

    var tp = netcomp ? extractNetcompsTgtPrice(fullBody, htmlBody) : null;
    if (!tp) {
      var bodyBeforeListing = fullBody.split('OEM EXCESS')[0];
      tp = extractTargetPrice(bodyBeforeListing) || extractTargetPrice(subject.split('|')[0]);
    }

    var qty = '';
    var qtyMatch = fullBody.match(/(\d[\d\s,.]*\d|\d)\s*(?:pcs?|pieces?|units?|ea|each)|QtyReq\s+(\d+)|Qty\s*[：:]\s*(\d[\d,]*)/i);
    if (qtyMatch) {
      var rawQty = qtyMatch[1] || qtyMatch[2] || qtyMatch[3] || '';
      var qtyNum = parseQtyValue(rawQty, country);
      if (qtyNum > 0) qty = String(qtyNum);
    }
    if (!qty && netcomp) {
      var netcompsQty = extractNetcompsQtyReq(htmlBody);
      if (netcompsQty) qty = netcompsQty;
    }
    if (!qty) {
      var kMatch = fullBody.match(/\b(\d+(?:\.\d+)?)\s*[Kk]\b/);
      if (kMatch) qty = String(Math.round(parseFloat(kMatch[1]) * 1000));
    }

    var oemResults = searchOEMExcess(mpn);
    var inStockResults = searchInStock(mpn);
    var stanResults = searchStanSheet(mpn);

    var origMsg = null;
    for (var i=0;i<messages.length;i++){
      if (messages[i].getFrom().indexOf(JOHN_EMAIL)<0&&messages[i].getFrom().indexOf('intransittech')<0){origMsg=messages[i];break;}
    }
    var replyToId = lastMsg.getId(), threadId = thread.getId();
    var handled = false;

    var hasBillExt = oemResults.length > 0 && (
      oemResults.some(function(r){return r.notes.indexOf('BILL EXT 117')>=0;})
      || fullBody.indexOf('BILL EXT 117') >= 0
    );
    var has2k = oemResults.length > 0 && oemResults.some(function(r){return r.notes.indexOf('$2000')>=0;});
    var tpMsg = (has2k && !hasBillExt) ? MSG_NEED_TP_2000 : MSG_NEED_TP_500;

    // IN STOCK branch
    if (inStockResults.length) {
      var ownStock = inStockResults.filter(function(r){return String(r.notes).indexOf('Warehouse#3')<0;});
      var w3Stock  = inStockResults.filter(function(r){return String(r.notes).indexOf('Warehouse#3')>=0;});
      var isW3 = w3Stock.length > 0;
      var w3Mpn = isW3 ? String(w3Stock[0].mpn) : mpn;
      var w3Quoted = isW3 && stanResults.length > 0 && stanResults[0].status === 'QUOTED';

      if (ownStock.length && isW3) {
        addToStanSheet(w3Mpn, country, qty, tp);
        Logger.log('Both own stock and W3 for '+mpn+' — W3 added to Stan sheet, leaving in inbox for John');

      } else if (ownStock.length) {
        var sr = ownStock[0];
        var priorQuotes = getPriorQuoteHistory(sr.mpn);
        var stockInfo = 'This is our stock<br><br>'
          + 'MPN: ' + sr.mpn + '<br>'
          + (sr.dc ? 'DC: ' + sr.dc + '<br>' : '')
          + (sr.qty ? 'QTY in stock: ' + sr.qty + '<br>' : '')
          + 'Price: $[FILL IN]<br><br>'
          + 'There is a $100 min on stock items'
          + priorQuotes;
        var hbStock = origMsg?buildDraftHTML(stockInfo,origMsg):'<div dir="ltr">'+stockInfo+getSignatureHTML()+'</div>';
        createThreadedDraft(replyTo,'Re: '+subject,hbStock,replyToId,threadId,null);
        hubPostDraft(threadId, mpn, replyTo, 'Re: '+subject, stockInfo);
        hubLog('draft_created', 'In-stock draft: '+mpn, {mpn:mpn, type:'own_stock', qty:sr.qty});
        Logger.log('IN STOCK own draft: '+mpn);
        handled = true;

      } else if (isW3 && w3Quoted) {
        var stanR = stanResults[0];
        var stanQuoteText = stanR.colB + (stanR.colC ? ' | ' + stanR.colC : '');
        var replyTextW3, htmlBodyW3;
        if (oemResults.length) {
          replyTextW3 = stanQuoteText + '\n\nWe also have additional quantity available. ' + tpMsg;
          Logger.log('Combined Stan+OEM draft: '+mpn);
        } else {
          replyTextW3 = stanQuoteText;
          Logger.log('Stan QUOTED draft: '+mpn);
        }
        htmlBodyW3 = origMsg?buildDraftHTML(replyTextW3,origMsg):'<div dir="ltr">'+replyTextW3+getSignatureHTML()+'</div>';
        createThreadedDraft(replyTo,'Re: '+subject,htmlBodyW3,replyToId,threadId,null);
        hubPostDraft(threadId, mpn, replyTo, 'Re: '+subject, replyTextW3);
        hubLog('draft_created', 'Stan quoted draft: '+mpn, {mpn:mpn, type:'stan_quoted'});
        handled = true;

      } else if (isW3 && !w3Quoted) {
        addToStanSheet(w3Mpn, country, qty, tp);
        if (oemResults.length) {
          var oemReplyText, oemHtmlBody;
          if (tp) {
            oemReplyText = hasBillExt ? MSG_BILL : MSG_CHECKING;
            oemHtmlBody = origMsg?buildDraftHTML(oemReplyText,origMsg):'<div dir="ltr">'+oemReplyText+getSignatureHTML()+'</div>';
            if (hasBillExt) createThreadedDraft(replyTo,'Re: '+subject,oemHtmlBody,replyToId,threadId,BILL_EMAIL);
            else createThreadedDraft(replyTo,'Re: '+subject,oemHtmlBody,replyToId,threadId,null);
          } else {
            oemReplyText = tpMsg;
            oemHtmlBody = origMsg?buildDraftHTML(oemReplyText,origMsg):'<div dir="ltr">'+oemReplyText+getSignatureHTML()+'</div>';
            createThreadedDraft(replyTo,'Re: '+subject,oemHtmlBody,replyToId,threadId,null);
          }
          hubPostDraft(threadId, mpn, replyTo, 'Re: '+subject, oemReplyText);
          hubLog('draft_created', 'W3+OEM draft: '+mpn, {mpn:mpn, type:'w3_oem', tp:tp||null});
          Logger.log('W3 not quoted + OEM draft for '+mpn);
          handled = true;
        } else {
          Logger.log('Added to Stan sheet (not yet quoted): '+w3Mpn+' — leaving in inbox for John');
          handled = true;
        }
      }
    }

    // OEM EXCESS only branch
    if (!handled && oemResults.length) {
      var forteMatches = checkForteForMPN(mpn,3650);
      var quotedEntry = null;
      for (var f=0;f<forteMatches.length;f++){if(forteMatches[f].status==='QUOTED'&&forteMatches[f].colH){quotedEntry=forteMatches[f];break;}}
      var replyText2, htmlBody2, draftType;
      if (quotedEntry) {
        var quotedPrice = quotedEntry.colH;
        if (quotedPrice && !isNaN(parseFloat(quotedPrice))) {
          quotedPrice = '$' + parseFloat(quotedPrice).toFixed(2) + ' each';
        }
        replyText2 = quotedPrice + (quotedEntry.colI?' | '+quotedEntry.colI:'');
        draftType = 'forte_quoted';
        htmlBody2 = origMsg?buildDraftHTML(replyText2,origMsg):'<div dir="ltr">'+replyText2+getSignatureHTML()+'</div>';
        createThreadedDraft(replyTo,'Re: '+subject,htmlBody2,replyToId,threadId,null);
        Logger.log('QUOTED Forte draft: '+mpn);
      } else if (tp) {
        if (hasBillExt) {
          replyText2 = MSG_BILL; draftType = 'bill';
          htmlBody2 = origMsg?buildDraftHTML(MSG_BILL,origMsg):'<div dir="ltr">'+MSG_BILL+getSignatureHTML()+'</div>';
          createThreadedDraft(replyTo,'Re: '+subject,htmlBody2,replyToId,threadId,BILL_EMAIL);
        } else {
          replyText2 = MSG_CHECKING; draftType = 'checking';
          htmlBody2 = origMsg?buildDraftHTML(MSG_CHECKING,origMsg):'<div dir="ltr">'+MSG_CHECKING+getSignatureHTML()+'</div>';
          createThreadedDraft(replyTo,'Re: '+subject,htmlBody2,replyToId,threadId,null);
        }
      } else {
        replyText2 = tpMsg; draftType = 'need_tp';
        htmlBody2 = origMsg?buildDraftHTML(tpMsg,origMsg):'<div dir="ltr">'+tpMsg+getSignatureHTML()+'</div>';
        createThreadedDraft(replyTo,'Re: '+subject,htmlBody2,replyToId,threadId,null);
      }
      hubPostDraft(threadId, mpn, replyTo, 'Re: '+subject, replyText2);
      hubLog('draft_created', 'OEM draft: '+mpn+' ('+draftType+')', {mpn:mpn, type:draftType, tp:tp||null, country:country});
      Logger.log('OEM only draft: '+mpn);
      handled = true;
    }

    if (handled) {
      thread.addLabel(label);
    } else {
      Logger.log('No match / leaving in inbox: ' + mpn);
    }
  });
}

// ============================================================
// TRIGGER 5 — Sent "checking on it now" → add to Forte
// ============================================================
function checkSentCheckingReplies() {
  var _cfg = getRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) { hubLog('run', 'checkSentCheckingReplies: disabled via hub config'); return; }
  var query = 'in:sent "checking on it now" newer_than:2d -label:oem-rfq-sent-processed';
  var threads = GmailApp.search(query,0,20);
  hubLog('run', 'checkSentCheckingReplies: ' + threads.length + ' thread(s)');
  if (!threads.length) return;
  var label = GmailApp.getUserLabelByName('oem-rfq-sent-processed')||GmailApp.createLabel('oem-rfq-sent-processed');
  threads.forEach(function(thread) {
    try {
      var messages = thread.getMessages();
      var sentMsg = null;
      for (var i=messages.length-1;i>=0;i--){
        if (messages[i].getFrom().indexOf(JOHN_EMAIL)>=0&&messages[i].getPlainBody().indexOf('checking on it now')>=0){sentMsg=messages[i];break;}
      }
      if (!sentMsg){thread.addLabel(label);return;}
      var mpn = extractMPN(thread.getFirstMessageSubject());
      if (!mpn || !searchOEMExcess(mpn).length) {
        var firstBuyerBody = '';
        for (var k=0;k<messages.length;k++){
          if (messages[k].getFrom().indexOf(JOHN_EMAIL)<0&&messages[k].getFrom().indexOf('intransittech')<0){
            firstBuyerBody = messages[k].getPlainBody(); break;
          }
        }
        var bodyMpn = extractMPNFromRFQBody(firstBuyerBody);
        if (bodyMpn && searchOEMExcess(bodyMpn).length) mpn = bodyMpn;
      }
      if (!mpn||!searchOEMExcess(mpn).length){thread.addLabel(label);return;}
      var tp = null, country = 'USA', qty = '';
      var johnHasReplied = false;
      for (var j=0;j<messages.length;j++){
        var msg = messages[j];
        if (msg.getFrom().indexOf(JOHN_EMAIL)>=0||msg.getFrom().indexOf('intransittech')>=0) {
          johnHasReplied = true; continue;
        }
        country = extractCountryFromEmail(msg.getFrom());
        var fullMsgBody = msg.getPlainBody();
        var directText = stripQuotedLines(fullMsgBody);
        if (johnHasReplied) {
          var buyerTP = null;
          if (!isNetcompEmail(msg.getFrom(), msg.getSubject())) {
            buyerTP = extractNetcompsTgtPrice(directText, msg.getBody());
          }
          if (!buyerTP) {
            var bodyClean = directText.split('OEM EXCESS')[0];
            buyerTP = extractTargetPrice(bodyClean);
          }
          if (buyerTP && !tp) tp = buyerTP;
        } else if (isNetcompEmail(msg.getFrom(), msg.getSubject())) {
          var netcompTP = extractNetcompsTgtPrice(directText, msg.getBody());
          if (netcompTP && !tp) tp = netcompTP;
        }
          var qtyMatch = fullMsgBody.match(/(\d[\d\s,.]*\d|\d)\s*(?:pcs?|pieces?|units?|ea|each)|QtyReq\s+(\d+)|Qty\s*[：:]\s*(\d[\d,]*)/i);
          if (qtyMatch && !qty) {
            var rawQty = qtyMatch[1] || qtyMatch[2] || qtyMatch[3] || '';
            var qtyNum = parseQtyValue(rawQty, country);
            if (qtyNum > 0) qty = String(qtyNum);
          }
          if (!qty && isNetcompEmail(msg.getFrom(), msg.getSubject())) {
            var netcompsQty = extractNetcompsQtyReq(msg.getBody());
            if (netcompsQty) qty = netcompsQty;
          }
          if (!qty) {
            var kMatchSent = fullMsgBody.match(/\b(\d+(?:\.\d+)?)\s*[Kk]\b/);
            if (kMatchSent) qty = String(Math.round(parseFloat(kMatchSent[1]) * 1000));
          }
      }
      if (!qty) {Logger.log('No QTY, skipping Forte: '+mpn);thread.addLabel(label);return;}
      if (tp) {
        var qtyNum = parseFloat(qty)||0;
        var priceNum = parseFloat(String(tp).replace(/[^0-9.]/g,''))||0;
        if (qtyNum > 0 && priceNum > 0 && qtyNum * priceNum < 500) {
          Logger.log('Below MOV, skipping Forte: '+mpn+' | '+qty+' x '+tp);
          thread.addLabel(label);return;
        }
      }
      var forteResults = checkForteForMPN(mpn,60);
      var hasRecent = forteResults.some(function(r){return r.recent&&r.status.toLowerCase()!=='closed';});
      if (!hasRecent) {
        addToForteSheet(mpn,qty,tp,country,'');
        hubLog('run', 'Added to Forte: '+mpn, {mpn:mpn, tp:tp||null, qty:qty, country:country});
        Logger.log('Added to Forte: '+mpn+' | TP: '+tp);
      } else {
        Logger.log('Forte 60-day skip: '+mpn);
      }
    } catch(e) { Logger.log('Error checkSentChecking: '+e.toString()); }
    thread.addLabel(label);
  });
}

// ============================================================
// TRIGGER 6 — Payment Advice → forward to Deb
// ============================================================
function checkInboxForPaymentAdvice() {
  var _cfg = getRemoteConfig(); applyRemoteConfig(_cfg);
  if (_cfg.enabled === false) { hubLog('run', 'checkInboxForPaymentAdvice: disabled via hub config'); return; }
  var query = 'in:inbox subject:"payment advice" -label:oem-payment-forwarded';
  var threads = GmailApp.search(query,0,10);
  hubLog('run', 'checkInboxForPaymentAdvice: ' + threads.length + ' thread(s)');
  if (!threads.length) return;
  var label = GmailApp.getUserLabelByName('oem-payment-forwarded')||GmailApp.createLabel('oem-payment-forwarded');
  threads.forEach(function(thread) {
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

// ============================================================
// APPEND DAVID LIST
// ============================================================
function appendDavidList() {
  var NOTE = 'OEM EXCESS! $500 MIN TP REQUIRED';
  var davidData = SpreadsheetApp.openById('1NbzAKOLkSfQCp5ex9QPgTF_lwE6V0jAB5aKNSdJ_NFE').getSheets()[0].getDataRange().getValues();
  var oemSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIN_SHEET_NAME);
  var startRow = (String(davidData[0][0]).toLowerCase().indexOf('part')>=0||String(davidData[0][0]).toLowerCase().indexOf('mpn')>=0)?1:0;
  var rows = [];
  for (var i=startRow;i<davidData.length;i++){var mpn=String(davidData[i][0]||'').trim();if(mpn)rows.push([mpn,String(davidData[i][1]||'').trim(),'',String(davidData[i][2]||'').trim(),NOTE]);}
  if (rows.length) oemSheet.getRange(oemSheet.getLastRow()+1,1,rows.length,5).setValues(rows);
  Logger.log('Appended '+rows.length+' rows.');
}

// ============================================================
// WEB APP
// ============================================================
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

// ============================================================
// TRIGGERS
// ============================================================
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger('checkDavidNoStockEmails').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('checkSentRemovals').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('checkInboxForTPReplies').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('checkInboxForNewRFQs').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('checkSentCheckingReplies').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('checkInboxForPaymentAdvice').timeBased().everyMinutes(5).create();
  Logger.log('All 6 triggers installed.');
}

// ============================================================
// TEST FUNCTIONS
// ============================================================
function testSearch(mpn) {
  Logger.log('--- OEM EXCESS ---'); searchOEMExcess(mpn);
  Logger.log('--- IN STOCK ---'); searchInStock(mpn);
  Logger.log('--- STAN SHEET ---'); searchStanSheet(mpn);
}
function testNewRFQs() { checkInboxForNewRFQs(); }
function testTPReplies() { checkInboxForTPReplies(); }
function testSentChecking() { checkSentCheckingReplies(); }

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

function addMissingForteEntries() {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow1 = sheet.getLastRow() + 1;
  sheet.appendRow(['6/2/2026', 'FLLXT971ABE.A4', 488, '$12', '', 'China', '=C'+nextRow1+'*D'+nextRow1, '', '', '', 'Open']);
  Logger.log('Added FLLXT971ABE.A4: QTY=488, TP=$12');
  var nextRow2 = sheet.getLastRow() + 1;
  sheet.appendRow(['6/2/2026', 'AD8606ARMZ', 126656, '$0.7', '', 'China', '=C'+nextRow2+'*D'+nextRow2, '', '', '', 'Open']);
  Logger.log('Added AD8606ARMZ: QTY=126656, TP=$0.70');
  var nextRow3 = sheet.getLastRow() + 1;
  sheet.appendRow(['6/2/2026', 'JX552-50003', 20, '$25', '', 'USA', '=C'+nextRow3+'*D'+nextRow3, '', '', '', 'Open']);
  Logger.log('Added JX552-50003: QTY=20, TP=$25');
}

function addMT46ToForte() {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  sheet.appendRow(['6/2/2026', 'MT46H32M32LFB5-5ITB', 3869, 15, '', 'Germany', 58035, '', '', '', 'Open']);
}

function fixM426ForteTP() {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]).trim() === 'M426R4GA3PB0-CWM' && data[i][3] == 5536) {
      sheet.getRange(i + 1, 4).setValue('$1650');
      Logger.log('Fixed M426R4GA3PB0-CWM Forte TP: 5536 -> $1650');
      return;
    }
  }
  Logger.log('M426R4GA3PB0-CWM row with TP=5536 not found');
}

function add1718628ToForte() {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['6/3/2026', '1-1718628-1', 3600, '$1.50', '', 'China', '=C'+nextRow+'*D'+nextRow, '', '', '', 'Open']);
  Logger.log('Added 1-1718628-1 to Forte: QTY=3600, TP=$1.50, China');
}

function addSN74ToForte() {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['6/2/2026', 'SN74LVC1G139DCUR', 12000, '$0.20', '', 'USA', '=C'+nextRow+'*D'+nextRow, '', '', '', 'Open']);
  Logger.log('Added SN74LVC1G139DCUR to Forte: QTY=12000, TP=$0.20, USA');
}

function fixTDA21240ForteQty() {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]).trim().toUpperCase() === 'TDA21240' && (!data[i][2] || data[i][2] === '')) {
      var rowNum = i + 1;
      sheet.getRange(rowNum, 3).setValue(10000);
      sheet.getRange(rowNum, 7).setFormula('=C' + rowNum + '*D' + rowNum);
      Logger.log('Fixed TDA21240 Forte row ' + rowNum + ': QTY → 10000, G → formula');
      return;
    }
  }
  Logger.log('TDA21240 Forte row with blank QTY not found');
}

function addCM1224ToForte() {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['6/16/2026', 'CM1224-04SO', 24000, '$0.115', '', 'USA', '=C'+nextRow+'*D'+nextRow, '', '', 'A2 Global / Bernard Benson — TP reply was $.115 (no leading zero bug)', 'Open']);
  Logger.log('Added CM1224-04SO: QTY=24000, TP=$0.115, USA');
}

function addLTC2914ToForte() {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['6/16/2026', 'LTC2914IDHC-1', 20000, '$3.00', '', 'Slovakia', '=C'+nextRow+'*D'+nextRow, '', '', 'Direct RFQ via netCOMPONENTS listing — Sparetronics 6/16/2026', 'Open']);
  Logger.log('Added LTC2914IDHC-1: QTY=20000, TP=$3.00, Slovakia');
}

function addPE60252ToForte() {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['6/16/2026', 'PE60252B1-000U-A99', 300, '$18.00', '', 'France', '=C'+nextRow+'*D'+nextRow, '', '', 'Alantys Technology / Fabio Buono — TP reply "18 usd" (extractNetcompsBuyerEmail <mailto:> bug)', 'Open']);
  Logger.log('Added PE60252B1-000U-A99: QTY=300, TP=$18.00, France');
}

function addFRJ2A11ToForte() {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['6/16/2026', 'FRJ-2A11', 1864, '$0.35', '', 'USA', '=C'+nextRow+'*D'+nextRow, '', '', 'Right Choice Electronics / Robert Campbell — TP reply was ".35" (bare decimal bug)', 'Open']);
  Logger.log('Added FRJ-2A11: QTY=1864, TP=$0.35, USA');
}

function addAK3076CYToForte() {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['6/16/2026', 'AK3-076C-Y', 300, '$5.00', '', 'Spain', '=C'+nextRow+'*D'+nextRow, '', '', 'TCX Micro S.L. / Jake Taylor — TP reply "TP 5$ ea" (<mailto:> bug + label-without-draft bug)', 'Open']);
  Logger.log('Added AK3-076C-Y: QTY=300, TP=$5.00, Spain');
}

function fixTDA21240() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  var deletedSheet = getOrCreateDeletedSheet(ss);
  var rowData = mainSheet.getRange(125117, 1, 1, 5).getValues()[0];
  if (String(rowData[0]).trim().toUpperCase() === 'TDA21240' && String(rowData[4]).indexOf('BILL EXT 117') >= 0) {
    logDeletion(deletedSheet, rowData, 'Bill Pratt - no longer available 6/2/2026');
    mainSheet.deleteRow(125117);
    Logger.log('Deleted: TDA21240 BILL EXT 117 row 125117');
  } else {
    Logger.log('WARNING: Row 125117 does not match expected TDA21240 BILL EXT 117 — NOT deleted. Content: ' + JSON.stringify(rowData));
  }
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  forteSheet.appendRow(['6/2/2026', 'TDA21240', '', '$4.30', '', 'China', '', '', '', '', 'Open']);
  Logger.log('Added TDA21240 to Forte: TP $4.30 | China');
}
