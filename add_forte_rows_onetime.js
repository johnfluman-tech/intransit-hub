// ONE-TIME scripts — paste into Apps Script and run as needed

// ─────────────────────────────────────────────────────────────────
// populateInStockPriceHistory()
//
// For every row in the IN STOCK sheet, searches all sent Gmail for
// that MPN and writes:
//   Col J (price_history)  — full text of up to 5 recent sent emails
//   Col F (price_to_quote) — last per-unit price extracted from emails,
//                            only written if col F is currently blank
//
// Skips rows that already have col J filled (unless it's "No sent quotes found").
// Supports resume: saves row index in Script Properties. If it times out,
// just run again — it picks up where it left off.
// Run clearPriceHistoryProgress() to reset and start from scratch.
// ─────────────────────────────────────────────────────────────────
function populateInStockPriceHistory() {
  var IN_STOCK_ID = '1iOFHUBiWRgA6EjtO2ujoGpz-8v1qTRkgCXSvCa2Gf54';
  var sheet = SpreadsheetApp.openById(IN_STOCK_ID).getSheets()[0];
  var data  = sheet.getDataRange().getValues();

  var props    = PropertiesService.getScriptProperties();
  var startIdx = parseInt(props.getProperty('priceHistoryProgress') || '1', 10);
  var MAX_PER_RUN = 75; // stay well inside 6-min limit (~3-4 s per MPN Gmail search)
  var processed = 0, updated = 0, skipped = 0;

  Logger.log('populateInStockPriceHistory: ' + data.length + ' total rows, starting at row ' + (startIdx + 1));

  for (var i = startIdx; i < data.length; i++) {
    if (processed >= MAX_PER_RUN) {
      props.setProperty('priceHistoryProgress', String(i));
      Logger.log('MAX_PER_RUN reached — saved progress at row ' + (i + 1) + '. Run again to continue.');
      return;
    }

    var mpn = String(data[i][0]).trim();
    if (!mpn) { skipped++; continue; } // blank row

    // Skip rows that already have real history content
    var existingHistory = String(data[i][9] || '').trim();
    if (existingHistory && existingHistory !== 'No sent quotes found') { skipped++; continue; }

    // Gmail search — uses getRecentSentQuotesFull() already in email_script_v24_hub.js
    var history = getRecentSentQuotesFull(mpn, 5);
    var hasQuotes = history &&
      history.indexOf('No prior sent') < 0 &&
      history.indexOf('No sent messages') < 0 &&
      history.indexOf('No sent quotes') < 0 &&
      history.indexOf('error') < 0;

    if (hasQuotes) {
      sheet.getRange(i + 1, 10).setValue(history);

      // Auto-populate col F if blank — extract per-unit price from history
      var existingPrice = String(data[i][5] || '').trim();
      if (!existingPrice) {
        var price = extractPerUnitPriceFromHistory_(history);
        if (price !== null) {
          sheet.getRange(i + 1, 6).setValue(price);
          Logger.log('Row ' + (i + 1) + ' (' + mpn + '): history written, price=$' + price + ' auto-filled in col F');
        } else {
          Logger.log('Row ' + (i + 1) + ' (' + mpn + '): history written, no clear per-unit price found');
        }
      } else {
        Logger.log('Row ' + (i + 1) + ' (' + mpn + '): history written, col F already has value');
      }
      updated++;
    } else {
      sheet.getRange(i + 1, 10).setValue('No sent quotes found');
      Logger.log('Row ' + (i + 1) + ' (' + mpn + '): no quotes found');
    }

    processed++;
    if (processed % 10 === 0) SpreadsheetApp.flush();
  }

  // Completed all rows
  props.deleteProperty('priceHistoryProgress');
  SpreadsheetApp.flush();
  Logger.log('populateInStockPriceHistory DONE. Processed=' + processed + ', Updated=' + updated + ', Skipped=' + skipped);
}


// Extracts the most likely per-unit price from a getRecentSentQuotesFull() string.
// Prefers decimal prices (e.g. $1.25, $12.50) over whole numbers.
// Ignores $500 and $2000 (minimum line values), and prices above $50,000.
function extractPerUnitPriceFromHistory_(historyText) {
  if (!historyText || historyText.indexOf('$') < 0) return null;

  // Decimal price (strong signal — per-unit prices almost always have cents)
  var decMatches = historyText.match(/\$\s*(\d{1,6}\.\d{1,4})/g) || [];
  for (var d = 0; d < decMatches.length; d++) {
    var p = parseFloat(decMatches[d].replace(/[$,\s]/g, ''));
    if (p > 0 && p !== 500 && p !== 2000 && p < 50000) return p;
  }

  // Whole-number price (weaker signal — exclude known line minimums and large totals)
  var intMatches = historyText.match(/\$\s*(\d{1,5})\b/g) || [];
  for (var n = 0; n < intMatches.length; n++) {
    var pi = parseInt(intMatches[n].replace(/[$,\s]/g, ''), 10);
    if (pi > 0 && pi !== 500 && pi !== 2000 && pi < 5000) return pi;
  }

  return null;
}


// Resets progress so the next populateInStockPriceHistory() run starts from row 2.
function clearPriceHistoryProgress() {
  PropertiesService.getScriptProperties().deleteProperty('priceHistoryProgress');
  Logger.log('Progress cleared. Next run will start from row 2.');
}




// ─────────────────────────────────────────────────────────────────
// Run removeDuplicateOEM_9GA0812P4H001() to remove 2 duplicate rows
// for 9GA0812P4H001 (rows 63627 and 63628 — same qty/notes as row 63626).
// Keeps row 63626, deletes the duplicates.
// ─────────────────────────────────────────────────────────────────
function removeDuplicateOEM_9GA0812P4H001() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var sheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  // Delete in descending order so row numbers stay valid
  sheet.deleteRow(63628);
  sheet.deleteRow(63627);
  Logger.log('Deleted duplicate OEM rows 63627 and 63628 for 9GA0812P4H001. Row 63626 remains.');
}

// ─────────────────────────────────────────────────────────────────
// Run fixDavidNoStkDrafts_Aug10() to:
//   1. Delete all wrong drafts in 5 David no-stk threads
//      (wrong drafts: "Ok, noted.", "claude", MSG_CHECKING)
//   2. Create correct "Ok, removed from listing." reply in each thread
//
// Threads:
//   19febdd64d4eefd8  MAX4886ETO+T  #4273  No Stk
//   19febdb03075bb5b  IRLML6402TRPBF #4272  No stk
//   19febd5cc71f7a5d  IS43TR16256B-125KBLI #4271 NO STOCK
//   19febcfe95391b16  MT40A1G16TB-062EIT:FTR  #4266 and 4267  No stk
//   19febce0bbb4a45c  AFB0612DH-TZUT #4260 NO STOCK
// ─────────────────────────────────────────────────────────────────
function fixDavidNoStkDrafts_Aug10() {
  var SIG_HTML =
    '<div><b><span style="color:rgb(31,73,125);font-family:Tahoma,sans-serif;font-size:10pt">Regards,</span></b></div>' +
    '<div><b><span style="color:rgb(31,73,125);font-family:Tahoma,sans-serif;font-size:10pt">John Fluman</span></b></div>' +
    '<div><b><span style="color:rgb(31,73,125);font-family:Arial,sans-serif;font-size:8pt">Intransit Technologies</span></b></div>' +
    '<div><a href="mailto:john.fluman@intransittech.com" style="font-family:Calibri;font-size:8pt">john.fluman@intransittech.com</a></div>' +
    '<div><i><span style="color:gray;font-family:Arial,sans-serif;font-size:7.5pt">An ISO 9001 Certified Company</span></i></div>' +
    '<div><span style="color:rgb(31,73,125);font-family:Tahoma,sans-serif;font-size:8pt">Toll (877) 677-5868 x101 - Local (949) 481-7935 x101</span></div>' +
    '<br><div><span style="color:rgb(166,166,166);font-family:Calibri,sans-serif;font-size:8pt">The information contained in this communication and its attachment(s) is intended only for the use of the individual to whom it is addressed and may contain information that is privileged, confidential, or exempt from disclosure. If the reader of this message is not the intended recipient, you are hereby notified that any dissemination, distribution, or copying of this communication is strictly prohibited. If you have received this communication in error, please notify <a href="mailto:john.fluman@intransittech.com">john.fluman@intransittech.com</a> and delete the communication without retaining any copies. Thank you.</span></div>';

  var BODY_TEXT = 'Ok, removed from listing.';
  var BODY_HTML = '<p>' + BODY_TEXT + '</p>' + SIG_HTML;

  var threadIds = [
    '19febdd64d4eefd8',
    '19febdb03075bb5b',
    '19febd5cc71f7a5d',
    '19febcfe95391b16',
    '19febce0bbb4a45c'
  ];

  // Step 1: Delete all existing wrong drafts in these threads
  var allDrafts = GmailApp.getDrafts();
  var deletedCount = 0;
  allDrafts.forEach(function(draft) {
    try {
      var threadId = draft.getMessage().getThread().getId();
      if (threadIds.indexOf(threadId) !== -1) {
        draft.deleteDraft();
        deletedCount++;
        Logger.log('Deleted wrong draft in thread ' + threadId);
      }
    } catch(e) {
      Logger.log('Error deleting draft: ' + e);
    }
  });
  Logger.log('Deleted ' + deletedCount + ' wrong draft(s)');

  // Step 2: Create correct "Ok, removed from listing." reply in each thread
  threadIds.forEach(function(threadId) {
    try {
      var thread = GmailApp.getThreadById(threadId);
      var messages = thread.getMessages();
      var lastMsg = messages[messages.length - 1];
      lastMsg.createDraftReply(BODY_TEXT, { htmlBody: BODY_HTML });
      Logger.log('Created correct draft in thread ' + threadId + ' (' + thread.getFirstMessageSubject() + ')');
    } catch(e) {
      Logger.log('Error creating draft for thread ' + threadId + ': ' + e);
    }
  });

  Logger.log('fixDavidNoStkDrafts_Aug10: DONE');
}


// ─────────────────────────────────────────────────────────────────
// Run davidNoStk_Aug10_ForteAndOEM() to stamp Forte + delete OEM EXCESS
// for the 5 new David no-stks received Aug 10:
//   AFB0612DH-TZUT       #4260
//   MT40A1G16TB-062EIT:FTR #4266 and #4267
//   IS43TR16256B-125KBLI  #4271
//   IRLML6402TRPBF        #4272
//   MAX4886ETO+T           #4273
// ─────────────────────────────────────────────────────────────────
function davidNoStk_Aug10_ForteAndOEM() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var OEM_EXCESS_ID  = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var stamp = 'NO STK - 8/10/2026';

  // Forte rows to stamp col K
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  [4260, 4266, 4267, 4271, 4272, 4273].forEach(function(rowNum) {
    try {
      var cell = forteSheet.getRange(rowNum, 11);
      var current = String(cell.getValue()).trim().toUpperCase();
      if (current.indexOf('NO STK') === -1 && current !== 'CLOSED') {
        cell.clearDataValidations();
        cell.setValue(stamp);
        cell.setBackground('#000000');
        cell.setFontColor('#FFFFFF');
        cell.setFontWeight('bold');
        Logger.log('Stamped Forte row ' + rowNum);
      } else {
        Logger.log('Forte row ' + rowNum + ' already stamped: ' + current);
      }
    } catch(e) { Logger.log('Error stamping row ' + rowNum + ': ' + e); }
  });

  // OEM EXCESS MPNs to delete
  var targets = ['AFB0612DH-TZUT', 'MT40A1G16TB-062EIT:FTR', 'IS43TR16256B-125KBLI', 'IRLML6402TRPBF', 'MAX4886ETO+T'];
  var ss = SpreadsheetApp.openById(OEM_EXCESS_ID);
  var mainSheet = ss.getSheetByName('sheet1');
  var deletedSheet = ss.getSheetByName('deleted');
  if (!deletedSheet) deletedSheet = ss.insertSheet('deleted');

  targets.forEach(function(target) {
    var data = mainSheet.getDataRange().getValues();
    var found = false;
    var normalTarget = target.replace(/-/g, '').replace(/:/g, '').toUpperCase();
    for (var i = data.length - 1; i >= 1; i--) {
      var mpn = String(data[i][0]).trim().replace(/-/g, '').replace(/:/g, '').toUpperCase();
      if (mpn === normalTarget) {
        found = true;
        var rowNum = i + 1;
        mainSheet.getRange(rowNum, 5).setValue('NO STK 8/10/2026');
        var logRow = data[i].slice();
        logRow.push('David no-stk 8/10/2026 - ' + target);
        deletedSheet.appendRow(logRow);
        mainSheet.deleteRow(rowNum);
        Logger.log('Deleted OEM EXCESS row ' + rowNum + ' (' + target + ')');
        data = mainSheet.getDataRange().getValues();
      }
    }
    if (!found) Logger.log(target + ' not found in OEM EXCESS (may already be removed)');
  });

  Logger.log('davidNoStk_Aug10_ForteAndOEM: DONE');
}


// Run fixTPS82130SILR_Aug10() to:
//   1. Delete TPS82130SIL from OEM EXCESS (row 129752) — same part as TPS82130SILR, David no-stk
//   2. Delete the wrong "similar MPN" draft to Demi Chen in thread 19fe9cec5da21d3f
function fixTPS82130SILR_Aug10() {
  var OEM_EXCESS_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var ss = SpreadsheetApp.openById(OEM_EXCESS_ID);
  var mainSheet = ss.getSheetByName('sheet1');
  var deletedSheet = ss.getSheetByName('deleted');
  if (!deletedSheet) deletedSheet = ss.insertSheet('deleted');

  var data = mainSheet.getDataRange().getValues();
  var found = false;
  for (var i = data.length - 1; i >= 1; i--) {
    var mpn = String(data[i][0]).trim().toUpperCase().replace(/-/g, '');
    if (mpn === 'TPS82130SIL' || mpn === 'TPS82130SILR') {
      var rowNum = i + 1;
      mainSheet.getRange(rowNum, 5).setValue('NO STK 8/10/2026');
      var logRow = data[i].slice();
      logRow.push('David no-stk 8/7/2026 - TPS82130SILR (listed as TPS82130SIL)');
      deletedSheet.appendRow(logRow);
      mainSheet.deleteRow(rowNum);
      Logger.log('Deleted OEM EXCESS row ' + rowNum + ' (TPS82130SIL)');
      data = mainSheet.getDataRange().getValues();
      found = true;
    }
  }
  if (!found) Logger.log('TPS82130SIL not found in OEM EXCESS — may already be removed');

  var drafts = GmailApp.getDrafts();
  for (var j = 0; j < drafts.length; j++) {
    try {
      if (drafts[j].getMessage().getThread().getId() === '19fe9cec5da21d3f') {
        drafts[j].deleteDraft();
        Logger.log('Deleted wrong similar-MPN draft for Demi Chen / TPS82130SILR');
        break;
      }
    } catch(e) {}
  }
  Logger.log('fixTPS82130SILR_Aug10: DONE');
}


// Run deleteForteRow4264() to remove wrong 2EDL23N06PJXUMA1 Forte entry
function deleteForteRow4264() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var mpn = String(sheet.getRange(4264, 2).getValue()).trim();
  Logger.log('Row 4264 MPN: ' + mpn);
  if (mpn.toUpperCase().indexOf('2EDL23N06') !== -1) {
    sheet.deleteRow(4264);
    Logger.log('Deleted Forte row 4264 (2EDL23N06PJXUMA1 — was BILL EXT)');
  } else {
    Logger.log('ERROR: Row 4264 MPN is "' + mpn + '" — not 2EDL23N06PJXUMA1. Check before deleting.');
  }
}


// DO NOT RUN — MT40A1G16TB-062EIT:FTR is now a David no-stk (#4266 and #4267).
// Running addForteAug9() would add a buyer entry for a part with no stock. Skip it.
// function addForteAug9() { ... }


// Run davidNoStk_Aug12() to stamp Forte rows + (OEM deletions already queued via command-queue #179-180):
//   OPA1678IDRGR        #4278  No stock
//   35MXC10000MEFCSN30X30 #4282  No stk
function davidNoStk_Aug12() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var stamp = 'NO STK - 8/12/2026';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  [4278, 4282].forEach(function(rowNum) {
    try {
      var cell = sheet.getRange(rowNum, 11);
      var cur = String(cell.getValue()).trim().toUpperCase();
      if (cur.indexOf('NO STK') === -1 && cur !== 'CLOSED') {
        cell.clearDataValidations();
        cell.setValue(stamp);
        cell.setBackground('#000000');
        cell.setFontColor('#FFFFFF');
        cell.setFontWeight('bold');
        Logger.log('Stamped Forte row ' + rowNum);
      } else {
        Logger.log('Row ' + rowNum + ' already stamped: ' + cur);
      }
    } catch(e) { Logger.log('Error row ' + rowNum + ': ' + e); }
  });
  Logger.log('davidNoStk_Aug12: DONE');
}


// Run davidNoStk_Aug12b() to stamp Forte rows for ADUM221N0BRWZ-RL #4290 and STM32G030C8T6 #4289
// OEM deletions already queued via command-queue #182-183
function davidNoStk_Aug12b() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var stamp = 'NO STK - 8/12/2026';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  [4290, 4289].forEach(function(rowNum) {
    try {
      var cell = sheet.getRange(rowNum, 11);
      var cur = String(cell.getValue()).trim().toUpperCase();
      if (cur.indexOf('NO STK') === -1 && cur !== 'CLOSED') {
        cell.clearDataValidations();
        cell.setValue(stamp);
        cell.setBackground('#000000');
        cell.setFontColor('#FFFFFF');
        cell.setFontWeight('bold');
        Logger.log('Stamped Forte row ' + rowNum);
      } else {
        Logger.log('Row ' + rowNum + ' already stamped: ' + cur);
      }
    } catch(e) { Logger.log('Error row ' + rowNum + ': ' + e); }
  });
  Logger.log('davidNoStk_Aug12b: DONE');
}


// ─────────────────────────────────────────────────────────────────
// Run removeAllDuplicateOEM() to scan the ENTIRE OEM EXCESS sheet,
// find every MPN that appears more than once, keep the first row,
// and delete all subsequent duplicate rows (in descending order).
// Safe to run multiple times — idempotent after first run.
// ─────────────────────────────────────────────────────────────────
function removeAllDuplicateOEM() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var ss = SpreadsheetApp.openById(OEM_SHEET_ID);
  var sheet = ss.getSheetByName('sheet1');

  var lastRow = sheet.getLastRow();
  Logger.log('Total rows in OEM EXCESS: ' + lastRow);

  // Read only column A (MPN) — much faster than getDataRange() on 60k+ rows
  var mpnValues = sheet.getRange(1, 1, lastRow, 1).getValues();

  var seen = {};    // normalized MPN → first row number
  var toDelete = []; // row numbers of duplicates

  // Row 1 is header; start at index 1 (row 2)
  for (var i = 1; i < mpnValues.length; i++) {
    var mpn = String(mpnValues[i][0]).trim().toUpperCase();
    if (!mpn) continue;  // skip blank rows
    var rowNum = i + 1;
    if (seen.hasOwnProperty(mpn)) {
      toDelete.push(rowNum);
    } else {
      seen[mpn] = rowNum;
    }
  }

  Logger.log('Duplicate rows found: ' + toDelete.length);
  if (toDelete.length === 0) {
    Logger.log('No duplicates found — sheet is clean.');
    return;
  }

  // Log sample of duplicate MPNs for review
  var mpnCounts = {};
  toDelete.forEach(function(rn) {
    var m = String(mpnValues[rn - 1][0]).trim().toUpperCase();
    mpnCounts[m] = (mpnCounts[m] || 0) + 1;
  });
  var sample = Object.keys(mpnCounts).slice(0, 30).map(function(m) {
    return m + '×' + (mpnCounts[m] + 1);
  });
  Logger.log('Sample duplicated MPNs (MPN×total_count): ' + sample.join(', '));

  // Delete bottom-up so row numbers above stay valid
  toDelete.sort(function(a, b) { return b - a; });

  var deleted = 0;
  for (var j = 0; j < toDelete.length; j++) {
    sheet.deleteRow(toDelete[j]);
    deleted++;
    if (deleted % 100 === 0) {
      Logger.log('Progress: ' + deleted + ' / ' + toDelete.length + ' deleted');
      SpreadsheetApp.flush();
    }
  }

  Logger.log('removeAllDuplicateOEM DONE. Deleted ' + deleted +
    ' duplicate rows. Sheet now has ~' + (lastRow - deleted) + ' rows.');
}


// Run addForteAug11() to add Forte entry for L6384ED013TR (msg_checking sent 8/11/2026).
// Buyer: Roger Zhang / HK Waykey Technology. OEM EXCESS 132,800 units confirmed.
function addForteAug11() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow([
    '8/11/2026',       // A — Date
    'L6384ED013TR',    // B — MPN
    132800,            // C — Qty
    0.15,              // D — Buyer TP
    '',                // E — John Buy (blank)
    'HK',              // F — Country
    '=C' + nextRow + '*D' + nextRow, // G — Potential
    '',                // H — John Quoted (blank)
    '',                // I — Notes (blank)
    '',                // J — History (auto-populated for duplicates)
    'Open'             // K — Status
  ]);
  Logger.log('Added L6384ED013TR — 132800 qty, $0.15 TP, HK buyer (HK Waykey Technology)');
}
