// ONE-TIME — Run fixStanSheet_July2() to:
//   1. Set QTY (col G) for LT1424CS8-5#PBF (row 327) = 200 (buyer: Bonnie Chan)
//   2. Set QTY (col G) for BCM5338MKQM (row 329) = 300 (buyer: Amy Gu)
//   3. Delete row 330 (LP2951ACMX-3.3/NOPB — wrongly added, OEM EXCESS part not W3)
//   4. Delete the wrong BCM5338MKQM draft (msg id 19f1e76df0013aa7) that had
//      model-hallucinated body text instead of "Warehouse is checking details..."
function fixStanSheet_July2() {
  var STAN_SHEET_ID = '1pGRDpkqftQNoEYna53MxRJfUY8jEf5_w32FNa56OUIM';
  var sheet = SpreadsheetApp.openById(STAN_SHEET_ID).getSheets()[0];

  // Fix QTYs (col G = column index 7, 1-based)
  sheet.getRange(327, 7).setValue(200);
  Logger.log('Updated Stan row 327 (LT1424CS8-5#PBF) QTY = 200');

  sheet.getRange(329, 7).setValue(300);
  Logger.log('Updated Stan row 329 (BCM5338MKQM) QTY = 300');

  // Delete LP2951ACMX-3.3/NOPB — row 330 (OEM EXCESS part, was wrongly added via fuzzy match)
  sheet.deleteRow(330);
  Logger.log('Deleted Stan row 330 (LP2951ACMX-3.3/NOPB — was OEM EXCESS, not W3)');

  // Delete the wrong BCM5338MKQM draft (message id: 19f1e76df0013aa7)
  // Body had hallucinated debugging text instead of the correct add_to_stan body
  var token = ScriptApp.getOAuthToken();
  var wrongMsgId = '19f1e76df0013aa7';
  var pageToken = '';
  var deleted = false;
  do {
    var url = 'https://gmail.googleapis.com/gmail/v1/users/me/drafts?maxResults=100' +
              (pageToken ? '&pageToken=' + pageToken : '');
    var resp = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
    var body = JSON.parse(resp.getContentText());
    var drafts = body.drafts || [];
    for (var i = 0; i < drafts.length; i++) {
      if (drafts[i].message && drafts[i].message.id === wrongMsgId) {
        UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + drafts[i].id, {
          method: 'delete',
          headers: { 'Authorization': 'Bearer ' + token }
        });
        Logger.log('Deleted wrong BCM5338MKQM draft id: ' + drafts[i].id);
        deleted = true;
        break;
      }
    }
    pageToken = body.nextPageToken || '';
  } while (pageToken && !deleted);
  if (!deleted) Logger.log('Wrong BCM5338MKQM draft not found — may already be deleted');
}

// ONE-TIME — Run fixAERI_RP604K331A() to:
//   1. Delete the wrong MOV-decline draft (r5557525408424450768) that said "7 pcs at $7/EA"
//      The qty parser grabbed 7 from "$7 each" (price) instead of 250 from the original RFQ.
//   2. Add correct Forte entry: RP604K331A-TR, 250 qty, $7 TP, US (Luzmaria Orozco / AERI)
//   NOTE: msg_checking draft to Luzmaria already created manually (r6502891987470216417).
function fixAERI_RP604K331A() {
  // Delete wrong draft
  var token = ScriptApp.getOAuthToken();
  UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/me/drafts/r5557525408424450768', {
    method: 'delete',
    headers: { 'Authorization': 'Bearer ' + token }
  });
  Logger.log('Deleted wrong MOV-decline draft r5557525408424450768 for RP604K331A-TR');

  // Add Forte entry
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['7/2/2026', 'RP604K331A-TR', 250, 7, '', 'US',
    '=C'+nextRow+'*D'+nextRow, '', '', '', 'Open']);
  Logger.log('Added RP604K331A-TR to Forte row ' + nextRow + ' (250 qty, $7 TP, US)');
}

// ONE-TIME — Run removeOemExcess_XGL4020() to stamp NO STK and delete OEM EXCESS
// row 133100 (XGL4020-472MEC, COILCRAFT, 44466 qty). David Poggi confirmed
// "Cant share" on 7/1/2026 (#3907 = Forte row reference).
function removeOemExcess_XGL4020() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var sheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
  sheet.getRange(133100, 5).setValue('NO STK ' + today);
  sheet.deleteRow(133100);
  Logger.log('Stamped NO STK and deleted OEM row 133100 for XGL4020-472MEC');
}

// ONE-TIME — Run addForte_ADG1606BRUZ_2v7() to add the Forte entry for Owen Dai's
// updated $2.7 TP on ADG1606BRUZ-REEL7 (qty=943, CN). Original $2.2 entry already
// exists at Forte row 3959. This adds a new row reflecting the raised offer.
function addForte_ADG1606BRUZ_2v7() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['7/2/2026', 'ADG1606BRUZ-REEL7', 943, 2.7, '', 'CN',
    '=C'+nextRow+'*D'+nextRow, '', '', '', 'Open']);
  Logger.log('Added ADG1606BRUZ-REEL7 $2.7 TP to Forte row ' + nextRow);
}

// ONE-TIME — Run updateStanQTY_June29() to fill in the missing QTY column for
// the 3 Warehouse#3 parts added to Stan's RFQ sheet on 6/29/2026 (rows 324-326).
// Quantities pulled from IN STOCK sheet: EP3SL150F780I3N=408, XC2S200E6FTG256C=120, EL4390CM=125
function updateStanQTY_June29() {
  var STAN_SHEET_ID = '1pGRDpkqftQNoEYna53MxRJfUY8jEf5_w32FNa56OUIM';
  var sheet = SpreadsheetApp.openById(STAN_SHEET_ID).getSheets()[0];
  // Col G (index 7, 1-based) = QTY
  var updates = [
    { row: 324, mpn: 'EP3SL150F780I3N', qty: 408 },
    { row: 325, mpn: 'XC2S200E6FTG256C', qty: 120 },
    { row: 326, mpn: 'EL4390CM',         qty: 125 },
  ];
  updates.forEach(function(u) {
    sheet.getRange(u.row, 7).setValue(u.qty);
    Logger.log('Updated Stan row ' + u.row + ' (' + u.mpn + ') QTY = ' + u.qty);
  });
}

// ONE-TIME — Run removeOemExcess_DA721700U32() to stamp NO STK and delete row
// 121688 (DA7217-00U32, BILL EXT 117, qty 12016). Bill said no bid 6/30/2026.
function removeOemExcess_DA721700U32() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var sheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
  sheet.getRange(121688, 5).setValue('NO STK ' + today);
  sheet.deleteRow(121688);
  Logger.log('Stamped NO STK and deleted OEM row 121688 for DA7217-00U32');
}

// ONE-TIME — Run addForte_WFM200S022XNN3() to add the missed Forte entry from
// the WFM200S022XNN3 msg_checking that automation filed incorrectly on Jun 29.
// Buyer: JICE ZHU / sales1@bzgj-ele.com, qty=500, TP=$3, CN
function addForte_WFM200S022XNN3() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['6/29/2026', 'WFM200S022XNN3', 500, 3, '', 'CN',
    '=C'+nextRow+'*D'+nextRow, '', '', '', 'Open']);
  Logger.log('Added WFM200S022XNN3 to Forte row ' + nextRow);
}

// ONE-TIME — Run removeOemExcess_AT25DF321A() to stamp NO STK and delete the 2
// OEM EXCESS rows for AT25DF321A-SH-T (rows 64430 and 64431).
// Bill confirmed Jun 25 2026: "they're no longer available."
function removeOemExcess_AT25DF321A() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var sheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];
  for (var i = 1; i < data.length; i++) {
    var mpn = String(data[i][1]).trim();
    if (mpn.toLowerCase() === 'at25df321a-sh-t') {
      // Stamp NO STK in Notes column (col E = index 4)
      sheet.getRange(i + 1, 5).setValue('NO STK ' + today);
      rowsToDelete.push(i + 1);
    }
  }
  // Delete rows bottom-up so indices stay valid
  rowsToDelete.reverse().forEach(function(r) { sheet.deleteRow(r); });
  Logger.log('removeOemExcess_AT25DF321A: stamped and deleted ' + rowsToDelete.length + ' rows for AT25DF321A-SH-T');
}

// ─── backfillForteHistory ─────────────────────────────────────────────────────
// ONE-TIME: Run once to populate col J (History) for duplicate MPNs in the Forte
// sheet from the last 3 months. For each row whose col J is blank and has at
// least one earlier entry for the same MPN, fills J with a formatted history of
// those prior entries (most recent first).
// Format per line: "M/D/YYYY | Qty: X | TP: Y | Quoted: Z | STATUS | Notes"
// ─────────────────────────────────────────────────────────────────────────────
function backfillForteHistory() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 90); // 3 months back from today

  // Build a map: normalizedMPN → array of {rowNum (1-indexed), date, qty, tp, quoted, notes, status, histJ}
  var byMpn = {};
  for (var i = 1; i < data.length; i++) {
    var mpn = String(data[i][1]).trim();
    if (!mpn) continue;
    var key = mpn.toLowerCase();
    if (!byMpn[key]) byMpn[key] = [];
    byMpn[key].push({
      rowNum:  i + 1,
      date:    data[i][0] ? new Date(data[i][0]) : null,
      qty:     String(data[i][2] || '').trim(),
      tp:      String(data[i][3] || '').trim(),
      quoted:  String(data[i][7] || '').trim(),
      notes:   String(data[i][8] || '').trim(),
      status:  String(data[i][10] || '').trim(),
      histJ:   String(data[i][9] || '').trim(),
    });
  }

  var updated = 0;
  for (var key in byMpn) {
    var rows = byMpn[key];
    if (rows.length < 2) continue; // Only duplicates

    // Skip MPNs with no entry in the last 3 months
    var hasRecent = rows.some(function(r) { return r.date && r.date >= cutoff; });
    if (!hasRecent) continue;

    for (var ri = 0; ri < rows.length; ri++) {
      var cur = rows[ri];
      if (cur.histJ) continue; // Already has history

      // Prior entries = those with an earlier date, or earlier row if same date
      var prior = rows.filter(function(r, idx) {
        if (idx === ri) return false;
        if (!cur.date) return idx < ri;
        if (!r.date)   return false;
        return r.date < cur.date || (r.date.getTime() === cur.date.getTime() && r.rowNum < cur.rowNum);
      });
      if (!prior.length) continue; // No prior entries, nothing to backfill

      // Sort prior most-recent first
      prior.sort(function(a, b) {
        if (!a.date && !b.date) return b.rowNum - a.rowNum;
        if (!a.date) return 1;
        if (!b.date) return -1;
        return b.date - a.date;
      });

      var histLines = prior.map(function(r) {
        var ds = r.date ? Utilities.formatDate(r.date, Session.getScriptTimeZone(), 'M/d/yyyy') : '?';
        var line = ds;
        if (r.qty)    line += ' | Qty: ' + r.qty;
        if (r.tp)     line += ' | TP: ' + r.tp;
        if (r.quoted) line += ' | Quoted: ' + r.quoted;
        if (r.status && r.status.toLowerCase() !== 'open') line += ' | ' + r.status;
        if (r.notes)  line += ' | ' + r.notes;
        return line;
      });

      sheet.getRange(cur.rowNum, 10).setValue(histLines.join('\n'));
      updated++;
    }
  }
  SpreadsheetApp.flush();
  Logger.log('backfillForteHistory: updated ' + updated + ' rows across ' + Object.keys(byMpn).length + ' MPNs scanned.');
}

// ONE-TIME SCRIPT — Paste in Apps Script editor and run addForteRowsToday()
// Adds missed Forte entries from 2026-06-19 where MSG_CHECKING was sent manually

var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';

function addForteRowsToday() {
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var today = '6/19/2026';

  var rows = [
    // [mpn, qty, buyerTP, country]
    ['ICE40LP1K-CB81',                      1000,    2.00,  'CA'],
    ['MT53E1G32D2FW-046 IT:B TR',            207,   50.00,  'CA'],
    ['MMQA33VT1G',                        200000,    0.19,  'NL'],
    ['SIT5356AI-FQ-33N0-10.000000Y',          45,   45.00,  'US'],
    // DEI1016B removed — BILL EXT 117 part, goes to Bill not Forte
  ];

  rows.forEach(function(r) {
    var nextRow = sheet.getLastRow() + 1;
    sheet.appendRow([
      today,        // A: Date
      r[0],         // B: MPN
      r[1],         // C: Qty (buyer need)
      r[2],         // D: Target to sell (buyer TP)
      '',           // E: Target to buy (john buy — fill in manually)
      r[3],         // F: Country
      '=C' + nextRow + '*D' + nextRow,  // G: Potential
      '',           // H: John Quoted
      '',           // I: Notes
      '',           // J: History (always blank)
      'Open'        // K: Status
    ]);
    Logger.log('Added: ' + r[0]);
  });

  Logger.log('Done — ' + rows.length + ' rows added to Forte sheet.');
}

// ONE-TIME — Run addMissingForteRows_Jun25() to fix Forte entries missed on 2026-06-25
// STPS20L15D: MSG_CHECKING draft was created then deleted, Forte never auto-added
// CAR1AP80DC12-S: checkSentCheckingReplies failed qty extraction (Qty= format), Forte skipped
function addMissingForteRows_Jun25() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var today = '6/25/2026';
  var rows = [
    // [mpn, qty, buyerTP, country]
    ['STPS20L15D',     5000, 0.35, 'SG'],  // chili.wu@ample.sg — Ample Electronics Singapore
    ['CAR1AP80DC12-S', 6982, 1.26, 'AU'],  // gparashou@electronic-components.com.au — Australia
    ['7114-5623-02',  16000, 0.40, 'ES'],  // f.sabelli@goldney.net — Goldney Electronics Spain; script wrongly declined (used netcomp listed qty=2 instead of buyer's 16K)
    ['FIS155NL',       1300, 0.8209, 'TR'],  // gulsah@boardelectronics.com — Board Elektronik Turkey; extractTargetPrice missed table/Unit Price($) format
  ];
  rows.forEach(function(r) {
    var nextRow = sheet.getLastRow() + 1;
    sheet.appendRow([
      today, r[0], r[1], r[2], '', r[3],
      '=C' + nextRow + '*D' + nextRow,
      '', '', '', 'Open'
    ]);
    Logger.log('Added: ' + r[0]);
  });
  Logger.log('Done — ' + rows.length + ' rows added.');
}

// ONE-TIME — Run addMissingForteRows_Jun25b() for Yazaki 7183-2412 missed due to European TP format bug
// Emre (velosi.co Slovakia) replied "0,18$" Jun 17 — extractTargetPrice failed, Label_166 applied, Forte skipped
// Using Jun 24 revised TP from Emre's follow-up: $0.17/ea, 2800 pcs
function addMissingForteRows_Jun25b() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var today = '6/25/2026';
  var rows = [
    // [mpn, qty, buyerTP, country]
    ['7183-2412', 2800, 0.17, 'SK'],  // emre.kurtulus@velosi.co — Velosi S.R.O. Slovakia; Jun 24 revised TP
  ];
  rows.forEach(function(r) {
    var nextRow = sheet.getLastRow() + 1;
    sheet.appendRow([
      today, r[0], r[1], r[2], '', r[3],
      '=C' + nextRow + '*D' + nextRow,
      '', '', '', 'Open'
    ]);
    Logger.log('Added: ' + r[0]);
  });
  Logger.log('Done — ' + rows.length + ' rows added.');
}

// ONE-TIME -- Run removeOemExcess_SN75ALS176DR() to stamp NO STK and delete OEM EXCESS row 120136.
// David confirmed 2026-06-29: "#3930 SN75ALS176DR No stock"
// NOTE: Forte row 3930 has buyerTP=500 (likely extracted from description bug, not a real buyer TP).
// John: review Forte row 3930 (SN75ALS176DR, Jun 26, qty=2500, TP=500, CN) and delete if bogus.
function removeOemExcess_SN75ALS176DR() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var sheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];
  for (var i = 1; i < data.length; i++) {
    var mpn = String(data[i][1]).trim();
    if (mpn.toUpperCase() === 'SN75ALS176DR') {
      sheet.getRange(i + 1, 5).setValue('NO STK ' + today);
      rowsToDelete.push(i + 1);
    }
  }
  rowsToDelete.reverse().forEach(function(r) { sheet.deleteRow(r); });
  Logger.log('removeOemExcess_SN75ALS176DR: stamped and deleted ' + rowsToDelete.length + ' row(s)');
}

// ONE-TIME — Run addForte_LMK00304SQXNOPB() to add missed msg_checking Forte entry.
// Carmen (carmen@mission-ic.com, Mission Electronics, CN) wrote "TP 4U" = $4/unit.
// Agent misread as no TP and created wrong request_tp_500 draft (Jun 30 2026).
function addForte_LMK00304SQXNOPB() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['6/30/2026', 'LMK00304SQX/NOPB', 5000, 4, '', 'CN',
    '=C'+nextRow+'*D'+nextRow, '', '', '', 'Open']);
  Logger.log('Added LMK00304SQX/NOPB to Forte row ' + nextRow);
}

// ONE-TIME — Run addForte_LCC110PTR() to add missed msg_checking Forte entry.
// Deniss Dedkovski (ddedkovski@class-ic.com, Classic Components, US) gave TgtPrice=3.
// Agent misread as no TP and created wrong request_tp_500 draft (Jun 30 2026).
function addForte_LCC110PTR() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['6/30/2026', 'LCC110PTR', 1100, 3, '', 'US',
    '=C'+nextRow+'*D'+nextRow, '', '', '', 'Open']);
  Logger.log('Added LCC110PTR to Forte row ' + nextRow);
}

// ONE-TIME — Run backfillForteNoStk() to stamp "NO STK - [date]" on all Forte rows
// for parts David confirmed no-stock between 2026-06-24 and 2026-07-01.
// The bug: executeWorkerDecision called deletePart() but never updateForteSheet(),
// so every automation-based removal left Forte status stuck at "Open".
// Each entry uses the date of David's email as the no-stock date.
// Skips rows already stamped "NO STK" or "CLOSED".
function backfillForteNoStk() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var data = sheet.getDataRange().getValues();

  // MPN → date of David's confirmed no-stock email (M/d/yyyy format)
  var removals = {
    // 2026-07-01
    'AD22283-B-R2':              '7/1/2026',
    'MRF1513NT1':                '7/1/2026',
    'S24SE05006PDFA':            '7/1/2026',
    'FPF2700MPX':                '7/1/2026',
    'BCR420UW6-7':               '7/1/2026',
    // 2026-06-30
    'VRF2933MP':                 '6/30/2026',
    'ISZ080N10NM6':              '6/30/2026',
    'UCC27282DR':                '6/30/2026',
    'F1778447M2ILB0':            '6/30/2026',
    'PM2120-330K-RC':            '6/30/2026',
    'PM2120330KRC':              '6/30/2026',  // variant without dashes
    'EPCS16SI8':                 '6/30/2026',
    'EPCS16SI8N':                '6/30/2026',
    'LMK00304SQX/NOPB':          '6/30/2026',
    'DA721700U32':               '6/30/2026',
    'DA7217-00U32':              '6/30/2026',  // variant with dashes
    // 2026-06-29
    'CL21A226MAYNNNE':           '6/29/2026',
    'ADG5206BCPZ':               '6/29/2026',
    'LTC2601IDDTRPBF':           '6/29/2026',
    'LOCTITE3609':               '6/29/2026',
    'MT25QL256ABA8ESF-0SIT':     '6/29/2026',
    'SN75ALS176DR':              '6/29/2026',
    // 2026-06-26
    '88E1112-C2-NNC1C000':       '6/26/2026',
    'BTS6143D':                  '6/26/2026',
    'LCMXO3LF-4300C-5BG256I':   '6/26/2026',
    // 2026-06-25
    'AM3352BZCZD60':             '6/25/2026',
    'CAR1AP80DC12-S':            '6/25/2026',
    'LM25037QMTX/NOPB':          '6/25/2026',
    'STPS20L15D':                '6/25/2026',
    'LTC4357IMS8PBF':            '6/25/2026',
    'AT25DF321A-SH-T':           '6/25/2026',
    // 2026-06-24
    'XC7A100T-1FTG256I':         '6/24/2026',
    'LMZM23600V3SILT':           '6/24/2026',
  };

  var updated = 0, skipped = 0;
  for (var i = 1; i < data.length; i++) {
    var mpn = String(data[i][1]).trim();
    var status = String(data[i][10] || '').trim();
    var noStkDate = removals[mpn] || removals[mpn.toUpperCase()] || removals[mpn.replace(/[-\/]/g, '')];
    if (!noStkDate) continue;
    // Skip already stamped or closed rows
    if (status.toUpperCase().indexOf('NO STK') >= 0 || status.toUpperCase() === 'CLOSED') {
      skipped++;
      continue;
    }
    var newStatus = 'NO STK - ' + noStkDate;
    var cell = sheet.getRange(i + 1, 11);  // col K = index 11 (1-based)
    cell.clearDataValidations();
    cell.setValue(newStatus);
    cell.setBackground('#000000');
    cell.setFontColor('#FFFFFF');
    cell.setFontWeight('bold');
    updated++;
    Logger.log('Stamped Forte row ' + (i + 1) + ': ' + mpn + ' → ' + newStatus);
  }
  SpreadsheetApp.flush();
  Logger.log('backfillForteNoStk complete — updated: ' + updated + ', skipped (already done): ' + skipped);
}

// ONE-TIME — Run removeForte_MRF1513NT1_S24SE05006PDFA() to delete the two wrong
// Forte entries created on 2026-07-01:
//   Row 3965 — MRF1513NT1: worker returned msg_checking but all OEM rows were BILL EXT;
//              should have been bill_handle (no Forte). Audit fixed the draft but Forte
//              was already written before the audit ran.
//   Row 3963 — S24SE05006PDFA: part was in OEM EXCESS at time of RFQ so msg_checking
//              was technically correct, but David immediately confirmed no stock
//              ("Cant find"), making the Forte entry useless/misleading.
function removeForte_MRF1513NT1_S24SE05006PDFA() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  // Delete row 3965 first (higher row number), then 3963, to keep indices valid
  sheet.deleteRow(3965);
  sheet.deleteRow(3963);
  Logger.log('Deleted wrong Forte rows 3963 (S24SE05006PDFA) and 3965 (MRF1513NT1)');
}

// ONE-TIME — Run deleteWrongDrafts() to remove duplicate drafts created by the
// race condition between Trigger 3 and Trigger 7 (both fired simultaneously before
// the optimistic-label fix was deployed on 2026-07-01). For each affected thread,
// keeps the newest draft and deletes all older ones.
function deleteWrongDrafts() {
  // Threads confirmed to have duplicate drafts as of 2026-07-01 list_drafts audit
  var TARGET_THREADS = [
    '19f1d56c8ab92920',  // AD5504BRUZ (also has old wrong-format draft)
    '19f1d5d1abfa7ea8',  // MRF1513NT1
    '19f1cc652555ffd6',  // AR8035-AL1B (Cyclops Electronics)
    '19f1cbdcf7fac6f0',  // LMX2594RHAT
    '19f1c917760eace1',  // OPA377AIDBVR
    '19f1c8cfdfd74d76',  // SCT070HU120G3AG
    '19f1c88d7eaf466d',  // RMLV1616AGSA-5S2#AA0 (Artel)
    '19f1c71f9111d2a2',  // BCX56-10
    '19f1c6147ae7e525',  // 357002K38M09 (thread 1)
    '19f1c597afdd914c',  // 357002K38M09 (thread 2)
    '19f1c445fc095bbb',  // HRF-AT4520-FL-TR
    '19f1c4478d1b79e5',  // M41T62LC6F
    '19f1c38fd8bae17d',  // CY8C5868AXI-LP035
    '19f1b9eb6683918b',  // 8A34001E-000AJG8
    '19f1b8ac36899753',  // RMLV1616AGSA-5S2#AA0 (Speed Supply)
  ];

  var allDrafts = GmailApp.getDrafts();
  var byThread = {};

  allDrafts.forEach(function(draft) {
    try {
      var threadId = draft.getMessage().getThread().getId();
      if (TARGET_THREADS.indexOf(threadId) === -1) return;
      if (!byThread[threadId]) byThread[threadId] = [];
      byThread[threadId].push({ draft: draft, date: draft.getMessage().getDate(), id: draft.getId() });
    } catch(e) {
      Logger.log('Error reading draft: ' + e.message);
    }
  });

  var deleted = 0;
  TARGET_THREADS.forEach(function(threadId) {
    var drafts = byThread[threadId];
    if (!drafts || drafts.length === 0) {
      Logger.log('Thread ' + threadId + ': no drafts found (already cleaned up)');
      return;
    }
    if (drafts.length === 1) {
      Logger.log('Thread ' + threadId + ': 1 draft — OK, nothing to delete');
      return;
    }
    // Keep newest, delete all older ones
    drafts.sort(function(a, b) { return b.date - a.date; });
    Logger.log('Thread ' + threadId + ': keeping ' + drafts[0].id + ' (' + drafts[0].date + ')');
    for (var i = 1; i < drafts.length; i++) {
      try {
        drafts[i].draft.deleteDraft();
        deleted++;
        Logger.log('  Deleted older draft: ' + drafts[i].id + ' (' + drafts[i].date + ')');
      } catch(e) {
        Logger.log('  Failed to delete ' + drafts[i].id + ': ' + e.message);
      }
    }
  });

  Logger.log('deleteWrongDrafts complete — deleted ' + deleted + ' duplicate draft(s)');
}

// ONE-TIME — Run addForte_MPM3695GRF250022() to add missed msg_checking Forte entry.
// kunhua yao (zhongshi@zssx.top, Zhong Shi Sheng Xin, CN) gave TgtPrice=12, qty=1050.
// Agent chose request_tp_500 (ignored TgtPrice column). Prior entry Apr 16 is outside 60 days.
function addForte_MPM3695GRF250022() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['6/30/2026', 'MPM3695GRF-25-0022', 1050, 12, '', 'CN',
    '=C'+nextRow+'*D'+nextRow, '', '', '', 'Open']);
  Logger.log('Added MPM3695GRF-25-0022 to Forte row ' + nextRow);
}

// ONE-TIME — Run addForte_AL8860MP13() to add missed msg_checking Forte entry.
// Joe Tucarella (joe@ableelectronics.com, Able Electronics, US) gave TgtPrice=0.09, qty=73473.
// Agent chose request_tp_500 (ignored TgtPrice column). Jun 30 2026.
function addForte_AL8860MP13() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['6/30/2026', 'AL8860MP-13', 73473, 0.09, '', 'US',
    '=C'+nextRow+'*D'+nextRow, '', '', '', 'Open']);
  Logger.log('Added AL8860MP-13 to Forte row ' + nextRow);
}

// ONE-TIME — Run removeForte_AFCT5765ATPZ() to delete the wrong Forte entry for
// AFCT5765ATPZ (row 3964). Bug: agent wrote forte_entry before forte_entry guard
// was in place. Line value was $125 (5 pcs × $25) — below $500 MOV, should have
// been declined, never added to Forte. Decline draft r4983561768500791146 created.
function removeForte_AFCT5765ATPZ() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  sheet.deleteRow(3964);
  Logger.log('Deleted wrong Forte row 3964 (AFCT5765ATPZ)');
}

// ONE-TIME — Run fixBillExtForteErrors() to:
//   1. Delete wrong Forte entries for PI4MSD5V9548AZDEX (row ~3983) and SC18IS606PWJ
//      (row ~3980) — both are BILL EXT parts, should never have been added to Forte.
//   2. Delete wrong msg_checking draft r-9072564021236759707 for PI4MSD5V9548AZDEX.
//      Correct bill_handle draft r-6308889964994501476 already created.
function fixBillExtForteErrors() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var targets = ['PI4MSD5V9548AZDEX', 'SC18IS606PWJ'];
  var rowsToDelete = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var mpn = String(data[i][1]).trim().toUpperCase();
    if (targets.indexOf(mpn) !== -1) {
      rowsToDelete.push(i + 1);
      Logger.log('Marking for deletion Forte row ' + (i + 1) + ': ' + mpn);
    }
  }
  rowsToDelete.sort(function(a, b) { return b - a; });
  rowsToDelete.forEach(function(r) {
    sheet.deleteRow(r);
    Logger.log('Deleted Forte row ' + r);
  });
  // Delete wrong PI4MSD5V9548AZDEX msg_checking draft
  var token = ScriptApp.getOAuthToken();
  UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/me/drafts/r-9072564021236759707', {
    method: 'delete',
    headers: { 'Authorization': 'Bearer ' + token }
  });
  Logger.log('Deleted wrong PI4MSD5V9548AZDEX msg_checking draft r-9072564021236759707');
  Logger.log('fixBillExtForteErrors complete');
}

// ONE-TIME — Run deleteOldWrongDrafts_Jul3() to delete automation-created drafts
// that were superseded by manually-correct drafts created 2026-07-03.
//   r3853834301542523820 — AM486DX5133V-16BHC (old, body unknown; new stan_quoted draft r-5364126226482229979 created)
//   r-4586403387947820516 — NRF52833-QIAA-R (old, body unknown; new request_tp_500 draft r-4410992122938450026 created)
//   r1716529454449499559 — BAS4002ARPPE6327 (old, body unknown; new msg_checking draft r-6402756441308139034 created)
//   r-1733567911978935294 — BCM56980B0KFSBG msg_checking (David confirmed "Cant find" 7/3/2026 — do NOT send to buyer)
//   r-3837120187750311243 — DTMH04-3PA msg_checking (stephen@haleinst.com) — wrong; OEM EXCESS + no TP = no_bid, never msg_checking
//   r2139607451800387398 — K471K15X7RF5UH5 msg_checking (tom.hull@cyclops-group.com) — David confirmed No stk 7/3/2026; delete before sending
//   r-3828300807569788190 — TLD5190QU request_tp_500 (Sunny@honortech-int.com) — OEM + no TP = no_bid per new rule; never send
//   r-2454284305981434702 — CN9130-2000-NGI-AUS-G request_tp_500 (janet.wang@unix-tech.com) — OEM + no TP = no_bid per new rule; never send
function deleteOldWrongDrafts_Jul3() {
  var token = ScriptApp.getOAuthToken();
  var toDelete = [
    'r3853834301542523820',
    'r-4586403387947820516',
    'r1716529454449499559',
    'r-1733567911978935294',
    'r-3837120187750311243',
    'r2139607451800387398',
    'r-3828300807569788190',
    'r-2454284305981434702',
  ];
  toDelete.forEach(function(draftId) {
    try {
      UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + draftId, {
        method: 'delete',
        headers: { 'Authorization': 'Bearer ' + token }
      });
      Logger.log('Deleted old draft: ' + draftId);
    } catch(e) {
      Logger.log('Could not delete ' + draftId + ': ' + e.message + ' (may already be deleted)');
    }
  });
  Logger.log('deleteOldWrongDrafts_Jul3 complete');
}

// ONE-TIME — Run addForte_BAS4002ARPPE6327() to add missed Forte entry.
// Daniel (daniel@szbshx.com, CN) asked about DC on 12204 pcs on 7/3/2026.
// msg_checking draft r-6402756441308139034 created; Forte was empty.
// No buyer TP given — adding entry with blank TP (qty drives it).
function addForte_BAS4002ARPPE6327() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['7/3/2026', 'BAS4002ARPPE6327', 12204, '', '', 'CN',
    '', '', '', '', 'Open']);
  Logger.log('Added BAS4002ARPPE6327 to Forte row ' + nextRow + ' (12204 qty, no TP, CN)');
}


// ONE-TIME — Run processAllDavidNoStk_Jul3() to process 12 David no-stk/cant-find
// emails received 2026-07-03. Performs:
//   1. Stamps "NO STK 7/3/2026" in OEM EXCESS col E for 16 rows, then deletes them.
//   2. Stamps "NO STK - 7/3/2026" in Forte col K for 8 MPNs (matched by name).
//   3. Deletes Forte row 3982 (diode — David says "no part number listed").
// IMPORTANT: Run this BEFORE fixBillExtForteErrors() — row 3982 shifts if SC18IS606PWJ
// (row 3980) is deleted first.
// OEM rows covered:
//   QCA7005-AL33-R-1: 115084, 115085 | LT3580IMS8ETRPBF: 95669 (BILL EXT), 95670
//   PI4MSD5V9548AZDEX: 112936 (BILL EXT), 112937 (BILL EXT) | SC18IS606PWJ: 120217 (BILL EXT)
//   DDTC113ZCA-7-F: 76764 | BCM56980B0KFSBG: 67843 (BILL EXT), 67844 (BILL EXT), 67845
//   STM32G474RBT6: 127151 (NOT 126975 — different part) | HMC547ALP3E: 86598
//   ADG1606BRUZ-REEL7: 62633 | MX66L1G45GXDI-08G: 107573 (BILL EXT) | MCDP6000C1: 102439
function processAllDavidNoStk_Jul3() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';

  // OEM EXCESS: stamp col E then delete (sorted descending to keep row numbers valid during deletion)
  var oemRows = [127151, 120217, 115085, 115084, 112937, 112936, 107573, 102439,
                 95670, 95669, 86598, 76764, 67845, 67844, 67843, 62633];
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  oemRows.forEach(function(r) { oemSheet.getRange(r, 5).setValue('NO STK 7/3/2026'); });
  SpreadsheetApp.flush();
  oemRows.forEach(function(r) {
    oemSheet.deleteRow(r);
    Logger.log('OEM EXCESS: deleted row ' + r);
  });

  // Forte: stamp NO STK by MPN name (searches full sheet, skips already-stamped rows)
  var mpnsToStamp = [
    'QCA7005-AL33-R-1', 'LT3580IMS8ETRPBF', 'DDTC113ZCA-7-F',
    'BCM56980B0KFSBG', 'STM32G474RBT6', 'HMC547ALP3E',
    'ADG1606BRUZ-REEL7', 'MCDP6000C1'
  ];
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var data = forteSheet.getDataRange().getValues();
  var stamped = 0;
  for (var i = 1; i < data.length; i++) {
    var mpn = String(data[i][1]).trim();
    var status = String(data[i][10] || '').trim();
    if (mpnsToStamp.indexOf(mpn) !== -1 && status.toUpperCase().indexOf('NO STK') < 0) {
      var cell = forteSheet.getRange(i + 1, 11);
      cell.clearDataValidations();
      cell.setValue('NO STK - 7/3/2026');
      cell.setBackground('#000000');
      cell.setFontColor('#FFFFFF');
      cell.setFontWeight('bold');
      Logger.log('Forte row ' + (i + 1) + ': ' + mpn + ' → NO STK - 7/3/2026');
      stamped++;
    }
  }

  // Forte row 3982: delete bad entry (David: "no part number listed" — no valid MPN)
  forteSheet.deleteRow(3982);
  Logger.log('Deleted bad Forte entry row 3982 (diode — no valid MPN)');

  SpreadsheetApp.flush();
  Logger.log('processAllDavidNoStk_Jul3 complete — OEM rows deleted: ' + oemRows.length + ', Forte rows stamped: ' + stamped);
}

// ONE-TIME — Run removeOem_K471K15X7RF5UH5_Jul3() to process David's "No stk" reply.
// David reported #3981 K471K15X7RF5UH5 No stk on 7/3/2026. Trigger 7 (old script) took no action.
// Actions: stamp OEM row 90459 NO STK + delete it; stamp Forte row 3981 NO STK - 7/3/2026.
// Reply draft to David (r-6630817906763203033) already created in Gmail — just run and send.
function removeOem_K471K15X7RF5UH5_Jul3() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';

  // Stamp OEM row 90459 then delete
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  oemSheet.getRange(90459, 5).setValue('NO STK 7/3/2026');
  SpreadsheetApp.flush();
  oemSheet.deleteRow(90459);
  Logger.log('Stamped and deleted OEM row 90459 (K471K15X7RF5UH5)');

  // Stamp Forte row 3981 NO STK
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var cell = forteSheet.getRange(3981, 11);
  cell.clearDataValidations();
  cell.setValue('NO STK - 7/3/2026');
  cell.setBackground('#000000');
  cell.setFontColor('#FFFFFF');
  cell.setFontWeight('bold');
  SpreadsheetApp.flush();
  Logger.log('Stamped Forte row 3981 K471K15X7RF5UH5 → NO STK - 7/3/2026');
}

// ONE-TIME — Run removeOem_INMP411ACEZ_Jul3()
// David reported INMP411ACEZ #3987 No stk on 7/3/2026.
// Forte row 3987 = Open. OEM EXCESS may have INMP411ACEZ-R7 row 87945 (fuzzy) — stamp only if exact found.
// Reply draft r6068198179012203386 already created.
function removeOem_INMP411ACEZ_Jul3() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  // Search for exact INMP411ACEZ in OEM (col B)
  var data = oemSheet.getRange(2, 2, oemSheet.getLastRow() - 1, 1).getValues();
  var oemRowToDelete = null;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === 'INMP411ACEZ') { oemRowToDelete = i + 2; break; }
  }
  if (oemRowToDelete) {
    oemSheet.getRange(oemRowToDelete, 5).setValue('NO STK 7/3/2026');
    SpreadsheetApp.flush();
    oemSheet.deleteRow(oemRowToDelete);
    Logger.log('Stamped and deleted OEM INMP411ACEZ (row ' + oemRowToDelete + ')');
  } else {
    Logger.log('OEM INMP411ACEZ exact row not found (may already be deleted)');
  }
  // Stamp Forte row 3987
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var cell = forteSheet.getRange(3987, 11);
  cell.clearDataValidations();
  cell.setValue('NO STK - 7/3/2026');
  cell.setBackground('#000000');
  cell.setFontColor('#FFFFFF');
  cell.setFontWeight('bold');
  SpreadsheetApp.flush();
  Logger.log('Stamped Forte row 3987 INMP411ACEZ → NO STK - 7/3/2026');
}

// ONE-TIME — Run removeDavidNoStk_3x_Jul3()
// Three David no-stk emails processed manually 7/3/2026:
//   676008001 #3985 Stock Sold  → OEM row 19632, Forte row 3985
//   IS43TR16640C-125JBLI #3984 No stk → OEM row 89034, Forte row 3984
//   BAS4002ARPPE6327 #3988 no stk → OEM row 66966, Forte row 3988
// Drafts: r-7499660713323468487, r5672943263510560525, r-7812586594929483963
function removeDavidNoStk_3x_Jul3() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var today = '7/3/2026';

  // OEM rows — stamp col E then delete (process in descending order to preserve row numbers)
  var oemRows = [
    {row: 89034, mpn: 'IS43TR16640C-125JBLI'},
    {row: 66966, mpn: 'BAS4002ARPPE6327'},
    {row: 19632, mpn: '676008001'},
  ].sort(function(a,b){ return b.row - a.row; });
  oemRows.forEach(function(r) {
    oemSheet.getRange(r.row, 5).setValue('NO STK ' + today);
    SpreadsheetApp.flush();
    oemSheet.deleteRow(r.row);
    Logger.log('Stamped + deleted OEM row ' + r.row + ' (' + r.mpn + ')');
  });

  // Forte rows — stamp col K black NO STK
  [{row:3985,mpn:'676008001'},{row:3984,mpn:'IS43TR16640C-125JBLI'},{row:3988,mpn:'BAS4002ARPPE6327'}].forEach(function(r){
    var cell = forteSheet.getRange(r.row, 11);
    cell.clearDataValidations();
    cell.setValue('NO STK - ' + today);
    cell.setBackground('#000000');
    cell.setFontColor('#FFFFFF');
    cell.setFontWeight('bold');
    Logger.log('Stamped Forte row ' + r.row + ' ' + r.mpn);
  });
  SpreadsheetApp.flush();
  Logger.log('removeDavidNoStk_3x_Jul3 complete');
}

// ONE-TIME — Run addForte_PVA1OAH21_Jul3()
// LAYTEC DESIGN & CONSULTING INC (ikalman@laytec.com) — website RFQ 7/3/2026
// PVA1OAH21.2NV2;PVA1OAH2SNA, 2200 qty, TP $0.60, Canada
// Automation missed: semicolon in MPN broke extractor → empty oem_results → no_bid
// msg_checking draft r2099275116509180978 created manually.
function addForte_PVA1OAH21_Jul3() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['7/3/2026', 'PVA1OAH21.2NV2;PVA1OAH2SNA', 2200, 0.60, '', 'CA',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added Forte: PVA1OAH21.2NV2;PVA1OAH2SNA qty=2200 TP=0.60 CA');
}

// ONE-TIME — Run removeOem_AD8630ARUZREEL_Jul3()
// David reported AD8630ARUZ-REEL #3983 no stk on 7/3/2026.
// OEM EXCESS has TWO rows (62416 + 62417) — deletePart returned FUZZY_REVIEW (>1 exact match),
// so automation never labeled the thread. Both rows stamped + deleted here.
// Forte row 3983 stamped black NO STK.
// Reply draft created manually (see project memory for draft ID).
function removeOem_AD8630ARUZREEL_Jul3() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var today = '7/3/2026';

  // Delete in descending order so row numbers stay valid
  [62417, 62416].forEach(function(row) {
    oemSheet.getRange(row, 5).setValue('NO STK ' + today);
    SpreadsheetApp.flush();
    oemSheet.deleteRow(row);
    Logger.log('Stamped + deleted OEM AD8630ARUZ-REEL row ' + row);
  });

  // Stamp Forte row 3983
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var cell = forteSheet.getRange(3983, 11);
  cell.clearDataValidations();
  cell.setValue('NO STK - ' + today);
  cell.setBackground('#000000');
  cell.setFontColor('#FFFFFF');
  cell.setFontWeight('bold');
  SpreadsheetApp.flush();
  Logger.log('Stamped Forte row 3983 AD8630ARUZ-REEL → NO STK - ' + today);
}

// ── One-time: remove oem-rfq-incoming-processed from David threads stuck by Trigger 3 bug ──
// Run once after pasting the fixed script. Trigger 7 will then pick them up on next 5-min cycle.
function unlabelStuckDavidThreads() {
  var label = GmailApp.getUserLabelByName('oem-rfq-incoming-processed');
  if (!label) { Logger.log('Label not found'); return; }
  var threads = GmailApp.search('from:david@fortetechno.com label:oem-rfq-incoming-processed -label:oem-agent-processed in:inbox', 0, 20);
  Logger.log('Found ' + threads.length + ' stuck David thread(s)');
  threads.forEach(function(t) {
    t.removeLabel(label);
    Logger.log('Unlabeled: ' + t.getFirstMessageSubject());
  });
  Logger.log('Done — Trigger 7 will process these on next run.');
}

// ── David no-stk: RP604K331A-TR #3975 (7/7/2026) ────────────────────────────
// Stamp OEM row 118520 NO STK, delete it, stamp Forte row 3975 black NO STK.
function removeOem_RP604K331ATR_3975() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var today = '7/7/2026';

  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  oemSheet.getRange(118520, 5).setValue('NO STK ' + today);
  SpreadsheetApp.flush();
  oemSheet.deleteRow(118520);
  SpreadsheetApp.flush();
  Logger.log('OEM row 118520 (RP604K331A-TR) stamped and deleted');

  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var cell = forteSheet.getRange(3975, 11);
  cell.clearDataValidations();
  cell.setValue('NO STK - ' + today);
  cell.setBackground('#000000');
  cell.setFontColor('#FFFFFF');
  cell.setFontWeight('bold');
  SpreadsheetApp.flush();
  Logger.log('Forte row 3975 stamped NO STK - ' + today);
}

// ── Add Forte entry: KSM26RS8/8HDI — Teresa/Motek, msg_checking sent 7/7/2026 ──
function addForte_KSM26RS8_Teresa_Jul7() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var today = '7/7/2026';
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow([today, 'KSM26RS8/8HDI', 30, 70, '', 'US',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added Forte row ' + nextRow + ': KSM26RS8/8HDI, qty 30, TP $70, US');
}

// ── Add Forte entry: 5EHM1S — Gabe/a2global, msg_checking sent 7/7/2026 ──
function addForte_5EHM1S_Gabe_Jul7() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var today = '7/7/2026';
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow([today, '5EHM1S', 250, 40, '', 'US',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added Forte row ' + nextRow + ': 5EHM1S, qty 250, TP $40, US');
}

// ── Add Forte entry: XP1001000-05R — Craig/Whistler, msg_checking sent 7/7/2026 ──
function addForte_XP1001000_Craig_Jul7() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var today = '7/7/2026';
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow([today, 'XP1001000-05R', 100, 28, '', 'GB',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added Forte row ' + nextRow + ': XP1001000-05R, qty 100, TP $28, GB');
}

// ── Add Forte entry: ISL83075EIBZA-T — Alisson/Chip1, msg_checking sent 7/7/2026 ──
function addForte_ISL83075_Alisson_Jul7() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var today = '7/7/2026';
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow([today, 'ISL83075EIBZA-T', 20000, 0.7, '', 'US',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added Forte row ' + nextRow + ': ISL83075EIBZA-T, qty 20000, TP $0.70, US');
}

// ── Add Forte entry: TMC2224-LA-T — Baalaji/Himil, msg_checking sent 7/7/2026 ──
function addForte_TMC2224_Baalaji_Jul7() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var today = '7/7/2026';
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow([today, 'TMC2224-LA-T', 1900, 1.5, '', 'IN',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added Forte row ' + nextRow + ': TMC2224-LA-T, qty 1900, TP $1.50, IN');
}

// ── Remove OEM EXCESS row — 926823-2 #3915, David no-stk 7/7/2026 ──
// OEM row 59256 (confirmed via web app — NOT 3915 which is the Forte row ref in subject)
function removeOem_9268232_3915() {
  var sheet = SpreadsheetApp.openById('1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4').getSheets()[0];
  sheet.getRange(59256, 5).setValue('NO STK 7/7/2026');
  sheet.deleteRow(59256);
  Logger.log('Removed OEM row 59256: 926823-2');
}

// ── Remove OEM EXCESS rows — David "Removing..." replies Jun 30–Jul 3 ──────────
// These MPNs are still in OEM EXCESS because John replied with "Removing [MPN] from
// OEM EXCESS" (not "Removed - MPN: [MPN]"), so Trigger 2's pattern never matched.
// Rows deleted highest-first. Forte stamped by MPN name (skips already-stamped rows).
// OEM rows:
//   VRF2933MP: 135471 | SAMSUNG CL05X226MR6NUW8: 120011 | S24SE05006PDFA: 119411
//   PM2120-330K-RC: 113580-113585 | MT53E1G32D2FW-046IT:B: 107005
//   FPF2700MPX: 83558 | EPCS16SI8: 80812-80813 | AL8860MP-13: 63633
function removeOemRows_MissingRemovals_Jun30Jul3() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];

  // Stamp then delete in descending order
  var oemRows = [
    { row: 135471, mpn: 'VRF2933MP',                  date: '6/30/2026' },
    { row: 120011, mpn: 'SAMSUNG CL05X226MR6NUW8',    date: '7/2/2026'  },
    { row: 119411, mpn: 'S24SE05006PDFA',              date: '7/1/2026'  },
    { row: 113585, mpn: 'PM2120-330K-RC',              date: '6/30/2026' },
    { row: 113584, mpn: 'PM2120-330K-RC',              date: '6/30/2026' },
    { row: 113583, mpn: 'PM2120-330K-RC',              date: '6/30/2026' },
    { row: 113582, mpn: 'PM2120-330K-RC',              date: '6/30/2026' },
    { row: 113581, mpn: 'PM2120-330K-RC',              date: '6/30/2026' },
    { row: 113580, mpn: 'PM2120-330K-RC',              date: '6/30/2026' },
    { row: 107005, mpn: 'MT53E1G32D2FW-046IT:B',       date: '7/3/2026'  },
    { row: 83558,  mpn: 'FPF2700MPX',                  date: '7/1/2026'  },
    { row: 80813,  mpn: 'EPCS16SI8',                   date: '6/30/2026' },
    { row: 80812,  mpn: 'EPCS16SI8',                   date: '6/30/2026' },
    { row: 63633,  mpn: 'AL8860MP-13',                 date: '7/1/2026'  },
  ];
  oemRows.forEach(function(r) {
    oemSheet.getRange(r.row, 5).setValue('NO STK ' + r.date);
    SpreadsheetApp.flush();
    oemSheet.deleteRow(r.row);
    Logger.log('Stamped + deleted OEM row ' + r.row + ' (' + r.mpn + ')');
  });

  // Stamp Forte by MPN name — skips rows already marked NO STK
  var forteStamps = [
    { mpn: 'AL8860MP-13',              date: '7/1/2026'  },
    { mpn: 'SAMSUNG CL05X226MR6NUW8', date: '7/2/2026'  },
    { mpn: 'VRF2933MP',               date: '6/30/2026' },
    { mpn: 'PM2120-330K-RC',          date: '6/30/2026' },
    { mpn: 'PM2120330KRC',            date: '6/30/2026' },
    { mpn: 'EPCS16SI8',               date: '6/30/2026' },
    { mpn: 'FPF2700MPX',              date: '7/1/2026'  },
    { mpn: 'S24SE05006PDFA',          date: '7/1/2026'  },
    { mpn: 'MT53E1G32D2FW-046IT:B',   date: '7/3/2026'  },
    { mpn: 'MT53E1G32D2FW-046 IT:B TR', date: '7/3/2026' },
  ];
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var data = forteSheet.getDataRange().getValues();
  var mpnSet = {};
  forteStamps.forEach(function(s) { mpnSet[s.mpn.toUpperCase()] = s.date; });
  var stamped = 0;
  for (var i = 1; i < data.length; i++) {
    var mpn = String(data[i][1]).trim();
    var status = String(data[i][10] || '').trim();
    var noStkDate = mpnSet[mpn.toUpperCase()];
    if (!noStkDate) continue;
    if (status.toUpperCase().indexOf('NO STK') >= 0 || status.toUpperCase() === 'CLOSED') continue;
    var cell = forteSheet.getRange(i + 1, 11);
    cell.clearDataValidations();
    cell.setValue('NO STK - ' + noStkDate);
    cell.setBackground('#000000');
    cell.setFontColor('#FFFFFF');
    cell.setFontWeight('bold');
    Logger.log('Forte row ' + (i + 1) + ': ' + mpn + ' → NO STK - ' + noStkDate);
    stamped++;
  }
  SpreadsheetApp.flush();
  Logger.log('removeOemRows_MissingRemovals_Jun30Jul3 complete — OEM rows: ' + oemRows.length + ', Forte rows stamped: ' + stamped);
}

// ── Remove OEM EXCESS rows — David no-stk replies, 7/7/2026 batch ──
// Run AFTER sending all 8 "Ok, removed from listing" replies.
// Row numbers confirmed via web app — the #XXXX in David's subject are Forte refs, NOT OEM rows.
// TPS23861PWR has TWO OEM rows. OEM rows deleted highest-first. Forte stamped by MPN name.
function removeOemRows_DavidNoStk_Jul7() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';

  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var oemRows = [
    { row: 136784, mpn: 'XCZU27DR-2FFVE1156I' },
    { row: 135261, mpn: 'VIPER26HDTR' },
    { row: 132506, mpn: 'TPW1R306PLL1QM' },
    { row: 131148, mpn: 'TPS23861PWR' },
    { row: 131147, mpn: 'TPS23861PWR' },
    { row: 106638, mpn: 'MT25QU02GCBB8E12-0AATTR' },
    { row: 96636,  mpn: 'LTC7150SJY-4#PBF' },
    { row: 76726,  mpn: 'DD-00429VP-200' },
    { row: 62092,  mpn: 'AD633JRZ' },
  ];
  oemRows.forEach(function(r) {
    oemSheet.getRange(r.row, 5).setValue('NO STK 7/7/2026');
    oemSheet.deleteRow(r.row);
    Logger.log('Removed OEM row ' + r.row + ': ' + r.mpn);
  });
  SpreadsheetApp.flush();

  // Stamp Forte by MPN name (search-based, safe against row-number shifts)
  var forteStamps = [
    'XCZU27DR-2FFVE1156I', 'VIPER26HDTR', 'TPW1R306PLL1QM', 'TPS23861PWR',
    'MT25QU02GCBB8E12-0AATTR', 'LTC7150SJY-4#PBF', 'DD-00429VP-200', 'AD633JRZ'
  ];
  var mpnSet = {};
  forteStamps.forEach(function(m) { mpnSet[m.toUpperCase()] = true; });
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var data = forteSheet.getDataRange().getValues();
  var stamped = 0;
  for (var i = 1; i < data.length; i++) {
    var mpn = String(data[i][1]).trim();
    var status = String(data[i][10] || '').trim();
    if (!mpnSet[mpn.toUpperCase()]) continue;
    if (status.toUpperCase().indexOf('NO STK') >= 0 || status.toUpperCase() === 'CLOSED') continue;
    var cell = forteSheet.getRange(i + 1, 11);
    cell.clearDataValidations();
    cell.setValue('NO STK - 7/7/2026');
    cell.setBackground('#000000');
    cell.setFontColor('#FFFFFF');
    cell.setFontWeight('bold');
    Logger.log('Forte row ' + (i + 1) + ': ' + mpn + ' → NO STK - 7/7/2026');
    stamped++;
  }
  SpreadsheetApp.flush();
  Logger.log('removeOemRows_DavidNoStk_Jul7 complete — OEM rows deleted: ' + oemRows.length + ', Forte rows stamped: ' + stamped);
}

// ── Audit OEM EXCESS for ICS upload quality issues ──────────────────────────
// Run to diagnose which rows are being rejected by ICS (post@icsource.com).
// Logs: rows with blank/non-numeric QTY, rows with whitespace in MPN,
// blank trailing rows, and qty values stored as comma-text.
// Does NOT modify the sheet — read-only diagnostic.
function auditOemExcessICSQuality() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var sheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var totalRows = data.length - 1; // exclude header
  var nullQty = [], commaTextQty = [], blankMpn = [], messyMpn = [];

  for (var i = 1; i < data.length; i++) {
    var mpn = String(data[i][0]);
    var mpnTrimmed = mpn.trim();
    var qty = data[i][3];
    var sheetRow = i + 1;

    if (!mpnTrimmed) {
      blankMpn.push(sheetRow);
      continue;
    }
    if (mpn !== mpnTrimmed) {
      messyMpn.push({ row: sheetRow, mpn: JSON.stringify(mpn) });
    }

    if (qty === '' || qty === null || qty === undefined) {
      nullQty.push({ row: sheetRow, mpn: mpnTrimmed, qty: '(empty)' });
    } else if (typeof qty === 'string' && qty.indexOf(',') >= 0) {
      var num = parseFloat(qty.replace(/,/g, ''));
      if (!isNaN(num)) {
        commaTextQty.push({ row: sheetRow, mpn: mpnTrimmed, qty: qty });
      } else {
        nullQty.push({ row: sheetRow, mpn: mpnTrimmed, qty: qty + ' (non-numeric)' });
      }
    } else {
      var n = parseFloat(String(qty).replace(/,/g, ''));
      if (isNaN(n) || n <= 0) {
        nullQty.push({ row: sheetRow, mpn: mpnTrimmed, qty: String(qty) });
      }
    }
  }

  Logger.log('=== OEM EXCESS ICS Quality Audit ===');
  Logger.log('Total data rows: ' + totalRows);
  Logger.log('Rows that will be REJECTED by ICS (null/non-numeric qty): ' + nullQty.length);
  nullQty.forEach(function(r) { Logger.log('  Row ' + r.row + ': ' + r.mpn + ' | QTY=' + r.qty); });
  Logger.log('Rows with comma-text qty (converted to number during upload): ' + commaTextQty.length);
  if (commaTextQty.length <= 20) commaTextQty.forEach(function(r) { Logger.log('  Row ' + r.row + ': ' + r.mpn + ' | QTY=' + r.qty); });
  Logger.log('Blank MPN rows (trailing/phantom rows): ' + blankMpn.length);
  if (blankMpn.length) Logger.log('  Rows: ' + blankMpn.join(', '));
  Logger.log('Rows with whitespace in MPN: ' + messyMpn.length);
  messyMpn.forEach(function(r) { Logger.log('  Row ' + r.row + ': ' + r.mpn); });
  Logger.log('=== End Audit ===');
}

// ── Delete blank trailing rows and clean MPN whitespace in OEM EXCESS ───────
// Run once after auditOemExcessICSQuality() to fix structural data quality.
// - Deletes rows where FullPartNumber (col A) is completely blank
// - Strips leading/trailing whitespace and tab characters from MPN values
// Does NOT add or remove qty values — ESR5-NO-31-24VAC-DC still needs a qty entered manually.
function cleanOemExcessStructure() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var sheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var blankRows = [], messyMpnRows = [];

  for (var i = data.length - 1; i >= 1; i--) { // bottom-up for safe deletion
    var mpn = String(data[i][0]);
    var mpnTrimmed = mpn.trim();
    if (!mpnTrimmed) {
      blankRows.push(i + 1);
      sheet.deleteRow(i + 1);
      Logger.log('Deleted blank row ' + (i + 1));
    } else if (mpn !== mpnTrimmed) {
      messyMpnRows.push(i + 1);
      sheet.getRange(i + 1, 1).setValue(mpnTrimmed);
      Logger.log('Cleaned MPN whitespace at row ' + (i + 1) + ': ' + JSON.stringify(mpn) + ' → ' + mpnTrimmed);
    }
  }
  SpreadsheetApp.flush();
  Logger.log('cleanOemExcessStructure complete — blank rows deleted: ' + blankRows.length + ', MPN whitespace fixed: ' + messyMpnRows.length);
  Logger.log('NOTE: ESR5-NO-31-24VAC-DC still needs a QTY value entered manually in the sheet.');
}

// ONE-TIME — Run after sending draft r-[id] to David.
// Removes SLI-343P8G3F from OEM EXCESS (row 122060) and stamps Forte.
// David email: "SLI-343P8G3F #4010  No stock" (2026-07-07)
function removeOem_SLI343P8G3F_David_Jul7() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var today = '7/7/2026';
  var noStkStamp = 'NO STK ' + today;

  // 1. Delete OEM EXCESS row 122060 (SLI-343P8G3F, qty 617)
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var oemRow = oemSheet.getRange(122060, 1, 1, 5).getValues()[0];
  Logger.log('OEM row 122060 before delete: ' + JSON.stringify(oemRow));
  if (String(oemRow[0]).trim().toUpperCase() !== 'SLI-343P8G3F') {
    Logger.log('ERROR: row 122060 MPN mismatch — expected SLI-343P8G3F, got ' + oemRow[0]);
    return;
  }
  oemSheet.getRange(122060, 5).setValue(noStkStamp);
  oemSheet.deleteRow(122060);
  Logger.log('Deleted OEM row 122060 (SLI-343P8G3F)');

  // 2. Stamp Forte — find all open SLI-343P8G3F rows and mark NO STK
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var data = forteSheet.getDataRange().getValues();
  var stamped = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toUpperCase() === 'SLI-343P8G3F' &&
        String(data[i][10]).trim().toUpperCase() !== 'CLOSED') {
      var cell = forteSheet.getRange(i + 1, 11); // col K = Status
      cell.setValue('NO STK - ' + today);
      cell.setBackground('#000000');
      cell.setFontColor('#FFFFFF');
      cell.setFontWeight('bold');
      Logger.log('Stamped Forte row ' + (i + 1) + ' (SLI-343P8G3F)');
      stamped++;
    }
  }
  Logger.log('removeOem_SLI343P8G3F_David_Jul7 complete — Forte rows stamped: ' + stamped);
}

// ONE-TIME — Run after sending draft r1787994097531601685 to David.
// Removes 7289-7643-30 from OEM EXCESS (row 54289) and stamps Forte.
// David email: "7289-7643-30 #3913  No stk" (2026-07-07)
function removeOem_728976433_David_Jul7() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var today = '7/7/2026';
  var noStkStamp = 'NO STK ' + today;

  // 1. Delete OEM EXCESS row 54289 (7289-7643-30, qty 1335)
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var oemRow = oemSheet.getRange(54289, 1, 1, 5).getValues()[0];
  Logger.log('OEM row 54289 before delete: ' + JSON.stringify(oemRow));
  if (String(oemRow[0]).trim().toUpperCase() !== '7289-7643-30') {
    Logger.log('ERROR: row 54289 MPN mismatch — expected 7289-7643-30, got ' + oemRow[0]);
    return;
  }
  oemSheet.getRange(54289, 5).setValue(noStkStamp);
  oemSheet.deleteRow(54289);
  Logger.log('Deleted OEM row 54289 (7289-7643-30)');

  // 2. Stamp Forte row #3913 and any other open rows for this MPN
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var data = forteSheet.getDataRange().getValues();
  var stamped = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toUpperCase() === '7289-7643-30' &&
        String(data[i][10]).trim().toUpperCase() !== 'CLOSED') {
      var cell = forteSheet.getRange(i + 1, 11);
      cell.setValue('NO STK - ' + today);
      cell.setBackground('#000000');
      cell.setFontColor('#FFFFFF');
      cell.setFontWeight('bold');
      Logger.log('Stamped Forte row ' + (i + 1) + ' (7289-7643-30)');
      stamped++;
    }
  }
  Logger.log('removeOem_728976433_David_Jul7 complete — Forte rows stamped: ' + stamped);
}

// ONE-TIME — Run after sending drafts for Jul 8 David no-stk batch.
// IRF7241TRPBF #4001 (row 88414), W25Q256JWEIM #4014 (rows 135762+135763), FTLF1519P1BCL #3853 (row 83848)
function removeOemRows_DavidNoStk_Jul8() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var today = '7/8/2026';
  var noStkStamp = 'NO STK ' + today;
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];

  // Delete in descending row order to avoid row-shift errors
  var oemRows = [
    { row: 135763, mpn: 'W25Q256JWEIM' },
    { row: 135762, mpn: 'W25Q256JWEIM' },
    { row: 88414,  mpn: 'IRF7241TRPBF' },
    { row: 83848,  mpn: 'FTLF1519P1BCL' }
  ];

  oemRows.forEach(function(item) {
    var rowData = oemSheet.getRange(item.row, 1, 1, 5).getValues()[0];
    Logger.log('OEM row ' + item.row + ' before delete: ' + JSON.stringify(rowData));
    if (String(rowData[0]).trim().toUpperCase() !== item.mpn.toUpperCase()) {
      Logger.log('ERROR: row ' + item.row + ' MPN mismatch — expected ' + item.mpn + ', got ' + rowData[0]);
      return;
    }
    oemSheet.getRange(item.row, 5).setValue(noStkStamp);
    oemSheet.deleteRow(item.row);
    Logger.log('Deleted OEM row ' + item.row + ' (' + item.mpn + ')');
  });

  // Stamp Forte by MPN name for all 3 parts
  var mpnsToStamp = ['IRF7241TRPBF', 'W25Q256JWEIM', 'FTLF1519P1BCL'];
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var data = forteSheet.getDataRange().getValues();
  var stamped = 0;
  for (var i = 1; i < data.length; i++) {
    var rowMpn = String(data[i][1]).trim().toUpperCase();
    if (mpnsToStamp.indexOf(rowMpn) >= 0 && String(data[i][10]).trim().toUpperCase() !== 'CLOSED') {
      var cell = forteSheet.getRange(i + 1, 11);
      cell.setValue('NO STK - ' + today);
      cell.setBackground('#000000');
      cell.setFontColor('#FFFFFF');
      cell.setFontWeight('bold');
      Logger.log('Stamped Forte row ' + (i + 1) + ' (' + data[i][1] + ')');
      stamped++;
    }
  }
  Logger.log('removeOemRows_DavidNoStk_Jul8 complete — Forte rows stamped: ' + stamped);
}

// ONE-TIME — Run after sending Jul 9 David no-stk drafts.
// MP9100-75.0-1 #4028 (row 105424), STM32G474RET6 #4025 (rows 127094+127093+126915),
// TPSI2140QDWQRQ1 #4024 (row 132436), AD8676ARMZ-REEL #4022 (row 62437)
function removeOemRows_DavidNoStk_Jul9() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var today = '7/9/2026';
  var noStkStamp = 'NO STK ' + today;
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];

  // Descending order to avoid row-shift errors
  var oemRows = [
    { row: 132436, mpn: 'TPSI2140QDWQRQ1' },
    { row: 127094, mpn: 'STM32G474RET6' },
    { row: 127093, mpn: 'STM32G474RET6' },
    { row: 126915, mpn: 'STM32G474RET6' },
    { row: 105424, mpn: 'MP9100-75.0-1' },
    { row: 62437,  mpn: 'AD8676ARMZ-REEL' }
  ];

  oemRows.forEach(function(item) {
    var rowData = oemSheet.getRange(item.row, 1, 1, 5).getValues()[0];
    Logger.log('OEM row ' + item.row + ': ' + JSON.stringify(rowData));
    if (String(rowData[0]).trim().toUpperCase() !== item.mpn.toUpperCase()) {
      Logger.log('ERROR: row ' + item.row + ' mismatch — expected ' + item.mpn + ', got ' + rowData[0]);
      return;
    }
    oemSheet.getRange(item.row, 5).setValue(noStkStamp);
    oemSheet.deleteRow(item.row);
    Logger.log('Deleted OEM row ' + item.row + ' (' + item.mpn + ')');
  });

  // Stamp Forte for all 4 MPNs
  var mpnsToStamp = ['MP9100-75.0-1', 'STM32G474RET6', 'TPSI2140QDWQRQ1', 'AD8676ARMZ-REEL'];
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var data = forteSheet.getDataRange().getValues();
  var stamped = 0;
  for (var i = 1; i < data.length; i++) {
    var rowMpn = String(data[i][1]).trim().toUpperCase();
    if (mpnsToStamp.indexOf(rowMpn) >= 0 && String(data[i][10]).trim().toUpperCase() !== 'CLOSED') {
      var cell = forteSheet.getRange(i + 1, 11);
      cell.setValue('NO STK - ' + today);
      cell.setBackground('#000000');
      cell.setFontColor('#FFFFFF');
      cell.setFontWeight('bold');
      Logger.log('Stamped Forte row ' + (i + 1) + ' (' + data[i][1] + ')');
      stamped++;
    }
  }
  Logger.log('removeOemRows_DavidNoStk_Jul9 complete — Forte rows stamped: ' + stamped);
}

// ONE-TIME — Run after sending Jul 9 (batch 2) David no-stk drafts.
// PIC32MZ2048EFH100-I/PT #4017 no stk  (row 113267)  → draft r-6505310211248415749
// NFL18ST506H1A3D #4026 Cant share      (row 108875)  → draft r9066069716555901252
// HSC100800RJ #4023 No stock            (row 86770)   → draft r57137538588965595
function removeOemRows_DavidNoStk_Jul9b() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var today = '7/9/2026';
  var noStkStamp = 'NO STK ' + today;
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];

  // Descending order to avoid row-shift errors
  var oemRows = [
    { row: 113267, mpn: 'PIC32MZ2048EFH100-I/PT' },
    { row: 108875, mpn: 'NFL18ST506H1A3D' },
    { row: 86770,  mpn: 'HSC100800RJ' }
  ];

  oemRows.forEach(function(item) {
    var rowData = oemSheet.getRange(item.row, 1, 1, 5).getValues()[0];
    Logger.log('OEM row ' + item.row + ': ' + JSON.stringify(rowData));
    oemSheet.getRange(item.row, 5).setValue(noStkStamp);
    oemSheet.deleteRow(item.row);
    Logger.log('Deleted OEM row ' + item.row + ' (' + item.mpn + ')');
  });

  // Stamp Forte by MPN name for all 3 parts
  var mpnsToStamp = ['PIC32MZ2048EFH100-I/PT', 'NFL18ST506H1A3D', 'HSC100800RJ'];
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var data = forteSheet.getDataRange().getValues();
  var stamped = 0;
  for (var i = 1; i < data.length; i++) {
    var rowMpn = String(data[i][1]).trim().toUpperCase();
    if (mpnsToStamp.indexOf(rowMpn) >= 0 && String(data[i][10]).trim().toUpperCase() !== 'CLOSED') {
      var cell = forteSheet.getRange(i + 1, 11);
      cell.clearDataValidations();
      cell.setValue('NO STK - ' + today);
      cell.setBackground('#000000');
      cell.setFontColor('#FFFFFF');
      cell.setFontWeight('bold');
      Logger.log('Stamped Forte row ' + (i + 1) + ' (' + data[i][1] + ')');
      stamped++;
    }
  }
  SpreadsheetApp.flush();
  Logger.log('removeOemRows_DavidNoStk_Jul9b complete — Forte rows stamped: ' + stamped);
}

// ONE-TIME — Run deleteWrongDraft_S40FC004() to remove the wrong Bill-removal draft
// created with "Done ? S40FC004C1B1I00010 removed from OEM EXCESS." (em dash encoding bug).
// Replacement draft r4170294018491356103 already created with correct wording.
// Part already removed from OEM EXCESS — just delete the bad draft and send the new one.
function deleteWrongDraft_S40FC004() {
  var token = ScriptApp.getOAuthToken();
  try {
    UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/me/drafts/r-1086416931308096189', {
      method: 'delete',
      headers: { 'Authorization': 'Bearer ' + token }
    });
    Logger.log('Deleted wrong Bill-removal draft r-1086416931308096189 for S40FC004C1B1I00010');
  } catch(e) {
    Logger.log('Draft already gone or error: ' + e);
  }
}

// ONE-TIME — Run removeOemScan_W25Q256JWEIM() to delete OEM EXCESS rows for W25Q256JWEIM/JWEIMS.
// The one-time Jul8 function skips these because the sheet has "W25Q256JWEIMS" (trailing S)
// while the expected MPN is "W25Q256JWEIM" — verification check fails.
// This scan version matches both spellings and deletes all matching rows safely.
function removeOemScan_W25Q256JWEIM() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var sheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var today = '7/9/2026';
  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var mpn = String(data[i][0]).trim().toUpperCase();
    if (mpn === 'W25Q256JWEIM' || mpn === 'W25Q256JWEIMS') {
      sheet.getRange(i + 1, 5).setValue('NO STK ' + today);
      rowsToDelete.push(i + 1);
      Logger.log('Marked row ' + (i + 1) + ': ' + data[i][0]);
    }
  }
  SpreadsheetApp.flush();
  rowsToDelete.forEach(function(r) { sheet.deleteRow(r); Logger.log('Deleted OEM row ' + r); });
  Logger.log('removeOemScan_W25Q256JWEIM complete — ' + rowsToDelete.length + ' row(s) deleted');
}

// ONE-TIME — Run fixWrongNoBid_29537331() to:
//   1. Delete wrong no_bid draft r393182235757971220 (Next Wiring Harness Tech / Suresh Ravi)
//   2. Add Forte entry: 29537331, qty=2586, TP=0.52, IN (India)
// Bug: Haiku misread "$500 MIN TP REQUIRED" in description as a per-unit floor → chose no_bid.
// Correct msg_checking draft r2185140518700090261 already created.
// Fix applied in auditAndCorrect: no_bid with inventory present now gets Sonnet second-check.
function fixWrongNoBid_29537331() {
  // Delete wrong no_bid draft
  var token = ScriptApp.getOAuthToken();
  try {
    UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/me/drafts/r393182235757971220', {
      method: 'delete',
      headers: { 'Authorization': 'Bearer ' + token }
    });
    Logger.log('Deleted wrong no_bid draft r393182235757971220 for 29537331');
  } catch(e) {
    Logger.log('Draft r393182235757971220 already gone or error: ' + e);
  }
  // Add Forte entry
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['7/9/2026', '29537331', 2586, 0.52, '', 'IN',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added 29537331 to Forte row ' + nextRow + ' (qty=2586, TP=$0.52, IN)');
}

// ONE-TIME — Run deleteWrongDraft_ATH35088736() to remove the wrong request_tp_500 draft
// created for Anthony Maida / ATH Electronics RFQ on 35088736 (Aptiv, our own IN STOCK).
// Bug: all-digit MPN bypassed IN STOCK lookup → wrong request_tp_500 instead of own_stock.
// Correct own_stock draft r5040325320735075191 already created ("108 pcs at $10.00 each").
function deleteWrongDraft_ATH35088736() {
  var token = ScriptApp.getOAuthToken();
  UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/me/drafts/r1503123174137764597', {
    method: 'delete',
    headers: { 'Authorization': 'Bearer ' + token }
  });
  Logger.log('Deleted wrong request_tp_500 draft r1503123174137764597 for ATH Electronics 35088736');
}

// ONE-TIME — Run deleteWrongDrafts_HMC1031_Bernardo() to remove both wrong drafts created when
// automation incorrectly processed an email addressed to bernardo.moreno@intransittech.com.
// Bug: forwarding delivered it into John's inbox; TO-field guard fix now prevents this.
// Draft r5534143167725617478 = wrong request_tp_500 to michaelg@aedelectronics.com (HMC1031MS8ETR)
// Draft r1503123174137764597 = wrong request_tp_500 to anthony.maida@athelectronics.com (35088736)
function deleteWrongDrafts_HMC1031_Bernardo() {
  var token = ScriptApp.getOAuthToken();
  ['r5534143167725617478', 'r1503123174137764597'].forEach(function(draftId) {
    try {
      UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/me/drafts/' + draftId, {
        method: 'delete',
        headers: { 'Authorization': 'Bearer ' + token }
      });
      Logger.log('Deleted draft ' + draftId);
    } catch(e) {
      Logger.log('Draft ' + draftId + ' already gone or error: ' + e);
    }
  });
}

// ONE-TIME — Run removeOem_900NVIDIA_BillRemoval() to delete the 7 OEM EXCESS rows
// for 900-13448-0020-000 (BILL EXT 117, rows 58386-58392).
// Bill said "not available" in thread 19f3718b959d0743 but Trigger 8 missed it because
// extractMPN() requires letters and "900-13448-0020-000" is all digits+dashes.
// The script fix (raw-subject fallback) prevents this for future occurrences.
function removeOem_900NVIDIA_BillRemoval() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var sheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  // Delete rows in reverse order so row numbers stay valid after each deletion
  var rows = [58392, 58391, 58390, 58389, 58388, 58387, 58386];
  rows.forEach(function(r) {
    sheet.deleteRow(r);
    Logger.log('Deleted OEM row ' + r + ' (900-13448-0020-000)');
  });
  SpreadsheetApp.flush();
  Logger.log('removeOem_900NVIDIA_BillRemoval complete — 7 rows deleted');
}

// ONE-TIME — Run addForte_558527_1() to add Forte entry for 558527-1 (Rebound Singapore)
// Buyer: vivien.teo@reboundeu.com | QTY: 80 | TP: $7 | Country: SG
// MSG_CHECKING draft r1744644503456409158 created and pending send by John
function addForte_558527_1() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var mpn = '558527-1';
  var today = '7/9/2026';
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow([today, mpn, 80, '$7', '', 'SG', '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added to Forte: 558527-1 | QTY: 80 | TP: $7 | SG');
}

// ONE-TIME — Run addForte_STM32F407VGT6_Jul9() to add Forte entry for STM32F407VGT6.
// Buyer: Li Ling (lilingseniorsales@gmail.com, CN) | QTY: 915 | TP: $1.50
// MSG_CHECKING draft r6393362527466539506 created 7/9/2026.
// Prior Forte entries: row 186 (3/21/2023, CLOSED), row 2468 (6/10/2025) — both > 60 days ago.
function addForte_STM32F407VGT6_Jul9() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['7/9/2026', 'STM32F407VGT6', 915, 1.50, '', 'CN',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added STM32F407VGT6 to Forte row ' + nextRow + ' (qty=915, TP=$1.50, CN)');
}

// ONE-TIME — Run addForte_MT35XU02GCBA1G12_Jul9() to add Forte entry for MT35XU02GCBA1G12-0SIT.
// Buyer: emma@liyijing.com.cn (CN) | QTY: 96 | TP: $40
// MSG_CHECKING draft r4640866340950046105 created 7/9/2026.
// No prior Forte entries for this MPN.
function addForte_MT35XU02GCBA1G12_Jul9() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['7/9/2026', 'MT35XU02GCBA1G12-0SIT', 96, 40, '', 'CN',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added MT35XU02GCBA1G12-0SIT to Forte row ' + nextRow + ' (qty=96, TP=$40, CN)');
}

// ONE-TIME — Run addForte_XC7A200T_Jul9() to add missed msg_checking Forte entry.
// Buyer: fan guibing (guibing@sz-xjsj.com.cn, SZ Luohu Dist. Xianjie Elec., CN)
// Worker sent wrong request_tp_2000 (couldn't parse TgtPrice=11 from netCOMPONENTS table).
// Corrective msg_checking draft r6888179817342035228 already created (reply to thread 19f41f1f120b991f).
// 60-day check: last Forte entry Dec 2024 — outside 60 days, safe to add.
function addForte_XC7A200T_Jul9() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['7/9/2026', 'XC7A200T-2FBG676I', 100, 11, '', 'CN',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added XC7A200T-2FBG676I to Forte row ' + nextRow + ' (qty=100, TP=$11, CN)');
}

// ONE-TIME — Run addForte_BTS500551TMAATMA1_Jul10() after sending msg_checking r3787864537891541025.
// Buyer: Joe Tucarella (joe@ableelectronics.com, Able Electronics, Bellport NY, US)
// QtyReq=7000, TgtPrice=$3.50 ($24,500 line). OEM EXCESS row 86411 (7667 qty, not BILL EXT).
// No prior Forte entries. Draft created manually — trigger stalled after 12:16.
function addForte_BTS500551TMAATMA1_Jul10() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['7/10/2026', 'BTS500551TMAATMA1', 7000, 3.50, '', 'US',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added BTS500551TMAATMA1 to Forte row ' + nextRow + ' (qty=7000, TP=$3.50, US)');
}

// ONE-TIME — Run removeOem_LMX2594RHAT_4019_Jul10() after sending reply draft r-2182396785982369923.
// David email: "LMX2594RHAT #4019  No stk" (2026-07-10). OEM row 21560 (43 qty, TI).
// Forte row 4019: qty 500, TP $11, CN (Jul 8 2026 buyer RFQ).
function removeOem_LMX2594RHAT_4019_Jul10() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var today = '7/10/2026';

  // Verify and delete OEM row 21560
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var rowData = oemSheet.getRange(21560, 1, 1, 5).getValues()[0];
  Logger.log('OEM row 21560 before delete: ' + JSON.stringify(rowData));
  if (String(rowData[0]).trim().toUpperCase() !== 'LMX2594RHAT') {
    Logger.log('ERROR: row 21560 MPN mismatch — expected LMX2594RHAT, got ' + rowData[0]);
    return;
  }
  oemSheet.getRange(21560, 5).setValue('NO STK ' + today);
  oemSheet.deleteRow(21560);
  Logger.log('Stamped and deleted OEM row 21560 (LMX2594RHAT)');

  // Stamp Forte row 4019
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var cell = forteSheet.getRange(4019, 11);
  cell.clearDataValidations();
  cell.setValue('NO STK - ' + today);
  cell.setBackground('#000000');
  cell.setFontColor('#FFFFFF');
  cell.setFontWeight('bold');
  SpreadsheetApp.flush();
  Logger.log('Stamped Forte row 4019 (LMX2594RHAT) → NO STK - ' + today);
}

// ONE-TIME — Run removeOem_TPS82084SILR_4033_Jul10() AFTER sending draft r968649894916425707.
// David email: "TPS82084SILR #4033 no stk" (2026-07-10). OEM row 16750 (19536 qty, BILL EXT 117).
// Forte row 4033: qty 19536, TP $0.40, US.
function removeOem_TPS82084SILR_4033_Jul10() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var noStkDate = '7/10/2026';
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var rowData = oemSheet.getRange(16750, 1, 1, 5).getValues()[0];
  Logger.log('OEM row 16750 before delete: ' + JSON.stringify(rowData));
  if (String(rowData[0]).trim().toUpperCase() !== 'TPS82084SILR') {
    Logger.log('ERROR: row 16750 MPN mismatch — expected TPS82084SILR, got ' + rowData[0]); return;
  }
  oemSheet.getRange(16750, 5).setValue('NO STK ' + noStkDate);
  oemSheet.deleteRow(16750);
  Logger.log('Stamped and deleted OEM row 16750 (TPS82084SILR)');
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var fData = forteSheet.getRange(4033, 1, 1, 11).getValues()[0];
  Logger.log('Forte row 4033 check: ' + JSON.stringify(fData));
  if (String(fData[1]).trim().toUpperCase() !== 'TPS82084SILR') {
    Logger.log('ERROR: Forte row 4033 MPN mismatch — got ' + fData[1]); return;
  }
  var cell = forteSheet.getRange(4033, 11);
  cell.clearDataValidations(); cell.setValue('NO STK - ' + noStkDate);
  cell.setBackground('#000000'); cell.setFontColor('#FFFFFF'); cell.setFontWeight('bold');
  SpreadsheetApp.flush();
  Logger.log('Stamped Forte row 4033 (TPS82084SILR) → NO STK - ' + noStkDate);
}

// ONE-TIME — Run stampMissedNoStk_Jul9() to backfill OEM EXCESS removals and Forte NO STK stamps
// for 6 David no-stk emails from 2026-07-09 where replies were sent but sheets were never updated.
// Bug cause: executeDecision had no handler for david_nostock action — deletePart() was never called.
// Fixed in email_script_v24_hub.js. Run this function once to clean up the backlog.
// Entries (OEM deleted descending row order):
//   STM32G474RET6 → OEM 126379 / Forte 4025
//   PIC32MZ2048EFH100-I/PT → OEM 116920 / Forte 4017
//   NFL18ST506H1A3D → OEM 113704 / Forte 4026
//   MP9100-75.0-1 → OEM 111286 / Forte 4028
//   HSC100800RJ → OEM 98131 / Forte 4023
//   TPSI2140QDWQRQ1 → OEM 16765 / Forte 4024 (BILL EXT 117)
function stampMissedNoStk_Jul9() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var noStkDate = '7/9/2026';
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];

  // OEM entries sorted descending by row so each deletion doesn't shift subsequent rows
  var oemEntries = [
    { row: 126379, expectedMpn: 'STM32G474RET6' },
    { row: 116920, expectedMpn: 'PIC32MZ2048EFH100-I/PT' },
    { row: 113704, expectedMpn: 'NFL18ST506H1A3D' },
    { row: 111286, expectedMpn: 'MP9100-75.0-1' },
    { row: 98131,  expectedMpn: 'HSC100800RJ' },
    { row: 16765,  expectedMpn: 'TPSI2140QDWQRQ1' },
  ];
  oemEntries.forEach(function(e) {
    var rowData = oemSheet.getRange(e.row, 1, 1, 5).getValues()[0];
    Logger.log('OEM row ' + e.row + ': ' + JSON.stringify(rowData));
    if (String(rowData[0]).trim().toUpperCase() !== e.expectedMpn.toUpperCase()) {
      Logger.log('SKIP: row ' + e.row + ' MPN mismatch — expected ' + e.expectedMpn + ', got ' + rowData[0]);
      return;
    }
    oemSheet.getRange(e.row, 5).setValue('NO STK ' + noStkDate);
    oemSheet.deleteRow(e.row);
    Logger.log('Stamped and deleted OEM row ' + e.row + ' (' + e.expectedMpn + ')');
  });

  // Forte entries: stamp specific rows by row number with MPN verification
  var forteEntries = [
    { row: 4025, expectedMpn: 'STM32G474RET6' },
    { row: 4017, expectedMpn: 'PIC32MZ2048EFH100-I/PT' },
    { row: 4026, expectedMpn: 'NFL18ST506H1A3D' },
    { row: 4028, expectedMpn: 'MP9100-75.0-1' },
    { row: 4023, expectedMpn: 'HSC100800RJ' },
    { row: 4024, expectedMpn: 'TPSI2140QDWQRQ1' },
  ];
  forteEntries.forEach(function(e) {
    var fData = forteSheet.getRange(e.row, 1, 1, 11).getValues()[0];
    Logger.log('Forte row ' + e.row + ' check: MPN=' + fData[1]);
    if (String(fData[1]).trim().toUpperCase() !== e.expectedMpn.toUpperCase()) {
      Logger.log('SKIP: Forte row ' + e.row + ' MPN mismatch — expected ' + e.expectedMpn + ', got ' + fData[1]);
      return;
    }
    var cell = forteSheet.getRange(e.row, 11);
    cell.clearDataValidations(); cell.setValue('NO STK - ' + noStkDate);
    cell.setBackground('#000000'); cell.setFontColor('#FFFFFF'); cell.setFontWeight('bold');
    Logger.log('Stamped Forte row ' + e.row + ' (' + e.expectedMpn + ') → NO STK - ' + noStkDate);
  });
  SpreadsheetApp.flush();
  Logger.log('stampMissedNoStk_Jul9 complete');
}

// ONE-TIME — Run addForte_M306N4FGTFPUKJ_Jul10() after sending msg_checking r4073434427945463352.
// Buyer: Beatriz Griman (beatriz@chipnonstop.com, Whitestonebridge S.L., Spain, ES)
// QtyReq=100, TgtPrice=$30 ($3,000 line). OEM EXCESS row 105890 (174 qty, not BILL EXT).
// No prior Forte entries. Buyer replied TP $30 to John's request_tp reply.
function addForte_M306N4FGTFPUKJ_Jul10() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['7/10/2026', 'M306N4FGTFPUKJ', 100, 30, '', 'ES',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added M306N4FGTFPUKJ to Forte row ' + nextRow + ' (qty=100, TP=$30, ES)');
}

// ONE-TIME — Run addForte_EPCQ64SI16N_Jul10() after sending msg_checking draft r8004372012075735770.
// Buyer: Min Liu (liumin@zhongceco.com), Zhong Ce Electronic Technology, CN.
// QtyReq=100, TgtPrice=$16, line=$1,600. OEM row 93821 (100 qty Intel, $500 min, no BILL EXT).
// Prior Forte row 3649 (May 10, $12 TP) is 61 days old — outside 60-day window, add new entry.
function addForte_EPCQ64SI16N_Jul10() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['7/10/2026', 'EPCQ64SI16N', 100, 16, '', 'CN',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added EPCQ64SI16N to Forte row ' + nextRow + ' (qty=100, TP=$16, CN)');
}

// ONE-TIME — Run removeOem_STM32L476JGY3TR_4043_Jul10() AFTER sending draft r-2893020824272751311.
// David email: "STM32L476JGY3TR #4043  No stock" (2026-07-10). OEM row 126395 (7324 qty, ST MICRO).
// Forte row 4043: qty 1480, TP $3, CN.
function removeOem_STM32L476JGY3TR_4043_Jul10() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var noStkDate = '7/10/2026';
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var rowData = oemSheet.getRange(126395, 1, 1, 5).getValues()[0];
  Logger.log('OEM row 126395 before delete: ' + JSON.stringify(rowData));
  if (String(rowData[0]).trim().toUpperCase() !== 'STM32L476JGY3TR') {
    Logger.log('ERROR: row 126395 MPN mismatch — expected STM32L476JGY3TR, got ' + rowData[0]); return;
  }
  oemSheet.getRange(126395, 5).setValue('NO STK ' + noStkDate);
  oemSheet.deleteRow(126395);
  Logger.log('Stamped and deleted OEM row 126395 (STM32L476JGY3TR)');
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var fData = forteSheet.getRange(4043, 1, 1, 11).getValues()[0];
  Logger.log('Forte row 4043 check: MPN=' + fData[1]);
  if (String(fData[1]).trim().toUpperCase() !== 'STM32L476JGY3TR') {
    Logger.log('ERROR: Forte row 4043 MPN mismatch — got ' + fData[1]); return;
  }
  var cell = forteSheet.getRange(4043, 11);
  cell.clearDataValidations(); cell.setValue('NO STK - ' + noStkDate);
  cell.setBackground('#000000'); cell.setFontColor('#FFFFFF'); cell.setFontWeight('bold');
  SpreadsheetApp.flush();
  Logger.log('Stamped Forte row 4043 (STM32L476JGY3TR) → NO STK - ' + noStkDate);
}

// ONE-TIME — Run removeOem_DEI1072ASESG_Jul10() AFTER sending draft r8019405357433069600 (Bill).
// Bill confirmed "no longer available" for DEI1072A-SES-G (BILL EXT 117).
// OEM rows: 134999 (85 qty) + 134998 (71 qty) — deleted descending. No Forte entries (BILL EXT).
function removeOem_DEI1072ASESG_Jul10() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var noStkDate = '7/10/2026';
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var entries = [
    { row: 134999, expectedMpn: 'DEI1072A-SES-G' },
    { row: 134998, expectedMpn: 'DEI1072A-SES-G' },
  ];
  entries.forEach(function(e) {
    var rowData = oemSheet.getRange(e.row, 1, 1, 5).getValues()[0];
    Logger.log('OEM row ' + e.row + ': ' + JSON.stringify(rowData));
    if (String(rowData[0]).trim().toUpperCase() !== e.expectedMpn.toUpperCase()) {
      Logger.log('SKIP: row ' + e.row + ' MPN mismatch — expected ' + e.expectedMpn + ', got ' + rowData[0]);
      return;
    }
    oemSheet.getRange(e.row, 5).setValue('NO STK ' + noStkDate);
    oemSheet.deleteRow(e.row);
    Logger.log('Stamped and deleted OEM row ' + e.row + ' (DEI1072A-SES-G)');
  });
  SpreadsheetApp.flush();
  Logger.log('removeOem_DEI1072ASESG_Jul10 complete');
}

// ONE-TIME — Run removeOem_STM32H750VBT6TR_4041_Jul10() AFTER sending draft r-1568268353597923251.
// David: "STM32H750VBT6TR #4041 No stock". OEM row 126367 (3780 qty). Forte row 4041 Open.
function removeOem_STM32H750VBT6TR_4041_Jul10() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var noStkDate = '7/10/2026';
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var rowData = oemSheet.getRange(126367, 1, 1, 5).getValues()[0];
  Logger.log('OEM row 126367: ' + JSON.stringify(rowData));
  if (String(rowData[0]).trim().toUpperCase() !== 'STM32H750VBT6TR') {
    Logger.log('ERROR: MPN mismatch at row 126367 — got ' + rowData[0]); return;
  }
  oemSheet.getRange(126367, 5).setValue('NO STK ' + noStkDate);
  oemSheet.deleteRow(126367);
  Logger.log('Stamped and deleted OEM row 126367 (STM32H750VBT6TR)');
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var fData = forteSheet.getRange(4041, 1, 1, 11).getValues()[0];
  Logger.log('Forte row 4041 MPN: ' + fData[1]);
  if (String(fData[1]).trim().toUpperCase() !== 'STM32H750VBT6TR') {
    Logger.log('ERROR: Forte row 4041 MPN mismatch — got ' + fData[1]); return;
  }
  var cell = forteSheet.getRange(4041, 11);
  cell.clearDataValidations(); cell.setValue('NO STK - ' + noStkDate);
  cell.setBackground('#000000'); cell.setFontColor('#FFFFFF'); cell.setFontWeight('bold');
  SpreadsheetApp.flush();
  Logger.log('Stamped Forte row 4041 (STM32H750VBT6TR) → NO STK - ' + noStkDate);
}

// ONE-TIME — Run removeOem_ISL9R3060G2_4038_Jul10() AFTER sending draft r-7941055968615600676.
// David: "ISL9R3060G2 #4038 No stock". OEM row 99783 (16236 qty). Forte row 4038 Open.
function removeOem_ISL9R3060G2_4038_Jul10() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var noStkDate = '7/10/2026';
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var rowData = oemSheet.getRange(99783, 1, 1, 5).getValues()[0];
  Logger.log('OEM row 99783: ' + JSON.stringify(rowData));
  if (String(rowData[0]).trim().toUpperCase() !== 'ISL9R3060G2') {
    Logger.log('ERROR: MPN mismatch at row 99783 — got ' + rowData[0]); return;
  }
  oemSheet.getRange(99783, 5).setValue('NO STK ' + noStkDate);
  oemSheet.deleteRow(99783);
  Logger.log('Stamped and deleted OEM row 99783 (ISL9R3060G2)');
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var fData = forteSheet.getRange(4038, 1, 1, 11).getValues()[0];
  Logger.log('Forte row 4038 MPN: ' + fData[1]);
  if (String(fData[1]).trim().toUpperCase() !== 'ISL9R3060G2') {
    Logger.log('ERROR: Forte row 4038 MPN mismatch — got ' + fData[1]); return;
  }
  var cell = forteSheet.getRange(4038, 11);
  cell.clearDataValidations(); cell.setValue('NO STK - ' + noStkDate);
  cell.setBackground('#000000'); cell.setFontColor('#FFFFFF'); cell.setFontWeight('bold');
  SpreadsheetApp.flush();
  Logger.log('Stamped Forte row 4038 (ISL9R3060G2) → NO STK - ' + noStkDate);
}

// ONE-TIME — Run removeOem_ADA4940_4037_Jul10() AFTER sending draft r5566033739422784193.
// David: "ADA4940-1ACPZ-R7 #4037 No stock". OEM row 81137 (1185 qty). Forte row 4037 Open.
function removeOem_ADA4940_4037_Jul10() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var noStkDate = '7/10/2026';
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  var rowData = oemSheet.getRange(81137, 1, 1, 5).getValues()[0];
  Logger.log('OEM row 81137: ' + JSON.stringify(rowData));
  if (String(rowData[0]).trim().toUpperCase() !== 'ADA4940-1ACPZ-R7') {
    Logger.log('ERROR: MPN mismatch at row 81137 — got ' + rowData[0]); return;
  }
  oemSheet.getRange(81137, 5).setValue('NO STK ' + noStkDate);
  oemSheet.deleteRow(81137);
  Logger.log('Stamped and deleted OEM row 81137 (ADA4940-1ACPZ-R7)');
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var fData = forteSheet.getRange(4037, 1, 1, 11).getValues()[0];
  Logger.log('Forte row 4037 MPN: ' + fData[1]);
  if (String(fData[1]).trim().toUpperCase() !== 'ADA4940-1ACPZ-R7') {
    Logger.log('ERROR: Forte row 4037 MPN mismatch — got ' + fData[1]); return;
  }
  var cell = forteSheet.getRange(4037, 11);
  cell.clearDataValidations(); cell.setValue('NO STK - ' + noStkDate);
  cell.setBackground('#000000'); cell.setFontColor('#FFFFFF'); cell.setFontWeight('bold');
  SpreadsheetApp.flush();
  Logger.log('Stamped Forte row 4037 (ADA4940-1ACPZ-R7) → NO STK - ' + noStkDate);
}

// ONE-TIME — Run addForte_STM32G474RET6_Huiwei_Jul12() AFTER sending draft r-38241329826949335.
// RFQ: Jun Wang (wang@huiweielectronic.cn), Suzhou Huiwei Electronics, CN.
// 20834 qty, TP $1.75. OEM row 126355 (ST MICRO). Not BILL EXT.
// Forte row 4025 is a separate HK buyer ($3.60) — this is a different buyer, add new row.
function addForte_STM32G474RET6_Huiwei_Jul12() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var date = '7/12/2026';
  var mpn = 'STM32G474RET6';
  var qty = 20834;
  var tp = 1.75;
  var country = 'CN';
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow([date, mpn, qty, tp, '', country,
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  SpreadsheetApp.flush();
  Logger.log('Added STM32G474RET6 (Huiwei CN, 20834 qty, $1.75) to Forte row ' + nextRow);
}

// ONE-TIME — Run addForte_STM32H573VIT6_JunWang_Jul12() AFTER sending draft r6140960264817724470.
// RFQ: Jun Wang (wang@huiweielectronic.cn), Suzhou Huiwei Electronics, CN.
// 379222 qty, TP $1.00. OEM row 126359 (379222 qty). No existing Forte entry.
function addForte_STM32H573VIT6_JunWang_Jul12() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['7/12/2026', 'STM32H573VIT6', 379222, 1.00, '', 'CN',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  SpreadsheetApp.flush();
  Logger.log('Added STM32H573VIT6 (Jun Wang CN, 379222 qty, $1.00) to Forte row ' + nextRow);
}

// ONE-TIME — Run addForte_AP7361C33E13_Vinrox_Jul12() AFTER sending draft r2276491220951466190.
// RFQ: Vinrox Technologies Private Limited (info@vinrox.com), India.
// 10000 qty requested, TP $0.10. OEM row 82123 (7359 qty, Diodes Inc.).
// Buyer originally asked Jun 25, replied $0.10 on Jul 11. Total value = $1000.
function addForte_AP7361C33E13_Vinrox_Jul12() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['7/12/2026', 'AP7361C-33E-13', 10000, 0.10, '', 'IN',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  SpreadsheetApp.flush();
  Logger.log('Added AP7361C-33E-13 (Vinrox IN, 10000 qty, $0.10) to Forte row ' + nextRow);
}

// ONE-TIME — Run removeNoStk_STM32G031G6U6TR_Jul12() to clean up David no-stk.
// David said no-stk; John replied "Ok, removed from listing"; David replied "Thank you".
// OEM EXCESS rows to DELETE: 126339 (450000 qty) and 126220 (1308 qty).
// Forte row 4039: stamp col E = "NO STK 7/12/2026".
// Delete higher row first to avoid row number shifting.
function removeNoStk_STM32G031G6U6TR_Jul12() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';

  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  oemSheet.deleteRow(126339);
  Logger.log('Deleted OEM row 126339 (STM32G031G6U6TR 450000 qty)');
  oemSheet.deleteRow(126220);
  Logger.log('Deleted OEM row 126220 (STM32G031G6U6TR 1308 qty)');

  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  forteSheet.getRange(4039, 5).setValue('NO STK 7/12/2026');
  SpreadsheetApp.flush();
  Logger.log('Stamped Forte row 4039 (STM32G031G6U6TR) NO STK 7/12/2026');
}

// ONE-TIME — Run removeNoStk_EPCQ64SI16N_Jul12() to clean up David no-stk.
// David said no-stk; John replied "Ok, removed from listing"; David replied "Thank you".
// OEM EXCESS row to DELETE: 93821 (100 qty).
// Forte rows 3649 and 4045: stamp col E = "NO STK 7/12/2026".
function removeNoStk_EPCQ64SI16N_Jul12() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';

  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  oemSheet.deleteRow(93821);
  Logger.log('Deleted OEM row 93821 (EPCQ64SI16N 100 qty)');

  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  forteSheet.getRange(4045, 5).setValue('NO STK 7/12/2026');
  forteSheet.getRange(3649, 5).setValue('NO STK 7/12/2026');
  SpreadsheetApp.flush();
  Logger.log('Stamped Forte rows 3649 and 4045 (EPCQ64SI16N) NO STK 7/12/2026');
}

// ONE-TIME — Run removeNoStk_STM32H573VIT6_Jul13()
// David sent "STM32H573VIT6 #4048 no stk" on 2026-07-13.
// IMPORTANT: Run this BEFORE removeNoStk_STM32G031G6U6TR_Jul12 to avoid row-shift issues.
// OEM EXCESS row to DELETE: 126359 (379222 qty) — matches Forte row 4048 qty.
// Row 126220 (1308 qty) appears shared with STM32G031G6U6TR — do NOT delete here.
// Forte row 4048: stamp col E = "NO STK 7/13/2026".
// NOTE: Do NOT run addForte_STM32H573VIT6_JunWang_Jul12 — OEM is no-stk as of today.
function removeNoStk_STM32H573VIT6_Jul13() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';

  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  oemSheet.deleteRow(126359);
  Logger.log('Deleted OEM row 126359 (STM32H573VIT6 379222 qty)');

  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  forteSheet.getRange(4048, 5).setValue('NO STK 7/13/2026');
  SpreadsheetApp.flush();
  Logger.log('Stamped Forte row 4048 (STM32H573VIT6) NO STK 7/13/2026');
}

// ONE-TIME — Run removeNoStk_DS2ESDC24V_Jul13()
// David sent "DS2E-S-DC24V #4058 no stk" on 2026-07-13.
// OEM EXCESS row to DELETE: 91585 (799 qty)
// Forte row 4058: stamp col E = "NO STK 7/13/2026" (3000 qty, $3 TP, CN, Open)
function removeNoStk_DS2ESDC24V_Jul13() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';

  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  oemSheet.deleteRow(91585);
  Logger.log('Deleted OEM row 91585 (DS2E-S-DC24V 799 qty)');

  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  forteSheet.getRange(4058, 5).setValue('NO STK 7/13/2026');
  SpreadsheetApp.flush();
  Logger.log('Stamped Forte row 4058 (DS2E-S-DC24V) NO STK 7/13/2026');
}

// ONE-TIME — Run removeNoStk_7286542140_Jul13()
// David sent "7286-5421-40 #4054 No stk" on 2026-07-13.
// OEM EXCESS row to DELETE: 74060 (504 qty, Yazaki Europe Ltd Be)
// Forte row 4054: stamp col E = "NO STK 7/13/2026" (504 qty, $1.50 TP, BE, Open — created by automation Jul 13)
// NOTE: Check for any MSG_CHECKING draft to EOS Electronic (excess@components-service.com) — do NOT send it.
function removeNoStk_7286542140_Jul13() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';

  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  oemSheet.deleteRow(74060);
  Logger.log('Deleted OEM row 74060 (7286-5421-40 504 qty)');

  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  forteSheet.getRange(4054, 5).setValue('NO STK 7/13/2026');
  SpreadsheetApp.flush();
  Logger.log('Stamped Forte row 4054 (7286-5421-40) NO STK 7/13/2026');
}

// ONE-TIME — Run removeNoStk_XC7A100T1FGG484I_Jul13()
// David sent "XC7A100T-1FGG484I #3994 No Stk" on 2026-07-13.
// OEM EXCESS row to DELETE: 24662 (138 qty, AMD)
// Forte row 3994: stamp col E = "NO STK 7/13/2026" (100 qty, $7 TP, CN, Open, Jul 6 2026)
function removeNoStk_XC7A100T1FGG484I_Jul13() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';

  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];
  oemSheet.deleteRow(24662);
  Logger.log('Deleted OEM row 24662 (XC7A100T-1FGG484I 138 qty)');

  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  forteSheet.getRange(3994, 5).setValue('NO STK 7/13/2026');
  SpreadsheetApp.flush();
  Logger.log('Stamped Forte row 3994 (XC7A100T-1FGG484I) NO STK 7/13/2026');
}

// ONE-TIME — Run removeNoStk_ZOE_CCM_Jul14() AFTER sending both David reply drafts:
//   r7869308420403183158 (ZOE-M8G-0 → David thread 19f6087d394bf95d)
//   r9156615699556308437 (CCM03-3512 → David thread 19f6080e0ad6c550)
// ZOE-M8G-0: OEM row 133619, Forte row 4063
// CCM03-3512 LFT R851B: OEM row 87788, Forte row 4035
function removeNoStk_ZOE_CCM_Jul14() {
  var OEM_SHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var noStkDate = '7/14/2026';
  var oemSheet = SpreadsheetApp.openById(OEM_SHEET_ID).getSheets()[0];

  // Delete OEM rows descending (133619 > 87788)
  var oemEntries = [
    { row: 133619, expectedMpn: 'ZOE-M8G-0' },
    { row: 87788,  expectedMpn: 'CCM03-3512' }
  ];
  oemEntries.forEach(function(e) {
    var rowData = oemSheet.getRange(e.row, 1, 1, 3).getValues()[0];
    Logger.log('OEM row ' + e.row + ': ' + JSON.stringify(rowData));
    if (String(rowData[0]).toUpperCase().indexOf(e.expectedMpn.toUpperCase()) < 0) {
      Logger.log('ERROR: row ' + e.row + ' MPN mismatch — expected ' + e.expectedMpn + ', got ' + rowData[0]);
      return;
    }
    oemSheet.getRange(e.row, 5).setValue('NO STK ' + noStkDate);
    oemSheet.deleteRow(e.row);
    Logger.log('Stamped and deleted OEM row ' + e.row + ' (' + e.expectedMpn + ')');
  });
  SpreadsheetApp.flush();

  // Stamp Forte col E = "NO STK date" at specific rows with MPN verification
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var forteEntries = [
    { row: 4063, expectedMpn: 'ZOE-M8G-0' },
    { row: 4035, expectedMpn: 'CCM03-3512' }
  ];
  forteEntries.forEach(function(e) {
    var fData = forteSheet.getRange(e.row, 1, 1, 3).getValues()[0];
    Logger.log('Forte row ' + e.row + ' MPN: ' + fData[1]);
    if (String(fData[1]).toUpperCase().indexOf(e.expectedMpn.toUpperCase()) < 0) {
      Logger.log('ERROR: Forte row ' + e.row + ' MPN mismatch — expected ' + e.expectedMpn + ', got ' + fData[1]);
      return;
    }
    forteSheet.getRange(e.row, 5).setValue('NO STK ' + noStkDate);
    Logger.log('Stamped Forte row ' + e.row + ' (' + e.expectedMpn + ') → NO STK ' + noStkDate);
  });
  SpreadsheetApp.flush();
  Logger.log('removeNoStk_ZOE_CCM_Jul14 complete');
}

// ONE-TIME — Run addForte_PVX6012PBF_Jul14() after sending msg_checking draft r-738142460803916118.
// Buyer: Ken (sales@holmeselectronics.com), Holmes Electronics LLC, US.
// Confirmed TP=$5, qty=114, line=$570. OEM row 118053 (not BILL EXT).
function addForte_PVX6012PBF_Jul14() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['7/14/2026', 'PVX6012PBF', 114, 5, '', 'US',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  SpreadsheetApp.flush();
  Logger.log('Added PVX6012PBF to Forte row ' + nextRow + ' (qty=114, TP=$5, US — Holmes Electronics)');
}

// ONE-TIME — Run deleteWrongDraft_LMK05318BRGZR_Jul14() to remove the wrong request_tp_500 draft
// created for LMK05318BRGZR (Sharon Shao / Amble Electronics). Bug: BILL EXT 117 + TP=$8 in
// original RFQ should have been bill_handle — automation missed both the BILL EXT tag and the TP.
// Correct drafts already created: r4343185015141196463 (reply to Sharon CC Bill), r2470618555022853058 (FW to Bill).
function deleteWrongDraft_LMK05318BRGZR_Jul14() {
  var token = ScriptApp.getOAuthToken();
  try {
    UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/me/drafts/r2658864817345049120', {
      method: 'delete',
      headers: { 'Authorization': 'Bearer ' + token }
    });
    Logger.log('Deleted wrong request_tp_500 draft r2658864817345049120 for LMK05318BRGZR (Sharon/Amble)');
  } catch(e) {
    Logger.log('Draft r2658864817345049120 already gone or error: ' + e);
  }
}

// ONE-TIME — Run addForteRows_Jul14() to add Forte entries for inbox TP-reply threads
// processed manually on 2026-07-14 (automation backlog sweep).
// F980J107MMAAXE (kw@ascglobal.com, qty=4000, tp=1.00, PL) — SKIPPED: forte_sheet already has entry (qty=2000, tp=1.4, QUOTED)
// MKS2B044701K00KSSD (jc.cruz@reboundeu.com, qty=4000, tp=0.30, AE) — SKIPPED: forte_sheet already has entry (qty=10000, tp=0.5, Open)
// 6098-8236: forte_sheet is empty → ADD (vivien.teo@reboundeu.com, qty=900, tp=1.30, SG)
function addForteRows_Jul14() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var today = '7/14/2026';
  var rows = [
    // [mpn, qty, buyerTP, country]
    ['6098-8236', 900, 1.30, 'SG'],  // vivien.teo@reboundeu.com — Rebound EU Singapore; OEM qty=2988
  ];
  rows.forEach(function(r) {
    var nextRow = sheet.getLastRow() + 1;
    sheet.appendRow([
      today, r[0], r[1], r[2], '', r[3],
      '=C' + nextRow + '*D' + nextRow,
      '', '', '', 'Open'
    ]);
    Logger.log('Added: ' + r[0]);
  });
  Logger.log('Done — ' + rows.length + ' row(s) added to Forte sheet.');
}

// ONE-TIME — Run addForte_LA100PSP13_AEService_Jul16() to add missed Forte entry.
// Cédric DERNONCOURT (cdernoncourt@aeservice.fr, A.E. Service, FR) replied TP=$20, qty=250.
// Automation processed TP reply (oem-tp-processed) but did NOT create MSG_CHECKING draft.
// Correct MSG_CHECKING draft created manually: r-6875666363410168244
function addForte_LA100PSP13_AEService_Jul16() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var nextRow = sheet.getLastRow() + 1;
  sheet.appendRow(['7/16/2026', 'LA100-P/SP13', 250, 20, '', 'FR',
    '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
  Logger.log('Added LA100-P/SP13 to Forte row ' + nextRow + ' (250 qty, $20 TP, FR)');
}

// ONE-TIME — Run addMissingStanRows_Jul16() to backfill all Stan RFQ sheet entries
// that were missed due to the executeDecision bug (add_to_stan only called
// addToStanSheet when forte_entry was present — fixed in Jul 16 commit).
// Covers "Warehouse is checking" replies sent May 26 – Jul 16, 2026 where
// no corresponding Stan sheet row exists. addToStanSheet() has built-in dedup
// so safe to run even if some rows already exist.
function addMissingStanRows_Jul16() {
  var entries = [
    // [mpn, country, qty, tp]  — tp = '' if buyer gave no TP
    ['ADM3491ARZ',          'CN', 1071, ''],    // Bonnie Chan / Shenzhen Hengchenxin — Jul 16
    ['L6202',               'NL', 50,   ''],    // Richard Cross / ChipSource Europe — Jul 16
    ['TR3B107M010C1400',    'CN', 2000, ''],    // YiMin Ke / Innovation Ray — Jul 15
    ['IDT72V235L15TF',      'HK', 75,   ''],    // Wan Yiu Ling / Linkduty Co. — Jul 14
    ['ADG5404BRUZ',         'CN', 6400, 2],     // Kevin Hu / Zhichenxing (earliest, has TP $2) — Jul 11-13
    ['XC6SLX150T-3FGG676I', 'GB', 33,   ''],    // Charmaine / Vital Electronics UK — Jul 7
    ['BCM5421A1IMLG',       'CN', 1300, ''],    // Binacupeng@foxmail.com — Jul 3
    ['EL4390CM',            'SG', 100,  ''],    // Sumit Gupta / India Electronics — Jun 30 (dedup handles if already in sheet)
    ['E28F800B5T90',        'US', 60,   2.50],  // Sales Group / Alpha-Micro Electronics — Jun 19
    ['EPM7128STC100-10',    'US', 200,  ''],    // Tiffany Hull / AERI — Jun 18
    ['K4S561632C-TC75',     'FR', 50,   ''],    // Lou Barbe / Club Electronics — Jun 17
    ['ADF4116BRUZ',         'CN', 133,  1.80],  // Bonnie Chan / Shenzhen Hengchenxin — Jun 4 (TP $1.80 given Jun 8)
    ['Z8S18020VEC',         'US', 80,   ''],    // George Beezer / Arcadia Components — Jun 1
    ['PPC440GX-3CF800C',    'CN', 40,   ''],    // David Zhang / Shenzhen Mino Industry — May 27
    ['XC2S150-5FG256I',     'US', 100,  ''],    // Olivia Zenteno / Velocity Electronics — May 26
    ['LTL533-11',           'CN', 5000, ''],    // Joey Zhang / HK DCY Technology — May 26
    ['IS42S32400E-7TLI',    'CN', 400,  ''],    // Bonnie Chan / Shenzhen Hengchenxin — May 26
    ['EPF10K30ABC356-1',    'US', 10,   ''],    // Han Taehoon / 4 Star Electronics — May 26
    ['AM28F010120JC',       'US', 250,  ''],    // Han Taehoon / 4 Star Electronics — May 26
  ];

  var added = 0, skipped = 0;
  entries.forEach(function(e) {
    var mpn = e[0], country = e[1], qty = e[2], tp = e[3];
    var existing = searchStanSheet(mpn);
    if (existing.length > 0) {
      Logger.log('SKIP (already in Stan): ' + mpn);
      skipped++;
      return;
    }
    var sheet = SpreadsheetApp.openById(STAN_SHEET_ID).getSheets()[0];
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
    sheet.appendRow(['', '', '', today, mpn, country, qty, tp || '']);
    Logger.log('ADDED to Stan: ' + mpn + ' | ' + country + ' | QTY:' + qty + ' | TP:' + (tp || 'none'));
    added++;
  });
  SpreadsheetApp.flush();
  Logger.log('addMissingStanRows_Jul16 complete — added: ' + added + ', skipped: ' + skipped);
}

// ONE-TIME — Run addMissing3Stan_Jul16b() to force-add the 3 entries that were
// still missing from Stan sheet after addMissingStanRows_Jul16() ran.
// Web app confirmed MISSING for these 3 on Jul 16 check.
function addMissing3Stan_Jul16b() {
  var sheet = SpreadsheetApp.openById(STAN_SHEET_ID).getSheets()[0];
  var entries = [
    ['ADM3491ARZ',       'CN', 1071, ''],  // Bonnie Chan / Shenzhen Hengchenxin — Jul 16
    ['L6202',            'NL', 50,   ''],  // Richard Cross / ChipSource Europe — Jul 16
    ['TR3B107M010C1400', 'CN', 2000, ''],  // YiMin Ke / Innovation Ray — Jul 15
  ];
  var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
  entries.forEach(function(e) {
    sheet.appendRow(['', '', '', date, e[0], e[1], e[2], e[3]]);
    Logger.log('ADDED: ' + e[0]);
  });
  SpreadsheetApp.flush();
  Logger.log('addMissing3Stan_Jul16b done — 3 rows added');
}

// ONE-TIME — Run deleteJul17NoStockOEM_oneTime() to clean up OEM EXCESS for:
//   1. DA7217-00U32 (BILL EXT): Bill confirmed sold, no longer available (Jul 17)
//   2. ADM232AARNZ-REEL7 (row 4094): David said no stock (Jul 17) — belt+suspenders
// Also stamps Forte row 4094 black/white in case automation hasn't run yet.
function deleteJul17NoStockOEM_oneTime() {
  // Delete DA7217-00U32 from OEM EXCESS (BILL EXT — not in Forte)
  var r1 = deletePart('DA7217-00U32', 'Jul17 Bill confirmed sold');
  Logger.log('DA7217-00U32 deletePart → ' + r1);

  // Delete ADM232AARNZ-REEL7 from OEM EXCESS + stamp Forte row 4094
  var r2 = deletePart('ADM232AARNZ-REEL7', 'Jul17 David no stock #4094');
  Logger.log('ADM232AARNZ-REEL7 deletePart → ' + r2);
  updateForteSheet('ADM232AARNZ-REEL7');
  Logger.log('ADM232AARNZ-REEL7 Forte stamped');

  Logger.log('deleteJul17NoStockOEM_oneTime: DONE');
}

// ONE-TIME — Run addForte230N3V14_Jul17() to add 230N3V14 to Forte.
// Wanda/Nexelec (US) gave TP=$2.00 for qty 250 (line=$500, meets $500 min).
// MSG_CHECKING draft sent Jul 17 2026.
function addForte230N3V14_Jul17() {
  var existing = checkForteForMPN('230N3V14', 60);
  if (existing) { Logger.log('230N3V14 already in Forte within 60 days — skipping'); return; }
  addToForteSheet('230N3V14', 250, 2.00, 'US', '');
  Logger.log('addForte230N3V14_Jul17: DONE');
}

// ONE-TIME — Run deleteRAA2100404_Jul17() for David no-stk #4095 (Jul 17 2026)
// Stamps Forte row 4095 NO STK (black/white/bold) + deletes from OEM EXCESS
function deleteRAA2100404_Jul17() {
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var cell = forteSheet.getRange(4095, FORTE_STATUS_COL + 1);
  var cur = String(cell.getValue()).trim().toUpperCase();
  if (cur.indexOf('NO STK') === -1 && cur !== 'CLOSED') {
    cell.clearDataValidations();
    cell.setValue('NO STK - 7/17/2026');
    cell.setBackground('#000000'); cell.setFontColor('#FFFFFF'); cell.setFontWeight('bold');
    Logger.log('Stamped Forte row 4095');
  } else {
    Logger.log('Row 4095 already stamped: ' + cur);
  }
  var r = deletePart('RAA2100404GLGMD0', 'David no-stk #4095 Jul17');
  Logger.log('deletePart RAA2100404GLGMD0 → ' + r);
}

// ── msg_checking Forte entries Jul 20 2026 ───────────────────
// 2 buyer TP replies processed manually. Drafts created via Gmail MCP.
// Run once to add Forte entries.
function addForteEntries_Jul20_oneTime() {
  var today = '7/20/2026';
  // SIM-ST-MFF2 — jason@southelectronics.com, 400pcs @ $2.50, US
  // Prior Forte entry from Jan 2025 (row 1935) — outside 60-day window, OK to add
  addToForteSheet('SIM-ST-MFF2', 400, 2.50, 'US', '');
  Logger.log('Added SIM-ST-MFF2');
  // 151-02245 — sofiag@chip1.com, 20000pcs @ $1.50, US
  // No prior Forte entry
  addToForteSheet('151-02245', 20000, 1.50, 'US', '');
  Logger.log('Added 151-02245');
}

// ── David no-stk Jul 20 2026 ──────────────────────────────────
// 5 no-stock emails from David. Drafts already created via Gmail MCP.
// Run this once to stamp Forte rows + delete from OEM EXCESS.
function davidNoStk_Jul20_oneTime() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var FORTE_STATUS_COL = 11;
  var stamp = 'NO STK - 7/20/2026';
  var noStks = [
    { row: 4102, mpn: 'CAT4237TD-GT3' },
    { row: 4106, mpn: 'BAT165-E6327' },
    { row: 4104, mpn: 'ICM-42670-P' },
    { row: 3695, mpn: 'BUK9875-100A/CUX' },
    { row: 4100, mpn: 'LDECD4100KA0N00' },
  ];
  noStks.forEach(function(item) {
    var cell = forteSheet.getRange(item.row, FORTE_STATUS_COL);
    var cur = cell.getValue();
    if (!cur || cur.toString().toUpperCase().indexOf('NO STK') === -1) {
      cell.clearDataValidations();
      cell.setValue(stamp);
      cell.setBackground('#000000');
      cell.setFontColor('#FFFFFF');
      cell.setFontWeight('bold');
      Logger.log('Stamped Forte row ' + item.row + ' (' + item.mpn + ')');
    } else {
      Logger.log('Row ' + item.row + ' already stamped: ' + cur);
    }
    var result = deletePart(item.mpn, 'David no-stk Jul20');
    Logger.log('deletePart ' + item.mpn + ' → ' + result);
  });
}

function davidNoStk_FT232RL_Jul20_oneTime() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var FORTE_STATUS_COL = 11;
  var stamp = 'NO STK - 7/20/2026';
  var cell = forteSheet.getRange(4114, FORTE_STATUS_COL);
  cell.clearDataValidations();
  cell.setValue(stamp);
  cell.setBackground('#000000');
  cell.setFontColor('#FFFFFF');
  cell.setFontWeight('bold');
  Logger.log('Stamped Forte row 4114 (FT232RL-REEL)');
  var result = deletePart('FT232RL-REEL', 'David no-stk Jul20');
  Logger.log('deletePart FT232RL-REEL → ' + result);
}

function davidNoStk_XP1001000_Jul20_oneTime() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var FORTE_STATUS_COL = 11;
  var stamp = 'NO STK - 7/20/2026';
  var cell = forteSheet.getRange(4008, FORTE_STATUS_COL);
  cell.clearDataValidations();
  cell.setValue(stamp);
  cell.setBackground('#000000');
  cell.setFontColor('#FFFFFF');
  cell.setFontWeight('bold');
  Logger.log('Stamped Forte row 4008 (XP1001000-05R)');
  var result = deletePart('XP1001000-05R', 'David no-stk Jul20');
  Logger.log('deletePart XP1001000-05R → ' + result);
}

function davidNoStk_LF353DR_Jul21_oneTime() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var FORTE_STATUS_COL = 11;
  var stamp = 'NO STK - 7/21/2026';
  var cell = forteSheet.getRange(3813, FORTE_STATUS_COL);
  cell.clearDataValidations();
  cell.setValue(stamp);
  cell.setBackground('#000000');
  cell.setFontColor('#FFFFFF');
  cell.setFontWeight('bold');
  Logger.log('Stamped Forte row 3813 (LF353DR)');
  var result = deletePart('LF353DR', 'David no-stk Jul21');
  Logger.log('deletePart LF353DR → ' + result);
}

function davidNoStk_NRF52840_Jul21_oneTime() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var FORTE_STATUS_COL = 11;
  var stamp = 'NO STK - 7/21/2026';
  var cell = forteSheet.getRange(4092, FORTE_STATUS_COL);
  cell.clearDataValidations();
  cell.setValue(stamp);
  cell.setBackground('#000000');
  cell.setFontColor('#FFFFFF');
  cell.setFontWeight('bold');
  Logger.log('Stamped Forte row 4092 (NRF52840-CKAA-R)');
  var result = deletePart('NRF52840-CKAA-R', 'David no-stk Jul21');
  Logger.log('deletePart NRF52840-CKAA-R → ' + result);
}

function davidNoStk_MT60B1G16HD_Jul21_oneTime() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var FORTE_STATUS_COL = 11;
  var stamp = 'NO STK - 7/21/2026';
  var cell = forteSheet.getRange(4103, FORTE_STATUS_COL);
  cell.clearDataValidations();
  cell.setValue(stamp);
  cell.setBackground('#000000');
  cell.setFontColor('#FFFFFF');
  cell.setFontWeight('bold');
  Logger.log('Stamped Forte row 4103 (MT60B1G16HD-64B:H)');
  var result = deletePart('MT60B1G16HD-64B:H', 'David no-stk Jul21');
  Logger.log('deletePart MT60B1G16HD-64B:H → ' + result);
}

function davidNoStk_PEX8724_Jul21_oneTime() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var FORTE_STATUS_COL = 11;
  var stamp = 'NO STK - 7/21/2026';
  var cell = forteSheet.getRange(4101, FORTE_STATUS_COL);
  cell.clearDataValidations();
  cell.setValue(stamp);
  cell.setBackground('#000000');
  cell.setFontColor('#FFFFFF');
  cell.setFontWeight('bold');
  Logger.log('Stamped Forte row 4101 (PEX8724-CA80BC)');
  var result = deletePart('PEX8724-CA80BC', 'David no-stk Jul21');
  Logger.log('deletePart PEX8724-CA80BC → ' + result);
}

// ONE-TIME — Run fixDavidNoStk_Jul21_Missed() to stamp Forte col K and delete from OEM EXCESS
// for 2 no-stk emails missed by automation on Jul 21 2026:
//   row 4027 → KLMAG1JETD-B041   row 4074 → PR-K-24
function fixDavidNoStk_Jul21_Missed() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var stamp = 'NO STK - 7/21/2026';
  var rows = [
    { row: 4027, mpn: 'KLMAG1JETD-B041' },
    { row: 4074, mpn: 'PR-K-24' },
  ];
  rows.forEach(function(r) {
    var cell = sheet.getRange(r.row, 11); // col K = Status
    var current = String(cell.getValue()).trim().toUpperCase();
    if (current.indexOf('NO STK') === -1 && current !== 'CLOSED') {
      cell.clearDataValidations();
      cell.setValue(stamp);
      cell.setBackground('#000000');
      cell.setFontColor('#FFFFFF');
      cell.setFontWeight('bold');
      Logger.log('Stamped Forte row ' + r.row + ' (' + r.mpn + ')');
    } else {
      Logger.log('Skipped Forte row ' + r.row + ' (' + r.mpn + ') — already: ' + current);
    }
    var result = deletePart(r.mpn, 'David no-stk Jul21');
    Logger.log('deletePart ' + r.mpn + ' → ' + result);
  });
}
