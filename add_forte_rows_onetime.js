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
function deleteOldWrongDrafts_Jul3() {
  var token = ScriptApp.getOAuthToken();
  var toDelete = [
    'r3853834301542523820',
    'r-4586403387947820516',
    'r1716529454449499559',
    'r-1733567911978935294',
    'r-3837120187750311243',
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
