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
