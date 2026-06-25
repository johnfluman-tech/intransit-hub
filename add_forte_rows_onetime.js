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
