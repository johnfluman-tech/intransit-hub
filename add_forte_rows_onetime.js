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
