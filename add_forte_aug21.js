// One-time: add 3 Forte entries for MSG_CHECKING threads (Aug 20, 2026)
// Run addForteAug21() once in Apps Script editor, then delete
function addForteAug21() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var today = '8/20/2026';
  var rows = [
    // [mpn, qty, buyerTP, country]
    ['BSC110N06NS3G', 3000, 0.25, 'DK'],  // Petra / Emporium Partners
    ['CC2640F128RSMR', 15000, 0.60, 'CN'], // Dorian / Shenzhen Solandi
    ['5009980900', 129, 6.00, 'CN'],        // Gehunte Power Supply
  ];
  rows.forEach(function(r) {
    var nextRow = sheet.getLastRow() + 1;
    sheet.appendRow([today, r[0], r[1], r[2], '', r[3],
      '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
    Logger.log('Added: ' + r[0]);
  });
}
