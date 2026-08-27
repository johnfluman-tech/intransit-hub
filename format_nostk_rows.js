// One-time: apply black background + white text to all "NO STK" rows in Forte sheet
// Run formatNoStkRows() once in Apps Script editor, then delete
function formatNoStkRows() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  // Read col K (index 10, col 11) for all data rows
  var colK = sheet.getRange(2, 11, lastRow - 1, 1).getValues();

  var count = 0;
  for (var i = 0; i < colK.length; i++) {
    var status = (colK[i][0] || '').toString().toUpperCase();
    if (status.indexOf('NO STK') !== -1) {
      var rowNum = i + 2;
      var rowRange = sheet.getRange(rowNum, 1, 1, lastCol);
      rowRange.setBackground('#000000');
      rowRange.setFontColor('#FFFFFF');
      count++;
      Logger.log('Formatted row ' + rowNum + ': ' + colK[i][0]);
    }
  }
  Logger.log('Done. Formatted ' + count + ' rows.');
  SpreadsheetApp.flush();
}
