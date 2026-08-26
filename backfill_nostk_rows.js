// One-time: stamp Forte rows 4380 and 4391 col K with "NO STK - [date]" + black bg / white text
// Run backfillNoStkRows() once in Apps Script editor, then delete
function backfillNoStkRows() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var rows = [4380, 4391];
  var today = '8/20/2026';

  rows.forEach(function(rowNum) {
    var mpn = sheet.getRange(rowNum, 2).getValue();
    var existingK = (sheet.getRange(rowNum, 11).getValue() || '').toString().trim();
    Logger.log('Row ' + rowNum + ' MPN=' + mpn + ' existingK=' + existingK);

    // Only stamp if not already stamped
    if (existingK.toUpperCase().indexOf('NO STK') === -1) {
      sheet.getRange(rowNum, 11).setValue('NO STK - ' + today);
      Logger.log('  Stamped K' + rowNum);
    } else {
      Logger.log('  Already stamped: ' + existingK);
    }

    // Apply black bg + white text to col K only
    var kCell = sheet.getRange(rowNum, 11);
    kCell.setBackground('#000000');
    kCell.setFontColor('#FFFFFF');
    Logger.log('  Formatted K' + rowNum);
  });

  SpreadsheetApp.flush();
  Logger.log('Done.');
}
