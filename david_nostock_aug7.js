// ONE-TIME — Run davidNoStockAug7() to process all David no-stk emails from Aug 6-7 2026:
//   CM4104032  #4119 #4120  Cant Share  (Aug 6)
//   TPS82130SILR #4252      no stk      (Aug 7)
//   IDW40G120C5BFKSA1 #4254 No stk     (Aug 7)
//
// Actions:
//   1. Stamp Forte col K (Status) with NO STK - 8/7/2026 for rows 4119, 4120, 4252, 4254
//   2. Delete CM4104032, TPS82130SILR, IDW40G120C5BFKSA1 from OEM EXCESS
function davidNoStockAug7() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var OEM_EXCESS_ID  = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
  var stamp = 'NO STK - 8/7/2026';
  var noStkStamp = 'NO STK 8/7/2026';

  // ── Step 1: Stamp Forte rows col K (index 11) ──
  var forteSheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  [4119, 4120, 4252, 4254].forEach(function(rowNum) {
    try {
      var cell = forteSheet.getRange(rowNum, 11);
      var current = String(cell.getValue()).trim();
      if (current.toUpperCase().indexOf('NO STK') === -1 && current.toUpperCase() !== 'CLOSED') {
        cell.clearDataValidations();
        cell.setValue(stamp);
        cell.setBackground('#000000');
        cell.setFontColor('#FFFFFF');
        cell.setFontWeight('bold');
        Logger.log('Stamped Forte row ' + rowNum);
      } else {
        Logger.log('Forte row ' + rowNum + ' already stamped: ' + current);
      }
    } catch(e) {
      Logger.log('Error stamping row ' + rowNum + ': ' + e);
    }
  });

  // ── Step 2: Delete MPNs from OEM EXCESS ──
  var targets = ['CM4104032', 'TPS82130SILR', 'IDW40G120C5BFKSA1'];
  var ss = SpreadsheetApp.openById(OEM_EXCESS_ID);
  var mainSheet = ss.getSheetByName('sheet1');

  var deletedSheet = ss.getSheetByName('deleted');
  if (!deletedSheet) deletedSheet = ss.insertSheet('deleted');

  targets.forEach(function(target) {
    var data = mainSheet.getDataRange().getValues();
    var found = false;
    // Scan bottom-up so row deletion doesn't shift indexes
    for (var i = data.length - 1; i >= 1; i--) {
      var mpn = String(data[i][0]).trim().replace(/-/g, '').toUpperCase();
      var normalTarget = target.replace(/-/g, '').toUpperCase();
      if (mpn === normalTarget || String(data[i][0]).trim() === target) {
        found = true;
        var rowNum = i + 1;
        mainSheet.getRange(rowNum, 5).setValue(noStkStamp);
        var logRow = data[i].slice();
        logRow.push('David no-stk 8/7/2026 - ' + target);
        deletedSheet.appendRow(logRow);
        mainSheet.deleteRow(rowNum);
        Logger.log('Deleted OEM EXCESS row ' + rowNum + ' (' + target + ')');
        // Re-read data after deletion
        data = mainSheet.getDataRange().getValues();
      }
    }
    if (!found) Logger.log(target + ' not found in OEM EXCESS (may already be removed)');
  });

  Logger.log('davidNoStockAug7: DONE');
}
