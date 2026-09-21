// ONE-TIME: Restore A-2014-2S-4-N-R to OEM EXCESS and reset Forte status
// Caused by: worker falsely fired remove_oem because historical "Future No stk" text
// in Aug 25 David message was found in full thread_content on Sep 20 processing.
// Run ONCE then delete this file.

var SPREADSHEET_ID = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
var TARGET_MPN     = 'A-2014-2S-4-N-R';

function restoreA2014() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var mainSheet    = ss.getSheetByName('sheet1');
  var deletedSheet = ss.getSheetByName('Deleted Rows');

  if (!deletedSheet) { Logger.log('ERROR: Deleted Rows sheet not found'); return; }

  // Find most recent deletion of this MPN in Deleted Rows
  // Columns: Date Deleted | FullPartNumber | Man | DC | QTY | Notes | Source
  var delData = deletedSheet.getDataRange().getValues();
  var foundRow = null;
  for (var i = delData.length - 1; i >= 1; i--) {
    if (String(delData[i][1]).trim().toLowerCase() === TARGET_MPN.toLowerCase()) {
      foundRow = delData[i];
      Logger.log('Found deleted row: ' + JSON.stringify(foundRow));
      break;
    }
  }

  if (!foundRow) {
    Logger.log('ERROR: No deleted row found for ' + TARGET_MPN);
    return;
  }

  // Restore original row: Col A=MPN, B=Man, C=DC, D=QTY, E=Notes
  var mpn   = foundRow[1];
  var man   = foundRow[2];
  var dc    = foundRow[3];
  var qty   = foundRow[4];
  var notes = foundRow[5];

  mainSheet.appendRow([mpn, man, dc, qty, notes]);
  Logger.log('Restored to OEM EXCESS: ' + mpn + ' | ' + man + ' | DC:' + dc + ' | QTY:' + qty + ' | ' + notes);

  // Reset Forte status from "NO STK - 9/20/2026" back to "Open"
  var forte = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var fData = forte.getDataRange().getValues();
  var FORTE_STATUS_COL = 10; // col K (0-indexed: 10)
  for (var j = 1; j < fData.length; j++) {
    if (String(fData[j][1]).trim().toLowerCase() === TARGET_MPN.toLowerCase()) {
      var statusCell = forte.getRange(j + 1, FORTE_STATUS_COL + 1);
      Logger.log('Resetting Forte row ' + (j+1) + ' status from "' + fData[j][FORTE_STATUS_COL] + '" to "Open"');
      statusCell.clearDataValidations();
      statusCell.setValue('Open');
    }
  }

  Logger.log('Done. Verify the restored OEM EXCESS row and Forte status.');
}
