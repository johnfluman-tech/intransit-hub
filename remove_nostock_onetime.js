// ============================================================
// STANDALONE one-time script — no dependencies on main email script.
// Paste this as the ENTIRE content of a new Apps Script project.
// Run removeNoStockParts() to delete the 3 parts David confirmed no-stock.
// ============================================================

var OEM_SHEET_ID   = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4';
var OEM_SHEET_NAME = 'sheet1';
var DELETED_TAB    = 'Deleted Rows';

function removeNoStockParts() {
  var parts = [
    { mpn: 'MCIMX535DVP1C2', note: '#3900 David — No stock' },
    { mpn: 'BD429B',          note: '#3904 David — Cant find' },
    { mpn: '10M08SAM153I7G',  note: '#3896 David — NO STOCK'  },
  ];

  var ss        = SpreadsheetApp.openById(OEM_SHEET_ID);
  var main      = ss.getSheetByName(OEM_SHEET_NAME);
  var delSheet  = getOrMakeDeletedSheet(ss);
  var data      = main.getDataRange().getValues();

  parts.forEach(function(item) {
    var found = false;
    // Search from bottom so row deletion doesn't shift indexes
    for (var i = data.length - 1; i >= 1; i--) {
      var cell = String(data[i][0]).trim();
      if (cell.toLowerCase() === item.mpn.toLowerCase()) {
        // Archive the row
        var archived = data[i].concat([new Date(), item.note]);
        delSheet.appendRow(archived);
        main.deleteRow(i + 1);
        Logger.log('DELETED row ' + (i + 1) + ': ' + item.mpn);
        found = true;
        break;
      }
    }
    if (!found) Logger.log('NOT FOUND (already removed?): ' + item.mpn);
  });

  Logger.log('Done. Check View > Logs for results.');
}

function getOrMakeDeletedSheet(ss) {
  var s = ss.getSheetByName(DELETED_TAB);
  if (!s) {
    s = ss.insertSheet(DELETED_TAB);
    s.appendRow(['MPN', 'QTY', 'Notes', 'Deleted At', 'Reason']);
  }
  return s;
}
