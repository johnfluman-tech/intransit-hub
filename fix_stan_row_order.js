// One-time fix: move LCMXO2 and MC68HC000FN16 from rows 1-2 to the bottom of Stan's sheet
// Paste into Apps Script editor and run fixStanSheetOrder() once
function fixStanSheetOrder() {
  var STAN_SHEET_ID = '1pGRDpkqftQNoEYna53MxRJfUY8jEf5_w32FNa56OUIM';
  var ss = SpreadsheetApp.openById(STAN_SHEET_ID);
  var sheet = ss.getSheets()[0];

  var row1 = sheet.getRange(1, 1, 1, 8).getValues()[0];
  var row2 = sheet.getRange(2, 1, 1, 8).getValues()[0];
  Logger.log('Row 1: ' + row1.join(' | '));
  Logger.log('Row 2: ' + row2.join(' | '));

  if (!row1[4] || row1[4].trim() !== 'LCMXO2-7000HC-4BG256I') {
    Logger.log('ERROR: Row 1 col E is not LCMXO2-7000HC-4BG256I — aborting');
    return;
  }
  if (!row2[4] || row2[4].trim() !== 'MC68HC000FN16') {
    Logger.log('ERROR: Row 2 col E is not MC68HC000FN16 — aborting');
    return;
  }

  // Delete both top rows (row 2 shifts to 1 after first delete)
  sheet.deleteRow(1);
  sheet.deleteRow(1);

  // Append at the true bottom
  sheet.appendRow(row1);
  sheet.appendRow(row2);

  Logger.log('Done — LCMXO2 and MC68HC000FN16 moved to bottom.');
}
