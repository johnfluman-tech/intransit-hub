/**
 * Forte sheet "Find" tab -> write-back to Sheet1.
 * Paste into the FORTE sheet: Extensions > Apps Script (bound project). Simple trigger, no install/auth needed.
 *
 * Find layout: B1 = search box, row 3 = headers, A4 = row-number formula (Sheet1 row ids),
 * B4 = data formula (Sheet1 A:AE, newest first; V:AE = Market pricing 1-10). Find column c  <->  Sheet1 column c-1.
 * Typing into a result cell breaks the B4 spill, so we copy the typed value(s) to Sheet1,
 * clear what was typed, and put both formulas back — the results refresh with the new value.
 * Delete works too: cells David selects and clears are cleared on Sheet1 — one result row at a time
 * (a multi-row Delete is ignored so a slip can't wipe a block of Sheet1).
 */
var FIND_TAB = 'Find';
var MAIN_TAB = 'Sheet1';
var FIND_FORMULA_A = "=IF(TRIM($B$1)=\"\",\"\",IFERROR(LET(r,SORT(FILTER(ROW(Sheet1!A3:A),ISNUMBER(SEARCH(TRIM($B$1),Sheet1!B3:B))),1,FALSE),ARRAYFORMULA(HYPERLINK(\"#gid=0&range=A\"&r&\":AE\"&r,r))),\"No match\"))";
var FIND_FORMULA_B = "=IF(TRIM($B$1)=\"\",\"\",IFERROR(SORT(FILTER(Sheet1!A3:AE,ISNUMBER(SEARCH(TRIM($B$1),Sheet1!B3:B))),FILTER(ROW(Sheet1!A3:A),ISNUMBER(SEARCH(TRIM($B$1),Sheet1!B3:B))),FALSE),\"\"))";

function onEdit(e) {
  var rng = e.range, sh = rng.getSheet();
  if (sh.getName() !== FIND_TAB || rng.getLastRow() < 4) return;   // search box / headers: normal edit

  var top = Math.max(rng.getRow(), 4), left = rng.getColumn();
  var nRows = rng.getLastRow() - top + 1, nCols = rng.getLastColumn() - left + 1;
  var single = rng.getNumRows() === 1 && rng.getNumColumns() === 1;
  var typed = single ? [[e.value === undefined ? '' : rng.getValue()]]   // single cell: trust the edit event
                     : sh.getRange(top, left, nRows, nCols).getValues();

  // Put the row-number formula back first (in case col A was typed over / deleted), then read the ids
  sh.getRange('A4').setFormula(FIND_FORMULA_A);
  SpreadsheetApp.flush();
  var ids = sh.getRange(top, 1, nRows, 1).getValues();
  var main = e.source.getSheetByName(MAIN_TAB);
  var allowClear = nRows === 1, cleared = [];

  for (var i = 0; i < nRows; i++) {
    var row = Number(ids[i][0]);
    if (!row || row < 3) continue;                                    // blank / "No match" line
    for (var j = 0; j < nCols; j++) {
      var c = left + j;
      if (c < 2 || c > 32) continue;                                  // col A = row #, beyond AE = not shown
      var cell = main.getRange(row, c - 1);
      if (typed[i][j] === '') {
        if (!allowClear) continue;                                    // multi-row Delete: never wipe a block of Sheet1
        var old = cell.getDisplayValue();
        if (old === '') continue;
        cleared.push(cell.getA1Notation() + ' "' + old + '"');
        cell.clearContent();
      } else {
        cell.setValue(typed[i][j]);
      }
    }
  }

  sh.getRange(top, left, nRows, nCols).clearContent();
  sh.getRange('A4').setFormula(FIND_FORMULA_A);
  sh.getRange('B4').setFormula(FIND_FORMULA_B);

  if (cleared.length) e.source.toast('Cleared on Sheet1: ' + cleared.join(', '), 'Find', 8);
  else if (!allowClear && typed.some(function (r) { return r.some(function (v) { return v === ''; }); }))
    e.source.toast('Delete one row at a time on Find (or delete on Sheet1).', 'Find', 6);
}
