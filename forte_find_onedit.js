/**
 * Forte sheet "Find" tab -> write-back to Sheet1.
 * Paste into the FORTE sheet: Extensions > Apps Script (bound project). Simple trigger, no install/auth needed.
 *
 * Find layout: B1 = search box, row 3 = headers, A4 = row-number formula (Sheet1 row ids),
 * B4 = data formula (Sheet1 A:AE, newest first; V:AE = Market pricing 1-10). Find column c  <->  Sheet1 column c-1.
 * Typing into a result cell breaks the B4 spill, so we copy the typed value(s) to Sheet1,
 * clear what was typed, and put both formulas back — the results refresh with the new value.
 */
var FIND_TAB = 'Find';
var MAIN_TAB = 'Sheet1';
var FIND_FORMULA_A = "=IF(TRIM($B$1)=\"\",\"\",IFERROR(LET(r,SORT(FILTER(ROW(Sheet1!A3:A),ISNUMBER(SEARCH(TRIM($B$1),Sheet1!B3:B))),1,FALSE),ARRAYFORMULA(HYPERLINK(\"#gid=0&range=A\"&r&\":W\"&r,r))),\"No match\"))";
var FIND_FORMULA_B = "=IF(TRIM($B$1)=\"\",\"\",IFERROR(SORT(FILTER(Sheet1!A3:AE,ISNUMBER(SEARCH(TRIM($B$1),Sheet1!B3:B))),FILTER(ROW(Sheet1!A3:A),ISNUMBER(SEARCH(TRIM($B$1),Sheet1!B3:B))),FALSE),\"\"))";

function onEdit(e) {
  var rng = e.range, sh = rng.getSheet();
  if (sh.getName() !== FIND_TAB || rng.getLastRow() < 4) return;   // search box / headers: normal edit

  var top = Math.max(rng.getRow(), 4), left = rng.getColumn();
  var nRows = rng.getLastRow() - top + 1, nCols = rng.getLastColumn() - left + 1;
  var typed = sh.getRange(top, left, nRows, nCols).getValues();
  var ids = sh.getRange(top, 1, nRows, 1).getValues();              // A spill survives B edits
  var main = e.source.getSheetByName(MAIN_TAB);

  for (var i = 0; i < nRows; i++) {
    var row = Number(ids[i][0]);
    if (!row || row < 3) continue;                                    // blank / "No match" line
    for (var j = 0; j < nCols; j++) {
      var c = left + j;
      if (c < 2 || c > 32) continue;                                  // col A = row #, beyond AE = not shown
      if (typed[i][j] === '') continue;   // Delete on B4 collapses the spill so every cell reads blank — never wipe Sheet1 from here
      main.getRange(row, c - 1).setValue(typed[i][j]);
    }
  }

  sh.getRange(top, left, nRows, nCols).clearContent();
  sh.getRange('A4').setFormula(FIND_FORMULA_A);
  sh.getRange('B4').setFormula(FIND_FORMULA_B);
}
