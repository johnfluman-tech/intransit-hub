// One-time catch-up: add all W3_CHECKING replies since Jul 9 that were never
// added to Stan's RFQ sheet.  Run addMissingStanRFQs_Jul14() once, verify, done.

var STAN_SHEET_ID = '1pGRDpkqftQNoEYna53MxRJfUY8jEf5_w32FNa56OUIM';

function addMissingStanRFQs_Jul14() {
  var sheet = SpreadsheetApp.openById(STAN_SHEET_ID).getSheets()[0];

  // Format: ['', '', '', date(M/d/yyyy), mpn, country, qty, tp]
  // Cols A-C blank, D=date, E=MPN, F=country, G=qty, H=TP ('' if none given)
  var rows = [
    ['', '', '', '7/14/2026', 'IDT72V235L15TF',      'HK', 75,   ''],   // Linkduty
    ['', '', '', '7/13/2026', 'ADG5404BRUZ',          'CN', 1285, ''],   // Hengchenxin/Janet
    ['', '', '', '7/13/2026', 'ADG5404BRUZ',          'CN', 1000, ''],   // Hongqi/Kevin Chu
    ['', '', '', '7/13/2026', 'LTC3311SEV#TRPBF',     'CN', 2500, ''],   // Hengchenxin/Bonnie
    ['', '', '', '7/13/2026', 'MPM3650GQW-P',         'NL', 104,  ''],   // Innova Technologies BV
    ['', '', '', '7/12/2026', 'ADG5404BRUZ',          'CN', 6400, 2],    // Zhichenxing/Kevin Hu (TP=$2 given)
    ['', '', '', '7/12/2026', 'PTMA210152M',          'CN', 238,  ''],   // Hengchenxin/Bonnie
    ['', '', '', '7/10/2026', 'LM150K/883',           'US', 1,    ''],   // Intercomp/Mariah
    ['', '', '', '7/10/2026', 'ADSP21065L-CS-240',    'CN', 40,   ''],   // Zhidian Hechuang/Bella
    ['', '', '', '7/7/2026',  'XCV600E-6FG680C',      'EU', 80,   ''],   // Converge EMEA
    ['', '', '', '7/9/2026',  'XC5VLX85T-1FFG1136',  'HK', 1000, ''],   // Jenny Penny/Sandy
    ['', '', '', '7/9/2026',  'DAC5688I',             'CN', 106,  ''],   // Hengchenxin/Bonnie
    ['', '', '', '7/9/2026',  'SA605D',               'HK', 141,  ''],   // Zhongkai/Mila
  ];

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, 8).setValues(rows);
  Logger.log('Done — added ' + rows.length + ' rows starting at row ' + (lastRow + 1));
}
