"""
process_forte_david_v3.py
Standard Forte raw-file conversion for David's new-format file.

Input : ForteCurrentAvailableExcess2 (Repaired) (5).xlsx  (no header, cols A-F)
         Col A = PartNumber, B = Manufacturer, C = Quantity, D = Location/customer, E = Notes
Output: forte_processed_output.csv  (header + data; 5 cols)
         FullPartNumber, Man, DC, QTY, Notes

Run from any directory:
    python "G:/My Drive/Download Internet/process_forte_david_v3.py"

After completion, run importForteBulk() in Apps Script to push to OEM EXCESS sheet.
"""

import re, csv, sys, math, time
import openpyxl
import gspread
from google.oauth2.service_account import Credentials

INPUT_FILE  = r'G:\My Drive\Download Internet\ForteCurrentAvailableExcess2 (Repaired) (5).xlsx'
OUTPUT_FILE = r'G:\My Drive\Download Internet\forte_processed_output.csv'

# ── Cheap-part filter ────────────────────────────────────────────────────────
# Match by known passive-component part-number PREFIXES only.
# Do NOT use substring matching (e.g. "CAP", "RES") — would hit IC part numbers.

CHEAP_PREFIXES = tuple(p.upper() for p in [
    # Resistors — Panasonic
    'ERJ', 'ERJP', 'ERJU', 'ERJS',
    # Resistors — Vishay / Dale
    'CRCW', 'WSL', 'WSLF', 'WSLT', 'WSHM', 'WSHA', 'RNMF', 'MFR', 'CMF', 'CFR',
    # Resistors — Rohm
    'MCR', 'MCRL',
    # Resistors — Susumu
    'RR', 'RG', 'KRL',
    # Capacitors / MLCCs — Murata
    'GRM', 'GRJ', 'GRC', 'GCD', 'GCE', 'GCM', 'LLL', 'KCM',
    # Capacitors / MLCCs — TDK
    'CGA',
    # Capacitors / MLCCs — Taiyo Yuden
    'EMK',
    # Inductors — Murata
    'LQH', 'LQM', 'LQG', 'LQW', 'LQP', 'MLG',
    # Ferrite beads — Murata / TDK
    'BLM', 'BLZ', 'BLL', 'MMZ',
    # Inductors — Bourns
    'SRR', 'SRU', 'SDR', 'SRL', 'SRF',
    # Crystals / resonators — Murata / Abracon
    'CSTCE', 'CSTCR', 'CSTLS', 'CSTCW', 'ABLS', 'ABM',
])

# Size-code-based SMD passive patterns (safe: RC0402, CC0603, etc.)
CHEAP_RE = [
    re.compile(r'^RC(0201|0402|0603|0805|1206|1210|2010|2512)', re.I),
    re.compile(r'^CC(0201|0402|0603|0805|1206|1210)', re.I),
]

def is_cheap(pn):
    u = pn.upper()
    if u.startswith(CHEAP_PREFIXES):
        return True
    return any(p.match(u) for p in CHEAP_RE)


# ── Cleaning helpers ─────────────────────────────────────────────────────────

_PN_TRANS = str.maketrans('', '', '()$#"\',*%@!')

def _is_empty(v):
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    return False

def clean_pn(raw):
    if _is_empty(raw):
        return ''
    s = str(raw).strip()
    if s.lower() == 'nan':
        return ''
    s = re.sub(r'[\r\n\t\xa0\u2002-\u200b\u202f\u3000]', '', s)
    s = s.replace(' ', '').translate(_PN_TRANS)
    return s[:35]

def clean_man(raw):
    if _is_empty(raw):
        return ''
    s = str(raw).strip()
    if s.lower() == 'nan':
        return ''
    s = re.sub(r'[\r\n\t\xa0]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s[:20]

def clean_text(raw):
    if raw is None:
        return ''
    return re.sub(r'[\r\n\t\xa0]+', ' ', str(raw)).strip()

def to_qty(raw):
    if _is_empty(raw):
        return None
    try:
        v = float(str(raw).replace(',', '').strip())
        if math.isnan(v) or math.isinf(v):
            return None
        i = int(v)
        return i if i == v else None
    except (ValueError, TypeError):
        return None

def has_curtiss(colD, colE):
    combined = (clean_text(colD) + ' ' + clean_text(colE)).upper()
    return 'CURTISS' in combined


# ── Main processing ──────────────────────────────────────────────────────────

print(f'Reading: {INPUT_FILE}')
wb = openpyxl.load_workbook(INPUT_FILE, read_only=True, data_only=True)
ws = wb.active

n_raw = n_blank = n_qty = n_cheap = 0
# key: (FullPartNumber, Notes) → {qty, mans}
groups = {}

for row in ws.iter_rows(values_only=True):
    n_raw += 1
    colA = row[0] if len(row) > 0 else None
    colB = row[1] if len(row) > 1 else None
    colC = row[2] if len(row) > 2 else None
    colD = row[3] if len(row) > 3 else None
    colE = row[4] if len(row) > 4 else None

    pn = clean_pn(colA)
    if not pn:
        n_blank += 1
        continue

    qty = to_qty(colC)
    if qty is None or qty < 10:
        n_qty += 1
        continue

    if is_cheap(pn):
        n_cheap += 1
        continue

    man  = clean_man(colB)
    note = ('OEM EXCESS! $2000 MIN TP REQUIRED'
            if has_curtiss(colD, colE)
            else 'OEM EXCESS! $500 MIN TP REQUIRED')

    key = (pn, note)
    if key in groups:
        groups[key]['qty'] += qty
        if man:
            groups[key]['mans'].add(man)
    else:
        groups[key] = {'qty': qty, 'mans': {man} if man else set()}

wb.close()

print(f'  Raw rows:      {n_raw:>10,}')
print(f'  Blank PN:      {n_blank:>10,}')
print(f'  QTY < 10:      {n_qty:>10,}')
print(f'  Cheap parts:   {n_cheap:>10,}')
print(f'  Unique groups: {len(groups):>10,}')


# ── Build output rows ────────────────────────────────────────────────────────

output = []
for (pn, note), v in groups.items():
    man_str = ', '.join(sorted(m for m in v['mans'] if m))[:20]
    output.append([pn, man_str, '', v['qty'], note])

output.sort(key=lambda r: r[0])


# ── Quality checks ────────────────────────────────────────────────────────────

errors = []
for i, (pn, man, dc, qty, note) in enumerate(output):
    if len(pn) > 35:
        errors.append(f'Row {i+1}: PN too long ({len(pn)}): {pn}')
    if len(man) > 20:
        errors.append(f'Row {i+1}: Man too long ({len(man)}): {man}')
    if dc != '':
        errors.append(f'Row {i+1}: DC not blank')
    if qty < 10:
        errors.append(f'Row {i+1}: QTY < 10: {qty}')
    if pn.lower() == 'nan' or man.lower() == 'nan':
        errors.append(f'Row {i+1}: NaN in PN or Man')
    if note not in ('OEM EXCESS! $500 MIN TP REQUIRED',
                    'OEM EXCESS! $2000 MIN TP REQUIRED'):
        errors.append(f'Row {i+1}: unexpected Notes value: {note}')

if errors:
    print(f'\n!!! {len(errors)} quality error(s) — first 20:')
    for e in errors[:20]:
        print('  ', e)
    sys.exit(1)

print('Quality checks: PASSED')


# ── Write backup CSV ─────────────────────────────────────────────────────────

print(f'Writing backup CSV → {OUTPUT_FILE}')
with open(OUTPUT_FILE, 'w', newline='', encoding='utf-8') as f:
    w = csv.writer(f)
    w.writerow(['FullPartNumber', 'Man', 'DC', 'QTY', 'Notes'])
    w.writerows(output)


# ── Update OEM EXCESS sheet ───────────────────────────────────────────────────

OEM_SHEET_ID  = '1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4'
CREDS_FILE    = r'G:\My Drive\Download Internet\google_credentials.json'
SHEET_TAB     = 'sheet1'
HEADER        = ['MPN', 'Man', 'DC', 'QTY', 'Notes']

print('\nConnecting to OEM EXCESS sheet...')
creds = Credentials.from_service_account_file(
    CREDS_FILE, scopes=['https://www.googleapis.com/auth/spreadsheets'])
gc  = gspread.authorize(creds)
ws  = gc.open_by_key(OEM_SHEET_ID).worksheet(SHEET_TAB)

# Read existing data to extract BILL EXT rows
print('Reading current sheet to find BILL EXT rows...')
current = ws.get_all_values()  # list of lists, row 0 = header
bill_rows = [r for r in current[1:] if 'BILL EXT' in str(r[4]).upper()]
print(f'  Current rows: {len(current)-1:,} | BILL EXT to keep: {len(bill_rows):,}')

# Build full replacement: header + bill rows + new data
# QTY stored as integer (no comma formatting needed — Sheets formats display)
new_data_rows = [[r[0], r[1], r[2], r[3], r[4]] for r in output]
all_rows = [HEADER] + bill_rows + new_data_rows
total_rows = len(all_rows)
print(f'  Total rows to write: {total_rows:,}')

# Clear sheet
print('Clearing sheet...')
ws.clear()

# Write in chunks of 10,000 rows to stay within API payload limits
CHUNK = 10000
print(f'Writing {total_rows:,} rows in chunks of {CHUNK:,}...')
for start in range(0, total_rows, CHUNK):
    chunk = all_rows[start:start + CHUNK]
    end_row = start + len(chunk)
    cell_range = f'A{start+1}:E{end_row}'
    ws.update(cell_range, chunk, value_input_option='RAW')
    print(f'  Written rows {start+1}–{end_row}')
    if end_row < total_rows:
        time.sleep(1)  # avoid Sheets API rate limit

n_curtiss = sum(1 for r in output if '$2000' in r[4])
print(f'\nDone.')
print(f'  Sheet rows: {total_rows:,} (header + {len(bill_rows)} BILL EXT + {len(new_data_rows):,} new)')
print(f'  $500 min:    {len(output)-n_curtiss:,}')
print(f'  $2000 Curtiss: {n_curtiss:,}')
