# Intransit Technologies — Claude Code Session Protocols
# John Fluman | john.fluman@intransittech.com
# LOAD THIS FULLY BEFORE ANY ACTION. These rules are non-negotiable.

## WHO I AM WORKING WITH
John Fluman, owner of Intransit Technologies — electronic components distributor (OEM excess, ISO 9001).
Phone: Toll (877) 677-5868 x101 · Local (949) 481-7935 x101

---

## CRITICAL RULES — NEVER VIOLATE

1. **NEVER add to Forte without a QTY.** No exceptions. No QTY = do not add.
2. **Forte history column (J):** `addToForteSheet()` auto-populates col J with prior history when a duplicate MPN is added. Do NOT manually pass a custom historyNote unless John explicitly specifies one.
3. **ALWAYS check OEM EXCESS via web app** before claiming a part is or isn't available. Never assume.
4. **ALWAYS include John's signature** on every manually created Gmail draft.
5. **ALWAYS read the full thread** before any email work. Never act on a snippet.
6. **60-day duplicate check** before adding any MPN to Forte. If already there within 60 days, skip.
7. **Skip @intransittech.com threads.** Internal emails need no draft help.
8. **"Checking on it" sent = Forte entry required** — even if buyer gave no formal TP.
9. **No advice blocks in drafts.** Do not add yellow "Note for John" boxes or [ADVICE: ...] blocks to any email draft — John stopped using this feature.

---

## INFRASTRUCTURE

| Resource | Value |
|---|---|
| Worker URL | `https://intransit-hub.intransit-sales.workers.dev` |
| HUB_SECRET | `InTransit!Hub#2026` |
| FORTE_SHEET_ID | `1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4` |
| OEM_EXCESS SHEET | `1FSYIiFFEd5jrSNoxngjI0d8ZI3Qfyq_c8GzfcK6XQu4` |
| IN_STOCK_ID | `1iOFHUBiWRgA6EjtO2ujoGpz-8v1qTRkgCXSvCa2Gf54` |
| STAN_SHEET_ID | `1pGRDpkqftQNoEYna53MxRJfUY8jEf5_w32FNa56OUIM` |
| Apps Script email file | `C:\Users\fluma\intransit-hub\email_script_v24_hub.js` |
| Forte one-time script | `C:\Users\fluma\intransit-hub\add_forte_rows_onetime.js` |
| Google Drive MCP | Read-only — cannot write/append rows to Sheets |

## OEM EXCESS WEB APP
```
https://script.google.com/macros/s/AKfycbyuuBmiYVW5mKI82D5YQGPh1nNGLJZzlLKoxuOdtmOUwUe75VlhhakqgwKooZu5LHFK/exec?key=baSDJ%23444FE%268&mpn=[MPN]
```
- Always use PowerShell `Invoke-WebRequest -MaximumRedirection 5` — WebFetch gets stuck in redirects.
- Returns JSON: `{ oem_excess: [...], in_stock: [...], stan_sheet: [...], forte_sheet: [...] }`

---

## FORTE SHEET — COLUMN LAYOUT

| Col | Field | Notes |
|---|---|---|
| A | Date | `M/d/yyyy` format |
| B | MPN | Exact as given |
| C | Qty | Buyer's requested quantity — REQUIRED |
| D | Buyer TP | Target price buyer gave |
| E | John Buy | Leave blank |
| F | Country | 2-letter code (CA, US, NL, etc.) |
| G | Potential | `=C[row]*D[row]` formula |
| H | John Quoted | Leave blank |
| I | Notes | Leave blank |
| J | History | Auto-populated by `addToForteSheet()` for duplicate MPNs (prior entries, most recent first). Blank on first occurrence. |
| K | Status | `Open` |

### addToForteSheet() signature
```javascript
addToForteSheet(mpn, qty, targetPrice, country, historyNote)
// historyNote: pass '' for normal calls; addToForteSheet() auto-builds col J from prior entries
// If a prior entry exists for this MPN, col J is auto-populated with all prior rows (most recent first)
```

### 60-day check
```javascript
checkForteForMPN(mpn, 60)  // 60 days, not 90
```

---

## EMAIL WORKFLOW — RFQ PROCESSING

### Step 1: Check OEM EXCESS
Run the web app PowerShell call. If found → proceed. If not found → no bid (silent, no reply needed).

### Step 2: If found in OEM EXCESS → Draft MSG_CHECKING
Reply to the buyer's last message in the thread. Use exact text:

```
We are checking on it now. If we get a response from the OEM, I will respond to you right away. If we do not respond back to you, please consider this a no bid. Thank you very much for the opportunity.
```

Always use the full HTML signature with disclaimer on every draft:
```html
<div><b><span style="color:rgb(31,73,125);font-family:Tahoma,sans-serif;font-size:10pt">Regards,</span></b></div>
<div><b><span style="color:rgb(31,73,125);font-family:Tahoma,sans-serif;font-size:10pt">John Fluman</span></b></div>
<div><b><span style="color:rgb(31,73,125);font-family:Arial,sans-serif;font-size:8pt">Intransit Technologies</span></b></div>
<div><a href="mailto:john.fluman@intransittech.com" style="font-family:Calibri;font-size:8pt">john.fluman@intransittech.com</a></div>
<div><i><span style="color:gray;font-family:Arial,sans-serif;font-size:7.5pt">An ISO 9001 Certified Company</span></i></div>
<div><span style="color:rgb(31,73,125);font-family:Tahoma,sans-serif;font-size:8pt">Toll (877) 677-5868 x101 - Local (949) 481-7935 x101</span></div>
<br>
<div><span style="color:rgb(166,166,166);font-family:Calibri,sans-serif;font-size:8pt">The information contained in this communication and its attachment(s) is intended only for the use of the individual to whom it is addressed and may contain information that is privileged, confidential, or exempt from disclosure. If the reader of this message is not the intended recipient, you are hereby notified that any dissemination, distribution, or copying of this communication is strictly prohibited. If you have received this communication in error, please notify <a href="mailto:john.fluman@intransittech.com" style="font-family:Calibri;font-size:8pt">john.fluman@intransittech.com</a> and delete the communication without retaining any copies. Thank you.</span></div>
```

### Step 3: Add to Forte
- qty = buyer's requested quantity
- targetPrice = buyer's TP
- country = buyer's country (2-letter)
- historyNote = '' (always blank)
- Check 60-day duplicate first

### When MSG_CHECKING is sent = always add to Forte (even without formal buyer TP)

### BILL parts — OEM EXCESS tagged "BILL EXT [number]"
- Do NOT add to Forte
- Do NOT send MSG_CHECKING to buyer
- After buyer gives TP → forward inquiry to Bill (bill.pratt@intransittech.com) and reply "Bill will help with this request"

---

## STANDARD REPLY TEMPLATES

**TP Request (no TP given):**
> We need a target price to proceed. Please note there is a $500 minimum line requirement. Once we have your target we will get back to you right away.

**$2,000 min version:**
> We need a target price to proceed. Please note there is a $2,000 minimum line requirement. Once we have your target we will get back to you right away.

**Checking on it (MSG_CHECKING):**
> We are checking on it now. If we get a response from the OEM, I will respond to you right away. If we do not respond back to you, please consider this a no bid. Thank you very much for the opportunity.

**Bill will help:**
> Bill will help with this request
> TO: buyer | CC: bill.pratt@intransittech.com — always CC Bill so he sees future buyer replies directly

---

## GMAIL THREAD HANDLING

- **Always fetch FULL_CONTENT** for threads before acting. Never rely on the snippet.
- **Large threads**: use MINIMAL first to map message IDs, then FULL_CONTENT on specific messages.
- **Reply target**: always `replyToMessageId` = the last message in the thread (buyer's most recent reply).
- **Internal threads** (@intransittech.com): skip entirely, no drafts needed.
- **oem-payment-forwarded label**: thread-level label. Always check `thread.getLabels()` before re-forwarding — Gmail's `-label:` query can re-match a thread when new messages arrive.

---

## GMAIL LABELS (key ones)

| Label | Purpose |
|---|---|
| `oem-rfq-incoming-processed` | RFQ was seen by automation |
| `oem-tp-processed` | Buyer TP reply was processed |
| `oem-payment-forwarded` | Payment advice forwarded to Deb |
| `oem-nostock-seen` | Checked, not in stock |
| `BATCH-FI` | Finance batch |

---

## EMAIL SCRIPT (Apps Script v25 — ACTIVE as of 2026-06-19)

File: `C:\Users\fluma\intransit-hub\email_script_v24_hub.js` (filename is v24, content is v25)

Key functions:
- `addToForteSheet(mpn, qty, targetPrice, country, historyNote)` — adds row to Forte
- `checkForteForMPN(mpn, 60)` — 60-day duplicate check
- `checkInboxForPaymentAdvice()` — Trigger 6, forwards payment advice to Deb
- `checkInboxForTPReplies()` — Trigger 4, processes TP replies
- `checkInboxForNewRFQs()` — Trigger 3, scans for new RFQs
- `processFixQueue()` — Trigger 9, repairs Gmail drafts queued via /api/fix-queue
- `processCommandQueue()` — Trigger 10, runs inventory commands (remove stock, send NetCOMPONENTS email)
- `setupTriggers()` — installs all 10 triggers (run once after script update)

BLOCKED_DOMAINS: `['sourceschip.com', 'bulechip.com', 'feelchips.com', 'chip-wintrading.com', 'qizhongsmart.com', 'lcgyzzonline.com']`

When John pastes a new version of the script into Apps Script editor:
1. Paste full file content
2. Run `setupTriggers()` once to reinstall all triggers

---

## PEOPLE

| Name | Email | Role |
|---|---|---|
| John Fluman | john.fluman@intransittech.com | Owner |
| Deb MacDonald | deb@intransittech.com | Finance |
| Bill Pratt | bill.pratt@intransittech.com | Sales |
| David | david@fortetechno.com | OEM Excess supplier |

---

## COMMON BUYERS (reference)

| Buyer | Email | Country | Notes |
|---|---|---|---|
| Alexio Canada | alexiocanada109@gmail.com | CA | IC Source / netCOMPONENTS |
| NF Smith (André) | andre@nfsmith.nl | NL | netCOMPONENTS |
| Jerry Gallegos / Accu-sembly | jgallegos@accu-sembly.com | US | Website RFQ via WebsiteRFQ@ |

---

## WRITING ONE-TIME FORTE SCRIPTS

When Google Drive MCP can't write to Sheets, create a one-time Apps Script function:

```javascript
function addMissingRows() {
  var FORTE_SHEET_ID = '1DbZsEC8AsZY8BGpBils7toGf517jn-oqT0MUNyTi_e4';
  var sheet = SpreadsheetApp.openById(FORTE_SHEET_ID).getSheets()[0];
  var today = 'M/D/YYYY';  // use actual date
  var rows = [
    // [mpn, qty, buyerTP, country]  — history is always ''
    ['MPN-HERE', qty, tp, 'CC'],
  ];
  rows.forEach(function(r) {
    var nextRow = sheet.getLastRow() + 1;
    sheet.appendRow([today, r[0], r[1], r[2], '', r[3],
      '=C' + nextRow + '*D' + nextRow, '', '', '', 'Open']);
    Logger.log('Added: ' + r[0]);
  });
}
```

Key: pass `''` for historyNote — the function auto-populates col J with prior history if the MPN is a duplicate.

---

## SESSION START CHECKLIST
Before doing anything else, mentally confirm:
- [ ] Have I read the full thread (not just the snippet)?
- [ ] Have I checked OEM EXCESS via the web app?
- [ ] Do I have a QTY before touching Forte?
- [ ] Is historyNote passed as '' (auto-history is built automatically for duplicates)?
- [ ] Does the draft have John's signature?
- [ ] Did I do a 60-day Forte duplicate check?
