"""Backfill today's email decisions into D1 so Training tab has data."""
import requests

HUB = "https://intransit-hub.intransit-sales.workers.dev"
SEC = "InTransit!Hub#2026"
H   = {"Authorization": f"Bearer {SEC}"}

TP500   = "We need a target price to proceed. Please note there is a $500 minimum line requirement. Once we have your target we will get back to you right away."
TP2000  = "We need a target price to proceed. Please note there is a $2,000 minimum line requirement. Once we have your target we will get back to you right away."
CHKING  = "We are checking on it now. If we get a response from the OEM, I will respond to you right away. If we do not respond back to you, please consider this a no bid. Thank you very much for the opportunity."
BILL    = "Hi, I'm passing this along to Bill who will be able to help you further."
CHKING_BUG = "[BUG — sent to wrong address (autosend@icsource.com instead of buyer)] " + CHKING

entries = [
    # thread_id,                  mpn,                    sender,                                    subject,                             draft_content
    ("19edc4110659d7ae", "IS25LP080D-JULE-TR",    "david@fortetechno.com",                  "RE: IS25LP080D-JULE-TR",            CHKING),
    ("19ed9b064b171f1d", "ISL99227FRZ-TR5784",    "XiaoYan@ranshendz.cn",                   "Re: ISL99227FRZ-TR5784",            TP500),
    ("19ed80097ef3c557", "EL9115ILZ-T7R5556",     "peter@rose.com",                          "Re: Request for Quote from your Website", TP500),
    ("19ed80097ef3c557", "EL9115ILZ-T7R5556",     "peter@rose.com",                          "Re: Request for Quote from your Website", CHKING_BUG),
    ("19edbf727b373f74", "BCM54210SB0IMLG",       "linda.lanni@converge.com",                "Re: BCM54210SB0IMLG",               CHKING),
    ("19ed61372352c067", "12110440-B",             "cameron@express-technology.com",          "Re: 12110440-B",                    CHKING),
    ("19edc041af365883", "9ZXL0851EKILFT",         "NarendrababuRajaraman.Naren@converge.com","Re: 9ZXL0851EKILFT",               BILL),
    ("19edbf3480e3a6d6", "MX52LM08A11XVI",         "danielle@bulechip.com",                  "Re: MX52LM08A11XVI",                TP500),
    ("19edbdb36ae38e9b", "IS42S16160J-7TLI",       "erickc@northshorecomponents.com",         "Re: IS42S16160J-7TLI",              TP500),
    ("19edbdb36ae38e9b", "IS42S16160J-7TLI",       "erickc@northshorecomponents.com",         "Re: IS42S16160J-7TLI",              CHKING),
    ("19edbdd5cde87bd3", "M83536/32-002L",          "jgallegos@accu-sembly.com",              "Re: Request for Quote from your Website", TP500),
    ("19edbd1314e4e075", "AT25DF321A-SH-T",         "jacquelineg@chip1.com",                  "Re: AT25DF321A-SH-T",               BILL),
    ("19ed7224dbfea92e", "8302401EX",               "luzmaria.orozco@aeri.com",               "Re: 8302401EX",                     TP500),
    ("19ed7224dbfea92e", "8302401EX",               "luzmaria.orozco@aeri.com",               "Re: 8302401EX",                     CHKING),
    ("19ed88fb70a0e9e0", "OP2177ARZ-REEL7",         "lark@hkyx-ic.com",                       "Re: OP2177ARZ-REEL7",               TP500),
    ("19ed8db61cf874a7", "M25P16-VMW6TG",           "amelia@hkyx-ic.com",                     "Re: M25P16-VMW6TG",                 TP500),
    ("19edb3b06ff5009e", "MP3V5050GP",              "sells@zztchip.com",                      "Re: MP3V5050GP",                    BILL),
    ("19edab9b1339598f", "MMQA33VT1G",              "andre@nfsmith.nl",                       "Re: MMQA33VT1G",                    TP500),
    ("19edb40cf2d80f9d", "1544650-2",               "gabriela.alvez@elan-sol.com",            "Re: 1544650-2",                     TP500),
    ("19ed9fdb3398d522", "W9812G6KH-6I",            "amelia@hkyx-ic.com",                     "Re: W9812G6KH-6I",                  TP500),
    ("19ed9b3cec6fac73", "7047-3249",               "pur@hermeses.com",                       "Re: 7047-3249",                     TP500),
    ("19ed7e328ccfff22", "2-968321-4",              "dtorres@stclairtech.com",                "Re: 2-968321-4",                    TP500),
]

ok = err = 0
for thread_id, mpn, sender, subject, draft in entries:
    r = requests.post(
        f"{HUB}/api/drafts",
        json={"thread_id": thread_id, "mpn": mpn, "sender": sender,
              "subject": subject, "draft_content": draft},
        headers=H, timeout=10,
    )
    if r.status_code == 200:
        ok += 1
        print(f"  OK  {mpn[:20]:20s}  {sender[:35]}")
    else:
        err += 1
        print(f"  ERR {r.status_code}  {mpn}  {r.text[:80]}")

print(f"\n{ok} inserted, {err} errors.")
