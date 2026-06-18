"""Seed default app configs into D1 via /api/configs."""
import requests

HUB = "https://intransit-hub.intransit-sales.workers.dev"
SEC = "InTransit!Hub#2026"
H = {"Authorization": f"Bearer {SEC}"}

defaults = {
    "email_automation": {
        "enabled": True,
        "MSG_NEED_TP_500":  "Hi,\n\nWe'd be happy to help! To process this request we require a target price with a minimum order value of $500.\n\nPlease let us know your target price and we'll get you a quote.\n\nBest regards,",
        "MSG_NEED_TP_2000": "Hi,\n\nWe'd be happy to help! To process this request we require a target price with a minimum order value of $2,000.\n\nPlease let us know your target price and we'll get you a quote.\n\nBest regards,",
        "MSG_CHECKING":     "Hi,\n\nThank you for reaching out. I'm checking on availability for this part and will get back to you shortly.\n\nBest regards,",
        "MSG_BILL":         "Hi, I'm passing this along to Bill who will be able to help you further. Best regards,",
        "DAVID_EMAIL":      "david@intransittech.com",
        "BILL_EMAIL":       "bill@intransittech.com",
        "DEB_EMAIL":        "deb@intransittech.com"
    },
    "tee_time_bot": {
        "enabled": True,
        "booking_hour":   7,
        "booking_minute": 0,
        "days_out":       7
    },
    "icsource_checker": {"enabled": True},
    "oem_excess":       {"enabled": True}
}

for app, cfg in defaults.items():
    r = requests.post(f"{HUB}/api/configs/{app}", json={"config": cfg}, headers=H, timeout=10)
    print(f"  {r.status_code}  {app}")

print("\nDone.")
