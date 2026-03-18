"""
============================================================
 send_telegram.py — Telegram Bot Alert Sender
============================================================
 Two SEPARATE functions:
   send_occ3_telegram() → Standard alert message
   send_occ5_telegram() → Escalation message (different format)
============================================================
"""

import os
import logging
import requests

log = logging.getLogger("hemmtrack")

TG_TOKEN   = os.getenv("TELEGRAM_BOT_TOKEN", "")
TG_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")
TG_API_URL = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"


def _send_tg_message(text: str) -> dict:
    """Send a message via Telegram Bot API."""
    if not TG_TOKEN or not TG_CHAT_ID:
        log.warning("Telegram not configured — skipping")
        return {"ok": False, "reason": "not_configured"}

    resp = requests.post(TG_API_URL, json={
        "chat_id": TG_CHAT_ID,
        "text": text,
        "parse_mode": "HTML"
    }, timeout=10)

    result = resp.json()
    if not result.get("ok"):
        log.error(f"Telegram API error: {result}")
    return result


# ===========================================================
#  OCC = 3 — Standard Telegram Alert
# ===========================================================
def send_occ3_telegram(data: dict) -> dict:
    """
    OCC=3 Telegram alert — matches the existing frontend format.
    Same message structure as the XHR-based alert.
    """
    model     = data.get("model", "Altroz / Nexon")
    failure   = data.get("failure", "Open Hem")
    station   = data.get("station", "-")
    occ       = str(data.get("occ", "03")).zfill(2)
    shift     = data.get("shift", "-")
    inspector = data.get("inspector", "-")
    date_str  = data.get("date", "-")

    message = (
        "🚨 Alarm Escalation 🚨\n"
        "🛑 200 Demerit 🛑\n"
        "\n"
        f"📅 Date: {date_str}\n"
        f"Issue: {failure}\n"
        f"Model: {model}\n"
        f"🏭 Station: {station}\n"
        f"Shift: {shift}\n"
        f"OCC: {occ}\n"
        f"Inspector ID: {inspector}\n"
        f"Reported At: Slat Buy-Off\n"
        "\n"
        "⚡ Immediate action required!\n"
        "— HemmTrack Pro V2"
    )

    return _send_tg_message(message)


# ===========================================================
#  OCC = 5 — Escalation Telegram Alert (different format)
# ===========================================================
def send_occ5_telegram(data: dict) -> dict:
    """
    OCC=5 Telegram escalation — more urgent format.
    Includes PPT + email confirmation.
    COMPLETELY SEPARATE from OCC=3.
    """
    model     = data.get("model", "Altroz / Nexon")
    failure   = data.get("failure", "Open Hem")
    station   = data.get("station", "-")
    occ       = str(data.get("occ", "05")).zfill(2)
    shift     = data.get("shift", "-")
    inspector = data.get("inspector", "-")
    date_str  = data.get("date", "-")
    press     = str(data.get("press", "-"))
    rc        = data.get("rc", "-")

    message = (
        f"🚨🚨 ESCALATION ALERT — OCC={occ} 🚨🚨\n"
        "⛔ MANAGEMENT ACTION REQUIRED ⛔\n"
        "\n"
        f"📅 Date: {date_str}\n"
        f"Issue: {failure}\n"
        f"Model: {model}\n"
        f"🏭 Station: {station}\n"
        f"Shift: {shift}\n"
        f"OCC: {occ} (ESCALATION LEVEL)\n"
        f"Inspector: {inspector}\n"
        f"💨 Pressure: {press} Bar\n"
        f"🔧 Root Cause: {rc}\n"
        f"Reported At: Slat Buy-Off\n"
        "\n"
        "📎 PPT One-Pager emailed with attachment\n"
        "📧 Escalation email sent via backend\n"
        "\n"
        "⚡ Immediate corrective action is MANDATORY!\n"
        "— HemmTrack Pro V2 (Flask Backend)"
    )

    return _send_tg_message(message)
