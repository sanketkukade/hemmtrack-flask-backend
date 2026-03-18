"""
============================================================
 send_email.py — Nodemailer-style Email via smtplib
============================================================
 Two SEPARATE functions:
   send_occ3_email()  → Simple text/HTML alert (no attachment)
   send_occ5_email_with_ppt() → HTML email WITH .pptx attachment
============================================================
"""

import os
import smtplib
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

log = logging.getLogger("hemmtrack")

# ── SMTP Config ──
SMTP_USER      = os.getenv("SMTP_USER", "")
SMTP_PASS      = os.getenv("SMTP_PASS", "")
ALERT_TO_EMAIL = os.getenv("ALERT_TO_EMAIL", SMTP_USER)


def _get_smtp_connection():
    """Create and return authenticated SMTP connection."""
    if not SMTP_USER or not SMTP_PASS:
        raise ValueError("SMTP_USER or SMTP_PASS not set in environment")
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(SMTP_USER, SMTP_PASS)
    return server


# ===========================================================
#  OCC = 3 — Simple Email Alert (NO attachment)
# ===========================================================
def send_occ3_email(data: dict) -> str:
    """Send OCC=3 alert email — plain text + HTML, no attachment."""

    model     = data.get("model", "Altroz / Nexon")
    failure   = data.get("failure", "Open Hem")
    station   = data.get("station", "-")
    occ       = str(data.get("occ", "03")).zfill(2)
    shift     = data.get("shift", "-")
    inspector = data.get("inspector", "-")
    side      = data.get("side", "-")
    press     = str(data.get("press", "-"))
    rc        = data.get("rc", "-")
    date_str  = data.get("date", datetime.now().strftime("%d %b %Y, %I:%M %p"))

    msg = MIMEMultipart("alternative")
    msg["From"]    = f'"🔔 HemmTrack Pro V2" <{SMTP_USER}>'
    msg["To"]      = ALERT_TO_EMAIL
    msg["Subject"] = f"🔔 OCC={occ} ALERT | {failure} | {station} | {date_str}"

    # ── Plain text ──
    text_body = f"""🔔 HIGH DEFECT ALERT — OCC={occ}

📅 Date: {date_str}
🏭 Station: {station}
🔍 Defect: {failure}
🚗 Model: {model}
⏰ Shift: {shift}
👤 Inspector: {inspector}
🚪 Side: {side}
💨 Pressure: {press} Bar
🔧 Root Cause: {rc}
📊 OCC Count: {occ}

⚡ Immediate action required!

— HemmTrack Pro V2 (Automated OCC=3 Alert)"""

    # ── HTML ──
    html_body = f"""
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
      <div style="background:#CC0000;color:#fff;padding:14px 20px;border-radius:8px 8px 0 0;">
        <h2 style="margin:0;font-size:16px;">🔔 HIGH DEFECT ALERT — OCC={occ}</h2>
        <p style="margin:4px 0 0;font-size:11px;opacity:0.9;">HemmTrack Pro V2 · {date_str}</p>
      </div>
      <div style="padding:16px 20px;background:#fff;border:1px solid #ddd;">
        <table style="width:100%;border-collapse:collapse;font-size:13px;">
          <tr style="background:#f8f9fa;">
            <td style="padding:8px 12px;border:1px solid #dee2e6;font-weight:700;width:35%;">📅 Date</td>
            <td style="padding:8px 12px;border:1px solid #dee2e6;">{date_str}</td>
          </tr>
          <tr>
            <td style="padding:8px 12px;border:1px solid #dee2e6;font-weight:700;">🏭 Station</td>
            <td style="padding:8px 12px;border:1px solid #dee2e6;color:#CC0000;font-weight:700;">{station}</td>
          </tr>
          <tr style="background:#f8f9fa;">
            <td style="padding:8px 12px;border:1px solid #dee2e6;font-weight:700;">🔍 Defect</td>
            <td style="padding:8px 12px;border:1px solid #dee2e6;">{failure}</td>
          </tr>
          <tr>
            <td style="padding:8px 12px;border:1px solid #dee2e6;font-weight:700;">🚗 Model</td>
            <td style="padding:8px 12px;border:1px solid #dee2e6;">{model}</td>
          </tr>
          <tr style="background:#f8f9fa;">
            <td style="padding:8px 12px;border:1px solid #dee2e6;font-weight:700;">⏰ Shift</td>
            <td style="padding:8px 12px;border:1px solid #dee2e6;">{shift}</td>
          </tr>
          <tr>
            <td style="padding:8px 12px;border:1px solid #dee2e6;font-weight:700;">👤 Inspector</td>
            <td style="padding:8px 12px;border:1px solid #dee2e6;">{inspector}</td>
          </tr>
          <tr style="background:#f8f9fa;">
            <td style="padding:8px 12px;border:1px solid #dee2e6;font-weight:700;">💨 Pressure</td>
            <td style="padding:8px 12px;border:1px solid #dee2e6;">{press} Bar</td>
          </tr>
          <tr>
            <td style="padding:8px 12px;border:1px solid #dee2e6;font-weight:700;">🔧 Root Cause</td>
            <td style="padding:8px 12px;border:1px solid #dee2e6;">{rc}</td>
          </tr>
          <tr style="background:#FFF2CC;">
            <td style="padding:10px 12px;border:2px solid #CC0000;font-weight:900;color:#CC0000;">📊 OCC</td>
            <td style="padding:10px 12px;border:2px solid #CC0000;font-weight:900;color:#CC0000;font-size:20px;">{occ}</td>
          </tr>
        </table>
      </div>
      <div style="background:#1F3864;color:#fff;padding:12px 20px;border-radius:0 0 8px 8px;font-size:11px;">
        <strong>HemmTrack Pro V2</strong> · OCC=3 Alert<br>
        Sanket Kukade · Weld Shop Quality Engineer · Tata Motors PVBU, Pune
      </div>
    </div>"""

    msg.attach(MIMEText(text_body, "plain"))
    msg.attach(MIMEText(html_body, "html"))

    server = _get_smtp_connection()
    server.sendmail(SMTP_USER, ALERT_TO_EMAIL, msg.as_string())
    server.quit()

    return "OCC=3 email sent"


# ===========================================================
#  OCC = 5 — Email WITH PPT Attachment
# ===========================================================
def send_occ5_email_with_ppt(data: dict, ppt_path: str) -> str:
    """Send OCC=5 escalation email with .pptx file attached."""

    model     = data.get("model", "Altroz / Nexon")
    failure   = data.get("failure", "Open Hem")
    station   = data.get("station", "-")
    occ       = str(data.get("occ", "05")).zfill(2)
    shift     = data.get("shift", "-")
    inspector = data.get("inspector", "-")
    side      = data.get("side", "-")
    press     = str(data.get("press", "-"))
    rc        = data.get("rc", "-")
    date_str  = data.get("date", datetime.now().strftime("%d %b %Y, %I:%M %p"))

    safe_date = date_str.replace(" ", "_").replace("/", "-")
    ppt_filename = f"ESCALATION_OCC{occ}_{safe_date}.pptx"

    msg = MIMEMultipart("mixed")
    msg["From"]    = f'"🚨 HemmTrack Pro V2" <{SMTP_USER}>'
    msg["To"]      = ALERT_TO_EMAIL
    msg["Subject"] = f"🚨 ESCALATION OCC={occ} | {failure} | {station} | PPT Attached | {date_str}"

    # ── Plain text ──
    text_body = f"""🚨🚨 ESCALATION ALERT — OCC={occ} 🚨🚨
⛔ MANAGEMENT ACTION REQUIRED

📅 Date: {date_str}
🏭 Station: {station}
🔍 Defect: {failure}
🚗 Model: {model}
⏰ Shift: {shift}
👤 Inspector: {inspector}
🚪 Side: {side}
💨 Pressure: {press} Bar
🔧 Root Cause: {rc}
📊 OCC Count: {occ}

📎 One-Pager PPT attached: {ppt_filename}
⚡ Immediate corrective action is MANDATORY!

— HemmTrack Pro V2 (Automated OCC=5 Escalation)"""

    # ── HTML ──
    html_body = f"""
    <div style="font-family:Arial,sans-serif;max-width:650px;margin:0 auto;">
      <div style="background:#8B0000;color:#fff;padding:16px 24px;border-radius:8px 8px 0 0;">
        <h2 style="margin:0;font-size:18px;">🚨 ESCALATION ALERT — OCC={occ}</h2>
        <p style="margin:4px 0 0;font-size:12px;opacity:0.9;">HemmTrack Pro V2 · Automated Escalation · {date_str}</p>
      </div>
      <div style="background:#FFC7CE;color:#8B0000;padding:10px 24px;font-weight:700;font-size:13px;border-left:4px solid #8B0000;border-right:4px solid #8B0000;">
        ⛔ MANAGEMENT ACTION REQUIRED — OCC Threshold Exceeded
      </div>
      <div style="padding:20px 24px;background:#fff;border:1px solid #ddd;">
        <table style="width:100%;border-collapse:collapse;font-size:13px;">
          <tr style="background:#f8f9fa;">
            <td style="padding:10px 14px;border:1px solid #dee2e6;font-weight:700;width:35%;">📅 Date</td>
            <td style="padding:10px 14px;border:1px solid #dee2e6;">{date_str}</td>
          </tr>
          <tr>
            <td style="padding:10px 14px;border:1px solid #dee2e6;font-weight:700;">🏭 Station</td>
            <td style="padding:10px 14px;border:1px solid #dee2e6;font-weight:700;color:#CC0000;">{station}</td>
          </tr>
          <tr style="background:#f8f9fa;">
            <td style="padding:10px 14px;border:1px solid #dee2e6;font-weight:700;">🔍 Defect Type</td>
            <td style="padding:10px 14px;border:1px solid #dee2e6;">{failure}</td>
          </tr>
          <tr>
            <td style="padding:10px 14px;border:1px solid #dee2e6;font-weight:700;">🚗 Model</td>
            <td style="padding:10px 14px;border:1px solid #dee2e6;">{model}</td>
          </tr>
          <tr style="background:#f8f9fa;">
            <td style="padding:10px 14px;border:1px solid #dee2e6;font-weight:700;">⏰ Shift</td>
            <td style="padding:10px 14px;border:1px solid #dee2e6;">{shift}</td>
          </tr>
          <tr>
            <td style="padding:10px 14px;border:1px solid #dee2e6;font-weight:700;">👤 Inspector</td>
            <td style="padding:10px 14px;border:1px solid #dee2e6;">{inspector}</td>
          </tr>
          <tr style="background:#f8f9fa;">
            <td style="padding:10px 14px;border:1px solid #dee2e6;font-weight:700;">🚪 Side</td>
            <td style="padding:10px 14px;border:1px solid #dee2e6;">{side}</td>
          </tr>
          <tr>
            <td style="padding:10px 14px;border:1px solid #dee2e6;font-weight:700;">💨 Pressure</td>
            <td style="padding:10px 14px;border:1px solid #dee2e6;">{press} Bar</td>
          </tr>
          <tr style="background:#f8f9fa;">
            <td style="padding:10px 14px;border:1px solid #dee2e6;font-weight:700;">🔧 Root Cause</td>
            <td style="padding:10px 14px;border:1px solid #dee2e6;">{rc}</td>
          </tr>
          <tr style="background:#FFF2CC;">
            <td style="padding:12px 14px;border:2px solid #8B0000;font-weight:900;color:#8B0000;font-size:15px;">📊 OCC Count</td>
            <td style="padding:12px 14px;border:2px solid #8B0000;font-weight:900;color:#8B0000;font-size:22px;">{occ}</td>
          </tr>
        </table>
      </div>
      <div style="background:#E8F5E9;padding:12px 24px;border:1px solid #A5D6A7;font-size:13px;color:#2E7D32;">
        📎 <strong>One-Pager PPT attached:</strong> {ppt_filename}
      </div>
      <div style="background:#1F3864;color:#fff;padding:14px 24px;border-radius:0 0 8px 8px;font-size:11px;">
        <strong>HemmTrack Pro V2</strong> · OCC=5 Automated Escalation System<br>
        Sanket Kukade · Weld Shop Quality Engineer · Tata Motors PVBU, Pune
      </div>
    </div>"""

    # Build alternative part for text/html
    alt_part = MIMEMultipart("alternative")
    alt_part.attach(MIMEText(text_body, "plain"))
    alt_part.attach(MIMEText(html_body, "html"))
    msg.attach(alt_part)

    # ── Attach PPT file ──
    if ppt_path and os.path.exists(ppt_path):
        with open(ppt_path, "rb") as f:
            ppt_part = MIMEBase("application", "vnd.openxmlformats-officedocument.presentationml.presentation")
            ppt_part.set_payload(f.read())
            encoders.encode_base64(ppt_part)
            ppt_part.add_header("Content-Disposition", f'attachment; filename="{ppt_filename}"')
            msg.attach(ppt_part)
        log.info(f"  📎 PPT attached: {ppt_filename}")
    else:
        log.warning(f"  ⚠ PPT file not found at {ppt_path} — sending without attachment")

    server = _get_smtp_connection()
    server.sendmail(SMTP_USER, ALERT_TO_EMAIL, msg.as_string())
    server.quit()

    return f"OCC=5 email sent with {ppt_filename}"
