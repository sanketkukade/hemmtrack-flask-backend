"""
HemmTrack Pro V2 — Flask Backend
=================================
Endpoint: POST /check_occ
  → Receives defect alert data from frontend
  → Generates One-Pager PPT (when OCC >= 5)
  → Sends email with PPT attachment via Gmail SMTP
  → Sends Telegram text alert + PPT document

Railway Environment Variables Required:
  SMTP_EMAIL       = sanketkukade111@gmail.com
  SMTP_PASSWORD    = <Gmail App Password (16 chars, no spaces)>

Optional overrides:
  SMTP_HOST        = smtp.gmail.com   (default)
  SMTP_PORT        = 587              (default)
  TELEGRAM_BOT_TOKEN = (has default from frontend)
  TELEGRAM_CHAT_ID   = (has default from frontend)
"""

import os
import io
import time
import json
import logging
import smtplib
import requests
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

from flask import Flask, request, jsonify
from flask_cors import CORS
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ═══════════════════════════════════════════════════
# LOGGING
# ═══════════════════════════════════════════════════
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
log = logging.getLogger('HemmTrack')

app = Flask(__name__)
CORS(app, origins=[
    "https://sanketkukade.github.io",
    "http://127.0.0.1:5500",
    "http://localhost:5500",
    "http://localhost:3000",
    "http://localhost:8080",
    "null",   # file:// protocol sends Origin: null
])


# ═══════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════
def get_smtp_config():
    """Read SMTP config from Railway environment variables.
    
    Handles variable name mismatches:
      Railway has ALERT_TO_EMAIL  → maps to SMTP email
      Railway has RESEND_API_KEY  → maps to SMTP password (Gmail App Password)
    
    Normalizes Gmail App Password:
      "bekc unmr ywxe vyad" → "bekcunmrywxevyad"
    """
    # Smart fallback chain: try exact name first, then alternatives
    email = (
        os.environ.get('SMTP_EMAIL', '').strip()
        or os.environ.get('SMTP_USER', '').strip()
        or os.environ.get('ALERT_TO_EMAIL', '').strip()
    )

    # Password: try all possible var names, then strip ALL spaces
    raw_password = (
        os.environ.get('SMTP_PASSWORD', '').strip()
        or os.environ.get('SMTP_PASS', '').strip()
        or os.environ.get('RESEND_API_KEY', '').strip()
        or os.environ.get('GMAIL_APP_PASSWORD', '').strip()
    )
    # Gmail App Passwords are displayed as "xxxx xxxx xxxx xxxx"
    # but SMTP auth requires "xxxxxxxxxxxxxxxx" (no spaces)
    password = raw_password.replace(' ', '')

    host = os.environ.get('SMTP_HOST', 'smtp.gmail.com').strip()
    port = int(os.environ.get('SMTP_PORT', '587'))

    log.info("📧 SMTP Config Check:")
    log.info(f"   SMTP_EMAIL    = {'✅ ' + email[:5] + '***' if email else '❌ NOT SET'}")
    if email and not os.environ.get('SMTP_EMAIL'):
        log.info(f"   (resolved from ALERT_TO_EMAIL)")
    log.info(f"   SMTP_PASSWORD = {'✅ SET (' + str(len(password)) + ' chars)' if password else '❌ NOT SET'}")
    if password and not os.environ.get('SMTP_PASSWORD'):
        log.info(f"   (resolved from RESEND_API_KEY, spaces removed)")
    log.info(f"   SMTP_HOST     = {host}")
    log.info(f"   SMTP_PORT     = {port}")

    return {
        'email': email,
        'password': password,
        'host': host,
        'port': port,
        'valid': bool(email and password)
    }


def get_telegram_config():
    """Read Telegram config — defaults match the frontend tokens."""
    token = os.environ.get(
        'TELEGRAM_BOT_TOKEN',
        '8626407803:AAHaOOIK_UILf5kTpYrBPt9yVHSs8y713TE'
    ).strip()
    chat_id = os.environ.get('TELEGRAM_CHAT_ID', '823556812').strip()
    return {'token': token, 'chat_id': chat_id}


# ═══════════════════════════════════════════════════
# PPT GENERATION
# ═══════════════════════════════════════════════════
def generate_one_pager_ppt(data):
    """Generate a One-Pager PPTX and return as bytes."""
    log.info("📊 Generating One-Pager PPT...")

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank

    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0x06, 0x0A, 0x12)

    # Title
    title_box = slide.shapes.add_textbox(
        Inches(0.3), Inches(0.2), Inches(12.7), Inches(0.7)
    )
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🚨 HEMMING DEFECT — ONE PAGER ALERT"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x00, 0xE5, 0xFF)
    p.alignment = PP_ALIGN.CENTER

    # Sub-header
    sub_box = slide.shapes.add_textbox(
        Inches(0.3), Inches(0.9), Inches(12.7), Inches(0.4)
    )
    tf = sub_box.text_frame
    p = tf.paragraphs[0]
    p.text = (
        f"Date: {data.get('date', '-')}  |  "
        f"Station: {data.get('station', '-')}  |  "
        f"Model: {data.get('model', '-')}  |  "
        f"Shift: {data.get('shift', '-')}"
    )
    p.font.size = Pt(14)
    p.font.color.rgb = RGBColor(0x4A, 0x6F, 0xA5)
    p.alignment = PP_ALIGN.CENTER

    # Info Grid
    info_items = [
        ("OCC Count",        data.get('occ', '-'),               "🔴"),
        ("Failure Mode",     data.get('failure', '-'),            "🔍"),
        ("Root Cause",       data.get('rc', '-'),                 "🔧"),
        ("Hemming Pressure", f"{data.get('press', '-')} Bar",    "💨"),
        ("Side",             data.get('side', '-'),               "🚪"),
        ("Inspector ID",     data.get('inspector', '-'),          "👤"),
        ("RAYBG Status",     data.get('raybg', '-'),              "📋"),
        ("Corrective Action",data.get('actions', '-'),            "⚡"),
        ("ECD (Target)",     data.get('ecd', '-'),                "📅"),
    ]

    y_start = 1.5
    col_width = 4.1
    row_height = 0.7

    for i, (label, value, icon) in enumerate(info_items):
        col = i % 3
        row = i // 3
        x = 0.5 + col * col_width
        y = y_start + row * row_height

        box = slide.shapes.add_textbox(
            Inches(x), Inches(y), Inches(3.8), Inches(0.6)
        )
        tf = box.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = f"{icon} {label}"
        p.font.size = Pt(10)
        p.font.color.rgb = RGBColor(0x4A, 0x6F, 0xA5)
        p.font.bold = True

        p2 = tf.add_paragraph()
        p2.text = str(value)
        p2.font.size = Pt(16)
        p2.font.bold = True
        p2.font.color.rgb = (
            RGBColor(0xFF, 0x3D, 0x5A) if label == "OCC Count"
            else RGBColor(0xD4, 0xE8, 0xFF)
        )

    # History Table
    history = data.get('highHistory', [])
    if history:
        tbl_y = 4.0
        hist_title = slide.shapes.add_textbox(
            Inches(0.5), Inches(tbl_y - 0.4), Inches(8), Inches(0.35)
        )
        tf = hist_title.text_frame
        p = tf.paragraphs[0]
        p.text = "📋 Recent HIGH Defect History (Last 5)"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0xFF, 0xAA, 0x00)

        cols = 5
        rows = min(len(history), 5) + 1
        table_shape = slide.shapes.add_table(
            rows, cols,
            Inches(0.5), Inches(tbl_y),
            Inches(12.3), Inches(0.35 * rows)
        )
        table = table_shape.table

        headers = ['Date', 'Station', 'Defect', 'Root Cause', 'Pressure']
        for j, h in enumerate(headers):
            cell = table.cell(0, j)
            cell.text = h
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(9)
                paragraph.font.bold = True
                paragraph.font.color.rgb = RGBColor(0x00, 0xE5, 0xFF)

        for r, entry in enumerate(history[:5]):
            vals = [
                entry.get('date', '-'),
                entry.get('station', '-'),
                entry.get('defect', '-'),
                entry.get('rootCause', '-'),
                str(entry.get('pressure', '-'))
            ]
            for j, v in enumerate(vals):
                cell = table.cell(r + 1, j)
                cell.text = v
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.size = Pt(9)
                    paragraph.font.color.rgb = RGBColor(0xD4, 0xE8, 0xFF)

    # Footer
    footer_box = slide.shapes.add_textbox(
        Inches(0.3), Inches(6.8), Inches(12.7), Inches(0.5)
    )
    tf = footer_box.text_frame
    p = tf.paragraphs[0]
    p.text = (
        "Generated by HemmTrack Pro V2  |  "
        "Sanket Kukade — M.Tech Manufacturing Engg, DYPIU Pune  |  "
        "Tata Motors PVBU"
    )
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(0x4A, 0x6F, 0xA5)
    p.alignment = PP_ALIGN.CENTER

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    ppt_bytes = buffer.getvalue()

    log.info(f"📊 PPT generated: {len(ppt_bytes)} bytes ({len(ppt_bytes)//1024} KB)")
    return ppt_bytes


# ═══════════════════════════════════════════════════
# EMAIL — Gmail SMTP with PPT attachment
# ═══════════════════════════════════════════════════
def send_email_with_ppt(smtp_cfg, data, ppt_bytes=None):
    """Send alert email via SMTP. Attach PPT if provided."""

    if not smtp_cfg['valid']:
        log.error("❌ SMTP credentials missing! Set SMTP_EMAIL and SMTP_PASSWORD in Railway.")
        log.error("   Railway Dashboard → Your Service → Variables → Add:")
        log.error("     SMTP_EMAIL    = sanketkukade111@gmail.com")
        log.error("     SMTP_PASSWORD = <Gmail App Password>")
        return 'skipped: no SMTP credentials'

    try:
        to_email = smtp_cfg['email']
        from_email = smtp_cfg['email']

        msg = MIMEMultipart()
        msg['From'] = f"HemmTrack Pro <{from_email}>"
        msg['To'] = to_email

        occ = data.get('occ', '?')
        station = data.get('station', '-')
        failure = data.get('failure', 'Open Hem')
        has_ppt = ppt_bytes is not None

        if has_ppt:
            msg['Subject'] = f"📊 ONE PAGER — OCC {occ} | {failure} | {station} — HemmTrack"
        else:
            msg['Subject'] = f"🚨 OCC ALERT {occ} — {failure} | {station} — HemmTrack"

        html_body = f"""
        <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;
                    background:#060a12;color:#d4e8ff;padding:24px;border-radius:12px;">
          <div style="text-align:center;padding:16px;
                      background:linear-gradient(135deg,#0a1628,#0f1928);
                      border-radius:10px;margin-bottom:16px;">
            <h1 style="color:#00e5ff;margin:0;font-size:22px;">
              {'📊 ONE PAGER ALERT' if has_ppt else '🚨 OCC ALERT'}
            </h1>
            <p style="color:#4a6fa5;margin:4px 0 0;font-size:12px;">
              HemmTrack Pro V2 — Automated Alert
            </p>
          </div>
          <table style="width:100%;border-collapse:collapse;margin:16px 0;">
            <tr><td style="padding:8px;color:#4a6fa5;">📅 Date</td>
                <td style="padding:8px;color:#d4e8ff;font-weight:700;">{data.get('date','-')}</td></tr>
            <tr><td style="padding:8px;color:#4a6fa5;">🔴 OCC Count</td>
                <td style="padding:8px;color:#ff3d5a;font-weight:900;font-size:18px;">{occ}</td></tr>
            <tr><td style="padding:8px;color:#4a6fa5;">🏭 Station</td>
                <td style="padding:8px;color:#d4e8ff;font-weight:700;">{station}</td></tr>
            <tr><td style="padding:8px;color:#4a6fa5;">🔍 Failure</td>
                <td style="padding:8px;color:#ff3d5a;font-weight:700;">{failure}</td></tr>
            <tr><td style="padding:8px;color:#4a6fa5;">🚪 Side</td>
                <td style="padding:8px;color:#d4e8ff;">{data.get('side','-')}</td></tr>
            <tr><td style="padding:8px;color:#4a6fa5;">🔧 Root Cause</td>
                <td style="padding:8px;color:#d4e8ff;">{data.get('rc','-')}</td></tr>
            <tr><td style="padding:8px;color:#4a6fa5;">💨 Pressure</td>
                <td style="padding:8px;color:#d4e8ff;">{data.get('press','-')} Bar</td></tr>
            <tr><td style="padding:8px;color:#4a6fa5;">👤 Inspector</td>
                <td style="padding:8px;color:#d4e8ff;">{data.get('inspector','-')}</td></tr>
            <tr><td style="padding:8px;color:#4a6fa5;">⚡ Action</td>
                <td style="padding:8px;color:#d4e8ff;">{data.get('actions','-')}</td></tr>
          </table>
          {'<p style="text-align:center;color:#3e86f6;font-weight:700;">📎 One Pager PPT attached</p>' if has_ppt else ''}
          <div style="text-align:center;padding:12px;margin-top:16px;border-top:1px solid #1e3050;">
            <p style="color:#4a6fa5;font-size:10px;margin:0;">
              Sanket Kukade — M.Tech Manufacturing Engg, DYPIU Pune<br>
              Tata Motors PVBU — HemmTrack Pro V2
            </p>
          </div>
        </div>
        """

        msg.attach(MIMEText(html_body, 'html'))

        if has_ppt:
            filename = (
                f"OnePager_OCC{occ}_{station}_"
                f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
            )
            part = MIMEBase(
                'application',
                'vnd.openxmlformats-officedocument.presentationml.presentation'
            )
            part.set_payload(ppt_bytes)
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
            msg.attach(part)
            log.info(f"📎 PPT attached: {filename} ({len(ppt_bytes)//1024} KB)")

        log.info(f"📤 Connecting to {smtp_cfg['host']}:{smtp_cfg['port']}...")

        with smtplib.SMTP(smtp_cfg['host'], smtp_cfg['port'], timeout=15) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            log.info(f"🔐 STARTTLS OK. Authenticating as {smtp_cfg['email'][:5]}***...")
            server.login(smtp_cfg['email'], smtp_cfg['password'])
            log.info("✅ SMTP Login successful!")
            server.sendmail(from_email, [to_email], msg.as_string())
            log.info(f"✅ Email sent to {to_email}")

        return 'sent'

    except smtplib.SMTPAuthenticationError as e:
        log.error(f"❌ SMTP Auth Failed: {e}")
        log.error("   → Gmail needs an App Password, NOT your regular password.")
        log.error("   → https://myaccount.google.com/apppasswords")
        return f'auth_failed: {str(e)[:80]}'
    except smtplib.SMTPException as e:
        log.error(f"❌ SMTP Error: {e}")
        return f'smtp_error: {str(e)[:80]}'
    except Exception as e:
        log.error(f"❌ Email Error: {e}")
        return f'error: {str(e)[:80]}'


# ═══════════════════════════════════════════════════
# TELEGRAM
# ═══════════════════════════════════════════════════
def send_telegram_alert(tg_cfg, data, ppt_bytes=None):
    """Send Telegram text + optional PPT document."""

    token = tg_cfg['token']
    chat_id = tg_cfg['chat_id']
    base_url = f"https://api.telegram.org/bot{token}"

    occ = data.get('occ', '?')
    has_ppt = ppt_bytes is not None

    text = (
        f"{'📊 ONE PAGER ALERT' if has_ppt else '🚨 OCC ALERT'}\n"
        f"{'🛑 200 Demerit 🛑' if not has_ppt else ''}\n\n"
        f"📅 Date: {data.get('date', '-')}\n"
        f"🔴 OCC: {occ}\n"
        f"Issue: {data.get('failure', 'Open Hem')}\n"
        f"Model: {data.get('model', '-')}\n"
        f"🏭 Station: {data.get('station', '-')}\n"
        f"🚪 Side: {data.get('side', '-')}\n"
        f"Shift: {data.get('shift', '-')}\n"
        f"👤 Inspector: {data.get('inspector', '-')}\n"
        f"🔧 Root Cause: {data.get('rc', '-')}\n"
        f"💨 Pressure: {data.get('press', '-')} Bar\n"
        f"⚡ Action: {data.get('actions', '-')}\n\n"
        f"{'📎 PPT One Pager attached below' if has_ppt else '⚡ Immediate action required!'}\n"
        f"— HemmTrack Pro V2"
    )

    text_status = 'error'
    doc_status = None

    try:
        log.info(f"📱 Sending Telegram message to chat {chat_id}...")
        resp = requests.post(
            f"{base_url}/sendMessage",
            json={"chat_id": chat_id, "text": text},
            timeout=10
        )
        if resp.status_code == 200 and resp.json().get('ok'):
            text_status = 'sent'
            log.info("✅ Telegram message sent!")
        else:
            log.error(f"❌ Telegram message failed: {resp.text[:200]}")
            text_status = f'error: {resp.status_code}'
    except Exception as e:
        log.error(f"❌ Telegram message error: {e}")
        text_status = f'error: {str(e)[:50]}'

    if has_ppt:
        try:
            filename = (
                f"OnePager_OCC{occ}_{data.get('station','ST')}_"
                f"{datetime.now().strftime('%H%M%S')}.pptx"
            )
            log.info(f"📎 Sending PPT via Telegram: {filename}...")
            resp = requests.post(
                f"{base_url}/sendDocument",
                data={
                    "chat_id": chat_id,
                    "caption": f"📊 One Pager — OCC {occ} | {data.get('failure','Open Hem')}"
                },
                files={
                    "document": (
                        filename,
                        io.BytesIO(ppt_bytes),
                        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                },
                timeout=20
            )
            if resp.status_code == 200 and resp.json().get('ok'):
                doc_status = 'sent'
                log.info("✅ Telegram PPT document sent!")
            else:
                doc_status = f'error: {resp.status_code}'
                log.error(f"❌ Telegram document failed: {resp.text[:200]}")
        except Exception as e:
            doc_status = f'error: {str(e)[:50]}'
            log.error(f"❌ Telegram document error: {e}")

    return text_status, doc_status


# ═══════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════

@app.route("/", methods=["GET"])
def health():
    smtp = get_smtp_config()
    tg = get_telegram_config()
    return jsonify({
        "status": "ok",
        "service": "HemmTrack Pro V2 Backend",
        "smtp_configured": smtp['valid'],
        "smtp_email": (smtp['email'][:5] + '***') if smtp['email'] else 'NOT SET',
        "telegram_configured": bool(tg['token'] and tg['chat_id']),
        "timestamp": datetime.now().isoformat(),
    })


@app.route("/check_occ", methods=["POST"])
def check_occ():
    """
    Main alert endpoint.
      OCC < 5 → Email + Telegram (no PPT)
      OCC >= 5 → PPT + Email with attachment + Telegram with document
    """
    start = time.time()

    try:
        data = request.get_json(force=True)
        log.info("=" * 60)
        log.info(f"📥 /check_occ — OCC: {data.get('occ','?')}")
        log.info(f"   Station: {data.get('station')} | "
                 f"Failure: {data.get('failure')} | "
                 f"Shift: {data.get('shift')}")

        occ_val = int(data.get('occ', 0))
        smtp_cfg = get_smtp_config()
        tg_cfg = get_telegram_config()

        ppt_bytes = None
        ppt_info = None

        if occ_val >= 5:
            log.info("📊 OCC >= 5 → Generating One Pager PPT...")
            ppt_bytes = generate_one_pager_ppt(data)
            ppt_info = {
                "filename": f"OnePager_OCC{data.get('occ','00')}_{data.get('station','ST')}.pptx",
                "sizeKB": len(ppt_bytes) // 1024
            }
        else:
            log.info(f"📋 OCC={occ_val} < 5 → Alert-only (no PPT)")

        email_status = send_email_with_ppt(smtp_cfg, data, ppt_bytes)
        tg_text_status, tg_doc_status = send_telegram_alert(tg_cfg, data, ppt_bytes)

        elapsed = f"{time.time() - start:.1f}s"

        result = {
            "success": True,
            "occ": data.get('occ'),
            "email": email_status,
            "telegram": tg_text_status,
            "telegramDoc": tg_doc_status,
            "ppt": ppt_info,
            "elapsed": elapsed,
        }

        log.info(f"📤 Result: email={email_status} | "
                 f"telegram={tg_text_status} | "
                 f"telegramDoc={tg_doc_status} | "
                 f"{elapsed}")
        log.info("=" * 60)

        return jsonify(result), 200

    except Exception as e:
        elapsed = f"{time.time() - start:.1f}s"
        log.error(f"❌ /check_occ FAILED: {e}", exc_info=True)
        return jsonify({
            "success": False,
            "error": str(e),
            "elapsed": elapsed,
        }), 500


@app.route("/get_stats", methods=["GET"])
def get_stats():
    """Dashboard placeholder."""
    return jsonify({
        "total_defects": 0, "high_defects": 0, "stations": 5,
        "alerts": 0, "by_station": {}, "defect_types": {},
        "by_shift": {}, "recent": [],
        "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })


@app.route("/debug/env", methods=["GET"])
def debug_env():
    """Debug — check if env vars are loaded (safe/masked)."""
    smtp = get_smtp_config()
    return jsonify({
        "SMTP_EMAIL": (smtp['email'][:8] + '***') if smtp['email'] else 'NOT SET',
        "SMTP_PASSWORD_LENGTH": len(smtp['password']) if smtp['password'] else 0,
        "SMTP_PASSWORD_SET": bool(smtp['password']),
        "SMTP_HOST": smtp['host'],
        "SMTP_PORT": smtp['port'],
        "VALID": smtp['valid'],
        "env_keys_found": [
            k for k in os.environ.keys()
            if any(x in k.upper() for x in ('SMTP', 'TELEGRAM', 'EMAIL'))
        ],
    })


# ═══════════════════════════════════════════════════
# RUN
# ═══════════════════════════════════════════════════
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    log.info(f"🚀 HemmTrack Pro Backend starting on port {port}...")
    smtp = get_smtp_config()
    if not smtp['valid']:
        log.warning("⚠️  SMTP NOT CONFIGURED — emails will be skipped!")
        log.warning("   Set SMTP_EMAIL + SMTP_PASSWORD in Railway Variables")
    app.run(host="0.0.0.0", port=port, debug=False)
