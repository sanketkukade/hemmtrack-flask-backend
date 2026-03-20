"""
HemmTrack Flask Backend — app.py
Corrected version with CORS, proper error handling, and both OCC + PPT routes.

Deploy on Railway with these environment variables:
  SMTP_EMAIL, SMTP_PASSWORD, TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
"""

import os
import json
import time
import traceback
from datetime import datetime

from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)

# ═══════════════════════════════════════════════════
# CORS — This is the #1 reason for "Failed to fetch"
# ═══════════════════════════════════════════════════
# Allow requests from GitHub Pages AND localhost dev
CORS(app, resources={r"/*": {
    "origins": [
        "https://sanketkukade.github.io",
        "http://localhost:*",
        "http://127.0.0.1:*",
        "null"  # for file:// protocol during local dev
    ],
    "methods": ["GET", "POST", "OPTIONS"],
    "allow_headers": ["Content-Type", "Authorization"],
    "max_age": 3600
}})


# ═══════════════════════════════════════════════════
# Health check — Railway uses this to verify the app is alive
# ═══════════════════════════════════════════════════
@app.route('/', methods=['GET'])
def health():
    return jsonify({
        "status": "ok",
        "service": "HemmTrack Flask Backend",
        "version": "2.1",
        "timestamp": datetime.utcnow().isoformat(),
        "endpoints": ["/check_occ", "/get_stats"]
    })


# ═══════════════════════════════════════════════════
# GET /get_stats — Live dashboard data endpoint
# ═══════════════════════════════════════════════════
@app.route('/get_stats', methods=['GET'])
def get_stats():
    """Returns current defect stats for the live dashboard."""
    try:
        # In production, this would query a database.
        # For now, return structured data that the frontend expects.
        return jsonify({
            "success": True,
            "total_defects": 0,
            "high_defects": 0,
            "fpy": 98.5,
            "open_actions": 0,
            "cpk": 1.33,
            "stations": {
                "ST-100": 0, "ST-120": 0,
                "ST-150": 0, "ST-200": 0, "ST-220": 0
            },
            "defect_types": {
                "Open Hem": 0, "Wrinkle": 0, "Overlap": 0,
                "Crack": 0, "Dent": 0, "Gap Issue": 0
            },
            "by_shift": {
                "Shift A": 0, "Shift B": 0, "Shift C": 0
            },
            "recent_defects": []
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


# ═══════════════════════════════════════════════════
# POST /check_occ — Handles BOTH OCC=3 and PPT=5 alerts
# ═══════════════════════════════════════════════════
@app.route('/check_occ', methods=['POST', 'OPTIONS'])
def check_occ():
    """
    Receives defect alert payload from frontend.
    - OCC=3: Sends email + Telegram text alert
    - PPT=5: Generates One-Pager PPT + sends email with attachment + Telegram
    """
    # Handle preflight CORS
    if request.method == 'OPTIONS':
        return '', 204

    start_time = time.time()

    try:
        data = request.get_json(force=True)
        if not data:
            return jsonify({"success": False, "error": "Empty payload"}), 400

        occ_count = int(data.get('occ', '0'))
        model = data.get('model', 'Nexon')
        failure = data.get('failure', 'Open Hem')
        station = data.get('station', 'ST-100')
        rc = data.get('rc', 'Under Analysis')
        press = data.get('press', '180')
        raybg = data.get('raybg', 'R')
        actions = data.get('actions', 'TBD')
        ecd = data.get('ecd', 'TBD')
        date_str = data.get('date', datetime.now().strftime('%d %b %Y'))
        shift = data.get('shift', '-')
        side = data.get('side', '-')
        inspector = data.get('inspector', '-')
        high_history = data.get('highHistory', [])

        app.logger.info(f"📥 Received alert: OCC={occ_count}, Station={station}, "
                        f"Defect={failure}, Model={model}")

        result = {
            "success": True,
            "occ": occ_count,
            "email": "pending",
            "telegram": "pending",
            "telegramDoc": "skipped",
            "ppt": None
        }

        # ── Build alert message ──
        alert_subject = f"🚨 HemmTrack Alert — OCC={occ_count:02d} | {failure} | {station}"
        alert_body = (
            f"HemmTrack Pro V2 — Defect Alert\n"
            f"{'='*45}\n"
            f"Date: {date_str}\n"
            f"Model: {model}\n"
            f"Station: {station} | Side: {side}\n"
            f"Defect: {failure} | Severity: HIGH\n"
            f"Root Cause: {rc}\n"
            f"Pressure: {press} bar\n"
            f"Inspector: {inspector} | Shift: {shift}\n"
            f"OCC Count: {occ_count}\n"
            f"RAYBG: {raybg}\n"
            f"Actions: {actions}\n"
            f"ECD: {ecd}\n"
        )

        if high_history:
            alert_body += f"\n{'='*45}\nRecent HIGH Defects:\n"
            for h in high_history[-5:]:
                alert_body += (f"  • {h.get('date','-')} | {h.get('station','-')} | "
                               f"{h.get('defect','-')} | {h.get('rootCause','-')}\n")

        # ── Send Email via SMTP ──
        try:
            result["email"] = _send_email(alert_subject, alert_body)
        except Exception as e:
            app.logger.error(f"Email failed: {e}")
            result["email"] = f"failed: {str(e)}"

        # ── Send Telegram ──
        try:
            result["telegram"] = _send_telegram(alert_body)
        except Exception as e:
            app.logger.error(f"Telegram failed: {e}")
            result["telegram"] = f"failed: {str(e)}"

        # ── PPT=5: Generate One-Pager PPT ──
        if occ_count >= 5:
            try:
                ppt_result = _generate_and_send_ppt(data, alert_subject)
                result["ppt"] = ppt_result
                result["telegramDoc"] = ppt_result.get("telegramDoc", "skipped")
            except Exception as e:
                app.logger.error(f"PPT generation failed: {e}")
                result["ppt"] = {"error": str(e)}

        elapsed = f"{(time.time() - start_time):.2f}s"
        result["elapsed"] = elapsed
        app.logger.info(f"✅ Alert processed in {elapsed}")

        return jsonify(result)

    except Exception as e:
        app.logger.error(f"❌ check_occ error: {traceback.format_exc()}")
        return jsonify({
            "success": False,
            "error": str(e),
            "elapsed": f"{(time.time() - start_time):.2f}s"
        }), 500


# ═══════════════════════════════════════════════════
# Helper: Send Email via SMTP
# ═══════════════════════════════════════════════════
def _send_email(subject, body, attachment_path=None):
    """Send email using Gmail SMTP."""
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email import encoders

    smtp_email = os.environ.get('SMTP_EMAIL')
    smtp_password = os.environ.get('SMTP_PASSWORD')

    if not smtp_email or not smtp_password:
        app.logger.warning("SMTP credentials not configured")
        return "skipped: no SMTP credentials"

    recipient = smtp_email  # Send to self

    msg = MIMEMultipart()
    msg['From'] = smtp_email
    msg['To'] = recipient
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Attach PPT if provided
    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',
                            f'attachment; filename={os.path.basename(attachment_path)}')
            msg.attach(part)

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(smtp_email, smtp_password)
            server.send_message(msg)
        return "sent"
    except smtplib.SMTPAuthenticationError:
        return "failed: SMTP auth error — check App Password"
    except Exception as e:
        return f"failed: {str(e)}"


# ═══════════════════════════════════════════════════
# Helper: Send Telegram Message
# ═══════════════════════════════════════════════════
def _send_telegram(text, document_path=None):
    """Send message (and optionally a document) via Telegram Bot API."""
    import requests

    bot_token = os.environ.get('TELEGRAM_BOT_TOKEN')
    chat_id = os.environ.get('TELEGRAM_CHAT_ID')

    if not bot_token or not chat_id:
        app.logger.warning("Telegram credentials not configured")
        return "skipped: no Telegram credentials"

    try:
        # Send text message
        url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
        resp = requests.post(url, json={
            "chat_id": chat_id,
            "text": text[:4096],  # Telegram limit
            "parse_mode": "HTML"
        }, timeout=10)

        if resp.status_code != 200:
            return f"failed: HTTP {resp.status_code}"

        # Send document if provided
        if document_path and os.path.exists(document_path):
            doc_url = f"https://api.telegram.org/bot{bot_token}/sendDocument"
            with open(document_path, 'rb') as f:
                requests.post(doc_url, data={"chat_id": chat_id},
                              files={"document": f}, timeout=30)

        return "sent"
    except requests.exceptions.Timeout:
        return "failed: timeout"
    except Exception as e:
        return f"failed: {str(e)}"


# ═══════════════════════════════════════════════════
# Helper: Generate One-Pager PPT (PPT=5)
# ═══════════════════════════════════════════════════
def _generate_and_send_ppt(data, subject):
    """Generate a simple One-Pager PPT and send via email + Telegram."""
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        import tempfile

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(0.8))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = f"ONE PAGER — {data.get('failure', 'Open Hem')} | {data.get('station', 'ST-100')}"
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

        # Details
        details = [
            f"Date: {data.get('date', '-')}",
            f"Model: {data.get('model', 'Nexon')}",
            f"Station: {data.get('station', '-')} | Side: {data.get('side', '-')}",
            f"Defect: {data.get('failure', '-')} | OCC: {data.get('occ', '-')}",
            f"Root Cause: {data.get('rc', '-')}",
            f"Pressure: {data.get('press', '-')} bar",
            f"RAYBG Status: {data.get('raybg', '-')}",
            f"Actions: {data.get('actions', '-')}",
            f"ECD: {data.get('ecd', '-')}",
            f"Inspector: {data.get('inspector', '-')} | Shift: {data.get('shift', '-')}",
        ]

        detail_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(12), Inches(5))
        tf2 = detail_box.text_frame
        tf2.word_wrap = True
        for line in details:
            p = tf2.add_paragraph()
            p.text = line
            p.font.size = Pt(16)
            p.space_after = Pt(6)

        # Save to temp file
        tmp = tempfile.NamedTemporaryFile(suffix='.pptx', delete=False,
                                           prefix='HemmTrack_OnePager_')
        prs.save(tmp.name)
        file_size = os.path.getsize(tmp.name)

        # Send via email + Telegram
        email_result = _send_email(subject + " [ONE PAGER]", 
                                    "One-Pager PPT attached.", tmp.name)
        telegram_doc = _send_telegram(
            f"📊 ONE PAGER GENERATED\n{data.get('failure','-')} | {data.get('station','-')}",
            tmp.name)

        # Cleanup
        try:
            os.unlink(tmp.name)
        except Exception:
            pass

        return {
            "filename": os.path.basename(tmp.name),
            "sizeKB": round(file_size / 1024, 1),
            "email": email_result,
            "telegramDoc": telegram_doc
        }

    except ImportError:
        app.logger.warning("python-pptx not installed, skipping PPT generation")
        return {"error": "python-pptx not installed", "email": "skipped", "telegramDoc": "skipped"}
    except Exception as e:
        return {"error": str(e), "email": "skipped", "telegramDoc": "skipped"}


# ═══════════════════════════════════════════════════
# Run
# ═══════════════════════════════════════════════════
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
