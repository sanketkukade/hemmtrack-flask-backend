"""
HemmTrack Pro V2 — Flask Backend
=================================
Endpoint: POST /check_occ
  → Receives defect alert data from frontend
  → Generates One-Pager PPT (when OCC >= 5)
  → Sends email with PPT attachment via Resend API
  → Sends Telegram text alert + PPT document

Railway Environment Variables Required:
  RESEND_API_KEY    = re_xxxxxxxx   (from https://resend.com/api-keys)
  RESEND_FROM_EMAIL = onboarding@resend.dev  (free tier sender)
  ALERT_TO_EMAIL    = sanketkukade111@gmail.com

  TELEGRAM_BOT_TOKEN = (has default from frontend)
  TELEGRAM_CHAT_ID   = (has default from frontend)
"""

import os
import io
import time
import json
import base64
import socket
import logging
import requests
import requests.adapters
import urllib3.util.retry
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
    "null",
])


# ═══════════════════════════════════════════════════
# RESILIENT HTTP — IPv4 Forcing + Retry Adapter
# ═══════════════════════════════════════════════════
# Railway containers sometimes fail on IPv6 DNS resolution.
# Force all outbound connections to use IPv4.
_orig_getaddrinfo = socket.getaddrinfo

def _ipv4_getaddrinfo(host, port, family=0, type=0, proto=0, flags=0):
    """Force IPv4 (AF_INET) for all DNS lookups."""
    return _orig_getaddrinfo(host, port, socket.AF_INET, type, proto, flags)

socket.getaddrinfo = _ipv4_getaddrinfo
log.info("🌐 IPv4 forced for all outbound connections")


def _build_resilient_session():
    """Build a requests.Session with automatic retry on transient failures.

    Retry strategy:
      - 3 total retries (so up to 4 attempts total)
      - Exponential backoff: 1s → 2s → 4s between retries
      - Retries on: 429 (rate limit), 500, 502, 503, 504
      - Retries on: ConnectionError, Timeout, incomplete reads
    """
    retry = urllib3.util.retry.Retry(
        total=3,
        backoff_factor=1,          # 1s, 2s, 4s
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["POST", "GET"],
        raise_on_status=False,     # Let us handle status codes
    )
    adapter = requests.adapters.HTTPAdapter(
        max_retries=retry,
        pool_connections=5,
        pool_maxsize=5,
    )
    session = requests.Session()
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


# Single global session — reused across all requests
http = _build_resilient_session()
log.info("🔄 Resilient HTTP session ready (3 retries, exponential backoff)")


# ═══════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════
RESEND_API_URL = "https://api.resend.com/emails"


def get_resend_config():
    """Read Resend email config from Railway environment variables."""
    api_key = os.environ.get('RESEND_API_KEY', '').strip()
    from_email = os.environ.get('RESEND_FROM_EMAIL', 'onboarding@resend.dev').strip()
    to_email = os.environ.get('ALERT_TO_EMAIL', '').strip()

    log.info("📧 Resend Config:")
    log.info(f"   RESEND_API_KEY    = {'✅ ' + api_key[:10] + '...' if api_key else '❌ NOT SET'}")
    log.info(f"   RESEND_FROM_EMAIL = {from_email}")
    log.info(f"   ALERT_TO_EMAIL    = {'✅ ' + to_email if to_email else '❌ NOT SET'}")

    return {
        'api_key': api_key,
        'from_email': from_email,
        'to_email': to_email,
        'valid': bool(api_key and to_email),
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
# PPT GENERATION  (unchanged from original)
# ═══════════════════════════════════════════════════
def generate_one_pager_ppt(data):
    """Generate a One-Pager PPTX and return as bytes."""
    log.info("📊 Generating One-Pager PPT...")

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0x06, 0x0A, 0x12)

    title_box = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(12.7), Inches(0.7))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🚨 HEMMING DEFECT — ONE PAGER ALERT"
    p.font.size = Pt(28); p.font.bold = True
    p.font.color.rgb = RGBColor(0x00, 0xE5, 0xFF)
    p.alignment = PP_ALIGN.CENTER

    sub_box = slide.shapes.add_textbox(Inches(0.3), Inches(0.9), Inches(12.7), Inches(0.4))
    tf = sub_box.text_frame
    p = tf.paragraphs[0]
    p.text = (f"Date: {data.get('date', '-')}  |  Station: {data.get('station', '-')}  |  "
              f"Model: {data.get('model', '-')}  |  Shift: {data.get('shift', '-')}")
    p.font.size = Pt(14); p.font.color.rgb = RGBColor(0x4A, 0x6F, 0xA5); p.alignment = PP_ALIGN.CENTER

    info_items = [
        ("OCC Count", data.get('occ', '-'), "🔴"), ("Failure Mode", data.get('failure', '-'), "🔍"),
        ("Root Cause", data.get('rc', '-'), "🔧"), ("Hemming Pressure", f"{data.get('press', '-')} Bar", "💨"),
        ("Side", data.get('side', '-'), "🚪"), ("Inspector ID", data.get('inspector', '-'), "👤"),
        ("RAYBG Status", data.get('raybg', '-'), "📋"), ("Corrective Action", data.get('actions', '-'), "⚡"),
        ("ECD (Target)", data.get('ecd', '-'), "📅"),
    ]

    for i, (label, value, icon) in enumerate(info_items):
        col, row = i % 3, i // 3
        x, y = 0.5 + col * 4.1, 1.5 + row * 0.7
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(3.8), Inches(0.6))
        tf = box.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = f"{icon} {label}"
        p.font.size = Pt(10); p.font.color.rgb = RGBColor(0x4A, 0x6F, 0xA5); p.font.bold = True
        p2 = tf.add_paragraph(); p2.text = str(value); p2.font.size = Pt(16); p2.font.bold = True
        p2.font.color.rgb = RGBColor(0xFF, 0x3D, 0x5A) if label == "OCC Count" else RGBColor(0xD4, 0xE8, 0xFF)

    history = data.get('highHistory', [])
    if history:
        hist_title = slide.shapes.add_textbox(Inches(0.5), Inches(3.6), Inches(8), Inches(0.35))
        tf = hist_title.text_frame; p = tf.paragraphs[0]
        p.text = "📋 Recent HIGH Defect History (Last 5)"
        p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = RGBColor(0xFF, 0xAA, 0x00)

        rows = min(len(history), 5) + 1
        table_shape = slide.shapes.add_table(rows, 5, Inches(0.5), Inches(4.0), Inches(12.3), Inches(0.35 * rows))
        table = table_shape.table
        for j, h in enumerate(['Date', 'Station', 'Defect', 'Root Cause', 'Pressure']):
            cell = table.cell(0, j); cell.text = h
            for par in cell.text_frame.paragraphs:
                par.font.size = Pt(9); par.font.bold = True; par.font.color.rgb = RGBColor(0x00, 0xE5, 0xFF)
        for r, entry in enumerate(history[:5]):
            for j, v in enumerate([entry.get('date','-'), entry.get('station','-'), entry.get('defect','-'),
                                   entry.get('rootCause','-'), str(entry.get('pressure','-'))]):
                cell = table.cell(r+1, j); cell.text = v
                for par in cell.text_frame.paragraphs:
                    par.font.size = Pt(9); par.font.color.rgb = RGBColor(0xD4, 0xE8, 0xFF)

    footer_box = slide.shapes.add_textbox(Inches(0.3), Inches(6.8), Inches(12.7), Inches(0.5))
    tf = footer_box.text_frame; p = tf.paragraphs[0]
    p.text = "Generated by HemmTrack Pro V2  |  Sanket Kukade — M.Tech Manufacturing Engg, DYPIU Pune  |  Tata Motors PVBU"
    p.font.size = Pt(10); p.font.color.rgb = RGBColor(0x4A, 0x6F, 0xA5); p.alignment = PP_ALIGN.CENTER

    buffer = io.BytesIO()
    prs.save(buffer); buffer.seek(0)
    ppt_bytes = buffer.getvalue()
    log.info(f"📊 PPT generated: {len(ppt_bytes)} bytes ({len(ppt_bytes)//1024} KB)")
    return ppt_bytes


# ═══════════════════════════════════════════════════
# EMAIL — Resend API + Retry + IPv4 + Timeout
# ═══════════════════════════════════════════════════
MAX_EMAIL_ATTEMPTS = 3
EMAIL_TIMEOUT = (10, 30)  # (connect_timeout, read_timeout)


def send_email_with_ppt(resend_cfg, data, ppt_bytes=None):
    """Send alert email via Resend API with retry resilience.

    Resilience features:
      ✅ IPv4 forced (socket monkey-patch above)
      ✅ 3 retry attempts with exponential backoff (1s → 2s → 4s)
      ✅ Separate connect (10s) and read (30s) timeouts
      ✅ Per-attempt logging with timing
      ✅ urllib3 Retry adapter handles 429/5xx automatically
      ✅ Manual retry for application-level errors
    """

    if not resend_cfg['valid']:
        log.error("❌ Resend not configured! Railway → Variables → Add:")
        log.error("   RESEND_API_KEY   = re_xxxxxxxx")
        log.error("   ALERT_TO_EMAIL   = sanketkukade111@gmail.com")
        return 'skipped: Resend not configured'

    occ = data.get('occ', '?')
    station = data.get('station', '-')
    failure = data.get('failure', 'Open Hem')
    has_ppt = ppt_bytes is not None

    subject = (f"📊 ONE PAGER — OCC {occ} | {failure} | {station} — HemmTrack" if has_ppt
               else f"🚨 OCC ALERT {occ} — {failure} | {station} — HemmTrack")

    html_body = f"""
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;
                background:#060a12;color:#d4e8ff;padding:24px;border-radius:12px;">
      <div style="text-align:center;padding:16px;background:linear-gradient(135deg,#0a1628,#0f1928);
                  border-radius:10px;margin-bottom:16px;">
        <h1 style="color:#00e5ff;margin:0;font-size:22px;">
          {'📊 ONE PAGER ALERT' if has_ppt else '🚨 OCC ALERT'}</h1>
        <p style="color:#4a6fa5;margin:4px 0 0;font-size:12px;">HemmTrack Pro V2 — Automated Alert</p>
      </div>
      <table style="width:100%;border-collapse:collapse;margin:16px 0;">
        <tr><td style="padding:8px;color:#4a6fa5;">📅 Date</td><td style="padding:8px;color:#d4e8ff;font-weight:700;">{data.get('date','-')}</td></tr>
        <tr><td style="padding:8px;color:#4a6fa5;">🔴 OCC Count</td><td style="padding:8px;color:#ff3d5a;font-weight:900;font-size:18px;">{occ}</td></tr>
        <tr><td style="padding:8px;color:#4a6fa5;">🏭 Station</td><td style="padding:8px;color:#d4e8ff;font-weight:700;">{station}</td></tr>
        <tr><td style="padding:8px;color:#4a6fa5;">🔍 Failure</td><td style="padding:8px;color:#ff3d5a;font-weight:700;">{failure}</td></tr>
        <tr><td style="padding:8px;color:#4a6fa5;">🚪 Side</td><td style="padding:8px;color:#d4e8ff;">{data.get('side','-')}</td></tr>
        <tr><td style="padding:8px;color:#4a6fa5;">🔧 Root Cause</td><td style="padding:8px;color:#d4e8ff;">{data.get('rc','-')}</td></tr>
        <tr><td style="padding:8px;color:#4a6fa5;">💨 Pressure</td><td style="padding:8px;color:#d4e8ff;">{data.get('press','-')} Bar</td></tr>
        <tr><td style="padding:8px;color:#4a6fa5;">👤 Inspector</td><td style="padding:8px;color:#d4e8ff;">{data.get('inspector','-')}</td></tr>
        <tr><td style="padding:8px;color:#4a6fa5;">⚡ Action</td><td style="padding:8px;color:#d4e8ff;">{data.get('actions','-')}</td></tr>
      </table>
      {'<p style="text-align:center;color:#3e86f6;font-weight:700;">📎 One Pager PPT attached</p>' if has_ppt else ''}
      <div style="text-align:center;padding:12px;margin-top:16px;border-top:1px solid #1e3050;">
        <p style="color:#4a6fa5;font-size:10px;margin:0;">Sanket Kukade — M.Tech Manufacturing Engg, DYPIU Pune<br>Tata Motors PVBU — HemmTrack Pro V2</p>
      </div>
    </div>"""

    payload = {
        "from": resend_cfg['from_email'],
        "to": [resend_cfg['to_email']],
        "subject": subject,
        "html": html_body,
    }

    if has_ppt:
        filename = f"OnePager_OCC{occ}_{station}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        payload["attachments"] = [{
            "filename": filename,
            "content": base64.b64encode(ppt_bytes).decode('utf-8'),
        }]
        log.info(f"📎 PPT attached: {filename} ({len(ppt_bytes)//1024} KB)")

    # ── Retry Loop ──
    last_error = None
    for attempt in range(1, MAX_EMAIL_ATTEMPTS + 1):
        attempt_start = time.time()
        try:
            log.info(f"📤 Resend attempt {attempt}/{MAX_EMAIL_ATTEMPTS} → {resend_cfg['to_email']}...")

            resp = http.post(
                RESEND_API_URL,
                headers={
                    "Authorization": f"Bearer {resend_cfg['api_key']}",
                    "Content-Type": "application/json",
                },
                json=payload,
                timeout=EMAIL_TIMEOUT,
            )

            elapsed_ms = int((time.time() - attempt_start) * 1000)

            if resp.status_code == 200:
                resp_json = resp.json()
                if resp_json.get('id'):
                    log.info(f"✅ Email SENT via Resend! ID: {resp_json['id']} ({elapsed_ms}ms, attempt {attempt})")
                    return 'sent'

            # Non-200: log and decide whether to retry
            err_msg = resp.text[:200]
            log.warning(f"⚠️  Attempt {attempt} failed ({resp.status_code}, {elapsed_ms}ms): {err_msg}")

            # Don't retry on client errors (4xx) — they won't change
            if 400 <= resp.status_code < 500 and resp.status_code != 429:
                log.error(f"❌ Resend client error {resp.status_code} — not retrying")
                if resp.status_code in (401, 403):
                    log.error("   → RESEND_API_KEY invalid. Get new key: https://resend.com/api-keys")
                if resp.status_code == 422:
                    log.error("   → Validation error. Check RESEND_FROM_EMAIL is verified.")
                return f'error: {resp.status_code}'

            last_error = f'http_{resp.status_code}'

        except requests.exceptions.ConnectTimeout:
            elapsed_ms = int((time.time() - attempt_start) * 1000)
            log.warning(f"⚠️  Attempt {attempt} connect timeout ({elapsed_ms}ms)")
            last_error = 'connect_timeout'

        except requests.exceptions.ReadTimeout:
            elapsed_ms = int((time.time() - attempt_start) * 1000)
            log.warning(f"⚠️  Attempt {attempt} read timeout ({elapsed_ms}ms)")
            last_error = 'read_timeout'

        except requests.exceptions.ConnectionError as e:
            elapsed_ms = int((time.time() - attempt_start) * 1000)
            log.warning(f"⚠️  Attempt {attempt} connection error ({elapsed_ms}ms): {str(e)[:100]}")
            last_error = 'connection_error'

        except Exception as e:
            elapsed_ms = int((time.time() - attempt_start) * 1000)
            log.warning(f"⚠️  Attempt {attempt} unexpected error ({elapsed_ms}ms): {e}")
            last_error = str(e)[:60]

        # Backoff before next attempt (skip if last attempt)
        if attempt < MAX_EMAIL_ATTEMPTS:
            wait = 2 ** (attempt - 1)  # 1s, 2s
            log.info(f"   ⏳ Waiting {wait}s before retry...")
            time.sleep(wait)

    # All attempts exhausted
    log.error(f"❌ Email FAILED after {MAX_EMAIL_ATTEMPTS} attempts. Last error: {last_error}")
    return f'failed: {last_error} (after {MAX_EMAIL_ATTEMPTS} attempts)'


# ═══════════════════════════════════════════════════
# TELEGRAM  (completely unchanged from original)
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
        resp = requests.post(f"{base_url}/sendMessage", json={"chat_id": chat_id, "text": text}, timeout=10)
        if resp.status_code == 200 and resp.json().get('ok'):
            text_status = 'sent'; log.info("✅ Telegram message sent!")
        else:
            log.error(f"❌ Telegram message failed: {resp.text[:200]}"); text_status = f'error: {resp.status_code}'
    except Exception as e:
        log.error(f"❌ Telegram message error: {e}"); text_status = f'error: {str(e)[:50]}'

    if has_ppt:
        try:
            filename = f"OnePager_OCC{occ}_{data.get('station','ST')}_{datetime.now().strftime('%H%M%S')}.pptx"
            log.info(f"📎 Sending PPT via Telegram: {filename}...")
            resp = requests.post(
                f"{base_url}/sendDocument",
                data={"chat_id": chat_id, "caption": f"📊 One Pager — OCC {occ} | {data.get('failure','Open Hem')}"},
                files={"document": (filename, io.BytesIO(ppt_bytes),
                       "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
                timeout=20
            )
            if resp.status_code == 200 and resp.json().get('ok'):
                doc_status = 'sent'; log.info("✅ Telegram PPT document sent!")
            else:
                doc_status = f'error: {resp.status_code}'; log.error(f"❌ Telegram document failed: {resp.text[:200]}")
        except Exception as e:
            doc_status = f'error: {str(e)[:50]}'; log.error(f"❌ Telegram document error: {e}")

    return text_status, doc_status


# ═══════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════

@app.route("/", methods=["GET"])
def health():
    resend = get_resend_config()
    tg = get_telegram_config()
    return jsonify({
        "status": "ok",
        "service": "HemmTrack Pro V2 Backend",
        "email_provider": "Resend API",
        "email_configured": resend['valid'],
        "from_email": resend['from_email'],
        "to_email": (resend['to_email'][:5] + '***') if resend['to_email'] else 'NOT SET',
        "telegram_configured": bool(tg['token'] and tg['chat_id']),
        "timestamp": datetime.now().isoformat(),
    })


@app.route("/check_occ", methods=["POST"])
def check_occ():
    """OCC < 5 → Email + Telegram (no PPT) | OCC >= 5 → PPT + Email + Telegram"""
    start = time.time()
    try:
        data = request.get_json(force=True)
        log.info("=" * 60)
        log.info(f"📥 /check_occ — OCC: {data.get('occ','?')}")
        log.info(f"   Station: {data.get('station')} | Failure: {data.get('failure')} | Shift: {data.get('shift')}")

        occ_val = int(data.get('occ', 0))
        resend_cfg = get_resend_config()
        tg_cfg = get_telegram_config()

        ppt_bytes = None; ppt_info = None
        if occ_val >= 5:
            log.info("📊 OCC >= 5 → Generating One Pager PPT...")
            ppt_bytes = generate_one_pager_ppt(data)
            ppt_info = {"filename": f"OnePager_OCC{data.get('occ','00')}_{data.get('station','ST')}.pptx",
                        "sizeKB": len(ppt_bytes) // 1024}
        else:
            log.info(f"📋 OCC={occ_val} < 5 → Alert-only (no PPT)")

        email_status = send_email_with_ppt(resend_cfg, data, ppt_bytes)
        tg_text_status, tg_doc_status = send_telegram_alert(tg_cfg, data, ppt_bytes)

        elapsed = f"{time.time() - start:.1f}s"
        result = {"success": True, "occ": data.get('occ'), "email": email_status,
                  "telegram": tg_text_status, "telegramDoc": tg_doc_status, "ppt": ppt_info, "elapsed": elapsed}
        log.info(f"📤 Result: email={email_status} | telegram={tg_text_status} | telegramDoc={tg_doc_status} | {elapsed}")
        log.info("=" * 60)
        return jsonify(result), 200

    except Exception as e:
        elapsed = f"{time.time() - start:.1f}s"
        log.error(f"❌ /check_occ FAILED: {e}", exc_info=True)
        return jsonify({"success": False, "error": str(e), "elapsed": elapsed}), 500


@app.route("/get_stats", methods=["GET"])
def get_stats():
    return jsonify({"total_defects": 0, "high_defects": 0, "stations": 5, "alerts": 0,
                    "by_station": {}, "defect_types": {}, "by_shift": {}, "recent": [],
                    "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})


@app.route("/debug/env", methods=["GET"])
def debug_env():
    resend = get_resend_config()
    return jsonify({
        "provider": "Resend API",
        "RESEND_API_KEY": ('✅ ' + resend['api_key'][:10] + '...') if resend['api_key'] else '❌ NOT SET',
        "RESEND_FROM_EMAIL": resend['from_email'],
        "ALERT_TO_EMAIL": resend['to_email'] or '❌ NOT SET',
        "READY": resend['valid'],
        "env_keys": [k for k in os.environ if any(x in k.upper() for x in ('RESEND','TELEGRAM','EMAIL','ALERT'))],
    })


# ═══════════════════════════════════════════════════
# RUN
# ═══════════════════════════════════════════════════
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    log.info(f"🚀 HemmTrack Pro Backend starting on port {port}...")
    r = get_resend_config()
    if not r['valid']:
        log.warning("⚠️  RESEND NOT CONFIGURED — emails will fail!")
        log.warning("   Set RESEND_API_KEY + ALERT_TO_EMAIL in Railway")
    else:
        log.info("✅ Resend API ready")
    app.run(host="0.0.0.0", port=port, debug=False)
