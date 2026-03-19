"""
HemmTrack Pro — Flask Backend v3
=================================
Matches frontend payload from hemmtrack_final_v3.html EXACTLY.

Frontend sends this payload to /check_occ:
{
  "occ": "03" or "05",
  "model": "Altroz / Nexon",
  "failure": "Open Hem",
  "station": "ST-200",
  "rc": "Under Analysis",
  "press": "180",
  "raybg": "R",
  "actions": "TBD",
  "ecd": "TBD",
  "date": "19 Mar 2026",
  "shift": "A",
  "side": "LH",
  "inspector": "Rajesh",
  "highHistory": [{date, station, defect, rootCause, pressure}, ...]
}

Frontend expects this response:
  OCC=3 → { "success": true, "elapsed": "1.2s" }
  OCC=5 → { "success": true, "ppt": { "sizeKB": "85" }, "elapsed": "2.4s" }
  Error → { "success": false, "error": "message" }

Environment Variables (Railway):
  RESEND_API_KEY, RESEND_FROM_EMAIL, ALERT_TO_EMAIL
  TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
"""

import os
import io
import json
import time
import base64
import logging
import traceback
import threading
from datetime import datetime, timezone

import requests
from flask import Flask, jsonify, request
from flask_cors import CORS

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ──────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("hemmtrack")

app = Flask(__name__)
CORS(app)  # Frontend is on GitHub Pages — needs CORS


# ══════════════════════════════════════════════════════════════
#  ENV HELPER
# ══════════════════════════════════════════════════════════════
def get_env(key):
    return os.environ.get(key, "").strip()

def env_status(key):
    val = get_env(key)
    return f"SET ({len(val)} chars)" if val else "EMPTY"


# ══════════════════════════════════════════════════════════════
#  RESEND EMAIL (with optional attachment)
# ══════════════════════════════════════════════════════════════
def send_email(subject, body_html, attachment_bytes=None, attachment_filename=None):
    """Send email via Resend HTTP API. Returns dict with status."""
    api_key = get_env("RESEND_API_KEY")
    from_email = get_env("RESEND_FROM_EMAIL")
    to_email = get_env("ALERT_TO_EMAIL")

    if not api_key or not from_email or not to_email:
        missing = []
        if not api_key: missing.append("RESEND_API_KEY")
        if not from_email: missing.append("RESEND_FROM_EMAIL")
        if not to_email: missing.append("ALERT_TO_EMAIL")
        return {"ok": False, "error": f"Missing: {', '.join(missing)}"}

    payload = {
        "from": from_email,
        "to": [to_email],
        "subject": subject,
        "html": body_html,
    }

    if attachment_bytes and attachment_filename:
        encoded = base64.b64encode(attachment_bytes).decode("utf-8")
        payload["attachments"] = [{"filename": attachment_filename, "content": encoded}]

    try:
        resp = requests.post(
            "https://api.resend.com/emails",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json=payload,
            timeout=30,
        )
        data = resp.json()

        if resp.status_code == 200 and "id" in data:
            logger.info(f"Email sent: {data['id']}")
            return {"ok": True, "id": data["id"]}
        else:
            err = data.get("message", f"HTTP {resp.status_code}")
            logger.error(f"Resend error: {err}")
            return {"ok": False, "error": err}

    except Exception as e:
        logger.error(f"Email exception: {e}")
        return {"ok": False, "error": str(e)}


# ══════════════════════════════════════════════════════════════
#  TELEGRAM (message + document)
# ══════════════════════════════════════════════════════════════
def send_telegram_message(text):
    token = get_env("TELEGRAM_BOT_TOKEN")
    chat_id = get_env("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        return {"ok": False, "error": "Telegram credentials missing"}

    try:
        resp = requests.post(
            f"https://api.telegram.org/bot{token}/sendMessage",
            json={"chat_id": chat_id, "text": text, "parse_mode": "HTML"},
            timeout=15,
        )
        data = resp.json()
        if data.get("ok"):
            return {"ok": True, "message_id": data["result"]["message_id"]}
        return {"ok": False, "error": data.get("description", "Unknown")}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def send_telegram_document(file_bytes, filename, caption=""):
    token = get_env("TELEGRAM_BOT_TOKEN")
    chat_id = get_env("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        return {"ok": False, "error": "Telegram credentials missing"}

    try:
        resp = requests.post(
            f"https://api.telegram.org/bot{token}/sendDocument",
            data={"chat_id": chat_id, "caption": caption, "parse_mode": "HTML"},
            files={"document": (filename, io.BytesIO(file_bytes),
                   "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
            timeout=30,
        )
        data = resp.json()
        if data.get("ok"):
            return {"ok": True, "message_id": data["result"]["message_id"]}
        return {"ok": False, "error": data.get("description", "Unknown")}
    except Exception as e:
        return {"ok": False, "error": str(e)}


# ══════════════════════════════════════════════════════════════
#  PPT GENERATION — One-Page Critical Defect Report
# ══════════════════════════════════════════════════════════════
def generate_defect_ppt(p):
    """
    Generate one-page PPT from frontend payload.
    p = payload dict with: occ, model, failure, station, rc, press,
        shift, side, inspector, date, actions, ecd, highHistory
    """
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Colors
    BG      = RGBColor(0x1A, 0x1A, 0x2E)
    RED     = RGBColor(0xE7, 0x4C, 0x3C)
    WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
    LIGHT   = RGBColor(0xEC, 0xF0, 0xF1)
    BLUE    = RGBColor(0x3E, 0x86, 0xF6)
    CARD    = RGBColor(0x24, 0x24, 0x3E)
    GREEN   = RGBColor(0x27, 0xAE, 0x60)
    ORANGE  = RGBColor(0xF3, 0x9C, 0x12)
    DIMTEXT = RGBColor(0x66, 0x66, 0x88)

    # Background
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = BG

    occ       = p.get("occ", "05")
    model     = p.get("model", "Altroz / Nexon")
    failure   = p.get("failure", "Open Hem")
    station   = p.get("station", "N/A")
    rc        = p.get("rc", "Under Analysis")
    press     = p.get("press", "N/A")
    shift     = p.get("shift", "N/A")
    side      = p.get("side", "N/A")
    inspector = p.get("inspector", "N/A")
    date      = p.get("date", datetime.now().strftime("%d %b %Y"))
    actions   = p.get("actions", "TBD")
    ecd       = p.get("ecd", "TBD")
    history   = p.get("highHistory", [])

    # ═══ 1. RED ALERT BANNER ═══
    banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.7))
    banner.fill.solid()
    banner.fill.fore_color.rgb = RED
    banner.line.fill.background()
    tf = banner.text_frame
    tf.word_wrap = True
    pr = tf.paragraphs[0]
    pr.alignment = PP_ALIGN.CENTER
    pr.space_before = Pt(8)
    run = pr.add_run()
    run.text = f"CRITICAL DEFECT ALERT — OCC {occ} — {failure.upper()}"
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.color.rgb = WHITE

    # ═══ 2. TITLE + SUBTITLE ═══
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.9), Inches(9), Inches(0.7))
    tf = tb.text_frame
    run = tf.paragraphs[0].add_run()
    run.text = f"CRITICAL DEFECT ALERT: {failure} — Rear Door"
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = WHITE

    tb = slide.shapes.add_textbox(Inches(0.6), Inches(1.55), Inches(9), Inches(0.4))
    tf = tb.text_frame
    run = tf.paragraphs[0].add_run()
    run.text = f"Model: {model}  |  Date: {date}  |  Shift: {shift}"
    run.font.size = Pt(12)
    run.font.color.rgb = LIGHT

    # ═══ 3. OCC BADGE (top right) ═══
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.2), Inches(0.85), Inches(2.6), Inches(1.1))
    badge.fill.solid()
    badge.fill.fore_color.rgb = RED
    badge.line.fill.background()
    tf = badge.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = tf.paragraphs[0].add_run()
    run.text = f"OCC {occ}"
    run.font.size = Pt(44)
    run.font.bold = True
    run.font.color.rgb = WHITE
    pr2 = tf.add_paragraph()
    pr2.alignment = PP_ALIGN.CENTER
    run2 = pr2.add_run()
    run2.text = "OCCURRENCES"
    run2.font.size = Pt(10)
    run2.font.bold = True
    run2.font.color.rgb = WHITE

    # ═══ 4. INFO CARDS ROW ═══
    cards = [
        ("STATION",   station,   BLUE),
        ("SIDE",      side,      GREEN),
        ("PRESSURE",  f"{press} Bar", ORANGE),
        ("INSPECTOR", inspector, BLUE),
    ]
    cw, ch = Inches(2.9), Inches(0.95)
    sx, cy, gap = Inches(0.6), Inches(2.2), Inches(0.2)

    for i, (label, value, color) in enumerate(cards):
        x = sx + i * (cw + gap)
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, cy, cw, ch)
        card.fill.solid()
        card.fill.fore_color.rgb = CARD
        card.line.fill.background()
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.15)
        tf.margin_top = Inches(0.08)

        run = tf.paragraphs[0].add_run()
        run.text = label
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = color

        pr = tf.add_paragraph()
        run = pr.add_run()
        run.text = str(value)
        run.font.size = Pt(20)
        run.font.bold = True
        run.font.color.rgb = WHITE

    # ═══ 5. ROOT CAUSE + ACTIONS BOX ═══
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(3.4), Inches(7.8), Inches(1.8))
    box.fill.solid()
    box.fill.fore_color.rgb = CARD
    box.line.fill.background()
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    tf.margin_top = Inches(0.12)
    tf.margin_right = Inches(0.2)

    run = tf.paragraphs[0].add_run()
    run.text = "ROOT CAUSE & CORRECTIVE ACTION"
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = BLUE

    pr = tf.add_paragraph()
    pr.space_before = Pt(8)
    run = pr.add_run()
    run.text = f"Root Cause: {rc}"
    run.font.size = Pt(14)
    run.font.color.rgb = WHITE

    pr = tf.add_paragraph()
    pr.space_before = Pt(4)
    run = pr.add_run()
    run.text = f"Action Taken: {actions}"
    run.font.size = Pt(13)
    run.font.color.rgb = LIGHT

    pr = tf.add_paragraph()
    pr.space_before = Pt(4)
    run = pr.add_run()
    run.text = f"ECD (Expected Closure Date): {ecd}"
    run.font.size = Pt(13)
    run.font.color.rgb = ORANGE

    # ═══ 6. DEFECT HISTORY TABLE (right side) ═══
    hist_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.6), Inches(3.4), Inches(4.2), Inches(1.8))
    hist_box.fill.solid()
    hist_box.fill.fore_color.rgb = CARD
    hist_box.line.fill.background()
    tf = hist_box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.15)
    tf.margin_top = Inches(0.1)
    tf.margin_right = Inches(0.1)

    run = tf.paragraphs[0].add_run()
    run.text = "LAST 5 HIGH DEFECT HISTORY"
    run.font.size = Pt(10)
    run.font.bold = True
    run.font.color.rgb = ORANGE

    if history:
        for i, h in enumerate(history[:5]):
            pr = tf.add_paragraph()
            pr.space_before = Pt(2)
            run = pr.add_run()
            run.text = f"{i+1}. {h.get('date','-')} | {h.get('station','-')} | {h.get('defect','-')}"
            run.font.size = Pt(10)
            run.font.color.rgb = LIGHT
    else:
        pr = tf.add_paragraph()
        pr.space_before = Pt(6)
        run = pr.add_run()
        run.text = "No history available"
        run.font.size = Pt(11)
        run.font.color.rgb = DIMTEXT

    # ═══ 7. CAPA NOTE ═══
    note = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(5.5), Inches(12.2), Inches(0.8))
    note.fill.solid()
    note.fill.fore_color.rgb = RGBColor(0x2D, 0x15, 0x15)
    note.line.fill.background()
    tf = note.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    tf.margin_top = Inches(0.1)

    run = tf.paragraphs[0].add_run()
    run.text = f"OCC {occ} reached — As per quality protocol, initiate CAPA / 8D investigation immediately. This report has been auto-sent to management via Email & Telegram."
    run.font.size = Pt(12)
    run.font.italic = True
    run.font.color.rgb = RED

    # ═══ 8. FOOTER ═══
    ft = slide.shapes.add_textbox(Inches(0.6), Inches(6.6), Inches(12.2), Inches(0.5))
    tf = ft.text_frame
    run = tf.paragraphs[0].add_run()
    run.text = f"HemmTrack Pro V2  |  Auto-Generated Defect Report  |  Tata Motors PVBU, Pune  |  {date}"
    run.font.size = Pt(10)
    run.font.color.rgb = DIMTEXT

    # Save to bytes
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════
#  MAIN ENDPOINT — /check_occ
# ══════════════════════════════════════════════════════════════
@app.route("/", methods=["GET"])
def health():
    return jsonify({
        "service": "HemmTrack Pro Backend",
        "version": "3.0",
        "status": "running",
        "endpoints": ["/check_occ", "/debug/env", "/debug/test_email",
                      "/debug/test_telegram", "/debug/test_ppt"],
    })


@app.route("/check_occ", methods=["POST", "OPTIONS"])
def check_occ():
    """
    Main endpoint called by frontend.

    OCC=3 → Email + Telegram (text alert only)
    OCC=5 → Email with PPT attachment + Telegram with PPT document

    Response format matches what frontend expects:
      { "success": true, "ppt": {"sizeKB": "85"}, "elapsed": "2.4s" }
    """
    # Handle CORS preflight
    if request.method == "OPTIONS":
        return "", 204

    start = time.time()

    try:
        p = request.get_json(silent=True) or {}
        occ_str = p.get("occ", "00")
        occ_num = int(occ_str)

        logger.info(f"/check_occ called — OCC={occ_str}, failure={p.get('failure')}, station={p.get('station')}")

        # ── Build alert content ──
        failure  = p.get("failure", "Open Hem")
        station  = p.get("station", "N/A")
        model    = p.get("model", "Altroz / Nexon")
        shift    = p.get("shift", "N/A")
        side     = p.get("side", "N/A")
        rc       = p.get("rc", "Under Analysis")
        press    = p.get("press", "N/A")
        inspector = p.get("inspector", "N/A")
        date     = p.get("date", datetime.now().strftime("%d %b %Y"))
        actions  = p.get("actions", "TBD")
        ecd      = p.get("ecd", "TBD")

        # ── Email HTML body ──
        email_html = f"""
        <div style="font-family:Arial,sans-serif;max-width:600px;">
            <div style="background:#CC0000;color:#fff;padding:12px 20px;font-size:18px;font-weight:bold;">
                CRITICAL DEFECT ALERT: {failure} — OCC {occ_str}
            </div>
            <div style="padding:16px;background:#f9f9f9;border:1px solid #ddd;">
                <table style="width:100%;border-collapse:collapse;font-size:14px;">
                    <tr><td style="padding:6px;font-weight:bold;width:140px;">Date</td><td style="padding:6px;">{date}</td></tr>
                    <tr><td style="padding:6px;font-weight:bold;">Model</td><td style="padding:6px;">{model}</td></tr>
                    <tr><td style="padding:6px;font-weight:bold;">Shift</td><td style="padding:6px;">{shift}</td></tr>
                    <tr><td style="padding:6px;font-weight:bold;">Station</td><td style="padding:6px;">{station}</td></tr>
                    <tr><td style="padding:6px;font-weight:bold;">Side</td><td style="padding:6px;">{side}</td></tr>
                    <tr><td style="padding:6px;font-weight:bold;">Defect</td><td style="padding:6px;color:#CC0000;font-weight:bold;">{failure}</td></tr>
                    <tr><td style="padding:6px;font-weight:bold;">OCC Count</td><td style="padding:6px;color:#CC0000;font-weight:bold;">{occ_str}</td></tr>
                    <tr><td style="padding:6px;font-weight:bold;">Pressure</td><td style="padding:6px;">{press} Bar</td></tr>
                    <tr><td style="padding:6px;font-weight:bold;">Inspector</td><td style="padding:6px;">{inspector}</td></tr>
                    <tr><td style="padding:6px;font-weight:bold;">Root Cause</td><td style="padding:6px;">{rc}</td></tr>
                    <tr><td style="padding:6px;font-weight:bold;">Actions</td><td style="padding:6px;">{actions}</td></tr>
                    <tr><td style="padding:6px;font-weight:bold;">ECD</td><td style="padding:6px;">{ecd}</td></tr>
                </table>
            </div>
            <div style="padding:10px;font-size:11px;color:#888;">
                Auto-generated by HemmTrack Pro V2 | Tata Motors PVBU, Pune
            </div>
        </div>
        """

        # ── Telegram message ──
        tg_text = (
            f"🚨 <b>CRITICAL DEFECT ALERT</b>\n\n"
            f"📅 Date: {date}\n"
            f"🏭 Station: {station}\n"
            f"🚪 Side: {side}\n"
            f"🔍 Defect: {failure}\n"
            f"📊 OCC: {occ_str}\n"
            f"💨 Pressure: {press} Bar\n"
            f"👤 Inspector: {inspector}\n"
            f"🔧 Root Cause: {rc}\n"
            f"Shift: {shift} | Model: {model}\n\n"
        )

        result = {"success": True}

        if occ_num >= 5:
            # ═══ OCC >= 5: Generate PPT + send with attachment ═══
            logger.info(f"OCC={occ_num} — Generating PPT...")

            ppt_bytes = generate_defect_ppt(p)
            fname = f"Defect_Report_OCC{occ_str}_{failure.replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
            size_kb = round(len(ppt_bytes) / 1024)

            logger.info(f"PPT generated: {fname} ({size_kb} KB)")

            # Email with PPT attachment
            email_subject = f"🚨 CRITICAL: {failure} — OCC {occ_str} — Auto Report Attached"
            email_html += f'<p style="color:#CC0000;font-weight:bold;font-size:14px;">📎 PPT Report Attached ({size_kb} KB)</p>'

            email_result = send_email(email_subject, email_html,
                                      attachment_bytes=ppt_bytes,
                                      attachment_filename=fname)

            # Telegram: message + PPT document
            tg_text += f"📎 PPT Report: {fname} ({size_kb} KB)\n— HemmTrack Pro V2"
            tg_msg_result = send_telegram_message(tg_text)
            tg_doc_result = send_telegram_document(ppt_bytes, fname,
                f"🚨 OCC {occ_str} — {failure} — Auto Report")

            result["ppt"] = {"sizeKB": str(size_kb), "filename": fname}
            result["email"] = "sent" if email_result["ok"] else email_result.get("error", "failed")
            result["telegram"] = "sent" if tg_msg_result["ok"] else tg_msg_result.get("error", "failed")
            result["telegramDoc"] = "sent" if tg_doc_result["ok"] else tg_doc_result.get("error", "failed")

            if not email_result["ok"] and not tg_msg_result["ok"]:
                result["success"] = False
                result["error"] = "Both email and Telegram failed"

        else:
            # ═══ OCC < 5 (e.g., OCC=3): Email + Telegram text only ═══
            email_subject = f"🔔 Defect Alert: {failure} — OCC {occ_str}"
            tg_text += "⚡ Immediate action required!\n— HemmTrack Pro V2"

            email_result = send_email(email_subject, email_html)
            tg_result = send_telegram_message(tg_text)

            result["email"] = "sent" if email_result["ok"] else email_result.get("error", "failed")
            result["telegram"] = "sent" if tg_result["ok"] else tg_result.get("error", "failed")

            if not email_result["ok"] and not tg_result["ok"]:
                result["success"] = False
                result["error"] = "Both email and Telegram failed"

        elapsed = f"{time.time() - start:.1f}s"
        result["elapsed"] = elapsed
        logger.info(f"/check_occ done — OCC={occ_str}, elapsed={elapsed}, success={result['success']}")

        return jsonify(result), 200

    except Exception as e:
        logger.error(f"check_occ error: {e}\n{traceback.format_exc()}")
        elapsed = f"{time.time() - start:.1f}s"
        return jsonify({"success": False, "error": str(e), "elapsed": elapsed}), 500


# ══════════════════════════════════════════════════════════════
#  DEBUG ENDPOINTS
# ══════════════════════════════════════════════════════════════
@app.route("/debug/env", methods=["GET"])
def debug_env():
    return jsonify({
        "RESEND_API_KEY": env_status("RESEND_API_KEY"),
        "RESEND_FROM_EMAIL": get_env("RESEND_FROM_EMAIL") or "EMPTY",
        "ALERT_TO_EMAIL": env_status("ALERT_TO_EMAIL"),
        "TELEGRAM_BOT_TOKEN": env_status("TELEGRAM_BOT_TOKEN"),
        "TELEGRAM_CHAT_ID": get_env("TELEGRAM_CHAT_ID") or "EMPTY",
    })


@app.route("/debug/test_email", methods=["GET"])
def debug_test_email():
    r = send_email("HemmTrack Pro — Email Test",
                   f"<h2>Email working</h2><p>{datetime.now(timezone.utc).isoformat()}</p>")
    return jsonify({"status": "success" if r["ok"] else "error", **r})


@app.route("/debug/test_telegram", methods=["GET"])
def debug_test_telegram():
    r = send_telegram_message(f"<b>HemmTrack Pro — Telegram Test</b>\n{datetime.now(timezone.utc).isoformat()}")
    return jsonify({"status": "success" if r["ok"] else "error", **r})


@app.route("/debug/test_ppt", methods=["GET"])
def debug_test_ppt():
    """Generate test PPT → send via email + Telegram."""
    test_payload = {
        "occ": "05", "model": "Altroz", "failure": "Open Hem",
        "station": "ST-200", "rc": "Roller pressure low", "press": "165",
        "shift": "A", "side": "LH", "inspector": "Test Inspector",
        "date": datetime.now().strftime("%d %b %Y"),
        "actions": "Roller pressure adjusted to 185 Bar", "ecd": "Immediate",
        "highHistory": [
            {"date": "18 Mar 2026", "station": "ST-200", "defect": "Open Hem", "rootCause": "Low pressure", "pressure": "160"},
            {"date": "17 Mar 2026", "station": "ST-150", "defect": "Wrinkle", "rootCause": "Sheet alignment", "pressure": "175"},
        ]
    }

    try:
        ppt_bytes = generate_defect_ppt(test_payload)
        fname = f"Test_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        size_kb = round(len(ppt_bytes) / 1024)

        email_r = send_email("HemmTrack Pro — Test PPT",
                            "<h2>Test PPT attached</h2>",
                            attachment_bytes=ppt_bytes, attachment_filename=fname)
        tg_r = send_telegram_document(ppt_bytes, fname, "Test PPT — HemmTrack Pro")

        return jsonify({
            "success": True,
            "ppt": {"sizeKB": str(size_kb), "filename": fname},
            "email": "sent" if email_r["ok"] else email_r.get("error"),
            "telegram": "sent" if tg_r["ok"] else tg_r.get("error"),
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/debug/test_all", methods=["GET"])
def debug_test_all():
    return jsonify({
        "env": {k: env_status(k) for k in ["RESEND_API_KEY", "TELEGRAM_BOT_TOKEN", "ALERT_TO_EMAIL"]},
        "email": send_email("Full Test", "<p>Test</p>"),
        "telegram": send_telegram_message("Full System Test — HemmTrack Pro"),
    })


# ══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    port = int(get_env("PORT") or "5000")
    logger.info(f"Starting HemmTrack Pro v3 on port {port}")
    app.run(host="0.0.0.0", port=port, debug=False)
