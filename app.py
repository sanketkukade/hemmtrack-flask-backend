"""
============================================================
 HemmTrack Pro V2 — Flask Backend (Railway Deploy Ready)
============================================================
 Endpoint: POST /check_occ
 
 OCC = 3 → Email alert + Telegram alert
 OCC = 5 → Email with PPT attachment + Telegram escalation
 
 Both alerts are handled by SEPARATE functions.
 Zero conflict between them.
============================================================
"""

import os
import logging
import time
import tempfile
from datetime import datetime

from flask import Flask, request, jsonify
from flask_cors import CORS

from send_email import send_occ3_email, send_occ5_email_with_ppt
from send_telegram import send_occ3_telegram, send_occ5_telegram
from generate_ppt import generate_escalation_ppt

# ── Logging ──────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%d-%b-%Y %H:%M:%S"
)
log = logging.getLogger("hemmtrack")

# ── Flask App ────────────────────────────────────────────
app = Flask(__name__)
CORS(app)  # Allow frontend (GitHub Pages) to call this


# ===========================================================
#  HEALTH CHECK
# ===========================================================
@app.route("/", methods=["GET"])
def index():
    return jsonify({
        "status": "ok",
        "service": "HemmTrack Pro V2 — Flask Backend",
        "version": "1.0.0",
        "endpoints": {
            "health": "GET /",
            "alert": "POST /check_occ"
        }
    })


@app.route("/health", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "smtp": bool(os.getenv("SMTP_USER")),
        "telegram": bool(os.getenv("TELEGRAM_BOT_TOKEN"))
    })


# ===========================================================
#  POST /check_occ — Main Alert Router
# ===========================================================
@app.route("/check_occ", methods=["POST"])
def check_occ():
    start = time.time()

    log.info("=" * 50)
    log.info("📥 /check_occ request received")
    log.info("=" * 50)

    try:
        data = request.get_json(force=True)
        if not data:
            return jsonify({"success": False, "error": "Empty JSON body"}), 400

        occ = data.get("occ")
        if occ is None:
            return jsonify({"success": False, "error": "Missing 'occ' field"}), 400

        # ── Parse OCC as integer ──
        try:
            occ_int = int(str(occ).strip().lstrip("0") or "0")
        except ValueError:
            return jsonify({"success": False, "error": f"Invalid OCC value: {occ}"}), 400

        log.info(f"📊 OCC Value = {occ_int}")
        log.info(f"📋 Data: {data}")

        # ─────────────────────────────────────────────
        #  OCC = 3 → Email + Telegram (simple alert)
        # ─────────────────────────────────────────────
        if occ_int == 3:
            log.info("🔔 OCC=3 — Triggering standard alert...")
            result = handle_occ3(data)
            elapsed = round((time.time() - start) * 1000)
            result["elapsed"] = f"{elapsed}ms"
            log.info(f"✅ OCC=3 complete in {elapsed}ms")
            return jsonify(result)

        # ─────────────────────────────────────────────
        #  OCC = 5 → Email with PPT + Telegram
        # ─────────────────────────────────────────────
        elif occ_int == 5:
            log.info("🚨 OCC=5 — Triggering escalation with PPT...")
            result = handle_occ5(data)
            elapsed = round((time.time() - start) * 1000)
            result["elapsed"] = f"{elapsed}ms"
            log.info(f"✅ OCC=5 complete in {elapsed}ms")
            return jsonify(result)

        # ─────────────────────────────────────────────
        #  Other OCC values → Acknowledged, no action
        # ─────────────────────────────────────────────
        else:
            log.info(f"ℹ️ OCC={occ_int} — No alert threshold. Acknowledged.")
            return jsonify({
                "success": True,
                "message": f"OCC={occ_int} received. No alert threshold reached.",
                "action": "none"
            })

    except Exception as e:
        log.exception("❌ Unhandled error in /check_occ")
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


# ===========================================================
#  OCC = 3 Handler — Email + Telegram
# ===========================================================
def handle_occ3(data: dict) -> dict:
    """
    OCC=3: Send simple email alert + Telegram notification.
    NO PPT attachment. Separate from OCC=5.
    """
    results = {
        "success": True,
        "occ": 3,
        "action": "alert",
        "email": {"sent": False},
        "telegram": {"sent": False}
    }

    # ── Step 1: Email ──
    try:
        log.info("  📧 Step 1/2 — Sending OCC=3 email...")
        email_result = send_occ3_email(data)
        results["email"] = {"sent": True, "detail": email_result}
        log.info(f"  ✅ Email sent")
    except Exception as e:
        log.error(f"  ❌ Email failed: {e}")
        results["email"] = {"sent": False, "error": str(e)}

    # ── Step 2: Telegram ──
    try:
        log.info("  📱 Step 2/2 — Sending OCC=3 Telegram...")
        tg_result = send_occ3_telegram(data)
        results["telegram"] = {"sent": True, "detail": tg_result}
        log.info(f"  ✅ Telegram sent")
    except Exception as e:
        log.error(f"  ❌ Telegram failed: {e}")
        results["telegram"] = {"sent": False, "error": str(e)}

    return results


# ===========================================================
#  OCC = 5 Handler — PPT Generation + Email + Telegram
# ===========================================================
def handle_occ5(data: dict) -> dict:
    """
    OCC=5: Generate PPT → Email with attachment → Telegram.
    COMPLETELY SEPARATE from OCC=3.
    """
    results = {
        "success": True,
        "occ": 5,
        "action": "escalation",
        "ppt": {"generated": False},
        "email": {"sent": False},
        "telegram": {"sent": False}
    }

    ppt_path = None

    try:
        # ── Step 1: Generate PPT ──
        log.info("  📄 Step 1/3 — Generating escalation PPT...")
        ppt_path = generate_escalation_ppt(data)
        file_size = os.path.getsize(ppt_path)
        file_name = os.path.basename(ppt_path)
        results["ppt"] = {
            "generated": True,
            "fileName": file_name,
            "sizeKB": round(file_size / 1024, 1)
        }
        log.info(f"  ✅ PPT generated: {file_name} ({results['ppt']['sizeKB']} KB)")

        # ── Step 2: Email with PPT ──
        try:
            log.info("  📧 Step 2/3 — Sending OCC=5 email with PPT attachment...")
            email_result = send_occ5_email_with_ppt(data, ppt_path)
            results["email"] = {"sent": True, "detail": email_result}
            log.info(f"  ✅ Email sent with attachment")
        except Exception as e:
            log.error(f"  ❌ Email failed: {e}")
            results["email"] = {"sent": False, "error": str(e)}

        # ── Step 3: Telegram ──
        try:
            log.info("  📱 Step 3/3 — Sending OCC=5 Telegram escalation...")
            tg_result = send_occ5_telegram(data)
            results["telegram"] = {"sent": True, "detail": tg_result}
            log.info(f"  ✅ Telegram sent")
        except Exception as e:
            log.error(f"  ❌ Telegram failed: {e}")
            results["telegram"] = {"sent": False, "error": str(e)}

    except Exception as e:
        log.error(f"  ❌ PPT generation failed: {e}")
        results["ppt"] = {"generated": False, "error": str(e)}
        results["success"] = False

    finally:
        # ── Cleanup temp PPT file ──
        if ppt_path and os.path.exists(ppt_path):
            try:
                os.remove(ppt_path)
                log.info(f"  🗑️ Temp PPT cleaned up")
            except OSError:
                pass

    return results


# ===========================================================
#  Run Server
# ===========================================================
if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    log.info("")
    log.info("=" * 52)
    log.info(" HemmTrack Pro V2 — Flask Backend")
    log.info("=" * 52)
    log.info(f" 🚀 Server starting on port {port}")
    log.info(f" 📧 SMTP: {os.getenv('SMTP_USER', '⚠ NOT SET')}")
    log.info(f" 📱 Telegram: {'✅' if os.getenv('TELEGRAM_BOT_TOKEN') else '⚠ NOT SET'}")
    log.info(f" 🎯 POST /check_occ")
    log.info("=" * 52)
    log.info("")
    app.run(host="0.0.0.0", port=port, debug=False)
