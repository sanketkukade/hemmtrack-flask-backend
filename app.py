"""
HemmTrack Pro — Flask Backend (Production-Ready v2)
====================================================
Railway deployment with Resend HTTP API (email) + Telegram Bot API.

Why Resend instead of SMTP?
- Railway BLOCKS outbound SMTP connections (ports 587, 465).
- OSError: [Errno 101] Network is unreachable — Railway's network policy.
- Resend uses HTTPS (port 443) which Railway allows.
- Free tier: 3,000 emails/month, no credit card needed.
- Single HTTP POST call — simpler than smtplib.

Required Railway Environment Variables:
  RESEND_API_KEY       → from resend.com/api-keys
  RESEND_FROM_EMAIL    → verified sender (onboarding@resend.dev for testing,
                         or alerts@yourdomain.com after domain verification)
  ALERT_TO_EMAIL       → recipient email address
  TELEGRAM_BOT_TOKEN   → from @BotFather
  TELEGRAM_CHAT_ID     → your chat/group ID
"""

import os
import logging
import traceback
from datetime import datetime, timezone

import requests
from flask import Flask, jsonify, request

# ──────────────────────────────────────────────────────────────
# Logging — Railway captures stdout
# ──────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("hemmtrack")

app = Flask(__name__)


# ══════════════════════════════════════════════════════════════
#  HELPER: Runtime Environment Variable Loader
# ══════════════════════════════════════════════════════════════
def get_env(key: str) -> str:
    """
    Read env var AT CALL TIME — never at import time.
    Railway injects vars at process startup; module-level reads
    can race and get empty strings.
    """
    return os.environ.get(key, "").strip()


def env_status(key: str) -> str:
    """Return 'SET (N chars)' or 'EMPTY' — never the actual value."""
    val = get_env(key)
    return f"SET ({len(val)} chars)" if val else "EMPTY"


# ══════════════════════════════════════════════════════════════
#  EMAIL: Resend HTTP API (Replaces blocked SMTP)
# ══════════════════════════════════════════════════════════════
def send_email_alert(subject: str, body_html: str, to_email: str = None) -> dict:
    """
    Send email via Resend HTTP API.

    Why Resend over SMTP?
    - Railway blocks SMTP (port 587/465) → OSError: Network is unreachable
    - Resend uses HTTPS (port 443) → works on Railway
    - Free tier: 3,000 emails/month
    - One POST request — no SMTP handshake/TLS/ehlo dance

    API docs: https://resend.com/docs/api-reference/emails/send-email
    """
    api_key = get_env("RESEND_API_KEY")
    from_email = get_env("RESEND_FROM_EMAIL")
    alert_to = to_email or get_env("ALERT_TO_EMAIL")

    # ── Pre-flight validation ──
    errors = []
    if not api_key:
        errors.append("RESEND_API_KEY is empty or not set")
    if not from_email:
        errors.append("RESEND_FROM_EMAIL is empty or not set")
    if not alert_to:
        errors.append("ALERT_TO_EMAIL is empty or not set")

    if errors:
        msg = "Pre-flight failed: " + "; ".join(errors)
        logger.error(msg)
        return {"status": "error", "stage": "validation", "reason": msg}

    # ── Resend API call ──
    url = "https://api.resend.com/emails"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "from": from_email,
        "to": [alert_to],
        "subject": subject,
        "html": body_html,
    }

    try:
        logger.info(f"Resend: Sending to {alert_to} from {from_email}")

        resp = requests.post(url, json=payload, headers=headers, timeout=15)
        data = resp.json()

        # ── Success: HTTP 200 with {"id": "email_xxx"} ──
        if resp.status_code == 200 and "id" in data:
            email_id = data["id"]
            logger.info(f"Resend: Sent (id={email_id})")
            return {
                "status": "success",
                "email_id": email_id,
                "to": alert_to,
                "from": from_email,
                "timestamp": datetime.now(timezone.utc).isoformat(),
            }

        # ── API error ──
        error_msg = data.get("message", "Unknown error")
        status_code = data.get("statusCode", resp.status_code)

        # Actionable hints for common errors
        if status_code == 401 or "API key" in str(error_msg):
            error_msg += (
                " → RESEND_API_KEY is invalid or expired. "
                "Go to resend.com/api-keys and generate a new one."
            )
        elif status_code == 403:
            error_msg += (
                " → Sender not verified. Use 'onboarding@resend.dev' for "
                "testing, or verify your domain at resend.com/domains."
            )
        elif status_code == 422:
            error_msg += (
                " → Invalid request. Check that RESEND_FROM_EMAIL and "
                "ALERT_TO_EMAIL are valid email addresses."
            )
        elif status_code == 429:
            error_msg += " → Rate limit hit. Free tier = 3,000 emails/month."

        reason = f"Resend API error {status_code}: {error_msg}"
        logger.error(reason)
        return {"status": "error", "stage": "api_response", "reason": reason}

    except requests.exceptions.Timeout:
        reason = "Resend API timed out after 15 seconds"
        logger.error(reason)
        return {"status": "error", "stage": "timeout", "reason": reason}

    except requests.exceptions.ConnectionError as e:
        reason = f"Cannot reach api.resend.com: {e}"
        logger.error(reason)
        return {"status": "error", "stage": "connection", "reason": reason}

    except Exception as e:
        reason = f"Unexpected: {type(e).__name__}: {e}\n{traceback.format_exc()}"
        logger.error(reason)
        return {"status": "error", "stage": "unknown", "reason": reason}


# ══════════════════════════════════════════════════════════════
#  TELEGRAM: Bot API Sender
# ══════════════════════════════════════════════════════════════
def send_telegram_alert(message: str) -> dict:
    """Send Telegram message via Bot API (HTTPS — works on Railway)."""
    bot_token = get_env("TELEGRAM_BOT_TOKEN")
    chat_id = get_env("TELEGRAM_CHAT_ID")

    if not bot_token:
        msg = "TELEGRAM_BOT_TOKEN is empty or not set"
        logger.error(msg)
        return {"status": "error", "stage": "validation", "reason": msg}

    if not chat_id:
        msg = "TELEGRAM_CHAT_ID is empty or not set"
        logger.error(msg)
        return {"status": "error", "stage": "validation", "reason": msg}

    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": message,
        "parse_mode": "HTML",
    }

    try:
        logger.info(f"Telegram: Sending to chat_id={chat_id}")
        resp = requests.post(url, json=payload, timeout=15)
        data = resp.json()

        if data.get("ok"):
            msg_id = data["result"]["message_id"]
            logger.info(f"Telegram: Sent (message_id={msg_id})")
            return {
                "status": "success",
                "chat_id": chat_id,
                "message_id": msg_id,
                "timestamp": datetime.now(timezone.utc).isoformat(),
            }
        else:
            error_code = data.get("error_code", "unknown")
            description = data.get("description", "No description")
            reason = f"Telegram error {error_code}: {description}"

            if error_code == 401:
                reason += " → Token invalid. Get fresh from @BotFather."
            elif error_code == 400 and "chat not found" in description.lower():
                reason += " → Send /start to the bot first."

            logger.error(reason)
            return {"status": "error", "stage": "api_response", "reason": reason}

    except requests.exceptions.Timeout:
        reason = "Telegram timed out after 15 seconds"
        logger.error(reason)
        return {"status": "error", "stage": "timeout", "reason": reason}

    except requests.exceptions.ConnectionError as e:
        reason = f"Cannot reach api.telegram.org: {e}"
        logger.error(reason)
        return {"status": "error", "stage": "connection", "reason": reason}

    except Exception as e:
        reason = f"Unexpected: {type(e).__name__}: {e}"
        logger.error(reason)
        return {"status": "error", "stage": "unknown", "reason": reason}


# ══════════════════════════════════════════════════════════════
#  API ENDPOINTS
# ══════════════════════════════════════════════════════════════

@app.route("/", methods=["GET"])
def health():
    return jsonify({
        "service": "HemmTrack Pro Backend",
        "version": "2.0 — Resend + Telegram",
        "status": "running",
        "email_method": "Resend HTTP API (SMTP removed — blocked by Railway)",
        "timestamp": datetime.now(timezone.utc).isoformat(),
    })


@app.route("/check_occ", methods=["POST"])
def check_occ():
    """
    Main OCC alert endpoint.
    Sends email (Resend) + Telegram.
    Returns actual success/failure — never silently fails.
    """
    try:
        data = request.get_json(silent=True) or {}
        logger.info(f"/check_occ called with keys: {list(data.keys())}")

        subject = data.get("subject", "🚨 OCC Alert — HemmTrack Pro")
        body_html = data.get("body", (
            "<h2>🚨 OCC Threshold Exceeded</h2>"
            "<p>An OCC alert has been triggered.</p>"
            "<p>Check the HemmTrack Pro dashboard for details.</p>"
            f"<p><small>Alert time: {datetime.now(timezone.utc).isoformat()}</small></p>"
        ))
        tg_message = data.get("message",
            "🚨 OCC Alert — HemmTrack Pro: Threshold exceeded. Check dashboard."
        )

        # ── Send both ──
        email_result = send_email_alert(subject, body_html)
        telegram_result = send_telegram_alert(tg_message)

        # ── Overall status ──
        email_ok = email_result.get("status") == "success"
        telegram_ok = telegram_result.get("status") == "success"

        if email_ok and telegram_ok:
            overall, http_code = "success", 200
        elif email_ok or telegram_ok:
            overall, http_code = "partial", 207
        else:
            overall, http_code = "failed", 502

        response = {
            "status": overall,
            "email": email_result,
            "telegram": telegram_result,
            "timestamp": datetime.now(timezone.utc).isoformat(),
        }

        logger.info(
            f"/check_occ → {overall} | "
            f"email={email_result['status']} | "
            f"telegram={telegram_result['status']}"
        )
        return jsonify(response), http_code

    except Exception as e:
        error_msg = f"Unhandled: {type(e).__name__}: {e}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return jsonify({"status": "error", "reason": error_msg}), 500


# ══════════════════════════════════════════════════════════════
#  DEBUG ENDPOINTS
# ══════════════════════════════════════════════════════════════

@app.route("/debug/env", methods=["GET"])
def debug_env():
    """Check env vars — shows SET/EMPTY, never actual values."""
    return jsonify({
        "RESEND_API_KEY": env_status("RESEND_API_KEY"),
        "RESEND_FROM_EMAIL": get_env("RESEND_FROM_EMAIL") or "EMPTY",
        "ALERT_TO_EMAIL": env_status("ALERT_TO_EMAIL"),
        "TELEGRAM_BOT_TOKEN": env_status("TELEGRAM_BOT_TOKEN"),
        "TELEGRAM_CHAT_ID": get_env("TELEGRAM_CHAT_ID") or "EMPTY",
        "PORT": get_env("PORT") or "default",
        "note": "SMTP vars removed — Railway blocks SMTP. Using Resend.",
        "old_smtp_vars": {
            "SMTP_USER": "still set (safe to remove)" if get_env("SMTP_USER") else "not set",
            "SMTP_PASS": "still set (safe to remove)" if get_env("SMTP_PASS") else "not set",
        },
    })


@app.route("/debug/test_email", methods=["GET"])
def debug_test_email():
    """Test email via Resend."""
    result = send_email_alert(
        subject="✅ HemmTrack Pro — Email Test (Resend)",
        body_html=(
            "<h2>✅ Email Test Successful</h2>"
            "<p>Resend HTTP API is working from Railway.</p>"
            "<p>SMTP is no longer needed.</p>"
            f"<p><small>{datetime.now(timezone.utc).isoformat()}</small></p>"
        ),
    )
    return jsonify(result), 200 if result["status"] == "success" else 502


@app.route("/debug/test_telegram", methods=["GET"])
def debug_test_telegram():
    """Test Telegram."""
    result = send_telegram_alert(
        f"✅ <b>HemmTrack Pro — Telegram Test</b>\n"
        f"Bot API working from Railway.\n"
        f"<i>{datetime.now(timezone.utc).isoformat()}</i>"
    )
    return jsonify(result), 200 if result["status"] == "success" else 502


@app.route("/debug/test_all", methods=["GET"])
def debug_test_all():
    """Run all tests at once."""
    env_data = {
        "RESEND_API_KEY": env_status("RESEND_API_KEY"),
        "RESEND_FROM_EMAIL": get_env("RESEND_FROM_EMAIL") or "EMPTY",
        "ALERT_TO_EMAIL": env_status("ALERT_TO_EMAIL"),
        "TELEGRAM_BOT_TOKEN": env_status("TELEGRAM_BOT_TOKEN"),
        "TELEGRAM_CHAT_ID": get_env("TELEGRAM_CHAT_ID") or "EMPTY",
    }

    email_result = send_email_alert(
        subject="✅ HemmTrack Pro — Full System Test",
        body_html="<h2>Full test passed</h2><p>Email via Resend working.</p>",
    )

    telegram_result = send_telegram_alert(
        "✅ HemmTrack Pro — Full System Test\nAll systems operational."
    )

    return jsonify({
        "env_check": env_data,
        "email_test": email_result,
        "telegram_test": telegram_result,
        "timestamp": datetime.now(timezone.utc).isoformat(),
    })


# ══════════════════════════════════════════════════════════════
#  ENTRYPOINT
# ══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    port = int(get_env("PORT") or "5000")
    logger.info(f"Starting HemmTrack Pro backend on port {port}")
    app.run(host="0.0.0.0", port=port, debug=False)
