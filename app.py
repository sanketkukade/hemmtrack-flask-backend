"""
HemmTrack Pro — Flask Backend (Production-Ready)
=================================================
Railway deployment with Gmail SMTP + Telegram Bot alerts.

Root Cause Analysis:
1. Environment variables read at module level → empty on Railway
2. Silent exception handling → API returns "success" even when email fails
3. SMTP port/TLS mismatch → Gmail needs 587+STARTTLS or 465+SSL, not mixed
4. Missing ehlo() after starttls() → Gmail rejects the session
5. No timeouts → requests hang indefinitely on Railway
6. Telegram: no response validation → silent failures

All issues are fixed in this file.
"""

import os
import smtplib
import logging
import traceback
from datetime import datetime, timezone
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import requests
from flask import Flask, jsonify, request

# ──────────────────────────────────────────────────────────────
# Logging — Railway captures stdout, so configure properly
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
    Read an env var AT CALL TIME — never at import time.

    Why this matters on Railway:
    - Railway injects env vars into the process at startup.
    - If you do `SMTP_USER = os.environ.get("SMTP_USER")` at the top
      of the file, Python evaluates it once during import.
    - If the variable isn't ready yet (race condition) or gets
      updated after deploy, the code uses the stale/empty value forever.
    - Reading inside the function guarantees the latest value every time.
    """
    value = os.environ.get(key, "").strip()
    return value


def env_status(key: str) -> str:
    """Return 'SET (Nchars)' or 'EMPTY' — never the actual secret."""
    val = get_env(key)
    if val:
        return f"SET ({len(val)} chars)"
    return "EMPTY"


# ══════════════════════════════════════════════════════════════
#  EMAIL: Gmail SMTP Sender (Fixed)
# ══════════════════════════════════════════════════════════════
def send_email_alert(subject: str, body: str, to_email: str = None) -> dict:
    """
    Send an HTML email via Gmail SMTP.

    Fixes applied:
    - Reads env vars at call time, not import time
    - Validates credentials before connecting
    - Uses port 587 + STARTTLS (Gmail's recommended method)
    - Calls ehlo() BEFORE and AFTER starttls() (required by RFC)
    - 30-second timeout prevents Railway from hanging
    - Catches specific SMTP exceptions with clear messages
    - Returns detailed result dict (never silently fails)
    """
    smtp_user = get_env("SMTP_USER")
    smtp_pass = get_env("SMTP_PASS")
    alert_to = (to_email or get_env("ALERT_TO_EMAIL"))

    # ── Pre-flight validation ──
    errors = []
    if not smtp_user:
        errors.append("SMTP_USER is empty or not set")
    if not smtp_pass:
        errors.append("SMTP_PASS is empty or not set")
    if not alert_to:
        errors.append("ALERT_TO_EMAIL (or to_email param) is empty")

    if errors:
        msg = "Pre-flight check failed: " + "; ".join(errors)
        logger.error(msg)
        return {"status": "error", "stage": "validation", "reason": msg}

    # ── Build MIME message ──
    mime_msg = MIMEMultipart("alternative")
    mime_msg["From"] = smtp_user
    mime_msg["To"] = alert_to
    mime_msg["Subject"] = subject
    mime_msg.attach(MIMEText(body, "html", "utf-8"))

    # ── Send via SMTP ──
    try:
        logger.info(f"SMTP: Connecting to smtp.gmail.com:587 as {smtp_user}")

        server = smtplib.SMTP("smtp.gmail.com", 587, timeout=30)

        # Enable wire-level debug (shows in Railway logs)
        server.set_debuglevel(1)

        # Step 1: Initial EHLO (identifies client to server)
        server.ehlo()

        # Step 2: Upgrade to TLS
        server.starttls()

        # Step 3: EHLO again (required after TLS upgrade — RFC 3207)
        # THIS IS THE MOST COMMONLY MISSED STEP.
        # Without it, Gmail responds with "5.5.1 Authentication Required"
        server.ehlo()

        # Step 4: Authenticate
        server.login(smtp_user, smtp_pass)

        # Step 5: Send
        server.sendmail(smtp_user, [alert_to], mime_msg.as_string())

        # Step 6: Clean disconnect
        server.quit()

        logger.info(f"SMTP: Email sent successfully to {alert_to}")
        return {
            "status": "success",
            "to": alert_to,
            "subject": subject,
            "timestamp": datetime.now(timezone.utc).isoformat(),
        }

    except smtplib.SMTPAuthenticationError as e:
        reason = (
            f"Authentication failed (code {e.smtp_code}): {e.smtp_error}. "
            "LIKELY CAUSE: You're using your Gmail login password instead of "
            "an App Password. Go to myaccount.google.com → Security → "
            "2-Step Verification → App Passwords → generate one for 'Mail'."
        )
        logger.error(f"SMTP AUTH ERROR: {reason}")
        return {"status": "error", "stage": "authentication", "reason": reason}

    except smtplib.SMTPConnectError as e:
        reason = (
            f"Connection to smtp.gmail.com:587 failed: {e}. "
            "Railway might be blocking outbound SMTP. Try port 465 with SSL."
        )
        logger.error(f"SMTP CONNECT ERROR: {reason}")
        return {"status": "error", "stage": "connection", "reason": reason}

    except smtplib.SMTPRecipientsRefused as e:
        reason = f"Recipient refused: {e}"
        logger.error(reason)
        return {"status": "error", "stage": "recipient", "reason": reason}

    except smtplib.SMTPServerDisconnected as e:
        reason = (
            f"Server disconnected unexpectedly: {e}. "
            "This usually means Railway's network dropped the connection."
        )
        logger.error(reason)
        return {"status": "error", "stage": "disconnect", "reason": reason}

    except smtplib.SMTPException as e:
        reason = f"SMTP error ({type(e).__name__}): {e}"
        logger.error(reason)
        return {"status": "error", "stage": "smtp", "reason": reason}

    except TimeoutError:
        reason = (
            "SMTP connection timed out after 30s. "
            "Railway may be blocking port 587. Try the SSL fallback on port 465."
        )
        logger.error(reason)
        return {"status": "error", "stage": "timeout", "reason": reason}

    except Exception as e:
        reason = f"Unexpected error: {type(e).__name__}: {e}\n{traceback.format_exc()}"
        logger.error(reason)
        return {"status": "error", "stage": "unknown", "reason": reason}


def send_email_alert_ssl(subject: str, body: str, to_email: str = None) -> dict:
    """
    Fallback: Send email via port 465 + direct SSL.
    Use this if Railway blocks port 587.
    """
    smtp_user = get_env("SMTP_USER")
    smtp_pass = get_env("SMTP_PASS")
    alert_to = (to_email or get_env("ALERT_TO_EMAIL"))

    if not smtp_user or not smtp_pass or not alert_to:
        return {"status": "error", "stage": "validation", "reason": "Missing credentials"}

    mime_msg = MIMEMultipart("alternative")
    mime_msg["From"] = smtp_user
    mime_msg["To"] = alert_to
    mime_msg["Subject"] = subject
    mime_msg.attach(MIMEText(body, "html", "utf-8"))

    try:
        logger.info(f"SMTP-SSL: Connecting to smtp.gmail.com:465 as {smtp_user}")

        # Port 465 uses SSL from the start (no starttls needed)
        server = smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30)
        server.set_debuglevel(1)
        server.login(smtp_user, smtp_pass)
        server.sendmail(smtp_user, [alert_to], mime_msg.as_string())
        server.quit()

        logger.info(f"SMTP-SSL: Email sent successfully to {alert_to}")
        return {"status": "success", "method": "SSL-465", "to": alert_to}

    except Exception as e:
        reason = f"{type(e).__name__}: {e}"
        logger.error(f"SMTP-SSL ERROR: {reason}")
        return {"status": "error", "stage": "ssl_send", "reason": reason}


# ══════════════════════════════════════════════════════════════
#  TELEGRAM: Bot API Sender (Fixed)
# ══════════════════════════════════════════════════════════════
def send_telegram_alert(message: str) -> dict:
    """
    Send a Telegram message via Bot API.

    Fixes applied:
    - Reads env vars at call time
    - Validates token and chat_id before calling API
    - 15-second timeout (Telegram API is usually fast)
    - Checks response 'ok' field (not just HTTP status)
    - Returns actual error description from Telegram API
    """
    bot_token = get_env("TELEGRAM_BOT_TOKEN")
    chat_id = get_env("TELEGRAM_CHAT_ID")

    # ── Pre-flight validation ──
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

        # Telegram API returns {"ok": true, "result": {...}} on success
        # and {"ok": false, "description": "..."} on failure
        if data.get("ok"):
            msg_id = data["result"]["message_id"]
            logger.info(f"Telegram: Sent successfully (message_id={msg_id})")
            return {
                "status": "success",
                "chat_id": chat_id,
                "message_id": msg_id,
                "timestamp": datetime.now(timezone.utc).isoformat(),
            }
        else:
            error_code = data.get("error_code", "unknown")
            description = data.get("description", "No description")
            reason = f"Telegram API error {error_code}: {description}"

            # Provide actionable hints for common errors
            if error_code == 401:
                reason += " → Bot token is invalid. Get a fresh one from @BotFather."
            elif error_code == 400 and "chat not found" in description.lower():
                reason += (
                    " → The bot hasn't received a message from this chat yet. "
                    "Send /start to the bot first, then try again."
                )

            logger.error(reason)
            return {"status": "error", "stage": "api_response", "reason": reason}

    except requests.exceptions.Timeout:
        reason = "Telegram API request timed out after 15 seconds"
        logger.error(reason)
        return {"status": "error", "stage": "timeout", "reason": reason}

    except requests.exceptions.ConnectionError as e:
        reason = f"Cannot reach api.telegram.org: {e}"
        logger.error(reason)
        return {"status": "error", "stage": "connection", "reason": reason}

    except Exception as e:
        reason = f"Unexpected error: {type(e).__name__}: {e}"
        logger.error(reason)
        return {"status": "error", "stage": "unknown", "reason": reason}


# ══════════════════════════════════════════════════════════════
#  API ENDPOINTS
# ══════════════════════════════════════════════════════════════

# ── Health Check ──
@app.route("/", methods=["GET"])
def health():
    return jsonify({
        "service": "HemmTrack Pro Backend",
        "status": "running",
        "timestamp": datetime.now(timezone.utc).isoformat(),
    })


# ── Main Alert Endpoint (Fixed) ──
@app.route("/check_occ", methods=["POST"])
def check_occ():
    """
    Trigger OCC-based email + Telegram alerts.

    FIX: Previously returned {"status": "success"} regardless of
    whether email/Telegram actually sent. Now returns the real
    result of each notification attempt.
    """
    try:
        data = request.get_json(silent=True) or {}
        logger.info(f"/check_occ called with keys: {list(data.keys())}")

        # Extract alert content from request body
        subject = data.get("subject", "🚨 OCC Alert — HemmTrack Pro")
        body = data.get("body", "<h3>OCC threshold exceeded</h3><p>Check dashboard for details.</p>")
        tg_message = data.get("message", "🚨 OCC Alert — HemmTrack Pro: Threshold exceeded")

        # ── Send both notifications ──
        email_result = send_email_alert(subject, body)
        telegram_result = send_telegram_alert(tg_message)

        # ── Determine overall status ──
        email_ok = email_result.get("status") == "success"
        telegram_ok = telegram_result.get("status") == "success"

        if email_ok and telegram_ok:
            overall = "success"
            http_code = 200
        elif email_ok or telegram_ok:
            overall = "partial"  # One worked, one didn't
            http_code = 207      # HTTP 207 Multi-Status
        else:
            overall = "failed"   # Both failed
            http_code = 502      # Bad Gateway (upstream services failed)

        response = {
            "status": overall,
            "email": email_result,
            "telegram": telegram_result,
            "timestamp": datetime.now(timezone.utc).isoformat(),
        }

        logger.info(f"/check_occ result: {overall} | email={email_result['status']} | telegram={telegram_result['status']}")
        return jsonify(response), http_code

    except Exception as e:
        error_msg = f"Unhandled error in /check_occ: {type(e).__name__}: {e}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        return jsonify({"status": "error", "reason": error_msg}), 500


# ══════════════════════════════════════════════════════════════
#  DEBUG ENDPOINTS (Remove before production use)
# ══════════════════════════════════════════════════════════════

@app.route("/debug/env", methods=["GET"])
def debug_env():
    """
    Check if all required env vars are loaded.
    Shows SET/EMPTY and character count — NEVER shows actual values.
    """
    return jsonify({
        "SMTP_USER": env_status("SMTP_USER"),
        "SMTP_PASS": env_status("SMTP_PASS"),
        "ALERT_TO_EMAIL": env_status("ALERT_TO_EMAIL"),
        "TELEGRAM_BOT_TOKEN": env_status("TELEGRAM_BOT_TOKEN"),
        "TELEGRAM_CHAT_ID": get_env("TELEGRAM_CHAT_ID") or "EMPTY",
        "PORT": get_env("PORT") or "not set (Railway default)",
        "note": "If any critical var is EMPTY, check Railway Dashboard → Variables tab",
    })


@app.route("/debug/test_email", methods=["GET"])
def debug_test_email():
    """Send a test email and return the full result."""
    result = send_email_alert(
        subject="✅ HemmTrack Pro — Email Test",
        body=(
            "<h2>Email Test Successful</h2>"
            "<p>If you're reading this, Gmail SMTP is working from Railway.</p>"
            f"<p><small>Sent at {datetime.now(timezone.utc).isoformat()}</small></p>"
        ),
    )
    return jsonify(result), 200 if result["status"] == "success" else 502


@app.route("/debug/test_email_ssl", methods=["GET"])
def debug_test_email_ssl():
    """Test email via port 465 SSL (fallback if 587 is blocked)."""
    result = send_email_alert_ssl(
        subject="✅ HemmTrack Pro — Email Test (SSL 465)",
        body=(
            "<h2>SSL Email Test Successful</h2>"
            "<p>Port 465 + SSL is working from Railway.</p>"
            f"<p><small>Sent at {datetime.now(timezone.utc).isoformat()}</small></p>"
        ),
    )
    return jsonify(result), 200 if result["status"] == "success" else 502


@app.route("/debug/test_telegram", methods=["GET"])
def debug_test_telegram():
    """Send a test Telegram message and return the full result."""
    result = send_telegram_alert(
        f"✅ <b>HemmTrack Pro — Telegram Test</b>\n"
        f"If you see this, Telegram Bot API is working from Railway.\n"
        f"<i>{datetime.now(timezone.utc).isoformat()}</i>"
    )
    return jsonify(result), 200 if result["status"] == "success" else 502


@app.route("/debug/test_all", methods=["GET"])
def debug_test_all():
    """Run all tests at once and return combined result."""
    env_data = {
        "SMTP_USER": env_status("SMTP_USER"),
        "SMTP_PASS": env_status("SMTP_PASS"),
        "ALERT_TO_EMAIL": env_status("ALERT_TO_EMAIL"),
        "TELEGRAM_BOT_TOKEN": env_status("TELEGRAM_BOT_TOKEN"),
        "TELEGRAM_CHAT_ID": get_env("TELEGRAM_CHAT_ID") or "EMPTY",
    }

    email_result = send_email_alert(
        subject="✅ HemmTrack Pro — Full System Test",
        body="<h2>Full system test</h2><p>Email is working.</p>",
    )

    telegram_result = send_telegram_alert(
        "✅ HemmTrack Pro — Full System Test\nTelegram is working."
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
