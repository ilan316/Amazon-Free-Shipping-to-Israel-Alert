"""
Email notifier — sends a Gmail alert when free shipping to Israel is detected.

Setup:
  1. Enable 2-Step Verification on your Google account
  2. Go to: myaccount.google.com/apppasswords
  3. Create an App Password (select "Mail" + "Windows Computer")
  4. Copy the 16-character password into .env as: GMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx
"""

import smtplib
import os
import logging
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

logger = logging.getLogger(__name__)


def _smtp_send(config: dict, subject: str, text_body: str, html_body: str):
    """Internal helper — connects to SMTP and sends the message."""
    email_cfg   = config.get("email", {})
    sender      = email_cfg.get("sender", "")
    recipient   = email_cfg.get("recipient", "")
    smtp_host   = email_cfg.get("smtp_host", "smtp.gmail.com")
    smtp_port   = email_cfg.get("smtp_port", 587)

    if not sender or not recipient:
        raise RuntimeError("Email not configured — open Settings and enter your email address.")

    app_password = os.environ.get("GMAIL_APP_PASSWORD", "").replace(" ", "")
    if not app_password:
        raise RuntimeError("GMAIL_APP_PASSWORD missing from .env file.")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = sender
    msg["To"]      = recipient
    msg.attach(MIMEText(text_body, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html",  "utf-8"))

    try:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=20) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender, app_password)
            server.sendmail(sender, [recipient], msg.as_string())
        logger.info(f"Email sent to {recipient}: {subject}")
    except smtplib.SMTPAuthenticationError as e:
        raise RuntimeError(f"Gmail authentication failed — invalid App Password. ({e})")
    except smtplib.SMTPException as e:
        raise RuntimeError(f"SMTP error: {e}")
    except OSError as e:
        raise RuntimeError(f"Network error while sending email: {e}")


def send_batch_free_shipping_alert(config: dict, items: list):
    """
    Sends ONE email listing all FREE-shipping products.

    items — list of dicts: [{"product": {...}, "shipping_text": "..."}, ...]
    """
    if not items:
        return

    checked_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if len(items) == 1:
        name    = items[0]["product"].get("name", items[0]["product"].get("asin", ""))
        subject = f"FREE Shipping to Israel: {name}"
    else:
        subject = f"FREE Shipping to Israel: {len(items)} products found"

    # ── plain-text body ──────────────────────────────────────
    lines = ["FREE Shipping to Israel Alert!\n"]
    for item in items:
        p    = item["product"]
        asin = p.get("asin", "")
        name = p.get("name", asin)
        url  = p.get("url", f"https://www.amazon.com/dp/{asin}")
        lines += [
            f"Product : {name}",
            f"ASIN    : {asin}",
            f"URL     : {url}",
            f"Shipping: {item['shipping_text']}",
            "",
        ]
    lines += [f"Checked at: {checked_at}", "", "---", "Amazon Free Shipping Monitor"]
    text_body = "\n".join(lines)

    # ── HTML body ────────────────────────────────────────────
    product_rows = ""
    for item in items:
        p    = item["product"]
        asin = p.get("asin", "")
        name = p.get("name", asin)
        url  = p.get("url", f"https://www.amazon.com/dp/{asin}")
        shipping = item["shipping_text"].replace("<", "&lt;").replace(">", "&gt;")
        product_rows += f"""
  <div style="border:1px solid #e0e0e0; border-radius:6px; padding:12px 16px; margin-bottom:14px;">
    <p style="margin:0 0 6px 0; font-size:15px; font-weight:bold;">{name}</p>
    <table style="border-collapse:collapse; font-size:13px; color:#444;">
      <tr>
        <td style="padding:2px 12px 2px 0; font-weight:bold;">ASIN</td>
        <td>{asin}</td>
      </tr>
      <tr>
        <td style="padding:2px 12px 2px 0; font-weight:bold;">Link</td>
        <td><a href="{url}" style="color:#0066cc;">{url}</a></td>
      </tr>
    </table>
    <pre style="background:#f5f5f5; padding:8px; border-radius:4px; white-space:pre-wrap;
                font-size:12px; margin:8px 0 0 0;">{shipping}</pre>
  </div>"""

    html_body = f"""
<html>
<body style="font-family: Arial, sans-serif; color: #222; max-width:640px; margin:auto;">
  <h2 style="color: #e47911;">&#128666; FREE Shipping to Israel — {len(items)} product(s)</h2>
  {product_rows}
  <p style="color: #888; font-size: 12px; margin-top:20px;">Checked at: {checked_at}</p>
</body>
</html>
"""

    _smtp_send(config, subject, text_body, html_body)


def send_free_shipping_alert(config: dict, product: dict, shipping_text: str):
    """Single-product shortcut — delegates to the batch sender."""
    send_batch_free_shipping_alert(config, [{"product": product, "shipping_text": shipping_text}])
