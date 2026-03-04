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

# ── Localized strings ─────────────────────────────────────────────────────────
_STRINGS = {
    "he": {
        "subject_single":   "פרסומת | 🚨 משלוח חינם לישראל: {name}",
        "subject_multi":    "פרסומת | 🚨 משלוח חינם לישראל: {n} מוצרים",
        "preheader":        "אל תפספס! המחיר עשוי להשתנות בכל עת — לחץ לפרטים",
        "header_title":     "🚚 משלוח חינם לישראל!",
        "header_sub":       "נמצאו {n} מוצרים עם משלוח חינם",
        "header_sub1":      "נמצא מוצר עם משלוח חינם",
        "shipping_badge":   "✅ משלוח חינם לישראל · הזמנות $49+",
        "btn_buy":          "בדוק את המחיר באמזון ←",
        "urgency":          "⏰ המחיר עשוי להשתנות בכל עת",
        "quick_tip_title":  "💡 טיפ לחיסכון",
        "quick_tip_body":   "הזמינו בין $49 ל-$74.99 כדי ליהנות ממשלוח חינם ללא מכס ישראלי.",
        "disclosure":       "קישור שותף — הקנייה לא עולה לך יותר, אך אנו עשויים לקבל עמלה קטנה.",
        "footer":           "נבדק: {checked_at} · Amazon Free Shipping Monitor",
        "aod_note":         "⚠️ המשלוח החינמי נמצא תחת <strong>\"כל אפשרויות הקנייה\"</strong>.<br>"
                            "פתח את עמוד המוצר &larr; לחץ <strong>\"ראה את כל אפשרויות הקנייה\"</strong>"
                            " &larr; בחר את ההצעה עם משלוח חינם.",
        "aod_plain":        "הערה: המשלוח החינמי נמצא תחת 'כל אפשרויות הקנייה'. "
                            "פתח את הקישור ← לחץ 'ראה את כל אפשרויות הקנייה' ← בחר הצעה עם משלוח חינם.",
        "plain_header":     "🚨 התראת משלוח חינם לישראל!\n",
        "plain_product":    "מוצר",
        "plain_url":        "קישור",
        "plain_urgency":    "⏰ המחיר עשוי להשתנות בכל עת",
        "plain_footer":     "נבדק: {checked_at}",
    },
    "en": {
        "subject_single":   "🚨 FREE Shipping to Israel: {name}",
        "subject_multi":    "🚨 FREE Shipping to Israel: {n} products found",
        "preheader":        "Don't miss out! Price may change at any time — check now",
        "header_title":     "🚚 FREE Shipping to Israel!",
        "header_sub":       "{n} products with free shipping found",
        "header_sub1":      "1 product with free shipping found",
        "shipping_badge":   "✅ FREE Shipping to Israel · Orders $49+",
        "btn_buy":          "Check Today's Price on Amazon →",
        "urgency":          "⏰ Price may change at any time",
        "quick_tip_title":  "💡 Money-Saving Tip",
        "quick_tip_body":   "Order between $49–$74.99 to enjoy free shipping without Israeli customs fees.",
        "disclosure":       "Affiliate link — no extra cost to you, but we may earn a small commission.",
        "footer":           "Checked at: {checked_at} · Amazon Free Shipping Monitor",
        "aod_note":         "⚠️ Free shipping found in <strong>All Buying Options</strong>.<br>"
                            "Open the product page &rarr; click <strong>&ldquo;See All Buying Options&rdquo;</strong>"
                            " &rarr; select the offer with free shipping.",
        "aod_plain":        "NOTE: Found in All Buying Options — open the link, "
                            "click 'See All Buying Options', select the free-shipping offer.",
        "plain_header":     "🚨 FREE Shipping to Israel Alert!\n",
        "plain_product":    "Product",
        "plain_url":        "URL    ",
        "plain_urgency":    "⏰ Price may change at any time",
        "plain_footer":     "Checked at: {checked_at}",
    },
}


def _t(lang: str, key: str, **kw) -> str:
    s = _STRINGS.get(lang, _STRINGS["en"]).get(key, _STRINGS["en"].get(key, ""))
    return s.format(**kw) if kw else s


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

    items — list of dicts: [{"product": {...}, "shipping_text": "...", "found_in_aod": bool}, ...]
    """
    if not items:
        return

    lang          = config.get("language", "he")
    affiliate_tag = os.environ.get("AMAZON_AFFILIATE_TAG", "").strip()
    checked_at    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    is_rtl        = lang == "he"
    txt_dir       = 'dir="rtl"' if is_rtl else ""
    txt_align     = "right"     if is_rtl else "left"
    n             = len(items)

    # ── Subject ──────────────────────────────────────────────
    first_name = items[0]["product"].get("name", items[0]["product"].get("asin", ""))
    subject = (
        _t(lang, "subject_single", name=first_name) if n == 1
        else _t(lang, "subject_multi", n=n)
    )

    # ── Plain-text body ───────────────────────────────────────
    lines = [_t(lang, "plain_header")]
    for item in items:
        p    = item["product"]
        asin = p.get("asin", "")
        name = p.get("name", asin)
        product_url = (
            f"https://www.amazon.com/dp/{asin}?tag={affiliate_tag}"
            if affiliate_tag else
            f"https://www.amazon.com/dp/{asin}"
        )
        aod_line = [_t(lang, "aod_plain")] if item.get("found_in_aod") else []
        lines += [
            f"{_t(lang, 'plain_product')} : {name}",
            f"ASIN    : {asin}",
            f"{_t(lang, 'plain_url')} : {product_url}",
            _t(lang, "plain_urgency"),
            *aod_line,
            "",
        ]
    lines.append(_t(lang, "plain_footer", checked_at=checked_at))
    text_body = "\n".join(lines)

    # ── CTA button builder (bulletproof table-based) ──────────
    def _cta_btn(url: str, label: str, bg: str = "#e47911") -> str:
        return f"""
        <tr>
          <td align="{txt_align}" style="padding:8px 0;">
            <table cellpadding="0" cellspacing="0" border="0">
              <tr>
                <td align="center" bgcolor="{bg}"
                    style="border-radius:6px; mso-padding-alt:0;">
                  <a href="{url}"
                     style="display:inline-block; background:{bg}; color:#ffffff;
                            font-family:Arial,sans-serif; font-size:16px; font-weight:bold;
                            text-decoration:none; padding:13px 32px; border-radius:6px;
                            letter-spacing:0.3px; mso-padding-alt:13px 32px;"
                     target="_blank">
                    {label}
                  </a>
                </td>
              </tr>
            </table>
          </td>
        </tr>"""

    # ── HTML product cards ────────────────────────────────────
    product_cards = ""
    for item in items:
        p    = item["product"]
        asin = p.get("asin", "")
        name = p.get("name", asin)
        product_url = (
            f"https://www.amazon.com/dp/{asin}?tag={affiliate_tag}"
            if affiliate_tag else
            f"https://www.amazon.com/dp/{asin}"
        )

        aod_block = ""
        if item.get("found_in_aod"):
            aod_block = f"""
        <tr>
          <td style="padding-top:10px;">
            <div style="background:#fff8e1; border-{('right' if is_rtl else 'left')}:4px solid #e47911;
                        padding:10px 14px; border-radius:4px; font-size:12px; color:#666;
                        line-height:1.6; text-align:{txt_align};" {txt_dir}>
              {_t(lang, "aod_note")}
            </div>
          </td>
        </tr>"""

        product_cards += f"""
      <table width="100%" cellpadding="0" cellspacing="0"
             style="background:#ffffff; border:1px solid #e8e8e8; border-radius:10px;
                    margin-bottom:16px; box-shadow:0 2px 8px rgba(0,0,0,0.06);">

        <!-- Product name + ASIN -->
        <tr>
          <td style="padding:20px 22px 4px;">
            <p style="margin:0 0 4px; font-size:17px; font-weight:bold; line-height:1.4;
                      text-align:{txt_align};" {txt_dir}>
              <a href="{product_url}" style="color:#0066cc; text-decoration:none;">{name}</a>
            </p>
            <p style="margin:0 0 10px; font-size:12px; color:#999; text-align:{txt_align};">
              ASIN: {asin}
            </p>
          </td>
        </tr>

        <!-- Shipping badge -->
        <tr>
          <td style="padding:0 22px 10px;">
            <p style="margin:0; font-size:13px; font-weight:bold; color:#2e7d32;
                      text-align:{txt_align};" {txt_dir}>
              {_t(lang, "shipping_badge")}
            </p>
          </td>
        </tr>

        <!-- Primary CTA button -->
        <tr>
          <td style="padding:0 22px 4px;">
            <table width="100%" cellpadding="0" cellspacing="0">
              {_cta_btn(product_url, _t(lang, "btn_buy"))}
            </table>
          </td>
        </tr>

        <!-- Urgency line -->
        <tr>
          <td style="padding:4px 22px 10px; text-align:{txt_align};" {txt_dir}>
            <p style="margin:0; font-size:12px; color:#c0392b; font-style:italic;">
              {_t(lang, "urgency")}
            </p>
          </td>
        </tr>

        <!-- Secondary CTA button (repeated) -->
        <tr>
          <td style="padding:0 22px 16px;">
            <table width="100%" cellpadding="0" cellspacing="0">
              {_cta_btn(product_url, _t(lang, "btn_buy"), bg="#c0392b")}
            </table>
          </td>
        </tr>

        {aod_block}
        <tr><td style="padding:0 0 4px;"></td></tr>
      </table>"""

    # ── Header sub-text ───────────────────────────────────────
    header_sub = _t(lang, "header_sub1") if n == 1 else _t(lang, "header_sub", n=n)

    # ── Affiliate disclosure row (only when tag is set) ───────
    disclosure_row = ""
    if affiliate_tag:
        disclosure_row = f"""
          <tr>
            <td style="padding:0 22px 14px; text-align:{txt_align};" {txt_dir}>
              <p style="margin:0; font-size:11px; color:#aaa; font-style:italic;">
                {_t(lang, "disclosure")}
              </p>
            </td>
          </tr>"""

    # ── Quick tip box ─────────────────────────────────────────
    quick_tip_row = f"""
          <tr>
            <td style="padding:0 22px 16px;">
              <table width="100%" cellpadding="0" cellspacing="0"
                     style="background:#e8f5e9; border-radius:8px; border:1px solid #a5d6a7;">
                <tr>
                  <td style="padding:14px 18px; text-align:{txt_align};" {txt_dir}>
                    <p style="margin:0 0 4px; font-size:13px; font-weight:bold; color:#2e7d32;">
                      {_t(lang, "quick_tip_title")}
                    </p>
                    <p style="margin:0; font-size:12px; color:#388e3c; line-height:1.5;">
                      {_t(lang, "quick_tip_body")}
                    </p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>"""

    # ── Full HTML ─────────────────────────────────────────────
    html_body = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin:0; padding:0; background:#f0f2f5;
             font-family:Arial,'Segoe UI',sans-serif;">

  <!-- Hidden preheader (shows in inbox preview) -->
  <div style="display:none; max-height:0; overflow:hidden; mso-hide:all;">
    {_t(lang, "preheader")}
    &nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;
    &nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;
  </div>

  <table width="100%" cellpadding="0" cellspacing="0"
         style="background:#f0f2f5; padding:28px 0;">
    <tr>
      <td align="center">
        <table width="600" cellpadding="0" cellspacing="0"
               style="max-width:600px; width:100%;">

          <!-- Header -->
          <tr>
            <td style="background:linear-gradient(135deg, #c45e00 0%, #e47911 50%, #f5a623 100%);
                       border-radius:12px 12px 0 0; padding:32px 28px; text-align:center;">
              <div style="font-size:48px; line-height:1; margin-bottom:14px;">🚚</div>
              <h1 style="margin:0 0 8px; color:white; font-size:26px; font-weight:bold;
                         letter-spacing:0.5px;" {txt_dir}>
                {_t(lang, "header_title")}
              </h1>
              <p style="margin:0; color:rgba(255,255,255,0.88); font-size:14px;" {txt_dir}>
                {header_sub}
              </p>
            </td>
          </tr>

          <!-- Disclosure (if affiliate) -->
          {disclosure_row}

          <!-- Products -->
          <tr>
            <td style="background:#f8f9fa; padding:22px 22px 8px;">
              {product_cards}
            </td>
          </tr>

          <!-- Quick tip -->
          <tr>
            <td style="background:#f8f9fa; padding:0 22px 8px;">
              <table width="100%" cellpadding="0" cellspacing="0">
                {quick_tip_row}
              </table>
            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="background:#efefef; border-top:1px solid #e0e0e0;
                       border-radius:0 0 12px 12px;
                       padding:14px 28px; text-align:center;">
              <p style="margin:0; color:#aaa; font-size:11px;" {txt_dir}>
                {_t(lang, "footer", checked_at=checked_at)}
              </p>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>"""

    _smtp_send(config, subject, text_body, html_body)


def send_free_shipping_alert(config: dict, product: dict, shipping_text: str):
    """Single-product shortcut — delegates to the batch sender."""
    send_batch_free_shipping_alert(config, [{"product": product, "shipping_text": shipping_text}])
