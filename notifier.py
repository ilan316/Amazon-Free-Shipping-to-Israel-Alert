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

# Keep product names readable in subject/body when Amazon titles are very long.
_MAX_NAME_SUBJECT = 72
_MAX_NAME_BODY = 88

# ── Localized strings ─────────────────────────────────────────────────────────
_STRINGS = {
    "he": {
        "subject_single":   "פרסומת | 🚨 משלוח חינם לישראל: {name}",
        "subject_multi":    "פרסומת | 🚨 משלוח חינם לישראל: {n} מוצרים",
        "preheader":        "אל תפספס! המחיר עשוי להשתנות בכל עת — לחץ לפרטים",
        "header_title":     "משלוח חינם לישראל 🚚",
        "header_sub":       "נמצאו {n} מוצרים עם משלוח חינם",
        "header_sub1":      "נמצא מוצר עם משלוח חינם",
        "shipping_badge":   "✅ משלוח חינם לישראל · הזמנות $49+",
        "btn_buy":          "לקנות באמזון →",
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
        "header_title":     "FREE Shipping to Israel 🚚",
        "header_sub":       "{n} products with free shipping found",
        "header_sub1":      "1 product with free shipping found",
        "shipping_badge":   "✅ FREE Shipping to Israel · Orders $49+",
        "btn_buy":          "Buy on Amazon →",
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


def _short_product_name(name: str, limit: int = _MAX_NAME_BODY) -> str:
    """Trim long product names without cutting words in the middle when possible."""
    if not name:
        return ""
    clean = " ".join(str(name).split())
    if len(clean) <= limit:
        return clean
    head = clean[: max(1, limit - 1)]
    cut = head.rfind(" ")
    if cut >= int(limit * 0.6):
        head = head[:cut]
    return f"{head.rstrip()}…"


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
    first_name = _short_product_name(
        items[0]["product"].get("name", items[0]["product"].get("asin", "")),
        _MAX_NAME_SUBJECT
    )
    subject = (
        _t(lang, "subject_single", name=first_name) if n == 1
        else _t(lang, "subject_multi", n=n)
    )

    # ── Plain-text body ───────────────────────────────────────
    lines = [_t(lang, "plain_header")]
    for item in items:
        p    = item["product"]
        asin = p.get("asin", "")
        name = _short_product_name(p.get("name", asin), _MAX_NAME_BODY)
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
    def _cta_btn(url: str, label: str) -> str:
        return f"""<table cellpadding="0" cellspacing="0" border="0" style="margin:8px 0 4px;">
              <tr>
                <td align="center" bgcolor="#FF9900"
                    style="border-radius:6px;">
                  <a href="{url}"
                     style="display:inline-block; background:#FF9900; color:#111111;
                            font-family:Arial,sans-serif; font-size:14px; font-weight:bold;
                            text-decoration:none; padding:11px 28px; border-radius:6px;
                            letter-spacing:0.2px; white-space:nowrap;"
                     target="_blank">{label}</a>
                </td>
              </tr>
            </table>"""

    # ── HTML product cards ────────────────────────────────────
    # RTL: image on right, text on left → reverse column order
    product_cards = ""
    for item in items:
        p       = item["product"]
        asin    = p.get("asin", "")
        name    = _short_product_name(p.get("name", asin), _MAX_NAME_BODY)
        img_url = f"https://m.media-amazon.com/images/P/{asin}.01._SL200_.jpg"
        product_url = (
            f"https://www.amazon.com/dp/{asin}?tag={affiliate_tag}"
            if affiliate_tag else
            f"https://www.amazon.com/dp/{asin}"
        )

        aod_block = ""
        if item.get("found_in_aod"):
            aod_block = f"""<tr>
              <td colspan="3" style="padding:10px 0 0;">
                <div style="background:#fff8e1;
                            border-{('right' if is_rtl else 'left')}:3px solid #FF9900;
                            padding:10px 14px; border-radius:4px;
                            font-size:12px; color:#555; line-height:1.6;
                            text-align:{txt_align};" {txt_dir}>
                  {_t(lang, "aod_note")}
                </div>
              </td>
            </tr>"""

        # Image cell and content cell — order depends on RTL
        img_cell = f"""<td width="120" valign="top"
               style="padding:{('16px 0 16px 16px' if is_rtl else '16px 16px 16px 0')};">
            <a href="{product_url}" target="_blank">
              <img src="{img_url}" width="110" height="110" alt="{name}"
                   style="display:block; border-radius:6px;
                          border:1px solid #e8e8e8; object-fit:contain;">
            </a>
          </td>"""

        content_cell = f"""<td valign="top"
               style="padding:{('16px 16px 16px 0' if is_rtl else '16px 0 16px 16px')};">

            <!-- Product name -->
            <p style="margin:0 0 3px; font-size:15px; font-weight:bold;
                      line-height:1.4; text-align:{txt_align};" {txt_dir}>
              <a href="{product_url}"
                 style="color:#0066cc; text-decoration:none;">{name}</a>
            </p>

            <!-- ASIN -->
            <p style="margin:0 0 8px; font-size:11px; color:#aaa;
                      text-align:{txt_align};">ASIN: {asin}</p>

            <!-- Shipping badge -->
            <p style="margin:0 0 10px; font-size:12px; font-weight:bold;
                      color:#007600; text-align:{txt_align};" {txt_dir}>
              {_t(lang, "shipping_badge")}
            </p>

            <!-- Primary CTA -->
            <div style="text-align:{txt_align};">
              {_cta_btn(product_url, _t(lang, "btn_buy"))}
            </div>

            <!-- Urgency -->
            <p style="margin:6px 0 6px; font-size:11px; color:#888;
                      font-style:italic; text-align:{txt_align};" {txt_dir}>
              {_t(lang, "urgency")}
            </p>

            <!-- Secondary CTA -->
            <div style="text-align:{txt_align};">
              {_cta_btn(product_url, _t(lang, "btn_buy"))}
            </div>

          </td>"""

        # Columns order: RTL → content | spacer | image, LTR → image | spacer | content
        if is_rtl:
            cols = content_cell + '<td width="1" style="background:#f0f0f0;"></td>' + img_cell
        else:
            cols = img_cell + '<td width="1" style="background:#f0f0f0;"></td>' + content_cell

        product_cards += f"""
      <table width="100%" cellpadding="0" cellspacing="0"
             style="background:#ffffff; border:1px solid #e8e8e8; border-radius:10px;
                    margin-bottom:14px;">
        <tr>
          {cols}
        </tr>
        {aod_block}
      </table>"""

    # ── Header sub-text ───────────────────────────────────────
    header_sub = _t(lang, "header_sub1") if n == 1 else _t(lang, "header_sub", n=n)

    # ── Affiliate disclosure row (only when tag is set) ───────
    disclosure_row = ""
    if affiliate_tag:
        disclosure_row = f"""
          <tr>
            <td style="padding:12px 24px 4px; text-align:{txt_align};" {txt_dir}>
              <p style="margin:0; font-size:11px; color:#aaa; font-style:italic;">
                {_t(lang, "disclosure")}
              </p>
            </td>
          </tr>"""

    # ── Quick tip box ─────────────────────────────────────────
    quick_tip_row = f"""
          <tr>
            <td style="padding:0 20px 20px;">
              <table width="100%" cellpadding="0" cellspacing="0"
                     style="background:#f0faf0; border-radius:8px;
                            border:1px solid #c8e6c9;">
                <tr>
                  <td style="padding:12px 16px; text-align:{txt_align};" {txt_dir}>
                    <p style="margin:0 0 3px; font-size:13px; font-weight:bold;
                               color:#2e7d32;">
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
<body style="margin:0; padding:0; background:#f3f3f3;
             font-family:Arial,'Segoe UI',sans-serif;">

  <!-- Hidden preheader -->
  <div style="display:none; max-height:0; overflow:hidden; mso-hide:all;">
    {_t(lang, "preheader")}
    &nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;
    &nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;
  </div>

  <table width="100%" cellpadding="0" cellspacing="0"
         style="background:#f3f3f3; padding:24px 0;">
    <tr>
      <td align="center">
        <table width="600" cellpadding="0" cellspacing="0"
               style="max-width:600px; width:100%;">

          <!-- Header -->
          <tr>
            <td style="background:#232f3e; border-radius:10px 10px 0 0;
                       padding:28px 24px; text-align:center;">
              <h1 style="margin:0 0 6px; color:#FF9900; font-size:22px;
                         font-weight:bold; letter-spacing:0.3px;" {txt_dir}>
                {_t(lang, "header_title")}
              </h1>
              <p style="margin:0; color:rgba(255,255,255,0.7); font-size:13px;"
                 {txt_dir}>
                {header_sub}
              </p>
            </td>
          </tr>

          <!-- Disclosure -->
          {disclosure_row}

          <!-- Products -->
          <tr>
            <td style="background:#f8f8f8; padding:20px 20px 6px;">
              {product_cards}
            </td>
          </tr>

          <!-- Quick tip -->
          <tr>
            <td style="background:#f8f8f8;">
              <table width="100%" cellpadding="0" cellspacing="0">
                {quick_tip_row}
              </table>
            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="background:#232f3e; border-radius:0 0 10px 10px;
                       padding:12px 24px; text-align:center;">
              <p style="margin:0; color:rgba(255,255,255,0.45); font-size:11px;"
                 {txt_dir}>
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
