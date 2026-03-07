"""
Amazon Israel Free Ship Alert — Tkinter GUI
Run with: python gui.py
"""

import sys
import os
import webbrowser

# Always set CWD to the project folder so that config.json / state.json /
# browser_profile are found correctly — regardless of how the app was launched
# (double-click, VBS launcher, Windows autostart registry, PyInstaller exe, etc.)
if getattr(sys, "frozen", False):
    os.chdir(os.path.dirname(sys.executable))
else:
    _script_dir = os.path.dirname(os.path.abspath(__file__))
    if _script_dir:
        os.chdir(_script_dir)

import asyncio
import queue
import re
import subprocess
import threading
import urllib.request
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
from datetime import datetime, timedelta

from dotenv import load_dotenv

load_dotenv()

_tray_import_error = ""
try:
    import pystray
    from PIL import Image as _PILImage, ImageDraw as _PILDraw
    _TRAY_AVAILABLE = True
except Exception as _e:
    _TRAY_AVAILABLE = False
    _tray_import_error = repr(_e)

import config as cfg_module
import state as state_module
from version import __version__


def _run_check_all_products(config: dict, state: dict):
    # Delay checker/playwright import until a check actually runs.
    from checker import check_all_products as _check_all_products
    return _check_all_products(config, state)


def _send_email_alert(config: dict, items: list):
    # Delay notifier import until we really send an email.
    from notifier import send_batch_free_shipping_alert as _send_batch_alert
    return _send_batch_alert(config, items)

# ──────────────────────────────────────────────
# Internationalisation (i18n)
# ──────────────────────────────────────────────

_STRINGS: dict = {
    "he": {
        "monitored_products":   "מוצרים במעקב",
        "col_product_name":     "שם מוצר",
        "col_asin":             "ASIN",
        "col_eligible":         "זכאות",
        "col_last_checked":     "בדיקה אחרונה",
        "btn_add_product":      "➕  הוסף מוצר",
        "btn_remove":           "🗑  הסר",
        "btn_pause":            "⏸  השהה",
        "btn_resume":           "▶  המשך",
        "btn_check_now":        "🔍  בדוק עכשיו",
        "btn_settings":         "⚙  הגדרות",
        "btn_contact":          "✉  צור קשר",
        "btn_start_monitoring": "▶  התחל ניטור",
        "btn_stop_monitoring":  "⏹  עצור ניטור",
        "log_label":            "יומן",
        "btn_clear":            "נקה",
        "status_ready":         "מוכן",
        "status_free":          "✅ זכאי — משלוח חינם",
        "status_not_eligible":  "❌ לא זכאי",
        "status_not_checked":   "— טרם נבדק",
        "status_paused":        "⏸  מושהה",
        "settings_title":       "הגדרות",
        "email_section":        'התראות דוא"ל',
        "email_label":          'כתובת הדוא"ל שלך:\u200f',
        "email_hint":           "התראות יישלחו לכתובת זו כשיזוהה משלוח חינם",
        "interval_section":     "מרווח בדיקה",
        "hours":                "שעות",
        "interval_6h":          "6 שעות",
        "interval_12h":         "12 שעות",
        "interval_24h":         "24 שעות",
        "interval_note":        "מייל יישלח עבור מוצרים שעברו למצב FREE SHIPPING, לכל היותר פעם אחת בכל 24 שעות, בהתאם למועד הבדיקה האחרונה.",
        "startup_section":      "הפעלה עם Windows",
        "startup_checkbox":     "הפעל אוטומטית עם הפעלת Windows",
        "btn_save":             "שמור",
        "language_section":     "שפה",
        "language_label":       "שפת ממשק:",
        "language_restart_note":"השינוי יכנס לתוקף עם הפתיחה הבאה",
        "add_title":            "הוסף מוצר/ים",
        "add_instruction":      "הזן כתובת מוצר או מספר מוצר (ASIN) — שורה אחת לכל מוצר:",
        "add_example":          "דוגמה: https://www.amazon.com/dp/B0CKH5GJFN או B0CKH5GJFN",
        "btn_add":              "הוסף",
        "btn_paste":            "הדבק",
        "check_now_title":      "בדיקה עכשיו",
        "check_now_prompt":     "האם לבצע בדיקה עכשיו?",
        "invalid_entries":      "ערכים לא תקינים",
        "no_selection":         "אין בחירה",
        "select_to_pause":      "בחר מוצר להשהיה/חידוש.",
        "select_to_remove":     "בחר מוצר להסרה.",
        "remove_title":         "הסר מוצר",
        "remove_confirm":       "להסיר ASIN {} מהניטור?",
        "no_products_title":    "אין מוצרים",
        "no_products_msg":      "הוסף לפחות מוצר אחד קודם.",
        "email_required_title": 'נדרש דוא"ל',
        "email_required_msg":   'אנא הזן כתובת דוא"ל לפני הרצת בדיקה.\n\nלחץ OK לפתיחת ההגדרות.',
        "busy_title":           "עסוק",
        "busy_msg":             "בדיקה כבר פועלת.",
        "email_not_conf_title": 'דוא"ל לא מוגדר',
        "email_not_conf_msg":   'לא הוגדרה כתובת דוא"ל.\n\nהתראות לא יישלחו עד שתגדיר כתובת.\n\nלחץ ⚙ הגדרות והזן כתובת דוא"ל.',
        "invalid_number":       "הזן מספרים תקינים.",
        "interval_too_short":   "המרווח חייב להיות לפחות דקה אחת.",
        "quit_title":           "יציאה",
        "quit_msg":             "הניטור פעיל. לעצור ולצאת?",
        "autostart_error":      "שגיאת הפעלה אוטומטית",
        "monitoring_every":     "ניטור — כל {} דק'",
        "monitoring_stopped":   "הניטור הופסק.",
        "check_running":        "מריץ בדיקה...",
        "check_complete":       "בדיקה הושלמה.",
        "free_detected":        "זוהה משלוח חינם!",
        "tray_open":            "פתח Amazon Israel Free Ship Alert",
        "tray_exit":            "יציאה",
        "log_added_products":   "נוספו {} מוצר/ים. לחץ 'בדוק עכשיו' לבדיקת משלוח.",
        "log_added_status":     "נוספו {} מוצר/ים.",
        "log_paused_action":    "הושהה: {}",
        "log_resumed_action":   "חודש: {}",
        "log_removed_action":   "מוצר הוסר: {}",
        "settings_saved":       "הגדרות נשמרו. מרווח: {} שעות.{}{}",
        "autostart_on_log":     "  הפעלה אוטומטית: פועל.",
        "autostart_off_log":    "  הפעלה אוטומטית: כבוי.",
        "email_set_log":        '  דוא"ל: {}.',
        "update_available_title": "עדכון זמין",
        "update_available_msg":   "גרסה {} זמינה (גרסה נוכחית: {}).\u200f\nלהוריד ולהתקין את העדכון?\u200f",
        "btn_update_now":         "הורד עדכון כעת",
        "btn_update_later":       "מאוחר יותר",
        "downloading_update":     "מוריד עדכון...",
        "update_failed":          "הורדת העדכון נכשלה: {}",
        "btn_copy_link":          "🔗  העתק קישור אפיליאט",
        "log_link_copied":        "קישור אפיליאט הועתק ללוח",
        "log_no_affiliate_tag":   "קוד אפיליאט לא מוגדר",
        "next_check_label":       "בדיקה הבאה:",
        "btn_log_show":           "▼ יומן",
        "btn_log_hide":           "▲ יומן",
    },
    "en": {
        "monitored_products":   "Monitored Products",
        "col_product_name":     "Product Name",
        "col_asin":             "ASIN",
        "col_eligible":         "Eligible",
        "col_last_checked":     "Last Checked",
        "btn_add_product":      "➕  Add Product",
        "btn_remove":           "🗑  Remove",
        "btn_pause":            "⏸  Pause",
        "btn_resume":           "▶  Resume",
        "btn_check_now":        "🔍  Check Now",
        "btn_settings":         "⚙  Settings",
        "btn_contact":          "✉  Contact",
        "btn_start_monitoring": "▶  Start Monitoring",
        "btn_stop_monitoring":  "⏹  Stop Monitoring",
        "log_label":            "Log",
        "btn_clear":            "Clear",
        "status_ready":         "Ready",
        "status_free":          "✅ Eligible — FREE Shipping",
        "status_not_eligible":  "❌ Not eligible",
        "status_not_checked":   "— Not checked yet",
        "status_paused":        "⏸  Paused",
        "settings_title":       "Settings",
        "email_section":        "Email alerts",
        "email_label":          "Your email address:",
        "email_hint":           "Alerts will be sent to this address when FREE shipping is detected",
        "interval_section":     "Check interval",
        "hours":                "Hours",
        "interval_6h":          "6 hours",
        "interval_12h":         "12 hours",
        "interval_24h":         "24 hours",
        "interval_note":        "An email will be sent for products that switched to FREE SHIPPING, at most once every 24 hours, based on the last check time.",
        "startup_section":      "Windows startup",
        "startup_checkbox":     "Start automatically when Windows boots",
        "btn_save":             "Save",
        "language_section":     "Language",
        "language_label":       "Interface language:",
        "language_restart_note":"Change takes effect on next app launch",
        "add_title":            "Add Product(s)",
        "add_instruction":      "Enter product URL or ASIN — one per line:",
        "add_example":          "Example: https://www.amazon.com/dp/B0CKH5GJFN or B0CKH5GJFN",
        "btn_add":              "Add",
        "btn_paste":            "Paste",
        "check_now_title":      "Check Now",
        "check_now_prompt":     "Run a check now?",
        "invalid_entries":      "Invalid entries",
        "no_selection":         "No selection",
        "select_to_pause":      "Select a product to pause/resume.",
        "select_to_remove":     "Select a product to remove.",
        "remove_title":         "Remove Product",
        "remove_confirm":       "Remove ASIN {} from monitoring?",
        "no_products_title":    "No Products",
        "no_products_msg":      "Add at least one product first.",
        "email_required_title": "Email required",
        "email_required_msg":   "Please enter your email address before running a check.\n\nClick OK to open Settings.",
        "busy_title":           "Busy",
        "busy_msg":             "A check is already running.",
        "email_not_conf_title": "Email not configured",
        "email_not_conf_msg":   "No recipient email address is set.\n\nAlerts will not be sent until you configure your email.\n\nPlease click \u2699 Settings and enter your email address.",
        "invalid_number":       "Enter valid numbers.",
        "interval_too_short":   "Interval must be at least 1 minute.",
        "quit_title":           "Quit",
        "quit_msg":             "Monitoring is running. Stop it and exit?",
        "autostart_error":      "Autostart Error",
        "monitoring_every":     "Monitoring — every {} min",
        "monitoring_stopped":   "Monitoring stopped.",
        "check_running":        "Running check...",
        "check_complete":       "Check complete.",
        "free_detected":        "FREE shipping detected!",
        "tray_open":            "Open Amazon Israel Free Ship Alert",
        "tray_exit":            "Exit",
        "log_added_products":   "Added {} product(s). Click Check Now to fetch names and check shipping.",
        "log_added_status":     "Added {} product(s).",
        "log_paused_action":    "Paused: {}",
        "log_resumed_action":   "Resumed: {}",
        "log_removed_action":   "Removed product: {}",
        "settings_saved":       "Settings saved. Interval: {} hours.{}{}",
        "autostart_on_log":     "  Auto-start: ON.",
        "autostart_off_log":    "  Auto-start: OFF.",
        "email_set_log":        "  Email: {}.",
        "update_available_title": "Update Available",
        "update_available_msg":   "Version {} is available (current: {}).\nDownload and install the update?",
        "btn_update_now":         "Download Update Now",
        "btn_update_later":       "Later",
        "downloading_update":     "Downloading update...",
        "update_failed":          "Update download failed: {}",
        "btn_copy_link":          "🔗  Copy Affiliate Link",
        "log_link_copied":        "Affiliate link copied to clipboard",
        "log_no_affiliate_tag":   "Affiliate tag not configured",
        "next_check_label":       "Next check:",
        "btn_log_show":           "▼ Log",
        "btn_log_hide":           "▲ Log",
    },
}

try:
    _LANG: str = cfg_module.load_config().get("language", "he")
except Exception:
    _LANG = "he"


def _t(key: str, *args) -> str:
    """Return translated string for the current language, falling back to English."""
    s = _STRINGS.get(_LANG, _STRINGS["en"]).get(key) or _STRINGS["en"].get(key, key)
    return s.format(*args) if args else s


# RTL layout helpers — used throughout to right-align text in Hebrew mode
_IS_RTL  = _LANG == "he"
_ANCHOR  = "e" if _IS_RTL else "w"     # widget/text anchor (east = right)
_JUSTIFY = "right" if _IS_RTL else "left"  # text justification inside widgets


# ──────────────────────────────────────────────
# Status display helpers
# ──────────────────────────────────────────────

STATUS_LABELS = {
    "FREE":    _t("status_free"),
    "PAID":    _t("status_not_eligible"),
    "NO_SHIP": _t("status_not_eligible"),
    "UNKNOWN": _t("status_not_eligible"),
    "ERROR":   _t("status_not_eligible"),
    "—":       _t("status_not_checked"),
    "PAUSED":  _t("status_paused"),
}

STATUS_COLORS = {
    "FREE":     "#1a7a1a",
    "PAID":     "#cc2200",
    "NO_SHIP":  "#cc2200",
    "UNKNOWN":  "#cc2200",
    "ERROR":    "#cc2200",
    "—":        "#aaaaaa",
    "PAUSED":   "#888888",
}

MAX_LOG_LINES = 300


# ──────────────────────────────────────────────
# Product name fetcher (no browser needed)
# ──────────────────────────────────────────────

def _fetch_product_name(asin: str) -> str:
    """
    Tries to fetch the product title from Amazon using a plain HTTP request.
    Returns the title string, or the ASIN if the request fails/is blocked.
    """
    import html as html_mod
    url = f"https://www.amazon.com/dp/{asin}"
    req = urllib.request.Request(url, headers={
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept-Language": "en-US,en;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    })
    try:
        with urllib.request.urlopen(req, timeout=12) as resp:
            html_content = resp.read().decode("utf-8", errors="ignore")

        # Prefer og:title — try both attribute orderings (Amazon varies)
        for pat in [
            r'<meta\b[^>]*\bproperty=["\']og:title["\'][^>]*\bcontent=["\'](.*?)["\']',
            r'<meta\b[^>]*\bcontent=["\'](.*?)["\'][^>]*\bproperty=["\']og:title["\']',
        ]:
            m = re.search(pat, html_content, re.IGNORECASE)
            if m and m.group(1).strip():
                return html_mod.unescape(m.group(1).strip())

        # Fall back to <title> tag — skip parts that are just "Amazon" or "Amazon.com"
        title_match = re.search(r'<title[^>]*>(.*?)</title>', html_content,
                                re.IGNORECASE | re.DOTALL)
        if title_match:
            title = html_mod.unescape(title_match.group(1).strip())
            for part in title.split(":"):
                part = part.strip()
                if part and part.lower() not in ("amazon", "amazon.com", ""):
                    return part
    except Exception:
        pass
    return asin  # fallback: use ASIN as name


# ──────────────────────────────────────────────
# Background monitor thread
# ──────────────────────────────────────────────

class MonitorThread(threading.Thread):
    def __init__(self, log_queue: queue.Queue, on_results, stop_event: threading.Event):
        super().__init__(daemon=True)
        self.log_queue = log_queue
        self.on_results = on_results
        self.stop_event = stop_event

    def _log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put(f"[{ts}] {msg}")

    def _compute_first_wait(self) -> int:
        """Return seconds to wait before the first check (0 = check immediately)."""
        try:
            config       = cfg_module.load_config()
            state        = state_module.load_state()
            interval_sec = config.get("check_interval_minutes", 60) * 60
            products     = [p for p in config.get("products", [])
                            if not p.get("paused", False)]
            if not products or not state:
                return 0
            latest = None
            for p in products:
                lc = state.get(p["asin"], {}).get("last_checked")
                if not lc:
                    return 0  # unchecked product — check immediately
                try:
                    dt = datetime.fromisoformat(lc)
                    if latest is None or dt > latest:
                        latest = dt
                except ValueError:
                    return 0
            if latest is None:
                return 0
            elapsed = (datetime.now() - latest).total_seconds()
            return max(0, int(interval_sec - elapsed))
        except Exception:
            return 0

    def run(self):
        first_wait = self._compute_first_wait()
        if first_wait > 0:
            h = first_wait // 3600
            m = (first_wait % 3600) // 60
            s = first_wait % 60
            wait_str = f"{h}h {m}m" if h else (f"{m}m {s}s" if m else f"{s}s")
            self._log(f"Smart resume: last check was recent — next check in {wait_str}.")
            next_run = datetime.now() + timedelta(seconds=first_wait)
            self.log_queue.put(f"__next_run__{next_run.strftime('%H:%M:%S')}")
            self.stop_event.wait(timeout=first_wait)
            if self.stop_event.is_set():
                self._log("Monitoring stopped.")
                return

        while not self.stop_event.is_set():
            config = cfg_module.load_config()
            state  = state_module.load_state()
            products = config.get("products", [])

            active   = [p for p in products if not p.get("paused", False)]
            paused_n = len(products) - len(active)

            if not products:
                self._log("No products configured — waiting...")
            elif not active:
                self._log("All products are paused — skipping check cycle.")
            else:
                skip_msg = f"  ({paused_n} paused)" if paused_n else ""
                self._log(f"Starting check cycle ({len(active)} product(s){skip_msg})...")
                self._log("  Setting delivery location to Israel...")
                check_config = {**config, "products": active}
                try:
                    results = asyncio.run(_run_check_all_products(check_config, state))
                    product_map = {p["asin"]: p for p in products}

                    free_items = []
                    names_updated = False
                    for result in results:
                        asin       = result.asin
                        status_str = result.status.value
                        product    = product_map.get(asin, {"asin": asin, "name": asin, "url": ""})
                        name       = product.get("name", asin)

                        # Auto-update name if checker fetched a real title
                        if result.product_name and result.product_name != asin \
                                and product.get("name", asin) == asin:
                            product["name"] = result.product_name
                            name = result.product_name
                            names_updated = True

                        cooldown = config.get("notification_cooldown_hours", 24)
                        notify = state_module.should_notify(state, asin, status_str, cooldown_hours=cooldown)
                        if notify:
                            free_items.append({"product": product, "shipping_text": result.raw_text,
                                               "found_in_aod": result.found_in_aod})

                        state = state_module.update_product_state(
                            state, asin, status_str, notified=notify,
                        )
                        label = STATUS_LABELS.get(status_str, status_str)
                        self._log(f"  {name}: {label}")

                    if names_updated:
                        cfg_module.save_config(config)

                    if free_items:
                        names = ", ".join(
                            i["product"].get("name", i["product"].get("asin", ""))
                            for i in free_items
                        )
                        self._log(f"FREE shipping for: {names} — sending email...")
                        try:
                            _send_email_alert(config, free_items)
                            self._log("Email alert sent successfully.")
                        except RuntimeError as mail_err:
                            self._log(f"EMAIL ERROR: {mail_err}")

                    state_module.save_state(state)
                    self._log("Check cycle complete.")
                    self.on_results()

                except Exception as e:
                    self._log(f"Error during check: {e}")

            interval = cfg_module.load_config().get("check_interval_minutes", 60)
            next_run = datetime.now() + timedelta(minutes=interval)
            self.log_queue.put(f"__next_run__{next_run.strftime('%H:%M:%S')}")
            self.stop_event.wait(timeout=interval * 60)

        self._log("Monitoring stopped.")


# ──────────────────────────────────────────────
# Main GUI class
# ──────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Amazon Israel Free Ship Alert  v{__version__}")
        self.resizable(True, True)
        self.minsize(740, 520)
        try:
            self.iconbitmap(os.path.join(os.getcwd(), "icon.ico"))
        except Exception:
            pass
        # Force taskbar icon via Windows API once the window is visible
        self.after(300, self._apply_taskbar_icon)

        self._monitor_thread: MonitorThread | None = None
        self._stop_event = threading.Event()
        self._log_queue: queue.Queue = queue.Queue()
        self._manual_check_running = False
        self._tray_icon = None
        self._tray_thread: threading.Thread | None = None
        self._sort_col: str | None = None
        self._sort_rev: bool = False
        self._checked: dict = {}  # asin → bool, True = checked (default)

        self._build_ui()
        self._refresh_table()
        self._poll_log()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        # Defer tray setup so the main window can appear faster.
        self.after(1000, self._init_tray)
        self._sync_autostart()                          # upgrade registry to VBS if needed
        self.after(800, self._maybe_autostart_monitoring)  # resume monitoring if it was running
        self.after(2000, self._check_for_updates)       # check for newer version on startup

    # ── UI construction ──────────────────────

    def _apply_taskbar_icon(self):
        """Force the Windows taskbar to use our icon via direct Win32 API calls."""
        try:
            import ctypes, ctypes.wintypes
            icon_path = os.path.join(os.getcwd(), "icon.ico")
            if not os.path.exists(icon_path):
                return

            u32 = ctypes.windll.user32

            # Declare proper 64-bit-safe types
            u32.LoadImageW.restype = ctypes.c_void_p
            u32.LoadImageW.argtypes = [
                ctypes.c_void_p, ctypes.c_wchar_p,
                ctypes.c_uint, ctypes.c_int, ctypes.c_int, ctypes.c_uint
            ]
            u32.SendMessageW.restype = ctypes.c_ssize_t
            u32.SendMessageW.argtypes = [
                ctypes.c_ssize_t, ctypes.c_uint,
                ctypes.c_size_t, ctypes.c_size_t
            ]
            u32.SetClassLongPtrW.restype = ctypes.c_size_t
            u32.SetClassLongPtrW.argtypes = [
                ctypes.c_ssize_t, ctypes.c_int, ctypes.c_size_t
            ]

            hIcon = u32.LoadImageW(
                None, icon_path,
                1,           # IMAGE_ICON
                0, 0,        # use default system sizes
                0x10 | 0x40  # LR_LOADFROMFILE | LR_DEFAULTSIZE
            )
            if not hIcon:
                return

            hwnd = self.winfo_id()
            WM_SETICON = 0x0080
            # Set instance icon (title bar + taskbar)
            u32.SendMessageW(hwnd, WM_SETICON, 1, hIcon)  # ICON_BIG
            u32.SendMessageW(hwnd, WM_SETICON, 0, hIcon)  # ICON_SMALL
            # Set class icon (overrides Python's default for this window class)
            u32.SetClassLongPtrW(hwnd, -14, hIcon)  # GCL_HICON
            u32.SetClassLongPtrW(hwnd, -34, hIcon)  # GCL_HICONSM
        except Exception:
            pass

    def _build_ui(self):
        # Logo-matched color palette
        # Orange: #F5A31A  |  Navy: #1E3252  |  Cream bg: #FAF6EA
        _BG = "#FAF6EA"
        self.configure(bg=_BG)
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading",
                         background="#1E3252", foreground="white",
                         font=("Segoe UI", 9, "bold"), relief="flat")
        style.map("Treeview.Heading", background=[("active", "#2A4470")])
        style.configure("Treeview",
                         background="#ffffff", fieldbackground="#ffffff",
                         rowheight=26, font=("Segoe UI", 9))
        style.map("Treeview",
                   background=[("selected", "#FEF3DC")],
                   foreground=[("selected", "#1E3252")])

        # Products table
        top = tk.Frame(self, padx=8, pady=6, bg=_BG)
        top.pack(fill=tk.BOTH, expand=True)

        # Logo banner above the table
        self._logo_img = None
        try:
            from PIL import Image as _PilImg, ImageTk as _PilImgTk
            _logo_path = os.path.join(os.getcwd(), "logo-new.png")
            if os.path.exists(_logo_path):
                _pil = _PilImg.open(_logo_path).convert("RGBA")
                _pil = _pil.resize((380, 157), _PilImg.LANCZOS)
                self._logo_img = _PilImgTk.PhotoImage(_pil)
                _logo_frame = tk.Frame(top, bg=_BG)
                _logo_frame.pack(anchor="center", pady=(0, 4))
                tk.Label(_logo_frame, image=self._logo_img, bg=_BG).pack()
        except Exception:
            pass

        self._products_label_var = tk.StringVar(value=_t("monitored_products"))
        tk.Label(top, textvariable=self._products_label_var,
                 font=("Segoe UI", 11, "bold"),
                 anchor=_ANCHOR, justify=_JUSTIFY, bg=_BG).pack(anchor=_ANCHOR, fill=tk.X)

        cols = ("check", "name", "asin", "status", "last_checked")
        self.tree = ttk.Treeview(top, columns=cols, show="headings", selectmode="extended", height=8)
        self.tree.heading("check", text="☑", command=self._toggle_all_checked)
        for _col, _txt in [
            ("name",         _t("col_product_name")),
            ("asin",         _t("col_asin")),
            ("status",       _t("col_eligible")),
            ("last_checked", _t("col_last_checked")),
        ]:
            self.tree.heading(_col, text=_txt, command=lambda c=_col: self._sort_by(c))
        self.tree.column("check",        width=30,  anchor="center", stretch=False)
        self.tree.column("name",         width=260, anchor=_ANCHOR)
        self.tree.column("asin",         width=110, anchor="center")
        self.tree.column("status",       width=175, anchor="center")
        self.tree.column("last_checked", width=140, anchor=_ANCHOR)

        vsb = ttk.Scrollbar(top, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.LEFT, fill=tk.Y)

        self.tree.bind("<Double-1>", self._on_product_double_click)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)

        for status, colour in STATUS_COLORS.items():
            self.tree.tag_configure(status, foreground=colour)

        # Action buttons
        mid = tk.Frame(self, padx=8, pady=4, bg=_BG)
        mid.pack(fill=tk.X)

        btn_cfg = {"padx": 9, "pady": 4, "relief": tk.FLAT, "bd": 0,
                   "font": ("Segoe UI", 9), "cursor": "hand2"}

        tk.Button(mid, text=_t("btn_add_product"),
                  bg="#F5A31A", fg="white", activebackground="#D9901A",
                  command=self._add_product, **btn_cfg).pack(side=tk.LEFT, padx=(0, 4))

        tk.Button(mid, text=_t("btn_remove"),
                  bg="#dc2626", fg="white", activebackground="#b91c1c",
                  command=self._remove_product, **btn_cfg).pack(side=tk.LEFT, padx=(0, 4))

        self._pause_btn = tk.Button(mid, text=_t("btn_pause"),
                                    bg="#d97706", fg="white", activebackground="#b45309",
                                    command=self._toggle_pause, **btn_cfg)
        self._pause_btn.pack(side=tk.LEFT, padx=(0, 4))

        tk.Button(mid, text=_t("btn_check_now"),
                  bg="#1E3252", fg="white", activebackground="#16284A",
                  command=self._check_now, **btn_cfg).pack(side=tk.LEFT, padx=(0, 4))

        tk.Button(mid, text=_t("btn_settings"),
                  bg="#6B5E45", fg="white", activebackground="#574A38",
                  command=self._show_settings, **btn_cfg).pack(side=tk.LEFT, padx=(0, 4))

        tk.Button(mid, text=_t("btn_contact"),
                  bg="#1E3252", fg="white", activebackground="#16284A",
                  command=lambda: webbrowser.open("https://www.amzfreeil.com/"),
                  **btn_cfg).pack(side=tk.LEFT, padx=(0, 16))

        self._start_btn = tk.Button(mid, text=_t("btn_start_monitoring"),
                                    bg="#F5A31A", fg="white", activebackground="#D9901A",
                                    command=self._toggle_monitoring, **btn_cfg)
        self._start_btn.pack(side=tk.LEFT, padx=(0, 4))

        self._interval_var = tk.StringVar(value="")
        tk.Label(mid, textvariable=self._interval_var,
                 font=("Segoe UI", 9), fg="#8B7355", bg=_BG).pack(side=tk.RIGHT)

        # Log area
        bot = tk.Frame(self, padx=8, bg=_BG)
        bot.pack(fill=tk.BOTH, expand=False, pady=(0, 8))

        self._log_visible = False

        hdr = tk.Frame(bot, bg=_BG)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text=_t("log_label"), font=("Segoe UI", 9, "bold"),
                 anchor=_ANCHOR, bg=_BG).pack(side=tk.LEFT if not _IS_RTL else tk.RIGHT)
        self._toggle_log_btn = tk.Button(
            hdr, text=_t("btn_log_show"), font=("Segoe UI", 8), relief=tk.FLAT,
            command=self._toggle_log, cursor="hand2", bg=_BG)
        self._toggle_log_btn.pack(side=tk.LEFT if _IS_RTL else tk.RIGHT, padx=(0, 4))
        tk.Button(hdr, text=_t("btn_clear"), font=("Segoe UI", 8), relief=tk.FLAT,
                  command=self._clear_log, cursor="hand2", bg=_BG).pack(side=tk.LEFT if _IS_RTL else tk.RIGHT)

        self._log_container = tk.Frame(bot, bg=_BG)
        # intentionally NOT packed — log is hidden by default

        self._log_text = tk.Text(
            self._log_container, height=10, state=tk.DISABLED,
            font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
            relief=tk.FLAT, padx=6, pady=4, wrap=tk.WORD,
        )
        lsb = ttk.Scrollbar(self._log_container, orient="vertical", command=self._log_text.yview)
        self._log_text.configure(yscrollcommand=lsb.set)
        self._log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        lsb.pack(side=tk.LEFT, fill=tk.Y)

        self._log_text.tag_configure("free",  foreground="#4ec94e")
        self._log_text.tag_configure("error", foreground="#f48771")
        self._log_text.tag_configure("info",  foreground="#d4d4d4")

        # Status bar
        self._status_var = tk.StringVar(value=_t("status_ready"))
        tk.Label(self, textvariable=self._status_var, relief=tk.FLAT,
                 anchor=_ANCHOR, font=("Segoe UI", 8), fg="#7A6535", bg="#F0E5C8",
                 padx=8, pady=3).pack(side=tk.BOTTOM, fill=tk.X)

    # ── Table helpers ─────────────────────────

    def _sort_by(self, col: str):
        if self._sort_col == col:
            self._sort_rev = not self._sort_rev
        else:
            self._sort_col = col
            self._sort_rev = False
        self._refresh_table()

    def _refresh_table(self):
        config = cfg_module.load_config()
        state  = state_module.load_state()

        # Update column headers with sort arrows
        _base = {
            "name":         _t("col_product_name"),
            "asin":         _t("col_asin"),
            "status":       _t("col_eligible"),
            "last_checked": _t("col_last_checked"),
        }
        for col, base in _base.items():
            arrow = (" ▼" if self._sort_rev else " ▲") if self._sort_col == col else ""
            self.tree.heading(col, text=base + arrow)

        # Sync _checked: add new products (default True), keep existing state
        current_asins = {p["asin"] for p in config.get("products", [])}
        for asin in current_asins:
            if asin not in self._checked:
                self._checked[asin] = True
        # Remove stale entries
        for asin in list(self._checked):
            if asin not in current_asins:
                del self._checked[asin]

        # Build row data
        rows = []
        for p in config.get("products", []):
            asin     = p["asin"]
            paused   = p.get("paused", False)
            s        = state.get(asin, {})
            raw_last = s.get("last_checked") or ""
            if raw_last:
                try:
                    _d = datetime.fromisoformat(raw_last)
                    disp_last = _d.strftime("%d/%m/%Y %H:%M:%S") if _IS_RTL else _d.strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    disp_last = raw_last.replace("T", " ")
            else:
                disp_last = "—"
            if paused:
                label, tag = STATUS_LABELS["PAUSED"], "PAUSED"
            else:
                sk    = s.get("last_status", "—")
                label = STATUS_LABELS.get(sk, sk)
                tag   = sk if sk in STATUS_COLORS else "—"
            rows.append({
                "asin": asin, "name": p.get("name", asin),
                "label": label, "disp_last": disp_last,
                "sort_last": raw_last, "tag": tag,
            })

        # Sort rows if a column is selected
        if self._sort_col:
            key_fn = {
                "name":         lambda r: r["name"].lower(),
                "asin":         lambda r: r["asin"].lower(),
                "status":       lambda r: r["label"].lower(),
                "last_checked": lambda r: r["sort_last"],
            }.get(self._sort_col, lambda r: "")
            rows.sort(key=key_fn, reverse=self._sort_rev)

        # Update heading with count
        self._products_label_var.set(f"{_t('monitored_products')} ({len(rows)})")

        # Repopulate tree
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for r in rows:
            chk_sym = "☑" if self._checked.get(r["asin"], True) else "☐"
            self.tree.insert("", tk.END, iid=r["asin"],
                             values=(chk_sym, r["name"], r["asin"], r["label"], r["disp_last"]),
                             tags=(r["tag"],))
        self._on_tree_select()

    def _on_product_double_click(self, event):
        """Double-clicking a product row opens its affiliate link (or regular URL) in the browser."""
        import webbrowser
        sel = self.tree.selection()
        if not sel:
            return
        asin = sel[0]
        tag = os.environ.get("AMAZON_AFFILIATE_TAG", "").strip()
        if tag:
            webbrowser.open(f"https://www.amazon.com/dp/{asin}?tag={tag}")
        else:
            config = cfg_module.load_config()
            for p in config.get("products", []):
                if p["asin"] == asin:
                    webbrowser.open(p.get("url", f"https://www.amazon.com/dp/{asin}"))
                    return

    def _on_tree_select(self, event=None):
        """Update Pause button label based on selected product's paused state."""
        if not hasattr(self, "_pause_btn"):
            return
        sel = self.tree.selection()
        if not sel:
            self._pause_btn.configure(text=_t("btn_pause"))
            return
        asin = sel[0]
        config = cfg_module.load_config()
        for p in config.get("products", []):
            if p["asin"] == asin:
                if p.get("paused", False):
                    self._pause_btn.configure(text=_t("btn_resume"))
                else:
                    self._pause_btn.configure(text=_t("btn_pause"))
                return

    def _on_tree_click(self, event):
        """Toggle checkbox when user clicks the check column cell."""
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        if col != "#1":  # check column is first (#1)
            return
        row = self.tree.identify_row(event.y)
        if not row:
            return
        self._checked[row] = not self._checked.get(row, True)
        self.tree.set(row, "check", "☑" if self._checked[row] else "☐")

    def _toggle_all_checked(self):
        """Select All / Deselect All via the check column heading."""
        if not self._checked:
            return
        all_checked = all(self._checked.values())
        new_state = not all_checked
        for asin in self._checked:
            self._checked[asin] = new_state
        self._refresh_table()

    def _checked_asins(self) -> list:
        """Return list of checked ASINs, or all ASINs if none are checked."""
        checked = [asin for asin, v in self._checked.items() if v]
        return checked if checked else list(self._checked)

    def _toggle_pause(self):
        targets = self._checked_asins()
        if not targets:
            messagebox.showinfo(_t("no_selection"), _t("select_to_pause"), parent=self)
            return
        config = cfg_module.load_config()
        actions = []
        for p in config.get("products", []):
            if p["asin"] in targets:
                p["paused"] = not p.get("paused", False)
                action = _t("log_paused_action", p.get("name", p["asin"])) if p["paused"] \
                         else _t("log_resumed_action", p.get("name", p["asin"]))
                actions.append(action)
        cfg_module.save_config(config)
        self._refresh_table()
        for action in actions:
            self._append_log(action)

    def _maybe_autostart_monitoring(self):
        """Auto-start monitoring on launch if it was running before shutdown."""
        config = cfg_module.load_config()
        if config.get("monitoring_active") and config.get("products"):
            self._toggle_monitoring()

    # ── Settings dialog ───────────────────────

    def _show_settings(self):
        """Opens a dialog to configure email, check interval, and language."""
        config = cfg_module.load_config()
        total_min = config.get("check_interval_minutes", 1440)
        # Snap to nearest supported interval (6h/12h/24h)
        _supported = [6, 12, 24]
        _init_hrs = min(_supported, key=lambda h: abs(h * 60 - total_min))

        dlg = tk.Toplevel(self)
        dlg.title(_t("settings_title"))
        dlg.resizable(False, False)
        dlg.grab_set()  # modal
        dlg.transient(self)

        # column helpers — in RTL col 0 is on the right visually (label side),
        # so we swap: widget goes into col 0, label into col 1.
        _c_lbl = 1 if _IS_RTL else 0   # column for text labels
        _c_wid = 0 if _IS_RTL else 1   # column for input widgets
        _pad_lbl = (8, 0) if _IS_RTL else (0, 8)   # label side-padding
        _lbl_anchor = "ne" if _IS_RTL else "nw"    # LabelFrame title position

        # ── Email alerts ──
        email_cfg = config.get("email", {})
        frm_email = tk.LabelFrame(dlg, text=_t("email_section"), padx=12, pady=8,
                                  font=("Segoe UI", 9), labelanchor=_lbl_anchor)
        frm_email.pack(padx=16, pady=(14, 6), fill=tk.X)

        ent_cfg = {"font": ("Segoe UI", 9), "relief": tk.SOLID, "bd": 1}
        email_var = tk.StringVar(value=email_cfg.get("recipient", ""))

        tk.Label(frm_email, text=_t("email_label"), font=("Segoe UI", 9),
                 anchor=_ANCHOR, justify=_JUSTIFY).grid(
            row=0, column=_c_lbl, sticky=_ANCHOR, padx=_pad_lbl)
        tk.Entry(frm_email, textvariable=email_var, width=32, **ent_cfg,
                 justify=_JUSTIFY).grid(row=0, column=_c_wid, sticky="ew")
        tk.Label(frm_email, text=_t("email_hint"),
                 font=("Segoe UI", 7), fg="#8B7355",
                 anchor=_ANCHOR, justify=_JUSTIFY).grid(
            row=1, column=0, columnspan=2, sticky="e" if _IS_RTL else "w", pady=(4, 0))
        frm_email.columnconfigure(_c_wid, weight=1)

        # ── Check interval ──
        frm = tk.LabelFrame(dlg, text=_t("interval_section"), padx=12, pady=8,
                             font=("Segoe UI", 9), labelanchor=_lbl_anchor)
        frm.pack(padx=16, pady=(0, 6), fill=tk.X)

        interval_var = tk.IntVar(value=_init_hrs)

        _radio_frame = tk.Frame(frm)
        _radio_frame.pack(anchor="e" if _IS_RTL else "w")
        for hrs, key in [(6, "interval_6h"), (12, "interval_12h"), (24, "interval_24h")]:
            tk.Radiobutton(
                _radio_frame, text=_t(key), variable=interval_var, value=hrs,
                font=("Segoe UI", 9), anchor=_ANCHOR,
            ).pack(side=tk.RIGHT if _IS_RTL else tk.LEFT, padx=(0, 16) if not _IS_RTL else (16, 0))

        tk.Label(frm, text=_t("interval_note"),
                 font=("Segoe UI", 7), fg="#8B7355", wraplength=340,
                 anchor=_ANCHOR, justify=_JUSTIFY).pack(
            anchor="e" if _IS_RTL else "w", pady=(6, 0))

        # ── Start with Windows ──
        frm3 = tk.LabelFrame(dlg, text=_t("startup_section"), padx=12, pady=8,
                              font=("Segoe UI", 9), labelanchor=_lbl_anchor)
        frm3.pack(padx=16, pady=(0, 6), fill=tk.X)

        autostart_var = tk.BooleanVar(value=self._get_autostart())
        _cb_frame = tk.Frame(frm3)
        _cb_frame.pack(anchor=_ANCHOR)
        if _IS_RTL:
            tk.Checkbutton(_cb_frame, text="", variable=autostart_var,
                           font=("Segoe UI", 9)).pack(side="right")
            tk.Label(_cb_frame, text=_t("startup_checkbox"),
                     font=("Segoe UI", 9), anchor="e").pack(side="right")
        else:
            tk.Checkbutton(_cb_frame, text=_t("startup_checkbox"),
                           variable=autostart_var,
                           font=("Segoe UI", 9)).pack(side="left")

        # ── Language ──
        frm_lang = tk.LabelFrame(dlg, text=_t("language_section"), padx=12, pady=8,
                                  font=("Segoe UI", 9), labelanchor=_lbl_anchor)
        frm_lang.pack(padx=16, pady=(0, 6), fill=tk.X)

        _lang_options = ["עברית (Hebrew)", "English"]
        _lang_display = "עברית (Hebrew)" if config.get("language", "he") == "he" else "English"
        lang_var = tk.StringVar(value=_lang_display)

        _inner_lang = tk.Frame(frm_lang)
        _inner_lang.grid(row=0, column=0, columnspan=2,
                         sticky="e" if _IS_RTL else "w")
        tk.Label(_inner_lang, text=_t("language_label"), font=("Segoe UI", 9),
                 anchor=_ANCHOR, justify=_JUSTIFY).grid(
            row=0, column=_c_lbl, padx=_pad_lbl, sticky=_ANCHOR)
        tk.OptionMenu(_inner_lang, lang_var, *_lang_options).grid(
            row=0, column=_c_wid, sticky=_ANCHOR)
        tk.Label(frm_lang, text=_t("language_restart_note"),
                 font=("Segoe UI", 7), fg="#8B7355",
                 anchor=_ANCHOR, justify=_JUSTIFY).grid(
            row=1, column=0, columnspan=2, sticky="e" if _IS_RTL else "w", pady=(4, 0))

        def _save():
            hrs = interval_var.get()
            total = hrs * 60
            # Save recipient email
            addr = email_var.get().strip()
            if addr:
                config.setdefault("email", {})["recipient"] = addr
            config["check_interval_minutes"] = total
            config["notification_cooldown_hours"] = 24  # fixed behind the scenes

            # Save language
            lang_key = "he" if lang_var.get().startswith("ע") else "en"
            lang_changed = lang_key != config.get("language", "he")
            config["language"] = lang_key

            cfg_module.save_config(config)
            self._set_autostart(autostart_var.get())

            autostart_msg = _t("autostart_on_log") if autostart_var.get() else _t("autostart_off_log")
            email_msg = _t("email_set_log", addr) if addr else ""
            self._append_log(_t("settings_saved", hrs, email_msg, autostart_msg))

            if lang_changed:
                messagebox.showinfo(_t("language_section"), _t("language_restart_note"), parent=dlg)

            dlg.destroy()

        tk.Button(dlg, text=_t("btn_save"), command=_save,
                  bg="#F5A31A", fg="white", relief=tk.FLAT,
                  padx=16, pady=4, font=("Segoe UI", 9),
                  cursor="hand2").pack(pady=(4, 14))

    # ── Button handlers ───────────────────────

    def _add_product(self):
        """
        Opens a multiline dialog — paste one or many ASINs / URLs (one per line).
        Product names are fetched automatically in background.
        """
        dlg = tk.Toplevel(self)
        dlg.title(_t("add_title"))
        dlg.resizable(True, True)
        dlg.grab_set()
        dlg.transient(self)
        dlg.minsize(500, 280)

        tk.Label(dlg,
                 text=_t("add_instruction"),
                 font=("Segoe UI", 9), anchor=_ANCHOR, justify=_JUSTIFY).pack(
            anchor=_ANCHOR, fill=tk.X, padx=12, pady=(12, 4))

        frm = tk.Frame(dlg, padx=12)
        frm.pack(fill=tk.BOTH, expand=True)
        txt = tk.Text(frm, height=9, font=("Consolas", 9), wrap=tk.NONE)
        sb  = ttk.Scrollbar(frm, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.LEFT, fill=tk.Y)
        txt.focus_set()

        tk.Label(dlg, text=_t("add_example"),
                 font=("Segoe UI", 8), fg="#8B7355",
                 anchor=_ANCHOR, justify=_JUSTIFY).pack(anchor=_ANCHOR, padx=12, pady=(3, 0))

        def _do_paste():
            try:
                clip = dlg.clipboard_get()
            except tk.TclError:
                return
            if clip:
                txt.insert(tk.INSERT, clip)

        def _do_add():
            raw = txt.get("1.0", tk.END).strip()
            dlg.destroy()
            if not raw:
                return
            lines = [l.strip() for l in raw.splitlines() if l.strip()]
            added_asins, errors = [], []
            for line in lines:
                before = {p["asin"] for p in cfg_module.load_config().get("products", [])}
                try:
                    cfg_module.add_product(line, line)
                    after = {p["asin"] for p in cfg_module.load_config().get("products", [])}
                    newly = after - before
                    if newly:
                        added_asins.extend(newly)
                    # duplicate → silently skip
                except ValueError:
                    errors.append(line)

            self._refresh_table()

            if added_asins:
                self._append_log(_t("log_added_products", len(added_asins)))
                self._status_var.set(_t("log_added_status", len(added_asins)))

            if errors:
                messagebox.showwarning(
                    _t("invalid_entries"),
                    f"{len(errors)} invalid entr{'y' if len(errors) == 1 else 'ies'}:\n\n" +
                    "\n".join(errors[:10]) + ("\n…" if len(errors) > 10 else ""),
                    parent=self,
                )

            if added_asins:
                if messagebox.askyesno(_t("check_now_title"), _t("check_now_prompt"), parent=self):
                    self._check_now()

        _btn_row = tk.Frame(dlg)
        _btn_row.pack(pady=(6, 12))
        tk.Button(_btn_row, text=_t("btn_paste"), command=_do_paste,
                  bg="#6B5E45", fg="white", relief=tk.FLAT,
                  padx=12, pady=4, font=("Segoe UI", 9),
                  cursor="hand2").pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(_btn_row, text=_t("btn_add"), command=_do_add,
                  bg="#F5A31A", fg="white", relief=tk.FLAT,
                  padx=16, pady=4, font=("Segoe UI", 9),
                  cursor="hand2").pack(side=tk.LEFT)

    def _remove_product(self):
        targets = self._checked_asins()
        if not targets:
            messagebox.showinfo(_t("no_selection"), _t("select_to_remove"), parent=self)
            return
        confirm_msg = _t("remove_confirm", ", ".join(targets))
        if messagebox.askyesno(_t("remove_title"), confirm_msg, parent=self):
            for asin in targets:
                cfg_module.remove_product(asin)
                self._append_log(_t("log_removed_action", asin))
            self._refresh_table()

    def _check_now(self):
        config = cfg_module.load_config()
        if not config.get("products"):
            messagebox.showinfo(_t("no_products_title"), _t("no_products_msg"), parent=self)
            return
        if not config.get("email", {}).get("recipient"):
            messagebox.showwarning(
                _t("email_required_title"),
                _t("email_required_msg"),
                parent=self,
            )
            self._show_settings()
            return
        if self._manual_check_running:
            messagebox.showinfo(_t("busy_title"), _t("busy_msg"), parent=self)
            return

        # Check only checked products; fall back to all when none checked
        checked = self._checked_asins()
        products_to_check = [p for p in config.get("products", [])
                             if p["asin"] in checked]
        if not products_to_check:
            products_to_check = config.get("products", [])
        check_config = {**config, "products": products_to_check}

        self._manual_check_running = True
        self._status_var.set(_t("check_running"))
        self._append_log("Manual check started...")

        def run():
            state = state_module.load_state()
            try:
                results     = asyncio.run(_run_check_all_products(check_config, state))
                product_map = {p["asin"]: p for p in check_config["products"]}

                free_items = []
                names_updated = False
                for r in results:
                    asin    = r.asin
                    s       = r.status.value
                    product = product_map.get(asin, {"asin": asin, "name": asin, "url": ""})
                    name    = product.get("name", asin)
                    label   = STATUS_LABELS.get(s, s)

                    # Auto-update name if checker fetched a real title
                    if r.product_name and r.product_name != asin \
                            and product.get("name", asin) == asin:
                        product["name"] = r.product_name
                        name = r.product_name
                        names_updated = True

                    cooldown = config.get("notification_cooldown_hours", 24)
                    notify = state_module.should_notify(state, asin, s, cooldown_hours=cooldown)
                    if notify:
                        free_items.append({"product": product, "shipping_text": r.raw_text,
                                           "found_in_aod": r.found_in_aod})

                    state = state_module.update_product_state(
                        state, asin, s, notified=notify,
                    )
                    self._log_queue.put(
                        f"__log__{datetime.now().strftime('%H:%M:%S')}  {name}: {label}"
                    )

                if names_updated:
                    cfg_module.save_config(config)

                if free_items:
                    names = ", ".join(
                        i["product"].get("name", i["product"].get("asin", ""))
                        for i in free_items
                    )
                    self._log_queue.put(f"__log__FREE shipping found for: {names}! Sending email...")
                    try:
                        _send_email_alert(config, free_items)
                        self._log_queue.put("__log__Email alert sent successfully.")
                    except RuntimeError as mail_err:
                        self._log_queue.put(f"__log__EMAIL ERROR: {mail_err}")

                state_module.save_state(state)
                self._log_queue.put("__log__Check complete.")
                self._log_queue.put("__refresh__")

            except Exception as e:
                self._log_queue.put(f"__log_error__Error: {e}")
            finally:
                self._manual_check_running = False

        threading.Thread(target=run, daemon=True).start()

    def _toggle_monitoring(self):
        if self._monitor_thread and self._monitor_thread.is_alive():
            self._stop_event.set()
            self._start_btn.configure(text=_t("btn_start_monitoring"),
                                      bg="#F5A31A", activebackground="#D9901A")
            self._status_var.set(_t("monitoring_stopped"))
            self._interval_var.set("")
            _cfg = cfg_module.load_config()
            _cfg["monitoring_active"] = False
            cfg_module.save_config(_cfg)
        else:
            config = cfg_module.load_config()
            if not config.get("products"):
                messagebox.showinfo(_t("no_products_title"), _t("no_products_msg"), parent=self)
                return
            email_cfg = config.get("email", {})
            if not email_cfg.get("recipient"):
                messagebox.showwarning(
                    _t("email_not_conf_title"),
                    _t("email_not_conf_msg"),
                    parent=self)
                self._show_settings()
                return
            self._stop_event = threading.Event()
            self._monitor_thread = MonitorThread(
                self._log_queue, self._refresh_table, self._stop_event
            )
            self._monitor_thread.start()
            self._start_btn.configure(text=_t("btn_stop_monitoring"),
                                      bg="#dc2626", activebackground="#b91c1c")
            interval = config.get("check_interval_minutes", 60)
            self._status_var.set(_t("monitoring_every", interval))
            config["monitoring_active"] = True
            cfg_module.save_config(config)

    # ── Log helpers ───────────────────────────

    def _append_log(self, msg: str, tag: str = "info"):
        ts   = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        self._log_text.configure(state=tk.NORMAL)
        self._log_text.insert(tk.END, line, tag)
        lines = int(self._log_text.index(tk.END).split(".")[0])
        if lines > MAX_LOG_LINES:
            self._log_text.delete("1.0", f"{lines - MAX_LOG_LINES}.0")
        self._log_text.see(tk.END)
        self._log_text.configure(state=tk.DISABLED)

    def _clear_log(self):
        self._log_text.configure(state=tk.NORMAL)
        self._log_text.delete("1.0", tk.END)
        self._log_text.configure(state=tk.DISABLED)

    def _toggle_log(self):
        if self._log_visible:
            self._log_container.pack_forget()
            self._log_visible = False
            self._toggle_log_btn.configure(text=_t("btn_log_show"))
        else:
            self._log_container.pack(fill=tk.BOTH, expand=True)
            self._log_visible = True
            self._toggle_log_btn.configure(text=_t("btn_log_hide"))

    def _poll_log(self):
        try:
            while True:
                msg = self._log_queue.get_nowait()
                if msg.startswith("__next_run__"):
                    self._interval_var.set(f"{_t('next_check_label')} {msg[len('__next_run__'):]}")
                elif msg.startswith("__refresh__"):
                    self._refresh_table()
                    self._status_var.set(_t("check_complete"))
                elif msg.startswith("__log_free__"):
                    self._append_log(msg[len("__log_free__"):], "free")
                    self._refresh_table()
                    self._status_var.set(_t("free_detected"))
                elif msg.startswith("__log_error__"):
                    self._append_log(msg[len("__log_error__"):], "error")
                    self._status_var.set("Error — see log.")
                elif msg.startswith("__status__"):
                    self._status_var.set(msg[len("__status__"):])
                elif msg.startswith("__log__"):
                    self._append_log(msg[len("__log__"):])
                else:
                    if "FREE" in msg and "detected" in msg:
                        self._append_log(msg, "free")
                        self._refresh_table()
                        self._status_var.set(_t("free_detected"))
                    elif "error" in msg.lower():
                        self._append_log(msg, "error")
                    else:
                        self._append_log(msg)
                        if "complete" in msg.lower():
                            self._refresh_table()
        except queue.Empty:
            pass
        self.after(400, self._poll_log)

    # ── Window close ─────────────────────────

    def _on_close(self):
        if _TRAY_AVAILABLE and self._tray_icon:
            self.withdraw()  # hide to tray — monitoring keeps running
        else:
            if self._monitor_thread and self._monitor_thread.is_alive():
                if not messagebox.askyesno(_t("quit_title"),
                                           _t("quit_msg"),
                                           parent=self):
                    return
                self._stop_event.set()
            self.destroy()

    # ── System tray ───────────────────────────

    @staticmethod
    def _create_tray_image():
        size = 64
        img = _PILImage.new('RGBA', (size, size), (0, 0, 0, 0))
        d = _PILDraw.Draw(img)
        d.ellipse([0, 0, size - 1, size - 1], fill='#FF9900')
        # white checkmark
        d.line([(15, 34), (27, 46), (50, 20)], fill='white', width=7)
        return img

    def _log_tray_error(self, msg: str):
        try:
            log_path = os.path.join(os.getcwd(), "tray_error.log")
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(msg + "\n")
        except Exception:
            pass

    def _tray_run_safe(self):
        try:
            self._tray_icon.run()
        except Exception as exc:
            self._log_tray_error(f"tray run() crashed: {repr(exc)}")

    def _init_tray(self):
        if not _TRAY_AVAILABLE:
            if _tray_import_error:
                self._log_tray_error(f"pystray/PIL import failed: {_tray_import_error}")
            return
        try:
            menu = pystray.Menu(
                pystray.MenuItem(_t("tray_open"), self._tray_open, default=True),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem(_t("tray_exit"), self._tray_exit),
            )
            self._tray_icon = pystray.Icon(
                'AmazonMonitor',
                self._create_tray_image(),
                'Amazon Israel Free Ship Alert',
                menu,
            )
            self._tray_thread = threading.Thread(
                target=self._tray_run_safe, daemon=True)
            self._tray_thread.start()
        except Exception as exc:
            self._tray_icon = None
            self._log_tray_error(f"tray init failed: {repr(exc)}")

    def _tray_open(self, icon=None, item=None):
        self.after(0, self.deiconify)
        self.after(0, self.lift)

    def _tray_exit(self, icon=None, item=None):
        if self._tray_icon:
            self._tray_icon.stop()
        self.after(0, self._force_quit)

    def _force_quit(self):
        if self._monitor_thread and self._monitor_thread.is_alive():
            self._stop_event.set()
        self.destroy()

    # ── Autostart ─────────────────────────────

    @staticmethod
    def _get_autostart() -> bool:
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                r"Software\Microsoft\Windows\CurrentVersion\Run") as k:
                winreg.QueryValueEx(k, "AmazonFreeShippingMonitor")
                return True
        except (FileNotFoundError, OSError):
            return False

    def _set_autostart(self, enabled: bool):
        try:
            import winreg
            with winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    r"Software\Microsoft\Windows\CurrentVersion\Run",
                    0, winreg.KEY_SET_VALUE) as k:
                if enabled:
                    # Use wscript.exe + VBS so Windows boots the app with no console
                    # window and with the correct working directory.
                    wscript  = os.path.join(
                        os.environ.get("SystemRoot", r"C:\Windows"),
                        "System32", "wscript.exe")
                    vbs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                            "Start Monitor.vbs")
                    if os.path.exists(vbs_path):
                        cmd = f'"{wscript}" "{vbs_path}"'
                    else:
                        # Fallback when VBS is missing (shouldn't happen in normal use)
                        pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
                        if not os.path.exists(pythonw):
                            pythonw = sys.executable
                        cmd = f'"{pythonw}" "{os.path.abspath(__file__)}"'
                    winreg.SetValueEx(k, "AmazonFreeShippingMonitor", 0,
                                      winreg.REG_SZ, cmd)
                else:
                    try:
                        winreg.DeleteValue(k, "AmazonFreeShippingMonitor")
                    except FileNotFoundError:
                        pass
        except Exception as exc:
            messagebox.showerror(_t("autostart_error"),
                                 f"Could not update startup entry:\n{exc}",
                                 parent=self)

    def _sync_autostart(self):
        """If autostart is registered, silently re-register using the current (VBS) format."""
        if self._get_autostart():
            self._set_autostart(True)

    # ── Auto-update ───────────────────────────

    _GITHUB_API = (
        "https://api.github.com/repos/"
        "ilan316/Amazon-Free-Shipping-to-Israel-Alert/releases/latest"
    )

    def _check_for_updates(self):
        """Fetch latest GitHub release in a background thread; show dialog if newer."""
        def _check():
            import json as _json, ssl
            try:
                ctx = ssl.create_default_context()
                ctx.check_hostname = False
                ctx.verify_mode = ssl.CERT_NONE
                req = urllib.request.Request(
                    self._GITHUB_API,
                    headers={"User-Agent": "AmazonIsraelFreeShipAlert"},
                )
                with urllib.request.urlopen(req, timeout=10, context=ctx) as resp:
                    data = _json.loads(resp.read().decode())
            except Exception as _exc:
                self.after(0, self._append_log,
                           f"[Update check failed: {_exc}]", "error")
                return

            tag = data.get("tag_name", "").lstrip("v")
            if not tag:
                return

            download_url = ""
            for asset in data.get("assets", []):
                if asset.get("name") == "AmazonIsraelFreeShipAlert.exe":
                    download_url = asset.get("browser_download_url", "")
                    break
            if not download_url:
                self.after(0, self._append_log,
                           f"[Update] v{tag} found but no matching asset — skipping")
                return

            try:
                current = tuple(int(x) for x in __version__.split("."))
                latest  = tuple(int(x) for x in tag.split("."))
            except ValueError:
                return

            if latest > current:
                self.after(0, self._append_log,
                           f"[Update] v{tag} available — showing dialog...")
                self.after(0, self._show_update_dialog, tag, download_url)

        threading.Thread(target=_check, daemon=True).start()

    def _show_update_dialog(self, version: str, download_url: str):
        """Show a modal dialog offering to download the new version."""
        # Bring main window to front (it may be hidden in tray)
        self.deiconify()
        self.lift()

        dlg = tk.Toplevel(self)
        dlg.withdraw()  # hide until positioned — prevents flash at (0,0)
        dlg.title(_t("update_available_title"))
        dlg.resizable(False, False)
        dlg.transient(self)
        try:
            dlg.iconbitmap(os.path.join(os.getcwd(), "icon.ico"))
        except Exception:
            pass

        msg_text = _t("update_available_msg", version, __version__)
        tk.Label(
            dlg, text=msg_text,
            font=("Segoe UI", 10), padx=24, pady=20,
            justify="right" if _IS_RTL else "left",
            anchor=_ANCHOR,
        ).pack(anchor=_ANCHOR)

        btn_row = tk.Frame(dlg, padx=20, pady=(0, 18))
        btn_row.pack()

        def _on_now():
            dlg.destroy()
            self._start_update_download(download_url)

        _btn_side = tk.RIGHT if _IS_RTL else tk.LEFT
        tk.Button(
            btn_row, text=_t("btn_update_now"),
            bg="#F5A31A", fg="white", relief=tk.FLAT,
            font=("Segoe UI", 10, "bold"), padx=16, pady=6,
            cursor="hand2", command=_on_now,
        ).pack(side=_btn_side, padx=(8, 0) if _IS_RTL else (0, 8))

        tk.Button(
            btn_row, text=_t("btn_update_later"),
            relief=tk.FLAT, font=("Segoe UI", 10),
            padx=16, pady=6, cursor="hand2",
            command=dlg.destroy,
        ).pack(side=_btn_side)

        dlg.update_idletasks()  # calculate widget sizes without showing
        dw = dlg.winfo_reqwidth()
        dh = dlg.winfo_reqheight()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        # Always center on screen — parent position is unreliable after tray restore
        x = max(0, min((sw - dw) // 2, sw - dw))
        y = max(0, min((sh - dh) // 2, sh - dh - 48))  # 48 ≈ taskbar
        dlg.geometry(f"+{x}+{y}")
        dlg.deiconify()  # show directly at correct position — no flash
        dlg.attributes("-topmost", True)   # force above all windows (Windows focus rules)
        dlg.lift()
        dlg.focus_force()
        dlg.after(300, lambda: dlg.attributes("-topmost", False))  # release topmost after 300ms
        dlg.grab_set()   # grab only after window is visible — avoids UI freeze

    def _start_update_download(self, download_url: str):
        """Download the new installer to a temp file, then launch it."""
        install_dir = os.path.dirname(os.path.abspath(__file__))
        self._append_log(_t("downloading_update"))
        self._status_var.set(_t("downloading_update"))

        def _download():
            import tempfile, ssl
            try:
                ctx = ssl.create_default_context()
                ctx.check_hostname = False
                ctx.verify_mode = ssl.CERT_NONE
                req = urllib.request.Request(
                    download_url,
                    headers={"User-Agent": "AmazonIsraelFreeShipAlert"},
                )
                fd, tmp_path = tempfile.mkstemp(
                    suffix=".exe", prefix="AmazonUpdate_")
                os.close(fd)
                with urllib.request.urlopen(req, timeout=120, context=ctx) as resp:
                    total = int(resp.headers.get("Content-Length", 0))
                    downloaded = 0
                    with open(tmp_path, "wb") as fh:
                        while True:
                            chunk = resp.read(65536)
                            if not chunk:
                                break
                            fh.write(chunk)
                            downloaded += len(chunk)
                            if total > 0:
                                pct = downloaded * 100 // total
                                self.after(
                                    0, self._status_var.set,
                                    f"{_t('downloading_update')} {pct}%",
                                )
            except Exception as exc:
                self.after(0, self._append_log,
                           _t("update_failed", str(exc)), "error")
                return
            self.after(0, self._launch_update, tmp_path, install_dir)

        threading.Thread(target=_download, daemon=True).start()

    def _launch_update(self, installer_path: str, install_dir: str):
        """Launch the downloaded installer and fully close this app."""
        try:
            subprocess.Popen(
                [installer_path, f"--dir={install_dir}", "--auto-update"],
                creationflags=0x00000008 | 0x08000000,  # DETACHED_PROCESS | CREATE_NO_WINDOW
            )
        except Exception as exc:
            self._append_log(f"[Update] Failed to launch installer: {exc}", "error")
            return
        self._append_log("[Update] Installer launched — closing app...")
        if self._tray_icon:
            self._tray_icon.stop()
        if self._monitor_thread and self._monitor_thread.is_alive():
            self._stop_event.set()
        self.destroy()


# ──────────────────────────────────────────────
# Single-instance guard
# ──────────────────────────────────────────────

def _ensure_single_instance() -> bool:
    """
    Returns True if this is the first running instance.
    If another instance already holds the mutex, brings its window to the
    foreground and returns False (caller should exit immediately).
    """
    try:
        import ctypes
        _k32 = ctypes.windll.kernel32
        _k32.CreateMutexW(None, True, "AmazonIsraelFreeShipAlert_SingleInstance")
        if _k32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
            _u32 = ctypes.windll.user32
            # FindWindowW requires exact title — use EnumWindows for partial match
            # (title includes version suffix, e.g. "Amazon Israel Free Ship Alert  v1.5.0")
            _found = [0]
            _WNDENUMPROC = ctypes.WINFUNCTYPE(
                ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p
            )
            @_WNDENUMPROC
            def _cb(hwnd, _):
                n = _u32.GetWindowTextLengthW(hwnd) + 1
                buf = ctypes.create_unicode_buffer(n)
                _u32.GetWindowTextW(hwnd, buf, n)
                if "Amazon Israel Free Ship Alert" in buf.value:
                    _found[0] = hwnd
                    return False   # stop enumerating
                return True
            _u32.EnumWindows(_cb, 0)
            if _found[0]:
                _u32.ShowWindow(_found[0], 5)      # SW_SHOW — unhides withdrawn window
                _u32.SetForegroundWindow(_found[0])
            return False
    except Exception:
        pass
    return True


# ──────────────────────────────────────────────
if __name__ == "__main__":
    # Tell Windows to treat this as a distinct app (not pythonw.exe),
    # so the taskbar shows our custom icon instead of the Python icon.
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "AmazonIsraelFreeShipAlert.1.0"
        )
    except Exception:
        pass

    if not _ensure_single_instance():
        sys.exit(0)
    app = App()
    app.mainloop()
