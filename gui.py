"""
Amazon Israel Free Ship Alert — Tkinter GUI
Run with: python gui.py
"""

import sys
import os

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
from checker import check_all_products, ShippingStatus
from notifier import send_batch_free_shipping_alert
from version import __version__

# ──────────────────────────────────────────────
# Status display helpers
# ──────────────────────────────────────────────

STATUS_LABELS = {
    "FREE":     "✅ Eligible — FREE Shipping",
    "PAID":     "❌ Not eligible",
    "NO_SHIP":  "❌ Not eligible",
    "UNKNOWN":  "❌ Not eligible",
    "ERROR":    "❌ Not eligible",
    "—":        "— Not checked yet",
    "PAUSED":   "⏸  Paused",
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
                    results = asyncio.run(check_all_products(check_config, state))
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
                            free_items.append({"product": product, "shipping_text": result.raw_text})

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
                            send_batch_free_shipping_alert(config, free_items)
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
        self._tray_icon = None
        self._tray_thread: threading.Thread | None = None
        self._sort_col: str | None = None
        self._sort_rev: bool = False

        self._build_ui()
        self._refresh_table()
        self._poll_log()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._init_tray()
        self._sync_autostart()                          # upgrade registry to VBS if needed
        self.after(800, self._maybe_autostart_monitoring)  # resume monitoring if it was running

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
        # Products table
        top = tk.Frame(self, padx=8, pady=6)
        top.pack(fill=tk.BOTH, expand=True)

        # Logo banner above the table
        self._logo_img = None
        try:
            from PIL import Image as _PilImg, ImageTk as _PilImgTk
            _logo_path = os.path.join(os.getcwd(), "image.jpg")
            if os.path.exists(_logo_path):
                _pil = _PilImg.open(_logo_path).convert("RGBA")
                _pil = _pil.resize((380, 157), _PilImg.LANCZOS)
                self._logo_img = _PilImgTk.PhotoImage(_pil)
                _logo_frame = tk.Frame(top, bg="white")
                _logo_frame.pack(anchor="center", pady=(0, 4))
                tk.Label(_logo_frame, image=self._logo_img, bg="white").pack()
        except Exception:
            pass

        tk.Label(top, text="Monitored Products", font=("Segoe UI", 11, "bold")).pack(anchor="w")

        cols = ("name", "asin", "status", "last_checked")
        self.tree = ttk.Treeview(top, columns=cols, show="headings", selectmode="browse", height=8)
        for _col, _txt in [("name", "Product Name"), ("asin", "ASIN"),
                            ("status", "Eligible"), ("last_checked", "Last Checked")]:
            self.tree.heading(_col, text=_txt, command=lambda c=_col: self._sort_by(c))
        self.tree.column("name",         width=260, anchor="w")
        self.tree.column("asin",         width=110, anchor="center")
        self.tree.column("status",       width=175, anchor="center")
        self.tree.column("last_checked", width=140, anchor="center")

        vsb = ttk.Scrollbar(top, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.LEFT, fill=tk.Y)

        self.tree.bind("<Double-1>", self._on_product_double_click)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        for status, colour in STATUS_COLORS.items():
            self.tree.tag_configure(status, foreground=colour)

        # Action buttons
        mid = tk.Frame(self, padx=8, pady=4)
        mid.pack(fill=tk.X)

        btn_cfg = {"padx": 9, "pady": 4, "relief": tk.FLAT, "bd": 0,
                   "font": ("Segoe UI", 9), "cursor": "hand2"}

        tk.Button(mid, text="➕  Add Product",
                  bg="#0066cc", fg="white", activebackground="#0055aa",
                  command=self._add_product, **btn_cfg).pack(side=tk.LEFT, padx=(0, 4))

        tk.Button(mid, text="🗑  Remove",
                  bg="#cc3300", fg="white", activebackground="#aa2200",
                  command=self._remove_product, **btn_cfg).pack(side=tk.LEFT, padx=(0, 4))

        self._pause_btn = tk.Button(mid, text="⏸  Pause",
                                    bg="#7a6000", fg="white", activebackground="#5e4800",
                                    command=self._toggle_pause, **btn_cfg)
        self._pause_btn.pack(side=tk.LEFT, padx=(0, 4))

        tk.Button(mid, text="🔍  Check Now",
                  bg="#4a4a4a", fg="white", activebackground="#333333",
                  command=self._check_now, **btn_cfg).pack(side=tk.LEFT, padx=(0, 4))

        tk.Button(mid, text="⚙  Settings",
                  bg="#555555", fg="white", activebackground="#444444",
                  command=self._show_settings, **btn_cfg).pack(side=tk.LEFT, padx=(0, 16))

        self._start_btn = tk.Button(mid, text="▶  Start Monitoring",
                                    bg="#1a7a1a", fg="white", activebackground="#145e14",
                                    command=self._toggle_monitoring, **btn_cfg)
        self._start_btn.pack(side=tk.LEFT, padx=(0, 4))

        self._interval_var = tk.StringVar(value="")
        tk.Label(mid, textvariable=self._interval_var,
                 font=("Segoe UI", 9), fg="#555555").pack(side=tk.RIGHT)

        # Log area
        bot = tk.Frame(self, padx=8)
        bot.pack(fill=tk.BOTH, expand=False, pady=(0, 8))

        hdr = tk.Frame(bot)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="Log", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        tk.Button(hdr, text="Clear", font=("Segoe UI", 8), relief=tk.FLAT,
                  command=self._clear_log, cursor="hand2").pack(side=tk.RIGHT)

        self._log_text = tk.Text(
            bot, height=10, state=tk.DISABLED,
            font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
            relief=tk.FLAT, padx=6, pady=4, wrap=tk.WORD,
        )
        lsb = ttk.Scrollbar(bot, orient="vertical", command=self._log_text.yview)
        self._log_text.configure(yscrollcommand=lsb.set)
        self._log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        lsb.pack(side=tk.LEFT, fill=tk.Y)

        self._log_text.tag_configure("free",  foreground="#4ec94e")
        self._log_text.tag_configure("error", foreground="#f48771")
        self._log_text.tag_configure("info",  foreground="#d4d4d4")

        # Status bar
        self._status_var = tk.StringVar(value="Ready")
        tk.Label(self, textvariable=self._status_var, relief=tk.SUNKEN,
                 anchor="w", font=("Segoe UI", 8), fg="#555555",
                 padx=8).pack(side=tk.BOTTOM, fill=tk.X)

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
        _base = {"name": "Product Name", "asin": "ASIN",
                 "status": "Eligible", "last_checked": "Last Checked"}
        for col, base in _base.items():
            arrow = (" ▼" if self._sort_rev else " ▲") if self._sort_col == col else ""
            self.tree.heading(col, text=base + arrow)

        # Build row data
        rows = []
        for p in config.get("products", []):
            asin     = p["asin"]
            paused   = p.get("paused", False)
            s        = state.get(asin, {})
            raw_last = s.get("last_checked") or ""
            disp_last = raw_last.replace("T", "  ") if raw_last else "—"
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

        # Repopulate tree
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for r in rows:
            self.tree.insert("", tk.END, iid=r["asin"],
                             values=(r["name"], r["asin"], r["label"], r["disp_last"]),
                             tags=(r["tag"],))
        self._on_tree_select()

    def _on_product_double_click(self, event):
        """Double-clicking a product row opens its Amazon page in the system browser."""
        import webbrowser
        sel = self.tree.selection()
        if not sel:
            return
        asin = sel[0]
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
            self._pause_btn.configure(text="⏸  Pause")
            return
        asin = sel[0]
        config = cfg_module.load_config()
        for p in config.get("products", []):
            if p["asin"] == asin:
                if p.get("paused", False):
                    self._pause_btn.configure(text="▶  Resume")
                else:
                    self._pause_btn.configure(text="⏸  Pause")
                return

    def _toggle_pause(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("No selection", "Select a product to pause/resume.", parent=self)
            return
        asin = sel[0]
        config = cfg_module.load_config()
        for p in config.get("products", []):
            if p["asin"] == asin:
                p["paused"] = not p.get("paused", False)
                action = "Paused" if p["paused"] else "Resumed"
                name   = p.get("name", asin)
                break
        cfg_module.save_config(config)
        self._refresh_table()
        self.tree.selection_set(asin)
        self._append_log(f"{action}: {name}")

    def _maybe_autostart_monitoring(self):
        """Auto-start monitoring on launch if it was running before shutdown."""
        config = cfg_module.load_config()
        if config.get("monitoring_active") and config.get("products"):
            self._toggle_monitoring()

    # ── Settings dialog ───────────────────────

    def _show_settings(self):
        """Opens a dialog to configure email, check interval, and notification cooldown."""
        config = cfg_module.load_config()
        total_min = config.get("check_interval_minutes", 60)
        init_days = total_min // (24 * 60)
        init_hrs  = (total_min % (24 * 60)) // 60
        init_mins = total_min % 60
        init_cooldown = config.get("notification_cooldown_hours", 24)

        dlg = tk.Toplevel(self)
        dlg.title("Settings")
        dlg.resizable(False, False)
        dlg.grab_set()  # modal
        dlg.transient(self)

        # ── Email alerts ──
        email_cfg = config.get("email", {})
        frm_email = tk.LabelFrame(dlg, text="Email alerts", padx=12, pady=8,
                                  font=("Segoe UI", 9))
        frm_email.pack(padx=16, pady=(14, 6), fill=tk.X)

        ent_cfg = {"font": ("Segoe UI", 9), "relief": tk.SOLID, "bd": 1}
        email_var = tk.StringVar(value=email_cfg.get("recipient", ""))

        tk.Label(frm_email, text="Your email address:", font=("Segoe UI", 9)).grid(
            row=0, column=0, sticky="w", padx=(0, 8))
        tk.Entry(frm_email, textvariable=email_var, width=32, **ent_cfg).grid(
            row=0, column=1, sticky="ew")
        tk.Label(frm_email,
                 text="Alerts will be sent to this address when FREE shipping is detected",
                 font=("Segoe UI", 7), fg="#777777").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))
        frm_email.columnconfigure(1, weight=1)

        # ── Check interval ──
        frm = tk.LabelFrame(dlg, text="Check interval", padx=12, pady=8,
                             font=("Segoe UI", 9))
        frm.pack(padx=16, pady=(14, 6), fill=tk.X)

        spin_cfg = {"width": 4, "font": ("Segoe UI", 10), "justify": "center"}

        days_var = tk.StringVar(value=str(init_days))
        hrs_var  = tk.StringVar(value=str(init_hrs))
        mins_var = tk.StringVar(value=str(init_mins))

        for col, label, var, lo, hi in [
            (0, "Days",    days_var, 0, 365),
            (2, "Hours",   hrs_var,  0, 23),
            (4, "Minutes", mins_var, 0, 59),
        ]:
            tk.Label(frm, text=label, font=("Segoe UI", 9)).grid(
                row=0, column=col, padx=(0, 2))
            tk.Spinbox(frm, from_=lo, to=hi, textvariable=var,
                       **spin_cfg).grid(row=0, column=col + 1, padx=(0, 10))

        # ── Notification cooldown ──
        frm2 = tk.LabelFrame(dlg, text="Email notification cooldown", padx=12, pady=8,
                              font=("Segoe UI", 9))
        frm2.pack(padx=16, pady=(0, 6), fill=tk.X)

        cooldown_var = tk.StringVar(value=str(init_cooldown))
        tk.Label(frm2, text="Hours between repeat emails for the same FREE product:",
                 font=("Segoe UI", 9)).grid(row=0, column=0, padx=(0, 8), sticky="w")
        tk.Spinbox(frm2, from_=1, to=720, textvariable=cooldown_var,
                   **spin_cfg).grid(row=0, column=1, padx=(0, 4))

        # ── Start with Windows ──
        frm3 = tk.LabelFrame(dlg, text="Windows startup", padx=12, pady=8,
                              font=("Segoe UI", 9))
        frm3.pack(padx=16, pady=(0, 6), fill=tk.X)

        autostart_var = tk.BooleanVar(value=self._get_autostart())
        tk.Checkbutton(frm3, text="Start automatically when Windows boots",
                       variable=autostart_var,
                       font=("Segoe UI", 9)).pack(anchor="w")

        def _save():
            try:
                d = max(0, int(days_var.get()))
                h = max(0, int(hrs_var.get()))
                m = max(0, int(mins_var.get()))
                cooldown = max(1, int(cooldown_var.get()))
            except ValueError:
                messagebox.showerror("Invalid", "Enter valid numbers.", parent=dlg)
                return
            total = d * 24 * 60 + h * 60 + m
            if total < 1:
                messagebox.showerror("Invalid",
                    "Interval must be at least 1 minute.", parent=dlg)
                return
            # Save recipient email
            addr = email_var.get().strip()
            if addr:
                config.setdefault("email", {})["recipient"] = addr
            config["check_interval_minutes"] = total
            config["notification_cooldown_hours"] = cooldown
            cfg_module.save_config(config)
            self._set_autostart(autostart_var.get())
            autostart_msg = "  Auto-start: ON." if autostart_var.get() else "  Auto-start: OFF."
            email_msg = f"  Email: {addr}." if addr else ""
            self._append_log(
                f"Settings saved. Interval: {d}d {h}h {m}m ({total} min). "
                f"Cooldown: {cooldown}h.{email_msg}{autostart_msg}"
            )
            dlg.destroy()

        tk.Button(dlg, text="Save", command=_save,
                  bg="#0066cc", fg="white", relief=tk.FLAT,
                  padx=16, pady=4, font=("Segoe UI", 9),
                  cursor="hand2").pack(pady=(4, 14))

    # ── Button handlers ───────────────────────

    def _add_product(self):
        """
        Opens a multiline dialog — paste one or many ASINs / URLs (one per line).
        Product names are fetched automatically in background.
        """
        dlg = tk.Toplevel(self)
        dlg.title("Add Product(s)")
        dlg.resizable(True, True)
        dlg.grab_set()
        dlg.transient(self)
        dlg.minsize(500, 280)

        tk.Label(dlg,
                 text="Enter one or more Amazon URLs or ASINs — one per line:",
                 font=("Segoe UI", 9)).pack(anchor="w", padx=12, pady=(12, 4))

        frm = tk.Frame(dlg, padx=12)
        frm.pack(fill=tk.BOTH, expand=True)
        txt = tk.Text(frm, height=9, font=("Consolas", 9), wrap=tk.NONE)
        sb  = ttk.Scrollbar(frm, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.LEFT, fill=tk.Y)
        txt.focus_set()

        tk.Label(dlg, text="Example:  B08N5WRWNW   or   https://www.amazon.com/dp/B08N5WRWNW",
                 font=("Segoe UI", 8), fg="#777777").pack(anchor="w", padx=12, pady=(3, 0))

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
                self._append_log(
                    f"Added {len(added_asins)} product(s). "
                    "Click Check Now to fetch names and check shipping.")
                self._status_var.set(f"Added {len(added_asins)} product(s).")

            if errors:
                messagebox.showwarning(
                    "Invalid entries",
                    f"{len(errors)} invalid entr{'y' if len(errors) == 1 else 'ies'}:\n\n" +
                    "\n".join(errors[:10]) + ("\n…" if len(errors) > 10 else ""),
                    parent=self,
                )

        tk.Button(dlg, text="Add", command=_do_add,
                  bg="#0066cc", fg="white", relief=tk.FLAT,
                  padx=16, pady=4, font=("Segoe UI", 9),
                  cursor="hand2").pack(pady=(6, 12))

    def _remove_product(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("No selection", "Select a product to remove.", parent=self)
            return
        asin = sel[0]
        if messagebox.askyesno("Remove Product",
                               f"Remove ASIN {asin} from monitoring?", parent=self):
            cfg_module.remove_product(asin)
            self._refresh_table()
            self._append_log(f"Removed product: {asin}")

    def _check_now(self):
        config = cfg_module.load_config()
        if not config.get("products"):
            messagebox.showinfo("No Products", "Add at least one product first.", parent=self)
            return
        if not config.get("email", {}).get("recipient"):
            messagebox.showwarning(
                "Email required",
                "Please enter your email address before running a check.\n\n"
                "Click OK to open Settings.",
                parent=self,
            )
            self._show_settings()
            return
        if self._monitor_thread and self._monitor_thread.is_alive():
            messagebox.showinfo("Busy", "A check is already running.", parent=self)
            return

        self._status_var.set("Running check...")
        self._append_log("Manual check started...")

        def run():
            state = state_module.load_state()
            try:
                results     = asyncio.run(check_all_products(config, state))
                product_map = {p["asin"]: p for p in config["products"]}

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
                        free_items.append({"product": product, "shipping_text": r.raw_text})

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
                        send_batch_free_shipping_alert(config, free_items)
                        self._log_queue.put("__log__Email alert sent successfully.")
                    except RuntimeError as mail_err:
                        self._log_queue.put(f"__log__EMAIL ERROR: {mail_err}")

                state_module.save_state(state)
                self._log_queue.put("__log__Check complete.")
                self._log_queue.put("__refresh__")

            except Exception as e:
                self._log_queue.put(f"__log_error__Error: {e}")

        threading.Thread(target=run, daemon=True).start()

    def _toggle_monitoring(self):
        if self._monitor_thread and self._monitor_thread.is_alive():
            self._stop_event.set()
            self._start_btn.configure(text="▶  Start Monitoring",
                                      bg="#1a7a1a", activebackground="#145e14")
            self._status_var.set("Monitoring stopped.")
            self._interval_var.set("")
            _cfg = cfg_module.load_config()
            _cfg["monitoring_active"] = False
            cfg_module.save_config(_cfg)
        else:
            config = cfg_module.load_config()
            if not config.get("products"):
                messagebox.showinfo("No Products", "Add at least one product first.", parent=self)
                return
            email_cfg = config.get("email", {})
            if not email_cfg.get("recipient"):
                messagebox.showwarning(
                    "Email not configured",
                    "No recipient email address is set.\n\n"
                    "Alerts will not be sent until you configure your email.\n\n"
                    "Please click \u2699 Settings and enter your email address.",
                    parent=self)
                self._show_settings()
                return
            self._stop_event = threading.Event()
            self._monitor_thread = MonitorThread(
                self._log_queue, self._refresh_table, self._stop_event
            )
            self._monitor_thread.start()
            self._start_btn.configure(text="⏹  Stop Monitoring",
                                      bg="#cc3300", activebackground="#aa2200")
            interval = config.get("check_interval_minutes", 60)
            self._status_var.set(f"Monitoring — every {interval} min")
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

    def _poll_log(self):
        try:
            while True:
                msg = self._log_queue.get_nowait()
                if msg.startswith("__next_run__"):
                    self._interval_var.set(f"Next: {msg[len('__next_run__'):]}")
                elif msg.startswith("__refresh__"):
                    self._refresh_table()
                    self._status_var.set("Check complete.")
                elif msg.startswith("__log_free__"):
                    self._append_log(msg[len("__log_free__"):], "free")
                    self._refresh_table()
                    self._status_var.set("FREE shipping detected!")
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
                        self._status_var.set("FREE shipping detected!")
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
                if not messagebox.askyesno("Quit",
                                           "Monitoring is running. Stop it and exit?",
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
                pystray.MenuItem('Open Amazon Israel Free Ship Alert', self._tray_open, default=True),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem('Exit', self._tray_exit),
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
            messagebox.showerror("Autostart Error",
                                 f"Could not update startup entry:\n{exc}",
                                 parent=self)

    def _sync_autostart(self):
        """If autostart is registered, silently re-register using the current (VBS) format."""
        if self._get_autostart():
            self._set_autostart(True)


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
            _hwnd = _u32.FindWindowW(None, "Amazon Israel Free Ship Alert")
            if _hwnd:
                _u32.ShowWindow(_hwnd, 9)      # SW_RESTORE (un-minimise / un-hide)
                _u32.SetForegroundWindow(_hwnd)
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
