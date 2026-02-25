# Amazon Free Shipping to Israel Alert

A Windows desktop app that monitors Amazon product pages and sends a Gmail alert the moment free shipping to Israel becomes available.

---

## How it works

The app opens each product page in a headless Chromium browser (Playwright), sets the delivery address to Israel, and reads the shipping block. Free shipping is detected when the page contains:

> *"FREE delivery … to Israel … on eligible orders"*

When that text appears, a Gmail notification is sent to your configured recipient email.

---

## Installation (end users)

1. Download **`AmazonIsraelFreeShipAlert_Setup_vX.X.X.exe`** from [Releases](../../releases)
2. Run it — the installer will automatically:
   - Download and install **Python 3.13** (if not already present)
   - Install **Visual C++ Redistributable 2022** (required for Playwright)
   - Install all Python packages
   - Install the **Chromium** browser
   - Create a Desktop shortcut and add the app to Windows startup
3. Open the app → click **Settings** → enter your **recipient email address**
4. Add product URLs (Amazon ASIN links) and click **Start Monitoring**

> The installer requires an internet connection (~150 MB total download).

---

## Features

- Monitors multiple products simultaneously
- Configurable check interval (default: every 3 hours)
- Gmail email alert on free-shipping detection
- System tray icon — runs silently in the background
- Autostart on Windows login
- Single-instance: clicking the shortcut while the app is running brings the window to the front

---

## Development

### Requirements

- Windows 10/11
- Python 3.13 (`python build_installer.py` uses whichever `python` is on PATH)
- PyInstaller (install into the same Python: `pip install pyinstaller`)
- Pillow (`pip install Pillow`)

### Run from source

```bash
pip install -r requirements.txt
python -m playwright install chromium
python gui.py
```

### Build the installer

```bash
# 1. Edit version.py → bump __version__
# 2. Run:
python build_installer.py
# Output: AmazonIsraelFreeShipAlert_Setup_v{VERSION}.exe
```

The build script:
1. Compiles a tiny launcher exe (`AmazonIsraelFreeShipAlert.exe`) via PyInstaller
2. Base64-encodes all source files into `install.py`
3. Bundles `install.py` into the Setup exe via PyInstaller

Do **not** commit `*.exe`, `install.py`, or `_build_*_tmp/` — they are generated artifacts and are excluded by `.gitignore`.

### Environment variables (`.env`)

Create a `.env` file (not committed) with the Gmail App Password used by the sender account:

```
GMAIL_APP_PASSWORD=your_app_password_here
```

This is baked into the installer at build time so end users don't need to configure it.

---

## Tech stack

| Component | Library |
|-----------|---------|
| GUI | Tkinter |
| System tray | pystray + Pillow |
| Browser automation | Playwright (async Chromium) |
| Scheduling | APScheduler |
| Email | Gmail SMTP (`smtplib`) |
| Packaging | PyInstaller |
