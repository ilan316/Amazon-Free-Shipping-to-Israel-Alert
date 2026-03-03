"""
Build a Windows installer for Amazon Israel Free Ship Alert.

Run this ONCE on the dev machine:
    python build_installer.py

Output:
    install.py                        — Python script installer (requires Python on target)
    AmazonIsraelFreeShipAlert_Setup.exe  — Standalone EXE installer (no Python needed to RUN it)

The EXE installer when run on any Windows PC:
    1. Asks where to install
    2. Extracts all project files
    3. Finds Python and installs pip packages
    4. Installs Chromium via playwright
    5. Creates a desktop shortcut + Start Monitor.vbs
"""

import base64
import os
import sys

PROJECT = os.path.dirname(os.path.abspath(__file__))

# Read version from version.py
def _read_version() -> str:
    ns = {}
    with open(os.path.join(PROJECT, "version.py"), encoding="utf-8") as fh:
        exec(fh.read(), ns)
    return ns.get("__version__", "0.0.0")

VERSION = _read_version()

# Project files to embed in the installer (AmazonIsraelFreeShipAlert.exe added dynamically)
INCLUDE = [
    "gui.py",
    "checker.py",
    "notifier.py",
    "config.py",
    "state.py",
    "scheduler.py",
    "version.py",
    "requirements.txt",
    "config.json",
    ".env",
    "logo-new.png",
    "icon.ico",
]

# ── Launcher script compiled into AmazonIsraelFreeShipAlert.exe ────────────
# Finds the system Python and spawns gui.py as a detached subprocess.
# This avoids ABI mismatches: C extensions (greenlet, PIL, etc.) are always
# loaded by the same Python interpreter that pip installed them into.
LAUNCHER_SCRIPT = '''
import sys, os, subprocess, shutil, glob

def _find_python():
    # Search known install locations FIRST — more reliable than PATH because:
    # 1. PATH may be stale after a fresh Python install (process env not updated)
    # 2. PATH may contain the Windows Store Python stub (WindowsApps\python.exe)
    #    which is NOT a real interpreter — it just redirects to the Store.
    home = os.path.expanduser("~")
    for pat in [
        os.path.join(home, "AppData", "Local", "Programs", "Python", "Python3*", "python.exe"),
        r"C:\\Python3*\\python.exe",
        r"C:\\Program Files\\Python3*\\python.exe",
    ]:
        hits = sorted(glob.glob(pat), reverse=True)
        for hit in hits:
            if "WindowsApps" not in hit:
                return hit
    # Fallback: PATH search — skip Microsoft Store Python stubs
    for cand in ("python", "python3", "py"):
        exe = shutil.which(cand)
        if exe and "WindowsApps" not in exe:
            return exe
    return ""

def _log(log_path, msg):
    try:
        import datetime
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"{datetime.datetime.now().strftime('%H:%M:%S')} {msg}\\n")
    except Exception:
        pass

def main():
    base = os.path.dirname(os.path.abspath(sys.executable))
    gui  = os.path.join(base, "gui.py")
    log_path = os.path.join(base, "launcher.log")

    _log(log_path, f"Launcher started. base={base}")
    _log(log_path, f"gui.py exists: {os.path.exists(gui)}")

    if not os.path.exists(gui):
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, f"gui.py not found in:\\n{base}", "Error", 0x10)
        return

    python = _find_python()
    _log(log_path, f"Python found: {python}")

    if not python:
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            0, "Python not found.\\nPlease install Python 3.11+ and re-run.",
            "Amazon Israel Free Ship Alert", 0x10)
        return

    # Use python.exe (NOT pythonw.exe) so stderr is capturable.
    # CREATE_NO_WINDOW keeps it invisible — same user experience as pythonw.
    error_log = os.path.join(base, "app_error.log")
    _log(log_path, f"error_log: {error_log}")

    try:
        import time
        # Prepend Python's own directory to PATH so the Windows DLL loader can
        # find python313.dll / vcruntime140.dll when C extensions (greenlet,
        # PIL, etc.) are imported.  Without this, a fresh machine that has never
        # had Python in its session PATH will hit:
        #   ImportError: DLL load failed while importing _greenlet:
        #   The specified module could not be found.
        env = os.environ.copy()
        python_dir = os.path.dirname(os.path.abspath(python))
        env["PATH"] = python_dir + os.pathsep + env.get("PATH", "")
        # Strip PyInstaller environment variables that confuse Python's Tkinter.
        # The Setup installer is itself a PyInstaller bundle and sets TCL_LIBRARY /
        # TK_LIBRARY pointing to its own _MEI<n> temp dir. If those are inherited
        # by python.exe gui.py, Tkinter fails:
        #   TclError: Can't find a usable init.tcl ...
        #             {C:\...\Temp\_MEI<n>\_tcl_data}  ← wrong dir, not our Python
        for _pyi_var in ("TCL_LIBRARY", "TK_LIBRARY", "TIX_LIBRARY"):
            env.pop(_pyi_var, None)
        err_f = open(error_log, "w", encoding="utf-8")
        proc = subprocess.Popen(
            [python, gui],
            cwd=base,
            env=env,
            stdout=subprocess.DEVNULL,
            stderr=err_f,
            creationflags=0x08000000,   # CREATE_NO_WINDOW
        )
        err_f.close()
        _log(log_path, f"Popen OK, pid={proc.pid}")

        # Wait 4 seconds — if the process exits that quickly it almost certainly crashed.
        time.sleep(4)
        rc = proc.poll()
        if rc is None:
            _log(log_path, "Process still running after 4 s — looks good.")
        elif rc == 0:
            # Clean exit (exit code 0 = success).
            # This is normal when the app is already running in the task tray:
            # the new instance detects the running one, brings it to the front,
            # and exits with code 0.  NOT an error — do nothing.
            _log(log_path, "Process exited with code 0 (clean exit / already running).")
        else:
            # Non-zero exit code = crash — show the error log.
            try:
                with open(error_log, "r", encoding="utf-8") as f:
                    err_content = f.read().strip()
            except Exception:
                err_content = ""
            _log(log_path, f"Process exited with rc={rc}, error={err_content[:200]}")
            import ctypes
            msg = err_content[:1200] if err_content else f"App exited with code {rc} (no error output)."
            ctypes.windll.user32.MessageBoxW(
                0, msg, "Amazon Israel Free Ship Alert — Startup Error", 0x10)

    except Exception as e:
        _log(log_path, f"Popen failed: {e}")
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, f"Launch failed:\\n{e}", "Error", 0x10)

if __name__ == "__main__":
    main()
'''

# ── Installer logic (embedded verbatim in install.py / the EXE) ─────────
INSTALLER_CODE = r'''
import base64, os, sys, subprocess, shutil
import json as _json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

DEFAULT_DIR = os.path.join(os.path.expanduser("~"), "Downloads", "amazon to israel free alert")
_NO_WIN = getattr(subprocess, "CREATE_NO_WINDOW", 0)

# Command-line args for auto-update mode (launched by the running app)
_ARGV_DIR = ""
_ARGV_AUTO_UPDATE = False
for _arg in sys.argv[1:]:
    if _arg.startswith("--dir="):
        _ARGV_DIR = _arg[6:]
    elif _arg == "--auto-update":
        _ARGV_AUTO_UPDATE = True

_UNINSTALL_PS1_TMPL = r"""
Add-Type -AssemblyName System.Windows.Forms

$app = "Amazon Free Shipping to Israel Alert"
$installDir = "__INSTALL_DIR__"
$regRun   = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$regUninst = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\AmazonFreeShippingAlert"

$ans = [System.Windows.Forms.MessageBox]::Show(
    "Are you sure you want to uninstall $app?",
    "Uninstall $app",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question)
if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { exit 0 }

# Kill running instance (launcher exe + python gui.py)
Get-Process -Name "AmazonIsraelFreeShipAlert" -ErrorAction SilentlyContinue |
    Stop-Process -Force -ErrorAction SilentlyContinue
Get-CimInstance Win32_Process -Filter "Name='python.exe'" |
    Where-Object { $_.CommandLine -like "*gui.py*" } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
Start-Sleep -Seconds 2

# Remove autostart
Remove-ItemProperty -Path $regRun -Name "AmazonFreeShippingMonitor" -ErrorAction SilentlyContinue

# Remove uninstall key
Remove-Item -Path $regUninst -Recurse -Force -ErrorAction SilentlyContinue

# Delete desktop shortcut
$lnk = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "Amazon Israel Free Ship Alert.lnk")
Remove-Item $lnk -Force -ErrorAction SilentlyContinue

# Schedule folder deletion after this script exits (cmd /c ping delays ~1 s per ping)
$cmd = "cmd.exe /c ping 127.0.0.1 -n 3 >nul & rmdir /s /q `"$installDir`""
Start-Process "cmd.exe" -ArgumentList "/c ping 127.0.0.1 -n 3 >nul & rmdir /s /q `"$installDir`"" -WindowStyle Hidden

[System.Windows.Forms.MessageBox]::Show(
    "$app has been uninstalled.",
    "Uninstall complete",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information)
"""


def _find_python():
    """Return a usable python executable path, or '' if not found."""
    if not getattr(sys, "frozen", False):
        return sys.executable
    # Search known install locations FIRST — more reliable than PATH because:
    # 1. PATH may be stale after a fresh Python install (process env not updated)
    # 2. PATH may contain the Windows Store Python stub (WindowsApps\python.exe)
    #    which is NOT a real interpreter — it just redirects to the Store.
    import glob
    home = os.path.expanduser("~")
    for pattern in [
        os.path.join(home, "AppData", "Local", "Programs", "Python", "Python3*", "python.exe"),
        r"C:\Python3*\python.exe",
        r"C:\Program Files\Python3*\python.exe",
    ]:
        hits = sorted(glob.glob(pattern), reverse=True)
        for hit in hits:
            if "WindowsApps" not in hit:
                return hit
    # Fallback: PATH search — skip Microsoft Store Python stubs
    for cand in ("python", "python3", "py"):
        exe = shutil.which(cand)
        if exe and "WindowsApps" not in exe:
            return exe
    return ""


class InstallerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Amazon Israel Free Ship Alert v{VERSION} — Setup")
        self.resizable(False, False)
        self.geometry("560x470")
        self._installing = False
        self._launcher_path = ""
        self._app_exe_path = ""
        self._build_ui()
        if _ARGV_DIR:
            self._dir_var.set(_ARGV_DIR)
        else:
            self._dir_var.set(InstallerApp._detect_install_dir())
        try:
            _icon = tk.PhotoImage(data=ICON_B64)
            self.iconphoto(True, _icon)
            self._icon_ref = _icon
        except Exception:
            pass
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        if _ARGV_AUTO_UPDATE:
            self.after(300, self._start_install)  # auto-start install (no user interaction needed)

    # ── UI ─────────────────────────────────────────────────────

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg="white", pady=4)
        hdr.pack(fill=tk.X)
        _logo_shown = False
        if LOGO_B64:
            try:
                _logo = tk.PhotoImage(data=LOGO_B64)
                _lbl = tk.Label(hdr, image=_logo, bg="white")
                _lbl.image = _logo  # keep reference
                _lbl.pack()
                _logo_shown = True
            except Exception:
                pass
        if not _logo_shown:
            tk.Label(hdr, text="Amazon Israel Free Ship Alert",
                     font=("Segoe UI", 14, "bold"),
                     bg="white", fg="#1a1a2e").pack()
            tk.Label(hdr, text=f"Setup Wizard  •  v{VERSION}",
                     font=("Segoe UI", 10), bg="white", fg="#555555").pack()

        # Bottom bar — packed BEFORE body so expand=True on body doesn't hide it
        bot = tk.Frame(self, padx=20, pady=10)
        bot.pack(side=tk.BOTTOM, fill=tk.X)
        self._install_btn = tk.Button(
            bot, text="  Install  ",
            bg="#0066cc", fg="white", relief=tk.FLAT,
            font=("Segoe UI", 10, "bold"),
            padx=20, pady=6, cursor="hand2",
            command=self._start_install)
        self._install_btn.pack(side=tk.LEFT)
        self._status_var = tk.StringVar(value="Ready.")
        tk.Label(bot, textvariable=self._status_var,
                 font=("Segoe UI", 8), fg="#555555").pack(side=tk.LEFT, padx=12)

        body = tk.Frame(self, padx=20, pady=8)
        body.pack(fill=tk.BOTH, expand=True)

        lbl  = {"font": ("Segoe UI", 9, "bold"), "anchor": "w"}
        hint = {"font": ("Segoe UI", 8), "fg": "#777777", "anchor": "w"}
        ent  = {"font": ("Segoe UI", 10), "relief": tk.SOLID, "bd": 1}

        # Install dir
        tk.Label(body, text="Install folder", **lbl).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(10, 1))
        frm = tk.Frame(body)
        frm.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(2, 0))
        self._dir_var = tk.StringVar(value="")
        tk.Entry(frm, textvariable=self._dir_var, width=44, **ent).pack(
            side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(frm, text=" Browse\u2026 ", command=self._browse,
                  font=("Segoe UI", 8), relief=tk.FLAT,
                  bg="#e8e8e8", cursor="hand2", pady=4).pack(side=tk.LEFT, padx=(4, 0))

        body.columnconfigure(0, weight=1)

        # Progress log
        tk.Label(body, text="Progress", **lbl).grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(12, 1))
        log_frm = tk.Frame(body)
        log_frm.grid(row=3, column=0, columnspan=2, sticky="nsew")
        body.rowconfigure(3, weight=1)

        self._log_txt = tk.Text(
            log_frm, height=9, state=tk.DISABLED,
            font=("Consolas", 8), bg="#1e1e1e", fg="#d4d4d4",
            relief=tk.FLAT, padx=6, pady=4, wrap=tk.WORD)
        sb = ttk.Scrollbar(log_frm, orient="vertical", command=self._log_txt.yview)
        self._log_txt.configure(yscrollcommand=sb.set)
        self._log_txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.LEFT, fill=tk.Y)

        # "Launch after install" checkbox
        self._launch_var = tk.BooleanVar(value=True)
        self._launch_cb = tk.Checkbutton(
            body, text="Launch Amazon Israel Free Ship Alert after installation",
            variable=self._launch_var, font=("Segoe UI", 9))
        self._launch_cb.grid(row=4, column=0, columnspan=2, sticky="w", pady=(6, 0))

    def _browse(self):
        init = self._dir_var.get() or os.path.expanduser("~")
        d = filedialog.askdirectory(initialdir=init, title="Choose install folder")
        if d:
            self._dir_var.set(d)

    @staticmethod
    def _detect_install_dir() -> str:
        """Return the existing install directory from registry, or DEFAULT_DIR."""
        import winreg, re as _re
        try:
            with winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    r"Software\Microsoft\Windows\CurrentVersion\Run") as k:
                cmd = winreg.QueryValueEx(k, "AmazonFreeShippingMonitor")[0]
                sysroot = os.environ.get("SystemRoot", r"C:\Windows").lower()
                for m in _re.findall(r'"([^"]+)"', cmd):
                    if m.lower().endswith((".exe", ".vbs")):
                        d = os.path.dirname(m)
                        if d.lower().startswith(sysroot):
                            continue
                        if os.path.isdir(d):
                            return d
        except Exception:
            pass
        if os.path.exists(os.path.join(DEFAULT_DIR, "config.json")):
            return DEFAULT_DIR
        return DEFAULT_DIR

    # ── Thread-safe UI helpers ──────────────────────────────────

    def _log(self, msg: str):
        self.after(0, self._do_log, msg)

    def _do_log(self, msg: str):
        self._log_txt.configure(state=tk.NORMAL)
        self._log_txt.insert(tk.END, msg + "\n")
        self._log_txt.see(tk.END)
        self._log_txt.configure(state=tk.DISABLED)

    def _status(self, msg: str):
        self.after(0, self._status_var.set, msg)

    # ── Install ─────────────────────────────────────────────────

    def _start_install(self):
        install_dir = self._dir_var.get().strip()
        if not install_dir:
            messagebox.showerror("Invalid folder",
                                 "Please choose an install folder.", parent=self)
            return
        python_exe = _find_python()
        if not python_exe:
            if messagebox.askyesno(
                    "Python not found",
                    "Python 3.11+ is required but was not found on this PC.\n\n"
                    "Download and install Python 3.13 automatically?\n"
                    "(\u226425 MB \u2014 requires internet connection)",
                    parent=self):
                self._install_btn.configure(state=tk.DISABLED, text="Installing\u2026")
                self._installing = True
                threading.Thread(
                    target=self._install_python_then_app,
                    args=(install_dir,),
                    daemon=True).start()
            return
        self._install_btn.configure(state=tk.DISABLED, text="Installing\u2026")
        self._installing = True
        threading.Thread(
            target=self._do_install,
            args=(install_dir, python_exe),
            daemon=True).start()

    def _install_python_then_app(self, install_dir: str):
        python_exe = self._download_and_install_python()
        if not python_exe:
            self._log("\nERROR: Python installation failed or was cancelled.")
            self._status("Python install failed.")
            self.after(0, lambda: self._install_btn.configure(
                state=tk.NORMAL, text="  Install  "))
            self._installing = False
            return
        self._do_install(install_dir, python_exe)

    def _download_and_install_python(self) -> str:
        """Download Python 3.12 installer, run it silently, return python.exe path."""
        import urllib.request, urllib.error, tempfile, glob, ssl
        url = "https://www.python.org/ftp/python/3.13.12/python-3.13.12-amd64.exe"
        self._status("Downloading Python 3.13\u2026")
        self._log("\nDownloading Python 3.13 (\u226425 MB)\u2026")
        tmp = ""
        try:
            fd, tmp = tempfile.mkstemp(suffix=".exe", prefix="py_setup_")
            os.close(fd)

            def _reporthook(count, block, total):
                if total > 0:
                    pct = min(100, count * block * 100 // total)
                    self._status(f"Downloading Python 3.13\u2026  {pct}%")

            # Use an unverified SSL context to handle environments (e.g. Windows Sandbox)
            # where the system certificate store may not be fully initialised.
            _ctx = ssl.create_default_context()
            _ctx.check_hostname = False
            _ctx.verify_mode = ssl.CERT_NONE
            _opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=_ctx))
            urllib.request.install_opener(_opener)
            urllib.request.urlretrieve(url, tmp, _reporthook)
            self._log("  Download complete.")
        except Exception as exc:
            self._log(f"  Download failed: {exc}")
            if tmp and os.path.exists(tmp):
                os.unlink(tmp)
            return ""

        self._status("Installing Python 3.13\u2026")
        self._log("Installing Python 3.13 (this may take a minute)\u2026")
        try:
            ret = subprocess.call(
                [tmp, "/passive", "InstallAllUsers=0",
                 "PrependPath=1", "Include_test=0"])
        except Exception as exc:
            self._log(f"  Installer error: {exc}")
            ret = -1
        finally:
            try:
                os.unlink(tmp)
            except Exception:
                pass

        if ret != 0:
            self._log(f"  Python installer exited with code {ret}.")
            return ""

        self._log("  Python 3.13 installed.  Searching for python.exe\u2026")
        # PATH in this process is stale; rely on known install locations
        home = os.path.expanduser("~")
        for pattern in [
            os.path.join(home, "AppData", "Local", "Programs", "Python", "Python3*", "python.exe"),
            r"C:\Python3*\python.exe",
            r"C:\Program Files\Python3*\python.exe",
        ]:
            hits = sorted(glob.glob(pattern), reverse=True)
            if hits:
                self._log(f"  Found: {hits[0]}")
                return hits[0]
        self._log("  Could not find python.exe after install.")
        return ""

    def _run_logged(self, cmd):
        """Run a subprocess and stream its output to the progress log in real-time."""
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace",
            creationflags=_NO_WIN, bufsize=1,
        )
        for line in proc.stdout:
            line = line.rstrip()
            if line:
                self._log(f"  {line}")
        proc.wait()
        if proc.returncode != 0:
            raise subprocess.CalledProcessError(proc.returncode, cmd)

    def _install_vcredist(self):
        """Download and silently install Visual C++ Redistributable 2015-2022 x64.

        greenlet (and other C++ extensions) require msvcp140.dll at runtime.
        Python only bundles vcruntime140.dll — NOT msvcp140.dll.
        A fresh Windows / Sandbox installation lacks msvcp140.dll in System32,
        causing: ImportError: DLL load failed while importing _greenlet.
        """
        import urllib.request, tempfile, ssl
        url = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
        self._status("Installing Visual C++ Redistributable\u2026")
        self._log("\nInstalling Visual C++ Redistributable 2015-2022 x64\u2026")
        self._log("  (Provides msvcp140.dll required by greenlet / Playwright)")
        tmp = ""
        try:
            fd, tmp = tempfile.mkstemp(suffix=".exe", prefix="vcredist_")
            os.close(fd)

            _ctx = ssl.create_default_context()
            _ctx.check_hostname = False
            _ctx.verify_mode = ssl.CERT_NONE
            _opener = urllib.request.build_opener(
                urllib.request.HTTPSHandler(context=_ctx))
            urllib.request.install_opener(_opener)

            def _hook(count, block, total):
                if total > 0:
                    pct = min(100, count * block * 100 // total)
                    self._status(f"Downloading VC++ Redist\u2026 {pct}%")

            urllib.request.urlretrieve(url, tmp, _hook)
            self._log("  Download complete. Installing silently\u2026")
            self._status("Installing Visual C++ Redistributable\u2026")

            ret = subprocess.call(
                [tmp, "/install", "/quiet", "/norestart"],
                creationflags=_NO_WIN,
            )
            # 0 = success
            # 1638 = already installed (a newer version is present) — fine
            # 3010 = success but a reboot is recommended
            if ret in (0, 1638):
                self._log("  Visual C++ Redistributable: OK.")
            elif ret == 3010:
                self._log("  Visual C++ Redistributable: installed (reboot may be needed).")
            else:
                self._log(f"  Visual C++ Redistributable: installer exited with code {ret}.")
        except Exception as exc:
            self._log(f"  WARNING: VC++ Redistributable install failed: {exc}")
            self._log("  The app may not start if msvcp140.dll is missing on this system.")
        finally:
            try:
                if tmp and os.path.exists(tmp):
                    os.unlink(tmp)
            except Exception:
                pass

    def _close_running_app(self):
        """Close the running app (including from tray) before installing.

        The launcher exe (AmazonIsraelFreeShipAlert.exe) spawns python.exe
        and exits after ~4 seconds, so the real long-running process is
        python.exe gui.py.  We find it by command-line pattern and kill it.
        """
        import time
        self._log("\nChecking for a running instance of the app...")
        killed = False

        # Primary: wmic — find python.exe with gui.py in its command line
        try:
            out = subprocess.check_output(
                ["wmic", "process", "where",
                 "name='python.exe' and CommandLine like '%gui.py%'",
                 "get", "ProcessId"],
                creationflags=_NO_WIN, text=True, encoding="utf-8",
                errors="ignore", stderr=subprocess.DEVNULL, timeout=8,
            )
            pids = [ln.strip() for ln in out.splitlines() if ln.strip().isdigit()]
            for pid in pids:
                ret = subprocess.call(
                    ["taskkill", "/F", "/PID", pid],
                    creationflags=_NO_WIN,
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                )
                if ret == 0:
                    self._log(f"  App process (PID {pid}) terminated.")
                    killed = True
        except Exception:
            # wmic not available — try PowerShell fallback
            try:
                _ps = os.path.join(
                    os.environ.get("SystemRoot", r"C:\Windows"),
                    "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
                _cmd = (
                    "Get-WmiObject Win32_Process -Filter \"name='python.exe'\" | "
                    "Where-Object {$_.CommandLine -like '*gui.py*'} | "
                    "ForEach-Object {Stop-Process -Id $_.ProcessId -Force "
                    "-ErrorAction SilentlyContinue}"
                )
                subprocess.call(
                    [_ps, "-NoProfile", "-NonInteractive",
                     "-ExecutionPolicy", "Bypass", "-Command", _cmd],
                    creationflags=_NO_WIN, timeout=15,
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                )
                killed = True
            except Exception:
                pass

        if killed:
            time.sleep(2)  # Let file handles release
        else:
            self._log("  No running instance found.")

    def _do_install(self, install_dir: str, python_exe: str):
        try:
            # Normalize path (filedialog returns forward-slashes on Windows)
            install_dir = os.path.normpath(install_dir)

            # Close any running instance so files can be safely overwritten
            self._close_running_app()

            # Save existing user config BEFORE extracting files so we can
            # restore products, email, interval, etc. after the update.
            _cfg_path_pre = os.path.join(install_dir, "config.json")
            _existing_cfg = {}
            if os.path.exists(_cfg_path_pre):
                try:
                    with open(_cfg_path_pre, "r", encoding="utf-8") as _fh:
                        _existing_cfg = _json.load(_fh)
                    _n = len(_existing_cfg.get("products", []))
                    if _n:
                        self._log(f"Existing install detected — preserving {_n} product(s) and settings.\n")
                    else:
                        self._log("Existing install detected — preserving settings.\n")
                except Exception:
                    _existing_cfg = {}

            # Extract files
            self._status("Extracting files\u2026")
            self._log(f"Installing to: {install_dir}\n")
            os.makedirs(install_dir, exist_ok=True)
            self._log("Extracting files\u2026")
            for name, b64 in FILES.items():
                dest = os.path.join(install_dir, name)
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                with open(dest, "wb") as fh:
                    fh.write(base64.b64decode(b64))
                self._log(f"  {name}")

            # Patch config.json — restore user data if this is an update,
            # or keep fresh defaults if this is a brand-new install.
            cfg_path = os.path.join(install_dir, "config.json")
            try:
                with open(cfg_path, "r", encoding="utf-8") as fh:
                    cfg = _json.load(fh)
                # Preserve monitoring_active from previous install so monitoring
                # resumes automatically after an update if it was running before.
                cfg["monitoring_active"] = bool(_existing_cfg.get("monitoring_active", False))
                # Restore user data from previous install (if any)
                if _existing_cfg.get("products"):
                    cfg["products"] = _existing_cfg["products"]
                    self._log(f"  Restored {len(cfg['products'])} product(s) from previous install.")
                else:
                    cfg["products"] = []
                if "check_interval_minutes" in _existing_cfg:
                    cfg["check_interval_minutes"] = _existing_cfg["check_interval_minutes"]
                if "notification_cooldown_hours" in _existing_cfg:
                    cfg["notification_cooldown_hours"] = _existing_cfg["notification_cooldown_hours"]
                if _existing_cfg.get("email", {}).get("recipient"):
                    cfg.setdefault("email", {})["recipient"] = _existing_cfg["email"]["recipient"]
                if "language" in _existing_cfg:
                    cfg["language"] = _existing_cfg["language"]
                with open(cfg_path, "w", encoding="utf-8") as fh:
                    _json.dump(cfg, fh, indent=2, ensure_ascii=False)
                self._log("  config.json  (settings preserved)")
            except Exception as exc:
                self._log(f"  WARNING: config.json: {exc}")

            # Install VC++ Redistributable BEFORE pip packages.
            # greenlet is a C++ extension and requires msvcp140.dll at runtime.
            # Python only bundles vcruntime140.dll — msvcp140.dll is NOT included.
            # Without VC++ Redist, importing greenlet/Playwright fails on a fresh
            # Windows machine with: "DLL load failed while importing _greenlet".
            self._install_vcredist()

            # upgrade pip first
            self._status("Upgrading pip\u2026")
            self._log("\nUpgrading pip\u2026")
            self._run_logged([python_exe, "-m", "pip", "install", "--upgrade", "pip"])

            # pip install packages
            # --only-binary :all:  → refuse source distributions entirely; binary wheels only.
            #   This is essential for greenlet: without a compiled C extension, Playwright fails.
            # --no-cache-dir      → never use a stale or corrupt cached package.
            self._status("Installing Python packages\u2026")
            self._log("\nInstalling Python packages\u2026")
            self._run_logged([
                python_exe, "-m", "pip", "install", "--upgrade",
                "--only-binary", ":all:",
                "--no-cache-dir",
                "greenlet",
                "playwright>=1.49.0", "apscheduler==3.10.4", "python-dotenv==1.0.1",
                "pystray>=0.19.5", "Pillow>=10.0.0",
            ])
            self._log("  Packages installed.")

            # playwright install chromium
            self._status("Installing Chromium browser (may take a few minutes)\u2026")
            self._log("\nInstalling Chromium browser\u2026")
            self._run_logged([python_exe, "-m", "playwright", "install", "chromium"])
            self._log("  Chromium ready.")

            # Resolve paths — use full paths so nothing depends on PATH
            _python_dir = os.path.dirname(os.path.abspath(python_exe))
            _pythonw = os.path.join(_python_dir, "pythonw.exe")
            if not os.path.exists(_pythonw):
                _pythonw = python_exe  # fallback: python.exe
            _gui_path  = os.path.join(install_dir, "gui.py")
            _app_exe   = os.path.join(install_dir, "AmazonIsraelFreeShipAlert.exe")
            _use_exe   = os.path.isfile(_app_exe)
            if _use_exe:
                self._app_exe_path = _app_exe

            # Start Monitor.vbs — only created as fallback when no exe launcher is present.
            # When the exe exists it is the launcher; no VBS needed (VBS triggers AV false positives).
            launcher = ""
            if not _use_exe:
                _vbs_launcher = (
                    'Set WshShell = CreateObject("WScript.Shell")\r\n'
                    f'WshShell.Run Chr(34) & "{_pythonw}" & Chr(34) & " " & Chr(34) & "{_gui_path}" & Chr(34), 0, False\r\n'
                    'Set WshShell = Nothing\r\n'
                )
                launcher = os.path.join(install_dir, "Start Monitor.vbs")
                with open(launcher, "w", encoding="utf-8") as fh:
                    fh.write(_vbs_launcher)
                self._launcher_path = launcher

            # Desktop shortcut — via PowerShell (works even when VBS is disabled)
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            if os.path.isdir(desktop):
                shortcut_lnk = os.path.join(desktop, "Amazon Israel Free Ship Alert.lnk")
                _icon_file   = os.path.join(install_dir, "icon.ico")
                if _use_exe:
                    _target = _app_exe
                    _args   = ""
                else:
                    _target = _pythonw
                    _args   = f'"{_gui_path}"'
                _ps = (
                    f'$s=(New-Object -COM WScript.Shell).CreateShortcut("{shortcut_lnk}");'
                    f'$s.TargetPath="{_target}";'
                    f'$s.Arguments="{_args}";'
                    f'$s.WorkingDirectory="{install_dir}";'
                    f'$s.IconLocation="{_icon_file},0";'
                    f'$s.Save()'
                )
                _powershell = os.path.join(
                    os.environ.get("SystemRoot", r"C:\Windows"),
                    "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
                try:
                    subprocess.call(
                        [_powershell, "-NoProfile", "-NonInteractive",
                         "-ExecutionPolicy", "Bypass", "-Command", _ps],
                        creationflags=_NO_WIN, timeout=30)
                    self._log(f"\nDesktop shortcut: {shortcut_lnk}")
                except subprocess.TimeoutExpired:
                    self._log("\nDesktop shortcut: skipped (timed out).")
                except Exception as _sc_err:
                    self._log(f"\nDesktop shortcut: skipped ({_sc_err}).")

            # Enable autostart by default
            try:
                import winreg as _wr
                _cmd = f'"{_app_exe}"' if _use_exe else f'"{_pythonw}" "{_gui_path}"'
                with _wr.OpenKey(_wr.HKEY_CURRENT_USER,
                                 r"Software\Microsoft\Windows\CurrentVersion\Run",
                                 0, _wr.KEY_SET_VALUE) as _k:
                    _wr.SetValueEx(_k, "AmazonFreeShippingMonitor", 0, _wr.REG_SZ, _cmd)
                self._log("  Autostart on Windows login: enabled.")
            except Exception as _ae:
                self._log(f"  Could not enable autostart: {_ae}")

            # Register in Windows "Installed Apps" (Add/Remove Programs)
            try:
                import winreg as _wr
                _unreg_key = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\AmazonFreeShippingAlert"
                _ps1_path  = os.path.join(install_dir, "_uninstall.ps1")
                _ps1_content = _UNINSTALL_PS1_TMPL.replace(
                    "__INSTALL_DIR__", install_dir.replace("\\", "\\\\"))
                with open(_ps1_path, "w", encoding="utf-8") as _pf:
                    _pf.write(_ps1_content)
                _powershell_exe = os.path.join(
                    os.environ.get("SystemRoot", r"C:\Windows"),
                    "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
                _ucmd = (f'"{_powershell_exe}" -ExecutionPolicy Bypass'
                         f' -File "{_ps1_path}"')
                with _wr.CreateKey(_wr.HKEY_CURRENT_USER, _unreg_key) as _k:
                    _wr.SetValueEx(_k, "DisplayName",     0, _wr.REG_SZ,    "Amazon Free Shipping to Israel Alert")
                    _wr.SetValueEx(_k, "DisplayVersion",  0, _wr.REG_SZ,    VERSION)
                    _wr.SetValueEx(_k, "Publisher",       0, _wr.REG_SZ,    "amzfreeil.com")
                    _wr.SetValueEx(_k, "InstallLocation", 0, _wr.REG_SZ,    install_dir)
                    _wr.SetValueEx(_k, "UninstallString", 0, _wr.REG_SZ,    _ucmd)
                    _wr.SetValueEx(_k, "DisplayIcon",     0, _wr.REG_SZ,    _app_exe + ",0")
                    _wr.SetValueEx(_k, "NoModify",        0, _wr.REG_DWORD, 1)
                    _wr.SetValueEx(_k, "NoRepair",        0, _wr.REG_DWORD, 1)
                self._log("  Registered in Windows Installed Apps.")
            except Exception as _re:
                self._log(f"  Could not register in Installed Apps: {_re}")

            self._log("\n" + "=" * 48)
            self._log("  Installation complete!")
            self._log(f"  Folder   :  {install_dir}")
            self._log(f"  Launcher :  {launcher or _app_exe}")
            self._log("=" * 48)
            self._status("Installation complete!")
            self.after(0, self._complete)

        except Exception as exc:
            self._log(f"\n\nERROR: {exc}")
            self._status(f"Failed \u2014 {exc}")
            self.after(0, lambda: self._install_btn.configure(
                state=tk.NORMAL, text="Retry"))
        finally:
            self._installing = False

    def _complete(self):
        self._install_btn.configure(
            state=tk.DISABLED, text="  Done  ", bg="#1a7a1a")
        _launch_now = self._launch_var.get()
        _launcher   = getattr(self, "_launcher_path", "")
        _app_exe    = getattr(self, "_app_exe_path", "")
        messagebox.showinfo(
            "Installation complete",
            "Amazon Israel Free Ship Alert installed successfully!\n\n"
            "Launch:\n"
            "  Use the Desktop shortcut.\n\n"
            "The app will also start automatically with Windows.",
            parent=self)
        self.destroy()
        if _launch_now:
            if _app_exe and os.path.exists(_app_exe):
                subprocess.Popen([_app_exe], creationflags=_NO_WIN)
            elif _launcher and os.path.exists(_launcher):
                _wscript = os.path.join(
                    os.environ.get("SystemRoot", r"C:\Windows"), "System32", "wscript.exe")
                subprocess.Popen([_wscript, _launcher], creationflags=_NO_WIN)

    def _on_close(self):
        if self._installing:
            if not messagebox.askyesno(
                    "Cancel?",
                    "Installation is in progress. Quit anyway?",
                    parent=self):
                return
        self.destroy()


def main():
    app = InstallerApp()
    app.mainloop()
'''


# ── Build helpers ────────────────────────────────────────────────────────

def _encode_files():
    encoded = {}
    for name in INCLUDE:
        path = os.path.join(PROJECT, name)
        if os.path.exists(path):
            with open(path, "rb") as fh:
                data = fh.read()
            encoded[name] = base64.b64encode(data).decode("ascii")
            print(f"  + {name}  ({len(data):,} bytes)")
        else:
            print(f"  - {name}  (not found — skipping)")
    return encoded


def _build_install_py(encoded: dict, logo_b64: str = "", icon_b64: str = "") -> str:
    """Return the full source of install.py."""
    lines = ['"""Amazon Israel Free Ship Alert — Installer\n'
             'Run:  python install.py\n'
             'Requires Python 3.11+ on the target machine.\n'
             '"""\n\n']

    # FILES dict
    lines.append("FILES = {\n")
    for name, b64 in encoded.items():
        lines.append(f"    {name!r}: {b64!r},\n")
    lines.append("}\n\n")

    # Logo (pre-resized PNG, base64-encoded — used by the installer UI header)
    lines.append(f"LOGO_B64 = {logo_b64!r}\n")

    # Icon (small square PNG, base64-encoded — used for iconphoto in the installer window)
    lines.append(f"ICON_B64 = {icon_b64!r}\n")

    # Version constant (baked in at build time so INSTALLER_CODE can reference it)
    lines.append(f"VERSION = {VERSION!r}\n")

    # The installer logic
    lines.append(INSTALLER_CODE)
    lines.append('\nif __name__ == "__main__":\n    main()\n')
    return "".join(lines)


def _write_install_py(encoded: dict, logo_b64: str = "", icon_b64: str = "") -> str:
    code = _build_install_py(encoded, logo_b64, icon_b64)
    out = os.path.join(PROJECT, "install.py")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(code)
    kb = os.path.getsize(out) // 1024
    print(f"\nGenerated: {out}  ({kb} KB)")
    return out


def _build_exe(install_py_path: str, icon_path: str = ""):
    """Bundle install.py into a standalone .exe using PyInstaller."""
    try:
        import PyInstaller.__main__ as pyi
    except ImportError:
        print("\nPyInstaller not installed — skipping .exe build.")
        print("To build the .exe:  pip install pyinstaller  then re-run.")
        return

    build_tmp  = os.path.join(PROJECT, "_build_installer_tmp")
    setup_name = "AmazonIsraelFreeShipAlert"
    print(f"\nBuilding {setup_name}.exe (installer) ...")
    args = [
        "--onefile",
        "--name",      setup_name,
        "--windowed",                      # GUI only — no console window
        "--distpath",  PROJECT,            # put the .exe in the project folder
        "--workpath",  build_tmp,
        "--specpath",  build_tmp,
        "--noconfirm",
    ]
    if icon_path and os.path.exists(icon_path):
        args.extend(["--icon", icon_path])
    args.append(install_py_path)
    pyi.run(args)

    exe = os.path.join(PROJECT, f"{setup_name}.exe")
    if os.path.exists(exe):
        mb = os.path.getsize(exe) / (1024 * 1024)
        print(f"\nInstaller EXE: {exe}  ({mb:.1f} MB)")
        print(f"\nSend  {setup_name}.exe  to the new computer.")
        print("Double-click it — no Python needed to run the installer itself.")
    else:
        print("\nWARNING: exe not found after build. Check PyInstaller output above.")


def _make_logo_b64(width: int = 420, height: int = 173) -> str:
    """Return base64-encoded resized PNG of the logo, or '' on failure."""
    try:
        from PIL import Image as _Img
        import io as _io
        path = os.path.join(PROJECT, "logo-new.png")
        with _Img.open(path) as img:
            img = img.resize((width, height), _Img.LANCZOS)
            buf = _io.BytesIO()
            img.save(buf, format="PNG")
            return base64.b64encode(buf.getvalue()).decode("ascii")
    except Exception as exc:
        print(f"  Warning: could not encode logo: {exc}")
        return ""


def _make_icon() -> str:
    """
    Create icon.ico (orange circle + white checkmark) in the project folder.
    Returns the path to the .ico file, or '' on failure.
    """
    try:
        from PIL import Image, ImageDraw
        size = 256
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        d.ellipse([0, 0, size - 1, size - 1], fill="#FF9900")
        # white checkmark — same proportions as the tray icon
        s = size
        pts = [
            (int(s * 0.23), int(s * 0.53)),
            (int(s * 0.42), int(s * 0.72)),
            (int(s * 0.78), int(s * 0.31)),
        ]
        d.line(pts, fill="white", width=int(s * 0.11))
        out = os.path.join(PROJECT, "icon.ico")
        img.save(out, format="ICO",
                 sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
        print(f"  Created: icon.ico")
        return out
    except Exception as exc:
        print(f"  Warning: could not create icon.ico: {exc}")
        return ""


def _make_icon_b64(size: int = 64) -> str:
    """Return base64-encoded PNG of the app icon (small square for iconphoto)."""
    try:
        from PIL import Image, ImageDraw
        import io as _io
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        d.ellipse([0, 0, size - 1, size - 1], fill="#FF9900")
        s = size
        pts = [
            (int(s * 0.23), int(s * 0.53)),
            (int(s * 0.42), int(s * 0.72)),
            (int(s * 0.78), int(s * 0.31)),
        ]
        d.line(pts, fill="white", width=max(1, int(s * 0.11)))
        buf = _io.BytesIO()
        img.save(buf, format="PNG")
        return base64.b64encode(buf.getvalue()).decode("ascii")
    except Exception as exc:
        print(f"  Warning: could not encode icon PNG: {exc}")
        return ""


def _build_launcher_exe(icon_path: str) -> str:
    """
    Compile LAUNCHER_SCRIPT into AmazonIsraelFreeShipAlert.exe with the custom icon
    embedded in the PE.  The exe runs gui.py in-process, so the taskbar inherits
    the exe's icon — no Win32 API tricks needed.
    """
    try:
        import PyInstaller.__main__ as pyi
    except ImportError:
        print("  PyInstaller not found — launcher exe skipped.")
        return ""

    import tempfile
    tmp_dir  = os.path.join(PROJECT, "_build_launcher_tmp")
    launcher = os.path.join(tmp_dir, "_launcher.py")
    os.makedirs(tmp_dir, exist_ok=True)
    with open(launcher, "w", encoding="utf-8") as fh:
        fh.write(LAUNCHER_SCRIPT)

    print("\nBuilding AmazonIsraelFreeShipAlert.exe (launcher) ...")
    args = [
        "--onefile",
        "--windowed",
        "--name",     "AmazonIsraelFreeShipAlert",
        "--distpath", PROJECT,
        "--workpath", tmp_dir,
        "--specpath", tmp_dir,
        "--noconfirm",
        # The launcher only needs ctypes (for MessageBoxW error dialogs).
        # All other packages (playwright, PIL, greenlet, etc.) are loaded by
        # the system Python subprocess — no bundling needed here.
        "--hidden-import", "ctypes",
        "--hidden-import", "ctypes.wintypes",
    ]
    if icon_path and os.path.exists(icon_path):
        args.extend(["--icon", icon_path])
    args.append(launcher)
    pyi.run(args)

    exe = os.path.join(PROJECT, "AmazonIsraelFreeShipAlert.exe")
    if os.path.exists(exe):
        mb = os.path.getsize(exe) / (1024 * 1024)
        print(f"  Created: AmazonIsraelFreeShipAlert.exe ({mb:.1f} MB)")
        return exe
    print("  WARNING: AmazonIsraelFreeShipAlert.exe not found after build.")
    return ""


def main():
    print("Creating icon...")
    icon_path = _make_icon()

    print("Building launcher exe with embedded icon...")
    app_exe = _build_launcher_exe(icon_path)
    if app_exe:
        INCLUDE.insert(0, "AmazonIsraelFreeShipAlert.exe")
    else:
        print("  Launcher exe not built — will fall back to pythonw at install time.")

    print("\nReading project files...")
    encoded = _encode_files()

    if not encoded:
        print("\nERROR: No source files found.")
        return

    print("Encoding logo...")
    logo_b64 = _make_logo_b64()

    print("Encoding icon PNG...")
    icon_b64 = _make_icon_b64()

    install_py = _write_install_py(encoded, logo_b64, icon_b64)
    _build_exe(install_py, icon_path)

    print()
    print("Done.")


if __name__ == "__main__":
    main()
