# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

## Release workflow

Version is the single source of truth in `version.py`. To release:

```
# 1. Edit version.py → bump __version__
# 2. Build everything:
python build_installer.py
# Output: AmazonIsraelFreeShipAlert_Setup_v{VERSION}.exe  (~17.5 MB)
```

**Claude is responsible for bumping the version before every release.** Scheme: PATCH for bug fixes, MINOR for new features.

The build also regenerates `AmazonIsraelFreeShipAlert.exe` (the launcher, ~7 MB) and embeds it inside the Setup exe. Do not commit `*.exe` or `install.py` — they are generated artifacts.

---

## Architecture

### Launch chain (Windows, end-user machine)

```
AmazonIsraelFreeShipAlert.exe   ← PyInstaller launcher (no Python bundled)
  └─ python.exe gui.py          ← system Python 3.13, subprocess
       ├─ checker.py            ← Playwright async browser automation
       ├─ notifier.py           ← Gmail SMTP
       ├─ scheduler.py          ← APScheduler (runs inside a thread in gui.py)
       ├─ config.py / config.json
       └─ state.py / state.json
```

The launcher is a tiny PyInstaller exe (no app code bundled). It spawns the real Python as a subprocess so C extensions (greenlet, PIL) are always loaded by the same interpreter that `pip` installed them into — avoiding ABI mismatches.

### Why the launcher cleans the subprocess environment

The Setup installer is itself a PyInstaller bundle that sets `TCL_LIBRARY` / `TK_LIBRARY` to its own `_MEI<n>` temp dir. If those leak into `python.exe gui.py`, Tkinter cannot find `init.tcl`. The launcher strips these vars before spawning. See `LAUNCHER_SCRIPT` in `build_installer.py`.

### Checker detection logic (`checker.py`)

Uses Playwright in **async** mode (headed Chromium). The browser profile persists the Israel delivery location so only the first run calls `setup_location_once()`. Subsequent checks just navigate and read the shipping block.

Free shipping is detected when the delivery block contains all three strings:
- `"free delivery"` · `"to israel"` · `"eligible orders"`

### Installer (`build_installer.py`)

Everything is embedded in one file. At build time it:
1. Reads all source files and base64-encodes them into `install.py`
2. Packages `install.py` + a Tkinter UI into `AmazonIsraelFreeShipAlert_Setup_v{VERSION}.exe` via PyInstaller

At install time the Setup exe:
1. Extracts source files to the chosen folder
2. Downloads + silently installs **Visual C++ Redistributable 2022** (`vc_redist.x64.exe`) — required because `msvcp140.dll` (needed by greenlet's C++ extension) is absent from fresh Windows / Sandbox installs
3. Upgrades pip and installs Python packages with `--only-binary :all:`
4. Installs Chromium via `playwright install chromium`
5. Creates desktop shortcut, `Start Monitor.vbs`, and autostart registry entry

### Email (`notifier.py`)

Sender is always `amazonisraelalert@gmail.com`. The Gmail App Password is baked into `.env` at build time. Users only configure their **recipient** email in Settings. Errors raise `RuntimeError` and surface in the GUI log.

### Config defaults (written on fresh install)

`config.json` is reset by the installer to:
```json
{ "check_interval_minutes": 180, "monitoring_active": false, "products": [], "recipient": "" }
```

---

## Key design decisions to preserve

- **`--only-binary :all:`** for pip — never allow source builds; greenlet must be a compiled wheel.
- **`_find_python()` in both launcher and installer** globs `~/AppData/Local/Programs/Python/Python3*/python.exe` first, skipping `WindowsApps` stubs. PATH is checked only as fallback.
- **Exit code 0 is not an error** — when the app is already running in the tray, a new launch exits with 0. The launcher must not show an error dialog for code 0, only for non-zero codes.
- **`PLAYWRIGHT_BROWSERS_PATH`** is always set explicitly to `%LOCALAPPDATA%\ms-playwright` in `checker.py`.

---

## MANDATORY: Post-Task Documentation (SR-PTD)

**CRITICAL: After completing ANY task that modifies files, you MUST invoke this skill:**

```
Skill tool -> skill: "sr-ptd-skill"
```

**This is NOT optional. Skipping this skill means the task is INCOMPLETE.**

When planning ANY development task, add as the FINAL item in your task list:
```
[ ] Create SR-PTD documentation
```

### Before Starting Any Task:
1. Create your task plan as usual
2. Add SR-PTD documentation as the last task item
3. This step is MANDATORY for: features, bug fixes, refactors, maintenance, research

### When Completing the SR-PTD Task:
1. Read `~/.claude/skills/sr-ptd-skill/SKILL.md` for full instructions
2. Choose template: Full (complex tasks) or Quick (simple tasks)
3. Create file: `SR-PTD_YYYY-MM-DD_[task-id]_[description].md`
4. Save to: `C:/projects/Skills/Dev_doc_for_skills`
5. Fill all applicable sections thoroughly

### Task Completion Criteria:
A task is NOT complete until SR-PTD documentation exists.

### If Conversation Continues After Task:
Update the existing SR-PTD document instead of creating a new one.
