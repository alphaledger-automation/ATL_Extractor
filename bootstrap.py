"""
ATL Extractor — Self-Setup (Bootstrap)
======================================
Makes the app install everything it needs by itself:

  Missing Python packages (customtkinter, openpyxl, pdfplumber, pillow)
  are installed with pip.

If anything needs downloading, a small progress window is shown (plain
tkinter, so it works even before customtkinter is installed). If
everything is already in place this module returns almost instantly and
shows nothing.

Every step is written to  <user-data-dir>/setup.log  for troubleshooting.
"""

from __future__ import annotations

import importlib
import importlib.util
import os
import pathlib
import queue
import site
import subprocess
import sys
import threading
import traceback
from typing import Callable, List, Optional

APP_TITLE    = "ATL Extractor"
APP_DIR_NAME = "ATL_Extractor"

# import-name -> pip-name. Used to decide what is missing.
REQUIRED_MODULES = {
    "customtkinter": "customtkinter",
    "openpyxl": "openpyxl",
    "pdfplumber": "pdfplumber",
    "PIL": "pillow",
}

_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)   # 0 on non-Windows

LogFn = Callable[[str], None]


# ── Paths ─────────────────────────────────────────────────────────────────────

def _data_dir() -> pathlib.Path:
    """Per-user writable folder for the setup log."""
    if sys.platform == "win32":
        base = pathlib.Path(os.environ.get("LOCALAPPDATA", pathlib.Path.home()))
    else:
        base = pathlib.Path.home() / ".local" / "share"
    d = base / APP_DIR_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d


def _app_dir() -> pathlib.Path:
    """Folder containing requirements.txt (next to main.py / this file)."""
    return pathlib.Path(__file__).resolve().parent


def log_file() -> pathlib.Path:
    return _data_dir() / "setup.log"


# ── Detection ─────────────────────────────────────────────────────────────────

def missing_python_packages() -> List[str]:
    """Return pip names of required packages that can't be imported."""
    return [
        pip_name for mod, pip_name in REQUIRED_MODULES.items()
        if importlib.util.find_spec(mod) is None
    ]


# ── Installation ──────────────────────────────────────────────────────────────

def _run_streaming(cmd: List[str], log: LogFn, env: Optional[dict] = None) -> int:
    """Run a command, forwarding each output line to `log`. Returns exit code."""
    log("$ " + " ".join(cmd))
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        stdin=subprocess.DEVNULL,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
        creationflags=_NO_WINDOW,
    )
    assert proc.stdout is not None
    for line in proc.stdout:            # text mode: '\r' progress ticks split too
        line = line.strip()
        if line:
            log(line)
    return proc.wait()


def _add_user_site_to_path() -> None:
    """After a `pip --user` install the user site may not be on sys.path yet."""
    try:
        us = site.getusersitepackages()
        if us and os.path.isdir(us) and us not in sys.path:
            sys.path.append(us)
    except Exception:
        pass
    importlib.invalidate_caches()


def install_python_packages(log: LogFn) -> None:
    """pip-install requirements.txt (falls back to --user / ensurepip)."""
    req = _app_dir() / "requirements.txt"
    if not req.exists():
        # requirements.txt missing -> install the known package list directly
        targets = list(REQUIRED_MODULES.values())
    else:
        targets = ["-r", str(req)]

    base = [sys.executable, "-m", "pip", "install",
            "--disable-pip-version-check"] + targets

    # Make sure pip itself exists
    if subprocess.call([sys.executable, "-m", "pip", "--version"],
                       stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL, creationflags=_NO_WINDOW) != 0:
        log("pip not found — bootstrapping it with ensurepip...")
        _run_streaming([sys.executable, "-m", "ensurepip", "--upgrade"], log)

    rc = _run_streaming(base, log)
    if rc != 0:
        log("Normal install failed — retrying with --user ...")
        rc = _run_streaming(base + ["--user"], log)
    if rc != 0:
        raise RuntimeError(
            "Could not install the required Python packages.\n"
            "Check your internet connection and try again."
        )
    _add_user_site_to_path()


# ── Orchestration ─────────────────────────────────────────────────────────────

def _setup_steps(log: LogFn, status: LogFn) -> None:
    """Do whatever is needed. Runs in a worker thread when a GUI is shown."""
    missing = missing_python_packages()
    if missing:
        status("Installing required Python packages...")
        log("Missing packages: " + ", ".join(missing))
        install_python_packages(log)
        still = missing_python_packages()
        if still:
            raise RuntimeError("Packages still missing after install: "
                               + ", ".join(still))
        log("Python packages installed.")


def _file_logger() -> LogFn:
    path = log_file()

    def _log(msg: str) -> None:
        try:
            with open(path, "a", encoding="utf-8") as fh:
                fh.write(msg + "\n")
        except Exception:
            pass
    return _log


def _needs_setup() -> bool:
    return bool(missing_python_packages())


def ensure_ready(show_gui: bool = True) -> None:
    """
    Make sure every dependency is present, installing what's missing.
    Exits the process with a message if setup cannot be completed.
    """
    flog = _file_logger()

    if not _needs_setup():
        return

    if show_gui:
        ok, err = _run_with_splash(flog)
    else:
        ok, err = True, ""
        try:
            _setup_steps(lambda m: (print(m), flog(m)),
                         lambda m: print(f"== {m}"))
        except Exception as exc:                       # noqa: BLE001
            flog(traceback.format_exc())
            ok, err = False, str(exc)

    if not ok:
        _fatal(f"{err}\n\nSetup log: {log_file()}")


def _fatal(message: str) -> None:
    """Show an error (GUI if possible, else console) and exit."""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(f"{APP_TITLE} — Setup failed", message)
        root.destroy()
    except Exception:
        print("SETUP FAILED:\n" + message)
    sys.exit(1)


# ── Splash / progress window (plain tkinter) ──────────────────────────────────

def _run_with_splash(flog: LogFn):
    """Show a progress window while setup runs. Returns (ok, error_text)."""
    try:
        import tkinter as tk
        from tkinter import ttk
    except Exception:
        # No tkinter -> fall back to console mode
        try:
            _setup_steps(lambda m: (print(m), flog(m)), lambda m: print(f"== {m}"))
            return True, ""
        except Exception as exc:                        # noqa: BLE001
            flog(traceback.format_exc())
            return False, str(exc)

    q: "queue.Queue[tuple]" = queue.Queue()
    result = {"ok": True, "err": ""}

    def log(msg: str) -> None:
        flog(msg)
        q.put(("log", msg))

    def status(msg: str) -> None:
        flog("== " + msg)
        q.put(("status", msg))

    def worker() -> None:
        try:
            _setup_steps(log, status)
        except Exception as exc:                        # noqa: BLE001
            flog(traceback.format_exc())
            result["ok"], result["err"] = False, str(exc)
        finally:
            q.put(("done", None))

    root = tk.Tk()
    root.title(f"{APP_TITLE} — First-time setup")
    scale = max(1.0, root.winfo_fpixels("1i") / 96.0)    # 1.25 on a 125% display
    root.geometry(f"{int(560 * scale)}x{int(300 * scale)}")
    root.resizable(False, False)

    tk.Label(root, text="Preparing the application…",
             font=("Segoe UI", 12, "bold")).pack(pady=(14, 2))
    lbl_status = tk.Label(root, text="Checking required components…",
                          font=("Segoe UI", 10))
    lbl_status.pack()
    bar = ttk.Progressbar(root, mode="indeterminate", length=int(500 * scale))
    bar.pack(pady=10)
    bar.start(12)
    tk.Label(root, text="This only happens once and needs an internet connection.",
             font=("Segoe UI", 9), fg="#666").pack()

    txt = tk.Text(root, height=8, width=68, font=("Consolas", 8),
                  state="disabled", bg="#f4f4f4")
    txt.pack(padx=12, pady=8)

    def pump() -> None:
        try:
            while True:
                kind, val = q.get_nowait()
                if kind == "log":
                    txt.configure(state="normal")
                    txt.insert("end", val + "\n")
                    txt.see("end")
                    txt.configure(state="disabled")
                elif kind == "status":
                    lbl_status.configure(text=val)
                elif kind == "done":
                    root.destroy()
                    return
        except queue.Empty:
            pass
        root.after(100, pump)

    root.protocol("WM_DELETE_WINDOW", lambda: None)      # can't close mid-install
    threading.Thread(target=worker, daemon=True).start()
    root.after(100, pump)
    root.mainloop()
    return result["ok"], result["err"]
