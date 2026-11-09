"""
ATL Extractor — Entry Point
===========================
Run this file:  python main.py     (or just double-click run.bat)

On start-up the app checks for everything it needs and installs whatever
is missing by itself (pdfplumber, openpyxl, customtkinter, pillow).
Nothing needs to be installed manually except Python.

Options:
    python main.py --setup-only            install/verify dependencies, then exit
    python main.py --cli "<pdf folder>"    extract without opening the window
"""

import sys


def _enable_hidpi():
    """
    Tell Windows this app draws at the screen's real resolution. Without it
    Windows stretches the window on 125% / 150% displays and all text looks
    blurry. Must run before the first window (incl. the setup splash) opens.
    """
    if sys.platform != "win32":
        return
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)      # per-monitor
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def _run_cli(argv):
    """Extract a folder and write the Excel file — no window, no prompts."""
    import pathlib
    import queue
    import threading

    from constants import DONE_RESULTS, default_pdf_dir, default_output_dir
    from pdf_reader import build_records
    from pdf_engine import ExtractWorker
    from report_writer import default_output_name

    folder = argv[0] if argv else str(default_pdf_dir())
    pdf_dir = pathlib.Path(folder).resolve()
    out_dir = pathlib.Path(argv[1]).resolve() if len(argv) > 1 \
        else default_output_dir(str(pdf_dir))

    records = build_records(str(pdf_dir))
    if not records:
        print(f"No PDF files found in {pdf_dir}")
        return 1

    q: "queue.Queue[tuple]" = queue.Queue()
    worker = ExtractWorker(
        records=records, pdf_dir=pdf_dir, output_dir=out_dir, out_q=q,
        stop_event=threading.Event(), pause_event=threading.Event(),
        output_name=default_output_name(),
    )
    worker.run()

    while True:                                   # drain the queue to stdout
        try:
            msg = q.get_nowait()
        except queue.Empty:
            break
        if msg[0] == "log":
            print(msg[2])

    done = sum(1 for r in records if r.result in DONE_RESULTS)
    print(f"\nRead {done} / {len(records)}")
    print(f"Excel file: {worker.report_path}")
    return 0 if done else 1


def main():
    setup_only = "--setup-only" in sys.argv
    cli = "--cli" in sys.argv

    if not cli:
        _enable_hidpi()

    # bootstrap uses only the standard library, so it can run before any
    # third-party package has been installed.
    from bootstrap import ensure_ready
    ensure_ready(show_gui=not (setup_only or cli))

    if setup_only:
        print("Setup complete. All dependencies are installed.")
        return

    if cli:
        rest = sys.argv[sys.argv.index("--cli") + 1:]
        sys.exit(_run_cli([a for a in rest if not a.startswith("--")]))

    from gui import ATLExtractorApp
    app = ATLExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
