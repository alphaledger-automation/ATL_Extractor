"""
ATL Extractor — GUI
===================
Dark navy / cyan interface built with customtkinter. Two pages, switched
from the left sidebar:

  Dashboard       — pick the PDF folder / output folder / file name,
                    Start / Pause / Cancel, summary tiles, activity log.
  Current Status  — Done / Total counter and a live table of every PDF
                    with the fields read out of it.
"""

import datetime
import json
import os
import pathlib
import platform
import queue
import subprocess
import threading
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional

import customtkinter as ctk

from constants import (
    APP_NAME, APP_VERSION, APP_SUBTITLE, C, HDR, FONT_UI, FONT_MONO,
    DONE_RESULTS, RES_PENDING, RES_PROCESSING, RES_EXTRACTED,
    CAT_ACTIVE, CAT_LATE_FILER, CAT_INACTIVE, CAT_NO_RECORD, RN_RENAMED,
    default_pdf_dir, default_output_dir,
)
from pdf_reader import Record, build_records, list_pdfs, rename_target
from pdf_engine import ExtractWorker, RenameWorker
from report_writer import default_output_name


# ── Appearance ────────────────────────────────────────────────────────────────
ctk.set_default_color_theme("blue")

MODES = ("Light", "Dark")
_SETTINGS_FILE = pathlib.Path(
    os.environ.get("LOCALAPPDATA", str(pathlib.Path.home()))) / "ATL_Extractor" / "settings.json"


def load_mode() -> str:
    """Last-used appearance mode (Dark by default)."""
    try:
        mode = json.loads(_SETTINGS_FILE.read_text(encoding="utf-8")).get("mode")
        return mode if mode in MODES else "Dark"
    except Exception:
        return "Dark"


def save_mode(mode: str):
    try:
        _SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
        _SETTINGS_FILE.write_text(json.dumps({"mode": mode}), encoding="utf-8")
    except Exception:
        pass


def pick(key: str) -> str:
    """Palette colour for the current mode — for plain tk / ttk widgets,
    which (unlike customtkinter ones) do not switch by themselves."""
    return C[key][1 if ctk.get_appearance_mode() == "Dark" else 0]


ctk.set_appearance_mode(load_mode())

PAGE_DASHBOARD = "Dashboard"
PAGE_STATUS    = "Current Status"

JOB_EXTRACT = "extract"
JOB_RENAME  = "rename"

RENAME_PATTERNS = {
    "Registration No            (0123456.pdf)": "{reg}",
    "Registration No + date     (0123456_31-Dec-2023.pdf)": "{reg}_{date}",
    "Registration No + name     (0123456_ALPHA TRADERS.pdf)": "{reg}_{name}",
}


def _now() -> str:
    return datetime.datetime.now().strftime("%H:%M:%S")


def _open_path(path: str):
    if platform.system() == "Windows":
        os.startfile(path)
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def _shorten(text: str, limit: int = 62) -> str:
    return text if len(text) <= limit else "..." + text[-(limit - 3):]


# ── Small widget factories (keep the palette in one place) ───────────────────

def _font(size: int = 13, bold: bool = False, mono: bool = False) -> ctk.CTkFont:
    return ctk.CTkFont(family=FONT_MONO if mono else FONT_UI, size=size,
                       weight="bold" if bold else "normal")


def _card(parent, **kw) -> ctk.CTkFrame:
    return ctk.CTkFrame(parent, fg_color=C["bg700"], border_width=1,
                        border_color=C["div"], corner_radius=8, **kw)


def _label(parent, text="", size=13, bold=False, color=None, **kw) -> ctk.CTkLabel:
    return ctk.CTkLabel(parent, text=text, font=_font(size, bold),
                        text_color=color or C["white"], **kw)


def _btn_primary(parent, text, command, **kw) -> ctk.CTkButton:
    return ctk.CTkButton(
        parent, text=text, command=command, font=_font(14, True),
        fg_color=C["a300"], hover_color=C["a200"], text_color=C["on_acc"],
        text_color_disabled=C["sl400"], corner_radius=7, **kw)


def _btn_outline(parent, text, command, color, hover, **kw) -> ctk.CTkButton:
    return ctk.CTkButton(
        parent, text=text, command=command, font=_font(13, True),
        fg_color="transparent", hover_color=hover, text_color=color,
        text_color_disabled=C["sl400"], border_width=1, border_color=color,
        corner_radius=7, **kw)


class ATLExtractorApp(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__(fg_color=C["bg800"])

        self.title(f"{APP_NAME} v{APP_VERSION}")
        self._fit_to_screen(1180, 720)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        try:
            icon = pathlib.Path(__file__).resolve().parent / "Assets" / "Icon.ico"
            if icon.exists():
                self.iconbitmap(str(icon))
        except Exception:
            pass

        # ── Internal state ────────────────────────────────────────────────
        self._pdf_dir: pathlib.Path = default_pdf_dir()
        self._records: List[Record] = []
        self._output_dir: pathlib.Path = default_output_dir(str(self._pdf_dir))
        self._output_dir_custom = False
        self._worker_thread: Optional[threading.Thread] = None
        self._stop_event  = threading.Event()
        self._pause_event = threading.Event()       # SET = paused
        self._q = queue.Queue()
        self._is_running = False
        self._job = ""
        self._report_path = ""
        self._pages = {}
        self._nav = {}
        self._scale = 1.0

        self._build_ui()
        self._apply_native_theme()
        self._show_page(PAGE_DASHBOARD)
        self._load_folder(str(self._pdf_dir), auto=True)
        self._poll_queue()

    # ==================================================================
    #  UI BUILD
    # ==================================================================

    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # ── Header bar ───────────────────────────────────────────────────
        header = ctk.CTkFrame(self, fg_color=HDR["bg"], corner_radius=0, height=62)
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(1, weight=1)

        _label(header, APP_NAME, size=19, bold=True, color=HDR["title"]).grid(
            row=0, column=0, sticky="w", padx=(22, 12), pady=(16, 16))
        _label(header, APP_SUBTITLE, size=13, color=HDR["sub"]).grid(
            row=0, column=1, sticky="w", pady=(19, 16))
        self._lbl_header_count = _label(header, "", size=14, bold=True,
                                        color=HDR["count"])
        self._lbl_header_count.grid(row=0, column=2, sticky="e", padx=22)

        ctk.CTkFrame(self, fg_color=C["a500"], corner_radius=0, height=2).grid(
            row=1, column=0, columnspan=2, sticky="ew")

        # ── Sidebar ──────────────────────────────────────────────────────
        side = ctk.CTkFrame(self, fg_color=C["side"], corner_radius=0, width=200)
        side.grid(row=2, column=0, sticky="nsw")
        side.grid_propagate(False)
        side.grid_columnconfigure(1, weight=1)
        side.grid_rowconfigure(10, weight=1)

        _label(side, "NAVIGATION", size=11, bold=True, color=C["a300"]).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=18, pady=(18, 8))

        for i, name in enumerate((PAGE_DASHBOARD, PAGE_STATUS), start=1):
            bar = ctk.CTkFrame(side, fg_color="transparent", corner_radius=0,
                               width=3, height=40)
            bar.grid(row=i, column=0, sticky="ns")
            btn = ctk.CTkButton(
                side, text="   " + name, anchor="w", height=40, corner_radius=0,
                font=_font(13), fg_color="transparent",
                hover_color=C["sl700"], text_color=C["sl200"],
                command=lambda n=name: self._show_page(n))
            btn.grid(row=i, column=1, sticky="ew")
            self._nav[name] = (bar, btn)

        _label(side, "APPEARANCE", size=11, bold=True, color=C["a300"]).grid(
            row=11, column=0, columnspan=2, sticky="w", padx=18, pady=(0, 6))
        self._mode_switch = ctk.CTkSegmentedButton(
            side, values=list(MODES), font=_font(12, True), height=30,
            fg_color=C["bg900"], unselected_color=C["bg900"],
            unselected_hover_color=C["sl700"], selected_color=C["seg_sel"],
            selected_hover_color=C["seg_hover"], text_color=C["white"],
            command=self._set_mode)
        self._mode_switch.set(ctk.get_appearance_mode())
        self._mode_switch.grid(row=12, column=0, columnspan=2, sticky="ew",
                               padx=16, pady=(0, 14))

        _label(side, f"Version {APP_VERSION}", size=12, color=C["sl300"]).grid(
            row=13, column=0, columnspan=2, sticky="w", padx=18, pady=(0, 14))

        # ── Pages ────────────────────────────────────────────────────────
        body = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        body.grid(row=2, column=1, sticky="nsew", padx=16, pady=14)
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        for name, builder in ((PAGE_DASHBOARD, self._build_dashboard),
                              (PAGE_STATUS, self._build_status)):
            page = ctk.CTkFrame(body, fg_color="transparent", corner_radius=0)
            page.grid(row=0, column=0, sticky="nsew")
            builder(page)
            self._pages[name] = page

    def _fit_to_screen(self, width: int, height: int):
        """Open centred, and never larger than the screen (small laptops)."""
        try:
            scale = float(self._get_window_scaling())
        except Exception:
            scale = 1.0
        screen_w, screen_h = self.winfo_screenwidth(), self.winfo_screenheight()
        width  = int(min(width,  screen_w / scale - 40))
        height = int(min(height, screen_h / scale - 100))    # taskbar + title bar
        x = max(0, int((screen_w - width * scale) / 2))
        y = max(0, int((screen_h - height * scale) / 2 - 30 * scale))
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.minsize(min(1000, width), min(600, height))

    def _show_page(self, name: str):
        for page_name, page in self._pages.items():
            bar, btn = self._nav[page_name]
            if page_name == name:
                page.tkraise()
                bar.configure(fg_color=C["acc"])
                btn.configure(fg_color=C["sl700"], text_color=C["a200"],
                              font=_font(13, True))
            else:
                bar.configure(fg_color="transparent")
                btn.configure(fg_color="transparent", text_color=C["sl200"],
                              font=_font(13))

    # ── Dashboard page ───────────────────────────────────────────────────

    def _build_dashboard(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(5, weight=1)          # log area stretches

        # 1. Input: PDF folder, output folder, file name
        inp = _card(tab)
        inp.grid(row=0, column=0, sticky="ew")
        inp.grid_columnconfigure(1, weight=1)

        _label(inp, "PDF Folder", bold=True).grid(
            row=0, column=0, sticky="w", padx=16, pady=(12, 4))
        folder_row = ctk.CTkFrame(inp, fg_color="transparent")
        folder_row.grid(row=0, column=1, sticky="ew", padx=8, pady=(12, 4))
        self._folder_label = _label(folder_row, "No folder selected",
                                    color=C["sl300"], anchor="w")
        self._folder_label.pack(side="left")
        self._folder_status = _label(folder_row, "", color=C["sl200"])
        self._folder_status.pack(side="left", padx=12)
        self._btn_browse = _btn_outline(
            inp, "Browse ...", self._browse_folder, C["a300"], C["sl700"],
            width=110)
        self._btn_browse.grid(row=0, column=2, padx=16, pady=(12, 4))

        _label(inp, "Save Excel To", bold=True).grid(
            row=1, column=0, sticky="w", padx=16, pady=4)
        self._out_label = _label(inp, str(self._output_dir), color=C["sl200"],
                                 anchor="w")
        self._out_label.grid(row=1, column=1, sticky="ew", padx=8, pady=4)
        self._btn_out = _btn_outline(
            inp, "Change ...", self._browse_output, C["a300"], C["sl700"],
            width=110)
        self._btn_out.grid(row=1, column=2, padx=16, pady=4)

        _label(inp, "Excel File Name", bold=True).grid(
            row=2, column=0, sticky="w", padx=16, pady=(4, 12))
        name_row = ctk.CTkFrame(inp, fg_color="transparent")
        name_row.grid(row=2, column=1, columnspan=2, sticky="ew", padx=8,
                      pady=(4, 12))
        self._name_entry = ctk.CTkEntry(
            name_row, width=300, font=_font(13), fg_color=C["bg900"],
            text_color=C["white"], border_color=C["sl600"], border_width=1)
        self._name_entry.insert(0, default_output_name())
        self._name_entry.pack(side="left")
        _label(name_row, "written and re-written while the run is in progress",
               color=C["sl200"]).pack(side="left", padx=12)

        # 2. Action buttons + options
        btn_frame = ctk.CTkFrame(tab, fg_color="transparent")
        btn_frame.grid(row=1, column=0, sticky="ew", pady=(12, 6))

        self._btn_start = _btn_primary(btn_frame, "Start", self._start,
                                       width=140, height=38)
        self._btn_start.pack(side="left", padx=(0, 8))
        self._btn_pause = _btn_outline(
            btn_frame, "Pause", self._toggle_pause, C["warn"], C["warn_hover"],
            width=110, height=38, state="disabled")
        self._btn_pause.pack(side="left", padx=(0, 8))
        self._btn_cancel = _btn_outline(
            btn_frame, "Cancel", self._cancel, C["err"], C["err_hover"],
            width=110, height=38, state="disabled")
        self._btn_cancel.pack(side="left")

        chk = dict(font=_font(13), text_color=C["sl100"], fg_color=C["a400"],
                   hover_color=C["a500"], border_color=C["sl400"],
                   checkmark_color=C["on_acc"], border_width=2,
                   checkbox_width=20, checkbox_height=20)
        self._chk_open = ctk.CTkCheckBox(
            btn_frame, text="Open Excel when finished", **chk)
        self._chk_open.pack(side="right", padx=(16, 0))
        self._chk_skip = ctk.CTkCheckBox(
            btn_frame, text="Skip files already extracted", **chk)
        self._chk_skip.pack(side="right")

        # 3. Summary tiles
        tiles = ctk.CTkFrame(tab, fg_color="transparent")
        tiles.grid(row=2, column=0, sticky="ew", pady=6)
        self._tiles = {}
        for i, (key, title, colour) in enumerate([
            ("total",     "PDF FILES", C["white"]),
            ("done",      "READ",      C["a300"]),
            ("pending",   "PENDING",   C["warn"]),
            ("failed",    "FAILED",    C["err"]),
            ("active",    "ACTIVE",    C["ok"]),
            ("late",      "LATE FILER", C["warn"]),
            ("inactive",  "IN-ACTIVE", C["err"]),
        ]):
            tiles.grid_columnconfigure(i, weight=1, uniform="tile")
            box = _card(tiles)
            box.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 8, 0))
            value = _label(box, "0", size=26, bold=True, color=colour)
            value.pack(pady=(6, 0))
            _label(box, title, size=11, bold=True, color=C["sl200"]).pack(pady=(0, 7))
            self._tiles[key] = value

        # 4. Progress bar
        prog = _card(tab)
        prog.grid(row=3, column=0, sticky="ew", pady=6)
        prog.grid_columnconfigure(0, weight=1)
        top = ctk.CTkFrame(prog, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=16, pady=(10, 4))
        top.grid_columnconfigure(0, weight=1)
        self._lbl_progress = _label(top, "Read 0 / 0", bold=True)
        self._lbl_progress.grid(row=0, column=0, sticky="w")
        self._lbl_current = _label(top, "", color=C["sl200"])
        self._lbl_current.grid(row=0, column=1, sticky="e")
        self._progress_bar = ctk.CTkProgressBar(
            prog, height=12, fg_color=C["bg900"], progress_color=C["a400"],
            border_width=1, border_color=C["sl700"])
        self._progress_bar.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 14))
        self._progress_bar.set(0)

        # 5. Activity log
        log_hdr = ctk.CTkFrame(tab, fg_color="transparent")
        log_hdr.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        _label(log_hdr, "ACTIVITY LOG", size=11, bold=True,
               color=C["a300"]).pack(side="left", padx=2)
        _btn_outline(log_hdr, "Clear", self._clear_log, C["sl200"], C["sl700"],
                     width=70, height=26).pack(side="right")

        self._log_box = ctk.CTkTextbox(
            tab, font=_font(12, mono=True), state="disabled", wrap="word",
            height=70, fg_color=C["bg900"], text_color=C["log_info"],
            border_width=1, border_color=C["div"], corner_radius=7,
            scrollbar_button_color=C["sl600"],
            scrollbar_button_hover_color=C["a500"])
        self._log_box.grid(row=5, column=0, sticky="nsew", pady=(6, 8))

        # 6. Footer: output + the rename job
        foot = ctk.CTkFrame(tab, fg_color="transparent")
        foot.grid(row=6, column=0, sticky="ew")
        _btn_outline(foot, "Open PDF Folder", self._open_pdf_folder,
                     C["a300"], C["sl700"], width=150, height=32).pack(
                         side="left", padx=(0, 8))
        self._btn_report = _btn_outline(
            foot, "Open Excel File", self._open_report, C["a300"], C["sl700"],
            width=150, height=32, state="disabled")
        self._btn_report.pack(side="left")

        self._btn_rename = _btn_outline(
            foot, "Rename PDFs ...", self._rename, C["warn"], C["warn_hover"],
            width=160, height=32, state="disabled")
        self._btn_rename.pack(side="right")
        self._rename_pattern = ctk.CTkComboBox(
            foot, values=list(RENAME_PATTERNS), width=330, state="readonly",
            font=_font(12), dropdown_font=_font(12), fg_color=C["bg900"],
            text_color=C["white"], border_color=C["sl600"],
            button_color=C["sl600"], button_hover_color=C["a500"],
            dropdown_fg_color=C["bg800"], dropdown_text_color=C["white"],
            dropdown_hover_color=C["a600"], height=32)
        self._rename_pattern.set(list(RENAME_PATTERNS)[0])
        self._rename_pattern.pack(side="right", padx=(0, 8))

    # ── Current Status page ──────────────────────────────────────────────

    def _build_status(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        top = _card(tab)
        top.grid(row=0, column=0, sticky="ew")
        top.grid_columnconfigure(1, weight=1)

        self._lbl_big = _label(top, "0 / 0", size=36, bold=True, color=C["a300"])
        self._lbl_big.grid(row=0, column=0, rowspan=2, padx=(20, 10), pady=12)
        _label(top, "READ / TOTAL", size=11, bold=True, color=C["sl200"]).grid(
            row=0, column=1, sticky="sw", padx=8, pady=(16, 0))
        self._lbl_status_detail = _label(top, "", color=C["sl100"], anchor="w")
        self._lbl_status_detail.grid(row=1, column=1, sticky="nw", padx=8,
                                     pady=(0, 12))
        self._status_bar = ctk.CTkProgressBar(
            top, height=12, fg_color=C["bg900"], progress_color=C["a400"],
            border_width=1, border_color=C["sl700"])
        self._status_bar.grid(row=2, column=0, columnspan=2, sticky="ew",
                              padx=20, pady=(0, 16))
        self._status_bar.set(0)

        table = _card(tab)
        table.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        table.grid_columnconfigure(0, weight=1)
        table.grid_rowconfigure(0, weight=1)

        # ttk widgets are not scaled by customtkinter, so do it here — this
        # keeps the table text the same crisp size as the rest of the UI.
        try:
            scale = float(self._get_window_scaling())
        except Exception:
            scale = 1.0
        self._scale = scale

        ttk.Style(self).theme_use("clam")           # the only theme that takes colours

        cols = [("no", "#", 44, "center"), ("reg", "Registration No", 130, "w"),
                ("type", "Type", 56, "center"), ("name", "Name", 180, "w"),
                ("business", "Business Name", 160, "w"),
                ("filing", "Filing Status", 130, "w"),
                ("date", "Checking Date", 112, "center"),
                ("file", "PDF File", 170, "w"),
                ("result", "Result", 150, "w")]
        self._tree = ttk.Treeview(table, columns=[c[0] for c in cols],
                                  show="headings", style="ATL.Treeview")
        for key, title, width, anchor in cols:
            self._tree.heading(key, text=title, anchor=anchor)
            self._tree.column(key, width=int(width * scale), anchor=anchor,
                              stretch=(key == "name"))

        scroll = ctk.CTkScrollbar(
            table, command=self._tree.yview, fg_color="transparent",
            button_color=C["sl600"], button_hover_color=C["a500"])
        self._tree.configure(yscrollcommand=scroll.set)
        self._tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)
        scroll.grid(row=0, column=1, sticky="ns", padx=(2, 6), pady=10)

    # ── Light / Dark ─────────────────────────────────────────────────────

    def _set_mode(self, mode: str):
        ctk.set_appearance_mode(mode)       # recolours every customtkinter widget
        self._apply_native_theme()
        save_mode(mode)

    def _apply_native_theme(self):
        """Colour the plain tk / ttk parts (table, log tags) for the current mode."""
        scale = self._scale
        px = lambda n: -max(1, int(round(n * scale)))   # negative = pixels

        style = ttk.Style(self)
        style.configure(
            "ATL.Treeview", rowheight=int(round(30 * scale)),
            font=(FONT_UI, px(13)), background=pick("bg900"),
            fieldbackground=pick("bg900"), foreground=pick("white"),
            borderwidth=0, relief="flat")
        style.configure(
            "ATL.Treeview.Heading", font=(FONT_UI, px(13), "bold"),
            background=pick("bg600"), foreground=pick("a200"), relief="flat",
            borderwidth=0, padding=(6, int(round(7 * scale))))
        style.map("ATL.Treeview",
                  background=[("selected", pick("a600"))],
                  foreground=[("selected", pick("sel_fg"))])
        style.map("ATL.Treeview.Heading",
                  background=[("active", pick("bg500"))])
        style.layout("ATL.Treeview",
                     [("Treeview.treearea", {"sticky": "nswe"})])   # no border

        self._tree.tag_configure("done",    foreground=pick("ok"))
        self._tree.tag_configure("partial", foreground=pick("warn"))
        self._tree.tag_configure("failed",  foreground=pick("err"))
        self._tree.tag_configure("working", foreground=pick("warn"))
        self._tree.tag_configure("pending", foreground=pick("sl200"))

        text = self._log_box._textbox
        text.tag_configure("ok",   foreground=pick("log_ok"))
        text.tag_configure("err",  foreground=pick("log_err"))
        text.tag_configure("warn", foreground=pick("log_warn"))
        text.tag_configure("info", foreground=pick("log_info"))

    # ==================================================================
    #  TABLE + COUNTERS
    # ==================================================================

    @staticmethod
    def _row_tag(rec: Record) -> str:
        if rec.result == RES_EXTRACTED:
            return "done"
        if rec.result in DONE_RESULTS:
            return "partial"
        if rec.result == RES_PROCESSING:
            return "working"
        if rec.result == RES_PENDING:
            return "pending"
        return "failed"

    @staticmethod
    def _row_values(rec: Record):
        return (rec.idx, rec.reg_no, rec.param_type, rec.name,
                rec.business_name, rec.filing_status, rec.checking_date,
                rec.source_file, rec.rename_status or rec.result)

    def _fill_table(self):
        self._tree.delete(*self._tree.get_children())
        for rec in self._records:
            self._tree.insert("", "end", iid=str(rec.idx),
                              values=self._row_values(rec),
                              tags=(self._row_tag(rec),))

    def _update_row(self, rec: Record):
        iid = str(rec.idx)
        if self._tree.exists(iid):
            self._tree.item(iid, values=self._row_values(rec),
                            tags=(self._row_tag(rec),))
            if rec.result == RES_PROCESSING:
                self._tree.see(iid)

    def _refresh_counts(self):
        total = len(self._records)
        done = failed = active = late = inactive = no_record = 0
        for rec in self._records:
            if rec.result in DONE_RESULTS:
                done += 1
            elif rec.result not in (RES_PENDING, RES_PROCESSING):
                failed += 1
            if rec.category == CAT_LATE_FILER:
                late += 1
            elif rec.category == CAT_INACTIVE:
                inactive += 1
            elif rec.category == CAT_NO_RECORD:
                no_record += 1
            elif rec.category == CAT_ACTIVE:
                active += 1
        pending = total - done - failed

        for key, value in (("total", total), ("done", done), ("pending", pending),
                           ("failed", failed), ("active", active),
                           ("late", late), ("inactive", inactive)):
            self._tiles[key].configure(text=f"{value:,}")

        fraction = (done + failed) / total if total else 0
        self._progress_bar.set(fraction)
        self._status_bar.set(fraction)
        self._lbl_progress.configure(text=f"Read {done:,} / {total:,}")
        self._lbl_big.configure(text=f"{done:,} / {total:,}")
        self._lbl_header_count.configure(
            text=f"Read {done:,} / {total:,}" if total else "")
        self._lbl_status_detail.configure(
            text=f"Pending {pending:,}   •   Failed {failed:,}   •   "
                 f"Active {active:,}   •   Late filer {late:,}   •   "
                 f"In-Active {inactive:,}   •   No record {no_record:,}")

        self._btn_rename.configure(
            state="normal" if (done and not self._is_running) else "disabled")

    # ==================================================================
    #  EVENT HANDLERS
    # ==================================================================

    def _browse_folder(self):
        folder = filedialog.askdirectory(
            title="Select the folder with ATL Status PDFs",
            initialdir=str(self._pdf_dir))
        if folder:
            self._load_folder(folder, auto=False)

    def _load_folder(self, folder: str, auto: bool):
        try:
            names = list_pdfs(folder)
        except Exception as exc:                            # noqa: BLE001
            self._folder_status.configure(text=str(exc).splitlines()[0],
                                          text_color=C["err"])
            return

        self._pdf_dir = pathlib.Path(folder)
        self._records = build_records(folder)
        self._folder_label.configure(text=_shorten(str(self._pdf_dir)),
                                     text_color=C["white"])
        text = f"{len(names):,} PDF file(s)"
        if auto:
            text = "Auto-detected — " + text
        self._folder_status.configure(
            text=text, text_color=C["ok"] if names else C["err"])

        if not self._output_dir_custom:
            self._output_dir = default_output_dir(folder)
            self._out_label.configure(text=_shorten(str(self._output_dir)))

        self._report_path = ""
        self._btn_report.configure(state="disabled")
        self._fill_table()
        self._refresh_counts()

    def _browse_output(self):
        folder = filedialog.askdirectory(
            title="Select the folder for the Excel file",
            initialdir=str(self._output_dir))
        if folder:
            self._output_dir = pathlib.Path(folder)
            self._output_dir_custom = True
            self._out_label.configure(text=_shorten(str(self._output_dir)))

    # ── Start / Pause / Cancel ────────────────────────────────────────────

    def _set_inputs_state(self, state: str):
        for widget in (self._btn_browse, self._btn_out, self._name_entry,
                       self._chk_skip, self._chk_open, self._btn_rename):
            widget.configure(state=state)
        self._rename_pattern.configure(
            state="readonly" if state == "normal" else "disabled")

    def _begin_job(self, job: str, worker) -> None:
        """Shared Start plumbing for both the extract and rename jobs."""
        self._stop_event.clear()
        self._pause_event.clear()
        self._is_running = True
        self._job = job

        self._btn_start.configure(state="disabled")
        self._btn_pause.configure(state="normal", text="Pause")
        self._btn_cancel.configure(state="normal")
        self._set_inputs_state("disabled")

        self._worker_thread = threading.Thread(target=worker.run, daemon=True)
        self._worker_thread.start()

    def _start(self):
        if self._is_running:
            return
        if not self._records:
            messagebox.showwarning(
                "No PDFs",
                "No PDF files found. Pick the folder that holds the ATL "
                "Status PDFs.")
            return

        name = self._name_entry.get().strip() or default_output_name()
        self._report_path = ""
        self._btn_report.configure(state="disabled")

        worker = ExtractWorker(
            records=self._records,
            pdf_dir=self._pdf_dir,
            output_dir=self._output_dir,
            out_q=self._q,
            stop_event=self._stop_event,
            pause_event=self._pause_event,
            output_name=name,
            skip_done=bool(self._chk_skip.get()),
        )
        self._begin_job(JOB_EXTRACT, worker)
        self._append_log(
            f"[{_now()}] Started — {len(self._records)} PDF file(s) from "
            f"{self._pdf_dir}", "info")

    def _rename(self):
        if self._is_running:
            return
        ready = [r for r in self._records if r.result in DONE_RESULTS]
        if not ready:
            messagebox.showwarning(
                "Nothing to rename",
                "Run the extraction first — the registration numbers come "
                "from inside the PDFs.")
            return

        pattern = RENAME_PATTERNS.get(self._rename_pattern.get(), "{reg}")
        will_change = [r for r in ready
                       if rename_target(r, pattern).lower() != r.source_file.lower()]
        no_reg = [r for r in ready if not r.reg_no]
        sample = ", ".join(rename_target(r, pattern)
                           for r in will_change[:3]) or "(none)"

        if not messagebox.askyesno(
                "Rename PDF files",
                f"{len(will_change)} of {len(ready)} file(s) will be renamed "
                f"inside:\n{self._pdf_dir}\n\n"
                f"For example:\n{sample}\n\n"
                + (f"{len(no_reg)} file(s) have no registration number and "
                   "will be left alone.\n\n" if no_reg else "")
                + "This changes the files on disk and cannot be undone "
                  "automatically. Continue?"):
            return

        worker = RenameWorker(
            records=self._records,
            pdf_dir=self._pdf_dir,
            out_q=self._q,
            stop_event=self._stop_event,
            pause_event=self._pause_event,
            pattern=pattern,
            output_dir=self._output_dir,
            output_name=self._name_entry.get().strip(),
        )
        self._begin_job(JOB_RENAME, worker)
        self._append_log(f"[{_now()}] Renaming started — pattern {pattern}",
                         "info")

    def _toggle_pause(self):
        if self._pause_event.is_set():
            self._pause_event.clear()
            self._btn_pause.configure(text="Pause")
            self._append_log(f"[{_now()}] Resumed.", "info")
        else:
            self._pause_event.set()
            self._btn_pause.configure(text="Resume")
            self._append_log(
                f"[{_now()}] Paused after the current file — click Resume to "
                "continue.", "warn")

    def _cancel(self):
        self._stop_event.set()
        self._pause_event.clear()
        self._btn_cancel.configure(state="disabled")
        self._btn_pause.configure(state="disabled")
        self._append_log(
            f"[{_now()}] Cancel requested — finishing the current file...",
            "warn")

    # ── Footer actions ────────────────────────────────────────────────────

    def _open_pdf_folder(self):
        if self._pdf_dir.is_dir():
            _open_path(str(self._pdf_dir))

    def _open_report(self):
        if self._report_path and os.path.exists(self._report_path):
            _open_path(self._report_path)

    def _clear_log(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    def _on_close(self):
        if self._is_running:
            if not messagebox.askyesno("Quit", "A run is in progress. Stop and quit?"):
                return
            self._stop_event.set()
            self._pause_event.clear()
        self.destroy()

    # ==================================================================
    #  QUEUE POLLING  (runs on the GUI / main thread every 150 ms)
    # ==================================================================

    def _poll_queue(self):
        changed = False
        for _ in range(400):
            try:
                msg = self._q.get_nowait()
            except queue.Empty:
                break
            kind = msg[0]
            if kind == "log":
                self._append_log(msg[2], msg[1])
            elif kind == "record":
                rec = msg[1]
                self._update_row(rec)
                if rec.result == RES_PROCESSING:
                    self._lbl_current.configure(
                        text=f"Current: {_shorten(rec.source_file, 44)}")
                changed = True
            elif kind == "report":
                self._report_path = msg[1]
            elif kind == "done":
                self._on_done()
                changed = True
        if changed:
            self._refresh_counts()
        self.after(150, self._poll_queue)

    def _append_log(self, text: str, tag: str = "info"):
        self._log_box.configure(state="normal")
        self._log_box._textbox.insert("end", text + "\n", tag)
        self._log_box._textbox.see("end")
        self._log_box.configure(state="disabled")

    def _on_done(self):
        job = self._job
        self._is_running = False
        self._job = ""
        self._btn_start.configure(state="normal")
        self._btn_pause.configure(state="disabled", text="Pause")
        self._btn_cancel.configure(state="disabled")
        self._set_inputs_state("normal")
        self._lbl_current.configure(text="")
        if self._report_path:
            self._btn_report.configure(state="normal")
            self._append_log(f"[{_now()}] Excel file: {self._report_path}", "ok")

        total = len(self._records)
        done = sum(1 for r in self._records if r.result in DONE_RESULTS)
        pending = sum(1 for r in self._records if r.result == RES_PENDING)

        if job == JOB_RENAME:
            renamed = sum(1 for r in self._records if r.rename_status == RN_RENAMED)
            messagebox.showinfo(
                "Renaming complete",
                f"Renamed: {renamed} / {done} file(s)\n\nFolder:\n{self._pdf_dir}")
            return

        messagebox.showinfo(
            "Complete" if not pending else "Stopped",
            f"Read: {done} / {total}\n"
            f"Failed: {total - done - pending}\n"
            + (f"Not processed: {pending}\n" if pending else "")
            + (f"\nExcel file:\n{self._report_path}" if self._report_path else ""),
        )
        if self._report_path and self._chk_open.get():
            self._open_report()
