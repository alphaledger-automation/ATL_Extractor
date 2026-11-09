"""
ATL Extractor — Constants & Configuration
=========================================
All field labels, patterns, timing values and the colour palette live here
so the rest of the code stays clean.
"""

import pathlib
import re

# ── Fields read from an ATL verification PDF ─────────────────────────────────
# Longest labels first — the alternation is matched in this order, so
# "Business Name" and "Filing Status Checking Date" are never mistaken for
# the shorter "Name" / "Filing Status".
PDF_LABELS = [
    "Filing Status Checking Date",
    "Registration No",
    "Business Name",
    "Filing Status",
    "Name",
]

LABEL_RE = re.compile(
    r"(?:^|(?<=\s))(" + "|".join(re.escape(lbl) for lbl in PDF_LABELS) + r")\s*:",
    re.MULTILINE,
)

# The FBR boiler-plate that follows the last field — everything from here on
# is dropped.
FOOTER_MARKERS = (
    "You can also check your Active Taxpayer Status",
    "Registration Number (e.g. ATL",
)

# "In-Active" results carry a long surcharge note on the same line.
NOTE_MARKER = "Note:"

# Text that marks a PDF with no result in it
NO_RECORD_MARKER = "no record exists"

# A PDF must contain this label to count as an ATL verification slip
REQUIRED_MARKER = "Registration No"

# ── Filing-status categories ─────────────────────────────────────────────────
CAT_ACTIVE      = "Active"
CAT_LATE_FILER  = "Active (Late Filer)"
CAT_INACTIVE    = "In-Active"
CAT_NO_RECORD   = "No record"
CAT_UNKNOWN     = "Unknown"

CATEGORIES = (CAT_ACTIVE, CAT_LATE_FILER, CAT_INACTIVE, CAT_NO_RECORD, CAT_UNKNOWN)

# ── Timing (seconds) ────────────────────────────────────────────────────────
INTER_FILE_DELAY_SECS = 0        # extraction is local — no throttling needed
REPORT_SAVE_EVERY     = 25       # re-save the Excel file every N records

# ── Input / output ──────────────────────────────────────────────────────────
PDF_SUBFOLDER = "ATL Status"     # folder name looked for next to the app

MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
          "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}
MONTH_NAMES = ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


def default_pdf_dir() -> pathlib.Path:
    """The 'ATL Status' folder next to the app, if it exists."""
    here = pathlib.Path(__file__).resolve().parent
    for base in (here, here.parent):
        candidate = base / PDF_SUBFOLDER
        if candidate.is_dir():
            return candidate
    return here


def default_output_dir(pdf_dir: str = "") -> pathlib.Path:
    """The Excel file is written next to the PDF folder by default."""
    if pdf_dir:
        return pathlib.Path(pdf_dir).resolve().parent
    return pathlib.Path.home() / "Downloads"


# ── Record result values ────────────────────────────────────────────────────
RES_PENDING    = "Pending"
RES_PROCESSING = "Processing"
RES_EXTRACTED  = "Extracted"
RES_PARTIAL    = "Extracted (incomplete)"
RES_NO_RECORD  = "No record in PDF"
RES_NOT_ATL    = "Not an ATL slip"
RES_FAILED     = "Failed"            # always followed by " - <reason>"

DONE_RESULTS = (RES_EXTRACTED, RES_PARTIAL, RES_NO_RECORD)

# Rename outcomes
RN_PENDING  = ""
RN_RENAMED  = "Renamed"
RN_SAME     = "Already named"
RN_EXISTS   = "Skipped - name taken"
RN_NO_REG   = "Skipped - no Reg No"
RN_FAILED   = "Rename failed"

# ── GUI Appearance ──────────────────────────────────────────────────────────
APP_NAME    = "ATL Extractor"
APP_VERSION = "2.0"
APP_SUBTITLE = "ATL Status PDFs  →  Excel"

# Colour palette. Every entry is (light, dark).
C = {
    # surfaces
    "bg800": ("#EEF3F8", "#132238"),      # window
    "side":  ("#E1E9F2", "#0E1A2E"),      # sidebar
    "bg700": ("#FFFFFF", "#182B45"),      # cards
    "bg900": ("#F7FAFC", "#0E1A2E"),      # inputs, log, table
    "bg600": ("#DCE6F0", "#1E3452"),      # table heading
    "bg500": ("#C9D7E6", "#243D60"),
    "div":   ("#CBD5E1", "#243D60"),      # card borders
    # cyan accent
    "a600": ("#BDE8F3", "#006B84"),       # selection background
    "a500": ("#008EAD", "#008EAD"),
    "a400": ("#0096B8", "#00B4D8"),       # progress bar, checkbox
    "a300": ("#00758F", "#22D3EE"),       # accent text / primary button
    "a200": ("#005F75", "#67E8F9"),
    "acc":  ("#00758F", "#22D3EE"),
    "on_acc": ("#FFFFFF", "#080F1C"),     # text on an accent-filled button
    "sel_fg": ("#0F172A", "#FFFFFF"),     # text on a selected table row
    "seg_sel": ("#9ADBEA", "#008EAD"),    # selected Light/Dark switch segment
    "seg_hover": ("#7CCFE2", "#00A0C4"),
    # slate
    "sl700": ("#DCE6F0", "#1C2E44"),      # hover background
    "sl600": ("#B6C4D4", "#2A4060"),      # input borders, scrollbars
    "sl500": ("#94A3B8", "#3D5A7A"),
    "sl400": ("#8496AB", "#5A7A9A"),      # disabled text
    "sl300": ("#64748B", "#7AA0BC"),
    "sl200": ("#475569", "#A8C4D8"),      # secondary text
    "sl100": ("#334155", "#D8EAF4"),
    # text + status
    "white": ("#0F172A", "#FFFFFF"),      # primary text
    "ok":    ("#0B8A5F", "#2EE8A8"),
    "warn":  ("#B45309", "#FBBC41"),
    "err":   ("#D12E43", "#F95C6B"),
    "warn_hover": ("#FDF3DC", "#3A3320"),
    "err_hover":  ("#FDE7EA", "#3A2030"),
    "log_info": ("#1E4E6B", "#90D5F0"), "log_ok": ("#0B8A5F", "#2EE8A8"),
    "log_warn": ("#B45309", "#FBBC41"), "log_err": ("#D12E43", "#F95C6B"),
}

# The header bar keeps the navy in both modes.
HDR = {"bg": "#080F1C", "title": "#FFFFFF", "sub": "#A8C4D8", "count": "#22D3EE"}

# Segoe UI / Consolas are hinted for Windows ClearType, so they stay sharp
# at small sizes (customtkinter's default Roboto does not).
FONT_UI   = "Segoe UI"
FONT_MONO = "Consolas"
