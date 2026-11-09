"""
ATL Extractor — PDF Input Handler
=================================
Lists the ATL verification PDFs in a folder and pulls the five fields out
of each one:

    Registration No | Name | Business Name | Filing Status |
    Filing Status Checking Date

The PDFs are only ever read here. Renaming is a separate, explicit step
(see pdf_engine.rename_files).
"""

from __future__ import annotations

import pathlib
import re
from dataclasses import dataclass
from typing import Dict, List, Tuple

import pdfplumber

from constants import (
    LABEL_RE, FOOTER_MARKERS, NOTE_MARKER, NO_RECORD_MARKER, REQUIRED_MARKER,
    MONTHS, MONTH_NAMES,
    CAT_ACTIVE, CAT_LATE_FILER, CAT_INACTIVE, CAT_NO_RECORD, CAT_UNKNOWN,
    RES_PENDING, RN_PENDING,
)


@dataclass
class Record:
    """One ATL PDF and everything we learn about it."""
    idx: int                        # 1-based position in the run
    source_file: str                # file name as found on disk
    reg_no: str = ""                # Registration No exactly as printed
    param_type: str = ""            # "NTN" or "CNIC"
    name: str = ""
    business_name: str = ""
    filing_status: str = ""         # e.g. "Active (Late Filer)"
    category: str = ""              # Active / Active (Late Filer) / In-Active ...
    checking_date: str = ""         # normalised to DD-MMM-YYYY
    note: str = ""                  # the surcharge note on In-Active slips
    result: str = RES_PENDING
    rename_status: str = RN_PENDING
    new_file: str = ""
    timestamp: str = ""


# ── Folder listing ──────────────────────────────────────────────────────────

def list_pdfs(folder: str) -> List[str]:
    """Every *.pdf in `folder` (not recursive), sorted by name."""
    path = pathlib.Path(folder)
    if not path.is_dir():
        raise ValueError(f"Folder not found:\n{folder}")
    return sorted(
        (p.name for p in path.iterdir()
         if p.is_file() and p.suffix.lower() == ".pdf"),
        key=lambda s: s.lower(),
    )


def build_records(folder: str) -> List[Record]:
    """One blank Record per PDF, in file-name order."""
    return [Record(idx=i, source_file=name)
            for i, name in enumerate(list_pdfs(folder), start=1)]


# ── Text extraction ─────────────────────────────────────────────────────────

def read_text(pdf_path: pathlib.Path) -> str:
    """All text of a PDF as one string. Raises on an unreadable file."""
    with pdfplumber.open(str(pdf_path)) as pdf:
        return "\n".join((page.extract_text() or "") for page in pdf.pages)


def _strip_footer(text: str) -> str:
    """Drop the FBR SMS boiler-plate and the browser print header/footer."""
    cut = len(text)
    for marker in FOOTER_MARKERS:
        pos = text.find(marker)
        if pos != -1:
            cut = min(cut, pos)
    return text[:cut]


def parse_fields(text: str) -> Dict[str, str]:
    """
    Pull the labelled fields out of the slip text.

    The slip prints each field as "<Label>: <value>", but the value of an
    In-Active slip runs on for three more lines (the surcharge note) and
    the browser's own header/footer lines sit above and below. So instead
    of matching one line at a time, every label is located and its value
    is everything up to the next label.
    """
    out = {"reg_no": "", "name": "", "business_name": "",
           "filing_status": "", "checking_date": "", "note": "",
           "no_record": False}

    body = _strip_footer(text)

    if NO_RECORD_MARKER in body.lower():
        out["filing_status"] = "No record exists"
        out["no_record"] = True
        return out

    matches = list(LABEL_RE.finditer(body))
    key_for = {
        "Registration No": "reg_no",
        "Name": "name",
        "Business Name": "business_name",
        "Filing Status": "filing_status",
        "Filing Status Checking Date": "checking_date",
    }

    for n, m in enumerate(matches):
        key = key_for[m.group(1)]
        end = matches[n + 1].start() if n + 1 < len(matches) else len(body)
        value = " ".join(body[m.end():end].split())      # collapse newlines
        if key == "filing_status":
            value, out["note"] = _split_note(value)
        # The first occurrence wins: a slip never repeats a label, but the
        # print header can echo a URL that happens to contain one.
        if not out[key]:
            out[key] = value

    out["checking_date"] = normalise_date(out["checking_date"])
    if not out["name"]:
        out["name"] = out["business_name"]
    return out


def _split_note(value: str) -> Tuple[str, str]:
    """Separate "In-Active Note: <surcharge text>" into status and note."""
    pos = value.find(NOTE_MARKER)
    if pos == -1:
        return value.strip(), ""
    return value[:pos].strip(), value[pos + len(NOTE_MARKER):].strip()


def normalise_date(raw: str) -> str:
    """
    Turn whatever the slip printed into DD-MMM-YYYY.
    Accepts 31,Dec,2023 / 31-Dec-2023 / 31/12/2023 / 2023-12-31.
    Unrecognised text is returned unchanged.
    """
    s = raw.strip()
    if not s:
        return ""

    m = re.match(r"^(\d{1,2})\s*[,\-./ ]\s*([A-Za-z]{3,})\s*[,\-./ ]\s*(\d{4})$", s)
    if m and m.group(2)[:3].lower() in MONTHS:
        day, month, year = int(m.group(1)), MONTHS[m.group(2)[:3].lower()], int(m.group(3))
        return f"{day:02d}-{MONTH_NAMES[month - 1]}-{year}"

    m = re.match(r"^(\d{1,2})\s*[,\-./ ]\s*(\d{1,2})\s*[,\-./ ]\s*(\d{4})$", s)
    if m:
        day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1 <= month <= 12:
            return f"{day:02d}-{MONTH_NAMES[month - 1]}-{year}"

    m = re.match(r"^(\d{4})\s*[,\-./ ]\s*(\d{1,2})\s*[,\-./ ]\s*(\d{1,2})$", s)
    if m:
        year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1 <= month <= 12:
            return f"{day:02d}-{MONTH_NAMES[month - 1]}-{year}"

    return s


def categorise(filing_status: str) -> str:
    """Bucket the printed status so the Excel file can be filtered on it."""
    low = filing_status.lower()
    if not low:
        return CAT_UNKNOWN
    if "no record" in low:
        return CAT_NO_RECORD
    if "in-active" in low or "inactive" in low:
        return CAT_INACTIVE
    if "late filer" in low:
        return CAT_LATE_FILER
    if "active" in low:
        return CAT_ACTIVE
    return CAT_UNKNOWN


def classify(reg_no: str) -> str:
    """
    Registration number -> "NTN" or "CNIC".

      13 digits         -> CNIC
      7 characters      -> NTN  (digits, or with a letter: A123456)
      5-8 characters    -> NTN
      anything else     -> ""
    """
    clean = re.sub(r"[^A-Za-z0-9]", "", reg_no)
    if not clean:
        return ""
    if clean.isdigit() and len(clean) == 13:
        return "CNIC"
    if 5 <= len(clean) <= 8:
        return "NTN"
    return ""


def is_atl_slip(text: str) -> bool:
    """True if the PDF looks like an ATL verification slip at all."""
    low = text.lower()
    return REQUIRED_MARKER.lower() in low or NO_RECORD_MARKER in low


# ── Target file name for the rename step ────────────────────────────────────

_BAD_NAME_CHARS = re.compile(r'[<>:"/\\|?*]')


def rename_target(rec: Record, pattern: str) -> str:
    """
    Build the new file name from a pattern. Supported placeholders:
        {reg}   Registration No
        {date}  Checking date (DD-MMM-YYYY)
        {name}  Name
    """
    if not rec.reg_no:
        return ""
    name = (pattern or "{reg}").format(
        reg=rec.reg_no,
        date=rec.checking_date or "",
        name=rec.name or "",
    )
    name = _BAD_NAME_CHARS.sub("-", name).strip(" .")
    if not name:
        return ""
    return name if name.lower().endswith(".pdf") else name + ".pdf"
