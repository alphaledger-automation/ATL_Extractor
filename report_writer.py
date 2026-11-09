"""
ATL Extractor — Excel Writer
============================
Writes ATL_Extract_<timestamp>.xlsx with one row per PDF:

  # | Registration No | Length | Type | Name | Business Name |
  Filing Status | Category | Filing Status Checking Date | PDF File |
  Result | Note

Sheet 2 ("Summary") holds the counts per category so the file can be
skimmed without filtering.

The file is re-saved while the run is in progress, so it is always up to
date — keep it closed in Excel during a run.
"""

from __future__ import annotations

import datetime
import pathlib
from typing import List

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from constants import (
    APP_NAME, APP_VERSION, CATEGORIES, DONE_RESULTS,
    RES_PENDING, RES_PROCESSING, RES_EXTRACTED,
)

# ── Styles ──────────────────────────────────────────────────────────────────
_HEADER_FILL  = PatternFill("solid", fgColor="1A5530")
_HEADER_FONT  = Font(bold=True, color="FFFFFF")
_HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

_FILL_ACTIVE   = PatternFill("solid", fgColor="E6F6EC")
_FILL_LATE     = PatternFill("solid", fgColor="FFF6DC")
_FILL_INACTIVE = PatternFill("solid", fgColor="FCE6E8")
_FILL_PROBLEM  = PatternFill("solid", fgColor="EEEEEE")

_CATEGORY_FILL = {
    "Active":              _FILL_ACTIVE,
    "Active (Late Filer)": _FILL_LATE,
    "In-Active":           _FILL_INACTIVE,
    "No record":           _FILL_PROBLEM,
    "Unknown":             _FILL_PROBLEM,
}

_HEADERS = ["#", "Registration No", "Length", "Type", "Name", "Business Name",
            "Filing Status", "Category", "Filing Status Checking Date",
            "PDF File", "Result", "Note"]
_COL_WIDTHS = [6, 20, 9, 8, 34, 32, 20, 20, 24, 40, 24, 60]

_SHEET_DATA    = "ATL Status"
_SHEET_SUMMARY = "Summary"


def default_output_name() -> str:
    """ATL_Extract_<YYYY-MM-DD_HH-MM-SS>.xlsx — unique per run."""
    stamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return f"ATL_Extract_{stamp}.xlsx"


def report_path(output_dir: pathlib.Path, output_name: str) -> pathlib.Path:
    name = output_name or default_output_name()
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return pathlib.Path(output_dir) / name


def _result_text(result: str) -> str:
    if result in (RES_PENDING, RES_PROCESSING):
        return RES_PENDING
    return result


def write_report(records: List, output_dir: pathlib.Path,
                 output_name: str = "") -> str:
    """Write the workbook and return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = _SHEET_DATA

    for col_idx, header in enumerate(_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font      = _HEADER_FONT
        cell.fill      = _HEADER_FILL
        cell.alignment = _HEADER_ALIGN
    for col_idx, width in enumerate(_COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[1].height = 30

    for row, rec in enumerate(records, start=2):
        ws.cell(row=row, column=1, value=rec.idx)
        # Registration numbers stay text: keeps leading zeros (0123456),
        # letters (A123456) and stops CNICs turning into 1.12E+12.
        reg_cell = ws.cell(row=row, column=2, value=rec.reg_no or None)
        reg_cell.number_format = "@"
        ws.cell(row=row, column=3, value=len(rec.reg_no) if rec.reg_no else None)
        ws.cell(row=row, column=4, value=rec.param_type or None)
        ws.cell(row=row, column=5, value=rec.name or None)
        ws.cell(row=row, column=6, value=rec.business_name or None)
        ws.cell(row=row, column=7, value=rec.filing_status or None)
        ws.cell(row=row, column=8, value=rec.category or None)
        ws.cell(row=row, column=9, value=rec.checking_date or None)
        ws.cell(row=row, column=10, value=rec.source_file)
        ws.cell(row=row, column=11, value=_result_text(rec.result))
        ws.cell(row=row, column=12, value=rec.note or None)

        fill = _CATEGORY_FILL.get(rec.category)
        if rec.result not in DONE_RESULTS:
            fill = _FILL_PROBLEM
        if fill:
            for col_idx in range(1, len(_HEADERS) + 1):
                ws.cell(row=row, column=col_idx).fill = fill

    last_row = max(len(records) + 1, 2)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(_HEADERS))}{last_row}"

    _write_summary(wb.create_sheet(_SHEET_SUMMARY), records)

    path = report_path(output_dir, output_name)
    wb.save(str(path))
    return str(path)


def _write_summary(ws, records: List) -> None:
    """A small counts-per-category sheet."""
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 14

    title = ws.cell(row=1, column=1, value=f"{APP_NAME} v{APP_VERSION} — Summary")
    title.font = Font(bold=True, size=12)
    ws.cell(row=2, column=1,
            value="Generated " + datetime.datetime.now().strftime("%d-%b-%Y %H:%M:%S"))

    row = 4
    for header in ("Filing Status", "Count"):
        cell = ws.cell(row=row, column=1 if header == "Filing Status" else 2,
                       value=header)
        cell.font      = _HEADER_FONT
        cell.fill      = _HEADER_FILL
        cell.alignment = _HEADER_ALIGN

    counts = {cat: 0 for cat in CATEGORIES}
    extracted = failed = 0
    for rec in records:
        if rec.result in DONE_RESULTS:
            counts[rec.category if rec.category in counts else "Unknown"] += 1
            if rec.result == RES_EXTRACTED:
                extracted += 1
        elif rec.result not in (RES_PENDING, RES_PROCESSING):
            failed += 1

    row = 5
    for cat in CATEGORIES:
        if counts[cat]:
            ws.cell(row=row, column=1, value=cat)
            ws.cell(row=row, column=2, value=counts[cat])
            fill = _CATEGORY_FILL.get(cat)
            if fill:
                ws.cell(row=row, column=1).fill = fill
                ws.cell(row=row, column=2).fill = fill
            row += 1

    row += 1
    for label, value in (("PDF files found", len(records)),
                         ("Fully extracted", extracted),
                         ("Could not be read", failed)):
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row, column=2, value=value).font = Font(bold=True)
        row += 1
