"""
ATL Extractor — Extraction Engine
=================================
Walks a folder of ATL verification PDFs and, for each one:

  1. reads the text with pdfplumber,
  2. pulls out Registration No / Name / Business Name / Filing Status /
     Checking Date,
  3. writes the running result into the Excel summary,
  4. reports progress back to the GUI through a queue.

A second, separate job renames the PDFs to their registration number —
that one changes files on disk, so the GUI always asks first.

Both jobs are designed to run inside a background daemon thread.
"""

from __future__ import annotations

import datetime
import pathlib
import queue
import threading
import time
from typing import List

from constants import (
    INTER_FILE_DELAY_SECS, REPORT_SAVE_EVERY,
    RES_PENDING, RES_PROCESSING, RES_EXTRACTED, RES_PARTIAL,
    RES_NO_RECORD, RES_NOT_ATL, RES_FAILED,
    RN_RENAMED, RN_SAME, RN_EXISTS, RN_NO_REG, RN_FAILED,
    DONE_RESULTS,
)
from pdf_reader import (
    Record, build_records, read_text, parse_fields, categorise, classify,
    is_atl_slip, rename_target,
)
from report_writer import write_report


def _now() -> str:
    return datetime.datetime.now().strftime("%H:%M:%S")


def _now_iso() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class _Job:
    """Shared plumbing: queue messages, pause/stop handling, logging.

    Messages put on `out_q` (all tuples):
        ("log", level, text)      level = info | ok | err | warn
        ("record", Record)        a record changed
        ("report", path)          the Excel file was (re)written
        ("done", None)            the job is over
    """

    def __init__(self, out_q: queue.Queue, stop_event: threading.Event,
                 pause_event: threading.Event):
        self.out_q = out_q
        self.stop_event = stop_event
        self.pause_event = pause_event

    def _log(self, msg: str, level: str = "info"):
        self.out_q.put(("log", level, f"[{_now()}] {msg}"))

    def _emit(self, rec: Record):
        self.out_q.put(("record", rec))

    def _wait_if_paused(self) -> bool:
        """Block while paused. Returns False if the job should stop."""
        while self.pause_event.is_set() and not self.stop_event.is_set():
            time.sleep(0.2)
        return not self.stop_event.is_set()


# ══════════════════════════════════════════════════════════════════════════
#  Extraction
# ══════════════════════════════════════════════════════════════════════════

class ExtractWorker(_Job):
    """Reads every PDF in `pdf_dir` and writes one Excel summary."""

    def __init__(
        self,
        records: List[Record],
        pdf_dir: pathlib.Path,
        output_dir: pathlib.Path,
        out_q: queue.Queue,
        stop_event: threading.Event,
        pause_event: threading.Event,
        output_name: str = "",
        skip_done: bool = False,
    ):
        super().__init__(out_q, stop_event, pause_event)
        self.records = records
        self.pdf_dir = pathlib.Path(pdf_dir)
        self.output_dir = pathlib.Path(output_dir)
        self.output_name = output_name
        self.skip_done = skip_done
        self._report_warned = False
        self.report_path = ""

    # ── Public entry point ────────────────────────────────────────────────

    def run(self):
        """Call this in a daemon thread."""
        try:
            self._run_inner()
        except Exception as exc:                            # noqa: BLE001
            self._log(f"Fatal error: {exc}", "err")
        finally:
            for rec in self.records:            # stopped mid-record
                if rec.result == RES_PROCESSING:
                    rec.result = RES_PENDING
                    self._emit(rec)
            self._save_report()
            self.out_q.put(("done", None))

    # ── Core loop ─────────────────────────────────────────────────────────

    def _run_inner(self):
        self.output_dir.mkdir(parents=True, exist_ok=True)

        pending = [r for r in self.records
                   if not (self.skip_done and r.result in DONE_RESULTS)]
        skipped = len(self.records) - len(pending)
        if skipped:
            self._log(f"{skipped} file(s) skipped (already extracted).")
        if not pending:
            self._log("Nothing to do — every file is already extracted.", "ok")
            return

        total = len(pending)
        self._log(f"Reading {total} PDF file(s) from {self.pdf_dir}")

        for n, rec in enumerate(pending, start=1):
            if not self._wait_if_paused():
                self._log("Run cancelled by user.", "warn")
                break

            rec.result = RES_PROCESSING
            self._emit(rec)

            self._process(rec)
            rec.timestamp = _now_iso()
            self._emit(rec)

            if rec.result in DONE_RESULTS:
                self._log(
                    f"[{n}/{total}] {rec.reg_no or rec.source_file} — "
                    f"{rec.filing_status or 'Unknown'} — {rec.name}",
                    "ok" if rec.result == RES_EXTRACTED else "warn")
            else:
                self._log(f"[{n}/{total}] {rec.source_file} — {rec.result}", "err")

            if n % REPORT_SAVE_EVERY == 0:
                self._save_report()
            if INTER_FILE_DELAY_SECS and n < total:
                time.sleep(INTER_FILE_DELAY_SECS)

        self._log("Extraction finished.", "ok")

    # ── Single file ───────────────────────────────────────────────────────

    def _process(self, rec: Record):
        path = self.pdf_dir / rec.source_file
        try:
            text = read_text(path)
        except Exception as exc:                            # noqa: BLE001
            reason = str(exc).splitlines()[0][:120] or exc.__class__.__name__
            rec.result = f"{RES_FAILED} - {reason}"
            return

        if not text.strip():
            rec.result = f"{RES_FAILED} - no text in PDF (scanned image?)"
            return
        if not is_atl_slip(text):
            rec.result = RES_NOT_ATL
            return

        fields = parse_fields(text)
        rec.reg_no        = fields["reg_no"]
        rec.name          = fields["name"]
        rec.business_name = fields["business_name"]
        rec.filing_status = fields["filing_status"]
        rec.checking_date = fields["checking_date"]
        rec.note          = fields["note"]
        rec.category      = categorise(rec.filing_status)
        rec.param_type    = classify(rec.reg_no)

        if fields["no_record"]:
            rec.result = RES_NO_RECORD
        elif rec.reg_no and rec.filing_status and rec.checking_date:
            rec.result = RES_EXTRACTED
        else:
            missing = [label for label, value in (
                ("Registration No", rec.reg_no),
                ("Filing Status", rec.filing_status),
                ("Checking Date", rec.checking_date),
            ) if not value]
            rec.result = RES_PARTIAL
            self._log(f"  {rec.source_file}: missing {', '.join(missing)}", "warn")

    # ── Excel ─────────────────────────────────────────────────────────────

    def _save_report(self):
        try:
            self.report_path = write_report(
                self.records, self.output_dir, self.output_name)
            self.out_q.put(("report", self.report_path))
            self._report_warned = False
        except Exception as exc:                            # noqa: BLE001
            if not self._report_warned:      # e.g. the file is open in Excel
                self._report_warned = True
                self._log("Could not write the Excel file (is it open in "
                          f"Excel? close it and it will be written): {exc}",
                          "warn")


# ══════════════════════════════════════════════════════════════════════════
#  Renaming  (changes files on disk — the GUI confirms first)
# ══════════════════════════════════════════════════════════════════════════

class RenameWorker(_Job):
    """Renames each PDF to its registration number."""

    def __init__(
        self,
        records: List[Record],
        pdf_dir: pathlib.Path,
        out_q: queue.Queue,
        stop_event: threading.Event,
        pause_event: threading.Event,
        pattern: str = "{reg}",
        output_dir: pathlib.Path = None,
        output_name: str = "",
    ):
        super().__init__(out_q, stop_event, pause_event)
        self.records = records
        self.pdf_dir = pathlib.Path(pdf_dir)
        self.pattern = pattern
        self.output_dir = pathlib.Path(output_dir) if output_dir else None
        self.output_name = output_name

    def run(self):
        """Call this in a daemon thread."""
        try:
            self._run_inner()
        except Exception as exc:                            # noqa: BLE001
            self._log(f"Fatal error: {exc}", "err")
        finally:
            self._refresh_report()
            self.out_q.put(("done", None))

    def _refresh_report(self):
        """Rewrite the Excel file so its "PDF File" column matches disk."""
        if not self.output_dir or not self.output_name:
            return
        try:
            path = write_report(self.records, self.output_dir, self.output_name)
            self.out_q.put(("report", path))
        except Exception as exc:                            # noqa: BLE001
            self._log("Could not update the Excel file with the new names "
                      f"(is it open in Excel?): {exc}", "warn")

    def _run_inner(self):
        targets = [r for r in self.records if r.result in DONE_RESULTS]
        total = len(targets)
        if not total:
            self._log("Nothing to rename — run the extraction first.", "warn")
            return

        self._log(f"Renaming {total} file(s) ...")
        renamed = 0

        for n, rec in enumerate(targets, start=1):
            if not self._wait_if_paused():
                self._log("Renaming cancelled by user.", "warn")
                break

            new_name = rename_target(rec, self.pattern)
            if not new_name:
                rec.rename_status = RN_NO_REG
                self._log(f"[{n}/{total}] {rec.source_file} — no registration "
                          "number, left alone", "warn")
                self._emit(rec)
                continue

            if new_name.lower() == rec.source_file.lower():
                rec.rename_status = RN_SAME
                rec.new_file = rec.source_file
                self._emit(rec)
                continue

            old_path = self.pdf_dir / rec.source_file
            new_path = self.pdf_dir / new_name

            if not old_path.exists():
                rec.rename_status = RN_FAILED
                self._log(f"[{n}/{total}] {rec.source_file} — file is gone", "err")
                self._emit(rec)
                continue
            if new_path.exists():
                rec.rename_status = RN_EXISTS
                self._log(f"[{n}/{total}] {new_name} already exists — "
                          f"{rec.source_file} left alone", "warn")
                self._emit(rec)
                continue

            try:
                old_path.rename(new_path)
            except Exception as exc:                        # noqa: BLE001
                rec.rename_status = RN_FAILED
                self._log(f"[{n}/{total}] {rec.source_file} — "
                          f"{str(exc).splitlines()[0][:100]}", "err")
            else:
                rec.rename_status = RN_RENAMED
                rec.new_file = new_name
                rec.source_file = new_name       # keep pointing at the real file
                renamed += 1
                self._log(f"[{n}/{total}] {new_name}", "ok")
            self._emit(rec)

        self._log(f"Renaming finished — {renamed} of {total} file(s) renamed.",
                  "ok")
