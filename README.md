# ATL Extractor

Extract and structure Active Taxpayer List data into spreadsheets for vendor/customer screening.

Reads a folder of FBR **ATL Status** verification PDFs and writes one Excel workbook holding the details from every slip. It can then rename each PDF to its registration number so the files line up with the spreadsheet rows.

Everything happens locally — the PDFs are parsed on your own machine and nothing is uploaded.

---

## Requirements

- Windows 10/11
- Python 3.10+

Dependencies are installed on first run. `run.bat` will install Python itself via `winget` if it is missing.

## Install and run

```bash
git clone https://github.com/alphaledger-automation/ATL_Extractor.git
cd ATL_Extractor
pip install -r requirements.txt
python main.py
```

Or double-click **`run.bat`** on Windows; `main.py` installs anything missing (`pdfplumber`, `openpyxl`, `customtkinter`, `pillow`).

### Command line

```bash
python main.py --cli "C:\path\to\ATL Status" ["C:\output\folder"]
python main.py --setup-only
```

`--cli` extracts a folder and writes the workbook without opening the window.

## Input

A folder of ATL Status PDFs downloaded from IRIS. The `ATL Status` folder next to the app is picked up automatically; use **Browse…** for any other location.

PDFs must have a text layer — scans and screenshots are reported as `Failed - no text in PDF (scanned image?)` rather than silently skipped.

## Output

```
<output folder>/ATL_Extract_<date>_<time>.xlsx
```

**Sheet `ATL Status`** — one row per PDF, colour-coded by filing status, with a filter row and frozen headers:

`Registration No | Length | Type (NTN/CNIC) | Name | Business Name | Filing Status | Category | Filing Status Checking Date | PDF File | Result | Note`

Registration numbers are stored as text, so leading zeros (`0123456`), letter prefixes (`A123456`) and 13-digit CNICs stay intact.

**Sheet `Summary`** — counts per filing status, files found, files fully extracted, files that could not be read.

The workbook is written as the run progresses, so it is always up to date. Keep it closed in Excel during a run.

## Filing status handling

The slip prints status as free text. It is copied verbatim into `Filing Status` and also bucketed into `Category` so the sheet can be filtered:

| Slip text | Category |
|---|---|
| Active | Active |
| Active (Late Filer) | Active (Late Filer) |
| In-Active | In-Active (the surcharge note on the same line goes to `Note`) |
| No record exists | No record |

`Active` and `Active (Late Filer)` are counted separately — a late filer is on the ATL but did not file on time.

## Renaming PDFs

After a run, the Dashboard offers three name patterns:

- `Registration No` → `0123456.pdf`
- `Registration No + date` → `0123456_31-Dec-2023.pdf`
- `Registration No + name` → `0123456_ALPHA TRADERS.pdf`

A confirmation box shows how many files will change and an example name before anything is touched. A file is left alone if it has no registration number or if the target name is taken. The workbook is rewritten afterwards so its `PDF File` column matches disk. Running it twice is a no-op (`Already named`).

## Features

- Dashboard page: PDF folder, output folder, output file name
- Current Status page: Read / Total counter and a live table per PDF
- Pause / Resume and Cancel, applied after the file in flight finishes
- **Skip files already extracted** — a second Start re-reads only what failed or was never reached
- **Open Excel when finished**
- Light / Dark theme, remembered between launches
- Partial extractions reported as `Extracted (incomplete)` with the missing field named in the activity log

## Project layout

| File | Purpose |
|---|---|
| `run.bat` | Windows launcher |
| `main.py` | Entry point; `--cli`, `--setup-only` |
| `bootstrap.py` | Dependency check and installer |
| `gui.py` | Dashboard / Current Status window |
| `pdf_engine.py` | Extraction and renaming jobs (background thread) |
| `pdf_reader.py` | Lists the PDFs and parses slip text |
| `report_writer.py` | Writes the Excel workbook |
| `constants.py` | Labels, patterns, palette |
| `Assets/` | Application icons |

`USER_GUIDE.txt` holds the end-user walkthrough and troubleshooting notes.

## Notes

- No taxpayer data or sample PDFs are included in this repository.
- Setup log: `%LocalAppData%\ATL_Extractor\setup.log`

## License

MIT — see [LICENSE](LICENSE). Copyright (c) 2026 Alpha Edge Advisory.
