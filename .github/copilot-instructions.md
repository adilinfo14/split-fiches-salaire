# Copilot Instructions for split-fiches-salaire

## Project Overview
A Python desktop utility for splitting multi-page French salary slip PDFs into individual files named by employee AVS number and pay period. The tool auto-detects payslip boundaries using regex pattern matching for "Période" (mm.YYYY) and "AVS" (XXX.XXXX.XXXX.XX), then exports renamed PDFs with comprehensive logging and error tracking.

## Architecture

### Single-File Design
- **[src/split_fiches.py](src/split_fiches.py)** (~740 lines): All code in one module containing:
  - **PDF parsing & extraction**: `extract_filename_year_month_avs()` uses hardcoded regex patterns to detect payslip headers
  - **Core split logic**: `split_pdf()` implements two modes:
    - `group_multipage=True` (default): Groups consecutive pages into single files when no AVS/Période detected on continuation pages (handles multi-page slips)
    - `group_multipage=False`: Saves each page individually, useful for already-separated docs
  - **Tkinter GUI**: `AppUI` class provides interactive file picker, progress bar, results table with threading for non-blocking operations
  - **CLI mode**: `run_cli()` for headless execution, `main()` routes between UI and CLI

### Data Flow
1. User selects PDF (GUI) or provides path (CLI)
2. `split_pdf()` iterates through pages, extracting text
3. For each page: regex matches detect Période (MM.YYYY) → year/month, AVS (XXX.XXXX.XXXX.XX)
4. Matched pages start new file; unmatched pages either append to current file or become orphans
5. Files written to `output/{timestamp}/`, errors to `errors/{timestamp}/`, logs to `logs/`
6. Results recorded in `Record` dataclass, exported to CSV

### Directory Structure
- **project_root/** (defined by `project_root()` from file location)
  - **input/** (folder for user to place source PDFs)
  - **output/** → split_TIMESTAMP/ (successfully named files)
  - **errors/** → split_TIMESTAMP/ (files missing Période/AVS, orphans, exceptions)
  - **logs/** (CSV and .log files from each run)

## Key Patterns & Conventions

### Naming Format
- **Expected input**: Any PDF name
- **Output pattern**: `{YYYY}-{month_fr}-{AVS}.pdf` (e.g., `2026-janvier-756.1234.5678.97.pdf`)
- Month names are French (`janvier`, `février`, etc.), extracted from MONTHS_FR dict
- Duplicate filenames get `_p{start_page:03d}` suffix to avoid overwrites

### Record Status Tracking
Four statuses in CSV export:
- **OK**: Detected Période + AVS, file successfully written to output/
- **FALLBACK**: Missing Période/AVS metadata, moved to errors/ (orphan or undetectable)
- **ERROR**: Exception during PDF writing or text extraction (syntax errors, corrupt pages)
- **ORPHAN**: Page appearing before any valid payslip header (multi-page PDFs)

### UI/CLI Threading Pattern
GUI runs heavy operations (`split_pdf()`) on daemon thread via `threading.Thread` with progress callback passed through `progress_cb` parameter. Main thread updates via `self.master.after()` to avoid blocking.

## Development Commands

### Environment Setup
```bash
cd "C:\path\to\split-fiches-salaire"
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### Run Application
```bash
# GUI mode (default)
python src/split_fiches.py

# CLI mode with automatic multi-page grouping
python src/split_fiches.py "path/to/file.pdf"

# CLI mode, disable multi-page grouping
python src/split_fiches.py "path/to/file.pdf" --no-multipage
```

### Build EXE (PyInstaller)
```bash
pip install pyinstaller
pyinstaller split-fiches-salaire.spec
# Output: dist/split-fiches-salaire.exe
```

## Critical Implementation Details

### Regex Patterns (case-sensitive, UTF-8)
- **Période**: `r"Période\s*:\s*(\d{2})\.(\d{4})"` — Captures MM and YYYY from text like "Période : 12.2025"
- **AVS**: `r"\b\d{3}\.\d{4}\.\d{4}\.\d{2}\b"` — Swiss AVS format (no capture groups, returns full match)
- Both patterns must exist on same page to start new payslip; orphans accumulate until next match

### CSV Export Fields
Semicolon-delimited; encoding UTF-8:
```
status | year | month | avs | pages | output_file | output_path | note
```
Pages formatted as `"1"` (single) or `"1-3"` (range), year/month/avs as "-" if unknown

### Threading Safety
- Main Tkinter loop guarded by `self.master.after()` callback from worker thread
- `self.records` list populated after `split_pdf()` completes, safe access during `_finish()`
- No locks needed (Python GIL + single result write)

## When Modifying Core Logic

1. **Changing regex patterns**: Update `MONTHS_FR` dict and regex strings (lines 21–34). Test against actual PDF samples.
2. **Adding statuses**: Update `Record` status field (line 139), UI table column handling, CSV export validation.
3. **Handling errors**: All page-level exceptions caught in try-except blocks; error types logged but not raised (graceful degradation).
4. **GUI layout changes**: Edit `AppUI._build()` method; remember `columnconfigure(weight=1)` for responsive layout.
5. **File paths**: Always use `Path` object with `.resolve()` for absolute paths; `.parent.mkdir(parents=True)` before writes.

## Common Issues & Solutions

- **Regex not matching**: Verify PDF contains literal "Période : MM.YYYY" text (some PDFs may use different formatting)
- **Empty output/**: Check console log for ERROR status; most likely text extraction failed or regex patterns don't match
- **Duplicate filenames**: Code automatically appends `_p{page_num}` suffix; check log for warnings
- **App won't run**: Ensure PyPDF is installed (`pip install pypdf`); Tkinter included with Python on Windows
