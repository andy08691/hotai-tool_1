# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A desktop GUI tool (Python/tkinter) for managing automotive customer lead lists: merging multiple priority-ranked Excel files, deduplicating contacts, filtering recent purchasers, and enriching records with short URLs.

## Setup & Running

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python lead_list_tool.py
```

**Single dependency:** `openpyxl>=3.1.0`

No test suite or linter is configured.

## Packaging

```bash
# macOS
pyinstaller --windowed --onefile lead_list_tool.py

# Windows
pyinstaller --noconsole --onefile lead_list_tool.py
```

## Architecture

The entire application lives in `lead_list_tool.py` (~941 lines). The structure is:

**Data layer:**
- `Record` dataclass — one customer contact, with source file metadata and multi-variant phone/ID fields
- `UnionFind` — groups duplicate records across files using phone numbers, ONE ID, and LINE ID as keys

**Processing pipeline (3 steps, each triggered by GUI):**
1. `merge_files()` — reads `G1_*.xlsx`, `G2_*.xlsx`, … from a folder; lower G-number = higher priority; deduplicates via UnionFind; winner fills in missing fields from lower-priority records
2. `remove_recent_orders()` — cross-references an orders Excel; removes matched contacts to a `Removed_RecentOrders` sheet
3. `add_short_urls()` — reads a CSV/Excel mapping file and fills a target column with matched short URLs; logs match methodology to `ShortURL_Log`

**Output workbook sheets:** `Working_List`, `Merged_Master`, `Dropped_Duplicates`, `Removed_RecentOrders`, `ShortURL_Log`, `Manifest`, `README`

**GUI layer:**
- `LeadListGUI` class wraps everything in a tkinter window with file/folder dialogs and a scrollable log pane
- `run_app()` is the entry point called from `__main__`

**Key utilities:** `normalize_phone()`, `normalize_id()`, `normalize_header()`, `clean_text()`, `file_sha1()`, `auto_width()`

## Data Conventions

- Input files follow the naming pattern `G{n}_*.xlsx` where `n` is the group priority (1 = highest)
- Phone numbers are stored in up to 4 variants (CR/member phone, SMS/sales phone) per record
- Null-like values (`""`, `"nan"`, `"none"`, `"null"`, `"-"`) are normalized to `""` by `clean_text()`
- The `Manifest` sheet records SHA1 hashes and timestamps for all input files
