# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A desktop GUI tool (Python/tkinter) for managing automotive customer lead lists: merging multiple priority-ranked Excel files, deduplicating contacts, filtering recent purchasers, and generating/importing short URL templates.

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

The entire application lives in `lead_list_tool.py`. The structure is:

**Data layer:**
- `Record` dataclass — one customer contact, with source file metadata and multi-variant phone/ID fields
- `UnionFind` — groups duplicate records across files using phone numbers, ONE ID, and LINE ID as keys

**Processing pipeline (4 steps, each triggered by GUI):**
1. `merge_files()` — reads `G1_*.xlsx`, `G2_*.xlsx`, … from a folder; lower G-number = higher priority; deduplicates via UnionFind; winner fills in missing fields from lower-priority records
2. `remove_recent_orders()` — cross-references an orders Excel; removes matched contacts to a `Removed_RecentOrders` sheet
3. `export_phone_template()` — exports SMS phone numbers from Working_List to a `Phone`-column CSV for uploading to the short URL platform
4. `collect_short_urls()` — reads the platform-returned CSV (`No, Phone, Url, Count`) and writes matched URLs and click counts back into Working_List; URL column is write-once (skip if exists), Count column always overwrites

**Output workbook sheets:** `Working_List`, `Merged_Master`, `Dropped_Duplicates`, `Filtered_DNC`, `Removed_RecentOrders`, `ShortURL_Log`, `Manifest`, `README`

**GUI layer:**
- `LeadListGUI` class wraps everything in a tkinter window with file/folder dialogs and a scrollable log pane
- `run_app()` is the entry point called from `__main__`

**Key utilities:** `normalize_phone()`, `normalize_id()`, `normalize_header()`, `clean_text()`, `file_sha1()`, `auto_width()`

## Data Conventions

- Input files follow the naming pattern `G{n}_*.xlsx` where `n` is the group priority (1 = highest)
- Phone numbers are stored in up to 4 variants (CR/member phone, SMS/sales phone) per record
- Null-like values (`""`, `"nan"`, `"none"`, `"null"`, `"-"`) are normalized to `""` by `clean_text()`
- DNC values (`"不聯繫"`, `"電話不聯繫"`, `"簡訊不聯繫"`, `"個資未授權"`, `"該電話不可聯繫"`, etc.) trigger filtering into `Filtered_DNC` sheet
- Records where all phone fields AND LINE ID are empty are excluded during merge
- The `Manifest` sheet records SHA1 hashes and timestamps for all input files
- Short URL platform CSV format: `No, Phone, Url, Count`
- CANONICAL_COLUMNS includes vehicle info fields: `現保有車款_T`, `現保有車交車年份_T`, `現保有車款_L`, `現保有車交車年份_L`
