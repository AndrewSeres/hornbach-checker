# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

### Local setup
```bash
pip install -r requirements.txt
playwright install chromium
```

### Run the CI scraper (writes to Google Sheets)
```bash
export GOOGLE_SHEETS_CREDS=$(cat service_account.json)
export GOOGLE_SHEET_ID=your_sheet_id
python scraper.py
```

### Run the desktop GUI tool (exports to Excel)
```bash
python hornbach_checker.py
```

## Architecture

This project has two independent entry points that share the same scraping logic:

**`scraper.py`** — CI/headless script for GitHub Actions. Scrapes hornbach.sk for wood briquette products, then writes results to three Google Sheets tabs: `Posledný beh` (latest state), `História` (append-only column per run), and `Rozdiel` (diff between last two runs). Triggered every Friday at ~16:00 CET via `.github/workflows/scrape.yml`.

**`hornbach_checker.py`** — Local desktop tool with a tkinter GUI. Contains the same scraper logic plus an Excel export (openpyxl) that saves a formatted `.xlsx` to the Desktop. Falls back to headless CLI if tkinter is unavailable.

**`index.html`** — Static live viewer. Fetches data directly from the public Google Sheet via the `gviz/tq?tqx=out:csv` API (no backend required). Has three tabs: current stock table + bar chart, diff comparison, and trend line chart. The Sheet ID is hardcoded in the JS (`SHEET_ID` constant). Requires the Google Sheet to be shared publicly as "Anyone with the link → Viewer".

### Key shared patterns
- Both Python scripts define the same `CANONICAL_STORES`, `KEYWORDS`, `EXCLUDE_KEYWORDS`, `canonicalize_store()`, and `store_sort_key()` — changes to store list or product filtering must be kept in sync across both files.
- Stock parsing from the availability modal uses line-by-line regex heuristics (look for "X balení/ks", standalone digits with contextual keywords, or out-of-stock phrases).
- Google Sheets credentials come from `GOOGLE_SHEETS_CREDS` env var (JSON string) in CI, or from a local `service_account.json` file for development.
