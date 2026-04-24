# SEC Financial Fetcher — Design Spec
Date: 2026-04-13

## Overview

A Windows Tkinter GUI tool that fetches structured financial statements from SEC EDGAR (via edgartools) and saves them as Excel files, one per company. Designed for a stock analyst who maintains a watchlist and needs complete GAAP financials updated each quarter, with Non-GAAP metrics extracted via AI from 8-K press releases.

---

## Goals

1. Fetch all available historical quarterly financial statements for any US-listed company
2. Save to Excel in a stable structure that survives updates without breaking user formulas
3. Support single-company and batch (watchlist) modes
4. Phase 1: GAAP data via edgartools. Phase 2: Non-GAAP via Gemini API from 8-K filings
5. Tool must be runnable by double-clicking a BAT file with no technical knowledge required

---

## Architecture

### File Structure

```
sec-fetcher/
├── 啟動器.bat              # Thin launcher, 2 lines only, English only
├── launcher.ps1            # All Chinese UI, env checks, installs venv
├── main.py                 # Tkinter GUI entry point
├── fetcher_gaap.py         # Fetches XBRL financial statements via edgartools
├── fetcher_nongaap.py      # Fetches Non-GAAP data from 8-K via AI API
├── excel_writer.py         # Writes/updates Excel files (Data_* sheets only)
├── config.example.json     # Committed to git — template with empty values
├── config.json             # Gitignored — real API key + watchlist
├── company_cache.json      # Committed to git — ticker → company name cache
├── requirements.txt
├── .gitignore
├── README.md
├── ARCHITECTURE.md
├── CHANGELOG.md
├── PITFALLS.md
└── output/                 # Gitignored — generated Excel files
    └── AAPL.xlsx
```

### .gitignore Contents
```
config.json
output/
venv/
__pycache__/
*.pyc
*.log
```

---

## GUI Design

### Main Window — Two Tabs

**Tab 1: 單一公司**
- Ticker input field with placeholder text
- Checkboxes: GAAP only / Non-GAAP only / Both
- [執行] button
- Progress log area (scrollable)

**Tab 2: 批量更新**
- Lists all watchlist companies with checkboxes
- [全選] / [全不選] buttons
- [開始批量更新] button
- Progress log area with per-company status

**Persistent buttons (always visible):**
- [管理 Watchlist] — opens popup window
- [進階設定] — opens popup window

### Watchlist Popup Window
```
┌──────────────────────────────────────────┐
│           管理 Watchlist                 │
│  ──────────────────────────────────────  │
│  AAPL   Apple Inc.                 [x]  │
│  MSFT   Microsoft Corporation      [x]  │
│  ──────────────────────────────────────  │
│  新增：[________] [查詢]                 │
│  → 查到：Apple Inc. — [加入] [取消]      │
│  ──────────────────────────────────────  │
│  [更新名稱庫]  上次更新：2026-04-13      │
│                              [關閉]      │
└──────────────────────────────────────────┘
```

- Add: enter ticker → click 查詢 → shows company name for confirmation → [加入]
- Remove: click [x] next to company
- 更新名稱庫: re-queries EDGAR for all watchlist tickers, updates company_cache.json
- Name lookup: checks company_cache.json first (instant), falls back to live EDGAR query if not found
- All changes auto-save to config.json

### Advanced Settings Popup
- AI Provider: dropdown (Google Gemini / OpenAI / Anthropic)
- Model name: text field, auto-fills with provider default when switched
  - Google Gemini → `gemini-flash-latest`
  - OpenAI → `gpt-4o-mini`
  - Anthropic → `claude-haiku-4-5-20251001`
- API Key: masked input with show/hide toggle (pattern_secret_entry.py)
- [測試連線] button — sends a test prompt to verify key + model
- Settings auto-save to config.json

---

## Excel Structure

### One file per company: `output/AAPL.xlsx`

### Python-owned sheets (Data_* prefix)

| Sheet | Content | Source |
|-------|---------|--------|
| `Data_IS` | GAAP Income Statement | XBRL |
| `Data_BS` | Balance Sheet | XBRL |
| `Data_CF` | Cash Flow Statement | XBRL |
| `Data_Equity` | Statement of Stockholders' Equity | XBRL |
| `Data_CI` | Comprehensive Income Statement | XBRL |
| `Data_NonGAAP` | Non-GAAP metrics from earnings release | Gemini + 8-K |
| `Data_Meta` | Filing dates, fiscal year, company info | EDGAR |

- Sheets are created only if the company actually has that statement
- Python **only ever modifies Data_* sheets** — all other sheets are untouched

### Column Layout (all Data_* sheets)

```
Col A  : Concept / Label (row identifier, e.g. "Revenue", "us-gaap_Revenues")
Col B+ : One column per quarter, ordered oldest → newest (left → right)
Row 1  : Fiscal quarter label (e.g. FY2020Q1, FY2020Q2 ...)
Row 2  : Filing/announcement date (e.g. 2020-02-01)
Row 3+ : Financial line items
```

Python uses all columns from A onwards. Users do NOT annotate inside Data_* sheets.
Instead, users create their own sheets (e.g. My_IS, My_BS) for analysis and annotations,
and reference Data_* sheets via XLOOKUP by label name.

### Update Strategy
- Full rewrite of all Data_* sheets on every update
- Ensures restatements and corrections from SEC are captured
- User's own sheets (non-Data_*) are never touched by Python
- User should reference data via XLOOKUP by label, not by fixed cell position

---

## Data Sources

### Phase 1 — GAAP (edgartools + SEC XBRL)

```python
from edgar import Company, set_identity
set_identity("...")
company = Company("AAPL")
financials = company.get_financials()

financials.income_statement()        # Data_IS
financials.balance_sheet()           # Data_BS
financials.cashflow_statement()      # Data_CF
financials.statement_of_equity()     # Data_Equity
financials.comprehensive_income()    # Data_CI
```

- Fetches all historical quarterly filings
- Each filing includes fiscal year, period dates, and announcement date
- `include_dimensions=True` to capture segment breakdowns where available

### Phase 2 — Non-GAAP (Gemini API + 8-K Exhibit 99.1)

1. Use edgartools to fetch all 8-K filings for the company
2. Find Exhibit 99.1 (earnings press release) from each filing
3. Extract HTML content of the exhibit
4. Send to Gemini API with a structured prompt requesting Non-GAAP metrics as JSON
5. Parse response and write to `Data_NonGAAP` sheet

- Default model: `gemini-flash-latest` (Google's canonical alias — version-agnostic)
- Provider and model are user-configurable in Advanced Settings
- If extraction fails for a quarter, that cell is left blank with an error note in Data_Meta

---

## Config Files

### config.json (gitignored)
```json
{
  "identity": "Your Name your@email.com",
  "output_dir": "output",
  "watchlist": [
    {"ticker": "AAPL", "name": "Apple Inc."},
    {"ticker": "MSFT", "name": "Microsoft Corporation"}
  ],
  "ai": {
    "provider": "google",
    "model": "gemini-flash-latest",
    "api_key": ""
  }
}
```

### config.example.json (committed)
Same structure as above with empty `api_key` and example watchlist entries.

### company_cache.json (committed)
```json
{
  "last_updated": "2026-04-13",
  "companies": {
    "AAPL": "Apple Inc.",
    "MSFT": "Microsoft Corporation"
  }
}
```

---

## Launcher (Windows Tool Rules)

- `啟動器.bat`: 2 lines only, English only (CP950 safe)
- `launcher.ps1`: all Chinese UI, UTF-8 BOM required, checks Python → uv → venv
- First run: copies `config.example.json` → `config.json`, prompts user to fill API key
- venv managed by `uv`, packages installed from `requirements.txt`
- Broken dist-info cleanup before each `uv pip install` (per pitfalls.md 地雷九)

---

## Key Requirements

| # | Requirement |
|---|------------|
| 1 | Double-click BAT → tool launches with zero technical knowledge |
| 2 | All historical quarterly data fetched, oldest→newest left→right |
| 3 | Every column includes fiscal quarter label + announcement date |
| 4 | Python never modifies non-Data_* sheets; users work in their own My_* sheets |
| 5 | company_cache.json used for instant name lookup, live EDGAR as fallback |
| 6 | Non-GAAP module is pluggable — Phase 1 works without it |
| 7 | config.json, output/ and venv/ never committed to git |
| 8 | AI provider and model are user-configurable in Advanced Settings |
| 9 | Model default is `gemini-flash-latest` — do not rename or alias |
| 10 | Cross-company comparison is a future extension — design must not prevent it |

---

## Out of Scope (This Version)

- Cross-company comparison export
- Automatic scheduled updates
- Non-US company filings (non-SEC)
- PDF-based press releases (Non-GAAP extraction limited to HTML exhibits)
