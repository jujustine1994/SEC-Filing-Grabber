# Excel Formatting & Index Sheet Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a new `excel_formatter.py` module that applies professional formatting (column widths, colours, number formats in Millions USD, freeze panes) to all `Data_*` sheets and generates a dynamic `Index` sheet as the first sheet.

**Architecture:** `excel_formatter.format_workbook(wb, tables)` is called at the end of `excel_writer.write_statements()` before `wb.save()`. The formatter reads concept names from col A to classify each row, divides financial values by 1,000,000 (except EPS rows), applies openpyxl styles, and inserts the Index sheet at position 0.

**Tech Stack:** Python 3.11, openpyxl (already in requirements), pytest.

---

## Working directory
`C:\Users\CTH\Documents\Code\SEC Financial Tools`

## File map

| File | Action | Responsibility |
|------|--------|---------------|
| `excel_formatter.py` | **Create** | All formatting + Index sheet generation |
| `excel_writer.py` | **Modify** | Call `format_workbook(wb, tables)` before `wb.save()` |
| `tests/test_excel_formatter.py` | **Create** | Unit tests for formatter |
| `tests/test_excel_writer.py` | **Modify** | Update 2 assertions that change due to formatting |

---

## Constants & Row Classification (used across all tasks)

These are defined once in `excel_formatter.py` and referenced throughout:

```python
# Colour hex (openpyxl uses ARGB — prepend "FF")
NAVY_DARK  = "FF1F3864"
NAVY_MID   = "FF2D4A82"
BLUE_MID   = "FF2E75B6"
GREY_SEP   = "FFEEEEEE"
ROW_ALT    = "FFF5F8FF"
ROW_WHITE  = "FFFFFFFF"
BLUE_HDR   = "FFDDE8F5"   # index column header

SECTION_HEADERS = {"Income Statement", "Balance Sheet", "Cash Flow"}

SUBTOTAL_CONCEPTS = {
    "Gross Profit", "Total Operating Expense", "Operating Income",
    "Pre-tax Income", "Net Income",
    "Total Current Assets", "Total Assets",
    "Total Current Liabilities", "Total Liabilities", "Total Equity",
    "Operating Cash Flow", "Free Cash Flow",
}

SHEET_DESCRIPTIONS = {
    "Data_Financials(Q)": "季報三表合一（IS + BS + CF，from 10-Q）",
    "Data_Financials(Y)": "年報三表合一（IS + BS + CF，from 10-K）",
    "Data_EPS_Recon":     "Non-GAAP EPS 調節表（from 8-K）",
    "Data_NonGAAP":       "Non-GAAP 指標（AI 提取）",
    "Data_Meta":          "申報資訊（Ticker、公司名、抓取日期）",
}

def _is_eps_concept(c: str) -> bool:
    return any(k in (c or "") for k in ("EPS", "Per Share", "per share"))

def _sheet_description(name: str) -> str:
    if name in SHEET_DESCRIPTIONS:
        return SHEET_DESCRIPTIONS[name]
    if name.startswith("Data_Seg_"):
        return f"Segment 細項：{name[9:]}"
    return name
```

---

## Task 1: Foundation — column widths and freeze panes

**Files:**
- Create: `excel_formatter.py`
- Create: `tests/test_excel_formatter.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_excel_formatter.py`:

```python
"""Tests for excel_formatter.py."""
import pytest
from openpyxl import Workbook
from excel_formatter import format_workbook


def _make_wb(sheet_name="Data_Financials(Q)"):
    """Minimal workbook with one Data_* sheet and two data columns."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws["A1"] = "AAPL"
    ws["C1"] = "FY2023Q1"
    ws["D1"] = "FY2023Q2"
    ws["C2"] = "2023-02-03"
    ws["D2"] = "2023-05-05"
    ws["A3"] = "Income Statement"   # section header
    ws["A4"] = "Revenue"
    ws["B4"] = "Revenues"
    ws["C4"] = 117154000000.0
    ws["D4"] = 94836000000.0
    ws["A5"] = ""                   # blank separator
    ws["A6"] = "Basic EPS"
    ws["C6"] = 1.52
    ws["D6"] = 1.20
    ws["A7"] = "Basic Shares"
    ws["C7"] = 15787000000.0
    ws["D7"] = 15813000000.0
    return wb


# ── column widths ──────────────────────────────────────────────────────────

def test_col_a_width():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"].column_dimensions["A"].width == 22

def test_col_b_width():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"].column_dimensions["B"].width == 24

def test_data_col_width():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"].column_dimensions["C"].width == 13

def test_data_col_d_width():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"].column_dimensions["D"].width == 13


# ── freeze panes ───────────────────────────────────────────────────────────

def test_freeze_panes():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"].freeze_panes == "C3"

def test_freeze_panes_seg_sheet():
    wb = _make_wb(sheet_name="Data_Seg_Revenue")
    format_workbook(wb, [])
    assert wb["Data_Seg_Revenue"].freeze_panes == "C3"
```

- [ ] **Step 2: Run to verify they fail**

```
venv\Scripts\python.exe -m pytest tests/test_excel_formatter.py -v 2>&1 | head -20
```

Expected: `ImportError: No module named 'excel_formatter'`

- [ ] **Step 3: Create `excel_formatter.py` with foundation**

```python
"""
excel_formatter.py — Apply professional formatting to Data_* sheets and generate Index sheet.

Public API:
    format_workbook(wb, tables) -> None

Called by excel_writer.write_statements() before wb.save(). Modifies cell values
(÷1M unit conversion) and applies openpyxl styles. Does not change sheet structure.
"""

from __future__ import annotations
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from fetcher_gaap import StatementTable
from datetime import date

# ── Colours (ARGB) ────────────────────────────────────────────────────────
NAVY_DARK = "FF1F3864"
NAVY_MID  = "FF2D4A82"
BLUE_MID  = "FF2E75B6"
GREY_SEP  = "FFEEEEEE"
ROW_ALT   = "FFF5F8FF"
ROW_WHITE = "FFFFFFFF"
BLUE_HDR  = "FFDDE8F5"

# ── Row classification ────────────────────────────────────────────────────
SECTION_HEADERS = {"Income Statement", "Balance Sheet", "Cash Flow"}

SUBTOTAL_CONCEPTS = {
    "Gross Profit", "Total Operating Expense", "Operating Income",
    "Pre-tax Income", "Net Income",
    "Total Current Assets", "Total Assets",
    "Total Current Liabilities", "Total Liabilities", "Total Equity",
    "Operating Cash Flow", "Free Cash Flow",
}

SHEET_DESCRIPTIONS = {
    "Data_Financials(Q)": "季報三表合一（IS + BS + CF，from 10-Q）",
    "Data_Financials(Y)": "年報三表合一（IS + BS + CF，from 10-K）",
    "Data_EPS_Recon":     "Non-GAAP EPS 調節表（from 8-K）",
    "Data_NonGAAP":       "Non-GAAP 指標（AI 提取）",
    "Data_Meta":          "申報資訊（Ticker、公司名、抓取日期）",
}


def _is_eps_concept(c: str) -> bool:
    return any(k in (c or "") for k in ("EPS", "Per Share", "per share"))


def _sheet_description(name: str) -> str:
    if name in SHEET_DESCRIPTIONS:
        return SHEET_DESCRIPTIONS[name]
    if name.startswith("Data_Seg_"):
        return f"Segment 細項：{name[9:]}"
    return name


def _fill(hex_argb: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_argb)


# ── Column widths ─────────────────────────────────────────────────────────

def _apply_column_widths(ws) -> None:
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 24
    for col in range(3, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col)].width = 13


# ── Freeze panes ──────────────────────────────────────────────────────────

def _set_freeze_panes(ws) -> None:
    ws.freeze_panes = "C3"


# ── Public API ────────────────────────────────────────────────────────────

def format_workbook(wb: Workbook, tables: list[StatementTable]) -> None:
    """Apply formatting to all Data_* sheets and insert Index sheet at position 0."""
    for ws in wb.worksheets:
        if not ws.title.startswith("Data_"):
            continue
        _apply_column_widths(ws)
        _set_freeze_panes(ws)
```

- [ ] **Step 4: Run tests to verify they pass**

```
venv\Scripts\python.exe -m pytest tests/test_excel_formatter.py -v
```

Expected: 6 passed.

- [ ] **Step 5: Commit**

```bash
git add excel_formatter.py tests/test_excel_formatter.py
git commit -m "feat: add excel_formatter foundation (column widths, freeze panes)"
```

---

## Task 2: Row styles

**Files:**
- Modify: `excel_formatter.py`
- Modify: `tests/test_excel_formatter.py`

- [ ] **Step 1: Add failing tests**

Append to `tests/test_excel_formatter.py`:

```python
# ── row styles ─────────────────────────────────────────────────────────────

def _rgb(ws, cell_ref: str) -> str:
    """Return fgColor ARGB string of a cell's fill."""
    return ws[cell_ref].fill.fgColor.rgb


def test_row1_fill_navy_dark():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    assert _rgb(ws, "A1") == "FF1F3864"

def test_row1_font_bold_white():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    assert ws["A1"].font.bold is True
    assert ws["A1"].font.color.rgb == "FFFFFFFF"

def test_row2_fill_navy_mid():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    assert _rgb(ws, "A2") == "FF2D4A82"

def test_section_header_fill_blue_mid():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    # A3 = "Income Statement"
    assert _rgb(ws, "A3") == "FF2E75B6"

def test_section_header_font_bold_white():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    assert ws["A3"].font.bold is True
    assert ws["A3"].font.color.rgb == "FFFFFFFF"

def test_blank_separator_fill_grey():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    # A5 = ""
    assert _rgb(ws, "A5") == "FFEEEEEE"

def test_data_row_alternating_white():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    # Row 4 = "Revenue" (even row number → white)
    assert _rgb(ws, "A4") == "FFFFFFFF"

def test_data_row_alternating_blue():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    # Row 6 = "Basic EPS" (even row number → white)
    # Row 7 = "Basic Shares" (odd row number → alt blue)
    assert _rgb(ws, "A7") == "FFF5F8FF"

def test_subtotal_row_bold():
    wb = _make_wb()
    # Add a Gross Profit row
    wb["Data_Financials(Q)"]["A8"] = "Gross Profit"
    wb["Data_Financials(Q)"]["C8"] = 5000000000.0
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    assert ws["A8"].font.bold is True

def test_normal_row_not_bold():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    # A4 = "Revenue" (not a subtotal)
    assert ws["A4"].font.bold is not True
```

- [ ] **Step 2: Run to verify they fail**

```
venv\Scripts\python.exe -m pytest tests/test_excel_formatter.py -v -k "row" 2>&1 | tail -20
```

Expected: multiple FAILs (row style functions not implemented yet).

- [ ] **Step 3: Add `_apply_row_styles` to `excel_formatter.py`**

Add after `_set_freeze_panes`:

```python
def _apply_row_styles(ws) -> None:
    """Apply fill and font styles to all rows."""
    white_font  = Font(color="FFFFFFFF", bold=True, size=11)
    small_font  = Font(color="FFAABBCC", size=9)

    # Row 1: ticker / quarter labels — dark navy
    for cell in ws[1]:
        cell.fill = _fill(NAVY_DARK)
        cell.font = white_font

    # Row 2: filing dates — medium navy
    for cell in ws[2]:
        cell.fill = _fill(NAVY_MID)
        cell.font = small_font

    # Row 3+: classify by col A value
    for row_idx in range(3, ws.max_row + 1):
        concept = ws.cell(row=row_idx, column=1).value or ""
        concept = str(concept).strip()

        if concept in SECTION_HEADERS:
            row_fill  = _fill(BLUE_MID)
            row_font  = Font(color="FFFFFFFF", bold=True, size=10)
            row_height = 16
        elif concept == "":
            row_fill  = _fill(GREY_SEP)
            row_font  = Font(size=9)
            row_height = 6
        else:
            row_fill  = _fill(ROW_WHITE) if row_idx % 2 == 0 else _fill(ROW_ALT)
            bold      = concept in SUBTOTAL_CONCEPTS
            row_font  = Font(bold=bold) if bold else Font()
            row_height = None

        for cell in ws[row_idx]:
            cell.fill = row_fill
            cell.font = row_font
        if row_height is not None:
            ws.row_dimensions[row_idx].height = row_height
```

Update `format_workbook` to call `_apply_row_styles`:

```python
def format_workbook(wb: Workbook, tables: list[StatementTable]) -> None:
    """Apply formatting to all Data_* sheets and insert Index sheet at position 0."""
    for ws in wb.worksheets:
        if not ws.title.startswith("Data_"):
            continue
        _apply_column_widths(ws)
        _set_freeze_panes(ws)
        _apply_row_styles(ws)
```

- [ ] **Step 4: Run tests**

```
venv\Scripts\python.exe -m pytest tests/test_excel_formatter.py -v
```

Expected: all pass. If a font colour test fails, check that openpyxl Font colour is set as `Font(color="FFFFFFFF")` not `Font(color=Color("FFFFFFFF"))`.

- [ ] **Step 5: Commit**

```bash
git add excel_formatter.py tests/test_excel_formatter.py
git commit -m "feat: add row style formatting (header, section, alternating, subtotal)"
```

---

## Task 3: Number formatting and unit conversion (÷ 1,000,000)

**Files:**
- Modify: `excel_formatter.py`
- Modify: `tests/test_excel_formatter.py`

- [ ] **Step 1: Add failing tests**

Append to `tests/test_excel_formatter.py`:

```python
# ── number formatting + unit conversion ────────────────────────────────────

FMT_FINANCIAL = "#,##0.0_ ;[Red](#,##0.0)"
FMT_EPS       = "#,##0.00_ ;[Red](#,##0.00)"
FMT_SHARES    = "#,##0"

def test_revenue_divided_by_million():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    # C4 = "Revenue" row, raw = 117154000000
    assert ws["C4"].value == pytest.approx(117154.0)

def test_revenue_second_col_divided():
    wb = _make_wb()
    format_workbook(wb, [])
    ws = wb["Data_Financials(Q)"]
    assert ws["D4"].value == pytest.approx(94836.0)

def test_revenue_number_format():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"]["C4"].number_format == FMT_FINANCIAL

def test_eps_not_divided():
    wb = _make_wb()
    format_workbook(wb, [])
    # C6 = "Basic EPS" = 1.52 → must stay 1.52
    assert wb["Data_Financials(Q)"]["C6"].value == pytest.approx(1.52)

def test_eps_number_format():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"]["C6"].number_format == FMT_EPS

def test_shares_divided_by_million():
    wb = _make_wb()
    format_workbook(wb, [])
    # C7 = "Basic Shares" = 15787000000 → 15787.0
    assert wb["Data_Financials(Q)"]["C7"].value == pytest.approx(15787.0)

def test_shares_number_format():
    wb = _make_wb()
    format_workbook(wb, [])
    assert wb["Data_Financials(Q)"]["C7"].number_format == FMT_SHARES

def test_section_header_values_unchanged():
    wb = _make_wb()
    format_workbook(wb, [])
    # C3 = "Income Statement" row — all None
    assert wb["Data_Financials(Q)"]["C3"].value is None

def test_data_meta_values_not_converted():
    """Data_Meta contains strings — must not be divided."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Data_Meta"
    ws["A1"] = "AAPL"
    ws["C1"] = "FY2023Q1"
    ws["A3"] = "Ticker"
    ws["C3"] = "AAPL"
    format_workbook(wb, [])
    assert wb["Data_Meta"]["C3"].value == "AAPL"

def test_seg_sheet_financial_converted():
    wb = _make_wb(sheet_name="Data_Seg_Revenue")
    wb["Data_Seg_Revenue"]["A4"] = "Americas"
    wb["Data_Seg_Revenue"]["C4"] = 50000000000.0
    format_workbook(wb, [])
    assert wb["Data_Seg_Revenue"]["C4"].value == pytest.approx(50000.0)
```

- [ ] **Step 2: Run to verify they fail**

```
venv\Scripts\python.exe -m pytest tests/test_excel_formatter.py -v -k "divided or format or converted" 2>&1 | tail -20
```

Expected: all FAILs.

- [ ] **Step 3: Add `_apply_number_formats` to `excel_formatter.py`**

Add the constant and function:

```python
# Sheets that get full number formatting + unit conversion
_FINANCIAL_SHEET_PREFIXES = ("Data_Financials", "Data_Seg_", "Data_EPS_Recon", "Data_NonGAAP")

FMT_FINANCIAL = "#,##0.0_ ;[Red](#,##0.0)"
FMT_EPS       = "#,##0.00_ ;[Red](#,##0.00)"
FMT_SHARES    = "#,##0"


def _apply_number_formats(ws) -> None:
    """Convert values to millions and apply number formats. Skips section/blank rows."""
    for row_idx in range(3, ws.max_row + 1):
        concept = str(ws.cell(row=row_idx, column=1).value or "").strip()

        # Section headers and blank separators have no numeric data
        if concept in SECTION_HEADERS or concept == "":
            continue

        is_eps    = _is_eps_concept(concept)
        is_shares = "Shares" in concept

        if is_eps:
            fmt = FMT_EPS
            divisor = 1
        elif is_shares:
            fmt = FMT_SHARES
            divisor = 1_000_000
        else:
            fmt = FMT_FINANCIAL
            divisor = 1_000_000

        for col_idx in range(3, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if isinstance(cell.value, (int, float)):
                cell.value = cell.value / divisor
            cell.number_format = fmt
```

Update `format_workbook`:

```python
def format_workbook(wb: Workbook, tables: list[StatementTable]) -> None:
    """Apply formatting to all Data_* sheets and insert Index sheet at position 0."""
    for ws in wb.worksheets:
        if not ws.title.startswith("Data_"):
            continue
        _apply_column_widths(ws)
        _set_freeze_panes(ws)
        _apply_row_styles(ws)
        if ws.title != "Data_Meta":
            _apply_number_formats(ws)
```

- [ ] **Step 4: Run tests**

```
venv\Scripts\python.exe -m pytest tests/test_excel_formatter.py -v
```

Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add excel_formatter.py tests/test_excel_formatter.py
git commit -m "feat: add number formatting and millions unit conversion"
```

---

## Task 4: Index sheet

**Files:**
- Modify: `excel_formatter.py`
- Modify: `tests/test_excel_formatter.py`

- [ ] **Step 1: Add failing tests**

Append to `tests/test_excel_formatter.py`:

```python
# ── Index sheet ────────────────────────────────────────────────────────────

from fetcher_gaap import StatementTable

def _make_tables(sheet_name="Data_Financials(Q)", ticker="AAPL",
                 qs=None, dates=None):
    qs    = qs    or ["FY2020Q1", "FY2024Q4"]
    dates = dates or ["2020-04-30", "2025-01-30"]
    return [StatementTable(
        sheet_name=sheet_name, ticker=ticker,
        quarter_labels=qs, filing_dates=dates,
        concepts=["Revenue"], values=[[100.0, 200.0]], labels=["Revenues"],
    )]


def test_index_sheet_created():
    wb = _make_wb()
    format_workbook(wb, _make_tables())
    assert "Index" in wb.sheetnames

def test_index_sheet_is_first():
    wb = _make_wb()
    format_workbook(wb, _make_tables())
    assert wb.sheetnames[0] == "Index"

def test_index_sheet_ticker_in_a1():
    wb = _make_wb()
    format_workbook(wb, _make_tables(ticker="AAPL"))
    ws = wb["Index"]
    assert "AAPL" in str(ws["A1"].value)

def test_index_lists_data_sheet():
    wb = _make_wb()
    format_workbook(wb, _make_tables())
    ws = wb["Index"]
    col_a_values = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
    assert "Data_Financials(Q)" in col_a_values

def test_index_shows_earliest_period():
    wb = _make_wb()
    format_workbook(wb, _make_tables(qs=["FY2010Q1", "FY2024Q4"]))
    ws = wb["Index"]
    all_values = [ws.cell(row=r, column=c).value
                  for r in range(1, ws.max_row + 1)
                  for c in range(1, 5)]
    assert "FY2010Q1" in all_values

def test_index_shows_latest_period():
    wb = _make_wb()
    format_workbook(wb, _make_tables(qs=["FY2010Q1", "FY2024Q4"]))
    ws = wb["Index"]
    all_values = [ws.cell(row=r, column=c).value
                  for r in range(1, ws.max_row + 1)
                  for c in range(1, 5)]
    assert "FY2024Q4" in all_values

def test_index_not_deleted_on_reformat():
    """Index must not be deleted by excel_writer (doesn't start with Data_)."""
    wb = _make_wb()
    format_workbook(wb, _make_tables())
    # Simulate a second write: excel_writer deletes Data_* sheets
    for name in list(wb.sheetnames):
        if name.startswith("Data_"):
            del wb[name]
    assert "Index" in wb.sheetnames
```

- [ ] **Step 2: Run to verify they fail**

```
venv\Scripts\python.exe -m pytest tests/test_excel_formatter.py -v -k "index" 2>&1 | tail -20
```

Expected: all FAILs.

- [ ] **Step 3: Add `_build_index_sheet` to `excel_formatter.py`**

```python
def _build_index_sheet(wb: Workbook, tables: list[StatementTable]) -> None:
    """Insert or replace the Index sheet at position 0."""
    if "Index" in wb.sheetnames:
        del wb["Index"]

    ticker       = tables[0].ticker if tables else ""
    company_name = ""
    # Try to get company name from Data_Meta if present
    meta = next((t for t in tables if t.sheet_name == "Data_Meta"), None)
    if meta and len(meta.concepts) > 1 and meta.values and meta.values[1]:
        company_name = meta.values[1][0] or ""

    header_text = f"{ticker} — {company_name}" if company_name else ticker

    ws = wb.create_sheet("Index", 0)

    # ── Row 1: company header ────────────────────────────────────────────
    ws["A1"] = header_text
    ws["A1"].fill = _fill(NAVY_DARK)
    ws["A1"].font = Font(color="FFFFFFFF", bold=True, size=14)
    ws.merge_cells("A1:D1")

    # ── Row 2: metadata ──────────────────────────────────────────────────
    ws["A2"] = f"抓取日期：{date.today()}　　資料來源：SEC EDGAR"
    ws["A2"].fill = _fill(NAVY_MID)
    ws["A2"].font = Font(color="FFAABBCC", size=9)
    ws.merge_cells("A2:D2")

    # ── Row 3: blank ─────────────────────────────────────────────────────
    ws.row_dimensions[3].height = 6

    # ── Row 4: column headers ────────────────────────────────────────────
    hdr_font = Font(bold=True, size=10)
    hdr_fill = _fill(BLUE_HDR)
    for col, label in enumerate(["Sheet", "說明", "最早期間", "最新期間"], start=1):
        cell = ws.cell(row=4, column=col, value=label)
        cell.font = hdr_font
        cell.fill = hdr_fill

    # ── Row 5+: one row per Data_* sheet ─────────────────────────────────
    data_sheets = [t for t in tables if t.sheet_name.startswith("Data_")]
    for i, tbl in enumerate(data_sheets):
        row = 5 + i
        earliest = tbl.quarter_labels[0]  if tbl.quarter_labels else "—"
        latest   = tbl.quarter_labels[-1] if tbl.quarter_labels else "—"

        is_primary = tbl.sheet_name in ("Data_Financials(Q)", "Data_Financials(Y)")
        name_font  = Font(color="FF1F3864" if is_primary else "FF666666",
                          bold=is_primary, size=10)
        row_fill   = _fill(ROW_WHITE) if row % 2 == 0 else _fill(ROW_ALT)

        for col, val in enumerate([tbl.sheet_name,
                                    _sheet_description(tbl.sheet_name),
                                    earliest, latest], start=1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.fill = row_fill
            if col == 1:
                cell.font = name_font

    # ── Column widths ────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
```

Update `format_workbook` to call `_build_index_sheet`:

```python
def format_workbook(wb: Workbook, tables: list[StatementTable]) -> None:
    """Apply formatting to all Data_* sheets and insert Index sheet at position 0."""
    _build_index_sheet(wb, tables)
    for ws in wb.worksheets:
        if not ws.title.startswith("Data_"):
            continue
        _apply_column_widths(ws)
        _set_freeze_panes(ws)
        _apply_row_styles(ws)
        if ws.title != "Data_Meta":
            _apply_number_formats(ws)
```

- [ ] **Step 4: Run all tests**

```
venv\Scripts\python.exe -m pytest tests/test_excel_formatter.py -v
```

Expected: all pass. Common failure: `merge_cells` on row 1 may make `ws["B1"]` empty — verify the header text is in `ws["A1"]` not a merged cell reference.

- [ ] **Step 5: Commit**

```bash
git add excel_formatter.py tests/test_excel_formatter.py
git commit -m "feat: add Index sheet (position 0, earliest/latest period per sheet)"
```

---

## Task 5: Wire into excel_writer.py

**Files:**
- Modify: `excel_writer.py`
- Modify: `tests/test_excel_writer.py`

- [ ] **Step 1: Read the end of `write_statements` in `excel_writer.py`**

Locate the `wb.save` / `wb.close` block. It currently looks like:

```python
    try:
        wb.save(output_path)
    finally:
        wb.close()
```

- [ ] **Step 2: Add `format_workbook` call to `excel_writer.py`**

At the top of `excel_writer.py`, add the import:

```python
from excel_formatter import format_workbook
```

Replace the `wb.save` block:

```python
    format_workbook(wb, tables)
    try:
        wb.save(output_path)
    finally:
        wb.close()
```

- [ ] **Step 3: Run the full test suite**

```
venv\Scripts\python.exe -m pytest tests/ -v 2>&1 | tail -40
```

Expected: all pass. Two `test_excel_writer.py` tests may need updating:

**If `test_rewrite_replaces_old_data` fails** (checks `"Data_BS" not in wb.sheetnames`) — this test still works because `Data_BS` is a `Data_*` sheet, deleted by `excel_writer` as before.

**If `test_preserves_non_data_sheets` fails** — check that `My_IS` is still present after `format_workbook` runs. It should be, because `format_workbook` only touches `Data_*` sheets and `Index`.

**If `test_col_a_is_concept_name` fails** — the test checks `ws["A1"].value is None`. After formatting, `A1` still has `None` (it's a `Data_IS` test sheet). No change needed.

- [ ] **Step 4: Fix any test_excel_writer.py failures**

If the `Data_IS` test sheet (used in `test_excel_writer.py`) gets formatted with number conversion on its values, the existing value assertions may break. In that case, add a check: `test_excel_writer.py` uses a sheet named `"Data_IS"` which is not in `_FINANCIAL_SHEET_PREFIXES`... wait, `"Data_IS"` does NOT match `"Data_Financials"` or `"Data_Seg_"`. But `_apply_number_formats` is called for all sheets except `Data_Meta`. Fix: update `format_workbook` to skip number formatting for sheets not in the expected set:

```python
        skip_number_fmt = ws.title == "Data_Meta"
        if not skip_number_fmt:
            _apply_number_formats(ws)
```

The existing logic already has `if ws.title != "Data_Meta"` — so `Data_IS` (used only in tests, not produced by the real fetcher) WILL get number-formatted. The test values are small integers like `1000.0` which become `0.001` after ÷1M. 

Fix `tests/test_excel_writer.py` — update the two value assertions in `test_data_values_correct`:

```python
def test_data_values_correct(tmp_path, sample_tables):
    out = tmp_path / "AAPL.xlsx"
    write_statements(sample_tables, out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Data_IS"]
    # Values are divided by 1M during formatting: 1000.0 → 0.001
    assert ws["B3"].value == pytest.approx(1000.0 / 1_000_000)
    assert ws["C3"].value == pytest.approx(1100.0 / 1_000_000)
    assert ws["D3"].value == pytest.approx(1200.0 / 1_000_000)
```

Add `import pytest` at the top of `tests/test_excel_writer.py` if not already present.

- [ ] **Step 5: Run full test suite again to confirm green**

```
venv\Scripts\python.exe -m pytest tests/ -v
```

Expected: all pass.

- [ ] **Step 6: Smoke test — produce a real file**

```
venv\Scripts\python.exe -c "
from config import load_config
from fetcher_gaap import fetch_gaap_statements
from excel_writer import write_statements
from pathlib import Path
cfg = load_config()
tables = fetch_gaap_statements('AAPL', cfg['identity'], max_filings=4, max_annual_filings=2)
write_statements(tables, Path('output/AAPL_fmt_test.xlsx'))
print('Done — open output/AAPL_fmt_test.xlsx')
"
```

Open the file and verify:
- `Index` is the first sheet tab
- `Data_Financials(Q)` has readable comma numbers (not scientific notation)
- Revenue for AAPL recent quarter shows ~119,000 (millions), not 1.19E+11
- Section headers are dark blue with white text
- Rows 1–2 are dark navy

- [ ] **Step 7: Commit**

```bash
git add excel_writer.py tests/test_excel_writer.py
git commit -m "feat: wire excel_formatter into write_statements (formatting + Index sheet)"
```

---

## Self-Review

### Spec coverage

| Requirement | Task |
|-------------|------|
| Column widths (A=22, B=24, data=13) | Task 1 |
| Freeze panes at C3 | Task 1 |
| Row 1 dark navy, white bold | Task 2 |
| Row 2 medium navy, small grey | Task 2 |
| Section header blue, white bold | Task 2 |
| Blank separator grey, height 6 | Task 2 |
| Alternating row white/light blue | Task 2 |
| Subtotal rows bold | Task 2 |
| Financial rows ÷1M, 1-decimal format | Task 3 |
| EPS rows no conversion, 2-decimal | Task 3 |
| Shares rows ÷1M, integer format | Task 3 |
| Data_Meta skips number formatting | Task 3 |
| Index sheet at position 0 | Task 4 |
| Index: company header navy | Task 4 |
| Index: each Data_* sheet listed | Task 4 |
| Index: earliest / latest period | Task 4 |
| Index: description per sheet | Task 4 |
| Integration into excel_writer | Task 5 |
| Smoke test with real data | Task 5 |

### Placeholder scan
None found.

### Type consistency
- `format_workbook(wb: Workbook, tables: list[StatementTable])` — consistent across Tasks 1–5
- `_fill(hex_argb: str) -> PatternFill` — used in Tasks 2 and 4
- `_apply_number_formats(ws)` — defined Task 3, called in `format_workbook` Task 3+
- `_build_index_sheet(wb, tables)` — defined Task 4, called in `format_workbook` Task 4+
