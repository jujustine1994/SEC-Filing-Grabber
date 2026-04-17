# Excel Formatting & Index Sheet Design Spec

**Date:** 2026-04-17  
**Status:** Approved

## Goal

Improve the Excel output from unreadable (scientific notation, no structure) to professional financial report quality. Python handles all formatting automatically — no manual steps required.

---

## New File: `excel_formatter.py`

Single public function:

```python
def format_workbook(wb: Workbook, ticker: str, tables: list[StatementTable]) -> None
```

Called at the end of `excel_writer.write_statements()`, after all data has been written. Applies formatting and inserts the Index sheet. Does NOT modify cell values.

---

## Number Format Rules

Applied to data rows in all `Data_Financials(Q)`, `Data_Financials(Y)`, `Data_Seg_*` sheets.

| Row type | Format string | Notes |
|----------|--------------|-------|
| Default financial row | `#,##0.0_ ;[Red](#,##0.0)` | Millions USD, 1 decimal, negative = red parentheses |
| EPS rows (Basic EPS, Diluted EPS) | `#,##0.00_ ;[Red](#,##0.00)` | 2 decimals, no unit conversion |
| Share rows (Basic Shares, Diluted Shares) | `#,##0` | Raw value (already in millions for most companies) |
| Section header rows | *(no number format)* | All-None values, not applicable |
| Blank separator rows | *(no number format)* | All-None values |

**Unit conversion:** All financial values divided by 1,000,000 before writing. EPS and share rows are excluded from conversion (detected by row concept name matching).

**Rows excluded from ÷1M conversion:**
- Concept contains any of: `"EPS"`, `"Shares"`, `"Per Share"`, `"per share"`

---

## Column Widths

Applied to all `Data_*` sheets.

| Column | Width (chars) | Content |
|--------|--------------|---------|
| A | 22 | Std Name (concept label) |
| B | 24 | Original Item (XBRL label) |
| C onward | 13 | Quarterly/annual data values |

---

## Row Styles

### Row 1 — Ticker / Quarter Labels
- Background: `#1F3864` (dark navy)
- Font: white, bold, size 11
- Row height: 18pt

### Row 2 — Filing Dates
- Background: `#2D4A82` (medium navy)
- Font: `#AABBCC`, size 9
- Row height: 14pt

### Section Header Rows
Detected by: concept value is one of `"Income Statement"`, `"Balance Sheet"`, `"Cash Flow"`.
- Background: `#2E75B6` (medium blue)
- Font: white, bold, size 10, letter-spacing via spaces
- Row height: 16pt

### Blank Separator Rows
Detected by: concept value is `""` (empty string).
- Background: `#EEEEEE`
- Row height: 6pt

### Subtotal Rows (bold)
Concepts that get bold formatting:
- Income Statement: `Gross Profit`, `Total Operating Expense`, `Operating Income`, `Pre-tax Income`, `Net Income`
- Balance Sheet: `Total Current Assets`, `Total Assets`, `Total Current Liabilities`, `Total Liabilities`, `Total Equity`
- Cash Flow: `Operating Cash Flow`, `Free Cash Flow`

### Alternating Data Rows
- Odd rows: white `#FFFFFF`
- Even rows: `#F5F8FF` (very light blue)

---

## Freeze Panes

All `Data_*` sheets: freeze at cell `C3`.
- Rows 1–2 (ticker/dates) fixed when scrolling vertically
- Columns A–B (concept names) fixed when scrolling horizontally

---

## Index Sheet

**Sheet name:** `Index`  
**Position:** First sheet (index 0) in workbook.  
**Not prefixed with `Data_`** — never deleted by `excel_writer`.

### Layout

```
Row 1:  [TICKER — Company Name]          (navy header, full width merge, large bold)
Row 2:  [抓取日期: YYYY-MM-DD   資料來源: SEC EDGAR]   (medium navy, small font)
Row 3:  (blank)
Row 4:  [Sheet] [說明] [最早期間] [最新期間]   (column headers, light blue bg)
Row 5+: one row per Data_* sheet
```

### Column widths
- A: 22 (Sheet name)
- B: 28 (說明)
- C: 12 (最早期間)
- D: 12 (最新期間)

### Sheet descriptions (hardcoded mapping)

| Sheet name prefix | 說明 |
|------------------|------|
| `Data_Financials(Q)` | 季報三表合一（IS + BS + CF，from 10-Q） |
| `Data_Financials(Y)` | 年報三表合一（IS + BS + CF，from 10-K） |
| `Data_Seg_*` | Segment 細項：`{suffix}` |
| `Data_EPS_Recon` | Non-GAAP EPS 調節表（from 8-K） |
| `Data_NonGAAP` | Non-GAAP 指標（AI 提取，from 8-K press release） |
| `Data_Meta` | 申報資訊（Ticker、公司名、抓取日期） |

### Earliest / latest period
- Read from `StatementTable.quarter_labels[0]` and `[-1]`
- `Data_Meta`, `Data_EPS_Recon`, `Data_NonGAAP`: show `—`

### Index row styling
- Sheet name column: navy blue font `#1F3864`, bold for `Data_Financials(Q/Y)`; grey `#666666` for Seg/Meta
- Alternating row background (same as data sheets)
- No number formats needed

---

## Integration

### `excel_writer.py` change
At the end of `write_statements()`, after `wb.save()` is removed, replace with:

```python
from excel_formatter import format_workbook
format_workbook(wb, ticker, tables)
wb.save(output_path)
wb.close()
```

`tables` is already available in `write_statements()` — pass it through so `format_workbook` can build the Index sheet.

### `write_statements()` signature change
Add `ticker: str = ""` parameter (already stored in `StatementTable.ticker` — use that as fallback).

### `main.py`
No changes required.

---

## Files Changed

| File | Change |
|------|--------|
| `excel_formatter.py` | **New** — all formatting logic |
| `excel_writer.py` | Call `format_workbook()` at end of `write_statements()` |
| `tests/test_excel_formatter.py` | **New** — unit tests for formatter |
| `tests/test_excel_writer.py` | Minor: update to expect formatted output |

---

## Exclusions (out of scope)

- No template.xlsx mechanism
- No user-customizable colors
- `My_*` sheets: formatting not touched (already guaranteed by existing `Data_`-prefix filter)
- `Data_Meta` sheet: apply column widths and freeze panes but skip number formatting (all values are strings)
