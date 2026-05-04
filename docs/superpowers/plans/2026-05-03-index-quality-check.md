# Index Sheet Quality Check Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 Excel Index sheet 加入報表完成度欄（E 欄）與品質明細區塊，讓分析師打開 Excel 第一頁就能看到哪些 key rows 有缺失。

**Architecture:** 全部改動集中在 `excel_formatter.py` 的 `_build_index_sheet()`。新增 `_compute_quality()` helper 呼叫既有的 `check_key_rows()`，回傳缺失清單；`_build_index_sheet()` 用結果寫入 E 欄完成度分數，並在表格下方加品質明細區塊。不動其他檔案。

**Tech Stack:** Python 3.11, openpyxl, pytest

---

## 檔案異動清單

| 檔案 | 動作 | 說明 |
|------|------|------|
| `excel_formatter.py` | 修改 | 新增顏色常數、`ALL_KEY_ROWS`、`_compute_quality()`；更新 `_build_index_sheet()` |
| `tests/test_excel_formatter.py` | 修改 | 新增品質欄與明細區塊的 unit tests |

---

## Task 1：`_compute_quality()` helper

**Files:**
- Modify: `excel_formatter.py`（在 `_build_index_sheet` 之前加入）
- Modify: `tests/test_excel_formatter.py`

- [ ] **Step 1: 寫失敗測試**

在 `tests/test_excel_formatter.py` 頂端 imports 加入：
```python
from fetcher_gaap import StatementTable
from excel_formatter import _compute_quality
```

在檔案尾端加入：

```python
# ── _compute_quality ──────────────────────────────────────────────────────

_ALL_KEY_ROWS = [
    "Revenue", "Operating Income", "Net Income", "Diluted EPS",
    "Total Assets", "Total Liabilities", "Total Equity — Parent",
    "Operating Cash Flow", "Capex",
]


def _make_q_table(missing=None):
    """StatementTable with all 9 key rows; rows in `missing` have all-None values."""
    missing = set(missing or [])
    values = [
        [None, None, None, None] if c in missing else [100.0, 200.0, 300.0, 400.0]
        for c in _ALL_KEY_ROWS
    ]
    return StatementTable(
        sheet_name="Data_Financials(Q)",
        quarter_labels=["FY2025Q1", "FY2025Q2", "FY2025Q3", "FY2025Q4"],
        filing_dates=["2025-02-01", "2025-05-01", "2025-08-01", "2025-11-01"],
        concepts=_ALL_KEY_ROWS[:],
        values=values,
        ticker="TEST",
    )


def test_compute_quality_all_ok():
    tbl = _make_q_table()
    score, total, missing = _compute_quality([tbl])
    assert total == 9
    assert score == 9
    assert missing == set()


def test_compute_quality_two_missing():
    tbl = _make_q_table(missing=["Operating Income", "Capex"])
    score, total, missing = _compute_quality([tbl])
    assert score == 7
    assert missing == {"Operating Income", "Capex"}


def test_compute_quality_no_q_table():
    result = _compute_quality([])
    assert result is None


def test_compute_quality_non_q_table_ignored():
    tbl = _make_q_table()
    tbl.sheet_name = "Data_Financials(Y)"
    result = _compute_quality([tbl])
    assert result is None
```

- [ ] **Step 2: 確認測試失敗**

```
python -m pytest tests/test_excel_formatter.py -k "compute_quality" -v
```
預期：`ImportError: cannot import name '_compute_quality'`

- [ ] **Step 3: 實作**

在 `excel_formatter.py` 的顏色常數區塊後加入新常數，並在 `_build_index_sheet` 之前加入 helper：

```python
# ── Quality check colours ─────────────────────────────────────────────────
QUALITY_GREEN   = "FF1A7A34"
QUALITY_ORANGE  = "FFC25C00"
QUALITY_MISS_BG = "FFFFF0E0"
QUALITY_MISS_FG = "FFC00000"

ALL_KEY_ROWS = [
    "Revenue", "Operating Income", "Net Income", "Diluted EPS",
    "Total Assets", "Total Liabilities", "Total Equity — Parent",
    "Operating Cash Flow", "Capex",
]


def _compute_quality(tables: list) -> tuple[int, int, set] | None:
    """Compute quality score for Data_Financials(Q).

    Returns (score, total, missing_set) or None if no Q table found.
    """
    from override_engine import check_key_rows
    q_tbl = next((t for t in tables if t.sheet_name == "Data_Financials(Q)"), None)
    if q_tbl is None:
        return None
    missing = set(
        check_key_rows(q_tbl.concepts, q_tbl.values, "IS") +
        check_key_rows(q_tbl.concepts, q_tbl.values, "BS") +
        check_key_rows(q_tbl.concepts, q_tbl.values, "CF")
    )
    total = len(ALL_KEY_ROWS)
    return total - len(missing), total, missing
```

- [ ] **Step 4: 確認測試通過**

```
python -m pytest tests/test_excel_formatter.py -k "compute_quality" -v
```
預期：4 PASSED

- [ ] **Step 5: 確認全套不破壞**

```
python -m pytest tests/test_excel_formatter.py -v
```
預期：全部 PASSED

- [ ] **Step 6: Commit**

```
git add excel_formatter.py tests/test_excel_formatter.py
git commit -m "feat: add _compute_quality helper for Index sheet quality check"
```

---

## Task 2：Index 表格加「完成度」E 欄

**Files:**
- Modify: `excel_formatter.py`（`_build_index_sheet()`）
- Modify: `tests/test_excel_formatter.py`

- [ ] **Step 1: 寫失敗測試**

在 `tests/test_excel_formatter.py` 尾端加入：

```python
# ── Index quality column ──────────────────────────────────────────────────

def _get_index_ws(tables=None):
    """Run format_workbook with given tables and return the Index sheet."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Data_Financials(Q)"
    return wb, format_workbook(wb, tables or []) or wb["Index"]


def _index_ws(tables=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data_Financials(Q)"
    format_workbook(wb, tables or [])
    return wb["Index"]


def test_index_quality_header_exists():
    ws = _index_ws()
    assert ws.cell(row=4, column=5).value == "完成度"


def test_index_quality_col_all_ok():
    tbl = _make_q_table()
    ws = _index_ws([tbl])
    # Data_Financials(Q) is row 5 when it's the only sheet
    cell = ws.cell(row=5, column=5)
    assert "9/9" in str(cell.value)
    assert "✓" in str(cell.value)
    assert cell.font.color.rgb == QUALITY_GREEN


def test_index_quality_col_missing():
    tbl = _make_q_table(missing=["Operating Income", "Capex"])
    ws = _index_ws([tbl])
    cell = ws.cell(row=5, column=5)
    assert "7/9" in str(cell.value)
    assert "⚠" in str(cell.value)
    assert cell.font.color.rgb == QUALITY_ORANGE


def test_index_quality_col_no_q_table():
    """When tables list is empty, Data_Financials(Q) row doesn't exist — no crash."""
    ws = _index_ws([])
    # Just verify header exists and no exception was thrown
    assert ws.cell(row=4, column=5).value == "完成度"


def test_index_header_merged_to_e():
    ws = _index_ws()
    merged = [str(r) for r in ws.merged_cells.ranges]
    assert any("E1" in r for r in merged), f"A1:E1 not merged, got: {merged}"
```

加在檔案最上方 imports 加入：
```python
from excel_formatter import _compute_quality, QUALITY_GREEN, QUALITY_ORANGE
```

- [ ] **Step 2: 確認測試失敗**

```
python -m pytest tests/test_excel_formatter.py -k "index_quality" -v
```
預期：多個 FAILED（`QUALITY_GREEN` 未定義或 E 欄為空）

- [ ] **Step 3: 實作**

在 `excel_formatter.py` 的 `_build_index_sheet()` 中做以下修改：

**3a. 計算品質分數（在函式頂部，`ws = wb.create_sheet(...)` 之前）：**
```python
    quality = _compute_quality(tables)
```

**3b. 調整 header 合併範圍（改 D → E）：**
```python
    ws.merge_cells("A1:E1")   # 原本是 A1:D1
    # ...
    ws.merge_cells("A2:E2")   # 原本是 A2:D2
```

**3c. 在 Row 4 header 加第 5 欄：**
```python
    for col, label in enumerate(["Sheet", "說明", "最早期間", "最新期間", "完成度"], start=1):
```

**3d. 在 Row 5+ 資料列加 E 欄：**

將現有的資料列迴圈：
```python
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

        # ── E 欄：完成度 ──
        e_cell = ws.cell(row=row, column=5)
        e_cell.fill = row_fill
        if tbl.sheet_name == "Data_Financials(Q)" and quality is not None:
            score, total, _ = quality
            if score == total:
                e_cell.value = f"{score}/{total} ✓"
                e_cell.font = Font(color=QUALITY_GREEN, size=10)
            else:
                e_cell.value = f"{score}/{total} ⚠"
                e_cell.font = Font(color=QUALITY_ORANGE, size=10)
        else:
            e_cell.value = "—"
            e_cell.font = Font(color="FF999999", size=10)
```

**3e. 加 E 欄寬度（在 Column widths 區塊）：**
```python
    ws.column_dimensions["E"].width = 10
```

- [ ] **Step 4: 確認測試通過**

```
python -m pytest tests/test_excel_formatter.py -k "index_quality" -v
```
預期：全部 PASSED

- [ ] **Step 5: 確認全套不破壞**

```
python -m pytest tests/test_excel_formatter.py -v
```
預期：全部 PASSED

- [ ] **Step 6: Commit**

```
git add excel_formatter.py tests/test_excel_formatter.py
git commit -m "feat: add quality score column E to Index sheet"
```

---

## Task 3：品質明細區塊（表格下方）

**Files:**
- Modify: `excel_formatter.py`（`_build_index_sheet()` 尾端）
- Modify: `tests/test_excel_formatter.py`

- [ ] **Step 1: 寫失敗測試**

在 `tests/test_excel_formatter.py` 尾端加入：

```python
# ── Index quality detail section ──────────────────────────────────────────

def _find_detail_header_row(ws) -> int | None:
    """Find the row containing the quality detail section header."""
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and "品質明細" in str(cell.value):
                return cell.row
    return None


def test_index_detail_section_present():
    tbl = _make_q_table()
    ws = _index_ws([tbl])
    assert _find_detail_header_row(ws) is not None, "品質明細 section header not found"


def test_index_detail_section_absent_when_no_q_table():
    ws = _index_ws([])
    assert _find_detail_header_row(ws) is None


def test_index_detail_all_ok_shows_check():
    tbl = _make_q_table()
    ws = _index_ws([tbl])
    hdr_row = _find_detail_header_row(ws)
    # Collect all B-col values after header
    b_vals = [ws.cell(row=hdr_row + 1 + i, column=2).value for i in range(9)]
    assert all(v == "✓" for v in b_vals), f"Expected all ✓, got: {b_vals}"


def test_index_detail_missing_row_shows_cross():
    tbl = _make_q_table(missing=["Operating Income", "Capex"])
    ws = _index_ws([tbl])
    hdr_row = _find_detail_header_row(ws)
    # Collect (concept, status) pairs
    rows = {
        ws.cell(row=hdr_row + 1 + i, column=1).value:
        ws.cell(row=hdr_row + 1 + i, column=2).value
        for i in range(9)
    }
    assert "✗" in rows.get("Operating Income", ""), f"Operating Income: {rows.get('Operating Income')}"
    assert "✗" in rows.get("Capex", ""), f"Capex: {rows.get('Capex')}"


def test_index_detail_missing_row_highlighted():
    tbl = _make_q_table(missing=["Capex"])
    ws = _index_ws([tbl])
    hdr_row = _find_detail_header_row(ws)
    # Find Capex row
    for i in range(9):
        a_val = ws.cell(row=hdr_row + 1 + i, column=1).value
        if a_val == "Capex":
            fill_rgb = ws.cell(row=hdr_row + 1 + i, column=1).fill.fgColor.rgb
            assert fill_rgb == QUALITY_MISS_BG, f"Expected orange bg, got: {fill_rgb}"
            return
    pytest.fail("Capex row not found in detail section")
```

imports 區加入：
```python
from excel_formatter import QUALITY_MISS_BG
```

- [ ] **Step 2: 確認測試失敗**

```
python -m pytest tests/test_excel_formatter.py -k "index_detail" -v
```
預期：全部 FAILED（品質明細區塊尚未實作）

- [ ] **Step 3: 實作**

在 `_build_index_sheet()` 的 Column widths 設定之後加入：

```python
    # ── 品質明細區塊 ───────────────────────────────────────────────────────
    if quality is None:
        return

    score, total, missing = quality
    next_row = 5 + len(data_sheets) + 2   # blank row gap

    # Section header
    hdr_cell = ws.cell(row=next_row, column=1, value=f"品質明細 — Data_Financials(Q)")
    hdr_cell.fill = _fill(NAVY_DARK)
    hdr_cell.font = Font(color="FFFFFFFF", bold=True, size=10)
    ws.merge_cells(
        start_row=next_row, start_column=1,
        end_row=next_row, end_column=2
    )
    next_row += 1

    for row_name in ALL_KEY_ROWS:
        is_missing = row_name in missing
        bg = _fill(QUALITY_MISS_BG) if is_missing else _fill(ROW_WHITE)
        status = "✗  缺失" if is_missing else "✓"
        fg = QUALITY_MISS_FG if is_missing else "FF1A7A34"

        a = ws.cell(row=next_row, column=1, value=row_name)
        b = ws.cell(row=next_row, column=2, value=status)
        a.fill = bg
        b.fill = bg
        a.font = Font(color=fg, size=10)
        b.font = Font(color=fg, size=10)
        next_row += 1
```

- [ ] **Step 4: 確認測試通過**

```
python -m pytest tests/test_excel_formatter.py -k "index_detail" -v
```
預期：全部 PASSED

- [ ] **Step 5: 確認全套不破壞**

```
python -m pytest tests/test_excel_formatter.py -v
```
預期：全部 PASSED

- [ ] **Step 6: 確認全專案測試無破壞**

```
python -m pytest tests/ --ignore=tests/test_live_snapshots.py -q
```
預期：全部 PASSED

- [ ] **Step 7: Commit**

```
git add excel_formatter.py tests/test_excel_formatter.py
git commit -m "feat: add quality detail section to Index sheet"
```

---

## 完成後驗收

- [ ] 執行全套 unit tests：`python -m pytest tests/ --ignore=tests/test_live_snapshots.py -v`
- [ ] 全部 PASSED
- [ ] 人工驗證：`python scripts/smoke_test_10.py`（或手動 fetch AAPL），開啟 Excel → Index sheet → 確認：
  - E 欄標題「完成度」出現
  - Data_Financials(Q) 行顯示分數（如「9/9 ✓」）
  - 表格下方出現品質明細區塊，9 行全 ✓
  - A1:E1 合併正確（公司名稱延伸到 E 欄）
