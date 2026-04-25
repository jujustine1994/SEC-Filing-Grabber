# Overflow Rows Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Append all XBRL rows not matched by the fixed template into each statement section as overflow rows; route Non-GAAP-labelled overflow to a separate `Data_Financials_NG` sheet.

**Architecture:** Three build functions (`_build_is_table`, `_build_bs_table`, `_build_cf_table`) each track which DataFrame indices are consumed by template matching, then collect the remaining non-abstract rows into two overflow dicts (GAAP and NG). Each function returns `(gaap_tbl, ng_tbl)`. `fetch_gaap_statements` merges each pair and produces `Data_Financials_NG(Q/Y)` when the NG tables are non-empty. `_merge_financials` is unchanged.

**Tech Stack:** Python 3.13, pandas, edgartools, openpyxl

**Spec:** `docs/superpowers/specs/2026-04-25-overflow-rows-design.md`

---

## File Map

| File | Change |
|------|--------|
| `fetcher_gaap.py` | Add helpers; modify 3 build functions + `fetch_gaap_statements` |
| `tests/test_fetcher_gaap.py` | New — unit tests for helpers |

No changes to `override_engine.py`, `excel_writer.py`, `excel_formatter.py`, `test_live_snapshots.py`, `test_override_engine.py`.

---

## Task 1: Add `_is_nongaap_label` and `_collect_overflow` helpers

**Files:**
- Modify: `fetcher_gaap.py` (add after `_row_key` / `_seg_sheet_suffix` helpers, before `_build_template_table`)
- Create: `tests/test_fetcher_gaap.py`

- [ ] **Step 1.1: Write failing unit tests**

Create `tests/test_fetcher_gaap.py`:

```python
"""Unit tests for fetcher_gaap helpers (no live EDGAR calls)."""
import pandas as pd
import pytest
from fetcher_gaap import _is_nongaap_label, _collect_overflow


# ── _is_nongaap_label ────────────────────────────────────────────────────────

def test_nongaap_label_non_gaap():
    assert _is_nongaap_label("Non-GAAP Revenue") is True

def test_nongaap_label_adjusted():
    assert _is_nongaap_label("Adjusted Operating Income") is True

def test_nongaap_label_excluding():
    assert _is_nongaap_label("Gross profit excluding discontinued ops") is True

def test_nongaap_label_gaap_row():
    assert _is_nongaap_label("Revenue") is False

def test_nongaap_label_total_assets():
    assert _is_nongaap_label("Total assets") is False

def test_nongaap_label_case_insensitive():
    assert _is_nongaap_label("NON-GAAP EPS") is True


# ── _collect_overflow ─────────────────────────────────────────────────────────

def _make_df():
    """Minimal DataFrame mimicking edgartools output with 3 rows."""
    return pd.DataFrame({
        "concept":                ["GrossProfit",   "OperatingLeaseAsset", "AdjustedGrossProfit"],
        "label":                  ["Gross profit",  "Operating lease ROU", "Adjusted gross profit"],
        "standard_concept":       ["GrossProfit",   None,                  None],
        "abstract":               [False,           False,                 False],
        "is_breakdown":           [False,           False,                 False],
        "dimension_member_label": [None,            None,                  None],
        "2024-03-31 (Q1)":        [50_000,          10_000,                55_000],
    })


def test_collect_overflow_gaap_row_captured():
    df = _make_df()
    consumed = {0}   # GrossProfit consumed by template
    gaap, ng = {}, {}
    _collect_overflow(df, consumed, "2024-03-31 (Q1)", "FY2024Q1", gaap, ng)
    assert "OperatingLeaseAsset" in gaap
    assert gaap["OperatingLeaseAsset"]["periods"]["FY2024Q1"] == 10_000


def test_collect_overflow_ng_row_routed():
    df = _make_df()
    consumed = {0}
    gaap, ng = {}, {}
    _collect_overflow(df, consumed, "2024-03-31 (Q1)", "FY2024Q1", gaap, ng)
    assert "AdjustedGrossProfit" in ng
    assert ng["AdjustedGrossProfit"]["periods"]["FY2024Q1"] == 55_000


def test_collect_overflow_consumed_excluded():
    df = _make_df()
    consumed = {0}
    gaap, ng = {}, {}
    _collect_overflow(df, consumed, "2024-03-31 (Q1)", "FY2024Q1", gaap, ng)
    assert "GrossProfit" not in gaap
    assert "GrossProfit" not in ng


def test_collect_overflow_none_value_not_stored():
    """None values are not stored in periods dict (collected only when non-None)."""
    df = _make_df()
    df.loc[1, "2024-03-31 (Q1)"] = None  # OperatingLeaseAsset has no value
    consumed = {0}
    gaap, ng = {}, {}
    _collect_overflow(df, consumed, "2024-03-31 (Q1)", "FY2024Q1", gaap, ng)
    # Key still added (concept exists), but periods is empty
    assert gaap["OperatingLeaseAsset"]["periods"] == {}


def test_collect_overflow_accumulates_across_quarters():
    """Calling _collect_overflow twice with different quarters merges periods."""
    df1 = _make_df()
    df2 = _make_df()
    df2.loc[1, "2024-06-30 (Q2)"] = 11_000
    df2 = df2.rename(columns={"2024-03-31 (Q1)": "2024-06-30 (Q2)"})
    consumed = {0}
    gaap, ng = {}, {}
    _collect_overflow(df1, consumed, "2024-03-31 (Q1)", "FY2024Q1", gaap, ng)
    _collect_overflow(df2, consumed, "2024-06-30 (Q2)", "FY2024Q2", gaap, ng)
    assert "FY2024Q1" in gaap["OperatingLeaseAsset"]["periods"]
    assert "FY2024Q2" in gaap["OperatingLeaseAsset"]["periods"]
```

- [ ] **Step 1.2: Run tests to confirm they fail**

```
pytest tests/test_fetcher_gaap.py -v
```
Expected: `ImportError` or `AttributeError` — `_is_nongaap_label` and `_collect_overflow` not yet defined.

- [ ] **Step 1.3: Add helpers to `fetcher_gaap.py`**

Add the following block **after the `_seg_sheet_suffix` function** (around line 373), before `_build_template_table`:

```python
# ── Overflow helpers ──────────────────────────────────────────────────────────

# Labels containing these substrings are routed to the Non-GAAP overflow sheet
# instead of the GAAP overflow section.  Matching is case-insensitive substring.
_NONGAAP_KEYWORDS: frozenset[str] = frozenset({
    "non-gaap", "non gaap", "adjusted", "excluding", "excl.", "ex-",
})


def _is_nongaap_label(label: str) -> bool:
    """Return True if label looks like a Non-GAAP / adjusted metric."""
    low = label.lower()
    return any(kw in low for kw in _NONGAAP_KEYWORDS)


def _collect_overflow(
    df: pd.DataFrame,
    consumed: set[int],
    data_col: str,
    quarter_label: str,
    gaap_out: dict,
    ng_out: dict,
) -> None:
    """Collect unmatched XBRL rows from df into gaap_out or ng_out.

    Rows whose index is in `consumed` are skipped (already used by template).
    Abstract, breakdown, and dimension rows are excluded via _consolidated_mask.
    Rows with None values are recorded in the dict but not added to periods
    (a later all-None check decides whether to include the row in output).
    """
    mask = _consolidated_mask(df)
    df_c = df[mask]
    remaining = df_c[~df_c.index.isin(consumed)]
    for _, row in remaining.iterrows():
        key = str(row.get("concept", "") or "")
        if not key or key == "nan":
            continue
        raw = str(row.get("label", "") or "")
        display = unicodedata.normalize("NFKC", raw)
        out = ng_out if _is_nongaap_label(display) else gaap_out
        if key not in out:
            out[key] = {"label": display, "periods": {}}
        val = _to_python_val(row.get(data_col))
        if val is not None:
            out[key]["periods"][quarter_label] = val
```

- [ ] **Step 1.4: Run tests — expect PASS**

```
pytest tests/test_fetcher_gaap.py -v
```
Expected: 10 passed.

- [ ] **Step 1.5: Commit**

```
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "feat: add _is_nongaap_label and _collect_overflow helpers"
```

---

## Task 2: Modify `_build_is_table`

**Files:**
- Modify: `fetcher_gaap.py` — `_build_is_table` function (lines ~448–574)

`_build_is_table` returns `StatementTable` today. Change it to return `tuple[StatementTable, StatementTable]` — `(gaap_tbl, ng_tbl)`.

- [ ] **Step 2.1: Write failing test for new return type**

Add to `tests/test_fetcher_gaap.py`:

```python
# ── return-type smoke (no live EDGAR — just check signature) ─────────────────

def test_build_is_table_returns_tuple(monkeypatch):
    """_build_is_table must return a 2-tuple of StatementTables."""
    from fetcher_gaap import _build_is_table, StatementTable
    # Pass empty filings list → both tables should be empty StatementTables
    result = _build_is_table([], max_filings=8)
    assert isinstance(result, tuple) and len(result) == 2
    gaap_tbl, ng_tbl = result
    assert isinstance(gaap_tbl, StatementTable)
    assert isinstance(ng_tbl, StatementTable)
    # GAAP table has all IS template rows (concepts list non-empty)
    assert len(gaap_tbl.concepts) > 0
    # NG table for empty filings has no rows
    assert ng_tbl.concepts == []
```

- [ ] **Step 2.2: Run test to confirm fail**

```
pytest tests/test_fetcher_gaap.py::test_build_is_table_returns_tuple -v
```
Expected: FAIL — `_build_is_table` returns `StatementTable`, not `tuple`.

- [ ] **Step 2.3: Modify `_build_is_table`**

Apply these changes to the function body:

**a) Change signature return type annotation:**
```python
def _build_is_table(filings, max_filings: int, is_overrides: dict | None = None) -> tuple[StatementTable, StatementTable]:
```

**b) Add overflow dicts before the filing loop (after `row_labels: dict`):**
```python
    gaap_overflow: dict[str, dict] = {}   # concept_key → {"label": str, "periods": {quarter: value}}
    ng_overflow:   dict[str, dict] = {}
```

**c) Add `consumed: set[int] = set()` at the top of each filing iteration (after the `label in periods` guard):**
```python
        consumed: set[int] = set()
```

**d) In the IS template loop, after every `idx = _match_is_row(df, ...)` call where `source != "CF"`, add `if idx is not None: consumed.add(idx)`:**

Replace the existing `else:` branch (source != "CF") inside the template row loop:
```python
            else:
                idx = _match_is_row(df, std_concept, fallback,
                                    match=match, label_hint=label_hint)
                if idx is not None:
                    consumed.add(idx)   # track consumed IS df index
                val = _to_python_val(df.loc[idx, q_col]) if idx is not None else None
                if idx is not None and i not in row_labels:
                    raw = str(df.loc[idx, "label"] or "")
                    row_labels[i] = unicodedata.normalize("NFKC", raw)
```

**e) In post-processing block 2 (ProfitLoss fallback), add consumed tracking:**
```python
        if row_vals.get(_NET_INCOME_IDX) is None:
            idx = _match_is_row(df, "ProfitLoss", "ProfitLoss")
            if idx is not None:
                consumed.add(idx)   # track this fallback index
                row_vals[_NET_INCOME_IDX] = _to_python_val(df.loc[idx, q_col])
                if _NET_INCOME_IDX not in row_labels:
                    row_labels[_NET_INCOME_IDX] = unicodedata.normalize(
                        "NFKC", str(df.loc[idx, "label"] or ""))
```

*Note: Post-processing blocks 1 (derived), 3 (D&A label fallback from cf_df), and 4 (derived gross profit) do NOT consume IS df indices — no change needed there.*

**f) Add overflow collection at the end of the filing iteration, just before `periods[label] = ...`:**
```python
        _collect_overflow(df, consumed, q_col, label, gaap_overflow, ng_overflow)
```

**g) Replace the `if not periods:` early-return block:**
```python
    if not periods:
        empty = StatementTable(
            sheet_name="Data_IS",
            quarter_labels=[],
            filing_dates=[],
            concepts=[row[0] for row in IS_TEMPLATE],
            values=[[] for _ in IS_TEMPLATE],
            labels=["" for _ in IS_TEMPLATE],
        )
        empty_ng = StatementTable(
            sheet_name="Data_IS_NG",
            quarter_labels=[], filing_dates=[],
            concepts=[], values=[], labels=[],
        )
        return empty, empty_ng
```

**h) Replace the final `return StatementTable(...)` block with overflow-aware construction:**
```python
    sorted_labels = sorted(periods.keys())
    filing_dates  = [periods[lbl][0] for lbl in sorted_labels]

    # ── Build GAAP table (template rows + GAAP overflow) ──────────────────
    concepts_g: list[str]        = [row[0] for row in IS_TEMPLATE]
    labels_g:   list[str]        = [row_labels.get(i, "") for i in range(len(IS_TEMPLATE))]
    values_g:   list[list[Any]]  = [
        [periods[lbl][1].get(i) for lbl in sorted_labels]
        for i in range(len(IS_TEMPLATE))
    ]
    for key in sorted(gaap_overflow):
        entry = gaap_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue   # skip entirely-empty overflow rows
        concepts_g.append(entry["label"] or key)
        labels_g.append(key)
        values_g.append(row)

    gaap_tbl = StatementTable(
        sheet_name="Data_IS",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=concepts_g,
        labels=labels_g,
        values=values_g,
    )

    # ── Build NG table (Non-GAAP overflow only, no template rows) ─────────
    concepts_n: list[str]       = []
    labels_n:   list[str]       = []
    values_n:   list[list[Any]] = []
    for key in sorted(ng_overflow):
        entry = ng_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue
        concepts_n.append(entry["label"] or key)
        labels_n.append(key)
        values_n.append(row)

    ng_tbl = StatementTable(
        sheet_name="Data_IS_NG",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=concepts_n,
        labels=labels_n,
        values=values_n,
    )

    return gaap_tbl, ng_tbl
```

- [ ] **Step 2.4: Run tests**

```
pytest tests/test_fetcher_gaap.py -v
```
Expected: all pass (including `test_build_is_table_returns_tuple`).

- [ ] **Step 2.5: Commit**

```
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "feat: _build_is_table returns (gaap_tbl, ng_tbl) with overflow rows"
```

---

## Task 3: Modify `_build_bs_table`

**Files:**
- Modify: `fetcher_gaap.py` — `_build_bs_table` function (lines ~583–664)

- [ ] **Step 3.1: Write failing test**

Add to `tests/test_fetcher_gaap.py`:

```python
def test_build_bs_table_returns_tuple(monkeypatch):
    from fetcher_gaap import _build_bs_table, StatementTable
    result = _build_bs_table([], max_filings=8)
    assert isinstance(result, tuple) and len(result) == 2
    gaap_tbl, ng_tbl = result
    assert isinstance(gaap_tbl, StatementTable)
    assert len(gaap_tbl.concepts) > 0   # BS template rows present
    assert ng_tbl.concepts == []
```

- [ ] **Step 3.2: Run test to confirm fail**

```
pytest tests/test_fetcher_gaap.py::test_build_bs_table_returns_tuple -v
```
Expected: FAIL.

- [ ] **Step 3.3: Modify `_build_bs_table`**

**a) Change signature:**
```python
def _build_bs_table(filings, max_filings: int, bs_overrides: dict | None = None) -> tuple[StatementTable, StatementTable]:
```

**b) Add overflow dicts after `row_labels`:**
```python
    gaap_overflow: dict[str, dict] = {}
    ng_overflow:   dict[str, dict] = {}
```

**c) Add `consumed: set[int] = set()` inside the filing loop (after the `label in periods` guard).**

**d) In the template row loop, add consumed tracking for every `_match_is_row` call:**
```python
            idx = _match_is_row(df, std_concept, fallback, match=match, label_hint=label_hint)
            if idx is not None:
                consumed.add(idx)
            val = _to_python_val(df.loc[idx, bs_col]) if idx is not None else None
            row_vals[i] = val
            if idx is not None and i not in row_labels:
                raw = str(df.loc[idx, "label"] or "")
                row_labels[i] = unicodedata.normalize("NFKC", raw)
```

**e) Add overflow collection before `periods[label] = ...`:**
```python
        _collect_overflow(df, consumed, bs_col, label, gaap_overflow, ng_overflow)
```

**f) Replace `if not periods:` early-return:**
```python
    if not periods:
        empty = StatementTable(
            sheet_name="Data_BS",
            quarter_labels=[], filing_dates=[],
            concepts=[row[0] for row in BS_TEMPLATE],
            values=[[] for _ in BS_TEMPLATE],
            labels=["" for _ in BS_TEMPLATE],
        )
        empty_ng = StatementTable(
            sheet_name="Data_BS_NG",
            quarter_labels=[], filing_dates=[],
            concepts=[], values=[], labels=[],
        )
        return empty, empty_ng
```

**g) Replace final `return StatementTable(...)` with overflow-aware construction:**
```python
    sorted_labels = sorted(periods.keys())
    filing_dates  = [periods[lbl][0] for lbl in sorted_labels]

    concepts_g = [row[0] for row in BS_TEMPLATE]
    labels_g   = [row_labels.get(i, "") for i in range(len(BS_TEMPLATE))]
    values_g   = [
        [periods[lbl][1].get(i) for lbl in sorted_labels]
        for i in range(len(BS_TEMPLATE))
    ]
    for key in sorted(gaap_overflow):
        entry = gaap_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue
        concepts_g.append(entry["label"] or key)
        labels_g.append(key)
        values_g.append(row)

    gaap_tbl = StatementTable(
        sheet_name="Data_BS",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=concepts_g, labels=labels_g, values=values_g,
    )

    concepts_n, labels_n, values_n = [], [], []
    for key in sorted(ng_overflow):
        entry = ng_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue
        concepts_n.append(entry["label"] or key)
        labels_n.append(key)
        values_n.append(row)

    ng_tbl = StatementTable(
        sheet_name="Data_BS_NG",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=concepts_n, labels=labels_n, values=values_n,
    )

    return gaap_tbl, ng_tbl
```

- [ ] **Step 3.4: Run tests**

```
pytest tests/test_fetcher_gaap.py -v
```
Expected: all pass.

- [ ] **Step 3.5: Commit**

```
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "feat: _build_bs_table returns (gaap_tbl, ng_tbl) with overflow rows"
```

---

## Task 4: Modify `_build_cf_table`

**Files:**
- Modify: `fetcher_gaap.py` — `_build_cf_table` function (lines ~669–794)

CF has YTD complexity: Q2/Q3 filings use cumulative YTD values that are subtracted later to produce standalone quarters. Overflow from YTD filings would require the same subtraction, which is complex. **Overflow is only collected from non-YTD filings (Q1 and FY)** to keep this tractable. Q2/Q3 overflow rows will show `None` for those quarters (acceptable limitation).

- [ ] **Step 4.1: Write failing test**

Add to `tests/test_fetcher_gaap.py`:

```python
def test_build_cf_table_returns_tuple():
    from fetcher_gaap import _build_cf_table, StatementTable
    result = _build_cf_table([], max_filings=8)
    assert isinstance(result, tuple) and len(result) == 2
    gaap_tbl, ng_tbl = result
    assert isinstance(gaap_tbl, StatementTable)
    assert len(gaap_tbl.concepts) > 0   # CF template rows present
    assert ng_tbl.concepts == []
```

- [ ] **Step 4.2: Run test to confirm fail**

```
pytest tests/test_fetcher_gaap.py::test_build_cf_table_returns_tuple -v
```
Expected: FAIL.

- [ ] **Step 4.3: Modify `_build_cf_table`**

**a) Change signature:**
```python
def _build_cf_table(filings, max_filings: int, cf_overrides: dict | None = None) -> tuple[StatementTable, StatementTable]:
```

**b) Add overflow dicts after `row_labels`:**
```python
    gaap_overflow: dict[str, dict] = {}
    ng_overflow:   dict[str, dict] = {}
```

**c) Add `consumed: set[int] = set()` inside the filing loop (after the `label in collected` guard).**

**d) In the CF template row loop, add consumed tracking for every `_match_is_row` call:**
```python
            idx = _match_is_row(df, std_concept, fallback, match=match, label_hint=label_hint)
            if idx is not None:
                consumed.add(idx)
            val = _to_python_val(df.loc[idx, data_col]) if idx is not None else None
            row_vals[i] = val
            if idx is not None and i not in row_labels:
                raw = str(df.loc[idx, "label"] or "")
                row_labels[i] = unicodedata.normalize("NFKC", raw)
```

**e) Add overflow collection before `collected[label] = ...`, but only for non-YTD filings:**
```python
        # Collect overflow only for non-YTD filings (Q1/FY have standalone values).
        # Q2/Q3 YTD overflow would require cross-filing subtraction — deferred.
        if not is_ytd:
            _collect_overflow(df, consumed, data_col, label, gaap_overflow, ng_overflow)
```

**f) Replace `if not collected:` early-return:**
```python
    if not collected:
        empty = StatementTable(
            sheet_name="Data_CF",
            quarter_labels=[], filing_dates=[],
            concepts=[row[0] for row in CF_TEMPLATE],
            values=[[] for _ in CF_TEMPLATE],
            labels=["" for _ in CF_TEMPLATE],
        )
        empty_ng = StatementTable(
            sheet_name="Data_CF_NG",
            quarter_labels=[], filing_dates=[],
            concepts=[], values=[], labels=[],
        )
        return empty, empty_ng
```

**g) Replace the final `return tbl` with overflow-aware construction. The FCF computation remains unchanged. After the existing FCF computation block:**

Replace:
```python
    return tbl
```

With:
```python
    # ── Append GAAP overflow rows to CF table ────────────────────────────
    for key in sorted(gaap_overflow):
        entry = gaap_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue
        tbl.concepts.append(entry["label"] or key)
        tbl.labels.append(key)
        tbl.values.append(row)

    # ── Build NG table ───────────────────────────────────────────────────
    concepts_n, labels_n, values_n = [], [], []
    for key in sorted(ng_overflow):
        entry = ng_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue
        concepts_n.append(entry["label"] or key)
        labels_n.append(key)
        values_n.append(row)

    ng_tbl = StatementTable(
        sheet_name="Data_CF_NG",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=concepts_n, labels=labels_n, values=values_n,
    )

    return tbl, ng_tbl
```

- [ ] **Step 4.4: Run tests**

```
pytest tests/test_fetcher_gaap.py -v
```
Expected: all pass.

- [ ] **Step 4.5: Commit**

```
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "feat: _build_cf_table returns (gaap_tbl, ng_tbl) with overflow rows"
```

---

## Task 5: Update `fetch_gaap_statements` to unpack tuples and produce NG sheets

**Files:**
- Modify: `fetcher_gaap.py` — `fetch_gaap_statements` function (lines ~1060–1139)

The function calls `_build_is_table`, `_build_bs_table`, `_build_cf_table` in two places:
1. For quarterly (10-Q) filings
2. For annual (10-K) filings (if available)
Also calls them again inside the auto-repair branch.

- [ ] **Step 5.1: Write failing test**

Add to `tests/test_fetcher_gaap.py`:

```python
def test_fetch_gaap_statements_not_broken_with_empty_results(monkeypatch):
    """fetch_gaap_statements must not raise when all build functions return empty tables."""
    from fetcher_gaap import fetch_gaap_statements, _build_is_table, _build_bs_table, _build_cf_table, StatementTable

    empty_is = (
        StatementTable("Data_IS", [], [], ["Revenue"], [[]], [""]),
        StatementTable("Data_IS_NG", [], [], [], [], []),
    )
    empty_bs = (
        StatementTable("Data_BS", [], [], ["Total Assets"], [[]], [""]),
        StatementTable("Data_BS_NG", [], [], [], [], []),
    )
    empty_cf = (
        StatementTable("Data_CF", [], [], ["Operating Cash Flow"], [[]], [""]),
        StatementTable("Data_CF_NG", [], [], [], [], []),
    )

    monkeypatch.setattr("fetcher_gaap._build_is_table", lambda *a, **kw: empty_is)
    monkeypatch.setattr("fetcher_gaap._build_bs_table", lambda *a, **kw: empty_bs)
    monkeypatch.setattr("fetcher_gaap._build_cf_table", lambda *a, **kw: empty_cf)

    # Also mock Company so no HTTP call is made
    class _FakeCompany:
        name = "FAKE"
        def get_filings(self, form, amendments):
            return [object()]   # one fake filing to avoid ValueError
    monkeypatch.setattr("fetcher_gaap.Company", lambda ticker: _FakeCompany())
    monkeypatch.setattr("fetcher_gaap.set_identity", lambda x: None)
    monkeypatch.setattr("fetcher_gaap.load_overrides", lambda ticker: {})

    # Should not raise; Data_Financials_NG should NOT be in results (all NG tables empty)
    results = fetch_gaap_statements("FAKE", "test@test.com", max_filings=1, max_annual_filings=0)
    sheet_names = [t.sheet_name for t in results]
    assert "Data_Financials(Q)" in sheet_names
    assert not any("NG" in s for s in sheet_names)
```

- [ ] **Step 5.2: Run test to confirm fail**

```
pytest tests/test_fetcher_gaap.py::test_fetch_gaap_statements_not_broken_with_empty_results -v
```
Expected: FAIL (tuple unpacking errors because build functions still return StatementTable not tuple).

- [ ] **Step 5.3: Update `fetch_gaap_statements` — quarterly section**

Replace the three build calls and related blocks:

**First occurrence (lines ~1073–1075):**
```python
    is_tbl, is_ng_tbl = _build_is_table(filings_q, max_filings, is_overrides=overrides.get("IS", {}))
    bs_tbl, bs_ng_tbl = _build_bs_table(filings_q, max_filings, bs_overrides=overrides.get("BS", {}))
    cf_tbl, cf_ng_tbl = _build_cf_table(filings_q, max_filings, cf_overrides=overrides.get("CF", {}))
```

**Auto-repair rebuild (lines ~1112–1114):**
```python
            is_tbl, is_ng_tbl = _build_is_table(filings_q, max_filings, is_overrides=overrides.get("IS", {}))
            bs_tbl, bs_ng_tbl = _build_bs_table(filings_q, max_filings, bs_overrides=overrides.get("BS", {}))
            cf_tbl, cf_ng_tbl = _build_cf_table(filings_q, max_filings, cf_overrides=overrides.get("CF", {}))
```

**Merge + NG sheet (lines ~1122–1123), replace with:**
```python
    quarterly_tbl = _merge_financials(is_tbl, bs_tbl, cf_tbl, sheet_name="Data_Financials(Q)")
    tables: list[StatementTable] = [quarterly_tbl]

    # Add Non-GAAP overflow sheet only when at least one NG table has data
    if any(t.quarter_labels for t in [is_ng_tbl, bs_ng_tbl, cf_ng_tbl]):
        ng_q_tbl = _merge_financials(is_ng_tbl, bs_ng_tbl, cf_ng_tbl,
                                     sheet_name="Data_Financials_NG(Q)")
        tables.append(ng_q_tbl)
```

**Annual section (lines ~1127–1131), replace with:**
```python
    filings_k = list(company.get_filings(form="10-K", amendments=False))
    if filings_k:
        is_ann, is_ng_ann = _build_is_table(filings_k, max_annual_filings, is_overrides=overrides.get("IS", {}))
        bs_ann, bs_ng_ann = _build_bs_table(filings_k, max_annual_filings, bs_overrides=overrides.get("BS", {}))
        cf_ann, cf_ng_ann = _build_cf_table(filings_k, max_annual_filings, cf_overrides=overrides.get("CF", {}))
        annual_tbl = _merge_financials(is_ann, bs_ann, cf_ann, sheet_name="Data_Financials(Y)")
        tables.append(annual_tbl)
        if any(t.quarter_labels for t in [is_ng_ann, bs_ng_ann, cf_ng_ann]):
            ng_y_tbl = _merge_financials(is_ng_ann, bs_ng_ann, cf_ng_ann,
                                         sheet_name="Data_Financials_NG(Y)")
            tables.append(ng_y_tbl)
```

- [ ] **Step 5.4: Run all unit tests**

```
pytest tests/test_fetcher_gaap.py tests/test_override_engine.py -v
```
Expected: all pass. `test_override_engine.py` must still pass unchanged (template row indices unaffected by overflow).

- [ ] **Step 5.5: Commit**

```
git add fetcher_gaap.py tests/test_fetcher_gaap.py
git commit -m "feat: fetch_gaap_statements produces Data_Financials_NG when Non-GAAP overflow found"
```

---

## Task 6: Update CHANGELOG

**Files:**
- Modify: `CHANGELOG.md`

- [ ] **Step 6.1: Prepend Session 13 entry**

Add at the top of `CHANGELOG.md` (after the title):

```markdown
## Session 13 — 2026-04-26: Overflow Rows (B1) + Non-GAAP Separation

### New
- `_is_nongaap_label(label)` — keyword-based Non-GAAP detector
- `_collect_overflow(df, consumed, col, quarter, gaap_out, ng_out)` — collects unmatched XBRL rows into GAAP / NG buckets
- `_build_is_table`, `_build_bs_table`, `_build_cf_table` now return `(gaap_tbl, ng_tbl)` tuples
- GAAP overflow rows appended after template rows in `Data_Financials(Q/Y)`
- `Data_Financials_NG(Q/Y)` sheet produced when Non-GAAP overflow rows exist
- `tests/test_fetcher_gaap.py` — new unit test file for fetcher helpers

### Changed
- `fetch_gaap_statements` unpacks 2-tuples from build functions; auto-repair branch updated accordingly

### Known Limitation
- CF Q2/Q3 overflow rows collected from non-YTD filings only (Q2/Q3 overflow = None); full YTD subtraction for overflow deferred (see spec)
```

- [ ] **Step 6.2: Commit**

```
git add CHANGELOG.md
git commit -m "docs: Session 13 changelog — overflow rows B1 + NG separation"
```

---

## Self-Review Notes

**Spec coverage check:**
- ✅ GAAP overflow appended after template rows
- ✅ Non-GAAP overflow in separate sheet
- ✅ Consumed index tracking in all three build functions
- ✅ `_is_nongaap_label` keyword detection
- ✅ All-None overflow rows skipped
- ✅ `_merge_financials` unchanged
- ✅ `check_key_rows` unaffected (template row positions unchanged)
- ✅ CF Q2/Q3 overflow limitation documented

**Placeholder scan:** No TBD/TODO in plan body.

**Type consistency:**
- `_collect_overflow` signature: `(df, consumed: set[int], data_col: str, quarter_label: str, gaap_out: dict, ng_out: dict) -> None` — used consistently in Tasks 2/3/4
- Build functions all return `tuple[StatementTable, StatementTable]` — callers in Task 5 unpack as `gaap_tbl, ng_tbl`
- `ng_tbl.sheet_name` = `"Data_IS_NG"` / `"Data_BS_NG"` / `"Data_CF_NG"` (intermediate); merged sheet = `"Data_Financials_NG(Q)"` — no collision
