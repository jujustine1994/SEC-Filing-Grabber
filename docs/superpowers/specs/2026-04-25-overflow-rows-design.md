# Overflow Rows Design (B1)

**Date:** 2026-04-25  
**Status:** Approved — ready for implementation planning  
**Goal:** Append all unmatched XBRL rows after each section's template rows, so no data is silently lost.

---

## Motivation

The fixed universal template (IS 22 rows, BS 41 rows, CF 26 rows) covers standard GAAP concepts across most non-financial companies. However, company-specific or industry-specific XBRL items that don't match any template entry currently produce `None` and are silently discarded.

User requirement: numbers must not be lost. Duplicate rows are acceptable. Template row positions and cross-company comparability must be preserved.

---

## Design

### Principle

Each `StatementTable` returned by `_build_is_table`, `_build_bs_table`, `_build_cf_table` is extended with **overflow rows** appended after the fixed template rows.

Overflow rows = all XBRL rows in the same statement DataFrame that:
1. Are non-abstract, non-breakdown, no dimension (same `_consolidated_mask` as template matching)
2. Were **not consumed** by any template match in that filing

The `_merge_financials` function requires no changes — it already iterates all rows of each `StatementTable` sequentially.

### StatementTable Layout (after change)

```
concepts[0..N-1]   = template row labels (unchanged)
concepts[N..]      = overflow: original XBRL label (Col A)

labels[0..N-1]     = original XBRL labels (unchanged)
labels[N..]        = overflow: XBRL concept name (Col B, used as identifier)

values[0..N-1]     = template row values (unchanged)
values[N..]        = overflow row values (None for quarters where concept absent)
```

### Excel Output Structure

```
Income Statement         ← section header
  [22 IS template rows]
  [IS overflow rows]     ← new, Col A = original label, Col B = concept name
                         ← blank separator (existing)
Balance Sheet            ← section header
  [41 BS template rows]
  [BS overflow rows]     ← new
                         ← blank separator (existing)
Cash Flow                ← section header
  [26 CF template rows]
  [CF overflow rows]     ← new
```

No additional separator row is inserted between template and overflow sections — the existing blank rows before each section header are sufficient.

---

## Implementation

### Changes to `fetcher_gaap.py`

Modify `_build_is_table`, `_build_bs_table`, `_build_cf_table`.

**Per-filing loop changes (each of the three functions):**

1. Add `consumed: set[int] = set()` before the template row loop.
2. After each `_match_is_row(df, ...)` call that returns a non-None index, add the index to `consumed`.
3. For IS: only track consumed indices from `df` (IS DataFrame). CF-sourced rows (`source == "CF"`) consume `cf_df` indices — do NOT add these to IS `consumed` (CF overflow is handled by `_build_cf_table`).
4. After the post-processing fallback calls (ProfitLoss, D&A label fallback etc.), also add their matched indices to `consumed`.

**Overflow collection (per filing, after template loop):**

```python
overflow_data: dict[str, dict] = {}   # initialized before the filing loop, outside it

mask = _consolidated_mask(df)
df_c = df[mask]
for _, row in df_c[~df_c.index.isin(consumed)].iterrows():
    key = str(row.get("concept", "") or "")
    if not key or key == "nan":
        continue
    if key not in overflow_data:
        raw = str(row.get("label", "") or "")
        overflow_data[key] = {
            "label": unicodedata.normalize("NFKC", raw),
            "periods": {}
        }
    val = _to_python_val(row.get(q_col))
    if val is not None:
        overflow_data[key]["periods"][label] = val  # label = FY2024Q1 etc.
```

**After the filing loop, before constructing StatementTable:**

```python
for key in sorted(overflow_data):
    entry = overflow_data[key]
    display_label = entry["label"] or key
    concepts.append(display_label)          # Col A = original XBRL label
    labels_col.append(key)                  # Col B = concept name
    values.append([entry["periods"].get(q) for q in sorted_labels])
```

**Overflow rows with all-None values are skipped** (no point adding them — concept exists in XBRL but has no numerical data).

### No changes required

- `_merge_financials` — already iterates all rows generically
- `override_engine.py` — `check_key_rows` finds template rows by name; overflow names don't collide
- `excel_formatter.py` — overflow rows get default formatting (no standard name to trigger special rules)
- `test_live_snapshots.py`, `test_override_engine.py` — template row indices unchanged

---

## Edge Cases

| Case | Handling |
|------|----------|
| Same concept in multiple filings | `overflow_data` dict deduplicates by concept key; values collected per quarter |
| Concept in filing A but not B | Value = `None` for quarters where concept absent (consistent with template rows) |
| IS template `source == "CF"` rows | Only IS `df` consumed indices tracked; CF consumed indices handled in `_build_cf_table` |
| CF YTD filings (Q2/Q3) | Overflow collection uses same `q_col` guard as template rows — YTD filings already skipped |
| Abstract / breakdown / dimension rows | Excluded by `_consolidated_mask` (same as template matching) |
| Overflow row with no numerical value | Skipped (`if val is not None` check) |
| `concept` field = "nan" or empty | Skipped |

---

## Impact on Tests

- **Unit tests** (`test_override_engine.py`): no changes needed — tests target template rows only
- **Live snapshot tests** (`test_live_snapshots.py`): no changes needed — `check_key_rows` unaffected
- **New smoke test** (optional): after implementation, run one ticker (e.g. COHR) and verify overflow rows appear for Basic/Diluted Shares

---

## Next Steps

1. Implement in `_build_is_table` (largest, most complex due to CF-source rows)
2. Implement in `_build_bs_table`
3. Implement in `_build_cf_table`
4. Smoke test: COHR (known to have Shares data missing in template) and one standard ticker (AAPL)
5. Verify `excel_formatter.py` handles unknown concept names gracefully
