# tests/test_fetcher_nongaap.py


def _ng_row(tbl, name):
    """取 Data_NonGAAP 某一列的值（模板化後不能再用 values[0]）。"""
    assert name in tbl.concepts, f"缺列 {name}"
    return tbl.values[tbl.concepts.index(name)]

import json
from pathlib import Path
from fetcher_nongaap import _load_cache, _save_cache, _period_to_quarter_label, _build_eps_recon_table, _build_nongaap_table, _normalize_nongaap_metrics, _filter_nongaap_by_year


def test_period_to_quarter_label_q1():
    assert _period_to_quarter_label("20240331") == "FY2024Q1"

def test_period_to_quarter_label_q2():
    assert _period_to_quarter_label("20240630") == "FY2024Q2"

def test_period_to_quarter_label_q3():
    assert _period_to_quarter_label("20240930") == "FY2024Q3"

def test_period_to_quarter_label_q4():
    assert _period_to_quarter_label("20241231") == "FY2024Q4"

def test_period_with_dashes():
    assert _period_to_quarter_label("2024-03-31") == "FY2024Q1"

def test_load_cache_missing_file():
    result = _load_cache(Path("/nonexistent/nongaap_cache.json"), "AAPL")
    assert result == {}

def test_save_and_load_cache(tmp_path):
    cache_path = tmp_path / "nongaap_cache.json"
    data = {"FY2024Q1": {"metrics": {"Non-GAAP EPS": 0.71}}}
    _save_cache(cache_path, "AAPL", data)
    loaded = _load_cache(cache_path, "AAPL")
    assert loaded["FY2024Q1"]["metrics"]["Non-GAAP EPS"] == 0.71

def test_cache_ticker_isolation(tmp_path):
    """Two tickers in the same output_dir must not interfere with each other."""
    cache_path = tmp_path / "nongaap_cache.json"
    aapl_data = {"FY2024Q1": {"metrics": {"Non-GAAP EPS": 1.50}}}
    nvda_data = {"FY2024Q1": {"metrics": {"Non-GAAP EPS": 5.20}}}
    _save_cache(cache_path, "AAPL", aapl_data)
    _save_cache(cache_path, "NVDA", nvda_data)
    assert _load_cache(cache_path, "AAPL")["FY2024Q1"]["metrics"]["Non-GAAP EPS"] == 1.50
    assert _load_cache(cache_path, "NVDA")["FY2024Q1"]["metrics"]["Non-GAAP EPS"] == 5.20

def test_cache_second_save_does_not_erase_other_ticker(tmp_path):
    """Saving AAPL a second time must not delete NVDA's data."""
    cache_path = tmp_path / "nongaap_cache.json"
    _save_cache(cache_path, "AAPL", {"FY2024Q1": {"metrics": {}}})
    _save_cache(cache_path, "NVDA", {"FY2024Q1": {"metrics": {}}})
    _save_cache(cache_path, "AAPL", {"FY2024Q1": {"metrics": {}}, "FY2024Q2": {"metrics": {}}})
    assert "FY2024Q1" in _load_cache(cache_path, "NVDA")

def test_cache_old_format_migration(tmp_path):
    """Old single-ticker cache (no ticker key) is loaded transparently."""
    import json
    cache_path = tmp_path / "nongaap_cache.json"
    old_data = {"FY2024Q1": {"metrics": {"Non-GAAP EPS": 0.99}}}
    cache_path.write_text(json.dumps(old_data), encoding="utf-8")
    loaded = _load_cache(cache_path, "AAPL")
    assert loaded["FY2024Q1"]["metrics"]["Non-GAAP EPS"] == 0.99

def test_cache_old_format_rewritten_on_save(tmp_path):
    """After loading old format and saving, file is in new multi-ticker format."""
    import json
    cache_path = tmp_path / "nongaap_cache.json"
    old_data = {"FY2024Q1": {"metrics": {"Non-GAAP EPS": 0.99}}}
    cache_path.write_text(json.dumps(old_data), encoding="utf-8")
    # Simulate: load old data, add a new quarter, save
    loaded = _load_cache(cache_path, "AAPL")
    loaded["FY2024Q2"] = {"metrics": {"Non-GAAP EPS": 1.10}}
    _save_cache(cache_path, "AAPL", loaded)
    # File should now be in new format
    on_disk = json.loads(cache_path.read_text(encoding="utf-8"))
    assert "AAPL" in on_disk                             # new format
    assert "FY2024Q1" in on_disk["AAPL"]                # old data preserved
    assert "FY2024Q2" in on_disk["AAPL"]                # new data added
    assert _load_cache(cache_path, "AAPL")["FY2024Q1"]["metrics"]["Non-GAAP EPS"] == 0.99


# ── EPS Recon and NonGAAP Table Builder Tests ───────────────────────────────

SAMPLE_CACHE = {
    "FY2024Q1": {
        "filing_date": "2024-01-24",
        "eps_recon": {"GAAP EPS": 0.53, "SBC": -0.12, "Non-GAAP EPS": 0.65},
        "metrics": {"Non-GAAP Net Income": 2513000000.0, "Adjusted EBITDA": 3800000000.0},
    },
    "FY2024Q2": {
        "filing_date": "2024-04-23",
        "eps_recon": {"GAAP EPS": 0.42, "SBC": -0.10, "Non-GAAP EPS": 0.52},
        "metrics": {"Non-GAAP Net Income": 1800000000.0},
    },
}

def test_build_eps_recon_table_sheet_name():
    tbl = _build_eps_recon_table("TSLA", SAMPLE_CACHE)
    assert tbl is not None
    assert tbl.sheet_name == "Data_EPS_Recon"

def test_build_eps_recon_table_quarters_oldest_to_newest():
    tbl = _build_eps_recon_table("TSLA", SAMPLE_CACHE)
    assert tbl.quarter_labels == ["FY2024Q1", "FY2024Q2"]

def test_build_eps_recon_table_concepts():
    tbl = _build_eps_recon_table("TSLA", SAMPLE_CACHE)
    assert "GAAP EPS" in tbl.concepts
    assert "Non-GAAP EPS" in tbl.concepts

def test_build_eps_recon_table_values():
    tbl = _build_eps_recon_table("TSLA", SAMPLE_CACHE)
    gaap_idx = tbl.concepts.index("GAAP EPS")
    assert tbl.values[gaap_idx] == [0.53, 0.42]

def test_build_nongaap_table_sheet_name():
    tbl = _build_nongaap_table("TSLA", SAMPLE_CACHE)
    assert tbl is not None
    assert tbl.sheet_name == "Data_NonGAAP"

def test_build_nongaap_table_union_of_metrics():
    tbl = _build_nongaap_table("TSLA", SAMPLE_CACHE)
    # Adjusted EBITDA only in Q1, should still appear with None in Q2
    assert "Adjusted EBITDA" in tbl.concepts
    ebitda_idx = tbl.concepts.index("Adjusted EBITDA")
    assert tbl.values[ebitda_idx] == [3800000000.0, None]

def test_build_eps_recon_table_empty_cache():
    assert _build_eps_recon_table("TSLA", {}) is None

def test_build_nongaap_table_empty_cache():
    assert _build_nongaap_table("TSLA", {}) is None


# ── Deduplication logic test ─────────────────────────────────────────────────

def test_period_to_quarter_label_dedup_logic():
    """Verify dedup keeps oldest filing per quarter (same algorithm as _get_earnings_filings)."""
    # Simulate edgartools newest-first order with a duplicate Q1
    raw_newest_first = [
        ("FY2024Q1", "filing_new_Q1", "ek_new_Q1"),   # newest Q1 (should be discarded)
        ("FY2024Q2", "filing_Q2",     "ek_Q2"),
        ("FY2024Q1", "filing_old_Q1", "ek_old_Q1"),   # oldest Q1 (should be kept)
    ]

    # Apply same dedup logic as _get_earnings_filings
    seen: set = set()
    deduped = []
    for label, filing, eight_k in reversed(raw_newest_first):  # iterate oldest-first
        if label not in seen:
            seen.add(label)
            deduped.append((label, filing, eight_k))
    result = list(reversed(deduped))  # flip back to newest-first (matches edgartools convention)

    # After dedup: oldest Q1 is kept, Q2 retained; result is newest-first
    assert len(result) == 2
    assert result[0] == ("FY2024Q2", "filing_Q2", "ek_Q2")
    assert result[1] == ("FY2024Q1", "filing_old_Q1", "ek_old_Q1")


# ── _normalize_nongaap_metrics tests ─────────────────────────────────────────

def test_normalize_strips_quarterly_suffix():
    raw = {"Non-GAAP Gross margin (Q4 FY26)": 75.2}
    assert _normalize_nongaap_metrics(raw) == {"Non-GAAP Gross margin": 75.2}

def test_normalize_strips_fy_suffix():
    """期間 token 剝掉，但年度值改帶 (FY) 標記另成一列（2026-08-01 起，見
    metric_rules.FY_ONLY_HANDLING）——年度數字不可佔用季欄位的原名。"""
    raw = {"Non-GAAP Net income (FY26)": 4000000000.0}
    assert _normalize_nongaap_metrics(raw) == {"Non-GAAP Net income (FY)": 4000000000.0}

def test_normalize_quarterly_keeps_its_own_column():
    """當季與年度並存時，季欄位必須是當季的數字；年度值另成 (FY) 列並存。
    （2026-08-01 前的行為是年度值直接被丟棄。）"""
    raw = {
        "Non-GAAP EPS (Q4 FY26)": 1.62,
        "Non-GAAP EPS (FY26)": 6.01,
    }
    result = _normalize_nongaap_metrics(raw)
    assert result["Non-GAAP EPS"] == 1.62
    assert result["Non-GAAP EPS (FY)"] == 6.01

def test_normalize_fy_kept_but_labelled_when_no_quarterly():
    """只有年度值時仍保留（不刪資料），但要標記成 (FY)，不可冒充當季值。"""
    raw = {"Non-GAAP Annual Tax Rate (FY26)": 17.0}
    assert _normalize_nongaap_metrics(raw) == {"Non-GAAP Annual Tax Rate (FY)": 17.0}

def test_normalize_drops_expected_prefix():
    raw = {"Expected Non-GAAP Gross margin (Q1 FY27)": 75.0}
    assert _normalize_nongaap_metrics(raw) == {}

def test_normalize_drops_outlook_prefix():
    raw = {"Outlook Non-GAAP Operating expenses": 7500000000.0}
    assert _normalize_nongaap_metrics(raw) == {}

def test_normalize_drops_outlook_in_name():
    raw = {"Non-GAAP Outlook Gross margin": 75.0}
    assert _normalize_nongaap_metrics(raw) == {}

def test_normalize_drops_guidance_prefix():
    raw = {"Guidance Non-GAAP EPS": 1.80}
    assert _normalize_nongaap_metrics(raw) == {}

def test_normalize_keeps_clean_names_unchanged():
    """Names without period suffix pass through unchanged."""
    raw = {"Non-GAAP Revenue": 39300000000.0, "Non-GAAP EPS": 1.62}
    assert _normalize_nongaap_metrics(raw) == raw

def test_normalize_fy2digit_and_4digit():
    """Both FY26 (2-digit) and FY2026 (4-digit) suffixes are stripped."""
    assert _normalize_nongaap_metrics({"Non-GAAP EPS (FY26)": 1.0}) == {"Non-GAAP EPS (FY)": 1.0}
    assert _normalize_nongaap_metrics({"Non-GAAP EPS (FY2026)": 1.0}) == {"Non-GAAP EPS (FY)": 1.0}

def test_normalize_period_in_middle():
    """Period token embedded in name: 'Non-GAAP Q4 FY26 Gross margin' → 'Non-GAAP Gross margin'."""
    raw = {"Non-GAAP Q4 FY26 Gross margin": 75.2}
    assert _normalize_nongaap_metrics(raw) == {"Non-GAAP Gross margin": 75.2}

def test_normalize_period_as_prefix():
    """Period token at front: 'Q2 FY26 Non-GAAP Revenue' → 'Non-GAAP Revenue'."""
    raw = {"Q2 FY26 Non-GAAP Revenue": 30040000000.0}
    assert _normalize_nongaap_metrics(raw) == {"Non-GAAP Revenue": 30040000000.0}

def test_normalize_comparison_periods_deduplicated():
    """When same metric appears for current + prior quarters, first occurrence wins."""
    raw = {
        "Non-GAAP Q4 FY26 Gross margin": 75.2,   # current Q4 — keep
        "Non-GAAP Q3 FY26 Gross margin": 73.6,   # prior Q — discard
        "Non-GAAP Q4 FY25 Gross margin": 73.5,   # year-ago Q — discard
    }
    result = _normalize_nongaap_metrics(raw)
    assert result == {"Non-GAAP Gross margin": 75.2}

def test_normalize_trailing_table_label_stripped():
    """Trailing '(Table)' noise label is removed."""
    raw = {"Non-GAAP Q4 FY26 Gross margin (Table)": 75.2}
    assert _normalize_nongaap_metrics(raw) == {"Non-GAAP Gross margin": 75.2}

def test_normalize_meaningful_parens_kept():
    """Parentheticals with actual content are NOT stripped."""
    raw = {"Non-GAAP EPS excluding H20 charges": 1.62}
    result = _normalize_nongaap_metrics(raw)
    assert "Non-GAAP EPS excluding H20 charges" in result

def test_normalize_empty_input():
    assert _normalize_nongaap_metrics({}) == {}

def test_normalize_mixed_nvda_real_data():
    """Mirrors actual NVDA Q4 FY26 AI output: period-in-name + FY dups + guidance."""
    raw = {
        "Non-GAAP Q4 FY26 Gross margin": 75.2,       # current Q — keep
        "Non-GAAP Q3 FY26 Gross margin": 73.6,        # comparison Q — discard
        "Non-GAAP Q4 FY25 Gross margin": 73.5,        # comparison Q — discard
        "Non-GAAP FY2026 Gross margin": 74.8,         # FY dup — discard (Q exists)
        "Non-GAAP Q4 FY26 Net income": 22067000000.0, # keep
        "Non-GAAP FY2026 Net income": 76019000000.0,  # FY dup — discard
        "Q2 FY26 Non-GAAP Revenue": 30040000000.0,    # period-as-prefix — keep
        "Expected Non-GAAP Gross margin (Q1 FY27)": 75.0,       # guidance — drop
        "Outlook Non-GAAP Operating expenses": 7500000000.0,    # guidance — drop
    }
    result = _normalize_nongaap_metrics(raw)
    # 當季列：三個。年度值不再被丟棄，改成 (FY) 列另放（2026-08-01 起）
    assert {k for k in result if not k.endswith("(FY)")} == {
        "Non-GAAP Gross margin",
        "Non-GAAP Net income",
        "Non-GAAP Revenue",
    }
    assert result["Non-GAAP Gross margin (FY)"] == 74.8
    assert result["Non-GAAP Net income (FY)"] == 76019000000.0
    assert result["Non-GAAP Gross margin"] == 75.2      # current Q value
    assert result["Non-GAAP Net income"] == 22067000000.0
    assert result["Non-GAAP Revenue"] == 30040000000.0


# ── _filter_nongaap_by_year ───────────────────────────────────────────────────

def _make_ng_filing(label: str):
    """Make a (label, filing, eight_k) tuple with a label like 'FY2021Q2'."""
    from unittest.mock import MagicMock
    filing = MagicMock()
    eight_k = MagicMock()
    return (label, filing, eight_k)


def test_filter_nongaap_no_bounds():
    filings = [_make_ng_filing("FY2020Q1"), _make_ng_filing("FY2022Q3")]
    result = _filter_nongaap_by_year(filings, None, None)
    assert len(result) == 2


def test_filter_nongaap_start_year():
    filings = [_make_ng_filing("FY2019Q4"), _make_ng_filing("FY2021Q1")]
    result = _filter_nongaap_by_year(filings, start_year=2020, end_year=None)
    assert len(result) == 1
    assert result[0][0] == "FY2021Q1"


def test_filter_nongaap_end_year():
    filings = [_make_ng_filing("FY2020Q2"), _make_ng_filing("FY2023Q1")]
    result = _filter_nongaap_by_year(filings, start_year=None, end_year=2021)
    assert len(result) == 1
    assert result[0][0] == "FY2020Q2"


def test_filter_nongaap_both_bounds():
    filings = [_make_ng_filing(f"FY{y}Q1") for y in [2018, 2020, 2022, 2024]]
    result = _filter_nongaap_by_year(filings, start_year=2019, end_year=2022)
    years = [int(lbl[2:6]) for lbl, _, _ in result]
    assert years == [2020, 2022]


from fetcher_nongaap import _list_earnings_filings


class FakeFiling:
    """Stand-in for edgartools EntityFiling — only the listing-level attributes."""
    def __init__(self, items, period_of_report, accession="0000000000-00-000000"):
        self.items = items
        self.period_of_report = period_of_report
        self.accession_no = accession
        self.filing_date = period_of_report
        self.obj_called = False

    def obj(self):
        self.obj_called = True
        raise AssertionError("obj() must not be called during listing-stage filtering")


class FakeCompany:
    """Stand-in for edgartools Company. Filings are newest-first, as EDGAR returns them."""
    def __init__(self, filings):
        self._filings = filings

    def get_filings(self, **kwargs):
        return self._filings


def test_list_earnings_filings_keeps_only_item_202():
    company = FakeCompany([
        FakeFiling("2.02,9.01", "2024-09-30"),
        FakeFiling("5.07",      "2024-08-15"),
        FakeFiling("8.01,9.01", "2024-07-01"),
        FakeFiling("2.02",      "2024-06-30"),
    ])
    result = _list_earnings_filings(company)
    assert [label for label, _ in result] == ["FY2024Q3", "FY2024Q2"]


def test_list_earnings_filings_never_downloads():
    """Listing-stage filtering must not call obj() — that is the whole point."""
    filings = [
        FakeFiling("2.02,9.01", "2024-09-30"),
        FakeFiling("5.07",      "2024-08-15"),
    ]
    _list_earnings_filings(FakeCompany(filings))
    assert all(f.obj_called is False for f in filings)


def test_list_earnings_filings_tolerates_missing_items():
    """items may be None or empty on some filings — must not raise."""
    company = FakeCompany([
        FakeFiling(None, "2024-09-30"),
        FakeFiling("",   "2024-06-30"),
        FakeFiling("2.02", "2024-03-31"),
    ])
    result = _list_earnings_filings(company)
    assert [label for label, _ in result] == ["FY2024Q1"]


def test_list_earnings_filings_skips_malformed_period():
    company = FakeCompany([
        FakeFiling("2.02", None),
        FakeFiling("2.02", ""),
        FakeFiling("2.02", "2024-03-31"),
    ])
    result = _list_earnings_filings(company)
    assert [label for label, _ in result] == ["FY2024Q1"]


def test_list_earnings_filings_dedupes_keeping_newest():
    """Same label twice with equally good metadata: keep the newest filing.

    A preliminary release always precedes the final one, so "newest" is the rule
    that keeps the official numbers. Amendments never reach here
    (``get_filings(amendments=False)``).
    """
    oldest = FakeFiling("2.02,9.01", "2024-09-30", accession="OLDEST")
    newer  = FakeFiling("2.02,9.01", "2024-09-30", accession="NEWER")
    company = FakeCompany([newer, oldest])   # newest-first, as EDGAR returns
    result = _list_earnings_filings(company)
    assert len(result) == 1
    assert result[0][1].accession_no == "NEWER"


def test_list_earnings_filings_dedupe_prefers_the_one_with_an_exhibit_wdc():
    """WDC FY2025Q1 (real case): keeping the oldest lost the whole quarter.

    0001193125-25-007725 is an Item 2.02+5.02 filing with **no exhibit** (no
    Item 9.01 → no press release attached); 0000106040-25-000005 filed 13 days
    later is the actual FY2025 Q2 earnings release. Listing metadata is verbatim
    from EDGAR.
    """
    no_exhibit = FakeFiling("2.02,5.02", "2025-01-10", accession="0001193125-25-007725")
    real       = FakeFiling("2.02,9.01", "2025-01-29", accession="0000106040-25-000005")
    company = FakeCompany([real, no_exhibit])
    result = _list_earnings_filings(company)
    assert len(result) == 1
    assert result[0][1].accession_no == "0000106040-25-000005"


def test_list_earnings_filings_dedupe_drops_preliminary_release_qrvo():
    """QRVO FY2025Q4 (real case): keeping the oldest kept the *preliminary* numbers.

    0000950103-25-013685 (2025-10-28) is titled "Preliminary Fiscal 2026 Second
    Quarter Results"; 0001628280-25-048216 (2025-11-03) is the official release.
    Both carry Item 9.01, so the exhibit test cannot separate them — recency can.
    """
    preliminary = FakeFiling("2.02,9.01", "2025-10-28", accession="0000950103-25-013685")
    official    = FakeFiling("2.02,9.01", "2025-11-03", accession="0001628280-25-048216")
    company = FakeCompany([official, preliminary])
    result = _list_earnings_filings(company)
    assert len(result) == 1
    assert result[0][1].accession_no == "0001628280-25-048216"


def test_list_earnings_filings_dedupe_falls_back_to_newest_without_exhibits():
    """Neither candidate carries Item 9.01: still keep the newest, never the oldest."""
    older = FakeFiling("2.02", "2024-09-15", accession="OLDER")
    newer = FakeFiling("2.02", "2024-09-30", accession="NEWER")
    company = FakeCompany([newer, older])
    result = _list_earnings_filings(company)
    assert [f.accession_no for _, f in result] == ["NEWER"]


def test_list_earnings_filings_dedupe_does_not_download():
    """The exhibit test must read listing metadata only — obj() stays untouched."""
    filings = [
        FakeFiling("2.02,9.01", "2025-01-29", accession="A"),
        FakeFiling("2.02,5.02", "2025-01-10", accession="B"),
    ]
    _list_earnings_filings(FakeCompany(filings))
    assert all(f.obj_called is False for f in filings)


def test_list_earnings_filings_applies_year_range():
    company = FakeCompany([
        FakeFiling("2.02", "2024-03-31"),
        FakeFiling("2.02", "2023-03-31"),
        FakeFiling("2.02", "2022-03-31"),
        FakeFiling("2.02", "2021-03-31"),
    ])
    result = _list_earnings_filings(company, start_year=2022, end_year=2023)
    assert [label for label, _ in result] == ["FY2023Q1", "FY2022Q1"]


def test_list_earnings_filings_max_filings_keeps_newest():
    company = FakeCompany([
        FakeFiling("2.02", "2024-12-31"),
        FakeFiling("2.02", "2024-09-30"),
        FakeFiling("2.02", "2024-06-30"),
        FakeFiling("2.02", "2024-03-31"),
    ])
    result = _list_earnings_filings(company, max_filings=2)
    assert [label for label, _ in result] == ["FY2024Q4", "FY2024Q3"]


def test_list_earnings_filings_skips_non_numeric_period():
    """A period_of_report that is long enough but non-numeric (e.g. 'UNKNOWN0') must be
    skipped, not abort the whole listing — _period_to_quarter_label raises ValueError
    on non-numeric month digits."""
    company = FakeCompany([
        FakeFiling("2.02", "UNKNOWN0"),
        FakeFiling("2.02", "2024-03-31"),
    ])
    result = _list_earnings_filings(company)
    assert [label for label, _ in result] == ["FY2024Q1"]


def test_list_earnings_filings_year_filter_runs_before_max_filings():
    """Year range narrows the pool first; max_filings then trims the narrowed pool."""
    company = FakeCompany([
        FakeFiling("2.02", "2024-12-31"),
        FakeFiling("2.02", "2024-09-30"),
        FakeFiling("2.02", "2023-12-31"),
        FakeFiling("2.02", "2023-09-30"),
    ])
    result = _list_earnings_filings(company, end_year=2023, max_filings=1)
    assert [label for label, _ in result] == ["FY2023Q4"]


from fetcher_nongaap import _find_missing_quarters


def test_find_missing_quarters_none_missing():
    assert _find_missing_quarters(["FY2024Q3", "FY2024Q2", "FY2024Q1"]) == []


def test_find_missing_quarters_single_gap():
    assert _find_missing_quarters(["FY2024Q3", "FY2024Q1"]) == ["FY2024Q2"]


def test_find_missing_quarters_spans_year_boundary():
    assert _find_missing_quarters(["FY2025Q1", "FY2024Q3"]) == ["FY2024Q4"]


def test_find_missing_quarters_ignores_outside_range():
    """Gaps only exist between the oldest and newest label — never before or after."""
    assert _find_missing_quarters(["FY2024Q2", "FY2024Q3"]) == []


def test_find_missing_quarters_multiple_gaps():
    assert _find_missing_quarters(["FY2024Q4", "FY2024Q2", "FY2023Q4"]) == [
        "FY2024Q1", "FY2024Q3",
    ]


def test_find_missing_quarters_handles_empty_and_single():
    assert _find_missing_quarters([]) == []
    assert _find_missing_quarters(["FY2024Q1"]) == []


def test_find_missing_quarters_ignores_unparseable_labels():
    assert _find_missing_quarters(["FY2024Q3", "GARBAGE", "FY2024Q1"]) == ["FY2024Q2"]


from fetcher_nongaap import _recover_missing_quarters


class FakeEightK:
    def __init__(self, has_earnings):
        self.has_earnings = has_earnings


class RecoverableFiling:
    """FakeFiling variant whose obj() returns a parsed 8-K instead of raising."""
    def __init__(self, items, period_of_report, has_earnings, accession="ACC"):
        self.items = items
        self.period_of_report = period_of_report
        self.accession_no = accession
        self.filing_date = period_of_report
        self._has_earnings = has_earnings
        self.obj_calls = 0

    def obj(self):
        self.obj_calls += 1
        return FakeEightK(self._has_earnings)


def test_recover_missing_quarters_finds_mislabelled_earnings():
    target = RecoverableFiling("8.01,9.01", "2024-06-30", has_earnings=True)
    company = FakeCompany([
        RecoverableFiling("2.02", "2024-09-30", has_earnings=True),
        target,
        RecoverableFiling("2.02", "2024-03-31", has_earnings=True),
    ])
    result = _recover_missing_quarters(company, ["FY2024Q2"])
    assert [label for label, _ in result] == ["FY2024Q2"]
    assert result[0][1] is target


def test_recover_missing_quarters_only_downloads_gap_candidates():
    """Filings outside the gap, or already tagged 2.02, must not be downloaded."""
    in_gap     = RecoverableFiling("5.07",  "2024-06-30", has_earnings=True)
    tagged     = RecoverableFiling("2.02",  "2024-06-30", has_earnings=True)
    other_qtr  = RecoverableFiling("8.01",  "2024-09-30", has_earnings=True)
    _recover_missing_quarters(FakeCompany([other_qtr, tagged, in_gap]), ["FY2024Q2"])
    assert in_gap.obj_calls == 1
    assert tagged.obj_calls == 0
    assert other_qtr.obj_calls == 0


def test_recover_missing_quarters_ignores_non_earnings():
    company = FakeCompany([RecoverableFiling("5.07", "2024-06-30", has_earnings=False)])
    assert _recover_missing_quarters(company, ["FY2024Q2"]) == []


def test_recover_missing_quarters_empty_gap_list_downloads_nothing():
    f = RecoverableFiling("5.07", "2024-06-30", has_earnings=True)
    assert _recover_missing_quarters(FakeCompany([f]), []) == []
    assert f.obj_calls == 0


def test_recover_missing_quarters_survives_download_failure():
    """A filing that fails to parse must not abort recovery of the others."""
    class ExplodingFiling(RecoverableFiling):
        def obj(self):
            self.obj_calls += 1
            raise ValueError("parse failed")

    good = RecoverableFiling("5.07", "2024-06-30", has_earnings=True)
    bad  = ExplodingFiling("8.01", "2024-06-30", has_earnings=True, accession="BAD")
    result = _recover_missing_quarters(FakeCompany([bad, good]), ["FY2024Q2"])
    assert [label for label, _ in result] == ["FY2024Q2"]
    assert result[0][1] is good


def test_recover_missing_quarters_skips_non_numeric_period():
    """A gap-candidate filing with a non-numeric period_of_report (e.g. 'UNKNOWN0')
    must be skipped, not abort recovery of the other gap-quarter candidates —
    _period_to_quarter_label raises ValueError on non-numeric month digits."""
    bad  = RecoverableFiling("8.01", "UNKNOWN0", has_earnings=True, accession="BAD")
    good = RecoverableFiling("8.01,9.01", "2024-06-30", has_earnings=True)
    result = _recover_missing_quarters(FakeCompany([bad, good]), ["FY2024Q2"])
    assert [label for label, _ in result] == ["FY2024Q2"]
    assert result[0][1] is good


def test_recover_missing_quarters_skips_unparseable_ordinal():
    """A period like 'ABCD0331' survives _period_to_quarter_label without raising
    (the year slice is never int()-ed) but produces a label whose ordinal is None —
    must be skipped rather than crash the sort with a TypeError."""
    bad  = RecoverableFiling("8.01", "ABCD0331", has_earnings=True, accession="BAD")
    good = RecoverableFiling("8.01,9.01", "2024-06-30", has_earnings=True)
    result = _recover_missing_quarters(FakeCompany([bad, good]), ["FY2024Q2"])
    assert [label for label, _ in result] == ["FY2024Q2"]
    assert result[0][1] is good


def test_fetch_nongaap_downloads_only_uncached_quarters(tmp_path, monkeypatch):
    """The whole point: one obj() call per quarter actually needed, and none for
    quarters already in the cache."""
    import fetcher_nongaap as fn

    q3 = RecoverableFiling("2.02", "2024-09-30", has_earnings=True)
    q2 = RecoverableFiling("2.02", "2024-06-30", has_earnings=True)
    q1 = RecoverableFiling("2.02", "2024-03-31", has_earnings=True)
    company = FakeCompany([q3, q2, q1])

    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})
    monkeypatch.setattr(fn, "_extract_nongaap_metrics", lambda ek, cfg: {"Non-GAAP EPS": 1.0})

    # FY2024Q1 is already cached — it must not be downloaded again
    fn._save_cache(tmp_path / fn.CACHE_FILENAME, "TEST",
                   {"FY2024Q1": {"filing_date": "2024-04-30", "eps_recon": {}, "metrics": {}}})

    fn.fetch_nongaap_statements("TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path)

    assert q3.obj_calls == 1
    assert q2.obj_calls == 1
    assert q1.obj_calls == 0


def test_fetch_nongaap_writes_metrics_to_cache(tmp_path, monkeypatch):
    import fetcher_nongaap as fn

    company = FakeCompany([RecoverableFiling("2.02", "2024-09-30", has_earnings=True)])
    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})
    monkeypatch.setattr(fn, "_extract_nongaap_metrics", lambda ek, cfg: {"Non-GAAP EPS": 2.5})

    fn.fetch_nongaap_statements("TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path)

    cached = fn._load_cache(tmp_path / fn.CACHE_FILENAME, "TEST")
    assert cached["FY2024Q3"]["metrics"]["Non-GAAP EPS"] == 2.5


def test_fetch_nongaap_recovers_gap_quarter_in_newest_first_order(tmp_path, monkeypatch):
    """Wiring test for the gap-recovery block in fetch_nongaap_statements: a quarter
    whose 8-K omits Item 2.02 (so _list_earnings_filings misses it) must still be
    found via _recover_missing_quarters, merged back into the run, downloaded, and
    cached — and the merge must preserve newest-first processing order. Without
    reverse=True on the merge sort, this would process oldest-first instead."""
    import fetcher_nongaap as fn

    q3 = RecoverableFiling("2.02", "2024-09-30", has_earnings=True)
    q2 = RecoverableFiling("5.07", "2024-06-30", has_earnings=True)  # not tagged 2.02 -> gap
    q1 = RecoverableFiling("2.02", "2024-03-31", has_earnings=True)
    company = FakeCompany([q3, q2, q1])

    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})
    monkeypatch.setattr(fn, "_extract_nongaap_metrics", lambda ek, cfg: {"Non-GAAP EPS": 1.0})

    seen_labels = []
    fn.fetch_nongaap_statements(
        "TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path,
        progress_cb=lambda i, total, label: seen_labels.append(label),
    )

    # The gap quarter was found by recovery (1 obj() call to check has_earnings)
    # and then downloaded again in the main loop to extract data (1 more call).
    assert q2.obj_calls == 2
    cached = fn._load_cache(tmp_path / fn.CACHE_FILENAME, "TEST")
    assert "FY2024Q2" in cached

    # Processing order stays newest-first with the recovered quarter slotted in
    # at its correct position, not appended at the end or reversed.
    assert len(seen_labels) == 3
    assert "FY2024Q3" in seen_labels[0]
    assert "FY2024Q2" in seen_labels[1]
    assert "FY2024Q1" in seen_labels[2]


# ═════════════════════════════════════════════════════════════════════════════
# 中文指標名稱處理（2026-08-01 新增，TODO 第 2 項）
#
# 背景：AI prompt 原為中文，回傳的指標名時中時英（同一 ticker 內都會混），
# 而下游三條規則（期間剝除、guidance 過濾、Excel ÷1M 豁免）全部只認英文。
# 規則表集中於 metric_rules.py，測試對著那張表寫。
# ═════════════════════════════════════════════════════════════════════════════

# ── L1: 中文期間 token 剝除 ────────────────────────────────────────────────

def test_normalize_strips_zh_quarter_prefix():
    """'2024年第四季 X' → 'X'（ARLO FY2025Q1 實際輸出樣式）。"""
    raw = {"2024年第四季 Non-GAAP 毛利率": 37.5}
    result = _normalize_nongaap_metrics(raw)
    assert list(result.values()) == [37.5]
    assert "2024" not in list(result)[0]

def test_normalize_strips_zh_quarter_prefix_with_du():
    """'2025年第四季度 X' — 帶「度」字的變體（ARLO FY2026Q1 樣式）。"""
    raw = {"2025年第四季度 Non-GAAP 毛利率": 47.8}
    assert "2025" not in list(_normalize_nongaap_metrics(raw))[0]

def test_normalize_strips_zh_fiscal_quarter():
    """'2026財年第三季度 X' — PANW 樣式。"""
    raw = {"2026財年第三季度 Non-GAAP 營業利潤": 814000000.0}
    assert "2026" not in list(_normalize_nongaap_metrics(raw))[0]

def test_normalize_strips_zh_bare_quarter():
    """'第一季 X' — 無年份的裸季度（CRM 樣式）。"""
    raw = {"第一季 Non-GAAP 攤薄每股盈餘": 3.88}
    assert "第一季" not in list(_normalize_nongaap_metrics(raw))[0]

def test_normalize_zh_annual_goes_to_fy_bucket():
    """'2024全年度 X' 必須歸 FY 桶——否則年度值會蓋掉當季值。"""
    raw = {
        "2024年第四季 Non-GAAP 毛利率": 37.5,   # 當季
        "2024全年度 Non-GAAP 毛利率": 37.6,     # 年度，另成 (FY) 列
    }
    result = _normalize_nongaap_metrics(raw)
    assert result["Non-GAAP 毛利率"] == 37.5
    assert result["Non-GAAP 毛利率 (FY)"] == 37.6

def test_normalize_zh_annual_with_nian_goes_to_fy_bucket():
    """'2025年全年度 X' — 帶「年」的年度變體（ARLO FY2026Q1 樣式）。"""
    raw = {
        "2025年第四季度 Non-GAAP 毛利率": 47.8,
        "2025年全年度 Non-GAAP 毛利率": 45.1,
    }
    result = _normalize_nongaap_metrics(raw)
    assert result["Non-GAAP 毛利率"] == 47.8
    assert result["Non-GAAP 毛利率 (FY)"] == 45.1

def test_normalize_zh_comparison_quarters_deduplicated():
    """同一指標的當季 + 上季 + 去年同季，只留第一個（新聞稿當季在前）。"""
    raw = {
        "2024年第四季 Non-GAAP 毛利率": 37.5,
        "2024年第三季 Non-GAAP 毛利率": 36.0,
        "2023年第四季 Non-GAAP 毛利率": 35.8,
    }
    assert list(_normalize_nongaap_metrics(raw).values()) == [37.5]

# ── L1: 中文 guidance 過濾 ─────────────────────────────────────────────────

def test_normalize_drops_zh_guidance_yuce():
    """'2025年第一季預測 X（低標）' — 預測 = guidance，丟。"""
    raw = {"2025年第一季預測 Non-GAAP 每股盈餘（低標）": 0.09}
    assert _normalize_nongaap_metrics(raw) == {}

def test_normalize_drops_zh_guidance_zhiyin():
    """'... 指引下限' — 指引 = guidance，且詞在名稱中間，不是開頭。"""
    raw = {"2026年第一季度 Non-GAAP 稀釋每股收益指引下限": 0.17}
    assert _normalize_nongaap_metrics(raw) == {}

def test_normalize_drops_zh_guidance_in_middle():
    """中文 guidance 詞出現在中間也要丟（英文版用 startswith 會漏）。"""
    raw = {"2026財年預期 Non-GAAP 營業利潤率上限": 29.0}
    assert _normalize_nongaap_metrics(raw) == {}

def test_normalize_keeps_zh_metric_without_guidance_word():
    """不含 guidance 詞的中文指標必須留下——防過度過濾。"""
    raw = {"Non-GAAP 毛利率": 50.1}
    assert len(_normalize_nongaap_metrics(raw)) == 1

# ── L2: 中英對照 ───────────────────────────────────────────────────────────

def test_canonicalize_zh_gross_margin():
    from fetcher_nongaap import _canonicalize_metric_name
    assert _canonicalize_metric_name("Non-GAAP 毛利率") == "Non-GAAP Gross Margin"

def test_canonicalize_zh_free_cash_flow():
    from fetcher_nongaap import _canonicalize_metric_name
    assert _canonicalize_metric_name("自由現金流") == "Free Cash Flow"

def test_canonicalize_composes_terms():
    """詞彙替換要能組合：自由現金流 + 利潤率 → Free Cash Flow Margin。"""
    from fetcher_nongaap import _canonicalize_metric_name
    assert _canonicalize_metric_name("自由現金流利潤率") == "Free Cash Flow Margin"

def test_canonicalize_longest_term_wins():
    """訂閱與服務毛利率 要整段命中，不可被拆成錯誤組合。"""
    from fetcher_nongaap import _canonicalize_metric_name
    assert _canonicalize_metric_name("Non-GAAP 訂閱與服務毛利率") == \
        "Non-GAAP Subscription and Services Gross Margin"

def test_canonicalize_zh_adjusted_ebitda():
    from fetcher_nongaap import _canonicalize_metric_name
    assert _canonicalize_metric_name("調整後 EBITDA 利潤率") == "Adjusted EBITDA Margin"

def test_canonicalize_english_alias_merges():
    """AI 的英文用詞本身也不一致：Net Income Per Share 與 稀釋每股收益 是同一列。"""
    from fetcher_nongaap import _canonicalize_metric_name
    a = _canonicalize_metric_name("Non-GAAP Net Income Per Share")
    b = _canonicalize_metric_name("Non-GAAP 稀釋每股收益")
    assert a == b

def test_canonicalize_unknown_name_passes_through():
    """對照表不可能窮舉——沒收錄的名稱要原樣留下，不可吞掉資料。"""
    from fetcher_nongaap import _canonicalize_metric_name
    assert _canonicalize_metric_name("Non-GAAP 某個沒收錄的指標") == \
        "Non-GAAP 某個沒收錄的指標"

def test_canonicalize_pure_english_unchanged():
    """純英文且已是標準名的，不可被改寫（避免動到既有 NVDA 行為）。"""
    from fetcher_nongaap import _canonicalize_metric_name
    assert _canonicalize_metric_name("Non-GAAP Revenue") == "Non-GAAP Revenue"

# ── L3: 表格組裝——對角線散開的直接斷言 ────────────────────────────────────

def test_build_nongaap_table_merges_zh_and_en_same_metric():
    """同一指標 Q1 回中文、Q2 回英文，必須合成一列兩格有值，不可散成兩列。"""
    cache = {
        "FY2025Q1": {"filing_date": "2025-02-01", "metrics": {"Non-GAAP 毛利率": 37.5}},
        "FY2025Q2": {"filing_date": "2025-05-01", "metrics": {"Non-GAAP Gross Margin": 41.4}},
    }
    tbl = _build_nongaap_table("ARLO", cache)
    assert _ng_row(tbl, "Non-GAAP Gross Margin") == [37.5, 41.4]

def test_build_nongaap_table_merges_case_variants():
    """AI 的英文大小寫也會漂移：Gross margin / Gross Margin 是同一列。"""
    cache = {
        "FY2025Q1": {"filing_date": "", "metrics": {"Non-GAAP Gross margin": 37.5}},
        "FY2025Q2": {"filing_date": "", "metrics": {"Non-GAAP Gross Margin": 41.4}},
    }
    tbl = _build_nongaap_table("ARLO", cache)
    assert _ng_row(tbl, "Non-GAAP Gross Margin") == [37.5, 41.4]
    assert tbl.concepts.count("Non-GAAP Gross Margin") == 1

def test_build_nongaap_table_renormalizes_legacy_cache():
    """舊快取存的是「未剝期間」的中文名（正規化以前只在寫入時做）。
    讀取時要再跑一次正規化，否則既有 nongaap_cache.json 全部要重抓。"""
    cache = {
        "FY2025Q1": {"filing_date": "", "metrics": {"2024年第四季 Non-GAAP 毛利率": 37.5}},
        "FY2025Q2": {"filing_date": "", "metrics": {"Non-GAAP 毛利率": 41.4}},
    }
    tbl = _build_nongaap_table("ARLO", cache)
    assert _ng_row(tbl, "Non-GAAP Gross Margin") == [37.5, 41.4]

def test_build_nongaap_table_drops_guidance_from_legacy_cache():
    """舊快取裡的中文 guidance 列，讀取時要被濾掉。"""
    cache = {
        "FY2025Q1": {"filing_date": "", "metrics": {
            "Non-GAAP 毛利率": 37.5,
            "2025年第一季預測 Non-GAAP 每股盈餘（低標）": 0.09,
        }},
    }
    tbl = _build_nongaap_table("ARLO", cache)
    assert not any("預測" in c or "低標" in c for c in tbl.concepts)

# ── L5: ARLO 真實快取 golden test（不連網、不呼叫 AI）───────────────────────

def _load_arlo_fixture():
    p = Path(__file__).parent / "fixtures" / "arlo_nongaap_raw.json"
    with open(p, encoding="utf-8") as f:
        return json.load(f)

def test_arlo_golden_metric_count_collapses():
    """修前 6 季共 59 個原始名稱、幾乎每個各自成列；修後當季列應收斂到 10 列以內。
    (FY) 列不算在內——那是新聞稿裡確實存在的年度數字，只是被標記出來另放。"""
    tbl = _build_nongaap_table("ARLO", _load_arlo_fixture())
    # core 與調節列固定存在，這裡看的是 overflow 區有沒有收斂：
    # 修前 6 季共 59 個原始名稱，幾乎每個各自成列。
    from nongaap_layout import SECTION_OTHER, SECTION_ANNUAL
    if SECTION_OTHER in tbl.concepts:
        start = tbl.concepts.index(SECTION_OTHER)
        end = (tbl.concepts.index(SECTION_ANNUAL)
               if SECTION_ANNUAL in tbl.concepts else len(tbl.concepts))
        assert len(tbl.concepts[start + 1:end]) <= 10

def test_arlo_golden_core_metrics_are_dense():
    """毛利率是 ARLO 每季都報的指標，6 季必須都有值。"""
    tbl = _build_nongaap_table("ARLO", _load_arlo_fixture())
    by_name = dict(zip(tbl.concepts, tbl.values))
    gm = by_name["Non-GAAP Gross Margin"]
    assert sum(v is not None for v in gm) == 6

def test_arlo_golden_eps_dense():
    tbl = _build_nongaap_table("ARLO", _load_arlo_fixture())
    by_name = dict(zip(tbl.concepts, tbl.values))
    eps = by_name["Non-GAAP Diluted EPS"]
    assert sum(v is not None for v in eps) == 6

def test_arlo_golden_no_guidance_rows():
    """預測／指引列全部不得出現。"""
    tbl = _build_nongaap_table("ARLO", _load_arlo_fixture())
    joined = " ".join(tbl.concepts)
    for word in ("預測", "指引", "低標", "高標", "Guidance", "Outlook"):
        assert word not in joined

def test_arlo_golden_no_period_tokens_left():
    """任何指標名都不得殘留年份或季度字樣。"""
    tbl = _build_nongaap_table("ARLO", _load_arlo_fixture())
    joined = " ".join(tbl.concepts)
    for token in ("2023", "2024", "2025", "2026", "第一季", "第三季", "第四季", "全年度"):
        assert token not in joined

def test_arlo_golden_values_in_plausible_range():
    """毛利率落在 30–60、EPS 落在 0–1——防止數值被錯位或縮放。"""
    tbl = _build_nongaap_table("ARLO", _load_arlo_fixture())
    by_name = dict(zip(tbl.concepts, tbl.values))
    for v in by_name["Non-GAAP Gross Margin"]:
        assert v is None or 30.0 <= v <= 60.0
    for v in by_name["Non-GAAP Diluted EPS"]:
        assert v is None or 0.0 <= v <= 1.0

def test_arlo_golden_q2fy2026_matches_press_release():
    """FY2026Q2 欄位對 8-K 原文：毛利率 50.1、EPS 0.28、Adjusted EBITDA 30.4M。"""
    tbl = _build_nongaap_table("ARLO", _load_arlo_fixture())
    col = tbl.quarter_labels.index("FY2026Q2")
    by_name = dict(zip(tbl.concepts, tbl.values))
    assert by_name["Non-GAAP Gross Margin"][col] == 50.1
    assert by_name["Non-GAAP Diluted EPS"][col] == 0.28
    assert by_name["Adjusted EBITDA"][col] == 30400000.0


# ── prompt 要求英文指標名（方案 c 的第一道防線）────────────────────────────

def test_prompt_requires_english_metric_names():
    """AI 回中回英是隨機的，prompt 必須明確要求英文，減少下游要接的中文。
    對照層仍然保留當第二道防線（AI 不聽話時接住）。"""
    from fetcher_nongaap import _NONGAAP_PROMPT
    assert "English" in _NONGAAP_PROMPT

def test_prompt_still_has_text_placeholder():
    """改 prompt 不可弄丟 format 佔位符，否則 _call_ai 會拋 KeyError。"""
    from fetcher_nongaap import _NONGAAP_PROMPT
    assert "{press_release_text}" in _NONGAAP_PROMPT
    _NONGAAP_PROMPT.format(press_release_text="x")   # 不可拋例外


def test_prompt_restricts_to_current_period():
    """實跑 ARLO 發現 Free Cash Flow 列混進全年度數字（FY2025Q1 的 48.6M 配 9.5%
    margin 是年度值，單季應該是 ~37%）。prompt 必須明確限定只取當期。"""
    from fetcher_nongaap import _NONGAAP_PROMPT
    low = _NONGAAP_PROMPT.lower()
    assert "full-year" in low or "full year" in low
    assert "year-to-date" in low


# ── AI 呼叫失敗不可污染快取（2026-08-01 實跑 PANW 撞 HTTP 429 後補）────────
#
# 原行為：_call_ai 失敗回 {} → _extract_nongaap_metrics 回 {} → 該季照樣寫進
# 快取（metrics 為空）→ 下次執行 `lbl not in cache` 命中 → **永遠不會再抓**。
# 一次暫時性的 429 或斷線，就讓那季資料無聲永久消失。
#
# 修法：失敗回 None（與「AI 有回應但沒找到指標」的 {} 區分開），
# 失敗的季度不寫快取，下次執行自動重試。

def test_call_ai_returns_none_on_failure(monkeypatch):
    """AI 呼叫拋例外時要回 None，不可回 {}——{} 代表「真的沒有指標」。"""
    import fetcher_nongaap as fn
    monkeypatch.setattr(fn, "_NONGAAP_PROMPT", "{press_release_text}")
    result = fn._call_ai("text", {"provider": "nonexistent-provider"})
    assert result is None

def test_extract_nongaap_metrics_propagates_none(monkeypatch):
    """_extract_nongaap_metrics 不可把 _call_ai 的 None 吃掉變成 {}。"""
    import fetcher_nongaap as fn
    monkeypatch.setattr(fn, "_call_ai", lambda text, cfg: None)

    class FakePR:
        def markdown(self): return "some press release text"
    class FakeEightK:
        press_releases = [FakePR()]

    assert fn._extract_nongaap_metrics(FakeEightK(), {}) is None

def test_failed_quarter_not_written_to_cache(tmp_path, monkeypatch):
    """AI 失敗的季度不可進快取，否則下次執行會跳過它。"""
    import fetcher_nongaap as fn

    company = FakeCompany([RecoverableFiling("2.02", "2024-09-30", has_earnings=True)])
    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})
    monkeypatch.setattr(fn, "_extract_nongaap_metrics", lambda ek, cfg: None)

    fn.fetch_nongaap_statements("TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path)

    cached = fn._load_cache(tmp_path / fn.CACHE_FILENAME, "TEST")
    assert "FY2024Q3" not in cached

def test_failed_quarter_retried_on_next_run(tmp_path, monkeypatch):
    """第一趟失敗、第二趟成功——第二趟必須真的重抓，而不是讀到空快取。"""
    import fetcher_nongaap as fn

    company = FakeCompany([RecoverableFiling("2.02", "2024-09-30", has_earnings=True)])
    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})

    monkeypatch.setattr(fn, "_extract_nongaap_metrics", lambda ek, cfg: None)
    fn.fetch_nongaap_statements("TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path)

    monkeypatch.setattr(fn, "_extract_nongaap_metrics",
                        lambda ek, cfg: {"Non-GAAP Gross Margin": 45.5})
    fn.fetch_nongaap_statements("TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path)

    cached = fn._load_cache(tmp_path / fn.CACHE_FILENAME, "TEST")
    assert cached["FY2024Q3"]["metrics"]["Non-GAAP Gross Margin"] == 45.5

def test_genuinely_empty_quarter_is_cached(tmp_path, monkeypatch):
    """AI 成功回應但新聞稿真的沒有 Non-GAAP 指標（{}）時，仍要寫快取，
    否則每次執行都會為了同一份沒有指標的 8-K 重複呼叫 AI。"""
    import fetcher_nongaap as fn

    company = FakeCompany([RecoverableFiling("2.02", "2024-09-30", has_earnings=True)])
    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})
    monkeypatch.setattr(fn, "_extract_nongaap_metrics", lambda ek, cfg: {})

    fn.fetch_nongaap_statements("TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path)

    cached = fn._load_cache(tmp_path / fn.CACHE_FILENAME, "TEST")
    assert "FY2024Q3" in cached
    assert cached["FY2024Q3"]["metrics"] == {}

def test_build_nongaap_table_tolerates_none_metrics():
    """防禦：舊快取若存過 None，組表不可炸。"""
    cache = {"FY2025Q1": {"filing_date": "", "metrics": None}}
    tbl = _build_nongaap_table("TEST", cache)
    # 沒有指標時仍回骨架表——讀不到 sheet 與讀到空 sheet 是兩種不同的訊號
    assert tbl is not None
    assert all(v is None for row in tbl.values for v in row)


# ── CRM 實跑暴露的規則表缺口（2026-08-01）────────────────────────────────────

def test_normalize_drops_english_guidance_in_middle():
    """'Non-GAAP Diluted Net Income Per Share Guidance (Low)' —— guidance 在中間。
    英文原本只用 startswith，這類整批漏掉，直接混進時間序列。"""
    raw = {"Non-GAAP Diluted Net Income Per Share Guidance (Low)": 3.11}
    assert _normalize_nongaap_metrics(raw) == {}

def test_normalize_drops_english_guidance_suffix():
    raw = {"Non-GAAP Operating Margin Guidance": 34.3}
    assert _normalize_nongaap_metrics(raw) == {}

def test_normalize_drops_english_high_low_bounds():
    raw = {"Free Cash Flow Growth Guidance (High)": 10.0}
    assert _normalize_nongaap_metrics(raw) == {}

def test_normalize_keeps_non_guidance_english_metric():
    """防過度過濾：一般英文指標不可被誤殺。"""
    raw = {"Non-GAAP Operating Margin": 34.8}
    assert _normalize_nongaap_metrics(raw) == {"Non-GAAP Operating Margin": 34.8}

def test_canonicalize_constant_currency():
    """CRM 的「恆定匯率」與「固定匯率」是同一件事（constant currency）。"""
    from fetcher_nongaap import _canonicalize_metric_name
    a = _canonicalize_metric_name("Non-GAAP 恆定匯率 Revenue 成長率")
    b = _canonicalize_metric_name("固定匯率 Revenue 成長率")
    assert a == b

def test_canonicalize_growth_rate():
    from fetcher_nongaap import _canonicalize_metric_name
    assert "Growth" in _canonicalize_metric_name("Free Cash Flow 年增率")

def test_canonicalize_crpo():
    """當期剩餘履約義務 = cRPO，SaaS 業最重要的領先指標之一。"""
    from fetcher_nongaap import _canonicalize_metric_name
    out = _canonicalize_metric_name("Non-GAAP 恆定匯率當期剩餘履約義務成長率")
    assert "cRPO" in out
    assert "履約" not in out

def test_canonicalize_leaves_no_chinese_for_crm_names():
    """CRM 實跑出現過的中文名，全部不得殘留中文字。"""
    from fetcher_nongaap import _canonicalize_metric_name
    names = [
        "Non-GAAP 恆定匯率 Revenue 成長率",
        "Non-GAAP 恆定匯率當期剩餘履約義務成長率",
        "Non-GAAP 恆定匯率 Subscription 與支援 Revenue 成長率",
        "Free Cash Flow 年增率",
        "固定匯率 Revenue 成長率",
        "固定匯率 Subscription 與支援 Revenue 成長率",
        "固定匯率當期未履約合約總額成長率",
    ]
    for n in names:
        out = _canonicalize_metric_name(n)
        assert not any("一" <= ch <= "鿿" for ch in out), f"{n} -> {out}"

def test_growth_rate_is_percent_in_excel():
    """成長率是百分比，不可除以 1M。"""
    from excel_formatter import _is_percent_concept
    assert _is_percent_concept("Non-GAAP Constant Currency Revenue Growth")


# ═════════════════════════════════════════════════════════════════════════════
# A：年度值不再填進季欄位，改為另成一列加 (FY) 標記（2026-08-01）
#
# 工具的目標是「照實把 8-K 的數字落地」，不做判斷。把全年數字填進季欄位是替
# 資料下判斷（而且無聲）；整個丟掉是刪資料。標記出來另成一列才是照實落地。
# 開關：metric_rules.FY_ONLY_HANDLING = "label" | "fill" | "drop"
# ═════════════════════════════════════════════════════════════════════════════

def test_fy_only_metric_gets_fy_suffix():
    """只有年度值、沒有當季值時，不可佔用季欄位的原名。"""
    raw = {"Non-GAAP Annual Tax Rate (FY26)": 17.0}
    assert _normalize_nongaap_metrics(raw) == {"Non-GAAP Annual Tax Rate (FY)": 17.0}

def test_fy_and_quarterly_coexist_as_separate_rows():
    """當季與年度同時存在時兩列都留，不可互相蓋掉，也不可合併。"""
    raw = {
        "Non-GAAP EPS (Q4 FY26)": 1.62,
        "Non-GAAP EPS (FY26)": 6.01,
    }
    result = _normalize_nongaap_metrics(raw)
    assert result["Non-GAAP EPS"] == 1.62
    assert result["Non-GAAP EPS (FY)"] == 6.01

def test_zh_annual_gets_fy_suffix():
    raw = {"2024全年度 Non-GAAP 毛利率": 37.6}
    assert _normalize_nongaap_metrics(raw) == {"Non-GAAP 毛利率 (FY)": 37.6}

def test_quarterly_value_never_replaced_by_annual():
    """最重要的一條：當季有值時，該格必須是當季的數字。"""
    raw = {
        "2024年第四季 Non-GAAP 毛利率": 37.5,
        "2024全年度 Non-GAAP 毛利率": 37.6,
    }
    assert _normalize_nongaap_metrics(raw)["Non-GAAP 毛利率"] == 37.5

def test_canonicalize_preserves_fy_suffix():
    """(FY) 標記不可被中英對照層吃掉，也不可害對照表比對不到。"""
    from fetcher_nongaap import _canonicalize_metric_name
    assert _canonicalize_metric_name("Non-GAAP 毛利率 (FY)") == "Non-GAAP Gross Margin (FY)"

def test_fy_row_does_not_merge_with_quarterly_row():
    """(FY) 列與當季列在組表時必須是兩列。"""
    cache = {
        "FY2025Q1": {"filing_date": "", "metrics": {
            "Non-GAAP Gross Margin": 37.5,
            "Non-GAAP Gross Margin (FY)": 37.6,
        }},
    }
    tbl = _build_nongaap_table("TEST", cache)
    assert _ng_row(tbl, "Non-GAAP Gross Margin") == [37.5]
    assert _ng_row(tbl, "Non-GAAP Gross Margin (FY)") == [37.6]


# ═════════════════════════════════════════════════════════════════════════════
# C：服務毛利率 與 訂閱與服務毛利率 不再合併（2026-08-01）
#
# 查 ARLO 原文：FY2025Q1 報的是 "non-GAAP service gross margin 81.7%"（配
# Service revenue $64.1M），FY2025Q2 起改成 "subscriptions and services gross
# margin 83.1%"（配 Subscriptions and services revenue $68.8M）。公司自己改了
# 名稱與營收基礎，認定兩者是同一條線屬於判斷，工具不做。
# ═════════════════════════════════════════════════════════════════════════════

def test_service_and_subscription_margins_stay_separate():
    from fetcher_nongaap import _canonicalize_metric_name
    a = _canonicalize_metric_name("Non-GAAP Service Gross Margin")
    b = _canonicalize_metric_name("Non-GAAP Subscriptions and Services Gross Margin")
    assert a != b

def test_english_subscription_variants_still_merge():
    """公司同一個說法的單複數／連接詞差異仍要併——那是寫法不是定義。"""
    from fetcher_nongaap import _canonicalize_metric_name
    a = _canonicalize_metric_name("Non-GAAP Subscriptions and Services Gross Margin")
    b = _canonicalize_metric_name("Non-GAAP Subscription and Services Gross Margin")
    assert a == b


# ═════════════════════════════════════════════════════════════════════════════
# D：AI 呼叫退避重試 + 跑完統計未取得季數（2026-08-01）
# ═════════════════════════════════════════════════════════════════════════════

def test_call_ai_retries_then_succeeds(monkeypatch):
    """暫時性失敗（429 尖峰）要能靠重試救回來。"""
    import fetcher_nongaap as fn
    calls = []
    def flaky(prompt, cfg):
        calls.append(1)
        if len(calls) < 3:
            raise RuntimeError("boom")
        return '{"Non-GAAP Gross Margin": 45.5}'
    monkeypatch.setattr(fn, "_ai_request", flaky)
    monkeypatch.setattr(fn.time, "sleep", lambda s: None)

    assert fn._call_ai("text", {"provider": "google"}) == {"Non-GAAP Gross Margin": 45.5}
    assert len(calls) == 3

def test_call_ai_gives_up_after_max_attempts(monkeypatch):
    import fetcher_nongaap as fn
    calls = []
    def always_fail(prompt, cfg):
        calls.append(1)
        raise RuntimeError("boom")
    monkeypatch.setattr(fn, "_ai_request", always_fail)
    monkeypatch.setattr(fn.time, "sleep", lambda s: None)

    assert fn._call_ai("text", {"provider": "google"}) is None
    assert len(calls) == fn.AI_MAX_ATTEMPTS

def test_call_ai_does_not_retry_on_success(monkeypatch):
    """成功就不要多打——AI 呼叫是要付費的。"""
    import fetcher_nongaap as fn
    calls = []
    def ok(prompt, cfg):
        calls.append(1)
        return "{}"
    monkeypatch.setattr(fn, "_ai_request", ok)
    monkeypatch.setattr(fn.time, "sleep", lambda s: None)

    assert fn._call_ai("text", {"provider": "google"}) == {}
    assert len(calls) == 1

def test_retry_backoff_is_applied(monkeypatch):
    """重試之間要真的等，否則對每分鐘限流沒有意義。"""
    import fetcher_nongaap as fn
    slept = []
    monkeypatch.setattr(fn, "_ai_request",
                        lambda p, c: (_ for _ in ()).throw(RuntimeError("boom")))
    monkeypatch.setattr(fn.time, "sleep", lambda s: slept.append(s))

    fn._call_ai("text", {"provider": "google"})
    assert len(slept) == fn.AI_MAX_ATTEMPTS - 1
    assert all(s > 0 for s in slept)

def test_failed_quarters_summarised_at_end(tmp_path, monkeypatch, capsys):
    """跑完要明確說「N 季沒拿到」，不能只有中途一行 stderr 捲過去。"""
    import fetcher_nongaap as fn

    company = FakeCompany([
        RecoverableFiling("2.02", "2024-09-30", has_earnings=True),
        RecoverableFiling("2.02", "2024-06-30", has_earnings=True),
    ])
    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})
    monkeypatch.setattr(fn, "_extract_nongaap_metrics", lambda ek, cfg: None)

    fn.fetch_nongaap_statements("TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path)

    err = capsys.readouterr().err
    assert "2" in err and "FY2024Q3" in err and "FY2024Q2" in err

def test_summary_reaches_progress_callback(tmp_path, monkeypatch):
    """GUI 只看得到 progress_cb，stderr 使用者看不到。"""
    import fetcher_nongaap as fn

    company = FakeCompany([RecoverableFiling("2.02", "2024-09-30", has_earnings=True)])
    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})
    monkeypatch.setattr(fn, "_extract_nongaap_metrics", lambda ek, cfg: None)

    msgs = []
    fn.fetch_nongaap_statements("TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path,
                                progress_cb=lambda i, t, m: msgs.append(m))
    assert any("重跑" in m or "未取得" in m for m in msgs)

def test_no_summary_when_everything_succeeds(tmp_path, monkeypatch, capsys):
    """全部成功時不可印雜訊。"""
    import fetcher_nongaap as fn

    company = FakeCompany([RecoverableFiling("2.02", "2024-09-30", has_earnings=True)])
    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})
    monkeypatch.setattr(fn, "_extract_nongaap_metrics", lambda ek, cfg: {"Non-GAAP Revenue": 1.0})

    fn.fetch_nongaap_statements("TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path)
    assert "未取得" not in capsys.readouterr().err


# ═════════════════════════════════════════════════════════════════════════════
# max_filings 在缺季回補後要重新裁切（TODO 第 4 項，2026-08-01）
#
# 原行為：_list_earnings_filings() 先套 max_filings 切片，_recover_missing_quarters()
# 補回缺季後**不再裁切**，所以要 4 季、保留區間有 2 個缺口時實際會下載 6 份，
# 每多一份就多一次 AI 呼叫（也就是多一次配額與費用）。
# ═════════════════════════════════════════════════════════════════════════════

def test_recovered_quarters_respect_max_filings(tmp_path, monkeypatch):
    """回補後總數不可超過 max_filings。"""
    import fetcher_nongaap as fn

    # 6 份財報，其中兩份沒標 2.02（清單階段會漏，靠回補找回）
    filings = [
        RecoverableFiling("2.02", "2024-12-31", has_earnings=True),
        RecoverableFiling("5.07", "2024-09-30", has_earnings=True),   # 缺口
        RecoverableFiling("2.02", "2024-06-30", has_earnings=True),
        RecoverableFiling("5.07", "2024-03-31", has_earnings=True),   # 缺口
        RecoverableFiling("2.02", "2023-12-31", has_earnings=True),
        RecoverableFiling("2.02", "2023-09-30", has_earnings=True),
    ]
    company = FakeCompany(filings)
    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})
    monkeypatch.setattr(fn, "_extract_nongaap_metrics", lambda ek, cfg: {"Non-GAAP Revenue": 1.0})

    seen = []
    fn.fetch_nongaap_statements(
        "TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path,
        progress_cb=lambda i, t, label: seen.append(label),
        max_filings=4,
    )

    cached = fn._load_cache(tmp_path / fn.CACHE_FILENAME, "TEST")
    assert len(cached) <= 4


def test_recovered_quarters_keep_the_newest(tmp_path, monkeypatch):
    """裁切要留最新的——分析看的是近幾季，不是十年前那幾季。"""
    import fetcher_nongaap as fn

    filings = [
        RecoverableFiling("2.02", "2024-12-31", has_earnings=True),
        RecoverableFiling("5.07", "2024-09-30", has_earnings=True),
        RecoverableFiling("2.02", "2024-06-30", has_earnings=True),
        RecoverableFiling("2.02", "2024-03-31", has_earnings=True),
        RecoverableFiling("2.02", "2023-12-31", has_earnings=True),
    ]
    company = FakeCompany(filings)
    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})
    monkeypatch.setattr(fn, "_extract_nongaap_metrics", lambda ek, cfg: {"Non-GAAP Revenue": 1.0})

    fn.fetch_nongaap_statements("TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path,
                                max_filings=3)

    cached = fn._load_cache(tmp_path / fn.CACHE_FILENAME, "TEST")
    assert set(cached) == {"FY2024Q4", "FY2024Q3", "FY2024Q2"}


def test_recovery_still_fills_gap_within_limit(tmp_path, monkeypatch):
    """裁切不可把回補功能整個廢掉——額度內的缺季仍要補回來。"""
    import fetcher_nongaap as fn

    filings = [
        RecoverableFiling("2.02", "2024-12-31", has_earnings=True),
        RecoverableFiling("5.07", "2024-09-30", has_earnings=True),   # 缺口，要補回
        RecoverableFiling("2.02", "2024-06-30", has_earnings=True),
    ]
    company = FakeCompany(filings)
    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})
    monkeypatch.setattr(fn, "_extract_nongaap_metrics", lambda ek, cfg: {"Non-GAAP Revenue": 1.0})

    fn.fetch_nongaap_statements("TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path,
                                max_filings=8)

    cached = fn._load_cache(tmp_path / fn.CACHE_FILENAME, "TEST")
    assert "FY2024Q3" in cached


def test_cached_quarters_do_not_consume_the_limit(tmp_path, monkeypatch):
    """max_filings 限的是「這一趟要處理幾季」。已在快取裡的季不該擠掉新的季，
    否則第二次執行會什麼都抓不到。"""
    import fetcher_nongaap as fn

    filings = [
        RecoverableFiling("2.02", "2024-12-31", has_earnings=True),
        RecoverableFiling("2.02", "2024-09-30", has_earnings=True),
    ]
    company = FakeCompany(filings)
    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})
    monkeypatch.setattr(fn, "_extract_nongaap_metrics", lambda ek, cfg: {"Non-GAAP Revenue": 1.0})

    fn.fetch_nongaap_statements("TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path,
                                max_filings=2)
    cached = fn._load_cache(tmp_path / fn.CACHE_FILENAME, "TEST")
    assert set(cached) == {"FY2024Q4", "FY2024Q3"}


# ═════════════════════════════════════════════════════════════════════════════
# 新聞稿截斷（2026-08-02 實跑 ARLO 抓到，調節表全空的真正原因）
#
# prompt 原本只送新聞稿前 12,000 字元。ARLO 的新聞稿全長 53,569 字元，
# 「Stock-based compensation」出現在 18,605 / 33,759 / 38,440 / 40,558，
# 「Amortization」在 40,848——**全部在截斷之後，AI 根本沒看到調節表**。
# 重點條列在最前面（所以毛利率、EPS 抓得到），調節表一律在文件尾端。
# ═════════════════════════════════════════════════════════════════════════════

def test_prompt_text_limit_covers_a_typical_press_release():
    """ARLO 53.6K、多數新聞稿在 30~80K 字元之間，上限要蓋得住。"""
    from fetcher_nongaap import PROMPT_TEXT_LIMIT
    assert PROMPT_TEXT_LIMIT >= 60_000


def test_full_text_reaches_the_model(monkeypatch):
    """截斷點之後的內容必須真的送進 prompt——這是調節表能不能被抓到的關鍵。"""
    import fetcher_nongaap as fn
    seen = {}
    def capture(prompt, cfg):
        seen["prompt"] = prompt
        return "{}"
    monkeypatch.setattr(fn, "_ai_request", capture)
    monkeypatch.setattr(fn.time, "sleep", lambda s: None)

    text = ("x" * 40_000) + "Stock-based compensation 12345"
    fn._call_ai(text, {"provider": "google"})
    assert "Stock-based compensation 12345" in seen["prompt"]


def test_oversized_text_is_still_bounded(monkeypatch):
    """仍要有上限——不可把超長文件整份送出去。"""
    import fetcher_nongaap as fn
    seen = {}
    monkeypatch.setattr(fn, "_ai_request",
                        lambda p, c: (seen.__setitem__("prompt", p), "{}")[1])
    monkeypatch.setattr(fn.time, "sleep", lambda s: None)

    fn._call_ai("y" * (fn.PROMPT_TEXT_LIMIT + 50_000), {"provider": "google"})
    assert len(seen["prompt"]) < fn.PROMPT_TEXT_LIMIT + 5_000


# ═════════════════════════════════════════════════════════════════════════════
# 額度用盡後停止重試（2026-08-03）
#
# Gemini 的額度是**按請求次數**算的。撞到 429（額度用盡）之後再重試，每一次都
# 必敗、而且每一次都扣一次額度。實測 CRM 一趟燒掉 12 次呼叫換到 0 筆資料
# （4 季 × 3 次嘗試）。熔斷後只會花 3 次，省下的 9 次可以換 9 季真實資料。
#
# 注意：只對「額度型」失敗熔斷。一般的暫時性錯誤（連線中斷等）仍要重試。
# ═════════════════════════════════════════════════════════════════════════════

class _Quota429(Exception):
    """模擬 SDK 的額度耗盡例外（帶 HTTP 429）。"""
    status_code = 429


def test_no_retry_after_quota_exhausted(monkeypatch):
    """429 不重試——重試也是必敗，而且每次都扣額度。"""
    import fetcher_nongaap as fn
    calls = []
    def quota_fail(prompt, cfg):
        calls.append(1)
        raise _Quota429("quota")
    monkeypatch.setattr(fn, "_ai_request", quota_fail)
    monkeypatch.setattr(fn.time, "sleep", lambda s: None)

    assert fn._call_ai("text", {"provider": "google"}) is None
    assert len(calls) == 1


def test_transient_failure_still_retries(monkeypatch):
    """非額度型的失敗仍要重試——那種重試是有機會成功的。"""
    import fetcher_nongaap as fn
    calls = []
    def flaky(prompt, cfg):
        calls.append(1)
        if len(calls) < 2:
            raise RuntimeError("connection reset")
        return '{"Non-GAAP Revenue": 1.0}'
    monkeypatch.setattr(fn, "_ai_request", flaky)
    monkeypatch.setattr(fn.time, "sleep", lambda s: None)

    assert fn._call_ai("text", {"provider": "google"}) == {"Non-GAAP Revenue": 1.0}
    assert len(calls) == 2


def test_quota_exhaustion_stops_remaining_quarters(tmp_path, monkeypatch):
    """一季撞到額度後，同一趟剩下的季不再呼叫 AI——否則每季再燒一次。"""
    import fetcher_nongaap as fn

    company = FakeCompany([
        RecoverableFiling("2.02", "2024-12-31", has_earnings=True),
        RecoverableFiling("2.02", "2024-09-30", has_earnings=True),
        RecoverableFiling("2.02", "2024-06-30", has_earnings=True),
    ])
    monkeypatch.setattr(fn, "set_identity", lambda *a, **k: None)
    monkeypatch.setattr(fn, "Company", lambda ticker: company)
    monkeypatch.setattr(fn, "_extract_eps_recon", lambda ek: {})

    calls = []
    def quota_fail(ek, cfg):
        calls.append(1)
        raise _Quota429("quota")
    monkeypatch.setattr(fn, "_extract_nongaap_metrics", quota_fail)
    monkeypatch.setattr(fn.time, "sleep", lambda s: None)

    fn.fetch_nongaap_statements("TEST", "CTH x@y.com", {"api_key": "k"}, tmp_path)
    assert len(calls) == 1

    # 失敗的季一律不寫快取，下次執行仍會全部重抓
    assert fn._load_cache(tmp_path / fn.CACHE_FILENAME, "TEST") == {}


def test_is_quota_error_detects_429():
    from fetcher_nongaap import _is_quota_error
    assert _is_quota_error(_Quota429("x")) is True
    assert _is_quota_error(RuntimeError("boom")) is False
