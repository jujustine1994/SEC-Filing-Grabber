# tests/test_fetcher_nongaap.py
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
    raw = {"Non-GAAP Net income (FY26)": 4000000000.0}
    assert _normalize_nongaap_metrics(raw) == {"Non-GAAP Net income": 4000000000.0}

def test_normalize_quarterly_wins_over_fy():
    """When Q and FY versions of same metric exist, quarterly value is kept."""
    raw = {
        "Non-GAAP EPS (Q4 FY26)": 1.62,
        "Non-GAAP EPS (FY26)": 6.01,
    }
    result = _normalize_nongaap_metrics(raw)
    assert result == {"Non-GAAP EPS": 1.62}

def test_normalize_fy_fills_gap_when_no_quarterly():
    """FY value is kept when no quarterly version of the metric exists."""
    raw = {"Non-GAAP Annual Tax Rate (FY26)": 17.0}
    assert _normalize_nongaap_metrics(raw) == {"Non-GAAP Annual Tax Rate": 17.0}

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
    assert _normalize_nongaap_metrics({"Non-GAAP EPS (FY26)": 1.0}) == {"Non-GAAP EPS": 1.0}
    assert _normalize_nongaap_metrics({"Non-GAAP EPS (FY2026)": 1.0}) == {"Non-GAAP EPS": 1.0}

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
    assert set(result.keys()) == {
        "Non-GAAP Gross margin",
        "Non-GAAP Net income",
        "Non-GAAP Revenue",
    }
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


import pytest
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


def test_list_earnings_filings_dedupes_keeping_oldest():
    """Same quarter filed twice (e.g. 8-K then corrected 8-K): keep the oldest filing."""
    oldest = FakeFiling("2.02", "2024-09-30", accession="OLDEST")
    newer  = FakeFiling("2.02", "2024-09-30", accession="NEWER")
    company = FakeCompany([newer, oldest])   # newest-first, as EDGAR returns
    result = _list_earnings_filings(company)
    assert len(result) == 1
    assert result[0][1].accession_no == "OLDEST"


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
