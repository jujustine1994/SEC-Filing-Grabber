# tests/test_fetcher_nongaap.py
import json
from pathlib import Path
from fetcher_nongaap import _load_cache, _save_cache, _period_to_quarter_label, _build_eps_recon_table, _build_nongaap_table


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
    result = _load_cache(Path("/nonexistent/nongaap_cache.json"))
    assert result == {}

def test_save_and_load_cache(tmp_path):
    cache_path = tmp_path / "nongaap_cache.json"
    data = {"FY2024Q1": {"metrics": {"Non-GAAP EPS": 0.71}}}
    _save_cache(cache_path, data)
    loaded = _load_cache(cache_path)
    assert loaded["FY2024Q1"]["metrics"]["Non-GAAP EPS"] == 0.71


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
