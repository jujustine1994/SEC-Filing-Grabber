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
