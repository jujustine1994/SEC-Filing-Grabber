# tests/test_fetcher_nongaap.py
import json
from pathlib import Path
from fetcher_nongaap import _load_cache, _save_cache, _period_to_quarter_label


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
