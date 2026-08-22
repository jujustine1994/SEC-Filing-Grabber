"""跨公司比較「選公司」自動完成比對邏輯（2026-08-21 TODO F4 第 1 項）。

CTH 回報：打公司全名（如「Intel」）搜不到，要打 `INTC` 才行。純函式，
不開 Tk 視窗（沿用 test_window_geometry.py 同一套作法）。
"""
import main

CACHE = {
    "INTC": "Intel Corporation",
    "NVDA": "NVIDIA Corporation",
    "AMD": "Advanced Micro Devices, Inc.",
    "IBM": "International Business Machines Corporation",
}


def test_matches_by_ticker_prefix():
    assert ("INTC", "Intel Corporation") in main.match_company_cache("INT", CACHE)


def test_matches_by_company_name_substring():
    """這是本次要修的 bug：打「Intel」原本搜不到 INTC。"""
    result = main.match_company_cache("Intel", CACHE)
    assert ("INTC", "Intel Corporation") in result


def test_matches_are_case_insensitive():
    assert ("NVDA", "NVIDIA Corporation") in main.match_company_cache("nvidia", CACHE)


def test_ticker_prefix_match_still_works_for_short_names():
    assert ("AMD", "Advanced Micro Devices, Inc.") in main.match_company_cache("AMD", CACHE)


def test_no_match_returns_empty_list():
    assert main.match_company_cache("ZZZZ", CACHE) == []


def test_empty_input_returns_empty_list():
    assert main.match_company_cache("", CACHE) == []
    assert main.match_company_cache("   ", CACHE) == []


def test_respects_limit():
    big_cache = {f"T{i:03d}": f"Test Company {i}" for i in range(20)}
    result = main.match_company_cache("T", big_cache, limit=5)
    assert len(result) == 5


def test_name_substring_match_not_limited_to_prefix():
    """公司名稱裡「含有」輸入字串就算符合，不用是開頭——例如打
    「Business Machines」要找得到 IBM。"""
    result = main.match_company_cache("Business Machines", CACHE)
    assert ("IBM", "International Business Machines Corporation") in result
