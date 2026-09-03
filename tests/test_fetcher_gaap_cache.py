"""本地 filing 快取掛在 `_filing_obj()` 上的整合行為。

`filing_cache.py` 自己的儲存層測試在 tests/test_filing_cache.py。這裡釘的是
「什麼時候會打網路、什麼時候不會」，以及那幾條踩到會餵錯資料的邊界。
"""
from unittest.mock import MagicMock

import pandas as pd
import pytest

import filing_cache
from fetcher_gaap import (
    _filing_obj,
    _bind_disk_cache,
    _disk_cache_scope,
    _parse_cache_scope,
    last_cache_stats,
)
from net_retry import NetworkDownError

ACC = "0001045810-25-000123"
ACC_OLD = "0001045810-19-000001"


@pytest.fixture
def cache_dir(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    return tmp_path / "SEC Financial Tools" / "filing_cache"


def _df():
    return pd.DataFrame({
        "concept": ["us-gaap_Revenue"],
        "label": ["Net sales"],
        "2025-12-27 (Q1)": [1000.0],
    })


def _fake_filing(acc=ACC, filing_date="2025-08-27", form="10-Q", financials=True):
    """一份假的 edgartools filing。`obj()` 被呼叫幾次是這組測試的主要斷言。"""
    stmt = MagicMock()
    stmt.to_dataframe.return_value = _df()
    fin = MagicMock()
    fin.income_statement.return_value = stmt
    fin.balance_sheet.return_value = stmt
    fin.cashflow_statement.return_value = stmt

    obj = MagicMock()
    obj.financials = fin if financials else None

    filing = MagicMock()
    filing.accession_no = acc
    filing.filing_date = filing_date
    filing.form = form
    filing.obj.return_value = obj
    return filing


# ── 命中／未命中 ──────────────────────────────────────────────────────────

def test_first_fetch_hits_the_network_and_writes_the_cache_file(cache_dir):
    filing = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(filing)
    assert filing.obj.call_count == 1
    assert filing_cache.filing_path("NVDA", ACC).exists()
    assert any(f["accession_no"] == ACC
               for f in filing_cache.read_manifest("NVDA")["filings"])


def test_second_fetch_reads_the_cache_and_never_touches_the_network(cache_dir):
    first = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(first)

    second = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        tenq = _filing_obj(second)
        df = tenq.financials.income_statement().to_dataframe()
    assert second.obj.call_count == 0
    pd.testing.assert_frame_equal(df, _df())
    assert last_cache_stats() == (1, 1)


def test_a_cache_hit_still_goes_through_the_in_memory_parse_cache(cache_dir):
    """G9 的記憶體快取存的就是 `_filing_obj()` 的回傳值——存真物件或存替身
    對它沒有差別，兩層快取不衝突。"""
    warm = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(warm)

    filing = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        assert _filing_obj(filing) is _filing_obj(filing)


def test_cache_is_off_when_nothing_is_bound(cache_dir):
    """沒綁定 ticker/cik（例如拿不到 cik）時，行為跟改動前完全一樣。"""
    filing = _fake_filing()
    with _parse_cache_scope():
        _filing_obj(filing)
    assert filing.obj.call_count == 1
    assert not filing_cache.cache_root().exists()


# ── 負向快取 vs 網路失敗 ──────────────────────────────────────────────────

def test_a_filing_with_no_financials_is_cached_negatively(cache_dir):
    """pre-XBRL 舊申報：記著「試過了、沒有財務資料」，下次不用再打 SEC。"""
    first = _fake_filing(acc=ACC_OLD, filing_date="2008-05-01", financials=False)
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(first)

    second = _fake_filing(acc=ACC_OLD, filing_date="2008-05-01", financials=False)
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        tenq = _filing_obj(second)
    assert second.obj.call_count == 0
    assert tenq.financials is None


def test_a_network_failure_is_never_recorded_as_a_negative_cache(cache_dir):
    """網路失敗是暫時性的，交給既有的 D11-B 缺漏帳本，每次都該重試。
    誤記成「沒有 financials」會讓那一期**永久**消失。"""
    boom = _fake_filing()
    boom.obj.side_effect = NetworkDownError("network down")
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        with pytest.raises(NetworkDownError):
            _filing_obj(boom)
    assert not filing_cache.filing_path("NVDA", ACC).exists()

    retry = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(retry)
    assert retry.obj.call_count == 1


def test_a_parse_failure_after_download_is_not_cached_either(cache_dir):
    """`financials` 拿得到但 `to_dataframe()` 炸掉——一樣不留快取，下次重試。"""
    filing = _fake_filing()
    filing.obj.return_value.financials.income_statement.side_effect = RuntimeError("bad xbrl")
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(filing)
    assert not filing_cache.filing_path("NVDA", ACC).exists()


# ── 逐份即時落檔 ──────────────────────────────────────────────────────────

def test_filings_parsed_before_a_crash_stay_on_disk(cache_dir):
    """一趟抓取可能好幾分鐘。中途斷線或使用者關視窗時，已經抓到的進度
    不可以全部白費——所以是逐份落檔，不是整趟跑完才一次寫入。"""
    ok = _fake_filing(acc=ACC)
    boom = _fake_filing(acc=ACC_OLD)
    boom.obj.side_effect = NetworkDownError("network down")

    with pytest.raises(NetworkDownError):
        with _disk_cache_scope(), _parse_cache_scope():
            _bind_disk_cache("NVDA", 1045810)
            _filing_obj(ok)
            _filing_obj(boom)

    assert filing_cache.filing_path("NVDA", ACC).exists()
    assert not filing_cache.filing_path("NVDA", ACC_OLD).exists()


# ── 查詢邊界不能漏抓 ──────────────────────────────────────────────────────

def test_a_wider_query_still_fetches_filings_that_were_never_cached(cache_dir):
    """先用窄範圍抓（只碰到 2020 年以後的 filing），再用完整期間抓同一家，
    第二次一定要真的去補抓從沒進過快取的舊 filing——不能因為「這家公司
    已經有快取」就少抓。快取的粒度是**一份 filing**，不是一家公司。"""
    recent = _fake_filing(acc=ACC, filing_date="2025-08-27")
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(recent)

    recent2 = _fake_filing(acc=ACC, filing_date="2025-08-27")
    old = _fake_filing(acc=ACC_OLD, filing_date="2019-05-01")
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(recent2)
        _filing_obj(old)
    assert recent2.obj.call_count == 0      # 已經有的照樣讀快取
    assert old.obj.call_count == 1          # 沒有的一定要補抓
    assert filing_cache.filing_path("NVDA", ACC_OLD).exists()
    assert last_cache_stats() == (1, 2)


def test_the_cache_is_keyed_per_company_not_shared(cache_dir):
    """ticker 會換手。`cik` 不符時整包視同無快取——這種錯不會報例外，
    只會安靜地把別家公司的數字餵給使用者。"""
    filing = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(filing)

    other = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 99999)     # 同 ticker、不同公司
        _filing_obj(other)
    assert other.obj.call_count == 1
