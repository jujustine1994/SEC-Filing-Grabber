"""本地 filing 快取掛在 `_filing_obj()` 上的整合行為。

`filing_cache.py` 自己的儲存層測試在 tests/test_filing_cache.py。這裡釘的是
「什麼時候會打網路、什麼時候不會」，以及那幾條踩到會餵錯資料的邊界。
"""
import threading
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
        tenq = _filing_obj(filing)
    assert filing.obj.call_count == 1
    # 沒命中磁碟快取（miss）時一定要回傳真物件，不能是 `filing_cache` 的替身——
    # 差一步的話，計畫驗收測試「冷抓 vs 熱快取逐格比對」會失去意義：兩條路
    # 都變成同一種替身物件，比對永遠一樣，測不出快取有沒有餵錯資料。
    assert tenq is filing.obj.return_value
    assert not isinstance(tenq, filing_cache._CachedFiling)
    assert filing_cache.filing_path("NVDA", ACC).exists()


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
    """沒綁定 ticker/cik（例如拿不到 cik）時，行為跟改動前完全一樣。

    真實情況下範圍是有開的（`fetch_gaap_statements()` 一律開
    `_disk_cache_scope()`）——只是 `Company.cik` 拿不到，`_bind_disk_cache()`
    提早回傳、`ctx["ticker"]` 留空。這裡刻意在打開的範圍內綁一個 `cik=None`，
    而不是乾脆不開範圍，才是這條分支實際會被走到的樣子。"""
    filing = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", None)
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
    # 正向對照：同一份 filing 真的抓成功後檔案就會出現——證明前面「沒有
    # 檔案」是因為網路失敗那條路徑真的沒寫，不是因為寫入功能整個是死的
    # （例如 `edgartools_version()` 回 None 導致讀寫全部關閉）。
    assert filing_cache.filing_path("NVDA", ACC).exists()


def test_a_parse_failure_after_download_is_not_cached_either(cache_dir):
    """`financials` 拿得到但 `to_dataframe()` 炸掉——一樣不留快取，下次重試。"""
    filing = _fake_filing()
    filing.obj.return_value.financials.income_statement.side_effect = RuntimeError("bad xbrl")
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(filing)
    assert not filing_cache.filing_path("NVDA", ACC).exists()

    # 正向對照：同一份 filing 換一次正常解析就會落檔——證明前面「沒有檔案」
    # 是解析失敗那條路徑真的沒寫，不是寫入功能整個是死的。
    retry = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(retry)
    assert filing_cache.filing_path("NVDA", ACC).exists()


class _BoomObj:
    """`.obj()` 的回傳值，`financials` 是個存取就炸的 lazy property——用來
    模擬殘缺／異常 XBRL 讓 edgartools 內部丟出 `AttributeError` 的情況。
    刻意不用 `MagicMock` + `PropertyMock`：後者要掛在 `type(instance)` 上，
    而 `MagicMock()` 的類別是全域共用的 `MagicMock`，掛上去會污染其他測試
    用到的所有 mock 物件。"""

    @property
    def financials(self):
        raise AttributeError("boom")


def test_an_attribute_error_reading_financials_is_never_cached_negatively(cache_dir):
    """`financials` 是個會炸的 lazy property（殘缺／異常 XBRL），不是「真的
    沒有 financials」——`getattr(obj, "financials", None)` 會把兩者混為一談，
    寫成負向快取就等於把這一期**永久**判死刑（比網路失敗更糟：網路失敗
    下次還會重試，這種一旦寫錯就再也不會重新嘗試）。分不清楚就不准寫。"""
    filing = _fake_filing()
    filing.obj.return_value = _BoomObj()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(filing)
    assert not filing_cache.filing_path("NVDA", ACC).exists()

    retry = _fake_filing()
    with _disk_cache_scope(), _parse_cache_scope():
        _bind_disk_cache("NVDA", 1045810)
        _filing_obj(retry)
    assert retry.obj.call_count == 1
    assert filing_cache.filing_path("NVDA", ACC).exists()


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


# ── 執行緒隔離 ────────────────────────────────────────────────────────────

def test_disk_cache_is_isolated_across_threads(cache_dir):
    """`main.py` 的批次抓取跟跨公司比較各自跑在自己的執行緒上，且目前並不
    保證兩者不會同時執行（`_compare_worker` 沒有設 `is_running`，不會被
    `_start_worker` 的互斥擋住）。如果磁碟快取範圍是模組級全域變數，兩條
    執行緒會共用同一個 ctx：後綁定的那條執行緒的 ticker/cik 會蓋掉先綁定
    的，檔案就會被寫進錯的公司資料夾。改成 `ContextVar`（跟既有的
    `_ledger_var` 同一招）後，每條新執行緒起手都是空的，天然互不干擾。

    用 `Barrier` 逼兩條執行緒的範圍真的同時存在（而不是先後執行、剛好沒
    撞在一起）——這是刻意設計成必然重疊，不是碰運氣的時序測試，不會 flaky。
    """
    barrier = threading.Barrier(2)
    results: dict[str, int] = {}

    def _run(ticker: str, cik: int, acc: str) -> None:
        filing = _fake_filing(acc=acc)
        with _disk_cache_scope(), _parse_cache_scope():
            _bind_disk_cache(ticker, cik)
            barrier.wait(timeout=5)   # 兩條執行緒的範圍在這一刻確定同時開著
            _filing_obj(filing)
        results[ticker] = filing.obj.call_count

    t1 = threading.Thread(target=_run, args=("NVDA", 1045810, ACC))
    t2 = threading.Thread(target=_run, args=("AAPL", 320193, ACC_OLD))
    t1.start()
    t2.start()
    t1.join(timeout=10)
    t2.join(timeout=10)

    assert not t1.is_alive() and not t2.is_alive()
    assert results == {"NVDA": 1, "AAPL": 1}
    assert filing_cache.filing_path("NVDA", ACC).exists()
    assert filing_cache.filing_path("AAPL", ACC_OLD).exists()
    # 沒有互相污染：誰的 accession 都沒有跑進對方的資料夾。
    assert not filing_cache.filing_path("NVDA", ACC_OLD).exists()
    assert not filing_cache.filing_path("AAPL", ACC).exists()


def test_last_cache_stats_is_isolated_across_threads(cache_dir):
    """驗證每條執行緒的 `last_cache_stats()` 只回自己的命中數/總數，不會撈到
    另一條執行緒的。原本 `_last_cache_stats` 是純模組級全域變數，並行抓取時
    會互相競爭、回報成別家公司的數字。改成 `ContextVar` 解決這個問題。

    用 `Barrier` 逼兩條執行緒的範圍真的同時存在，不是碰運氣的時序測試。
    """
    barrier = threading.Barrier(2)
    results: dict[str, tuple[int, int]] = {}

    def _run(ticker: str, cik: int, hits: int, misses: int) -> None:
        with _disk_cache_scope() as ctx:
            ctx["ticker"] = ticker
            ctx["cik"] = cik
            ctx["hits"] = hits
            ctx["misses"] = misses
            barrier.wait(timeout=5)   # 兩條執行緒的範圍在這一刻確定同時開著
        # 範圍離開後，統計數字才會寫進 ContextVar
        results[ticker] = last_cache_stats()

    t1 = threading.Thread(target=_run, args=("NVDA", 1045810, 24, 1))
    t2 = threading.Thread(target=_run, args=("AAPL", 320193, 5, 10))
    t1.start()
    t2.start()
    t1.join(timeout=10)
    t2.join(timeout=10)

    assert not t1.is_alive() and not t2.is_alive()
    # 各自的統計數字不能互相污染
    assert results == {
        "NVDA": (24, 25),  # hits=24, total=hits+misses=24+1=25
        "AAPL": (5, 15),   # hits=5, total=5+10=15
    }
