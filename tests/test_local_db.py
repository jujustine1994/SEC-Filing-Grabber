"""Tests for local_db.py — 本地財報資料庫的狀態層（TODO J1–J4）。

三塊互相獨立的邏輯，這裡分三段釘：

1. **`_meta.json`**（J2）——寫進去讀回來、跟目錄對不上就重建、schema 不符就重建。
   原則是「掃目錄為準，meta 只是快照」，所以每一條測試都在驗「meta 說謊時
   誰贏」。
2. **`reached_bottom` 推導**（J2）——純函式，餵假的 filing 清單。這是「整家跳過」
   的判斷依據，錯了會變成每次都全部重抓（或反過來，永遠不再往下挖）。
3. **更新名單**（J1）與**版本不符偵測**（J4）——config 讀寫，不碰網路。

抓取流程（J3）用注入的假 `list_filings`／`fetch` 測，完全離線。
"""
import json
from datetime import date
from pathlib import Path

import pytest

import filing_cache
import local_db


@pytest.fixture
def cache_dir(tmp_path, monkeypatch):
    """把快取根目錄導到 tmp_path（跟 test_filing_cache.py 同一招）。"""
    monkeypatch.setenv("SEC_LOCAL_DB_ROOT", str(tmp_path))
    return tmp_path / "filing_cache"


def _acc(n: int) -> str:
    """第 n 份的假 accession，格式合法（`ACCESSION_RE` 會擋不合法的）。"""
    return f"0000320193-{n // 100 % 100:02d}-{n % 100:06d}"


def _write_filing(ticker: str, accession: str, *, form: str, filing_date: str,
                  version: str = "5.29.0", cik: int = 320193) -> Path:
    """直接寫一份 filing 快取檔。不走 `save_filing()`——那個會用「現在安裝的」
    edgartools 版本，測版本不符時就假不出來了。"""
    path = filing_cache.ticker_dir(ticker) / f"{accession}.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({
        "schema_version": filing_cache.SCHEMA_VERSION,
        "accession_no": accession,
        "form": form,
        "filing_date": filing_date,
        "cached_at": "2026-09-05T00:00:00+08:00",
        "cik": cik,
        "edgartools_version": version,
        "has_financials": True,
        "dataframes": {k: None for k in filing_cache.STATEMENT_KEYS},
    }, ensure_ascii=False), encoding="utf-8")
    return path


# ── J2：reached_bottom 推導（純函式）──────────────────────────────────────

def test_reached_bottom_is_none_when_something_is_still_unfetched():
    available = [("a", "2025-01-01"), ("b", "2024-01-01")]
    assert local_db.derive_reached_bottom(available, {"a"}) is None


def test_reached_bottom_is_no_more_filings_when_the_list_runs_out():
    """META 2013 才上市——清單裡沒有任何 2008 之前的申報，抓完就是真的沒了。"""
    available = [("a", "2013-02-01"), ("b", "2014-02-01")]
    assert local_db.derive_reached_bottom(available, {"a", "b"}) == "no_more_filings"


def test_reached_bottom_is_xbrl_cutoff_when_older_filings_exist_but_are_pre_xbrl():
    """AAPL 有 2008 之前的申報，但那些沒有 XBRL，抓不到也不該一直重試。"""
    available = [("a", "2008-05-01"), ("old", "2007-11-01")]
    assert local_db.derive_reached_bottom(available, {"a"}) == "xbrl_cutoff"


def test_reached_bottom_ignores_extra_cached_accessions():
    """快取裡有清單上沒有的（改用不同表單類型、或 SEC 撤下），不影響判定。"""
    available = [("a", "2013-02-01")]
    assert local_db.derive_reached_bottom(available, {"a", "zz"}) == "no_more_filings"


def test_reached_bottom_treats_unparseable_dates_as_inside_the_window():
    """日期壞掉時寧可判「還沒到底」也不要誤判到底——誤判到底會讓那家公司
    再也不往下挖，而且不會有任何症狀。"""
    available = [("a", ""), ("b", "2013-02-01")]
    assert local_db.derive_reached_bottom(available, {"b"}) is None
    assert local_db.derive_reached_bottom(available, {"a", "b"}) == "no_more_filings"


def test_cutoff_matches_the_one_the_fetch_loop_actually_uses():
    """這個常數在兩個檔案各寫一份（local_db 不想 import edgartools 才這樣），
    分岔了會讓「到底」判定跟實際抓取行為不一致。"""
    import fetcher_gaap
    assert local_db.XBRL_CUTOFF == fetcher_gaap._XBRL_CUTOFF


# ── J3：整家跳過的判定（純函式）───────────────────────────────────────────

def _plan(listings, cached, version_ok=True):
    return local_db.plan_ticker(listings, cached, version_ok=version_ok)


def test_plan_skips_when_every_form_reached_bottom_and_nothing_is_new():
    listings = {"10-Q": [("q1", "2013-02-01")], "10-K": [("k1", "2013-11-01")]}
    plan = _plan(listings, {"q1", "k1"})
    assert plan["skip"] is True
    assert plan["new_count"] == 0


def test_plan_does_not_skip_when_a_new_filing_appeared():
    listings = {"10-Q": [("q2", "2026-08-01"), ("q1", "2013-02-01")],
                "10-K": [("k1", "2013-11-01")]}
    plan = _plan(listings, {"q1", "k1"})
    assert plan["skip"] is False
    assert plan["new_count"] == 1


def test_plan_does_not_skip_when_one_form_is_still_unfinished():
    """10-K 到底了但 10-Q 還沒——合記會誤判成整家到底，所以要分 form 記。"""
    listings = {"10-Q": [("q1", "2021-02-01"), ("q0", "2020-02-01")],
                "10-K": [("k1", "2013-11-01")]}
    plan = _plan(listings, {"q1", "k1"})
    assert plan["skip"] is False
    assert plan["forms"]["10-K"]["reached_bottom"] == "no_more_filings"
    assert plan["forms"]["10-Q"]["reached_bottom"] is None


def test_plan_never_skips_when_the_cached_data_is_from_another_edgartools_version():
    """版本不符時 `load_filing()` 一律回 None，那些檔案等同不存在——
    這時跳過會讓那家公司永遠停在失效狀態。"""
    listings = {"10-Q": [("q1", "2013-02-01")], "10-K": [("k1", "2013-11-01")]}
    assert _plan(listings, {"q1", "k1"}, version_ok=False)["skip"] is False


# ── J2：_meta.json 的讀寫與自癒 ───────────────────────────────────────────

def test_rebuild_meta_counts_per_form_with_oldest_and_newest(cache_dir):
    _write_filing("AAPL", _acc(1), form="10-Q", filing_date="2025-08-01")
    _write_filing("AAPL", _acc(2), form="10-Q", filing_date="2024-08-01")
    _write_filing("AAPL", _acc(3), form="10-K", filing_date="2025-11-01")
    meta = local_db.rebuild_meta("AAPL")
    assert meta["ticker"] == "AAPL"
    assert meta["file_count"] == 3
    assert meta["cik"] == 320193
    assert meta["forms"]["10-Q"] == {
        "count": 2, "oldest": "2024-08-01", "newest": "2025-08-01",
        # `_write_filing` 寫的 dataframes 是 {k: None}，抽不出期末日 → None。
        # 有期間的情況由 test_rebuild_meta_records_real_fiscal_periods 蓋。
        "period_oldest": None, "period_newest": None,
        "reached_bottom": None, "reached_bottom_stale": False,
    }
    assert meta["forms"]["10-K"]["count"] == 1


def test_load_meta_uses_the_snapshot_when_the_file_count_matches(cache_dir):
    """快取路徑：`file_count` 對得上就直接用，不去讀 881 個 JSON。
    這裡把 meta 裡的 count 動手腳成 999，讀回來還是 999 就證明沒重算。"""
    _write_filing("AAPL", _acc(1), form="10-Q", filing_date="2025-08-01")
    meta = local_db.rebuild_meta("AAPL")
    meta["forms"]["10-Q"]["count"] = 999
    local_db.write_meta("AAPL", meta)
    assert local_db.load_meta("AAPL")["forms"]["10-Q"]["count"] == 999


def test_load_meta_rebuilds_when_the_directory_grew_behind_its_back(cache_dir):
    _write_filing("AAPL", _acc(1), form="10-Q", filing_date="2025-08-01")
    local_db.write_meta("AAPL", local_db.rebuild_meta("AAPL"))
    _write_filing("AAPL", _acc(2), form="10-Q", filing_date="2026-08-01")
    meta = local_db.load_meta("AAPL")
    assert meta["file_count"] == 2
    assert meta["forms"]["10-Q"]["newest"] == "2026-08-01"


def test_load_meta_rebuilds_when_the_schema_version_does_not_match(cache_dir):
    _write_filing("AAPL", _acc(1), form="10-Q", filing_date="2025-08-01")
    local_db.write_meta("AAPL", {"schema_version": 999, "ticker": "AAPL",
                                 "file_count": 1, "forms": {}})
    meta = local_db.load_meta("AAPL")
    assert meta["schema_version"] == local_db.META_SCHEMA_VERSION
    assert meta["forms"]["10-Q"]["count"] == 1


def test_load_meta_rebuilds_when_the_file_is_corrupt(cache_dir):
    _write_filing("AAPL", _acc(1), form="10-Q", filing_date="2025-08-01")
    local_db.meta_path("AAPL").write_text("{ this is not json", encoding="utf-8")
    assert local_db.load_meta("AAPL")["file_count"] == 1


def test_rebuild_keeps_reached_bottom_but_marks_it_stale(cache_dir):
    """重算 `reached_bottom` 要連網拿完整清單，不該為了顯示一列就連 201 次網。
    所以保留舊值、標記過期，下次「更新本地庫」跑到時再重算。"""
    _write_filing("AAPL", _acc(1), form="10-Q", filing_date="2025-08-01")
    meta = local_db.rebuild_meta("AAPL")
    meta["forms"]["10-Q"]["reached_bottom"] = "xbrl_cutoff"
    local_db.write_meta("AAPL", meta)
    _write_filing("AAPL", _acc(2), form="10-Q", filing_date="2026-08-01")
    healed = local_db.load_meta("AAPL")
    assert healed["forms"]["10-Q"]["reached_bottom"] == "xbrl_cutoff"
    assert healed["forms"]["10-Q"]["reached_bottom_stale"] is True


def test_load_meta_returns_none_for_a_company_with_no_cache(cache_dir):
    assert local_db.load_meta("NOPE") is None


def test_meta_json_is_not_counted_as_a_filing(cache_dir):
    """`_meta.json` 混在同一個資料夾，`ACCESSION_RE` 那道閘要擋住它，
    不然份數會多一。"""
    _write_filing("AAPL", _acc(1), form="10-Q", filing_date="2025-08-01")
    local_db.write_meta("AAPL", local_db.rebuild_meta("AAPL"))
    count, _size = filing_cache._dir_stats(filing_cache.ticker_dir("AAPL"))
    assert count == 1
    assert local_db.load_meta("AAPL")["file_count"] == 1


def test_a_folder_with_only_meta_is_not_listed_as_a_cached_company(cache_dir):
    """清除之後只剩 `_meta.json` 的資料夾，GUI 不該出現一列「0 份」。"""
    local_db.write_meta("GHOST", {"schema_version": local_db.META_SCHEMA_VERSION,
                                  "ticker": "GHOST", "file_count": 0, "forms": {}})
    assert [r["ticker"] for r in filing_cache.list_cached_tickers()] == []


def test_clearing_a_ticker_takes_its_meta_with_it(cache_dir):
    _write_filing("AAPL", _acc(1), form="10-Q", filing_date="2025-08-01")
    local_db.write_meta("AAPL", local_db.rebuild_meta("AAPL"))
    assert filing_cache.clear_ticker("AAPL") is True
    assert not local_db.meta_path("AAPL").exists()


# ── J1：更新名單 ──────────────────────────────────────────────────────────

def test_update_list_defaults_to_empty():
    from config import DEFAULT_CONFIG
    assert DEFAULT_CONFIG[local_db.UPDATE_LIST_KEY] == []


def test_update_list_normalises_case_whitespace_and_duplicates():
    cfg = {}
    local_db.set_update_list(cfg, [" aapl ", "MSFT", "aapl", "", None])
    assert cfg[local_db.UPDATE_LIST_KEY] == ["AAPL", "MSFT"]


def test_add_tickers_returns_only_the_ones_that_were_actually_new():
    cfg = {local_db.UPDATE_LIST_KEY: ["AAPL"]}
    assert local_db.add_tickers(cfg, ["aapl", "NVDA"]) == ["NVDA"]
    assert cfg[local_db.UPDATE_LIST_KEY] == ["AAPL", "NVDA"]


def test_remove_ticker():
    cfg = {local_db.UPDATE_LIST_KEY: ["AAPL", "NVDA"]}
    local_db.remove_ticker(cfg, "nvda")
    assert cfg[local_db.UPDATE_LIST_KEY] == ["AAPL"]


def test_import_from_watchlist_reads_the_ticker_field():
    cfg = {"watchlist": [{"ticker": "AAPL", "name": "Apple"}, {"ticker": "NVDA"}],
           local_db.UPDATE_LIST_KEY: ["AAPL"]}
    assert local_db.import_from_watchlist(cfg) == ["NVDA"]
    assert cfg[local_db.UPDATE_LIST_KEY] == ["AAPL", "NVDA"]


def test_import_from_cache_reads_the_cache_directory(cache_dir):
    _write_filing("AAPL", _acc(1), form="10-Q", filing_date="2025-08-01")
    _write_filing("NVDA", _acc(2), form="10-Q", filing_date="2025-08-01")
    cfg = {local_db.UPDATE_LIST_KEY: []}
    assert sorted(local_db.import_from_cache(cfg)) == ["AAPL", "NVDA"]


def test_update_list_and_watchlist_stay_independent():
    """兩份名單刻意分開：合併會讓 Tab 2 一按產 201 份 Excel。"""
    cfg = {"watchlist": [{"ticker": "AAPL"}], local_db.UPDATE_LIST_KEY: []}
    local_db.add_tickers(cfg, ["NVDA"])
    assert [w["ticker"] for w in cfg["watchlist"]] == ["AAPL"]


# ── J4：版本鎖與版本不符偵測 ──────────────────────────────────────────────

def test_requirements_pins_edgartools_to_an_exact_version():
    """不鎖的話任何人重跑一次 `pip install -r requirements.txt` 就可能
    讓整個本地庫失效。"""
    req = (Path(__file__).parent.parent / "requirements.txt").read_text(encoding="utf-8")
    assert any(line.strip().startswith("edgartools==") for line in req.splitlines())


def test_pinned_version_matches_what_is_installed():
    assert local_db.pinned_edgartools_version() == filing_cache.edgartools_version()


def test_stale_summary_is_empty_when_every_company_matches(cache_dir, monkeypatch):
    monkeypatch.setattr(filing_cache, "edgartools_version", lambda: "5.29.0")
    _write_filing("AAPL", _acc(1), form="10-Q", filing_date="2025-08-01",
                  version="5.29.0")
    summary = local_db.stale_cache_summary()
    assert summary["companies"] == []
    assert summary["n_filings"] == 0


def test_stale_summary_lists_companies_parsed_by_another_version(cache_dir, monkeypatch):
    monkeypatch.setattr(filing_cache, "edgartools_version", lambda: "5.31.0")
    _write_filing("AAPL", _acc(1), form="10-Q", filing_date="2025-08-01",
                  version="5.29.0")
    _write_filing("AAPL", _acc(2), form="10-K", filing_date="2025-11-01",
                  version="5.29.0")
    _write_filing("NVDA", _acc(3), form="10-Q", filing_date="2025-08-01",
                  version="5.31.0")
    summary = local_db.stale_cache_summary()
    assert summary["companies"] == ["AAPL"]
    assert summary["n_filings"] == 2
    assert summary["old_versions"] == ["5.29.0"]
    assert summary["current"] == "5.31.0"
    assert summary["estimated_seconds"] > 0


def test_stale_summary_is_empty_when_the_version_cannot_be_read(cache_dir, monkeypatch):
    """取不到版本時不該恐嚇使用者說全部要重抓。"""
    monkeypatch.setattr(filing_cache, "edgartools_version", lambda: None)
    _write_filing("AAPL", _acc(1), form="10-Q", filing_date="2025-08-01")
    assert local_db.stale_cache_summary()["companies"] == []


# ── J3：更新本地庫的流程 ──────────────────────────────────────────────────

class _FakeEdgar:
    """假的 EDGAR。`listings` 是 {ticker: {form: [(accession, date), ...]}}，
    `fetch()` 把清單裡的東西寫進快取——就是真 `fetch_gaap_statements()` 對
    快取的副作用。"""

    def __init__(self, listings, fail: set[str] | None = None):
        self.listings = listings
        self.fail = fail or set()
        self.fetched: list[str] = []
        self.listed: list[str] = []

    def list_filings(self, ticker, identity):
        self.listed.append(ticker)
        if ticker in self.fail:
            raise RuntimeError("boom")
        return self.listings[ticker], 320193

    def fetch(self, ticker, identity, max_filings, max_annual_filings):
        self.fetched.append(ticker)
        for form, rows in self.listings[ticker].items():
            for acc, filing_date in rows:
                if filing_date and filing_date >= "2008-01-01":
                    _write_filing(ticker, acc, form=form, filing_date=filing_date)
        return None


def _run(edgar, tickers):
    return local_db.update_local_db(
        tickers, "tester tester@example.com",
        list_filings=edgar.list_filings, fetch=edgar.fetch)


def test_update_fetches_a_company_that_has_no_cache_at_all(cache_dir):
    edgar = _FakeEdgar({"META": {"10-Q": [(_acc(1), "2013-05-01")],
                                 "10-K": [(_acc(2), "2013-02-01")]}})
    report = _run(edgar, ["META"])
    assert edgar.fetched == ["META"]
    assert report.updated == 1 and report.skipped == 0
    meta = local_db.load_meta("META")
    assert meta["forms"]["10-Q"]["reached_bottom"] == "no_more_filings"
    assert meta["forms"]["10-Q"]["reached_bottom_stale"] is False


def test_second_run_skips_the_whole_company(cache_dir):
    """驗收條件：已經到底又沒有新財報的公司，第二次執行整家跳過，
    完全不進抓取迴圈——這是「不要每次全部重抓」的核心。"""
    edgar = _FakeEdgar({"META": {"10-Q": [(_acc(1), "2013-05-01")],
                                 "10-K": [(_acc(2), "2013-02-01")]}})
    _run(edgar, ["META"])
    report = _run(edgar, ["META"])
    assert edgar.fetched == ["META"]          # 沒有第二次
    assert edgar.listed == ["META", "META"]   # 但清單還是查了（很便宜）
    assert report.skipped == 1 and report.updated == 0


def test_a_new_filing_pulls_the_company_back_into_the_fetch_loop(cache_dir):
    listings = {"META": {"10-Q": [(_acc(1), "2013-05-01")],
                         "10-K": [(_acc(2), "2013-02-01")]}}
    edgar = _FakeEdgar(listings)
    _run(edgar, ["META"])
    listings["META"]["10-Q"].insert(0, (_acc(3), "2026-08-01"))
    report = _run(edgar, ["META"])
    assert edgar.fetched == ["META", "META"]
    assert report.updated == 1
    assert local_db.load_meta("META")["forms"]["10-Q"]["newest"] == "2026-08-01"


def test_one_company_failing_does_not_stop_the_rest(cache_dir):
    edgar = _FakeEdgar(
        {"AAPL": {"10-Q": [(_acc(1), "2013-05-01")], "10-K": []},
         "NVDA": {"10-Q": [(_acc(2), "2013-05-01")], "10-K": []}},
        fail={"AAPL"})
    report = _run(edgar, ["AAPL", "NVDA"])
    assert report.failed == 1 and report.updated == 1
    assert [r.ticker for r in report.results if r.status == "failed"] == ["AAPL"]
    assert "RuntimeError" in report.results[0].error


def test_update_reports_progress_for_every_company(cache_dir):
    edgar = _FakeEdgar({"META": {"10-Q": [(_acc(1), "2013-05-01")], "10-K": []}})
    seen = []
    local_db.update_local_db(["META"], "x", list_filings=edgar.list_filings,
                             fetch=edgar.fetch, progress=seen.append)
    kinds = [e["event"] for e in seen]
    assert kinds[0] == "start" and kinds[-1] == "done"
    assert any(e["event"] == "ticker_done" and e["ticker"] == "META" for e in seen)


def test_update_stops_early_when_the_caller_asks_it_to(cache_dir):
    """GUI 關視窗／CLI Ctrl-C 時要停得下來。已抓到的份數本來就都在磁碟上。"""
    edgar = _FakeEdgar({"AAPL": {"10-Q": [(_acc(1), "2013-05-01")], "10-K": []},
                        "NVDA": {"10-Q": [(_acc(2), "2013-05-01")], "10-K": []}})
    report = local_db.update_local_db(
        ["AAPL", "NVDA"], "x", list_filings=edgar.list_filings, fetch=edgar.fetch,
        should_stop=lambda: len(edgar.fetched) >= 1)
    assert edgar.fetched == ["AAPL"]
    assert report.stopped is True


def test_update_skips_blank_and_duplicate_tickers(cache_dir):
    edgar = _FakeEdgar({"AAPL": {"10-Q": [(_acc(1), "2013-05-01")], "10-K": []}})
    _run(edgar, ["aapl", "AAPL", "", None])
    assert edgar.listed == ["AAPL"]


# ── GUI 那一列的文字（純函式，Tk 的部分照專案現況用探針手動驗）─────────────

def test_row_text_shows_the_span_across_both_forms():
    from main import local_db_row_text
    meta = {"forms": {
        "10-Q": {"oldest": "2008-02-01", "newest": "2026-07-31",
                 "reached_bottom": "xbrl_cutoff"},
        "10-K": {"oldest": "2008-11-05", "newest": "2025-10-31",
                 "reached_bottom": "xbrl_cutoff"}}}
    span, bottom = local_db_row_text(meta)
    assert span == "2008-02~2026-07"
    assert "?" not in bottom


def test_row_text_says_partial_when_one_form_is_unfinished():
    """兩個 form 只要有一個沒到底就顯示未到底——「還要不要再挖」是整家一起
    決定的。"""
    from main import local_db_row_text
    meta = {"forms": {"10-Q": {"oldest": "2021-02-01", "newest": "2026-02-01",
                               "reached_bottom": None},
                      "10-K": {"oldest": "2013-11-01", "newest": "2025-11-01",
                               "reached_bottom": "no_more_filings"}}}
    _span, bottom = local_db_row_text(meta)
    import i18n
    assert bottom == i18n.t("gui.lbl.db_bottom_no")


def test_row_text_marks_a_stale_reached_bottom_with_a_question_mark():
    from main import local_db_row_text
    meta = {"forms": {
        "10-Q": {"oldest": "2013-02-01", "newest": "2026-02-01",
                 "reached_bottom": "no_more_filings", "reached_bottom_stale": True},
        "10-K": {"oldest": "2013-02-01", "newest": "2026-02-01",
                 "reached_bottom": "no_more_filings"}}}
    assert local_db_row_text(meta)[1].endswith("?")


def test_row_text_survives_a_missing_meta():
    """meta 還沒建（或剛被刪）時 GUI 不能炸——這個清單每次切到 Tab3 都會畫。"""
    from main import local_db_row_text
    assert local_db_row_text(None) == ("—", "—")
    assert local_db_row_text({"forms": {}}) == ("—", "—")


# ── 「便宜」這件事要真的便宜（2026-09-04 自我複查抓到的兩處）─────────────

def test_stale_summary_does_not_open_every_filing_when_meta_is_fresh(cache_dir):
    """啟動時的版本偵測**不可以**去讀每一份 filing。

    201 家拓到底是 16,000 份檔案——原本的寫法對每一家呼叫 `scan_filings()`，
    等於每次開程式都把整個本地庫讀一遍。meta 新鮮時它就已經記著版本了。
    """
    _write_filing("AAPL", _acc(1), form="10-Q", filing_date="2025-08-01",
                  version="5.29.0")
    local_db.write_meta("AAPL", local_db.rebuild_meta("AAPL"))

    def boom(_ticker):
        raise AssertionError("不該為了偵測版本去讀 filing 檔")

    original = local_db.scan_filings
    local_db.scan_filings = boom
    try:
        summary = local_db.stale_cache_summary()
    finally:
        local_db.scan_filings = original
    assert summary["companies"] == []


def test_stale_summary_still_finds_stale_companies_from_meta(cache_dir, monkeypatch):
    monkeypatch.setattr(filing_cache, "edgartools_version", lambda: "5.31.0")
    _write_filing("AAPL", _acc(1), form="10-Q", filing_date="2025-08-01",
                  version="5.29.0")
    local_db.write_meta("AAPL", local_db.rebuild_meta("AAPL"))
    summary = local_db.stale_cache_summary()
    assert summary["companies"] == ["AAPL"]
    assert summary["n_filings"] == 1
    assert summary["old_versions"] == ["5.29.0"]


def test_skipping_a_company_does_not_reread_its_filings(cache_dir):
    """「整家跳過」要真的便宜。原本跳過之後還是照樣重建一次 meta
    （＝把那家的 75 份檔案全部開一遍），跳過的意義就少了一半。"""
    edgar = _FakeEdgar({"META": {"10-Q": [(_acc(1), "2013-05-01")],
                                 "10-K": [(_acc(2), "2013-02-01")]}})
    _run(edgar, ["META"])                      # 第一輪：建好 meta

    calls = []
    original = local_db.scan_filings

    def counting(ticker):
        calls.append(ticker)
        return original(ticker)

    local_db.scan_filings = counting
    try:
        report = _run(edgar, ["META"])         # 第二輪：整家跳過
    finally:
        local_db.scan_filings = original
    assert report.skipped == 1
    assert calls == [], f"跳過的公司不該重讀 filing，實際讀了 {calls}"


def test_skipping_still_refreshes_reached_bottom_and_timestamp(cache_dir):
    """便宜歸便宜，`reached_bottom` 還是要更新成這一輪剛連網算出來的，
    過期標記也要清掉——不然 GUI 上那個「?」永遠拿不掉。"""
    edgar = _FakeEdgar({"META": {"10-Q": [(_acc(1), "2013-05-01")], "10-K": []}})
    _run(edgar, ["META"])
    meta = local_db.read_meta("META")
    meta["forms"]["10-Q"]["reached_bottom"] = None
    meta["forms"]["10-Q"]["reached_bottom_stale"] = True
    meta["updated_at"] = "2000-01-01T00:00:00+08:00"
    local_db.write_meta("META", meta)

    _run(edgar, ["META"])
    healed = local_db.read_meta("META")
    assert healed["forms"]["10-Q"]["reached_bottom"] == "no_more_filings"
    assert healed["forms"]["10-Q"]["reached_bottom_stale"] is False
    assert healed["updated_at"] > "2000-01-01"


def test_skip_path_falls_back_to_rebuild_when_meta_is_missing_a_form(cache_dir):
    """殘缺的 meta 不可以走「沿用」那條路。

    份數對得上所以 `load_meta()` 不會自癒它——沿用的話會生出一個只有
    `reached_bottom`、沒有 `count` 的條目，而且**會一直錯下去**。
    """
    edgar = _FakeEdgar({"META": {"10-Q": [(_acc(1), "2013-05-01")], "10-K": []}})
    _run(edgar, ["META"])
    broken = local_db.read_meta("META")
    del broken["forms"]["10-K"]
    local_db.write_meta("META", broken)

    report = _run(edgar, ["META"])
    assert report.skipped == 1
    healed = local_db.read_meta("META")
    assert healed["forms"]["10-K"]["count"] == 0
    assert healed["forms"]["10-Q"]["count"] == 1


@pytest.mark.parametrize("meta, ok", [
    (None, False),
    ({"file_count": 9, "forms": {"10-Q": {"count": 1}, "10-K": {"count": 0}}}, False),
    ({"file_count": 1, "forms": None}, False),
    ({"file_count": 1, "forms": {"10-Q": {"count": 1}}}, False),
    ({"file_count": 1, "forms": {"10-Q": {}, "10-K": {"count": 0}}}, False),
    ({"file_count": 1, "forms": {"10-Q": {"count": 1}, "10-K": {"count": 0}}}, True),
])
def test_meta_is_reusable(meta, ok):
    assert local_db._meta_is_reusable(meta, 1) is ok


# ── J7：資料庫總覽（唯讀純函式）──────────────────────────────────────────

def _row(ticker="AAPL", **kw):
    """總覽的一列，只填測試關心的欄位。"""
    base = {"ticker": ticker, "filings": 10, "size_bytes": 1000,
            "filed_from": "2008-05-01", "filed_to": "2026-06-01",
            "period_from": "2008-03-29", "period_to": "2026-03-28",
            "years": 18.0, "bottom": local_db.BOTTOM_YES, "bottom_stale": False,
            "in_list": True, "meta_ok": True, "edgartools_version": "5.29.0",
            "updated_at": "2026-09-18T10:00:00+08:00", "stale_days": 0}
    return {**base, **kw}


def test_overview_row_reports_filed_dates_not_fiscal_periods(cache_dir):
    """`filed_from`／`filed_to` 是 SEC 收件日。這條釘的是語意，不是格式——
    顯示端要據此標成「申報日期」，標成「涵蓋期間」就是講假話。"""
    _write_filing("AAPL", _acc(1), form="10-K", filing_date="2008-05-01")
    _write_filing("AAPL", _acc(2), form="10-Q", filing_date="2026-06-01")
    local_db.load_meta("AAPL")          # 先讓 meta 生出來

    cached = {"ticker": "AAPL", "count": 2, "size_bytes": 999}
    row = local_db.overview_row("AAPL", cached, in_list=False)
    assert row["filed_from"] == "2008-05-01"
    assert row["filed_to"] == "2026-06-01"
    assert row["years"] == 18.1


def test_overview_row_never_writes_to_disk(cache_dir):
    """總覽是唯讀的。`load_meta()` 會自癒並寫檔，`overview_row()` 不可以——
    244 家「瞄一眼」不該變成上萬次檔案讀取加一輪寫入。"""
    _write_filing("NVDA", _acc(1), form="10-Q", filing_date="2020-05-01")
    meta_file = local_db.meta_path("NVDA")
    assert not meta_file.exists()        # 還沒有 meta，正是最會誘發自癒的狀態

    cached = {"ticker": "NVDA", "count": 1, "size_bytes": 100}
    row = local_db.overview_row("NVDA", cached, in_list=False)

    assert not meta_file.exists()        # 沒有被偷偷建出來
    assert row["meta_ok"] is False       # 而且照實回報「對不上」
    assert row["filings"] == 1           # 份數仍然正確（來自目錄，不是 meta）


def test_overview_row_flags_meta_that_disagrees_with_the_directory(cache_dir):
    _write_filing("META", _acc(1), form="10-Q", filing_date="2020-05-01")
    local_db.load_meta("META")
    _write_filing("META", _acc(2), form="10-Q", filing_date="2021-05-01")

    cached = {"ticker": "META", "count": 2, "size_bytes": 100}
    assert local_db.overview_row("META", cached, in_list=False)["meta_ok"] is False


def test_overview_rows_are_alphabetical_and_mark_the_update_list(cache_dir):
    for ticker in ("NVDA", "AAPL", "MSFT"):
        _write_filing(ticker, _acc(1), form="10-Q", filing_date="2020-05-01")

    rows = local_db.overview_rows({"local_db_tickers": ["AAPL"]})
    assert [r["ticker"] for r in rows] == ["AAPL", "MSFT", "NVDA"]
    assert [r["in_list"] for r in rows] == [True, False, False]


def test_overview_summary_totals():
    rows = [_row("AAPL", filings=75, size_bytes=1000),
            _row("NVDA", filings=60, size_bytes=500)]
    assert local_db.overview_summary(rows) == {
        "companies": 2, "filings": 135, "size_bytes": 1500}


def test_filter_matches_anywhere_in_the_ticker_ignoring_case():
    rows = [_row("AMD"), _row("AAPL"), _row("NVDA")]
    assert [r["ticker"] for r in local_db.filter_overview_rows(rows, "md")] == ["AMD"]
    # NVDA 也含 "A"，所以三家裡有兩家中
    assert len(local_db.filter_overview_rows(rows, "a")) == 3
    assert len(local_db.filter_overview_rows(rows, "")) == 3


def test_sort_puts_missing_values_last_in_both_directions():
    """算不出年數的公司是「資料不全」，升冪降冪都該沉底。
    降冪若直接 `reverse=True`，`None` 會被翻到最前面，把真正的極端值埋掉。"""
    rows = [_row("AAPL", years=18.1), _row("BLK", years=None), _row("CRM", years=2.0)]

    asc = local_db.sort_overview_rows(rows, "years")
    assert [r["ticker"] for r in asc] == ["CRM", "AAPL", "BLK"]

    desc = local_db.sort_overview_rows(rows, "years", descending=True)
    assert [r["ticker"] for r in desc] == ["AAPL", "CRM", "BLK"]


def test_sort_falls_back_to_ticker_when_the_key_is_unknown():
    """點壞一個欄位標題不該讓整頁炸掉。"""
    rows = [_row("NVDA"), _row("AAPL")]
    assert [r["ticker"] for r in local_db.sort_overview_rows(rows, "nope")] \
        == ["AAPL", "NVDA"]


# ── J7：總覽的顯示層（`main.py` 的純函式，Treeview 本身用探針驗）──────────

def test_overview_cells_show_fiscal_periods_not_filed_dates(cache_dir):
    """主欄位顯示的必須是**財報期間**，不是 SEC 收件日。兩者差一整個期間，
    分析師問的是「我有哪幾季的數字」。"""
    from i18n import set_lang
    from main import DB_OVERVIEW_COLUMNS, db_overview_cells
    set_lang("zh_tw")

    from locales.zh_tw import STRINGS
    assert "財報期間" == STRINGS["gui.col.db_period_span"]

    cells = db_overview_cells(_row(period_from="2008-03-29", period_to="2026-03-28",
                                   filed_from="2008-05-01", filed_to="2026-06-01"))
    assert len(cells) == len(DB_OVERVIEW_COLUMNS)
    # 顯示 03/29 那組（財報期間），不是 05/01 那組（收件日）
    assert cells[DB_OVERVIEW_COLUMNS.index("period_span")] == "2008-03 ~ 2026-03"


def test_overview_cells_render_missing_values_as_dashes():
    from main import DB_OVERVIEW_COLUMNS, db_overview_cells
    cells = db_overview_cells(_row(period_from=None, period_to=None, years=None,
                                   bottom=local_db.BOTTOM_UNKNOWN, stale_days=None))
    assert cells[DB_OVERVIEW_COLUMNS.index("period_span")] == "—"
    assert cells[DB_OVERVIEW_COLUMNS.index("checked")] == "—"
    assert cells[DB_OVERVIEW_COLUMNS.index("years")] == "—"
    assert cells[DB_OVERVIEW_COLUMNS.index("bottom")] == "—"


def test_overview_cells_mark_stale_bottom_with_a_question_mark():
    """`reached_bottom` 是上一輪留下的值就加問號——不加的話，一個「已到底」
    看起來跟真的重算過一樣可信。"""
    from main import DB_OVERVIEW_COLUMNS, db_overview_cells
    idx = DB_OVERVIEW_COLUMNS.index("bottom")
    assert db_overview_cells(_row(bottom_stale=True))[idx].endswith("?")
    assert not db_overview_cells(_row(bottom_stale=False))[idx].endswith("?")
    # 本來就沒有 meta 的不加問號——「—?」沒有意義
    assert db_overview_cells(
        _row(bottom=local_db.BOTTOM_UNKNOWN, bottom_stale=True))[idx] == "—"


def test_overview_cells_flag_rows_whose_snapshot_is_stale():
    from main import DB_OVERVIEW_COLUMNS, db_overview_cells
    idx = DB_OVERVIEW_COLUMNS.index("note")
    assert db_overview_cells(_row(meta_ok=True))[idx] == ""
    assert db_overview_cells(_row(meta_ok=False))[idx] != ""


def test_overview_csv_has_a_header_and_one_line_per_company():
    from main import DB_OVERVIEW_COLUMNS, db_overview_csv
    text = db_overview_csv([_row("AAPL"), _row("NVDA")])
    lines = text.strip().splitlines()
    assert len(lines) == 3                       # 標題 + 2 家
    assert len(lines[0].split(",")) == len(DB_OVERVIEW_COLUMNS)
    assert lines[1].startswith("AAPL")
    assert lines[2].startswith("NVDA")


# ── J7 後續：真正的財報期間、上次更新時間、只抓幾家（2026-09-18）──────────

def _write_filing_with_periods(ticker, accession, *, form, filing_date, period_cols):
    """帶真實 DataFrame 欄名的快取檔——期末日就是從這些欄名抽出來的。"""
    import pandas as pd
    df = pd.DataFrame({"concept": ["Revenue"], "label": ["Revenues"],
                       **{c: [1.0] for c in period_cols}})
    path = filing_cache.ticker_dir(ticker) / f"{accession}.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({
        "schema_version": filing_cache.SCHEMA_VERSION,
        "accession_no": accession, "form": form, "filing_date": filing_date,
        "cached_at": "2026-09-05T00:00:00+08:00", "cik": 320193,
        "edgartools_version": "5.29.0", "has_financials": True,
        "dataframes": {"income_statement": filing_cache.df_to_payload(df),
                       "balance_sheet": None, "cashflow_statement": None},
    }, ensure_ascii=False), encoding="utf-8")


def test_period_end_picks_the_latest_column_not_the_comparative():
    """一份 10-Q 的欄位含去年同期，當期是**最新**那個。挑錯會讓期間整段偏一年。"""
    import pandas as pd
    df = pd.DataFrame({"concept": ["Rev"], "label": ["Revenues"],
                       "2026-03-29 (Q1)": [1.0], "2025-03-30 (Q1)": [0.9]})
    entry = {"dataframes": {"income_statement": filing_cache.df_to_payload(df)}}
    assert local_db.period_end_of(entry) == "2026-03-29"


def test_period_end_handles_bare_dates_from_the_balance_sheet():
    """資產負債表的欄名是裸日期（instant），沒有 `(Q1)` 後綴。"""
    import pandas as pd
    df = pd.DataFrame({"concept": ["Cash"], "2024-03-31": [1.0]})
    entry = {"dataframes": {"balance_sheet": filing_cache.df_to_payload(df)}}
    assert local_db.period_end_of(entry) == "2024-03-31"


def test_period_end_is_none_for_negative_cache_entries():
    """pre-XBRL 的負向快取 `dataframes` 是 null，本來就沒有期間，不該炸。"""
    assert local_db.period_end_of({"dataframes": None}) is None
    assert local_db.period_end_of({}) is None


def test_rebuild_meta_records_real_fiscal_periods(cache_dir):
    """meta 要同時存收件日與財報期間——兩者差一整個期間。"""
    _write_filing_with_periods("AAPL", _acc(1), form="10-K",
                               filing_date="2008-05-01",
                               period_cols=["2007-12-29 (FY)", "2006-12-30 (FY)"])
    _write_filing_with_periods("AAPL", _acc(2), form="10-K",
                               filing_date="2026-06-01",
                               period_cols=["2026-03-28 (FY)"])
    meta = local_db.rebuild_meta("AAPL")
    form = meta["forms"]["10-K"]
    assert (form["oldest"], form["newest"]) == ("2008-05-01", "2026-06-01")
    assert (form["period_oldest"], form["period_newest"]) \
        == ("2007-12-29", "2026-03-28")


def test_overview_years_come_from_fiscal_periods_not_filed_dates(cache_dir):
    """年數要用財報期間算——那才是「我手上有幾年的數字」。"""
    _write_filing_with_periods("AAPL", _acc(1), form="10-K",
                               filing_date="2008-05-01",
                               period_cols=["2007-12-29 (FY)"])
    _write_filing_with_periods("AAPL", _acc(2), form="10-K",
                               filing_date="2026-06-01",
                               period_cols=["2025-12-27 (FY)"])
    local_db.load_meta("AAPL")
    row = local_db.overview_row("AAPL", {"ticker": "AAPL", "count": 2,
                                         "size_bytes": 100}, in_list=False)
    assert (row["period_from"], row["period_to"]) == ("2007-12-29", "2025-12-27")
    # 期間跨 18.0 年，收件日跨 18.1 年——取的是前者
    assert row["years"] == local_db._span_years("2007-12-29", "2025-12-27")


def test_schema_bump_invalidates_old_metas(cache_dir):
    """schema 從 1 升到 2，舊 meta 沒有 period_* 欄位，必須整份重建。"""
    _write_filing("AAPL", _acc(1), form="10-Q", filing_date="2020-05-01")
    local_db.write_meta("AAPL", {"schema_version": 1, "ticker": "AAPL",
                                 "file_count": 1, "forms": {}})
    assert local_db.read_meta("AAPL") is None       # 舊版一律判無效


@pytest.mark.parametrize("stamp, expected", [
    ("2026-09-18T10:00:00+08:00", 0),
    ("2026-09-15T10:00:00+08:00", 3),
    ("", None),
    (None, None),
    ("not a timestamp", None),
])
def test_days_since_converts_updated_at(stamp, expected):
    """`updated_at` 在 2026-09-18 之前寫了但從來沒人讀，等於白存。"""
    assert local_db.days_since(stamp, today=date(2026, 9, 18)) == expected


def test_days_since_never_returns_negative():
    """時鐘不同步／時區換算讓 updated_at 看起來在未來時，不要顯示「-1 天前」。"""
    assert local_db.days_since("2026-09-20T10:00:00+08:00",
                               today=date(2026, 9, 18)) == 0


def test_overview_row_exposes_when_we_last_checked(cache_dir):
    _write_filing("NVDA", _acc(1), form="10-Q", filing_date="2020-05-01")
    local_db.load_meta("NVDA")
    row = local_db.overview_row("NVDA", {"ticker": "NVDA", "count": 1,
                                         "size_bytes": 10}, in_list=False)
    assert row["updated_at"]                 # meta 剛寫，有時間戳
    assert row["stale_days"] == 0


def test_checked_text_reads_as_relative_time():
    from i18n import set_lang
    from main import db_checked_text
    set_lang("zh_tw")
    assert db_checked_text(0) == "今天"
    assert "3" in db_checked_text(3)
    assert db_checked_text(None) == "—"


def test_rows_can_be_sorted_by_how_stale_they_are():
    """「哪幾家最久沒查」是實際會問的問題，所以 stale_days 要能排序。"""
    rows = [_row("AAPL", stale_days=0), _row("BLK", stale_days=90),
            _row("CRM", stale_days=None)]
    order = [r["ticker"] for r in
             local_db.sort_overview_rows(rows, "stale_days", descending=True)]
    assert order == ["BLK", "AAPL", "CRM"]      # 最久沒查的在最前，未知沉底
