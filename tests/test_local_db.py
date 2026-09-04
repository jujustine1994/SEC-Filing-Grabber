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
    monkeypatch.setenv("APPDATA", str(tmp_path))
    return tmp_path / "SEC Financial Tools" / "filing_cache"


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
