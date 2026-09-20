"""Tests for filing_cache.py — 本地 filing 快取的儲存層。

快取的是「解析層」的輸出（edgartools 解出來的三張 DataFrame），比對層
永遠在快取之上即時重跑。所以這裡釘的是「存進去再讀回來，跟原本一模一樣」，
以及「任何一種對不上的情況都要安靜地退回無快取，不能拋例外、不能餵錯資料」。
"""
import importlib.metadata
import json
import os
from pathlib import Path

import pandas as pd
import pytest

import filing_cache


def _sample_df() -> pd.DataFrame:
    """一張像 edgartools 真的會吐出來的表：str / float64 / int64 / bool 四種
    dtype，外加一整欄都是 None——那一欄是 read_json 會推錯成 float64 的地雷。"""
    return pd.DataFrame({
        "concept":  ["us-gaap_Revenue", "us-gaap_NetIncomeLoss"],
        "label":    ["Net sales", "Net income"],
        "level":    [4, 3],
        "abstract": [False, False],
        "2025-12-27 (Q1)": [1000.0, 200.0],
        "dimension_member_label": [None, None],
    })


# ── 序列化：存進去讀回來要一模一樣 ────────────────────────────────────────

def test_payload_roundtrip_keeps_values_and_dtypes():
    df = _sample_df()
    back = filing_cache.payload_to_df(filing_cache.df_to_payload(df))
    pd.testing.assert_frame_equal(back, df, check_like=False)


def test_payload_is_a_json_object_not_an_escaped_string():
    """存 `df.to_json()` 的字串會讓整份內容被逃逸一次，檔案膨脹 10~15%
    而且文字編輯器打開完全不能看。要存解析過的物件。"""
    payload = filing_cache.df_to_payload(_sample_df())
    assert isinstance(payload["data"], dict)
    assert set(payload["data"]) >= {"columns", "index", "data"}


def test_all_null_column_keeps_its_original_dtype():
    """`to_json(orient="split")` 不含 dtype，整欄皆 null 會被推成 float64。
    這是 spike 實測抓到的唯一一個還原不回去的細節。"""
    df = _sample_df()
    back = filing_cache.payload_to_df(filing_cache.df_to_payload(df))
    assert back["dimension_member_label"].dtype == df["dimension_member_label"].dtype


def test_datetime_column_does_not_come_back_reinterpreted_as_1970():
    """`to_json(orient="split")` 把 datetime64 欄寫成 epoch **毫秒**的整數，
    `pd.DataFrame(...)` 讀回來變成 int64。若照存檔時記下的 dtype 硬
    `astype("datetime64[us]")`，pandas 會把那串整數**當成微秒**重新解讀
    ——毫秒被誤讀成微秒，時間軸整個縮 1000 倍，2025-12-27 會變成
    1970-01-21。這一步不拋例外，`except (TypeError, ValueError)` 完全
    抓不到，所以要直接釘住「不會被還原成 1970 年代」，而不是只驗「沒拋例外」。
    """
    df = pd.DataFrame({"filing_date": pd.to_datetime(["2025-12-27"])})
    back = filing_cache.payload_to_df(filing_cache.df_to_payload(df))
    col = back["filing_date"]
    if pd.api.types.is_datetime64_any_dtype(col):
        # 若還原邏輯宣稱認得 datetime64，數值必須是原本的日期。
        assert col.iloc[0] == df["filing_date"].iloc[0]
    else:
        # 目前的作法：datetime64 不在安全還原清單內，留在 pandas 從 JSON
        # 推斷出來的原始整數（epoch 毫秒，2025-12-27 = 1766793600000）。
        assert col.iloc[0] == 1766793600000


def test_none_and_empty_dataframe_are_not_the_same_thing():
    """`is_stmt is None` 與「空表」在下游（`_current_q_col`）行為不同，
    存檔再讀回來不可以混成同一種。"""
    assert filing_cache.df_to_payload(None) is None
    assert filing_cache.payload_to_df(None) is None

    empty = _sample_df().iloc[0:0]
    back = filing_cache.payload_to_df(filing_cache.df_to_payload(empty))
    assert back is not None
    assert len(back) == 0
    assert list(back.columns) == list(empty.columns)


# ── 替身物件 ──────────────────────────────────────────────────────────────
#
# 快取命中時 `_filing_obj()` 回傳的東西。四個 builder 對 filing 物件的用法
# 只有一種：`.financials` → `income_statement()`/`balance_sheet()`/
# `cashflow_statement()` → `.to_dataframe()`，全部無參數。

def _entry(has_financials=True, is_df="sample", bs_df="sample", cf_df=None):
    def _p(v):
        if isinstance(v, str) and v == "sample":
            return filing_cache.df_to_payload(_sample_df())
        return filing_cache.df_to_payload(v)
    return {
        "schema_version": filing_cache.SCHEMA_VERSION,
        "accession_no": "0001045810-25-000123",
        "form": "10-Q",
        "filing_date": "2025-08-27",
        "cik": 1045810,
        "has_financials": has_financials,
        "dataframes": None if not has_financials else {
            "income_statement": _p(is_df),
            "balance_sheet": _p(bs_df),
            "cashflow_statement": _p(cf_df),
        },
    }


def test_cached_filing_exposes_financials_attribute_directly():
    """有兩處直接寫 `tenq.financials.xxx()` 繞過 `_financials_of()`
    （fetcher_gaap.py:1123、2584-2586），這兩條路也要能吃替身。"""
    filing = filing_cache.cached_filing(_entry())
    df = filing.financials.income_statement().to_dataframe()
    pd.testing.assert_frame_equal(df, _sample_df())


def test_cached_filing_raises_attribute_error_for_anything_else():
    """`_financials_of()` 是 `getattr(tenq, "financials", None)`。替身若對未知
    屬性用 `__getattr__` 兜底回 None，以後有人新用到 filing 的其他屬性時，
    快取命中的路徑會**安靜地把整份 filing 當成沒資料**，清快取重跑卻是好的
    ——這種 bug 極難查。所以未知屬性一律照 Python 預設拋 AttributeError。"""
    filing = filing_cache.cached_filing(_entry())
    with pytest.raises(AttributeError):
        filing.has_earnings
    with pytest.raises(AttributeError):
        filing.obj


def test_missing_statement_comes_back_as_none_not_empty():
    """`is_stmt is None` 這種判斷在 fetcher_gaap.py:1330 / 1499 都有。"""
    filing = filing_cache.cached_filing(_entry(cf_df=None))
    assert filing.financials.cashflow_statement() is None


def test_empty_statement_comes_back_as_an_empty_dataframe_not_none():
    """跟上一條成對：空表不是「沒有這張表」，下游行為不一樣。"""
    empty = _sample_df().iloc[0:0]
    filing = filing_cache.cached_filing(_entry(cf_df=empty))
    stmt = filing.financials.cashflow_statement()
    assert stmt is not None
    assert len(stmt.to_dataframe()) == 0


def test_negative_cache_entry_has_financials_none():
    """pre-XBRL 舊申報：`financials` 本來就是 None，替身要照樣回 None，
    讓 `_financials_of()` 走既有的 `continue` 分支。"""
    filing = filing_cache.cached_filing(_entry(has_financials=False))
    assert filing.financials is None


# ── 替身的 DataFrame memo（TODO I7）───────────────────────────────────────


def test_repeated_to_dataframe_parses_the_payload_only_once():
    """同一份 filing 的同一張表，四個 builder 各自呼叫一次。

    memo 之前 `payload_to_df()` 每次都重跑一輪 JSON → DataFrame → astype，
    ARLO 預設參數實測 **224 次、合計 0.37s**。memo 之後每張表只解析一次。
    """
    filing = filing_cache.cached_filing(_entry())
    calls = {"n": 0}
    orig = filing_cache.payload_to_df

    def counted(payload):
        calls["n"] += 1
        return orig(payload)

    filing_cache.payload_to_df = counted
    try:
        for _ in range(4):
            filing.financials.income_statement().to_dataframe()
    finally:
        filing_cache.payload_to_df = orig
    assert calls["n"] == 1


def test_each_call_returns_an_independent_dataframe():
    """**memo 不可以讓四個 builder 共用同一個 DataFrame 物件。**

    現行程式碼沒有任何一處改動報表 dataframe（全庫零 `inplace=True`、
    零欄位指派），所以共用「現在」是安全的——但哪天有人寫 `df["x"] = ...`
    就會靜默污染其他 builder，而且症狀是「另一張表莫名多一欄」，極難查。

    實測深複製比重新解析便宜 **9.8 倍**（0.17ms vs 1.67ms），
    所以隔離幾乎是免費的，沒有理由省。
    """
    filing = filing_cache.cached_filing(_entry())
    stmt = filing.financials.income_statement()
    first = stmt.to_dataframe()
    second = stmt.to_dataframe()
    assert first is not second
    pd.testing.assert_frame_equal(first, second)

    first["injected"] = 1
    third = stmt.to_dataframe()
    assert "injected" not in third.columns


def test_memo_is_shared_across_repeated_getter_calls():
    """`_CachedFinancials._stmt()` 每次都 new 一個 `_CachedStatement`，
    所以 memo 只放在 statement 身上是沒用的——四個 builder 各自呼叫
    `fin.income_statement()`，拿到的是四個不同物件。memo 要能跨 getter 共用。"""
    filing = filing_cache.cached_filing(_entry())
    assert filing.financials.income_statement() is filing.financials.income_statement()


def test_memo_does_not_leak_between_statements():
    """三張表各自 memo，不可以互相拿到對方的 DataFrame。"""
    bs = _sample_df().rename(columns={"label": "bs_label"})
    filing = filing_cache.cached_filing(_entry(bs_df=bs))
    is_df = filing.financials.income_statement().to_dataframe()
    bs_df = filing.financials.balance_sheet().to_dataframe()
    assert "label" in is_df.columns
    assert "bs_label" in bs_df.columns


# ── 路徑與原子寫入 ────────────────────────────────────────────────────────

@pytest.fixture
def cache_dir(tmp_path, monkeypatch):
    """把快取根目錄導到 tmp_path。`cache_root()` 每次呼叫重讀環境變數，
    所以 monkeypatch 就夠了，不用改模組層常數。"""
    monkeypatch.setenv("SEC_LOCAL_DB_ROOT", str(tmp_path))
    return tmp_path / "filing_cache"


def test_cache_root_defaults_to_the_project_local_db_folder(monkeypatch):
    monkeypatch.delenv("SEC_LOCAL_DB_ROOT", raising=False)
    assert filing_cache.cache_root() == filing_cache._project_root() / "local_db" / "filing_cache"


def test_cache_root_honors_the_override_env_var(cache_dir):
    assert filing_cache.cache_root() == cache_dir


def test_ticker_dir_is_uppercase(cache_dir):
    assert filing_cache.ticker_dir("nvda").name == "NVDA"


def test_filing_path_rejects_anything_that_is_not_an_accession_number(cache_dir):
    assert filing_cache.filing_path("NVDA", "0001045810-25-000123") is not None
    for bad in ("../../etc/passwd", "", "abc", "0001045810-25-00012"):
        assert filing_cache.filing_path("NVDA", bad) is None


def test_atomic_write_leaves_no_tmp_file_behind(cache_dir):
    path = cache_dir / "NVDA" / "x.json"
    assert filing_cache.atomic_write_json(path, {"a": 1}) is True
    assert json.loads(path.read_text(encoding="utf-8")) == {"a": 1}
    assert list(path.parent.glob("*.tmp")) == []


def test_atomic_write_uses_a_pid_suffixed_tmp_then_replaces(cache_dir, monkeypatch):
    """兩個實例（批次抓取＋跨公司比較）有機會同時寫同一個檔名。
    tmp 檔名不帶 PID 的話兩邊會蓋到對方的暫存檔。"""
    seen = {}

    real_replace = os.replace

    def _spy(src, dst):
        seen["src"] = str(src)
        return real_replace(src, dst)

    monkeypatch.setattr(filing_cache.os, "replace", _spy)
    path = cache_dir / "NVDA" / "y.json"
    filing_cache.atomic_write_json(path, {"a": 1})
    assert str(os.getpid()) in seen["src"]
    assert seen["src"].endswith(".tmp")


def test_atomic_write_returns_false_instead_of_raising_when_disk_write_fails(
        cache_dir, monkeypatch):
    """磁碟滿／權限問題時只記 log 繼續跑，不能拖垮整趟抓取。"""
    def _boom(*a, **kw):
        raise OSError("disk full")

    monkeypatch.setattr(filing_cache, "open", _boom, raising=False)
    assert filing_cache.atomic_write_json(cache_dir / "NVDA" / "z.json", {"a": 1}) is False


def test_edgartools_version_is_read_from_package_metadata():
    """edgartools 是硬相依（requirements.txt），在本環境裡一定裝著。要驗證
    版本字串真的會被取回，不是虛幻的 None。"""
    v = filing_cache.edgartools_version()
    assert v is not None
    assert isinstance(v, str)
    assert v[0].isdigit()


def test_edgartools_version_returns_none_when_package_not_found(monkeypatch):
    """實測 `edgar.__version__` 不存在（AttributeError），只能走
    importlib.metadata。取不到就回 None，呼叫端把 None 當成
    「這次不要用快取」——不可以填一個預設值混進檔案裡。"""
    def _raise(*a, **kw):
        raise importlib.metadata.PackageNotFoundError("edgartools")

    monkeypatch.setattr("importlib.metadata.version", _raise)
    assert filing_cache.edgartools_version() is None


# ── 單份 filing 的讀寫與五道閘 ────────────────────────────────────────────

ACC = "0001045810-25-000123"


def _save_sample(ticker="NVDA", cik=1045810, cf=None):
    return filing_cache.save_filing(
        ticker, ACC, form="10-Q", filing_date="2025-08-27", cik=cik,
        dataframes={"income_statement": _sample_df(),
                    "balance_sheet": _sample_df(),
                    "cashflow_statement": cf},
        has_financials=True,
    )


def test_save_then_load_returns_the_same_three_dataframes(cache_dir):
    assert _save_sample() is True
    entry = filing_cache.load_filing("NVDA", ACC, 1045810)
    assert entry is not None
    filing = filing_cache.cached_filing(entry)
    pd.testing.assert_frame_equal(
        filing.financials.income_statement().to_dataframe(), _sample_df())
    assert filing.financials.cashflow_statement() is None


def test_load_returns_none_when_the_file_does_not_exist(cache_dir):
    assert filing_cache.load_filing("NVDA", ACC, 1045810) is None


def test_load_returns_none_when_the_cik_differs(cache_dir):
    """ticker 會換手（公司更名、代號被回收給別家）。這種錯不會報例外，
    只會安靜地把別家公司的數字餵給使用者，比任何其他失效情境都危險。"""
    _save_sample(cik=1045810)
    assert filing_cache.load_filing("NVDA", ACC, 99999) is None


def test_load_returns_none_when_the_edgartools_version_differs(cache_dir, monkeypatch):
    """套件升版可能讓同一份 filing 解出不一樣的結果（新的 standardization
    mapping、XBRL parser 修 bug）。這條軸線跟我們自己的比對規則是兩件事。"""
    _save_sample()
    monkeypatch.setattr(filing_cache, "edgartools_version", lambda: "99.0.0")
    assert filing_cache.load_filing("NVDA", ACC, 1045810) is None


def test_load_returns_none_when_the_schema_version_differs(cache_dir):
    _save_sample()
    path = filing_cache.filing_path("NVDA", ACC)
    raw = json.loads(path.read_text(encoding="utf-8"))
    raw["schema_version"] = filing_cache.SCHEMA_VERSION + 1
    path.write_text(json.dumps(raw), encoding="utf-8")
    assert filing_cache.load_filing("NVDA", ACC, 1045810) is None


def test_load_returns_none_for_corrupt_json_instead_of_raising(cache_dir):
    """損毀的快取不可以拖垮整趟抓取——當作沒這份，重抓後覆蓋掉。"""
    _save_sample()
    filing_cache.filing_path("NVDA", ACC).write_text("{ not json", encoding="utf-8")
    assert filing_cache.load_filing("NVDA", ACC, 1045810) is None


def test_cache_is_disabled_entirely_when_the_version_is_unavailable(
        cache_dir, monkeypatch):
    """取不到 edgartools 版本時不讀也不寫，不留下無法比對版本的檔案。"""
    monkeypatch.setattr(filing_cache, "edgartools_version", lambda: None)
    assert _save_sample() is False
    assert filing_cache.load_filing("NVDA", ACC, 1045810) is None


def test_negative_cache_records_a_filing_with_no_financials(cache_dir):
    """pre-XBRL 舊申報（2009 年前常見）：記著「這份試過了、沒有財務資料」，
    下次不用再打一次 SEC 重試。"""
    assert filing_cache.save_filing(
        "NVDA", ACC, form="10-Q", filing_date="2008-05-01", cik=1045810,
        dataframes=None, has_financials=False) is True
    entry = filing_cache.load_filing("NVDA", ACC, 1045810)
    assert entry is not None
    assert entry["has_financials"] is False
    assert filing_cache.cached_filing(entry).financials is None


def test_save_rejects_a_malformed_accession(cache_dir):
    assert filing_cache.save_filing(
        "NVDA", "not-an-accession", form="10-Q", filing_date="2025-08-27",
        cik=1045810, dataframes=None, has_financials=False) is False


def test_load_rejects_an_entry_whose_version_is_null_when_the_environment_has_none_either(
        cache_dir, monkeypatch):
    """edgartools 版本查不到時不讀也不寫。但如果磁碟上的舊檔案記的是
    `edgartools_version: null`，而當下環境也查不到版本，比對會是
    `None != None` → False，檔案會被**誤認為有效**——這是讀側的單獨防線。"""
    # 先用能夠決定版本的環境存檔
    _save_sample()
    # 改檔案上的版本欄位為 null
    path = filing_cache.filing_path("NVDA", ACC)
    raw = json.loads(path.read_text(encoding="utf-8"))
    raw["edgartools_version"] = None
    path.write_text(json.dumps(raw), encoding="utf-8")
    # 環境也查不到版本
    monkeypatch.setattr(filing_cache, "edgartools_version", lambda: None)
    # 讀側要拒絕（不能讓 None != None 的誤認成過）
    assert filing_cache.load_filing("NVDA", ACC, 1045810) is None


# ── fetched_keys：分得清「這張表不存在」與「當初根本沒抓」──────────────────
#
# `dataframes` 裡的 `null` 有兩個意思，光看值分不出來：
#   - 這家公司真的沒有這張表（正常，例如很多公司沒有獨立的綜合損益表）
#   - 存檔當下的 `STATEMENT_KEYS` 還沒有這張表，根本沒去抓
#
# 分不出來的代價在 2026-09-18 付過一次：三張表擴成六張時，舊檔的
# `statement_of_equity` 會被讀成「這家公司沒有股東權益變動表」，不報錯、
# Excel 默默少一張表——只好整庫作廢重抓 11 小時。`fetched_keys` 把「當初抓了
# 哪些」明寫進檔案，以後再擴 `STATEMENT_KEYS` 就不必升 `SCHEMA_VERSION`。


def test_save_records_which_statements_were_fetched(cache_dir):
    """存檔時把當下的 `STATEMENT_KEYS` 明寫進去——這是 `null` 有沒有歧義的關鍵。"""
    _save_sample()
    raw = json.loads(filing_cache.filing_path("NVDA", ACC).read_text(encoding="utf-8"))
    assert raw["fetched_keys"] == list(filing_cache.STATEMENT_KEYS)


def test_load_infers_fetched_keys_for_files_written_before_the_field_existed(cache_dir):
    """14,417 份既有快取都沒有這個欄位，但它們全是 `schema_version=2`，
    而 v2 的定義就是「六張表全抓」——推導得出來就不必升版重抓。

    推導結果要寫回 `entry`，下游才能一律讀 `entry["fetched_keys"]`，
    不必每個呼叫點各自再判一次舊檔。
    """
    _save_sample()
    path = filing_cache.filing_path("NVDA", ACC)
    raw = json.loads(path.read_text(encoding="utf-8"))
    del raw["fetched_keys"]                      # 模擬既有的 14,417 份
    path.write_text(json.dumps(raw), encoding="utf-8")

    entry = filing_cache.load_filing("NVDA", ACC, 1045810)
    assert entry is not None
    assert entry["fetched_keys"] == list(filing_cache.STATEMENT_KEYS)


def test_load_returns_none_when_a_required_statement_was_never_fetched(cache_dir):
    """要一張當初沒抓的表 → 整份視同無快取，照舊連網重抓。

    回 `None` 而不是回一份缺料的 entry：跟既有四道閘同一個路數，呼叫端
    不必多學一種「半有效」狀態。這是未來擴 `STATEMENT_KEYS` 時真正會走到的路。
    """
    _save_sample()
    path = filing_cache.filing_path("NVDA", ACC)
    raw = json.loads(path.read_text(encoding="utf-8"))
    raw["fetched_keys"] = [k for k in raw["fetched_keys"] if k != "cover"]
    path.write_text(json.dumps(raw), encoding="utf-8")

    assert filing_cache.load_filing("NVDA", ACC, 1045810, require_keys=("cover",)) is None
    # 沒被拿掉的那些照樣讀得出來——缺一張不該讓整份報廢
    assert filing_cache.load_filing(
        "NVDA", ACC, 1045810, require_keys=("income_statement",)) is not None


def test_load_requires_only_the_core_statements_by_default(cache_dir):
    """預設只要 IS/BS/CF——日常產 Excel 的路徑不該因為某張 extras 沒抓就重抓。"""
    _save_sample()
    path = filing_cache.filing_path("NVDA", ACC)
    raw = json.loads(path.read_text(encoding="utf-8"))
    raw["fetched_keys"] = list(filing_cache.CORE_STATEMENT_KEYS)
    path.write_text(json.dumps(raw), encoding="utf-8")

    assert filing_cache.load_filing("NVDA", ACC, 1045810) is not None
    # 反面：同一份檔案，明講要 extras 就該擋下來。少了這半條，整個測試在
    # 「還沒實作」跟「預設值寫錯成六張」兩種情況下都會通過，等於沒在測東西。
    assert filing_cache.load_filing(
        "NVDA", ACC, 1045810, require_keys=("cover",)) is None


def test_required_keys_check_does_not_reject_a_negative_cache(cache_dir):
    """負向快取（pre-XBRL）根本沒有 `dataframes`，「哪張表沒抓」對它沒有意義。
    拿 `require_keys` 去擋它會讓那些舊申報每趟都重打 SEC，永遠不收斂。"""
    filing_cache.save_filing(
        "NVDA", ACC, form="10-Q", filing_date="2008-05-01", cik=1045810,
        dataframes=None, has_financials=False)
    entry = filing_cache.load_filing(
        "NVDA", ACC, 1045810, require_keys=filing_cache.STATEMENT_KEYS)
    assert entry is not None
    assert entry["has_financials"] is False


# ── GUI 用的統計與清除 ────────────────────────────────────────────────────────

def test_listing_scans_the_folder_no_global_index_needed(cache_dir):
    """快取了哪些公司直接掃資料夾就知道——不維護容易跟磁碟脫鉤的全域清單。"""
    _save_sample(ticker="NVDA")
    _save_sample(ticker="AMD")
    rows = filing_cache.list_cached_tickers()
    assert {r["ticker"] for r in rows} == {"NVDA", "AMD"}
    assert all(r["count"] == 1 and r["size_bytes"] > 0 for r in rows)


def test_listing_is_sorted_by_size_descending(cache_dir):
    _save_sample(ticker="BIG")
    filing_cache.save_filing("SML", ACC, form="10-Q", filing_date="2025-08-27",
                             cik=1, dataframes=None, has_financials=False)
    rows = filing_cache.list_cached_tickers()
    assert [r["ticker"] for r in rows] == ["BIG", "SML"]


def test_listing_is_empty_when_nothing_is_cached(cache_dir):
    assert filing_cache.list_cached_tickers() == []
    assert filing_cache.total_size_bytes() == 0


def test_listing_survives_a_file_vanishing_mid_scan(cache_dir, monkeypatch):
    """正在被清除、或另一個實例正在寫入時檔案可能瞬間消失。"""
    _save_sample()

    # 窄化 monkeypatch：只讓指定目錄的 stat 失敗，不影響 pytest 內部
    original_stat = Path.stat
    ticker_path = filing_cache.ticker_dir("NVDA")

    def _stat_with_boom(self, *args, **kwargs):
        if self.parent == ticker_path:
            raise OSError("gone")
        return original_stat(self, *args, **kwargs)

    monkeypatch.setattr(Path, "stat", _stat_with_boom)
    rows = filing_cache.list_cached_tickers()
    assert rows == [] or rows[0]["size_bytes"] == 0


def test_total_size_is_the_sum_of_every_ticker(cache_dir):
    _save_sample(ticker="NVDA")
    _save_sample(ticker="AMD")
    rows = filing_cache.list_cached_tickers()
    assert filing_cache.total_size_bytes() == sum(r["size_bytes"] for r in rows)


def test_clear_ticker_removes_the_whole_folder(cache_dir):
    _save_sample(ticker="NVDA")
    _save_sample(ticker="AMD")
    assert filing_cache.clear_ticker("NVDA") is True
    assert not filing_cache.ticker_dir("NVDA").exists()
    assert [r["ticker"] for r in filing_cache.list_cached_tickers()] == ["AMD"]


def test_clear_ticker_on_something_that_is_not_cached_is_harmless(cache_dir):
    assert filing_cache.clear_ticker("ZZZZ") is False


def test_clear_all_removes_every_ticker(cache_dir):
    _save_sample(ticker="NVDA")
    _save_sample(ticker="AMD")
    assert filing_cache.clear_all() == 2
    assert filing_cache.list_cached_tickers() == []


def test_listing_tolerates_one_ticker_vanishing_mid_scan(cache_dir, monkeypatch):
    """一個 ticker 目錄在掃描中被刪掉（concurrent clear 或另一實例 rmtree），
    不該拖垮整份列表——只跳過那一個，其他 ticker 還是要出現。"""
    _save_sample(ticker="NVDA")
    _save_sample(ticker="AMD")

    # 窄化 monkeypatch：只讓 NVDA 目錄的 is_dir() 失敗
    original_is_dir = Path.is_dir
    nvda_path = filing_cache.ticker_dir("NVDA")

    def _is_dir_with_boom(self):
        if self == nvda_path:
            raise OSError("gone")
        return original_is_dir(self)

    monkeypatch.setattr(Path, "is_dir", _is_dir_with_boom)
    rows = filing_cache.list_cached_tickers()
    # NVDA 消失了，但 AMD 還在
    assert [r["ticker"] for r in rows] == ["AMD"]


# ── J9：離線退路用的本地盤點 ──────────────────────────────────────────────

def _write_entry(cache_dir, ticker, accession, *, form, filing_date,
                 schema=None):
    path = cache_dir / ticker / f"{accession}.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({
        "schema_version": filing_cache.SCHEMA_VERSION if schema is None else schema,
        "accession_no": accession, "form": form, "filing_date": filing_date,
        "cached_at": "2026-09-05T00:00:00+08:00", "cik": 320193,
        "edgartools_version": "5.29.0", "has_financials": True,
        "dataframes": {k: None for k in filing_cache.STATEMENT_KEYS},
    }, ensure_ascii=False), encoding="utf-8")


def test_list_cached_filings_returns_newest_first(cache_dir):
    _write_entry(cache_dir, "NVDA", "0000320193-20-000001",
                 form="10-Q", filing_date="2020-05-01")
    _write_entry(cache_dir, "NVDA", "0000320193-26-000002",
                 form="10-Q", filing_date="2026-05-01")
    rows = filing_cache.list_cached_filings("NVDA")
    assert [r["filing_date"] for r in rows] == ["2026-05-01", "2020-05-01"]
    assert rows[0]["accession"] == "0000320193-26-000002"


def test_list_cached_filings_can_filter_by_form(cache_dir):
    _write_entry(cache_dir, "NVDA", "0000320193-26-000001",
                 form="10-Q", filing_date="2026-05-01")
    _write_entry(cache_dir, "NVDA", "0000320193-26-000002",
                 form="10-K", filing_date="2026-02-01")
    assert [r["form"] for r in filing_cache.list_cached_filings("NVDA", form="10-K")] \
        == ["10-K"]


def test_list_cached_filings_skips_stale_schema(cache_dir):
    """舊 schema 在 `load_filing()` 也讀不出來，列進清單只會讓離線退路
    拿到一堆叫不出內容的項目。"""
    _write_entry(cache_dir, "NVDA", "0000320193-26-000001",
                 form="10-Q", filing_date="2026-05-01", schema=0)
    assert filing_cache.list_cached_filings("NVDA") == []


def test_list_cached_filings_ignores_the_meta_file(cache_dir):
    _write_entry(cache_dir, "NVDA", "0000320193-26-000001",
                 form="10-Q", filing_date="2026-05-01")
    (cache_dir / "NVDA" / "_meta.json").write_text("{}", encoding="utf-8")
    assert len(filing_cache.list_cached_filings("NVDA")) == 1


def test_list_cached_filings_is_empty_for_an_unknown_ticker(cache_dir):
    assert filing_cache.list_cached_filings("NOPE") == []


# ── 6-K 的 R 檔判定快取（TODO D9 A 路線）─────────────────────────────────
#
# 外國私人發行人的 6-K 是大雜燴：ARM 33 份裡只有 9 份有財報。判準是
# `R*.htm` 的數量 > 1，但問一份 6-K 有幾個 R 檔要一次 index 請求——
# Sony 有 1,055 份 6-K，每輪重問等於 1,055 次請求。所以判定結果要落檔，
# 而且**跟 filing 快取分開存**（判定不是財報內容，不該讓 `file_count` 變動）。

def test_sixk_probe_roundtrips_the_r_file_count_per_accession(cache_dir):
    filing_cache.save_sixk_probe("ARM", {"0001045810-25-000123": 68,
                                         "0001045810-25-000124": 1})
    assert filing_cache.load_sixk_probe("ARM") == {
        "0001045810-25-000123": 68,
        "0001045810-25-000124": 1,
    }


def test_sixk_probe_is_empty_for_a_ticker_that_was_never_probed(cache_dir):
    assert filing_cache.load_sixk_probe("ARM") == {}


def test_sixk_probe_falls_back_to_empty_when_the_file_is_corrupt(cache_dir):
    """半截檔不能讓抓取炸掉——判定沒了就重問一次，成本是請求不是正確性。"""
    (cache_dir / "ARM").mkdir(parents=True)
    (cache_dir / "ARM" / filing_cache.SIXK_PROBE_FILENAME).write_text(
        '{"0001045810-25-0001', encoding="utf-8")
    assert filing_cache.load_sixk_probe("ARM") == {}


def test_sixk_probe_drops_entries_that_are_not_accession_to_int(cache_dir):
    """手改過／舊版寫的髒資料一律不採信，回頭重問。"""
    (cache_dir / "ARM").mkdir(parents=True)
    (cache_dir / "ARM" / filing_cache.SIXK_PROBE_FILENAME).write_text(
        json.dumps({"0001045810-25-000123": 68, "junk": 5,
                    "0001045810-25-000124": "many"}), encoding="utf-8")
    assert filing_cache.load_sixk_probe("ARM") == {"0001045810-25-000123": 68}


def test_sixk_probe_file_is_not_counted_as_a_cached_filing(cache_dir):
    """跟 `_meta.json` 同一個理由：它在 filing 資料夾裡，但它不是 filing。
    算進去會讓 `_meta_is_reusable()` 的 `file_count` 對不上，每次都重建 meta。"""
    _write_entry(cache_dir, "ARM", "0001045810-25-000123",
                 form="6-K", filing_date="2026-07-29")
    filing_cache.save_sixk_probe("ARM", {"0001045810-25-000124": 1})
    assert filing_cache._dir_stats(cache_dir / "ARM")[0] == 1
    assert len(filing_cache.list_cached_filings("ARM")) == 1
