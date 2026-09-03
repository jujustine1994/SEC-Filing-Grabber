"""
fetcher_gaap.py — Fetch XBRL GAAP financial statements from SEC EDGAR via edgartools.

Fetches quarterly data from up to `max_filings` 10-Q filings (newest first),
and annual data from up to `max_annual_filings` 10-K filings (newest first).

Public API:
    fetch_gaap_statements(ticker, identity, max_filings=80, max_annual_filings=20) -> list[StatementTable]

Sheet outputs:
    Data_Financials(Q) — quarterly IS + BS + CF merged (from 10-Q)
    Data_Financials(Y) — annual IS + BS + CF merged (from 10-K)
    Data_Seg_*         — one sheet per IS concept that has segment breakdowns
    Data_Meta          — ticker / company / date / quarter count

StatementTable layout (A / B / C+):
    Col A  = Std Name (standardised display label)
    Col B  = Original Item (company's XBRL label from edgartools)
    Col C+ = quarterly data, oldest → newest
"""

from __future__ import annotations

import math
import re
import sys
import time
import unicodedata
from dataclasses import dataclass, field, replace
from datetime import date, date as _date
from typing import Any, Callable

import pandas as pd
from edgar import Company, set_identity as set_identity

from contextlib import contextmanager
from contextvars import ContextVar

import filing_cache
from fetch_ledger import FetchLedger
from i18n import t
from net_retry import NetworkDownError, with_retry
from override_engine import load_overrides, run_diagnosis, check_key_rows

# 9 個關鍵科目的清單（品質檢查用）。定義在 excel_formatter，避免兩處各記一份。
_ALL_KEY_ROWS_LAZY = None


# ── Data contract ────────────────────────────────────────────────────────

@dataclass
class StatementTable:
    """One financial statement pre-structured for Excel output."""
    sheet_name:     str
    quarter_labels: list[str]
    filing_dates:   list[str]
    concepts:       list[str]
    values:         list[list[Any]]
    ticker:         str = ""
    labels:         list[str] = field(default_factory=list)   # B-col: original XBRL labels
    # 各期真正的期末結算日（YYYY-MM-DD）。美股多為 52/53 週制，期末日不是月底
    # （ARLO FY2026Q1 結束在 03-29 不是 03-31），從財季標籤反推只能得到月份。
    # 來源是 XBRL 欄名 "2026-03-29 (Q1)"。沒帶到的表留空，下游自行退回反推。
    period_ends:    list[str] = field(default_factory=list)


# ── Constants ─────────────────────────────────────────────────────────────

META_COLS: set[str] = {
    'concept', 'label', 'standard_concept', 'level', 'abstract',
    'dimension', 'is_breakdown', 'dimension_axis', 'dimension_member',
    'dimension_member_label', 'dimension_label', 'unit', 'point_in_time',
    'balance', 'weight', 'preferred_sign', 'parent_concept', 'parent_abstract_concept',
}

# EDGAR began requiring XBRL from ~2009; filings before this date have no XBRL data.
# Filing lists are returned newest-first, so hitting a pre-cutoff filing means we can
# break the loop immediately rather than continuing through decades of empty filings.
_XBRL_CUTOFF: _date = _date(2008, 1, 1)


# ── 缺漏帳本（CTH 2026-08-17 定案）────────────────────────────────────────
#
# 抓不到的期數要被記下來並在最後講出來，但**不中止抓取**——「抓得太嚴格
# 讓資料永遠抓不出來」比缺幾期更糟。完整的演進理由見 fetch_ledger.py。
#
# 用 ContextVar 而不是傳參數：帳本要穿過九個抓取函式，全部改簽名會動到
# 一堆既有測試，而這件事跟那些函式的職責無關。ContextVar 每個執行緒各自
# 一份，GUI 的背景 worker 與 CLI 都正確。
_ledger_var: ContextVar[FetchLedger | None] = ContextVar("_fetch_ledger", default=None)


@contextmanager
def collect_gaps(ledger: FetchLedger | None = None):
    """開一本帳本涵蓋這一趟抓取。`with collect_gaps() as led: ...`"""
    led = ledger if ledger is not None else FetchLedger()
    token = _ledger_var.set(led)
    try:
        yield led
    finally:
        _ledger_var.reset(token)


def _ledger() -> FetchLedger | None:
    return _ledger_var.get()


def _note_gap(where, exc: BaseException) -> None:
    """記一期沒抓到。沒開帳本時（單獨呼叫某個 builder 的測試）靜靜略過。"""
    led = _ledger()
    if led is not None:
        led.record(str(where), exc)


def _note_ok() -> None:
    led = _ledger()
    if led is not None:
        led.succeeded()


# ── 缺漏自動重試（D11-B，2026-08-24 CTH 決定要做）───────────────────────────
#
# 帳本有網路造成的缺漏時（`network_blamed`），退避幾秒後整趟重跑一次、
# 逐格合併：原本有值的格子保留，None 的格子用重試那輪的值補。只重試一次，
# 不遞迴重試到底——網路真的斷線時不該乾等。`data` 類缺漏不重試，SEC 有
# 回應代表那是資料本身的性質，重試沒用（跟 `fetch_ledger.py` 既有的分類
# 判斷同一套邏輯，不重新發明）。
#
# 整趟重跑而不是只重試失敗的那幾份 filing：帳本現在只記期間標籤跟例外
# 分類，沒存 filing 物件本身，要做到「只補失敗那幾份」得動 7 個
# `_note_gap()` 呼叫點的呼叫鏈，範圍與風險都大得多。單一 ticker 重跑一次
# 多花的時間可接受，因為只有真的撞到 `network_blamed` 才會觸發。

_RETRY_BACKOFF_SECONDS = 5.0


def _merge_retry_tables(orig: list[StatementTable],
                         retry: list[StatementTable]) -> list[StatementTable]:
    """逐格合併：`orig` 有值的格子保留，`None` 的格子用 `retry` 同位置的值補。

    用 `sheet_name` 配對；`quarter_labels`／`concepts` 結構對不上（理論上
    不該發生，同一組參數兩輪本該一致）就整張表保留原樣，不硬湊、不猜。
    """
    retry_by_name = {t.sheet_name: t for t in retry}
    merged = []
    for tbl in orig:
        rt = retry_by_name.get(tbl.sheet_name)
        if (rt is None
                or rt.quarter_labels != tbl.quarter_labels
                or rt.concepts != tbl.concepts):
            merged.append(tbl)
            continue
        new_values = [
            [o if o is not None else r for o, r in zip(orig_row, retry_row)]
            for orig_row, retry_row in zip(tbl.values, rt.values)
        ]
        merged.append(replace(tbl, values=new_values))
    return merged


def _patch_meta_gap_note(tables: list[StatementTable], led: "FetchLedger") -> None:
    """重試吸收帳本之後，把 `Data_Meta` 的 `Fetch Gaps` 欄位改成最新的
    `led.summary()`——不然 Index 還是會顯示重試前的舊缺漏數。"""
    note = led.summary() or t("xls.meta.none")
    for tbl in tables:
        if tbl.sheet_name == "Data_Meta" and "Fetch Gaps" in tbl.concepts:
            idx = tbl.concepts.index("Fetch Gaps")
            tbl.values[idx] = [note] * len(tbl.values[idx])


def _fetch_with_retry(
    tables: list[StatementTable],
    led: FetchLedger,
    retry_once: Callable[[], tuple[list[StatementTable], FetchLedger]],
    sleep: Callable[[float], None] = time.sleep,
) -> list[StatementTable]:
    """第一輪 `tables`／`led` 已經跑完了；帳本有網路造成的缺漏就退避重試一次並合併。

    `led` 是呼叫端當下在用的那本帳本（自己開的或外層 caller 開的都一樣）——
    這個函式不負責開第一輪的帳本，只決定要不要重試、重試完怎麼合併回去。
    `retry_once` 要在自己獨立的一輪 `collect_gaps()` 範圍裡重跑一次、回傳
    `(tables, ledger)`；純函式不碰網路，測試把它換成假的即可。
    """
    if led.network_blamed:
        sleep(_RETRY_BACKOFF_SECONDS)
        try:
            retry_tables, retry_led = retry_once()
        except Exception:
            # 重試路徑本身出錯不該把「第一輪已經抓到大半資料」變成整趟
            # 失敗——退回第一輪的結果，帳本維持原本記的缺漏就好。
            return tables
        tables = _merge_retry_tables(tables, retry_tables)
        led.absorbed_by_retry(retry_led)
        _patch_meta_gap_note(tables, led)
    return tables


# ── 抓取進度回報（2026-08-18，TODO E12）────────────────────────────────────
#
# GAAP 抓取耗時過久時看起來像卡死——`fetch_gaap_statements` 一份 filing 要
# 分別建 IS/BS/CF 三張表，各自對同一批 filings 各跑一輪、各發一次網路請求，
# 幾十份 filing 跑下來可能要幾分鐘，中途 GUI 進度條完全不動。
#
# 沿用帳本（`collect_gaps`）同一套 ContextVar 手法，不改 `_build_is_table`／
# `_build_bs_table`／`_build_cf_table` 的函式簽名——這三個函式已經在每份
# filing 處理完（不管成功失敗）呼叫一次 `_note_ok()`/`_note_gap()`，這裡搭
# 同一個掛鉤點回報進度，不用另外傳參數穿過呼叫鏈。
class _ProgressState:
    __slots__ = ("cb", "total", "current")

    def __init__(self, cb):
        self.cb = cb
        self.total = 0
        self.current = 0

    def tick(self, label: str = "") -> None:
        self.current += 1
        if self.cb is None:
            return
        try:
            # current 可能因為早停（`max_filings` 上限、pre-XBRL 篩掉）跑不到
            # total——夾住上限，不然進度條會顯示超過 100%
            self.cb(min(self.current, self.total), self.total, label)
        except Exception:
            pass  # 進度回報是錦上添花，回呼本身出錯不能拖垮抓取


_progress_var: ContextVar[_ProgressState | None] = ContextVar("_fetch_progress", default=None)


@contextmanager
def report_progress(cb):
    """開一個進度回報範圍。`cb(current, total, label)` 在每處理完一份 filing
    （建 IS/BS/CF 表其中一步）時被呼叫一次。`cb=None` 時整段是 no-op，呼叫端
    不用另外判斷要不要包這層。
    """
    state = _ProgressState(cb) if cb is not None else None
    token = _progress_var.set(state)
    try:
        yield
    finally:
        _progress_var.reset(token)


def _set_progress_total(n: int) -> None:
    """列完 filings、知道實際要處理幾份之後呼叫一次，設定進度條的分母。"""
    state = _progress_var.get()
    if state is not None:
        state.total = max(n, 1)


def _tick_progress(label: str = "") -> None:
    state = _progress_var.get()
    if state is not None:
        state.tick(label)


def _filing_ref(filing) -> str:
    """這份 filing 拿來給人看的稱呼。

    失敗當下還不知道財季標籤——那是從抓下來的 DataFrame 欄名推出來的，
    而我們正是沒抓到。退而求其次用期末日（`period_of_report`），對使用者
    來說一樣認得出是哪一期。
    """
    for attr in ("period_of_report", "filing_date"):
        value = getattr(filing, attr, None)
        if value:
            return str(value)
    return "?"


# ── 解析快取（G9）─────────────────────────────────────────────────────────
#
# IS/BS/CF/segments 四個 build pass 各自對**同一批 filing** 重新解析一次。
# 實測 ARLO（25 份 filing、66 秒）：`_filing_obj` 被呼叫 96 次（每份 3.8 次），
# 其中 `financials`（真正的 XBRL 解析）花 19.9 秒、`to_dataframe` 花 28.4 秒。
#
# edgartools **不會跨呼叫快取解析結果**——同一支 ticker 在同一個 process 連跑
# 兩次是 64.5s vs 67.3s，完全沒變快。所以這個重複成本是真的。
#
# 快取的生命週期**只能是一次抓取**（`_parse_cache_scope()`）：跨 ticker 殘留會
# 吃掉別家的資料，跨執行殘留會拿到過期的申報。沒開範圍時照常運作、不快取。
_parse_cache: dict | None = None


@contextmanager
def _parse_cache_scope():
    """一次抓取的解析快取範圍。離開時清空，不讓資料跨 ticker／跨執行殘留。"""
    global _parse_cache
    outer = _parse_cache
    _parse_cache = {} if outer is None else outer
    try:
        yield
    finally:
        if outer is None:
            _parse_cache = None


def _cache_key(filing) -> str | None:
    """filing 的快取鍵。拿不到 accession 就不快取（回 None）——寧可慢也不要錯拿。"""
    acc = getattr(filing, "accession_no", None)
    return str(acc) if acc else None


# ── 本地磁碟快取（跨執行有效）────────────────────────────────────────────
#
# 跟上面的 `_parse_cache`（G9，只活在一次執行的記憶體裡）是**兩層不同的快取**，
# 不衝突：磁碟快取讀回來的結果一樣會進記憶體快取，避免同一次執行內重複讀檔。
#
# 綁定分兩步是為了避免把 `_fetch_gaap_impl()` 整段重新縮排：範圍在
# `fetch_gaap_statements()` 開，`ticker`/`cik` 等 `Company(ticker)` 建好之後
# 才由 `_bind_disk_cache()` 填進去。沒綁定（拿不到 cik）就整個不用快取，
# 行為跟改動前一模一樣。
_disk_cache: dict | None = None
_last_cache_stats: tuple[int, int] = (0, 0)


@contextmanager
def _disk_cache_scope():
    """一次抓取的磁碟快取範圍。離開時更新 manifest 並記下命中統計。"""
    global _disk_cache, _last_cache_stats
    outer = _disk_cache
    ctx = {"ticker": None, "cik": None, "hits": 0, "misses": 0} if outer is None else outer
    _disk_cache = ctx
    try:
        yield ctx
    finally:
        if outer is None:
            if ctx["ticker"]:
                # 收尾重建索引：反映這趟跑完後磁碟上實際的內容。
                # manifest 只是給 GUI 看的，重建失敗不影響任何抓取結果。
                try:
                    filing_cache.rebuild_manifest(ctx["ticker"], ctx["cik"])
                except OSError:
                    pass
            _last_cache_stats = (ctx["hits"], ctx["hits"] + ctx["misses"])
            _disk_cache = None


def _bind_disk_cache(ticker: str, cik) -> None:
    """把這趟的公司身分填進磁碟快取範圍。cik 拿不到就不啟用快取——
    ticker 只是別名，會換手；cik 才是跟 SEC 打交道真正的鍵。"""
    if _disk_cache is None or not ticker or cik is None:
        return
    try:
        _disk_cache["cik"] = int(cik)
    except (TypeError, ValueError):
        return
    _disk_cache["ticker"] = str(ticker).strip().upper()


def last_cache_stats() -> tuple[int, int]:
    """上一趟抓取的 (命中份數, 處理份數)。給 log 用（見 main.py）。"""
    return _last_cache_stats


def _save_to_disk_cache(ctx: dict, filing, obj) -> None:
    """把剛解析出來的三張 DataFrame 逐份即時落檔。

    ⚠ 只有「`filing.obj()` 成功回來」才會走到這裡——網路失敗會在上面直接
    往外拋，不留任何快取（正向或負向都不留），下次照樣重試。
    解析成功但 `financials` 是 None（pre-XBRL）才寫負向快取。
    """
    acc = _cache_key(filing)
    fd = getattr(filing, "filing_date", None)
    meta = {
        "form": str(getattr(filing, "form", "") or ""),
        "filing_date": str(fd) if fd else "",
        "cik": ctx["cik"],
    }
    fin = _financials_of(obj)
    if fin is None:
        filing_cache.save_filing(ctx["ticker"], acc, dataframes=None,
                                 has_financials=False, **meta)
        return
    try:
        dfs = {}
        for key, getter in (("income_statement", fin.income_statement),
                            ("balance_sheet", fin.balance_sheet),
                            ("cashflow_statement", fin.cashflow_statement)):
            stmt = getter()
            dfs[key] = None if stmt is None else stmt.to_dataframe()
    except Exception:
        return   # 解析失敗不留快取，下次重試（跟網路失敗同一個處理）
    filing_cache.save_filing(ctx["ticker"], acc, dataframes=dfs,
                             has_financials=True, **meta)


def _list_filings(company, form: str, amendments: bool = False,
                   sleep: Callable[[float], None] = time.sleep) -> list:
    """列出這家公司某個表單類型的所有 filing，網路問題退避重試（2026-08-25）。

    這一步比逐份 filing 更早——`_filing_obj()` 抓內容早就有 `with_retry`
    保護瞬斷，但列清單這一步完全沒被蓋到。2026-08-25 實測 201 家重建撞到
    6 家逾時，全部發生在這裡：一次瞬斷直接讓整趟 `fetch_gaap_statements()`
    拋例外，連 D11-B 的缺漏帳本都還沒開始記，重試不到。
    """
    return with_retry(lambda: list(company.get_filings(form=form, amendments=amendments)),
                       sleep=sleep)


def _filing_obj(filing):
    """下載並解析一份 filing，網路問題會退避重試（CTH 2026-08-17）。

    重試是為了救閃斷——一次瞬斷不該讓那一季永久消失。救不回來就拋
    `NetworkDownError`，呼叫端記進帳本後**繼續抓下一期**，不中止。

    帳本判定網路已經斷掉之後（連續多期失敗）就不再退避重試：整個網路
    斷掉時 40 份財報各等 2+4 秒等於乾等 4 分鐘，剩下的快速失敗完，
    照樣寫檔、照樣提示。

    快取有兩層：記憶體（G9，一次執行內）與磁碟（跨執行）。磁碟命中時回傳的
    是 `filing_cache` 的替身物件，只實作 `.financials` 那一條鏈——四個 builder
    對 filing 物件的用法就只有這一種。
    """
    key = _cache_key(filing)
    if _parse_cache is not None and key is not None and key in _parse_cache:
        return _parse_cache[key]

    ctx = _disk_cache
    if ctx is not None and ctx["ticker"] and key is not None:
        entry = filing_cache.load_filing(ctx["ticker"], key, ctx["cik"])
        if entry is not None:
            ctx["hits"] += 1
            obj = filing_cache.cached_filing(entry)
            if _parse_cache is not None:
                _parse_cache[key] = obj
            return obj

    led = _ledger()
    attempts = 1 if (led is not None and led.give_up_retrying) else 3
    obj = with_retry(lambda: filing.obj(), attempts=attempts)

    if ctx is not None and ctx["ticker"] and key is not None:
        ctx["misses"] += 1
        _save_to_disk_cache(ctx, filing, obj)

    if _parse_cache is not None and key is not None:
        _parse_cache[key] = obj
    return obj


def _financials_of(tenq):
    """取 filing 物件上的 financials，沒有就回 None。

    舊申報（2009 年前不強制 XBRL）解出來的物件 `financials` 會是 None，
    再對它取 `.income_statement()` 就是 AttributeError。那不是錯誤而是
    「這份沒有可解的資料」，明講出來比讓它拋例外再被攔下來清楚。
    """
    return getattr(tenq, "financials", None)


def _filter_filings_by_year(
    filings: list,
    start_year: int | None,
    end_year: int | None,
) -> list:
    """Filter filings list to only those within [start_year, end_year] (inclusive).

    Handles both date objects and ISO date strings ('YYYY-MM-DD').
    Returns filings unchanged when both bounds are None.
    """
    if start_year is None and end_year is None:
        return filings
    result = []
    for f in filings:
        fd = getattr(f, "filing_date", None)
        if fd is None:
            result.append(f)
            continue
        year = fd.year if isinstance(fd, _date) else int(str(fd)[:4])
        if start_year is not None and year < start_year:
            continue
        if end_year is not None and year > end_year:
            continue
        result.append(f)
    return result

# Tuple: (label, std_concept, fallback_suffix, source, match, label_hint)
#   label         — display name (Col A)
#   std_concept   — primary: standard_concept == value
#   fallback_suffix — secondary: concept contains value
#   source        — "IS" | "CF" | "BS" | "DERIVED"
#   match         — "first" | "last"  (which occurrence when multiple rows match)
#   label_hint    — filter: 該優先層的候選裡只留 label 含這個字串的（濾空整層跳過）
#   label_fallback— 第三層：concept 兩層都比不到時，改用 label 比對。
#                   **公司自訂延伸 tag（`nvda_...`）唯一抓得到的方式**——那種 concept
#                   名字每家自己取，只有 label 對得上。寫得下去就要窄，第三層沒有
#                   任何東西再擋它了（見 H4 spec）
# ── H6（2026-08-25）：四條 label_hint 的措辭清單 ────────────────────────────
#
# 201 家最新 10-Q 實測掃出來的（`scripts/diag_hintsweep.py`，原始輸出與人工分類在
# `output/_hintsweep_201/`）。下面每個措辭都對應到具體公司，不是想像出來的。
# **hint 只在該優先層有候選時才過濾**，濾空整層就跳過——所以放太寬會吃到錯的列，
# 放太窄會整家全損。這四條都是「concept 層本來就對得上、純粹被 hint 卡住」。

# Capex：14 家全損（AEP/AMP/APD/AXP/BK/COF/F/GIS/HSY/ITW/KMB/MAR/PEP/UNP），
# 它們寫 Capital spending／Capital investments／Purchases of premises and equipment，
# 舊 hint 的 `propert|capital expenditure` 兩個詞根一個都不含。
# ⚠ 前面那段 negative lookahead 是必要的：UNP／AMD 的現金流量表底下另有一列
# `CapitalExpendituresIncurredButNotYetPaid`（std_concept 同樣是 `CapitalExpenses`），
# 那是非現金揭露，**加總會重複計算**。
_CAPEX_HINT = (
    r"^(?!.*(?:accrued|not yet paid|payable))"
    r".*(?:propert|capital expenditure|capital spending|capital investment"
    r"|capital addition|capital and technology|premises and equipment"
    r"|plant and equipment|land, buildings|generation facilities)"
)

# Cash：5 家全損（ETN/APD/IP/KR/SLB）。std_concept 多數已經正確命中
# `CashAndCashEquivalentsAtCarryingValue`，純粹被措辭卡掉。
# ⚠ 不可以放寬到吃進銀行的 `CashAndDueFromBanks`（AXP/BAC/BK/C/COF/JPM/WFC 7 家）
# ——那是概念取捨題，CTH 還沒決定（TODO H6），H6 刻意不動。
_CASH_HINT = r"cash and (?:cash )?equivalents|^cash\s*$|cash and cash items|cash and temporary"

# 第三層（label 比對）。ASU 2016-18 只要求**現金流量表**的期初期末總額包含受限
# 現金，**資產負債表沒有要求合併列示**——多數公司 BS 仍分開列、附註做 reconciliation。
# INTC 2022~2025 那 15 期就是這樣：BS 印的字是「Cash and cash equivalents」，
# 但 tag 挑了 ASU 2016-18 的合併 element
# `CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents`
# （⚠ 名字裡沒有 "And"，所以 std 與 fallback_suffix 兩層都比不中）。
#
# 抓「公司自己印在報表表面的那一行」是對的口徑；真的把受限現金併進列示的公司，
# label 會寫成「Cash, cash equivalents and restricted cash」之類，**要窄到吃不到
# 那種**——第三層後面沒有任何東西再擋它。
_CASH_LABEL_FALLBACK = r"^cash and cash equivalents$"

# Common Stock & APIC：12 家全損。愛爾蘭／英國／瑞士註冊或改遷冊的公司寫
# Ordinary shares（ACN/AON/ETN/JCI/LIN/MDT），另一批寫 Common shares
# （ABT/AMP/AXP/CB/KR/UNP）。concept 全部是 `CommonStockValue(Outstanding)`、
# std_concept 全部是 `CommonEquity`——分類表原本說這批是「concept 層先失守」是
# 錯的（2026-08-25 用 diag_rowprobe.py 逐家查過原始 10-Q）。
#
# ⚠ 要排除庫藏股列：LIN／ABT／AMP／KR 的 `TreasuryStockCommonValue` 的
# std_concept 同樣是 `CommonEquity`（實測 LIN [28]、ABT [51]、AMP [77]、KR [37]）。
# **但不可以「label 含 treasury 就踢掉」**——NSC 的普通股列自己就寫
# 「Common stock, net of treasury shares」，那樣寫會誤傷它（201 家重掃實際踩到）。
# 真正的庫藏股列都帶「at cost」或「in treasury」，用那個當判準。
_TREASURY_ROW = r"(?:.*\bin treasury\b)|(?:.*treasury.{0,40}at cost)|(?:\s*less[:\s\-]+treasury\b)|(?:\s*treasury\b)"
_COMMON_STOCK_HINT = (
    rf"^(?!{_TREASURY_ROW})"
    r".*(?:common stock|paid-in capital|ordinary shares|common shares)"
)

# Cost of Revenue：36 家被擋，其中**只有 6 家是真缺口**——CVX/COP/PSX（採購原油
# 商品）、AEP/EXC（採購電力燃料）、CMG（食材包材）。其餘 29 家是銀行／保險／
# 交易所／鐵路／REIT，概念上本來就沒有 COGS（與 D8 同一類），維持空白才對。
# ⚠ 所以只加「purchased」與 CMG 那個措辭，**不可以放寬到吃進
# `LaborAndRelatedExpense`（Compensation and benefits／Labor and Fringe）**。
_COGS_HINT = r"cost|^purchased|food, beverage"

# D&A（`CF_TEMPLATE` 的 `D&A` 與 `IS_TEMPLATE` 的 `D&A (CF memo)` 共用，兩列取的
# 是同一個東西）。2026-08-25（G10）實測 AMD／MRVL：
#
#   concept = us-gaap_OtherDepreciationAndAmortization
#   standard_concept = NonoperatingIncomeExpense   ← edgartools 標錯，跟折舊攤銷無關
#   label = "Depreciation and amortization"
#
# 第一層（std_concept）被標錯的值擋掉，第二層舊的 `DepreciationDepletionAnd
# Amortization` 又比不中那個 concept 名字 → 整列全損（AMD 2/25 期、MRVL 0/19 期）。
# 放寬成整個 Depreciation…Amortization 家族，把 `OtherDepreciationAndAmortization`
# 與 `DepreciationAmortizationAndAccretionNet` 都收進來。
_DA_FALLBACK = r"Depreciation\w*Amortization"

# 第三層：公司自訂延伸 tag 只有 label 對得上（TSLA 用
# `tsla_DepreciationAmortizationAndImpairment`，std_concept 是 nan）。
# ⚠ 一定要 `^depreciation` 開頭，不可以只寫 `amortization`——現金流量表上另外還有
# 「Amortization of acquisition-related intangibles」（AMD/MRVL 都有獨立一列）、
# 債務發行成本攤銷、遞延佣金攤銷，吃進來會變成拿無形資產攤銷當 D&A。
_DA_LABEL_FALLBACK = r"^depreciation"

_T = tuple[str, str | None, str, str, str, str | None, str | None]

IS_TEMPLATE: list[_T] = [
    ("Revenue",                    "Revenue",                        r"RevenueFromContractWithCustomer|SalesRevenueNet|SalesRevenueGoodsNet|_Revenues$|^Revenues$", "IS", "first", None, None),
    ("Cost of Revenue",            "CostOfGoodsAndServicesSold",     "CostOfGoodsSold",                                       "IS", "first", _COGS_HINT, None),
    ("Gross Profit",               "GrossProfit",                    "GrossProfit",                                            "IS", "first", None, None),
    ("R&D Expense",                "ResearchAndDevelopmentExpenses", "ResearchAndDevelopment",                                 "IS", "first", None, None),
    ("SG&A Expense",               "SellingGeneralAndAdminExpenses", "SellingGeneralAndAdmin",                                 "IS", "first", None, None),
    ("D&A (CF memo)",              "DepreciationExpense",            _DA_FALLBACK,                                             "CF", "first", None, _DA_LABEL_FALLBACK),
    ("Other Operating Expense",    None,                             "OtherCostAndExpenseOperating|OtherOperatingIncomeExpenseNet|OtherOperatingExpense", "IS", "first", None, None),
    ("Total Operating Expense",    "TotalOperatingExpenses",         "OperatingExpenses",                                      "IS", "first", None, None),
    ("Total Costs and Expenses",   None,                             "^us-gaap_CostsAndExpenses$",                             "IS", "first", None, None),
    ("Operating Income",           "OperatingIncomeLoss",            "OperatingIncomeLoss",                                    "IS", "first", None, None),
    ("Interest Expense",           "InterestExpense",                "InterestExpense",                                        "IS", "first", None, None),
    ("Interest Income",            "InterestIncome",                 r"(?<!Non)InterestIncome(?!Expense)|InvestmentIncomeInterest", "IS", "first", None, None),
    ("Other Non-op Inc/(Exp)",     None,                             "OtherNonoperatingIncome",                                "IS", "first", None, None),
    ("Total Non-op Income/(Loss)", "NonoperatingIncomeExpense",      "NonoperatingIncome",                                     "IS", "first", None, None),
    ("Pre-tax Income",             "PretaxIncomeLoss",               "IncomeLossFromContinuingOperationsBeforeIncomeTax",       "IS", "first", None, None),
    ("Income Tax",                 "IncomeTaxes",                    "IncomeTaxExpense",                                       "IS", "first", None, None),
    ("Net Income",                 "NetIncome",                      "NetIncomeLoss|NetIncomeLossAttributableToParent",         "IS", "first", None, None),
    ("Minority Interest",          None,                             "NetIncomeLossAttributableToNoncontrollingInterest",       "IS", "first", None, None),
    # 含少數股權的淨利。有 NCI 結構的公司會把「合併淨利」與「歸屬母公司淨利」分開報，
    # 上面的 Net Income 只認 NetIncomeLoss（歸屬母公司），這一列補的是合併數。
    ("Net Income incl. NCI",       None,                             "^us-gaap_ProfitLoss$",                                   "IS", "first", None, None),
    ("SBC",                        "StockBasedCompensationExpense",  "ShareBasedCompensation",                                 "CF", "first", None, None),
    ("Basic EPS",                  None,                             "EarningsPerShareBasic",                                  "IS", "first", None, None),
    ("Diluted EPS",                None,                             "EarningsPerShareDiluted",                                "IS", "first", None, None),
    ("Basic Shares",               "SharesAverage",                  "WeightedAverageNumberOfSharesOutstandingBasic",          "IS", "first", None, None),
    ("Diluted Shares",             "SharesFullyDilutedAverage",      "WeightedAverageNumberOfDilutedSharesOutstanding",        "IS", "first", None, None),
]

BS_TEMPLATE: list[_T] = [
    # ── Assets ──────────────────────────────────────────────────────────
    ("Cash",                           "CashAndMarketableSecurities",             "CashAndCashEquivalents",                                    "BS", "first", _CASH_HINT, _CASH_LABEL_FALLBACK),
    ("Short-term Investments",         "ShortTermInvestments",                    "ShortTermInvestments",                                      "BS", "first", None, None),
    ("Accounts Receivable",            "TradeReceivables",                        "AccountsReceivable",                                        "BS", "first", "receivable", None),
    ("Inventories",                    "Inventories",                             "Inventories",                                               "BS", "first", None, None),
    ("Other Current Assets",           "OtherNonOperatingCurrentAssets",          "OtherCurrentAssets",                                        "BS", "first", "other", None),
    ("Total Current Assets",           "CurrentAssetsTotal",                      "AssetsCurrent",                                             "BS", "first", None, None),
    ("PP&E, net",                      "PlantPropertyEquipmentNet",               "PropertyPlantAndEquipmentNet",                              "BS", "first", None, None),
    ("Operating Lease ROU Assets",     "OperatingLeaseRightOfUseAsset",           "OperatingLeaseRightOfUseAsset",                             "BS", "first", None, None),
    ("Long-term Investments",          "LongtermInvestments",                     "LongTermInvestments",                                       "BS", "first", None, None),
    ("Goodwill",                       "Goodwill",                                "Goodwill",                                                  "BS", "first", None, None),
    ("Intangible Assets, net",         "IntangibleAssets",                        "IntangibleAssetsNet",                                       "BS", "first", None, None),
    ("Deferred Tax Assets",            "DeferredTaxNoncurrentAssets",             "DeferredIncomeTaxAssetsNet",                                "BS", "first", None, None),
    ("Other Non-current Assets",       "OtherNonOperatingNonCurrentAssets",       "OtherAssetsNoncurrent",                                     "BS", "last",  "other|miscellaneous", None),
    ("Total Non-current Assets",       None,                                      "^us-gaap_AssetsNoncurrent$",                                "BS", "first", None, None),
    ("Total Assets",                   "Assets",                                  "Assets",                                                    "BS", "last",  None, None),
    # ── Liabilities ─────────────────────────────────────────────────────
    ("Accounts Payable",               "TradePayables",                           "AccountsPayable",                                           "BS", "first", None, None),
    ("Short-term Debt",                "ShortTermDebt",                           "ShortTermBorrowings",                                       "BS", "first", None, None),
    ("Current Portion of LT Debt",     "CurrentPortionOfLongTermDebt",            "LongTermDebtCurrent",                                       "BS", "first", None, None),
    ("Op. Lease Liabilities, current", "OperatingLeaseCurrentDebtEquivalent",     "OperatingLeaseLiabilityCurrent",                            "BS", "first", None, None),
    ("Accrued Compensation",           "AccruedCompensation",                     "EmployeeRelatedLiabilitiesCurrent",                         "BS", "first", None, None),
    ("Deferred Revenue, current",      None,                                      "ContractWithCustomerLiabilityCurrent|DeferredRevenueCurrent", "BS", "first", None, None),
    ("Income Tax Payable",             "AccruedIncomeTaxes",                      "AccruedIncomeTaxesCurrent",                                 "BS", "first", None, None),
    ("Other Current Liabilities",      "OtherNonOperatingCurrentLiabilities",     "OtherLiabilitiesCurrent",                                   "BS", "first", None, None),
    ("Total Current Liabilities",      "CurrentLiabilitiesTotal",                 "LiabilitiesCurrent",                                        "BS", "first", None, None),
    ("Long-term Debt",                 "LongTermDebt",                            r"LongTerm(?:Debt|NotesAndLoans|Borrowings)(?!\w*(?<!Non)Current$)", "BS", "first", None, None),
    ("Op. Lease Liabilities, LT",      "OperatingLeaseNonCurrentDebtEquivalent",  "OperatingLeaseLiabilityNoncurrent",                         "BS", "first", None, None),
    ("Finance Lease Liabilities, LT",  None,                                      "FinanceLeaseLiabilityNoncurrent",                           "BS", "first", "finance lease", None),
    ("Deferred Revenue, LT",           "ContractLiabilities",                     "ContractWithCustomerLiabilityNoncurrent",                   "BS", "first", None, None),
    ("Deferred Tax Liability, LT",     "DeferredTaxNonCurrentLiabilities",        "DeferredIncomeTaxLiabilitiesNet",                           "BS", "first", None, None),
    ("Pension & Retirement Oblig.",    "PensionObligations",                      "PensionAndOtherPostretirementDefinedBenefitPlans",          "BS", "first", None, None),
    ("Other Non-current Liabilities",  "OtherNonOperatingNonCurrentLiabilities",  "OtherLiabilitiesNoncurrent",                                "BS", "first", None, None),
    ("Total Non-current Liabilities",  None,                                      "^us-gaap_LiabilitiesNoncurrent$",                           "BS", "first", None, None),
    ("Total Liabilities",              "Liabilities",                             "Liabilities",                                               "BS", "last",  None, None),
    # ── Equity ──────────────────────────────────────────────────────────
    ("Preferred Stock",                "PreferredStock",                          "PreferredStockValue",                                       "BS", "first", None, None),
    ("Common Stock & APIC",            "CommonEquity",                            "CommonStockValue",                                          "BS", "first", _COMMON_STOCK_HINT, None),
    ("Additional Paid-in Capital",     "AdditionalPaidInCapital",                 "AdditionalPaidInCapitalCommonStock",                        "BS", "first", None, None),
    ("Treasury Stock",                 "TreasuryShares",                          "TreasuryStockValue",                                        "BS", "first", None, None),
    ("Retained Earnings",              "RetainedEarnings",                        "RetainedEarningsAccumulatedDeficit",                        "BS", "first", None, None),
    ("AOCI",                           "AccumulatedOtherComprehensiveIncome",     "AccumulatedOtherComprehensiveIncomeLossNetOfTax",            "BS", "first", None, None),
    ("Total Equity — Parent",          "AllEquityBalance",                        "StockholdersEquity",                                        "BS", "first", None, None),
    ("Noncontrolling Interests",       "MinorityInterestBalance",                 "MinorityInterest",                                          "BS", "first", None, None),
    ("Total Equity incl. NCI",         "AllEquityBalanceIncludingMinorityInterest","StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest", "BS", "first", None, None),
    ("Total Liabilities & Equity",     "LiabilitiesAndEquity",                    "LiabilitiesAndStockholdersEquity",                          "BS", "first", None, None),
    # 期末在外流通股數（時點值）。與 IS 的 Basic/Diluted Shares 不同——那兩個是
    # 算 EPS 用的**加權平均**，在有買回或增發的季度會與期末股數差很多。
    ("Shares Outstanding",             "CommonSharesOutstanding",                 "CommonStockSharesOutstanding|EntityCommonStockSharesOutstanding", "BS", "last", None, None),
]

CF_TEMPLATE: list[_T] = [
    # ── Operating ────────────────────────────────────────────────────────
    ("Net Income",                 "NetIncome",                          "NetIncomeLoss|ProfitLoss",                              "CF", "first", None, None),
    ("D&A",                        "DepreciationExpense",                _DA_FALLBACK,                                            "CF", "first", None, _DA_LABEL_FALLBACK),
    ("SBC",                        "StockBasedCompensationExpense",      "ShareBasedCompensation",                                "CF", "first", None, None),
    ("Amortization of Intangibles","AmortizationOfIntangibles",          "AmortizationOfIntangibleAssets",                        "CF", "first", None, None),
    ("Change in Receivables",      "ChangeInReceivables",                "IncreaseDecreaseInAccountsReceivable",                  "CF", "first", "receivable", None),
    ("Change in Inventories",      None,                                 "IncreaseDecreaseIn(?:RetailRelated)?Inventories",         "CF", "first", "inventor", None),
    ("Change in Accounts Payable",     None,  "IncreaseDecreaseInAccountsPayable",                          "CF", "first", None, None),
    ("Change in Prepaid & Other Assets", None, "IncreaseDecreaseInPrepaidDeferredExpenseAndOtherAssets",     "CF", "first", None, None),
    ("Change in Other Operating Assets", None, "IncreaseDecreaseInOtherOperatingAssets",                     "CF", "first", None, None),
    ("Change in Deferred Revenue", "ChangeInDeferredRevenue",            "IncreaseDecreaseInDeferredRevenue",                     "CF", "first", None, None),
    ("Other Working Capital",      "ChangeInOtherWorkingCapital",        "IncreaseDecreaseInOtherOperatingLiabilities",           "CF", "first", None, None),
    ("Other Non-cash Items",       "OtherNonCashItemsCF",                "OtherNoncashIncomeExpense",                             "CF", "first", None, None),
    ("Operating Cash Flow",        "NetCashFromOperatingActivities",     "NetCashProvidedByUsedInOperatingActivities",            "CF", "last",  "^net cash|^cash|^total", None),
    # ── Investing ────────────────────────────────────────────────────────
    # label_fallback 是為了 NVDA 這種用自訂延伸 tag 的公司：它的 10-K 從 FY2013 到
    # FY2023 共 11 年 tag 成 `nvda_PurchasesOfPropertyAndEquipmentAndIntangibleAssets`，
    # concept 兩層都比不到，年報那格空掉之後連 Q4 都合成不出來（見 H4 spec）。
    # **刻意用 `^purchases` 錨在開頭**——第三層後面沒有任何東西再擋它了，寫寬一點
    # 就會吃到「Proceeds from sales of property」與「Depreciation of property」。
    ("Capex",                      "CapitalExpenses",                    "PaymentsToAcquirePropertyPlantAndEquipment",            "CF", "first", _CAPEX_HINT, r"^purchases (?:of|related to).*propert"),
    ("Acquisitions",               "AcquisitionsNet",                    "PaymentsToAcquireBusinessesNetOfCashAcquired",          "CF", "first", None, None),
    ("Investment Purchases",       "InvestmentPurchases",                "PaymentsToAcquireInvestments",                          "CF", "first", None, None),
    ("Investment Proceeds",        "InvestmentProceeds",                 "ProceedsFromSaleOfInvestments",                         "CF", "first", None, None),
    ("Investing Cash Flow",        "NetCashFromInvestingActivities",     "NetCashProvidedByUsedInInvestingActivities",            "CF", "last",  "^net cash|^cash|^total", None),
    # ── Financing ────────────────────────────────────────────────────────
    # 借款／還款這兩列**刻意不在這裡比對**，值一律由下面的 `_sum_matching_rows()`
    # 加總（見 `_DEBT_PROCEEDS_PATTERNS`）。理由：公司常常同時有長期借款、商業本票
    # 等好幾條借款線，要的是總額；而且模板一旦比對到其中一條，`consumed` 會讓加總
    # 跳過它，加總再覆蓋回 row_vals 就只剩「其他幾條的和」——比不比對還錯。
    ("Debt Proceeds",              None,                                 "",                                                      "CF", "first", None, None),
    ("Debt Repayments",            None,                                 "",                                                      "CF", "first", None, None),
    ("Share Repurchases",          None,                                 "PaymentsForRepurchaseOfCommonStock",                    "CF", "first", None, None),
    ("Dividends Paid",             None,                                  "PaymentsOfDividends|PaymentsOfDividendsCommonStock|PaymentsOfOrdinaryDividends", "CF", "first", "dividend", None),
    ("Financing Cash Flow",        "NetCashFromFinancingActivities",     "NetCashProvidedByUsedInFinancingActivities",            "CF", "last",  "^net cash|^cash|^total", None),
    # ── Other ────────────────────────────────────────────────────────────
    ("FX Effect on Cash",          "ForeignExchangeEffectOnCash",        "EffectOfExchangeRateOnCashAndCashEquivalents",          "CF", "first", None, None),
    ("Net Change in Cash",         "NetChangeInCash",                    "CashAndCashEquivalentsPeriodIncreaseDecrease",          "CF", "first", None, None),
    ("Ending Cash",                "CashAndCashEquivalents",             "CashAndCashEquivalentsAtCarryingValue",                 "CF", "last",  None, None),
    ("Cash Taxes Paid",            None,                                 "IncomeTaxesPaid",                                       "CF", "first", None, None),
    ("Cash Interest Paid",         None,                                 "InterestPaid",                                          "CF", "first", "interest|^total", None),
    # ── Derived (computed, not from XBRL) ────────────────────────────────
    ("Free Cash Flow",             None,                                 "",                                                      "DERIVED", "first", None, None),
]

# ── Index maps for post-processing derived / fallback rows ────────────────

_IS_IDX: dict[str, int] = {row[0]: i for i, row in enumerate(IS_TEMPLATE)}
_NONOP_TOTAL_IDX   = _IS_IDX["Total Non-op Income/(Loss)"]
_OP_INCOME_IDX     = _IS_IDX["Operating Income"]
_PRETAX_IDX        = _IS_IDX["Pre-tax Income"]
_NET_INCOME_IDX    = _IS_IDX["Net Income"]
_DA_CF_IDX         = _IS_IDX["D&A (CF memo)"]
_REVENUE_IDX       = _IS_IDX["Revenue"]
_COGS_IDX          = _IS_IDX["Cost of Revenue"]
_GROSS_PROFIT_IDX  = _IS_IDX["Gross Profit"]

_BS_IDX: dict[str, int] = {row[0]: i for i, row in enumerate(BS_TEMPLATE)}
_TCA_IDX = _BS_IDX["Total Current Assets"]
_NCA_IDX = _BS_IDX["Total Non-current Assets"]
_TA_IDX  = _BS_IDX["Total Assets"]
_TCL_IDX = _BS_IDX["Total Current Liabilities"]
_NCL_IDX = _BS_IDX["Total Non-current Liabilities"]
_TL_IDX  = _BS_IDX["Total Liabilities"]

_CF_IDX: dict[str, int] = {row[0]: i for i, row in enumerate(CF_TEMPLATE)}
_CF_NET_INCOME_IDX      = _CF_IDX["Net Income"]
_CF_DA_IDX              = _CF_IDX["D&A"]
_CF_OP_CASH_IDX         = _CF_IDX["Operating Cash Flow"]
_CF_CAPEX_IDX           = _CF_IDX["Capex"]
_CF_FCF_IDX             = _CF_IDX["Free Cash Flow"]

# 現金流量表裡混著**時點值**：`Ending Cash` 是期末現金「餘額」，不是本期發生額。
# `_build_cf_table()` 對 YTD 欄做「本季 YTD − 上季 YTD」還原單季，那對流量項
# 是對的，對餘額就完全錯了——減出來是「現金變動額」不是「餘額」。
#
# 實測 AAPL（2026-08-22 做 companyfacts 逐格比對時發現）：
#     2026-03-28   錯誤     255,000,000   正確 45,572,000,000
#     2026-06-27   錯誤  -6,028,000,000   正確 39,544,000,000
# 52 家裡 50 家中招。
#
# 期初現金沒有獨立列（`Net Change in Cash` 是流量、是對的），所以目前只有這一列。
# 日後在 CF_TEMPLATE 加時點值列時**要記得加進來**。
_CF_POINT_IN_TIME_IDX = frozenset({_CF_IDX["Ending Cash"]})

_CF_INV_PURCHASES_IDX  = _CF_IDX["Investment Purchases"]
_CF_INV_PROCEEDS_IDX   = _CF_IDX["Investment Proceeds"]
_CF_DEBT_PROCEEDS_IDX   = _CF_IDX["Debt Proceeds"]
_CF_DEBT_REPAYMENTS_IDX = _CF_IDX["Debt Repayments"]

_INV_PROCEEDS_PATTERNS: list[str] = [
    r"ProceedsFromSaleOfInvestments",
    r"ProceedsFromSaleOfAvailableForSaleSecurities",
    r"ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities",
    r"ProceedsFromSaleAndMaturityOfMarketableSecurities",
    r"ProceedsFromSaleOfShortTermInvestments",
    r"ProceedsFromSaleMaturityAndCollectionOfShorttermInvestments",
]
# 借款／還款一律「把所有借款線加總」，不是挑一條——實測 INTC、PG、XOM、COST
# 都同時有長期借款與商業本票兩條以上。
#
# ⚠ `ProceedsFromRepaymentsOf...` 是**淨額**（借款減還款），不能算進任何一邊：
# 它可正可負，塞進 Debt Repayments 會讓「還款」出現負數。CAT、GE、WMT、NKE、
# XOM 都有這種列，所以兩個正則都要把它排掉（借款用 `(?!Repayments)`，還款用
# `(?<!ProceedsFrom)`）。排掉之後它們會落到 overflow，資料不會消失。
_DEBT_PROCEEDS_PATTERNS: list[str] = [
    r"ProceedsFromIssuanceOfDebt$",
    r"ProceedsFrom(?!Repayments)\w*(?:LongTermDebt|ShortTermDebt|ShortTermBorrowings"
    r"|CommercialPaper|ConvertibleDebt|NotesPayable|LinesOfCredit|MediumTermNotes"
    r"|DebtNetOfIssuanceCosts|DebtMaturingIn)",
]
_DEBT_REPAYMENTS_PATTERNS: list[str] = [
    r"(?<!ProceedsFrom)RepaymentsOf\w*(?:Debt|Borrowings|CommercialPaper"
    r"|NotesPayable|LinesOfCredit|MediumTermNotes)",
]


# ── Helpers ────────────────────────────────────────────────────────────────

def _col_to_quarter_label(col_name: str, fy_end_month: int = 12) -> str:
    """Convert edgartools period column name to FY label.

    fy_end_month: company's fiscal year end month (1-12). Default 12 = calendar year.
    For non-December FY companies, quarterly periods ending after fy_end_month belong
    to the next fiscal year (e.g. AAPL Sep FY: Dec 2023 Q1 → FY2024Q1).
    Annual (FY) labels are never adjusted.

    Examples (default fy_end_month=12):
        "2023-03-31 (Q1)"  -> "FY2023Q1"
        "2024-12-31 (FY)"  -> "FY2024"
    Examples (fy_end_month=9, AAPL):
        "2023-12-30 (Q1)"  -> "FY2024Q1"
        "2024-09-28 (FY)"  -> "FY2024"

    財季編號**由期末日反推，不採信欄名裡的 `(Qn)`**：edgartools 對 52/53 週
    財年制的公司會標錯。實測 NVDA 2016-05-01（FY2017 Q1）被標成 `(Q2)`、
    INTC 2023-04-01（2023 Q1，13 週制溢出到 4 月）也被標成 `(Q2)`。兩份不同
    期間的 10-Q 因此算出同一個 label，`_build_*_table` 的 dedup
    （`if label in periods: continue`）就把舊的那一季靜默丟掉，連帶讓
    `_synthesize_q4()` 缺 Q1/Q2/Q3 而合成不出 Q4。`(Qn)` 現在只用來分辨
    「這是季度欄還是年度欄」，編號一律自己算。
    """
    m = re.match(r"(\d{4})-(\d{2})-\d{2}\s+\((\w+)\)", col_name.strip())
    if m:
        year, period = int(m.group(1)), m.group(3)
        if period.upper() == "FY":
            return f"FY{year}"
        # 延後 import：fiscal_input -> excel_formatter -> fetcher_gaap 會循環匯入
        from fiscal_input import fiscal_quarter_of, fy_start_month
        label = fiscal_quarter_of(_col_to_period_end(col_name),
                                  fy_start_month(fy_end_month))
        return label or f"FY{year}{period}"
    return col_name


def _col_to_period_end(col_name: str) -> str:
    """edgartools 欄名 → 期末日。`"2026-03-29 (Q1)"` → `"2026-03-29"`。

    抓不到回空字串——下游會退回從財季標籤反推的年月。
    """
    m = re.match(r"(\d{4}-\d{2}-\d{2})\s+\(\w+\)", (col_name or "").strip())
    return m.group(1) if m else ""


def _detect_fy_end_month(filings_k: list) -> int:
    """Detect company's fiscal year end month from 10-K filings.

    Looks for a column labeled '(FY)' in the IS statement of the first 3 10-K filings.
    Returns the month number (1-12), defaulting to 12 (December) if not detected.
    """
    for filing in filings_k[:3]:
        try:
            tenq = _filing_obj(filing)
            fin = _financials_of(tenq)
            if fin is None:
                continue
            is_stmt = fin.income_statement()
            if is_stmt is None:
                continue
            df = is_stmt.to_dataframe()
            for col in df.columns:
                if col in META_COLS:
                    continue
                mm = re.search(r"\d{4}-(\d{2})-\d{2}\s+\(FY\)", col)
                if mm:
                    return int(mm.group(1))
        except Exception as exc:
            # 這一步失敗會退回 12 月（見函式結尾），而財年結束月是所有期間
            # 標籤的基準——AAPL 退成 12 月的話每個標籤都差一季。所以一定要
            # 記進帳本讓使用者看到警告，不能安靜地用錯的基準跑完。
            _note_gap(_filing_ref(filing), exc)
            continue
    return 12


def _is_q_col(col_name: str) -> bool:
    """True if column is a quarterly period (Qx or FY), False for YTD."""
    m = re.search(r"\((\w+)\)", col_name)
    if not m:
        return False
    period = m.group(1).upper()
    return bool(re.match(r"Q\d+$", period)) or period == "FY"


def _current_q_col(df) -> str | None:
    """Return the first quarterly (non-YTD) period column from a filing's DataFrame."""
    for col in df.columns:
        if col in META_COLS:
            continue
        if _is_q_col(col):
            return col
    return None


def _ytd_col(df) -> str | None:
    """Return the first YTD period column (labeled '(YTD)'), or None."""
    for col in df.columns:
        if col in META_COLS:
            continue
        m = re.search(r"\((\w+)\)", col)
        if m and m.group(1).upper() == "YTD":
            return col
    return None


def _prev_quarter_label(label: str) -> str | None:
    """Return the label of the previous quarter, or None for Q1 / annual.

    Examples:
        "FY2025Q2" → "FY2025Q1"
        "FY2025Q1" → None
        "FY2025"   → None
    """
    m = re.match(r"(FY\d{4})Q(\d+)$", label)
    if not m:
        return None
    fy, q = m.group(1), int(m.group(2))
    if q <= 1:
        return None
    return f"{fy}Q{q - 1}"


def _consolidated_mask(df):
    """Boolean mask: non-abstract, non-breakdown, no dimension."""
    mask = ~df.get("abstract", False).astype(bool)
    mask &= ~df.get("is_breakdown", False).astype(bool)
    dim_col = df.get("dimension_member_label")
    if dim_col is not None:
        mask &= dim_col.isna() | (dim_col.astype(str) == "nan")
    return mask


# TODO: Verify fallback accuracy against ≥10 tickers.
#       Tested: AAPL, TSLA, BA, XOM (Session 9).
#       Still needed: MSFT, AMZN, META, GOOGL, NVDA, JPM, GS, JNJ.
def _match_is_row(df, std_concept: str | None, fallback_suffix: str,
                   label_fallback: str | None = None,
                   match: str = "first",
                   label_hint: str | None = None) -> int | None:
    """Find the row index in df matching a template entry.

    Priority order (each level only tried when previous level produces no usable result):
        1. standard_concept == std_concept (consolidated rows only)
        2. concept column contains fallback_suffix (case-insensitive, consolidated only)
        3. label column contains label_fallback (case-insensitive, consolidated only)

    label_hint: when candidates are found at a priority level, filter to those whose
        label contains this string.  If the filter leaves no candidates, the entire
        priority level is skipped and the next level is tried — candidates that fail
        label_hint are never returned as a fallback.
    match: "first" → earliest matching row; "last" → latest matching row.

    Returns None if no match found at any priority level.
    """
    mask = _consolidated_mask(df)
    df_c = df[mask]

    def _pick(rows) -> int | None:
        """Apply label_hint filter and return the selected index, or None if filtered out."""
        if rows.empty:
            return None
        if label_hint:
            hinted = rows[rows["label"].astype(str).str.contains(label_hint, case=False, na=False)]
            if hinted.empty:
                return None          # hint not satisfied → skip this priority level
            rows = hinted
        return rows.index[-1] if match == "last" else rows.index[0]

    # Priority 1: standard_concept exact match
    if std_concept:
        rows = df_c[df_c["standard_concept"].astype(str) == std_concept]
        result = _pick(rows)
        if result is not None:
            return result

    # Priority 2: concept contains fallback_suffix (supports regex OR via "|")
    if fallback_suffix:
        rows = df_c[df_c["concept"].astype(str).str.contains(fallback_suffix, case=False, na=False)]
        result = _pick(rows)
        if result is not None:
            return result

    # Priority 3: label contains label_fallback
    if label_fallback:
        rows = df_c[df_c["label"].astype(str).str.contains(label_fallback, case=False, na=False)]
        result = _pick(rows)
        if result is not None:
            return result

    return None


def _apply_row_override(df: pd.DataFrame, col: str, override_entry: dict) -> Any:
    """Look up a value from df using a pre-diagnosed override entry.

    concept_override: re-query df with the override's std_concept.
    structural_absence: return None immediately (confirmed missing in XBRL).
    """
    fix_type = override_entry.get("fix_type")
    if fix_type == "structural_absence":
        return None
    if fix_type == "concept_override":
        sc = override_entry.get("std_concept", "")
        if sc and col in df.columns:
            idx = _match_is_row(df, sc, sc)
            if idx is not None:
                return _to_python_val(df.loc[idx, col])
    return None


def _to_python_val(val) -> Any:
    """Convert pandas NA / float NaN / None to None; leave other values as-is."""
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    return val


def _row_key(df_row) -> str:
    """Unique key for a data row: concept + optional dimension_member_label."""
    concept = str(df_row.get("concept", "") or "")
    dim = str(df_row.get("dimension_member_label", "") or "")
    if dim and dim != "nan":
        return f"{concept}|{dim}"
    return concept


def _seg_sheet_suffix(concept: str, standard_concept: str | None) -> str:
    """Generate a ≤22-char alphanumeric suffix for a segment sheet name."""
    raw = standard_concept if standard_concept and standard_concept != "nan" else concept
    raw = re.sub(r"^[a-z_]+[_:]", "", raw)
    raw = re.sub(r"[^A-Za-z0-9]", "", raw)
    return raw[:22]


# ── Overflow helpers ──────────────────────────────────────────────────────────

# Labels containing any of these substrings are routed to the Non-GAAP overflow
# sheet instead of the GAAP overflow section.  Matching is case-insensitive.
_NONGAAP_KEYWORDS: frozenset[str] = frozenset({
    "non-gaap", "non gaap", "adjusted", "excluding", "excl.", "ex-",
})


def _is_nongaap_label(label: str) -> bool:
    """Return True if label looks like a Non-GAAP / adjusted metric."""
    low = label.lower()
    return any(kw in low for kw in _NONGAAP_KEYWORDS)


def _collect_overflow(
    df: pd.DataFrame,
    consumed: set[int],
    data_col: str,
    quarter_label: str,
    gaap_out: dict,
    ng_out: dict,
) -> None:
    """Collect unmatched XBRL rows from df into gaap_out or ng_out dicts.

    Rows whose index is in `consumed` are skipped (already captured by template).
    Abstract, breakdown, and dimension rows are excluded via _consolidated_mask.
    When a value is None the key is still recorded (so the concept appears in
    output even when it has no data for this specific quarter); periods dict
    only stores non-None values — the caller decides whether to drop all-None rows.
    """
    mask = _consolidated_mask(df)
    df_c = df[mask]
    remaining = df_c[~df_c.index.isin(consumed)]
    for _, row in remaining.iterrows():
        key = str(row.get("concept", "") or "")
        if not key or key == "nan":
            continue
        raw = str(row.get("label", "") or "")
        display = unicodedata.normalize("NFKC", raw)
        out = ng_out if _is_nongaap_label(display) else gaap_out
        if key not in out:
            out[key] = {"label": display, "periods": {}}
        val = _to_python_val(row.get(data_col))
        if val is not None:
            out[key]["periods"][quarter_label] = val


def _sum_matching_rows(
    df: pd.DataFrame,
    col: str,
    patterns: list[str],
    consumed: set[int],
) -> tuple[Any, list[int]]:
    """Sum values from consolidated rows whose concept matches any pattern in patterns.

    Skips rows already in consumed. Returns (total_or_None, list_of_matched_indices).
    """
    mask = _consolidated_mask(df)
    df_c = df[mask]
    total: float | None = None
    indices: list[int] = []
    seen_concepts: set[str] = set()
    for pattern in patterns:
        matches = df_c[df_c["concept"].astype(str).str.contains(pattern, case=False, na=False, regex=True)]
        for idx, row in matches.iterrows():
            if idx in consumed:
                continue
            concept = str(row.get("concept", "") or "")
            if concept in seen_concepts:
                continue
            seen_concepts.add(concept)
            val = _to_python_val(row.get(col))
            if val is not None:
                total = (total or 0.0) + val
                indices.append(idx)
    return total, indices


def _build_template_table(filings, template: list[_T], sheet_name: str,
                           stmt_method: str, max_filings: int,
                           fy_end_month: int = 12) -> StatementTable:
    """Generic fixed-template builder used by IS, BS, and CF."""
    periods: dict[str, tuple[str, dict[int, Any]]] = {}
    row_labels: dict[int, str] = {}   # first available original XBRL label per row

    for filing in filings:
        if len(periods) >= max_filings:
            break
        try:
            tenq = _filing_obj(filing)
            fin = _financials_of(tenq)
            if fin is None:
                continue
            stmt = getattr(fin, stmt_method)()
            if stmt is None:
                continue
            df = stmt.to_dataframe()
            _note_ok()
        except Exception as exc:
            _note_gap(_filing_ref(filing), exc)
            print(f"[fetcher_gaap] {sheet_name} warning: {type(exc).__name__}", file=sys.stderr)
            continue

        q_col = _current_q_col(df)
        if q_col is None:
            continue

        label = _col_to_quarter_label(q_col, fy_end_month)
        if label in periods:
            continue

        row_vals: dict[int, Any] = {}
        for i, (_, std_concept, fallback, source, match, label_hint, lbl_fb) in enumerate(template):
            if source == "DERIVED":
                row_vals[i] = None   # filled in post-processing
                continue
            idx = _match_is_row(df, std_concept, fallback, label_fallback=lbl_fb,
                                 match=match, label_hint=label_hint)
            val = _to_python_val(df.loc[idx, q_col]) if idx is not None else None
            row_vals[i] = val
            if idx is not None and i not in row_labels:
                raw = str(df.loc[idx, "label"] or "")
                row_labels[i] = unicodedata.normalize("NFKC", raw)

        periods[label] = (str(filing.filing_date), row_vals)

    if not periods:
        return StatementTable(
            sheet_name=sheet_name,
            quarter_labels=[],
            filing_dates=[],
            concepts=[row[0] for row in template],
            values=[[] for _ in template],
            labels=["" for _ in template],
        )

    sorted_labels = sorted(periods.keys())
    filing_dates  = [periods[lbl][0] for lbl in sorted_labels]

    values: list[list[Any]] = []
    for i in range(len(template)):
        values.append([periods[lbl][1].get(i) for lbl in sorted_labels])

    labels_list = [row_labels.get(i, "") for i in range(len(template))]

    return StatementTable(
        sheet_name=sheet_name,
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=[row[0] for row in template],
        values=values,
        labels=labels_list,
    )


# ── IS: template-based fetch ────────────────────────────────────────────────

def _build_is_table(
    filings, max_filings: int, is_overrides: dict | None = None,
    fy_end_month: int = 12,
) -> tuple[StatementTable, StatementTable]:
    """Build IS StatementTables from 10-Q filings using the fixed IS template.

    Returns (gaap_tbl, ng_tbl):
      gaap_tbl — template rows + GAAP overflow rows (unmatched XBRL items)
      ng_tbl   — Non-GAAP overflow rows (labels containing "adjusted", "non-gaap", etc.)
                 Empty table when no Non-GAAP rows are found.
    """
    is_overrides = is_overrides or {}
    period_end_map: dict[str, str] = {}
    periods: dict[str, tuple[str, dict[int, Any]]] = {}
    row_labels: dict[int, str] = {}
    # Overflow dicts accumulate across all filings; key = XBRL concept name
    gaap_overflow: dict[str, dict] = {}
    ng_overflow:   dict[str, dict] = {}

    for filing in filings:
        if len(periods) >= max_filings:
            break
        _fd = getattr(filing, "filing_date", None)
        if isinstance(_fd, _date) and _fd < _XBRL_CUTOFF:
            break   # filings are newest-first; everything older is also pre-XBRL
        try:
            tenq = _filing_obj(filing)
            fin = _financials_of(tenq)
            if fin is None:
                continue
            stmt = fin.income_statement()
            if stmt is None:
                continue
            df = stmt.to_dataframe()
            _note_ok()
            _tick_progress("IS")
        except Exception as exc:
            _note_gap(_filing_ref(filing), exc)
            _tick_progress("IS")
            print(f"[fetcher_gaap] IS warning: {type(exc).__name__}", file=sys.stderr)
            continue

        q_col = _current_q_col(df)
        if q_col is None:
            continue

        label = _col_to_quarter_label(q_col, fy_end_month)
        if label in periods:
            continue
        # 真正的期末結算日（52/53 週制不是月底）。從財季標籤反推只能得到月份，
        # Data_Std 需要精確日期，所以趁還看得到欄名先存下來。
        period_end_map[label] = _col_to_period_end(q_col)

        # Fetch CF statement for D&A / SBC rows
        cf_df: pd.DataFrame | None = None
        cf_q_col: str | None = None
        try:
            cf_stmt = tenq.financials.cashflow_statement()
            if cf_stmt is not None:
                cf_df = cf_stmt.to_dataframe()
                cf_q_col = _current_q_col(cf_df)
        except Exception:
            pass

        # Tracks which IS df indices are consumed by template matching this filing.
        # CF-sourced rows (source == "CF") consume cf_df indices — not tracked here.
        consumed: set[int] = set()

        row_vals: dict[int, Any] = {}
        for i, (row_name, std_concept, fallback, source, match, label_hint, lbl_fb) in enumerate(IS_TEMPLATE):
            # Apply override if one exists for this row (concept_override or structural_absence)
            if row_name in is_overrides:
                ov = is_overrides[row_name]
                if ov.get("fix_type") == "structural_absence":
                    row_vals[i] = None
                    continue
                # concept_override: try in IS df; CF-sourced rows are not in KEY_ROWS so no IS overrides expected
                val = _apply_row_override(df, q_col, ov)
                row_vals[i] = val
                continue

            if source == "CF":
                # CF-sourced rows look in cf_df; those indices are NOT tracked in IS consumed set
                # (CF overflow is handled separately by _build_cf_table)
                if cf_df is not None and cf_q_col is not None:
                    idx = _match_is_row(cf_df, std_concept, fallback, label_fallback=lbl_fb,
                                        match=match, label_hint=label_hint)
                    val = _to_python_val(cf_df.loc[idx, cf_q_col]) if idx is not None else None
                    if idx is not None and i not in row_labels:
                        raw = str(cf_df.loc[idx, "label"] or "")
                        row_labels[i] = unicodedata.normalize("NFKC", raw)
                else:
                    val = None
            else:
                idx = _match_is_row(df, std_concept, fallback, label_fallback=lbl_fb,
                                    match=match, label_hint=label_hint)
                if idx is not None:
                    consumed.add(idx)   # mark as consumed so overflow skips this row
                val = _to_python_val(df.loc[idx, q_col]) if idx is not None else None
                if idx is not None and i not in row_labels:
                    raw = str(df.loc[idx, "label"] or "")
                    row_labels[i] = unicodedata.normalize("NFKC", raw)
            row_vals[i] = val

        # ── Post-processing: fallbacks not expressible in the 6-tuple ──

        # 1. Total Non-op: DERIVED = Pre-tax − Operating Income
        #    Guard: skip if discontinued operations present (would distort the difference)
        if row_vals.get(_NONOP_TOTAL_IDX) is None:
            op_val     = row_vals.get(_OP_INCOME_IDX)
            pretax_val = row_vals.get(_PRETAX_IDX)
            has_discontinued = _match_is_row(df, None, "DiscontinuedOperations") is not None
            if op_val is not None and pretax_val is not None and not has_discontinued:
                row_vals[_NONOP_TOTAL_IDX] = pretax_val - op_val

        # 2. Net Income fallback chain
        if row_vals.get(_NET_INCOME_IDX) is None:
            # 2a. Parent-only net income (more precise than ProfitLoss)
            idx = _match_is_row(df, "NetIncomeLossAttributableToParent",
                                 "NetIncomeLossAttributableToParent")
            if idx is not None:
                consumed.add(idx)
                row_vals[_NET_INCOME_IDX] = _to_python_val(df.loc[idx, q_col])
                if _NET_INCOME_IDX not in row_labels:
                    row_labels[_NET_INCOME_IDX] = unicodedata.normalize(
                        "NFKC", str(df.loc[idx, "label"] or ""))

        if row_vals.get(_NET_INCOME_IDX) is None:
            # 2b. ProfitLoss last resort (includes NCI — use only when parent-only unavailable)
            idx = _match_is_row(df, "ProfitLoss", "ProfitLoss")
            if idx is not None:
                consumed.add(idx)
                row_vals[_NET_INCOME_IDX] = _to_python_val(df.loc[idx, q_col])
                if _NET_INCOME_IDX not in row_labels:
                    row_labels[_NET_INCOME_IDX] = unicodedata.normalize(
                        "NFKC", str(df.loc[idx, "label"] or ""))

        # 3. D&A label fallback: for companies where standard_concept = nan (TSLA)
        #    This searches cf_df, not IS df — no IS consumed tracking needed
        if row_vals.get(_DA_CF_IDX) is None and cf_df is not None and cf_q_col is not None:
            idx = _match_is_row(cf_df, None, "", label_fallback="depreciation")
            if idx is not None:
                row_vals[_DA_CF_IDX] = _to_python_val(cf_df.loc[idx, cf_q_col])
                if _DA_CF_IDX not in row_labels:
                    row_labels[_DA_CF_IDX] = unicodedata.normalize(
                        "NFKC", str(cf_df.loc[idx, "label"] or ""))

        # 4. Gross Profit: DERIVED = Revenue − COGS (companies without explicit GP in XBRL)
        if row_vals.get(_GROSS_PROFIT_IDX) is None:
            rev  = row_vals.get(_REVENUE_IDX)
            cogs = row_vals.get(_COGS_IDX)
            if rev is not None and cogs is not None:
                row_vals[_GROSS_PROFIT_IDX] = rev - cogs

        # Collect unmatched IS df rows into overflow buckets
        _collect_overflow(df, consumed, q_col, label, gaap_overflow, ng_overflow)

        periods[label] = (str(filing.filing_date), row_vals)

    if not periods:
        empty = StatementTable(
            sheet_name="Data_IS",
            quarter_labels=[],
            filing_dates=[],
            concepts=[row[0] for row in IS_TEMPLATE],
            values=[[] for _ in IS_TEMPLATE],
            labels=["" for _ in IS_TEMPLATE],
        )
        empty_ng = StatementTable(
            sheet_name="Data_IS_NG",
            quarter_labels=[], filing_dates=[],
            concepts=[], values=[], labels=[],
        )
        return empty, empty_ng

    sorted_labels = sorted(periods.keys())
    filing_dates  = [periods[lbl][0] for lbl in sorted_labels]

    # ── Build GAAP table: template rows + GAAP overflow ───────────────────
    concepts_g: list[str]       = [row[0] for row in IS_TEMPLATE]
    labels_g:   list[str]       = [row_labels.get(i, "") for i in range(len(IS_TEMPLATE))]
    values_g:   list[list[Any]] = [
        [periods[lbl][1].get(i) for lbl in sorted_labels]
        for i in range(len(IS_TEMPLATE))
    ]
    for key in sorted(gaap_overflow):
        entry = gaap_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue   # skip entirely-empty overflow rows
        concepts_g.append(entry["label"] or key)
        labels_g.append(key)
        values_g.append(row)

    period_ends = [period_end_map.get(lbl, "") for lbl in sorted_labels]

    gaap_tbl = StatementTable(
        sheet_name="Data_IS",
        quarter_labels=sorted_labels,
        period_ends=period_ends,
        filing_dates=filing_dates,
        concepts=concepts_g,
        labels=labels_g,
        values=values_g,
    )

    # ── Build NG table: Non-GAAP overflow rows only (no template rows) ────
    concepts_n: list[str]       = []
    labels_n:   list[str]       = []
    values_n:   list[list[Any]] = []
    for key in sorted(ng_overflow):
        entry = ng_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue
        concepts_n.append(entry["label"] or key)
        labels_n.append(key)
        values_n.append(row)

    ng_tbl = StatementTable(
        sheet_name="Data_IS_NG",
        quarter_labels=sorted_labels,
        period_ends=period_ends,
        filing_dates=filing_dates,
        concepts=concepts_n,
        labels=labels_n,
        values=values_n,
    )

    return gaap_tbl, ng_tbl


# ── BS: template-based fetch ────────────────────────────────────────────────

def _build_bs_table(filings, max_filings: int, bs_overrides: dict | None = None,
                    fy_end_month: int = 12) -> tuple[StatementTable, StatementTable]:
    """Build Data_BS StatementTable using the fixed BS template.

    Balance sheet columns in edgartools are instant (bare date, e.g. "2024-03-31")
    rather than period ("2024-03-31 (Q1)"), so _current_q_col cannot find them.
    We derive the quarter label from the IS statement (same filing) for merge alignment.

    Returns (gaap_tbl, ng_tbl). gaap_tbl = template rows + GAAP overflow;
    ng_tbl = Non-GAAP overflow rows only (sheet "Data_BS_NG").
    """
    bs_overrides = bs_overrides or {}
    periods: dict[str, tuple[str, dict[int, Any]]] = {}
    row_labels: dict[int, str] = {}
    gaap_overflow: dict[str, dict] = {}  # {concept_key: {"label": str, "periods": {q: val}}}
    ng_overflow: dict[str, dict] = {}

    for filing in filings:
        if len(periods) >= max_filings:
            break
        _fd = getattr(filing, "filing_date", None)
        if isinstance(_fd, _date) and _fd < _XBRL_CUTOFF:
            break   # filings are newest-first; everything older is also pre-XBRL
        try:
            tenq = _filing_obj(filing)
            fin = _financials_of(tenq)
            if fin is None:
                continue

            # Get quarter label from IS (has "(Q1)"/"(FY)" format)
            is_stmt = fin.income_statement()
            is_df = is_stmt.to_dataframe() if is_stmt is not None else None
            is_q_col = _current_q_col(is_df) if is_df is not None else None

            bs_stmt = fin.balance_sheet()
            if bs_stmt is None:
                continue
            df = bs_stmt.to_dataframe()
            _note_ok()
            _tick_progress("BS")
        except Exception as exc:
            _note_gap(_filing_ref(filing), exc)
            _tick_progress("BS")
            print(f"[fetcher_gaap] BS warning: {type(exc).__name__}", file=sys.stderr)
            continue

        # BS columns are bare dates; pick first non-meta column
        bs_col = next((c for c in df.columns if c not in META_COLS), None)
        if bs_col is None:
            continue

        label = _col_to_quarter_label(is_q_col, fy_end_month) if is_q_col else _col_to_quarter_label(bs_col, fy_end_month)
        if label in periods:
            continue

        consumed: set[int] = set()
        row_vals: dict[int, Any] = {}
        for i, (row_name, std_concept, fallback, source, match, label_hint, lbl_fb) in enumerate(BS_TEMPLATE):
            if source == "DERIVED":
                row_vals[i] = None
                continue
            if row_name in bs_overrides:
                ov = bs_overrides[row_name]
                if ov.get("fix_type") == "structural_absence":
                    row_vals[i] = None
                    continue
                row_vals[i] = _apply_row_override(df, bs_col, ov)
                continue
            idx = _match_is_row(df, std_concept, fallback, label_fallback=lbl_fb,
                                match=match, label_hint=label_hint)
            val = _to_python_val(df.loc[idx, bs_col]) if idx is not None else None
            row_vals[i] = val
            if idx is not None:
                consumed.add(idx)
                if i not in row_labels:
                    raw = str(df.loc[idx, "label"] or "")
                    row_labels[i] = unicodedata.normalize("NFKC", raw)

        # ── Post-processing: non-current subtotals ──────────────────────────
        # 多數公司不直接標記 AssetsNoncurrent / LiabilitiesNoncurrent，用
        # Total - Current 相減補上，跟 IS 的 Total Non-op 做法一致。
        if row_vals.get(_NCA_IDX) is None:
            ta, tca = row_vals.get(_TA_IDX), row_vals.get(_TCA_IDX)
            if ta is not None and tca is not None:
                row_vals[_NCA_IDX] = ta - tca

        if row_vals.get(_NCL_IDX) is None:
            tl, tcl = row_vals.get(_TL_IDX), row_vals.get(_TCL_IDX)
            if tl is not None and tcl is not None:
                row_vals[_NCL_IDX] = tl - tcl

        _collect_overflow(df, consumed, bs_col, label, gaap_overflow, ng_overflow)
        periods[label] = (str(filing.filing_date), row_vals)

    empty_ng = StatementTable(
        sheet_name="Data_BS_NG", quarter_labels=[], filing_dates=[],
        concepts=[], values=[], labels=[],
    )
    if not periods:
        return StatementTable(
            sheet_name="Data_BS",
            quarter_labels=[],
            filing_dates=[],
            concepts=[row[0] for row in BS_TEMPLATE],
            values=[[] for _ in BS_TEMPLATE],
            labels=["" for _ in BS_TEMPLATE],
        ), empty_ng

    sorted_labels = sorted(periods.keys())
    filing_dates = [periods[lbl][0] for lbl in sorted_labels]

    # ── Build GAAP table: template rows + GAAP overflow ─────────────────────
    concepts_g: list[str]       = [row[0] for row in BS_TEMPLATE]
    labels_g:   list[str]       = [row_labels.get(i, "") for i in range(len(BS_TEMPLATE))]
    values_g:   list[list[Any]] = [
        [periods[lbl][1].get(i) for lbl in sorted_labels]
        for i in range(len(BS_TEMPLATE))
    ]
    for key in sorted(gaap_overflow):
        entry = gaap_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue
        concepts_g.append(entry["label"] or key)
        labels_g.append(key)
        values_g.append(row)

    gaap_tbl = StatementTable(
        sheet_name="Data_BS",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=concepts_g,
        labels=labels_g,
        values=values_g,
    )

    # ── Build NG table: Non-GAAP overflow rows only (no template rows) ───────
    concepts_n: list[str]       = []
    labels_n:   list[str]       = []
    values_n:   list[list[Any]] = []
    for key in sorted(ng_overflow):
        entry = ng_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue
        concepts_n.append(entry["label"] or key)
        labels_n.append(key)
        values_n.append(row)

    ng_tbl = StatementTable(
        sheet_name="Data_BS_NG",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=concepts_n,
        labels=labels_n,
        values=values_n,
    )
    return gaap_tbl, ng_tbl


# ── CF: template-based fetch ────────────────────────────────────────────────

def _build_cf_table(filings, max_filings: int, cf_overrides: dict | None = None,
                    fy_end_month: int = 12) -> tuple[StatementTable, StatementTable]:
    """Build Data_CF StatementTable using the fixed CF template.

    Q1 and FY filings have standalone period columns (Q1/FY) and are used directly.
    Q2 and Q3 filings have YTD (cumulative) CF columns; standalone quarter values are
    derived by subtracting the prior period's YTD: Q2 = Q2_YTD − Q1, Q3 = Q3_YTD − Q2_YTD.
    Overflow rows use the same subtraction logic: raw values are collected per filing and
    converted to standalone after the filing loop.

    Returns (gaap_tbl, ng_tbl). gaap_tbl = template rows + GAAP overflow;
    ng_tbl = Non-GAAP overflow rows only (sheet "Data_CF_NG").
    """
    cf_overrides = cf_overrides or {}
    # collected: label → (filing_date, {row_i: raw_value}, is_ytd)
    collected: dict[str, tuple[str, dict[int, Any], bool]] = {}
    # ytd_raw: stores raw values (standalone for Q1/FY, cumulative for YTD)
    # used as the subtraction base for the next YTD period
    ytd_raw: dict[str, dict[int, Any]] = {}
    row_labels: dict[int, str] = {}
    gaap_overflow: dict[str, dict] = {}  # {concept_key: {"label": str, "periods": {q: val}}}
    ng_overflow: dict[str, dict] = {}
    # Raw overflow per filing; YTD subtraction applied after loop (mirrors template logic)
    # {q_label: {concept_key: (display_label, is_nongaap, raw_val)}}
    overflow_per_filing: dict[str, dict] = {}

    for filing in filings:
        if len(collected) >= max_filings:
            break
        _fd = getattr(filing, "filing_date", None)
        if isinstance(_fd, _date) and _fd < _XBRL_CUTOFF:
            break   # filings are newest-first; everything older is also pre-XBRL
        try:
            tenq = _filing_obj(filing)
            fin = _financials_of(tenq)
            if fin is None:
                continue
            is_stmt = fin.income_statement()
            is_df = is_stmt.to_dataframe() if is_stmt is not None else None
            is_q_col = _current_q_col(is_df) if is_df is not None else None

            cf_stmt = fin.cashflow_statement()
            if cf_stmt is None:
                continue
            df = cf_stmt.to_dataframe()
            _note_ok()
            _tick_progress("CF")
        except Exception as exc:
            _note_gap(_filing_ref(filing), exc)
            _tick_progress("CF")
            print(f"[fetcher_gaap] CF warning: {type(exc).__name__}", file=sys.stderr)
            continue

        q_col = _current_q_col(df)
        if q_col is not None:
            label = _col_to_quarter_label(q_col, fy_end_month)
            if label in collected:
                continue
            is_ytd = False
            data_col = q_col
        else:
            ytd_col = _ytd_col(df)
            if ytd_col is None or is_q_col is None:
                continue
            label = _col_to_quarter_label(is_q_col, fy_end_month)
            if label in collected:
                continue
            is_ytd = True
            data_col = ytd_col

        consumed: set[int] = set()
        row_vals: dict[int, Any] = {}
        for i, (row_name, std_concept, fallback, source, match, label_hint, lbl_fb) in enumerate(CF_TEMPLATE):
            if source == "DERIVED":
                row_vals[i] = None
                continue
            if row_name in cf_overrides:
                ov = cf_overrides[row_name]
                if ov.get("fix_type") == "structural_absence":
                    row_vals[i] = None
                    continue
                row_vals[i] = _apply_row_override(df, data_col, ov)
                continue
            idx = _match_is_row(df, std_concept, fallback, label_fallback=lbl_fb,
                                match=match, label_hint=label_hint)
            val = _to_python_val(df.loc[idx, data_col]) if idx is not None else None
            row_vals[i] = val
            if idx is not None:
                consumed.add(idx)
                if i not in row_labels:
                    raw = str(df.loc[idx, "label"] or "")
                    row_labels[i] = unicodedata.normalize("NFKC", raw)

        # Post-processing (BEFORE overflow): Investment Proceeds — sum all relevant rows
        inv_proc_val, inv_proc_indices = _sum_matching_rows(df, data_col, _INV_PROCEEDS_PATTERNS, consumed)
        if inv_proc_val is not None:
            row_vals[_CF_INV_PROCEEDS_IDX] = inv_proc_val
            consumed.update(inv_proc_indices)

        # Post-processing (BEFORE overflow): Debt Proceeds — sum LT + ST + credit lines
        debt_proc_val, debt_proc_indices = _sum_matching_rows(df, data_col, _DEBT_PROCEEDS_PATTERNS, consumed)
        if debt_proc_val is not None:
            row_vals[_CF_DEBT_PROCEEDS_IDX] = debt_proc_val
            consumed.update(debt_proc_indices)

        # Post-processing (BEFORE overflow): Debt Repayments — sum LT + ST + credit lines
        debt_rep_val, debt_rep_indices = _sum_matching_rows(df, data_col, _DEBT_REPAYMENTS_PATTERNS, consumed)
        if debt_rep_val is not None:
            row_vals[_CF_DEBT_REPAYMENTS_IDX] = debt_rep_val
            consumed.update(debt_rep_indices)

        # Collect raw overflow for all filings (incl. YTD); YTD subtraction applied after loop
        df_c = df[_consolidated_mask(df)]
        remaining = df_c[~df_c.index.isin(consumed)]
        filing_ov: dict[str, tuple[str, bool, Any]] = {}
        for _, ov_row in remaining.iterrows():
            ov_key = str(ov_row.get("concept", "") or "")
            if not ov_key or ov_key == "nan":
                continue
            ov_raw = str(ov_row.get("label", "") or "")
            ov_display = unicodedata.normalize("NFKC", ov_raw)
            ov_val = _to_python_val(ov_row.get(data_col))
            filing_ov[ov_key] = (ov_display, _is_nongaap_label(ov_display), ov_val)
        overflow_per_filing[label] = filing_ov

        collected[label] = (str(filing.filing_date), row_vals, is_ytd)
        ytd_raw[label] = row_vals  # Q1 standalone doubles as Q1 YTD base

    empty_ng = StatementTable(
        sheet_name="Data_CF_NG", quarter_labels=[], filing_dates=[],
        concepts=[], values=[], labels=[],
    )
    if not collected:
        return StatementTable(
            sheet_name="Data_CF",
            quarter_labels=[],
            filing_dates=[],
            concepts=[row[0] for row in CF_TEMPLATE],
            values=[[] for _ in CF_TEMPLATE],
            labels=["" for _ in CF_TEMPLATE],
        ), empty_ng

    sorted_labels = sorted(collected.keys())

    # Convert YTD periods to standalone quarters via subtraction
    standalone: dict[str, dict[int, Any]] = {}
    for label in sorted_labels:
        _, row_vals, is_ytd = collected[label]
        if not is_ytd:
            standalone[label] = row_vals
        else:
            prev_label = _prev_quarter_label(label)
            if prev_label and prev_label in ytd_raw:
                prev = ytd_raw[prev_label]
                standalone[label] = {
                    # 時點值（期末餘額）直接採用，不參與相減——見
                    # `_CF_POINT_IN_TIME_IDX` 的說明
                    i: (row_vals.get(i) if i in _CF_POINT_IN_TIME_IDX else
                        (row_vals.get(i) - prev.get(i)
                         if row_vals.get(i) is not None and prev.get(i) is not None
                         else None))
                    for i in range(len(CF_TEMPLATE))
                }
            else:
                standalone[label] = row_vals  # no prior YTD — keep cumulative as best-effort

    # ── Convert YTD overflow to standalone (mirrors template subtraction above) ──
    # Union of all concepts seen in any filing's overflow
    all_overflow_keys: dict[str, tuple[str, bool]] = {}  # concept_key → (label, is_nongaap)
    for filing_ov in overflow_per_filing.values():
        for ov_key, (ov_lbl, ov_is_ng, _) in filing_ov.items():
            if ov_key not in all_overflow_keys:
                all_overflow_keys[ov_key] = (ov_lbl, ov_is_ng)

    for ov_key, (display_label, is_ng) in all_overflow_keys.items():
        out = ng_overflow if is_ng else gaap_overflow
        if ov_key not in out:
            out[ov_key] = {"label": display_label, "periods": {}}
        for q_lbl in sorted_labels:
            _, _, lbl_is_ytd = collected[q_lbl]
            filing_ov = overflow_per_filing.get(q_lbl, {})
            raw_val = filing_ov[ov_key][2] if ov_key in filing_ov else None
            if not lbl_is_ytd:
                if raw_val is not None:
                    out[ov_key]["periods"][q_lbl] = raw_val
            else:
                prev_lbl = _prev_quarter_label(q_lbl)
                prev_ov = overflow_per_filing.get(prev_lbl, {})
                prev_val = prev_ov[ov_key][2] if ov_key in prev_ov else None
                if raw_val is not None and prev_val is not None:
                    out[ov_key]["periods"][q_lbl] = raw_val - prev_val

    filing_dates = [collected[lbl][0] for lbl in sorted_labels]
    values: list[list[Any]] = [
        [standalone[lbl].get(i) for lbl in sorted_labels]
        for i in range(len(CF_TEMPLATE))
    ]

    tbl = StatementTable(
        sheet_name="Data_CF",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=[row[0] for row in CF_TEMPLATE],
        values=values,
        labels=[row_labels.get(i, "") for i in range(len(CF_TEMPLATE))],
    )

    # FCF = OCF − |Capex| (abs normalises sign: some companies report Capex negative)
    for j in range(len(sorted_labels)):
        op_cf = tbl.values[_CF_OP_CASH_IDX][j]
        capex = tbl.values[_CF_CAPEX_IDX][j]
        if op_cf is not None and capex is not None:
            tbl.values[_CF_FCF_IDX][j] = op_cf - abs(capex)

    # ── Build GAAP table: template rows (from tbl) + GAAP overflow ──────────
    concepts_g = tbl.concepts[:]
    labels_g   = tbl.labels[:]
    values_g   = [row[:] for row in tbl.values]
    for key in sorted(gaap_overflow):
        entry = gaap_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue
        concepts_g.append(entry["label"] or key)
        labels_g.append(key)
        values_g.append(row)

    gaap_tbl = StatementTable(
        sheet_name="Data_CF",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=concepts_g,
        labels=labels_g,
        values=values_g,
    )

    # ── Build NG table: Non-GAAP overflow rows only (no template rows) ───────
    concepts_n: list[str]       = []
    labels_n:   list[str]       = []
    values_n:   list[list[Any]] = []
    for key in sorted(ng_overflow):
        entry = ng_overflow[key]
        row = [entry["periods"].get(q) for q in sorted_labels]
        if all(v is None for v in row):
            continue
        concepts_n.append(entry["label"] or key)
        labels_n.append(key)
        values_n.append(row)

    ng_tbl = StatementTable(
        sheet_name="Data_CF_NG",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=concepts_n,
        labels=labels_n,
        values=values_n,
    )
    return gaap_tbl, ng_tbl


# ── Three-statement merge ───────────────────────────────────────────────────

_STD_QUARTER_RE = re.compile(r"FY(\d{4})Q([1-4])$")


# ── 期間換算 ────────────────────────────────────────────────────────────────

def _fiscal_period_end(label: str, fy_end_month: int) -> tuple[int, int] | None:
    """財季標籤 → (西元年, 月)，指該財季**結束**的年月。無法解析回 None。

    財年 Y 結束於西元 Y 年的 fy_end_month 月（SEC 慣例）。第 q 財季結束於
    財年結束前 (4-q) 季，也就是 fy_end_month - 3*(4-q) 個月。
    """
    m = _STD_QUARTER_RE.match((label or "").strip())
    if m is None:
        return None
    year, q = int(m.group(1)), int(m.group(2))
    month = fy_end_month - 3 * (4 - q)
    while month <= 0:
        month += 12
        year -= 1
    return year, month


def _calendar_quarter(label: str, fy_end_month: int, period_end: str = "") -> str:
    """財季標籤 → 日曆季標籤（`2026Q1`）。年度標籤或無法解析回空字串。

    有真實期末日就直接用它換算（最準）；沒有才靠結算月反推。

    **判準一律委派給 `fiscal_input.calendar_quarter_of(basis="end")`，這裡不自己算。**
    2026-08-22 之前這裡是第三套獨立實作，而且忘了內縮——INTC 結束在 2023-04-01
    的那一季（實際涵蓋 1~3 月）會被算成 `2023Q2`。平常這個值會被
    `fiscal_input._apply_to_sheet()` 的公式蓋掉看不到，但第 5 列不是完整 ISO
    日期的殘留格（合成 Q4 的年報期末日有時只有 `2010-01`）會保留它，那就是錯值。

    退路（期末日抓不到，靠財季標籤反推）保留：`_fiscal_period_end()` 算出來的
    月份本來就是「該季最後一個月」，不需要再內縮，直接取日曆季即可。
    """
    if period_end and re.match(r"\d{4}-\d{2}-\d{2}", period_end):
        # 延後 import：fiscal_input -> excel_formatter -> fetcher_gaap 會循環匯入
        from fiscal_input import calendar_quarter_of
        return calendar_quarter_of(period_end, basis="end")
    parsed = _fiscal_period_end(label, fy_end_month)
    if parsed is None:
        return ""
    year, month = parsed
    return f"{year}Q{(month - 1) // 3 + 1}"


def _fiscal_quarter(label: str) -> str:
    """`FY2026Q1` → `FY2026FQ1`。

    財季用 `FQ` 標記、財年用 `FY`，與第 4 列的日曆季（`2026Q1`）在視覺上就分得開。
    這兩列最容易被搞混——非 12 月結算的公司同一欄可能是 FY2026FQ1 但日曆 2025Q4，
    看錯就是整整一季的誤差，所以刻意讓兩種寫法長得不一樣。
    """
    m = _STD_QUARTER_RE.match((label or "").strip())
    return f"FY{m.group(1)}FQ{m.group(2)}" if m else ""


def _period_end(label: str, fy_end_month: int) -> str:
    """財季標籤 → 期末年月（`2026-03`）。無法解析回空字串。"""
    parsed = _fiscal_period_end(label, fy_end_month)
    if parsed is None:
        return ""
    year, month = parsed
    return f"{year}-{month:02d}"


# overflow 區的分隔標題。放公司特有、不在固定模板裡的 XBRL 科目——
# 實測每家中位數 IS 4 / BS 2 / CF 10 個，合計約 16 列。
OVERFLOW_SECTION = "Other (as reported)"

# 三表之間空幾列。捲動時要一眼看出換表了，光靠標題底色不夠。
SECTION_GAP = 5


def _next_fiscal_label(label: str) -> str:
    """`FY2025Q3` → `FY2025Q4` → `FY2026Q1`。不是財季標籤回空字串。"""
    m = re.match(r"^FY(\d{4})Q([1-4])$", label)
    if m is None:
        return ""
    year, quarter = int(m.group(1)), int(m.group(2))
    return f"FY{year + 1}Q1" if quarter == 4 else f"FY{year}Q{quarter + 1}"


def _with_gap_columns(all_qs: list[str], period_ends: list[str]) -> list[str]:
    """抓不到的季度不要整欄消失，補一個空白欄位進去（G6，2026-08-25）。

    現況欄位清單是「成功抓到什麼就放什麼」，某一季掛掉整欄消失，畫面上
    FY2025Q1 直接跳到 FY2025Q3，使用者與 AI 都看不出中間漏了一季。補出來的
    欄位沒有任何值，第 5 列期末日退回由財季標籤反推的年月（`2025-06`），
    在 Excel 上就是一整欄空白——「有漏」這件事因此看得見。

    缺口判定沿用 `data_quality.missing_quarters()`（`round(天數差/91) - 1`，
    單一缺口上限 4 季），**不要在這裡另外寫一份**：那條公式與上限是 52 家、
    1,482 對相鄰期間實測定下來的，固定門檻會把 COSTCO 的 16 週第四季誤判成
    缺一季。

    年度表（`FY2025` 這種標籤）不處理——季度那套天數算法不適用。
    """
    from data_quality import missing_quarters

    if not all(re.match(r"^FY\d{4}Q[1-4]$", q) for q in all_qs):
        return all_qs

    end_to_label = {end: q for q, end in zip(all_qs, period_ends) if end}
    filled = set(all_qs)
    for gap in missing_quarters(period_ends):
        label = end_to_label.get(gap.after, "")
        for _ in range(gap.count):
            label = _next_fiscal_label(label)
            if not label:
                break
            filled.add(label)
    return sorted(filled)


def _merge_financials(is_tbl: StatementTable,
                       bs_tbl: StatementTable,
                       cf_tbl: StatementTable,
                       sheet_name: str = "Data_Financials(Q)",
                       fy_end_month: int = 12) -> StatementTable:
    """Merge IS + BS + CF into a single StatementTable.

    Quarter union is taken across all three statements; missing values are None.
    Section header rows ("Income Statement", "Balance Sheet", "Cash Flow")
    are inserted as separator rows with all-None values and empty labels.
    """
    all_qs = sorted(
        set(is_tbl.quarter_labels)
        | set(bs_tbl.quarter_labels)
        | set(cf_tbl.quarter_labels)
    )

    # Build date map (IS takes priority over BS over CF)
    date_map: dict[str, str] = {}
    for tbl in [cf_tbl, bs_tbl, is_tbl]:
        for lbl, dt in zip(tbl.quarter_labels, tbl.filing_dates):
            date_map[lbl] = dt
    filing_dates = [date_map.get(q, "") for q in all_qs]

    # 期末日同樣三表取聯集（IS 優先）。缺的留空，Data_Std 會退回反推年月。
    end_map: dict[str, str] = {}
    for tbl in [cf_tbl, bs_tbl, is_tbl]:
        for lbl, end in zip(tbl.quarter_labels, tbl.period_ends or []):
            if end:
                end_map[lbl] = end
    period_ends = [end_map.get(q, "") for q in all_qs]

    # 抓不到的季度留一整欄空白（G6）。要放在 filing_dates / period_ends 算完
    # **之後**——缺口判定吃的就是那份期末日序列；補進來的欄位在兩個 map 裡都
    # 查不到，自然落成空字串，第 5 列再由標籤反推年月。
    with_gaps = _with_gap_columns(all_qs, period_ends)
    if with_gaps != all_qs:
        all_qs = with_gaps
        filing_dates = [date_map.get(q, "") for q in all_qs]
        period_ends = [end_map.get(q, "") for q in all_qs]

    concepts:    list[str]        = []
    labels_col:  list[str]        = []
    values:      list[list[Any]]  = []

    def _add_header(title: str) -> None:
        concepts.append(title)
        labels_col.append("")
        values.append([None] * len(all_qs))

    def _add_blank() -> None:
        concepts.append("")
        labels_col.append("")
        values.append([None] * len(all_qs))

    # overflow（公司特有、不在模板裡的 XBRL 科目）全部集中到 sheet 最底部。
    #
    # 原本 overflow 接在每個 section 的模板列之後，IS 多幾行 overflow，BS 整段
    # 就往下推——實測 11 個輸出檔裡 `Cash` 落在第 28~56 列之間，跨檔案公式因此
    # 完全寫不出來。移到底部後模板列號跨公司固定。
    overflow: list[tuple[str, str, list[Any]]] = []

    def _add_rows(tbl: StatementTable, template_names: set[str]) -> None:
        q_idx = {q: j for j, q in enumerate(tbl.quarter_labels)}
        for i, concept in enumerate(tbl.concepts):
            label = tbl.labels[i] if tbl.labels else ""
            row = [_to_python_val(tbl.values[i][q_idx[q]])
                   if q in q_idx else None
                   for q in all_qs]
            if concept in template_names:
                concepts.append(concept)
                labels_col.append(label)
                values.append(row)
            else:
                overflow.append((concept, label, row))

    # 期間標籤三列放最上面。財季用 FY/FQ、日曆季用純數字，兩者視覺上分得開——
    # 非 12 月結算的公司同一欄可能是 FY2026FQ1 但日曆 2025Q4，看錯就是整整一季。
    def _add_label_row(name: str, vals: list[Any]) -> None:
        concepts.append(name)
        labels_col.append("")
        values.append(vals)

    _add_label_row("Fiscal Quarter", [_fiscal_quarter(q) for q in all_qs])
    _add_label_row("Calendar Quarter",
                   [_calendar_quarter(q, fy_end_month, period_ends[i] if i < len(period_ends) else "")
                    for i, q in enumerate(all_qs)])
    _add_label_row("Period End",
                   [(period_ends[i] if i < len(period_ends) and period_ends[i]
                     else _period_end(q, fy_end_month)) for i, q in enumerate(all_qs)])
    _add_blank()

    _add_header("Income Statement")
    _add_rows(is_tbl, {r[0] for r in IS_TEMPLATE})
    for _ in range(SECTION_GAP):
        _add_blank()
    _add_header("Balance Sheet")
    _add_rows(bs_tbl, {r[0] for r in BS_TEMPLATE})
    for _ in range(SECTION_GAP):
        _add_blank()
    _add_header("Cash Flow")
    _add_rows(cf_tbl, {r[0] for r in CF_TEMPLATE})

    if overflow:
        for _ in range(SECTION_GAP):
            _add_blank()
        _add_header(OVERFLOW_SECTION)
        for concept, label, row in overflow:
            concepts.append(concept)
            labels_col.append(label)
            values.append(row)

    return StatementTable(
        sheet_name=sheet_name,
        quarter_labels=all_qs,
        filing_dates=filing_dates,
        period_ends=period_ends,
        concepts=concepts,
        values=values,
        labels=labels_col,
    )


# ── Q4 synthesis (D0-1, 2026-08-20 CTH 決定要修) ─────────────────────────────
#
# SEC 沒有 Q4 的 10-Q——公司只交 Q1/Q2/Q3，Q4 數字要嘛在 10-K 年報裡，要嘛靠推算。
# 有年報（10-K）可用時，用年度值反推單季 Q4：
#   IS/CF（流量項）：Q4 = 年報 FY 值 − Q1 − Q2 − Q3
#   BS（存量項，資產負債表本來就是年底時點數字）：Q4 = 年報 FY 值直接取用，不相減
#
# 只處理模板列（`n_template_rows` 以內）——overflow 列（公司特有科目）在季報
# 與年報兩邊出現的順序不保證對齊，沒有可靠的列對應，Q4 一律留 None。
# IS 模板裡 source=="CF" 的列 → 對應 CF 模板的哪一列。
# 兩邊名字刻意不同（IS 那邊帶「(CF memo)」提醒讀者這不是損益表原生科目），
# 所以要有這張對照表，不能靠名字相等。
_IS_ROWS_SOURCED_FROM_CF = {
    "SBC": "SBC",
    "D&A (CF memo)": "D&A",
}


def _backfill_cf_sourced_rows(is_tbl: StatementTable,
                              cf_tbl: StatementTable) -> StatementTable:
    """用已建好的 CF 表回填 IS 裡 source=="CF" 的那幾列。

    為什麼要這樣做（2026-08-22，TODO G3）：10-Q 的現金流量表是 **YTD 累計**，
    Q2/Q3 的 filing 沒有單季欄。`_build_is_table()` 原本用
    `_current_q_col(cf_df)` 直接找單季欄，找不到就整格留空——實測 NVDA 的
    `SBC` 與 `D&A (CF memo)` 缺 51/68 期，缺的全是 Q2/Q3（Q1 有值，因為 Q1 的
    YTD 就等於單季），連帶 `_synthesize_q4()` 也算不出那兩列的 Q4。

    `_build_cf_table()` 已經做過 YTD 拆算（本季 YTD − 上季 YTD，見該函式的
    `is_ytd` 分支），所以 CF 區同名兩列是好的。這裡直接共用它的結果，
    **不要在 IS 再寫一份 YTD 拆算**——那就是第二份會漂移的實作。

    依 `quarter_labels` 對照，不靠欄位位置（兩張表的期數不保證一樣）。
    CF 那格是 None 就不動 IS 原本的值，不要用 None 蓋掉真資料。
    """
    cf_idx_by_label = {q: i for i, q in enumerate(cf_tbl.quarter_labels)}
    new_values = [list(row) for row in is_tbl.values]

    for is_name, cf_name in _IS_ROWS_SOURCED_FROM_CF.items():
        if is_name not in is_tbl.concepts or cf_name not in cf_tbl.concepts:
            continue
        is_row = is_tbl.concepts.index(is_name)
        cf_row = cf_tbl.concepts.index(cf_name)
        for j, label in enumerate(is_tbl.quarter_labels):
            k = cf_idx_by_label.get(label)
            if k is None:
                continue
            val = cf_tbl.values[cf_row][k]
            if val is not None:
                new_values[is_row][j] = val

    return replace(is_tbl, values=new_values)


def _synthesize_q4(q_tbl: StatementTable, ann_tbl: StatementTable,
                    n_template_rows: int, is_balance: bool,
                    point_in_time_idx: frozenset[int] = frozenset()) -> StatementTable:
    """Insert a synthesized Q4 column into q_tbl for each fiscal year covered by ann_tbl.

    Skips a fiscal year when:
      - Q4 already exists in q_tbl (never overwrite real data)
      - is_balance=False and Q1/Q2/Q3 aren't all present in q_tbl
      - the derived column would be entirely None

    Returns q_tbl unchanged (same object) if ann_tbl has no annual periods.

    `point_in_time_idx`：**流量表裡混著的餘額列**。整張表 `is_balance=False`
    時這些列仍然要直接取年報值，不可以做「年報 − Q1 − Q2 − Q3」。
    現金流量表的 `Ending Cash` 就是這種——2026-08-22 實測 AAPL 的 Q4 欄位算出
    −58,796,000,000（正確是 +45,317,000,000）。見 `_CF_POINT_IN_TIME_IDX`。
    """
    if not ann_tbl.quarter_labels:
        return q_tbl

    q_idx = {q: i for i, q in enumerate(q_tbl.quarter_labels)}
    new_labels = list(q_tbl.quarter_labels)
    new_dates = list(q_tbl.filing_dates)
    new_ends = list(q_tbl.period_ends) if q_tbl.period_ends else [""] * len(new_labels)
    new_values = [list(row) for row in q_tbl.values]
    added = False

    for fy_i, fy_label in enumerate(ann_tbl.quarter_labels):
        if "Q" in fy_label:
            continue   # not a pure FY label (shouldn't normally happen for an annual table)
        q4_label = f"{fy_label}Q4"
        if q4_label in q_idx:
            continue

        if is_balance:
            col_vals = [ann_tbl.values[r][fy_i] if r < len(ann_tbl.values) else None
                        for r in range(n_template_rows)]
        else:
            q1, q2, q3 = f"{fy_label}Q1", f"{fy_label}Q2", f"{fy_label}Q3"
            if not (q1 in q_idx and q2 in q_idx and q3 in q_idx):
                continue
            i1, i2, i3 = q_idx[q1], q_idx[q2], q_idx[q3]
            col_vals = []
            for r in range(n_template_rows):
                fy_val = ann_tbl.values[r][fy_i] if r < len(ann_tbl.values) else None
                if r in point_in_time_idx:
                    col_vals.append(fy_val)      # 餘額：年報上的期末值就是 Q4 的值
                    continue
                v1 = q_tbl.values[r][i1] if r < len(q_tbl.values) else None
                v2 = q_tbl.values[r][i2] if r < len(q_tbl.values) else None
                v3 = q_tbl.values[r][i3] if r < len(q_tbl.values) else None
                if fy_val is None or v1 is None or v2 is None or v3 is None:
                    col_vals.append(None)
                else:
                    col_vals.append(fy_val - v1 - v2 - v3)

        if all(v is None for v in col_vals):
            continue

        new_labels.append(q4_label)
        new_dates.append(ann_tbl.filing_dates[fy_i] if fy_i < len(ann_tbl.filing_dates) else "")
        ann_end = (ann_tbl.period_ends[fy_i]
                   if ann_tbl.period_ends and fy_i < len(ann_tbl.period_ends) else "")
        new_ends.append(ann_end)
        for r in range(len(new_values)):
            new_values[r].append(col_vals[r] if r < len(col_vals) else None)
        added = True

    if not added:
        return q_tbl

    order = sorted(range(len(new_labels)), key=lambda i: new_labels[i])
    return StatementTable(
        sheet_name=q_tbl.sheet_name,
        quarter_labels=[new_labels[i] for i in order],
        filing_dates=[new_dates[i] for i in order],
        period_ends=[new_ends[i] for i in order],
        concepts=q_tbl.concepts,
        values=[[row[i] for i in order] for row in new_values],
        ticker=q_tbl.ticker,
        labels=q_tbl.labels,
    )


# ── BS/CF: dynamic row-union fetch (kept for reference / fallback) ──────────

def _build_dynamic_table(filings, stmt_method: str, sheet_name: str,
                          max_filings: int,
                          fy_end_month: int = 12) -> StatementTable | None:
    """Build BS or CF StatementTable using a dynamic row union across all filings."""
    concept_labels: dict[str, str] = {}
    periods: dict[str, tuple[str, dict[str, Any]]] = {}

    for filing in filings:
        if len(periods) >= max_filings:
            break
        try:
            fin = _financials_of(_filing_obj(filing))
            if fin is None:
                continue
            stmt = getattr(fin, stmt_method)()
            if stmt is None:
                continue
            df = stmt.to_dataframe()
            _note_ok()
        except Exception as exc:
            _note_gap(_filing_ref(filing), exc)
            print(f"[fetcher_gaap] {sheet_name} warning: {type(exc).__name__}", file=sys.stderr)
            continue

        q_col = _current_q_col(df)
        if q_col is None:
            continue

        label = _col_to_quarter_label(q_col, fy_end_month)
        if label in periods:
            continue

        mask = _consolidated_mask(df)
        df_c = df[mask].reset_index(drop=True)

        period_vals: dict[str, Any] = {}
        for _, row in df_c.iterrows():
            key = _row_key(row)
            if key not in concept_labels:
                raw = str(row.get("label", "") or key)
                concept_labels[key] = unicodedata.normalize("NFKC", raw)
            period_vals[key] = _to_python_val(row.get(q_col))

        periods[label] = (str(filing.filing_date), period_vals)

    if not periods or not concept_labels:
        return None

    sorted_labels = sorted(periods.keys())
    filing_dates = [periods[lbl][0] for lbl in sorted_labels]
    concepts_ordered = list(concept_labels.keys())

    values: list[list[Any]] = []
    for key in concepts_ordered:
        values.append([periods[lbl][1].get(key) for lbl in sorted_labels])

    return StatementTable(
        sheet_name=sheet_name,
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=[concept_labels[k] for k in concepts_ordered],
        values=values,
    )


# ── Segment breakdown sheets ────────────────────────────────────────────────

def _dimension_axis(row) -> str:
    """從 XBRL 列取維度軸名稱，例如 `us-gaap:StatementBusinessSegmentsAxis`。

    edgartools 有 `dimension_axis` 欄；沒有時退回從 `dimension_label`
    （格式為 `"軸: 成員"`）取冒號前那段。都取不到回空字串。
    """
    axis = str(row.get("dimension_axis", "") or "").strip()
    if axis and axis != "nan":
        return axis
    label = str(row.get("dimension_label", "") or "")
    return label.split(":", 2)[0].strip() + ":" + label.split(":", 2)[1].strip()         if label.count(":") >= 2 else (label.split(":")[0].strip() if ":" in label else "")


def _build_segment_tables(filings, max_filings: int, fy_end_month: int = 12) -> list[StatementTable]:
    """Build one StatementTable per IS concept that has segment/dimension rows."""
    seg_data: dict[str, dict] = {}
    periods_seen: set[str] = set()

    for filing in filings:
        if len(periods_seen) >= max_filings:
            break
        _fd = getattr(filing, "filing_date", None)
        if isinstance(_fd, _date) and _fd < _XBRL_CUTOFF:
            break   # filings are newest-first; everything older is also pre-XBRL
        try:
            fin = _financials_of(_filing_obj(filing))
            if fin is None:
                continue
            stmt = fin.income_statement()
            if stmt is None:
                continue
            df = stmt.to_dataframe()
            _note_ok()
        except Exception as exc:
            _note_gap(_filing_ref(filing), exc)
            print(f"[fetcher_gaap] Seg warning: {type(exc).__name__}", file=sys.stderr)
            continue

        q_col = _current_q_col(df)
        if q_col is None:
            continue

        period_label = _col_to_quarter_label(q_col, fy_end_month)
        filing_date = str(filing.filing_date)

        dim_col = df.get("dimension_member_label")
        if dim_col is None:
            continue
        mask_dim = ~(dim_col.isna() | (dim_col.astype(str) == "nan"))
        mask_not_abstract = ~df.get("abstract", False).astype(bool)
        df_dim = df[mask_dim & mask_not_abstract]

        for _, row in df_dim.iterrows():
            concept_xbrl = str(row.get("concept", "") or "")
            if not concept_xbrl:
                continue
            std = str(row.get("standard_concept", "") or "nan")
            member = str(row.get("dimension_member_label", "") or "")

            if concept_xbrl not in seg_data:
                seg_data[concept_xbrl] = {"std": std, "members": {}, "axes": {}, "periods": {}}

            seg_data[concept_xbrl]["members"].setdefault(member, member)
            # 記下這個 member 屬於哪個維度軸。沒有軸就分不出「業務別營收」與
            # 「權益項目別」——MSFT 實測會混進 `Retained earnings`、`Service Life`
            # 這種根本不是 segment 的東西。
            seg_data[concept_xbrl]["axes"].setdefault(member, _dimension_axis(row))

            if period_label not in seg_data[concept_xbrl]["periods"]:
                seg_data[concept_xbrl]["periods"][period_label] = (filing_date, {})

            seg_data[concept_xbrl]["periods"][period_label][1][member] = _to_python_val(row.get(q_col))

        periods_seen.add(period_label)

    tables: list[StatementTable] = []
    used_sheet_names: set[str] = set()

    for concept_xbrl, data in seg_data.items():
        if not data["periods"]:
            continue

        suffix = _seg_sheet_suffix(concept_xbrl, data["std"] if data["std"] != "nan" else None)
        sheet_name = f"Data_Seg_{suffix}"
        base = sheet_name
        n = 2
        while sheet_name in used_sheet_names:
            sheet_name = f"{base[:28]}_{n}"
            n += 1
        used_sheet_names.add(sheet_name)

        sorted_periods = sorted(data["periods"].keys())
        members_ordered = list(data["members"].keys())

        tables.append(StatementTable(
            sheet_name=sheet_name,
            quarter_labels=sorted_periods,
            filing_dates=[data["periods"][lbl][0] for lbl in sorted_periods],
            concepts=members_ordered,
            values=[[data["periods"][lbl][1].get(m) for lbl in sorted_periods]
                    for m in members_ordered],
            # B 欄借放維度軸，讓 segments.py 組長格式時能標出每一列屬於哪個軸
            labels=[data["axes"].get(m, "") for m in members_ordered],
        ))

    return tables


# ── Meta sheet ─────────────────────────────────────────────────────────────

# ── 期末流通股數（走封面頁 dei fact，不在三表裡）──────────────────────────
#
# 實測 ARLO / AAPL / NVDA / MSFT / COHR 五家都**沒有**在資產負債表 tag
# `us-gaap:CommonStockSharesOutstanding`——股數只寫在 `CommonStockValue` 的
# label 文字裡（"shares issued and outstanding: 108,745,373 at March 29, 2026"）。
#
# 真正拿得到的是封面頁的 `dei:EntityCommonStockSharesOutstanding`，走
# `Company.get_facts()`，ARLO 有 32 筆、AAPL 70 筆，2009 年起逐季都有。
#
# ⚠ 這個 fact 的日期是封面頁的「最近可行日期」，比財季結束**晚幾週**
#   （ARLO FY2025Q1 財季結束 2025-03-30，股數是 2025-05-02 的 103,400,957）。
#   它是公開資料裡最接近的時點股數，但不是財季結束當天的數字。用它算 BVPS
#   等於「期末權益 ÷ 幾週後的股數」，量級無虞但不是同一天。

_SHARES_CONCEPT = "dei:EntityCommonStockSharesOutstanding"
# 年度標籤（`FY2025`），用來把 10-K 的封面頁股數同時補進季表的 Q4
_ANNUAL_LABEL = re.compile(r"FY\d{4}")


def _shares_label(fiscal_year, fiscal_period) -> str | None:
    """fact 的 (fiscal_year, fiscal_period) → 本專案的期間標籤。無法對映回 None。"""
    if fiscal_year is None or not fiscal_period:
        return None
    period = str(fiscal_period).strip().upper()
    try:
        year = int(fiscal_year)
    except (TypeError, ValueError):
        return None
    if period == "FY":
        return f"FY{year}"
    if re.fullmatch(r"Q[1-4]", period):
        return f"FY{year}{period}"
    return None


def _shares_map_from_records(records) -> dict[str, float]:
    """fact 記錄 → {期間標籤: 股數}。重複申報取最後一筆（更正後的）。

    **10-K 的封面頁 fact 標的是 `fp='FY'`**，對出來的標籤是 `FY2025`；季表要的
    標籤是 `FY2025Q4`，對不上就每年缺一格。52 家實測 43 家的 `Shares Outstanding`
    「中間有洞」，成因全部是這個——AAPL/NVDA/WMT/COST/MU/ADBE 都是 Q1~Q3 全中、
    Q4 全空。年表照舊用 `FY2025`，季表另外補一份 `FY2025Q4`。

    公司若真的另外標了 `fp='Q4'`，那筆比封面頁的年度 fact 更貼近季末，**不覆蓋**。
    """
    out: dict[str, float] = {}
    annual: dict[str, float] = {}
    for rec in records or []:
        label = _shares_label(rec.get("fiscal_year"), rec.get("fiscal_period"))
        if label is None:
            continue
        value = rec.get("numeric_value")
        if value is None:
            continue
        try:
            out[label] = float(value)
        except (TypeError, ValueError):
            continue
        if _ANNUAL_LABEL.fullmatch(label):
            annual[label + "Q4"] = out[label]
    for label, value in annual.items():
        out.setdefault(label, value)
    return out


def _fetch_shares_outstanding(company) -> dict[str, float]:
    """抓封面頁流通股數的完整歷史序列。任何失敗回空 dict（該列留白即可）。"""
    try:
        facts = company.get_facts()
        df = facts.query().by_concept(_SHARES_CONCEPT).to_dataframe()
        if df is None or df.empty:
            return {}
        cols = [c for c in ("fiscal_year", "fiscal_period", "numeric_value")
                if c in df.columns]
        if len(cols) < 3:
            return {}
        return _shares_map_from_records(df[cols].to_dict("records"))
    except Exception as exc:
        print(f"[fetcher_gaap] 流通股數取得失敗: {type(exc).__name__}", file=sys.stderr)
        return {}


def _apply_shares_outstanding(tables: list[StatementTable],
                              shares_map: dict[str, float]) -> None:
    """把股數填進各表的 `Shares Outstanding` 列（就地修改）。

    沒有那一列的表直接跳過；對映不到的季度留 None，不用鄰近季度頂替。
    """
    if not shares_map:
        return
    for tbl in tables:
        if "Shares Outstanding" not in tbl.concepts:
            continue
        idx = tbl.concepts.index("Shares Outstanding")
        tbl.values[idx] = [shares_map.get(lbl) for lbl in tbl.quarter_labels]


def _build_meta_table(ticker: str, company_name: str,
                       tables: list[StatementTable],
                       fy_end_month: int = 12,
                       gap_note: str = "") -> StatementTable:
    """Build Data_Meta sheet with filing summary info."""
    n_quarters = 0
    quarter_labels: list[str] = []
    filing_dates: list[str] = []
    for tbl in tables:
        if tbl.sheet_name != "Data_Meta" and tbl.quarter_labels:
            n_quarters    = len(tbl.quarter_labels)
            quarter_labels = tbl.quarter_labels
            filing_dates   = tbl.filing_dates
            break

    # 品質檢查：9 個關鍵科目，各檢查「最近 4 期是否全部為空」——全空才算缺。
    # 只要有任一期有值就通過，所以 9/9 是「都至少抓到一期」，不代表每期都完整。
    score, missing_txt = "", ""
    q_tbl = next((t for t in tables if t.sheet_name == "Data_Financials(Q)"), None)
    if q_tbl is not None:
        missing = sorted(set(
            check_key_rows(q_tbl.concepts, q_tbl.values, "IS")
            + check_key_rows(q_tbl.concepts, q_tbl.values, "BS")
            + check_key_rows(q_tbl.concepts, q_tbl.values, "CF")
        ))
        from excel_formatter import ALL_KEY_ROWS as _ALL_KEY_ROWS
        total = len(_ALL_KEY_ROWS)
        score = f"{total - len(missing)}/{total}"
        missing_txt = t("xls.meta.sep").join(missing) if missing else t("xls.meta.none")

    # 最新期間：這份檔案的資料抓到哪一季、那一季實際結束在哪天。
    latest_label, latest_end = "", ""
    if q_tbl is not None and q_tbl.quarter_labels:
        latest_label = q_tbl.quarter_labels[-1]
        ends = q_tbl.period_ends or []
        latest_end = ends[-1] if ends and ends[-1] else _period_end(latest_label, fy_end_month)

    # 財年起訖：結算月的下個月為起月。AAPL 9 月結算 → 財年 10 月起。
    start_month = fy_end_month % 12 + 1
    fy_span = t("xls.meta.fy_span_value", start=start_month, end=fy_end_month)

    return StatementTable(
        sheet_name="Data_Meta",
        quarter_labels=quarter_labels,
        filing_dates=filing_dates,
        # Fiscal Year End Month 是換算日曆季的依據——沒有它就無法把不同結算月
        # 公司的 FY 標籤對齊到同一個日曆季，是這張表唯一「程式在用」的欄位。
        # 品質檢查（原本在已移除的 Index sheet）也併到這裡。
        concepts=["Ticker", "Company Name", "Fetched Date", "Quarters Available",
                  "Fiscal Year End Month", "Fiscal Year Span", "Latest Period", "Latest Period End",
                  "Key Rows Complete", "Key Rows Missing", "Fetch Gaps"],
        values=[
            [ticker]            * n_quarters,
            [company_name]      * n_quarters,
            [str(date.today())] * n_quarters,
            [str(n_quarters)]   * n_quarters,
            [str(fy_end_month)] * n_quarters,
            [fy_span]           * n_quarters,
            [latest_label]      * n_quarters,
            [latest_end]        * n_quarters,
            [score]             * n_quarters,
            [missing_txt]       * n_quarters,
            # 抓取缺漏。GUI 的 log 關掉就沒了，但這份 Excel 三天後再打開
            # 還在——使用者真正會搞混的時點是那時候。
            [gap_note or t("xls.meta.none")] * n_quarters,
        ],
    )


# ── Public API ─────────────────────────────────────────────────────────────

def fetch_gaap_statements(ticker: str, identity: str,
                           max_filings: int = 80,
                           max_annual_filings: int = 20,
                           ai_config: dict | None = None,
                           start_year: int | None = None,
                           end_year: int | None = None,
                           fetch_quarterly: bool = True,
                           fetch_annual: bool = True,
                           excluded_sheets: set[str] | None = None) -> list[StatementTable]:
    """Fetch quarterly and/or annual GAAP statements for a ticker.

    Args:
        ticker:              Stock ticker, e.g. "AAPL"
        identity:            SEC EDGAR identity string
        max_filings:         Max 10-Q filings to process (default 80, ~20 years)
        max_annual_filings:  Max 10-K filings to process (default 20, ~20 years)
        ai_config:           AI config dict (provider/model/api_key) for E2 diagnosis
        start_year:          Only include filings from this year onwards (None = no limit)
        end_year:            Only include filings up to this year (None = no limit)
        fetch_quarterly:     Whether to fetch 10-Q data (default True)
        fetch_annual:        Whether to fetch 10-K data (default True)
        excluded_sheets:     Set of sheet names to skip in the output

    Returns:
        List of StatementTable

    Raises:
        ValueError: No filings found for the requested form type(s)
    """
    # 帳本涵蓋整趟。抓不到的期數記在裡面，最後寫進 Data_Meta 的 Fetch Gaps
    # 讓 GUI 與 Excel 的 Index 頁都讀得到。呼叫端不必知道它存在；
    # 想自己拿到明細就在外面包一層 collect_gaps()。
    if _ledger() is None:
        with collect_gaps():
            return fetch_gaap_statements(
                ticker, identity, max_filings, max_annual_filings, ai_config,
                start_year, end_year, fetch_quarterly, fetch_annual,
                excluded_sheets,
            )

    # 解析快取涵蓋整趟（G9）。開在這裡而不是上面那個 `_ledger() is None` 分支
    # 裡面——`main.py` 與 `cli.py` 會自己先開 `collect_gaps()`，那條路不會走到
    # 上面的遞迴，快取就不會生效。`_parse_cache_scope()` 可以巢狀，重複開無害。
    #
    # D11-B：`led` 在這裡一定是「當下在用的那本帳本」，不管是誰開的
    # （上面的遞迴自己開的，或是 `main.py`／`cli.py` 先開好才呼叫進來的）——
    # 兩條路都會走到這裡，重試才不會只在其中一條路生效。
    led = _ledger()
    with _disk_cache_scope(), _parse_cache_scope():
        tables = _fetch_gaap_impl(
            ticker, identity, max_filings, max_annual_filings, ai_config,
            start_year, end_year, fetch_quarterly, fetch_annual, excluded_sheets,
        )

    def _retry_once() -> tuple[list[StatementTable], FetchLedger]:
        with collect_gaps() as retry_led, _disk_cache_scope(), _parse_cache_scope():
            retry_tables = _fetch_gaap_impl(
                ticker, identity, max_filings, max_annual_filings, ai_config,
                start_year, end_year, fetch_quarterly, fetch_annual, excluded_sheets,
            )
        return retry_tables, retry_led

    return _fetch_with_retry(tables, led, _retry_once)


def _fetch_gaap_impl(ticker: str, identity: str,
                     max_filings: int, max_annual_filings: int,
                     ai_config: dict | None,
                     start_year: int | None, end_year: int | None,
                     fetch_quarterly: bool, fetch_annual: bool,
                     excluded_sheets: set | None) -> list[StatementTable]:
    """`fetch_gaap_statements()` 的本體。拆出來只是為了讓帳本與解析快取
    這兩個 context manager 包在外面，不用把 120 行本體整段重新縮排。"""
    ai_config = ai_config or {}
    excluded_sheets = excluded_sheets or set()
    set_identity(identity)
    company = Company(ticker)
    # cik 才是跟 SEC 打交道真正的鍵；ticker 只是會換手的別名。拿不到就不用快取。
    _bind_disk_cache(ticker, getattr(company, "cik", None))

    filings_q = _list_filings(company, "10-Q") if fetch_quarterly else []
    filings_k = _list_filings(company, "10-K") if fetch_annual else []

    if fetch_quarterly and not filings_q:
        raise ValueError(
            f"No 10-Q filings found for ticker '{ticker}'. "
            "The ticker may be invalid or the company may not file 10-Qs."
        )
    if not fetch_quarterly and not filings_k:
        raise ValueError(
            f"No 10-K filings found for ticker '{ticker}'. "
            "The ticker may be invalid or the company may not file 10-Ks."
        )

    # Apply year range filter
    filings_q = _filter_filings_by_year(filings_q, start_year, end_year)
    filings_k = _filter_filings_by_year(filings_k, start_year, end_year)

    # 進度條分母：每份 filing 要建 IS/BS/CF 三張表，各跑一輪 = 3 個 tick。
    # `min(len, max_filings)` 只是上限估計——`_build_*_table` 內部可能因為
    # pre-XBRL 篩掉或重複期別提早結束，真正跑到的份數可能更少，屆時
    # `_tick_progress` 會把顯示值夾在 total 以內，不會超過 100%
    _n_q = min(len(filings_q), max_filings) if fetch_quarterly else 0
    _n_k = min(len(filings_k), max_annual_filings) if fetch_annual else 0
    _set_progress_total((_n_q + _n_k) * 3)

    overrides = load_overrides(ticker)
    if filings_k:
        fy_end_month = _detect_fy_end_month(filings_k)
    elif fetch_quarterly and filings_q:
        _probe_k = _list_filings(company, "10-K")[:1]
        fy_end_month = _detect_fy_end_month(_probe_k) if _probe_k else 12
    else:
        fy_end_month = 12

    tables: list[StatementTable] = []

    # 年報表要先建好——季報表要用它反推 Q4（見 _synthesize_q4）。
    is_ann = bs_ann = cf_ann = None
    is_ann_ng = bs_ann_ng = cf_ann_ng = None
    if fetch_annual and filings_k:
        is_ann, is_ann_ng = _build_is_table(filings_k, max_annual_filings, is_overrides=overrides.get("IS", {}), fy_end_month=fy_end_month)
        bs_ann, bs_ann_ng = _build_bs_table(filings_k, max_annual_filings, bs_overrides=overrides.get("BS", {}), fy_end_month=fy_end_month)
        cf_ann, cf_ann_ng = _build_cf_table(filings_k, max_annual_filings, cf_overrides=overrides.get("CF", {}), fy_end_month=fy_end_month)

    if fetch_quarterly and filings_q:
        is_tbl, is_ng = _build_is_table(filings_q, max_filings, is_overrides=overrides.get("IS", {}), fy_end_month=fy_end_month)
        bs_tbl, bs_ng = _build_bs_table(filings_q, max_filings, bs_overrides=overrides.get("BS", {}), fy_end_month=fy_end_month)
        cf_tbl, cf_ng = _build_cf_table(filings_q, max_filings, cf_overrides=overrides.get("CF", {}), fy_end_month=fy_end_month)

        # Diagnose key rows that are all-None in recent quarters
        missing_is = check_key_rows(is_tbl.concepts, is_tbl.values, "IS")
        missing_bs = check_key_rows(bs_tbl.concepts, bs_tbl.values, "BS")
        missing_cf = check_key_rows(cf_tbl.concepts, cf_tbl.values, "CF")

        if missing_is or missing_bs or missing_cf:
            try:
                tenq_latest = _filing_obj(filings_q[0])
                latest_is_df = tenq_latest.financials.income_statement().to_dataframe()
                latest_bs_df = tenq_latest.financials.balance_sheet().to_dataframe()
                latest_cf_df = tenq_latest.financials.cashflow_statement().to_dataframe()
            except Exception as exc:
                # 這是選用的診斷路徑（補 override），失敗不影響主要資料，
                # 也不算缺一期，所以不記帳本。
                print(f"[{ticker}] 診斷：無法取得最新 filing DataFrame — "
                      f"{type(exc).__name__}", file=sys.stderr)
                latest_is_df = latest_bs_df = latest_cf_df = None

            new_overrides: dict[str, dict] = {}
            if missing_is and latest_is_df is not None:
                fixes = run_diagnosis(ticker, "IS", latest_is_df, missing_is, ai_config)
                if fixes:
                    new_overrides["IS"] = fixes
            if missing_bs and latest_bs_df is not None:
                fixes = run_diagnosis(ticker, "BS", latest_bs_df, missing_bs, ai_config)
                if fixes:
                    new_overrides["BS"] = fixes
            if missing_cf and latest_cf_df is not None:
                fixes = run_diagnosis(ticker, "CF", latest_cf_df, missing_cf, ai_config)
                if fixes:
                    new_overrides["CF"] = fixes

            if new_overrides:
                total_fixes = sum(len(v) for v in new_overrides.values())
                print(f"[{ticker}] 自動修復：找到 {total_fixes} 項缺失指標修復方案，重新建表。", file=sys.stderr)
                overrides = load_overrides(ticker)
                is_tbl, is_ng = _build_is_table(filings_q, max_filings, is_overrides=overrides.get("IS", {}), fy_end_month=fy_end_month)
                bs_tbl, bs_ng = _build_bs_table(filings_q, max_filings, bs_overrides=overrides.get("BS", {}), fy_end_month=fy_end_month)
                cf_tbl, cf_ng = _build_cf_table(filings_q, max_filings, cf_overrides=overrides.get("CF", {}), fy_end_month=fy_end_month)
            else:
                remaining = missing_is + missing_bs + missing_cf
                if remaining:
                    no_key = "" if ai_config.get("api_key") else "（未設 AI API key，E2 診斷已跳過）"
                    print(f"[{ticker}] 警告：{remaining} 在 EDGAR 中無對應概念{no_key}。", file=sys.stderr)

        # Q4 補值（D0-1）：有年報可用時，用年度值反推單季 Q4。
        if is_ann is not None:
            is_tbl = _synthesize_q4(is_tbl, is_ann, len(IS_TEMPLATE), is_balance=False)
            bs_tbl = _synthesize_q4(bs_tbl, bs_ann, len(BS_TEMPLATE), is_balance=True)
            cf_tbl = _synthesize_q4(cf_tbl, cf_ann, len(CF_TEMPLATE), is_balance=False,
                                    point_in_time_idx=_CF_POINT_IN_TIME_IDX)

        # IS 的 CF-sourced 列（SBC / D&A (CF memo)）改用 CF 表已經拆算好的單季值。
        # **一定要放在 _synthesize_q4() 之後**：CF 表的 Q4 是在那一步才補上的，
        # 放在前面的話 IS 那兩列的 Q4 仍然會是空的（見 _backfill_cf_sourced_rows
        # 的說明與 TODO G3）。
        is_tbl = _backfill_cf_sourced_rows(is_tbl, cf_tbl)

        quarterly_tbl = _merge_financials(is_tbl, bs_tbl, cf_tbl, sheet_name="Data_Financials(Q)", fy_end_month=fy_end_month)
        tables.append(quarterly_tbl)
        if any(tbl.concepts for tbl in [is_ng, bs_ng, cf_ng]):
            ng_q_tbl = _merge_financials(is_ng, bs_ng, cf_ng, sheet_name="Data_Financials_NG(Q)", fy_end_month=fy_end_month)
            tables.append(ng_q_tbl)

    if fetch_annual and filings_k:
        annual_tbl = _merge_financials(is_ann, bs_ann, cf_ann, sheet_name="Data_Financials(Y)", fy_end_month=fy_end_month)
        tables.append(annual_tbl)
        if any(tbl.concepts for tbl in [is_ann_ng, bs_ann_ng, cf_ann_ng]):
            ng_y_tbl = _merge_financials(is_ann_ng, bs_ann_ng, cf_ann_ng, sheet_name="Data_Financials_NG(Y)", fy_end_month=fy_end_month)
            tables.append(ng_y_tbl)

    if fetch_quarterly and filings_q:
        seg_tables = _build_segment_tables(filings_q, max_filings, fy_end_month=fy_end_month)
        tables.extend(t for t in seg_tables if t.sheet_name not in excluded_sheets)

    company_name = getattr(company, "name", ticker) or ticker
    # 期末流通股數不在三表裡，走封面頁 dei fact 另外補（見 _fetch_shares_outstanding）
    _apply_shares_outstanding(tables, _fetch_shares_outstanding(company))

    gap_note = _ledger().summary() if _ledger() is not None else ""
    tables.append(_build_meta_table(ticker, company_name, tables, fy_end_month,
                                    gap_note=gap_note))

    for tbl in tables:
        tbl.ticker = ticker
    return tables


def preview_sheets(ticker: str, identity: str) -> dict[str, Any]:
    """Quick scan: fetch only the latest 10-Q to detect segment sheet names
    and report the newest quarter currently available on EDGAR.

    Predicts sheet names + reports latest quarter without a full fetch.
    Takes ~5-15 seconds (one HTTP request for the latest filing).

    Returns:
        {
            "sheets": [...],            # Fixed sheets (Financials Q/Y, Meta)
                                         # always included; Data_Seg_* detected
                                         # from the latest 10-Q.
            "latest_label": "FY2026Q1", # "" if undetectable
            "latest_period_end": "2025-12-27",
            "filing_date": "2026-02-01",  # 送件/公開日，不是 SEC accepted 時間戳
        }
    """
    from fiscal_input import fiscal_quarter_of  # 延後 import：fiscal_input -> excel_formatter -> fetcher_gaap 會循環匯入

    fixed = ["Data_Financials(Q)", "Data_Financials(Y)", "Data_Meta"]
    empty = {"sheets": fixed, "latest_label": "", "latest_period_end": "", "filing_date": ""}

    set_identity(identity)
    company = Company(ticker)
    filings_q = _list_filings(company, "10-Q")
    if not filings_q:
        return empty

    latest = filings_q[0]
    period_end = str(getattr(latest, "period_of_report", "") or "")
    filing_date = str(getattr(latest, "filing_date", "") or "")

    # 財年結束月：走 Company.fiscal_year_end 屬性（一次請求，已經在 company 物件上），
    # 不用 _detect_fy_end_month()——那個要 filing.obj() 抓 10-K 全文，太慢，快速掃描用不起。
    raw_fy = str(getattr(company, "fiscal_year_end", "") or "").strip()
    fy_end_month = int(raw_fy[:2]) if len(raw_fy) == 4 and raw_fy.isdigit() and 1 <= int(raw_fy[:2]) <= 12 else 12
    start_month = fy_end_month % 12 + 1
    latest_label = fiscal_quarter_of(period_end, start_month)

    try:
        seg_tables = _build_segment_tables([latest], max_filings=1)
        seg_names = [t.sheet_name for t in seg_tables]
    except Exception as exc:
        # 掃描是「這家有哪些 sheet」的預覽，缺 segment 不影響後續抓取，
        # 不記帳本也不中止——真正的抓取會自己再判斷一次。
        print(f"[preview_sheets] Segment scan failed: {type(exc).__name__}",
              file=sys.stderr)
        seg_names = []

    return {
        "sheets": fixed + seg_names,
        "latest_label": latest_label,
        "latest_period_end": period_end,
        "filing_date": filing_date,
    }
