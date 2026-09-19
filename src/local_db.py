"""
local_db.py — 本地財報資料庫的**狀態層**（TODO J1–J4）。

設計書：`docs/superpowers/specs/2026-09-04-local-filing-db-design.md`。

`filing_cache.py` 管的是「一份 filing 怎麼存、怎麼讀」，已經是正確的形狀，
這個模組不改它的儲存格式，只在上面補三塊「狀態與體驗」：

- **更新名單**（J1）——`config["local_db_tickers"]`，跟 `watchlist` 分開的第三份
  清單。`watchlist` 是「批次產 Excel 的對象」，更新名單是「要保持新鮮的資料」
- **`_meta.json`**（J2）——一家一份，放該公司資料夾。涵蓋期間、份數、
  `reached_bottom`、上次更新、寫入時的 edgartools 版本
- **「更新本地庫」**（J3）與**版本不符偵測**（J4）

⚠ **不動 `fetcher_gaap` 的抓取迴圈一行。** 「到底了沒」在抓取迴圈外面推導
（比對完整 filing 清單與已快取的 accession），不讓 builder 回報停止原因——
那要穿過 3 個 builder 與 8 個呼叫點，就是 TODO G13 (a) 那個坑。

⚠ **meta 只是快照，事實來源永遠是目錄本身。** 對不上就重建，重建很便宜
（比對只用目錄列舉，不讀檔內容）。`filing_cache.py` 開頭那條「不維護額外索引檔」
的原則沒有被破壞：meta 刪掉、寫壞、跟目錄不同步，功能都照樣正確，只是慢一點。

edgartools 在這裡是**延遲載入**的（只在真的要連網時才 import）——meta 與名單
這兩塊純邏輯不該為了跑一個單元測試去載一個幾秒的套件。
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import filing_cache

META_FILENAME = "_meta.json"
# 2：`forms[].period_oldest`／`period_newest`（真正的財報期間，
# 不是 SEC 收件日）。升版會讓舊 meta 全部判無效並重建——正確，
# 舊快照裡本來就沒有期間欄位。
META_SCHEMA_VERSION = 2

# 一家公司的表單組是「季報表單 ＋ 年報表單」兩格。分開記是必要的——
# `max_filings`（季）與 `max_annual_filings`（年）是兩個獨立上限，一家公司
# 可能年報到底了、季報還沒，合記會誤判成整家到底，然後**永遠不再往下挖**。
#
# 兩組表單（TODO D9）：美國國內申報人交 10-Q／10-K，外國私人發行人（ARM）
# 交 6-K／20-F。⚠ **不能寫死一組**——寫死 10-Q／10-K 的話 FPI 會整組判空：
# 「到底了沒」永遠是 None（每輪重掃到底）、meta 永遠判不可沿用（每輪重建），
# 而且兩種症狀都不會報錯。
FORMS = ("10-Q", "10-K")
FPI_FORMS = ("6-K", "20-F")


def form_set(names) -> tuple[str, str]:
    """這家公司是哪一組表單。看到任何一個 FPI 表單就整組算 FPI，
    認不出來（空的、舊 meta）一律回國內那組——那是 214 家的常態。"""
    seen = {str(n or "") for n in (names or ())}
    return FPI_FORMS if seen & set(FPI_FORMS) else FORMS

# EDGAR 從 2008 才開始要求 XBRL，更早的申報解析不出三張表。
# ⚠ 這個值跟 `fetcher_gaap._XBRL_CUTOFF` 必須一致（`test_local_db.py` 有釘）。
# 沒有直接 import 是為了不讓這個模組在載入時就把 edgartools 拖進來。
XBRL_CUTOFF = date(2008, 1, 1)

# 「更新本地庫」對每家公司開的抓取窗。200/50 不是「要抓 200 份」，是
# 「大到不會是它先喊停」的餘裕值：XBRL 從 2008 起算最多 18 年，
# ≈72 份 10-Q ＋ 18 份 10-K。實際由 `_XBRL_CUTOFF` 或清單用完停止。
DEEP_MAX_FILINGS = 200
DEEP_MAX_ANNUAL_FILINGS = 50

_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


# ── 小工具 ────────────────────────────────────────────────────────────────

def _as_date(value) -> date | None:
    """`"2025-08-01"` → `date`。認不出來回 None，呼叫端一律往「保守」那邊倒。"""
    if isinstance(value, date):
        return value
    text = str(value or "").strip()[:10]
    if not _DATE_RE.match(text):
        return None
    try:
        return date.fromisoformat(text)
    except ValueError:
        return None


def normalize_tickers(items) -> list[str]:
    """去空白、轉大寫、去重（保留出現順序）。None／空字串直接丟掉。"""
    out: list[str] = []
    seen: set[str] = set()
    for item in items or []:
        ticker = str(item or "").strip().upper()
        if not ticker or ticker in seen:
            continue
        seen.add(ticker)
        out.append(ticker)
    return out


# ── J2：reached_bottom 推導（純函式，不連網）──────────────────────────────

def derive_reached_bottom(available, cached_accessions) -> str | None:
    """這個 form 有沒有「抓到底」。

    Args:
        available: 完整 filing 清單，`[(accession, filing_date), ...]`。
            這是 `_list_filings()` 本來就會回傳的東西，所以「更新本地庫」
            跑的時候順手就有，不必額外連網。
        cached_accessions: 該公司資料夾裡已有的 accession 集合。

    Returns:
        `"no_more_filings"` 清單用完了（例如 META 2013 才上市）；
        `"xbrl_cutoff"`     清單還有更舊的，但那些在 2008 之前、沒有 XBRL；
        `None`              還沒抓完，下次要繼續挖。

    ⚠ 日期解析不出來的一律**當成在窗內**（＝還沒抓到就判 None）。誤判「還沒到底」
    的代價是多查一次清單，誤判「到底」的代價是那家公司永遠不再往下挖，
    而且完全沒有症狀。
    """
    cached = set(cached_accessions or ())
    in_window: list[str] = []
    has_older = False
    for accession, filing_date in available or ():
        parsed = _as_date(filing_date)
        if parsed is not None and parsed < XBRL_CUTOFF:
            has_older = True
        else:
            in_window.append(accession)
    if not set(in_window).issubset(cached):
        return None
    return "xbrl_cutoff" if has_older else "no_more_filings"


def _new_accessions(available, cached_accessions) -> list[str]:
    """清單上有、快取裡沒有、而且在 XBRL 窗內的 accession。"""
    cached = set(cached_accessions or ())
    out = []
    for accession, filing_date in available or ():
        parsed = _as_date(filing_date)
        if parsed is not None and parsed < XBRL_CUTOFF:
            continue
        if accession not in cached:
            out.append(accession)
    return out


def plan_ticker(listings: dict, cached_accessions, *, version_ok: bool = True) -> dict:
    """這家公司這一輪要不要進抓取迴圈。

    `skip=True` 就是設計書裡「不要每次全部重抓」的具體實現：兩個 form 都到底、
    又沒有新 filing 的公司，整家跳過，只花一次 filing 清單的網路。

    `version_ok=False`（快取是別的 edgartools 版本解出來的）**一律不跳過**——
    那些檔案在 `load_filing()` 眼裡等同不存在，跳過會讓那家公司永遠停在失效狀態。
    """
    cached = set(cached_accessions or ())
    forms: dict[str, dict] = {}
    new_count = 0
    used = form_set(listings.keys())
    for form in used:
        available = listings.get(form) or []
        new = _new_accessions(available, cached)
        new_count += len(new)
        forms[form] = {
            "reached_bottom": derive_reached_bottom(available, cached),
            "available": len(available),
            "new": len(new),
        }
    skip = (version_ok and new_count == 0
            and all(forms[f]["reached_bottom"] is not None for f in used))
    return {"skip": skip, "new_count": new_count, "forms": forms}


# ── J2：_meta.json ────────────────────────────────────────────────────────

def meta_path(ticker: str) -> Path:
    return filing_cache.ticker_dir(ticker) / META_FILENAME


def cached_accessions(ticker: str) -> set[str]:
    """該公司已快取的 accession 集合。**只做目錄列舉，不讀檔內容**——
    201 家若每次都要讀 881 個 JSON 才能顯示清單，GUI 會卡住。"""
    try:
        paths = list(filing_cache.ticker_dir(ticker).glob("*.json"))
    except OSError:
        return set()
    return {p.stem for p in paths if filing_cache.ACCESSION_RE.match(p.stem)}


_PERIOD_COL_RE = re.compile(r"(\d{4}-\d{2}-\d{2})")


def period_end_of(entry: dict) -> str | None:
    """一份 filing 涵蓋到的**財報期末日**，從快取的 DataFrame 欄名抽出來。

    edgartools 的欄名是 `"2026-03-29 (Q1)"`（損益表／現金流量表，duration）
    或 `"2024-03-31"`（資產負債表，instant）。**取最大值**＝這份申報的當期：
    其餘日期欄是去年同期、YTD 等比較欄，一定比當期舊。

    ⚠ 這跟 `filing_date`（SEC 收件日）是兩件事，差一整個期間——一份 2008-05
    申報的 10-K 蓋的是 2007 年度。`_meta.json` 原本只存 `filing_date`，
    「我有哪幾季的數字」這個問題答不出來，2026-09-18 才補上這個。

    抓不到回 None（pre-XBRL 的負向快取 `dataframes` 是 null，本來就沒有期間）。
    """
    frames = (entry or {}).get("dataframes") or {}
    best: str | None = None
    for payload in frames.values():
        columns = ((payload or {}).get("data") or {}).get("columns") or []
        for col in columns:
            m = _PERIOD_COL_RE.match(str(col).strip())
            if m and (best is None or m.group(1) > best):
                best = m.group(1)
    return best


def scan_filings(ticker: str) -> list[dict]:
    """讀出該公司每一份快取檔的 (accession, form, filing_date, 期末日, 版本)。

    比 `cached_accessions()` 貴得多（要真的開檔），**只在重建 meta 時走**。
    壞掉的檔案跳過不算——它在 `load_filing()` 那邊也一樣會被判無效。

    ⚠ 這支是 `rebuild_meta` 慢的元兇（實測 2.75 秒/家）：為了幾個小欄位把
    每份 70KB 的 JSON 整個 `json.load()` 進來。**但期末日是免費的**——
    整份都已經解進記憶體了，多讀一組欄名不花額外成本。
    """
    rows: list[dict] = []
    for accession in sorted(cached_accessions(ticker)):
        path = filing_cache.ticker_dir(ticker) / f"{accession}.json"
        try:
            with open(path, "r", encoding="utf-8") as f:
                entry = json.load(f)
        except (OSError, ValueError):
            continue
        if not isinstance(entry, dict):
            continue
        rows.append({
            "accession": accession,
            "form": str(entry.get("form") or ""),
            "filing_date": str(entry.get("filing_date") or ""),
            "period_end": period_end_of(entry),
            "cached_at": str(entry.get("cached_at") or ""),
            "edgartools_version": entry.get("edgartools_version"),
            "cik": entry.get("cik"),
        })
    return rows


def read_meta(ticker: str) -> dict | None:
    """原封不動讀回 `_meta.json`。不存在／壞掉／schema 不符一律回 None。
    不做自癒——那是 `load_meta()` 的事。"""
    path = meta_path(ticker)
    try:
        with open(path, "r", encoding="utf-8") as f:
            meta = json.load(f)
    except (OSError, ValueError):
        return None
    if not isinstance(meta, dict):
        return None
    if meta.get("schema_version") != META_SCHEMA_VERSION:
        return None
    return meta


def write_meta(ticker: str, meta: dict) -> bool:
    """寫 `_meta.json`。走 `atomic_write_json()`，跟 filing 同一套——
    多視窗同時跑時不會寫到一半被讀走。寫失敗只回 False，不拋。"""
    return filing_cache.atomic_write_json(meta_path(ticker), meta)


def rebuild_meta(ticker: str, previous: dict | None = None) -> dict:
    """掃目錄重建 meta。**目錄是事實來源，meta 只是快照。**

    `reached_bottom` 重算要連網拿完整 filing 清單，這裡不做——保留上一版的值
    並標記 `reached_bottom_stale`，下次「更新本地庫」跑到這家時再重算。
    """
    ticker = (ticker or "").strip().upper()
    rows = scan_filings(ticker)
    prev_forms = (previous or {}).get("forms") or {}
    forms: dict[str, dict] = {}
    # 目錄是事實來源，所以表單組先看快取檔實際是什麼 form；目錄空的（清過、
    # 全是負向快取）才退回上一版 meta 記的那組。
    for form in form_set([r["form"] for r in rows] or list(prev_forms)):
        dates = sorted(r["filing_date"] for r in rows
                       if r["form"] == form and r["filing_date"])
        count = sum(1 for r in rows if r["form"] == form)
        old = prev_forms.get(form) or {}
        carried = old.get("reached_bottom")
        # 期末日：pre-XBRL 的負向快取沒有期間，濾掉 None 再取範圍
        periods = sorted(r["period_end"] for r in rows
                         if r["form"] == form and r["period_end"])
        forms[form] = {
            "count": count,
            # ⚠ oldest／newest 是 **SEC 收件日**，period_* 才是財報期間。
            # 兩個都留著——「上次申報是什麼時候」與「我有哪幾季的數字」
            # 是不同的問題，而且差一整個期間。
            "oldest": dates[0] if dates else None,
            "newest": dates[-1] if dates else None,
            "period_oldest": periods[0] if periods else None,
            "period_newest": periods[-1] if periods else None,
            "reached_bottom": carried,
            # 帶著舊值就一定標過期——目錄跟 meta 對不上代表份數變了，
            # 「到底了沒」很可能也跟著變。
            "reached_bottom_stale": carried is not None,
        }
    # 版本混雜時取「最近寫進去的那一份」的版本：`load_filing()` 是逐份比對的，
    # 混雜狀態下這裡填哪一個都不完全準，取最新的至少反映最後一次抓取。
    newest_row = max(rows, key=lambda r: r["cached_at"], default=None)
    ciks = {r["cik"] for r in rows if r["cik"] is not None}
    return {
        "schema_version": META_SCHEMA_VERSION,
        "ticker": ticker,
        "cik": ciks.pop() if len(ciks) == 1 else None,
        "file_count": len(rows),
        "updated_at": filing_cache._now_iso(),
        "edgartools_version": newest_row["edgartools_version"] if newest_row else None,
        "forms": forms,
    }


def load_meta(ticker: str) -> dict | None:
    """拿這家公司的 meta，跟目錄對不上就當場重建並寫回。

    快取路徑（`file_count` 對得上）只做一次目錄列舉，不開任何檔——GUI 列 201 家
    才不會卡住。沒有任何快取檔的公司回 None。
    """
    ticker = (ticker or "").strip().upper()
    accessions = cached_accessions(ticker)
    meta = read_meta(ticker)
    if not accessions:
        return meta if meta and meta.get("file_count") == 0 else None
    if meta is not None and meta.get("file_count") == len(accessions):
        return meta
    rebuilt = rebuild_meta(ticker, previous=meta)
    write_meta(ticker, rebuilt)
    return rebuilt


# ── J7：資料庫總覽（GUI 分頁與 CLI `db-status` 共用的純資料層）────────────
#
# ⚠ **這一整段是唯讀的**，刻意走 `read_meta()` 而不是 `load_meta()`：後者對不上
# 會當場 `rebuild_meta()` 並寫檔，一家要開 75 個 JSON——244 家全部過期時，
# 「看一眼資料庫有什麼」會變成上萬次檔案讀取加 244 次寫入。總覽頁與 CLI 查詢
# 都是「瞄一眼」的操作，不該有副作用，對不上就照實說「需重算」，把重算留給
# 真正會動資料的「更新本地庫」。
#
# ⚠ 回傳的都是**英文機器鍵**（`bottom` 的 yes／no／unknown 等），顯示文字由
# 呼叫端查 i18n——跟 `Data_Ratios` A 欄／B 欄那套同一個原則。

BOTTOM_YES = "yes"
BOTTOM_NO = "no"
BOTTOM_UNKNOWN = "unknown"

OVERVIEW_SORT_KEYS = ("ticker", "filings", "size_bytes", "filed_from",
                      "period_from", "years", "stale_days")


def _span_years(filed_from: str | None, filed_to: str | None) -> float | None:
    """兩個申報日之間的年數，一位小數。`None` 代表算不出來。"""
    start, end = _as_date(filed_from), _as_date(filed_to)
    if start is None or end is None:
        return None
    return round((end - start).days / 365.25, 1)


def days_since(iso_timestamp: str | None, *, today: date | None = None) -> int | None:
    """`updated_at`（帶時區的 ISO 字串）→ 距今幾天。認不得回 None。

    **為什麼要這個**：`updated_at` 是「上次去 SEC 查這家」的時間，跟
    `forms[].newest`（最新申報日）是兩件事。一家顯示「已到底、最新申報
    2026-06」，如果那是三個月前查的，中間很可能已經出了新財報而我們不知道。
    2026-09-18 之前這個欄位**寫了但從來沒有人讀**，等於白存。
    """
    raw = str(iso_timestamp or "").strip()
    if not raw:
        return None
    try:
        # `_now_iso()` 產的是帶時區偏移的字串，比較前先轉成本地日期
        stamp = datetime.fromisoformat(raw)
    except ValueError:
        return None
    seen = stamp.date()
    reference = today or date.today()
    return max(0, (reference - seen).days)


def overview_row(ticker: str, cached: dict, in_list: bool) -> dict:
    """一家公司在總覽頁的一列。`cached` 是 `list_cached_tickers()` 的那個 dict。

    ⚠ `filed_from`／`filed_to` 是 **SEC 收件日**，不是財報期間——`_meta.json`
    存的就是 `filing_date`（見 `rebuild_meta`）。一份 2008-05 申報的 10-K 蓋的
    是 2007 年度，兩者差一整個期間。顯示端必須標成「申報日期」，標成「涵蓋
    期間」會變成對財報分析師講假話。真正的財報期間在 Excel 的 `Data_Meta`
    （`Oldest Period`／`Latest Period`，TODO J6 方向②）。
    """
    meta = read_meta(ticker)
    forms = (meta or {}).get("forms") or {}
    dates = [d for f in forms.values()
             for d in (f.get("oldest"), f.get("newest")) if d]
    filed_from = min(dates) if dates else None
    filed_to = max(dates) if dates else None
    periods = [d for f in forms.values()
               for d in (f.get("period_oldest"), f.get("period_newest")) if d]
    period_from = min(periods) if periods else None
    period_to = max(periods) if periods else None

    used = form_set(forms)
    states = [forms.get(f, {}).get("reached_bottom") for f in used]
    if not forms:
        bottom = BOTTOM_UNKNOWN
    elif all(s is not None for s in states):
        bottom = BOTTOM_YES
    else:
        bottom = BOTTOM_NO

    return {
        "ticker": ticker,
        "filings": cached["count"],
        "size_bytes": cached["size_bytes"],
        "filed_from": filed_from,
        "filed_to": filed_to,
        "period_from": period_from,
        "period_to": period_to,
        # 年數改用**財報期間**算——那才是「我手上有幾年的數字」。
        # 期間抓不到（舊 schema 或全是負向快取）才退回用申報日估。
        "years": (_span_years(period_from, period_to)
                  if period_from and period_to
                  else _span_years(filed_from, filed_to)),
        "bottom": bottom,
        # 上一輪留下來的值，這輪還沒重算——顯示端加問號
        "bottom_stale": any(forms.get(f, {}).get("reached_bottom_stale")
                            for f in used),
        "in_list": in_list,
        # meta 跟目錄對不上（或根本沒有 meta）。唯讀路徑不自癒，照實回報。
        "meta_ok": bool(meta) and meta.get("file_count") == cached["count"],
        "edgartools_version": (meta or {}).get("edgartools_version"),
        # 「上次去 SEC 查這家」的時間。跟 filed_to（最新申報日）是兩件事：
        # 前者是我們的動作，後者是公司的動作。
        "updated_at": (meta or {}).get("updated_at"),
        "stale_days": days_since((meta or {}).get("updated_at")),
    }


def overview_rows(cfg: dict | None = None) -> list[dict]:
    """資料庫總覽的全部列。預設照 ticker 字母序——總覽頁要的是「找得到某一家」，
    跟 Tab3 舊面板照容量排序（要的是「誰佔空間」）的用途不同。"""
    in_list = set(get_update_list(cfg or {}))
    rows = [overview_row(row["ticker"], row, row["ticker"] in in_list)
            for row in filing_cache.list_cached_tickers()]
    rows.sort(key=lambda r: r["ticker"])
    return rows


def overview_summary(rows) -> dict:
    return {
        "companies": len(rows),
        "filings": sum(r["filings"] for r in rows),
        "size_bytes": sum(r["size_bytes"] for r in rows),
    }


def filter_overview_rows(rows, query: str):
    """依 ticker 篩選，不分大小寫、比對「包含」而非「開頭」——記得是 `AMD`
    但打成 `md` 也要找得到。空字串回全部。"""
    needle = (query or "").strip().upper()
    if not needle:
        return list(rows)
    return [r for r in rows if needle in r["ticker"]]


def sort_overview_rows(rows, key: str = "ticker", descending: bool = False):
    """依欄位排序。不認得的 key 退回 ticker，不拋——排序是顯示層的事，
    使用者點壞一個欄位標題不該讓整頁炸掉。

    `None` 一律排在最後（不管升冪降冪）：算不出年數、沒有申報日的公司是
    「資料不全」，把它們夾在中間會讓真正的極端值被埋起來。
    """
    if key not in OVERVIEW_SORT_KEYS:
        key = "ticker"

    def sort_key(row):
        value = row.get(key)
        return (value is None, value if value is not None else "")

    ordered = sorted(rows, key=sort_key, reverse=descending)
    if descending:
        # reverse=True 會把「None 排最後」也一起翻過來，撥回去
        missing = [r for r in ordered if r.get(key) is None]
        present = [r for r in ordered if r.get(key) is not None]
        ordered = present + missing
    return ordered


# ── J1：更新名單 ──────────────────────────────────────────────────────────

UPDATE_LIST_KEY = "local_db_tickers"


def get_update_list(cfg: dict) -> list[str]:
    return normalize_tickers((cfg or {}).get(UPDATE_LIST_KEY) or [])


def set_update_list(cfg: dict, tickers) -> list[str]:
    cfg[UPDATE_LIST_KEY] = normalize_tickers(tickers)
    return cfg[UPDATE_LIST_KEY]


def add_tickers(cfg: dict, tickers) -> list[str]:
    """加進更新名單，回傳**真正新加的**那幾個（給 GUI 報「新增了 N 家」用）。"""
    current = get_update_list(cfg)
    existing = set(current)
    added = [t for t in normalize_tickers(tickers) if t not in existing]
    cfg[UPDATE_LIST_KEY] = current + added
    return added


def remove_ticker(cfg: dict, ticker: str) -> bool:
    target = str(ticker or "").strip().upper()
    current = get_update_list(cfg)
    if target not in current:
        return False
    cfg[UPDATE_LIST_KEY] = [t for t in current if t != target]
    return True


def import_from_watchlist(cfg: dict) -> list[str]:
    """便利動作一：把 watchlist 全部加進更新名單。

    兩份名單刻意分開（合併會讓 Tab 2 一按產 201 份 Excel），但分開維護很煩，
    所以給一鍵匯入。watchlist 的元素是 `{"ticker": ..., "name": ...}`；
    容忍純字串是為了不讓一個舊格式的 config 炸掉整個功能。
    """
    tickers = []
    for item in (cfg or {}).get("watchlist") or []:
        tickers.append(item.get("ticker") if isinstance(item, dict) else item)
    return add_tickers(cfg, tickers)


def import_from_cache(cfg: dict) -> list[str]:
    """便利動作二：把快取裡已有的公司全部加進更新名單。"""
    return add_tickers(cfg, [r["ticker"] for r in filing_cache.list_cached_tickers()])


# ── J4：版本鎖與版本不符偵測 ──────────────────────────────────────────────

# 重抓一份 filing 大約要多久（秒）。**實測值，取冷跑、取保守的那個**：
# 2026-09-04 連續抓 15 家沒抓過的公司拓到底，整趟 783 份／1,930s＝2.46 s/份，
# 其中中段 900 秒的窗是 321 份／2.8 s/份。取 2.8——這個數字只用來估
# 「重抓要幾小時」，估太樂觀比估太保守糟。
#
# ⚠ 不要用「對 META 量到的 1.8 s/份」——那家在 `~/.edgar/_tcache`（edgartools
# 自己那層持久化 HTTP 快取，跟本專案的 filing_cache 完全獨立、清除動作也碰不到）
# 裡已經是熱的，量到的是「本地重解析」不是「對 SEC 重新抓一次」。
# ARCHITECTURE.md 記過同一個坑讓第一次的快取效能量測整組作廢。
#
# 只拿來估「重抓要幾小時」給使用者參考，估錯不影響任何正確性。
SECONDS_PER_FILING = 2.8


def pinned_edgartools_version() -> str | None:
    """`requirements.txt` 裡鎖的版本。沒鎖或讀不到回 None。"""
    req = Path(__file__).parent.parent / "requirements.txt"
    try:
        text = req.read_text(encoding="utf-8")
    except OSError:
        return None
    for line in text.splitlines():
        line = line.strip()
        if line.startswith("edgartools=="):
            return line.split("==", 1)[1].strip()
    return None


def stale_cache_summary() -> dict:
    """哪些公司的快取是**別的 edgartools 版本**解出來的（J4）。

    `load_filing()` 拿存檔時記的版本跟現在安裝的做字串完全比對，不符就回 None
    ——`5.29.0 → 5.29.1` 也全滅。這個嚴格度是刻意的：快取存的是「那個版本的
    parser 吐出來的 DataFrame」，edgartools 修了解析 bug 的話，舊快取裡的數字
    就是帶著那個 bug 的，**而且不會報錯，只是數字錯**。

    取不到目前版本時回空——這時 `load_filing()` 本來就整個停用快取，
    再跳一個「全部要重抓」的對話框只是嚇人。
    """
    current = filing_cache.edgartools_version()
    empty = {"current": current, "companies": [], "n_companies": 0, "n_filings": 0,
             "size_bytes": 0, "old_versions": [], "estimated_seconds": 0}
    if not current:
        return empty
    companies: list[str] = []
    n_filings = 0
    size_bytes = 0
    old_versions: set[str] = set()
    for row in filing_cache.list_cached_tickers():
        # ⚠ **走 meta，不要逐份開檔。** 這個函式是在**程式啟動時**跑的，
        # 而 201 家拓到底是 ≈16,000 份檔案——原本的寫法對每一家呼叫
        # `scan_filings()`，等於每次開程式都把整個本地庫讀一遍才畫得出主視窗。
        # meta 本來就記著版本，`load_meta()` 在 meta 新鮮時只做一次目錄列舉。
        meta = load_meta(row["ticker"])
        version = (meta or {}).get("edgartools_version")
        # meta 拿不到版本（剛清空、或舊版留下的）就當它是相符的：這裡的職責是
        # 「講出確定失效的部分」，寧可少報也不要拿不確定的東西嚇人。真的失效的
        # 那幾份在 `load_filing()` 那道閘還是會被擋下來重抓，不會用到錯的數字。
        if version is None or version == current:
            continue
        companies.append(row["ticker"])
        n_filings += row["count"]
        size_bytes += row["size_bytes"]
        old_versions.add(str(version))
    return {
        "current": current,
        "companies": companies,
        "n_companies": len(companies),
        "n_filings": n_filings,
        "size_bytes": size_bytes,
        "old_versions": sorted(old_versions),
        "estimated_seconds": int(n_filings * SECONDS_PER_FILING),
    }


# ── J3：更新本地庫 ────────────────────────────────────────────────────────

@dataclass
class TickerResult:
    ticker: str
    status: str                   # "skipped" | "updated" | "failed"
    new_filings: int = 0
    error: str = ""
    gaps: int = 0
    forms: dict = field(default_factory=dict)


@dataclass
class UpdateReport:
    results: list[TickerResult] = field(default_factory=list)
    stopped: bool = False

    def _count(self, status: str) -> int:
        return sum(1 for r in self.results if r.status == status)

    @property
    def skipped(self) -> int:
        return self._count("skipped")

    @property
    def updated(self) -> int:
        return self._count("updated")

    @property
    def failed(self) -> int:
        return self._count("failed")

    @property
    def gap_tickers(self) -> list[str]:
        """有抓取缺漏的公司（TODO D11：連續大量抓取時 SEC 會偶發失敗、
        **靜默少格**）。這些之後單獨重跑即可——第二輪會從本地快取讀已經成功
        的部分，只重抓失敗那幾份。"""
        return [r.ticker for r in self.results if r.gaps]

    def summary(self) -> str:
        return (f"updated={self.updated} skipped={self.skipped} "
                f"failed={self.failed} gaps={len(self.gap_tickers)}"
                + (" stopped=1" if self.stopped else ""))


def _meta_is_reusable(meta: dict | None, file_count: int) -> bool:
    """跳過那條路可不可以直接沿用這份 meta，不重建。

    要求份數對得上、而且**兩個 form 的完整統計都在**。少一個 form（舊版寫的、
    手改過的、寫到一半的）就退回重建——沿用會生出一個只有 `reached_bottom`、
    沒有 `count`／`oldest`／`newest` 的殘缺條目，GUI 那一列會顯示成「—」，
    而且份數對得上所以**下次 `load_meta()` 不會自癒它**，會一直錯下去。
    """
    if not isinstance(meta, dict) or meta.get("file_count") != file_count:
        return False
    forms = meta.get("forms")
    if not isinstance(forms, dict):
        return False
    return all(isinstance(forms.get(f), dict) and "count" in forms[f]
               for f in form_set(forms))


def _default_list_filings(ticker: str, identity: str) -> tuple[dict, int | None]:
    """真的去 EDGAR 拿完整 filing 清單。一家一次網路，很便宜。

    ⚠ 這是整個「跳過」判斷唯一的網路成本。`_list_filings()` 本身帶退避重試
    （2026-08-25 實測 201 家重建撞到 6 家逾時全部發生在這一步）。
    """
    import fetcher_gaap
    fetcher_gaap.set_identity(identity)
    company = fetcher_gaap.Company(ticker)

    # ⚠ **表單組要跟抓取端用同一個判斷**（`_resolve_listings`）。各自判的話，
    # 清單這邊列 A 組、抓取那邊抓 B 組，`derive_reached_bottom()` 比對的就是
    # 兩份不相干的東西——永遠判不到底（每輪重抓）或永遠判到底（永遠不補），
    # 兩種都不報錯。6-K 的 R 檔過濾也在那支裡面，這裡自然跟著一致。
    resolved = fetcher_gaap._resolve_listings(
        ticker, lambda form: fetcher_gaap._list_filings(company, form))

    def _rows(filings):
        out = []
        for filing in filings:
            accession = fetcher_gaap._cache_key(filing)
            if accession is None:
                continue
            out.append((accession, str(getattr(filing, "filing_date", "") or "")))
        return out

    quarterly_form, annual_form = resolved["forms"]
    listings = {quarterly_form: _rows(resolved["quarterly"]),
                annual_form: _rows(resolved["annual"])}
    return listings, getattr(company, "cik", None)


def _default_fetch(ticker: str, identity: str, max_filings: int,
                   max_annual_filings: int):
    """跑一趟抓取，**結果直接丟棄**——要的是它的副作用（把快取填滿）。

    自己開 `collect_gaps()` 才拿得到這一家的缺漏帳本；`fetch_gaap_statements()`
    沒開的話會自己遞迴開一本，我們就看不到了。
    """
    from fetcher_gaap import collect_gaps, fetch_gaap_statements
    with collect_gaps() as ledger:
        fetch_gaap_statements(ticker, identity, max_filings=max_filings,
                              max_annual_filings=max_annual_filings)
    return ledger


def update_local_db(tickers, identity: str, *,
                    progress=None,
                    should_stop=None,
                    list_filings=None,
                    fetch=None,
                    max_filings: int = DEEP_MAX_FILINGS,
                    max_annual_filings: int = DEEP_MAX_ANNUAL_FILINGS) -> UpdateReport:
    """把更新名單上的公司一路抓到底，只暖快取**不產 Excel**。

    每家公司四步（設計書第六節）：
      1. 拿完整 filing 清單（一次網路）
      2. 跟快取現況比對 → 到底又沒有新財報就**整家跳過**
      3. `fetch_gaap_statements()` 拓到底，結果丟棄
      4. 用步驟 1 的清單重算 `_meta.json`

    **單一公司失敗不中斷整體**，記錄後繼續下一家（比照 `comparison.py` 的
    `CompanyFetchError` 原則：公司層級跳過，跟同一家公司內部的科目缺漏是兩回事）。

    `list_filings` / `fetch` 是注入點，測試換掉就能完全離線跑。
    `should_stop()` 每家開始前問一次——GUI 關視窗、CLI Ctrl-C 時停得下來，
    而且不會白費：`save_filing()` 是逐份即時落檔的。
    """
    list_filings = list_filings or _default_list_filings
    fetch = fetch or _default_fetch
    targets = normalize_tickers(tickers)
    report = UpdateReport()

    def emit(event: str, **kw):
        if progress is not None:
            progress({"event": event, **kw})

    emit("start", total=len(targets))
    current_version = filing_cache.edgartools_version()

    for index, ticker in enumerate(targets):
        if should_stop is not None and should_stop():
            report.stopped = True
            break
        emit("ticker_start", ticker=ticker, index=index, total=len(targets))
        try:
            listings, cik = list_filings(ticker, identity)
        except Exception as exc:                      # noqa: BLE001 — 一家壞掉不能拖垮整批
            report.results.append(TickerResult(ticker, "failed",
                                               error=f"{type(exc).__name__}: {exc}"))
            emit("ticker_done", ticker=ticker, status="failed", index=index,
                 total=len(targets))
            continue

        cached = cached_accessions(ticker)
        # 用 `load_meta()`（會自癒）不是 `read_meta()`：既有的 34 家在這個功能
        # 上線前沒有 meta，用 raw 讀會全部判成「版本不明 → 不可跳過」，
        # 等於第一輪一定全部進抓取迴圈。自癒重建會去讀檔拿到真正的版本。
        meta = load_meta(ticker)
        version_ok = bool(current_version) and (
            not cached or (meta or {}).get("edgartools_version") == current_version)
        plan = plan_ticker(listings, cached, version_ok=version_ok)

        status = "skipped"
        gaps = 0
        error = ""
        if plan["skip"]:
            emit("ticker_skip", ticker=ticker, index=index, total=len(targets))
        else:
            try:
                ledger = fetch(ticker, identity, max_filings, max_annual_filings)
                gaps = len(getattr(ledger, "gaps", ()) or ())
                status = "updated"
            except Exception as exc:                  # noqa: BLE001
                status = "failed"
                error = f"{type(exc).__name__}: {exc}"

        if status != "failed":
            # 步驟 4：用步驟 1 的完整清單重算 meta。這裡的 `reached_bottom`
            # 是**新鮮的**（剛連網拿到清單），所以 stale 旗標清掉。
            new_cached = cached_accessions(ticker)
            # 跳過的公司**不重建 meta**——「整家跳過」要真的便宜。重建會把那家
            # 的 75 份檔案全部開一遍，201 家就是 16,000 次開檔，跳過的意義少一半。
            # 上面的 `load_meta()` 已經確認過它跟目錄對得上，份數沒變、內容沒變，
            # 唯一要更新的就是下面那圈剛算出來的 `reached_bottom` 與時間戳。
            # ⚠ 表單組一律看**這一輪的清單**，不要寫死 10-Q／10-K——FPI 的
            # meta 裡根本沒有那兩個鍵，寫死會整家 `KeyError`（2026-09-19 實跑
            # `cli.py update-db ARM` 撞到，12 份檔案照樣落地但沒有 meta）。
            used_forms = form_set(listings.keys())
            if plan["skip"] and _meta_is_reusable(meta, len(new_cached)):
                forms = meta["forms"]
                meta = dict(meta)
                meta["forms"] = {f: dict(forms[f]) for f in used_forms}
                meta["updated_at"] = filing_cache._now_iso()
            else:
                meta = rebuild_meta(ticker, previous=read_meta(ticker))
            if cik is not None:
                try:
                    meta["cik"] = int(cik)
                except (TypeError, ValueError):
                    pass
            for form in used_forms:
                meta["forms"].setdefault(form, {"count": 0})
                meta["forms"][form]["reached_bottom"] = derive_reached_bottom(
                    listings.get(form) or [], new_cached)
                meta["forms"][form]["reached_bottom_stale"] = False
            write_meta(ticker, meta)

        report.results.append(TickerResult(
            ticker, status, new_filings=plan["new_count"], error=error, gaps=gaps,
            forms={f: r["reached_bottom"] for f, r in plan["forms"].items()}))
        emit("ticker_done", ticker=ticker, status=status, index=index,
             total=len(targets), new_filings=plan["new_count"], gaps=gaps,
             error=error)

    emit("done", updated=report.updated, skipped=report.skipped,
         failed=report.failed, stopped=report.stopped)
    return report
