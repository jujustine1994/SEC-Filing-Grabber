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
from datetime import date
from pathlib import Path

import filing_cache

META_FILENAME = "_meta.json"
META_SCHEMA_VERSION = 1

# 這個資料庫只認兩種表單：10-Q 與 10-K。分開記是必要的——`max_filings`（10-Q）
# 與 `max_annual_filings`（10-K）是兩個獨立上限，一家公司可能 10-K 到底了、
# 10-Q 還沒，合記會誤判成整家到底，然後**永遠不再往下挖**。
FORMS = ("10-Q", "10-K")

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
    for form in FORMS:
        available = listings.get(form) or []
        new = _new_accessions(available, cached)
        new_count += len(new)
        forms[form] = {
            "reached_bottom": derive_reached_bottom(available, cached),
            "available": len(available),
            "new": len(new),
        }
    skip = (version_ok and new_count == 0
            and all(forms[f]["reached_bottom"] is not None for f in FORMS))
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


def scan_filings(ticker: str) -> list[dict]:
    """讀出該公司每一份快取檔的 (accession, form, filing_date, 版本)。

    比 `cached_accessions()` 貴得多（要真的開檔），**只在重建 meta 時走**。
    壞掉的檔案跳過不算——它在 `load_filing()` 那邊也一樣會被判無效。
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
    for form in FORMS:
        dates = sorted(r["filing_date"] for r in rows
                       if r["form"] == form and r["filing_date"])
        count = sum(1 for r in rows if r["form"] == form)
        old = prev_forms.get(form) or {}
        carried = old.get("reached_bottom")
        forms[form] = {
            "count": count,
            "oldest": dates[0] if dates else None,
            "newest": dates[-1] if dates else None,
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

# 重抓一份 filing 大約要多久（秒）。**實測值，取冷跑那個**：2026-09-04
# 連續抓 15 家沒抓過的公司，取中段 900 秒的窗量到 321 份 → **2.8 s/份**。
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
        stale = [r for r in scan_filings(row["ticker"])
                 if r["edgartools_version"] != current]
        if not stale:
            continue
        companies.append(row["ticker"])
        n_filings += len(stale)
        size_bytes += row["size_bytes"]
        old_versions.update(str(r["edgartools_version"]) for r in stale)
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


def _default_list_filings(ticker: str, identity: str) -> tuple[dict, int | None]:
    """真的去 EDGAR 拿完整 filing 清單。一家一次網路，很便宜。

    ⚠ 這是整個「跳過」判斷唯一的網路成本。`_list_filings()` 本身帶退避重試
    （2026-08-25 實測 201 家重建撞到 6 家逾時全部發生在這一步）。
    """
    from fetcher_gaap import Company, _cache_key, _list_filings, set_identity
    set_identity(identity)
    company = Company(ticker)
    listings = {}
    for form in FORMS:
        rows = []
        for filing in _list_filings(company, form):
            accession = _cache_key(filing)
            if accession is None:
                continue
            rows.append((accession, str(getattr(filing, "filing_date", "") or "")))
        listings[form] = rows
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
            meta = rebuild_meta(ticker, previous=read_meta(ticker))
            if cik is not None:
                try:
                    meta["cik"] = int(cik)
                except (TypeError, ValueError):
                    pass
            for form in FORMS:
                meta["forms"][form]["reached_bottom"] = derive_reached_bottom(
                    listings.get(form) or [], new_cached)
                meta["forms"][form]["reached_bottom_stale"] = False
            write_meta(ticker, meta)

        report.results.append(TickerResult(
            ticker, status, new_filings=plan["new_count"], error=error, gaps=gaps,
            forms={f: plan["forms"][f]["reached_bottom"] for f in FORMS}))
        emit("ticker_done", ticker=ticker, status=status, index=index,
             total=len(targets), new_filings=plan["new_count"], gaps=gaps,
             error=error)

    emit("done", updated=report.updated, skipped=report.skipped,
         failed=report.failed, stopped=report.stopped)
    return report
