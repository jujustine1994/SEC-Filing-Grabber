"""
filing_cache.py — 本地 filing 解析快取（%APPDATA%\\SEC Financial Tools\\filing_cache）。

快取卡在**解析層與比對層之間**：存的是 edgartools 解出來的三張 DataFrame
（income statement / balance sheet / cashflow statement），比對層
（`IS/BS/CF_TEMPLATE` 那套科目對照）永遠在快取之上即時重跑。所以以後改
hint regex、加比率、調 Q4 合成邏輯都不會讓快取失效——但 **edgartools 升版
會**，那是另一條軸線，靠 `edgartools_version` 欄位擋（見 `load_filing`）。

事實來源是 `<accession>.json` 檔案本身，也是唯一的落地狀態——「哪些公司有
快取」直接掃 `filing_cache/` 底下有哪些子資料夾回答（見 `list_cached_tickers()`），
不維護額外的索引檔。
"""
from __future__ import annotations

import json
import os
import re
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd

SCHEMA_VERSION = 1

# SEC 的 accession number 格式固定，拿來當檔名前先驗——這同時是路徑注入的防線。
ACCESSION_RE = re.compile(r"^\d{10}-\d{2}-\d{6}$")

STATEMENT_KEYS = ("income_statement", "balance_sheet", "cashflow_statement")


def _now_iso() -> str:
    """本地時間帶時區偏移，例如 "2026-09-03T14:22:10+08:00"。純顯示用。"""
    return datetime.now().astimezone().isoformat(timespec="seconds")


# ── 路徑 ──────────────────────────────────────────────────────────────────
#
# 沿用 `config.py` 的 `%APPDATA%\SEC Financial Tools\`（**有空格**那個；
# `override_engine.py` 用的是底線版 `SEC_Financial_Tools`，兩者歷史上就分岔了，
# 這裡跟 config.py 對齊）。每次呼叫重讀環境變數，測試才好導到 tmp。

def cache_root() -> Path:
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / "SEC Financial Tools" / "filing_cache"
    return Path.home() / ".sec_financial_tools" / "filing_cache"


def ticker_dir(ticker: str) -> Path:
    """一個 ticker 一個資料夾，名稱就是大寫 ticker——不用查表，在檔案總管
    肉眼就看得出哪些公司有快取、大概多大。"""
    return cache_root() / (ticker or "").strip().upper()


def filing_path(ticker: str, accession: str) -> Path | None:
    """`<accession>.json` 的完整路徑。accession 格式不合法回 None（不快取）
    ——這同時擋掉把奇怪字串當檔名寫出去的可能。"""
    if not ACCESSION_RE.match(str(accession or "")):
        return None
    return ticker_dir(ticker) / f"{accession}.json"


# ── 原子寫入 ──────────────────────────────────────────────────────────────

def atomic_write_json(path: Path, obj) -> bool:
    """tmp + `os.replace()`。tmp 檔名帶 PID，避免兩個實例互相蓋到暫存檔。

    寫不進去（磁碟滿、權限）只回 False，不拋——快取只是加速層，
    寫入失敗不該影響這次抓取的結果。
    """
    tmp = Path(str(path) + f".{os.getpid()}.tmp")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False)
        os.replace(tmp, path)
        return True
    except Exception:
        try:
            tmp.unlink()
        except OSError:
            pass
        return False


# ── edgartools 版本 ───────────────────────────────────────────────────────

def edgartools_version() -> str | None:
    """實測 `edgar.__version__` **不存在**（AttributeError），只能走
    package metadata（實測回 "5.29.0"）。取不到回 None，呼叫端把 None 當成
    「這次不要用快取」——不可以填一個預設值混進檔案裡。"""
    try:
        from importlib.metadata import version
        return version("edgartools")
    except Exception:
        return None


# ── DataFrame 序列化 ──────────────────────────────────────────────────────

def df_to_payload(df: pd.DataFrame | None) -> dict | None:
    """DataFrame → 可放進 JSON 的物件。`None` 原樣傳遞（代表這張表不存在）。

    存 `json.loads(...)` 的**物件**不是 `to_json()` 的字串：字串塞進外層
    JSON 會被整份逃逸一次，檔案膨脹 10~15%，而且打開來完全不能看。
    """
    if df is None:
        return None
    return {
        "data": json.loads(df.to_json(orient="split")),
        "dtypes": {str(col): str(dt) for col, dt in df.dtypes.items()},
    }


def payload_to_df(payload: dict | None) -> pd.DataFrame | None:
    """payload → DataFrame。`None` 原樣傳遞。

    `orient="split"` 不帶 dtype，整欄皆 null 的欄位會被推成 float64——所以
    一定要照存檔時記下的 `dtypes` 明確 `astype()` 回去，不能靠自動推斷。
    """
    if payload is None:
        return None
    raw = payload["data"]
    df = pd.DataFrame(raw["data"], index=raw["index"], columns=raw["columns"])
    for col, dtype in (payload.get("dtypes") or {}).items():
        if col not in df.columns:
            continue
        # 只還原「不可能還原錯」的型別：object / bool / int*／uint*／float*。
        # datetime64 是實測踩到的地雷——`to_json(orient="split")` 把 datetime
        # 欄寫成 epoch **毫秒**的整數，`pd.DataFrame(...)` 讀回來自然變成
        # int64；這裡若照樣 `astype("datetime64[us]")`，pandas 會把那串整數
        # **當成微秒**重新解讀（毫秒被錯讀成微秒，時間軸整個縮 1000 倍），
        # 例如 2025-12-27 會變成 1970-01-21 10:46:33.6——`astype` 本身不拋
        # 例外，`except (TypeError, ValueError)` 完全抓不到，數字看起來正常
        # 但整欄都是錯的，比拋例外更危險。分不清楚就一律不寫回去，讓該欄
        # 留在 pandas 從 JSON 自動推斷出來的樣子（值本身沒錯，只是 dtype
        # 標籤不是原本那個）。
        if not (dtype in ("object", "bool") or dtype.startswith(("int", "uint", "float"))):
            continue
        try:
            df[col] = df[col].astype(dtype)
        except (TypeError, ValueError):
            # 型別還原失敗不該讓整份快取報廢——數值本身是對的，
            # 下游只有極少數地方在意 dtype。
            pass
    return df


# ── 替身物件（快取命中時 `_filing_obj()` 的回傳值）──────────────────────────
#
# ⚠ 這三個類別**刻意不定義 `__getattr__`**。`_financials_of()` 是
# `getattr(tenq, "financials", None)`，替身若對未知屬性兜底回 None，以後有人
# 在某個 builder 新用到 filing 物件的其他屬性，快取命中的路徑會安靜地把整份
# filing 當成沒資料、清快取重跑卻是好的——這種 bug 極難查。只實作有人真的
# 在用的那幾條路徑，其餘一律照 Python 預設失敗。

_UNSET = object()


class _CachedStatement:
    """替身的一張報表。只有 `to_dataframe()`，因為呼叫端只用這一個。

    **解析結果 memo 起來**（TODO I7）。四個 builder（IS／BS／CF／segments）
    各自對同一份 filing 的同一張表呼叫一次，memo 之前 `payload_to_df()` 會把
    「JSON → DataFrame → astype」整輪重跑 —— ARLO 預設參數實測 **224 次、
    合計 0.37s**。

    **每次仍回傳 copy，不共用同一個 DataFrame 物件。** 現行程式碼沒有任何一處
    改動報表 dataframe（全庫零 `inplace=True`、零欄位指派），共用「現在」是安全
    的；但哪天有人寫 `df["x"] = ...`，症狀會是「另一張表莫名多一欄」，極難查。
    實測深複製比重新解析便宜 **9.8 倍**（0.17ms vs 1.67ms），隔離幾乎免費。
    """

    def __init__(self, payload: dict):
        self._payload = payload
        self._df = _UNSET

    def to_dataframe(self) -> pd.DataFrame:
        if self._df is _UNSET:
            self._df = payload_to_df(self._payload)
        return None if self._df is None else self._df.copy()


class _CachedFinancials:
    """替身的 `financials`。三個 getter 全部無參數，跟真物件的用法一致。

    getter 回傳的 `_CachedStatement` 也要 memo：每次 new 一個的話，上面那層
    memo 形同虛設——四個 builder 各自呼叫 `fin.income_statement()`，拿到的會是
    四個各自空白的物件。
    """

    def __init__(self, dataframes: dict):
        self._dfs = dataframes or {}
        self._stmts: dict[str, "_CachedStatement | None"] = {}

    def _stmt(self, key: str) -> "_CachedStatement | None":
        if key not in self._stmts:
            payload = self._dfs.get(key)
            self._stmts[key] = None if payload is None else _CachedStatement(payload)
        return self._stmts[key]

    def income_statement(self):
        return self._stmt("income_statement")

    def balance_sheet(self):
        return self._stmt("balance_sheet")

    def cashflow_statement(self):
        return self._stmt("cashflow_statement")


class _CachedFiling:
    """替身的 filing 物件。只有 `.financials` 一個屬性。"""

    def __init__(self, entry: dict):
        self.financials = (_CachedFinancials(entry.get("dataframes") or {})
                           if entry.get("has_financials") else None)


def cached_filing(entry: dict) -> _CachedFiling:
    """快取檔內容 → 可以餵給既有 builder 的替身 filing 物件。"""
    return _CachedFiling(entry)


# ── 單份 filing 的讀寫 ────────────────────────────────────────────────────
#
# `<accession>.json` **是否存在，才是「這份 filing 有沒有快取」的事實來源**。

def load_filing(ticker: str, accession: str, cik: int) -> dict | None:
    """讀一份快取。四道閘任一沒過就回 None（視同無快取，照舊打 SEC 重抓）：
    JSON 可解析、`schema_version`、`cik`、`edgartools_version`。

    正確性優先於速度——寧可那次變慢，也不要餵錯公司的資料或吃到舊版
    parser 的 bug。任何情況都不拋例外。
    """
    version = edgartools_version()
    if version is None:
        return None
    path = filing_path(ticker, accession)
    if path is None or not path.exists():
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            entry = json.load(f)
    except (OSError, ValueError):
        return None
    if not isinstance(entry, dict):
        return None
    if entry.get("schema_version") != SCHEMA_VERSION:
        return None
    if entry.get("edgartools_version") != version:
        return None
    if entry.get("cik") != cik:
        return None
    return entry


def save_filing(ticker: str, accession: str, *, form: str, filing_date: str,
                cik: int, dataframes: dict | None, has_financials: bool) -> bool:
    """寫一份快取。**逐份即時落檔**——一趟抓取可能好幾分鐘，中途斷線或關視窗
    時，已經抓到的進度不該全部白費。

    `has_financials=False` 是負向快取（pre-XBRL 舊申報）。**網路失敗絕對不能
    走到這裡**——那是暫時性的，交給既有的 D11-B 缺漏帳本，每次都該重試。
    """
    version = edgartools_version()
    if version is None:
        return False
    path = filing_path(ticker, accession)
    if path is None:
        return False
    payloads = None
    if has_financials:
        payloads = {k: df_to_payload((dataframes or {}).get(k)) for k in STATEMENT_KEYS}
    entry = {
        "schema_version": SCHEMA_VERSION,
        "accession_no": accession,
        "form": form,
        "filing_date": filing_date,
        "cached_at": _now_iso(),
        "cik": cik,
        "edgartools_version": version,
        "has_financials": bool(has_financials),
        "dataframes": payloads,
    }
    return atomic_write_json(path, entry)


# ── GUI：統計與清除 ───────────────────────────────────────────────────────

def _dir_stats(directory: Path) -> tuple[int, int]:
    """(filing 份數, 位元組數)。`except OSError` 是必要的——清除中或另一個
    實例正在寫入時，檔案可能在 glob 與 stat 之間消失。"""
    count = 0
    size = 0
    try:
        paths = list(directory.glob("*.json"))
    except OSError:
        return 0, 0
    for path in paths:
        # 只算 filing 本身。同一個資料夾裡還有 `_meta.json`（見 local_db.py）
        # ——它既不是 filing，也不該讓一個「清空後只剩 meta」的資料夾在 GUI
        # 上顯示成一列「0 份」（下面 `list_cached_tickers()` 靠 size 來過濾）。
        if not ACCESSION_RE.match(path.stem):
            continue
        try:
            size += path.stat().st_size
        except OSError:
            continue
        count += 1
    return count, size


def list_cached_tickers() -> list[dict]:
    """快取了哪些公司：直接掃 `filing_cache/` 底下有哪些子資料夾。
    公司數量最多幾十家，掃資料夾的成本可以忽略。"""
    root = cache_root()
    rows: list[dict] = []
    try:
        entries = root.iterdir()
    except OSError:
        return []
    # 過濾出目錄，但要容忍單一目錄在 is_dir() 時消失——不能因為一個目錄的
    # stat 失敗就拋棄整份掃描。正在被清除或另一個實例正在寫入時常見。
    entries_list = []
    for p in entries:
        try:
            if p.is_dir():
                entries_list.append(p)
        except OSError:
            continue
    entries_list.sort(key=lambda p: p.name)
    for directory in entries_list:
        count, size = _dir_stats(directory)
        if count == 0 and size == 0:
            continue
        rows.append({"ticker": directory.name, "count": count, "size_bytes": size})
    rows.sort(key=lambda r: (-r["size_bytes"], r["ticker"]))
    return rows


def total_size_bytes() -> int:
    return sum(r["size_bytes"] for r in list_cached_tickers())


def clear_ticker(ticker: str) -> bool:
    """整個刪掉那家公司的資料夾。下次抓這家會當作全新開始。"""
    directory = ticker_dir(ticker)
    if not directory.exists():
        return False
    try:
        shutil.rmtree(directory)
        return True
    except OSError:
        return False


def clear_all() -> int:
    """刪掉所有公司的快取，回傳刪掉幾家。"""
    return sum(1 for row in list_cached_tickers() if clear_ticker(row["ticker"]))
