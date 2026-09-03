"""
filing_cache.py — 本地 filing 解析快取（%APPDATA%\\SEC Financial Tools\\filing_cache）。

快取卡在**解析層與比對層之間**：存的是 edgartools 解出來的三張 DataFrame
（income statement / balance sheet / cashflow statement），比對層
（`IS/BS/CF_TEMPLATE` 那套科目對照）永遠在快取之上即時重跑。所以以後改
hint regex、加比率、調 Q4 合成邏輯都不會讓快取失效——但 **edgartools 升版
會**，那是另一條軸線，靠 `edgartools_version` 欄位擋（見 `load_filing`）。

事實來源是 `<accession>.json` 檔案本身；`_manifest.json` 只是給 GUI 看的
衍生索引，壞了直接從資料夾重建。
"""
from __future__ import annotations

import json
import os
import re
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
    except OSError:
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

class _CachedStatement:
    """替身的一張報表。只有 `to_dataframe()`，因為呼叫端只用這一個。"""

    def __init__(self, payload: dict):
        self._payload = payload

    def to_dataframe(self) -> pd.DataFrame:
        return payload_to_df(self._payload)


class _CachedFinancials:
    """替身的 `financials`。三個 getter 全部無參數，跟真物件的用法一致。"""

    def __init__(self, dataframes: dict):
        self._dfs = dataframes or {}

    def _stmt(self, key: str) -> _CachedStatement | None:
        payload = self._dfs.get(key)
        return None if payload is None else _CachedStatement(payload)

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
# manifest 只是給 GUI 看的衍生索引，查快取不問它。

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
    if cik is not None and entry.get("cik") != cik:
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


# ── manifest ──────────────────────────────────────────────────────────────
#
# ⚠ manifest **不是事實來源**，是衍生索引，只給 GUI 顯示用。查快取一律直接問
# `<accession>.json` 存不存在（掃 100 個檔名是微秒級成本，比維護「manifest
# 跟磁碟是否同步」的心智負擔低得多）。所以這裡沒有「修正」邏輯，只有重建。

MANIFEST_NAME = "_manifest.json"


def read_manifest(ticker: str) -> dict | None:
    """讀 manifest。不存在、壞掉、版本不符一律回 None（呼叫端重建）。"""
    path = ticker_dir(ticker) / MANIFEST_NAME
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return None
    if not isinstance(data, dict) or data.get("schema_version") != SCHEMA_VERSION:
        return None
    return data


def rebuild_manifest(ticker: str, cik: int | None = None) -> dict:
    """從資料夾實際內容重建並寫回。manifest 遺失／損毀／跟磁碟對不上都走這條。"""
    directory = ticker_dir(ticker)
    rows: list[dict] = []
    try:
        paths = sorted(directory.glob("*.json"))
    except OSError:
        paths = []
    for path in paths:
        if path.name == MANIFEST_NAME or not ACCESSION_RE.match(path.stem):
            continue
        try:
            size = path.stat().st_size
            with open(path, "r", encoding="utf-8") as f:
                entry = json.load(f)
        except (OSError, ValueError):
            continue    # 壞掉的那一份不進索引，下次抓取會重抓並覆蓋
        if not isinstance(entry, dict):
            continue
        rows.append({
            "accession_no": entry.get("accession_no", path.stem),
            "form": entry.get("form", ""),
            "filing_date": entry.get("filing_date", ""),
            "cached_at": entry.get("cached_at", ""),
            "edgartools_version": entry.get("edgartools_version", ""),
            "has_financials": bool(entry.get("has_financials")),
            "size_bytes": size,
        })
    manifest = {
        "schema_version": SCHEMA_VERSION,
        "ticker": (ticker or "").strip().upper(),
        "cik": cik,
        # 上次去 SEC 查 filing 清單的時間。純顯示用，**不是有效期限**
        # ——每次抓取都會重查，不靠這個判斷要不要查。
        "last_checked_at": _now_iso(),
        "filings": rows,
    }
    if rows or directory.exists():
        atomic_write_json(directory / MANIFEST_NAME, manifest)
    return manifest
