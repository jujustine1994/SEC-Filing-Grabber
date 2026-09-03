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
from pathlib import Path

import pandas as pd

SCHEMA_VERSION = 1

# SEC 的 accession number 格式固定，拿來當檔名前先驗——這同時是路徑注入的防線。
ACCESSION_RE = re.compile(r"^\d{10}-\d{2}-\d{6}$")

STATEMENT_KEYS = ("income_statement", "balance_sheet", "cashflow_statement")


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
