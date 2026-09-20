"""
filing_cache.py — 本地 filing 解析快取（`<專案根目錄>/local_db/filing_cache/`）。

快取卡在**解析層與比對層之間**：存的是 edgartools 解出來的三張 DataFrame
（income statement / balance sheet / cashflow statement），比對層
（`IS/BS/CF_TEMPLATE` 那套科目對照）永遠在快取之上即時重跑。所以以後改
hint regex、加比率、調 Q4 合成邏輯都不會讓快取失效——但 **edgartools 升版
會**，那是另一條軸線，靠 `edgartools_version` 欄位擋（見 `load_filing`）。

⚠ **加欄位前先問：能不能從舊檔既有資訊推導出來？能，就不要動
`SCHEMA_VERSION`。** 升版是核彈級動作——全庫作廢、重抓 11 小時。
`fetched_keys`（2026-09-20）是第一個照這條規矩做的案例：既有 14,417 份都
沒有這個欄位，但它們全是 `schema_version=2`，而 v2 的定義就是「六張表全抓」，
所以 `load_filing()` 讀到缺欄位時直接推導補上，檔案一個位元組都不用改。

事實來源是 `<accession>.json` 檔案本身，也是唯一的落地狀態——「哪些公司有
快取」直接掃 `filing_cache/` 底下有哪些子資料夾回答（見 `list_cached_tickers()`），
不維護額外的索引檔。

**2026-09-18 從 `%APPDATA%` 搬到專案資料夾固定路徑**：這份資料是花真金白銀
（SEC 網路請求時間）抓下來的，語意上是「永久資料庫」，不是「可隨時重建的
快取」——但它原本躺在 `%APPDATA%`，那個位置在 Windows 語意上就是「系統可以
清掉的東西」（重灌、系統清理工具、防毒軟體都可能動它）。201 家、13,921 份
的資料就是這樣憑空消失的（起因是 GUI 的「全部清除」按鈕被按過，但放在
`%APPDATA%` 這件事本身也是風險）。搬進專案資料夾底下的 `local_db/` 後，跟
其他你會留意、會備份的專案檔案放在一起，`.gitignore` 排除掉不進版控。
"""
from __future__ import annotations

import json
import os
import re
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd

# 2（2026-09-18）：`STATEMENT_KEYS` 從三張表擴成六張（加 statement_of_equity、
# comprehensive_income、cover）。**必須升版**——舊快取沒有那三個 key，
# `payload_to_df(None)` 回 None 的語意是「這張表本來就不存在」，跟「當初根本
# 沒抓」是兩件事。不升版的話舊檔會被當成「這家公司沒有股東權益變動表」，
# 而且完全不報錯，正是這個專案最不能接受的失效模式。
SCHEMA_VERSION = 2

# SEC 的 accession number 格式固定，拿來當檔名前先驗——這同時是路徑注入的防線。
ACCESSION_RE = re.compile(r"^\d{10}-\d{2}-\d{6}$")

# edgartools 的 `Financials` 能給的六張表**全部都存**（2026-09-18 CTH 決定：
# 「反正都要抓了，看有沒有辦法全抓」）。理由是快取卡在解析層與比對層之間——
# 存下來的東西以後改模板、加新報表都不必重抓 SEC，而抓一次要 11 小時。
#
# 前三張是既有模板在用的；後三張目前**沒有任何下游在讀**，純粹是先存著：
#   - statement_of_equity   股東權益變動表（實測 AAPL 25 列、ARLO 40 列）
#   - comprehensive_income  綜合損益表（AAPL 15 列、ARLO 52 列）
#   - cover                 封面頁（在外流通股數、entity 資訊、財年結束日等）
STATEMENT_KEYS = ("income_statement", "balance_sheet", "cashflow_statement",
                  "statement_of_equity", "comprehensive_income", "cover")

# 模板比對層真正會讀的那幾張。`_CachedFinancials` 的替身只保證這幾個方法
# 的行為跟真的 edgartools 物件一致。
CORE_STATEMENT_KEYS = ("income_statement", "balance_sheet", "cashflow_statement")


def _now_iso() -> str:
    """本地時間帶時區偏移，例如 "2026-09-03T14:22:10+08:00"。純顯示用。"""
    return datetime.now().astimezone().isoformat(timespec="seconds")


# ── 路徑 ──────────────────────────────────────────────────────────────────
#
# 固定放在專案資料夾底下的 `local_db/filing_cache/`——刻意不跟 `config.py`
# 一樣走 `%APPDATA%`：這份資料是永久資料庫，不是可隨時重建的快取，不該放在
# 語意上「系統可以清掉」的地方。`SEC_LOCAL_DB_ROOT` 環境變數可覆寫（只給
# 測試用，導去 tmp_path），每次呼叫重讀環境變數。

def _project_root() -> Path:
    return Path(__file__).resolve().parent.parent


def cache_root() -> Path:
    override = os.environ.get("SEC_LOCAL_DB_ROOT")
    if override:
        return Path(override) / "filing_cache"
    return _project_root() / "local_db" / "filing_cache"


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

    # 2026-09-18 起也存這三張。目前沒有下游在讀，但替身要跟真物件一致——
    # 少了的話「快取命中時拿不到、清快取重跑卻拿得到」，正是最難查的那種
    # 不一致（見下面 `_CachedFiling` 刻意不定義 `__getattr__` 的理由）。
    def statement_of_equity(self):
        return self._stmt("statement_of_equity")

    def comprehensive_income(self):
        return self._stmt("comprehensive_income")

    def cover(self):
        return self._stmt("cover")


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

def load_filing(ticker: str, accession: str, cik: int,
                require_keys=CORE_STATEMENT_KEYS) -> dict | None:
    """讀一份快取。五道閘任一沒過就回 None（視同無快取，照舊打 SEC 重抓）：
    JSON 可解析、`schema_version`、`cik`、`edgartools_version`、`require_keys`。

    正確性優先於速度——寧可那次變慢，也不要餵錯公司的資料或吃到舊版
    parser 的 bug。任何情況都不拋例外。

    `require_keys` 是「這次要用到哪幾張表」。預設只要核心三張——日常產
    Excel 的路徑不該因為某張沒人在讀的 extras 沒抓就整份重抓。要用到
    extras 的新功能自己傳（例如 `require_keys=("statement_of_equity",)`），
    只有當初真的沒抓的那些會回 None 去重抓。

    回傳的 entry 保證有 `fetched_keys`（舊檔會就地推導補上），下游一律讀
    它就好，不必每個呼叫點各自再判一次舊檔。
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
    # 第五道閘。`fetched_keys` 是 2026-09-20 加的欄位，既有的 14,417 份都沒有
    # ——但它們全是 `schema_version=2`，而 v2 的定義就是「六張表全抓」
    # （2026-09-18 那次改版），所以推導得出來，**不必升版讓全庫作廢重抓**。
    # 這次加欄位本身就是這條規矩的第一個案例：能從舊檔既有資訊推導出來的
    # 欄位，就不要動 `SCHEMA_VERSION`。
    fetched = entry.get("fetched_keys")
    if not isinstance(fetched, list):
        fetched = list(STATEMENT_KEYS)
    entry["fetched_keys"] = fetched
    # ⚠ 負向快取（pre-XBRL，`has_financials=False`）不受這道閘管。它根本沒有
    # `dataframes`，「哪張表沒抓」對它沒有意義；拿 `require_keys` 去擋它會讓
    # 那些舊申報每趟都重打一次 SEC，永遠不收斂。
    if entry.get("has_financials") and not set(require_keys or ()) <= set(fetched):
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
        # 存檔當下的 `STATEMENT_KEYS`。有了它，`dataframes` 裡的 `null` 才沒有
        # 歧義：key 在這份清單裡 → 這家公司真的沒有這張表；不在 → 當初沒抓。
        # 分不出這兩件事的代價 2026-09-18 付過一次（三張擴六張，舊檔的
        # `statement_of_equity` 被讀成「這家公司沒有股東權益變動表」，不報錯、
        # Excel 默默少一張表，只能整庫重抓 11 小時）。
        # 負向快取也照寫——那時這個欄位不帶意義（沒有 `dataframes`），但一律
        # 存在能讓下游少一種分支，`load_filing()` 的第五道閘本來就會跳過它。
        "fetched_keys": list(STATEMENT_KEYS),
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


def list_cached_filings(ticker: str, form: str | None = None) -> list[dict]:
    """盤點本地有這家公司的哪些 filing → `[{accession, form, filing_date}, ...]`，
    **新到舊**（跟 SEC 的清單順序一致）。

    ⚠ 這是 TODO J9「離線退路」專用的。平常的抓取流程**不該**呼叫它——
    正常情況一定要問 SEC「有哪些財報」，只看本地會永遠發現不了新申報
    （見 `docs/TODO.md` J9 的「清單不可以本地優先」）。

    要開每一份檔案才讀得到 `form`／`filing_date`，所以不便宜（實測約
    2.75 秒/家）。這是例外路徑，可以接受。
    """
    version = edgartools_version()
    if version is None:
        return []                         # 讀不到版本時 `load_filing()` 也全關
    rows: list[dict] = []
    directory = ticker_dir(ticker)
    try:
        paths = list(directory.glob("*.json"))
    except OSError:
        return []
    for path in paths:
        if not ACCESSION_RE.match(path.stem):
            continue                      # `_meta.json` 之類的不是 filing
        try:
            with open(path, "r", encoding="utf-8") as f:
                entry = json.load(f)
        except (OSError, ValueError):
            continue
        if not isinstance(entry, dict):
            continue
        # ⚠ 要走**跟 `load_filing()` 同一組閘**，不能只擋 schema。
        # 只擋 schema 的話，edgartools 升版後這裡照樣回報「有 25 份」，但每一
        # 份 `load_filing()` 都回 None——離線退路會宣告「已用 25 份本地資料」，
        # 實際上一份都讀不出來，最後產出一份空 Excel 被標成「離線資料」而不是
        # 「讀不出來」。那樣 `if not rows: raise` 那道保護形同虛設
        # （2026-09-18 code review 實測抓到）。
        if entry.get("schema_version") != SCHEMA_VERSION:
            continue
        if entry.get("edgartools_version") != version:
            continue
        row_form = str(entry.get("form") or "")
        if form is not None and row_form != form:
            continue
        rows.append({
            "accession": path.stem,
            "form": row_form,
            "filing_date": str(entry.get("filing_date") or ""),
        })
    rows.sort(key=lambda r: r["filing_date"], reverse=True)
    return rows


def cached_cik(ticker: str) -> int | None:
    """從本地任一份快取檔撈這家公司的 cik（TODO J9）。

    **離線時唯一還拿得到 cik 的地方**。`Company(ticker).cik` 要連網，而
    `_bind_disk_cache()` 沒有 cik 就整段關閉磁碟快取——那樣離線退路盤出來的
    filing 全部會走到 `.obj()` 然後拋例外，變成「宣告用了本地資料、實際拿到
    空表」。cik 是上次抓成功時寫進每一份快取檔的，直接讀回來就好。
    """
    for row in list_cached_filings(ticker)[:1]:
        path = ticker_dir(ticker) / f"{row['accession']}.json"
        try:
            with open(path, "r", encoding="utf-8") as f:
                cik = json.load(f).get("cik")
        except (OSError, ValueError, AttributeError):
            return None
        try:
            return int(cik)
        except (TypeError, ValueError):
            return None
    return None


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


# ── 6-K 的 R 檔判定快取（TODO D9 A 路線）─────────────────────────────────

SIXK_PROBE_FILENAME = "_sixk_probe.json"


def sixk_probe_path(ticker: str) -> Path:
    return ticker_dir(ticker) / SIXK_PROBE_FILENAME


def load_sixk_probe(ticker: str) -> dict[str, int]:
    """`{accession: R*.htm 檔數}`。問一份 6-K 有幾個 R 檔要一次 index 請求，
    存下來之後就永遠不必再問（一份申報的附件不會變）。

    ⚠ **獨立檔，不併進 `_meta.json`**：判定不是財報內容，混進去會讓
    `file_count` 這條「目錄是事實來源」的規則失焦。檔名不合 `ACCESSION_RE`，
    所以 `_dir_stats()`／`list_cached_filings()` 本來就不會把它當 filing。

    認不出來的一律丟掉回 `{}`／略過那一筆——重問的成本是請求，採信髒資料的
    成本是那幾季的財報被永久跳過（而且完全沒有症狀）。
    """
    try:
        with open(sixk_probe_path(ticker), encoding="utf-8") as f:
            raw = json.load(f)
    except (OSError, ValueError):
        return {}
    if not isinstance(raw, dict):
        return {}
    return {k: v for k, v in raw.items()
            if ACCESSION_RE.match(str(k)) and isinstance(v, int)
            and not isinstance(v, bool)}


def save_sixk_probe(ticker: str, probes: dict) -> bool:
    return atomic_write_json(sixk_probe_path(ticker), dict(probes or {}))
