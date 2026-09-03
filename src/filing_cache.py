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
