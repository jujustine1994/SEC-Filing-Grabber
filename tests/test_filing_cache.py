"""Tests for filing_cache.py — 本地 filing 快取的儲存層。

快取的是「解析層」的輸出（edgartools 解出來的三張 DataFrame），比對層
永遠在快取之上即時重跑。所以這裡釘的是「存進去再讀回來，跟原本一模一樣」，
以及「任何一種對不上的情況都要安靜地退回無快取，不能拋例外、不能餵錯資料」。
"""
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
