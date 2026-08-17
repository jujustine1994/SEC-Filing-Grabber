"""一期都沒抓到時不可以寫檔（CTH 2026-08-17）。

缺幾期留空是可以接受的——使用者會看到警告。但**全部**都沒抓到時寫出去
就是一份空殼 Excel，而它會蓋掉使用者原本好好的舊檔。那是這整件事裡唯一
真正不可逆的傷害。

`tables` 本身不會是空 list：即使一期都沒抓到，_merge_financials 仍會產出
空的 Data_Financials(Q)、Data_Meta 等結構。所以不能只檢查 `if not tables`。
"""

import pytest

from fetcher_gaap import StatementTable
from main import has_any_data


def _tbl(name, quarters=("FY2024Q1",), values=None):
    n = len(quarters)
    return StatementTable(
        sheet_name=name,
        quarter_labels=list(quarters),
        filing_dates=["2024-01-01"] * n,
        concepts=["Revenue"],
        values=values if values is not None else [[1.0] * n],
        ticker="TEST",
        labels=[""],
    )


def _empty(name):
    return StatementTable(
        sheet_name=name, quarter_labels=[], filing_dates=[],
        concepts=[], values=[], ticker="TEST", labels=[],
    )


def test_a_normal_result_counts_as_data():
    assert has_any_data([_tbl("Data_Financials(Q)")])


def test_an_empty_list_is_not_data():
    assert not has_any_data([])


def test_structure_without_any_period_is_not_data():
    """一期都沒抓到時仍會產出空的表結構——那不算有資料，寫出去就是空殼。"""
    assert not has_any_data([
        _empty("Data_Financials(Q)"),
        _empty("Data_Financials(Y)"),
    ])


def test_meta_alone_is_not_data():
    """Data_Meta 永遠有值（ticker、抓取日期），它有東西不代表抓到財報。
    只看「有沒有任何非空的表」會被它騙過去。"""
    assert not has_any_data([
        _empty("Data_Financials(Q)"),
        _tbl("Data_Meta", quarters=("FY2024Q1",)),
    ])


def test_ratios_alone_is_not_data():
    """Data_Ratios 是從三表算出來的。三表全空時它也該是空的，
    但它的欄位結構仍在——同樣不可以當成「有抓到」。"""
    assert not has_any_data([
        _empty("Data_Financials(Q)"),
        _tbl("Data_Ratios", quarters=("FY2024Q1",), values=[[None]]),
    ])


def test_annual_only_still_counts():
    """只勾年報時季報表是空的，那是使用者自己選的，不是失敗。"""
    assert has_any_data([
        _empty("Data_Financials(Q)"),
        _tbl("Data_Financials(Y)"),
    ])


def test_all_values_none_is_not_data():
    """有欄位標籤但每一格都是 None——版面在、數字沒有，一樣是空殼。"""
    assert not has_any_data([
        _tbl("Data_Financials(Q)", quarters=("FY2024Q1", "FY2024Q2"),
             values=[[None, None]]),
    ])
