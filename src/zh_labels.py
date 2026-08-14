"""
zh_labels.py — 三表科目與維度軸的顯示說明（Excel B 欄）。

⚠ **譯文已於 2026-08-14 遷入 `locales/*.py`**，這個檔案只剩兩個薄 wrapper。
要改用詞請改 `src/locales/<語言>.py` 裡的 `acct.*` / `axis.*` 條目。
檔名與函式簽名保留是刻意的——`excel_writer.py` 與 `segments.py` 直接
import 這兩個函式，改名等於為了搬字串去動不相干的檔案。

Excel 版面：
    A 欄  英文標準名（程式內部用這個對映，**不要改**）
    B 欄  說明（本檔查表，跟著介面語言走）
    C 欄  公司原始財報標籤（`Net sales`，永遠英文原文，拿去 10-Q Ctrl+F 用）
    D 欄起 各期數據

B 欄只是給人看的，程式一律用 A 欄的英文名做比對。所以譯文隨便你改用詞，
改錯也不會弄壞任何計算——最壞的情況只是 B 欄顯示怪怪的。

沒收錄的科目（例如 overflow 區的公司特有科目）B 欄留白。
"""

from __future__ import annotations

from i18n import t


def _lookup(ns: str, key: str) -> str:
    """查 `<ns>.<key>`，查不到回空字串（而不是 t() 預設回傳的 key 本身）。

    B 欄查不到就該留白——overflow 科目、未收錄的比率都是常態，把
    `ratio.Foo (%)` 這種 key 印進儲存格只是雜訊。
    """
    key = (key or "").strip()
    if not key:
        return ""
    full = f"{ns}.{key}"
    val = t(full)
    return "" if val == full else val


def zh_label(concept: str) -> str:
    """取科目說明。沒收錄回空字串——overflow 區的公司特有科目本來就沒有。

    函式名保留 `zh_`：改名要動 excel_writer 與既有測試，而它現在回傳的是
    「當前語言」的說明，不再限於中文。名字是歷史包袱，行為是對的。
    """
    # 查不到留白：overflow 區的公司特有科目沒收錄是常態不是錯誤
    return _lookup("acct", concept)


def ratio_label(name: str) -> str:
    """Data_Ratios 列名 → 說明（B 欄）。A 欄是英文機器鍵，這裡回當前語言。"""
    return _lookup("ratio", name)


def meta_label(name: str) -> str:
    """Data_Meta 欄位名 → 說明（B 欄）。同上。"""
    return _lookup("meta", name)


def axis_label(axis: str) -> str:
    """維度軸 → 分類說明（Data_Segments 的 B 欄）。

    XBRL 的分類細項掛在不同的「軸」上。沒有軸就分不出這一列是業務別營收，
    還是權益變動表的項目別——MSFT 實測會混進 `Retained earnings`（權益項目
    軸）與 `Service Life`（耐用年限軸），那些根本不是 segment。

    **不過濾、不丟棄**，只把軸標出來讓你自己篩。沒收錄的軸標成「其他維度」，
    不回空字串——空白會讓人以為沒有軸，實際上是有軸但我們沒收錄。
    """
    if not (axis or "").strip():
        return ""
    return _lookup("axis", axis) or t("axis._other")
