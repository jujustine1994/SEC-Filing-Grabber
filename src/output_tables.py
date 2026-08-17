"""output_tables.py — 決定最後要寫進 Excel 的 sheet 清單。

從 `main.py` 抽出來（2026-08-07）。原因：`cli.py` 需要跟 GUI 走**完全一樣**的
輸出組裝邏輯，但 `main.py` 一 import 就會拉進 tkinter。複製一份到 CLI 是最糟的
選擇——兩邊會慢慢長歪，而「GUI 產的 Excel 跟 CLI 產的不一樣」這種 bug 極難發現。

行為與抽出前逐字相同，`main.py` 改成 import 這裡。
"""
from __future__ import annotations

from ratios import build_ratio_table
from segments import build_segments_long


def append_ratio_table(tables: list) -> None:
    """把 Data_Segments（長格式）與 Data_Ratios 加進輸出清單（就地修改）。

    來源固定是 Data_Financials(Q)——比率要看季度趨勢，年報只有 4 個點。
    沒抓 GAAP（只勾 Non-GAAP）時沒有來源表，安靜跳過。
    """
    seg_long = build_segments_long(tables)
    if seg_long is not None:
        tables.append(seg_long)

    # 輸出精簡（2026-08-03）：只留有意義的原始資料表。
    # 砍掉的都是「同資料換個排法」或「幾乎全空」的：
    #   Data_Seg_*          寬格式，與 Data_Segments 同源，且每家 sheet 名稱都不同
    #   Data_Financials_NG  XBRL overflow 的 Non-GAAP 分流，內容已在主表 overflow 區
    #   Data_NonGAAP        暫停（等 skill 方案，見 TODO B）
    #   Data_EPS_Recon      edgartools 從未回傳過內容
    #   Data_Std            列位固定已移進三表本身，不需要獨立一張
    tables[:] = [t for t in tables
                 if not t.sheet_name.startswith("Data_Seg_")
                 and not t.sheet_name.startswith("Data_Financials_NG")
                 and t.sheet_name not in ("Data_NonGAAP", "Data_EPS_Recon", "Data_Std")]

    q_tbl = next((t for t in tables if t.sheet_name == "Data_Financials(Q)"), None)
    if q_tbl is None:
        return
    ratio_tbl = build_ratio_table(q_tbl)
    if ratio_tbl is not None:
        tables.append(ratio_tbl)


# 這幾張表就算三表全空也會有內容或欄位結構，不能拿來當「有抓到資料」的證據：
# Data_Meta 永遠有 ticker 與抓取日期；Data_Ratios 是從三表算出來的衍生表。
_NOT_EVIDENCE_OF_DATA = frozenset({"Data_Meta", "Data_Ratios", "Index"})


def has_any_data(tables) -> bool:
    """這一趟到底有沒有抓到東西？

    缺幾期留空可以接受（使用者會看到警告），但**全部**都沒抓到時寫出去
    就是一份空殼 Excel，而它會蓋掉使用者原本好好的舊檔——那是這整件事裡
    唯一真正不可逆的傷害。

    不能只檢查 `if not tables`：一期都沒抓到時 _merge_financials 仍會產出
    空的 Data_Financials(Q) 等結構，list 不是空的。
    """
    for tbl in tables:
        if tbl.sheet_name in _NOT_EVIDENCE_OF_DATA:
            continue
        if not tbl.quarter_labels:
            continue
        # 欄位標籤在但每一格都是 None——版面在、數字沒有，一樣是空殼
        if any(v is not None for row in tbl.values for v in row):
            return True
    return False
