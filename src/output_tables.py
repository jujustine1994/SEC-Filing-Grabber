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
