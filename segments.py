"""
segments.py — 把 `Data_Seg_*` 的寬格式彙總成單一張長格式 `Data_Segments`。

═══════════════════════════════════════════════════════════════════════════════
為什麼要長格式
═══════════════════════════════════════════════════════════════════════════════
現行做法是每個分類軸開一張 sheet，所以**每家公司的 sheet 名稱與數量都不同**：
NVDA 有 BusinessSegment / Geographical 兩張，ARLO 可能是三張別的名字。下游 skill
得先列出所有 sheet、猜哪張是它要的——這是它最難處理的一種結構。

長格式把全部塞進固定名稱、固定欄位的一張表：

    A 欄（Metric — Member）        B 欄（來源 sheet）      FY2025Q1  FY2025Q2
    Revenue — Data Center         Data_Seg_Revenue          22,563    26,272
    Revenue — Gaming              Data_Seg_Revenue           2,544     2,880
    RevenueGeo — United States    Data_Seg_RevenueGeo       13,506    14,514

skill 讀 `Data_Segments` 然後 filter 就好，不必先探索。人要看寬表的話，
`Data_Seg_*` **原樣保留**——兩者資料同源，成本只是多寫一張 sheet。

═══════════════════════════════════════════════════════════════════════════════
為什麼從寬表衍生而不是重抓一次 XBRL
═══════════════════════════════════════════════════════════════════════════════
兩張表同源才不會有「長表和寬表對不起來」這種問題。多開一條抽取路徑等於多一個
會各自漂移的資料來源。
"""

from __future__ import annotations

from typing import Any

from fetcher_gaap import StatementTable
from zh_labels import axis_label

SHEET_NAME  = "Data_Segments"
SEG_PREFIX  = "Data_Seg_"
NAME_JOINER = " — "


def _metric_name(sheet_name: str) -> str:
    """`Data_Seg_RevenueGeo` → `RevenueGeo`。"""
    return sheet_name[len(SEG_PREFIX):] if sheet_name.startswith(SEG_PREFIX) else sheet_name


def build_segments_long(tables: list[StatementTable]) -> StatementTable | None:
    """把所有 `Data_Seg_*` 併成一張長格式表。沒有任何 segment 表時回 None。

    欄位取各軸季度的**聯集**並排序；某軸沒有那一季就留 None。
    值一律依季度標籤對位——不可假設各軸的季度範圍一致，那會讓整列左移，
    每個數字都掛到錯的季而且看不出來。
    """
    seg_tables = [t for t in tables
                  if t.sheet_name.startswith(SEG_PREFIX) and t.concepts]
    if not seg_tables:
        return None

    all_periods = sorted({q for t in seg_tables for q in t.quarter_labels})

    # 申報日：同一季各軸應相同，取第一個非空的
    filing_dates: list[str] = []
    for period in all_periods:
        date = ""
        for t in seg_tables:
            if period in t.quarter_labels:
                candidate = t.filing_dates[t.quarter_labels.index(period)]
                if candidate:
                    date = candidate
                    break
        filing_dates.append(date)

    concepts: list[str] = []
    labels: list[str] = []
    values: list[list[Any]] = []

    for t in seg_tables:
        metric = _metric_name(t.sheet_name)
        index_of = {q: i for i, q in enumerate(t.quarter_labels)}
        axes = list(t.labels or [""] * len(t.concepts))
        for i, (member, row) in enumerate(zip(t.concepts, t.values)):
            concepts.append(f"{metric}{NAME_JOINER}{member}")
            # B 欄放軸的中文分類、C 欄放原始軸名，讓使用者能篩掉非 segment 的維度
            axis = axes[i] if i < len(axes) else ""
            labels.append(axis)
            values.append([
                row[index_of[p]] if p in index_of and index_of[p] < len(row) else None
                for p in all_periods
            ])

    return StatementTable(
        sheet_name=SHEET_NAME,
        quarter_labels=all_periods,
        filing_dates=filing_dates,
        concepts=concepts,
        values=values,
        ticker=seg_tables[0].ticker,
        labels=labels,
    )
