"""press_release_tables.py — 8-K 新聞稿表格的確定性解析（TODO B3）。

**零 AI**：只用 `pandas.read_html` + 版面規則，不呼叫任何 API。

要解決的問題：多數美股新聞稿由 Workiva 產生，同一份 HTML 表格裡塞滿版面用的
垃圾欄——同一個數字重複寫進相鄰兩欄、`$` 與 `%` 各自佔一欄、期間之間插全空的
間隔欄。`pd.read_html` 忠實還原這些欄位，ARLO 的調節表因此是 24×30 的網格，
真正的資料只有 14×6。

清理流程（`clean_grid`）：

    1. 全空列砍掉
    2. 前面連續、完全沒有數字的欄 = 標籤欄，合併成第 0 欄
    3. 剩下的欄以「全空欄」為界切成群組（全空欄是 Workiva 的期間間隔欄，
       **不能先刪再合併**，否則兩個期間的數字會黏在一起）
    4. 每個群組每列收斂成一格：丟掉 `$`、把 `%` 併回數字、重複值去重

第 3 步是整件事的關鍵。第一版先刪空欄再收斂相鄰重複欄，ARLO 的
「Three Months Ended」與「Six Months Ended」就會合併成同一欄。

設計取捨（2026-08-07，CTH 睡覺時決定，可推翻）：
  - **表格標題不從前後文抓**。`pd.read_html` 的順序與 `//table` 的文件順序在
    ARLO 上一致，但巢狀表格會讓兩者錯開，抓錯標題比沒有標題更糟。改用表格
    自己開頭的純文字列當 `caption`。
  - 同一格收斂出兩個不同數字時**併排保留**（`"1 2"`）而不是取其一。資料異常
    要看得見，不能靜默丟。
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from typing import Any

import pandas as pd

# `$` 是純粹的貨幣記號欄，收斂時丟掉；`%` 要併回數字（48.2 → 48.2%）。
_CURRENCY_TOKENS = {"$", "US$", "NT$", "€", "£", "¥"}
_PERCENT_TOKEN = "%"

# 破折號是「本期無此項」的排版寫法，一律正規化成 em dash 方便下游辨識。
_DASHES = {"-", "–", "—", "‒", "―"}

_NUMBER_RE = re.compile(r"^[(\-+]?\s*[\d,]+(?:\.\d+)?\s*\)?$")

# 判斷一張表要不要留給 skill 看：調節表一定會出現這些字。
_NONGAAP_PATTERNS = (
    re.compile(r"non-?gaap", re.IGNORECASE),
    re.compile(r"reconcil", re.IGNORECASE),
)


@dataclass
class PressTable:
    """清理後的新聞稿表格。第 0 欄是標籤，其餘是期間欄。"""

    index: int
    rows: list[list[str]] = field(default_factory=list)
    caption: str = ""

    @property
    def n_rows(self) -> int:
        return len(self.rows)

    @property
    def n_cols(self) -> int:
        return len(self.rows[0]) if self.rows else 0

    def text(self) -> str:
        """管線分隔的純文字。供關鍵字篩選與 CLI 的非 JSON 輸出使用。"""
        return "\n".join(" | ".join(cell for cell in row) for row in self.rows)

    def to_dict(self) -> dict[str, Any]:
        return {
            "index": self.index,
            "caption": self.caption,
            "n_rows": self.n_rows,
            "n_cols": self.n_cols,
            "rows": self.rows,
        }


# ── 儲存格層級 ──────────────────────────────────────────────────────────────

def _norm(cell: Any) -> str:
    """統一成字串：NaN → ""、壓掉連續空白、破折號正規化。"""
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return ""
    text = " ".join(str(cell).split())
    if text.lower() in ("nan", "none"):
        return ""
    if text in _DASHES:
        return "—"
    return text


def _is_number(token: str) -> bool:
    return bool(_NUMBER_RE.match(token.replace(" ", " ").strip()))


def _has_number(column: list[str]) -> bool:
    return any(_is_number(c) for c in column if c)


def _dedupe(tokens: list[str]) -> list[str]:
    """去掉重複值但保留順序——Workiva 的重複欄與真正的兩個數字要分得開。"""
    out: list[str] = []
    for t in tokens:
        if t not in out:
            out.append(t)
    return out


def _collapse(cells: list[str]) -> str:
    """一個欄位群組在某一列上的所有儲存格 → 一格。"""
    tokens = [c for c in cells if c]
    pct = _PERCENT_TOKEN in tokens
    tokens = [t for t in tokens if t not in _CURRENCY_TOKENS and t != _PERCENT_TOKEN]
    tokens = _dedupe(tokens)
    if not tokens:
        return _PERCENT_TOKEN if pct else ""
    value = " ".join(tokens)
    return f"{value}{_PERCENT_TOKEN}" if pct else value


# ── 網格層級 ────────────────────────────────────────────────────────────────

def _column(grid: list[list[str]], idx: int) -> list[str]:
    return [row[idx] for row in grid]


def _label_span(grid: list[list[str]]) -> int:
    """前面有幾欄是標籤欄（完全不含數字的連續前導欄）。

    至少回傳 1：整張表都是數字時，第 0 欄仍當標籤處理，這樣 `rows[i][0]`
    的意義在所有表上一致，下游不必分兩種情況。
    """
    n_cols = len(grid[0])
    span = 0
    while span < n_cols and not _has_number(_column(grid, span)):
        span += 1
    return max(1, min(span, n_cols))


def _data_rows(grid: list[list[str]]) -> list[list[str]]:
    """有數字的列。表頭列不算——它們會把間隔欄填滿，害間隔判斷失效。"""
    return [row for row in grid if any(_is_number(c) for c in row)]


def _value_groups(grid: list[list[str]], start: int) -> list[list[int]]:
    """把 `start` 之後的欄以間隔欄為界切成群組。

    間隔欄＝**在所有資料列上都是空的**欄。不能用「整欄都空」判斷：Workiva 的
    表頭是 colspan 展開的，`Three Months Ended` 會把 15 個欄位（含中間的間隔欄）
    全部填滿，用整欄判斷就一個間隔都找不到，三個期間的數字會併成一格。
    """
    rows = _data_rows(grid) or grid
    groups: list[list[int]] = []
    current: list[int] = []
    for idx in range(start, len(grid[0])):
        if any(_column(rows, idx)):
            current.append(idx)
        elif current:
            groups.append(current)
            current = []
    if current:
        groups.append(current)
    return groups


def clean_grid(grid: list[list[str]]) -> list[list[str]]:
    """把 read_html 出來的原始網格清成「標籤 + 每期一欄」。"""
    rows = [[_norm(c) for c in row] for row in grid]
    width = max((len(r) for r in rows), default=0)
    rows = [r + [""] * (width - len(r)) for r in rows]
    rows = [r for r in rows if any(r)]
    if not rows:
        return []

    span = _label_span(rows)
    groups = _value_groups(rows, span)

    out: list[list[str]] = []
    for row in rows:
        label = _collapse(row[:span])
        values = [_collapse([row[i] for i in group]) for group in groups]
        out.append([label] + values)

    # colspan 展開的標題列（`NVIDIA CORPORATION` 重複 4 次）只留第一格。
    # 限定「該列沒有數字」才做，否則各期數字剛好相同的列會被清空。
    for row in out:
        filled = [c for c in row if c]
        if len(filled) > 1 and len(set(filled)) == 1 and not _is_number(filled[0]):
            for i in range(1, len(row)):
                row[i] = ""

    # 收斂後才知道哪些欄整欄空（例如整群都是 `$`），這裡再砍一次
    if out:
        keep = [0] + [i for i in range(1, len(out[0])) if any(r[i] for r in out)]
        out = [[r[i] for i in keep] for r in out]
    return [r for r in out if any(r)]


# ── 對外 API ────────────────────────────────────────────────────────────────

def _own_caption(rows: list[list[str]]) -> str:
    """表格開頭那幾列只有標籤欄有字的，當成標題（如「(In thousands)」）。"""
    parts: list[str] = []
    for row in rows:
        if row[0] and not any(row[1:]):
            parts.append(row[0])
        elif any(row[1:]):
            break
    return " / ".join(parts[:2])


def _is_title_block(rows: list[list[str]]) -> bool:
    """整張表只有一欄文字、沒有任何數字——這是新聞稿的標題區塊不是資料表。

    Workiva 把「ARLO TECHNOLOGIES, INC. / RECONCILIATIONS OF GAAP MEASURES TO
    NON-GAAP MEASURES」排成獨立的 `<table>`，緊接在真正的資料表前面。單獨輸出
    是雜訊，但併成下一張表的標題就正好是判斷「這張是不是調節表」的依據。
    """
    if not rows or len(rows[0]) != 1 or len(rows) > 6:
        return False
    # 用「有沒有數字字元」而不是「能不能 parse 成數字」判斷。ARLO 的財測區塊
    # 也是單欄（`GAAP $140 - $150 $(0.06) - $0.00`），parse 不成數字，用寬鬆
    # 條件會被當標題吃掉——那是真的資料，不能丟。
    return not any(any(ch.isdigit() for ch in c) for row in rows for c in row)


def parse_tables(html: str) -> list[PressTable]:
    """解析新聞稿 HTML，回傳清理後的表格（依文件順序）。

    HTML 裡沒有表格時回傳空清單，不拋例外——很多 8-K 的附件是純文字。
    """
    try:
        frames = pd.read_html(io.StringIO(html))
    except ValueError:
        return []

    tables: list[PressTable] = []
    pending_title = ""
    for i, frame in enumerate(frames):
        raw = [list(row) for row in frame.itertuples(index=False)]
        rows = clean_grid(raw)
        if not rows:
            continue
        if _is_title_block(rows):
            block = " / ".join(r[0] for r in rows if r[0])
            # 連續兩個標題區塊要串起來，直接覆蓋等於把前一個丟掉
            pending_title = f"{pending_title} / {block}" if pending_title else block
            continue
        caption = " / ".join(p for p in (pending_title, _own_caption(rows)) if p)
        pending_title = ""
        tables.append(PressTable(index=i, rows=rows, caption=caption))

    if pending_title:
        # 落單的標題區塊（後面沒有資料表）還是輸出，寧可多一張也不要靜默丟資料
        tables.append(PressTable(index=len(frames), rows=[[pending_title]],
                                 caption=pending_title))
    return tables


def is_nongaap_table(table: PressTable) -> bool:
    """這張表跟 Non-GAAP 調節有關嗎？

    `non-gaap` 出現在任何地方都算；`reconcil` **只認標題**。理由：現金流量表
    裡固定有一列「Reconciliation of cash, cash equivalents and restricted
    cash」，認內文的話 ARLO 每季會多帶 2,330 字元的 GAAP 現金流量表進來，
    而那份資料 `Data_Q` 已經有了。
    """
    if _NONGAAP_PATTERNS[0].search(table.text()):
        return True
    return bool(_NONGAAP_PATTERNS[1].search(table.caption))


def filter_nongaap(tables: list[PressTable]) -> list[PressTable]:
    """只留與 Non-GAAP／調節有關的表。

    ARLO 實測：15 張表 → 3 張，字元數 320K（原文）→ 約 2K。
    """
    return [t for t in tables if is_nongaap_table(t)]
