"""data_quality.py — 抓取結果的缺漏判斷。

## 為什麼要重寫（TODO G5）

原本的完成度只看 9 個關鍵列，而且**只在「最近 4 期全部是 None」才算缺**。
判準寬到形同虛設：NVDA 那份檔案 Index 顯示 `9/9 ✓`，實際上 95 個欄位裡有
27 個幾乎全空、3 季完全沒有資料。

## CTH 2026-08-22 定案的三個判斷

可信度由高到低，**三個都不需要「同業基準表」**：

    A  季度斷層        相鄰兩期差太多 → 中間漏了幾季        誤判率 0
    B  中間有洞        同一列有些期有、有些期沒有            誤判率 0
    C  整列全空且矛盾  空白，但相關欄位顯示它應該要有        誤判率低

原本的提案主打「52 家同業普及率」，被降級成參考資訊——它做不到 C 做得到的
事（分辨「這家公司真的沒有」與「我們抓漏了」），而且**它正是「公司真的沒有
某個科目就被永遠標紅」這種誤判的來源**（CTH 直接點出來的問題）。

## 三個判斷各自的關鍵細節

**A** 用 `round(天數差 / 91) - 1`，不能用固定門檻。52 家 1,482 對相鄰期間裡，
111~150 天的 16 筆**全部是 COSTCO**——它的第四季是 16 週（112~119 天），
是正常的一季。固定門檻會把它全部誤判成缺季。

**B** 只看「第一個有值」到「最後一個有值」**之間**有沒有洞，前後空白不算。
`Operating Lease ROU Assets` 只有 28/67 期不是漏抓，是租賃準則 ASC 842 從
2019 才適用。這條沒處理好會製造一大堆假警報。

**C** 用同一家公司的相關欄位互相驗證，不跟別家比。NVDA 實測：有 74 億長期
負債、有利息費用，卻完全沒有借還款現金流——這家公司自己的資料就對不起來。
反過來，如果負債類欄位全都空白，那沒有借還款紀錄是**一致的**，不該標紅。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date

_ISO = re.compile(r"^(\d{4})-(\d{2})-(\d{2})$")

# 13 週 × 7 天。**不是隨手取的數字**，而且不能換成固定門檻——見模組說明的 A。
_QUARTER_DAYS = 91
# 單一缺口最多認幾季。52 家實測沒有任何 >210 天（>2 季）的缺口，
# 真的出現就是資料異常，不該讓程式生出一長串假期間。
_MAX_GAP_QUARTERS = 4
# 一欄有值的模板列低於這個比例就算「整欄稀疏」——欄位在、但幾乎沒東西。
# 實測合成 Q4 失敗時每個流量列都會在那一期出現一個洞，不收攏的話同一件事
# 會被報成 40 個獨立的列問題。
_SPARSE_PERIOD_RATIO = 0.5
# 少於這麼多列就不做稀疏判斷——「一半的列是空的」在只有兩三列的表上沒有意義，
# 那時一個 None 就會讓整欄被判成稀疏。實際的表有 95 列，這道門檻只擋掉退化情況。
_SPARSE_MIN_ROWS = 8
# 稀疏欄超過總期數這個比例 → 不是「某幾期有問題」，是**模板不適用這家公司**。
# 52 家實測：BAC / GS / SCHW / PLD 的稀疏欄佔 90~100%，因為 IS/BS/CF 模板是為
# 製造業設計的，金融股與 REIT 的報表結構完全不同（TODO D8 已記錄，另外處理）。
# 對這種情況列出 21 行稀疏欄沒有意義，直接講「模板不適用」才是有用的訊息。
_TEMPLATE_MISMATCH_RATIO = 0.5


@dataclass(frozen=True)
class QuarterGap:
    after: str      # 缺口前的那個期末日
    before: str     # 缺口後的那個期末日
    count: int      # 中間漏了幾季


@dataclass(frozen=True)
class HoledRow:
    row: str
    have: int       # 在「首末有值之間」實際有幾期
    span: int       # 首末有值之間共幾期


@dataclass(frozen=True)
class Contradiction:
    row: str
    evidence: str   # 憑什麼說它該有值——要講得出來，不然使用者無從判斷


@dataclass(frozen=True)
class SparsePeriod:
    period_end: str   # 期末日；抓不到時退回財季標籤（合成 Q4 常常沒有期末日）
    filled: int       # 這一欄有幾個模板列有值
    total: int        # 模板列共幾列


@dataclass
class QualityReport:
    total_periods: int = 0
    sparse_periods: list[SparsePeriod] = field(default_factory=list)
    missing_quarters: list[QuarterGap] = field(default_factory=list)
    holed: list[HoledRow] = field(default_factory=list)
    contradictions: list[Contradiction] = field(default_factory=list)
    empty_but_plausible: int = 0     # 整列全空、但沒有矛盾可指認的列數
    # 一半以上的期間都整欄稀疏 → 模板不適用這家公司（金融股／REIT），
    # 此時上面那些逐列的判斷意義不大，應該先講這件事
    template_mismatch: bool = False


# 「這一列整列空白時，只有在下列欄位也都空白的情況下才算正常」。
# 全部是會計上必然的關係，不是統計猜測：
#   有負債就會有借還款現金流；有存貨才有存貨變動；有庫藏股才有買回。
_COHERENCE: dict[str, tuple[str, ...]] = {
    # 有負債餘額就會有借還款現金流。**證據刻意不放 `Interest Expense`**——
    # 融資租賃也會產生利息費用，用它當證據會對「有租賃、無借款」的公司誤報
    # （實測 ARLO 就是這樣被誤判）。
    "Debt Proceeds":              ("Long-term Debt", "Short-term Debt"),
    "Debt Repayments":            ("Long-term Debt", "Short-term Debt"),
    "Current Portion of LT Debt": ("Long-term Debt", "Short-term Debt"),
    "Change in Inventories":      ("Inventories",),
    "Share Repurchases":          ("Treasury Stock",),
    "Minority Interest":          ("Noncontrolling Interests",),
    "Noncontrolling Interests":   ("Minority Interest",),
    # 認列了使用權資產就一定有對應的租賃負債，這是 ASC 842 的結構要求
    "Op. Lease Liabilities, current": ("Operating Lease ROU Assets",),
    "Op. Lease Liabilities, LT":      ("Operating Lease ROU Assets",),
    # 刻意**不放** `Amortization of Intangibles ← Intangible Assets, net`：
    # 無形資產可能是非確定年限（不攤銷），攤銷也常併進 D&A 一起報。
    # 實測 52 家幾乎人人中招，是規則太寬不是真的漏抓。
}


def _parse(d: str | None) -> date | None:
    m = _ISO.match((d or "").strip())
    if m is None:
        return None
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def missing_quarters(period_ends: list[str]) -> list[QuarterGap]:
    """A：期末日序列 → 中間漏了哪幾段、各漏幾季。

    解析不了的日期直接跳過（不猜）。相同日期重複出現不算缺口——實測 SNOW 有
    兩欄期末日都是 `2022-01-31`，那是重複列的問題（TODO G13），不是缺季。
    """
    dates = sorted({d for d in (_parse(e) for e in period_ends) if d is not None})
    gaps: list[QuarterGap] = []
    for a, b in zip(dates, dates[1:]):
        n = round((b - a).days / _QUARTER_DAYS) - 1
        if 0 < n <= _MAX_GAP_QUARTERS:
            gaps.append(QuarterGap(after=a.isoformat(), before=b.isoformat(), count=n))
    return gaps


def _template_rows() -> frozenset[str]:
    """三張模板的列名。延後 import 避免匯入期相依。"""
    from fetcher_gaap import BS_TEMPLATE, CF_TEMPLATE, IS_TEMPLATE
    return frozenset(r[0] for T in (IS_TEMPLATE, BS_TEMPLATE, CF_TEMPLATE) for r in T)


def _rows(table, known: frozenset[str] | None = None) -> list[tuple[str, list]]:
    """{列名: 值} —— 同名列取第一個，跟 Excel 上看到的一致。

    **只評估模板列，`Other (as reported)` 的 overflow 列一律排除。** overflow
    的 key 是 XBRL concept name，公司改一次 tag 就長出新的一列、舊列從此空白，
    本來就會斷斷續續（見 TODO G4）。把它們算進來的話 NVDA 會報 85 列有洞，
    其中絕大多數是「這家公司特有的科目只在某幾年出現」這種非問題。
    """
    known = _template_rows() if known is None else known
    out: list[tuple[str, list]] = []
    seen: set[str] = set()
    n = len(table.period_ends or table.quarter_labels)
    for i, name in enumerate(table.concepts):
        if not name or name in seen or name not in known:
            continue
        seen.add(name)
        out.append((name, list(table.values[i][:n])))
    return out


def _has(v) -> bool:
    return isinstance(v, (int, float))


def sparse_periods(table, known: frozenset[str] | None = None) -> list[SparsePeriod]:
    """D：欄位存在、但整排幾乎都空的期間。

    典型成因是那一年的 Q4 合成失敗——每個流量列都會在那一期留一個洞。
    那是**一個期間問題**，不是 40 個列問題。
    """
    rows = _rows(table, known)
    if len(rows) < _SPARSE_MIN_ROWS:
        return []
    ends = list(table.period_ends or [])
    labels = list(table.quarter_labels or [])
    n = max(len(ends), len(labels))
    out = []
    for j in range(n):
        filled = sum(1 for _, vals in rows if j < len(vals) and _has(vals[j]))
        if filled < len(rows) * _SPARSE_PERIOD_RATIO:
            # 合成 Q4 的欄位常常沒有期末日（年報只帶到年月），退回財季標籤，
            # 不然畫面上會出現一個沒有名字的期間，使用者不知道在講哪一欄
            end = str(ends[j]) if j < len(ends) and ends[j] else ""
            out.append(SparsePeriod(period_end=end or (labels[j] if j < len(labels) else "?"),
                                    filled=filled, total=len(rows)))
    return out


def holed_rows(table, known: frozenset[str] | None = None) -> list[HoledRow]:
    """B：只看「第一個有值」到「最後一個有值」之間的洞。

    兩個排除，兩個都是實測踩出來的：

    1. **前後空白不算**——那多半是會計準則開始適用（`Operating Lease ROU
       Assets` 從 ASC 842 的 2019 年起）或科目停用，不是漏抓
    2. **整欄稀疏的期間不算**——那一期是所有列一起缺，歸 D 報一次就好，
       不要讓同一件事在這裡被報 40 次
    """
    sparse = sparse_periods(table, known)
    ends_raw = list(table.period_ends or [])
    labels = list(table.quarter_labels or [])
    n = max(len(ends_raw), len(labels))
    ends = [(str(ends_raw[j]) if j < len(ends_raw) and ends_raw[j] else "")
            or (labels[j] if j < len(labels) else "?") for j in range(n)]
    skip = {sp.period_end for sp in sparse}
    out: list[HoledRow] = []
    for name, vals in _rows(table, known):
        keep = [v for j, v in enumerate(vals) if j >= len(ends) or ends[j] not in skip]
        idx = [i for i, v in enumerate(keep) if _has(v)]
        if len(idx) < 2:
            continue                       # 全空歸 C；只有一期無從判斷
        span = idx[-1] - idx[0] + 1
        if len(idx) < span:
            out.append(HoledRow(row=name, have=len(idx), span=span))
    return sorted(out, key=lambda h: (h.have / h.span, h.have))


def contradictions(table, known: frozenset[str] | None = None) -> list[Contradiction]:
    """C：整列全空，但同一家公司的相關欄位顯示它應該要有值。"""
    rows = dict(_rows(table, known))
    filled = {k for k, v in rows.items() if any(_has(x) for x in v)}
    out: list[Contradiction] = []
    for name, related in _COHERENCE.items():
        if name not in rows or name in filled:
            continue                       # 沒這一列、或有值 → 不是 C 的範圍
        evidence = [r for r in related if r in filled]
        if evidence:
            out.append(Contradiction(row=name, evidence="、".join(evidence)))
    return out


def assess(table, known: frozenset[str] | None = None) -> QualityReport:
    """跑完三個判斷。純函式，不打網路、不看別家公司。"""
    rows = _rows(table, known)
    holed = holed_rows(table, known)
    contra = contradictions(table, known)
    contra_names = {c.row for c in contra}
    empty = sum(1 for name, vals in rows
                if not any(_has(v) for v in vals) and name not in contra_names)
    sparse = sparse_periods(table, known)
    n_periods = len(table.period_ends or table.quarter_labels)
    return QualityReport(
        total_periods=n_periods,
        sparse_periods=sparse,
        template_mismatch=bool(n_periods) and len(sparse) > n_periods * _TEMPLATE_MISMATCH_RATIO,
        missing_quarters=missing_quarters(list(table.period_ends or [])),
        holed=holed,
        contradictions=contra,
        empty_but_plausible=empty,
    )
