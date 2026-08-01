"""
nongaap_layout.py — `Data_NonGAAP` 的固定模板版面。

═══════════════════════════════════════════════════════════════════════════════
為什麼要固定模板
═══════════════════════════════════════════════════════════════════════════════
Non-GAAP 依定義是公司自訂的，如果照 AI 回什麼就排什麼，每家公司的列都不一樣，
下游 skill 得先探索才知道有什麼，而且沒跑過的公司品質必然比跑過的差。

解法沿用 `Data_Financials` 已經驗證過的模式：**固定行數萬能模板 + overflow**。
Core 是跨產業都存在的指標（永遠有那一列，沒資料就空白），公司自己的特殊指標
掉進 overflow 區——overflow 不需要事先認識這家公司，這就是通用性的來源。

═══════════════════════════════════════════════════════════════════════════════
Core 收哪些：實測 32 家（大中小型跨產業）8-K 新聞稿的覆蓋率
═══════════════════════════════════════════════════════════════════════════════
（`scripts/survey_nongaap_metrics.py`，扣掉完全不報 Non-GAAP 的 AAPL/AMZN/COST
後 29 家為分母）

    Non-GAAP Net Income        79%      Non-GAAP Operating Margin   48%
    Non-GAAP Diluted EPS       79%      Non-GAAP Operating Expenses 41%
    Free Cash Flow             76%      Non-GAAP Revenue            38%
    Non-GAAP Operating Income  66%      Non-GAAP Gross Profit       38%
    Non-GAAP Effective Tax Rate 62%     Non-GAAP Diluted Shares     31%
    Non-GAAP Gross Margin      59%      Adjusted EBITDA             21%
    **Non-GAAP Net Margin       0%**  ← 沒有任何一家會寫，只能推導

判準是「跨產業都有、定義沒爭議」。SaaS 專屬指標（ARR / RPO / cRPO / Billings /
NRR）**刻意不進 core**——收進來等於開產業別模板，那是特例的另一種形式。
Adjusted EBITDA 覆蓋率只有 21% 但仍收，因為它是中小型公司的主要指標
（ARLO / AEIS / FORM 都報），對大公司也只是多兩行空白。

═══════════════════════════════════════════════════════════════════════════════
GAAP 對照行為什麼從新聞稿抓，不從 Data_Financials 拉
═══════════════════════════════════════════════════════════════════════════════
Non-GAAP 的季度標籤系統性晚一季（TODO 第 4 項的 off-by-one）。把 XBRL 的
`FY2026Q2` 毛利率放進本表的 `FY2026Q2` 欄，兩者其實差一季——會變成一個**無聲的
錯誤對比**，數字看起來完全正常。

8-K 的調節表本來就同時列 GAAP 與 Non-GAAP（ARLO 原文：「GAAP subscriptions and
services gross margin of 83.7% **and** ... non-GAAP ... 85.4%」），同一份文件、
同一個期間。所以對照值一律從同一份新聞稿取，等 off-by-one 修好之前不跨表拉。

═══════════════════════════════════════════════════════════════════════════════
調節表的「其他」用殘差倒算
═══════════════════════════════════════════════════════════════════════════════
    其他 = Non-GAAP 淨利 − GAAP 淨利 − 具名項目合計

好處是這張表會自己對帳：AI 漏抓了某個調整項，殘差就會變大，一眼看得出「有一塊
沒被解釋」，而不是默默少一塊。殘差為 0 代表具名項目完整解釋了整座橋。

具名項目取實測覆蓋率 ≥ 59% 的前七名（29 家為分母）：
    SBC 90% / 重組資遣 79% / 減損 69% / 訴訟和解 66% /
    無形資產攤銷 66% / 併購相關 62% / 調整項稅務影響 59%
"""

from __future__ import annotations

from typing import Any

from fetcher_gaap import StatementTable

# ── 版面常數（可調整）───────────────────────────────────────────────────────

SECTION_CORE   = "Non-GAAP Core"
SECTION_RECON  = "GAAP → Non-GAAP 調節"
SECTION_OTHER  = "Other Non-GAAP (as reported)"
SECTION_ANNUAL = "Annual (FY)"

ALL_SECTIONS = (SECTION_CORE, SECTION_RECON, SECTION_OTHER, SECTION_ANNUAL)

GAAP_PREFIX  = "  GAAP "          # 對照行縮排，視覺上掛在 Non-GAAP 那列下面
RESIDUAL_ROW = "  + 其他 Non-GAAP 調整（殘差）"
FY_SUFFIX    = " (FY)"

SRC_PR      = "8-K press release"
SRC_PR_GAAP = "8-K press release (GAAP 對照)"

# ── Core 模板 ───────────────────────────────────────────────────────────────
#
# (顯示名稱, 從指標表查的鍵, GAAP 對照行的顯示尾綴 or None)
# 順序即 sheet 上的列序。要增減 core 行就改這張表。

CORE_ROWS: list[tuple[str, str, str | None]] = [
    ("Non-GAAP Revenue",            "Non-GAAP Revenue",            "Revenue"),
    ("Non-GAAP Gross Profit",       "Non-GAAP Gross Profit",       "Gross Profit"),
    ("Non-GAAP Gross Margin",       "Non-GAAP Gross Margin",       "Gross Margin"),
    ("Non-GAAP Operating Expenses", "Non-GAAP Operating Expenses", None),
    ("Non-GAAP Operating Income",   "Non-GAAP Operating Income",   "Operating Income"),
    ("Non-GAAP Operating Margin",   "Non-GAAP Operating Margin",   "Operating Margin"),
    ("Non-GAAP Net Income",         "Non-GAAP Net Income",         "Net Income"),
    ("Non-GAAP Net Margin",         "Non-GAAP Net Margin",         "Net Margin"),
    ("Non-GAAP Diluted EPS",        "Non-GAAP Diluted EPS",        "Diluted EPS"),
    ("Non-GAAP Diluted Shares",     "Non-GAAP Diluted Shares",     None),
    ("Non-GAAP Effective Tax Rate", "Non-GAAP Effective Tax Rate", None),
    ("Adjusted EBITDA",             "Adjusted EBITDA",             None),
    ("Adjusted EBITDA Margin",      "Adjusted EBITDA Margin",      None),
    ("Free Cash Flow",              "Free Cash Flow",              None),
    ("Free Cash Flow Margin",       "Free Cash Flow Margin",       None),
]

# 需要推導的 core 行（新聞稿不會寫，只能自己算）
# 淨利率：32 家全部沒寫。分母優先用 Non-GAAP Revenue，沒有就退 GAAP Revenue，
# 兩者都沒有就留空——不拿別的科目硬湊。
DERIVED_ROWS = {
    "Non-GAAP Net Margin": (
        "DERIVED = Non-GAAP Net Income / Revenue",
        ("Non-GAAP Net Income", ("Non-GAAP Revenue", "GAAP Revenue")),
    ),
    "  GAAP Net Margin": (
        "DERIVED = GAAP Net Income / GAAP Revenue",
        ("GAAP Net Income", ("GAAP Revenue", "Non-GAAP Revenue")),
    ),
}

# ── 調節表 ──────────────────────────────────────────────────────────────────
#
# (顯示名稱, 從指標表查的鍵)。值是「加到 GAAP 淨利上以得到 Non-GAAP 淨利」的
# 帶號金額，所以稅務影響通常是負的。

ADDBACK_ROWS: list[tuple[str, str]] = [
    ("  + 股權獎酬 SBC",        "Stock-Based Compensation"),
    ("  + 無形資產攤銷",         "Amortization of Intangibles"),
    ("  + 重組／資遣",           "Restructuring Charges"),
    ("  + 減損",                 "Impairment Charges"),
    ("  + 訴訟／和解",           "Litigation and Settlement"),
    ("  + 併購相關費用",         "Acquisition-Related Costs"),
    ("  + 調整項之稅務影響",     "Tax Effect of Adjustments"),
]

_RECON_TOP    = "GAAP Net Income"
_RECON_BOTTOM = "= Non-GAAP Net Income"


# ── 工具 ────────────────────────────────────────────────────────────────────

def _get(metrics: dict[str, Any], key: str) -> Any:
    """不分大小寫、忽略前後空白地取值。"""
    if key in metrics:
        return metrics[key]
    folded = key.casefold().strip()
    for name, value in metrics.items():
        if name.casefold().strip() == folded:
            return value
    return None


def _derive_margin(metrics: dict[str, Any], numerator_key: str,
                   denominator_keys: tuple[str, ...]) -> float | None:
    num = _get(metrics, numerator_key)
    if num is None:
        return None
    for key in denominator_keys:
        den = _get(metrics, key)
        if den is None:
            continue
        try:
            if float(den) == 0.0:
                continue
            return float(num) / float(den) * 100.0
        except (TypeError, ValueError):
            continue
    return None


def _residual(metrics: dict[str, Any]) -> float | None:
    """殘差 = Non-GAAP 淨利 − GAAP 淨利 − 具名調整項合計。

    橋的兩端缺任一端就回 None——用 0 代替會讓「沒資料」看起來像「完全對帳」。
    """
    ng = _get(metrics, "Non-GAAP Net Income")
    gaap = _get(metrics, "GAAP Net Income")
    if ng is None or gaap is None:
        return None
    named = 0.0
    for _display, key in ADDBACK_ROWS:
        value = _get(metrics, key)
        if value is not None:
            try:
                named += float(value)
            except (TypeError, ValueError):
                pass
    try:
        return float(ng) - float(gaap) - named
    except (TypeError, ValueError):
        return None


def _consumed_keys() -> set[str]:
    """已經被固定模板吃掉的鍵——不可再出現在 overflow 區重複一次。"""
    keys = {k.casefold() for _d, k, _g in CORE_ROWS}
    keys |= {k.casefold() for _d, k in ADDBACK_ROWS}
    keys |= {f"GAAP {g}".casefold() for _d, _k, g in CORE_ROWS if g}
    keys |= {"gaap net income", "non-gaap net income"}
    return keys


# ── Public API ──────────────────────────────────────────────────────────────

def build_nongaap_table(
    ticker: str,
    per_quarter: dict[str, dict[str, Any]],
    quarter_labels: list[str],
    filing_dates: list[str],
) -> StatementTable:
    """組出 Data_NonGAAP。

    Args:
        per_quarter: {季度標籤: {正規化後的指標名: 值}}，已經過
                     `_normalize_nongaap_metrics` 與 `_canonicalize_metric_name`
        quarter_labels / filing_dates: 欄位（舊→新）

    即使完全沒有資料也會回一張只有骨架的表——讀不到 sheet 與讀到空 sheet 是兩種
    不同的訊號，前者無法區分「這家沒報 Non-GAAP」與「抓取失敗」。
    """
    concepts: list[str] = []
    labels: list[str] = []
    values: list[list[Any]] = []

    def add(name: str, source: str, getter) -> None:
        concepts.append(name)
        labels.append(source)
        values.append([getter(per_quarter.get(q, {})) for q in quarter_labels])

    def add_section(name: str) -> None:
        concepts.append(name)
        labels.append("")
        values.append([None] * len(quarter_labels))

    # ── Core ────────────────────────────────────────────────────────────
    add_section(SECTION_CORE)
    for display, key, gaap_tail in CORE_ROWS:
        if display in DERIVED_ROWS:
            formula, (num, dens) = DERIVED_ROWS[display]
            add(display, formula, lambda m, n=num, d=dens: _derive_margin(m, n, d))
        else:
            add(display, SRC_PR, lambda m, k=key: _get(m, k))

        if gaap_tail:
            gaap_display = f"{GAAP_PREFIX}{gaap_tail}"
            if gaap_display in DERIVED_ROWS:
                formula, (num, dens) = DERIVED_ROWS[gaap_display]
                add(gaap_display, formula,
                    lambda m, n=num, d=dens: _derive_margin(m, n, d))
            else:
                add(gaap_display, SRC_PR_GAAP,
                    lambda m, k=f"GAAP {gaap_tail}": _get(m, k))

    # ── 調節表 ──────────────────────────────────────────────────────────
    add_section(SECTION_RECON)
    add(_RECON_TOP, SRC_PR_GAAP, lambda m: _get(m, "GAAP Net Income"))
    for display, key in ADDBACK_ROWS:
        add(display, SRC_PR, lambda m, k=key: _get(m, k))
    add(RESIDUAL_ROW, "DERIVED = Non-GAAP 淨利 − GAAP 淨利 − 具名項目合計", _residual)
    add(_RECON_BOTTOM, SRC_PR, lambda m: _get(m, "Non-GAAP Net Income"))

    # ── Overflow：模板沒收的公司自訂指標，照原名保留 ─────────────────────
    consumed = _consumed_keys()
    other: list[str] = []
    annual: list[str] = []
    for q in quarter_labels:
        for name in per_quarter.get(q, {}):
            if name.casefold() in consumed:
                continue
            bucket = annual if name.endswith(FY_SUFFIX) else other
            if name not in bucket:
                bucket.append(name)

    if other:
        add_section(SECTION_OTHER)
        for name in other:
            add(name, SRC_PR, lambda m, k=name: _get(m, k))

    if annual:
        add_section(SECTION_ANNUAL)
        for name in annual:
            add(name, SRC_PR, lambda m, k=name: _get(m, k))

    return StatementTable(
        sheet_name="Data_NonGAAP",
        quarter_labels=list(quarter_labels),
        filing_dates=list(filing_dates),
        concepts=concepts,
        values=values,
        ticker=ticker,
        labels=labels,
    )
