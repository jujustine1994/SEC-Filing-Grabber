"""comparison.py — 跨公司財務比較的資料抓取與重組。

把多家公司各自的 fetch_gaap_statements() 結果，重組成
{指標名: {ticker: {period_label: value}}} 這種給 comparison_writer.py
直接寫表用的形狀。單一公司抓取失敗不中斷整體流程，記錄下來繼續下一家
（比照 fetcher_gaap.collect_gaps() 的「跳過不中斷」原則，但這裡是公司
層級的跳過，不是同一家公司內部的科目缺漏，所以不共用同一套機制）。
"""

from __future__ import annotations

import calendar
from dataclasses import dataclass, field
from datetime import date
from typing import Callable, Literal

from fetcher_gaap import StatementTable, fetch_gaap_statements, report_progress
from fiscal_input import calendarized_quarter_of, calendarized_year_of
from ratios import build_ratio_table, RATIO_DEFS


_RATIO_NAMES = {name for name, _, _, _ in RATIO_DEFS}


@dataclass(frozen=True)
class CompanyFetchError:
    ticker: str
    error_type: str


@dataclass
class ComparisonResult:
    metrics: dict[str, dict[str, dict[str, float | None]]] = field(default_factory=dict)
    period_ends: dict[str, dict[str, str]] = field(default_factory=dict)
    failures: list[CompanyFetchError] = field(default_factory=list)
    # {ticker: {日曆季: 該公司自己的財季標籤}}。欄位鍵是跨公司對齊過的日曆季，
    # 各公司在同一欄的財季不一樣（NVDA 的 2025Q2 是 FY2026Q2、AMD 是 FY2025Q2）
    # ——Compare_Data 最上方那張對應表就是在講這件事（G2）。
    fiscal_labels: dict[str, dict[str, str]] = field(default_factory=dict)
    # {ticker: {推算 Q4 的日曆季}}。季報表裡出現 Q4 一定是推算來的：SEC 沒有
    # Q4 的 10-Q，那一欄只可能由 fetcher_gaap._synthesize_q4() 補進來
    # （「年報 − Q1 − Q2 − Q3」）。給說明 sheet 的第 5 條打勾用（G7）。
    synthetic_q4: dict[str, set[str]] = field(default_factory=dict)


def _sheet_name_for(frequency: Literal["quarterly", "annual"]) -> str:
    return "Data_Financials(Q)" if frequency == "quarterly" else "Data_Financials(Y)"


def _aligned_labels(
    table: StatementTable, frequency: Literal["quarterly", "annual"]
) -> list[str]:
    """把各公司自己的財季／財年標籤換成跨公司共用的日曆期間標籤。

    2026-08-22 修：原本直接拿財季標籤（`FY2026Q2`）當共同欄位鍵。財年結束月
    不同的公司，同一個標籤指的根本不是同一段時間——NVDA 一月結算，它的
    FY2026Q2 結束在 2025-07-27，卻跟 AMD 結束在 2026-06-27 的 FY2026Q2 疊在
    同一欄，整條線在日曆時間上偏移約一年。改成用期末日算出的日曆期間
    （`2025Q2` / `2025`，判準見 `fiscal_input.calendarized_quarter_of()`）。

    期末日抓不到（空字串）時退回原本的財季標籤——寧可那一欄自己站一格，
    也不要因為算不出來就把整季資料丟掉。
    """
    convert = calendarized_quarter_of if frequency == "quarterly" else calendarized_year_of
    ends = table.period_ends or []
    return [
        convert(ends[i] if i < len(ends) else "") or label
        for i, label in enumerate(table.quarter_labels)
    ]


def _fiscal_label_map(aligned: list[str], table: StatementTable) -> dict[str, str]:
    """{日曆季: 原始財季標籤}。逐期從實際期末日算出來的對應關係，公司哪一年
    改過財年，那一欄自己就會反映出來，不需要任何例外處理（G2 選對應表而不是
    「只寫財年開始月份」的理由）。"""
    return {
        label: table.quarter_labels[i]
        for i, label in enumerate(aligned)
        if i < len(table.quarter_labels)
    }


def parse_period_bound(value: str | int | None, *, is_end: bool) -> str | None:
    """把使用者輸入的期間邊界轉成用來比對 `period_ends` 的 ISO 日期字串
    （`period_ends` 本身就是 `YYYY-MM-DD`，字典序比較即等於日期先後）。

    接受 `YYYY`／`YYYY-MM`／`YYYY-MM-DD`（也接受 `int`，向後相容舊 config
    存的純年份）。起始邊界取當期第一天、結束邊界取當期最後一天，好讓
    `2024` 涵蓋整年、`2024-06` 涵蓋整月——跟這格「快照時間點」YYYYMMDD
    輸入是同一個 tkinter 沒有原生日期選擇器的取捨，純文字＋容錯解析。
    """
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None

    parts = text.split("-")
    try:
        if len(parts) == 1:
            year = int(parts[0])
            return f"{year:04d}-12-31" if is_end else f"{year:04d}-01-01"
        if len(parts) == 2:
            year, month = int(parts[0]), int(parts[1])
            if is_end:
                last_day = calendar.monthrange(year, month)[1]
                return f"{year:04d}-{month:02d}-{last_day:02d}"
            return f"{year:04d}-{month:02d}-01"
        if len(parts) == 3:
            year, month, day = int(parts[0]), int(parts[1]), int(parts[2])
            date(year, month, day)  # 只借來驗證日期合法，值本身仍用字串格式化
            return f"{year:04d}-{month:02d}-{day:02d}"
    except ValueError:
        pass
    raise ValueError(f"Cannot parse period {text!r}; expected YYYY, YYYY-MM, or YYYY-MM-DD")


def _filter_by_period_end(table: StatementTable, start: str | None, end: str | None) -> StatementTable:
    """依 period_ends（期末日）篩選欄位，比對的是 `parse_period_bound()` 算出來的
    ISO 日期字串邊界。沒有 period_ends 資料的欄一律保留（不因為篩不了就整欄
    丟掉，寧可多顯示也不要漏資料）。"""
    if start is None and end is None:
        return table

    keep = []
    for i, period_end in enumerate(table.period_ends or []):
        if not period_end:
            keep.append(i)
            continue
        if start is not None and period_end < start:
            continue
        if end is not None and period_end > end:
            continue
        keep.append(i)

    return StatementTable(
        sheet_name=table.sheet_name,
        quarter_labels=[table.quarter_labels[i] for i in keep],
        filing_dates=[table.filing_dates[i] for i in keep] if table.filing_dates else [],
        concepts=table.concepts,
        values=[[row[i] for i in keep] for row in table.values],
        ticker=table.ticker,
        labels=table.labels,
        period_ends=[table.period_ends[i] for i in keep] if table.period_ends else [],
    )


def build_comparison(
    tickers: list[str],
    identity: str,
    metric_names: list[str],
    *,
    frequency: Literal["quarterly", "annual"],
    start_year: str | int | None,
    end_year: str | int | None,
    max_filings: int = 80,
    max_annual_filings: int = 20,
    on_company_start: Callable[[str, int, int], None] | None = None,
    progress_cb: Callable[[int, int, str], None] | None = None,
) -> ComparisonResult:
    """對每個 ticker 抓資料、抽出選定指標，重組成跨公司比較用的資料結構。

    `start_year`／`end_year`（參數名沿用舊稱，實際接受 `YYYY`／`YYYY-MM`／
    `YYYY-MM-DD` 三種寫法，見 `parse_period_bound()`；F7，2026-09-03）篩的是
    **期末日**落在區間內，不是純年份比較。

    `on_company_start(ticker, current, total)` 在每家公司開始抓之前呼叫一次
    （就算後面失敗了也會呼叫過），讓呼叫端知道現在正在試哪一家，不會整段
    安靜到看起來像卡死。`progress_cb(current, total, label)` 原樣轉給
    `fetcher_gaap.report_progress()`，取得單一公司內部逐份 filing 的細部進度
    （跟 Tab1 GAAP 抓取用的是同一套機制），一家公司要抓很多季時能看到進度
    在動，不用等整家公司抓完才有畫面反應。
    """
    result = ComparisonResult(metrics={name: {} for name in metric_names})
    sheet_name = _sheet_name_for(frequency)
    start_bound = parse_period_bound(start_year, is_end=False)
    end_bound = parse_period_bound(end_year, is_end=True)

    valid_tickers = [t.strip().upper() for t in tickers if t.strip()]
    total = len(valid_tickers)

    for i, ticker in enumerate(valid_tickers, start=1):
        if on_company_start is not None:
            try:
                on_company_start(ticker, i, total)
            except Exception:
                pass  # 進度回報是錦上添花，回呼本身出錯不能拖垮抓取
        try:
            with report_progress(progress_cb):
                tables = fetch_gaap_statements(
                    ticker, identity, max_filings=max_filings,
                    max_annual_filings=max_annual_filings,
                    fetch_quarterly=(frequency == "quarterly"),
                    # 季度模式也要抓年報：D0-1 的 Q4 合成
                    # （fetcher_gaap._synthesize_q4()）靠「年報－Q1－Q2－Q3」
                    # 湊出 10-Q 沒有的 Q4，沒有年報這步就整個跳過，日曆 Q4
                    # 對齊到日曆年結束的公司（如 AMD/INTC）會整欄空白。
                    # 這裡永遠抓年報，只是當 Q4 合成材料用，不影響輸出頻率。
                    fetch_annual=True,
                )
        except Exception as e:
            result.failures.append(CompanyFetchError(ticker=ticker, error_type=type(e).__name__))
            continue

        raw_table = next((t for t in tables if t.sheet_name == sheet_name), None)
        if raw_table is None:
            result.failures.append(CompanyFetchError(ticker=ticker, error_type="NoDataForFrequency"))
            continue

        raw_table = _filter_by_period_end(raw_table, start_bound, end_bound)
        ratio_table = build_ratio_table(raw_table)

        aligned = _aligned_labels(raw_table, frequency)
        result.fiscal_labels[ticker] = _fiscal_label_map(aligned, raw_table)
        # 季報表的 Q4 一定是推算的（見 ComparisonResult.synthetic_q4）；年報
        # 標籤是 `FY2025` 沒有 Q，這裡自然算出空集合，不用另外分支。
        result.synthetic_q4[ticker] = {
            label for label, fiscal in result.fiscal_labels[ticker].items()
            if fiscal.endswith("Q4")
        }

        period_map: dict[str, str] = {}
        for j, label in enumerate(aligned):
            end = raw_table.period_ends[j] if j < len(raw_table.period_ends or []) else ""
            period_map[label] = end
        result.period_ends[ticker] = period_map

        for metric_name in metric_names:
            source_table = ratio_table if metric_name in _RATIO_NAMES else raw_table
            if source_table is None or metric_name not in source_table.concepts:
                continue
            # build_ratio_table() 逐欄對應 raw_table，欄序一致，所以共用同一份
            # aligned 標籤；長度不同（理論上不該發生）時以較短的為準，不硬配。
            row = source_table.values[source_table.concepts.index(metric_name)]
            result.metrics.setdefault(metric_name, {})[ticker] = dict(
                zip(aligned, row)
            )

    return result
