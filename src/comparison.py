"""comparison.py — 跨公司財務比較的資料抓取與重組。

把多家公司各自的 fetch_gaap_statements() 結果，重組成
{指標名: {ticker: {period_label: value}}} 這種給 comparison_writer.py
直接寫表用的形狀。單一公司抓取失敗不中斷整體流程，記錄下來繼續下一家
（比照 fetcher_gaap.collect_gaps() 的「跳過不中斷」原則，但這裡是公司
層級的跳過，不是同一家公司內部的科目缺漏，所以不共用同一套機制）。
"""

from __future__ import annotations

from dataclasses import dataclass, field
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


def _filter_by_year(table: StatementTable, start_year: int | None, end_year: int | None) -> StatementTable:
    """依 period_ends 的年份篩選欄位。沒有 period_ends 資料的欄一律保留
    （不因為篩不了就整欄丟掉，寧可多顯示也不要漏資料）。"""
    if start_year is None and end_year is None:
        return table

    keep = []
    for i, end in enumerate(table.period_ends or []):
        if not end:
            keep.append(i)
            continue
        try:
            year = int(end[:4])
        except (TypeError, ValueError):
            keep.append(i)
            continue
        if start_year is not None and year < start_year:
            continue
        if end_year is not None and year > end_year:
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
    start_year: int | None,
    end_year: int | None,
    max_filings: int = 80,
    max_annual_filings: int = 20,
    on_company_start: Callable[[str, int, int], None] | None = None,
    progress_cb: Callable[[int, int, str], None] | None = None,
) -> ComparisonResult:
    """對每個 ticker 抓資料、抽出選定指標，重組成跨公司比較用的資料結構。

    `on_company_start(ticker, current, total)` 在每家公司開始抓之前呼叫一次
    （就算後面失敗了也會呼叫過），讓呼叫端知道現在正在試哪一家，不會整段
    安靜到看起來像卡死。`progress_cb(current, total, label)` 原樣轉給
    `fetcher_gaap.report_progress()`，取得單一公司內部逐份 filing 的細部進度
    （跟 Tab1 GAAP 抓取用的是同一套機制），一家公司要抓很多季時能看到進度
    在動，不用等整家公司抓完才有畫面反應。
    """
    result = ComparisonResult(metrics={name: {} for name in metric_names})
    sheet_name = _sheet_name_for(frequency)

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

        raw_table = _filter_by_year(raw_table, start_year, end_year)
        ratio_table = build_ratio_table(raw_table)

        aligned = _aligned_labels(raw_table, frequency)

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
