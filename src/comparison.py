"""comparison.py — 跨公司財務比較的資料抓取與重組。

把多家公司各自的 fetch_gaap_statements() 結果，重組成
{指標名: {ticker: {period_label: value}}} 這種給 comparison_writer.py
直接寫表用的形狀。單一公司抓取失敗不中斷整體流程，記錄下來繼續下一家
（比照 fetcher_gaap.collect_gaps() 的「跳過不中斷」原則，但這裡是公司
層級的跳過，不是同一家公司內部的科目缺漏，所以不共用同一套機制）。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

from fetcher_gaap import StatementTable, fetch_gaap_statements
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
) -> ComparisonResult:
    """對每個 ticker 抓資料、抽出選定指標，重組成跨公司比較用的資料結構。"""
    result = ComparisonResult(metrics={name: {} for name in metric_names})
    sheet_name = _sheet_name_for(frequency)

    for ticker in tickers:
        ticker = ticker.strip().upper()
        if not ticker:
            continue
        try:
            tables = fetch_gaap_statements(
                ticker, identity, max_filings=max_filings,
                max_annual_filings=max_annual_filings,
                fetch_quarterly=(frequency == "quarterly"),
                fetch_annual=(frequency == "annual"),
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

        period_map: dict[str, str] = {}
        for i, label in enumerate(raw_table.quarter_labels):
            end = raw_table.period_ends[i] if i < len(raw_table.period_ends or []) else ""
            period_map[label] = end
        result.period_ends[ticker] = period_map

        for metric_name in metric_names:
            source_table = ratio_table if metric_name in _RATIO_NAMES else raw_table
            if source_table is None or metric_name not in source_table.concepts:
                continue
            row = source_table.values[source_table.concepts.index(metric_name)]
            result.metrics.setdefault(metric_name, {})[ticker] = dict(
                zip(source_table.quarter_labels, row)
            )

    return result
