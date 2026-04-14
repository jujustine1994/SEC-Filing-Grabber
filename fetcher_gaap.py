"""
fetcher_gaap.py — Fetch XBRL GAAP financial statements from SEC EDGAR via edgartools.

Fetches quarterly data from up to `max_filings` 10-Q filings (newest first),
building a time-series with quarterly columns oldest→newest.

Public API:
    fetch_gaap_statements(ticker, identity, max_filings=80) -> list[StatementTable]

Sheet outputs:
    Data_IS        — 18-row fixed template (NaN for missing items)
    Data_BS        — all consolidated rows, dynamic union across quarters
    Data_CF        — all consolidated rows, dynamic union across quarters
    Data_Seg_*     — one sheet per IS concept that has segment breakdowns
    Data_Meta      — ticker / company / date / quarter count
"""

from __future__ import annotations

import math
import re
import sys
from dataclasses import dataclass
from datetime import date
from typing import Any

import pandas as pd
from edgar import Company, set_identity as set_identity


# ── Data contract ────────────────────────────────────────────────────────

@dataclass
class StatementTable:
    """One financial statement pre-structured for Excel output."""
    sheet_name:     str
    quarter_labels: list[str]
    filing_dates:   list[str]
    concepts:       list[str]
    values:         list[list[Any]]
    ticker:         str = ""


# ── Constants ─────────────────────────────────────────────────────────────

META_COLS: set[str] = {
    'concept', 'label', 'standard_concept', 'level', 'abstract',
    'dimension', 'is_breakdown', 'dimension_axis', 'dimension_member',
    'dimension_member_label', 'dimension_label', 'unit', 'point_in_time',
    'balance', 'weight', 'preferred_sign', 'parent_concept', 'parent_abstract_concept',
}

IS_TEMPLATE: list[tuple[str, str | None, str, str]] = [
    ("Revenue",                "Revenue",                        "RevenueFromContractWithCustomer",                          "IS"),
    ("Cost of Revenue",        "CostOfGoodsAndServicesSold",     "CostOfGoodsSold",                                          "IS"),
    ("Gross Profit",           "GrossProfit",                    "GrossProfit",                                              "IS"),
    ("R&D Expense",            "ResearchAndDevelopmentExpenses", "ResearchAndDevelopment",                                   "IS"),
    ("SG&A Expense",           "SellingGeneralAndAdminExpenses", "SellingGeneralAndAdmin",                                   "IS"),
    ("D&A",                    "DepreciationExpense",            "DepreciationDepletionAndAmortization",                     "CF"),
    ("Other Operating Expense","OtherOperatingExpenses",         "OtherOperatingExpense",                                    "IS"),
    ("Total Operating Expense","TotalOperatingExpenses",         "OperatingExpenses",                                        "IS"),
    ("Operating Income",       "OperatingIncomeLoss",            "OperatingIncome",                                          "IS"),
    ("Interest Expense",       "InterestExpense",                "InterestExpense",                                          "IS"),
    ("Interest Income",        "InterestIncome",                 "InterestIncome",                                           "IS"),
    ("Other Non-op Inc/(Exp)", "NonoperatingIncomeExpense",      "NonoperatingIncome",                                       "IS"),
    ("Pre-tax Income",         "PretaxIncomeLoss",               "IncomeLossFromContinuingOperationsBeforeIncomeTax",         "IS"),
    ("Income Tax",             "IncomeTaxes",                    "IncomeTaxExpense",                                         "IS"),
    ("Net Income",             "NetIncome",                      "NetIncomeLoss",                                            "IS"),
    ("Minority Interest",      None,                             "NetIncomeLossAttributableToNoncontrollingInterest",         "IS"),
    ("SBC",                    "StockBasedCompensationExpense",  "ShareBasedCompensation",                                   "CF"),
    ("Basic EPS",              None,                             "EarningsPerShareBasic",                                    "IS"),
    ("Diluted EPS",            None,                             "EarningsPerShareDiluted",                                  "IS"),
    ("Basic Shares",           "SharesAverage",                  "WeightedAverageNumberOfSharesOutstandingBasic",            "IS"),
    ("Diluted Shares",         "SharesFullyDilutedAverage",      "WeightedAverageNumberOfDilutedSharesOutstanding",          "IS"),
]


# ── Helpers ────────────────────────────────────────────────────────────────

def _col_to_quarter_label(col_name: str) -> str:
    """Convert edgartools period column name to FY label.

    Examples:
        "2023-03-31 (Q1)"  -> "FY2023Q1"
        "2024-12-31 (FY)"  -> "FY2024"
        "2023-03-31"       -> "2023-03-31"
    """
    m = re.match(r"(\d{4})-\d{2}-\d{2}\s+\((\w+)\)", col_name.strip())
    if m:
        year, period = m.group(1), m.group(2)
        return f"FY{year}" if period.upper() == "FY" else f"FY{year}{period}"
    return col_name


def _is_q_col(col_name: str) -> bool:
    """True if column is a quarterly period (Qx or FY), False for YTD."""
    m = re.search(r"\((\w+)\)", col_name)
    if not m:
        return False
    period = m.group(1).upper()
    return bool(re.match(r"Q\d+$", period)) or period == "FY"


def _current_q_col(df) -> str | None:
    """Return the first quarterly (non-YTD) period column from a filing's DataFrame."""
    for col in df.columns:
        if col in META_COLS:
            continue
        if _is_q_col(col):
            return col
    return None


def _consolidated_mask(df):
    """Boolean mask: non-abstract, non-breakdown, no dimension."""
    mask = ~df.get("abstract", False).astype(bool)
    mask &= ~df.get("is_breakdown", False).astype(bool)
    dim_col = df.get("dimension_member_label")
    if dim_col is not None:
        mask &= dim_col.isna() | (dim_col.astype(str) == "nan")
    return mask


def _match_is_row(df, std_concept: str | None, fallback_suffix: str) -> int | None:
    """Find the row index in df matching a template entry.

    Priority:
        1. standard_concept == std_concept (consolidated rows only)
        2. concept column contains fallback_suffix (case-insensitive, consolidated only)

    Returns None if no match found.
    """
    mask = _consolidated_mask(df)
    df_c = df[mask]

    if std_concept:
        rows = df_c[df_c["standard_concept"].astype(str) == std_concept]
        if not rows.empty:
            return rows.index[0]

    if fallback_suffix:
        rows = df_c[df_c["concept"].astype(str).str.contains(fallback_suffix, case=False, na=False)]
        if not rows.empty:
            return rows.index[0]

    return None


def _to_python_val(val) -> Any:
    """Convert pandas NA / float NaN / None to None; leave other values as-is."""
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    return val


def _row_key(df_row) -> str:
    """Unique key for a data row: concept + optional dimension_member_label."""
    concept = str(df_row.get("concept", "") or "")
    dim = str(df_row.get("dimension_member_label", "") or "")
    if dim and dim != "nan":
        return f"{concept}|{dim}"
    return concept


def _seg_sheet_suffix(concept: str, standard_concept: str | None) -> str:
    """Generate a ≤22-char alphanumeric suffix for a segment sheet name."""
    raw = standard_concept if standard_concept and standard_concept != "nan" else concept
    raw = re.sub(r"^[a-z_]+[_:]", "", raw)
    raw = re.sub(r"[^A-Za-z0-9]", "", raw)
    return raw[:22]


# ── IS: template-based fetch ────────────────────────────────────────────────

def _build_is_table(filings, max_filings: int) -> StatementTable:
    """Build Data_IS StatementTable from 10-Q filings using the fixed template."""
    periods: dict[str, tuple[str, dict[int, Any]]] = {}

    for filing in filings:
        if len(periods) >= max_filings:
            break
        try:
            tenq = filing.obj()
            stmt = tenq.financials.income_statement()
            if stmt is None:
                continue
            df = stmt.to_dataframe()
        except Exception as exc:
            print(f"[fetcher_gaap] IS warning: {exc!r}", file=sys.stderr)
            continue

        q_col = _current_q_col(df)
        if q_col is None:
            continue

        label = _col_to_quarter_label(q_col)
        if label in periods:
            continue

        # Fetch CF df for rows with source == "CF"
        cf_df: pd.DataFrame | None = None
        cf_q_col: str | None = None
        if any(row[3] == "CF" for row in IS_TEMPLATE):
            try:
                cf_stmt = tenq.financials.cashflow_statement()
                if cf_stmt is not None:
                    cf_df = cf_stmt.to_dataframe()
                    cf_q_col = _current_q_col(cf_df)
            except Exception:
                pass

        row_vals: dict[int, Any] = {}
        for i, (_, std_concept, fallback, source) in enumerate(IS_TEMPLATE):
            if source == "CF":
                if cf_df is not None and cf_q_col is not None:
                    idx = _match_is_row(cf_df, std_concept, fallback)
                    val = _to_python_val(cf_df.loc[idx, cf_q_col]) if idx is not None else None
                else:
                    val = None
            else:
                idx = _match_is_row(df, std_concept, fallback)
                val = _to_python_val(df.loc[idx, q_col]) if idx is not None else None
            row_vals[i] = val

        periods[label] = (str(filing.filing_date), row_vals)

    if not periods:
        return StatementTable(
            sheet_name="Data_IS",
            quarter_labels=[],
            filing_dates=[],
            concepts=[row[0] for row in IS_TEMPLATE],
            values=[[] for _ in IS_TEMPLATE],
        )

    sorted_labels = sorted(periods.keys())
    filing_dates = [periods[lbl][0] for lbl in sorted_labels]

    values: list[list[Any]] = []
    for i in range(len(IS_TEMPLATE)):
        values.append([periods[lbl][1].get(i) for lbl in sorted_labels])

    return StatementTable(
        sheet_name="Data_IS",
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=[row[0] for row in IS_TEMPLATE],
        values=values,
    )


# ── BS/CF: dynamic row-union fetch ─────────────────────────────────────────

def _build_dynamic_table(filings, stmt_method: str, sheet_name: str,
                          max_filings: int) -> StatementTable | None:
    """Build BS or CF StatementTable using a dynamic row union across all filings."""
    concept_labels: dict[str, str] = {}
    periods: dict[str, tuple[str, dict[str, Any]]] = {}

    for filing in filings:
        if len(periods) >= max_filings:
            break
        try:
            stmt = getattr(filing.obj().financials, stmt_method)()
            if stmt is None:
                continue
            df = stmt.to_dataframe()
        except Exception as exc:
            print(f"[fetcher_gaap] {sheet_name} warning: {exc!r}", file=sys.stderr)
            continue

        q_col = _current_q_col(df)
        if q_col is None:
            continue

        label = _col_to_quarter_label(q_col)
        if label in periods:
            continue

        mask = _consolidated_mask(df)
        df_c = df[mask].reset_index(drop=True)

        period_vals: dict[str, Any] = {}
        for _, row in df_c.iterrows():
            key = _row_key(row)
            if key not in concept_labels:
                concept_labels[key] = str(row.get("label", "") or key)
            period_vals[key] = _to_python_val(row.get(q_col))

        periods[label] = (str(filing.filing_date), period_vals)

    if not periods or not concept_labels:
        return None

    sorted_labels = sorted(periods.keys())
    filing_dates = [periods[lbl][0] for lbl in sorted_labels]
    concepts_ordered = list(concept_labels.keys())

    values: list[list[Any]] = []
    for key in concepts_ordered:
        values.append([periods[lbl][1].get(key) for lbl in sorted_labels])

    return StatementTable(
        sheet_name=sheet_name,
        quarter_labels=sorted_labels,
        filing_dates=filing_dates,
        concepts=[concept_labels[k] for k in concepts_ordered],
        values=values,
    )


# ── Segment breakdown sheets ────────────────────────────────────────────────

def _build_segment_tables(filings, max_filings: int) -> list[StatementTable]:
    """Build one StatementTable per IS concept that has segment/dimension rows."""
    seg_data: dict[str, dict] = {}
    periods_seen: set[str] = set()

    for filing in filings:
        if len(periods_seen) >= max_filings:
            break
        try:
            stmt = filing.obj().financials.income_statement()
            if stmt is None:
                continue
            df = stmt.to_dataframe()
        except Exception as exc:
            print(f"[fetcher_gaap] Seg warning: {exc!r}", file=sys.stderr)
            continue

        q_col = _current_q_col(df)
        if q_col is None:
            continue

        period_label = _col_to_quarter_label(q_col)
        filing_date = str(filing.filing_date)

        dim_col = df.get("dimension_member_label")
        if dim_col is None:
            continue
        mask_dim = ~(dim_col.isna() | (dim_col.astype(str) == "nan"))
        mask_not_abstract = ~df.get("abstract", False).astype(bool)
        df_dim = df[mask_dim & mask_not_abstract]

        for _, row in df_dim.iterrows():
            concept_xbrl = str(row.get("concept", "") or "")
            if not concept_xbrl:
                continue
            std = str(row.get("standard_concept", "") or "nan")
            member = str(row.get("dimension_member_label", "") or "")

            if concept_xbrl not in seg_data:
                seg_data[concept_xbrl] = {"std": std, "members": {}, "periods": {}}

            seg_data[concept_xbrl]["members"].setdefault(member, member)

            if period_label not in seg_data[concept_xbrl]["periods"]:
                seg_data[concept_xbrl]["periods"][period_label] = (filing_date, {})

            seg_data[concept_xbrl]["periods"][period_label][1][member] = _to_python_val(row.get(q_col))

        periods_seen.add(period_label)

    tables: list[StatementTable] = []
    used_sheet_names: set[str] = set()

    for concept_xbrl, data in seg_data.items():
        if not data["periods"]:
            continue

        suffix = _seg_sheet_suffix(concept_xbrl, data["std"] if data["std"] != "nan" else None)
        sheet_name = f"Data_Seg_{suffix}"
        base = sheet_name
        n = 2
        while sheet_name in used_sheet_names:
            sheet_name = f"{base[:28]}_{n}"
            n += 1
        used_sheet_names.add(sheet_name)

        sorted_periods = sorted(data["periods"].keys())
        members_ordered = list(data["members"].keys())

        tables.append(StatementTable(
            sheet_name=sheet_name,
            quarter_labels=sorted_periods,
            filing_dates=[data["periods"][lbl][0] for lbl in sorted_periods],
            concepts=members_ordered,
            values=[[data["periods"][lbl][1].get(m) for lbl in sorted_periods]
                    for m in members_ordered],
        ))

    return tables


# ── Meta sheet ─────────────────────────────────────────────────────────────

def _build_meta_table(ticker: str, company_name: str,
                       tables: list[StatementTable]) -> StatementTable:
    """Build Data_Meta sheet with filing summary info."""
    n_quarters = 0
    quarter_labels: list[str] = []
    filing_dates: list[str] = []
    for tbl in tables:
        if tbl.sheet_name != "Data_Meta" and tbl.quarter_labels:
            n_quarters    = len(tbl.quarter_labels)
            quarter_labels = tbl.quarter_labels
            filing_dates   = tbl.filing_dates
            break

    return StatementTable(
        sheet_name="Data_Meta",
        quarter_labels=quarter_labels,
        filing_dates=filing_dates,
        concepts=["Ticker", "Company Name", "Fetched Date", "Quarters Available"],
        values=[
            [ticker]            * n_quarters,
            [company_name]      * n_quarters,
            [str(date.today())] * n_quarters,
            [str(n_quarters)]   * n_quarters,
        ],
    )


# ── Public API ─────────────────────────────────────────────────────────────

def fetch_gaap_statements(ticker: str, identity: str,
                           max_filings: int = 80) -> list[StatementTable]:
    """Fetch quarterly GAAP statements from 10-Q filings for a ticker.

    Args:
        ticker:      Stock ticker, e.g. "AAPL"
        identity:    SEC EDGAR identity string
        max_filings: Maximum number of 10-Q filings to process (default 80, ~20 years)

    Returns:
        List of StatementTable: Data_IS, Data_BS, Data_CF, Data_Seg_*, Data_Meta

    Raises:
        ValueError: No 10-Q filings found for ticker
    """
    set_identity(identity)
    company = Company(ticker)

    filings = list(company.get_filings(form="10-Q", amendments=False))
    if not filings:
        raise ValueError(
            f"No 10-Q filings found for ticker '{ticker}'. "
            "The ticker may be invalid or the company may not file 10-Qs."
        )

    tables: list[StatementTable] = []
    tables.append(_build_is_table(filings, max_filings))

    for method_name, sheet_name in [("balance_sheet", "Data_BS"),
                                     ("cashflow_statement", "Data_CF")]:
        tbl = _build_dynamic_table(filings, method_name, sheet_name, max_filings)
        if tbl is not None:
            tables.append(tbl)

    tables.extend(_build_segment_tables(filings, max_filings))

    company_name = getattr(company, "name", ticker) or ticker
    tables.append(_build_meta_table(ticker, company_name, tables))

    for tbl in tables:
        tbl.ticker = ticker
    return tables
