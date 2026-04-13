"""
fetcher_gaap.py — Fetch XBRL GAAP financial statements from SEC EDGAR via edgartools.

Public API:
    fetch_gaap_statements(ticker, identity) -> list[StatementTable]

Each StatementTable is pre-structured for direct writing by excel_writer.py.
"""

import sys
from dataclasses import dataclass
from datetime import date
from typing import Any

from edgar import Company, set_identity as set_identity


# ---- Data contract ----

@dataclass
class StatementTable:
    """One financial statement pre-structured for Excel output."""
    sheet_name: str              # e.g. "Data_IS"
    quarter_labels: list[str]    # Row 1: ["FY2023Q1", ...]
    filing_dates: list[str]      # Row 2: ["2023-02-03", ...]
    concepts: list[str]          # Col A, Row 3+
    values: list[list[Any]]      # values[concept_idx][quarter_idx]


# ---- Statement type map ----

STATEMENT_MAP = [
    ("income_statement",     "Data_IS"),
    ("balance_sheet",        "Data_BS"),
    ("cashflow_statement",   "Data_CF"),
    ("statement_of_equity",  "Data_Equity"),
    ("comprehensive_income", "Data_CI"),
]


# ---- Helpers ----

def _parse_fiscal_label(fiscal_year: str, fiscal_period: str) -> str:
    """Build label like 'FY2024Q1' or 'FY2024' from edgartools period info."""
    if str(fiscal_period).upper() == "FY":
        return f"FY{fiscal_year}"
    return f"FY{fiscal_year}{fiscal_period}"


def _col_to_quarter_label(col_name: str) -> str:
    """Parse edgartools period column name to FY quarter label format.

    Examples:
        "2023-03-31 (Q1)"  -> "FY2023Q1"
        "2024-12-31 (FY)"  -> "FY2024"
        "2023-03-31"       -> "2023-03-31"  (instant, no parens — return as-is)
    """
    import re
    m = re.match(r"(\d{4})-\d{2}-\d{2}\s+\((\w+)\)", col_name.strip())
    if m:
        year, period = m.group(1), m.group(2)
        if period.upper() == "FY":
            return f"FY{year}"
        return f"FY{year}{period}"
    # No parentheses — instant column, return as-is
    return col_name


def _stmt_to_table(stmt, sheet_name: str) -> StatementTable | None:
    """Convert an edgartools Statement to a StatementTable.

    Returns None if the statement has no data.

    edgartools v5.29 Statement.to_dataframe() returns a flat DataFrame where:
      - 'label'    column: human-readable concept name (e.g. "Revenues")
      - 'concept'  column: XBRL concept name (e.g. "us-gaap_Revenues")
      - 'abstract' column: bool, True for section headers with no values
      - 'level'    column: indentation level
      - period columns: named like "2024-03-31", "2024-03-31 (Q1)", "2024-03-31 (FY)"
      - Index is a plain RangeIndex (NOT concept names)
    """
    if stmt is None:
        return None
    try:
        df = stmt.to_dataframe()
    except Exception as exc:
        print(f"[fetcher_gaap] WARNING: to_dataframe() raised {exc!r}", file=sys.stderr)
        return None
    if df is None or df.empty:
        return None

    # Identify metadata columns — everything else is a period column
    META_COLS = {
        'concept', 'label', 'standard_concept', 'level', 'abstract',
        'dimension', 'is_breakdown', 'dimension_axis', 'dimension_member',
        'dimension_member_label', 'dimension_label', 'unit', 'point_in_time',
        'balance', 'weight', 'preferred_sign',
    }
    period_cols = [c for c in df.columns if c not in META_COLS]

    if not period_cols:
        return None

    # Drop abstract rows (section headers have no numeric values)
    if 'abstract' in df.columns:
        df = df[~df['abstract'].astype(bool)].reset_index(drop=True)

    if df.empty:
        return None

    # Quarter labels: parse period column names to FY format (e.g. "FY2024Q1")
    quarter_labels = [_col_to_quarter_label(str(c)) for c in period_cols]
    # No filing-date metadata on Statement objects — leave blank
    filing_dates = [""] * len(period_cols)

    # Concept labels: prefer 'label' column, fall back to 'concept'
    if 'label' in df.columns:
        concepts = list(df['label'].fillna("").astype(str))
    elif 'concept' in df.columns:
        concepts = list(df['concept'].fillna("").astype(str))
    else:
        concepts = [str(i) for i in df.index]

    # Values: one list per concept, one value per period
    values = [list(df[period_cols].iloc[i].values) for i in range(len(df))]

    return StatementTable(
        sheet_name=sheet_name,
        quarter_labels=quarter_labels,
        filing_dates=filing_dates,
        concepts=concepts,
        values=values,
    )


def _build_meta_table(ticker: str, company_name: str, tables: list[StatementTable]) -> StatementTable:
    """Build a Data_Meta sheet summarising filing info across all statements."""
    # Count unique quarters from the first available statement
    n_quarters = 0
    quarter_labels: list[str] = []
    filing_dates: list[str] = []
    for tbl in tables:
        if tbl.sheet_name != "Data_Meta" and tbl.quarter_labels:
            n_quarters    = len(tbl.quarter_labels)
            quarter_labels = tbl.quarter_labels
            filing_dates   = tbl.filing_dates
            break

    meta_concepts = ["Ticker", "Company Name", "Fetched Date", "Quarters Available"]
    meta_values = [
        [ticker] * n_quarters,
        [company_name] * n_quarters,
        [str(date.today())] * n_quarters,
        [str(n_quarters)] * n_quarters,
    ]

    return StatementTable(
        sheet_name="Data_Meta",
        quarter_labels=quarter_labels,
        filing_dates=filing_dates,
        concepts=meta_concepts,
        values=meta_values,
    )


# ---- Public API ----

def fetch_gaap_statements(ticker: str, identity: str) -> list[StatementTable]:
    """Fetch all available historical quarterly GAAP statements for a ticker.

    Args:
        ticker:   Stock ticker, e.g. "AAPL"
        identity: SEC EDGAR identity string, e.g. "John Smith john@example.com"

    Returns:
        List of StatementTable objects (IS, BS, CF, Equity, CI, Meta).
        Skips statement types not available for this company.

    Raises:
        ValueError: ticker not found on EDGAR
        Exception:  propagates network/API errors from edgartools
    """
    set_identity(identity)
    company    = Company(ticker)
    financials = company.get_financials()

    if financials is None:
        raise ValueError(
            f"No annual filing found for ticker '{ticker}'. "
            "The company may not have an annual 10-K/20-F/40-F on EDGAR, "
            "or the ticker may be invalid."
        )

    tables: list[StatementTable] = []
    for method_name, sheet_name in STATEMENT_MAP:
        method = getattr(financials, method_name, None)
        if method is None:
            continue
        try:
            stmt = method()
        except Exception:
            continue
        tbl = _stmt_to_table(stmt, sheet_name)
        if tbl is not None:
            tables.append(tbl)

    company_name = getattr(company, "name", ticker) or ticker
    meta = _build_meta_table(ticker, company_name, tables)
    tables.append(meta)

    return tables
