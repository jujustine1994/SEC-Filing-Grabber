"""
fetcher_gaap.py — Fetch XBRL GAAP financial statements from SEC EDGAR via edgartools.

Public API:
    fetch_gaap_statements(ticker, identity) -> list[StatementTable]

Each StatementTable is pre-structured for direct writing by excel_writer.py.
"""

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


def _stmt_to_table(stmt, sheet_name: str) -> StatementTable | None:
    """Convert an edgartools Statement to a StatementTable.

    Returns None if the statement has no data.

    edgartools Statement objects expose:
      - .to_dataframe() -> DataFrame (concepts as index, period labels as columns)
      - .periods        -> list of period objects with .fiscal_year, .fiscal_period, .filed
    """
    if stmt is None:
        return None
    try:
        df = stmt.to_dataframe()
    except Exception:
        return None
    if df is None or df.empty:
        return None

    # Build quarter labels and filing dates from period metadata
    try:
        periods = stmt.periods
        quarter_labels = [
            _parse_fiscal_label(str(p.fiscal_year), str(p.fiscal_period))
            for p in periods
        ]
        filing_dates = [str(p.filed) if p.filed else "" for p in periods]
    except AttributeError:
        # Fallback: use DataFrame column headers directly
        quarter_labels = [str(c) for c in df.columns]
        filing_dates   = [""] * len(df.columns)

    concepts = list(df.index.astype(str))
    values   = [list(df.iloc[i].values) for i in range(len(df))]

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
