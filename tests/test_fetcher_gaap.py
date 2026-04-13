import pytest
from unittest.mock import patch, MagicMock
import pandas as pd
from fetcher_gaap import fetch_gaap_statements, StatementTable, _parse_fiscal_label, _col_to_quarter_label


def test_parse_fiscal_label_quarter():
    assert _parse_fiscal_label("2024", "Q1") == "FY2024Q1"
    assert _parse_fiscal_label("2024", "Q4") == "FY2024Q4"


def test_parse_fiscal_label_annual():
    assert _parse_fiscal_label("2024", "FY") == "FY2024"


def test_col_to_quarter_label_quarter():
    assert _col_to_quarter_label("2023-03-31 (Q1)") == "FY2023Q1"
    assert _col_to_quarter_label("2023-06-30 (Q2)") == "FY2023Q2"


def test_col_to_quarter_label_annual():
    assert _col_to_quarter_label("2024-12-31 (FY)") == "FY2024"


def test_col_to_quarter_label_instant_fallback():
    assert _col_to_quarter_label("2023-03-31") == "2023-03-31"


def test_statement_table_structure():
    tbl = StatementTable(
        sheet_name="Data_IS",
        quarter_labels=["FY2023Q1", "FY2023Q2"],
        filing_dates=["2023-02-03", "2023-05-05"],
        concepts=["Revenues", "NetIncomeLoss"],
        values=[[1000, 1100], [200, 220]],
    )
    assert tbl.sheet_name == "Data_IS"
    assert len(tbl.quarter_labels) == 2
    assert len(tbl.values) == 2       # 2 concepts
    assert len(tbl.values[0]) == 2    # 2 quarters


def test_fetch_returns_list_of_statement_tables(mock_edgar_company):
    with patch("fetcher_gaap.Company") as MockCompany, \
         patch("fetcher_gaap.set_identity"):
        MockCompany.return_value = mock_edgar_company
        result = fetch_gaap_statements("AAPL", identity="Test User test@test.com")
    assert isinstance(result, list)
    assert all(isinstance(t, StatementTable) for t in result)


def test_fetch_includes_income_statement(mock_edgar_company):
    with patch("fetcher_gaap.Company") as MockCompany, \
         patch("fetcher_gaap.set_identity"):
        MockCompany.return_value = mock_edgar_company
        result = fetch_gaap_statements("AAPL", identity="Test User test@test.com")
    sheet_names = [t.sheet_name for t in result]
    assert "Data_IS" in sheet_names


def test_fetch_includes_meta(mock_edgar_company):
    with patch("fetcher_gaap.Company") as MockCompany, \
         patch("fetcher_gaap.set_identity"):
        MockCompany.return_value = mock_edgar_company
        result = fetch_gaap_statements("AAPL", identity="Test User test@test.com")
    sheet_names = [t.sheet_name for t in result]
    assert "Data_Meta" in sheet_names


def test_fetch_consistent_lengths(mock_edgar_company):
    with patch("fetcher_gaap.Company") as MockCompany, \
         patch("fetcher_gaap.set_identity"):
        MockCompany.return_value = mock_edgar_company
        result = fetch_gaap_statements("AAPL", identity="Test User test@test.com")
    for tbl in result:
        if tbl.sheet_name == "Data_Meta":
            continue  # Meta uses a different layout
        n_quarters = len(tbl.quarter_labels)
        assert len(tbl.filing_dates) == n_quarters
        for row in tbl.values:
            assert len(row) == n_quarters


def test_fetch_quarter_labels_format(mock_edgar_company):
    with patch("fetcher_gaap.Company") as MockCompany, \
         patch("fetcher_gaap.set_identity"):
        MockCompany.return_value = mock_edgar_company
        result = fetch_gaap_statements("AAPL", identity="Test User test@test.com")
    is_table = next(t for t in result if t.sheet_name == "Data_IS")
    assert is_table.quarter_labels == ["FY2023Q1", "FY2023Q2"]


@pytest.fixture
def mock_edgar_company():
    """Minimal mock of edgartools Company with income statement data.

    Mimics the real edgartools v5.29 Statement.to_dataframe() output:
    - Flat DataFrame with RangeIndex (NOT concept names as index)
    - 'concept' column: XBRL concept name
    - 'label'   column: human-readable name
    - 'abstract' column: bool (True = section header, no values)
    - 'level'   column: int
    - Period columns named like "2023-03-31 (Q1)", "2023-06-30 (Q2)"
    """
    mock_df = pd.DataFrame({
        "concept":  ["us-gaap_Revenues", "us-gaap_NetIncomeLoss"],
        "label":    ["Revenues", "Net Income"],
        "abstract": [False, False],
        "level":    [1, 1],
        "2023-03-31 (Q1)": [1000.0, 200.0],
        "2023-06-30 (Q2)": [1100.0, 220.0],
    })

    mock_stmt = MagicMock()
    mock_stmt.to_dataframe.return_value = mock_df

    mock_financials = MagicMock()
    mock_financials.income_statement.return_value = mock_stmt
    mock_financials.balance_sheet.return_value = None
    mock_financials.cashflow_statement.return_value = None
    mock_financials.statement_of_equity.return_value = None
    mock_financials.comprehensive_income.return_value = None

    mock_company = MagicMock()
    mock_company.name = "Apple Inc."
    mock_company.get_financials.return_value = mock_financials
    return mock_company
