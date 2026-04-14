import pytest
import openpyxl
from pathlib import Path
from fetcher_gaap import StatementTable
from excel_writer import write_statements


def test_a1_shows_ticker_when_set(tmp_path):
    tbl = StatementTable(
        sheet_name="Data_IS",
        quarter_labels=["FY2023Q1"],
        filing_dates=["2023-02-03"],
        concepts=["Revenue"],
        values=[[1000.0]],
        ticker="AAPL",
    )
    out = tmp_path / "AAPL.xlsx"
    write_statements([tbl], out)
    wb = openpyxl.load_workbook(out)
    assert wb["Data_IS"]["A1"].value == "AAPL"


def test_a1_is_none_when_ticker_empty(tmp_path):
    tbl = StatementTable(
        sheet_name="Data_IS",
        quarter_labels=["FY2023Q1"],
        filing_dates=["2023-02-03"],
        concepts=["Revenue"],
        values=[[1000.0]],
        ticker="",
    )
    out = tmp_path / "AAPL.xlsx"
    write_statements([tbl], out)
    wb = openpyxl.load_workbook(out)
    assert wb["Data_IS"]["A1"].value is None


@pytest.fixture
def sample_tables():
    return [
        StatementTable(
            sheet_name="Data_IS",
            quarter_labels=["FY2023Q1", "FY2023Q2", "FY2023Q3"],
            filing_dates=["2023-02-03", "2023-05-05", "2023-08-04"],
            concepts=["Revenues", "NetIncomeLoss", "EPS"],
            values=[
                [1000.0, 1100.0, 1200.0],
                [200.0,  220.0,  240.0],
                [1.23,   1.35,   1.47],
            ],
        ),
        StatementTable(
            sheet_name="Data_BS",
            quarter_labels=["FY2023Q1", "FY2023Q2"],
            filing_dates=["2023-02-03", "2023-05-05"],
            concepts=["Assets", "Liabilities"],
            values=[[50000.0, 52000.0], [30000.0, 31000.0]],
        ),
    ]


def test_write_creates_file(tmp_path, sample_tables):
    out = tmp_path / "AAPL.xlsx"
    write_statements(sample_tables, out)
    assert out.exists()


def test_write_creates_correct_sheets(tmp_path, sample_tables):
    out = tmp_path / "AAPL.xlsx"
    write_statements(sample_tables, out)
    wb = openpyxl.load_workbook(out)
    assert "Data_IS" in wb.sheetnames
    assert "Data_BS" in wb.sheetnames


def test_col_a_is_concept_name(tmp_path, sample_tables):
    out = tmp_path / "AAPL.xlsx"
    write_statements(sample_tables, out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Data_IS"]
    # Col A: row 1 and 2 are empty; row 3+ = concept names
    assert ws["A1"].value is None
    assert ws["A2"].value is None
    assert ws["A3"].value == "Revenues"
    assert ws["A4"].value == "NetIncomeLoss"
    assert ws["A5"].value == "EPS"


def test_row1_is_quarter_labels(tmp_path, sample_tables):
    out = tmp_path / "AAPL.xlsx"
    write_statements(sample_tables, out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Data_IS"]
    assert ws["B1"].value == "FY2023Q1"
    assert ws["C1"].value == "FY2023Q2"
    assert ws["D1"].value == "FY2023Q3"


def test_row2_is_filing_dates(tmp_path, sample_tables):
    out = tmp_path / "AAPL.xlsx"
    write_statements(sample_tables, out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Data_IS"]
    assert ws["B2"].value == "2023-02-03"
    assert ws["C2"].value == "2023-05-05"


def test_data_values_correct(tmp_path, sample_tables):
    out = tmp_path / "AAPL.xlsx"
    write_statements(sample_tables, out)
    wb = openpyxl.load_workbook(out)
    ws = wb["Data_IS"]
    # Revenues row: B3=1000, C3=1100, D3=1200
    assert ws["B3"].value == 1000.0
    assert ws["C3"].value == 1100.0
    assert ws["D3"].value == 1200.0


def test_preserves_non_data_sheets(tmp_path, sample_tables):
    """Python must NOT touch any sheet that doesn't start with Data_."""
    out = tmp_path / "AAPL.xlsx"
    # Pre-create workbook with user's custom sheet
    wb = openpyxl.Workbook()
    ws_user = wb.create_sheet("My_IS")
    ws_user["A1"] = "User annotation"
    ws_user["B1"] = "=Data_IS!B3"
    wb.save(out)
    wb.close()

    write_statements(sample_tables, out)

    wb2 = openpyxl.load_workbook(out)
    assert "My_IS" in wb2.sheetnames
    assert wb2["My_IS"]["A1"].value == "User annotation"


def test_rewrite_replaces_old_data(tmp_path, sample_tables):
    """Second write must replace all Data_* content (handles restatements)."""
    out = tmp_path / "AAPL.xlsx"
    write_statements(sample_tables, out)

    updated = [
        StatementTable(
            sheet_name="Data_IS",
            quarter_labels=["FY2023Q1", "FY2023Q2", "FY2023Q3", "FY2023Q4"],
            filing_dates=["2023-02-03", "2023-05-05", "2023-08-04", "2023-11-03"],
            concepts=["Revenues", "NetIncomeLoss", "EPS"],
            values=[
                [1000.0, 1100.0, 1200.0, 1300.0],
                [200.0,  220.0,  240.0,  260.0],
                [1.23,   1.35,   1.47,   1.60],
            ],
        )
    ]
    write_statements(updated, out)

    wb = openpyxl.load_workbook(out)
    ws = wb["Data_IS"]
    assert ws["E1"].value == "FY2023Q4"   # 4th quarter now present
    assert "Data_BS" not in wb.sheetnames  # not in updated list → removed
