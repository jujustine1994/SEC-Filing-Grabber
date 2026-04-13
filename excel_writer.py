"""
excel_writer.py — Write StatementTable objects to Data_* sheets in an Excel file.

Rules:
- Only writes/deletes sheets with the Data_* prefix
- Never modifies sheets with other prefixes (My_*, Analysis, etc.)
- Full rewrite on every call — ensures restatements are captured
- Col A: Concept/Label; Col B+: quarters oldest→newest
- Row 1: quarter labels; Row 2: filing dates; Row 3+: financial data
"""

from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from fetcher_gaap import StatementTable


def write_statements(tables: list[StatementTable], output_path: str | Path) -> None:
    """Write StatementTable list to an Excel file.

    Replaces all existing Data_* sheets. Preserves all other sheets.
    Creates the file (and parent directory) if absent.

    Args:
        tables:      List of StatementTable objects to write.
        output_path: Path to .xlsx file.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if output_path.exists():
        wb = load_workbook(output_path)
    else:
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    # Remove all existing Data_* sheets (full rewrite)
    for name in list(wb.sheetnames):
        if name.startswith("Data_"):
            del wb[name]

    # Write each StatementTable as a new Data_* sheet
    for tbl in tables:
        ws = wb.create_sheet(tbl.sheet_name)
        _write_sheet(ws, tbl)

    try:
        wb.save(output_path)
    finally:
        wb.close()


def _write_sheet(ws: Worksheet, tbl: StatementTable) -> None:
    """Write one StatementTable into a worksheet.

    Layout:
        Row 1: A=empty, B..=quarter labels
        Row 2: A=empty, B..=filing dates
        Row 3+: A=concept name, B..=values
    """
    # Row 1: quarter labels (starting at column B)
    ws.cell(row=1, column=1, value=None)
    for col_idx, label in enumerate(tbl.quarter_labels, start=2):
        ws.cell(row=1, column=col_idx, value=label)

    # Row 2: filing dates (starting at column B)
    ws.cell(row=2, column=1, value=None)
    for col_idx, date_str in enumerate(tbl.filing_dates, start=2):
        ws.cell(row=2, column=col_idx, value=date_str)

    # Row 3+: concept name (col A) + values (col B+)
    for row_offset, (concept, row_values) in enumerate(zip(tbl.concepts, tbl.values)):
        row = 3 + row_offset
        ws.cell(row=row, column=1, value=concept)
        for col_idx, val in enumerate(row_values, start=2):
            ws.cell(row=row, column=col_idx, value=val)
