import pytest
import pandas as pd
import os
from src.parser import ExcelParser

DATA_FILE = "data/complex_sample.xlsx"

@pytest.fixture
def parser():
    if not os.path.exists(DATA_FILE):
        pytest.fail(f"{DATA_FILE} not found. Run 'python scripts/generate_sample_excel.py' first.")
    return ExcelParser(DATA_FILE)

def test_sheet_names(parser):
    names = parser.get_sheet_names()
    assert "Sales Report" in names
    assert "Inventory" in names

def test_merged_cells_propagation(parser):
    # Testing Sales Report where "North" (A3:A4) is merged
    df = parser.parse_sheet("Sales Report", header_rows=2)

    # Check that "North" exists in both row 0 and row 1 of the DataFrame
    # (Corresponds to Excel rows 3 and 4)
    assert df.iloc[0, 0] == "North"
    assert df.iloc[1, 0] == "North"

def test_multi_row_headers(parser):
    df = parser.parse_sheet("Sales Report", header_rows=2)
    cols = df.columns

    assert "Q1 2023 - Sales" in cols
    assert "Q1 2023 - Profit" in cols
    # Check Region header (merged vertical)
    # The parser combines them. If both A1 and A2 have "Region" (due to unmerge), it becomes "Region - Region"
    # Or "Region" if logic handled dedupe (current logic doesn't, so we expect "Region - Region")
    assert "Region - Region" in cols

def test_simple_sheet(parser):
    df = parser.parse_sheet("Inventory", header_rows=1)
    assert len(df) == 3
    assert "Item ID" in df.columns
    assert df.iloc[0]["Name"] == "Widget A"

def test_missing_sheet(parser):
    with pytest.raises(ValueError):
        parser.parse_sheet("NonExistentSheet")
