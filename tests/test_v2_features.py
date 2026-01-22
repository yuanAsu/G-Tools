import pytest
import os
from src.parser import ExcelParser

DATA_FILE = "data/complex_v2.xlsx"

@pytest.fixture
def parser():
    if not os.path.exists(DATA_FILE):
        pytest.fail(f"{DATA_FILE} not found. Run scripts/generate_v2_data.py first.")
    return ExcelParser(DATA_FILE)

def test_anchor_finding(parser):
    # Test Mixed Data sheet
    row, col = parser.find_anchor("Mixed Data", "Project Info")
    assert row == 1
    assert col == 1

    row, col = parser.find_anchor("Mixed Data", "Funding Assumptions")
    assert row == 7
    assert col == 1

def test_anchor_table_extraction(parser):
    # Parse Project Info
    df1 = parser.parse_table_at_anchor("Mixed Data", "Project Info", relative_header_offset=1)
    assert not df1.empty
    assert "ID" in df1.columns
    assert "Name" in df1.columns
    assert df1.iloc[0]["Name"] == "Alpha"
    # Should stop before Funding Assumptions
    assert "Funding Assumptions" not in df1.iloc[:, 0].values

    # Parse Funding Assumptions
    df2 = parser.parse_table_at_anchor("Mixed Data", "Funding Assumptions", relative_header_offset=1)
    assert not df2.empty
    assert "Source" in df2.columns
    assert df2.iloc[0]["Source"] == "Bank Loan"

def test_form_parsing(parser):
    data = parser.parse_form("Project Form")
    assert data["Project Name"] == "Apollo Mission"
    assert data["Project Manager"] == "Jules"
    assert data["Total Budget"] == 1500000
    assert data["Start Date"] == "2024-01-01"

def test_hierarchy_parsing(parser):
    tree = parser.parse_hierarchy("Financial Statement")

    assert "Assets" in tree
    assert "Current Assets" in tree["Assets"]
    assert tree["Assets"]["Current Assets"]["Cash"] == 5000

    assert "Fixed Assets" in tree
    assert tree["Fixed Assets"]["Equipment"] == 10000
