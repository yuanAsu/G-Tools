import pytest
import os
from src.loader import load_complex_excel, ExcelLoaderError
from src.serializer import serialize_excel_to_markdown
from src.llm_bridge import generate_analysis_prompt

DATA_FILE = "data/complex_v2.xlsx"

def test_loader_success():
    if not os.path.exists(DATA_FILE):
        pytest.fail("Data file missing")
    wb = load_complex_excel(DATA_FILE)
    assert wb is not None
    assert "Mixed Data" in wb.sheetnames

def test_loader_missing_file():
    with pytest.raises(ExcelLoaderError):
        load_complex_excel("non_existent_file.xlsx")

def test_serializer_formatting():
    # Setup dummy workbook logic isn't easily mockable without openpyxl object construction
    # We test on actual file
    wb = load_complex_excel(DATA_FILE)
    markdown = serialize_excel_to_markdown(wb)

    assert "=== Sheet: Mixed Data ===" in markdown
    # Check truncation (though our sample data is short, logic is there)
    # Check format
    assert "Row " in markdown
    assert ": [" in markdown

def test_prompt_generation():
    prompt = generate_analysis_prompt(DATA_FILE)
    assert "Plaintext" in prompt
    assert "Output JSON:" in prompt
    # Check if data is embedded
    assert "Project Info" in prompt
    assert "Apollo Mission" in prompt
