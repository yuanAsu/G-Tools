import pytest
import os
import zipfile
from src.preprocessor import preprocess_excel
from src.parser import ExcelParser

DATA_FILE = "data/complex_v2.xlsx"

def test_preprocessor_runs_on_valid_file():
    """Test that preprocessing a valid file doesn't break it."""
    if not os.path.exists(DATA_FILE):
        pytest.fail("Data file missing")

    processed_path = preprocess_excel(DATA_FILE)
    assert os.path.exists(processed_path)
    assert processed_path != DATA_FILE

    # Check if it's a valid zip
    assert zipfile.is_zipfile(processed_path)

    # Check if we can load it
    try:
        parser = ExcelParser(DATA_FILE) # This now calls preprocessor internally
        assert len(parser.get_sheet_names()) > 0
    finally:
        if os.path.exists(processed_path):
            os.remove(processed_path)

def test_preprocessor_removes_calc_chain(tmp_path):
    """Mock a zip with calcChain.xml and ensure it is removed."""
    # Create a dummy zip
    dummy_zip = tmp_path / "dummy.xlsx"
    with zipfile.ZipFile(dummy_zip, 'w') as z:
        z.writestr("xl/calcChain.xml", "<dummy></dummy>")
        z.writestr("xl/workbook.xml", "<workbook></workbook>")
        z.writestr("[Content_Types].xml", "<Types></Types>") # Minimal valid structure

    processed_path = preprocess_excel(str(dummy_zip))

    with zipfile.ZipFile(processed_path, 'r') as z:
        files = z.namelist()
        assert "xl/calcChain.xml" not in files
        assert "xl/workbook.xml" in files

    os.remove(processed_path)

def test_preprocessor_namespace_replacement(tmp_path):
    """Test namespace replacement in workbook.xml."""
    dummy_zip = tmp_path / "dummy_ns.xlsx"
    old_ns = "http://purl.oclc.org/ooxml/spreadsheetml/main"
    content = f'<workbook xmlns="{old_ns}"><sheets/></workbook>'

    with zipfile.ZipFile(dummy_zip, 'w') as z:
        z.writestr("xl/workbook.xml", content)

    processed_path = preprocess_excel(str(dummy_zip))

    with zipfile.ZipFile(processed_path, 'r') as z:
        new_content = z.read("xl/workbook.xml").decode('utf-8')
        assert "http://schemas.openxmlformats.org/spreadsheetml/2006/main" in new_content
        assert old_ns not in new_content

    os.remove(processed_path)

def test_preprocessor_metadata_surgery(tmp_path):
    """Test removal of definedNames."""
    dummy_zip = tmp_path / "dummy_meta.xlsx"
    content = '<workbook><definedNames><definedName>Test</definedName></definedNames><sheets/></workbook>'

    with zipfile.ZipFile(dummy_zip, 'w') as z:
        z.writestr("xl/workbook.xml", content)

    processed_path = preprocess_excel(str(dummy_zip))

    with zipfile.ZipFile(processed_path, 'r') as z:
        new_content = z.read("xl/workbook.xml").decode('utf-8')
        assert "<definedNames>" not in new_content
        assert "<sheets/>" in new_content

    os.remove(processed_path)

def test_preprocessor_relationship_pruning(tmp_path):
    """Test removal of externalLink relationships."""
    dummy_zip = tmp_path / "dummy_rels.xlsx"
    content = '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLink1.xml"/>
</Relationships>'''

    with zipfile.ZipFile(dummy_zip, 'w') as z:
        z.writestr("xl/_rels/workbook.xml.rels", content)
        z.writestr("xl/workbook.xml", "<workbook></workbook>") # Needed for loop

    processed_path = preprocess_excel(str(dummy_zip))

    with zipfile.ZipFile(processed_path, 'r') as z:
        new_content = z.read("xl/_rels/workbook.xml.rels").decode('utf-8')
        assert 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"' in new_content
        assert 'externalLink' not in new_content

    os.remove(processed_path)
