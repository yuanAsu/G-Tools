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

def test_preprocessor_physical_removal(tmp_path):
    """Test removal of forbidden files and folders."""
    dummy_zip = tmp_path / "dummy_clean.xlsx"
    with zipfile.ZipFile(dummy_zip, 'w') as z:
        z.writestr("xl/calcChain.xml", "<dummy/>")
        z.writestr("xl/vbaProject.bin", "binaryjunk")
        z.writestr("xl/externalLinks/externalLink1.xml", "<dummy/>")
        z.writestr("xl/printerSettings/printerSettings1.bin", "binaryjunk")
        z.writestr("xl/workbook.xml", "<workbook/>")
        z.writestr("[Content_Types].xml", "<Types/>")

    processed_path = preprocess_excel(str(dummy_zip))

    with zipfile.ZipFile(processed_path, 'r') as z:
        files = z.namelist()
        assert "xl/calcChain.xml" not in files
        assert "xl/vbaProject.bin" not in files
        assert not any('externalLinks' in f for f in files)
        assert not any('printerSettings' in f for f in files)
        assert "xl/workbook.xml" in files

    os.remove(processed_path)

def test_preprocessor_global_namespace_transplant(tmp_path):
    """Test namespace replacement in multiple files."""
    dummy_zip = tmp_path / "dummy_global.xlsx"

    # Strict OOXML namespaces
    ns_ss = "http://purl.oclc.org/ooxml/spreadsheetml/main"
    ns_rel = "http://purl.oclc.org/ooxml/officeDocument/relationships"
    ns_draw = "http://purl.oclc.org/ooxml/drawingml/main"

    # Standard MS namespaces
    ms_ss = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ms_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ms_draw = "http://schemas.openxmlformats.org/drawingml/2006/main"

    content = f'<root xmlns="{ns_ss}" xmlns:r="{ns_rel}" xmlns:a="{ns_draw}"></root>'

    with zipfile.ZipFile(dummy_zip, 'w') as z:
        z.writestr("xl/workbook.xml", content)
        z.writestr("xl/worksheets/sheet1.xml", content)
        z.writestr("xl/drawing/drawing1.xml", content)

    processed_path = preprocess_excel(str(dummy_zip))

    with zipfile.ZipFile(processed_path, 'r') as z:
        for fname in ["xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/drawing/drawing1.xml"]:
            data = z.read(fname).decode('utf-8')
            assert ms_ss in data
            assert ms_rel in data
            assert ms_draw in data
            assert ns_ss not in data

    os.remove(processed_path)

def test_preprocessor_workbook_surgery(tmp_path):
    """Test removal of definedNames and externalReferences in workbook.xml only."""
    dummy_zip = tmp_path / "dummy_surgery.xlsx"

    # Case 1: Block tags with namespaces
    bad_content_block = '<workbook><x:definedNames>BAD</x:definedNames><x:externalReferences>BAD</x:externalReferences><sheets/></workbook>'

    # Case 2: Self-closing tags with namespaces
    bad_content_self = '<workbook><x:definedNames/><x:externalReferences/><sheets/></workbook>'

    for i, content in enumerate([bad_content_block, bad_content_self]):
        zip_name = dummy_zip.parent / f"dummy_surgery_{i}.xlsx"
        with zipfile.ZipFile(zip_name, 'w') as z:
            z.writestr("xl/workbook.xml", content)

        processed_path = preprocess_excel(str(zip_name))

        with zipfile.ZipFile(processed_path, 'r') as z:
            wb_content = z.read("xl/workbook.xml").decode('utf-8')
            assert "definedNames" not in wb_content
            assert "externalReferences" not in wb_content
            assert "<sheets/>" in wb_content

        os.remove(processed_path)

def test_preprocessor_rels_cleaning(tmp_path):
    """Test cleaning of .rels files."""
    dummy_zip = tmp_path / "dummy_rels.xlsx"

    # Old namespace for package relationships
    ns_pkg_rel_old = "http://purl.oclc.org/ooxml/package/relationships"
    ms_pkg_rel = "http://schemas.openxmlformats.org/package/2006/relationships"

    content = f'<Relationships xmlns="{ns_pkg_rel_old}"><Relationship Type="...externalLink..." Target="..."/></Relationships>'

    with zipfile.ZipFile(dummy_zip, 'w') as z:
        z.writestr("xl/workbook.xml.rels", content)

    processed_path = preprocess_excel(str(dummy_zip))

    with zipfile.ZipFile(processed_path, 'r') as z:
        rels_content = z.read("xl/workbook.xml.rels").decode('utf-8')
        assert ms_pkg_rel in rels_content
        assert ns_pkg_rel_old not in rels_content
        assert "externalLink" not in rels_content

    os.remove(processed_path)
