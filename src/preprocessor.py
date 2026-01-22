import zipfile
import re
import os
import tempfile

def preprocess_excel(input_path: str) -> str:
    """
    Preprocesses an Excel file to handle WPS/Strict OOXML issues via Global Gene Repair.
    Returns the path to the processed temporary file.
    """
    print(f"正在进行全域基因修复: {input_path} ...")

    # Create a temp file for the output
    fd, output_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)

    try:
        with zipfile.ZipFile(input_path, 'r') as zin, zipfile.ZipFile(output_path, 'w') as zout:
            for item in zin.infolist():
                filename = item.filename

                # 1. Physical Removal
                # User logic: if 'externalLinks' in item.filename: continue
                if 'calcChain' in filename: continue
                if 'vbaProject' in filename: continue
                if 'externalLinks' in filename: continue
                if 'printerSettings' in filename: continue

                content = zin.read(filename)

                # Check if it is an XML-like file that needs processing
                if filename.endswith('.xml') and (
                    filename.endswith('workbook.xml') or
                    'worksheets/sheet' in filename or
                    'sharedStrings.xml' in filename or
                    'styles.xml' in filename or
                    'drawing' in filename # Covered by 'endswith .xml' in my logic usually, but let's be safe
                ):
                    content_str = content.decode('utf-8')

                    # 2. Global Namespace Transplant
                    # SpreadsheetML Main
                    content_str = content_str.replace(
                        "http://purl.oclc.org/ooxml/spreadsheetml/main",
                        "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    )
                    # Relationships
                    content_str = content_str.replace(
                        "http://purl.oclc.org/ooxml/officeDocument/relationships",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    )
                    # DrawingML
                    content_str = content_str.replace(
                        "http://purl.oclc.org/ooxml/drawingml/main",
                        "http://schemas.openxmlformats.org/drawingml/2006/main"
                    )

                    # 3. Workbook Surgery (Specific to workbook.xml)
                    if filename.endswith('workbook.xml'):
                        # Aggressive Regex for definedNames (handles namespaces and self-closing)
                        # Remove block <...definedNames...>...</...definedNames>
                        content_str = re.sub(r'<[\w:]*?definedNames.*?>.*?</[\w:]*?definedNames>', '', content_str, flags=re.DOTALL)
                        # Remove self-closing <...definedNames.../>
                        content_str = re.sub(r'<[\w:]*?definedNames.*?/>', '', content_str)

                        # Aggressive Regex for externalReferences
                        content_str = re.sub(r'<[\w:]*?externalReferences.*?>.*?</[\w:]*?externalReferences>', '', content_str, flags=re.DOTALL)
                        content_str = re.sub(r'<[\w:]*?externalReferences.*?/>', '', content_str)

                    content = content_str.encode('utf-8')

                # 3. .rels Cleaning
                elif filename.endswith('.rels'):
                    content_str = content.decode('utf-8')

                    # Replace package relationships namespace
                    content_str = content_str.replace(
                        'http://purl.oclc.org/ooxml/package/relationships',
                        'http://schemas.openxmlformats.org/package/2006/relationships'
                    )
                    # Also do the standard ones just in case
                    content_str = content_str.replace(
                        "http://purl.oclc.org/ooxml/officeDocument/relationships",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    )

                    # Remove externalLink relationships
                    if 'workbook.xml.rels' in filename:
                        content_str = re.sub(r'<Relationship[^>]*?externalLink[^>]*?/>', '', content_str)
                        # Also handle non-self-closing if any (rare for Relationships)

                    content = content_str.encode('utf-8')

                zout.writestr(item, content)

        print("全域修复完毕")
        return output_path

    except Exception as e:
        # If failure, clean up and raise
        if os.path.exists(output_path):
            os.remove(output_path)
        raise e
