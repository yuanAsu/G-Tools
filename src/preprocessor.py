import zipfile
import re
import os
import tempfile

def preprocess_excel(input_path: str) -> str:
    """
    Preprocesses an Excel file to handle WPS/Strict OOXML issues via Global Gene Repair.
    Returns the path to the processed temporary file.
    """

    # Create a temp file for the output
    fd, output_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)

    try:
        with zipfile.ZipFile(input_path, 'r') as zin, zipfile.ZipFile(output_path, 'w') as zout:
            for item in zin.infolist():
                filename = item.filename

                # 1. Physical Removal
                # Direct file matches
                if filename in ["xl/calcChain.xml", "xl/vbaProject.bin"]:
                    continue
                # Directory matches (prefixes)
                if filename.startswith("xl/externalLinks/") or filename.startswith("xl/printerSettings/"):
                    continue

                content = zin.read(filename)

                # Check if it is an XML-like file that needs processing
                # Applying Global Namespace Transplant to all XML/rels files
                is_xml = filename.endswith(".xml") or filename.endswith(".rels")

                if is_xml:
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
                    if filename == "xl/workbook.xml":
                        # Remove <definedNames>...</definedNames>
                        content_str = re.sub(r"<definedNames>.*?</definedNames>", "", content_str, flags=re.DOTALL)
                        # Remove <externalReferences>...</externalReferences>
                        content_str = re.sub(r"<externalReferences>.*?</externalReferences>", "", content_str, flags=re.DOTALL)

                        # Note: Previous implementation removed customWorkbookViews, keeping it minimal as per strict request
                        # but often definedNames removal is the critical one for crashes.

                    content = content_str.encode('utf-8')

                zout.writestr(item, content)

        return output_path

    except Exception as e:
        # If failure, clean up and raise
        if os.path.exists(output_path):
            os.remove(output_path)
        raise e
