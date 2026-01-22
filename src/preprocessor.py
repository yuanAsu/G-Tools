import zipfile
import re
import os
import tempfile
import shutil

def preprocess_excel(input_path: str) -> str:
    """
    Preprocesses an Excel file to handle WPS/non-standard issues.
    Returns the path to the processed temporary file.
    """

    # Create a temp file for the output
    fd, output_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)

    try:
        with zipfile.ZipFile(input_path, 'r') as zin, zipfile.ZipFile(output_path, 'w') as zout:
            for item in zin.infolist():
                filename = item.filename

                # 1. Physical Removal: calcChain.xml
                if filename == "xl/calcChain.xml":
                    continue

                content = zin.read(filename)

                if filename == "xl/workbook.xml":
                    content_str = content.decode('utf-8')

                    # 2. Namespace Normalization
                    content_str = content_str.replace(
                        "http://purl.oclc.org/ooxml/spreadsheetml/main",
                        "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    )
                    content_str = content_str.replace(
                        "http://purl.oclc.org/ooxml/officeDocument/relationships",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    )

                    # 3. Metadata Surgery
                    # Using regex to remove blocks. DOTALL needed for multiline match if any.
                    # Patterns are non-greedy .*? to match start to end tag.
                    patterns = [
                        r"<definedNames>.*?</definedNames>",
                        r"<externalReferences>.*?</externalReferences>",
                        r"<customWorkbookViews>.*?</customWorkbookViews>"
                    ]

                    for p in patterns:
                        content_str = re.sub(p, "", content_str, flags=re.DOTALL)

                    content = content_str.encode('utf-8')

                elif filename == "xl/_rels/workbook.xml.rels":
                    content_str = content.decode('utf-8')

                    # 2. Namespace Normalization (rels specific)
                    # Note: Usually relationships namespace is standard, but user asked for strict replacement.
                    # The user instruction said: "Read xl/workbook.xml and xl/_rels/workbook.xml.rels... perform string replacement"
                    content_str = content_str.replace(
                        "http://purl.oclc.org/ooxml/spreadsheetml/main",
                        "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    )
                    content_str = content_str.replace(
                        "http://purl.oclc.org/ooxml/officeDocument/relationships",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    )

                    # 4. Relationship Pruning
                    # Remove <Relationship ... Type="...externalLink..." ... /> or Target="...externalLink..."
                    # We can use regex.
                    # Pattern: <Relationship [^>]*externalLink[^>]*/>
                    # Or specific attribute check.
                    # Let's simple regex for any Relationship tag containing "externalLink"

                    # Regex for self-closing Relationship tag
                    # <Relationship ... />

                    def rel_replacer(match):
                        tag = match.group(0)
                        if "externalLink" in tag:
                            return ""
                        return tag

                    # Match <Relationship ...> or <Relationship ... />
                    # Note: Relationships usually don't have body content, they are self-closing.
                    content_str = re.sub(r"<Relationship\s+.*?>", rel_replacer, content_str)

                    content = content_str.encode('utf-8')

                zout.writestr(item, content)

        return output_path

    except Exception as e:
        # If failure, clean up and raise
        if os.path.exists(output_path):
            os.remove(output_path)
        raise e
