import openpyxl
import os
from openpyxl import Workbook
from src.preprocessor import preprocess_excel

class ExcelLoaderError(Exception):
    """Custom exception for Excel loading failures."""
    pass

def load_complex_excel(filepath: str) -> Workbook:
    """
    Loads an Excel file after running the strict pre-processing protocol.
    Acts as a 'Gatekeeper' to ensure only clean files are loaded.
    """
    if not os.path.exists(filepath):
        raise ExcelLoaderError(f"File not found: {filepath}")

    processed_path = None
    try:
        # 1. Pre-process (Physical Removal, Namespace Transplant, Surgery)
        processed_path = preprocess_excel(filepath)

        # 2. Load with OpenPyXL (Strict flags)
        # read_only=False: Force dimension recalc
        # keep_links=False: Ignore broken external links
        wb = openpyxl.load_workbook(
            processed_path,
            data_only=True,
            read_only=False,
            keep_links=False
        )
        return wb

    except Exception as e:
        raise ExcelLoaderError(f"文件格式受损或加密，请另存为标准 .xlsx 后重试。Error: {str(e)}")

    finally:
        # Clean up the temp file
        if processed_path and os.path.exists(processed_path):
            try:
                os.remove(processed_path)
            except OSError:
                pass
