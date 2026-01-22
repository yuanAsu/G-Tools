import openpyxl
import os
import warnings
from openpyxl import Workbook
from src.preprocessor import preprocess_excel

# Suppress OpenPyXL warnings as requested
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

class ExcelLoaderError(Exception):
    """Custom exception for Excel loading failures."""
    pass

def load_complex_excel(filepath: str) -> Workbook:
    """
    Loads an Excel file after running the strict pre-processing protocol.
    Acts as a 'Gatekeeper' to ensure only clean files are loaded.
    Implements a fallback strategy: read_only=False -> read_only=True.
    """
    if not os.path.exists(filepath):
        raise ExcelLoaderError(f"File not found: {filepath}")

    processed_path = None
    try:
        # 1. Pre-process (Physical Removal, Namespace Transplant, Surgery)
        processed_path = preprocess_excel(filepath)

        # 2. Load with OpenPyXL
        try:
            # First Attempt: Standard Loading (data_only=True, read_only=False)
            # read_only=False is preferred to recalc dimensions
            wb = openpyxl.load_workbook(
                processed_path,
                data_only=True,
                read_only=False,
                keep_links=False
            )
            return wb
        except Exception as e:
            print(f"标准模式加载异常 ({e})，切换只读模式...")
            # Fallback: Read-Only Mode
            # This is safer for very broken files but might miss dimensions
            wb = openpyxl.load_workbook(
                processed_path,
                data_only=True,
                read_only=True,
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
