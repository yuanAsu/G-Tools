import pandas as pd
import openpyxl
from openpyxl.utils import range_boundaries

class ExcelParser:
    def __init__(self, filepath):
        self.filepath = filepath
        self.wb = openpyxl.load_workbook(filepath, data_only=True)

    def get_sheet_names(self):
        return self.wb.sheetnames

    def _fill_merged_cells(self, sheet):
        """
        Unmerges all cells in the sheet and fills them with the value
        from the top-left cell of the merged range.
        """
        # Create a list of ranges to process (copying the list because we'll be modifying it)
        merged_ranges = list(sheet.merged_cells.ranges)

        for merged_range in merged_ranges:
            min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
            top_left_value = sheet.cell(row=min_row, column=min_col).value

            sheet.unmerge_cells(str(merged_range))

            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    sheet.cell(row=row, column=col).value = top_left_value

    def parse_sheet(self, sheet_name, header_rows=1):
        """
        Parses a specific sheet.

        Args:
            sheet_name (str): Name of the sheet.
            header_rows (int): Number of rows to treat as headers.

        Returns:
            pd.DataFrame: The parsed data.
        """
        if sheet_name not in self.wb.sheetnames:
            raise ValueError(f"Sheet {sheet_name} not found.")

        # We work on a copy logic or reload to avoid persisting changes to self.wb if we want to be safe,
        # but for now we'll modify the loaded wb in memory.
        sheet = self.wb[sheet_name]

        # 1. Handle Merged Cells
        self._fill_merged_cells(sheet)

        # 2. Extract Data
        data = list(sheet.values)

        if not data:
            return pd.DataFrame()

        # 3. Process Headers
        if header_rows > 0:
            # Extract header rows
            raw_headers = data[:header_rows]
            body = data[header_rows:]

            # Combine headers if multi-row
            final_headers = []
            num_cols = len(raw_headers[0])

            for col_idx in range(num_cols):
                col_header_parts = []
                for row_idx in range(header_rows):
                    val = raw_headers[row_idx][col_idx]
                    if val and str(val).strip():
                        col_header_parts.append(str(val).strip())

                # Join with a separator, e.g., " - "
                final_headers.append(" - ".join(col_header_parts))

            # Create DataFrame
            df = pd.DataFrame(body, columns=final_headers)
        else:
            df = pd.DataFrame(data)

        return df

    def parse_all(self):
        """Returns a dict of DataFrames for all sheets."""
        results = {}
        for name in self.get_sheet_names():
            # Default to 1 header row, user might need to customize this per sheet in a real app
            results[name] = self.parse_sheet(name, header_rows=1)
        return results
