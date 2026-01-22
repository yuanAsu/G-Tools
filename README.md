# Complex Excel Parser

A Python application for parsing Excel files with complex structures, such as merged cells and multi-level headers.

## Features

- **Merged Cell Handling:** Automatically propagates values from merged cells to all cells in the range.
- **Multi-Level Headers:** Flattens multiple header rows into a single column name (e.g., "Q1 - Sales").
- **Modular Design:** Separation of parsing logic (`src/parser.py`) from usage.

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Generate sample data (optional, for testing):
   ```bash
   python scripts/generate_sample_excel.py
   ```

## Usage

```python
from src.parser import ExcelParser

parser = ExcelParser("data/complex_sample.xlsx")
df = parser.parse_sheet("Sales Report", header_rows=2)
print(df)
```

## Tools

- `scripts/analyze_excel.py`: Inspects an Excel file and reports its structure (sheet names, merged cells, etc.).
  ```bash
  python scripts/analyze_excel.py data/complex_sample.xlsx
  ```

## Testing

Run tests with pytest:
```bash
python -m pytest tests/
```
**Note:** You must run the sample generator script first to create the test data.
