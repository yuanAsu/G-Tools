# [中文](README.md) | [English](README_EN.md)

# Complex Excel Parser

A Python application for parsing complex Excel files. It supports merged cells, multi-level headers, semantic anchoring, key-value forms, and hierarchical data extraction.

## Features

- **Standard Table Parsing:**
  - Handles merged cells (propagates values).
  - Flattens multi-level headers.

- **Semantic Anchoring (`parse_table_at_anchor`):**
  - Locates specific tables within a mixed sheet using keywords (e.g., "Funding Assumptions").
  - Extracts the data region relative to the anchor.

- **Form Extraction (`parse_form`):**
  - Scans sheets for Key-Value pairs (e.g., "Project Name: Apollo").
  - Useful for project details, cover sheets, and unstructured layouts.

- **Hierarchy Extraction (`parse_hierarchy`):**
  - Builds nested JSON/Dictionary structures from indented columns.
  - Ideal for financial statements (Assets -> Current Assets -> Cash).

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Generate sample data:
   ```bash
   python scripts/generate_sample_excel.py  # Basic V1 data
   python scripts/generate_v2_data.py       # Advanced V2 data (Anchors, Forms, Hierarchy)
   ```

## Usage

```python
from src.parser import ExcelParser
import json

parser = ExcelParser("data/complex_v2.xlsx")

# 1. Semantic Anchor
df_funding = parser.parse_table_at_anchor("Mixed Data", "Funding Assumptions")
print(df_funding)

# 2. Form Extraction
project_info = parser.parse_form("Project Form")
print(project_info)
# {'Project Name': 'Apollo Mission', ...}

# 3. Hierarchy Extraction
financials = parser.parse_hierarchy("Financial Statement")
print(json.dumps(financials, indent=2))
```

## Tools

- `scripts/analyze_excel.py`: Inspects an Excel file and reports its structure (sheet names, merged cells, etc.).
  ```bash
  python scripts/analyze_excel.py data/complex_v2.xlsx
  ```

## Testing

Run tests with pytest:
```bash
python -m pytest tests/
```
