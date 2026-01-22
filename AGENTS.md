# Project: Complex Excel Parser

## Goal
Build an application capable of parsing complex Excel files. "Complex" generally refers to files with:
- Merged cells (row or column spans)
- Multi-level headers
- Non-tabular data layout
- Multiple sheets with cross-references

## Guidelines
- **Structure vs Data:** Use `openpyxl` when layout preservation or cell property inspection (merges, styles) is required. Use `pandas` for efficient data manipulation once the structure is normalized.
- **Modularity:** Keep parsing logic separate from business logic.
- **Testing:** All parsing logic must be tested against synthetic complex data.
- **Types:** Use Python type hints.
