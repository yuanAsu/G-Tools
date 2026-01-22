import sys
import openpyxl

def analyze_excel(filepath):
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    except Exception as e:
        print(f"Error loading {filepath}: {e}")
        return

    print(f"Analysis of: {filepath}")
    print("=" * 60)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        print(f"\nSheet: [{sheet_name}]")
        print(f"  Dimensions: {ws.dimensions}")
        print(f"  Max Row: {ws.max_row}, Max Col: {ws.max_column}")

        merged_ranges = list(ws.merged_cells.ranges)
        if merged_ranges:
            print(f"  Merged Cells ({len(merged_ranges)}):")
            for rng in merged_ranges[:5]:
                print(f"    - {rng}")
            if len(merged_ranges) > 5:
                print(f"    ... and {len(merged_ranges)-5} more.")
        else:
            print("  Merged Cells: None")

        print("  Preview (First 10 rows):")
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True)):
            # Check if empty
            if all(c is None for c in row):
                continue # Skip empty rows in preview to save space? Or show them?
                # Showing them is better for debugging anchors.
                pass

            formatted_row = [str(val)[:20] + "..." if val and len(str(val)) > 20 else str(val) if val is not None else "." for val in row]
            print(f"    Row {i+1:2d}: {formatted_row}")

def main():
    if len(sys.argv) < 2:
        print("Usage: python analyze_excel.py <filepath>")
        sys.exit(1)

    analyze_excel(sys.argv[1])

if __name__ == "__main__":
    main()
