"""Inspect the recalibrated S2 objectives file."""
from openpyxl import load_workbook

SRC = "/home/z/my-project/upload/DOC-20260706-WA0017.xlsx"

wb = load_workbook(SRC, data_only=True)
print(f"Sheets: {wb.sheetnames}")

for sn in wb.sheetnames:
    ws = wb[sn]
    print(f"\n=== Sheet: {sn} (max_row={ws.max_row}, max_col={ws.max_column}) ===")
    
    # Print first 30 rows
    for r in range(1, min(35, ws.max_row + 1)):
        row_vals = []
        for c in range(1, min(ws.max_column + 1, 20)):
            v = ws.cell(row=r, column=c).value
            if v is not None:
                row_vals.append(f"C{c}={repr(v)[:50]}")
        if row_vals:
            print(f"  Row {r}: {' | '.join(row_vals)}")
