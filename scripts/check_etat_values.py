"""Check the distinct values in the 'État' column (column P / index 16) for each sheet."""
from openpyxl import load_workbook
from collections import Counter

SRC = "/home/z/my-project/upload/ventes janv a juin 2026.xlsx"

wb = load_workbook(SRC, read_only=True, data_only=True)

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    print(f"\n{'='*80}")
    print(f"SHEET: {sheet_name}")
    print('='*80)
    counter = Counter()
    total_data_rows = 0
    for row in ws.iter_rows(min_row=3, values_only=True):
        total_data_rows += 1
        etat = row[15] if len(row) > 15 else None  # column P = index 15 (0-based)
        counter[etat] += 1
    print(f"Total data rows: {total_data_rows}")
    print("État value distribution:")
    for val, count in counter.most_common():
        print(f"  {repr(val)}: {count}")

wb.close()
