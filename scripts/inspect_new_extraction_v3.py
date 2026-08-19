"""Inspect the new ERP extraction (18).xlsx - determine coverage."""
import openpyxl
from collections import Counter

SRC = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (18).xlsx"
wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
ws = wb.active
print(f"Sheet: {ws.title}")

rows = list(ws.iter_rows(values_only=True))
print(f"Total rows: {len(rows)}")
print(f"Row 0 (title): {rows[0]}")
print(f"Row 1 (header): {rows[1]}")

header = list(rows[1])
data_rows = rows[2:]
print(f"\nData rows count: {len(data_rows)}")

col_map = {h: i for i, h in enumerate(header) if h}
state_idx = col_map.get("État") or col_map.get("Etat")
date_idx = col_map.get("Date de commande")
print(f"\nIndices: state={state_idx}, date={date_idx}")

states = Counter()
dates_count = Counter()
livree_dates = Counter()
for r in data_rows:
    if not r or len(r) <= max(state_idx, date_idx):
        continue
    if r[state_idx] is not None:
        states[str(r[state_idx])] += 1
    d = r[date_idx]
    if d is not None:
        ds = d if isinstance(d, str) else d.strftime("%Y-%m-%d")
        dates_count[ds] += 1
        s = str(r[state_idx]).strip() if r[state_idx] else ""
        if s in {"Livrée", "Livree", "LIVREE", "LIVRÉE"}:
            livree_dates[ds] += 1

print("\nTop 10 states:")
for s, c in states.most_common(10):
    print(f"  {s!r}: {c}")

print("\nDate distribution (Livree only), last 25:")
for d, c in sorted(livree_dates.items())[-25:]:
    print(f"  {d}: {c}")

print(f"\nTotal Livree rows: {sum(livree_dates.values())}")
print(f"First date: {min(dates_count.keys())}")
print(f"Last date: {max(dates_count.keys())}")

aug_dates = sorted([d for d in livree_dates if d.startswith("2026-08")])
print(f"\nAugust dates covered (Livree): {aug_dates}")
aug_total = sum(livree_dates[d] for d in aug_dates)
print(f"August total Livree rows: {aug_total}")

print("\n5 latest Livree dates:")
for d, c in sorted(livree_dates.items())[-5:]:
    print(f"  {d}: {c} rows")
