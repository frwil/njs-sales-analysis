"""Inspect the new ERP extraction (5)(1).xlsx - corrected for the actual header."""
import openpyxl
from collections import Counter, defaultdict
from datetime import datetime

SRC = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (5) (1).xlsx"
wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
ws = wb.active
print(f"Sheet: {ws.title}")

rows = list(ws.iter_rows(values_only=True))
print(f"Total rows: {len(rows)}")
print(f"Row 0 (title): {rows[0]}")
print(f"Row 1 (header): {rows[1]}")
print(f"Row 2 (sample): {rows[2]}")

# Header is on row 1 (index 1)
header = list(rows[1])
data_rows = rows[2:]
print(f"\nData rows count: {len(data_rows)}")

# Column mapping
col_map = {h: i for i, h in enumerate(header) if h}
print(f"\nColumn map:")
for k, v in col_map.items():
    print(f"  '{k}' -> [{v}]")

# Indices
state_idx = col_map.get("État") or col_map.get("Etat") or col_map.get("Statut")
date_idx = col_map.get("Date de commande")
art_idx = col_map.get("Réf. produit")
art_desc_idx = col_map.get("Description du produit")
qty_idx = col_map.get("Qté commandée")
client_idx = col_map.get("Tiers")
agence_idx = col_map.get("agence")
ref_idx = col_map.get("Réf.")
print(f"\nIndices: state={state_idx}, date={date_idx}, art_ref={art_idx}, art_desc={art_desc_idx}, qte={qty_idx}, client={client_idx}, agence={agence_idx}, ref={ref_idx}")

# Filter Livree rows
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
        # Livree check
        s = str(r[state_idx]).strip() if r[state_idx] else ""
        if s in {"Livrée", "Livree", "LIVREE", "LIVRÉE"}:
            livree_dates[ds] += 1

print("\nTop 10 states:")
for s, c in states.most_common(10):
    print(f"  {s!r}: {c}")

print("\nDate distribution (all rows), last 30:")
for d, c in sorted(dates_count.items())[-30:]:
    print(f"  {d}: {c}")

print(f"\nTotal Livree rows: {sum(livree_dates.values())}")
print("Date distribution (Livree only), last 25:")
for d, c in sorted(livree_dates.items())[-25:]:
    print(f"  {d}: {c}")

# August coverage
aug_dates = sorted([d for d in livree_dates if d.startswith("2026-08")])
print(f"\nAugust dates covered (Livree): {aug_dates}")
aug_total = sum(livree_dates[d] for d in aug_dates)
print(f"August total Livree rows: {aug_total}")

# Per-day breakdown by category in August
def detect_cat(art_ref, art_desc):
    a = (art_ref or "").upper()
    d = (art_desc or "").upper()
    if "CONCENT" in d or "CONCENTR" in d or a.startswith("C1") or a.startswith("CB") or a.startswith("PB") or a.startswith("DB"):
        return "CONCENTRES"
    if "SOJA" in d or "TOURTEAU" in d or a.startswith("T1") or a == "SPC":
        return "TOURTEAUX/SOJA"
    if "MAIS" in d or a.startswith("M1"):
        return "MAIS"
    if "INGREDIENT" in d or a.startswith("I1") or a.startswith("B1") or a.startswith("E1"):
        return "INGREDIENTS"
    if "PREMIX" in d or a.startswith("PX"):
        return "PREMIX"
    return "OTHER"

aug_daily = defaultdict(lambda: defaultdict(float))
aug_count = defaultdict(int)
for r in data_rows:
    if not r or len(r) <= max(state_idx, date_idx, qty_idx):
        continue
    s = str(r[state_idx]).strip() if r[state_idx] else ""
    if s not in {"Livrée", "Livree", "LIVREE", "LIVRÉE"}:
        continue
    d = r[date_idx]
    if not isinstance(d, datetime):
        continue
    if d.strftime("%Y-%m") != "2026-08":
        continue
    ds = d.strftime("%Y-%m-%d")
    art_ref = r[art_idx] if art_idx is not None else ""
    art_desc = r[art_desc_idx] if art_desc_idx is not None else ""
    qty = r[qty_idx] if qty_idx is not None and r[qty_idx] else 0
    cat = detect_cat(art_ref, art_desc)
    aug_daily[ds][cat] += float(qty)
    aug_count[ds] += 1

print("\nAugust daily breakdown (Livree, qty by category):")
print(f"{'Date':12} {'CONC':>8} {'SOJA':>8} {'MAIS':>8} {'INGR':>8} {'PREMIX':>8} {'OTHER':>8} {'N_cmd':>6}")
for d in sorted(aug_daily):
    c = aug_daily[d]
    print(f"{d:12} {c.get('CONCENTRES',0):>8.0f} {c.get('TOURTEAUX/SOJA',0):>8.0f} {c.get('MAIS',0):>8.0f} {c.get('INGREDIENTS',0):>8.0f} {c.get('PREMIX',0):>8.0f} {c.get('OTHER',0):>8.0f} {aug_count[d]:>6}")

# Check first/last date
print(f"\nFirst date in extraction: {min(dates_count.keys())}")
print(f"Last date in extraction: {max(dates_count.keys())}")

# Check last 2 days - if morning extraction
last_date_str = max(livree_dates.keys())
print(f"\nLast Livree date: {last_date_str} with {livree_dates[last_date_str]} rows")
last_dt = datetime.strptime(last_date_str, "%Y-%m-%d")
# Check today
today = datetime(2026, 8, 15)
print(f"Today: {today.strftime('%Y-%m-%d')}")
days_diff = (today - last_dt).days
print(f"Days since last Livree: {days_diff}")
