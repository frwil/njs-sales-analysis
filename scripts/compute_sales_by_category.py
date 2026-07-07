"""
Compute actual sales (Jan-Jun 2026) per category, per month, per agency.
Compare with objectives.
Calculate average price per ton per category.
Forecast S2 2026 in 3 scenarios.
"""
import os
import json
import re
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict

# ===== LOAD DATA =====
SRC_VENTES = "/home/z/my-project/download/ventes_livrees.xlsx"
SRC_OBJECTIFS = "/home/z/my-project/upload/objectifs 2026 - Takou.xlsx"
OUT_DIR = "/home/z/my-project/scripts/ventes_objectifs_data"

with open("/home/z/my-project/scripts/product_category_map.json", "r", encoding="utf-8") as f:
    PRODUCT_CATEGORY = json.load(f)

# Load excluded clients
with open("/home/z/my-project/scripts/excluded_clients.json", "r", encoding="utf-8") as f:
    excluded_tiers = set(json.load(f))

# ===== WEIGHT PARSER =====
WEIGHT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(KG|Kg|kg|KG|GRAMMES?|G|L)\b", re.IGNORECASE)
def parse_weight_kg(desc):
    if not desc: return 0.0
    matches = WEIGHT_RE.findall(str(desc))
    if not matches: return 0.0
    val_str, unit = matches[-1]
    val = float(val_str)
    u = unit.upper()
    if u == "KG": return val
    if u in ("G", "GRAMME", "GRAMMES"): return val/1000.0
    if u == "L": return val
    return 0.0

# ===== AGENCY NAME NORMALIZATION =====
# Sales file uses "AGENCE FAMLA", objectives file uses "Famla"
def normalize_agence(agence_str):
    """Convert 'AGENCE FAMLA' to 'Famla' to match objectives file."""
    if not agence_str:
        return ""
    s = str(agence_str).strip()
    # Remove 'AGENCE ' prefix
    if s.upper().startswith("AGENCE "):
        s = s[7:]
    # Remove 'AGENCE DE ' prefix
    if s.upper().startswith("DE "):
        s = s[3:]
    # Special cases
    if "BAMENDA" in s.upper() and "MBOUDA" in s.upper():
        return "Mbouda"
    if "BERTOUA" in s.upper():
        return "Bertoua"
    # Capitalize first letter
    if s:
        s = s[0].upper() + s[1:].lower()
    return s

# ===== READ SALES FILE =====
print("Reading sales file...")
wb_src = load_workbook(SRC_VENTES, read_only=True, data_only=True)

# Per-category aggregates
# category_sales[category][month] = {"vol_kg": float, "ca": float}
category_sales = defaultdict(lambda: defaultdict(lambda: {"vol_kg": 0.0, "ca": 0.0}))
# Per-agency per-category
agency_category_sales = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {"vol_kg": 0.0, "ca": 0.0})))
# Per-product (for verification)
product_sales = defaultdict(lambda: {"vol_kg": 0.0, "ca": 0.0})

sheet_to_month = {"Sheet 1": 1, "Feuil1": 2, "Feuil2": 3, "Feuil3": 4, "Feuil4": 5, "Feuil5": 6}

for sheet_name in wb_src.sheetnames:
    ws = wb_src[sheet_name]
    month_num = sheet_to_month.get(sheet_name)
    if month_num is None: continue
    for row in ws.iter_rows(min_row=3, values_only=True):
        tiers = row[5] if len(row) > 5 else None
        if tiers is None or str(tiers).strip() == "": continue
        if tiers in excluded_tiers: continue
        ref_prod = row[0] if len(row) > 0 else None
        desc = row[1] if len(row) > 1 else None
        qte = row[2] if len(row) > 2 else 0
        ca_ht = row[8] if len(row) > 8 else 0
        agence = row[17] if len(row) > 17 else None

        ref_prod_str = str(ref_prod).strip() if ref_prod is not None else ""
        if not ref_prod_str: continue
        category = PRODUCT_CATEGORY.get(ref_prod_str, "DIVERS")

        # Compute weight
        # We don't have product_weight pre-loaded; parse from desc
        weight_kg = parse_weight_kg(desc)

        try: qte_f = float(qte) if qte is not None else 0.0
        except: qte_f = 0.0
        try: ca_f = float(ca_ht) if ca_ht is not None else 0.0
        except: ca_f = 0.0
        vol_kg = qte_f * weight_kg

        category_sales[category][month_num]["vol_kg"] += vol_kg
        category_sales[category][month_num]["ca"] += ca_f
        product_sales[ref_prod_str]["vol_kg"] += vol_kg
        product_sales[ref_prod_str]["ca"] += ca_f

        # Agence normalization
        agence_norm = normalize_agence(agence)
        if agence_norm:
            agency_category_sales[agence_norm][category][month_num]["vol_kg"] += vol_kg
            agency_category_sales[agence_norm][category][month_num]["ca"] += ca_f

wb_src.close()

print("Sales data loaded.")
print(f"\nCategories found in sales: {sorted(category_sales.keys())}")

# Convert kg to tons
print("\nActual sales per category (Jan-Jun 2026) in TONS:")
print(f"{'Category':<28}{'Jan':>8}{'Fev':>8}{'Mar':>8}{'Avr':>8}{'Mai':>8}{'Juin':>8}{'TOTAL':>10}{'CA (M FCFA)':>14}")
for cat in sorted(category_sales.keys()):
    total_vol = 0
    total_ca = 0
    months_str = ""
    for m in range(1, 7):
        v = category_sales[cat][m]["vol_kg"] / 1000.0
        total_vol += v
        total_ca += category_sales[cat][m]["ca"]
        months_str += f"{v:>8.1f}"
    print(f"{cat:<28}{months_str}{total_vol:>10.1f}{total_ca/1e6:>14.1f}")

# Compute average price per ton per category (over 6 months)
print("\nAverage price per ton per category (Jan-Jun 2026):")
price_per_ton = {}
for cat in sorted(category_sales.keys()):
    total_vol_kg = sum(category_sales[cat][m]["vol_kg"] for m in range(1, 7))
    total_ca = sum(category_sales[cat][m]["ca"] for m in range(1, 7))
    if total_vol_kg > 0:
        price = total_ca / (total_vol_kg / 1000.0)  # FCFA per ton
    else:
        price = 0
    price_per_ton[cat] = price
    print(f"  {cat:<28} {price:>12,.0f} FCFA/t  (vol={total_vol_kg/1000.0:>10.1f} t, CA={total_ca:>15,.0f})")

# Save data for next steps
data = {
    "category_sales_tons": {cat: {str(m): category_sales[cat][m]["vol_kg"]/1000.0 for m in range(1,7)} for cat in category_sales},
    "category_sales_ca": {cat: {str(m): category_sales[cat][m]["ca"] for m in range(1,7)} for cat in category_sales},
    "price_per_ton": price_per_ton,
    "agency_category_sales_tons": {ag: {cat: {str(m): agency_category_sales[ag][cat][m]["vol_kg"]/1000.0 for m in range(1,7)} for cat in agency_category_sales[ag]} for ag in agency_category_sales},
    "agency_category_sales_ca": {ag: {cat: {str(m): agency_category_sales[ag][cat][m]["ca"] for m in range(1,7)} for cat in agency_category_sales[ag]} for ag in agency_category_sales},
}

with open(f"{OUT_DIR}.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2, default=str)
print(f"\nData saved to {OUT_DIR}.json")
