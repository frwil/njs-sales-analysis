"""
Deep-dive analysis on CONCENTRES (highest margin category).
Breakdown by sub-category (Chair/Porc/Ponte, 10%/5%), by agency, by month.
Compare with objectives, YoY, and forecast.
"""
import os
import json
import re
from openpyxl import load_workbook
from collections import defaultdict

SRC_VENTES_2026 = "/home/z/my-project/download/ventes_livrees.xlsx"
SRC_VENTES_2025 = "/home/z/my-project/upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx"

with open("/home/z/my-project/scripts/excluded_clients.json", "r", encoding="utf-8") as f:
    excluded_tiers = set(json.load(f))

COMPTOIR_RE = re.compile(r"CLIENTS?\s+COMPTOIR", re.IGNORECASE)
INTERNAL_RE = re.compile(r"FILIALE\s+GROUPE\s+NJS|SOLDE\s+COMPTA|BELGOCAM\b|NJS\s+GROUP\b", re.IGNORECASE)

# Concentrés sub-categories (corrigé selon descriptions réelles produits)
CONC_SUBCATS = {
    "BELGO 10% Chair": {"C104", "C1042", "C1043", "C1044"},
    "BELGO 5% Chair": {"C105", "C1053", "C1054", "C1055"},
    "BELGO 10% Ponte": {"C102", "C1022"},
    "BELGO 5% Ponte": {"C101"},
    "BELGO 10% Porc": {"C103"},
    "BELGO Ruminant": {"C108"},
    "BELGO Rabbit": {"ALAP25"},
}

# Weight parser
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

def get_conc_subcat(ref):
    for subcat, refs in CONC_SUBCATS.items():
        if ref in refs:
            return subcat
    return "Autres"

# ===== READ 2026 SALES =====
print("Reading 2026 sales for CONCENTRES deep-dive...")
wb = load_workbook(SRC_VENTES_2026, read_only=True, data_only=True)

# subcat_2026[subcat][month] = {"vol_t": float, "ca": float}
subcat_2026 = defaultdict(lambda: defaultdict(lambda: {"vol_t": 0.0, "ca": 0.0}))
# agency_conc_2026[agency][month] = {"vol_t": float, "ca": float}
agency_conc_2026 = defaultdict(lambda: defaultdict(lambda: {"vol_t": 0.0, "ca": 0.0}))
# product_2026[ref][month] = {"vol_t": float, "ca": float}
product_2026 = defaultdict(lambda: defaultdict(lambda: {"vol_t": 0.0, "ca": 0.0}))

CONC_REFS = set()
for refs in CONC_SUBCATS.values():
    CONC_REFS.update(refs)

sheet_to_month = {"Sheet 1": 1, "Feuil1": 2, "Feuil2": 3, "Feuil3": 4, "Feuil4": 5, "Feuil5": 6}

product_weight_cache = {}

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    month_num = sheet_to_month.get(sheet_name)
    if month_num is None: continue
    for row in ws.iter_rows(min_row=3, values_only=True):
        tiers = row[5] if len(row) > 5 else None
        if tiers is None or str(tiers).strip() == "": continue
        if tiers in excluded_tiers or COMPTOIR_RE.search(str(tiers)) or INTERNAL_RE.search(str(tiers)):
            continue
        ref_prod = row[0] if len(row) > 0 else None
        desc = row[1] if len(row) > 1 else None
        qte = row[2] if len(row) > 2 else 0
        ca_ht = row[8] if len(row) > 8 else 0
        agence = row[17] if len(row) > 17 else None
        etat = row[15] if len(row) > 15 else None

        ref_prod_str = str(ref_prod).strip() if ref_prod is not None else ""
        if ref_prod_str not in CONC_REFS: continue

        # Filter: Livrée OR (service + Validée)
        etat_str = str(etat) if etat else ""
        if etat_str != "Livrée":
            # Pour les CONCENTRES, on ne garde que Livrée (pas de services)
            continue

        if ref_prod_str not in product_weight_cache:
            product_weight_cache[ref_prod_str] = parse_weight_kg(desc)
        weight_kg = product_weight_cache[ref_prod_str]

        try: qte_f = float(qte) if qte is not None else 0.0
        except: qte_f = 0.0
        try: ca_f = float(ca_ht) if ca_ht is not None else 0.0
        except: ca_f = 0.0
        vol_t = qte_f * weight_kg / 1000.0

        subcat = get_conc_subcat(ref_prod_str)
        subcat_2026[subcat][month_num]["vol_t"] += vol_t
        subcat_2026[subcat][month_num]["ca"] += ca_f
        product_2026[ref_prod_str][month_num]["vol_t"] += vol_t
        product_2026[ref_prod_str][month_num]["ca"] += ca_f

        # Agence normalization
        ag_str = str(agence).strip() if agence else ""
        if ag_str.upper().startswith("AGENCE "):
            ag_str = ag_str[7:]
        if ag_str.upper().startswith("DE "):
            ag_str = ag_str[3:]
        if ag_str:
            ag_str = ag_str[0].upper() + ag_str[1:].lower()
            if "BAMENDA" in ag_str.upper() and "MBOUDA" in ag_str.upper():
                ag_str = "Mbouda"
            agency_conc_2026[ag_str][month_num]["vol_t"] += vol_t
            agency_conc_2026[ag_str][month_num]["ca"] += ca_f

wb.close()

# ===== READ 2025 SALES for CONCENTRES =====
print("Reading 2025 sales for CONCENTRES YoY...")
wb2 = load_workbook(SRC_VENTES_2025, read_only=True, data_only=True)
ws2 = wb2.active

subcat_2025 = defaultdict(lambda: defaultdict(lambda: {"vol_t": 0.0, "ca": 0.0}))

for row in ws2.iter_rows(min_row=2, values_only=True):
    tiers = row[4] if len(row) > 4 else None
    if tiers is None or str(tiers).strip() == "": continue
    if str(tiers) in excluded_tiers or COMPTOIR_RE.search(str(tiers)) or INTERNAL_RE.search(str(tiers)):
        continue
    cat = row[10] if len(row) > 10 else None
    if cat is None or str(cat).strip() != "CONCENTRES": continue

    ref_prod = row[0] if len(row) > 0 else None
    vol_t = row[11] if len(row) > 11 else 0
    ca_ht = row[6] if len(row) > 6 else 0
    date_cmd = row[5] if len(row) > 5 else None

    ref_prod_str = str(ref_prod).strip() if ref_prod is not None else ""
    subcat = get_conc_subcat(ref_prod_str)

    month = None
    if date_cmd:
        if hasattr(date_cmd, 'month'):
            month = date_cmd.month
        else:
            s = str(date_cmd)
            parts = s.split("/")
            if len(parts) == 3:
                try: month = int(parts[1])
                except: pass

    if month is None: continue

    try: vol_t_f = float(vol_t) if vol_t is not None else 0.0
    except: vol_t_f = 0.0
    try: ca_f = float(ca_ht) if ca_ht is not None else 0.0
    except: ca_f = 0.0

    subcat_2025[subcat][month]["vol_t"] += vol_t_f
    subcat_2025[subcat][month]["ca"] += ca_f

wb2.close()

# ===== PRINT RESULTS =====
print("\n=== CONCENTRÉS 2026 (S1) par sous-catégorie ===")
print(f"{'Sous-catégorie':<25}{'S1 vol (t)':>12}{'S1 CA (M FCFA)':>16}{'Prix moy (FCFA/t)':>20}{'% du vol':>10}")
total_vol = 0
total_ca = 0
for subcat in sorted(subcat_2026.keys()):
    vol = sum(subcat_2026[subcat][m]["vol_t"] for m in range(1, 7))
    ca = sum(subcat_2026[subcat][m]["ca"] for m in range(1, 7))
    price = ca / vol if vol > 0 else 0
    total_vol += vol
    total_ca += ca
    print(f"{subcat:<25}{vol:>12.1f}{ca/1e6:>16.1f}{price:>20,.0f}{vol:>10.1%}")
print(f"{'TOTAL':<25}{total_vol:>12.1f}{total_ca/1e6:>16.1f}{'':>20}{'100,0%':>10}")

print(f"\n=== CONCENTRÉS YoY (S1 2025 vs S1 2026) par sous-catégorie ===")
print(f"{'Sous-catégorie':<25}{'S1 2025 (t)':>14}{'S1 2026 (t)':>14}{'YoY %':>10}{'S2 2025 (t)':>14}")
for subcat in sorted(set(list(subcat_2025.keys()) + list(subcat_2026.keys()))):
    s1_25 = sum(subcat_2025.get(subcat, {}).get(m, {"vol_t": 0})["vol_t"] for m in range(1, 7))
    s1_26 = sum(subcat_2026.get(subcat, {}).get(m, {"vol_t": 0})["vol_t"] for m in range(1, 7))
    s2_25 = sum(subcat_2025.get(subcat, {}).get(m, {"vol_t": 0})["vol_t"] for m in range(7, 13))
    yoy = ((s1_26 - s1_25) / s1_25 * 100) if s1_25 > 0 else 0
    print(f"{subcat:<25}{s1_25:>14.1f}{s1_26:>14.1f}{yoy:>+9.1f}%{s2_25:>14.1f}")

print(f"\n=== CONCENTRÉS par produit (S1 2026) ===")
print(f"{'Réf':<10}{'S1 vol (t)':>12}{'S1 CA (M FCFA)':>16}{'Prix moy (FCFA/t)':>20}")
for ref in sorted(product_2026.keys()):
    vol = sum(product_2026[ref][m]["vol_t"] for m in range(1, 7))
    ca = sum(product_2026[ref][m]["ca"] for m in range(1, 7))
    price = ca / vol if vol > 0 else 0
    print(f"{ref:<10}{vol:>12.1f}{ca/1e6:>16.1f}{price:>20,.0f}")

print(f"\n=== CONCENTRÉS par agence (S1 2026) ===")
print(f"{'Agence':<20}{'S1 vol (t)':>12}{'S1 CA (M FCFA)':>16}")
for ag in sorted(agency_conc_2026.keys(), key=lambda a: -sum(agency_conc_2026[a][m]["vol_t"] for m in range(1,7))):
    vol = sum(agency_conc_2026[ag][m]["vol_t"] for m in range(1, 7))
    ca = sum(agency_conc_2026[ag][m]["ca"] for m in range(1, 7))
    print(f"{ag:<20}{vol:>12.1f}{ca/1e6:>16.1f}")

# Monthly evolution
print(f"\n=== CONCENTRÉS évolution mensuelle 2026 ===")
print(f"{'Sous-catégorie':<25}", end='')
for m in ['Jan','Fév','Mar','Avr','Mai','Juin']:
    print(f'{m:>8}', end='')
print(f"{'TOTAL':>10}")
for subcat in sorted(subcat_2026.keys()):
    print(f"{subcat:<25}", end='')
    total = 0
    for m in range(1, 7):
        v = subcat_2026[subcat][m]["vol_t"]
        total += v
        print(f"{v:>8.1f}", end='')
    print(f"{total:>10.1f}")

# Save data
output = {
    "subcat_2026": {sc: {str(m): subcat_2026[sc][m] for m in range(1,7)} for sc in subcat_2026},
    "subcat_2025": {sc: {str(m): subcat_2025.get(sc, {}).get(m, {"vol_t": 0, "ca": 0}) for m in range(1,13)} for sc in set(list(subcat_2025.keys()) + list(subcat_2026.keys()))},
    "product_2026": {ref: {str(m): product_2026[ref][m] for m in range(1,7)} for ref in product_2026},
    "agency_conc_2026": {ag: {str(m): agency_conc_2026[ag][m] for m in range(1,7)} for ag in agency_conc_2026},
}
with open("/home/z/my-project/scripts/concentres_deepdive.json", "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False, indent=2, default=str)
print(f"\nData saved to /home/z/my-project/scripts/concentres_deepdive.json")
