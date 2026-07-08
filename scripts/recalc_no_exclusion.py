"""
Recalculate ALL sales data for ventes vs objectifs:
- NO exclusion of SPC/PDC/comptoirs (they are sales channels, not external clients)
- Add weight=50kg for Maïs (M1051) which has no weight in description
- Recompute category_sales_tons, category_sales_ca, price_per_ton
- Recompute objectives comparison
- Recompute YoY (2025 also without exclusion + Maïs 50kg)
- Recompute forecast
"""
import os
import json
import re
from openpyxl import load_workbook
from collections import defaultdict

SRC_2026 = "/home/z/my-project/download/ventes_livrees.xlsx"
SRC_2025 = "/home/z/my-project/upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx"

with open("/home/z/my-project/scripts/product_category_map.json", "r", encoding="utf-8") as f:
    PRODUCT_CATEGORY = json.load(f)

# Override: Manual weights for misconfigured products
MANUAL_WEIGHTS = {
    "M1051": 50.0,   # Maïs en sacs de 50 kg
    "CF101": 1.0,    # Carbonate de calcium 1 kg (confirmé par ratio prix vs CF1012)
    "S101": 1.0,     # Sel en sachets de 1 kg
}

# Services dont l'état 'Validée' doit être inclus (jamais 'Livrée')
SERVICES_VALIDEE = {"PONT_BASCULE", "CONTRIBUTION_CARBURANT"}

WEIGHT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(KG|Kg|kg|KG|GRAMMES?|G|L)\b", re.IGNORECASE)
def parse_weight_kg(desc, ref=""):
    # Manual override first
    if ref in MANUAL_WEIGHTS:
        return MANUAL_WEIGHTS[ref]
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

CATEGORIES = ['ALIMENT COMPLET', 'INGREDIENTS', 'COMPLEMENT ALIMENTAIRE', 'MATERIEL ELEVAGE',
              'PREMIX', 'CONCENTRES', 'TOURTEAUX', 'ALVEOLE', 'DIVERS']

# ===== 2026 SALES (NO EXCLUSION) =====
print("Reading 2026 sales (NO exclusion, with Maïs 50kg)...")
wb = load_workbook(SRC_2026, read_only=True, data_only=True)

category_sales_2026 = defaultdict(lambda: defaultdict(lambda: {"vol_kg": 0.0, "ca": 0.0}))
product_weight = {}

sheet_to_month = {"Sheet 1": 1, "Feuil1": 2, "Feuil2": 3, "Feuil3": 4, "Feuil4": 5, "Feuil5": 6}

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    month_num = sheet_to_month.get(sheet_name)
    if month_num is None: continue
    for row in ws.iter_rows(min_row=3, values_only=True):
        tiers = row[5] if len(row) > 5 else None
        if tiers is None or str(tiers).strip() == "": continue
        # NO EXCLUSION — all sales count for ventes vs objectifs
        ref_prod = row[0] if len(row) > 0 else None
        desc = row[1] if len(row) > 1 else None
        qte = row[2] if len(row) > 2 else 0
        ca_ht = row[8] if len(row) > 8 else 0
        etat = row[15] if len(row) > 15 else None

        ref_prod_str = str(ref_prod).strip() if ref_prod is not None else ""
        if not ref_prod_str: continue

        # Filter: Livrée OR (service + Validée)
        etat_str = str(etat) if etat else ""
        if etat_str != "Livrée":
            if not (ref_prod_str in SERVICES_VALIDEE and etat_str == "Validée"):
                continue

        # Exclure M1051 (Maïs) à CA=0 (régularisation stock)
        try: ca_check = float(ca_ht) if ca_ht is not None else 0.0
        except: ca_check = 0.0
        if ref_prod_str == "M1051" and ca_check == 0:
            continue

        category = PRODUCT_CATEGORY.get(ref_prod_str, "DIVERS")

        if ref_prod_str not in product_weight:
            product_weight[ref_prod_str] = parse_weight_kg(desc, ref_prod_str)
        weight_kg = product_weight[ref_prod_str]

        try: qte_f = float(qte) if qte is not None else 0.0
        except: qte_f = 0.0
        try: ca_f = float(ca_ht) if ca_ht is not None else 0.0
        except: ca_f = 0.0
        vol_kg = qte_f * weight_kg

        category_sales_2026[category][month_num]["vol_kg"] += vol_kg
        category_sales_2026[category][month_num]["ca"] += ca_f

wb.close()

# Print results
print("\n=== VENTES 2026 SANS EXCLUSION (avec Maïs 50kg) ===")
print(f"{'Catégorie':<28}", end='')
for m in ['Jan','Fév','Mar','Avr','Mai','Juin']:
    print(f'{m:>8}', end='')
print(f"{'TOTAL (t)':>12}{'CA (M FCFA)':>14}")
price_per_ton = {}
for cat in sorted(category_sales_2026.keys()):
    print(f"{cat:<28}", end='')
    total_vol = 0
    total_ca = 0
    for m in range(1, 7):
        v = category_sales_2026[cat][m]["vol_kg"] / 1000.0
        total_vol += v
        total_ca += category_sales_2026[cat][m]["ca"]
        print(f"{v:>8.1f}", end='')
    price = (total_ca / total_vol) if total_vol > 0 else 0
    price_per_ton[cat] = price
    print(f"{total_vol:>12.1f}{total_ca/1e6:>14.1f}")

# Save
sales_data = {
    "category_sales_tons": {cat: {str(m): category_sales_2026[cat][m]["vol_kg"]/1000.0 for m in range(1,7)} for cat in category_sales_2026},
    "category_sales_ca": {cat: {str(m): category_sales_2026[cat][m]["ca"] for m in range(1,7)} for cat in category_sales_2026},
    "price_per_ton": price_per_ton,
}
with open("/home/z/my-project/scripts/ventes_objectifs_data.json", "w", encoding="utf-8") as f:
    json.dump(sales_data, f, ensure_ascii=False, indent=2, default=str)
print(f"\nSaved ventes_objectifs_data.json")

# ===== 2025 SALES (NO EXCLUSION, with Maïs 50kg) =====
print("\nReading 2025 sales (NO exclusion, with Maïs 50kg)...")
wb2 = load_workbook(SRC_2025, read_only=True, data_only=True)
ws2 = wb2.active

category_2025 = defaultdict(lambda: defaultdict(lambda: {"vol_t": 0.0, "ca": 0.0}))
for row in ws2.iter_rows(min_row=2, values_only=True):
    cat = row[10] if len(row) > 10 else None
    if cat is None: continue
    cat_str = str(cat).strip()
    if cat_str == "DIVERS2": cat_str = "DIVERS"

    ref_prod = row[0] if len(row) > 0 else None
    vol_t = row[11] if len(row) > 11 else 0  # already in tons
    ca_ht = row[6] if len(row) > 6 else 0
    date_cmd = row[5] if len(row) > 5 else None

    # Special case: Maïs volume is 0 in source (no weight), recalculate
    ref_str = str(ref_prod).strip() if ref_prod else ""
    if ref_str == "M1051":
        qte = row[2] if len(row) > 2 else 0
        try: qte_f = float(qte) if qte else 0
        except: qte_f = 0
        vol_t = qte_f * MANUAL_WEIGHTS["M1051"] / 1000.0

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

    category_2025[cat_str][month]["vol_t"] += vol_t_f
    category_2025[cat_str][month]["ca"] += ca_f

wb2.close()

print("\n=== VENTES 2025 SANS EXCLUSION (avec Maïs 50kg) ===")
for cat in sorted(category_2025.keys()):
    total = sum(category_2025[cat][m]["vol_t"] for m in range(1, 13))
    ca = sum(category_2025[cat][m]["ca"] for m in range(1, 13))
    print(f"  {cat:<28} total={total:>10.1f} t  CA={ca/1e6:>10.1f} M")

# ===== RECOMPUTE OBJECTIVES COMPARISON =====
print("\n=== RECOMPUTING OBJECTIVES COMPARISON ===")
with open("/home/z/my-project/scripts/objectives_comparison.json", "r") as f:
    obj_data = json.load(f)

comparison_data = []
for cat in CATEGORIES + ["Innovations"]:
    q1_obj = sum(obj_data['global_objectives'].get(cat, {}).get(str(m), 0) for m in [1,2,3])
    q2_obj = sum(obj_data['global_objectives'].get(cat, {}).get(str(m), 0) for m in [4,5,6])
    s1_obj = q1_obj + q2_obj

    q1_real = sum(sales_data["category_sales_tons"].get(cat, {}).get(str(m), 0) for m in [1,2,3])
    q2_real = sum(sales_data["category_sales_tons"].get(cat, {}).get(str(m), 0) for m in [4,5,6])
    s1_real = q1_real + q2_real

    pct_q1 = (q1_real / q1_obj * 100) if q1_obj > 0 else 0
    pct_q2 = (q2_real / q2_obj * 100) if q2_obj > 0 else 0
    pct_s1 = (s1_real / s1_obj * 100) if s1_obj > 0 else 0

    comparison_data.append({
        "category": cat,
        "q1_obj": q1_obj, "q1_real": q1_real, "pct_q1": pct_q1,
        "q2_obj": q2_obj, "q2_real": q2_real, "pct_q2": pct_q2,
        "s1_obj": s1_obj, "s1_real": s1_real, "pct_s1": pct_s1,
        "price_per_ton": sales_data["price_per_ton"].get(cat, 0),
    })

q1_obj_tot = sum(d["q1_obj"] for d in comparison_data)
q2_obj_tot = sum(d["q2_obj"] for d in comparison_data)
s1_obj_tot = sum(d["s1_obj"] for d in comparison_data)
q1_real_tot = sum(d["q1_real"] for d in comparison_data)
q2_real_tot = sum(d["q2_real"] for d in comparison_data)
s1_real_tot = sum(d["s1_real"] for d in comparison_data)

obj_data['comparison'] = comparison_data
obj_data['totals'] = {
    "q1_obj": q1_obj_tot, "q1_real": q1_real_tot,
    "q2_obj": q2_obj_tot, "q2_real": q2_real_tot,
    "s1_obj": s1_obj_tot, "s1_real": s1_real_tot,
}
with open("/home/z/my-project/scripts/objectives_comparison.json", "w", encoding="utf-8") as f:
    json.dump(obj_data, f, ensure_ascii=False, indent=2, default=str)

print(f"\nTotal S1 objectif : {s1_obj_tot:,.1f} t")
print(f"Total S1 réel     : {s1_real_tot:,.1f} t")
print(f"% atteinte S1     : {s1_real_tot/s1_obj_tot*100:.1f}%")

# ===== RECOMPUTE YoY + FORECAST =====
print("\n=== RECOMPUTING YoY + FORECAST ===")
s1_2025 = {}
s2_2025 = {}
for cat in set(list(category_2025.keys()) + CATEGORIES):
    mapped = cat if cat != "DIVERS2" else "DIVERS"
    if mapped not in s1_2025:
        s1_2025[mapped] = {"vol_t": 0, "ca": 0}
        s2_2025[mapped] = {"vol_t": 0, "ca": 0}
    for m in range(1, 7):
        s1_2025[mapped]["vol_t"] += category_2025.get(cat, {}).get(m, {"vol_t": 0})["vol_t"]
        s1_2025[mapped]["ca"] += category_2025.get(cat, {}).get(m, {"ca": 0})["ca"]
    for m in range(7, 13):
        s2_2025[mapped]["vol_t"] += category_2025.get(cat, {}).get(m, {"vol_t": 0})["vol_t"]
        s2_2025[mapped]["ca"] += category_2025.get(cat, {}).get(m, {"ca": 0})["ca"]

yoy_data = {}
for cat in CATEGORIES:
    s1_25 = s1_2025.get(cat, {"vol_t": 0})["vol_t"]
    d = next((x for x in comparison_data if x['category'] == cat), None)
    s1_26 = d['s1_real'] if d else 0
    s2_25 = s2_2025.get(cat, {"vol_t": 0})["vol_t"]
    yoy = ((s1_26 - s1_25) / s1_25 * 100) if s1_25 > 0 else 0
    yoy_data[cat] = {"s1_2025": s1_25, "s1_2026": s1_26, "yoy_pct": yoy, "s2_2025": s2_25}

# Forecast
forecast_refined = {}
for cat in CATEGORIES:
    s2_25 = yoy_data[cat]["s2_2025"]
    yoy = yoy_data[cat]["yoy_pct"] / 100
    s2_obj = sum(obj_data['global_objectives'].get(cat, {}).get(str(m), 0) for m in range(7, 13))
    pess_mult = max(0, 1 + yoy - 0.10)
    real_mult = 1 + yoy
    opt_mult = 1 + yoy + 0.10
    forecast_refined[cat] = {"s2_2025": s2_25, "yoy_pct": yoy*100,
                              "pess": s2_25 * pess_mult, "real": s2_25 * real_mult,
                              "opt": s2_25 * opt_mult, "s2_obj": s2_obj}

total_s2_25 = sum(forecast_refined[c]["s2_2025"] for c in CATEGORIES)
total_pess = sum(forecast_refined[c]["pess"] for c in CATEGORIES)
total_real = sum(forecast_refined[c]["real"] for c in CATEGORIES)
total_opt = sum(forecast_refined[c]["opt"] for c in CATEGORIES)

yoy_output = {
    "category_2025_monthly": {cat: {str(m): category_2025.get(cat, {}).get(m, {"vol_t": 0, "ca": 0}) for m in range(1, 13)} for cat in set(list(category_2025.keys()) + CATEGORIES)},
    "yoy_data": yoy_data,
    "forecast_refined": forecast_refined,
    "totals": {"s2_2025": total_s2_25, "pess": total_pess, "real": total_real, "opt": total_opt},
}
with open("/home/z/my-project/scripts/yoy_forecast.json", "w", encoding="utf-8") as f:
    json.dump(yoy_output, f, ensure_ascii=False, indent=2, default=str)

print(f"\nYoY + forecast recalculé:")
print(f"  S1 2025: {sum(yoy_data[c]['s1_2025'] for c in CATEGORIES):,.1f} t")
print(f"  S1 2026: {sum(yoy_data[c]['s1_2026'] for c in CATEGORIES):,.1f} t")
print(f"  YoY S1 : {(sum(yoy_data[c]['s1_2026'] for c in CATEGORIES)/sum(yoy_data[c]['s1_2025'] for c in CATEGORIES)-1)*100:+.1f}%")
print(f"  Forecast S2 réaliste: {total_real:,.1f} t")

# Print detailed comparison
print(f"\n=== COMPARAISON VENTES vs OBJECTIFS (SANS EXCLUSION, avec Maïs) ===")
print(f"{'Catégorie':<28}{'S1 obj (t)':>12}{'S1 réel (t)':>12}{'% S1':>8}")
for d in comparison_data:
    print(f"{d['category']:<28}{d['s1_obj']:>12.1f}{d['s1_real']:>12.1f}{d['pct_s1']:>7.1f}%")
print(f"{'TOTAL':<28}{s1_obj_tot:>12.1f}{s1_real_tot:>12.1f}{s1_real_tot/s1_obj_tot*100:>7.1f}%")

# Verify Maïs impact
mais_vol = sum(category_sales_2026.get("INGREDIENTS", {}).get(m, {}).get("vol_kg", 0) for m in range(1,7)) / 1000.0
mais_ca = 0
for row_ref in ["M1051"]:
    # Already in INGREDIENTS
    pass
print(f"\nMaïs (M1051) maintenant en INGREDIENTS: {mais_vol:.1f} t (vs 0.0 t avant)")
