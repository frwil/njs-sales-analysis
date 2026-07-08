"""
1. Fix P109 (Premix Multi Rumi) weight = 25kg
2. Compute sales vs objectives by agency (all agencies)
3. Compute CONCENTRES by agency
4. Compute monthly trend by category (all categories)
5. Compute monthly trend by agency (top 5)
"""
import os, json, re
from openpyxl import load_workbook
from collections import defaultdict

SRC_2026 = "/home/z/my-project/download/ventes_livrees.xlsx"
SRC_OBJ = "/home/z/my-project/upload/objectifs 2026 - Takou.xlsx"
SRC_OBJ_S2 = "/home/z/my-project/upload/DOC-20260706-WA0017.xlsx"

with open("/home/z/my-project/scripts/product_category_map.json", "r") as f:
    PRODUCT_CATEGORY = json.load(f)

# Manual weight overrides for misconfigured products
MANUAL_WEIGHTS = {
    "M1051": 50.0,   # Maïs en sacs de 50 kg
    "CF101": 1.0,    # Carbonate de calcium 1 kg
    "S101": 1.0,     # Sel en sachets de 1 kg
    "P109": 25.0,    # Premix Multi Rumi 25 kg
}
SERVICES_VALIDEE = {"PONT_BASCULE", "CONTRIBUTION_CARBURANT"}

WEIGHT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(KG|Kg|kg|KG|GRAMMES?|G|L)\b", re.IGNORECASE)
def parse_weight_kg(desc, ref=""):
    if ref in MANUAL_WEIGHTS: return MANUAL_WEIGHTS[ref]
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

# Agency name normalization
def normalize_agence(agence_str):
    if not agence_str: return ""
    s = str(agence_str).strip()
    if s.upper().startswith("AGENCE "): s = s[7:]
    if s.upper().startswith("DE "): s = s[3:]
    if "BAMENDA" in s.upper() and "MBOUDA" in s.upper(): return "Mbouda"
    if "BERTOUA" in s.upper(): return "Bertoua"
    if s: s = s[0].upper() + s[1:].lower()
    return s

# ===== READ 2026 SALES (NO exclusion) with agency =====
print("Reading 2026 sales by agency...")
wb = load_workbook(SRC_2026, read_only=True, data_only=True)

# agency_category[agency][category][month] = {"vol_t": float, "ca": float}
agency_cat_month = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {"vol_t": 0.0, "ca": 0.0})))
# category_month[category][month] = {"vol_t": float, "ca": float}
cat_month = defaultdict(lambda: defaultdict(lambda: {"vol_t": 0.0, "ca": 0.0}))
# agency_month[agency][month] = {"vol_t": float, "ca": float}
ag_month = defaultdict(lambda: defaultdict(lambda: {"vol_t": 0.0, "ca": 0.0}))

product_weight_cache = {}
sheet_to_month = {"Sheet 1":1, "Feuil1":2, "Feuil2":3, "Feuil3":4, "Feuil4":5, "Feuil5":6}

for sn in wb.sheetnames:
    ws = wb[sn]
    month_num = sheet_to_month.get(sn)
    if not month_num: continue
    for row in ws.iter_rows(min_row=3, values_only=True):
        tiers = row[5] if len(row)>5 else None
        if tiers is None or str(tiers).strip()=="": continue
        ref_prod = row[0] if len(row)>0 else None
        desc = row[1] if len(row)>1 else None
        qte = row[2] if len(row)>2 else 0
        ca = row[8] if len(row)>8 else 0
        agence = row[17] if len(row)>17 else None
        etat = row[15] if len(row)>15 else None

        ref_str = str(ref_prod).strip() if ref_prod else ""
        if not ref_str: continue

        # Filter: Livrée OR (service + Validée)
        etat_str = str(etat) if etat else ""
        if etat_str != "Livrée":
            if not (ref_str in SERVICES_VALIDEE and etat_str == "Validée"):
                continue

        # Exclure M1051 (Maïs) à CA=0 (régularisation stock)
        try: ca_check = float(ca) if ca else 0.0
        except: ca_check = 0.0
        if ref_str == "M1051" and ca_check == 0:
            continue

        cat = PRODUCT_CATEGORY.get(ref_str, "DIVERS")

        if ref_str not in product_weight_cache:
            product_weight_cache[ref_str] = parse_weight_kg(desc, ref_str)
        weight = product_weight_cache[ref_str]

        try: qte_f = float(qte) if qte else 0
        except: qte_f = 0
        try: ca_f = float(ca) if ca else 0
        except: ca_f = 0
        vol_t = qte_f * weight / 1000.0

        cat_month[cat][month_num]["vol_t"] += vol_t
        cat_month[cat][month_num]["ca"] += ca_f

        ag_norm = normalize_agence(agence)
        if ag_norm:
            agency_cat_month[ag_norm][cat][month_num]["vol_t"] += vol_t
            agency_cat_month[ag_norm][cat][month_num]["ca"] += ca_f
            ag_month[ag_norm][month_num]["vol_t"] += vol_t
            ag_month[ag_norm][month_num]["ca"] += ca_f

wb.close()

# ===== READ OBJECTIVES BY AGENCY =====
print("Reading objectives by agency...")
wb_obj = load_workbook(SRC_OBJ, data_only=True)
ws_obj = wb_obj.active
CATEGORIES = ['ALIMENT COMPLET','INGREDIENTS','COMPLEMENT ALIMENTAIRE','MATERIEL ELEVAGE',
              'PREMIX','CONCENTRES','TOURTEAUX','ALVEOLE','DIVERS']

# agency_obj[agency][category] = S1 objective (tons)
agency_obj = defaultdict(lambda: defaultdict(float))
for block_start in range(16, 327, 12):
    agence_cell = ws_obj.cell(row=block_start, column=1).value
    if agence_cell is None: continue
    agence_name = str(agence_cell).replace("Total ", "").strip()
    if not agence_name or agence_name == "0": continue
    for i, cat in enumerate(CATEGORIES):
        r = block_start + 1 + i
        s1 = sum(ws_obj.cell(row=r, column=5+j).value or 0 for j in range(6))
        agency_obj[agence_name][cat] = float(s1)
wb_obj.close()

# ===== BUILD AGENCY COMPARISON =====
print("\n=== VENTES vs OBJECTIFS PAR AGENCE (S1 2026) ===")
print(f"{'Agence':<16}{'S1 obj (t)':>12}{'S1 réel (t)':>12}{'% atteinte':>12}{'CA (M FCFA)':>14}")
print("-"*66)

agency_comparison = []
for ag in sorted(agency_obj.keys()):
    obj_total = sum(agency_obj[ag][c] for c in CATEGORIES)
    real_vol = sum(agency_cat_month.get(ag, {}).get(c, {}).get(m, {}).get("vol_t", 0)
                   for c in CATEGORIES for m in range(1, 7))
    real_ca = sum(agency_cat_month.get(ag, {}).get(c, {}).get(m, {}).get("ca", 0)
                  for c in CATEGORIES for m in range(1, 7))
    pct = (real_vol / obj_total * 100) if obj_total > 0 else 0
    print(f"{ag:<16}{obj_total:>12.1f}{real_vol:>12.1f}{pct:>11.1f}%{real_ca/1e6:>14.1f}")
    agency_comparison.append({
        "agence": ag, "s1_obj": obj_total, "s1_real": real_vol,
        "pct": pct, "ca": real_ca
    })

# ===== CONCENTRES BY AGENCY =====
print(f"\n=== CONCENTRES PAR AGENCE (S1 2026) ===")
print(f"{'Agence':<16}{'Vol (t)':>12}{'CA (M FCFA)':>14}{'Obj S1 (t)':>12}{'% atteinte':>12}")
print("-"*66)
conc_by_ag = []
for ag in sorted(agency_cat_month.keys(),
                 key=lambda a: -sum(agency_cat_month[a].get("CONCENTRES", {}).get(m, {}).get("vol_t", 0) for m in range(1,7))):
    vol = sum(agency_cat_month[ag].get("CONCENTRES", {}).get(m, {}).get("vol_t", 0) for m in range(1,7))
    ca = sum(agency_cat_month[ag].get("CONCENTRES", {}).get(m, {}).get("ca", 0) for m in range(1,7))
    obj = agency_obj.get(ag, {}).get("CONCENTRES", 0)
    pct = (vol / obj * 100) if obj > 0 else 0
    if vol > 0:
        print(f"{ag:<16}{vol:>12.1f}{ca/1e6:>14.1f}{obj:>12.1f}{pct:>11.1f}%")
        conc_by_ag.append({"agence": ag, "vol": vol, "ca": ca, "obj": obj, "pct": pct})

# ===== MONTHLY TREND BY CATEGORY =====
print(f"\n=== TENDANCE MENSUELLE PAR CATÉGORIE (2026) ===")
cats_trend = ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'ALIMENT COMPLET', 'COMPLEMENT ALIMENTAIRE', 'PREMIX']
print(f"{'Catégorie':<28}", end='')
for m in ['Jan','Fév','Mar','Avr','Mai','Juin']:
    print(f'{m:>8}', end='')
print()
for cat in cats_trend:
    print(f"{cat:<28}", end='')
    for m in range(1,7):
        v = cat_month.get(cat, {}).get(m, {}).get("vol_t", 0)
        print(f"{v:>8.0f}", end='')
    print()

# ===== MONTHLY TREND BY AGENCY (top 5) =====
print(f"\n=== TENDANCE MENSUELLE PAR AGENCE (top 5, 2026) ===")
top5_ag = sorted(ag_month.keys(), key=lambda a: -sum(ag_month[a][m]["vol_t"] for m in range(1,7)))[:5]
print(f"{'Agence':<16}", end='')
for m in ['Jan','Fév','Mar','Avr','Mai','Juin']:
    print(f'{m:>8}', end='')
print(f"{'TOTAL':>10}")
for ag in top5_ag:
    print(f"{ag:<16}", end='')
    total = 0
    for m in range(1,7):
        v = ag_month[ag][m]["vol_t"]
        total += v
        print(f"{v:>8.0f}", end='')
    print(f"{total:>10.0f}")

# Save all data
output = {
    "agency_comparison": agency_comparison,
    "conc_by_agency": conc_by_ag,
    "cat_month_2026": {cat: {str(m): cat_month[cat][m] for m in range(1,7)} for cat in cat_month},
    "agency_month_2026": {ag: {str(m): ag_month[ag][m] for m in range(1,7)} for ag in ag_month},
    "top5_agencies": top5_ag,
}
with open("/home/z/my-project/scripts/agency_monthly_analysis.json", "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False, indent=2, default=str)
print(f"\nData saved to agency_monthly_analysis.json")

# Verify P109 fix
p109_vol = sum(cat_month.get("PREMIX", {}).get(m, {}).get("vol_t", 0) for m in range(1,7))
print(f"\nPREMIX total S1 (with P109 at 25kg): {p109_vol:.1f} t")
PYEOF
