"""
Recalcule les données régionales correctes:
- NDOBO est dans le Littoral (pas le Centre)
- Les totaux doivent être cohérents avec le global 41 148 t / 99,8%
"""
import openpyxl
import json
import re
from collections import defaultdict

with open('/home/z/my-project/scripts/product_category_map.json') as f:
    cat_map = json.load(f)

MANUAL_WEIGHTS = {"M1051": 50.0, "CF101": 1.0, "S101": 1.0, "P109": 25.0}
SERVICES_VALIDEE = {"PONT_BASCULE", "CONTRIBUTION_CARBURANT"}
WEIGHT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(KG|Kg|kg|KG|GRAMMES?|G|L)\b", re.IGNORECASE)

def parse_weight_kg(ref, desc):
    ref_str = str(ref).strip() if ref else ""
    if ref_str in MANUAL_WEIGHTS: return MANUAL_WEIGHTS[ref_str]
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

def normalize_agence(agence_str):
    if not agence_str: return ""
    s = str(agence_str).strip()
    if s.upper().startswith("SPC") or s.upper().startswith("PDC"): return ""
    if "AGENCE " in s.upper(): s = s[7:]
    if s.upper().startswith("DE "): s = s[3:]
    if "BAMENDA" in s.upper() and "MBOUDA" in s.upper(): return "Mbouda"
    if "BERTOUA" in s.upper(): return "Bertoua"
    if s: s = s[0].upper() + s[1:].lower()
    return s

# Mapping correct
AGENCE_REGION = {
    "Ahala": "Centre", "Bertoua": "Centre", "Buea": "Littoral",
    "Djeleng": "Ouest", "Famla": "Ouest", "Mbouda": "Ouest",
    "Messassi": "Centre", "Ndobo": "Littoral", "Ngaoundere": "Centre",
    "Nkoabang": "Centre", "Nkolbisson": "Centre", "Nkongsamba": "Littoral",
    "Pk11": "Littoral", "Village": "Littoral",
}

# ===== LIRE VENTES 2026 =====
wb26 = openpyxl.load_workbook('/home/z/my-project/upload/ventes janv a juin 2026.xlsx', read_only=True, data_only=True)

region_vol_2026 = defaultdict(float)
region_ca_2026 = defaultdict(float)
region_conc_vol_2026 = defaultdict(float)

for sheet_name in wb26.sheetnames:
    ws = wb26[sheet_name]
    if ws.cell(2, 1).value != 'Réf. produit': continue
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row or len(row) < 16: continue
        ref = row[0]
        desc = row[1]
        qte = row[2]
        ca = row[8]
        etat = row[15]
        agence = row[17] if len(row) > 17 else None
        
        if not ref: continue
        ref_str = str(ref).strip()
        etat_str = str(etat) if etat else ""
        if etat_str != "Livrée":
            if not (ref_str in SERVICES_VALIDEE and etat_str == "Validée"):
                continue
        try:
            q = float(qte) if qte else 0
            c = float(ca) if ca else 0
        except: continue
        if ref_str == "M1051" and c == 0: continue
        
        agence_norm = normalize_agence(agence)
        if not agence_norm: continue
        region = AGENCE_REGION.get(agence_norm, "Autre")
        
        weight = parse_weight_kg(ref_str, desc)
        vol_t = q * weight / 1000.0
        
        region_vol_2026[region] += vol_t
        region_ca_2026[region] += c
        
        cat = cat_map.get(ref_str, "DIVERS")
        if cat == "CONCENTRES":
            region_conc_vol_2026[region] += vol_t

wb26.close()

# ===== LIRE OBJECTIFS PAR AGENCE =====
with open('/home/z/my-project/scripts/objectives_comparison.json') as f:
    obj_data = json.load(f)

# Somme des objectifs par agence → région
region_vol_obj = defaultdict(float)
region_conc_obj = defaultdict(float)

for agence, cats in obj_data.get('agence_objectives', {}).items():
    agence_norm = normalize_agence(agence)
    region = AGENCE_REGION.get(agence_norm, "Autre")
    for cat, months in cats.items():
        vol = sum(float(months.get(str(m), 0)) for m in range(1, 7))
        region_vol_obj[region] += vol
        if cat == "CONCENTRES":
            region_conc_obj[region] += vol

# ===== AFFICHAGE =====
print("=" * 110)
print("ANALYSE RÉGIONALE CORRIGÉE (S1 2026)")
print("=" * 110)
print(f"{'Région':<10} {'Agences':<50} {'Vol obj':>10} {'Vol réel':>10} {'% vol':>8} {'CA (M)':>10} {'Conc vol':>10} {'Conc obj':>10} {'% conc':>8}")
print("-" * 130)

regions_info = {
    "Ouest": "FAMLA, DJELENG, MBOUDA",
    "Centre": "MESSASSI, AHALA, BERTOUA, NGAOUNDERE, NKOABANG, NKOLBISSON",
    "Littoral": "NDOBO, BUEA, NKONGSAMBA, PK11, VILLAGE",
}

total_vol_obj = 0
total_vol_real = 0
total_ca = 0
total_conc_vol = 0
total_conc_obj = 0

for r in ["Ouest", "Centre", "Littoral"]:
    vo = region_vol_obj.get(r, 0)
    vr = region_vol_2026.get(r, 0)
    ca = region_ca_2026.get(r, 0)
    cv = region_conc_vol_2026.get(r, 0)
    co = region_conc_obj.get(r, 0)
    pct_v = vr/vo*100 if vo > 0 else 0
    pct_c = cv/co*100 if co > 0 else 0
    print(f"{r:<10} {regions_info[r]:<50} {vo:>10.0f} {vr:>10.0f} {pct_v:>7.1f}% {ca/1e6:>10.0f} {cv:>10.0f} {co:>10.0f} {pct_c:>7.1f}%")
    total_vol_obj += vo
    total_vol_real += vr
    total_ca += ca
    total_conc_vol += cv
    total_conc_obj += co

print("-" * 130)
pct_total_v = total_vol_real/total_vol_obj*100 if total_vol_obj > 0 else 0
pct_total_c = total_conc_vol/total_conc_obj*100 if total_conc_obj > 0 else 0
print(f"{'TOTAL':<10} {'14 agences':<50} {total_vol_obj:>10.0f} {total_vol_real:>10.0f} {pct_total_v:>7.1f}% {total_ca/1e6:>10.0f} {total_conc_vol:>10.0f} {total_conc_obj:>10.0f} {pct_total_c:>7.1f}%")

# YoY depuis yoy_agence_region.json
with open('scripts/yoy_agence_region.json') as f:
    yoy_data = json.load(f)

print("\n=== YoY PAR RÉGION ===")
print(f"{'Région':<10} {'YoY vol':>10} {'YoY CA':>10}")
for r in ["Ouest", "Centre", "Littoral"]:
    d = yoy_data['yoy_region'].get(r, {})
    print(f"{r:<10} {d.get('yoy_vol_pct', 0):>+9.1f}% {d.get('yoy_ca_pct', 0):>+9.1f}%")

# Sauvegarder
output = {
    r: {
        "agences": regions_info[r],
        "vol_obj": region_vol_obj.get(r, 0),
        "vol_real": region_vol_2026.get(r, 0),
        "ca": region_ca_2026.get(r, 0),
        "conc_vol": region_conc_vol_2026.get(r, 0),
        "conc_obj": region_conc_obj.get(r, 0),
        "yoy_vol_pct": yoy_data['yoy_region'].get(r, {}).get('yoy_vol_pct', 0),
        "yoy_ca_pct": yoy_data['yoy_region'].get(r, {}).get('yoy_ca_pct', 0),
    } for r in ["Ouest", "Centre", "Littoral"]
}
with open('scripts/region_obj_analysis_corrected.json', 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)
print("\n✓ Sauvegardé dans region_obj_analysis_corrected.json")
