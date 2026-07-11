"""
Recalcule les YoY depuis le fichier source 2025 avec:
- Même mapping cat_map (unifié)
- Avec comptoirs (pas d'exclusion)
- Maïs séparé des autres ingrédients
- Filtre S1 (Jan-Jun) et S2 (Jul-Dec) 2025
"""
import openpyxl
import json
import re
from collections import defaultdict

with open('/home/z/my-project/scripts/product_category_map.json') as f:
    cat_map = json.load(f)

MANUAL_WEIGHTS = {"M1051": 50.0, "CF101": 1.0, "S101": 1.0, "P109": 25.0}

# ===== LIRE 2025 =====
wb25 = openpyxl.load_workbook('/home/z/my-project/upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx', read_only=True, data_only=True)
ws25 = wb25.active

# Par catégorie et par produit (avec comptoirs, sans exclusion)
cat_s1_2025 = defaultdict(float)
cat_s2_2025 = defaultdict(float)
prod_s1_2025 = defaultdict(float)
prod_s2_2025 = defaultdict(float)

for row in ws25.iter_rows(min_row=2, values_only=True):
    if not row or len(row) < 14: continue
    ref = row[0]
    ca = row[6]
    etat = row[9]
    date_cmd = row[5]
    
    if not ref: continue
    if str(etat) != "Livrée": continue
    
    # Filtre S1/S2
    month = None
    if date_cmd:
        if hasattr(date_cmd, 'month'): month = date_cmd.month
        else:
            try:
                s = str(date_cmd)
                parts = s.split("/")
                if len(parts) == 3: month = int(parts[1])
                elif len(s) >= 7: month = int(s[5:7])
            except: pass
    if month is None: continue
    
    try: c = float(ca) if ca else 0
    except: continue
    
    ref_str = str(ref).strip()
    cat = cat_map.get(ref_str, "DIVERS")
    
    # Volume: pour Maïs, recalculer avec qte × 50kg
    vol_t = 0
    if ref_str == "M1051":
        try:
            q = float(row[2]) if row[2] else 0
            vol_t = q * 50 / 1000
        except: vol_t = 0
    else:
        try: vol_t = float(row[11]) if row[11] else 0
        except: vol_t = 0
    
    if month <= 6:
        cat_s1_2025[cat] += vol_t
        prod_s1_2025[ref_str] += vol_t
    else:
        cat_s2_2025[cat] += vol_t
        prod_s2_2025[ref_str] += vol_t

wb25.close()

# ===== LIRE 2026 (déjà calculé, avec comptoirs) =====
with open('/home/z/my-project/scripts/objectives_comparison.json') as f:
    obj = json.load(f)

# S1 2026 par catégorie (avec comptoirs)
cat_s1_2026 = {}
for d in obj['comparison']:
    cat_s1_2026[d['category']] = d['s1_real']

# ===== SÉPARATION MAÏS / AUTRES INGRÉDIENTS =====
mais_s1_2025 = prod_s1_2025.get("M1051", 0)
mais_s2_2025 = prod_s2_2025.get("M1051", 0)
ing_hors_mais_s1_2025 = cat_s1_2025.get("INGREDIENTS", 0) - mais_s1_2025
ing_hors_mais_s2_2025 = cat_s2_2025.get("INGREDIENTS", 0) - mais_s2_2025

# 2026
mais_s1_2026 = 1892  # déjà calculé (hors régularisation)
ing_hors_mais_s1_2026 = cat_s1_2026.get("INGREDIENTS", 0) - mais_s1_2026

# ===== AFFICHAGE =====
print("=" * 100)
print("YOY RECALCULÉ DEPUIS FICHIER SOURCE 2025 (avec comptoirs, Maïs séparé)")
print("=" * 100)
print(f"{'Catégorie':<30} {'S1 2025 (t)':>12} {'S1 2026 (t)':>12} {'YoY %':>10} {'S2 2025 (t)':>12}")
print("-" * 80)

categories = [
    ("TOURTEAUX", cat_s1_2025.get("TOURTEAUX", 0), cat_s1_2026.get("TOURTEAUX", 0), cat_s2_2025.get("TOURTEAUX", 0)),
    ("CONCENTRES", cat_s1_2025.get("CONCENTRES", 0), cat_s1_2026.get("CONCENTRES", 0), cat_s2_2025.get("CONCENTRES", 0)),
    ("ALIMENT COMPLET", cat_s1_2025.get("ALIMENT COMPLET", 0), cat_s1_2026.get("ALIMENT COMPLET", 0), cat_s2_2025.get("ALIMENT COMPLET", 0)),
    ("INGREDIENTS (hors Maïs)", ing_hors_mais_s1_2025, ing_hors_mais_s1_2026, ing_hors_mais_s2_2025),
    ("MAÏS (M1051)", mais_s1_2025, mais_s1_2026, mais_s2_2025),
    ("COMPLEMENT ALIM.", cat_s1_2025.get("COMPLEMENT ALIMENTAIRE", 0), cat_s1_2026.get("COMPLEMENT ALIMENTAIRE", 0), cat_s2_2025.get("COMPLEMENT ALIMENTAIRE", 0)),
    ("PREMIX", cat_s1_2025.get("PREMIX", 0), cat_s1_2026.get("PREMIX", 0), cat_s2_2025.get("PREMIX", 0)),
]

total_s1_2025 = 0
total_s1_2026 = 0
total_s2_2025 = 0
for cat, s25, s26, s2_25 in categories:
    yoy = (s26 - s25) / s25 * 100 if s25 > 0 else 0
    print(f"{cat:<30} {s25:>12.1f} {s26:>12.1f} {yoy:>+9.1f}% {s2_25:>12.1f}")
    total_s1_2025 += s25
    total_s1_2026 += s26
    total_s2_2025 += s2_25

yoy_total = (total_s1_2026 - total_s1_2025) / total_s1_2025 * 100
print("-" * 80)
print(f"{'TOTAL':<30} {total_s1_2025:>12.1f} {total_s1_2026:>12.1f} {yoy_total:>+9.1f}% {total_s2_2025:>12.1f}")

# Sauvegarder
output = {
    'yoy_corrected': {
        cat: {
            's1_2025': s25,
            's1_2026': s26,
            'yoy_pct': (s26 - s25) / s25 * 100 if s25 > 0 else 0,
            's2_2025': s2_25,
        } for cat, s25, s26, s2_25 in categories
    },
    'totals': {
        's1_2025': total_s1_2025,
        's1_2026': total_s1_2026,
        'yoy_pct': yoy_total,
        's2_2025': total_s2_2025,
    }
}
with open('/home/z/my-project/scripts/yoy_corrected.json', 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)
print("\n✓ Sauvegardé dans yoy_corrected.json")

# Forecast S2 2026 (sans Maïs)
print("\n" + "=" * 100)
print("FORECAST S2 2026 (sans Maïs — produit opportuniste)")
print("=" * 100)
print(f"{'Catégorie':<30} {'S2 2025 (t)':>12} {'YoY %':>10} {'🔴 Pess.':>10} {'🟡 Real.':>10} {'🟢 Opt.':>10}")
print("-" * 85)

forecast = {}
for cat, s25, s26, s2_25 in categories:
    if cat == "MAÏS (M1051)":
        # Pas de forecast pour le Maïs (opportuniste)
        print(f"{cat:<30} {s2_25:>12.1f} {'—':>10} {'—':>10} {'—':>10} {'—':>10}")
        forecast[cat] = {'s2_2025': s2_25, 'pess': None, 'real': None, 'opt': None}
        continue
    yoy = (s26 - s25) / s25 * 100 if s25 > 0 else 0
    yoy_dec = yoy / 100
    pess = s2_25 * (1 + yoy_dec - 0.10)
    real = s2_25 * (1 + yoy_dec)
    opt = s2_25 * (1 + yoy_dec + 0.10)
    print(f"{cat:<30} {s2_25:>12.1f} {yoy:>+9.1f}% {pess:>10.1f} {real:>10.1f} {opt:>10.1f}")
    forecast[cat] = {'s2_2025': s2_25, 'pess': pess, 'real': real, 'opt': opt}

# Total forecast (sans Maïs)
total_pess = sum(f['pess'] for f in forecast.values() if f['pess'] is not None)
total_real = sum(f['real'] for f in forecast.values() if f['real'] is not None)
total_opt = sum(f['opt'] for f in forecast.values() if f['opt'] is not None)
total_s2_25 = sum(f['s2_2025'] for f in forecast.values())
print("-" * 85)
print(f"{'TOTAL (hors Maïs)':<30} {total_s2_25:>12.1f} {'':>10} {total_pess:>10.1f} {total_real:>10.1f} {total_opt:>10.1f}")

# Sauvegarder forecast
with open('/home/z/my-project/scripts/forecast_corrected.json', 'w', encoding='utf-8') as f:
    json.dump({
        'forecast': forecast,
        'totals': {'pess': total_pess, 'real': total_real, 'opt': total_opt, 's2_2025': total_s2_25}
    }, f, ensure_ascii=False, indent=2)
print("\n✓ Forecast sauvegardé dans forecast_corrected.json")
