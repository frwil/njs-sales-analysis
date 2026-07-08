"""
Save real CA objectives and recompute all CA comparisons.
Uses the CA objectives provided directly by the user (not calculated from volume × price).
"""
import json
from collections import defaultdict

# ===== CA OBJECTIFS RÉELS (fournis par l'utilisateur) =====
# Monthly CA per category (Jan-Dec)
CA_OBJ_MONTHLY = {
    "ALIMENT COMPLET": [79453851, 70999192, 78610181, 64250610, 70162616, 70999192,
                         62557835, 55790143, 58336188, 76075004, 76922461, 93805137],
    "INGREDIENTS": [147636291, 131383263, 147636291, 119214956, 131383263, 132722580,
                     116500267, 104313742, 108364857, 142225199, 144923586, 176058199],
    "COMPLEMENT ALIMENTAIRE": [2625488]*12,
    "MATERIEL ELEVAGE": [29647784, 27896550, 29534633, 26839511, 27865895, 28235159,
                          26646430, 25214885, 25959839, 29150649, 29392509, 50480729],
    "PREMIX": [20365839, 18514804, 20365839, 16667540, 18514804, 18514804,
                16667540, 14812597, 14812597, 18514804, 20365839, 24064340],
    "CONCENTRES": [1513938189, 1344990844, 1504697079, 1218280540, 1337728488, 1357527607,
                    1189246213, 1064514504, 1110707563, 1451242612, 1472358335, 1791113140],
    "TOURTEAUX": [1595479046, 1417622692, 1586324658, 1284229855, 1409775747, 1431027005,
                   1253497267, 1121739469, 1170453890, 1529436675, 1551995703, 1887438634],
    "ALVEOLE": [16665500, 14809000, 16566000, 13418500, 14725000, 14950000,
                 13092500, 11717500, 12227000, 15972500, 16211500, 19715000],
    "DIVERS": [862848, 767535, 856411, 712164, 767535, 781729,
                685180, 638802, 649118, 830665, 839659, 1010975],
    "Innovations": [0]*12,
}

# Compute S1 and S2 CA objectives
ca_obj_S1 = {cat: sum(CA_OBJ_MONTHLY[cat][:6]) for cat in CA_OBJ_MONTHLY}
ca_obj_S2 = {cat: sum(CA_OBJ_MONTHLY[cat][6:]) for cat in CA_OBJ_MONTHLY}
ca_obj_annual = {cat: sum(CA_OBJ_MONTHLY[cat]) for cat in CA_OBJ_MONTHLY}

# ===== LOAD ACTUAL SALES (no exclusion, with Maïs 50kg + P109 25kg) =====
with open("scripts/ventes_objectifs_data.json", "r") as f:
    sales = json.load(f)

# Actual CA per category (S1)
ca_real_S1 = {}
for cat in sales["category_sales_ca"]:
    ca_real_S1[cat] = sum(sales["category_sales_ca"][cat].get(str(m), 0) for m in range(1, 7))

# ===== COMPARISON =====
CATEGORIES = ["TOURTEAUX", "CONCENTRES", "ALIMENT COMPLET", "INGREDIENTS",
              "COMPLEMENT ALIMENTAIRE", "PREMIX", "MATERIEL ELEVAGE", "ALVEOLE", "DIVERS", "Innovations"]

print("=" * 100)
print("VENTES vs OBJECTIFS CA RÉELS (S1 2026)")
print("=" * 100)
print(f"\n{'Catégorie':<28}{'CA obj réel S1':>16}{'CA réel S1':>14}{'Écart (M)':>10}{'% atteinte':>12}")
print("-" * 80)

total_obj = 0
total_real = 0
comparison = []
for cat in CATEGORIES:
    obj = ca_obj_S1.get(cat, 0)
    real = ca_real_S1.get(cat, 0)
    ecart = real - obj
    pct = (real / obj * 100) if obj > 0 else 0
    total_obj += obj
    total_real += real
    comparison.append({"category": cat, "ca_obj_S1": obj, "ca_real_S1": real, "pct": pct})
    marker = " ⚠️" if pct < 90 and obj > 0 else ""
    print(f"{cat:<28}{obj/1e6:>16.1f}{real/1e6:>14.1f}{ecart/1e6:>+10.1f}{pct:>11.1f}%{marker}")

print("-" * 80)
print(f"{'TOTAL':<28}{total_obj/1e6:>16.1f}{total_real/1e6:>14.1f}{(total_real-total_obj)/1e6:>+10.1f}{total_real/total_obj*100:>11.1f}%")

# ===== S2 FORECAST CA =====
print(f"\n{'=' * 100}")
print("FORECAST CA S2 (basé sur YoY + CA objectifs réels)")
print(f"{'=' * 100}")

# Load YoY data
with open("scripts/yoy_forecast.json", "r") as f:
    yoy = json.load(f)

print(f"\n{'Catégorie':<28}{'CA obj S2 réel':>16}{'CA forecast S2':>16}{'Atteinte %':>12}")
print("-" * 72)
total_obj_s2 = 0
total_fc = 0
for cat in CATEGORIES:
    obj_s2 = ca_obj_S2.get(cat, 0)
    # Forecast = S2 2025 CA × (1 + YoY volume growth)
    # But we don't have S2 2025 CA easily. Use volume forecast × real price per ton.
    # Actually, better: use the ratio of S1 real CA / S1 obj CA, applied to S2 obj CA
    # This gives a CA-based forecast, not volume-based
    s1_ratio = ca_real_S1.get(cat, 0) / ca_obj_S1.get(cat, 1) if ca_obj_S1.get(cat, 0) > 0 else 1
    fc_ca = obj_s2 * s1_ratio
    pct = (fc_ca / obj_s2 * 100) if obj_s2 > 0 else 0
    total_obj_s2 += obj_s2
    total_fc += fc_ca
    print(f"{cat:<28}{obj_s2/1e6:>16.1f}{fc_ca/1e6:>16.1f}{pct:>11.1f}%")

print("-" * 72)
print(f"{'TOTAL':<28}{total_obj_s2/1e6:>16.1f}{total_fc/1e6:>16.1f}{total_fc/total_obj_s2*100:>11.1f}%")

# ===== SAVE DATA =====
output = {
    "ca_obj_monthly": CA_OBJ_MONTHLY,
    "ca_obj_S1": ca_obj_S1,
    "ca_obj_S2": ca_obj_S2,
    "ca_obj_annual": ca_obj_annual,
    "ca_real_S1": ca_real_S1,
    "comparison_S1": comparison,
    "totals": {
        "ca_obj_S1": total_obj,
        "ca_real_S1": total_real,
        "ca_obj_S2": total_obj_s2,
        "ca_forecast_S2": total_fc,
    }
}
with open("scripts/ca_obj_real.json", "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False, indent=2, default=str)
print(f"\nData saved to ca_obj_real.json")
