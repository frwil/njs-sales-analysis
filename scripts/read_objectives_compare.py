"""
Read objectives from the objectifs file, per category x month (global) and per agence x category x month.
Then build comparison vs actual sales and forecast S2.
"""
import os
import json
from openpyxl import load_workbook

SRC_OBJECTIFS = "/home/z/my-project/upload/objectifs 2026 - Takou.xlsx"
DATA_FILE = "/home/z/my-project/scripts/ventes_objectifs_data.json"

# Load sales data
with open(DATA_FILE, "r", encoding="utf-8") as f:
    sales_data = json.load(f)

# ===== READ OBJECTIVES =====
print("Reading objectives file...")
wb = load_workbook(SRC_OBJECTIFS, data_only=True)
ws = wb.active

# Categories list (rows 5-14 in global section, same structure repeats per agence)
CATEGORIES = ["ALIMENT COMPLET", "INGREDIENTS", "COMPLEMENT ALIMENTAIRE", "MATERIEL ELEVAGE",
              "PREMIX", "CONCENTRES", "TOURTEAUX", "ALVEOLE", "DIVERS", "Innovations"]
MONTHS = list(range(1, 13))  # 1=Jan ... 12=Dec

# Global objectives: rows 5-14 (10 categories)
global_objectives = {}  # cat -> {month: tons}
for i, cat in enumerate(CATEGORIES):
    r = 5 + i
    global_objectives[cat] = {}
    for j, m in enumerate(MONTHS):
        v = ws.cell(row=r, column=5 + j).value  # E=P(5+11=16)
        global_objectives[cat][m] = float(v) if v else 0.0

# Print global objectives
print("\nGlobal objectives 2026 (tons) per category x month:")
print(f"{'Category':<28}", end='')
for m in range(1, 13):
    print(f"  M{m:>2}", end='')
print(f"{'TOTAL':>10}{'Q1':>8}{'Q2':>8}{'S1':>8}{'S2':>8}")
for cat in CATEGORIES:
    print(f"{cat[:28]:<28}", end='')
    total = 0
    q1 = sum(global_objectives[cat][m] for m in [1,2,3])
    q2 = sum(global_objectives[cat][m] for m in [4,5,6])
    s2 = sum(global_objectives[cat][m] for m in [7,8,9,10,11,12])
    for m in range(1, 13):
        v = global_objectives[cat][m]
        total += v
        print(f"{v:>6.0f}", end='')
    print(f"{total:>10.0f}{q1:>8.0f}{q2:>8.0f}{q1+q2:>8.0f}{s2:>8.0f}")

# Read agence-level objectives
# Structure: every 12 rows, starting at row 16
# Row 16: Total Village (agence header), Row 17-26: 10 categories
# Agence name is in column A of the "Total" row
agence_objectives = {}  # agence -> {cat -> {month -> tons}}
for block_start in range(16, 327, 12):
    agence_cell = ws.cell(row=block_start, column=1).value
    if agence_cell is None: continue
    agence_name = str(agence_cell).replace("Total ", "").strip()
    if not agence_name or agence_name == "0": continue
    agence_objectives[agence_name] = {}
    for i, cat in enumerate(CATEGORIES):
        r = block_start + 1 + i
        agence_objectives[agence_name][cat] = {}
        for j, m in enumerate(MONTHS):
            v = ws.cell(row=r, column=5 + j).value
            agence_objectives[agence_name][cat][m] = float(v) if v else 0.0

print(f"\nAgences in objectives file: {len(agence_objectives)}")
for ag in sorted(agence_objectives.keys()):
    total_q1 = sum(agence_objectives[ag][cat][m] for cat in CATEGORIES for m in [1,2,3])
    total_q2 = sum(agence_objectives[ag][cat][m] for cat in CATEGORIES for m in [4,5,6])
    print(f"  {ag:<25} Q1={total_q1:>8.1f} t  Q2={total_q2:>8.1f} t  S1={total_q1+total_q2:>8.1f} t")

# ===== BUILD COMPARISON SALES VS OBJECTIVES (Q1, Q2, S1) =====
print("\n" + "="*80)
print("COMPARISON: Actual Sales vs Objectives (in tons)")
print("="*80)
print(f"\n{'Category':<28}{'Q1 obj':>10}{'Q1 real':>10}{'% Q1':>8}{'Q2 obj':>10}{'Q2 real':>10}{'% Q2':>8}{'S1 obj':>10}{'S1 real':>10}{'% S1':>8}")

comparison_data = []
for cat in CATEGORIES:
    q1_obj = sum(global_objectives[cat][m] for m in [1,2,3])
    q2_obj = sum(global_objectives[cat][m] for m in [4,5,6])
    s1_obj = q1_obj + q2_obj

    q1_real = sum(sales_data["category_sales_tons"].get(cat, {}).get(str(m), 0) for m in [1,2,3])
    q2_real = sum(sales_data["category_sales_tons"].get(cat, {}).get(str(m), 0) for m in [4,5,6])
    s1_real = q1_real + q2_real

    pct_q1 = (q1_real / q1_obj * 100) if q1_obj > 0 else 0
    pct_q2 = (q2_real / q2_obj * 100) if q2_obj > 0 else 0
    pct_s1 = (s1_real / s1_obj * 100) if s1_obj > 0 else 0

    print(f"{cat[:28]:<28}{q1_obj:>10.1f}{q1_real:>10.1f}{pct_q1:>7.1f}%{q2_obj:>10.1f}{q2_real:>10.1f}{pct_q2:>7.1f}%{s1_obj:>10.1f}{s1_real:>10.1f}{pct_s1:>7.1f}%")
    comparison_data.append({
        "category": cat,
        "q1_obj": q1_obj, "q1_real": q1_real, "pct_q1": pct_q1,
        "q2_obj": q2_obj, "q2_real": q2_real, "pct_q2": pct_q2,
        "s1_obj": s1_obj, "s1_real": s1_real, "pct_s1": pct_s1,
        "price_per_ton": sales_data["price_per_ton"].get(cat, 0),
    })

# Totals
q1_obj_tot = sum(global_objectives[cat][m] for cat in CATEGORIES for m in [1,2,3])
q2_obj_tot = sum(global_objectives[cat][m] for cat in CATEGORIES for m in [4,5,6])
s1_obj_tot = q1_obj_tot + q2_obj_tot
q1_real_tot = sum(sales_data["category_sales_tons"].get(cat, {}).get(str(m), 0) for cat in CATEGORIES for m in [1,2,3])
q2_real_tot = sum(sales_data["category_sales_tons"].get(cat, {}).get(str(m), 0) for cat in CATEGORIES for m in [4,5,6])
s1_real_tot = q1_real_tot + q2_real_tot
print(f"{'TOTAL':<28}{q1_obj_tot:>10.1f}{q1_real_tot:>10.1f}{q1_real_tot/q1_obj_tot*100:>7.1f}%{q2_obj_tot:>10.1f}{q2_real_tot:>10.1f}{q2_real_tot/q2_obj_tot*100:>7.1f}%{s1_obj_tot:>10.1f}{s1_real_tot:>10.1f}{s1_real_tot/s1_obj_tot*100:>7.1f}%")

# Save comparison + objectives data
output_data = {
    "global_objectives": {cat: {str(m): global_objectives[cat][m] for m in range(1,13)} for cat in CATEGORIES},
    "agence_objectives": {ag: {cat: {str(m): agence_objectives[ag][cat][m] for m in range(1,13)} for cat in CATEGORIES} for ag in agence_objectives},
    "comparison": comparison_data,
    "totals": {
        "q1_obj": q1_obj_tot, "q1_real": q1_real_tot,
        "q2_obj": q2_obj_tot, "q2_real": q2_real_tot,
        "s1_obj": s1_obj_tot, "s1_real": s1_real_tot,
    }
}

with open("/home/z/my-project/scripts/objectives_comparison.json", "w", encoding="utf-8") as f:
    json.dump(output_data, f, ensure_ascii=False, indent=2, default=str)
print(f"\nData saved to /home/z/my-project/scripts/objectives_comparison.json")
