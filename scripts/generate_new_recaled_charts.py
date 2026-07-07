"""
Regenerate all charts that depend on the recalibrated S2 objectives.
"""
import os
import json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

plt.rcParams['font.sans-serif'] = ['Liberation Sans', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

OUT_DIR = '/home/z/my-project/scripts/pdf_charts'
NAVY = '#1F4E78'
NAVY_LIGHT = '#2E75B6'
GOLD = '#FFC000'
GOLD_DARK = '#BF8F00'
RED = '#C00000'
GREEN = '#548235'
GRAY = '#595959'
GRAY_LIGHT = '#A6A6A6'

# Load all data
with open('/home/z/my-project/scripts/s2_recaled_objectives.json', 'r') as f:
    recaled = json.load(f)
with open('/home/z/my-project/scripts/objectives_comparison.json', 'r') as f:
    old_obj = json.load(f)
with open('/home/z/my-project/scripts/yoy_forecast.json', 'r') as f:
    yoy = json.load(f)
with open('/home/z/my-project/scripts/ventes_objectifs_data.json', 'r') as f:
    sales = json.load(f)

prices = sales['price_per_ton']
MONTHS_S2 = [7, 8, 9, 10, 11, 12]

# ============ CHART I (UPDATED): Forecast vs Objectifs S2 (anciens + nouveaux recalibrés) ============
fig, ax = plt.subplots(figsize=(11, 5), constrained_layout=True)
cats = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT COMPLET', 'INGREDIENTS', 'COMPLEMENT ALIM.', 'PREMIX']
old_obj_vals = [sum(old_obj['global_objectives'].get(c.replace('ALIM.', 'ALIMENTAIRE') if c == 'COMPLEMENT ALIM.' else c, {}).get(str(m), 0) for m in MONTHS_S2) for c in cats]
new_obj_vals = [sum(recaled['global_s2_recaled'].get(c.replace('ALIM.', 'ALIMENTAIRE') if c == 'COMPLEMENT ALIM.' else c, {}).get(str(m), 0) for m in MONTHS_S2) for c in cats]
forecast_vals = [yoy['forecast_refined'].get(c.replace('ALIM.', 'ALIMENTAIRE') if c == 'COMPLEMENT ALIM.' else c, {}).get('real', 0) for c in cats]

x = np.arange(len(cats))
width = 0.25
b1 = ax.bar(x - width, old_obj_vals, width, label='Objectif initial S2', color=GRAY_LIGHT, edgecolor='white')
b2 = ax.bar(x, new_obj_vals, width, label='Objectif recalibré S2', color=NAVY, edgecolor='white')
b3 = ax.bar(x + width, forecast_vals, width, label='Forecast YoY (réaliste)', color=GOLD_DARK, edgecolor='white')

ax.set_xticks(x)
ax.set_xticklabels([c.replace(' ', '\n', 1) for c in cats], fontsize=8)
ax.set_ylabel('Volume S2 (tonnes)', fontsize=10)
ax.set_title('Forecast YoY vs Objectifs S2 (initial et recalibré) par catégorie', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='upper right', fontsize=8)
ax.grid(axis='y', alpha=0.3, linestyle='--')

plt.savefig(f'{OUT_DIR}/chartI_forecast_vs_recaled.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartI_forecast_vs_recaled.png (updated)")

# ============ CHART J (UPDATED): S2 recalibrated by agency ============
fig, ax = plt.subplots(figsize=(12, 5), constrained_layout=True)
agency_totals = {}
for ag in recaled['agency_s2_recaled']:
    total = sum(recaled['agency_s2_recaled'][ag][cat][str(m)] for cat in recaled['agency_s2_recaled'][ag] for m in MONTHS_S2)
    agency_totals[ag] = total

sorted_ag = sorted(agency_totals.items(), key=lambda x: -x[1])
ag_names = [a[0] for a in sorted_ag]
ag_vals = [a[1] for a in sorted_ag]

x = np.arange(len(ag_names))
bars = ax.bar(x, ag_vals, color=NAVY, edgecolor='white', linewidth=1)
ax.set_xticks(x)
ax.set_xticklabels(ag_names, fontsize=8, rotation=30, ha='right')
ax.set_ylabel('Volume S2 recalibré (tonnes)', fontsize=10)
ax.set_title('Objectifs S2 recalibrés par agence (Juillet-Décembre 2026)', fontsize=11, fontweight='bold', color=NAVY)
ax.grid(axis='y', alpha=0.3, linestyle='--')
for bar, val in zip(bars, ag_vals):
    ax.text(bar.get_x() + bar.get_width()/2., val + 50, f'{val:,.0f}'.replace(',', ' '), ha='center', va='bottom', fontsize=7)

plt.savefig(f'{OUT_DIR}/chartJ_s2_by_agence.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartJ_s2_by_agence.png (updated)")

# ============ NEW CHART O: CA impact of recalibration (before vs after) ============
fig, ax = plt.subplots(figsize=(10, 5), constrained_layout=True)
cats_ca = ['CONCENTRES', 'TOURTEAUX', 'INGREDIENTS', 'ALIMENT COMPLET', 'PREMIX', 'COMPLEMENT ALIM.']
ca_before = []
ca_after = []
for c in cats_ca:
    key = c.replace('ALIM.', 'ALIMENTAIRE') if c == 'COMPLEMENT ALIM.' else c
    vol_old = sum(old_obj['global_objectives'].get(key, {}).get(str(m), 0) for m in MONTHS_S2)
    vol_new = sum(recaled['global_s2_recaled'].get(key, {}).get(str(m), 0) for m in MONTHS_S2)
    price = prices.get(key, 0)
    ca_before.append(vol_old * price / 1e6)
    ca_after.append(vol_new * price / 1e6)

x = np.arange(len(cats_ca))
width = 0.35
b1 = ax.bar(x - width/2, ca_before, width, label='CA objectif initial', color=GRAY_LIGHT, edgecolor='white')
b2 = ax.bar(x + width/2, ca_after, width, label='CA objectif recalibré', color=NAVY, edgecolor='white')

ax.set_xticks(x)
ax.set_xticklabels([c.replace(' ', '\n', 1) for c in cats_ca], fontsize=8)
ax.set_ylabel('CA HT S2 (millions FCFA)', fontsize=10)
ax.set_title('Incidence du recalibrage sur le CA HT S2 par catégorie', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='upper right', fontsize=9)
ax.grid(axis='y', alpha=0.3, linestyle='--')

# Add delta labels
for i in range(len(cats_ca)):
    delta = ca_after[i] - ca_before[i]
    color = GREEN if delta >= 0 else RED
    max_val = max(ca_before[i], ca_after[i])
    ax.text(x[i], max_val + 100, f'{delta:+.0f} M', ha='center', va='bottom', fontsize=8, color=color, fontweight='bold')

plt.savefig(f'{OUT_DIR}/chartO_ca_impact_recaled.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartO_ca_impact_recaled.png (new)")

# ============ NEW CHART P: CONCENTRES - can we reach +5% to +10% vs 2025? ============
fig, ax = plt.subplots(figsize=(10, 5), constrained_layout=True)

# Data
total_2025 = 17926.1
s1_2026 = 8592.1
s2_2025 = 9256.2
forecast_real = yoy['forecast_refined']['CONCENTRES']['real']
forecast_pess = yoy['forecast_refined']['CONCENTRES']['pess']
forecast_opt = yoy['forecast_refined']['CONCENTRES']['opt']
new_obj_conc = sum(recaled['global_s2_recaled']['CONCENTRES'][str(m)] for m in MONTHS_S2)
target_5 = total_2025 * 1.05 - s1_2026  # S2 needed for +5%
target_10 = total_2025 * 1.10 - s1_2026  # S2 needed for +10%

labels = ['S2 2025\n(référence)', '🔴 Forecast\npessimiste', '🟡 Forecast\nréaliste', '🟢 Forecast\noptimiste', 'Objectif S2\nrecalibré', 'Cible +5%\nvs 2025', 'Cible +10%\nvs 2025']
values = [s2_2025, forecast_pess, forecast_real, forecast_opt, new_obj_conc, target_5, target_10]
colors_bar = [GRAY_LIGHT, RED, GOLD_DARK, GREEN, NAVY, '#70AD47', '#375623']

bars = ax.bar(labels, values, color=colors_bar, edgecolor='white', linewidth=1.5)
ax.set_ylabel('Volume S2 CONCENTRÉS (tonnes)', fontsize=10)
ax.set_title('CONCENTRÉS — Peut-on boucler +5% à +10% vs 2025 ?', fontsize=11, fontweight='bold', color=NAVY)
ax.grid(axis='y', alpha=0.3, linestyle='--')

# Add value labels
for bar, val in zip(bars, values):
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., h + 100, f'{val:,.0f}'.replace(',', ' '), ha='center', va='bottom', fontsize=8, fontweight='bold')

# Horizontal line for S2 2025 reference
ax.axhline(y=s2_2025, color=GRAY, linewidth=1, linestyle='--', alpha=0.5)

ax.set_ylim(0, max(values) * 1.15)
plt.xticks(fontsize=8)

plt.savefig(f'{OUT_DIR}/chartP_conc_target_5_10pct.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartP_conc_target_5_10pct.png (new)")

print(f"\nAll updated charts saved to {OUT_DIR}")
