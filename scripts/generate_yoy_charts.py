"""
Generate charts for the YoY 2025 vs 2026 analysis.
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

with open('/home/z/my-project/scripts/yoy_forecast.json', 'r') as f:
    data = json.load(f)

# ============ CHART F: YoY Growth S1 2025 vs S1 2026 ============
fig, ax = plt.subplots(figsize=(11, 5), constrained_layout=True)
cats = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT COMPLET', 'INGREDIENTS', 'COMPLEMENT ALIMENTAIRE', 'PREMIX']
s1_2025 = [data['yoy_data'][c]['s1_2025'] for c in cats]
s1_2026 = [data['yoy_data'][c]['s1_2026'] for c in cats]

x = np.arange(len(cats))
width = 0.35
b1 = ax.bar(x - width/2, s1_2025, width, label='S1 2025', color=GRAY_LIGHT, edgecolor='white')
b2 = ax.bar(x + width/2, s1_2026, width, label='S1 2026', color=NAVY, edgecolor='white')

ax.set_xticks(x)
ax.set_xticklabels([c.replace(' ', '\n', 1) for c in cats], fontsize=8)
ax.set_ylabel('Volume (tonnes)', fontsize=10)
ax.set_title('Comparaison YoY : S1 2025 vs S1 2026 par catégorie', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='upper right', fontsize=9)
ax.grid(axis='y', alpha=0.3, linestyle='--')

# Add YoY % labels
for i, cat in enumerate(cats):
    yoy = data['yoy_data'][cat]['yoy_pct']
    color = GREEN if yoy > 0 else RED
    max_val = max(s1_2025[i], s1_2026[i])
    ax.text(x[i], max_val + max(s1_2025 + s1_2026)*0.02, f'{yoy:+.1f}%', ha='center', va='bottom', fontsize=9, color=color, fontweight='bold')

plt.savefig(f'{OUT_DIR}/chartF_yoy_2025_2026.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartF_yoy_2025_2026.png")

# ============ CHART G: Évolution mensuelle 2025 (full year) vs 2026 S1 ============
fig, axes = plt.subplots(1, 2, figsize=(13, 4.5), constrained_layout=True)
months_labels_12 = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc']

for idx, cat in enumerate(['TOURTEAUX', 'CONCENTRES']):
    ax = axes[idx]
    # 2025 full year (12 months)
    v_2025 = [data['category_2025_monthly'].get(cat, {}).get(str(m), {}).get('vol_t', 0) for m in range(1, 13)]
    # 2026 S1 (6 months)
    v_2026 = []
    with open('/home/z/my-project/scripts/ventes_objectifs_data.json', 'r') as f:
        sales_data = json.load(f)
    v_2026 = [sales_data['category_sales_tons'].get(cat, {}).get(str(m), 0) for m in range(1, 7)]

    x_2025 = np.arange(12)
    x_2026 = np.arange(6)

    ax.plot(x_2025, v_2025, marker='o', linewidth=2.5, markersize=7, color=GRAY, label='2025')
    ax.plot(x_2026, v_2026, marker='s', linewidth=2.5, markersize=7, color=NAVY, label='2026 S1')
    # Projection line for S2 2026 (realist scenario)
    s2_forecast = data['forecast_refined'].get(cat, {})
    if s2_forecast:
        # Distribute S2 forecast across months using 2025 seasonality
        s2_2025_monthly = [data['category_2025_monthly'].get(cat, {}).get(str(m), {}).get('vol_t', 0) for m in range(7, 13)]
        s2_2025_total = sum(s2_2025_monthly)
        if s2_2025_total > 0:
            # Realist scenario
            real_total = s2_forecast['real']
            v_forecast = [real_total * (v / s2_2025_total) for v in s2_2025_monthly]
            x_forecast = np.arange(6, 12)
            ax.plot(x_forecast, v_forecast, marker='^', linewidth=2.5, markersize=7, color=GOLD_DARK, label='2026 S2 (forecast realiste)', linestyle='--')

    ax.set_xticks(range(12))
    ax.set_xticklabels(months_labels_12, fontsize=8)
    ax.set_ylabel('Tonnes', fontsize=9)
    ax.set_title(f'{cat}', fontsize=11, fontweight='bold', color=NAVY)
    ax.legend(fontsize=7, loc='best')
    ax.grid(axis='y', alpha=0.3, linestyle='--')

plt.suptitle('Évolution mensuelle 2025 (année complète) vs 2026 (S1 réel + S2 forecast)', fontsize=11, fontweight='bold', color=NAVY, y=1.02)
plt.savefig(f'{OUT_DIR}/chartG_evolution_2025_2026.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartG_evolution_2025_2026.png")

# ============ CHART H: Forecast S2 refined (3 scenarios) vs objectif ============
fig, ax = plt.subplots(figsize=(11, 5), constrained_layout=True)
cats_f = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT COMPLET', 'INGREDIENTS', 'COMPLEMENT ALIMENTAIRE', 'PREMIX']
s2_2025_vals = [data['forecast_refined'][c]['s2_2025'] for c in cats_f]
pess_vals = [data['forecast_refined'][c]['pess'] for c in cats_f]
real_vals = [data['forecast_refined'][c]['real'] for c in cats_f]
opt_vals = [data['forecast_refined'][c]['opt'] for c in cats_f]
s2_obj_vals = [data['forecast_refined'][c]['s2_obj'] for c in cats_f]

x = np.arange(len(cats_f))
width = 0.18
b0 = ax.bar(x - 2*width, s2_2025_vals, width, label='S2 2025 (réf.)', color=GRAY_LIGHT, edgecolor='white')
b1 = ax.bar(x - width, pess_vals, width, label='Pessimiste', color=RED, edgecolor='white')
b2 = ax.bar(x, real_vals, width, label='Realiste', color=GOLD_DARK, edgecolor='white')
b3 = ax.bar(x + width, opt_vals, width, label='Optimiste', color=GREEN, edgecolor='white')
b4 = ax.bar(x + 2*width, s2_obj_vals, width, label='Objectif S2', color=NAVY, edgecolor='white')

ax.set_xticks(x)
ax.set_xticklabels([c.replace(' ', '\n', 1) for c in cats_f], fontsize=8)
ax.set_ylabel('Volume S2 (tonnes)', fontsize=10)
ax.set_title('Forecast S2 2026 affiné (basé sur 2025 + YoY) vs Objectif', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='upper right', fontsize=8)
ax.grid(axis='y', alpha=0.3, linestyle='--')

plt.savefig(f'{OUT_DIR}/chartH_forecast_refined.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartH_forecast_refined.png")

print(f"\nAll charts saved to {OUT_DIR}")
