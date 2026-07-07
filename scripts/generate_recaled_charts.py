"""Generate chart for recalibrated S2 objectives vs forecast."""
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
GOLD_DARK = '#BF8F00'
RED = '#C00000'
GREEN = '#548235'
GRAY_LIGHT = '#A6A6A6'

with open('/home/z/my-project/scripts/s2_recaled_objectives.json', 'r') as f:
    s2_data = json.load(f)
with open('/home/z/my-project/scripts/yoy_forecast.json', 'r') as f:
    yoy_data = json.load(f)

# ============ CHART I: Forecast YoY vs Objectifs S2 Recalibrés ============
fig, ax = plt.subplots(figsize=(11, 5), constrained_layout=True)
cats = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT COMPLET', 'INGREDIENTS', 'COMPLEMENT ALIM.', 'PREMIX']
old_obj = [26043, 12242, 501, 585, 6, 59]  # old S2 objectives
recaled_obj = [sum(s2_data['global_s2_recaled'][c.replace('ALIM.', 'ALIMENTAIRE')][str(m)] for m in range(7, 13)) if c == 'COMPLEMENT ALIM.' else sum(s2_data['global_s2_recaled'][c][str(m)] for m in range(7, 13)) for c in cats]
forecast = [yoy_data['forecast_refined'].get(c.replace('ALIM.', 'ALIMENTAIRE') if c == 'COMPLEMENT ALIM.' else c, {}).get('real', 0) for c in cats]

x = np.arange(len(cats))
width = 0.25
b1 = ax.bar(x - width, old_obj, width, label='Ancien objectif S2', color=GRAY_LIGHT, edgecolor='white')
b2 = ax.bar(x, recaled_obj, width, label='Objectif S2 recalibré', color=NAVY, edgecolor='white')
b3 = ax.bar(x + width, forecast, width, label='Forecast YoY (réaliste)', color=GOLD_DARK, edgecolor='white')

ax.set_xticks(x)
ax.set_xticklabels([c.replace(' ', '\n', 1) for c in cats], fontsize=8)
ax.set_ylabel('Volume S2 (tonnes)', fontsize=10)
ax.set_title('Forecast YoY vs Objectifs S2 (anciens et recalibrés) par catégorie', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='upper right', fontsize=8)
ax.grid(axis='y', alpha=0.3, linestyle='--')

plt.savefig(f'{OUT_DIR}/chartI_forecast_vs_recaled.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartI_forecast_vs_recaled.png")

# ============ CHART J: Par agence - S2 recalibré vs S1 réel ============
fig, ax = plt.subplots(figsize=(12, 5), constrained_layout=True)
# Compute S2 recalibrated total per agency
agency_s2_totals = {}
for ag in s2_data['agency_s2_recaled']:
    total = sum(s2_data['agency_s2_recaled'][ag][cat][str(m)] for cat in s2_data['agency_s2_recaled'][ag] for m in range(7, 13))
    agency_s2_totals[ag] = total

# Sort by S2 objective descending
sorted_ag = sorted(agency_s2_totals.items(), key=lambda x: -x[1])
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
print("  ✓ chartJ_s2_by_agence.png")

print(f"\nAll charts saved to {OUT_DIR}")
