"""Generate CONCENTRES deep-dive charts."""
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

with open('/home/z/my-project/scripts/concentres_deepdive.json', 'r') as f:
    data = json.load(f)

# ============ CHART K: CONCENTRÉS by sub-category (volume + CA) ============
fig, axes = plt.subplots(1, 2, figsize=(13, 4.5), constrained_layout=True)
subcats = ['BELGO 10% Chair', 'BELGO 5% Ponte', 'BELGO 10% Porc', 'BELGO 5% Chair', 'BELGO 10% Ponte', 'BELGO Rabbit']
vols = []
cas = []
for sc in subcats:
    v = sum(data['subcat_2026'].get(sc, {}).get(str(m), {}).get('vol_t', 0) for m in range(1, 7))
    c = sum(data['subcat_2026'].get(sc, {}).get(str(m), {}).get('ca', 0) for m in range(1, 7))
    vols.append(v)
    cas.append(c / 1e6)

# Volume
colors_k = [NAVY, NAVY_LIGHT, RED, GOLD_DARK, GRAY, GREEN]
axes[0].barh(range(len(subcats)), vols, color=colors_k, edgecolor='white')
axes[0].set_yticks(range(len(subcats)))
axes[0].set_yticklabels([s.replace('BELGO ', '') for s in subcats], fontsize=8)
axes[0].invert_yaxis()
axes[0].set_xlabel('Volume S1 2026 (tonnes)', fontsize=9)
axes[0].set_title('Volume par sous-catégorie', fontsize=10, fontweight='bold', color=NAVY)
axes[0].grid(axis='x', alpha=0.3, linestyle='--')
for i, v in enumerate(vols):
    axes[0].text(v + 50, i, f'{v:.0f} t', va='center', fontsize=8)

# CA
axes[1].barh(range(len(subcats)), cas, color=colors_k, edgecolor='white')
axes[1].set_yticks(range(len(subcats)))
axes[1].set_yticklabels([s.replace('BELGO ', '') for s in subcats], fontsize=8)
axes[1].invert_yaxis()
axes[1].set_xlabel('CA S1 2026 (millions FCFA)', fontsize=9)
axes[1].set_title('CA HT par sous-catégorie', fontsize=10, fontweight='bold', color=NAVY)
axes[1].grid(axis='x', alpha=0.3, linestyle='--')
for i, c in enumerate(cas):
    axes[1].text(c + 20, i, f'{c:.0f} M', va='center', fontsize=8)

plt.suptitle('CONCENTRÉS — Décomposition par sous-catégorie (S1 2026)', fontsize=11, fontweight='bold', color=NAVY, y=1.02)
plt.savefig(f'{OUT_DIR}/chartK_conc_subcat.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartK_conc_subcat.png")

# ============ CHART L: CONCENTRÉS YoY by sub-category ============
fig, ax = plt.subplots(figsize=(10, 5), constrained_layout=True)
subcats_yoy = ['BELGO 10% Chair', 'BELGO 5% Ponte', 'BELGO 10% Porc', 'BELGO 5% Chair', 'BELGO 10% Ponte']
s1_2025_vals = []
s1_2026_vals = []
yoy_pcts = []
for sc in subcats_yoy:
    s25 = sum(data['subcat_2025'].get(sc, {}).get(str(m), {}).get('vol_t', 0) for m in range(1, 7))
    s26 = sum(data['subcat_2026'].get(sc, {}).get(str(m), {}).get('vol_t', 0) for m in range(1, 7))
    s1_2025_vals.append(s25)
    s1_2026_vals.append(s26)
    yoy = ((s26 - s25) / s25 * 100) if s25 > 0 else 0
    yoy_pcts.append(yoy)

x = np.arange(len(subcats_yoy))
width = 0.35
b1 = ax.bar(x - width/2, s1_2025_vals, width, label='S1 2025', color=GRAY_LIGHT, edgecolor='white')
b2 = ax.bar(x + width/2, s1_2026_vals, width, label='S1 2026', color=NAVY, edgecolor='white')

ax.set_xticks(x)
ax.set_xticklabels([s.replace('BELGO ', '') for s in subcats_yoy], fontsize=8, rotation=15)
ax.set_ylabel('Volume (tonnes)', fontsize=10)
ax.set_title('CONCENTRÉS — Croissance YoY par sous-catégorie (S1 2025 vs S1 2026)', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='upper right', fontsize=9)
ax.grid(axis='y', alpha=0.3, linestyle='--')

for i, yoy in enumerate(yoy_pcts):
    color = GREEN if yoy > 0 else RED
    max_val = max(s1_2025_vals[i], s1_2026_vals[i])
    ax.text(x[i], max_val + 100, f'{yoy:+.1f}%', ha='center', va='bottom', fontsize=9, color=color, fontweight='bold')

plt.savefig(f'{OUT_DIR}/chartL_conc_yoy.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartL_conc_yoy.png")

# ============ CHART M: CONCENTRÉS monthly evolution 2026 ============
fig, ax = plt.subplots(figsize=(10, 5), constrained_layout=True)
months_labels = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin']
subcats_monthly = ['BELGO 10% Chair', 'BELGO 5% Ponte', 'BELGO 10% Porc', 'BELGO 5% Chair']
colors_m = [NAVY, GREEN, RED, GOLD_DARK]
markers = ['o', 's', '^', 'D']

for i, sc in enumerate(subcats_monthly):
    vals = [data['subcat_2026'].get(sc, {}).get(str(m), {}).get('vol_t', 0) for m in range(1, 7)]
    ax.plot(range(6), vals, marker=markers[i], linewidth=2.5, markersize=7, color=colors_m[i], label=sc.replace('BELGO ', ''))

ax.set_xticks(range(6))
ax.set_xticklabels(months_labels, fontsize=9)
ax.set_ylabel('Volume (tonnes)', fontsize=10)
ax.set_title('CONCENTRÉS — Évolution mensuelle 2026 par sous-catégorie', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='best', fontsize=8)
ax.grid(axis='y', alpha=0.3, linestyle='--')

plt.savefig(f'{OUT_DIR}/chartM_conc_monthly.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartM_conc_monthly.png")

# ============ CHART N: CONCENTRÉS by agence ============
fig, ax = plt.subplots(figsize=(11, 5), constrained_layout=True)
agencies = []
ag_vols = []
for ag in sorted(data['agency_conc_2026'].keys(),
                  key=lambda a: -sum(data['agency_conc_2026'][a].get(str(m), {}).get('vol_t', 0) for m in range(1, 7))):
    vol = sum(data['agency_conc_2026'][ag].get(str(m), {}).get('vol_t', 0) for m in range(1, 7))
    if vol > 0:
        agencies.append(ag)
        ag_vols.append(vol)

x = np.arange(len(agencies))
bars = ax.bar(x, ag_vols, color=NAVY, edgecolor='white', linewidth=1)
ax.set_xticks(x)
ax.set_xticklabels(agencies, fontsize=8, rotation=30, ha='right')
ax.set_ylabel('Volume S1 2026 (tonnes)', fontsize=10)
ax.set_title('CONCENTRÉS — Volume par agence (S1 2026)', fontsize=11, fontweight='bold', color=NAVY)
ax.grid(axis='y', alpha=0.3, linestyle='--')
for bar, val in zip(bars, ag_vols):
    ax.text(bar.get_x() + bar.get_width()/2., val + 20, f'{val:.0f}', ha='center', va='bottom', fontsize=7)

plt.savefig(f'{OUT_DIR}/chartN_conc_by_agence.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartN_conc_by_agence.png")

print(f"\nAll CONCENTRÉS charts saved to {OUT_DIR}")
