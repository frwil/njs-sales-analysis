"""Generate 4 new charts for agency + monthly trend analysis."""
import os, json
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

with open('/home/z/my-project/scripts/agency_monthly_analysis.json', 'r') as f:
    data = json.load(f)

# ============ CHART Q: Ventes vs Objectifs par agence ============
fig, ax = plt.subplots(figsize=(12, 5.5), constrained_layout=True)
ag_data = [(d['agence'], d['s1_obj'], d['s1_real'], d['pct']) for d in data['agency_comparison']
           if d['s1_obj'] > 0 and d['s1_real'] > 0]
ag_data.sort(key=lambda x: -x[1])

names = [d[0] for d in ag_data]
objs = [d[1] for d in ag_data]
reals = [d[2] for d in ag_data]
pcts = [d[3] for d in ag_data]

x = np.arange(len(names))
width = 0.35
b1 = ax.bar(x - width/2, objs, width, label='Objectif S1', color=NAVY_LIGHT, edgecolor='white')
b2 = ax.bar(x + width/2, reals, width, label='Ventes réelles S1', color=GOLD_DARK, edgecolor='white')

ax.set_xticks(x)
ax.set_xticklabels(names, fontsize=7, rotation=30, ha='right')
ax.set_ylabel('Volume (tonnes)', fontsize=10)
ax.set_title('Ventes vs Objectifs S1 2026 par agence (tonnes)', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='upper right', fontsize=9)
ax.grid(axis='y', alpha=0.3, linestyle='--')

for i, pct in enumerate(pcts):
    color = GREEN if pct >= 100 else (GOLD_DARK if pct >= 80 else RED)
    max_val = max(objs[i], reals[i])
    ax.text(x[i], max_val + 50, f'{pct:.0f}%', ha='center', va='bottom', fontsize=7, color=color, fontweight='bold')

plt.savefig(f'{OUT_DIR}/chartQ_ventes_vs_obj_agence.png', dpi=150, bbox_inches='tight')
plt.close()
print("  OK chartQ_ventes_vs_obj_agence.png")

# ============ CHART R: CONCENTRES par agence ============
fig, ax = plt.subplots(figsize=(12, 5), constrained_layout=True)
conc_data = data['conc_by_agency']
# Remove Pk11 (obj=0)
conc_data = [d for d in conc_data if d['obj'] > 0]
conc_data.sort(key=lambda x: -x['vol'])

names = [d['agence'] for d in conc_data]
vols = [d['vol'] for d in conc_data]
objs = [d['obj'] for d in conc_data]
pcts = [d['pct'] for d in conc_data]

x = np.arange(len(names))
width = 0.35
b1 = ax.bar(x - width/2, objs, width, label='Objectif S1', color=NAVY_LIGHT, edgecolor='white')
b2 = ax.bar(x + width/2, vols, width, label='Ventes réelles S1', color=RED, edgecolor='white')

ax.set_xticks(x)
ax.set_xticklabels(names, fontsize=7, rotation=30, ha='right')
ax.set_ylabel('Volume CONCENTRES (tonnes)', fontsize=10)
ax.set_title('CONCENTRES — Ventes vs Objectifs S1 2026 par agence', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='upper right', fontsize=9)
ax.grid(axis='y', alpha=0.3, linestyle='--')

for i, pct in enumerate(pcts):
    color = GREEN if pct >= 90 else (GOLD_DARK if pct >= 70 else RED)
    max_val = max(objs[i], vols[i])
    ax.text(x[i], max_val + 20, f'{pct:.0f}%', ha='center', va='bottom', fontsize=7, color=color, fontweight='bold')

plt.savefig(f'{OUT_DIR}/chartR_conc_by_agence.png', dpi=150, bbox_inches='tight')
plt.close()
print("  OK chartR_conc_by_agence.png")

# ============ CHART S: Tendance mensuelle par catégorie ============
fig, ax = plt.subplots(figsize=(11, 5), constrained_layout=True)
cats = ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'ALIMENT COMPLET', 'COMPLEMENT ALIM.', 'PREMIX']
cat_keys = ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'ALIMENT COMPLET', 'COMPLEMENT ALIMENTAIRE', 'PREMIX']
colors_s = [NAVY, RED, GREEN, GOLD_DARK, NAVY_LIGHT, GRAY]
markers = ['o', 's', '^', 'D', 'v', 'P']

months_labels = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin']
for i, (cat, key) in enumerate(zip(cats, cat_keys)):
    vals = [data['cat_month_2026'].get(key, {}).get(str(m), {}).get('vol_t', 0) for m in range(1, 7)]
    ax.plot(range(6), vals, marker=markers[i], linewidth=2.5, markersize=7, color=colors_s[i], label=cat)

ax.set_xticks(range(6))
ax.set_xticklabels(months_labels, fontsize=9)
ax.set_ylabel('Volume (tonnes)', fontsize=10)
ax.set_title('Tendance mensuelle par catégorie (Janvier-Juin 2026)', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='best', fontsize=8)
ax.grid(axis='y', alpha=0.3, linestyle='--')
ax.set_yscale('log')

plt.savefig(f'{OUT_DIR}/chartS_monthly_by_cat.png', dpi=150, bbox_inches='tight')
plt.close()
print("  OK chartS_monthly_by_cat.png")

# ============ CHART T: Tendance mensuelle par agence (top 5) ============
fig, ax = plt.subplots(figsize=(11, 5), constrained_layout=True)
top5 = data['top5_agencies']
colors_t = [NAVY, RED, GREEN, GOLD_DARK, NAVY_LIGHT]
markers_t = ['o', 's', '^', 'D', 'v']

for i, ag in enumerate(top5):
    vals = [data['agency_month_2026'].get(ag, {}).get(str(m), {}).get('vol_t', 0) for m in range(1, 7)]
    ax.plot(range(6), vals, marker=markers_t[i], linewidth=2.5, markersize=7, color=colors_t[i], label=ag)

ax.set_xticks(range(6))
ax.set_xticklabels(months_labels, fontsize=9)
ax.set_ylabel('Volume total (tonnes)', fontsize=10)
ax.set_title('Tendance mensuelle par agence — Top 5 (Janvier-Juin 2026)', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='best', fontsize=9)
ax.grid(axis='y', alpha=0.3, linestyle='--')

plt.savefig(f'{OUT_DIR}/chartT_monthly_by_agence.png', dpi=150, bbox_inches='tight')
plt.close()
print("  OK chartT_monthly_by_agence.png")

print(f"\nAll 4 new charts saved to {OUT_DIR}")
