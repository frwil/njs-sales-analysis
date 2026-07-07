"""
Generate charts for the Ventes vs Objectifs + Forecast sections.
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
os.makedirs(OUT_DIR, exist_ok=True)

NAVY = '#1F4E78'
NAVY_LIGHT = '#2E75B6'
GOLD = '#FFC000'
GOLD_DARK = '#BF8F00'
RED = '#C00000'
GREEN = '#548235'
GRAY = '#595959'
GRAY_LIGHT = '#A6A6A6'

# Load data
with open('/home/z/my-project/scripts/objectives_comparison.json', 'r') as f:
    data = json.load(f)

# ============ CHART A: Ventes vs Objectifs S1 par catégorie ============
fig, ax = plt.subplots(figsize=(10, 5), constrained_layout=True)
cats = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT COMPLET', 'INGREDIENTS', 'COMPLEMENT ALIMENTAIRE', 'PREMIX']
obj_s1 = [data['global_objectives'][c]['1'] + data['global_objectives'][c]['2'] + data['global_objectives'][c]['3']
          + data['global_objectives'][c]['4'] + data['global_objectives'][c]['5'] + data['global_objectives'][c]['6']
          for c in cats]
real_s1 = [next((d['s1_real'] for d in data['comparison'] if d['category'] == c), 0) for c in cats]

x = np.arange(len(cats))
width = 0.35
b1 = ax.bar(x - width/2, obj_s1, width, label='Objectif S1', color=NAVY_LIGHT, edgecolor='white')
b2 = ax.bar(x + width/2, real_s1, width, label='Ventes réelles S1', color=GOLD_DARK, edgecolor='white')

ax.set_xticks(x)
ax.set_xticklabels([c.replace(' ', '\n', 1) for c in cats], fontsize=8)
ax.set_ylabel('Volume (tonnes)', fontsize=10)
ax.set_title('Ventes vs Objectifs S1 2026 par catégorie (en tonnes)', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='upper right', fontsize=9)
ax.grid(axis='y', alpha=0.3, linestyle='--')

for i, (obj, real) in enumerate(zip(obj_s1, real_s1)):
    pct = (real / obj * 100) if obj > 0 else 0
    color = GREEN if pct >= 90 else (GOLD_DARK if pct >= 70 else RED)
    ax.text(x[i] + width/2, real + max(obj_s1)*0.02, f'{pct:.0f}%', ha='center', va='bottom', fontsize=8, color=color, fontweight='bold')

plt.savefig(f'{OUT_DIR}/chartA_ventes_vs_objectifs.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartA_ventes_vs_objectifs.png")

# ============ CHART B: Évolution mensuelle vs Objectifs (TOURTEAUX + CONCENTRES) ============
fig, axes = plt.subplots(1, 2, figsize=(12, 4.5), constrained_layout=True)

with open('/home/z/my-project/scripts/ventes_objectifs_data.json', 'r') as f:
    sales_data = json.load(f)

months_labels = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin']
for idx, cat in enumerate(['TOURTEAUX', 'CONCENTRES']):
    ax = axes[idx]
    obj_monthly = [data['global_objectives'][cat][str(m)] for m in range(1, 7)]
    real_monthly = [sales_data['category_sales_tons'].get(cat, {}).get(str(m), 0) for m in range(1, 7)]
    x = np.arange(6)
    ax.plot(x, obj_monthly, marker='o', linewidth=2.5, markersize=8, color=NAVY, label='Objectif')
    ax.plot(x, real_monthly, marker='s', linewidth=2.5, markersize=8, color=GOLD_DARK, label='Ventes réelles')
    ax.fill_between(x, obj_monthly, real_monthly, where=[r >= o for r, o in zip(real_monthly, obj_monthly)],
                    alpha=0.3, color=GREEN, label='Dépassement')
    ax.fill_between(x, obj_monthly, real_monthly, where=[r < o for r, o in zip(real_monthly, obj_monthly)],
                    alpha=0.3, color=RED, label='Sous-performance')
    ax.set_xticks(x)
    ax.set_xticklabels(months_labels, fontsize=9)
    ax.set_ylabel('Tonnes', fontsize=9)
    ax.set_title(f'{cat}', fontsize=11, fontweight='bold', color=NAVY)
    ax.legend(fontsize=8, loc='best')
    ax.grid(axis='y', alpha=0.3, linestyle='--')

plt.suptitle('Évolution mensuelle : ventes réelles vs objectifs (Jan-Juin 2026)', fontsize=12, fontweight='bold', color=NAVY, y=1.02)
plt.savefig(f'{OUT_DIR}/chartB_evolution_vs_objectifs.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartB_evolution_vs_objectifs.png")

# ============ CHART C: % atteinte objectif par catégorie (Q1, Q2, S1) ============
fig, ax = plt.subplots(figsize=(11, 5), constrained_layout=True)
cats_full = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT COMPLET', 'INGREDIENTS', 'COMPLEMENT ALIMENTAIRE', 'PREMIX']
pct_q1 = []
pct_q2 = []
pct_s1 = []
for c in cats_full:
    d = next((x for x in data['comparison'] if x['category'] == c), None)
    if d:
        pct_q1.append(min(d['pct_q1'], 200))  # cap at 200% for readability
        pct_q2.append(min(d['pct_q2'], 200))
        pct_s1.append(min(d['pct_s1'], 200))
    else:
        pct_q1.append(0); pct_q2.append(0); pct_s1.append(0)

x = np.arange(len(cats_full))
width = 0.25
b1 = ax.bar(x - width, pct_q1, width, label='% Q1', color=NAVY_LIGHT)
b2 = ax.bar(x, pct_q2, width, label='% Q2', color=GOLD)
b3 = ax.bar(x + width, pct_s1, width, label='% S1', color=NAVY)

# 100% line
ax.axhline(y=100, color='black', linewidth=1, linestyle='--', alpha=0.7, label='Objectif (100%)')

ax.set_xticks(x)
ax.set_xticklabels([c.replace(' ', '\n', 1) for c in cats_full], fontsize=8)
ax.set_ylabel('% atteinte objectif', fontsize=10)
ax.set_title('Taux d\'atteinte des objectifs par catégorie (Q1, Q2, S1 2026)', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='upper right', fontsize=9)
ax.grid(axis='y', alpha=0.3, linestyle='--')
ax.set_ylim(0, 220)

for bars in [b1, b2, b3]:
    for bar in bars:
        h = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., h + 3, f'{h:.0f}%', ha='center', va='bottom', fontsize=7)

plt.savefig(f'{OUT_DIR}/chartC_pct_atteinte.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartC_pct_atteinte.png")

# ============ CHART D: Forecast S2 (3 scénarios) par catégorie principale ============
# Compute forecast scenarios
forecast_data = {}
for cat in ['TOURTEAUX', 'CONCENTRES', 'ALIMENT COMPLET', 'INGREDIENTS', 'COMPLEMENT ALIMENTAIRE', 'PREMIX']:
    d = next((x for x in data['comparison'] if x['category'] == cat), None)
    if not d: continue
    # S2 objective
    s2_obj = sum(data['global_objectives'][cat][str(m)] for m in range(7, 13))
    # S1 real
    s1_real = d['s1_real']
    # S1 achievement rate
    s1_pct = d['pct_s1'] / 100
    # 3 scenarios:
    # Pessimist: S1 rate - 10% (degradation)
    # Realist: S1 rate (status quo)
    # Optimist: S1 rate + 15% (improvement from plan + concurrent rupture)
    # Apply to S2 objective
    pess = s2_obj * max(0, s1_pct - 0.10)
    real = s2_obj * s1_pct
    opt = s2_obj * min(1.20, s1_pct + 0.15)
    forecast_data[cat] = {'s2_obj': s2_obj, 'pess': pess, 'real': real, 'opt': opt}

fig, ax = plt.subplots(figsize=(11, 5), constrained_layout=True)
cats_f = list(forecast_data.keys())
s2_obj_vals = [forecast_data[c]['s2_obj'] for c in cats_f]
pess_vals = [forecast_data[c]['pess'] for c in cats_f]
real_vals = [forecast_data[c]['real'] for c in cats_f]
opt_vals = [forecast_data[c]['opt'] for c in cats_f]

x = np.arange(len(cats_f))
width = 0.20
b0 = ax.bar(x - 1.5*width, s2_obj_vals, width, label='Objectif S2', color=GRAY_LIGHT, edgecolor='white')
b1 = ax.bar(x - 0.5*width, pess_vals, width, label='🔴 Pessimiste', color=RED, edgecolor='white')
b2 = ax.bar(x + 0.5*width, real_vals, width, label='🟡 Réaliste', color=GOLD_DARK, edgecolor='white')
b3 = ax.bar(x + 1.5*width, opt_vals, width, label='🟢 Optimiste', color=GREEN, edgecolor='white')

ax.set_xticks(x)
ax.set_xticklabels([c.replace(' ', '\n', 1) for c in cats_f], fontsize=8)
ax.set_ylabel('Volume S2 (tonnes)', fontsize=10)
ax.set_title('Forecast S2 2026 (3 scénarios) vs Objectif, par catégorie', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='upper right', fontsize=9)
ax.grid(axis='y', alpha=0.3, linestyle='--')

plt.savefig(f'{OUT_DIR}/chartD_forecast_s2.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartD_forecast_s2.png")

# ============ CHART E: Total forecast S2 vs Objectif ============
fig, ax = plt.subplots(figsize=(8, 4.5), constrained_layout=True)
total_s2_obj = sum(forecast_data[c]['s2_obj'] for c in forecast_data)
total_pess = sum(forecast_data[c]['pess'] for c in forecast_data)
total_real = sum(forecast_data[c]['real'] for c in forecast_data)
total_opt = sum(forecast_data[c]['opt'] for c in forecast_data)
total_s1_real = sum(d['s1_real'] for d in data['comparison'])

labels = ['S1 réel\n(Jan-Juin)', 'S2 objectif', 'S2 Pessimiste', 'S2 Réaliste', 'S2 Optimiste']
values = [total_s1_real, total_s2_obj, total_pess, total_real, total_opt]
colors_bar = [NAVY_LIGHT, GRAY_LIGHT, RED, GOLD_DARK, GREEN]
bars = ax.bar(labels, values, color=colors_bar, edgecolor='white', linewidth=1.5)
ax.set_ylabel('Volume total (tonnes)', fontsize=10)
ax.set_title('Forecast total S2 2026 vs Objectif (en tonnes)', fontsize=12, fontweight='bold', color=NAVY)
ax.grid(axis='y', alpha=0.3, linestyle='--')
for bar, val in zip(bars, values):
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., h + max(values)*0.01, f'{val:,.0f}'.replace(',', ' '), ha='center', va='bottom', fontsize=9, fontweight='bold')
ax.set_ylim(0, max(values) * 1.15)
plt.savefig(f'{OUT_DIR}/chartE_forecast_total.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chartE_forecast_total.png")

# Save forecast data
with open('/home/z/my-project/scripts/forecast_s2.json', 'w', encoding='utf-8') as f:
    json.dump(forecast_data, f, ensure_ascii=False, indent=2, default=str)
print(f"\nForecast data saved to /home/z/my-project/scripts/forecast_s2.json")
print(f"\nTotal forecast S2:")
print(f"  S1 réel: {total_s1_real:,.0f} t")
print(f"  S2 objectif: {total_s2_obj:,.0f} t")
print(f"  S2 Pessimiste: {total_pess:,.0f} t ({total_pess/total_s2_obj*100:.1f}% obj)")
print(f"  S2 Réaliste: {total_real:,.0f} t ({total_real/total_s2_obj*100:.1f}% obj)")
print(f"  S2 Optimiste: {total_opt:,.0f} t ({total_opt/total_s2_obj*100:.1f}% obj)")
