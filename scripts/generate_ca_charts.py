"""Generate CA comparison chart with real CA objectives."""
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

with open('/home/z/my-project/scripts/ca_obj_real.json', 'r') as f:
    data = json.load(f)

# ============ CHART U: CA réel vs CA objectif réel par catégorie ============
fig, ax = plt.subplots(figsize=(11, 5), constrained_layout=True)
cats = ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'ALIMENT COMPLET', 'MATERIEL ELEVAGE', 'COMPLEMENT ALIM.']
cat_keys = ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'ALIMENT COMPLET', 'MATERIEL ELEVAGE', 'COMPLEMENT ALIMENTAIRE']

obj_vals = [data['ca_obj_S1'].get(k, 0) / 1e6 for k in cat_keys]
real_vals = [data['ca_real_S1'].get(k, 0) / 1e6 for k in cat_keys]

x = np.arange(len(cats))
width = 0.35
b1 = ax.bar(x - width/2, obj_vals, width, label='CA objectif S1', color=NAVY, edgecolor='white')
b2 = ax.bar(x + width/2, real_vals, width, label='CA réel S1', color=GOLD_DARK, edgecolor='white')

ax.set_xticks(x)
ax.set_xticklabels([c.replace(' ', '\n', 1) for c in cats], fontsize=8)
ax.set_ylabel('CA HT (millions FCFA)', fontsize=10)
ax.set_title('CA réel vs CA objectif S1 2026 par catégorie (objectifs CA réels)', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='upper right', fontsize=9)
ax.grid(axis='y', alpha=0.3, linestyle='--')

for i, k in enumerate(cat_keys):
    obj = data['ca_obj_S1'].get(k, 0)
    real = data['ca_real_S1'].get(k, 0)
    pct = (real / obj * 100) if obj > 0 else 0
    color = GREEN if pct >= 100 else (GOLD_DARK if pct >= 80 else RED)
    max_val = max(obj_vals[i], real_vals[i])
    ax.text(x[i], max_val + 100, f'{pct:.0f}%', ha='center', va='bottom', fontsize=8, color=color, fontweight='bold')

plt.savefig(f'{OUT_DIR}/chartU_ca_real_vs_obj.png', dpi=150, bbox_inches='tight')
plt.close()
print("  OK chartU_ca_real_vs_obj.png")
