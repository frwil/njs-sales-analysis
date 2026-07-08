"""Generate price analysis chart."""
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
GRAY = '#595959'
GRAY_LIGHT = '#A6A6A6'

with open('/home/z/my-project/scripts/price_analysis.json', 'r') as f:
    data = json.load(f)

# ============ CHART V: Prix moyen comparison ============
fig, ax = plt.subplots(figsize=(10, 5), constrained_layout=True)

cats = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT COMPLET']
cat_labels = ['Tourteaux\n(soja)', 'Concentr\u00e9s\n(BELGO 10%+5%)', 'Aliment complet\n(Booster+Rabbit+Fish)']

prices_obj = [data[c]['obj']['price'] / 1e3 for c in cats]  # in K FCFA/t
prices_2026 = [data[c]['2026']['price'] / 1e3 for c in cats]
prices_2025 = [data[c]['2025']['price'] / 1e3 for c in cats]

x = np.arange(len(cats))
width = 0.25
b1 = ax.bar(x - width, prices_2025, width, label='2025', color=GRAY_LIGHT, edgecolor='white')
b2 = ax.bar(x, prices_obj, width, label='Objectif 2026', color=NAVY, edgecolor='white')
b3 = ax.bar(x + width, prices_2026, width, label='S1 2026', color=GOLD_DARK, edgecolor='white')

ax.set_xticks(x)
ax.set_xticklabels(cat_labels, fontsize=9)
ax.set_ylabel('Prix moyen (milliers FCFA/t)', fontsize=10)
ax.set_title('Prix moyen par cat\u00e9gorie : 2025 vs Objectif 2026 vs S1 2026', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(loc='upper left', fontsize=9)
ax.grid(axis='y', alpha=0.3, linestyle='--')

# Add YoY and vs obj labels
for i, cat in enumerate(cats):
    yoy = data[cat]['yoy_price']
    vs_obj = data[cat]['vs_obj_price']
    max_val = max(prices_2025[i], prices_obj[i], prices_2026[i])
    ax.text(x[i], max_val + 15, f'YoY: {yoy:+.1f}%\nvs obj: {vs_obj:+.1f}%', ha='center', va='bottom', fontsize=8, color=NAVY)

# Add international soja price line
soja_world = 322 * 600 / 1e3  # 322 USD/t * 600 FCFA/USD / 1000 = K FCFA/t
ax.axhline(y=soja_world, color=RED, linewidth=1.5, linestyle='--', alpha=0.7, label=f'Prix mondial soja ({soja_world:.0f} K FCFA/t)')

ax.legend(loc='upper left', fontsize=8)
ax.set_ylim(0, max(prices_2025 + prices_obj + prices_2026) * 1.25)

plt.savefig(f'{OUT_DIR}/chartV_price_analysis.png', dpi=150, bbox_inches='tight')
plt.close()
print("  OK chartV_price_analysis.png")
