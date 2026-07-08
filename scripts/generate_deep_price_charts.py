"""Generate deep price analysis charts."""
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

with open('/home/z/my-project/scripts/deep_price_analysis.json', 'r') as f:
    data = json.load(f)

# ============ CHART W: Price by CONCENTRES sub-category ============
fig, ax = plt.subplots(figsize=(10, 5), constrained_layout=True)
subcats = ['BELGO 10% Chair', 'BELGO 5% Chair', 'BELGO 10% Porc', 'BELGO 5% Ponte', 'BELGO 10% Ponte']
prices = [data['conc_subcat_price'][sc]['price']/1e3 for sc in subcats]
vols = [data['conc_subcat_price'][sc]['vol_t'] for sc in subcats]
labels = [sc.replace('BELGO ', '') for sc in subcats]
colors_w = [NAVY, GOLD_DARK, RED, GREEN, GRAY_LIGHT]

x = np.arange(len(subcats))
bars = ax.bar(x, prices, color=colors_w, edgecolor='white', linewidth=1.5)
ax.set_xticks(x)
ax.set_xticklabels(labels, fontsize=9, rotation=15)
ax.set_ylabel('Prix moyen (milliers FCFA/t)', fontsize=10)
ax.set_title('Prix moyen par sous-cat\u00e9gorie de concentr\u00e9s (S1 2026)', fontsize=11, fontweight='bold', color=NAVY)
ax.grid(axis='y', alpha=0.3, linestyle='--')

for i, (bar, price, vol) in enumerate(zip(bars, prices, vols)):
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., h + 5, f'{price:.0f} K\n({vol:.0f} t)', ha='center', va='bottom', fontsize=8)

ax.set_ylim(0, max(prices) * 1.25)
plt.savefig(f'{OUT_DIR}/chartW_conc_subcat_price.png', dpi=150, bbox_inches='tight')
plt.close()
print("  OK chartW_conc_subcat_price.png")

# ============ CHART X: Price by C104 format ============
fig, ax = plt.subplots(figsize=(8, 4.5), constrained_layout=True)
fmts = ['50Kg', '25Kg', '5Kg', '1Kg']
refs = ['C104', 'C1044', 'C1043', 'C1042']
prices_fmt = [data['c104_formats'][r]['price']/1e3 for r in refs]
vols_fmt = [data['c104_formats'][r]['vol_t'] for r in refs]

x = np.arange(len(fmts))
bars = ax.bar(x, prices_fmt, color=[NAVY, NAVY, NAVY, NAVY], edgecolor='white', linewidth=1.5)
# Gradient effect: darker for smaller formats
for i, bar in enumerate(bars):
    alpha = 1.0 - i * 0.2
    bar.set_alpha(alpha)
ax.set_xticks(x)
ax.set_xticklabels(fmts, fontsize=10)
ax.set_ylabel('Prix moyen (milliers FCFA/t)', fontsize=10)
ax.set_title('C104 (BELGO 10% Chair) — Prix par format (S1 2026)', fontsize=11, fontweight='bold', color=NAVY)
ax.grid(axis='y', alpha=0.3, linestyle='--')

for i, (bar, price, vol) in enumerate(zip(bars, prices_fmt, vols_fmt)):
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., h + 3, f'{price:.0f} K\n({vol:.0f} t)', ha='center', va='bottom', fontsize=8)

ax.set_ylim(0, max(prices_fmt) * 1.3)
plt.savefig(f'{OUT_DIR}/chartX_c104_format_price.png', dpi=150, bbox_inches='tight')
plt.close()
print("  OK chartX_c104_format_price.png")

# ============ CHART Y: Maïs vs other ingredients ============
fig, ax = plt.subplots(figsize=(8, 4.5), constrained_layout=True)
cats = ['Maïs', 'Autres ingrédients', 'Mixte INGREDIENTS']
prices_ing = [data['mais_vs_other']['mais']['price']/1e3,
              data['mais_vs_other']['other']['price']/1e3,
              data['mais_vs_other']['mixed_price']/1e3]
vols_ing = [data['mais_vs_other']['mais']['vol_t'],
            data['mais_vs_other']['other']['vol_t'],
            data['mais_vs_other']['mais']['vol_t'] + data['mais_vs_other']['other']['vol_t']]

x = np.arange(len(cats))
colors_y = [GREEN, NAVY, GRAY_LIGHT]
bars = ax.bar(x, prices_ing, color=colors_y, edgecolor='white', linewidth=1.5)
ax.set_xticks(x)
ax.set_xticklabels(cats, fontsize=9)
ax.set_ylabel('Prix moyen (milliers FCFA/t)', fontsize=10)
ax.set_title('INGREDIENTS — S\u00e9paration Maïs vs autres (S1 2026)', fontsize=11, fontweight='bold', color=NAVY)
ax.grid(axis='y', alpha=0.3, linestyle='--')

for i, (bar, price, vol) in enumerate(zip(bars, prices_ing, vols_ing)):
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., h + 10, f'{price:.0f} K\n({vol:.0f} t)', ha='center', va='bottom', fontsize=8)

ax.set_ylim(0, max(prices_ing) * 1.25)
plt.savefig(f'{OUT_DIR}/chartY_mais_vs_other.png', dpi=150, bbox_inches='tight')
plt.close()
print("  OK chartY_mais_vs_other.png")

# ============ CHART Z: Price by region ============
fig, axes = plt.subplots(1, 2, figsize=(12, 4.5), constrained_layout=True)

# All categories by region
regions = ['Ouest', 'Centre', 'Littoral']
prices_reg = []
vols_reg = []
ca_reg = []
for r in regions:
    d = data['region_data'][r]
    price = d['ca'] / d['vol_t'] if d['vol_t'] > 0 else 0
    prices_reg.append(price / 1e3)
    vols_reg.append(d['vol_t'])
    ca_reg.append(d['ca'] / 1e6)

axes[0].bar(regions, prices_reg, color=[NAVY, GOLD_DARK, GREEN], edgecolor='white', linewidth=1.5)
axes[0].set_ylabel('Prix moyen (milliers FCFA/t)', fontsize=9)
axes[0].set_title('Toutes cat\u00e9gories', fontsize=10, fontweight='bold', color=NAVY)
axes[0].grid(axis='y', alpha=0.3, linestyle='--')
for i, (r, p, v) in enumerate(zip(regions, prices_reg, vols_reg)):
    axes[0].text(i, p + 5, f'{p:.0f} K\n({v:.0f} t)', ha='center', va='bottom', fontsize=8)
axes[0].set_ylim(0, max(prices_reg) * 1.25)

# CONCENTRES by region
prices_conc_reg = []
vols_conc_reg = []
for r in regions:
    d = data['region_conc'][r]
    price = d['ca'] / d['vol_t'] if d['vol_t'] > 0 else 0
    prices_conc_reg.append(price / 1e3)
    vols_conc_reg.append(d['vol_t'])

axes[1].bar(regions, prices_conc_reg, color=[NAVY, GOLD_DARK, GREEN], edgecolor='white', linewidth=1.5)
axes[1].set_ylabel('Prix moyen CONCENTRÉS (milliers FCFA/t)', fontsize=9)
axes[1].set_title('CONCENTRÉS uniquement', fontsize=10, fontweight='bold', color=NAVY)
axes[1].grid(axis='y', alpha=0.3, linestyle='--')
for i, (r, p, v) in enumerate(zip(regions, prices_conc_reg, vols_conc_reg)):
    axes[1].text(i, p + 5, f'{p:.0f} K\n({v:.0f} t)', ha='center', va='bottom', fontsize=8)
axes[1].set_ylim(0, max(prices_conc_reg) * 1.25)

plt.suptitle('Prix moyen par r\u00e9gion (S1 2026)', fontsize=11, fontweight='bold', color=NAVY, y=1.02)
plt.savefig(f'{OUT_DIR}/chartZ_price_by_region.png', dpi=150, bbox_inches='tight')
plt.close()
print("  OK chartZ_price_by_region.png")

print(f"\nAll 4 deep price charts saved.")
