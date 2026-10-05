"""
Génère les graphiques matplotlib (PNG) pour le PDF report
Style corporate sobre: bleu marine + gris + accents dorés
"""
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.ticker as mticker
import os
import warnings
warnings.filterwarnings('ignore')

# Register fonts (skip if variable font fails - DejaVu is enough for our needs)
try:
    fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
    fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf')
except Exception:
    pass

plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 10

# Corporate sober palette
COLOR_PRIMARY = "#1F3864"      # Bleu marine
COLOR_SECONDARY = "#2E5C8A"
COLOR_ACCENT = "#C9A961"        # Doré sobre
COLOR_CHICK = "#7C3AED"         # Violet
COLOR_PIGLET = "#D97706"        # Orange
COLOR_GREY = "#6B7280"
COLOR_LIGHT = "#E5E7EB"
COLOR_BG = "#F8FAFC"

# Load data
WORK = "/home/z/my-project/work"
CHARTS = "/home/z/my-project/work/charts"
os.makedirs(CHARTS, exist_ok=True)

df = pd.read_csv(os.path.join(WORK, 'booster_consolidated.csv'))
df['Date de commande'] = pd.to_datetime(df['Date de commande'])

months_order = ['2025-10', '2025-11', '2025-12', '2026-01', '2026-02', '2026-03',
                '2026-04', '2026-05', '2026-06', '2026-07', '2026-08', '2026-09', '2026-10']
month_labels_fr = ['Oct 25', 'Nov 25', 'Déc 25', 'Jan 26', 'Fév 26', 'Mar 26',
                   'Avr 26', 'Mai 26', 'Juin 26', 'Juil 26', 'Août 26', 'Sep 26', 'Oct 26*']

# ============================================================
# Chart 1: Évolution mensuelle globale - Stacked bar Chick + Piglet
# ============================================================
print("Chart 1: Évolution mensuelle globale...")
pivot = df.pivot_table(values='qte_tonnes', index='mois', columns='famille', aggfunc='sum', fill_value=0)
pivot = pivot.reindex(months_order, fill_value=0)

fig, ax = plt.subplots(figsize=(11, 5.5), constrained_layout=True)
x = np.arange(len(months_order))
width = 0.7

bars1 = ax.bar(x, pivot['Chick Booster'], width, label='Chick Booster', color=COLOR_CHICK, edgecolor='white', linewidth=0.5)
bars2 = ax.bar(x, pivot['Piglet Booster'], width, bottom=pivot['Chick Booster'], label='Piglet Booster', color=COLOR_PIGLET, edgecolor='white', linewidth=0.5)

# Total line on top
totals = pivot['Chick Booster'] + pivot['Piglet Booster']
ax.plot(x, totals, color=COLOR_PRIMARY, marker='o', markersize=6, linewidth=2.2, label='Total Booster', zorder=5)

# Value labels
for i, t in enumerate(totals):
    ax.annotate(f'{t:.1f}', xy=(i, t), xytext=(0, 8), textcoords='offset points',
                ha='center', fontsize=9, color=COLOR_PRIMARY, fontweight='bold')

ax.set_xticks(x)
ax.set_xticklabels(month_labels_fr, rotation=0, fontsize=9)
ax.set_ylabel('Volume (tonnes)', fontsize=10, color=COLOR_PRIMARY)
ax.set_title('Évolution mensuelle des ventes Chick & Piglet Booster',
             fontsize=13, color=COLOR_PRIMARY, fontweight='bold', pad=12)
ax.yaxis.grid(True, linestyle='--', alpha=0.4, color=COLOR_GREY)
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color(COLOR_GREY)
ax.spines['bottom'].set_color(COLOR_GREY)
ax.tick_params(colors=COLOR_GREY)
ax.legend(loc='upper left', frameon=False, fontsize=10)
ax.set_ylim(0, max(totals) * 1.18)

plt.savefig(os.path.join(CHARTS, '01_evolution_globale.png'), dpi=180, facecolor='white')
plt.close()
print("  -> 01_evolution_globale.png")

# ============================================================
# Chart 2: Évolution par région - Multi-line
# ============================================================
print("Chart 2: Évolution par région...")
pivot_reg = df.pivot_table(values='qte_tonnes', index='mois', columns='region', aggfunc='sum', fill_value=0)
pivot_reg = pivot_reg.reindex(months_order, fill_value=0)
# Sort regions by total descending (exclude Non spécifié)
region_totals = pivot_reg.sum().sort_values(ascending=False)
region_order = [r for r in region_totals.index if r != 'Non spécifié']
# Add Non spécifié at end if exists
if 'Non spécifié' in region_totals.index:
    region_order.append('Non spécifié')

fig, ax = plt.subplots(figsize=(11, 5.5), constrained_layout=True)
colors = [COLOR_PRIMARY, COLOR_PIGLET, COLOR_CHICK, COLOR_GREY]
markers = ['o', 's', '^', 'D']

for i, region in enumerate(region_order):
    vals = pivot_reg[region].values
    ax.plot(x, vals, color=colors[i % len(colors)], marker=markers[i % len(markers)],
            markersize=6, linewidth=2, label=region)

ax.set_xticks(x)
ax.set_xticklabels(month_labels_fr, rotation=0, fontsize=9)
ax.set_ylabel('Volume (tonnes)', fontsize=10, color=COLOR_PRIMARY)
ax.set_title('Évolution mensuelle par région',
             fontsize=13, color=COLOR_PRIMARY, fontweight='bold', pad=12)
ax.yaxis.grid(True, linestyle='--', alpha=0.4, color=COLOR_GREY)
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color(COLOR_GREY)
ax.spines['bottom'].set_color(COLOR_GREY)
ax.tick_params(colors=COLOR_GREY)
ax.legend(loc='upper left', frameon=False, fontsize=10)

plt.savefig(os.path.join(CHARTS, '02_evolution_region.png'), dpi=180, facecolor='white')
plt.close()
print("  -> 02_evolution_region.png")

# ============================================================
# Chart 3: Top 10 agences - Horizontal bar
# ============================================================
print("Chart 3: Top 10 agences...")
agence_totals = df.groupby('agence')['qte_tonnes'].sum().sort_values(ascending=False).head(10)

fig, ax = plt.subplots(figsize=(11, 5.5), constrained_layout=True)
y = np.arange(len(agence_totals))
bars = ax.barh(y, agence_totals.values, color=COLOR_PRIMARY, edgecolor='white', linewidth=0.5)

# Color top 3 with accent
for i in range(min(3, len(bars))):
    bars[i].set_color(COLOR_ACCENT)

# Value labels
for i, v in enumerate(agence_totals.values):
    ax.text(v + 1, i, f'{v:.1f} t', va='center', fontsize=9, color=COLOR_PRIMARY, fontweight='bold')

ax.set_yticks(y)
ax.set_yticklabels(agence_totals.index, fontsize=10)
ax.invert_yaxis()
ax.set_xlabel('Volume (tonnes)', fontsize=10, color=COLOR_PRIMARY)
ax.set_title('Top 10 agences par volume cumulé (Oct 25 - Oct 26)',
             fontsize=13, color=COLOR_PRIMARY, fontweight='bold', pad=12)
ax.xaxis.grid(True, linestyle='--', alpha=0.4, color=COLOR_GREY)
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color(COLOR_GREY)
ax.spines['bottom'].set_color(COLOR_GREY)
ax.tick_params(colors=COLOR_GREY)
ax.set_xlim(0, max(agence_totals.values) * 1.15)

plt.savefig(os.path.join(CHARTS, '03_top_agences.png'), dpi=180, facecolor='white')
plt.close()
print("  -> 03_top_agences.png")

# ============================================================
# Chart 4: Évolution par produit (4 lines)
# ============================================================
print("Chart 4: Évolution par produit...")
products = ['CHICK BOOSTER 25 Kg', 'CHICK BOOSTER 5Kg', 'PIGLET BOOSTER 25Kg', 'PIGLET BOOSTER 5Kg']
product_labels = ['Chick Booster 25Kg', 'Chick Booster 5Kg', 'Piglet Booster 25Kg', 'Piglet Booster 5Kg']
product_colors = [COLOR_CHICK, '#A78BFA', COLOR_PIGLET, '#FCD34D']
product_markers = ['o', 'o', 's', 's']

pivot_pr = df.pivot_table(values='qte_tonnes', index='mois', columns='Description du produit', aggfunc='sum', fill_value=0)
pivot_pr = pivot_pr.reindex(months_order, fill_value=0)

fig, ax = plt.subplots(figsize=(11, 5.5), constrained_layout=True)
for i, (prod, label) in enumerate(zip(products, product_labels)):
    vals = pivot_pr[prod].values if prod in pivot_pr.columns else np.zeros(len(months_order))
    ax.plot(x, vals, color=product_colors[i], marker=product_markers[i],
            markersize=6, linewidth=2, label=label)

ax.set_xticks(x)
ax.set_xticklabels(month_labels_fr, rotation=0, fontsize=9)
ax.set_ylabel('Volume (tonnes)', fontsize=10, color=COLOR_PRIMARY)
ax.set_title('Évolution mensuelle par produit',
             fontsize=13, color=COLOR_PRIMARY, fontweight='bold', pad=12)
ax.yaxis.grid(True, linestyle='--', alpha=0.4, color=COLOR_GREY)
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color(COLOR_GREY)
ax.spines['bottom'].set_color(COLOR_GREY)
ax.tick_params(colors=COLOR_GREY)
ax.legend(loc='upper left', frameon=False, fontsize=9, ncol=2)

plt.savefig(os.path.join(CHARTS, '04_evolution_produit.png'), dpi=180, facecolor='white')
plt.close()
print("  -> 04_evolution_produit.png")

# ============================================================
# Chart 5: Pie chart - Ventilation par famille et format
# ============================================================
print("Chart 5: Ventilation famille/format...")
fam_fmt = df.groupby(['famille', 'format'])['qte_tonnes'].sum().sort_values(ascending=False)

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 5), constrained_layout=True)

# Pie 1: par famille
fam_totals = df.groupby('famille')['qte_tonnes'].sum()
colors_fam = [COLOR_CHICK, COLOR_PIGLET]
wedges, texts, autotexts = ax1.pie(fam_totals.values, labels=fam_totals.index,
                                    colors=colors_fam, autopct='%1.1f%%',
                                    startangle=90, textprops={'fontsize': 11, 'color': COLOR_PRIMARY},
                                    wedgeprops={'edgecolor': 'white', 'linewidth': 2})
for at in autotexts:
    at.set_color('white')
    at.set_fontweight('bold')
ax1.set_title('Répartition par famille', fontsize=12, color=COLOR_PRIMARY, fontweight='bold', pad=10)

# Pie 2: par format
fmt_totals = df.groupby('format')['qte_tonnes'].sum()
colors_fmt = [COLOR_PRIMARY, COLOR_ACCENT]
wedges2, texts2, autotexts2 = ax2.pie(fmt_totals.values, labels=fmt_totals.index,
                                      colors=colors_fmt, autopct='%1.1f%%',
                                      startangle=90, textprops={'fontsize': 11, 'color': COLOR_PRIMARY},
                                      wedgeprops={'edgecolor': 'white', 'linewidth': 2})
for at in autotexts2:
    at.set_color('white')
    at.set_fontweight('bold')
ax2.set_title('Répartition par format', fontsize=12, color=COLOR_PRIMARY, fontweight='bold', pad=10)

plt.savefig(os.path.join(CHARTS, '05_ventilation.png'), dpi=180, facecolor='white')
plt.close()
print("  -> 05_ventilation.png")

# ============================================================
# Chart 6: Heatmap Agence x Mois (top 15 agences)
# ============================================================
print("Chart 6: Heatmap Agence x Mois...")
top15_agences = df.groupby('agence')['qte_tonnes'].sum().sort_values(ascending=False).head(15).index
df_top = df[df['agence'].isin(top15_agences)]
pivot_hm = df_top.pivot_table(values='qte_tonnes', index='agence', columns='mois', aggfunc='sum', fill_value=0)
pivot_hm = pivot_hm.reindex(columns=months_order, fill_value=0)
# Reorder agences by total descending
pivot_hm = pivot_hm.loc[top15_agences]

fig, ax = plt.subplots(figsize=(11, 6), constrained_layout=True)
im = ax.imshow(pivot_hm.values, aspect='auto', cmap='Blues')

# Annotations
for i in range(len(pivot_hm.index)):
    for j in range(len(pivot_hm.columns)):
        val = pivot_hm.values[i, j]
        if val > 0:
            color = 'white' if val > pivot_hm.values.max() * 0.5 else COLOR_PRIMARY
            ax.text(j, i, f'{val:.1f}', ha='center', va='center',
                    fontsize=8, color=color, fontweight='bold')

ax.set_xticks(range(len(months_order)))
ax.set_xticklabels(month_labels_fr, fontsize=9, rotation=0)
ax.set_yticks(range(len(pivot_hm.index)))
ax.set_yticklabels(pivot_hm.index, fontsize=10)
ax.set_title('Heatmap des ventes mensuelles - Top 15 agences (tonnes)',
             fontsize=13, color=COLOR_PRIMARY, fontweight='bold', pad=12)

# Colorbar
cbar = plt.colorbar(im, ax=ax, shrink=0.8)
cbar.set_label('Volume (tonnes)', fontsize=10, color=COLOR_PRIMARY)
cbar.ax.tick_params(colors=COLOR_GREY)

ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color(COLOR_GREY)
ax.spines['bottom'].set_color(COLOR_GREY)
ax.tick_params(colors=COLOR_GREY)

plt.savefig(os.path.join(CHARTS, '06_heatmap_agence.png'), dpi=180, facecolor='white')
plt.close()
print("  -> 06_heatmap_agence.png")

# ============================================================
# Chart 7: Bar chart régions (cumul annuel)
# ============================================================
print("Chart 7: Cumul par région...")
region_tot = df.groupby('region')['qte_tonnes'].sum().sort_values(ascending=False)
region_tot = region_tot[region_tot.index != 'Non spécifié']  # exclude

fig, ax = plt.subplots(figsize=(9, 4.5), constrained_layout=True)
colors_reg = [COLOR_PRIMARY, COLOR_SECONDARY, COLOR_ACCENT]
bars = ax.bar(region_tot.index, region_tot.values, color=colors_reg[:len(region_tot)],
              edgecolor='white', linewidth=1.5, width=0.6)

for bar, v in zip(bars, region_tot.values):
    ax.text(bar.get_x() + bar.get_width()/2, v + 5,
            f'{v:.1f} t\n({v/region_tot.sum()*100:.1f}%)',
            ha='center', va='bottom', fontsize=10, color=COLOR_PRIMARY, fontweight='bold')

ax.set_ylabel('Volume (tonnes)', fontsize=10, color=COLOR_PRIMARY)
ax.set_title('Volume cumulé par région (Oct 25 - Oct 26)',
             fontsize=13, color=COLOR_PRIMARY, fontweight='bold', pad=12)
ax.yaxis.grid(True, linestyle='--', alpha=0.4, color=COLOR_GREY)
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color(COLOR_GREY)
ax.spines['bottom'].set_color(COLOR_GREY)
ax.tick_params(colors=COLOR_GREY)
ax.set_ylim(0, max(region_tot.values) * 1.2)

plt.savefig(os.path.join(CHARTS, '07_cumul_region.png'), dpi=180, facecolor='white')
plt.close()
print("  -> 07_cumul_region.png")

print(f"\n✅ Tous les graphiques générés dans {CHARTS}")
print(f"   Fichiers: {sorted(os.listdir(CHARTS))}")
