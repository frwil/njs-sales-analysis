"""
Generate PNG charts for the BELGOCAM PDF report.
Charts are saved to /home/z/my-project/scripts/pdf_charts/ and embedded in the PDF via ReportLab Image().
"""
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np

# Font setup (Liberation Sans = clean Latin)
plt.rcParams['font.sans-serif'] = ['Liberation Sans', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

OUT_DIR = '/home/z/my-project/scripts/pdf_charts'
os.makedirs(OUT_DIR, exist_ok=True)

# BELGOCAM colors
NAVY = '#1F4E78'
NAVY_LIGHT = '#2E75B6'
GOLD = '#FFC000'
GOLD_DARK = '#BF8F00'
RED = '#C00000'
GREEN = '#548235'
GRAY = '#595959'
GRAY_LIGHT = '#A6A6A6'

# ============ CHART 1: Évolution mensuelle du nombre de clients actifs ============
fig, ax = plt.subplots(figsize=(8, 4.5), constrained_layout=True)
months = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin']
all_products = [941, 836, 859, 836, 853, 958]
cible = [815, 690, 718, 705, 738, 856]
concentres = [634, 534, 560, 547, 573, 672]
x = np.arange(len(months))
ax.plot(x, all_products, marker='o', linewidth=2.5, markersize=8, color=NAVY, label='Tous produits')
ax.plot(x, cible, marker='s', linewidth=2.5, markersize=8, color=GOLD_DARK, label='Ciblé (16 produits)')
ax.plot(x, concentres, marker='^', linewidth=2.5, markersize=8, color='#ED7D31', label='Concentrés (12 produits)')
ax.set_xticks(x)
ax.set_xticklabels(months, fontsize=10)
ax.set_ylabel('Nombre de clients actifs', fontsize=10)
ax.set_title('Évolution mensuelle du nombre de clients actifs', fontsize=12, fontweight='bold', color=NAVY)
ax.legend(loc='lower right', framealpha=0.95, fontsize=9)
ax.grid(axis='y', alpha=0.3, linestyle='--')
ax.set_ylim(0, max(all_products) * 1.15)
for i, v in enumerate(all_products):
    ax.text(i, v + 25, str(v), ha='center', va='bottom', fontsize=8, color=NAVY)
plt.savefig(f'{OUT_DIR}/chart1_evolution.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chart1_evolution.png")

# ============ CHART 2: Segments de transition Q1→Q2 (ciblé) ============
fig, ax = plt.subplots(figsize=(8, 4.5), constrained_layout=True)
segments = ['Persistants\n(Zero→Zero)', 'Réactivés\n(Zero→Active)', 'Retenus\n(Active→Active)', 'Churned\n(Active→Zero)']
counts = [118, 234, 850, 192]
colors_seg = [GRAY_LIGHT, GREEN, NAVY_LIGHT, RED]
bars = ax.bar(segments, counts, color=colors_seg, edgecolor='white', linewidth=1.5)
ax.set_ylabel('Nombre de clients', fontsize=10)
ax.set_title('Segments de transition Q1 → Q2 (16 produits ciblés)', fontsize=12, fontweight='bold', color=NAVY)
ax.grid(axis='y', alpha=0.3, linestyle='--')
for bar, val in zip(bars, counts):
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., h + 15, f'{val}', ha='center', va='bottom', fontsize=11, fontweight='bold')
ax.set_ylim(0, max(counts) * 1.15)
plt.savefig(f'{OUT_DIR}/chart2_segments.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chart2_segments.png")

# ============ CHART 3: Top 10 pertes Q1 (CA en millions) ============
fig, ax = plt.subplots(figsize=(9, 5), constrained_layout=True)
clients = ['SIGHELO SARL', 'COMPAGNIE\nFERMIERE', 'TEIKING\nJEAN MARIE', 'FEUDJIO\nBACK ARMEL',
           'KENMEGNE\nALAIN', 'STE SATI SARL', 'FOTSO\nVINCENT', 'DJOUSSI\nAURORE',
           'STE IPACAM\n& FILS', 'BAHO']
pertes = [23.7, 21.0, 19.6, 17.5, 16.6, 13.6, 13.5, 13.3, 14.8, 10.9]
y_pos = np.arange(len(clients))
bars = ax.barh(y_pos, pertes, color=RED, edgecolor='white', linewidth=1)
ax.set_yticks(y_pos)
ax.set_yticklabels(clients, fontsize=8)
ax.invert_yaxis()
ax.set_xlabel('Perte Q1 estimée (millions FCFA)', fontsize=10)
ax.set_title('Top 10 pertes Q1 — Clients zéro achat global (16 produits)', fontsize=11, fontweight='bold', color=NAVY)
ax.grid(axis='x', alpha=0.3, linestyle='--')
for bar, val in zip(bars, pertes):
    ax.text(val + 0.3, bar.get_y() + bar.get_height()/2., f'{val:.1f} M', va='center', fontsize=8)
ax.set_xlim(0, max(pertes) * 1.15)
plt.savefig(f'{OUT_DIR}/chart3_top10_pertes.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chart3_top10_pertes.png")

# ============ CHART 4: Zero achat par agence (stacked bar) ============
fig, ax = plt.subplots(figsize=(9, 5), constrained_layout=True)
agences = ['FAMLA', 'MESSASSI', 'NDOBO', 'DJELENG', 'NGAOUNDERE', 'NKONGSAMBA', 'Autres']
zero_2080 = [7, 0, 4, 2, 0, 0, 1]
zero_autres = [49, 56, 41, 25, 31, 17, 121]
actifs = [237, 111, 109, 80, 60, 80, 525]
x = np.arange(len(agences))
width = 0.55
b1 = ax.bar(x, zero_2080, width, label='Zero achat 20/80', color=GOLD, edgecolor='white')
b2 = ax.bar(x, zero_autres, width, bottom=zero_2080, label='Zero achat autres', color=GRAY_LIGHT, edgecolor='white')
b3 = ax.bar(x, actifs, width, bottom=[a+b for a,b in zip(zero_2080, zero_autres)], label='Actifs Q1', color=NAVY_LIGHT, edgecolor='white')
ax.set_xticks(x)
ax.set_xticklabels(agences, fontsize=9)
ax.set_ylabel('Nombre de clients', fontsize=10)
ax.set_title('Répartition zero achat Q1 (ciblé) par agence', fontsize=12, fontweight='bold', color=NAVY)
ax.legend(loc='upper right', framealpha=0.95, fontsize=9)
ax.grid(axis='y', alpha=0.3, linestyle='--')
plt.savefig(f'{OUT_DIR}/chart4_zero_agence.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chart4_zero_agence.png")

# ============ CHART 5: Bilan net Q1→Q2 (ciblé vs concentrés) ============
fig, ax = plt.subplots(figsize=(8, 4.5), constrained_layout=True)
categories = ['Ciblé (16 produits)', 'Concentrés (12 produits)']
gains = [549.7, 145.5]
pertes_b = [273.3, 93.1]
nets = [g-p for g, p in zip(gains, pertes_b)]
x = np.arange(len(categories))
width = 0.25
b1 = ax.bar(x - width, gains, width, label='Gain (réactivés)', color=GREEN)
b2 = ax.bar(x, pertes_b, width, label='Perte (churned)', color=RED)
b3 = ax.bar(x + width, nets, width, label='Bilan net', color=NAVY)
ax.set_xticks(x)
ax.set_xticklabels(categories, fontsize=10)
ax.set_ylabel('Montant (millions FCFA)', fontsize=10)
ax.set_title('Bilan net Q1 → Q2 : Ciblé vs Concentrés', fontsize=12, fontweight='bold', color=NAVY)
ax.legend(loc='upper right', fontsize=9)
ax.grid(axis='y', alpha=0.3, linestyle='--')
ax.axhline(y=0, color='black', linewidth=0.8)
for bars in [b1, b2, b3]:
    for bar in bars:
        h = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., h + (10 if h >= 0 else -25), f'{h:.0f}', ha='center', va='bottom' if h>=0 else 'top', fontsize=9)
plt.savefig(f'{OUT_DIR}/chart5_bilan_net.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chart5_bilan_net.png")

# ============ CHART 6: Synergie Chick × Concentrés (pie) ============
fig, axes = plt.subplots(1, 2, figsize=(10, 4.5), constrained_layout=True)
# Pie 1: Chick
chick_labels = ['Cross-sell\n(Chair/Ponte)', 'Chick only\n(opportunité)']
chick_values = [270, 73]
axes[0].pie(chick_values, labels=chick_labels, colors=[GREEN, RED], autopct='%1.1f%%',
            startangle=90, textprops={'fontsize': 10})
axes[0].set_title('Chick Booster × Conc. Chair/Ponte\n(343 clients Chick)', fontsize=11, fontweight='bold', color=NAVY)
# Pie 2: Piglet
piglet_labels = ['Cross-sell\n(Porc)', 'Piglet only\n(opportunité)']
piglet_values = [162, 68]
axes[1].pie(piglet_values, labels=piglet_labels, colors=[GREEN, RED], autopct='%1.1f%%',
            startangle=90, textprops={'fontsize': 10})
axes[1].set_title('Piglet Booster × Conc. Porc\n(230 clients Piglet)', fontsize=11, fontweight='bold', color=NAVY)
plt.savefig(f'{OUT_DIR}/chart6_synergie_booster.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chart6_synergie_booster.png")

# ============ CHART 7: Projection CA additionnel (3 scénarios) ============
fig, ax = plt.subplots(figsize=(9, 4.5), constrained_layout=True)
scenarios = ['Pessimiste', 'Realiste', 'Optimiste']
reactivation = [34, 68, 103]
reconquete = [27, 81, 135]
acquisition = [3, 16, 63]
recup_pertes = [58, 174, 348]
concurrents = [100, 300, 600]
x = np.arange(len(scenarios))
width = 0.15
ax.bar(x - 2*width, reactivation, width, label='Réactivation 20/80', color=NAVY)
ax.bar(x - width, reconquete, width, label='Reconquête churned', color=NAVY_LIGHT)
ax.bar(x, acquisition, width, label='Acquisition persistants', color=GOLD_DARK)
ax.bar(x + width, recup_pertes, width, label='Récupération pertes Q1', color='#ED7D31')
ax.bar(x + 2*width, concurrents, width, label='Conquête concurrents', color=GREEN)
ax.set_xticks(x)
ax.set_xticklabels(scenarios, fontsize=10)
ax.set_ylabel('CA additionnel (millions FCFA)', fontsize=10)
ax.set_title('Projection CA additionnel à 6 mois par scénario', fontsize=12, fontweight='bold', color=NAVY)
ax.legend(loc='upper left', fontsize=8, ncol=2)
ax.grid(axis='y', alpha=0.3, linestyle='--')
plt.savefig(f'{OUT_DIR}/chart7_projection.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chart7_projection.png")

# ============ CHART 8: Distribution fréquence d'achat ============
fig, ax = plt.subplots(figsize=(8, 4.5), constrained_layout=True)
freq_labels = ['6 mois\n(fidèles)', '5 mois', '4 mois', '3 mois\n(semi-fid.)', '2 mois', '1 mois\n(one-shot)']
freq_values = [452, 141, 131, 145, 196, 294]
colors_freq = [NAVY, NAVY_LIGHT, GOLD_DARK, '#ED7D31', GRAY_LIGHT, RED]
bars = ax.bar(freq_labels, freq_values, color=colors_freq, edgecolor='white', linewidth=1.5)
ax.set_ylabel('Nombre de clients', fontsize=10)
ax.set_title('Distribution de la fréquence d\'achat (sur 6 mois)', fontsize=12, fontweight='bold', color=NAVY)
ax.grid(axis='y', alpha=0.3, linestyle='--')
for bar, val in zip(bars, freq_values):
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., h + 10, f'{val}', ha='center', va='bottom', fontsize=10, fontweight='bold')
ax.set_ylim(0, max(freq_values) * 1.15)
plt.savefig(f'{OUT_DIR}/chart8_frequence.png', dpi=150, bbox_inches='tight')
plt.close()
print("  ✓ chart8_frequence.png")

print(f"\nAll charts saved to: {OUT_DIR}")
