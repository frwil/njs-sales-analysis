"""
Génère le graphique chartAA_region_ventes_vs_obj.png avec les données régionales corrigées.
Affiche: volume total vs objectif + volume CONCENTRÉS vs objectif, par région.
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
import json

# Fonts — use Liberation Sans (available, works with matplotlib)
plt.rcParams['font.sans-serif'] = ['Liberation Sans', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# Charger les données corrigées
with open('/home/z/my-project/scripts/region_obj_analysis_corrected.json') as f:
    data = json.load(f)

regions = ["Ouest", "Centre", "Littoral"]
vol_obj = [data[r]["vol_obj"] for r in regions]
vol_real = [data[r]["vol_real"] for r in regions]
conc_obj = [data[r]["conc_obj"] for r in regions]
conc_vol = [data[r]["conc_vol"] for r in regions]

# Couleurs BELGOCAM
NAVY = "#1F3A5F"
GOLD = "#C9A961"
GRAY = "#8C8C8C"
LIGHT_NAVY = "#3A5F8F"
LIGHT_GOLD = "#E0C891"

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5.5), constrained_layout=True)

x = np.arange(len(regions))
width = 0.35

# --- Graphique 1: Volume total ---
bars1 = ax1.bar(x - width/2, vol_obj, width, label='Objectif S1', color=GRAY, alpha=0.7, edgecolor='white')
bars2 = ax1.bar(x + width/2, vol_real, width, label='Réalisé S1', color=NAVY, edgecolor='white')

ax1.set_xlabel('Région', fontsize=11, fontweight='bold')
ax1.set_ylabel('Volume (tonnes)', fontsize=11, fontweight='bold')
ax1.set_title('Volume total par région — S1 2026', fontsize=13, fontweight='bold', color=NAVY)
ax1.set_xticks(x)
ax1.set_xticklabels(regions, fontsize=11)
ax1.legend(fontsize=9, loc='upper right')
ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.yaxis.grid(True, alpha=0.3, linestyle='--')

# Annotations % atteinte
for i, (vo, vr) in enumerate(zip(vol_obj, vol_real)):
    pct = vr/vo*100 if vo > 0 else 0
    color = "#2E7D32" if pct >= 100 else "#C62828"
    ax1.annotate(f'{pct:.1f}%', xy=(i + width/2, vr), xytext=(0, 5),
                textcoords='offset points', ha='center', fontsize=10, fontweight='bold', color=color)
    # Valeurs
    ax1.annotate(f'{vo:,.0f}'.replace(',', ' '), xy=(i - width/2, vo), xytext=(0, 3),
                textcoords='offset points', ha='center', fontsize=8, color=GRAY)
    ax1.annotate(f'{vr:,.0f}'.replace(',', ' '), xy=(i + width/2, vr), xytext=(0, -12),
                textcoords='offset points', ha='center', fontsize=8, color=NAVY, fontweight='bold')

# --- Graphique 2: CONCENTRÉS ---
bars3 = ax2.bar(x - width/2, conc_obj, width, label='Objectif S1', color=GRAY, alpha=0.7, edgecolor='white')
bars4 = ax2.bar(x + width/2, conc_vol, width, label='Réalisé S1', color=GOLD, edgecolor='white')

ax2.set_xlabel('Région', fontsize=11, fontweight='bold')
ax2.set_ylabel('Volume CONCENTRÉS (tonnes)', fontsize=11, fontweight='bold')
ax2.set_title('CONCENTRÉS par région — S1 2026', fontsize=13, fontweight='bold', color=NAVY)
ax2.set_xticks(x)
ax2.set_xticklabels(regions, fontsize=11)
ax2.legend(fontsize=9, loc='upper right')
ax2.spines['top'].set_visible(False)
ax2.spines['right'].set_visible(False)
ax2.yaxis.grid(True, alpha=0.3, linestyle='--')

# Annotations % atteinte
for i, (co, cv) in enumerate(zip(conc_obj, conc_vol)):
    pct = cv/co*100 if co > 0 else 0
    color = "#C62828"  # toutes sous 100%
    ax2.annotate(f'{pct:.1f}%', xy=(i + width/2, cv), xytext=(0, 5),
                textcoords='offset points', ha='center', fontsize=10, fontweight='bold', color=color)
    # Valeurs
    ax2.annotate(f'{co:,.0f}'.replace(',', ' '), xy=(i - width/2, co), xytext=(0, 3),
                textcoords='offset points', ha='center', fontsize=8, color=GRAY)
    ax2.annotate(f'{cv:,.0f}'.replace(',', ' '), xy=(i + width/2, cv), xytext=(0, -12),
                textcoords='offset points', ha='center', fontsize=8, color=GOLD, fontweight='bold')

plt.savefig('/home/z/my-project/scripts/pdf_charts/chartAA_region_ventes_vs_obj.png', dpi=150, facecolor='white')
plt.close()
print("✓ chartAA_region_ventes_vs_obj.png généré avec données corrigées")

# Vérification
print(f"\nDonnées utilisées:")
for r in regions:
    d = data[r]
    print(f"  {r}: vol obj={d['vol_obj']:.0f} t, vol réel={d['vol_real']:.0f} t ({d['vol_real']/d['vol_obj']*100:.1f}%) | conc obj={d['conc_obj']:.0f} t, conc réel={d['conc_vol']:.0f} t ({d['conc_vol']/d['conc_obj']*100:.1f}%)")
