"""
Phase Execute - Génération des graphiques de visualisation pour les PDFs.
"""
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
import os

fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

os.makedirs("/home/z/my-project/download/forecast_q4_2026/charts", exist_ok=True)

# Copy AED charts
import shutil
for f in os.listdir("/home/z/my-project/scripts/aed_charts"):
    shutil.copy(f"/home/z/my-project/scripts/aed_charts/{f}", 
                f"/home/z/my-project/download/forecast_q4_2026/charts/{f}")

# Load forecast
fcst = pd.read_csv("/home/z/my-project/scripts/forecast_q4_2026.csv", parse_dates=['date'])

# Colors
SCEN_COLORS = {
    'S1_rupture': '#FCE4D6',
    'S2_reappro_50': '#FFF2CC',
    'S3_reappro_100': '#C6EFCE',
    'S4_baisse_prix': '#BDD7EE',
}
SCEN_LABELS = {
    'S1_rupture': 'S1 Rupture totale',
    'S2_reappro_50': 'S2 Réappro 50%',
    'S3_reappro_100': 'S3 Réappro 100%',
    'S4_baisse_prix': 'S4 Baisse prix',
}

# Chart 7: Comparaison volumes par scénario
fig, ax = plt.subplots(figsize=(10, 6), constrained_layout=True)
scenarios = ['S1_rupture', 'S2_reappro_50', 'S3_reappro_100', 'S4_baisse_prix']
synth_t = fcst.groupby('scenario')['tonnes'].sum()
colors = [SCEN_COLORS[s] for s in scenarios]
bars = ax.bar([SCEN_LABELS[s] for s in scenarios], [synth_t[s] for s in scenarios], 
              color=colors, edgecolor='black', linewidth=1.5)
ax.set_title('BELGOCAM - Volume total Q4 2026 par scénario', fontsize=13, fontweight='bold')
ax.set_ylabel('Tonnes')
for bar, s in zip(bars, scenarios):
    height = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., height + 500,
            f'{height:,.0f} t', ha='center', va='bottom', fontweight='bold')
ax.grid(True, alpha=0.3, axis='y')
plt.savefig("/home/z/my-project/download/forecast_q4_2026/charts/07_volume_par_scenario.png", dpi=120)
plt.close()
print("Chart 7 saved")

# Chart 8: Comparaison CA par scénario
fig, ax = plt.subplots(figsize=(10, 6), constrained_layout=True)
synth_ca = fcst.groupby('scenario')['ca_m_fcfa'].sum()
bars = ax.bar([SCEN_LABELS[s] for s in scenarios], [synth_ca[s] for s in scenarios], 
              color=colors, edgecolor='black', linewidth=1.5)
ax.set_title('BELGOCAM - CA total Q4 2026 par scénario', fontsize=13, fontweight='bold')
ax.set_ylabel('M FCFA')
for bar, s in zip(bars, scenarios):
    height = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., height + 200,
            f'{height:,.0f} M', ha='center', va='bottom', fontweight='bold')
ax.grid(True, alpha=0.3, axis='y')
plt.savefig("/home/z/my-project/download/forecast_q4_2026/charts/08_ca_par_scenario.png", dpi=120)
plt.close()
print("Chart 8 saved")

# Chart 9: Évolution mensuelle par scénario
fig, ax = plt.subplots(figsize=(12, 6), constrained_layout=True)
for scn in scenarios:
    monthly_t = fcst[fcst['scenario'] == scn].groupby('month')['tonnes'].sum()
    ax.plot(monthly_t.index, monthly_t.values, marker='o', label=SCEN_LABELS[scn], 
            color=SCEN_COLORS[scn], linewidth=2.5, markersize=10, markeredgecolor='black')
ax.set_title('BELGOCAM - Évolution mensuelle du volume Q4 2026 par scénario', fontsize=13, fontweight='bold')
ax.set_xlabel('Mois')
ax.set_ylabel('Tonnes')
ax.set_xticks([9, 10, 11, 12])
ax.set_xticklabels(['Septembre', 'Octobre', 'Novembre', 'Décembre'])
ax.legend(loc='best')
ax.grid(True, alpha=0.3)
plt.savefig("/home/z/my-project/download/forecast_q4_2026/charts/09_evolution_mensuelle.png", dpi=120)
plt.close()
print("Chart 9 saved")

# Chart 10: Volume par famille et scénario
fig, ax = plt.subplots(figsize=(12, 6), constrained_layout=True)
synth_fam = fcst.groupby(['scenario', 'family'])['tonnes'].sum().unstack(fill_value=0)
families = ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'MAIS', 'ALIMENT_COMPLET']
synth_fam = synth_fam[families]
synth_fam.T.plot(kind='bar', ax=ax, edgecolor='black')
ax.set_title('BELGOCAM - Volume Q4 2026 par famille et scénario', fontsize=13, fontweight='bold')
ax.set_xlabel('Famille')
ax.set_ylabel('Tonnes')
ax.set_xticklabels(ax.get_xticklabels(), rotation=0)
ax.legend([SCEN_LABELS[s] for s in scenarios], loc='best')
ax.grid(True, alpha=0.3, axis='y')
plt.savefig("/home/z/my-project/download/forecast_q4_2026/charts/10_par_famille_scenario.png", dpi=120)
plt.close()
print("Chart 10 saved")

# Chart 11: Volume par région et scénario
fig, ax = plt.subplots(figsize=(10, 6), constrained_layout=True)
synth_reg = fcst.groupby(['scenario', 'region'])['tonnes'].sum().unstack(fill_value=0)
synth_reg = synth_reg[['Ouest', 'Centre', 'Littoral']]
synth_reg.T.plot(kind='bar', ax=ax, edgecolor='black')
ax.set_title('BELGOCAM - Volume Q4 2026 par région et scénario', fontsize=13, fontweight='bold')
ax.set_xlabel('Région')
ax.set_ylabel('Tonnes')
ax.set_xticklabels(ax.get_xticklabels(), rotation=0)
ax.legend([SCEN_LABELS[s] for s in scenarios], loc='best')
ax.grid(True, alpha=0.3, axis='y')
plt.savefig("/home/z/my-project/download/forecast_q4_2026/charts/11_par_region_scenario.png", dpi=120)
plt.close()
print("Chart 11 saved")

# Chart 12: Top 10 agences par CA (scénario S3)
fig, ax = plt.subplots(figsize=(12, 7), constrained_layout=True)
s3 = fcst[fcst['scenario'] == 'S3_reappro_100']
top_ag = s3.groupby(['agence', 'region'])['ca_m_fcfa'].sum().reset_index().sort_values('ca_m_fcfa', ascending=False).head(10)
colors_reg = {'Ouest': '#1F4E78', 'Centre': '#C9A961', 'Littoral': '#A0522D'}
colors = [colors_reg.get(r, '#888') for r in top_ag['region']]
ax.barh(top_ag['agence'], top_ag['ca_m_fcfa'], color=colors, edgecolor='black')
ax.set_title('BELGOCAM - Top 10 agences par CA Q4 2026 (Scénario S3 - Réappro 100%)', fontsize=13, fontweight='bold')
ax.set_xlabel('CA (M FCFA)')
ax.invert_yaxis()
from matplotlib.patches import Patch
legend_elements = [Patch(facecolor=c, label=r) for r, c in colors_reg.items()]
ax.legend(handles=legend_elements, loc='lower right')
ax.grid(True, alpha=0.3, axis='x')
plt.savefig("/home/z/my-project/download/forecast_q4_2026/charts/12_top_agences_ca.png", dpi=120)
plt.close()
print("Chart 12 saved")

# Chart 13: PACE methodology flowchart
fig, ax = plt.subplots(figsize=(14, 4), constrained_layout=True)
ax.set_xlim(0, 14)
ax.set_ylim(0, 4)
ax.axis('off')

phases = [
    (1, 'P - PREPARE', 'Préparation\ndes données', '#1F4E78'),
    (4, 'A - ANALYZE', 'Analyse\nexploratoire', '#C9A961'),
    (7, 'C - CONSTRUCT', 'Construction\nmodèle Prophet', '#A0522D'),
    (10, 'E - EXECUTE', 'Génération\nlivrables', '#7B68EE'),
]

for x, title, desc, color in phases:
    # Box
    rect = plt.Rectangle((x, 1), 3, 2, facecolor=color, edgecolor='black', linewidth=2, alpha=0.7)
    ax.add_patch(rect)
    # Title
    ax.text(x + 1.5, 2.5, title, ha='center', va='center', fontsize=14, fontweight='bold', color='white')
    # Description
    ax.text(x + 1.5, 1.7, desc, ha='center', va='center', fontsize=10, color='white')
    # Arrow
    if x < 10:
        ax.annotate('', xy=(x + 3.5, 2), xytext=(x + 3, 2),
                    arrowprops=dict(arrowstyle='->', lw=2, color='black'))

ax.set_title('BELGOCAM - Méthodologie PACE pour le Forecast Q4 2026', fontsize=15, fontweight='bold', pad=20)
plt.savefig("/home/z/my-project/download/forecast_q4_2026/charts/13_pace_flowchart.png", dpi=120)
plt.close()
print("Chart 13 saved")

print("\n=== ALL CHARTS SAVED ===")
print(f"Location: /home/z/my-project/download/forecast_q4_2026/charts/")
