"""
Phase Analyze (AED) - Analyse Exploratoire des Données.
- Statistiques descriptives
- Saisonnalité mensuelle
- Tendance
- Décomposition par famille, produit, agence, région
- Génération de graphiques pour le guide méthodologique
"""
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import json
import os

# Font setup
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# Load data
df = pd.read_csv("/home/z/my-project/scripts/dataset_consolide.csv", parse_dates=['date'])
print(f"Loaded {len(df)} records from {df['date'].min().date()} to {df['date'].max().date()}")

# Create output dir
os.makedirs("/home/z/my-project/scripts/aed_charts", exist_ok=True)
os.makedirs("/home/z/my-project/download/forecast_q4_2026/charts", exist_ok=True)

# === 1. Volumes mensuels par famille ===
print("\n=== 1. VOLUMES MENSUELS PAR FAMILLE (tonnes) ===")
df['year_month'] = df['date'].dt.to_period('M')
monthly_family = df.groupby(['year_month', 'family'])['tonnes'].sum().unstack(fill_value=0)
print(monthly_family.tail(12))

# === 2. Statistiques descriptives par famille ===
print("\n=== 2. STATISTIQUES PAR FAMILLE (2025-2026) ===")
stats_family = df.groupby('family').agg(
    tonnes_total=('tonnes', 'sum'),
    tonnes_moy_mois=('tonnes', lambda x: x.sum() / df['year_month'].nunique()),
    n_cmdes=('tonnes', 'count'),
    n_clients=('client', 'nunique'),
    n_agences=('agence', 'nunique'),
    n_produits=('ref', 'nunique'),
).round(1)
print(stats_family)

# === 3. Saisonnalité mensuelle (2025) ===
print("\n=== 3. SAISONNALITÉ MENSUELLE (2025) ===")
df_2025 = df[df['year'] == 2025].copy()
monthly_2025 = df_2025.groupby(['month', 'family'])['tonnes'].sum().unstack(fill_value=0)
# Calculate seasonal index (month / annual avg)
annual_avg_2025 = df_2025.groupby('family')['tonnes'].sum() / 12
seasonal_index = monthly_2025.div(annual_avg_2025, axis=1).round(2)
print("\nIndice saisonnalité (1.0 = moyenne annuelle):")
print(seasonal_index)

# === 4. Top produits par famille ===
print("\n=== 4. TOP PRODUITS PAR FAMILLE ===")
top_products = df.groupby(['family', 'ref', 'description'])['tonnes'].sum().reset_index()
top_products = top_products.sort_values(['family', 'tonnes'], ascending=[True, False])
print(top_products.groupby('family').head(3).to_string(index=False))

# === 5. Volumes par agence ===
print("\n=== 5. VOLUMES PAR AGENCE (2025 + S1 2026 + Juil-Août 2026) ===")
vol_agence = df.groupby(['agence', 'region'])['tonnes'].sum().reset_index()
vol_agence = vol_agence.sort_values('tonnes', ascending=False)
print(vol_agence.to_string(index=False))

# === 6. Tendance Q4 2025 (pour validation forecast Q4 2026) ===
print("\n=== 6. TENDANCE Q4 2025 (sept-déc) ===")
df_q4_2025 = df_2025[df_2025['month'].isin([9, 10, 11, 12])]
q4_2025 = df_q4_2025.groupby(['month', 'family'])['tonnes'].sum().unstack(fill_value=0)
print(q4_2025)

# Q4 2025 represents what % of full year 2025?
q4_share = df_q4_2025.groupby('family')['tonnes'].sum() / df_2025.groupby('family')['tonnes'].sum() * 100
print("\nPart du Q4 dans l'année 2025 (%):")
print(q4_share.round(1))

# === GRAPHIQUES ===

# Chart 1: Volumes mensuels par famille
fig, ax = plt.subplots(figsize=(14, 6), constrained_layout=True)
monthly_family.plot(ax=ax, marker='o')
ax.set_title('BELGOCAM - Volumes mensuels par famille (2025 - août 2026)', fontsize=13, fontweight='bold')
ax.set_xlabel('Mois')
ax.set_ylabel('Tonnes')
ax.legend(loc='best', fontsize=9)
ax.grid(True, alpha=0.3)
plt.savefig("/home/z/my-project/scripts/aed_charts/01_volumes_mensuels_famille.png", dpi=120)
plt.close()
print("\nChart 1 saved: 01_volumes_mensuels_famille.png")

# Chart 2: Saisonnalité mensuelle (indice 2025)
fig, ax = plt.subplots(figsize=(12, 6), constrained_layout=True)
seasonal_index.plot(ax=ax, marker='o')
ax.axhline(y=1.0, color='black', linestyle='--', alpha=0.5, label='Moyenne annuelle')
ax.set_title('BELGOCAM - Indice de saisonnalité mensuelle (base 2025)', fontsize=13, fontweight='bold')
ax.set_xlabel('Mois')
ax.set_ylabel('Indice (1.0 = moyenne annuelle)')
ax.set_xticks(range(1, 13))
ax.set_xticklabels(['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc'])
ax.legend(loc='best', fontsize=9)
ax.grid(True, alpha=0.3)
plt.savefig("/home/z/my-project/scripts/aed_charts/02_saisonnalite_mensuelle.png", dpi=120)
plt.close()
print("Chart 2 saved: 02_saisonnalite_mensuelle.png")

# Chart 3: Top 10 agences par volume total
top_agences = vol_agence.head(10)
fig, ax = plt.subplots(figsize=(12, 6), constrained_layout=True)
colors = ['#1F4E78' if r == 'Ouest' else '#C9A961' if r == 'Centre' else '#A0522D' for r in top_agences['region']]
ax.barh(top_agences['agence'], top_agences['tonnes'], color=colors)
ax.set_title('BELGOCAM - Top 10 agences par volume total (2025 - août 2026)', fontsize=13, fontweight='bold')
ax.set_xlabel('Tonnes')
ax.invert_yaxis()
# Legend
from matplotlib.patches import Patch
legend_elements = [
    Patch(facecolor='#1F4E78', label='Ouest'),
    Patch(facecolor='#C9A961', label='Centre'),
    Patch(facecolor='#A0522D', label='Littoral'),
]
ax.legend(handles=legend_elements, loc='lower right')
ax.grid(True, alpha=0.3, axis='x')
plt.savefig("/home/z/my-project/scripts/aed_charts/03_top_agences.png", dpi=120)
plt.close()
print("Chart 3 saved: 03_top_agences.png")

# Chart 4: Top 15 produits par volume
top_produits = df.groupby(['ref', 'description', 'family'])['tonnes'].sum().reset_index().sort_values('tonnes', ascending=False).head(15)
fig, ax = plt.subplots(figsize=(12, 7), constrained_layout=True)
colors_fam = {'TOURTEAUX': '#1F4E78', 'CONCENTRES': '#C9A961', 'INGREDIENTS': '#A0522D', 'ALIMENT_COMPLET': '#7B68EE', 'MAIS': '#FFA500'}
colors = [colors_fam.get(f, '#888888') for f in top_produits['family']]
ax.barh(top_produits['description'], top_produits['tonnes'], color=colors)
ax.set_title('BELGOCAM - Top 15 produits par volume (2025 - août 2026)', fontsize=13, fontweight='bold')
ax.set_xlabel('Tonnes')
ax.invert_yaxis()
legend_elements = [Patch(facecolor=c, label=f) for f, c in colors_fam.items()]
ax.legend(handles=legend_elements, loc='lower right')
ax.grid(True, alpha=0.3, axis='x')
plt.savefig("/home/z/my-project/scripts/aed_charts/04_top_produits.png", dpi=120)
plt.close()
print("Chart 4 saved: 04_top_produits.png")

# Chart 5: Évolution soja vs concentrés (ratio bundle)
monthly_soja = df[df['family'] == 'TOURTEAUX'].groupby('year_month')['sacs_50'].sum()
monthly_conc = df[df['family'] == 'CONCENTRES'].groupby('year_month')['sacs_50'].sum()
ratio_bundle = (monthly_soja / monthly_conc).round(2)
fig, ax1 = plt.subplots(figsize=(14, 6), constrained_layout=True)
ax1.bar(monthly_soja.index.astype(str), monthly_soja.values, alpha=0.7, label='Soja (sacs)', color='#1F4E78')
ax1.bar(monthly_conc.index.astype(str), monthly_conc.values, alpha=0.7, label='Concentrés (sacs)', color='#C9A961')
ax1.set_ylabel('Sacs éq. 50kg')
ax1.set_xlabel('Mois')
ax1.legend(loc='upper left')
ax1.tick_params(axis='x', rotation=90)
ax2 = ax1.twinx()
ax2.plot(ratio_bundle.index.astype(str), ratio_bundle.values, color='red', marker='o', label='Ratio soja:conc', linewidth=2)
ax2.set_ylabel('Ratio soja:conc', color='red')
ax2.tick_params(axis='y', labelcolor='red')
ax2.axhline(y=3, color='red', linestyle='--', alpha=0.5)
ax2.legend(loc='upper right')
plt.title('BELGOCAM - Évolution du ratio bundle soja:concentrés (2025 - août 2026)', fontsize=13, fontweight='bold')
plt.savefig("/home/z/my-project/scripts/aed_charts/05_ratio_bundle.png", dpi=120)
plt.close()
print("Chart 5 saved: 05_ratio_bundle.png")

# Chart 6: Volumes Q4 2025 par famille (pour validation forecast)
fig, ax = plt.subplots(figsize=(10, 6), constrained_layout=True)
q4_2025.T.plot(kind='bar', ax=ax)
ax.set_title('BELGOCAM - Volumes Q4 2025 par famille et par mois', fontsize=13, fontweight='bold')
ax.set_xlabel('Famille')
ax.set_ylabel('Tonnes')
ax.set_xticklabels(ax.get_xticklabels(), rotation=0)
ax.legend(['Sept', 'Oct', 'Nov', 'Déc'], loc='best')
ax.grid(True, alpha=0.3, axis='y')
plt.savefig("/home/z/my-project/scripts/aed_charts/06_q4_2025_par_famille.png", dpi=120)
plt.close()
print("Chart 6 saved: 06_q4_2025_par_famille.png")

# === Save AED summary to JSON ===
aed_summary = {
    'data_period': f"{df['date'].min().date()} to {df['date'].max().date()}",
    'total_records': len(df),
    'total_tonnes': round(df['tonnes'].sum(), 1),
    'total_ca_ttc_m': round(df['montant_ttc'].sum() / 1e6, 1),
    'n_products': int(df['ref'].nunique()),
    'n_agences': int(df['agence'].nunique()),
    'n_regions': int(df['region'].nunique()),
    'n_clients': int(df['client'].nunique()),
    'stats_by_family': stats_family.to_dict(orient='index'),
    'seasonal_index': seasonal_index.to_dict(),
    'q4_2025_share_pct': q4_share.round(1).to_dict(),
    'top_10_agences': vol_agence.head(10).to_dict(orient='records'),
    'top_15_produits': top_produits.to_dict(orient='records'),
}
with open("/home/z/my-project/scripts/aed_summary.json", 'w') as f:
    json.dump(aed_summary, f, indent=2, default=str)
print("\nAED summary saved: /home/z/my-project/scripts/aed_summary.json")
print("\n=== AED COMPLETE ===")
