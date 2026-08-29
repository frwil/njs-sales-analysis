"""
Analyse SPC - Définition des quantités à commander auprès de SPC
pour approvisionner les agences BELGOCAM en aliment complet.

Approche:
1. Calculer le volume soja "non disponible" par agence (scénarios S1/S3)
2. Convertir en aliment complet (ratio 1:3)
3. Répartir par produit SPC selon le mix BELGOCAM (adjusté vs mix SPC)
4. Étendre aux autres régions proportionnellement au soja
5. Split par format 5kg / 10kg

Scénarios de bascule: 10% (prudent), 30% (réaliste), 50% (optimiste)
Toutes agences sauf BUEA
"""
import pandas as pd
import numpy as np
import json
import os

# Load dataset
df = pd.read_csv("/home/z/my-project/scripts/dataset_consolide_v2.csv", parse_dates=['date'], low_memory=False)
print(f"Loaded {len(df)} records")

# === 1. Volumes soja par agence (historique mensuel moyen) ===
print("\n=== 1. VOLUMES SOJA PAR AGENCE (moyenne mensuelle) ===")
soja = df[df['family'] == 'TOURTEAUX'].copy()
soja['year_month'] = soja['date'].dt.to_period('M')

# Monthly average per agency (last 20 months)
soja_monthly = soja.groupby(['agence', 'region', 'year_month'])['sacs_50'].sum().reset_index()
soja_avg = soja_monthly.groupby(['agence', 'region'])['sacs_50'].mean().round(0).reset_index()
soja_avg.columns = ['agence', 'region', 'sacs_soja_mois']
soja_avg = soja_avg.sort_values('sacs_soja_mois', ascending=False)
print(soja_avg.to_string(index=False))

# Total soja per region
soja_region = soja_avg.groupby('region')['sacs_soja_mois'].sum().round(0)
print(f"\nTotal soja par région (sacs/mois):")
print(soja_region)

# === 2. Mix BELGOCAM par agence (via concentrés) ===
print("\n=== 2. MIX BELGOCAM PAR AGENCE (via concentrés) ===")
# Use concentrés sales to infer the type of elevage
# C101 = BELGO 5% PONTE → ponte
# C102 = BELGO 10% PONTE → ponte
# C103 = BELGO 5% CHAIR → chair
# C104 = BELGO 10% CHAIR → chair
# C105 = BELGO 10% PORC → porc
# C108 = BELGO RUMINANT → autre

conc = df[df['family'] == 'CONCENTRES'].copy()

# Map products to elevage type
def get_elevage_type(ref):
    ref = str(ref)
    if ref in ('C101', 'C102', 'C1022'): return 'PONTE'
    if ref in ('C103', 'C104', 'C1042', 'C1043', 'C1044'): return 'CHAIR'
    if ref in ('C105', 'C1053', 'C1054', 'C1055'): return 'PORC'
    if ref == 'C108': return 'AUTRE'
    return 'AUTRE'

conc['elevage'] = conc['ref'].apply(get_elevage_type)
conc_vol = conc.groupby(['agence', 'region', 'elevage'])['sacs_50'].sum().reset_index()

# Compute mix per agency
conc_pivot = conc_vol.pivot_table(index=['agence', 'region'], columns='elevage', values='sacs_50', fill_value=0)
conc_pivot['TOTAL'] = conc_pivot.sum(axis=1)
for col in ['CHAIR', 'PONTE', 'PORC', 'AUTRE']:
    if col in conc_pivot.columns:
        conc_pivot[f'{col}_pct'] = (conc_pivot[col] / conc_pivot['TOTAL'] * 100).round(1)

print("\nMix BELGOCAM par agence (%):")
print(conc_pivot[['CHAIR_pct', 'PONTE_pct', 'PORC_pct'] if 'CHAIR_pct' in conc_pivot.columns else []].to_string())

# === 3. SPC product mix ===
print("\n=== 3. SPC PRODUCT MIX (fichier Littoral) ===")
# SPC products and their categories
SPC_PRODUCTS = {
    # VOLAILLE CHAIR (BM) - RETENUS
    'BM1': {'cat': 'CHAIR', 'desc': 'BM1 ALIMENT CHAIR', 'spc_vol': 1999, 'spc_pct': 38.5},
    'BM2': {'cat': 'CHAIR', 'desc': 'BM2 ALIMENT CHAIR', 'spc_vol': 2633, 'spc_pct': 50.7},
    'BM3': {'cat': 'CHAIR', 'desc': 'BM3 ALIMENT CHAIR', 'spc_vol': 2376, 'spc_pct': 45.7},
    'BM1E': {'cat': 'CHAIR', 'desc': 'BM1 EXTRA', 'spc_vol': 470, 'spc_pct': 9.0},
    # VOLAILLE PONTE (AM) - RETENUS
    'AM1': {'cat': 'PONTE', 'desc': 'AM1 ALIMENT PONTE', 'spc_vol': 119, 'spc_pct': 22.9},
    'AM2': {'cat': 'PONTE', 'desc': 'AM2 ALIMENT PONTE', 'spc_vol': 15, 'spc_pct': 2.9},
    'AM3': {'cat': 'PONTE', 'desc': 'AM3 ALIMENT PONTE', 'spc_vol': 267, 'spc_pct': 51.4},
}

# SPC mix by category - ONLY CHAIR + PONTE (per DCM directive)
# Recalculated to sum to 100% within the retained categories
spc_cat_mix = {
    'CHAIR': 95.2,   # BM1+BM2+BM3+BM1E = 7478 / (7478+401) = 95.2%
    'PONTE': 4.8,    # AM1+AM2+AM3 = 401 / (7478+401) = 4.8%
}

print("Mix SPC Littoral par catégorie:")
for cat, pct in sorted(spc_cat_mix.items(), key=lambda x: -x[1]):
    print(f"  {cat}: {pct}%")

# === 4. BELGOCAM-adjusted mix ===
print("\n=== 4. MIX BELGOCAM AJUSTÉ (vs SPC) ===")
# Per DCM directive: ONLY CHAIR (BM) + PONTE (AM) products are retained
# Per SPC commercial: "Chair contributes more, then porc, then ponte at SPC"
# "This won't be the same at BELGOCAM" → BELGOCAM has proportionally more PONTE

# Compute BELGOCAM global mix (CHAIR + PONTE only, excluding PORC and AUTRE)
conc_global = conc.groupby('elevage')['sacs_50'].sum()
chair_global = conc_global.get('CHAIR', 0)
ponte_global = conc_global.get('PONTE', 0)
total_cp = chair_global + ponte_global
belgocam_mix = {
    'CHAIR': round(chair_global / total_cp * 100, 1) if total_cp > 0 else 95.2,
    'PONTE': round(ponte_global / total_cp * 100, 1) if total_cp > 0 else 4.8,
}
print("Mix BELGOCAM global (CHAIR + PONTE uniquement, via concentrés):")
for cat, pct in sorted(belgocam_mix.items(), key=lambda x: -x[1]):
    print(f"  {cat}: {pct}%")

# Compute adjusted SPC mix per agency (CHAIR + PONTE only)
print("\n=== 5. MIX SPC AJUSTÉ PAR AGENCE (CHAIR + PONTE uniquement) ===")
adjusted_mixes = {}

for _, row in soja_avg.iterrows():
    agence = row['agence']
    region = row['region']
    
    if agence == 'Buea':
        continue  # Excluded
    
    # Get BELGOCAM mix for this agency (CHAIR + PONTE only)
    if (agence, region) in conc_pivot.index:
        chair_vol = conc_pivot.loc[(agence, region), 'CHAIR'] if 'CHAIR' in conc_pivot.columns else 0
        ponte_vol = conc_pivot.loc[(agence, region), 'PONTE'] if 'PONTE' in conc_pivot.columns else 0
        total_cp_ag = chair_vol + ponte_vol
        if total_cp_ag > 0:
            chair_pct = chair_vol / total_cp_ag * 100
            ponte_pct = ponte_vol / total_cp_ag * 100
        else:
            chair_pct, ponte_pct = 95.2, 4.8  # fallback to SPC mix
    else:
        chair_pct, ponte_pct = 95.2, 4.8
    
    adjusted_mix = {
        'CHAIR': chair_pct,
        'PONTE': ponte_pct,
    }
    
    adjusted_mixes[(agence, region)] = adjusted_mix
    
    if agence in ['Ndobo', 'Famla', 'Messassi']:  # Show sample
        print(f"\n  {agence} ({region}):")
        print(f"    BELGOCAM mix: CHAIR={chair_pct:.0f}%, PONTE={ponte_pct:.0f}%")

# === 6. Compute SPC orders ===
print("\n=== 6. CALCUL DES COMMANDES SPC ===")

# Parameters
SUBSTITUTION_RATIO = 3  # 1 sac soja = 3 sacs aliment complet
FORMAT_SPLIT = {'5kg': 0.30, '10kg': 0.70}  # 30% en 5kg, 70% en 10kg
SCENARIOS = {'prudent': 0.10, 'realiste': 0.30, 'optimiste': 0.50}

# For SPC products, within each category, use SPC proportions
# ONLY CHAIR + PONTE (per DCM directive)
spc_within_cat = {}
for cat in ['CHAIR', 'PONTE']:
    cat_products = {ref: info for ref, info in SPC_PRODUCTS.items() if info['cat'] == cat}
    cat_total = sum(info['spc_vol'] for info in cat_products.values())
    if cat_total > 0:
        spc_within_cat[cat] = {ref: info['spc_vol'] / cat_total for ref, info in cat_products.items()}
    else:
        spc_within_cat[cat] = {ref: 1.0 / len(cat_products) for ref in cat_products}

all_orders = []

for _, row in soja_avg.iterrows():
    agence = row['agence']
    region = row['region']
    sacs_soja = row['sacs_soja_mois']
    
    if agence == 'Buea':
        continue
    
    mix = adjusted_mixes.get((agence, region), {})
    
    for scn_name, scn_rate in SCENARIOS.items():
        # Volume soja non disponible (substitution demand)
        sacs_soja_subst = sacs_soja * scn_rate
        
        # Convert to aliment complet (1:3 ratio)
        # But only the "food" portion converts to BM/AM/P
        # The "other" (lapin, poisson, etc.) is independent demand
        sacs_aliment_food = sacs_soja_subst * SUBSTITUTION_RATIO  # Total food sacs
        
        # Split by category using adjusted mix
        for cat, cat_pct in mix.items():
            cat_sacs = sacs_aliment_food * cat_pct / 100
            
            # Split by product within category
            for ref, within_pct in spc_within_cat[cat].items():
                product_sacs = cat_sacs * within_pct
                
                # Split by format
                for fmt, fmt_pct in FORMAT_SPLIT.items():
                    # Convert to actual units
                    # 5kg format: 1 sac 50kg eq = 10 sacs de 5kg
                    # 10kg format: 1 sac 50kg eq = 5 sacs de 10kg
                    if fmt == '5kg':
                        units = product_sacs * fmt_pct * 10  # 10 sacs de 5kg per sac 50kg eq
                    else:  # 10kg
                        units = product_sacs * fmt_pct * 5  # 5 sacs de 10kg per sac 50kg eq
                    
                    all_orders.append({
                        'agence': agence,
                        'region': region,
                        'scenario': scn_name,
                        'taux_bascule': f"{int(scn_rate*100)}%",
                        'produit_spc': ref,
                        'description': SPC_PRODUCTS[ref]['desc'],
                        'categorie': SPC_PRODUCTS[ref]['cat'],
                        'format': fmt,
                        'sacs_50_eq': round(product_sacs * fmt_pct, 1),
                        'unites_commander': round(units, 0),
                        'sacs_soja_substitue': round(sacs_soja_subst, 0),
                    })

orders_df = pd.DataFrame(all_orders)
print(f"\n{len(orders_df)} lignes de commandes générées")

# === Synthèse ===
print("\n=== SYNTHÈSE PAR SCÉNARIO ===")
for scn in ['prudent', 'realiste', 'optimiste']:
    sub = orders_df[orders_df['scenario'] == scn]
    total_units = sub['unites_commander'].sum()
    total_sacs_50 = sub['sacs_50_eq'].sum()
    print(f"\n  {scn.upper()} (taux {sub['taux_bascule'].iloc[0]}):")
    print(f"    Total unités à commander: {total_units:,.0f}")
    print(f"    Total sacs éq. 50kg: {total_sacs_50:,.0f}")
    
    # By category
    print(f"    Par catégorie:")
    cat_sum = sub.groupby('categorie')['unites_commander'].sum().sort_values(ascending=False)
    for cat, units in cat_sum.items():
        print(f"      {cat}: {units:,.0f} unités ({units/total_units*100:.1f}%)")

# By agency (realiste scenario)
print("\n=== SYNTHÈSE PAR AGENCE (SCÉNARIO RÉALISTE 30%) ===")
realiste = orders_df[orders_df['scenario'] == 'realiste']
ag_sum = realiste.groupby(['agence', 'region']).agg(
    unites=('unites_commander', 'sum'),
    sacs_50_eq=('sacs_50_eq', 'sum'),
).reset_index().sort_values('unites', ascending=False)
print(ag_sum.to_string(index=False))

# By format
print("\n=== SYNTHÈSE PAR FORMAT (SCÉNARIO RÉALISTE 30%) ===")
fmt_sum = realiste.groupby('format')['unites_commander'].sum()
print(fmt_sum)

# === Save ===
output_path = "/home/z/my-project/scripts/spc_orders.csv"
orders_df.to_csv(output_path, index=False)
print(f"\nSaved: {output_path}")

# Save summary
summary = {
    'substitution_ratio': SUBSTITUTION_RATIO,
    'format_split': FORMAT_SPLIT,
    'scenarios': SCENARIOS,
    'excluded_agencies': ['Buea'],
    'spc_cat_mix': spc_cat_mix,
    'belgocam_mix': belgocam_mix,
    'total_orders_lines': len(orders_df),
}
with open("/home/z/my-project/scripts/spc_summary.json", 'w') as f:
    json.dump(summary, f, indent=2, default=str)

print("\n=== ANALYSE SPC COMPLETE ===")
