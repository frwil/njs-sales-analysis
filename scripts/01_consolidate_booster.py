"""
Consolidation des données de ventes Chick Booster + Piglet Booster
Période: Octobre 2025 - Septembre 2026
Sources:
  - 86d96135...xlsx : Jan-Dec 2025 (avec Agence + Region)
  - ventes janv a juin 2026.xlsx : Jan-Jun 2026 (avec agence, sans region)
  - NJS GROUP ERP (40).xlsx : Sep 2026 (latest, plus complet)
  - NJS GROUP ERP (25).xlsx : Août 2026 complet
  - NJS GROUP ERP (9).xlsx : Juillet 2026 complet
"""
import pandas as pd
import os
import warnings
warnings.filterwarnings('ignore')

UPLOAD = "/home/z/my-project/upload"
WORK = "/home/z/my-project/work"
os.makedirs(WORK, exist_ok=True)

# ============================================================
# 1. Load 2025 full year data (file 86d96135...)
# ============================================================
print("=" * 70)
print("1. Loading 2025 full year data...")
fp25 = os.path.join(UPLOAD, "86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx")
df25 = pd.read_excel(fp25, sheet_name="Feuil1")
df25['Date de commande'] = pd.to_datetime(df25['Date de commande'], errors='coerce')
df25['source'] = '2025_full'
# Standardize columns: rename to common schema
df25 = df25.rename(columns={
    'tableauProprieteAgences.Agence': 'agence_raw',
    'tableauProprieteAgences.Region': 'region_raw',
    'Qté commandée (en tonnes)': 'qte_tonnes',
    'Qté commandée': 'qte_sacs',
})
print(f"   2025 rows: {len(df25):,}")

# ============================================================
# 2. Load H1 2026 (ventes janv a juin 2026.xlsx) - 6 monthly sheets
# ============================================================
print("\n2. Loading H1 2026 data (6 monthly sheets)...")
fp_h1 = os.path.join(UPLOAD, "ventes janv a juin 2026.xlsx")
sheets_h1 = ['Sheet 1', 'Feuil1', 'Feuil2', 'Feuil3', 'Feuil4', 'Feuil5']
df_h1_list = []
for sh in sheets_h1:
    df_m = pd.read_excel(fp_h1, sheet_name=sh, header=1)
    df_m['source_sheet'] = sh
    df_h1_list.append(df_m)
df_h1 = pd.concat(df_h1_list, ignore_index=True)
df_h1['Date de commande'] = pd.to_datetime(df_h1['Date de commande'], errors='coerce', dayfirst=True)
df_h1['source'] = '2026_h1'
# Rename for consistency
df_h1 = df_h1.rename(columns={'agence': 'agence_raw'})
df_h1['region_raw'] = None  # Will be mapped later
df_h1['qte_tonnes'] = None  # Will be computed later
df_h1 = df_h1.rename(columns={'Qté commandée': 'qte_sacs'})
print(f"   H1 2026 rows: {len(df_h1):,}")

# ============================================================
# 3. Load H2 2026 (Jul, Aug, Sep from NJS GROUP ERP)
# ============================================================
print("\n3. Loading H2 2026 ERP files (Jul, Aug, Sep)...")
# Jul: file (9).xlsx (Jul 1-31, most complete)
# Aug: file (25).xlsx (Aug 1-31, complete)
# Sep: file (6) (1).xlsx (Sep 1-19, latest)
h2_files = [
    ("Jul", os.path.join(UPLOAD, "NJS GROUP ERP - Lignes de commandes + multicompany (9).xlsx")),
    ("Aug", os.path.join(UPLOAD, "NJS GROUP ERP - Lignes de commandes + multicompany (25).xlsx")),
    ("Sep", os.path.join(UPLOAD, "NJS GROUP ERP - Lignes de commandes + multicompany (6) (1).xlsx")),
]
df_h2_list = []
for month_label, fp in h2_files:
    print(f"   Loading {month_label}: {os.path.basename(fp)}")
    df_m = pd.read_excel(fp, sheet_name="Sheet 1", header=1)
    df_m['source'] = f'2026_{month_label.lower()}'
    df_h2_list.append(df_m)
df_h2 = pd.concat(df_h2_list, ignore_index=True)
df_h2['Date de commande'] = pd.to_datetime(df_h2['Date de commande'], errors='coerce', dayfirst=True)
df_h2 = df_h2.rename(columns={'agence': 'agence_raw'})
df_h2['region_raw'] = None
df_h2['qte_tonnes'] = None
df_h2 = df_h2.rename(columns={'Qté commandée': 'qte_sacs'})
print(f"   H2 2026 rows: {len(df_h2):,}")

# ============================================================
# 4. Concatenate all
# ============================================================
print("\n4. Concatenating all sources...")
# Keep only common columns + key info
common_cols = ['Description du produit', 'qte_sacs', 'Date de commande', 'Montant HT',
               'agence_raw', 'region_raw', 'qte_tonnes', 'source']
# Add 'source' to df25
df25['source'] = '2025_full'

# Ensure all DFs have the columns
for df_x in [df25, df_h1, df_h2]:
    for c in common_cols:
        if c not in df_x.columns:
            df_x[c] = None

df_all = pd.concat([df25[common_cols], df_h1[common_cols], df_h2[common_cols]],
                   ignore_index=True)
print(f"   Total rows: {len(df_all):,}")
print(f"   Date range: {df_all['Date de commande'].min()} to {df_all['Date de commande'].max()}")

# ============================================================
# 5. Filter to Booster products only
# ============================================================
print("\n5. Filtering Booster products...")
mask = df_all['Description du produit'].astype(str).str.contains('booster', case=False, na=False)
df_booster = df_all[mask].copy()
print(f"   Booster rows: {len(df_booster):,}")
print(f"   Unique products: {df_booster['Description du produit'].unique()}")

# ============================================================
# 6. Normalize agence names (strip "AGENCE " prefix, etc.)
# ============================================================
print("\n6. Normalizing agence names...")
def normalize_agence(s):
    if pd.isna(s):
        return None
    s = str(s).strip()
    # Remove "AGENCE " prefix
    if s.startswith('AGENCE '):
        s = s[7:]
    # Special cases
    if s.startswith('DE BAMENDA - DEPOT MBOUDA'):
        s = 'Bamenda-Mbouda'
    if s == 'SPC BAF-CHEFFERIE':
        s = 'SPC Chef.'
    # Known acronyms to preserve
    acronyms = {'SPC', 'PDC', 'DLA', 'PK11', 'PK15', 'TPO', 'YASSA', 'BERI', 'NDERE', 'MBOUDA', 'DSCHANG', 'KYE-OSSI'}
    parts = s.split()
    normalized_parts = []
    for p in parts:
        # Hyphenated acronym like "SPC-YASSA" or "SPC-DLA-BERI"
        if '-' in p and any(part.upper() in acronyms for part in p.split('-')):
            # Keep as-is if it contains acronyms
            normalized_parts.append(p)
        elif p.isupper() and len(p) > 4:
            # Likely a normal word in uppercase -> Title Case (e.g., MESSASSI -> Messassi, NGAOUNDERE -> Ngaoundere)
            normalized_parts.append(p.title())
        elif p.isupper() and len(p) <= 4 and p.upper() not in acronyms:
            # Short word, not an acronym (e.g., BUEA -> Buea)
            normalized_parts.append(p.title())
        else:
            # Acronym or already mixed case - keep as-is
            normalized_parts.append(p)
    return ' '.join(normalized_parts)

df_booster['agence'] = df_booster['agence_raw'].apply(normalize_agence)

# Build region mapping from 2025 data (goldmine file)
region_map = (
    df25.dropna(subset=['agence_raw', 'region_raw'])
        .assign(agence_norm=lambda d: d['agence_raw'].apply(normalize_agence))
        .drop_duplicates(['agence_norm', 'region_raw'])
        .set_index('agence_norm')['region_raw']
        .to_dict()
)
print(f"   Region mapping built for {len(region_map)} agences:")
for k, v in sorted(region_map.items(), key=lambda x: (str(x[1]), x[0])):
    print(f"     {k:25s} -> {v}")

# Apply region mapping
df_booster['region'] = df_booster['agence'].map(region_map)

# For unmapped agences, assign based on geography (manual fallback)
fallback_region = {
    'Bamenda-Mbouda': 'Ouest',
    'Buea': 'Littoral',  # Match existing data, even if geographically Sud-Ouest
    'SPC BUEA': 'Littoral',
    'SPC VILLAGE': 'Littoral',
    'SPC PK15': 'Littoral',
    'SPC-YASSA': 'Littoral',
    'SPC-DLA-BERI': 'Littoral',
    'SPC-TPO': 'Littoral',
    'SPC-NDERE': 'Centre',
    'SPC KYE-OSSI': 'Centre',
    'SPC-DSCHANG': 'Centre',
    'PDC Emana': 'Centre',
    'PDC BERTOUA': 'Centre',
    'SPC MBOUDA': 'Ouest',
    'SPC Chef.': 'Ouest',
    'SPC DSCHANG': 'Centre',
    'SPC NDERE': 'Centre',
    'SPC TPO': 'Littoral',
}
for ag, reg in fallback_region.items():
    if ag not in region_map:
        region_map[ag] = reg

df_booster['region'] = df_booster['agence'].map(region_map)
# Replace NaN with "Non spécifié"
df_booster['region'] = df_booster['region'].fillna('Non spécifié')

# ============================================================
# 7. Compute qte_tonnes where missing
# ============================================================
print("\n7. Computing volumes in tonnes...")
# Conversion: 25Kg = 0.025 t, 5Kg = 0.005 t
def qte_to_tonnes(row):
    if pd.notna(row['qte_tonnes']) and row['qte_tonnes'] > 0:
        return row['qte_tonnes']
    qte_sacs = row['qte_sacs']
    if pd.isna(qte_sacs):
        return 0
    product = str(row['Description du produit']).upper()
    if '25' in product:
        return qte_sacs * 0.025
    elif '5' in product and '25' not in product:
        return qte_sacs * 0.005
    else:
        return qte_sacs * 0.025  # default to 25Kg

df_booster['qte_tonnes'] = df_booster.apply(qte_to_tonnes, axis=1)

# ============================================================
# 8. Add month column and product category
# ============================================================
df_booster['mois'] = df_booster['Date de commande'].dt.to_period('M').astype(str)
df_booster['annee'] = df_booster['Date de commande'].dt.year
df_booster['mois_num'] = df_booster['Date de commande'].dt.month

# Product family
def product_family(s):
    s = str(s).upper()
    if 'CHICK' in s:
        return 'Chick Booster'
    elif 'PIGLET' in s:
        return 'Piglet Booster'
    return 'Autre'

df_booster['famille'] = df_booster['Description du produit'].apply(product_family)

# Format (5kg vs 25kg)
def product_format(s):
    s = str(s).upper()
    if '25' in s:
        return '25 Kg'
    elif '5' in s and '25' not in s:
        return '5 Kg'
    return 'Autre'

df_booster['format'] = df_booster['Description du produit'].apply(product_format)

# Filter to target period: Oct 2025 - Sep 2026
mask_period = (df_booster['Date de commande'] >= '2025-10-01') & (df_booster['Date de commande'] <= '2026-09-30')
df_booster = df_booster[mask_period].copy()
print(f"\n8. Filtered to Oct 2025 - Sep 2026: {len(df_booster):,} rows")

# ============================================================
# 9. Save consolidated data
# ============================================================
output_cols = ['Date de commande', 'mois', 'annee', 'mois_num', 'agence', 'region',
               'Description du produit', 'famille', 'format', 'qte_sacs', 'qte_tonnes',
               'Montant HT', 'source']
df_booster[output_cols].to_csv(os.path.join(WORK, 'booster_consolidated.csv'), index=False)
print(f"\n9. Saved to {WORK}/booster_consolidated.csv")
print(f"   Final shape: {df_booster.shape}")

# Quick stats
print("\n" + "=" * 70)
print("SUMMARY")
print("=" * 70)
print(f"Total Booster rows: {len(df_booster):,}")
print(f"Date range: {df_booster['Date de commande'].min().date()} to {df_booster['Date de commande'].max().date()}")
print(f"\nVolume total (tonnes): {df_booster['qte_tonnes'].sum():.2f}")
print(f"\nVolume par famille:")
print(df_booster.groupby('famille')['qte_tonnes'].sum().round(2))
print(f"\nVolume par region:")
print(df_booster.groupby('region')['qte_tonnes'].sum().round(2))
print(f"\nVolume par mois:")
print(df_booster.groupby('mois')['qte_tonnes'].sum().round(2))
print(f"\nTop 10 agences par volume:")
print(df_booster.groupby('agence')['qte_tonnes'].sum().sort_values(ascending=False).head(10).round(2))
