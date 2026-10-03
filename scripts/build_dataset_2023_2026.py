"""
Construction du dataset consolidé 2023-2026 (4 ans).
Inclut l'année 2023 (extraite de LY_21_24) + le dataset 2024-2026 déjà construit.

CHANGEMENTS vs version 2024-2026:
1. AJOUT de l'année 2023 (Jan-Dec 2023, extrait de LY_21_24 sheet)
2. EXCLUSION de V305 (BELGOKILL 200L) — seul V300 1L conservé pour COMPLEMENT_ALIMENTAIRE

Sources:
  - LY_21_24 (Jan 2021 - Jun 2024): filtre 2023 uniquement
  - dataset_2024_2026 (Jul 2024 - Août 2026): déjà construit

Familles finales (7):
  - TOURTEAUX, CONCENTRES, INGREDIENTS, ALIMENT_COMPLET
  - COMPLEMENT_ALIMENTAIRE (BELGOKILL V300 1L only + autres CA001-CA008)
  - MATERIEL_ELEVAGE (tonnes=0)
  - PREMIX (tonnes=0)

Conversion:
  - 1L = 1kg pour COMPLEMENT_ALIMENTAIRE
  - MATERIEL_ELEVAGE: tonnes=0, CA only
"""
import openpyxl
import pandas as pd
import os

# === Source files ===
SRC_21_24 = "/home/z/my-project/upload/21_24.xlsx"  # LY_21_24 sheet (contains 2023 data)
EXISTING_DATASET = "/home/z/my-project/scripts/dataset_2024_2026.csv"

# === Product family mapping (from catProduit2) ===
def get_family_from_cat(cat):
    cat = str(cat).strip() if cat else ''
    if cat == 'TOURTEAUX': return 'TOURTEAUX'
    if cat == 'CONCENTRES': return 'CONCENTRES'
    if cat == 'ALIMENT COMPLET': return 'ALIMENT_COMPLET'
    if cat == 'INGREDIENTS': return 'INGREDIENTS'
    if cat == 'PREMIX': return 'PREMIX'
    if cat == 'MATERIELS ELEVAGE': return 'MATERIEL_ELEVAGE'
    if cat == 'COMPLEMENT ALIMENTAIRE': return 'COMPLEMENT_ALIMENTAIRE'
    return None  # DIVERS → excluded

# === Region mapping ===
def map_region(reg):
    reg = str(reg).strip() if reg else ''
    if reg in ('Yde', 'Est', 'Nord'): return 'Centre'
    if reg in ('Ouest', 'Littoral', 'Centre'): return reg
    return None

# === Agence mapping for LY_21_24 (Agence2 col 14) ===
AGENCE_MAP_21_24 = {
    'FAMLA': ('Famla', 'Ouest'), 'DJELENG': ('Djeleng', 'Ouest'), 'MBOUDA': ('Mbouda', 'Ouest'),
    'MESSASSI': ('Messassi', 'Centre'), 'AHALA': ('Ahala', 'Centre'),
    'BERTOUA': ('Bertoua', 'Centre'), 'NGAOUNDERE': ('Ngaoundere', 'Centre'),
    'NKOLBISSON': ('Nkolbisson', 'Centre'), 'NKOABANG': ('Nkoabang', 'Centre'),
    'NDOBO': ('Ndobo', 'Littoral'), 'VILLAGE': ('Village', 'Littoral'),
    'PK11': ('Pk11', 'Littoral'), 'NKONGSAMBA': ('Nkongsamba', 'Littoral'),
    'BUEA': ('Buea', 'Littoral'),
}

# Products to exclude
EXCLUDED_REFS = {'M1051', 'M1052', 'V305'}  # MAIS + V305 (BELGOKILL 200L — keep only V300 1L)

# Weight map for COMPLEMENT_ALIMENTAIRE (only V300 1L kept)
COMPLEMENT_WEIGHTS = {
    'V300': 1,        # BELGOKILL 1L = 1 kg
    'CA003.1': 1, 'CA004.1': 1, 'CA006.1': 1, 'CA001.1': 1,
    'CA002.1': 1, 'CA005.1': 1, 'CA007.1': 1, 'CA008.1': 1,
}

# === Load LY_21_24 - filter 2023 only ===
print("=" * 70)
print("Loading LY_21_24 - filtering year 2023...")
print("=" * 70)
wb = openpyxl.load_workbook(SRC_21_24, read_only=True, data_only=True)
ws = wb['LY_21_24']
rows = list(ws.iter_rows(values_only=True))
print(f"  {len(rows)-1} total rows in LY_21_24")

records_2023 = []
for r in rows[1:]:
    if not r or not r[0]: continue
    ref = str(r[0])
    if ref in EXCLUDED_REFS: continue
    cat = str(r[7]).strip() if r[7] else ''
    family = get_family_from_cat(cat)
    if family is None: continue
    
    # Filter 2023 only
    date = r[4]
    if date is None: continue
    if isinstance(date, str):
        try: date = pd.to_datetime(date, errors='coerce')
        except: continue
    if pd.isna(date): continue
    if date.year != 2023: continue
    
    # MATERIEL_ELEVAGE: tonnes=0
    if family == 'MATERIEL_ELEVAGE' and r[8] == 0:
        tonnes = 0  # Non-comptabilisable
    elif family == 'MATERIEL_ELEVAGE':
        tonnes = 0  # Always 0 per user instruction
    elif family == 'COMPLEMENT_ALIMENTAIRE':
        # 1L = 1kg: qte = kg, tonnes = qte/1000
        qte = r[3] if r[3] else 0
        weight = COMPLEMENT_WEIGHTS.get(ref, 1)
        kg = qte * weight
        tonnes = kg / 1000
    else:
        tonnes = r[10] if r[10] else 0  # qteCmdTonne (col 10)
    
    agence_raw = str(r[14]).strip() if r[14] else ''  # Agence2 (col 14)
    if agence_raw not in AGENCE_MAP_21_24: continue
    agence, region_default = AGENCE_MAP_21_24[agence_raw]
    
    region_raw = str(r[11]).strip() if r[11] else region_default
    region = map_region(region_raw) or region_default
    
    montant_ht = r[2] if r[2] else 0  # montantHT (col 2)
    # Règle : une vente dont le montant HT est à 0 n'est pas intégrée
    if not isinstance(montant_ht, (int, float)) or montant_ht <= 0: continue
    qte = r[3] if r[3] else 0
    kg = tonnes * 1000
    sacs_50 = kg / 50 if family not in ('MATERIEL_ELEVAGE', 'COMPLEMENT_ALIMENTAIRE') else 0
    
    records_2023.append({
        'date': date, 'year': date.year, 'month': date.month,
        'ref': ref, 'family': family, 'description': r[1],
        'client': '', 'agence': agence, 'region': region,
        'qte': qte, 'weight_kg': 0, 'kg': kg, 'tonnes': tonnes,
        'sacs_50': sacs_50,
        'montant_ttc': 0, 'montant_ht': montant_ht,
        'source': 'LY_2023', 'interne': False
    })

print(f"  2023 records loaded: {len(records_2023)}")
df_2023 = pd.DataFrame(records_2023)
print(f"  2023 by family:")
print(df_2023['family'].value_counts())
print(f"  2023 COMPLEMENT_ALIMENTAIRE tonnes: {df_2023[df_2023['family']=='COMPLEMENT_ALIMENTAIRE']['tonnes'].sum():.2f}")

# === Load existing 2024-2026 dataset ===
print("\n" + "=" * 70)
print("Loading existing 2024-2026 dataset...")
print("=" * 70)
df_24_26 = pd.read_csv(EXISTING_DATASET, parse_dates=['date'], low_memory=False)
print(f"  {len(df_24_26)} records (Jul 2024 - Aug 2026)")

# Remove V305 from existing dataset (per user request)
v305_before = len(df_24_26)
df_24_26 = df_24_26[df_24_26['ref'] != 'V305']
print(f"  Removed V305 (BELGOKILL 200L): {v305_before - len(df_24_26)} records excluded")
print(f"  After V305 removal: {len(df_24_26)} records")

# Re-verify MATERIEL_ELEVAGE tonnes=0 (defensive)
df_24_26.loc[df_24_26['family'] == 'MATERIEL_ELEVAGE', 'tonnes'] = 0
df_24_26.loc[df_24_26['family'] == 'MATERIEL_ELEVAGE', 'sacs_50'] = 0

# === Merge 2023 + 2024-2026 ===
print("\n" + "=" * 70)
print("Merging 2023 + 2024-2026...")
print("=" * 70)
df_all = pd.concat([df_2023, df_24_26], ignore_index=True)
print(f"Total records: {len(df_all)}")
print(f"Date range: {df_all['date'].min().date()} → {df_all['date'].max().date()}")
print(f"\nRecords by year:")
print(df_all.groupby(df_all['date'].dt.year).size())
print(f"\nRecords by family:")
print(df_all['family'].value_counts())

# Save
output_path = "/home/z/my-project/scripts/dataset_2023_2026.csv"
df_all.to_csv(output_path, index=False)
print(f"\nDataset saved: {output_path}")
print(f"File size: {os.path.getsize(output_path) / 1024:.0f} KB")

# Stats by family
print("\n" + "=" * 70)
print("STATS PAR FAMILLE (2023-2026)")
print("=" * 70)
stats = df_all.groupby('family').agg(
    tonnes=('tonnes', 'sum'),
    ca_ht_m_fcfa=('montant_ht', lambda x: x.sum() / 1e6),
    ca_ttc_m_fcfa=('montant_ttc', lambda x: x.sum() / 1e6),
    n_records=('tonnes', 'count'),
    n_produits=('ref', 'nunique'),
    n_agences=('agence', 'nunique'),
).round({'tonnes': 1, 'ca_ht_m_fcfa': 1, 'ca_ttc_m_fcfa': 1})
print(stats)

# Stats by year
print("\n" + "=" * 70)
print("STATS PAR ANNÉE")
print("=" * 70)
stats_year = df_all.groupby(df_all['date'].dt.year).agg(
    tonnes=('tonnes', 'sum'),
    ca_ht_m_fcfa=('montant_ht', lambda x: x.sum() / 1e6),
    ca_ttc_m_fcfa=('montant_ttc', lambda x: x.sum() / 1e6),
    n_records=('tonnes', 'count'),
).round({'tonnes': 0, 'ca_ht_m_fcfa': 1, 'ca_ttc_m_fcfa': 1})
print(stats_year)

# Stats COMPLEMENT_ALIMENTAIRE par année (verify V305 excluded)
print("\n" + "=" * 70)
print("COMPLEMENT_ALIMENTAIRE par année (V300 1L only)")
print("=" * 70)
ca = df_all[df_all['family'] == 'COMPLEMENT_ALIMENTAIRE']
print(f"Refs in COMPLEMENT_ALIMENTAIRE: {sorted(ca['ref'].unique())}")
print(f"V305 present? {'V305' in ca['ref'].unique()}")
ca_by_year = ca.groupby(ca['date'].dt.year).agg(
    tonnes=('tonnes', 'sum'),
    litres=('qte', 'sum'),
    ca_ht_m=('montant_ht', lambda x: x.sum() / 1e6),
    n_records=('tonnes', 'count'),
).round({'tonnes': 2, 'ca_ht_m': 1})
print(ca_by_year)

print("\n" + "=" * 70)
print(f"DATASET 2023-2026 PRÊT: {len(df_all):,} records, {df_all['ref'].nunique()} produits, {df_all['agence'].nunique()} agences")
print(f"44 mois d'historique (Jan 2023 - Août 2026)")
print(f"MATERIEL_ELEVAGE: tonnes=0 (CA only)")
print(f"COMPLEMENT_ALIMENTAIRE: V300 1L only (V305 200L exclu)")
print("=" * 70)
