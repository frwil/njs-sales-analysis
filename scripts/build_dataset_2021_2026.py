"""
Construction du dataset consolidé 2021-2026 (5,5 ans) avec données 21_24.xlsx.
Intègre les données 2021-2024 dans le dataset v2 existant (2025-2026).

Sources:
  - LY_21_24 (Jan 2021 - Juin 2024): 26,045 rows
  - LY_24 (Jul-Dec 2024): 54,944 rows (excl #N/A)
  - 2025 (Jan-Dec): déjà dans dataset_v2
  - 2026 (Jan-Août): déjà dans dataset_v2

Règles:
  - Agence2 (col 14) pour LY_21_24
  - Col O (14) pour LY_24, exclure #N/A
  - Régions: Yde/Est/Nord → Centre
  - MATERIELS ELEVAGE → MATERIEL_ELEVAGE
  - Exclure DIVERS, COMPLEMENT ALIMENTAIRE
  - Exclure MAIS (produit opportuniste)
  - Montant HT uniquement (pas de TTC)
  - comptabilisable=0 = garder avec tonnes=0 (matériel non exprimable)
"""
import openpyxl
import pandas as pd
from collections import defaultdict
import os

SRC_21_24 = "/home/z/my-project/upload/21_24.xlsx"

# === Product family mapping (from catProduit2) ===
def get_family_from_cat(cat):
    cat = str(cat).strip() if cat else ''
    if cat == 'TOURTEAUX': return 'TOURTEAUX'
    if cat == 'CONCENTRES': return 'CONCENTRES'
    if cat == 'ALIMENT COMPLET': return 'ALIMENT_COMPLET'
    if cat == 'INGREDIENTS': return 'INGREDIENTS'
    if cat == 'PREMIX': return 'PREMIX'
    if cat == 'MATERIELS ELEVAGE': return 'MATERIEL_ELEVAGE'
    return None  # DIVERS, COMPLEMENT ALIMENTAIRE → excluded

# === Region mapping ===
def map_region(reg):
    reg = str(reg).strip() if reg else ''
    if reg in ('Yde', 'Est', 'Nord'): return 'Centre'
    return reg  # Ouest, Littoral

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
EXCLUDED_REFS = {'M1051', 'M1052'}  # MAIS

# === Load LY_21_24 (Jan 2021 - Jun 2024) ===
print("Loading LY_21_24 (Jan 2021 - Jun 2024)...")
wb = openpyxl.load_workbook(SRC_21_24, read_only=True, data_only=True)
ws = wb['LY_21_24']
rows = list(ws.iter_rows(values_only=True))
print(f"  {len(rows)-1} rows")

records_21_24 = []
for r in rows[1:]:
    if not r or not r[0]: continue
    ref = str(r[0])
    if ref in EXCLUDED_REFS: continue
    cat = str(r[7]).strip() if r[7] else ''
    family = get_family_from_cat(cat)
    if family is None: continue  # DIVERS, COMPLEMENT ALIMENTAIRE
    if family == 'MATERIEL_ELEVAGE' and r[8] == 0:
        # Non-comptabilisable material - keep with tonnes=0
        tonnes = 0
    else:
        tonnes = r[10] if r[10] else 0  # qteCmdTonne (col 10)
    
    agence_raw = str(r[14]).strip() if r[14] else ''  # Agence2 (col 14)
    if agence_raw not in AGENCE_MAP_21_24: continue
    agence, region_default = AGENCE_MAP_21_24[agence_raw]
    
    # Region from col 11, mapped
    region_raw = str(r[11]).strip() if r[11] else region_default
    region = map_region(region_raw)
    
    montant_ht = r[2] if r[2] else 0  # montantHT (col 2)
    date = r[4]  # dateCmd (col 4)
    if date is None: continue
    if isinstance(date, str):
        date = pd.to_datetime(date, errors='coerce')
    if pd.isna(date): continue
    
    qte = r[3] if r[3] else 0  # qteCmd (col 3)
    kg = tonnes * 1000
    sacs_50 = kg / 50 if family != 'MATERIEL_ELEVAGE' else 0
    
    records_21_24.append({
        'date': date, 'year': date.year if hasattr(date, 'year') else pd.Timestamp(date).year,
        'month': date.month if hasattr(date, 'month') else pd.Timestamp(date).month,
        'ref': ref, 'family': family, 'description': r[1],
        'agence': agence, 'region': region,
        'qte': qte, 'weight_kg': 0, 'kg': kg, 'tonnes': tonnes,
        'sacs_50': sacs_50,
        'montant_ttc': 0, 'montant_ht': montant_ht,
        'source': '2021_2024'
    })

print(f"  Records loaded: {len(records_21_24)}")

# === Load LY_24 (Jul-Dec 2024) ===
print("\nLoading LY_24 (Jul-Dec 2024)...")
ws = wb['LY_24']
rows_24 = list(ws.iter_rows(values_only=True))
print(f"  {len(rows_24)-1} rows")

# LY_24 columns: 0=ref, 1=desc, 2=qte, 5=date, 7=montantHT, 10=etat, 12=tonnes, 13=cat, 14=agence, 19=region
records_24 = []
for r in rows_24[1:]:
    if not r or not r[0]: continue
    ref = str(r[0])
    if ref in EXCLUDED_REFS: continue
    cat = str(r[13]).strip() if r[13] else ''
    family = get_family_from_cat(cat)
    if family is None: continue
    etat = str(r[10]).strip() if r[10] else ''
    if etat != 'Livrée': continue
    agence = str(r[14]).strip() if r[14] else ''
    if agence == '#N/A' or not agence: continue
    
    # Map agence - LY_24 uses same format as Agence2 (Famla, Ndobo, etc.)
    agence_title = agence.capitalize() if agence.isupper() else agence
    # Find in map
    found = False
    for key, (a, reg) in AGENCE_MAP_21_24.items():
        if a.lower() == agence.lower():
            agence_short = a
            region_default = reg
            found = True
            break
    if not found: continue
    
    # Region from col 19
    region_raw = str(r[19]).strip() if r[19] else region_default
    region = map_region(region_raw)
    
    tonnes = r[12] if r[12] else 0  # qteCmdTonne (col M, index 12)
    if family == 'MATERIEL_ELEVAGE':
        tonnes = 0  # Non-tonnage
    
    montant_ht = r[7] if r[7] else 0  # Montant HT (col 7)
    date_str = str(r[5])[:10] if r[5] else ''
    try:
        date = pd.to_datetime(date_str, format='%d/%m/%Y')
    except:
        continue
    if pd.isna(date): continue
    
    qte = r[2] if r[2] else 0
    kg = tonnes * 1000
    sacs_50 = kg / 50 if family != 'MATERIEL_ELEVAGE' else 0
    
    records_24.append({
        'date': date, 'year': date.year, 'month': date.month,
        'ref': ref, 'family': family, 'description': r[1],
        'agence': agence_short, 'region': region,
        'qte': qte, 'weight_kg': 0, 'kg': kg, 'tonnes': tonnes,
        'sacs_50': sacs_50,
        'montant_ttc': 0, 'montant_ht': montant_ht,
        'source': 'S2_2024'
    })

print(f"  Records loaded: {len(records_24)}")

# === Load existing dataset_v2 (2025-2026) ===
print("\nLoading existing dataset_v2 (2025-2026)...")
df_existing = pd.read_csv("/home/z/my-project/scripts/dataset_consolide_v2.csv", parse_dates=['date'], low_memory=False)
print(f"  {len(df_existing)} records")

# === Merge all ===
print("\nMerging all data...")
df_21_24 = pd.DataFrame(records_21_24)
df_24 = pd.DataFrame(records_24)
df_all = pd.concat([df_21_24, df_24, df_existing], ignore_index=True)
print(f"Total records: {len(df_all)}")
print(f"Date range: {df_all['date'].min().date()} → {df_all['date'].max().date()}")
print(f"\nBy year:")
print(df_all.groupby(df_all['date'].dt.year).agg(records=('tonnes','count'), tonnes=('tonnes','sum')).round({'tonnes':0}))
print(f"\nBy family:")
print(df_all['family'].value_counts())

# Save
output_path = "/home/z/my-project/scripts/dataset_2021_2026.csv"
df_all.to_csv(output_path, index=False)
print(f"\nSaved: {output_path}")
print(f"Size: {os.path.getsize(output_path)/1024:.0f} KB")
