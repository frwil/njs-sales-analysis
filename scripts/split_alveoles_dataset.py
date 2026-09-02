"""
Construction du dataset 2023-2026 v2 - avec ALVEOLES comme famille séparée.

CHANGEMENTS vs v1:
1. ALVEOLES extraite de MATERIEL_ELEVAGE comme 8e famille distincte
   - Refs: MAT011-80010002, MAT014-80010003, MAT015, MAT017 (4 refs)
   - Identification: description contient "ALVEOLE"
2. MATERIEL_ELEVAGE maintenant = 32 refs (hors alvéoles) — toujours tonnes=0
3. ALVEOLES: tonnes=0 comme MATERIEL_ELEVAGE (CA only, non exprimable en volume)

8 familles au total (vs 7 avant):
  TOURTEAUX, CONCENTRES, INGREDIENTS, ALIMENT_COMPLET, MATERIEL_ELEVAGE, PREMIX,
  COMPLEMENT_ALIMENTAIRE, ALVEOLES (NOUVEAU)
"""
import pandas as pd
import os

# === ALVEOLES refs (identified by description containing "ALVEOLE") ===
ALVEOLES_REFS = {'MAT011-80010002', 'MAT014-80010003', 'MAT015', 'MAT017'}

# === Load existing dataset 2023-2026 ===
print("=" * 70)
print("Loading existing dataset 2023-2026...")
print("=" * 70)
df = pd.read_csv("/home/z/my-project/scripts/dataset_2023_2026.csv", parse_dates=['date'], low_memory=False)
print(f"  {len(df)} records loaded")
print(f"  Date range: {df['date'].min().date()} → {df['date'].max().date()}")
print(f"\n  Before split:")
print(df['family'].value_counts())

# === Identify ALVEOLES by ref OR by description containing "ALVEOLE" ===
def is_alveole(ref, desc):
    if ref in ALVEOLES_REFS: return True
    if pd.notna(desc) and 'ALVEOLE' in str(desc).upper(): return True
    return False

# Apply split
mask_alv = df.apply(lambda r: r['family'] == 'MATERIEL_ELEVAGE' and is_alveole(r['ref'], r.get('description', '')), axis=1)
print(f"\n  ALVEOLES records to extract: {mask_alv.sum()}")

# Move ALVEOLES records to new family
df.loc[mask_alv, 'family'] = 'ALVEOLES'

# Verify ALVEOLES is still tonnes=0 (it's a physical product like cages, not bulk)
df.loc[df['family'] == 'ALVEOLES', 'tonnes'] = 0
df.loc[df['family'] == 'ALVEOLES', 'sacs_50'] = 0

print(f"\n  After split:")
print(df['family'].value_counts())

# Stats by family
print("\n" + "=" * 70)
print("STATS PAR FAMILLE (2023-2026) — AVEC ALVEOLES SÉPARÉS")
print("=" * 70)
stats = df.groupby('family').agg(
    tonnes=('tonnes', 'sum'),
    ca_ht_m_fcfa=('montant_ht', lambda x: x.sum() / 1e6),
    ca_ttc_m_fcfa=('montant_ttc', lambda x: x.sum() / 1e6),
    n_records=('tonnes', 'count'),
    n_produits=('ref', 'nunique'),
    n_agences=('agence', 'nunique'),
).round({'tonnes': 1, 'ca_ht_m_fcfa': 1, 'ca_ttc_m_fcfa': 1})
print(stats)

# Stats ALVEOLES vs MATERIEL_ELEVAGE by year
print("\n" + "=" * 70)
print("ALVEOLES vs MATERIEL_ELEVAGE par année")
print("=" * 70)
df['year'] = df['date'].dt.year
for fam in ['ALVEOLES', 'MATERIEL_ELEVAGE']:
    sub = df[df['family'] == fam]
    if len(sub) > 0:
        print(f"\n{fam}:")
        by_year = sub.groupby('year').agg(
            records=('tonnes', 'count'),
            ca_ht_m=('montant_ht', lambda x: x.sum() / 1e6),
            ca_ttc_m=('montant_ttc', lambda x: x.sum() / 1e6),
        ).round({'ca_ht_m': 1, 'ca_ttc_m': 1})
        print(by_year)

# ALVEOLES refs verification
print("\n" + "=" * 70)
print("ALVEOLES — Références")
print("=" * 70)
alv = df[df['family'] == 'ALVEOLES']
print(f"Nb refs: {alv['ref'].nunique()}")
print(f"Refs: {sorted(alv['ref'].unique().tolist())}")
print(f"Descriptions: {alv.groupby('ref')['description'].first().to_dict()}")

# MATERIEL_ELEVAGE refs verification
print("\n" + "=" * 70)
print("MATERIEL_ELEVAGE — Références (hors alvéoles)")
print("=" * 70)
mat = df[df['family'] == 'MATERIEL_ELEVAGE']
print(f"Nb refs: {mat['ref'].nunique()}")
print(f"Refs: {sorted(mat['ref'].unique().tolist())}")

# Save
output_path = "/home/z/my-project/scripts/dataset_2023_2026.csv"  # Overwrite same file
df = df.drop(columns=['year'])  # Don't keep year column
df.to_csv(output_path, index=False)
print(f"\nDataset saved: {output_path}")
print(f"File size: {os.path.getsize(output_path) / 1024:.0f} KB")

print("\n" + "=" * 70)
print(f"DATASET 2023-2026 v2 PRÊT: {len(df):,} records, {df['ref'].nunique()} produits, {df['family'].nunique()} familles")
print(f"8 familles: TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, ALIMENT_COMPLET, MATERIEL_ELEVAGE, PREMIX, COMPLEMENT_ALIMENTAIRE, ALVEOLES")
print("=" * 70)
