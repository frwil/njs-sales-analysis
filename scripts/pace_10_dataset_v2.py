"""
Reconstruction du dataset avec intégration du MATERIEL_ELEVAGE (incluant alvéoles).
Exclusion du MAIS (produit opportuniste).

Familles finales:
  - TOURTEAUX (soja)
  - CONCENTRES (BELGO Chair/Ponte/Porc)
  - INGREDIENTS (Belgotox, Bicarbonate, etc.)
  - ALIMENT_COMPLET (Chick/Piglet Booster)
  - MATERIEL_ELEVAGE (alvéoles, abreuvoirs, mangeoires, cages, radiants, etc.) - CA ONLY

Exclus:
  - MAIS (M1051, M1052) - produit opportuniste hors portefeuille
  - DIVERS (manuels, pierre à lécher)
  - DIVERS2 (contribution carburant, pont bascule, sac réemploi)
  - COMPLEMENT ALIMENTAIRE (BELGOKILL, BELGO HARMONY, etc.)
  - PREMIX (P102N2, P104N2, P109) - à valider
"""
import openpyxl
import pandas as pd
import json
import os

FILE_2025 = "/home/z/my-project/upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx"
FILE_S1_2026 = "/home/z/my-project/upload/ventes janv a juin 2026.xlsx"
FILE_JUIL = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (9).xlsx"
FILE_AOUT = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (24).xlsx"

# === Product refs and weights ===
SOJA_REFS = {'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25}
CONC_REFS = {
    'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
    'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50,
    'C1022': 5,  # BELGO 10% PONTE 5Kg
}
INGREDIENT_REFS = {
    'B100': 25, 'E101': 25, 'I105': 25, 'B1001': 1, 'B1003': 5, 'B1004': 25,
    'E1011': 1, 'E1013': 5, 'E1014': 0.2,  # 200g
    'I1051': 1, 'I1053': 5, 'I1054': 25,
    'I106': 25, 'I1061': 1, 'I107': 25, 'I1071': 1,  # Methionine, Lysine
    'P105': 25, 'P1051': 1, 'P1053': 5,  # Belgofos
    'F114': 50, 'F1145': 50, 'F1146': 25, 'F1147': 1,  # Farine poisson
}
ALIMENT_REFS = {'CB100': 25, 'CB200': 25, 'CB101': 5, 'CB201': 5, 
                'PB100': 25, 'PB200': 25, 'DB100': 25, 'DB200': 25,
                'ALAP25': 25}  # BELGO RABBIT
# MATERIEL_ELEVAGE refs (from 2025 file - catégorie MATERIEL ELEVAGE)
# Inclut alvéoles (MAT014, MAT011, MAT015, MAT017) + abreuvoirs, mangeoires, etc.
MATERIEL_REFS = {
    'MAT003': 1, 'MAT004': 1, 'MAT005': 1, 'MAT006': 1, 'MAT007': 1, 'MAT008': 1, 'MAT009': 1,
    'MAT011': 1, 'MAT014': 1, 'MAT015': 1, 'MAT017': 1,  # ALVEOLES
    'MAT020': 1, 'MAT033': 1, 'MAT039': 1, 'MAT040': 1, 'MAT042': 1,
    'MAT047': 1, 'MAT049': 1, 'MAT050': 1, 'MAT054': 1, 'MAT055': 1, 'MAT073': 1,
    # 2025 file has MAT014-80010003, MAT011-80010002 (with suffix)
    'MAT014-80010003': 1, 'MAT011-80010002': 1,
}
# PREMIX refs (à intégrer car faisant partie du portefeuille)
PREMIX_REFS = {'P102N2': 25, 'P104N2': 25, 'P109': 25, 'PX101': 25, 'PX102': 25, 'PX103': 25, 'PX104': 25, 'PX105': 25}

# Products to EXCLUDE (opportunistic, not in regular portfolio)
EXCLUDED_REFS = {
    'M1051', 'M1052',  # MAIS - opportunistic
    # ELVOR TONIC, CARBONATE DE CALCIUM - non trouvés dans données mais à exclure si jamais vendus
    # Note: These will be flagged in documentation
}

# Internal clients to exclude
INTERNAL_CLIENT_PATTERNS = ['SPC', 'PDC', 'COMPTOIR', 'EMANA']

AGENCE_MAP = {
    'AGENCE FAMLA': ('Famla', 'Ouest'), 'Famla': ('Famla', 'Ouest'),
    'AGENCE MESSASSI': ('Messassi', 'Centre'), 'Messassi': ('Messassi', 'Centre'),
    'AGENCE BERTOUA': ('Bertoua', 'Centre'), 'Bertoua': ('Bertoua', 'Centre'),
    'AGENCE NDOBO': ('Ndobo', 'Littoral'), 'Ndobo': ('Ndobo', 'Littoral'),
    'AGENCE DJELENG': ('Djeleng', 'Ouest'), 'Djeleng': ('Djeleng', 'Ouest'),
    'AGENCE VILLAGE': ('Village', 'Littoral'), 'Village': ('Village', 'Littoral'),
    'AGENCE PK11': ('Pk11', 'Littoral'), 'Pk11': ('Pk11', 'Littoral'),
    'AGENCE NGAOUNDERE': ('Ngaoundere', 'Centre'), 'Ngaoundere': ('Ngaoundere', 'Centre'),
    'AGENCE AHALA': ('Ahala', 'Centre'), 'Ahala': ('Ahala', 'Centre'),
    'AGENCE NKONGSAMBA': ('Nkongsamba', 'Littoral'), 'Nkongsamba': ('Nkongsamba', 'Littoral'),
    'AGENCE NKOLBISSON': ('Nkolbisson', 'Centre'), 'Nkolbisson': ('Nkolbisson', 'Centre'),
    'AGENCE NKOABANG': ('Nkoabang', 'Centre'), 'Nkoabang': ('Nkoabang', 'Centre'),
    'AGENCE DE BAMENDA - DEPOT MBOUDA': ('Mbouda', 'Ouest'), 'Mbouda': ('Mbouda', 'Ouest'),
    'AGENCE BUEA': ('Buea', 'Littoral'), 'Buea': ('Buea', 'Littoral'),
}


def get_family_v2(ref):
    """Map ref to product family (v2 - with MATERIEL_ELEVAGE)."""
    if ref in EXCLUDED_REFS: return 'EXCLUDED'
    if ref in SOJA_REFS: return 'TOURTEAUX'
    if ref in CONC_REFS: return 'CONCENTRES'
    if ref in INGREDIENT_REFS: return 'INGREDIENTS'
    if ref in ALIMENT_REFS: return 'ALIMENT_COMPLET'
    if ref in MATERIEL_REFS: return 'MATERIEL_ELEVAGE'
    if ref in PREMIX_REFS: return 'PREMIX'
    # Fallback: check by ref prefix
    if ref.startswith('MAT'): return 'MATERIEL_ELEVAGE'
    if ref.startswith('M'): return 'EXCLUDED'  # MAIS et autres
    return 'AUTRES'


def get_weight_v2(ref):
    if ref in SOJA_REFS: return SOJA_REFS[ref]
    if ref in CONC_REFS: return CONC_REFS[ref]
    if ref in INGREDIENT_REFS: return INGREDIENT_REFS[ref]
    if ref in ALIMENT_REFS: return ALIMENT_REFS[ref]
    if ref in MATERIEL_REFS: return MATERIEL_REFS[ref]
    if ref in PREMIX_REFS: return PREMIX_REFS[ref]
    return 1  # Default for material (1 unit = 1 piece)


def is_internal_client(client_str):
    if not client_str: return False
    s = str(client_str).upper()
    for p in INTERNAL_CLIENT_PATTERNS:
        if p in s: return True
    return False


def detect_cols_2025(ws):
    header = None
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        header = row
    indices = {}
    for i, h in enumerate(header):
        if h == 'Réf. produit': indices['ref'] = i
        elif h == 'Description du produit': indices['desc'] = i
        elif h == 'Qté commandée': indices['qte'] = i
        elif h == 'Tiers': indices['client'] = i
        elif h == 'Date de commande': indices['date'] = i
        elif h == 'Montant HT': indices['montant_ht'] = i
        elif h == 'Montant TTC': indices['montant_ttc'] = i
        elif h == 'État': indices['etat'] = i
        elif h == 'tableauProprieteAgences.Agence': indices['agence'] = i
        elif h == 'tableauProprieteAgences.Region': indices['region'] = i
    return indices


def detect_cols_2026(ws, header_row=2):
    header = None
    for row in ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True):
        header = row
    indices = {}
    for i, h in enumerate(header):
        if h == 'Réf. produit': indices['ref'] = i
        elif h == 'Description du produit': indices['desc'] = i
        elif h == 'Qté commandée': indices['qte'] = i
        elif h == 'Tiers': indices['client'] = i
        elif h == 'Date de commande': indices['date'] = i
        elif h == 'Montant HT': indices['montant_ht'] = i
        elif h == 'Montant TTC': indices['montant_ttc'] = i
        elif h == 'État': indices['etat'] = i
        elif h == 'agence': indices['agence'] = i
    return indices


def parse_date(d):
    if isinstance(d, pd.Timestamp): return d
    if isinstance(d, str):
        try: return pd.to_datetime(d, format='%d/%m/%Y')
        except: 
            try: return pd.to_datetime(d)
            except: return None
    return pd.to_datetime(d, errors='coerce')


def process_rows(ws, cols, source_label, date_filter=None):
    records = []
    for r in ws.iter_rows(min_row=(3 if source_label != '2025' else 2), values_only=True):
        if not r or len(r) <= max(cols.values()): continue
        if r[cols['ref']] == 'Total': continue
        etat = str(r[cols['etat']]).strip() if r[cols['etat']] else ''
        if etat != 'Livrée': continue
        ref = str(r[cols['ref']])
        family = get_family_v2(ref)
        if family in ('EXCLUDED', 'AUTRES'): continue
        client = r[cols['client']]
        if is_internal_client(client): continue
        agence_raw = r[cols['agence']]
        agence_short, region = AGENCE_MAP.get(agence_raw, (None, None))
        if not agence_short: continue
        qte = r[cols['qte']] or 0
        weight = get_weight_v2(ref)
        kg = qte * weight
        montant_ttc = r[cols['montant_ttc']] or 0
        montant_ht = r[cols['montant_ht']] if 'montant_ht' in cols and r[cols['montant_ht']] else 0
        date = parse_date(r[cols['date']])
        if date is None: continue
        if date_filter and date_filter not in str(date): continue
        records.append({
            'date': date, 'year': date.year, 'month': date.month,
            'ref': ref, 'family': family, 'description': r[cols['desc']],
            'client': client, 'agence': agence_short, 'region': region,
            'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': kg / 1000,
            'sacs_50': kg / 50 if family != 'MATERIEL_ELEVAGE' else 0,  # Pas de sacs pour matériel
            'montant_ttc': montant_ttc, 'montant_ht': montant_ht,
            'source': source_label
        })
    return records


# === Load all sources ===
print("Loading 2025 historical data...")
wb = openpyxl.load_workbook(FILE_2025, read_only=True, data_only=True)
ws = wb['Feuil1']
cols = detect_cols_2025(ws)
records_2025 = process_rows(ws, cols, '2025')
print(f"  2025: {len(records_2025)} records")

print("\nLoading S1 2026...")
wb = openpyxl.load_workbook(FILE_S1_2026, read_only=True, data_only=True)
records_s1 = []
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    cols = detect_cols_2026(ws, header_row=2)
    if 'ref' not in cols: continue
    records_s1.extend(process_rows(ws, cols, 'S1_2026'))
print(f"  S1 2026: {len(records_s1)} records")

print("\nLoading Juillet 2026...")
wb = openpyxl.load_workbook(FILE_JUIL, read_only=True, data_only=True)
ws = wb['Sheet 1']
cols = detect_cols_2026(ws, header_row=2)
records_juil = process_rows(ws, cols, 'Juil_2026', date_filter='2026-07')
print(f"  Juillet 2026: {len(records_juil)} records")

print("\nLoading Août 2026...")
wb = openpyxl.load_workbook(FILE_AOUT, read_only=True, data_only=True)
ws = wb['Sheet 1']
cols = detect_cols_2026(ws, header_row=2)
records_aout = []
for r in ws.iter_rows(min_row=3, values_only=True):
    if not r or len(r) <= max(cols.values()): continue
    if r[cols['ref']] == 'Total': continue
    etat = str(r[cols['etat']]).strip() if r[cols['etat']] else ''
    if etat != 'Livrée': continue
    if not r[cols['date']] or '/08/2026' not in str(r[cols['date']]): continue
    if str(r[cols['date']]) == '27/08/2026': continue
    ref = str(r[cols['ref']])
    family = get_family_v2(ref)
    if family in ('EXCLUDED', 'AUTRES'): continue
    client = r[cols['client']]
    if is_internal_client(client): continue
    agence_raw = r[cols['agence']]
    agence_short, region = AGENCE_MAP.get(agence_raw, (None, None))
    if not agence_short: continue
    qte = r[cols['qte']] or 0
    weight = get_weight_v2(ref)
    kg = qte * weight
    montant_ttc = r[cols['montant_ttc']] or 0
    montant_ht = r[cols['montant_ht']] or 0
    date = parse_date(r[cols['date']])
    if date is None: continue
    records_aout.append({
        'date': date, 'year': date.year, 'month': date.month,
        'ref': ref, 'family': family, 'description': r[cols['desc']],
        'client': client, 'agence': agence_short, 'region': region,
        'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': kg / 1000,
        'sacs_50': kg / 50 if family != 'MATERIEL_ELEVAGE' else 0,
        'montant_ttc': montant_ttc, 'montant_ht': montant_ht,
        'source': 'Aout_2026'
    })
print(f"  Août 2026: {len(records_aout)} records")

# === Consolidate ===
print("\nConsolidating...")
all_records = records_2025 + records_s1 + records_juil + records_aout
df = pd.DataFrame(all_records)
print(f"Total records: {len(df)}")
print(f"Date range: {df['date'].min()} → {df['date'].max()}")
print(f"\nRecords by family:")
print(df['family'].value_counts())

# Save
output_path = "/home/z/my-project/scripts/dataset_consolide_v2.csv"
df.to_csv(output_path, index=False)
print(f"\nDataset saved: {output_path}")
print(f"File size: {os.path.getsize(output_path) / 1024:.0f} KB")

# Stats by family
print("\n=== STATS PAR FAMILLE ===")
stats = df.groupby('family').agg(
    tonnes=('tonnes', 'sum'),
    ca_m_fcfa=('montant_ttc', lambda x: x.sum() / 1e6),
    n_records=('tonnes', 'count'),
    n_produits=('ref', 'nunique'),
).round({'tonnes': 1, 'ca_m_fcfa': 1})
print(stats)
