"""
Phase Prepare - Construction du dataset consolide pour forecast Q4 2026.
Sources:
  - Historique 2025 (86d96135...): Jan-Dec 2025
  - S1 2026 (ventes janv a juin 2026.xlsx): Jan-Juin 2026
  - Juillet 2026 (NJS GROUP ERP (9).xlsx)
  - Aout 2026 (NJS GROUP ERP (21).xlsx): 01-26/08

Output: dataset_consolidé.csv (1 row par commande produit × date × agence)
"""
import openpyxl
import pandas as pd
from collections import defaultdict
import json
import os

# === Source files ===
FILE_2025 = "/home/z/my-project/upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx"
FILE_S1_2026 = "/home/z/my-project/upload/ventes janv a juin 2026.xlsx"
FILE_JUIL = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (9).xlsx"
FILE_AOUT = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (21).xlsx"

# Product refs and weights
SOJA_REFS = {'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25}
CONC_REFS = {
    'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
    'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50,
}
MAIS_REFS = {'M1051': 50, 'M1052': 50}
INGREDIENT_REFS = {
    'B100': 25, 'E101': 25, 'I105': 25, 'B1001': 1, 'B1003': 5, 'B1004': 25,
    'E1011': 1, 'E1013': 5, 'E1014': 25, 'I1051': 1, 'I1053': 5, 'I1054': 25,
}
PREMIX_REFS = {'PX101': 25, 'PX102': 25, 'PX103': 25, 'PX104': 25, 'PX105': 25}
ALIMENT_REFS = {'CB100': 25, 'CB200': 25, 'PB100': 25, 'PB200': 25, 'DB100': 25, 'DB200': 25}

AGENCE_MAP = {
    'AGENCE FAMLA': ('Famla', 'Ouest'),
    'AGENCE MESSASSI': ('Messassi', 'Centre'),
    'AGENCE BERTOUA': ('Bertoua', 'Centre'),
    'AGENCE NDOBO': ('Ndobo', 'Littoral'),
    'AGENCE DJELENG': ('Djeleng', 'Ouest'),
    'AGENCE VILLAGE': ('Village', 'Littoral'),
    'AGENCE PK11': ('Pk11', 'Littoral'),
    'AGENCE NGAOUNDERE': ('Ngaoundere', 'Centre'),
    'AGENCE AHALA': ('Ahala', 'Centre'),
    'AGENCE NKONGSAMBA': ('Nkongsamba', 'Littoral'),
    'AGENCE NKOLBISSON': ('Nkolbisson', 'Centre'),
    'AGENCE NKOABANG': ('Nkoabang', 'Centre'),
    'AGENCE DE BAMENDA - DEPOT MBOUDA': ('Mbouda', 'Ouest'),
    'AGENCE BUEA': ('Buea', 'Littoral'),
    'SPC BAF-CHEFFERIE': ('Baf-Chefferie', 'Ouest'),
    'PDC Emana': ('Emana', 'Centre'),
    'SPC-NDERE': ('Ndere', 'Centre'),
    'SPC-DSCHANG': ('Dschang', 'Ouest'),
    'SPC BUEA': ('Buea-SPC', 'Littoral'),
    'SPC-YASSA': ('Yassa', 'Littoral'),
    # 2025 format (short names)
    'Famla': ('Famla', 'Ouest'),
    'Messassi': ('Messassi', 'Centre'),
    'Bertoua': ('Bertoua', 'Centre'),
    'Ndobo': ('Ndobo', 'Littoral'),
    'Djeleng': ('Djeleng', 'Ouest'),
    'Village': ('Village', 'Littoral'),
    'Pk11': ('Pk11', 'Littoral'),
    'Ngaoundere': ('Ngaoundere', 'Centre'),
    'Ahala': ('Ahala', 'Centre'),
    'Nkongsamba': ('Nkongsamba', 'Littoral'),
    'Nkolbisson': ('Nkolbisson', 'Centre'),
    'Nkoabang': ('Nkoabang', 'Centre'),
    'Mbouda': ('Mbouda', 'Ouest'),
    'Buea': ('Buea', 'Littoral'),
}

INTERNAL_CLIENT_PATTERNS = ['SPC', 'PDC', 'COMPTOIR', 'EMANA']


def get_family(ref):
    if ref in SOJA_REFS: return 'TOURTEAUX'
    if ref in CONC_REFS: return 'CONCENTRES'
    if ref in MAIS_REFS: return 'MAIS'
    if ref in INGREDIENT_REFS: return 'INGREDIENTS'
    if ref in PREMIX_REFS: return 'PREMIX'
    if ref in ALIMENT_REFS: return 'ALIMENT_COMPLET'
    return 'AUTRES'


def get_weight(ref):
    if ref in SOJA_REFS: return SOJA_REFS[ref]
    if ref in CONC_REFS: return CONC_REFS[ref]
    if ref in MAIS_REFS: return MAIS_REFS[ref]
    if ref in INGREDIENT_REFS: return INGREDIENT_REFS[ref]
    if ref in PREMIX_REFS: return PREMIX_REFS[ref]
    if ref in ALIMENT_REFS: return ALIMENT_REFS[ref]
    return 50  # default


def is_internal_client(client_str):
    if not client_str: return False
    s = str(client_str).upper()
    for p in INTERNAL_CLIENT_PATTERNS:
        if p in s: return True
    return False


def detect_cols_2025(ws):
    """2025 file has different header layout."""
    header = None
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        header = row
        break
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
        elif h == 'tableauPropietesProduits.CategorieProduit': indices['cat'] = i
        elif h == 'tableauProprieteAgences.Agence': indices['agence'] = i
        elif h == 'tableauProprieteAgences.Region': indices['region'] = i
    return indices


def detect_cols_2026(ws, header_row=2):
    """2026 file: title on row 1, header on row 2."""
    header = None
    for row in ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True):
        header = row
        break
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
        try:
            # Try DD/MM/YYYY
            return pd.to_datetime(d, format='%d/%m/%Y')
        except Exception:
            try:
                return pd.to_datetime(d)
            except Exception:
                return None
    return pd.to_datetime(d, errors='coerce')


# === Load 2025 data ===
print("Loading 2025 historical data...")
wb = openpyxl.load_workbook(FILE_2025, read_only=True, data_only=True)
ws = wb['Feuil1']
cols = detect_cols_2025(ws)
print(f"  2025 columns: {cols}")

records_2025 = []
for r in ws.iter_rows(min_row=2, values_only=True):
    if not r or len(r) <= max(cols.values()): continue
    if r[cols['ref']] == 'Total': continue
    etat = str(r[cols['etat']]).strip() if r[cols['etat']] else ''
    if etat != 'Livrée': continue
    ref = r[cols['ref']]
    family = get_family(ref)
    if family == 'AUTRES': continue  # Skip DIVERS, MATERIEL ELEVAGE, etc.
    client = r[cols['client']]
    if is_internal_client(client): continue
    agence_raw = r[cols['agence']]
    agence_short, region = AGENCE_MAP.get(agence_raw, (None, None))
    if not agence_short: continue  # Skip unknown agencies
    qte = r[cols['qte']] or 0
    weight = get_weight(ref)
    kg = qte * weight
    montant_ttc = r[cols['montant_ttc']] or 0
    montant_ht = r[cols['montant_ht']] or 0
    date = parse_date(r[cols['date']])
    if date is None: continue
    records_2025.append({
        'date': date, 'year': date.year, 'month': date.month,
        'ref': ref, 'family': family, 'description': r[cols['desc']],
        'client': client, 'agence': agence_short, 'region': region,
        'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': kg / 1000,
        'sacs_50': kg / 50, 'montant_ttc': montant_ttc, 'montant_ht': montant_ht,
        'source': '2025'
    })
print(f"  2025: {len(records_2025)} records loaded")

# === Load S1 2026 (multi-sheet) ===
print("\nLoading S1 2026...")
wb = openpyxl.load_workbook(FILE_S1_2026, read_only=True, data_only=True)
records_s1 = []
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    cols = detect_cols_2026(ws, header_row=2)
    if 'ref' not in cols: continue
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) <= max(cols.values()): continue
        if r[cols['ref']] == 'Total': continue
        etat = str(r[cols['etat']]).strip() if r[cols['etat']] else ''
        if etat != 'Livrée': continue
        ref = r[cols['ref']]
        family = get_family(ref)
        if family == 'AUTRES': continue
        client = r[cols['client']]
        if is_internal_client(client): continue
        agence_raw = r[cols['agence']]
        agence_short, region = AGENCE_MAP.get(agence_raw, (None, None))
        if not agence_short: continue
        qte = r[cols['qte']] or 0
        weight = get_weight(ref)
        kg = qte * weight
        montant_ttc = r[cols['montant_ttc']] if 'montant_ttc' in cols and r[cols['montant_ttc']] else 0
        montant_ht = r[cols['montant_ht']] if 'montant_ht' in cols and r[cols['montant_ht']] else 0
        date = parse_date(r[cols['date']])
        if date is None: continue
        records_s1.append({
            'date': date, 'year': date.year, 'month': date.month,
            'ref': ref, 'family': family, 'description': r[cols['desc']],
            'client': client, 'agence': agence_short, 'region': region,
            'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': kg / 1000,
            'sacs_50': kg / 50, 'montant_ttc': montant_ttc, 'montant_ht': montant_ht,
            'source': 'S1_2026'
        })
print(f"  S1 2026: {len(records_s1)} records loaded")

# === Load Juillet 2026 ===
print("\nLoading Juillet 2026...")
wb = openpyxl.load_workbook(FILE_JUIL, read_only=True, data_only=True)
ws = wb['Sheet 1']
cols = detect_cols_2026(ws, header_row=2)
records_juil = []
for r in ws.iter_rows(min_row=3, values_only=True):
    if not r or len(r) <= max(cols.values()): continue
    if r[cols['ref']] == 'Total': continue
    etat = str(r[cols['etat']]).strip() if r[cols['etat']] else ''
    if etat != 'Livrée': continue
    if not r[cols['date']] or '/07/2026' not in str(r[cols['date']]): continue
    ref = r[cols['ref']]
    family = get_family(ref)
    if family == 'AUTRES': continue
    client = r[cols['client']]
    if is_internal_client(client): continue
    agence_raw = r[cols['agence']]
    agence_short, region = AGENCE_MAP.get(agence_raw, (None, None))
    if not agence_short: continue
    qte = r[cols['qte']] or 0
    weight = get_weight(ref)
    kg = qte * weight
    montant_ttc = r[cols['montant_ttc']] or 0
    montant_ht = r[cols['montant_ht']] or 0
    date = parse_date(r[cols['date']])
    if date is None: continue
    records_juil.append({
        'date': date, 'year': date.year, 'month': date.month,
        'ref': ref, 'family': family, 'description': r[cols['desc']],
        'client': client, 'agence': agence_short, 'region': region,
        'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': kg / 1000,
        'sacs_50': kg / 50, 'montant_ttc': montant_ttc, 'montant_ht': montant_ht,
        'source': 'Juil_2026'
    })
print(f"  Juillet 2026: {len(records_juil)} records loaded")

# === Load Août 2026 ===
print("\nLoading Août 2026 (au 26/08)...")
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
    if str(r[cols['date']]) == '27/08/2026': continue  # skip if matinal
    ref = r[cols['ref']]
    family = get_family(ref)
    if family == 'AUTRES': continue
    client = r[cols['client']]
    if is_internal_client(client): continue
    agence_raw = r[cols['agence']]
    agence_short, region = AGENCE_MAP.get(agence_raw, (None, None))
    if not agence_short: continue
    qte = r[cols['qte']] or 0
    weight = get_weight(ref)
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
        'sacs_50': kg / 50, 'montant_ttc': montant_ttc, 'montant_ht': montant_ht,
        'source': 'Aout_2026'
    })
print(f"  Août 2026: {len(records_aout)} records loaded")

# === Consolidate ===
print("\nConsolidating all records...")
all_records = records_2025 + records_s1 + records_juil + records_aout
df = pd.DataFrame(all_records)
print(f"Total records: {len(df)}")
print(f"Date range: {df['date'].min()} → {df['date'].max()}")
print(f"\nRecords by source:")
print(df['source'].value_counts())
print(f"\nRecords by family:")
print(df['family'].value_counts())
print(f"\nRecords by year-month (last 12):")
df['year_month'] = df['date'].dt.to_period('M')
print(df['year_month'].value_counts().sort_index().tail(12))

# === Save ===
output_path = "/home/z/my-project/scripts/dataset_consolide.csv"
df.to_csv(output_path, index=False)
print(f"\nDataset saved: {output_path}")
print(f"File size: {os.path.getsize(output_path) / 1024:.0f} KB")

# Summary stats
summary = {
    'total_records': len(df),
    'date_min': str(df['date'].min()),
    'date_max': str(df['date'].max()),
    'by_source': df['source'].value_counts().to_dict(),
    'by_family': df['family'].value_counts().to_dict(),
    'n_products': df['ref'].nunique(),
    'n_agences': df['agence'].nunique(),
    'n_regions': df['region'].nunique(),
    'n_clients': df['client'].nunique(),
}
with open("/home/z/my-project/scripts/dataset_summary.json", 'w') as f:
    json.dump(summary, f, indent=2, default=str)
print(f"\nSummary saved: /home/z/my-project/scripts/dataset_summary.json")
