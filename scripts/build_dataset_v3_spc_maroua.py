"""
Mise à jour du dataset 2023-2026 v3 - Intégration des agences SPC + Maroua.

CHANGEMENTS vs v2:
1. Intégration des 11 agences SPC (auparavant exclues car non dans AGENCE_MAP):
   - SPC BAF-CHEFFERIE (Ouest)
   - SPC BUEA, SPC VILLAGE/SPC Village, SPC-DLA-BERI, SPC-YASSA, SPC PK15, SPC-TPO (Littoral)
   - SPC KYE-OSSI, SPC-NDERE (Centre)
   - SPC-DSCHANG (Ouest)
2. Ajout de l'agence MAROUA (Nord → Centre par convention)
3. Maintien de l'exclusion SPC comme CLIENT (Tiers column) - c'est différent

Résultat: 25 agences (14 BELGOCAM + 11 SPC) + Maroua = 26 agences total
Régions: Ouest, Centre, Littoral (Maroua → Centre par convention, comme Yde/Est/Nord)

Note: Les alvéoles de SPC BAF-CHEFFERIE ont explosé en 2025 (311 M)
mais quasi-disparu en 2026. Pour la saisonnalité ALVEOLES, voir forecast.
"""
import openpyxl
import pandas as pd
import os
import re

# === 11 SPC agencies mapping ===
SPC_AGENCES = {
    'SPC BAF-CHEFFERIE': ('SPC Baf-Chefferie', 'Ouest'),
    'SPC BUEA': ('SPC Buea', 'Littoral'),
    'SPC KYE-OSSI': ('SPC Kye-Ossi', 'Centre'),
    'SPC VILLAGE': ('SPC Village', 'Littoral'),
    'SPC Village': ('SPC Village', 'Littoral'),
    'SPC-DLA-BERI': ('SPC DLA-Beri', 'Littoral'),
    'SPC-DSCHANG': ('SPC Dschang', 'Ouest'),
    'SPC-NDERE': ('SPC Ndere', 'Centre'),
    'SPC-YASSA': ('SPC Yassa', 'Littoral'),
    'SPC PK15': ('SPC PK15', 'Littoral'),
    'SPC-TPO': ('SPC TPO', 'Littoral'),
}

# Maroua (only seen in août 2026, soja T102 sales)
MAROUA_MAPPING = {
    'AGENCE MAROUA': ('Maroua', 'Centre'),  # Centre convention (Nord → Centre)
}

# Existing BELGOCAM agencies
BELGOCAM_AGENCES = {
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
    # Short forms (for LY_24)
    'FAMLA': ('Famla', 'Ouest'), 'DJELENG': ('Djeleng', 'Ouest'), 'MBOUDA': ('Mbouda', 'Ouest'),
    'MESSASSI': ('Messassi', 'Centre'), 'AHALA': ('Ahala', 'Centre'),
    'BERTOUA': ('Bertoua', 'Centre'), 'NGAOUNDERE': ('Ngaoundere', 'Centre'),
    'NKOLBISSON': ('Nkolbisson', 'Centre'), 'NKOABANG': ('Nkoabang', 'Centre'),
    'NDOBO': ('Ndobo', 'Littoral'), 'VILLAGE': ('Village', 'Littoral'),
    'PK11': ('Pk11', 'Littoral'), 'NKONGSAMBA': ('Nkongsamba', 'Littoral'),
    'BUEA': ('Buea', 'Littoral'),
}

# Combined map
ALL_AGENCES = {**BELGOCAM_AGENCES, **SPC_AGENCES, **MAROUA_MAPPING}

# === Product refs ===
SOJA_REFS = {'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25}
CONC_REFS = {'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
             'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50, 'C1022': 5}
ALIMENT_REFS = {'CB100': 25, 'CB200': 25, 'CB101': 5, 'CB201': 5, 'PB100': 25, 'PB200': 25, 
                'DB100': 25, 'DB200': 25, 'ALAP25': 25}
INGREDIENT_REFS = {'B100': 25, 'E101': 25, 'I105': 25, 'B1001': 1, 'B1003': 5, 'B1004': 25,
                   'E1011': 1, 'E1013': 5, 'E1014': 0.2, 'I1051': 1, 'I1053': 5, 'I1054': 25,
                   'I106': 25, 'I1061': 1, 'I107': 25, 'I1071': 1, 'P105': 25, 'P1051': 1, 'P1053': 5,
                   'F114': 50, 'F1145': 50, 'F1146': 25, 'F1147': 1}
PREMIX_REFS = {'P102N2': 25, 'P104N2': 25, 'P109': 25, 'PX101': 25, 'PX102': 25, 'PX103': 25, 'PX104': 25, 'PX105': 25}
MATERIEL_REFS = {f'MAT{i:03d}': 1 for i in range(1, 100)}
MATERIEL_REFS.update({'MAT014-80010003': 1, 'MAT011-80010002': 1})
MATERIEL_REFS.update({f'ME{i:03d}': 1 for i in range(100, 200)})
MATERIEL_REFS.update({'ME100': 1, 'ME1001': 1, 'ME101': 1, 'ME102': 1, 'ME103': 1, 'ME104': 1, 'ME1041': 1, 'ME105': 1, 'ME106': 1, 'ME107': 1})
COMPLEMENT_REFS = {
    'V300': 1, 'CA003.1': 1, 'CA004.1': 1, 'CA006.1': 1, 'CA001.1': 1,
    'CA002.1': 1, 'CA005.1': 1, 'CA007.1': 1, 'CA008.1': 1,
}
ALVEOLES_REFS = {'MAT011-80010002', 'MAT014-80010003', 'MAT015', 'MAT017'}

# Products to exclude
EXCLUDED_REFS = {'M1051', 'M1052', 'V305'}

# Internal clients to exclude (Tiers column - SPC SA, PDC, COMPTOIR, EMANA)
INTERNAL_CLIENT_PATTERNS = ['SPC SA', 'PDC', 'COMPTOIR', 'EMANA']

def get_family(ref):
    if ref in EXCLUDED_REFS: return 'EXCLUDED'
    if ref in SOJA_REFS: return 'TOURTEAUX'
    if ref in CONC_REFS: return 'CONCENTRES'
    if ref in ALIMENT_REFS: return 'ALIMENT_COMPLET'
    if ref in INGREDIENT_REFS: return 'INGREDIENTS'
    if ref in PREMIX_REFS: return 'PREMIX'
    if ref in ALVEOLES_REFS: return 'ALVEOLES'
    if ref in MATERIEL_REFS or ref.startswith('MAT') or ref.startswith('ME'): return 'MATERIEL_ELEVAGE'
    if ref in COMPLEMENT_REFS: return 'COMPLEMENT_ALIMENTAIRE'
    return 'AUTRES'

def get_weight(ref):
    if ref in SOJA_REFS: return SOJA_REFS[ref]
    if ref in CONC_REFS: return CONC_REFS[ref]
    if ref in ALIMENT_REFS: return ALIMENT_REFS[ref]
    if ref in INGREDIENT_REFS: return INGREDIENT_REFS[ref]
    if ref in PREMIX_REFS: return PREMIX_REFS[ref]
    if ref in ALVEOLES_REFS: return 1
    if ref in MATERIEL_REFS: return 1
    if ref in COMPLEMENT_REFS: return 1
    return 1

def is_internal_client(client_str):
    """Check if client (Tiers) is internal (SPC SA as client, PDC, COMPTOIR, EMANA).
    NOTE: 'SPC' alone in Tiers column may match agency names — we exclude SPC SA specifically."""
    if not client_str: return False
    s = str(client_str).upper()
    # Internal clients are SPC SA (the sister company), PDC, COMPTOIR, EMANA
    for p in INTERNAL_CLIENT_PATTERNS:
        if p in s: return True
    return False

def parse_date(d):
    if isinstance(d, pd.Timestamp): return d
    if isinstance(d, str):
        try: return pd.to_datetime(d, format='%d/%m/%Y')
        except:
            try: return pd.to_datetime(d)
            except: return None
    return pd.to_datetime(d, errors='coerce')

# === Process sources ===
all_records = []

# === LY_24 (Jul-Dec 2024) - includes SPC agencies ===
print("=" * 70)
print("Loading LY_24 (Jul-Dec 2024) - avec SPC agences...")
print("=" * 70)
wb = openpyxl.load_workbook("/home/z/my-project/upload/21_24.xlsx", read_only=True, data_only=True)
ws = wb['LY_24']
rows_24 = list(ws.iter_rows(values_only=True))

# LY_24 has agence in col 14, no Tiers col
for r in rows_24[1:]:
    if not r or not r[0]: continue
    ref = str(r[0])
    if ref in EXCLUDED_REFS: continue
    family = get_family(ref)
    if family in ('EXCLUDED', 'AUTRES'): continue
    etat = str(r[10]).strip() if r[10] else ''
    if etat != 'Livrée': continue
    agence = str(r[14]).strip() if r[14] else ''
    if agence == '#N/A' or not agence: continue
    # Check if agence is in our map (now includes SPC)
    if agence not in ALL_AGENCES: continue
    agence_short, region = ALL_AGENCES[agence]
    
    qte = r[2] if r[2] else 0
    weight = get_weight(ref)
    kg = qte * weight
    sacs_50 = kg / 50 if family not in ('MATERIEL_ELEVAGE', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES') else 0
    if family in ('MATERIEL_ELEVAGE', 'ALVEOLES'):
        tonnes_val = 0  # CA only
    elif family == 'COMPLEMENT_ALIMENTAIRE':
        tonnes_val = kg / 1000
    else:
        tonnes_val = r[12] if r[12] else kg / 1000  # qteCmdTonne
    
    montant_ht = r[7] if r[7] else 0
    date_str = str(r[5])[:10] if r[5] else ''
    try:
        date = pd.to_datetime(date_str, format='%d/%m/%Y')
    except:
        continue
    if pd.isna(date): continue
    
    all_records.append({
        'date': date, 'year': date.year, 'month': date.month,
        'ref': ref, 'family': family, 'description': r[1],
        'client': '', 'agence': agence_short, 'region': region,
        'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': tonnes_val,
        'sacs_50': sacs_50,
        'montant_ttc': 0, 'montant_ht': montant_ht,
        'source': 'LY_24'
    })
print(f"  LY_24 records (with SPC): {len([r for r in all_records if r['source']=='LY_24'])}")

# === LY_21_24 filter 2023 (similar approach, no Tiers column) ===
print("\nLoading LY_21_24 (filter 2023) - avec SPC agences...")
ws2 = wb['LY_21_24']
rows_23 = list(ws2.iter_rows(values_only=True))
for r in rows_23[1:]:
    if not r or not r[0]: continue
    ref = str(r[0])
    if ref in EXCLUDED_REFS: continue
    family = get_family(ref)
    if family in ('EXCLUDED', 'AUTRES'): continue
    date = r[4]
    if date is None: continue
    if isinstance(date, str):
        try: date = pd.to_datetime(date, errors='coerce')
        except: continue
    if pd.isna(date): continue
    if date.year != 2023: continue
    
    agence_raw = str(r[14]).strip() if r[14] else ''
    if agence_raw not in ALL_AGENCES: continue
    agence_short, region = ALL_AGENCES[agence_raw]
    
    if family in ('MATERIEL_ELEVAGE', 'ALVEOLES'):
        tonnes_val = 0
    elif family == 'COMPLEMENT_ALIMENTAIRE':
        qte = r[3] if r[3] else 0
        weight = get_weight(ref)
        kg = qte * weight
        tonnes_val = kg / 1000
    else:
        tonnes_val = r[10] if r[10] else 0
        kg = tonnes_val * 1000
    
    qte = r[3] if r[3] else 0
    weight = get_weight(ref)
    sacs_50 = kg / 50 if family not in ('MATERIEL_ELEVAGE', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES') else 0
    montant_ht = r[2] if r[2] else 0
    
    all_records.append({
        'date': date, 'year': date.year, 'month': date.month,
        'ref': ref, 'family': family, 'description': r[1],
        'client': '', 'agence': agence_short, 'region': region,
        'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': tonnes_val,
        'sacs_50': sacs_50,
        'montant_ttc': 0, 'montant_ht': montant_ht,
        'source': 'LY_2023'
    })
print(f"  LY_2023 records: {len([r for r in all_records if r['source']=='LY_2023'])}")

# === 2025 ===
print("\nLoading 2025 (avec SPC agences)...")
wb = openpyxl.load_workbook("/home/z/my-project/upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx", read_only=True, data_only=True)
ws = wb['Feuil1']
header = None
for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
    header = row
cols_2025 = {}
for i, h in enumerate(header):
    if h == 'Réf. produit': cols_2025['ref'] = i
    elif h == 'Description du produit': cols_2025['desc'] = i
    elif h == 'Qté commandée': cols_2025['qte'] = i
    elif h == 'Tiers': cols_2025['client'] = i
    elif h == 'Date de commande': cols_2025['date'] = i
    elif h == 'Montant HT': cols_2025['montant_ht'] = i
    elif h == 'Montant TTC': cols_2025['montant_ttc'] = i
    elif h == 'État': cols_2025['etat'] = i
    elif h == 'tableauProprieteAgences.Agence': cols_2025['agence'] = i
    elif h == 'tableauProprieteAgences.Region': cols_2025['region'] = i

count_2025 = 0
for r in ws.iter_rows(min_row=2, values_only=True):
    if not r or len(r) <= max(cols_2025.values()): continue
    if r[cols_2025['ref']] == 'Total': continue
    etat = str(r[cols_2025['etat']]).strip() if r[cols_2025['etat']] else ''
    if etat != 'Livrée': continue
    ref = str(r[cols_2025['ref']])
    family = get_family(ref)
    if family in ('EXCLUDED', 'AUTRES'): continue
    client = r[cols_2025['client']]
    if is_internal_client(client): continue  # Exclude SPC SA, PDC, COMPTOIR as client
    agence_raw = r[cols_2025['agence']]
    if agence_raw not in ALL_AGENCES: continue
    agence_short, region = ALL_AGENCES[agence_raw]
    qte = r[cols_2025['qte']] or 0
    weight = get_weight(ref)
    kg = qte * weight
    if family in ('MATERIEL_ELEVAGE', 'ALVEOLES'):
        tonnes_val = 0
    elif family == 'COMPLEMENT_ALIMENTAIRE':
        tonnes_val = kg / 1000
    else:
        tonnes_val = kg / 1000
    sacs_50 = kg / 50 if family not in ('MATERIEL_ELEVAGE', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES') else 0
    montant_ttc = r[cols_2025['montant_ttc']] or 0
    montant_ht = r[cols_2025['montant_ht']] or 0
    date = parse_date(r[cols_2025['date']])
    if date is None: continue
    all_records.append({
        'date': date, 'year': date.year, 'month': date.month,
        'ref': ref, 'family': family, 'description': r[cols_2025['desc']],
        'client': client, 'agence': agence_short, 'region': region,
        'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': tonnes_val,
        'sacs_50': sacs_50,
        'montant_ttc': montant_ttc, 'montant_ht': montant_ht,
        'source': '2025'
    })
    count_2025 += 1
print(f"  2025 records: {count_2025}")

# === S1 2026 + Juil + Août ===
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

def process_2026(ws, source_label, date_filter=None):
    cols = detect_cols_2026(ws, header_row=2)
    if 'ref' not in cols: return []
    records = []
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) <= max(cols.values()): continue
        if r[cols['ref']] == 'Total': continue
        etat = str(r[cols['etat']]).strip() if r[cols['etat']] else ''
        if etat != 'Livrée': continue
        ref = str(r[cols['ref']])
        family = get_family(ref)
        if family in ('EXCLUDED', 'AUTRES'): continue
        client = r[cols.get('client', 0)] if 'client' in cols else ''
        if is_internal_client(client): continue
        agence_raw = r[cols['agence']]
        if agence_raw not in ALL_AGENCES: continue
        agence_short, region = ALL_AGENCES[agence_raw]
        date_str = str(r[cols['date']]) if r[cols['date']] else ''
        if date_filter and date_filter not in date_str: continue
        qte = r[cols['qte']] or 0
        weight = get_weight(ref)
        kg = qte * weight
        if family in ('MATERIEL_ELEVAGE', 'ALVEOLES'):
            tonnes_val = 0
        elif family == 'COMPLEMENT_ALIMENTAIRE':
            tonnes_val = kg / 1000
        else:
            tonnes_val = kg / 1000
        sacs_50 = kg / 50 if family not in ('MATERIEL_ELEVAGE', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES') else 0
        montant_ttc = r[cols['montant_ttc']] or 0
        montant_ht = r[cols['montant_ht']] or 0
        date = parse_date(r[cols['date']])
        if date is None: continue
        records.append({
            'date': date, 'year': date.year, 'month': date.month,
            'ref': ref, 'family': family, 'description': r[cols['desc']],
            'client': client, 'agence': agence_short, 'region': region,
            'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': tonnes_val,
            'sacs_50': sacs_50,
            'montant_ttc': montant_ttc, 'montant_ht': montant_ht,
            'source': source_label
        })
    return records

# S1 2026 (Jan-June)
print("\nLoading S1 2026 (avec SPC agences)...")
wb = openpyxl.load_workbook("/home/z/my-project/upload/ventes janv a juin 2026.xlsx", read_only=True, data_only=True)
records_s1 = []
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    cols = detect_cols_2026(ws, header_row=2)
    if 'ref' not in cols: continue
    records_s1.extend(process_2026(ws, 'S1_2026'))
print(f"  S1 2026 records: {len(records_s1)}")
all_records.extend(records_s1)

# Juillet 2026
print("\nLoading Juillet 2026 (avec SPC agences)...")
wb = openpyxl.load_workbook("/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (9).xlsx", read_only=True, data_only=True)
ws = wb['Sheet 1']
records_juil = process_2026(ws, 'Juil_2026', date_filter='/07/2026')
print(f"  Juil 2026 records: {len(records_juil)}")
all_records.extend(records_juil)

# Août 2026
print("\nLoading Août 2026 (avec SPC agences)...")
wb = openpyxl.load_workbook("/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (27).xlsx", read_only=True, data_only=True)
ws = wb['Sheet 1']
records_aout = []
for r in ws.iter_rows(min_row=3, values_only=True):
    if not r or len(r) < 16: continue
    if r[0] == 'Total': continue
    etat = str(r[13]).strip() if r[13] else ''
    if etat != 'Livrée': continue
    if not r[6] or '/08/2026' not in str(r[6]): continue
    if str(r[6]) == '27/08/2026': continue  # skip partial
    ref = str(r[0])
    family = get_family(ref)
    if family in ('EXCLUDED', 'AUTRES'): continue
    client = r[5] if r[5] else ''
    if is_internal_client(client): continue
    agence_raw = r[15]
    if agence_raw not in ALL_AGENCES: continue
    agence_short, region = ALL_AGENCES[agence_raw]
    qte = r[2] or 0
    weight = get_weight(ref)
    kg = qte * weight
    if family in ('MATERIEL_ELEVAGE', 'ALVEOLES'):
        tonnes_val = 0
    elif family == 'COMPLEMENT_ALIMENTAIRE':
        tonnes_val = kg / 1000
    else:
        tonnes_val = kg / 1000
    sacs_50 = kg / 50 if family not in ('MATERIEL_ELEVAGE', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES') else 0
    montant_ttc = r[9] or 0
    montant_ht = r[8] or 0
    date = parse_date(r[6])
    if date is None: continue
    records_aout.append({
        'date': date, 'year': date.year, 'month': date.month,
        'ref': ref, 'family': family, 'description': r[1],
        'client': client, 'agence': agence_short, 'region': region,
        'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': tonnes_val,
        'sacs_50': sacs_50,
        'montant_ttc': montant_ttc, 'montant_ht': montant_ht,
        'source': 'Aout_2026'
    })
print(f"  Août 2026 records: {len(records_aout)}")
all_records.extend(records_aout)

# === Consolidate ===
df = pd.DataFrame(all_records)
print(f"\nTotal records (incl SPC agencies + Maroua): {len(df)}")
print(f"Date range: {df['date'].min().date()} → {df['date'].max().date()}")
print(f"\nAgences ({df['agence'].nunique()}):")
for ag in sorted(df['agence'].unique()):
    sub = df[df['agence'] == ag]
    print(f"  {ag} ({sub['region'].iloc[0]}): {len(sub)} records, HT={sub['montant_ht'].sum()/1e6:.1f} M")

print(f"\nBy family:")
print(df['family'].value_counts())

print(f"\nBy year:")
print(df.groupby(df['date'].dt.year).size())

# Stats ALVEOLES by year (verify the spike pattern)
print("\nALVEOLES by year (HT in M FCFA):")
alv = df[df['family'] == 'ALVEOLES']
print(alv.groupby(alv['date'].dt.year).agg(
    records=('tonnes', 'count'),
    ht_m=('montant_ht', lambda x: x.sum() / 1e6),
    ttc_m=('montant_ttc', lambda x: x.sum() / 1e6),
).round({'ht_m': 1, 'ttc_m': 1}))

# ALVEOLES by year AND by agence
print("\nALVEOLES by year + agence (top 5):")
alv_by_ag = alv.groupby([alv['date'].dt.year, 'agence']).agg(
    records=('tonnes', 'count'),
    ht_m=('montant_ht', lambda x: x.sum() / 1e6),
).round({'ht_m': 1}).reset_index()
print(alv_by_ag.sort_values('ht_m', ascending=False).head(15))

# === Save ===
output_path = "/home/z/my-project/scripts/dataset_2023_2026.csv"
df.to_csv(output_path, index=False)
print(f"\nDataset saved: {output_path}")
print(f"File size: {os.path.getsize(output_path) / 1024:.0f} KB")
