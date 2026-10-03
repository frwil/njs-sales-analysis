"""
Construction du dataset consolidé 2024-2026 (3 ans) avec COMPLEMENT ALIMENTAIRE.

CHANGEMENTS vs version précédente (2021-2026):
1. Exclusion des années 2021-2023 (utilisateur: "utiliser les années 2024 à 2026 pour tout le forecast")
2. AJOUT de la famille COMPLEMENT_ALIMENTAIRE (BELGOKILL, BELGO HARMONY, etc.)
   - Conversion 1L = 1kg (donc 1 qte = 1 kg = 0.001 tonnes)
   - V305 (BELGOKILL 200L) = 200 kg par unité
3. Utilisation des tendances BELGOKILL (V300 + V305) comme proxy pour toute la famille

Sources:
  - LY_24 (Jul-Dec 2024): 54,944 rows (excl #N/A)
  - 2025 (Jan-Dec): 86d96135 file
  - 2026 (Jan-Sep): extractions globales BELGOCAM (54) Jan-Mar, (55) Avr-Jun, (56) Jul-Sep

Règle : toute vente dont le montant HT ou TTC est à 0 n'est pas intégrée.

Ventes internes (clients SPC/PDC/COMPTOIR/EMANA) : drapeau `interne=True`.
Dans le monthly performance, les clients internes rattachés aux agences classiques
sont comptés intégralement (CA + volumes) ; seules les agences SPC non mappées
(SPC BUEA, SPC PK15…) gardent leur CA mais pas leurs volumes.
Les analyses qui excluent (forecasts PDC/SPC, zéro achat) filtrent sur le drapeau.

Familles finales (7):
  - TOURTEAUX (soja)
  - CONCENTRES (BELGO Chair/Ponte/Porc)
  - INGREDIENTS (Belgotox, Bicarbonate, etc.)
  - ALIMENT_COMPLET (Chick/Piglet Booster, BELGO FISH)
  - COMPLEMENT_ALIMENTAIRE (BELGOKILL, BELGO HARMONY, etc.) - 1L=1kg
  - MATERIEL_ELEVAGE - CA only
  - PREMIX - CA only

Exclus:
  - MAIS (M1051, M1052) - produit opportuniste
  - DIVERS
"""
import openpyxl
import pandas as pd
import json
import os

# === Source files ===
FILE_LY_24 = "/home/z/my-project/upload/21_24.xlsx"  # LY_24 sheet (Jul-Dec 2024)
FILE_2025 = "/home/z/my-project/upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx"
# Extractions globales BELGOCAM 2026 : (fichier, (mois min, mois max)) — 17 colonnes, dates en chaînes
FILE_2026 = [
    ("/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (54).xlsx", (1, 3)),
    ("/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (55).xlsx", (4, 6)),
    ("/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (56).xlsx", (7, 9)),
]

# === Product refs and weights (kg per unit) ===
SOJA_REFS = {'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25}
CONC_REFS = {
    'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
    'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50,
    'C1022': 5,
}
INGREDIENT_REFS = {
    'B100': 25, 'E101': 25, 'I105': 25, 'B1001': 1, 'B1003': 5, 'B1004': 25,
    'E1011': 1, 'E1013': 5, 'E1014': 0.2,
    'I1051': 1, 'I1053': 5, 'I1054': 25,
    'I106': 25, 'I1061': 1, 'I107': 25, 'I1071': 1,
    'P105': 25, 'P1051': 1, 'P1053': 5,
    'F114': 50, 'F1145': 50, 'F1146': 25, 'F1147': 1,
    'F1143': 25, 'I1063': 5,
}
ALIMENT_REFS = {'CB100': 25, 'CB200': 25, 'CB101': 5, 'CB201': 5,
                'PB100': 25, 'PB200': 25, 'DB100': 25, 'DB200': 25,
                'ALAP25': 25,
                # BELGOFISH (aliments poissons) : séparés uniquement pour les forecasts 2027,
                # intégrés aux ALIMENTS COMPLETS dans les datasets et monthly performance
                'APCL2': 15, 'APCL25': 5, 'APCL3': 15, 'APCL30': 1, 'APCL35': 5,
                'APCL4.5': 15, 'APCL4.55': 5, 'APCL450': 1, 'APCL6': 15, 'APCL65': 5,
                'APCL8': 15, 'APCL80': 1, 'APCL85': 5}

# === NOUVEAU: COMPLEMENT ALIMENTAIRE (liquides) ===
# Tous les produits BELGOxxx liquides (1L = 1kg) + V305 (200L = 200kg)
# BELGOKILL = V300 (1L) + V305 (200L) - utilisé comme proxy pour la tendance famille
COMPLEMENT_REFS = {
    'V300': 1,        # BELGOKILL 1L = 1 kg
    'V305': 200,      # BELGOKILL 200L = 200 kg
    'CA003.1': 1,     # BELGO HARMONY 1L = 1 kg
    'CA004.1': 1,     # BELGO PROTECT 1L = 1 kg
    'CA006.1': 1,     # BELGO DRY LIT 1L = 1 kg
    'CA001.1': 1,     # BELGO WATER CLEAN 1L = 1 kg
    'CA002.1': 1,     # BELGO VIT Ese 1L = 1 kg
    'CA005.1': 1,     # BELGO THERMO 1L = 1 kg
    'CA007.1': 1,     # BELGO BIO SELECT 1L = 1 kg
    'CA008.1': 1,     # BELGO FRESH 1L = 1 kg
}

# MATERIEL ELEVAGE : pièces/paquets sans poids (règle métier) → poids 0, volume 0, CA uniquement
MATERIEL_REFS = {
    'MAT003': 0, 'MAT004': 0, 'MAT005': 0, 'MAT006': 0, 'MAT007': 0, 'MAT008': 0, 'MAT009': 0,
    'MAT011': 0, 'MAT014': 0, 'MAT015': 0, 'MAT017': 0,
    'MAT020': 0, 'MAT033': 0, 'MAT039': 0, 'MAT040': 0, 'MAT042': 0,
    'MAT047': 0, 'MAT049': 0, 'MAT050': 0, 'MAT054': 0, 'MAT055': 0, 'MAT073': 0,
    'MAT014-80010003': 0, 'MAT011-80010002': 0,
}
PREMIX_REFS = {'P102N2': 25, 'P104N2': 25, 'P109': 25, 'PX101': 25, 'PX102': 25, 'PX103': 25, 'PX104': 25, 'PX105': 25}

# Products to EXCLUDE
EXCLUDED_REFS = {'M1051', 'M1052'}  # MAIS

# Internal clients
INTERNAL_CLIENT_PATTERNS = ['SPC', 'PDC', 'COMPTOIR', 'EMANA']

# Agence map (covers both formats)
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

# Agence map for LY_24 (uses short names like 'FAMLA', 'NDOBO')
AGENCE_MAP_LY24 = {
    'FAMLA': ('Famla', 'Ouest'), 'DJELENG': ('Djeleng', 'Ouest'), 'MBOUDA': ('Mbouda', 'Ouest'),
    'MESSASSI': ('Messassi', 'Centre'), 'AHALA': ('Ahala', 'Centre'),
    'BERTOUA': ('Bertoua', 'Centre'), 'NGAOUNDERE': ('Ngaoundere', 'Centre'),
    'NKOLBISSON': ('Nkolbisson', 'Centre'), 'NKOABANG': ('Nkoabang', 'Centre'),
    'NDOBO': ('Ndobo', 'Littoral'), 'VILLAGE': ('Village', 'Littoral'),
    'PK11': ('Pk11', 'Littoral'), 'NKONGSAMBA': ('Nkongsamba', 'Littoral'),
    'BUEA': ('Buea', 'Littoral'),
}

# Region map for LY_24
def map_region(reg):
    reg = str(reg).strip() if reg else ''
    if reg in ('Yde', 'Est', 'Nord'): return 'Centre'
    return reg if reg in ('Ouest', 'Littoral', 'Centre') else None


def get_family(ref):
    if ref in EXCLUDED_REFS: return 'EXCLUDED'
    if ref in SOJA_REFS: return 'TOURTEAUX'
    if ref in CONC_REFS: return 'CONCENTRES'
    if ref in INGREDIENT_REFS: return 'INGREDIENTS'
    if ref in ALIMENT_REFS: return 'ALIMENT_COMPLET'
    if ref in COMPLEMENT_REFS: return 'COMPLEMENT_ALIMENTAIRE'
    if ref in MATERIEL_REFS: return 'MATERIEL_ELEVAGE'
    if ref in PREMIX_REFS: return 'PREMIX'
    if ref.startswith('MAT'): return 'MATERIEL_ELEVAGE'
    if ref.startswith('M'): return 'EXCLUDED'  # MAIS et autres
    return 'AUTRES'


def get_weight(ref):
    if ref in SOJA_REFS: return SOJA_REFS[ref]
    if ref in CONC_REFS: return CONC_REFS[ref]
    if ref in INGREDIENT_REFS: return INGREDIENT_REFS[ref]
    if ref in ALIMENT_REFS: return ALIMENT_REFS[ref]
    if ref in COMPLEMENT_REFS: return COMPLEMENT_REFS[ref]
    if ref in MATERIEL_REFS: return MATERIEL_REFS[ref]
    if ref in PREMIX_REFS: return PREMIX_REFS[ref]
    return 1


def is_internal_client(client_str):
    if not client_str: return False
    s = str(client_str).upper()
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


# === Load LY_24 (Jul-Dec 2024) ===
print("=" * 70)
print("Loading LY_24 (Jul-Dec 2024)...")
print("=" * 70)
wb = openpyxl.load_workbook(FILE_LY_24, read_only=True, data_only=True)
ws = wb['LY_24']
rows_24 = list(ws.iter_rows(values_only=True))
print(f"  {len(rows_24)-1} rows")

# LY_24 columns: 0=ref, 1=desc, 2=qte, 5=date, 7=montantHT, 10=etat, 13=cat, 14=agence, 19=region
records_24 = []
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
    
    # Map agence using LY_24 format
    found = False
    agence_short, region_default = None, None
    for key, (a, reg) in AGENCE_MAP_LY24.items():
        if a.lower() == agence.lower():
            agence_short = a
            region_default = reg
            found = True
            break
    if not found: continue
    
    region_raw = str(r[19]).strip() if r[19] else region_default
    region = map_region(region_raw) or region_default
    
    qte = r[2] if r[2] else 0
    weight = get_weight(ref)
    # MATERIEL_ELEVAGE : pas de volume, uniquement le CA (règle métier)
    kg = 0 if family == 'MATERIEL_ELEVAGE' else qte * weight
    # For COMPLEMENT_ALIMENTAIRE and MATERIEL, no "sacs_50" notion
    sacs_50 = kg / 50 if family not in ('MATERIEL_ELEVAGE', 'COMPLEMENT_ALIMENTAIRE') else 0
    
    montant_ht = r[7] if r[7] else 0
    # Règle : une vente dont le montant HT est à 0 n'est pas intégrée
    if not isinstance(montant_ht, (int, float)) or montant_ht <= 0: continue
    date_str = str(r[5])[:10] if r[5] else ''
    try:
        date = pd.to_datetime(date_str, format='%d/%m/%Y')
    except:
        continue
    if pd.isna(date): continue
    
    records_24.append({
        'date': date, 'year': date.year, 'month': date.month,
        'ref': ref, 'family': family, 'description': r[1],
        'client': '', 'agence': agence_short, 'region': region,
        'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': kg / 1000,
        'sacs_50': sacs_50,
        'montant_ttc': 0, 'montant_ht': montant_ht,
        'source': 'LY_24', 'interne': False
    })
print(f"  Records loaded: {len(records_24)}")
df_24 = pd.DataFrame(records_24)


# === Load 2025 (Jan-Dec) ===
print("\n" + "=" * 70)
print("Loading 2025 (Jan-Dec)...")
print("=" * 70)
wb = openpyxl.load_workbook(FILE_2025, read_only=True, data_only=True)
ws = wb['Feuil1']

# Detect columns
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

records_2025 = []
for r in ws.iter_rows(min_row=2, values_only=True):
    if not r or len(r) <= max(cols_2025.values()): continue
    if r[cols_2025['ref']] == 'Total': continue
    etat = str(r[cols_2025['etat']]).strip() if r[cols_2025['etat']] else ''
    if etat != 'Livrée': continue
    ref = str(r[cols_2025['ref']])
    family = get_family(ref)
    if family in ('EXCLUDED', 'AUTRES'): continue
    client = r[cols_2025['client']]
    # Clients internes (COMPTOIR/PDC/SPC/EMANA) rattachés aux agences classiques :
    # comptés intégralement (CA + volumes). Agences SPC non mappées : CA oui, volumes non.
    # Le drapeau `interne` reste posé pour les analyses qui excluent (forecasts, zéro achat).
    interne = is_internal_client(client)
    agence_raw = r[cols_2025['agence']]
    if agence_raw in AGENCE_MAP:
        agence_short, region = AGENCE_MAP[agence_raw]
        vol_zero = False
    elif interne:
        agence_short, region = (str(agence_raw).strip() or 'SPC'), None
        vol_zero = True
    else:
        continue
    qte = r[cols_2025['qte']] or 0
    weight = get_weight(ref)
    # MATERIEL_ELEVAGE : pas de volume, uniquement le CA (règle métier)
    kg = 0 if (vol_zero or family == 'MATERIEL_ELEVAGE') else qte * weight
    sacs_50 = 0 if (vol_zero or family in ('MATERIEL_ELEVAGE', 'COMPLEMENT_ALIMENTAIRE')) else kg / 50
    montant_ttc = r[cols_2025['montant_ttc']] or 0
    montant_ht = r[cols_2025['montant_ht']] or 0
    # Règle : une vente dont le montant HT ou TTC est à 0 n'est pas intégrée
    if not isinstance(montant_ht, (int, float)) or montant_ht <= 0 \
       or not isinstance(montant_ttc, (int, float)) or montant_ttc == 0: continue
    date = parse_date(r[cols_2025['date']])
    if date is None: continue
    records_2025.append({
        'date': date, 'year': date.year, 'month': date.month,
        'ref': ref, 'family': family, 'description': r[cols_2025['desc']],
        'client': client, 'agence': agence_short, 'region': region,
        'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': kg / 1000,
        'sacs_50': sacs_50,
        'montant_ttc': montant_ttc, 'montant_ht': montant_ht,
        'source': '2025', 'interne': interne
    })
print(f"  2025: {len(records_2025)} records")


# === Load 2026 (Jan-Sep) : extractions globales BELGOCAM (54)/(55)/(56) ===
print("\n" + "=" * 70)
print("Loading 2026 (Jan-Sep) — extractions globales (54)/(55)/(56)...")
print("=" * 70)

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

records_2026 = []
for path, (mois_min, mois_max) in FILE_2026:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['Sheet 1']
    cols = detect_cols_2026(ws, header_row=2)
    if 'ref' not in cols: continue
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) <= max(cols.values()): continue
        if r[cols['ref']] == 'Total': continue
        etat = str(r[cols['etat']]).strip() if r[cols['etat']] else ''
        if etat != 'Livrée': continue
        ref = str(r[cols['ref']])
        family = get_family(ref)
        if family in ('EXCLUDED', 'AUTRES'): continue
        client = r[cols['client']]
        # Clients internes (COMPTOIR/PDC/SPC/EMANA) rattachés aux agences classiques :
        # comptés intégralement (CA + volumes). Agences SPC non mappées : CA oui, volumes non.
        # Le drapeau `interne` reste posé pour les analyses qui excluent (forecasts, zéro achat).
        interne = is_internal_client(client)
        agence_raw = r[cols['agence']]
        if agence_raw in AGENCE_MAP:
            agence_short, region = AGENCE_MAP[agence_raw]
            vol_zero = False
        elif interne:
            agence_short, region = (str(agence_raw).strip() or 'SPC'), None
            vol_zero = True
        else:
            continue
        date = parse_date(r[cols['date']])
        if date is None: continue
        # Filtre mensuel du fichier (garde-fou anti-doublon / hors période)
        if date.year != 2026 or not (mois_min <= date.month <= mois_max): continue
        qte = r[cols['qte']] or 0
        weight = get_weight(ref)
        # MATERIEL_ELEVAGE : pas de volume, uniquement le CA (règle métier)
        kg = 0 if (vol_zero or family == 'MATERIEL_ELEVAGE') else qte * weight
        sacs_50 = 0 if (vol_zero or family in ('MATERIEL_ELEVAGE', 'COMPLEMENT_ALIMENTAIRE')) else kg / 50
        montant_ttc = r[cols['montant_ttc']] or 0
        montant_ht = r[cols['montant_ht']] or 0
        # Règle : une vente dont le montant HT ou TTC est à 0 n'est pas intégrée
        if not isinstance(montant_ht, (int, float)) or montant_ht <= 0 \
           or not isinstance(montant_ttc, (int, float)) or montant_ttc == 0: continue
        records_2026.append({
            'date': date, 'year': date.year, 'month': date.month,
            'ref': ref, 'family': family, 'description': r[cols['desc']],
            'client': client, 'agence': agence_short, 'region': region,
            'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': kg / 1000,
            'sacs_50': sacs_50,
            'montant_ttc': montant_ttc, 'montant_ht': montant_ht,
            'source': 'BELGOCAM_2026', 'interne': interne
        })
    wb.close()
print(f"  2026 (Jan-Sep): {len(records_2026)} records")


# === Consolidate ===
print("\n" + "=" * 70)
print("Consolidating 2024-2026...")
print("=" * 70)
all_records = records_24 + records_2025 + records_2026
df = pd.DataFrame(all_records)
print(f"Total records: {len(df)}")
print(f"Date range: {df['date'].min().date()} → {df['date'].max().date()}")
print(f"\nRecords by year:")
print(df.groupby(df['date'].dt.year).size())
print(f"\nRecords by family:")
print(df['family'].value_counts())

# Save
output_path = "/home/z/my-project/scripts/dataset_2024_2026.csv"
df.to_csv(output_path, index=False)
print(f"\nDataset saved: {output_path}")
print(f"File size: {os.path.getsize(output_path) / 1024:.0f} KB")

# Stats by family
print("\n" + "=" * 70)
print("STATS PAR FAMILLE (2024-2026)")
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

# Stats COMPLEMENT_ALIMENTAIRE par année
print("\n" + "=" * 70)
print("COMPLEMENT_ALIMENTAIRE par année (proxy BELGOKILL)")
print("=" * 70)
ca = df[df['family'] == 'COMPLEMENT_ALIMENTAIRE']
print(ca.groupby(ca['date'].dt.year).agg(
    tonnes=('tonnes', 'sum'),
    litres=('qte', 'sum'),
    ca_ht_m=('montant_ht', lambda x: x.sum() / 1e6),
    n_records=('tonnes', 'count'),
).round({'tonnes': 2, 'ca_ht_m': 1}))

# Stats COMPLEMENT par ref
print("\n=== COMPLEMENT_ALIMENTAIRE par ref ===")
ca_by_ref = ca.groupby('ref').agg(
    tonnes=('tonnes', 'sum'),
    litres=('qte', 'sum'),
    ca_ht_m=('montant_ht', lambda x: x.sum() / 1e6),
    n_records=('tonnes', 'count'),
).round({'tonnes': 2, 'ca_ht_m': 1}).sort_values('ca_ht_m', ascending=False)
print(ca_by_ref)

print("\n" + "=" * 70)
print(f"DATASET 2024-2026 PRÊT: {len(df):,} records, {df['ref'].nunique()} produits, {df['agence'].nunique()} agences")
print("=" * 70)
