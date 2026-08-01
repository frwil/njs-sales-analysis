"""
BELGOCAM SA - Recompute July MTD metrics for zero achat + bundle PDFs.
Uses new extraction (au 24/07/2026) and S1 2026 data.

Outputs:
- /home/z/my-project/scripts/juillet_mtd_24.json  (clients, habitudes, top risque)
- /home/z/my-project/scripts/cross_sell_mtd_24.json  (soja→conc, booster→conc)
- /home/z/my-project/scripts/volumes_juillet_24.json  (volumes by category, projection)
- /home/z/my-project/scripts/bundle_metrics_24.json  (bundle global + par agence)
"""
import openpyxl
from collections import defaultdict, Counter
import datetime
import json

# ===== CONFIG =====
S1_FILE = '/home/z/my-project/upload/ventes janv a juin 2026.xlsx'
JULY_FILE = '/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (3) (10) (1).xlsx'

JULY_END_DAY = 31  # au 31/07/2026 — bilan complet définitif (27j lun-sam)
JULY_DAYS_LUN_SAM = sum(1 for d in range(1, JULY_END_DAY + 1) if datetime.date(2026, 7, d).weekday() < 6)
JULY_TOTAL_DAYS_LUN_SAM = sum(1 for d in range(1, 32) if datetime.date(2026, 7, d).weekday() < 6)
TEMPS_ECOULE_PCT = JULY_DAYS_LUN_SAM / JULY_TOTAL_DAYS_LUN_SAM * 100

# Product references (sac weight in kg)
SOJA_REFS = {'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25}
CONC_REFS = {
    'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
    'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50
}
# Booster products (Chick/Piglet Booster)
BOOSTER_REFS = {'CB100', 'CB101', 'CB200', 'CB201', 'ALAP25', 'APCL2', 'APCL25', 'APCL3', 'APCL4.5', 'APCL4.55', 'APCL6', 'APCL65', 'APCL8', 'APCL85'}
# Product category mapping (subset)
PROD_CAT = {
    'T102': 'TOURTEAUX', 'T1021': 'TOURTEAUX', 'T1023': 'TOURTEAUX', 'T1024': 'TOURTEAUX',
    'C101': 'CONCENTRES', 'C102': 'CONCENTRES', 'C103': 'CONCENTRES', 'C104': 'CONCENTRES',
    'C1042': 'CONCENTRES', 'C1043': 'CONCENTRES', 'C1044': 'CONCENTRES',
    'C105': 'CONCENTRES', 'C1053': 'CONCENTRES', 'C1054': 'CONCENTRES', 'C1055': 'CONCENTRES', 'C108': 'CONCENTRES',
    'P102N2': 'PREMIX', 'P104N2': 'PREMIX', 'P109': 'PREMIX',
    'M1051': 'INGREDIENTS',  # Maïs
    'I105': 'INGREDIENTS', 'I1051': 'INGREDIENTS', 'I106': 'INGREDIENTS', 'I1061': 'INGREDIENTS',
    'I107': 'INGREDIENTS', 'I1071': 'INGREDIENTS',
    'F114': 'INGREDIENTS', 'F1142': 'INGREDIENTS', 'F1145': 'INGREDIENTS', 'F1146': 'INGREDIENTS', 'F1147': 'INGREDIENTS',
    'B100': 'INGREDIENTS', 'B1001': 'INGREDIENTS',
    'CF101': 'INGREDIENTS', 'CF1012': 'INGREDIENTS',
    'S101': 'INGREDIENTS',
    'E101': 'INGREDIENTS', 'E1011': 'INGREDIENTS', 'E1014': 'INGREDIENTS',
    'P105': 'INGREDIENTS', 'P1051': 'INGREDIENTS', 'P1053': 'INGREDIENTS',
    'ALAP25': 'ALIMENT COMPLET', 'CB100': 'ALIMENT COMPLET', 'CB101': 'ALIMENT COMPLET',
    'CB200': 'ALIMENT COMPLET', 'CB201': 'ALIMENT COMPLET',
    'APCL2': 'ALIMENT COMPLET', 'APCL25': 'ALIMENT COMPLET', 'APCL3': 'ALIMENT COMPLET',
    'APCL4.5': 'ALIMENT COMPLET', 'APCL4.55': 'ALIMENT COMPLET', 'APCL6': 'ALIMENT COMPLET',
    'APCL65': 'ALIMENT COMPLET', 'APCL8': 'ALIMENT COMPLET', 'APCL85': 'ALIMENT COMPLET',
}
# Manual weights (for some ingredients)
MANUAL_WEIGHTS = {'M1051': 50.0, 'CF101': 1.0, 'S101': 1.0}

# Weights for INGREDIENTS products (kg per unit)
ING_WEIGHTS = {
    'I105': 25, 'I1051': 1,        # Sulfate de fer
    'I106': 25, 'I1061': 1,        # Methionine
    'I107': 25, 'I1071': 1,        # Lysine
    'F114': 50, 'F1142': 25, 'F1145': 50, 'F1146': 25, 'F1147': 1,  # Farine poisson
    'B100': 25, 'B1001': 1,        # Bicarbonate
    'CF101': 1, 'CF1012': 50,      # Coquillage (already mapped)
    'S101': 1,                      # Sel
    'E101': 25, 'E1011': 1, 'E1014': 5,  # Belgotox
    'P105': 25, 'P1051': 1, 'P1053': 5,  # Belgofos
}

# Agence map
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
}

# Internal clients (excluded from zero achat analysis)
INTERNAL_CLIENTS_FILE = '/home/z/my-project/scripts/excluded_clients.json'


def load_internal_clients():
    try:
        with open(INTERNAL_CLIENTS_FILE) as f:
            return set(json.load(f))
    except Exception:
        return set()


def load_s1_data():
    """Load S1 2026 sales data. Returns list of rows."""
    wb = openpyxl.load_workbook(S1_FILE, read_only=True)
    rows = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        for i, r in enumerate(ws.iter_rows(min_row=3, values_only=True)):
            if not r or len(r) < 18:
                continue
            if r[0] == 'Total':
                continue
            if r[15] != 'Livrée':
                continue
            rows.append(r)
    return rows


def load_july_data():
    """Load July 2026 sales data (Livrée only)."""
    wb = openpyxl.load_workbook(JULY_FILE, read_only=True)
    ws = wb['Sheet 1']
    rows = []
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) < 18:
            continue
        if r[0] == 'Total':
            continue
        if r[15] != 'Livrée':
            continue
        rows.append(r)
    return rows


def get_client_key(tiers):
    """Extract client code from Tiers field 'CU2407-00794 -  PENKA DEFFO DANIEL'."""
    if not tiers:
        return None
    s = str(tiers)
    if ' - ' in s:
        return s.split(' - ', 1)[0].strip()
    return s.strip()


def get_client_name(tiers):
    """Extract client name from Tiers field."""
    if not tiers:
        return ''
    s = str(tiers)
    if ' - ' in s:
        return s.split(' - ', 1)[1].strip()
    return s.strip()


def is_comptoir(name):
    """Check if client name is a 'CLIENT COMPTOIR' (aggregate, internal)."""
    if not name:
        return False
    n = name.upper()
    return ('COMPTOIR' in n or 'CLIENT COMPTOIR' in n)


def compute_volumes_by_cat(rows):
    """Compute volumes (t) by product category. Maïs is separated from other INGREDIENTS."""
    cat_kg = defaultdict(float)
    for r in rows:
        ref = r[0]
        qte = r[2] or 0
        if ref == 'M1051':
            # Maïs - separated from other INGREDIENTS
            kg = qte * MANUAL_WEIGHTS[ref]
            cat_kg['MAÏS'] += kg
            continue
        if ref in MANUAL_WEIGHTS:
            kg = qte * MANUAL_WEIGHTS[ref]
        elif ref in SOJA_REFS:
            kg = qte * SOJA_REFS[ref]
        elif ref in CONC_REFS:
            kg = qte * CONC_REFS[ref]
        elif ref in ING_WEIGHTS:
            kg = qte * ING_WEIGHTS[ref]
        elif ref in PROD_CAT:
            cat = PROD_CAT[ref]
            if cat == 'PREMIX':
                # Premix usually 5kg or 25kg - check 25 first to avoid matching "5 KG" in "25 KG"
                desc = str(r[1] or '').upper()
                if '25KG' in desc.replace(' ', '') or '25 KG' in desc:
                    kg = qte * 25
                elif '5KG' in desc.replace(' ', '') or ' 5 KG' in desc or desc.endswith('5 KG'):
                    kg = qte * 5
                elif '1KG' in desc.replace(' ', '') or ' 1 KG' in desc:
                    kg = qte * 1
                else:
                    kg = qte * 25  # default for premix
            elif cat == 'ALIMENT COMPLET':
                desc = str(r[1] or '').upper()
                if '50KG' in desc.replace(' ', '') or '50 KG' in desc:
                    kg = qte * 50
                elif '25KG' in desc.replace(' ', '') or '25 KG' in desc:
                    kg = qte * 25
                else:
                    kg = qte * 25  # default for small bags
            else:
                # Default 50kg
                kg = qte * 50
        else:
            continue
        cat = PROD_CAT.get(ref, 'AUTRE')
        cat_kg[cat] += kg
    return cat_kg


def compute_clients_per_month(rows):
    """For each month, get set of unique client keys."""
    clients_by_month = defaultdict(set)
    for r in rows:
        date = str(r[6])
        if not date or len(date) < 7:
            continue
        month = date[3:10]  # MM/YYYY
        ck = get_client_key(r[5])
        if ck:
            clients_by_month[month].add(ck)
    return clients_by_month


def compute_juillet_mtd_metrics(s1_rows, july_rows, internal_clients):
    """Compute the main July MTD metrics for zero achat PDF."""
    # S1 clients (unique external clients)
    s1_clients = {}  # key -> {'name', 'agence', 'region', 'ca_s1', 'nb_achats', 'jours'}
    s1_client_days = defaultdict(list)  # key -> [days of month when purchased]
    s1_client_ca = defaultdict(float)
    s1_client_agence = {}
    s1_client_nb = defaultdict(int)

    for r in s1_rows:
        ck = get_client_key(r[5])
        if not ck or ck in internal_clients:
            continue
        name = get_client_name(r[5])
        if is_comptoir(name):
            continue  # exclude aggregate comptoir clients
        ca = r[8] or 0  # Montant HT
        date_str = str(r[6])
        try:
            day = int(date_str[:2])
        except:
            continue
        agence_raw = r[17] or ''
        agence, region = AGENCE_MAP.get(agence_raw, (agence_raw, '?'))
        s1_clients[ck] = {'name': name, 'agence': agence, 'region': region}
        s1_client_days[ck].append(day)
        s1_client_ca[ck] += ca
        s1_client_nb[ck] += 1
        s1_client_agence[ck] = (agence, region)

    # July clients
    july_clients = set()
    for r in july_rows:
        ck = get_client_key(r[5])
        if not ck or ck in internal_clients:
            continue
        name = get_client_name(r[5])
        if is_comptoir(name):
            continue
        july_clients.add(ck)

    # S1 averages per client
    s1_total = len(s1_clients)
    maintenus = 0
    nouveaux = 0
    risque_eleve = 0
    a_surveiller = 0
    pas_inquietant = 0
    ca_risque_eleve = 0
    ca_a_surveiller = 0
    ca_pas_inquietant = 0
    ca_maintenus = 0
    top_risque = []

    for ck, info in s1_clients.items():
        ca_s1 = s1_client_ca[ck]
        ca_s1_m = ca_s1 / 6  # monthly average
        days = s1_client_days[ck]
        # Compute mean day of purchase
        if days:
            jour_moyen = sum(days) / len(days)
        else:
            jour_moyen = 0
        nb_achats = s1_client_nb[ck]

        if ck in july_clients:
            maintenus += 1
            ca_maintenus += ca_s1
        else:
            # Classify by mean day
            if jour_moyen <= 21:
                risque_eleve += 1
                ca_risque_eleve += ca_s1
                top_risque.append({
                    'nom': info['name'],
                    'agence': info['agence'],
                    'region': info['region'],
                    'ca_s1_m': round(ca_s1_m / 1_000_000, 1),  # in millions
                    'jour_moyen': round(jour_moyen, 1),
                    'nb_achats': nb_achats,
                })
            elif jour_moyen <= 25:
                a_surveiller += 1
                ca_a_surveiller += ca_s1
            else:
                pas_inquietant += 1
                ca_pas_inquietant += ca_s1

    # Nouveaux = July clients not in S1
    nouveaux = len(july_clients - set(s1_clients.keys()))

    # Sort top risque by CA
    top_risque.sort(key=lambda x: -x['ca_s1_m'])
    top_risque_10 = top_risque[:10]

    return {
        'temps_ecoule_pct': round(TEMPS_ECOULE_PCT, 1),
        'july_days_lun_sam': JULY_DAYS_LUN_SAM,
        'july_total_days_lun_sam': JULY_TOTAL_DAYS_LUN_SAM,
        'clients': {
            's1_total': s1_total,
            'juil_mtd': len(july_clients),
            'maintenus': maintenus,
            'nouveaux': nouveaux,
            'risque_eleve': risque_eleve,
            'a_surveiller': a_surveiller,
            'pas_inquietant': pas_inquietant,
        },
        'ca': {
            'risque_eleve': round(ca_risque_eleve, 0),
            'a_surveiller': round(ca_a_surveiller, 0),
            'pas_inquietant': round(ca_pas_inquietant, 0),
            'maintenus': round(ca_maintenus, 0),
        },
        'top_risque': top_risque_10,
    }


def compute_cross_sell(s1_rows, july_rows, internal_clients):
    """Compute cross-sell rates: soja→conc, booster→conc, S1 vs July."""
    # For each client, track: bought_soja, bought_conc, bought_booster (per month for S1, total for July)
    # S1: per month
    s1_month_clients = defaultdict(lambda: {'soja': False, 'conc': False, 'booster': False})
    for r in s1_rows:
        ck = get_client_key(r[5])
        if not ck or ck in internal_clients:
            continue
        name = get_client_name(r[5])
        if is_comptoir(name):
            continue
        date_str = str(r[6])
        if not date_str or len(date_str) < 7:
            continue
        month = date_str[3:10]
        ref = r[0]
        if ref in SOJA_REFS:
            s1_month_clients[(month, ck)]['soja'] = True
        if ref in CONC_REFS:
            s1_month_clients[(month, ck)]['conc'] = True
        if ref in BOOSTER_REFS:
            s1_month_clients[(month, ck)]['booster'] = True

    # S1 averages
    s1_soja_clients = set()
    s1_soja_conc_clients = set()
    s1_soja_only_clients = set()
    s1_booster_clients = set()
    s1_booster_conc_clients = set()
    s1_booster_only_clients = set()
    s1_soja_total_months = 0
    s1_soja_conc_total_months = 0
    s1_booster_total_months = 0
    s1_booster_conc_total_months = 0

    # Group by client
    s1_client_months = defaultdict(lambda: {'soja_months': 0, 'conc_months': 0, 'booster_months': 0,
                                             'soja_conc_months': 0, 'booster_conc_months': 0})
    for (month, ck), cats in s1_month_clients.items():
        if cats['soja']:
            s1_client_months[ck]['soja_months'] += 1
            if cats['conc']:
                s1_client_months[ck]['soja_conc_months'] += 1
        if cats['booster']:
            s1_client_months[ck]['booster_months'] += 1
            if cats['conc']:
                s1_client_months[ck]['booster_conc_months'] += 1

    # Client is "soja client" if bought soja at least once in S1
    # Cross-sell rate = # months where soja+conc / # months where soja
    for ck, m in s1_client_months.items():
        if m['soja_months'] > 0:
            s1_soja_clients.add(ck)
            s1_soja_total_months += m['soja_months']
            s1_soja_conc_total_months += m['soja_conc_months']
            if m['soja_conc_months'] > 0:
                s1_soja_conc_clients.add(ck)
            else:
                s1_soja_only_clients.add(ck)
        if m['booster_months'] > 0:
            s1_booster_clients.add(ck)
            s1_booster_total_months += m['booster_months']
            s1_booster_conc_total_months += m['booster_conc_months']
            if m['booster_conc_months'] > 0:
                s1_booster_conc_clients.add(ck)
            else:
                s1_booster_only_clients.add(ck)

    # July
    july_client = defaultdict(lambda: {'soja': False, 'conc': False, 'booster': False})
    for r in july_rows:
        ck = get_client_key(r[5])
        if not ck or ck in internal_clients:
            continue
        name = get_client_name(r[5])
        if is_comptoir(name):
            continue
        ref = r[0]
        if ref in SOJA_REFS:
            july_client[ck]['soja'] = True
        if ref in CONC_REFS:
            july_client[ck]['conc'] = True
        if ref in BOOSTER_REFS:
            july_client[ck]['booster'] = True

    july_soja_clients = set()
    july_soja_conc_clients = set()
    july_soja_only_clients = set()
    july_booster_clients = set()
    july_booster_conc_clients = set()
    july_booster_only_clients = set()

    for ck, cats in july_client.items():
        if cats['soja']:
            july_soja_clients.add(ck)
            if cats['conc']:
                july_soja_conc_clients.add(ck)
            else:
                july_soja_only_clients.add(ck)
        if cats['booster']:
            july_booster_clients.add(ck)
            if cats['conc']:
                july_booster_conc_clients.add(ck)
            else:
                july_booster_only_clients.add(ck)

    # Regression: clients who had soja+conc in S1 but only soja in July
    soja_regression = len(s1_soja_conc_clients & july_soja_only_clients)
    # Conversion: clients who had soja-only in S1 but soja+conc in July
    soja_conversion = len(s1_soja_only_clients & july_soja_conc_clients)
    booster_regression = len(s1_booster_conc_clients & july_booster_only_clients)
    booster_conversion = len(s1_booster_only_clients & july_booster_conc_clients)

    # Cross-sell rates (% of soja clients who also buy conc)
    s1_soja_conc_pct = round(len(s1_soja_conc_clients) / len(s1_soja_clients) * 100, 0) if s1_soja_clients else 0
    july_soja_conc_pct = round(len(july_soja_conc_clients) / len(july_soja_clients) * 100, 0) if july_soja_clients else 0
    s1_booster_conc_pct = round(len(s1_booster_conc_clients) / len(s1_booster_clients) * 100, 0) if s1_booster_clients else 0
    july_booster_conc_pct = round(len(july_booster_conc_clients) / len(july_booster_clients) * 100, 0) if july_booster_clients else 0

    return {
        'soja': {
            's1_total': len(s1_soja_clients),
            's1_soja_conc': len(s1_soja_conc_clients),
            's1_soja_conc_pct': s1_soja_conc_pct,
            's1_soja_only': len(s1_soja_only_clients),
            's1_soja_only_pct': 100 - s1_soja_conc_pct,
            'juil_soja': len(july_soja_clients),
            'juil_soja_conc': len(july_soja_conc_clients),
            'juil_soja_conc_pct': july_soja_conc_pct,
            'juil_soja_only': len(july_soja_only_clients),
            'juil_soja_only_pct': 100 - july_soja_conc_pct,
            'regression': soja_regression,
            'conversion': soja_conversion,
        },
        'booster': {
            's1_total': len(s1_booster_clients),
            's1_booster_conc': len(s1_booster_conc_clients),
            's1_booster_conc_pct': s1_booster_conc_pct,
            's1_booster_only': len(s1_booster_only_clients),
            's1_booster_only_pct': 100 - s1_booster_conc_pct,
            'juil_booster': len(july_booster_clients),
            'juil_booster_conc': len(july_booster_conc_clients),
            'juil_booster_conc_pct': july_booster_conc_pct,
            'juil_booster_only': len(july_booster_only_clients),
            'juil_booster_only_pct': 100 - july_booster_conc_pct,
            'regression': booster_regression,
            'conversion': booster_conversion,
        },
    }


def compute_volumes(s1_rows, july_rows):
    """Compute volumes by category for S1 (monthly avg) and July MTD."""
    s1_cat = compute_volumes_by_cat(s1_rows)
    july_cat = compute_volumes_by_cat(july_rows)
    # S1 monthly avg = total / 6
    s1_monthly = {k: v / 6 / 1000 for k, v in s1_cat.items()}  # in tons
    july_mtd = {k: v / 1000 for k, v in july_cat.items()}  # in tons
    # Projection = july_mtd / temps_ecoule
    proj_factor = JULY_TOTAL_DAYS_LUN_SAM / JULY_DAYS_LUN_SAM
    july_proj = {k: v * proj_factor for k, v in july_mtd.items()}
    return {
        's1_monthly_t': {k: round(v, 0) for k, v in s1_monthly.items()},
        'juil_mtd_t': {k: round(v, 0) for k, v in july_mtd.items()},
        'juil_proj_t': {k: round(v, 0) for k, v in july_proj.items()},
    }


def compute_bundle_metrics(s1_rows, july_rows, internal_clients):
    """Compute bundle metrics: price, budget, ratio by client - global + per agency."""
    # For each client-month in S1, compute soja_sacs, conc_sacs, soja_CA, conc_CA
    s1_client_month = defaultdict(lambda: {
        'soja_sacs': 0, 'conc_sacs': 0, 'soja_ca': 0, 'conc_ca': 0,
        'agence': None, 'region': None
    })
    for r in s1_rows:
        ck = get_client_key(r[5])
        if not ck or ck in internal_clients:
            continue
        name = get_client_name(r[5])
        if is_comptoir(name):
            continue
        date_str = str(r[6])
        if not date_str or len(date_str) < 7:
            continue
        month = date_str[3:10]
        ref = r[0]
        qte = r[2] or 0
        ca = r[8] or 0
        agence_raw = r[17] or ''
        agence, region = AGENCE_MAP.get(agence_raw, (agence_raw, '?'))
        key = (ck, month)
        if ref in SOJA_REFS:
            s1_client_month[key]['soja_sacs'] += qte * SOJA_REFS[ref] / 50
            s1_client_month[key]['soja_ca'] += ca
            s1_client_month[key]['agence'] = agence
            s1_client_month[key]['region'] = region
        if ref in CONC_REFS:
            s1_client_month[key]['conc_sacs'] += qte * CONC_REFS[ref] / 50
            s1_client_month[key]['conc_ca'] += ca
            s1_client_month[key]['agence'] = agence
            s1_client_month[key]['region'] = region

    # July
    july_client = defaultdict(lambda: {
        'soja_sacs': 0, 'conc_sacs': 0, 'soja_ca': 0, 'conc_ca': 0,
        'agence': None, 'region': None
    })
    for r in july_rows:
        ck = get_client_key(r[5])
        if not ck or ck in internal_clients:
            continue
        name = get_client_name(r[5])
        if is_comptoir(name):
            continue
        ref = r[0]
        qte = r[2] or 0
        ca = r[8] or 0
        agence_raw = r[17] or ''
        agence, region = AGENCE_MAP.get(agence_raw, (agence_raw, '?'))
        if ref in SOJA_REFS:
            july_client[ck]['soja_sacs'] += qte * SOJA_REFS[ref] / 50
            july_client[ck]['soja_ca'] += ca
            july_client[ck]['agence'] = agence
            july_client[ck]['region'] = region
        if ref in CONC_REFS:
            july_client[ck]['conc_sacs'] += qte * CONC_REFS[ref] / 50
            july_client[ck]['conc_ca'] += ca
            july_client[ck]['agence'] = agence
            july_client[ck]['region'] = region

    # Filter: bundle clients = those with soja AND conc > 0
    s1_bundle_per_month = [v for v in s1_client_month.values() if v['soja_sacs'] > 0 and v['conc_sacs'] > 0]
    july_bundle = [v for v in july_client.values() if v['soja_sacs'] > 0 and v['conc_sacs'] > 0]

    # Global metrics - S1
    s1_total_soja_sacs = sum(v['soja_sacs'] for v in s1_bundle_per_month)
    s1_total_conc_sacs = sum(v['conc_sacs'] for v in s1_bundle_per_month)
    s1_total_soja_ca = sum(v['soja_ca'] for v in s1_bundle_per_month)
    s1_total_conc_ca = sum(v['conc_ca'] for v in s1_bundle_per_month)
    s1_n_clients = len(s1_bundle_per_month)  # client-months
    s1_n_unique_clients = len(set(ck for (ck, m) in s1_client_month.keys() if s1_client_month[(ck, m)]['soja_sacs'] > 0 and s1_client_month[(ck, m)]['conc_sacs'] > 0))

    s1_avg_soja_per_client = s1_total_soja_sacs / s1_n_clients if s1_n_clients else 0
    s1_avg_conc_per_client = s1_total_conc_sacs / s1_n_clients if s1_n_clients else 0
    s1_price_soja = s1_total_soja_ca / s1_total_soja_sacs if s1_total_soja_sacs else 0
    s1_price_conc = s1_total_conc_ca / s1_total_conc_sacs if s1_total_conc_sacs else 0
    s1_budget = (s1_total_soja_ca + s1_total_conc_ca) / s1_n_clients if s1_n_clients else 0
    s1_ratio = s1_avg_soja_per_client / s1_avg_conc_per_client if s1_avg_conc_per_client else 0

    # July metrics
    july_total_soja_sacs = sum(v['soja_sacs'] for v in july_bundle)
    july_total_conc_sacs = sum(v['conc_sacs'] for v in july_bundle)
    july_total_soja_ca = sum(v['soja_ca'] for v in july_bundle)
    july_total_conc_ca = sum(v['conc_ca'] for v in july_bundle)
    july_n_clients = len(july_bundle)

    july_avg_soja_per_client = july_total_soja_sacs / july_n_clients if july_n_clients else 0
    july_avg_conc_per_client = july_total_conc_sacs / july_n_clients if july_n_clients else 0
    july_price_soja = july_total_soja_ca / july_total_soja_sacs if july_total_soja_sacs else 0
    july_price_conc = july_total_conc_ca / july_total_conc_sacs if july_total_conc_sacs else 0
    july_budget = (july_total_soja_ca + july_total_conc_ca) / july_n_clients if july_n_clients else 0
    july_ratio = july_avg_soja_per_client / july_avg_conc_per_client if july_avg_conc_per_client else 0

    # Per agency
    by_agence_s1 = defaultdict(lambda: {'clients': 0, 'soja_sacs': 0, 'conc_sacs': 0, 'soja_ca': 0, 'conc_ca': 0, 'region': None})
    for v in s1_bundle_per_month:
        ag = v['agence']
        by_agence_s1[ag]['clients'] += 1
        by_agence_s1[ag]['soja_sacs'] += v['soja_sacs']
        by_agence_s1[ag]['conc_sacs'] += v['conc_sacs']
        by_agence_s1[ag]['soja_ca'] += v['soja_ca']
        by_agence_s1[ag]['conc_ca'] += v['conc_ca']
        by_agence_s1[ag]['region'] = v['region']

    by_agence_july = defaultdict(lambda: {'clients': 0, 'soja_sacs': 0, 'conc_sacs': 0, 'soja_ca': 0, 'conc_ca': 0, 'region': None})
    for v in july_bundle:
        ag = v['agence']
        by_agence_july[ag]['clients'] += 1
        by_agence_july[ag]['soja_sacs'] += v['soja_sacs']
        by_agence_july[ag]['conc_sacs'] += v['conc_sacs']
        by_agence_july[ag]['soja_ca'] += v['soja_ca']
        by_agence_july[ag]['conc_ca'] += v['conc_ca']
        by_agence_july[ag]['region'] = v['region']

    agence_data = []
    for ag in sorted(set(list(by_agence_s1.keys()) + list(by_agence_july.keys()))):
        s1 = by_agence_s1.get(ag, {'clients': 0, 'soja_sacs': 0, 'conc_sacs': 0, 'soja_ca': 0, 'conc_ca': 0, 'region': None})
        ju = by_agence_july.get(ag, {'clients': 0, 'soja_sacs': 0, 'conc_sacs': 0, 'soja_ca': 0, 'conc_ca': 0, 'region': None})
        s1_avg_soja = s1['soja_sacs'] / s1['clients'] if s1['clients'] else 0
        s1_avg_conc = s1['conc_sacs'] / s1['clients'] if s1['clients'] else 0
        ju_avg_soja = ju['soja_sacs'] / ju['clients'] if ju['clients'] else 0
        ju_avg_conc = ju['conc_sacs'] / ju['clients'] if ju['clients'] else 0
        s1_p_soja = s1['soja_ca'] / s1['soja_sacs'] if s1['soja_sacs'] else 0
        s1_p_conc = s1['conc_ca'] / s1['conc_sacs'] if s1['conc_sacs'] else 0
        ju_p_soja = ju['soja_ca'] / ju['soja_sacs'] if ju['soja_sacs'] else 0
        ju_p_conc = ju['conc_ca'] / ju['conc_sacs'] if ju['conc_sacs'] else 0
        s1_budget = (s1['soja_ca'] + s1['conc_ca']) / s1['clients'] if s1['clients'] else 0
        ju_budget = (ju['soja_ca'] + ju['conc_ca']) / ju['clients'] if ju['clients'] else 0
        s1_ratio = s1_avg_soja / s1_avg_conc if s1_avg_conc else 0
        ju_ratio = ju_avg_soja / ju_avg_conc if ju_avg_conc else 0
        agence_data.append({
            'agence': ag,
            'region': s1['region'] or ju['region'],
            's1_clients': s1['clients'],
            'juil_clients': ju['clients'],
            's1_price_soja': round(s1_p_soja, 0),
            'juil_price_soja': round(ju_p_soja, 0),
            's1_price_conc': round(s1_p_conc, 0),
            'juil_price_conc': round(ju_p_conc, 0),
            's1_budget': round(s1_budget, 0),
            'juil_budget': round(ju_budget, 0),
            's1_sacs_soja': round(s1_avg_soja, 1),
            'juil_sacs_soja': round(ju_avg_soja, 1),
            's1_sacs_conc': round(s1_avg_conc, 1),
            'juil_sacs_conc': round(ju_avg_conc, 1),
            's1_ratio': round(s1_ratio, 2),
            'juil_ratio': round(ju_ratio, 2),
        })

    return {
        'global': {
            's1': {
                'n_client_months': s1_n_clients,
                'n_unique_clients': s1_n_unique_clients,
                'price_soja': round(s1_price_soja, 0),
                'price_conc': round(s1_price_conc, 0),
                'budget': round(s1_budget, 0),
                'sacs_soja_per_client': round(s1_avg_soja_per_client, 1),
                'sacs_conc_per_client': round(s1_avg_conc_per_client, 1),
                'ratio': round(s1_ratio, 2),
            },
            'juil': {
                'n_clients': july_n_clients,
                'price_soja': round(july_price_soja, 0),
                'price_conc': round(july_price_conc, 0),
                'budget': round(july_budget, 0),
                'sacs_soja_per_client': round(july_avg_soja_per_client, 1),
                'sacs_conc_per_client': round(july_avg_conc_per_client, 1),
                'ratio': round(july_ratio, 2),
            },
        },
        'by_agence': agence_data,
    }


def main():
    internal = load_internal_clients()
    print(f'Internal clients excluded: {len(internal)}')

    print('Loading S1 2026 data...')
    s1_rows = load_s1_data()
    print(f'  {len(s1_rows)} Livrée rows')

    print('Loading July 2026 data (au 24/07)...')
    july_rows = load_july_data()
    print(f'  {len(july_rows)} Livrée rows')

    print(f'Temps écoulé: {TEMPS_ECOULE_PCT:.1f}% ({JULY_DAYS_LUN_SAM}/{JULY_TOTAL_DAYS_LUN_SAM} jours lun-sam)')

    # 1. Juillet MTD metrics
    print('\n=== Juillet MTD metrics (zero achat) ===')
    mtd = compute_juillet_mtd_metrics(s1_rows, july_rows, internal)
    print(f"Clients S1: {mtd['clients']['s1_total']}")
    print(f"Clients juillet MTD: {mtd['clients']['juil_mtd']}")
    print(f"  Maintenus: {mtd['clients']['maintenus']}")
    print(f"  Nouveaux: {mtd['clients']['nouveaux']}")
    print(f"  Risque élevé: {mtd['clients']['risque_eleve']} (CA {mtd['ca']['risque_eleve']/1e6:.0f} M)")
    print(f"  À surveiller: {mtd['clients']['a_surveiller']} (CA {mtd['ca']['a_surveiller']/1e6:.0f} M)")
    print(f"  Pas inquiétant: {mtd['clients']['pas_inquietant']} (CA {mtd['ca']['pas_inquietant']/1e6:.0f} M)")
    print(f"Top 10 risque:")
    for i, c in enumerate(mtd['top_risque'], 1):
        print(f"  {i}. {c['nom'][:40]:40s} | {c['agence']:12s} | {c['region']:10s} | CA S1: {c['ca_s1_m']} M | jour moy: {c['jour_moyen']}")

    with open('/home/z/my-project/scripts/juillet_mtd_31.json', 'w', encoding='utf-8') as f:
        json.dump(mtd, f, indent=2, ensure_ascii=False)
    print('Saved: juillet_mtd_31.json')

    # 2. Cross-sell
    print('\n=== Cross-sell metrics ===')
    cs = compute_cross_sell(s1_rows, july_rows, internal)
    print(f"Soja: S1 {cs['soja']['s1_total']} clients ({cs['soja']['s1_soja_conc_pct']}% avec conc) → Juil {cs['soja']['juil_soja']} ({cs['soja']['juil_soja_conc_pct']}%)")
    print(f"  Régression: {cs['soja']['regression']}, Conversion: {cs['soja']['conversion']}")
    print(f"Booster: S1 {cs['booster']['s1_total']} ({cs['booster']['s1_booster_conc_pct']}% avec conc) → Juil {cs['booster']['juil_booster']} ({cs['booster']['juil_booster_conc_pct']}%)")
    print(f"  Régression: {cs['booster']['regression']}, Conversion: {cs['booster']['conversion']}")

    with open('/home/z/my-project/scripts/cross_sell_mtd_31.json', 'w', encoding='utf-8') as f:
        json.dump(cs, f, indent=2, ensure_ascii=False)
    print('Saved: cross_sell_mtd_31.json')

    # 3. Volumes by category
    print('\n=== Volumes by category ===')
    vols = compute_volumes(s1_rows, july_rows)
    for cat in ['TOURTEAUX', 'CONCENTRES', 'PREMIX', 'INGREDIENTS']:
        s1m = vols['s1_monthly_t'].get(cat, 0)
        ju = vols['juil_mtd_t'].get(cat, 0)
        proj = vols['juil_proj_t'].get(cat, 0)
        delta_pct = (proj - s1m) / s1m * 100 if s1m else 0
        print(f"  {cat}: S1 moy {s1m:.0f} t → Juil MTD {ju:.0f} t (proj {proj:.0f} t, {delta_pct:+.0f}%)")

    with open('/home/z/my-project/scripts/volumes_juillet_31.json', 'w', encoding='utf-8') as f:
        json.dump(vols, f, indent=2, ensure_ascii=False)
    print('Saved: volumes_juillet_31.json')

    # 4. Bundle metrics
    print('\n=== Bundle metrics ===')
    bm = compute_bundle_metrics(s1_rows, july_rows, internal)
    g_s1 = bm['global']['s1']
    g_ju = bm['global']['juil']
    print(f"Global S1: {g_s1['n_unique_clients']} clients bundle, prix soja {g_s1['price_soja']:,} FCFA/sac, prix conc {g_s1['price_conc']:,}, ratio {g_s1['ratio']}:1")
    print(f"Global Juil: {g_ju['n_clients']} clients bundle, prix soja {g_ju['price_soja']:,} FCFA/sac, prix conc {g_ju['price_conc']:,}, ratio {g_ju['ratio']}:1")
    print(f"  Budget: {g_s1['budget']:,} → {g_ju['budget']:,} ({(g_ju['budget']/g_s1['budget']-1)*100:+.0f}%)")
    print(f"  Sacs soja/client: {g_s1['sacs_soja_per_client']} → {g_ju['sacs_soja_per_client']} ({(g_ju['sacs_soja_per_client']/g_s1['sacs_soja_per_client']-1)*100:+.0f}%)")
    print(f"  Sacs conc/client: {g_s1['sacs_conc_per_client']} → {g_ju['sacs_conc_per_client']} ({(g_ju['sacs_conc_per_client']/g_s1['sacs_conc_per_client']-1)*100:+.0f}%)")
    print('\nPar agence:')
    for ag in sorted(bm['by_agence'], key=lambda x: -x['juil_ratio']):
        if ag['juil_clients'] == 0:
            continue
        print(f"  {ag['agence']:15s} {ag['region']:10s} | S1 {ag['s1_clients']:3d}→Juil {ag['juil_clients']:3d} | ratio S1 {ag['s1_ratio']:.1f}→Juil {ag['juil_ratio']:.1f} | prix soja {ag['s1_price_soja']:>5.0f}→{ag['juil_price_soja']:>5.0f}")

    with open('/home/z/my-project/scripts/bundle_metrics_31.json', 'w', encoding='utf-8') as f:
        json.dump(bm, f, indent=2, ensure_ascii=False)
    print('Saved: bundle_metrics_31.json')


if __name__ == '__main__':
    main()
