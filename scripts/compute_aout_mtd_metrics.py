"""Compute August MTD metrics (au 14/08/2026) for the zero-achat and bundle PDFs."""
import openpyxl
from collections import defaultdict, Counter
from datetime import datetime, date, timedelta
import json
import re

# Sources
AOUT_SRC = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (18).xlsx"
JUIN_SRC = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (9).xlsx"  # juillet complet (for S1 + juillet comparison)
JUILLET_FULL_SRC = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (9).xlsx"

# Product refs
SOJA_REFS = {'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25}
CONC_REFS = {
    'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
    'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50
}
MAIS_REFS = {'M1051': 50, 'M1052': 50}
# Ingredient refs (excluding Mais)
INGREDIENT_REFS = {
    'B100': 25, 'E101': 25, 'I105': 25, 'B1001': 1, 'B1003': 5, 'B1004': 25,
    'E1011': 1, 'E1013': 5, 'E1014': 25,
    'I1051': 1, 'I1053': 5, 'I1054': 25,
}
PREMIX_REFS = {'PX101': 25, 'PX102': 25, 'PX103': 25, 'PX104': 25, 'PX105': 25}

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

# Internal clients to exclude (SPC/PDC/Comptoir)
INTERNAL_CLIENT_PATTERNS = ['SPC', 'PDC', 'COMPTOIR', 'EMANA']

# Objectives (monthly targets) — August (CORRECTION: CONCENTRES = 1534 t, somme des objectifs agence)
OBJ = {
    'TOURTEAUX': 3850,  # t
    'CONCENTRES': 1534,  # t (objectif réel août, somme des objectifs agence)
    'MAIS': 130,  # t
    'INGREDIENTS': 320,  # t
}

# Stock BEKOKO (au 08/08/2026)
STOCK_BEKOKO = {'50kg': 80384, '1kg': 2245, '5kg': 4, '25kg': 1}
SPC_ALLOCATION = 9200
STOCK_DATE = '08/08/2026'
PROD_CONC = {'min': 4500, 'moy': 4900, 'max': 5300}


def detect_cols(ws):
    header = None
    for row in ws.iter_rows(min_row=2, max_row=2, values_only=True):
        header = row
        break
    indices = {}
    for i, h in enumerate(header):
        if h == 'État':
            indices['etat'] = i
        elif h == 'agence':
            indices['agence'] = i
    if 'etat' not in indices:
        indices['etat'] = 13 if len(header) <= 16 else 15
    if 'agence' not in indices:
        indices['agence'] = 15 if len(header) <= 16 else 17
    return indices


def load_livree(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['Sheet 1']
    col_idx = detect_cols(ws)
    rows = list(ws.iter_rows(min_row=3, values_only=True))
    out = []
    for r in rows:
        if not r or len(r) < max(col_idx['etat'], col_idx['agence']) + 1:
            continue
        if r[0] == 'Total':
            continue
        if r[col_idx['etat']] != 'Livrée':
            continue
        out.append(r)
    return out, col_idx


def parse_date(d):
    if isinstance(d, datetime):
        return d.strftime('%d/%m/%Y')
    if isinstance(d, str):
        return d[:10]
    return str(d)[:10]


def is_internal_client(client_str):
    if not client_str:
        return False
    s = str(client_str).upper()
    for p in INTERNAL_CLIENT_PATTERNS:
        if p in s:
            return True
    return False


# === Compute August metrics ===
print("Loading August extraction (18).xlsx...")
rows_aout, col_idx_aout = load_livree(AOUT_SRC)
print(f"  {len(rows_aout)} Livree rows")

# Filter only complete days (exclude 18/08 - matinal extraction with 5 rows)
aout_complete = []
for r in rows_aout:
    d = parse_date(r[6])
    if d == '18/08/2026':  # matinal - only 5 Livree rows
        continue
    aout_complete.append(r)
print(f"  {len(aout_complete)} Livree rows (excluding 18/08 matinal)")

# Days elapsed (lun-sam) — Aug 1-17
days_elapsed = 0
for d in range(1, 18):  # 1-17
    dt = date(2026, 8, d)
    if dt.weekday() < 6:
        days_elapsed += 1
print(f"  Days elapsed (lun-sam): {days_elapsed}")

# Total days in August (lun-sam)
total_days_aug = 0
for d in range(1, 32):
    dt = date(2026, 8, d)
    if dt.weekday() < 6:
        total_days_aug += 1
print(f"  Total days in August (lun-sam): {total_days_aug}")
pct_elapsed = days_elapsed / total_days_aug * 100

# === Volume by category (August MTD) ===
vol = defaultdict(float)  # kg
vol_sacs = defaultdict(float)
vol_by_agence = defaultdict(lambda: defaultdict(float))  # agence -> cat -> kg

# Filter out internal clients (SPC/PDC/Comptoir)
external_rows = []
internal_count = 0
for r in aout_complete:
    client = r[5] if len(r) > 5 else ''
    if is_internal_client(client):
        internal_count += 1
        continue
    external_rows.append(r)
print(f"  External rows (after excluding internal clients): {len(external_rows)} (excluded: {internal_count})")

for r in external_rows:
    ref = r[0]
    qte = r[2] or 0
    date_str = parse_date(r[6])
    if '/08/2026' not in date_str:
        continue
    agence_raw = r[col_idx_aout['agence']] or ''
    agence_short, region = AGENCE_MAP.get(agence_raw, (agence_raw, '?'))
    
    if ref in SOJA_REFS:
        cat = 'TOURTEAUX'
        kg = qte * SOJA_REFS[ref]
    elif ref in CONC_REFS:
        cat = 'CONCENTRES'
        kg = qte * CONC_REFS[ref]
    elif ref in MAIS_REFS:
        cat = 'MAIS'
        kg = qte * MAIS_REFS[ref]
    elif ref in INGREDIENT_REFS:
        cat = 'INGREDIENTS'
        kg = qte * INGREDIENT_REFS[ref]
    elif ref in PREMIX_REFS:
        cat = 'PREMIX'
        kg = qte * PREMIX_REFS[ref]
    else:
        continue
    
    vol[cat] += kg
    vol_sacs[cat] += kg / 50
    vol_by_agence[agence_short][cat] += kg

print(f"\n=== VOLUMES AOÛT MTD (au 14/08, {days_elapsed}j lun-sam) ===")
for cat in ['TOURTEAUX', 'CONCENTRES', 'MAIS', 'INGREDIENTS', 'PREMIX']:
    t = vol[cat] / 1000
    moy = t / days_elapsed
    proj = moy * total_days_aug
    obj = OBJ.get(cat, 0)
    pct = proj / obj * 100 if obj > 0 else 0
    print(f"  {cat}: {t:.1f} t (moy {moy:.1f} t/j, proj {proj:.0f} t, obj {obj} t → {pct:.0f}%)")

# === CONCENTRES detail by agence ===
print(f"\n=== CONCENTRÉS par agence (au 14/08) ===")
conc_by_agence = []
for ag, cats in vol_by_agence.items():
    conc_t = cats.get('CONCENTRES', 0) / 1000
    soja_t = cats.get('TOURTEAUX', 0) / 1000
    if conc_t > 0 or soja_t > 0:
        conc_by_agence.append((ag, conc_t, soja_t, conc_t / days_elapsed if days_elapsed else 0))
conc_by_agence.sort(key=lambda x: -x[1])
print(f"{'Agence':15} {'Conc (t)':>10} {'Soja (t)':>10} {'Conc/j (t)':>12} {'Proj (t)':>10}")
for ag, conc_t, soja_t, moy in conc_by_agence:
    proj = moy * total_days_aug
    print(f"{ag:15} {conc_t:>10.1f} {soja_t:>10.1f} {moy:>12.1f} {proj:>10.0f}")

# === Bundle ratio (soja:conc) ===
total_soja_sacs = vol_sacs['TOURTEAUX']
total_conc_sacs = vol_sacs['CONCENTRES']
ratio = total_soja_sacs / total_conc_sacs if total_conc_sacs > 0 else 0
print(f"\n=== RATIO BUNDLE AOÛT ===")
print(f"  Total soja: {total_soja_sacs:.0f} sacs ({vol['TOURTEAUX']/1000:.0f} t)")
print(f"  Total conc: {total_conc_sacs:.0f} sacs ({vol['CONCENTRES']/1000:.0f} t)")
print(f"  Ratio soja:conc = {ratio:.1f}:1")

# === Soja-only commands ===
cmds = defaultdict(lambda: {'soja_kg': 0, 'conc_kg': 0, 'agence': None, 'client': None, 'date': None, 'cmd_ref': None, 'soja_sacs_50': 0, 'conc_sacs_50': 0})
for r in external_rows:
    ref = r[0]
    qte = r[2] or 0
    cmd = r[3]
    date_str = parse_date(r[6])
    if '/08/2026' not in date_str:
        continue
    agence_raw = r[col_idx_aout['agence']] or ''
    agence_short, region = AGENCE_MAP.get(agence_raw, (agence_raw, '?'))
    if ref in SOJA_REFS:
        cmds[cmd]['soja_kg'] += qte * SOJA_REFS[ref]
        cmds[cmd]['soja_sacs_50'] += qte * SOJA_REFS[ref] / 50
        cmds[cmd]['agence'] = agence_short
        cmds[cmd]['region'] = region
        cmds[cmd]['client'] = r[5]
        cmds[cmd]['date'] = date_str
        cmds[cmd]['cmd_ref'] = cmd
    if ref in CONC_REFS:
        cmds[cmd]['conc_kg'] += qte * CONC_REFS[ref]
        cmds[cmd]['conc_sacs_50'] += qte * CONC_REFS[ref] / 50
        cmds[cmd]['agence'] = agence_short
        cmds[cmd]['region'] = region
        cmds[cmd]['client'] = r[5]
        cmds[cmd]['date'] = date_str
        cmds[cmd]['cmd_ref'] = cmd

cmds_soja = {k: v for k, v in cmds.items() if v['soja_kg'] > 0}
cmds_bundle = {k: v for k, v in cmds.items() if v['soja_kg'] > 0 and v['conc_kg'] > 0}
cmds_soja_only = {k: v for k, v in cmds.items() if v['soja_kg'] > 0 and v['conc_kg'] == 0}
cmds_conc_only = {k: v for k, v in cmds.items() if v['conc_kg'] > 0 and v['soja_kg'] == 0}

print(f"\n=== COMMANDES AOÛT (au 14/08) ===")
print(f"  Total cmds soja: {len(cmds_soja)}")
print(f"  Bundle (soja+conc): {len(cmds_bundle)}")
print(f"  Soja-only: {len(cmds_soja_only)}")
print(f"  Conc-only: {len(cmds_conc_only)}")
print(f"  % Bundle (sur cmds soja): {len(cmds_bundle)/len(cmds_soja)*100:.0f}%" if cmds_soja else "n/a")

# Distribution of ratios
ratios = []
for c in cmds_bundle.values():
    if c['conc_sacs_50'] > 0:
        ratios.append(c['soja_sacs_50'] / c['conc_sacs_50'])
dist = {
    '<= 3:1': sum(1 for r in ratios if r <= 3),
    '3-5:1': sum(1 for r in ratios if 3 < r <= 5),
    '5-10:1': sum(1 for r in ratios if 5 < r <= 10),
    '10-20:1': sum(1 for r in ratios if 10 < r <= 20),
    '> 20:1': sum(1 for r in ratios if r > 20),
}
total_r = len(ratios)
print(f"\n=== DISTRIBUTION RATIOS BUNDLE ===")
for k, v in dist.items():
    pct = v / total_r * 100 if total_r else 0
    print(f"  {k}: {v} ({pct:.0f}%)")

# === STOCK ===
stock_50 = STOCK_BEKOKO['50kg']
stock_1 = STOCK_BEKOKO['1kg']
stock_5 = STOCK_BEKOKO['5kg']
stock_25 = STOCK_BEKOKO['25kg']
stock_eq_50_brut = stock_50 + stock_1 * 1/50 + stock_5 * 5/50 + stock_25 * 25/50
stock_t_brut = stock_eq_50_brut * 50 / 1000
stock_eq_50_net = stock_eq_50_brut - SPC_ALLOCATION
stock_t_net = stock_eq_50_net * 50 / 1000

# Vente soja / jour
vente_soja_sacs_jour = vol_sacs['TOURTEAUX'] / days_elapsed
vente_soja_sacs_sem = vente_soja_sacs_jour * 6
conso_moy = vente_soja_sacs_sem + PROD_CONC['moy']
jours_stock_moy = stock_eq_50_net / conso_moy * 7
rupture_date = date(2026, 8, 17) + timedelta(days=int(jours_stock_moy))

print(f"\n=== STOCK SOJA (BEKOKO au {STOCK_DATE}) ===")
print(f"  Stock brut: {stock_eq_50_brut:.0f} sacs ({stock_t_brut:.0f} t)")
print(f"  Stock net (excl SPC): {stock_eq_50_net:.0f} sacs ({stock_t_net:.0f} t)")
print(f"  Vente soja directe: {vente_soja_sacs_sem:.0f} sacs/sem ({vente_soja_sacs_jour:.0f} sacs/j)")
print(f"  Conso totale moy (vente + prod conc): {conso_moy:.0f} sacs/sem")
print(f"  Jours de stock: {int(jours_stock_moy)}")
print(f"  Date rupture probable: {rupture_date.strftime('%d/%m/%Y')}")

# === Zero-achat analysis ===
# Compare S1 (Jan-June) clients vs August clients
print(f"\n=== ZERO ACHAT AOÛT ===")
# Use juin extraction as S1 source
rows_s1, col_idx_s1 = load_livree(JUIN_SRC)
s1_clients = set()
for r in rows_s1:
    if not r or not r[5]:
        continue
    if is_internal_client(r[5]):
        continue
    s1_clients.add(r[5])

aout_clients = set()
for r in external_rows:
    if '/08/2026' not in parse_date(r[6]):
        continue
    if not r[5]:
        continue
    aout_clients.add(r[5])

print(f"  Clients S1 (Jan-Juil): {len(s1_clients)}")
print(f"  Clients Août MTD: {len(aout_clients)}")
churned = s1_clients - aout_clients
new_aout = aout_clients - s1_clients
print(f"  Churned (S1 sans achat août): {len(churned)}")
print(f"  Nouveaux août (hors S1): {len(new_aout)}")

# Cross-sell analysis: clients qui achetaient soja + conc en S1
s1_soja_clients = set()
s1_conc_clients = set()
for r in rows_s1:
    if not r or not r[5] or is_internal_client(r[5]):
        continue
    ref = r[0]
    if ref in SOJA_REFS:
        s1_soja_clients.add(r[5])
    elif ref in CONC_REFS:
        s1_conc_clients.add(r[5])

aout_soja_clients = set()
aout_conc_clients = set()
for r in external_rows:
    if '/08/2026' not in parse_date(r[6]):
        continue
    if not r[5]:
        continue
    ref = r[0]
    if ref in SOJA_REFS:
        aout_soja_clients.add(r[5])
    elif ref in CONC_REFS:
        aout_conc_clients.add(r[5])

# Cross-sell: S1 clients who bought BOTH soja+conc
s1_bundle = s1_soja_clients & s1_conc_clients
aout_bundle = aout_soja_clients & aout_conc_clients
print(f"\n  S1 bundle (soja+conc): {len(s1_bundle)}")
print(f"  Août bundle (soja+conc): {len(aout_bundle)}")

# Cross-sell soja->conc: S1 soja clients, did they buy conc in août?
s1_soja_conc_aout = s1_soja_clients & aout_conc_clients
s1_soja_no_conc_aout = s1_soja_clients - aout_conc_clients
print(f"\n  S1 soja clients: {len(s1_soja_clients)}")
print(f"  Dont ont aussi acheté conc en août: {len(s1_soja_conc_aout)} ({len(s1_soja_conc_aout)/len(s1_soja_clients)*100:.0f}%)")
print(f"  Dont n'ont pas acheté conc en août: {len(s1_soja_no_conc_aout)}")

# Save summary
summary = {
    'update_date': '17/08/2026',
    'days_elapsed': days_elapsed,
    'total_days_aug': total_days_aug,
    'pct_elapsed': round(pct_elapsed, 1),
    'volumes': {k: {'t': round(v/1000, 1), 'moy_t_j': round(v/1000/days_elapsed, 1), 
                    'proj_t': round(v/1000/days_elapsed*total_days_aug, 0),
                    'obj_t': OBJ.get(k, 0),
                    'pct_obj': round(v/1000/days_elapsed*total_days_aug/OBJ.get(k, 1)*100, 0)} 
                for k, v in vol.items()},
    'bundle': {
        'total_soja_sacs': round(total_soja_sacs, 0),
        'total_conc_sacs': round(total_conc_sacs, 0),
        'ratio': round(ratio, 1),
        'cmds_soja': len(cmds_soja),
        'cmds_bundle': len(cmds_bundle),
        'cmds_soja_only': len(cmds_soja_only),
        'cmds_conc_only': len(cmds_conc_only),
        'pct_bundle': round(len(cmds_bundle)/len(cmds_soja)*100, 0) if cmds_soja else 0,
        'dist': dist,
    },
    'stock': {
        'date': STOCK_DATE,
        'brut_sacs': round(stock_eq_50_brut, 0),
        'brut_t': round(stock_t_brut, 0),
        'net_sacs': round(stock_eq_50_net, 0),
        'net_t': round(stock_t_net, 0),
        'spc_exclu': SPC_ALLOCATION,
        'vente_sacs_jour': round(vente_soja_sacs_jour, 0),
        'vente_sacs_sem': round(vente_soja_sacs_sem, 0),
        'conso_moy_sacs_sem': round(conso_moy, 0),
        'jours_stock': int(jours_stock_moy),
        'rupture_date': rupture_date.strftime('%d/%m/%Y'),
    },
    'zero_achat': {
        's1_clients': len(s1_clients),
        'aout_clients': len(aout_clients),
        'churned': len(churned),
        'new_aout': len(new_aout),
    },
    'cross_sell': {
        's1_soja': len(s1_soja_clients),
        's1_conc': len(s1_conc_clients),
        'aout_soja': len(aout_soja_clients),
        'aout_conc': len(aout_conc_clients),
        's1_bundle': len(s1_bundle),
        'aout_bundle': len(aout_bundle),
        'soja_to_conc_aout': len(s1_soja_conc_aout),
        'soja_no_conc_aout': len(s1_soja_no_conc_aout),
        'pct_soja_with_conc_aout': round(len(s1_soja_conc_aout)/len(s1_soja_clients)*100, 0) if s1_soja_clients else 0,
    },
    'conc_by_agence': [{'agence': ag, 'conc_t': round(c, 1), 'soja_t': round(s, 1), 
                        'moy_t_j': round(c/days_elapsed, 1), 'proj_t': round(c/days_elapsed*total_days_aug, 0)}
                       for ag, c, s in [(ag, cats.get('CONCENTRES', 0), cats.get('TOURTEAUX', 0)) 
                                          for ag, cats in vol_by_agence.items()]
                       if c > 0 or s > 0]
}

with open('/home/z/my-project/scripts/aout_mtd_18.json', 'w', encoding='utf-8') as f:
    json.dump(summary, f, indent=2, ensure_ascii=False, default=str)
print(f"\nSaved summary: /home/z/my-project/scripts/aout_mtd_18.json")
