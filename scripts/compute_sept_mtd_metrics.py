"""Compute September MTD metrics from NJS GROUP ERP extraction (33).xlsx.
Based on compute_aout_mtd_metrics.py — adapted for September 2026.

Outputs:
- /home/z/my-project/scripts/sept_mtd_01.json  (latest metrics)
"""
import openpyxl
from collections import defaultdict, Counter
from datetime import datetime, date, timedelta
import json
import re

# Source (latest extraction: (36).xlsx as of 11/09/2026)
SEPT_SRC = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (6) (1).xlsx"

# Product refs
SOJA_REFS = {'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25}
CONC_REFS = {
    'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
    'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50
}
MAIS_REFS = {'M1051': 50, 'M1052': 50}
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
    'AGENCE MAROUA': ('Maroua', 'Centre'),
    'SPC BAF-CHEFFERIE': ('Baf-Chefferie', 'Ouest'),
    'PDC Emana': ('Emana', 'Centre'),
    'SPC-NDERE': ('Ndere', 'Centre'),
    'SPC-DSCHANG': ('Dschang', 'Ouest'),
    'SPC BUEA': ('Buea-SPC', 'Littoral'),
    'SPC-YASSA': ('Yassa', 'Littoral'),
}

# Internal clients to exclude (SPC/PDC/Comptoir)
INTERNAL_CLIENT_PATTERNS = ['SPC', 'PDC', 'EMANA']  # COMPTOIR inclus (ventes au comptoir à compter)

# Objectives September (monthly targets) — same as Aug for first estimate
OBJ = {
    'TOURTEAUX': 3850,  # t
    'CONCENTRES': 1534,  # t (objectif réel)
    'MAIS': 130,  # t
    'INGREDIENTS': 320,  # t
}

# Stock BEKOKO (au 08/08/2026, à mettre à jour si nouveau stock)
STOCK_BEKOKO = {'50kg': 80384, '1kg': 2245, '5kg': 4, '25kg': 1}
SPC_ALLOCATION = 9200
STOCK_DATE = '08/08/2026'  # À mettre à jour
PROD_CONC = {'min': 4500, 'moy': 4900, 'max': 5300}


def detect_cols(ws):
    """Detect column indices based on header row."""
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
    # Fallback to standard positions
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


# === Compute September metrics ===
print(f"Loading September extraction: {SEPT_SRC.split('/')[-1]}...")
rows_sept, col_idx_sept = load_livree(SEPT_SRC)
print(f"  {len(rows_sept)} Livrée rows total")

# Identify latest date in the extraction (to compute "days elapsed")
latest_date = None
for r in rows_sept:
    date_str = parse_date(r[6])
    if '/09/2026' in date_str:
        try:
            d = datetime.strptime(date_str, '%d/%m/%Y').date()
            if latest_date is None or d > latest_date:
                latest_date = d
        except Exception:
            pass
if latest_date is None:
    latest_date = date(2026, 9, 1)
print(f"  Latest Livrée date in extraction: {latest_date.strftime('%d/%m/%Y')}")

# Days elapsed (lun-sam) — Sep 1 to latest_date inclusive
days_elapsed = 0
for d_offset in range((latest_date - date(2026, 9, 1)).days + 1):
    dt = date(2026, 9, 1) + timedelta(days=d_offset)
    if dt.weekday() < 6:
        days_elapsed += 1
print(f"  Days elapsed (lun-sam, Sep 1 to {latest_date.strftime('%d/%m/%Y')}): {days_elapsed}")

# Total days in September (lun-sam)
total_days_sep = 0
for d in range(1, 31):
    dt = date(2026, 9, d)
    if dt.weekday() < 6:
        total_days_sep += 1
print(f"  Total days in September (lun-sam): {total_days_sep}")
pct_elapsed = days_elapsed / total_days_sep * 100 if total_days_sep > 0 else 0

# === Volume by category (Sept MTD) ===
vol = defaultdict(float)
vol_sacs = defaultdict(float)
vol_by_agence = defaultdict(lambda: defaultdict(float))

# Filter out internal clients
external_rows = []
internal_count = 0
for r in rows_sept:
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
    if '/09/2026' not in date_str:
        continue
    agence_raw = r[col_idx_sept['agence']] or ''
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

print(f"\n=== VOLUMES SEPTEMBRE MTD (au {latest_date.strftime('%d/%m/%Y')}, {days_elapsed}j lun-sam) ===")
for cat in ['TOURTEAUX', 'CONCENTRES', 'MAIS', 'INGREDIENTS', 'PREMIX']:
    t = vol[cat] / 1000
    moy_j = t / days_elapsed if days_elapsed > 0 else 0
    proj = moy_j * total_days_sep
    obj = OBJ.get(cat, 0)
    pct = (proj / obj * 100) if obj > 0 else 0
    print(f"  {cat:12s}: {t:6.1f} t ({moy_j:.1f} t/j, projection {proj:.0f} t vs obj {obj} t = {pct:.0f}%)")

# === Bundle soja-concentrés (Sept MTD) ===
print("\n=== BUNDLE SEPTEMBRE MTD ===")
cmds_soja = set()
cmds_conc = set()
cmds_bundle = set()
cmds_soja_only = set()
cmds_conc_only = set()

total_soja_sacs = 0
total_conc_sacs = 0

dist = Counter()

for r in external_rows:
    ref = r[0]
    qte = r[2] or 0
    date_str = parse_date(r[6])
    if '/09/2026' not in date_str:
        continue
    cmd_ref = r[3] if len(r) > 3 else None  # Réf. commande client
    if not cmd_ref:
        continue
    
    is_soja = ref in SOJA_REFS
    is_conc = ref in CONC_REFS
    
    if is_soja:
        cmds_soja.add(cmd_ref)
        kg = qte * SOJA_REFS[ref]
        total_soja_sacs += kg / 50
    elif is_conc:
        cmds_conc.add(cmd_ref)
        kg = qte * CONC_REFS[ref]
        total_conc_sacs += kg / 50

# Bundle = cmds that have both soja AND conc
# Need to check each command's products
cmd_to_prods = defaultdict(set)
cmd_to_soja_sacs = defaultdict(float)
cmd_to_conc_sacs = defaultdict(float)
for r in external_rows:
    ref = r[0]
    qte = r[2] or 0
    date_str = parse_date(r[6])
    if '/09/2026' not in date_str:
        continue
    cmd_ref = r[3] if len(r) > 3 else None
    if not cmd_ref:
        continue
    if ref in SOJA_REFS:
        cmd_to_prods[cmd_ref].add('soja')
        cmd_to_soja_sacs[cmd_ref] += qte * SOJA_REFS[ref] / 50
    elif ref in CONC_REFS:
        cmd_to_prods[cmd_ref].add('conc')
        cmd_to_conc_sacs[cmd_ref] += qte * CONC_REFS[ref] / 50

cmds_bundle = {c for c, p in cmd_to_prods.items() if 'soja' in p and 'conc' in p}
cmds_soja_only = {c for c, p in cmd_to_prods.items() if 'soja' in p and 'conc' not in p}
cmds_conc_only = {c for c, p in cmd_to_prods.items() if 'conc' in p and 'soja' not in p}
cmds_soja = {c for c, p in cmd_to_prods.items() if 'soja' in p}
cmds_conc = {c for c, p in cmd_to_prods.items() if 'conc' in p}

# Distribution of bundle ratios (soja:conc)
for cmd_ref, prods in cmd_to_prods.items():
    if 'soja' in prods and 'conc' in prods:
        s = cmd_to_soja_sacs[cmd_ref]
        c = cmd_to_conc_sacs[cmd_ref]
        if c > 0:
            ratio = s / c
            if ratio <= 3:
                dist['<= 3:1'] += 1
            elif ratio <= 5:
                dist['3-5:1'] += 1
            elif ratio <= 10:
                dist['5-10:1'] += 1
            elif ratio <= 20:
                dist['10-20:1'] += 1
            else:
                dist['> 20:1'] += 1

ratio_global = total_soja_sacs / total_conc_sacs if total_conc_sacs > 0 else 0
pct_bundle = len(cmds_bundle) / len(cmds_soja) * 100 if cmds_soja else 0

print(f"  Total soja sacs: {total_soja_sacs:.0f}")
print(f"  Total conc sacs: {total_conc_sacs:.0f}")
print(f"  Ratio global soja/conc: {ratio_global:.1f}:1")
print(f"  Commandes soja: {len(cmds_soja)}")
print(f"  Commandes bundle (soja+conc): {len(cmds_bundle)} ({pct_bundle:.0f}% des cmds soja)")
print(f"  Commandes soja-only: {len(cmds_soja_only)}")
print(f"  Commandes conc-only: {len(cmds_conc_only)}")
print(f"  Distribution ratios bundle:")
for k in ['<= 3:1', '3-5:1', '5-10:1', '10-20:1', '> 20:1']:
    print(f"    {k}: {dist[k]}")

# === Stock (placeholder, will need actual stock data) ===
stock_eq_50_brut = sum(STOCK_BEKOKO[k] * (1 if k == '50kg' else (1/50 if k == '1kg' else 5/50 if k == '5kg' else 25/50)) for k in STOCK_BEKOKO)
stock_t_brut = stock_eq_50_brut * 50 / 1000
stock_eq_50_net = stock_eq_50_brut - SPC_ALLOCATION
stock_t_net = stock_eq_50_net * 50 / 1000

# Vente soja sacs/jour (sept MTD)
vente_soja_sacs_jour = total_soja_sacs / days_elapsed if days_elapsed > 0 else 0
vente_soja_sacs_sem = vente_soja_sacs_jour * 6
conso_moy = max(vente_soja_sacs_sem, 22383)  # keep max with S1 conso

if vente_soja_sacs_jour > 0:
    jours_stock = int(stock_eq_50_net / vente_soja_sacs_jour)
    rupture_date = latest_date + timedelta(days=jours_stock)
else:
    jours_stock = 0
    rupture_date = latest_date

print(f"\n=== STOCK (au {STOCK_DATE}) ===")
print(f"  Stock brut: {stock_eq_50_brut:.0f} sacs ({stock_t_brut:.0f} t)")
print(f"  Stock net (hors SPC): {stock_eq_50_net:.0f} sacs ({stock_t_net:.0f} t)")
print(f"  Vente soja sept: {vente_soja_sacs_jour:.0f} sacs/jour, {vente_soja_sacs_sem:.0f} sacs/sem")
print(f"  Jours de stock: {jours_stock}j — rupture probable {rupture_date.strftime('%d/%m/%Y')}")

# === Zero-achat Sept vs S1 2026 ===
# Need S1 baseline (Jan-Jun 2026)
# Use historical dataset for S1 clients
import pandas as pd
hist_df = pd.read_csv("/home/z/my-project/scripts/dataset_2023_2026.csv", parse_dates=['date'], low_memory=False)
s1_hist = hist_df[(hist_df['date'].dt.year == 2026) & (hist_df['date'].dt.month <= 6)]
print(f"\n=== ZERO-ACHAT ANALYSE ===")
print(f"  S1 2026 records: {len(s1_hist)}")

# Get S1 unique clients
s1_clients = set(s1_hist['tiers'].dropna().unique()) if 'tiers' in s1_hist.columns else set()
if not s1_clients and 'client' in s1_hist.columns:
    s1_clients = set(s1_hist['client'].dropna().unique())

# September clients (from external rows)
sept_clients = set()
for r in external_rows:
    date_str = parse_date(r[6])
    if '/09/2026' not in date_str:
        continue
    client = r[5] if len(r) > 5 else None
    if client:
        sept_clients.add(client)

churned = s1_clients - sept_clients
new_sept = sept_clients - s1_clients
print(f"  S1 2026 clients: {len(s1_clients)}")
print(f"  Sept MTD clients: {len(sept_clients)}")
print(f"  Churned (S1 mais pas sept): {len(churned)}")
print(f"  New (nouveaux sept): {len(new_sept)}")

# === Cross-sell analysis ===
# S1 soja and conc clients
s1_soja_clients = set()
s1_conc_clients = set()
for _, row in s1_hist.iterrows():
    ref = row.get('ref', None)
    client = row.get('tiers', None) or row.get('client', None)
    if not client:
        continue
    if ref in SOJA_REFS:
        s1_soja_clients.add(client)
    elif ref in CONC_REFS:
        s1_conc_clients.add(client)

sept_soja_clients = set()
sept_conc_clients = set()
for r in external_rows:
    date_str = parse_date(r[6])
    if '/09/2026' not in date_str:
        continue
    if not r[5]:
        continue
    ref = r[0]
    if ref in SOJA_REFS:
        sept_soja_clients.add(r[5])
    elif ref in CONC_REFS:
        sept_conc_clients.add(r[5])

s1_bundle = s1_soja_clients & s1_conc_clients
sept_bundle = sept_soja_clients & sept_conc_clients

s1_soja_conc_sept = s1_soja_clients & sept_conc_clients
s1_soja_no_conc_sept = s1_soja_clients - sept_conc_clients

print(f"\n  S1 bundle (soja+conc): {len(s1_bundle)}")
print(f"  Sept bundle (soja+conc): {len(sept_bundle)}")
print(f"  S1 soja clients: {len(s1_soja_clients)}")
print(f"  Dont ont aussi acheté conc en sept: {len(s1_soja_conc_sept)} ({len(s1_soja_conc_sept)/len(s1_soja_clients)*100 if s1_soja_clients else 0:.0f}%)")
print(f"  Dont n'ont pas acheté conc en sept: {len(s1_soja_no_conc_sept)}")

# === CONCENTRES par agence (Sept MTD) ===
print(f"\n=== CONCENTRES PAR AGENCE (Sept MTD) ===")
conc_by_ag = []
for ag, cats in vol_by_agence.items():
    c = cats.get('CONCENTRES', 0)
    s = cats.get('TOURTEAUX', 0)
    if c > 0 or s > 0:
        moy_j = c/1000/days_elapsed if days_elapsed > 0 else 0
        proj = moy_j * total_days_sep
        conc_by_ag.append({'agence': ag, 'conc_t': round(c/1000, 1), 'soja_t': round(s/1000, 1),
                          'moy_t_j': round(moy_j, 2), 'proj_t': round(proj, 0)})
        print(f"  {ag:20s}: conc {c/1000:.1f} t, soja {s/1000:.1f} t (proj conc {proj:.0f} t)")

# Save summary
summary = {
    'update_date': latest_date.strftime('%d/%m/%Y'),
    'extraction_file': 'NJS GROUP ERP - Lignes de commandes + multicompany (33).xlsx',
    'days_elapsed': days_elapsed,
    'total_days_sep': total_days_sep,
    'pct_elapsed': round(pct_elapsed, 1),
    'volumes': {k: {'t': round(v/1000, 1), 'moy_t_j': round(v/1000/days_elapsed, 2) if days_elapsed > 0 else 0,
                    'proj_t': round(v/1000/days_elapsed*total_days_sep, 0) if days_elapsed > 0 else 0,
                    'obj_t': OBJ.get(k, 0),
                    'pct_obj': round(v/1000/days_elapsed*total_days_sep/OBJ.get(k, 1)*100, 0) if days_elapsed > 0 else 0}
                for k, v in vol.items()},
    'bundle': {
        'total_soja_sacs': round(total_soja_sacs, 0),
        'total_conc_sacs': round(total_conc_sacs, 0),
        'ratio': round(ratio_global, 1),
        'cmds_soja': len(cmds_soja),
        'cmds_bundle': len(cmds_bundle),
        'cmds_soja_only': len(cmds_soja_only),
        'cmds_conc_only': len(cmds_conc_only),
        'pct_bundle': round(pct_bundle, 0),
        'dist': dict(dist),
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
        'jours_stock': jours_stock,
        'rupture_date': rupture_date.strftime('%d/%m/%Y'),
    },
    'zero_achat': {
        's1_clients': len(s1_clients),
        'sept_clients': len(sept_clients),
        'churned': len(churned),
        'new_sept': len(new_sept),
    },
    'cross_sell': {
        's1_soja': len(s1_soja_clients),
        's1_conc': len(s1_conc_clients),
        'sept_soja': len(sept_soja_clients),
        'sept_conc': len(sept_conc_clients),
        's1_bundle': len(s1_bundle),
        'sept_bundle': len(sept_bundle),
        'soja_to_conc_sept': len(s1_soja_conc_sept),
        'soja_no_conc_sept': len(s1_soja_no_conc_sept),
        'pct_soja_with_conc_sept': round(len(s1_soja_conc_sept)/len(s1_soja_clients)*100, 0) if s1_soja_clients else 0,
    },
    'conc_by_agence': sorted(conc_by_ag, key=lambda x: -x['conc_t']),
}

OUT = '/home/z/my-project/scripts/sept_mtd_07.json'

# Preserve stock info from previous sept_mtd_01.json (manually updated 07/09 with new stock central)
import os
prev_stock = None
prev_path = '/home/z/my-project/scripts/sept_mtd_01.json'
if os.path.exists(prev_path):
    prev_data = json.load(open(prev_path))
    if 'stock' in prev_data and prev_data['stock'].get('brut_sacs', 0) < 10000:
        # Use previous stock info (manually updated with real stock central BEKOKO)
        summary['stock'] = prev_data['stock']
        print(f"\n[INFO] Preserved stock info from {prev_path} (BEKOKO central: {prev_data['stock'].get('brut_sacs')} sacs)")

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(summary, f, indent=2, ensure_ascii=False, default=str)
print(f"\n=== SAVED: {OUT} ===")
