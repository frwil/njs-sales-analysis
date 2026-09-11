"""Analyse Soja/Concentrés : Période 05-10/09 vs 3 semaines précédentes (15/08-04/09).
Compare les volumes globaux, par région et par agence.
Période 1 : 05/09 au 10/09/2026 (6 jours calendaire, 5 jours ouvrés lun-sam)
Période 2 : 15/08 au 04/09/2026 (21 jours calendaire, 18 jours ouvrés lun-sam)

Source données :
- 15-31/08 : dataset_2023_2026.csv (Livrée historique)
- 01-10/09 : extraction NJS GROUP ERP (36).xlsx du 11/09/2026
"""
import os
import openpyxl
import pandas as pd
from collections import defaultdict
from datetime import datetime, date, timedelta
import json

# === Constants ===
SOJA_REFS = {'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25}
CONC_REFS = {
    'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
    'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50
}

AGENCE_MAP = {
    'AGENCE FAMLA': ('Famla', 'Ouest'), 'AGENCE MESSASSI': ('Messassi', 'Centre'),
    'AGENCE BERTOUA': ('Bertoua', 'Centre'), 'AGENCE NDOBO': ('Ndobo', 'Littoral'),
    'AGENCE DJELENG': ('Djeleng', 'Ouest'), 'AGENCE VILLAGE': ('Village', 'Littoral'),
    'AGENCE PK11': ('Pk11', 'Littoral'), 'AGENCE NGAOUNDERE': ('Ngaoundere', 'Centre'),
    'AGENCE AHALA': ('Ahala', 'Centre'), 'AGENCE NKONGSAMBA': ('Nkongsamba', 'Littoral'),
    'AGENCE NKOLBISSON': ('Nkolbisson', 'Centre'), 'AGENCE NKOABANG': ('Nkoabang', 'Centre'),
    'AGENCE DE BAMENDA - DEPOT MBOUDA': ('Mbouda', 'Ouest'), 'AGENCE BUEA': ('Buea', 'Littoral'),
    'AGENCE MAROUA': ('Maroua', 'Centre'),
    'SPC BAF-CHEFFERIE': ('Baf-Chefferie', 'Ouest'), 'PDC Emana': ('Emana', 'Centre'),
    'SPC-NDERE': ('Ndere', 'Centre'), 'SPC-DSCHANG': ('Dschang', 'Ouest'),
    'SPC BUEA': ('Buea-SPC', 'Littoral'), 'SPC-YASSA': ('Yassa', 'Littoral'),
}

INTERNAL = ['SPC', 'PDC', 'COMPTOIR', 'EMANA']

def is_internal(c):
    if not c: return False
    s = str(c).upper()
    return any(p in s for p in INTERNAL)

# === Periods ===
# Period 1: 05/09 to 10/09/2026
P1_START = date(2026, 9, 5)
P1_END = date(2026, 9, 10)
# Period 2: 15/08 to 04/09/2026 (3 weeks before)
P2_START = date(2026, 8, 15)
P2_END = date(2026, 9, 4)

def count_working_days(d_start, d_end):
    """Count Mon-Sat (Cameroon work week)."""
    n = 0
    d = d_start
    while d <= d_end:
        if d.weekday() < 6:  # Mon=0, Sun=6
            n += 1
        d += timedelta(days=1)
    return n

P1_DAYS_CAL = (P1_END - P1_START).days + 1
P2_DAYS_CAL = (P2_END - P2_START).days + 1
P1_DAYS_WORK = count_working_days(P1_START, P1_END)
P2_DAYS_WORK = count_working_days(P2_START, P2_END)

print(f"Période 1: {P1_START.strftime('%d/%m/%Y')} - {P1_END.strftime('%d/%m/%Y')} = {P1_DAYS_CAL}j cal / {P1_DAYS_WORK}j ouvrés")
print(f"Période 2: {P2_START.strftime('%d/%m/%Y')} - {P2_END.strftime('%d/%m/%Y')} = {P2_DAYS_CAL}j cal / {P2_DAYS_WORK}j ouvrés")
print()

# === Data loading ===
# Load August from dataset (for P2 15-31/08)
hist_df = pd.read_csv('/home/z/my-project/scripts/dataset_2023_2026.csv', parse_dates=['date'], low_memory=False)
print(f"Loaded dataset: {len(hist_df)} records")

# Filter Aug 15-31
aug_data = hist_df[(hist_df['date'].dt.year == 2026) & (hist_df['date'].dt.month == 8) & (hist_df['date'].dt.day >= 15)]
print(f"Aug 15-31 records: {len(aug_data)}")

# Load September from extraction (36)
SEPT_SRC = '/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (36).xlsx'
wb = openpyxl.load_workbook(SEPT_SRC, read_only=True, data_only=True)
ws = wb['Sheet 1']
sep_rows = list(ws.iter_rows(min_row=3, values_only=True))
print(f"Sept extraction rows: {len(sep_rows)}")

# === Aggregation function ===
def aggregate_period(rows_data, source='dataset'):
    """Aggregate soja and conc by region and agence.
    rows_data: list of tuples (ref, qte, date_str, agence_raw, client, etat)
    source: 'dataset' for hist_df rows, 'extraction' for ERP rows
    """
    nationwide = {'soja_t': 0, 'conc_t': 0}
    by_region = defaultdict(lambda: {'soja_t': 0, 'conc_t': 0})
    by_agence = defaultdict(lambda: {'soja_t': 0, 'conc_t': 0, 'region': ''})
    
    for r in rows_data:
        if source == 'dataset':
            ref = r.get('ref')
            qte = r.get('qte', 0)
            date_obj = r.get('date')
            agence = r.get('agence', '')
            # Determine agence_raw from agence (reverse map)
            agence_raw = None
            for k, v in AGENCE_MAP.items():
                if v[0] == agence:
                    agence_raw = k
                    break
            if not agence_raw:
                agence_raw = agence
            client = r.get('tiers', '') or ''
            etat = 'Livrée'  # dataset only contains Livrées
            qte = float(qte) if qte else 0
        else:
            # ERP extraction row
            if not r or r[0] == 'Total': continue
            ref = r[0]
            qte = float(r[2]) if r[2] else 0
            date_str = str(r[6])[:10] if r[6] else ''
            try:
                date_obj = datetime.strptime(date_str, '%d/%m/%Y')
            except:
                continue
            agence_raw = r[15] if len(r) > 15 else ''
            client = r[5] if len(r) > 5 else ''
            etat = r[13] if len(r) > 13 else None
        
        # Filter: Livrées only
        if etat != 'Livrée': continue
        # Filter: external clients
        if is_internal(client): continue
        # Agence mapping
        if agence_raw in AGENCE_MAP:
            agence_short, region = AGENCE_MAP[agence_raw]
        else:
            agence_short = str(agence_raw)
            region = '?'
        
        # Convert to kg
        if ref in SOJA_REFS:
            kg = qte * SOJA_REFS[ref]
            t = kg / 1000
            nationwide['soja_t'] += t
            by_region[region]['soja_t'] += t
            by_agence[agence_short]['soja_t'] += t
            by_agence[agence_short]['region'] = region
        elif ref in CONC_REFS:
            kg = qte * CONC_REFS[ref]
            t = kg / 1000
            nationwide['conc_t'] += t
            by_region[region]['conc_t'] += t
            by_agence[agence_short]['conc_t'] += t
            by_agence[agence_short]['region'] = region
    
    return nationwide, by_region, by_agence

# === Build P1 data (05-10/09) ===
# From extraction (36) - filter date 05-10/09
P1_rows = []
for r in sep_rows:
    if not r or r[0] == 'Total': continue
    if r[13] != 'Livrée': continue
    date_str = str(r[6])[:10] if r[6] else ''
    try:
        d = datetime.strptime(date_str, '%d/%m/%Y').date()
    except:
        continue
    if P1_START <= d <= P1_END:
        P1_rows.append(r)

print(f"\nP1 rows (05-10/09): {len(P1_rows)}")
n1, r1, a1 = aggregate_period(P1_rows, source='extraction')

# === Build P2 data (15/08-04/09) ===
# Part A: Aug 15-31 from dataset
P2_rows_dataset = []
for _, row in aug_data.iterrows():
    # Convert dataset row to a dict
    P2_rows_dataset.append(row)
print(f"P2 part A (Aug 15-31): {len(P2_rows_dataset)} rows")

# Part B: Sep 1-4 from extraction (36)
P2_rows_extraction = []
for r in sep_rows:
    if not r or r[0] == 'Total': continue
    if r[13] != 'Livrée': continue
    date_str = str(r[6])[:10] if r[6] else ''
    try:
        d = datetime.strptime(date_str, '%d/%m/%Y').date()
    except:
        continue
    if P2_START <= d <= P2_END and d.month == 9:
        P2_rows_extraction.append(r)
print(f"P2 part B (Sep 1-4): {len(P2_rows_extraction)} rows")

n2a, r2a, a2a = aggregate_period(P2_rows_dataset, source='dataset')
n2b, r2b, a2b = aggregate_period(P2_rows_extraction, source='extraction')

# Merge P2
def merge(d1, d2):
    out = defaultdict(lambda: {'soja_t': 0, 'conc_t': 0})
    for k, v in d1.items():
        out[k]['soja_t'] += v['soja_t']
        out[k]['conc_t'] += v['conc_t']
    for k, v in d2.items():
        out[k]['soja_t'] += v['soja_t']
        out[k]['conc_t'] += v['conc_t']
    return dict(out)

n2 = {'soja_t': n2a['soja_t'] + n2b['soja_t'], 'conc_t': n2a['conc_t'] + n2b['conc_t']}
r2 = merge(r2a, r2b)

# Merge by agence
a2 = defaultdict(lambda: {'soja_t': 0, 'conc_t': 0, 'region': ''})
for ag, data in a2a.items():
    a2[ag]['soja_t'] += data['soja_t']
    a2[ag]['conc_t'] += data['conc_t']
    a2[ag]['region'] = data['region']
for ag, data in a2b.items():
    a2[ag]['soja_t'] += data['soja_t']
    a2[ag]['conc_t'] += data['conc_t']
    a2[ag]['region'] = data['region']
a2 = dict(a2)

# === Compute averages ===
def avg(d, days):
    return {'soja_t': d['soja_t']/days, 'conc_t': d['conc_t']/days}

n1_avg = avg(n1, P1_DAYS_WORK)
n2_avg = avg(n2, P2_DAYS_WORK)

# === Print comparison ===
def fmt(x):
    return f"{x:,.1f}".replace(',', ' ').replace('.', ',')

print("\n" + "="*80)
print("ANALYSE SOJA / CONCENTRÉS — P1 (05-10/09) vs P2 (15/08-04/09)")
print("="*80)
print(f"\nPériode 1: 05/09 - 10/09/2026 ({P1_DAYS_WORK} j ouvrés)")
print(f"Période 2: 15/08 - 04/09/2026 ({P2_DAYS_WORK} j ouvrés = 3 semaines)")

print("\n--- GLOBAL ---")
print(f"               P1 (5-10/09)    P2 (15/08-04/09)    Var vol    Var moy/j")
print(f"Soja (t)         {n1['soja_t']:8.1f}        {n2['soja_t']:8.1f}      {(n1['soja_t']/n2['soja_t']-1)*100:+6.1f}%     {(n1_avg['soja_t']/n2_avg['soja_t']-1)*100:+6.1f}%")
print(f"Soja moy/j (t)   {n1_avg['soja_t']:8.2f}        {n2_avg['soja_t']:8.2f}")
print(f"Conc (t)         {n1['conc_t']:8.1f}        {n2['conc_t']:8.1f}      {(n1['conc_t']/n2['conc_t']-1)*100:+6.1f}%     {(n1_avg['conc_t']/n2_avg['conc_t']-1)*100:+6.1f}%")
print(f"Conc moy/j (t)   {n1_avg['conc_t']:8.2f}        {n2_avg['conc_t']:8.2f}")
total1 = n1['soja_t'] + n1['conc_t']
total2 = n2['soja_t'] + n2['conc_t']
print(f"Total (t)        {total1:8.1f}        {total2:8.1f}      {(total1/total2-1)*100:+6.1f}%")
ratio1 = n1['soja_t'] / n1['conc_t'] if n1['conc_t'] > 0 else 0
ratio2 = n2['soja_t'] / n2['conc_t'] if n2['conc_t'] > 0 else 0
print(f"Ratio soja/conc  {ratio1:8.2f}:1      {ratio2:8.2f}:1")

print("\n--- PAR RÉGION ---")
print(f"{'Région':<12} {'P1 soja':<10} {'P2 soja':<10} {'Δ%':<8} {'P1 conc':<10} {'P2 conc':<10} {'Δ%':<8} {'Ratio P1':<10} {'Ratio P2':<10}")
for region in ['Ouest', 'Centre', 'Littoral']:
    p1s = r1.get(region, {'soja_t': 0})['soja_t']
    p2s = r2.get(region, {'soja_t': 0})['soja_t']
    p1c = r1.get(region, {'conc_t': 0})['conc_t']
    p2c = r2.get(region, {'conc_t': 0})['conc_t']
    var_s = (p1s/p2s - 1)*100 if p2s > 0 else 0
    var_c = (p1c/p2c - 1)*100 if p2c > 0 else 0
    r1_ratio = p1s/p1c if p1c > 0 else 0
    r2_ratio = p2s/p2c if p2c > 0 else 0
    print(f"{region:<12} {fmt(p1s):<10} {fmt(p2s):<10} {var_s:+6.1f}%  {fmt(p1c):<10} {fmt(p2c):<10} {var_c:+6.1f}%  {r1_ratio:>5.2f}:1   {r2_ratio:>5.2f}:1")

print("\n--- PAR AGENCE (top 15 par total P1) ---")
print(f"{'Agence':<15} {'Région':<10} {'P1 soja':<10} {'P2 soja':<10} {'Δ%':<8} {'P1 conc':<10} {'P2 conc':<10} {'Δ%':<8}")
# Sort agences by total P1 volume (soja + conc)
agences_sorted = sorted(a1.keys(), key=lambda a: -(a1[a]['soja_t'] + a1[a]['conc_t']))
for ag in agences_sorted[:15]:
    p1s = a1.get(ag, {'soja_t': 0})['soja_t']
    p2s = a2.get(ag, {'soja_t': 0})['soja_t']
    p1c = a1.get(ag, {'conc_t': 0})['conc_t']
    p2c = a2.get(ag, {'conc_t': 0})['conc_t']
    region = a1.get(ag, {}).get('region', '') or a2.get(ag, {}).get('region', '')
    var_s = (p1s/p2s - 1)*100 if p2s > 0 else (100 if p1s > 0 else 0)
    var_c = (p1c/p2c - 1)*100 if p2c > 0 else (100 if p1c > 0 else 0)
    print(f"{ag:<15} {region:<10} {fmt(p1s):<10} {fmt(p2s):<10} {var_s:+6.1f}%  {fmt(p1c):<10} {fmt(p2c):<10} {var_c:+6.1f}%")

# === Save summary as JSON ===
summary = {
    'p1_dates': {'start': P1_START.strftime('%d/%m/%Y'), 'end': P1_END.strftime('%d/%m/%Y'), 'work_days': P1_DAYS_WORK, 'cal_days': P1_DAYS_CAL},
    'p2_dates': {'start': P2_START.strftime('%d/%m/%Y'), 'end': P2_END.strftime('%d/%m/%Y'), 'work_days': P2_DAYS_WORK, 'cal_days': P2_DAYS_CAL},
    'global': {
        'p1': {'soja_t': n1['soja_t'], 'conc_t': n1['conc_t'], 'soja_moy_j': n1_avg['soja_t'], 'conc_moy_j': n1_avg['conc_t']},
        'p2': {'soja_t': n2['soja_t'], 'conc_t': n2['conc_t'], 'soja_moy_j': n2_avg['soja_t'], 'conc_moy_j': n2_avg['conc_t']},
        'var_soja_pct': (n1['soja_t']/n2['soja_t']-1)*100 if n2['soja_t'] > 0 else 0,
        'var_conc_pct': (n1['conc_t']/n2['conc_t']-1)*100 if n2['conc_t'] > 0 else 0,
        'var_soja_moy_j_pct': (n1_avg['soja_t']/n2_avg['soja_t']-1)*100 if n2_avg['soja_t'] > 0 else 0,
        'var_conc_moy_j_pct': (n1_avg['conc_t']/n2_avg['conc_t']-1)*100 if n2_avg['conc_t'] > 0 else 0,
    },
    'regions': {},
    'agences': {},
}

for region in ['Ouest', 'Centre', 'Littoral']:
    p1s = r1.get(region, {'soja_t': 0})['soja_t']
    p2s = r2.get(region, {'soja_t': 0})['soja_t']
    p1c = r1.get(region, {'conc_t': 0})['conc_t']
    p2c = r2.get(region, {'conc_t': 0})['conc_t']
    summary['regions'][region] = {
        'p1': {'soja_t': p1s, 'conc_t': p1c},
        'p2': {'soja_t': p2s, 'conc_t': p2c},
        'var_soja_pct': (p1s/p2s - 1)*100 if p2s > 0 else 0,
        'var_conc_pct': (p1c/p2c - 1)*100 if p2c > 0 else 0,
    }

for ag in agences_sorted[:20]:
    p1s = a1.get(ag, {'soja_t': 0})['soja_t']
    p2s = a2.get(ag, {'soja_t': 0})['soja_t']
    p1c = a1.get(ag, {'conc_t': 0})['conc_t']
    p2c = a2.get(ag, {'conc_t': 0})['conc_t']
    region = a1.get(ag, {}).get('region', '') or a2.get(ag, {}).get('region', '')
    summary['agences'][ag] = {
        'region': region,
        'p1': {'soja_t': p1s, 'conc_t': p1c},
        'p2': {'soja_t': p2s, 'conc_t': p2c},
        'var_soja_pct': (p1s/p2s - 1)*100 if p2s > 0 else (100 if p1s > 0 else 0),
        'var_conc_pct': (p1c/p2c - 1)*100 if p2c > 0 else (100 if p1c > 0 else 0),
    }

with open('/home/z/my-project/scripts/sept_p1_vs_p2_analysis.json', 'w', encoding='utf-8') as f:
    json.dump(summary, f, indent=2, ensure_ascii=False, default=str)
print(f"\n=== SAVED: /home/z/my-project/scripts/sept_p1_vs_p2_analysis.json ===")
