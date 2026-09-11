"""Analyse Soja/Concentrés : P1 (05-10/09) vs 3 semaines individuelles précédentes (W-1, W-2, W-3).
Chaque période = 6 jours calendaire (Sam-Thu) = 5 jours ouvrés (lun-sam).
Permet de voir la TENDANCE (évolution semaine par semaine).

P1   : 05-10/09/2026 (post-hausse prix 04/09)
W-1 : 29/08-03/09/2026 (semaine juste avant hausse)
W-2 : 22-27/08/2026
W-3 : 15-20/08/2026

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

# === Periods (each = 6 calendar days, 5 working days lun-sam) ===
PERIODS = [
    {'name': 'P1', 'start': date(2026, 9, 5), 'end': date(2026, 9, 10), 'label': '05-10/09 (post-hausse)'},
    {'name': 'W-1', 'start': date(2026, 8, 29), 'end': date(2026, 9, 3), 'label': '29/08-03/09 (pré-hausse)'},
    {'name': 'W-2', 'start': date(2026, 8, 22), 'end': date(2026, 8, 27), 'label': '22-27/08'},
    {'name': 'W-3', 'start': date(2026, 8, 15), 'end': date(2026, 8, 20), 'label': '15-20/08'},
]

def count_working_days(d_start, d_end):
    n = 0
    d = d_start
    while d <= d_end:
        if d.weekday() < 6:
            n += 1
        d += timedelta(days=1)
    return n

for p in PERIODS:
    p['cal_days'] = (p['end'] - p['start']).days + 1
    p['work_days'] = count_working_days(p['start'], p['end'])
    print(f"{p['name']}: {p['start'].strftime('%d/%m/%Y')} - {p['end'].strftime('%d/%m/%Y')} = {p['cal_days']}j cal / {p['work_days']}j ouvrés ({p['label']})")

# === Data loading ===
hist_df = pd.read_csv('/home/z/my-project/scripts/dataset_2023_2026.csv', parse_dates=['date'], low_memory=False)
print(f"\nLoaded dataset: {len(hist_df)} records")

SEPT_SRC = '/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (36).xlsx'
wb = openpyxl.load_workbook(SEPT_SRC, read_only=True, data_only=True)
ws = wb['Sheet 1']
sep_rows = list(ws.iter_rows(min_row=3, values_only=True))
print(f"Sept extraction rows: {len(sep_rows)}")

# === Aggregation function ===
def aggregate_period(rows_data, source, p_start, p_end):
    """Aggregate soja and conc by region/agence for a given period."""
    nationwide = {'soja_t': 0, 'conc_t': 0}
    by_region = defaultdict(lambda: {'soja_t': 0, 'conc_t': 0})
    by_agence = defaultdict(lambda: {'soja_t': 0, 'conc_t': 0, 'region': ''})
    
    for r in rows_data:
        if source == 'dataset':
            ref = r.get('ref')
            qte = r.get('qte', 0)
            date_obj = r.get('date')
            agence = r.get('agence', '')
            agence_raw = None
            for k, v in AGENCE_MAP.items():
                if v[0] == agence:
                    agence_raw = k
                    break
            if not agence_raw:
                agence_raw = agence
            client = r.get('tiers', '') or ''
            etat = 'Livrée'
            qte = float(qte) if qte else 0
            d = date_obj.date() if hasattr(date_obj, 'date') else date_obj
        else:
            if not r or r[0] == 'Total': continue
            ref = r[0]
            qte = float(r[2]) if r[2] else 0
            date_str = str(r[6])[:10] if r[6] else ''
            try:
                d = datetime.strptime(date_str, '%d/%m/%Y').date()
            except:
                continue
            agence_raw = r[15] if len(r) > 15 else ''
            client = r[5] if len(r) > 5 else ''
            etat = r[13] if len(r) > 13 else None
        
        # Filter date
        if not (p_start <= d <= p_end):
            continue
        # Filter etat
        if etat != 'Livrée': continue
        # Filter internal clients
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
    
    return nationwide, dict(by_region), dict(by_agence)

# === Compute aggregates for each period ===
results = {}
for p in PERIODS:
    # Filter rows for this period from historical dataset
    aug_data_p = hist_df[(hist_df['date'].dt.date >= p['start']) & (hist_df['date'].dt.date <= p['end'])]
    
    # Aggregate from dataset (Aug periods)
    n, r, a = {'soja_t': 0, 'conc_t': 0}, defaultdict(lambda: {'soja_t': 0, 'conc_t': 0}), defaultdict(lambda: {'soja_t': 0, 'conc_t': 0, 'region': ''})
    for _, row in aug_data_p.iterrows():
        ref = row.get('ref')
        qte = float(row.get('qte', 0) or 0)
        date_obj = row.get('date')
        agence = row.get('agence', '')
        agence_raw = None
        for k, v in AGENCE_MAP.items():
            if v[0] == agence:
                agence_raw = k
                break
        if not agence_raw:
            agence_raw = agence
        client = row.get('tiers', '') or ''
        etat = 'Livrée'  # dataset contains only Livrées
        d = date_obj.date() if hasattr(date_obj, 'date') else date_obj
        
        # Filter internal clients
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
            n['soja_t'] += t
            r[region]['soja_t'] += t
            a[agence_short]['soja_t'] += t
            a[agence_short]['region'] = region
        elif ref in CONC_REFS:
            kg = qte * CONC_REFS[ref]
            t = kg / 1000
            n['conc_t'] += t
            r[region]['conc_t'] += t
            a[agence_short]['conc_t'] += t
            a[agence_short]['region'] = region
    
    # Also include extraction rows for periods overlapping September
    n2, r2, a2 = aggregate_period(sep_rows, 'extraction', p['start'], p['end'])
    
    # Merge
    n['soja_t'] += n2['soja_t']
    n['conc_t'] += n2['conc_t']
    for region in set(list(r.keys()) + list(r2.keys())):
        if region not in r: r[region] = {'soja_t': 0, 'conc_t': 0}
        if region in r2:
            r[region]['soja_t'] += r2[region]['soja_t']
            r[region]['conc_t'] += r2[region]['conc_t']
    for ag in set(list(a.keys()) + list(a2.keys())):
        if ag not in a: a[ag] = {'soja_t': 0, 'conc_t': 0, 'region': ''}
        if ag in a2:
            a[ag]['soja_t'] += a2[ag]['soja_t']
            a[ag]['conc_t'] += a2[ag]['conc_t']
            a[ag]['region'] = a2[ag]['region'] or a[ag]['region']
    
    n['soja_moy_j'] = n['soja_t'] / p['work_days']
    n['conc_moy_j'] = n['conc_t'] / p['work_days']
    n['ratio'] = n['soja_t'] / n['conc_t'] if n['conc_t'] > 0 else 0
    n['total_t'] = n['soja_t'] + n['conc_t']
    
    results[p['name']] = {
        'nationwide': n, 'regions': dict(r), 'agences': dict(a),
        'dates': {'start': p['start'].strftime('%d/%m/%Y'), 'end': p['end'].strftime('%d/%m/%Y'),
                  'work_days': p['work_days'], 'cal_days': p['cal_days'], 'label': p['label']},
    }

# === Print comparison ===
def fmt(x):
    return f"{x:,.1f}".replace(',', ' ').replace('.', ',')

def fmt_signed(x):
    s = '+' if x >= 0 else ''
    return f"{s}{x:.1f}%".replace('.', ',')

print("\n" + "="*100)
print("ANALYSE SOJA / CONCENTRÉS — ÉVOLUTION SUR 4 PÉRIODES (mêmes nb de jours = 5j ouvrés)")
print("="*100)

print("\n--- GLOBAL — ÉVOLUTION SEMAINE PAR SEMAINE ---")
print(f"{'Période':<12} {'Dates':<22} {'Soja (t)':<10} {'Conc (t)':<10} {'Total (t)':<10} {'Soja moy/j':<12} {'Conc moy/j':<12} {'Ratio':<8}")
for p in PERIODS:
    r = results[p['name']]
    n = r['nationwide']
    print(f"{p['name']:<12} {r['dates']['label']:<22} {fmt(n['soja_t']):<10} {fmt(n['conc_t']):<10} {fmt(n['total_t']):<10} {fmt(n['soja_moy_j']):<12} {fmt(n['conc_moy_j']):<12} {n['ratio']:.2f}:1")

print("\n--- VARIATION % vs PÉRIODE PRÉCÉDENTE ---")
print(f"{'Transition':<20} {'Δ Soja':<12} {'Δ Conc':<12} {'Δ Total':<12} {'Δ Soja moy/j':<14} {'Δ Conc moy/j':<14}")
transitions = [('W-3 → W-2', 'W-3', 'W-2'), ('W-2 → W-1', 'W-2', 'W-1'), ('W-1 → P1', 'W-1', 'P1')]
for label, prev, curr in transitions:
    n_prev = results[prev]['nationwide']
    n_curr = results[curr]['nationwide']
    var_s = (n_curr['soja_t']/n_prev['soja_t'] - 1)*100 if n_prev['soja_t'] > 0 else 0
    var_c = (n_curr['conc_t']/n_prev['conc_t'] - 1)*100 if n_prev['conc_t'] > 0 else 0
    var_t = (n_curr['total_t']/n_prev['total_t'] - 1)*100 if n_prev['total_t'] > 0 else 0
    var_s_moy = (n_curr['soja_moy_j']/n_prev['soja_moy_j'] - 1)*100 if n_prev['soja_moy_j'] > 0 else 0
    var_c_moy = (n_curr['conc_moy_j']/n_prev['conc_moy_j'] - 1)*100 if n_prev['conc_moy_j'] > 0 else 0
    print(f"{label:<20} {fmt_signed(var_s):<12} {fmt_signed(var_c):<12} {fmt_signed(var_t):<12} {fmt_signed(var_s_moy):<14} {fmt_signed(var_c_moy):<14}")

print("\n--- VARIATION % vs P1 (actuel) ---")
print(f"{'Période':<12} {'Δ Soja':<12} {'Δ Conc':<12} {'Δ Total':<12}")
n_p1 = results['P1']['nationwide']
for prev in ['W-3', 'W-2', 'W-1']:
    n_prev = results[prev]['nationwide']
    var_s = (n_p1['soja_t']/n_prev['soja_t'] - 1)*100 if n_prev['soja_t'] > 0 else 0
    var_c = (n_p1['conc_t']/n_prev['conc_t'] - 1)*100 if n_prev['conc_t'] > 0 else 0
    var_t = (n_p1['total_t']/n_prev['total_t'] - 1)*100 if n_prev['total_t'] > 0 else 0
    print(f"{prev+' → P1':<12} {fmt_signed(var_s):<12} {fmt_signed(var_c):<12} {fmt_signed(var_t):<12}")

print("\n--- PAR RÉGION — ÉVOLUTION ---")
for region in ['Ouest', 'Centre', 'Littoral']:
    print(f"\n[{region.upper()}]")
    print(f"{'Période':<8} {'Soja (t)':<10} {'Conc (t)':<10} {'Total (t)':<10} {'Soja moy/j':<12} {'Conc moy/j':<12} {'Ratio':<8}")
    for p in PERIODS:
        n = results[p['name']]['regions'].get(region, {'soja_t': 0, 'conc_t': 0})
        work_days = results[p['name']]['dates']['work_days']
        soja_moy = n['soja_t']/work_days
        conc_moy = n['conc_t']/work_days
        ratio = n['soja_t']/n['conc_t'] if n['conc_t'] > 0 else 0
        total = n['soja_t'] + n['conc_t']
        print(f"{p['name']:<8} {fmt(n['soja_t']):<10} {fmt(n['conc_t']):<10} {fmt(total):<10} {fmt(soja_moy):<12} {fmt(conc_moy):<12} {ratio:.2f}:1")

print("\n--- PAR AGENCE — TOP 15 (tri par volume P1) ---")
print(f"{'Agence':<15} {'Région':<10} | {'P1 Soja':<10} {'W-1 Soja':<10} {'W-2 Soja':<10} {'W-3 Soja':<10} | {'P1 Conc':<10} {'W-1 Conc':<10} {'W-2 Conc':<10} {'W-3 Conc':<10}")
# Sort agences by P1 total
agences_sorted = sorted(results['P1']['agences'].keys(),
                       key=lambda a: -(results['P1']['agences'][a]['soja_t'] + results['P1']['agences'][a]['conc_t']))
for ag in agences_sorted[:15]:
    region = results['P1']['agences'][ag]['region']
    soja_vals = [results[p['name']]['agences'].get(ag, {'soja_t': 0})['soja_t'] for p in PERIODS]
    conc_vals = [results[p['name']]['agences'].get(ag, {'conc_t': 0})['conc_t'] for p in PERIODS]
    print(f"{ag:<15} {region:<10} | {fmt(soja_vals[0]):<10} {fmt(soja_vals[1]):<10} {fmt(soja_vals[2]):<10} {fmt(soja_vals[3]):<10} | {fmt(conc_vals[0]):<10} {fmt(conc_vals[1]):<10} {fmt(conc_vals[2]):<10} {fmt(conc_vals[3]):<10}")

# === Save summary as JSON ===
summary = {
    'periods': [{'name': p['name'], 'dates': results[p['name']]['dates']} for p in PERIODS],
    'global': {p['name']: results[p['name']]['nationwide'] for p in PERIODS},
    'regions': {region: {p['name']: results[p['name']]['regions'].get(region, {'soja_t': 0, 'conc_t': 0})
                        for p in PERIODS} for region in ['Ouest', 'Centre', 'Littoral']},
}

# Agences (top 20) — explicit loop for clarity
summary['agences'] = {}
for ag in agences_sorted[:20]:
    summary['agences'][ag] = {
        'region': results['P1']['agences'][ag]['region'],
    }
    for p in PERIODS:
        summary['agences'][ag][p['name']] = {
            'soja_t': results[p['name']]['agences'].get(ag, {'soja_t': 0})['soja_t'],
            'conc_t': results[p['name']]['agences'].get(ag, {'conc_t': 0})['conc_t'],
        }

# Compute variations vs P1
summary['variations_vs_p1'] = {
    'global': {prev: {
        'soja': (results['P1']['nationwide']['soja_t']/results[prev]['nationwide']['soja_t'] - 1)*100 if results[prev]['nationwide']['soja_t'] > 0 else 0,
        'conc': (results['P1']['nationwide']['conc_t']/results[prev]['nationwide']['conc_t'] - 1)*100 if results[prev]['nationwide']['conc_t'] > 0 else 0,
    } for prev in ['W-3', 'W-2', 'W-1']},
    'regions': {region: {prev: {
        'soja': (results['P1']['regions'].get(region, {'soja_t': 0})['soja_t']/results[prev]['regions'].get(region, {'soja_t': 1})['soja_t'] - 1)*100,
        'conc': (results['P1']['regions'].get(region, {'conc_t': 0})['conc_t']/results[prev]['regions'].get(region, {'conc_t': 1})['conc_t'] - 1)*100,
    } for prev in ['W-3', 'W-2', 'W-1']} for region in ['Ouest', 'Centre', 'Littoral']}
}

# Compute trend (week-over-week variations)
summary['trend_week_over_week'] = {}
for label, prev, curr in transitions:
    summary['trend_week_over_week'][label] = {
        'soja': (results[curr]['nationwide']['soja_t']/results[prev]['nationwide']['soja_t'] - 1)*100 if results[prev]['nationwide']['soja_t'] > 0 else 0,
        'conc': (results[curr]['nationwide']['conc_t']/results[prev]['nationwide']['conc_t'] - 1)*100 if results[prev]['nationwide']['conc_t'] > 0 else 0,
    }

with open('/home/z/my-project/scripts/sept_trend_analysis.json', 'w', encoding='utf-8') as f:
    json.dump(summary, f, indent=2, ensure_ascii=False, default=str)
print(f"\n=== SAVED: /home/z/my-project/scripts/sept_trend_analysis.json ===")
