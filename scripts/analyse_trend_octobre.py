"""Analyse Soja/Concentrés : P1 (05-10/10) et P2 (12-17/10) vs 3 semaines individuelles précédentes (W-1, W-2, W-3).
Chaque période = 6 jours calendaire (lun-sam) = 6 jours ouvrés.
Permet de voir la TENDANCE (évolution semaine par semaine).

P1   : 05-10/10/2026 (post-baisse prix 07/10 : 20 500 → 17 600 F/sac)
P2   : 12-17/10/2026 (première semaine complète post-baisse)
W-1 : 28/09-03/10/2026
W-2 : 21-26/09/2026
W-3 : 14-19/09/2026

Source données : dataset_2023_2026.csv (Livrée historique, reconstruit avec
l'extraction (67) Octobre 01-17/10). Les périodes de septembre proviennent du
bilan définitif (61) — cohérent avec les autres livrables d'octobre.
"""
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

INTERNAL_ALL = ['SPC', 'PDC', 'EMANA']  # COMPTOIR inclus (ventes au comptoir, à compter)

def is_internal(c):
    if not c: return False
    s = str(c).upper()
    return any(p in s for p in INTERNAL_ALL)

# === Periods (each = 6 calendar days, lun-sam = 6 jours ouvrés) ===
PERIODS = [
    {'name': 'P1', 'start': date(2026, 10, 5), 'end': date(2026, 10, 10), 'label': '05-10/10 (post-baisse 07/10)'},
    {'name': 'P2', 'start': date(2026, 10, 12), 'end': date(2026, 10, 17), 'label': '12-17/10'},
    {'name': 'W-1', 'start': date(2026, 9, 28), 'end': date(2026, 10, 3), 'label': '28/09-03/10'},
    {'name': 'W-2', 'start': date(2026, 9, 21), 'end': date(2026, 9, 26), 'label': '21-26/09'},
    {'name': 'W-3', 'start': date(2026, 9, 14), 'end': date(2026, 9, 19), 'label': '14-19/09'},
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

# === Data loading (dataset seul : couvre 14/09 → 17/10) ===
hist_df = pd.read_csv('/home/z/my-project/scripts/dataset_2023_2026.csv', parse_dates=['date'], low_memory=False)
print(f"\nLoaded dataset: {len(hist_df)} records")

# === Aggregation per period ===
results = {}
for p in PERIODS:
    rows = hist_df[(hist_df['date'].dt.date >= p['start']) & (hist_df['date'].dt.date <= p['end'])]
    nationwide = {'soja_t': 0, 'conc_t': 0, 'soja_ca': 0.0, 'soja_sacs': 0.0}
    by_region = defaultdict(lambda: {'soja_t': 0, 'conc_t': 0})
    by_agence = defaultdict(lambda: {'soja_t': 0, 'conc_t': 0, 'region': ''})

    for _, row in rows.iterrows():
        ref = row.get('ref')
        qte = float(row.get('qte', 0) or 0)
        agence = row.get('agence', '')
        client = row.get('client', '') or ''
        if is_internal(client): continue  # dataset Livrée, clients externes

        agence_raw = None
        for k, v in AGENCE_MAP.items():
            if v[0] == agence:
                agence_raw = k
                break
        agence_raw = agence_raw or agence
        if agence_raw in AGENCE_MAP:
            agence_short, region = AGENCE_MAP[agence_raw]
        else:
            agence_short, region = str(agence_raw), '?'

        if ref in SOJA_REFS:
            t = qte * SOJA_REFS[ref] / 1000
            nationwide['soja_t'] += t
            nationwide['soja_ca'] += float(row.get('montant_ttc', 0) or 0)
            nationwide['soja_sacs'] += qte * SOJA_REFS[ref] / 50
            by_region[region]['soja_t'] += t
            by_agence[agence_short]['soja_t'] += t
            by_agence[agence_short]['region'] = region
        elif ref in CONC_REFS:
            t = qte * CONC_REFS[ref] / 1000
            nationwide['conc_t'] += t
            by_region[region]['conc_t'] += t
            by_agence[agence_short]['conc_t'] += t
            by_agence[agence_short]['region'] = region

    nationwide['soja_moy_j'] = nationwide['soja_t'] / p['work_days']
    nationwide['conc_moy_j'] = nationwide['conc_t'] / p['work_days']
    nationwide['ratio'] = nationwide['soja_t'] / nationwide['conc_t'] if nationwide['conc_t'] > 0 else 0
    nationwide['total_t'] = nationwide['soja_t'] + nationwide['conc_t']
    nationwide['prix_soja_sac'] = nationwide['soja_ca'] / nationwide['soja_sacs'] if nationwide['soja_sacs'] > 0 else 0

    results[p['name']] = {
        'nationwide': nationwide, 'regions': dict(by_region), 'agences': dict(by_agence),
        'dates': {'start': p['start'].strftime('%d/%m/%Y'), 'end': p['end'].strftime('%d/%m/%Y'),
                  'work_days': p['work_days'], 'cal_days': p['cal_days'], 'label': p['label']},
    }

# === Print comparison ===
def fmt(x):
    return f"{x:,.1f}".replace(',', ' ').replace('.', ',')

def fmt_signed(x):
    s = '+' if x >= 0 else ''
    return f"{s}{x:.1f}%".replace('.', ',')

print("\n" + "="*110)
print("ANALYSE SOJA / CONCENTRÉS — OCTOBRE : ÉVOLUTION SUR 4 PÉRIODES (6 jours lun-sam chacune)")
print("="*110)

print("\n--- GLOBAL — ÉVOLUTION SEMAINE PAR SEMAINE ---")
print(f"{'Période':<8} {'Dates':<24} {'Soja (t)':<10} {'Conc (t)':<10} {'Ratio':<8} {'Soja moy/j':<12} {'Conc moy/j':<12} {'Prix soja/sac':<15}")
for p in PERIODS:
    n = results[p['name']]['nationwide']
    print(f"{p['name']:<8} {p['label']:<24} {fmt(n['soja_t']):<10} {fmt(n['conc_t']):<10} {n['ratio']:.2f}:1   {fmt(n['soja_moy_j']):<12} {fmt(n['conc_moy_j']):<12} {n['prix_soja_sac']:,.0f} F")

print("\n--- VARIATION % vs PÉRIODE PRÉCÉDENTE ---")
transitions = [('W-3 → W-2', 'W-3', 'W-2'), ('W-2 → W-1', 'W-2', 'W-1'), ('W-1 → P1', 'W-1', 'P1'), ('P1 → P2', 'P1', 'P2')]
print(f"{'Transition':<16} {'Δ Soja':<10} {'Δ Conc':<10} {'Δ Total':<10} {'Δ Soja moy/j':<14} {'Δ Conc moy/j':<14}")
for label, prev, curr in transitions:
    n_prev = results[prev]['nationwide']
    n_curr = results[curr]['nationwide']
    var_s = (n_curr['soja_t']/n_prev['soja_t'] - 1)*100 if n_prev['soja_t'] > 0 else 0
    var_c = (n_curr['conc_t']/n_prev['conc_t'] - 1)*100 if n_prev['conc_t'] > 0 else 0
    var_t = (n_curr['total_t']/n_prev['total_t'] - 1)*100 if n_prev['total_t'] > 0 else 0
    var_sm = (n_curr['soja_moy_j']/n_prev['soja_moy_j'] - 1)*100 if n_prev['soja_moy_j'] > 0 else 0
    var_cm = (n_curr['conc_moy_j']/n_prev['conc_moy_j'] - 1)*100 if n_prev['conc_moy_j'] > 0 else 0
    print(f"{label:<16} {fmt_signed(var_s):<10} {fmt_signed(var_c):<10} {fmt_signed(var_t):<10} {fmt_signed(var_sm):<14} {fmt_signed(var_cm):<14}")

print("\n--- VARIATION % vs P2 (actuel) ---")
n_p2 = results['P2']['nationwide']
for prev in ['W-3', 'W-2', 'W-1', 'P1']:
    n_prev = results[prev]['nationwide']
    var_s = (n_p2['soja_t']/n_prev['soja_t'] - 1)*100 if n_prev['soja_t'] > 0 else 0
    var_c = (n_p2['conc_t']/n_prev['conc_t'] - 1)*100 if n_prev['conc_t'] > 0 else 0
    print(f"{prev+' → P2':<14} {fmt_signed(var_s):<10} {fmt_signed(var_c):<10}")

print("\n--- PAR RÉGION ---")
for region in ['Ouest', 'Centre', 'Littoral']:
    print(f"\n[{region.upper()}]")
    print(f"{'Période':<8} {'Soja (t)':<10} {'Conc (t)':<10} {'Ratio':<8}")
    for p in PERIODS:
        n = results[p['name']]['regions'].get(region, {'soja_t': 0, 'conc_t': 0})
        ratio = n['soja_t']/n['conc_t'] if n['conc_t'] > 0 else 0
        print(f"{p['name']:<8} {fmt(n['soja_t']):<10} {fmt(n['conc_t']):<10} {ratio:.2f}:1")

print("\n--- PAR AGENCE — TOP 15 (tri par volume P2) ---")
agences_sorted = sorted(results['P2']['agences'].keys(),
                       key=lambda a: -(results['P2']['agences'][a]['soja_t'] + results['P2']['agences'][a]['conc_t']))
for ag in agences_sorted[:15]:
    region = results['P2']['agences'][ag]['region']
    soja_vals = [results[p['name']]['agences'].get(ag, {'soja_t': 0})['soja_t'] for p in PERIODS]
    conc_vals = [results[p['name']]['agences'].get(ag, {'conc_t': 0})['conc_t'] for p in PERIODS]
    print(f"{ag:<15} {region:<10} | P2 {fmt(soja_vals[1])}/{fmt(conc_vals[1])} | P1 {fmt(soja_vals[0])}/{fmt(conc_vals[0])} | W-1 {fmt(soja_vals[2])}/{fmt(conc_vals[2])} | W-2 {fmt(soja_vals[3])}/{fmt(conc_vals[3])} | W-3 {fmt(soja_vals[4])}/{fmt(conc_vals[4])}")

# === Save summary as JSON ===
summary = {
    'periods': [{'name': p['name'], 'dates': results[p['name']]['dates']} for p in PERIODS],
    'global': {p['name']: results[p['name']]['nationwide'] for p in PERIODS},
    'regions': {region: {p['name']: results[p['name']]['regions'].get(region, {'soja_t': 0, 'conc_t': 0})
                        for p in PERIODS} for region in ['Ouest', 'Centre', 'Littoral']},
}
summary['agences'] = {}
for ag in agences_sorted[:20]:
    summary['agences'][ag] = {'region': results['P2']['agences'][ag]['region']}
    for p in PERIODS:
        summary['agences'][ag][p['name']] = {
            'soja_t': results[p['name']]['agences'].get(ag, {'soja_t': 0})['soja_t'],
            'conc_t': results[p['name']]['agences'].get(ag, {'conc_t': 0})['conc_t'],
        }
summary['variations_vs_p2'] = {
    'global': {prev: {
        'soja': (results['P2']['nationwide']['soja_t']/results[prev]['nationwide']['soja_t'] - 1)*100 if results[prev]['nationwide']['soja_t'] > 0 else 0,
        'conc': (results['P2']['nationwide']['conc_t']/results[prev]['nationwide']['conc_t'] - 1)*100 if results[prev]['nationwide']['conc_t'] > 0 else 0,
    } for prev in ['W-3', 'W-2', 'W-1', 'P1']},
}
summary['trend_week_over_week'] = {}
for label, prev, curr in transitions:
    summary['trend_week_over_week'][label] = {
        'soja': (results[curr]['nationwide']['soja_t']/results[prev]['nationwide']['soja_t'] - 1)*100 if results[prev]['nationwide']['soja_t'] > 0 else 0,
        'conc': (results[curr]['nationwide']['conc_t']/results[prev]['nationwide']['conc_t'] - 1)*100 if results[prev]['nationwide']['conc_t'] > 0 else 0,
    }

with open('/home/z/my-project/scripts/oct_trend_analysis.json', 'w', encoding='utf-8') as f:
    json.dump(summary, f, indent=2, ensure_ascii=False, default=str)
print(f"\n=== SAVED: /home/z/my-project/scripts/oct_trend_analysis.json ===")
