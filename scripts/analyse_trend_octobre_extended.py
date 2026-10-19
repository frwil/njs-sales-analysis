"""Extended trend analysis — octobre 2026 : 5 métriques par période.
- Volume (t), clients uniques, commandes (proxy date+client)
- Panier moyen par client (kg), panier moyen par commande (kg)
Question : la baisse de prix du 07/10 (20 500 → 17 600 F/sac) a-t-elle déjà
élargi la clientèle / augmenté les paniers, ou seulement les volumes ?

Périodes (6 jours lun-sam chacune) :
P1   : 05-10/10/2026 (post-baisse 07/10)
P2   : 12-17/10/2026 (première semaine complète post-baisse)
W-1 : 28/09-03/10/2026
W-2 : 21-26/09/2026
W-3 : 14-19/09/2026

Source : dataset_2023_2026.csv (reconstruit avec extraction (67), Livrée,
clients externes ; COMPTOIR inclus comme dans l'analyse trend de septembre).
"""
import pandas as pd
from collections import defaultdict
from datetime import date, timedelta
import json

# === Constants ===
SOJA_REFS = {'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25}
CONC_REFS = {
    'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
    'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50
}
INTERNAL = ['SPC', 'PDC', 'EMANA']  # COMPTOIR inclus

def is_internal(c):
    if not c: return False
    s = str(c).upper()
    return any(p in s for p in INTERNAL)

PERIODS = [
    {'name': 'W-3', 'start': date(2026, 9, 14), 'end': date(2026, 9, 19), 'label': '14-19/09'},
    {'name': 'W-2', 'start': date(2026, 9, 21), 'end': date(2026, 9, 26), 'label': '21-26/09'},
    {'name': 'W-1', 'start': date(2026, 9, 28), 'end': date(2026, 10, 3), 'label': '28/09-03/10'},
    {'name': 'P1', 'start': date(2026, 10, 5), 'end': date(2026, 10, 10), 'label': '05-10/10 (post-baisse 07/10)'},
    {'name': 'P2', 'start': date(2026, 10, 12), 'end': date(2026, 10, 17), 'label': '12-17/10'},
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

hist_df = pd.read_csv('/home/z/my-project/scripts/dataset_2023_2026.csv', parse_dates=['date'], low_memory=False)
print(f"Loaded dataset: {len(hist_df)} records")

results = {}
for p in PERIODS:
    metrics = {
        'soja_t': 0, 'conc_t': 0,
        'soja_clients': set(), 'conc_clients': set(),
        'soja_orders': set(), 'conc_orders': set(),
        'soja_baskets_per_client': [], 'conc_baskets_per_client': [],
        'soja_baskets_per_order': [], 'conc_baskets_per_order': [],
    }
    rows = hist_df[(hist_df['date'].dt.date >= p['start']) & (hist_df['date'].dt.date <= p['end'])]
    client_soja_vol = defaultdict(float)
    client_conc_vol = defaultdict(float)
    order_soja_vol = defaultdict(float)
    order_conc_vol = defaultdict(float)

    for _, row in rows.iterrows():
        ref = row.get('ref')
        qte = float(row.get('qte', 0) or 0)
        client = row.get('client', '') or ''
        d = row.get('date').date()
        order_id = f"{d}_{client}"  # proxy commande = date+client (méthode septembre)

        if is_internal(client): continue

        if ref in SOJA_REFS:
            kg = qte * SOJA_REFS[ref]
            metrics['soja_t'] += kg / 1000
            metrics['soja_clients'].add(client)
            metrics['soja_orders'].add(order_id)
            client_soja_vol[client] += kg
            order_soja_vol[order_id] += kg
        elif ref in CONC_REFS:
            kg = qte * CONC_REFS[ref]
            metrics['conc_t'] += kg / 1000
            metrics['conc_clients'].add(client)
            metrics['conc_orders'].add(order_id)
            client_conc_vol[client] += kg
            order_conc_vol[order_id] += kg

    metrics['n_soja_clients'] = len(metrics['soja_clients'])
    metrics['n_conc_clients'] = len(metrics['conc_clients'])
    metrics['n_soja_orders'] = len(metrics['soja_orders'])
    metrics['n_conc_orders'] = len(metrics['conc_orders'])
    metrics['soja_panier_client_kg'] = sum(client_soja_vol.values()) / metrics['n_soja_clients'] if metrics['n_soja_clients'] > 0 else 0
    metrics['conc_panier_client_kg'] = sum(client_conc_vol.values()) / metrics['n_conc_clients'] if metrics['n_conc_clients'] > 0 else 0
    metrics['soja_panier_order_kg'] = sum(order_soja_vol.values()) / metrics['n_soja_orders'] if metrics['n_soja_orders'] > 0 else 0
    metrics['conc_panier_order_kg'] = sum(order_conc_vol.values()) / metrics['n_conc_orders'] if metrics['n_conc_orders'] > 0 else 0
    metrics['soja_panier_client_t'] = metrics['soja_panier_client_kg'] / 1000
    metrics['conc_panier_client_t'] = metrics['conc_panier_client_kg'] / 1000
    metrics['soja_panier_order_t'] = metrics['soja_panier_order_kg'] / 1000
    metrics['conc_panier_order_t'] = metrics['conc_panier_order_kg'] / 1000
    metrics['soja_moy_j'] = metrics['soja_t'] / p['work_days']
    metrics['conc_moy_j'] = metrics['conc_t'] / p['work_days']
    metrics['work_days'] = p['work_days']
    results[p['name']] = metrics

def fmt(x):
    return f"{x:,.1f}".replace(',', ' ').replace('.', ',')

def fmt_signed(x):
    s = '+' if x >= 0 else ''
    return f"{s}{x:.1f}%".replace('.', ',')

print("\n" + "="*120)
print("ANALYSE ÉTENDUE OCTOBRE — SOJA vs CONCENTRÉS : Volume, Clients, Commandes, Panier moyen")
print("="*120)

print("\n--- SOJA — 4 métriques sur 4 périodes ---")
print(f"{'Période':<8} {'Volume (t)':<12} {'Clients':<10} {'Commandes':<12} {'Panier/client (kg)':<20} {'Panier/cmd (kg)':<18} {'Vol moy/j (t)':<15}")
for p in PERIODS:
    m = results[p['name']]
    print(f"{p['name']:<8} {fmt(m['soja_t']):<12} {m['n_soja_clients']:<10} {m['n_soja_orders']:<12} {fmt(m['soja_panier_client_kg']):<20} {fmt(m['soja_panier_order_kg']):<18} {fmt(m['soja_moy_j']):<15}")

print("\n--- CONCENTRÉS — 4 métriques sur 4 périodes ---")
print(f"{'Période':<8} {'Volume (t)':<12} {'Clients':<10} {'Commandes':<12} {'Panier/client (kg)':<20} {'Panier/cmd (kg)':<18} {'Vol moy/j (t)':<15}")
for p in PERIODS:
    m = results[p['name']]
    print(f"{p['name']:<8} {fmt(m['conc_t']):<12} {m['n_conc_clients']:<10} {m['n_conc_orders']:<12} {fmt(m['conc_panier_client_kg']):<20} {fmt(m['conc_panier_order_kg']):<18} {fmt(m['conc_moy_j']):<15}")

transitions = [('W-3 → W-2', 'W-3', 'W-2'), ('W-2 → W-1', 'W-2', 'W-1'), ('W-1 → P1', 'W-1', 'P1'), ('P1 → P2', 'P1', 'P2')]
print("\n--- VARIATIONS % vs SEMAINE PRÉCÉDENTE ---")
for fam in ['soja', 'conc']:
    print(f"\n[{fam.upper()}]")
    print(f"{'Transition':<15} {'Δ Vol':<10} {'Δ Clients':<12} {'Δ Cmds':<10} {'Δ Panier/cli':<14} {'Δ Panier/cmd':<14}")
    for label, prev, curr in transitions:
        m_prev, m_curr = results[prev], results[curr]
        var_vol = (m_curr[f'{fam}_t']/m_prev[f'{fam}_t'] - 1)*100 if m_prev[f'{fam}_t'] > 0 else 0
        var_cli = (m_curr[f'n_{fam}_clients']/m_prev[f'n_{fam}_clients'] - 1)*100 if m_prev[f'n_{fam}_clients'] > 0 else 0
        var_cmd = (m_curr[f'n_{fam}_orders']/m_prev[f'n_{fam}_orders'] - 1)*100 if m_prev[f'n_{fam}_orders'] > 0 else 0
        var_pc = (m_curr[f'{fam}_panier_client_kg']/m_prev[f'{fam}_panier_client_kg'] - 1)*100 if m_prev[f'{fam}_panier_client_kg'] > 0 else 0
        var_po = (m_curr[f'{fam}_panier_order_kg']/m_prev[f'{fam}_panier_order_kg'] - 1)*100 if m_prev[f'{fam}_panier_order_kg'] > 0 else 0
        print(f"{label:<15} {fmt_signed(var_vol):<10} {fmt_signed(var_cli):<12} {fmt_signed(var_cmd):<10} {fmt_signed(var_pc):<14} {fmt_signed(var_po):<14}")

m_w3, m_p2 = results['W-3'], results['P2']
print("\n--- VARIATIONS CUMULÉES W-3 → P2 ---")
for fam in ['soja', 'conc']:
    print(f"[{fam.upper()}]")
    print(f"  Volume:        {fmt_signed((m_p2[f'{fam}_t']/m_w3[f'{fam}_t']-1)*100):<10} ({fmt(m_w3[f'{fam}_t'])} → {fmt(m_p2[f'{fam}_t'])} t)")
    print(f"  Clients:       {fmt_signed((m_p2[f'n_{fam}_clients']/m_w3[f'n_{fam}_clients']-1)*100):<10} ({m_w3[f'n_{fam}_clients']} → {m_p2[f'n_{fam}_clients']} clients)")
    print(f"  Commandes:     {fmt_signed((m_p2[f'n_{fam}_orders']/m_w3[f'n_{fam}_orders']-1)*100):<10} ({m_w3[f'n_{fam}_orders']} → {m_p2[f'n_{fam}_orders']} cmd)")
    print(f"  Panier/client: {fmt_signed((m_p2[f'{fam}_panier_client_kg']/m_w3[f'{fam}_panier_client_kg']-1)*100):<10}")
    print(f"  Panier/cmd:    {fmt_signed((m_p2[f'{fam}_panier_order_kg']/m_w3[f'{fam}_panier_order_kg']-1)*100):<10}")

def trend_direction(values):
    if len(values) < 2: return 'unknown'
    directions = []
    for i in range(1, len(values)):
        directions.append('down' if values[i] < values[i-1] else ('up' if values[i] > values[i-1] else 'flat'))
    if all(d == 'down' for d in directions): return 'down (consistent)'
    if all(d == 'up' for d in directions): return 'up (consistent)'
    return f'mixed ({directions})'

out = {
    'periods': [{'name': p['name'], 'dates': p['label'], 'work_days': p['work_days']} for p in PERIODS],
    'soja': {p['name']: {
        'volume_t': results[p['name']]['soja_t'],
        'n_clients': results[p['name']]['n_soja_clients'],
        'n_orders': results[p['name']]['n_soja_orders'],
        'panier_client_kg': results[p['name']]['soja_panier_client_kg'],
        'panier_order_kg': results[p['name']]['soja_panier_order_kg'],
        'vol_moy_j': results[p['name']]['soja_moy_j'],
    } for p in PERIODS},
    'conc': {p['name']: {
        'volume_t': results[p['name']]['conc_t'],
        'n_clients': results[p['name']]['n_conc_clients'],
        'n_orders': results[p['name']]['n_conc_orders'],
        'panier_client_kg': results[p['name']]['conc_panier_client_kg'],
        'panier_order_kg': results[p['name']]['conc_panier_order_kg'],
        'vol_moy_j': results[p['name']]['conc_moy_j'],
    } for p in PERIODS},
    'cumul_w3_to_p2': {
        'soja': {
            'volume': (m_p2['soja_t']/m_w3['soja_t']-1)*100,
            'clients': (m_p2['n_soja_clients']/m_w3['n_soja_clients']-1)*100,
            'orders': (m_p2['n_soja_orders']/m_w3['n_soja_orders']-1)*100,
            'panier_client': (m_p2['soja_panier_client_kg']/m_w3['soja_panier_client_kg']-1)*100,
            'panier_order': (m_p2['soja_panier_order_kg']/m_w3['soja_panier_order_kg']-1)*100,
        },
        'conc': {
            'volume': (m_p2['conc_t']/m_w3['conc_t']-1)*100,
            'clients': (m_p2['n_conc_clients']/m_w3['n_conc_clients']-1)*100,
            'orders': (m_p2['n_conc_orders']/m_w3['n_conc_orders']-1)*100,
            'panier_client': (m_p2['conc_panier_client_kg']/m_w3['conc_panier_client_kg']-1)*100,
            'panier_order': (m_p2['conc_panier_order_kg']/m_w3['conc_panier_order_kg']-1)*100,
        },
    },
}
with open('/home/z/my-project/scripts/oct_trend_extended.json', 'w', encoding='utf-8') as f:
    json.dump(out, f, indent=2, ensure_ascii=False, default=str)
print(f"\n=== SAVED: /home/z/my-project/scripts/oct_trend_extended.json ===")
