"""Extended trend analysis — adds 3 new metrics per period:
- Number of unique clients (soja, conc)
- Number of orders (commandes soja, conc)
- Average basket per client (panier moyen en tonnes)
- Average basket per order (panier moyen par commande)

Question: do these metrics confirm that SOJA is "stable" (oscillating) and CONC is "declining"?
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

# === Data loading ===
hist_df = pd.read_csv('/home/z/my-project/scripts/dataset_2023_2026.csv', parse_dates=['date'], low_memory=False)
print(f"Loaded dataset: {len(hist_df)} records")

SEPT_SRC = '/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (36).xlsx'
wb = openpyxl.load_workbook(SEPT_SRC, read_only=True, data_only=True)
ws = wb['Sheet 1']
sep_rows = list(ws.iter_rows(min_row=3, values_only=True))
print(f"Sept extraction rows: {len(sep_rows)}")

# === Compute metrics for each period ===
results = {}
for p in PERIODS:
    metrics = {
        'soja_t': 0, 'conc_t': 0,
        'soja_clients': set(), 'conc_clients': set(),
        'soja_orders': set(), 'conc_orders': set(),
        'soja_baskets_per_client': [],  # kg per client
        'conc_baskets_per_client': [],
        'soja_baskets_per_order': [],  # kg per order
        'conc_baskets_per_order': [],
    }
    
    # Process Aug periods from dataset
    aug_data_p = hist_df[(hist_df['date'].dt.date >= p['start']) & (hist_df['date'].dt.date <= p['end'])]
    
    # Build client→volume and order→volume dicts
    client_soja_vol = defaultdict(float)
    client_conc_vol = defaultdict(float)
    order_soja_vol = defaultdict(float)
    order_conc_vol = defaultdict(float)
    
    for _, row in aug_data_p.iterrows():
        ref = row.get('ref')
        qte = float(row.get('qte', 0) or 0)
        agence = row.get('agence', '')
        client = row.get('client', '') or ''
        date_obj = row.get('date')
        d = date_obj.date() if hasattr(date_obj, 'date') else date_obj
        
        # The dataset doesn't have order ref, so we use (date, client) as order ID proxy
        # Actually, we should use date+client as order proxy
        order_id = f"{d}_{client}"
        
        # Filter internal clients
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
    
    # Process Sep extraction (rows from extraction file)
    for r in sep_rows:
        if not r or r[0] == 'Total': continue
        if r[13] != 'Livrée': continue
        date_str = str(r[6])[:10] if r[6] else ''
        try:
            d = datetime.strptime(date_str, '%d/%m/%Y').date()
        except:
            continue
        if not (p['start'] <= d <= p['end']): continue
        
        ref = r[0]
        qte = float(r[2]) if r[2] else 0
        agence_raw = r[15] if len(r) > 15 else ''
        client = r[5] if len(r) > 5 else ''
        # Use (date, client) as order proxy (same as dataset for fair comparison)
        order_id = f"{d}_{client}"
        
        if is_internal(client): continue
        if not client: continue
        
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
    
    # Convert sets to counts
    metrics['n_soja_clients'] = len(metrics['soja_clients'])
    metrics['n_conc_clients'] = len(metrics['conc_clients'])
    metrics['n_soja_orders'] = len(metrics['soja_orders'])
    metrics['n_conc_orders'] = len(metrics['conc_orders'])
    
    # Compute averages
    metrics['soja_panier_client_kg'] = sum(client_soja_vol.values()) / metrics['n_soja_clients'] if metrics['n_soja_clients'] > 0 else 0
    metrics['conc_panier_client_kg'] = sum(client_conc_vol.values()) / metrics['n_conc_clients'] if metrics['n_conc_clients'] > 0 else 0
    metrics['soja_panier_order_kg'] = sum(order_soja_vol.values()) / metrics['n_soja_orders'] if metrics['n_soja_orders'] > 0 else 0
    metrics['conc_panier_order_kg'] = sum(order_conc_vol.values()) / metrics['n_conc_orders'] if metrics['n_conc_orders'] > 0 else 0
    
    # Convert to tonnes for readability
    metrics['soja_panier_client_t'] = metrics['soja_panier_client_kg'] / 1000
    metrics['conc_panier_client_t'] = metrics['conc_panier_client_kg'] / 1000
    metrics['soja_panier_order_t'] = metrics['soja_panier_order_kg'] / 1000
    metrics['conc_panier_order_t'] = metrics['conc_panier_order_kg'] / 1000
    
    # Daily averages for volumes
    metrics['soja_moy_j'] = metrics['soja_t'] / p['work_days']
    metrics['conc_moy_j'] = metrics['conc_t'] / p['work_days']
    metrics['work_days'] = p['work_days']
    
    results[p['name']] = metrics

# === Print comparison ===
def fmt(x):
    return f"{x:,.1f}".replace(',', ' ').replace('.', ',')

def fmt_signed(x):
    s = '+' if x >= 0 else ''
    return f"{s}{x:.1f}%".replace('.', ',')

print("\n" + "="*120)
print("ANALYSE ÉTENDUE — SOJA vs CONCENTRÉS : Volume, Clients, Commandes, Panier moyen")
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

# === Variations week-over-week ===
print("\n--- VARIATIONS % vs SEMAINE PRÉCÉDENTE ---")
transitions = [('W-3 → W-2', 'W-3', 'W-2'), ('W-2 → W-1', 'W-2', 'W-1'), ('W-1 → P1', 'W-1', 'P1')]

print("\n[SOJA]")
print(f"{'Transition':<15} {'Δ Vol':<10} {'Δ Clients':<12} {'Δ Cmds':<10} {'Δ Panier/cli':<14} {'Δ Panier/cmd':<14}")
for label, prev, curr in transitions:
    m_prev = results[prev]
    m_curr = results[curr]
    var_vol = (m_curr['soja_t']/m_prev['soja_t'] - 1)*100 if m_prev['soja_t'] > 0 else 0
    var_cli = (m_curr['n_soja_clients']/m_prev['n_soja_clients'] - 1)*100 if m_prev['n_soja_clients'] > 0 else 0
    var_cmd = (m_curr['n_soja_orders']/m_prev['n_soja_orders'] - 1)*100 if m_prev['n_soja_orders'] > 0 else 0
    var_pc = (m_curr['soja_panier_client_kg']/m_prev['soja_panier_client_kg'] - 1)*100 if m_prev['soja_panier_client_kg'] > 0 else 0
    var_po = (m_curr['soja_panier_order_kg']/m_prev['soja_panier_order_kg'] - 1)*100 if m_prev['soja_panier_order_kg'] > 0 else 0
    print(f"{label:<15} {fmt_signed(var_vol):<10} {fmt_signed(var_cli):<12} {fmt_signed(var_cmd):<10} {fmt_signed(var_pc):<14} {fmt_signed(var_po):<14}")

print("\n[CONCENTRÉS]")
print(f"{'Transition':<15} {'Δ Vol':<10} {'Δ Clients':<12} {'Δ Cmds':<10} {'Δ Panier/cli':<14} {'Δ Panier/cmd':<14}")
for label, prev, curr in transitions:
    m_prev = results[prev]
    m_curr = results[curr]
    var_vol = (m_curr['conc_t']/m_prev['conc_t'] - 1)*100 if m_prev['conc_t'] > 0 else 0
    var_cli = (m_curr['n_conc_clients']/m_prev['n_conc_clients'] - 1)*100 if m_prev['n_conc_clients'] > 0 else 0
    var_cmd = (m_curr['n_conc_orders']/m_prev['n_conc_orders'] - 1)*100 if m_prev['n_conc_orders'] > 0 else 0
    var_pc = (m_curr['conc_panier_client_kg']/m_prev['conc_panier_client_kg'] - 1)*100 if m_prev['conc_panier_client_kg'] > 0 else 0
    var_po = (m_curr['conc_panier_order_kg']/m_prev['conc_panier_order_kg'] - 1)*100 if m_prev['conc_panier_order_kg'] > 0 else 0
    print(f"{label:<15} {fmt_signed(var_vol):<10} {fmt_signed(var_cli):<12} {fmt_signed(var_cmd):<10} {fmt_signed(var_pc):<14} {fmt_signed(var_po):<14}")

# === Cumul variations W-3 → P1 ===
print("\n--- VARIATIONS CUMULÉES W-3 → P1 ---")
m_w3 = results['W-3']
m_p1 = results['P1']

print("\n[SOJA]")
print(f"  Volume:        {fmt_signed((m_p1['soja_t']/m_w3['soja_t']-1)*100):<10} ({fmt(m_w3['soja_t'])} → {fmt(m_p1['soja_t'])} t)")
print(f"  Clients:       {fmt_signed((m_p1['n_soja_clients']/m_w3['n_soja_clients']-1)*100):<10} ({m_w3['n_soja_clients']} → {m_p1['n_soja_clients']} clients)")
print(f"  Commandes:     {fmt_signed((m_p1['n_soja_orders']/m_w3['n_soja_orders']-1)*100):<10} ({m_w3['n_soja_orders']} → {m_p1['n_soja_orders']} cmd)")
print(f"  Panier/client: {fmt_signed((m_p1['soja_panier_client_kg']/m_w3['soja_panier_client_kg']-1)*100):<10} ({fmt(m_w3['soja_panier_client_kg'])} → {fmt(m_p1['soja_panier_client_kg'])} kg)")
print(f"  Panier/cmd:    {fmt_signed((m_p1['soja_panier_order_kg']/m_w3['soja_panier_order_kg']-1)*100):<10} ({fmt(m_w3['soja_panier_order_kg'])} → {fmt(m_p1['soja_panier_order_kg'])} kg)")

print("\n[CONCENTRÉS]")
print(f"  Volume:        {fmt_signed((m_p1['conc_t']/m_w3['conc_t']-1)*100):<10} ({fmt(m_w3['conc_t'])} → {fmt(m_p1['conc_t'])} t)")
print(f"  Clients:       {fmt_signed((m_p1['n_conc_clients']/m_w3['n_conc_clients']-1)*100):<10} ({m_w3['n_conc_clients']} → {m_p1['n_conc_clients']} clients)")
print(f"  Commandes:     {fmt_signed((m_p1['n_conc_orders']/m_w3['n_conc_orders']-1)*100):<10} ({m_w3['n_conc_orders']} → {m_p1['n_conc_orders']} cmd)")
print(f"  Panier/client: {fmt_signed((m_p1['conc_panier_client_kg']/m_w3['conc_panier_client_kg']-1)*100):<10} ({fmt(m_w3['conc_panier_client_kg'])} → {fmt(m_p1['conc_panier_client_kg'])} kg)")
print(f"  Panier/cmd:    {fmt_signed((m_p1['conc_panier_order_kg']/m_w3['conc_panier_order_kg']-1)*100):<10} ({fmt(m_w3['conc_panier_order_kg'])} → {fmt(m_p1['conc_panier_order_kg'])} kg)")

# === Analysis: trend direction ===
print("\n" + "="*80)
print("ANALYSE : LES NOUVELLES MÉTRIQUES CONFIRMENT-ELLES LES CONCLUSIONS ?")
print("="*80)

# SOJA trend per metric
soja_vol_trend = [results[p['name']]['soja_t'] for p in PERIODS]
soja_cli_trend = [results[p['name']]['n_soja_clients'] for p in PERIODS]
soja_cmd_trend = [results[p['name']]['n_soja_orders'] for p in PERIODS]
soja_pc_trend = [results[p['name']]['soja_panier_client_kg'] for p in PERIODS]
soja_po_trend = [results[p['name']]['soja_panier_order_kg'] for p in PERIODS]

# CONC trend per metric
conc_vol_trend = [results[p['name']]['conc_t'] for p in PERIODS]
conc_cli_trend = [results[p['name']]['n_conc_clients'] for p in PERIODS]
conc_cmd_trend = [results[p['name']]['n_conc_orders'] for p in PERIODS]
conc_pc_trend = [results[p['name']]['conc_panier_client_kg'] for p in PERIODS]
conc_po_trend = [results[p['name']]['conc_panier_order_kg'] for p in PERIODS]

def trend_direction(values):
    """Return 'down', 'up', or 'mixed' based on consecutive week-over-week variations."""
    if len(values) < 2: return 'unknown'
    directions = []
    for i in range(1, len(values)):
        if values[i] < values[i-1]:
            directions.append('down')
        elif values[i] > values[i-1]:
            directions.append('up')
        else:
            directions.append('flat')
    if all(d == 'down' for d in directions):
        return 'down (consistent)'
    if all(d == 'up' for d in directions):
        return 'up (consistent)'
    return f'mixed ({directions})'

print(f"\n--- SOJA — Tendance par métrique ---")
print(f"  Volume (t):        {trend_direction(soja_vol_trend)}")
print(f"  Clients:           {trend_direction(soja_cli_trend)}")
print(f"  Commandes:         {trend_direction(soja_cmd_trend)}")
print(f"  Panier/client (kg): {trend_direction(soja_pc_trend)}")
print(f"  Panier/cmd (kg):    {trend_direction(soja_po_trend)}")

print(f"\n--- CONCENTRÉS — Tendance par métrique ---")
print(f"  Volume (t):        {trend_direction(conc_vol_trend)}")
print(f"  Clients:           {trend_direction(conc_cli_trend)}")
print(f"  Commandes:         {trend_direction(conc_cmd_trend)}")
print(f"  Panier/client (kg): {trend_direction(conc_pc_trend)}")
print(f"  Panier/cmd (kg):    {trend_direction(conc_po_trend)}")

# Save to JSON
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
    'trends': {
        'soja': {
            'volume': trend_direction(soja_vol_trend),
            'clients': trend_direction(soja_cli_trend),
            'orders': trend_direction(soja_cmd_trend),
            'panier_client': trend_direction(soja_pc_trend),
            'panier_order': trend_direction(soja_po_trend),
        },
        'conc': {
            'volume': trend_direction(conc_vol_trend),
            'clients': trend_direction(conc_cli_trend),
            'orders': trend_direction(conc_cmd_trend),
            'panier_client': trend_direction(conc_pc_trend),
            'panier_order': trend_direction(conc_po_trend),
        },
    },
    'cumul_w3_to_p1': {
        'soja': {
            'volume': (m_p1['soja_t']/m_w3['soja_t']-1)*100,
            'clients': (m_p1['n_soja_clients']/m_w3['n_soja_clients']-1)*100,
            'orders': (m_p1['n_soja_orders']/m_w3['n_soja_orders']-1)*100,
            'panier_client': (m_p1['soja_panier_client_kg']/m_w3['soja_panier_client_kg']-1)*100,
            'panier_order': (m_p1['soja_panier_order_kg']/m_w3['soja_panier_order_kg']-1)*100,
        },
        'conc': {
            'volume': (m_p1['conc_t']/m_w3['conc_t']-1)*100,
            'clients': (m_p1['n_conc_clients']/m_w3['n_conc_clients']-1)*100,
            'orders': (m_p1['n_conc_orders']/m_w3['n_conc_orders']-1)*100,
            'panier_client': (m_p1['conc_panier_client_kg']/m_w3['conc_panier_client_kg']-1)*100,
            'panier_order': (m_p1['conc_panier_order_kg']/m_w3['conc_panier_order_kg']-1)*100,
        },
    },
}

with open('/home/z/my-project/scripts/sept_trend_extended.json', 'w', encoding='utf-8') as f:
    json.dump(out, f, indent=2, ensure_ascii=False, default=str)
print(f"\n=== SAVED: /home/z/my-project/scripts/sept_trend_extended.json ===")
