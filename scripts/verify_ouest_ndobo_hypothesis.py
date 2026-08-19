"""Hypothèse: clients Ouest (S1) ayant acheté à NDOBO en juillet-août 2026.
Inclut les états "Livrée", "Validée", "En cours".

S1 = Janvier-Juin 2026 (cumul des extractions disponibles)
Juillet = juillet complet
Août = 01/08 au 17/08
"""
import openpyxl
from collections import defaultdict, Counter
from datetime import datetime
import json

# Sources
S1_FILE = "/home/z/my-project/upload/ventes janv a juin 2026.xlsx"  # S1 Jan-Juin 2026 (6 sheets)
JUIN_FILE = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (9).xlsx"  # juillet
AOUT_FILE = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (18).xlsx"  # août au 17/08

# États à inclure (validé + en cours + livré)
ETATS_INCLUS = {"Livrée", "Validée", "En cours"}

# Agences Ouest
OUEST_AGENCES = {"AGENCE FAMLA", "AGENCE DJELENG", "AGENCE DE BAMENDA - DEPOT MBOUDA"}
NDOBO = "AGENCE NDOBO"

# Product refs
SOJA_REFS = {'T102': 'Tourteau soja 50kg', 'T1021': 'Tourteau soja 1kg', 'T1023': 'Tourteau soja 5kg', 'T1024': 'Tourteau soja 25kg'}
CONC_REFS = {
    'C101': 'C101 BELGO Chair 50kg', 'C102': 'C102 BELGO Ponte 50kg', 'C103': 'C103 BELGO Porc 50kg',
    'C104': 'C104 BELGO 10% Chair 50kg', 'C1042': 'C1042 BELGO 10% Chair 1kg', 'C1043': 'C1043 BELGO 10% Chair 5kg', 'C1044': 'C1044 BELGO 10% Chair 25kg',
    'C105': 'C105 BELGO Chair 50kg', 'C1053': 'C1053 BELGO Chair 1kg', 'C1054': 'C1054 BELGO Chair 5kg', 'C1055': 'C1055 BELGO Chair 25kg',
    'C108': 'C108 BELGO 25% Chair 50kg',
}


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
        indices['etat'] = 13
    if 'agence' not in indices:
        indices['agence'] = 15
    return indices


def load_all_rows(path, etats_inclus):
    """Load all rows with the given states from a multi-sheet ERP file.
    Returns list of dicts."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # Detect header on row 2 (row 1 is title)
        header = None
        for row in ws.iter_rows(min_row=2, max_row=2, values_only=True):
            header = row
            break
        if not header:
            continue
        col_idx = {}
        for i, h in enumerate(header):
            if h == 'État':
                col_idx['etat'] = i
            elif h == 'agence':
                col_idx['agence'] = i
        if 'etat' not in col_idx:
            col_idx['etat'] = 13 if len(header) <= 16 else 15
        if 'agence' not in col_idx:
            col_idx['agence'] = 15 if len(header) <= 16 else 17
        rows = list(ws.iter_rows(min_row=3, values_only=True))
        for r in rows:
            if not r or len(r) <= max(col_idx['etat'], col_idx['agence']):
                continue
            if r[0] == 'Total':
                continue
            etat = str(r[col_idx['etat']]).strip() if r[col_idx['etat']] else ""
            if etat not in etats_inclus:
                continue
            out.append({
                'date': str(r[6])[:10],
                'client': r[5],
                'agence': r[col_idx['agence']],
                'ref': r[0],
                'qte': r[2] or 0,
                'etat': etat,
                'cmd_ref': r[3],
                'art_desc': r[1],
                'montant_ttc': r[9] if len(r) > 9 else 0,
            })
    return out


def get_client_agence_freq(rows, target_agence):
    """For each client, compute total volume (sacs) bought at target_agence."""
    client_data = defaultdict(lambda: {
        'qte_sacs': 0, 'cmds': set(), 'products': Counter(),
        'refs': Counter(), 'dates': [], 'agences_seen': set(),
        'sample_cmd': None, 'total_ttc': 0,
    })
    for r in rows:
        if not r['client']:
            continue
        client = r['client']
        agence = r['agence']
        client_data[client]['agences_seen'].add(agence)
        if agence == target_agence:
            ref = r['ref']
            qte = r['qte']
            # Convert to sacs eq 50kg
            if ref in SOJA_REFS:
                sacs = qte * {50:1, 1:1/50, 5:5/50, 25:25/50}.get({'T102':50,'T1021':1,'T1023':5,'T1024':25}[ref], 1)
            elif ref in CONC_REFS:
                sacs = qte * {50:1, 1:1/50, 5:5/50, 25:25/50}.get(
                    {'C101':50,'C102':50,'C103':50,'C104':50,'C1042':1,'C1043':5,'C1044':25,
                     'C105':50,'C1053':1,'C1054':5,'C1055':25,'C108':50}[ref], 1)
            else:
                sacs = 0
            client_data[client]['qte_sacs'] += sacs
            client_data[client]['cmds'].add(r['cmd_ref'])
            client_data[client]['products'][r['art_desc']] += qte
            client_data[client]['refs'][ref] += qte
            client_data[client]['dates'].append(r['date'])
            client_data[client]['total_ttc'] += r['montant_ttc'] or 0
            if not client_data[client]['sample_cmd']:
                client_data[client]['sample_cmd'] = r['cmd_ref']
    return client_data


# === S1: Identifier les clients Ouest (Jan-Juin) ===
print("Loading S1 (Jan-Juin) - ventes janv a juin 2026.xlsx...")
s1_rows = load_all_rows(S1_FILE, ETATS_INCLUS)
print(f"  {len(s1_rows)} rows in S1 (Jan-Juin 2026)")

print("\nLoading juillet (file 9)...")
juil_rows = load_all_rows(JUIN_FILE, ETATS_INCLUS)
juil_rows = [r for r in juil_rows if r['date'] and '/07/2026' in r['date']]
print(f"  Juillet: {len(juil_rows)} rows")

print("\nLoading août (file 18)...")
aout_rows = load_all_rows(AOUT_FILE, ETATS_INCLUS)
aout_rows = [r for r in aout_rows if r['date'] and '/08/2026' in r['date']]
aout_rows = [r for r in aout_rows if r['date'] != '18/08/2026']
print(f"  Août (au 17/08): {len(aout_rows)} rows")

# === Identifier clients Ouest en S1 ===
print("\n=== IDENTIFICATION CLIENTS OUEST EN S1 ===")
s1_ouest_clients = defaultdict(set)  # client -> set of agences Ouest où il a acheté
s1_ouest_volume = defaultdict(lambda: defaultdict(float))  # client -> agence -> sacs
for r in s1_rows:
    if not r['client']:
        continue
    if r['agence'] in OUEST_AGENCES:
        s1_ouest_clients[r['client']].add(r['agence'])
        ref = r['ref']
        qte = r['qte']
        if ref in SOJA_REFS:
            sacs = qte * {50:1, 1:1/50, 5:5/50, 25:25/50}.get({'T102':50,'T1021':1,'T1023':5,'T1024':25}[ref], 1)
        elif ref in CONC_REFS:
            sacs = qte * {50:1, 1:1/50, 5:5/50, 25:25/50}.get(
                {'C101':50,'C102':50,'C103':50,'C104':50,'C1042':1,'C1043':5,'C1044':25,
                 'C105':50,'C1053':1,'C1054':5,'C1055':25,'C108':50}[ref], 1)
        else:
            sacs = 0
        s1_ouest_volume[r['client']][r['agence']] += sacs

print(f"  Total clients Ouest en S1: {len(s1_ouest_clients)}")

# Détail par agence Ouest
ouest_by_agence = defaultdict(set)
for client, agences in s1_ouest_clients.items():
    for ag in agences:
        ouest_by_agence[ag].add(client)
for ag, clients in ouest_by_agence.items():
    print(f"    {ag}: {len(clients)} clients S1")

# === Identifier les clients Ouest S1 qui ont acheté à NDOBO en juillet ou août ===
print("\n=== CLIENTS OUEST S1 AYANT ACHETE A NDOBO EN JUILLET-AOUT ===")

# Get clients NDOBO juillet
ndobo_juil_clients = set()
ndobo_juil_data = defaultdict(lambda: {'cmds': set(), 'qte_sacs': 0, 'products': Counter(), 'refs': Counter(), 'dates': [], 'montant_ttc': 0})
for r in juil_rows:
    if r['agence'] == NDOBO and r['client']:
        ndobo_juil_clients.add(r['client'])
        ref = r['ref']
        qte = r['qte']
        if ref in SOJA_REFS:
            sacs = qte * {50:1, 1:1/50, 5:5/50, 25:25/50}.get({'T102':50,'T1021':1,'T1023':5,'T1024':25}[ref], 1)
        elif ref in CONC_REFS:
            sacs = qte * {50:1, 1:1/50, 5:5/50, 25:25/50}.get(
                {'C101':50,'C102':50,'C103':50,'C104':50,'C1042':1,'C1043':5,'C1044':25,
                 'C105':50,'C1053':1,'C1054':5,'C1055':25,'C108':50}[ref], 1)
        else:
            sacs = 0
        ndobo_juil_data[r['client']]['cmds'].add(r['cmd_ref'])
        ndobo_juil_data[r['client']]['qte_sacs'] += sacs
        ndobo_juil_data[r['client']]['products'][r['art_desc']] += qte
        ndobo_juil_data[r['client']]['refs'][ref] += qte
        ndobo_juil_data[r['client']]['dates'].append(r['date'])
        ndobo_juil_data[r['client']]['montant_ttc'] += r['montant_ttc'] or 0
print(f"  Total clients NDOBO en juillet: {len(ndobo_juil_clients)}")

# Get clients NDOBO août
ndobo_aout_clients = set()
ndobo_aout_data = defaultdict(lambda: {'cmds': set(), 'qte_sacs': 0, 'products': Counter(), 'refs': Counter(), 'dates': [], 'montant_ttc': 0})
for r in aout_rows:
    if r['agence'] == NDOBO and r['client']:
        ndobo_aout_clients.add(r['client'])
        ref = r['ref']
        qte = r['qte']
        if ref in SOJA_REFS:
            sacs = qte * {50:1, 1:1/50, 5:5/50, 25:25/50}.get({'T102':50,'T1021':1,'T1023':5,'T1024':25}[ref], 1)
        elif ref in CONC_REFS:
            sacs = qte * {50:1, 1:1/50, 5:5/50, 25:25/50}.get(
                {'C101':50,'C102':50,'C103':50,'C104':50,'C1042':1,'C1043':5,'C1044':25,
                 'C105':50,'C1053':1,'C1054':5,'C1055':25,'C108':50}[ref], 1)
        else:
            sacs = 0
        ndobo_aout_data[r['client']]['cmds'].add(r['cmd_ref'])
        ndobo_aout_data[r['client']]['qte_sacs'] += sacs
        ndobo_aout_data[r['client']]['products'][r['art_desc']] += qte
        ndobo_aout_data[r['client']]['refs'][ref] += qte
        ndobo_aout_data[r['client']]['dates'].append(r['date'])
        ndobo_aout_data[r['client']]['montant_ttc'] += r['montant_ttc'] or 0
print(f"  Total clients NDOBO en août: {len(ndobo_aout_clients)}")

# Cross: Ouest S1 ∩ NDOBO juillet/août
ouest_ndobo_juil = s1_ouest_clients.keys() & ndobo_juil_clients
ouest_ndobo_aout = s1_ouest_clients.keys() & ndobo_aout_clients
ouest_ndobo_both = ouest_ndobo_juil & ouest_ndobo_aout
ouest_ndobo_either = ouest_ndobo_juil | ouest_ndobo_aout

print(f"\n  Clients Ouest S1 ayant acheté à NDOBO en juillet: {len(ouest_ndobo_juil)}")
print(f"  Clients Ouest S1 ayant acheté à NDOBO en août: {len(ouest_ndobo_aout)}")
print(f"  Clients Ouest S1 ayant acheté à NDOBO en juillet ET août: {len(ouest_ndobo_both)}")
print(f"  Clients Ouest S1 ayant acheté à NDOBO en juillet OU août: {len(ouest_ndobo_either)}")

# Volume total concerné (juillet + août à NDOBO)
total_sacs_juil = sum(ndobo_juil_data[c]['qte_sacs'] for c in ouest_ndobo_juil)
total_sacs_aout = sum(ndobo_aout_data[c]['qte_sacs'] for c in ouest_ndobo_aout)
total_ttc_juil = sum(ndobo_juil_data[c]['montant_ttc'] for c in ouest_ndobo_juil)
total_ttc_aout = sum(ndobo_aout_data[c]['montant_ttc'] for c in ouest_ndobo_aout)
print(f"\n  Volume total NDOBO (clients Ouest S1):")
print(f"    Juillet: {total_sacs_juil:.0f} sacs, {total_sacs_juil*50/1000:.1f} t, CA {total_ttc_juil/1e6:.0f} M FCFA")
print(f"    Août:    {total_sacs_aout:.0f} sacs, {total_sacs_aout*50/1000:.1f} t, CA {total_ttc_aout/1e6:.0f} M FCFA")
print(f"    Total:   {total_sacs_juil+total_sacs_aout:.0f} sacs, {(total_sacs_juil+total_sacs_aout)*50/1000:.1f} t, CA {(total_ttc_juil+total_ttc_aout)/1e6:.0f} M FCFA")

# Comparaison: ces clients, qu'ont-ils acheté à FAMLA en S1 vs juillet-août ?
print("\n=== COMPARE: ACHAT FAMLA S1 vs JUILLET-AOUT ===")
# Build S1 FAMLA volume per client
s1_famla_volume = defaultdict(float)
for r in s1_rows:
    if r['agence'] == 'AGENCE FAMLA' and r['client']:
        ref = r['ref']
        qte = r['qte']
        if ref in SOJA_REFS:
            sacs = qte * {50:1, 1:1/50, 5:5/50, 25:25/50}.get({'T102':50,'T1021':1,'T1023':5,'T1024':25}[ref], 1)
        elif ref in CONC_REFS:
            sacs = qte * {50:1, 1:1/50, 5:5/50, 25:25/50}.get(
                {'C101':50,'C102':50,'C103':50,'C104':50,'C1042':1,'C1043':5,'C1044':25,
                 'C105':50,'C1053':1,'C1054':5,'C1055':25,'C108':50}[ref], 1)
        else:
            sacs = 0
        s1_famla_volume[r['client']] += sacs

# Build Juil+Août FAMLA volume per client
ja_famla_volume = defaultdict(float)
for r in juil_rows + aout_rows:
    if r['agence'] == 'AGENCE FAMLA' and r['client']:
        ref = r['ref']
        qte = r['qte']
        if ref in SOJA_REFS:
            sacs = qte * {50:1, 1:1/50, 5:5/50, 25:25/50}.get({'T102':50,'T1021':1,'T1023':5,'T1024':25}[ref], 1)
        elif ref in CONC_REFS:
            sacs = qte * {50:1, 1:1/50, 5:5/50, 25:25/50}.get(
                {'C101':50,'C102':50,'C103':50,'C104':50,'C1042':1,'C1043':5,'C1044':25,
                 'C105':50,'C1053':1,'C1054':5,'C1055':25,'C108':50}[ref], 1)
        else:
            sacs = 0
        ja_famla_volume[r['client']] += sacs

# Pour les clients Ouest ayant acheté à NDOBO en Juil+Août, vérifier s'ils ont réduit leurs achats à FAMLA
print(f"\n{'Client':60} {'Sacs FAMLA S1':>15} {'Sacs FAMLA JA':>15} {'Sacs NDOBO JA':>15}")
print("-" * 110)
total_famla_s1 = 0
total_famla_ja = 0
total_ndobo_ja = 0
for client in sorted(ouest_ndobo_either, key=lambda c: -(s1_famla_volume[c])):
    s1 = s1_famla_volume[client]
    ja_famla = ja_famla_volume[client]
    ja_ndobo = ndobo_juil_data[client]['qte_sacs'] + ndobo_aout_data[client]['qte_sacs']
    total_famla_s1 += s1
    total_famla_ja += ja_famla
    total_ndobo_ja += ja_ndobo
    client_short = (client[:55] + '...') if len(str(client)) > 55 else client
    print(f"{str(client_short):60} {s1:>15.0f} {ja_famla:>15.0f} {ja_ndobo:>15.0f}")

print("-" * 110)
print(f"{'TOTAL':60} {total_famla_s1:>15.0f} {total_famla_ja:>15.0f} {total_ndobo_ja:>15.0f}")
print(f"\nBaisse FAMLA S1 → Juil+Août: {total_famla_s1-total_famla_ja:.0f} sacs ({(total_famla_s1-total_famla_ja)/total_famla_s1*100:.0f}%)")
print(f"Volume transféré à NDOBO: {total_ndobo_ja:.0f} sacs")
print(f"Ratio NDOBO / baisse FAMLA: {total_ndobo_ja/(total_famla_s1-total_famla_ja)*100:.0f}%" if total_famla_s1 > total_famla_ja else "")

# Détail des produits achetés à NDOBO par ces clients
print("\n=== PRODUITS ACHETES A NDOBO PAR LES CLIENTS OUEST ===")
all_products = Counter()
for c in ouest_ndobo_either:
    for p, q in ndobo_juil_data[c]['products'].items():
        all_products[p] += q
    for p, q in ndobo_aout_data[c]['products'].items():
        all_products[p] += q

print(f"{'Produit':50} {'Qté totale':>15}")
print("-" * 70)
for p, q in all_products.most_common():
    print(f"{str(p)[:50]:50} {q:>15.0f}")

# Save detail to JSON
detail = []
for client in ouest_ndobo_either:
    detail.append({
        'client': str(client),
        's1_famla_sacs': round(s1_famla_volume[client], 0),
        's1_djeleng_sacs': round(s1_ouest_volume[client].get('AGENCE DJELENG', 0), 0),
        's1_mbouda_sacs': round(s1_ouest_volume[client].get('AGENCE DE BAMENDA - DEPOT MBOUDA', 0), 0),
        'ja_famla_sacs': round(ja_famla_volume[client], 0),
        'ndobo_juil_sacs': round(ndobo_juil_data[client]['qte_sacs'], 0),
        'ndobo_juil_cmds': len(ndobo_juil_data[client]['cmds']),
        'ndobo_juil_ca_m': round(ndobo_juil_data[client]['montant_ttc']/1e6, 1),
        'ndobo_aout_sacs': round(ndobo_aout_data[client]['qte_sacs'], 0),
        'ndobo_aout_cmds': len(ndobo_aout_data[client]['cmds']),
        'ndobo_aout_ca_m': round(ndobo_aout_data[client]['montant_ttc']/1e6, 1),
        'ndobo_products_juil': dict(ndobo_juil_data[client]['products']),
        'ndobo_products_aout': dict(ndobo_aout_data[client]['products']),
        'ndobo_refs_juil': dict(ndobo_juil_data[client]['refs']),
        'ndobo_refs_aout': dict(ndobo_aout_data[client]['refs']),
        'ndobo_dates_juil': sorted(set(ndobo_juil_data[client]['dates'])),
        'ndobo_dates_aout': sorted(set(ndobo_aout_data[client]['dates'])),
    })

with open('/home/z/my-project/scripts/ouest_ndobo_transfer.json', 'w', encoding='utf-8') as f:
    json.dump(detail, f, indent=2, ensure_ascii=False, default=str)
print(f"\nDétail sauvegardé: /home/z/my-project/scripts/ouest_ndobo_transfer.json")
