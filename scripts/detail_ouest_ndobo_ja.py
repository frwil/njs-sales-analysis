"""Détail des achats des 16 clients Ouest transférés à NDOBO en juillet-août 2026.
Vue mois par mois × famille d'articles (TOURTEAUX/SOJA, CONCENTRES, AUTRES).
"""
import openpyxl
from collections import defaultdict, Counter
import json

# Sources
JUIN_FILE = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (9).xlsx"  # juillet
AOUT_FILE = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (18).xlsx"  # août au 17/08

ETATS_INCLUS = {"Livrée", "Validée", "En cours"}
NDOBO = "AGENCE NDOBO"

# Les 16 clients identifiés
CLIENTS_OUEST_NDOBO = {
    "CU2407-00783 -  METAFE GNITEYO SONYA M",
    "CU2407-00875 -  NOUMSI BLAISE",
    "CU2407-00290 -  M. SONFACK FERDINAND",
    "CU2407-00558 -  DUBOIS SARL",
    "CU2407-00676 -  KAMGO",
    "CU2604-14447 -  KOUOKAM ALPHONSE STEPHANE",
    "CU2407-00288 -  LIDEL SARL",
    "CU2407-00815 -  NGADEU BRICE",
    "CU2407-00658 -  INGRID BANDJA",
    "MOUKAM JOSEPH",
    "CU2509-14118 -  FEUDJO JEAN CLOVIS (ETS ESPOIR DU CAMEROUN)",
    "CU2407-00780 -  MELI JEAN",
    "CU2408-01042 -  CLIENT AGROCAM (INTERNAL)",
    "CU2407-00989 -  NSUH EDWIN CHEO (ETS GILLI BROWN CONSTRUCTION)",
    "CU2605-14465 -  FONDA CHRISTIAN CHO",
    "CU2407-00266 -  EMMANUEL WIRSENYUY",
}

# Product refs → famille
SOJA_REFS = {'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25}
CONC_REFS = {
    'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
    'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50,
    'CB200': 25, 'CB100': 25,  # Chick/Piglet booster
    'PB100': 25, 'PB200': 25,
    'DB100': 25, 'DB200': 25,
}
PREMIX_REFS = {'PX101': 25, 'PX102': 25, 'PX103': 25, 'PX104': 25, 'PX105': 25}
MAIS_REFS = {'M1051': 50, 'M1052': 50}
INGREDIENT_REFS = {
    'B100': 25, 'E101': 25, 'I105': 25, 'B1001': 1, 'B1003': 5, 'B1004': 25,
    'E1011': 1, 'E1013': 5, 'E1014': 25,
    'I1051': 1, 'I1053': 5, 'I1054': 25,
}


def get_family(ref):
    """Map ref to product family."""
    if ref in SOJA_REFS:
        return 'TOURTEAUX/SOJA'
    if ref in CONC_REFS:
        return 'CONCENTRES'
    if ref in PREMIX_REFS:
        return 'PREMIX'
    if ref in MAIS_REFS:
        return 'MAIS'
    if ref in INGREDIENT_REFS:
        return 'INGREDIENTS'
    return 'AUTRES'


def get_sacs(ref, qte):
    """Convert quantity to sacs eq 50kg based on ref."""
    if ref in SOJA_REFS:
        return qte * SOJA_REFS[ref] / 50
    if ref in CONC_REFS:
        return qte * CONC_REFS[ref] / 50
    if ref in PREMIX_REFS:
        return qte * PREMIX_REFS[ref] / 50
    if ref in MAIS_REFS:
        return qte * MAIS_REFS[ref] / 50
    if ref in INGREDIENT_REFS:
        return qte * INGREDIENT_REFS[ref] / 50
    return 0  # AUTRES non comptabilisés en sacs


def load_rows(path, month_filter):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
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
            col_idx['etat'] = 13
        if 'agence' not in col_idx:
            col_idx['agence'] = 15
        rows = list(ws.iter_rows(min_row=3, values_only=True))
        for r in rows:
            if not r or len(r) <= max(col_idx['etat'], col_idx['agence']):
                continue
            if r[0] == 'Total':
                continue
            etat = str(r[col_idx['etat']]).strip() if r[col_idx['etat']] else ""
            if etat not in ETATS_INCLUS:
                continue
            date_str = str(r[6])[:10]
            if f'/{month_filter}/2026' not in date_str:
                continue
            out.append({
                'date': date_str,
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


# Load juillet + août rows
print("Loading juillet (file 9)...")
juil_rows = load_rows(JUIN_FILE, '07')
print(f"  {len(juil_rows)} rows in juillet")

print("Loading août (file 18)...")
aout_rows = load_rows(AOUT_FILE, '08')
aout_rows = [r for r in aout_rows if r['date'] != '18/08/2026']
print(f"  {len(aout_rows)} rows in août (au 17/08)")

# Filter NDOBO only + target clients
juil_ndobo = [r for r in juil_rows if r['agence'] == NDOBO and r['client'] in CLIENTS_OUEST_NDOBO]
aout_ndobo = [r for r in aout_rows if r['agence'] == NDOBO and r['client'] in CLIENTS_OUEST_NDOBO]
print(f"\n  NDOBO Juil (16 clients Ouest): {len(juil_ndobo)} lignes")
print(f"  NDOBO Août (16 clients Ouest): {len(aout_ndobo)} lignes")

# Build per-client × per-month × per-family breakdown
clients_data = defaultdict(lambda: {
    '07': defaultdict(lambda: {'sacs': 0, 'qte': 0, 'cmds': set(), 'ca_ttc': 0, 'refs': Counter(), 'produits': Counter(), 'dates': set(), 'etats': Counter()}),
    '08': defaultdict(lambda: {'sacs': 0, 'qte': 0, 'cmds': set(), 'ca_ttc': 0, 'refs': Counter(), 'produits': Counter(), 'dates': set(), 'etats': Counter()}),
})

for r in juil_ndobo:
    family = get_family(r['ref'])
    sacs = get_sacs(r['ref'], r['qte'])
    c = clients_data[r['client']]['07'][family]
    c['sacs'] += sacs
    c['qte'] += r['qte']
    c['cmds'].add(r['cmd_ref'])
    c['ca_ttc'] += r['montant_ttc'] or 0
    c['refs'][r['ref']] += r['qte']
    c['produits'][r['art_desc']] += r['qte']
    c['dates'].add(r['date'])
    c['etats'][r['etat']] += 1

for r in aout_ndobo:
    family = get_family(r['ref'])
    sacs = get_sacs(r['ref'], r['qte'])
    c = clients_data[r['client']]['08'][family]
    c['sacs'] += sacs
    c['qte'] += r['qte']
    c['cmds'].add(r['cmd_ref'])
    c['ca_ttc'] += r['montant_ttc'] or 0
    c['refs'][r['ref']] += r['qte']
    c['produits'][r['art_desc']] += r['qte']
    c['dates'].add(r['date'])
    c['etats'][r['etat']] += 1

# === Print: synthèse par famille et par mois (tous 16 clients confondus) ===
print("\n" + "=" * 100)
print("SYNTHÈSE GLOBALE — 16 CLIENTS OUEST TRANSFÉRÉS À NDOBO")
print("=" * 100)

FAMILLES = ['TOURTEAUX/SOJA', 'CONCENTRES', 'PREMIX', 'INGREDIENTS', 'MAIS', 'AUTRES']

print(f"\n{'Famille':25} {'Juil sacs':>12} {'Juil t':>10} {'Juil CA(M)':>12} {'Août sacs':>12} {'Août t':>10} {'Août CA(M)':>12} {'Total sacs':>12}")
print("-" * 117)
totals = {'07': {'sacs': 0, 't': 0, 'ca': 0}, '08': {'sacs': 0, 't': 0, 'ca': 0}}
for fam in FAMILLES:
    s7 = sum(c['07'][fam]['sacs'] for c in clients_data.values())
    t7 = s7 * 50 / 1000
    ca7 = sum(c['07'][fam]['ca_ttc'] for c in clients_data.values()) / 1e6
    s8 = sum(c['08'][fam]['sacs'] for c in clients_data.values())
    t8 = s8 * 50 / 1000
    ca8 = sum(c['08'][fam]['ca_ttc'] for c in clients_data.values()) / 1e6
    print(f"{fam:25} {s7:>12.0f} {t7:>10.1f} {ca7:>12.1f} {s8:>12.0f} {t8:>10.1f} {ca8:>12.1f} {s7+s8:>12.0f}")
    totals['07']['sacs'] += s7
    totals['07']['t'] += t7
    totals['07']['ca'] += ca7
    totals['08']['sacs'] += s8
    totals['08']['t'] += t8
    totals['08']['ca'] += ca8

print("-" * 117)
print(f"{'TOTAL':25} {totals['07']['sacs']:>12.0f} {totals['07']['t']:>10.1f} {totals['07']['ca']:>12.1f} {totals['08']['sacs']:>12.0f} {totals['08']['t']:>10.1f} {totals['08']['ca']:>12.1f} {totals['07']['sacs']+totals['08']['sacs']:>12.0f}")

# === Print: détail par client ===
print("\n" + "=" * 100)
print("DÉTAIL PAR CLIENT — Achats à NDOBO en juillet-août")
print("=" * 100)

# Sort clients by total sacs (juil+aout) desc
client_totals = []
for client, months in clients_data.items():
    total_sacs = sum(months['07'][f]['sacs'] for f in FAMILLES) + sum(months['08'][f]['sacs'] for f in FAMILLES)
    total_ca = sum(months['07'][f]['ca_ttc'] for f in FAMILLES) + sum(months['08'][f]['ca_ttc'] for f in FAMILLES)
    client_totals.append((client, total_sacs, total_ca, months))
client_totals.sort(key=lambda x: -x[1])

for idx, (client, total_sacs, total_ca, months) in enumerate(client_totals, 1):
    client_short = (client[:65] + '...') if len(str(client)) > 65 else client
    print(f"\n--- {idx}. {client_short} ---")
    print(f"    Total: {total_sacs:.0f} sacs, CA {total_ca/1e6:.1f} M FCFA")
    
    # Print by month and family
    print(f"    {'Mois':6} {'Famille':25} {'Sacs':>10} {'Tonnes':>10} {'CA(M FCFA)':>12} {'Cmds':>6} {'Réfs produits (qté)':>50}")
    print("    " + "-" * 125)
    for month in ['07', '08']:
        month_label = 'JUIL' if month == '07' else 'AOÛT'
        for fam in FAMILLES:
            c = months[month][fam]
            if c['sacs'] > 0 or c['qte'] > 0:
                refs_str = ', '.join(f'{k}({v:.0f})' for k, v in c['refs'].most_common(5))
                print(f"    {month_label:6} {fam:25} {c['sacs']:>10.0f} {c['sacs']*50/1000:>10.1f} {c['ca_ttc']/1e6:>12.1f} {len(c['cmds']):>6} {refs_str:>50}")
        # Subtotal month
        month_sacs = sum(months[month][f]['sacs'] for f in FAMILLES)
        month_ca = sum(months[month][f]['ca_ttc'] for f in FAMILLES)
        if month_sacs > 0:
            print(f"    {month_label:6} {'SOUS-TOTAL':25} {month_sacs:>10.0f} {month_sacs*50/1000:>10.1f} {month_ca/1e6:>12.1f}")

# Save detail to JSON
detail = []
for client, months in clients_data.items():
    entry = {'client': str(client), 'months': {}}
    for month in ['07', '08']:
        entry['months'][month] = {}
        for fam in FAMILLES:
            c = months[month][fam]
            if c['sacs'] > 0 or c['qte'] > 0:
                entry['months'][month][fam] = {
                    'sacs': round(c['sacs'], 0),
                    'tonnes': round(c['sacs'] * 50 / 1000, 1),
                    'ca_m_fcfa': round(c['ca_ttc'] / 1e6, 2),
                    'cmds': len(c['cmds']),
                    'refs': dict(c['refs']),
                    'produits': dict(c['produits']),
                    'dates': sorted(c['dates']),
                    'etats': dict(c['etats']),
                }
    detail.append(entry)

# Sort by total sacs
detail.sort(key=lambda x: -sum(m.get('sacs', 0) for m in x['months'].values() for fam in m.values() if isinstance(fam, dict) for k, v in fam.items() if k == 'sacs'))

with open('/home/z/my-project/scripts/ouest_ndobo_detail.json', 'w', encoding='utf-8') as f:
    json.dump(detail, f, indent=2, ensure_ascii=False, default=str)
print(f"\n\nDétail sauvegardé: /home/z/my-project/scripts/ouest_ndobo_detail.json")
