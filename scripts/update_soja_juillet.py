"""
Update analyse_soja_juillet.xlsx with the latest extraction (au 27/07/2026).

Source: /home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (4).xlsx
Output: /home/z/my-project/download/analyse_soja_juillet.xlsx (overwritten)

4 sheets:
  T1 - Moy jour soja par agence (01/07 au 27/07, 23j, Livree)
  T2 - Commandes soja 27/07 (avec evolution vs 24/07)
  T3 - Soja-only 27/07 (liste detaillee)
  T4 - Stock soja et rupture (mise a jour avec moy/jour 27/07)
"""
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict, Counter
import datetime
import copy
import json

SRC = '/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (4).xlsx'
PREV_SRC = '/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (3) (7).xlsx'  # au 24/07
OUT = '/home/z/my-project/download/analyse_soja_juillet.xlsx'
PREV_OUT = '/home/z/my-project/download/analyse_soja_juillet.xlsx'  # previous output for comparison

# Product references
SOJA_REFS = {'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25}
CONC_REFS = {
    'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
    'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50
}

# Agence normalization -> (short_name, region)
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

# Region order
REGION_ORDER = ['Ouest', 'Centre', 'Littoral']

# Stock soja
STOCK_SACS = 160000  # 160k sacs de 50kg
STOCK_T = STOCK_SACS * 50 / 1000  # 8000 t


def load_livree(path):
    """Load only Livrée rows from the ERP extraction."""
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb['Sheet 1']
    rows = list(ws.iter_rows(min_row=3, values_only=True))
    out = []
    for r in rows:
        if not r or len(r) < 18:
            continue
        if r[0] == 'Total':
            continue
        if r[15] != 'Livrée':
            continue
        out.append(r)
    return out


def cmd_soja_conc(rows, date_filter=None):
    """For each command, compute soja_kg, conc_kg, soja_sacs_50, conc_sacs_50.
    Returns dict {cmd_ref: {...}} of commands that contain soja OR conc.
    Bundle = both > 0, soja_only = soja > 0 & conc == 0, conc_only = conc > 0 & soja == 0.
    The ratio is total soja / total conc across ALL these commands (portfolio mix)."""
    cmds = defaultdict(lambda: {
        'soja_kg': 0, 'conc_kg': 0, 'soja_sacs_50': 0, 'conc_sacs_50': 0,
        'agence': None, 'client': None, 'etat': None, 'date': None, 'cmd_ref': None
    })
    for r in rows:
        date = str(r[6])[:10]
        if date_filter and date != date_filter:
            continue
        ref = r[0]
        qte = r[2] or 0
        cmd = r[3]
        agence_raw = r[17] or ''
        agence_short, region = AGENCE_MAP.get(agence_raw, (agence_raw, '?'))
        if ref in SOJA_REFS:
            cmds[cmd]['soja_kg'] += qte * SOJA_REFS[ref]
            cmds[cmd]['soja_sacs_50'] += qte * SOJA_REFS[ref] / 50
            cmds[cmd]['agence'] = agence_short
            cmds[cmd]['region'] = region
            cmds[cmd]['client'] = r[5]
            cmds[cmd]['etat'] = r[15]
            cmds[cmd]['date'] = date
            cmds[cmd]['cmd_ref'] = cmd
        if ref in CONC_REFS:
            cmds[cmd]['conc_kg'] += qte * CONC_REFS[ref]
            cmds[cmd]['conc_sacs_50'] += qte * CONC_REFS[ref] / 50
            cmds[cmd]['agence'] = agence_short
            cmds[cmd]['region'] = region
            cmds[cmd]['client'] = r[5]
            cmds[cmd]['etat'] = r[15]
            cmds[cmd]['date'] = date
            cmds[cmd]['cmd_ref'] = cmd
    # Keep commands that have soja OR conc
    return {k: v for k, v in cmds.items() if v['soja_kg'] > 0 or v['conc_kg'] > 0}


def compute_t1(rows, end_date_str):
    """T1: Moyenne journaliere soja par agence/region, 01/07 to end_date."""
    # Days elapsed (lun-sam)
    end_d = int(end_date_str[:2])
    days_elapsed = 0
    for d in range(1, end_d + 1):
        dt = datetime.date(2026, 7, d)
        if dt.weekday() < 6:
            days_elapsed += 1

    # Total soja per agence
    soja_by_agence = defaultdict(lambda: {'kg': 0, 'agence': None, 'region': None})
    for r in rows:
        ref = r[0]
        qte = r[2] or 0
        agence_raw = r[17] or ''
        agence_short, region = AGENCE_MAP.get(agence_raw, (agence_raw, '?'))
        if ref in SOJA_REFS:
            soja_by_agence[agence_short]['kg'] += qte * SOJA_REFS[ref]
            soja_by_agence[agence_short]['agence'] = agence_short
            soja_by_agence[agence_short]['region'] = region

    # Build sorted list (sorted by total desc)
    items = [(v['agence'], v['region'], v['kg'] / 1000) for v in soja_by_agence.values()]
    items.sort(key=lambda x: -x[2])

    # Group by region for subtotals (in kg, not converted back)
    by_region = defaultdict(lambda: {'kg': 0})
    for ag, reg, t in items:
        by_region[reg]['kg'] += t * 1000  # convert t back to kg
        by_region[reg]['region'] = reg

    total_kg = sum(v['kg'] for v in soja_by_agence.values())

    return {
        'days_elapsed': days_elapsed,
        'end_date': end_date_str,
        'items': items,
        'regions': [(reg, by_region[reg]['kg']) for reg in REGION_ORDER if reg in by_region],
        'total_kg': total_kg,
        'total_t': total_kg / 1000,
        'moy_jour_t': total_kg / 1000 / days_elapsed,
        'moy_jour_sacs': total_kg / 50 / days_elapsed,
    }


def compute_t2(rows, date_str):
    """T2: Commandes soja du jour par agence.
    cmds = commandes avec soja (soja_only + bundle).
    bundle = commandes avec soja ET conc.
    soja_only = commandes avec soja mais sans conc.
    ratio = total soja / total conc (toutes commandes incl. conc-only - portfolio mix)."""
    cmds = cmd_soja_conc(rows, date_filter=date_str)
    by_agence = defaultdict(lambda: {
        'cmds_soja': 0, 'bundle': 0, 'soja_only': 0,
        'soja_sacs_50': 0, 'conc_sacs_50': 0, 'region': None
    })
    for c in cmds.values():
        ag = c['agence']
        by_agence[ag]['region'] = c['region']
        # Accumulate totals (includes conc-only commands)
        by_agence[ag]['soja_sacs_50'] += c['soja_sacs_50']
        by_agence[ag]['conc_sacs_50'] += c['conc_sacs_50']
        if c['soja_kg'] > 0:
            by_agence[ag]['cmds_soja'] += 1
            if c['conc_kg'] > 0:
                by_agence[ag]['bundle'] += 1
            else:
                by_agence[ag]['soja_only'] += 1

    items = []
    for ag, v in sorted(by_agence.items(), key=lambda x: -x[1]['cmds_soja']):
        # Skip agencies with no soja commands (conc-only)
        if v['cmds_soja'] == 0:
            continue
        ratio = v['soja_sacs_50'] / v['conc_sacs_50'] if v['conc_sacs_50'] > 0 else float('inf')
        ratio_str = f'{ratio:.1f}:1' if ratio != float('inf') else '∞'
        items.append({
            'agence': ag, 'region': v['region'], 'cmds': v['cmds_soja'],
            'bundle': v['bundle'], 'soja_only': v['soja_only'],
            'ratio': ratio, 'ratio_str': ratio_str,
            'soja_sacs_50': v['soja_sacs_50'], 'conc_sacs_50': v['conc_sacs_50'],
        })

    # Total: includes ALL cmds (soja-only + bundle + conc-only) for the ratio
    total_soja_all = sum(c['soja_sacs_50'] for c in cmds.values())
    total_conc_all = sum(c['conc_sacs_50'] for c in cmds.values())
    total = {
        'cmds': sum(i['cmds'] for i in items),  # soja cmds only
        'bundle': sum(i['bundle'] for i in items),
        'soja_only': sum(i['soja_only'] for i in items),
        'soja_sacs_50': total_soja_all,  # all
        'conc_sacs_50': total_conc_all,  # all (incl conc-only)
    }
    total['ratio'] = total['soja_sacs_50'] / total['conc_sacs_50'] if total['conc_sacs_50'] > 0 else float('inf')
    total['ratio_str'] = f"{total['ratio']:.1f}:1" if total['ratio'] != float('inf') else '∞'

    # Distribution of ratios per cmd (bundle only)
    ratios = []
    for c in cmds.values():
        if c['soja_kg'] > 0 and c['conc_sacs_50'] > 0:
            ratios.append(c['soja_sacs_50'] / c['conc_sacs_50'])
    dist = {
        '<= 3:1': sum(1 for r in ratios if r <= 3),
        '3-5:1': sum(1 for r in ratios if 3 < r <= 5),
        '5-10:1': sum(1 for r in ratios if 5 < r <= 10),
        '10-20:1': sum(1 for r in ratios if 10 < r <= 20),
        '> 20:1': sum(1 for r in ratios if r > 20),
    }
    total_ratios = sum(dist.values())
    dist_pct = {k: (v, f'{v/total_ratios*100:.0f}%' if total_ratios > 0 else '0%') for k, v in dist.items()}

    return {
        'date': date_str,
        'items': items,
        'total': total,
        'dist': dist,
        'dist_pct': dist_pct,
        'n_ratios': total_ratios,
    }


def compute_t3(rows, date_str):
    """T3: List of soja-only commands on date (commands with soja but NO conc)."""
    cmds = cmd_soja_conc(rows, date_filter=date_str)
    soja_only = [c for c in cmds.values() if c['soja_kg'] > 0 and c['conc_kg'] == 0]
    soja_only.sort(key=lambda c: -c['soja_kg'])
    return {
        'date': date_str,
        'items': soja_only,
        'total_kg': sum(c['soja_kg'] for c in soja_only),
        'total_sacs': sum(c['soja_sacs_50'] for c in soja_only),
    }


def add_business_days(start_date, n_days):
    """Add n_days business days (Mon-Sat) to start_date. start_date counts as day 0."""
    if n_days <= 0:
        return start_date
    added = 0
    current = start_date
    while added < n_days:
        current += datetime.timedelta(days=1)
        if current.weekday() < 6:  # Mon-Sat
            added += 1
    return current


def compute_t4(t1, start_date_str='27/07/2026', price_hike_date='23/07/2026'):
    """T4: Stock soja et rupture (with moy/jour from T1).
    Jours stock = business days (lun-sam). Date rupture = start + business days."""
    moy_sacs = t1['moy_jour_sacs']
    moy_t = t1['moy_jour_t']
    # Round to nearest whole day (matches previous convention)
    days_realiste = round(STOCK_SACS / moy_sacs)
    days_plus20 = round(STOCK_SACS / (moy_sacs * 1.2))
    days_moins20 = round(STOCK_SACS / (moy_sacs * 0.8))

    start = datetime.date(2026, 7, int(start_date_str[:2]))
    rupture_realiste = add_business_days(start, days_realiste)
    rupture_plus20 = add_business_days(start, days_plus20)
    rupture_moins20 = add_business_days(start, days_moins20)

    return {
        'moy_sacs': moy_sacs,
        'moy_t': moy_t,
        'scenarios': [
            {'name': 'Realiste', 'sacs_jour': round(moy_sacs), 't_jour': moy_t,
             'jours': days_realiste, 'date': rupture_realiste.strftime('%d/%m/%Y')},
            {'name': 'Acceleration +20%', 'sacs_jour': round(moy_sacs * 1.2),
             't_jour': moy_t * 1.2, 'jours': days_plus20, 'date': rupture_plus20.strftime('%d/%m/%Y')},
            {'name': 'Ralentissement -20%', 'sacs_jour': round(moy_sacs * 0.8),
             't_jour': moy_t * 0.8, 'jours': days_moins20, 'date': rupture_moins20.strftime('%d/%m/%Y')},
        ],
        'stock_sacs': STOCK_SACS,
        'stock_t': STOCK_T,
        'ventes_juillet_sacs': round(t1['total_kg'] / 50),
        'ventes_juillet_t': t1['total_t'],
        'moy_jour_sacs': round(moy_sacs),
        'moy_jour_t': moy_t,
        'jours_ouvrables': '6j/sem (lun-sam)',
        'depart': start_date_str,
        'price_hike_date': price_hike_date,
    }


# Style helpers
HEAD_FILL = PatternFill('solid', fgColor='1F4E78')
HEAD_FONT = Font(bold=True, color='FFFFFF', size=11)
SUBHEAD_FILL = PatternFill('solid', fgColor='D9E1F2')
SUBHEAD_FONT = Font(bold=True, color='1F4E78', size=11)
TOTAL_FILL = PatternFill('solid', fgColor='FFF2CC')
TOTAL_FONT = Font(bold=True, size=11)
EVOL_FILL = PatternFill('solid', fgColor='E2EFDA')
EVOL_FONT = Font(italic=True, color='375623', size=10)
RED_FILL = PatternFill('solid', fgColor='FCE4D6')
RED_FONT = Font(color='9C0006', size=10)
THIN = Side(border_style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header_row(ws, row, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER


def style_total_row(ws, row, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = TOTAL_FILL
        cell.font = TOTAL_FONT
        cell.border = BORDER


def write_t1(ws, t1, t1_prev):
    ws['A1'] = 'TABLEAU 1 - Moyenne journaliere SOJA juillet par agence/region'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws['A2'] = f"01/07 au {t1['end_date']} ({t1['days_elapsed']}j). Livree. lun-sam. Mise a jour 27/07/2026."
    ws['A2'].font = Font(italic=True, size=10, color='595959')

    headers = ['Agence', 'Region', 'Total soja (t)', 'Moy/jour (t)', 'Evolution vs 24/07']
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, len(headers))

    row = 5
    for ag, reg, t in t1['items']:
        ws.cell(row=row, column=1, value=ag)
        ws.cell(row=row, column=2, value=reg)
        ws.cell(row=row, column=3, value=round(t, 1))
        ws.cell(row=row, column=4, value=round(t / t1['days_elapsed'], 1))
        # Evolution vs 24/07
        prev_match = next((p for p in t1_prev['items'] if p[0] == ag), None)
        if prev_match:
            prev_t_per_day = prev_match[2] / t1_prev['days_elapsed']
            delta = (t / t1['days_elapsed']) - prev_t_per_day
            if abs(delta) < 0.05:
                evol = 'stable'
            else:
                evol = f"{delta:+.1f} t/j"
            ws.cell(row=row, column=5, value=evol)
            if delta < -0.1:
                ws.cell(row=row, column=5).font = RED_FONT
            elif delta > 0.1:
                ws.cell(row=row, column=5).fill = EVOL_FILL
                ws.cell(row=row, column=5).font = EVOL_FONT
        else:
            ws.cell(row=row, column=5, value='nouveau')
            ws.cell(row=row, column=5).fill = EVOL_FILL
            ws.cell(row=row, column=5).font = EVOL_FONT
        for c in range(1, len(headers) + 1):
            ws.cell(row=row, column=c).border = BORDER
        row += 1

    # Region subtotals
    for reg, kg in t1['regions']:
        ws.cell(row=row, column=1, value=f'TOTAL {reg}')
        ws.cell(row=row, column=2, value=reg)
        ws.cell(row=row, column=3, value=round(kg / 1000, 1))
        ws.cell(row=row, column=4, value=round(kg / 1000 / t1['days_elapsed'], 1))
        for c in range(1, len(headers) + 1):
            ws.cell(row=row, column=c).fill = SUBHEAD_FILL
            ws.cell(row=row, column=c).font = SUBHEAD_FONT
            ws.cell(row=row, column=c).border = BORDER
        row += 1

    # Grand total
    ws.cell(row=row, column=1, value='TOTAL')
    ws.cell(row=row, column=3, value=round(t1['total_t'], 1))
    ws.cell(row=row, column=4, value=round(t1['moy_jour_t'], 1))
    prev_total_moy = t1_prev['moy_jour_t']
    delta = t1['moy_jour_t'] - prev_total_moy
    evol = f"{delta:+.1f} t/j vs {prev_total_moy:.1f}"
    ws.cell(row=row, column=5, value=evol)
    if delta < 0:
        ws.cell(row=row, column=5).font = Font(color='008000', bold=True, size=10)  # green = good (less consumption)
    else:
        ws.cell(row=row, column=5).font = Font(color='C00000', bold=True, size=10)
    style_total_row(ws, row, len(headers))
    row += 2

    # Note about price increase
    ws.cell(row=row, column=1, value='NOTE: Le prix du soja a augmente de +2 000 XAF/sac le 23/07/2026.')
    ws.cell(row=row, column=1).font = Font(italic=True, color='C00000', size=10)
    row += 1
    ws.cell(row=row, column=1, value=f"Moyenne/jour globale: {t1['moy_jour_t']:.1f} t/j ({int(t1['moy_jour_sacs'])} sacs) sur {t1['days_elapsed']}j.")
    ws.cell(row=row, column=1).font = Font(bold=True, size=10)
    row += 1
    ws.cell(row=row, column=1, value=f"Precedent (24/07): {prev_total_moy:.1f} t/j ({int(t1_prev['moy_jour_sacs'])} sacs) sur {t1_prev['days_elapsed']}j.")
    ws.cell(row=row, column=1).font = Font(italic=True, size=10, color='595959')

    # Column widths
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 24


def write_t2(ws, t2, t2_prev):
    ws['A1'] = f"TABLEAU 2 - Commandes SOJA le {t2['date']} par agence"
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws['A2'] = 'Livree. Ratio en sacs 50kg-equivalent. Mise a jour 27/07/2026.'
    ws['A2'].font = Font(italic=True, size=10, color='595959')

    headers = ['Agence', 'Region', 'Cmds', 'Bundle', 'Soja-only', 'Ratio moy', 'Evolution cmdes vs 24/07']
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, len(headers))

    row = 5
    for it in t2['items']:
        ws.cell(row=row, column=1, value=it['agence'])
        ws.cell(row=row, column=2, value=it['region'])
        ws.cell(row=row, column=3, value=it['cmds'])
        ws.cell(row=row, column=4, value=it['bundle'])
        ws.cell(row=row, column=5, value=it['soja_only'])
        ws.cell(row=row, column=6, value=it['ratio_str'])
        # Evolution
        prev_match = next((p for p in t2_prev['items'] if p['agence'] == it['agence']), None)
        if prev_match:
            delta = it['cmds'] - prev_match['cmds']
            if delta == 0:
                evol = 'stable'
            else:
                evol = f"{delta:+d} cmdes"
            ws.cell(row=row, column=7, value=evol)
            if delta < 0:
                ws.cell(row=row, column=7).font = Font(color='008000', size=10)
            elif delta > 0:
                ws.cell(row=row, column=7).font = Font(color='C00000', size=10)
        else:
            ws.cell(row=row, column=7, value='nouveau')
            ws.cell(row=row, column=7).fill = EVOL_FILL
        for c in range(1, len(headers) + 1):
            ws.cell(row=row, column=c).border = BORDER
        row += 1

    # Total
    ws.cell(row=row, column=1, value='TOTAL')
    ws.cell(row=row, column=3, value=t2['total']['cmds'])
    ws.cell(row=row, column=4, value=t2['total']['bundle'])
    ws.cell(row=row, column=5, value=t2['total']['soja_only'])
    ws.cell(row=row, column=6, value=t2['total']['ratio_str'])
    delta_total = t2['total']['cmds'] - t2_prev['total']['cmds']
    evol = f"{delta_total:+d} cmdes vs {t2_prev['total']['cmds']}"
    ws.cell(row=row, column=7, value=evol)
    if delta_total < 0:
        ws.cell(row=row, column=7).font = Font(color='008000', bold=True, size=10)
    elif delta_total > 0:
        ws.cell(row=row, column=7).font = Font(color='C00000', bold=True, size=10)
    style_total_row(ws, row, len(headers))
    row += 2

    # Distribution ratios
    ws.cell(row=row, column=1, value='Distribution ratios (bundle seulement)')
    ws.cell(row=row, column=1).font = SUBHEAD_FONT
    row += 1
    ws.cell(row=row, column=1, value='Tranche')
    ws.cell(row=row, column=2, value='Nb cmdes')
    ws.cell(row=row, column=3, value='Part')
    style_header_row(ws, row, 3)
    row += 1
    for label, (nb, pct) in t2['dist_pct'].items():
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=nb)
        ws.cell(row=row, column=3, value=pct)
        for c in range(1, 4):
            ws.cell(row=row, column=c).border = BORDER
        row += 1

    row += 2
    ws.cell(row=row, column=1, value=f"NOTE: Le prix du soja a augmente de +2 000 XAF/sac le 23/07/2026 (effet plein visible a partir de ce jour).")
    ws.cell(row=row, column=1).font = Font(italic=True, color='C00000', size=10)
    row += 1
    ws.cell(row=row, column=1, value=f"Ratio global: {t2['total']['ratio_str']} (vs {t2_prev['total']['ratio_str']} le 24/07).")
    ws.cell(row=row, column=1).font = Font(bold=True, size=10)
    row += 2
    ws.cell(row=row, column=1, value="Note: T2/T3 portent sur le 25/07 (dernier jour complet, 66 cmdes soja). Le 27/07 est une extraction matinale (85 lignes au total, 0 commande soja Livree).")
    ws.cell(row=row, column=1).font = Font(italic=True, size=9, color='595959')

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 14
    ws.column_dimensions['G'].width = 28


def write_t3(ws, t3, t3_prev):
    ws['A1'] = f"TABLEAU 3 - SOJA sans concentres le {t3['date']}"
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws['A2'] = 'Livree. Liste detaillee des commandes soja-only.'
    ws['A2'].font = Font(italic=True, size=10, color='595959')

    headers = ['#', 'N cmd', 'Client', 'Agence', 'Region', 'Soja (t)', 'Sacs', 'Etat', 'Nouveau vs 24/07']
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, len(headers))

    prev_refs = set(c['cmd_ref'] for c in t3_prev['items'])
    row = 5
    for i, c in enumerate(t3['items'], 1):
        ws.cell(row=row, column=1, value=i)
        ws.cell(row=row, column=2, value=c['cmd_ref'])
        ws.cell(row=row, column=3, value=c['client'])
        ws.cell(row=row, column=4, value=c['agence'])
        ws.cell(row=row, column=5, value=c['region'])
        ws.cell(row=row, column=6, value=round(c['soja_kg'] / 1000, 2))
        ws.cell(row=row, column=7, value=round(c['soja_sacs_50']))
        ws.cell(row=row, column=8, value=c['etat'])
        is_new = c['cmd_ref'] not in prev_refs
        ws.cell(row=row, column=9, value='oui' if is_new else '')
        if is_new:
            ws.cell(row=row, column=9).fill = EVOL_FILL
            ws.cell(row=row, column=9).font = EVOL_FONT
        for c2 in range(1, len(headers) + 1):
            ws.cell(row=row, column=c2).border = BORDER
        row += 1

    # Total
    ws.cell(row=row, column=3, value='TOTAL')
    ws.cell(row=row, column=6, value=round(t3['total_kg'] / 1000, 1))
    ws.cell(row=row, column=7, value=round(t3['total_sacs']))
    style_total_row(ws, row, len(headers))
    row += 2

    ws.cell(row=row, column=1, value=f"Total soja-only: {len(t3['items'])} cmdes ({round(t3['total_kg']/1000, 1)} t) le {t3['date']}")
    ws.cell(row=row, column=1).font = Font(bold=True, size=10)
    row += 1
    ws.cell(row=row, column=1, value=f"Precedent (24/07): {len(t3_prev['items'])} cmdes ({round(t3_prev['total_kg']/1000, 1)} t)")
    ws.cell(row=row, column=1).font = Font(italic=True, size=10, color='595959')

    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 8
    ws.column_dimensions['H'].width = 10
    ws.column_dimensions['I'].width = 16


def write_t4(ws, t4, t4_prev):
    ws['A1'] = 'TABLEAU 4 - Stock SOJA et date rupture'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws['A2'] = f"160 000 sacs (8 000 t). Moy/jour sur juillet ({t4['moy_jour_t']:.1f} t/j, Livree). Mise a jour 27/07/2026."
    ws['A2'].font = Font(italic=True, size=10, color='595959')

    headers = ['Scenario', 'Sacs/jour', 't/jour', 'Jours stock', 'Date rupture', 'Evolution date rupture']
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, len(headers))

    row = 5
    prev_scn = {s['name']: s for s in t4_prev['scenarios']}
    for s in t4['scenarios']:
        ws.cell(row=row, column=1, value=s['name'])
        ws.cell(row=row, column=2, value=s['sacs_jour'])
        ws.cell(row=row, column=3, value=round(s['t_jour'], 1))
        ws.cell(row=row, column=4, value=s['jours'])
        ws.cell(row=row, column=5, value=s['date'])
        if s['name'] in prev_scn:
            prev = prev_scn[s['name']]
            delta_days = s['jours'] - prev['jours']
            evol = f"{delta_days:+d} jours vs {prev['date']}"
            ws.cell(row=row, column=6, value=evol)
            if delta_days > 0:
                ws.cell(row=row, column=6).font = Font(color='008000', size=10)
            elif delta_days < 0:
                ws.cell(row=row, column=6).font = Font(color='C00000', size=10)
        for c in range(1, len(headers) + 1):
            ws.cell(row=row, column=c).border = BORDER
        row += 1

    row += 1
    ws.cell(row=row, column=1, value='Parametres')
    ws.cell(row=row, column=1).font = SUBHEAD_FONT
    row += 1
    params = [
        ('Stock', f"{t4['stock_sacs']:,} sacs".replace(',', ' '), f"{int(t4['stock_t']):,} t".replace(',', ' ')),
        ('Ventes juillet', f"{t4['ventes_juillet_sacs']:,} sacs".replace(',', ' '), f"{t4['ventes_juillet_t']:.1f} t"),
        ('Moyenne/jour', f"{t4['moy_jour_sacs']:,} sacs".replace(',', ' '), f"{t4['moy_jour_t']:.1f} t"),
        ('Jours ouvrables', t4['jours_ouvrables'], ''),
        ('Depart', t4['depart'], ''),
        ('Hausse prix soja', f"+2 000 XAF/sac le {t4['price_hike_date']}", ''),
    ]
    for label, v1, v2 in params:
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=v1)
        ws.cell(row=row, column=3, value=v2)
        for c in range(1, 4):
            ws.cell(row=row, column=c).border = BORDER
        row += 1

    row += 2
    ws.cell(row=row, column=1, value=f"ANALYSE: La moyenne journaliere est remontee a {t4['moy_jour_t']:.1f} t/j (vs {t4_prev['moy_jour_t']:.1f} t/j au 24/07).")
    ws.cell(row=row, column=1).font = Font(bold=True, size=10)
    row += 1
    delta_moy = t4['moy_jour_t'] - t4_prev['moy_jour_t']
    if abs(delta_moy) < 2:
        msg = f"Variation faible ({delta_moy:+.1f} t/j) — dans la marge de fluctuation hebdomadaire."
    elif delta_moy < 0:
        msg = f"Baisse de {abs(delta_moy):.1f} t/j — a surveiller (effet possible de la hausse de prix)."
    else:
        msg = f"Hausse de {delta_moy:+.1f} t/j — pas d'effet visible de la hausse de prix pour le moment."
    ws.cell(row=row, column=1, value=msg)
    ws.cell(row=row, column=1).font = Font(italic=True, size=10, color='595959')
    row += 1
    ws.cell(row=row, column=1, value="L'effet plein de la hausse +2 000 XAF/sac (effective 23/07) sera visible sur les jours a venir.")
    ws.cell(row=row, column=1).font = Font(italic=True, size=10, color='C00000')
    row += 1
    ws.cell(row=row, column=1, value=f"Date de rupture probable: {t4['scenarios'][0]['date']} ({t4['scenarios'][0]['jours']} jours de stock).")
    ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='C00000')

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 32


def main():
    print('Loading new extraction (au 27/07)...')
    rows_new = load_livree(SRC)
    print(f'  {len(rows_new)} Livree rows')
    print('Loading previous extraction (au 24/07)...')
    rows_prev = load_livree(PREV_SRC)
    print(f'  {len(rows_prev)} Livree rows')

    # Compute tables for new (au 27/07)
    # T1: aggregate 1-27/07 (23 days lun-sam)
    t1_new = compute_t1(rows_new, '27/07/2026')
    # T2/T3: use 25/07 (last full business day with substantial soja data, since 27/07 is partial morning extraction)
    t2_new = compute_t2(rows_new, '25/07/2026')
    t3_new = compute_t3(rows_new, '25/07/2026')
    t4_new = compute_t4(t1_new, start_date_str='27/07/2026')

    # Compute tables for previous (au 24/07) - using the PREVIOUS extraction file
    t1_prev = compute_t1(rows_prev, '24/07/2026')
    t2_prev = compute_t2(rows_prev, '24/07/2026')
    t3_prev = compute_t3(rows_prev, '24/07/2026')
    t4_prev = compute_t4(t1_prev, start_date_str='24/07/2026')

    print(f"\nT1 (au 27/07): {t1_new['total_t']:.1f} t, moy/jour {t1_new['moy_jour_t']:.1f} t ({t1_new['days_elapsed']}j)")
    print(f"T1 (au 24/07): {t1_prev['total_t']:.1f} t, moy/jour {t1_prev['moy_jour_t']:.1f} t ({t1_prev['days_elapsed']}j)")
    print(f"\nT2 (27/07): {t2_new['total']['cmds']} cmdes, ratio {t2_new['total']['ratio_str']}")
    print(f"T2 (24/07): {t2_prev['total']['cmds']} cmdes, ratio {t2_prev['total']['ratio_str']}")
    print(f"\nT3 (27/07): {len(t3_new['items'])} cmdes soja-only, {t3_new['total_kg']/1000:.1f} t")
    print(f"T3 (24/07): {len(t3_prev['items'])} cmdes soja-only, {t3_prev['total_kg']/1000:.1f} t")
    print(f"\nT4 (au 27/07): rupture realiste {t4_new['scenarios'][0]['date']} ({t4_new['scenarios'][0]['jours']}j)")
    print(f"T4 (au 24/07): rupture realiste {t4_prev['scenarios'][0]['date']} ({t4_prev['scenarios'][0]['jours']}j)")

    # Write Excel
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws1 = wb.create_sheet('T1 - Moy jour soja par agence')
    write_t1(ws1, t1_new, t1_prev)

    ws2 = wb.create_sheet('T2 - Commandes soja 25-07')
    write_t2(ws2, t2_new, t2_prev)

    ws3 = wb.create_sheet('T3 - Soja-only 25-07')
    write_t3(ws3, t3_new, t3_prev)

    ws4 = wb.create_sheet('T4 - Stock soja et rupture')
    write_t4(ws4, t4_new, t4_prev)

    wb.save(OUT)
    print(f'\nSaved: {OUT}')

    # Also save a summary JSON for traceability
    summary = {
        'update_date': '27/07/2026',
        'source_file': SRC.split('/')[-1],
        'rows_livree_new': len(rows_new),
        'rows_livree_prev': len(rows_prev),
        't1_new': {'total_t': t1_new['total_t'], 'moy_jour_t': t1_new['moy_jour_t'], 'days_elapsed': t1_new['days_elapsed']},
        't1_prev': {'total_t': t1_prev['total_t'], 'moy_jour_t': t1_prev['moy_jour_t'], 'days_elapsed': t1_prev['days_elapsed']},
        't2_new': {'cmds': t2_new['total']['cmds'], 'bundle': t2_new['total']['bundle'], 'soja_only': t2_new['total']['soja_only'], 'ratio': t2_new['total']['ratio_str']},
        't2_prev': {'cmds': t2_prev['total']['cmds'], 'bundle': t2_prev['total']['bundle'], 'soja_only': t2_prev['total']['soja_only'], 'ratio': t2_prev['total']['ratio_str']},
        't3_new': {'count': len(t3_new['items']), 'total_t': t3_new['total_kg'] / 1000},
        't3_prev': {'count': len(t3_prev['items']), 'total_t': t3_prev['total_kg'] / 1000},
        't4_new': {'rupture_date': t4_new['scenarios'][0]['date'], 'jours_stock': t4_new['scenarios'][0]['jours'], 'moy_jour_t': t4_new['moy_jour_t']},
        't4_prev': {'rupture_date': t4_prev['scenarios'][0]['date'], 'jours_stock': t4_prev['scenarios'][0]['jours'], 'moy_jour_t': t4_prev['moy_jour_t']},
        'price_hike': {'date': '23/07/2026', 'amount': '+2 000 XAF/sac 50kg'},
    }
    with open('/home/z/my-project/scripts/soja_juillet_27.json', 'w', encoding='utf-8') as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)
    print('Summary saved: /home/z/my-project/scripts/soja_juillet_27.json')


if __name__ == '__main__':
    main()
