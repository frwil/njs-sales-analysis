"""
Update analyse_soja_aout.xlsx with the latest extraction (au 27/07/2026).

Source: /home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (4).xlsx
Output: /home/z/my-project/download/analyse_soja_aout.xlsx (overwritten)

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

SRC = '/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (3) (11) (1).xlsx'
# Pour août, pas d'extraction précédente — on compare vs juillet complet
PREV_SRC = '/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (9).xlsx'  # juillet complet (pour comparaison)
OUT = '/home/z/my-project/download/analyse_soja_aout.xlsx'
PREV_OUT = '/home/z/my-project/download/analyse_soja_aout.xlsx'  # previous output for comparison

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
STOCK_SACS = 160000  # 160k sacs de 50kg — stock physique au 21/07/2026
STOCK_T = STOCK_SACS * 50 / 1000  # 8000 t
STOCK_DATE = '01/08/2026'  # Date de référence du stock physique
STOCK_DAY = 1  # jour du mois (stock au 01/08)


def detect_column_indices(ws):
    """Auto-detect column indices from header row (row 2).
    Handles both 18-col format (with Date création/Date clôture) and 16-col format."""
    header = None
    for row in ws.iter_rows(min_row=2, max_row=2, values_only=True):
        header = row
        break
    if not header:
        return {'etat': 15, 'agence': 17}  # default old format
    indices = {}
    for i, h in enumerate(header):
        if h == 'État':
            indices['etat'] = i
        elif h == 'agence':
            indices['agence'] = i
    # Defaults if not found
    if 'etat' not in indices:
        indices['etat'] = 13 if len(header) <= 16 else 15
    if 'agence' not in indices:
        indices['agence'] = 15 if len(header) <= 16 else 17
    return indices


def load_livree(path):
    """Load only Livrée rows from the ERP extraction. Auto-detects column format."""
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb['Sheet 1']
    col_idx = detect_column_indices(ws)
    etat_col = col_idx['etat']
    agence_col = col_idx['agence']
    min_cols = max(etat_col, agence_col) + 1
    rows = list(ws.iter_rows(min_row=3, values_only=True))
    out = []
    for r in rows:
        if not r or len(r) < min_cols:
            continue
        if r[0] == 'Total':
            continue
        if r[etat_col] != 'Livrée':
            continue
        out.append(r)
    return out, col_idx


def cmd_soja_conc(rows, date_filter=None, col_idx=None):
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
        agence_raw = r[col_idx['agence']] or ''
        agence_short, region = AGENCE_MAP.get(agence_raw, (agence_raw, '?'))
        if ref in SOJA_REFS:
            cmds[cmd]['soja_kg'] += qte * SOJA_REFS[ref]
            cmds[cmd]['soja_sacs_50'] += qte * SOJA_REFS[ref] / 50
            cmds[cmd]['agence'] = agence_short
            cmds[cmd]['region'] = region
            cmds[cmd]['client'] = r[5]
            cmds[cmd]['etat'] = r[col_idx['etat']]
            cmds[cmd]['date'] = date
            cmds[cmd]['cmd_ref'] = cmd
        if ref in CONC_REFS:
            cmds[cmd]['conc_kg'] += qte * CONC_REFS[ref]
            cmds[cmd]['conc_sacs_50'] += qte * CONC_REFS[ref] / 50
            cmds[cmd]['agence'] = agence_short
            cmds[cmd]['region'] = region
            cmds[cmd]['client'] = r[5]
            cmds[cmd]['etat'] = r[col_idx['etat']]
            cmds[cmd]['date'] = date
            cmds[cmd]['cmd_ref'] = cmd
    # Keep commands that have soja OR conc
    return {k: v for k, v in cmds.items() if v['soja_kg'] > 0 or v['conc_kg'] > 0}


def compute_t1(rows, end_date_str, col_idx=None):
    """T1: Moyenne journaliere soja par agence/region, 01/MM to end_date."""
    # Days elapsed (lun-sam)
    end_d = int(end_date_str[:2])
    month = int(end_date_str[3:5])
    days_elapsed = 0
    for d in range(1, end_d + 1):
        dt = datetime.date(2026, month, d)
        if dt.weekday() < 6:
            days_elapsed += 1

    # Total soja per agence
    soja_by_agence = defaultdict(lambda: {'kg': 0, 'agence': None, 'region': None})
    for r in rows:
        ref = r[0]
        qte = r[2] or 0
        agence_raw = r[col_idx['agence']] or ''
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


def compute_t2(rows, date_str, col_idx=None):
    """T2: Commandes soja du jour par agence.
    cmds = commandes avec soja (soja_only + bundle).
    bundle = commandes avec soja ET conc.
    soja_only = commandes avec soja mais sans conc.
    ratio = total soja / total conc (toutes commandes incl. conc-only - portfolio mix)."""
    cmds = cmd_soja_conc(rows, date_filter=date_str, col_idx=col_idx)
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


def compute_t3(rows, date_str, col_idx=None):
    """T3: List of soja-only commands on date (commands with soja but NO conc)."""
    cmds = cmd_soja_conc(rows, date_filter=date_str, col_idx=col_idx)
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


def compute_t4_post_stock(rows, col_idx=None, current_date_str='04/08/2026', stock_date_str='01/08/2026',
                            price_hike_date='23/07/2026'):
    """T4: Stock soja et rupture — méthode corrigée.

    Le stock physique (160 000 sacs) est mesuré au 21/07 (stock_date).
    Les ventes 01→21/07 ont déjà été consommées et ne sont plus dans le stock.

    Calcul correct:
    1. Ventes post-stock = ventes soja du (stock_day+1) au current_date
    2. Moy/jour post-stock = ventes_post_stock / nb_jours_ouvres_post_stock
    3. Stock restant au current_date = stock_initial - ventes_post_stock (en sacs)
    4. Jours de stock restants = stock_restant / moy_jour_post_stock
    5. Date rupture = current_date + jours_stock (en jours ouvrables lun-sam)
    """
    stock_day = int(stock_date_str[:2])
    stock_month = int(stock_date_str[3:5])
    current_day = int(current_date_str[:2])
    current_month = int(current_date_str[3:5])

    # Jours ouvrables post-stock (du lendemain du stock_day au current_day, multi-month)
    post_stock_days = []
    if stock_month == current_month:
        # Same month
        for d in range(stock_day + 1, current_day + 1):
            dt = datetime.date(2026, stock_month, d)
            if dt.weekday() < 6:  # lun-sam
                post_stock_days.append(f'{d:02d}/{stock_month:02d}/2026')
    else:
        # Multi-month: from stock_day+1 to end of stock_month, then 1 to current_day of current_month
        # First month
        import calendar
        last_day_stock_month = calendar.monthrange(2026, stock_month)[1]
        for d in range(stock_day + 1, last_day_stock_month + 1):
            dt = datetime.date(2026, stock_month, d)
            if dt.weekday() < 6:
                post_stock_days.append(f'{d:02d}/{stock_month:02d}/2026')
        # Current month
        for d in range(1, current_day + 1):
            dt = datetime.date(2026, current_month, d)
            if dt.weekday() < 6:
                post_stock_days.append(f'{d:02d}/{current_month:02d}/2026')

    # Ventes soja sur ces jours
    post_stock_kg = 0
    for r in rows:
        ref = r[0]
        if ref not in SOJA_REFS:
            continue
        date_str = str(r[6])[:10]
        if date_str in post_stock_days:
            qte = r[2] or 0
            post_stock_kg += qte * SOJA_REFS[ref]

    # Jours ouvrables post-stock (lun-sam)
    nb_jours_post = len(post_stock_days)
    if nb_jours_post == 0:
        # Pas de ventes post-stock, utiliser moyenne globale comme fallback
        moy_t_post = 240  # fallback
        moy_sacs_post = moy_t_post * 20
    else:
        moy_t_post = (post_stock_kg / 1000) / nb_jours_post
        moy_sacs_post = (post_stock_kg / 50) / nb_jours_post

    # Stock restant au current_date
    post_stock_sacs = post_stock_kg / 50
    stock_restant_sacs = STOCK_SACS - post_stock_sacs
    stock_restant_t = stock_restant_sacs * 50 / 1000

    # Jours de stock restants
    if moy_sacs_post > 0:
        days_realiste = round(stock_restant_sacs / moy_sacs_post)
        days_plus20 = round(stock_restant_sacs / (moy_sacs_post * 1.2))
        days_moins20 = round(stock_restant_sacs / (moy_sacs_post * 0.8))
    else:
        days_realiste = days_plus20 = days_moins20 = 0

    # Date de rupture = current_date + jours_stock (lun-sam)
    start = datetime.date(2026, current_month, current_day)
    rupture_realiste = add_business_days(start, days_realiste)
    rupture_plus20 = add_business_days(start, days_plus20)
    rupture_moins20 = add_business_days(start, days_moins20)

    # Ventes totales août (pour info)
    total_août_kg = 0
    for r in rows:
        ref = r[0]
        if ref in SOJA_REFS:
            total_août_kg += (r[2] or 0) * SOJA_REFS[ref]

    return {
        'moy_sacs_post': moy_sacs_post,
        'moy_t_post': moy_t_post,
        'scenarios': [
            {'name': 'Realiste', 'sacs_jour': round(moy_sacs_post), 't_jour': moy_t_post,
             'jours': days_realiste, 'date': rupture_realiste.strftime('%d/%m/%Y')},
            {'name': 'Acceleration +20%', 'sacs_jour': round(moy_sacs_post * 1.2),
             't_jour': moy_t_post * 1.2, 'jours': days_plus20, 'date': rupture_plus20.strftime('%d/%m/%Y')},
            {'name': 'Ralentissement -20%', 'sacs_jour': round(moy_sacs_post * 0.8),
             't_jour': moy_t_post * 0.8, 'jours': days_moins20, 'date': rupture_moins20.strftime('%d/%m/%Y')},
        ],
        'stock_initial_sacs': STOCK_SACS,
        'stock_initial_t': STOCK_T,
        'stock_date': stock_date_str,
        'post_stock_days': post_stock_days,
        'nb_jours_post': nb_jours_post,
        'post_stock_sacs': round(post_stock_sacs),
        'post_stock_t': round(post_stock_kg / 1000, 1),
        'stock_restant_sacs': round(stock_restant_sacs),
        'stock_restant_t': round(stock_restant_t, 1),
        'ventes_août_sacs': round(total_août_kg / 50),
        'ventes_août_t': round(total_août_kg / 1000, 1),
        'current_date': current_date_str,
        'price_hike_date': price_hike_date,
    }


def compute_t5_trend(rows, end_date_str='04/08/2026', col_idx=None):
    """T5: Tendance journalière soja et concentrés en août.
    Retourne pour chaque jour ouvré (lun-sam): date, soja (t), conc (t), ratio soja:conc.
    """
    # Build dict date -> {soja_kg, conc_kg}
    daily = defaultdict(lambda: {'soja_kg': 0, 'conc_kg': 0})
    for r in rows:
        ref = r[0]
        date_str = str(r[6])[:10]
        if not date_str.startswith('/') and len(date_str) >= 10:
            qte = r[2] or 0
            if ref in SOJA_REFS:
                daily[date_str]['soja_kg'] += qte * SOJA_REFS[ref]
            if ref in CONC_REFS:
                daily[date_str]['conc_kg'] += qte * CONC_REFS[ref]

    # Build sorted list of dates in month (01 → end_date)
    end_day = int(end_date_str[:2])
    month = int(end_date_str[3:5])
    items = []
    for d in range(1, end_day + 1):
        dt = datetime.date(2026, month, d)
        date_str = f'{d:02d}/{month:02d}/2026'
        is_ouvre = dt.weekday() < 6  # lun-sam
        weekday_name = ['Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam', 'Dim'][dt.weekday()]
        soja_t = daily.get(date_str, {}).get('soja_kg', 0) / 1000
        conc_t = daily.get(date_str, {}).get('conc_kg', 0) / 1000
        ratio = (daily.get(date_str, {}).get('soja_kg', 0) / 50) / (daily.get(date_str, {}).get('conc_kg', 0) / 50) if daily.get(date_str, {}).get('conc_kg', 0) > 0 else None
        items.append({
            'date': date_str,
            'jour': d,
            'weekday': weekday_name,
            'is_ouvre': is_ouvre,
            'soja_t': round(soja_t, 1),
            'conc_t': round(conc_t, 1),
            'ratio': round(ratio, 2) if ratio else None,
        })
    return {'items': items, 'end_date': end_date_str}


def write_t5(ws, t5):
    """T5: Feuille avec table des ventes journalières + graphique de tendance soja/concentrés."""
    from openpyxl.chart import LineChart, Reference, BarChart
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.layout import Layout, ManualLayout

    ws['A1'] = 'TABLEAU 5 - Tendance journalière SOJA et CONCENTRÉS (Août 2026)'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws['A2'] = f"Données Livrées au {t5['end_date']}. Jours ouvrés (lun-sam) en couleur, dimanches en gris."
    ws['A2'].font = Font(italic=True, size=10, color='595959')

    headers = ['Date', 'Jour', 'Jour sem.', 'Soja (t)', 'Concentrés (t)', 'Ratio soja:conc']
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, len(headers))

    row = 5
    for it in t5['items']:
        ws.cell(row=row, column=1, value=it['date'])
        ws.cell(row=row, column=2, value=it['jour'])
        ws.cell(row=row, column=3, value=it['weekday'])
        ws.cell(row=row, column=4, value=it['soja_t'])
        ws.cell(row=row, column=5, value=it['conc_t'])
        ws.cell(row=row, column=6, value=f"{it['ratio']}:1" if it['ratio'] else '—')
        # Highlight dimanches
        if not it['is_ouvre']:
            for c in range(1, len(headers) + 1):
                ws.cell(row=row, column=c).fill = PatternFill('solid', fgColor='F2F2F2')
                ws.cell(row=row, column=c).font = Font(italic=True, color='808080')
        # Highlight hausse prix 23/07
        if it['jour'] == 23:
            for c in range(1, len(headers) + 1):
                ws.cell(row=row, column=c).fill = PatternFill('solid', fgColor='FCE4D6')
                ws.cell(row=row, column=c).font = Font(bold=True, color='C00000')
        for c in range(1, len(headers) + 1):
            ws.cell(row=row, column=c).border = BORDER
        row += 1

    # Line chart: Soja vs Concentrés
    chart = LineChart()
    chart.title = 'Tendance journalière — Soja vs Concentrés (Août 2026)'
    chart.style = 12
    chart.y_axis.title = 'Tonnes'
    chart.x_axis.title = 'Date'
    chart.height = 12  # cm
    chart.width = 24

    data = Reference(ws, min_col=4, min_row=4, max_col=5, max_row=row - 1)
    cats = Reference(ws, min_col=1, min_row=5, max_row=row - 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    # Style series
    if len(chart.series) >= 2:
        from openpyxl.chart.marker import Marker
        from openpyxl.drawing.line import LineProperties
        from openpyxl.drawing.colors import ColorChoice
        # Soja = navy, Conc = gold
        chart.series[0].graphicalProperties = openpyxl.chart.series.GraphicalProperties(solidFill='1F4E78')
        chart.series[1].graphicalProperties = openpyxl.chart.series.GraphicalProperties(solidFill='C9A961')

    ws.add_chart(chart, 'H4')

    # Bar chart: Ratio soja:conc (only for days with ratio)
    chart2 = BarChart()
    chart2.type = 'col'
    chart2.title = 'Ratio soja:concentrés par jour (objectif 3:1)'
    chart2.style = 10
    chart2.y_axis.title = 'Ratio (soja:conc)'
    chart2.x_axis.title = 'Date'
    chart2.height = 10
    chart2.width = 24

    # Filter ratio data (replace None with 0)
    ratio_col = 7  # use a helper column for ratios
    ws.cell(row=4, column=ratio_col, value='Ratio numérique')
    style_header_row(ws, 4, ratio_col)
    row2 = 5
    for it in t5['items']:
        ws.cell(row=row2, column=ratio_col, value=it['ratio'] if it['ratio'] else None)
        row2 += 1

    data2 = Reference(ws, min_col=ratio_col, min_row=4, max_row=row2 - 1)
    cats2 = Reference(ws, min_col=1, min_row=5, max_row=row2 - 1)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.legend = None

    ws.add_chart(chart2, 'H30')

    # Notes
    row_note = max(row, row2) + 3
    ws.cell(row=row_note, column=1, value='NOTES DE LECTURE:')
    ws.cell(row=row_note, column=1).font = Font(bold=True, size=11, color='1F4E78')
    row_note += 1
    notes = [
        '• La ligne "Soja" (bleu) montre la consommation journalière de tourteaux de soja — impactée par la hausse +1 000 FCFA/sac (début août) puis +2 000 FCFA/sac (23/07).',
        '• La ligne "Concentrés" (or) montre la consommation journalière de concentrés (BELGO Chair/Ponte/Porc).',
        '• Le ratio soja:conc cible est 3:1 (objectif bundle). Un ratio supérieur signifie que les clients achètent plus de soja que de concentrés — dégradation du bundle.',
        '• Les dimanches sont grisés (pas de ventes). Le 23/07 est surligné en rouge (date de la 2e hausse tarifaire +2 000 FCFA/sac).',
        '• À observer: la tendance soja après le 23/07 vs avant — l\'effet prix devrait ralentir la consommation soja si les clients sont sensibles au prix.',
    ]
    for note in notes:
        ws.cell(row=row_note, column=1, value=note)
        ws.cell(row=row_note, column=1).font = Font(size=10)
        row_note += 1

    # Column widths
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 8
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 16
    ws.column_dimensions['F'].width = 16
    ws.column_dimensions['G'].width = 16


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
    ws['A1'] = 'TABLEAU 1 - Moyenne journaliere SOJA août par agence/region'
    ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
    ws['A2'] = f"01/07 au {t1['end_date']} ({t1['days_elapsed']}j). Livree. lun-sam. Mise a jour 04/08/2026."
    ws['A2'].font = Font(italic=True, size=10, color='595959')

    headers = ['Agence', 'Region', 'Total soja (t)', 'Moy/jour (t)', 'Evolution vs 31/07']
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, len(headers))

    row = 5
    for ag, reg, t in t1['items']:
        ws.cell(row=row, column=1, value=ag)
        ws.cell(row=row, column=2, value=reg)
        ws.cell(row=row, column=3, value=round(t, 1))
        ws.cell(row=row, column=4, value=round(t / t1['days_elapsed'], 1))
        # Evolution vs 31/07
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
    ws.cell(row=row, column=1, value=f"Precedent (31/07): {prev_total_moy:.1f} t/j ({int(t1_prev['moy_jour_sacs'])} sacs) sur {t1_prev['days_elapsed']}j.")
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
    ws['A2'] = 'Livree. Ratio en sacs 50kg-equivalent. Mise a jour 04/08/2026.'
    ws['A2'].font = Font(italic=True, size=10, color='595959')

    headers = ['Agence', 'Region', 'Cmds', 'Bundle', 'Soja-only', 'Ratio moy', 'Evolution cmdes vs 31/07']
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
    ws.cell(row=row, column=1, value=f"Ratio global: {t2['total']['ratio_str']} (vs {t2_prev['total']['ratio_str']} le 31/07).")
    ws.cell(row=row, column=1).font = Font(bold=True, size=10)
    row += 2
    ws.cell(row=row, column=1, value="Note: T2/T3 portent sur le 04/08 (dernier jour complet, 82 cmdes soja). Comparaison vs 31/07 (bilan juillet complet).")
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

    headers = ['#', 'N cmd', 'Client', 'Agence', 'Region', 'Soja (t)', 'Sacs', 'Etat', 'Nouveau vs 31/07']
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
    ws.cell(row=row, column=1, value=f"Precedent (31/07): {len(t3_prev['items'])} cmdes ({round(t3_prev['total_kg']/1000, 1)} t)")
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
    ws['A2'] = (f"Stock physique au {t4['stock_date']}: {t4['stock_initial_sacs']:,} sacs ({int(t4['stock_initial_t'])} t). "
                f"Methode corrigee: ventes post-stock {t4['stock_date']}→{t4['current_date']} deduites du stock. "
                f"Mise a jour {t4['current_date']}.").replace(',', ' ')
    ws['A2'].font = Font(italic=True, size=10, color='595959')

    headers = ['Scenario', 'Sacs/jour (post-stock)', 't/jour', 'Jours stock', 'Date rupture', 'Evolution date rupture']
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
        ('Stock initial (au 21/07)', f"{t4['stock_initial_sacs']:,} sacs".replace(',', ' '), f"{int(t4['stock_initial_t'])} t"),
        ('Ventes post-stock 02→04/08', f"{t4['post_stock_sacs']:,} sacs".replace(',', ' '), f"{t4['post_stock_t']} t"),
        ('Stock restant au 04/08', f"{t4['stock_restant_sacs']:,} sacs".replace(',', ' '), f"{t4['stock_restant_t']} t"),
        ('Jours ouvrables post-stock', f"{t4['nb_jours_post']} j (lun-sam)", ''),
        ('Moyenne/jour post-stock', f"{int(t4['moy_sacs_post']):,} sacs".replace(',', ' '), f"{t4['moy_t_post']:.1f} t"),
        ('Ventes août (total, info)', f"{t4['ventes_août_sacs']:,} sacs".replace(',', ' '), f"{t4['ventes_août_t']} t"),
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
    ws.cell(row=row, column=1, value=f"ANALYSE: Stock restant au {t4['current_date']} = {t4['stock_restant_sacs']:,} sacs ({t4['stock_restant_t']} t).".replace(',', ' '))
    ws.cell(row=row, column=1).font = Font(bold=True, size=10)
    row += 1
    delta_moy = t4['moy_t_post'] - t4_prev['moy_t_post']
    if abs(delta_moy) < 2:
        msg = f"Moy/jour post-stock stable ({delta_moy:+.1f} t/j vs precedent)."
    elif delta_moy < 0:
        msg = f"Baisse moy/jour post-stock ({delta_moy:+.1f} t/j) — effet possible de la hausse de prix."
    else:
        msg = f"Hausse moy/jour post-stock ({delta_moy:+.1f} t/j) — pas d'effet visible de la hausse de prix."
    ws.cell(row=row, column=1, value=msg)
    ws.cell(row=row, column=1).font = Font(italic=True, size=10, color='595959')
    row += 1
    ws.cell(row=row, column=1, value=f"Methode corrigee: le stock de 160 000 sacs est mesure au 21/07. Les ventes 01→21/07 sont deja consommees.")
    ws.cell(row=row, column=1).font = Font(italic=True, size=9, color='595959')
    row += 1
    ws.cell(row=row, column=1, value=f"Seules les ventes 02→04/08 (post-stock) sont deduites du stock pour calculer le stock restant au 04/08.")
    ws.cell(row=row, column=1).font = Font(italic=True, size=9, color='595959')
    row += 1
    ws.cell(row=row, column=1, value=f"Date de rupture probable: {t4['scenarios'][0]['date']} ({t4['scenarios'][0]['jours']} jours de stock restant).")
    ws.cell(row=row, column=1).font = Font(bold=True, size=11, color='C00000')

    ws.column_dimensions['A'].width = 32
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 32


def main():
    print('Loading new extraction (au 04/08)...')
    rows_new, col_idx_new = load_livree(SRC)
    print(f'  {len(rows_new)} Livree rows (cols: etat={col_idx_new["etat"]}, agence={col_idx_new["agence"]})')
    print('Loading previous extraction (au 31/07)...')
    rows_prev, col_idx_prev = load_livree(PREV_SRC)
    print(f'  {len(rows_prev)} Livree rows (cols: etat={col_idx_prev["etat"]}, agence={col_idx_prev["agence"]})')

    # Compute tables for new (au 31/07 — bilan complet définitif)
    # T1: aggregate 1-31/07 (27 days lun-sam — mois complet)
    t1_new = compute_t1(rows_new, '04/08/2026', col_idx=col_idx_new)
    # T2/T3: use 31/07 (now complete with 109 soja commands Livrées)
    t2_new = compute_t2(rows_new, '04/08/2026', col_idx=col_idx_new)
    t3_new = compute_t3(rows_new, '04/08/2026', col_idx=col_idx_new)
    # T4: méthode corrigée — stock physique au 21/07, ventes post-stock à déduire
    t4_new = compute_t4_post_stock(rows_new, col_idx=col_idx_new, current_date_str='04/08/2026', stock_date_str='01/08/2026')

    # Compute tables for previous (extraction au 31/07 matinale — used 30/07 for T2/T3)
    t1_prev = compute_t1(rows_prev, '31/07/2026', col_idx=col_idx_prev)
    t2_prev = compute_t2(rows_prev, '31/07/2026', col_idx=col_idx_prev)
    t3_prev = compute_t3(rows_prev, '31/07/2026', col_idx=col_idx_prev)
    # T4 précédent: stock au 21/07, ventes 02-04/08 (post-stock)
    t4_prev = compute_t4_post_stock(rows_prev, col_idx=col_idx_prev, current_date_str='31/07/2026', stock_date_str='21/07/2026')

    print(f"\nT1 (au 04/08): {t1_new['total_t']:.1f} t, moy/jour {t1_new['moy_jour_t']:.1f} t ({t1_new['days_elapsed']}j)")
    print(f"T1 (au 31/07): {t1_prev['total_t']:.1f} t, moy/jour {t1_prev['moy_jour_t']:.1f} t ({t1_prev['days_elapsed']}j)")
    print(f"\nT2 (04/08): {t2_new['total']['cmds']} cmdes, ratio {t2_new['total']['ratio_str']}")
    print(f"T2 (31/07): {t2_prev['total']['cmds']} cmdes, ratio {t2_prev['total']['ratio_str']}")
    print(f"\nT3 (04/08): {len(t3_new['items'])} cmdes soja-only, {t3_new['total_kg']/1000:.1f} t")
    print(f"T3 (31/07): {len(t3_prev['items'])} cmdes soja-only, {t3_prev['total_kg']/1000:.1f} t")
    print(f"\nT4 (au 04/08): stock restant {t4_new['stock_restant_sacs']:,} sacs ({t4_new['stock_restant_t']} t)".replace(',', ' '))
    print(f"  Ventes post-stock 02-04/08: {t4_new['post_stock_sacs']:,} sacs ({t4_new['post_stock_t']} t)".replace(',', ' '))
    print(f"  Moy/jour post-stock: {t4_new['moy_t_post']:.1f} t/j ({int(t4_new['moy_sacs_post'])} sacs/j)")
    print(f"  Rupture realiste: {t4_new['scenarios'][0]['date']} ({t4_new['scenarios'][0]['jours']}j)")
    print(f"T4 (au 31/07): rupture realiste {t4_prev['scenarios'][0]['date']} ({t4_prev['scenarios'][0]['jours']}j)")

    # Write Excel
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws1 = wb.create_sheet('T1 - Moy jour soja par agence')
    write_t1(ws1, t1_new, t1_prev)

    ws2 = wb.create_sheet('T2 - Commandes soja 04-08')
    write_t2(ws2, t2_new, t2_prev)

    ws3 = wb.create_sheet('T3 - Soja-only 04-08')
    write_t3(ws3, t3_new, t3_prev)

    ws4 = wb.create_sheet('T4 - Stock soja et rupture')
    write_t4(ws4, t4_new, t4_prev)

    # T5: tendance journalière soja + concentrés avec graphiques
    t5_new = compute_t5_trend(rows_new, end_date_str='04/08/2026', col_idx=col_idx_new)
    ws5 = wb.create_sheet('T5 - Tendance Août')
    write_t5(ws5, t5_new)

    wb.save(OUT)
    print(f'\nSaved: {OUT}')

    # Also save a summary JSON for traceability
    summary = {
        'update_date': '04/08/2026',
        'source_file': SRC.split('/')[-1],
        'rows_livree_new': len(rows_new),
        'rows_livree_prev': len(rows_prev),
        't1_new': {'total_t': t1_new['total_t'], 'moy_jour_t': t1_new['moy_jour_t'], 'days_elapsed': t1_new['days_elapsed']},
        't1_prev': {'total_t': t1_prev['total_t'], 'moy_jour_t': t1_prev['moy_jour_t'], 'days_elapsed': t1_prev['days_elapsed']},
        't2_new': {'cmds': t2_new['total']['cmds'], 'bundle': t2_new['total']['bundle'], 'soja_only': t2_new['total']['soja_only'], 'ratio': t2_new['total']['ratio_str']},
        't2_prev': {'cmds': t2_prev['total']['cmds'], 'bundle': t2_prev['total']['bundle'], 'soja_only': t2_prev['total']['soja_only'], 'ratio': t2_prev['total']['ratio_str']},
        't3_new': {'count': len(t3_new['items']), 'total_t': t3_new['total_kg'] / 1000},
        't3_prev': {'count': len(t3_prev['items']), 'total_t': t3_prev['total_kg'] / 1000},
        't4_new': {'rupture_date': t4_new['scenarios'][0]['date'], 'jours_stock': t4_new['scenarios'][0]['jours'],
                    'moy_jour_t': t4_new['moy_t_post'], 'stock_restant_sacs': t4_new['stock_restant_sacs'],
                    'stock_restant_t': t4_new['stock_restant_t'], 'post_stock_t': t4_new['post_stock_t']},
        't4_prev': {'rupture_date': t4_prev['scenarios'][0]['date'], 'jours_stock': t4_prev['scenarios'][0]['jours'],
                    'moy_jour_t': t4_prev['moy_t_post'], 'stock_restant_sacs': t4_prev['stock_restant_sacs'],
                    'stock_restant_t': t4_prev['stock_restant_t'], 'post_stock_t': t4_prev['post_stock_t']},
        'price_hike': {'date': '23/07/2026', 'amount': '+2 000 XAF/sac 50kg'},
    }
    with open('/home/z/my-project/scripts/soja_aout_31.json', 'w', encoding='utf-8') as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)
    print('Summary saved: /home/z/my-project/scripts/soja_aout_31.json')


if __name__ == '__main__':
    main()
