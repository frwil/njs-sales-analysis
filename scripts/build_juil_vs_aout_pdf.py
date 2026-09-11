"""Generate PDF comparison July 2026 vs August 2026 — Volumes + CA + Objectifs.
Includes realization % vs objectives for each month.
"""
import os
import json
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, PageBreak, Image
)

# === Fonts ===
pdfmetrics.registerFont(TTFont('DejaVuSans', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))

# === Colors ===
NAVY = colors.HexColor('#1F3A5F')
GOLD = colors.HexColor('#C9A961')
RED = colors.HexColor('#C00000')
GREEN = colors.HexColor('#548235')
ORANGE = colors.HexColor('#ED7D31')
GRAY = colors.HexColor('#595959')
LIGHT_GRAY = colors.HexColor('#F2F2F2')
LIGHT_GREEN = colors.HexColor('#E2EFDA')
LIGHT_RED = colors.HexColor('#FCE4EC')
LIGHT_YELLOW = colors.HexColor('#FFF9E6')

styles = getSampleStyleSheet()
H1 = ParagraphStyle('H1', parent=styles['Heading1'], fontName='DejaVuSans-Bold', fontSize=18, textColor=NAVY, spaceAfter=12, spaceBefore=8)
H2 = ParagraphStyle('H2', parent=styles['Heading2'], fontName='DejaVuSans-Bold', fontSize=13, textColor=NAVY, spaceAfter=8, spaceBefore=12)
H3 = ParagraphStyle('H3', parent=styles['Heading3'], fontName='DejaVuSans-Bold', fontSize=11, textColor=GOLD, spaceAfter=6, spaceBefore=8)
BODY = ParagraphStyle('Body', parent=styles['Normal'], fontName='DejaVuSans', fontSize=10, leading=14, alignment=TA_JUSTIFY, spaceAfter=6)
BULLET = ParagraphStyle('Bullet', parent=BODY, leftIndent=20, bulletIndent=10, spaceAfter=4)
SMALL = ParagraphStyle('Small', parent=BODY, fontSize=8, textColor=GRAY, leading=10)

CELL_HEADER = ParagraphStyle('CellHeader', fontName='DejaVuSans-Bold', fontSize=8, leading=10, alignment=TA_CENTER, textColor=colors.white)
CELL_CENTER = ParagraphStyle('CellCenter', fontName='DejaVuSans', fontSize=8, leading=10, alignment=TA_CENTER)
CELL_LEFT = ParagraphStyle('CellLeft', fontName='DejaVuSans', fontSize=8, leading=10, alignment=TA_LEFT)

def wrap_cell(content, style=CELL_CENTER):
    s = str(content) if content is not None else ''
    return Paragraph(s, style)

def fmt(x):
    if isinstance(x, (int, float)):
        return f"{x:,.1f}".replace(',', ' ').replace('.', ',')
    return str(x)

def fmt_int(x):
    if isinstance(x, (int, float)):
        return f"{int(round(x)):,}".replace(',', ' ')
    return str(x)

def fmt_signed(x):
    s = '+' if x >= 0 else ''
    return f"{s}{x:.1f}%".replace('.', ',')

def color_for_pct(pct):
    """Color based on % of objective achieved."""
    if pct >= 100: return LIGHT_GREEN
    if pct >= 80: return LIGHT_YELLOW
    return LIGHT_RED

def color_for_var(x):
    if x >= 0: return LIGHT_GREEN
    if x <= -20: return LIGHT_RED
    if x <= -5: return LIGHT_YELLOW
    return LIGHT_GRAY

# === Load data ===
print("Loading data...")
df = pd.read_csv('/home/z/my-project/scripts/dataset_2023_2026.csv', parse_dates=['date'], low_memory=False)
print(f"  Dataset: {len(df)} records")

# Load objectives (CA + Volume)
ca_obj = json.load(open('/home/z/my-project/scripts/ca_obj_real.json'))
vol_obj = json.load(open('/home/z/my-project/scripts/s2_recaled_objectives.json'))
print(f"  Objectives CA: {len(ca_obj.get('ca_obj_monthly', {}))} families")
print(f"  Objectives Volume: {len(vol_obj.get('global_s2_recaled', {}))} families")

# Filter July and August 2026
juil = df[(df['date'].dt.year == 2026) & (df['date'].dt.month == 7)].copy()
aout = df[(df['date'].dt.year == 2026) & (df['date'].dt.month == 8)].copy()

# CA combined (TTC fallback to HT)
juil['ca_combined'] = juil['montant_ttc'].where(juil['montant_ttc'] > 0, juil['montant_ht'])
aout['ca_combined'] = aout['montant_ttc'].where(aout['montant_ttc'] > 0, aout['montant_ht'])

# Mapping between dataset family names and objectives file family names
FAMILY_MAP = {
    'TOURTEAUX': 'TOURTEAUX',
    'CONCENTRES': 'CONCENTRES',
    'INGREDIENTS': 'INGREDIENTS',
    'ALIMENT_COMPLET': 'ALIMENT COMPLET',
    'PREMIX': 'PREMIX',
    'COMPLEMENT_ALIMENTAIRE': 'COMPLEMENT ALIMENTAIRE',
    'ALVEOLES': 'ALVEOLE',
    'MATERIEL_ELEVAGE': 'MATERIEL ELEVAGE',
    'MAIS': 'DIVERS',
}

# === Compute aggregates ===
print("\nComputing aggregates...")

def get_realization(month_data, month_num):
    """Returns dict family → (volume_t, ca_m_fcfa, vol_obj_t, ca_obj_m_fcfa, pct_vol, pct_ca)"""
    result = {}
    for ds_fam, obj_fam in FAMILY_MAP.items():
        fam_data = month_data[month_data['family'] == ds_fam]
        vol_t = fam_data['tonnes'].sum()
        ca_m = fam_data['ca_combined'].sum() / 1e6  # M FCFA
        
        # Volume objective for this month (key 7 or 8)
        vol_obj_t = vol_obj.get('global_s2_recaled', {}).get(obj_fam, {}).get(str(month_num), 0)
        
        # CA objective for this month (index 6 for July, 7 for August)
        ca_obj_fcfa = ca_obj.get('ca_obj_monthly', {}).get(obj_fam, [0]*12)[month_num - 1]
        ca_obj_m = ca_obj_fcfa / 1e6  # Convert to M FCFA
        
        pct_vol = (vol_t / vol_obj_t * 100) if vol_obj_t > 0 else 0
        pct_ca = (ca_m / ca_obj_m * 100) if ca_obj_m > 0 else 0
        
        result[ds_fam] = {
            'vol_t': vol_t,
            'ca_m': ca_m,
            'vol_obj_t': vol_obj_t,
            'ca_obj_m': ca_obj_m,
            'pct_vol': pct_vol,
            'pct_ca': pct_ca,
        }
    return result

juil_data = get_realization(juil, 7)
aout_data = get_realization(aout, 8)

# Totals
total_juil = {'vol_t': sum(d['vol_t'] for d in juil_data.values()),
              'ca_m': sum(d['ca_m'] for d in juil_data.values()),
              'vol_obj_t': sum(d['vol_obj_t'] for d in juil_data.values()),
              'ca_obj_m': sum(d['ca_obj_m'] for d in juil_data.values())}
total_aout = {'vol_t': sum(d['vol_t'] for d in aout_data.values()),
              'ca_m': sum(d['ca_m'] for d in aout_data.values()),
              'vol_obj_t': sum(d['vol_obj_t'] for d in aout_data.values()),
              'ca_obj_m': sum(d['ca_obj_m'] for d in aout_data.values())}
total_juil['pct_vol'] = (total_juil['vol_t'] / total_juil['vol_obj_t'] * 100) if total_juil['vol_obj_t'] > 0 else 0
total_juil['pct_ca'] = (total_juil['ca_m'] / total_juil['ca_obj_m'] * 100) if total_juil['ca_obj_m'] > 0 else 0
total_aout['pct_vol'] = (total_aout['vol_t'] / total_aout['vol_obj_t'] * 100) if total_aout['vol_obj_t'] > 0 else 0
total_aout['pct_ca'] = (total_aout['ca_m'] / total_aout['ca_obj_m'] * 100) if total_aout['ca_obj_m'] > 0 else 0

# === Build PDF ===
OUT = '/home/z/my-project/download/comparaison_juillet_aout_2026.pdf'
doc = SimpleDocTemplate(OUT, pagesize=A4, topMargin=1.5*cm, bottomMargin=1.5*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
story = []

# === Cover ===
story.append(Paragraph("BELGOCAM SA", ParagraphStyle('CT', parent=styles['Title'], fontName='DejaVuSans-Bold', fontSize=22, textColor=NAVY, alignment=TA_CENTER, spaceAfter=6)))
story.append(Paragraph("Comparaison Juillet vs Août 2026", ParagraphStyle('CS', parent=styles['Title'], fontName='DejaVuSans-Bold', fontSize=18, textColor=GOLD, alignment=TA_CENTER, spaceAfter=4)))
story.append(Paragraph("Volumes + CA + Réalisation vs Objectifs", ParagraphStyle('CS2', parent=BODY, fontSize=11, textColor=GRAY, alignment=TA_CENTER, spaceAfter=10)))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.3*cm))

# === 1. Vue Globale ===
story.append(Paragraph("1. Vue Globale — Toutes familles", H2))

glob_data = [
    ["Indicateur", "Juillet 2026", "Août 2026", "Variation", "Δ"],
    ["Volume réalisé (t)", fmt(total_juil['vol_t']), fmt(total_aout['vol_t']),
     fmt_signed((total_aout['vol_t']/total_juil['vol_t']-1)*100), ""],
    ["Volume objectif (t)", fmt(total_juil['vol_obj_t']), fmt(total_aout['vol_obj_t']),
     fmt_signed((total_aout['vol_obj_t']/total_juil['vol_obj_t']-1)*100), ""],
    ["% réalisation Volume", f"{total_juil['pct_vol']:.1f}%", f"{total_aout['pct_vol']:.1f}%",
     fmt_signed(total_aout['pct_vol'] - total_juil['pct_vol']), "pts"],
    ["CA réalisé (M FCFA)", fmt(total_juil['ca_m']), fmt(total_aout['ca_m']),
     fmt_signed((total_aout['ca_m']/total_juil['ca_m']-1)*100), ""],
    ["CA objectif (M FCFA)", fmt(total_juil['ca_obj_m']), fmt(total_aout['ca_obj_m']),
     fmt_signed((total_aout['ca_obj_m']/total_juil['ca_obj_m']-1)*100), ""],
    ["% réalisation CA", f"{total_juil['pct_ca']:.1f}%", f"{total_aout['pct_ca']:.1f}%",
     fmt_signed(total_aout['pct_ca'] - total_juil['pct_ca']), "pts"],
]

wrapped_glob = []
for ri, row in enumerate(glob_data):
    prow = []
    for ci, cell in enumerate(row):
        if ri == 0:
            prow.append(wrap_cell(cell, CELL_HEADER))
        elif ci == 0:
            prow.append(wrap_cell(cell, CELL_LEFT))
        else:
            prow.append(wrap_cell(cell, CELL_CENTER))
    wrapped_glob.append(prow)

t_glob = Table(wrapped_glob, colWidths=[4*cm, 3*cm, 3*cm, 3*cm, 1.5*cm], repeatRows=1)
style_glob = [
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
    ('TOPPADDING', (0,0), (-1,-1), 5), ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ('LEFTPADDING', (0,0), (-1,-1), 4), ('RIGHTPADDING', (0,0), (-1,-1), 4),
    # Color % realization cells
    ('BACKGROUND', (1,3), (2,3), color_for_pct(total_juil['pct_vol'])),
    ('BACKGROUND', (1,6), (2,6), color_for_pct(total_juil['pct_ca'])),
    # Color % cells for Août
    ('BACKGROUND', (2,3), (2,3), color_for_pct(total_aout['pct_vol'])),
    ('BACKGROUND', (2,6), (2,6), color_for_pct(total_aout['pct_ca'])),
    # Bold realization rows
    ('FONT', (0,3), (-1,3), 'DejaVuSans-Bold', 8),
    ('FONT', (0,6), (-1,6), 'DejaVuSans-Bold', 8),
]
# Color variation cells
style_glob.append(('BACKGROUND', (3,1), (3,1), color_for_var((total_aout['vol_t']/total_juil['vol_t']-1)*100)))
style_glob.append(('BACKGROUND', (3,4), (3,4), color_for_var((total_aout['ca_m']/total_juil['ca_m']-1)*100)))
t_glob.setStyle(TableStyle(style_glob))
story.append(t_glob)
story.append(Spacer(1, 0.3*cm))

# Global analysis
var_vol = (total_aout['vol_t']/total_juil['vol_t']-1)*100
var_ca = (total_aout['ca_m']/total_juil['ca_m']-1)*100
gap_vol_juil = total_juil['vol_obj_t'] - total_juil['vol_t']
gap_ca_juil = total_juil['ca_obj_m'] - total_juil['ca_m']
gap_vol_aout = total_aout['vol_obj_t'] - total_aout['vol_t']
gap_ca_aout = total_aout['ca_obj_m'] - total_aout['ca_m']

glob_analysis = f"""
<b>Lecture globale</b> : 
<br/>• <b>Volume</b> : {fmt(total_juil['vol_t'])} t en juillet → {fmt(total_aout['vol_t'])} t en août ({fmt_signed(var_vol)}). 
L'objectif volume est passé de {fmt(total_juil['vol_obj_t'])} t à {fmt(total_aout['vol_obj_t'])} t entre les deux mois (objectif août plus bas que juillet de {fmt_signed((total_aout['vol_obj_t']/total_juil['vol_obj_t']-1)*100)}).
<br/>• <b>CA</b> : {fmt(total_juil['ca_m'])} M FCFA → {fmt(total_aout['ca_m'])} M FCFA ({fmt_signed(var_ca)}). 
L'objectif CA est passé de {fmt(total_juil['ca_obj_m'])} M à {fmt(total_aout['ca_obj_m'])} M ({fmt_signed((total_aout['ca_obj_m']/total_juil['ca_obj_m']-1)*100)}).
<br/>• <b>Réalisation Volume</b> : juillet à {total_juil['pct_vol']:.0f}% de l'objectif (gap {fmt_int(gap_vol_juil)} t), août à {total_aout['pct_vol']:.0f}% (gap {fmt_int(gap_vol_aout)} t).
<br/>• <b>Réalisation CA</b> : juillet à {total_juil['pct_ca']:.0f}% de l'objectif (gap {fmt_int(gap_ca_juil)} M), août à {total_aout['pct_ca']:.0f}% (gap {fmt_int(gap_ca_aout)} M).
"""
story.append(Paragraph(glob_analysis, BODY))
story.append(Spacer(1, 0.3*cm))

# === 2. Par Famille ===
story.append(PageBreak())
story.append(Paragraph("2. Vue par Famille — Volumes + CA + Objectifs", H2))

# Volume table by family
story.append(Paragraph("<b>2.1 Volumes par famille (réalisé vs objectif)</b>", H3))

fam_data_vol = [["Famille", "Vol Juil (t)", "Obj Juil (t)", "% Juil", "Vol Août (t)", "Obj Août (t)", "% Août", "Δ Vol Juil→Août"]]
families_order = ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'ALIMENT_COMPLET', 'PREMIX', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES', 'MATERIEL_ELEVAGE']

for fam in families_order:
    j = juil_data[fam]
    a = aout_data[fam]
    var = (a['vol_t']/j['vol_t']-1)*100 if j['vol_t'] > 0 else 0
    fam_data_vol.append([fam, fmt(j['vol_t']), fmt(j['vol_obj_t']), f"{j['pct_vol']:.0f}%",
                        fmt(a['vol_t']), fmt(a['vol_obj_t']), f"{a['pct_vol']:.0f}%",
                        fmt_signed(var) if j['vol_t'] > 0 else "—"])

# Total row
fam_data_vol.append(["TOTAL", fmt(total_juil['vol_t']), fmt(total_juil['vol_obj_t']), f"{total_juil['pct_vol']:.0f}%",
                    fmt(total_aout['vol_t']), fmt(total_aout['vol_obj_t']), f"{total_aout['pct_vol']:.0f}%",
                    fmt_signed(var_vol)])

wrapped_fam_vol = []
for ri, row in enumerate(fam_data_vol):
    prow = []
    for ci, cell in enumerate(row):
        if ri == 0:
            prow.append(wrap_cell(cell, CELL_HEADER))
        elif ci == 0:
            prow.append(wrap_cell(cell, CELL_LEFT))
        else:
            prow.append(wrap_cell(cell, CELL_CENTER))
    wrapped_fam_vol.append(prow)

t_fam_vol = Table(wrapped_fam_vol, colWidths=[2.8*cm, 1.7*cm, 1.7*cm, 1.3*cm, 1.7*cm, 1.7*cm, 1.3*cm, 2.3*cm], repeatRows=1)
style_fam_vol = [
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
    ('TOPPADDING', (0,0), (-1,-1), 4), ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ('LEFTPADDING', (0,0), (-1,-1), 3), ('RIGHTPADDING', (0,0), (-1,-1), 3),
    # Total row in yellow
    ('BACKGROUND', (0,len(fam_data_vol)-1), (-1,len(fam_data_vol)-1), colors.HexColor('#FFF2CC')),
    ('FONT', (0,len(fam_data_vol)-1), (-1,len(fam_data_vol)-1), 'DejaVuSans-Bold', 8),
]
# Color % cells based on realization
for i, fam in enumerate(families_order, 1):
    style_fam_vol.append(('BACKGROUND', (3, i), (3, i), color_for_pct(juil_data[fam]['pct_vol'])))
    style_fam_vol.append(('BACKGROUND', (6, i), (6, i), color_for_pct(aout_data[fam]['pct_vol'])))
t_fam_vol.setStyle(TableStyle(style_fam_vol))
story.append(t_fam_vol)
story.append(Spacer(1, 0.3*cm))

# CA table by family
story.append(Paragraph("<b>2.2 CA par famille (réalisé vs objectif)</b>", H3))

fam_data_ca = [["Famille", "CA Juil (M)", "Obj Juil (M)", "% Juil", "CA Août (M)", "Obj Août (M)", "% Août", "Δ CA Juil→Août"]]

for fam in families_order:
    j = juil_data[fam]
    a = aout_data[fam]
    var = (a['ca_m']/j['ca_m']-1)*100 if j['ca_m'] > 0 else 0
    fam_data_ca.append([fam, fmt(j['ca_m']), fmt(j['ca_obj_m']), f"{j['pct_ca']:.0f}%",
                       fmt(a['ca_m']), fmt(a['ca_obj_m']), f"{a['pct_ca']:.0f}%",
                       fmt_signed(var) if j['ca_m'] > 0 else "—"])

fam_data_ca.append(["TOTAL", fmt(total_juil['ca_m']), fmt(total_juil['ca_obj_m']), f"{total_juil['pct_ca']:.0f}%",
                   fmt(total_aout['ca_m']), fmt(total_aout['ca_obj_m']), f"{total_aout['pct_ca']:.0f}%",
                   fmt_signed(var_ca)])

wrapped_fam_ca = []
for ri, row in enumerate(fam_data_ca):
    prow = []
    for ci, cell in enumerate(row):
        if ri == 0:
            prow.append(wrap_cell(cell, CELL_HEADER))
        elif ci == 0:
            prow.append(wrap_cell(cell, CELL_LEFT))
        else:
            prow.append(wrap_cell(cell, CELL_CENTER))
    wrapped_fam_ca.append(prow)

t_fam_ca = Table(wrapped_fam_ca, colWidths=[2.8*cm, 1.7*cm, 1.7*cm, 1.3*cm, 1.7*cm, 1.7*cm, 1.3*cm, 2.3*cm], repeatRows=1)
style_fam_ca = [
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
    ('TOPPADDING', (0,0), (-1,-1), 4), ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ('LEFTPADDING', (0,0), (-1,-1), 3), ('RIGHTPADDING', (0,0), (-1,-1), 3),
    ('BACKGROUND', (0,len(fam_data_ca)-1), (-1,len(fam_data_ca)-1), colors.HexColor('#FFF2CC')),
    ('FONT', (0,len(fam_data_ca)-1), (-1,len(fam_data_ca)-1), 'DejaVuSans-Bold', 8),
]
for i, fam in enumerate(families_order, 1):
    style_fam_ca.append(('BACKGROUND', (3, i), (3, i), color_for_pct(juil_data[fam]['pct_ca'])))
    style_fam_ca.append(('BACKGROUND', (6, i), (6, i), color_for_pct(aout_data[fam]['pct_ca'])))
t_fam_ca.setStyle(TableStyle(style_fam_ca))
story.append(t_fam_ca)
story.append(Spacer(1, 0.3*cm))

# Family analysis
story.append(Paragraph("<b>Lecture par famille</b>", H3))
fam_analysis = """
<b>TOURTEAUX</b> : Chute majeure du volume (-44%) et du CA (-44%) entre juillet et août. Réalisation juillet : 194% (pic rupture concurrente) → août : 120% (toujours au-dessus de l'objectif). <b>Plus grande contribution à la baisse globale</b>.
<br/><b>CONCENTRES</b> : Baisse plus contenue (-10% vol, -9% CA). Réalisation juillet : 110%, août : 111%. <b>Stabilité relative vs objectif</b>.
<br/><b>INGREDIENTS</b> : Baisse légère (-3% vol, -5% CA). Juillet à 121%, août à 130% de l'objectif. <b>Sur-performance les deux mois</b>.
<br/><b>ALIMENT_COMPLET</b> : Baisse marquée (-19% vol, -18% CA). Juillet : 119%, août : 109% (recul sous la performance juillet).
<br/><b>PREMIX</b> : Baisse de -35% en volume. Juillet : 110%, août : 78% — <b>sous-objectif en août</b>.
<br/><b>MATERIEL_ELEVAGE</b> : Volumes nuls (prévu). CA +132% en août (ponctuel).
"""
story.append(Paragraph(fam_analysis, BODY))
story.append(Spacer(1, 0.3*cm))

# === 3. Par Région ===
story.append(PageBreak())
story.append(Paragraph("3. Vue par Région — Volumes + CA", H2))

# Compute regional aggregates
reg_data_juil = {}
reg_data_aout = {}
for reg in ['Ouest', 'Centre', 'Littoral']:
    j_reg = juil[juil['region'] == reg]
    a_reg = aout[aout['region'] == reg]
    reg_data_juil[reg] = {'vol_t': j_reg['tonnes'].sum(), 'ca_m': j_reg['ca_combined'].sum()/1e6}
    reg_data_aout[reg] = {'vol_t': a_reg['tonnes'].sum(), 'ca_m': a_reg['ca_combined'].sum()/1e6}

reg_table = [["Région", "Vol Juil (t)", "Vol Août (t)", "Δ Vol %", "CA Juil (M)", "CA Août (M)", "Δ CA %"]]
for reg in ['Ouest', 'Centre', 'Littoral']:
    j = reg_data_juil[reg]
    a = reg_data_aout[reg]
    var_v = (a['vol_t']/j['vol_t']-1)*100 if j['vol_t'] > 0 else 0
    var_c = (a['ca_m']/j['ca_m']-1)*100 if j['ca_m'] > 0 else 0
    reg_table.append([reg, fmt(j['vol_t']), fmt(a['vol_t']), fmt_signed(var_v),
                     fmt(j['ca_m']), fmt(a['ca_m']), fmt_signed(var_c)])

# Total
total_v_j = sum(r['vol_t'] for r in reg_data_juil.values())
total_v_a = sum(r['vol_t'] for r in reg_data_aout.values())
total_c_j = sum(r['ca_m'] for r in reg_data_juil.values())
total_c_a = sum(r['ca_m'] for r in reg_data_aout.values())
reg_table.append(["TOTAL", fmt(total_v_j), fmt(total_v_a), fmt_signed((total_v_a/total_v_j-1)*100),
                 fmt(total_c_j), fmt(total_c_a), fmt_signed((total_c_a/total_c_j-1)*100)])

wrapped_reg = []
for ri, row in enumerate(reg_table):
    prow = []
    for ci, cell in enumerate(row):
        if ri == 0:
            prow.append(wrap_cell(cell, CELL_HEADER))
        elif ci == 0:
            prow.append(wrap_cell(cell, CELL_LEFT))
        else:
            prow.append(wrap_cell(cell, CELL_CENTER))
    wrapped_reg.append(prow)

t_reg = Table(wrapped_reg, colWidths=[2.5*cm, 2.5*cm, 2.5*cm, 2*cm, 2.5*cm, 2.5*cm, 2*cm], repeatRows=1)
style_reg = [
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
    ('TOPPADDING', (0,0), (-1,-1), 4), ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ('LEFTPADDING', (0,0), (-1,-1), 4), ('RIGHTPADDING', (0,0), (-1,-1), 4),
    ('BACKGROUND', (0,len(reg_table)-1), (-1,len(reg_table)-1), colors.HexColor('#FFF2CC')),
    ('FONT', (0,len(reg_table)-1), (-1,len(reg_table)-1), 'DejaVuSans-Bold', 8),
]
# Color variation cells
for i, reg in enumerate(['Ouest', 'Centre', 'Littoral'], 1):
    var_v = (reg_data_aout[reg]['vol_t']/reg_data_juil[reg]['vol_t']-1)*100
    var_c = (reg_data_aout[reg]['ca_m']/reg_data_juil[reg]['ca_m']-1)*100
    style_reg.append(('BACKGROUND', (3, i), (3, i), color_for_var(var_v)))
    style_reg.append(('BACKGROUND', (6, i), (6, i), color_for_var(var_c)))
t_reg.setStyle(TableStyle(style_reg))
story.append(t_reg)
story.append(Spacer(1, 0.3*cm))

reg_analysis = """
<b>Lecture régionale</b> :
<br/>• <b>Ouest</b> : Plus forte baisse (-45% vol, -35% CA). FAMLA principalement responsable.
<br/>• <b>Centre</b> : Baisse la plus modérée (-27% vol, -17% CA). Région la plus résiliente.
<br/>• <b>Littoral</b> : Baisse marquée (-35% vol, -39% CA). Effet NDOBO (pic en juillet).
"""
story.append(Paragraph(reg_analysis, BODY))
story.append(Spacer(1, 0.3*cm))

# === 4. Insights & Conclusion ===
story.append(Paragraph("4. Insights & Conclusion", H2))

insights = [
    f"<b>Baisse globale marquée</b> : Volume -37,3% et CA -31,5% entre juillet et août 2026. La baisse volume est plus forte que la baisse CA → effet prix (hausse soja) qui amortit partiellement la baisse CA.",
    f"<b>Objectif août plus bas que juillet</b> : L'objectif volume est passé de {fmt(total_juil['vol_obj_t'])} t (juil) à {fmt(total_aout['vol_obj_t'])} t (août), soit {fmt_signed((total_aout['vol_obj_t']/total_juil['vol_obj_t']-1)*100)}. Idem pour le CA ({fmt_signed((total_aout['ca_obj_m']/total_juil['ca_obj_m']-1)*100)}).",
    f"<b>Réalisation vs objectif</b> : Volume juillet à {total_juil['pct_vol']:.0f}% de l'objectif, août à {total_aout['pct_vol']:.0f}%. CA juillet à {total_juil['pct_ca']:.0f}%, août à {total_aout['pct_ca']:.0f}%. Les deux mois dépassent l'objectif malgré la baisse.",
    f"<b>TOURTEAUX = principal contributeur</b> : 80% de la baisse globale vient du TOURTEAUX (-44% vol, -44% CA). Le pic juillet (194% de l'objectif) s'explique par la rupture concurrente ; août (120%) reste au-dessus de l'objectif malgré la normalisation.",
    f"<b>CONCENTRES = stabilité</b> : Baisse limitée à -10% vol / -9% CA. Réalisation stable à 110-111% de l'objectif. La discipline bundle (cross-sell soja+conc) a amorti la baisse.",
    f"<b>PREMIX en alerte</b> : Réalisation août à 78% de l'objectif (vs 110% en juillet). Baisse de -35% en volume. À surveiller.",
    f"<b>Ouest = région la plus touchée</b> (-45% vol, -35% CA), Littoral suit (-35% vol, -39% CA), Centre plus résilient (-27% vol, -17% CA).",
    f"<b>Contexte</b> : La baisse juillet→août ne traduit pas une dégradation structurelle mais une <b>normalisation après pic de juillet</b> (rupture concurrente). Les volumes août restent au-dessus des objectifs → performance saine malgré la baisse apparente.",
]
for ins in insights:
    story.append(Paragraph(f"• {ins}", BULLET))
story.append(Spacer(1, 0.3*cm))

# === Footer ===
story.append(HRFlowable(width="100%", thickness=0.5, color=GRAY))
story.append(Paragraph(
    "<b>Sources</b> : dataset_2023_2026.csv (records Livrées juillet + août 2026). "
    "<b>Objectifs</b> : ca_obj_real.json (CA mensuel par famille) + s2_recaled_objectives.json (volume mensuel par famille). "
    "<b>CA</b> : montant_ttc avec fallback sur montant_ht si TTC=0. "
    "<b>% réalisation</b> = (réalisé / objectif) × 100. Code couleur : vert ≥100%, jaune 80-100%, rouge <80%.",
    SMALL))

doc.build(story)
print(f"\n=== PDF GENERATED ===")
print(f"Path: {OUT}")
print(f"Size: {os.path.getsize(OUT)/1024:.0f} KB")
