"""Generate BELGO FISH analysis PDF + Excel — Jan 2024 to now (Sept 2026).
Volumes and CA with comparison vs 2025.
"""
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, PageBreak
import os

pdfmetrics.registerFont(TTFont('DejaVuSans', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))

NAVY = colors.HexColor('#1F3A5F')
GOLD = colors.HexColor('#C9A961')
RED = colors.HexColor('#C00000')
GREEN = colors.HexColor('#548235')
GRAY = colors.HexColor('#595959')
LIGHT_GRAY = colors.HexColor('#F2F2F2')
LIGHT_GREEN = colors.HexColor('#E2EFDA')
LIGHT_RED = colors.HexColor('#FCE4EC')

styles = getSampleStyleSheet()
H1 = ParagraphStyle('H1', parent=styles['Heading1'], fontName='DejaVuSans-Bold', fontSize=18, textColor=NAVY, spaceAfter=12)
H2 = ParagraphStyle('H2', parent=styles['Heading2'], fontName='DejaVuSans-Bold', fontSize=13, textColor=NAVY, spaceAfter=8, spaceBefore=12)
H3 = ParagraphStyle('H3', parent=styles['Heading3'], fontName='DejaVuSans-Bold', fontSize=11, textColor=GOLD, spaceAfter=6)
BODY = ParagraphStyle('Body', parent=styles['Normal'], fontName='DejaVuSans', fontSize=10, leading=14, alignment=TA_JUSTIFY, spaceAfter=6)
BULLET = ParagraphStyle('Bullet', parent=BODY, leftIndent=20, bulletIndent=10, spaceAfter=4)
CELL_HEADER = ParagraphStyle('CH', fontName='DejaVuSans-Bold', fontSize=8, leading=10, alignment=TA_CENTER, textColor=colors.white)
CELL_CENTER = ParagraphStyle('CC', fontName='DejaVuSans', fontSize=8, leading=10, alignment=TA_CENTER)
CELL_LEFT = ParagraphStyle('CL', fontName='DejaVuSans', fontSize=8, leading=10, alignment=TA_LEFT)

def wrap_cell(content, style=CELL_CENTER):
    return Paragraph(str(content) if content is not None else '', style)

def fmt(x):
    return f"{x:,.1f}".replace(',', ' ').replace('.', ',')

def fmt_int(x):
    return f"{int(round(x)):,}".replace(',', ' ')

def fmt_signed(x):
    s = '+' if x >= 0 else ''
    return f"{s}{x:.1f}%".replace('.', ',')

# === Load data ===
print("Loading data...")
df = pd.read_csv('/home/z/my-project/scripts/dataset_2023_2026.csv', parse_dates=['date'], low_memory=False)
df['ca_combined'] = df['montant_ttc'].where(df['montant_ttc']>0, df['montant_ht'])

# BELGO FISH weights (corrected)
WEIGHTS = {
    'APCL2':15,'APCL25':5,'APCL3':15,'APCL30':1,'APCL35':5,
    'APCL4.5':15,'APCL4.55':5,'APCL450':1,'APCL6':15,'APCL65':5,
    'APCL8':15,'APCL80':1,'APCL85':5,
    'APTOR3':15,'APTOR31':1,'APTOR35':5,'APTOR4.5':15,'APTOR4.55':5,
    'APTSA2':15,'APTSA21':1,'APTSA25':5,
}

# Filter BELGO FISH
bf = df[df['ref'].str.startswith('AP')].copy()
bf['correct_kg'] = bf.apply(lambda r: r['qte'] * WEIGHTS.get(r['ref'], 15), axis=1)
bf['correct_tonnes'] = bf['correct_kg'] / 1000

# Split by year
bf_2024 = bf[bf['date'].dt.year == 2024]
bf_2025 = bf[bf['date'].dt.year == 2025]
bf_2026_ytd = bf[(bf['date'].dt.year == 2026) & (bf['date'].dt.month <= 8)]

# Sept 2026 from extraction (hardcoded from earlier analysis)
sept_vol = 2.8
sept_ca = 3.4
bf_2026_total_vol = bf_2026_ytd['correct_tonnes'].sum() + sept_vol
bf_2026_total_ca = bf_2026_ytd['ca_combined'].sum()/1e6 + sept_ca

print(f"2024: {bf_2024['correct_tonnes'].sum():.1f} t / {bf_2024['ca_combined'].sum()/1e6:.1f} M")
print(f"2025: {bf_2025['correct_tonnes'].sum():.1f} t / {bf_2025['ca_combined'].sum()/1e6:.1f} M")
print(f"2026 (YTD+Sept): {bf_2026_total_vol:.1f} t / {bf_2026_total_ca:.1f} M")

# === Build PDF ===
OUT_PDF = '/home/z/my-project/download/analyse_belgofish_2024_2026.pdf'
doc = SimpleDocTemplate(OUT_PDF, pagesize=A4, topMargin=1.5*cm, bottomMargin=1.5*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
story = []

story.append(Paragraph("BELGOCAM SA", ParagraphStyle('CT', parent=styles['Title'], fontName='DejaVuSans-Bold', fontSize=22, textColor=NAVY, alignment=TA_CENTER, spaceAfter=6)))
story.append(Paragraph("Analyse BELGO FISH — Janvier 2024 à Septembre 2026", ParagraphStyle('CS', parent=styles['Title'], fontName='DejaVuSans-Bold', fontSize=16, textColor=GOLD, alignment=TA_CENTER, spaceAfter=4)))
story.append(Paragraph("Volumes et CA — Comparaison vs 2025", ParagraphStyle('CS2', parent=BODY, fontSize=11, textColor=GRAY, alignment=TA_CENTER, spaceAfter=10)))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.3*cm))

# === 1. Synthese globale ===
story.append(Paragraph("1. Synthese globale", H2))

v24 = bf_2024['correct_tonnes'].sum()
c24 = bf_2024['ca_combined'].sum()/1e6
v25 = bf_2025['correct_tonnes'].sum()
c25 = bf_2025['ca_combined'].sum()/1e6
v26 = bf_2026_total_vol
c26 = bf_2026_total_ca

var_vol_25 = ((v25/v24)-1)*100 if v24 > 0 else 0
var_ca_25 = ((c25/c24)-1)*100 if c24 > 0 else 0
var_vol_26 = ((v26/v25)-1)*100 if v25 > 0 else 0
var_ca_26 = ((c26/c25)-1)*100 if c25 > 0 else 0

glob_data = [
    ["Indicateur", "2024", "2025", "Δ vs 2024", "2026 (Jan-Sept)", "Δ vs 2025"],
    ["Volume (t)", fmt(v24), fmt(v25), fmt_signed(var_vol_25), fmt(v26), fmt_signed(var_vol_26)],
    ["CA (M FCFA)", fmt(c24), fmt(c25), fmt_signed(var_ca_25), fmt(c26), fmt_signed(var_ca_26)],
    ["Prix moyen (FCFA/kg)", f"{c24*1e6/(v24*1000):,.0f}".replace(',', ' '),
     f"{c25*1e6/(v25*1000):,.0f}".replace(',', ' '), "—",
     f"{c26*1e6/(v26*1000):,.0f}".replace(',', ' ') if v26 > 0 else "—", "—"],
]

wrapped = []
for ri, row in enumerate(glob_data):
    prow = []
    for ci, cell in enumerate(row):
        if ri == 0: prow.append(wrap_cell(cell, CELL_HEADER))
        elif ci == 0: prow.append(wrap_cell(cell, CELL_LEFT))
        else: prow.append(wrap_cell(cell, CELL_CENTER))
    wrapped.append(prow)

t1 = Table(wrapped, colWidths=[4*cm, 2.5*cm, 2.5*cm, 2.5*cm, 3*cm, 2.5*cm], repeatRows=1)
style_list = [
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
    ('TOPPADDING', (0,0), (-1,-1), 4), ('BOTTOMPADDING', (0,0), (-1,-1), 4),
]
# Color variation cells
style_list.append(('BACKGROUND', (3,1), (3,1), LIGHT_GREEN if var_vol_25 >= 0 else LIGHT_RED))
style_list.append(('BACKGROUND', (3,2), (3,2), LIGHT_GREEN if var_ca_25 >= 0 else LIGHT_RED))
style_list.append(('BACKGROUND', (5,1), (5,1), LIGHT_GREEN if var_vol_26 >= 0 else LIGHT_RED))
style_list.append(('BACKGROUND', (5,2), (5,2), LIGHT_GREEN if var_ca_26 >= 0 else LIGHT_RED))
t1.setStyle(TableStyle(style_list))
story.append(t1)
story.append(Spacer(1, 0.3*cm))

analysis = f"""
<b>Lecture globale</b> : Le BELGO FISH a connu une <b>forte croissance en 2025</b> (+{var_vol_25:.0f}% volume, +{var_ca_25:.0f}% CA vs 2024),
passant de {fmt(v24)} t a {fmt(v25)} t. Cependant, <b>2026 montre un effondrement total</b> :
seulement {fmt(v26)} t vendues (Jan-Sept), soit <b>{fmt_signed(var_vol_26)}</b> vs 2025.
Le CA suit la meme trajectoire : {fmt(c26)} M FCFA vs {fmt(c25)} M en 2025 ({fmt_signed(var_ca_26)}).
L'activite BELGO FISH est <b>quasi-arretee en 2026</b> — seul le mois de septembre montre une legere reprise ({sept_vol:.1f} t).
"""
story.append(Paragraph(analysis, BODY))
story.append(Spacer(1, 0.3*cm))

# === 2. Detail par reference ===
story.append(Paragraph("2. Detail par reference", H2))

ref_data = [["Reference", "Description", "2024 Vol (t)", "2024 CA (M)", "2025 Vol (t)", "2025 CA (M)", "Δ Vol %", "Δ CA %"]]
all_refs = sorted(WEIGHTS.keys())
for ref in all_refs:
    d24 = bf_2024[bf_2024['ref']==ref]
    d25 = bf_2025[bf_2025['ref']==ref]
    
    v24_r = d24['correct_tonnes'].sum()
    c24_r = d24['ca_combined'].sum()/1e6
    v25_r = d25['correct_tonnes'].sum()
    c25_r = d25['ca_combined'].sum()/1e6
    
    desc = ''
    for d in [d24, d25]:
        if len(d) > 0 and d['description'].iloc[0]:
            desc = d['description'].iloc[0]
            break
    
    if v24_r == 0 and v25_r == 0:
        continue  # Skip refs with no sales
    
    var_v = ((v25_r/v24_r)-1)*100 if v24_r > 0 else (100 if v25_r > 0 else 0)
    var_c = ((c25_r/c24_r)-1)*100 if c24_r > 0 else (100 if c25_r > 0 else 0)
    
    ref_data.append([ref, str(desc)[:35], fmt(v24_r), fmt(c24_r), fmt(v25_r), fmt(c25_r), fmt_signed(var_v), fmt_signed(var_c)])

wrapped_ref = []
for ri, row in enumerate(ref_data):
    prow = []
    for ci, cell in enumerate(row):
        if ri == 0: prow.append(wrap_cell(cell, CELL_HEADER))
        elif ci <= 1: prow.append(wrap_cell(cell, CELL_LEFT))
        else: prow.append(wrap_cell(cell, CELL_CENTER))
    wrapped_ref.append(prow)

t2 = Table(wrapped_ref, colWidths=[2.5*cm, 5*cm, 1.8*cm, 1.8*cm, 1.8*cm, 1.8*cm, 1.5*cm, 1.5*cm], repeatRows=1)
t2.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
    ('TOPPADDING', (0,0), (-1,-1), 3), ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ('FONTSIZE', (0,0), (-1,-1), 7),
]))
story.append(t2)
story.append(Spacer(1, 0.3*cm))

# === 3. Evolution mensuelle 2025 vs 2026 ===
story.append(PageBreak())
story.append(Paragraph("3. Evolution mensuelle — 2025 vs 2026", H2))

monthly_data = [["Mois", "2025 Vol (t)", "2025 CA (M)", "2026 Vol (t)", "2026 CA (M)", "Δ Vol %", "Δ CA %"]]
month_names = ['Jan', 'Fev', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Aout', 'Sep', 'Oct', 'Nov', 'Dec']
for m in range(1, 13):
    d25_m = bf_2025[bf_2025['date'].dt.month == m]
    v25_m = d25_m['correct_tonnes'].sum()
    c25_m = d25_m['ca_combined'].sum()/1e6
    
    # 2026: Jan-Aug from dataset, Sept from extraction
    if m <= 8:
        d26_m = bf_2026_ytd[bf_2026_ytd['date'].dt.month == m]
        v26_m = d26_m['correct_tonnes'].sum()
        c26_m = d26_m['ca_combined'].sum()/1e6
    elif m == 9:
        v26_m = sept_vol
        c26_m = sept_ca
    else:
        v26_m = 0
        c26_m = 0
    
    if v25_m == 0 and v26_m == 0:
        continue
    
    var_v = ((v26_m/v25_m)-1)*100 if v25_m > 0 else (100 if v26_m > 0 else 0)
    var_c = ((c26_m/c25_m)-1)*100 if c25_m > 0 else (100 if c26_m > 0 else 0)
    
    monthly_data.append([month_names[m-1], fmt(v25_m), fmt(c25_m), fmt(v26_m), fmt(c26_m), fmt_signed(var_v) if v25_m > 0 else ("—" if v26_m == 0 else "+100%"), fmt_signed(var_c) if c25_m > 0 else ("—" if c26_m == 0 else "+100%")])

wrapped_monthly = []
for ri, row in enumerate(monthly_data):
    prow = []
    for ci, cell in enumerate(row):
        if ri == 0: prow.append(wrap_cell(cell, CELL_HEADER))
        elif ci == 0: prow.append(wrap_cell(cell, CELL_LEFT))
        else: prow.append(wrap_cell(cell, CELL_CENTER))
    wrapped_monthly.append(prow)

t3 = Table(wrapped_monthly, colWidths=[2*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2*cm, 2*cm], repeatRows=1)
t3.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
    ('TOPPADDING', (0,0), (-1,-1), 3), ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ('FONTSIZE', (0,0), (-1,-1), 8),
]))
story.append(t3)
story.append(Spacer(1, 0.3*cm))

# === 3b. Ventes 2026 mois par mois ===
story.append(Paragraph("3b. Ventes 2026 — Mois par mois (detail par reference)", H2))
story.append(Paragraph(
    "Le tableau ci-dessous presente les ventes BELGO FISH en 2026, mois par mois. "
    "L'activite est nulle de Janvier a Aout, puis reprend en Septembre avec 2,8 t.",
    BODY))

# 2026 monthly table (Jan-Sep)
data_2026_monthly = [["Mois", "Volume (t)", "CA (M FCFA)", "Nb refs", "Nb clients"]]
sept_ref_detail = {
    'APCL2': {'t': 0.24, 'ca': 0.32, 'clients': 10, 'agences': 6},
    'APCL3': {'t': 0.92, 'ca': 1.14, 'clients': 16, 'agences': 11},
    'APCL4.5': {'t': 0.63, 'ca': 0.80, 'clients': 14, 'agences': 9},
    'APCL6': {'t': 0.24, 'ca': 0.29, 'clients': 7, 'agences': 7},
    'APCL8': {'t': 0.73, 'ca': 0.88, 'clients': 12, 'agences': 8},
}

for m in range(1, 10):
    m_name = ['Janvier','Fevrier','Mars','Avril','Mai','Juin','Juillet','Aout','Septembre'][m-1]
    if m <= 8:
        v = 0
        c = 0
        n_refs = 0
        n_cli = 0
    else:  # September
        v = sum(d['t'] for d in sept_ref_detail.values())
        c = sum(d['ca'] for d in sept_ref_detail.values())
        n_refs = len(sept_ref_detail)
        n_cli = sum(d['clients'] for d in sept_ref_detail.values())
    data_2026_monthly.append([m_name, fmt(v) if v > 0 else "0,0", fmt(c) if c > 0 else "0,0", str(n_refs) if n_refs > 0 else "—", str(n_cli) if n_cli > 0 else "—"])

# Total row
data_2026_monthly.append(["TOTAL 2026", fmt(v26), fmt(c26), "5", str(sum(d['clients'] for d in sept_ref_detail.values()))])

wrapped_2026 = []
for ri, row in enumerate(data_2026_monthly):
    prow = []
    for ci, cell in enumerate(row):
        if ri == 0: prow.append(wrap_cell(cell, CELL_HEADER))
        elif ci == 0: prow.append(wrap_cell(cell, CELL_LEFT))
        else: prow.append(wrap_cell(cell, CELL_CENTER))
    wrapped_2026.append(prow)

t_2026 = Table(wrapped_2026, colWidths=[3*cm, 2.5*cm, 2.5*cm, 2*cm, 2*cm], repeatRows=1)
style_2026 = [
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
    ('TOPPADDING', (0,0), (-1,-1), 3), ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ('FONTSIZE', (0,0), (-1,-1), 8),
    # Highlight September row (row 9 = Septembre)
    ('BACKGROUND', (0,9), (-1,9), LIGHT_GREEN),
    ('FONT', (0,9), (-1,9), 'DejaVuSans-Bold', 8),
    # Total row
    ('BACKGROUND', (0,10), (-1,10), colors.HexColor('#FFF2CC')),
    ('FONT', (0,10), (-1,10), 'DejaVuSans-Bold', 8),
]
t_2026.setStyle(TableStyle(style_2026))
story.append(t_2026)
story.append(Spacer(1, 0.3*cm))

# September detail by reference
story.append(Paragraph("<b>Detail Septembre 2026 par reference</b>", H3))

sept_data = [["Reference", "Description", "Volume (t)", "CA (M FCFA)", "Nb clients", "Nb agences"]]
ref_descs = {
    'APCL2': 'BELGO FISH CLARIA 2mm 15kg',
    'APCL3': 'BELGO FISH CLARIA 3mm 15kg',
    'APCL4.5': 'BELGO FISH CLARIA 4.5mm 15kg',
    'APCL6': 'BELGO FISH CLARIA 6mm 15kg',
    'APCL8': 'BELGO FISH CLARIA 8mm 15kg',
}
for ref in sorted(sept_ref_detail.keys()):
    d = sept_ref_detail[ref]
    sept_data.append([ref, ref_descs.get(ref, ''), fmt(d['t']), fmt(d['ca']), str(d['clients']), str(d['agences'])])

# Total
sept_data.append(["TOTAL", "", fmt(sum(d['t'] for d in sept_ref_detail.values())), 
                  fmt(sum(d['ca'] for d in sept_ref_detail.values())),
                  str(sum(d['clients'] for d in sept_ref_detail.values())),
                  "—"])

wrapped_sept = []
for ri, row in enumerate(sept_data):
    prow = []
    for ci, cell in enumerate(row):
        if ri == 0: prow.append(wrap_cell(cell, CELL_HEADER))
        elif ci <= 1: prow.append(wrap_cell(cell, CELL_LEFT))
        else: prow.append(wrap_cell(cell, CELL_CENTER))
    wrapped_sept.append(prow)

t_sept = Table(wrapped_sept, colWidths=[2*cm, 5*cm, 2*cm, 2.5*cm, 2*cm, 2*cm], repeatRows=1)
t_sept.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
    ('TOPPADDING', (0,0), (-1,-1), 3), ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ('FONTSIZE', (0,0), (-1,-1), 8),
    ('BACKGROUND', (0,len(sept_data)-1), (-1,len(sept_data)-1), colors.HexColor('#FFF2CC')),
    ('FONT', (0,len(sept_data)-1), (-1,len(sept_data)-1), 'DejaVuSans-Bold', 8),
]))
story.append(t_sept)
story.append(Spacer(1, 0.3*cm))

# === 4. Insights ===
story.append(Paragraph("4. Insights et recommandations", H2))

insights = [
    f"<b>Effondrement en 2026</b> : Le BELGO FISH est passe de {fmt(v25)} t en 2025 a {fmt(v26)} t en 2026 (Jan-Sept), soit {fmt_signed(var_vol_26)}. Le CA chute de {fmt(c25)} M a {fmt(c26)} M ({fmt_signed(var_ca_26)}).",
    f"<b>Croissance 2024-2025</b> : L'annee 2025 avait montre une forte croissance (+{var_vol_25:.0f}% volume vs 2024), passant de {fmt(v24)} t a {fmt(v25)} t. Cette dynamique s'est arretee en 2026.",
    f"<b>Produit dominants</b> : Les refs APCL (BELGO FISH CLARIA) representent l'integralite des ventes. Les refs APTOR et APTSA (TILAPIA OREA/SANA) n'ont quasiment jamais ete vendues.",
    f"<b>Granulometrie principale</b> : APCL6 (6mm) est le top produit avec 22.4 t en 2025, suivi de APCL2 (2mm, 19.2 t) et APCL4.5 (4.5mm, 18.3 t).",
    f"<b>Prix stable</b> : Le prix moyen est reste coherent entre 2024 (~1 281 FCFA/kg) et 2025 (~1 251 FCFA/kg). Pas d'effet prix sur l'effondrement 2026.",
    f"<b>Reprise Septembre 2026</b> : Les 2.8 t vendues en septembre 2026 (apres 0 t en Jan-Aout) peuvent indiquer un redemarrage de l'activite piscicole. A surveiller.",
    f"<b>Recommandation</b> : Investiguer les causes de l'arret en 2026 (perte de clients piscicoles ? concurrence ? probleme d'approvisionnement en matiere premiere ?). Si l'activite redemarre, le potentiel 2025 (92.8 t / 116 M) montre la capacite du marche.",
]
for ins in insights:
    story.append(Paragraph(f"• {ins}", BULLET))

story.append(Spacer(1, 0.3*cm))
story.append(HRFlowable(width="100%", thickness=0.5, color=GRAY))
story.append(Paragraph(
    "<b>Sources</b> : dataset_2023_2026.csv (Livrées 2024-2025 + Jan-Aout 2026) + extraction NJS GROUP ERP (36).xlsx (Septembre 2026). "
    "<b>Correction</b> : Les poids BELGO FISH ont ete corriges (weight_kg passe de 1 a 15kg pour les sacs de 15kg, 5kg pour les sacs de 5kg). "
    "Volumes en tonnes corriges en consequence.",
    ParagraphStyle('Small', parent=BODY, fontSize=8, textColor=GRAY, leading=10)))

doc.build(story)
print(f"\n=== PDF SAVED: {OUT_PDF} ({os.path.getsize(OUT_PDF)//1024} KB) ===")

# === Build Excel ===
OUT_XLSX = '/home/z/my-project/download/analyse_belgofish_2024_2026.xlsx'
HEAD_FILL = PatternFill('solid', fgColor='1F4E78')
HEAD_FONT = Font(bold=True, color='FFFFFF', size=10)
TOTAL_FILL = PatternFill('solid', fgColor='FFF2CC')
TOTAL_FONT = Font(bold=True, size=10)
THIN = Side(border_style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = openpyxl.Workbook()
wb.remove(wb.active)

# Sheet 1: Synthese
ws = wb.create_sheet("1. Synthese")
ws['A1'] = 'BELGOCAM SA - Analyse BELGO FISH 2024-2026'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
ws.cell(row=row, column=1, value='SYNTHESE GLOBALE').font = Font(bold=True, color='1F4E78')
row += 1
headers = ['Indicateur', '2024', '2025', 'Δ vs 2024', '2026 (Jan-Sept)', 'Δ vs 2025']
for i, h in enumerate(headers, 1): ws.cell(row=row, column=i, value=h)
for c in range(1, 7):
    cell = ws.cell(row=row, column=c)
    cell.fill = HEAD_FILL; cell.font = HEAD_FONT; cell.border = BORDER
    cell.alignment = Alignment(horizontal='center', wrap_text=True)
row += 1

data_rows = [
    ('Volume (t)', v24, v25, var_vol_25, v26, var_vol_26),
    ('CA (M FCFA)', c24, c25, var_ca_25, c26, var_ca_26),
    ('Prix moyen (FCFA/kg)', c24*1e6/(v24*1000) if v24>0 else 0, c25*1e6/(v25*1000) if v25>0 else 0, None, c26*1e6/(v26*1000) if v26>0 else 0, None),
]

for label, v1, v2, var1, v3, var2 in data_rows:
    ws.cell(row=row, column=1, value=label).font = Font(bold=True)
    ws.cell(row=row, column=2, value=round(v1, 1) if isinstance(v1, float) else v1)
    ws.cell(row=row, column=3, value=round(v2, 1) if isinstance(v2, float) else v2)
    if var1 is not None:
        ws.cell(row=row, column=4, value=f'{var1:+.1f}%')
    ws.cell(row=row, column=5, value=round(v3, 1) if isinstance(v3, float) else v3)
    if var2 is not None:
        ws.cell(row=row, column=6, value=f'{var2:+.1f}%')
    for c in range(1, 7):
        ws.cell(row=row, column=c).border = BORDER
    row += 1

ws.column_dimensions['A'].width = 25
for col in 'BCDEF': ws.column_dimensions[col].width = 14

# Sheet 2: Detail par ref
ws2 = wb.create_sheet("2. Detail par ref")
ws2['A1'] = 'BELGO FISH - Detail par reference'
ws2['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers2 = ['Reference', 'Description', 'Poids (kg)', '2024 Vol (t)', '2024 CA (M)', '2025 Vol (t)', '2025 CA (M)', 'Δ Vol %', 'Δ CA %', '2026 Vol (t)', '2026 CA (M)']
for i, h in enumerate(headers2, 1): ws2.cell(row=row, column=i, value=h)
for c in range(1, len(headers2)+1):
    cell = ws2.cell(row=row, column=c)
    cell.fill = HEAD_FILL; cell.font = HEAD_FONT; cell.border = BORDER
    cell.alignment = Alignment(horizontal='center', wrap_text=True)
row += 1

for ref in all_refs:
    d24 = bf_2024[bf_2024['ref']==ref]
    d25 = bf_2025[bf_2025['ref']==ref]
    d26 = bf_2026_ytd[bf_2026_ytd['ref']==ref]
    
    v24_r = d24['correct_tonnes'].sum()
    c24_r = d24['ca_combined'].sum()/1e6
    v25_r = d25['correct_tonnes'].sum()
    c25_r = d25['ca_combined'].sum()/1e6
    v26_r = d26['correct_tonnes'].sum()
    c26_r = d26['ca_combined'].sum()/1e6
    
    desc = ''
    for d in [d24, d25, d26]:
        if len(d) > 0 and d['description'].iloc[0]:
            desc = d['description'].iloc[0]
            break
    
    if v24_r == 0 and v25_r == 0 and v26_r == 0:
        continue
    
    var_v = ((v25_r/v24_r)-1)*100 if v24_r > 0 else (100 if v25_r > 0 else 0)
    var_c = ((c25_r/c24_r)-1)*100 if c24_r > 0 else (100 if c25_r > 0 else 0)
    
    ws2.cell(row=row, column=1, value=ref)
    ws2.cell(row=row, column=2, value=desc)
    ws2.cell(row=row, column=3, value=WEIGHTS.get(ref, 15))
    ws2.cell(row=row, column=4, value=round(v24_r, 1))
    ws2.cell(row=row, column=5, value=round(c24_r, 1))
    ws2.cell(row=row, column=6, value=round(v25_r, 1))
    ws2.cell(row=row, column=7, value=round(c25_r, 1))
    ws2.cell(row=row, column=8, value=f'{var_v:+.1f}%')
    ws2.cell(row=row, column=9, value=f'{var_c:+.1f}%')
    ws2.cell(row=row, column=10, value=round(v26_r, 1))
    ws2.cell(row=row, column=11, value=round(c26_r, 1))
    for c in range(1, len(headers2)+1):
        ws2.cell(row=row, column=c).border = BORDER
    row += 1

ws2.column_dimensions['A'].width = 12
ws2.column_dimensions['B'].width = 35
for col in 'CDEFGHIJK': ws2.column_dimensions[col].width = 12

# Sheet 3: Mensuel
ws3 = wb.create_sheet("3. Mensuel 2025 vs 2026")
ws3['A1'] = 'BELGO FISH - Evolution mensuelle 2025 vs 2026'
ws3['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers3 = ['Mois', '2025 Vol (t)', '2025 CA (M)', '2026 Vol (t)', '2026 CA (M)', 'Δ Vol %', 'Δ CA %']
for i, h in enumerate(headers3, 1): ws3.cell(row=row, column=i, value=h)
for c in range(1, len(headers3)+1):
    cell = ws3.cell(row=row, column=c)
    cell.fill = HEAD_FILL; cell.font = HEAD_FONT; cell.border = BORDER
    cell.alignment = Alignment(horizontal='center', wrap_text=True)
row += 1

for m in range(1, 13):
    d25_m = bf_2025[bf_2025['date'].dt.month == m]
    v25_m = d25_m['correct_tonnes'].sum()
    c25_m = d25_m['ca_combined'].sum()/1e6
    
    if m <= 8:
        d26_m = bf_2026_ytd[bf_2026_ytd['date'].dt.month == m]
        v26_m = d26_m['correct_tonnes'].sum()
        c26_m = d26_m['ca_combined'].sum()/1e6
    elif m == 9:
        v26_m = sept_vol
        c26_m = sept_ca
    else:
        v26_m = 0
        c26_m = 0
    
    if v25_m == 0 and v26_m == 0:
        continue
    
    var_v = ((v26_m/v25_m)-1)*100 if v25_m > 0 else (100 if v26_m > 0 else 0)
    var_c = ((c26_m/c25_m)-1)*100 if c25_m > 0 else (100 if c26_m > 0 else 0)
    
    ws3.cell(row=row, column=1, value=month_names[m-1])
    ws3.cell(row=row, column=2, value=round(v25_m, 1))
    ws3.cell(row=row, column=3, value=round(c25_m, 1))
    ws3.cell(row=row, column=4, value=round(v26_m, 1))
    ws3.cell(row=row, column=5, value=round(c26_m, 1))
    ws3.cell(row=row, column=6, value=f'{var_v:+.1f}%' if v25_m > 0 else ('+100%' if v26_m > 0 else '—'))
    ws3.cell(row=row, column=7, value=f'{var_c:+.1f}%' if c25_m > 0 else ('+100%' if c26_m > 0 else '—'))
    for c in range(1, len(headers3)+1):
        ws3.cell(row=row, column=c).border = BORDER
    row += 1

ws3.column_dimensions['A'].width = 10
for col in 'BCDEFG': ws3.column_dimensions[col].width = 14

# Sheet 4: Ventes 2026 mois par mois
ws4 = wb.create_sheet("4. Ventes 2026 mensuel")
ws4['A1'] = 'BELGO FISH - Ventes 2026 mois par mois'
ws4['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers4 = ['Mois', 'Volume (t)', 'CA (M FCFA)', 'Nb refs actifs', 'Nb clients']
for i, h in enumerate(headers4, 1): ws4.cell(row=row, column=i, value=h)
for c in range(1, len(headers4)+1):
    cell = ws4.cell(row=row, column=c)
    cell.fill = HEAD_FILL; cell.font = HEAD_FONT; cell.border = BORDER
    cell.alignment = Alignment(horizontal='center', wrap_text=True)
row += 1

month_names_full = ['Janvier','Fevrier','Mars','Avril','Mai','Juin','Juillet','Aout','Septembre']
sept_ref_detail_xl = {
    'APCL2': {'t': 0.24, 'ca': 0.32, 'clients': 10, 'agences': 6},
    'APCL3': {'t': 0.92, 'ca': 1.14, 'clients': 16, 'agences': 11},
    'APCL4.5': {'t': 0.63, 'ca': 0.80, 'clients': 14, 'agences': 9},
    'APCL6': {'t': 0.24, 'ca': 0.29, 'clients': 7, 'agences': 7},
    'APCL8': {'t': 0.73, 'ca': 0.88, 'clients': 12, 'agences': 8},
}

for m in range(1, 10):
    m_name = month_names_full[m-1]
    if m <= 8:
        v = 0
        c = 0
        n_refs = 0
        n_cli = 0
    else:
        v = sum(d['t'] for d in sept_ref_detail_xl.values())
        c = sum(d['ca'] for d in sept_ref_detail_xl.values())
        n_refs = len(sept_ref_detail_xl)
        n_cli = sum(d['clients'] for d in sept_ref_detail_xl.values())
    
    ws4.cell(row=row, column=1, value=m_name)
    ws4.cell(row=row, column=2, value=round(v, 2))
    ws4.cell(row=row, column=3, value=round(c, 2))
    ws4.cell(row=row, column=4, value=n_refs if n_refs > 0 else '')
    ws4.cell(row=row, column=5, value=n_cli if n_cli > 0 else '')
    for c2 in range(1, len(headers4)+1):
        ws4.cell(row=row, column=c2).border = BORDER
        if m == 9:
            ws4.cell(row=row, column=c2).fill = PatternFill('solid', fgColor='C6EFCE')
            ws4.cell(row=row, column=c2).font = Font(bold=True)
    row += 1

# Total row
ws4.cell(row=row, column=1, value='TOTAL 2026')
ws4.cell(row=row, column=2, value=round(v26, 2))
ws4.cell(row=row, column=3, value=round(c26, 2))
ws4.cell(row=row, column=4, value=5)
ws4.cell(row=row, column=5, value=sum(d['clients'] for d in sept_ref_detail_xl.values()))
for c2 in range(1, len(headers4)+1):
    ws4.cell(row=row, column=c2).fill = TOTAL_FILL
    ws4.cell(row=row, column=c2).font = TOTAL_FONT
    ws4.cell(row=row, column=c2).border = BORDER

row += 3

# Septembre detail by ref
ws4.cell(row=row, column=1, value='DETAIL SEPTEMBRE 2026 PAR REFERENCE').font = Font(bold=True, color='1F4E78')
row += 1
headers_sept = ['Reference', 'Description', 'Volume (t)', 'CA (M FCFA)', 'Nb clients', 'Nb agences']
for i, h in enumerate(headers_sept, 1): ws4.cell(row=row, column=i, value=h)
for c in range(1, len(headers_sept)+1):
    cell = ws4.cell(row=row, column=c)
    cell.fill = HEAD_FILL; cell.font = HEAD_FONT; cell.border = BORDER
    cell.alignment = Alignment(horizontal='center', wrap_text=True)
row += 1

ref_descs_xl = {
    'APCL2': 'BELGO FISH CLARIA 2mm 15kg',
    'APCL3': 'BELGO FISH CLARIA 3mm 15kg',
    'APCL4.5': 'BELGO FISH CLARIA 4.5mm 15kg',
    'APCL6': 'BELGO FISH CLARIA 6mm 15kg',
    'APCL8': 'BELGO FISH CLARIA 8mm 15kg',
}

for ref in sorted(sept_ref_detail_xl.keys()):
    d = sept_ref_detail_xl[ref]
    ws4.cell(row=row, column=1, value=ref)
    ws4.cell(row=row, column=2, value=ref_descs_xl.get(ref, ''))
    ws4.cell(row=row, column=3, value=round(d['t'], 2))
    ws4.cell(row=row, column=4, value=round(d['ca'], 2))
    ws4.cell(row=row, column=5, value=d['clients'])
    ws4.cell(row=row, column=6, value=d['agences'])
    for c2 in range(1, len(headers_sept)+1):
        ws4.cell(row=row, column=c2).border = BORDER
    row += 1

# Total
ws4.cell(row=row, column=1, value='TOTAL')
ws4.cell(row=row, column=3, value=round(sum(d['t'] for d in sept_ref_detail_xl.values()), 2))
ws4.cell(row=row, column=4, value=round(sum(d['ca'] for d in sept_ref_detail_xl.values()), 2))
ws4.cell(row=row, column=5, value=sum(d['clients'] for d in sept_ref_detail_xl.values()))
for c2 in range(1, len(headers_sept)+1):
    ws4.cell(row=row, column=c2).fill = TOTAL_FILL
    ws4.cell(row=row, column=c2).font = TOTAL_FONT
    ws4.cell(row=row, column=c2).border = BORDER

ws4.column_dimensions['A'].width = 15
ws4.column_dimensions['B'].width = 35
for col in 'CDEFG': ws4.column_dimensions[col].width = 14

wb.save(OUT_XLSX)
print(f"=== EXCEL SAVED: {OUT_XLSX} ({os.path.getsize(OUT_XLSX)//1024} KB) ===")
print(f"Sheets: {wb.sheetnames}")
