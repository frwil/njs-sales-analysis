"""
Rapport de Performance Mensuelle — BELGOCAM SA / NJS GROUP (LIVRABLE MENSUEL)
Mois : Septembre 2026 (données complètes 01-30/09/2026)

6 sections :
 1. Performance globale par famille vs objectifs de septembre
 2. Performance par agence × famille vs objectifs de septembre
 3. Performance par région × famille vs objectifs de septembre
 4. Performance (tendance mensuelle) par famille vs objectifs Jan → Sept (YTD)
 5. Performance (tendance mensuelle) par agence × (Soja & Concentrés) YTD vs objectifs
 6. Performance (tendance mensuelle) par région × (Soja & Concentrés) YTD vs objectifs

Usage mensuel : mettre à jour DATA_PATH + MONTH_LABEL puis relancer.
"""
import os
import json
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle,
    HRFlowable, KeepTogether, Image
)

# ============================================================
# PARAMÈTRES DU MOIS
# ============================================================
MONTH_LABEL = "Septembre 2026"
DATA_PATH = "/home/z/my-project/scripts/perf_mensuelle_2026_09.json"
OUT_PDF = "/home/z/my-project/download/performance_mensuelle_septembre_2026.pdf"
CHARTS_DIR = "/home/z/my-project/work/charts"

# === Fonts ===
FONT_DIR = '/usr/share/fonts'
pdfmetrics.registerFont(TTFont('DejaVuSans', f'{FONT_DIR}/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', f'{FONT_DIR}/truetype/dejavu/DejaVuSans-Bold.ttf'))

# === Colors (BELGOCAM corporate) ===
NAVY = colors.HexColor('#1F3A5F')
GOLD = colors.HexColor('#C9A961')
RED = colors.HexColor('#C00000')
GREEN = colors.HexColor('#548235')
GRAY = colors.HexColor('#595959')
LIGHT_GRAY = colors.HexColor('#F2F2F2')
LIGHT_GREEN = colors.HexColor('#E2EFDA')
LIGHT_RED = colors.HexColor('#F8CBAD')
LIGHT_ORANGE = colors.HexColor('#FCE4D6')

styles = getSampleStyleSheet()
H1 = ParagraphStyle('H1', parent=styles['Heading1'], fontName='DejaVuSans-Bold', fontSize=18, textColor=NAVY, spaceAfter=14, spaceBefore=10, keepWithNext=1)
H2 = ParagraphStyle('H2', parent=styles['Heading2'], fontName='DejaVuSans-Bold', fontSize=13, textColor=NAVY, spaceAfter=10, spaceBefore=14)
H3 = ParagraphStyle('H3', parent=styles['Heading3'], fontName='DejaVuSans-Bold', fontSize=11, textColor=GOLD, spaceAfter=8, spaceBefore=10, keepWithNext=1)
BODY = ParagraphStyle('Body', parent=styles['Normal'], fontName='DejaVuSans', fontSize=10, leading=14, alignment=TA_JUSTIFY, spaceAfter=8)
BULLET = ParagraphStyle('Bullet', parent=BODY, leftIndent=20, bulletIndent=10, spaceAfter=4)
SMALL = ParagraphStyle('Small', parent=BODY, fontSize=8, textColor=GRAY)
CELL = ParagraphStyle('Cell', parent=BODY, fontName='DejaVuSans', fontSize=9, leading=11, alignment=TA_LEFT, spaceAfter=0)
TITLE = ParagraphStyle('Title2', parent=styles['Title'], fontName='DejaVuSans-Bold', fontSize=26, textColor=NAVY, alignment=TA_LEFT, spaceAfter=6)

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            if self._pageNumber > 1:
                self.draw_page_header()
                self.draw_page_footer(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_header(self):
        self.saveState()
        self.setFont('DejaVuSans-Bold', 8)
        self.setFillColor(NAVY)
        self.drawString(2*cm, A4[1]-1.0*cm, "BELGOCAM SA — NJS GROUP")
        self.setFont('DejaVuSans', 8)
        self.setFillColor(GRAY)
        self.drawRightString(A4[0]-2*cm, A4[1]-1.0*cm, f"Performance mensuelle — {MONTH_LABEL}")
        self.setStrokeColor(GOLD)
        self.setLineWidth(0.8)
        self.line(2*cm, A4[1]-1.25*cm, A4[0]-2*cm, A4[1]-1.25*cm)
        self.restoreState()

    def draw_page_footer(self, num_pages):
        self.saveState()
        self.setStrokeColor(colors.HexColor('#BFBFBF'))
        self.setLineWidth(0.5)
        self.line(2*cm, 1.5*cm, A4[0]-2*cm, 1.5*cm)
        self.setFont('DejaVuSans', 8)
        self.setFillColor(GRAY)
        self.drawString(2*cm, 1.0*cm, "BELGOCAM SA — Performance mensuelle — Confidentiel")
        self.drawRightString(A4[0]-2*cm, 1.0*cm, f"Page {self._pageNumber} / {num_pages}")
        self.restoreState()

# === Load data ===
DATA = json.load(open(DATA_PATH, encoding='utf-8'))
print(f"Loaded {DATA_PATH} ({DATA['meta']['label']})")

# Helpers
def fmt(x):
    if x is None:
        return "—"
    if isinstance(x, float):
        return f"{x:,.0f}".replace(',', ' ').replace('.', ',')
    return str(x)

def fmt1(x):
    return f"{x:,.1f}".replace(',', ' ').replace('.', ',')

def pct(x):
    if x is None:
        return "—"
    return f"{x:.1f}%".replace('.', ',')

def pct_cell(x):
    """% with status color background."""
    if x is None:
        return Paragraph("—", CELL)
    s = ParagraphStyle('PC', parent=CELL, alignment=TA_CENTER,
                       textColor=(GREEN if x >= 100 else RED),
                       fontName='DejaVuSans-Bold')
    return Paragraph(f"{x:.0f}%".replace('.', ','), s)

def make_table(data, col_widths=None, font_size=9, header_color=NAVY, highlight_rows=None):
    cell_style = ParagraphStyle('CD', parent=CELL, fontSize=font_size, leading=font_size+2)
    cell_header = ParagraphStyle('CH', parent=cell_style, fontName='DejaVuSans-Bold', textColor=colors.white, alignment=TA_CENTER)
    cell_center = ParagraphStyle('CC', parent=cell_style, alignment=TA_CENTER)
    processed = []
    for ri, row in enumerate(data):
        prow = []
        for ci, cell in enumerate(row):
            if isinstance(cell, Paragraph):
                prow.append(cell)
            else:
                s = str(cell) if cell is not None else ''
                if ri == 0:
                    prow.append(Paragraph(s, cell_header))
                elif len(s) <= 15:
                    prow.append(Paragraph(s, cell_center))
                else:
                    prow.append(Paragraph(s, cell_style))
        processed.append(prow)
    t = Table(processed, colWidths=col_widths, repeatRows=1)
    style_list = [
        ('BACKGROUND', (0, 0), (-1, 0), header_color),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#BFBFBF')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, LIGHT_GRAY]),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]
    if highlight_rows:
        for r in highlight_rows:
            style_list.append(('BACKGROUND', (0, r), (-1, r), LIGHT_ORANGE))
            style_list.append(('FONTNAME', (0, r), (-1, r), 'DejaVuSans-Bold'))
    t.setStyle(TableStyle(style_list))
    return t

def img(path, width=16.5*cm):
    return Image(path, width=width, height=width * 0.5)

# ============================================================
# Story
# ============================================================
story = []

# --- COVER ---
story.append(Spacer(1, 3.5*cm))
story.append(Paragraph("BELGOCAM SA", TITLE))
story.append(Paragraph("NJS GROUP — Direction Générale", ParagraphStyle('Sub', parent=BODY, fontSize=12, textColor=GRAY, spaceAfter=30)))
story.append(HRFlowable(width='100%', thickness=2, color=GOLD, spaceAfter=20))
story.append(Paragraph("Rapport de Performance Mensuelle", H1))
story.append(Paragraph(f"{MONTH_LABEL} — données complètes du mois (ERP, ventes livrées)", ParagraphStyle('S2', parent=BODY, fontSize=13, textColor=NAVY, spaceAfter=24)))

kpi = DATA['sc_ytd']
cover_kpis = [
    ["Indicateur", "Valeur"],
    ["Soja Septembre (réel vs obj)", f"{fmt(DATA['global_sept']['TOURTEAUX']['t'])} t vs {fmt(DATA['global_sept']['TOURTEAUX']['obj'])} t ({pct(DATA['global_sept']['TOURTEAUX']['pct'])})"],
    ["Concentrés Septembre (réel vs obj)", f"{fmt(DATA['global_sept']['CONCENTRES']['t'])} t vs {fmt(DATA['global_sept']['CONCENTRES']['obj'])} t ({pct(DATA['global_sept']['CONCENTRES']['pct'])})"],
    ["Soja YTD Jan-Sep (réel vs obj)", f"{fmt(kpi['soja']['t'])} t vs {fmt(kpi['soja']['obj'])} t ({pct(kpi['soja']['pct'])})"],
    ["Concentrés YTD Jan-Sep (réel vs obj)", f"{fmt(kpi['conc']['t'])} t vs {fmt(kpi['conc']['obj'])} t ({pct(kpi['conc']['pct'])})"],
    ["Ratio soja:concentrés YTD", f"{DATA['ratio_soja_conc_ytd']} (cible bundle ≤ 2,5)"],
]
story.append(make_table(cover_kpis, col_widths=[7.5*cm, 9*cm], font_size=10, highlight_rows=[1, 2]))
story.append(Spacer(1, 1.2*cm))
story.append(Paragraph(
    "Ce rapport présente la performance commerciale du mois comparée aux objectifs recalibrés, "
    "ainsi que la tendance cumulée depuis janvier (YTD), par famille, par agence et par région. "
    "Il est mis à jour chaque mois à partir de l'extraction ERP des ventes livrées.",
    BODY))
story.append(Spacer(1, 0.4*cm))
story.append(Paragraph(f"Données au {DATA['meta']['update_date']} — Source : {DATA['meta']['source_erp']}", SMALL))
story.append(PageBreak())

# --- SECTION 1 : Performance globale par famille ---
story.append(Paragraph("1. Performance globale par famille — Septembre 2026", H1))
story.append(Paragraph(
    "Lecture des volumes livrés en septembre comparés aux objectifs S2 recalibrés du mois "
    "(objectifs Takou pour les mois 1-6, recalibrage S2 pour les mois 7-12). "
    f"Le mois de septembre a été marqué par la baisse tarifaire du soja au 22/09/2026 "
    f"(cadence passée de 139 t/j à 278 t/j, +100%), qui a fortement tiré les volumes de fin de mois.",
    BODY))

g = DATA['global_sept']
rows = [["Famille", "Réel Sept (t)", "Objectif (t)", "% Objectif", "Écart (t)", "CA Sept (M FCFA)"]]
for fam in ['TOURTEAUX', 'CONCENTRES', 'ALIMENT_COMPLET', 'INGREDIENTS', 'PREMIX', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES', 'MATERIEL_ELEVAGE']:
    d = g[fam]
    statut = "✓" if (d['pct'] or 0) >= 100 else "✗"
    rows.append([fam, fmt(d['t']), fmt(d['obj']), pct(d['pct']),
                 f"+{fmt(d['ecart'])}" if d['ecart'] >= 0 else fmt(d['ecart']).replace('-', '- '), fmt(d['ca'])])
story.append(make_table(rows, col_widths=[3.6*cm, 2.4*cm, 2.3*cm, 2.0*cm, 2.2*cm, 3.0*cm], font_size=8.5, highlight_rows=[1, 2]))
story.append(Spacer(1, 0.35*cm))
story.append(img(f'{CHARTS_DIR}/perf_01_global_mois.png'))
story.append(Spacer(1, 0.3*cm))
story.append(Paragraph(
    f"<b>Points clés</b> : Soja à <b>{pct(g['TOURTEAUX']['pct'])}</b> de l'objectif ({fmt(g['TOURTEAUX']['t'])} t vs {fmt(g['TOURTEAUX']['obj'])} t) "
    f"et Concentrés à <b>{pct(g['CONCENTRES']['pct'])}</b> ({fmt(g['CONCENTRES']['t'])} t vs {fmt(g['CONCENTRES']['obj'])} t). "
    f"Le ratio soja:concentrés du mois est de "
    f"<b>{g['TOURTEAUX']['t']/g['CONCENTRES']['t']:.1f}:1</b> (cible bundle ≤ 2,5:1).",
    BODY))
story.append(PageBreak())

# --- SECTION 2 : Performance par agence × famille ---
story.append(Paragraph("2. Performance par agence × famille — Septembre 2026", H1))
story.append(Paragraph(
    "Soja et Concentrés par agence (les deux familles cœur), puis % d'objectif des autres familles par agence. "
    "Objectifs agence = déclinaison S2 recalibrée du mois (mois 9).", BODY))

a = DATA['agence_sept']
rows = [["Agence", "Région", "Soja réel", "Soja obj", "% Soja", "Conc réel", "Conc obj", "% Conc"]]
for ag in ['Famla', 'Djeleng', 'Mbouda', 'Messassi', 'Bertoua', 'Ngaoundere', 'Ahala', 'Nkolbisson', 'Nkoabang',
           'Ndobo', 'Village', 'Pk11', 'Nkongsamba', 'Buea']:
    reg = {'Famla': 'Ouest', 'Djeleng': 'Ouest', 'Mbouda': 'Ouest',
           'Messassi': 'Centre', 'Bertoua': 'Centre', 'Ngaoundere': 'Centre', 'Ahala': 'Centre',
           'Nkolbisson': 'Centre', 'Nkoabang': 'Centre',
           'Ndobo': 'Littoral', 'Village': 'Littoral', 'Pk11': 'Littoral', 'Nkongsamba': 'Littoral', 'Buea': 'Littoral'}[ag]
    s = a[ag]['TOURTEAUX']
    c = a[ag]['CONCENTRES']
    rows.append([ag, reg, fmt(s['t']), fmt(s['obj']), pct_cell(s['pct']),
                 fmt(c['t']), fmt(c['obj']), pct_cell(c['pct'])])
story.append(make_table(rows, col_widths=[2.5*cm, 1.7*cm, 1.9*cm, 1.9*cm, 1.5*cm, 1.9*cm, 1.9*cm, 1.5*cm], font_size=8))

story.append(Spacer(1, 0.35*cm))
rows2 = [["Agence", "ALIM. COMPLET", "INGRÉDIENTS", "PREMIX", "COMPL. ALIM."]]
for ag in ['Famla', 'Djeleng', 'Mbouda', 'Messassi', 'Bertoua', 'Ngaoundere', 'Ahala', 'Nkolbisson', 'Nkoabang',
           'Ndobo', 'Village', 'Pk11', 'Nkongsamba', 'Buea']:
    rows2.append([ag,
                  pct_cell(a[ag]['ALIMENT_COMPLET']['pct']),
                  pct_cell(a[ag]['INGREDIENTS']['pct']),
                  pct_cell(a[ag]['PREMIX']['pct']),
                  pct_cell(a[ag]['COMPLEMENT_ALIMENTAIRE']['pct'])])
story.append(make_table(rows2, col_widths=[2.5*cm, 3.0*cm, 3.0*cm, 3.0*cm, 3.0*cm], font_size=8))

story.append(Spacer(1, 0.35*cm))
story.append(img(f'{CHARTS_DIR}/perf_02_agence_mois.png', width=15*cm))
story.append(Spacer(1, 0.3*cm))
top_conc = max(a, key=lambda ag: a[ag]['CONCENTRES']['t'])
top_soja = max(a, key=lambda ag: a[ag]['TOURTEAUX']['t'])
story.append(Paragraph(
    f"<b>Points clés</b> : {top_soja} en tête du soja ({fmt(a[top_soja]['TOURTEAUX']['t'])} t), "
    f"{top_conc} en tête des concentrés ({fmt(a[top_conc]['CONCENTRES']['t'])} t). "
    f"Les agences Littoral (Ndobo, Village, Pk11, Nkongsamba, Buea) restent le point de vigilance sur les deux familles.",
    BODY))
story.append(PageBreak())

# --- SECTION 3 : Performance par région × famille ---
story.append(Paragraph("3. Performance par région × famille — Septembre 2026", H1))
story.append(Paragraph(
    "Agrégation régionale des 14 agences, comparée aux objectifs S2 recalibrés du mois.", BODY))

r = DATA['region_sept']
rows = [["Région", "Soja réel", "Soja obj", "% Soja", "Conc réel", "Conc obj", "% Conc", "Ratio soja:conc"]]
for reg in ['Ouest', 'Centre', 'Littoral']:
    s = r[reg]['TOURTEAUX']
    c = r[reg]['CONCENTRES']
    ratio = s['t'] / c['t'] if c['t'] > 0 else None
    rows.append([reg, fmt(s['t']), fmt(s['obj']), pct_cell(s['pct']),
                 fmt(c['t']), fmt(c['obj']), pct_cell(c['pct']),
                 f"{ratio:.1f}" if ratio else "—"])
story.append(make_table(rows, col_widths=[2.2*cm, 1.9*cm, 1.9*cm, 1.5*cm, 1.9*cm, 1.9*cm, 1.5*cm, 2.0*cm], font_size=8.5))

story.append(Spacer(1, 0.3*cm))
rows2 = [["Région", "ALIM. COMPLET", "INGRÉDIENTS", "PREMIX", "COMPL. ALIM."]]
for reg in ['Ouest', 'Centre', 'Littoral']:
    rows2.append([reg,
                  pct_cell(r[reg]['ALIMENT_COMPLET']['pct']),
                  pct_cell(r[reg]['INGREDIENTS']['pct']),
                  pct_cell(r[reg]['PREMIX']['pct']),
                  pct_cell(r[reg]['COMPLEMENT_ALIMENTAIRE']['pct'])])
story.append(make_table(rows2, col_widths=[2.2*cm, 3.2*cm, 3.2*cm, 3.2*cm, 3.2*cm], font_size=8.5))

story.append(Spacer(1, 0.35*cm))
story.append(img(f'{CHARTS_DIR}/perf_03_region_mois.png'))
story.append(Spacer(1, 0.3*cm))
story.append(Paragraph(
    f"<b>Points clés</b> : Ouest à <b>{pct(r['Ouest']['TOURTEAUX']['pct'])}</b> de l'objectif soja, "
    f"Centre à <b>{pct(r['Centre']['TOURTEAUX']['pct'])}</b>, Littoral à <b>{pct(r['Littoral']['TOURTEAUX']['pct'])}</b>. "
    f"Sur les concentrés : Ouest {pct(r['Ouest']['CONCENTRES']['pct'])}, Centre {pct(r['Centre']['CONCENTRES']['pct'])} "
    f"et Littoral {pct(r['Littoral']['CONCENTRES']['pct'])} — le Littoral est la région en décrochage sur le mois.",
    BODY))
story.append(PageBreak())

# --- SECTION 4 : Tendance mensuelle par famille (YTD) ---
story.append(Paragraph("4. Tendance mensuelle par famille vs objectifs — Jan → Sept (YTD)", H1))
story.append(Paragraph(
    "Volumes mensuels réalisés vs objectif mensuel (Takou mois 1-6, S2 recalibré mois 7-9) "
    "pour les deux familles cœur, puis % d'objectif mensuel pour les autres familles.", BODY))

fm = DATA['fam_monthly']
MONTHS = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep']
rows = [["Mois", "Soja réel", "Soja obj", "% Soja", "Conc réel", "Conc obj", "% Conc"]]
for m in range(1, 10):
    s = fm['TOURTEAUX'][str(m)]
    c = fm['CONCENTRES'][str(m)]
    rows.append([MONTHS[m-1], fmt(s['t']), fmt(s['obj']), pct_cell(s['t']/s['obj']*100 if s['obj'] else None),
                 fmt(c['t']), fmt(c['obj']), pct_cell(c['t']/c['obj']*100 if c['obj'] else None)])
ytd_s = DATA['sc_ytd']['soja']
ytd_c = DATA['sc_ytd']['conc']
rows.append(["YTD", fmt(ytd_s['t']), fmt(ytd_s['obj']), pct_cell(ytd_s['pct']),
             fmt(ytd_c['t']), fmt(ytd_c['obj']), pct_cell(ytd_c['pct'])])
story.append(make_table(rows, col_widths=[1.7*cm, 2.0*cm, 2.0*cm, 1.6*cm, 2.0*cm, 2.0*cm, 1.6*cm], font_size=8.5, highlight_rows=[10]))

story.append(Spacer(1, 0.3*cm))
rows2 = [["Famille"] + MONTHS + ["YTD"]]
for fam in ['ALIMENT_COMPLET', 'INGREDIENTS', 'PREMIX', 'COMPLEMENT_ALIMENTAIRE']:
    row = [fam]
    for m in range(1, 10):
        d = fm[fam][str(m)]
        row.append(pct_cell(d['t']/d['obj']*100 if d['obj'] else None))
    d_ytd = DATA['fam_ytd'][fam]
    row.append(pct_cell(d_ytd['pct']))
    rows2.append(row)
story.append(make_table(rows2, col_widths=[2.8*cm] + [1.2*cm]*9 + [1.2*cm], font_size=7.5))

story.append(Spacer(1, 0.35*cm))
story.append(img(f'{CHARTS_DIR}/perf_04_fam_mensuel.png'))
story.append(Spacer(1, 0.3*cm))
story.append(Paragraph(
    f"<b>Points clés</b> : Le soja dépasse l'objectif cumulé (<b>{pct(ytd_s['pct'])}</b>, {fmt(ytd_s['t'])} t vs {fmt(ytd_s['obj'])} t) "
    f"grâce à juillet-août (rupture concurrente) et à la baisse tarifaire du 22/09. "
    f"Les concentrés sont en retard cumulé (<b>{pct(ytd_c['pct'])}</b>, {fmt(ytd_c['t'])} t vs {fmt(ytd_c['obj'])} t) : "
    f"le ratio soja:concentrés YTD de <b>{DATA['ratio_soja_conc_ytd']}</b> reste au-dessus de la cible bundle de 2,5.",
    BODY))
story.append(PageBreak())

# --- SECTION 5 : Tendance mensuelle par agence × Soja/Conc (YTD) ---
story.append(Paragraph("5. Tendance mensuelle par agence × (Soja & Concentrés) — YTD", H1))
story.append(Paragraph(
    "Heatmaps du % d'objectif mensuel par agence (100 = objectif atteint), puis cumul YTD par agence "
    "comparé à l'objectif cumulé (Takou mois 1-6 + S2 mois 7-9).", BODY))

story.append(img(f'{CHARTS_DIR}/perf_05_heatmap_agence.png'))

story.append(Spacer(1, 0.35*cm))
ay = DATA['agence_sc_ytd']
rows = [["Agence", "Région", "Soja YTD", "Soja obj", "% Soja", "Conc YTD", "Conc obj", "% Conc"]]
for ag in ['Famla', 'Djeleng', 'Mbouda', 'Messassi', 'Bertoua', 'Ngaoundere', 'Ahala', 'Nkolbisson', 'Nkoabang',
           'Ndobo', 'Village', 'Pk11', 'Nkongsamba', 'Buea']:
    reg = {'Famla': 'Ouest', 'Djeleng': 'Ouest', 'Mbouda': 'Ouest',
           'Messassi': 'Centre', 'Bertoua': 'Centre', 'Ngaoundere': 'Centre', 'Ahala': 'Centre',
           'Nkolbisson': 'Centre', 'Nkoabang': 'Centre',
           'Ndobo': 'Littoral', 'Village': 'Littoral', 'Pk11': 'Littoral', 'Nkongsamba': 'Littoral', 'Buea': 'Littoral'}[ag]
    s = ay[ag]['soja']
    c = ay[ag]['conc']
    rows.append([ag, reg, fmt(s['t']), fmt(s['obj']), pct_cell(s['pct']),
                 fmt(c['t']), fmt(c['obj']), pct_cell(c['pct'])])
story.append(make_table(rows, col_widths=[2.5*cm, 1.7*cm, 1.9*cm, 1.9*cm, 1.5*cm, 1.9*cm, 1.9*cm, 1.5*cm], font_size=8))

story.append(Spacer(1, 0.3*cm))
story.append(Paragraph(
    f"<b>Points clés</b> : la quasi-totalité des agences dépasse l'objectif soja YTD, mais la plupart "
    f"restent sous l'objectif concentrés YTD — le retard concentrés est généralisé et pas seulement littoral. "
    f"Le bundle (2,5:1) doit redevenir le levier prioritaire du T4.",
    BODY))
story.append(PageBreak())

# --- SECTION 6 : Tendance mensuelle par région × Soja/Conc (YTD) ---
story.append(Paragraph("6. Tendance mensuelle par région × (Soja & Concentrés) — YTD", H1))
story.append(Paragraph(
    "Tendance mensuelle régionale comparée à l'objectif national mensuel, puis cumul YTD par région.", BODY))

story.append(img(f'{CHARTS_DIR}/perf_06_region_mensuel.png'))

story.append(Spacer(1, 0.35*cm))
ry = DATA['region_sc_ytd']
rows = [["Région", "Soja YTD", "Soja obj", "% Soja", "Conc YTD", "Conc obj", "% Conc", "Ratio YTD"]]
for reg in ['Ouest', 'Centre', 'Littoral']:
    s = ry[reg]['soja']
    c = ry[reg]['conc']
    ratio = s['t'] / c['t'] if c['t'] > 0 else None
    rows.append([reg, fmt(s['t']), fmt(s['obj']), pct_cell(s['pct']),
                 fmt(c['t']), fmt(c['obj']), pct_cell(c['pct']),
                 f"{ratio:.2f}" if ratio else "—"])
rows.append(["NATIONAL", fmt(ytd_s['t']), fmt(ytd_s['obj']), pct_cell(ytd_s['pct']),
             fmt(ytd_c['t']), fmt(ytd_c['obj']), pct_cell(ytd_c['pct']), f"{DATA['ratio_soja_conc_ytd']}"])
story.append(make_table(rows, col_widths=[2.2*cm, 1.9*cm, 1.9*cm, 1.5*cm, 1.9*cm, 1.9*cm, 1.5*cm, 1.8*cm], font_size=8.5, highlight_rows=[4]))

story.append(Spacer(1, 0.3*cm))
story.append(Paragraph(
    f"<b>Points clés</b> : le soja YTD est porté par l'Ouest ({fmt(ry['Ouest']['soja']['t'])} t) et le Centre "
    f"({fmt(ry['Centre']['soja']['t'])} t) ; le Littoral est à {pct(ry['Littoral']['soja']['pct'])} de son objectif soja. "
    f"Sur les concentrés, les trois régions sont sous objectif YTD — Littoral {pct(ry['Littoral']['conc']['pct'])}, "
    f"Centre {pct(ry['Centre']['conc']['pct'])}, Ouest {pct(ry['Ouest']['conc']['pct'])}.",
    BODY))

# --- Recommandations ---
story.append(Spacer(1, 0.4*cm))
story.append(Paragraph("Recommandations pour le mois suivant", H2))
recos = [
    "<b>Bundle soja:concentrés</b> — Le ratio YTD de 3,2:1 dépasse la cible de 2,5:1 : conditionner toute vente de soja à la vente de concentrés, priorité sur le Littoral.",
    "<b>Concentrés</b> — Retard YTD de 2 895 t vs objectif : plan de rattrapage T4 avec objectifs hebdomadaires par agence.",
    "<b>Littoral</b> — Région en décrochage sur les deux familles : revue des comptes clés (Ndobo, Village, Pk11, Nkongsamba, Buea) et plan de relance.",
    "<b>Prix soja</b> — Suivre l'effet de la baisse tarifaire du 22/09 (cadence x2) et sécuriser le réapprovisionnement avant la rupture prévue début octobre.",
    "<b>Rituel mensuel</b> — Ce rapport sera mis à jour chaque mois avec l'extraction ERP du mois clos.",
]
for rec in recos:
    story.append(Paragraph(f"• {rec}", BULLET))

# --- Méthodologie ---
story.append(Spacer(1, 0.4*cm))
story.append(Paragraph("Méthodologie & sources", H2))
methodo = [
    f"<b>Périmètre</b> : ventes livrées (état « Livrée ») des 14 agences BELGOCAM, toutes familles, extraction ERP du {DATA['meta']['update_date']}.",
    f"<b>Actuals</b> : janvier-août 2026 issus du dataset historique consolidé ; {MONTH_LABEL} issu de l'extraction {DATA['meta']['source_erp']}.",
    "<b>Objectifs</b> : mois 1-6 = objectifs Takou ; mois 7-12 = objectifs S2 recalibrés (14 agences, cohérents avec le national).",
    "<b>Unités</b> : volumes en tonnes ; CA en M FCFA (TTC) ; % d'objectif = réel / objectif × 100.",
    "<b>Familles</b> : TOURTEAUX (soja), CONCENTRÉS, ALIMENT COMPLET, INGRÉDIENTS, PREMIX, COMPLÉMENT ALIMENTAIRE, ALVÉOLES, MATÉRIEL D'ÉLEVAGE.",
]
for m in methodo:
    story.append(Paragraph(f"• {m}", BULLET))

# === Build ===
doc = SimpleDocTemplate(OUT_PDF, pagesize=A4,
                        leftMargin=2*cm, rightMargin=2*cm, topMargin=2.2*cm, bottomMargin=2*cm,
                        title=f"Performance mensuelle — {MONTH_LABEL}",
                        author="BELGOCAM SA — NJS GROUP")
doc.build(story, canvasmaker=NumberedCanvas)
print(f"PDF généré: {OUT_PDF} ({os.path.getsize(OUT_PDF)//1024} KB)")
