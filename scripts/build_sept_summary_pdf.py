"""Generate September update PDF — combined analysis: zero-achat + bundle soja-concentrés + performance CONCENTRES.
Does NOT overwrite the original juillet/août PDFs. Saves as analyse_septembre_upd.pdf.

Based on sept_mtd_10.json (computed by compute_sept_mtd_metrics.py — mois complet 01-30/09/2026).
"""
import os
import json
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, HRFlowable, KeepTogether
)

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
LIGHT_ORANGE = colors.HexColor('#FCE4D6')

styles = getSampleStyleSheet()
H1 = ParagraphStyle('H1', parent=styles['Heading1'], fontName='DejaVuSans-Bold', fontSize=18, textColor=NAVY, spaceAfter=14, spaceBefore=10, keepWithNext=1)
H2 = ParagraphStyle('H2', parent=styles['Heading2'], fontName='DejaVuSans-Bold', fontSize=14, textColor=NAVY, spaceAfter=10, spaceBefore=14)
H3 = ParagraphStyle('H3', parent=styles['Heading3'], fontName='DejaVuSans-Bold', fontSize=12, textColor=GOLD, spaceAfter=8, spaceBefore=10, keepWithNext=1)
BODY = ParagraphStyle('Body', parent=styles['Normal'], fontName='DejaVuSans', fontSize=10, leading=14, alignment=TA_JUSTIFY, spaceAfter=8)
BULLET = ParagraphStyle('Bullet', parent=BODY, leftIndent=20, bulletIndent=10, spaceAfter=4)
SMALL = ParagraphStyle('Small', parent=BODY, fontSize=8, textColor=GRAY)
CELL = ParagraphStyle('Cell', parent=BODY, fontName='DejaVuSans', fontSize=9, leading=11, alignment=TA_LEFT, spaceAfter=0)

# Footer with page numbers ("Page X / Y") — skipped on the cover (page 1)
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
        self.drawRightString(A4[0]-2*cm, A4[1]-1.0*cm, f"Analyse Septembre 2026 — Données au {DATA['update_date']}")
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
        self.drawString(2*cm, 1.0*cm, "BELGOCAM SA — Analyse Septembre 2026 — Confidentiel")
        self.drawRightString(A4[0]-2*cm, 1.0*cm, f"Page {self._pageNumber} / {num_pages}")
        self.restoreState()

# === Load data ===
DATA = json.load(open('/home/z/my-project/scripts/sept_mtd_10.json'))
print(f"Loaded sept_mtd_10.json (update {DATA['update_date']})")

# Helpers
def fmt(x):
    """Format number with French thousands separator (space)."""
    if isinstance(x, float):
        return f"{x:,.0f}".replace(',', ' ').replace('.', ',')
    return str(x)

def fmt1(x):
    return f"{x:,.1f}".replace(',', ' ').replace('.', ',')

def make_table(data, col_widths=None, font_size=9, header_color=NAVY, highlight_rows=None):
    cell_style = ParagraphStyle('CD', parent=CELL, fontSize=font_size, leading=font_size+2)
    cell_header = ParagraphStyle('CH', parent=cell_style, fontName='DejaVuSans-Bold', textColor=colors.white, alignment=TA_CENTER)
    cell_center = ParagraphStyle('CC', parent=cell_style, alignment=TA_CENTER)
    processed = []
    for ri, row in enumerate(data):
        prow = []
        for ci, cell in enumerate(row):
            s = str(cell) if cell is not None else ''
            if ri == 0: prow.append(Paragraph(s, cell_header))
            elif len(s) <= 15: prow.append(Paragraph(s, cell_center))
            else: prow.append(Paragraph(s, cell_style))
        processed.append(prow)
    t = Table(processed, colWidths=col_widths, repeatRows=1)
    style_list = [
        ('BACKGROUND', (0,0), (-1,0), header_color),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_GRAY]),
        ('TOPPADDING', (0,0), (-1,-1), 4), ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4), ('RIGHTPADDING', (0,0), (-1,-1), 4),
    ]
    if highlight_rows:
        for ridx in highlight_rows:
            style_list.append(('BACKGROUND', (0, ridx), (-1, ridx), colors.HexColor('#FFE699')))
    t.setStyle(TableStyle(style_list))
    return t

def heading_block(text):
    """Section title + separator — returned as a list so each call site can keep it
    together with the section's first block (a title never ends a page alone)."""
    return [
        Paragraph(text, H1),
        HRFlowable(width="100%", thickness=1, color=NAVY),
        Spacer(1, 0.3*cm),
    ]

# === Build PDF ===
OUT = '/home/z/my-project/download/analyse_septembre_upd.pdf'
doc = SimpleDocTemplate(OUT, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []

# === COVER ===
story.append(Spacer(1, 2*cm))
story.append(Paragraph("BELGOCAM SA", ParagraphStyle('CL', parent=styles['Title'], fontName='DejaVuSans-Bold', fontSize=32, textColor=NAVY, alignment=TA_CENTER)))
story.append(Spacer(1, 0.5*cm))
story.append(HRFlowable(width="60%", thickness=2, color=GOLD, spaceBefore=10, spaceAfter=20, hAlign='CENTER'))
story.append(Spacer(1, 1*cm))
story.append(Paragraph("Analyse Septembre 2026 — Mise à jour", ParagraphStyle('CT', parent=styles['Title'], fontName='DejaVuSans-Bold', fontSize=24, textColor=NAVY, alignment=TA_CENTER, spaceAfter=20)))
story.append(Paragraph("Performance CONCENTRES + Bundle Soja-Concentrés + Zero-Achat", ParagraphStyle('CS', parent=styles['Title'], fontName='DejaVuSans', fontSize=14, textColor=GOLD, alignment=TA_CENTER, spaceAfter=30)))
story.append(Spacer(1, 2*cm))
story.append(HRFlowable(width="40%", thickness=1, color=GRAY, spaceBefore=10, spaceAfter=10, hAlign='CENTER'))
story.append(Paragraph(f"<b>Mise à jour</b> : {DATA['update_date']} — <b>mois complet</b> ({DATA['days_elapsed']}/{DATA['total_days_sep']} jours ouvrés, données définitives)", ParagraphStyle('CI', parent=BODY, fontSize=11, alignment=TA_CENTER, textColor=GRAY)))
story.append(Paragraph(f"<b>⚠ Événement du mois — Baisse prix soja le 22/09/2026 (mi-journée)</b> : 26 000 → <b>20 000 FCFA/sac</b> (-23%) → doublement de la cadence soja (139 → 278 t/j)", ParagraphStyle('CI4', parent=BODY, fontName='DejaVuSans-Bold', fontSize=11, alignment=TA_CENTER, textColor=RED)))
story.append(Paragraph(f"<b>Source</b> : {DATA['extraction_file']}", ParagraphStyle('CI2', parent=BODY, fontSize=10, alignment=TA_CENTER, textColor=GRAY)))
story.append(Paragraph("William Francis Fohom — Data Analyst, Administrateur National de Ventes", ParagraphStyle('CI3', parent=BODY, fontSize=10, alignment=TA_CENTER, textColor=GRAY, spaceBefore=10)))

story.append(PageBreak())

# === Section 1: Performance CONCENTRES septembre ===
# (titre + 1er tableau soudés pour éviter un titre orphelin en bas de page)

v_conc = DATA['volumes']['CONCENTRES']
v_soja = DATA['volumes']['TOURTEAUX']

# KPI box (mois complet : le réel = résultat définitif)
ecart_conc = v_conc['t'] - v_conc['obj_t']
ecart_soja = v_soja['t'] - v_soja['obj_t']
kpi_data = [
    ["Indicateur", "Réel Septembre", "Objectif", "% Objectif", "Écart", "Statut"],
    ["Volume CONCENTRES (t)", fmt(v_conc['t']), fmt(v_conc['obj_t']), fmt(v_conc['pct_obj']) + "%",
     f"{'+' if ecart_conc >= 0 else ''}{fmt(ecart_conc)} t",
     "✅ Objectif dépassé" if v_conc['pct_obj'] >= 100 else ("⚠ Sous objectif" if v_conc['pct_obj'] >= 80 else "❌ Critique")],
    ["Moyenne quotidienne (t/j)", fmt1(v_conc['moy_t_j']), fmt1(v_conc['obj_t']/DATA['total_days_sep']), "—",
     f"{'+' if v_conc['moy_t_j'] >= v_conc['obj_t']/DATA['total_days_sep'] else ''}{fmt1(v_conc['moy_t_j'] - v_conc['obj_t']/DATA['total_days_sep'])} t/j",
     "✅ Au-dessus objectif" if v_conc['moy_t_j'] >= v_conc['obj_t']/DATA['total_days_sep'] else "⚠ Sous objectif"],
    ["Volume TOURTEAUX (t)", fmt(v_soja['t']), fmt(v_soja['obj_t']), fmt(v_soja['pct_obj']) + "%",
     f"{'+' if ecart_soja >= 0 else ''}{fmt(ecart_soja)} t",
     "✅ Objectif dépassé" if v_soja['pct_obj'] >= 100 else ("⚠ Sous objectif" if v_soja['pct_obj'] >= 80 else "❌ Critique")],
]
kpi_table = make_table(kpi_data, col_widths=[4*cm, 2.5*cm, 2*cm, 2*cm, 2.5*cm, 4*cm], font_size=9, highlight_rows=[])
story.append(KeepTogether(heading_block("1. Performance CONCENTRES — Septembre 2026 (mois complet)")
                          + [Paragraph("<b>KPIs clés — CONCENTRES Septembre (définitif)</b>", H3), kpi_table]))
story.append(Spacer(1, 0.3*cm))

# Comparison avec août
import os
aout_path = '/home/z/my-project/scripts/aout_mtd_30.json'
if os.path.exists(aout_path):
    AOUT = json.load(open(aout_path))
    a_conc = AOUT['volumes']['CONCENTRES']
    a_soja = AOUT['volumes']['TOURTEAUX']

    var_conc_t = ((v_conc['t'] / a_conc['t']) - 1) * 100 if a_conc['t'] > 0 else 0
    var_soja_t = ((v_soja['t'] / a_soja['t']) - 1) * 100 if a_soja['t'] > 0 else 0

    var_conc_moy = ((v_conc['moy_t_j'] / a_conc['moy_t_j']) - 1) * 100 if a_conc['moy_t_j'] > 0 else 0
    var_soja_moy = ((v_soja['moy_t_j'] / a_soja['moy_t_j']) - 1) * 100 if a_soja['moy_t_j'] > 0 else 0

    comp_data = [
        ["Indicateur", "Août complet", "Septembre complet", "Variation", "Lecture"],
        ["CONCENTRES (t)", fmt(a_conc['t']), fmt(v_conc['t']), f"{'+' if var_conc_t>=0 else ''}{var_conc_t:.1f}%",
         "Progression" if var_conc_t >= -5 else "Recul significatif"],
        ["CONCENTRES moy/j (t/j)", fmt1(a_conc['moy_t_j']), fmt1(v_conc['moy_t_j']), f"{'+' if var_conc_moy>=0 else ''}{var_conc_moy:.1f}%",
         "Cadence en hausse" if var_conc_moy >= -5 else "Ralentissement"],
        ["TOURTEAUX (t)", fmt(a_soja['t']), fmt(v_soja['t']), f"{'+' if var_soja_t>=0 else ''}{var_soja_t:.1f}%",
         "Stable" if abs(var_soja_t) < 10 else ("Hausse" if var_soja_t > 0 else "Baisse")],
        ["TOURTEAUX moy/j (t/j)", fmt1(a_soja['moy_t_j']), fmt1(v_soja['moy_t_j']), f"{'+' if var_soja_moy>=0 else ''}{var_soja_moy:.1f}%",
         "Stable" if abs(var_soja_moy) < 10 else ("Hausse" if var_soja_moy > 0 else "Baisse")],
    ]
    comp_table = make_table(comp_data, col_widths=[4*cm, 2.5*cm, 3*cm, 2.5*cm, 4.5*cm], font_size=9, highlight_rows=[])
    story.append(KeepTogether([Paragraph("<b>Comparaison Septembre complet vs Août complet</b>", H3), comp_table]))
    story.append(Spacer(1, 0.3*cm))

# CONCENTRES par agence (mois complet)
ag_data = [["Rang", "Agence", "Conc Sept (t)", "Soja Sept (t)", "Moy/j (t)"]]
for i, ag in enumerate(DATA['conc_by_agence'], 1):
    ag_data.append([str(i), ag['agence'], fmt(ag['conc_t']), fmt(ag['soja_t']), fmt1(ag['moy_t_j'])])
# Total row
total_conc_mtd = sum(a['conc_t'] for a in DATA['conc_by_agence'])
total_soja_mtd = sum(a['soja_t'] for a in DATA['conc_by_agence'])
ag_data.append(["", f"TOTAL ({len(DATA['conc_by_agence'])} agences)", fmt(total_conc_mtd), fmt(total_soja_mtd),
               fmt1(total_conc_mtd/DATA['days_elapsed'])])
ag_table = make_table(ag_data, col_widths=[1.2*cm, 3.5*cm, 2.8*cm, 2.8*cm, 2.2*cm], font_size=9, highlight_rows=[len(ag_data)-1])
story.append(KeepTogether([Paragraph("<b>CONCENTRES par agence — Septembre complet</b>", H3), ag_table]))
story.append(Spacer(1, 0.3*cm))

# Analyse & lecture
story.append(Paragraph("<b>Analyse & Lecture</b>", H3))
analysis_text = f"""
<b>Performance CONCENTRES Septembre 2026 (mois complet)</b> :
{v_conc['t']} t livrées en {DATA['days_elapsed']} jours ouvrés, soit une cadence de
<b>{v_conc['moy_t_j']} t/jour</b> contre un objectif de {v_conc['obj_t']} t.
Le mois se clôture à <b>{v_conc['pct_obj']}% de l'objectif</b> (écart {ecart_conc:+.1f} t).
"""
if v_conc['pct_obj'] >= 100:
    analysis_text += "✅ <b>Objectif dépassé</b> — mois concluant pour les concentrés. "
elif v_conc['pct_obj'] >= 80:
    analysis_text += "✅ <b>Bonne performance</b> — objectif quasi atteint. "
elif v_conc['pct_obj'] >= 60:
    analysis_text += "⚠ <b>À surveiller</b> — sous-performance à corriger en octobre. "
else:
    analysis_text += "❌ <b>Critique</b> — action commerciale urgente requise. "

if os.path.exists(aout_path):
    a_conc = AOUT['volumes']['CONCENTRES']
    var = ((v_conc['t']/a_conc['t'])-1)*100
    if var >= 0:
        analysis_text += f"Le mois de septembre ({v_conc['t']} t) est <b>supérieur de {var:.1f}%</b> au volume d'août ({a_conc['t']} t) — <b>troisième mois consécutif de progression</b> de la dynamique concentrés."
    else:
        analysis_text += f"Le mois de septembre ({v_conc['t']} t) est <b>inférieur de {abs(var):.1f}%</b> au volume d'août ({a_conc['t']} t) — <b>ralentissement à surveiller</b>."

story.append(Paragraph(analysis_text, BODY))
story.append(Spacer(1, 0.2*cm))

# Encadré événements prix soja (mois complet)
story.append(Paragraph(
    "<b>⚠ Contexte prix du mois</b> : La hausse du soja T102 (50kg) au 04/09/2026 (25 000 → 26 000 FCFA/sac, +4%) avait freiné les "
    "ventes soja en début de mois (comportement attentiste) : <b>139,1 t/j en moyenne du 01 au 22/09</b>. "
    "<b>Le 22/09/2026 en milieu de journée, le prix est retombé à 20 000 FCFA/sac (-23%)</b>, déclenchant un puissant effet de "
    "rattrapage : <b>278,1 t/j du 23 au 30/09 (+100%)</b>, avec un pic à 471 t le 23/09. Sur le mois, la baisse de prix a ainsi "
    "contribué à la fois à un volume record (4 590 t de soja) et à une réduction du prix moyen de vente — arbitrage "
    "prix/volume à intégrer dans la politique tarifaire d'octobre. Le cross-sell CONCENTRES est resté préservé "
    f"({DATA['bundle']['pct_bundle']}% des commandes soja incluent du concentré) — la discipline bundle s'est maintenue malgré la volatilité prix.",
    ParagraphStyle('Box', parent=BODY, fontSize=9, leading=12, backColor=LIGHT_ORANGE, borderColor=RED, borderWidth=0.5, borderPadding=6, alignment=TA_JUSTIFY, spaceAfter=6)
))
story.append(Spacer(1, 0.3*cm))

story.append(PageBreak())

# === Section 2: Bundle Soja-Concentrés ===

b = DATA['bundle']
bundle_data = [
    ["Indicateur", "Valeur Septembre", "Comparaison"],
    ["Sacs soja vendus", fmt(b['total_soja_sacs']), f"≈ {b['total_soja_sacs']*50/1000:.0f} t équivalent"],
    ["Sacs concentrés vendus", fmt(b['total_conc_sacs']), f"≈ {b['total_conc_sacs']*50/1000:.0f} t équivalent"],
    ["Ratio global soja/conc", f"{b['ratio']}:1", "✅ Excellent (objectif ≤ 2,5:1)" if b['ratio'] <= 2.5 else "⚠ Au-dessus objectif"],
    ["Commandes soja", str(b['cmds_soja']), "—"],
    ["Commandes bundle (soja+conc)", str(b['cmds_bundle']), f"{b['pct_bundle']}% des commandes soja"],
    ["Commandes soja-only (sans conc)", str(b['cmds_soja_only']), "Cross-sell à améliorer" if b['cmds_soja_only'] > 50 else "✅ Bon cross-sell"],
    ["Commandes conc-only (sans soja)", str(b['cmds_conc_only']), "—"],
]
bundle_table = make_table(bundle_data, col_widths=[5*cm, 4*cm, 7*cm], font_size=9)
story.append(KeepTogether(heading_block("2. Bundle Soja-Concentrés — Septembre 2026") + [bundle_table]))
story.append(Spacer(1, 0.3*cm))

# Distribution des ratios
dist = b['dist']
dist_total = sum(dist.values()) if dist else 1
dist_data = [
    ["Tranche ratio", "Nombre commandes", "Part des commandes bundle"],
    ["≤ 3:1 (excellent)", str(dist.get('<= 3:1', 0)), f"{dist.get('<= 3:1', 0)/dist_total*100:.1f}%"],
    ["3-5:1 (acceptable)", str(dist.get('3-5:1', 0)), f"{dist.get('3-5:1', 0)/dist_total*100:.1f}%"],
    ["5-10:1 (à surveiller)", str(dist.get('5-10:1', 0)), f"{dist.get('5-10:1', 0)/dist_total*100:.1f}%"],
    ["10-20:1 (action requise)", str(dist.get('10-20:1', 0)), f"{dist.get('10-20:1', 0)/dist_total*100:.1f}%"],
    ["> 20:1 (critique)", str(dist.get('> 20:1', 0)), f"{dist.get('> 20:1', 0)/dist_total*100:.1f}%"],
    ["TOTAL", str(dist_total), "100.0%"],
]
dist_table = make_table(dist_data, col_widths=[5*cm, 5*cm, 6*cm], font_size=9, highlight_rows=[6])
story.append(KeepTogether([Paragraph("<b>Distribution des ratios bundle (commandes avec soja + conc)</b>", H3), dist_table]))
story.append(Spacer(1, 0.3*cm))

bundle_analysis = f"""
<b>Lecture Bundle Septembre</b> :
Le ratio global soja/concentrés est de <b>{b['ratio']}:1</b> —
{"<b>conforme à l'objectif 2,5:1</b>, témoignant d'un excellent cross-sell" if b['ratio'] <= 2.5 else "<b>au-dessus de l'objectif 2,5:1</b>, à surveiller"}.
Sur {b['cmds_soja']} commandes soja, <b>{b['cmds_bundle']} ({b['pct_bundle']}%) incluent des concentrés</b> —
{'niveau de cross-sell excellent' if b['pct_bundle'] >= 90 else ('cross-sell solide' if b['pct_bundle'] >= 75 else 'cross-sell à améliorer')}.
Seulement {b['cmds_soja_only']} commandes soja-only ({b['cmds_soja_only']/b['cmds_soja']*100:.1f}% des commandes soja) —
{'✅ aligné avec le plan d\'action bundle' if b['cmds_soja_only']/b['cmds_soja']*100 <= 10 else '⚠ taux de soja-only élevé'}.
La baisse de prix du 22/09 n'a pas dégradé la discipline bundle : le ratio est resté stable à 2,5:1 sur l'ensemble du mois.
"""
story.append(Paragraph(bundle_analysis, BODY))
story.append(Spacer(1, 0.3*cm))

story.append(PageBreak())

# === Section 3: Stock Soja ===

s = DATA['stock']
stock_data = [
    ["Indicateur", "Valeur", "Détail"],
    ["Date de référence stock", s['date'], f"Inventaire BEKOKO au {s['date']} — hyp. stock inchangé depuis"],
    ["Stock BRUT central (sacs 50kg éq.)", fmt(s['brut_sacs']), f"{s['brut_t']} t — T102: {s['t102_sacs']} + T1021: {s['t1021_kg']} kg (1kg) + T1023: {s['t1023_kg']} kg (5kg) + T1024: {s['t1024_kg']} kg (25kg)"],
    ["Vente soja septembre (sacs/jour)", fmt(s['vente_sacs_jour']), f"≈ {s['vente_sacs_jour']*50/1000:.0f} t/jour (moyenne mois complet)"],
    ["Jours de stock central", f"{s['jours_stock']} jours", "⚠ Sous le seuil de sécurité (10j) — couvre les 4 premiers jours d'octobre"],
    ["Date rupture CENTRALE estimée", s['rupture_date'], "⚠ RÉAPPRO URGENT avant cette date"],
    ["Besoin couvert fin septembre", "0 sac", "✅ Mois clôturé sans rupture (couverture assurée jusqu'au 30/09)"],
    ["Manque vs besoin début octobre", f"{fmt(s['vente_sacs_jour'])} sacs/j", "Réappro à programmer immédiatement pour octobre"],
    ["Stock précédent", fmt(s['precedent_stock_brut_sacs']), f"au {s['precedent_stock_date']} — delta +{s['variation_pct']}% (réappro reçu)"],
]
stock_table = make_table(stock_data, col_widths=[5.5*cm, 4*cm, 6.5*cm], font_size=8, highlight_rows=[4,5,7])
story.append(KeepTogether(heading_block("3. Stock Soja BEKOKO — Situation au 30/09") + [stock_table]))
story.append(Spacer(1, 0.3*cm))

stock_analysis = f"""
<b>⚠ STOCK SOUS LE SEUIL DE SÉCURITÉ — URGENCE OCTOBRE</b> : Le magasin central BEKOKO contient <b>{s['brut_sacs']} sacs</b>
(équivalent 50kg, soit {s['brut_t']} t) au {s['date']} — sous hypothèse que ce stock n'a pas été entamé depuis l'inventaire.
Au rythme de vente de septembre ({s['vente_sacs_jour']} sacs/jour), ceci représente <b>{s['jours_stock']} jours de stock</b> —
<b>rupture centrale estimée au {s['rupture_date']}</b>. Le mois de septembre est clôturé sans rupture, mais la couverture
n'excède pas les 4 premiers jours d'octobre : <b>le réapprovisionnement doit être déclenché immédiatement</b>.

<b>Cadence de fin de mois soutenue par la baisse de prix</b> : depuis la baisse du soja à 20 000 FCFA/sac le 22/09, la cadence
s'est établie à <b>278,1 t/j</b> (vs 139,1 t/j du 01 au 22/09), soit un rythme de consommation doublé qui pèsera sur le stock
dès les premiers jours d'octobre si la demande reste soutenue.

<b>Réapprovisionnement reçu en septembre</b> : le stock est passé de {s['precedent_stock_brut_sacs']} sacs ({s['precedent_stock_date']})
à {s['brut_sacs']} sacs ({s['date']}), soit +{s['variation_pct']}% (+{s['brut_sacs'] - s['precedent_stock_brut_sacs']} sacs).

<b>Besoin début octobre</b> : {fmt(s['vente_sacs_jour'])} sacs/jour au rythme actuel — à la cadence post-baisse de prix
(≈ 5 560 sacs/j), le stock central serait consommé en {s['jours_stock']} jour(s) à peine. Un nouveau réappro est
<b>nécessaire avant le {s['rupture_date']}</b>.
"""
story.append(Paragraph(stock_analysis, BODY))
story.append(Spacer(1, 0.3*cm))

# === Section 4: Zero-Achat & Cross-sell ===

z = DATA['zero_achat']
c = DATA['cross_sell']

za_data = [
    ["Indicateur", "Valeur", "Lecture"],
    ["Clients actifs S1 2026", str(z['s1_clients']), "Base de référence (Jan-Juin 2026)"],
    ["Clients actifs Septembre", str(z['sept_clients']), f"{z['sept_clients']/z['s1_clients']*100:.1f}% de la base S1"],
    ["Churned (S1 sans achat Sept)", str(z['churned']), f"{z['churned']/z['s1_clients']*100:.1f}% de churn"],
    ["Nouveaux clients Sept", str(z['new_sept']), "Clients hors S1 ayant acheté en sept"],
    ["S1 clients soja", str(c['s1_soja']), "—"],
    ["S1 clients conc", str(c['s1_conc']), "—"],
    ["S1 clients bundle (soja+conc)", str(c['s1_bundle']), f"{c['s1_bundle']/c['s1_soja']*100:.0f}% des clients soja S1"],
    ["Sept clients soja", str(c['sept_soja']), "—"],
    ["Sept clients conc", str(c['sept_conc']), "—"],
    ["Sept clients bundle", str(c['sept_bundle']), f"{c['sept_bundle']/c['sept_soja']*100:.0f}% des clients soja Sept"],
    ["S1 soja clients → conc en Sept", str(c['soja_to_conc_sept']), f"{c['pct_soja_with_conc_sept']}% des S1 soja"],
]
za_table = make_table(za_data, col_widths=[5*cm, 3*cm, 8*cm], font_size=9)
story.append(KeepTogether(heading_block("4. Zero-Achat & Cross-Sell — Septembre 2026 (mois complet)") + [za_table]))
story.append(Spacer(1, 0.3*cm))

za_analysis = f"""
<b>Lecture Zero-Achat (mois complet)</b> : Sur les {z['s1_clients']} clients actifs en S1 2026, <b>{z['sept_clients']}
({z['sept_clients']/z['s1_clients']*100:.1f}%) ont acheté en septembre</b>. <b>{z['churned']} clients S1
({z['churned']/z['s1_clients']*100:.1f}%) sont restés sans achat sur tout le mois</b> — ce sont les cibles prioritaires de
réactivation pour octobre. <b>{z['new_sept']} nouveaux clients</b> (hors S1) ont acheté en septembre, compensant partiellement
le churn : le solde net de la base clients est de {z['new_sept'] - z['churned']} clients. La forte dynamique de fin de mois
(poste baisse de prix du 22/09) montre que la demande latente existe — la réactivation des churned passe par la continuité
de la politique prix d'octobre et le démarchage ciblé par agence.

<b>Cross-sell</b> : Sur {c['s1_soja']} clients S1 soja, <b>{c['soja_to_conc_sept']} ({c['pct_soja_with_conc_sept']}%)
ont aussi acheté du concentré en septembre</b>. Ce taux est plus faible que la performance bundle observée
({b['pct_bundle']}% des commandes), ce qui indique que le cross-sell est meilleur sur les nouveaux clients
septembre que sur les clients S1 historiques — opportunité de réactivation commerciale sur les S1 soja-only.
"""
story.append(Paragraph(za_analysis, BODY))
story.append(Spacer(1, 0.3*cm))

# === Section 5: Recommandations ===

recos = []
status_conc = "✅" if v_conc['pct_obj'] >= 100 else ("⚠" if v_conc['pct_obj'] >= 80 else "❌")
recos.append(f"{status_conc} <b>CONCENTRES — Résultat {v_conc['t']} t</b> vs objectif {v_conc['obj_t']} t = <b>{v_conc['pct_obj']}%</b> ({ecart_conc:+.0f} t). "
            + ("Objectif dépassé — capitaliser en octobre avec la cadence de fin de mois (78,2 t/j post 22/09)." if v_conc['pct_obj'] >= 100 else "Renforcer la cadence — action commerciale sur les agences en retard."))

status_bundle = "✅" if b['ratio'] <= 2.5 else "⚠"
recos.append(f"{status_bundle} <b>Bundle ratio {b['ratio']}:1</b> — "
            + (f"Conforme à l'objectif 2,5:1, cross-sell à {b['pct_bundle']}% des commandes soja — maintenir la discipline." if b['ratio'] <= 2.5 else "Au-dessus de l'objectif 2,5:1. Renforcer le push concentrés."))

recos.append(f"🚨 <b>STOCK SOJA — URGENCE DÉBUT OCTOBRE</b> : Magasin central BEKOKO = {s['brut_sacs']} sacs ({s['brut_t']} t, inventaire {s['date']}). "
            f"Au rythme de septembre ({s['vente_sacs_jour']} sacs/j), <b>{s['jours_stock']} jour(s) de stock</b> — rupture centrale estimée <b>{s['rupture_date']}</b>. "
            f"<b>Déclencher le réapprovisionnement immédiatement</b> (besoin ≈ {fmt(s['vente_sacs_jour'])} sacs/j, potentiellement 5 500+ sacs/j si la cadence post-baisse de prix se maintient).")

recos.append(f"🎯 <b>Cross-sell S1 soja → conc</b> : {c['pct_soja_with_conc_sept']}% des S1 soja ont pris du conc en sept. "
            f"Action de réactivation sur les {c['s1_soja'] - c['soja_to_conc_sept']} clients S1 soja sans conc en sept.")

recos.append(f"📉 <b>Politique prix octobre</b> : la baisse à 20 000 FCFA/sac a doublé la cadence soja (139,1 → 278,1 t/j) et porté le mois à "
            f"121% de l'objectif. Arbitrage prix/volume : au prix de 26 000 FCFA, la cadence était de 139,1 t/j (attentisme) ; à 20 000 FCFA, "
            f"278,1 t/j. Le maintien du prix bas soutient les volumes mais réduit la marge unitaire — décision DG requise pour octobre.")

# Top agence to push
if DATA['conc_by_agence']:
    top_ag = DATA['conc_by_agence'][0]
    recos.append(f"🏆 <b>Top agence Sept</b> : {top_ag['agence']} ({top_ag['conc_t']} t de concentrés) — "
                "utiliser comme modèle de performance pour les autres agences.")

# Bottom agence to push
if len(DATA['conc_by_agence']) > 5:
    bottom_ag = DATA['conc_by_agence'][-1]
    recos.append(f"⚠ <b>Agence à surveiller</b> : {bottom_ag['agence']} ({bottom_ag['conc_t']} t de concentrés) — "
                "renforcer le support commercial.")

reco_paras = [Paragraph(f"• {r}", BULLET) for r in recos]
story.append(KeepTogether(heading_block("5. Recommandations & Plan d'action") + reco_paras))

story.append(Spacer(1, 0.5*cm))
methodo_note = Paragraph(
    f"Les données de cette analyse sont issues de l'extraction <b>{DATA['extraction_file']}</b> "
    f"(mois complet, données définitives au {DATA['update_date']} — {DATA['days_elapsed']} jours ouvrés sur {DATA['total_days_sep']}). "
    f"Les objectifs de septembre sont les <b>objectifs S2 recalibrés</b> (TOURTEAUX {fmt(v_soja['obj_t'])} t, CONCENTRES {fmt(v_conc['obj_t'])} t). "
    "Les comparaisons avec août utilisent les données de synthèse d'août (extraction du 31/08/2026, objectifs S2 recalibrés août). "
    "Événements prix : hausse du soja le 04/09 (25 000 → 26 000 FCFA/sac) puis baisse mi-journée du 22/09 (26 000 → 20 000 FCFA/sac) "
    "avec doublement de la cadence sur les 7 derniers jours ouvrés. "
    "Le stock central BEKOKO est daté du 16/09/2026 (dernier inventaire communiqué) — les jours de couverture sont estimés sous hypothèse de stock inchangé. "
    "Tous les clients de l'extraction ERP sont inclus (y compris comptoirs internes SPC/PDC), conformément au périmètre de l'analyse.",
    SMALL)
story.append(KeepTogether([Paragraph("<b>Note méthodologique</b>", H3), methodo_note]))

# === Save PDF ===
doc.build(story, canvasmaker=NumberedCanvas)
print(f"\n=== PDF GENERATED ===")
print(f"Path: {OUT}")
print(f"Size: {os.path.getsize(OUT)/1024:.0f} KB")
