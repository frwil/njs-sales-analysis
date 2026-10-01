"""Generate DG pitch PDF — 1-page executive summary for the Director General.
Includes nationwide volumes vs objectives, regional breakdown, key insights, recommendations and actions.

Based on sept_pitch_data_10.json and sept_mtd_10.json (mois complet 01-30/09/2026).
"""
import os
import json
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, KeepTogether
)

# === Fonts ===
FONT_DIR = '/usr/share/fonts'
pdfmetrics.registerFont(TTFont('DejaVuSans', f'{FONT_DIR}/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', f'{FONT_DIR}/truetype/dejavu/DejaVuSans-Bold.ttf'))

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

styles = getSampleStyleSheet()
H1 = ParagraphStyle('H1', parent=styles['Heading1'], fontName='DejaVuSans-Bold', fontSize=16, textColor=NAVY, spaceAfter=8, spaceBefore=4)
H2 = ParagraphStyle('H2', parent=styles['Heading2'], fontName='DejaVuSans-Bold', fontSize=12, textColor=NAVY, spaceAfter=6, spaceBefore=8)
H3 = ParagraphStyle('H3', parent=styles['Heading3'], fontName='DejaVuSans-Bold', fontSize=10, textColor=GOLD, spaceAfter=4, spaceBefore=6)
BODY = ParagraphStyle('Body', parent=styles['Normal'], fontName='DejaVuSans', fontSize=9, leading=12, alignment=TA_JUSTIFY, spaceAfter=4)
BULLET = ParagraphStyle('Bullet', parent=BODY, leftIndent=15, bulletIndent=5, spaceAfter=3)
CELL = ParagraphStyle('Cell', parent=BODY, fontName='DejaVuSans', fontSize=9, leading=11, alignment=TA_LEFT, spaceAfter=0)
SMALL = ParagraphStyle('Small', parent=BODY, fontSize=7.5, textColor=GRAY, leading=10)

# Footer with page numbers ("Page X / Y") on every page
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
            self.draw_page_header()
            self.draw_page_footer(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_header(self):
        self.saveState()
        self.setFont('DejaVuSans-Bold', 7.5)
        self.setFillColor(NAVY)
        self.drawString(1.5*cm, A4[1]-0.75*cm, "BELGOCAM SA — NJS GROUP")
        self.setFont('DejaVuSans', 7.5)
        self.setFillColor(GRAY)
        self.drawRightString(A4[0]-1.5*cm, A4[1]-0.75*cm, f"Pitch DG — Données au {MTD['update_date']}")
        self.setStrokeColor(GOLD)
        self.setLineWidth(0.8)
        self.line(1.5*cm, A4[1]-1.0*cm, A4[0]-1.5*cm, A4[1]-1.0*cm)
        self.restoreState()

    def draw_page_footer(self, num_pages):
        self.saveState()
        self.setStrokeColor(colors.HexColor('#BFBFBF'))
        self.setLineWidth(0.5)
        self.line(1.5*cm, 1.15*cm, A4[0]-1.5*cm, 1.15*cm)
        self.setFont('DejaVuSans', 7.5)
        self.setFillColor(GRAY)
        self.drawString(1.5*cm, 0.75*cm, "BELGOCAM SA — Pitch Septembre 2026 — Confidentiel")
        self.drawRightString(A4[0]-1.5*cm, 0.75*cm, f"Page {self._pageNumber} / {num_pages}")
        self.restoreState()

# Cell styles for tables (wrapping)
CELL_HEADER_P = ParagraphStyle('CellHeaderP', fontName='DejaVuSans-Bold', fontSize=8, leading=10, alignment=TA_CENTER, textColor=colors.white)
CELL_BODY_P = ParagraphStyle('CellBodyP', fontName='DejaVuSans', fontSize=8, leading=10, alignment=TA_LEFT)
CELL_BODY_CENTER_P = ParagraphStyle('CellBodyCenterP', fontName='DejaVuSans', fontSize=8, leading=10, alignment=TA_CENTER)

def wrap_cell_p(content, style=CELL_BODY_P):
    """Wrap content in a Paragraph for proper text wrapping in tables."""
    s = str(content) if content is not None else ''
    return Paragraph(s, style)

# === Load data ===
PITCH = json.load(open('/home/z/my-project/scripts/sept_pitch_data_10.json'))
MTD = json.load(open('/home/z/my-project/scripts/sept_mtd_10.json'))
print(f"Loaded sept data (update {MTD['update_date']})")

# Helpers
def fmt(x):
    if isinstance(x, (int, float)):
        return f"{int(round(x)):,}".replace(',', ' ')
    return str(x)

def fmt1(x):
    return f"{x:,.1f}".replace(',', ' ').replace('.', ',')

def status_color(pct):
    if pct >= 90: return LIGHT_GREEN
    if pct >= 70: return colors.HexColor('#FFF9E6')  # light yellow
    return LIGHT_RED

def status_text(pct):
    if pct >= 90: return "✅"
    if pct >= 70: return "⚠"
    return "❌"

# === Build PDF ===
OUT = '/home/z/my-project/download/pitch_dg_septembre_upd.pdf'
doc = SimpleDocTemplate(OUT, pagesize=A4, topMargin=1.2*cm, bottomMargin=1.2*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
story = []

# === HEADER ===
story.append(Paragraph("BELGOCAM SA — Pitch Septembre 2026 (mois complet)", H1))
story.append(Paragraph(f"Performance Commerciale définitive au {MTD['update_date']} ({MTD['days_elapsed']} jours ouvrés — mois complet)",
                       ParagraphStyle('SubH', parent=BODY, fontSize=10, textColor=GRAY, alignment=TA_CENTER, spaceAfter=4)))
story.append(Paragraph(f"<b>⚠ Baisse prix soja le 22/09/2026 (mi-journée) : 26 000 → 20 000 FCFA/sac (-23%) → cadence doublée (139 → 278 t/j)</b>",
                       ParagraphStyle('Alert', parent=BODY, fontName='DejaVuSans-Bold', fontSize=10, textColor=RED, alignment=TA_CENTER, spaceAfter=4)))
story.append(Paragraph(f"<b>🚨 STOCK SOUS SEUIL — Magasin central BEKOKO : {fmt(MTD['stock']['brut_sacs'])} sacs ({fmt(MTD['stock']['brut_t'])} t, inventaire {MTD['stock']['date']}). Rupture estimée {MTD['stock']['rupture_date']} ({MTD['stock']['jours_stock']} jours de stock) — réappro urgent</b>",
                       ParagraphStyle('Alert2', parent=BODY, fontName='DejaVuSans-Bold', fontSize=10, textColor=RED, alignment=TA_CENTER, spaceAfter=8, backColor=colors.HexColor('#FCE4EC'), borderPadding=4)))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.2*cm))

# === 1. NATIONWIDE ===
n = PITCH['nationwide']
nat_data = [
    ["Indicateur", "Réel Sept", "Moy/j", "Objectif", "% Obj", "Statut"],
    ["TOURTEAUX (Soja)", f"{fmt(n['soja_t_mtd'])} t", f"{fmt1(n['soja_moy_t_j'])} t/j", f"{fmt(n['soja_obj_t'])} t", f"{n['soja_pct_obj']:.0f}%", status_text(n['soja_pct_obj'])],
    ["CONCENTRÉS", f"{fmt(n['conc_t_mtd'])} t", f"{fmt1(n['conc_moy_t_j'])} t/j", f"{fmt(n['conc_obj_t'])} t", f"{n['conc_pct_obj']:.0f}%", status_text(n['conc_pct_obj'])],
    ["Total Soja+Conc", f"{fmt(n['soja_t_mtd']+n['conc_t_mtd'])} t", f"{fmt1(n['soja_moy_t_j']+n['conc_moy_t_j'])} t/j", f"{fmt(n['soja_obj_t']+n['conc_obj_t'])} t", f"{(n['soja_t_mtd']+n['conc_t_mtd'])/(n['soja_obj_t']+n['conc_obj_t'])*100:.0f}%", "—"],
    ["Ratio bundle", f"{n['ratio_global']:.1f}:1", "—", "≤ 2,5:1", "—", "✅" if n['ratio_global'] <= 2.5 else "⚠"],
    ["Mix ventes", f"Soja {n['pct_soja_volume']:.0f}% / Conc {n['pct_conc_volume']:.0f}%", "—", "—", "—", "—"],
]

# Build table with conditional coloring
t = Table(nat_data, colWidths=[3.6*cm, 2.3*cm, 1.7*cm, 2.1*cm, 1.4*cm, 1.3*cm], repeatRows=1)
style_list = [
    ('FONT', (0,0), (-1,0), 'DejaVuSans-Bold', 9),
    ('FONT', (0,1), (-1,-1), 'DejaVuSans', 9),
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
    ('ALIGN', (1,0), (-1,-1), 'CENTER'),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
    ('TOPPADDING', (0,0), (-1,-1), 4), ('BOTTOMPADDING', (0,0), (-1,-1), 4),
]
# Color soja row (1) based on pct
style_list.append(('BACKGROUND', (0,1), (-1,1), status_color(n['soja_pct_obj'])))
style_list.append(('BACKGROUND', (0,2), (-1,2), status_color(n['conc_pct_obj'])))
style_list.append(('BACKGROUND', (0,3), (-1,3), colors.HexColor('#FFF2CC')))  # total in yellow
style_list.append(('BACKGROUND', (0,4), (-1,4), colors.HexColor('#E2EFDA')))  # ratio in green
style_list.append(('BACKGROUND', (0,5), (-1,5), LIGHT_GRAY))  # mix
t.setStyle(TableStyle(style_list))
story.append(KeepTogether([Paragraph("1. VUES NATIONALES — Ventes vs Objectifs", H2), t]))
story.append(Spacer(1, 0.15*cm))

# Key insight for nationwide
nationwide_insight = f"""
<b>Synthèse nationwide</b> : Septembre se clôture à <b>{fmt(n['soja_t_mtd']+n['conc_t_mtd'])} t</b>
(soja + concentrés), soit <b>{(n['soja_t_mtd']+n['conc_t_mtd'])/(n['soja_obj_t']+n['conc_obj_t'])*100:.0f}% de l'objectif combiné</b>.
Le <b>SOJA atteint {n['soja_pct_obj']:.0f}%</b> de l'objectif {status_text(n['soja_pct_obj'])} — la <b>baisse de prix du 22/09 (26 000 → 20 000 FCFA/sac)</b> a doublé la cadence :
139,1 t/j (01-22/09) → <b>278,1 t/j (23-30/09, +100%)</b>, avec un pic à 471 t le 23/09. Le <b>CONCENTRES atteint {n['conc_pct_obj']:.0f}%</b> {status_text(n['conc_pct_obj'])}.
Le <b>ratio bundle {n['ratio_global']:.1f}:1</b> est conforme à l'objectif (≤ 2,5:1) — le cross-sell s'est maintenu malgré la volatilité prix.
"""
story.append(Paragraph(nationwide_insight, BODY))
story.append(Spacer(1, 0.2*cm))

# === 2. RÉGIONS ===
reg_data = [["Région", "Soja Sept (t)", "Soja obj (t)", "% Soja", "Conc Sept (t)", "Conc obj (t)", "% Conc", "Ratio"]]
total_soja_mtd = 0
total_conc_mtd = 0
total_soja_obj = 0
total_conc_obj = 0
for region in ['Ouest', 'Centre', 'Littoral']:
    r = PITCH['regions'][region]
    total_soja_mtd += r['soja_t_mtd']
    total_conc_mtd += r['conc_t_mtd']
    total_soja_obj += r['soja_obj_t']
    total_conc_obj += r['conc_obj_t']
    reg_data.append([
        region,
        fmt(r['soja_t_mtd']),
        fmt(r['soja_obj_t']),
        f"{r['soja_pct_obj']:.0f}%",
        fmt(r['conc_t_mtd']),
        fmt(r['conc_obj_t']),
        f"{r['conc_pct_obj']:.0f}%",
        f"{r['ratio']:.1f}:1",
    ])
reg_data.append([
    "TOTAL",
    fmt(total_soja_mtd),
    fmt(total_soja_obj),
    f"{total_soja_mtd/total_soja_obj*100:.0f}%",
    fmt(total_conc_mtd),
    fmt(total_conc_obj),
    f"{total_conc_mtd/total_conc_obj*100:.0f}%",
    f"{total_soja_mtd/total_conc_mtd:.1f}:1" if total_conc_mtd > 0 else "—",
])

t2 = Table(reg_data, colWidths=[2.0*cm, 1.9*cm, 1.7*cm, 1.2*cm, 1.9*cm, 1.7*cm, 1.2*cm, 1.4*cm], repeatRows=1)
style_list2 = [
    ('FONT', (0,0), (-1,0), 'DejaVuSans-Bold', 8),
    ('FONT', (0,1), (-1,-1), 'DejaVuSans', 8),
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
    ('ALIGN', (1,0), (-1,-1), 'CENTER'),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
    ('TOPPADDING', (0,0), (-1,-1), 3), ('BOTTOMPADDING', (0,0), (-1,-1), 3),
]
# Color regional rows
for i, region in enumerate(['Ouest', 'Centre', 'Littoral'], 1):
    r = PITCH['regions'][region]
    # Color % cells based on performance
    style_list2.append(('BACKGROUND', (3, i), (3, i), status_color(r['soja_pct_obj'])))
    style_list2.append(('BACKGROUND', (6, i), (6, i), status_color(r['conc_pct_obj'])))
# Total row
style_list2.append(('BACKGROUND', (0, 4), (-1, 4), colors.HexColor('#FFF2CC')))
style_list2.append(('FONT', (0, 4), (-1, 4), 'DejaVuSans-Bold', 8))
t2.setStyle(TableStyle(style_list2))
story.append(KeepTogether([Paragraph("2. VUES RÉGIONALES — Performances par région", H2), t2]))
story.append(Spacer(1, 0.15*cm))

# Regional insight
ouest = PITCH['regions']['Ouest']
centre = PITCH['regions']['Centre']
littoral = PITCH['regions']['Littoral']
reg_insight = f"""
<b>Lecture régionale</b> :
<b>OUEST</b> est la région leader ({fmt(ouest['soja_t_mtd']+ouest['conc_t_mtd'])} t, soja {ouest['soja_pct_obj']:.0f}% {status_text(ouest['soja_pct_obj'])}, conc {ouest['conc_pct_obj']:.0f}% {status_text(ouest['conc_pct_obj'])}).
<b>CENTRE</b> enregistre la plus forte surperformance (soja {centre['soja_pct_obj']:.0f}% {status_text(centre['soja_pct_obj'])}, conc {centre['conc_pct_obj']:.0f}% {status_text(centre['conc_pct_obj'])}).
<b>LITTORAL</b> est la seule région sous l'objectif : soja {littoral['soja_pct_obj']:.0f}% {status_text(littoral['soja_pct_obj'])}, conc {littoral['conc_pct_obj']:.0f}% {status_text(littoral['conc_pct_obj'])}, ratio dégradé {littoral['ratio']:.1f}:1 —
la baisse de prix du 22/09 et le maintien à 20 000 FCFA/sac sont une fenêtre pour y relancer les volumes soja (NDOBO, VILLAGE) et renforcer le cross-sell concentrés.
"""
story.append(Paragraph(reg_insight, BODY))
story.append(Spacer(1, 0.2*cm))

# === 3. INSIGHTS CLÉS ===

# Get bundle stats from MTD
b = MTD['bundle']
s = MTD['stock']
z = MTD['zero_achat']
top_ag = MTD['conc_by_agence'][0]
total_conc_mtd = sum(a['conc_t'] for a in MTD['conc_by_agence'])
bottom_ags = [a['agence'] for a in MTD['conc_by_agence'][-2:]]

insights_data = [
    [Paragraph("<b>Insight</b>", CELL), Paragraph("<b>Donnée</b>", CELL), Paragraph("<b>Implication</b>", CELL)],
    [Paragraph("Baisse prix soja 22/09 (mi-journée)", CELL),
     Paragraph("26 000 → 20 000 FCFA/sac (-23%)", CELL),
     Paragraph("Doublement de la cadence : 139,1 t/j (01-22/09) → <b>278,1 t/j (23-30/09, +100%)</b>, pic 471 t le 23/09. Mois clôturé à 121% de l'objectif soja — arbitrage prix/volume à trancher pour octobre", CELL)],
    [Paragraph(f"Bundle ratio {b['ratio']}:1", CELL),
     Paragraph(f"{b['cmds_bundle']}/{b['cmds_soja']} cmds soja avec conc = {b['pct_bundle']}% cross-sell", CELL),
     Paragraph("✅ Conforme à l'objectif 2,5:1 — le cross-sell s'est maintenu malgré la volatilité prix (hausse 04/09 puis baisse 22/09)", CELL)],
    [Paragraph("🔴 Stock soja BEKOKO", CELL),
     Paragraph(f"<b>{fmt(s['brut_sacs'])} sacs</b> ({fmt(s['brut_t'])} t, inventaire {s['date']})", CELL),
     Paragraph(f"<b>{s['jours_stock']} jour(s) de stock central</b> — rupture estimée <b>{s['rupture_date']}</b>. Septembre clôturé sans rupture mais couverture quasi nulle : <b>réappro urgent</b> pour début octobre", CELL)],
    [Paragraph("Top agence CONCENTRES", CELL),
     Paragraph(f"{top_ag['agence'].upper()} — {fmt(top_ag['conc_t'])} t en septembre", CELL),
     Paragraph(f"{top_ag['conc_t']/total_conc_mtd*100:.0f}% du volume CONCENTRES national — pilier de la performance", CELL)],
    [Paragraph("Agences en retrait", CELL),
     Paragraph(", ".join(bottom_ags), CELL),
     Paragraph("Dernières du classement CONCENTRES — activer le relais de la baisse prix soja pour doper le bundle", CELL)],
]
t3 = Table(insights_data, colWidths=[4*cm, 6*cm, 7.5*cm], repeatRows=1)
t3.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
    ('FONT', (0,0), (-1,0), 'DejaVuSans-Bold', 9),
    ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_GRAY]),
    ('TOPPADDING', (0,0), (-1,-1), 4), ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ('LEFTPADDING', (0,0), (-1,-1), 4), ('RIGHTPADDING', (0,0), (-1,-1), 4),
]))
story.append(KeepTogether([Paragraph("3. Insights clés", H2), t3]))
story.append(Spacer(1, 0.2*cm))

# === 4. RECOMMANDATIONS ===
recos = [
    ("ACTION IMMÉDIATE", "🚨 RÉAPPRO URGENT soja central",
     f"Stock BEKOKO = {fmt(s['brut_sacs'])} sacs ({fmt(s['brut_t'])} t, inventaire {s['date']}). <b>{s['jours_stock']} jour(s) de stock</b> — rupture estimée {s['rupture_date']}. À la cadence post-baisse de prix (≈ 5 560 sacs/j), le stock est consommé en {s['jours_stock']} j. <b>Réappro à positionner avant le {s['rupture_date']}</b>.",
     "Logistique / Direction Achats / DG", "AVANT 04/10/2026"),
    ("ACTION COMMERCIALE", "Relancer le SOJA sur Littoral",
     f"Littoral = {PITCH['regions']['Littoral']['soja_t_mtd']:.0f} t vs objectif {PITCH['regions']['Littoral']['soja_obj_t']:.0f} t ({PITCH['regions']['Littoral']['soja_pct_obj']:.0f}% obj), conc {PITCH['regions']['Littoral']['conc_pct_obj']:.0f}% obj — seule région sous l'objectif. Capitaliser sur le prix bas (20 000 FCFA/sac) pour relancer soja (NDOBO, VILLAGE) et renforcer le cross-sell.",
     "RA Littoral + Direction Commerciale", "Semaine 40 (28/09-04/10)"),
    ("ACTION COMMERCIALE", "Maintenir le cross-sell bundle",
     f"Ratio {b['ratio']}:1 en septembre ✅ — poursuivre la discipline bundle (objectif ≤ 2,5:1). Push CONCENTRES en cross-sell sur les {b['cmds_soja_only']} cmds soja-only et les {b['cmds_conc_only']} cmds conc-only.",
     "Toutes agences", "Continu"),
    ("SURVEILLANCE", "Suivi stock central vs cadence post-baisse",
     f"Cadence soja doublée depuis le 22/09 (278 t/j ≈ 5 560 sacs/j) — surveiller quotidiennement le stock central (4 j de couverture) et les livraisons agences face à l'accélération des ventes.",
     "Data Analyst + Direction Commerciale", "Point quotidien 01-10/10"),
    ("STRATÉGIQUE", "Trancher la politique prix octobre",
     f"Élasticité observée : -23% de prix → +100% de cadence (139 → 278 t/j), mois à 121% de l'objectif soja. Maintien à 20 000 FCFA/sac = volumes soutenus mais marge réduite ; retour à 26 000 = attentisme (cadence 139 t/j). <b>Décision DG requise</b> avant la semaine 41.",
     "DG + Direction Financière", "Semaine 40"),
    ("OPPORTUNITÉ", "Réactivation des clients S1 sans achat Sept",
     f"{z['churned']} clients S1 sur {z['s1_clients']} ({z['churned']/z['s1_clients']*100:.0f}%) sans achat en septembre (mois complet). Le prix bas actuel est une fenêtre de réactivation — campagne ciblée sur les top 200 clients S1 inactifs.",
     "Administrateurs de Vente", "Semaines 40-41"),
]

recos_data = [[wrap_cell_p(h, CELL_HEADER_P) for h in ["Priorité", "Action", "Détail", "Responsable", "Échéance"]]]
for r in recos:
    recos_data.append([
        wrap_cell_p(r[0], CELL_BODY_CENTER_P),
        wrap_cell_p(r[1], CELL_BODY_P),
        wrap_cell_p(r[2], CELL_BODY_P),
        wrap_cell_p(r[3], CELL_BODY_P),
        wrap_cell_p(r[4], CELL_BODY_CENTER_P),
    ])

t4 = Table(recos_data, colWidths=[2.5*cm, 3.2*cm, 6.3*cm, 3*cm, 2.5*cm], repeatRows=1)
style_list4 = [
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
    ('TOPPADDING', (0,0), (-1,-1), 4), ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ('LEFTPADDING', (0,0), (-1,-1), 4), ('RIGHTPADDING', (0,0), (-1,-1), 4),
]
# Color priority column
priority_colors = {
    'ACTION IMMÉDIATE': LIGHT_RED,
    'ACTION COMMERCIALE': colors.HexColor('#FFF2CC'),
    'SURVEILLANCE': LIGHT_GRAY,
    'STRATÉGIQUE': colors.HexColor('#E2EFDA'),
    'OPPORTUNITÉ': colors.HexColor('#DEEBF7'),
}
for i, r in enumerate(recos, 1):
    pc = priority_colors.get(r[0], LIGHT_GRAY)
    style_list4.append(('BACKGROUND', (0, i), (0, i), pc))
    style_list4.append(('FONT', (0, i), (0, i), 'DejaVuSans-Bold', 7.5))
t4.setStyle(TableStyle(style_list4))

# === FOOTER ===
# Section 4 + méthodo soudés : ni tableau coupé, ni fin de paragraphe orpheline en bas de page
methodo_note = Paragraph(
    f"<b>Source</b> : {MTD['extraction_file']} (mois complet au {MTD['update_date']} — {MTD['days_elapsed']} jours ouvrés). "
    f"<b>Méthodologie</b> : Volumes Livrées uniquement, tous clients inclus. Les objectifs sont les <b>objectifs S2 recalibrés</b> de septembre "
    f"(TOURTEAUX {fmt(n['soja_obj_t'])} t, CONCENTRES {fmt(n['conc_obj_t'])} t ; objectifs régionaux = somme des objectifs agences S2 recalibrés). "
    f"<b>Événements prix soja</b> : hausse le 04/09/2026 (25 000 → 26 000 FCFA/sac), puis baisse mi-journée du 22/09/2026 (26 000 → 20 000 FCFA/sac, -23%) avec doublement de la cadence (139,1 → 278,1 t/j). "
    f"<b>Stock central BEKOKO</b> : {MTD['stock']['brut_sacs']} sacs ({MTD['stock']['brut_t']} t) au {MTD['stock']['date']} — hyp. stock inchangé depuis l'inventaire.",
    SMALL)
story.append(KeepTogether([Paragraph("4. Recommandations & Actions", H2), t4,
                           HRFlowable(width="100%", thickness=0.5, color=GRAY),
                           Spacer(1, 0.1*cm), methodo_note]))

# === Save PDF ===
doc.build(story, canvasmaker=NumberedCanvas)
print(f"\n=== PITCH PDF GENERATED ===")
print(f"Path: {OUT}")
print(f"Size: {os.path.getsize(OUT)/1024:.0f} KB")
