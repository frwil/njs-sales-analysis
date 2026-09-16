"""Generate DG pitch PDF — 1-page executive summary for the Director General.
Includes nationwide volumes vs objectives, regional breakdown, key insights, recommendations and actions.

Based on sept_pitch_data.json and sept_mtd_01.json (computed September 5, 2026).
Highlights: Soja price increased to 27,000 FCFA on 04/09/2026.
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
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
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

# === Load data ===
PITCH = json.load(open('/home/z/my-project/scripts/sept_pitch_data_04.json'))
MTD = json.load(open('/home/z/my-project/scripts/sept_mtd_04.json'))
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
story.append(Paragraph("BELGOCAM SA — Pitch Septembre 2026 MTD", H1))
story.append(Paragraph(f"Performance Commerciale au {MTD['update_date']} ({MTD['days_elapsed']}j/{MTD['total_days_sep']}j = {MTD['pct_elapsed']}% du mois)", 
                       ParagraphStyle('SubH', parent=BODY, fontSize=10, textColor=GRAY, alignment=TA_CENTER, spaceAfter=4)))
story.append(Paragraph(f"<b>⚠ Hausse prix soja le 04/09/2026 : 25 000 → 26 000 FCFA/sac (+4%)</b>",
                       ParagraphStyle('Alert', parent=BODY, fontName='DejaVuSans-Bold', fontSize=10, textColor=RED, alignment=TA_CENTER, spaceAfter=4)))
story.append(Paragraph(f"<b>🔴 ALERTE STOCK CRITIQUE — Magasin central BEKOKO : 2 716 sacs (135 t) + 296 prod = 3 012 sacs. Rupture 08/09/2026 (1,6 jour de stock)</b>",
                       ParagraphStyle('Alert2', parent=BODY, fontName='DejaVuSans-Bold', fontSize=10, textColor=RED, alignment=TA_CENTER, spaceAfter=8, backColor=colors.HexColor('#FCE4EC'), borderPadding=4)))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.2*cm))

# === 1. NATIONWIDE ===
story.append(Paragraph("1. VUES NATIONALES — Ventes vs Objectifs", H2))

n = PITCH['nationwide']
nat_data = [
    ["Indicateur", "Volume MTD", "Moy/j", "Projection fin Sept", "Objectif", "% Obj", "Statut"],
    ["TOURTEAUX (Soja)", f"{fmt(n['soja_t_mtd'])} t", f"{fmt1(n['soja_moy_t_j'])} t/j", f"{fmt(n['soja_proj_t'])} t", f"{fmt(n['soja_obj_t'])} t", f"{n['soja_pct_obj']:.0f}%", status_text(n['soja_pct_obj'])],
    ["CONCENTRÉS", f"{fmt(n['conc_t_mtd'])} t", f"{fmt1(n['conc_moy_t_j'])} t/j", f"{fmt(n['conc_proj_t'])} t", f"{fmt(n['conc_obj_t'])} t", f"{n['conc_pct_obj']:.0f}%", status_text(n['conc_pct_obj'])],
    ["Total Soja+Conc", f"{fmt(n['soja_t_mtd']+n['conc_t_mtd'])} t", f"{fmt1(n['soja_moy_t_j']+n['conc_moy_t_j'])} t/j", f"{fmt(n['soja_proj_t']+n['conc_proj_t'])} t", f"{fmt(n['soja_obj_t']+n['conc_obj_t'])} t", f"{(n['soja_proj_t']+n['conc_proj_t'])/(n['soja_obj_t']+n['conc_obj_t'])*100:.0f}%", "—"],
    ["Ratio bundle", f"{n['ratio_global']:.1f}:1", "—", "—", "≤ 2,5:1", "—", "✅" if n['ratio_global'] <= 2.5 else "⚠"],
    ["Mix ventes", f"Soja {n['pct_soja_volume']:.0f}% / Conc {n['pct_conc_volume']:.0f}%", "—", "—", "—", "—", "—"],
]

# Build table with conditional coloring
t = Table(nat_data, colWidths=[3.5*cm, 2.2*cm, 1.6*cm, 2.7*cm, 1.8*cm, 1.4*cm, 1.2*cm], repeatRows=1)
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
story.append(t)
story.append(Spacer(1, 0.15*cm))

# Key insight for nationwide
nationwide_insight = f"""
<b>Synthèse nationwide</b> : Sur {MTD['days_elapsed']} jours ouvrés ({MTD['pct_elapsed']}% du mois), les ventes représentent <b>{fmt(n['soja_t_mtd']+n['conc_t_mtd'])} t</b> 
(soja + concentrés). La projection fin septembre s'établit à <b>{fmt(n['soja_proj_t']+n['conc_proj_t'])} t</b> 
({(n['soja_proj_t']+n['conc_proj_t'])/(n['soja_obj_t']+n['conc_obj_t'])*100:.0f}% de l'objectif combiné). 
Le <b>CONCENTRES est à {n['conc_pct_obj']:.0f}%</b> de l'objectif {'✅ sur trajectoire' if n['conc_pct_obj'] >= 80 else '⚠ sous objectif'}, mais le <b>SOJA est en retrait à {n['soja_pct_obj']:.0f}%</b> ❌ — reflet de la hausse tarifaire du 04/09 (26 000 FCFA/sac, +4% vs 25 000) 
qui ralentit temporairement la demande. Le <b>ratio bundle {n['ratio_global']:.1f}:1</b> reste excellent (objectif ≤ 2,5:1) — 
signe que le cross-sell se maintient malgré le choc prix.
"""
story.append(Paragraph(nationwide_insight, BODY))
story.append(Spacer(1, 0.2*cm))

# === 2. RÉGIONS ===
story.append(Paragraph("2. VUES RÉGIONALES — Performances par région", H2))

reg_data = [["Région", "Soja MTD (t)", "Soja proj (t)", "Soja obj (t)", "% Soja", "Conc MTD (t)", "Conc proj (t)", "Conc obj (t)", "% Conc", "Ratio"]]
total_soja_mtd = 0
total_conc_mtd = 0
total_soja_proj = 0
total_conc_proj = 0
total_soja_obj = 0
total_conc_obj = 0
for region in ['Ouest', 'Centre', 'Littoral']:
    r = PITCH['regions'][region]
    total_soja_mtd += r['soja_t_mtd']
    total_conc_mtd += r['conc_t_mtd']
    total_soja_proj += r['soja_proj_t']
    total_conc_proj += r['conc_proj_t']
    total_soja_obj += r['soja_obj_t']
    total_conc_obj += r['conc_obj_t']
    reg_data.append([
        region,
        fmt(r['soja_t_mtd']),
        fmt(r['soja_proj_t']),
        fmt(r['soja_obj_t']),
        f"{r['soja_pct_obj']:.0f}%",
        fmt(r['conc_t_mtd']),
        fmt(r['conc_proj_t']),
        fmt(r['conc_obj_t']),
        f"{r['conc_pct_obj']:.0f}%",
        f"{r['ratio']:.1f}:1",
    ])
reg_data.append([
    "TOTAL",
    fmt(total_soja_mtd),
    fmt(total_soja_proj),
    fmt(total_soja_obj),
    f"{total_soja_proj/total_soja_obj*100:.0f}%",
    fmt(total_conc_mtd),
    fmt(total_conc_proj),
    fmt(total_conc_obj),
    f"{total_conc_proj/total_conc_obj*100:.0f}%",
    f"{total_soja_mtd/total_conc_mtd:.1f}:1" if total_conc_mtd > 0 else "—",
])

t2 = Table(reg_data, colWidths=[1.8*cm, 1.7*cm, 1.7*cm, 1.5*cm, 1.2*cm, 1.7*cm, 1.7*cm, 1.5*cm, 1.2*cm, 1.2*cm], repeatRows=1)
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
    style_list2.append(('BACKGROUND', (4, i), (4, i), status_color(r['soja_pct_obj'])))
    style_list2.append(('BACKGROUND', (8, i), (8, i), status_color(r['conc_pct_obj'])))
# Total row
style_list2.append(('BACKGROUND', (0, 4), (-1, 4), colors.HexColor('#FFF2CC')))
style_list2.append(('FONT', (0, 4), (-1, 4), 'DejaVuSans-Bold', 8))
t2.setStyle(TableStyle(style_list2))
story.append(t2)
story.append(Spacer(1, 0.15*cm))

# Regional insight
ouest = PITCH['regions']['Ouest']
centre = PITCH['regions']['Centre']
littoral = PITCH['regions']['Littoral']
reg_insight = f"""
<b>Lecture régionale</b> : 
<b>OUEST</b> est la région leader ({fmt(ouest['soja_proj_t']+ouest['conc_proj_t'])} t projetés, {ouest['conc_pct_obj']:.0f}% obj conc ✅). 
<b>CENTRE</b> suit avec une performance CONCENTRES à {centre['conc_pct_obj']:.0f}% mais un SOJA en retrait à {centre['soja_pct_obj']:.0f}% ❌. 
<b>LITTORAL</b> présente le plus grand écart ({littoral['conc_pct_obj']:.0f}% conc, {littoral['soja_pct_obj']:.0f}% soja ❌) — 
NDOBO et VILLAGE particulièrement en retrait, à cibler en priorité pour les actions commerciales.
"""
story.append(Paragraph(reg_insight, BODY))
story.append(Spacer(1, 0.2*cm))

# === 3. INSIGHTS CLÉS ===
story.append(Paragraph("3. Insights clés", H2))

# Get bundle stats from MTD
b = MTD['bundle']
s = MTD['stock']
z = MTD['zero_achat']

insights_data = [
    [Paragraph("<b>Insight</b>", CELL), Paragraph("<b>Donnée</b>", CELL), Paragraph("<b>Implication</b>", CELL)],
    [Paragraph("Hausse prix soja 04/09", CELL),
     Paragraph(f"25 000 → 26 000 FCFA/sac (+4%)", CELL),
     Paragraph("Ralentissement temporaire de la demande soja — à surveiller sur les 2 prochaines semaines", CELL)],
    [Paragraph("Bundle ratio 1,8:1", CELL),
     Paragraph(f"357/370 cmds soja avec conc = {b['pct_bundle']}% cross-sell", CELL),
     Paragraph("✅ Maintien du cross-sell malgré la hausse prix — discipline commerciale préservée", CELL)],
    [Paragraph("🔴 Stock soja BEKOKO", CELL),
     Paragraph(f"<b>{fmt(s['brut_sacs'])} sacs</b> ({fmt(s['brut_t'])} t) + {fmt(s['production_en_cours_sacs'])} prod = {fmt(s['stock_avec_prod_sacs'])} sacs", CELL),
     Paragraph(f"<b>{s['jours_stock_brut']} jour(s) de stock central</b> — rupture <b>{s['rupture_date_brut']}</b>. Manque sept: {fmt(s['manque_septembre_sacs'])} sacs. RÉAPPRO URGENT", CELL)],
    [Paragraph("Top agence CONCENTRES", CELL),
     Paragraph("FAMLA (Ouest) — 66 t MTD, proj 341 t", CELL),
     Paragraph("22% du volume CONCENTRES national — pilier de la performance", CELL)],
    [Paragraph("Agences en retrait", CELL),
     Paragraph("NDOBO, VILLAGE, NKONGSAMBA (Littoral)", CELL),
     Paragraph("Littoral sous-performe — action commerciale ciblée requise", CELL)],
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
story.append(t3)
story.append(Spacer(1, 0.2*cm))

# === 4. RECOMMANDATIONS ===
story.append(Paragraph("4. Recommandations & Actions", H2))

recos = [
    ("ACTION IMMÉDIATE", "🔴 RÉAPPRO URGENT soja central",
     f"Stock BEKOKO = {fmt(s['brut_sacs'])} sacs ({fmt(s['brut_t'])} t) + {fmt(s['production_en_cours_sacs'])} prod = {fmt(s['stock_avec_prod_sacs'])} sacs. <b>{s['jours_stock_brut']} jour(s) de stock</b> — rupture {s['rupture_date_brut']}. Besoin sept = {fmt(s['besoin_septembre_sacs'])} sacs, <b>manque {fmt(s['manque_septembre_sacs'])} sacs</b>.",
     "Logistique / Direction Achats / DG", "AVANT 08/09/2026"),
    ("ACTION COMMERCIALE", "Push CONCENTRES sur Littoral",
     f"NDOBO + VILLAGE + NKONGSAMBA = {PITCH['regions']['Littoral']['conc_t_mtd']:.0f} t MTD vs projection {PITCH['regions']['Littoral']['conc_proj_t']:.0f} t (66% obj). Activer promotions bundle.",
     "RA Littoral + Direction Commerciale", "Semaine 38 (08-14/09)"),
    ("ACTION COMMERCIALE", "Maintenir le cross-sell bundle",
     f"Ratio 1,8:1 actuellement ✅ — poursuivre la discipline bundle (objectif ≤ 2,5:1) malgré la hausse prix. Push CONCENTRES en cross-sell sur les 13 cmds soja-only.",
     "Toutes agences", "Continu"),
    ("SURVEILLANCE", "Suivi cadence SOJA post-hausse 04/09",
     f"Soja à 59% objectif ❌ — surveiller la reprise sur 7-10 j. Si cadence < 80 t/j persistante, ajuster le forecast Q4 2026 (actuellement 20 051 t).",
     "Data Analyst + Direction Commerciale", "Point hebdo 12/09 + 19/09"),
    ("STRATÉGIQUE", "Ajuster prix forecast 2027",
     f"Le prix 26 000 FCFA/sac (vs 17 170 prévu Q4 2026 et 16 800 prévu 2027) dépasse les hypothèses forecast. Réviser les hypothèses prix si la hausse se confirme durable.",
     "Direction Financière + Data Analyst", "Décision d'ici 30/09"),
    ("OPPORTUNITÉ", "Réactivation 1 055 clients S1 sans achat Sept",
     f"76% des clients S1 n'ont pas encore acheté en Sept (mais mois entamé à 19%). Campagne téléphonique ciblée sur les top 200 clients S1 inactifs.",
     "Administrateurs de Vente", "Semaine 38-39"),
]

recos_data = [["Priorité", "Action", "Détail", "Responsable", "Échéance"]]
for r in recos:
    recos_data.append([r[0], r[1], r[2], r[3], r[4]])

t4 = Table(recos_data, colWidths=[2.5*cm, 3.2*cm, 6.3*cm, 3*cm, 2.5*cm], repeatRows=1)
style_list4 = [
    ('FONT', (0,0), (-1,0), 'DejaVuSans-Bold', 8),
    ('FONT', (0,1), (-1,-1), 'DejaVuSans', 8),
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
    ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
    ('TOPPADDING', (0,0), (-1,-1), 3), ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ('LEFTPADDING', (0,0), (-1,-1), 3), ('RIGHTPADDING', (0,0), (-1,-1), 3),
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
story.append(t4)
story.append(Spacer(1, 0.15*cm))

# === FOOTER ===
story.append(HRFlowable(width="100%", thickness=0.5, color=GRAY))
story.append(Spacer(1, 0.1*cm))
story.append(Paragraph(
    f"<b>Source</b> : {MTD['extraction_file']} (extraction au {MTD['update_date']}) — analyse au {MTD['update_date']} ({MTD['days_elapsed']}j/{MTD['total_days_sep']}j = {MTD['pct_elapsed']}% du mois). "
    f"<b>Méthodologie</b> : Volumes Livrées uniquement, clients internes (SPC/PDC/Comptoir) exclus. Projection fin septembre = moyenne quotidienne × 26 jours ouvrés (lun-sam). "
    f"<b>Hausse prix soja</b> : à partir du 04/09/2026, le prix du soja T102 (50kg) passe de 25 000 à 26 000 FCFA/sac (+4%). "
    f"<b>Stock central BEKOKO</b> : 2 716 sacs (135 t) + 296 sacs en production au 07/09/2026 (post-chargement agences weekend 05-06/09).",
    SMALL))

# === Save PDF ===
doc.build(story)
print(f"\n=== PITCH PDF GENERATED ===")
print(f"Path: {OUT}")
print(f"Size: {os.path.getsize(OUT)/1024:.0f} KB")
