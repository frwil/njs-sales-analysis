"""Generate September update PDF — combined analysis: zero-achat + bundle soja-concentrés + performance CONCENTRES.
Does NOT overwrite the original juillet/août PDFs. Saves as analyse_septembre_upd.pdf.

Based on sept_mtd_01.json (computed by compute_sept_mtd_metrics.py).
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
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, HRFlowable
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
H1 = ParagraphStyle('H1', parent=styles['Heading1'], fontName='DejaVuSans-Bold', fontSize=18, textColor=NAVY, spaceAfter=14, spaceBefore=10)
H2 = ParagraphStyle('H2', parent=styles['Heading2'], fontName='DejaVuSans-Bold', fontSize=14, textColor=NAVY, spaceAfter=10, spaceBefore=14)
H3 = ParagraphStyle('H3', parent=styles['Heading3'], fontName='DejaVuSans-Bold', fontSize=12, textColor=GOLD, spaceAfter=8, spaceBefore=10)
BODY = ParagraphStyle('Body', parent=styles['Normal'], fontName='DejaVuSans', fontSize=10, leading=14, alignment=TA_JUSTIFY, spaceAfter=8)
BULLET = ParagraphStyle('Bullet', parent=BODY, leftIndent=20, bulletIndent=10, spaceAfter=4)
SMALL = ParagraphStyle('Small', parent=BODY, fontSize=8, textColor=GRAY)
CELL = ParagraphStyle('Cell', parent=BODY, fontName='DejaVuSans', fontSize=9, leading=11, alignment=TA_LEFT, spaceAfter=0)

# === Load data ===
DATA = json.load(open('/home/z/my-project/scripts/sept_mtd_01.json'))
print(f"Loaded sept_mtd_01.json (update {DATA['update_date']})")

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
story.append(Paragraph(f"<b>Mise à jour</b> : {DATA['update_date']} ({DATA['days_elapsed']}/{DATA['total_days_sep']} jours, {DATA['pct_elapsed']}% du mois)", ParagraphStyle('CI', parent=BODY, fontSize=11, alignment=TA_CENTER, textColor=GRAY)))
story.append(Paragraph(f"<b>Source</b> : {DATA['extraction_file']}", ParagraphStyle('CI2', parent=BODY, fontSize=10, alignment=TA_CENTER, textColor=GRAY)))
story.append(Paragraph("William Francis Fohom — Data Analyst, Administrateur National de Ventes", ParagraphStyle('CI3', parent=BODY, fontSize=10, alignment=TA_CENTER, textColor=GRAY, spaceBefore=10)))

story.append(PageBreak())

# === Section 1: Performance CONCENTRES septembre ===
story.append(Paragraph("1. Performance CONCENTRES — Septembre 2026 MTD", H1))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.3*cm))

v_conc = DATA['volumes']['CONCENTRES']
v_soja = DATA['volumes']['TOURTEAUX']

# KPI box
story.append(Paragraph("<b>KPIs clés — CONCENTRES Septembre MTD</b>", H3))
kpi_data = [
    ["Indicateur", "Valeur MTD", "Projection fin Sept", "Objectif", "% Objectif", "Statut"],
    ["Volume CONCENTRES (t)", fmt(v_conc['t']), fmt(v_conc['proj_t']), fmt(v_conc['obj_t']), fmt(v_conc['pct_obj']) + "%",
     "✅ Sur trajectoire" if v_conc['pct_obj'] >= 80 else ("⚠ À surveiller" if v_conc['pct_obj'] >= 60 else "❌ Critique")],
    ["Moyenne quotidienne (t/j)", fmt1(v_conc['moy_t_j']), "—", fmt1(v_conc['obj_t']/DATA['total_days_sep']), "—",
     "✅ Au-dessus objectif" if v_conc['moy_t_j'] >= v_conc['obj_t']/DATA['total_days_sep'] else "⚠ Sous objectif"],
    ["Volume TOURTEAUX (t)", fmt(v_soja['t']), fmt(v_soja['proj_t']), fmt(v_soja['obj_t']), fmt(v_soja['pct_obj']) + "%",
     "✅ Sur trajectoire" if v_soja['pct_obj'] >= 80 else ("⚠ À surveiller" if v_soja['pct_obj'] >= 60 else "❌ Critique")],
]
story.append(make_table(kpi_data, col_widths=[4*cm, 2.5*cm, 3*cm, 2*cm, 2*cm, 3.5*cm], font_size=9, highlight_rows=[]))
story.append(Spacer(1, 0.3*cm))

# Comparison avec août
import os
aout_path = '/home/z/my-project/scripts/aout_mtd_30.json'
if os.path.exists(aout_path):
    AOUT = json.load(open(aout_path))
    story.append(Paragraph("<b>Comparaison Septembre MTD vs Août complet</b>", H3))
    a_conc = AOUT['volumes']['CONCENTRES']
    a_soja = AOUT['volumes']['TOURTEAUX']
    
    var_conc_t = ((v_conc['proj_t'] / a_conc['t']) - 1) * 100 if a_conc['t'] > 0 else 0
    var_soja_t = ((v_soja['proj_t'] / a_soja['t']) - 1) * 100 if a_soja['t'] > 0 else 0
    
    var_conc_moy = ((v_conc['moy_t_j'] / a_conc['moy_t_j']) - 1) * 100 if a_conc['moy_t_j'] > 0 else 0
    var_soja_moy = ((v_soja['moy_t_j'] / a_soja['moy_t_j']) - 1) * 100 if a_soja['moy_t_j'] > 0 else 0
    
    comp_data = [
        ["Indicateur", "Août complet", "Sept MTD (proj.)", "Variation", "Lecture"],
        ["CONCENTRES (t)", fmt(a_conc['t']), fmt(v_conc['proj_t']), f"{'+' if var_conc_t>=0 else ''}{var_conc_t:.1f}%",
         "Maintien ou progression" if var_conc_t >= -5 else "Recul significatif"],
        ["CONCENTRES moy/j (t/j)", fmt1(a_conc['moy_t_j']), fmt1(v_conc['moy_t_j']), f"{'+' if var_conc_moy>=0 else ''}{var_conc_moy:.1f}%",
         "Cadence préservée" if var_conc_moy >= -5 else "Ralentissement"],
        ["TOURTEAUX (t)", fmt(a_soja['t']), fmt(v_soja['proj_t']), f"{'+' if var_soja_t>=0 else ''}{var_soja_t:.1f}%",
         "Stable" if abs(var_soja_t) < 10 else ("Hausse" if var_soja_t > 0 else "Baisse")],
        ["TOURTEAUX moy/j (t/j)", fmt1(a_soja['moy_t_j']), fmt1(v_soja['moy_t_j']), f"{'+' if var_soja_moy>=0 else ''}{var_soja_moy:.1f}%",
         "Stable" if abs(var_soja_moy) < 10 else ("Hausse" if var_soja_moy > 0 else "Baisse")],
    ]
    story.append(make_table(comp_data, col_widths=[4*cm, 2.5*cm, 3*cm, 2.5*cm, 4.5*cm], font_size=9, highlight_rows=[]))
    story.append(Spacer(1, 0.3*cm))

# CONCENTRES par agence
story.append(Paragraph("<b>CONCENTRES par agence — Septembre MTD + projection</b>", H3))
ag_data = [["Rang", "Agence", "Conc MTD (t)", "Soja MTD (t)", "Moy/j (t)", "Proj fin Sept (t)"]]
for i, ag in enumerate(DATA['conc_by_agence'][:10], 1):
    ag_data.append([str(i), ag['agence'], fmt(ag['conc_t']), fmt(ag['soja_t']), fmt1(ag['moy_t_j']), fmt(ag['proj_t'])])
# Total row
total_conc_mtd = sum(a['conc_t'] for a in DATA['conc_by_agence'])
total_soja_mtd = sum(a['soja_t'] for a in DATA['conc_by_agence'])
total_proj = sum(a['proj_t'] for a in DATA['conc_by_agence'])
ag_data.append(["", "TOTAL (10 agences)", fmt(total_conc_mtd), fmt(total_soja_mtd),
               fmt1(total_conc_mtd/DATA['days_elapsed']), fmt(total_proj)])
story.append(make_table(ag_data, col_widths=[1.2*cm, 3.5*cm, 2.8*cm, 2.8*cm, 2.2*cm, 3.2*cm], font_size=9, highlight_rows=[len(ag_data)-1]))
story.append(Spacer(1, 0.3*cm))

# Analyse & lecture
story.append(Paragraph("<b>Analyse & Lecture</b>", H3))
analysis_text = f"""
<b>Performance CONCENTRES Septembre MTD (au {DATA['update_date']})</b> : 
{v_conc['t']} t écoulées en {DATA['days_elapsed']} jours ouvrés, soit une cadence de 
<b>{v_conc['moy_t_j']} t/jour</b>. Sur cette base, la projection fin septembre s'établit à 
<b>{v_conc['proj_t']} t</b> contre un objectif de {v_conc['obj_t']} t, soit 
<b>{v_conc['pct_obj']}% de l'objectif</b>. 
"""
if v_conc['pct_obj'] >= 100:
    analysis_text += "✅ <b>Objectif en passe d'être dépassé</b> — cadence supérieure à la cible. "
elif v_conc['pct_obj'] >= 80:
    analysis_text += "✅ <b>Sur trajectoire</b> — l'objectif est atteignable en fin de mois. "
elif v_conc['pct_obj'] >= 60:
    analysis_text += "⚠ <b>À surveiller</b> — risque de sous-performance si la cadence ne s'améliore pas. "
else:
    analysis_text += "❌ <b>Critique</b> — cadence insuffisante, action commerciale urgente requise. "

if os.path.exists(aout_path):
    a_conc = AOUT['volumes']['CONCENTRES']
    var = ((v_conc['proj_t']/a_conc['t'])-1)*100
    if var >= 0:
        analysis_text += f"La projection septembre ({v_conc['proj_t']} t) est <b>supérieure de {var:.1f}%</b> au volume août complet ({a_conc['t']} t) — <b>continuité de la dynamique</b> observée en août."
    else:
        analysis_text += f"La projection septembre ({v_conc['proj_t']} t) est <b>inférieure de {abs(var):.1f}%</b> au volume août complet ({a_conc['t']} t) — <b>ralentissement à surveiller</b>."

story.append(Paragraph(analysis_text, BODY))
story.append(Spacer(1, 0.3*cm))

story.append(PageBreak())

# === Section 2: Bundle Soja-Concentrés ===
story.append(Paragraph("2. Bundle Soja-Concentrés — Septembre MTD", H1))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.3*cm))

b = DATA['bundle']
bundle_data = [
    ["Indicateur", "Valeur Sept MTD", "Comparaison"],
    ["Sacs soja vendus", fmt(b['total_soja_sacs']), f"≈ {b['total_soja_sacs']*50/1000:.0f} t équivalent"],
    ["Sacs concentrés vendus", fmt(b['total_conc_sacs']), f"≈ {b['total_conc_sacs']*50/1000:.0f} t équivalent"],
    ["Ratio global soja/conc", f"{b['ratio']}:1", "✅ Excellent (objectif ≤ 2,5:1)" if b['ratio'] <= 2.5 else "⚠ Au-dessus objectif"],
    ["Commandes soja", str(b['cmds_soja']), "—"],
    ["Commandes bundle (soja+conc)", str(b['cmds_bundle']), f"{b['pct_bundle']}% des commandes soja"],
    ["Commandes soja-only (sans conc)", str(b['cmds_soja_only']), "Cross-sell à améliorer" if b['cmds_soja_only'] > 50 else "✅ Bon cross-sell"],
    ["Commandes conc-only (sans soja)", str(b['cmds_conc_only']), "—"],
]
story.append(make_table(bundle_data, col_widths=[5*cm, 4*cm, 7*cm], font_size=9))
story.append(Spacer(1, 0.3*cm))

# Distribution des ratios
story.append(Paragraph("<b>Distribution des ratios bundle (commandes avec soja + conc)</b>", H3))
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
story.append(make_table(dist_data, col_widths=[5*cm, 5*cm, 6*cm], font_size=9, highlight_rows=[6]))
story.append(Spacer(1, 0.3*cm))

bundle_analysis = f"""
<b>Lecture Bundle Septembre MTD</b> : 
Le ratio global soja/concentrés est de <b>{b['ratio']}:1</b> — 
{"<b>inférieur à l'objectif 2,5:1</b>, ce qui témoigne d'un excellent cross-sell" if b['ratio'] <= 2.5 else "<b>au-dessus de l'objectif 2,5:1</b>, à surveiller"}. 
Sur {b['cmds_soja']} commandes soja, <b>{b['cmds_bundle']} ({b['pct_bundle']}%) incluent des concentrés</b> — 
{'niveau de cross-sell excellent' if b['pct_bundle'] >= 90 else ('cross-sell solide' if b['pct_bundle'] >= 75 else 'cross-sell à améliorer')}. 
Seulement {b['cmds_soja_only']} commandes soja-only ({b['cmds_soja_only']/b['cmds_soja']*100:.1f}% des commandes soja) — 
{'✅ aligné avec le plan d\'action bundle' if b['cmds_soja_only']/b['cmds_soja']*100 <= 10 else '⚠ taux de soja-only élevé'}.
"""
story.append(Paragraph(bundle_analysis, BODY))
story.append(Spacer(1, 0.3*cm))

story.append(PageBreak())

# === Section 3: Stock Soja ===
story.append(Paragraph("3. Stock Soja BEKOKO — Septembre MTD", H1))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.3*cm))

s = DATA['stock']
stock_data = [
    ["Indicateur", "Valeur", "Détail"],
    ["Date de référence stock", s['date'], "Dernier inventaire connu"],
    ["Stock brut (sacs 50kg éq.)", fmt(s['brut_sacs']), f"{s['brut_t']} t"],
    ["Stock net (hors SPC)", fmt(s['net_sacs']), f"{s['net_t']} t (SPC exclu: {s['spc_exclu']} sacs)"],
    ["Vente soja septembre (sacs/jour)", fmt(s['vente_sacs_jour']), f"≈ {s['vente_sacs_jour']*50/1000:.0f} t/jour"],
    ["Vente soja semaine (sacs/sem)", fmt(s['vente_sacs_sem']), "Projection sur 6 jours ouvrés"],
    ["Jours de stock restants", f"{s['jours_stock']} jours", "Au rythme de vente actuel"],
    ["Date rupture probable", s['rupture_date'], "À programmer réappro avant cette date"],
]
story.append(make_table(stock_data, col_widths=[5*cm, 4*cm, 7*cm], font_size=9, highlight_rows=[7]))
story.append(Spacer(1, 0.3*cm))

stock_analysis = f"""
<b>Lecture Stock</b> : Au rythme de vente actuel ({s['vente_sacs_jour']} sacs/jour), le stock net de {s['net_sacs']} sacs 
couvre <b>{s['jours_stock']} jours</b> — rupture probable estimée au <b>{s['rupture_date']}</b>. 
{'✅ Stock confortable, pas d\'urgence immédiate' if s['jours_stock'] >= 30 else '⚠ Stock à surveiller, réappro à programmer'}. 
La hausse du rythme de vente septembre ({s['vente_sacs_jour']} sacs/j vs ~2 914 sacs/j en août) reflète la saisonnalité Q4 et confirme la trajectoire haute du forecast.
"""
story.append(Paragraph(stock_analysis, BODY))
story.append(Spacer(1, 0.3*cm))

# === Section 4: Zero-Achat & Cross-sell ===
story.append(Paragraph("4. Zero-Achat & Cross-Sell — Septembre MTD", H1))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.3*cm))

z = DATA['zero_achat']
c = DATA['cross_sell']

za_data = [
    ["Indicateur", "Valeur", "Lecture"],
    ["Clients actifs S1 2026", str(z['s1_clients']), "Base de référence (Jan-Juin 2026)"],
    ["Clients actifs Sept MTD", str(z['sept_clients']), f"{z['sept_clients']/z['s1_clients']*100:.1f}% de la base S1"],
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
story.append(make_table(za_data, col_widths=[5*cm, 3*cm, 8*cm], font_size=9))
story.append(Spacer(1, 0.3*cm))

za_analysis = f"""
<b>Lecture Zero-Achat</b> : Sur les {z['s1_clients']} clients actifs en S1 2026, seuls <b>{z['sept_clients']} 
({z['sept_clients']/z['s1_clients']*100:.1f}%) ont déjà acheté en septembre MTD</b>. 
Toutefois, le mois n'est entamé que depuis {DATA['days_elapsed']} jours ({DATA['pct_elapsed']}% du mois) — ce chiffre va 
naturellement monter jusqu'en fin septembre. <b>{z['new_sept']} nouveaux clients</b> (hors S1) ont déjà acheté en septembre, 
ce qui dénote une dynamique commerciale positive. 

<b>Cross-sell</b> : Sur {c['s1_soja']} clients S1 soja, <b>{c['soja_to_conc_sept']} ({c['pct_soja_with_conc_sept']}%) 
ont aussi acheté du concentré en septembre</b>. Ce taux est plus faible que la performance bundle observée 
({b['pct_bundle']}% des commandes), ce qui indique que le cross-sell est meilleur sur les nouveaux clients 
septembre que sur les clients S1 historiques — opportunité de réactivation commerciale sur les S1 soja-only.
"""
story.append(Paragraph(za_analysis, BODY))
story.append(Spacer(1, 0.3*cm))

# === Section 5: Recommandations ===
story.append(Paragraph("5. Recommandations & Plan d'action", H1))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.3*cm))

recos = []
status_conc = "✅" if v_conc['pct_obj'] >= 80 else ("⚠" if v_conc['pct_obj'] >= 60 else "❌")
recos.append(f"{status_conc} <b>CONCENTRES — Cadence {v_conc['moy_t_j']} t/j</b> vs objectif {v_conc['obj_t']/DATA['total_days_sep']:.1f} t/j. "
            f"Projection {v_conc['proj_t']} t vs obj {v_conc['obj_t']} t = {v_conc['pct_obj']}%. "
            + ("Maintenir la cadence, objectif atteignable." if v_conc['pct_obj'] >= 80 else "Accélérer la cadence — action commerciale sur les agences en retard."))

status_bundle = "✅" if b['ratio'] <= 2.5 else "⚠"
recos.append(f"{status_bundle} <b>Bundle ratio {b['ratio']}:1</b> — "
            + ("Excellent, sous l'objectif 2,5:1. Cross-sell à 96% — maintenir." if b['ratio'] <= 2.5 else "Au-dessus de l'objectif 2,5:1. Renforcer le push concentrés."))

recos.append(f"📦 <b>Stock soja</b> : {s['net_sacs']} sacs au {s['date']} — {s['jours_stock']} jours de stock, rupture probable {s['rupture_date']}. "
            + ("Pas d'urgence immédiate." if s['jours_stock'] >= 30 else "Programmer le réapprovisionnement."))

recos.append(f"🎯 <b>Cross-sell S1 soja → conc</b> : {c['pct_soja_with_conc_sept']}% des S1 soja ont pris du conc en sept. "
            f"Action de réactivation sur les {c['s1_soja'] - c['soja_to_conc_sept']} clients S1 soja sans conc en sept.")

# Top agence to push
if DATA['conc_by_agence']:
    top_ag = DATA['conc_by_agence'][0]
    recos.append(f"🏆 <b>Top agence Sept</b> : {top_ag['agence']} ({top_ag['conc_t']} t, proj {top_ag['proj_t']} t) — "
                "utiliser comme modèle de performance pour les autres agences.")

# Bottom agence to push
if len(DATA['conc_by_agence']) > 5:
    bottom_ag = DATA['conc_by_agence'][-1]
    recos.append(f"⚠ <b>Agence à surveiller</b> : {bottom_ag['agence']} ({bottom_ag['conc_t']} t, proj {bottom_ag['proj_t']} t) — "
                "renforcer le support commercial.")

for r in recos:
    story.append(Paragraph(f"• {r}", BULLET))

story.append(Spacer(1, 0.5*cm))
story.append(Paragraph("<b>Note méthodologique</b>", H3))
story.append(Paragraph(
    f"Les données de cette analyse sont issues de l'extraction <b>{DATA['extraction_file']}</b> "
    f"(mise à jour au {DATA['update_date']}). Les volumes MTD ne couvrent que {DATA['days_elapsed']} jours ouvrés sur {DATA['total_days_sep']} "
    f"({DATA['pct_elapsed']}% du mois). Les projections fin septembre sont calculées en multipliant la moyenne quotidienne "
    f"par le nombre total de jours ouvrés (lundi-samedi). Les comparaisons avec août utilisent l'extraction (27).xlsx du 31/08/2026. "
    "Les clients internes (SPC, PDC, Comptoir, Emana) sont exclus des analyses bundle et zero-achat. "
    "Les chiffres préliminaires peuvent évoluer avec les prochaines extractions ERP.",
    SMALL))

# === Save PDF ===
doc.build(story)
print(f"\n=== PDF GENERATED ===")
print(f"Path: {OUT}")
print(f"Size: {os.path.getsize(OUT)/1024:.0f} KB")
