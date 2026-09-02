"""
Forecast 2027 - Génération des 5 PDFs PACE en un seul script optimisé.
Résumé exécutif + Proposition + RACI + Stratégie + Guide.
"""
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle,
    Image, HRFlowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY

pdfmetrics.registerFont(TTFont('DejaVuSans', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))

NAVY = HexColor('#1F4E78'); GOLD = HexColor('#C9A961'); GRAY = HexColor('#595959')
LIGHT_GRAY = HexColor('#F2F2F2'); GREEN = HexColor('#C6EFCE'); S3_COLOR = 'C6EFCE'
from reportlab.lib import colors as rl_colors
from reportlab.platypus import TableStyle
THIN_SIDE = __import__('reportlab.lib.colors', fromlist=['Color']).Color  # not used directly
# Use TableStyle directly for borders - no need for Side/Border objects

styles = getSampleStyleSheet()
H1 = ParagraphStyle('H1', parent=styles['Heading1'], fontName='DejaVuSans-Bold', fontSize=18, textColor=NAVY, spaceAfter=14, spaceBefore=10)
H2 = ParagraphStyle('H2', parent=styles['Heading2'], fontName='DejaVuSans-Bold', fontSize=14, textColor=NAVY, spaceAfter=10, spaceBefore=14)
H3 = ParagraphStyle('H3', parent=styles['Heading3'], fontName='DejaVuSans-Bold', fontSize=12, textColor=GOLD, spaceAfter=8, spaceBefore=10)
BODY = ParagraphStyle('Body', parent=styles['Normal'], fontName='DejaVuSans', fontSize=10, leading=14, alignment=TA_JUSTIFY, spaceAfter=8)
BODY_BOLD = ParagraphStyle('BodyBold', parent=BODY, fontName='DejaVuSans-Bold')
BULLET = ParagraphStyle('Bullet', parent=BODY, leftIndent=20, bulletIndent=10, spaceAfter=4)
SMALL = ParagraphStyle('Small', parent=BODY, fontSize=8, textColor=GRAY)
CAPTION = ParagraphStyle('Caption', parent=BODY, fontSize=9, textColor=GRAY, alignment=TA_CENTER)

CELL_STYLE = ParagraphStyle('CellStyle', parent=BODY, fontName='DejaVuSans', fontSize=9, leading=11, alignment=TA_LEFT, spaceAfter=0, spaceBefore=0)

OUT_DIR = "/home/z/my-project/download/forecast_2027"
os.makedirs(OUT_DIR, exist_ok=True)

def make_table(data, col_widths=None, font_size=9, header_color=NAVY):
    cell_style = ParagraphStyle('CD', parent=CELL_STYLE, fontSize=font_size, leading=font_size+2)
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
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), header_color),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_GRAY]),
        ('TOPPADDING', (0,0), (-1,-1), 4), ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4), ('RIGHTPADDING', (0,0), (-1,-1), 4),
    ]))
    return t

def cover_page(title, subtitle, doc_type, date_str="2 septembre 2026"):
    elements = []
    elements.append(Spacer(1, 4*cm))
    elements.append(Paragraph("BELGOCAM SA", ParagraphStyle('CL', parent=styles['Title'], fontName='DejaVuSans-Bold', fontSize=32, textColor=NAVY, alignment=TA_CENTER)))
    elements.append(Spacer(1, 0.5*cm))
    elements.append(HRFlowable(width="60%", thickness=2, color=GOLD, spaceBefore=10, spaceAfter=20, hAlign='CENTER'))
    elements.append(Spacer(1, 1*cm))
    elements.append(Paragraph(title, ParagraphStyle('CT', parent=styles['Title'], fontName='DejaVuSans-Bold', fontSize=28, textColor=NAVY, alignment=TA_CENTER, spaceAfter=20)))
    elements.append(Paragraph(subtitle, ParagraphStyle('CS', parent=styles['Title'], fontName='DejaVuSans', fontSize=16, textColor=GOLD, alignment=TA_CENTER, spaceAfter=30)))
    elements.append(Spacer(1, 2*cm))
    elements.append(HRFlowable(width="40%", thickness=1, color=GRAY, spaceBefore=10, spaceAfter=10, hAlign='CENTER'))
    elements.append(Paragraph(f"<b>{doc_type}</b>", ParagraphStyle('CI', parent=BODY, fontSize=11, alignment=TA_CENTER, textColor=GRAY)))
    elements.append(Paragraph("William Francis Fohom", ParagraphStyle('CI2', parent=BODY, fontSize=11, alignment=TA_CENTER, textColor=GRAY)))
    elements.append(Paragraph("Data Analyst | Administrateur National de Ventes", ParagraphStyle('CI3', parent=BODY, fontSize=11, alignment=TA_CENTER, textColor=GRAY)))
    elements.append(Spacer(1, 0.5*cm))
    elements.append(Paragraph(date_str, ParagraphStyle('CI4', parent=BODY, fontSize=11, alignment=TA_CENTER, textColor=GRAY)))
    elements.append(PageBreak())
    return elements

def style_header_row(ws, row, n):
    pass  # Not needed for PDF

# ==================== 1. RÉSUMÉ EXÉCUTIF ====================
print("Generating 1. Résumé exécutif 2027...")
exec_path = f"{OUT_DIR}/01_resume_executif.pdf"
doc = SimpleDocTemplate(exec_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []

story.append(Paragraph("RÉSUMÉ EXÉCUTIF", H1))
story.append(Paragraph("Forecast 2027 - Volumes et Valeurs (12 mois)", H2))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Contexte</b>", H3))
story.append(Paragraph(
    "BELGOCAM SA doit anticiper ses volumes de ventes et son chiffre d'affaires pour l'année 2027 complète "
    "(janvier - décembre, 12 mois). Ce forecast annuel s'appuie sur <b>115 592 enregistrements</b> couvrant "
    "20 mois d'historique (janvier 2025 - août 2026), intégrant les commandes En cours et Validées d'août 2026. "
    "L'effet soja exceptionnel de juillet-août 2026 (rupture concurrente) a été <b>désaisonnalisé</b> pour éviter "
    "de biaiser les prédictions. Le prix du soja a été actualisé à <b>25 000 FCFA/sac</b> (hausse du 24/08/2026).",
    BODY))

story.append(Paragraph("<b>Méthodologie</b>", H3))
story.append(Paragraph(
    "Le modèle <b>Prophet</b> (Facebook/Meta) a été entraîné sur 4 familles alimentaires (TOURTEAUX, CONCENTRÉS, "
    "INGRÉDIENTS, ALIMENT COMPLET) au niveau famille × région. Le MATERIEL ÉLEVAGE et les PREMIX sont projetés "
    "par extrapolation de la moyenne historique. Le forecast couvre <b>6 familles, 69 produits, 14 agences</b> et "
    "3 régions. Le scénario S3 (réappro soja 100%, situation normale) est utilisé comme référence.",
    BODY))

story.append(Paragraph("<b>Résultats clés 2027</b>", H3))
synth_data = [
    ["Indicateur", "Valeur", "Détail"],
    ["Volume total 2027", "109 288 t", "6 familles, 69 produits, 14 agences"],
    ["CA total 2027", "58 550 M FCFA", "Prix soja 25 000 FCFA/sac"],
    ["Période", "12 mois (Jan-Déc 2027)", "Forecast complet annuel"],
    ["Scénario", "S3 (réappro 100%)", "Situation normale"],
    ["Données historiques", "115 592 enregistrements", "Jan 2025 - Août 2026 + En cours/Validées"],
    ["Désaisonnalisation", "Effet soja Jul-Août 2026 neutralisé", "Cap à moyenne S1"],
]
story.append(make_table(synth_data, col_widths=[5*cm, 4*cm, 8*cm], font_size=9))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Synthèse par famille</b>", H3))
fam_data = [
    ["Famille", "Volume (t)", "CA (M FCFA)", "Part CA"],
    ["TOURTEAUX", "87 399", "43 701", "74,6%"],
    ["CONCENTRÉS", "19 565", "13 063", "22,3%"],
    ["ALIMENT_COMPLET", "1 327", "1 038", "1,8%"],
    ["MATERIEL_ELEVAGE", "—", "376", "0,6%"],
    ["INGREDIENTS", "996", "209", "0,4%"],
    ["PREMIX", "—", "163", "0,3%"],
    ["MAÏS (exclu)", "0", "0", "0,0%"],
    ["TOTAL", "109 288", "58 550", "100%"],
]
story.append(make_table(fam_data, col_widths=[4*cm, 3.5*cm, 3.5*cm, 2.5*cm], font_size=9))
story.append(Paragraph("<i>Note: Le maïs (M1051) a été retiré des ventes 2026 avant le forecast 2027 car produit opportuniste hors portefeuille régulier BELGOCAM. Ligne MAÏS à 0 pour traçabilité.</i>", SMALL))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Synthèse par trimestre</b>", H3))
q_data = [
    ["Trimestre", "Volume (t)", "CA (M FCFA)", "Part CA"],
    ["Q1 (Jan-Mar)", "19 309", "10 599", "18,1%"],
    ["Q2 (Avr-Juin)", "29 614", "15 673", "26,8%"],
    ["Q3 (Juil-Sept)", "10 950", "6 427", "11,0%"],
    ["Q4 (Oct-Déc)", "49 414", "25 851", "44,1%"],
    ["TOTAL", "109 288", "58 550", "100%"],
]
story.append(make_table(q_data, col_widths=[4*cm, 3.5*cm, 3.5*cm, 2.5*cm], font_size=9))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Top 5 agences par CA 2027</b>", H3))
top_data = [
    ["Rang", "Agence", "Région", "CA (M FCFA)"],
    ["1", "FAMLA", "Ouest", "14 196"],
    ["2", "NDOBO", "Littoral", "9 715"],
    ["3", "MESSASSI", "Centre", "5 779"],
    ["4", "DJELENG", "Ouest", "4 029"],
    ["5", "VILLAGE", "Littoral", "3 800"],
]
story.append(make_table(top_data, col_widths=[1.5*cm, 4*cm, 3*cm, 4.5*cm], font_size=9))

story.append(PageBreak())


story.append(Paragraph("<b>Réalisation 2026 YTD + Projection fin d'année (contexte)</b>", H3))

story.append(Paragraph(
    "Le forecast 2027 s'appuie sur la réalisation 2026. Le tableau ci-dessous présente la trajectoire 2026 "
    "(YTD réel Jan-Août + Q4 forecast S3) vs objectif annuel, qui sert de base de comparaison.",
    BODY))

ytd_data = [
    ["Famille", "YTD 2026 (t)", "Q4 fcst (t)", "Total 2026 (t)", "Obj annuel (t)", "% Annuel", "Statut"],
    ["TOURTEAUX", "39 718", "33 853", "73 571", "54 196", "136%", "✅"],
    ["CONCENTRÉS", "12 066", "6 393", "18 459", "24 192", "76%", "❌"],
    ["ALIMENT COMPLET", "525", "368", "893", "1 046", "85%", "❌"],
    ["INGRÉDIENTS", "548", "307", "855", "1 064", "80%", "❌"],
    ["PREMIX", "61", "0", "61", "130", "47%", "❌"],
    ["MAÏS (exclu)", "0", "0", "0", "—", "—", "—"],
    ["TOTAL", "52 918", "40 921", "93 839", "80 628", "116%", "✅"],
]
story.append(make_table(ytd_data, col_widths=[2.8*cm, 2*cm, 1.8*cm, 2.2*cm, 2.5*cm, 1.5*cm, 1*cm], font_size=8))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph(
    "<b>Lecture</b> : En 2026, le TOURTEAUX surperforme à 136% (effet rupture concurrente exceptionnel) "
    "tandis que les CONCENTRÉS restent à 76% de l'objectif. Le forecast 2027 projette une normalisation : "
    "109 288 t (vs 93 839 t en 2026, +16%) et 58 550 M FCFA. La désaisonnalisation de l'effet soja 2026 "
    "évite de répliquer ce phénomène non récurrent dans les prédictions 2027.",
    BODY))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Recommandations principales</b>", H3))
recos = [
    "<b>Planification annuelle 2027</b> — Utiliser le forecast S3 (109 288 t, 58 550 M FCFA) comme base budgétaire.",
    "<b>Saisonnalité Q4</b> — Le Q4 2027 représente 44% du CA annuel (25 851 M FCFA). Préparer les stocks et les ressources dès septembre 2027.",
    "<b>Pic d'octobre</b> — Octobre 2027 = pic de l'année (20 602 t, 10 746 M FCFA). Anticiper le réapprovisionnement soja avant septembre 2027.",
    "<b>Maintien du bundle</b> — Le ratio bundle 2,3:1 atteint en août 2026 doit être maintenu en 2027 pour optimiser les CONCENTRÉS.",
    "<b>Mise à jour trimestrielle</b> — Actualiser le forecast chaque trimestre avec les nouvelles données ERP.",
    "<b>Surveillance FAMLA et NDOBO</b> — Ces 2 agences représentent 41% du CA 2027, leur performance est critique.",
]
for r in recos:
    story.append(Paragraph(f"• {r}", BULLET))

story.append(Spacer(1, 0.3*cm))
story.append(Paragraph("<b>Livrables produits</b>", H3))
livrables = [
    "Excel forecast 2027 (7 feuilles: synthèse, par famille × mois, par région × mois, par agence, par produit, détail complet, hypothèses)",
    "Résumé exécutif PDF (ce document)",
    "Proposition de projet PDF",
    "Matrice RACI PDF",
    "Document stratégique PACE PDF",
    "Guide méthodologique PDF",
]
for l in livrables:
    story.append(Paragraph(f"• {l}", BULLET))

story.append(Spacer(1, 0.3*cm))
story.append(Paragraph("<b>Conclusion</b>", H3))
story.append(Paragraph(
    "Le forecast 2027 projette <b>109 288 tonnes</b> pour un CA de <b>58 550 M FCFA</b>, avec une forte concentration "
    "saisonnière au Q4 (44% du CA). La désaisonnalisation de l'effet soja exceptionnel de 2026 et l'actualisation "
    "du prix à 25 000 FCFA/sac permettent une projection réaliste. Le pic d'octobre (20 602 t) nécessitera une "
    "anticipation renforcée du réapprovisionnement soja dès septembre 2027.",
    BODY))

doc.build(story)
print(f"✓ Résumé exécutif: {exec_path} ({os.path.getsize(exec_path)/1024:.0f} KB)")

# ==================== 2. PROPOSITION DE PROJET ====================
print("Generating 2. Proposition de projet 2027...")
prop_path = f"{OUT_DIR}/02_proposition_projet.pdf"
doc = SimpleDocTemplate(prop_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []

story.extend(cover_page("Proposition de Projet", "Forecast 2027 - Volumes et Valeurs (12 mois)", "PROPOSITION DE PROJET"))

story.append(Paragraph("1. Contexte et justification", H1))
story.append(Paragraph(
    "BELGOCAM SA doit anticiper ses volumes et CA pour l'année 2027 complète. Ce forecast annuel s'appuie sur "
    "20 mois d'historique (janvier 2025 - août 2026) et intègre les innovations suivantes : désaisonnalisation "
    "de l'effet soja exceptionnel, inclusion des commandes En cours/Validées, actualisation du prix soja à "
    "25 000 FCFA/sac. Le forecast couvre 6 familles (sans Maïs), 69 produits et 14 agences.",
    BODY))

story.append(Paragraph("2. Objectifs", H1))
story.append(Paragraph("<b>Objectif principal</b>", H3))
story.append(Paragraph(
    "Produire un forecast annuel 2027 (12 mois) en volumes et valeurs, désagrégé par produit, famille, "
    "agence et région, selon le scénario S3 (réappro soja 100%, situation normale).",
    BODY))
story.append(Paragraph("<b>Objectifs spécifiques</b>", H3))
objs = [
    "Modéliser la tendance et la saisonnalité mensuelles via Prophet (4 familles alimentaires)",
    "Désaisonnaliser l'effet soja exceptionnel de juillet-août 2026 pour éviter le biais",
    "Intégrer les commandes En cours et Validées comme potentielles ventes Livrées",
    "Actualiser le prix soja à 25 000 FCFA/sac (hausse du 24/08/2026)",
    "Extrapoler le MATERIEL_ELEVAGE et les PREMIX par moyenne historique",
    "Exclure le Maïs et les produits opportunistes (ELVOR TONIC, CARBONATE DE CALCIUM)",
    "Produire les livrables PACE complets (guide, stratégie, RACI, résumé, Excel)",
]
for o in objs:
    story.append(Paragraph(f"• {o}", BULLET))

story.append(Paragraph("3. Périmètre", H1))
scope_data = [
    ["Dimension", "Détail", "Volume"],
    ["Période forecast", "Janvier - Décembre 2027 (12 mois)", "12 mois"],
    ["Familles incluses", "TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, ALIMENT COMPLET, MATERIEL ÉLEVAGE, PREMIX", "6 familles"],
    ["Produits", "69 références (T102, C101-C108, MAT014/MAT011/MAT017 alvéoles, P102N2, etc.)", "69 produits"],
    ["Agences", "14 agences BELGOCAM (FAMLA, NDOBO, DJELENG, etc.)", "14 agences"],
    ["Régions", "Ouest, Centre, Littoral", "3 régions"],
    ["Niveau détail", "Produit × Agence × Mois", "7 224 lignes"],
    ["Scénario", "S3 - Réappro soja 100% (situation normale)", "1 scénario"],
    ["Données historiques", "115 592 enregistrements (Jan 2025 - Août 2026) + En cours/Validées", "20 mois"],
    ["Désaisonnalisation", "Effet soja Jul-Août 2026 neutralisé (cap moyenne S1)", "—"],
    ["Prix soja", "25 000 FCFA/sac (actualisé au 24/08/2026)", "—"],
]
story.append(make_table(scope_data, col_widths=[3.5*cm, 9.5*cm, 4*cm], font_size=9))

story.append(Paragraph("4. Méthodologie PACE", H1))
story.append(Paragraph(
    "La méthodologie PACE (Prepare-Analyze-Construct-Execute) structure le projet en 4 phases. "
    "L'innovation majeure vs le forecast Q4 2026 est la <b>désaisonnalisation</b> de l'effet soja exceptionnel "
    "et l'inclusion des commandes En cours/Validées.",
    BODY))

pace_data = [
    ["Phase", "Activités", "Livrables"],
    ["P - PREPARE", "Consolidation 115 592 records + En cours/Validées, calcul prix, désaisonnalisation soja", "Dataset consolidé, prix 2027"],
    ["A - ANALYZE", "AED, saisonnalité 2025, top produits/agences, analyse effet soja", "Graphiques, synthèse AED"],
    ["C - CONSTRUCT", "Prophet (4 familles) + extrapolation (2 familles), forecast 12 mois S3, désagrégation", "Forecast 2027 (7 224 lignes)"],
    ["E - EXECUTE", "Excel multi-feuilles, 5 PDFs PACE, graphiques", "7 livrables finaux"],
]
story.append(make_table(pace_data, col_widths=[3*cm, 7.5*cm, 6.5*cm], font_size=8))

story.append(Paragraph("5. Équipe et gouvernance", H1))
team_data = [
    ["Rôle", "Responsabilité", "Phase PACE"],
    ["Data Analyst (W. F. Fohom)", "Modélisation, AED, forecast, livrables", "P + A + C + E"],
    ["Direction Commerciale", "Validation hypothèses et objectifs", "A + E"],
    ["Direction Production", "Stocks, capacité, réappro soja", "P + C"],
    ["Direction Générale", "Décisions stratégiques", "E"],
    ["Contrôle de Gestion", "Cohérence financière", "E"],
]
story.append(make_table(team_data, col_widths=[4.5*cm, 8.5*cm, 4*cm], font_size=9))

story.append(Paragraph("6. Risques et mitigation", H1))
risks_data = [
    ["Risque", "Probabilité", "Mitigation"],
    ["Saisonnalité sous-estimée (20 mois d'historique)", "Moyenne", "Désaisonnalisation soja + mise à jour trimestrielle"],
    ["Prix soja volatil (déjà +47% en 2 mois)", "Élevée", "Prix actualisé 25 000 FCFA, scénario S3"],
    ["Effet soja 2026 non récurrent en 2027", "Élevée", "Cap désaisonnalisation à moyenne S1"],
    ["Commandes En cours/Validées non converties", "Faible", "Filtre qualité, exclusion Annulées"],
    ["Nouvelles hausses tarifaires en 2027", "Moyenne", "Hypothèse prix stable en 2027"],
]
story.append(make_table(risks_data, col_widths=[7*cm, 3*cm, 7*cm], font_size=9))

story.append(Paragraph("7. Critères de succès", H1))
success = [
    "Forecast 2027 produit sur 12 mois avec désagrégation complète (produit × agence × mois)",
    "Désaisonnalisation effective de l'effet soja exceptionnel",
    "En cours et Validées intégrés comme potentielles ventes",
    "Prix soja actualisé à 25 000 FCFA/sac",
    "Écart forecast vs réalité ≤ 15% par trimestre",
    "Livrables PACE complets (7 documents) produits et diffusés",
    "Réutilisabilité du modèle pour les mises à jour trimestrielles",
]
for s in success:
    story.append(Paragraph(f"• {s}", BULLET))

doc.build(story)
print(f"✓ Proposition: {prop_path} ({os.path.getsize(prop_path)/1024:.0f} KB)")

# ==================== 3. MATRICE RACI ====================
print("Generating 3. Matrice RACI 2027...")
raci_path = f"{OUT_DIR}/03_matrice_raci.pdf"
doc = SimpleDocTemplate(raci_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
story = []

story.extend(cover_page("Matrice RACI", "Forecast 2027", "MATRICE RACI"))

story.append(Paragraph("1. Définition des rôles RACI", H1))
raci_def = [
    ["Lettre", "Rôle", "Description"],
    ["R", "Responsable", "Celui qui réalise la tâche"],
    ["A", "Approbateur", "Celui qui rend compte et valide"],
    ["C", "Consulté", "Celui dont l'avis est sollicité"],
    ["I", "Informé", "Celui qui est tenu au courant"],
]
story.append(make_table(raci_def, col_widths=[2*cm, 5*cm, 11*cm], font_size=9))

story.append(Paragraph("2. Acteurs du projet", H1))
actors = [
    ["Code", "Acteur", "Implication"],
    ["DA", "Data Analyst (W. F. Fohom)", "Pleine (R sur toutes les tâches techniques)"],
    ["DC", "Direction Commerciale", "Validation hypothèses et objectifs"],
    ["DP", "Direction Production", "Stocks, capacité, réappro soja"],
    ["DG", "Direction Générale", "Décision finale et allocation ressources"],
    ["CG", "Contrôle de Gestion", "Cohérence financière"],
    ["RA", "Responsables d'Agences (14)", "Informés et consultés sur leur périmètre"],
]
story.append(make_table(actors, col_widths=[1.5*cm, 6*cm, 10.5*cm], font_size=9))

story.append(Paragraph("3. Matrice RACI détaillée par étape PACE", H1))
raci_matrix = [
    ["Phase", "Tâche", "DA", "DC", "DP", "DG", "CG", "RA"],
    ["PREPARE", "1.1 Consolidation données + En cours/Validées", "R/A", "I", "C", "I", "I", "I"],
    ["", "1.2 Désaisonnalisation effet soja", "R/A", "C", "C", "I", "I", "I"],
    ["", "1.3 Calcul prix 2027 (soja 25 000)", "R/A", "C", "I", "I", "C", "I"],
    ["", "1.4 Validation dataset", "R", "C", "C", "I", "A", "I"],
    ["ANALYZE", "2.1 AED + saisonnalité", "R/A", "I", "I", "I", "I", "I"],
    ["", "2.2 Analyse effet soja exceptionnel", "R/A", "C", "C", "I", "I", "I"],
    ["", "2.3 Validation hypothèses AED", "R", "C", "C", "I", "A", "I"],
    ["CONSTRUCT", "3.1 Prophet (4 familles × 3 régions)", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.2 Extrapolation MAT+PREMIX", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.3 Forecast 12 mois S3", "R", "C", "I", "A", "C", "I"],
    ["", "3.4 Désagrégation produit × agence", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.5 Validation forecast 2027", "R", "C", "I", "A", "C", "I"],
    ["EXECUTE", "4.1 Excel 7 feuilles", "R/A", "I", "I", "I", "I", "I"],
    ["", "4.2 PDFs PACE (5 documents)", "R/A", "I", "I", "I", "I", "I"],
    ["", "4.3 Présentation CODIR", "R", "C", "C", "A", "C", "I"],
    ["", "4.4 Diffusion agences", "R", "A", "I", "I", "I", "C"],
    ["", "4.5 Mise à jour trimestrielle", "R/A", "C", "C", "I", "I", "I"],
]
story.append(make_table(raci_matrix, col_widths=[2.5*cm, 6.5*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm], font_size=7.5))

story.append(Paragraph("4. Points de contrôle", H1))
control_data = [
    ["Jalon", "Phase", "Date", "Approbateur", "Livrable"],
    ["Kick-off", "PREPARE", "02/09/2026", "DG", "Proposition validée"],
    ["Validation dataset", "PREPARE", "03/09/2026", "CG", "Dataset + désaisonnalisation"],
    ["Validation AED", "ANALYZE", "04/09/2026", "DC", "Synthèse AED"],
    ["Validation forecast", "CONSTRUCT", "05/09/2026", "DG + DC", "Forecast 2027 complet"],
    ["Livrables finaux", "EXECUTE", "06/09/2026", "DG", "7 livrables"],
    ["Présentation CODIR", "EXECUTE", "07/09/2026", "DG", "Validation officielle"],
    ["Mise à jour Q1", "—", "Mars 2027", "DA", "Réactualisation"],
]
story.append(make_table(control_data, col_widths=[3.5*cm, 2.5*cm, 3*cm, 3*cm, 5*cm], font_size=8))

doc.build(story)
print(f"✓ Matrice RACI: {raci_path} ({os.path.getsize(raci_path)/1024:.0f} KB)")

# ==================== 4. DOCUMENT STRATÉGIQUE PACE ====================
print("Generating 4. Document stratégique PACE 2027...")
strat_path = f"{OUT_DIR}/04_document_strategique_pace.pdf"
doc = SimpleDocTemplate(strat_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []

story.extend(cover_page("Document Stratégique PACE", "Forecast 2027 - Volumes et Valeurs (12 mois)", "STRATÉGIE PACE"))

story.append(Paragraph("1. Vision et objectifs", H1))
story.append(Paragraph("1.1 Vision", H2))
story.append(Paragraph(
    "Faire du forecast 2027 un <b>outil de planification annuelle</b> permettant à BELGOCAM SA d'anticiper "
    "109 288 tonnes de ventes et 58 550 M FCFA de chiffre d'affaires, avec une précision opérationnelle "
    "et une désaisonnalisation de l'effet soja exceptionnel de 2026.",
    BODY))

story.append(Paragraph("1.2 Objectifs stratégiques", H2))
obj_data = [
    ["Objectif", "Description", "KPI"],
    ["Forecast annuel", "12 mois 2027 en volume + valeur", "109 288 t, 58 550 M FCFA"],
    ["Désaisonnalisation", "Neutraliser l'effet soja Jul-Août 2026", "Cap moyenne S1 appliqué"],
    ["En cours + Validées", "Intégrer comme potentielles ventes", "255 commandes incluses"],
    ["Prix actualisé", "Soja 25 000 FCFA/sac", "vs 20 600 précédent"],
    ["Désagrégation", "Produit × agence × mois", "7 224 lignes"],
    ["Adoption", "Diffusion CODIR + 14 agences", "100% agences informées"],
]
story.append(make_table(obj_data, col_widths=[3.5*cm, 8.5*cm, 5*cm], font_size=9))

story.append(Paragraph("2. Phase PREPARE - Préparation des données", H1))
story.append(Paragraph("2.1 Sources consolidées", H2))
sources_data = [
    ["Source", "Période", "Records"],
    ["Historique 2025", "Jan-Déc 2025", "64 580"],
    ["S1 2026", "Jan-Juin 2026", "35 621"],
    ["Juillet 2026", "Juillet 2026", "7 410"],
    ["Août 2026", "01-31/08/2026", "6 431"],
    ["En cours + Validées", "Août 2026", "~255"],
    ["TOTAL", "20 mois", "115 592"],
]
story.append(make_table(sources_data, col_widths=[5*cm, 5*cm, 5*cm], font_size=9))

story.append(Paragraph("2.2 Désaisonnalisation de l'effet soja", H2))
story.append(Paragraph(
    "Les volumes soja de juillet-août 2026 ont été exceptionnellement élevés en raison de la rupture concurrente. "
    "Pour éviter que Prophet n'extrapole ce phénomène non récurrent, les volumes soja Jul-Août 2026 ont été "
    "<b>cappés à la moyenne S1</b> (janvier-juin 2026). Cette désaisonnalisation permet de préserver le signal "
    "de saisonnalité annuelle tout en neutralisant l'effet exceptionnel.",
    BODY))

story.append(Paragraph("3. Phase CONSTRUCT - Modélisation", H1))
story.append(Paragraph("3.1 Configuration Prophet", H2))
config_data = [
    ["Paramètre", "Valeur", "Justification"],
    ["yearly_seasonality", "True", "Capture la saisonnalité annuelle (pic Q4)"],
    ["seasonality_mode", "multiplicative", "Saisonnalité proportionnelle à la tendance"],
    ["changepoint_prior_scale", "0.05", "Tendance modérément flexible"],
    ["interval_width", "0.8", "Intervalle de confiance 80%"],
    ["Période forecast", "12 mois (Jan-Déc 2027)", "Année complète"],
    ["Future dates", "Explicites (Jan 2027 - Déc 2027)", "Pas de make_future_dataframe"],
]
story.append(make_table(config_data, col_widths=[5*cm, 5*cm, 7*cm], font_size=9))

story.append(Paragraph("4. Résultats par famille et trimestre", H1))
story.append(Paragraph("4.1 Par famille", H2))
fam_detail = [
    ["Famille", "Volume (t)", "CA (M FCFA)", "Part CA", "Méthode"],
    ["TOURTEAUX", "87 399", "43 701", "74,6%", "Prophet"],
    ["CONCENTRÉS", "19 565", "13 063", "22,3%", "Prophet"],
    ["ALIMENT_COMPLET", "1 327", "1 038", "1,8%", "Prophet"],
    ["MATERIEL_ELEVAGE", "—", "376", "0,6%", "Extrapolation CA"],
    ["INGREDIENTS", "996", "209", "0,4%", "Prophet"],
    ["PREMIX", "—", "163", "0,3%", "Extrapolation CA"],
]
story.append(make_table(fam_detail, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2*cm, 4*cm], font_size=9))

story.append(Paragraph("4.2 Par trimestre", H2))
q_detail = [
    ["Trimestre", "Volume (t)", "CA (M FCFA)", "Part CA", "Lecture"],
    ["Q1 (Jan-Mar)", "19 309", "10 599", "18,1%", "Démarrage progressif"],
    ["Q2 (Avr-Juin)", "29 614", "15 673", "26,8%", "Montée en charge"],
    ["Q3 (Juil-Sept)", "10 950", "6 427", "11,0%", "Creux saisonnier"],
    ["Q4 (Oct-Déc)", "49 414", "25 851", "44,1%", "Pic saisonnier"],
]
story.append(make_table(q_detail, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2*cm, 4.5*cm], font_size=9))


story.append(PageBreak())
story.append(Paragraph("4.5 Réalisation 2026 YTD + Projection fin d'année (contexte)", H1))
story.append(Paragraph(
    "Cette section présente la trajectoire 2026 (YTD réel Jan-Août + Q4 forecast S3) qui sert de base "
    "de comparaison au forecast 2027. L'analyse révèle une surperformance exceptionnelle du TOURTEAUX "
    "(136% de l'objectif annuel) due à la rupture concurrente, et un déficit structurel des CONCENTRÉS (76%).",
    BODY))

proj_data = [
    ["Famille", "YTD réel (t)", "Q4 fcst (t)", "Total 2026 (t)", "Obj annuel (t)", "% Annuel", "Écart (t)"],
    ["TOURTEAUX", "39 718", "33 853", "73 571", "54 196", "136%", "+19 375"],
    ["CONCENTRÉS", "12 066", "6 393", "18 459", "24 192", "76%", "-5 733"],
    ["ALIMENT COMPLET", "525", "368", "893", "1 046", "85%", "-153"],
    ["INGRÉDIENTS", "548", "307", "855", "1 064", "80%", "-209"],
    ["PREMIX", "61", "0", "61", "130", "47%", "❌"],
    ["MAÏS (exclu)", "0", "0", "0", "—", "—", "—"],
    ["TOTAL", "52 918", "40 921", "93 839", "80 628", "116%", "+13 211"],
]
story.append(make_table(proj_data, col_widths=[3*cm, 2*cm, 2.5*cm, 2.5*cm, 2.5*cm, 1.5*cm, 2.5*cm], font_size=8))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph(
    "<b>Analyse</b> : Le TOURTEAUX surperforme massivement (+19 375 t) grâce à la rupture concurrente "
    "de juillet-août 2026. Cette surperformance est exceptionnelle et non récurrente. Les CONCENTRÉS restent "
    "à 76% de l'objectif annuel (-5 733 t) : malgré un mois d'août à 113%, le retard accumulé sur S1 n'est pas "
    "rattrapable en Q4. Au niveau global, l'objectif annuel est dépassé à 116% (+13 211 t), mais cette performance "
    "est entièrement tirée par le TOURTEAUX. Cette analyse justifie la désaisonnalisation de l'effet soja "
    "dans le forecast 2027, pour éviter de projeter un phénomène non récurrent.",
    BODY))


story.append(Paragraph("5. Plan de déploiement", H1))
deploy_data = [
    ["Étape", "Date", "Action", "Cible"],
    ["1", "02/09/2026", "Validation proposition", "DG"],
    ["2", "06/09/2026", "Livrables finaux prêts", "DA"],
    ["3", "07/09/2026", "Présentation CODIR", "DG + DC + DP + CG"],
    ["4", "08/09/2026", "Diffusion aux agences", "14 RA"],
    ["5", "Mars 2027", "Mise à jour Q1 avec données réelles", "DA"],
    ["6", "Juin 2027", "Mise à jour Q2", "DA"],
    ["7", "Sept 2027", "Mise à jour Q3", "DA"],
]
story.append(make_table(deploy_data, col_widths=[1*cm, 3*cm, 7*cm, 6*cm], font_size=9))

story.append(Paragraph("6. Conclusion", H1))
story.append(Paragraph(
    "Le forecast 2027 projette <b>109 288 tonnes</b> pour <b>58 550 M FCFA</b>, avec une concentration "
    "saisonnière au Q4 (44%). La désaisonnalisation de l'effet soja et l'actualisation du prix à 25 000 FCFA "
    "permettent une projection réaliste. Le pic d'octobre (20 602 t) nécessitera une anticipation renforcée "
    "du réapprovisionnement soja dès septembre 2027.",
    BODY))

doc.build(story)
print(f"✓ Document stratégique: {strat_path} ({os.path.getsize(strat_path)/1024:.0f} KB)")

# ==================== 5. GUIDE MÉTHODOLOGIQUE ====================
print("Generating 5. Guide méthodologique 2027...")
guide_path = f"{OUT_DIR}/05_guide_methodologique.pdf"
doc = SimpleDocTemplate(guide_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []

story.extend(cover_page("Guide Méthodologique", "Forecast 2027 - Méthodes Prophet et désaisonnalisation", "GUIDE MÉTHODOLOGIQUE"))

story.append(Paragraph("Sommaire", H1))
story.append(Paragraph(
    "1. Introduction et contexte<br/>"
    "2. Méthodologie PACE<br/>"
    "3. Outils et environnement<br/>"
    "4. Phase PREPARE - Désaisonnalisation soja<br/>"
    "5. Phase PREPARE - En cours + Validées<br/>"
    "6. Phase ANALYZE - AED<br/>"
    "7. Phase CONSTRUCT - Prophet<br/>"
    "8. Phase EXECUTE - Livrables<br/>"
    "Annexes",
    BODY))

story.append(PageBreak())
story.append(Paragraph("1. Introduction et contexte", H1))
story.append(Paragraph(
    "Ce guide décrit la méthodologie complète du forecast 2027 de BELGOCAM SA. Il couvre 12 mois "
    "(janvier-décembre 2027) avec 6 familles de produits et 14 agences. Les innovations majeures vs "
    "le forecast Q4 2026 sont : (1) désaisonnalisation de l'effet soja exceptionnel, "
    "(2) inclusion des commandes En cours/Validées, (3) actualisation du prix soja à 25 000 FCFA, "
    "(4) forecast annuel complet (12 mois vs 4 mois).",
    BODY))

story.append(Paragraph("2. Méthodologie PACE", H1))
story.append(Paragraph(
    "La méthodologie PACE (Prepare-Analyze-Construct-Execute) structure le projet en 4 phases successives, "
    "chacune avec des livrables vérifiables.",
    BODY))
pace_detail = [
    ["Phase", "Objectif", "Livrables", "Durée"],
    ["P - PREPARE", "Préparer données + désaisonnalisation", "Dataset 115 592 records, prix 2027", "2 jours"],
    ["A - ANALYZE", "Comprendre données + effet soja", "Graphiques, synthèse AED", "1 jour"],
    ["C - CONSTRUCT", "Modéliser Prophet 12 mois", "Forecast 2027 (7 224 lignes)", "2 jours"],
    ["E - EXECUTE", "Produire livrables finaux", "Excel + 5 PDFs", "1 jour"],
]
story.append(make_table(pace_detail, col_widths=[3*cm, 4*cm, 5*cm, 3*cm], font_size=9))

story.append(Paragraph("3. Outils et environnement", H1))
stack_data = [
    ["Outil", "Version", "Usage"],
    ["Python", "3.12", "Langage principal"],
    ["pandas", "2.x", "Manipulation données"],
    ["Prophet", "1.1+", "Modélisation prédictive"],
    ["openpyxl", "3.x", "Génération Excel"],
    ["ReportLab", "4.x", "Génération PDFs"],
    ["matplotlib", "3.x", "Visualisations"],
]
story.append(make_table(stack_data, col_widths=[4*cm, 3*cm, 8*cm], font_size=9))

story.append(PageBreak())
story.append(Paragraph("4. Phase PREPARE - Désaisonnalisation soja", H1))
story.append(Paragraph(
    "L'effet soja exceptionnel de juillet-août 2026 (rupture concurrente) a généré des volumes anormalement "
    "élevés qui ne sont pas récurrents. Pour éviter que Prophet n'extrapole ce phénomène, les volumes soja "
    "de Jul-Août 2026 ont été cappés à la moyenne S1.",
    BODY))
story.append(Paragraph("<b>Méthode</b>", H3))
story.append(Paragraph(
    "1. Calcul de la moyenne mensuelle soja S1 (Jan-Juin 2026)<br/>"
    "2. Identification du mois Jul-Août 2026 avec le volume le plus élevé<br/>"
    "3. Calcul du facteur de cap = moyenne S1 / volume max Jul-Août<br/>"
    "4. Application du facteur à tous les enregistrements soja Jul-Août 2026<br/>"
    "5. Les volumes sont réduits proportionnellement, préservant le signal de saisonnalité",
    BODY))

story.append(Paragraph("5. Phase PREPARE - En cours + Validées", H1))
story.append(Paragraph(
    "Les commandes En cours (98) et Validées (157) d'août 2026 sont incluses dans le dataset comme "
    "potentielles ventes Livrées. Les commandes Annulées (159) et Brouillon (96) sont exclues. "
    "Cette intégration permet de capturer le pipeline commercial en cours au moment de l'extraction.",
    BODY))

story.append(Paragraph("6. Phase ANALYZE - AED", H1))
story.append(Paragraph(
    "L'AED révèle : (1) TOURTEAUX = 76% du volume historique, (2) Q4 = 35-46% du volume annuel selon les familles, "
    "(3) pic octobre (indice 1.86 pour TOURTEAUX), (4) FAMLA = 31.6% du volume, (5) ratio bundle 2,3:1 en août 2026.",
    BODY))

story.append(Paragraph("7. Phase CONSTRUCT - Prophet", H1))
story.append(Paragraph("7.1 Configuration", H3))
story.append(Paragraph(
    "Prophet est configuré avec yearly_seasonality=True (saisonnalité annuelle), seasonality_mode='multiplicative' "
    "(saisonnalité proportionnelle), changepoint_prior_scale=0.05 (tendance modérée), interval_width=0.8. "
    "Les future dates sont explicites (Jan-Déc 2027) pour garantir 12 mois de forecast.",
    BODY))

story.append(Paragraph("7.2 Stratégie de modélisation", H3))
story.append(Paragraph(
    "13 modèles Prophet famille × région (4 familles × 3 régions + extrapolation pour MATERIEL_ELEVAGE et PREMIX). "
    "Désagrégation par produit × agence selon les parts historiques de CA. Prix 2027 : 25 000 FCFA/sac soja, "
    "prix août 2026 pour les autres familles.",
    BODY))

story.append(Paragraph("8. Phase EXECUTE - Livrables", H1))
livrables_data = [
    ["#", "Livrable", "Format", "Usage"],
    ["1", "Résumé exécutif", "PDF", "Direction Générale"],
    ["2", "Proposition de projet", "PDF", "CODIR"],
    ["3", "Matrice RACI", "PDF", "Gouvernance"],
    ["4", "Document stratégique PACE", "PDF", "Référence stratégique"],
    ["5", "Guide méthodologique", "PDF", "Référence méthodes"],
    ["6", "Excel forecast 2027", "XLSX 7f", "Pilotage opérationnel"],
    ["7", "Graphiques", "PNG", "Présentations"],
]
story.append(make_table(livrables_data, col_widths=[1*cm, 5*cm, 3*cm, 6*cm], font_size=9))

story.append(Paragraph("Annexe - Glossaire", H1))
glossaire = [
    ["Terme", "Définition"],
    ["Prophet", "Bibliothèque prévision séries temporelles (Facebook/Meta)"],
    ["Désaisonnalisation", "Neutralisation d'un effet exceptionnel non récurrent"],
    ["S3", "Scénario réappro soja 100% (situation normale)"],
    ["En cours + Validées", "Commandes non encore Livrées, potentielles ventes"],
    ["Cap S1", "Plafonnement des volumes Jul-Août 2026 à la moyenne S1"],
    ["Q4", "Quatrième trimestre (oct-déc) = 44% du CA 2027"],
    ["PACE", "Prepare-Analyze-Construct-Execute"],
    ["RACI", "Responsible-Accountable-Consulted-Informed"],
]
story.append(make_table(glossaire, col_widths=[4*cm, 11*cm], font_size=9))

doc.build(story)
print(f"✓ Guide méthodologique: {guide_path} ({os.path.getsize(guide_path)/1024:.0f} KB)")

print("\n=== TOUS LES 5 PDFs 2027 GÉNÉRÉS ===")
for f in sorted(os.listdir(OUT_DIR)):
    if f.endswith('.pdf'):
        print(f"  {f}: {os.path.getsize(os.path.join(OUT_DIR, f))/1024:.0f} KB")
