"""
Forecast 2027 - Génération des 5 PDFs PACE - VERSION 2
AVEC COMPLEMENT_ALIMENTAIRE et données 2024-2026

Changements vs version 1:
- 7 familles au lieu de 6 (ajout COMPLEMENT_ALIMENTAIRE)
- 78 produits (vs 69)
- 118 608 t, 64 114 M FCFA (vs 75 112 t, 41 219 M FCFA)
- Données 2024-2026 (32 mois) au lieu de 2025-2026 (20 mois)
- Proxy BELGOKILL pour COMPLEMENT_ALIMENTAIRE
- Conversion 1L = 1kg
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
NEW_FAMILY_COLOR = 'FFE699'  # Jaune pour COMPLEMENT_ALIMENTAIRE

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

def make_table(data, col_widths=None, font_size=9, header_color=NAVY, highlight_rows=None):
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
            style_list.append(('BACKGROUND', (0, ridx), (-1, ridx), HexColor('#FFE699')))
    t.setStyle(TableStyle(style_list))
    return t

def cover_page(title, subtitle, doc_type, date_str="3 septembre 2026"):
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

# ==================== 1. RÉSUMÉ EXÉCUTIF ====================
print("Generating 1. Résumé exécutif 2027 (v2)...")
exec_path = f"{OUT_DIR}/01_resume_executif.pdf"
doc = SimpleDocTemplate(exec_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []

story.append(Paragraph("RÉSUMÉ EXÉCUTIF", H1))
story.append(Paragraph("Forecast 2027 - Volumes et Valeurs (12 mois) — Version 2", H2))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Contexte</b>", H3))
story.append(Paragraph(
    "BELGOCAM SA doit anticiper ses volumes de ventes et son chiffre d'affaires pour l'année 2027 complète "
    "(janvier - décembre, 12 mois). Cette version 2 du forecast s'appuie sur <b>170 560 enregistrements</b> couvrant "
    "32 mois d'historique (juillet 2024 - août 2026), intégrant les commandes En cours et Validées d'août 2026. "
    "L'effet soja exceptionnel de juillet-août 2026 (rupture concurrente) a été <b>désaisonnalisé</b> pour éviter "
    "de biaiser les prédictions. Le prix du soja a été actualisé à <b>25 000 FCFA/sac</b> (hausse du 24/08/2026).",
    BODY))
story.append(Paragraph(
    "<b>NOUVEAUTÉ VERSION 2</b> : Ajout de la famille <b>COMPLEMENT_ALIMENTAIRE</b> (BELGOKILL, BELGO HARMONY, "
    "BELGO PROTECT, etc.) avec conversion 1L = 1kg. La tendance <b>BELGOKILL</b> (V300 = 38% du CA famille) est "
    "utilisée comme proxy pour toute la famille. Les années 2021-2023 ont été exclues (focus sur 2024-2026).",
    BODY))

story.append(Paragraph("<b>Méthodologie</b>", H3))
story.append(Paragraph(
    "Le modèle <b>Prophet</b> (Facebook/Meta) a été entraîné sur 5 familles alimentaires (TOURTEAUX, CONCENTRÉS, "
    "INGRÉDIENTS, ALIMENT COMPLET, COMPLEMENT ALIMENTAIRE) au niveau famille × région. Le MATERIEL ÉLEVAGE et les "
    "PREMIX sont projetés par extrapolation de la moyenne historique. Le forecast couvre <b>7 familles, 78 produits, "
    "14 agences</b> et 3 régions. Le scénario S3 (réappro soja 100%, situation normale) est utilisé comme référence.",
    BODY))

story.append(Paragraph("<b>Résultats clés 2027</b>", H3))
synth_data = [
    ["Indicateur", "Valeur", "Détail"],
    ["Volume total 2027", "118 608 t", "7 familles, 78 produits, 14 agences"],
    ["CA total 2027", "64 114 M FCFA", "Prix soja 25 000 FCFA/sac"],
    ["Période", "12 mois (Jan-Déc 2027)", "Forecast complet annuel"],
    ["Scénario", "S3 (réappro 100%)", "Situation normale"],
    ["Données historiques", "170 560 enregistrements", "Jul 2024 - Août 2026 + En cours/Validées"],
    ["Désaisonnalisation", "Effet soja Jul-Août 2026 neutralisé", "Cap à moyenne S1 2026"],
    ["NOUVEAU: COMPLEMENT_ALIMENTAIRE", "3 t / 17 M FCFA", "10 produits (BELGOKILL proxy)"],
]
story.append(make_table(synth_data, col_widths=[5*cm, 4*cm, 8*cm], font_size=9, highlight_rows=[7]))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Synthèse par famille</b>", H3))
fam_data = [
    ["Famille", "Volume (t)", "CA (M FCFA)", "Part CA"],
    ["TOURTEAUX", "92 517", "46 261", "72,2%"],
    ["CONCENTRÉS", "24 680", "16 504", "25,7%"],
    ["ALIMENT_COMPLET", "890", "688", "1,1%"],
    ["MATERIEL_ELEVAGE", "—", "377", "0,6%"],
    ["INGREDIENTS", "518", "107", "0,2%"],
    ["PREMIX", "—", "160", "0,3%"],
    ["COMPLEMENT_ALIMENTAIRE", "3", "17", "0,0%"],
    ["MAÏS (exclu)", "0", "0", "0,0%"],
    ["TOTAL", "118 608", "64 114", "100%"],
]
story.append(make_table(fam_data, col_widths=[5*cm, 3*cm, 3*cm, 2.5*cm], font_size=9, highlight_rows=[7]))
story.append(Paragraph("<i>Note 1: Le maïs (M1051) a été retiré des ventes 2026 avant le forecast 2027 car produit opportuniste hors portefeuille régulier BELGOCAM.</i>", SMALL))
story.append(Paragraph("<i>Note 2: COMPLEMENT_ALIMENTAIRE (NOUVEAU) = produits liquides (BELGOKILL, BELGO HARMONY, etc.), conversion 1L=1kg. Faible volume mais forte valeur unitaire.</i>", SMALL))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Synthèse par trimestre</b>", H3))
q_data = [
    ["Trimestre", "Volume (t)", "CA (M FCFA)", "Part CA"],
    ["Q1 (Jan-Mar)", "22 616", "12 451", "19,4%"],
    ["Q2 (Avr-Juin)", "31 562", "16 806", "26,2%"],
    ["Q3 (Juil-Sept)", "17 089", "9 915", "15,5%"],
    ["Q4 (Oct-Déc)", "47 340", "24 942", "38,9%"],
    ["TOTAL", "118 608", "64 114", "100%"],
]
story.append(make_table(q_data, col_widths=[4*cm, 3.5*cm, 3.5*cm, 2.5*cm], font_size=9))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Coefficients saisonniers 2025 vs 2027</b>", H3))
story.append(Paragraph(
    "Les coefficients saisonniers (1.0 = moyenne annuelle) révèlent la concentration saisonnière des ventes. "
    "Le pic d'octobre (coefficient 1,86 en 2025 pour TOURTEAUX) explique pourquoi le Q4 représente près de 39% du CA annuel.",
    BODY))

seasonal_data = [
    ["Famille", "Année", "Jan", "Fév", "Mar", "Avr", "Mai", "Juin", "Juil", "Août", "Sep", "Oct", "Nov", "Déc"],
    ["TOURTEAUX", "2025", "0,81", "0,84", "0,79", "0,95", "0,84", "0,72", "0,72", "0,77", "0,81", "1,86", "1,25", "1,65"],
    ["", "2027 fcst", "1,04", "0,49", "0,59", "0,65", "0,97", "1,77", "0,67", "0,13", "0,45", "1,94", "1,26", "2,04"],
    ["CONCENTRÉS", "2025", "1,04", "0,92", "0,95", "1,00", "0,97", "0,96", "1,05", "0,90", "0,93", "0,95", "1,15", "1,18"],
    ["", "2027 fcst", "0,95", "0,90", "0,99", "0,89", "0,78", "0,84", "1,13", "1,02", "1,24", "0,64", "1,27", "1,34"],
]
story.append(make_table(seasonal_data, col_widths=[2.2*cm, 1.8*cm, 1*cm, 1*cm, 1*cm, 1*cm, 1*cm, 1*cm, 1*cm, 1*cm, 1*cm, 1*cm, 1*cm, 1*cm], font_size=7))
story.append(Spacer(1, 0.2*cm))

story.append(Paragraph("<b>Taux de progression 2027 vs 2026 (annualisé)</b>", H3))

prog_data = [
    ["Famille", "2026 YTD (t)", "2026 annualisé", "2027 fcst (t)", "Progression", "CA 2027 (M)"],
    ["TOURTEAUX", "39 045", "~58 568", "92 517", "+58%", "46 261"],
    ["CONCENTRÉS", "11 932", "~17 898", "24 680", "+38%", "16 504"],
    ["ALIMENT COMPLET", "502", "~753", "890", "+18%", "688"],
    ["INGRÉDIENTS", "531", "~797", "518", "-35%", "107"],
    ["PREMIX", "60", "~90", "0 (CA)", "—", "160"],
    ["COMPLEMENT_ALIM.", "3", "~4", "3", "-25%", "17"],
    ["MAÏS (exclu)", "0", "0", "0", "—", "0"],
    ["TOTAL", "52 079", "~78 118", "118 608", "+52%", "64 114"],
]
story.append(make_table(prog_data, col_widths=[2.8*cm, 1.8*cm, 1.8*cm, 1.8*cm, 1.5*cm, 2.5*cm], font_size=7.5, highlight_rows=[6]))
story.append(Spacer(1, 0.2*cm))

story.append(Paragraph(
    "<b>Insights clés</b> : (1) Le pic Q4 2027 (39% du CA) est confirmé par un coefficient saisonnier de 1,94 en octobre pour TOURTEAUX. "
    "(2) La désaisonnalisation de l'effet soja 2026 évite de répliquer le pic artificiel de juillet-août sur 2027. "
    "(3) La progression globale de +52% en volume reflète la tendance haussière observée sur 2024-2026, tirée principalement par le TOURTEAUX. "
    "(4) En CA, l'effet prix soja (25 000 FCFA/sac) amplifie la progression à +64 M FCFA. "
    "(5) NOUVEAU : COMPLEMENT_ALIMENTAIRE ajouté avec tendance BELGOKILL comme proxy (3 t, 17 M FCFA — petits volumes mais forte valeur unitaire).",
    BODY))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Top 5 agences par CA 2027</b>", H3))
top_data = [
    ["Rang", "Agence", "Région", "CA (M FCFA)"],
    ["1", "FAMLA", "Ouest", "18 085"],
    ["2", "NDOBO", "Littoral", "9 476"],
    ["3", "MESSASSI", "Centre", "5 848"],
    ["4", "DJELENG", "Ouest", "5 166"],
    ["5", "VILLAGE", "Littoral", "3 749"],
]
story.append(make_table(top_data, col_widths=[1.5*cm, 4*cm, 3*cm, 4.5*cm], font_size=9))

story.append(PageBreak())

story.append(Paragraph("<b>Réalisation 2024-2026 vs Forecast 2027 (contexte)</b>", H3))
story.append(Paragraph(
    "Le forecast 2027 s'appuie sur 32 mois d'historique (Jul 2024 - Août 2026). Le tableau ci-dessous présente "
    "l'évolution par famille et par année, qui sert de base de comparaison au forecast 2027.",
    BODY))

hist_data = [
    ["Famille", "2024 S2 (t)", "2025 (t)", "2026 YTD (t)", "2027 fcst (t)", "CA 2027 (M)"],
    ["TOURTEAUX", "19 296", "44 584", "39 045", "92 517", "46 261"],
    ["CONCENTRÉS", "8 207", "17 319", "11 932", "24 680", "16 504"],
    ["ALIMENT_COMPLET", "296", "500", "502", "890", "688"],
    ["INGRÉDIENTS", "606", "733", "531", "518", "107"],
    ["PREMIX", "41", "86", "60", "0 (CA)", "160"],
    ["COMPLEMENT_ALIM.", "3", "6", "3", "3", "17"],
    ["MAÏS (exclu)", "0", "0", "0", "0", "0"],
    ["TOTAL", "28 449", "63 228", "52 079", "118 608", "64 114"],
]
story.append(make_table(hist_data, col_widths=[2.8*cm, 2.2*cm, 2*cm, 2.2*cm, 2.2*cm, 2.5*cm], font_size=8, highlight_rows=[6]))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph(
    "<b>Lecture</b> : L'historique 2024-2026 montre une tendance haussière continue pour le TOURTEAUX (soja), "
    "passant de 19 296 t (S2 2024) à 44 584 t (2025) et 39 045 t (YTD 2026 sur 8 mois). Le forecast 2027 projette "
    "cette tendance à 92 517 t (+58% vs 2026 annualisé). La famille COMPLEMENT_ALIMENTAIRE reste stable autour de "
    "3-6 t par an, portée par les liquides BELGOKILL/BELGO HARMONY.",
    BODY))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Recommandations principales</b>", H3))
recos = [
    "<b>Planification annuelle 2027</b> — Utiliser le forecast S3 (118 608 t, 64 114 M FCFA) comme base budgétaire.",
    "<b>Saisonnalité Q4</b> — Le Q4 2027 représente 39% du CA annuel (24 942 M FCFA). Préparer les stocks et les ressources dès septembre 2027.",
    "<b>Pic d'octobre</b> — Octobre 2027 = pic de l'année (16 355 t, 8 454 M FCFA). Anticiper le réapprovisionnement soja avant septembre 2027.",
    "<b>Maintien du bundle</b> — Le ratio bundle 2,3:1 atteint en août 2026 doit être maintenu en 2027 pour optimiser les CONCENTRÉS.",
    "<b>Surveillance COMPLEMENT_ALIMENTAIRE</b> — Famille ajoutée au forecast (BELGOKILL proxy). 17 M FCFA de CA prévu, à suivre pour ajustement trimestriel.",
    "<b>Mise à jour trimestrielle</b> — Actualiser le forecast chaque trimestre avec les nouvelles données ERP.",
    "<b>Surveillance FAMLA et NDOBO</b> — Ces 2 agences représentent 43% du CA 2027, leur performance est critique.",
]
for r in recos:
    story.append(Paragraph(f"• {r}", BULLET))

story.append(Spacer(1, 0.3*cm))
story.append(Paragraph("<b>Livrables produits</b>", H3))
livrables = [
    "Excel forecast 2027 (8 feuilles: synthèse, par famille × mois, par région × mois, par agence, par produit, détail complet, hypothèses, saisonnalité)",
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
    "Le forecast 2027 projette <b>118 608 tonnes</b> pour un CA de <b>64 114 M FCFA</b>, avec une forte concentration "
    "saisonnière au Q4 (39% du CA). La désaisonnalisation de l'effet soja exceptionnel de 2026, l'actualisation "
    "du prix à 25 000 FCFA/sac et l'ajout de la famille COMPLEMENT_ALIMENTAIRE (proxy BELGOKILL) permettent une "
    "projection réaliste. Le pic d'octobre (16 355 t) nécessitera une anticipation renforcée du réapprovisionnement "
    "soja dès septembre 2027.",
    BODY))

doc.build(story)
print(f"✓ Résumé exécutif: {exec_path} ({os.path.getsize(exec_path)/1024:.0f} KB)")

# ==================== 2. PROPOSITION DE PROJET ====================
print("Generating 2. Proposition de projet 2027 (v2)...")
prop_path = f"{OUT_DIR}/02_proposition_projet.pdf"
doc = SimpleDocTemplate(prop_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []

story.extend(cover_page("Proposition de Projet", "Forecast 2027 - Volumes et Valeurs (12 mois) — Version 2", "PROPOSITION DE PROJET"))

story.append(Paragraph("1. Contexte et justification", H1))
story.append(Paragraph(
    "BELGOCAM SA doit anticiper ses volumes et CA pour l'année 2027 complète. Cette version 2 du forecast annuel "
    "s'appuie sur 32 mois d'historique (juillet 2024 - août 2026) et intègre les innovations suivantes : "
    "désaisonnalisation de l'effet soja exceptionnel, inclusion des commandes En cours/Validées, actualisation "
    "du prix soja à 25 000 FCFA/sac, et ajout de la famille COMPLEMENT_ALIMENTAIRE avec proxy BELGOKILL. "
    "Le forecast couvre 7 familles (sans Maïs), 78 produits et 14 agences.",
    BODY))

story.append(Paragraph("2. Objectifs", H1))
story.append(Paragraph("<b>Objectif principal</b>", H3))
story.append(Paragraph(
    "Produire un forecast annuel 2027 (12 mois) en volumes et valeurs, désagrégé par produit, famille, "
    "agence et région, selon le scénario S3 (réappro soja 100%, situation normale).",
    BODY))
story.append(Paragraph("<b>Objectifs spécifiques</b>", H3))
objs = [
    "Modéliser la tendance et la saisonnalité mensuelles via Prophet (5 familles alimentaires)",
    "Désaisonnaliser l'effet soja exceptionnel de juillet-août 2026 pour éviter le biais",
    "Intégrer les commandes En cours et Validées comme potentielles ventes Livrées",
    "Actualiser le prix soja à 25 000 FCFA/sac (hausse du 24/08/2026)",
    "NOUVEAU : Ajouter la famille COMPLEMENT_ALIMENTAIRE avec proxy BELGOKILL (1L=1kg)",
    "NOUVEAU : Utiliser les données 2024-2026 (32 mois) au lieu de 2025-2026 (20 mois)",
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
    ["Familles incluses", "TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, ALIMENT COMPLET, MATERIEL ÉLEVAGE, PREMIX, COMPLEMENT_ALIMENTAIRE", "7 familles"],
    ["Produits", "78 références (T102, C101-C108, MAT014/MAT011/MAT017 alvéoles, P102N2, V300 BELGOKILL, etc.)", "78 produits"],
    ["Agences", "14 agences BELGOCAM (FAMLA, NDOBO, DJELENG, etc.)", "14 agences"],
    ["Régions", "Ouest, Centre, Littoral", "3 régions"],
    ["Niveau détail", "Produit × Agence × Mois", "8 868 lignes"],
    ["Scénario", "S3 - Réappro soja 100% (situation normale)", "1 scénario"],
    ["Données historiques", "170 560 enregistrements (Jul 2024 - Août 2026) + En cours/Validées", "32 mois"],
    ["Désaisonnalisation", "Effet soja Jul-Août 2026 neutralisé (cap moyenne S1)", "—"],
    ["Prix soja", "25 000 FCFA/sac (actualisé au 24/08/2026)", "—"],
    ["NOUVEAU: COMPLEMENT_ALIM.", "10 produits liquides (BELGOKILL, BELGO HARMONY, etc.), 1L=1kg", "10 produits"],
]
story.append(make_table(scope_data, col_widths=[3.5*cm, 9.5*cm, 4*cm], font_size=9, highlight_rows=[11]))

story.append(Paragraph("4. Méthodologie PACE", H1))
story.append(Paragraph(
    "La méthodologie PACE (Prepare-Analyze-Construct-Execute) structure le projet en 4 phases. "
    "L'innovation majeure de la version 2 est l'<b>ajout de la famille COMPLEMENT_ALIMENTAIRE</b> avec proxy BELGOKILL "
    "et l'utilisation des données 2024-2026 (32 mois d'historique).",
    BODY))

pace_data = [
    ["Phase", "Activités", "Livrables"],
    ["P - PREPARE", "Consolidation 170 560 records + En cours/Validées, calcul prix, désaisonnalisation soja, ajout COMPLEMENT_ALIM.", "Dataset 2024-2026, prix 2027"],
    ["A - ANALYZE", "AED, saisonnalité 2024-2025, top produits/agences, analyse effet soja", "Graphiques, synthèse AED"],
    ["C - CONSTRUCT", "Prophet (5 familles) + extrapolation (2 familles), forecast 12 mois S3, désagrégation", "Forecast 2027 (8 868 lignes)"],
    ["E - EXECUTE", "Excel multi-feuilles, 5 PDFs PACE, graphiques", "8 livrables finaux"],
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
    ["Historique 2024-2026 insuffisant (32 mois)", "Faible", "Prophet adapté, mise à jour trimestrielle"],
    ["Prix soja volatil (déjà +47% en 2 mois)", "Élevée", "Prix actualisé 25 000 FCFA, scénario S3"],
    ["Effet soja 2026 non récurrent en 2027", "Élevée", "Cap désaisonnalisation à moyenne S1"],
    ["Commandes En cours/Validées non converties", "Faible", "Filtre qualité, exclusion Annulées"],
    ["Nouvelles hausses tarifaires en 2027", "Moyenne", "Hypothèse prix stable en 2027"],
    ["COMPLEMENT_ALIMENTAIRE : volumes faibles", "Moyenne", "Proxy BELGOKILL, CA extrapolé si volume nul"],
    ["Nouvelles hausses tarifaires en 2027", "Moyenne", "Hypothèse prix stable en 2027"],
]
story.append(make_table(risks_data, col_widths=[7*cm, 3*cm, 7*cm], font_size=9))

story.append(Paragraph("7. Critères de succès", H1))
success = [
    "Forecast 2027 produit sur 12 mois avec désagrégation complète (produit × agence × mois)",
    "Désaisonnalisation effective de l'effet soja exceptionnel",
    "En cours et Validées intégrés comme potentielles ventes",
    "Prix soja actualisé à 25 000 FCFA/sac",
    "Famille COMPLEMENT_ALIMENTAIRE ajoutée avec proxy BELGOKILL",
    "Écart forecast vs réalité ≤ 15% par trimestre",
    "Livrables PACE complets (8 documents) produits et diffusés",
    "Réutilisabilité du modèle pour les mises à jour trimestrielles",
]
for s in success:
    story.append(Paragraph(f"• {s}", BULLET))

doc.build(story)
print(f"✓ Proposition: {prop_path} ({os.path.getsize(prop_path)/1024:.0f} KB)")

# ==================== 3. MATRICE RACI ====================
print("Generating 3. Matrice RACI 2027 (v2)...")
raci_path = f"{OUT_DIR}/03_matrice_raci.pdf"
doc = SimpleDocTemplate(raci_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
story = []

story.extend(cover_page("Matrice RACI", "Forecast 2027 - Version 2", "MATRICE RACI"))

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
    ["PREPARE", "1.1 Consolidation 2024-2026 + En cours/Validées", "R/A", "I", "C", "I", "I", "I"],
    ["", "1.2 Désaisonnalisation effet soja", "R/A", "C", "C", "I", "I", "I"],
    ["", "1.3 Calcul prix 2027 (soja 25 000)", "R/A", "C", "I", "I", "C", "I"],
    ["", "1.4 NOUVEAU: Ajout COMPLEMENT_ALIMENTAIRE", "R/A", "C", "I", "I", "I", "I"],
    ["", "1.5 Validation dataset", "R", "C", "C", "I", "A", "I"],
    ["ANALYZE", "2.1 AED + saisonnalité", "R/A", "I", "I", "I", "I", "I"],
    ["", "2.2 Analyse effet soja exceptionnel", "R/A", "C", "C", "I", "I", "I"],
    ["", "2.3 Validation hypothèses AED", "R", "C", "C", "I", "A", "I"],
    ["CONSTRUCT", "3.1 Prophet (5 familles × 3 régions)", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.2 Extrapolation MAT+PREMIX", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.3 Forecast 12 mois S3", "R", "C", "I", "A", "C", "I"],
    ["", "3.4 Désagrégation produit × agence", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.5 Validation forecast 2027", "R", "C", "I", "A", "C", "I"],
    ["EXECUTE", "4.1 Excel 8 feuilles", "R/A", "I", "I", "I", "I", "I"],
    ["", "4.2 PDFs PACE (5 documents)", "R/A", "I", "I", "I", "I", "I"],
    ["", "4.3 Présentation CODIR", "R", "C", "C", "A", "C", "I"],
    ["", "4.4 Diffusion agences", "R", "A", "I", "I", "I", "C"],
    ["", "4.5 Mise à jour trimestrielle", "R/A", "C", "C", "I", "I", "I"],
]
story.append(make_table(raci_matrix, col_widths=[2.5*cm, 6.5*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm], font_size=7.5))

story.append(Paragraph("4. Points de contrôle", H1))
control_data = [
    ["Jalon", "Phase", "Date", "Approbateur", "Livrable"],
    ["Kick-off", "PREPARE", "03/09/2026", "DG", "Proposition validée"],
    ["Validation dataset v2", "PREPARE", "03/09/2026", "CG", "Dataset 2024-2026 + COMPLEMENT_ALIM."],
    ["Validation AED", "ANALYZE", "04/09/2026", "DC", "Synthèse AED"],
    ["Validation forecast", "CONSTRUCT", "05/09/2026", "DG + DC", "Forecast 2027 complet"],
    ["Livrables finaux", "EXECUTE", "05/09/2026", "DG", "8 livrables"],
    ["Présentation CODIR", "EXECUTE", "06/09/2026", "DG", "Validation officielle"],
    ["Mise à jour Q1", "—", "Mars 2027", "DA", "Réactualisation"],
]
story.append(make_table(control_data, col_widths=[3.5*cm, 2.5*cm, 3*cm, 3*cm, 5*cm], font_size=8))

doc.build(story)
print(f"✓ Matrice RACI: {raci_path} ({os.path.getsize(raci_path)/1024:.0f} KB)")

# ==================== 4. DOCUMENT STRATÉGIQUE PACE ====================
print("Generating 4. Document stratégique PACE 2027 (v2)...")
strat_path = f"{OUT_DIR}/04_document_strategique_pace.pdf"
doc = SimpleDocTemplate(strat_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []

story.extend(cover_page("Document Stratégique PACE", "Forecast 2027 - Version 2 (7 familles, données 2024-2026)", "STRATÉGIE PACE"))

story.append(Paragraph("1. Vision et objectifs", H1))
story.append(Paragraph("1.1 Vision", H2))
story.append(Paragraph(
    "Faire du forecast 2027 un <b>outil de planification annuelle</b> permettant à BELGOCAM SA d'anticiper "
    "118 608 tonnes de ventes et 64 114 M FCFA de chiffre d'affaires, avec une précision opérationnelle, "
    "une désaisonnalisation de l'effet soja exceptionnel de 2026, et l'intégration de la famille "
    "COMPLEMENT_ALIMENTAIRE avec proxy BELGOKILL.",
    BODY))

story.append(Paragraph("1.2 Objectifs stratégiques", H2))
obj_data = [
    ["Objectif", "Description", "KPI"],
    ["Forecast annuel", "12 mois 2027 en volume + valeur", "118 608 t, 64 114 M FCFA"],
    ["Désaisonnalisation", "Neutraliser l'effet soja Jul-Août 2026", "Cap moyenne S1 2026"],
    ["En cours + Validées", "Intégrer comme potentielles ventes", "255 commandes incluses"],
    ["Prix actualisé", "Soja 25 000 FCFA/sac", "vs 20 600 précédent"],
    ["Désagrégation", "Produit × agence × mois", "8 868 lignes"],
    ["Adoption", "Diffusion CODIR + 14 agences", "100% agences informées"],
    ["NOUVEAU: COMPLEMENT_ALIM.", "10 produits liquides (1L=1kg), proxy BELGOKILL", "3 t, 17 M FCFA"],
]
story.append(make_table(obj_data, col_widths=[3.5*cm, 8.5*cm, 5*cm], font_size=9, highlight_rows=[7]))

story.append(Paragraph("2. Phase PREPARE - Préparation des données", H1))
story.append(Paragraph("2.1 Sources consolidées (2024-2026)", H2))
sources_data = [
    ["Source", "Période", "Records"],
    ["LY_24 (Jul-Dec 2024)", "Juillet-Décembre 2024", "52 215"],
    ["Historique 2025", "Jan-Déc 2025", "66 206"],
    ["S1 2026", "Jan-Juin 2026", "37 745"],
    ["Juillet 2026", "Juillet 2026", "7 537"],
    ["Août 2026", "01-31/08/2026", "6 857"],
    ["En cours + Validées", "Août 2026", "~255"],
    ["TOTAL", "32 mois", "170 560"],
]
story.append(make_table(sources_data, col_widths=[5*cm, 5*cm, 5*cm], font_size=9))

story.append(Paragraph("2.2 Désaisonnalisation de l'effet soja", H2))
story.append(Paragraph(
    "Les volumes soja de juillet-août 2026 ont été exceptionnellement élevés en raison de la rupture concurrente. "
    "Pour éviter que Prophet n'extrapole ce phénomène non récurrent, les volumes soja Jul-Août 2026 ont été "
    "<b>cappés à la moyenne S1 2026</b> (janvier-juin 2026). Cette désaisonnalisation permet de préserver le signal "
    "de saisonnalité annuelle tout en neutralisant l'effet exceptionnel.",
    BODY))

story.append(Paragraph("2.3 NOUVEAU: Famille COMPLEMENT_ALIMENTAIRE", H2))
story.append(Paragraph(
    "La famille COMPLEMENT_ALIMENTAIRE (10 produits liquides) a été ajoutée au forecast 2027. "
    "Ces produits (BELGOKILL V300/V305, BELGO HARMONY, BELGO PROTECT, BELGO DRY LIT, BELGO WATER CLEAN, "
    "BELGO VIT Ese, BELGO THERMO, BELGO BIO SELECT, BELGO FRESH) sont des liquides conditionnés en bidons "
    "de 1L (ou 200L pour V305). La conversion <b>1L = 1kg</b> permet d'exprimer les volumes en tonnes. "
    "La tendance <b>BELGOKILL</b> (V300, produit dominant avec 38% du CA famille) est utilisée comme proxy "
    "pour toute la famille via Prophet.",
    BODY))

ca_ref_data = [
    ["Réf.", "Produit", "CA 2024-2026 (M FCFA)", "Part famille"],
    ["V300", "BELGOKILL 1L", "20,4", "38%"],
    ["CA003.1", "BELGO HARMONY 1L", "8,8", "16%"],
    ["CA002.1", "BELGO VIT Ese 1L", "7,6", "14%"],
    ["CA005.1", "BELGO THERMO 1L", "4,5", "8%"],
    ["CA004.1", "BELGO PROTECT 1L", "3,9", "7%"],
    ["CA008.1", "BELGO FRESH 1L", "2,9", "5%"],
    ["CA007.1", "BELGO BIO SELECT 1L", "2,8", "5%"],
    ["CA006.1", "BELGO DRY LIT 1L", "1,8", "3%"],
    ["CA001.1", "BELGO WATER CLEAN 1L", "1,5", "3%"],
    ["TOTAL", "10 produits", "54,0", "100%"],
]
story.append(make_table(ca_ref_data, col_widths=[2*cm, 6*cm, 4*cm, 3*cm], font_size=9))
story.append(Spacer(1, 0.3*cm))

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
    ["Familles Prophet", "5 (TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, ALIMENT_COMPLET, COMPLEMENT_ALIMENTAIRE)", "Nouveau: COMPLEMENT_ALIMENTAIRE"],
    ["Familles extrapolation", "2 (MATERIEL_ELEVAGE, PREMIX)", "CA only"],
]
story.append(make_table(config_data, col_widths=[5*cm, 5*cm, 7*cm], font_size=9, highlight_rows=[7]))

story.append(Paragraph("4. Résultats par famille et trimestre", H1))
story.append(Paragraph("4.1 Par famille", H2))
fam_detail = [
    ["Famille", "Volume (t)", "CA (M FCFA)", "Part CA", "Méthode"],
    ["TOURTEAUX", "92 517", "46 261", "72,2%", "Prophet"],
    ["CONCENTRÉS", "24 680", "16 504", "25,7%", "Prophet"],
    ["ALIMENT_COMPLET", "890", "688", "1,1%", "Prophet"],
    ["MATERIEL_ELEVAGE", "—", "377", "0,6%", "Extrapolation CA"],
    ["INGREDIENTS", "518", "107", "0,2%", "Prophet"],
    ["PREMIX", "—", "160", "0,3%", "Extrapolation CA"],
    ["COMPLEMENT_ALIM.", "3", "17", "0,0%", "Prophet (proxy BELGOKILL)"],
]
story.append(make_table(fam_detail, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2*cm, 4*cm], font_size=9, highlight_rows=[7]))

story.append(Paragraph("4.2 Par trimestre", H2))
q_detail = [
    ["Trimestre", "Volume (t)", "CA (M FCFA)", "Part CA", "Lecture"],
    ["Q1 (Jan-Mar)", "22 616", "12 451", "19,4%", "Démarrage progressif"],
    ["Q2 (Avr-Juin)", "31 562", "16 806", "26,2%", "Montée en charge"],
    ["Q3 (Juil-Sept)", "17 089", "9 915", "15,5%", "Creux saisonnier"],
    ["Q4 (Oct-Déc)", "47 340", "24 942", "38,9%", "Pic saisonnier"],
]
story.append(make_table(q_detail, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2*cm, 4.5*cm], font_size=9))

story.append(PageBreak())
story.append(Paragraph("4.5 Historique 2024-2026 vs Forecast 2027", H1))
story.append(Paragraph(
    "Cette section présente l'évolution des volumes par famille sur 2024-2026 (historique) vs 2027 (forecast). "
    "L'analyse révèle une tendance haussière continue pour le TOURTEAUX et les CONCENTRÉS, tandis que la famille "
    "COMPLEMENT_ALIMENTAIRE reste stable autour de 3-6 t par an.",
    BODY))

hist_detail = [
    ["Famille", "2024 S2 (t)", "2025 (t)", "2026 YTD (t)", "2027 fcst (t)", "Évolution"],
    ["TOURTEAUX", "19 296", "44 584", "39 045", "92 517", "Tendance haussière"],
    ["CONCENTRÉS", "8 207", "17 319", "11 932", "24 680", "Croissance modérée"],
    ["ALIMENT_COMPLET", "296", "500", "502", "890", "Croissance +18%"],
    ["INGRÉDIENTS", "606", "733", "531", "518", "Stable"],
    ["PREMIX", "41", "86", "60", "0 (CA)", "CA only"],
    ["COMPLEMENT_ALIM.", "3", "6", "3", "3", "Stable (BELGOKILL proxy)"],
    ["MAÏS (exclu)", "0", "0", "0", "0", "—"],
    ["TOTAL", "28 449", "63 228", "52 079", "118 608", "+52% vs 2026 ann."],
]
story.append(make_table(hist_detail, col_widths=[3*cm, 2*cm, 2*cm, 2.2*cm, 2.2*cm, 3.5*cm], font_size=8, highlight_rows=[6]))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph(
    "<b>Analyse</b> : Le TOURTEAUX montre une trajectoire de croissance soutenue (19 296 t en S2 2024 → 44 584 t en 2025 → "
    "39 045 t YTD 2026 sur 8 mois → 92 517 t forecast 2027). Cette tendance est amplifiée par l'effet prix (25 000 FCFA/sac) "
    "qui propulse le CA à 46 261 M FCFA. Les CONCENTRÉS progressent également (+38% vs 2026 annualisé), portés par le "
    "maintien du bundle 2,3:1. La famille COMPLEMENT_ALIMENTAIRE reste stable, avec BELGOKILL comme produit phare (38% du CA famille).",
    BODY))

story.append(Paragraph("5. Plan de déploiement", H1))
deploy_data = [
    ["Étape", "Date", "Action", "Cible"],
    ["1", "03/09/2026", "Validation proposition v2", "DG"],
    ["2", "05/09/2026", "Livrables finaux prêts", "DA"],
    ["3", "06/09/2026", "Présentation CODIR", "DG + DC + DP + CG"],
    ["4", "07/09/2026", "Diffusion aux agences", "14 RA"],
    ["5", "Mars 2027", "Mise à jour Q1 avec données réelles", "DA"],
    ["6", "Juin 2027", "Mise à jour Q2", "DA"],
    ["7", "Sept 2027", "Mise à jour Q3", "DA"],
]
story.append(make_table(deploy_data, col_widths=[1*cm, 3*cm, 7*cm, 6*cm], font_size=9))

story.append(Paragraph("6. Conclusion", H1))
story.append(Paragraph(
    "Le forecast 2027 version 2 projette <b>118 608 tonnes</b> pour <b>64 114 M FCFA</b>, avec une concentration "
    "saisonnière au Q4 (39%). La désaisonnalisation de l'effet soja, l'actualisation du prix à 25 000 FCFA, "
    "l'ajout de la famille COMPLEMENT_ALIMENTAIRE (proxy BELGOKILL) et l'utilisation des données 2024-2026 (32 mois) "
    "permettent une projection réaliste. Le pic d'octobre (16 355 t) nécessitera une anticipation renforcée "
    "du réapprovisionnement soja dès septembre 2027.",
    BODY))

doc.build(story)
print(f"✓ Document stratégique: {strat_path} ({os.path.getsize(strat_path)/1024:.0f} KB)")

# ==================== 5. GUIDE MÉTHODOLOGIQUE ====================
print("Generating 5. Guide méthodologique 2027 (v2)...")
guide_path = f"{OUT_DIR}/05_guide_methodologique.pdf"
doc = SimpleDocTemplate(guide_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []

story.extend(cover_page("Guide Méthodologique", "Forecast 2027 - Version 2 (Prophet + COMPLEMENT_ALIMENTAIRE)", "GUIDE MÉTHODOLOGIQUE"))

story.append(Paragraph("Sommaire", H1))
story.append(Paragraph(
    "1. Introduction et contexte<br/>"
    "2. Méthodologie PACE<br/>"
    "3. Outils et environnement<br/>"
    "4. Phase PREPARE - Données 2024-2026<br/>"
    "5. Phase PREPARE - Désaisonnalisation soja<br/>"
    "6. Phase PREPARE - En cours + Validées<br/>"
    "7. NOUVEAU: Phase PREPARE - COMPLEMENT_ALIMENTAIRE<br/>"
    "8. Phase ANALYZE - AED<br/>"
    "9. Phase CONSTRUCT - Prophet<br/>"
    "10. Phase EXECUTE - Livrables<br/>"
    "Annexes",
    BODY))

story.append(PageBreak())
story.append(Paragraph("1. Introduction et contexte", H1))
story.append(Paragraph(
    "Ce guide décrit la méthodologie complète du forecast 2027 de BELGOCAM SA (version 2). Il couvre 12 mois "
    "(janvier-décembre 2027) avec 7 familles de produits (incluant la nouvelle famille COMPLEMENT_ALIMENTAIRE) "
    "et 14 agences. Les innovations majeures vs la version 1 sont : "
    "(1) ajout de la famille COMPLEMENT_ALIMENTAIRE avec proxy BELGOKILL et conversion 1L=1kg, "
    "(2) utilisation des données 2024-2026 (32 mois) au lieu de 2025-2026 (20 mois), "
    "(3) désaisonnalisation de l'effet soja exceptionnel, "
    "(4) inclusion des commandes En cours/Validées, "
    "(5) actualisation du prix soja à 25 000 FCFA.",
    BODY))

story.append(Paragraph("2. Méthodologie PACE", H1))
story.append(Paragraph(
    "La méthodologie PACE (Prepare-Analyze-Construct-Execute) structure le projet en 4 phases successives, "
    "chacune avec des livrables vérifiables.",
    BODY))
pace_detail = [
    ["Phase", "Objectif", "Livrables", "Durée"],
    ["P - PREPARE", "Préparer données 2024-2026 + désaisonnalisation + COMPLEMENT_ALIM.", "Dataset 170 560 records, prix 2027", "2 jours"],
    ["A - ANALYZE", "Comprendre données + effet soja", "Graphiques, synthèse AED", "1 jour"],
    ["C - CONSTRUCT", "Modéliser Prophet 12 mois (7 familles)", "Forecast 2027 (8 868 lignes)", "2 jours"],
    ["E - EXECUTE", "Produire livrables finaux", "Excel + 5 PDFs", "1 jour"],
]
story.append(make_table(pace_detail, col_widths=[3*cm, 4*cm, 5*cm, 3*cm], font_size=9))

story.append(Paragraph("3. Outils et environnement", H1))
stack_data = [
    ["Outil", "Version", "Usage"],
    ["Python", "3.12", "Langage principal"],
    ["pandas", "2.x", "Manipulation données"],
    ["Prophet", "1.1+", "Modélisation prédictive (5 familles × 3 régions = 15 modèles)"],
    ["openpyxl", "3.x", "Génération Excel (8 feuilles)"],
    ["ReportLab", "4.x", "Génération PDFs (5 documents PACE)"],
    ["matplotlib", "3.x", "Visualisations"],
]
story.append(make_table(stack_data, col_widths=[4*cm, 3*cm, 8*cm], font_size=9))

story.append(PageBreak())
story.append(Paragraph("4. Phase PREPARE - Données 2024-2026", H1))
story.append(Paragraph(
    "Le dataset consolidé couvre 32 mois (juillet 2024 - août 2026), soit 170 560 enregistrements. "
    "Les sources sont : LY_24 (Jul-Dec 2024, 52 215 records), Historique 2025 (66 206 records), "
    "S1 2026 (37 745 records), Juillet 2026 (7 537 records), Août 2026 (6 857 records), En cours + Validées (~255).",
    BODY))
story.append(Paragraph(
    "<b>Exclusion des années 2021-2023</b> : La version 2 se concentre sur les données récentes (2024-2026) "
    "pour capturer la tendance actuelle du marché. Les années 2021-2023 sont exclues car les conditions de marché "
    "ont significativement changé (notamment avec la hausse du soja et la rupture concurrente de 2026).",
    BODY))

story.append(Paragraph("5. Phase PREPARE - Désaisonnalisation soja", H1))
story.append(Paragraph(
    "L'effet soja exceptionnel de juillet-août 2026 (rupture concurrente) a généré des volumes anormalement "
    "élevés qui ne sont pas récurrents. Pour éviter que Prophet n'extrapole ce phénomène, les volumes soja "
    "de Jul-Août 2026 ont été cappés à la moyenne S1 2026 (Jan-Juin 2026).",
    BODY))
story.append(Paragraph("<b>Méthode</b>", H3))
story.append(Paragraph(
    "1. Calcul de la moyenne mensuelle soja S1 2026 (Jan-Juin 2026)<br/>"
    "2. Identification du mois Jul-Août 2026 avec le volume le plus élevé<br/>"
    "3. Calcul du facteur de cap = moyenne S1 / volume max Jul-Août<br/>"
    "4. Application du facteur à tous les enregistrements soja Jul-Août 2026<br/>"
    "5. Les volumes sont réduits proportionnellement, préservant le signal de saisonnalité",
    BODY))

story.append(Paragraph("6. Phase PREPARE - En cours + Validées", H1))
story.append(Paragraph(
    "Les commandes En cours et Validées d'août 2026 sont incluses dans le dataset comme "
    "potentielles ventes Livrées. Les commandes Annulées et Brouillon sont exclues. "
    "Cette intégration permet de capturer le pipeline commercial en cours au moment de l'extraction.",
    BODY))

story.append(PageBreak())
story.append(Paragraph("7. NOUVEAU: Phase PREPARE - Famille COMPLEMENT_ALIMENTAIRE", H1))
story.append(Paragraph(
    "La famille COMPLEMENT_ALIMENTAIRE regroupe 10 produits liquides de la gamme BELGO : "
    "BELGOKILL (V300 1L, V305 200L), BELGO HARMONY, BELGO PROTECT, BELGO DRY LIT, BELGO WATER CLEAN, "
    "BELGO VIT Ese, BELGO THERMO, BELGO BIO SELECT, BELGO FRESH. Ces produits sont conditionnés en bidons "
    "de 1L (ou 200L pour BELGOKILL V305).",
    BODY))

story.append(Paragraph("<b>Conversion 1L = 1kg</b>", H3))
story.append(Paragraph(
    "Pour exprimer les volumes en tonnes (nécessaire pour Prophet), la conversion 1L = 1kg est appliquée : "
    "<b>1 qte = 1 kg, tonnes = qte / 1000</b>. Cette approximation est valable pour les liquides aqueux "
    "(densité ~1). Pour V305 (BELGOKILL 200L), 1 qte = 200 kg (= 200L × 1 kg/L).",
    BODY))

story.append(Paragraph("<b>Proxy BELGOKILL</b>", H3))
story.append(Paragraph(
    "BELGOKILL (V300 1L + V305 200L) représente <b>38% du CA famille</b> sur 2024-2026 (20,4 M FCFA sur 54 M FCFA total). "
    "La tendance BELGOKILL est utilisée comme proxy pour modéliser la famille entière via Prophet. "
    "Les autres produits (BELGO HARMONY, BELGO PROTECT, etc.) sont ensuite désagrégés selon leurs parts historiques.",
    BODY))

ca_proxy = [
    ["Réf.", "Produit", "Poids (kg/qte)", "Prix 2027 (FCFA/L)", "Part CA famille"],
    ["V300", "BELGOKILL 1L", "1", "2 500", "38%"],
    ["V305", "BELGOKILL 200L", "200", "500 000", "—"],
    ["CA003.1", "BELGO HARMONY 1L", "1", "8 000", "16%"],
    ["CA004.1", "BELGO PROTECT 1L", "1", "9 800", "7%"],
    ["CA006.1", "BELGO DRY LIT 1L", "1", "6 500", "3%"],
    ["CA001.1", "BELGO WATER CLEAN 1L", "1", "5 500", "3%"],
    ["CA002.1", "BELGO VIT Ese 1L", "1", "10 500", "14%"],
    ["CA005.1", "BELGO THERMO 1L", "1", "15 000", "8%"],
    ["CA007.1", "BELGO BIO SELECT 1L", "1", "8 000", "5%"],
    ["CA008.1", "BELGO FRESH 1L", "1", "14 000", "5%"],
]
story.append(make_table(ca_proxy, col_widths=[2*cm, 5*cm, 3*cm, 3.5*cm, 3*cm], font_size=9))

story.append(Paragraph("8. Phase ANALYZE - AED", H1))
story.append(Paragraph(
    "L'AED révèle : (1) TOURTEAUX = 76% du volume historique, (2) Q4 = 35-46% du volume annuel selon les familles, "
    "(3) pic octobre (indice 1.86 pour TOURTEAUX en 2025), (4) FAMLA = 31.6% du volume, "
    "(5) ratio bundle 2,3:1 en août 2026, (6) COMPLEMENT_ALIMENTAIRE = 3-6 t/an (faible volume, forte valeur unitaire).",
    BODY))

story.append(Paragraph("9. Phase CONSTRUCT - Prophet", H1))
story.append(Paragraph("9.1 Configuration", H3))
story.append(Paragraph(
    "Prophet est configuré avec yearly_seasonality=True (saisonnalité annuelle), seasonality_mode='multiplicative' "
    "(saisonnalité proportionnelle), changepoint_prior_scale=0.05 (tendance modérée), interval_width=0.8. "
    "Les future dates sont explicites (Jan-Déc 2027) pour garantir 12 mois de forecast.",
    BODY))

story.append(Paragraph("9.2 Stratégie de modélisation", H3))
story.append(Paragraph(
    "15 modèles Prophet famille × région (5 familles Prophet × 3 régions) + extrapolation pour MATERIEL_ELEVAGE et "
    "PREMIX (2 familles × 3 régions = 6 extrapolations). Désagrégation par produit × agence selon les parts "
    "historiques de CA. Prix 2027 : 25 000 FCFA/sac soja, prix septembre 2026 pour les autres familles, "
    "prix par litre pour COMPLEMENT_ALIMENTAIRE (2 500 à 15 000 FCFA/L selon produit).",
    BODY))

story.append(Paragraph("10. Phase EXECUTE - Livrables", H1))
livrables_data = [
    ["#", "Livrable", "Format", "Usage"],
    ["1", "Résumé exécutif", "PDF", "Direction Générale"],
    ["2", "Proposition de projet", "PDF", "CODIR"],
    ["3", "Matrice RACI", "PDF", "Gouvernance"],
    ["4", "Document stratégique PACE", "PDF", "Référence stratégique"],
    ["5", "Guide méthodologique", "PDF", "Référence méthodes"],
    ["6", "Excel forecast 2027 (8 feuilles)", "XLSX", "Pilotage opérationnel"],
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
    ["Cap S1", "Plafonnement des volumes Jul-Août 2026 à la moyenne S1 2026"],
    ["Q4", "Quatrième trimestre (oct-déc) = 39% du CA 2027"],
    ["PACE", "Prepare-Analyze-Construct-Execute"],
    ["RACI", "Responsible-Accountable-Consulted-Informed"],
    ["NOUVEAU: COMPLEMENT_ALIM.", "Famille ajoutée en v2: 10 produits liquides BELGOxxx (1L=1kg)"],
    ["NOUVEAU: Proxy BELGOKILL", "Tendance BELGOKILL (V300, 38% du CA famille) utilisée comme proxy"],
    ["NOUVEAU: 1L=1kg", "Conversion appliquée pour exprimer les volumes liquides en tonnes"],
]
story.append(make_table(glossaire, col_widths=[5*cm, 11*cm], font_size=9, highlight_rows=[9, 10, 11]))

doc.build(story)
print(f"✓ Guide méthodologique: {guide_path} ({os.path.getsize(guide_path)/1024:.0f} KB)")

print("\n=== TOUS LES 5 PDFs 2027 V2 GÉNÉRÉS ===")
for f in sorted(os.listdir(OUT_DIR)):
    if f.endswith('.pdf'):
        print(f"  {f}: {os.path.getsize(os.path.join(OUT_DIR, f))/1024:.0f} KB")
