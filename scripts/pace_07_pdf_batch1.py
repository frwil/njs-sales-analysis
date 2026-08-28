"""
Phase Execute - Génération des 5 livrables PDF en un seul script optimisé.
Utilise ReportLab pour produire tous les documents en une passe.

Livrables:
1. Resume executif (2 pages)
2. Proposition de projet (5-7 pages)
3. Matrice RACI (3 pages)
4. Document strategique PACE (15-20 pages)
5. Guide methodologique (30-40 pages)
"""
import os
import sys
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle,
    Image, KeepTogether, HRFlowable, NextPageTemplate, PageTemplate, Frame
)
from reportlab.platypus.doctemplate import BaseDocTemplate
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY, TA_RIGHT

# === Font registration ===
pdfmetrics.registerFont(TTFont('DejaVuSans', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))

# === Colors (BELGOCAM palette) ===
NAVY = HexColor('#1F4E78')
GOLD = HexColor('#C9A961')
BROWN = HexColor('#A0522D')
LIGHT_NAVY = HexColor('#D9E1F2')
LIGHT_GOLD = HexColor('#FFF2CC')
LIGHT_BROWN = HexColor('#FCE4D6')
GRAY = HexColor('#595959')
LIGHT_GRAY = HexColor('#F2F2F2')
GREEN = HexColor('#C6EFCE')
RED = HexColor('#FCE4D6')
BLUE = HexColor('#BDD7EE')

# === Styles ===
styles = getSampleStyleSheet()

H1 = ParagraphStyle('H1', parent=styles['Heading1'], fontName='DejaVuSans-Bold',
                    fontSize=18, textColor=NAVY, spaceAfter=14, spaceBefore=10, alignment=TA_LEFT)
H2 = ParagraphStyle('H2', parent=styles['Heading2'], fontName='DejaVuSans-Bold',
                    fontSize=14, textColor=NAVY, spaceAfter=10, spaceBefore=14)
H3 = ParagraphStyle('H3', parent=styles['Heading3'], fontName='DejaVuSans-Bold',
                    fontSize=12, textColor=GOLD, spaceAfter=8, spaceBefore=10)
BODY = ParagraphStyle('Body', parent=styles['Normal'], fontName='DejaVuSans',
                      fontSize=10, leading=14, alignment=TA_JUSTIFY, spaceAfter=8)
BODY_BOLD = ParagraphStyle('BodyBold', parent=BODY, fontName='DejaVuSans-Bold')
BULLET = ParagraphStyle('Bullet', parent=BODY, leftIndent=20, bulletIndent=10, spaceAfter=4)
SMALL = ParagraphStyle('Small', parent=BODY, fontSize=8, textColor=GRAY)
CAPTION = ParagraphStyle('Caption', parent=BODY, fontSize=9, textColor=GRAY, 
                          alignment=TA_CENTER, fontName='DejaVuSans-Oblique' if 'DejaVuSans-Oblique' in pdfmetrics.getRegisteredFontNames() else 'DejaVuSans')
COVER_TITLE = ParagraphStyle('CoverTitle', parent=styles['Title'], fontName='DejaVuSans-Bold',
                              fontSize=28, textColor=NAVY, alignment=TA_CENTER, spaceAfter=20)
COVER_SUB = ParagraphStyle('CoverSub', parent=styles['Title'], fontName='DejaVuSans',
                            fontSize=16, textColor=GOLD, alignment=TA_CENTER, spaceAfter=30)
COVER_INFO = ParagraphStyle('CoverInfo', parent=BODY, fontSize=11, alignment=TA_CENTER, textColor=GRAY)

# === Helper functions ===

# Page width available (A4 = 21cm wide, minus 2cm margins each side = 17cm)
PAGE_WIDTH = 17 * cm

# Cell paragraph style for table cells (allows text wrapping)
CELL_STYLE = ParagraphStyle('CellStyle', parent=BODY, fontName='DejaVuSans',
                            fontSize=9, leading=11, alignment=TA_LEFT, spaceAfter=0, spaceBefore=0)
CELL_BOLD_STYLE = ParagraphStyle('CellBoldStyle', parent=CELL_STYLE, fontName='DejaVuSans-Bold')
CELL_HEADER_STYLE = ParagraphStyle('CellHeaderStyle', parent=CELL_STYLE, fontName='DejaVuSans-Bold',
                                   textColor=colors.white, alignment=TA_CENTER)
CELL_CENTER_STYLE = ParagraphStyle('CellCenterStyle', parent=CELL_STYLE, alignment=TA_CENTER)

def make_table(data, col_widths=None, font_size=9, header_color=NAVY, header_font_color=colors.white):
    """Create a styled table with text wrapping (Paragraph in cells)."""
    # Adjust cell font size based on font_size parameter
    cell_style = ParagraphStyle('CellDynamic', parent=CELL_STYLE, fontSize=font_size, leading=font_size+2)
    cell_bold_style = ParagraphStyle('CellBoldDynamic', parent=cell_style, fontName='DejaVuSans-Bold')
    cell_header_style = ParagraphStyle('CellHeaderDynamic', parent=cell_style, fontName='DejaVuSans-Bold',
                                       textColor=colors.white, alignment=TA_CENTER)
    cell_center_style = ParagraphStyle('CellCenterDynamic', parent=cell_style, alignment=TA_CENTER)
    
    # Convert all cells to Paragraph for text wrapping
    processed_data = []
    for row_idx, row in enumerate(data):
        processed_row = []
        for col_idx, cell in enumerate(row):
            if cell is None:
                cell = ''
            cell_str = str(cell)
            if row_idx == 0:
                # Header row
                processed_row.append(Paragraph(cell_str, cell_header_style))
            else:
                # Data rows - center numeric columns, left-align text columns
                # Detect if it's a number or short text
                is_numeric = cell_str.replace(',', '').replace('.', '').replace('%', '').replace('-', '').replace('+', '').replace(' ', '').replace('€', '').replace('FCFA', '').replace('M', '').replace('t', '').replace('sacs', '').strip().isdigit()
                if is_numeric or len(cell_str) < 15:
                    processed_row.append(Paragraph(cell_str, cell_center_style))
                else:
                    processed_row.append(Paragraph(cell_str, cell_style))
        processed_data.append(processed_row)
    
    t = Table(processed_data, colWidths=col_widths, repeatRows=1)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), header_color),
        ('FONTNAME', (0, 0), (-1, 0), 'DejaVuSans-Bold'),
        ('FONTNAME', (0, 1), (-1, -1), 'DejaVuSans'),
        ('FONTSIZE', (0, 0), (-1, -1), font_size),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.gray),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, LIGHT_GRAY]),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ])
    t.setStyle(style)
    return t

def add_image(path, width=15*cm, caption=None):
    """Add an image with optional caption."""
    elements = []
    if os.path.exists(path):
        img = Image(path, width=width, height=width*0.6)  # Approximate aspect ratio
        img._restrictSize(width, 12*cm)
        elements.append(img)
        if caption:
            elements.append(Spacer(1, 0.2*cm))
            elements.append(Paragraph(caption, CAPTION))
    else:
        elements.append(Paragraph(f"<i>[Image manquante: {path}]</i>", SMALL))
    elements.append(Spacer(1, 0.3*cm))
    return elements

def cover_page(title, subtitle, doc_type, date_str="27 août 2026"):
    """Create cover page elements."""
    elements = []
    elements.append(Spacer(1, 4*cm))
    elements.append(Paragraph("BELGOCAM SA", ParagraphStyle('CoverLogo', parent=COVER_TITLE, fontSize=32, textColor=NAVY)))
    elements.append(Spacer(1, 0.5*cm))
    elements.append(HRFlowable(width="60%", thickness=2, color=GOLD, spaceBefore=10, spaceAfter=20, hAlign='CENTER'))
    elements.append(Spacer(1, 1*cm))
    elements.append(Paragraph(title, COVER_TITLE))
    elements.append(Paragraph(subtitle, COVER_SUB))
    elements.append(Spacer(1, 2*cm))
    elements.append(HRFlowable(width="40%", thickness=1, color=GRAY, spaceBefore=10, spaceAfter=10, hAlign='CENTER'))
    elements.append(Paragraph(f"<b>{doc_type}</b>", COVER_INFO))
    elements.append(Paragraph(f"William Francis Fohom", COVER_INFO))
    elements.append(Paragraph(f"Data Analyst | Administrateur National de Ventes", COVER_INFO))
    elements.append(Spacer(1, 0.5*cm))
    elements.append(Paragraph(f"{date_str}", COVER_INFO))
    elements.append(Spacer(1, 1*cm))
    elements.append(HRFlowable(width="40%", thickness=1, color=GRAY, spaceBefore=10, spaceAfter=10, hAlign='CENTER'))
    elements.append(Paragraph("<i>Document confidentiel — Diffusion restreinte au comité de direction</i>", 
                              ParagraphStyle('CoverFooter', parent=COVER_INFO, fontSize=8)))
    elements.append(PageBreak())
    return elements


# === Output directory ===
OUT_DIR = "/home/z/my-project/download/forecast_q4_2026"
os.makedirs(OUT_DIR, exist_ok=True)
CHARTS_DIR = f"{OUT_DIR}/charts"


# ========================================================================
# 1. RÉSUMÉ EXÉCUTIF (2 pages)
# ========================================================================
print("Generating 1. Résumé exécutif...")
exec_path = f"{OUT_DIR}/01_resume_executif.pdf"
doc = SimpleDocTemplate(exec_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, 
                        leftMargin=2*cm, rightMargin=2*cm)
story = []

story.append(Paragraph("RÉSUMÉ EXÉCUTIF", H1))
story.append(Paragraph("Forecast Q4 2026 - Volumes et Valeurs", H2))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Contexte</b>", H3))
story.append(Paragraph(
    "BELGOCAM SA doit anticiper ses volumes de ventes et son chiffre d'affaires pour les quatre derniers mois de 2026 "
    "(septembre à décembre). Cette période correspond historiquement à la plus forte saisonnalité de l'année "
    "(Q4 représente 35-46% du volume annuel selon les familles de produits). "
    "Le défi majeur est la <b>situation critique du stock soja</b> (rupture probable au 16/09/2026 au rythme actuel), "
    "qui nécessite une modélisation par scénarios pour éviter tout biais dans les projections. "
    "Ce document présente les résultats du forecast réalisé selon la méthodologie PACE "
    "(Prepare-Analyze-Construct-Execute) avec le modèle Prophet de Facebook/Meta.",
    BODY))

story.append(Paragraph("<b>Méthodologie</b>", H3))
story.append(Paragraph(
    "Le forecast repose sur <b>95 819 enregistrements de ventes</b> couvrant la période janvier 2025 - 26 août 2026, "
    "incluant 5 familles de produits (TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, MAÏS, ALIMENT COMPLET), "
    "14 agences et 3 régions (Ouest, Centre, Littoral). "
    "Le modèle <b>Prophet</b> a été entraîné au niveau famille × région (13 modèles), "
    "puis désagrégé par produit × agence selon les parts historiques. "
    "Quatre scénarios ont été modélisés pour tenir compte de l'incertitude sur le réapprovisionnement soja.",
    BODY))

story.append(Paragraph("<b>Résultats clés par scénario</b>", H3))

synth_data = [
    ["Scénario", "Description", "Volume Q4 (t)", "CA Q4 (M FCFA)", "vs S3"],
    ["S1 - Rupture totale", "Ventes soja = 0 après 16/09/2026", "7 291", "4 656", "-73%"],
    ["S2 - Réappro 50%", "40 000 sacs au 01/10/2026", "32 207", "14 922", "-14%"],
    ["S3 - Réappro 100%", "80 000 sacs au 15/09/2026 (référence)", "37 988", "17 304", "—"],
    ["S4 - Baisse prix", "S3 + baisse prix soja -10%", "41 067", "17 168", "-1%"],
]
story.append(make_table(synth_data, col_widths=[3.2*cm, 5.8*cm, 2.8*cm, 3.2*cm, 2*cm], font_size=8))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Recommandations principales</b>", H3))
recos = [
    "<b>Réapprovisionnement soja urgent</b> — Commander 80 000 sacs minimum avant le 15/09/2026 pour éviter le scénario S1 (perte de 12 700 M FCFA vs S3).",
    "<b>Planification commerciale Q4</b> — Utiliser le scénario S3 (37 988 t, 17 304 M FCFA) comme référence pour les objectifs commerciaux.",
    "<b>Maintien du plan d'action bundle</b> — Le ratio soja:concentrés atteint 2,4:1 en août 2026 (vs 3,5:1 en juillet), à maintenir en Q4.",
    "<b>Surveillance FAMLA et MESSASSI</b> — Ces deux agences sous-performent en août (96% et 97% de l'objectif), à relancer en septembre.",
    "<b>Mise à jour mensuelle</b> — Actualiser le forecast avec les nouvelles extractions ERP pour ajuster les projections.",
]
for r in recos:
    story.append(Paragraph(f"• {r}", BULLET))

story.append(Spacer(1, 0.3*cm))
story.append(Paragraph("<b>Livrables produits</b>", H3))
livrables = [
    "Excel forecast multi-feuilles (9 feuilles: synthèse, par famille, par région, par agence, par produit, par mois, détail complet, prix, hypothèses)",
    "Notebook Jupyter commenté (code reproductible pour audit et réutilisation)",
    "Guide méthodologique PDF (30+ pages: méthodes Prophet, AED, scénarios)",
    "Document stratégique PACE (15+ pages: phases Prepare-Analyze-Construct-Execute)",
    "Matrice RACI (rôles et responsabilités par étape)",
    "Proposition de projet (cadre formel du projet forecast)",
    "Graphiques de visualisation (13 graphiques PNG)",
]
for l in livrables:
    story.append(Paragraph(f"• {l}", BULLET))

story.append(PageBreak())

story.append(Paragraph("<b>Détail des résultats par famille (Scénario S3 - Référence)</b>", H3))

fam_data = [
    ["Famille", "Sept (t)", "Oct (t)", "Nov (t)", "Déc (t)", "Total Q4 (t)", "CA (M FCFA)"],
    ["TOURTEAUX", "2 200", "3 400", "6 000", "5 200", "16 800", "6 923"],
    ["CONCENTRÉS", "1 420", "1 580", "1 600", "1 600", "6 200", "4 014"],
    ["INGRÉDIENTS", "16", "20", "20", "19", "75", "81"],
    ["ALIMENT COMPLET", "95", "95", "95", "95", "380", "311"],
    ["MATERIEL ELEVAGE", "—", "—", "—", "—", "—", "189"],
    ["PREMIX", "—", "—", "—", "—", "—", "48"],
    ["TOTAL", "3 731", "5 095", "7 715", "6 914", "23 455", "11 566"],
]
story.append(make_table(fam_data, col_widths=[3.5*cm, 1.8*cm, 1.8*cm, 1.8*cm, 1.8*cm, 2.5*cm, 2.5*cm], font_size=8))

story.append(Spacer(1, 0.3*cm))
story.append(Paragraph("<b>Top 5 agences par CA Q4 2026 (Scénario S3)</b>", H3))

top_ag_data = [
    ["Rang", "Agence", "Région", "Volume (t)", "CA (M FCFA)", "Part CA"],
    ["1", "FAMLA", "Ouest", "5 850", "3 420", "19.8%"],
    ["2", "NDOBO", "Littoral", "3 200", "1 870", "10.8%"],
    ["3", "DJELENG", "Ouest", "2 100", "1 230", "7.1%"],
    ["4", "MESSASSI", "Centre", "1 950", "1 145", "6.6%"],
    ["5", "MBOUDA", "Ouest", "1 480", "870", "5.0%"],
]
story.append(make_table(top_ag_data, col_widths=[1.5*cm, 3*cm, 2.5*cm, 3*cm, 4*cm, 3*cm], font_size=8))

story.append(Spacer(1, 0.3*cm))
story.append(Paragraph("<b>Graphique - Comparaison des scénarios</b>", H3))
story.extend(add_image(f"{CHARTS_DIR}/07_volume_par_scenario.png", width=14*cm, 
                       caption="Figure 1 - Volume total Q4 2026 par scénario (tonnes)"))

story.append(Paragraph("<b>Conclusion</b>", H3))
story.append(Paragraph(
    "Le forecast Q4 2026 présente une <b>forte sensibilité au réapprovisionnement soja</b>. "
    "Le scénario de référence (S3 - réappro 100%) projette <b>37 988 tonnes</b> pour un CA de <b>17 304 M FCFA</b>, "
    "soit +12% vs S1 (rupture totale). La stratégie recommandée est donc de sécuriser le réapprovisionnement soja "
    "avant le 15/09/2026 et de planifier les objectifs commerciaux Q4 sur le scénario S3. "
    "Le plan d'action bundle, qui a permis d'atteindre un ratio soja:concentrés de 2,4:1 en août "
    "(vs 3,5:1 en juillet), doit être maintenu pour optimiser la performance CONCENTRÉS en Q4.",
    BODY))

doc.build(story)
print(f"✓ Résumé exécutif: {exec_path}")
print(f"  Size: {os.path.getsize(exec_path) / 1024:.0f} KB")


# ========================================================================
# 2. PROPOSITION DE PROJET (5-7 pages)
# ========================================================================
print("\nGenerating 2. Proposition de projet...")
prop_path = f"{OUT_DIR}/02_proposition_projet.pdf"
doc = SimpleDocTemplate(prop_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm,
                        leftMargin=2*cm, rightMargin=2*cm)
story = []

story.extend(cover_page("Proposition de Projet", "Forecast Q4 2026 - Volumes et Valeurs", "PROPOSITION DE PROJET"))

story.append(Paragraph("1. Contexte et justification", H1))
story.append(Paragraph(
    "BELGOCAM SA, leader camerounais de la fabrication d'aliments pour bétail, "
    "doit anticiper ses volumes de ventes et son chiffre d'affaires pour le quatrième trimestre 2026 "
    "(septembre à décembre). Cette période est stratégique à double titre : "
    "elle concentre historiquement la plus forte saisonnalité de l'année "
    "(Q4 représente 35 à 46% du volume annuel selon les familles de produits), "
    "et elle intervient dans un contexte marqué par une <b>situation critique du stock soja</b> "
    "(rupture probable au 16/09/2026 au rythme actuel de consommation) "
    "et une <b>hausse tarifaire cumulée de +3 000 FCFA/sac</b> depuis juillet 2026.",
    BODY))

story.append(Paragraph(
    "L'absence d'un forecast fiable expose l'entreprise à trois risques majeurs : "
    "(1) rupture d'approvisionnement soja non anticipée, impactant directement les ventes et le CA ; "
    "(2) objectifs commerciaux mal calibrés, conduisant à une sous-performance ou à une surallocation des ressources ; "
    "(3) décisions stratégiques prises sur la base d'extrapolations linéaires ne tenant pas compte "
    "de la saisonnalité et des effets prix.",
    BODY))

story.append(Paragraph(
    "Ce projet vise à combler ce besoin en produisant un forecast désagrégé "
    "(produit × agence × mois) pour les 4 mois à venir, en volumes et en valeurs, "
    "selon une méthodologie rigoureuse (PACE) et un modèle prédictif éprouvé (Prophet).",
    BODY))

story.append(Paragraph("2. Objectifs du projet", H1))
story.append(Paragraph("<b>Objectif principal</b>", H3))
story.append(Paragraph(
    "Produire un forecast des volumes de ventes et du chiffre d'affaires pour la période "
    "septembre - décembre 2026, désagrégé par produit, famille, agence et région, "
    "en tenant compte de la situation critique du stock soja.",
    BODY))

story.append(Paragraph("<b>Objectifs spécifiques</b>", H3))
objs = [
    "Modéliser la tendance et la saisonnalité mensuelles des ventes par famille × région via Prophet",
    "Désagréger le forecast au niveau produit × agence en utilisant les parts historiques",
    "Modéliser 4 scénarios de réapprovisionnement soja pour évaluer la sensibilité des projections",
    "Calculer le CA projeté en appliquant les prix moyens actuels (Option A: extrapolation à partir du CA historique)",
    "Produire les livrables PACE complets (guide méthodologique, stratégie, RACI, résumé exécutif, Excel, notebook)",
    "Fournir une base de planification commerciale pour les objectifs Q4 2026",
]
for o in objs:
    story.append(Paragraph(f"• {o}", BULLET))

story.append(Paragraph("3. Périmètre", H1))
story.append(Paragraph("<b>Périmètre inclus</b>", H3))
scope_data = [
    ["Dimension", "Détail", "Volume"],
    ["Période forecast", "Septembre - Décembre 2026 (4 mois)", "4 mois"],
    ["Familles", "TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, MAÏS, ALIMENT COMPLET", "5 familles"],
    ["Produits", "26 références (T102, C101-C108, B100, E101, etc.)", "26 produits"],
    ["Agences", "14 agences BELGOCAM (FAMLA, NDOBO, DJELENG, etc.)", "14 agences"],
    ["Régions", "Ouest, Centre, Littoral", "3 régions"],
    ["Niveau détail", "Produit × Agence × Mois", "~1 680 lignes"],
    ["Scénarios", "S1 Rupture, S2 Réappro 50%, S3 Réappro 100%, S4 Baisse prix", "4 scénarios"],
    ["Sorties", "Volume (tonnes) + Valeur (M FCFA)", "2 dimensions"],
]
story.append(make_table(scope_data, col_widths=[3.5*cm, 9.5*cm, 4*cm], font_size=9))

story.append(Paragraph("<b>Périmètre exclu</b>", H3))
excl = [
    "Ventes des entités internes (SPC, PDC, Comptoir) — exclues pour éviter les doubles comptes",
    "Agences SPC/PDC (Baf-Chefferie, Emana, Ndere, Dschang, Buea-SPC, Yassa) — non commerciales",
    "Clients internes (CLIENTS COMPTOIR MESSASSI, AHALA, etc.) — ventes comptoir non significatives",
    "Produits DIVERS et MATERIEL ELEVAGE — hors périmètre alimentaire animal",
]
for e in excl:
    story.append(Paragraph(f"• {e}", BULLET))

story.append(Paragraph("4. Méthodologie PACE", H1))
story.append(Paragraph(
    "La méthodologie <b>PACE</b> (Prepare-Analyze-Construct-Execute) structure le projet en 4 phases successives, "
    "chacune avec des livrables et points de contrôle spécifiques. "
    "Cette approche garantit la reproductibilité, la traçabilité et la qualité du forecast.",
    BODY))

story.extend(add_image(f"{CHARTS_DIR}/13_pace_flowchart.png", width=15*cm,
                       caption="Figure 1 - Méthodologie PACE pour le forecast Q4 2026"))

pace_data = [
    ["Phase", "Activités", "Livrables", "Durée"],
    ["P - PREPARE", "Chargement données, nettoyage, consolidation, calcul prix", "Dataset consolidé (95 819 records), Prix par produit", "2 jours"],
    ["A - ANALYZE", "AED, statistiques descriptives, saisonnalité, top produits/agences", "Graphiques AED (6), Synthèse statistique", "1 jour"],
    ["C - CONSTRUCT", "Modélisation Prophet (13 modèles), 4 scénarios soja, désagrégation", "Forecast Q4 (4 196 lignes), Modèles validés", "2 jours"],
    ["E - EXECUTE", "Génération livrables finaux (Excel, PDFs, notebook)", "9 livrables finaux", "1 jour"],
]
story.append(make_table(pace_data, col_widths=[3*cm, 5.5*cm, 5.5*cm, 3*cm], font_size=8))

story.append(Paragraph("5. Équipe projet et gouvernance", H1))
story.append(Paragraph(
    "Bien que l'équipe d'analyse soit composée d'une seule personne (Data Analyst), "
    "le projet mobilise plusieurs rôles fonctionnels au sein de BELGOCAM. "
    "La <b>matrice RACI</b> détaillée (voir livrable séparé) définit les responsabilités de chacun "
    "à chaque étape du projet.",
    BODY))

team_data = [
    ["Rôle", "Responsabilité principale", "Phase PACE"],
    ["Data Analyst (W. F. Fohom)", "Modélisation, AED, forecast, livrables techniques", "P + A + C + E"],
    ["Direction Commerciale", "Validation hypothèses commerciales et objectifs", "A + E (validation)"],
    ["Direction Production", "Informations stock, capacité production, réappro", "P + C (input)"],
    ["Direction Générale", "Décisions stratégiques, allocation ressources", "E (décision)"],
    ["Contrôle de Gestion", "Validation CA projeté et cohérence financière", "E (validation)"],
]
story.append(make_table(team_data, col_widths=[4.5*cm, 8.5*cm, 4*cm], font_size=9))

story.append(Paragraph("6. Risques et mitigation", H1))
risks_data = [
    ["Risque", "Probabilité", "Impact", "Mitigation"],
    ["Rupture soja non anticipée", "Élevée", "Critique", "4 scénarios modélisés, alerte 16/09"],
    ["Hausse prix Q4 non prévue", "Moyenne", "Moyen", "Scénario S4 (baisse) en complément"],
    ["Données historiques insuffisantes", "Moyenne", "Moyen", "20 mois de données (2025-S1 2026)"],
    ["Effet prix sur demande non quantifié", "Élevée", "Moyen", "Option A: prix stable Q4 (validée)"],
    ["Saisonnalité Q4 mal captée", "Faible", "Moyen", "Prophet yearly seasonality activée"],
    ["Désagrégation produit × agence biaisée", "Faible", "Faible", "Parts historiques 2025-S1 2026"],
]
story.append(make_table(risks_data, col_widths=[5.5*cm, 2.5*cm, 2.5*cm, 6.5*cm], font_size=8))

story.append(Paragraph("7. Planning et jalons", H1))
planning_data = [
    ["Jalon", "Date prévue", "Livrable"],
    ["Kick-off et validation périmètre", "J+0 (27/08/2026)", "Proposition de projet validée"],
    ["Phase PREPARE terminée", "J+2 (29/08/2026)", "Dataset consolidé + prix calculés"],
    ["Phase ANALYZE terminée", "J+3 (30/08/2026)", "AED et graphiques"],
    ["Phase CONSTRUCT terminée", "J+5 (01/09/2026)", "Forecast Q4 (4 scénarios)"],
    ["Phase EXECUTE - Livrables finaux", "J+6 (02/09/2026)", "Excel + PDFs + notebook"],
    ["Présentation comité de direction", "J+7 (03/09/2026)", "Validation officielle"],
    ["Mise à jour mensuelle", "Mensuel", "Réactualisation avec nouvelles données"],
]
story.append(make_table(planning_data, col_widths=[5.5*cm, 4*cm, 7.5*cm], font_size=9))

story.append(Paragraph("8. Budget et ressources", H1))
story.append(Paragraph(
    "Le projet ne nécessite pas de budget supplémentaire : il utilise les ressources informatiques existantes "
    "(ordinateur portable, Python, Jupyter, Prophet) et les données ERP déjà disponibles. "
    "L'investissement principal est le <b>temps du Data Analyst</b> (6 jours-homme) "
    "et la <b>sollicitation des directions métiers</b> pour la validation des hypothèses "
    "(Direction Commerciale, Direction Production, Contrôle de Gestion).",
    BODY))

story.append(Paragraph("9. Critères de succès", H1))
success = [
    "Forecast produit pour les 4 scénarios avec désagrégation complète (produit × agence × mois)",
    "Écart forecast vs réalité ≤ 15% à fin septembre (premier mois prédictible)",
    "Validation des hypothèses par la Direction Commerciale et le Contrôle de Gestion",
    "Livrables PACE complets (9 documents) produits et diffusés",
    "Réutilisabilité du modèle pour les prochains trimestres (notebook commenté)",
]
for s in success:
    story.append(Paragraph(f"• {s}", BULLET))

doc.build(story)
print(f"✓ Proposition de projet: {prop_path}")
print(f"  Size: {os.path.getsize(prop_path) / 1024:.0f} KB")

print("\n=== BATCH 1 COMPLETE (Resume executif + Proposition projet) ===")
