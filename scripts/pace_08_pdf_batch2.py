"""
Phase Execute - Batch 2 PDFs:
3. Matrice RACI
4. Document strategique PACE
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

# Fonts
pdfmetrics.registerFont(TTFont('DejaVuSans', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))

# Colors
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

styles = getSampleStyleSheet()
H1 = ParagraphStyle('H1', parent=styles['Heading1'], fontName='DejaVuSans-Bold',
                    fontSize=18, textColor=NAVY, spaceAfter=14, spaceBefore=10)
H2 = ParagraphStyle('H2', parent=styles['Heading2'], fontName='DejaVuSans-Bold',
                    fontSize=14, textColor=NAVY, spaceAfter=10, spaceBefore=14)
H3 = ParagraphStyle('H3', parent=styles['Heading3'], fontName='DejaVuSans-Bold',
                    fontSize=12, textColor=GOLD, spaceAfter=8, spaceBefore=10)
BODY = ParagraphStyle('Body', parent=styles['Normal'], fontName='DejaVuSans',
                      fontSize=10, leading=14, alignment=TA_JUSTIFY, spaceAfter=8)
BODY_BOLD = ParagraphStyle('BodyBold', parent=BODY, fontName='DejaVuSans-Bold')
BULLET = ParagraphStyle('Bullet', parent=BODY, leftIndent=20, bulletIndent=10, spaceAfter=4)
SMALL = ParagraphStyle('Small', parent=BODY, fontSize=8, textColor=GRAY)
CAPTION = ParagraphStyle('Caption', parent=BODY, fontSize=9, textColor=GRAY, alignment=TA_CENTER)

def make_table(data, col_widths=None, font_size=9, header_color=NAVY):
    t = Table(data, colWidths=col_widths, repeatRows=1)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), header_color),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'DejaVuSans-Bold'),
        ('FONTNAME', (0, 1), (-1, -1), 'DejaVuSans'),
        ('FONTSIZE', (0, 0), (-1, -1), font_size),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.gray),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, LIGHT_GRAY]),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ])
    t.setStyle(style)
    return t

def add_image(path, width=15*cm, caption=None):
    elements = []
    if os.path.exists(path):
        img = Image(path, width=width, height=width*0.6)
        img._restrictSize(width, 12*cm)
        elements.append(img)
        if caption:
            elements.append(Spacer(1, 0.2*cm))
            elements.append(Paragraph(caption, CAPTION))
    elements.append(Spacer(1, 0.3*cm))
    return elements

def cover_page(title, subtitle, doc_type, date_str="27 août 2026"):
    elements = []
    elements.append(Spacer(1, 4*cm))
    elements.append(Paragraph("BELGOCAM SA", ParagraphStyle('CoverLogo', parent=styles['Title'], 
        fontName='DejaVuSans-Bold', fontSize=32, textColor=NAVY, alignment=TA_CENTER)))
    elements.append(Spacer(1, 0.5*cm))
    elements.append(HRFlowable(width="60%", thickness=2, color=GOLD, spaceBefore=10, spaceAfter=20, hAlign='CENTER'))
    elements.append(Spacer(1, 1*cm))
    elements.append(Paragraph(title, ParagraphStyle('CT', parent=styles['Title'], 
        fontName='DejaVuSans-Bold', fontSize=28, textColor=NAVY, alignment=TA_CENTER, spaceAfter=20)))
    elements.append(Paragraph(subtitle, ParagraphStyle('CS', parent=styles['Title'], 
        fontName='DejaVuSans', fontSize=16, textColor=GOLD, alignment=TA_CENTER, spaceAfter=30)))
    elements.append(Spacer(1, 2*cm))
    elements.append(HRFlowable(width="40%", thickness=1, color=GRAY, spaceBefore=10, spaceAfter=10, hAlign='CENTER'))
    elements.append(Paragraph(f"<b>{doc_type}</b>", ParagraphStyle('CI', parent=BODY, fontSize=11, alignment=TA_CENTER, textColor=GRAY)))
    elements.append(Paragraph("William Francis Fohom", ParagraphStyle('CI2', parent=BODY, fontSize=11, alignment=TA_CENTER, textColor=GRAY)))
    elements.append(Paragraph("Data Analyst | Administrateur National de Ventes", ParagraphStyle('CI3', parent=BODY, fontSize=11, alignment=TA_CENTER, textColor=GRAY)))
    elements.append(Spacer(1, 0.5*cm))
    elements.append(Paragraph(date_str, ParagraphStyle('CI4', parent=BODY, fontSize=11, alignment=TA_CENTER, textColor=GRAY)))
    elements.append(PageBreak())
    return elements

OUT_DIR = "/home/z/my-project/download/forecast_q4_2026"
CHARTS_DIR = f"{OUT_DIR}/charts"


# ========================================================================
# 3. MATRICE RACI (3 pages)
# ========================================================================
print("Generating 3. Matrice RACI...")
raci_path = f"{OUT_DIR}/03_matrice_raci.pdf"
doc = SimpleDocTemplate(raci_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm,
                        leftMargin=1.5*cm, rightMargin=1.5*cm)
story = []

story.extend(cover_page("Matrice RACI", "Forecast Q4 2026", "MATRICE RACI"))

story.append(Paragraph("1. Définition des rôles RACI", H1))
story.append(Paragraph(
    "La matrice RACI est un outil de gestion de projet qui clarifie les rôles et responsabilités "
    "de chaque acteur à chaque étape du projet. L'acronyme RACI signifie :",
    BODY))

raci_def = [
    ["Lettre", "Rôle", "Description"],
    ["R", "Responsable (Responsible)", "Celui qui réalise la tâche. Il est opérationnel et accountable du livrable."],
    ["A", "Approbateur (Accountable)", "Celui qui rend compte au final. Un seul par tâche. Il valide le livrable."],
    ["C", "Consulté (Consulted)", "Celui dont l'avis est sollicité avant la décision. Communication bidirectionnelle."],
    ["I", "Informé (Informed)", "Celui qui est tenu au courant de la décision après coup. Communication unidirectionnelle."],
]
story.append(make_table(raci_def, col_widths=[1.5*cm, 5*cm, 11*cm], font_size=9))

story.append(Paragraph("2. Acteurs du projet", H1))
story.append(Paragraph(
    "Bien que l'équipe d'analyse soit composée d'une seule personne, "
    "le projet forecast Q4 2026 mobilise plusieurs acteurs internes de BELGOCAM. "
    "Chacun intervient à des étapes spécifiques selon son rôle fonctionnel.",
    BODY))

actors = [
    ["Code", "Acteur", "Fonction", "Implication"],
    ["DA", "Data Analyst", "William Francis Fohom - Administrateur National de Ventes", "Pleine (R sur toutes les tâches techniques)"],
    ["DC", "Direction Commerciale", "Responsable des objectifs commerciaux et du réseau d'agences", "Validation des hypothèses et objectifs"],
    ["DP", "Direction Production", "Responsable de la production et des stocks BEKOKO", "Fourniture des données stock et capacité"],
    ["DG", "Direction Générale", "Direction stratégique de BELGOCAM SA", "Décision finale et allocation ressources"],
    ["CG", "Contrôle de Gestion", "Service financier en charge du CA et budgets", "Validation de la cohérence financière"],
    ["RA", "Responsables d'Agences", "14 responsables d'agences BELGOCAM", "Informés et consultés sur leurs périmètres"],
]
story.append(make_table(actors, col_widths=[1.2*cm, 3.5*cm, 6.5*cm, 5.5*cm], font_size=9))

story.append(PageBreak())
story.append(Paragraph("3. Matrice RACI détaillée par étape PACE", H1))
story.append(Paragraph(
    "Le tableau ci-dessous présente la répartition des rôles RACI pour chaque tâche du projet, "
    "organisée selon les 4 phases de la méthodologie PACE.",
    BODY))

# Matrice RACI détaillée
raci_matrix = [
    ["Phase", "Tâche", "DA", "DC", "DP", "DG", "CG", "RA"],
    ["PREPARE", "1.1 Définition du périmètre", "R", "C", "C", "A", "I", "I"],
    ["", "1.2 Collecte des données ERP", "R", "I", "C", "I", "I", "I"],
    ["", "1.3 Nettoyage et consolidation", "R/A", "I", "I", "I", "I", "I"],
    ["", "1.4 Calcul des prix moyens", "R/A", "C", "I", "I", "C", "I"],
    ["", "1.5 Validation dataset", "R", "C", "C", "I", "A", "I"],
    ["ANALYZE", "2.1 AED (stats descriptives)", "R/A", "I", "I", "I", "I", "I"],
    ["", "2.2 Analyse saisonnalité", "R/A", "I", "I", "I", "I", "I"],
    ["", "2.3 Top produits/agences", "R/A", "C", "I", "I", "I", "I"],
    ["", "2.4 Graphiques AED", "R/A", "I", "I", "I", "I", "I"],
    ["", "2.5 Validation hypothèses AED", "R", "C", "C", "I", "A", "I"],
    ["CONSTRUCT", "3.1 Configuration Prophet", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.2 Entraînement 13 modèles", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.3 Définition 4 scénarios soja", "R", "C", "C", "A", "C", "I"],
    ["", "3.4 Application des scénarios", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.5 Désagrégation produit × agence", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.6 Validation forecast", "R", "C", "I", "A", "C", "I"],
    ["EXECUTE", "4.1 Excel multi-feuilles", "R/A", "I", "I", "I", "I", "I"],
    ["", "4.2 Notebook Jupyter", "R/A", "I", "I", "I", "I", "I"],
    ["", "4.3 Guide méthodologique PDF", "R/A", "I", "I", "I", "I", "I"],
    ["", "4.4 Document stratégique PACE", "R/A", "I", "I", "I", "I", "I"],
    ["", "4.5 Matrice RACI", "R/A", "C", "C", "I", "C", "I"],
    ["", "4.6 Résumé exécutif", "R", "C", "I", "A", "C", "I"],
    ["", "4.7 Proposition de projet", "R", "C", "C", "A", "C", "I"],
    ["", "4.8 Présentation au CODIR", "R", "C", "C", "A", "C", "I"],
    ["", "4.9 Diffusion aux responsables d'agences", "R", "A", "I", "I", "I", "C"],
    ["", "4.10 Mise à jour mensuelle", "R/A", "C", "C", "I", "I", "I"],
]
story.append(make_table(raci_matrix, col_widths=[2.2*cm, 5.5*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm], font_size=7.5))

story.append(Spacer(1, 0.3*cm))
story.append(Paragraph(
    "<b>Lecture de la matrice</b> : Le Data Analyst (DA) est Responsable (R) de toutes les tâches techniques. "
    "La Direction Générale (DG) est Approbateur (A) sur les tâches stratégiques "
    "(définition du périmètre, scénarios, validation forecast, diffusion). "
    "La Direction Commerciale (DC) et le Contrôle de Gestion (CG) sont Consultés (C) "
    "sur les hypothèses et validations. Les Responsables d'Agences (RA) sont Informés (I) "
    "et Consultés (C) uniquement pour la diffusion des résultats.",
    BODY))

story.append(PageBreak())
story.append(Paragraph("4. Synthèse des responsabilités par acteur", H1))

story.append(Paragraph("<b>Data Analyst (DA) — William Francis Fohom</b>", H3))
story.append(Paragraph(
    "Le Data Analyst est le <b>moteur opérationnel</b> du projet. Il est Responsable (R) de toutes les tâches "
    "techniques : préparation des données, AED, modélisation Prophet, génération des livrables. "
    "Il est également Accountable (A) sur les livrables techniques (dataset, prix, forecast, Excel, PDFs, notebook). "
    "Son rôle couvre l'intégralité du cycle PACE.",
    BODY))

story.append(Paragraph("<b>Direction Commerciale (DC)</b>", H3))
story.append(Paragraph(
    "La Direction Commerciale est <b>Consultée (C)</b> sur les hypothèses commerciales "
    "(saisonnalité, effets prix, objectifs par agence) et <b>Approbatrice (A)</b> sur la diffusion "
    "aux responsables d'agences. Elle valide la cohérence du forecast avec la stratégie commerciale.",
    BODY))

story.append(Paragraph("<b>Direction Production (DP)</b>", H3))
story.append(Paragraph(
    "La Direction Production est <b>Consultée (C)</b> sur les données de stock (BEKOKO), "
    "la capacité de production concentrés (4 500-5 300 sacs/sem) et les hypothèses de réapprovisionnement soja. "
    "Elle fournit les inputs critiques pour les scénarios S1-S4.",
    BODY))

story.append(Paragraph("<b>Direction Générale (DG)</b>", H3))
story.append(Paragraph(
    "La Direction Générale est <b>Approbatrice (A)</b> sur les jalons stratégiques : "
    "validation du périmètre, choix du scénario de référence, validation finale du forecast, "
    "et décision de mise en œuvre. Elle arbitre les arbitrages budget/ressources.",
    BODY))

story.append(Paragraph("<b>Contrôle de Gestion (CG)</b>", H3))
story.append(Paragraph(
    "Le Contrôle de Gestion est <b>Consulté (C)</b> sur la cohérence financière du forecast "
    "(CA projeté vs objectifs annuels, validation des prix extrapolés). "
    "Il valide la conformité aux normes comptables et financières de BELGOCAM.",
    BODY))

story.append(Paragraph("<b>Responsables d'Agences (RA)</b>", H3))
story.append(Paragraph(
    "Les 14 responsables d'agences sont <b>Informés (I)</b> des résultats du forecast "
    "et <b>Consultés (C)</b> lors de la diffusion pour validation de leur périmètre spécifique. "
    "Ils sont les utilisateurs finaux du forecast pour le pilotage de leur agence.",
    BODY))

story.append(Paragraph("5. Points de contrôle et jalons de validation", H1))
control_data = [
    ["Jalon", "Phase", "Date", "Approbateur", "Livrable validé"],
    ["Kick-off", "PREPARE", "27/08/2026", "DG", "Proposition de projet"],
    ["Validation dataset", "PREPARE", "29/08/2026", "CG", "Dataset consolidé"],
    ["Validation AED", "ANALYZE", "30/08/2026", "DC", "Synthèse AED + graphiques"],
    ["Validation scénarios", "CONSTRUCT", "31/08/2026", "DG + DP", "4 scénarios soja"],
    ["Validation forecast", "CONSTRUCT", "01/09/2026", "DG + DC", "Forecast Q4 complet"],
    ["Livrables finaux", "EXECUTE", "02/09/2026", "DG", "9 livrables PACE"],
    ["Présentation CODIR", "EXECUTE", "03/09/2026", "DG", "Validation officielle"],
]
story.append(make_table(control_data, col_widths=[3*cm, 2.5*cm, 2.5*cm, 3*cm, 5*cm], font_size=8))

doc.build(story)
print(f"✓ Matrice RACI: {raci_path}")
print(f"  Size: {os.path.getsize(raci_path) / 1024:.0f} KB")


# ========================================================================
# 4. DOCUMENT STRATÉGIQUE PACE (15-20 pages)
# ========================================================================
print("\nGenerating 4. Document stratégique PACE...")
strat_path = f"{OUT_DIR}/04_document_strategique_pace.pdf"
doc = SimpleDocTemplate(strat_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm,
                        leftMargin=2*cm, rightMargin=2*cm)
story = []

story.extend(cover_page("Document Stratégique PACE", "Forecast Q4 2026 - Volumes et Valeurs", "STRATÉGIE PACE"))

# === Sommaire ===
story.append(Paragraph("Sommaire", H1))
story.append(Paragraph(
    "1. Vision et objectifs stratégiques<br/>"
    "2. Phase PREPARE - Préparation des données<br/>"
    "3. Phase ANALYZE - Analyse Exploratoire (AED)<br/>"
    "4. Phase CONSTRUCT - Construction du modèle<br/>"
    "5. Phase EXECUTE - Génération des livrables<br/>"
    "6. Scénarios stratégiques et analyse de sensibilité<br/>"
    "7. Plan de déploiement et adoption<br/>"
    "8. Mesure de la valeur et KPIs de succès<br/>"
    "9. Gouvernance et comité de pilotage<br/>"
    "10. Conclusion et prochaines étapes",
    BODY))

story.append(PageBreak())

# === Section 1 ===
story.append(Paragraph("1. Vision et objectifs stratégiques", H1))
story.append(Paragraph("1.1 Vision", H2))
story.append(Paragraph(
    "Faire du forecast Q4 2026 un <b>outil de pilotage stratégique</b> permettant à BELGOCAM SA "
    "d'anticiper ses volumes de ventes et son chiffre d'affaires avec une précision opérationnelle, "
    "tout en intégrant l'incertitude liée à la situation critique du stock soja. "
    "Ce forecast doit servir de référence pour les décisions commerciales, "
    "industrielles et financières du quatrième trimestre 2026.",
    BODY))

story.append(Paragraph("1.2 Objectifs stratégiques", H2))
obj_data = [
    ["Objectif", "Description", "KPI de succès"],
    ["Anticipation", "Produire un forecast Q4 2026 fiable (précision ≥ 85%)", "Écart forecast vs réel ≤ 15%"],
    ["Désagrégation", "Détail par produit × agence × mois pour le pilotage opérationnel", "1 680+ lignes"],
    ["Scénarios", "Modéliser 4 scénarios de réapprovisionnement soja", "S1-S4 produits"],
    ["Valeur", "Calculer le CA projeté (M FCFA) en plus du volume (tonnes)", "2 dimensions"],
    ["Adoption", "Diffuser auprès des 14 agences et du CODIR", "100% agences informées"],
    ["Réutilisabilité", "Notebook Jupyter reproductible pour les prochains trimestres", "Code commenté"],
]
story.append(make_table(obj_data, col_widths=[3*cm, 9*cm, 4*cm], font_size=9))

story.append(Paragraph("1.3 Alignement avec la stratégie BELGOCAM", H2))
story.append(Paragraph(
    "Ce projet s'inscrit dans la stratégie globale de BELGOCAM SA de <b>pilotage par les données</b> (data-driven decision making). "
    "Il complète les analyses mensuelles déjà en place (analyse zero achat, bundle soja-concentrés, suivi soja mensuel) "
    "en apportant une dimension prédictive sur 4 mois. "
    "Le forecast Q4 2026 alimente directement la planification commerciale, "
    "la gestion des stocks et le budget financier du quatrième trimestre.",
    BODY))

story.append(PageBreak())

# === Section 2 - PREPARE ===
story.append(Paragraph("2. Phase PREPARE - Préparation des données", H1))
story.append(Paragraph("2.1 Sources de données consolidées", H2))
story.append(Paragraph(
    "Le dataset consolidé regroupe <b>95 819 enregistrements de ventes Livrées</b> "
    "couvrant la période janvier 2025 - 26 août 2026 (20 mois d'historique). "
    "Cette consolidation a nécessité l'intégration de 4 sources ERP hétérogènes, "
    "chacune avec un format de colonnes légèrement différent.",
    BODY))

sources_data = [
    ["Source", "Période", "Records", "Usage"],
    ["86d96135-...xlsx", "Jan-Déc 2025", "52 818", "Historique annuel (saisonnalité)"],
    ["ventes janv a juin 2026.xlsx", "Jan-Juin 2026", "31 458", "S1 2026 (tendance récente)"],
    ["NJS GROUP ERP (9).xlsx", "Juillet 2026", "6 410", "Mois complet"],
    ["NJS GROUP ERP (21).xlsx", "01-26/08/2026", "5 133", "Août MTD (dernier mois)"],
    ["TOTAL", "20 mois", "95 819", "Dataset consolidé"],
]
story.append(make_table(sources_data, col_widths=[5*cm, 3.5*cm, 2.5*cm, 5*cm], font_size=9))

story.append(Paragraph("2.2 Nettoyage et normalisation", H2))
story.append(Paragraph(
    "Plusieurs étapes de nettoyage ont été réalisées pour garantir la qualité du dataset :",
    BODY))
cleaning = [
    "Filtrage sur l'état 'Livrée' uniquement (exclusion des 'Validée', 'En cours', 'Annulée', 'Brouillon')",
    "Exclusion des clients internes (SPC, PDC, COMPTOIR, EMANA) pour éviter les doubles comptes",
    "Exclusion des agences non commerciales (SPC BAF-CHEFFERIE, PDC Emana, SPC-NDERE, SPC-DSCHANG, SPC BUEA, SPC-YASSA)",
    "Exclusion des produits hors périmètre alimentaire animal (DIVERS, MATERIEL ELEVAGE)",
    "Normalisation des noms d'agences (formats 2025 courts vs 2026 avec préfixe 'AGENCE')",
    "Conversion des quantités en tonnes via le poids unitaire par référence produit",
    "Calcul des sacs équivalent 50 kg pour les comparaisons inters-produits",
]
for c in cleaning:
    story.append(Paragraph(f"• {c}", BULLET))

story.append(Paragraph("2.3 Calcul des prix moyens (Option A)", H2))
story.append(Paragraph(
    "Les prix unitaires par produit ont été extrapolés à partir du CA HT/TTC et des quantités "
    "des extractions existantes. Pour chaque référence produit, le prix moyen mensuel a été calculé "
    "en divisant le montant TTC par le nombre de sacs équivalent 50 kg. "
    "Les valeurs aberrantes (1er-99e percentile) ont été filtrées pour garantir la robustesse. "
    "Le prix retenu pour le forecast Q4 2026 est le prix le plus récent disponible (août 2026).",
    BODY))

story.append(Paragraph(
    "Précision estimée de l'extrapolation : <b>~90%</b> (vs ~99% si un fichier prix officiel était fourni). "
    "Pour le scénario S4 (baisse prix), une baisse de -10% est appliquée sur les prix soja uniquement, "
    "modélisant l'hypothèse d'un ajustement commercial si le stock se stabilise.",
    BODY))

story.append(PageBreak())

# === Section 3 - ANALYZE ===
story.append(Paragraph("3. Phase ANALYZE - Analyse Exploratoire (AED)", H1))
story.append(Paragraph("3.1 Statistiques descriptives", H2))
story.append(Paragraph(
    "L'AED a permis de dégager les enseignements clés suivants sur les 20 mois d'historique :",
    BODY))

stats_data = [
    ["Famille", "Volume total (t)", "CA total (M FCFA)", "Nb produits", "Top agence"],
    ["TOURTEAUX", "83 327", "34 290", "4", "FAMLA (34 444 t)"],
    ["CONCENTRÉS", "24 547", "16 600", "12", "FAMLA (8 950 t)"],
    ["INGRÉDIENTS", "628", "264", "7", "MESSASSI (180 t)"],
    ["ALIMENT COMPLET", "535", "218", "2", "FAMLA (245 t)"],
    ["MAIS", "8", "0", "1", "FAMLA (8 t)"],
    ["TOTAL", "109 045", "51 372", "26", "—"],
]
story.append(make_table(stats_data, col_widths=[3*cm, 3*cm, 3.5*cm, 2.5*cm, 4*cm], font_size=9))

story.append(Paragraph("3.2 Saisonnalité mensuelle", H2))
story.append(Paragraph(
    "L'analyse de saisonnalité (base 2025) révèle une <b>forte concentration des volumes en Q4</b>, "
    "particulièrement marquée pour les TOURTEAUX (octobre = indice 1.86) et l'ALIMENT COMPLET (septembre = 1.58). "
    "Cette saisonnalité est liée à la rentrée scolaire (septembre), aux fêtes de fin d'année (décembre) "
    "et à la préparation de la période de pic de ponte/ponte en octobre-novembre.",
    BODY))

story.extend(add_image(f"{CHARTS_DIR}/02_saisonnalite_mensuelle.png", width=14*cm,
                       caption="Figure 1 - Indice de saisonnalité mensuelle par famille (base 2025)"))

q4_share_data = [
    ["Famille", "Part Q4 2025 dans l'année (%)"],
    ["TOURTEAUX", "46.3%"],
    ["CONCENTRÉS", "35.2%"],
    ["INGRÉDIENTS", "36.6%"],
    ["ALIMENT COMPLET", "42.0%"],
]
story.append(make_table(q4_share_data, col_widths=[6*cm, 8*cm], font_size=9))
story.append(Paragraph(
    "<b>Implication pour le forecast</b> : Le Q4 représente entre 35% et 46% du volume annuel selon les familles. "
    "Le modèle Prophet doit impérativement capturer cette saisonnalité forte via le paramètre "
    "<code>yearly_seasonality=True</code>.",
    BODY))

story.append(PageBreak())
story.append(Paragraph("3.3 Top produits et agences", H2))
story.append(Paragraph(
    "L'analyse des top produits et agences permet d'identifier les leviers principaux du forecast :",
    BODY))

story.append(Paragraph("<b>Top 5 produits par volume</b>", H3))
top_prod_data = [
    ["Rang", "Réf", "Description", "Famille", "Volume (t)", "Part du total"],
    ["1", "T102", "TOURTEAUX DE SOJA 50 KG", "TOURTEAUX", "83 305", "76.4%"],
    ["2", "C104", "BELGO 10% CHAIR 50Kg", "CONCENTRÉS", "14 404", "13.2%"],
    ["3", "C101", "BELGO 5% PONTE 50 Kg", "CONCENTRÉS", "5 616", "5.2%"],
    ["4", "C103", "BELGO 5% CHAIR 50 Kg", "CONCENTRÉS", "4 487", "4.1%"],
    ["5", "CB100", "CHICK BOOSTER 25 Kg", "ALIMENT COMPLET", "500", "0.5%"],
]
story.append(make_table(top_prod_data, col_widths=[1.5*cm, 1.5*cm, 5*cm, 3*cm, 2.5*cm, 2.5*cm], font_size=8))

story.append(Paragraph(
    "<b>Concentration</b> : Le soja (T102) représente à lui seul <b>76% du volume total</b>. "
    "Les 3 principaux concentrés (C104, C101, C103) représentent 22% supplémentaires. "
    "La modélisation doit donc accorder une attention particulière à ces 4 produits.",
    BODY))

story.append(Paragraph("<b>Top 5 agences par volume</b>", H3))
top_ag_data = [
    ["Rang", "Agence", "Région", "Volume (t)", "Part du total"],
    ["1", "FAMLA", "Ouest", "34 444", "31.6%"],
    ["2", "NDOBO", "Littoral", "14 958", "13.7%"],
    ["3", "MESSASSI", "Centre", "10 775", "9.9%"],
    ["4", "DJELENG", "Ouest", "9 779", "9.0%"],
    ["5", "MBOUDA", "Ouest", "6 567", "6.0%"],
]
story.append(make_table(top_ag_data, col_widths=[1.5*cm, 3*cm, 3*cm, 3*cm, 3*cm], font_size=9))

story.append(Paragraph(
    "<b>Concentration géographique</b> : L'Ouest (FAMLA + DJELENG + MBOUDA) représente <b>46.6%</b> du volume total. "
    "Le Littoral (NDOBO + autres) pèse 28%, le Centre (MESSASSI + autres) 25%. "
    "Cette concentration justifie la désagrégation par région dans le forecast.",
    BODY))

story.append(PageBreak())

# === Section 4 - CONSTRUCT ===
story.append(Paragraph("4. Phase CONSTRUCT - Construction du modèle", H1))
story.append(Paragraph("4.1 Choix méthodologique : Prophet", H2))
story.append(Paragraph(
    "Le modèle <b>Prophet</b> (développé par Facebook/Meta) a été retenu pour ce forecast pour plusieurs raisons :",
    BODY))

prophet_reasons = [
    "<b>Gestion de la saisonnalité</b> : Prophet détecte automatiquement les saisonnalités annuelles, hebdomadaires et journalières. Le paramètre <code>yearly_seasonality=True</code> est crucial pour capturer les pics Q4.",
    "<b>Robustesse aux données manquantes</b> : Prophet gère naturellement les mois sans ventes (cas du MAÏS en Q4 2025).",
    "<b>Interprétabilité</b> : La décomposition tendance + saisonnalité + vacances est lisible par les décideurs non-techniques.",
    "<b>Intervals de confiance</b> : Prophet fournit nativement des bornes inférieures et supérieures (paramètre <code>interval_width=0.8</code>).",
    "<b>Rapidité</b> : L'entraînement d'un modèle Prophet prend quelques secondes, permettant de tester rapidement plusieurs configurations.",
    "<b>Mode multiplicatif</b> : Le paramètre <code>seasonality_mode='multiplicative'</code> est adapté aux séries où la saisonnalité croît avec la tendance (cas BELGOCAM).",
]
for r in prophet_reasons:
    story.append(Paragraph(f"• {r}", BULLET))

story.append(Paragraph("4.2 Stratégie de modélisation : famille × région", H2))
story.append(Paragraph(
    "Plutôt que de modéliser chaque combinaison produit × agence (130+ modèles, temps de calcul prohibitif), "
    "nous avons opté pour une <b>stratégie en deux niveaux</b> :",
    BODY))

strategy = [
    "Niveau 1 : Entraînement de 13 modèles Prophet au niveau famille × région (5 familles × 3 régions, moins les combinaisons sans données)",
    "Niveau 2 : Désagrégation du forecast famille × région vers le niveau produit × agence, en appliquant les parts historiques de chaque combinaison",
]
for s in strategy:
    story.append(Paragraph(f"• {s}", BULLET))

story.append(Paragraph(
    "<b>Avantages</b> : Réduction du temps de calcul (de 30+ minutes à moins de 2 minutes), "
    "meilleure robustesse statistique (plus de données par modèle), "
    "préservation de la structure produit × agence via les parts historiques. "
    "<b>Limite</b> : La désagrégation suppose que les parts historiques restent stables en Q4 2026, "
    "ce qui peut sous-estimer les effets de réallocations (ex: transfert clients FAMLA → NDOBO observé en juillet-août).",
    BODY))

story.append(Paragraph("4.3 Configuration Prophet", H2))
config_data = [
    ["Paramètre", "Valeur", "Justification"],
    ["yearly_seasonality", "True", "Capture la saisonnalité annuelle (pics Q4)"],
    ["weekly_seasonality", "False", "Données mensuelles, pas de saisonnalité hebdo"],
    ["daily_seasonality", "False", "Données mensuelles, pas de saisonnalité journalière"],
    ["seasonality_mode", "multiplicative", "Saisonnalité proportionnelle à la tendance"],
    ["changepoint_prior_scale", "0.05", "Tendance modérément flexible (évite surajustement)"],
    ["interval_width", "0.8", "Intervalle de confiance à 80%"],
    ["mcmc_samples", "0", "Désactivation MCMC pour rapidité (méthode MAP)"],
]
story.append(make_table(config_data, col_widths=[5*cm, 3*cm, 8*cm], font_size=9))

story.append(PageBreak())

# === Section 5 - EXECUTE ===
story.append(Paragraph("5. Phase EXECUTE - Génération des livrables", H1))
story.append(Paragraph("5.1 Livrables produits", H2))
livrables_data = [
    ["#", "Livrable", "Format", "Usage"],
    ["1", "Résumé exécutif", "PDF (2 pages)", "Direction Générale - Décision rapide"],
    ["2", "Proposition de projet", "PDF (7 pages)", "Cadre formel du projet - Validation"],
    ["3", "Matrice RACI", "PDF (3 pages)", "Gouvernance - Rôles et responsabilités"],
    ["4", "Document stratégique PACE", "PDF (15+ pages)", "Référence méthodologique complète"],
    ["5", "Guide méthodologique", "PDF (30+ pages)", "Détail des méthodes et outils"],
    ["6", "Excel forecast", "XLSX (9 feuilles)", "Données détaillées - Pilotage opérationnel"],
    ["7", "Notebook Jupyter", ".ipynb", "Code reproductible - Audit et réutilisation"],
    ["8", "Graphiques de visualisation", "PNG (13 graphiques)", "Supports de présentation"],
]
story.append(make_table(livrables_data, col_widths=[1*cm, 5*cm, 4*cm, 6*cm], font_size=9))

story.append(Paragraph("5.2 Excel forecast multi-feuilles", H2))
story.append(Paragraph(
    "Le fichier Excel <code>forecast_q4_2026_volumes_valeurs.xlsx</code> contient 9 feuilles :",
    BODY))
excel_sheets = [
    "<b>1. Synthèse</b> - Vue globale par scénario (volumes et CA)",
    "<b>2. Par Famille</b> - Détail par famille × mois × scénario",
    "<b>3. Par Région</b> - Détail par région × mois × scénario",
    "<b>4. Par Agence</b> - Top 14 agences par scénario avec part de CA",
    "<b>5. Par Produit</b> - 26 références produits par scénario",
    "<b>6. Par Mois</b> - Évolution mensuelle sept-déc par famille et scénario",
    "<b>7. Détail complet</b> - Toutes les combinaisons produit × agence × mois × scénario (4 196 lignes)",
    "<b>8. Prix utilisés</b> - Prix par produit (stable + baisse) avec source",
    "<b>9. Hypothèses</b> - Tous les paramètres et scénarios documentés",
]
for s in excel_sheets:
    story.append(Paragraph(f"• {s}", BULLET))

story.append(PageBreak())

# === Section 6 - Scénarios ===
story.append(Paragraph("6. Scénarios stratégiques et analyse de sensibilité", H1))
story.append(Paragraph(
    "La situation critique du stock soja (rupture probable au 16/09/2026) rend le forecast Q4 2026 "
    "particulièrement sensible aux décisions de réapprovisionnement. "
    "Quatre scénarios ont été modélisés pour couvrir le spectre des possibles.",
    BODY))

story.append(Paragraph("6.1 Description des scénarios", H2))
scenarios_data = [
    ["Scénario", "Hypothèse réappro", "Sept", "Oct", "Nov", "Déc"],
    ["S1 - Rupture totale", "Aucun réappro", "50%", "0%", "0%", "0%"],
    ["S2 - Réappro 50%", "40 000 sacs au 01/10", "50%", "70%", "70%", "100%"],
    ["S3 - Réappro 100%", "80 000 sacs au 15/09", "70%", "100%", "100%", "100%"],
    ["S4 - Baisse prix", "S3 + baisse prix -10%", "70% × 1.05", "100% × 1.10", "100% × 1.10", "100% × 1.10"],
]
story.append(make_table(scenarios_data, col_widths=[3.5*cm, 4.5*cm, 2*cm, 2*cm, 2*cm, 2*cm], font_size=9))

story.append(Paragraph("6.2 Résultats comparés", H2))
results_data = [
    ["Scénario", "Volume Q4 (t)", "CA Q4 (M FCFA)", "vs S3 (CA)", "Recommandation"],
    ["S1 - Rupture totale", "7 291", "4 656", "-73%", "À ÉVITER ABSOLUMENT"],
    ["S2 - Réappro 50%", "32 207", "14 922", "-14%", "Plan B (si réappro partiel)"],
    ["S3 - Réappro 100%", "37 988", "17 304", "—", "RÉFÉRENCE (recommandé)"],
    ["S4 - Baisse prix", "41 067", "17 168", "-1%", "Optimiste (si stock stabilisé)"],
]
story.append(make_table(results_data, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2*cm, 5.5*cm], font_size=9))

story.extend(add_image(f"{CHARTS_DIR}/08_ca_par_scenario.png", width=14*cm,
                       caption="Figure 2 - CA Q4 2026 par scénario (M FCFA)"))

story.append(Paragraph("6.3 Analyse de sensibilité", H2))
story.append(Paragraph(
    "L'analyse de sensibilité révèle deux enseignements stratégiques :",
    BODY))

story.append(Paragraph(
    "<b>1. Impact critique du réapprovisionnement</b> : L'écart entre S1 (rupture) et S3 (réappro 100%) "
    "est de <b>12 648 M FCFA</b> (soit -73% de CA). C'est de loin le facteur de sensibilité n°1. "
    "La décision de réapprovisionner ou non le soja avant le 15/09/2026 conditionne 73% du CA Q4.",
    BODY))

story.append(Paragraph(
    "<b>2. Effet volume vs effet prix (S3 vs S4)</b> : Le scénario S4 (baisse prix -10%) génère "
    "<b>+3 079 t de volume</b> vs S3, mais <b>-136 M FCFA de CA</b>. "
    "La baisse de prix stimule la demande (+8% en volume) mais l'effet volume ne compense pas la baisse unitaire. "
    "Sur le plan purement financier, <b>maintenir les prix (S3) est plus rentable que baisser les prix (S4)</b>, "
    "sauf si la baisse permet de capter des parts de marché concurrentielles.",
    BODY))

story.append(PageBreak())

# === Section 7 ===
story.append(Paragraph("7. Plan de déploiement et adoption", H1))
story.append(Paragraph("7.1 Calendrier de déploiement", H2))
deploy_data = [
    ["Étape", "Date", "Action", "Cible"],
    ["1", "27/08/2026", "Validation proposition de projet", "DG"],
    ["2", "02/09/2026", "Livrables finaux prêts", "DA"],
    ["3", "03/09/2026", "Présentation au CODIR", "DG + DC + DP + CG"],
    ["4", "04/09/2026", "Diffusion Excel aux responsables d'agences", "14 RA"],
    ["5", "10/09/2026", "Briefing individuel agences sous-performantes", "FAMLA, NDOBO, MESSASSI"],
    ["6", "15/09/2026", "Point de situation J+15 (réappro soja ?)", "CODIR"],
    ["7", "01/10/2026", "Mise à jour mensuelle du forecast", "DA"],
    ["8", "Mensuel", "Suivi écart forecast vs réel", "DA + DC"],
]
story.append(make_table(deploy_data, col_widths=[1*cm, 2.5*cm, 7*cm, 5.5*cm], font_size=9))

story.append(Paragraph("7.2 Stratégie d'adoption", H2))
story.append(Paragraph(
    "Pour garantir l'adoption du forecast par les équipes commerciales, plusieurs leviers sont activés :",
    BODY))
adoption = [
    "<b>Formation</b> : Session de 1h pour les 14 responsables d'agences sur la lecture de l'Excel forecast",
    "<b>Appropriation</b> : Chaque responsable reçoit son forecast spécifique avec sa part de CA projetée",
    "<b>Suivi mensuel</b> : Comparaison forecast vs réalisé au niveau agence, avec analyse des écarts",
    "<b>Mise à jour</b> : Réactualisation mensuelle du forecast avec les nouvelles données ERP",
    "<b>Feedback</b> : Recueil des retours des utilisateurs pour améliorer le modèle (V2 en janvier 2027)",
]
for a in adoption:
    story.append(Paragraph(f"• {a}", BULLET))

story.append(Paragraph("8. Mesure de la valeur et KPIs de succès", H1))
story.append(Paragraph(
    "Le succès du projet sera mesuré sur 3 dimensions : précision, adoption et valeur business.",
    BODY))

kpi_data = [
    ["Dimension", "KPI", "Cible", "Mesure"],
    ["Précision", "Écart forecast vs réel (sept 2026)", "≤ 15%", "Comparaison fin octobre"],
    ["Précision", "Écart forecast vs réel (Q4 2026)", "≤ 20%", "Comparaison fin janvier 2027"],
    ["Adoption", "Taux de consultation Excel forecast", "≥ 80% agences", "Enquête fin septembre"],
    ["Adoption", "Taux d'utilisation pour planification commerciale", "≥ 70% agences", "Enquête fin octobre"],
    ["Valeur", "Décision de réappro soja prise avant 15/09", "Oui/Non", "Vérification DG"],
    ["Valeur", "Objectifs Q4 2026 calibrés sur forecast", "Oui/Non", "Validation DC"],
    ["Valeur", "Réutilisation du modèle pour Q1 2027", "Oui/Non", "Vérification DA"],
]
story.append(make_table(kpi_data, col_widths=[2.5*cm, 6*cm, 3*cm, 4.5*cm], font_size=9))

story.append(Paragraph("9. Gouvernance et comité de pilotage", H1))
story.append(Paragraph(
    "Le projet est piloté par un <b>comité de pilotage restreint</b> composé du Data Analyst (pilote opérationnel), "
    "de la Direction Commerciale (validateur métier) et de la Direction Générale (validateur stratégique). "
    "Le Contrôle de Gestion et la Direction Production sont consultés selon les besoins.",
    BODY))

story.append(Paragraph(
    "Le comité se réunit à 3 moments clés : (1) kick-off (27/08), (2) validation forecast (01/09), "
    "(3) présentation au CODIR (03/09). Un point de situation J+15 (15/09) permet de vérifier "
    "la décision de réapprovisionnement soja et d'ajuster le scénario de référence si nécessaire.",
    BODY))

story.append(Paragraph("10. Conclusion et prochaines étapes", H1))
story.append(Paragraph(
    "Le forecast Q4 2026 fournit à BELGOCAM SA un <b>outil de pilotage stratégique</b> "
    "qui transforme 95 819 enregistrements historiques en projections actionnables. "
    "Les 4 scénarios modélisés permettent d'anticiper l'impact des décisions de réapprovisionnement soja "
    "et de calibrer les objectifs commerciaux Q4 avec une précision opérationnelle.",
    BODY))

story.append(Paragraph(
    "<b>Prochaines étapes immédiates</b> :",
    BODY_BOLD))
next_steps = [
    "Validation de la proposition de projet par la Direction Générale (27/08/2026)",
    "Prise de décision sur le réapprovisionnement soja avant le 15/09/2026",
    "Présentation du forecast au CODIR le 03/09/2026",
    "Diffusion aux 14 responsables d'agences le 04/09/2026",
    "Mise à jour mensuelle du forecast à partir d'octobre 2026",
]
for s in next_steps:
    story.append(Paragraph(f"• {s}", BULLET))

story.append(Paragraph(
    "<b>Prochaines étapes à moyen terme</b> :",
    BODY_BOLD))
mid_term = [
    "Réutilisation du modèle pour le forecast Q1 2027 (janvier 2027)",
    "Intégration de variables externes (prix matière première, météo, calendrier scolaire) en V2",
    "Extension à d'autres familles de produits (DIVERS, MATERIEL ELEVAGE) si pertinent",
    "Industrialisation via un pipeline automatisé (Airflow, dbt) si le volume de mises à jour l justifie",
]
for s in mid_term:
    story.append(Paragraph(f"• {s}", BULLET))

doc.build(story)
print(f"✓ Document stratégique PACE: {strat_path}")
print(f"  Size: {os.path.getsize(strat_path) / 1024:.0f} KB")

print("\n=== BATCH 2 COMPLETE (Matrice RACI + Document stratégique PACE) ===")
