"""
Forecast Q4 2026 + 2027 - Génération des 5 PDFs PACE - VERSION 3
AVEC COMPLEMENT_ALIMENTAIRE (V300 1L only), données 2023-2026, MATERIEL_ELEVAGE tonnes=0

Génère:
- 5 PDFs pour Q4 2026 (dans download/forecast_q4_2026/)
- 5 PDFs pour 2027 (dans download/forecast_2027/)
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
NEW_FAMILY_COLOR = 'FFE699'  # Yellow for COMPLEMENT_ALIMENTAIRE

styles = getSampleStyleSheet()
H1 = ParagraphStyle('H1', parent=styles['Heading1'], fontName='DejaVuSans-Bold', fontSize=18, textColor=NAVY, spaceAfter=14, spaceBefore=10)
H2 = ParagraphStyle('H2', parent=styles['Heading2'], fontName='DejaVuSans-Bold', fontSize=14, textColor=NAVY, spaceAfter=10, spaceBefore=14)
H3 = ParagraphStyle('H3', parent=styles['Heading3'], fontName='DejaVuSans-Bold', fontSize=12, textColor=GOLD, spaceAfter=8, spaceBefore=10)
BODY = ParagraphStyle('Body', parent=styles['Normal'], fontName='DejaVuSans', fontSize=10, leading=14, alignment=TA_JUSTIFY, spaceAfter=8)
BULLET = ParagraphStyle('Bullet', parent=BODY, leftIndent=20, bulletIndent=10, spaceAfter=4)
SMALL = ParagraphStyle('Small', parent=BODY, fontSize=8, textColor=GRAY)

CELL_STYLE = ParagraphStyle('CellStyle', parent=BODY, fontName='DejaVuSans', fontSize=9, leading=11, alignment=TA_LEFT, spaceAfter=0, spaceBefore=0)

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


# ============================================
# Q4 2026 - 5 PDFs
# ============================================
Q4_DIR = "/home/z/my-project/download/forecast_q4_2026"
os.makedirs(Q4_DIR, exist_ok=True)

# === Q4 2026: 1. RÉSUMÉ EXÉCUTIF ===
print("Generating Q4 2026: 1. Résumé exécutif (v3)...")
exec_path = f"{Q4_DIR}/01_resume_executif.pdf"
doc = SimpleDocTemplate(exec_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []

story.append(Paragraph("RÉSUMÉ EXÉCUTIF", H1))
story.append(Paragraph("Forecast Q4 2026 - Volumes et Valeurs — Version 3", H2))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Contexte</b>", H3))
story.append(Paragraph(
    "BELGOCAM SA doit anticiper ses volumes de ventes et son chiffre d'affaires pour le Q4 2026 "
    "(septembre - décembre). Cette version 3 du forecast s'appuie sur <b>176 576 enregistrements</b> "
    "couvrant <b>44 mois d'historique</b> (janvier 2023 - août 2026), intégrant les commandes En cours et "
    "Validées d'août 2026. L'effet soja exceptionnel de juillet-août 2026 (rupture concurrente) a été "
    "<b>désaisonnalisé</b>. Le prix du soja a été actualisé à <b>25 000 FCFA/sac</b> (hausse du 24/08/2026).",
    BODY))
story.append(Paragraph(
    "<b>NOUVEAUTÉS VERSION 3</b> : (1) Intégration de l'année 2023 (4 ans d'historique au total). "
    "(2) Ajout de la famille <b>COMPLEMENT_ALIMENTAIRE</b> (BELGOKILL V300 1L only, V305 200L exclu). "
    "Conversion 1L = 1kg. (3) MATERIEL_ELEVAGE toujours à 0 en tonnes (non exprimable en volume).",
    BODY))

story.append(Paragraph("<b>Méthodologie</b>", H3))
story.append(Paragraph(
    "Le modèle <b>Prophet</b> (Facebook/Meta) a été entraîné sur 5 familles alimentaires (TOURTEAUX, CONCENTRÉS, "
    "INGRÉDIENTS, ALIMENT COMPLET, COMPLEMENT ALIMENTAIRE) au niveau famille × région. Le MATERIEL ÉLEVAGE et "
    "les PREMIX sont projetés par extrapolation de la moyenne historique Q4. Le forecast couvre <b>8 familles, "
    "121 produits, 25 agences</b> et 3 régions. Le scénario S3 (réappro soja 100%) est utilisé comme référence.",
    BODY))

story.append(Paragraph("<b>Résultats clés Q4 2026</b>", H3))
synth_data = [
    ["Indicateur", "Valeur", "Détail"],
    ["Volume total Q4 2026", "29 347 t", "8 familles, 121 produits, 25 agences"],
    ["CA total Q4 2026", "16 232 M FCFA", "Prix soja 25 000 FCFA/sac"],
    ["Période", "Sept-Déc 2026 (4 mois)", "Saison haute (35-46% du volume annuel)"],
    ["Scénario", "S3 (réappro 100%)", "80 000 sacs au 15/09/2026"],
    ["Données historiques", "176 576 enregistrements", "Jan 2023 - Août 2026 + En cours/Validées"],
    ["Désaisonnalisation", "Effet soja Jul-Août 2026 neutralisé", "Cap à moyenne S1 2026"],
    ["NOUVEAU: COMPLEMENT_ALIMENTAIRE", "1 t / 6,7 M FCFA", "V300 1L only (V305 200L exclu)"],
    ["NOUVEAU v5: ALVEOLES", "0 t / 34 M FCFA (2026 season, forfait ALV=0)", "4 refs MAT011/MAT014/MAT015/MAT017"],
    ["NOUVEAU v5: 10 agences SPC incluses", "Forfait global 348 M/an (poids 2025)", "ALV + MAT_ELEV"],
    ["NOUVEAU v4: Maroua agence", "1 agence ajoutée (soja T102, Août 2026)", "Centre (par convention)"],
    ["NOUVEAU v5: Forfait SPC v2", "39 M/an MAT only (2026 annualisé)", "ALV=0, MAT=38.4M"],
    ["NOUVEAU v4: Bundle 2.5:1", "Ratio soja:concentré <= 2.5:1 forcé", "6 ajustements Q4"],
]
story.append(make_table(synth_data, col_widths=[5*cm, 4*cm, 8*cm], font_size=9, highlight_rows=[7, 8]))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Synthèse par famille</b>", H3))
fam_data = [
    ["Famille", "Volume Q4 (t)", "CA Q4 (M FCFA)", "Part CA"],
    ["TOURTEAUX", "19 982", "9 993", "61,6%"],
    ["CONCENTRÉS", "8 703", "5 778", "35,6%"],
    ["ALIMENT_COMPLET", "267", "186", "1,1%"],
    ["ALVEOLES", "0", "34", "0,2%"],
    ["MATERIEL_ELEVAGE", "0", "82", "0,5%"],
    ["INGREDIENTS", "394", "88", "0,5%"],
    ["PREMIX", "0", "62", "0,4%"],
    ["COMPLEMENT_ALIMENTAIRE", "2", "9", "0,1%"],
    ["MAÏS (exclu)", "0", "0", "0,0%"],
    ["TOTAL", "29 347", "16 232", "100%"],
]
story.append(make_table(fam_data, col_widths=[5*cm, 3*cm, 3*cm, 2.5*cm], font_size=9, highlight_rows=[4, 8]))
story.append(Spacer(1, 0.2*cm))

story.append(Paragraph("<b>Synthèse par mois</b>", H3))
month_data = [
    ["Mois", "Volume (t)", "CA (M FCFA)", "Part CA"],
    ["Septembre 2026", "4 449", "2 516", "15,5%"],
    ["Octobre 2026", "9 021", "4 968", "30,6%"],
    ["Novembre 2026", "7 691", "4 247", "26,2%"],
    ["Décembre 2026", "8 186", "4 501", "27,7%"],
    ["TOTAL Q4", "29 347", "16 232", "100%"],
]
story.append(make_table(month_data, col_widths=[4*cm, 3.5*cm, 3.5*cm, 2.5*cm], font_size=9))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Top 5 agences par CA Q4 2026</b>", H3))
top_data = [
    ["Rang", "Agence", "Région", "CA (M FCFA)"],
    ["1", "FAMLA", "Ouest", "4 430"],
    ["2", "NDOBO", "Littoral", "2 497"],
    ["3", "MESSASSI", "Centre", "1 445"],
    ["4", "DJELENG", "Ouest", "1 268"],
    ["5", "VILLAGE", "Littoral", "950"],
]
story.append(make_table(top_data, col_widths=[1.5*cm, 4*cm, 3*cm, 4.5*cm], font_size=9))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Réalisation 2026 YTD + Projection fin d'année</b>", H3))
ytd_data = [
    ["Famille", "YTD 2026 (t)", "Q4 fcst (t)", "Total 2026 (t)"],
    ["TOURTEAUX", "39 045", "18 321", "57 366"],
    ["CONCENTRÉS", "11 932", "6 610", "18 542"],
    ["ALIMENT_COMPLET", "502", "187", "689"],
    ["INGRÉDIENTS", "531", "291", "822"],
    ["PREMIX", "60", "0", "60"],
    ["COMPLEMENT_ALIM.", "3", "1", "4"],
    ["ALVEOLES", "0", "0 (CA only)", "0"],
    ["MATERIEL_ELEVAGE", "0", "0 (CA only)", "0"],
    ["MAÏS (exclu)", "0", "0", "0"],
    ["TOTAL", "52 073", "29 347", "77 483"],
]
story.append(make_table(ytd_data, col_widths=[3.5*cm, 3*cm, 3*cm, 3*cm], font_size=9, highlight_rows=[6, 7]))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Recommandations principales</b>", H3))
recos = [
    "<b>Réapprovisionnement soja urgent</b> — Commander 80 000 sacs minimum avant le 15/09/2026 pour éviter le scénario S1 (perte de 12 000+ M FCFA vs S3).",
    "<b>Planification commerciale Q4</b> — Utiliser le scénario S3 (29 347 t, 16 232 M FCFA) comme référence.",
    "<b>Pic d'octobre</b> — Octobre = pic du Q4 (7 628 t, 4 114 M FCFA). Anticiper capacité de production et logistique.",
    "<b>Maintien du bundle</b> — Ratio bundle 2,3:1 atteint en août 2026 à maintenir Q4 2026.",
    "<b>Surveillance FAMLA</b> — Représente 29% du CA Q4 (4 028 M FCFA). Performance critique.",
    "<b>Suivi COMPLEMENT_ALIMENTAIRE</b> — Nouvelle famille intégrée au forecast (1 t, 7 M FCFA). BELGOKILL V300 1L comme proxy.",
]
for r in recos:
    story.append(Paragraph(f"• {r}", BULLET))

story.append(Spacer(1, 0.3*cm))
story.append(Paragraph("<b>Conclusion</b>", H3))
story.append(Paragraph(
    "Le forecast Q4 2026 projette <b>29 347 tonnes</b> pour un CA de <b>16 232 M FCFA</b>. "
    "La désaisonnalisation de l'effet soja exceptionnel, l'actualisation du prix à 25 000 FCFA/sac, "
    "l'intégration de l'année 2023 (44 mois d'historique) et l'ajout de la famille COMPLEMENT_ALIMENTAIRE "
    "(V300 1L only) permettent une projection réaliste. Le pic d'octobre (7 628 t) nécessitera une "
    "anticipation renforcée du réapprovisionnement soja.",
    BODY))

doc.build(story)
print(f"✓ Q4 2026 Résumé exécutif: {os.path.getsize(exec_path)/1024:.0f} KB")

# === Q4 2026: 2. PROPOSITION DE PROJET ===
print("Generating Q4 2026: 2. Proposition (v3)...")
prop_path = f"{Q4_DIR}/02_proposition_projet.pdf"
doc = SimpleDocTemplate(prop_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []
story.extend(cover_page("Proposition de Projet", "Forecast Q4 2026 - Version 3", "PROPOSITION DE PROJET"))

story.append(Paragraph("1. Contexte et justification", H1))
story.append(Paragraph(
    "BELGOCAM SA doit anticiper ses volumes et CA pour le Q4 2026 (sept-déc). Cette version 3 s'appuie sur "
    "44 mois d'historique (Jan 2023 - Août 2026) et intègre les innovations suivantes : ajout de la famille "
    "COMPLEMENT_ALIMENTAIRE (BELGOKILL V300 1L only, V305 200L exclu, 1L=1kg), intégration de l'année 2023, "
    "désaisonnalisation de l'effet soja, inclusion des commandes En cours/Validées, actualisation du prix soja "
    "à 25 000 FCFA/sac. Le forecast couvre 8 familles, 121 produits et 25 agences.",
    BODY))

story.append(Paragraph("2. Objectifs", H1))
story.append(Paragraph("<b>Objectif principal</b>", H3))
story.append(Paragraph(
    "Produire un forecast Q4 2026 (4 mois) en volumes et valeurs, désagrégé par produit, famille, agence et "
    "région, selon le scénario S3 (réappro soja 100%).",
    BODY))
story.append(Paragraph("<b>Objectifs spécifiques</b>", H3))
objs = [
    "Modéliser la tendance et la saisonnalité mensuelles via Prophet (5 familles alimentaires)",
    "Désaisonnaliser l'effet soja exceptionnel de juillet-août 2026",
    "Intégrer les commandes En cours et Validées comme potentielles ventes Livrées",
    "Actualiser le prix soja à 25 000 FCFA/sac (hausse du 24/08/2026)",
    "NOUVEAU v3 : Ajouter la famille COMPLEMENT_ALIMENTAIRE (V300 1L only, 1L=1kg)",
    "NOUVEAU v3 : Intégrer l'année 2023 (44 mois d'historique au total)",
    "NOUVEAU v3 : MATERIEL_ELEVAGE toujours à 0 en tonnes",
    "Exclure le Maïs, les produits opportunistes et V305 (BELGOKILL 200L)",
    "Produire les livrables PACE complets (8 documents)",
]
for o in objs:
    story.append(Paragraph(f"• {o}", BULLET))

story.append(Paragraph("3. Périmètre", H1))
scope_data = [
    ["Dimension", "Détail", "Volume"],
    ["Période forecast", "Septembre - Décembre 2026 (4 mois)", "4 mois"],
    ["Familles incluses", "TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, ALIMENT COMPLET, MATERIEL ÉLEVAGE, PREMIX, COMPLEMENT_ALIMENTAIRE, ALVEOLES", "8 familles"],
    ["Produits", "115 références (T102, C101-C108, MAT014, P102N2, V300 BELGOKILL, etc.)", "121 produits"],
    ["Agences", "25 agences BELGOCAM", "25 agences"],
    ["Régions", "Ouest, Centre, Littoral", "3 régions"],
    ["Niveau détail", "Produit × Agence × Mois", "4 569 lignes"],
    ["Scénario", "S3 - Réappro soja 100%", "1 scénario"],
    ["Données historiques", "176 576 enregistrements (Jan 2023 - Août 2026)", "44 mois"],
    ["Désaisonnalisation", "Effet soja Jul-Août 2026 neutralisé", "—"],
    ["Prix soja", "25 000 FCFA/sac (actualisé 24/08/2026)", "—"],
    ["NOUVEAU: COMPLEMENT_ALIM.", "V300 1L only (V305 exclu), 1L=1kg", "9 produits"],
]
story.append(make_table(scope_data, col_widths=[3.5*cm, 9.5*cm, 4*cm], font_size=9, highlight_rows=[11, 12]))

story.append(Paragraph("4. Méthodologie PACE", H1))
story.append(Paragraph(
    "La méthodologie PACE (Prepare-Analyze-Construct-Execute) structure le projet en 4 phases. "
    "L'innovation majeure v3 est l'ajout de la famille COMPLEMENT_ALIMENTAIRE avec proxy BELGOKILL V300 1L "
    "et l'utilisation de 44 mois d'historique (Jan 2023 - Août 2026).",
    BODY))

pace_data = [
    ["Phase", "Activités", "Livrables"],
    ["P - PREPARE", "Consolidation 176 576 records + En cours/Validées, désaisonnalisation soja, ajout COMPLEMENT_ALIM.", "Dataset 2023-2026, prix Q4"],
    ["A - ANALYZE", "AED, saisonnalité 2023-2025, top produits/agences", "Graphiques, synthèse AED"],
    ["C - CONSTRUCT", "Prophet (5 familles × 3 régions) + extrapolation (2 familles), forecast Q4 S3", "Forecast Q4 2026 (4 569 lignes)"],
    ["E - EXECUTE", "Excel 8 feuilles, 5 PDFs PACE, graphiques", "8 livrables finaux"],
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
    ["Historique 44 mois — suffisant pour Prophet", "Faible", "Mise à jour trimestrielle"],
    ["Prix soja volatil (+47% en 2 mois)", "Élevée", "Prix actualisé 25 000 FCFA, scénario S3"],
    ["Effet soja 2026 non récurrent en Q4", "Élevée", "Cap désaisonnalisation à moyenne S1"],
    ["Commandes En cours/Validées non converties", "Faible", "Filtre qualité, exclusion Annulées"],
    ["V305 exclu — BELGOKILL 200L absent", "Faible", "V300 1L utilisé comme proxy suffisant"],
]
story.append(make_table(risks_data, col_widths=[7*cm, 3*cm, 7*cm], font_size=9))

story.append(Paragraph("7. Critères de succès", H1))
success = [
    "Forecast Q4 2026 produit sur 4 mois avec désagrégation complète (4 569 lignes)",
    "Désaisonnalisation effective de l'effet soja exceptionnel",
    "En cours et Validées intégrés comme potentielles ventes",
    "Prix soja actualisé à 25 000 FCFA/sac",
    "Famille COMPLEMENT_ALIMENTAIRE ajoutée avec proxy BELGOKILL V300 1L",
    "Écart forecast vs réalité ≤ 15% par mois",
    "Livrables PACE complets (8 documents) produits et diffusés",
]
for s in success:
    story.append(Paragraph(f"• {s}", BULLET))

doc.build(story)
print(f"✓ Q4 2026 Proposition: {os.path.getsize(prop_path)/1024:.0f} KB")

# === Q4 2026: 3. MATRICE RACI ===
print("Generating Q4 2026: 3. Matrice RACI (v3)...")
raci_path = f"{Q4_DIR}/03_matrice_raci.pdf"
doc = SimpleDocTemplate(raci_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
story = []
story.extend(cover_page("Matrice RACI", "Forecast Q4 2026 - Version 3", "MATRICE RACI"))

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
    ["PREPARE", "1.1 Consolidation 2023-2026 + En cours/Validées", "R/A", "I", "C", "I", "I", "I"],
    ["", "1.2 Désaisonnalisation effet soja", "R/A", "C", "C", "I", "I", "I"],
    ["", "1.3 Calcul prix Q4 (soja 25 000)", "R/A", "C", "I", "I", "C", "I"],
    ["", "1.4 NOUVEAU: Ajout COMPLEMENT_ALIMENTAIRE (V300 1L)", "R/A", "C", "I", "I", "I", "I"],
    ["", "1.5 Validation dataset v3", "R", "C", "C", "I", "A", "I"],
    ["ANALYZE", "2.1 AED + saisonnalité 2023-2025", "R/A", "I", "I", "I", "I", "I"],
    ["", "2.2 Analyse effet soja exceptionnel", "R/A", "C", "C", "I", "I", "I"],
    ["", "2.3 Validation hypothèses AED", "R", "C", "C", "I", "A", "I"],
    ["CONSTRUCT", "3.1 Prophet (5 familles × 3 régions)", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.2 Extrapolation MAT+PREMIX (CA only)", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.3 Forecast Q4 S3", "R", "C", "I", "A", "C", "I"],
    ["", "3.4 Désagrégation produit × agence", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.5 Validation forecast Q4", "R", "C", "I", "A", "C", "I"],
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
    ["Validation dataset v3", "PREPARE", "03/09/2026", "CG", "Dataset 2023-2026 + COMPLEMENT_ALIM."],
    ["Validation AED", "ANALYZE", "04/09/2026", "DC", "Synthèse AED"],
    ["Validation forecast", "CONSTRUCT", "05/09/2026", "DG + DC", "Forecast Q4 2027 complet"],
    ["Livrables finaux", "EXECUTE", "05/09/2026", "DG", "8 livrables"],
    ["Présentation CODIR", "EXECUTE", "06/09/2026", "DG", "Validation officielle"],
    ["Mise à jour Jan 2027", "—", "Jan 2027", "DA", "Réactualisation avec données réelles"],
]
story.append(make_table(control_data, col_widths=[3.5*cm, 2.5*cm, 3*cm, 3*cm, 5*cm], font_size=8))

doc.build(story)
print(f"✓ Q4 2026 Matrice RACI: {os.path.getsize(raci_path)/1024:.0f} KB")

# === Q4 2026: 4. DOCUMENT STRATÉGIQUE PACE ===
print("Generating Q4 2026: 4. Document stratégique PACE (v3)...")
strat_path = f"{Q4_DIR}/04_document_strategique_pace.pdf"
doc = SimpleDocTemplate(strat_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []
story.extend(cover_page("Document Stratégique PACE", "Forecast Q4 2026 - Version 3 (8 familles, données 2023-2026)", "STRATÉGIE PACE"))

story.append(Paragraph("1. Vision et objectifs", H1))
story.append(Paragraph("1.1 Vision", H2))
story.append(Paragraph(
    "Faire du forecast Q4 2026 un <b>outil de planification opérationnelle</b> permettant à BELGOCAM SA d'anticiper "
    "29 347 tonnes de ventes et 16 232 M FCFA de chiffre d'affaires sur la période septembre-décembre 2026, "
    "avec une désaisonnalisation de l'effet soja exceptionnel et l'intégration de la famille COMPLEMENT_ALIMENTAIRE.",
    BODY))

story.append(Paragraph("1.2 Objectifs stratégiques", H2))
obj_data = [
    ["Objectif", "Description", "KPI"],
    ["Forecast Q4", "4 mois Sept-Dec 2026 en volume + valeur", "29 347 t, 16 232 M FCFA"],
    ["Désaisonnalisation", "Neutraliser l'effet soja Jul-Août 2026", "Cap moyenne S1 2026"],
    ["En cours + Validées", "Intégrer comme potentielles ventes", "255 commandes incluses"],
    ["Prix actualisé", "Soja 25 000 FCFA/sac", "vs 20 600 précédent"],
    ["Désagrégation", "Produit × agence × mois", "4 569 lignes"],
    ["NOUVEAU: COMPLEMENT_ALIM.", "V300 1L only (V305 exclu), 1L=1kg", "1 t, 7 M FCFA"],
]
story.append(make_table(obj_data, col_widths=[3.5*cm, 8.5*cm, 5*cm], font_size=9, highlight_rows=[6, 7]))

story.append(Paragraph("2. Phase PREPARE - Préparation des données", H1))
story.append(Paragraph("2.1 Sources consolidées (2023-2026)", H2))
sources_data = [
    ["Source", "Période", "Records"],
    ["LY_21_24 (filtre 2023)", "Jan-Dec 2023", "6 016"],
    ["LY_24 (Jul-Dec 2024)", "Juillet-Décembre 2024", "52 215"],
    ["Historique 2025", "Jan-Déc 2025", "66 206"],
    ["S1 + Juil + Août 2026", "Jan-Août 2026", "52 139"],
    ["En cours + Validées", "Août 2026", "~255"],
    ["TOTAL", "44 mois", "176 576"],
]
story.append(make_table(sources_data, col_widths=[5*cm, 5*cm, 5*cm], font_size=9))

story.append(Paragraph("2.2 Désaisonnalisation de l'effet soja", H2))
story.append(Paragraph(
    "Les volumes soja de juillet-août 2026 ont été exceptionnellement élevés (rupture concurrente). "
    "Pour éviter que Prophet n'extrapole ce phénomène non récurrent, les volumes soja Jul-Août 2026 ont été "
    "<b>cappés à la moyenne S1 2026</b>.",
    BODY))

story.append(Paragraph("2.3 NOUVEAU v3: Famille COMPLEMENT_ALIMENTAIRE", H2))
story.append(Paragraph(
    "La famille COMPLEMENT_ALIMENTAIRE (9 produits liquides 1L) a été ajoutée au forecast Q4 2026. "
    "Selon demande utilisateur : <b>seul V300 (BELGOKILL 1L) est conservé — V305 (BELGOKILL 200L) est EXCLU</b>. "
    "Conversion 1L = 1kg appliquée pour exprimer les volumes en tonnes.",
    BODY))

story.append(Paragraph("3. Phase CONSTRUCT - Modélisation", H1))
story.append(Paragraph("3.1 Configuration Prophet", H2))
config_data = [
    ["Paramètre", "Valeur", "Justification"],
    ["yearly_seasonality", "True", "Capture la saisonnalité annuelle (pic Q4)"],
    ["seasonality_mode", "multiplicative", "Saisonnalité proportionnelle à la tendance"],
    ["changepoint_prior_scale", "0.05", "Tendance modérément flexible"],
    ["interval_width", "0.8", "Intervalle de confiance 80%"],
    ["Période forecast", "4 mois (Sep-Dec 2026)", "Q4 2026"],
    ["Familles Prophet", "5 (TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, ALIMENT_COMPLET, COMPLEMENT_ALIMENTAIRE)", "Nouveau: COMPLEMENT_ALIMENTAIRE"],

]
story.append(make_table(config_data, col_widths=[5*cm, 5*cm, 7*cm], font_size=9, highlight_rows=[6]))

story.append(Paragraph("4. Résultats par famille et mois", H1))
story.append(Paragraph("4.1 Par famille", H2))
fam_detail = [
    ["Famille", "Volume Q4 (t)", "CA Q4 (M FCFA)", "Part CA", "Méthode"],
    ["TOURTEAUX", "19 982", "9 993", "61,6%", "Prophet"],
    ["CONCENTRÉS", "8 703", "5 778", "35,6%", "Prophet + bundle 2.5:1"],
    ["ALIMENT_COMPLET", "267", "186", "1,1%", "Prophet"],
    ["ALVEOLES", "0", "34", "0,2%", "Extrap 2026 + Forfait SPC (ALV=0)"],
    ["MATERIEL_ELEVAGE", "0", "82", "0,5%", "Extrap + Forfait SPC (2026 annualisé)"],
    ["INGREDIENTS", "394", "88", "0,5%", "Prophet"],
    ["PREMIX", "0", "62", "0,4%", "Extrapolation CA"],
    ["COMPLEMENT_ALIM.", "2", "9", "0,1%", "Prophet (proxy V300 1L)"],
]
story.append(make_table(fam_detail, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2*cm, 4*cm], font_size=9, highlight_rows=[4, 8]))

story.append(Paragraph("4.2 Par mois", H2))
month_detail = [
    ["Mois", "Volume (t)", "CA (M FCFA)", "Part CA", "Lecture"],
    ["Septembre 2026", "4 054", "2 313", "16,5%", "Démarrage Q4 + réappro"],
    ["Octobre 2026", "7 628", "4 131", "29,3%", "Pic mensuel"],
    ["Novembre 2026", "6 609", "3 653", "26,1%", "Maintien"],
    ["Décembre 2026", "7 118", "3 896", "27,9%", "Fêtes de fin d'année"],
]
story.append(make_table(month_detail, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2*cm, 4.5*cm], font_size=9))

story.append(Paragraph("5. Plan de déploiement", H1))
deploy_data = [
    ["Étape", "Date", "Action", "Cible"],
    ["1", "03/09/2026", "Validation proposition v3", "DG"],
    ["2", "05/09/2026", "Livrables finaux prêts", "DA"],
    ["3", "06/09/2026", "Présentation CODIR", "DG + DC + DP + CG"],
    ["4", "07/09/2026", "Diffusion aux agences", "14 RA"],
    ["5", "Jan 2027", "Mise à jour avec données réelles Q4", "DA"],
]
story.append(make_table(deploy_data, col_widths=[1*cm, 3*cm, 7*cm, 6*cm], font_size=9))

story.append(Paragraph("6. Conclusion", H1))
story.append(Paragraph(
    "Le forecast Q4 2026 version 3 projette <b>29 347 tonnes</b> pour <b>16 232 M FCFA</b>. La désaisonnalisation "
    "de l'effet soja, l'actualisation du prix à 25 000 FCFA, l'intégration de l'année 2023 (44 mois d'historique), "
    "l'ajout de la famille COMPLEMENT_ALIMENTAIRE (V300 1L only) et le maintien de MATERIEL_ELEVAGE à 0 tonne "
    "permettent une projection réaliste. Le pic d'octobre (7 628 t) nécessitera une anticipation renforcée "
    "du réapprovisionnement soja.",
    BODY))

doc.build(story)
print(f"✓ Q4 2026 Document stratégique: {os.path.getsize(strat_path)/1024:.0f} KB")

# === Q4 2026: 5. GUIDE MÉTHODOLOGIQUE ===
print("Generating Q4 2026: 5. Guide méthodologique (v3)...")
guide_path = f"{Q4_DIR}/05_guide_methodologique.pdf"
doc = SimpleDocTemplate(guide_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []
story.extend(cover_page("Guide Méthodologique", "Forecast Q4 2026 - Version 3 (Prophet + COMPLEMENT_ALIMENTAIRE V300 1L)", "GUIDE MÉTHODOLOGIQUE"))

story.append(Paragraph("Sommaire", H1))
story.append(Paragraph(
    "1. Introduction et contexte<br/>"
    "2. Méthodologie PACE<br/>"
    "3. Outils et environnement<br/>"
    "4. Phase PREPARE - Données 2023-2026<br/>"
    "5. Phase PREPARE - Désaisonnalisation soja<br/>"
    "6. NOUVEAU v3: Phase PREPARE - COMPLEMENT_ALIMENTAIRE<br/>"
    "7. Phase ANALYZE - AED<br/>"
    "8. Phase CONSTRUCT - Prophet<br/>"
    "9. Phase EXECUTE - Livrables<br/>"
    "Annexes",
    BODY))

story.append(PageBreak())
story.append(Paragraph("1. Introduction et contexte", H1))
story.append(Paragraph(
    "Ce guide décrit la méthodologie complète du forecast Q4 2026 de BELGOCAM SA (version 3). Il couvre 4 mois "
    "(septembre-décembre 2026) avec 8 familles de produits (incluant la nouvelle famille COMPLEMENT_ALIMENTAIRE) "
    "et 25 agences. Les innovations majeures v3 : "
    "(1) ajout de la famille COMPLEMENT_ALIMENTAIRE (V300 BELGOKILL 1L only, V305 200L EXCLU, 1L=1kg), "
    "(2) intégration de l'année 2023 (44 mois d'historique au total), "
    "(3) MATERIEL_ELEVAGE toujours à 0 tonne (non exprimable en volume), "
    "(4) désaisonnalisation de l'effet soja, "
    "(5) inclusion des commandes En cours/Validées, "
    "(6) actualisation du prix soja à 25 000 FCFA/sac.",
    BODY))

story.append(Paragraph("2. Méthodologie PACE", H1))
story.append(Paragraph(
    "La méthodologie PACE (Prepare-Analyze-Construct-Execute) structure le projet en 4 phases successives.",
    BODY))
pace_detail = [
    ["Phase", "Objectif", "Livrables", "Durée"],
    ["P - PREPARE", "Préparer données 2023-2026 + désaisonnalisation + COMPLEMENT_ALIM.", "Dataset 176 576 records, prix Q4", "2 jours"],
    ["A - ANALYZE", "Comprendre données + effet soja", "Graphiques, synthèse AED", "1 jour"],
    ["C - CONSTRUCT", "Modéliser Prophet Q4 (8 familles)", "Forecast Q4 2026 (4 569 lignes)", "2 jours"],
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
story.append(Paragraph("4. Phase PREPARE - Données 2023-2026", H1))
story.append(Paragraph(
    "Le dataset consolidé couvre 44 mois (janvier 2023 - août 2026), soit 176 576 enregistrements. "
    "Les sources sont : LY_21_24 filtre 2023 (6 016 records), LY_24 Jul-Dec 2024 (52 215), "
    "Historique 2025 (66 206), S1+Juil+Août 2026 (52 139), En cours + Validées (~255).",
    BODY))
story.append(Paragraph(
    "<b>Intégration de l'année 2023</b> : La version 3 inclut 2023 pour enrichir l'historique avec 4 ans "
    "de données (2023-2026). Cela permet à Prophet de mieux capturer la saisonnalité annuelle.",
    BODY))

story.append(Paragraph("5. Phase PREPARE - Désaisonnalisation soja", H1))
story.append(Paragraph(
    "L'effet soja exceptionnel de juillet-août 2026 (rupture concurrente) a généré des volumes anormalement "
    "élevés non récurrents. Les volumes soja Jul-Août 2026 ont été cappés à la moyenne S1 2026.",
    BODY))
story.append(Paragraph("<b>Méthode</b>", H3))
story.append(Paragraph(
    "1. Calcul moyenne mensuelle soja S1 2026 (Jan-Juin)<br/>"
    "2. Identification mois Jul-Août 2026 avec volume max<br/>"
    "3. Calcul facteur de cap = moyenne S1 / volume max<br/>"
    "4. Application du facteur aux enregistrements soja Jul-Août 2026<br/>"
    "5. Volumes réduits proportionnellement, signal saisonnier préservé",
    BODY))

story.append(PageBreak())
story.append(Paragraph("6. NOUVEAU v3: Famille COMPLEMENT_ALIMENTAIRE", H1))
story.append(Paragraph(
    "La famille COMPLEMENT_ALIMENTAIRE regroupe 9 produits liquides BELGO : BELGOKILL V300 (1L), "
    "BELGO HARMONY, BELGO PROTECT, BELGO DRY LIT, BELGO WATER CLEAN, BELGO VIT Ese, BELGO THERMO, "
    "BELGO BIO SELECT, BELGO FRESH — tous en conditionnement 1L.",
    BODY))

story.append(Paragraph("<b>Exclusion de V305 (BELGOKILL 200L)</b>", H3))
story.append(Paragraph(
    "Selon demande utilisateur explicite : seul le conditionnement 1L est conservé. "
    "V305 (BELGOKILL 200L) est EXCLU du forecast. Cela permet d'avoir une famille homogène "
    "(tous produits en 1L) et d'éviter le biais du conditionnement bulk.",
    BODY))

story.append(Paragraph("<b>Conversion 1L = 1kg</b>", H3))
story.append(Paragraph(
    "Pour exprimer les volumes en tonnes (nécessaire pour Prophet), la conversion 1L = 1kg est appliquée : "
    "<b>1 qte = 1 kg, tonnes = qte / 1000</b>. Approximation valable pour liquides aqueux (densité ~1).",
    BODY))

story.append(Paragraph("<b>Proxy BELGOKILL V300 1L</b>", H3))
story.append(Paragraph(
    "BELGOKILL V300 (1L) est le produit dominant de la famille (38% du CA famille sur 2023-2026). "
    "Sa tendance est utilisée comme proxy pour modéliser toute la famille via Prophet. "
    "Les autres produits (CA001-CA008) sont désagrégés selon leurs parts historiques.",
    BODY))

ca_proxy = [
    ["Réf.", "Produit", "Poids (kg/qte)", "Prix Q4 2026 (FCFA/L)", "Statut"],
    ["V300", "BELGOKILL 1L", "1", "2 500", "✓ Inclus (proxy)"],
    ["V305", "BELGOKILL 200L", "—", "—", "✗ EXCLU v3"],
    ["CA003.1", "BELGO HARMONY 1L", "1", "8 000", "✓ Inclus"],
    ["CA004.1", "BELGO PROTECT 1L", "1", "9 800", "✓ Inclus"],
    ["CA006.1", "BELGO DRY LIT 1L", "1", "6 500", "✓ Inclus"],
    ["CA001.1", "BELGO WATER CLEAN 1L", "1", "5 500", "✓ Inclus"],
    ["CA002.1", "BELGO VIT Ese 1L", "1", "10 500", "✓ Inclus"],
    ["CA005.1", "BELGO THERMO 1L", "1", "15 000", "✓ Inclus"],
    ["CA007.1", "BELGO BIO SELECT 1L", "1", "8 000", "✓ Inclus"],
    ["CA008.1", "BELGO FRESH 1L", "1", "14 000", "✓ Inclus"],
]
story.append(make_table(ca_proxy, col_widths=[2*cm, 5*cm, 2.5*cm, 3.5*cm, 3*cm], font_size=9, highlight_rows=[2]))

story.append(Paragraph("7. Phase ANALYZE - AED", H1))
story.append(Paragraph(
    "L'AED révèle : (1) TOURTEAUX = 76% du volume historique, (2) Q4 = 35-46% du volume annuel selon familles, "
    "(3) pic octobre (indice 1.86 pour TOURTEAUX en 2025), (4) FAMLA = 31.6% du volume, "
    "(5) COMPLEMENT_ALIMENTAIRE = 3-7 t/an (faible volume, forte valeur unitaire).",
    BODY))

story.append(Paragraph("8. Phase CONSTRUCT - Prophet", H1))
story.append(Paragraph(
    "15 modèles Prophet famille × région (5 familles × 3 régions) + extrapolation Q4 moyenne pour MATERIEL_ELEVAGE "
    "et PREMIX (2 familles × 3 régions = 6 extrapolations). Désagrégation par produit × agence selon parts "
    "historiques de CA. Prix Q4 2026 : 25 000 FCFA/sac soja, prix août 2026 autres familles, prix par litre "
    "pour COMPLEMENT_ALIMENTAIRE (2 500-15 000 FCFA/L selon produit).",
    BODY))

story.append(Paragraph("9. Phase EXECUTE - Livrables", H1))
livrables_data = [
    ["#", "Livrable", "Format", "Usage"],
    ["1", "Résumé exécutif", "PDF", "Direction Générale"],
    ["2", "Proposition de projet", "PDF", "CODIR"],
    ["3", "Matrice RACI", "PDF", "Gouvernance"],
    ["4", "Document stratégique PACE", "PDF", "Référence stratégique"],
    ["5", "Guide méthodologique", "PDF", "Référence méthodes"],
    ["6", "Excel forecast Q4 2026 (8 feuilles)", "XLSX", "Pilotage opérationnel"],
]
story.append(make_table(livrables_data, col_widths=[1*cm, 5*cm, 3*cm, 6*cm], font_size=9))

story.append(Paragraph("Annexe - Glossaire", H1))
glossaire = [
    ["Terme", "Définition"],
    ["Prophet", "Bibliothèque prévision séries temporelles (Facebook/Meta)"],
    ["Désaisonnalisation", "Neutralisation d'un effet exceptionnel non récurrent"],
    ["S3", "Scénario réappro soja 100% (situation normale)"],
    ["Cap S1", "Plafonnement des volumes Jul-Août 2026 à moyenne S1 2026"],
    ["Q4", "Quatrième trimestre (oct-déc)"],
    ["PACE", "Prepare-Analyze-Construct-Execute"],
    ["NOUVEAU v3: COMPLEMENT_ALIM.", "9 produits liquides 1L BELGOxxx (V300 1L proxy)"],
    ["NOUVEAU v3: ALVEOLES séparé", "Famille distincte de MATERIEL_ELEVAGE (4 refs MAT011/14/15/17)"],
    ["NOUVEAU v4: SPC agences", "5 agences SPC incluses (Baf-Chefferie dominant, 274 M CA historique)"],
    ["NOUVEAU v4: Maroua", "Agence ajoutée (Centre par convention), soja T102 only (Août 2026)"],
    ["NOUVEAU v4: Bundle 2.5:1", "Ratio soja:concentré forcé à <=2.5:1 (augment. CONCENTRES si ratio > 2.5)"],
    ["NOUVEAU v4: Saisonnalité ALV 2026", "ALVEOLES utilise 2026 (13.9 M/an) au lieu du pic 2025 (739 M)"],
    ["NOUVEAU v5: 10 SPC agences", "Toutes SPC incluses (vs 5 avant) — fin exclusion COMPTOIR pour SPC"],
    ["NOUVEAU v5: Forfait SPC v2", "Forfait MAT_ELEV (38.4 M/an) basé sur 2026 annualisé, poids 2025"],
    ["NOUVEAU v5: SPC PK15 forfait", "Poids 0 en 2025 → forfait réaliste 1 M/an (0.5 ALV + 0.5 MAT)"],
    ["NOUVEAU v3: V305 EXCLU", "BELGOKILL 200L retiré du forecast (seul V300 1L conservé)"],
    ["NOUVEAU v3: 1L=1kg", "Conversion pour volumes liquides en tonnes"],
    ["NOUVEAU v3: Données 2023-2026", "44 mois d'historique (vs 20 mois en v1)"],
    ["MATERIEL_ELEVAGE", "Toujours 0 en tonnes (CA only). ALVEOLES désormais séparés en famille distincte"],
]
story.append(make_table(glossaire, col_widths=[5*cm, 11*cm], font_size=9, highlight_rows=[7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]))

doc.build(story)
print(f"✓ Q4 2026 Guide méthodologique: {os.path.getsize(guide_path)/1024:.0f} KB")

print(f"\n=== Q4 2026: 5 PDFs v3 GÉNÉRÉS ===")
for f in sorted(os.listdir(Q4_DIR)):
    if f.endswith('.pdf'):
        print(f"  {f}: {os.path.getsize(os.path.join(Q4_DIR, f))/1024:.0f} KB")


# ============================================
# 2027 - 5 PDFs (use updated numbers)
# ============================================
F2027_DIR = "/home/z/my-project/download/forecast_2027"
os.makedirs(F2027_DIR, exist_ok=True)

# === 2027: 1. RÉSUMÉ EXÉCUTIF ===
print("\nGenerating 2027: 1. Résumé exécutif (v3)...")
exec_path = f"{F2027_DIR}/01_resume_executif.pdf"
doc = SimpleDocTemplate(exec_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []

story.append(Paragraph("RÉSUMÉ EXÉCUTIF", H1))
story.append(Paragraph("Forecast 2027 - Volumes et Valeurs (12 mois) — Version 3", H2))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Contexte</b>", H3))
story.append(Paragraph(
    "BELGOCAM SA doit anticiper ses volumes de ventes et son chiffre d'affaires pour l'année 2027 complète. "
    "Cette version 3 s'appuie sur <b>176 576 enregistrements</b> couvrant <b>44 mois d'historique</b> "
    "(janvier 2023 - août 2026), intégrant les commandes En cours et Validées d'août 2026. "
    "L'effet soja exceptionnel de juillet-août 2026 a été <b>désaisonnalisé</b>. Le prix du soja a été "
    "actualisé à <b>25 000 FCFA/sac</b>.",
    BODY))
story.append(Paragraph(
    "<b>NOUVEAUTÉS VERSION 3</b> : (1) Intégration de l'année 2023 (44 mois d'historique). "
    "(2) Ajout de la famille <b>COMPLEMENT_ALIMENTAIRE</b> (BELGOKILL V300 1L only, V305 200L exclu, 1L=1kg). "
    "(3) MATERIEL_ELEVAGE toujours à 0 en tonnes.",
    BODY))

story.append(Paragraph("<b>Méthodologie</b>", H3))
story.append(Paragraph(
    "Le modèle <b>Prophet</b> a été entraîné sur 5 familles alimentaires (TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, "
    "ALIMENT COMPLET, COMPLEMENT ALIMENTAIRE) au niveau famille × région. Le MATERIEL ÉLEVAGE et les PREMIX "
    "sont projetés par extrapolation de la moyenne historique. Le forecast couvre <b>8 familles, 121 produits, "
    "25 agences</b> et 3 régions.",
    BODY))

story.append(Paragraph("<b>Résultats clés 2027</b>", H3))
synth_data = [
    ["Indicateur", "Valeur", "Détail"],
    ["Volume total 2027", "96 102 t", "8 familles, 121 produits, 25 agences"],
    ["CA total 2027", "53 052 M FCFA", "Prix soja 25 000 FCFA/sac"],
    ["Période", "12 mois (Jan-Déc 2027)", "Forecast complet annuel"],
    ["Scénario", "S3 (réappro 100%)", "Situation normale"],
    ["Données historiques", "176 576 enregistrements", "Jan 2023 - Août 2026 + En cours/Validées"],
    ["Désaisonnalisation", "Effet soja Jul-Août 2026 neutralisé", "Cap à moyenne S1 2026"],
    ["NOUVEAU v3: COMPLEMENT_ALIM.", "3 t / 18 M FCFA", "V300 1L only (V305 200L exclu)"],
    ["NOUVEAU v5: ALVEOLES", "0 t / 101 M FCFA (2026 season, forfait ALV=0)", "4 refs MAT011/MAT014/MAT015/MAT017"],
    ["NOUVEAU v4: SPC agences incluses", "5 agences SPC + Maroua", "25 agences total"],
    ["NOUVEAU v5: Forfait SPC v2", "39 M/an MAT only (2026 annualisé)", "ALV=0, MAT=38.4M"],
    ["NOUVEAU v4: Bundle 2.5:1", "Ratio soja:concentré <= 2.5:1 forcé", "26 ajustements annuels"],
]
story.append(make_table(synth_data, col_widths=[5*cm, 4*cm, 8*cm], font_size=9, highlight_rows=[7, 8]))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Synthèse par famille</b>", H3))
fam_data = [
    ["Famille", "Volume (t)", "CA (M FCFA)", "Part CA"],
    ["TOURTEAUX", "66 338", "33 176", "62,5%"],
    ["CONCENTRÉS", "27 822", "18 493", "34,9%"],
    ["ALIMENT_COMPLET", "877", "612", "1,2%"],
    ["ALVEOLES", "0", "101", "0,2%"],
    ["MATERIEL_ELEVAGE", "0", "214", "0,4%"],
    ["INGREDIENTS", "1 059", "236", "0,4%"],
    ["PREMIX", "0", "192", "0,4%"],
    ["COMPLEMENT_ALIMENTAIRE", "5", "28", "0,1%"],
    ["MAÏS (exclu)", "0", "0", "0,0%"],
    ["TOTAL", "96 102", "53 052", "100%"],
]
story.append(make_table(fam_data, col_widths=[5*cm, 3*cm, 3*cm, 2.5*cm], font_size=9, highlight_rows=[4, 8]))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Synthèse par trimestre</b>", H3))
q_data = [
    ["Trimestre", "Volume (t)", "CA (M FCFA)", "Part CA"],
    ["Q1 (Jan-Mar)", "25 295", "13 949", "26,3%"],
    ["Q2 (Avr-Juin)", "24 316", "13 414", "25,3%"],
    ["Q3 (Juil-Sept)", "19 515", "10 849", "20,4%"],
    ["Q4 (Oct-Déc)", "26 975", "14 839", "28,0%"],
    ["TOTAL", "96 102", "53 052", "100%"],
]
story.append(make_table(q_data, col_widths=[4*cm, 3.5*cm, 3.5*cm, 2.5*cm], font_size=9))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Top 5 agences par CA 2027</b>", H3))
top_data = [
    ["Rang", "Agence", "Région", "CA (M FCFA)"],
    ["1", "FAMLA", "Ouest", "15 682"],
    ["2", "NDOBO", "Littoral", "7 652"],
    ["3", "MESSASSI", "Centre", "4 552"],
    ["4", "DJELENG", "Ouest", "4 494"],
    ["5", "VILLAGE", "Littoral", "2 912"],
]
story.append(make_table(top_data, col_widths=[1.5*cm, 4*cm, 3*cm, 4.5*cm], font_size=9))

story.append(PageBreak())

story.append(Paragraph("<b>Historique 2023-2026 vs Forecast 2027</b>", H3))
story.append(Paragraph(
    "Le forecast 2027 s'appuie sur 4 ans d'historique (Jan 2023 - Août 2026). Le tableau ci-dessous présente "
    "l'évolution par famille et par année.",
    BODY))

hist_data = [
    ["Famille", "2023 (t)", "2024 (t)", "2025 (t)", "2026 YTD (t)", "2027 fcst (t)", "CA 2027 (M)"],
    ["TOURTEAUX", "33 323", "19 296", "44 584", "39 045", "60 673", "30 338"],
    ["CONCENTRÉS", "14 963", "8 207", "17 319", "11 932", "19 658", "13 094"],
    ["ALIMENT_COMPLET", "709", "296", "500", "502", "561", "393"],
    ["INGRÉDIENTS", "964", "606", "733", "531", "686", "151"],
    ["PREMIX", "75", "41", "86", "60", "0 (CA)", "157"],
    ["COMPLEMENT_ALIM.", "8", "3", "6", "3", "3", "18"],
    ["MAÏS (exclu)", "0", "0", "0", "0", "0", "0"],
    ["TOTAL", "50 041", "28 450", "63 228", "52 073", "96 102", "53 052"],
]
story.append(make_table(hist_data, col_widths=[2.8*cm, 1.8*cm, 1.8*cm, 1.8*cm, 2.2*cm, 2.2*cm, 2.4*cm], font_size=8, highlight_rows=[6, 7]))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph(
    "<b>Lecture</b> : Le TOURTEAUX montre une trajectoire haussière (33 323 t en 2023 → 60 673 t forecast 2027, +82%). "
    "Les CONCENTRÉS progressent également (+31% vs 2026 annualisé). COMPLEMENT_ALIMENTAIRE reste stable autour de "
    "3-7 t/an (V300 1L proxy). Progression globale 2027 vs 2026 annualisé : <b>+4,4%</b> en volume.",
    BODY))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Recommandations principales</b>", H3))
recos = [
    "<b>Planification annuelle 2027</b> — Utiliser le forecast S3 (96 102 t, 53 052 M FCFA) comme base budgétaire.",
    "<b>Saisonnalité Q4</b> — Le Q4 2027 représente 28% du CA annuel (12 245 M FCFA). Préparer stocks dès septembre 2027.",
    "<b>Pic d'octobre</b> — Octobre 2027 = pic de l'année. Anticiper le réapprovisionnement soja avant septembre 2027.",
    "<b>Maintien du bundle</b> — Ratio bundle 2,3:1 atteint en août 2026 à maintenir en 2027.",
    "<b>Surveillance COMPLEMENT_ALIMENTAIRE</b> — Famille ajoutée au forecast (V300 1L proxy). 18 M FCFA de CA prévu.",
    "<b>Surveillance FAMLA et NDOBO</b> — Ces 2 agences représentent 43% du CA 2027.",
    "<b>Mise à jour trimestrielle</b> — Actualiser le forecast chaque trimestre avec données ERP.",
]
for r in recos:
    story.append(Paragraph(f"• {r}", BULLET))

story.append(Spacer(1, 0.3*cm))
story.append(Paragraph("<b>Conclusion</b>", H3))
story.append(Paragraph(
    "Le forecast 2027 projette <b>96 102 tonnes</b> pour un CA de <b>53 052 M FCFA</b>. La désaisonnalisation "
    "de l'effet soja, l'actualisation du prix à 25 000 FCFA, l'intégration de l'année 2023 (44 mois d'historique), "
    "l'ajout de la famille COMPLEMENT_ALIMENTAIRE (V300 1L only) et le maintien de MATERIEL_ELEVAGE à 0 tonne "
    "permettent une projection réaliste. Le pic d'octobre nécessitera une anticipation renforcée du "
    "réapprovisionnement soja dès septembre 2027.",
    BODY))

doc.build(story)
print(f"✓ 2027 Résumé exécutif: {os.path.getsize(exec_path)/1024:.0f} KB")

# === 2027: 2-5 PDFs (proposition, RACI, stratégique, guide) ===
# Create concise versions reusing cover_page and table structures

# 2. PROPOSITION
print("Generating 2027: 2. Proposition (v3)...")
prop_path = f"{F2027_DIR}/02_proposition_projet.pdf"
doc = SimpleDocTemplate(prop_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []
story.extend(cover_page("Proposition de Projet", "Forecast 2027 - Version 3", "PROPOSITION DE PROJET"))

story.append(Paragraph("1. Contexte", H1))
story.append(Paragraph(
    "BELGOCAM SA doit anticiper ses volumes et CA pour l'année 2027 complète. Cette version 3 s'appuie sur "
    "44 mois d'historique (Jan 2023 - Août 2026) et intègre : désaisonnalisation de l'effet soja exceptionnel, "
    "inclusion des commandes En cours/Validées, actualisation du prix soja à 25 000 FCFA/sac, "
    "ajout de la famille COMPLEMENT_ALIMENTAIRE (V300 1L only, V305 200L exclu), intégration de l'année 2023, "
    "MATERIEL_ELEVAGE à 0 tonne. Le forecast couvre 8 familles, 121 produits et 25 agences.",
    BODY))

story.append(Paragraph("2. Objectifs", H1))
story.append(Paragraph("<b>Objectif principal</b>", H3))
story.append(Paragraph(
    "Produire un forecast annuel 2027 (12 mois) en volumes et valeurs, désagrégé par produit, famille, agence "
    "et région, selon le scénario S3.",
    BODY))
story.append(Paragraph("<b>Objectifs spécifiques</b>", H3))
objs = [
    "Modéliser tendance + saisonnalité mensuelles via Prophet (5 familles)",
    "Désaisonnaliser l'effet soja exceptionnel de juillet-août 2026",
    "Intégrer les commandes En cours et Validées",
    "Actualiser le prix soja à 25 000 FCFA/sac",
    "NOUVEAU v3: Ajouter COMPLEMENT_ALIMENTAIRE (V300 1L only, 1L=1kg)",
    "NOUVEAU v3: Intégrer l'année 2023 (44 mois d'historique)",
    "NOUVEAU v3: MATERIEL_ELEVAGE toujours à 0 tonne",
    "Exclure le Maïs, V305 (BELGOKILL 200L) et produits opportunistes",
    "Produire les livrables PACE complets",
]
for o in objs:
    story.append(Paragraph(f"• {o}", BULLET))

story.append(Paragraph("3. Périmètre", H1))
scope_data = [
    ["Dimension", "Détail", "Volume"],
    ["Période forecast", "Janvier - Décembre 2027 (12 mois)", "12 mois"],
    ["Familles incluses", "8 familles (TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, ALIMENT COMPLET, MATERIEL ÉLEVAGE, PREMIX, COMPLEMENT_ALIMENTAIRE)", "8"],
    ["Produits", "115 références (T102, C101-C108, MAT014, P102N2, V300 BELGOKILL 1L, etc.)", "121"],
    ["Agences", "25 agences BELGOCAM", "20"],
    ["Régions", "Ouest, Centre, Littoral", "3"],
    ["Niveau détail", "Produit × Agence × Mois", "13 704 lignes"],
    ["Scénario", "S3 - Réappro soja 100%", "1"],
    ["Données historiques", "176 576 enregistrements (Jan 2023 - Août 2026)", "44 mois"],
    ["Désaisonnalisation", "Effet soja Jul-Août 2026 neutralisé", "—"],
    ["Prix soja", "25 000 FCFA/sac (actualisé 24/08/2026)", "—"],
    ["NOUVEAU v3: COMPLEMENT_ALIM.", "V300 1L only (V305 200L exclu), 1L=1kg", "9 produits"],
    ["NOUVEAU v3: ALVEOLES séparé", "4 refs (MAT011/MAT014/MAT015/MAT017), tonnes=0", "CA only"],
]
story.append(make_table(scope_data, col_widths=[3.5*cm, 9.5*cm, 4*cm], font_size=9, highlight_rows=[11, 12]))

story.append(Paragraph("4. Méthodologie PACE", H1))
pace_data = [
    ["Phase", "Activités", "Livrables"],
    ["P - PREPARE", "Consolidation 176 576 records + En cours/Validées, désaisonnalisation soja, ajout COMPLEMENT_ALIM.", "Dataset 2023-2026, prix 2027"],
    ["A - ANALYZE", "AED, saisonnalité 2023-2025, top produits/agences", "Graphiques, synthèse AED"],
    ["C - CONSTRUCT", "Prophet (5 familles × 3 régions) + extrapolation (2 familles), forecast 12 mois S3", "Forecast 2027 (13 704 lignes)"],
    ["E - EXECUTE", "Excel 8 feuilles, 5 PDFs PACE", "8 livrables finaux"],
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
    ["Historique 44 mois — suffisant", "Faible", "Mise à jour trimestrielle"],
    ["Prix soja volatil (+47% en 2 mois)", "Élevée", "Prix actualisé 25 000 FCFA, scénario S3"],
    ["Effet soja 2026 non récurrent en 2027", "Élevée", "Cap désaisonnalisation à moyenne S1"],
    ["V305 exclu — BELGOKILL 200L absent", "Faible", "V300 1L utilisé comme proxy suffisant"],
    ["Nouvelles hausses tarifaires en 2027", "Moyenne", "Hypothèse prix stable en 2027"],
]
story.append(make_table(risks_data, col_widths=[7*cm, 3*cm, 7*cm], font_size=9))

story.append(Paragraph("7. Critères de succès", H1))
success = [
    "Forecast 2027 produit sur 12 mois avec désagrégation complète",
    "Désaisonnalisation effective de l'effet soja",
    "Prix soja actualisé à 25 000 FCFA/sac",
    "Famille COMPLEMENT_ALIMENTAIRE ajoutée (V300 1L only)",
    "Écart forecast vs réalité ≤ 15% par trimestre",
    "Livrables PACE complets (8 documents) produits et diffusés",
]
for s in success:
    story.append(Paragraph(f"• {s}", BULLET))

doc.build(story)
print(f"✓ 2027 Proposition: {os.path.getsize(prop_path)/1024:.0f} KB")

# 3. MATRICE RACI 2027
print("Generating 2027: 3. Matrice RACI (v3)...")
raci_path = f"{F2027_DIR}/03_matrice_raci.pdf"
doc = SimpleDocTemplate(raci_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
story = []
story.extend(cover_page("Matrice RACI", "Forecast 2027 - Version 3", "MATRICE RACI"))

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
    ["PREPARE", "1.1 Consolidation 2023-2026 + En cours/Validées", "R/A", "I", "C", "I", "I", "I"],
    ["", "1.2 Désaisonnalisation effet soja", "R/A", "C", "C", "I", "I", "I"],
    ["", "1.3 Calcul prix 2027 (soja 25 000)", "R/A", "C", "I", "I", "C", "I"],
    ["", "1.4 NOUVEAU: Ajout COMPLEMENT_ALIMENTAIRE (V300 1L)", "R/A", "C", "I", "I", "I", "I"],
    ["", "1.5 Validation dataset v3", "R", "C", "C", "I", "A", "I"],
    ["ANALYZE", "2.1 AED + saisonnalité 2023-2025", "R/A", "I", "I", "I", "I", "I"],
    ["", "2.2 Analyse effet soja exceptionnel", "R/A", "C", "C", "I", "I", "I"],
    ["", "2.3 Validation hypothèses AED", "R", "C", "C", "I", "A", "I"],
    ["CONSTRUCT", "3.1 Prophet (5 familles × 3 régions)", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.2 Extrapolation MAT+PREMIX (CA only)", "R/A", "I", "I", "I", "I", "I"],
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
    ["Validation dataset v3", "PREPARE", "03/09/2026", "CG", "Dataset 2023-2026 + COMPLEMENT_ALIM."],
    ["Validation AED", "ANALYZE", "04/09/2026", "DC", "Synthèse AED"],
    ["Validation forecast", "CONSTRUCT", "05/09/2026", "DG + DC", "Forecast 2027 complet"],
    ["Livrables finaux", "EXECUTE", "05/09/2026", "DG", "8 livrables"],
    ["Présentation CODIR", "EXECUTE", "06/09/2026", "DG", "Validation officielle"],
    ["Mise à jour Q1", "—", "Mars 2027", "DA", "Réactualisation"],
]
story.append(make_table(control_data, col_widths=[3.5*cm, 2.5*cm, 3*cm, 3*cm, 5*cm], font_size=8))

doc.build(story)
print(f"✓ 2027 Matrice RACI: {os.path.getsize(raci_path)/1024:.0f} KB")

# 4. DOCUMENT STRATÉGIQUE PACE 2027
print("Generating 2027: 4. Document stratégique PACE (v3)...")
strat_path = f"{F2027_DIR}/04_document_strategique_pace.pdf"
doc = SimpleDocTemplate(strat_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []
story.extend(cover_page("Document Stratégique PACE", "Forecast 2027 - Version 3 (8 familles, données 2023-2026)", "STRATÉGIE PACE"))

story.append(Paragraph("1. Vision et objectifs", H1))
story.append(Paragraph("1.1 Vision", H2))
story.append(Paragraph(
    "Faire du forecast 2027 un <b>outil de planification annuelle</b> permettant à BELGOCAM SA d'anticiper "
    "96 102 tonnes de ventes et 53 052 M FCFA de chiffre d'affaires, avec désaisonnalisation de l'effet soja "
    "et intégration de la famille COMPLEMENT_ALIMENTAIRE (V300 1L only).",
    BODY))

story.append(Paragraph("1.2 Objectifs stratégiques", H2))
obj_data = [
    ["Objectif", "Description", "KPI"],
    ["Forecast annuel", "12 mois 2027 en volume + valeur", "96 102 t, 53 052 M FCFA"],
    ["Désaisonnalisation", "Neutraliser l'effet soja Jul-Août 2026", "Cap moyenne S1 2026"],
    ["En cours + Validées", "Intégrer comme potentielles ventes", "255 commandes incluses"],
    ["Prix actualisé", "Soja 25 000 FCFA/sac", "vs 20 600 précédent"],
    ["Désagrégation", "Produit × agence × mois", "13 704 lignes"],
    ["Adoption", "Diffusion CODIR + 25 agences", "100% agences informées"],
    ["NOUVEAU v3: COMPLEMENT_ALIM.", "V300 1L only (V305 exclu), 1L=1kg", "3 t, 18 M FCFA"],
    ["NOUVEAU v3: ALVEOLES séparé", "4 refs (MAT011/MAT014/MAT015/MAT017), tonnes=0", "0 t, 438 M FCFA"],
]
story.append(make_table(obj_data, col_widths=[3.5*cm, 8.5*cm, 5*cm], font_size=9, highlight_rows=[6, 7]))

story.append(Paragraph("2. Phase PREPARE - Préparation des données", H1))
story.append(Paragraph("2.1 Sources consolidées (2023-2026)", H2))
sources_data = [
    ["Source", "Période", "Records"],
    ["LY_21_24 (filtre 2023)", "Jan-Dec 2023", "6 016"],
    ["LY_24 (Jul-Dec 2024)", "Juillet-Décembre 2024", "52 215"],
    ["Historique 2025", "Jan-Déc 2025", "66 206"],
    ["S1 + Juil + Août 2026", "Jan-Août 2026", "52 139"],
    ["En cours + Validées", "Août 2026", "~255"],
    ["TOTAL", "44 mois", "176 576"],
]
story.append(make_table(sources_data, col_widths=[5*cm, 5*cm, 5*cm], font_size=9))

story.append(Paragraph("2.2 Désaisonnalisation de l'effet soja", H2))
story.append(Paragraph(
    "Les volumes soja de juillet-août 2026 ont été exceptionnellement élevés (rupture concurrente). "
    "Pour éviter que Prophet n'extrapole ce phénomène non récurrent, les volumes soja Jul-Août 2026 ont été "
    "<b>cappés à la moyenne S1 2026</b>.",
    BODY))

story.append(Paragraph("2.3 NOUVEAU v3: Famille COMPLEMENT_ALIMENTAIRE", H2))
story.append(Paragraph(
    "La famille COMPLEMENT_ALIMENTAIRE (9 produits liquides 1L) a été ajoutée au forecast 2027. "
    "Selon demande utilisateur : <b>seul V300 (BELGOKILL 1L) est conservé — V305 (BELGOKILL 200L) est EXCLU</b>. "
    "Conversion 1L = 1kg appliquée pour exprimer les volumes en tonnes.",
    BODY))

story.append(Paragraph("3. Phase CONSTRUCT - Modélisation", H1))
story.append(Paragraph("3.1 Configuration Prophet", H2))
config_data = [
    ["Paramètre", "Valeur", "Justification"],
    ["yearly_seasonality", "True", "Capture la saisonnalité annuelle"],
    ["seasonality_mode", "multiplicative", "Saisonnalité proportionnelle"],
    ["changepoint_prior_scale", "0.05", "Tendance modérée"],
    ["interval_width", "0.8", "Intervalle 80%"],
    ["Période forecast", "12 mois (Jan-Déc 2027)", "Année complète"],
    ["Familles Prophet", "5 (incluant COMPLEMENT_ALIMENTAIRE)", "Nouveau v3"],

]
story.append(make_table(config_data, col_widths=[5*cm, 5*cm, 7*cm], font_size=9, highlight_rows=[6]))

story.append(Paragraph("4. Résultats par famille et trimestre", H1))
story.append(Paragraph("4.1 Par famille", H2))
fam_detail = [
    ["Famille", "Volume (t)", "CA (M FCFA)", "Part CA", "Méthode"],
    ["TOURTEAUX", "66 338", "33 176", "62,5%", "Prophet"],
    ["CONCENTRÉS", "27 822", "18 493", "34,9%", "Prophet + bundle 2.5:1"],
    ["ALIMENT_COMPLET", "877", "612", "1,2%", "Prophet"],
    ["ALVEOLES", "0", "101", "0,2%", "Extrap 2026 + Forfait SPC (ALV=0)"],
    ["MATERIEL_ELEVAGE", "0", "214", "0,4%", "Extrap + Forfait SPC (2026 annualisé)"],
    ["INGREDIENTS", "1 059", "236", "0,4%", "Prophet"],
    ["PREMIX", "0", "192", "0,4%", "Extrapolation CA"],
    ["COMPLEMENT_ALIM.", "5", "28", "0,1%", "Prophet (V300 1L proxy)"],
]
story.append(make_table(fam_detail, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2*cm, 4*cm], font_size=9, highlight_rows=[4, 8]))

story.append(Paragraph("4.2 Par trimestre", H2))
q_detail = [
    ["Trimestre", "Volume (t)", "CA (M FCFA)", "Part CA", "Lecture"],
    ["Q1 (Jan-Mar)", "25 295", "13 949", "26,3%", "Démarrage"],
    ["Q2 (Avr-Juin)", "24 316", "13 414", "25,3%", "Montée"],
    ["Q3 (Juil-Sept)", "19 515", "10 849", "20,4%", "Creux"],
    ["Q4 (Oct-Déc)", "26 975", "14 839", "28,0%", "Pic saisonnier"],
]
story.append(make_table(q_detail, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2*cm, 4.5*cm], font_size=9))

story.append(Paragraph("5. Historique 2023-2026 vs Forecast 2027", H1))
hist_detail = [
    ["Famille", "2023 (t)", "2024 (t)", "2025 (t)", "2026 YTD (t)", "2027 fcst (t)", "CA 2027 (M)"],
    ["TOURTEAUX", "33 323", "19 296", "44 584", "39 045", "60 673", "30 338"],
    ["CONCENTRÉS", "14 963", "8 207", "17 319", "11 932", "19 658", "13 094"],
    ["ALIMENT_COMPLET", "709", "296", "500", "502", "561", "393"],
    ["INGRÉDIENTS", "964", "606", "733", "531", "686", "151"],
    ["PREMIX", "75", "41", "86", "60", "0 (CA)", "157"],
    ["COMPLEMENT_ALIM.", "8", "3", "6", "3", "3", "18"],
    ["TOTAL", "50 041", "28 450", "63 228", "52 073", "96 102", "53 052"],
]
story.append(make_table(hist_detail, col_widths=[2.8*cm, 1.8*cm, 1.8*cm, 1.8*cm, 2.2*cm, 2.2*cm, 2.4*cm], font_size=8, highlight_rows=[6, 7]))

story.append(Paragraph("6. Plan de déploiement", H1))
deploy_data = [
    ["Étape", "Date", "Action", "Cible"],
    ["1", "03/09/2026", "Validation proposition v3", "DG"],
    ["2", "05/09/2026", "Livrables finaux prêts", "DA"],
    ["3", "06/09/2026", "Présentation CODIR", "DG + DC + DP + CG"],
    ["4", "07/09/2026", "Diffusion aux agences", "14 RA"],
    ["5", "Mars 2027", "Mise à jour Q1", "DA"],
]
story.append(make_table(deploy_data, col_widths=[1*cm, 3*cm, 7*cm, 6*cm], font_size=9))

story.append(Paragraph("7. Conclusion", H1))
story.append(Paragraph(
    "Le forecast 2027 version 3 projette <b>96 102 tonnes</b> pour <b>53 052 M FCFA</b>. La désaisonnalisation "
    "de l'effet soja, l'actualisation du prix à 25 000 FCFA, l'intégration de l'année 2023 (44 mois d'historique), "
    "l'ajout de la famille COMPLEMENT_ALIMENTAIRE (V300 1L only) et le maintien de MATERIEL_ELEVAGE à 0 tonne "
    "permettent une projection réaliste.",
    BODY))

doc.build(story)
print(f"✓ 2027 Document stratégique: {os.path.getsize(strat_path)/1024:.0f} KB")

# 5. GUIDE MÉTHODOLOGIQUE 2027
print("Generating 2027: 5. Guide méthodologique (v3)...")
guide_path = f"{F2027_DIR}/05_guide_methodologique.pdf"
doc = SimpleDocTemplate(guide_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []
story.extend(cover_page("Guide Méthodologique", "Forecast 2027 - Version 3 (Prophet + COMPLEMENT_ALIMENTAIRE V300 1L)", "GUIDE MÉTHODOLOGIQUE"))

story.append(Paragraph("Sommaire", H1))
story.append(Paragraph(
    "1. Introduction et contexte<br/>"
    "2. Méthodologie PACE<br/>"
    "3. Outils et environnement<br/>"
    "4. Phase PREPARE - Données 2023-2026<br/>"
    "5. Phase PREPARE - Désaisonnalisation soja<br/>"
    "6. NOUVEAU v3: Phase PREPARE - COMPLEMENT_ALIMENTAIRE (V300 1L)<br/>"
    "7. Phase ANALYZE - AED<br/>"
    "8. Phase CONSTRUCT - Prophet<br/>"
    "9. Phase EXECUTE - Livrables<br/>"
    "Annexes",
    BODY))

story.append(PageBreak())
story.append(Paragraph("1. Introduction et contexte", H1))
story.append(Paragraph(
    "Ce guide décrit la méthodologie complète du forecast 2027 de BELGOCAM SA (version 3). Il couvre 12 mois "
    "(janvier-décembre 2027) avec 8 familles de produits (incluant COMPLEMENT_ALIMENTAIRE) et 25 agences. "
    "Innovations v3 : (1) ajout COMPLEMENT_ALIMENTAIRE (V300 1L only, V305 200L EXCLU, 1L=1kg), "
    "(2) intégration année 2023 (44 mois d'historique), (3) MATERIEL_ELEVAGE toujours 0 tonne, "
    "(4) désaisonnalisation effet soja, (5) inclusion En cours/Validées, (6) prix soja 25 000 FCFA.",
    BODY))

story.append(Paragraph("2. Méthodologie PACE", H1))
pace_detail = [
    ["Phase", "Objectif", "Livrables", "Durée"],
    ["P - PREPARE", "Préparer données 2023-2026 + désaisonnalisation + COMPLEMENT_ALIM.", "Dataset 176 576 records, prix 2027", "2 jours"],
    ["A - ANALYZE", "Comprendre données + effet soja", "Graphiques, synthèse AED", "1 jour"],
    ["C - CONSTRUCT", "Modéliser Prophet 12 mois (8 familles)", "Forecast 2027 (13 704 lignes)", "2 jours"],
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
]
story.append(make_table(stack_data, col_widths=[4*cm, 3*cm, 8*cm], font_size=9))

story.append(PageBreak())
story.append(Paragraph("4. Phase PREPARE - Données 2023-2026", H1))
story.append(Paragraph(
    "Le dataset consolidé couvre 44 mois (janvier 2023 - août 2026), soit 176 576 enregistrements. "
    "Sources : LY_21_24 filtre 2023 (6 016), LY_24 Jul-Dec 2024 (52 215), Historique 2025 (66 206), "
    "S1+Juil+Août 2026 (52 139), En cours + Validées (~255).",
    BODY))
story.append(Paragraph(
    "<b>Intégration de l'année 2023</b> : La version 3 inclut 2023 pour enrichir l'historique avec 4 ans "
    "de données. Prophet dispose ainsi d'une saisonnalité annuelle complète.",
    BODY))

story.append(Paragraph("5. Phase PREPARE - Désaisonnalisation soja", H1))
story.append(Paragraph(
    "L'effet soja exceptionnel de juillet-août 2026 (rupture concurrente) a généré des volumes anormalement "
    "élevés non récurrents. Les volumes soja Jul-Août 2026 ont été cappés à la moyenne S1 2026.",
    BODY))

story.append(PageBreak())
story.append(Paragraph("6. NOUVEAU v3: Famille COMPLEMENT_ALIMENTAIRE", H1))
story.append(Paragraph(
    "La famille COMPLEMENT_ALIMENTAIRE regroupe 9 produits liquides BELGO en conditionnement 1L : "
    "BELGOKILL V300 (1L), BELGO HARMONY, BELGO PROTECT, BELGO DRY LIT, BELGO WATER CLEAN, BELGO VIT Ese, "
    "BELGO THERMO, BELGO BIO SELECT, BELGO FRESH.",
    BODY))

story.append(Paragraph("<b>Exclusion de V305 (BELGOKILL 200L)</b>", H3))
story.append(Paragraph(
    "Selon demande utilisateur explicite : seul le conditionnement 1L est conservé. "
    "V305 (BELGOKILL 200L) est EXCLU du forecast. Cela permet d'avoir une famille homogène (tous 1L) "
    "et d'éviter le biais du conditionnement bulk.",
    BODY))

story.append(Paragraph("<b>Conversion 1L = 1kg</b>", H3))
story.append(Paragraph(
    "Pour exprimer les volumes en tonnes (nécessaire pour Prophet), la conversion 1L = 1kg est appliquée : "
    "<b>1 qte = 1 kg, tonnes = qte / 1000</b>.",
    BODY))

story.append(Paragraph("<b>Proxy BELGOKILL V300 1L</b>", H3))
story.append(Paragraph(
    "BELGOKILL V300 (1L) est le produit dominant de la famille (38% du CA famille sur 2023-2026). "
    "Sa tendance est utilisée comme proxy pour modéliser toute la famille via Prophet.",
    BODY))

ca_proxy = [
    ["Réf.", "Produit", "Poids (kg/qte)", "Prix 2027 (FCFA/L)", "Statut"],
    ["V300", "BELGOKILL 1L", "1", "2 500", "✓ Inclus (proxy)"],
    ["V305", "BELGOKILL 200L", "—", "—", "✗ EXCLU v3"],
    ["CA003.1", "BELGO HARMONY 1L", "1", "8 000", "✓ Inclus"],
    ["CA004.1", "BELGO PROTECT 1L", "1", "9 800", "✓ Inclus"],
    ["CA006.1", "BELGO DRY LIT 1L", "1", "6 500", "✓ Inclus"],
    ["CA001.1", "BELGO WATER CLEAN 1L", "1", "5 500", "✓ Inclus"],
    ["CA002.1", "BELGO VIT Ese 1L", "1", "10 500", "✓ Inclus"],
    ["CA005.1", "BELGO THERMO 1L", "1", "15 000", "✓ Inclus"],
    ["CA007.1", "BELGO BIO SELECT 1L", "1", "8 000", "✓ Inclus"],
    ["CA008.1", "BELGO FRESH 1L", "1", "14 000", "✓ Inclus"],
]
story.append(make_table(ca_proxy, col_widths=[2*cm, 5*cm, 2.5*cm, 3.5*cm, 3*cm], font_size=9, highlight_rows=[2]))

story.append(Paragraph("7. Phase ANALYZE - AED", H1))
story.append(Paragraph(
    "L'AED révèle : (1) TOURTEAUX = 76% du volume historique, (2) Q4 = 35-46% du volume annuel, "
    "(3) pic octobre (indice 1.86 pour TOURTEAUX en 2025), (4) FAMLA = 31.6% du volume, "
    "(5) COMPLEMENT_ALIMENTAIRE = 3-7 t/an (faible volume, forte valeur unitaire).",
    BODY))

story.append(Paragraph("8. Phase CONSTRUCT - Prophet", H1))
story.append(Paragraph(
    "15 modèles Prophet famille × région (5 familles Prophet × 3 régions) + extrapolation pour MATERIEL_ELEVAGE "
    "et PREMIX (CA only, tonnes=0). Désagrégation par produit × agence selon parts historiques de CA. "
    "Prix 2027 : 25 000 FCFA/sac soja, prix par litre pour COMPLEMENT_ALIMENTAIRE (2 500-15 000 FCFA/L).",
    BODY))

story.append(Paragraph("9. Phase EXECUTE - Livrables", H1))
livrables_data = [
    ["#", "Livrable", "Format", "Usage"],
    ["1", "Résumé exécutif", "PDF", "Direction Générale"],
    ["2", "Proposition de projet", "PDF", "CODIR"],
    ["3", "Matrice RACI", "PDF", "Gouvernance"],
    ["4", "Document stratégique PACE", "PDF", "Référence stratégique"],
    ["5", "Guide méthodologique", "PDF", "Référence méthodes"],
    ["6", "Excel forecast 2027 (8 feuilles)", "XLSX", "Pilotage opérationnel"],
]
story.append(make_table(livrables_data, col_widths=[1*cm, 5*cm, 3*cm, 6*cm], font_size=9))

story.append(Paragraph("Annexe - Glossaire", H1))
glossaire = [
    ["Terme", "Définition"],
    ["Prophet", "Bibliothèque prévision séries temporelles (Facebook/Meta)"],
    ["Désaisonnalisation", "Neutralisation d'un effet exceptionnel non récurrent"],
    ["S3", "Scénario réappro soja 100%"],
    ["Cap S1", "Plafonnement Jul-Août 2026 à moyenne S1 2026"],
    ["Q4", "Quatrième trimestre (oct-déc)"],
    ["PACE", "Prepare-Analyze-Construct-Execute"],
    ["NOUVEAU v3: COMPLEMENT_ALIM.", "9 produits liquides 1L BELGOxxx (V300 1L proxy)"],
    ["NOUVEAU v3: ALVEOLES séparé", "Famille distincte de MATERIEL_ELEVAGE (4 refs MAT011/14/15/17)"],
    ["NOUVEAU v4: SPC agences", "5 agences SPC incluses (Baf-Chefferie dominant, 274 M CA historique)"],
    ["NOUVEAU v4: Maroua", "Agence ajoutée (Centre par convention), soja T102 only (Août 2026)"],
    ["NOUVEAU v4: Bundle 2.5:1", "Ratio soja:concentré forcé à <=2.5:1 (augment. CONCENTRES si ratio > 2.5)"],
    ["NOUVEAU v4: Saisonnalité ALV 2026", "ALVEOLES utilise 2026 (13.9 M/an) au lieu du pic 2025 (739 M)"],
    ["NOUVEAU v5: 10 SPC agences", "Toutes SPC incluses (vs 5 avant) — fin exclusion COMPTOIR pour SPC"],
    ["NOUVEAU v5: Forfait SPC v2", "Forfait MAT_ELEV (38.4 M/an) basé sur 2026 annualisé, poids 2025"],
    ["NOUVEAU v5: SPC PK15 forfait", "Poids 0 en 2025 → forfait réaliste 1 M/an (0.5 ALV + 0.5 MAT)"],
    ["NOUVEAU v3: V305 EXCLU", "BELGOKILL 200L retiré (seul V300 1L conservé)"],
    ["NOUVEAU v3: 1L=1kg", "Conversion pour volumes liquides en tonnes"],
    ["NOUVEAU v3: Données 2023-2026", "44 mois d'historique (vs 20 mois en v1)"],
    ["MATERIEL_ELEVAGE", "Toujours 0 en tonnes (CA only). ALVEOLES désormais séparés en famille distincte"],
]
story.append(make_table(glossaire, col_widths=[5*cm, 11*cm], font_size=9, highlight_rows=[7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]))

doc.build(story)
print(f"✓ 2027 Guide méthodologique: {os.path.getsize(guide_path)/1024:.0f} KB")

print(f"\n=== 2027: 5 PDFs v3 GÉNÉRÉS ===")
for f in sorted(os.listdir(F2027_DIR)):
    if f.endswith('.pdf'):
        print(f"  {f}: {os.path.getsize(os.path.join(F2027_DIR, f))/1024:.0f} KB")

print("\n=== TOUS LES 10 PDFs v3 GÉNÉRÉS ===")
print(f"Q4 2026: {Q4_DIR}")
print(f"2027: {F2027_DIR}")
