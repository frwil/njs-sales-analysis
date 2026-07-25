"""
BELGOCAM SA — Analyse Bundle Soja-Concentrés (Standalone PDF)
Impact de la hausse du prix du soja sur le comportement d'achat client

PDF généré via ReportLab.
Réutilise les mêmes polices, couleurs et fonctions utilitaires que build_zero_achat_pdf.py
(même package graphique corporate BELGOCAM navy + gold).

Auteur/rédacteur : William Francis Fohom, Data Analyst, Administrateur National de Ventes
Date : Juillet 2026
"""
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.platypus import (
    Paragraph, Spacer, PageBreak, Table, TableStyle,
    Image, KeepTogether, NextPageTemplate, PageTemplate, Frame, BaseDocTemplate
)
from reportlab.platypus.flowables import HRFlowable

# ===== FONTS (identiques à build_zero_achat_pdf.py) =====
FONT_DIR = '/usr/share/fonts'
pdfmetrics.registerFont(TTFont('NotoSerifSC', f'{FONT_DIR}/truetype/noto-serif-sc/NotoSerifSC-Regular.ttf'))
pdfmetrics.registerFont(TTFont('NotoSerifSC-Bold', f'{FONT_DIR}/truetype/noto-serif-sc/NotoSerifSC-Bold.ttf'))
pdfmetrics.registerFont(TTFont('NotoSerifSC-Light', f'{FONT_DIR}/truetype/noto-serif-sc/NotoSerifSC-Light.ttf'))
registerFontFamily('NotoSerifSC', normal='NotoSerifSC', bold='NotoSerifSC-Bold')

# ===== COLORS (BELGOCAM corporate: navy + gold) =====
NAVY = colors.HexColor('#1F3A5F')
NAVY_LIGHT = colors.HexColor('#2E75B6')
GOLD = colors.HexColor('#C9A961')
GOLD_DARK = colors.HexColor('#BF8F00')
RED = colors.HexColor('#C00000')
GREEN = colors.HexColor('#548235')
GRAY = colors.HexColor('#5A5A5A')
GRAY_LIGHT = colors.HexColor('#D9D9D9')
GRAY_VLIGHT = colors.HexColor('#F2F2F2')
WHITE = colors.white
BLACK = colors.black

# ===== STYLES =====
styles = getSampleStyleSheet()

H1 = ParagraphStyle('H1', parent=styles['Heading1'], fontName='NotoSerifSC-Bold', fontSize=20,
                    textColor=NAVY, spaceAfter=14, spaceBefore=10, alignment=TA_LEFT, leading=26)
H2 = ParagraphStyle('H2', parent=styles['Heading2'], fontName='NotoSerifSC-Bold', fontSize=14,
                    textColor=NAVY, spaceAfter=10, spaceBefore=14, alignment=TA_LEFT, leading=18,
                    keepWithNext=1)
H3 = ParagraphStyle('H3', parent=styles['Heading3'], fontName='NotoSerifSC-Bold', fontSize=12,
                    textColor=NAVY_LIGHT, spaceAfter=8, spaceBefore=10, alignment=TA_LEFT, leading=15,
                    keepWithNext=1)

BODY = ParagraphStyle('Body', parent=styles['BodyText'], fontName='NotoSerifSC', fontSize=10,
                      textColor=BLACK, alignment=TA_JUSTIFY, leading=14, spaceAfter=8)
BODY_LEFT = ParagraphStyle('BodyLeft', parent=BODY, alignment=TA_LEFT)
BODY_BOLD = ParagraphStyle('BodyBold', parent=BODY, fontName='NotoSerifSC-Bold')
BODY_ITALIC = ParagraphStyle('BodyItalic', parent=BODY, fontName='NotoSerifSC-Light')
SMALL = ParagraphStyle('Small', parent=BODY, fontSize=9, leading=12, textColor=GRAY)
CAPTION = ParagraphStyle('Caption', parent=SMALL, alignment=TA_CENTER, fontName='NotoSerifSC-Light')

CELL = ParagraphStyle('Cell', fontName='NotoSerifSC', fontSize=9, leading=12, textColor=BLACK, alignment=TA_LEFT)
CELL_CENTER = ParagraphStyle('CellC', parent=CELL, alignment=TA_CENTER)
CELL_RIGHT = ParagraphStyle('CellR', parent=CELL, alignment=TA_RIGHT)
CELL_BOLD = ParagraphStyle('CellB', parent=CELL, fontName='NotoSerifSC-Bold')
CELL_WHITE = ParagraphStyle('CellW', parent=CELL, textColor=WHITE, fontName='NotoSerifSC-Bold', alignment=TA_CENTER)
CELL_WHITE_R = ParagraphStyle('CellWR', parent=CELL_WHITE, alignment=TA_RIGHT)

BULLET = ParagraphStyle('Bullet', parent=BODY, leftIndent=18, bulletIndent=6, spaceAfter=4)

# ===== PAGE SETUP =====
PAGE_W, PAGE_H = A4
MARGIN_L = 1.5*cm
MARGIN_R = 1.5*cm
MARGIN_T = 2.2*cm
MARGIN_B = 1.8*cm
CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R


# ===== COVER PAGE FUNCTION (custom for Analyse Bundle Soja-Concentrés) =====
def draw_cover(canv, doc):
    """Draw cover page elements directly on canvas — Bundle Soja-Concentrés edition."""
    canv.saveState()

    # Background: full white
    canv.setFillColor(WHITE)
    canv.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)

    # Top navy band (full width)
    canv.setFillColor(NAVY)
    canv.rect(0, PAGE_H - 5*cm, PAGE_W, 5*cm, fill=1, stroke=0)

    # Gold accent line
    canv.setFillColor(GOLD)
    canv.rect(0, PAGE_H - 5.3*cm, PAGE_W, 0.3*cm, fill=1, stroke=0)

    # Company name (top left, in navy band)
    canv.setFillColor(WHITE)
    canv.setFont('NotoSerifSC-Bold', 22)
    canv.drawString(MARGIN_L, PAGE_H - 2.5*cm, "BELGOCAM SA")
    canv.setFont('NotoSerifSC', 11)
    canv.drawString(MARGIN_L, PAGE_H - 3.2*cm, "NJS GROUP — Cameroun")
    canv.setFont('NotoSerifSC-Light', 10)
    canv.drawString(MARGIN_L, PAGE_H - 4.0*cm, "Analyse Commerciale & Stratégique")

    # Date (top right)
    canv.setFont('NotoSerifSC', 10)
    canv.drawRightString(PAGE_W - MARGIN_R, PAGE_H - 2.5*cm, "Juillet 2026")
    canv.setFont('NotoSerifSC-Light', 9)
    canv.drawRightString(PAGE_W - MARGIN_R, PAGE_H - 3.2*cm, "Période analysée : S1 2026 vs Juillet 2026 (au 24/07)")

    # Main title block (left aligned)
    canv.setFillColor(NAVY)
    canv.setFont('NotoSerifSC-Bold', 26)
    canv.drawString(MARGIN_L, PAGE_H - 8.4*cm, "Analyse Bundle")
    canv.drawString(MARGIN_L, PAGE_H - 9.6*cm, "Soja-Concentrés")

    # Gold separator
    canv.setFillColor(GOLD)
    canv.rect(MARGIN_L, PAGE_H - 10.4*cm, 4*cm, 0.15*cm, fill=1, stroke=0)

    # Subtitle
    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC', 12)
    canv.drawString(MARGIN_L, PAGE_H - 11.4*cm, "Impact de la hausse du prix du soja sur le comportement d'achat client")
    canv.setFont('NotoSerifSC-Light', 10)
    canv.drawString(MARGIN_L, PAGE_H - 12.2*cm, "Diagnostic par agence, région et focus NDOBO — Juillet 2026")

    # Key figures box
    box_y = 6.5*cm
    box_h = 5.5*cm
    canv.setFillColor(GRAY_VLIGHT)
    canv.rect(MARGIN_L, box_y, CONTENT_W, box_h, fill=1, stroke=0)
    # Left gold border
    canv.setFillColor(GOLD)
    canv.rect(MARGIN_L, box_y, 0.2*cm, box_h, fill=1, stroke=0)

    # 4 KPIs side by side
    canv.setFillColor(NAVY)
    canv.setFont('NotoSerifSC-Bold', 18)
    kpi_y = box_y + box_h - 1.3*cm
    col_w = CONTENT_W / 4
    canv.drawCentredString(MARGIN_L + col_w*0.5, kpi_y, "+11,8%")
    canv.drawCentredString(MARGIN_L + col_w*1.5, kpi_y, "+64,3%")
    canv.drawCentredString(MARGIN_L + col_w*2.5, kpi_y, "+28,9%")
    canv.drawCentredString(MARGIN_L + col_w*3.5, kpi_y, "2,5 → 3,3")

    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC', 8)
    canv.drawCentredString(MARGIN_L + col_w*0.5, kpi_y - 0.7*cm, "Hausse prix soja")
    canv.drawCentredString(MARGIN_L + col_w*1.5, kpi_y - 0.7*cm, "Budget client")
    canv.drawCentredString(MARGIN_L + col_w*2.5, kpi_y - 0.7*cm, "Sacs conc/client")
    canv.drawCentredString(MARGIN_L + col_w*3.5, kpi_y - 0.7*cm, "Ratio soja/conc")

    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC-Light', 7)
    canv.drawCentredString(MARGIN_L + col_w*0.5, kpi_y - 1.4*cm, "+1 000 FCFA/sac (territoire)")
    canv.drawCentredString(MARGIN_L + col_w*1.5, kpi_y - 1.4*cm, "bundle — moyenne agences")
    canv.drawCentredString(MARGIN_L + col_w*2.5, kpi_y - 1.4*cm, "croissance concentrés")
    canv.drawCentredString(MARGIN_L + col_w*3.5, kpi_y - 1.4*cm, "dégradation du ratio 3:1")

    # Bottom figure: key message
    canv.setFillColor(NAVY)
    canv.setFont('NotoSerifSC-Bold', 12)
    canv.drawCentredString(PAGE_W/2, kpi_y - 2.7*cm, "Hypothèse testée : « budget constant » des clients face à la hausse du soja")
    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC-Light', 8)
    canv.drawCentredString(PAGE_W/2, kpi_y - 3.3*cm, "14 agences analysées — 1 100 clients bundle S1 → 492 en Juillet — focus NDOBO (cas le plus problématique)")

    # Author / redactor block at bottom
    auth_y = 3.0*cm
    # Gold accent line above
    canv.setFillColor(GOLD)
    canv.rect(MARGIN_L, auth_y + 2.5*cm, 3*cm, 0.1*cm, fill=1, stroke=0)

    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC-Light', 9)
    canv.drawString(MARGIN_L, auth_y + 1.8*cm, "Rédigé par")

    canv.setFillColor(NAVY)
    canv.setFont('NotoSerifSC-Bold', 14)
    canv.drawString(MARGIN_L, auth_y + 1.0*cm, "William Francis Fohom")

    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC', 10)
    canv.drawString(MARGIN_L, auth_y + 0.4*cm, "Data Analyst")
    canv.drawString(MARGIN_L, auth_y - 0.1*cm, "Administrateur National de Ventes")

    # Bottom band
    canv.setFillColor(NAVY)
    canv.rect(0, 0, PAGE_W, 1.5*cm, fill=1, stroke=0)
    canv.setFillColor(WHITE)
    canv.setFont('NotoSerifSC', 9)
    canv.drawString(MARGIN_L, 0.55*cm, "BELGOCAM SA — Document confidentiel — Usage interne")
    canv.setFont('NotoSerifSC-Light', 8)
    canv.drawRightString(PAGE_W - MARGIN_R, 0.55*cm, "Analyse Bundle Soja-Concentrés — Juillet 2026")

    canv.restoreState()


def draw_body_page(canv, doc):
    """Header/footer for body pages — Bundle Soja-Concentrés."""
    canv.saveState()
    # Top header line
    canv.setStrokeColor(NAVY)
    canv.setLineWidth(0.5)
    canv.line(MARGIN_L, PAGE_H - 1.3*cm, PAGE_W - MARGIN_R, PAGE_H - 1.3*cm)
    canv.setFillColor(NAVY)
    canv.setFont('NotoSerifSC-Bold', 9)
    canv.drawString(MARGIN_L, PAGE_H - 1.0*cm, "BELGOCAM SA")
    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC', 9)
    canv.drawString(MARGIN_L + 3*cm, PAGE_H - 1.0*cm, "Analyse Bundle Soja-Concentrés — Juillet 2026")
    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC-Light', 8)
    canv.drawRightString(PAGE_W - MARGIN_R, PAGE_H - 1.0*cm, "William Francis Fohom — Data Analyst")

    # Footer
    canv.setStrokeColor(GRAY_LIGHT)
    canv.setLineWidth(0.3)
    canv.line(MARGIN_L, 1.3*cm, PAGE_W - MARGIN_R, 1.3*cm)
    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC-Light', 8)
    canv.drawString(MARGIN_L, 0.9*cm, "BELGOCAM SA — Document confidentiel")
    page_num = canv.getPageNumber() - 1  # subtract cover
    canv.drawRightString(PAGE_W - MARGIN_R, 0.9*cm, f"Page {page_num}")
    canv.restoreState()


# ===== HELPER FUNCTIONS (identiques à build_zero_achat_pdf.py) =====
def make_table(data, col_widths=None, header_row=True, font_size=9, align='LEFT'):
    """Create a styled table with BELGOCAM colors. Auto-scales to full content width."""
    if col_widths is None:
        n_cols = len(data[0])
        col_widths = [CONTENT_W / n_cols] * n_cols
    else:
        total = sum(col_widths)
        if total < CONTENT_W * 0.95:
            scale = CONTENT_W / total
            col_widths = [w * scale for w in col_widths]

    cell_style = ParagraphStyle('cell_fs', parent=CELL, fontSize=font_size, leading=font_size + 3)
    cell_style_r = ParagraphStyle('cell_r_fs', parent=CELL_RIGHT, fontSize=font_size, leading=font_size + 3)
    cell_style_c = ParagraphStyle('cell_c_fs', parent=CELL_CENTER, fontSize=font_size, leading=font_size + 3)
    cell_style_w = ParagraphStyle('cell_w_fs', parent=CELL_WHITE, fontSize=font_size, leading=font_size + 3)

    wrapped = []
    for i, row in enumerate(data):
        wrapped_row = []
        for j, cell in enumerate(row):
            if isinstance(cell, str):
                if i == 0 and header_row:
                    style = cell_style_w
                else:
                    style = cell_style_r if align == 'RIGHT' else cell_style
                wrapped_row.append(Paragraph(cell, style))
            else:
                wrapped_row.append(cell)
        wrapped.append(wrapped_row)

    t = Table(wrapped, colWidths=col_widths, repeatRows=1 if header_row else 0)
    ts = TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('GRID', (0, 0), (-1, -1), 0.4, GRAY_LIGHT),
    ])
    if header_row:
        ts.add('BACKGROUND', (0, 0), (-1, 0), NAVY)
        ts.add('TEXTCOLOR', (0, 0), (-1, 0), WHITE)
        ts.add('FONTNAME', (0, 0), (-1, 0), 'NotoSerifSC-Bold')
    for i in range(1, len(data)):
        if i % 2 == 0:
            ts.add('BACKGROUND', (0, i), (-1, i), GRAY_VLIGHT)
    t.setStyle(ts)
    return t


def kpi_card(label, value, sublabel=""):
    """Create a small KPI card as a Table."""
    data = [
        [Paragraph(f'<b>{value}</b>', ParagraphStyle('kpi_v', fontName='NotoSerifSC-Bold', fontSize=13, textColor=NAVY, alignment=TA_CENTER, leading=16))],
        [Paragraph(label, ParagraphStyle('kpi_l', fontName='NotoSerifSC', fontSize=8, textColor=GRAY, alignment=TA_CENTER, leading=10))],
    ]
    if sublabel:
        data.append([Paragraph(sublabel, ParagraphStyle('kpi_s', fontName='NotoSerifSC-Light', fontSize=7, textColor=GRAY, alignment=TA_CENTER, leading=9))])
    t = Table(data, colWidths=[3.2*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), GRAY_VLIGHT),
        ('LINEBELOW', (0, 0), (-1, 0), 0.5, GOLD),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('BOX', (0, 0), (-1, -1), 0.3, GRAY_LIGHT),
    ]))
    return t


def kpi_row(cards):
    """Place multiple KPI cards in a row."""
    n = len(cards)
    w = CONTENT_W / n
    t = Table([cards], colWidths=[w]*n)
    t.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]))
    return t


def insight_box(text_html, color=NAVY):
    """Highlight a key insight in a tinted box with gold left border."""
    p = Paragraph(text_html, ParagraphStyle('insight', parent=BODY, fontName='NotoSerifSC', fontSize=10, leading=14, textColor=BLACK, alignment=TA_LEFT))
    t = Table([[p]], colWidths=[CONTENT_W])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), GRAY_VLIGHT),
        ('LINEBEFORE', (0, 0), (0, -1), 3, GOLD),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('BOX', (0, 0), (-1, -1), 0.3, GRAY_LIGHT),
    ]))
    return t


def section_divider():
    return HRFlowable(width="100%", thickness=0.5, color=GOLD, spaceBefore=6, spaceAfter=10)


# ===== BUILD CONTENT =====
def build_story():
    story = []

    # ---------- PAGE 1: COVER (drawn by canvas) ----------
    story.append(Spacer(1, PAGE_H - MARGIN_T - MARGIN_B))  # fill the page
    story.append(PageBreak())

    # =====================================================================
    # SECTION 1 — CONTEXTE ET PROBLÉMATIQUE
    # =====================================================================
    story.append(Paragraph("1. Contexte et problématique", H1))
    story.append(section_divider())

    story.append(Paragraph(
        "<b>En juillet 2026</b>, le prix du soja (tourteaux) a subi <b>deux hausses successives</b> : <b>+1 000 FCFA/sac</b> en début de mois, puis <b>+2 000 FCFA/sac supplémentaires</b> par sac de 50 kg à partir du <b>23/07/2026</b> — soit <b>+3 000 FCFA/sac cumulé</b> au total sur le mois. "
        "Cette décision tarifaire intervient alors que BELGOCAM bénéficie d'une fenêtre d'opportunité exceptionnelle — "
        "la rupture de stock du principal concurrent sur le soja — qui se traduit par une forte croissance des volumes "
        "TOURTEAUX (+46,5% YoY en S1 2026).",
        BODY
    ))
    story.append(Paragraph(
        "Le <b>bundle soja-concentrés (ratio 3:1)</b> est l'action n°1 du plan d'action commercial BELGOCAM : "
        "chaque commande de soja doit inclure au moins 1 sac de concentré pour 3 sacs de soja. "
        "Ce ratio vise à sécuriser le cross-sell vers les concentrés — produit à plus forte marge — et à éviter "
        "que les clients ne profitent du soja (produit d'appel) sans consommer les concentrés.",
        BODY
    ))
    story.append(Paragraph("<b>Objectif de l'analyse</b>", H3))
    story.append(Paragraph(
        "Vérifier si la hausse du prix du soja en juillet 2026 <b>impacte le comportement d'achat des concentrés</b> "
        "auprès des clients bundle (clients ayant acheté à la fois du soja et des concentrés dans le même mois).",
        BODY
    ))
    story.append(Paragraph("<b>Hypothèse du client</b>", H3))
    story.append(insight_box(
        "<b>« Les clients ne voulant pas augmenter leur budget et voulant respecter le bundle préfèrent équilibrer "
        "en prenant moins de concentrés. »</b><br/><br/>"
        "Cette hypothèse sera testée agence par agence et région par région, puis validée ou infirmée "
        "à travers une analyse comparative S1 2026 (moyenne mensuelle Janvier-Juin) vs Juillet 2026 (au 24/07, "
        "commandes Livrées)."
    ))
    story.append(Spacer(1, 0.3*cm))

    story.append(PageBreak())

    # =====================================================================
    # SECTION 2 — MÉTHODOLOGIE
    # =====================================================================
    story.append(Paragraph("2. Méthodologie", H1))
    story.append(section_divider())

    story.append(Paragraph("2.1 Périmètre d'analyse", H2))
    story.append(Paragraph(
        "L'analyse porte exclusivement sur les <b>clients bundle</b>, définis comme les clients ayant acheté "
        "<b>à la fois du soja ET des concentrés dans le même mois</b>. Cette définition filtre les clients mono-produit "
        "et permet d'isoler le comportement d'achat couplé soja + concentrés — cœur de l'enjeu du ratio 3:1.",
        BODY
    ))
    story.append(Paragraph(
        "<b>Base de clients bundle :</b> 958 clients en moyenne mensuelle S1 2026 → 540 clients actifs en juillet 2026 "
        "(au 24/07, projection partielle du mois).",
        BODY_BOLD
    ))

    story.append(Paragraph("2.2 Comparaison temporelle", H2))
    story.append(Paragraph(
        "Deux périodes sont comparées :<br/>"
        "• <b>S1 2026 (référence)</b> : moyenne mensuelle calculée sur la période Janvier-Juin 2026.<br/>"
        "• <b>Juillet 2026</b> : données au 24/07, portant sur les commandes au statut <b>Livrée</b> "
        "(commandes fermes, hors panier/brouillon).",
        BODY
    ))

    story.append(Paragraph("2.3 Indicateurs suivis", H2))
    ind_data = [
        ["Indicateur", "Définition", "Unité"],
        ["Prix moyen soja/sac", "CA soja ÷ volume soja (en sacs)", "FCFA/sac"],
        ["Prix moyen conc/sac", "CA concentrés ÷ volume concentrés (en sacs)", "FCFA/sac"],
        ["Budget moyen par client", "CA total (soja + conc) ÷ nb clients bundle", "FCFA"],
        ["Sacs soja/client", "Volume soja (sacs) ÷ nb clients bundle", "sacs"],
        ["Sacs conc/client", "Volume concentrés (sacs) ÷ nb clients bundle", "sacs"],
        ["Ratio soja/conc", "Sacs soja/client ÷ sacs conc/client", "ratio X:1"],
    ]
    story.append(make_table(ind_data, col_widths=[5*cm, 8*cm, 3.5*cm]))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("2.4 Niveaux d'analyse", H2))
    story.append(Paragraph(
        "Trois niveaux d'agrégation sont mobilisés pour croiser les lectures :<br/>"
        "• <b>Analyse globale</b> : toutes agences confondues (14 agences BELGOCAM).<br/>"
        "• <b>Analyse par agence</b> : comparaison des 14 agences sur les indicateurs clés et verdict.<br/>"
        "• <b>Analyse par région</b> : agrégation par région commerciale (Ouest, Centre, Littoral).<br/>"
        "• <b>Focus NDOBO</b> : analyse approfondie de l'agence la plus problématique (ratio 5,5:1 en juillet).",
        BODY
    ))

    story.append(PageBreak())

    # =====================================================================
    # SECTION 3 — ANALYSE GLOBALE (toutes agences)
    # =====================================================================
    story.append(Paragraph("3. Analyse globale (toutes agences)", H1))
    story.append(section_divider())

    story.append(Paragraph(
        "Vue agrégée des indicateurs bundle sur l'ensemble des 14 agences BELGOCAM, "
        "comparant la moyenne mensuelle S1 2026 et les données Juillet 2026 (au 24/07).",
        BODY
    ))

    story.append(Paragraph("3.1 Tableau de bord global", H2))
    global_data = [
        ["Indicateur", "Moyenne S1", "Juillet", "Évolution", "Lecture"],
        ["Prix soja/sac", "16 198", "18 322", "+13,1%", "Hausse cumulée +3 000 FCFA (1 000 + 2 000)"],
        ["Prix conc/sac", "33 382", "33 560", "+0,5%", "Stabilité quasi totale"],
        ["Budget/client", "2 974 944", "3 925 013", "+31,9%", "Les clients augmentent leur budget"],
        ["Sacs soja/client", "132,7", "140,9", "+6,2%", "Achat massif de soja (rupture concurrente)"],
        ["Sacs conc/client", "49,3", "40,0", "-18,9%", "Les concentrés reculent en volume par client"],
        ["Ratio soja/conc", "2,7:1", "3,5:1", "dégradation", "Dégradation du ratio au-delà de 3:1"],
    ]
    story.append(make_table(global_data, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2*cm, 5.5*cm]))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("3.2 Lecture des résultats", H2))
    story.append(Paragraph(
        "<b>Le prix du soja a augmenté comme attendu</b> (+13,1%, soit précisément +2 124 FCFA/sac en moyenne sur juillet, "
        "reflétant l'effet combiné de la hausse de début de mois (+1 000) et de la hausse du 23/07 (+2 000) qui ne joue que sur 2 jours), "
        "tandis que le prix des concentrés est resté stable (+0,5%). "
        "La première conséquence mécanique aurait dû être une <b>compression du budget client</b> "
        "ou une <b>réduction des volumes de soja</b>.",
        BODY
    ))
    story.append(Paragraph(
        "<b>Or, c'est l'inverse qui se produit :</b> le budget client bondit de +31,9% et les sacs de soja par client "
        "augmentent de +6,2%. Ce paradoxe s'explique par la <b>rupture de stock du concurrent</b> — "
        "les clients viennent massivement chez BELGOCAM pour s'approvisionner en soja, malgré la hausse tarifaire. "
        "Le soja joue pleinement son rôle de produit d'appel.",
        BODY_BOLD
    ))
    story.append(Paragraph(
        "<b>Côté concentrés :</b> les sacs par client reculent de <b>-18,9%</b>. La dégradation est plus prononcée "
        "qu'au point précédent (18/07), ce qui confirme la dérive du bundle. Le différentiel de croissance "
        "(soja +6,2% vs conc -18,9%) entraîne mécaniquement une <b>dégradation nette du ratio soja/conc</b>, "
        "qui passe de 2,7:1 à 3,5:1 — le ratio cible de 3:1 est franchi dans le mauvais sens.",
        BODY
    ))

    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph("3.3 Insight clé", H2))
    story.append(insight_box(
        "<b>L'hypothèse du « budget constant » n'est PAS confirmée globalement</b> — le budget client a augmenté de +64%.<br/><br/>"
        "<b>MAIS le ratio se dégrade :</b> les clients achètent proportionnellement plus de soja et moins de concentrés. "
        "Le bundle est respecté en apparence (les clients prennent les deux produits) mais le <b>ratio 3:1 n'est pas tenu</b>."
    ))

    story.append(PageBreak())

    # =====================================================================
    # SECTION 4 — ANALYSE PAR AGENCE
    # =====================================================================
    story.append(Paragraph("4. Analyse par agence (14 agences)", H1))
    story.append(section_divider())

    story.append(Paragraph(
        "Comparaison détaillée des 14 agences BELGOCAM sur les indicateurs clés du bundle soja-concentrés. "
        "Le verdict synthétise la dynamique observée : ratio stable, en hausse, en dégradation ou cas critique (conc stagnant/baisse).",
        BODY
    ))

    agence_data = [
        ["Agence", "Région", "Prix soja Δ%", "Budget Δ%", "Sacs conc Δ%", "Ratio S1", "Ratio Juil", "Verdict"],
        ["Famla", "Ouest", "+13,7%", "+2,5%", "-16,4%", "3,2:1", "3,9:1", "⚠️ Ratio se dégrade"],
        ["Ndobo", "Littoral", "+12,8%", "+5,7%", "-39,2%", "2,7:1", "5,5:1", "🔴 Conc en chute libre"],
        ["Messassi", "Centre", "+12,9%", "-0,1%", "-22,1%", "2,6:1", "3,4:1", "⚠️ Ratio se dégrade"],
        ["Djeleng", "Ouest", "+12,6%", "+1,0%", "-35,2%", "2,2:1", "4,1:1", "🔴 Conc en chute"],
        ["Mbouda", "Ouest", "+13,9%", "+7,0%", "-11,7%", "2,6:1", "3,2:1", "⚠️ Conc stagnant"],
        ["Village", "Littoral", "+12,8%", "-16,5%", "-24,9%", "2,7:1", "2,8:1", "⚠️ Budget en baisse"],
        ["Bertoua", "Centre", "+9,7%", "+33,4%", "+3,9%", "2,0:1", "2,8:1", "✅ Bonne croissance"],
        ["Nkongsamba", "Littoral", "+13,6%", "-34,0%", "-39,9%", "2,5:1", "2,5:1", "🔴 Budget et conc en baisse"],
        ["Nkoabang", "Centre", "+12,1%", "-26,1%", "-39,5%", "3,1:1", "3,8:1", "🔴 Conc en chute"],
        ["Ngaoundere", "Centre", "+12,6%", "+18,9%", "+15,3%", "2,2:1", "2,1:1", "✅ Excellente"],
        ["Buea", "Littoral", "+15,1%", "-4,3%", "-1,4%", "3,5:1", "2,9:1", "✅ Ratio s'améliore"],
        ["Ahala", "Centre", "+12,7%", "+19,1%", "-4,8%", "2,2:1", "2,9:1", "✅ Bonne croissance"],
        ["Nkolbisson", "Centre", "+14,3%", "-19,5%", "-30,0%", "2,1:1", "2,5:1", "🔴 Conc en baisse"],
        ["Pk11", "Littoral", "+13,7%", "+31,6%", "-4,8%", "3,2:1", "4,5:1", "⚠️ Ratio se dégrade"],
    ]
    story.append(make_table(agence_data, col_widths=[2.2*cm, 1.6*cm, 1.8*cm, 1.8*cm, 1.8*cm, 1.6*cm, 1.6*cm, 3.4*cm], font_size=7.5))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("4.1 Lecture par cluster de verdict", H2))
    story.append(Paragraph(
        "<b>🔴 Agences critiques (conc en chute libre) :</b> <b>NDOBO</b> (-39,2% conc, ratio 5,5:1), <b>DJELENG</b> (-35,2%, ratio 4,1:1), "
        "<b>NKOABANG</b> (-39,5%, ratio 3,8:1), <b>NKOLBISSON</b> (-30,0%), <b>NKONGSAMBA</b> (-39,9%). "
        "À NDOBO, les concentrés s'effondrent alors que le soja explose — le ratio se dégrade fortement. "
        "À NKONGSAMBA, le budget client baisse fortement (-34,0%) ET les concentrés chutent : l'hypothèse du budget constant y est confirmée.",
        BODY_BOLD
    ))
    story.append(Paragraph(
        "<b>⚠️ Agences en dégradation du ratio :</b> FAMLA (3,9:1), MESSASSI (3,4:1), MBOUDA (3,2:1), VILLAGE (2,8:1 mais budget en baisse), PK11 (4,5:1 — deuxième pire ratio). "
        "Ces agences ont une croissance des concentrés négative ou insuffisante pour suivre celle du soja.",
        BODY
    ))
    story.append(Paragraph(
        "<b>✅ Agences performantes :</b> BERTOUA (+3,9% conc), NGAOUNDERE (+15,3% conc), BUEA (-1,4% conc mais ratio s'améliore), AHALA (-4,8% conc mais budget +19,1%). "
        "Le Centre reste globalement meilleur (4 de ses 6 agences en verdict positif), porté par NGAOUNDERE et BERTOUA "
        "qui affichent des croissances de concentrés positives.",
        BODY
    ))

    story.append(PageBreak())

    # =====================================================================
    # SECTION 5 — ANALYSE PAR RÉGION
    # =====================================================================
    story.append(Paragraph("5. Analyse par région", H1))
    story.append(section_divider())

    story.append(Paragraph(
        "Agrégation par région commerciale : <b>Ouest</b> (3 agences), <b>Centre</b> (6 agences), <b>Littoral</b> (5 agences). "
        "Cette vue révèle les dynamiques géographiques contrastées du bundle soja-concentrés.",
        BODY
    ))

    # ----- 5.1 OUEST -----
    story.append(Paragraph("5.1 Région OUEST (3 agences : Famla, Djeleng, Mbouda)", H2))
    story.append(Paragraph(
        "<b>Base clients bundle :</b> 1 106 clients S1 → 206 clients en Juillet (au 24/07).",
        BODY_BOLD
    ))
    ouest_data = [
        ["Indicateur", "S1", "Juillet", "Évolution"],
        ["Prix soja", "16 163", "18 341", "+13,5%"],
        ["Prix conc", "34 190", "34 137", "-0,2%"],
        ["Budget/client", "4 641 702", "4 760 536", "+2,6%"],
        ["Sacs conc/client", "58,0", "46,2", "-20,3%"],
        ["Ratio soja/conc", "2,8:1", "3,8:1", "dégradation"],
    ]
    story.append(make_table(ouest_data, col_widths=[4.5*cm, 3.5*cm, 3.5*cm, 4*cm]))
    story.append(Spacer(1, 0.2*cm))
    story.append(insight_box(
        "<b>Verdict OUEST :</b> Budget quasi-stable (+2,6%) — conc en chute (-20,3%) et ratio se dégrade fortement (2,8 → 3,8). DJELENG tire la région vers le bas (-35,2% conc)."
    ))
    story.append(Spacer(1, 0.3*cm))

    # ----- 5.2 CENTRE -----
    story.append(Paragraph("5.2 Région CENTRE (6 agences)", H2))
    story.append(Paragraph(
        "<b>Base clients bundle :</b> 954 clients S1 → 166 clients en Juillet (au 24/07).",
        BODY_BOLD
    ))
    centre_data = [
        ["Indicateur", "S1", "Juillet", "Évolution"],
        ["Prix soja", "16 695", "18 712", "+12,1%"],
        ["Prix conc", "33 750", "33 956", "+0,6%"],
        ["Budget/client", "3 682 305", "3 941 327", "+7,0%"],
        ["Sacs conc/client", "50,2", "43,8", "-12,7%"],
        ["Ratio soja/conc", "2,4:1", "3,0:1", "dégradation modérée"],
    ]
    story.append(make_table(centre_data, col_widths=[4.5*cm, 3.5*cm, 3.5*cm, 4*cm]))
    story.append(Spacer(1, 0.2*cm))
    story.append(insight_box(
        "<b>Verdict CENTRE :</b> ✅ Conc en baisse modérée (-12,7%), ratio limite (3,0:1). Meilleure région sur le bundle — NGAOUNDERE et BERTOUA tirent la région vers le haut."
    ))
    story.append(Spacer(1, 0.3*cm))

    story.append(PageBreak())

    # ----- 5.3 LITTORAL -----
    story.append(Paragraph("5.3 Région LITTORAL (5 agences)", H2))
    story.append(Paragraph(
        "<b>Base clients bundle :</b> 1 054 clients S1 → 168 clients en Juillet (au 24/07).",
        BODY_BOLD
    ))
    littoral_data = [
        ["Indicateur", "S1", "Juillet", "Évolution"],
        ["Prix soja", "15 773", "17 828", "+13,0%"],
        ["Prix conc", "31 715", "31 825", "+0,3%"],
        ["Budget/client", "3 014 013", "2 884 384", "-4,3%"],
        ["Sacs conc/client", "39,5", "28,7", "-27,2%"],
        ["Ratio soja/conc", "2,8:1", "3,8:1", "forte dégradation"],
    ]
    story.append(make_table(littoral_data, col_widths=[4.5*cm, 3.5*cm, 3.5*cm, 4*cm]))
    story.append(Spacer(1, 0.2*cm))
    story.append(insight_box(
        "<b>Verdict LITTORAL :</b> 🔴 Conc en forte chute (-27,2%), ratio se dégrade fortement (2,8 → 3,8) ET budget en baisse (-4,3%). "
        "Région la plus touchée — NDOBO (5,5:1), NKONGSAMBA (-39,9% conc) et PK11 (4,5:1) en sont les principaux contributeurs."
    ))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("5.4 Synthèse régionale comparée", H2))
    synth_region = [
        ["Région", "Budget Δ%", "Conc Δ%", "Ratio S1", "Ratio Juil", "Verdict"],
        ["Ouest (3 agences)", "+2,6%", "-20,3%", "2,8:1", "3,8:1", "⚠️ Dégradation"],
        ["Centre (6 agences)", "+7,0%", "-12,7%", "2,4:1", "3,0:1", "✅ Modéré"],
        ["Littoral (5 agences)", "-4,3%", "-27,2%", "2,8:1", "3,8:1", "🔴 Forte dégradation"],
    ]
    story.append(make_table(synth_region, col_widths=[3.5*cm, 2.3*cm, 2.3*cm, 2.3*cm, 2.3*cm, 3.8*cm]))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Lecture comparée :</b> le Centre tire son épingle du jeu grâce à une baisse limitée des concentrés "
        "(-12,7%), portée par NGAOUNDERE (+15,3%) et BERTOUA (+3,9%) qui sont les seules agences en croissance de concentrés. "
        "Le Littoral est pénalisé par NDOBO, NKONGSAMBA et PK11 où les concentrés s'effondrent (-27,2% en moyenne). "
        "L'Ouest présente un profil intermédiaire avec budget quasi-stable mais une dégradation nette du ratio.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # =====================================================================
    # SECTION 6 — FOCUS NDOBO (CAS LE PLUS PROBLÉMATIQUE)
    # =====================================================================
    story.append(Paragraph("6. Focus NDOBO (cas le plus problématique)", H1))
    story.append(section_divider())

    story.append(Paragraph(
        "NDOBO est l'agence la plus problématique du réseau sur le bundle soja-concentrés : "
        "ratio de 5,5:1 en juillet (vs 2,7:1 en S1), chute libre des concentrés (-39,2%) et budget en hausse limitée (+5,7%). "
        "Cette agence concentre les symptômes qui valident l'hypothèse du « budget constant » émise par les clients.",
        BODY
    ))

    # ----- 6.1 Évolution mensuelle -----
    story.append(Paragraph("6.1 Évolution mensuelle Janvier → Juillet (NDOBO)", H2))
    story.append(Paragraph(
        "Vue mensuelle des indicateurs bundle sur l'agence NDOBO, avec la moyenne S1 comme référence.",
        BODY
    ))
    ndobo_mensuel = [
        ["Mois", "Clients", "Lignes", "Vol conc (t)", "CA conc (M)", "Vol/client (t)", "CA/client (M)", "Vol/ligne (t)", "Bundle %", "Ratio soja/conc"],
        ["Janvier", "74", "315", "139,8", "88,3", "1,9", "1,2", "0,44", "84%", "2,3:1"],
        ["Février", "65", "291", "135,5", "86,4", "2,1", "1,3", "0,47", "80%", "1,9:1"],
        ["Mars", "72", "312", "162,4", "103,5", "2,3", "1,4", "0,52", "85%", "1,8:1"],
        ["Avril", "60", "268", "128,0", "79,8", "2,1", "1,3", "0,48", "85%", "2,3:1"],
        ["Mai", "65", "274", "126,6", "81,1", "1,9", "1,2", "0,46", "88%", "2,9:1"],
        ["Juin", "68", "320", "141,4", "89,3", "2,1", "1,3", "0,44", "93%", "3,6:1"],
        ["Juillet (24j)", "55", "190", "69,6", "44,2", "1,3", "0,8", "0,37", "95%", "5,5:1"],
        ["Moy S1", "67", "297", "138,9", "88,1", "2,1", "1,3", "0,47", "86%", "2,7:1"],
    ]
    story.append(make_table(ndobo_mensuel, col_widths=[1.7*cm, 1.4*cm, 1.4*cm, 1.7*cm, 1.7*cm, 1.8*cm, 1.8*cm, 1.7*cm, 1.5*cm, 2*cm], font_size=7.5))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Lecture de la dynamique mensuelle :</b> le ratio soja/conc passe de 2,3:1 en janvier à 5,5:1 en juillet, "
        "avec une rupture nette entre mai (2,9:1) et juin (3,6:1) qui s'aggrave en juillet. La hausse tarifaire du soja en juillet "
        "(+1 000 puis +2 000 FCFA/sac au 23/07) aggrave fortement le mouvement. Sur la même période, le volume de concentrés par client "
        "passe de 2,1 t à 1,3 t (-38%). Le taux de bundle augmente (84% → 95%), ce qui signifie que les clients prennent bien les deux produits — "
        "mais dans des proportions de plus en plus déséquilibrées.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ----- 6.2 Budget client NDOBO -----
    story.append(Paragraph("6.2 Budget client NDOBO — S1 vs Juillet", H2))
    story.append(Paragraph(
        "Comparaison détaillée du budget bundle client à NDOBO entre la moyenne mensuelle S1 et juillet 2026.",
        BODY
    ))
    ndobo_budget = [
        ["Indicateur", "Moy S1", "Juillet", "Évolution"],
        ["Prix soja/sac", "15 692", "17 696", "+12,8%"],
        ["Prix conc/sac", "31 560", "31 868", "+1,0%"],
        ["Sacs soja/client", "114,0", "139,8", "+22,6%"],
        ["Sacs conc/client", "41,6", "25,3", "-39,2%"],
        ["Budget/client", "3 103 620", "3 279 596", "+5,7%"],
        ["Ratio soja/conc", "2,7:1", "5,5:1", "dégradation critique"],
    ]
    story.append(make_table(ndobo_budget, col_widths=[4.5*cm, 3.5*cm, 3.5*cm, 4*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Constat NDOBO :</b> les clients augmentent fortement leurs achats de soja (+22,6% — effet de la rupture concurrente) mais réduisent "
        "drastiquement leurs achats de concentrés (-39,2%, soit près de 40%). Le budget total augmente légèrement de +5,7%, "
        "ce qui valide partiellement l'hypothèse du « budget constant » : <b>les clients compensent la hausse du prix du soja en réduisant les concentrés</b> "
        "plutôt qu'en augmentant significativement leur budget global. Le ratio explose de 2,7:1 à 5,5:1.",
        BODY_BOLD
    ))

    story.append(Spacer(1, 0.3*cm))

    # ----- 6.3 Simulation budget constant -----
    story.append(Paragraph("6.3 Simulation budget constant NDOBO", H2))
    story.append(Paragraph(
        "Pour confirmer l'hypothèse, simulons le budget qu'aurait un client NDOBO s'il maintenait son niveau de concentrés "
        "de S1 (41,6 sacs) avec les nouveaux prix de juillet.",
        BODY
    ))
    story.append(insight_box(
        "<b>Si le client maintenait ses 41,6 sacs de concentrés (niveau S1) avec le nouveau prix du soja, "
        "son budget passerait de 3 103 620 F à 3 605 446 F (+16,2%).</b> "
        "En réduisant les concentrés à 25,3 sacs, le client économise 519 174 FCFA et limite la hausse de son budget "
        "à +5,7%."
    ))
    story.append(Spacer(1, 0.2*cm))
    simul_data = [
        ["Scénario", "Sacs soja", "Sacs conc", "Budget", "vs S1"],
        ["S1 (référence)", "114,0", "41,6", "3 103 620 F", "—"],
        ["Juillet réel", "139,8", "25,3", "3 279 596 F", "+5,7%"],
        ["Juillet simulé (conc maintenu)", "139,8", "41,6", "3 605 446 F", "+16,2%"],
    ]
    story.append(make_table(simul_data, col_widths=[5.5*cm, 2.5*cm, 2.5*cm, 3.5*cm, 2*cm]))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Conclusion de la simulation :</b> l'hypothèse client est confirmée à NDOBO. Le client arbitre "
        "explicitement entre deux options : <b>(a) augmenter son budget de +16,2%</b> pour conserver son niveau de concentrés, "
        "ou <b>(b) réduire les concentrés de -39,2%</b> pour limiter la hausse du budget à +5,7%. "
        "Le client choisit massivement l'option (b) — d'où la dégradation du ratio à 5,5:1.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ----- 6.4 Distribution des ratios en juillet -----
    story.append(Paragraph("6.4 Distribution des ratios en juillet (NDOBO)", H2))
    story.append(Paragraph(
        "Répartition des commandes bundle de juillet à NDOBO par tranche de ratio soja/conc. "
        "Permet de visualiser la dispersion du déséquilibre.",
        BODY
    ))
    distrib_data = [
        ["Tranche", "Part des commandes", "Lecture"],
        ["≤ 3:1 (objectif)", "≈ 30%", "Base saine mais minoritaire"],
        ["3:1 à 5:1", "≈ 30%", "Déséquilibre modéré"],
        ["5:1 à 10:1", "≈ 25%", "Déséquilibre marqué"],
        ["> 10:1", "≈ 15%", "Déséquilibre critique — conc symbolique"],
        ["TOTAL", "100%", "55 commandes bundle au 24/07"],
    ]
    story.append(make_table(distrib_data, col_widths=[5*cm, 4*cm, 6*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Lecture de la distribution :</b><br/>"
        "• <b>≈30% des commandes respectent l'objectif 3:1</b> — une base saine existe à NDOBO mais elle est minoritaire.<br/>"
        "• <b>≈30% sont entre 3:1 et 5:1</b> — déséquilibre modéré, récupérable par négociation commerciale.<br/>"
        "• <b>≈25% sont entre 5:1 et 10:1</b> — déséquilibre marqué, le concentré devient accessoire.<br/>"
        "• <b>≈15% dépassent 10:1</b> — déséquilibre critique, le concentré est purement symbolique. "
        "Ces commandes représentent un échec total du bundle.<br/>",
        BODY_BOLD
    ))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Recommandation opérationnelle :</b> imposer le ratio 3:1 strictement — refuser ou alerter sur les commandes "
        "avec ratio &gt; 5:1 (≈40% des commandes à NDOBO en juillet), et traiter en priorité les ratios &gt; 10:1 (≈15% des commandes) "
        "qui témoignent d'un contournement manifeste du bundle.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # =====================================================================
    # SECTION 7 — CONCLUSIONS ET RECOMMANDATIONS
    # =====================================================================
    story.append(Paragraph("7. Conclusions et recommandations", H1))
    story.append(section_divider())

    # ----- 7.1 Constats -----
    story.append(Paragraph("7.1 Constats", H2))
    story.append(Paragraph(
        "<b>1. Le phénomène est GÉNÉRAL</b> — toutes agences confondues, le ratio soja/conc passe de 2,7 à 3,5. "
        "Aucune agence n'est épargnée par la dégradation du bundle en juillet 2026.",
        BODY
    ))
    story.append(Paragraph(
        "<b>2. Mais il est PLUS PRONONCÉ à NDOBO</b> (ratio 5,5:1) <b>et PK11</b> (4,5:1). "
        "Ces deux agences du Littoral concentrent les cas les plus critiques, avec des clients qui réduisent "
        "massivement leurs achats de concentrés pour compenser la hausse du prix du soja.",
        BODY_BOLD
    ))
    story.append(Paragraph(
        "<b>3. Le Littoral est la région la plus touchée</b> — ratio 2,8:1 → 3,8:1, avec une chute des concentrés "
        "de -27,2% (vs -12,7% au Centre et -20,3% à l'Ouest). La dynamique de dégradation y est structurelle.",
        BODY
    ))
    story.append(Paragraph(
        "<b>4. Le Centre s'en sort mieux</b> — ratio 2,4:1 → 3,0:1, avec deux agences en croissance de concentrés "
        "(NGAOUNDERE +15,3%, BERTOUA +3,9%). Le Centre est un modèle à reproduire.",
        BODY
    ))
    story.append(Paragraph(
        "<b>5. L'hypothèse « budget constant » est confirmée à NDOBO</b> (budget +5,7% seulement malgré +22,6% de soja) <b>mais pas globalement</b> "
        "(+31,9% de budget toutes agences). À NDOBO, les clients arbitrent explicitement contre les concentrés ; "
        "ailleurs, ils acceptent d'augmenter leur budget mais privilégient tout de même le soja.",
        BODY_BOLD
    ))

    story.append(Spacer(1, 0.3*cm))

    # ----- 7.2 Recommandations -----
    story.append(Paragraph("7.2 Recommandations", H2))
    story.append(Paragraph(
        "Plan d'action en 9 axes pour endiguer la dégradation du ratio et restaurer l'équilibre du bundle soja-concentrés.",
        BODY
    ))
    reco_data = [
        ["#", "Recommandation", "Cible / Périmètre", "Délai"],
        ["1", "Imposer le ratio 3:1 strictement — refuser les commandes avec ratio > 5:1", "Toutes agences — système de facturation", "Immédiat"],
        ["2", "Surveiller les ratios > 10:1 (20% des commandes à NDOBO en juillet)", "NDOBO + PK11 en priorité", "Hebdomadaire"],
        ["3", "Envisager une remise sur les concentrés dans le bundle pour compenser la hausse du soja", "Direction commerciale + produit", "30 jours"],
        ["4", "Focus terrain sur NDOBO et PK11 (agences les plus dégradées)", "Équipe commerciale Littoral", "15 jours"],
        ["5", "Capitaliser sur les bonnes pratiques du Centre (DJELENG, AHALA, NGAOUNDERE)", "Reproduire le modèle Centre", "60 jours"],
        ["6", "Comparer les prix des concentrés vs concurrents — si BELGOCAM est plus cher, facteur aggravant", "Direction produit + benchmark", "30 jours"],
        ["7", "Analyser le profil des clients à NDOBO — probablement plus de grossistes/provendiers sensibles au prix", "Data Analyst + commercial NDOBO", "45 jours"],
        ["8", "Annulation des commandes ≥ 10 t en soja sauf si concentré en ratio ≥ 2,5:1 (voir 7.3)", "Système de facturation + commerciaux", "Immédiat"],
        ["9", "Déployer les packs de bundle prédéfinis (voir 7.4) comme alternative aux commandes libres", "Toutes agences — argumentaire commercial", "15 jours"],
    ]
    story.append(make_table(reco_data, col_widths=[0.8*cm, 8*cm, 4.5*cm, 2.5*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph(
        "<b>Priorisation :</b> les recommandations 1, 8 et 9 sont à mener en parallèle et en urgence — "
        "le verrouillage du ratio 3:1 (recommandation 1) et l'annulation des commandes ≥ 10 t sans concentré suffisant (recommandation 8) "
        "ne produiront d'effet que si les commerciaux terrain sont mobilisés sur NDOBO et PK11 (recommandation 4) "
        "et outillés avec les packs de bundle prédéfinis (recommandation 9). "
        "La remise bundle (recommandation 3) est le levier structurel de moyen terme pour réconcilier "
        "la hausse du soja et le maintien des concentrés.",
        BODY_BOLD
    ))

    story.append(Spacer(1, 0.4*cm))

    # ----- 7.3 Règle d'annulation des commandes ≥ 10 t en soja -----
    story.append(Paragraph("7.3 Règle d'annulation — Commandes ≥ 10 t en soja sans concentré suffisant", H2))
    story.append(Paragraph(
        "<b>Règle opérationnelle :</b> toute commande de soja ≥ 10 tonnes doit obligatoirement inclure "
        "des concentrés dans un ratio ≥ 2,5:1 (soja:conc). En dessous de ce ratio, la commande est annulée "
        "ou suspendue jusqu'à ajustement par le client.<br/><br/>"
        "<b>Logique :</b> 10 tonnes de soja représentent environ 200 sacs de 50 kg. "
        "Avec un ratio minimum de 2,5:1, le client doit commander au minimum 80 sacs de concentrés "
        "(4 t) pour valider sa commande de soja. Cela garantit un volume minimal de concentrés "
        "sur les gros volumes de soja, qui sont précisément les commandes où le ratio se dégrade le plus.<br/><br/>"
        "<b>Seuil de 10 t :</b> ce seuil cible les gros acheteurs (provendiers, éleveurs industriels) "
        "qui représentent l'essentiel du volume soja mais aussi la plus grande déperdition de concentrés. "
        "Les petites commandes (< 10 t) ne sont pas concernées par cette règle — le bundle 3:1 classique s'applique.",
        BODY
    ))

    story.append(Paragraph("Tableau de référence — Ratio minimum 2,5:1 pour commandes ≥ 10 t de soja", H3))
    seuil_data = [
        ["Soja (t)", "Soja (sacs 50kg)", "Conc minimum (t)", "Conc minimum (sacs 50kg)", "Ratio imposé", "CA conc min (M FCFA)*"],
        ["10", "200", "4,0", "80", "2,5:1", "2,4"],
        ["15", "300", "6,0", "120", "2,5:1", "3,6"],
        ["20", "400", "8,0", "160", "2,5:1", "4,8"],
        ["25", "500", "10,0", "200", "2,5:1", "6,0"],
        ["30", "600", "12,0", "240", "2,5:1", "7,2"],
        ["50", "1 000", "20,0", "400", "2,5:1", "12,0"],
    ]
    story.append(make_table(seuil_data, col_widths=[1.8*cm, 2.8*cm, 2.5*cm, 3*cm, 2*cm, 3*cm], font_size=8))
    story.append(Paragraph(
        "<i>* CA concentrés minimum calculé à 30 000 FCFA/sac de 50 kg (prix moyen S1 2026).</i>",
        ParagraphStyle('fn', parent=SMALL, fontName='NotoSerifSC-Light', fontSize=7, textColor=GRAY, spaceBefore=2)
    ))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Impact estimé :</b> à NDOBO en juillet, environ 8 commandes avaient un ratio > 10:1 sur des volumes de soja significatifs. "
        "En appliquant la règle des 2,5:1 minimum, ces 8 commandes auraient généré un volume supplémentaire "
        "de concentrés estimé à ~15-20 t — soit +20 à +25% du volume concentrés réalisé en juillet à NDOBO.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ----- 7.4 Proposition de packs de bundle -----
    story.append(Paragraph("7.4 Proposition de packs de bundle prédéfinis", H2))
    story.append(Paragraph(
        "Plutôt que de laisser le client composer sa commande librement (ce qui conduit à des ratios déséquilibrés), "
        "les commerciaux proposent des <b>packs prédéfinis</b> avec un ratio soja:concentrés garanti. "
        "Chaque pack correspond à un profil d'éleveur et à un volume d'activité. "
        "Le client choisit un pack plutôt que de commander produit par produit.",
        BODY
    ))

    story.append(Paragraph("7.4.1 Packs pour élevage volaille (Chair + Ponte)", H3))
    packs_volaille = [
        ["Pack", "Soja (t)", "Conc Chair (t)", "Conc Ponte (t)", "Total conc (t)", "Ratio soja:conc", "CA pack (M FCFA)*", "Profil client"],
        ["Pack Découverte", "2,5", "0,5", "0,5", "1,0", "2,5:1", "1,6", "Petit éleveur (<500 sujets)"],
        ["Pack Standard", "5,0", "1,0", "1,0", "2,0", "2,5:1", "3,2", "Éleveur moyen (500-2000)"],
        ["Pack Confort", "5,0", "1,5", "1,0", "2,5", "2,0:1", "3,6", "Éleveur moyen+ (cross-sell renforcé)"],
        ["Pack Pro", "10,0", "2,0", "2,0", "4,0", "2,5:1", "6,4", "Grand éleveur (2000-5000)"],
        ["Pack Industrie", "20,0", "4,0", "4,0", "8,0", "2,5:1", "12,8", "Éleveur industriel (>5000)"],
    ]
    story.append(make_table(packs_volaille, col_widths=[2*cm, 1.5*cm, 1.8*cm, 1.8*cm, 1.5*cm, 1.8*cm, 2*cm, 3.5*cm], font_size=7))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("7.4.2 Packs pour élevage porcin", H3))
    packs_porc = [
        ["Pack", "Soja (t)", "Conc Porc (t)", "Ratio soja:conc", "CA pack (M FCFA)*", "Profil client"],
        ["Pack Découverte Porc", "2,5", "1,0", "2,5:1", "1,6", "Petit élevage porcin (<100 têtes)"],
        ["Pack Standard Porc", "5,0", "2,0", "2,5:1", "3,2", "Élevage moyen (100-500)"],
        ["Pack Pro Porc", "10,0", "4,0", "2,5:1", "6,4", "Grand élevage (>500)"],
    ]
    story.append(make_table(packs_porc, col_widths=[3*cm, 1.5*cm, 1.8*cm, 2*cm, 2.5*cm, 4.5*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("7.4.3 Packs mixtes (multi-espèces)", H3))
    packs_mixte = [
        ["Pack", "Soja (t)", "Conc Chair (t)", "Conc Ponte (t)", "Conc Porc (t)", "Total conc (t)", "Ratio", "Profil client"],
        ["Pack Mixte Standard", "5,0", "0,8", "0,7", "0,5", "2,0", "2,5:1", "Éleveur polyvalent"],
        ["Pack Mixte Pro", "10,0", "1,5", "1,5", "1,0", "4,0", "2,5:1", "Ferme diversifiée"],
        ["Pack Mixte Booster", "5,0", "1,0", "1,0", "0,5", "2,5", "2,0:1", "Avec Chick/Piglet Booster inclus"],
    ]
    story.append(make_table(packs_mixte, col_widths=[2.5*cm, 1.3*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.3*cm, 3.5*cm], font_size=7))
    story.append(Paragraph(
        "<i>* CA packs estimé à ~1 280 000 FCFA/t (soja à 18 000 F/sac + conc à 30 000 F/sac, mix 2,5:1).</i>",
        ParagraphStyle('fn2', parent=SMALL, fontName='NotoSerifSC-Light', fontSize=7, textColor=GRAY, spaceBefore=2)
    ))

    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Avantages des packs prédéfinis :</b><br/>"
        "• <b>Ratio garanti</b> — le commercial ne négocie plus les quantités produit par produit, il propose un pack.<br/>"
        "• <b>Simplicité</b> — le client comprend immédiatement ce qu'il achète et pourquoi.<br/>"
        "• <b>Tarification transparente</b> — chaque pack a un prix global, pas de calcul ligne par ligne.<br/>"
        "• <b>Cross-sell intégré</b> — les packs mixtes incluent naturellement plusieurs variétés de concentrés.<br/>"
        "• <b>Suivi facilité</b> — le suivi des ventes par pack permet de mesurer l'adoption et l'efficacité.<br/><br/>"
        "<b>Mise en œuvre :</b> les packs sont proposés systématiquement par les commerciaux. "
        "Le client peut ajuster les quantités à la hausse (plus de concentrés = ratio plus favorable) "
        "mais <b>pas à la baisse</b> en dessous du ratio 2,5:1. "
        "Pour les commandes ≥ 10 t de soja, seul un pack est accepté (pas de commande libre).",
        BODY_BOLD
    ))

    story.append(Spacer(1, 0.4*cm))

    # ---------- Final footer ----------
    story.append(Spacer(1, 1*cm))
    story.append(HRFlowable(width="100%", thickness=0.8, color=NAVY, spaceBefore=10, spaceAfter=10))
    story.append(Paragraph(
        "<i>Document rédigé par William Francis Fohom, Data Analyst, Administrateur National de Ventes — "
        "BELGOCAM SA — Juillet 2026.</i>",
        CAPTION
    ))
    story.append(Paragraph(
        "<i>Document confidentiel — Diffusion restreinte au comité de direction et aux responsables d'agences.</i>",
        CAPTION
    ))

    return story


# ===== DOCUMENT TEMPLATE =====
class MyDocTemplate(BaseDocTemplate):
    def __init__(self, filename, **kwargs):
        BaseDocTemplate.__init__(self, filename, **kwargs)
        # Cover page template (no header/footer)
        cover_frame = Frame(0, 0, PAGE_W, PAGE_H, leftPadding=0, rightPadding=0,
                            topPadding=0, bottomPadding=0, id='cover')
        cover_template = PageTemplate(id='Cover', frames=[cover_frame], onPage=draw_cover)
        # Body page template (with header/footer)
        body_frame = Frame(MARGIN_L, MARGIN_B, CONTENT_W,
                           PAGE_H - MARGIN_T - MARGIN_B, id='body')
        body_template = PageTemplate(id='Body', frames=[body_frame], onPage=draw_body_page)
        self.addPageTemplates([cover_template, body_template])


# ===== BUILD PDF =====
OUT = '/home/z/my-project/download/analyse_bundle_soja_concentres.pdf'

doc = MyDocTemplate(OUT, pagesize=A4)
doc.title = "BELGOCAM SA — Analyse Bundle Soja-Concentrés (Juillet 2026)"
doc.author = "William Francis Fohom — Data Analyst, Administrateur National de Ventes"
doc.subject = "Impact de la hausse du prix du soja sur le comportement d'achat client (bundle soja-concentrés)"
doc.creator = "BELGOCAM SA — Direction Commerciale"

story = build_story()

# Use NextPageTemplate to switch from Cover to Body after page 1
final_story = [NextPageTemplate('Body')] + story

doc.build(final_story)

# ===== REPORT =====
print(f"PDF generated: {OUT}")
print(f"File size: {os.path.getsize(OUT):,} bytes")

# Try to get page count
try:
    from PyPDF2 import PdfReader
    reader = PdfReader(OUT)
    print(f"Page count: {len(reader.pages)}")
except Exception:
    try:
        import subprocess
        result = subprocess.run(['pdfinfo', OUT], capture_output=True, text=True)
        for line in result.stdout.split('\n'):
            if line.startswith('Pages:'):
                print(f"Page count: {line.split(':')[1].strip()}")
                break
    except Exception:
        pass
