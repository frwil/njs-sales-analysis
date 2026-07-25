"""
BELGOCAM SA - Analyse Zéro Achat (Standalone PDF)
Suivi continu Janvier-Decembre 2026

PDF genere via ReportLab.
Reutilise les memes polices, couleurs et fonctions utilitaires que build_pdf_report_v3.py.

Auteur/redacteur: William Francis Fohom, Data Analyst, Administrateur National de Ventes
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

# ===== FONTS =====
FONT_DIR = '/usr/share/fonts'
pdfmetrics.registerFont(TTFont('NotoSerifSC', f'{FONT_DIR}/truetype/noto-serif-sc/NotoSerifSC-Regular.ttf'))
pdfmetrics.registerFont(TTFont('NotoSerifSC-Bold', f'{FONT_DIR}/truetype/noto-serif-sc/NotoSerifSC-Bold.ttf'))
pdfmetrics.registerFont(TTFont('NotoSerifSC-Light', f'{FONT_DIR}/truetype/noto-serif-sc/NotoSerifSC-Light.ttf'))
registerFontFamily('NotoSerifSC', normal='NotoSerifSC', bold='NotoSerifSC-Bold')

# ===== COLORS (BELGOCAM corporate: navy + gold) — same as build_pdf_report_v3.py =====
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

# Title styles
H1 = ParagraphStyle('H1', parent=styles['Heading1'], fontName='NotoSerifSC-Bold', fontSize=20,
                    textColor=NAVY, spaceAfter=14, spaceBefore=10, alignment=TA_LEFT, leading=26)
H2 = ParagraphStyle('H2', parent=styles['Heading2'], fontName='NotoSerifSC-Bold', fontSize=14,
                    textColor=NAVY, spaceAfter=10, spaceBefore=14, alignment=TA_LEFT, leading=18,
                    keepWithNext=1)
H3 = ParagraphStyle('H3', parent=styles['Heading3'], fontName='NotoSerifSC-Bold', fontSize=12,
                    textColor=NAVY_LIGHT, spaceAfter=8, spaceBefore=10, alignment=TA_LEFT, leading=15,
                    keepWithNext=1)

# Body styles
BODY = ParagraphStyle('Body', parent=styles['BodyText'], fontName='NotoSerifSC', fontSize=10,
                      textColor=BLACK, alignment=TA_JUSTIFY, leading=14, spaceAfter=8)
BODY_LEFT = ParagraphStyle('BodyLeft', parent=BODY, alignment=TA_LEFT)
BODY_BOLD = ParagraphStyle('BodyBold', parent=BODY, fontName='NotoSerifSC-Bold')
BODY_ITALIC = ParagraphStyle('BodyItalic', parent=BODY, fontName='NotoSerifSC-Light')
SMALL = ParagraphStyle('Small', parent=BODY, fontSize=9, leading=12, textColor=GRAY)
CAPTION = ParagraphStyle('Caption', parent=SMALL, alignment=TA_CENTER, fontName='NotoSerifSC-Light')

# Cell styles
CELL = ParagraphStyle('Cell', fontName='NotoSerifSC', fontSize=9, leading=12, textColor=BLACK, alignment=TA_LEFT)
CELL_CENTER = ParagraphStyle('CellC', parent=CELL, alignment=TA_CENTER)
CELL_RIGHT = ParagraphStyle('CellR', parent=CELL, alignment=TA_RIGHT)
CELL_BOLD = ParagraphStyle('CellB', parent=CELL, fontName='NotoSerifSC-Bold')
CELL_WHITE = ParagraphStyle('CellW', parent=CELL, textColor=WHITE, fontName='NotoSerifSC-Bold', alignment=TA_CENTER)
CELL_WHITE_R = ParagraphStyle('CellWR', parent=CELL_WHITE, alignment=TA_RIGHT)

# Bullet styles
BULLET = ParagraphStyle('Bullet', parent=BODY, leftIndent=18, bulletIndent=6, spaceAfter=4)


# ===== PAGE SETUP =====
PAGE_W, PAGE_H = A4
MARGIN_L = 1.5*cm
MARGIN_R = 1.5*cm
MARGIN_T = 2.2*cm
MARGIN_B = 1.8*cm
CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R


# ===== COVER PAGE FUNCTION (custom for Analyse Zéro Achat) =====
def draw_cover(canv, doc):
    """Draw cover page elements directly on canvas — Analyse Zéro Achat edition."""
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
    canv.drawRightString(PAGE_W - MARGIN_R, PAGE_H - 3.2*cm, "Période analysée : Janvier - Décembre 2026")

    # Main title block (left aligned)
    canv.setFillColor(NAVY)
    canv.setFont('NotoSerifSC-Bold', 30)
    canv.drawString(MARGIN_L, PAGE_H - 9*cm, "Analyse Zéro Achat")

    # Gold separator
    canv.setFillColor(GOLD)
    canv.rect(MARGIN_L, PAGE_H - 9.8*cm, 4*cm, 0.15*cm, fill=1, stroke=0)

    # Subtitle
    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC', 13)
    canv.drawString(MARGIN_L, PAGE_H - 10.8*cm, "Suivi continu Janvier-Décembre 2026")
    canv.setFont('NotoSerifSC-Light', 11)
    canv.drawString(MARGIN_L, PAGE_H - 11.6*cm, "Diagnostic, pertes Q1, transition Q1→Q2 et plan d'action")

    # Key figures box
    box_y = 7.5*cm
    box_h = 5*cm
    canv.setFillColor(GRAY_VLIGHT)
    canv.rect(MARGIN_L, box_y, CONTENT_W, box_h, fill=1, stroke=0)
    # Left gold border
    canv.setFillColor(GOLD)
    canv.rect(MARGIN_L, box_y, 0.2*cm, box_h, fill=1, stroke=0)

    # 4 KPIs side by side
    canv.setFillColor(NAVY)
    canv.setFont('NotoSerifSC-Bold', 20)
    kpi_y = box_y + box_h - 1.5*cm
    col_w = CONTENT_W / 4
    canv.drawCentredString(MARGIN_L + col_w*0.5, kpi_y, "338")
    canv.drawCentredString(MARGIN_L + col_w*1.5, kpi_y, "526")
    canv.drawCentredString(MARGIN_L + col_w*2.5, kpi_y, "14")
    canv.drawCentredString(MARGIN_L + col_w*3.5, kpi_y, "118")

    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC', 8)
    canv.drawCentredString(MARGIN_L + col_w*0.5, kpi_y - 0.7*cm, "Zéro achat global Q1")
    canv.drawCentredString(MARGIN_L + col_w*1.5, kpi_y - 0.7*cm, "Zéro achat concentrés Q1")
    canv.drawCentredString(MARGIN_L + col_w*2.5, kpi_y - 0.7*cm, "Clients 20/80 prioritaires")
    canv.drawCentredString(MARGIN_L + col_w*3.5, kpi_y - 0.7*cm, "Jamais acquis (6 mois)")

    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC-Light', 7)
    canv.drawCentredString(MARGIN_L + col_w*0.5, kpi_y - 1.4*cm, "16 produits ciblés")
    canv.drawCentredString(MARGIN_L + col_w*1.5, kpi_y - 1.4*cm, "12 produits concentrés")
    canv.drawCentredString(MARGIN_L + col_w*2.5, kpi_y - 1.4*cm, "à réactiver en priorité")
    canv.drawCentredString(MARGIN_L + col_w*3.5, kpi_y - 1.4*cm, "clients persistants")

    # Bottom figure: bilan net
    canv.setFillColor(NAVY)
    canv.setFont('NotoSerifSC-Bold', 14)
    canv.drawCentredString(PAGE_W/2, kpi_y - 2.6*cm, "Bilan net Q1→Q2 : +276 M FCFA (ciblé)")
    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC-Light', 8)
    canv.drawCentredString(PAGE_W/2, kpi_y - 3.2*cm, "Perte Q1 estimée : 1 599 t  /  581 M FCFA  —  1 359 clients externes analysés")

    # Author / redactor block at bottom
    auth_y = 3.5*cm
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
    canv.drawRightString(PAGE_W - MARGIN_R, 0.55*cm, "Analyse Zéro Achat — Juillet 2026")

    canv.restoreState()


def draw_body_page(canv, doc):
    """Header/footer for body pages — Analyse Zéro Achat."""
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
    canv.drawString(MARGIN_L + 3*cm, PAGE_H - 1.0*cm, "Analyse Zéro Achat — Suivi continu 2026")
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


# ===== HELPER FUNCTIONS (identical to build_pdf_report_v3.py) =====
def make_table(data, col_widths=None, header_row=True, font_size=9, align='LEFT'):
    """Create a styled table with BELGOCAM colors. Auto-scales to full content width."""
    if col_widths is None:
        n_cols = len(data[0])
        col_widths = [CONTENT_W / n_cols] * n_cols
    else:
        # Auto-scale: if sum of col_widths < CONTENT_W, scale up proportionally
        total = sum(col_widths)
        if total < CONTENT_W * 0.95:
            scale = CONTENT_W / total
            col_widths = [w * scale for w in col_widths]

    # Wrap text cells in Paragraphs for proper wrapping
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
    # Banding
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


def section_divider():
    return HRFlowable(width="100%", thickness=0.5, color=GOLD, spaceBefore=6, spaceAfter=10)


# ===== BUILD CONTENT =====
def build_story():
    story = []

    # ---------- PAGE 1: COVER (drawn by canvas) ----------
    story.append(Spacer(1, PAGE_H - MARGIN_T - MARGIN_B))  # fill the page
    story.append(PageBreak())

    # ---------- PAGE 2: SOMMAIRE ----------
    story.append(Paragraph("Sommaire", H1))
    story.append(section_divider())
    story.append(Spacer(1, 0.3*cm))

    toc_data = [
        ["Section", "Titre", "Page"],
        ["1", "Analyse Zéro Achat S1 2026 (Janvier-Juin)", "3"],
        ["1.1", "Vue d'ensemble", "3"],
        ["1.2", "Top 14 clients 20/80 prioritaires", "4"],
        ["1.3", "Zéro achat par agence", "5"],
        ["1.4", "Zéro achat concentrés (analyse distincte)", "6"],
        ["2", "Pertes Q1 estimées (méthode fréquence)", "7"],
        ["2.1", "Méthode de calcul", "7"],
        ["2.2", "Résultats agrégés", "7"],
        ["2.3", "Top 10 pertes par client", "8"],
        ["2.4", "Lecture de l'impact de la fréquence", "9"],
        ["2.5", "Pertes sur concentrés (296 t, 199 M FCFA)", "9"],
        ["3", "Transition Q1→Q2", "10"],
        ["3.1", "Matrice de transition", "10"],
        ["3.2", "Bilan net Q1→Q2 (+276 M FCFA ciblé)", "11"],
        ["3.3", "Destin des clients fidèles Q1", "12"],
        ["4", "Plan d'action zéro achat", "13"],
        ["4.1", "Tableau récapitulatif des priorités", "13"],
        ["4.2", "Stratégie de récupération Région Ouest (13.3.9)", "14"],
        ["5", "Suivi S2 2026 (Juil-Déc)", "15"],
    ]
    story.append(make_table(toc_data, col_widths=[1.8*cm, 12*cm, 2.5*cm], header_row=True))
    story.append(Spacer(1, 0.6*cm))
    story.append(Paragraph(
        "<i>Ce document est une extraction dédiée de l'analyse complète BELGOCAM SA. Il se concentre sur le diagnostic "
        "zéro achat, l'estimation des pertes Q1, la dynamique de transition Q1→Q2 et le plan d'action de réactivation. "
        "L'analyse porte sur 1 359 clients externes (après exclusion de 35 clients internes) et 16 produits ciblés "
        "(tourteaux de soja + concentrés).</i>",
        BODY_ITALIC
    ))
    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph(
        "<b>Glossaire des abréviations :</b> Q1 = Trimestre 1 (Janvier-Mars 2026) ; Q2 = Trimestre 2 (Avril-Juin 2026) ; "
        "S1 = Semestre 1 (Janvier-Juin 2026) ; S2 = Semestre 2 (Juillet-Décembre 2026) ; CA = Chiffre d'Affaires ; "
        "HT = Hors Taxes ; FCFA = Franc CFA ; 20/80 = segmentation Pareto (20% des clients = 80% du CA).",
        SMALL
    ))

    story.append(PageBreak())

    # =====================================================================
    # SECTION 1 — ANALYSE ZÉRO ACHAT S1 2026
    # =====================================================================
    story.append(Paragraph("1. Analyse Zéro Achat S1 2026 (Janvier-Juin)", H1))
    story.append(section_divider())

    # ----- 1.1 Vue d'ensemble -----
    story.append(Paragraph("1.1 Vue d'ensemble", H2))
    cards3 = [
        kpi_card("Zéro achat global Q1", "338", "clients (16 produits)"),
        kpi_card("Zéro achat concentrés Q1", "526", "clients (12 concentrés)"),
        kpi_card("20/80 en zéro achat Q1", "14", "clients prioritaires"),
        kpi_card("Jamais acquis (6 mois)", "118", "clients persistants"),
    ]
    story.append(kpi_row(cards3))
    story.append(Spacer(1, 0.4*cm))

    story.append(Paragraph(
        "<b>État des lieux détaillé :</b> sur une base de <b>1 359 clients externes</b> analysés, "
        "<b>338 clients sont en zéro achat combiné (soja + concentré)</b> et <b>526 clients sont en zéro achat sur le concentré seul</b>. "
        "L'écart entre ces deux chiffres (526 − 338 = 188) correspond aux clients qui ont acheté du soja mais pas de concentré — "
        "ils représentent la cible prioritaire du bundle soja-concentrés. "
        "Les <b>14 clients du top 20/80 en zéro achat</b> sont répartis géographiquement comme suit : "
        "FAMLA (7), NDOBO (4), DJELENG (2), AHALA (1). "
        "FAMLA concentre donc la moitié des clients prioritaires à réactiver — c'est l'agence où l'enjeu de sauvetage est le plus fort.",
        BODY
    ))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Lecture dynamique S1 :</b> le gap de volume sur le soja a été <b>rattrapé et dépassé</b> grâce à la rupture concurrente "
        "(+46,5% YoY en volume TOURTEAUX). En revanche, le gap sur les concentrés <b>reste non comblé</b> — la dynamique positive du soja "
        "ne s'est pas transmise aux concentrés malgré le potentiel de cross-sell. "
        "<b>⚠️ Attention à la lecture de juin</b> : une légère amélioration est observée sur le nombre de clients zéro achat concentrés en juin, "
        "mais elle est <b>trompeuse</b> — elle s'explique par le retour de certains clients chez BELGOCAM <b>faute de stock concurrent</b>, "
        "non par une fidélisation réelle. Ces clients repartiront chez le concurrent dès que son stock sera reconstitué. "
        "La fenêtre d'opportunité commerciale est donc <b>limitée dans le temps</b> — d'où l'urgence d'agir maintenant.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ----- 1.2 Top 14 clients 20/80 prioritaires -----
    story.append(Paragraph("1.2 Top 14 clients 20/80 prioritaires à réactiver", H2))
    story.append(Paragraph(
        "Ces 14 clients font partie du top 20/80 (★) mais n'ont acheté AUCUN produit ciblé en Q1. "
        "Ils sont à recontacter en priorité absolue pendant la rupture concurrente de soja.",
        BODY
    ))
    top14_data = [
        ["N°", "Client", "Agence", "CA 6 mois (FCFA)", "Perte Q1 (FCFA)", "Priorité"],
        ["1", "SIGHELO SARL", "NDOBO", "23 680 000", "23 680 000", "ÉLEVÉE"],
        ["2", "COMPAGNIE FERMIERE CAMEROUNAISE", "NDOBO", "21 000 000", "21 000 000", "ÉLEVÉE"],
        ["3", "TEIKING JEAN MARIE", "FAMLA", "20 822 000", "19 632 000", "ÉLEVÉE"],
        ["4", "FEUDJIO BACK ARMEL", "FAMLA", "17 460 000", "17 460 000", "ÉLEVÉE"],
        ["5", "KENMEGNE ALAIN", "FAMLA", "20 584 000", "16 600 000", "ÉLEVÉE"],
        ["6", "STE SATI SARL", "NDOBO", "13 600 000", "13 600 000", "ÉLEVÉE"],
        ["7", "FOTSO VINCENT", "FAMLA", "17 589 000", "13 464 000", "ÉLEVÉE"],
        ["8", "DJOUSSI AURORE FLORINDA", "FAMLA", "13 280 000", "13 280 000", "ÉLEVÉE"],
        ["9", "SADO", "DJELENG", "12 320 000", "12 320 000", "ÉLEVÉE"],
        ["10", "ETS MEKATALEO", "DJELENG", "12 424 250", "11 975 000", "ÉLEVÉE"],
        ["11", "MAKOUGANG KEUMBOU HERMINE", "FAMLA", "12 377 600", "11 922 600", "ÉLEVÉE"],
        ["12", "BAHO", "FAMLA", "14 234 500", "10 920 000", "ÉLEVÉE"],
        ["13", "MAKUETE Epse MANFOUO EDITH", "NDOBO", "15 400 000", "—", "FAIBLE"],
        ["14", "TChouabe siakam therese", "AHALA", "13 183 000", "—", "FAIBLE"],
    ]
    story.append(make_table(top14_data, col_widths=[0.8*cm, 5.5*cm, 2.5*cm, 3.2*cm, 3.2*cm, 2.6*cm], font_size=8))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Concentration géographique :</b> FAMLA (7 clients), NDOBO (4 clients), DJELENG (2 clients), AHALA (1). "
        "Perte Q1 cumulée des 12 premiers = 173 M FCFA. Action commerciale territoriale à organiser agence par agence.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ----- 1.3 Zéro achat par agence -----
    story.append(Paragraph("1.3 Zéro achat par agence", H2))
    story.append(Paragraph(
        "Répartition des clients zéro achat Q1 (ciblé) par agence principale, avec décomposition 20/80 vs autres.",
        BODY
    ))
    ag_za_data = [
        ["Agence", "Total clients", "Actifs Q1", "Zéro achat 20/80", "Zéro achat autres", "Total zéro achat", "% zéro achat"],
        ["FAMLA", "293", "237", "7", "49", "56", "19,1%"],
        ["MESSASSI", "167", "111", "0", "56", "56", "33,5%"],
        ["NDOBO", "154", "109", "4", "41", "45", "29,2%"],
        ["DJELENG", "107", "80", "2", "25", "27", "25,2%"],
        ["NGAOUNDERE", "91", "60", "0", "31", "31", "34,1%"],
        ["Autres agences", "647", "525", "1", "121", "122", "18,9%"],
        ["TOTAL", "1 359", "1 022", "14", "324", "338", "24,9%"],
    ]
    story.append(make_table(ag_za_data, col_widths=[3.2*cm, 1.8*cm, 1.6*cm, 2.2*cm, 2*cm, 2.2*cm, 1.6*cm]))
    story.append(Spacer(1, 0.3*cm))
    # Chart: Zero achat par agence
    chart_path = '/home/z/my-project/scripts/pdf_charts/chart4_zero_agence.png'
    if os.path.exists(chart_path):
        img = Image(chart_path, width=15*cm, height=8.3*cm)
        story.append(img)
        story.append(Paragraph("Figure 1 — Répartition zéro achat Q1 (ciblé) par agence", CAPTION))
        story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Constat :</b> MESSASSI et NGAOUNDERE affichent un taux de zéro achat critique (≥33%), alors qu'aucun client 20/80 "
        "n'est en zéro achat dans ces agences. Cela suggère un problème de couverture commerciale sur les clients moyens — "
        "potentiellement un sous-dimensionnement de l'effectif commercial ou un manque de visibilité locale. "
        "FAMLA concentre la moitié des clients 20/80 en zéro achat (7 sur 14), ce qui en fait l'agence prioritaire pour l'action de sauvetage.",
        BODY
    ))

    story.append(PageBreak())

    # ----- 1.4 Zéro achat concentrés -----
    story.append(Paragraph("1.4 Zéro achat concentrés (analyse distincte)", H2))
    story.append(Paragraph(
        "526 clients n'ont acheté aucun concentré en Q1 (vs 338 sur le ciblé global). La différence (188 clients) correspond "
        "aux clients qui ont acheté du soja en Q1 mais pas de concentrés — ce sont les cibles idéales pour une action de "
        "cross-sell vers les concentrés.",
        BODY
    ))
    za_c_data = [
        ["Catégorie", "Nb clients", "Dont 20/80", "Action"],
        ["Zéro achat ciblé Q1 (16 produits)", "338", "14", "Réactivation complète (soja + concentrés)"],
        ["Zéro achat concentrés Q1 (12 produits)", "526", "41", "Cross-sell concentrés (soja déjà acheté pour 188)"],
        ["Jamais acheté de concentrés en 6 mois", "≈ 380", "—", "Prospection longue — offre découverte"],
    ]
    story.append(make_table(za_c_data, col_widths=[6*cm, 2.5*cm, 2.5*cm, 6.5*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Insight stratégique :</b> les 188 clients qui achètent du soja mais pas de concentrés sont la cible #1 du cross-sell concentrés. "
        "Ils connaissent déjà BELGOCAM, ils ont une activité d'élevage, mais ils ne consomment pas le produit le plus marginal. "
        "Une offre bundle \"soja + concentrés préférentiels\" pourrait convertir une partie significative de ces 188 clients.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # =====================================================================
    # SECTION 2 — PERTES Q1 ESTIMÉES
    # =====================================================================
    story.append(Paragraph("2. Pertes Q1 estimées", H1))
    story.append(section_divider())

    # ----- 2.1 Méthode de calcul -----
    story.append(Paragraph("2.1 Méthode de calcul (fréquence d'achat)", H2))
    story.append(Paragraph(
        "La perte Q1 estimée pour chaque client \"zéro achat\" est calculée à partir de son comportement Q2, en tenant compte de sa fréquence d'achat :",
        BODY
    ))
    story.append(Paragraph(
        "<b>Perte Q1 = Σ(moyenne mensuelle par produit) × Fréquence × 3 mois</b>",
        BODY_BOLD
    ))
    story.append(Paragraph(
        "où :<br/>"
        "• <b>Moyenne mensuelle par produit P</b> = (volume total Q2 du produit P) / (nombre de mois Q2 où P a été acheté)<br/>"
        "• <b>Fréquence</b> = (nombre de mois Q2 avec achat de la catégorie ciblée) / 3<br/>"
        "• <b>3 mois</b> = nombre de mois \"perdus\" en Q1 (Jan-Mar)",
        BODY
    ))
    story.append(Paragraph(
        "Cette méthode permet de distinguer un client qui aurait acheté tous les mois (fréquence = 1, perte élevée) "
        "d'un client occasionnel qui n'aurait acheté qu'un mois sur trois (fréquence = 0,33, perte plus faible).",
        BODY
    ))

    # ----- 2.2 Résultats agrégés -----
    story.append(Paragraph("2.2 Résultats agrégés", H2))
    story.append(Paragraph(
        "Synthèse des pertes Q1 estimées par segment client (base 1 359 clients externes analysés) :",
        BODY
    ))
    pertes_agg_data = [
        ["Catégorie", "Clients", "Volume concerné (t)", "CA concerné (FCFA)", "Lecture"],
        ["Récupérables (actifs Q1+Q2)", "850", "≈ 16 721", "≈ 7,02 Md", "Cœur de clientèle stable — à fidéliser"],
        ["Zéro achat (ciblé Q1)", "338", "1 598,66", "580 819 762", "Perte Q1 à récupérer via réactivation"],
        ["Churned (actifs Q1 → zéro Q2)", "192", "758,01", "273 296 995", "Perte sèche — reconquête prioritaire"],
        ["TOTAL", "1 380", "≈ 19 078", "≈ 7,87 Md", "Base de pilotage du plan d'action"],
    ]
    story.append(make_table(pertes_agg_data, col_widths=[4.5*cm, 1.6*cm, 2.6*cm, 3.2*cm, 4.9*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>581 millions FCFA</b> de CA potentiel perdu en Q1 sur les seuls clients zéro achat global. "
        "C'est l'équivalent de 80% du CA total 6 mois de ces mêmes clients, ce qui démontre l'ampleur de l'opportunité commerciale. "
        "À cela s'ajoutent 273 M FCFA perdus sur les 192 clients churned Q1→Q2 — soit un enjeu total de <b>≈ 854 M FCFA</b>.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ----- 2.3 Top 10 pertes par client -----
    story.append(Paragraph("2.3 Top 10 pertes par client (ciblé global)", H2))
    top10_data = [
        ["N°", "Client", "Agence", "Région", "Fréquence Q2", "Perte Vol. (t)", "Perte CA (FCFA)"],
        ["1", "SIGHELO SARL", "NDOBO", "Littoral", "0,33 (1 mois)", "74,00", "23 680 000"],
        ["2", "TEIKING JEAN MARIE", "FAMLA", "Ouest", "0,67 (2 mois)", "61,00", "19 632 000"],
        ["3", "COMPAGNIE FERMIERE CAM.", "NDOBO", "Littoral", "0,33 (1 mois)", "60,00", "21 000 000"],
        ["4", "FEUDJIO BACK ARMEL", "FAMLA", "Ouest", "0,67 (2 mois)", "55,00", "17 460 000"],
        ["5", "KENMEGNE ALAIN", "FAMLA", "Ouest", "0,33 (1 mois)", "50,00", "16 600 000"],
        ["6", "FOTSO VINCENT", "FAMLA", "Ouest", "1,00 (3 mois)", "42,00", "13 464 000"],
        ["7", "DJOUSSI AURORE FLORINDA", "FAMLA", "Ouest", "0,33 (1 mois)", "40,00", "13 280 000"],
        ["8", "STE SATI SARL", "NDOBO", "Littoral", "0,33 (1 mois)", "40,00", "13 600 000"],
        ["9", "STE IPACAM & FILS", "MESSASSI", "Centre", "1,00 (3 mois)", "36,75", "14 754 000"],
        ["10", "BAHO", "FAMLA", "Ouest", "0,33 (1 mois)", "35,00", "10 920 000"],
    ]
    story.append(make_table(top10_data, col_widths=[0.8*cm, 4*cm, 1.5*cm, 1.5*cm, 2.2*cm, 2*cm, 3*cm], font_size=7.5))
    story.append(Spacer(1, 0.3*cm))
    # Chart: Top 10 pertes
    chart_path = '/home/z/my-project/scripts/pdf_charts/chart3_top10_pertes.png'
    if os.path.exists(chart_path):
        img = Image(chart_path, width=15*cm, height=8.3*cm)
        story.append(img)
        story.append(Paragraph("Figure 2 — Top 10 pertes Q1 (clients zéro achat global, en millions FCFA)", CAPTION))

    story.append(PageBreak())

    # ----- 2.4 Lecture impact fréquence -----
    story.append(Paragraph("2.4 Lecture de l'impact de la fréquence", H2))
    story.append(Paragraph(
        "La comparaison entre SIGHELO SARL et FOTSO VINCENT illustre bien l'effet de la fréquence sur le calcul de perte :",
        BODY
    ))
    freq_compare = [
        ["Client", "Moy. mens. Q2", "Fréquence Q2", "Perte Q1 (t)", "Lecture"],
        ["SIGHELO SARL", "74,00 t", "0,33 (1 mois/3)", "74,00 t", "Gros volume ponctuel — 1 mois aurait suffi"],
        ["FOTSO VINCENT", "14,00 t", "1,00 (3 mois/3)", "42,00 t", "Petit volume régulier — 3 mois pleins auraient été faits"],
        ["TEIKING JEAN MARIE", "30,50 t", "0,67 (2 mois/3)", "61,00 t", "Volume moyen régulier — 2 mois sur 3"],
    ]
    story.append(make_table(freq_compare, col_widths=[4*cm, 2.8*cm, 3*cm, 2.5*cm, 4.7*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Insight :</b> un client à haute fréquence d'achat (FOTSO, 3/3 mois) est plus \"récupérable\" à court terme "
        "qu'un client à faible fréquence mais gros volume (SIGHELO, 1/3 mois). "
        "Les actions de réactivation doivent donc prioriser en premier lieu les clients à forte fréquence Q2, "
        "qui démontrent un besoin d'achat récurrent.",
        BODY
    ))

    # ----- 2.5 Pertes sur concentrés -----
    story.append(Paragraph("2.5 Pertes sur concentrés (296 t, 199 M FCFA)", H2))
    story.append(Paragraph(
        "Sur les 526 clients zéro achat concentrés en Q1, la perte estimée est de <b>296 tonnes et 199 millions FCFA</b>. "
        "C'est moins que sur le ciblé global (car certains de ces clients ont quand même acheté du soja), "
        "mais c'est plus préoccupant stratégiquement car les concentrés sont le cœur de marge.",
        BODY
    ))
    story.append(Paragraph(
        "<b>Top 5 pertes sur concentrés :</b>",
        BODY_BOLD
    ))
    top5_conc = [
        ["Client", "Fréquence Q2", "Perte Volume (t)", "Perte CA (FCFA)"],
        ["SOCIETE COOPERATIVE SIMPLIFIEE DATCHIO", "0,67", "15,80", "10 114 000"],
        ["PROVENDERIE ELI EMONBA 2", "0,67", "10,00", "6 700 000"],
        ["PVD LE GROSSISTE SARL", "1,00", "9,00", "5 940 000"],
        ["STE IPACAM & FILS", "1,00", "9,00", "5 985 000"],
        ["MOTSOU DJEUMAKOU VICTOR", "1,00", "8,40", "6 048 000"],
    ]
    story.append(make_table(top5_conc, col_widths=[6*cm, 3*cm, 3.5*cm, 4*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "Ces 5 clients représentent à eux seuls 52 tonnes de perte concentrés, soit ~17% de la perte totale concentrés. "
        "Les clients \"PVD LE GROSSISTE\" et \"STE IPACAM\" sont à fréquence maximale (1,00) et donc prioritaires pour réactivation.",
        BODY
    ))

    story.append(PageBreak())

    # =====================================================================
    # SECTION 3 — TRANSITION Q1→Q2
    # =====================================================================
    story.append(Paragraph("3. Transition Q1→Q2", H1))
    story.append(section_divider())

    # ----- 3.1 Matrice de transition -----
    story.append(Paragraph("3.1 Matrice de transition Q1→Q2 (16 produits ciblés)", H2))
    matrice = [
        ["", "Q2 : Zéro achat", "Q2 : Active", "Total Q1"],
        ["Q1 : Zéro achat", "118 (persistants)", "234 (réactivés) ★", "352"],
        ["Q1 : Active", "192 (churned) ⚠", "850 (retenus)", "1 042"],
        ["Total Q2", "310", "1 084", "1 359"],
    ]
    story.append(make_table(matrice, col_widths=[3.5*cm, 4*cm, 4.5*cm, 3*cm]))
    story.append(Spacer(1, 0.3*cm))
    # Chart: Segments de transition
    chart_path = '/home/z/my-project/scripts/pdf_charts/chart2_segments.png'
    if os.path.exists(chart_path):
        img = Image(chart_path, width=14*cm, height=7.9*cm)
        story.append(img)
        story.append(Paragraph("Figure 3 — Segments de transition Q1 → Q2 (16 produits ciblés)", CAPTION))
        story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Quatre segments :</b><br/>"
        "• <b>Persistants (118)</b> : clients qui n'ont jamais acheté de ciblé sur 6 mois — opportunité non exploitée, à prospecter.<br/>"
        "• <b>Réactivés (234)</b> : GAIN — clients sans achat Q1 qui ont acheté en Q2. CA généré : 549,7 M FCFA.<br/>"
        "• <b>Retenus (850)</b> : STABLES — cœur de clientèle, achats Q1 et Q2.<br/>"
        "• <b>Churned (192)</b> : PERTE — clients actifs Q1 devenus inactifs Q2. CA perdu : 273,3 M FCFA.",
        BODY
    ))

    story.append(PageBreak())

    # ----- 3.2 Bilan net Q1→Q2 -----
    story.append(Paragraph("3.2 Bilan net Q1→Q2 (+276 M FCFA ciblé)", H2))
    bilan_data = [
        ["", "Clients", "Volume (t)", "CA (FCFA)"],
        ["GAIN — Réactivés (Q1 Zéro achat → Q2 Active)", "234", "+1 524,74", "+549 689 847"],
        ["PERTE — Churned (Q1 Active → Q2 Zéro achat)", "192", "-758,01", "-273 296 995"],
        ["BILAN NET", "+42", "+766,73", "+276 392 852"],
    ]
    story.append(make_table(bilan_data, col_widths=[7*cm, 2.5*cm, 3*cm, 3.5*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Bilan positif :</b> BELGOCAM a gagné <b>276 millions FCFA nets</b> sur les produits ciblés entre Q1 et Q2. "
        "Le gain de réactivation (549 M) dépasse largement la perte par churn (273 M). "
        "Cependant, ce bilan est principalement porté par le soja — le bilan concentrés seul est beaucoup plus fragile (+52 M FCFA seulement).",
        BODY_BOLD
    ))
    story.append(Spacer(1, 0.2*cm))
    # Chart: Bilan net
    chart_path = '/home/z/my-project/scripts/pdf_charts/chart5_bilan_net.png'
    if os.path.exists(chart_path):
        img = Image(chart_path, width=14*cm, height=7.9*cm)
        story.append(img)
        story.append(Paragraph("Figure 4 — Bilan net Q1 → Q2 : Ciblé vs Concentrés (millions FCFA)", CAPTION))

    story.append(PageBreak())

    # ----- 3.3 Destin des clients fidèles Q1 -----
    story.append(Paragraph("3.3 Destin des clients fidèles Q1 (récupérés, déclin, churned)", H2))
    story.append(Paragraph(
        "Sur les 527 clients fidèles en Q1 (3 mois d'achat ciblé), voici leur destin en Q2 :",
        BODY
    ))
    fidel_data = [
        ["Statut Q2", "Nb clients", "% des Q1 Fidèles", "Commentaire"],
        ["Q2 Fidèle (3 mois) — récupérés", "420", "79,7%", "Maintenus — cœur de clientèle stable"],
        ["Q2 Semi-fidèle (1-2 mois) — déclin", "97", "18,4%", "Déclin de fréquence — à surveiller (risque churn)"],
        ["Q2 Zéro achat (churned)", "10", "1,9%", "Perte sèche — anciens fidèles devenus inactifs"],
    ]
    story.append(make_table(fidel_data, col_widths=[5*cm, 2.5*cm, 3*cm, 6*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "Les <b>97 fidèles en déclin (18,4%)</b> sont un signal d'alerte pré-churn à traiter en priorité. "
        "Une visite physique par le commercial dédié permettrait de comprendre la cause du déclin "
        "(insatisfaction, concurrence, changement d'activité) avant que le client ne churn totalement.",
        BODY
    ))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Lecture stratégique :</b> le taux de churn des fidèles Q1 est de 1,9% sur le ciblé global, "
        "mais il grimpe à 3,1% sur les concentrés seuls — les concentrés sont donc plus sensibles à l'attrition. "
        "Le suivi rapproché des 97 fidèles en déclin est la meilleure prévention contre le churn structurel.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # =====================================================================
    # SECTION 4 — PLAN D'ACTION ZÉRO ACHAT
    # =====================================================================
    story.append(Paragraph("4. Plan d'action zéro achat", H1))
    story.append(section_divider())

    # ----- 4.1 Tableau récapitulatif des priorités -----
    story.append(Paragraph("4.1 Tableau récapitulatif des priorités", H2))
    story.append(Paragraph(
        "Plan d'action priorisé par niveau d'urgence — couvre la réactivation des 14 clients 20/80, le recontact des churned, "
        "le bundle soja-concentrés et la prospection des persistants.",
        BODY
    ))
    rec_data = [
        ["Priorité", "Axe d'action", "Nb clients", "Délai", "Pilote"],
        ["🔴 CRITIQUE", "Sauvetage des 14 clients 20/80 en zéro achat Q1 (ciblé + concentrés)", "14", "30 j", "Directeur Commercial + Responsables agences"],
        ["🔴 CRITIQUE", "Bundle soja-concentrés obligatoire : 3 sacs soja pour 1 sac concentré sur chaque commande", "Tous", "Immédiat", "Direction Commerciale"],
        ["🟠 ÉLEVÉE", "Recontact des 192 clients churned (actifs Q1 → Zéro achat Q2) par SAV + commercial", "192", "15 j", "SAV + Commerciaux terrain"],
        ["🟠 ÉLEVÉE", "Sauvetage des 97 anciens fidèles Q1 en déclin Q2 — visite physique + suivi SAV", "97", "30 j", "Commerciaux + SAV"],
        ["🟠 ÉLEVÉE", "Suivi SAV systématique des clients concentrés à risque (appels de satisfaction mensuels)", "850+", "Continu", "Service Après-Vente"],
        ["🟡 MOYENNE", "Prospection des 118 clients jamais acquis (persistants zéro achat)", "118", "60 j", "Marketing + Commerciaux"],
        ["🟡 MOYENNE", "Capitalisation sur les 234 clients réactivés Q2 (étude qualitative)", "234", "45 j", "Marketing + Direction commerciale"],
        ["🟡 MOYENNE", "Motivation de l'équipe commerciale : primes liées aux ventes de concentrés (objectif individuel)", "Tous", "30 j", "Direction + RH"],
        ["🟢 STRUCTURANTE", "Analyse du déclin volume concentrés (segment retenu -248 t) + benchmark tarifaire", "672", "90 j", "Direction Produit + Direction Commerciale"],
    ]
    story.append(make_table(rec_data, col_widths=[2.5*cm, 6.5*cm, 1.2*cm, 1.3*cm, 5.5*cm], font_size=7))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph(
        "<b>Synthèse du plan d'action :</b>",
        BODY_BOLD
    ))
    synth_pa = [
        ["Indicateur", "Valeur"],
        ["Enjeu total identifié (pertes Q1 + risques churn)", "≈ 854 M FCFA"],
        ["Gain potentiel (si réactivation complète des 20/80 zéro achat)", "≈ 175 M FCFA"],
        ["Nb total de clients concernés par une action prioritaire", "≈ 524 clients + 850+ suivis SAV"],
        ["Bundle soja-concentrés (ratio 3:1)", "Déploiement immédiat — toutes agences"],
        ["Suivi SAV clients concentrés", "2 ETP dédiés — appels mensuels 850+ clients"],
        ["Motivation équipe commerciale", "Primes sur concentrés + formation bundles (30 j)"],
        ["Horizon de mise en œuvre", "30-90 jours (priorités CRITIQUE et ÉLEVÉE sous 30 jours)"],
    ]
    story.append(make_table(synth_pa, col_widths=[8*cm, 9*cm]))

    story.append(PageBreak())

    # ----- 4.2 Stratégie de récupération Région Ouest (13.3.9) -----
    story.append(Paragraph("4.2 Stratégie de récupération — Région Ouest (13.3.9)", H2))
    story.append(Paragraph(
        "<b>Constat :</b> BELGOCAM dispose actuellement d'un avantage sur le soja dans la région Ouest (rupture concurrente). "
        "Cet avantage doit être exploité immédiatement, mais avec méthode.",
        BODY_BOLD
    ))
    story.append(Paragraph(
        "<b>Principes d'action :</b><br/>"
        "• <b>Conditionner l'accès soja à l'achat de concentré</b> — transformer l'avantage soja en levier de cross-sell.<br/>"
        "• <b>Benchmarking prix concurrent indispensable</b> avant toute approche des éleveurs — "
        "ne pas se présenter sans connaître le positionnement tarifaire du concurrent.<br/>"
        "• <b>Déploiement séquentiel par zone, pas simultané</b> — se donner un délai par agence pour ajuster en fonction des retours terrain.<br/>"
        "• <b>Horizon maximal : 3 mois</b> pour implémenter les actions (pas 6 mois) — "
        "objectif : récupérer au moins 50% des zones cibles d'ici fin du semestre.",
        BODY
    ))
    story.append(Paragraph(
        "<b>⚠️ Risque majeur :</b> l'avantage concurrentiel n'est pas indéfini. Le concurrent peut reconstituer son stock à tout moment. "
        "La fenêtre d'action est <b>limitée et imprévisible</b> — d'où l'urgence du déploiement sous 3 mois.",
        BODY_BOLD
    ))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Scénarios de récupération sur les 14 clients top 20/80 :</b>",
        BODY_BOLD
    ))
    scenario_data = [
        ["Scénario", "Clients récupérés (sur 14)", "% récupération", "CA récupéré estimé", "Lecture"],
        ["🔴 Pessimiste", "4", "~10%", "≈ 34 M FCFA", "Rupture concurrente courte (1-2 mois)"],
        ["🟡 Réaliste", "8", "~30%", "≈ 68 M FCFA", "Rupture modérée (3-4 mois)"],
        ["🟢 Optimiste", "12", "~60%", "≈ 103 M FCFA", "Rupture prolongée (6+ mois)"],
    ]
    story.append(make_table(scenario_data, col_widths=[3*cm, 2.8*cm, 2.2*cm, 3*cm, 6*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Action :</b> déployer la stratégie Ouest en séquentiel : <b>FAMLA en semaine 1</b> (7 clients 20/80 à réactiver), "
        "<b>DJELENG en semaine 2</b> (2 clients), puis extension vers les autres régions avec les apprentissages terrain.",
        BODY
    ))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Calendrier de déploiement séquentiel (3 mois) :</b>",
        BODY_BOLD
    ))
    cal_ouest = [
        ["Période", "Zone / Agence", "Action principale", "Cible"],
        ["Semaine 1", "FAMLA", "RDV direct commercial — 7 clients 20/80 zéro achat", "7 clients 20/80"],
        ["Semaine 2", "DJELENG", "Visite physique + offre bundle soja-concentrés", "2 clients 20/80"],
        ["Semaines 3-4", "NDOBO", "Recontact 4 clients 20/80 + 41 zéro achat autres", "4 clients 20/80 + 41"],
        ["Mois 2", "AHALA + autres agences Ouest", "Extension avec apprentissages terrain", "1 client 20/80 + 121 autres"],
        ["Mois 3", "Toutes agences Ouest", "Consolidation + verrouillage contrats 6 mois", "Récupération ≥ 50% des zones cibles"],
    ]
    story.append(make_table(cal_ouest, col_widths=[2.2*cm, 3.5*cm, 7.5*cm, 4*cm], font_size=8))

    story.append(PageBreak())

    # =====================================================================
    # SECTION 5 — SUIVI S2 2026 (JUIL-DÉC)
    # =====================================================================
    story.append(Paragraph("5. Suivi S2 2026 (Juil-Déc)", H1))
    story.append(section_divider())

    story.append(Paragraph(
        "Cette section suit l'évolution mensuelle des indicateurs zéro achat et l'efficacité du plan d'action. "
        "À chaque mise à jour des données, les valeurs réelles sont comparées à la référence S1 2026 pour mesurer "
        "l'impact des actions menées et ajuster le tir si nécessaire.",
        BODY
    ))

    # ===== 5.1 Premier point de suivi — Juillet 2026 (au 24/07) =====
    story.append(Paragraph("5.1 Premier point de suivi — Juillet 2026 (au 24/07)", H2))
    story.append(Paragraph(
        "<b>Données au 24 juillet 2026</b> — temps écoulé : 24 jours sur 31 (78% du mois, 21j ouvrables lun-sam sur 27). "
        "Les comparaisons avec S1 2026 sont ajustées au prorata du temps écoulé pour assurer une lecture juste. "
        "La projection juillet complet = réel 24 jours × (31/24).",
        BODY_BOLD
    ))

    # Tableau d'évolution des ventes par catégorie
    story.append(Paragraph("5.1.1 Évolution des ventes par catégorie — S1 vs Juillet", H3))
    story.append(Paragraph(
        "Comparaison du volume réalisé et du CA par catégorie entre la moyenne mensuelle S1 2026 et juillet 2026 (24 jours, projection 31 jours). Le Maïs est séparé des autres ingrédients (produit opportuniste).",
        BODY
    ))

    evol_ventes_data = [
        ["Catégorie", "Moy. mensuelle S1 (t)", "Juil 24j (t)", "Projection Juil (t)", "vs S1 (%)", "Obj Juil (t)", "% ajusté*"],
        ["TOURTEAUX", "4 844", "5 203", "6 689", "+38%", "3 834", "174%"],
        ["CONCENTRES", "1 513", "1 240", "1 594", "+5%", "1 802", "88%"],
        ["ALIMENT COMPLET", "102", "86", "110", "+8%", "74", "149%"],
        ["INGREDIENTS (hors Maïs)", "84", "59", "76", "-10%", "86", "88%"],
        ["MAÏS", "599", "0", "0", "-100%", "—", "—"],
        ["PREMIX", "12", "9", "12", "0%", "9", "128%"],
        ["TOTAL", "7 154", "6 597", "8 481", "+19%", "5 806", "146%"],
    ]
    story.append(make_table(evol_ventes_data, col_widths=[3*cm, 2.5*cm, 2.2*cm, 2.5*cm, 1.8*cm, 2*cm, 1.8*cm], font_size=7.5))
    story.append(Paragraph(
        "<i>* % ajusté = réel 24j / (objectif mois × 78%) × 100. Un % ajusté ≥ 100% signifie que le rythme est conforme à l'objectif au prorata du temps écoulé.</i>",
        ParagraphStyle('fn', parent=SMALL, fontName='NotoSerifSC-Light', fontSize=7, textColor=GRAY, spaceBefore=2)
    ))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Lecture juillet 2026 (au 24/07) :</b><br/>"
        "• <b>Volume global à 146% ajusté</b> — le rythme de juillet reste supérieur à l'objectif au prorata. Très bon démarrage S2.<br/>"
        "• <b>TOURTEAUX à 174% ajusté</b> — la dynamique de rupture concurrente se maintient vigoureusement. La projection (6 689 t) reste très supérieure à la moyenne S1 (4 844 t).<br/>"
        "• <b>CONCENTRES à 88% ajusté</b> — toujours sous l'objectif au prorata. La projection juillet (1 594 t) est légèrement supérieure à la moyenne S1 (1 513 t) — <b>amélioration timide vs début juillet</b>, mais le plan d'action doit accélérer.<br/>"
        "• <b>PREMIX à 128% ajusté</b> — performance solide, conforme à l'objectif.<br/>"
        "• <b>MAÏS à 0 t</b> — aucun achat en juillet (produit opportuniste, vendu uniquement sur disponibilité et demande).",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ===== 5.1.2 Évolution du portefeuille client — Analyse par habitudes d'achat =====
    story.append(Paragraph("5.1.2 Évolution du portefeuille client — Analyse par habitudes d'achat", H3))
    story.append(Paragraph(
        "Nous sommes au 24 juillet — il est normal que certains clients n'aient pas encore commandé. "
        "Une analyse pertinente doit croiser les habitudes d'achat de chaque client (jour habituel de commande en S1) "
        "avec son activité juillet. Un client qui achète habituellement en début de mois et qui n'a pas commandé au 24/07 "
        "doit attirer notre attention — un client qui achète en fin de mois n'est pas encore inquiétant.",
        BODY
    ))

    habitudes_data = [
        ["Catégorie", "Définition", "Nb clients", "CA S1 (M)", "Action"],
        ["✅ Maintenus", "Clients S1 ayant déjà acheté en juillet", "775", "14 560", "Aucune — base fidèle"],
        ["🔴 À risque élevé", "Jour d'achat habituel ≤ 21, pas d'achat juillet", "497", "1 851", "Recontact immédiat"],
        ["🟡 À surveiller", "Jour d'achat habituel 22-25, pas d'achat juillet", "56", "116", "À recontacter dès le 26/07"],
        ["🟢 Pas encore inquiétant", "Jour d'achat habituel > 25, pas d'achat juillet", "41", "42", "Attendre fin juillet"],
        ["TOTAL", "", "1 369", "16 569", ""],
    ]
    story.append(make_table(habitudes_data, col_widths=[3*cm, 5*cm, 1.8*cm, 2*cm, 4*cm], font_size=7.5))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Lecture par habitudes d'achat (au 24/07) :</b><br/>"
        "• <b>775 clients maintenus (57%)</b> — ont déjà acheté en juillet. Base fidèle, CA S1 = 14 560 M.<br/>"
        "• <b>497 clients à risque élevé</b> — ces clients achètent habituellement avant le 21 du mois. "
        "Au 24/07, ils auraient dû commander. <b>1 851 M FCFA de CA S1</b> sont potentiellement perdus. "
        "Recontact immédiat nécessaire — vérifier s'ils sont partis chez la concurrence.<br/>"
        "• <b>56 clients à surveiller</b> — achètent habituellement entre le 22 et le 25. "
        "À recontacter dès le 26/07 si pas d'achat.<br/>"
        "• <b>41 clients pas encore inquiétants</b> — achètent en fin de mois (après le 25). "
        "Attendre la fin juillet avant de s'inquiéter.",
        BODY_BOLD
    ))
    story.append(Spacer(1, 0.3*cm))

    # Top 10 clients à risque élevé
    story.append(Paragraph("5.1.3 Top 10 clients à risque élevé — recontact immédiat", H3))
    story.append(Paragraph(
        "Ces clients achetaient habituellement avant le 21 du mois mais n'ont pas commandé en juillet (au 24/07). "
        "Ils représentent un CA S1 cumulé de 69 M FCFA pour les 10 premiers.",
        BODY
    ))

    top_risque_data = [
        ["#", "Client", "Agence", "Région", "CA S1 (M)", "Jour moy.", "Nb achats"],
        ["1", "TCHEUTCHOUA TCHINDA ERIC", "FAMLA", "Ouest", "11", "14", "124"],
        ["2", "GIC ESPOIR KELENG (POKAM MERLINE)", "FAMLA", "Ouest", "11", "15", "33"],
        ["3", "MALENDOMA DORIS TATIANA FLORE", "NDOBO", "Littoral", "8", "16", "25"],
        ["4", "METAFE GNITEYO SONYA M", "FAMLA", "Ouest", "8", "21", "11"],
        ["5", "JEUGO ALICE", "FAMLA", "Ouest", "7", "16", "62"],
        ["6", "FEEMDISPRO SARL", "MESSASSI", "Centre", "6", "12", "62"],
        ["7", "SOP EVELYNE", "DJELENG", "Ouest", "5", "16", "36"],
        ["8", "TSEUFFA EMMANUEL", "BUEA", "Littoral", "4", "18", "44"],
        ["9", "M. SONFACK FERDINAND (ETS ACCOUVEURS)", "NDOBO", "Littoral", "4", "12", "30"],
        ["10", "GIC EMERGENCE", "BERTOUA", "Centre", "4", "19", "33"],
    ]
    story.append(make_table(top_risque_data, col_widths=[0.8*cm, 4.5*cm, 1.8*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.5*cm], font_size=7.5))
    story.append(Paragraph(
        "<b>Action :</b> ces 10 clients doivent être recontactés <b>aujourd'hui</b> par leur commercial. "
        "TCHEUTCHOUA (124 achats S1) et FEEMDISPRO (62 achats) sont des clients très réguliers — "
        "leur silence en juillet est anormal et doit déclencher une alerte. FAMLA concentre 5 des 10 clients à risque.",
        BODY_BOLD
    ))

    story.append(Spacer(1, 0.4*cm))

    # Bilan CA gagné / perdu
    story.append(Paragraph("5.1.4 Bilan CA — gagné vs perdu (lecture ajustée)", H3))
    bilan_ca_data = [
        ["Indicateur", "Juillet 24j", "Lecture"],
        ["CA clients maintenus (déjà actifs juillet)", "14 560 M", "✅ Base fidèle, bon rythme"],
        ["CA à risque élevé (clients début de mois, pas d'achat)", "1 851 M", "🔴 À récupérer immédiatement"],
        ["CA à surveiller (clients mi-mois, pas d'achat)", "116 M", "🟡 À vérifier au 26/07"],
        ["CA pas inquiétant (clients fin de mois)", "42 M", "🟢 Attendre fin juillet"],
        ["CA gagné (nouveaux clients)", "—", "⚠️ 47 nouveaux clients en juillet"],
    ]
    story.append(make_table(bilan_ca_data, col_widths=[6*cm, 2.5*cm, 7*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))

    # Évolution concentrés
    story.append(Paragraph("5.1.5 Évolution spécifique concentrés", H3))
    conc_evol_data = [
        ["Indicateur", "S1 2026", "Juillet 24j", "Projection 31j"],
        ["Clients acheteurs concentrés", "1 016", "≈ 700", "≈ 900"],
        ["Clients concentrés maintenus", "—", "≈ 660 (65%)", "≈ 850"],
        ["Clients concentrés à risque élevé (début de mois)", "—", "≈ 360", "Recontact immédiat"],
        ["Nouveaux clients concentrés", "—", "≈ 40", "≈ 50"],
    ]
    story.append(make_table(conc_evol_data, col_widths=[5.5*cm, 2.8*cm, 2.8*cm, 3.5*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Lecture concentrés :</b> sur 1 016 clients qui ont acheté des concentrés en S1, environ 660 (65%) ont déjà acheté en juillet "
        "(24 jours). La projection à 31 jours (≈ 850) suggère un léger recul vs S1 — <b>la base se maintient mais ne croît pas</b>. "
        "Le plan d'action doit produire des conversions supplémentaires en fin de mois. "
        "Les clients concentrés à risque élevé (habitude début de mois, pas d'achat) doivent être recontactés en priorité.",
        BODY_BOLD
    ))

    # Notes prix
    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph("5.1.6 Évolution des prix — Juillet 2026", H3))
    story.append(Paragraph(
        "<b>⚠️ Augmentations tarifaires intervenues en juillet 2026 :</b><br/>"
        "• <b>Soja (tourteaux)</b> : hausse de <b>+1 000 FCFA</b> en début de mois, puis <b>+2 000 FCFA supplémentaires</b> par sac de 50 kg à partir du <b>23/07/2026</b> — soit <b>+3 000 FCFA/sac cumulé</b> au total sur juillet.<br/>"
        "• <b>PREMIX</b> : hausse de <b>+3 000 FCFA</b> sur tout le territoire<br/><br/>"
        "<b>Impact observé (au 24/07) :</b> le prix moyen soja est passé de 16 198 FCFA/sac (S1) à 18 322 FCFA/sac en juillet (soit +13,1%), ce qui correspond à l'effet combiné des deux hausses sur 24 jours (les 2 derniers jours à +3 000 FCFA/sac). <br/>"
        "<b>Risque :</b> la seconde hausse (effective 23/07) n'a pas encore produit son plein effet sur les volumes. "
        "Le suivi de fin juillet et d'août devra vérifier si ces hausses entraînent une accélération du churn, particulièrement si "
        "le concurrent reconstitue ses stocks de soja.<br/><br/>"
        "<b>Croisement avec l'analyse des habitudes :</b> parmi les 497 clients à risque élevé (début de mois, pas d'achat), "
        "il faut vérifier si la hausse du prix du soja est un facteur explicatif. Un appel téléphonique permet de distinguer "
        "les clients partis chez la concurrence (prix) de ceux qui retardent simplement leur commande.",
        BODY_BOLD
    ))

    story.append(Spacer(1, 0.4*cm))

    # ===== 5.1.7 Corrélation Booster × Concentrés =====
    story.append(Paragraph("5.1.7 Corrélation Chick/Piglet Booster × Concentrés — Évolution S1 → Juillet", H3))
    story.append(Paragraph(
        "Analyse du cross-sell entre les aliments complets (Chick/Piglet Booster) et les concentrés. "
        "Permet de mesurer si le plan d'action de cross-sell produit des résultats.",
        BODY
    ))

    booster_data = [
        ["Indicateur", "S1 2026", "Juillet 24j", "Évolution"],
        ["Clients Booster (ALIMENT COMPLET)", "494", "171", "—"],
        ["Booster + Concentrés (cross-sell)", "357 (72%)", "120 (70%)", "⚠️ -2 points"],
        ["Booster ONLY (pas de concentrés)", "137 (28%)", "51 (30%)", "⚠️ +2 points"],
        ["✅ Booster-only S1 → ont commencé conc. en juillet", "—", "8", "Conversion faible"],
        ["⚠️ Booster+Conc S1 → devenus booster-only en juillet", "—", "9", "Régression!"],
    ]
    story.append(make_table(booster_data, col_widths=[6*cm, 2.8*cm, 2.8*cm, 4*cm], font_size=7.5))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Lecture Booster × Concentrés :</b><br/>"
        "• Le taux de cross-sell Booster→Concentrés chute de 72% à 70% en juillet — <b>régression légère mais contenue</b>.<br/>"
        "• <b>9 clients</b> qui achetaient Booster + Concentrés en S1 n'achètent plus que le Booster en juillet — "
        "ils abandonnent les concentrés.<br/>"
        "• <b>8 clients</b> Booster-only ont été convertis vers les concentrés — le cross-sell fonctionne timidement.<br/>"
        "• Le bundle Booster+Concentrés doit être renforcé, particulièrement auprès des 51 clients Booster-only de juillet.",
        BODY_BOLD
    ))

    story.append(Spacer(1, 0.3*cm))

    # ===== 5.1.8 Clients soja sans concentrés =====
    story.append(Paragraph("5.1.8 Clients soja sans concentrés — Évolution du cross-sell", H3))
    story.append(Paragraph(
        "Suivi du taux de cross-sell soja→concentrés. C'est l'indicateur clé du bundle 3:1 — "
        "chaque commande de soja doit inclure au moins 1 sac de concentré.",
        BODY
    ))

    soja_conc_data = [
        ["Indicateur", "S1 2026", "Juillet 24j", "Évolution"],
        ["Clients acheteurs soja", "1 215", "704", "—"],
        ["Soja + Concentrés (cross-sell)", "958 (79%)", "540 (77%)", "⚠️ -2 points"],
        ["Soja ONLY (pas de concentrés)", "257 (21%)", "164 (23%)", "⚠️ +2 points"],
        ["✅ Soja-only S1 → ont commencé concentrés en juillet", "—", "30", "Conversion en cours"],
        ["⚠️ Soja+Conc S1 → arrêt concentrés en juillet", "—", "82", "Régression massive!"],
    ]
    story.append(make_table(soja_conc_data, col_widths=[6*cm, 2.8*cm, 2.8*cm, 4*cm], font_size=7.5))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Lecture soja → concentrés :</b><br/>"
        "• <b>Le taux de cross-sell soja→concentrés chute de 79% à 77%</b> (-2 points) — le bundle 3:1 n'est que partiellement appliqué.<br/>"
        "• <b>82 clients</b> qui achetaient soja + concentrés en S1 n'achètent plus que le soja en juillet — c'est une <b>régression massive</b>.<br/>"
        "• <b>30 clients</b> soja-only ont été convertis vers les concentrés — le plan d'action cross-sell commence à produire des résultats.<br/>"
        "• <b>164 clients</b> achètent du soja sans concentrés en juillet — ce sont les cibles immédiates du bundle.<br/>"
        "• La hausse du prix du soja (+1 000 puis +2 000 FCFA/sac au 23/07) pourrait expliquer en partie cette régression : les clients viennent chercher le soja (produit d'appel) "
        "mais ne veulent pas ou ne peuvent pas ajouter du concentré à leur commande.",
        BODY_BOLD
    ))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Conclusion comportement client (juillet) :</b> les deux indicateurs de cross-sell (Booster→Conc et Soja→Conc) "
        "montrent une <b>dégradation modérée</b>. Le bundle soja-concentrés 3:1, qui est l'action #1 du plan, "
        "n'est manifestement <b>pas encore pleinement déployé en agence</b>. Les 82 clients en régression soja+conc→soja-only "
        "représentent une perte directe de CA concentrés qui s'ajoute à la sous-performance structurelle. "
        "<b>Vérification terrain urgente : le bundle est-il communiqué aux commerciaux ? Sont-ils outillés pour l'imposer ?</b>",
        BODY_BOLD
    ))

    story.append(PageBreak())
    story.append(Paragraph("5.2 Tableau de bord — Efficacité du plan d'action", H2))
    story.append(Paragraph(
        "Ce tableau sera enrichi à chaque mise à jour mensuelle. Les colonnes 'Août', 'Septembre', etc. seront remplies "
        "au fur et à mesure de la disponibilité des données. L'objectif est de visualiser immédiatement si les actions "
        "produisent les résultats attendus.",
        BODY
    ))

    dashboard_data = [
        ["KPI", "Réf S1", "Cible S2", "Juil (24j)", "Août", "Sept", "Oct", "Nov", "Déc"],
        ["Clients maintenus (S1 → actifs)", "1 369", "Maximiser", "775 (57%)", "—", "—", "—", "—", "—"],
        ["Nouveaux clients", "—", "Maximiser", "47", "—", "—", "—", "—", "—"],
        ["Clients à risque élevé (habitude ≤21j)", "—", "Réduire", "497", "—", "—", "—", "—", "—"],
        ["Top 14 clients 20/80 réactivés", "0/14", "≥ 8/14", "À suivre", "—", "—", "—", "—", "—"],
        ["192 churned Q1→Q2 reconquis", "0", "≥ 58", "À suivre", "—", "—", "—", "—", "—"],
        ["CA gagné nouveaux clients (M)", "—", "Maximiser", "—", "—", "—", "—", "—", "—"],
        ["CA à risque élevé (M)", "—", "Récupérer", "1 851", "—", "—", "—", "—", "—"],
        ["Bundle soja-concentrés (%)", "79%", "≥ 80%", "77%", "—", "—", "—", "—", "—"],
        ["Volume CONCENTRES (t)", "1 513/mois", "≥ 1 942", "1 594 (proj.)", "—", "—", "—", "—", "—"],
        ["Atteinte obj. CONCENTRES", "72%", "≥ 100%", "88% (ajusté)", "—", "—", "—", "—", "—"],
        ["Nouveaux clients concentrés", "—", "Maximiser", "≈ 40", "—", "—", "—", "—", "—"],
        ["Cross-sell Booster→Conc (%)", "72%", "≥ 72%", "70%", "—", "—", "—", "—", "—"],
        ["Prix soja (FCFA/sac)", "Stable", "—", "+1 000 puis +2 000 (23/07)", "—", "—", "—", "—", "—"],
        ["Prix PREMIX (FCFA/sac)", "Stable", "—", "+3 000", "—", "—", "—", "—", "—"],
    ]
    story.append(make_table(dashboard_data, col_widths=[4*cm, 2*cm, 2*cm, 2.6*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.5*cm], font_size=7))
    story.append(Paragraph(
        "<i>* 497 clients S1 avec habitude d'achat ≤ 21j n'ont pas commandé au 24/07. À relativiser : 775 clients (57%) sont déjà actifs.</i>",
        ParagraphStyle('fn2', parent=SMALL, fontName='NotoSerifSC-Light', fontSize=7, textColor=GRAY, spaceBefore=2)
    ))

    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Première lecture du plan d'action (au 24/07) :</b><br/>"
        "• <b>CONCENTRES</b> : la projection juillet (1 594 t) est légèrement supérieure à la moyenne S1 (1 513 t) — <b>amélioration timide</b>. "
        "Le plan d'action sur les concentrés commence à produire des effets mais reste sous l'objectif au prorata (88%). Les actions critiques (bundle 3:1, sauvetage 14 clients 20/80, recontact 192 churned) doivent être accélérées.<br/>"
        "• <b>TOURTEAUX</b> : la projection (6 689 t) reste très supérieure à la moyenne S1 (4 844 t) — la rupture concurrente se maintient vigoureusement malgré la hausse tarifaire.<br/>"
        "• <b>Nouveaux clients</b> : 47 en 24 jours (projection ~60/mois) — l'acquisition reste insuffisante. La prospection des 118 jamais acquis doit être lancée.<br/>"
        "• <b>Cross-sell</b> : le taux soja→conc chute de 79% à 77% (-2 points). 82 clients ont arrêté les concentrés. Le bundle 3:1 n'est que partiellement déployé.<br/>"
        "• <b>Conclusion</b> : à 24 jours, le plan d'action produit des effets visibles mais insuffisants sur les concentrés. La fenêtre d'opportunité (rupture concurrente) se maintient mais le temps presse — il reste 7 jours pour rattraper l'écart à l'objectif.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ===== 5.3 Calendrier de suivi =====
    story.append(Paragraph("5.3 Calendrier de suivi et modalités", H2))
    story.append(Paragraph(
        "<b>Modalités de suivi :</b> revue mensuelle en comité de direction le 1er lundi de chaque mois. "
        "Le tableau de bord est alimenté par le contrôle de gestion à partir des données de ventes journalières. "
        "Alerte automatique si un indicateur s'écarte de plus de 15% de sa trajectoire cible.",
        BODY
    ))

    cal_suivi_data = [
        ["Échéance", "Données", "Action", "Responsable"],
        ["Fin juillet 2026", "Ventes juillet complet (31j)", "1er bilan S2 — comparaison vs S1 et vs cibles", "Data Analyst + Direction Commerciale"],
        ["Fin août 2026", "Ventes août complet", "Bilan n°1 — mesure d'impact du plan d'action 30 jours", "Data Analyst + Direction Commerciale"],
        ["Fin septembre 2026", "Ventes Q3 (Juil-Sept)", "Revue trimestrielle Q3 — ajustement du plan", "Comité de direction"],
        ["Fin octobre 2026", "Ventes octobre", "Bilan n°2 — mi-parcours S2", "Data Analyst + Direction Commerciale"],
        ["Fin novembre 2026", "Ventes novembre", "Bilan n°3 — préparation clôture S2", "Data Analyst + Direction Commerciale"],
        ["Fin décembre 2026", "Ventes S2 complet (Juil-Déc)", "Bilan global S2 et préparation du plan 2027", "Comité de direction"],
    ]
    story.append(make_table(cal_suivi_data, col_widths=[2.8*cm, 3.5*cm, 5.5*cm, 4.5*cm], font_size=8))

    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph(
        "<b>Questions à se poser à chaque mise à jour :</b><br/>"
        "1. <b>Le nombre de clients zéro achat diminue-t-il ?</b> Si non, le plan de réactivation n'est pas efficace — ajuster l'approche.<br/>"
        "2. <b>Les 14 clients 20/80 sont-ils réactivés ?</b> Si ≤ 4 sur 8 attendus, escalader au DG pour RDV direct.<br/>"
        "3. <b>Le bundle soja-concentrés est-il déployé ?</b> Si le taux < 50%, le déploiement n'est pas effectif — vérifier l'application en agence.<br/>"
        "4. <b>Le volume CONCENTRES progresse-t-il vs S1 ?</b> Si stagnation (cas de juillet), le plan d'action doit être renforcé.<br/>"
        "5. <b>Les 192 churned sont-ils reconquis ?</b> Si < 10% reconquis après 30 jours, revoir la stratégie de recontact.",
        BODY_BOLD
    ))

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
OUT = '/home/z/my-project/download/analyse_zero_achat.pdf'

doc = MyDocTemplate(OUT, pagesize=A4)
doc.title = "BELGOCAM SA - Analyse Zéro Achat (Suivi continu Janvier-Décembre 2026)"
doc.author = "William Francis Fohom — Data Analyst, Administrateur National de Ventes"
doc.subject = "Analyse Zéro Achat — diagnostic, pertes Q1, transition Q1→Q2 et plan d'action"
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
