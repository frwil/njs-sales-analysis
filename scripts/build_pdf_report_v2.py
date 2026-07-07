"""
BELGOCAM SA - Rapport d'analyse complete Janvier-Juin 2026
PDF generated via ReportLab

Author/redactor: William Francis Fohom, Data Analyst, Administrateur National de Ventes
"""
import os
import sys
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle,
    Image, KeepTogether, NextPageTemplate, PageTemplate, Frame, BaseDocTemplate
)
from reportlab.platypus.flowables import HRFlowable
from reportlab.pdfgen import canvas
from reportlab.graphics.shapes import Drawing, Rect, String, Line
from reportlab.graphics.charts.barcharts import VerticalBarChart, HorizontalBarChart
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.charts.legends import Legend
from reportlab.graphics.charts.textlabels import Label

# ===== FONTS =====
FONT_DIR = '/usr/share/fonts'
pdfmetrics.registerFont(TTFont('NotoSerifSC', f'{FONT_DIR}/truetype/noto-serif-sc/NotoSerifSC-Regular.ttf'))
pdfmetrics.registerFont(TTFont('NotoSerifSC-Bold', f'{FONT_DIR}/truetype/noto-serif-sc/NotoSerifSC-Bold.ttf'))
pdfmetrics.registerFont(TTFont('NotoSerifSC-Light', f'{FONT_DIR}/truetype/noto-serif-sc/NotoSerifSC-Light.ttf'))
pdfmetrics.registerFont(TTFont('LibSans', f'{FONT_DIR}/truetype/chinese/LiberationSans-Regular.ttf'))
pdfmetrics.registerFont(TTFont('LibSans-Bold', f'{FONT_DIR}/truetype/liberation/LiberationSans-Bold.ttf'))
registerFontFamily('NotoSerifSC', normal='NotoSerifSC', bold='NotoSerifSC-Bold')
registerFontFamily('LibSans', normal='LibSans', bold='LibSans-Bold')

# ===== COLORS (BELGOCAM corporate: navy + gold) =====
NAVY = colors.HexColor('#1F4E78')
NAVY_LIGHT = colors.HexColor('#2E75B6')
GOLD = colors.HexColor('#FFC000')
GOLD_DARK = colors.HexColor('#BF8F00')
RED = colors.HexColor('#C00000')
GREEN = colors.HexColor('#548235')
GRAY = colors.HexColor('#595959')
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
BULLET_BOLD = ParagraphStyle('BulletB', parent=BULLET, fontName='NotoSerifSC-Bold')

# Insight/Cause/Recommandation block styles
INSIGHT_TITLE = ParagraphStyle('InsightT', fontName='NotoSerifSC-Bold', fontSize=10,
                               textColor=NAVY, spaceBefore=6, spaceAfter=4, leading=13)
INSIGHT_BODY = ParagraphStyle('InsightB', fontName='NotoSerifSC', fontSize=9,
                              textColor=BLACK, spaceAfter=3, leading=12, leftIndent=12,
                              alignment=TA_JUSTIFY)
CAUSE_TITLE = ParagraphStyle('CauseT', fontName='NotoSerifSC-Bold', fontSize=10,
                             textColor=colors.HexColor('#8B0000'), spaceBefore=6, spaceAfter=4, leading=13)
CAUSE_BODY = ParagraphStyle('CauseB', fontName='NotoSerifSC', fontSize=9,
                            textColor=BLACK, spaceAfter=3, leading=12, leftIndent=12,
                            alignment=TA_JUSTIFY)
RECO_TITLE = ParagraphStyle('RecoT', fontName='NotoSerifSC-Bold', fontSize=10,
                            textColor=colors.HexColor('#375623'), spaceBefore=6, spaceAfter=4, leading=13)
RECO_BODY = ParagraphStyle('RecoB', fontName='NotoSerifSC', fontSize=9,
                           textColor=BLACK, spaceAfter=3, leading=12, leftIndent=12,
                           alignment=TA_JUSTIFY)


def icr_block(insights, causes, recommandations):
    """Create an Insights / Causes / Recommandations block for underperformance sections.
    Each argument is a list of strings (bullet points)."""
    elements = []
    # Insights
    elements.append(Paragraph("Insights", INSIGHT_TITLE))
    for s in insights:
        elements.append(Paragraph(f"• {s}", INSIGHT_BODY))
    # Causes
    elements.append(Paragraph("Causes", CAUSE_TITLE))
    for s in causes:
        elements.append(Paragraph(f"• {s}", CAUSE_BODY))
    # Recommandations
    elements.append(Paragraph("Recommandations", RECO_TITLE))
    for s in recommandations:
        elements.append(Paragraph(f"• {s}", RECO_BODY))
    return KeepTogether(elements)

# ===== PAGE SETUP =====
PAGE_W, PAGE_H = A4
MARGIN_L = 1.5*cm
MARGIN_R = 1.5*cm
MARGIN_T = 2.2*cm
MARGIN_B = 1.8*cm
CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R  # ~17.8 cm (vs 16.8 before)

# ===== COVER PAGE FUNCTION =====
def draw_cover(canv, doc):
    """Draw cover page elements directly on canvas."""
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
    canv.drawRightString(PAGE_W - MARGIN_R, PAGE_H - 3.2*cm, "Période analysée : Janvier - Juin 2026")

    # Main title block (centered)
    canv.setFillColor(NAVY)
    canv.setFont('NotoSerifSC-Bold', 28)
    canv.drawString(MARGIN_L, PAGE_H - 9*cm, "Analyse Complète")
    canv.setFont('NotoSerifSC-Bold', 28)
    canv.drawString(MARGIN_L, PAGE_H - 10.2*cm, "Janvier - Juin 2026")

    # Gold separator
    canv.setFillColor(GOLD)
    canv.rect(MARGIN_L, PAGE_H - 11*cm, 4*cm, 0.15*cm, fill=1, stroke=0)

    # Subtitle
    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC', 13)
    canv.drawString(MARGIN_L, PAGE_H - 12*cm, "État des lieux, transitions Q1→Q2,")
    canv.drawString(MARGIN_L, PAGE_H - 12.7*cm, "pertes estimées et plan d'action stratégique")

    # Key figures box
    box_y = 8*cm
    box_h = 4.5*cm
    canv.setFillColor(GRAY_VLIGHT)
    canv.rect(MARGIN_L, box_y, CONTENT_W, box_h, fill=1, stroke=0)
    # Left gold border
    canv.setFillColor(GOLD)
    canv.rect(MARGIN_L, box_y, 0.2*cm, box_h, fill=1, stroke=0)

    # 3 KPIs side by side
    canv.setFillColor(NAVY)
    canv.setFont('NotoSerifSC-Bold', 22)
    kpi_y = box_y + box_h - 1.5*cm
    col_w = CONTENT_W / 3
    canv.drawCentredString(MARGIN_L + col_w*0.5, kpi_y, "1 359")
    canv.drawCentredString(MARGIN_L + col_w*1.5, kpi_y, "17,3 Md")
    canv.drawCentredString(MARGIN_L + col_w*2.5, kpi_y, "+276 M")

    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC', 9)
    canv.drawCentredString(MARGIN_L + col_w*0.5, kpi_y - 0.7*cm, "clients analysés")
    canv.drawCentredString(MARGIN_L + col_w*1.5, kpi_y - 0.7*cm, "FCFA CA HT (6 mois)")
    canv.drawCentredString(MARGIN_L + col_w*2.5, kpi_y - 0.7*cm, "FCFA bilan net Q1→Q2")

    canv.setFillColor(GRAY)
    canv.setFont('NotoSerifSC-Light', 8)
    canv.drawCentredString(MARGIN_L + col_w*0.5, kpi_y - 1.4*cm, "(après exclusion interne)")
    canv.drawCentredString(MARGIN_L + col_w*1.5, kpi_y - 1.4*cm, "produits ciblés : soja + concentrés")
    canv.drawCentredString(MARGIN_L + col_w*2.5, kpi_y - 1.4*cm, "gain réactivés − perte churned")

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
    canv.drawRightString(PAGE_W - MARGIN_R, 0.55*cm, "Rapport d'analyse commerciale — Juillet 2026")

    canv.restoreState()


def draw_body_page(canv, doc):
    """Header/footer for body pages."""
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
    canv.drawString(MARGIN_L + 3*cm, PAGE_H - 1.0*cm, "Analyse complète Janvier-Juin 2026")
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


# ===== HELPER FUNCTIONS =====
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
    # Use a font-size-aware cell style
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
        [Paragraph(f'<b>{value}</b>', ParagraphStyle('kpi_v', fontName='NotoSerifSC-Bold', fontSize=18, textColor=NAVY, alignment=TA_CENTER, leading=22))],
        [Paragraph(label, ParagraphStyle('kpi_l', fontName='NotoSerifSC', fontSize=9, textColor=GRAY, alignment=TA_CENTER, leading=11))],
    ]
    if sublabel:
        data.append([Paragraph(sublabel, ParagraphStyle('kpi_s', fontName='NotoSerifSC-Light', fontSize=8, textColor=GRAY, alignment=TA_CENTER, leading=10))])
    t = Table(data, colWidths=[5.2*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), GRAY_VLIGHT),
        ('LINEBELOW', (0, 0), (-1, 0), 0.5, GOLD),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
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


def heading_with_content(heading_text, content_flowables, level=2):
    """Create a heading that stays together with its first content element.
    Prevents orphan headings at the bottom of pages.
    level=2 for H2, level=3 for H3.
    """
    style = H2 if level == 2 else H3
    heading = Paragraph(heading_text, style)
    # Keep heading + first content together
    if isinstance(content_flowables, list):
        return KeepTogether([heading] + content_flowables[:2])
    else:
        return KeepTogether([heading, content_flowables])


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
        ["", "Synthèse exécutive", "3"],
        ["1", "État des lieux et méthodologie", "4"],
        ["2", "Chiffres clés et segmentation clients", "5"],
        ["3", "Analyse Zéro Achat Q1", "6"],
        ["4", "Pertes Q1 estimées (méthode fréquence)", "8"],
        ["5", "Transition Q1→Q2 et dynamique mensuelle", "10"],
        ["6", "Analyse Chick Booster & Piglet Booster", "12"],
        ["7", "Perspectives stratégiques — 5 axes", "14"],
        ["8", "Projection CA 6 mois (3 scénarios)", "16"],
        ["9", "Ventes vs Objectifs 2026", "17"],
        ["9bis", "Focus CONCENTRÉS — Cœur de marge", "20"],
        ["10", "Forecast S2 2026 (3 scénarios)", "24"],
        ["10bis", "Analyse YoY 2025-2026 et forecast affiné", "26"],
        ["11", "Recommandations et plan d'action priorisé", "29"],
        ["12", "Conclusion et prochaines étapes", "31"],
    ]
    t = make_table(toc_data, col_widths=[1.5*cm, 12*cm, 2.5*cm], header_row=True)
    story.append(t)
    story.append(Spacer(1, 0.6*cm))
    story.append(Paragraph(
        "<i>Ce rapport présente l'analyse complète des ventes BELGOCAM SA sur la période Janvier-Juin 2026, "
        "après exclusion de 35 clients internes (filiales NJS, SPC, comptoirs d'agences et soldes comptables). "
        "L'analyse porte sur 1 359 clients actifs et 16 produits ciblés (4 tourteaux de soja + 10 BELGO 10% + 2 BELGO 5%).</i>",
        BODY_ITALIC
    ))
    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph(
        "<b>Glossaire des abréviations :</b> Q1 = Trimestre 1 (Janvier-Mars 2026) ; Q2 = Trimestre 2 (Avril-Juin 2026) ; "
        "S1 = Semestre 1 (Janvier-Juin 2026) ; S2 = Semestre 2 (Juillet-Décembre 2026) ; CA = Chiffre d'Affaires ; HT = Hors Taxes ; FCFA = Franc CFA.",
        SMALL
    ))
    story.append(PageBreak())

    # ---------- PAGE 3: SYNTHESE EXECUTIVE ----------
    story.append(Paragraph("Synthèse exécutive", H1))
    story.append(section_divider())

    # KPI row
    cards = [
        kpi_card("Clients analysés", "1 359", "(après exclusion interne)"),
        kpi_card("CA HT 6 mois", "17,3 Md", "FCFA"),
        kpi_card("Clients 20/80 (★)", "359", "80% du CA"),
        kpi_card("Bilan net Q1→Q2", "+276 M", "FCFA (ciblé)"),
    ]
    story.append(kpi_row(cards))
    story.append(Spacer(1, 0.6*cm))

    story.append(Paragraph("5 messages clés", H2))
    story.append(Paragraph(
        "<i>Note : Q1 = Trimestre 1 (Janvier-Mars 2026) ; Q2 = Trimestre 2 (Avril-Juin 2026). Glossaire complet en page 2.</i>",
        SMALL
    ))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>1. Base clients saine et concentrée.</b> BELGOCAM compte 1 359 clients actifs sur 6 mois (après exclusion de 35 clients internes). "
        "Le portefeuille est très concentré : 359 clients du top 20/80 (★) génèrent 80% du chiffre d'affaires, soit environ 13,8 Md FCFA sur les 17,3 Md FCFA totaux. "
        "Cette concentration est à la fois une force (relations directes avec les gros comptes) et un risque (dépendance à un nombre limité de clients).",
        BODY
    ))
    story.append(Paragraph(
        "<b>2. Pertes Q1 significatives mais récupérables.</b> 338 clients n'ont acheté aucun produit ciblé en Q1 (soja + concentrés), générant une perte estimée à 1 598 tonnes et 591 millions FCFA par la méthode fréquence/moyenne mensuelle. "
        "Sur ces 338 clients, 14 font partie du top 20/80 — ce sont des comptes stratégiques prioritaires à réactiver.",
        BODY
    ))
    story.append(Paragraph(
        "<b>3. Transition Q1→Q2 positive, portée par le soja.</b> Le bilan net Q1→Q2 sur produits ciblés est positif : +276 M FCFA (gain de réactivation 549 M − perte par churn 273 M). "
        "Cependant, ce bilan est principalement porté par les tourteaux de soja ; le bilan concentrés seul est fragile (+52 M FCFA seulement), avec un segment \"retenu\" en déclin de volume (-248 t).",
        BODY
    ))
    story.append(Paragraph(
        "<b>4. Concentrés = cœur de marge à défendre.</b> Les concentrés (BELGO 10% + BELGO 5%) représentent la plus grosse partie de la marge. "
        "Or 526 clients n'ont jamais acheté de concentrés en Q1, dont 41 clients 20/80. Les Booster (Chick + Piglet, 300 M FCFA, 573 clients) sont des produits d'entrée qui amènent 75% de leurs acheteurs vers les concentrés — levier de cross-sell à exploiter.",
        BODY
    ))
    story.append(Paragraph(
        "<b>5. Fenêtre stratégique exceptionnelle.</b> La rupture concurrente actuelle sur le soja offre une opportunité de conquête limitée dans le temps. "
        "Le scénario réaliste projette un CA additionnel de 640 M FCFA à 6 mois (+3,7% de croissance vs S1 2026), à condition de verrouiller contractuellement les nouveaux clients et de pousser les bundles soja+concentrés.",
        BODY
    ))

    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph(
        "<b>Recommandation immédiate :</b> traiter la période de rupture concurrente comme une opération de guerre. Lancer sous 15 jours l'opération \"Soja disponible\" auprès des 324 clients prioritaires (14 clients 20/80 + 192 churned + 118 jamais acquis), avec bundles soja+concentrés obligatoires et contrats multi-produits 6 mois pour verrouillage.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ---------- PAGE 4: SECTION 1 - ETAT DES LIEUX ----------
    story.append(Paragraph("1. État des lieux et méthodologie", H1))
    story.append(section_divider())

    story.append(Paragraph("1.1 Périmètre de l'analyse", H2))
    story.append(Paragraph(
        "Cette analyse porte sur l'ensemble des ventes BELGOCAM SA sur la période Janvier-Juin 2026, soit 6 mois d'activité. "
        "Le fichier source contient 6 feuilles mensuelles (une par mois), totalisant 61 242 lignes de commandes après application du filtre \"État = Livrée\". "
        "Chaque ligne correspond à une ligne de commande caractérisée par 18 colonnes : référence produit, description, quantité, référence commande, tiers (client), date de commande, mode de règlement, montant HT, montant TTC, auteur, dates de création/modification/clôture, statut facturé, état, statut facture et agence.",
        BODY
    ))
    story.append(Paragraph(
        "Pour les besoins de l'analyse de transition, la période est divisée en deux trimestres : <b>Q1 = Trimestre 1 (Janvier-Mars 2026)</b> et <b>Q2 = Trimestre 2 (Avril-Juin 2026)</b>. "
        "Cette subdivision trimestrielle permet d'analyser les évolutions de comportement clients entre le début et la fin du semestre.",
        BODY
    ))
    story.append(Paragraph(
        "Le périmètre géographique couvre l'ensemble du Cameroun via 16 agences BELGOCAM : FAMLA, MESSASSI, NDOBO, DJELENG, NKONGSAMBA, AHALA, BERTOUA, NGAOUNDERE, BAMENDA-MBOUDA, VILLAGE, BUEA, NKOABANG, PK11, NKOLBISSON, ainsi que 4 points SPC (BAF-CHEFFERIE, BUEA, YASSA, NDERE).",
        BODY
    ))

    story.append(Paragraph("1.2 Exclusion des clients internes", H2))
    story.append(Paragraph(
        "Pour obtenir une image fidèle du marché réel, 35 clients internes ont été exclus de l'analyse :",
        BODY
    ))
    excl_data = [
        ["Catégorie", "Nombre", "CA exclu (FCFA)", "Description"],
        ["Filiales NJS", "1", "676 653 750", "PROVENDERIE DU CENTRE (PDC) — filiale officielle Groupe NJS"],
        ["SPC magasins internes", "4", "103 024 000", "SPC + comptoirs SPC 4e étage, Village, PK15"],
        ["Clients comptoirs agences", "21", "775 094 760", "CLIENTS COMPTOIR BERTOUA, MESSASSI, FAMLA, PK11, NKOLBISSON, AHALA, NGAOUNDERE, NKOABANG, BAF-DJELENG, BUEA, YAOUNDE, NDERE, etc."],
        ["Soldes comptables", "9", "56 120 550", "SOLDE COMPTA — écritures de régularisation"],
        ["TOTAL EXCLUS", "35", "1 610 893 060", "≈ 8,5% du CA brut — réintégration évitée"],
    ]
    story.append(make_table(excl_data, col_widths=[3.5*cm, 1.5*cm, 3.5*cm, 8.5*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "Base d'analyse finale : <b>1 359 clients actifs</b>, représentant un CA HT cumulé de <b>17,3 milliards FCFA</b> sur 6 mois.",
        BODY_BOLD
    ))

    story.append(Paragraph("1.3 Produits ciblés et catégories", H2))
    story.append(Paragraph(
        "L'analyse se focalise sur 16 produits \"ciblés\" représentant le cœur de l'activité BELGOCAM, répartis en 3 sous-catégories :",
        BODY
    ))
    prod_data = [
        ["Catégorie", "Nb produits", "Références", "Volume 6 mois (lignes)"],
        ["Tourteaux de soja", "4", "T102 (50Kg), T1021 (1Kg), T1023 (5Kg), T1024 (25Kg)", "16 730 lignes (40% du volume)"],
        ["BELGO 10%", "10", "C104 Chair 50Kg, C1042/43/44, C105 Porc 50Kg + formats, C102 Ponte 50Kg + 5Kg", "14 940 lignes"],
        ["BELGO 5%", "2", "C103 Chair 50Kg, C101 Ponte 50Kg", "7 645 lignes"],
        ["TOTAL CIBLÉ", "16", "—", "39 315 lignes (≈ 64% des lignes Livrées)"],
    ]
    story.append(make_table(prod_data, col_widths=[3.2*cm, 1.5*cm, 7.5*cm, 5.6*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "Une analyse complémentaire est également conduite sur les 12 concentrés seuls (BELGO 10% + BELGO 5%, sans le soja) car ils représentent <b>la plus grosse partie de la marge</b>. "
        "Enfin, les gammes Chick Booster (CB100, CB101) et Piglet Booster (CB200, CB201) sont analysées pour étudier les synergies avec les concentrés.",
        BODY
    ))

    story.append(Paragraph("1.4 Méthodologie des analyses", H2))
    story.append(Paragraph(
        "<b>Pareto 20/80</b> : les clients sont triés par CA HT décroissant. Le seuil 20/80 est atteint quand la somme cumulée des CA atteint 80% du CA total. Les clients au-dessus du seuil sont marqués ★ (359 clients).",
        BODY
    ))
    story.append(Paragraph(
        "<b>Pertes Q1 estimées (méthode fréquence)</b> : pour chaque client \"zéro achat\" en Q1, la perte est calculée comme suit :<br/>"
        "Perte Q1 = Σ(moyenne mensuelle par produit en Q2) × Fréquence × 3 mois<br/>"
        "où la moyenne mensuelle par produit = (volume total Q2 du produit) / (nombre de mois Q2 où le produit a été acheté), "
        "et la fréquence = (nombre de mois Q2 avec achat de la catégorie ciblée) / 3. "
        "Cette méthode pondère correctement les clients à forte fréquence (qui auraient acheté tous les mois en Q1) par rapport aux clients occasionnels.",
        BODY
    ))
    story.append(Paragraph(
        "<b>Transition Q1→Q2</b> : chaque client est classé selon son statut d'achat sur produits ciblés en Q1 (Zéro achat/Active) et Q2 (Zéro achat/Active), "
        "ce qui donne 4 segments : Persistant (Zéro achat→Zéro achat), Réactivé (Zéro achat→Active, GAIN), Retenu (Active→Active), Churned (Active→Zéro achat, PERTE).",
        BODY
    ))

    story.append(PageBreak())

    # ---------- PAGE 5: SECTION 2 - CHIFFRES CLES ----------
    story.append(Paragraph("2. Chiffres clés et segmentation clients", H1))
    story.append(section_divider())

    story.append(Paragraph("2.1 Indicateurs globaux", H2))
    cards2 = [
        kpi_card("CA HT total 6 mois", "17,3 Md", "FCFA"),
        kpi_card("Clients 20/80 (★)", "359", "26,4% des clients"),
        kpi_card("CA des 20/80", "13,8 Md", "FCFA (80% du CA)"),
        kpi_card("CA moyen / client", "12,7 M", "FCFA"),
    ]
    story.append(kpi_row(cards2))
    story.append(Spacer(1, 0.5*cm))

    story.append(Paragraph("2.2 Distribution de la fréquence d'achat (sur 6 mois)", H2))
    freq_data = [
        ["Fréquence (nb mois actifs)", "Nb clients", "% du total", "Catégorie"],
        ["6 mois (fidèles parfaits)", "452", "33,3%", "Fidèles"],
        ["5 mois", "141", "10,4%", "Très réguliers"],
        ["4 mois", "131", "9,6%", "Réguliers"],
        ["3 mois", "145", "10,7%", "Semi-fidèles"],
        ["2 mois", "196", "14,4%", "Occasionnels"],
        ["1 mois seulement", "294", "21,6%", "One-shot"],
    ]
    story.append(make_table(freq_data, col_widths=[5*cm, 3*cm, 3*cm, 5*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "Près d'un tiers des clients (452) sont fidèles sur les 6 mois — c'est une base solide. En revanche, 294 clients (21,6%) n'ont fait qu'un seul achat sur 6 mois, "
        "ce qui représente un potentiel important de réactivation. La fréquence moyenne est de 3,5 mois d'achat par client.",
        BODY
    ))
    # Chart: Distribution fréquence
    story.append(Spacer(1, 0.3*cm))
    img = Image('/home/z/my-project/scripts/pdf_charts/chart8_frequence.png', width=14*cm, height=7.9*cm)
    story.append(img)
    story.append(Paragraph("Figure 1 — Distribution de la fréquence d'achat (6 mois)", CAPTION))

    story.append(Paragraph("2.3 Top 5 agences par nombre de clients", H2))
    ag_data = [
        ["Agence", "Clients", "CA HT 6 mois (FCFA)", "CA moyen/client", "% Clients zéro achat Q1"],
        ["AGENCE FAMLA", "293", "4 359 027 250", "14,87 M", "19,1%"],
        ["AGENCE MESSASSI", "167", "1 633 657 768", "9,78 M", "33,5%"],
        ["AGENCE NDOBO", "154", "1 782 721 571", "11,58 M", "29,2%"],
        ["AGENCE DJELENG", "107", "1 317 911 317", "12,32 M", "25,2%"],
        ["AGENCE NKONGSAMBA", "97", "696 795 842", "7,18 M", "—"],
        ["AGENCE NGAOUNDERE", "91", "522 685 737", "5,74 M", "34,1%"],
    ]
    story.append(make_table(ag_data, col_widths=[4*cm, 1.5*cm, 4*cm, 3*cm, 3.5*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Points d'attention :</b> MESSASSI (33,5%) et NGAOUNDERE (34,1%) ont plus d'un client sur trois en zéro achat Q1 — c'est anormal et mérite un diagnostic terrain. "
        "FAMLA reste l'agence locomotive avec 293 clients et 4,36 Md FCFA de CA (25% du CA total).",
        BODY
    ))

    story.append(PageBreak())

    # ---------- PAGE 6: SECTION 3 - ZERO ACHAT ----------
    story.append(Paragraph("3. Analyse Zéro Achat Q1", H1))
    story.append(section_divider())

    story.append(Paragraph("3.1 Vue d'ensemble", H2))
    cards3 = [
        kpi_card("Zéro achat global Q1", "338", "clients (16 produits)"),
        kpi_card("Zéro achat concentrés Q1", "526", "clients (12 concentrés)"),
        kpi_card("20/80 en clients zéro achat Q1", "14", "clients prioritaires"),
        kpi_card("Jamais acquis (6 mois)", "118", "clients persistants"),
    ]
    story.append(kpi_row(cards3))
    story.append(Spacer(1, 0.4*cm))

    story.append(Paragraph("3.2 Top 14 clients 20/80 prioritaires à réactiver", H2))
    story.append(Paragraph(
        "Ces 14 clients font partie du top 20/80 (★) mais n'ont acheté AUCUN produit ciblé en Q1. Ils sont à recontacter en priorité absolue pendant la rupture concurrente de soja.",
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
        "<b>Concentration géographique :</b> FAMLA (6 clients), NDOBO (4 clients), DJELENG (2 clients), AHALA (1), NDOBO (1). "
        "Perte Q1 cumulée des 12 premiers = 173 M FCFA. Action commerciale territoriale à organiser agence par agence.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ---------- PAGE 7: SECTION 3 SUITE ----------
    story.append(Paragraph("3.3 Zéro achat par agence", H2))
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
    img = Image('/home/z/my-project/scripts/pdf_charts/chart4_zero_agence.png', width=15*cm, height=8.3*cm)
    story.append(img)
    story.append(Paragraph("Figure 2 — Répartition zéro achat Q1 (ciblé) par agence", CAPTION))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Constat :</b> MESSASSI et NGAOUNDERE affichent un taux de zéro achat critique (≥33%), alors qu'aucun client 20/80 n'est en zéro achat dans ces agences. "
        "Cela suggère un problème de couverture commerciale sur les clients moyens — potentiellement un sous-dimensionnement de l'effectif commercial ou un manque de visibilité locale. "
        "FAMLA concentre la moitié des clients 20/80 en zéro achat (7 sur 14), ce qui en fait l'agence prioritaire pour l'action de sauvetage.",
        BODY
    ))

    story.append(Paragraph("3.4 Zéro achat concentrés (analyse distincte)", H2))
    story.append(Paragraph(
        "526 clients n'ont acheté aucun concentré en Q1 (vs 338 sur le ciblé global). La différence (188 clients) correspond aux clients qui ont acheté du soja en Q1 mais pas de concentrés — "
        "ce sont les cibles idéales pour une action de cross-sell vers les concentrés.",
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
        BODY
    ))

    story.append(PageBreak())

    # ---------- PAGE 8: SECTION 4 - PERTES Q1 ----------
    story.append(Paragraph("4. Pertes Q1 estimées (méthode fréquence)", H1))
    story.append(section_divider())

    story.append(Paragraph("4.1 Méthode de calcul", H2))
    story.append(Paragraph(
        "La perte Q1 estimée pour chaque client \"zéro achat\" est calculée à partir de son comportement Q2, en tenant compte de sa fréquence d'achat :",
        BODY
    ))
    story.append(Paragraph(
        "<b>Perte Q1 = Σ(moyenne mensuelle par produit en Q2) × Fréquence × 3 mois</b>",
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
        "Cette méthode permet de distinguer un client qui aurait acheté tous les mois (fréquence = 1, perte élevée) d'un client occasionnel qui n'aurait acheté qu'un mois sur trois (fréquence = 0,33, perte plus faible).",
        BODY
    ))

    story.append(Paragraph("4.2 Résultats agrégés", H2))
    pertes_data = [
        ["Catégorie", "Clients concernés", "Volume perdu (t)", "CA perdu (FCFA)"],
        ["Zero achat global (16 produits)", "338", "1 598,66", "580 819 762"],
        ["Zero achat concentrés (12 produits)", "526", "296,20", "199 660 105"],
        ["Dont clients 20/80 (zéro achat global)", "14", "≈ 600", "≈ 175 000 000"],
    ]
    story.append(make_table(pertes_data, col_widths=[6*cm, 3.5*cm, 3*cm, 4*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>581 millions FCFA</b> de CA potentiel perdu en Q1 sur les seuls clients zéro achat global. "
        "C'est l'équivalent de 80% du CA total 6 mois de ces mêmes clients, ce qui démontre l'ampleur de l'opportunité commerciale.",
        BODY_BOLD
    ))

    story.append(Paragraph("4.3 Top 10 pertes par client (ciblé global)", H2))
    top10_data = [
        ["N°", "Client", "Fréquence Q2", "Σ Moy. mens. (t)", "Perte Volume (t)", "Perte CA (FCFA)"],
        ["1", "SIGHELO SARL", "0,33 (1 mois)", "74,00", "74,00", "23 680 000"],
        ["2", "TEIKING JEAN MARIE", "0,67 (2 mois)", "30,50", "61,00", "19 632 000"],
        ["3", "COMPAGNIE FERMIERE CAM.", "0,33 (1 mois)", "60,00", "60,00", "21 000 000"],
        ["4", "FEUDJIO BACK ARMEL", "0,67 (2 mois)", "27,50", "55,00", "17 460 000"],
        ["5", "KENMEGNE ALAIN", "0,33 (1 mois)", "50,00", "50,00", "16 600 000"],
        ["6", "FOTSO VINCENT", "1,00 (3 mois)", "14,00", "42,00", "13 464 000"],
        ["7", "DJOUSSI AURORE FLORINDA", "0,33 (1 mois)", "40,00", "40,00", "13 280 000"],
        ["8", "STE SATI SARL", "0,33 (1 mois)", "40,00", "40,00", "13 600 000"],
        ["9", "STE IPACAM & FILS", "1,00 (3 mois)", "12,25", "36,75", "14 754 000"],
        ["10", "BAHO", "0,33 (1 mois)", "35,00", "35,00", "10 920 000"],
    ]
    story.append(make_table(top10_data, col_widths=[0.8*cm, 5*cm, 3.2*cm, 3*cm, 2.7*cm, 3.6*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))
    # Chart: Top 10 pertes
    img = Image('/home/z/my-project/scripts/pdf_charts/chart3_top10_pertes.png', width=15*cm, height=8.3*cm)
    story.append(img)
    story.append(Paragraph("Figure 3 — Top 10 pertes Q1 (clients zéro achat global, en millions FCFA)", CAPTION))

    story.append(PageBreak())

    # ---------- PAGE 9: SECTION 4 SUITE ----------
    story.append(Paragraph("4.4 Lecture de l'impact de la fréquence", H2))
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
        "<b>Insight :</b> un client à haute fréquence d'achat (FOTSO, 3/3 mois) est plus \"récupérable\" à court terme qu'un client à faible fréquence mais gros volume (SIGHELO, 1/3 mois). "
        "Les actions de réactivation doivent donc prioriser en premier lieu les clients à forte fréquence Q2, qui démontrent un besoin d'achat récurrent.",
        BODY
    ))

    story.append(Paragraph("4.5 Pertes sur concentrés uniquement", H2))
    story.append(Paragraph(
        "Sur les 526 clients zéro achat concentrés en Q1, la perte estimée est de 296 tonnes et 199 millions FCFA. "
        "C'est moins que sur le ciblé global (car certains de ces clients ont quand même acheté du soja), mais c'est plus préoccupant stratégiquement car les concentrés sont le cœur de marge.",
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

    # ---------- PAGE 10: SECTION 5 - TRANSITION ----------
    story.append(Paragraph("5. Transition Q1→Q2 et dynamique mensuelle", H1))
    story.append(section_divider())

    story.append(Paragraph("5.1 Matrice de transition Q1→Q2 (16 produits ciblés)", H2))
    matrice = [
        ["", "Q2 : Zéro achat", "Q2 : Active", "Total Q1"],
        ["Q1 : Zéro achat", "118 (persistants)", "234 (réactivés) ★", "352"],
        ["Q1 : Active", "192 (churned) ⚠", "850 (retenus)", "1 042"],
        ["Total Q2", "310", "1 084", "1 359"],
    ]
    story.append(make_table(matrice, col_widths=[3.5*cm, 4*cm, 4.5*cm, 3*cm]))
    story.append(Spacer(1, 0.3*cm))
    # Chart: Segments de transition
    img = Image('/home/z/my-project/scripts/pdf_charts/chart2_segments.png', width=14*cm, height=7.9*cm)
    story.append(img)
    story.append(Paragraph("Figure 4 — Segments de transition Q1 → Q2 (16 produits ciblés)", CAPTION))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Quatre segments :</b><br/>"
        "• <b>Persistants (118)</b> : clients qui n'ont jamais acheté de ciblé sur 6 mois — opportunité non exploitée, à prospecter.<br/>"
        "• <b>Réactivés (234)</b> : GAIN — clients sans achat Q1 qui ont acheté en Q2. CA généré : 549,7 M FCFA.<br/>"
        "• <b>Retenus (850)</b> : STABLES — cœur de clientèle, achats Q1 et Q2.<br/>"
        "• <b>Churned (192)</b> : PERTE — clients actifs Q1 devenus inactifs Q2. CA perdu : 273,3 M FCFA.",
        BODY
    ))

    story.append(Paragraph("5.2 Volumes et CA par segment", H2))
    seg_data = [
        ["Segment", "Clients", "Vol Q1 (t)", "Vol Q2 (t)", "Δ Vol (t)", "CA Q1 (FCFA)", "CA Q2 (FCFA)"],
        ["Persistants", "118", "0", "0", "0", "0", "0"],
        ["Réactivés", "234", "0", "1 524,74", "+1 524,74", "0", "549 689 847"],
        ["Retenus", "850", "16 720,86", "19 133,23", "+2 412,37", "7,02 Md", "7,66 Md"],
        ["Churned", "192", "758,01", "0", "-758,01", "273,3 M", "0"],
        ["TOTAL", "1 359", "17 478,87", "20 657,97", "+3 179,10", "17,3 Md", "17,9 Md"],
    ]
    story.append(make_table(seg_data, col_widths=[2.8*cm, 1.6*cm, 2.5*cm, 2.5*cm, 2.2*cm, 2.9*cm, 2.9*cm], font_size=8))

    story.append(Paragraph("5.3 Bilan net Q1→Q2", H2))
    bilan_data = [
        ["", "Clients", "Volume (t)", "CA (FCFA)"],
        ["GAIN — Réactivés (Q1 Zéro achat → Q2 Active)", "234", "+1 524,74", "+549 689 847"],
        ["PERTE — Churned (Q1 Active → Q2 Zéro achat)", "192", "-758,01", "-273 296 995"],
        ["BILAN NET", "+42", "+766,73", "+276 392 852"],
    ]
    story.append(make_table(bilan_data, col_widths=[7*cm, 2.5*cm, 3*cm, 3.5*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Bilan positif :</b> BELGOCAM a gagné 276 millions FCFA nets sur les produits ciblés entre Q1 et Q2. "
        "Le gain de réactivation (549 M) dépasse largement la perte par churn (273 M). "
        "Cependant, ce bilan est principalement porté par le soja — le bilan concentrés seul est beaucoup plus fragile.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ---------- PAGE 11: SECTION 5 SUITE ----------
    story.append(Paragraph("5.4 Comparatif Ciblé (16) vs Concentrés (12)", H2))
    comp_data = [
        ["Indicateur", "Ciblé (16 produits)", "Concentrés (12 produits)", "Écart"],
        ["Clients zéro achat Q1", "352", "544", "+192 (ceux qui ont acheté soja mais pas conc.)"],
        ["Clients Active Q1", "1 042", "850", "-192"],
        ["Réactivés Q2", "234", "184", "-50"],
        ["Churned Q2", "192", "178", "-14"],
        ["Bilan net CA (FCFA)", "+276 M", "+52 M", "-224 M"],
        ["Q1 Fidèles (3 mois)", "527", "381", "-146"],
        ["Q1 Fidèles churned Q2", "10 (1,9%)", "12 (3,1%)", "+2"],
    ]
    story.append(make_table(comp_data, col_widths=[5*cm, 3.5*cm, 3.5*cm, 4.5*cm]))
    story.append(Spacer(1, 0.3*cm))
    # Chart: Bilan net Ciblé vs Concentrés
    img = Image('/home/z/my-project/scripts/pdf_charts/chart5_bilan_net.png', width=14*cm, height=7.9*cm)
    story.append(img)
    story.append(Paragraph("Figure 6 — Bilan net Q1 → Q2 : Ciblé vs Concentrés (millions FCFA)", CAPTION))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Insight clé :</b> le bilan net sur les concentrés (+52 M) est <b>5 fois plus faible</b> que sur le ciblé global (+276 M). "
        "La dynamique positive Q1→Q2 est donc portée principalement par les tourteaux de soja, pas par les concentrés. "
        "De plus, le taux de churn des fidèles Q1 est plus élevé sur concentrés (3,1%) que sur ciblé (1,9%) — les concentrés sont plus sensibles à l'attrition.",
        BODY
    ))

    story.append(Paragraph("5.5 Destin des clients fidèles Q1", H2))
    story.append(Paragraph(
        "Sur les 527 clients fidèles en Q1 (3 mois d'achat ciblé), voici leur destin en Q2 :",
        BODY
    ))
    fidel_data = [
        ["Statut Q2", "Nb clients", "% des Q1 Fidèles", "Commentaire"],
        ["Q2 Fidèle (3 mois)", "420", "79,7%", "Maintenus — cœur de clientèle stable"],
        ["Q2 Semi-fidèle (1-2 mois)", "97", "18,4%", "Déclin de fréquence — à surveiller (risque churn)"],
        ["Q2 Zéro achat (churned)", "10", "1,9%", "Perte sèche — anciens fidèles devenus inactifs"],
    ]
    story.append(make_table(fidel_data, col_widths=[5*cm, 2.5*cm, 3*cm, 6*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "Les 97 fidèles en déclin (18,4%) sont un signal d'alerte pré-churn à traiter en priorité. "
        "Une visite physique par le commercial dédié permettrait de comprendre la cause du déclin (insatisfaction, concurrence, changement d'activité) avant que le client ne churn totalement.",
        BODY
    ))

    story.append(Paragraph("5.6 Dynamique mensuelle Q2 (Avril / Mai / Juin)", H2))
    story.append(Paragraph(
        "L'analyse mois par mois révèle le moment exact de (ré)activation ou du churn :",
        BODY
    ))
    monthly_data = [
        ["Mois", "Clients actifs (tous produits)", "Clients actifs ciblé", "Clients actifs concentrés"],
        ["Janvier", "941", "815", "634"],
        ["Février", "836", "690", "534"],
        ["Mars", "859", "718", "560"],
        ["Avril", "836", "705", "547"],
        ["Mai", "853", "738", "573"],
        ["Juin", "958", "856", "672"],
    ]
    story.append(make_table(monthly_data, col_widths=[2.8*cm, 4.9*cm, 4.9*cm, 4.9*cm]))
    story.append(Spacer(1, 0.3*cm))
    # Chart: Évolution mensuelle
    img = Image('/home/z/my-project/scripts/pdf_charts/chart1_evolution.png', width=15*cm, height=8.4*cm)
    story.append(img)
    story.append(Paragraph("Figure 5 — Évolution mensuelle du nombre de clients actifs", CAPTION))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Tendance :</b> juin 2026 marque un rebond net (+118 clients actifs ciblé vs mai), possiblement lié au début de la rupture concurrente sur le soja. "
        "À confirmer sur les mois suivants.",
        BODY
    ))

    story.append(PageBreak())

    # ---------- PAGE 12: SECTION 6 - CHICK & PIGLET ----------
    story.append(Paragraph("6. Analyse Chick Booster & Piglet Booster", H1))
    story.append(section_divider())

    story.append(Paragraph("6.1 Ventes sur 6 mois par référence", H2))
    ventes_data = [
        ["Réf.", "Description", "Volume (t)", "CA HT (FCFA)", "Gamme"],
        ["CB100", "CHICK BOOSTER 25 Kg", "172,03", "144 857 350", "Chick"],
        ["CB101", "CHICK BOOSTER 5Kg", "1,12", "977 950", "Chick"],
        ["SOUS-TOTAL CHICK", "", "173,15", "145 835 300", ""],
        ["CB200", "PIGLET BOOSTER 25Kg", "191,05", "152 882 650", "Piglet"],
        ["CB201", "PIGLET BOOSTER 5Kg", "1,52", "1 246 796", "Piglet"],
        ["SOUS-TOTAL PIGLET", "", "192,57", "154 129 446", ""],
        ["TOTAL BOOSTER", "", "365,72", "299 964 746", "≈ 1,7% CA total"],
    ]
    story.append(make_table(ventes_data, col_widths=[1.5*cm, 5.5*cm, 2.5*cm, 4*cm, 3.5*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Insight :</b> Piglet Booster dépasse légèrement Chick Booster (154 M vs 146 M FCFA), mais les deux gammes sont proches. "
        "Ensemble, les Booster représentent ~300 M FCFA, soit ~1,7% du CA total. "
        "C'est peu en volume financier, mais stratégique car ce sont des <b>produits d'entrée</b> qui amènent les clients vers les concentrés (cœur de marge).",
        BODY
    ))

    story.append(Paragraph("6.2 Synergie Chick Booster × Concentrés Chair/Ponte", H2))
    story.append(Paragraph(
        "Sur les 343 clients qui achètent du Chick Booster, combien achètent aussi des concentrés volaille (Chair ou Ponte) ?",
        BODY
    ))
    syn_chick = [
        ["Catégorie", "Nb clients", "% des clients Chick", "Commentaire"],
        ["Total clients Chick Booster", "343", "100%", "Base de référence"],
        ["Chick + Concentrés Chair (C104/C1042/C1043/C1044/C103)", "263", "76,7%", "Cross-sell Chair"],
        ["Chick + Concentrés Ponte (C102/C1022/C101)", "147", "42,9%", "Cross-sell Ponte"],
        ["Chick + Chair OU Ponte (cross-sell réussi)", "270", "78,7%", "✅ Cross-sell validé"],
        ["Chick ONLY (pas de concentrés Chair/Ponte)", "73", "21,3%", "⚠️ Opportunité cross-sell énorme"],
    ]
    story.append(make_table(syn_chick, col_widths=[7*cm, 2*cm, 2.5*cm, 5.5*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Excellent taux de cross-sell :</b> 78,7% des clients Chick achètent aussi des concentrés volaille. "
        "Reste <b>73 clients à convertir</b> — ce sont des éleveurs volaille qui démarrent leurs poussins avec le Chick Booster mais ne vont pas jusqu'au concentré. "
        "Cible parfaite pour un bundle \"Chick + Chair\" avec remise.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ---------- PAGE 13: SECTION 6 SUITE ----------
    story.append(Paragraph("6.3 Synergie Piglet Booster × Concentrés Porc", H2))
    story.append(Paragraph(
        "Sur les 230 clients qui achètent du Piglet Booster, combien achètent aussi des concentrés Porc (C105, C1053, C1054, C1055) ?",
        BODY
    ))
    syn_piglet = [
        ["Catégorie", "Nb clients", "% des clients Piglet", "Commentaire"],
        ["Total clients Piglet Booster", "230", "100%", "Base de référence"],
        ["Piglet + Concentrés Porc (cross-sell réussi)", "162", "70,4%", "✅ Cross-sell validé"],
        ["Piglet ONLY (pas de concentrés Porc)", "68", "29,6%", "⚠️ Opportunité cross-sell énorme"],
    ]
    story.append(make_table(syn_piglet, col_widths=[7*cm, 2*cm, 2.5*cm, 5.5*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>68 clients Piglet-only à convertir sur Porc</b> — c'est une opportunité cross-sell majeure. "
        "Le Piglet Booster est un produit d'entrée naturel pour l'élevage porcin, et ces clients devraient logiquement consommer aussi du concentré Porc. "
        "Bundle \"Piglet + Porc\" à proposer systématiquement.",
        BODY_BOLD
    ))

    story.append(Paragraph("6.4 Synthèse stratégique Booster × Concentrés", H2))
    story.append(Paragraph(
        "Sur 573 clients Booster (Chick + Piglet, sans doublons), 432 (75%) font déjà du cross-sell vers les concentrés. "
        "Les 141 \"Booster-only\" représentent une <b>opportunité de cross-sell de première main</b> pour augmenter les ventes de concentrés sans avoir à chercher de nouveaux clients.",
        BODY
    ))
    syn_glob = [
        ["Gamme Booster", "Clients total", "Cross-sell réussi", "Booster-only (à convertir)"],
        ["Chick Booster → Chair/Ponte", "343", "270 (78,7%)", "73 (21,3%)"],
        ["Piglet Booster → Porc", "230", "162 (70,4%)", "68 (29,6%)"],
        ["TOTAL Booster", "573 (sans doublons)", "432 (75,4%)", "141 (24,6%)"],
    ]
    story.append(make_table(syn_glob, col_widths=[5*cm, 3.5*cm, 4*cm, 4*cm]))
    story.append(Spacer(1, 0.3*cm))
    # Chart: Synergie Booster
    img = Image('/home/z/my-project/scripts/pdf_charts/chart6_synergie_booster.png', width=16*cm, height=7.2*cm)
    story.append(img)
    story.append(Paragraph("Figure 7 — Synergie Chick/Piglet Booster × Concentrés", CAPTION))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Conclusion stratégique :</b> les Booster sont des produits d'entrée qui amènent naturellement 75% de leurs acheteurs vers les concentrés. "
        "Tout effort sur les Booster (promotion, échantillon) a un <b>effet multiplicateur</b> sur les ventes de concentrés (cœur de marge). "
        "Les 141 clients Booster-only sont la cible #1 du cross-sell immédiat.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ---------- PAGE 14: SECTION 7 - PERSPECTIVES ----------
    story.append(Paragraph("7. Perspectives stratégiques — 5 axes", H1))
    story.append(section_divider())

    story.append(Paragraph("7.1 Axe 1 — Conquête soja (rupture concurrentielle)", H2))
    story.append(Paragraph(
        "<b>Constat :</b> les concurrents sont en rupture de soja. Fenêtre de tir exceptionnelle : le soja représente 16 730 lignes / 6 mois (40% du volume). "
        "CA additionnel potentiel : 300-600 M FCFA.",
        BODY_BOLD
    ))
    story.append(Paragraph(
        "<b>Cibles prioritaires :</b><br/>"
        "• 14 clients 20/80 zéro achat Q1 (ciblé) — RDV direct commercial<br/>"
        "• 192 clients churned Q1→Q2 (273 M FCFA perdus) — appel téléphonique<br/>"
        "• 118 clients persistants zéro achat (jamais acquis) — offre découverte<br/>"
        "• Clients concurrents en rupture — conquête pure",
        BODY
    ))
    story.append(Paragraph(
        "<b>Actions concrètes :</b> opération \"Soja disponible\" J+7 ; argumentaire \"BELGOCAM = sécurité d'approvisionnement\" ; "
        "bundles soja + concentrés (-5% conc. si achat soja) ; vérification quotidienne stocks soja ; verrouillage contractuel 6 mois multi-produits.",
        BODY
    ))
    story.append(Paragraph("<b>KPI :</b> Nb nouveaux clients soja, Volume soja additionnel (t), CA additionnel (FCFA), Taux de conversion par segment. <b>Délai :</b> 15 jours (urgence).", BODY))

    story.append(Paragraph("7.2 Axe 2 — Défense concentrés (faiblesse structurelle)", H2))
    story.append(Paragraph(
        "<b>Constat :</b> bilan net Q1→Q2 concentrés fragile (+52 M FCFA vs +276 M ciblé). Segment \"retenu\" en déclin -248 t / -164 M FCFA. "
        "Taux churn fidèles 3,1% (vs 1,9% ciblé). 526 clients jamais acquis concentrés.",
        BODY_BOLD
    ))
    story.append(Paragraph(
        "<b>Actions :</b> benchmark prix concurrents (12 références) ; enquête qualitative 30 clients churned ; ajustement tarifaire si nécessaire (C104, C103) ; "
        "visite physique 97 fidèles en déclin ; bundle soja+concentrés obligatoire ; test nouveaux formats (1Kg, 5Kg, 25Kg).",
        BODY
    ))
    story.append(Paragraph("<b>KPI :</b> Volume concentrés Q3 vs Q2, Taux de rétention fidèles, Prix moyen vs concurrents. <b>Délai :</b> 30-60 jours.", BODY))

    story.append(Paragraph("7.3 Axe 3 — Stratégie par agence (disparités majeures)", H2))
    story.append(Paragraph(
        "<b>Constat :</b> disparités importantes — FAMLA (293 clients, 19% zéro achat) vs MESSASSI (167, 33,5% zéro achat) vs NGAOUNDERE (91, 34,1% zéro achat). "
        "1 client sur 3 en zéro achat dans certaines agences = anomalie.",
        BODY_BOLD
    ))
    story.append(Paragraph(
        "<b>Actions :</b> diagnostic terrain MESSASSI + NGAOUNDERE ; renforcement effectif commercial si sous-dimensionné ; visibilité locale (PLV, événements) ; "
        "plan marketing agence par agence ; objectifs commerciaux individualisés.",
        BODY
    ))
    story.append(Paragraph("<b>KPI :</b> % zéro achat par agence (cible <25%), Nb nouveaux clients par agence/mois. <b>Délai :</b> 60-90 jours.", BODY))

    story.append(PageBreak())

    # ---------- PAGE 15: SECTION 7 SUITE ----------
    story.append(Paragraph("7.4 Axe 4 — Risques à surveiller", H2))
    story.append(Paragraph(
        "<b>Rupture concurrente temporaire :</b> risque de retour client si pas de verrouillage. "
        "<b>Stock soja à surveiller</b> pour pouvoir servir la demande. "
        "<b>Cannibalisation possible</b> soja vs concentrés.",
        BODY_BOLD
    ))
    story.append(Paragraph(
        "<b>Actions :</b> suivi quotidien stock soja + plan réappro ; contrats multi-produits 6 mois (verrouillage) ; bundles obligatoires (anti-cannibalisation) ; "
        "veille concurrentielle hebdo ; tableau de bord churn mensuel.",
        BODY
    ))
    story.append(Paragraph("<b>KPI :</b> Niveau stock soja (jours), Taux de rétention nouveaux clients, Mix produit, Marge brute par produit. <b>Délai :</b> continu (revue mensuelle).", BODY))

    story.append(Paragraph("7.5 Axe 5 — Plan 90 jours (3 phases)", H2))
    story.append(Paragraph(
        "<b>Phase 1 (J1-15) — Urgence conquête soja :</b> lancement opération \"Soja disponible\" auprès des 324 clients prioritaires (14+192+118) ; "
        "communication agences (argumentaire sécurité d'approvisionnement) ; vérification quotidienne stocks soja + plan réapprovisionnement.",
        BODY
    ))
    story.append(Paragraph(
        "<b>Phase 2 (J15-30) — Verrouillage :</b> tous les nouveaux clients soja signent un engagement 6 mois multi-produits avec remise dégressive ; "
        "bundles soja + concentrés (remise -5% sur concentrés si achat soja) ; lancement enquête qualitative auprès de 30 clients churned concentrés.",
        BODY
    ))
    story.append(Paragraph(
        "<b>Phase 3 (J30-60) — Reconquête concentrés :</b> benchmark prix concurrents sur les 12 références concentrés ; ajustement tarifaire si nécessaire ; "
        "visite physique des 97 fidèles Q1 en déclin Q2.",
        BODY
    ))
    story.append(Paragraph(
        "<b>Phase 4 (J60-90) — Consolidation :</b> bilan conquest soja (combien de nouveaux clients, quel CA additionnel) ; "
        "plan marketing agences prioritaires (Messassi, Ngaoundéré) ; revue des 14 clients 20/80 zéro achat : combien réactivés, quel CA récupéré.",
        BODY
    ))

    story.append(Paragraph("7.6 Synthèse stratégique", H2))
    story.append(Paragraph(
        "<b>1.</b> La rupture concurrente sur le soja est une <b>aubaine à 6 mois</b> : traiter comme opération de guerre (moyens commerciaux maximaux, prix légèrement premium, mais verrouillage contractuel).",
        BODY_BOLD
    ))
    story.append(Paragraph(
        "<b>2.</b> Le vrai sujet stratégique est la <b>faiblesse des concentrés</b> : si BELGOCAM ne résout pas le déclin de volume du segment \"retenu\" (-248 t en Q2), la rentabilité long-terme est menacée car les concentrés sont plus marginaux que le soja brut.",
        BODY_BOLD
    ))
    story.append(Paragraph(
        "<b>3.</b> Les 14 <b>clients 20/80 en zéro achat Q1</b> sont le test ultime : si on ne les réactive pas pendant cette période de rupture concurrente (où l'argument \"stock disponible\" est imbattable), c'est qu'il y a un problème structurel (prix, qualité, relation client) qu'il faut diagnostiquer urgemment.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ---------- PAGE 16: SECTION 8 - PROJECTION ----------
    story.append(Paragraph("8. Projection CA 6 mois (3 scénarios)", H1))
    story.append(section_divider())

    story.append(Paragraph("8.1 Hypothèses par scénario", H2))
    hyp_data = [
        ["Hypothèse", "🔴 Pessimiste", "🟡 Réaliste", "🟢 Optimiste"],
        ["Contexte", "Rupture courte (1-2 mois)", "Rupture modérée (3-4 mois)", "Rupture prolongée (6+ mois)"],
        ["% clients 20/80 réactivés (sur 14)", "30% (4 clients)", "60% (8 clients)", "90% (12 clients)"],
        ["% clients churned reconquis", "10%", "30%", "50%"],
        ["% clients persistants acquis (sur 118)", "5% (6)", "15% (18)", "30% (35)"],
        ["% pertes Q1 récupérées", "10%", "30%", "60%"],
        ["CA capté sur clients concurrents", "100 M FCFA", "300 M FCFA", "600 M FCFA"],
    ]
    story.append(make_table(hyp_data, col_widths=[5.5*cm, 3.5*cm, 3.5*cm, 3.5*cm]))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("8.2 CA additionnel estimé à 6 mois", H2))
    ca_data = [
        ["Segment", "🔴 Pessimiste", "🟡 Réaliste", "🟢 Optimiste"],
        ["Réactivation 14 clients 20/80", "34 M", "68 M", "103 M"],
        ["Reconquête 192 churned", "27 M", "81 M", "135 M"],
        ["Acquisition 118 persistants zéro achat", "3 M", "16 M", "63 M"],
        ["Récupération pertes Q1", "58 M", "174 M", "348 M"],
        ["Conquête clients concurrents", "100 M", "300 M", "600 M"],
        ["CA ADDITIONNEL TOTAL", "222 M", "640 M", "1 249 M"],
    ]
    story.append(make_table(ca_data, col_widths=[6*cm, 3.3*cm, 3.3*cm, 3.3*cm]))
    story.append(Spacer(1, 0.3*cm))
    # Chart: Projection CA additionnel
    img = Image('/home/z/my-project/scripts/pdf_charts/chart7_projection.png', width=15*cm, height=7.5*cm)
    story.append(img)
    story.append(Paragraph("Figure 8 — Projection CA additionnel à 6 mois par scénario (millions FCFA)", CAPTION))

    story.append(Paragraph("8.3 Projection CA total S2 2026", H2))
    proj_data = [
        ["Indicateur", "Actuel (S1)", "🔴 Pessimiste", "🟡 Réaliste", "🟢 Optimiste"],
        ["CA HT 6 mois", "17,3 Md", "17,3 Md", "17,3 Md", "17,3 Md"],
        ["CA additionnel estimé", "—", "+222 M", "+640 M", "+1 249 M"],
        ["CA projeté S2 2026", "17,3 Md", "17,5 Md", "17,9 Md", "18,5 Md"],
        ["Croissance vs S1 2026", "—", "+1,3%", "+3,7%", "+7,2%"],
    ]
    story.append(make_table(proj_data, col_widths=[5*cm, 3*cm, 3*cm, 3*cm, 3*cm]))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Objectif S2 2026 :</b> +640 M FCFA de CA additionnel (scénario réaliste), soit +3,7% de croissance vs S1 2026. "
        "Cet objectif est atteignable si les actions de conquête soja sont lancées sous 15 jours et si le verrouillage contractuel est efficace.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ---------- PAGE 17: SECTION 9 - VENTES VS OBJECTIFS ----------
    story.append(Paragraph("9. Ventes vs Objectifs 2026", H1))
    story.append(section_divider())

    story.append(Paragraph("9.1 Vue d'ensemble (S1 2026 — Janvier-Juin)", H2))
    story.append(Paragraph(
        "Les objectifs 2026 ont été définis par catégorie de produits et par agence. Sur le semestre 1 (Janvier-Juin 2026), "
        "le volume total atteint est de <b>37 322 tonnes</b>, soit <b>90,5% de l'objectif S1</b> (41 230 tonnes). "
        "L'écart de 3 908 tonnes correspond principalement à la sous-performance des CONCENTRES (écart -3 950 t) et des INGREDIENTS (écart -463 t), "
        "partiellement compensé par le dépassement des TOURTEAUX (+1 183 t) lié à la rupture concurrente sur le soja.",
        BODY
    ))

    story.append(Paragraph("9.2 Analyse par catégorie (Q1 / Q2 / S1)", H2))
    cat_obj_data = [
        ["Catégorie", "Q1 obj. (t)", "Q1 réel (t)", "% Q1", "Q2 obj. (t)", "Q2 réel (t)", "% Q2", "S1 obj. (t)", "S1 réel (t)", "% S1"],
        ["TOURTEAUX", "14 068", "12 096", "86,0%", "12 617", "15 772", "125,0%", "26 685", "27 868", "104,4%"],
        ["CONCENTRES", "6 612", "4 369", "66,1%", "5 930", "4 223", "71,2%", "12 542", "8 592", "68,5%"],
        ["ALIMENT COMPLET", "271", "200", "73,8%", "243", "210", "86,4%", "514", "410", "79,7%"],
        ["INGREDIENTS", "315", "71", "22,6%", "283", "64", "22,5%", "598", "135", "22,6%"],
        ["COMPLEMENT ALIM.", "3", "145", "4845%", "3", "125", "4165%", "6", "270", "4505%"],
        ["PREMIX", "32", "22", "68,1%", "29", "21", "71,5%", "61", "43", "69,7%"],
        ["MATERIEL ELEVAGE", "0", "4", "—", "0", "1", "—", "0", "5", "—"],
        ["Innovations", "435", "0", "0%", "389", "0", "0%", "824", "0", "0%"],
        ["TOTAL", "21 736", "16 906", "77,8%", "19 494", "20 416", "104,7%", "41 230", "37 322", "90,5%"],
    ]
    story.append(make_table(cat_obj_data, col_widths=[3*cm, 1.5*cm, 1.5*cm, 1.2*cm, 1.5*cm, 1.5*cm, 1.2*cm, 1.5*cm, 1.5*cm, 1.2*cm], font_size=7))
    story.append(Spacer(1, 0.3*cm))

    # Chart A: Ventes vs Objectifs S1
    img = Image('/home/z/my-project/scripts/pdf_charts/chartA_ventes_vs_objectifs.png', width=15.5*cm, height=7.7*cm)
    story.append(img)
    story.append(Paragraph("Figure 9 — Ventes réelles vs Objectifs S1 2026 par catégorie (tonnes)", CAPTION))

    story.append(PageBreak())

    # ---------- PAGE 18: SECTION 9 SUITE ----------
    story.append(Paragraph("9.3 Lecture par catégorie — sous-performances", H2))
    story.append(Paragraph(
        "<b>TOURTEAUX (104,4% de l'objectif S1) :</b> seule catégorie majeure à dépasser l'objectif. "
        "Q1 légèrement sous l'objectif (86,0%) mais Q2 en fort dépassement (125,0%) grâce à la rupture concurrente. "
        "C'est la catégorie qui porte le résultat global — sans elle, le S1 serait à seulement 70% de l'objectif. "
        "Pas de sous-performance — mais dépendance excessive à cette catégorie.",
        BODY
    ))

    story.append(Spacer(1, 0.2*cm))
    story.append(icr_block(
        insights=[
            "CONCENTRÉS à 68,5% de l'objectif S1 avec écart de -3 950 tonnes — sous-performance critique et constante (Q1 : 66,1%, Q2 : 71,2%).",
            "Aucune agence n'atteint 100% de son objectif concentrés — problème systémique national.",
            "Le déclin YoY de -0,9% confirme le caractère structurel, pas conjoncturel.",
        ],
        causes=[
            "Positionnement tarifaire potentiellement non compétitif vs concurrents sur les concentrés.",
            "Qualité perçue en baisse ou absence de différenciation perçue par les éleveurs.",
            "Couverture commerciale insuffisante sur les clients moyens (pas seulement les 20/80).",
            "Absence de bundle soja-concentrés — les clients qui viennent pour le soja ne cross-sellent pas vers les concentrés.",
            "Migration des clients du BELGO 10% vers le BELGO 5% (dégradation du mix produit).",
        ],
        recommandations=[
            "Benchmark tarifaire immédiat sur les 12 références concentrés vs concurrents (délai : 30 jours).",
            "Imposer un bundle soja-concentrés sur chaque commande : 3 sacs de soja pour 1 sac de concentré (ratio 3:1 obligatoire).",
            "Mener une enquête qualitative auprès de 30 clients churned concentrés pour comprendre les causes de départ.",
            "Ajustement tarifaire si le benchmark confirme un écart — priorité sur C104 (BELGO 10% Chair) qui représente 50% du volume.",
            "Renforcer le suivi client par le SAV : appels de satisfaction systématiques sur les clients concentrés à risque.",
        ],
    ))

    story.append(Spacer(1, 0.3*cm))
    story.append(icr_block(
        insights=[
            "INGREDIENTS à 22,6% de l'objectif S1 — sous-performance majeure de -463 tonnes.",
            "Effondrement YoY de -65,6% (392 t en S1 2025 vs 135 t en S1 2026).",
            "Les matières premières (lysine, méthionine, farine de poisson, sulfate de fer) sont vendues bien en deçà des objectifs.",
        ],
        causes=[
            "Objectifs potentiellement irréalistes — fixés sans tenir compte de la base réelle de clients acheteurs d'ingrédients.",
            "Perte de parts de marché au profit de concurrents spécialisés en ingrédients (importateurs directs).",
            "Changement de stratégie produit — les clients peuvent s'approvisionner ailleurs pour les matières premières.",
        ],
        recommandations=[
            "Réviser les objectifs INGREDIENTS à la baisse pour les aligner sur la réalité du marché (objectif S1 réel ≈ 135 t vs 598 t prévus).",
            "Analyser les clients 2025 qui n'ont plus acheté d'ingrédients en 2026 — identifier s'ils sont partis chez la concurrence.",
            "Évaluer la rentabilité de cette catégorie — si la marge est faible, réallouer les efforts commerciaux vers les concentrés.",
        ],
    ))

    story.append(Spacer(1, 0.3*cm))
    story.append(icr_block(
        insights=[
            "ALIMENT COMPLET à 79,7% de l'objectif S1 — sous-performance modérée mais stratégique.",
            "Les Booster (Chick + Piglet) sont des produits d'entrée vers les concentrés — leur déficit aggrave la sous-performance des concentrés.",
        ],
        causes=[
            "Manque de promotion sur les gammes Booster auprès des nouveaux éleveurs.",
            "Absence de bundle Booster-concentrés dans l'offre commerciale.",
        ],
        recommandations=[
            "Campagne de promotion Booster auprès des nouveaux clients : échantillon gratuit à la première commande.",
            "Bundle Booster + concentré correspondant (Chick + Chair, Piglet + Porc) avec remise incitative.",
            "Suivi SAV systématique des clients Booster-only (141 clients) pour les convertir aux concentrés.",
        ],
    ))

    story.append(Spacer(1, 0.3*cm))
    story.append(icr_block(
        insights=[
            "PREMIX à 69,7% de l'objectif S1 — sous-performance de 30%.",
            "Innovations à 0% — aucune vente réalisée alors que 824 t étaient prévues.",
        ],
        causes=[
            "Concurrence sur les mélanges techniques (premix personnalisés par les concurrents).",
            "Innovations : produits non lancés ou non référencés dans la base de ventes.",
        ],
        recommandations=[
            "Investiguer la concurrence sur les premix techniques — comparer offres et prix.",
            "Clarifier le statut de la catégorie Innovations : produits en développement ? Lancés mais non référencés ?",
        ],
    ))

    story.append(Paragraph("9.4 Évolution mensuelle vs objectifs (catégories majeures)", H2))
    # Chart B: Évolution mensuelle
    img = Image('/home/z/my-project/scripts/pdf_charts/chartB_evolution_vs_objectifs.png', width=16*cm, height=6*cm)
    story.append(img)
    story.append(Paragraph("Figure 10 — Évolution mensuelle ventes réelles vs objectifs (TOURTEAUX + CONCENTRES)", CAPTION))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "Pour les TOURTEAUX, on observe un décrochage en février-mars (ventes réelles sous objectif) suivi d'un fort rebond à partir d'avril, "
        "qui coïncide avec le début de la rupture concurrente. Le mois de juin atteint 6 930 t (objectif : 4 377 t, soit 158%). "
        "Pour les CONCENTRES, les ventes restent constamment sous l'objectif tout au long du semestre, sans signe d'amélioration — "
        "ce qui confirme le caractère structurel de la sous-performance.",
        BODY
    ))

    story.append(PageBreak())

    # ---------- PAGE 19: SECTION 9 SUITE 2 ----------
    story.append(Paragraph("9.5 Taux d'atteinte par catégorie (Q1, Q2, S1)", H2))
    img = Image('/home/z/my-project/scripts/pdf_charts/chartC_pct_atteinte.png', width=16*cm, height=7.3*cm)
    story.append(img)
    story.append(Paragraph("Figure 11 — Taux d'atteinte des objectifs par catégorie (Q1, Q2, S1)", CAPTION))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "L'analyse des taux d'atteinte révèle deux dynamiques opposées :<br/>"
        "• <b>TOURTEAUX</b> : progression continue Q1→Q2 (86% → 125%), soutenue par la rupture concurrente.<br/>"
        "• <b>CONCENTRES</b> : stagnation (66% → 71%), sous-performance structurelle sans signal d'amélioration.<br/>"
        "• <b>ALIMENT COMPLET</b> : amélioration modeste (74% → 86%), à surveiller.<br/>"
        "• <b>INGREDIENTS</b> : stagnation critique (~22%), problème structurel.<br/>"
        "• <b>PREMIX</b> : stagnation (~70%), à investiguer.",
        BODY
    ))

    story.append(Paragraph("9.6 Analyse par agence (top 5)", H2))
    ag_obj_data = [
        ["Agence", "S1 obj. (t)", "S1 réel (t)", "% S1", "Écart (t)", "Commentaire"],
        ["FAMLA", "10 860", "≈ 9 800", "≈ 90%", "≈ -1 060", "Locomotive — proche objectif"],
        ["NDOBO", "6 651", "≈ 5 900", "≈ 89%", "≈ -750", "Bonne performance"],
        ["MESSASSI", "4 179", "≈ 3 200", "≈ 77%", "≈ -980", "Sous-performante — à diagnostiquer"],
        ["DJELENG", "3 083", "≈ 2 800", "≈ 91%", "≈ -280", "Bonne performance"],
        ["BUEA", "1 830", "≈ 1 500", "≈ 82%", "≈ -330", "Légèrement sous"],
        ["Autres agences", "8 850", "≈ 7 200", "≈ 81%", "≈ -1 650", "Mix hétérogène"],
    ]
    story.append(make_table(ag_obj_data, col_widths=[3*cm, 2.2*cm, 2.2*cm, 1.5*cm, 2*cm, 6.5*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "FAMLA et NDOBO (les deux plus grosses agences) tiennent le cap. MESSASSI sous-performe de 23%, ce qui confirme "
        "le diagnostic de la section 3 (zéro achat critique dans cette agence). Un renforcement commercial y est prioritaire.",
        BODY
    ))

    story.append(PageBreak())

    # ---------- PAGE 20: SECTION 9 BIS - FOCUS CONCENTRÉS ----------
    story.append(Paragraph("9bis. Focus CONCENTRÉS — Cœur de marge", H1))
    story.append(section_divider())

    story.append(Paragraph(
        "Les CONCENTRÉS (BELGO 10% + BELGO 5%) représentent <b>la plus grosse partie de la marge</b> de BELGOCAM. "
        "Avec 8 592 tonnes et 5,73 milliards FCFA de CA HT en S1 2026, ils sont la catégorie stratégique par excellence. "
        "Pourtant, ils n'atteignent que 68,5% de l'objectif S1 et stagnent en YoY (-0,9%). "
        "Cette section propose une analyse approfondie par sous-catégorie, par produit, par agence et dans le temps.",
        BODY
    ))

    story.append(Paragraph("9bis.1 Décomposition par sous-catégorie (S1 2026)", H2))
    conc_subcat_data = [
        ["Sous-catégorie", "Volume S1 (t)", "% vol", "CA HT (M FCFA)", "Prix moyen (FCFA/t)", "YoY S1"],
        ["BELGO 10% Chair", "4 308,0", "50,1%", "2 752,4", "638 908", "-3,2%"],
        ["BELGO 5% Ponte", "1 703,5", "19,8%", "1 230,1", "722 104", "+19,3%"],
        ["BELGO 10% Porc", "1 164,2", "13,5%", "640,9", "550 499", "-19,6%"],
        ["BELGO 5% Chair", "1 369,5", "15,9%", "1 084,1", "791 635", "+6,1%"],
        ["BELGO 10% Ponte", "36,8", "0,4%", "21,6", "586 569", "-27,9%"],
        ["BELGO Rabbit", "10,1", "0,1%", "4,1", "402 963", "Nouveau"],
        ["BELGO Ruminant", "0,1", "0,0%", "0,0", "600 000", "Nouveau"],
        ["TOTAL", "8 592,1", "100%", "5 733,2", "667 263", "-0,9%"],
    ]
    story.append(make_table(conc_subcat_data, col_widths=[3.2*cm, 2*cm, 1.2*cm, 2.2*cm, 2.5*cm, 1.5*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))

    # Chart K: CONCENTRÉS by sub-category
    img = Image('/home/z/my-project/scripts/pdf_charts/chartK_conc_subcat.png', width=16*cm, height=5.5*cm)
    story.append(img)
    story.append(Paragraph("Figure 19 — CONCENTRÉS : volume et CA HT par sous-catégorie (S1 2026)", CAPTION))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Hiérarchie des sous-catégories :</b><br/>"
        "• <b>BELGO 10% Chair</b> (C104) domine avec 50% du volume et 48% du CA. C'est le produit locomotive des concentrés. "
        "Son déclin YoY de -3,2% tire l'ensemble de la catégorie vers le bas.<br/>"
        "• <b>BELGO 5% Ponte</b> (C101) est le seul produit en forte croissance (+19,3% YoY) — bonne dynamique sur l'élevage de ponte.<br/>"
        "• <b>BELGO 10% Porc</b> (C105) subit le plus fort déclin YoY (-19,6%) — alerte rouge sur l'élevage porcin.<br/>"
        "• <b>BELGO 5% Chair</b> (C103) progresse modesteement (+6,1%) — les éleveurs migrent du 10% vers le 5% ?",
        BODY
    ))

    story.append(PageBreak())

    # ---------- PAGE 21: SECTION 9 BIS SUITE ----------
    story.append(Paragraph("9bis.2 Croissance YoY par sous-catégorie", H2))
    img = Image('/home/z/my-project/scripts/pdf_charts/chartL_conc_yoy.png', width=15*cm, height=7.5*cm)
    story.append(img)
    story.append(Paragraph("Figure 20 — CONCENTRÉS : croissance YoY S1 2025 vs S1 2026 par sous-catégorie", CAPTION))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Analyse YoY détaillée :</b><br/>"
        "• <b>BELGO 10% Chair -3,2%</b> : déclin modéré mais inquiétant car c'est 50% du volume. "
        "Perte de 143 tonnes YoY. À surveiller de près — c'est le produit qui fait le plus de marge.<br/>"
        "• <b>BELGO 10% Porc -19,6%</b> : effondrement de 284 tonnes. L'élevage porcin semble en difficulté. "
        "Vérifier si c'est une conjoncture (maladie, prix de l'aliment) ou structurel (concurrence).<br/>"
        "• <b>BELGO 10% Ponte -27,9%</b> : chute de 14 tonnes (base faible). Quasi-disparition de ce produit.<br/>"
        "• <b>BELGO 5% Ponte +19,3%</b> : seule vraie bonne nouvelle. +275 tonnes YoY. "
        "L'élevage de ponte se porte bien et migre vers le BELGO 5%.<br/>"
        "• <b>BELGO 5% Chair +6,1%</b> : croissance de 78 tonnes. Confirmation d'une migration 10%→5% Chair.",
        BODY
    ))

    story.append(Paragraph("9bis.3 Évolution mensuelle 2026", H2))
    img = Image('/home/z/my-project/scripts/pdf_charts/chartM_conc_monthly.png', width=15*cm, height=7.5*cm)
    story.append(img)
    story.append(Paragraph("Figure 21 — CONCENTRÉS : évolution mensuelle 2026 par sous-catégorie", CAPTION))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "L'évolution mensuelle montre une <b>stabilité globale sans tendance haussière</b> : "
        "le BELGO 10% Chair oscille entre 660 et 803 t/mois sans progression nette. "
        "Le BELGO 5% Ponte est le seul à montrer une légère tendance haussière vers juin. "
        "Le BELGO 10% Porc reste stable autour de 190-209 t/mois — pas de signal de redressement.",
        BODY
    ))

    story.append(PageBreak())

    # ---------- PAGE 22: SECTION 9 BIS SUITE 2 ----------
    story.append(Paragraph("9bis.4 Répartition par agence (S1 2026)", H2))
    img = Image('/home/z/my-project/scripts/pdf_charts/chartN_conc_by_agence.png', width=16*cm, height=7.3*cm)
    story.append(img)
    story.append(Paragraph("Figure 22 — CONCENTRÉS : volume S1 2026 par agence (tonnes)", CAPTION))
    story.append(Spacer(1, 0.2*cm))

    conc_agence_data = [
        ["Agence", "Vol S1 (t)", "CA (M FCFA)", "% du total", "Objectif S1 (t)", "% atteinte"],
        ["FAMLA", "2 107,1", "1 465,1", "24,5%", "≈ 3 200", "≈ 66%"],
        ["MESSASSI", "871,7", "569,3", "10,1%", "≈ 1 400", "≈ 62%"],
        ["NDOBO", "791,4", "500,1", "9,2%", "≈ 1 200", "≈ 66%"],
        ["DJELENG", "778,5", "521,3", "9,1%", "≈ 1 100", "≈ 71%"],
        ["MBOUDA", "619,3", "411,6", "7,2%", "≈ 900", "≈ 69%"],
        ["AHALA", "572,7", "386,7", "6,7%", "≈ 800", "≈ 72%"],
        ["Autres (7 agences)", "2 851,4", "1 979,1", "33,2%", "≈ 3 942", "≈ 72%"],
        ["TOTAL", "8 592,1", "5 733,2", "100%", "12 542", "68,5%"],
    ]
    story.append(make_table(conc_agence_data, col_widths=[3*cm, 2*cm, 2.2*cm, 1.5*cm, 2.2*cm, 1.8*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Constat par agence :</b> FAMLA concentre 24,5% du volume concentrés mais n'atteint que 66% de son objectif. "
        "MESSASSI sous-performe à 62% — confirmant les diagnostics précédents. "
        "Aucune agence n'atteint 100% de son objectif concentrés — c'est un problème systémique, pas localisé.",
        BODY_BOLD
    ))

    story.append(Paragraph("9bis.5 Focus produit : BELGO 10% Chair (C104) — la locomotive", H2))
    story.append(Paragraph(
        "Le C104 (BELGO 10% Chair 50Kg) est le produit stratégique numéro 1 de BELGOCAM : "
        "<b>4 292 tonnes et 2,74 milliards FCFA de CA en S1 2026</b>, soit à lui seul 16% du CA total de l'entreprise. "
        "Son prix moyen est de 638 721 FCFA/t. Son déclin YoY de -3,2% représente une perte de 143 tonnes et ~91 millions FCFA.",
        BODY
    ))
    story.append(Paragraph(
        "<b>Recommandation spécifique C104 :</b> ce produit doit faire l'objet d'un plan de sauvegarne dédié. "
        "Les 143 tonnes perdues YoY correspondent à environ 5 700 sacs de 50Kg. "
        "Identifier les clients qui ont arrêté d'acheter le C104 en 2026 vs 2025 et mener une enquête de satisfaction. "
        "Vérifier la concurrence sur le segment Chair 10%.",
        BODY_BOLD
    ))

    story.append(Paragraph("9bis.6 Synthèse stratégique CONCENTRÉS", H2))
    story.append(Paragraph(
        "<b>1. Stagnation globale (-0,9% YoY, 68,5% de l'objectif S1).</b> "
        "Le problème est structurel, pas conjoncturel. La rupture concurrente sur le soja n'a pas bénéficié aux concentrés — "
        "les clients qui viennent pour le soja ne cross-sellent pas vers les concentrés.<br/><br/>"
        "<b>2. Dynamiques contradictoires entre sous-catégories.</b> "
        "Le BELGO 5% Ponte (+19,3%) et le BELGO 5% Chair (+6,1%) progressent, suggérant une migration des clients du 10% vers le 5%. "
        "Mais le 10% Chair (-3,2%) et le 10% Porc (-19,6%) déclinent fortement. "
        "Le mix produit se dégrade — migration vers des produits moins marginaux ?<br/><br/>"
        "<b>3. Aucune agence n'atteint son objectif concentrés.</b> "
        "C'est un problème systémique qui require une action au niveau national : "
        "benchmark tarifaire, qualité perçue, argumentaire commercial, et peut-être une revisite du positionnement produit.<br/><br/>"
        "<b>4. Le C104 (BELGO 10% Chair) est le produit critique.</b> "
        "Avec 4 292 t et 2,74 Md FCFA, il représente à lui seul 16% du CA BELGOCAM. "
        "Sa sauvegarde est prioritaire — un plan dédié doit être mis en place dès le S2 2026.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ---------- PAGE 23: SECTION 10 - FORECAST S2 ----------
    story.append(Paragraph("10. Forecast S2 2026 (3 scénarios)", H1))
    story.append(section_divider())

    story.append(Paragraph("10.1 Méthodologie du forecast", H2))
    story.append(Paragraph(
        "Le forecast S2 (Juillet-Décembre 2026) est construit à partir du taux d'atteinte des objectifs observé en S1, "
        "appliqué aux objectifs S2 (qui tiennent compte de la saisonnalité BELGOCAM — activité plus faible en juillet-août, "
        "rebond à partir d'octobre). Trois scénarios sont proposés :",
        BODY
    ))
    story.append(Paragraph(
        "• <b>🔴 Pessimiste</b> : taux d'atteinte S1 moins 10 points (dégradation, fin rupture concurrente, pas d'action commerciale)<br/>"
        "• <b>🟡 Réaliste</b> : taux d'atteinte S1 maintenu (statu quo, actions partielles)<br/>"
        "• <b>🟢 Optimiste</b> : taux d'atteinte S1 plus 15 points (effet du plan d'action + maintien rupture concurrente)",
        BODY
    ))

    story.append(Paragraph("10.2 Forecast par catégorie (tons)", H2))
    forecast_table = [
        ["Catégorie", "S2 obj. (t)", "🔴 Pessimiste", "🟡 Réaliste", "🟢 Optimiste", "% S1 atteint"],
        ["TOURTEAUX", "26 043", "24 741", "27 215", "30 440", "104,4%"],
        ["CONCENTRES", "12 242", "7 099", "8 385", "10 222", "68,5%"],
        ["ALIMENT COMPLET", "501", "350", "399", "476", "79,7%"],
        ["INGREDIENTS", "585", "74", "132", "220", "22,6%"],
        ["COMPLEMENT ALIM.", "6", "264", "270", "276", "4505%"],
        ["PREMIX", "59", "35", "41", "50", "69,7%"],
        ["TOTAL (6 cat.)", "39 436", "32 563", "36 442", "41 684", "—"],
    ]
    story.append(make_table(forecast_table, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))

    # Chart D: Forecast S2 par catégorie
    img = Image('/home/z/my-project/scripts/pdf_charts/chartD_forecast_s2.png', width=16*cm, height=7.3*cm)
    story.append(img)
    story.append(Paragraph("Figure 12 — Forecast S2 (3 scénarios) vs Objectif, par catégorie majeure", CAPTION))

    story.append(PageBreak())

    # ---------- PAGE 21: SECTION 10 SUITE ----------
    story.append(Paragraph("10.3 Forecast total S2 vs Objectif", H2))
    img = Image('/home/z/my-project/scripts/pdf_charts/chartE_forecast_total.png', width=14*cm, height=7.9*cm)
    story.append(img)
    story.append(Paragraph("Figure 13 — Forecast total S2 2026 vs Objectif (en tonnes)", CAPTION))
    story.append(Spacer(1, 0.3*cm))

    total_data = [
        ["Scénario", "Volume S2 (t)", "vs Objectif S2", "vs S1 réel", "Hypothèses"],
        ["S1 réel (référence)", "37 322", "—", "—", "Base observée Jan-Juin 2026"],
        ["S2 objectif", "39 436", "100,0%", "+5,7%", "Objectif officiel BELGOCAM"],
        ["🔴 S2 Pessimiste", "32 563", "82,6%", "-12,8%", "Fin rupture concurrente, pas d'action"],
        ["🟡 S2 Réaliste", "36 442", "92,4%", "-2,4%", "Statu quo, actions partielles"],
        ["🟢 S2 Optimiste", "41 684", "105,7%", "+11,7%", "Plan d'action + maintien rupture"],
    ]
    story.append(make_table(total_data, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2.5*cm, 5.5*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("10.4 Estimation CA HT S2 (basée sur prix moyen S1)", H2))
    ca_forecast_data = [
        ["Catégorie", "Prix moyen S1 (FCFA/t)", "🔴 Pess. CA (M FCFA)", "🟡 Real. CA (M FCFA)", "🟢 Opt. CA (M FCFA)"],
        ["TOURTEAUX", "324 210", "8 022", "8 824", "9 872"],
        ["CONCENTRES", "667 263", "4 737", "5 595", "6 821"],
        ["ALIMENT COMPLET", "862 215", "302", "344", "410"],
        ["INGREDIENTS", "1 559 956", "115", "206", "343"],
        ["COMPLEMENT ALIM.", "1 331 123", "352", "359", "367"],
        ["PREMIX", "1 852 792", "65", "76", "93"],
        ["TOTAL", "—", "13 593", "15 404", "17 906"],
    ]
    story.append(make_table(ca_forecast_data, col_widths=[3.5*cm, 3.2*cm, 3*cm, 3*cm, 3*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Conclusion forecast (méthode 1 — basée sur objectifs) :</b> le scénario réaliste projette un CA HT S2 2026 de <b>15,4 milliards FCFA</b>, "
        "soit <b>-11% vs S1 réel</b> (17,3 Md). Le scénario optimiste atteindrait 17,9 Md FCFA (+3% vs S1), "
        "à condition que le plan d'action soit pleinement déployé et que la rupture concurrente se maintienne. "
        "Le risque principal reste la sous-performance des CONCENTRES : sans redressement, le scénario pessimiste (-1,7 Md vs S1) est plausible.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ---------- PAGE 22: SECTION 10 BIS - YoY 2025-2026 + FORECAST REFINED ----------
    story.append(Paragraph("10bis. Analyse YoY 2025-2026 et forecast affiné", H1))
    story.append(section_divider())

    story.append(Paragraph("10bis.1 Méthodologie — comparaison Year-over-Year", H2))
    story.append(Paragraph(
        "Les ventes 2025 (Janvier-Décembre, 71 477 lignes après exclusion des clients internes) ont été intégrées pour :<br/>"
        "• Mesurer la croissance YoY (S1 2026 vs S1 2025) par catégorie<br/>"
        "• Extraire la saisonnalité 2025 (notamment le ratio S2/S1)<br/>"
        "• Affiner le forecast S2 2026 en appliquant la croissance YoY observée sur la base saisonnière 2025",
        BODY
    ))
    story.append(Paragraph(
        "<b>Méthode du forecast affiné :</b> S2 2026 = S2 2025 × (1 + croissance YoY S1) ± ajustement scénario.<br/>"
        "• 🔴 Pessimiste : YoY - 10 points (fin rupture concurrente)<br/>"
        "• 🟡 Réaliste : YoY maintenu (statu quo)<br/>"
        "• 🟢 Optimiste : YoY + 10 points (plan d'action + maintien rupture)",
        BODY
    ))

    story.append(Paragraph("10bis.2 Croissance YoY par catégorie (S1 2025 vs S1 2026)", H2))
    yoy_table = [
        ["Catégorie", "S1 2025 (t)", "S1 2026 (t)", "YoY %", "S2 2025 (t)", "Lecture"],
        ["TOURTEAUX", "19 023", "27 868", "+46,5%", "27 358", "Effet rupture concurrente massif"],
        ["CONCENTRES", "8 670", "8 592", "-0,9%", "9 243", "Stagnation inquiétante (cœur de marge)"],
        ["ALIMENT COMPLET", "250", "410", "+64,0%", "316", "Belle progression (Chick/Piglet)"],
        ["INGREDIENTS", "392", "135", "-65,6%", "375", "Effondrement — à investiguer"],
        ["COMPLEMENT ALIM.", "3", "270", "+8225%", "3", "Boom (objectif sous-estimé)"],
        ["PREMIX", "45", "43", "-4,8%", "43", "Léger déclin"],
        ["TOTAL", "28 383", "37 252", "+31,2%", "36 338", "Croissance globale tirée par TOURTEAUX"],
    ]
    story.append(make_table(yoy_table, col_widths=[3.2*cm, 2.3*cm, 2.3*cm, 1.5*cm, 2.3*cm, 5*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))

    # Chart F: YoY
    img = Image('/home/z/my-project/scripts/pdf_charts/chartF_yoy_2025_2026.png', width=16*cm, height=7.3*cm)
    story.append(img)
    story.append(Paragraph("Figure 14 — Comparaison YoY S1 2025 vs S1 2026 par catégorie (tonnes)", CAPTION))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "<b>Insights YoY clés :</b><br/>"
        "• <b>TOURTEAUX +46,5%</b> : la rupture concurrente a généré un gain exceptionnel de 8 845 tonnes en S1 2026.<br/>"
        "• <b>CONCENTRES -0,9%</b> : stagnation quasi parfaite — confirmant le caractère structurel du problème, pas conjoncturel.<br/>"
        "• <b>INGREDIENTS -65,6%</b> : effondrement de 257 tonnes. Vérifier si c'est une perte de marché ou un changement de stratégie produit.<br/>"
        "• <b>ALIMENT COMPLET +64%</b> : les Booster progressent bien — confirmer la synergie avec les concentrés.<br/>"
        "• <b>COMPLEMENT ALIMENTAIRE +8225%</b> : explosion liée à la gamme BELGOFOS/BELGOTOX — objectif 2026 est à revoir à la hausse.",
        BODY
    ))

    story.append(PageBreak())

    # ---------- PAGE 23: SECTION 10 BIS SUITE ----------
    story.append(Paragraph("10bis.3 Évolution mensuelle 2025 (année complète) vs 2026", H2))
    img = Image('/home/z/my-project/scripts/pdf_charts/chartG_evolution_2025_2026.png', width=16*cm, height=6*cm)
    story.append(img)
    story.append(Paragraph("Figure 15 — Évolution mensuelle 2025 (année complète) vs 2026 (S1 réel + S2 forecast réaliste)", CAPTION))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        "L'analyse de la saisonnalité 2025 révèle :<br/>"
        "• <b>TOURTEAUX</b> : forte saisonnalité avec pic en octobre 2025 (7 178 t) — lié aux fêtes de fin d'année. "
        "Le forecast S2 2026 anticiperait un pic similaire, voire supérieur, si la rupture concurrente se maintient.<br/>"
        "• <b>CONCENTRES</b> : activité relativement stable sur 2025 (1 335-1 777 t/mois), sans saisonnalité marquée. "
        "Le forecast S2 2026 reste dans cette fourchette, confirmant la stagnation.",
        BODY
    ))

    story.append(Paragraph("10bis.4 Forecast S2 2026 affiné (basé sur YoY + saisonnalité 2025)", H2))
    img = Image('/home/z/my-project/scripts/pdf_charts/chartH_forecast_refined.png', width=16*cm, height=7.3*cm)
    story.append(img)
    story.append(Paragraph("Figure 16 — Forecast S2 2026 affiné (3 scénarios) vs S2 2025 et Objectif, par catégorie", CAPTION))
    story.append(Spacer(1, 0.2*cm))

    forecast_refined_table = [
        ["Catégorie", "S2 2025 (t)", "YoY S1", "🔴 Pess. (t)", "🟡 Real. (t)", "🟢 Opt. (t)", "S2 obj. (t)"],
        ["TOURTEAUX", "27 358", "+46,5%", "37 342", "40 078", "42 814", "26 043"],
        ["CONCENTRES", "9 243", "-0,9%", "8 236", "9 160", "10 085", "12 242"],
        ["ALIMENT COMPLET", "316", "+64,0%", "487", "519", "550", "501"],
        ["INGREDIENTS", "375", "-65,6%", "92", "129", "167", "585"],
        ["COMPLEMENT ALIM.", "3", "+8225%", "214", "214", "214", "6"],
        ["PREMIX", "43", "-4,8%", "37", "41", "45", "59"],
        ["TOTAL", "37 339", "+31,2%", "46 407", "50 141", "53 875", "40 238"],
    ]
    story.append(make_table(forecast_refined_table, col_widths=[3*cm, 2*cm, 1.5*cm, 2*cm, 2*cm, 2*cm, 2*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph(
        "<b>Confrontation des deux méthodes de forecast :</b><br/>"
        "• Méthode 1 (basée sur objectifs S2 × taux atteinte S1) : S2 réaliste = 36 442 t (92% obj)<br/>"
        "• Méthode 2 (basée sur 2025 + YoY) : S2 réaliste = 50 141 t (125% obj)<br/><br/>"
        "L'écart significatif (14 000 tonnes) s'explique par la croissance YoY exceptionnelle des TOURTEAUX (+46,5%), "
        "que la méthode 1 ne capte pas. <b>La méthode 2 est plus réaliste</b> si la rupture concurrente se maintient en S2.<br/><br/>"
        "⚠️ <b>Risque principal</b> : si la rupture concurrente s'arrête en S2, les TOURTEAUX reviendraient vers les niveaux 2025 (~27 000 t sur S2), "
        "soit un forecast réaliste révisé à 37 000-38 000 t (proche de la méthode 1).",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ---------- PAGE 24: SECTION 10 BIS SUITE 2 - OBJECTIFS RECALIBRÉS ----------
    story.append(Paragraph("10bis.5 Objectifs S2 recalibrés — Confrontation avec le forecast", H2))
    story.append(Paragraph(
        "Les objectifs S2 2026 ont été recalibrés par agence pour tenir compte des performances réelles du S1. "
        "Le total reste stable (~40 239 t) mais la répartition entre catégories a été ajustée :",
        BODY
    ))
    recaled_table = [
        ["Catégorie", "Objectif initial (t)", "Recalibré (t)", "Écart (t)", "Écart %", "Lecture"],
        ["TOURTEAUX", "26 043", "27 511", "+1 468", "+5,6%", "Hausse modérée — intègre rupture concurrente"],
        ["CONCENTRES", "12 242", "11 650", "-592", "-4,8%", "Légère baisse — objectif ambitieux"],
        ["INGREDIENTS", "585", "466", "-119", "-20,3%", "Baisse — ajustement réaliste"],
        ["ALIMENT COMPLET", "501", "532", "+31", "+6,3%", "Légère hausse"],
        ["PREMIX", "59", "69", "+10", "+16,1%", "Hausse modérée"],
        ["COMPLEMENT ALIM.", "6", "3", "-3", "-52,0%", "Toujours sous-estimé (270 t en S1)"],
        ["TOTAL", "40 238", "40 239", "+1", "+0,0%", "Total stable — reallocation interne"],
    ]
    story.append(make_table(recaled_table, col_widths=[3*cm, 2.2*cm, 2.2*cm, 1.5*cm, 1.3*cm, 5.5*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Insight stratégique du recalibrage :</b> contrairement à une simple baisse des objectifs concentrés, "
        "ce recalibrage maintient un objectif CONCENTRÉS ambitieux (11 650 t, soit +25,9% vs S2 2025) tout en ajustant modérément les TOURTEAUX à la hausse (+5,6%). "
        "Le signal envoyé aux équipes commerciales est clair : les CONCENTRÉS restent la priorité, avec un objectif volontariste "
        "qui, s'il est atteint, permettrait de boucler +13% vs 2025 — bien au-dessus de la cible +5-10%.",
        BODY_BOLD
    ))

    story.append(Paragraph("10bis.6 Incidence du recalibrage sur le CA HT", H2))
    img = Image('/home/z/my-project/scripts/pdf_charts/chartO_ca_impact_recaled.png', width=15*cm, height=7.5*cm)
    story.append(img)
    story.append(Paragraph("Figure 17 — Incidence du recalibrage sur le CA HT S2 par catégorie (millions FCFA)", CAPTION))
    story.append(Spacer(1, 0.2*cm))
    ca_impact = [
        ["Catégorie", "CA initial (M FCFA)", "CA recalibré (M FCFA)", "Écart CA (M)", "% écart"],
        ["CONCENTRÉS", "8 169", "7 774", "-395", "-4,8%"],
        ["TOURTEAUX", "8 443", "8 919", "+476", "+5,6%"],
        ["INGREDIENTS", "913", "728", "-185", "-20,3%"],
        ["ALIMENT COMPLET", "432", "459", "+27", "+6,3%"],
        ["PREMIX", "109", "127", "+18", "+16,1%"],
        ["COMPLEMENT ALIM.", "8", "4", "-4", "-52,0%"],
        ["TOTAL", "18 074", "18 010", "-64", "-0,4%"],
    ]
    story.append(make_table(ca_impact, col_widths=[3.2*cm, 2.8*cm, 2.8*cm, 2*cm, 1.5*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Bilan CA du recalibrage :</b> la perte sur les CONCENTRÉS (-395 M FCFA) est plus que compensée par le gain sur les TOURTEAUX (+476 M FCFA). "
        "Le bilan net CONCENTRÉS + TOURTEAUX est <b>positif (+81 M FCFA)</b> — le recalibrage ne détruit pas de CA, il le réalloue. "
        "La perte totale est de seulement 64 M FCFA (-0,4%), principalement due aux INGREDIENTS (-185 M). "
        "Le recalibrage est donc <b>neutre sur le CA global</b> tout en fixant un objectif CONCENTRÉS plus ambitieux.",
        BODY_BOLD
    ))

    story.append(PageBreak())

    # ---------- PAGE 25: SECTION 10 BIS SUITE 3 ----------
    story.append(Paragraph("10bis.7 Forecast YoY vs Objectifs S2 recalibrés", H2))
    img = Image('/home/z/my-project/scripts/pdf_charts/chartI_forecast_vs_recaled.png', width=16*cm, height=7.3*cm)
    story.append(img)
    story.append(Paragraph("Figure 18 — Forecast YoY (réaliste) vs Objectifs S2 initial et recalibré, par catégorie", CAPTION))
    story.append(Spacer(1, 0.2*cm))

    fc_vs_obj = [
        ["Catégorie", "Forecast YoY (t)", "Obj. Recalibré (t)", "Atteinte %", "Lecture"],
        ["TOURTEAUX", "40 078", "27 511", "145,7%", "Dépassement massif si rupture maintenue"],
        ["CONCENTRES", "9 160", "11 650", "78,6%", "Objectif ambitieux — nécessite plan d'action"],
        ["ALIMENT COMPLET", "519", "532", "97,4%", "Quasi-atteinte"],
        ["INGREDIENTS", "129", "466", "27,7%", "Très sous — problème structurel"],
        ["COMPLEMENT ALIM.", "214", "3", "7432%", "Objectif aberrant — à corriger"],
        ["PREMIX", "41", "69", "59,9%", "Sous-performance persistante"],
        ["TOTAL", "50 141", "40 239", "124,6%", "Dépassement global porté par TOURTEAUX"],
    ]
    story.append(make_table(fc_vs_obj, col_widths=[3.2*cm, 2.5*cm, 2.5*cm, 1.5*cm, 6*cm], font_size=8))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Conclusion forecast vs objectifs recalibrés :</b><br/>"
        "• Les CONCENTRÉS ont un objectif ambitieux (11 650 t) que le forecast réaliste (9 160 t) n'atteint qu'à 78,6%. "
        "L'écart de 2 490 t doit être comblé par le plan d'action.<br/>"
        "• Les TOURTEAUX dépasseraient l'objectif recalibré de 46% si la rupture concurrente se maintient.<br/>"
        "• Le total S2 serait à 124,6% des objectifs recalibrés — porté par les TOURTEAUX.<br/>"
        "• ⚠️ <b>Risque</b> : ce dépassement est entièrement porté par les TOURTEAUX. Si la rupture s'arrête, le total retomberait à ~37 000 t (92% de l'objectif recalibré).",
        BODY_BOLD
    ))

    story.append(KeepTogether([
        Paragraph("10bis.8 CONCENTRÉS — Peut-on boucler +5% à +10% vs 2025 ?", H2),
        Image('/home/z/my-project/scripts/pdf_charts/chartP_conc_target_5_10pct.png', width=15*cm, height=7.5*cm),
    ]))
    story.append(Paragraph("Figure 19 — CONCENTRÉS : forecast S2 vs cibles +5% et +10% vs 2025", CAPTION))
    story.append(Spacer(1, 0.2*cm))

    conc_target = [
        ["Scénario / Cible", "S2 2026 (t)", "Total 2026 (t)", "YoY vs 2025", "Cible +5% ?", "Cible +10% ?"],
        ["S2 2025 (référence)", "9 256", "17 926", "—", "—", "—"],
        ["🔴 Forecast pessimiste", "8 236", "16 828", "-6,1%", "Non", "Non"],
        ["🟡 Forecast réaliste", "9 160", "17 752", "-1,0%", "Non", "Non"],
        ["🟢 Forecast optimiste", "10 085", "18 677", "+4,2%", "Non (proche)", "Non"],
        ["Objectif S2 recalibré", "11 650", "20 242", "+13,0%", "Oui", "Oui"],
        ["Cible +5% vs 2025", "10 230", "18 822", "+5,0%", "—", "—"],
        ["Cible +10% vs 2025", "11 127", "19 719", "+10,0%", "—", "—"],
    ]
    story.append(make_table(conc_target, col_widths=[3.5*cm, 2*cm, 2.2*cm, 1.8*cm, 1.5*cm, 1.5*cm], font_size=7))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>Verdict :</b> avec les dynamiques actuelles (stagnation -1% YoY), le scénario réaliste ne permet pas d'atteindre +5%. "
        "Le forecast optimiste (+4,2%) est juste en-dessous de la cible +5%. "
        "Pour boucler +5%, il faut produire 10 230 t en S2 (vs 9 160 t forecast), soit <b>+1 070 t</b> à trouver — "
        "ce qui correspond aux leviers identifiés dans le plan d'action (réactivation 14 clients 20/80 ~600 t, "
        "reconquête 178 churned ~300 t, cross-sell Booster ~200 t, sauvegarde C104 ~143 t = potentiel total ~1 243 t). "
        "<b>La cible +5% est atteignable si le plan d'action est exécuté avec rigueur.</b> "
        "Pour +10%, il manque encore ~723 t après épuisement de ces leviers — des actions supplémentaires seraient nécessaires.",
        BODY_BOLD
    ))

    story.append(Paragraph("10bis.9 Objectifs S2 recalibrés par agence", H2))
    img = Image('/home/z/my-project/scripts/pdf_charts/chartJ_s2_by_agence.png', width=16*cm, height=6.7*cm)
    story.append(img)
    story.append(Paragraph("Figure 20 — Objectifs S2 recalibrés par agence (tonnes, Juillet-Décembre 2026)", CAPTION))

    story.append(PageBreak())

    # ---------- PAGE 25: SECTION 11 - RECOMMANDATIONS ----------
    story.append(Paragraph("11. Recommandations et plan d'action priorisé", H1))
    story.append(section_divider())

    rec_data = [
        ["Priorité", "Axe d'action", "Nb clients", "Délai", "Pilote"],
        ["🔴 CRITIQUE", "Sauvetage des 14 clients 20/80 en clients zéro achat Q1 (ciblé + concentrés)", "14", "30 j", "Directeur Commercial + Responsables agences"],
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

    story.append(Paragraph("11.1 Actions concrètes par recommandation", H2))
    story.append(Paragraph(
        "<b>1. Sauvetage des 14 clients 20/80 (CRITIQUE) :</b> RDV individuel avec chaque client par le directeur commercial. "
        "Offre dédiée : remise volume + échantillon + livraison gratuite. Vérifier s'il y a un problème qualité/prix/concurrence. "
        "Enjeu : 175 M FCFA de perte Q1 estimée sur ces 14 clients. Le SAV prépare le terrain avant chaque RDV en collectant "
        "l'historique des réclamations et le niveau de satisfaction du client.",
        BODY
    ))
    story.append(Paragraph(
        "<b>2. Bundle soja-concentrés obligatoire (CRITIQUE) :</b> chaque commande de tourteaux de soja doit inclure "
        "une fraction de concentré — ratio 3:1 (3 sacs de soja pour 1 sac de concentré). "
        "Cette règle transforme le soja (produit d'appel) en moteur de cross-sell vers les concentrés (cœur de marge). "
        "Les clients qui viennent pour le soja pendant la rupture concurrente repartiront systématiquement avec du concentré. "
        "Mise en œuvre : conditionner la vente de soja à l'achat d'au moins 1 sac de concentré pour 3 sacs de soja. "
        "Remise de -5% sur le concentré dans le cadre du bundle pour inciter.",
        BODY_BOLD
    ))
    story.append(Paragraph(
        "<b>3. Recontact des 192 churned par SAV + commercial (ÉLEVÉE) :</b> le SAV effectue un premier appel "
        "dans les 7 jours pour recueillir le motif de départ (qualité, prix, concurrence, livraison). "
        "Le commercial prend le relais dans les 15 jours avec une offre de retour : -5% sur première commande de réactivation. "
        "Le SAV assure ensuite un suivi mensuel pendant 3 mois pour éviter les rechutes. Enjeu : 273 M FCFA perdus.",
        BODY
    ))
    story.append(Paragraph(
        "<b>4. Sauvetage des 97 fidèles en déclin (ÉLEVÉE) :</b> visite physique par le commercial. "
        "Le SAV identifie en amont les signaux faibles (baisse de fréquence, réclamations récentes, retards de paiement). "
        "Diagnostic : changement d'activité ? Concurrence ? Insatisfaction ? "
        "Mise en place d'un plan de fidélisation sur 3 mois (tarif préférentiel, livraison prioritaire, suivi SAV dédié).",
        BODY
    ))
    story.append(Paragraph(
        "<b>5. Suivi SAV systématique des clients concentrés (ÉLEVÉE) :</b> le SAV met en place un programme "
        "d'appels de satisfaction mensuels sur les 850+ clients actifs concentrés. Chaque appel permet de : "
        "(a) détecter les signaux d'attrition avant qu'ils ne se transforment en churn, "
        "(b) collecter le feedback sur la qualité et les prix, "
        "(c) identifier les opportunités de cross-sell (Booster, compléments). "
        "Le SAV devient le système d'alerte précoce de la Direction Commerciale. "
        "Budget : 2 ETP SAV dédiés aux concentrés pendant 6 mois.",
        BODY_BOLD
    ))
    story.append(Paragraph(
        "<b>6. Prospection des 118 persistants zéro achat (MOYENNE) :</b> campagne d'échantillonnage + invitation à une démonstration produit. "
        "Tarif découverte sur première commande. Identifier ceux qui achètent chez les concurrents (analyse de marché locale).",
        BODY
    ))
    story.append(Paragraph(
        "<b>7. Capitalisation sur les 234 réactivés Q2 (MOYENNE) :</b> enquête qualitative auprès de 30 clients : "
        "qu'est-ce qui a déclenché l'achat Q2 ? (visite commerciale, promo, rupture concurrentielle, nouveauté produit). "
        "Reproduire les leviers efficaces à grande échelle.",
        BODY
    ))
    story.append(Paragraph(
        "<b>8. Motivation de l'équipe commerciale (MOYENNE) :</b> la motivation de l'équipe commerciale est une arme stratégique. "
        "Mise en place d'un système de primes individualisées liées aux ventes de concentrés (et non de soja qui se vend seul). "
        "Chaque commercial se voit fixer un objectif individuel de ventes concentrés avec : "
        "(a) prime mensuelle si objectif atteint, (b) bonus trimestriel sur le cross-sell soja→concentrés, "
        "(c) reconnaissance interne (classement mensuel des meilleurs vendeurs de concentrés). "
        "L'objectif est d'aligner les intérêts des commerciaux sur la marge (concentrés) plutôt que sur le volume (soja). "
        "Formation spécifique sur l'argumentaire concentrés + bundles à organiser dans les 30 jours.",
        BODY_BOLD
    ))
    story.append(Paragraph(
        "<b>9. Analyse déclin volume concentrés (STRUCTURANTE) :</b> le segment \"retenu\" sur concentrés décline en volume malgré la rétention. "
        "Enquête : prix trop élevés vs concurrents ? Qualité perçue en baisse ? Substitution par soja ? "
        "Lancer un atelier interne produit/marché.",
        BODY
    ))

    story.append(PageBreak())

    # ---------- PAGE 18: SECTION 9 SUITE ----------
    story.append(Paragraph("11.2 Synthèse du plan d'action", H2))
    synth_pa = [
        ["Indicateur", "Valeur"],
        ["Enjeu total identifié (pertes Q1 + risques churn)", "≈ 854 M FCFA"],
        ["Gain potentiel (si réactivation complète des 20/80 zéro achat)", "≈ 175 M FCFA"],
        ["Nb total de clients concernés par une action prioritaire", "≈ 524 clients + 850+ suivis SAV"],
        ["Bundle soja-concentrés (ratio 3:1)", "Déploiement immédiat — toutes agences"],
        ["Suivi SAV clients concentrés", "2 ETP dédiés — appels mensuels 850+ clients"],
        ["Motivation équipe commerciale", "Primes sur concentrés + formation bundles (30 j)"],
        ["Horizon de mise en œuvre", "30-90 jours (priorités CRITIQUE et ÉLEVÉE sous 30 jours)"],
        ["Pilotage", "Revue mensuelle CA concentrés + tableau de bord SAV + classement commercial"],
    ]
    story.append(make_table(synth_pa, col_widths=[8*cm, 9*cm]))
    story.append(Spacer(1, 0.5*cm))

    story.append(Paragraph("11.3 Calendrier de mise en œuvre", H2))
    cal_data = [
        ["Période", "Action principale", "Livrable attendu"],
        ["Immédiat", "Bundle soja-concentrés 3:1 obligatoire sur chaque commande", "Règle déployée dans toutes les agences"],
        ["J+7", "Lancement opération \"Soja disponible\" + SAV prépare les 14 RDV 20/80", "RDV planifiés + 1er contact effectué"],
        ["J+15", "Recontact 192 churned par SAV (1er appel) + commercial (offre de retour)", "50% des 192 clients contactés"],
        ["J+30", "Verrouillage contrats 6 mois + formation commerciale bundles", "60% des nouveaux verrouillés + formation complète"],
        ["J+30", "Sauvetage 14 clients 20/80 + 97 fidèles en déclin + lancement primes concentrés", "Premiers retours + système de primes actif"],
        ["J+30", "SAV : lancement du programme d'appels mensuels clients concentrés", "Programme SAV opérationnel (850+ clients)"],
        ["J+45", "Enquête qualitative 30 clients réactivés", "Rapport d'enquête + leviers identifiés"],
        ["J+60", "Benchmark prix concentrés concurrents", "Rapport benchmark + recommandations tarifaires"],
        ["J+90", "Bilan conquest soja + plan marketing agences + bilan primes", "Bilan CA additionnel vs scénario réaliste"],
    ]
    story.append(make_table(cal_data, col_widths=[2.2*cm, 7.5*cm, 8*cm]))

    story.append(PageBreak())

    # ---------- PAGE 19: SECTION 10 - CONCLUSION ----------
    story.append(Paragraph("12. Conclusion et prochaines étapes", H1))
    story.append(section_divider())

    story.append(Paragraph("12.1 Synthèse en 3 points", H2))
    story.append(Paragraph(
        "<b>1. La rupture concurrente sur le soja est une aubaine à 6 mois.</b> "
        "BELGOCAM doit la traiter comme une opération de guerre : moyens commerciaux maximaux, prix légèrement premium, mais surtout verrouillage contractuel des nouveaux clients pour éviter qu'ils ne repartent quand les concurrents se réapprovisionneront. "
        "Le soja représente 40% du volume BELGOCAM, c'est donc le produit d'appel idéal pour conquérir de nouveaux clients et les amener ensuite vers les concentrés (cœur de marge).",
        BODY
    ))
    story.append(Paragraph(
        "<b>2. Le vrai sujet stratégique est la faiblesse des concentrés.</b> "
        "Sans résoudre le déclin de volume du segment \"retenu\" (-248 t en Q2), la rentabilité long-terme est menacée. "
        "Les concentrés BELGO 10% et BELGO 5% sont plus marginaux que le soja brut, et leur déclin progressif (+52 M FCFA seulement de bilan net Q1→Q2 vs +276 M pour le soja) cache une dégradation silencieuse. "
        "Un benchmark prix concurrents et une enquête qualitative clients churned sont indispensables sous 60 jours.",
        BODY
    ))
    story.append(Paragraph(
        "<b>3. Les 14 clients 20/80 en zéro achat Q1 sont le test ultime.</b> "
        "Si BELGOCAM ne les réactive pas pendant cette période de rupture concurrente — où l'argument \"stock disponible\" est imbattable — c'est qu'il y a un problème structurel (prix, qualité, relation client) à diagnostiquer urgemment. "
        "Ces 14 clients (SIGHELO SARL, COMPAGNIE FERMIERE CAMEROUNAISE, TEIKING, FEUDJIO, KENMEGNE, etc.) représentent 175 M FCFA de perte Q1 estimée et doivent faire l'objet d'un RDV direct commercial sous 7 jours.",
        BODY
    ))

    story.append(Paragraph("12.2 Objectif S2 2026", H2))
    story.append(Paragraph(
        "<b>+640 M FCFA de CA additionnel</b> (scénario réaliste), soit <b>+3,7% de croissance vs S1 2026</b>. "
        "Cet objectif suppose : 60% des 14 clients 20/80 réactivés, 30% des 192 churned reconquis, 15% des 118 persistants acquis, "
        "30% des pertes Q1 récupérées, et 300 M FCFA de CA capté sur les clients concurrents en rupture de soja.",
        BODY_BOLD
    ))

    story.append(Paragraph("12.3 Prochaines étapes immédiates", H2))
    story.append(Paragraph(
        "<b>1. Comité de direction J+7 :</b> présentation du rapport, validation du plan d'action 90 jours, allocation des ressources commerciales. "
        "Décision sur l'offre \"Soja disponible\" (tarifs, remises, conditions de verrouillage).",
        BODY
    ))
    story.append(Paragraph(
        "<b>2. Lancement opération \"Soja disponible\" J+15 :</b> communication à toutes les agences, argumentaire commercial, "
        "liste nominative des 324 clients prioritaires (14+192+118) transmise aux commerciaux terrain, "
        "vérification quotidienne des stocks soja + plan de réapprovisionnement.",
        BODY
    ))
    story.append(Paragraph(
        "<b>3. Suivi mensuel des KPI :</b> tableau de bord mensuel retraçant le CA additionnel par segment (réactivés, reconquis, nouveaux), "
        "le taux de verrouillage contractuel, le taux de cross-sell Booster → Concentrés, "
        "et l'évolution du % de zéro achat par agence. Revue en comité de direction le 1er lundi de chaque mois.",
        BODY
    ))

    story.append(Spacer(1, 1*cm))
    story.append(HRFlowable(width="100%", thickness=0.8, color=NAVY, spaceBefore=10, spaceAfter=10))
    story.append(Paragraph(
        "<i>Rapport rédigé par William Francis Fohom, Data Analyst, Administrateur National de Ventes — BELGOCAM SA — Juillet 2026.</i>",
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
        cover_frame = Frame(0, 0, PAGE_W, PAGE_H, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id='cover')
        cover_template = PageTemplate(id='Cover', frames=[cover_frame], onPage=draw_cover)
        # Body page template (with header/footer)
        body_frame = Frame(MARGIN_L, MARGIN_B, CONTENT_W, PAGE_H - MARGIN_T - MARGIN_B, id='body')
        body_template = PageTemplate(id='Body', frames=[body_frame], onPage=draw_body_page)
        self.addPageTemplates([cover_template, body_template])


# ===== BUILD PDF =====
OUT = '/home/z/my-project/download/analyse_complete_belgocam.pdf'

doc = MyDocTemplate(OUT, pagesize=A4)
doc.title = "BELGOCAM SA - Analyse complète Janvier-Juin 2026"
doc.author = "William Francis Fohom — Data Analyst, Administrateur National de Ventes"
doc.subject = "Analyse commerciale et plan d'action stratégique"
doc.creator = "BELGOCAM SA — Direction Commerciale"

story = build_story()

# Use NextPageTemplate to switch from Cover to Body after page 1
final_story = [NextPageTemplate('Body')] + story

doc.build(final_story)

print(f"PDF generated: {OUT}")
print(f"File size: {os.path.getsize(OUT):,} bytes")
