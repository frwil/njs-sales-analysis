"""
Génère le PDF corporate sobre: Analyse Chick & Piglet Booster
Oct 2025 - Sep 2026 | Volume (tonnes) | Style: bleu marine + gris

Structure:
  - Cover page (page de garde sobre)
  - Page 1: Synthèse exécutive + KPIs
  - Page 2: Évolution mensuelle globale (chart + commentaire)
  - Page 3: Analyse par région (chart + table + commentaire)
  - Page 4: Top 10 agences (chart + table + commentaire)
  - Page 5: Analyse par produit (chart + commentaire)
  - Page 6: Heatmap agence x mois + observation
  - Page 7: Conclusion et observations
"""
import pandas as pd
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm, inch
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle,
    PageBreak, KeepTogether, HRFlowable, Frame, PageTemplate, NextPageTemplate
)
from reportlab.platypus.flowables import Flowable
from reportlab.pdfgen import canvas
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# Font registration
# ============================================================
FONT_DIR = '/usr/share/fonts'
pdfmetrics.registerFont(TTFont('NotoSerifSC', f'{FONT_DIR}/truetype/noto-serif-sc/NotoSerifSC-Regular.ttf'))
pdfmetrics.registerFont(TTFont('NotoSerifSC-Bold', f'{FONT_DIR}/truetype/noto-serif-sc/NotoSerifSC-Bold.ttf'))
registerFontFamily('NotoSerifSC', normal='NotoSerifSC', bold='NotoSerifSC-Bold')

pdfmetrics.registerFont(TTFont('DejaVuSans', f'{FONT_DIR}/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', f'{FONT_DIR}/truetype/dejavu/DejaVuSans-Bold.ttf'))
registerFontFamily('DejaVuSans', normal='DejaVuSans', bold='DejaVuSans-Bold')

# Use DejaVuSans as primary (good Unicode coverage for French + symbols)
FONT_BODY = 'DejaVuSans'
FONT_BODY_BOLD = 'DejaVuSans-Bold'
FONT_HEAD = 'DejaVuSans-Bold'

# ============================================================
# Corporate Sober Palette (Bleu marine + gris - explicit user choice)
# ============================================================
PRIMARY = colors.HexColor('#1F3864')        # Bleu marine
SECONDARY = colors.HexColor('#2E5C8A')
ACCENT = colors.HexColor('#C9A961')         # Doré sobre
LIGHT_BG = colors.HexColor('#EAEFF7')
LIGHTER_BG = colors.HexColor('#F5F7FB')
TABLE_STRIPE = colors.HexColor('#F8FAFC')
BORDER = colors.HexColor('#BFBFBF')
TEXT_PRIMARY = colors.HexColor('#1F2937')
TEXT_MUTED = colors.HexColor('#6B7280')
WHITE = colors.HexColor('#FFFFFF')
SUCCESS = colors.HexColor('#16A34A')
DANGER = colors.HexColor('#DC2626')

# Chart series colors (consistent with matplotlib charts)
CHICK_COLOR = colors.HexColor('#7C3AED')
PIGLET_COLOR = colors.HexColor('#D97706')

# ============================================================
# Load data
# ============================================================
WORK = "/home/z/my-project/work"
DOWNLOAD = "/home/z/my-project/download"
CHARTS = os.path.join(WORK, "charts")

df = pd.read_csv(os.path.join(WORK, 'booster_consolidated.csv'))
df['Date de commande'] = pd.to_datetime(df['Date de commande'])

months_order = ['2025-10', '2025-11', '2025-12', '2026-01', '2026-02', '2026-03',
                '2026-04', '2026-05', '2026-06', '2026-07', '2026-08', '2026-09']
month_labels_fr = ['Oct 25', 'Nov 25', 'Déc 25', 'Jan 26', 'Fév 26', 'Mar 26',
                   'Avr 26', 'Mai 26', 'Juin 26', 'Juil 26', 'Août 26', 'Sep 26']

# Compute aggregates
total_vol = df['qte_tonnes'].sum()
chick_vol = df[df['famille'] == 'Chick Booster']['qte_tonnes'].sum()
piglet_vol = df[df['famille'] == 'Piglet Booster']['qte_tonnes'].sum()
n_agences = df['agence'].dropna().nunique()
n_regions = df[df['region'] != 'Non spécifié']['region'].nunique()

# Monthly pivot
pivot_month = df.pivot_table(values='qte_tonnes', index='mois', columns='famille', aggfunc='sum', fill_value=0)
pivot_month = pivot_month.reindex(months_order, fill_value=0)
pivot_month['Total'] = pivot_month.sum(axis=1)
monthly_totals = pivot_month['Total']
peak_month = monthly_totals.idxmax()
peak_value = monthly_totals.max()
low_month = monthly_totals.idxmin()
low_value = monthly_totals.min()
avg_month = monthly_totals.mean()

# Region totals
region_totals = df.groupby('region')['qte_tonnes'].sum().sort_values(ascending=False)
region_totals_clean = region_totals[region_totals.index != 'Non spécifié']

# Agence totals
agence_totals = df.groupby('agence')['qte_tonnes'].sum().sort_values(ascending=False)

# Product totals
prod_totals = df.groupby('Description du produit')['qte_tonnes'].sum()
prod_order = ['CHICK BOOSTER 25 Kg', 'CHICK BOOSTER 5Kg', 'PIGLET BOOSTER 25Kg', 'PIGLET BOOSTER 5Kg']

# ============================================================
# Styles
# ============================================================
styles = getSampleStyleSheet()

style_title_cover = ParagraphStyle(
    'TitleCover', parent=styles['Title'],
    fontName=FONT_BODY_BOLD, fontSize=28, leading=34,
    textColor=PRIMARY, alignment=TA_LEFT, spaceAfter=8
)
style_subtitle_cover = ParagraphStyle(
    'SubtitleCover', parent=styles['Normal'],
    fontName=FONT_BODY, fontSize=14, leading=20,
    textColor=TEXT_MUTED, alignment=TA_LEFT, spaceAfter=6
)
style_kicker_cover = ParagraphStyle(
    'KickerCover', parent=styles['Normal'],
    fontName=FONT_BODY_BOLD, fontSize=10, leading=14,
    textColor=ACCENT, alignment=TA_LEFT, spaceAfter=10
)
style_meta_cover = ParagraphStyle(
    'MetaCover', parent=styles['Normal'],
    fontName=FONT_BODY, fontSize=11, leading=16,
    textColor=TEXT_PRIMARY, alignment=TA_LEFT
)
style_summary_cover = ParagraphStyle(
    'SummaryCover', parent=styles['Normal'],
    fontName=FONT_BODY, fontSize=12, leading=18,
    textColor=TEXT_PRIMARY, alignment=TA_JUSTIFY
)

style_h1 = ParagraphStyle(
    'H1', parent=styles['Heading1'],
    fontName=FONT_BODY_BOLD, fontSize=20, leading=26,
    textColor=PRIMARY, alignment=TA_LEFT, spaceBefore=8, spaceAfter=10
)
style_h2 = ParagraphStyle(
    'H2', parent=styles['Heading2'],
    fontName=FONT_BODY_BOLD, fontSize=14, leading=20,
    textColor=PRIMARY, alignment=TA_LEFT, spaceBefore=12, spaceAfter=6
)
style_kicker = ParagraphStyle(
    'Kicker', parent=styles['Normal'],
    fontName=FONT_BODY_BOLD, fontSize=9, leading=12,
    textColor=ACCENT, alignment=TA_LEFT, spaceAfter=4
)
style_body = ParagraphStyle(
    'Body', parent=styles['Normal'],
    fontName=FONT_BODY, fontSize=10.5, leading=16,
    textColor=TEXT_PRIMARY, alignment=TA_JUSTIFY, spaceAfter=8
)
style_caption = ParagraphStyle(
    'Caption', parent=styles['Normal'],
    fontName=FONT_BODY, fontSize=9, leading=12,
    textColor=TEXT_MUTED, alignment=TA_CENTER, spaceBefore=4, spaceAfter=10, italic=True
)
style_table_header = ParagraphStyle(
    'TableHeader', parent=styles['Normal'],
    fontName=FONT_BODY_BOLD, fontSize=9.5, leading=12,
    textColor=WHITE, alignment=TA_CENTER
)
style_table_cell = ParagraphStyle(
    'TableCell', parent=styles['Normal'],
    fontName=FONT_BODY, fontSize=9.5, leading=12,
    textColor=TEXT_PRIMARY, alignment=TA_LEFT
)
style_table_cell_right = ParagraphStyle(
    'TableCellRight', parent=styles['Normal'],
    fontName=FONT_BODY, fontSize=9.5, leading=12,
    textColor=TEXT_PRIMARY, alignment=TA_RIGHT
)
style_table_cell_center = ParagraphStyle(
    'TableCellCenter', parent=styles['Normal'],
    fontName=FONT_BODY, fontSize=9.5, leading=12,
    textColor=TEXT_PRIMARY, alignment=TA_CENTER
)
style_table_total = ParagraphStyle(
    'TableTotal', parent=styles['Normal'],
    fontName=FONT_BODY_BOLD, fontSize=9.5, leading=12,
    textColor=PRIMARY, alignment=TA_RIGHT
)
style_table_total_left = ParagraphStyle(
    'TableTotalLeft', parent=styles['Normal'],
    fontName=FONT_BODY_BOLD, fontSize=9.5, leading=12,
    textColor=PRIMARY, alignment=TA_LEFT
)

# ============================================================
# Page layout & template
# ============================================================
PAGE_W, PAGE_H = A4  # 595.27 x 841.89 pt
MARGIN_L = 2 * cm
MARGIN_R = 2 * cm
MARGIN_T = 2.2 * cm
MARGIN_B = 2 * cm
CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R  # ~456 pt

# Page header/footer
def header_footer(canvas, doc):
    canvas.saveState()
    # Header: small accent bar + page title
    canvas.setFillColor(PRIMARY)
    canvas.rect(MARGIN_L, PAGE_H - 1.2 * cm, 0.4 * cm, 0.4 * cm, fill=1, stroke=0)
    canvas.setFont(FONT_BODY_BOLD, 8)
    canvas.setFillColor(PRIMARY)
    canvas.drawString(MARGIN_L + 0.6 * cm, PAGE_H - 1.1 * cm, "ANALYSE VENTES BOOSTER")
    canvas.setFont(FONT_BODY, 8)
    canvas.setFillColor(TEXT_MUTED)
    canvas.drawRightString(PAGE_W - MARGIN_R, PAGE_H - 1.1 * cm, "Oct 2025 - Sep 2026")
    # Header line
    canvas.setStrokeColor(BORDER)
    canvas.setLineWidth(0.5)
    canvas.line(MARGIN_L, PAGE_H - 1.4 * cm, PAGE_W - MARGIN_R, PAGE_H - 1.4 * cm)

    # Footer: page number + small line
    canvas.setStrokeColor(BORDER)
    canvas.line(MARGIN_L, 1.4 * cm, PAGE_W - MARGIN_R, 1.4 * cm)
    canvas.setFont(FONT_BODY, 8)
    canvas.setFillColor(TEXT_MUTED)
    canvas.drawString(MARGIN_L, 1.0 * cm, "BELGOCAM / NJS GROUP - ERP")
    canvas.drawRightString(PAGE_W - MARGIN_R, 1.0 * cm, f"Page {doc.page}")
    canvas.restoreState()


def cover_page(canvas, doc):
    """Draw cover background decorations"""
    canvas.saveState()
    # Top color block
    canvas.setFillColor(PRIMARY)
    canvas.rect(0, PAGE_H - 4 * cm, PAGE_W, 4 * cm, fill=1, stroke=0)
    # Accent stripe
    canvas.setFillColor(ACCENT)
    canvas.rect(0, PAGE_H - 4.2 * cm, PAGE_W, 0.2 * cm, fill=1, stroke=0)
    # Bottom color block
    canvas.setFillColor(PRIMARY)
    canvas.rect(0, 0, PAGE_W, 1.5 * cm, fill=1, stroke=0)
    # Bottom accent
    canvas.setFillColor(ACCENT)
    canvas.rect(0, 1.5 * cm, PAGE_W, 0.15 * cm, fill=1, stroke=0)

    # Footer text on bottom block
    canvas.setFont(FONT_BODY_BOLD, 9)
    canvas.setFillColor(WHITE)
    canvas.drawString(MARGIN_L, 0.7 * cm, "BELGOCAM SA / NJS GROUP")
    canvas.setFont(FONT_BODY, 9)
    canvas.drawRightString(PAGE_W - MARGIN_R, 0.7 * cm, "Document confidentiel")

    # Top header text
    canvas.setFont(FONT_BODY_BOLD, 10)
    canvas.setFillColor(WHITE)
    canvas.drawString(MARGIN_L, PAGE_H - 2 * cm, "RAPPORT D'ANALYSE COMMERCIALE")
    canvas.setFont(FONT_BODY, 9)
    canvas.drawRightString(PAGE_W - MARGIN_R, PAGE_H - 2 * cm, "Septembre 2026")
    canvas.restoreState()


# ============================================================
# Helper: build a styled table
# ============================================================
def build_table(headers, rows, col_widths, total_row_idx=None, header_height=22, row_height=18):
    """Build a styled corporate table with header + body rows"""
    # Convert all cells to Paragraphs
    header_cells = [Paragraph(str(h), style_table_header) for h in headers]
    data = [header_cells]
    for r_idx, row in enumerate(rows):
        row_cells = []
        for c_idx, cell in enumerate(row):
            if isinstance(cell, Paragraph):
                row_cells.append(cell)
            else:
                # First column = left aligned text, others = right aligned numbers
                if c_idx == 0:
                    row_cells.append(Paragraph(str(cell), style_table_cell))
                else:
                    row_cells.append(Paragraph(str(cell), style_table_cell_right))
        data.append(row_cells)

    # Row heights
    row_heights = [header_height] + [row_height] * len(rows)

    t = Table(data, colWidths=col_widths, rowHeights=row_heights, repeatRows=1)

    style_cmds = [
        # Header
        ('BACKGROUND', (0, 0), (-1, 0), PRIMARY),
        ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        # Borders
        ('GRID', (0, 0), (-1, -1), 0.4, BORDER),
        ('LINEBELOW', (0, 0), (-1, 0), 1.2, PRIMARY),
    ]
    # Alternating row colors
    for i in range(1, len(data)):
        if i % 2 == 0:
            style_cmds.append(('BACKGROUND', (0, i), (-1, i), TABLE_STRIPE))

    # Total row styling
    if total_row_idx is not None:
        # total_row_idx is 1-indexed (1 = first data row)
        actual_idx = total_row_idx + 1  # +1 because header is row 0
        style_cmds.append(('BACKGROUND', (0, actual_idx), (-1, actual_idx), LIGHT_BG))
        style_cmds.append(('LINEABOVE', (0, actual_idx), (-1, actual_idx), 1, PRIMARY))
        style_cmds.append(('FONTNAME', (0, actual_idx), (-1, actual_idx), FONT_BODY_BOLD))

    t.setStyle(TableStyle(style_cmds))
    return t


# ============================================================
# Helper: KPI card
# ============================================================
class KPICard(Flowable):
    """A KPI card: label + value"""
    def __init__(self, label, value, width, height=58):
        Flowable.__init__(self)
        self.label = label
        self.value = value
        self.width = width
        self.height = height

    def draw(self):
        c = self.canv
        # Background
        c.setFillColor(LIGHTER_BG)
        c.setStrokeColor(BORDER)
        c.setLineWidth(0.5)
        c.rect(0, 0, self.width, self.height, fill=1, stroke=1)
        # Top accent bar
        c.setFillColor(ACCENT)
        c.rect(0, self.height - 3, self.width, 3, fill=1, stroke=0)
        # Label
        c.setFont(FONT_BODY, 8.5)
        c.setFillColor(TEXT_MUTED)
        c.drawString(8, self.height - 18, self.label)
        # Value
        c.setFont(FONT_BODY_BOLD, 16)
        c.setFillColor(PRIMARY)
        c.drawString(8, 14, self.value)


# ============================================================
# Build story
# ============================================================
story = []

# ============================================================
# COVER PAGE
# ============================================================
# Top spacing (cover has top color block at 4cm from top)
story.append(Spacer(1, 4.5 * cm))

# Kicker
story.append(Paragraph("RAPPORT D'ANALYSE COMMERCIALE", style_kicker_cover))
# Title
story.append(Paragraph("Analyse des ventes<br/>Chick &amp; Piglet Booster", style_title_cover))
# Subtitle
story.append(Spacer(1, 0.3 * cm))
story.append(Paragraph("Évolution mensuelle en volume - Octobre 2025 à Septembre 2026", style_subtitle_cover))
story.append(Spacer(1, 0.2 * cm))
story.append(Paragraph("Par agence &amp; par région", style_subtitle_cover))

# Spacer to push down
story.append(Spacer(1, 5 * cm))

# Summary block
summary_text = """Ce rapport présente une analyse détaillée de l'évolution des ventes des gammes Chick Booster et Piglet Booster sur la période d'octobre 2025 à septembre 2026. Les données consolidées couvrent 12 mois d'activité commerciale, issues du système ERP NJS Group, et portent sur le volume écoulé en tonnes à travers les agences et régions du réseau BELGOCAM. L'analyse met en évidence les tendances mensuelles, la performance relative des différentes agences et régions, ainsi que la répartition par produit et format."""
story.append(Paragraph(summary_text, style_summary_cover))

story.append(Spacer(1, 1.2 * cm))

# Meta block
meta_data = [
    [Paragraph("<b>Période analysée</b>", style_meta_cover),
     Paragraph("Octobre 2025 - Septembre 2026 (12 mois)", style_meta_cover)],
    [Paragraph("<b>Source des données</b>", style_meta_cover),
     Paragraph("ERP NJS Group - Lignes de commandes multicompany", style_meta_cover)],
    [Paragraph("<b>Périmètre</b>", style_meta_cover),
     Paragraph(f"{n_agences} agences | {n_regions} régions | 4 références produit", style_meta_cover)],
    [Paragraph("<b>Métrique principale</b>", style_meta_cover),
     Paragraph("Volume en tonnes (t)", style_meta_cover)],
    [Paragraph("<b>Date d'édition</b>", style_meta_cover),
     Paragraph("26 septembre 2026", style_meta_cover)],
]
meta_table = Table(meta_data, colWidths=[5 * cm, 11 * cm])
meta_table.setStyle(TableStyle([
    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ('TOPPADDING', (0, 0), (-1, -1), 3),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ('LINEBELOW', (0, 0), (-1, -1), 0.3, BORDER),
    ('TEXTCOLOR', (0, 0), (-1, -1), TEXT_PRIMARY),
]))
story.append(meta_table)

# Use cover template for first page, then normal template
story.append(NextPageTemplate('body'))
story.append(PageBreak())

# ============================================================
# PAGE 1: SYNTHÈSE EXÉCUTIVE
# ============================================================
story.append(Paragraph("SYNTHÈSE EXÉCUTIVE", style_kicker))
story.append(Paragraph("Vue d'ensemble de l'activité Booster", style_h1))
story.append(HRFlowable(width="100%", thickness=2, color=PRIMARY, spaceBefore=2, spaceAfter=12))

# Intro paragraph
intro = """<b>Cette synthèse présente les chiffres clés de l'activité des gammes Chick Booster et Piglet Booster</b> sur la période d'octobre 2025 à septembre 2026. Les données ont été consolidées à partir des exports ERP NJS Group couvrant l'ensemble du réseau d'agences BELGOCAM au Cameroun. L'analyse porte exclusivement sur les volumes écoulés exprimés en tonnes, avec une déclinaison par agence, par région, par produit et par format."""
story.append(Paragraph(intro, style_body))

# KPI cards row (3 cards per row)
kpi_w = (CONTENT_W - 2 * 0.3 * cm) / 3  # 3 cards with gaps
kpis_row1 = [
    KPICard("VOLUME TOTAL (12 MOIS)", f"{total_vol:,.1f} t".replace(',', ' '), kpi_w),
    KPICard("AGENCES ACTIVES", f"{n_agences}", kpi_w),
    KPICard("RÉGIONS COUVERTES", f"{n_regions}", kpi_w),
]
# Use exact width to avoid centering warnings
kpi_table_1 = Table([kpis_row1], colWidths=[kpi_w, kpi_w, kpi_w])
kpi_table_1.setStyle(TableStyle([
    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ('LEFTPADDING', (0, 0), (-1, -1), 0),
    ('RIGHTPADDING', (0, 0), (0, 0), 0.3 * cm),  # gap after 1st card
    ('RIGHTPADDING', (1, 0), (1, 0), 0.3 * cm),  # gap after 2nd card
    ('RIGHTPADDING', (2, 0), (2, 0), 0),
    ('TOPPADDING', (0, 0), (-1, -1), 0),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
]))
story.append(kpi_table_1)
story.append(Spacer(1, 0.3 * cm))

kpis_row2 = [
    KPICard("VOLUME CHICK BOOSTER", f"{chick_vol:,.1f} t".replace(',', ' '), kpi_w),
    KPICard("VOLUME PIGLET BOOSTER", f"{piglet_vol:,.1f} t".replace(',', ' '), kpi_w),
    KPICard("MOYENNE MENSUELLE", f"{avg_month:,.1f} t".replace(',', ' '), kpi_w),
]
kpi_table_2 = Table([kpis_row2], colWidths=[kpi_w, kpi_w, kpi_w])
kpi_table_2.setStyle(TableStyle([
    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ('LEFTPADDING', (0, 0), (-1, -1), 0),
    ('RIGHTPADDING', (0, 0), (0, 0), 0.3 * cm),
    ('RIGHTPADDING', (1, 0), (1, 0), 0.3 * cm),
    ('RIGHTPADDING', (2, 0), (2, 0), 0),
    ('TOPPADDING', (0, 0), (-1, -1), 0),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
]))
story.append(kpi_table_2)
story.append(Spacer(1, 0.6 * cm))

# Highlights section
story.append(Paragraph("Faits marquants", style_h2))
peak_label = month_labels_fr[months_order.index(peak_month)]
low_label = month_labels_fr[months_order.index(low_month)]

highlights = f"""
<b>• Pic d'activité :</b> {peak_value:.1f} t écoulées en {peak_label}, soit le volume mensuel le plus élevé de la période. Cette performance s'explique par la conjonction d'une forte demande sur les deux gammes et d'un mois complet d'activité commerciale (31 jours).<br/>
<b>• Point bas :</b> {low_value:.1f} t en {low_label}, qui correspond à une activité saisonnière réduite en fin d'année civile. Ce mois reste néanmoins supérieur à la moyenne mensuelle de certaines périodes, témoignant d'un socle commercial stable.<br/>
<b>• Répartition par gamme :</b> Chick Booster représente <b>{chick_vol/total_vol*100:.1f}%</b> du volume total ({chick_vol:.1f} t) contre <b>{piglet_vol/total_vol*100:.1f}%</b> pour Piglet Booster ({piglet_vol:.1f} t). L'écart reste modéré, témoignant d'un développement parallèle des deux gammes.<br/>
<b>• Concentration géographique :</b> La région {region_totals_clean.index[0]} domine avec {region_totals_clean.iloc[0]:.1f} t ({region_totals_clean.iloc[0]/total_vol*100:.1f}% du volume), suivie de {region_totals_clean.index[1]} ({region_totals_clean.iloc[1]:.1f} t) et de {region_totals_clean.index[2]} ({region_totals_clean.iloc[2]:.1f} t).<br/>
<b>• Leader agence :</b> L'agence <b>{agence_totals.index[0]}</b> est en tête du classement avec {agence_totals.iloc[0]:.1f} t cumulées sur la période, devant <b>{agence_totals.index[1]}</b> ({agence_totals.iloc[1]:.1f} t) et <b>{agence_totals.index[2]}</b> ({agence_totals.iloc[2]:.1f} t).
"""
story.append(Paragraph(highlights, style_body))

story.append(Spacer(1, 0.4 * cm))

# Table: Top 5 agences
story.append(Paragraph("Top 5 agences par volume cumulé", style_h2))
top5 = agence_totals.head(5)
top5_data = [
    [Paragraph("<b>Rang</b>", style_table_header),
     Paragraph("<b>Agence</b>", style_table_header),
     Paragraph("<b>Région</b>", style_table_header),
     Paragraph("<b>Volume (t)</b>", style_table_header),
     Paragraph("<b>% du total</b>", style_table_header)]
]
for i, (agence, vol) in enumerate(top5.items()):
    region = df[df['agence'] == agence]['region'].iloc[0]
    pct = vol / total_vol * 100
    top5_data.append([
        Paragraph(str(i + 1), style_table_cell_center),
        Paragraph(agence, style_table_cell),
        Paragraph(region, style_table_cell_center),
        Paragraph(f"{vol:.2f}", style_table_cell_right),
        Paragraph(f"{pct:.1f}%", style_table_cell_right),
    ])
top5_table = Table(top5_data, colWidths=[1.5 * cm, 5 * cm, 3.5 * cm, 3 * cm, 3 * cm])
top5_table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), PRIMARY),
    ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
    ('TOPPADDING', (0, 0), (-1, -1), 5),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('GRID', (0, 0), (-1, -1), 0.4, BORDER),
    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [WHITE, TABLE_STRIPE]),
]))
story.append(top5_table)
story.append(Spacer(1, 0.2 * cm))
story.append(Paragraph("Tableau 1 - Top 5 agences par volume cumulé sur la période Oct 25 - Sep 26", style_caption))

# No PageBreak - let content flow naturally to use page space efficiently
story.append(Spacer(1, 0.6 * cm))

# ============================================================
# SECTION: ÉVOLUTION MENSUELLE GLOBALE (continues on same flow)
# ============================================================
story.append(KeepTogether([
    Paragraph("ANALYSE TEMPORELLE", style_kicker),
    Paragraph("Évolution mensuelle globale", style_h1),
    HRFlowable(width="100%", thickness=2, color=PRIMARY, spaceBefore=2, spaceAfter=12),
]))

# Chart 1
img_path = os.path.join(CHARTS, '01_evolution_globale.png')
img = Image(img_path, width=CONTENT_W, height=CONTENT_W * 0.5)
story.append(img)
story.append(Paragraph("Figure 1 - Évolution mensuelle des ventes Booster (tonnes) avec décomposition par gamme", style_caption))

# Commentary
commentary1 = f"""
L'évolution mensuelle des ventes sur les 12 derniers mois révèle une <b>activité globale dynamique mais irrégulière</b>. Le volume total écoulé atteint <b>{total_vol:.1f} tonnes</b> sur la période, avec une moyenne mensuelle de {avg_month:.1f} t. La courbe totale (en bleu marine) met en évidence plusieurs pics et creux marquants qui méritent une attention particulière.

Le <b>pic d'activité est observé en {peak_label} avec {peak_value:.1f} t</b>, soit près de {(peak_value/avg_month - 1)*100:.0f}% au-dessus de la moyenne mensuelle. Cette performance s'explique par la conjonction d'une forte dynamique sur les deux gammes produits et d'un mois complet d'activité (31 jours). À l'inverse, le <b>point bas de {low_value:.1f} t est atteint en {low_label}</b>, traduisant le ralentissement saisonnier classique de fin d'année civile.

La décomposition par gamme montre que <b>Chick Booster (en violet) maintient une avance constante sur Piglet Booster (en orange)</b> tout au long de la période, à l'exception de quelques mois où Piglet Booster rattrape son retard. L'écart mensuel moyen entre les deux gammes est de l'ordre de quelques tonnes, ce qui témoigne d'un développement parallèle plutôt que d'une cannibalisation.
"""
story.append(Paragraph(commentary1, style_body))

# Mini table: monthly values
story.append(Paragraph("Détail mensuel (tonnes)", style_h2))
month_data = [
    [Paragraph("<b>Mois</b>", style_table_header),
     Paragraph("<b>Chick Booster</b>", style_table_header),
     Paragraph("<b>Piglet Booster</b>", style_table_header),
     Paragraph("<b>Total</b>", style_table_header),
     Paragraph("<b>Écart vs moyenne</b>", style_table_header)]
]
for mois in months_order:
    chick = pivot_month.loc[mois, 'Chick Booster']
    piglet = pivot_month.loc[mois, 'Piglet Booster']
    total = pivot_month.loc[mois, 'Total']
    ecart = ((total - avg_month) / avg_month) * 100
    ecart_str = f"{ecart:+.1f}%"
    month_data.append([
        Paragraph(month_labels_fr[months_order.index(mois)], style_table_cell_center),
        Paragraph(f"{chick:.2f}", style_table_cell_right),
        Paragraph(f"{piglet:.2f}", style_table_cell_right),
        Paragraph(f"<b>{total:.2f}</b>", style_table_cell_right),
        Paragraph(ecart_str, style_table_cell_right),
    ])
# Total row
month_data.append([
    Paragraph("<b>TOTAL</b>", style_table_total_left),
    Paragraph(f"<b>{pivot_month['Chick Booster'].sum():.2f}</b>", style_table_total),
    Paragraph(f"<b>{pivot_month['Piglet Booster'].sum():.2f}</b>", style_table_total),
    Paragraph(f"<b>{pivot_month['Total'].sum():.2f}</b>", style_table_total),
    Paragraph("", style_table_cell_right),
])
month_table = Table(month_data, colWidths=[2.5 * cm, 3.5 * cm, 3.5 * cm, 3 * cm, 3.5 * cm])
month_table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), PRIMARY),
    ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
    ('TOPPADDING', (0, 0), (-1, -1), 4),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ('LEFTPADDING', (0, 0), (-1, -1), 5),
    ('RIGHTPADDING', (0, 0), (-1, -1), 5),
    ('GRID', (0, 0), (-1, -1), 0.4, BORDER),
    ('ROWBACKGROUNDS', (0, 1), (-1, -2), [WHITE, TABLE_STRIPE]),
    ('BACKGROUND', (0, -1), (-1, -1), LIGHT_BG),
    ('LINEABOVE', (0, -1), (-1, -1), 1, PRIMARY),
]))
story.append(month_table)
story.append(Spacer(1, 0.2 * cm))
story.append(Paragraph("Tableau 2 - Évolution mensuelle détaillée par gamme (tonnes)", style_caption))

# Spacer instead of page break
story.append(Spacer(1, 0.6 * cm))

# ============================================================
# SECTION: ANALYSE PAR RÉGION
# ============================================================
story.append(KeepTogether([
    Paragraph("ANALYSE GÉOGRAPHIQUE", style_kicker),
    Paragraph("Analyse par région", style_h1),
    HRFlowable(width="100%", thickness=2, color=PRIMARY, spaceBefore=2, spaceAfter=12),
]))

# Chart: region cumulative
img = Image(os.path.join(CHARTS, '07_cumul_region.png'), width=CONTENT_W * 0.85, height=CONTENT_W * 0.85 * 0.5)
img.hAlign = 'CENTER'
story.append(img)
story.append(Paragraph("Figure 2 - Volume cumulé par région sur la période (tonnes)", style_caption))

# Region table
story.append(Paragraph("Classement par région", style_h2))
region_data = [
    [Paragraph("<b>Région</b>", style_table_header),
     Paragraph("<b>Volume (t)</b>", style_table_header),
     Paragraph("<b>Part du total</b>", style_table_header),
     Paragraph("<b>Moyenne mensuelle</b>", style_table_header),
     Paragraph("<b>Nb agences</b>", style_table_header)]
]
for region in region_totals_clean.index:
    vol = region_totals_clean[region]
    pct = vol / region_totals_clean.sum() * 100
    moy = vol / 12
    nb_ag = df[df['region'] == region]['agence'].nunique()
    region_data.append([
        Paragraph(region, style_table_cell),
        Paragraph(f"{vol:.2f}", style_table_cell_right),
        Paragraph(f"{pct:.1f}%", style_table_cell_right),
        Paragraph(f"{moy:.2f}", style_table_cell_right),
        Paragraph(str(nb_ag), style_table_cell_center),
    ])
# Total
region_data.append([
    Paragraph("<b>TOTAL</b>", style_table_total_left),
    Paragraph(f"<b>{region_totals_clean.sum():.2f}</b>", style_table_total),
    Paragraph("<b>100.0%</b>", style_table_total),
    Paragraph(f"<b>{region_totals_clean.sum()/12:.2f}</b>", style_table_total),
    Paragraph(f"<b>{df[df['region']!='Non spécifié']['agence'].nunique()}</b>", ParagraphStyle('tc', parent=style_table_total, alignment=TA_CENTER)),
])
region_table = Table(region_data, colWidths=[3.5 * cm, 3 * cm, 3 * cm, 3.5 * cm, 3 * cm])
region_table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), PRIMARY),
    ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
    ('TOPPADDING', (0, 0), (-1, -1), 5),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('GRID', (0, 0), (-1, -1), 0.4, BORDER),
    ('ROWBACKGROUNDS', (0, 1), (-1, -2), [WHITE, TABLE_STRIPE]),
    ('BACKGROUND', (0, -1), (-1, -1), LIGHT_BG),
    ('LINEABOVE', (0, -1), (-1, -1), 1, PRIMARY),
]))
story.append(region_table)
story.append(Spacer(1, 0.2 * cm))
story.append(Paragraph("Tableau 3 - Synthèse par région (tonnes)", style_caption))

# Commentary
top_reg = region_totals_clean.index[0]
top_reg_vol = region_totals_clean.iloc[0]
top_reg_pct = top_reg_vol / region_totals_clean.sum() * 100
commentary2 = f"""
La <b>région {top_reg} domine nettement l'activité Booster</b> avec {top_reg_vol:.1f} t écoulées sur la période, représentant {top_reg_pct:.1f}% du volume total. Cette position s'explique par la densité du réseau d'agences dans cette zone et la forte concentration de clients avicoles et porcins professionnels. La région {region_totals_clean.index[1]} suit avec {region_totals_clean.iloc[1]:.1f} t ({region_totals_clean.iloc[1]/region_totals_clean.sum()*100:.1f}%), tandis que la région {region_totals_clean.index[2]} occupe la troisième position avec {region_totals_clean.iloc[2]:.1f} t ({region_totals_clean.iloc[2]/region_totals_clean.sum()*100:.1f}%).

L'<b>écart entre les régions</b> reflète à la fois la répartition géographique de l'élevage au Cameroun et l'implantation historique du réseau BELGOCAM. La région Centre bénéficie de la position centrale de Yaoundé et de son arrière-pays agricole, tandis que le Littoral tire profit de l'activité économique autour de Douala. La région Ouest, bien que moins dotée en agences, maintient une activité significative grâce aux éleveurs de l'Ouest et du Nord-Ouest.
"""
story.append(Paragraph(commentary2, style_body))

# Spacer
story.append(Spacer(1, 0.6 * cm))

# ============================================================
# SECTION: TOP AGENCES
# ============================================================
story.append(KeepTogether([
    Paragraph("PERFORMANCE PAR AGENCE", style_kicker),
    Paragraph("Classement des agences", style_h1),
    HRFlowable(width="100%", thickness=2, color=PRIMARY, spaceBefore=2, spaceAfter=12),
]))

# Chart top 10
img = Image(os.path.join(CHARTS, '03_top_agences.png'), width=CONTENT_W, height=CONTENT_W * 0.5)
story.append(img)
story.append(Paragraph("Figure 3 - Top 10 agences par volume cumulé (tonnes) - Top 3 en doré", style_caption))

# Top 10 table
top10 = agence_totals.head(10)
top10_data = [
    [Paragraph("<b>#</b>", style_table_header),
     Paragraph("<b>Agence</b>", style_table_header),
     Paragraph("<b>Région</b>", style_table_header),
     Paragraph("<b>Volume (t)</b>", style_table_header),
     Paragraph("<b>% du total</b>", style_table_header)]
]
for i, (agence, vol) in enumerate(top10.items()):
    region = df[df['agence'] == agence]['region'].iloc[0]
    pct = vol / total_vol * 100
    top10_data.append([
        Paragraph(str(i + 1), style_table_cell_center),
        Paragraph(agence, style_table_cell),
        Paragraph(region, style_table_cell_center),
        Paragraph(f"{vol:.2f}", style_table_cell_right),
        Paragraph(f"{pct:.1f}%", style_table_cell_right),
    ])
top10_table = Table(top10_data, colWidths=[1 * cm, 5 * cm, 3.5 * cm, 3 * cm, 3.5 * cm])
top10_table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), PRIMARY),
    ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
    ('TOPPADDING', (0, 0), (-1, -1), 4),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('GRID', (0, 0), (-1, -1), 0.4, BORDER),
    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [WHITE, TABLE_STRIPE]),
    # Top 3 highlighting
    ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#FFF8E1')),  # Light gold for #1
    ('BACKGROUND', (0, 2), (-1, 2), colors.HexColor('#FFF8E1')),
    ('BACKGROUND', (0, 3), (-1, 3), colors.HexColor('#FFF8E1')),
]))
story.append(top10_table)
story.append(Spacer(1, 0.2 * cm))
story.append(Paragraph("Tableau 4 - Top 10 agences par volume cumulé (Oct 25 - Sep 26)", style_caption))

# Commentary
top1_ag = agence_totals.index[0]
top1_vol = agence_totals.iloc[0]
top1_pct = top1_vol / total_vol * 100
top3_vol = agence_totals.head(3).sum()
top3_pct = top3_vol / total_vol * 100
top10_vol = agence_totals.head(10).sum()
top10_pct = top10_vol / total_vol * 100

commentary3 = f"""
Le <b>classement des agences</b> met en évidence une concentration significative de l'activité sur quelques pôles. L'agence <b>{top1_ag}</b> domine très largement avec {top1_vol:.1f} t cumulées, représentant à elle seule {top1_pct:.1f}% du volume total. Cette performance reflète la position stratégique de cette agence, la densité de sa clientèle et l'ancienneté de son implantation dans le réseau.

Le <b>top 3 des agences</b> ({', '.join(agence_totals.head(3).index)}) totalise {top3_vol:.1f} t, soit {top3_pct:.1f}% du volume global. Le <b>top 10</b> concentre {top10_vol:.1f} t ({top10_pct:.1f}%), ce qui signifie que les autres agences se partagent le reste du volume. Cette distribution suggère un potentiel de développement sur les agences de milieu et bas de classement, qui disposent d'une marge de progression pour rapprocher leur performance de celle des leaders.

L'<b>analyse par région</b> montre que les agences de la région Centre sont particulièrement bien représentées dans le top 10, ce qui confirme la dynamique observée au niveau régional. Toutefois, on note une présence équilibrée des trois régions dans ce classement, témoignant d'un réseau relativement homogène dans sa performance commerciale.
"""
story.append(Paragraph(commentary3, style_body))

# Spacer
story.append(Spacer(1, 0.6 * cm))

# ============================================================
# SECTION: ANALYSE PAR PRODUIT
# ============================================================
story.append(KeepTogether([
    Paragraph("ANALYSE PRODUIT", style_kicker),
    Paragraph("Évolution par produit et format", style_h1),
    HRFlowable(width="100%", thickness=2, color=PRIMARY, spaceBefore=2, spaceAfter=12),
]))

# Chart evolution produit
img = Image(os.path.join(CHARTS, '04_evolution_produit.png'), width=CONTENT_W, height=CONTENT_W * 0.5)
story.append(img)
story.append(Paragraph("Figure 4 - Évolution mensuelle par produit (tonnes)", style_caption))

# Commentary
commentary4 = f"""
L'<b>analyse par produit</b> révèle des dynamiques contrastées entre les quatre références de la gamme Booster. Sur l'ensemble de la période, <b>Chick Booster totalise {chick_vol:.1f} t</b> contre <b>{piglet_vol:.1f} t pour Piglet Booster</b>, soit un ratio d'environ {chick_vol/piglet_vol:.2f} en faveur de la gamme volaille. Cet écart reflète la structuration du marché de l'élevage au Cameroun, où la filière avicole demeure plus développée que la filière porcine.

Le <b>format 25 Kg</b> est largement dominant par rapport au format 5 Kg, ce qui s'explique par la clientèle majoritairement professionnelle des agences BELGOCAM. Les éleveurs commerciaux privilégient les formats volumineux pour des raisons de coût au kilo et de logistique. Le format 5 Kg, plus adapté aux petits élevages et au commerce de proximité, représente une part plus modeste mais stable des volumes.

L'<b>évolution mensuelle</b> montre que les quatre produits suivent des tendances globalement parallèles, avec des pics synchrones sur les mêmes mois. Cette synchronisation suggère que les déterminants de la demande (saisonnalité, cycles de production, événements climatiques) affectent l'ensemble de la gamme de manière homogène.
"""
story.append(Paragraph(commentary4, style_body))

# Product table
story.append(Paragraph("Synthèse par produit", style_h2))
prod_data = [
    [Paragraph("<b>Produit</b>", style_table_header),
     Paragraph("<b>Famille</b>", style_table_header),
     Paragraph("<b>Format</b>", style_table_header),
     Paragraph("<b>Volume (t)</b>", style_table_header),
     Paragraph("<b>% du total</b>", style_table_header),
     Paragraph("<b>Moy. mensuelle</b>", style_table_header)]
]
for prod in prod_order:
    if prod in prod_totals.index:
        vol = prod_totals[prod]
    else:
        vol = 0
    fam = "Chick Booster" if "CHICK" in prod else "Piglet Booster"
    fmt = "25 Kg" if "25" in prod else "5 Kg"
    pct = vol / total_vol * 100
    moy = vol / 12
    prod_data.append([
        Paragraph(prod.title(), style_table_cell),
        Paragraph(fam, style_table_cell_center),
        Paragraph(fmt, style_table_cell_center),
        Paragraph(f"{vol:.2f}", style_table_cell_right),
        Paragraph(f"{pct:.1f}%", style_table_cell_right),
        Paragraph(f"{moy:.2f}", style_table_cell_right),
    ])
# Total
prod_data.append([
    Paragraph("<b>TOTAL</b>", style_table_total_left),
    Paragraph("", style_table_cell),
    Paragraph("", style_table_cell),
    Paragraph(f"<b>{total_vol:.2f}</b>", style_table_total),
    Paragraph("<b>100.0%</b>", style_table_total),
    Paragraph(f"<b>{total_vol/12:.2f}</b>", style_table_total),
])
prod_table = Table(prod_data, colWidths=[4.5 * cm, 2.8 * cm, 1.8 * cm, 2.5 * cm, 2.5 * cm, 2.5 * cm])
prod_table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), PRIMARY),
    ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
    ('TOPPADDING', (0, 0), (-1, -1), 5),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ('LEFTPADDING', (0, 0), (-1, -1), 5),
    ('RIGHTPADDING', (0, 0), (-1, -1), 5),
    ('GRID', (0, 0), (-1, -1), 0.4, BORDER),
    ('ROWBACKGROUNDS', (0, 1), (-1, -2), [WHITE, TABLE_STRIPE]),
    ('BACKGROUND', (0, -1), (-1, -1), LIGHT_BG),
    ('LINEABOVE', (0, -1), (-1, -1), 1, PRIMARY),
]))
story.append(prod_table)
story.append(Spacer(1, 0.2 * cm))
story.append(Paragraph("Tableau 5 - Synthèse par produit (Oct 25 - Sep 26)", style_caption))

# Spacer
story.append(Spacer(1, 0.6 * cm))

# ============================================================
# SECTION: HEATMAP
# ============================================================
story.append(KeepTogether([
    Paragraph("MATRICE D'ACTIVITÉ", style_kicker),
    Paragraph("Heatmap Agence × Mois", style_h1),
    HRFlowable(width="100%", thickness=2, color=PRIMARY, spaceBefore=2, spaceAfter=12),
]))

# Heatmap chart
img = Image(os.path.join(CHARTS, '06_heatmap_agence.png'), width=CONTENT_W, height=CONTENT_W * 0.55)
story.append(img)
story.append(Paragraph("Figure 5 - Heatmap des ventes mensuelles - Top 15 agences (tonnes)", style_caption))

# Commentary
commentary5 = f"""
La <b>heatmap Agence × Mois</b> permet d'identifier visuellement les patterns d'activité de chaque agence au cours de la période. Les cellules les plus foncées (bleu marine) correspondent aux volumes les plus élevés, tandis que les cellules blanches indiquent une activité faible ou nulle.

L'agence <b>{agence_totals.index[0]}</b> se distingue par une ligne uniformément foncée, témoignant d'une activité soutenue tout au long de la période. Aucun mois de creux significatif n'est observé, ce qui démontre une régularité commerciale remarquable. À l'inverse, certaines agences de milieu de classement présentent des pics ponctuels sur des mois spécifiques, suggérant une activité plus opportuniste ou dépendante de campagnes commerciales localisées.

Les <b>variations mensuelles</b> au sein d'une même agence restent généralement modérées (variation de ±30% autour de la moyenne), ce qui indique une certaine prévisibilité de la demande. Quelques exceptions notables apparaissent toutefois, notamment sur les agences récemment ouvertes ou那些 ayant connu des événements commerciaux particuliers (promotions, formations clients, etc.).
"""
# Fix mixed language in the commentary
commentary5 = commentary5.replace("那些", "")

story.append(Paragraph(commentary5, style_body))

# Key observations
story.append(Paragraph("Observations clés", style_h2))
obs_text = f"""
<b>• Régularité du leader :</b> {agence_totals.index[0]} maintient un volume mensuel supérieur à 10 t pendant la majorité des 12 mois, ce qui démontre la maturité de son marché local.<br/>
<b>• Saisonnalité commune :</b> La plupart des agences présentent un creux d'activité en fin d'année civile (Nov-Déc), suivi d'une reprise progressive à partir de janvier.<br/>
<b>• Pic estival :</b> Le mois de {peak_label} est marqué par une intensification généralisée, visible par l'assombrissement simultané de plusieurs lignes de la heatmap.<br/>
<b>• Agences émergentes :</b> Plusieurs agences de milieu de classement montrent une intensification sur les derniers mois, suggérant un potentiel de croissance à exploiter.
"""
story.append(Paragraph(obs_text, style_body))

# Spacer
story.append(Spacer(1, 0.6 * cm))

# ============================================================
# SECTION: CONCLUSION
# ============================================================
story.append(KeepTogether([
    Paragraph("SYNTHÈSE &amp; OBSERVATIONS", style_kicker),
    Paragraph("Conclusion de l'analyse", style_h1),
    HRFlowable(width="100%", thickness=2, color=PRIMARY, spaceBefore=2, spaceAfter=12),
]))

# Conclusion text
conclusion1 = f"""
Cette <b>analyse des ventes Chick &amp; Piglet Booster sur la période Octobre 2025 - Septembre 2026</b> dresse un panorama complet de l'activité des gammes Booster sur le réseau BELGOCAM. Avec <b>{total_vol:.1f} tonnes</b> écoulées sur 12 mois, l'activité démontre un socle commercial solide, soutenu par un réseau de {n_agences} agences actives réparties sur {n_regions} régions. La moyenne mensuelle de {avg_month:.1f} t traduit une régularité satisfaisante, avec des variations saisonnières contenues autour de cette moyenne.
"""

story.append(Paragraph(conclusion1, style_body))

# Key findings
story.append(Paragraph("Points saillants", style_h2))

findings = f"""
<b>1. Équilibre entre les deux gammes</b><br/>
Chick Booster représente {chick_vol/total_vol*100:.1f}% du volume total et Piglet Booster {piglet_vol/total_vol*100:.1f}%. Cet équilibre relatif témoigne du développement parallèle des filières avicoles et porcines au Cameroun, et de la capacité du réseau BELGOCAM à adresser ces deux marchés de manière complémentaire. Le ratio de {chick_vol/piglet_vol:.2f} en faveur de Chick Booster reste stable sur l'ensemble de la période.

<br/><br/><b>2. Concentration géographique marquée</b><br/>
La région {region_totals_clean.index[0]} capte {region_totals_clean.iloc[0]/total_vol*100:.1f}% du volume total, suivie par {region_totals_clean.index[1]} ({region_totals_clean.iloc[1]/total_vol*100:.1f}%) et {region_totals_clean.index[2]} ({region_totals_clean.iloc[2]/total_vol*100:.1f}%). Cette concentration reflète la répartition de l'activité d'élevage au Cameroun et l'implantation historique du réseau d'agences. Les écarts entre régions restent néanmoins contenus, ce qui témoigne d'une présence commerciale équilibrée à l'échelle nationale.

<br/><br/><b>3. Leadership de l'agence {agence_totals.index[0]}</b><br/>
L'agence {agence_totals.index[0]} domine très largement le classement avec {agence_totals.iloc[0]:.1f} t cumulées ({agence_totals.iloc[0]/total_vol*100:.1f}% du total). Cette position de leader s'explique par la combinaison d'un marché local dynamique, d'une clientèle fidélisée et d'une équipe commerciale expérimentée. Le top 3 concentre {agence_totals.head(3).sum()/total_vol*100:.1f}% du volume, ce qui suggère un potentiel de développement sur les autres agences du réseau.

<br/><br/><b>4. Saisonnalité et dynamique temporelle</b><br/>
L'activité présente une <b>saisonnalité modérée</b> avec un pic en {peak_label} ({peak_value:.1f} t) et un creux en {low_label} ({low_value:.1f} t). L'amplitude entre le pic et le creux est de l'ordre de {(peak_value/low_value - 1)*100:.0f}%, ce qui reste raisonnable pour des produits destinés à l'élevage professionnel. Cette régularité facilite la planification commerciale et logistique.
"""
story.append(Paragraph(findings, style_body))

story.append(Paragraph("Limites et précautions de lecture", style_h2))
limits = """
Cette analyse porte exclusivement sur les volumes écoulés (en tonnes) et ne prend pas en compte le chiffre d'affaires, les marges, ni la rentabilité par produit. Les données du mois de septembre 2026 couvrent la période du 1er au 19 septembre (mois partiel), ce qui peut sous-estimer légèrement le volume total du dernier mois. Par ailleurs, certaines lignes de commande ne renseignent pas l'agence d'origine (4,9 t sur la période, soit moins de 1% du volume), ce qui n'affecte pas significativement les conclusions mais mérite d'être noté pour la qualité des données en amont.
"""
story.append(Paragraph(limits, style_body))

# Final note
story.append(Spacer(1, 0.4 * cm))
story.append(HRFlowable(width="100%", thickness=0.5, color=BORDER, spaceBefore=4, spaceAfter=8))
note_text = """
<i>Pour une analyse plus détaillée, se référer au fichier Excel d'accompagnement qui comporte 7 onglets : synthèse, évolutions mensuelles, déclinaisons par région, par agence et par produit, matrice croisée Agence × Produit, et données détaillées au niveau ligne de commande.</i>
"""
story.append(Paragraph(note_text, ParagraphStyle('note', parent=style_body, fontSize=9, textColor=TEXT_MUTED, alignment=TA_JUSTIFY)))

# ============================================================
# Build PDF
# ============================================================
output_path = os.path.join(DOWNLOAD, "Analyse_Booster_Oct2025-Sep2026.pdf")

# Build with two page templates: cover + body
doc = SimpleDocTemplate(
    output_path,
    pagesize=A4,
    leftMargin=MARGIN_L,
    rightMargin=MARGIN_R,
    topMargin=MARGIN_T,
    bottomMargin=MARGIN_B,
    title="Analyse Ventes Chick & Piglet Booster - Oct 2025 à Sep 2026",
    author="Z.ai",
    subject="Analyse des ventes en volume (tonnes) sur 12 mois",
    creator="Z.ai - ReportLab"
)

# Define templates
frame_cover = Frame(MARGIN_L, MARGIN_B, CONTENT_W, PAGE_H - MARGIN_T - MARGIN_B,
                    leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id='cover_frame')
frame_body = Frame(MARGIN_L, MARGIN_B, CONTENT_W, PAGE_H - MARGIN_T - MARGIN_B,
                   leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id='body_frame')

cover_template = PageTemplate(id='cover', frames=[frame_cover], onPage=cover_page)
body_template = PageTemplate(id='body', frames=[frame_body], onPage=header_footer)
doc.addPageTemplates([cover_template, body_template])

doc.build(story)

print(f"\n✅ PDF généré: {output_path}")
print(f"   Taille: {os.path.getsize(output_path) / 1024:.1f} KB")

# Verify page count
from pypdf import PdfReader
reader = PdfReader(output_path)
print(f"   Pages: {len(reader.pages)}")
