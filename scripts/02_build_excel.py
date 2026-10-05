"""
Génère le livrable Excel: Analyse Chick Booster + Piglet Booster
Oct 2025 - Sep 2026 | Volume (tonnes) | Par agence, région, produit, mois

7 onglets:
  1. Synthèse
  2. Évolution Mensuelle Globale
  3. Par Région et Mois
  4. Par Agence et Mois
  5. Par Produit et Mois
  6. Cross-Tab Agence x Produit
  7. Données Détaillées
"""
import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference, BarChart3D
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils.dataframe import dataframe_to_rows
import warnings
warnings.filterwarnings('ignore')

WORK = "/home/z/my-project/work"
DOWNLOAD = "/home/z/my-project/download"
os.makedirs(DOWNLOAD, exist_ok=True)

# Load consolidated data
df = pd.read_csv(os.path.join(WORK, 'booster_consolidated.csv'))
df['Date de commande'] = pd.to_datetime(df['Date de commande'])
df['mois_dt'] = pd.to_datetime(df['mois'])

# Order months chronologically
months_order = sorted(df['mois'].unique(), key=lambda x: pd.Period(x, freq='M'))
month_labels_fr = {
    '2025-10': 'Oct 25', '2025-11': 'Nov 25', '2025-12': 'Déc 25',
    '2026-01': 'Jan 26', '2026-02': 'Fév 26', '2026-03': 'Mar 26',
    '2026-04': 'Avr 26', '2026-05': 'Mai 26', '2026-06': 'Juin 26',
    '2026-07': 'Juil 26', '2026-08': 'Août 26', '2026-09': 'Sep 26',
    '2026-10': 'Oct 26*',
}

# ============================================================
# Design tokens (corporate sobre: bleu marine + gris)
# ============================================================
COLOR_PRIMARY = "1F3864"        # Bleu marine
COLOR_SECONDARY = "2E5C8A"      # Bleu plus clair
COLOR_ACCENT = "C9A961"         # Doré sobre
COLOR_LIGHT_BG = "EAEFF7"       # Bleu très clair
COLOR_LIGHTER_BG = "F5F7FB"
COLOR_TEXT = "1F2937"
COLOR_MUTED = "6B7280"
COLOR_WHITE = "FFFFFF"
COLOR_CHICK = "8B5CF6"          # Violet pour Chick
COLOR_PIGLET = "F59E0B"         # Orange pour Piglet

FONT_NAME = "Calibri"

# Borders
thin_grey = Side(style='thin', color='BFBFBF')
thin_dark = Side(style='thin', color='1F3864')
border_all_grey = Border(left=thin_grey, right=thin_grey, top=thin_grey, bottom=thin_grey)
border_header = Border(left=thin_dark, right=thin_dark, top=thin_dark, bottom=thin_dark)

# Fonts
font_title = Font(name=FONT_NAME, size=18, bold=True, color=COLOR_PRIMARY)
font_subtitle = Font(name=FONT_NAME, size=12, bold=False, color=COLOR_MUTED, italic=True)
font_section = Font(name=FONT_NAME, size=13, bold=True, color=COLOR_PRIMARY)
font_header = Font(name=FONT_NAME, size=11, bold=True, color=COLOR_WHITE)
font_total = Font(name=FONT_NAME, size=11, bold=True, color=COLOR_PRIMARY)
font_body = Font(name=FONT_NAME, size=11, color=COLOR_TEXT)
font_small = Font(name=FONT_NAME, size=9, color=COLOR_MUTED, italic=True)
font_kpi_value = Font(name=FONT_NAME, size=20, bold=True, color=COLOR_PRIMARY)
font_kpi_label = Font(name=FONT_NAME, size=10, color=COLOR_MUTED)

# Fills
fill_primary = PatternFill(start_color=COLOR_PRIMARY, end_color=COLOR_PRIMARY, fill_type='solid')
fill_secondary = PatternFill(start_color=COLOR_SECONDARY, end_color=COLOR_SECONDARY, fill_type='solid')
fill_light = PatternFill(start_color=COLOR_LIGHT_BG, end_color=COLOR_LIGHT_BG, fill_type='solid')
fill_lighter = PatternFill(start_color=COLOR_LIGHTER_BG, end_color=COLOR_LIGHTER_BG, fill_type='solid')
fill_total = PatternFill(start_color=COLOR_LIGHT_BG, end_color=COLOR_LIGHT_BG, fill_type='solid')
fill_accent = PatternFill(start_color=COLOR_ACCENT, end_color=COLOR_ACCENT, fill_type='solid')

# Alignments
align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center')


def apply_header_style(ws, row, start_col, end_col):
    """Apply header style to a row range"""
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = font_header
        cell.fill = fill_primary
        cell.alignment = align_center
        cell.border = border_header


def apply_total_row(ws, row, start_col, end_col):
    """Apply total row style"""
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = font_total
        cell.fill = fill_total
        cell.border = border_all_grey


def write_title_block(ws, title, subtitle, end_col):
    """Write a title block at the top of a sheet"""
    ws.cell(row=1, column=1, value=title).font = font_title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    ws.cell(row=1, column=1).alignment = align_left
    ws.row_dimensions[1].height = 32

    ws.cell(row=2, column=1, value=subtitle).font = font_subtitle
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_col)
    ws.cell(row=2, column=1).alignment = align_left
    ws.row_dimensions[2].height = 18


# ============================================================
# Create Workbook
# ============================================================
wb = Workbook()
wb.remove(wb.active)
wb.properties.creator = "William Francis Fohom, Data Analyst"
wb.properties.lastModifiedBy = "William Francis Fohom, Data Analyst"
wb.properties.title = "Analyse Chick & Piglet Booster - Oct 25 à Oct 26"
wb.properties.subject = "Évolution des ventes en volume (tonnes)"

# ============================================================
# SHEET 1: SYNTHÈSE
# ============================================================
print("Building Sheet 1: Synthèse...")
ws = wb.create_sheet("1. Synthèse")
ws.sheet_view.showGridLines = False
end_col = 6

write_title_block(ws, "Analyse des ventes Chick & Piglet Booster",
                  "Période : Octobre 2025 - Octobre 2026 | Métrique : Volume (tonnes) | Source : ERP NJS Group", end_col)

# KPIs
total_vol = df['qte_tonnes'].sum()
n_agences = df['agence'].dropna().nunique()
n_regions = df[df['region'] != 'Non spécifié']['region'].nunique()
n_months = df['mois'].nunique()
chick_vol = df[df['famille'] == 'Chick Booster']['qte_tonnes'].sum()
piglet_vol = df[df['famille'] == 'Piglet Booster']['qte_tonnes'].sum()

# KPI row (row 4)
ws.row_dimensions[4].height = 22
ws.cell(row=4, column=1, value="INDICATEURS CLÉS").font = font_section
ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=end_col)

# KPI cards (row 5-7)
kpis = [
    ("Volume total", f"{total_vol:,.1f} t".replace(',', ' ')),
    ("Agences actives", f"{n_agences}"),
    ("Régions couvertes", f"{n_regions}"),
    ("Mois analysés", f"{n_months}"),
    ("Volume Chick Booster", f"{chick_vol:,.1f} t".replace(',', ' ')),
    ("Volume Piglet Booster", f"{piglet_vol:,.1f} t".replace(',', ' ')),
]

for i, (label, value) in enumerate(kpis):
    col = i + 1
    # Label
    ws.cell(row=5, column=col, value=label).font = font_kpi_label
    ws.cell(row=5, column=col).alignment = align_center
    ws.cell(row=5, column=col).fill = fill_lighter
    # Value
    ws.cell(row=6, column=col, value=value).font = font_kpi_value
    ws.cell(row=6, column=col).alignment = align_center
    ws.cell(row=6, column=col).fill = fill_lighter
    ws.cell(row=6, column=col).border = border_all_grey

ws.row_dimensions[5].height = 22
ws.row_dimensions[6].height = 36

# Top 5 Agences
ws.cell(row=9, column=1, value="TOP 5 AGENCES PAR VOLUME").font = font_section
ws.merge_cells(start_row=9, start_column=1, end_row=9, end_column=3)

top5_agences = df.groupby('agence')['qte_tonnes'].sum().sort_values(ascending=False).head(5)
ws.cell(row=10, column=1, value="Rang").font = font_header
ws.cell(row=10, column=2, value="Agence").font = font_header
ws.cell(row=10, column=3, value="Volume (t)").font = font_header
apply_header_style(ws, 10, 1, 3)

for i, (agence, vol) in enumerate(top5_agences.items()):
    r = 11 + i
    ws.cell(row=r, column=1, value=i + 1).font = font_body
    ws.cell(row=r, column=2, value=agence).font = font_body
    ws.cell(row=r, column=3, value=round(vol, 2)).font = font_body
    ws.cell(row=r, column=3).number_format = '#,##0.00'
    for c in range(1, 4):
        ws.cell(row=r, column=c).border = border_all_grey
        ws.cell(row=r, column=c).alignment = align_center if c == 1 else align_left if c == 2 else align_right

# Top région
ws.cell(row=9, column=5, value="VOLUME PAR RÉGION").font = font_section
ws.merge_cells(start_row=9, start_column=5, end_row=9, end_column=6)

region_vol = df.groupby('region')['qte_tonnes'].sum().sort_values(ascending=False)
ws.cell(row=10, column=5, value="Région").font = font_header
ws.cell(row=10, column=6, value="Volume (t)").font = font_header
apply_header_style(ws, 10, 5, 6)

for i, (region, vol) in enumerate(region_vol.items()):
    r = 11 + i
    ws.cell(row=r, column=5, value=region).font = font_body
    ws.cell(row=r, column=6, value=round(vol, 2)).font = font_body
    ws.cell(row=r, column=6).number_format = '#,##0.00'
    for c in range(5, 7):
        ws.cell(row=r, column=c).border = border_all_grey
        ws.cell(row=r, column=c).alignment = align_left if c == 5 else align_right

# Ventilation par famille
ws.cell(row=18, column=1, value="VENTILATION PAR FAMILLE ET FORMAT").font = font_section
ws.merge_cells(start_row=18, start_column=1, end_row=18, end_column=4)

ws.cell(row=19, column=1, value="Famille").font = font_header
ws.cell(row=19, column=2, value="Format").font = font_header
ws.cell(row=19, column=3, value="Volume (t)").font = font_header
ws.cell(row=19, column=4, value="% du total").font = font_header
apply_header_style(ws, 19, 1, 4)

famille_format = df.groupby(['famille', 'format'])['qte_tonnes'].sum().sort_values(ascending=False)
r = 20
for (fam, fmt), vol in famille_format.items():
    ws.cell(row=r, column=1, value=fam).font = font_body
    ws.cell(row=r, column=2, value=fmt).font = font_body
    ws.cell(row=r, column=3, value=round(vol, 2)).font = font_body
    ws.cell(row=r, column=3).number_format = '#,##0.00'
    pct = vol / total_vol * 100
    ws.cell(row=r, column=4, value=pct / 100).font = font_body
    ws.cell(row=r, column=4).number_format = '0.0%'
    for c in range(1, 5):
        ws.cell(row=r, column=c).border = border_all_grey
        ws.cell(row=r, column=c).alignment = align_left if c <= 2 else align_right
    r += 1

# Total row
ws.cell(row=r, column=1, value="TOTAL").font = font_total
ws.cell(row=r, column=3, value=round(total_vol, 2)).font = font_total
ws.cell(row=r, column=3).number_format = '#,##0.00'
ws.cell(row=r, column=4, value=1.0).font = font_total
ws.cell(row=r, column=4).number_format = '0.0%'
apply_total_row(ws, r, 1, 4)

# Note
ws.cell(row=r + 2, column=1, value="Note : Les volumes sont exprimés en tonnes (1 sac 25Kg = 0,025 t ; 1 sac 5Kg = 0,005 t). Les données Sep 2026 couvrent la période 01-26/09/2026 (mois partiel).").font = font_small
ws.merge_cells(start_row=r + 2, start_column=1, end_row=r + 2, end_column=end_col)
ws.row_dimensions[r + 2].height = 28

# Column widths
ws.column_dimensions['A'].width = 24
ws.column_dimensions['B'].width = 18
ws.column_dimensions['C'].width = 16
ws.column_dimensions['D'].width = 14
ws.column_dimensions['E'].width = 22
ws.column_dimensions['F'].width = 16


# ============================================================
# SHEET 2: ÉVOLUTION MENSUELLE GLOBALE
# ============================================================
print("Building Sheet 2: Évolution Mensuelle Globale...")
ws = wb.create_sheet("2. Évolution Mensuelle")
ws.sheet_view.showGridLines = False
end_col = 5

write_title_block(ws, "Évolution mensuelle globale",
                  "Volume en tonnes par famille de produits - Oct 2025 à Sep 2026", end_col)

# Headers
headers = ['Mois', 'Chick Booster (t)', 'Piglet Booster (t)', 'Total Booster (t)', 'Moyenne mobile 3 mois']
for i, h in enumerate(headers):
    ws.cell(row=4, column=i + 1, value=h)
apply_header_style(ws, 4, 1, end_col)
ws.row_dimensions[4].height = 30

# Data
pivot = df.pivot_table(values='qte_tonnes', index='mois', columns='famille', aggfunc='sum', fill_value=0)
pivot = pivot.reindex(months_order)
pivot['Total'] = pivot.sum(axis=1)
pivot['MA3'] = pivot['Total'].rolling(window=3, min_periods=1).mean()

r = 5
for mois in months_order:
    ws.cell(row=r, column=1, value=month_labels_fr.get(mois, mois)).font = font_body
    ws.cell(row=r, column=2, value=round(pivot.loc[mois, 'Chick Booster'], 2)).font = font_body
    ws.cell(row=r, column=2).number_format = '#,##0.00'
    ws.cell(row=r, column=3, value=round(pivot.loc[mois, 'Piglet Booster'], 2)).font = font_body
    ws.cell(row=r, column=3).number_format = '#,##0.00'
    ws.cell(row=r, column=4, value=round(pivot.loc[mois, 'Total'], 2)).font = font_body
    ws.cell(row=r, column=4).number_format = '#,##0.00'
    ws.cell(row=r, column=5, value=round(pivot.loc[mois, 'MA3'], 2)).font = font_body
    ws.cell(row=r, column=5).number_format = '#,##0.00'
    for c in range(1, end_col + 1):
        ws.cell(row=r, column=c).border = border_all_grey
        ws.cell(row=r, column=c).alignment = align_center if c == 1 else align_right
        if r % 2 == 0:
            ws.cell(row=r, column=c).fill = fill_lighter
    r += 1

# Total row
total_row = r
ws.cell(row=total_row, column=1, value="TOTAL").font = font_total
ws.cell(row=total_row, column=2, value=round(pivot['Chick Booster'].sum(), 2)).font = font_total
ws.cell(row=total_row, column=2).number_format = '#,##0.00'
ws.cell(row=total_row, column=3, value=round(pivot['Piglet Booster'].sum(), 2)).font = font_total
ws.cell(row=total_row, column=3).number_format = '#,##0.00'
ws.cell(row=total_row, column=4, value=round(pivot['Total'].sum(), 2)).font = font_total
ws.cell(row=total_row, column=4).number_format = '#,##0.00'
ws.cell(row=total_row, column=5, value="").font = font_total
apply_total_row(ws, total_row, 1, end_col)

# Average row
avg_row = total_row + 1
ws.cell(row=avg_row, column=1, value="MOY/MOIS").font = font_total
ws.cell(row=avg_row, column=2, value=round(pivot['Chick Booster'].mean(), 2)).font = font_total
ws.cell(row=avg_row, column=2).number_format = '#,##0.00'
ws.cell(row=avg_row, column=3, value=round(pivot['Piglet Booster'].mean(), 2)).font = font_total
ws.cell(row=avg_row, column=3).number_format = '#,##0.00'
ws.cell(row=avg_row, column=4, value=round(pivot['Total'].mean(), 2)).font = font_total
ws.cell(row=avg_row, column=4).number_format = '#,##0.00'
apply_total_row(ws, avg_row, 1, end_col)

# Column widths
ws.column_dimensions['A'].width = 14
for col in 'BCDE':
    ws.column_dimensions[col].width = 22

# Chart: Bar chart (stacked) - Chick + Piglet
chart = BarChart()
chart.type = "col"
chart.style = 11
chart.grouping = "stacked"
chart.overlap = 100
chart.title = "Évolution mensuelle des ventes Booster (tonnes)"
chart.y_axis.title = 'Volume (tonnes)'
chart.x_axis.title = 'Mois'
chart.height = 11
chart.width = 22

data = Reference(ws, min_col=2, min_row=4, max_col=3, max_row=4 + len(months_order))
cats = Reference(ws, min_col=1, min_row=5, max_row=4 + len(months_order))
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

# Style
chart.series[0].graphicalProperties.solidFill = COLOR_CHICK
chart.series[1].graphicalProperties.solidFill = COLOR_PIGLET

ws.add_chart(chart, f"G4")

# Second chart: Total evolution line + MA3
chart2 = LineChart()
chart2.title = "Évolution du volume total (tonnes) avec moyenne mobile 3 mois"
chart2.style = 12
chart2.y_axis.title = 'Volume (tonnes)'
chart2.x_axis.title = 'Mois'
chart2.height = 11
chart2.width = 22

data2 = Reference(ws, min_col=4, min_row=4, max_col=5, max_row=4 + len(months_order))
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats)

chart2.series[0].graphicalProperties.line.solidFill = COLOR_PRIMARY
chart2.series[0].graphicalProperties.line.width = 28000
chart2.series[1].graphicalProperties.line.solidFill = COLOR_ACCENT
chart2.series[1].graphicalProperties.line.dashStyle = "dash"
chart2.series[1].graphicalProperties.line.width = 22000

ws.add_chart(chart2, f"G25")


# ============================================================
# SHEET 3: PAR RÉGION ET MOIS
# ============================================================
print("Building Sheet 3: Par Région et Mois...")
ws = wb.create_sheet("3. Par Région")
ws.sheet_view.showGridLines = False
end_col = len(months_order) + 2  # Region + 12 months + Total

write_title_block(ws, "Évolution mensuelle par région",
                  "Volume en tonnes par région - Oct 2025 à Sep 2026", end_col)

# Headers
ws.cell(row=4, column=1, value="Région")
for i, mois in enumerate(months_order):
    ws.cell(row=4, column=2 + i, value=month_labels_fr.get(mois, mois))
ws.cell(row=4, column=2 + len(months_order), value="Total (t)")
apply_header_style(ws, 4, 1, end_col)
ws.row_dimensions[4].height = 30

# Data: pivot region x month
pivot_reg = df.pivot_table(values='qte_tonnes', index='region', columns='mois', aggfunc='sum', fill_value=0)
# Reorder columns
pivot_reg = pivot_reg.reindex(columns=months_order, fill_value=0)
# Add total column
pivot_reg['Total'] = pivot_reg.sum(axis=1)
# Sort by total descending
pivot_reg = pivot_reg.sort_values('Total', ascending=False)

r = 5
for region in pivot_reg.index:
    ws.cell(row=r, column=1, value=region).font = font_body
    for i, mois in enumerate(months_order):
        val = pivot_reg.loc[region, mois]
        ws.cell(row=r, column=2 + i, value=round(val, 2)).font = font_body
        ws.cell(row=r, column=2 + i).number_format = '#,##0.00'
    ws.cell(row=r, column=2 + len(months_order), value=round(pivot_reg.loc[region, 'Total'], 2)).font = font_total
    ws.cell(row=r, column=2 + len(months_order)).number_format = '#,##0.00'
    for c in range(1, end_col + 1):
        ws.cell(row=r, column=c).border = border_all_grey
        ws.cell(row=r, column=c).alignment = align_left if c == 1 else align_right
    r += 1

# Total row
total_row = r
ws.cell(row=total_row, column=1, value="TOTAL").font = font_total
for i, mois in enumerate(months_order):
    total = pivot_reg[mois].sum()
    ws.cell(row=total_row, column=2 + i, value=round(total, 2)).font = font_total
    ws.cell(row=total_row, column=2 + i).number_format = '#,##0.00'
ws.cell(row=total_row, column=2 + len(months_order), value=round(pivot_reg['Total'].sum(), 2)).font = font_total
ws.cell(row=total_row, column=2 + len(months_order)).number_format = '#,##0.00'
apply_total_row(ws, total_row, 1, end_col)

# Apply conditional formatting (heatmap) on month cells
data_range = f"B5:{get_column_letter(1 + len(months_order))}{total_row - 1}"
color_scale = ColorScaleRule(
    start_type='min', start_color='FFFFFF',
    mid_type='percentile', mid_value=50, mid_color='9DC3E6',
    end_type='max', end_color='1F3864'
)
ws.conditional_formatting.add(data_range, color_scale)

# Column widths
ws.column_dimensions['A'].width = 18
for i in range(len(months_order) + 1):
    ws.column_dimensions[get_column_letter(2 + i)].width = 12

# Chart: line chart of regional evolution
chart = LineChart()
chart.title = "Évolution mensuelle par région (tonnes)"
chart.style = 12
chart.y_axis.title = 'Volume (tonnes)'
chart.x_axis.title = 'Mois'
chart.height = 11
chart.width = 24

n_regions = len(pivot_reg.index)
data = Reference(ws, min_col=1, min_row=4, max_col=1 + len(months_order), max_row=4 + n_regions)
chart.add_data(data, titles_from_data=True, from_rows=True)
cats = Reference(ws, min_col=2, min_row=4, max_col=1 + len(months_order), max_row=4)
chart.set_categories(cats)

# Style each series with different colors
region_colors = [COLOR_PRIMARY, COLOR_PIGLET, COLOR_CHICK, COLOR_MUTED]
for i, s in enumerate(chart.series):
    s.graphicalProperties.line.solidFill = region_colors[i % len(region_colors)]
    s.graphicalProperties.line.width = 22000

ws.add_chart(chart, f"A{total_row + 3}")


# ============================================================
# SHEET 4: PAR AGENCE ET MOIS
# ============================================================
print("Building Sheet 4: Par Agence et Mois...")
ws = wb.create_sheet("4. Par Agence")
ws.sheet_view.showGridLines = False
end_col = len(months_order) + 3  # Agence + Region + 12 months + Total

write_title_block(ws, "Évolution mensuelle par agence",
                  "Volume en tonnes par agence - Oct 2025 à Sep 2026", end_col)

# Headers
ws.cell(row=4, column=1, value="Agence")
ws.cell(row=4, column=2, value="Région")
for i, mois in enumerate(months_order):
    ws.cell(row=4, column=3 + i, value=month_labels_fr.get(mois, mois))
ws.cell(row=4, column=3 + len(months_order), value="Total (t)")
apply_header_style(ws, 4, 1, end_col)
ws.row_dimensions[4].height = 30

# Data: pivot agence x month
pivot_ag = df.pivot_table(values='qte_tonnes', index=['agence', 'region'], columns='mois', aggfunc='sum', fill_value=0)
pivot_ag = pivot_ag.reindex(columns=months_order, fill_value=0)
pivot_ag['Total'] = pivot_ag.sum(axis=1)
pivot_ag = pivot_ag.sort_values('Total', ascending=False)
pivot_ag = pivot_ag.reset_index()

r = 5
for _, row in pivot_ag.iterrows():
    ws.cell(row=r, column=1, value=row['agence']).font = font_body
    ws.cell(row=r, column=2, value=row['region']).font = font_body
    for i, mois in enumerate(months_order):
        val = row[mois]
        ws.cell(row=r, column=3 + i, value=round(val, 2)).font = font_body
        ws.cell(row=r, column=3 + i).number_format = '#,##0.00'
    ws.cell(row=r, column=3 + len(months_order), value=round(row['Total'], 2)).font = font_total
    ws.cell(row=r, column=3 + len(months_order)).number_format = '#,##0.00'
    for c in range(1, end_col + 1):
        ws.cell(row=r, column=c).border = border_all_grey
        ws.cell(row=r, column=c).alignment = align_left if c <= 2 else align_right
    r += 1

# Total row
total_row = r
ws.cell(row=total_row, column=1, value="TOTAL").font = font_total
ws.cell(row=total_row, column=2, value="").font = font_total
for i, mois in enumerate(months_order):
    total = pivot_ag[mois].sum()
    ws.cell(row=total_row, column=3 + i, value=round(total, 2)).font = font_total
    ws.cell(row=total_row, column=3 + i).number_format = '#,##0.00'
ws.cell(row=total_row, column=3 + len(months_order), value=round(pivot_ag['Total'].sum(), 2)).font = font_total
ws.cell(row=total_row, column=3 + len(months_order)).number_format = '#,##0.00'
apply_total_row(ws, total_row, 1, end_col)

# Apply heatmap
data_range = f"C5:{get_column_letter(2 + len(months_order))}{total_row - 1}"
color_scale = ColorScaleRule(
    start_type='min', start_color='FFFFFF',
    mid_type='percentile', mid_value=50, mid_color='9DC3E6',
    end_type='max', end_color='1F3864'
)
ws.conditional_formatting.add(data_range, color_scale)

# Column widths
ws.column_dimensions['A'].width = 22
ws.column_dimensions['B'].width = 16
for i in range(len(months_order) + 1):
    ws.column_dimensions[get_column_letter(3 + i)].width = 11

# Freeze panes
ws.freeze_panes = "C5"


# ============================================================
# SHEET 5: PAR PRODUIT ET MOIS
# ============================================================
print("Building Sheet 5: Par Produit et Mois...")
ws = wb.create_sheet("5. Par Produit")
ws.sheet_view.showGridLines = False
end_col = len(months_order) + 2

write_title_block(ws, "Évolution mensuelle par produit",
                  "Volume en tonnes par produit - Oct 2025 à Sep 2026", end_col)

# Headers
ws.cell(row=4, column=1, value="Produit")
for i, mois in enumerate(months_order):
    ws.cell(row=4, column=2 + i, value=month_labels_fr.get(mois, mois))
ws.cell(row=4, column=2 + len(months_order), value="Total (t)")
apply_header_style(ws, 4, 1, end_col)
ws.row_dimensions[4].height = 30

# Data: pivot product x month
pivot_pr = df.pivot_table(values='qte_tonnes', index='Description du produit', columns='mois', aggfunc='sum', fill_value=0)
pivot_pr = pivot_pr.reindex(columns=months_order, fill_value=0)
pivot_pr['Total'] = pivot_pr.sum(axis=1)
# Sort: Chick 25Kg, Chick 5Kg, Piglet 25Kg, Piglet 5Kg
product_order = ['CHICK BOOSTER 25 Kg', 'CHICK BOOSTER 5Kg', 'PIGLET BOOSTER 25Kg', 'PIGLET BOOSTER 5Kg']
pivot_pr = pivot_pr.reindex(product_order)

r = 5
# Chick products section
ws.cell(row=r, column=1, value="CHICK BOOSTER").font = font_section
ws.cell(row=r, column=1).fill = fill_light
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=end_col)
r += 1
for prod in ['CHICK BOOSTER 25 Kg', 'CHICK BOOSTER 5Kg']:
    ws.cell(row=r, column=1, value=prod).font = font_body
    for i, mois in enumerate(months_order):
        val = pivot_pr.loc[prod, mois]
        ws.cell(row=r, column=2 + i, value=round(val, 2)).font = font_body
        ws.cell(row=r, column=2 + i).number_format = '#,##0.00'
    ws.cell(row=r, column=2 + len(months_order), value=round(pivot_pr.loc[prod, 'Total'], 2)).font = font_body
    ws.cell(row=r, column=2 + len(months_order)).number_format = '#,##0.00'
    for c in range(1, end_col + 1):
        ws.cell(row=r, column=c).border = border_all_grey
        ws.cell(row=r, column=c).alignment = align_left if c == 1 else align_right
    r += 1
# Subtotal Chick
ws.cell(row=r, column=1, value="  Sous-total Chick").font = font_total
chick_data = pivot_pr.loc[['CHICK BOOSTER 25 Kg', 'CHICK BOOSTER 5Kg']]
for i, mois in enumerate(months_order):
    ws.cell(row=r, column=2 + i, value=round(chick_data[mois].sum(), 2)).font = font_total
    ws.cell(row=r, column=2 + i).number_format = '#,##0.00'
ws.cell(row=r, column=2 + len(months_order), value=round(chick_data['Total'].sum(), 2)).font = font_total
ws.cell(row=r, column=2 + len(months_order)).number_format = '#,##0.00'
apply_total_row(ws, r, 1, end_col)
r += 1

# Piglet products section
ws.cell(row=r, column=1, value="PIGLET BOOSTER").font = font_section
ws.cell(row=r, column=1).fill = fill_light
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=end_col)
r += 1
for prod in ['PIGLET BOOSTER 25Kg', 'PIGLET BOOSTER 5Kg']:
    ws.cell(row=r, column=1, value=prod).font = font_body
    for i, mois in enumerate(months_order):
        val = pivot_pr.loc[prod, mois]
        ws.cell(row=r, column=2 + i, value=round(val, 2)).font = font_body
        ws.cell(row=r, column=2 + i).number_format = '#,##0.00'
    ws.cell(row=r, column=2 + len(months_order), value=round(pivot_pr.loc[prod, 'Total'], 2)).font = font_body
    ws.cell(row=r, column=2 + len(months_order)).number_format = '#,##0.00'
    for c in range(1, end_col + 1):
        ws.cell(row=r, column=c).border = border_all_grey
        ws.cell(row=r, column=c).alignment = align_left if c == 1 else align_right
    r += 1
# Subtotal Piglet
ws.cell(row=r, column=1, value="  Sous-total Piglet").font = font_total
piglet_data = pivot_pr.loc[['PIGLET BOOSTER 25Kg', 'PIGLET BOOSTER 5Kg']]
for i, mois in enumerate(months_order):
    ws.cell(row=r, column=2 + i, value=round(piglet_data[mois].sum(), 2)).font = font_total
    ws.cell(row=r, column=2 + i).number_format = '#,##0.00'
ws.cell(row=r, column=2 + len(months_order), value=round(piglet_data['Total'].sum(), 2)).font = font_total
ws.cell(row=r, column=2 + len(months_order)).number_format = '#,##0.00'
apply_total_row(ws, r, 1, end_col)
r += 1

# Grand total
ws.cell(row=r, column=1, value="TOTAL GÉNÉRAL").font = font_total
for i, mois in enumerate(months_order):
    ws.cell(row=r, column=2 + i, value=round(pivot_pr[mois].sum(), 2)).font = font_total
    ws.cell(row=r, column=2 + i).number_format = '#,##0.00'
ws.cell(row=r, column=2 + len(months_order), value=round(pivot_pr['Total'].sum(), 2)).font = font_total
ws.cell(row=r, column=2 + len(months_order)).number_format = '#,##0.00'
apply_total_row(ws, r, 1, end_col)
total_row = r

# Column widths
ws.column_dimensions['A'].width = 28
for i in range(len(months_order) + 1):
    ws.column_dimensions[get_column_letter(2 + i)].width = 12

# Chart: bar chart by product
chart = BarChart()
chart.type = "col"
chart.style = 11
chart.grouping = "clustered"
chart.title = "Évolution mensuelle par produit (tonnes)"
chart.y_axis.title = 'Volume (tonnes)'
chart.x_axis.title = 'Mois'
chart.height = 12
chart.width = 26

# Add data per product (skip subtotals)
data = Reference(ws, min_col=1, min_row=6, max_col=1 + len(months_order), max_row=7)  # Chick 25Kg + 5Kg
chart.add_data(data, titles_from_data=True, from_rows=True)

data2 = Reference(ws, min_col=1, min_row=10, max_col=1 + len(months_order), max_row=11)  # Piglet 25Kg + 5Kg
chart.add_data(data2, titles_from_data=True, from_rows=True)

cats = Reference(ws, min_col=2, min_row=4, max_col=1 + len(months_order), max_row=4)
chart.set_categories(cats)

prod_colors = [COLOR_CHICK, "C4B5FD", COLOR_PIGLET, "FCD34D"]
for i, s in enumerate(chart.series):
    s.graphicalProperties.solidFill = prod_colors[i]

ws.add_chart(chart, f"A{total_row + 3}")


# ============================================================
# SHEET 6: CROSS-TAB AGENCE x PRODUIT
# ============================================================
print("Building Sheet 6: Cross-Tab Agence x Produit...")
ws = wb.create_sheet("6. Agence x Produit")
ws.sheet_view.showGridLines = False
products_order = ['CHICK BOOSTER 25 Kg', 'CHICK BOOSTER 5Kg', 'PIGLET BOOSTER 25Kg', 'PIGLET BOOSTER 5Kg']
end_col = 4 + 3  # Agence + Region + 4 products + Total

write_title_block(ws, "Matrice Agence x Produit",
                  "Volume annuel en tonnes par agence et par produit - Oct 2025 à Sep 2026", end_col)

# Headers
ws.cell(row=4, column=1, value="Agence")
ws.cell(row=4, column=2, value="Région")
for i, prod in enumerate(products_order):
    short_name = prod.replace(' BOOSTER', '').replace(' Kg', 'Kg').replace('Kg', ' Kg').title()
    ws.cell(row=4, column=3 + i, value=short_name)
ws.cell(row=4, column=7, value="Total (t)")
apply_header_style(ws, 4, 1, end_col)
ws.row_dimensions[4].height = 30

# Pivot agence x product
pivot_ap = df.pivot_table(values='qte_tonnes', index=['agence', 'region'], columns='Description du produit', aggfunc='sum', fill_value=0)
for p in products_order:
    if p not in pivot_ap.columns:
        pivot_ap[p] = 0
pivot_ap = pivot_ap[products_order]
pivot_ap['Total'] = pivot_ap.sum(axis=1)
pivot_ap = pivot_ap.sort_values('Total', ascending=False).reset_index()

r = 5
for _, row in pivot_ap.iterrows():
    ws.cell(row=r, column=1, value=row['agence']).font = font_body
    ws.cell(row=r, column=2, value=row['region']).font = font_body
    for i, prod in enumerate(products_order):
        val = row[prod]
        ws.cell(row=r, column=3 + i, value=round(val, 2)).font = font_body
        ws.cell(row=r, column=3 + i).number_format = '#,##0.00'
    ws.cell(row=r, column=7, value=round(row['Total'], 2)).font = font_total
    ws.cell(row=r, column=7).number_format = '#,##0.00'
    for c in range(1, end_col + 1):
        ws.cell(row=r, column=c).border = border_all_grey
        ws.cell(row=r, column=c).alignment = align_left if c <= 2 else align_right
    r += 1

# Total row
total_row = r
ws.cell(row=total_row, column=1, value="TOTAL").font = font_total
ws.cell(row=total_row, column=2, value="").font = font_total
for i, prod in enumerate(products_order):
    total = pivot_ap[prod].sum()
    ws.cell(row=total_row, column=3 + i, value=round(total, 2)).font = font_total
    ws.cell(row=total_row, column=3 + i).number_format = '#,##0.00'
ws.cell(row=total_row, column=7, value=round(pivot_ap['Total'].sum(), 2)).font = font_total
ws.cell(row=total_row, column=7).number_format = '#,##0.00'
apply_total_row(ws, total_row, 1, end_col)

# Heatmap on product columns
data_range = f"C5:F{total_row - 1}"
color_scale = ColorScaleRule(
    start_type='min', start_color='FFFFFF',
    mid_type='percentile', mid_value=50, mid_color='9DC3E6',
    end_type='max', end_color='1F3864'
)
ws.conditional_formatting.add(data_range, color_scale)

# Column widths
ws.column_dimensions['A'].width = 22
ws.column_dimensions['B'].width = 16
for col in 'CDEF':
    ws.column_dimensions[col].width = 18
ws.column_dimensions['G'].width = 14


# ============================================================
# SHEET 7: DONNÉES DÉTAILLÉES
# ============================================================
print("Building Sheet 7: Données Détaillées...")
ws = wb.create_sheet("7. Données Détaillées")
ws.sheet_view.showGridLines = False

# Filter to important columns
detail_cols = ['Date de commande', 'mois', 'agence', 'region', 'Description du produit',
               'famille', 'format', 'qte_sacs', 'qte_tonnes', 'Montant HT']
df_detail = df[detail_cols].copy()
df_detail = df_detail.sort_values(['Date de commande', 'agence'])

# Headers
for i, c in enumerate(detail_cols):
    ws.cell(row=1, column=i + 1, value=c)
apply_header_style(ws, 1, 1, len(detail_cols))
ws.row_dimensions[1].height = 30

# Data
for r_idx, row in enumerate(df_detail.itertuples(index=False), start=2):
    for c_idx, col in enumerate(detail_cols, start=1):
        val = getattr(row, col.replace(' ', '_').replace('.', '_')) if hasattr(row, col.replace(' ', '_').replace('.', '_')) else None
        # Use dict access instead
        val = row[c_idx - 1]
        cell = ws.cell(row=r_idx, column=c_idx, value=val)
        cell.font = font_body
        cell.alignment = align_left
        cell.border = border_all_grey
        if col == 'qte_tonnes':
            cell.number_format = '#,##0.00'
            cell.alignment = align_right
        elif col == 'qte_sacs':
            cell.number_format = '#,##0'
            cell.alignment = align_right
        elif col == 'Montant HT':
            cell.number_format = '#,##0'
            cell.alignment = align_right
        elif col == 'Date de commande':
            cell.number_format = 'DD/MM/YYYY'
            cell.alignment = align_center
        if r_idx % 2 == 0:
            cell.fill = fill_lighter

# Auto-filter
ws.auto_filter.ref = f"A1:{get_column_letter(len(detail_cols))}{len(df_detail) + 1}"

# Freeze panes
ws.freeze_panes = "A2"

# Column widths
col_widths = {'A': 14, 'B': 10, 'C': 22, 'D': 14, 'E': 26, 'F': 16, 'G': 10, 'H': 12, 'I': 14, 'J': 14}
for col, w in col_widths.items():
    ws.column_dimensions[col].width = w


# ============================================================
# Save workbook
# ============================================================
output_path = os.path.join(DOWNLOAD, "Analyse_Booster_Oct2025-Oct2026.xlsx")
wb.save(output_path)
print(f"\n✅ Excel file saved: {output_path}")
print(f"   Size: {os.path.getsize(output_path) / 1024:.1f} KB")
print(f"   Sheets: {wb.sheetnames}")
