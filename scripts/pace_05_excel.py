"""
Phase Execute - Excel forecast multi-feuilles Q4 2026.
Feuilles:
  1. Synthèse - Vue globale par scénario
  2. Par Famille - Tonnes et CA par famille × scénario
  3. Par Région - Tonnes et CA par région × scénario
  4. Par Agence - Détail par agence × scénario
  5. Par Produit - Détail par produit × scénario
  6. Par Mois - Évolution mensuelle sept-déc
  7. Détail complet - Toutes les combinaisons produit × agence × mois × scénario
  8. Prix - Prix utilisés par produit (stable + baisse)
  9. Hypothèses - Scénarios et hypothèses
"""
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference, BarChart3D
from openpyxl.chart.label import DataLabelList
import json
import os

# Load data
fcst = pd.read_csv("/home/z/my-project/scripts/forecast_q4_2026.csv", parse_dates=['date'])
prix_stable = pd.read_csv("/home/z/my-project/scripts/prix_forecast_stable.csv")
prix_baisse = pd.read_csv("/home/z/my-project/scripts/prix_forecast_baisse.csv")

print(f"Loaded {len(fcst)} forecast records")

# Styles
HEAD_FILL = PatternFill('solid', fgColor='1F4E78')
HEAD_FONT = Font(bold=True, color='FFFFFF', size=11)
SUBHEAD_FILL = PatternFill('solid', fgColor='D9E1F2')
SUBHEAD_FONT = Font(bold=True, color='1F4E78', size=11)
TOTAL_FILL = PatternFill('solid', fgColor='FFF2CC')
TOTAL_FONT = Font(bold=True, size=11)
SCEN_COLORS = {
    'S1_rupture': 'FCE4D6',
    'S2_reappro_50': 'FFF2CC',
    'S3_reappro_100': 'C6EFCE',
    'S4_baisse_prix': 'BDD7EE',
}
THIN = Side(border_style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def style_header_row(ws, row, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER

def style_total_row(ws, row, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = TOTAL_FILL
        cell.font = TOTAL_FONT
        cell.border = BORDER


# === Create workbook ===
wb = openpyxl.Workbook()
wb.remove(wb.active)

# === Sheet 1: Synthèse ===
print("Creating sheet 1: Synthèse...")
ws = wb.create_sheet("1. Synthèse")
ws['A1'] = 'BELGOCAM SA - Forecast Q4 2026 (Septembre - Décembre)'
ws['A1'].font = Font(bold=True, size=16, color='1F4E78')
ws['A2'] = 'Modèle Prophet × 4 scénarios soja | Données: Jan 2025 - 26 Août 2026 | Généré le 27/08/2026'
ws['A2'].font = Font(italic=True, size=10, color='595959')

# Synthèse par scénario
ws['A4'] = 'SYNTHÈSE PAR SCÉNARIO'
ws['A4'].font = SUBHEAD_FONT

headers = ['Scénario', 'Description', 'Tonnes Q4', 'CA (M FCFA)', 'vs S3 (%)']
for i, h in enumerate(headers, 1):
    ws.cell(row=5, column=i, value=h)
style_header_row(ws, 5, len(headers))

scenarios_info = [
    ('S1_rupture', 'Rupture totale soja après 16/09/2026'),
    ('S2_reappro_50', 'Réappro 50% (40 000 sacs) au 01/10/2026'),
    ('S3_reappro_100', 'Réappro 100% (80 000 sacs) au 15/09/2026'),
    ('S4_baisse_prix', 'Réappro 100% + baisse prix soja -10%'),
]

synth = fcst.groupby('scenario').agg(
    tonnes=('tonnes', 'sum'),
    ca_m_fcfa=('ca_m_fcfa', 'sum'),
).round({'tonnes': 0, 'ca_m_fcfa': 1})

ref_ca = synth.loc['S3_reappro_100', 'ca_m_fcfa'] if 'S3_reappro_100' in synth.index else 1

row = 6
for scn, desc in scenarios_info:
    if scn in synth.index:
        t = synth.loc[scn, 'tonnes']
        ca = synth.loc[scn, 'ca_m_fcfa']
        pct = (ca / ref_ca - 1) * 100 if ref_ca > 0 else 0
        ws.cell(row=row, column=1, value=scn)
        ws.cell(row=row, column=2, value=desc)
        ws.cell(row=row, column=3, value=t)
        ws.cell(row=row, column=4, value=ca)
        ws.cell(row=row, column=5, value=f"{pct:+.1f}%")
        # Color the row by scenario
        for c in range(1, 6):
            ws.cell(row=row, column=c).fill = PatternFill('solid', fgColor=SCEN_COLORS[scn])
            ws.cell(row=row, column=c).border = BORDER
        row += 1

# Total row
ws.cell(row=row, column=1, value='Note')
ws.cell(row=row, column=2, value='S3 est le scénario de référence (réappro 100%)')
ws.cell(row=row, column=2).font = Font(italic=True, color='595959')
row += 2

# Synthèse par famille (4 scénarios en colonnes)
ws.cell(row=row, column=1, value='SYNTHÈSE PAR FAMILLE (TONNES)')
ws.cell(row=row, column=1).font = SUBHEAD_FONT
row += 1
headers = ['Famille', 'S1 Rupture', 'S2 Réappro 50%', 'S3 Réappro 100%', 'S4 Baisse prix']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

synth_fam = fcst.groupby(['scenario', 'family'])['tonnes'].sum().unstack(fill_value=0)
families_order = ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'MAIS', 'ALIMENT_COMPLET']
for fam in families_order:
    if fam in synth_fam.columns:
        ws.cell(row=row, column=1, value=fam)
        for i, scn in enumerate(['S1_rupture', 'S2_reappro_50', 'S3_reappro_100', 'S4_baisse_prix'], 2):
            if scn in synth_fam.index:
                ws.cell(row=row, column=i, value=round(synth_fam.loc[scn, fam], 0))
        for c in range(1, 6):
            ws.cell(row=row, column=c).border = BORDER
        row += 1

# Total row
ws.cell(row=row, column=1, value='TOTAL')
for i, scn in enumerate(['S1_rupture', 'S2_reappro_50', 'S3_reappro_100', 'S4_baisse_prix'], 2):
    ws.cell(row=row, column=i, value=round(synth.loc[scn, 'tonnes'], 0))
style_total_row(ws, row, 5)
row += 2

# Synthèse par famille CA
ws.cell(row=row, column=1, value='SYNTHÈSE PAR FAMILLE (CA M FCFA)')
ws.cell(row=row, column=1).font = SUBHEAD_FONT
row += 1
headers = ['Famille', 'S1 Rupture', 'S2 Réappro 50%', 'S3 Réappro 100%', 'S4 Baisse prix']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

synth_fam_ca = fcst.groupby(['scenario', 'family'])['ca_m_fcfa'].sum().unstack(fill_value=0)
for fam in families_order:
    if fam in synth_fam_ca.columns:
        ws.cell(row=row, column=1, value=fam)
        for i, scn in enumerate(['S1_rupture', 'S2_reappro_50', 'S3_reappro_100', 'S4_baisse_prix'], 2):
            if scn in synth_fam_ca.index:
                ws.cell(row=row, column=i, value=round(synth_fam_ca.loc[scn, fam], 1))
        for c in range(1, 6):
            ws.cell(row=row, column=c).border = BORDER
        row += 1

ws.cell(row=row, column=1, value='TOTAL')
for i, scn in enumerate(['S1_rupture', 'S2_reappro_50', 'S3_reappro_100', 'S4_baisse_prix'], 2):
    ws.cell(row=row, column=i, value=round(synth.loc[scn, 'ca_m_fcfa'], 1))
style_total_row(ws, row, 5)

# Column widths
ws.column_dimensions['A'].width = 25
ws.column_dimensions['B'].width = 45
ws.column_dimensions['C'].width = 15
ws.column_dimensions['D'].width = 18
ws.column_dimensions['E'].width = 12

# === Sheet 2: Par Famille ===
print("Creating sheet 2: Par Famille...")
ws = wb.create_sheet("2. Par Famille")
ws['A1'] = 'FORECAST Q4 2026 - DÉTAIL PAR FAMILLE × MOIS × SCÉNARIO'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
ws['A2'] = 'Volumes en tonnes | CA en M FCFA'
ws['A2'].font = Font(italic=True, size=10, color='595959')

row = 4
headers = ['Scénario', 'Famille', 'Sept (t)', 'Oct (t)', 'Nov (t)', 'Déc (t)', 'Total Q4 (t)', 'Total CA (M FCFA)']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

for scn in ['S1_rupture', 'S2_reappro_50', 'S3_reappro_100', 'S4_baisse_prix']:
    for fam in families_order:
        sub = fcst[(fcst['scenario'] == scn) & (fcst['family'] == fam)]
        if len(sub) == 0: continue
        monthly_t = sub.groupby('month')['tonnes'].sum().to_dict()
        total_t = sub['tonnes'].sum()
        total_ca = sub['ca_m_fcfa'].sum()
        ws.cell(row=row, column=1, value=scn)
        ws.cell(row=row, column=2, value=fam)
        ws.cell(row=row, column=3, value=round(monthly_t.get(9, 0), 0))
        ws.cell(row=row, column=4, value=round(monthly_t.get(10, 0), 0))
        ws.cell(row=row, column=5, value=round(monthly_t.get(11, 0), 0))
        ws.cell(row=row, column=6, value=round(monthly_t.get(12, 0), 0))
        ws.cell(row=row, column=7, value=round(total_t, 0))
        ws.cell(row=row, column=8, value=round(total_ca, 1))
        for c in range(1, 9):
            ws.cell(row=row, column=c).border = BORDER
            ws.cell(row=row, column=c).fill = PatternFill('solid', fgColor=SCEN_COLORS[scn])
        row += 1
    # Subtotal per scenario
    sub_total = fcst[fcst['scenario'] == scn]
    monthly_t_total = sub_total.groupby('month')['tonnes'].sum().to_dict()
    ws.cell(row=row, column=1, value=scn)
    ws.cell(row=row, column=2, value='TOTAL')
    ws.cell(row=row, column=3, value=round(monthly_t_total.get(9, 0), 0))
    ws.cell(row=row, column=4, value=round(monthly_t_total.get(10, 0), 0))
    ws.cell(row=row, column=5, value=round(monthly_t_total.get(11, 0), 0))
    ws.cell(row=row, column=6, value=round(monthly_t_total.get(12, 0), 0))
    ws.cell(row=row, column=7, value=round(sub_total['tonnes'].sum(), 0))
    ws.cell(row=row, column=8, value=round(sub_total['ca_m_fcfa'].sum(), 1))
    style_total_row(ws, row, 8)
    row += 2

ws.column_dimensions['A'].width = 18
ws.column_dimensions['B'].width = 22
for col in 'CDEFGH':
    ws.column_dimensions[col].width = 16

# === Sheet 3: Par Région ===
print("Creating sheet 3: Par Région...")
ws = wb.create_sheet("3. Par Région")
ws['A1'] = 'FORECAST Q4 2026 - DÉTAIL PAR RÉGION × MOIS × SCÉNARIO'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers = ['Scénario', 'Région', 'Sept (t)', 'Oct (t)', 'Nov (t)', 'Déc (t)', 'Total Q4 (t)', 'CA (M FCFA)']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

for scn in ['S1_rupture', 'S2_reappro_50', 'S3_reappro_100', 'S4_baisse_prix']:
    for reg in ['Ouest', 'Centre', 'Littoral']:
        sub = fcst[(fcst['scenario'] == scn) & (fcst['region'] == reg)]
        if len(sub) == 0: continue
        monthly_t = sub.groupby('month')['tonnes'].sum().to_dict()
        total_t = sub['tonnes'].sum()
        total_ca = sub['ca_m_fcfa'].sum()
        ws.cell(row=row, column=1, value=scn)
        ws.cell(row=row, column=2, value=reg)
        ws.cell(row=row, column=3, value=round(monthly_t.get(9, 0), 0))
        ws.cell(row=row, column=4, value=round(monthly_t.get(10, 0), 0))
        ws.cell(row=row, column=5, value=round(monthly_t.get(11, 0), 0))
        ws.cell(row=row, column=6, value=round(monthly_t.get(12, 0), 0))
        ws.cell(row=row, column=7, value=round(total_t, 0))
        ws.cell(row=row, column=8, value=round(total_ca, 1))
        for c in range(1, 9):
            ws.cell(row=row, column=c).border = BORDER
            ws.cell(row=row, column=c).fill = PatternFill('solid', fgColor=SCEN_COLORS[scn])
        row += 1
    # Total
    sub_total = fcst[fcst['scenario'] == scn]
    monthly_t_total = sub_total.groupby('month')['tonnes'].sum().to_dict()
    ws.cell(row=row, column=1, value=scn)
    ws.cell(row=row, column=2, value='TOTAL')
    ws.cell(row=row, column=3, value=round(monthly_t_total.get(9, 0), 0))
    ws.cell(row=row, column=4, value=round(monthly_t_total.get(10, 0), 0))
    ws.cell(row=row, column=5, value=round(monthly_t_total.get(11, 0), 0))
    ws.cell(row=row, column=6, value=round(monthly_t_total.get(12, 0), 0))
    ws.cell(row=row, column=7, value=round(sub_total['tonnes'].sum(), 0))
    ws.cell(row=row, column=8, value=round(sub_total['ca_m_fcfa'].sum(), 1))
    style_total_row(ws, row, 8)
    row += 2

ws.column_dimensions['A'].width = 18
ws.column_dimensions['B'].width = 15
for col in 'CDEFGH':
    ws.column_dimensions[col].width = 15

# === Sheet 4: Par Agence ===
print("Creating sheet 4: Par Agence...")
ws = wb.create_sheet("4. Par Agence")
ws['A1'] = 'FORECAST Q4 2026 - DÉTAIL PAR AGENCE × SCÉNARIO'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers = ['Scénario', 'Agence', 'Région', 'Tonnes Q4', 'CA (M FCFA)', 'Part CA (%)']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

for scn in ['S1_rupture', 'S2_reappro_50', 'S3_reappro_100', 'S4_baisse_prix']:
    sub = fcst[fcst['scenario'] == scn]
    by_agence = sub.groupby(['agence', 'region']).agg(
        tonnes=('tonnes', 'sum'),
        ca_m_fcfa=('ca_m_fcfa', 'sum'),
    ).reset_index().sort_values('ca_m_fcfa', ascending=False)
    total_ca = by_agence['ca_m_fcfa'].sum()
    
    for _, r in by_agence.iterrows():
        ws.cell(row=row, column=1, value=scn)
        ws.cell(row=row, column=2, value=r['agence'])
        ws.cell(row=row, column=3, value=r['region'])
        ws.cell(row=row, column=4, value=round(r['tonnes'], 0))
        ws.cell(row=row, column=5, value=round(r['ca_m_fcfa'], 1))
        ws.cell(row=row, column=6, value=f"{r['ca_m_fcfa']/total_ca*100:.1f}%")
        for c in range(1, 7):
            ws.cell(row=row, column=c).border = BORDER
            ws.cell(row=row, column=c).fill = PatternFill('solid', fgColor=SCEN_COLORS[scn])
        row += 1
    # Total
    ws.cell(row=row, column=1, value=scn)
    ws.cell(row=row, column=2, value='TOTAL')
    ws.cell(row=row, column=3, value='—')
    ws.cell(row=row, column=4, value=round(sub['tonnes'].sum(), 0))
    ws.cell(row=row, column=5, value=round(sub['ca_m_fcfa'].sum(), 1))
    ws.cell(row=row, column=6, value='100.0%')
    style_total_row(ws, row, 6)
    row += 2

ws.column_dimensions['A'].width = 18
ws.column_dimensions['B'].width = 18
ws.column_dimensions['C'].width = 12
for col in 'DEF':
    ws.column_dimensions[col].width = 15

# === Sheet 5: Par Produit ===
print("Creating sheet 5: Par Produit...")
ws = wb.create_sheet("5. Par Produit")
ws['A1'] = 'FORECAST Q4 2026 - DÉTAIL PAR PRODUIT × SCÉNARIO'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers = ['Scénario', 'Réf', 'Description', 'Famille', 'Tonnes Q4', 'Sacs 50kg', 'CA (M FCFA)']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

# Load descriptions
desc_df = fcst[['ref', 'family']].drop_duplicates()
# Get description from prix file
prix_df = pd.read_csv("/home/z/my-project/scripts/prix_forecast_stable.csv")[['ref', 'description']]
desc_df = desc_df.merge(prix_df, on='ref', how='left')

for scn in ['S1_rupture', 'S2_reappro_50', 'S3_reappro_100', 'S4_baisse_prix']:
    sub = fcst[fcst['scenario'] == scn]
    by_prod = sub.groupby(['ref', 'family']).agg(
        tonnes=('tonnes', 'sum'),
        sacs_50=('sacs_50', 'sum'),
        ca_m_fcfa=('ca_m_fcfa', 'sum'),
    ).reset_index().sort_values('ca_m_fcfa', ascending=False)
    by_prod = by_prod.merge(desc_df[['ref', 'description']].drop_duplicates(), on='ref', how='left')
    
    for _, r in by_prod.iterrows():
        ws.cell(row=row, column=1, value=scn)
        ws.cell(row=row, column=2, value=r['ref'])
        ws.cell(row=row, column=3, value=r['description'])
        ws.cell(row=row, column=4, value=r['family'])
        ws.cell(row=row, column=5, value=round(r['tonnes'], 1))
        ws.cell(row=row, column=6, value=round(r['sacs_50'], 0))
        ws.cell(row=row, column=7, value=round(r['ca_m_fcfa'], 1))
        for c in range(1, 8):
            ws.cell(row=row, column=c).border = BORDER
            ws.cell(row=row, column=c).fill = PatternFill('solid', fgColor=SCEN_COLORS[scn])
        row += 1
    # Total
    ws.cell(row=row, column=1, value=scn)
    ws.cell(row=row, column=2, value='TOTAL')
    ws.cell(row=row, column=3, value='—')
    ws.cell(row=row, column=4, value='—')
    ws.cell(row=row, column=5, value=round(sub['tonnes'].sum(), 0))
    ws.cell(row=row, column=6, value=round(sub['sacs_50'].sum(), 0))
    ws.cell(row=row, column=7, value=round(sub['ca_m_fcfa'].sum(), 1))
    style_total_row(ws, row, 7)
    row += 2

ws.column_dimensions['A'].width = 18
ws.column_dimensions['B'].width = 10
ws.column_dimensions['C'].width = 30
ws.column_dimensions['D'].width = 18
for col in 'EFG':
    ws.column_dimensions[col].width = 15

# === Sheet 6: Par Mois ===
print("Creating sheet 6: Par Mois...")
ws = wb.create_sheet("6. Par Mois")
ws['A1'] = 'FORECAST Q4 2026 - ÉVOLUTION MENSUELLE'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers = ['Scénario', 'Famille', 'Sept (t)', 'Oct (t)', 'Nov (t)', 'Déc (t)', 'Total Q4 (t)']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

for scn in ['S1_rupture', 'S2_reappro_50', 'S3_reappro_100', 'S4_baisse_prix']:
    for fam in families_order:
        sub = fcst[(fcst['scenario'] == scn) & (fcst['family'] == fam)]
        if len(sub) == 0: continue
        monthly_t = sub.groupby('month')['tonnes'].sum().to_dict()
        total_t = sub['tonnes'].sum()
        ws.cell(row=row, column=1, value=scn)
        ws.cell(row=row, column=2, value=fam)
        ws.cell(row=row, column=3, value=round(monthly_t.get(9, 0), 0))
        ws.cell(row=row, column=4, value=round(monthly_t.get(10, 0), 0))
        ws.cell(row=row, column=5, value=round(monthly_t.get(11, 0), 0))
        ws.cell(row=row, column=6, value=round(monthly_t.get(12, 0), 0))
        ws.cell(row=row, column=7, value=round(total_t, 0))
        for c in range(1, 8):
            ws.cell(row=row, column=c).border = BORDER
            ws.cell(row=row, column=c).fill = PatternFill('solid', fgColor=SCEN_COLORS[scn])
        row += 1

ws.column_dimensions['A'].width = 18
ws.column_dimensions['B'].width = 22
for col in 'CDEFG':
    ws.column_dimensions[col].width = 15

# === Sheet 7: Détail complet ===
print("Creating sheet 7: Détail complet...")
ws = wb.create_sheet("7. Détail complet")
ws['A1'] = 'FORECAST Q4 2026 - DÉTAIL COMPLET (Produit × Agence × Mois × Scénario)'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers = ['Scénario', 'Réf', 'Description', 'Famille', 'Agence', 'Région', 'Mois', 'Année', 'Tonnes', 'Sacs 50kg', 'Prix TTC/sac', 'CA (M FCFA)']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

# Merge description
desc_map = prix_df.set_index('ref')['description'].to_dict()

for scn in ['S1_rupture', 'S2_reappro_50', 'S3_reappro_100', 'S4_baisse_prix']:
    sub = fcst[fcst['scenario'] == scn].sort_values(['family', 'ref', 'agence', 'month'])
    for _, r in sub.iterrows():
        ws.cell(row=row, column=1, value=scn)
        ws.cell(row=row, column=2, value=r['ref'])
        ws.cell(row=row, column=3, value=desc_map.get(r['ref'], ''))
        ws.cell(row=row, column=4, value=r['family'])
        ws.cell(row=row, column=5, value=r['agence'])
        ws.cell(row=row, column=6, value=r['region'])
        ws.cell(row=row, column=7, value=f"{int(r['month']):02d}/2026")
        ws.cell(row=row, column=8, value=int(r['year']))
        ws.cell(row=row, column=9, value=round(r['tonnes'], 2))
        ws.cell(row=row, column=10, value=round(r['sacs_50'], 1))
        ws.cell(row=row, column=11, value=int(r['prix_ttc_sac']))
        ws.cell(row=row, column=12, value=round(r['ca_m_fcfa'], 2))
        for c in range(1, 13):
            ws.cell(row=row, column=c).border = BORDER
        row += 1

ws.column_dimensions['A'].width = 18
ws.column_dimensions['B'].width = 10
ws.column_dimensions['C'].width = 28
ws.column_dimensions['D'].width = 18
ws.column_dimensions['E'].width = 15
ws.column_dimensions['F'].width = 12
ws.column_dimensions['G'].width = 10
ws.column_dimensions['H'].width = 10
for col in 'IJKL':
    ws.column_dimensions[col].width = 13

# Freeze panes
ws.freeze_panes = 'A4'

# === Sheet 8: Prix ===
print("Creating sheet 8: Prix...")
ws = wb.create_sheet("8. Prix utilisés")
ws['A1'] = 'PRIX UTILISÉS POUR LE FORECAST Q4 2026'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
ws['A2'] = 'Option A: prix extrapolés à partir du CA HT/TTC et des quantités (précision ~90%)'
ws['A2'].font = Font(italic=True, size=10, color='595959')

row = 4
headers = ['Réf', 'Description', 'Famille', 'Prix stable (TTC/sac)', 'Prix baisse -10% soja (TTC/sac)', 'Source']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

for _, r in prix_stable.iterrows():
    prix_b = prix_baisse[prix_baisse['ref'] == r['ref']]['prix_ttc'].iloc[0] if len(prix_baisse[prix_baisse['ref'] == r['ref']]) > 0 else r['prix_ttc']
    ws.cell(row=row, column=1, value=r['ref'])
    ws.cell(row=row, column=2, value=r['description'])
    ws.cell(row=row, column=3, value=r['family'])
    ws.cell(row=row, column=4, value=int(r['prix_ttc']))
    ws.cell(row=row, column=5, value=int(prix_b))
    ws.cell(row=row, column=6, value='Août 2026 (dernier prix connu)')
    for c in range(1, 7):
        ws.cell(row=row, column=c).border = BORDER
    row += 1

ws.column_dimensions['A'].width = 10
ws.column_dimensions['B'].width = 30
ws.column_dimensions['C'].width = 18
ws.column_dimensions['D'].width = 22
ws.column_dimensions['E'].width = 30
ws.column_dimensions['F'].width = 28

# === Sheet 9: Hypothèses ===
print("Creating sheet 9: Hypothèses...")
ws = wb.create_sheet("9. Hypothèses")
ws['A1'] = 'HYPOTHÈSES DU FORECAST Q4 2026'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

hypotheses = [
    ('PÉRIODE', 'Septembre - Décembre 2026 (4 mois)'),
    ('DONNÉES HISTORIQUES', 'Janvier 2025 - 26 août 2026 (95 819 enregistrements Livrés)'),
    ('MÉTHODE', 'Prophet (saisonnalité multiplicative, yearly seasonality)'),
    ('NIVEAU DE DÉTAIL', '13 modèles famille × région, désagrégés en produit × agence'),
    ('', ''),
    ('SCÉNARIOS SOJA', ''),
    ('S1 - Rupture totale', 'Ventes soja = 0 à partir du 17/09/2026 (rupture probable 16/09). Sept = 50% du forecast.'),
    ('S2 - Réappro 50%', 'Réappro 40 000 sacs au 01/10/2026. Sept = 50%, Oct-Nov = 70%, Déc = 100%.'),
    ('S3 - Réappro 100%', 'Réappro 80 000 sacs au 15/09/2026. Sept = 70%, Oct-Déc = 100%. [Scénario de référence]'),
    ('S4 - Baisse prix', 'S3 + baisse prix soja -10% (si stock se stabilise). Volume boost +10%.'),
    ('', ''),
    ('PRIX', ''),
    ('Option A', 'Prix extrapolés à partir du CA HT/TTC et des quantités (précision ~90%)'),
    ('Option 3 (validée)', 'Prix stable août → décembre (pas de 3ème hausse)'),
    ('Bonus S4', 'Baisse prix soja -10% (hypothèse: stock stabilisé → ajustement commercial)'),
    ('', ''),
    ('STOCK SOJA (BEKOKO)', ''),
    ('Stock physique au 08/08/2026', '80 430 sacs (4 021 t) - tous formats'),
    ('Allocation SPC (entité soeur)', '9 200 sacs'),
    ('Stock net BELGOCAM', '71 230 sacs (3 561 t)'),
    ('Vente directe moyenne', '18 700 sacs/sem (3 117 sacs/jour)'),
    ('Production concentrés', '4 900 sacs/sem (conso interne soja)'),
    ('Conso totale', '23 600 sacs/sem'),
    ('Jours de stock restants', '21 jours'),
    ('Date rupture probable', '16/09/2026'),
    ('', ''),
    ('SAISONNALITÉ Q4 (2025)', ''),
    ('TOURTEAUX', 'Q4 = 46.3% du volume annuel 2025 (pic en octobre)'),
    ('CONCENTRES', 'Q4 = 35.2% du volume annuel 2025 (croissance déc)'),
    ('INGREDIENTS', 'Q4 = 36.6% du volume annuel 2025'),
    ('ALIMENT_COMPLET', 'Q4 = 42.0% du volume annuel 2025'),
    ('', ''),
    ('LIMITES', ''),
    ('Historique', '20 mois seulement (Prophet recommande 2+ ans pour yearly seasonality)'),
    ('Effet prix', 'Hausse +3 000 FCFA/sac cumulée depuis juillet - effet sur Q4 à confirmer'),
    ('Bundle ratio', 'Préservé via les ratios historiques de désagrégation'),
    ('Événements ponctuels', 'Non modélisables (promotions, ouvertures d\'agences)'),
]

row = 3
for label, value in hypotheses:
    if label and not value:
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=1).font = SUBHEAD_FONT
        ws.cell(row=row, column=1).fill = SUBHEAD_FILL
    elif label:
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row, column=2, value=value)
        ws.cell(row=row, column=1).border = BORDER
        ws.cell(row=row, column=2).border = BORDER
    row += 1

ws.column_dimensions['A'].width = 35
ws.column_dimensions['B'].width = 80

# === Save ===
output_path = "/home/z/my-project/download/forecast_q4_2026_volumes_valeurs.xlsx"
wb.save(output_path)
print(f"\n=== EXCEL SAVED ===")
print(f"Path: {output_path}")
print(f"Size: {os.path.getsize(output_path) / 1024:.0f} KB")
print(f"Sheets: {wb.sheetnames}")
