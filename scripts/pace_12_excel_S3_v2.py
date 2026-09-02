"""
Génération Excel forecast Q4 2026 - Scénario S3 - VERSION 2 (avec COMPLEMENT_ALIMENTAIRE, données 2023-2026)
- Sans MAIS
- Avec MATERIEL_ELEVAGE en CA (tonnes=0)
- Avec PREMIX en CA (tonnes=0)
- NOUVEAU: Avec COMPLEMENT_ALIMENTAIRE (V300 1L only, 1L=1kg)
- 7 familles au total
"""
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os

# Load forecast
fcst = pd.read_csv("/home/z/my-project/scripts/forecast_q4_2026_S3.csv", parse_dates=['date'])
print(f"Loaded {len(fcst)} forecast records")

# Load descriptions from new 2023-2026 dataset
desc_df = pd.read_csv("/home/z/my-project/scripts/dataset_2023_2026.csv", low_memory=False)
desc_map = desc_df[['ref', 'description']].drop_duplicates().set_index('ref')['description'].to_dict()

# Styles
HEAD_FILL = PatternFill('solid', fgColor='1F4E78')
HEAD_FONT = Font(bold=True, color='FFFFFF', size=11)
SUBHEAD_FILL = PatternFill('solid', fgColor='D9E1F2')
SUBHEAD_FONT = Font(bold=True, color='1F4E78', size=11)
TOTAL_FILL = PatternFill('solid', fgColor='FFF2CC')
TOTAL_FONT = Font(bold=True, size=11)
S3_COLOR = 'C6EFCE'  # Green for S3
NEW_FAMILY_COLOR = 'FFE699'  # Yellow for COMPLEMENT_ALIMENTAIRE
THIN = Side(border_style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# 7 familles (ajout COMPLEMENT_ALIMENTAIRE)
FAMILIES = ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'ALIMENT_COMPLET', 'MATERIEL_ELEVAGE', 'PREMIX', 'COMPLEMENT_ALIMENTAIRE']

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

def style_data_row(ws, row, n_cols, color=None):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.border = BORDER
        if color:
            cell.fill = PatternFill('solid', fgColor=color)

# === Create workbook ===
wb = openpyxl.Workbook()
wb.remove(wb.active)

# === Sheet 1: Synthèse ===
print("Creating sheet 1: Synthèse...")
ws = wb.create_sheet("1. Synthèse")
ws['A1'] = 'BELGOCAM SA - Forecast Q4 2026 (Septembre - Décembre) - Version 2'
ws['A1'].font = Font(bold=True, size=16, color='1F4E78')
ws['A2'] = 'S3 | Données 2023-2026 (44 mois) | 7 familles | Sans Maïs | MATERIEL_ELEVAGE tonnes=0 | COMPLEMENT_ALIMENTAIRE V300 1L only'
ws['A2'].font = Font(italic=True, size=10, color='595959')

ws['A4'] = 'SYNTHÈSE GLOBALE Q4 2026'
ws['A4'].font = SUBHEAD_FONT

headers = ['Indicateur', 'Valeur', 'Unité']
for i, h in enumerate(headers, 1):
    ws.cell(row=5, column=i, value=h)
style_header_row(ws, 5, 3)

total_t = fcst['tonnes'].sum()
total_ca = fcst['ca_m_fcfa'].sum()
total_sacs = fcst['sacs_50'].sum()

global_data = [
    ('Volume total Q4', f"{total_t:,.0f}", 'tonnes'),
    ('CA total Q4', f"{total_ca:,.1f}", 'M FCFA'),
    ('Sacs éq. 50kg total Q4', f"{total_sacs:,.0f}", 'sacs'),
    ('Période forecast', 'Septembre - Décembre 2026', '4 mois'),
    ('Scénario', 'S3 - Réappro soja 100%', '80 000 sacs au 15/09'),
    ('Nb familles', '7', 'TOURTEAUX, CONCENTRES, INGREDIENTS, ALIMENT_COMPLET, MATERIEL_ELEVAGE, PREMIX, COMPLEMENT_ALIMENTAIRE'),
    ('Nb produits', f"{fcst['ref'].nunique()}", 'références'),
    ('Nb agences', f"{fcst['agence'].nunique()}", 'agences'),
    ('Nb régions', f"{fcst['region'].nunique()}", 'régions'),
    ('Méthode', 'Prophet (5 familles) + Extrapolation (MAT_ELEVAGE, PREMIX)', '—'),
    ('Prix', 'Prix août 2026 + soja 25 000 FCFA/sac', 'Option A - extrapolation CA/qté'),
    ('NOUVEAU: COMPLEMENT_ALIMENTAIRE', 'V300 1L only (V305 200L exclu), 1L=1kg', '9 produits'),
    ('NOUVEAU: Données 2023-2026', '44 mois d\'historique (Jan 2023 - Août 2026)', '176 576 records'),
]

row = 6
for label, val, unit in global_data:
    ws.cell(row=row, column=1, value=label).font = Font(bold=True)
    ws.cell(row=row, column=2, value=val)
    ws.cell(row=row, column=3, value=unit)
    style_data_row(ws, row, 3, color=S3_COLOR)
    row += 1

row += 1
ws.cell(row=row, column=1, value='SYNTHÈSE PAR FAMILLE')
ws.cell(row=row, column=1).font = SUBHEAD_FONT
row += 1

headers = ['Famille', 'Volume Q4 (t)', 'CA Q4 (M FCFA)', 'Part CA (%)', 'Méthode forecast']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

fam_synth = fcst.groupby('family').agg(
    tonnes=('tonnes', 'sum'),
    ca_m_fcfa=('ca_m_fcfa', 'sum'),
).reindex(FAMILLES := ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'ALIMENT_COMPLET', 'MATERIEL_ELEVAGE', 'PREMIX', 'COMPLEMENT_ALIMENTAIRE'])

for fam in FAMILLES:
    if fam in fam_synth.index:
        t = fam_synth.loc[fam, 'tonnes']
        ca = fam_synth.loc[fam, 'ca_m_fcfa']
        pct = ca / total_ca * 100 if total_ca > 0 else 0
        method = 'Prophet' if fam in ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'ALIMENT_COMPLET', 'COMPLEMENT_ALIMENTAIRE'] else 'Extrapolation moyenne Q4'
        ws.cell(row=row, column=1, value=fam)
        ws.cell(row=row, column=2, value=round(t, 0))
        ws.cell(row=row, column=3, value=round(ca, 1))
        ws.cell(row=row, column=4, value=f"{pct:.1f}%")
        ws.cell(row=row, column=5, value=method)
        color = NEW_FAMILY_COLOR if fam == 'COMPLEMENT_ALIMENTAIRE' else S3_COLOR
        style_data_row(ws, row, 5, color=color)
        row += 1

# Total
ws.cell(row=row, column=1, value='TOTAL')
ws.cell(row=row, column=2, value=round(total_t, 0))
ws.cell(row=row, column=3, value=round(total_ca, 1))
ws.cell(row=row, column=4, value='100.0%')
ws.cell(row=row, column=5, value='—')
style_total_row(ws, row, 5)

ws.column_dimensions['A'].width = 30
ws.column_dimensions['B'].width = 18
ws.column_dimensions['C'].width = 20
ws.column_dimensions['D'].width = 12
ws.column_dimensions['E'].width = 35

# === Sheet 2: Par Famille ===
print("Creating sheet 2: Par Famille...")
ws = wb.create_sheet("2. Par Famille")
ws['A1'] = 'FORECAST Q4 2026 (S3) - DÉTAIL PAR FAMILLE × MOIS'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers = ['Famille', 'Sept (t)', 'Oct (t)', 'Nov (t)', 'Déc (t)', 'Total Q4 (t)', 'Total CA (M FCFA)', 'Part CA (%)']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

for fam in FAMILLES:
    sub = fcst[fcst['family'] == fam]
    if len(sub) == 0: continue
    monthly_t = sub.groupby('month')['tonnes'].sum().to_dict()
    monthly_ca = sub.groupby('month')['ca_m_fcfa'].sum().to_dict()
    total_t = sub['tonnes'].sum()
    total_ca_fam = sub['ca_m_fcfa'].sum()
    pct = total_ca_fam / total_ca * 100 if total_ca > 0 else 0
    ws.cell(row=row, column=1, value=fam)
    ws.cell(row=row, column=2, value=round(monthly_t.get(9, 0), 0))
    ws.cell(row=row, column=3, value=round(monthly_t.get(10, 0), 0))
    ws.cell(row=row, column=4, value=round(monthly_t.get(11, 0), 0))
    ws.cell(row=row, column=5, value=round(monthly_t.get(12, 0), 0))
    ws.cell(row=row, column=6, value=round(total_t, 0))
    ws.cell(row=row, column=7, value=round(total_ca_fam, 1))
    ws.cell(row=row, column=8, value=f"{pct:.1f}%")
    style_data_row(ws, row, 8, color=S3_COLOR)
    row += 1

# Total
ws.cell(row=row, column=1, value='TOTAL')
monthly_t_total = fcst.groupby('month')['tonnes'].sum().to_dict()
ws.cell(row=row, column=2, value=round(monthly_t_total.get(9, 0), 0))
ws.cell(row=row, column=3, value=round(monthly_t_total.get(10, 0), 0))
ws.cell(row=row, column=4, value=round(monthly_t_total.get(11, 0), 0))
ws.cell(row=row, column=5, value=round(monthly_t_total.get(12, 0), 0))
ws.cell(row=row, column=6, value=round(total_t, 0))
ws.cell(row=row, column=7, value=round(total_ca, 1))
ws.cell(row=row, column=8, value='100.0%')
style_total_row(ws, row, 8)

ws.column_dimensions['A'].width = 22
for col in 'BCDEFGH':
    ws.column_dimensions[col].width = 16

# === Sheet 3: Par Région ===
print("Creating sheet 3: Par Région...")
ws = wb.create_sheet("3. Par Région")
ws['A1'] = 'FORECAST Q4 2026 (S3) - DÉTAIL PAR RÉGION × MOIS'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers = ['Région', 'Sept (t)', 'Oct (t)', 'Nov (t)', 'Déc (t)', 'Total Q4 (t)', 'CA (M FCFA)', 'Part CA (%)']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

for reg in ['Ouest', 'Centre', 'Littoral']:
    sub = fcst[fcst['region'] == reg]
    if len(sub) == 0: continue
    monthly_t = sub.groupby('month')['tonnes'].sum().to_dict()
    total_t_reg = sub['tonnes'].sum()
    total_ca_reg = sub['ca_m_fcfa'].sum()
    pct = total_ca_reg / total_ca * 100 if total_ca > 0 else 0
    ws.cell(row=row, column=1, value=reg)
    ws.cell(row=row, column=2, value=round(monthly_t.get(9, 0), 0))
    ws.cell(row=row, column=3, value=round(monthly_t.get(10, 0), 0))
    ws.cell(row=row, column=4, value=round(monthly_t.get(11, 0), 0))
    ws.cell(row=row, column=5, value=round(monthly_t.get(12, 0), 0))
    ws.cell(row=row, column=6, value=round(total_t_reg, 0))
    ws.cell(row=row, column=7, value=round(total_ca_reg, 1))
    ws.cell(row=row, column=8, value=f"{pct:.1f}%")
    style_data_row(ws, row, 8, color=S3_COLOR)
    row += 1

ws.cell(row=row, column=1, value='TOTAL')
ws.cell(row=row, column=2, value=round(monthly_t_total.get(9, 0), 0))
ws.cell(row=row, column=3, value=round(monthly_t_total.get(10, 0), 0))
ws.cell(row=row, column=4, value=round(monthly_t_total.get(11, 0), 0))
ws.cell(row=row, column=5, value=round(monthly_t_total.get(12, 0), 0))
ws.cell(row=row, column=6, value=round(total_t, 0))
ws.cell(row=row, column=7, value=round(total_ca, 1))
ws.cell(row=row, column=8, value='100.0%')
style_total_row(ws, row, 8)

ws.column_dimensions['A'].width = 15
for col in 'BCDEFGH':
    ws.column_dimensions[col].width = 15

# === Sheet 4: Par Agence ===
print("Creating sheet 4: Par Agence...")
ws = wb.create_sheet("4. Par Agence")
ws['A1'] = 'FORECAST Q4 2026 (S3) - DÉTAIL PAR AGENCE'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers = ['Agence', 'Région', 'Tonnes Q4', 'CA (M FCFA)', 'Part CA (%)']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

by_agence = fcst.groupby(['agence', 'region']).agg(
    tonnes=('tonnes', 'sum'),
    ca_m_fcfa=('ca_m_fcfa', 'sum'),
).reset_index().sort_values('ca_m_fcfa', ascending=False)

for _, r in by_agence.iterrows():
    pct = r['ca_m_fcfa'] / total_ca * 100 if total_ca > 0 else 0
    ws.cell(row=row, column=1, value=r['agence'])
    ws.cell(row=row, column=2, value=r['region'])
    ws.cell(row=row, column=3, value=round(r['tonnes'], 0))
    ws.cell(row=row, column=4, value=round(r['ca_m_fcfa'], 1))
    ws.cell(row=row, column=5, value=f"{pct:.1f}%")
    style_data_row(ws, row, 5, color=S3_COLOR)
    row += 1

ws.cell(row=row, column=1, value='TOTAL')
ws.cell(row=row, column=2, value='—')
ws.cell(row=row, column=3, value=round(total_t, 0))
ws.cell(row=row, column=4, value=round(total_ca, 1))
ws.cell(row=row, column=5, value='100.0%')
style_total_row(ws, row, 5)

ws.column_dimensions['A'].width = 18
ws.column_dimensions['B'].width = 12
for col in 'CDE':
    ws.column_dimensions[col].width = 15

# === Sheet 5: Par Produit ===
print("Creating sheet 5: Par Produit...")
ws = wb.create_sheet("5. Par Produit")
ws['A1'] = 'FORECAST Q4 2026 (S3) - DÉTAIL PAR PRODUIT'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers = ['Réf', 'Description', 'Famille', 'Tonnes Q4', 'Sacs 50kg', 'CA (M FCFA)', 'Part CA (%)']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

by_prod = fcst.groupby(['ref', 'family']).agg(
    tonnes=('tonnes', 'sum'),
    sacs_50=('sacs_50', 'sum'),
    ca_m_fcfa=('ca_m_fcfa', 'sum'),
).reset_index().sort_values('ca_m_fcfa', ascending=False)

for _, r in by_prod.iterrows():
    pct = r['ca_m_fcfa'] / total_ca * 100 if total_ca > 0 else 0
    desc = desc_map.get(r['ref'], '')
    ws.cell(row=row, column=1, value=r['ref'])
    ws.cell(row=row, column=2, value=desc)
    ws.cell(row=row, column=3, value=r['family'])
    ws.cell(row=row, column=4, value=round(r['tonnes'], 1))
    ws.cell(row=row, column=5, value=round(r['sacs_50'], 0))
    ws.cell(row=row, column=6, value=round(r['ca_m_fcfa'], 1))
    ws.cell(row=row, column=7, value=f"{pct:.1f}%")
    color = NEW_FAMILY_COLOR if r['family'] == 'COMPLEMENT_ALIMENTAIRE' else S3_COLOR
    style_data_row(ws, row, 7, color=color)
    row += 1

ws.cell(row=row, column=1, value='TOTAL')
ws.cell(row=row, column=2, value='—')
ws.cell(row=row, column=3, value='—')
ws.cell(row=row, column=4, value=round(total_t, 0))
ws.cell(row=row, column=5, value=round(total_sacs, 0))
ws.cell(row=row, column=6, value=round(total_ca, 1))
ws.cell(row=row, column=7, value='100.0%')
style_total_row(ws, row, 7)

ws.column_dimensions['A'].width = 12
ws.column_dimensions['B'].width = 35
ws.column_dimensions['C'].width = 22
for col in 'DEFG':
    ws.column_dimensions[col].width = 14

# === Sheet 6: Par Mois ===
print("Creating sheet 6: Par Mois...")
ws = wb.create_sheet("6. Par Mois")
ws['A1'] = 'FORECAST Q4 2026 (S3) - ÉVOLUTION MENSUELLE PAR FAMILLE'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers = ['Famille', 'Sept (t)', 'Oct (t)', 'Nov (t)', 'Déc (t)', 'Total Q4 (t)', 'CA Sept (M)', 'CA Oct (M)', 'CA Nov (M)', 'CA Déc (M)', 'CA Total (M)']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

for fam in FAMILLES:
    sub = fcst[fcst['family'] == fam]
    if len(sub) == 0: continue
    monthly_t = sub.groupby('month')['tonnes'].sum().to_dict()
    monthly_ca = sub.groupby('month')['ca_m_fcfa'].sum().to_dict()
    total_t_fam = sub['tonnes'].sum()
    total_ca_fam = sub['ca_m_fcfa'].sum()
    ws.cell(row=row, column=1, value=fam)
    ws.cell(row=row, column=2, value=round(monthly_t.get(9, 0), 0))
    ws.cell(row=row, column=3, value=round(monthly_t.get(10, 0), 0))
    ws.cell(row=row, column=4, value=round(monthly_t.get(11, 0), 0))
    ws.cell(row=row, column=5, value=round(monthly_t.get(12, 0), 0))
    ws.cell(row=row, column=6, value=round(total_t_fam, 0))
    ws.cell(row=row, column=7, value=round(monthly_ca.get(9, 0), 1))
    ws.cell(row=row, column=8, value=round(monthly_ca.get(10, 0), 1))
    ws.cell(row=row, column=9, value=round(monthly_ca.get(11, 0), 1))
    ws.cell(row=row, column=10, value=round(monthly_ca.get(12, 0), 1))
    ws.cell(row=row, column=11, value=round(total_ca_fam, 1))
    color = NEW_FAMILY_COLOR if fam == 'COMPLEMENT_ALIMENTAIRE' else S3_COLOR
    style_data_row(ws, row, 11, color=color)
    row += 1

ws.cell(row=row, column=1, value='TOTAL')
for i, m in enumerate([9, 10, 11, 12], 2):
    ws.cell(row=row, column=i, value=round(monthly_t_total.get(m, 0), 0))
ws.cell(row=row, column=6, value=round(total_t, 0))
monthly_ca_total = fcst.groupby('month')['ca_m_fcfa'].sum().to_dict()
for i, m in enumerate([9, 10, 11, 12], 7):
    ws.cell(row=row, column=i, value=round(monthly_ca_total.get(m, 0), 1))
ws.cell(row=row, column=11, value=round(total_ca, 1))
style_total_row(ws, row, 11)

ws.column_dimensions['A'].width = 22
for col in 'BCDEFGHIJK':
    ws.column_dimensions[col].width = 13

# === Sheet 7: Détail complet ===
print("Creating sheet 7: Détail complet...")
ws = wb.create_sheet("7. Détail complet")
ws['A1'] = 'FORECAST Q4 2026 (S3) - DÉTAIL COMPLET (Produit × Agence × Mois)'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers = ['Réf', 'Description', 'Famille', 'Agence', 'Région', 'Mois', 'Année', 'Tonnes', 'Sacs 50kg', 'Prix TTC/sac', 'CA (M FCFA)']
for i, h in enumerate(headers, 1):
    ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

sub_sorted = fcst.sort_values(['family', 'ref', 'agence', 'month'])
for _, r in sub_sorted.iterrows():
    desc = desc_map.get(r['ref'], '')
    ws.cell(row=row, column=1, value=r['ref'])
    ws.cell(row=row, column=2, value=desc)
    ws.cell(row=row, column=3, value=r['family'])
    ws.cell(row=row, column=4, value=r['agence'])
    ws.cell(row=row, column=5, value=r['region'])
    ws.cell(row=row, column=6, value=f"{int(r['month']):02d}/2026")
    ws.cell(row=row, column=7, value=int(r['year']))
    ws.cell(row=row, column=8, value=round(r['tonnes'], 2))
    ws.cell(row=row, column=9, value=round(r['sacs_50'], 1))
    ws.cell(row=row, column=10, value=int(r['prix_ttc_sac']) if r['prix_ttc_sac'] else 0)
    ws.cell(row=row, column=11, value=round(r['ca_m_fcfa'], 2))
    style_data_row(ws, row, 11)
    row += 1

ws.column_dimensions['A'].width = 12
ws.column_dimensions['B'].width = 32
ws.column_dimensions['C'].width = 22
ws.column_dimensions['D'].width = 15
ws.column_dimensions['E'].width = 12
for col in 'FGHIJK':
    ws.column_dimensions[col].width = 13
ws.freeze_panes = 'A4'

# === Sheet 8: Hypothèses ===
print("Creating sheet 8: Hypothèses...")
ws = wb.create_sheet("8. Hypothèses")
ws['A1'] = 'HYPOTHÈSES DU FORECAST Q4 2026 (S3)'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

hypotheses = [
    ('PÉRIODE', 'Septembre - Décembre 2026 (4 mois)'),
    ('SCÉNARIO', 'S3 - Réappro 100% (80 000 sacs au 15/09/2026)'),
    ('MÉTHODE', 'Prophet (TOURTEAUX, CONCENTRES, INGREDIENTS, ALIMENT_COMPLET, COMPLEMENT_ALIMENTAIRE) + Extrapolation moyenne (MATERIEL_ELEVAGE, PREMIX)'),
    ('NIVEAU DE DÉTAIL', '15 modèles Prophet famille × région + 6 extrapolations, désagrégés en produit × agence'),
    ('', ''),
    ('NOUVEAUTÉS VERSION 2', ''),
    ('Données 2023-2026', '44 mois d\'historique (Jan 2023 - Août 2026), 176 576 records'),
    ('Famille COMPLEMENT_ALIMENTAIRE', '9 produits liquides (V300 BELGOKILL 1L + CA001-CA008), 1L=1kg'),
    ('V305 (BELGOKILL 200L)', 'EXCLU du forecast selon demande utilisateur'),
    ('MATERIEL_ELEVAGE', 'Tonnes toujours à 0 (CA only, non exprimable en volume)'),
    ('', ''),
    ('FAMILLES INCLUES (7)', ''),
    ('TOURTEAUX', 'Soja T102, T1021, T1023, T1024 (76% du volume historique)'),
    ('CONCENTRES', 'BELGO Chair/Ponte/Porc C101-C108 (13 références)'),
    ('INGREDIENTS', 'Belgotox, Bicarbonate, Methionine, Lysine, Belgofos, Farine poisson (18 refs)'),
    ('ALIMENT_COMPLET', 'Chick/Piglet Booster CB100-CB201, BELGO Rabbit (5 refs)'),
    ('MATERIEL_ELEVAGE', 'Alvéoles MAT014/MAT011/MAT015/MAT017 + autres refs matériel (26 refs) - CA ONLY'),
    ('PREMIX', 'P102N2, P104N2, P109 (3 refs) - CA ONLY'),
    ('COMPLEMENT_ALIMENTAIRE', 'BELGOKILL V300 1L + BELGO HARMONY/PROTECT/DRY LIT/WATER CLEAN/VIT/THERMO/BIO SELECT/FRESH (9 refs) - 1L=1kg'),
    ('', ''),
    ('FAMILLES EXCLUES', ''),
    ('MAIS', 'M1051, M1052 - Produit opportuniste hors portefeuille BELGOCAM'),
    ('DIVERS', 'Manuels, Pierre à lécher - hors périmètre alimentaire'),
    ('V305 (BELGOKILL 200L)', 'Exclu selon demande utilisateur - seul V300 1L conservé'),
    ('', ''),
    ('SCÉNARIO S3 - DÉTAILS', ''),
    ('Hypothèse réappro', '80 000 sacs (4 000 t) de soja au 15/09/2026'),
    ('Septembre', 'Volume soja × 0.7 (réappro mi-mois)'),
    ('Octobre - Décembre', 'Volume soja × 1.0 (stock reconstitué)'),
    ('', ''),
    ('PRIX', 'Prix août 2026 + soja 25 000 FCFA/sac (actualisé 24/08/2026)'),
    ('STOCK SOJA (BEKOKO)', '71 230 sacs net au 08/08/2026, rupture probable 16/09 sans réappro'),
    ('SAISONNALITÉ Q4', 'Q4 = 35-46% du volume annuel selon les familles'),
    ('LIMITES', '44 mois d\'historique (suffisant pour Prophet yearly seasonality)'),
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
output_path = "/home/z/my-project/download/forecast_q4_2026_S3_volumes_valeurs.xlsx"
wb.save(output_path)
print(f"\n=== EXCEL S3 SAVED ===")
print(f"Path: {output_path}")
print(f"Size: {os.path.getsize(output_path) / 1024:.0f} KB")
print(f"Sheets: {wb.sheetnames}")
