from openpyxl.utils import get_column_letter
"""
Excel - Forecast 2027 complet (12 mois, S3)
VERSION 2 - Avec COMPLEMENT_ALIMENTAIRE et données 2023-2026

Changements:
- 7 familles au lieu de 6 (ajout COMPLEMENT_ALIMENTAIRE)
- Données historiques 2023-2026 (vs 2025-2026 avant)
- 78 produits (vs 69 avant)
- 8 868 lignes détaillées (vs 12 600 avant — certaines familles n'ont pas toutes les agences)
- Sheet 8: Coefficients saisonniers 2024-2025 vs 2027
"""
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import os

fcst = pd.read_csv("/home/z/my-project/scripts/forecast_2027_S3.csv", parse_dates=['date'])
print(f"Loaded {len(fcst)} forecast records")

# Load descriptions from new dataset
desc_df = pd.read_csv("/home/z/my-project/scripts/dataset_2023_2026.csv", low_memory=False)
desc_map = desc_df[['ref', 'description']].drop_duplicates().set_index('ref')['description'].to_dict()

HEAD_FILL = PatternFill('solid', fgColor='1F4E78')
HEAD_FONT = Font(bold=True, color='FFFFFF', size=11)
SUBHEAD_FILL = PatternFill('solid', fgColor='D9E1F2')
SUBHEAD_FONT = Font(bold=True, color='1F4E78', size=11)
TOTAL_FILL = PatternFill('solid', fgColor='FFF2CC')
TOTAL_FONT = Font(bold=True, size=11)
S3_COLOR = 'C6EFCE'
NEW_FAMILY_COLOR = 'FFE699'  # Yellow for COMPLEMENT_ALIMENTAIRE
THIN = Side(border_style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# 7 familles (ordre: ordre alphabétique, mais COMPLEMENT_ALIMENTAIRE mis en évidence à la fin)
FAMILIES = ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'ALIMENT_COMPLET', 'MATERIEL_ELEVAGE', 'PREMIX', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES']
MONTHS = list(range(1, 13))
MONTH_NAMES = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc']
QUARTERS = {'Q1 (Jan-Mar)': [1,2,3], 'Q2 (Avr-Juin)': [4,5,6], 'Q3 (Juil-Sept)': [7,8,9], 'Q4 (Oct-Déc)': [10,11,12]}

def style_header_row(ws, row, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEAD_FILL; cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER

def style_total_row(ws, row, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = TOTAL_FILL; cell.font = TOTAL_FONT; cell.border = BORDER

def style_data_row(ws, row, n_cols, color=None):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.border = BORDER
        if color: cell.fill = PatternFill('solid', fgColor=color)

wb = openpyxl.Workbook()
wb.remove(wb.active)

# === Sheet 1: Synthèse ===
ws = wb.create_sheet("1. Synthèse")
ws['A1'] = 'BELGOCAM SA - Forecast 2027 (Janvier - Décembre)'
ws['A1'].font = Font(bold=True, size=16, color='1F4E78')
ws['A2'] = 'Scénario S3 | Désaisonnalisation effet soja | En cours+Validées inclus | Prix soja 25 000 FCFA/sac | Données 2023-2026'
ws['A2'].font = Font(italic=True, size=10, color='595959')

total_t = fcst['tonnes'].sum()
total_ca = fcst['ca_m_fcfa'].sum()

row = 4
ws.cell(row=row, column=1, value='SYNTHÈSE GLOBALE 2027')
ws.cell(row=row, column=1).font = SUBHEAD_FONT
row += 1

headers = ['Indicateur', 'Valeur', 'Unité']
for i, h in enumerate(headers, 1): ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, 3)
row += 1

global_data = [
    ('Volume total 2027', f"{total_t:,.0f}", 'tonnes'),
    ('CA total 2027', f"{total_ca:,.1f}", 'M FCFA'),
    ('Période forecast', 'Janvier - Décembre 2027', '12 mois'),
    ('Scénario', 'S3 - Réappro soja 100%', '—'),
    ('Nb familles', '8', 'TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, ALIMENT_COMPLET, MATERIEL_ELEVAGE, PREMIX, COMPLEMENT_ALIMENTAIRE'),
    ('Nb produits', f"{fcst['ref'].nunique()}", 'références'),
    ('Nb agences', f"{fcst['agence'].nunique()}", 'agences'),
    ('Méthode', 'Prophet (5 familles) + Extrapolation (MAT+PREMIX)', '—'),
    ('Désaisonnalisation', 'Effet soja Jul-Août 2026 neutralisé (cap S1 avg)', '—'),
    ('En cours + Validées', 'Incluses comme potentielles Livrées', '—'),
    ('Données historiques', '176 576 enregistrements (Jan 2023 - Août 2026)', '44 mois'),
    ('NOUVEAU: COMPLEMENT_ALIMENTAIRE', 'Ajouté au forecast (1L=1kg, proxy BELGOKILL)', '10 produits'),
]
for label, val, unit in global_data:
    ws.cell(row=row, column=1, value=label).font = Font(bold=True)
    ws.cell(row=row, column=2, value=val)
    ws.cell(row=row, column=3, value=unit)
    is_new = 'NOUVEAU' in label
    style_data_row(ws, row, 3, color=NEW_FAMILY_COLOR if is_new else S3_COLOR)
    row += 1

row += 1
ws.cell(row=row, column=1, value='SYNTHÈSE PAR FAMILLE')
ws.cell(row=row, column=1).font = SUBHEAD_FONT
row += 1

headers = ['Famille', 'Volume 2027 (t)', 'CA 2027 (M FCFA)', 'Part CA (%)']
for i, h in enumerate(headers, 1): ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, 4)
row += 1

fam_synth = fcst.groupby('family').agg(tonnes=('tonnes', 'sum'), ca_m_fcfa=('ca_m_fcfa', 'sum')).reindex(FAMILIES)
for fam in FAMILIES:
    if fam in fam_synth.index:
        t = fam_synth.loc[fam, 'tonnes']
        ca = fam_synth.loc[fam, 'ca_m_fcfa']
        pct = ca / total_ca * 100 if total_ca > 0 else 0
        ws.cell(row=row, column=1, value=fam)
        ws.cell(row=row, column=2, value=round(t, 0))
        ws.cell(row=row, column=3, value=round(ca, 1))
        ws.cell(row=row, column=4, value=f"{pct:.1f}%")
        color = NEW_FAMILY_COLOR if fam in ('COMPLEMENT_ALIMENTAIRE', 'ALVEOLES') else S3_COLOR
        style_data_row(ws, row, 4, color=color)
        row += 1
ws.cell(row=row, column=1, value='TOTAL')
ws.cell(row=row, column=2, value=round(total_t, 0))
ws.cell(row=row, column=3, value=round(total_ca, 1))
ws.cell(row=row, column=4, value='100.0%')
style_total_row(ws, row, 4)

row += 2
ws.cell(row=row, column=1, value='SYNTHÈSE PAR TRIMESTRE')
ws.cell(row=row, column=1).font = SUBHEAD_FONT
row += 1
headers = ['Trimestre', 'Volume (t)', 'CA (M FCFA)', 'Part CA (%)']
for i, h in enumerate(headers, 1): ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, 4)
row += 1
for qname, qmonths in QUARTERS.items():
    sub = fcst[fcst['month'].isin(qmonths)]
    t = sub['tonnes'].sum()
    ca = sub['ca_m_fcfa'].sum()
    pct = ca / total_ca * 100 if total_ca > 0 else 0
    ws.cell(row=row, column=1, value=qname)
    ws.cell(row=row, column=2, value=round(t, 0))
    ws.cell(row=row, column=3, value=round(ca, 1))
    ws.cell(row=row, column=4, value=f"{pct:.1f}%")
    style_data_row(ws, row, 4, color=S3_COLOR)
    row += 1
ws.cell(row=row, column=1, value='TOTAL')
ws.cell(row=row, column=2, value=round(total_t, 0))
ws.cell(row=row, column=3, value=round(total_ca, 1))
ws.cell(row=row, column=4, value='100.0%')
style_total_row(ws, row, 4)

ws.column_dimensions['A'].width = 35
ws.column_dimensions['B'].width = 20
ws.column_dimensions['C'].width = 20
ws.column_dimensions['D'].width = 12

# === Sheet 2: Par Famille × Mois ===
ws = wb.create_sheet("2. Par Famille × Mois")
ws['A1'] = 'FORECAST 2027 - DÉTAIL PAR FAMILLE × MOIS'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers = ['Famille'] + MONTH_NAMES + ['Total (t)', 'CA (M FCFA)']
for i, h in enumerate(headers, 1): ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

for fam in FAMILIES:
    sub = fcst[fcst['family'] == fam]
    if len(sub) == 0: continue
    monthly_t = sub.groupby('month')['tonnes'].sum().to_dict()
    monthly_ca = sub.groupby('month')['ca_m_fcfa'].sum().to_dict()
    total_t_fam = sub['tonnes'].sum()
    total_ca_fam = sub['ca_m_fcfa'].sum()
    ws.cell(row=row, column=1, value=fam)
    for i, m in enumerate(MONTHS, 2):
        ws.cell(row=row, column=i, value=round(monthly_t.get(m, 0), 0))
    ws.cell(row=row, column=14, value=round(total_t_fam, 0))
    ws.cell(row=row, column=15, value=round(total_ca_fam, 1))
    color = NEW_FAMILY_COLOR if fam in ('COMPLEMENT_ALIMENTAIRE', 'ALVEOLES') else S3_COLOR
    style_data_row(ws, row, 15, color=color)
    row += 1
ws.cell(row=row, column=1, value='TOTAL')
for i, m in enumerate(MONTHS, 2):
    ws.cell(row=row, column=i, value=round(fcst.groupby('month')['tonnes'].sum().get(m, 0), 0))
ws.cell(row=row, column=14, value=round(total_t, 0))
ws.cell(row=row, column=15, value=round(total_ca, 1))
style_total_row(ws, row, 15)

ws.column_dimensions['A'].width = 25
for col in [get_column_letter(c) for c in range(2, 16)]:
    ws.column_dimensions[col].width = 10

# === Sheet 3: Par Région × Mois ===
ws = wb.create_sheet("3. Par Région × Mois")
ws['A1'] = 'FORECAST 2027 - DÉTAIL PAR RÉGION × MOIS'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers = ['Région'] + MONTH_NAMES + ['Total (t)', 'CA (M FCFA)']
for i, h in enumerate(headers, 1): ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1
for reg in ['Ouest', 'Centre', 'Littoral']:
    sub = fcst[fcst['region'] == reg]
    monthly_t = sub.groupby('month')['tonnes'].sum().to_dict()
    ws.cell(row=row, column=1, value=reg)
    for i, m in enumerate(MONTHS, 2):
        ws.cell(row=row, column=i, value=round(monthly_t.get(m, 0), 0))
    ws.cell(row=row, column=14, value=round(sub['tonnes'].sum(), 0))
    ws.cell(row=row, column=15, value=round(sub['ca_m_fcfa'].sum(), 1))
    style_data_row(ws, row, 15, color=S3_COLOR)
    row += 1
ws.cell(row=row, column=1, value='TOTAL')
for i, m in enumerate(MONTHS, 2):
    ws.cell(row=row, column=i, value=round(fcst.groupby('month')['tonnes'].sum().get(m, 0), 0))
ws.cell(row=row, column=14, value=round(total_t, 0))
ws.cell(row=row, column=15, value=round(total_ca, 1))
style_total_row(ws, row, 15)
ws.column_dimensions['A'].width = 15
for col in [get_column_letter(c) for c in range(2, 16)]:
    ws.column_dimensions[col].width = 10

# === Sheet 4: Par Agence ===
ws = wb.create_sheet("4. Par Agence")
ws['A1'] = 'FORECAST 2027 - DÉTAIL PAR AGENCE'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
row = 3
headers = ['Agence', 'Région', 'Volume 2027 (t)', 'CA 2027 (M FCFA)', 'Part CA (%)']
for i, h in enumerate(headers, 1): ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, 5)
row += 1
by_ag = fcst.groupby(['agence', 'region']).agg(tonnes=('tonnes', 'sum'), ca_m_fcfa=('ca_m_fcfa', 'sum')).reset_index().sort_values('ca_m_fcfa', ascending=False)
for _, r in by_ag.iterrows():
    pct = r['ca_m_fcfa'] / total_ca * 100 if total_ca > 0 else 0
    ws.cell(row=row, column=1, value=r['agence'])
    ws.cell(row=row, column=2, value=r['region'])
    ws.cell(row=row, column=3, value=round(r['tonnes'], 0))
    ws.cell(row=row, column=4, value=round(r['ca_m_fcfa'], 1))
    ws.cell(row=row, column=5, value=f"{pct:.1f}%")
    style_data_row(ws, row, 5, color=S3_COLOR)
    row += 1
ws.cell(row=row, column=1, value='TOTAL')
ws.cell(row=row, column=3, value=round(total_t, 0))
ws.cell(row=row, column=4, value=round(total_ca, 1))
ws.cell(row=row, column=5, value='100.0%')
style_total_row(ws, row, 5)
ws.column_dimensions['A'].width = 18; ws.column_dimensions['B'].width = 12
for col in 'CDE': ws.column_dimensions[col].width = 18

# === Sheet 5: Par Produit ===
ws = wb.create_sheet("5. Par Produit")
ws['A1'] = 'FORECAST 2027 - DÉTAIL PAR PRODUIT'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
row = 3
headers = ['Réf', 'Description', 'Famille', 'Volume 2027 (t)', 'CA 2027 (M FCFA)', 'Part CA (%)']
for i, h in enumerate(headers, 1): ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, 6)
row += 1
by_prod = fcst.groupby(['ref', 'family']).agg(tonnes=('tonnes', 'sum'), ca_m_fcfa=('ca_m_fcfa', 'sum')).reset_index().sort_values('ca_m_fcfa', ascending=False)
for _, r in by_prod.iterrows():
    pct = r['ca_m_fcfa'] / total_ca * 100 if total_ca > 0 else 0
    ws.cell(row=row, column=1, value=r['ref'])
    ws.cell(row=row, column=2, value=desc_map.get(r['ref'], ''))
    ws.cell(row=row, column=3, value=r['family'])
    ws.cell(row=row, column=4, value=round(r['tonnes'], 1))
    ws.cell(row=row, column=5, value=round(r['ca_m_fcfa'], 1))
    ws.cell(row=row, column=6, value=f"{pct:.1f}%")
    color = NEW_FAMILY_COLOR if r['family'] in ('COMPLEMENT_ALIMENTAIRE', 'ALVEOLES') else S3_COLOR
    style_data_row(ws, row, 6, color=color)
    row += 1
ws.cell(row=row, column=1, value='TOTAL')
ws.cell(row=row, column=4, value=round(total_t, 0))
ws.cell(row=row, column=5, value=round(total_ca, 1))
ws.cell(row=row, column=6, value='100.0%')
style_total_row(ws, row, 6)
ws.column_dimensions['A'].width = 12; ws.column_dimensions['B'].width = 35; ws.column_dimensions['C'].width = 25
for col in 'DEF': ws.column_dimensions[col].width = 18

# === Sheet 6: Détail complet ===
ws = wb.create_sheet("6. Détail complet")
ws['A1'] = 'FORECAST 2027 - DÉTAIL COMPLET (Produit × Agence × Mois)'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
row = 3
headers = ['Réf', 'Description', 'Famille', 'Agence', 'Région', 'Mois', 'Année', 'Tonnes', 'Sacs 50kg', 'Prix TTC/sac', 'CA (M FCFA)']
for i, h in enumerate(headers, 1): ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, 11)
row += 1
for _, r in fcst.sort_values(['family', 'ref', 'agence', 'month']).iterrows():
    ws.cell(row=row, column=1, value=r['ref'])
    ws.cell(row=row, column=2, value=desc_map.get(r['ref'], ''))
    ws.cell(row=row, column=3, value=r['family'])
    ws.cell(row=row, column=4, value=r['agence'])
    ws.cell(row=row, column=5, value=r['region'])
    ws.cell(row=row, column=6, value=f"{int(r['month']):02d}/2027")
    ws.cell(row=row, column=7, value=int(r['year']))
    ws.cell(row=row, column=8, value=round(r['tonnes'], 2))
    ws.cell(row=row, column=9, value=round(r['sacs_50'], 1))
    ws.cell(row=row, column=10, value=int(r['prix_ttc_sac']) if r['prix_ttc_sac'] else 0)
    ws.cell(row=row, column=11, value=round(r['ca_m_fcfa'], 2))
    style_data_row(ws, row, 11)
    row += 1
ws.column_dimensions['A'].width = 12; ws.column_dimensions['B'].width = 32; ws.column_dimensions['C'].width = 25
ws.column_dimensions['D'].width = 15; ws.column_dimensions['E'].width = 12
for col in 'FGHIJK': ws.column_dimensions[col].width = 13
ws.freeze_panes = 'A4'

# === Sheet 7: Hypothèses ===
ws = wb.create_sheet("7. Hypothèses")
ws['A1'] = 'HYPOTHÈSES DU FORECAST 2027 (VERSION 2 - AVEC COMPLEMENT ALIMENTAIRE)'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
hyp = [
    ('PÉRIODE', 'Janvier - Décembre 2027 (12 mois)'),
    ('SCÉNARIO', 'S3 - Réappro soja 100% (situation normale)'),
    ('MÉTHODE', 'Prophet (5 familles alimentaires) + Extrapolation (MAT+PREMIX)'),
    ('', ''),
    ('NOUVEAUTÉS VS VERSION PRÉCÉDENTE', ''),
    ('Données historiques', '44 mois (Jan 2023 - Août 2026). Années 2021-2022 exclues, 2023 incluse.'),
    ('Famille COMPLEMENT_ALIMENTAIRE', 'AJOUTÉE: 10 produits (BELGOKILL V300/V305, BELGO HARMONY, BELGO PROTECT, etc.)'),
    ('Conversion 1L=1kg', 'Pour COMPLEMENT_ALIMENTAIRE: 1 litre = 1 kg (qte = kg, tonnes = qte/1000)'),
    ('Proxy BELGOKILL', 'BELGOKILL (V300 1L + V305 200L) représente 38% du CA famille — tendance utilisée comme proxy'),
    ('', ''),
    ('INNOVATIONS CONSERVÉES', ''),
    ('En cours + Validées', 'Commandes En cours et Validées d\'août 2026 incluses comme potentielles Livrées'),
    ('Désaisonnalisation soja', 'Volumes soja Jul-Août 2026 neutralisés (cap à moyenne S1 2026) pour éviter le biais de la rupture concurrente'),
    ('Prix soja actualisé', '25 000 FCFA/sac (prix au 24/08/2026) au lieu de 20 600'),
    ('Forecast 12 mois', 'Période complète Jan-Déc 2027 (vs 4 mois pour Q4 2026)'),
    ('', ''),
    ('FAMILLES INCLUSES (7)', 'TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, ALIMENT_COMPLET, MATERIEL_ELEVAGE, PREMIX, COMPLEMENT_ALIMENTAIRE'),
    ('FAMILLES EXCLUES', 'MAIS (produit opportuniste), DIVERS'),
    ('PRODUITS OPPORTUNISTES', 'MAÏs (M1051, M1052) — exclus du périmètre'),
    ('', ''),
    ('DONNÉES HISTORIQUES', '176 576 enregistrements (Jan 2023 - Août 2026) + En cours/Validées août'),
    ('LIMITES', '44 mois d\'historique (Prophet recommande 2+ ans pour yearly seasonality — OK)'),
    ('PRODUITS COMPLEMENT_ALIMENTAIRE', 'V300 (BELGOKILL 1L), V305 (BELGOKILL 200L), CA001-CA008 (BELGO xxx 1L)'),
    ('MISE À JOUR', 'Pipeline: python scripts/forecast_2027_v3.py + scripts/forecast_2027_excel_v2.py'),
]
row = 3
for label, value in hyp:
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
ws.column_dimensions['A'].width = 35; ws.column_dimensions['B'].width = 95

# === Sheet 8: Saisonnalité (NOUVEAU) ===
ws = wb.create_sheet("8. Saisonnalité")
ws['A1'] = 'COEFFICIENTS SAISONNIERS - 2025 (HISTORIQUE) VS 2027 (FORECAST)'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')

# Calculate seasonal coefficients from 2025 historical data
df_hist = pd.read_csv("/home/z/my-project/scripts/dataset_2023_2026.csv", parse_dates=['date'], low_memory=False)
df_2025 = df_hist[df_hist['date'].dt.year == 2025]
df_2025_fam = df_2025.groupby(['family', df_2025['date'].dt.month])['tonnes'].sum().reset_index()
df_2025_fam.columns = ['family', 'month', 'tonnes']
df_2025_avg = df_2025_fam.groupby('family')['tonnes'].mean().reset_index()
df_2025_avg.columns = ['family', 'avg_tonnes']

# Forecast 2027 seasonal
df_2027_fam = fcst.groupby(['family', 'month'])['tonnes'].sum().reset_index()
df_2027_avg = df_2027_fam.groupby('family')['tonnes'].mean().reset_index()
df_2027_avg.columns = ['family', 'avg_tonnes']

row = 3
ws.cell(row=row, column=1, value='Coefficient = volume mensuel / moyenne annuelle (1.0 = moyenne)')
ws.cell(row=row, column=1).font = Font(italic=True, color='595959')
row += 2

headers = ['Famille', 'Année'] + MONTH_NAMES + ['Moyenne (t)']
for i, h in enumerate(headers, 1): ws.cell(row=row, column=i, value=h)
style_header_row(ws, row, len(headers))
row += 1

for fam in FAMILIES:
    # 2025 historical
    sub_2025 = df_2025_fam[df_2025_fam['family'] == fam]
    avg_2025 = df_2025_avg[df_2025_avg['family'] == fam]['avg_tonnes'].values
    avg_2025 = avg_2025[0] if len(avg_2025) > 0 and avg_2025[0] > 0 else 1
    
    ws.cell(row=row, column=1, value=fam)
    ws.cell(row=row, column=2, value='2025 (hist)')
    for i, m in enumerate(MONTHS, 3):
        v = sub_2025[sub_2025['month'] == m]['tonnes'].values
        v = v[0] if len(v) > 0 else 0
        coef = v / avg_2025 if avg_2025 > 0 else 0
        ws.cell(row=row, column=i, value=round(coef, 2))
    ws.cell(row=row, column=15, value=round(avg_2025, 0))
    style_data_row(ws, row, 15, color=S3_COLOR)
    row += 1
    
    # 2027 forecast
    sub_2027 = df_2027_fam[df_2027_fam['family'] == fam]
    avg_2027 = df_2027_avg[df_2027_avg['family'] == fam]['avg_tonnes'].values
    avg_2027 = avg_2027[0] if len(avg_2027) > 0 and avg_2027[0] > 0 else 1
    
    ws.cell(row=row, column=1, value='')
    ws.cell(row=row, column=2, value='2027 (fcst)')
    for i, m in enumerate(MONTHS, 3):
        v = sub_2027[sub_2027['month'] == m]['tonnes'].values
        v = v[0] if len(v) > 0 else 0
        coef = v / avg_2027 if avg_2027 > 0 else 0
        ws.cell(row=row, column=i, value=round(coef, 2))
    ws.cell(row=row, column=15, value=round(avg_2027, 0))
    color = NEW_FAMILY_COLOR if fam in ('COMPLEMENT_ALIMENTAIRE', 'ALVEOLES') else S3_COLOR
    style_data_row(ws, row, 15, color=color)
    row += 1

ws.cell(row=row, column=1, value='Note: Coefficients > 1.0 indiquent un mois avec volume supérieur à la moyenne annuelle.')
ws.cell(row=row, column=1).font = Font(italic=True, color='595959')
row += 1
ws.cell(row=row, column=1, value='Pour TOURTEAUX: pic Oct-Déc (Q4) = saisonnalité annuelle forte.')
ws.cell(row=row, column=1).font = Font(italic=True, color='595959')
row += 1
ws.cell(row=row, column=1, value='Pour COMPLEMENT_ALIMENTAIRE: tendance BELGOKILL (V300) dominante, 38% du CA famille.')
ws.cell(row=row, column=1).font = Font(italic=True, color='595959')

ws.column_dimensions['A'].width = 25
ws.column_dimensions['B'].width = 14
for col in [get_column_letter(c) for c in range(3, 16)]:
    ws.column_dimensions[col].width = 8

output_path = "/home/z/my-project/download/forecast_2027_S3_volumes_valeurs.xlsx"
wb.save(output_path)
print(f"\n=== EXCEL 2027 V2 SAVED ===")
print(f"Path: {output_path}")
print(f"Size: {os.path.getsize(output_path) / 1024:.0f} KB")
print(f"Sheets: {wb.sheetnames}")
print(f"  7 familles: {FAMILIES}")
print(f"  Total: {total_t:,.0f} t, {total_ca:,.1f} M FCFA")
print(f"  {fcst['ref'].nunique()} produits, {fcst['agence'].nunique()} agences, {len(fcst)} lignes")
