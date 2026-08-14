"""
BELGOCAM SA — Modèle Prophet de prédiction des ventes
Prédiction TOURTEAUX et CONCENTRÉS pour S2 2026 (Sep-Déc)

Utilise 20 mois de données (Jan 2025 → Août 2026).
Intègre la hausse de prix du soja (+2 000 FCFA/sac le 23/07/2026) comme "holiday/event".

Output: /home/z/my-project/download/predictions_prophet.xlsx (3 feuilles)
  - Prédictions TOURTEAUX
  - Prédictions CONCENTRÉS
  - Synthèse vs objectifs
"""
import openpyxl
from collections import defaultdict
import json
import datetime
import pandas as pd
import numpy as np
from prophet import Prophet
from prophet.plot import plot_plotly, plot_components_plotly
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# ===== CONFIG =====
SOJA_REFS = {'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25}
CONC_REFS = {'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
             'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50}
PROD_CAT = json.load(open('/home/z/my-project/scripts/product_category_map.json'))
MANUAL_WEIGHTS = {'M1051': 50.0, 'CF101': 1.0, 'S101': 1.0}
ING_WEIGHTS = {
    'I105': 25, 'I1051': 1, 'I106': 25, 'I1061': 1, 'I107': 25, 'I1071': 1,
    'F114': 50, 'F1142': 25, 'F1145': 50, 'F1146': 25, 'F1147': 1,
    'B100': 25, 'B1001': 1, 'CF1012': 50, 'S101': 1,
    'E101': 25, 'E1011': 1, 'E1014': 5, 'P105': 25, 'P1051': 1, 'P1053': 5,
}

def get_weight(ref, desc=''):
    if ref in MANUAL_WEIGHTS: return MANUAL_WEIGHTS[ref]
    if ref in SOJA_REFS: return SOJA_REFS[ref]
    if ref in CONC_REFS: return CONC_REFS[ref]
    if ref in ING_WEIGHTS: return ING_WEIGHTS[ref]
    if ref in PROD_CAT:
        cat = PROD_CAT[ref]
        if cat == 'PREMIX':
            d = str(desc or '').upper()
            if '25KG' in d.replace(' ', '') or '25 KG' in d: return 25
            elif '5KG' in d.replace(' ', '') or ' 5 KG' in d: return 5
            elif '1KG' in d.replace(' ', '') or ' 1 KG' in d: return 1
            else: return 25
        elif cat == 'ALIMENT COMPLET':
            d = str(desc or '').upper()
            if '50KG' in d.replace(' ', '') or '50 KG' in d: return 50
            elif '25KG' in d.replace(' ', '') or '25 KG' in d: return 25
            else: return 25
        else: return 50
    return None

def get_category(ref):
    if ref in SOJA_REFS: return 'TOURTEAUX'
    if ref in CONC_REFS: return 'CONCENTRES'
    if ref == 'M1051': return 'MAIS'
    return PROD_CAT.get(ref, 'AUTRE')

def parse_date(val):
    if isinstance(val, datetime.datetime):
        return val.strftime('%Y-%m-%d')
    if isinstance(val, datetime.date):
        return val.strftime('%Y-%m-%d')
    s = str(val or '')
    if len(s) >= 10:
        # Convert DD/MM/YYYY to YYYY-MM-DD
        parts = s[:10].split('/')
        if len(parts) == 3:
            return f'{parts[2]}-{parts[1]}-{parts[0]}'
    return None


# ===== 1. CHARGER LES DONNÉES =====
print("=" * 80)
print("1. CHARGEMENT DES DONNÉES (Jan 2025 → Août 2026)")
print("=" * 80)

daily_data = defaultdict(lambda: {'TOURTEAUX': 0, 'CONCENTRES': 0})  # date_str -> cat -> kg

# --- 2025 (15 cols, header row 1, no État column, date as datetime) ---
print("  Chargement 2025...")
wb2025 = openpyxl.load_workbook('/home/z/my-project/upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx', read_only=True)
ws2025 = wb2025['Feuil1']
count = 0
for r in ws2025.iter_rows(min_row=2, values_only=True):
    if not r or len(r) < 6 or r[0] == 'Total': continue
    ref = r[0]
    if not ref: continue
    qte = r[2] or 0
    desc = r[1] or ''
    date_str = parse_date(r[5])
    if not date_str: continue
    weight = get_weight(ref, desc)
    cat = get_category(ref)
    if weight and cat in ('TOURTEAUX', 'CONCENTRES'):
        kg = qte * weight
        daily_data[date_str][cat] += kg
        count += 1
print(f"    {count} lignes")

# --- S1 2026 (18 cols, header row 2, État=col15) ---
print("  Chargement S1 2026...")
wb_s1 = openpyxl.load_workbook('/home/z/my-project/upload/ventes janv a juin 2026.xlsx', read_only=True)
count = 0
for sn in wb_s1.sheetnames:
    ws = wb_s1[sn]
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) < 18 or r[0] == 'Total': continue
        if r[15] != 'Livrée': continue
        ref = r[0]
        if not ref: continue
        qte = r[2] or 0
        desc = r[1] or ''
        date_str = parse_date(r[6])
        if not date_str: continue
        weight = get_weight(ref, desc)
        cat = get_category(ref)
        if weight and cat in ('TOURTEAUX', 'CONCENTRES'):
            kg = qte * weight
            daily_data[date_str][cat] += kg
            count += 1
print(f"    {count} lignes")

# --- Juillet 2026 (16 cols, header row 2, État=col13) ---
print("  Chargement Juillet 2026...")
wb_juil = openpyxl.load_workbook('/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (9).xlsx', read_only=True)
ws_juil = wb_juil['Sheet 1']
header_j = list(ws_juil.iter_rows(min_row=2, max_row=2, values_only=True))[0]
etat_j = 13 if len(header_j) <= 16 else 15
count = 0
for r in ws_juil.iter_rows(min_row=3, values_only=True):
    if not r or len(r) < 14 or r[0] == 'Total': continue
    if r[etat_j] != 'Livrée': continue
    ref = r[0]
    if not ref: continue
    qte = r[2] or 0
    desc = r[1] or ''
    date_str = parse_date(r[6])
    if not date_str: continue
    weight = get_weight(ref, desc)
    cat = get_category(ref)
    if weight and cat in ('TOURTEAUX', 'CONCENTRES'):
        kg = qte * weight
        daily_data[date_str][cat] += kg
        count += 1
print(f"    {count} lignes")

# --- Août 2026 (16 cols, exclure 14/08 matinal) ---
print("  Chargement Août 2026...")
wb_aout = openpyxl.load_workbook('/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (15).xlsx', read_only=True)
ws_aout = wb_aout['Sheet 1']
header_a = list(ws_aout.iter_rows(min_row=2, max_row=2, values_only=True))[0]
etat_a = 13 if len(header_a) <= 16 else 15
count = 0
for r in ws_aout.iter_rows(min_row=3, values_only=True):
    if not r or len(r) < 14 or r[0] == 'Total': continue
    if r[etat_a] != 'Livrée': continue
    date_str = parse_date(r[6])
    if date_str == '2026-08-14': continue  # Exclude matinal
    ref = r[0]
    if not ref: continue
    qte = r[2] or 0
    desc = r[1] or ''
    if not date_str: continue
    weight = get_weight(ref, desc)
    cat = get_category(ref)
    if weight and cat in ('TOURTEAUX', 'CONCENTRES'):
        kg = qte * weight
        daily_data[date_str][cat] += kg
        count += 1
print(f"    {count} lignes")

# ===== 2. PRÉPARER LES DATAFRAMES PROPHET =====
print("\n" + "=" * 80)
print("2. PRÉPARATION DES DONNÉES PROPHET")
print("=" * 80)

# Convertir en DataFrame avec dates continues (fill 0 for missing days)
all_dates = sorted(daily_data.keys())
start_date = datetime.date(2025, 1, 1)
end_date = datetime.date(2026, 8, 13)  # Dernier jour complet août

# Fill missing dates with 0
dates_list = []
current = start_date
while current <= end_date:
    ds = current.strftime('%Y-%m-%d')
    dates_list.append({
        'ds': ds,
        'TOURTEAUX': daily_data.get(ds, {}).get('TOURTEAUX', 0) / 1000,  # tonnes
        'CONCENTRES': daily_data.get(ds, {}).get('CONCENTRES', 0) / 1000,  # tonnes
        'prix_soja': 1 if ds >= '2026-07-23' else 0,  # Régresseur: hausse prix
    })
    current += datetime.timedelta(days=1)

df = pd.DataFrame(dates_list)
df['ds'] = pd.to_datetime(df['ds'])

# Prophet doesn't know about Sundays = 0 sales. We need to tell it.
# Method: set capacity to 0 on Sundays (floor=0, cap=0) so Prophet predicts ~0
# Better method: use a custom regressor "is_sunday" and add it as a regressor
df['is_sunday'] = df['ds'].dt.weekday.eq(6).astype(int)  # 1 if Sunday, 0 otherwise
# Also add is_saturday (lower sales on Saturdays)
df['is_saturday'] = df['ds'].dt.weekday.eq(5).astype(int)

print(f"  DataFrame: {len(df)} jours")
print(f"  Période: {df['ds'].min().date()} → {df['ds'].max().date()}")
print(f"  TOURTEAUX: moy {df['TOURTEAUX'].mean():.1f} t/j, max {df['TOURTEAUX'].max():.1f} t")
print(f"  CONCENTRES: moy {df['CONCENTRES'].mean():.1f} t/j, max {df['CONCENTRES'].max():.1f} t")

# Holidays: hausse de prix
holidays = pd.DataFrame([
    {'holiday': 'hausse_prix_soja_1000', 'ds': pd.Timestamp('2026-07-01'),
     'lower_window': 0, 'upper_window': 30, 'seasonality': False},
    {'holiday': 'hausse_prix_soja_2000', 'ds': pd.Timestamp('2026-07-23'),
     'lower_window': 0, 'upper_window': 20, 'seasonality': False},
])

# ===== 3. MODÈLE TOURTEAUX =====
print("\n" + "=" * 80)
print("3. MODÈLE PROPHET — TOURTEAUX")
print("=" * 80)

df_soja = df[['ds', 'TOURTEAUX', 'prix_soja', 'is_sunday', 'is_saturday']].rename(columns={'TOURTEAUX': 'y'})

m_soja = Prophet(
    holidays=holidays,
    weekly_seasonality=True,
    yearly_seasonality=True,
    daily_seasonality=False,
    changepoint_prior_scale=0.05,
    seasonality_prior_scale=10,
    interval_width=0.8,
)
m_soja.add_regressor('prix_soja')
m_soja.add_regressor('is_sunday')
m_soja.add_regressor('is_saturday')
m_soja.fit(df_soja)

# Forecast 150 jours (jusqu'au 31/12/2026)
future_soja = m_soja.make_future_dataframe(periods=150, freq='D')
future_soja['prix_soja'] = 1  # Hausse prix active
future_soja['is_sunday'] = future_soja['ds'].dt.weekday.eq(6).astype(int)
future_soja['is_saturday'] = future_soja['ds'].dt.weekday.eq(5).astype(int)
forecast_soja = m_soja.predict(future_soja)

# Clamp negative predictions to 0
forecast_soja['yhat'] = forecast_soja['yhat'].clip(lower=0)
forecast_soja['yhat_lower'] = forecast_soja['yhat_lower'].clip(lower=0)

# Force Sundays to 0 (BELGOCAM doesn't sell on Sundays)
sunday_mask_soja = forecast_soja['ds'].dt.weekday.eq(6)
forecast_soja.loc[sunday_mask_soja, 'yhat'] = 0
forecast_soja.loc[sunday_mask_soja, 'yhat_lower'] = 0
forecast_soja.loc[sunday_mask_soja, 'yhat_upper'] = 0

print(f"  Forecast: {len(forecast_soja)} jours")
print(f"  Prédiction Sep-Déc 2026:")
future_soja_only = forecast_soja[forecast_soja['ds'] >= pd.Timestamp('2026-09-01')]
monthly_pred_soja = future_soja_only.groupby(future_soja_only['ds'].dt.to_period('M')).agg({
    'yhat': 'sum', 'yhat_lower': 'sum', 'yhat_upper': 'sum'
}).reset_index()
for _, row in monthly_pred_soja.iterrows():
    print(f"    {row['ds']}: {row['yhat']:.0f} t (IC80: {row['yhat_lower']:.0f}-{row['yhat_upper']:.0f})")

# ===== 4. MODÈLE CONCENTRÉS =====
print("\n" + "=" * 80)
print("4. MODÈLE PROPHET — CONCENTRÉS")
print("=" * 80)

df_conc = df[['ds', 'CONCENTRES', 'prix_soja', 'is_sunday', 'is_saturday']].rename(columns={'CONCENTRES': 'y'})

m_conc = Prophet(
    holidays=holidays,
    weekly_seasonality=True,
    yearly_seasonality=True,
    daily_seasonality=False,
    changepoint_prior_scale=0.05,
    seasonality_prior_scale=10,
    interval_width=0.8,
)
m_conc.add_regressor('prix_soja')
m_conc.add_regressor('is_sunday')
m_conc.add_regressor('is_saturday')
m_conc.fit(df_conc)

future_conc = m_conc.make_future_dataframe(periods=150, freq='D')
future_conc['prix_soja'] = 1
future_conc['is_sunday'] = future_conc['ds'].dt.weekday.eq(6).astype(int)
future_conc['is_saturday'] = future_conc['ds'].dt.weekday.eq(5).astype(int)
forecast_conc = m_conc.predict(future_conc)

# Clamp negative predictions to 0
forecast_conc['yhat'] = forecast_conc['yhat'].clip(lower=0)
forecast_conc['yhat_lower'] = forecast_conc['yhat_lower'].clip(lower=0)

# Force Sundays to 0 (BELGOCAM doesn't sell on Sundays)
sunday_mask_conc = forecast_conc['ds'].dt.weekday.eq(6)
forecast_conc.loc[sunday_mask_conc, 'yhat'] = 0
forecast_conc.loc[sunday_mask_conc, 'yhat_lower'] = 0
forecast_conc.loc[sunday_mask_conc, 'yhat_upper'] = 0

print(f"  Forecast: {len(forecast_conc)} jours")
print(f"  Prédiction Sep-Déc 2026:")
future_conc_only = forecast_conc[forecast_conc['ds'] >= pd.Timestamp('2026-09-01')]
monthly_pred_conc = future_conc_only.groupby(future_conc_only['ds'].dt.to_period('M')).agg({
    'yhat': 'sum', 'yhat_lower': 'sum', 'yhat_upper': 'sum'
}).reset_index()
for _, row in monthly_pred_conc.iterrows():
    print(f"    {row['ds']}: {row['yhat']:.0f} t (IC80: {row['yhat_lower']:.0f}-{row['yhat_upper']:.0f})")

# ===== 5. SYNTHÈSE VS OBJECTIFS =====
print("\n" + "=" * 80)
print("5. SYNTHÈSE — PRÉDICTIONS vs OBJECTIFS (S2 2026)")
print("=" * 80)

# Objectifs recalibrés S2
with open('/home/z/my-project/scripts/s2_recaled_objectives.json') as f:
    s2_obj = json.load(f)

obj_soja = s2_obj['global_s2_recaled']['TOURTEAUX']
obj_conc = s2_obj['global_s2_recaled']['CONCENTRES']

print(f"\n{'Mois':8s} {'TOURTEAUX':>30s} {'CONCENTRES':>30s}")
print(f"{'':8s} {'Prédiction (t)':>10s} {'Obj (t)':>8s} {'%':>6s}   {'Prédiction (t)':>10s} {'Obj (t)':>8s} {'%':>6s}")
print("-" * 80)

for mois_num in [9, 10, 11, 12]:
    mois_str = f'2026-{mois_num:02d}'
    # TOURTEAUX
    soja_rows = forecast_soja[forecast_soja['ds'].dt.to_period('M').astype(str) == mois_str]
    soja_pred = soja_rows['yhat'].sum() if len(soja_rows) > 0 else 0
    soja_obj = obj_soja.get(str(mois_num), 0)
    soja_pct = soja_pred / soja_obj * 100 if soja_obj > 0 else 0
    
    # CONCENTRES
    conc_rows = forecast_conc[forecast_conc['ds'].dt.to_period('M').astype(str) == mois_str]
    conc_pred = conc_rows['yhat'].sum() if len(conc_rows) > 0 else 0
    conc_obj = obj_conc.get(str(mois_num), 0)
    conc_pct = conc_pred / conc_obj * 100 if conc_obj > 0 else 0
    
    month_name = ['Sep', 'Oct', 'Nov', 'Déc'][mois_num - 9]
    print(f"{month_name:8s} {soja_pred:10.0f} {soja_obj:8.0f} {soja_pct:5.0f}%   {conc_pred:10.0f} {conc_obj:8.0f} {conc_pct:5.0f}%")

# Total S2
total_soja_pred = sum(forecast_soja[forecast_soja['ds'] >= pd.Timestamp('2026-07-01')]['yhat'])
total_conc_pred = sum(forecast_conc[forecast_conc['ds'] >= pd.Timestamp('2026-07-01')]['yhat'])
total_soja_obj = sum(obj_soja.values())
total_conc_obj = sum(obj_conc.values())
print(f"\n{'Total S2':8s} {total_soja_pred:10.0f} {total_soja_obj:8.0f} {total_soja_pred/total_soja_obj*100:5.0f}%   {total_conc_pred:10.0f} {total_conc_obj:8.0f} {total_conc_pred/total_conc_obj*100:5.0f}%")

# ===== 6. SAUVEGARDER EXCEL =====
print("\n" + "=" * 80)
print("6. SAUVEGARDE DU FICHIER EXCEL")
print("=" * 80)

OUT = '/home/z/my-project/download/predictions_prophet.xlsx'
wb_out = openpyxl.Workbook()
wb_out.remove(wb_out.active)

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
HEAD_FILL = PatternFill('solid', fgColor='1F4E78')
HEAD_FONT = Font(bold=True, color='FFFFFF', size=11)
THIN = Side(border_style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Feuille 1: Prédictions TOURTEAUX
ws1 = wb_out.create_sheet('TOURTEAUX — Prédictions')
ws1['A1'] = 'PRÉDICTIONS PROPHET — TOURTEAUX (S2 2026)'
ws1['A1'].font = Font(bold=True, size=14, color='1F4E78')
ws1['A2'] = f'Modèle: Prophet (weekly + yearly seasonality, holidays: hausse prix 23/07, regressor: prix_soja). Données: Jan 2025 → Août 2026.'
ws1['A2'].font = Font(italic=True, size=9, color='595959')

headers1 = ['Date', 'Jour', 'Prédiction (t)', 'Borne inf (IC80%)', 'Borne sup (IC80%)', 'Objectif mois (t)', '% obj mois']
for i, h in enumerate(headers1, 1):
    cell = ws1.cell(row=4, column=i, value=h)
    cell.fill = HEAD_FILL; cell.font = HEAD_FONT; cell.border = BORDER

# Fill only Sep-Déc 2026
future_only = forecast_soja[(forecast_soja['ds'] >= pd.Timestamp('2026-09-01')) & (forecast_soja['ds'] <= pd.Timestamp('2026-12-31'))]
row = 5
for _, fr in future_only.iterrows():
    date = fr['ds']
    mois = date.month
    import calendar as cal_mod
    days_lun_sam = sum(1 for d in range(1, cal_mod.monthrange(2026, mois)[1]+1) if datetime.date(2026, mois, d).weekday() < 6)
    obj = obj_soja.get(str(mois), 0) / days_lun_sam if days_lun_sam > 0 else 0  # Daily objective (lun-sam only)
    # Sundays: 0 objective (BELGOCAM doesn't sell on Sundays)
    if date.weekday() == 6:
        obj = 0
        pct = 0
    else:
        pred = fr['yhat']
        pct = pred / obj * 100 if obj > 0 else 0
    pred = fr['yhat']
    lower = fr['yhat_lower']
    upper = fr['yhat_upper']
    
    ws1.cell(row=row, column=1, value=date.strftime('%d/%m/%Y')).border = BORDER
    ws1.cell(row=row, column=2, value=date.strftime('%a')).border = BORDER
    ws1.cell(row=row, column=3, value=round(pred, 1)).border = BORDER
    ws1.cell(row=row, column=4, value=round(lower, 1)).border = BORDER
    ws1.cell(row=row, column=5, value=round(upper, 1)).border = BORDER
    ws1.cell(row=row, column=6, value=round(obj, 1)).border = BORDER
    ws1.cell(row=row, column=7, value=round(pct, 0)).border = BORDER
    row += 1

# Monthly summary at bottom
row += 2
ws1.cell(row=row, column=1, value='SYNTHÈSE MENSUELLE').font = Font(bold=True, size=11, color='1F4E78')
row += 1
for i, h in enumerate(['Mois', 'Prédiction (t)', 'Borne inf', 'Borne sup', 'Objectif (t)', '% atteinte'], 1):
    cell = ws1.cell(row=row, column=i, value=h)
    cell.fill = HEAD_FILL; cell.font = HEAD_FONT; cell.border = BORDER
row += 1
for mois_num in [9, 10, 11, 12]:
    mois_str = f'2026-{mois_num:02d}'
    rows_m = forecast_soja[forecast_soja['ds'].dt.to_period('M').astype(str) == mois_str]
    pred = rows_m['yhat'].sum()
    lower = rows_m['yhat_lower'].sum()
    upper = rows_m['yhat_upper'].sum()
    obj = obj_soja.get(str(mois_num), 0)
    pct = pred / obj * 100 if obj > 0 else 0
    month_name = ['Sept', 'Oct', 'Nov', 'Déc'][mois_num - 9]
    ws1.cell(row=row, column=1, value=month_name).border = BORDER
    ws1.cell(row=row, column=2, value=round(pred, 0)).border = BORDER
    ws1.cell(row=row, column=3, value=round(lower, 0)).border = BORDER
    ws1.cell(row=row, column=4, value=round(upper, 0)).border = BORDER
    ws1.cell(row=row, column=5, value=round(obj, 0)).border = BORDER
    ws1.cell(row=row, column=6, value=round(pct, 0)).border = BORDER
    row += 1

for col, width in [('A', 14), ('B', 8), ('C', 14), ('D', 14), ('E', 14), ('F', 14), ('G', 10)]:
    ws1.column_dimensions[col].width = width

# Feuille 2: Prédictions CONCENTRÉS
ws2 = wb_out.create_sheet('CONCENTRÉS — Prédictions')
ws2['A1'] = 'PRÉDICTIONS PROPHET — CONCENTRÉS (S2 2026)'
ws2['A1'].font = Font(bold=True, size=14, color='1F4E78')
ws2['A2'] = f'Modèle: Prophet (weekly + yearly seasonality, holidays: hausse prix 23/07, regressor: prix_soja). Données: Jan 2025 → Août 2026.'
ws2['A2'].font = Font(italic=True, size=9, color='595959')

headers2 = ['Date', 'Jour', 'Prédiction (t)', 'Borne inf (IC80%)', 'Borne sup (IC80%)', 'Objectif mois (t)', '% obj mois']
for i, h in enumerate(headers2, 1):
    cell = ws2.cell(row=4, column=i, value=h)
    cell.fill = HEAD_FILL; cell.font = HEAD_FONT; cell.border = BORDER

future_conc_only = forecast_conc[(forecast_conc['ds'] >= pd.Timestamp('2026-09-01')) & (forecast_conc['ds'] <= pd.Timestamp('2026-12-31'))]
row = 5
for _, fr in future_conc_only.iterrows():
    date = fr['ds']
    mois = date.month
    days_lun_sam_c = sum(1 for d in range(1, cal_mod.monthrange(2026, mois)[1]+1) if datetime.date(2026, mois, d).weekday() < 6)
    obj = obj_conc.get(str(mois), 0) / days_lun_sam_c if days_lun_sam_c > 0 else 0  # Daily objective (lun-sam only)
    # Sundays: 0 objective (BELGOCAM doesn't sell on Sundays)
    if date.weekday() == 6:
        obj = 0
        pct = 0
    else:
        pct = fr['yhat'] / obj * 100 if obj > 0 else 0
    pred = fr['yhat']
    lower = fr['yhat_lower']
    upper = fr['yhat_upper']
    
    ws2.cell(row=row, column=1, value=date.strftime('%d/%m/%Y')).border = BORDER
    ws2.cell(row=row, column=2, value=date.strftime('%a')).border = BORDER
    ws2.cell(row=row, column=3, value=round(pred, 1)).border = BORDER
    ws2.cell(row=row, column=4, value=round(lower, 1)).border = BORDER
    ws2.cell(row=row, column=5, value=round(upper, 1)).border = BORDER
    ws2.cell(row=row, column=6, value=round(obj, 1)).border = BORDER
    ws2.cell(row=row, column=7, value=round(pct, 0)).border = BORDER
    row += 1

row += 2
ws2.cell(row=row, column=1, value='SYNTHÈSE MENSUELLE').font = Font(bold=True, size=11, color='1F4E78')
row += 1
for i, h in enumerate(['Mois', 'Prédiction (t)', 'Borne inf', 'Borne sup', 'Objectif (t)', '% atteinte'], 1):
    cell = ws2.cell(row=row, column=i, value=h)
    cell.fill = HEAD_FILL; cell.font = HEAD_FONT; cell.border = BORDER
row += 1
for mois_num in [9, 10, 11, 12]:
    mois_str = f'2026-{mois_num:02d}'
    rows_m = forecast_conc[forecast_conc['ds'].dt.to_period('M').astype(str) == mois_str]
    pred = rows_m['yhat'].sum()
    lower = rows_m['yhat_lower'].sum()
    upper = rows_m['yhat_upper'].sum()
    obj = obj_conc.get(str(mois_num), 0)
    pct = pred / obj * 100 if obj > 0 else 0
    month_name = ['Sept', 'Oct', 'Nov', 'Déc'][mois_num - 9]
    ws2.cell(row=row, column=1, value=month_name).border = BORDER
    ws2.cell(row=row, column=2, value=round(pred, 0)).border = BORDER
    ws2.cell(row=row, column=3, value=round(lower, 0)).border = BORDER
    ws2.cell(row=row, column=4, value=round(upper, 0)).border = BORDER
    ws2.cell(row=row, column=5, value=round(obj, 0)).border = BORDER
    ws2.cell(row=row, column=6, value=round(pct, 0)).border = BORDER
    row += 1

for col, width in [('A', 14), ('B', 8), ('C', 14), ('D', 14), ('E', 14), ('F', 14), ('G', 10)]:
    ws2.column_dimensions[col].width = width

# Feuille 3: Synthèse
ws3 = wb_out.create_sheet('Synthèse', 0)
ws3['A1'] = 'SYNTHÈSE — PRÉDICTIONS PROPHET vs OBJECTIFS (S2 2026)'
ws3['A1'].font = Font(bold=True, size=14, color='1F4E78')
ws3['A2'] = f'Modèle: Prophet | Données: Jan 2025 → Août 2026 | Holidays: hausse prix 23/07 | Regressor: prix_soja'
ws3['A2'].font = Font(italic=True, size=9, color='595959')

headers3 = ['Mois', 'TOURTEAUX Préd (t)', 'TOURTEAUX Obj (t)', '%', 'CONCENTRES Préd (t)', 'CONCENTRES Obj (t)', '%']
for i, h in enumerate(headers3, 1):
    cell = ws3.cell(row=4, column=i, value=h)
    cell.fill = HEAD_FILL; cell.font = HEAD_FONT; cell.border = BORDER

row = 5
for mois_num in [9, 10, 11, 12]:
    mois_str = f'2026-{mois_num:02d}'
    soja_rows = forecast_soja[forecast_soja['ds'].dt.to_period('M').astype(str) == mois_str]
    conc_rows = forecast_conc[forecast_conc['ds'].dt.to_period('M').astype(str) == mois_str]
    soja_pred = soja_rows['yhat'].sum()
    conc_pred = conc_rows['yhat'].sum()
    soja_obj = obj_soja.get(str(mois_num), 0)
    conc_obj = obj_conc.get(str(mois_num), 0)
    soja_pct = soja_pred / soja_obj * 100 if soja_obj > 0 else 0
    conc_pct = conc_pred / conc_obj * 100 if conc_obj > 0 else 0
    month_name = ['Septembre', 'Octobre', 'Novembre', 'Décembre'][mois_num - 9]
    ws3.cell(row=row, column=1, value=month_name).border = BORDER
    ws3.cell(row=row, column=2, value=round(soja_pred, 0)).border = BORDER
    ws3.cell(row=row, column=3, value=round(soja_obj, 0)).border = BORDER
    ws3.cell(row=row, column=4, value=round(soja_pct, 0)).border = BORDER
    ws3.cell(row=row, column=5, value=round(conc_pred, 0)).border = BORDER
    ws3.cell(row=row, column=6, value=round(conc_obj, 0)).border = BORDER
    ws3.cell(row=row, column=7, value=round(conc_pct, 0)).border = BORDER
    row += 1

# Total
ws3.cell(row=row, column=1, value='TOTAL S2').font = Font(bold=True)
total_pred_soja = sum(forecast_soja[forecast_soja['ds'] >= pd.Timestamp('2026-09-01')]['yhat'])
total_pred_conc = sum(forecast_conc[forecast_conc['ds'] >= pd.Timestamp('2026-09-01')]['yhat'])
total_obj_soja = sum(obj_soja.get(str(m), 0) for m in [9, 10, 11, 12])
total_obj_conc = sum(obj_conc.get(str(m), 0) for m in [9, 10, 11, 12])
ws3.cell(row=row, column=2, value=round(total_pred_soja, 0)).font = Font(bold=True)
ws3.cell(row=row, column=3, value=round(total_obj_soja, 0)).font = Font(bold=True)
ws3.cell(row=row, column=4, value=round(total_pred_soja / total_obj_soja * 100, 0)).font = Font(bold=True)
ws3.cell(row=row, column=5, value=round(total_pred_conc, 0)).font = Font(bold=True)
ws3.cell(row=row, column=6, value=round(total_obj_conc, 0)).font = Font(bold=True)
ws3.cell(row=row, column=7, value=round(total_pred_conc / total_obj_conc * 100, 0)).font = Font(bold=True)

row += 2
ws3.cell(row=row, column=1, value='NOTES:').font = Font(bold=True, size=11, color='1F4E78')
row += 1
notes = [
    '• Modèle Prophet avec saisonnalité hebdomadaire + annuelle + holidays (hausse prix 23/07) + regressor (prix_soja)',
    '• Intervalles de confiance à 80% (IC80)',
    '• Les prédictions supposent que la hausse de prix reste en vigueur (prix_soja = 1)',
    '• Les prédictions ne tiennent pas compte de la rupture de stock soja (date rupture 02/09/2026)',
    '• Si rupture de stock avant réapprovisionnement, les volumes TOURTEAUX réels seront inférieurs aux prédictions',
]
for note in notes:
    ws3.cell(row=row, column=1, value=note).font = Font(size=9)
    row += 1

for col, width in [('A', 14), ('B', 18), ('C', 18), ('D', 8), ('E', 18), ('F', 18), ('G', 8)]:
    ws3.column_dimensions[col].width = width

wb_out.save(OUT)
print(f"\nSaved: {OUT}")
print(f"Sheets: {wb_out.sheetnames}")
