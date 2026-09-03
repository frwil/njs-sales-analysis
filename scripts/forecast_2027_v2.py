"""
Forecast 2027 avec 5,7 ans d'historique (2021-2026).
Utilise le dataset_2021_2026.csv (190 451 records, 68 mois).

Innovations:
- 5,7 ans d'historique (vs 20 mois précédemment) → Prophet yearly seasonality robuste
- Désaisonnalisation effet soja Jul-Août 2026 conservée
- En cours + Validées incluses
- Prix soja 25 000 FCFA/sac
- MAIS exclu (ligne à 0 pour traçabilité)
"""
import pandas as pd
import numpy as np
import json
import os
import warnings
import logging
warnings.filterwarnings('ignore')
logging.getLogger('cmdstanpy').setLevel(logging.WARNING)
logging.getLogger('prophet').setLevel(logging.WARNING)

from prophet import Prophet

# === Load 2021-2026 dataset ===
df = pd.read_csv("/home/z/my-project/scripts/dataset_2021_2026.csv", parse_dates=['date'], low_memory=False)
print(f"Loaded {len(df)} records")
print(f"Date range: {df['date'].min().date()} → {df['date'].max().date()}")
print(f"Months of history: {df['date'].dt.to_period('M').nunique()}")

# === Add En cours + Validées from latest extraction ===
import openpyxl
FILE_AOUT = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (27).xlsx"
wb = openpyxl.load_workbook(FILE_AOUT, read_only=True, data_only=True)
ws = wb['Sheet 1']

SOJA_REFS = {'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25}
CONC_REFS = {'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
             'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50, 'C1022': 5}
ALIMENT_REFS = {'CB100': 25, 'CB200': 25, 'CB101': 5, 'CB201': 5, 'PB100': 25, 'PB200': 25, 'DB100': 25, 'DB200': 25, 'ALAP25': 25}
INGREDIENT_REFS = {'B100': 25, 'E101': 25, 'I105': 25, 'B1001': 1, 'B1003': 5, 'B1004': 25,
                   'E1011': 1, 'E1013': 5, 'E1014': 0.2, 'I1051': 1, 'I1053': 5, 'I1054': 25,
                   'I106': 25, 'I1061': 1, 'I107': 25, 'I1071': 1, 'P105': 25, 'P1051': 1, 'P1053': 5,
                   'F114': 50, 'F1145': 50, 'F1146': 25, 'F1147': 1}
PREMIX_REFS = {'P102N2': 25, 'P104N2': 25, 'P109': 25, 'PX101': 25, 'PX102': 25, 'PX103': 25, 'PX104': 25, 'PX105': 25}

def get_family_v2(ref):
    if ref in SOJA_REFS: return 'TOURTEAUX'
    if ref in CONC_REFS: return 'CONCENTRES'
    if ref in ALIMENT_REFS: return 'ALIMENT_COMPLET'
    if ref in INGREDIENT_REFS: return 'INGREDIENTS'
    if ref in PREMIX_REFS: return 'PREMIX'
    if ref.startswith('MAT'): return 'MATERIEL_ELEVAGE'
    return None

AGENCE_MAP = {
    'AGENCE FAMLA': ('Famla', 'Ouest'), 'AGENCE MESSASSI': ('Messassi', 'Centre'),
    'AGENCE BERTOUA': ('Bertoua', 'Centre'), 'AGENCE NDOBO': ('Ndobo', 'Littoral'),
    'AGENCE DJELENG': ('Djeleng', 'Ouest'), 'AGENCE VILLAGE': ('Village', 'Littoral'),
    'AGENCE PK11': ('Pk11', 'Littoral'), 'AGENCE NGAOUNDERE': ('Ngaoundere', 'Centre'),
    'AGENCE AHALA': ('Ahala', 'Centre'), 'AGENCE NKONGSAMBA': ('Nkongsamba', 'Littoral'),
    'AGENCE NKOLBISSON': ('Nkolbisson', 'Centre'), 'AGENCE NKOABANG': ('Nkoabang', 'Centre'),
    'AGENCE DE BAMENDA - DEPOT MBOUDA': ('Mbouda', 'Ouest'), 'AGENCE BUEA': ('Buea', 'Littoral'),
}

extra_records = []
for r in ws.iter_rows(min_row=3, values_only=True):
    if not r or not r[0] or r[0] == 'Total': continue
    ref = str(r[0])
    family = get_family_v2(ref)
    if family is None: continue
    etat = str(r[13]).strip() if r[13] else ''
    if etat not in ('En cours', 'Validée'): continue
    date_str = str(r[6])[:10] if r[6] else ''
    if '/08/2026' not in date_str: continue
    agence_raw = r[15] if r[15] else ''
    if agence_raw not in AGENCE_MAP: continue
    agence, region = AGENCE_MAP[agence_raw]
    qte = r[2] or 0
    weight = {**SOJA_REFS, **CONC_REFS, **ALIMENT_REFS, **INGREDIENT_REFS, **PREMIX_REFS}.get(ref, 1)
    kg = qte * weight
    try:
        date = pd.to_datetime(date_str, format='%d/%m/%Y')
    except:
        continue
    extra_records.append({
        'date': date, 'year': date.year, 'month': date.month,
        'ref': ref, 'family': family, 'description': r[1],
        'agence': agence, 'region': region,
        'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': kg / 1000,
        'sacs_50': kg / 50 if family != 'MATERIEL_ELEVAGE' else 0,
        'montant_ttc': r[9] or 0, 'montant_ht': 0,
        'source': 'En_cours_Validee'
    })

print(f"\nEn cours + Validées: {len(extra_records)} records")
df_extra = pd.DataFrame(extra_records)
df_all = pd.concat([df, df_extra], ignore_index=True)
print(f"Dataset total: {len(df_all)} records, {df_all['date'].min().date()} → {df_all['date'].max().date()}")

# === Désaisonnalisation effet soja Jul-Août 2026 ===
print("\n=== DÉSAISONNALISATION EFFET SOJA ===")
df_all['year_month'] = df_all['date'].dt.to_period('M')
soja_s1 = df_all[(df_all['family'] == 'TOURTEAUX') & (df_all['date'] < '2026-07-01')]
soja_s1_monthly = soja_s1.groupby('year_month')['tonnes'].sum()
soja_s1_avg = soja_s1_monthly.mean()
print(f"  S1 average monthly soja: {soja_s1_avg:.0f} t")

soja_jul_aug = df_all[(df_all['family'] == 'TOURTEAUX') & 
                      (df_all['date'] >= '2026-07-01') & 
                      (df_all['date'] <= '2026-08-31')]
soja_max = soja_jul_aug.groupby('year_month')['tonnes'].sum().max()
cap_factor = soja_s1_avg / soja_max if soja_max > 0 else 1.0
print(f"  Cap factor: {cap_factor:.2f}")

mask_cap = (df_all['family'] == 'TOURTEAUX') & (df_all['date'] >= '2026-07-01') & (df_all['date'] <= '2026-08-31')
df_all.loc[mask_cap, 'tonnes'] = df_all.loc[mask_cap, 'tonnes'] * cap_factor
df_all.loc[mask_cap, 'sacs_50'] = df_all.loc[mask_cap, 'sacs_50'] * cap_factor
df_all.loc[mask_cap, 'kg'] = df_all.loc[mask_cap, 'kg'] * cap_factor

# === Prix 2027 ===
prix_forecast = json.load(open("/home/z/my-project/scripts/prix_forecast.json"))
prix_2027 = prix_forecast['stable_Q4'].copy()
for ref in SOJA_REFS:
    if ref in prix_2027:
        prix_2027[ref] = 25000 if ref == 'T102' else int(prix_2027[ref] * 25000 / 20600)
PREMIX_PRICES = {'P102N2': 95147, 'P104N2': 91096, 'P109': 34000}
prix_2027.update(PREMIX_PRICES)

# === Agrégation mensuelle par famille × région ===
print("\n=== AGRÉGATION MENSUELLE ===")
monthly_fr_vol = df_all.groupby(['family', 'region', 'year_month'])['tonnes'].sum().reset_index()
monthly_fr_vol['date'] = monthly_fr_vol['year_month'].dt.to_timestamp()

monthly_fr_ca = df_all.groupby(['family', 'region', 'year_month'])['montant_ht'].sum().reset_index()
monthly_fr_ca['date'] = monthly_fr_ca['year_month'].dt.to_timestamp()
# For 2025-2026, use montant_ttc if available; for 2021-2024, montant_ht
# Simplest: use montant_ht for all (consistent)
monthly_fr_ca['ca_m_fcfa'] = monthly_fr_ca['montant_ht'] / 1e6

# Parts historiques produit × agence
monthly_pa = df_all.groupby(['ref', 'agence', 'family', 'region']).agg(
    tonnes=('tonnes', 'sum'),
    ca_ttc=('montant_ht', 'sum'),  # Using HT as proxy
).reset_index()
total_fr_ca = monthly_pa.groupby(['family', 'region'])['ca_ttc'].transform('sum')
monthly_pa['share_ca'] = monthly_pa['ca_ttc'] / total_fr_ca
total_fr_t = monthly_pa.groupby(['family', 'region'])['tonnes'].transform('sum')
monthly_pa['share_tonnes'] = monthly_pa['tonnes'] / total_fr_t.replace(0, 1)

combos_fr = df_all.groupby(['family', 'region']).size().reset_index()
combos_fr.columns = ['family', 'region', 'n_months']
print(f"  {len(combos_fr)} combinaisons famille × région")

# === Prophet ===
def fit_prophet_fast(history_df, periods=12, freq='MS'):
    future_dates = pd.date_range(start='2027-01-01', periods=periods, freq=freq)
    if len(history_df) < 4:
        avg = history_df['y'].mean() if len(history_df) > 0 else 0
        return pd.DataFrame({'ds': future_dates, 'yhat': [avg] * periods,
                             'yhat_lower': [avg * 0.8] * periods, 'yhat_upper': [avg * 1.2] * periods})
    try:
        m = Prophet(yearly_seasonality=True, weekly_seasonality=False, daily_seasonality=False,
                    seasonality_mode='multiplicative', changepoint_prior_scale=0.05,
                    interval_width=0.8, mcmc_samples=0)
        m.fit(history_df[['ds', 'y']])
        future = pd.DataFrame({'ds': future_dates})
        forecast = m.predict(future)
        return forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']]
    except Exception as e:
        print(f"    Prophet error: {e}")
        avg = history_df['y'].mean()
        return pd.DataFrame({'ds': future_dates, 'yhat': [avg] * periods,
                             'yhat_lower': [avg * 0.8] * periods, 'yhat_upper': [avg * 1.2] * periods})

def extrapolate_mean(history_df, periods=12, freq='MS', value_col='y'):
    future_dates = pd.date_range(start='2027-01-01', periods=periods, freq=freq)
    avg = history_df[value_col].mean() if len(history_df) > 0 else 0
    return pd.DataFrame({'ds': future_dates, 'yhat': [avg] * periods,
                         'yhat_lower': [avg * 0.7] * periods, 'yhat_upper': [avg * 1.3] * periods})

# === Forecast 2027 ===
print("\n=== FORECAST 2027 (12 mois, S3, 5,7 ans d'historique) ===")
forecasts_fr = []
PROPHET_FAMILIES = ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'ALIMENT_COMPLET']
EXTRAPOL_FAMILIES = ['MATERIEL_ELEVAGE', 'PREMIX']

for idx, row in combos_fr.iterrows():
    family = row['family']
    region = row['region']
    
    if family in EXTRAPOL_FAMILIES:
        history = monthly_fr_ca[(monthly_fr_ca['family'] == family) & (monthly_fr_ca['region'] == region)].copy()
        history = history[['date', 'ca_m_fcfa']].rename(columns={'date': 'ds', 'ca_m_fcfa': 'y'})
        history = history.sort_values('ds')
        history['y'] = history['y'].clip(lower=0)
        forecast = extrapolate_mean(history, periods=12, freq='MS', value_col='y')
        is_ca_only = True
    else:
        history = monthly_fr_vol[(monthly_fr_vol['family'] == family) & (monthly_fr_vol['region'] == region)].copy()
        history = history[['date', 'tonnes']].rename(columns={'date': 'ds', 'tonnes': 'y'})
        history = history.sort_values('ds')
        history['y'] = history['y'].clip(lower=0)
        forecast = fit_prophet_fast(history, periods=12, freq='MS')
        is_ca_only = False
    
    forecast['yhat'] = forecast['yhat'].clip(lower=0)
    forecast['family'] = family
    forecast['region'] = region
    forecast['is_ca_only'] = is_ca_only
    forecasts_fr.append(forecast)
    val_sum = forecast['yhat'].sum()
    unit = 'M FCFA' if is_ca_only else 't'
    print(f"  ✓ {family} × {region}: forecast {val_sum:.1f} {unit}")

forecasts_fr_df = pd.concat(forecasts_fr, ignore_index=True)

# === Désagrégation ===
print("\n=== DÉSAGRÉGATION ===")
all_forecasts = []

for _, fcst_row in forecasts_fr_df.iterrows():
    family = fcst_row['family']
    region = fcst_row['region']
    date = fcst_row['ds']
    month = date.month
    year = date.year
    is_ca_only = fcst_row['is_ca_only']
    
    if is_ca_only:
        ca_total_m_fcfa = fcst_row['yhat']
        tonnes_total = 0
    else:
        tonnes_total = fcst_row['yhat']
        pa_subset = monthly_pa[(monthly_pa['family'] == family) & (monthly_pa['region'] == region)]
        ca_total_m_fcfa = 0
        for _, pa_row in pa_subset.iterrows():
            ref = pa_row['ref']
            share_t = pa_row['share_tonnes']
            tonnes_combo = tonnes_total * share_t
            sacs_50 = tonnes_combo * 1000 / 50
            prix = prix_2027.get(ref, 0)
            ca_total_m_fcfa += sacs_50 * prix / 1e6
    
    pa_subset = monthly_pa[(monthly_pa['family'] == family) & (monthly_pa['region'] == region)]
    
    for _, pa_row in pa_subset.iterrows():
        ref = pa_row['ref']
        agence = pa_row['agence']
        share_ca = pa_row['share_ca']
        ca_m_fcfa_combo = ca_total_m_fcfa * share_ca
        
        if is_ca_only:
            tonnes_combo = 0
            sacs_50 = 0
            prix = 0
        else:
            tonnes_combo = tonnes_total * pa_row['share_tonnes']
            sacs_50 = tonnes_combo * 1000 / 50
            prix = prix_2027.get(ref, 0)
        
        all_forecasts.append({
            'scenario': 'S3_reappro_100',
            'ref': ref, 'family': family, 'agence': agence, 'region': region,
            'date': date, 'month': month, 'year': year,
            'tonnes': round(tonnes_combo, 2),
            'sacs_50': round(sacs_50, 1),
            'prix_ttc_sac': prix,
            'ca_m_fcfa': round(ca_m_fcfa_combo, 2),
        })

fcst_df = pd.DataFrame(all_forecasts)
fcst_df = fcst_df[fcst_df['year'] == 2027]
print(f"\n{len(fcst_df)} forecasts générés")

# Save
fcst_df.to_csv("/home/z/my-project/scripts/forecast_2027_S3.csv", index=False)
print(f"Saved: forecast_2027_S3.csv")

# === Synthèse ===
print("\n=== SYNTHÈSE FORECAST 2027 (5,7 ans d'historique) ===")
print("\n--- Par famille ---")
synth_fam = fcst_df.groupby('family').agg(tonnes=('tonnes','sum'), ca_m_fcfa=('ca_m_fcfa','sum')).round({'tonnes':0,'ca_m_fcfa':1})
total_ca = fcst_df['ca_m_fcfa'].sum()
synth_fam['part_ca_pct'] = (synth_fam['ca_m_fcfa'] / total_ca * 100).round(1)
print(synth_fam)

print(f"\n--- TOTAL 2027 ---")
print(f"Volume: {fcst_df['tonnes'].sum():.0f} t")
print(f"CA: {fcst_df['ca_m_fcfa'].sum():.1f} M FCFA")

print("\n--- Par trimestre ---")
fcst_df['quarter'] = ((fcst_df['month'] - 1) // 3) + 1
q = fcst_df.groupby('quarter').agg(tonnes=('tonnes','sum'), ca_m_fcfa=('ca_m_fcfa','sum')).round({'tonnes':0,'ca_m_fcfa':1})
q.index = ['Q1','Q2','Q3','Q4']
print(q)

print("\n--- Par mois ---")
m = fcst_df.groupby('month').agg(tonnes=('tonnes','sum'), ca_m_fcfa=('ca_m_fcfa','sum')).round({'tonnes':0,'ca_m_fcfa':1})
month_names = ['Jan','Fév','Mar','Avr','Mai','Juin','Juil','Août','Sep','Oct','Nov','Déc']
m.index = month_names
print(m)
