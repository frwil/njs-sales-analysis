"""
Forecast Q4 2026 - Scenario S3 uniquement - VERSION CORRIGÉE.
- Exclut le MAIS
- Inclut MATERIEL_ELEVAGE (alvéoles + matériel) en CA
- Inclut PREMIX avec prix calculés
- Utilise Prophet pour TOURTEAUX, CONCENTRES, INGREDIENTS, ALIMENT_COMPLET
- Utilise extrapolation moyenne pour MATERIEL_ELEVAGE et PREMIX (séries volatiles)
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

# Load v2 dataset
df = pd.read_csv("/home/z/my-project/scripts/dataset_consolide_v2.csv", parse_dates=['date'], low_memory=False)
prix_forecast = json.load(open("/home/z/my-project/scripts/prix_forecast.json"))

# Add PREMIX prices (calculated from data)
PREMIX_PRICES = {'P102N2': 95147, 'P104N2': 91096, 'P109': 34000}
prix_forecast['stable_Q4'].update(PREMIX_PRICES)

print(f"Loaded {len(df)} records")
print(f"Families: {df['family'].unique()}")

# === Agrégation mensuelle ===
df['year_month'] = df['date'].dt.to_period('M')

monthly_fr_vol = df.groupby(['family', 'region', 'year_month'])['tonnes'].sum().reset_index()
monthly_fr_vol['date'] = monthly_fr_vol['year_month'].dt.to_timestamp()

monthly_fr_ca = df.groupby(['family', 'region', 'year_month'])['montant_ttc'].sum().reset_index()
monthly_fr_ca['date'] = monthly_fr_ca['year_month'].dt.to_timestamp()
monthly_fr_ca['ca_m_fcfa'] = monthly_fr_ca['montant_ttc'] / 1e6

# Parts historiques produit × agence
monthly_pa = df.groupby(['ref', 'agence', 'family', 'region']).agg(
    tonnes=('tonnes', 'sum'),
    ca_ttc=('montant_ttc', 'sum'),
    qte=('qte', 'sum'),
).reset_index()
total_fr = monthly_pa.groupby(['family', 'region'])['ca_ttc'].transform('sum')
monthly_pa['share_ca'] = monthly_pa['ca_ttc'] / total_fr
# Also compute share by tonnes for food families
total_fr_t = monthly_pa.groupby(['family', 'region'])['tonnes'].transform('sum')
monthly_pa['share_tonnes'] = monthly_pa['tonnes'] / total_fr_t.replace(0, 1)

# === Liste des combos famille × région ===
combos_fr = df.groupby(['family', 'region']).size().reset_index()
combos_fr.columns = ['family', 'region', 'n_months']
print(f"\n{len(combos_fr)} combinaisons famille × région à modéliser")

# === Fonction Prophet ===
def fit_prophet_fast(history_df, periods=4, freq='MS'):
    if len(history_df) < 4:
        avg = history_df['y'].mean() if len(history_df) > 0 else 0
        future_dates = pd.date_range(start='2026-09-01', periods=periods, freq=freq)
        return pd.DataFrame({'ds': future_dates, 'yhat': [avg] * periods,
                             'yhat_lower': [avg * 0.8] * periods, 'yhat_upper': [avg * 1.2] * periods})
    try:
        m = Prophet(yearly_seasonality=True, weekly_seasonality=False, daily_seasonality=False,
                    seasonality_mode='multiplicative', changepoint_prior_scale=0.05,
                    interval_width=0.8, mcmc_samples=0)
        m.fit(history_df[['ds', 'y']])
        future = m.make_future_dataframe(periods=periods, freq=freq, include_history=False)
        forecast = m.predict(future)
        return forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']]
    except Exception as e:
        avg_q4 = history_df[history_df['ds'].dt.month.isin([9,10,11,12])]['y'].mean()
        if pd.isna(avg_q4) or avg_q4 == 0:
            avg_q4 = history_df['y'].mean()
        future_dates = pd.date_range(start='2026-09-01', periods=periods, freq=freq)
        return pd.DataFrame({'ds': future_dates, 'yhat': [avg_q4] * periods,
                             'yhat_lower': [avg_q4 * 0.8] * periods, 'yhat_upper': [avg_q4 * 1.2] * periods})


def extrapolate_mean(history_df, periods=4, freq='MS', value_col='y'):
    """Simple extrapolation using historical monthly average."""
    avg = history_df[value_col].mean() if len(history_df) > 0 else 0
    # Apply Q4 seasonal factor if available
    q4_history = history_df[history_df['ds'].dt.month.isin([9, 10, 11, 12])]
    if len(q4_history) >= 4:
        q4_avg = q4_history[value_col].mean()
        # Use Q4 average for forecast
        avg = q4_avg
    future_dates = pd.date_range(start='2026-09-01', periods=periods, freq=freq)
    return pd.DataFrame({'ds': future_dates, 'yhat': [avg] * periods,
                         'yhat_lower': [avg * 0.7] * periods, 'yhat_upper': [avg * 1.3] * periods})


# === Forecast Q4 2026 par famille × région (S3 uniquement) ===
print("\n=== Forecast Q4 2026 (Scénario S3 - Réappro 100%) ===")
forecasts_fr = []

PROPHET_FAMILIES = ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'ALIMENT_COMPLET']
EXTRAPOL_FAMILIES = ['MATERIEL_ELEVAGE', 'PREMIX']  # Volatile, use mean extrapolation

for idx, row in combos_fr.iterrows():
    family = row['family']
    region = row['region']
    
    if family in EXTRAPOL_FAMILIES:
        # Extrapolation CA for material/premix
        history = monthly_fr_ca[(monthly_fr_ca['family'] == family) & (monthly_fr_ca['region'] == region)].copy()
        history = history[['date', 'ca_m_fcfa']].rename(columns={'date': 'ds', 'ca_m_fcfa': 'y'})
        history = history.sort_values('ds')
        history['y'] = history['y'].clip(lower=0)
        forecast = extrapolate_mean(history, periods=4, freq='MS', value_col='y')
        is_ca_only = True
        val_sum = forecast['yhat'].sum()
        unit = 'M FCFA'
    else:
        # Prophet for tonnes
        history = monthly_fr_vol[(monthly_fr_vol['family'] == family) & (monthly_fr_vol['region'] == region)].copy()
        history = history[['date', 'tonnes']].rename(columns={'date': 'ds', 'tonnes': 'y'})
        history = history.sort_values('ds')
        history['y'] = history['y'].clip(lower=0)
        forecast = fit_prophet_fast(history, periods=4, freq='MS')
        is_ca_only = False
        val_sum = forecast['yhat'].sum()
        unit = 't'
    
    forecast['yhat'] = forecast['yhat'].clip(lower=0)
    forecast['family'] = family
    forecast['region'] = region
    forecast['is_ca_only'] = is_ca_only
    
    forecasts_fr.append(forecast)
    print(f"  ✓ {family} × {region}: forecast {val_sum:.1f} {unit}")

forecasts_fr_df = pd.concat(forecasts_fr, ignore_index=True)

# === Désagrégation par produit × agence (S3) ===
print("\n=== Désagrégation par produit × agence (S3) ===")
all_forecasts = []

for _, fcst_row in forecasts_fr_df.iterrows():
    family = fcst_row['family']
    region = fcst_row['region']
    date = fcst_row['ds']
    month = date.month
    is_ca_only = fcst_row['is_ca_only']
    
    if is_ca_only:
        ca_total_m_fcfa = fcst_row['yhat']
        tonnes_total = 0
    else:
        tonnes_total = fcst_row['yhat']
        # Apply S3 scenario for TOURTEAUX
        if family == 'TOURTEAUX' and month == 9:
            tonnes_total *= 0.7
        # Calculate CA via prices
        pa_subset = monthly_pa[(monthly_pa['family'] == family) & (monthly_pa['region'] == region)]
        ca_total_m_fcfa = 0
        for _, pa_row in pa_subset.iterrows():
            ref = pa_row['ref']
            share_t = pa_row['share_tonnes']
            tonnes_combo = tonnes_total * share_t
            sacs_50 = tonnes_combo * 1000 / 50
            prix = prix_forecast['stable_Q4'].get(ref, 0)
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
            prix = prix_forecast['stable_Q4'].get(ref, 0)
        
        all_forecasts.append({
            'scenario': 'S3_reappro_100',
            'ref': ref,
            'family': family,
            'agence': agence,
            'region': region,
            'date': date,
            'month': month,
            'year': 2026,
            'tonnes': round(tonnes_combo, 2),
            'sacs_50': round(sacs_50, 1),
            'prix_ttc_sac': prix,
            'ca_m_fcfa': round(ca_m_fcfa_combo, 2),
        })

fcst_df = pd.DataFrame(all_forecasts)
fcst_df = fcst_df[(fcst_df['year'] == 2026) & (fcst_df['month'].isin([9, 10, 11, 12]))]
print(f"\n{len(fcst_df)} forecasts détaillés générés (S3 uniquement)")

# Save
output_path = "/home/z/my-project/scripts/forecast_q4_2026_S3.csv"
fcst_df.to_csv(output_path, index=False)
print(f"Saved: {output_path}")

# === Synthèse ===
print("\n=== SYNTHÈSE Q4 2026 (S3 - Réappro 100%) ===")
print("\n--- Par famille ---")
synth_fam = fcst_df.groupby('family').agg(
    tonnes=('tonnes', 'sum'),
    ca_m_fcfa=('ca_m_fcfa', 'sum'),
).round({'tonnes': 0, 'ca_m_fcfa': 1})
print(synth_fam)

print("\n--- Par région ---")
synth_reg = fcst_df.groupby('region').agg(
    tonnes=('tonnes', 'sum'),
    ca_m_fcfa=('ca_m_fcfa', 'sum'),
).round({'tonnes': 0, 'ca_m_fcfa': 1})
print(synth_reg)

print("\n--- Par mois ---")
synth_month = fcst_df.groupby('month').agg(
    tonnes=('tonnes', 'sum'),
    ca_m_fcfa=('ca_m_fcfa', 'sum'),
).round({'tonnes': 0, 'ca_m_fcfa': 1})
print(synth_month)

print("\n--- TOTAL Q4 2026 ---")
print(f"Tonnes: {fcst_df['tonnes'].sum():.0f}")
print(f"CA: {fcst_df['ca_m_fcfa'].sum():.1f} M FCFA")
