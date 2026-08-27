"""
Phase Construct - Modélisation Prophet OPTIMISÉE.
Stratégie: 1 modèle Prophet par famille × région (12 modèles au lieu de 130+),
puis désagrégation par produit × agence proportionnellement à la part historique.

4 scénarios soja:
  S1 - Rupture totale: ventes soja = 0 après le 16/09/2026
  S2 - Réappro 50%: 40 000 sacs au 01/10/2026
  S3 - Réappro 100%: 80 000 sacs au 15/09/2026
  S4 - Baisse prix: Réappro 100% + baisse prix soja -10% (si stock se stabilise)
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

# Load data
df = pd.read_csv("/home/z/my-project/scripts/dataset_consolide.csv", parse_dates=['date'])
prix_forecast = json.load(open("/home/z/my-project/scripts/prix_forecast.json"))

print(f"Loaded {len(df)} records")
print(f"Date range: {df['date'].min().date()} → {df['date'].max().date()}")

# === Agrégation mensuelle par famille × région ===
print("\nAgrégation mensuelle par famille × région...")
df['year_month'] = df['date'].dt.to_period('M')
monthly_fr = df.groupby(['family', 'region', 'year_month'])['tonnes'].sum().reset_index()
monthly_fr['date'] = monthly_fr['year_month'].dt.to_timestamp()
print(f"  {len(monthly_fr)} combinaisons famille × région × mois")

# Also aggregate at produit × agence level for proportion calculation
monthly_pa = df.groupby(['ref', 'agence', 'family', 'region'])['tonnes'].sum().reset_index()
print(f"  {len(monthly_pa)} combinaisons produit × agence (historique total)")

# Compute shares: for each family × region, what is each produit × agence's share?
total_fr = monthly_pa.groupby(['family', 'region'])['tonnes'].transform('sum')
monthly_pa['share'] = monthly_pa['tonnes'] / total_fr
print(f"  Shares computed (sum should be 1.0 per family × region)")
print(monthly_pa.groupby(['family', 'region'])['share'].sum().head(10))

# === Liste des combos famille × région à modéliser ===
combos_fr = monthly_fr.groupby(['family', 'region']).size().reset_index()
combos_fr.columns = ['family', 'region', 'n_months']
print(f"\n{len(combos_fr)} combinaisons famille × région à modéliser avec Prophet")

# === Fonction Prophet simplifiée ===
def fit_prophet_fast(history_df, periods=4, freq='MS'):
    """Fit Prophet fast configuration."""
    if len(history_df) < 4:
        avg = history_df['y'].mean() if len(history_df) > 0 else 0
        future_dates = pd.date_range(start='2026-09-01', periods=periods, freq=freq)
        return pd.DataFrame({'ds': future_dates, 'yhat': [avg] * periods, 
                             'yhat_lower': [avg * 0.8] * periods, 'yhat_upper': [avg * 1.2] * periods})
    
    try:
        m = Prophet(
            yearly_seasonality=True,
            weekly_seasonality=False,
            daily_seasonality=False,
            seasonality_mode='multiplicative',
            changepoint_prior_scale=0.05,
            interval_width=0.8,
            # Disable MCMC for speed
            mcmc_samples=0,
            stan_backend='CMDSTANPY',
        )
        m.fit(history_df)
        future = m.make_future_dataframe(periods=periods, freq=freq, include_history=False)
        forecast = m.predict(future)
        return forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']]
    except Exception as e:
        print(f"  Prophet error on {history_df['family'].iloc[0] if 'family' in history_df else '?'}: {e}")
        # Fallback: simple seasonal naive forecast
        # Use Q4 2025 as proxy if available
        avg_q4 = history_df[history_df['ds'].dt.month.isin([9, 10, 11, 12])]['y'].mean()
        if pd.isna(avg_q4) or avg_q4 == 0:
            avg_q4 = history_df['y'].mean()
        future_dates = pd.date_range(start='2026-09-01', periods=periods, freq=freq)
        return pd.DataFrame({'ds': future_dates, 'yhat': [avg_q4] * periods,
                             'yhat_lower': [avg_q4 * 0.8] * periods, 'yhat_upper': [avg_q4 * 1.2] * periods})


# === Forecast Q4 2026 par famille × région ===
print("\n=== Forecast Q4 2026 par famille × région ===")
forecasts_fr = []

for idx, row in combos_fr.iterrows():
    family = row['family']
    region = row['region']
    
    # Get history for this family × region
    history = monthly_fr[(monthly_fr['family'] == family) & (monthly_fr['region'] == region)].copy()
    history = history[['date', 'tonnes']].rename(columns={'date': 'ds', 'tonnes': 'y'})
    history = history.sort_values('ds')
    history['y'] = history['y'].clip(lower=0)
    
    # Forecast
    forecast = fit_prophet_fast(history, periods=4, freq='MS')
    forecast['yhat'] = forecast['yhat'].clip(lower=0)
    forecast['family'] = family
    forecast['region'] = region
    
    forecasts_fr.append(forecast)
    print(f"  ✓ {family} × {region}: forecast {forecast['yhat'].sum():.0f} t Q4 2026")

forecasts_fr_df = pd.concat(forecasts_fr, ignore_index=True)
print(f"\n{len(forecasts_fr_df)} forecasts famille × région générés")

# === Désagrégation par produit × agence ===
print("\n=== Désagrégation par produit × agence ===")
all_forecasts = []

# Pour chaque forecast famille × région, désagréger
for _, fcst_row in forecasts_fr_df.iterrows():
    family = fcst_row['family']
    region = fcst_row['region']
    date = fcst_row['ds']
    month = date.month
    tonnes_total = fcst_row['yhat']
    tonnes_lower = fcst_row['yhat_lower']
    tonnes_upper = fcst_row['yhat_upper']
    
    # Get all produit × agence in this family × region
    pa_subset = monthly_pa[(monthly_pa['family'] == family) & (monthly_pa['region'] == region)]
    
    for _, pa_row in pa_subset.iterrows():
        ref = pa_row['ref']
        agence = pa_row['agence']
        share = pa_row['share']
        
        # Compute volume for this combo
        tonnes_combo = tonnes_total * share
        tonnes_lower_combo = tonnes_lower * share
        tonnes_upper_combo = tonnes_upper * share
        
        # Apply scenarios
        for scenario_name in ['S1_rupture', 'S2_reappro_50', 'S3_reappro_100', 'S4_baisse_prix']:
            tonnes_scn = tonnes_combo
            lower_scn = tonnes_lower_combo
            upper_scn = tonnes_upper_combo
            
            # Apply scenario adjustments
            if scenario_name == 'S1_rupture' and family == 'TOURTEAUX':
                if month >= 10:  # Oct, Nov, Dec = 0
                    tonnes_scn = 0
                    lower_scn = 0
                    upper_scn = 0
                elif month == 9:  # Sept = 50%
                    tonnes_scn *= 0.5
                    lower_scn *= 0.5
                    upper_scn *= 0.5
            
            elif scenario_name == 'S2_reappro_50' and family == 'TOURTEAUX':
                if month == 9:
                    tonnes_scn *= 0.5
                    lower_scn *= 0.5
                    upper_scn *= 0.5
                elif month in [10, 11]:
                    tonnes_scn *= 0.7
                    lower_scn *= 0.7
                    upper_scn *= 0.7
            
            elif scenario_name == 'S3_reappro_100' and family == 'TOURTEAUX':
                if month == 9:
                    tonnes_scn *= 0.7
                    lower_scn *= 0.7
                    upper_scn *= 0.7
            
            elif scenario_name == 'S4_baisse_prix' and family == 'TOURTEAUX':
                if month == 9:
                    tonnes_scn *= 0.7 * 1.05
                else:
                    tonnes_scn *= 1.10
                lower_scn *= 1.10
                upper_scn *= 1.10
            
            # Apply price
            if scenario_name == 'S4_baisse_prix':
                prix = prix_forecast['baisse_prix_soja_-10%'].get(ref, prix_forecast['stable_Q4'].get(ref, 0))
            else:
                prix = prix_forecast['stable_Q4'].get(ref, 0)
            
            sacs_50 = tonnes_scn * 1000 / 50
            ca_ttc = sacs_50 * prix
            ca_m_fcfa = ca_ttc / 1e6
            
            all_forecasts.append({
                'scenario': scenario_name,
                'ref': ref,
                'family': family,
                'agence': agence,
                'region': region,
                'date': date,
                'month': month,
                'year': 2026,
                'tonnes': round(tonnes_scn, 2),
                'tonnes_lower': round(lower_scn, 2),
                'tonnes_upper': round(upper_scn, 2),
                'sacs_50': round(sacs_50, 1),
                'prix_ttc_sac': prix,
                'ca_ttc_fcfa': round(ca_ttc, 0),
                'ca_m_fcfa': round(ca_m_fcfa, 2),
            })

fcst_df = pd.DataFrame(all_forecasts)
print(f"\n{len(fcst_df)} forecasts détaillés générés")

# Save
output_path = "/home/z/my-project/scripts/forecast_q4_2026.csv"
fcst_df.to_csv(output_path, index=False)
print(f"Saved: {output_path}")
print(f"File size: {os.path.getsize(output_path) / 1024:.0f} KB")

# === Synthèse par scénario ===
print("\n=== SYNTHÈSE PAR SCÉNARIO (Q4 2026 = sept + oct + nov + déc) ===")
print("\n--- Tonnes par scénario et famille ---")
synth_t = fcst_df.groupby(['scenario', 'family'])['tonnes'].sum().unstack(fill_value=0).round(0)
print(synth_t)

print("\n--- CA (M FCFA) par scénario et famille ---")
synth_ca = fcst_df.groupby(['scenario', 'family'])['ca_m_fcfa'].sum().unstack(fill_value=0).round(1)
print(synth_ca)

print("\n--- TOTAL Q4 2026 par scénario ---")
total_synth = fcst_df.groupby('scenario').agg(
    tonnes=('tonnes', 'sum'),
    ca_m_fcfa=('ca_m_fcfa', 'sum'),
).round({'tonnes': 0, 'ca_m_fcfa': 1})
print(total_synth)

# Save synthèse
total_synth.to_csv("/home/z/my-project/scripts/forecast_q4_synthesis.csv")
synth_t.to_csv("/home/z/my-project/scripts/forecast_q4_by_family_tonnes.csv")
synth_ca.to_csv("/home/z/my-project/scripts/forecast_q4_by_family_ca.csv")

# Synthèse par région
print("\n--- Tonnes par scénario et région ---")
synth_reg = fcst_df.groupby(['scenario', 'region'])['tonnes'].sum().unstack(fill_value=0).round(0)
print(synth_reg)
synth_reg.to_csv("/home/z/my-project/scripts/forecast_q4_by_region.csv")

# Synthèse par mois
print("\n--- Tonnes par scénario et mois ---")
synth_month = fcst_df.groupby(['scenario', 'month'])['tonnes'].sum().unstack(fill_value=0).round(0)
print(synth_month)

print("\n=== FORECAST COMPLETE ===")
