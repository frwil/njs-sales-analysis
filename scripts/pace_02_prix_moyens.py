"""
Phase Prepare - Calcul des prix moyens par produit (Option A).
Extrapolation à partir du CA HT/TTC et des quantités des extractions existantes.

Output: prix_moyens.csv (1 row par ref produit × year_month)
        prix_moyens_current.json (prix le plus récent par ref pour le forecast Q4)
"""
import pandas as pd
import numpy as np
import json

df = pd.read_csv("/home/z/my-project/scripts/dataset_consolide.csv", parse_dates=['date'])
print(f"Loaded {len(df)} records")

# Compute price per sac_50 (équivalent sac de 50 kg) for each ref × year_month
# Filter out records with 0 sacs or 0 montant
df_valid = df[(df['sacs_50'] > 0) & (df['montant_ttc'] > 0)].copy()
df_valid['prix_sac_50_ttc'] = df_valid['montant_ttc'] / df_valid['sacs_50']
df_valid['prix_sac_50_ht'] = df_valid['montant_ht'] / df_valid['sacs_50']

# Sanity check: filter outliers (prices outside 1st-99th percentile)
def filter_outliers(group):
    if len(group) < 5:
        return group
    q1 = group['prix_sac_50_ttc'].quantile(0.01)
    q99 = group['prix_sac_50_ttc'].quantile(0.99)
    return group[(group['prix_sac_50_ttc'] >= q1) & (group['prix_sac_50_ttc'] <= q99)]

df_clean = df_valid.groupby('ref', group_keys=False).apply(filter_outliers)
print(f"After outlier filtering: {len(df_clean)} records")

# === Prix moyen par ref × year_month ===
df_clean['year_month'] = df_clean['date'].dt.to_period('M').astype(str)
prix_monthly = df_clean.groupby(['ref', 'year_month']).agg(
    prix_ttc=('prix_sac_50_ttc', 'mean'),
    prix_ht=('prix_sac_50_ht', 'mean'),
    n_records=('sacs_50', 'count'),
    total_sacs=('sacs_50', 'sum'),
    total_ttc=('montant_ttc', 'sum'),
).reset_index()
prix_monthly['prix_ttc'] = prix_monthly['prix_ttc'].round(0)
prix_monthly['prix_ht'] = prix_monthly['prix_ht'].round(0)

print(f"\nPrix mensuels: {len(prix_monthly)} rows")
print("\nSample (TOURTEAUX T102):")
print(prix_monthly[prix_monthly['ref'] == 'T102'].tail(15))

print("\nSample (CONCENTRES C101):")
print(prix_monthly[prix_monthly['ref'] == 'C101'].tail(15))

# === Prix le plus récent (Q4 forecast) ===
# Get the latest price for each ref (most recent month with data)
prix_current = df_clean.sort_values('date').groupby('ref').last().reset_index()[['ref', 'prix_sac_50_ttc', 'prix_sac_50_ht']]
prix_current.columns = ['ref', 'prix_ttc', 'prix_ht']
prix_current = prix_current.merge(df[['ref', 'description', 'family']].drop_duplicates(), on='ref')
prix_current['prix_ttc'] = prix_current['prix_ttc'].round(0)
prix_current['prix_ht'] = prix_current['prix_ht'].round(0)

# Apply Option 3 (price stability Q4): use August 2026 prices as forecast prices
# But check if August data exists for each ref; if not, use latest available
print("\n=== PRIX ACTUELS (Q4 2026 forecast) ===")
print(prix_current.sort_values(['family', 'ref']).to_string(index=False))

# === Apply price hike context ===
# Soja: +3 000 FCFA/sac cumulé (1 000 + 2 000 au 23/07)
# Concentrés: prix stable depuis début 2026
# For Q4 forecast, we use Option A: prices stable at August 2026 level
prix_forecast = prix_current.copy()
prix_forecast['scenario'] = 'stable_Q4'

# === Bonus scenario: price decrease if stock stabilizes ===
# Hypothetical -10% on soja only
prix_forecast_baisse = prix_current.copy()
prix_forecast_baisse.loc[prix_forecast_baisse['family'] == 'TOURTEAUX', 'prix_ttc'] = (
    prix_forecast_baisse.loc[prix_forecast_baisse['family'] == 'TOURTEAUX', 'prix_ttc'] * 0.90
).round(0)
prix_forecast_baisse.loc[prix_forecast_baisse['family'] == 'TOURTEAUX', 'prix_ht'] = (
    prix_forecast_baisse.loc[prix_forecast_baisse['family'] == 'TOURTEAUX', 'prix_ht'] * 0.90
).round(0)
prix_forecast_baisse['scenario'] = 'baisse_prix_soja_-10%'

# Save outputs
prix_monthly.to_csv("/home/z/my-project/scripts/prix_moyens_monthly.csv", index=False)
prix_current.to_csv("/home/z/my-project/scripts/prix_moyens_current.csv", index=False)
prix_forecast.to_csv("/home/z/my-project/scripts/prix_forecast_stable.csv", index=False)
prix_forecast_baisse.to_csv("/home/z/my-project/scripts/prix_forecast_baisse.csv", index=False)

# Save as JSON for easy access
prix_dict = {
    'stable_Q4': prix_forecast.set_index('ref')['prix_ttc'].to_dict(),
    'baisse_prix_soja_-10%': prix_forecast_baisse.set_index('ref')['prix_ttc'].to_dict(),
}
with open("/home/z/my-project/scripts/prix_forecast.json", 'w') as f:
    json.dump(prix_dict, f, indent=2)

print(f"\n=== OUTPUTS ===")
print(f"  prix_moyens_monthly.csv: {len(prix_monthly)} rows")
print(f"  prix_moyens_current.csv: {len(prix_current)} refs")
print(f"  prix_forecast_stable.csv: {len(prix_forecast)} refs (Q4 stable)")
print(f"  prix_forecast_baisse.csv: {len(prix_forecast_baisse)} refs (Q4 avec baisse soja -10%)")
print(f"  prix_forecast.json: dictionnaire de scénarios")

# Summary by family
print(f"\n=== PRIX MOYEN PAR FAMILLE (Q4 stable) ===")
print(prix_current.groupby('family')['prix_ttc'].agg(['mean', 'min', 'max', 'count']).round(0))
