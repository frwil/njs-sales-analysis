"""Apply BELGO FISH separation to all 3 forecast versions:
1. Remove BELGO FISH (AP*) refs from ALIMENT_COMPLET
2. Create new BELGO_FISH family with 216 t/week target
3. Apply seasonality (dry season peak)
4. Distribute by agence using historical 2025 weights
5. Apply new price (-4,000 FCFA/sac 15kg)
6. Save updated CSVs for all 3 versions
"""
import pandas as pd
import json
import os

config = json.load(open('belgofish_forecast_config.json'))
CLARIA_REFS = config['claria_refs']
WEIGHTS = {k: int(v) for k, v in config['weights'].items()}
SEASONALITY = {int(k): float(v) for k, v in config['seasonality'].items()}
AG_WEIGHTS = {k: float(v) for k, v in config['agence_weights'].items()}
NEW_PRICE_KG = config['new_price_kg']
NEW_PRICE_SAC15 = config['new_price_sac15']

# Load descriptions
desc_df = pd.read_csv('dataset_2023_2026.csv', low_memory=False)
desc_map = desc_df[['ref', 'description']].drop_duplicates().set_index('ref')['description'].to_dict()

# Agence → region mapping
AGENCE_MAP = {
    'Ahala': ('Ahala', 'Centre'), 'Messassi': ('Messassi', 'Centre'),
    'Ndobo': ('Ndobo', 'Littoral'), 'Famla': ('Famla', 'Ouest'),
    'Bertoua': ('Bertoua', 'Centre'), 'Village': ('Village', 'Littoral'),
    'Nkolbisson': ('Nkolbisson', 'Centre'), 'Nkoabang': ('Nkoabang', 'Centre'),
    'Nkongsamba': ('Nkongsamba', 'Littoral'), 'Djeleng': ('Djeleng', 'Ouest'),
    'Pk11': ('Pk11', 'Littoral'), 'Ngaoundere': ('Ngaoundere', 'Centre'),
    'Buea': ('Buea', 'Littoral'), 'Mbouda': ('Mbouda', 'Ouest'),
}

# Distribute CLARIA refs within each agence (proportional to 2025 historical)
df = pd.read_csv('dataset_2023_2026.csv', parse_dates=['date'], low_memory=False)
bf_2025 = df[(df['date'].dt.year == 2025) & (df['ref'].isin(CLARIA_REFS))].copy()
bf_2025['correct_kg'] = bf_2025.apply(lambda r: r['qte'] * WEIGHTS.get(r['ref'], 15), axis=1)

# Ref weights within each agence (proportional to 2025 volume)
ref_weights_by_ag = {}
for ag in AG_WEIGHTS.keys():
    ag_data = bf_2025[bf_2025['agence'] == ag]
    if len(ag_data) > 0:
        ref_vol = ag_data.groupby('ref')['correct_kg'].sum()
        total = ref_vol.sum()
        if total > 0:
            ref_weights_by_ag[ag] = {r: v/total for r, v in ref_vol.items()}
    if ag not in ref_weights_by_ag:
        # Default: equal distribution across top 5 CLARIA refs
        ref_weights_by_ag[ag] = {'APCL3': 0.25, 'APCL4.5': 0.25, 'APCL6': 0.20, 'APCL2': 0.15, 'APCL8': 0.15}

def generate_belgofish_records():
    """Generate BELGO FISH forecast records for all 12 months × agences × refs."""
    records = []
    target_annual = config['target_annual']
    
    for month, season_pct in SEASONALITY.items():
        month_vol = target_annual * season_pct
        
        for ag, ag_pct in AG_WEIGHTS.items():
            ag_vol = month_vol * ag_pct / 100
            agence_short, region = AGENCE_MAP.get(ag, (ag, '?'))
            
            ref_dist = ref_weights_by_ag.get(ag, {'APCL3': 0.25, 'APCL4.5': 0.25, 'APCL6': 0.20, 'APCL2': 0.15, 'APCL8': 0.15})
            
            for ref, ref_pct in ref_dist.items():
                ref_vol = ag_vol * ref_pct
                weight_kg = WEIGHTS.get(ref, 15)
                units = ref_vol * 1000 / weight_kg
                sacs_50 = ref_vol * 1000 / 50  # standard sacs équivalent
                ca_m = ref_vol * 1000 * NEW_PRICE_KG / 1e6
                
                records.append({
                    'scenario': 'S3_reappro_100',
                    'ref': ref,
                    'family': 'BELGO_FISH',
                    'agence': agence_short,
                    'region': region,
                    'date': f'2027-{month:02d}-01',
                    'month': month,
                    'year': 2027,
                    'tonnes': round(ref_vol, 2),
                    'sacs_50': round(units, 1),
                    'prix_ttc_sac': int(NEW_PRICE_SAC15) if weight_kg == 15 else int(NEW_PRICE_SAC15 * weight_kg / 15),
                    'ca_m_fcfa': round(ca_m, 2),
                })
    
    return pd.DataFrame(records)

# Generate BELGO FISH records
print("Generating BELGO FISH records...")
bf_records = generate_belgofish_records()
print(f"  {len(bf_records)} records")
print(f"  Total: {bf_records['tonnes'].sum():.0f} t / {bf_records['ca_m_fcfa'].sum():.0f} M FCFA")
print(f"  Agences: {bf_records['agence'].nunique()}")
print(f"  Refs: {sorted(bf_records['ref'].unique())}")

# Process all 3 versions
versions = [
    ('forecast_2027_S3.csv', 'forecast_2027_S3.csv', 'ORIGINAL'),
    ('forecast_2027_S3_upd.csv', 'forecast_2027_S3_upd.csv', '_upd'),
    ('forecast_2027_S3_35percent.csv', 'forecast_2027_S3_35percent.csv', '_35pct'),
]

for csv_in, csv_out, label in versions:
    csv_path = f'/home/z/my-project/scripts/{csv_in}'
    print(f"\n{'='*60}")
    print(f"Processing {label}...")
    
    fc = pd.read_csv(csv_path, parse_dates=['date'])
    
    # 1. Remove BELGO FISH (AP*) from ALIMENT_COMPLET
    bf_mask = fc['ref'].str.startswith('AP')
    bf_removed = fc[bf_mask].copy()
    fc_clean = fc[~bf_mask].copy()
    print(f"  Removed BELGO FISH from ALIMENT_COMPLET: {len(bf_removed)} records ({bf_removed['tonnes'].sum():.1f} t)")
    
    # 2. Add new BELGO_FISH family
    bf_records_copy = bf_records.copy()
    # Convert date to datetime for consistency
    bf_records_copy['date'] = pd.to_datetime(bf_records_copy['date'])
    
    # 3. Merge
    fc_final = pd.concat([fc_clean, bf_records_copy], ignore_index=True)
    
    # Save
    out_path = f'/home/z/my-project/scripts/{csv_out}'
    fc_final.to_csv(out_path, index=False)
    
    # Summary
    print(f"  ALIMENT_COMPLET (sans BF): {fc_clean[fc_clean['family']=='ALIMENT_COMPLET']['tonnes'].sum():.0f} t / {fc_clean[fc_clean['family']=='ALIMENT_COMPLET']['ca_m_fcfa'].sum():.0f} M")
    print(f"  BELGO_FISH (nouvelle): {bf_records_copy['tonnes'].sum():.0f} t / {bf_records_copy['ca_m_fcfa'].sum():.0f} M")
    print(f"  Total: {fc_final['tonnes'].sum():.0f} t / {fc_final['ca_m_fcfa'].sum():.0f} M")
    print(f"  Families: {sorted(fc_final['family'].unique())}")
    print(f"  Saved: {out_path}")

print(f"\n{'='*60}")
print("ALL 3 VERSIONS UPDATED")
