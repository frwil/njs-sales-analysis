"""Generate 2 forecast simulations for 2027:
1. _upd: +23% SOJA, +50% CB100+CB200, +3% PREMIX, maintain COMPLEMENT_ALIM, +25% ALVEOLES CA, +30% MAT_ELEVAGE CA
2. _35_percent: Non-linear +35% CA total (TOURTEAUX +38.6%, CONC keep Prophet, INGREDIENTS +20%, ALIMENT_COMPLET +60%, PREMIX +15%, COMPLEMENT_ALIM +30%, ALVEOLES +35% CA, MAT_ELEVAGE +35% CA)

Also adds 2024-2026 historical data (vol + CA) by family to the exec summary.

Approach: Take the existing disaggregated forecast CSV (forecast_2027_S3.csv) and apply
scaling factors per family to the tonnes and ca_m_fcfa columns. The distribution across
agences/months is preserved (proportional scaling).
"""
import pandas as pd
import numpy as np
import json
import os

# === Load source data ===
print("Loading source data...")
fc = pd.read_csv('/home/z/my-project/scripts/forecast_2027_S3.csv', parse_dates=['date'])
df_hist = pd.read_csv('/home/z/my-project/scripts/dataset_2023_2026.csv', parse_dates=['date'], low_memory=False)
df_hist['ca_combined'] = df_hist['montant_ttc'].where(df_hist['montant_ttc']>0, df_hist['montant_ht'])
q4 = pd.read_csv('/home/z/my-project/scripts/forecast_q4_2026_S3.csv')
print(f"  Loaded {len(fc)} forecast records, {len(df_hist)} historical records")

# === Compute 2026 LY values by family ===
print("\nComputing 2026 LY values by family...")
ly_2026 = {}
for fam in fc['family'].unique():
    ytd = df_hist[(df_hist['date'].dt.year==2026) & (df_hist['date'].dt.month<=8) & (df_hist['family']==fam)]
    q4_fam = q4[q4['family']==fam]
    vol = ytd['tonnes'].sum() + q4_fam['tonnes'].sum()
    ca = ytd['ca_combined'].sum()/1e6 + q4_fam['ca_m_fcfa'].sum()
    ly_2026[fam] = {'vol': vol, 'ca': ca}

# Also compute for CB100+CB200 specifically
ytd_cb = df_hist[(df_hist['date'].dt.year==2026) & (df_hist['date'].dt.month<=8) & (df_hist['family']=='ALIMENT_COMPLET') & (df_hist['ref'].isin(['CB100', 'CB200']))]
q4_cb = q4[(q4['family']=='ALIMENT_COMPLET') & (q4['ref'].isin(['CB100', 'CB200']))]
cb_vol = ytd_cb['tonnes'].sum() + q4_cb['tonnes'].sum()
cb_ca = ytd_cb['ca_combined'].sum()/1e6 + q4_cb['ca_m_fcfa'].sum()
ly_2026_cb = {'vol': cb_vol, 'ca': cb_ca}

print("  2026 LY by family:")
for fam, vals in ly_2026.items():
    print(f"    {fam}: {vals['vol']:.0f} t / {vals['ca']:.1f} M")
print(f"  CB100+CB200: {cb_vol:.0f} t / {cb_ca:.1f} M")

# === Compute 2024, 2025 historical ===
hist_by_year = {}
for year in [2024, 2025]:
    hist_by_year[year] = {}
    for fam in fc['family'].unique():
        d = df_hist[(df_hist['date'].dt.year==year) & (df_hist['family']==fam)]
        vol = d['tonnes'].sum()
        ca = d['ca_combined'].sum()/1e6
        hist_by_year[year][fam] = {'vol': vol, 'ca': ca}
# Add MAIS for historical
for year in [2024, 2025]:
    d = df_hist[(df_hist['date'].dt.year==year) & (df_hist['family']=='MAIS')]
    hist_by_year[year]['MAIS'] = {'vol': d['tonnes'].sum(), 'ca': d['ca_combined'].sum()/1e6}

# === SIMULATION 1: _upd (+23% SOJA cross-sell) ===
print("\n" + "="*80)
print("SIMULATION 1: _upd (+23% SOJA cross-sell)")
print("="*80)

fc_upd = fc.copy()

# Rules for _upd
# 1. TOURTEAUX: scale to 2026 LY × 1.23
# 2. CONCENTRES: keep as-is (Prophet + bundle already applied)
# 3. ALIMENT_COMPLET CB100+CB200: scale to 2026 LY × 1.50
# 4. ALIMENT_COMPLET other refs: keep as-is
# 5. PREMIX: scale to 2026 LY × 1.03
# 6. COMPLEMENT_ALIMENTAIRE: scale to 2026 LY (maintain)
# 7. ALVEOLES: scale CA to 2026 LY × 1.25
# 8. MATERIEL_ELEVAGE: scale CA to 2026 LY × 1.30

for fam in fc_upd['family'].unique():
    fam_mask = fc_upd['family'] == fam
    
    if fam == 'TOURTEAUX':
        # Scale to 2026 LY × 1.23
        target_vol = ly_2026[fam]['vol'] * 1.23
        current_vol = fc_upd[fam_mask]['tonnes'].sum()
        if current_vol > 0:
            scale = target_vol / current_vol
            fc_upd.loc[fam_mask, 'tonnes'] *= scale
            fc_upd.loc[fam_mask, 'ca_m_fcfa'] *= scale
        print(f"  TOURTEAUX: scaled by {scale:.4f} -> {fc_upd[fam_mask]['tonnes'].sum():.0f} t / {fc_upd[fam_mask]['ca_m_fcfa'].sum():.0f} M")
    
    elif fam == 'PREMIX':
        target_vol = ly_2026[fam]['vol'] * 1.03
        current_vol = fc_upd[fam_mask]['tonnes'].sum()
        if current_vol > 0:
            scale = target_vol / current_vol
            fc_upd.loc[fam_mask, 'tonnes'] *= scale
            fc_upd.loc[fam_mask, 'ca_m_fcfa'] *= scale
        print(f"  PREMIX: scaled by {scale:.4f} -> {fc_upd[fam_mask]['tonnes'].sum():.0f} t / {fc_upd[fam_mask]['ca_m_fcfa'].sum():.0f} M")
    
    elif fam == 'COMPLEMENT_ALIMENTAIRE':
        target_vol = ly_2026[fam]['vol']  # maintain 2026
        current_vol = fc_upd[fam_mask]['tonnes'].sum()
        if current_vol > 0:
            scale = target_vol / current_vol
            fc_upd.loc[fam_mask, 'tonnes'] *= scale
            fc_upd.loc[fam_mask, 'ca_m_fcfa'] *= scale
        print(f"  COMPLEMENT_ALIM: scaled by {scale:.4f} -> {fc_upd[fam_mask]['tonnes'].sum():.0f} t / {fc_upd[fam_mask]['ca_m_fcfa'].sum():.0f} M")
    
    elif fam == 'ALIMENT_COMPLET':
        # CB100+CB200: +50% vs 2026; others: keep Prophet
        cb_mask = fam_mask & fc_upd['ref'].isin(['CB100', 'CB200'])
        other_mask = fam_mask & ~fc_upd['ref'].isin(['CB100', 'CB200'])
        
        # CB100+CB200
        target_cb_vol = ly_2026_cb['vol'] * 1.50
        current_cb_vol = fc_upd[cb_mask]['tonnes'].sum()
        if current_cb_vol > 0:
            scale_cb = target_cb_vol / current_cb_vol
            fc_upd.loc[cb_mask, 'tonnes'] *= scale_cb
            fc_upd.loc[cb_mask, 'ca_m_fcfa'] *= scale_cb
        print(f"  ALIMENT_COMPLET CB100+CB200: scaled by {scale_cb:.4f} -> {fc_upd[cb_mask]['tonnes'].sum():.0f} t / {fc_upd[cb_mask]['ca_m_fcfa'].sum():.0f} M")
        print(f"  ALIMENT_COMPLET autres: keep Prophet -> {fc_upd[other_mask]['tonnes'].sum():.0f} t / {fc_upd[other_mask]['ca_m_fcfa'].sum():.0f} M")
    
    elif fam == 'ALVEOLES':
        # CA only: +25% vs 2026
        target_ca = ly_2026[fam]['ca'] * 1.25
        current_ca = fc_upd[fam_mask]['ca_m_fcfa'].sum()
        if current_ca > 0:
            scale = target_ca / current_ca
            fc_upd.loc[fam_mask, 'ca_m_fcfa'] *= scale
        print(f"  ALVEOLES: CA scaled by {scale:.4f} -> {fc_upd[fam_mask]['ca_m_fcfa'].sum():.0f} M")
    
    elif fam == 'MATERIEL_ELEVAGE':
        # CA only: +30% vs 2026
        target_ca = ly_2026[fam]['ca'] * 1.30
        current_ca = fc_upd[fam_mask]['ca_m_fcfa'].sum()
        if current_ca > 0:
            scale = target_ca / current_ca
            fc_upd.loc[fam_mask, 'ca_m_fcfa'] *= scale
        print(f"  MATERIEL_ELEVAGE: CA scaled by {scale:.4f} -> {fc_upd[fam_mask]['ca_m_fcfa'].sum():.0f} M")
    
    else:
        # CONCENTRES, INGREDIENTS: keep as-is
        print(f"  {fam}: keep Prophet -> {fc_upd[fam_mask]['tonnes'].sum():.0f} t / {fc_upd[fam_mask]['ca_m_fcfa'].sum():.0f} M")

# Save _upd CSV
upd_path = '/home/z/my-project/scripts/forecast_2027_S3_upd.csv'
fc_upd.to_csv(upd_path, index=False)
print(f"\n  Saved: {upd_path}")
print(f"  Total: {fc_upd['tonnes'].sum():.0f} t / {fc_upd['ca_m_fcfa'].sum():.0f} M")

# === SIMULATION 2: _35_percent (non-linear +35% CA) ===
print("\n" + "="*80)
print("SIMULATION 2: _35_percent (non-linear +35% CA)")
print("="*80)

fc_35 = fc.copy()

# Rules for _35_percent
rules_35 = {
    'TOURTEAUX': {'vol_mult': 1.386, 'ca_mult': 1.386},  # +38.6% to hit +35% total
    'CONCENTRES': None,  # Keep Prophet
    'INGREDIENTS': {'vol_mult': 1.20, 'ca_mult': 1.20},
    'ALIMENT_COMPLET': {'vol_mult': 1.60, 'ca_mult': 1.60},  # +60% on entire family
    'PREMIX': {'vol_mult': 1.15, 'ca_mult': 1.15},
    'COMPLEMENT_ALIMENTAIRE': {'vol_mult': 1.30, 'ca_mult': 1.30},
    'ALVEOLES': {'vol_mult': None, 'ca_mult': 1.35},  # CA only
    'MATERIEL_ELEVAGE': {'vol_mult': None, 'ca_mult': 1.35},  # CA only
}

for fam, rule in rules_35.items():
    if rule is None:
        print(f"  {fam}: keep Prophet -> {fc_35[fc_35['family']==fam]['tonnes'].sum():.0f} t / {fc_35[fc_35['family']==fam]['ca_m_fcfa'].sum():.0f} M")
        continue
    
    fam_mask = fc_35['family'] == fam
    current_vol = fc_35[fam_mask]['tonnes'].sum()
    current_ca = fc_35[fam_mask]['ca_m_fcfa'].sum()
    
    # Scale to 2026 LY × multiplier
    if rule['vol_mult']:
        target_vol = ly_2026[fam]['vol'] * rule['vol_mult']
        if current_vol > 0:
            vol_scale = target_vol / current_vol
            fc_35.loc[fam_mask, 'tonnes'] *= vol_scale
    
    if rule['ca_mult']:
        target_ca = ly_2026[fam]['ca'] * rule['ca_mult']
        if current_ca > 0:
            ca_scale = target_ca / current_ca
            fc_35.loc[fam_mask, 'ca_m_fcfa'] *= ca_scale
    
    new_vol = fc_35[fam_mask]['tonnes'].sum()
    new_ca = fc_35[fam_mask]['ca_m_fcfa'].sum()
    pct_vol = (new_vol / ly_2026[fam]['vol'] - 1) * 100 if ly_2026[fam]['vol'] > 0 else 0
    pct_ca = (new_ca / ly_2026[fam]['ca'] - 1) * 100 if ly_2026[fam]['ca'] > 0 else 0
    print(f"  {fam}: {new_vol:.0f} t ({pct_vol:+.1f}%) / {new_ca:.0f} M ({pct_ca:+.1f}%)")

# Save _35_percent CSV
csv_35_path = '/home/z/my-project/scripts/forecast_2027_S3_35percent.csv'
fc_35.to_csv(csv_35_path, index=False)
print(f"\n  Saved: {csv_35_path}")
print(f"  Total: {fc_35['tonnes'].sum():.0f} t / {fc_35['ca_m_fcfa'].sum():.0f} M")

# === Save historical data JSON for PDF generation ===
hist_data = {'by_year': {}, '2026_ly': {}}
for year in [2024, 2025]:
    hist_data['by_year'][str(year)] = {}
    for fam in ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'ALIMENT_COMPLET', 'PREMIX', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES', 'MATERIEL_ELEVAGE', 'MAIS']:
        d = hist_by_year[year].get(fam, {'vol': 0, 'ca': 0})
        hist_data['by_year'][str(year)][fam] = {'vol': float(d['vol']), 'ca': float(d['ca'])}

for fam in ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'ALIMENT_COMPLET', 'PREMIX', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES', 'MATERIEL_ELEVAGE', 'MAIS']:
    d = ly_2026.get(fam, {'vol': 0, 'ca': 0})
    hist_data['2026_ly'][fam] = {'vol': float(d['vol']), 'ca': float(d['ca'])}

# Add MAIS 2026 LY
d_mais_2026 = df_hist[(df_hist['date'].dt.year==2026) & (df_hist['date'].dt.month<=8) & (df_hist['family']=='MAIS')]
hist_data['2026_ly']['MAIS'] = {'vol': float(d_mais_2026['tonnes'].sum()), 'ca': float(d_mais_2026['ca_combined'].sum()/1e6)}

hist_path = '/home/z/my-project/scripts/historical_2024_2026.json'
with open(hist_path, 'w') as f:
    json.dump(hist_data, f, indent=2)
print(f"\n  Historical data saved: {hist_path}")

# === Summary ===
print("\n" + "="*80)
print("SUMMARY")
print("="*80)
print(f"{'Simulation':<35} {'Volume (t)':>12} {'CA (M)':>12} {'D vs 2026 vol':>14} {'D vs 2026 CA':>14}")
print('-'*90)
total_2026_vol = sum(v['vol'] for v in ly_2026.values())
total_2026_ca = sum(v['ca'] for v in ly_2026.values())
print(f"{'2026 LY (hors MAIS)':<35} {total_2026_vol:>12,.0f} {total_2026_ca:>12,.0f} {'(reference)':>14} {'(reference)':>14}".replace(',', ' '))
print(f"{'_upd (+23% SOJA cross-sell)':<35} {fc_upd['tonnes'].sum():>12,.0f} {fc_upd['ca_m_fcfa'].sum():>12,.0f} {(fc_upd['tonnes'].sum()/total_2026_vol-1)*100:>+13.1f}% {(fc_upd['ca_m_fcfa'].sum()/total_2026_ca-1)*100:>+13.1f}%".replace(',', ' '))
print(f"{'_35_percent (non-linear +35%)':<35} {fc_35['tonnes'].sum():>12,.0f} {fc_35['ca_m_fcfa'].sum():>12,.0f} {(fc_35['tonnes'].sum()/total_2026_vol-1)*100:>+13.1f}% {(fc_35['ca_m_fcfa'].sum()/total_2026_ca-1)*100:>+13.1f}%".replace(',', ' '))
print()
print("CSV files generated:")
print(f"  1. {upd_path}")
print(f"  2. {csv_35_path}")
print(f"  3. {hist_path}")
