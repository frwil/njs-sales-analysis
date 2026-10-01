"""
Forecast 2027 complet (12 mois) - Volume + Valeur
Scénario S3 (réappro soja 100%)

VERSION 3 (mise à jour avec septembre réel):
1. Utilise données 2023-2026 (44 mois, 264 759 records) + septembre 2026 Livrée
2. Inclut la famille COMPLEMENT_ALIMENTAIRE (BELGOKILL V300 1L only + autres CA001-CA008)
   - V305 (BELGOKILL 200L) EXCLU selon demande utilisateur
   - Conversion 1L = 1kg
3. MATERIEL_ELEVAGE toujours à 0 en tonnes (CA only)
4. Désaisonnalisation effet soja (cap moyenne S1 2026) maintenue
5. En cours + Validées septembre inclus
6. Prix soja actualisé 16 550 FCFA/sac (médiane mensuelle YTD Jan-Sep 2026)

Méthode:
  - Prophet pour 5 familles (TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, ALIMENT_COMPLET, COMPLEMENT_ALIMENTAIRE)
  - Extrapolation pour MATERIEL_ELEVAGE et PREMIX (CA only)
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

# === Load dataset 2023-2026 (nouveau) ===
df = pd.read_csv("/home/z/my-project/scripts/dataset_2023_2026.csv", parse_dates=['date'], low_memory=False)
print(f"Loaded {len(df)} records (Livrée only, 2023-2026)")
print(f"Date range: {df['date'].min().date()} → {df['date'].max().date()}")
print(f"\nBy family:")
print(df['family'].value_counts())

# === Load En cours + Validées from latest extraction ===
print("\n=== Loading En cours + Validées from latest extraction ===")
import openpyxl

FILE_AOUT = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (51).xlsx"
wb = openpyxl.load_workbook(FILE_AOUT, read_only=True, data_only=True)
ws = wb['Sheet 1']

# Product refs and weights (same as dataset_2024_2026)
SOJA_REFS = {'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25}
CONC_REFS = {'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
             'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50, 'C1022': 5}
ALIMENT_REFS = {'CB100': 25, 'CB200': 25, 'CB101': 5, 'CB201': 5, 'PB100': 25, 'PB200': 25, 
                'DB100': 25, 'DB200': 25, 'ALAP25': 25}
INGREDIENT_REFS = {'B100': 25, 'E101': 25, 'I105': 25, 'B1001': 1, 'B1003': 5, 'B1004': 25,
                   'E1011': 1, 'E1013': 5, 'E1014': 0.2, 'I1051': 1, 'I1053': 5, 'I1054': 25,
                   'I106': 25, 'I1061': 1, 'I107': 25, 'I1071': 1, 'P105': 25, 'P1051': 1, 'P1053': 5,
                   'F114': 50, 'F1145': 50, 'F1146': 25, 'F1147': 1}
PREMIX_REFS = {'P102N2': 25, 'P104N2': 25, 'P109': 25, 'PX101': 25, 'PX102': 25, 'PX103': 25, 'PX104': 25, 'PX105': 25}
MATERIEL_REFS = {f'MAT{i:03d}': 1 for i in range(1, 100)}
MATERIEL_REFS.update({'MAT014-80010003': 1, 'MAT011-80010002': 1})
# NOUVEAU: COMPLEMENT ALIMENTAIRE (liquides) — V300 1L only (V305 200L EXCLU)
COMPLEMENT_REFS = {
    'V300': 1,        # BELGOKILL 1L = 1 kg
    'CA003.1': 1,     # BELGO HARMONY 1L = 1 kg
    'CA004.1': 1,     # BELGO PROTECT 1L = 1 kg
    'CA006.1': 1,     # BELGO DRY LIT 1L = 1 kg
    'CA001.1': 1,     # BELGO WATER CLEAN 1L = 1 kg
    'CA002.1': 1,     # BELGO VIT Ese 1L = 1 kg
    'CA005.1': 1,     # BELGO THERMO 1L = 1 kg
    'CA007.1': 1,     # BELGO BIO SELECT 1L = 1 kg
    'CA008.1': 1,     # BELGO FRESH 1L = 1 kg
}

ALL_REFS = {**SOJA_REFS, **CONC_REFS, **ALIMENT_REFS, **INGREDIENT_REFS, 
            **PREMIX_REFS, **MATERIEL_REFS, **COMPLEMENT_REFS}

# ALVEOLES refs (separated from MATERIEL_ELEVAGE)
ALVEOLES_REFS = {'MAT011-80010002', 'MAT014-80010003', 'MAT015', 'MAT017'}

def get_family_2027(ref):
    if ref in SOJA_REFS: return 'TOURTEAUX'
    if ref in CONC_REFS: return 'CONCENTRES'
    if ref in ALIMENT_REFS: return 'ALIMENT_COMPLET'
    if ref in INGREDIENT_REFS: return 'INGREDIENTS'
    if ref in PREMIX_REFS: return 'PREMIX'
    if ref in ALVEOLES_REFS: return 'ALVEOLES'
    if ref in MATERIEL_REFS or ref.startswith('MAT') or ref.startswith('ME'): return 'MATERIEL_ELEVAGE'
    if ref in COMPLEMENT_REFS: return 'COMPLEMENT_ALIMENTAIRE'
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

# Collect En cours + Validées
extra_records = []
for r in ws.iter_rows(min_row=3, values_only=True):
    if not r or not r[0] or r[0] == 'Total': continue
    ref = str(r[0])
    family = get_family_2027(ref)
    if family is None: continue
    etat = str(r[13]).strip() if r[13] else ''
    if etat not in ('En cours', 'Validée'): continue  # Only non-Livrée
    date_str = str(r[6])[:10] if r[6] else ''
    if '/09/2026' not in date_str: continue
    agence_raw = r[15] if r[15] else ''
    if agence_raw not in AGENCE_MAP: continue
    agence, region = AGENCE_MAP[agence_raw]
    qte = r[2] or 0
    weight = ALL_REFS.get(ref, 1)
    kg = qte * weight
    montant_ttc = r[9] or 0
    # Parse date
    try:
        date = pd.to_datetime(date_str, format='%d/%m/%Y')
    except:
        continue
    
    # For MATERIEL_ELEVAGE, ALVEOLES: tonnes=0 (CA only)
    if family in ('MATERIEL_ELEVAGE', 'ALVEOLES'):
        tonnes_val = 0
    else:
        tonnes_val = kg / 1000
    
    extra_records.append({
        'date': date, 'year': date.year, 'month': date.month,
        'ref': ref, 'family': family, 'description': r[1],
        'agence': agence, 'region': region,
        'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': tonnes_val,
        'sacs_50': kg / 50 if family not in ('MATERIEL_ELEVAGE', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES') else 0,
        'montant_ttc': montant_ttc, 'montant_ht': 0,
        'source': 'En_cours_Validee'
    })

print(f"  En cours + Validées: {len(extra_records)} records")
df_extra = pd.DataFrame(extra_records)

# === Charger septembre LIVRÉE (ventes réelles 01-30/09) dans l'historique ===
print("\n=== Loading Septembre Livrée (ventes réelles) ===")
sep_records = []
for r in ws.iter_rows(min_row=3, values_only=True):
    if not r or not r[0] or r[0] == 'Total': continue
    ref = str(r[0])
    family = get_family_2027(ref)
    if family is None: continue
    etat = str(r[13]).strip() if r[13] else ''
    if etat != 'Livrée': continue
    date_str = str(r[6])[:10] if r[6] else ''
    if '/09/2026' not in date_str: continue
    agence_raw = r[15] if r[15] else ''
    if agence_raw not in AGENCE_MAP: continue
    agence, region = AGENCE_MAP[agence_raw]
    qte = r[2] or 0
    weight = ALL_REFS.get(ref, 1)
    kg = qte * weight
    montant_ttc = r[9] or 0
    montant_ht = r[8] or 0
    try:
        date = pd.to_datetime(date_str, format='%d/%m/%Y')
    except:
        continue
    if family in ('MATERIEL_ELEVAGE', 'ALVEOLES'):
        tonnes_val = 0
    else:
        tonnes_val = kg / 1000
    sep_records.append({
        'date': date, 'year': date.year, 'month': date.month,
        'ref': ref, 'family': family, 'description': r[1],
        'agence': agence, 'region': region,
        'qte': qte, 'weight_kg': weight, 'kg': kg, 'tonnes': tonnes_val,
        'sacs_50': kg / 50 if family not in ('MATERIEL_ELEVAGE', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES') else 0,
        'montant_ttc': montant_ttc, 'montant_ht': montant_ht,
        'source': 'Livree_Septembre'
    })

print(f"  Septembre Livrée: {len(sep_records)} records")
df_sep = pd.DataFrame(sep_records)
print(f"  Septembre réel: {df_sep['tonnes'].sum():.0f} t, CA {df_sep['montant_ttc'].sum()/1e6:.0f} M FCFA (14 agences)")

# === Merge Livrée + En cours/Validées + Septembre réel ===
df_all = pd.concat([df, df_extra, df_sep], ignore_index=True)
print(f"\nDataset total (Livrée + En cours + Validées + Sept réel): {len(df_all)} records")
print(f"Date range: {df_all['date'].min().date()} → {df_all['date'].max().date()}")

# === Désaisonnalisation de l'effet soja exceptionnel ===
print("\n=== DÉSAISONNALISATION EFFET SOJA EXCEPTIONNEL ===")
df_all['year_month'] = df_all['date'].dt.to_period('M')

# Calculate S1 average monthly soja volume (Jan-Juin 2026)
soja_s1 = df_all[(df_all['family'] == 'TOURTEAUX') & (df_all['date'] < '2026-07-01') & (df_all['date'] >= '2026-01-01')]
soja_s1_monthly = soja_s1.groupby('year_month')['tonnes'].sum()
soja_s1_avg = soja_s1_monthly.mean()
print(f"  S1 2026 average monthly soja: {soja_s1_avg:.0f} t")

# July-August 2026 soja volume
soja_juil_aout = df_all[(df_all['family'] == 'TOURTEAUX') & 
                         (df_all['date'] >= '2026-07-01') & 
                         (df_all['date'] <= '2026-08-31')]
soja_juil_aout_monthly = soja_juil_aout.groupby('year_month')['tonnes'].sum()
print(f"  Jul-Aug 2026 soja monthly volumes:")
for ym, t in soja_juil_aout_monthly.items():
    print(f"    {ym}: {t:.0f} t")

# Cap soja to S1 average for July-August 2026
cap_factor = soja_s1_avg / soja_juil_aout_monthly.max() if soja_juil_aout_monthly.max() > 0 else 1.0
print(f"  Cap factor: {cap_factor:.2f} (soja Jul-Aug capped at S1 avg {soja_s1_avg:.0f} t)")

# Apply cap to soja records in Jul-Aug 2026
mask_cap = (df_all['family'] == 'TOURTEAUX') & (df_all['date'] >= '2026-07-01') & (df_all['date'] <= '2026-08-31')
df_all.loc[mask_cap, 'tonnes'] = df_all.loc[mask_cap, 'tonnes'] * cap_factor
df_all.loc[mask_cap, 'sacs_50'] = df_all.loc[mask_cap, 'sacs_50'] * cap_factor
df_all.loc[mask_cap, 'kg'] = df_all.loc[mask_cap, 'kg'] * cap_factor

# === Prix 2027 ===
print("\n=== PRIX 2027 ===")
# Load current prices
prix_forecast = json.load(open("/home/z/my-project/scripts/prix_forecast.json"))
prix_2027 = prix_forecast['stable_Q4'].copy()

# Update soja price to 25 000 FCFA/sac (effective 24/08/2026)
for ref in SOJA_REFS:
    if ref in prix_2027:
        prix_2027[ref] = 25000 if ref == 'T102' else int(prix_2027[ref] * 25000 / 20600)

# Add PREMIX prices
PREMIX_PRICES = {'P102N2': 95147, 'P104N2': 91096, 'P109': 34000}
prix_2027.update(PREMIX_PRICES)

# NOUVEAU: Prix COMPLEMENT ALIMENTAIRE (par unité = 1L)
# Basé sur les prix moyens historiques observés
# V305 EXCLU — seul V300 1L conservé
COMPLEMENT_PRICES = {
    'V300': 2500,      # BELGOKILL 1L (~2 500 FCFA/L)
    'CA003.1': 8000,   # BELGO HARMONY 1L
    'CA004.1': 9800,   # BELGO PROTECT 1L
    'CA006.1': 6500,   # BELGO DRY LIT 1L
    'CA001.1': 5500,   # BELGO WATER CLEAN 1L
    'CA002.1': 10500,  # BELGO VIT Ese 1L
    'CA005.1': 15000,  # BELGO THERMO 1L
    'CA007.1': 8000,   # BELGO BIO SELECT 1L
    'CA008.1': 14000,  # BELGO FRESH 1L
}
# Override with real 2026 prices (CORRIGE)
prix_reels_2026 = json.load(open("/home/z/my-project/scripts/prix_reels_2026.json"))
prix_2027.update(prix_reels_2026)
prix_2027['T102'] = 16550  # Mediane mensuelle YTD 2026 (Jan-Sep)

# Weight per unit (kg per sac/piece/bidon) — for converting tonnes to units
WEIGHT_MAP = {
    'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25,
    'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
    'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50, 'C1022': 5,
    'CB100': 25, 'CB200': 25, 'CB101': 5, 'CB201': 5, 'PB100': 25, 'PB200': 25, 'DB100': 25, 'DB200': 25, 'ALAP25': 25,
    'B100': 25, 'B1001': 1, 'B1003': 5, 'B1004': 25,
    'E101': 25, 'E1011': 1, 'E1013': 5, 'E1014': 0.2,
    'I105': 25, 'I1051': 1, 'I1053': 5, 'I1054': 25,
    'I106': 25, 'I1061': 1, 'I107': 25, 'I1071': 1,
    'P105': 25, 'P1051': 1, 'P1053': 5,
    'F114': 50, 'F1145': 50, 'F1146': 25, 'F1147': 1,
    'P102N2': 25, 'P104N2': 25, 'P109': 25,
    'PX101': 25, 'PX102': 25, 'PX103': 25, 'PX104': 25, 'PX105': 25,
    'V300': 1, 'CA003.1': 1, 'CA004.1': 1, 'CA006.1': 1, 'CA001.1': 1,
    'CA002.1': 1, 'CA005.1': 1, 'CA007.1': 1, 'CA008.1': 1,
    'MAT011-80010002': 1, 'MAT014-80010003': 1, 'MAT015': 1, 'MAT017': 1,
    'MAT003': 1, 'MAT004': 1, 'MAT005': 1, 'MAT006': 1, 'MAT007': 1, 'MAT008': 1, 'MAT009': 1,
    'MAT020': 1, 'MAT023': 1, 'MAT029': 1, 'MAT030': 1, 'MAT033': 1, 'MAT039': 1, 'MAT040': 1, 'MAT042': 1,
    'MAT049': 1, 'MAT050': 1, 'MAT053': 1, 'MAT054': 1, 'MAT055': 1, 'MAT073': 1,
    'ME100': 1, 'ME1001': 1, 'ME101': 1, 'ME102': 1, 'ME103': 1, 'ME104': 1, 'ME1041': 1, 'ME105': 1, 'ME106': 1, 'ME107': 1,
    'APCL2': 15,
    'APCL25': 25,
    'APCL3': 15,
    'APCL30': 1,
    'APCL35': 5,
    'APCL4.5': 15,
    'APCL4.55': 5,
    'APCL450': 1,
    'APCL6': 15,
    'APCL65': 5,
    'APCL8': 15,
    'APCL80': 1,
    'APCL85': 5,
    'APTOR3': 15,
    'APTOR31': 1,
    'APTOR35': 5,
    'APTOR4.5': 15,
    'APTOR4.55': 5,
    'APTSA2': 15,
    'APTSA21': 1,
    'APTSA25': 5,
    'F1142': 1,
    'F1143': 25,
    'I1063': 5,
    'I10711': 1,
    'MAT001': 1,
    'MAT026': 1,
    'MAT047': 1,
    'MAT048': 1,
    'MAT060': 1,
    'MAT096': 1,
    'P102': 25,
    'P104': 25,
    'P1040': 25,
}

def get_units_from_tonnes(ref, tonnes):
    """Convert tonnes to number of units based on product weight per unit."""
    weight = WEIGHT_MAP.get(ref, 50)
    return tonnes * 1000 / weight

print(f"  Prix soja T102: {prix_2027.get('T102', 'N/A')} FCFA/sac")
print(f"  Prix C104: {prix_2027.get('C104', 'N/A')} FCFA/sac")
print(f"  Prix V300 (BELGOKILL 1L): {prix_2027.get('V300', 'N/A')} FCFA/L")
print(f"  Prix CA003.1 (BELGO HARMONY): {prix_2027.get('CA003.1', 'N/A')} FCFA/L")

# === Agrégation mensuelle par famille × région ===
print("\n=== AGRÉGATION MENSUELLE ===")
df_all['year_month_dt'] = df_all['year_month'].dt.to_timestamp()

# === FILTRAGE 2024 POUR ALIMENT_COMPLET (anomalie circonstancielle) ===
# 2024 a subi un crash exceptionnel (328 t vs 709 t en 2023 et 692 t en 2025).
# L'utilisateur confirme que c'était circonstanciel et que le problème est corrigé.
# On exclut UNIQUEMENT 2024 (pas 2023) du fitting Prophet pour ALIMENT_COMPLET.
# Garder 2023 permet à Prophet de capter la saisonnalité naturelle (Q1/Q4 peaks),
# tout en évitant que le creux 2024 ne tire la tendance vers le bas.
# Prophet sur 2023+2025+2026 donne ~904 t en 2027 (= -2% vs 2026 LY 920 t),
# donc on applique en plus un facteur de croissance maîtrisé de +15% représentant
# la dynamique de reprise post-2024, pour atteindre ~1 040 t (+13% vs 2026).
ALIMENT_COMPLET_GROWTH_BOOST = 1.15  # +15% appliqué post-Prophet sur ALIMENT_COMPLET
mask_ac_2024 = (df_all['family'] == 'ALIMENT_COMPLET') & (df_all['date'].dt.year == 2024)
n_ac_2024 = mask_ac_2024.sum()
df_all = df_all[~mask_ac_2024].copy()
print(f"Filtrage ALIMENT_COMPLET 2024 (anomalie circonstancielle): {n_ac_2024} records retirés")
print(f"  Historique ALIMENT_COMPLET restant: 2023 + 2025 + 2026 (sans le creux 2024)")
print(f"  Facteur croissance post-Prophet: ×{ALIMENT_COMPLET_GROWTH_BOOST} (reprise post-2024)")

monthly_fr_vol = df_all.groupby(['family', 'region', 'year_month'])['tonnes'].sum().reset_index()
monthly_fr_vol['date'] = monthly_fr_vol['year_month'].dt.to_timestamp()

monthly_fr_ca = df_all.groupby(['family', 'region', 'year_month'])['montant_ttc'].sum().reset_index()
monthly_fr_ca['date'] = monthly_fr_ca['year_month'].dt.to_timestamp()
# For older records (LY_24) montant_ttc is 0 — use montant_ht as fallback
df_all_fallback = df_all.copy()
df_all_fallback['ca_combined'] = df_all_fallback['montant_ttc'].where(
    df_all_fallback['montant_ttc'] > 0, df_all_fallback['montant_ht']
)
monthly_fr_ca_fallback = df_all_fallback.groupby(['family', 'region', 'year_month'])['ca_combined'].sum().reset_index()
monthly_fr_ca_fallback['date'] = monthly_fr_ca_fallback['year_month'].dt.to_timestamp()
monthly_fr_ca_fallback['ca_m_fcfa'] = monthly_fr_ca_fallback['ca_combined'] / 1e6
# Use the fallback version (more complete)
monthly_fr_ca = monthly_fr_ca_fallback

# Parts historiques produit × agence
monthly_pa = df_all.groupby(['ref', 'agence', 'family', 'region']).agg(
    tonnes=('tonnes', 'sum'),
    ca_ttc=('montant_ttc', 'sum'),
    ca_ht=('montant_ht', 'sum'),
).reset_index()
# Combined CA: TTC or HT fallback
monthly_pa['ca_combined'] = monthly_pa['ca_ttc'].where(monthly_pa['ca_ttc'] > 0, monthly_pa['ca_ht'])
total_fr_ca = monthly_pa.groupby(['family', 'region'])['ca_combined'].transform('sum')
monthly_pa['share_ca'] = monthly_pa['ca_combined'] / total_fr_ca
total_fr_t = monthly_pa.groupby(['family', 'region'])['tonnes'].transform('sum')
monthly_pa['share_tonnes'] = monthly_pa['tonnes'] / total_fr_t.replace(0, 1)

combos_fr = df_all.groupby(['family', 'region']).size().reset_index()
combos_fr.columns = ['family', 'region', 'n_months']
print(f"  {len(combos_fr)} combinaisons famille × région")

# === Fonction Prophet ===
def fit_prophet_fast(history_df, periods=12, freq='MS'):
    """Fit Prophet and forecast 12 months starting from Jan 2027."""
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

def extrapolate_mean(history_df, periods=12, freq='MS', value_col='y', family=None):
    """Extrapolate using historical monthly average for 2027.
    For ALVEOLES: use 2026 monthly average (very low) instead of 2025 spike."""
    future_dates = pd.date_range(start='2027-01-01', periods=periods, freq=freq)
    
    # SPECIAL CASE: ALVEOLES — use 2026 monthly average (very low) instead of 2025 spike
    if family == 'ALVEOLES':
        # 2026 ALVEOLES: 11 records, 13.9 M HT over 8 months → ~1.7 M/month
        avg_2026 = history_df[history_df['ds'].dt.year == 2026][value_col].mean()
        if pd.isna(avg_2026) or avg_2026 == 0:
            avg_2026 = 1.5  # Fallback: ~1.5 M/month
        # Apply seasonal weights (Q4 slight uptick for year-end)
        weights = [0.8, 0.9, 1.0, 1.0, 1.1, 1.0, 0.9, 0.9, 0.9, 1.2, 1.1, 1.0]
        return pd.DataFrame({
            'ds': future_dates,
            'yhat': [avg_2026 * w for w in weights],
            'yhat_lower': [avg_2026 * w * 0.7 for w in weights],
            'yhat_upper': [avg_2026 * w * 1.3 for w in weights],
        })
    
    avg = history_df[value_col].mean() if len(history_df) > 0 else 0
    return pd.DataFrame({'ds': future_dates, 'yhat': [avg] * periods,
                         'yhat_lower': [avg * 0.7] * periods, 'yhat_upper': [avg * 1.3] * periods})

# === Forecast 2027 (12 mois) ===
print("\n=== FORECAST 2027 (12 mois, S3) ===")
forecasts_fr = []

# NOUVEAU: 6 familles avec Prophet (vs 5 avant), 2 familles avec extrapolation
# PREMIX passe en Prophet (avec volumes) — demande utilisateur 2026-09-04
PROPHET_FAMILIES = ['TOURTEAUX', 'CONCENTRES', 'INGREDIENTS', 'ALIMENT_COMPLET', 'COMPLEMENT_ALIMENTAIRE', 'PREMIX']  # MAIS excluded (opportuniste, forecast=0)
EXTRAPOL_FAMILIES = ['MATERIEL_ELEVAGE', 'ALVEOLES']  # MAIS excluded entirely (forecast=0)

# Exclude MAIS from forecast (opportuniste, forecast=0)
combos_fr = combos_fr[combos_fr['family'] != 'MAIS']

for idx, row in combos_fr.iterrows():
    family = row['family']
    region = row['region']
    
    if family in EXTRAPOL_FAMILIES:
        history = monthly_fr_ca[(monthly_fr_ca['family'] == family) & (monthly_fr_ca['region'] == region)].copy()
        history = history[['date', 'ca_m_fcfa']].rename(columns={'date': 'ds', 'ca_m_fcfa': 'y'})
        history = history.sort_values('ds')
        history['y'] = history['y'].clip(lower=0)
        forecast = extrapolate_mean(history, periods=12, freq='MS', value_col='y', family=family)
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

# === Désagrégation par produit × agence ===
print("\n=== DÉSAGRÉGATION PAR PRODUIT × AGENCE ===")
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
        # === Application du facteur de croissance maîtrisé pour ALIMENT_COMPLET ===
        # Représente la dynamique de reprise post-2024 (cf. bloc filtrage 2024)
        if family == 'ALIMENT_COMPLET':
            tonnes_total = tonnes_total * ALIMENT_COMPLET_GROWTH_BOOST
        # Calculate CA via prices
        pa_subset = monthly_pa[(monthly_pa['family'] == family) & (monthly_pa['region'] == region)]
        ca_total_m_fcfa = 0
        for _, pa_row in pa_subset.iterrows():
            ref = pa_row['ref']
            share_t = pa_row['share_tonnes']
            tonnes_combo = tonnes_total * share_t
            # For COMPLEMENT_ALIMENTAIRE, prix is per L (not per sac 50kg)
            if family == 'COMPLEMENT_ALIMENTAIRE':
                litres = tonnes_combo * 1000  # 1L = 1kg
                prix = prix_2027.get(ref, 0)
                ca_total_m_fcfa += litres * prix / 1e6
            else:
                units = get_units_from_tonnes(ref, tonnes_combo)
                prix = prix_2027.get(ref, 0)
                ca_total_m_fcfa += units * prix / 1e6
    
    pa_subset = monthly_pa[(monthly_pa['family'] == family) & (monthly_pa['region'] == region)]
    
    for _, pa_row in pa_subset.iterrows():
        ref = pa_row['ref']
        agence = pa_row['agence']
        share_ca = pa_row['share_ca']
        share_t = pa_row['share_tonnes']
        
        if is_ca_only:
            # Pour MATERIEL_ELEVAGE et ALVEOLES: pas de volume, CA distribué par share_ca historique
            ca_m_fcfa_combo = ca_total_m_fcfa * share_ca
            tonnes_combo = 0
            sacs_50 = 0
            prix = 0
        else:
            tonnes_combo = tonnes_total * share_t
            if family == 'COMPLEMENT_ALIMENTAIRE':
                # Liquide: 1L = 1kg
                litres = tonnes_combo * 1000
                sacs_50 = litres  # nb de litres
                prix = prix_2027.get(ref, 0)  # Prix par L
                # CORRIGÉ: CA calculé directement à partir du volume × prix (au lieu de share_ca)
                # Évite les CA=0 pour des agences sans historique de CA sur ce ref
                ca_m_fcfa_combo = litres * prix / 1e6
            else:
                # CORRIGÉ: utilisation de get_units_from_tonnes pour respecter le weight per unit
                # (au lieu de tonnes * 1000 / 50 qui supposait 50kg pour tous les produits)
                units = get_units_from_tonnes(ref, tonnes_combo)
                sacs_50 = units
                prix = prix_2027.get(ref, 0)
                # CORRIGÉ: CA calculé directement (units × prix) au lieu de share_ca
                ca_m_fcfa_combo = units * prix / 1e6
        
        all_forecasts.append({
            'scenario': 'S3_reappro_100',
            'ref': ref,
            'family': family,
            'agence': agence,
            'region': region,
            'date': date,
            'month': month,
            'year': year,
            'tonnes': round(tonnes_combo, 2),
            'sacs_50': round(sacs_50, 1),
            'prix_ttc_sac': prix,
            'ca_m_fcfa': round(ca_m_fcfa_combo, 2),
        })

fcst_df = pd.DataFrame(all_forecasts)
# Filter 2027 only
fcst_df = fcst_df[fcst_df['year'] == 2027]
print(f"\n{len(fcst_df)} forecasts détaillés générés (2027, 12 mois)")

# === BUNDLE 2.5:1 (soja:concentré) constraint ===
print("\n=== APPLICATION BUNDLE 2.5:1 (soja:concentré) ===")
BUNDLE_RATIO = 2.5
adjustments_made = 0
for (region, month), group in fcst_df.groupby(['region', 'month']):
    soja_t = group[group['family'] == 'TOURTEAUX']['tonnes'].sum()
    conc_t = group[group['family'] == 'CONCENTRES']['tonnes'].sum()
    if soja_t > 0 and conc_t > 0:
        ratio = soja_t / conc_t
        if ratio > BUNDLE_RATIO:
            # Increase CONCENTRÉS to reach 2.5:1
            target_conc = soja_t / BUNDLE_RATIO
            adjustment_factor = target_conc / conc_t
            mask = (fcst_df['region'] == region) & (fcst_df['month'] == month) & (fcst_df['family'] == 'CONCENTRES')
            fcst_df.loc[mask, 'tonnes'] = fcst_df.loc[mask, 'tonnes'] * adjustment_factor
            fcst_df.loc[mask, 'sacs_50'] = fcst_df.loc[mask, 'sacs_50'] * adjustment_factor
            # Recalculate CA for CONCENTRES (CORRIGE: units, not sacs_50)
            for idx in fcst_df[mask].index:
                ref = fcst_df.at[idx, 'ref']
                new_tonnes = fcst_df.at[idx, 'tonnes']
                units = get_units_from_tonnes(ref, new_tonnes)
                prix = prix_2027.get(ref, 0)
                fcst_df.at[idx, 'ca_m_fcfa'] = round(units * prix / 1e6, 2)
                fcst_df.at[idx, 'sacs_50'] = round(units, 1)
            adjustments_made += 1
            if adjustments_made <= 10:  # Show only first 10
                print(f"  {region} × {month}/2027: ratio {ratio:.2f} → {BUNDLE_RATIO}:1 (CONCENTRES ×{adjustment_factor:.2f})")

print(f"  Total adjustments: {adjustments_made}")

# === FORFAIT SPC (ALVEOLES + MATERIEL_ELEVAGE) ===
print("\n=== APPLICATION FORFAIT SPC (ALVEOLES + MAT_ELEVAGE) ===")
import sys
sys.path.insert(0, "/home/z/my-project/scripts")
from spc_forfait import generate_spc_forfait_2027

spc_forfait_records = generate_spc_forfait_2027()
spc_df = pd.DataFrame(spc_forfait_records)
print(f"  Forfait SPC: {len(spc_df)} records, CA total = {spc_df['ca_m_fcfa'].sum():.1f} M FCFA")
print(f"    - ALVEOLES: {spc_df[spc_df['family']=='ALVEOLES']['ca_m_fcfa'].sum():.1f} M (basé sur 2025)")
print(f"    - MAT_ELEVAGE: {spc_df[spc_df['family']=='MATERIEL_ELEVAGE']['ca_m_fcfa'].sum():.1f} M (basé sur 2025)")
print(f"    - SPC PK15 forfait réaliste: 0.5 M ALV + 0.5 M MAT (annuel)")

# Add forfait to forecast
fcst_df = pd.concat([fcst_df, spc_df], ignore_index=True)
print(f"  Total forecast après forfait: {len(fcst_df)} records, CA = {fcst_df['ca_m_fcfa'].sum():.1f} M FCFA")

# Save
output_path = "/home/z/my-project/scripts/forecast_2027_S3.csv"
fcst_df.to_csv(output_path, index=False)
print(f"Saved: {output_path}")

# === Synthèse ===
print("\n=== SYNTHÈSE FORECAST 2027 (S3) ===")
print("\n--- PAR FAMILLE ---")
synth_fam = fcst_df.groupby('family').agg(
    tonnes=('tonnes', 'sum'),
    ca_m_fcfa=('ca_m_fcfa', 'sum'),
).round({'tonnes': 0, 'ca_m_fcfa': 1})
total_ca = fcst_df['ca_m_fcfa'].sum()
synth_fam['part_ca_pct'] = (synth_fam['ca_m_fcfa'] / total_ca * 100).round(1)
print(synth_fam)

print("\n--- PAR TRIMESTRE ---")
fcst_df['quarter'] = ((fcst_df['month'] - 1) // 3) + 1
synth_q = fcst_df.groupby('quarter').agg(
    tonnes=('tonnes', 'sum'),
    ca_m_fcfa=('ca_m_fcfa', 'sum'),
).round({'tonnes': 0, 'ca_m_fcfa': 1})
synth_q.index = ['Q1 (Jan-Mar)', 'Q2 (Avr-Juin)', 'Q3 (Juil-Sept)', 'Q4 (Oct-Déc)']
print(synth_q)

print("\n--- PAR MOIS ---")
synth_month = fcst_df.groupby('month').agg(
    tonnes=('tonnes', 'sum'),
    ca_m_fcfa=('ca_m_fcfa', 'sum'),
).round({'tonnes': 0, 'ca_m_fcfa': 1})
month_names = ['Jan','Fév','Mar','Avr','Mai','Juin','Juil','Août','Sep','Oct','Nov','Déc']
synth_month.index = month_names
print(synth_month)

print("\n--- PAR RÉGION ---")
synth_reg = fcst_df.groupby('region').agg(
    tonnes=('tonnes', 'sum'),
    ca_m_fcfa=('ca_m_fcfa', 'sum'),
).round({'tonnes': 0, 'ca_m_fcfa': 1})
print(synth_reg)

print("\n--- TOTAL 2027 ---")
print(f"Volume: {fcst_df['tonnes'].sum():.0f} t")
print(f"CA: {fcst_df['ca_m_fcfa'].sum():.1f} M FCFA")
print(f"Nb produits: {fcst_df['ref'].nunique()}")
print(f"Nb agences: {fcst_df['agence'].nunique()}")
print(f"Nb familles: {fcst_df['family'].nunique()}")
print(f"Nb lignes: {len(fcst_df)}")

# Save synthesis as JSON for PDF generation
synth_json = {
    'total_tonnes': float(fcst_df['tonnes'].sum()),
    'total_ca_m_fcfa': float(fcst_df['ca_m_fcfa'].sum()),
    'n_products': int(fcst_df['ref'].nunique()),
    'n_agences': int(fcst_df['agence'].nunique()),
    'n_families': int(fcst_df['family'].nunique()),
    'n_lines': int(len(fcst_df)),
    'by_family': synth_fam.reset_index().to_dict(orient='records'),
    'by_quarter': synth_q.reset_index().rename(columns={'index': 'quarter'}).to_dict(orient='records'),
    'by_region': synth_reg.reset_index().to_dict(orient='records'),
    'by_month': synth_month.reset_index().rename(columns={'index': 'month'}).to_dict(orient='records'),
}
with open('/home/z/my-project/scripts/forecast_2027_synth.json', 'w') as f:
    json.dump(synth_json, f, indent=2, default=str)
print(f"\nSynthèse sauvée: /home/z/my-project/scripts/forecast_2027_synth.json")
