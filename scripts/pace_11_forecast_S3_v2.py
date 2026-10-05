"""
Forecast Q4 2026 (Sept-Dec) - Scenario S3 - VERSION 2 (mis à jour avec septembre réel)
Utilise les données 2023-2026 (44 mois) + COMPLEMENT_ALIMENTAIRE (V300 1L only)

Changements vs version 1:
1. Utilise dataset_2023_2026.csv (264 759 records, 44 mois) au lieu de dataset_consolide_v2.csv (115 086 records, 20 mois)
2. AJOUT de la famille COMPLEMENT_ALIMENTAIRE (V300 1L + CA001-CA008, 1L=1kg)
   - V305 (BELGOKILL 200L) EXCLU
3. MATERIEL_ELEVAGE toujours à 0 en tonnes
4. Désaisonnalisation de l'effet soja Jul-Août 2026 (cap moyenne S1 2026)
5. En cours + Validées septembre inclus (commandes ouvertes fin sept)
6. Prix soja actualisé 17 678 FCFA/sac (moyenne pondérée YTD Jan-Sep 2026)

Mise à jour septembre réel (fichier _upd):
  - Septembre Livrée (01-30/09) ajouté à l'historique Prophet
  - Les lignes septembre du forecast sont remplacées par les VENTES RÉELLES de septembre
    (plus de facteur 0,7 TOURTEAUX sur septembre) — Oct/Nov/Déc restent prévisionnels
  - Événement prix : hausse 04/09 (25 000 → 26 000), baisse 22/09 (26 000 → 20 000 FCFA/sac)

Méthode:
  - Prophet pour 5 familles (TOURTEAUX, CONCENTRES, INGRÉDIENTS, ALIMENT_COMPLET, COMPLEMENT_ALIMENTAIRE)
  - Extrapolation pour MATERIEL_ELEVAGE et PREMIX (CA only)
  - Forecast: Sep, Oct, Nov, Dec 2026 (4 mois) — Sept = réel
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

# === Load dataset 2023-2026 ===
df = pd.read_csv("/home/z/my-project/scripts/dataset_2023_2026.csv", parse_dates=['date'], low_memory=False)
print(f"Loaded {len(df)} records (Livrée only, 2023-2026)")
print(f"Date range: {df['date'].min().date()} → {df['date'].max().date()}")
print(f"\nBy family:")
print(df['family'].value_counts())

# === Load En cours + Validées from latest extraction ===
print("\n=== Loading En cours + Validées from latest extraction (septembre) ===")
import openpyxl

# Extraction finale septembre (01-30/09) — format 17 colonnes (54)/(56)/(61) : état=14, agence=16
FILE_AOUT = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (61).xlsx"
wb = openpyxl.load_workbook(FILE_AOUT, read_only=True, data_only=True)
ws = wb['Sheet 1']

SOJA_REFS = {'T102': 50, 'T1021': 1, 'T1023': 5, 'T1024': 25}
CONC_REFS = {'C101': 50, 'C102': 50, 'C103': 50, 'C104': 50, 'C1042': 1, 'C1043': 5, 'C1044': 25,
             'C105': 50, 'C1053': 1, 'C1054': 5, 'C1055': 25, 'C108': 50, 'C1022': 5}
ALIMENT_REFS = {'CB100': 25, 'CB200': 25, 'CB101': 5, 'CB201': 5, 'PB100': 25, 'PB200': 25,
                'DB100': 25, 'DB200': 25, 'ALAP25': 25,
                'APCL2': 15, 'APCL25': 5, 'APCL3': 15, 'APCL30': 1, 'APCL35': 5,
                'APCL4.5': 15, 'APCL4.55': 5, 'APCL450': 1, 'APCL6': 15, 'APCL65': 5,
                'APCL8': 15, 'APCL80': 1, 'APCL85': 5}  # BELGOFISH intégrés aux ALIMENTS COMPLETS
INGREDIENT_REFS = {'B100': 25, 'E101': 25, 'I105': 25, 'B1001': 1, 'B1003': 5, 'B1004': 25,
                   'E1011': 1, 'E1013': 5, 'E1014': 0.2, 'I1051': 1, 'I1053': 5, 'I1054': 25,
                   'I106': 25, 'I1061': 1, 'I107': 25, 'I1071': 1, 'P105': 25, 'P1051': 1, 'P1053': 5,
                   'F114': 50, 'F1145': 50, 'F1146': 25, 'F1147': 1,
                   'F1143': 25, 'I1063': 5}  # farine de poisson 25 kg + méthionine 5 kg
PREMIX_REFS = {'P102N2': 25, 'P104N2': 25, 'P109': 25, 'PX101': 25, 'PX102': 25, 'PX103': 25, 'PX104': 25, 'PX105': 25}
MATERIEL_REFS = {f'MAT{i:03d}': 1 for i in range(1, 100)}
MATERIEL_REFS.update({'MAT014-80010003': 1, 'MAT011-80010002': 1})
# COMPLEMENT ALIMENTAIRE — V300 1L only (V305 EXCLU)
COMPLEMENT_REFS = {
    'V300': 1,        # BELGOKILL 1L = 1 kg
    'CA003.1': 1, 'CA004.1': 1, 'CA006.1': 1, 'CA001.1': 1,
    'CA002.1': 1, 'CA005.1': 1, 'CA007.1': 1, 'CA008.1': 1,
}

ALL_REFS = {**SOJA_REFS, **CONC_REFS, **ALIMENT_REFS, **INGREDIENT_REFS, 
            **PREMIX_REFS, **MATERIEL_REFS, **COMPLEMENT_REFS}

# ALVEOLES refs (separated from MATERIEL_ELEVAGE)
ALVEOLES_REFS = {'MAT011-80010002', 'MAT014-80010003', 'MAT015', 'MAT017'}

def get_family_q4(ref):
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
    family = get_family_q4(ref)
    if family is None: continue
    etat = str(r[14]).strip() if r[14] else ''
    if etat not in ('En cours', 'Validée'): continue
    date_str = str(r[6])[:10] if r[6] else ''
    if '/09/2026' not in date_str: continue
    agence_raw = r[16] if r[16] else ''
    if agence_raw not in AGENCE_MAP: continue
    agence, region = AGENCE_MAP[agence_raw]
    qte = r[2] or 0
    weight = ALL_REFS.get(ref, 1)
    kg = qte * weight
    montant_ttc = r[9] or 0
    montant_ht = r[8] or 0
    # Règle : une vente dont le montant HT ou TTC est à 0 n'est pas intégrée
    if not isinstance(montant_ht, (int, float)) or montant_ht <= 0 \
       or not isinstance(montant_ttc, (int, float)) or montant_ttc == 0: continue
    try:
        date = pd.to_datetime(date_str, format='%d/%m/%Y')
    except:
        continue

    # For MATERIEL_ELEVAGE, ALVEOLES, PREMIX: tonnes=0 (CA only)
    if family in ('MATERIEL_ELEVAGE', 'ALVEOLES'):
        tonnes_val = 0
    elif family == 'COMPLEMENT_ALIMENTAIRE':
        tonnes_val = kg / 1000  # 1L = 1kg conversion
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

# === Charger septembre LIVRÉE (ventes réelles 01-30/09) ===
print("\n=== Loading Septembre Livrée (ventes réelles) ===")
sep_records = []
for r in ws.iter_rows(min_row=3, values_only=True):
    if not r or not r[0] or r[0] == 'Total': continue
    ref = str(r[0])
    family = get_family_q4(ref)
    if family is None: continue
    etat = str(r[14]).strip() if r[14] else ''
    if etat != 'Livrée': continue
    date_str = str(r[6])[:10] if r[6] else ''
    if '/09/2026' not in date_str: continue
    agence_raw = r[16] if r[16] else ''
    if agence_raw not in AGENCE_MAP: continue
    agence, region = AGENCE_MAP[agence_raw]
    qte = r[2] or 0
    weight = ALL_REFS.get(ref, 1)
    kg = qte * weight
    montant_ttc = r[9] or 0
    montant_ht = r[8] or 0
    # Règle : une vente dont le montant HT ou TTC est à 0 n'est pas intégrée
    if not isinstance(montant_ht, (int, float)) or montant_ht <= 0 \
       or not isinstance(montant_ttc, (int, float)) or montant_ttc == 0: continue
    try:
        date = pd.to_datetime(date_str, format='%d/%m/%Y')
    except:
        continue
    if family in ('MATERIEL_ELEVAGE', 'ALVEOLES'):
        tonnes_val = 0
    elif family == 'COMPLEMENT_ALIMENTAIRE':
        tonnes_val = kg / 1000
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
sep_vol_t = df_sep['tonnes'].sum()
sep_ca_m = df_sep['montant_ttc'].sum() / 1e6
print(f"  Septembre réel: {sep_vol_t:.0f} t, CA {sep_ca_m:.0f} M FCFA (14 agences)")

# Merge
df_all = pd.concat([df, df_extra, df_sep], ignore_index=True)
print(f"\nDataset total: {len(df_all)} records")
print(f"Date range: {df_all['date'].min().date()} → {df_all['date'].max().date()}")

# === Désaisonnalisation de l'effet soja exceptionnel ===
print("\n=== DÉSAISONNALISATION EFFET SOJA EXCEPTIONNEL ===")
df_all['year_month'] = df_all['date'].dt.to_period('M')

soja_s1 = df_all[(df_all['family'] == 'TOURTEAUX') & (df_all['date'] < '2026-07-01') & (df_all['date'] >= '2026-01-01')]
soja_s1_monthly = soja_s1.groupby('year_month')['tonnes'].sum()
soja_s1_avg = soja_s1_monthly.mean()
print(f"  S1 2026 average monthly soja: {soja_s1_avg:.0f} t")

soja_juil_aout = df_all[(df_all['family'] == 'TOURTEAUX') & 
                         (df_all['date'] >= '2026-07-01') & 
                         (df_all['date'] <= '2026-08-31')]
soja_juil_aout_monthly = soja_juil_aout.groupby('year_month')['tonnes'].sum()
print(f"  Jul-Aug 2026 soja monthly volumes:")
for ym, t in soja_juil_aout_monthly.items():
    print(f"    {ym}: {t:.0f} t")

cap_factor = soja_s1_avg / soja_juil_aout_monthly.max() if soja_juil_aout_monthly.max() > 0 else 1.0
print(f"  Cap factor: {cap_factor:.2f}")

mask_cap = (df_all['family'] == 'TOURTEAUX') & (df_all['date'] >= '2026-07-01') & (df_all['date'] <= '2026-08-31')
df_all.loc[mask_cap, 'tonnes'] = df_all.loc[mask_cap, 'tonnes'] * cap_factor
df_all.loc[mask_cap, 'sacs_50'] = df_all.loc[mask_cap, 'sacs_50'] * cap_factor
df_all.loc[mask_cap, 'kg'] = df_all.loc[mask_cap, 'kg'] * cap_factor

# === Prix Q4 2026 — CORRIGE avec prix reels 2026 ===
print("\n=== PRIX Q4 2026 (corriges avec prix reels 2026) ===")
# Load real 2026 prices per unit (par sac/piece/bidon)
prix_reels = json.load(open("/home/z/my-project/scripts/prix_reels_2026.json"))
# Load original forecast prices (for 50kg products)
prix_forecast = json.load(open("/home/z/my-project/scripts/prix_forecast.json"))
prix_q4 = prix_forecast['stable_Q4'].copy()
# Override with real 2026 prices (per unit) for non-50kg products
prix_q4.update(prix_reels)
# Update soja T102 price to 17 869 FCFA/sac (moyenne pondérée YTD Jan-Sep, réel ERP (61), incl. hausse septembre)
prix_q4['T102'] = 17869  # Moyenne ponderee YTD 2026 (Jan-Sep, reel ERP (61) 01-30/09)
# === Weight per unit (kg per sac/piece/bidon) — for converting tonnes to units ===
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
    weight = WEIGHT_MAP.get(ref, 50)  # default 50kg
    units = tonnes * 1000 / weight
    return units

# === Agrégation mensuelle ===
print("\n=== AGRÉGATION MENSUELLE ===")
df_all['year_month_dt'] = df_all['year_month'].dt.to_timestamp()

# === FILTRAGE 2024 POUR ALIMENT_COMPLET (anomalie circonstancielle) ===
# 2024 a subi un crash exceptionnel (328 t vs 709 t en 2023 et 692 t en 2025).
# L'utilisateur confirme que c'était circonstanciel et que le problème est corrigé.
# On exclut UNIQUEMENT 2024 du fitting Prophet pour ALIMENT_COMPLET.
# Un facteur de croissance maîtrisé de +15% est appliqué post-Prophet pour
# représenter la dynamique de reprise post-2024 (cf. bloc désagrégation).
ALIMENT_COMPLET_GROWTH_BOOST = 1.15
mask_ac_2024 = (df_all['family'] == 'ALIMENT_COMPLET') & (df_all['date'].dt.year == 2024)
n_ac_2024 = mask_ac_2024.sum()
df_all = df_all[~mask_ac_2024].copy()
print(f"Filtrage ALIMENT_COMPLET 2024 (anomalie circonstancielle): {n_ac_2024} records retirés")

monthly_fr_vol = df_all.groupby(['family', 'region', 'year_month'])['tonnes'].sum().reset_index()
monthly_fr_vol['date'] = monthly_fr_vol['year_month'].dt.to_timestamp()

# Combined CA (TTC fallback to HT for older records)
df_all_fallback = df_all.copy()
df_all_fallback['ca_combined'] = df_all_fallback['montant_ttc'].where(
    df_all_fallback['montant_ttc'] > 0, df_all_fallback['montant_ht']
)
monthly_fr_ca = df_all_fallback.groupby(['family', 'region', 'year_month'])['ca_combined'].sum().reset_index()
monthly_fr_ca['date'] = monthly_fr_ca['year_month'].dt.to_timestamp()
monthly_fr_ca['ca_m_fcfa'] = monthly_fr_ca['ca_combined'] / 1e6

# Parts historiques produit × agence
monthly_pa = df_all.groupby(['ref', 'agence', 'family', 'region']).agg(
    tonnes=('tonnes', 'sum'),
    ca_ttc=('montant_ttc', 'sum'),
    ca_ht=('montant_ht', 'sum'),
).reset_index()
monthly_pa['ca_combined'] = monthly_pa['ca_ttc'].where(monthly_pa['ca_ttc'] > 0, monthly_pa['ca_ht'])
total_fr_ca = monthly_pa.groupby(['family', 'region'])['ca_combined'].transform('sum')
monthly_pa['share_ca'] = monthly_pa['ca_combined'] / total_fr_ca
total_fr_t = monthly_pa.groupby(['family', 'region'])['tonnes'].transform('sum')
monthly_pa['share_tonnes'] = monthly_pa['tonnes'] / total_fr_t.replace(0, 1)

combos_fr = df_all.groupby(['family', 'region']).size().reset_index()
combos_fr.columns = ['family', 'region', 'n_months']
print(f"  {len(combos_fr)} combinaisons famille × région")

# === Prophet functions ===
def fit_prophet_fast_q4(history_df, periods=4, freq='MS'):
    """Fit Prophet and forecast Q4 2026 (Sep, Oct, Nov, Dec)."""
    future_dates = pd.date_range(start='2026-09-01', periods=periods, freq=freq)
    if len(history_df) < 4:
        # Use Q4 historical average if available
        q4_history = history_df[history_df['ds'].dt.month.isin([9, 10, 11, 12])]
        if len(q4_history) >= 4:
            avg = q4_history['y'].mean()
        else:
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
        q4_history = history_df[history_df['ds'].dt.month.isin([9, 10, 11, 12])]
        avg = q4_history['y'].mean() if len(q4_history) > 0 else history_df['y'].mean()
        if pd.isna(avg): avg = 0
        return pd.DataFrame({'ds': future_dates, 'yhat': [avg] * periods,
                             'yhat_lower': [avg * 0.8] * periods, 'yhat_upper': [avg * 1.2] * periods})

def extrapolate_mean_q4(history_df, periods=4, freq='MS', value_col='y', family=None):
    """Extrapolate using Q4 historical average.
    For ALVEOLES: use 2026 patterns (low) instead of 2025 spike."""
    future_dates = pd.date_range(start='2026-09-01', periods=periods, freq=freq)
    
    # SPECIAL CASE: ALVEOLES — use 2026 monthly average (very low) instead of 2025 spike
    if family == 'ALVEOLES':
        # 2026 ALVEOLES: 11 records, 13.9 M HT over 8 months → ~1.7 M/month
        # Spread across Q4 (Sep, Oct, Nov, Dec)
        avg_2026 = history_df[history_df['ds'].dt.year == 2026][value_col].mean()
        if pd.isna(avg_2026) or avg_2026 == 0:
            avg_2026 = 1.5  # Fallback: ~1.5 M/month
        # Apply Q4 weight (slightly higher than annual avg due to year-end purchases)
        q4_weights = [0.8, 1.2, 1.0, 0.9]  # Sep, Oct, Nov, Dec
        return pd.DataFrame({
            'ds': future_dates,
            'yhat': [avg_2026 * w for w in q4_weights],
            'yhat_lower': [avg_2026 * w * 0.7 for w in q4_weights],
            'yhat_upper': [avg_2026 * w * 1.3 for w in q4_weights],
        })
    
    q4_history = history_df[history_df['ds'].dt.month.isin([9, 10, 11, 12])]
    if len(q4_history) >= 4:
        avg = q4_history[value_col].mean()
    else:
        avg = history_df[value_col].mean() if len(history_df) > 0 else 0
    if pd.isna(avg): avg = 0
    return pd.DataFrame({'ds': future_dates, 'yhat': [avg] * periods,
                         'yhat_lower': [avg * 0.7] * periods, 'yhat_upper': [avg * 1.3] * periods})

# === Forecast Q4 2026 ===
print("\n=== FORECAST Q4 2026 (S3) ===")
forecasts_fr = []

# NOUVEAU: 6 familles avec Prophet (PREMIX ajouté avec volumes), 2 avec extrapolation
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
        forecast = extrapolate_mean_q4(history, periods=4, freq='MS', value_col='y', family=family)
        is_ca_only = True
    else:
        history = monthly_fr_vol[(monthly_fr_vol['family'] == family) & (monthly_fr_vol['region'] == region)].copy()
        history = history[['date', 'tonnes']].rename(columns={'date': 'ds', 'tonnes': 'y'})
        history = history.sort_values('ds')
        history['y'] = history['y'].clip(lower=0)
        forecast = fit_prophet_fast_q4(history, periods=4, freq='MS')
        is_ca_only = False
    
    forecast['yhat'] = forecast['yhat'].clip(lower=0)
    forecast['family'] = family
    forecast['region'] = region
    forecast['is_ca_only'] = is_ca_only
    
    forecasts_fr.append(forecast)
    val_sum = forecast['yhat'].sum()
    unit = 'M FCFA' if is_ca_only else 't'
    print(f"  ✓ {family} × {region}: forecast Q4 {val_sum:.1f} {unit}")

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
        # Septembre sera remplacé par les ventes réelles après désagrégation (plus de facteur 0,7)
        # === Facteur croissance maîtrisé pour ALIMENT_COMPLET (reprise post-2024) ===
        if family == 'ALIMENT_COMPLET':
            tonnes_total *= ALIMENT_COMPLET_GROWTH_BOOST
        # Calculate CA via real prices per unit (CORRIGE)
        pa_subset = monthly_pa[(monthly_pa['family'] == family) & (monthly_pa['region'] == region)]
        ca_total_m_fcfa = 0
        for _, pa_row in pa_subset.iterrows():
            ref = pa_row['ref']
            share_t = pa_row['share_tonnes']
            tonnes_combo = tonnes_total * share_t
            # CORRIGE: units = tonnes * 1000 / weight_per_unit; CA = units * prix_unit
            units = get_units_from_tonnes(ref, tonnes_combo)
            prix = prix_q4.get(ref, 0)
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
            units = 0
            prix = 0
        else:
            tonnes_combo = tonnes_total * share_t
            units = get_units_from_tonnes(ref, tonnes_combo)
            prix = prix_q4.get(ref, 0)
            # CORRIGÉ: CA calculé directement (units × prix) au lieu de share_ca
            # Évite les CA=0 pour des agences sans historique de CA sur ce ref
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
            'sacs_50': round(units, 1),  # Now contains units (not sacs 50kg)
            'prix_ttc_sac': prix,  # Now contains prix per unit
            'ca_m_fcfa': round(ca_m_fcfa_combo, 2),
        })

fcst_df = pd.DataFrame(all_forecasts)
# Filter Q4 2026 only
fcst_df = fcst_df[(fcst_df['year'] == 2026) & (fcst_df['month'].isin([9, 10, 11, 12]))]
print(f"\n{len(fcst_df)} forecasts détaillés générés (Q4 2026 S3)")

# === REMPLACER SEPTEMBRE PAR LES VENTES RÉELLES ===
print("\n=== SEPTEMBRE = VENTES RÉELLES (remplacement du forecast) ===")
fcst_df = fcst_df[fcst_df['month'] != 9].copy()

sep_act = df_sep.groupby(['ref', 'family', 'agence', 'region']).agg(
    tonnes=('tonnes', 'sum'),
    montant_ttc=('montant_ttc', 'sum'),
    qte=('qte', 'sum'),
).reset_index()

sep_act_rows = []
for _, row in sep_act.iterrows():
    ref = row['ref']
    weight = ALL_REFS.get(ref, 1)
    units = row['tonnes'] * 1000 / weight if weight > 0 else 0
    prix_act = row['montant_ttc'] / units if units > 0 else 0
    sep_act_rows.append({
        'scenario': 'S3_reappro_100',
        'ref': ref,
        'family': row['family'],
        'agence': row['agence'],
        'region': row['region'],
        'date': pd.Timestamp('2026-09-01'),
        'month': 9,
        'year': 2026,
        'tonnes': round(row['tonnes'], 2),
        'sacs_50': round(units, 1),
        'prix_ttc_sac': round(prix_act, 0),
        'ca_m_fcfa': round(row['montant_ttc'] / 1e6, 2),
    })

sep_act_df = pd.DataFrame(sep_act_rows)
print(f"  Septembre réel intégré: {len(sep_act_df)} lignes, "
      f"{sep_act_df['tonnes'].sum():.0f} t, CA {sep_act_df['ca_m_fcfa'].sum():.1f} M FCFA")
fcst_df = pd.concat([fcst_df, sep_act_df], ignore_index=True)

# === BUNDLE 2.5:1 (soja:concentré) constraint ===
# For each region × month: ensure ratio soja/concentré <= 2.5
# If ratio > 2.5, increase CONCENTRÉS to match (upward adjustment)
# Septembre (mois 9) est exclu : ventes réelles, déjà à 2,5:1
print("\n=== APPLICATION BUNDLE 2.5:1 (soja:concentré) ===")
BUNDLE_RATIO = 2.5
adjustments_made = 0
for (region, month), group in fcst_df.groupby(['region', 'month']):
    if month == 9:
        continue  # Septembre = réel, ne pas ajuster
    soja_t = group[group['family'] == 'TOURTEAUX']['tonnes'].sum()
    conc_t = group[group['family'] == 'CONCENTRES']['tonnes'].sum()
    if soja_t > 0 and conc_t > 0:
        ratio = soja_t / conc_t
        if ratio > BUNDLE_RATIO:
            # Increase CONCENTRÉS to reach 2.5:1
            target_conc = soja_t / BUNDLE_RATIO
            adjustment_factor = target_conc / conc_t
            # Apply to all CONCENTRES rows for this region × month
            mask = (fcst_df['region'] == region) & (fcst_df['month'] == month) & (fcst_df['family'] == 'CONCENTRES')
            fcst_df.loc[mask, 'tonnes'] = fcst_df.loc[mask, 'tonnes'] * adjustment_factor
            fcst_df.loc[mask, 'sacs_50'] = fcst_df.loc[mask, 'sacs_50'] * adjustment_factor
            # Recalculate CA for CONCENTRES (CORRIGE: units, not sacs_50)
            for idx in fcst_df[mask].index:
                ref = fcst_df.at[idx, 'ref']
                new_tonnes = fcst_df.at[idx, 'tonnes']
                units = get_units_from_tonnes(ref, new_tonnes)
                prix = prix_q4.get(ref, 0)
                fcst_df.at[idx, 'ca_m_fcfa'] = round(units * prix / 1e6, 2)
                fcst_df.at[idx, 'sacs_50'] = round(units, 1)
            adjustments_made += 1
            print(f"  {region} × {month}/2026: ratio {ratio:.2f} → {BUNDLE_RATIO}:1 (CONCENTRES ×{adjustment_factor:.2f})")

print(f"  Total adjustments: {adjustments_made}")

# === FORFAIT SPC (ALVEOLES + MATERIEL_ELEVAGE) ===
print("\n=== APPLICATION FORFAIT SPC (ALVEOLES + MAT_ELEVAGE) ===")
import sys
sys.path.insert(0, "/home/z/my-project/scripts")
from spc_forfait import generate_spc_forfait_q4_2026

spc_forfait_records = generate_spc_forfait_q4_2026()
spc_df = pd.DataFrame(spc_forfait_records)
# Septembre = ventes réelles : retirer le forfait SPC de septembre (éviter le double comptage)
spc_df = spc_df[spc_df['month'] != 9].copy()
print(f"  Forfait SPC: {len(spc_df)} records, CA total = {spc_df['ca_m_fcfa'].sum():.1f} M FCFA")
print(f"    - ALVEOLES: {spc_df[spc_df['family']=='ALVEOLES']['ca_m_fcfa'].sum():.1f} M (basé sur 2025)")
print(f"    - MAT_ELEVAGE: {spc_df[spc_df['family']=='MATERIEL_ELEVAGE']['ca_m_fcfa'].sum():.1f} M (basé sur 2025)")
print(f"    - SPC PK15 forfait réaliste: 0.5 M ALV + 0.5 M MAT")

# Add forfait to forecast
fcst_df = pd.concat([fcst_df, spc_df], ignore_index=True)
print(f"  Total forecast après forfait: {len(fcst_df)} records, CA = {fcst_df['ca_m_fcfa'].sum():.1f} M FCFA")

# Save (version _upd : les fichiers originaux S3 restent intacts)
output_path = "/home/z/my-project/scripts/forecast_q4_2026_S3_upd.csv"
fcst_df.to_csv(output_path, index=False)
print(f"Saved: {output_path}")

# === Synthèse ===
print("\n=== SYNTHÈSE FORECAST Q4 2026 (S3) ===")
print("\n--- PAR FAMILLE ---")
synth_fam = fcst_df.groupby('family').agg(
    tonnes=('tonnes', 'sum'),
    ca_m_fcfa=('ca_m_fcfa', 'sum'),
).round({'tonnes': 0, 'ca_m_fcfa': 1})
total_ca = fcst_df['ca_m_fcfa'].sum()
synth_fam['part_ca_pct'] = (synth_fam['ca_m_fcfa'] / total_ca * 100).round(1)
print(synth_fam)

print("\n--- PAR MOIS ---")
synth_month = fcst_df.groupby('month').agg(
    tonnes=('tonnes', 'sum'),
    ca_m_fcfa=('ca_m_fcfa', 'sum'),
).round({'tonnes': 0, 'ca_m_fcfa': 1})
month_names = ['Sep', 'Oct', 'Nov', 'Déc']
synth_month.index = [month_names[m-9] for m in synth_month.index]
print(synth_month)

print("\n--- PAR RÉGION ---")
synth_reg = fcst_df.groupby('region').agg(
    tonnes=('tonnes', 'sum'),
    ca_m_fcfa=('ca_m_fcfa', 'sum'),
).round({'tonnes': 0, 'ca_m_fcfa': 1})
print(synth_reg)

print("\n--- TOTAL Q4 2026 ---")
print(f"Volume: {fcst_df['tonnes'].sum():.0f} t")
print(f"CA: {fcst_df['ca_m_fcfa'].sum():.1f} M FCFA")
print(f"Nb produits: {fcst_df['ref'].nunique()}")
print(f"Nb agences: {fcst_df['agence'].nunique()}")
print(f"Nb familles: {fcst_df['family'].nunique()}")
print(f"Nb lignes: {len(fcst_df)}")

# Save synthese JSON
synth_json = {
    'total_tonnes': float(fcst_df['tonnes'].sum()),
    'total_ca_m_fcfa': float(fcst_df['ca_m_fcfa'].sum()),
    'n_products': int(fcst_df['ref'].nunique()),
    'n_agences': int(fcst_df['agence'].nunique()),
    'n_families': int(fcst_df['family'].nunique()),
    'n_lines': int(len(fcst_df)),
    'by_family': synth_fam.reset_index().to_dict(orient='records'),
    'by_region': synth_reg.reset_index().to_dict(orient='records'),
    'by_month': synth_month.reset_index().to_dict(orient='records'),
}
with open('/home/z/my-project/scripts/forecast_q4_2026_synth_upd.json', 'w') as f:
    json.dump(synth_json, f, indent=2, default=str)
print(f"\nSynthèse sauvée: /home/z/my-project/scripts/forecast_q4_2026_synth_upd.json")
