"""
Forfait SPC v2 (révisé) pour ALVEOLES + MATERIEL_ELEVAGE.

METHODE: 
- Poids par agence = basé sur 2025 (relatif, en %)
- Forfait total = basé sur la saisonnalité 2026 annualisée

Résultats 2026 (YTD 8 mois, SPC agencies):
- ALVEOLES: 0 M FCFA en 2026 → forfait ALV annuel = 0 M
- MATERIEL_ELEVAGE: 25.63 M en 2026 (8 mois) → annualisé = 38.44 M

Pour les poids par agence (2025):
- SPC Baf-Chefferie: 79.79% (dominant)
- SPC Yassa: 5.98%
- SPC Kye-Ossi: 5.71%
- SPC DLA-Beri: 4.19%
- SPC Ndere: 2.53%
- SPC Village: 0.91%
- SPC Dschang: 0.71%
- SPC Buea: 0.18%
- SPC PK15: 0% (forfait réaliste 0.5 M/an)
- SPC TPO: 0.01%

Saisonnalité mensuelle (2025 weights, puisque 2026 n'a que 3 mois de data):
  Mai: 2.68%, Juin: 18.05%, Juil: 21.0%, Août: 28.39%, Sep: 12.82%, Oct: 10.47%, Nov: 3.58%, Déc: 3.02%
"""
import pandas as pd
import json

# 2025 SPC weights (relatifs, en %)
SPC_WEIGHTS_2025 = {
    'SPC Baf-Chefferie': {'ALVEOLES': 97.83, 'MATERIEL_ELEVAGE': 79.79},
    'SPC TPO':           {'ALVEOLES': 0.90,  'MATERIEL_ELEVAGE': 0.01},
    'SPC Ndere':         {'ALVEOLES': 0.64,  'MATERIEL_ELEVAGE': 2.53},
    'SPC Kye-Ossi':      {'ALVEOLES': 0.13,  'MATERIEL_ELEVAGE': 5.71},
    'SPC DLA-Beri':      {'ALVEOLES': 0.20,  'MATERIEL_ELEVAGE': 4.19},
    'SPC Yassa':         {'ALVEOLES': 0.03,  'MATERIEL_ELEVAGE': 5.98},
    'SPC Village':       {'ALVEOLES': 0.11,  'MATERIEL_ELEVAGE': 0.91},
    'SPC Dschang':       {'ALVEOLES': 0.12,  'MATERIEL_ELEVAGE': 0.71},
    'SPC Buea':          {'ALVEOLES': 0.05,  'MATERIEL_ELEVAGE': 0.18},
}

SPC_REGIONS = {
    'SPC Baf-Chefferie': 'Ouest',
    'SPC TPO': 'Littoral',
    'SPC Ndere': 'Centre',
    'SPC Kye-Ossi': 'Centre',
    'SPC DLA-Beri': 'Littoral',
    'SPC Yassa': 'Littoral',
    'SPC Village': 'Littoral',
    'SPC Dschang': 'Ouest',
    'SPC Buea': 'Littoral',
    'SPC PK15': 'Littoral',
}

# FORFAIT ANNUEL basé sur 2026 annualisé (NOUVELLE METHODE v2)
FORFAIT_ALVEOLES_ANNUAL_2026 = 0.0  # ALV: 0 M en 2026 (pic 2025 non récurrent)
FORFAIT_MAT_ELEVAGE_ANNUAL_2026 = 38.44  # MAT: 25.63 M YTD × 12/8 = 38.44 M annualisé

# SPC PK15: 0 en 2025 ET 0.06 M en 2026 (1 seule vente) → forfait réaliste
FORFAIT_PK15_ALVEOLES = 0.0  # 0 ALV en 2026
FORFAIT_PK15_MAT_ELEVAGE = 0.5  # Forfait réaliste (taille similaire à SPC Yassa/Village)

# Saisonnalité mensuelle (basée sur 2025 SPC MAT_ELEVAGE, car 2026 n'a que 3 mois)
# 2025 weights: Mai 2.68%, Juin 18.05%, Juil 21.0%, Août 28.39%, Sep 12.82%, Oct 10.47%, Nov 3.58%, Déc 3.02%
# Complété pour Q1 (qui était 0 en 2025) avec une légère présence (~5% par mois)
MONTH_WEIGHTS_ANNUAL = {
    1: 0.04, 2: 0.04, 3: 0.04,    # Q1: 12% (extrapolé)
    4: 0.04, 5: 0.027, 6: 0.180,  # Q2: 24.7%
    7: 0.210, 8: 0.284, 9: 0.128, # Q3: 62.2% (pic été)
    10: 0.105, 11: 0.036, 12: 0.030  # Q4: 17.1%
}

# Q4 monthly weights (specific to Q4 2026 forecast)
MONTH_WEIGHTS_Q4 = {9: 0.128, 10: 0.105, 11: 0.036, 12: 0.030}  # Sep, Oct, Nov, Dec


def generate_spc_forfait_q4_2026():
    """Generate SPC forfait records for Q4 2026 (Sep, Oct, Nov, Dec).
    
    Forfait ALV: 0 (2026 = 0)
    Forfait MAT: based on 2026 annualized (38.44 M) × Q4 weights
    """
    records = []
    
    # Q4 portion of annual forfait (based on 2025 monthly weights for Q4)
    q4_total_weight = sum(MONTH_WEIGHTS_Q4.values())  # ~30%
    forfait_mat_q4 = FORFAIT_MAT_ELEVAGE_ANNUAL_2026 * q4_total_weight  # ~11.5 M
    forfait_alv_q4 = FORFAIT_ALVEOLES_ANNUAL_2026 * q4_total_weight  # 0
    
    # For each SPC agency, distribute forfait by 2025 weight
    for agence, weights in SPC_WEIGHTS_2025.items():
        region = SPC_REGIONS[agence]
        mat_share_pct = weights['MATERIEL_ELEVAGE'] / 100
        alv_share_pct = weights['ALVEOLES'] / 100  # Will be 0 anyway
        
        mat_ca_agence = forfait_mat_q4 * mat_share_pct
        alv_ca_agence = forfait_alv_q4 * alv_share_pct  # = 0
        
        for month, weight_m in MONTH_WEIGHTS_Q4.items():
            norm_weight = weight_m / q4_total_weight
            
            mat_month = mat_ca_agence * norm_weight
            alv_month = alv_ca_agence * norm_weight  # = 0
            
            # ALVEOLES record (always 0 in v2)
            if alv_month > 0.001:
                records.append({
                    'scenario': 'S3_reappro_100',
                    'ref': 'MAT014-80010003',
                    'family': 'ALVEOLES',
                    'agence': agence,
                    'region': region,
                    'date': pd.Timestamp(f'2026-{month:02d}-01'),
                    'month': month, 'year': 2026,
                    'tonnes': 0,
                    'sacs_50': 0,
                    'prix_ttc_sac': 0,
                    'ca_m_fcfa': round(alv_month, 2),
                    'source': 'forfait_spc_2026_saisonnalite'
                })
            
            # MATERIEL_ELEVAGE record
            if mat_month > 0.001:
                records.append({
                    'scenario': 'S3_reappro_100',
                    'ref': 'MAT003',
                    'family': 'MATERIEL_ELEVAGE',
                    'agence': agence,
                    'region': region,
                    'date': pd.Timestamp(f'2026-{month:02d}-01'),
                    'month': month, 'year': 2026,
                    'tonnes': 0,
                    'sacs_50': 0,
                    'prix_ttc_sac': 0,
                    'ca_m_fcfa': round(mat_month, 2),
                    'source': 'forfait_spc_2026_saisonnalite'
                })
    
    # SPC PK15 forfait réaliste (Q4 portion)
    pk15_mat_q4 = FORFAIT_PK15_MAT_ELEVAGE * q4_total_weight
    pk15_alv_q4 = FORFAIT_PK15_ALVEOLES * q4_total_weight  # 0
    
    for month, weight_m in MONTH_WEIGHTS_Q4.items():
        norm_weight = weight_m / q4_total_weight
        mat_month = pk15_mat_q4 * norm_weight
        alv_month = pk15_alv_q4 * norm_weight  # 0
        
        # PK15 ALV (0)
        if alv_month > 0.001:
            records.append({
                'scenario': 'S3_reappro_100',
                'ref': 'MAT014-80010003',
                'family': 'ALVEOLES',
                'agence': 'SPC PK15',
                'region': 'Littoral',
                'date': pd.Timestamp(f'2026-{month:02d}-01'),
                'month': month, 'year': 2026,
                'tonnes': 0, 'sacs_50': 0, 'prix_ttc_sac': 0,
                'ca_m_fcfa': round(alv_month, 2),
                'source': 'forfait_pk15_realiste'
            })
        
        # PK15 MAT (forfait réaliste 0.5 M/an)
        if mat_month > 0.001:
            records.append({
                'scenario': 'S3_reappro_100',
                'ref': 'MAT003',
                'family': 'MATERIEL_ELEVAGE',
                'agence': 'SPC PK15',
                'region': 'Littoral',
                'date': pd.Timestamp(f'2026-{month:02d}-01'),
                'month': month, 'year': 2026,
                'tonnes': 0, 'sacs_50': 0, 'prix_ttc_sac': 0,
                'ca_m_fcfa': round(mat_month, 2),
                'source': 'forfait_pk15_realiste'
            })
    
    return records


def generate_spc_forfait_2027():
    """Generate SPC forfait records for 2027 (12 months).
    
    Forfait ALV: 0 (2026 = 0, pic 2025 non récurrent)
    Forfait MAT: 38.44 M/an (2026 annualisé)
    Saisonnalité: 2025 weights (Q3 pic)
    """
    records = []
    
    for agence, weights in SPC_WEIGHTS_2025.items():
        region = SPC_REGIONS[agence]
        alv_share_pct = weights['ALVEOLES'] / 100
        mat_share_pct = weights['MATERIEL_ELEVAGE'] / 100
        
        alv_ca_agence = FORFAIT_ALVEOLES_ANNUAL_2026 * alv_share_pct  # = 0
        mat_ca_agence = FORFAIT_MAT_ELEVAGE_ANNUAL_2026 * mat_share_pct
        
        for month, weight_m in MONTH_WEIGHTS_ANNUAL.items():
            alv_month = alv_ca_agence * weight_m  # = 0
            mat_month = mat_ca_agence * weight_m
            
            if alv_month > 0.001:
                records.append({
                    'scenario': 'S3_reappro_100',
                    'ref': 'MAT014-80010003',
                    'family': 'ALVEOLES',
                    'agence': agence,
                    'region': region,
                    'date': pd.Timestamp(f'2027-{month:02d}-01'),
                    'month': month, 'year': 2027,
                    'tonnes': 0, 'sacs_50': 0, 'prix_ttc_sac': 0,
                    'ca_m_fcfa': round(alv_month, 2),
                    'source': 'forfait_spc_2026_saisonnalite'
                })
            
            if mat_month > 0.001:
                records.append({
                    'scenario': 'S3_reappro_100',
                    'ref': 'MAT003',
                    'family': 'MATERIEL_ELEVAGE',
                    'agence': agence,
                    'region': region,
                    'date': pd.Timestamp(f'2027-{month:02d}-01'),
                    'month': month, 'year': 2027,
                    'tonnes': 0, 'sacs_50': 0, 'prix_ttc_sac': 0,
                    'ca_m_fcfa': round(mat_month, 2),
                    'source': 'forfait_spc_2026_saisonnalite'
                })
    
    # SPC PK15 forfait réaliste (annuel)
    for month, weight_m in MONTH_WEIGHTS_ANNUAL.items():
        alv_month = FORFAIT_PK15_ALVEOLES * weight_m  # 0
        mat_month = FORFAIT_PK15_MAT_ELEVAGE * weight_m
        
        # PK15 ALV (0)
        if alv_month > 0.001:
            records.append({
                'scenario': 'S3_reappro_100',
                'ref': 'MAT014-80010003',
                'family': 'ALVEOLES',
                'agence': 'SPC PK15',
                'region': 'Littoral',
                'date': pd.Timestamp(f'2027-{month:02d}-01'),
                'month': month, 'year': 2027,
                'tonnes': 0, 'sacs_50': 0, 'prix_ttc_sac': 0,
                'ca_m_fcfa': round(alv_month, 2),
                'source': 'forfait_pk15_realiste'
            })
        # PK15 MAT (0.5 M/an)
        if mat_month > 0.001:
            records.append({
                'scenario': 'S3_reappro_100',
                'ref': 'MAT003',
                'family': 'MATERIEL_ELEVAGE',
                'agence': 'SPC PK15',
                'region': 'Littoral',
                'date': pd.Timestamp(f'2027-{month:02d}-01'),
                'month': month, 'year': 2027,
                'tonnes': 0, 'sacs_50': 0, 'prix_ttc_sac': 0,
                'ca_m_fcfa': round(mat_month, 2),
                'source': 'forfait_pk15_realiste'
            })
    
    return records


if __name__ == '__main__':
    print('=== FORFAIT SPC v2 (basé sur saisonnalité 2026, poids 2025) ===\n')
    print(f"Forfait annuel ALVEOLES 2026: {FORFAIT_ALVEOLES_ANNUAL_2026} M FCFA/an (0 = activité nulle en 2026)")
    print(f"Forfait annuel MATERIEL_ELEVAGE 2026: {FORFAIT_MAT_ELEVAGE_ANNUAL_2026} M FCFA/an (annualisé)")
    print(f"Forfait SPC PK15 MAT_ELEVAGE: {FORFAIT_PK15_MAT_ELEVAGE} M FCFA/an (réaliste)")
    print(f"Forfait SPC PK15 ALVEOLES: {FORFAIT_PK15_ALVEOLES} M FCFA/an (0)")
    print(f"TOTAL annuel SPC v2: {FORFAIT_ALVEOLES_ANNUAL_2026 + FORFAIT_MAT_ELEVAGE_ANNUAL_2026 + FORFAIT_PK15_ALVEOLES + FORFAIT_PK15_MAT_ELEVAGE} M FCFA/an")
    print(f"\nMéthode: Poids 2025 par agence × Forfait total 2026 annualisé")
    print()
    
    print('=== Q4 2026 Forfait SPC v2 ===')
    q4_records = generate_spc_forfait_q4_2026()
    print(f'Total records: {len(q4_records)}')
    q4_df = pd.DataFrame(q4_records)
    if len(q4_df) > 0:
        print(q4_df.groupby(['agence', 'family'])['ca_m_fcfa'].sum().round(2))
        print(f'\nTotal CA Q4: {q4_df["ca_m_fcfa"].sum():.2f} M FCFA')
        print(f'  - ALVEOLES: {q4_df[q4_df["family"]=="ALVEOLES"]["ca_m_fcfa"].sum():.2f} M (0)')
        print(f'  - MAT_ELEVAGE: {q4_df[q4_df["family"]=="MATERIEL_ELEVAGE"]["ca_m_fcfa"].sum():.2f} M')
    
    print()
    print('=== 2027 Forfait SPC v2 ===')
    f2027_records = generate_spc_forfait_2027()
    print(f'Total records: {len(f2027_records)}')
    f2027_df = pd.DataFrame(f2027_records)
    if len(f2027_df) > 0:
        print(f2027_df.groupby(['agence', 'family'])['ca_m_fcfa'].sum().round(2))
        print(f'\nTotal CA 2027: {f2027_df["ca_m_fcfa"].sum():.2f} M FCFA')
        print(f'  - ALVEOLES: {f2027_df[f2027_df["family"]=="ALVEOLES"]["ca_m_fcfa"].sum():.2f} M (0)')
        print(f'  - MAT_ELEVAGE: {f2027_df[f2027_df["family"]=="MATERIEL_ELEVAGE"]["ca_m_fcfa"].sum():.2f} M')
