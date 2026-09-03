"""
Forfait SPC pour ALVEOLES + MATERIEL_ELEVAGE.

Définit un forfait global basé sur 2025 pour les SPC agencies, distribué par poids 2025.
Pour SPC PK15 (0 en 2025), applique un forfait réaliste.

Forfait global annuel 2025 (basé sur 2025 SPC agency sales):
  ALVEOLES: 318.2 M FCFA/an (10 agences SPC, SPC Baf-Chefferie 97.83%)
  MATERIEL_ELEVAGE: 30.3 M FCFA/an (10 agences SPC, SPC Baf-Chefferie 79.79%)
  TOTAL: 348.5 M FCFA/an

Pour 2027 (forecast 12 mois), on applique:
  - Forfait ALVEOLES annuel: 318.2 M FCFA/an (maintien de l'activité 2025)
  - Forfait MAT_ELEVAGE annuel: 30.3 M FCFA/an (maintien)
  - Total SPC: 348.5 M FCFA/an

Pour Q4 2026 (4 mois): on applique 1/3 du forfait annuel (Q4 ≈ 33% activité annuelle):
  - Forfait ALVEOLES Q4: ~105 M FCFA
  - Forfait MAT_ELEVAGE Q4: ~10 M FCFA
  - Total SPC Q4: ~115 M FCFA

Poids 2025 par agence SPC:
  SPC Baf-Chefferie: ALV 97.83%, MAT 79.79%
  SPC TPO:           ALV 0.90%,  MAT 0.01%
  SPC Ndere:         ALV 0.64%,  MAT 2.53%
  SPC Kye-Ossi:      ALV 0.13%,  MAT 5.71%
  SPC DLA-Beri:      ALV 0.20%,  MAT 4.19%
  SPC Yassa:         ALV 0.03%,  MAT 5.98%
  SPC Village:       ALV 0.11%,  MAT 0.91%
  SPC Dschang:       ALV 0.12%,  MAT 0.71%
  SPC Buea:          ALV 0.05%,  MAT 0.18%
  SPC PK15:          ALV 0%,     MAT 0% (forfait réaliste: 0.5 M/an ALV + 0.5 M/an MAT)

Pour SPC PK15 (0 en 2025), forfait réaliste:
  ALVEOLES: 0.5 M FCFA/an (basé sur taille similaire à SPC Yassa/Village)
  MATERIEL_ELEVAGE: 0.5 M FCFA/an
"""
import pandas as pd
import json

# 2025 SPC weights
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
    # SPC PK15: 0% in 2025, forfait réaliste defined below
}

# SPC agencies regions (for forecast integration)
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

# Forfait global annuel (basé sur 2025)
FORFAIT_ALVEOLES_ANNUAL = 318.2  # M FCFA/an (total SPC ALVEOLES 2025)
FORFAIT_MAT_ELEVAGE_ANNUAL = 30.3  # M FCFA/an (total SPC MATERIEL_ELEVAGE 2025)

# Forfait SPC PK15 (poids nul en 2025) — forfait réaliste
FORFAIT_PK15_ALVEOLES = 0.5  # M FCFA/an
FORFAIT_PK15_MAT_ELEVAGE = 0.5  # M FCFA/an

# Saisonnalité approximative (Q4 ≈ 33% de l'année, pics en octobre et décembre)
MONTH_WEIGHTS_Q4 = {9: 0.20, 10: 0.35, 11: 0.20, 12: 0.25}  # Sep, Oct, Nov, Dec
MONTH_WEIGHTS_ANNUAL = {
    1: 0.07, 2: 0.06, 3: 0.07, 4: 0.07, 5: 0.08, 6: 0.08,
    7: 0.07, 8: 0.06, 9: 0.09, 10: 0.13, 11: 0.10, 12: 0.12
}

def generate_spc_forfait_q4_2026():
    """Generate SPC forfait records for Q4 2026 (Sep, Oct, Nov, Dec)."""
    records = []
    
    # Forfait total Q4 = (ALV + MAT) * 1/3 (Q4 ≈ 33% of year)
    forfait_alv_q4 = FORFAIT_ALVEOLES_ANNUAL * 0.33  # ~105 M
    forfait_mat_q4 = FORFAIT_MAT_ELEVAGE_ANNUAL * 0.33  # ~10 M
    
    for agence, weights in SPC_WEIGHTS_2025.items():
        region = SPC_REGIONS[agence]
        # ALVEOLES share
        alv_share_pct = weights['ALVEOLES'] / 100
        mat_share_pct = weights['MATERIEL_ELEVAGE'] / 100
        
        # ALVEOLES CA per agency
        alv_ca_agence = forfait_alv_q4 * alv_share_pct
        # MAT_ELEVAGE CA per agency
        mat_ca_agence = forfait_mat_q4 * mat_share_pct
        
        # Distribute across Q4 months
        for month, weight_m in MONTH_WEIGHTS_Q4.items():
            # Normalize Q4 weights to sum to 1
            q4_total = sum(MONTH_WEIGHTS_Q4.values())
            norm_weight = weight_m / q4_total
            
            alv_month = alv_ca_agence * norm_weight
            mat_month = mat_ca_agence * norm_weight
            
            # ALVEOLES record
            if alv_month > 0.01:
                records.append({
                    'scenario': 'S3_reappro_100',
                    'ref': 'MAT014-80010003',  # Main ALVEOLE ref
                    'family': 'ALVEOLES',
                    'agence': agence,
                    'region': region,
                    'date': pd.Timestamp(f'2026-{month:02d}-01'),
                    'month': month, 'year': 2026,
                    'tonnes': 0,  # ALVEOLES not expressed in tonnes
                    'sacs_50': 0,
                    'prix_ttc_sac': 0,
                    'ca_m_fcfa': round(alv_month, 2),
                    'source': 'forfait_spc_2025'
                })
            
            # MATERIEL_ELEVAGE record
            if mat_month > 0.01:
                records.append({
                    'scenario': 'S3_reappro_100',
                    'ref': 'MAT003',  # ABREUVOIR AUTO (main ref)
                    'family': 'MATERIEL_ELEVAGE',
                    'agence': agence,
                    'region': region,
                    'date': pd.Timestamp(f'2026-{month:02d}-01'),
                    'month': month, 'year': 2026,
                    'tonnes': 0,
                    'sacs_50': 0,
                    'prix_ttc_sac': 0,
                    'ca_m_fcfa': round(mat_month, 2),
                    'source': 'forfait_spc_2025'
                })
    
    # SPC PK15 (0 en 2025) — forfait réaliste
    pk15_alv_q4 = FORFAIT_PK15_ALVEOLES * 0.33
    pk15_mat_q4 = FORFAIT_PK15_MAT_ELEVAGE * 0.33
    for month, weight_m in MONTH_WEIGHTS_Q4.items():
        q4_total = sum(MONTH_WEIGHTS_Q4.values())
        norm_weight = weight_m / q4_total
        
        alv_month = pk15_alv_q4 * norm_weight
        mat_month = pk15_mat_q4 * norm_weight
        
        records.append({
            'scenario': 'S3_reappro_100',
            'ref': 'MAT014-80010003',
            'family': 'ALVEOLES',
            'agence': 'SPC PK15',
            'region': 'Littoral',
            'date': pd.Timestamp(f'2026-{month:02d}-01'),
            'month': month, 'year': 2026,
            'tonnes': 0,
            'sacs_50': 0,
            'prix_ttc_sac': 0,
            'ca_m_fcfa': round(alv_month, 2),
            'source': 'forfait_pk15_realiste'
        })
        records.append({
            'scenario': 'S3_reappro_100',
            'ref': 'MAT003',
            'family': 'MATERIEL_ELEVAGE',
            'agence': 'SPC PK15',
            'region': 'Littoral',
            'date': pd.Timestamp(f'2026-{month:02d}-01'),
            'month': month, 'year': 2026,
            'tonnes': 0,
            'sacs_50': 0,
            'prix_ttc_sac': 0,
            'ca_m_fcfa': round(mat_month, 2),
            'source': 'forfait_pk15_realiste'
        })
    
    return records


def generate_spc_forfait_2027():
    """Generate SPC forfait records for 2027 (12 months)."""
    records = []
    
    for agence, weights in SPC_WEIGHTS_2025.items():
        region = SPC_REGIONS[agence]
        alv_share_pct = weights['ALVEOLES'] / 100
        mat_share_pct = weights['MATERIEL_ELEVAGE'] / 100
        
        alv_ca_agence = FORFAIT_ALVEOLES_ANNUAL * alv_share_pct
        mat_ca_agence = FORFAIT_MAT_ELEVAGE_ANNUAL * mat_share_pct
        
        for month, weight_m in MONTH_WEIGHTS_ANNUAL.items():
            alv_month = alv_ca_agence * weight_m
            mat_month = mat_ca_agence * weight_m
            
            if alv_month > 0.01:
                records.append({
                    'scenario': 'S3_reappro_100',
                    'ref': 'MAT014-80010003',
                    'family': 'ALVEOLES',
                    'agence': agence,
                    'region': region,
                    'date': pd.Timestamp(f'2027-{month:02d}-01'),
                    'month': month, 'year': 2027,
                    'tonnes': 0,
                    'sacs_50': 0,
                    'prix_ttc_sac': 0,
                    'ca_m_fcfa': round(alv_month, 2),
                    'source': 'forfait_spc_2025'
                })
            
            if mat_month > 0.01:
                records.append({
                    'scenario': 'S3_reappro_100',
                    'ref': 'MAT003',
                    'family': 'MATERIEL_ELEVAGE',
                    'agence': agence,
                    'region': region,
                    'date': pd.Timestamp(f'2027-{month:02d}-01'),
                    'month': month, 'year': 2027,
                    'tonnes': 0,
                    'sacs_50': 0,
                    'prix_ttc_sac': 0,
                    'ca_m_fcfa': round(mat_month, 2),
                    'source': 'forfait_spc_2025'
                })
    
    # SPC PK15 forfait réaliste
    for month, weight_m in MONTH_WEIGHTS_ANNUAL.items():
        alv_month = FORFAIT_PK15_ALVEOLES * weight_m
        mat_month = FORFAIT_PK15_MAT_ELEVAGE * weight_m
        
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
    # Test: print forfait summary
    print('=== FORFAIT SPC (basé sur 2025) ===\n')
    print(f"Forfait annuel ALVEOLES: {FORFAIT_ALVEOLES_ANNUAL} M FCFA/an")
    print(f"Forfait annuel MATERIEL_ELEVAGE: {FORFAIT_MAT_ELEVAGE_ANNUAL} M FCFA/an")
    print(f"Forfait SPC PK15 ALVEOLES: {FORFAIT_PK15_ALVEOLES} M FCFA/an (réaliste)")
    print(f"Forfait SPC PK15 MAT_ELEVAGE: {FORFAIT_PK15_MAT_ELEVAGE} M FCFA/an (réaliste)")
    print(f"TOTAL annuel SPC: {FORFAIT_ALVEOLES_ANNUAL + FORFAIT_MAT_ELEVAGE_ANNUAL + FORFAIT_PK15_ALVEOLES + FORFAIT_PK15_MAT_ELEVAGE} M FCFA/an")
    print()
    
    # Generate Q4 2026 forfait
    print('=== Q4 2026 Forfait SPC ===')
    q4_records = generate_spc_forfait_q4_2026()
    print(f'Total records: {len(q4_records)}')
    q4_df = pd.DataFrame(q4_records)
    print(q4_df.groupby(['agence', 'family'])['ca_m_fcfa'].sum().round(2))
    print(f'\nTotal CA Q4: {q4_df["ca_m_fcfa"].sum():.2f} M FCFA')
    print(f'  - ALVEOLES: {q4_df[q4_df["family"]=="ALVEOLES"]["ca_m_fcfa"].sum():.2f} M')
    print(f'  - MAT_ELEVAGE: {q4_df[q4_df["family"]=="MATERIEL_ELEVAGE"]["ca_m_fcfa"].sum():.2f} M')
    
    print()
    print('=== 2027 Forfait SPC ===')
    f2027_records = generate_spc_forfait_2027()
    print(f'Total records: {len(f2027_records)}')
    f2027_df = pd.DataFrame(f2027_records)
    print(f2027_df.groupby(['agence', 'family'])['ca_m_fcfa'].sum().round(2))
    print(f'\nTotal CA 2027: {f2027_df["ca_m_fcfa"].sum():.2f} M FCFA')
    print(f'  - ALVEOLES: {f2027_df[f2027_df["family"]=="ALVEOLES"]["ca_m_fcfa"].sum():.2f} M')
    print(f'  - MAT_ELEVAGE: {f2027_df[f2027_df["family"]=="MATERIEL_ELEVAGE"]["ca_m_fcfa"].sum():.2f} M')
