"""
Calcul de la performance mensuelle — LIVRABLE MENSUEL (réutilisable chaque mois)
Mois courant : Septembre 2026 (01-30/09/2026, mois complet)

Calcule :
 1. Performance globale par famille vs objectifs du mois
 2. Performance par agence × famille vs objectifs du mois
 3. Performance par région × famille vs objectifs du mois
 4. Tendance mensuelle par famille vs objectifs (Jan → mois courant, YTD)
 5. Tendance mensuelle par agence × (Soja, Concentrés) YTD vs objectifs
 6. Tendance mensuelle par région × (Soja, Concentrés) YTD vs objectifs

Objectifs :
 - Mois 1-6  : objectifs Takou (objectives_comparison.json)
 - Mois 7-12 : objectifs S2 recalibrés (s2_recaled_objectives.json)

Actuals :
 - Jan-Août : scripts/dataset_2023_2026.csv (Livrée, 14 agences)
 - Mois courant : extraction ERP upload/ (Livrée + Validée + En cours —
   ces commandes restent rattachées au mois dans la configuration ERP)

Nouveautés septembre 2026 :
 7. Analyse comparée volumes vs CA : le CA suit-il les volumes ?
 8. Encaissements (StatutFacture) : Payée / Créance / Impayée par agence
 9. Mix-produit × encaissements : combos agence × produit gagnants

Usage mensuel : mettre à jour MONTH_NUM, MONTH_LABEL, ERP_FILE puis relancer.
"""
import os
import json
import warnings
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import openpyxl
from openpyxl import load_workbook
from collections import defaultdict

warnings.filterwarnings('ignore')

# ============================================================
# PARAMÈTRES DU MOIS (à mettre à jour chaque mois)
# ============================================================
MONTH_NUM = 9
MONTH_LABEL = "Septembre 2026"
ERP_FILE = "/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (51).xlsx"
OUT_JSON = f"/home/z/my-project/scripts/perf_mensuelle_2026_{MONTH_NUM:02d}.json"
CHARTS_DIR = "/home/z/my-project/work/charts"

# ============================================================
# Référentiels produits / agences (identiques aux scripts forecast)
# ============================================================
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
COMPLEMENT_REFS = {'V300': 1, 'CA003.1': 1, 'CA004.1': 1, 'CA006.1': 1, 'CA001.1': 1,
                   'CA002.1': 1, 'CA005.1': 1, 'CA007.1': 1, 'CA008.1': 1}
ALVEOLES_REFS = {'MAT011-80010002', 'MAT014-80010003', 'MAT015', 'MAT017'}

ALL_REFS = {**SOJA_REFS, **CONC_REFS, **ALIMENT_REFS, **INGREDIENT_REFS,
            **PREMIX_REFS, **MATERIEL_REFS, **COMPLEMENT_REFS}

def get_family(ref):
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

# Familles objectifs (Takou/S2) → familles actuals
FAM_OBJ_TO_ACT = {
    'ALIMENT COMPLET': 'ALIMENT_COMPLET',
    'COMPLEMENT ALIMENTAIRE': 'COMPLEMENT_ALIMENTAIRE',
    'MATERIEL ELEVAGE': 'MATERIEL_ELEVAGE',
    'ALVEOLE': 'ALVEOLES',
}
MAIN_FAMILIES = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT_COMPLET', 'INGREDIENTS',
                 'PREMIX', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES', 'MATERIEL_ELEVAGE']
REGIONS = ['Ouest', 'Centre', 'Littoral']

# ============================================================
# 1. Actuals Jan-Août (dataset historique) + mois courant (ERP Livrée)
# ============================================================
print("=" * 70)
print(f"PERFORMANCE MENSUELLE — {MONTH_LABEL} (mois {MONTH_NUM})")
print("=" * 70)

print("\n1. Chargement actuals Jan-Août 2026 (dataset)...")
hist = pd.read_csv("/home/z/my-project/scripts/dataset_2023_2026.csv",
                   parse_dates=['date'], low_memory=False)
hist = hist[(hist['date'].dt.year == 2026) & (hist['date'].dt.month < MONTH_NUM)].copy()
print(f"   {len(hist):,} records Jan-{MONTH_NUM - 1} 2026")

print(f"\n2. Chargement {MONTH_LABEL} Livrée + Validée + En cours (ERP)...")
wb = openpyxl.load_workbook(ERP_FILE, read_only=True, data_only=True)
ws = wb['Sheet 1']
sep_records = []
etat_counts = {'Livrée': 0, 'Validée': 0, 'En cours': 0}
for r in ws.iter_rows(min_row=3, values_only=True):
    if not r or not r[0] or r[0] == 'Total':
        continue
    ref = str(r[0])
    family = get_family(ref)
    if family is None:
        continue
    etat = str(r[13]).strip() if r[13] else ''
    # Actuals du mois = Livrée + Validée + En cours (restent rattachées au mois dans l'ERP)
    if etat not in ('Livrée', 'Validée', 'En cours'):
        continue
    etat_counts[etat] += 1
    date_str = str(r[6])[:10] if r[6] else ''
    if f'/{MONTH_NUM:02d}/2026' not in date_str:
        continue
    agence_raw = r[15] if r[15] else ''
    if agence_raw not in AGENCE_MAP:
        continue
    agence, region = AGENCE_MAP[agence_raw]
    qte = r[2] or 0
    kg = qte * ALL_REFS.get(ref, 1)
    sep_records.append({
        'date': pd.to_datetime(date_str, format='%d/%m/%Y'),
        'year': 2026, 'month': MONTH_NUM,
        'family': family, 'agence': agence, 'region': region,
        'tonnes': 0 if family in ('MATERIEL_ELEVAGE', 'ALVEOLES') else kg / 1000,
        'sacs_50': kg / 50 if family not in ('MATERIEL_ELEVAGE', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES') else 0,
        'montant_ttc': r[9] or 0,
        'etat': etat,
        'statut_facture': str(r[14]).strip() if r[14] else '(vide)',
    })
sep_df = pd.DataFrame(sep_records)
print(f"   {len(sep_df):,} lignes {MONTH_LABEL} (Livrée {etat_counts['Livrée']}, Validée {etat_counts['Validée']}, "
      f"En cours {etat_counts['En cours']}): {sep_df['tonnes'].sum():.0f} t, CA {sep_df['montant_ttc'].sum()/1e6:.0f} M FCFA")

cols = ['year', 'month', 'family', 'agence', 'region', 'tonnes', 'sacs_50', 'montant_ttc']
hist = hist[cols]
act_df = pd.concat([hist, sep_df[cols]], ignore_index=True)
# sep complet (avec état + statut facture) pour les analyses CA/encaissements
sep_full = sep_df.copy()

# ============================================================
# 3. Objectifs : Takou (mois 1-6) + S2 recalibrés (mois 7-12)
# ============================================================
print("\n3. Chargement objectifs (Takou mois 1-6 + S2 recalibrés mois 7-12)...")
obj_takou = json.load(open('/home/z/my-project/scripts/objectives_comparison.json', encoding='utf-8'))
obj_s2 = json.load(open('/home/z/my-project/scripts/s2_recaled_objectives.json', encoding='utf-8'))
TAK_G = obj_takou['global_objectives']
TAK_A = obj_takou['agence_objectives']
S2_G = obj_s2['global_s2_recaled']
S2_A = obj_s2['agency_s2_recaled']

def fam_obj_name(f):
    return FAM_OBJ_TO_ACT.get(f, f)

def global_obj(fam, m):
    src = TAK_G if m <= 6 else S2_G
    f = next((k for k in src if fam_obj_name(k) == fam), None)
    if f is None:
        return 0.0
    return float(src[f].get(str(m), 0))

def agency_obj(agence, fam, m):
    # Les JSON d'objectifs utilisent la casse 'PK11' (majuscules)
    ag = {'Pk11': 'PK11'}.get(agence, agence)
    src = TAK_A if m <= 6 else S2_A
    if ag not in src:
        return 0.0
    f = next((k for k in src[ag] if fam_obj_name(k) == fam), None)
    if f is None:
        return 0.0
    return float(src[ag][f].get(str(m), 0))

# ============================================================
# 3b. Objectifs CA : S1 fournis (Obj S1.xlsx) + S2 dérivés
#     (prix moyens S1 × volumes S2 recalibrés)
# ============================================================
print("\n3b. Objectifs CA : S1 fournis + S2 dérivés (prix S1 × volumes recalibrés)...")

def _obj_s1_totals(sheet_name):
    wb = load_workbook('/home/z/my-project/upload/Obj S1.xlsx', read_only=True, data_only=True)
    ws = wb[sheet_name]
    d = defaultdict(float)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[1] is None or not isinstance(row[1], str) or not row[1].strip():
            continue  # lignes 'Total X' : la 2e colonne porte la part (%), pas une catégorie
        d[row[1].strip()] += sum(float(v) if v is not None else 0.0 for v in row[3:9])
    wb.close()
    return d

vol_s1 = _obj_s1_totals('Vol')
ca_s1 = _obj_s1_totals('CA')
prix_s1 = {cat: (ca_s1.get(cat, 0.0) / v if v > 0 else None) for cat, v in vol_s1.items()}
for cat, p in prix_s1.items():
    if p:
        print(f"   Prix moyen S1 {cat}: {p:,.0f} F/t")

# CA objectifs Takou mensuels fournis (12 mois)
ca_takou = json.load(open('/home/z/my-project/scripts/ca_obj_real.json', encoding='utf-8'))['ca_obj_monthly']

def obj_name(fam):
    inv = {v: k for k, v in FAM_OBJ_TO_ACT.items()}
    return inv.get(fam, fam)

def ca_obj(fam, m):
    """Objectif CA du mois m (FCFA) : fourni Takou (m ≤ 6), dérivé prix S1 × volume S2 (m ≥ 7)."""
    on = obj_name(fam)
    if m <= 6 or not prix_s1.get(on):
        return float(ca_takou.get(on, [0] * 12)[m - 1])
    return global_obj(fam, m) * prix_s1[on]

# ============================================================
# 4. Sections 1-3 : performance du mois (global / agence / région)
# ============================================================
print("\n4. Sections 1-3 : performance du mois...")
act_month = act_df[act_df['month'] == MONTH_NUM]

global_sept = {}
for fam in MAIN_FAMILIES:
    sub = act_month[act_month['family'] == fam]
    t = float(sub['tonnes'].sum())
    obj = global_obj(fam, MONTH_NUM)
    global_sept[fam] = {
        't': round(t, 1), 'obj': round(obj, 1),
        'pct': round(t / obj * 100, 1) if obj > 0 else None,
        'ecart': round(t - obj, 1),
        'ca': round(float(sub['montant_ttc'].sum()) / 1e6, 1),
    }

agence_sept = {}
for agence, _ in AGENCE_MAP.values():
    agence_sept[agence] = {}
    for fam in MAIN_FAMILIES:
        t = float(act_month[(act_month['family'] == fam) & (act_month['agence'] == agence)]['tonnes'].sum())
        obj = agency_obj(agence, fam, MONTH_NUM)
        agence_sept[agence][fam] = {
            't': round(t, 1), 'obj': round(obj, 1),
            'pct': round(t / obj * 100, 1) if obj > 0 else None,
        }

region_sept = {}
for region in REGIONS:
    region_sept[region] = {}
    ags = [a for a, r in AGENCE_MAP.values() if r == region]
    for fam in MAIN_FAMILIES:
        t = float(act_month[(act_month['family'] == fam) & (act_month['region'] == region)]['tonnes'].sum())
        obj = sum(agency_obj(a, fam, MONTH_NUM) for a in ags)
        region_sept[region][fam] = {
            't': round(t, 1), 'obj': round(obj, 1),
            'pct': round(t / obj * 100, 1) if obj > 0 else None,
        }

# ============================================================
# 5. Section 4 : tendance mensuelle par famille (Jan → mois courant)
# ============================================================
print("\n5. Section 4 : tendance mensuelle par famille...")
fam_monthly = {}
for fam in MAIN_FAMILIES:
    fam_monthly[fam] = {}
    for m in range(1, MONTH_NUM + 1):
        t = float(act_df[(act_df['family'] == fam) & (act_df['month'] == m)]['tonnes'].sum())
        fam_monthly[fam][str(m)] = {'t': round(t, 1), 'obj': round(global_obj(fam, m), 1)}

fam_ytd = {}
for fam in MAIN_FAMILIES:
    t = float(act_df[act_df['family'] == fam]['tonnes'].sum())
    obj = sum(global_obj(fam, m) for m in range(1, MONTH_NUM + 1))
    fam_ytd[fam] = {
        't': round(t, 1), 'obj': round(obj, 1),
        'pct': round(t / obj * 100, 1) if obj > 0 else None,
        'ca': round(float(act_df[act_df['family'] == fam]['montant_ttc'].sum()) / 1e6, 1),
    }

# ============================================================
# 6. Sections 5-6 : tendance mensuelle Soja/Concentrés par agence / région
# ============================================================
print("\n6. Sections 5-6 : tendance Soja/Concentrés agence & région...")
SC_FAM = {'soja': 'TOURTEAUX', 'conc': 'CONCENTRES'}

agence_sc_monthly = {}
for agence, _ in AGENCE_MAP.values():
    agence_sc_monthly[agence] = {}
    for key, fam in SC_FAM.items():
        agence_sc_monthly[agence][key] = {}
        agence_sc_monthly[agence][key + '_obj'] = {}
        for m in range(1, MONTH_NUM + 1):
            t = float(act_df[(act_df['family'] == fam) & (act_df['agence'] == agence) & (act_df['month'] == m)]['tonnes'].sum())
            agence_sc_monthly[agence][key][str(m)] = round(t, 1)
            agence_sc_monthly[agence][key + '_obj'][str(m)] = round(agency_obj(agence, fam, m), 1)

region_sc_monthly = {}
for region in REGIONS:
    region_sc_monthly[region] = {}
    for key, fam in SC_FAM.items():
        region_sc_monthly[region][key] = {}
        region_sc_monthly[region][key + '_obj'] = {}
        for m in range(1, MONTH_NUM + 1):
            t = float(act_df[(act_df['family'] == fam) & (act_df['region'] == region) & (act_df['month'] == m)]['tonnes'].sum())
            region_sc_monthly[region][key][str(m)] = round(t, 1)
            ags = [a for a, r in AGENCE_MAP.values() if r == region]
            obj = sum(agency_obj(a, fam, m) for a in ags)
            region_sc_monthly[region][key + '_obj'][str(m)] = round(obj, 1)

sc_ytd = {}
for key, fam in SC_FAM.items():
    t = float(act_df[act_df['family'] == fam]['tonnes'].sum())
    obj = sum(global_obj(fam, m) for m in range(1, MONTH_NUM + 1))
    sc_ytd[key] = {'t': round(t, 1), 'obj': round(obj, 1),
                   'pct': round(t / obj * 100, 1) if obj > 0 else None}

agence_sc_ytd = {}
for agence, _ in AGENCE_MAP.values():
    agence_sc_ytd[agence] = {}
    for key, fam in SC_FAM.items():
        t = float(act_df[(act_df['family'] == fam) & (act_df['agence'] == agence)]['tonnes'].sum())
        obj = sum(agency_obj(agence, fam, m) for m in range(1, MONTH_NUM + 1))
        agence_sc_ytd[agence][key] = {'t': round(t, 1), 'obj': round(obj, 1),
                                      'pct': round(t / obj * 100, 1) if obj > 0 else None}

region_sc_ytd = {}
for region in REGIONS:
    region_sc_ytd[region] = {}
    ags = [a for a, r in AGENCE_MAP.values() if r == region]
    for key, fam in SC_FAM.items():
        t = float(act_df[(act_df['family'] == fam) & (act_df['region'] == region)]['tonnes'].sum())
        obj = sum(agency_obj(a, fam, m) for a in ags for m in range(1, MONTH_NUM + 1))
        region_sc_ytd[region][key] = {'t': round(t, 1), 'obj': round(obj, 1),
                                      'pct': round(t / obj * 100, 1) if obj > 0 else None}

# Ratio soja:concentrés (bundle)
ratio_month = round(sc_ytd['soja']['t'] / sc_ytd['conc']['t'], 2) if sc_ytd['conc']['t'] > 0 else None

# ============================================================
# 7. Analyse comparée Volumes vs CA + Encaissements
# ============================================================
print("\n7. Analyse comparée volumes vs CA + encaissements...")
# Prix de référence : prix moyen réalisé YTD (Jan → mois courant) par famille
prix_ref = {}
for fam in MAIN_FAMILIES:
    t_ytd_f = float(act_df[act_df['family'] == fam]['tonnes'].sum())
    ca_ytd_f = float(act_df[act_df['family'] == fam]['montant_ttc'].sum())
    prix_ref[fam] = round(ca_ytd_f / t_ytd_f, 0) if t_ytd_f > 0 else None  # FCFA/t

# CA réalisé vs CA attendu — attendu = objectif CA (Takou fourni m ≤ 6, dérivé prix S1 × volume S2 m ≥ 7)
ca_sept = {}
for fam in MAIN_FAMILIES:
    sub = sep_full[sep_full['family'] == fam]
    ca = float(sub['montant_ttc'].sum()) / 1e6
    t = float(sub['tonnes'].sum())
    att = ca_obj(fam, MONTH_NUM) / 1e6
    ca_sept[fam] = {
        'ca': round(ca, 1), 'attendu': round(att, 1) if att > 0 else None,
        'pct_ca': round(ca / att * 100, 1) if att > 0 else None,
        'prix_moy': round(ca * 1e6 / t, 0) if t > 0 else None,
        'prix_obj': round(prix_s1[obj_name(fam)], 0) if prix_s1.get(obj_name(fam)) else None,
        'prix_ref': prix_ref[fam],
        'pct_vol': global_sept[fam]['pct'],
    }

ca_ytd = {}
for fam in MAIN_FAMILIES:
    t = float(act_df[act_df['family'] == fam]['tonnes'].sum())
    ca = float(act_df[act_df['family'] == fam]['montant_ttc'].sum()) / 1e6
    att = sum(ca_obj(fam, m) for m in range(1, MONTH_NUM + 1)) / 1e6
    ca_ytd[fam] = {
        'ca': round(ca, 1), 'attendu': round(att, 1) if att > 0 else None,
        'pct_ca': round(ca / att * 100, 1) if att > 0 else None,
        'pct_vol': fam_ytd[fam]['pct'],
    }

# Agence (soja + concentrés) : % volume vs % CA du mois — le CA suit-il les volumes ?
ca_agence_sept = {}
for agence, _ in AGENCE_MAP.values():
    sc_mask = (act_month['agence'] == agence) & (act_month['family'].isin(['TOURTEAUX', 'CONCENTRES']))
    t_ag = float(act_month[sc_mask]['tonnes'].sum())
    ca_ag = float(act_month[sc_mask]['montant_ttc'].sum()) / 1e6
    obj_ag = sum(agency_obj(agence, fam, MONTH_NUM) for fam in ['TOURTEAUX', 'CONCENTRES'])
    att_ag = sum(agency_obj(agence, fam, MONTH_NUM) * (prix_s1.get(fam, 0) / 1e6)
                 for fam in ['TOURTEAUX', 'CONCENTRES'])
    pct_vol = round(t_ag / obj_ag * 100, 1) if obj_ag > 0 else None
    pct_ca = round(ca_ag / att_ag * 100, 1) if att_ag else None
    ca_agence_sept[agence] = {
        't': round(t_ag, 1), 'obj': round(obj_ag, 1),
        'ca': round(ca_ag, 1), 'attendu': round(att_ag, 1),
        'pct_vol': pct_vol, 'pct_ca': pct_ca,
        'ecart_ca_vol': round(pct_ca - pct_vol, 1) if (pct_ca is not None and pct_vol is not None) else None,
    }

# 8. Encaissements : StatutFacture (Payée / Créance / Impayée)
stat_global = sep_full.groupby('statut_facture')['montant_ttc'].sum() / 1e6
encaissements = {
    'payee': round(float(stat_global.get('Payée', 0)), 1),
    'creance': round(float(stat_global.get('Créance', 0)), 1),
    'impayee': round(float(stat_global.get('Impayée', 0)), 1),
    'vide': round(float(stat_global.get('(vide)', 0)), 1),
    'total': round(float(sep_full['montant_ttc'].sum()) / 1e6, 1),
    'taux_encaissement': None,
    'par_agence': {},
}
if encaissements['total'] > 0:
    encaissements['taux_encaissement'] = round(encaissements['payee'] / encaissements['total'] * 100, 1)
for agence, _ in AGENCE_MAP.values():
    sub = sep_full[sep_full['agence'] == agence]
    ca_a = float(sub['montant_ttc'].sum()) / 1e6
    if ca_a <= 0:
        continue
    g = sub.groupby('statut_facture')['montant_ttc'].sum() / 1e6
    payee = float(g.get('Payée', 0))
    encaissements['par_agence'][agence] = {
        'ca': round(ca_a, 1),
        'payee': round(payee, 1),
        'creance': round(float(g.get('Créance', 0)), 1),
        'impayee': round(float(g.get('Impayée', 0)), 1),
        'taux': round(payee / ca_a * 100, 1) if ca_a > 0 else None,
    }

# ============================================================
# 9. Mix-produit × encaissements : le combo gagnant
# ============================================================
print("\n9. Mix-produit × encaissements (combos agence × produit)...")
ca_total_sept = float(sep_full['montant_ttc'].sum()) / 1e6
mix_famille = {}
for fam in MAIN_FAMILIES:
    sub = sep_full[sep_full['family'] == fam]
    ca = float(sub['montant_ttc'].sum()) / 1e6
    if ca <= 0:
        continue
    g = sub.groupby('statut_facture')['montant_ttc'].sum() / 1e6
    payee = float(g.get('Payée', 0))
    cre = float(g.get('Créance', 0)) + float(g.get('Impayée', 0))
    mix_famille[fam] = {
        'ca': round(ca, 1),
        'part_ca': round(ca / ca_total_sept * 100, 1) if ca_total_sept > 0 else None,
        'payee': round(payee, 1),
        'non_encaisse': round(cre, 1),
        'taux': round(payee / ca * 100, 1) if ca > 0 else None,
    }

# Combos agence × produit (soja & concentrés) : CA, part du CA total, taux d'encaissement
combo_rows = []
for agence, _ in AGENCE_MAP.values():
    for fam in ['TOURTEAUX', 'CONCENTRES']:
        sub = sep_full[(sep_full['agence'] == agence) & (sep_full['family'] == fam)]
        ca = float(sub['montant_ttc'].sum()) / 1e6
        if ca <= 0:
            continue
        payee = float(sub[sub['statut_facture'] == 'Payée']['montant_ttc'].sum()) / 1e6
        cre = float(sub[sub['statut_facture'].isin(['Créance', 'Impayée'])]['montant_ttc'].sum()) / 1e6
        combo_rows.append({
            'agence': agence, 'famille': fam,
            'ca': round(ca, 1),
            'part_ca': round(ca / ca_total_sept * 100, 1) if ca_total_sept > 0 else None,
            'payee': round(payee, 1),
            'non_encaisse': round(cre, 1),
            'taux': round(payee / ca * 100, 1) if ca > 0 else None,
        })
combos = sorted(combo_rows, key=lambda r: r['ca'], reverse=True)

# ============================================================
# 8. Sauvegarde JSON
# ============================================================
out = {
    'meta': {
        'month_num': MONTH_NUM, 'label': MONTH_LABEL,
        'update_date': '30/09/2026',
        'jours_ouvres': 26 if MONTH_NUM == 9 else None,
        'source_erp': os.path.basename(ERP_FILE),
    },
    'global_sept': global_sept,
    'agence_sept': agence_sept,
    'region_sept': region_sept,
    'fam_monthly': fam_monthly,
    'fam_ytd': fam_ytd,
    'agence_sc_monthly': agence_sc_monthly,
    'region_sc_monthly': region_sc_monthly,
    'sc_ytd': sc_ytd,
    'agence_sc_ytd': agence_sc_ytd,
    'region_sc_ytd': region_sc_ytd,
    'ratio_soja_conc_ytd': ratio_month,
    'ca_sept': ca_sept,
    'ca_ytd': ca_ytd,
    'ca_agence_sept': ca_agence_sept,
    'encaissements': encaissements,
    'prix_obj_s1': {fam: (round(prix_s1[obj_name(fam)], 0) if prix_s1.get(obj_name(fam)) else None)
                    for fam in MAIN_FAMILIES},
    'mix_famille': mix_famille,
    'combos': combos,
}
with open(OUT_JSON, 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
print(f"\nSauvegardé: {OUT_JSON}")

# ============================================================
# 8. Graphiques
# ============================================================
print("\n8. Génération des graphiques...")
os.makedirs(CHARTS_DIR, exist_ok=True)
try:
    fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
    fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf')
except Exception:
    pass
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

NAVY = '#1F3A5F'
GOLD = '#C9A961'
GREEN = '#548235'
RED = '#C00000'
GRAY = '#8C8C8C'

MONTHS_FR = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc']
x = list(range(1, MONTH_NUM + 1))
xlabels = MONTHS_FR[:MONTH_NUM]

# --- Chart 1 : global par famille (mois courant, actual vs obj) ---
fam_show = [f for f in MAIN_FAMILIES if fam_ytd[f]['t'] > 0 or fam_ytd[f]['obj'] > 0]
vals_t = [global_sept[f]['t'] for f in fam_show]
vals_o = [global_sept[f]['obj'] for f in fam_show]
ypos = np.arange(len(fam_show))
fig, ax = plt.subplots(figsize=(9, 4.2))
ax.barh(ypos + 0.2, vals_t, height=0.36, color=NAVY, label='Réel')
ax.barh(ypos - 0.2, vals_o, height=0.36, color=GOLD, label='Objectif')
for i, (t, o) in enumerate(zip(vals_t, vals_o)):
    ax.text(t + max(vals_o) * 0.01, i + 0.22, f"{t:,.0f}".replace(',', ' '), va='center', fontsize=8, color=NAVY)
    ax.text(o + max(vals_o) * 0.01, i - 0.2, f"{o:,.0f}".replace(',', ' '), va='center', fontsize=8, color=GRAY)
ax.set_yticks(ypos)
ax.set_yticklabels(fam_show, fontsize=9)
ax.set_xlabel('Tonnes')
ax.set_title(f'Performance par famille — {MONTH_LABEL} (réel vs objectif)', fontsize=11, color=NAVY, fontweight='bold')
ax.legend(loc='lower right', fontsize=8)
ax.spines[['top', 'right']].set_visible(False)
plt.tight_layout()
plt.savefig(f'{CHARTS_DIR}/perf_01_global_mois.png', dpi=150)
plt.close()

# --- Chart 2 : agence × Soja/Concentrés (mois courant) ---
ags = [a for a, _ in AGENCE_MAP.values()]
soja_t = [agence_sept[a]['TOURTEAUX']['t'] for a in ags]
conc_t = [agence_sept[a]['CONCENTRES']['t'] for a in ags]
ypos = np.arange(len(ags))
fig, ax = plt.subplots(figsize=(9, 5.5))
ax.barh(ypos + 0.2, soja_t, height=0.36, color=NAVY, label='Soja réel')
ax.barh(ypos - 0.2, conc_t, height=0.36, color=GOLD, label='Concentrés réel')
ax.set_yticks(ypos)
ax.set_yticklabels(ags, fontsize=8)
ax.set_xlabel('Tonnes')
ax.set_title(f'Soja & Concentrés par agence — {MONTH_LABEL}', fontsize=11, color=NAVY, fontweight='bold')
ax.legend(fontsize=8)
ax.spines[['top', 'right']].set_visible(False)
plt.tight_layout()
plt.savefig(f'{CHARTS_DIR}/perf_02_agence_mois.png', dpi=150)
plt.close()

# --- Chart 3 : région × Soja/Concentrés (mois courant) ---
ypos = np.arange(len(REGIONS))
soja_t = [region_sept[r]['TOURTEAUX']['t'] for r in REGIONS]
soja_o = [region_sept[r]['TOURTEAUX']['obj'] for r in REGIONS]
conc_t = [region_sept[r]['CONCENTRES']['t'] for r in REGIONS]
conc_o = [region_sept[r]['CONCENTRES']['obj'] for r in REGIONS]
fig, ax = plt.subplots(figsize=(8, 3.8))
w = 0.2
ax.bar(ypos - 1.5*w, soja_t, width=w, color=NAVY, label='Soja réel')
ax.bar(ypos - 0.5*w, soja_o, width=w, color='#8FA8C8', label='Soja obj')
ax.bar(ypos + 0.5*w, conc_t, width=w, color=GOLD, label='Conc réel')
ax.bar(ypos + 1.5*w, conc_o, width=w, color='#E3D3AC', label='Conc obj')
for i, v in enumerate(soja_t):
    ax.text(i - 1.5*w, v + 12, f"{v:,.0f}".replace(',', ' '), ha='center', fontsize=7, color=NAVY)
for i, v in enumerate(conc_t):
    ax.text(i + 0.5*w, v + 12, f"{v:,.0f}".replace(',', ' '), ha='center', fontsize=7, color='#8A6D1D')
ax.set_xticks(ypos)
ax.set_xticklabels(REGIONS)
ax.set_ylabel('Tonnes')
ax.set_title(f'Régions × Soja/Concentrés — {MONTH_LABEL} (réel vs objectif)', fontsize=11, color=NAVY, fontweight='bold')
ax.legend(fontsize=7.5, ncol=4)
ax.spines[['top', 'right']].set_visible(False)
plt.tight_layout()
plt.savefig(f'{CHARTS_DIR}/perf_03_region_mois.png', dpi=150)
plt.close()

# --- Chart 4 : tendance mensuelle Soja & Concentrés (global, réel vs obj) ---
fig, axes = plt.subplots(1, 2, figsize=(11, 4))
for ax, key, fam, title in [
    (axes[0], 'soja', 'TOURTEAUX', 'Soja (TOURTEAUX)'),
    (axes[1], 'conc', 'CONCENTRES', 'Concentrés'),
]:
    t = [fam_monthly[fam][str(m)]['t'] for m in x]
    o = [fam_monthly[fam][str(m)]['obj'] for m in x]
    ax.plot(x, o, color=GOLD, linestyle='--', marker='o', markersize=4, label='Objectif')
    ax.plot(x, t, color=NAVY, marker='s', markersize=4, label='Réel')
    ax.fill_between(x, 0, t, color=NAVY, alpha=0.08)
    ax.set_xticks(x)
    ax.set_xticklabels(xlabels, fontsize=8)
    ax.set_title(f'{title} — tendance mensuelle YTD', fontsize=10, color=NAVY, fontweight='bold')
    ax.legend(fontsize=8)
    ax.set_ylabel('Tonnes', fontsize=8)
    ax.spines[['top', 'right']].set_visible(False)
plt.suptitle(f'Tendance Jan → {MONTH_LABEL} (réel vs objectif mensuel)', fontsize=12, color=NAVY, fontweight='bold')
plt.tight_layout()
plt.savefig(f'{CHARTS_DIR}/perf_04_fam_mensuel.png', dpi=150)
plt.close()

# --- Chart 5 : heatmaps agence × mois (Soja % obj et Conc % obj) ---
fig, axes = plt.subplots(1, 2, figsize=(11, 5.5))
for ax, key, fam, title in [
    (axes[0], 'soja', 'TOURTEAUX', 'Soja : % objectif atteint'),
    (axes[1], 'conc', 'CONCENTRES', 'Concentrés : % objectif atteint'),
]:
    grid = np.zeros((len(ags), MONTH_NUM))
    for i, a in enumerate(ags):
        for j, m in enumerate(x):
            t = agence_sc_monthly[a][key][str(m)]
            o = agence_sc_monthly[a][key + '_obj'][str(m)]
            grid[i, j] = t / o * 100 if o > 0 else np.nan
    cmap = plt.cm.RdYlGn
    cmap.set_bad('#DDDDDD')
    im = ax.imshow(grid, cmap=cmap, vmin=40, vmax=160, aspect='auto')
    ax.set_xticks(range(MONTH_NUM))
    ax.set_xticklabels(xlabels, fontsize=8)
    ax.set_yticks(range(len(ags)))
    ax.set_yticklabels(ags, fontsize=7)
    for i in range(grid.shape[0]):
        for j in range(grid.shape[1]):
            v = grid[i, j]
            if not np.isnan(v):
                ax.text(j, i, f"{v:.0f}", ha='center', va='center', fontsize=6.5,
                        color='white' if v < 60 or v > 140 else 'black')
    ax.set_title(title, fontsize=10, color=NAVY, fontweight='bold')
    fig.colorbar(im, ax=ax, fraction=0.03)
plt.suptitle(f'Agences × mois — % d\'objectif (100 = atteint) — Jan → {MONTH_LABEL}', fontsize=12, color=NAVY, fontweight='bold')
plt.tight_layout()
plt.savefig(f'{CHARTS_DIR}/perf_05_heatmap_agence.png', dpi=150)
plt.close()

# --- Chart 6 : tendance mensuelle par région (Soja & Conc, réel vs obj) ---
fig, axes = plt.subplots(1, 2, figsize=(11, 4))
colors_r = {'Ouest': NAVY, 'Centre': GREEN, 'Littoral': RED}
for ax, key, fam, title in [
    (axes[0], 'soja', 'TOURTEAUX', 'Soja par région'),
    (axes[1], 'conc', 'CONCENTRES', 'Concentrés par région'),
]:
    tot_obj = [sum(region_sc_monthly[r][key + '_obj'][str(m)] for r in REGIONS) for m in x]
    ax.plot(x, tot_obj, color='black', linestyle=':', linewidth=1.2, label='Objectif national')
    for r in REGIONS:
        t = [region_sc_monthly[r][key][str(m)] for m in x]
        ax.plot(x, t, color=colors_r[r], marker='o', markersize=3.5, label=r)
    ax.set_xticks(x)
    ax.set_xticklabels(xlabels, fontsize=8)
    ax.set_title(f'{title} — tendance mensuelle', fontsize=10, color=NAVY, fontweight='bold')
    ax.legend(fontsize=7.5)
    ax.set_ylabel('Tonnes', fontsize=8)
    ax.spines[['top', 'right']].set_visible(False)
plt.suptitle(f'Tendance régionale Jan → {MONTH_LABEL}', fontsize=12, color=NAVY, fontweight='bold')
plt.tight_layout()
plt.savefig(f'{CHARTS_DIR}/perf_06_region_mensuel.png', dpi=150)
plt.close()

print(f"   6 graphiques générés dans {CHARTS_DIR}")

# --- Chart 7 : par agence, % volume vs % CA (soja + concentrés) ---
ag_ca = [a for a in ags if ca_agence_sept[a]['pct_vol'] is not None and ca_agence_sept[a]['pct_ca'] is not None]
ag_ca.sort(key=lambda a: ca_agence_sept[a]['ecart_ca_vol'] or 0)
ypos = np.arange(len(ag_ca))
v_vol = [ca_agence_sept[a]['pct_vol'] for a in ag_ca]
v_ca = [ca_agence_sept[a]['pct_ca'] for a in ag_ca]
fig, ax = plt.subplots(figsize=(9, 5.5))
ax.barh(ypos + 0.2, v_vol, height=0.36, color=NAVY, label='% volume (t)')
ax.barh(ypos - 0.2, v_ca, height=0.36, color=GOLD, label='% CA')
ax.axvline(100, color=RED, linestyle='--', linewidth=1, alpha=0.7)
ax.text(100.5, len(ag_ca) - 0.4, 'Objectif 100%', color=RED, fontsize=8)
for i, a in enumerate(ag_ca):
    ax.text(v_vol[i] + 1, i + 0.2, f"{v_vol[i]:.0f}%", va='center', fontsize=7.5, color=NAVY)
    ax.text(v_ca[i] + 1, i - 0.2, f"{v_ca[i]:.0f}%", va='center', fontsize=7.5, color='#8A6D1D')
ax.set_yticks(ypos)
ax.set_yticklabels(ag_ca, fontsize=8)
ax.set_xlabel('% objectif')
ax.set_title(f'Volumes vs CA par agence — {MONTH_LABEL} (soja + concentrés)', fontsize=11, color=NAVY, fontweight='bold')
ax.legend(fontsize=8, loc='lower right')
ax.spines[['top', 'right']].set_visible(False)
plt.tight_layout()
plt.savefig(f'{CHARTS_DIR}/perf_07_agence_ca_vs_vol.png', dpi=150)
plt.close()

# --- Chart 8 : encaissements par agence (Payée / Créance / Impayée) ---
enc_ags = sorted(
    [a for a in encaissements['par_agence']],
    key=lambda a: encaissements['par_agence'][a]['creance'] + encaissements['par_agence'][a]['impayee'],
    reverse=True,
)
ypos = np.arange(len(enc_ags))
v_pay = [encaissements['par_agence'][a]['payee'] for a in enc_ags]
v_cre = [encaissements['par_agence'][a]['creance'] for a in enc_ags]
v_imp = [encaissements['par_agence'][a]['impayee'] for a in enc_ags]
fig, ax = plt.subplots(figsize=(9, 5.5))
ax.barh(ypos, v_pay, height=0.6, color=GREEN, label='Payée')
ax.barh(ypos, v_cre, height=0.6, left=v_pay, color=GOLD, label='Créance')
ax.barh(ypos, v_imp, height=0.6, left=[p + c for p, c in zip(v_pay, v_cre)], color=RED, label='Impayée')
ax.set_yticks(ypos)
ax.set_yticklabels(enc_ags, fontsize=8)
ax.set_xlabel('M FCFA')
ax.set_title(f'Encaissements par agence — {MONTH_LABEL} (StatutFacture)', fontsize=11, color=NAVY, fontweight='bold')
ax.legend(fontsize=8, loc='lower right')
ax.spines[['top', 'right']].set_visible(False)
plt.tight_layout()
plt.savefig(f'{CHARTS_DIR}/perf_08_encaissements.png', dpi=150)
plt.close()

print(f"   8 graphiques générés dans {CHARTS_DIR}")

# --- Chart 9 : mix-produit × encaissements (CA par famille, Payée vs non encaissé) ---
mix_ags = sorted(mix_famille, key=lambda f: mix_famille[f]['ca'], reverse=True)
ypos = np.arange(len(mix_ags))
v_pay = [mix_famille[f]['payee'] for f in mix_ags]
v_cre = [mix_famille[f]['non_encaisse'] for f in mix_ags]
fig, ax = plt.subplots(figsize=(9, 5))
ax.barh(ypos, v_pay, height=0.6, color=GREEN, label='Encaissé (Payée)')
ax.barh(ypos, v_cre, height=0.6, left=v_pay, color=RED, label='Créance + impayée')
for i, f in enumerate(mix_ags):
    tot = v_pay[i] + v_cre[i]
    part = mix_famille[f]['part_ca']
    ax.text(tot + 15, i, f"{tot:,.0f} M ({part}%)".replace(',', ' '), va='center', fontsize=8, color=NAVY)
ax.set_yticks(ypos)
ax.set_yticklabels(mix_ags, fontsize=8)
ax.set_xlabel('M FCFA')
ax.set_title(f'Mix-produit et encaissement — {MONTH_LABEL}', fontsize=11, color=NAVY, fontweight='bold')
ax.legend(fontsize=8, loc='lower right')
ax.spines[['top', 'right']].set_visible(False)
plt.tight_layout()
plt.savefig(f'{CHARTS_DIR}/perf_09_mix_encaissements.png', dpi=150)
plt.close()

# --- Chart 10 : combos agence × produit (CA vs taux d'encaissement) ---
fig, ax = plt.subplots(figsize=(9, 5.5))
for fam, col, lab in [('TOURTEAUX', NAVY, 'Soja'), ('CONCENTRES', GOLD, 'Concentrés')]:
    sub = [c for c in combos if c['famille'] == fam]
    xs = [c['ca'] for c in sub]
    ys = [c['taux'] for c in sub]
    ax.scatter(xs, ys, s=60, color=col, label=lab, zorder=3)
    for c in sub:
        ax.annotate(c['agence'], (c['ca'], c['taux']), fontsize=7,
                    xytext=(4, 4), textcoords='offset points', color=col)
ax.axhline(100, color=GREEN, linestyle='--', linewidth=0.8, alpha=0.7)
ax.text(ax.get_xlim()[1] * 0.05, 100.15, 'Encaissement total (100%)', color=GREEN, fontsize=8)
ax.set_xlabel('CA Septembre (M FCFA)')
ax.set_ylabel("Taux d'encaissement (%)")
ax.set_title(f'Combos agence × produit : CA et encaissement — {MONTH_LABEL}', fontsize=11, color=NAVY, fontweight='bold')
ax.legend(fontsize=8, loc='lower right')
ax.spines[['top', 'right']].set_visible(False)
plt.tight_layout()
plt.savefig(f'{CHARTS_DIR}/perf_10_combos.png', dpi=150)
plt.close()

print(f"   10 graphiques générés dans {CHARTS_DIR}")

# ============================================================
# Résumé console
# ============================================================
print("\n" + "=" * 70)
print("RÉSUMÉ")
print("=" * 70)
print(f"Soja {MONTH_LABEL}: {global_sept['TOURTEAUX']['t']:,.0f} t vs obj {global_sept['TOURTEAUX']['obj']:,.0f} t ({global_sept['TOURTEAUX']['pct']}%)")
print(f"Concentrés {MONTH_LABEL}: {global_sept['CONCENTRES']['t']:,.0f} t vs obj {global_sept['CONCENTRES']['obj']:,.0f} t ({global_sept['CONCENTRES']['pct']}%)")
print(f"Soja YTD: {sc_ytd['soja']['t']:,.0f} t vs obj {sc_ytd['soja']['obj']:,.0f} t ({sc_ytd['soja']['pct']}%)")
print(f"Concentrés YTD: {sc_ytd['conc']['t']:,.0f} t vs obj {sc_ytd['conc']['obj']:,.0f} t ({sc_ytd['conc']['pct']}%)")
print(f"Ratio soja:conc YTD: {ratio_month}")
print(f"\nCA {MONTH_LABEL}: {encaissements['total']:,.1f} M FCFA | encaissé {encaissements['payee']:,.1f} M ({encaissements['taux_encaissement']}%)")
print(f"  Créances: {encaissements['creance']:,.1f} M | Impayées: {encaissements['impayee']:,.1f} M")
for fam in MAIN_FAMILIES:
    if ca_sept[fam]['pct_ca'] is not None:
        print(f"CA {fam} {MONTH_LABEL}: {ca_sept[fam]['ca']:,.1f} M vs attendu {ca_sept[fam]['attendu']:,.1f} M ({ca_sept[fam]['pct_ca']}%) — volume {ca_sept[fam]['pct_vol']}%")
alerts = sorted(
    [(a, ca_agence_sept[a]['ecart_ca_vol']) for a in ca_agence_sept
     if ca_agence_sept[a]['ecart_ca_vol'] is not None and ca_agence_sept[a]['ecart_ca_vol'] < -10],
    key=lambda x: x[1],
)
if alerts:
    print("\n⚠ CA en retard sur volumes (>10 pts):")
    for a, e in alerts:
        print(f"  {a}: vol {ca_agence_sept[a]['pct_vol']}% vs CA {ca_agence_sept[a]['pct_ca']}% (écart {e} pts)")

print("\nMix-produit (part du CA, taux d'encaissement):")
for f, v in sorted(mix_famille.items(), key=lambda kv: kv[1]['ca'], reverse=True):
    print(f"  {f:25s} CA {v['ca']:>8,.1f} M ({v['part_ca']:>5}% du CA) | encaissé {v['taux']}% | non encaissé {v['non_encaisse']:,.1f} M")
print("\nCombos gagnants (CA élevé + encaissement total):")
for c in combos[:5]:
    star = " ★" if c['taux'] == 100.0 else ""
    print(f"  {c['famille']:12s} × {c['agence']:12s} CA {c['ca']:>8,.1f} M ({c['part_ca']}% du CA total), taux {c['taux']}%{star}")
risky = [c for c in combos if c['non_encaisse'] > 0]
if risky:
    print("\nCombos à risque (non encaissé):")
    for c in risky:
        print(f"  {c['famille']:12s} × {c['agence']:12s} non encaissé {c['non_encaisse']:,.1f} M (taux {c['taux']}%)")
