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
 - Mois courant : extraction ERP upload/ (Livrée uniquement)

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

print(f"\n2. Chargement {MONTH_LABEL} Livrée (ERP)...")
wb = openpyxl.load_workbook(ERP_FILE, read_only=True, data_only=True)
ws = wb['Sheet 1']
sep_records = []
for r in ws.iter_rows(min_row=3, values_only=True):
    if not r or not r[0] or r[0] == 'Total':
        continue
    ref = str(r[0])
    family = get_family(ref)
    if family is None:
        continue
    etat = str(r[13]).strip() if r[13] else ''
    if etat != 'Livrée':
        continue
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
    })
sep_df = pd.DataFrame(sep_records)
print(f"   {len(sep_df):,} lignes Livrée {MONTH_LABEL}: {sep_df['tonnes'].sum():.0f} t, CA {sep_df['montant_ttc'].sum()/1e6:.0f} M FCFA")

cols = ['year', 'month', 'family', 'agence', 'region', 'tonnes', 'sacs_50', 'montant_ttc']
hist = hist[cols]
act_df = pd.concat([hist, sep_df[cols]], ignore_index=True)

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
# 7. Sauvegarde JSON
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
