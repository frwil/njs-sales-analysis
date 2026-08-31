"""Compute CONCENTRÉS regional performance for August MTD (au 14/08)."""
import json

# August objectives by agence (from s2_recaled_objectives.json, month 8)
OBJ_AOUT = {
    'Ahala': 94.6, 'Bertoua': 72.3, 'Buea': 55.8, 'Djeleng': 112.1,
    'Famla': 394.7, 'Mbouda': 98.5, 'Messassi': 194.8, 'Ndobo': 196.0,
    'Ngaoundere': 45.9, 'Nkoabang': 48.2, 'Nkolbisson': 35.6,
    'Nkongsamba': 67.8, 'Pk11': 47.8, 'Village': 70.6,
}

REGION_MAP = {
    'Ahala': 'Centre', 'Bertoua': 'Centre', 'Ngaoundere': 'Centre',
    'Nkoabang': 'Centre', 'Nkolbisson': 'Centre', 'Messassi': 'Centre',
    'Famla': 'Ouest', 'Djeleng': 'Ouest', 'Mbouda': 'Ouest',
    'Ndobo': 'Littoral', 'Buea': 'Littoral', 'Nkongsamba': 'Littoral',
    'Pk11': 'Littoral', 'Village': 'Littoral',
}

# Projections August (from compute_aout_mtd_metrics.py)
# Format: agence -> (conc_mtd_t, proj_t)
with open('/home/z/my-project/scripts/aout_mtd_30.json') as f:
    summary = json.load(f)

PROJ = {}
for item in summary['conc_by_agence']:
    ag = item['agence']
    # Normalize names
    if ag == 'Pk11': ag = 'Pk11'
    if ag == 'Ngaoundere': ag = 'Ngaoundere'
    # The proj_t value is in tonnes - convert to numeric float (not multiplied)
    proj_t = float(item['proj_t'])
    # The value appears to be in kg or scaled - let's check the raw value
    # From compute_aout_mtd_metrics.py: 'proj_t': round(c/days_elapsed*total_days_aug, 0)
    # where c is in kg → so proj_t is in kg wrongly labeled. We need /1000.
    PROJ[ag] = proj_t / 1000  # convert kg to t
print("Projections par agence:")
for ag, p in sorted(PROJ.items()):
    print(f"  {ag}: {p} t")

# Compute by region
reg_proj = {'Ouest': 0, 'Centre': 0, 'Littoral': 0}
reg_obj = {'Ouest': 0, 'Centre': 0, 'Littoral': 0}
for ag, proj in PROJ.items():
    if ag in REGION_MAP and ag in OBJ_AOUT:
        reg = REGION_MAP[ag]
        reg_proj[reg] += proj
        reg_obj[reg] += OBJ_AOUT[ag]

print("\n=== PERFORMANCE PAR RÉGION (CONCENTRÉS Août MTD au 14/08) ===")
print(f"{'Région':10} {'Proj (t)':>10} {'Obj (t)':>10} {'%':>8} {'Écart (t)':>10}")
total_proj = 0
total_obj = 0
for reg in ['Ouest', 'Centre', 'Littoral']:
    p = reg_proj[reg]
    o = reg_obj[reg]
    pct = p / o * 100 if o > 0 else 0
    ecart = p - o
    print(f"{reg:10} {p:>10.0f} {o:>10.1f} {pct:>7.1f}% {ecart:>+10.0f}")
    total_proj += p
    total_obj += o
pct_total = total_proj / total_obj * 100
print(f"{'TOTAL':10} {total_proj:>10.0f} {total_obj:>10.1f} {pct_total:>7.1f}% {total_proj-total_obj:>+10.0f}")

# Per-agence performance
print("\n=== PERFORMANCE PAR AGENCE (CONCENTRÉS Août MTD au 14/08) ===")
print(f"{'Agence':15} {'Région':10} {'Proj (t)':>10} {'Obj (t)':>10} {'%':>8} {'Écart (t)':>10}")
agences = []
for ag, proj in PROJ.items():
    if ag in OBJ_AOUT:
        obj = OBJ_AOUT[ag]
        pct = proj / obj * 100
        ecart = proj - obj
        agences.append((ag, REGION_MAP[ag], proj, obj, pct, ecart))

# Sort by % desc
agences.sort(key=lambda x: -x[4])
for ag, reg, proj, obj, pct, ecart in agences:
    marker = '✅' if pct >= 100 else '❌'
    print(f"{ag:15} {reg:10} {proj:>10.0f} {obj:>10.1f} {pct:>7.0f}% {ecart:>+10.0f} {marker}")

# Stats
nb_above = sum(1 for a in agences if a[4] >= 100)
nb_below = sum(1 for a in agences if a[4] < 100)
print(f"\nAgences au-dessus de l'objectif: {nb_above}/14")
print(f"Agences en-dessous: {nb_below}/14")

# Top performers et critiques
top = agences[0]
worst = agences[-1]
print(f"\nTop performer: {top[0]} ({top[4]:.0f}%, +{top[5]:.0f} t)")
print(f"Plus critique: {worst[0]} ({worst[4]:.0f}%, {worst[5]:.0f} t)")

# Sans NDOBO
littoral_sans_ndobo_proj = reg_proj['Littoral'] - PROJ.get('Ndobo', 0)
littoral_sans_ndobo_obj = reg_obj['Littoral'] - OBJ_AOUT['Ndobo']
pct_littoral_sans = littoral_sans_ndobo_proj / littoral_sans_ndobo_obj * 100
print(f"\nLittoral sans NDOBO: {littoral_sans_ndobo_proj:.0f} / {littoral_sans_ndobo_obj:.1f} = {pct_littoral_sans:.0f}%")

# Si NDOBO retrouve son niveau objectif (196 t)
proj_avec_ndobo_obj = reg_proj['Littoral'] - PROJ.get('Ndobo', 0) + OBJ_AOUT['Ndobo']
pct_littoral_avec_ndobo_obj = proj_avec_ndobo_obj / reg_obj['Littoral'] * 100
print(f"Si NDOBO à 100% obj (+{OBJ_AOUT['Ndobo']-PROJ.get('Ndobo',0):.0f} t): Littoral → {pct_littoral_avec_ndobo_obj:.0f}%")
total_proj_avec_ndobo = total_proj - PROJ.get('Ndobo', 0) + OBJ_AOUT['Ndobo']
print(f"Total CONCENTRES: {total_proj_avec_ndobo:.0f} t vs {total_obj:.0f} t obj → {total_proj_avec_ndobo/total_obj*100:.0f}%")
