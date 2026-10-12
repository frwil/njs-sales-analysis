"""
Métriques zéro-achat — Octobre 2026 (MTD au 10/10).

Calcule depuis dataset_2023_2026.csv toutes les valeurs nécessaires à la
section 5.4 « Point de suivi — Octobre 2026 (MTD au 10/10) » de
build_zero_achat_pdf.py.

Cadre MTD : au 10/10, la plupart des clients S1 n'ont pas encore atteint leur
jour d'achat habituel. On ne parle donc pas de « churn » mais de
« pas encore venus », segmenté par rapport à leur jour moyen d'achat S1 :
- jour moyen ≤ 10 : auraient dû acheter avant le 10/10 → suivi (risque) ;
- jour moyen 11-25 : jour habituel pas encore atteint → normal ;
- jour moyen > 25 : acheteurs de fin de mois → normal.

Règles métier :
- Clients internes (COMPTOIR/PDC/SPC/EMANA) exclus des métriques clients.
- Ventes par catégorie : toutes les lignes du dataset (cohérent avec la perf mensuelle).
- Clients S1 = au moins un achat entre janvier et juin 2026.
- Nouveau client : premier achat en octobre (aucun achat janvier-septembre).
- Réactivé : aucun achat S1, achat en octobre (hors nouveaux).
- Bundle = commande (client, jour) contenant soja ET concentrés.

Output : scripts/zero_achat_octobre.json
"""
import json
import pandas as pd

df = pd.read_csv('scripts/dataset_2023_2026.csv', parse_dates=['date'], low_memory=False)
d26 = df[df['date'].dt.year == 2026].copy()

ext = d26[~d26['interne']].copy()
ext['jour'] = ext['date'].dt.day
ext['mois'] = ext['date'].dt.month

s1 = ext[ext['mois'].between(1, 6)]
sept = ext[ext['mois'] == 9]
octo = ext[ext['mois'] == 10]
# Septembre MTD (01-10/09) pour comparaison équitable avec octobre MTD
sept_mtd = ext[(ext['mois'] == 9) & (ext['jour'] <= 10)]

clients_s1 = set(s1['client'].dropna())
clients_octo = set(octo['client'].dropna())
clients_avant_oct = set(ext[ext['mois'].between(1, 9)]['client'].dropna())

# --- Ventes par catégorie (toutes lignes du dataset, internes inclus) ---
def cat_tonnes(sub):
    out = {}
    for fam in ['TOURTEAUX', 'CONCENTRES', 'ALIMENT_COMPLET', 'INGREDIENTS',
                'PREMIX', 'COMPLEMENT_ALIMENTAIRE', 'MATERIEL_ELEVAGE']:
        out[fam] = round(float(sub[sub['family'] == fam]['tonnes'].sum()), 1)
    mais = sub[sub['family'] == 'MAIS']['tonnes'].sum()
    out['MAIS'] = round(float(mais), 1)
    out['INGREDIENTS_HORS_MAIS'] = round(out['INGREDIENTS'] - out['MAIS'], 1)
    out['TOTAL'] = round(float(sub['tonnes'].sum()), 1)
    return out

t_s1 = cat_tonnes(s1)
t_s1_moy = {k: round(v / 6, 1) if v else 0 for k, v in t_s1.items()}
t_sept = cat_tonnes(sept)
t_sept_mtd = cat_tonnes(sept_mtd)
t_octo = cat_tonnes(octo)

# Objectifs octobre : repris du livrable mensuel régénéré (cohérence)
perf = json.load(open('scripts/perf_mensuelle_2026_10.json', encoding='utf-8'))
obj = {}
for fam in ['TOURTEAUX', 'CONCENTRES', 'ALIMENT_COMPLET', 'INGREDIENTS',
            'PREMIX', 'COMPLEMENT_ALIMENTAIRE', 'MATERIEL_ELEVAGE']:
    g = perf['global_sept'].get(fam, {})
    obj[fam] = round(g.get('obj', 0), 1) if g.get('obj') else None

# --- Habitudes d'achat S1 par client ---
jour_moy = s1.groupby('client')['jour'].mean().round(1)
nb_achats = s1.groupby('client').size()
ca_s1_client = s1.groupby('client')['montant_ttc'].sum() / 1e6
dernier_s1 = s1.sort_values('date').groupby('client').tail(1).set_index('client')

# S1 pas encore venus en octobre, segmentés par jour moyen d'achat
pas_encore = clients_s1 - clients_octo
def segment(jm):
    if jm <= 10:
        return 'depasse'    # auraient dû acheter avant le 10/10
    if jm <= 25:
        return 'normal'     # jour habituel pas encore atteint
    return 'fin_mois'       # acheteurs de fin de mois

pas_encore_par_cat = {'depasse': 0, 'normal': 0, 'fin_mois': 0}
ca_pas_encore = {'depasse': 0.0, 'normal': 0.0, 'fin_mois': 0.0}
for c in pas_encore:
    cat = segment(jour_moy.get(c, 0))
    pas_encore_par_cat[cat] += 1
    ca_pas_encore[cat] += ca_s1_client.get(c, 0)

venus = clients_s1 & clients_octo
nouveaux = clients_octo - clients_avant_oct
reactives = clients_octo - clients_s1  # dont nouveaux
reactives_vrais = reactives - nouveaux

# CA octobre des nouveaux clients
ca_octo_client = octo.groupby('client')['montant_ttc'].sum() / 1e6
ca_gagne_nouveaux = round(float(sum(ca_octo_client.get(c, 0) for c in nouveaux)), 1)

# Top 10 S1 à jour moyen ≤ 10 pas encore venus (CA S1 desc) — suivi prioritaire
top10 = []
for c in pas_encore:
    if segment(jour_moy.get(c, 0)) == 'depasse':
        ag = dernier_s1.loc[c, 'agence'] if c in dernier_s1.index else '—'
        reg = dernier_s1.loc[c, 'region'] if c in dernier_s1.index else '—'
        top10.append({
            'client': c,
            'agence': ag,
            'region': reg,
            'ca_s1_m': round(float(ca_s1_client.get(c, 0)), 1),
            'jour_moy': float(jour_moy.get(c, 0)),
            'nb_achats': int(nb_achats.get(c, 0)),
        })
top10.sort(key=lambda x: -x['ca_s1_m'])
top10 = top10[:10]

# --- Concentrés ---
clients_conc_s1 = set(s1[s1['family'] == 'CONCENTRES']['client'].dropna())
clients_conc_octo = set(octo[octo['family'] == 'CONCENTRES']['client'].dropna())
conc_venus = clients_conc_s1 & clients_conc_octo
conc_pas_encore = clients_conc_s1 - clients_conc_octo
conc_nouveaux = clients_conc_octo - set(ext[ext['mois'].between(1, 9) & (ext['family'] == 'CONCENTRES')]['client'].dropna())

# --- Soja / cross-sell ---
clients_soja_octo = set(octo[octo['family'] == 'TOURTEAUX']['client'].dropna())
clients_soja_conc_octo = clients_soja_octo & clients_conc_octo
soja_only = clients_soja_octo - clients_conc_octo
clients_soja_s1 = set(s1[s1['family'] == 'TOURTEAUX']['client'].dropna())

# --- Booster (ALIMENT COMPLET) ---
clients_booster_octo = set(octo[octo['family'] == 'ALIMENT_COMPLET']['client'].dropna())
booster_conc_octo = clients_booster_octo & clients_conc_octo

# --- Bundle : commandes (client, jour) soja avec/sans concentrés ---
octo_soja = octo[octo['family'] == 'TOURTEAUX']
octo_conc = octo[octo['family'] == 'CONCENTRES']
cmd_soja = set(zip(octo_soja['client'], octo_soja['date']))
cmd_conc = set(zip(octo_conc['client'], octo_conc['date']))
cmd_bundle = cmd_soja & cmd_conc
pct_bundle = round(len(cmd_bundle) / len(cmd_soja) * 100, 1) if cmd_soja else None

# Ratio tonnes soja:conc octobre (toutes lignes)
ratio_sc = round(t_octo['TOURTEAUX'] / t_octo['CONCENTRES'], 1) if t_octo['CONCENTRES'] else None

# Prix moyen soja octobre (FCFA / sac 50 kg)
prix_soja = round(float(octo[octo['family'] == 'TOURTEAUX']['montant_ttc'].sum()
                        / octo[octo['family'] == 'TOURTEAUX']['sacs_50'].sum()), 0)

# CA octobre total (toutes lignes, pour cohérence avec la perf)
ca_octo_total = float(d26[d26['date'].dt.month == 10]['montant_ttc'].sum()) / 1e6

out = {
    'ventes': {
        's1_moy': t_s1_moy,
        'septembre': t_sept,
        'septembre_mtd': t_sept_mtd,
        'octobre_mtd': t_octo,
        'obj_octobre': obj,
    },
    'clients': {
        's1_total': len(clients_s1),
        'venus': len(venus),
        'venus_pct': round(len(venus) / len(clients_s1) * 100, 1) if clients_s1 else None,
        'pas_encore_total': len(pas_encore),
        'pas_encore_depasse': pas_encore_par_cat['depasse'],
        'pas_encore_normal': pas_encore_par_cat['normal'],
        'pas_encore_fin_mois': pas_encore_par_cat['fin_mois'],
        'ca_depasse_m': round(ca_pas_encore['depasse'], 1),
        'ca_normal_m': round(ca_pas_encore['normal'], 1),
        'ca_fin_mois_m': round(ca_pas_encore['fin_mois'], 1),
        'ca_s1_venus_m': round(float(sum(ca_s1_client.get(c, 0) for c in venus)), 1),
        'nouveaux': len(nouveaux),
        'ca_gagne_nouveaux_m': ca_gagne_nouveaux,
        'reactives': len(reactives_vrais),
        'top10_depasse': top10,
    },
    'concentres': {
        'clients_s1': len(clients_conc_s1),
        'clients_octo': len(clients_conc_octo),
        'venus': len(conc_venus),
        'venus_pct': round(len(conc_venus) / len(clients_conc_s1) * 100, 1) if clients_conc_s1 else None,
        'pas_encore': len(conc_pas_encore),
        'nouveaux': len(conc_nouveaux),
        'volume_octo_t': t_octo['CONCENTRES'],
        'obj_octo_t': obj.get('CONCENTRES'),
        'pct_obj': round(t_octo['CONCENTRES'] / obj['CONCENTRES'] * 100, 1) if obj.get('CONCENTRES') else None,
    },
    'cross_sell': {
        'clients_soja_octo': len(clients_soja_octo),
        'soja_conc': len(clients_soja_conc_octo),
        'soja_conc_pct': round(len(clients_soja_conc_octo) / len(clients_soja_octo) * 100, 1) if clients_soja_octo else None,
        'soja_only': len(soja_only),
        'clients_booster_octo': len(clients_booster_octo),
        'booster_conc': len(booster_conc_octo),
        'booster_conc_pct': round(len(booster_conc_octo) / len(clients_booster_octo) * 100, 1) if clients_booster_octo else None,
        'cmd_soja': len(cmd_soja),
        'cmd_bundle': len(cmd_bundle),
        'pct_bundle': pct_bundle,
        'ratio_soja_conc_t': ratio_sc,
    },
    'prix_soja_sac': prix_soja,
    'ca_octo_total_m': round(ca_octo_total, 1),
}

with open('scripts/zero_achat_octobre.json', 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=2, default=str)

print(json.dumps({
    'ventes octo': t_octo,
    'sept_mtd': t_sept_mtd,
    's1_moy': t_s1_moy,
    'obj': obj,
    'clients': out['clients'],
    'concentres': out['concentres'],
    'cross_sell': out['cross_sell'],
    'prix_soja': prix_soja, 'ca_octo': round(ca_octo_total, 1),
}, ensure_ascii=False, indent=1, default=str))
print("\nSauvegardé : scripts/zero_achat_octobre.json")
