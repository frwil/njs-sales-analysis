"""
Métriques zéro-achat — Septembre 2026 (bilan définitif au 30/09).

Calcule depuis dataset_2023_2026.csv toutes les valeurs nécessaires à la
section 5.3 « Point de suivi — Septembre 2026 » de build_zero_achat_pdf.py
et remplit la colonne Septembre du tableau de bord.

Règles métier :
- Clients internes (COMPTOIR/PDC/SPC/EMANA) exclus des métriques clients.
- Ventes par catégorie : toutes les lignes du dataset (cohérent avec la perf mensuelle).
- Maïs (M1051) séparé des autres INGREDIENTS.
- Clients S1 = au moins un achat entre janvier et juin 2026.
- Churned risque élevé : jour d'achat moyen S1 ≤ 21, pas d'achat en septembre.
- Churned à surveiller : jour moyen S1 22-25, pas d'achat en septembre.
- Pas inquiétant : jour moyen S1 > 25, pas d'achat en septembre.
- Nouveau client : premier achat en septembre (aucun achat janvier-août).
- Réactivé : aucun achat S1, achat en septembre.
- Bundle = commande (client, jour) contenant soja ET concentrés.

Output : scripts/zero_achat_septembre.json
"""
import json
import pandas as pd

df = pd.read_csv('scripts/dataset_2023_2026.csv', parse_dates=['date'], low_memory=False)
d26 = df[df['date'].dt.year == 2026].copy()

# Clients externes uniquement pour les métriques clients (règle zéro achat)
ext = d26[~d26['interne']].copy()
ext['jour'] = ext['date'].dt.day
ext['mois'] = ext['date'].dt.month

s1 = ext[ext['mois'].between(1, 6)]
jul = ext[ext['mois'] == 7]
aou = ext[ext['mois'] == 8]
sept = ext[ext['mois'] == 9]

clients_s1 = set(s1['client'].dropna())
clients_sept = set(sept['client'].dropna())
clients_avant = set(ext[ext['mois'].between(1, 8)]['client'].dropna())

# --- Ventes par catégorie (toutes lignes du dataset, internes inclus) ---
def cat_tonnes(sub):
    out = {}
    for fam in ['TOURTEAUX', 'CONCENTRES', 'ALIMENT_COMPLET', 'INGREDIENTS',
                'PREMIX', 'COMPLEMENT_ALIMENTAIRE', 'MATERIEL_ELEVAGE']:
        out[fam] = round(float(sub[sub['family'] == fam]['tonnes'].sum()), 1)
    # MAIS = famille dédiée dans le dataset (M1051/M1052, sacs 50 kg)
    mais = sub[sub['family'] == 'MAIS']['tonnes'].sum()
    out['MAIS'] = round(float(mais), 1)
    out['INGREDIENTS_HORS_MAIS'] = round(out['INGREDIENTS'] - out['MAIS'], 1)
    out['TOTAL'] = round(float(sub['tonnes'].sum()), 1)
    return out

t_s1 = cat_tonnes(s1)
t_s1_moy = {k: round(v / 6, 1) if v else 0 for k, v in t_s1.items()}  # moyenne mensuelle S1
t_jul = cat_tonnes(jul)
t_aou = cat_tonnes(aou)
t_sept = cat_tonnes(sept)

# Objectifs septembre : repris du livrable mensuel régénéré (cohérence)
perf = json.load(open('scripts/perf_mensuelle_2026_09.json', encoding='utf-8'))
obj = {}
for fam in ['TOURTEAUX', 'CONCENTRES', 'ALIMENT_COMPLET', 'INGREDIENTS',
            'PREMIX', 'COMPLEMENT_ALIMENTAIRE', 'MATERIEL_ELEVAGE']:
    g = perf['global_sept'].get(fam, {})
    obj[fam] = round(g.get('obj', 0), 1) if g.get('obj') else None

# --- Habitudes d'achat S1 par client ---
jour_moy = s1.groupby('client')['jour'].mean().round(1)
nb_achats = s1.groupby('client').size()
ca_s1_client = s1.groupby('client')['montant_ttc'].sum() / 1e6
# agence du dernier achat S1
dernier_s1 = s1.sort_values('date').groupby('client').tail(1).set_index('client')

churned = clients_s1 - clients_sept
def risque(jm):
    if jm <= 21:
        return 'eleve'
    if jm <= 25:
        return 'surveiller'
    return 'fin_mois'

churned_par_cat = {'eleve': 0, 'surveiller': 0, 'fin_mois': 0}
ca_churned = {'eleve': 0.0, 'surveiller': 0.0, 'fin_mois': 0.0}
for c in churned:
    jm = jour_moy.get(c, 0)
    cat = risque(jm)
    churned_par_cat[cat] += 1
    ca_churned[cat] += ca_s1_client.get(c, 0)

maintenus = clients_s1 & clients_sept
nouveaux = clients_sept - clients_avant
reactives = clients_sept - clients_s1  # dont nouveaux : distinction = achat avant sept ou non
reactives_vrais = reactives - nouveaux

# CA septembre des nouveaux clients
ca_sept_client = sept.groupby('client')['montant_ttc'].sum() / 1e6
ca_gagne_nouveaux = round(float(sum(ca_sept_client.get(c, 0) for c in nouveaux)), 1)

# Top 10 churned à risque élevé (CA S1 desc)
top10 = []
for c in churned:
    if risque(jour_moy.get(c, 0)) == 'eleve':
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
clients_conc_sept = set(sept[sept['family'] == 'CONCENTRES']['client'].dropna())
conc_maintenus = clients_conc_s1 & clients_conc_sept
conc_churned = clients_conc_s1 - clients_conc_sept
conc_nouveaux = clients_conc_sept - set(ext[ext['mois'].between(1, 8) & (ext['family'] == 'CONCENTRES')]['client'].dropna())

# --- Soja / cross-sell ---
clients_soja_sept = set(sept[sept['family'] == 'TOURTEAUX']['client'].dropna())
clients_soja_conc_sept = clients_soja_sept & clients_conc_sept
soja_only = clients_soja_sept - clients_conc_sept
clients_soja_s1 = set(s1[s1['family'] == 'TOURTEAUX']['client'].dropna())

# --- Booster (ALIMENT COMPLET) ---
clients_booster_sept = set(sept[sept['family'] == 'ALIMENT_COMPLET']['client'].dropna())
booster_conc_sept = clients_booster_sept & clients_conc_sept

# --- Bundle : commandes (client, jour) soja avec/sans concentrés ---
sept_soja = sept[sept['family'] == 'TOURTEAUX']
sept_conc = sept[sept['family'] == 'CONCENTRES']
cmd_soja = set(zip(sept_soja['client'], sept_soja['date']))
cmd_conc = set(zip(sept_conc['client'], sept_conc['date']))
cmd_bundle = cmd_soja & cmd_conc
pct_bundle = round(len(cmd_bundle) / len(cmd_soja) * 100, 1) if cmd_soja else None

# Ratio tonnes soja:conc septembre (toutes lignes)
ratio_sc = round(t_sept['TOURTEAUX'] / t_sept['CONCENTRES'], 1) if t_sept['CONCENTRES'] else None

# Prix moyen soja septembre (FCFA / sac 50 kg)
prix_soja = round(float(sept[sept['family'] == 'TOURTEAUX']['montant_ttc'].sum()
                        / sept[sept['family'] == 'TOURTEAUX']['sacs_50'].sum()), 0)

# CA septembre total (toutes lignes, pour cohérence avec la perf)
ca_sept_total = float(d26[d26['date'].dt.month == 9]['montant_ttc'].sum()) / 1e6

# Vérification de cohérence avec la perf mensuelle
t_sept_perf = perf['global_sept']
out = {
    'ventes': {
        's1_moy': t_s1_moy, 'juillet': t_jul, 'aout': t_aou, 'septembre': t_sept,
        'obj_septembre': obj,
    },
    'clients': {
        's1_total': len(clients_s1),
        'maintenus': len(maintenus),
        'maintenus_pct': round(len(maintenus) / len(clients_s1) * 100, 1) if clients_s1 else None,
        'churned_total': len(churned),
        'churned_eleve': churned_par_cat['eleve'],
        'churned_surveiller': churned_par_cat['surveiller'],
        'churned_fin_mois': churned_par_cat['fin_mois'],
        'ca_churned_eleve_m': round(ca_churned['eleve'], 1),
        'ca_churned_surveiller_m': round(ca_churned['surveiller'], 1),
        'ca_churned_fin_mois_m': round(ca_churned['fin_mois'], 1),
        'ca_s1_maintenus_m': round(float(sum(ca_s1_client.get(c, 0) for c in maintenus)), 1),
        'nouveaux': len(nouveaux),
        'ca_gagne_nouveaux_m': ca_gagne_nouveaux,
        'reactives': len(reactives_vrais),
        'top10_churned': top10,
    },
    'concentres': {
        'clients_s1': len(clients_conc_s1),
        'clients_sept': len(clients_conc_sept),
        'maintenus': len(conc_maintenus),
        'maintenus_pct': round(len(conc_maintenus) / len(clients_conc_s1) * 100, 1) if clients_conc_s1 else None,
        'churned': len(conc_churned),
        'nouveaux': len(conc_nouveaux),
        'volume_sept_t': t_sept['CONCENTRES'],
        'obj_sept_t': obj.get('CONCENTRES'),
        'pct_obj': round(t_sept['CONCENTRES'] / obj['CONCENTRES'] * 100, 1) if obj.get('CONCENTRES') else None,
    },
    'cross_sell': {
        'clients_soja_sept': len(clients_soja_sept),
        'soja_conc': len(clients_soja_conc_sept),
        'soja_conc_pct': round(len(clients_soja_conc_sept) / len(clients_soja_sept) * 100, 1) if clients_soja_sept else None,
        'soja_only': len(soja_only),
        'clients_booster_sept': len(clients_booster_sept),
        'booster_conc': len(booster_conc_sept),
        'booster_conc_pct': round(len(booster_conc_sept) / len(clients_booster_sept) * 100, 1) if clients_booster_sept else None,
        'cmd_soja': len(cmd_soja),
        'cmd_bundle': len(cmd_bundle),
        'pct_bundle': pct_bundle,
        'ratio_soja_conc_t': ratio_sc,
    },
    'prix_soja_sac': prix_soja,
    'ca_sept_total_m': round(ca_sept_total, 1),
    'coherence_perf': {
        'sept_t_perf': {k: round(v['t'], 1) for k, v in t_sept_perf.items() if isinstance(v, dict)},
        'sept_t_dataset': t_sept,
    },
}

with open('scripts/zero_achat_septembre.json', 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=2, default=str)

print(json.dumps({
    'ventes sept': t_sept,
    's1_moy': t_s1_moy,
    'juillet': t_jul, 'aout': t_aou,
    'obj': obj,
    'clients': out['clients'],
    'concentres': out['concentres'],
    'cross_sell': out['cross_sell'],
    'prix_soja': prix_soja, 'ca_sept': round(ca_sept_total, 1),
}, ensure_ascii=False, indent=1, default=str))
print("\nSauvegardé : scripts/zero_achat_septembre.json")
