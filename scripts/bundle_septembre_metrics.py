"""
Métriques bundle soja→concentrés — Septembre 2026 (bilan définitif au 30/09).

Calcule toutes les valeurs de la section 9 « Point de suivi — Septembre 2026 »
de build_bundle_analysis_pdf.py, avec des sources cohérentes avec les livrables
déjà publiés :

- Volumes officiels septembre = perf_mensuelle_2026_09.json (Livrée+Validée+
  En cours, internes rattachés aux agences classiques inclus) — identiques au
  livrable perf mensuelle et à la section 5.3 zéro-achat.
- Métriques commandes (bundle %, ratio) = dataset Livrée, clients externes —
  identiques à la section 5.3 zéro-achat (source unique).
- Distribution des ratios par commande bundle = dataset Livrée, clients externes.
- Colonne de comparaison Août définitif = même méthode que septembre.

Output : scripts/bundle_septembre.json
"""
import json
import pandas as pd

# --- Sources ---
za = json.load(open('/home/z/my-project/scripts/zero_achat_septembre.json', encoding='utf-8'))
perf_sept = json.load(open('/home/z/my-project/scripts/perf_mensuelle_2026_09.json', encoding='utf-8'))
gs = perf_sept['global_sept']

d26 = pd.read_csv('/home/z/my-project/scripts/dataset_2023_2026.csv',
                  parse_dates=['date'], low_memory=False)


def mois_metrics(mois):
    """Commandes soja/conc + distribution des ratios bundle (dataset Livrée, ext)."""
    ext = d26[(d26['date'].dt.year == 2026) & (d26['date'].dt.month == mois)
              & (~d26['interne'])].copy()
    ext['jour'] = ext['date'].dt.date
    soj = ext[ext['family'] == 'TOURTEAUX']
    con = ext[ext['family'] == 'CONCENTRES']
    cmd_s = set(zip(soj['client'], soj['jour']))
    cmd_c = set(zip(con['client'], con['jour']))
    cmd_b = cmd_s & cmd_c
    dist = []
    for (c, d) in cmd_b:
        s = soj[(soj['client'] == c) & (soj['jour'] == d)]['sacs_50'].sum()
        k = con[(con['client'] == c) & (con['jour'] == d)]['sacs_50'].sum()
        if k:
            dist.append(s / k)
    buckets = {'<=3': 0, '3-5': 0, '5-10': 0, '10-20': 0, '>20': 0}
    for r_ in dist:
        if r_ <= 3: buckets['<=3'] += 1
        elif r_ <= 5: buckets['3-5'] += 1
        elif r_ <= 10: buckets['5-10'] += 1
        elif r_ <= 20: buckets['10-20'] += 1
        else: buckets['>20'] += 1
    total = len(dist)
    pct = {k: round(v / total * 100, 1) if total else 0 for k, v in buckets.items()}
    return {
        't_soja': round(float(soj['tonnes'].sum()), 0),
        't_conc': round(float(con['tonnes'].sum()), 0),
        'ratio_t': round(float(soj['tonnes'].sum() / con['tonnes'].sum()), 1) if len(con) else None,
        'cmd_soja': len(cmd_s),
        'cmd_bundle': len(cmd_b),
        'cmd_soja_only': len(cmd_s - cmd_c),
        'pct_bundle': round(len(cmd_b) / len(cmd_s) * 100, 1) if cmd_s else None,
        'prix_soja_sac': round(float(soj['montant_ttc'].sum() / soj['sacs_50'].sum()), 0),
        'ratio_dist': {'total': total, **buckets, 'pct': pct},
    }


aout = mois_metrics(8)
sept_dist = mois_metrics(9)['ratio_dist']

out = {
    'septembre': {
        # Volumes officiels = perf mensuelle septembre (Livrée+Validée+En cours)
        't_soja': round(gs['TOURTEAUX']['t'], 1),
        't_conc': round(gs['CONCENTRES']['t'], 1),
        'ratio_t': round(gs['TOURTEAUX']['t'] / gs['CONCENTRES']['t'], 1),
        'pct_obj_soja': round(gs['TOURTEAUX']['pct'], 0),
        'pct_obj_conc': round(gs['CONCENTRES']['pct'], 0),
        # Métriques commandes = zéro-achat (dataset Livrée, clients externes)
        'cmd_soja': za['cross_sell']['cmd_soja'],
        'cmd_bundle': za['cross_sell']['cmd_bundle'],
        'cmd_soja_only': za['cross_sell']['cmd_soja'] - za['cross_sell']['cmd_bundle'],
        'pct_bundle': za['cross_sell']['pct_bundle'],
        'prix_soja_sac': za['prix_soja_sac'],
        'ratio_dist': sept_dist,
    },
    'objectifs_sept': {
        't_soja': round(gs['TOURTEAUX']['obj'], 1),
        't_conc': round(gs['CONCENTRES']['obj'], 1),
    },
    'aout_final': aout,
}

with open('/home/z/my-project/scripts/bundle_septembre.json', 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

print(json.dumps(out, ensure_ascii=False, indent=1))
print("\nSauvegardé : scripts/bundle_septembre.json")
