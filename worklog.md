---
Task ID: soja-juillet-23
Agent: main
Task: Mettre a jour le fichier analyse_soja_juillet.xlsx avec la nouvelle extraction "Les ventes a l'instant" (au 23/07/2026).

Work Log:
- Identifie le fichier source le plus recent: `NJS GROUP ERP - Lignes de commandes + multicompany (3) (6).xlsx` (modifie 23/07 16:47).
- Verifie la couverture: donnees Livree du 01/07 au 23/07/2026 (8 464 lignes Livree).
- Script persistant: `/home/z/my-project/scripts/update_soja_juillet.py` (sauvegarde + re-execution).
- Reproduit la methodologie des 4 tableaux existants (T1 moy/jour, T2 commandes du jour, T3 soja-only, T4 stock rupture) avec colonnes d'evolution vs 22/07.
- Calcule le ratio global en sacs 50kg-equivalent (total soja / total conc, toutes commandes incluant conc-only) pour matcher la convention precedente (2.5:1 le 22/07).
- T4: utilise les jours ouvrables (lun-sam) pour le calcul de date de rupture, en ajoutant n business days via `add_business_days()`.

Stage Summary:
- T1 (au 23/07, 20j): 4 851,4 t soja vendu, moy/jour 242,6 t (vs 244,2 t/j au 22/07 sur 19j, soit -1,7 t/j). Baisse legere, dans la marge hebdomadaire.
- T2 (23/07): 75 commandes soja, 58 bundle, 17 soja-only, ratio 2,6:1 (vs 2,5:1 le 22/07). Chute de 17 commandes (jour de hausse de prix).
- T3 (23/07): 17 commandes soja-only pour 5,7 t (vs 22 cmdes, 10,3 t le 22/07). Amelioration: -5 cmdes, -4,6 t.
- T4: 33 jours de stock (lun-sam), rupture probable 31/08/2026 (vs 29/08 le 22/07 — decalage de 2 jours calendar dus au décalage du depart).
- Distribution des ratios bundle: 79% <= 3:1, 19% 3-5:1, 2% 5-10:1 (vs 66%/30%/4% le 22/07) — amelioration de la distribution.
- Output: `/home/z/my-project/download/analyse_soja_juillet.xlsx` (4 feuilles, 11 Ko).
- Summary JSON: `/home/z/my-project/scripts/soja_juillet_23.json`.
- Conclusion: Le 23/07 (1er jour d'application de la hausse +2 000 XAF/sac) montre une baisse des commandes soja (-17 cmdes) et une legere degradation du ratio global (2,5 -> 2,6:1), mais une amelioration de la distribution des ratios et du nombre de soja-only. Effet plein a confirmer sur les prochains jours.

---
Task ID: zero-achat-bundle-24
Agent: main
Task: Mettre a jour l'analyse zero achat et les fichiers d'analyse soja concentres avec la nouvelle extraction au 24/07/2026.

Work Log:
- Identifie la nouvelle extraction: `NJS GROUP ERP - Lignes de commandes + multicompany (3) (7).xlsx` (modifiee 25/07 07:20, couverture 01-24/07).
- Script persistant: `/home/z/my-project/scripts/compute_july_mtd_metrics.py` — calcule les metriques July MTD (zero achat, cross-sell, volumes, bundle) a partir de S1 + July.
- Corrige les poids des ingredients (B100=25kg, E101=25kg, I105=25kg, etc.) qui etaient defaults a 50kg.
- Separe le Mais des autres ingredients (M1051 isole, conforme a la regle "toujours separer le mais").
- Script `update_soja_juillet.py` mis a jour pour pointer vers la nouvelle extraction (7) et utiliser 24/07 comme date de reference.
- `build_zero_achat_pdf.py` (1228 lignes) — remplacement cible des valeurs dans la section 5 (Suivi S2):
  - Date "21/07" -> "24/07" partout
  - "21 jours / 68%" -> "24 jours / 78%"
  - Habitudes d'achat: 721->775 maintenus, 440->497 risque eleve, 157->56 a surveiller, 35->47 nouveaux
  - CA: 1819->1851 M risque eleve, 483->116 M a surveiller
  - Top 10 risque: remplacement complet (TCHEUTCHOUA en tete, 5/10 FAMLA)
  - Cross-sell soja: 113->82 regression, 15->30 conversion, 197->164 soja-only
  - Cross-sell booster: 16->9 regression, 2->8 conversion
  - Volumes: TOURTEAUX 4910/7249 -> 5203/6689, CONCENTRES 1021/1507 -> 1240/1594
  - Note prix: ajout de la seconde hausse +2 000 XAF/sac effective 23/07 (en plus de la premiere +1 000)
- `build_bundle_analysis_pdf.py` (1027 lignes) — remplacement cible:
  - Date "18/07" -> "24/07"
  - "+1 000 FCFA/sac" -> "deux hausses successives +1 000 puis +2 000 au 23/07 = +3 000 cumule"
  - Base clients: 1100->958 S1, 492->540 Juil
  - Tableau global: prix soja 16184->16198 (S1), 18092->18322 (Juil); budget +64%->+32%; sacs soja/client 88.6->132.7; sacs conc/client 29.4->49.3; ratio 2.5->2.7 (S1) et 3.3->3.5 (Juil)
  - Tableau par agence (14 agences) — toutes les valeurs mises a jour
  - Tableaux par region (Ouest, Centre, Littoral) — recalcules par aggregation
  - Focus NDOBO: ratio 4.6:1 -> 5.5:1, conc -42% -> -39%, budget -5% -> +6%
  - Recommandations: ajustement des references aux nouvelles valeurs

Stage Summary:
- analyse_soja_juillet.xlsx (4 feuilles, 11 Ko):
  - T1: 247.7 t/jour (vs 242.6 au 23/07, +5.2 t/j) — la hausse de prix n'a pas encore ralenti la consommation
  - T2 (24/07): 70 cmdes, ratio 1.7:1 (vs 2.6:1 le 23/07) — belle amelioration du bundle
  - T3: 8 cmdes soja-only (9.2 t) vs 17 cmdes (5.7 t) le 23/07 — moins de soja-only en volume moyen
  - T4: rupture probable 31/08/2026 (32 jours de stock, lun-sam)
- analyse_zero_achat.pdf (21 pages, 334 Ko): section 5 entierement refraichie avec donnees au 24/07
  - 775 clients maintenus (57%), 497 a risque eleve, 47 nouveaux
  - Cross-sell soja->conc: 79% -> 77% (-2 pts), 82 regressions, 30 conversions
  - Projection CONCENTRES juillet: 1594 t (+5% vs S1) — amelioration timide vs debut juillet
  - Note prix: double hausse +1 000 (debut mois) + +2 000 (23/07) = +3 000 cumule
- analyse_bundle_soja_concentres.pdf (14 pages, 78 Ko): entierement refraichi
  - Global: 958 -> 540 clients bundle, ratio 2.7:1 -> 3.5:1
  - 5 agences en verdict rouge (NDOBO 5.5:1, DJELENG 4.1:1, NKOABANG 3.8:1, NKOLBISSON, NKONGSAMBA)
  - 4 agences en verdict vert (BERTOUA, NGAOUNDERE, BUEA, AHALA)
  - Littoral region la plus touchee: ratio 2.8 -> 3.8, conc -27%
  - Centre meilleure region: ratio 2.4 -> 3.0, conc -13%
- Outputs JSON de traceabilite: juillet_mtd_24.json, cross_sell_mtd_24.json, volumes_juillet_24.json, bundle_metrics_24.json, soja_juillet_24.json
- Conclusion generale: l'effet plein de la hausse +2 000 XAF (effective 23/07) n'est pas encore visible sur 24j (seulement 2 jours a +3 000 XAF cumule). La consommation soja continue de progresser (+5 t/j) — la fenetre d'opportunité (rupture concurrente) l'emporte encore sur l'effet prix. Le ratio bundle s'ameliore le 24/07 (1.7:1) mais la tendance mensuelle reste degradee (3.5:1 vs 2.7:1 en S1).
