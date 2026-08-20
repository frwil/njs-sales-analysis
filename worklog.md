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

---
Task ID: zero-achat-bundle-27
Agent: main
Task: Mettre a jour l'analyse zero achat et bundle soja-concentres avec la nouvelle extraction au 27/07/2026. Ajouter la regle premix (1 sac/tonne soja pour commandes soja-only).

Work Log:
- Identifie la nouvelle extraction: `NJS GROUP ERP - Lignes de commandes + multicompany (4).xlsx` (modifiee 27/07 08:09, couverture 01-27/07, 9337 lignes Livree).
- Le 27/07 est une extraction matinale (85 lignes, 0 commande soja Livree) — utilise le 25/07 (samedi, 66 cmdes soja) pour T2/T3.
- Temps ecoule: 23 jours ouvres (lun-sam) sur 27 = 85.2% du mois.
- Script `update_soja_juillet.py` mis a jour pour pointer vers la nouvelle extraction (4) et utiliser 27/07 comme date de reference.
- Script `compute_july_mtd_metrics.py` mis a jour pour la nouvelle extraction. Outputs renommes en *_27.json.
- `build_zero_achat_pdf.py` — section 5 refraichie avec donnees au 27/07:
  - 789 clients maintenus (58%), 485 a risque eleve, 48 nouveaux
  - CA risque eleve: 1 772 M FCFA
  - Cross-sell soja: 82 regressions, 33 conversions, taux 77% (-2 pts vs S1)
  - Cross-sell booster: 9 regressions, 9 conversions, taux 71% (stable)
  - Volumes: TOURTEAUX 5465/6415 t (proj +32% vs S1, mais moy/jour -4% vs 24/07 — effet prix commence a se manifester)
  - CONCENTRES: 1322/1551 t (proj +3% vs S1, 80% ajuste)
  - Note prix: double hausse +1 000 (debut mois) + +2 000 (23/07) = +3 000 FCFA/sac cumule
- `build_bundle_analysis_pdf.py` — mis a jour avec donnees au 27/07:
  - Global: 958 -> 557 clients bundle, ratio 2.7:1 -> 3.5:1, budget +37%, sacs soja +10%, sacs conc -16%
  - 14 agences mises a jour (NDOBO 5.5:1, PK11 4.7:1, DJELENG 4.1:1 toujours critiques)
  - Regions: Ouest ratio 2.8->3.7, Centre 2.4->3.0, Littoral 2.8->3.8
  - Conclusions ajustees avec nouvelles valeurs
- NOUVELLE SECTION 7.5 ajoutee — Regle premix pour commandes soja-only:
  - Regle: 1 sac de premix (25 kg) minimum par tonne de soja pour les commandes soja-only sans concentre
  - Impact analyse: 1087 commandes soja-only sans premix en juillet (39% des commandes soja)
  - Volume soja concerne: 2860 t -> 2860 sacs premix requis (71.5 t)
  - CA potentiel: 71.5 M FCFA/mois, 858 M FCFA annualise
  - Top 10 commandes (SEPTENTRION VETERINAIRE 165 t cumulees, SAGUEN DEFFO 125 t)
  - Recommandation 10 ajoutee au tableau des recommandations (10 axes au lieu de 9)
  - Synergie avec regle 7.3 (annulation >= 10 t sans conc) mise en evidence

Stage Summary:
- analyse_soja_juillet.xlsx (4 feuilles, 11 Ko): au 27/07, moy/jour 237.6 t (-10 t/j vs 24/07), rupture probable 04/09/2026 (34j stock)
- analyse_zero_achat.pdf (21 pages, 334 Ko): section 5 entierement refraichie au 27/07
  - 789 clients maintenus (58%), 485 a risque eleve (CA 1772 M)
  - Projection CONCENTRES juillet: 1551 t (+3% vs S1, 80% ajuste)
  - Effet prix commence a se manifester: moy/jour TOURTEAUX -4% vs 24/07
- analyse_bundle_soja_concentres.pdf (15 pages, 84 Ko): entierement refraichi + nouvelle section 7.5
  - Global: 557 clients bundle, ratio 3.5:1, budget +37%
  - 5 agences rouges (NDOBO 5.5:1, DJELENG 4.1:1, NKOABANG 3.8:1, NKOLBISSON, NKONGSAMBA)
  - 4 agences vertes (BERTOUA, NGAOUNDERE, BUEA, AHALA)
  - NOUVEAU: Section 7.5 regle premix — potentiel CA additionnel 71.5 M FCFA/mois
- Outputs JSON: juillet_mtd_27.json, cross_sell_mtd_27.json, volumes_juillet_27.json, bundle_metrics_27.json, soja_juillet_27.json
- Conclusion: la hausse +2 000 XAF/sac (effective 23/07) commence a ralentir la consommation soja (-4% en 3 jours). La fenetre d'opportunité (rupture concurrente) commence a se refermer. La nouvelle regle premix (1 sac/tonne pour soja-only) represente un levier immediate de 71.5 M FCFA/mois.

---
Task ID: aout-update-15
Agent: main
Task: Mettre a jour l'analyse zero achat, bundle soja-concentres et soja_aout avec la nouvelle extraction au 14/08/2026. Fournir le point performance CONCENTRES d'aout dans la conversation.

Work Log:
- Identifie la nouvelle extraction: `NJS GROUP ERP - Lignes de commandes + multicompany (5) (1).xlsx` (modifiee 15/08 10:06, 5 087 lignes Livree).
- Couverture: 01/08 au 15/08. Le 15/08 est une extraction matinale (22 lignes Livree) — exclue comme partielle. Donnees utilisees: 01/08 au 14/08 (5 065 lignes Livree externes).
- 12 jours ouvres (lun-sam) sur 26 (46% du mois).
- Script `inspect_new_extraction.py` — verification couverture, header sur row 2 (16 colonnes, format identique au precedent).
- Script `compute_aout_mtd_metrics.py` — calcul des metriques August MTD:
  - TOURTEAUX: 1 840 t (moy 153 t/j, proj 3 986 t, obj 3 850 -> 104% ajuste)
  - CONCENTRES: 717 t (moy 59,7 t/j, proj 1 553 t, obj 1 480 -> 105% ajuste)
  - Bundle ratio: 2,6:1 (vs 3,5:1 juillet) — retour sous cible 3:1
  - Cross-sell: 95% bundle (880/924 cmds soja), 44 soja-only (5%)
  - Dist ratios: 65% <= 3:1, 33% 3-5:1, 2% >5:1
  - Stock BEKOKO net: 71 230 sacs (3 561 t), rupture probable 04/09/2026 (21 jours)
  - Churned: 405 clients S1 sans achat aout a 12j (vs 993 au 04/08)
- Script `update_soja_aout.py` mis a jour pour la nouvelle extraction (5)(1):
  - SRC -> (5) (1).xlsx
  - end_date -> 14/08/2026 (12j)
  - Output: /home/z/my-project/download/analyse_soja_aout.xlsx (5 feuilles T1-T5, 16 Ko)
- Script `build_zero_achat_pdf.py` mis a jour:
  - Section 5.2 "Premier point de suivi — Août 2026 (au 04/08)" -> "Point de suivi — Août 2026 (au 14/08)"
  - Tableau evol_aout: complete avec juillet et moyenne/jour
  - Dashboard: mise a jour Août (12j) avec nouveaux KPIs (95% bundle, 2,6:1 ratio, 71 230 sacs, 04/09 rupture)
  - Bilan intermediaire mi-aout: TOURTEAUX effet prix visible (-45% vs juillet), CONCENTRES dynamique positive, bundle ameliore
  - Output: /home/z/my-project/download/analyse_zero_achat.pdf (21 pages, 337 Ko)
- Script `build_bundle_analysis_pdf.py` mis a jour:
  - Section 8 "Premier point de suivi — Août 2026 (au 04/08)" -> "Point de suivi — Août 2026 (au 14/08)"
  - Tableau 8.1: ajout lignes projection, stock, commandes detail
  - NOUVEAU Tableau 8.2: distribution des ratios bundle (65% <= 3:1, 33% 3-5:1, 2% >5:1)
  - Lecture mi-aout complete avec 6 points cles
  - Output: /home/z/my-project/download/analyse_bundle_soja_concentres.pdf (17 pages, 90 Ko)
- Outputs JSON de traceabilite: aout_mtd_15.json, soja_aout_15.json

Stage Summary:
- analyse_soja_aout.xlsx (5 feuilles T1-T5, 16 Ko): au 14/08, moy/j soja 153 t (vs 158 t/j au 13/08), rupture probable 04/09/2026 (21j stock net)
- analyse_zero_achat.pdf (21 pages, 337 Ko): section 5.2 entierement refraichie au 14/08
  - 514 clients maintenus (37% du S1 actif), 95 nouveaux, 405 churned (a 12j, vs 993 au 04/08)
  - Projection CONCENTRES août: 1 553 t (105% ajuste)
  - Projection TOURTEAUX août: 3 986 t (104% ajuste)
  - Bundle ratio 2,6:1 (vs 3,5:1 juillet) — retour sous cible 3:1
  - Stock soja critique: 71 230 sacs net (3 561 t), rupture probable 04/09/2026
- analyse_bundle_soja_concentres.pdf (17 pages, 90 Ko): section 8 entierement refraichie
  - Tableau de bord global: ratio 2,6:1, 95% bundle, 880/924 cmds soja en bundle
  - NOUVEAU tableau distribution des ratios: 65% <= 3:1, 33% 3-5:1, 2% >5:1
  - Conclusion mi-aout: plan d'action bundle produit ses effets, CONCENTRES tiennent le cap
- Conclusion generale: la combinaison hausse tarifaire (+3 000 FCFA/sac cumul) + plan d'action bundle produit la trajectoire attendue. Les CONCENTRES sont en passe d'atteindre l'objectif pour le 2e mois consecutif (105% ajuste vs 106% en juillet). Le ratio bundle s'est nettement ameliore (2,6:1 vs 3,5:1). Le volume TOURTEAUX se normalise apres le pic de juillet (153 t/j vs 280 t/j, -45%) — effet prix visible. Le stock soja est critique (rupture probable 04/09) — reapprovisionnement a programmer avant le 25/08.

---
Task ID: aout-update-16
Agent: main
Task: Mettre a jour l'analyse zero achat, bundle soja-concentres et soja_aout avec la nouvelle extraction au 15/08/2026. Fournir le point performance CONCENTRES d'aout dans la conversation.

Work Log:
- Identifie la nouvelle extraction: `NJS GROUP ERP - Lignes de commandes + multicompany (3) (14) (1).xlsx` (modifiee 16/08 20:47, 5 364 lignes Livree).
- Couverture: 01/08 au 15/08. Le 15/08 a 217 lignes Livree (vs 22 precedemment) — c'est maintenant une journee complete (plus matinal). Donnees utilisees: 01/08 au 15/08 (13 jours ouvres lun-sam, 50% du mois).
- Scripts mis a jour:
  - `update_soja_aout.py` — SRC -> (3) (14) (1).xlsx, end_date -> 15/08/2026, output JSON: soja_aout_16.json
  - `compute_aout_mtd_metrics.py` — AOUT_SRC -> nouveau fichier, days_elapsed -> 1-15 (13j), rupture_date depart -> 15/08, output: aout_mtd_16.json
  - `build_zero_achat_pdf.py` — Section 5.2 refraichie au 15/08: 13j, 50% du mois
  - `build_bundle_analysis_pdf.py` — Section 8 refraichie au 15/08
- Metriques August MTD (au 15/08, 13j):
  - TOURTEAUX: 1 933 t (moy 148,7 t/j, proj 3 866 t, obj 3 850 -> 100% ajuste, vs 4 844 t S1 -20%)
  - CONCENTRES: 752 t (moy 57,8 t/j, proj 1 504 t, obj 1 480 -> 102% ajuste, vs 1 513 t S1 -1%)
  - Bundle ratio: 2,6:1 (stable vs 14/08, vs 3,5:1 juillet)
  - Cross-sell: 95% bundle (918/968 cmds soja), 50 soja-only (5%)
  - Dist ratios: 65% <= 3:1, 33% 3-5:1, 2% >5:1 (stable)
  - Stock BEKOKO net: 71 230 sacs (3 561 t), rupture probable 05/09/2026 (21 jours)
  - Churned: 388 clients S1 sans achat aout a 13j (vs 405 au 14/08, 993 au 04/08 — chute continue)
  - Nouveaux clients aout: 98 (vs 95 au 14/08)
  - Cross-sell: 425 clients soja S1 avec conc en aout (53% des 804 soja S1)
- Performance CONCENTRES par region:
  - Ouest: 619 / 605 = 102% (+14 t) — DJELENG 158% (+65 t) tire la region, FAMLA 82% (-69 t) sous-performe
  - Centre: 509 / 491 = 104% (+18 t) — 4/6 agences au-dessus
  - Littoral: 375 / 438 = 86% (-63 t) — NDOBO 61% (-77 t) plombe la region (sans NDOBO -> 106%)
  - TOTAL: 1 504 / 1 535 = 98% (-31 t) — 9/14 agences au-dessus de l'objectif
  - Si NDOBO retrouve obj (+77 t): total -> 103% (1 581 t / 1 535 t)
- Tendance vs precedent point (14/08):
  - CONCENTRES projection: 1 553 -> 1 504 t (-49 t, ajuste -3 pts)
  - TOURTEAUX projection: 3 986 -> 3 866 t (-120 t, ajuste -4 pts)
  - Bundle ratio: 2,6:1 stable
  - Churned: 405 -> 388 (-17), nouveaux 95 -> 98 (+3)

Stage Summary:
- analyse_soja_aout.xlsx (5 feuilles, 16 Ko): au 15/08, moy/j soja 148,7 t (vs 153,3 au 14/08, -4,6 t/j), rupture probable 05/09/2026 (21j stock net)
- analyse_zero_achat.pdf (21 pages, 337 Ko): section 5.2 refraichie au 15/08
  - 531 clients S1 actifs (37%), 98 nouveaux, 388 churned (a 13j, vs 993 au 04/08)
  - Projection CONCENTRES août: 1 504 t (102% ajuste) — dynamique stable
  - Projection TOURTEAUX août: 3 866 t (100% ajuste) — effet prix plein (-47% vs juillet)
  - Bundle ratio 2,6:1 (vs 3,5:1 juillet) — retour sous cible 3:1
  - Stock soja critique: 71 230 sacs net (3 561 t), rupture probable 05/09/2026
- analyse_bundle_soja_concentres.pdf (17 pages, 90 Ko): section 8 refraichie au 15/08
  - Tableau de bord global: ratio 2,6:1, 95% bundle, 918/968 cmds soja en bundle
  - Distribution des ratios: 65% <= 3:1, 33% 3-5:1, 2% >5:1
  - Conclusion mi-aout: plan d'action bundle produit ses effets, CONCENTRES tiennent le cap (102% ajuste)
- Outputs JSON: aout_mtd_16.json, soja_aout_16.json
- Conclusion generale: la dynamique se stabilise à mi-août. Les CONCENTRES restent au-dessus de l'objectif (102% ajuste) malgre la baisse du soja. L'effet prix plein sur les TOURTEAUX se confirme (-47% vs juillet), la fenêtre de rupture concurrente se referme. Le ratio bundle s'est ameliore (2,6:1 vs 3,5:1 en juillet) — le plan d'action bundle produit ses effets. Le stock soja reste critique (rupture probable 05/09) — reapprovisionnement a programmer avant le 25/08.

---
Task ID: aout-update-18
Agent: main
Task: Mettre a jour l'analyse zero achat, bundle soja-concentres et soja_aout avec la nouvelle extraction au 17/08/2026.

Work Log:
- Identifie la nouvelle extraction: `NJS GROUP ERP - Lignes de commandes + multicompany (18).xlsx` (modifiee 18/08 11:12, 5 758 lignes Livree).
- Couverture: 01/08 au 18/08. Le 18/08 a 5 lignes Livree (extraction matinale) — exclue. Donnees utilisees: 01/08 au 17/08 (14 joursouvres lun-sam, 54% du mois).
- Metriques August MTD (au 17/08, 14j):
  - TOURTEAUX: 2 052 t (moy 146,6 t/j, proj 3 812 t, obj 3 850 -> 99% ajuste, vs 4 844 t S1 -21%)
  - CONCENTRES: 806 t (moy 57,6 t/j, proj 1 496 t, obj 1 480 -> 101% ajuste)
  - Bundle ratio: 2,5:1 (vs 2,6:1 au 15/08, vs 3,5:1 juillet) — amelioration continue
  - Cross-sell: 95% bundle (977/1 029 cmds soja), 52 soja-only (5%)
  - Dist ratios: 66% <= 3:1, 32% 3-5:1, 2% >5:1
  - Stock BEKOKO net: 71 230 sacs (3 561 t), rupture probable 08/09/2026 (22 jours)
  - Churned: 372 clients S1 sans achat aout a 14j (vs 388 au 15/08, 993 au 04/08)
  - Nouveaux clients: 102 (vs 98 au 15/08)
  - Cross-sell: 434 clients soja S1 avec conc en aout (54%)
- Performance CONCENTRES par region (au 17/08):
  - Centre: 505 / 491 = 103% (+14 t)
  - Ouest: 597 / 605 = 99% (-8 t) — FAMLA 80% (-78 t) plombe
  - Littoral: 394 / 438 = 90% (-44 t) — NDOBO 68% (-63 t) ameliore
  - TOTAL: 1 496 / 1 535 = 98% (-38 t) — 7/14 agences au-dessus (regression vs 9/14 au 15/08)

Stage Summary:
- analyse_soja_aout.xlsx (5 feuilles, 16 Ko): au 17/08, moy/j soja 146,6 t, rupture probable 08/09/2026
- analyse_zero_achat.pdf (22 pages, 338 Ko): section 5.2 refraichie au 17/08
- analyse_bundle_soja_concentres.pdf (17 pages, 90 Ko): section 8 refraichie au 17/08
- Conclusion: leffet prix +3 000 FCFA/sac produit son plein effet (TOURTEAUX -48% vs juillet), la fenetre de rupture concurrente est refermee. Les CONCENTRES tiennent le cap (101% ajuste) malgre la baisse generale. Bundle 2,5:1 — amelioration continue. Stock soja critique (rupture 08/09).

---
Task ID: aout-corr-1534
Agent: main
Task: Correction de l'objectif CONCENTRES août (1480 -> 1534 t) suite à retour utilisateur.

Work Log:
- Correction de l'objectif CONCENTRES août dans compute_aout_mtd_metrics.py: 1480 -> 1534 t (somme des objectifs agence, deja utilisee dans compute_concentres_regional.py).
- Recalcul des metriques August MTD (au 17/08, 14j):
  - CONCENTRES: proj 1 496 t / obj 1 534 t = 98% ajuste (au lieu de 101% annonce precedemment)
  - TOURTEAUX: proj 3 812 t / obj 3 850 t = 99% ajuste (inchangé)
  - Bundle ratio 2,5:1, cross-sell 95% (inchangés)
- Mise a jour des PDFs:
  - build_zero_achat_pdf.py: tableau evol_aout (obj CONCENTRES 1534, % 98%), dashboard (obj CONCENTRES 1534, % 98%), lecture et bilan intermediaire ajustes
  - build_bundle_analysis_pdf.py: tableau 8.1 (CONCENTRES 98% obj), projection CONCENTRES ajustee, conclusion mi-aout ajustee
- Output: analyse_zero_achat.pdf (22 pages, 338 Ko), analyse_bundle_soja_concentres.pdf (17 pages, 90 Ko)

Stage Summary:
- Correction importante: l'objectif CONCENTRES août est 1 534 t (et non 1 480 t utilise par erreur dans le script principal).
- Performance reelle au 17/08: CONCENTRES a 98% de l'objectif (projection 1 496 t vs 1 534 t), ecart -38 t.
- Sans rattrapage FAMLA (-78 t) et NDOBO (-63 t), l'objectif sera manque de ~38 t (-2,5%).
- Si FAMLA et NDOBO rattrapent partiellement (50% de leur ecart), l'objectif sera atteint.

---
Task ID: aout-update-20
Agent: main
Task: Mettre a jour analyse zero achat, bundle soja-concentres et soja_aout avec la nouvelle extraction au 19/08/2026.

Work Log:
- Identifie la nouvelle extraction: `NJS GROUP ERP - Lignes de commandes + multicompany (19).xlsx` (modifiee 20/08 10:38, 7 225 lignes Livree).
- Couverture: 01/08 au 20/08. Le 20/08 a 67 lignes Livree (extraction matinale) — exclue. Donnees utilisees: 01/08 au 19/08 (16 jours ouvres lun-sam, 62% du mois).
- Metriques August MTD (au 19/08, 16j):
  - TOURTEAUX: 2 683 t (moy 167,7 t/j, proj 4 360 t, obj 3 850 -> 113% ajuste, vs 4 844 t S1 -10%)
  - CONCENTRES: 1 064 t (moy 66,5 t/j, proj 1 728 t, obj 1 534 -> 113% ajuste, +194 t d'avance)
  - Bundle ratio: 2,5:1 (stable vs 17/08)
  - Cross-sell: 95% bundle (1 271/1 331 cmds soja), 60 soja-only (5%)
  - Dist ratios: 69% <= 3:1 (+3 pts), 29% 3-5:1, 2% >5:1
  - Stock BEKOKO net: 71 230 sacs (3 561 t), rupture probable 07/09/2026 (19 jours, vs 22j au 17/08 — conso accelere)
  - Churned: 320 clients S1 sans achat aout a 16j (vs 372 au 17/08, 993 au 04/08)
  - Nouveaux clients: 125 (vs 102 au 17/08)
  - Cross-sell: 490 clients soja S1 avec conc en aout (61%)
- Performance CONCENTRES par region (au 19/08):
  - Ouest: 724 / 605 = 120% (+119 t) — DJELENG 186% (+96 t) tire la region, FAMLA a 96% (-14 t) a rattrape
  - Centre: 538 / 491 = 109% (+46 t) — 5/6 agences au-dessus, MESSASSI a 100%
  - Littoral: 467 / 438 = 107% (+29 t) — NDOBO a 97% (-6 t) a rattrape (vs 68% au 17/08)
  - TOTAL: 1 728 / 1 535 = 113% (+194 t) — 11/14 agences au-dessus (vs 7/14 au 17/08)

Stage Summary:
- analyse_soja_aout.xlsx (5 feuilles, 16 Ko): au 19/08, moy/j soja 167,7 t (vs 147 au 17/08, +14%), rupture probable 07/09/2026
- analyse_zero_achat.pdf (22 pages, 338 Ko): section 5.2 refraichie au 19/08
  - 599 clients S1 actifs (41%), 125 nouveaux, 320 churned
  - Projection CONCENTRES août: 1 728 t (113% ajuste, +194 t d'avance)
  - Projection TOURTEAUX août: 4 360 t (113% ajuste) — dynamique retrouvee
  - Bundle ratio 2,5:1, cross-sell 95%, 11/14 agences au-dessus
  - Stock soja critique: rupture probable 07/09/2026
- analyse_bundle_soja_concentres.pdf (17 pages, 90 Ko): section 8 refraichie au 19/08
  - Tableau de bord global: ratio 2,5:1, 95% bundle, 1 271/1 331 cmds soja en bundle
  - Distribution des ratios: 69% <= 3:1, 29% 3-5:1, 2% >5:1
  - Conclusion mi-aout: excellente dynamique, CONCENTRES depassent l'objectif de 13% (+194 t)
- Outputs JSON: aout_mtd_20.json, soja_aout_20.json
- Conclusion generale: Excellente dynamique mi-aout avancée. Les CONCENTRES depassent l'objectif de 13% (113% ajuste, +194 t d'avance) avec 11/14 agences au-dessus. FAMLA et NDOBO ont rattrape leur retard (FAMLA 80% -> 96%, NDOBO 68% -> 97%). L'effet prix sur les TOURTEAUX se tasse (moy/j remonte de 147 a 168 t/j). Le stock soja reste critique (rupture probable 07/09) — reapprovisionnement urgent a programmer avant le 27/08.
