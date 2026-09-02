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

---
Task ID: aout-update-21
Agent: main
Task: Mettre a jour analyse zero achat, bundle soja-concentres et soja_aout avec la nouvelle extraction au 20/08/2026.

Work Log:
- Identifie la nouvelle extraction: `NJS GROUP ERP - Lignes de commandes + multicompany (20).xlsx` (modifiee 21/08 10:14, 7 568 lignes Livree).
- Couverture: 01/08 au 21/08. Le 21/08 a 5 lignes Livree (extraction matinale) — exclue. Donnees utilisees: 01/08 au 20/08 (17 jours ouvres lun-sam, 65% du mois).
- Metriques August MTD (au 20/08, 17j):
  - TOURTEAUX: 2 819 t (moy 165,8 t/j, proj 4 312 t, obj 3 850 -> 112% ajuste, +462 t d'avance)
  - CONCENTRES: 1 123 t (moy 66,0 t/j, proj 1 717 t, obj 1 534 -> 112% ajuste, +183 t d'avance)
  - Bundle ratio: 2,5:1 (stable vs 19/08)
  - Cross-sell: 96% bundle (1 355/1 417 cmds soja), 62 soja-only (4%)
  - Dist ratios: 70% <= 3:1 (+1 pt), 28% 3-5:1, 2% >5:1
  - Stock BEKOKO net: 71 230 sacs (3 561 t), rupture probable 09/09/2026 (20 jours, vs 19j au 19/08 — conso se stabilise)
  - Churned: 307 clients S1 sans achat aout a 17j (vs 320 au 19/08, 993 au 04/08)
  - Nouveaux clients: 130 (vs 125 au 19/08)
  - Cross-sell: 503 clients soja S1 avec conc en aout (63%)
- Performance CONCENTRES par region (au 20/08):
  - Ouest: 722 / 605 = 119% (+116 t)
  - Centre: 529 / 491 = 108% (+38 t)
  - Littoral: 466 / 438 = 106% (+28 t)
  - TOTAL: 1 717 / 1 535 = 112% (+183 t) — 10/14 agences au-dessus (vs 11/14 au 19/08 — leger tassement)
  - DJELENG 180% (+89 t), PK11 150% (+24 t), MBOUDA 136% (+35 t) — top performers
  - FAMLA 98% (-8 t), MESSASSI 98% (-5 t), NDOBO 94% (-13 t), NKOABANG 97% (-2 t) — sous l'objectif

Stage Summary:
- analyse_soja_aout.xlsx (5 feuilles, 16 Ko): au 20/08, moy/j soja 165,8 t (stable vs 19/08), rupture probable 09/09/2026
- analyse_zero_achat.pdf (22 pages, 338 Ko): section 5.2 refraichie au 20/08
  - 612 clients S1 actifs (42%), 130 nouveaux, 307 churned
  - Projection CONCENTRES août: 1 717 t (112% ajuste, +183 t d'avance)
  - Projection TOURTEAUX août: 4 312 t (112% ajuste)
  - Bundle ratio 2,5:1, cross-sell 96%, 10/14 agences au-dessus
  - Stock soja critique: rupture probable 09/09/2026
- analyse_bundle_soja_concentres.pdf (17 pages, 90 Ko): section 8 refraichie au 20/08
  - Tableau de bord global: ratio 2,5:1, 96% bundle, 1 355/1 417 cmds soja en bundle
  - Distribution des ratios: 70% <= 3:1, 28% 3-5:1, 2% >5:1
  - Conclusion mi-aout: excellente dynamique maintenue, CONCENTRES depassent l'objectif de 12% (+183 t)
- Outputs JSON: aout_mtd_21.json, soja_aout_21.json
- Conclusion generale: Excellente dynamique maintenue au 20/08. Les CONCENTRES depassent l'objectif de 12% (112% ajuste, +183 t d'avance) avec 10/14 agences au-dessus. Leger tassement vs 19/08 (113%, 11/14) — FAMLA, MESSASSI, NDOBO et NKOABANG sont passes sous l'objectif (correction normale). Le stock soja reste critique (rupture probable 09/09) — reapprovisionnement urgent a programmer avant le 29/08.

---
Task ID: aout-update-22
Agent: main
Task: Mettre a jour analyse zero achat, bundle soja-concentres et soja_aout avec la nouvelle extraction au 21/08/2026.

Work Log:
- Identifie la nouvelle extraction: `NJS GROUP ERP - Lignes de commandes + multicompany (3) (15) (1).xlsx` (modifiee 22/08 09:48, 7 891 lignes Livree).
- Couverture: 01/08 au 22/08. Le 22/08 a 2 lignes Livree (extraction matinale) — exclue. Donnees utilisees: 01/08 au 21/08 (18 jours ouvres lun-sam, 69% du mois).
- Metriques August MTD (au 21/08, 18j):
  - TOURTEAUX: 2 916 t (moy 162,0 t/j, proj 4 213 t, obj 3 850 -> 109% ajuste, +363 t d'avance)
  - CONCENTRES: 1 172 t (moy 65,1 t/j, proj 1 693 t, obj 1 534 -> 110% ajuste, +159 t d'avance)
  - Bundle ratio: 2,5:1 (stable vs 20/08)
  - Cross-sell: 96% bundle (1 398/1 462 cmds soja), 64 soja-only (4%)
  - Dist ratios: 70% <= 3:1 (stable), 28% 3-5:1, 2% >5:1
  - Stock BEKOKO net: 71 230 sacs (3 561 t), rupture probable 10/09/2026 (20 jours, stable vs 20/08)
  - Churned: 299 clients S1 sans achat aout a 18j (vs 307 au 20/08, 993 au 04/08)
  - Nouveaux clients: 134 (vs 130 au 20/08)
  - Cross-sell: 509 clients soja S1 avec conc en aout (63%, stable)
- Performance CONCENTRES par region (au 21/08):
  - Ouest: 690 / 605 = 114% (+84 t) — leger tassement (vs 119% au 20/08)
  - Centre: 534 / 491 = 109% (+42 t) — stable
  - Littoral: 470 / 438 = 107% (+32 t) — leger progres
  - TOTAL: 1 693 / 1 535 = 110% (+159 t) — 10/14 agences au-dessus (stable vs 20/08)
  - DJELENG 170% (+78 t), PK11 141% (+20 t), VILLAGE 136% (+26 t) — top performers
  - FAMLA 94% (-23 t, vs 98% au 20/08), NDOBO 91% (-17 t, vs 94% au 20/08) — tassement
  - AHALA 99% (-1 t) — sous l'objectif pour la 1ere fois

Stage Summary:
- analyse_soja_aout.xlsx (5 feuilles, 16 Ko): au 21/08, moy/j soja 162 t (leger tassement vs 166 au 20/08), rupture probable 10/09/2026
- analyse_zero_achat.pdf (21 pages, 337 Ko): section 5.2 refraichie au 21/08
  - 620 clients S1 actifs (43%), 134 nouveaux, 299 churned
  - Projection CONCENTRES août: 1 693 t (110% ajuste, +159 t d'avance)
  - Projection TOURTEAUX août: 4 213 t (109% ajuste)
  - Bundle ratio 2,5:1, cross-sell 96%, 10/14 agences au-dessus
  - Stock soja critique: rupture probable 10/09/2026
- analyse_bundle_soja_concentres.pdf (17 pages, 90 Ko): section 8 refraichie au 21/08
  - Tableau de bord global: ratio 2,5:1, 96% bundle, 1 398/1 462 cmds soja en bundle
  - Distribution des ratios: 70% <= 3:1, 28% 3-5:1, 2% >5:1
  - Conclusion mi-aout: dynamique positive maintenue, CONCENTRES depassent l'objectif de 10% (+159 t)
- Outputs JSON: aout_mtd_22.json, soja_aout_22.json
- Conclusion generale: Dynamique positive maintenue au 21/08. Les CONCENTRES depassent l'objectif de 10% (110% ajuste, +159 t d'avance) avec 10/14 agences au-dessus. Leger tassement vs 20/08 (112%, 10/14) — FAMLA et NDOBO ont legerement recule. Le stock soja reste critique (rupture probable 10/09) — reapprovisionnement urgent a programmer avant le 30/08.

---
Task ID: aout-update-23
Agent: main
Task: Mettre a jour analyse zero achat, bundle soja-concentres et soja_aout avec la nouvelle extraction au 22/08/2026.

Work Log:
- Identifie la nouvelle extraction: `NJS GROUP ERP - Lignes de commandes + multicompany (3) (16) (1).xlsx` (modifiee 22/08 20:51, 8 136 lignes Livree).
- Couverture: 01/08 au 22/08. Le 22/08 a 218 lignes Livree (journee complete). Donnees utilisees: 01/08 au 22/08 (19 jours ouvres lun-sam, 73% du mois).
- Metriques August MTD (au 22/08, 19j):
  - TOURTEAUX: 2 960 t (moy 155,8 t/j, proj 4 051 t, obj 3 850 -> 105% ajuste, +201 t d'avance)
  - CONCENTRES: 1 192 t (moy 62,8 t/j, proj 1 632 t, obj 1 534 -> 106% ajuste, +97 t d'avance)
  - Bundle ratio: 2,5:1 (stable vs 21/08)
  - Cross-sell: 96% bundle (1 427/1 493 cmds soja), 66 soja-only (4%)
  - Dist ratios: 70% <= 3:1 (stable), 28% 3-5:1, 2% >5:1
  - Stock BEKOKO net: 71 230 sacs (3 561 t), rupture probable 12/09/2026 (21 jours, +1j vs 21/08 — conso a ralenti)
  - Churned: 294 clients S1 sans achat aout a 19j (vs 299 au 21/08, 993 au 04/08)
  - Nouveaux clients: 137 (vs 134 au 21/08)
  - Cross-sell: 514 clients soja S1 avec conc en aout (64%)
- Performance CONCENTRES par region (au 22/08):
  - Ouest: 655 / 605 = 108% (+50 t) — tassement (vs 114% au 21/08)
  - Centre: 521 / 491 = 106% (+29 t) — tassement (vs 109% au 21/08)
  - Littoral: 456 / 438 = 104% (+18 t) — tassement (vs 107% au 21/08)
  - TOTAL: 1 632 / 1 535 = 106% (+97 t) — 8/14 agences au-dessus (vs 10/14 au 21/08 — décroissance)
  - DJELENG 161% (+68 t), PK11 145% (+21 t), VILLAGE 129% (+20 t) — top performers
  - FAMLA 89% (-43 t, vs 94% au 21/08), NDOBO 88% (-23 t, vs 91% au 21/08) — tassement marque
  - MESSASSI 100% (-1 t), AHALA 96% (-4 t), NKONGSAMBA 96% (-3 t), NKOABANG 96% (-2 t) — sous l'objectif

Stage Summary:
- analyse_soja_aout.xlsx (5 feuilles, 16 Ko): au 22/08, moy/j soja 156 t (leger tassement vs 162 au 21/08), rupture probable 12/09/2026
- analyse_zero_achat.pdf (22 pages, 338 Ko): section 5.2 refraichie au 22/08
  - 625 clients S1 actifs (43%), 137 nouveaux, 294 churned
  - Projection CONCENTRES août: 1 632 t (106% ajuste, +97 t d'avance)
  - Projection TOURTEAUX août: 4 051 t (105% ajuste)
  - Bundle ratio 2,5:1, cross-sell 96%, 8/14 agences au-dessus
  - Stock soja: rupture probable 12/09/2026
- analyse_bundle_soja_concentres.pdf (17 pages, 90 Ko): section 8 refraichie au 22/08
  - Tableau de bord global: ratio 2,5:1, 96% bundle, 1 427/1 493 cmds soja en bundle
  - Distribution des ratios: 70% <= 3:1, 28% 3-5:1, 2% >5:1
  - Conclusion mi-aout: dynamique positive mais qui ralentit, CONCENTRES depassent l'objectif de 6% (+97 t)
- Outputs JSON: aout_mtd_23.json, soja_aout_23.json
- Conclusion generale: Dynamique positive mais qui ralentit au 22/08 (73% du mois). Les CONCENTRES depassent l'objectif de 6% (106% ajuste, +97 t d'avance) avec 8/14 agences au-dessus (vs 10/14 au 21/08 — décroissance). FAMLA et NDOBO ont encore tasse (-5 et -3 pts). Le stock soja reste critique (rupture probable 12/09) — reapprovisionnement a programmer avant le 01/09.

---
Task ID: aout-update-26
Agent: main
Task: Mettre a jour analyse zero achat, bundle soja-concentres et soja_aout avec la nouvelle extraction au 25/08/2026.

Work Log:
- Identifie la nouvelle extraction: `NJS GROUP ERP - Lignes de commandes + multicompany (3) (18) (1).xlsx` (modifiee 26/08 06:34, 9 293 lignes Livree).
- Couverture: 01/08 au 25/08 (weekend 23-24/08 inclus via journee 24/08 a 675 rows). Le 25/08 a 390 lignes Livree (journee complete). Donnees utilisees: 01/08 au 25/08 (21 jours ouvres lun-sam, 81% du mois).
- Metriques August MTD (au 25/08, 21j):
  - TOURTEAUX: 3 336 t (moy 158,9 t/j, proj 4 131 t, obj 3 850 -> 107% ajuste, +281 t d'avance)
  - CONCENTRES: 1 368 t (moy 65,2 t/j, proj 1 694 t, obj 1 534 -> 110% ajuste, +160 t d'avance)
  - Bundle ratio: 2,4:1 (vs 2,5:1 au 22/08 — NOUVELLE amelioration)
  - Cross-sell: 96% bundle (1 636/1 709 cmds soja), 73 soja-only (4%)
  - Dist ratios: 72% <= 3:1 (+2 pts), 26% 3-5:1 (-2 pts), 2% >5:1
  - Stock BEKOKO net: 71 230 sacs (3 561 t), rupture probable 14/09/2026 (20 jours, stable)
  - Churned: 271 clients S1 sans achat aout a 21j (vs 294 au 22/08, 993 au 04/08)
  - Nouveaux clients: 163 (vs 137 au 22/08, +26 — forte acceleration)
  - Cross-sell: 531 clients soja S1 avec conc en aout (66%)
- Performance CONCENTRES par region (au 25/08):
  - Ouest: 714 / 605 = 118% (+108 t) — fort rebond (vs 108% au 22/08)
  - Centre: 524 / 491 = 107% (+33 t) — stable
  - Littoral: 457 / 438 = 104% (+19 t) — stable
  - TOTAL: 1 694 / 1 535 = 110% (+160 t) — 7/14 agences au-dessus (stable vs 22/08)
  - DJELENG 180% (+90 t), PK11 143% (+20 t), NKOLBISSON 136% (+13 t) — top performers
  - FAMLA 96% (-17 t, vs 89% au 22/08 — rattrapage), NDOBO 91% (-17 t, stable vs 88% au 22/08)
  - MESSASSI 97% (-7 t), AHALA 97% (-3 t), NKONGSAMBA 98% (-1 t), NKOABANG 99% (-1 t) — sous l'objectif

Stage Summary:
- analyse_soja_aout.xlsx (5 feuilles, 16 Ko): au 25/08, moy/j soja 159 t (rebond vs 156 au 22/08), rupture probable 14/09/2026
- analyse_zero_achat.pdf (22 pages, 338 Ko): section 5.2 refraichie au 25/08
  - 648 clients S1 actifs (45%), 163 nouveaux, 271 churned
  - Projection CONCENTRES août: 1 694 t (110% ajuste, +160 t d'avance)
  - Projection TOURTEAUX août: 4 131 t (107% ajuste)
  - Bundle ratio 2,4:1 (nouvelle amelioration), cross-sell 96%, 7/14 agences au-dessus
  - Stock soja: rupture probable 14/09/2026
- analyse_bundle_soja_concentres.pdf (17 pages, 90 Ko): section 8 refraichie au 25/08
  - Tableau de bord global: ratio 2,4:1, 96% bundle, 1 636/1 709 cmds soja en bundle
  - Distribution des ratios: 72% <= 3:1, 26% 3-5:1, 2% >5:1
  - Conclusion mi-aout: rebond de la dynamique, CONCENTRES depassent l'objectif de 10% (+160 t)
- Outputs JSON: aout_mtd_26.json, soja_aout_26.json
- Conclusion generale: Rebond de la dynamique au 25/08 apres le creux du 22/08. Les CONCENTRES depassent l'objectif de 10% (110% ajuste, +160 t d'avance) avec 7/14 agences au-dessus. Le bundle ratio s'amelioire encore a 2,4:1 (vs 2,5:1). FAMLA a rattrape (89% -> 96%). Le stock soja reste critique (rupture probable 14/09) — reapprovisionnement a programmer avant le 01/09.

---
Task ID: aout-update-27
Agent: main
Task: Mettre a jour analyse zero achat, bundle soja-concentres et soja_aout avec la nouvelle extraction au 26/08/2026.

Work Log:
- Identifie la nouvelle extraction: `NJS GROUP ERP - Lignes de commandes + multicompany (21).xlsx` (modifiee 26/08 20:32, 9 688 lignes Livree).
- Couverture: 01/08 au 26/08. Le 26/08 a 288 lignes Livree (journee complete). Donnees utilisees: 01/08 au 26/08 (22 jours ouvres lun-sam, 85% du mois).
- Metriques August MTD (au 26/08, 22j):
  - TOURTEAUX: 3 437 t (moy 156,2 t/j, proj 4 062 t, obj 3 850 -> 106% ajuste, +212 t d'avance)
  - CONCENTRES: 1 421 t (moy 64,6 t/j, proj 1 679 t, obj 1 534 -> 109% ajuste, +144 t d'avance)
  - Bundle ratio: 2,4:1 (stable vs 25/08)
  - Cross-sell: 96% bundle (1 702/1 778 cmds soja), 76 soja-only (4%)
  - Dist ratios: 72% <= 3:1 (stable), 26% 3-5:1, 2% >5:1
  - Stock BEKOKO net: 71 230 sacs (3 561 t), rupture probable 16/09/2026 (21 jours, +1j vs 25/08 — conso a legerement ralenti)
  - Churned: 264 clients S1 sans achat aout a 22j (vs 271 au 25/08, 993 au 04/08)
  - Nouveaux clients: 171 (vs 163 au 25/08)
  - Cross-sell: 537 clients soja S1 avec conc en aout (67%)
- Performance CONCENTRES par region (au 26/08):
  - Ouest: 709 / 605 = 117% (+104 t) — stable vs 25/08 (118%)
  - Centre: 514 / 491 = 105% (+23 t) — leger tassement (vs 107% au 25/08)
  - Littoral: 456 / 438 = 104% (+18 t) — stable vs 25/08 (104%)
  - TOTAL: 1 679 / 1 535 = 109% (+144 t) — 7/14 agences au-dessus (stable vs 25/08)
  - DJELENG 179% (+89 t), PK11 147% (+22 t), MBOUDA 137% (+37 t) — top performers
  - FAMLA 94% (-22 t, vs 96% au 25/08), NDOBO 91% (-18 t, stable vs 25/08)
  - MESSASSI 92% (-15 t, vs 97% au 25/08), NKONGSAMBA 93% (-4 t), BUEA 91% (-5 t), AHALA 99% (-1 t), NKOABANG 98% (-1 t)

Stage Summary:
- analyse_soja_aout.xlsx (5 feuilles, 16 Ko): au 26/08, moy/j soja 156 t (leger tassement vs 159 au 25/08), rupture probable 16/09/2026
- analyse_zero_achat.pdf (22 pages, 338 Ko): section 5.2 refraichie au 26/08
  - 655 clients S1 actifs (45%), 171 nouveaux, 264 churned
  - Projection CONCENTRES août: 1 679 t (109% ajuste, +144 t d'avance)
  - Projection TOURTEAUX août: 4 062 t (106% ajuste)
  - Bundle ratio 2,4:1, cross-sell 96%, 7/14 agences au-dessus
  - Stock soja: rupture probable 16/09/2026
- analyse_bundle_soja_concentres.pdf (17 pages, 90 Ko): section 8 refraichie au 26/08
  - Tableau de bord global: ratio 2,4:1, 96% bundle, 1 702/1 778 cmds soja en bundle
  - Distribution des ratios: 72% <= 3:1, 26% 3-5:1, 2% >5:1
  - Conclusion mi-aout: dynamique maintenue au-dessus de l'objectif, CONCENTRES depassent l'objectif de 9% (+144 t)
- Outputs JSON: aout_mtd_27.json, soja_aout_27.json
- Conclusion generale: Dynamique maintenue au 26/08. Les CONCENTRES depassent l'objectif de 9% (109% ajuste, +144 t d'avance) avec 7/14 agences au-dessus. Le bundle ratio se maintient a 2,4:1. Le stock soja reste critique (rupture probable 16/09) — reapprovisionnement a programmer avant le 03/09.

---
Task ID: forecast-q4-2026-pace
Agent: main
Task: Forecast Q4 2026 (sept-dec) en volumes et valeurs avec methode PACE. Production de tous les livrables (guide methodologique, document strategique PACE, matrice RACI, resume executif, proposition de projet, Excel forecast, notebook Jupyter, graphiques).

Work Log:
- Phase PREPARE: Construction dataset consolide (95 819 records Jan 2025 - 26 aout 2026)
  - Sources: 86d96135 (2025), ventes janv a juin 2026, NJS ERP (9) juillet, NJS ERP (21) aout
  - Filtrage: etat Livree uniquement, exclusion clients internes (SPC/PDC/COMPTOIR/EMANA)
  - Mapping agences unifie (formats 2025 court + 2026 long)
  - Calcul prix moyens Option A (extrapolation CA/quantite, precision ~90%)
  - 26 refs produits, 14 agences, 3 regions
- Phase ANALYZE (AED):
  - Stats descriptives: 109 045 t total, 51 372 M FCFA CA
  - Saisonnalite mensuelle (base 2025): Q4 = 35-46% du volume annuel
  - Top produits: T102 (soja) = 76% du volume, C104 = 13%, C101 = 5%
  - Top agences: FAMLA 31.6%, NDOBO 13.7%, MESSASSI 9.9%
  - 6 graphiques AED generes
- Phase CONSTRUCT:
  - Strategie: 13 modeles Prophet famille × region, desagreges en produit × agence
  - Configuration Prophet: yearly_seasonality=True, multiplicative, changepoint=0.05, interval=0.8
  - 4 scenarios soja: S1 rupture (7 291 t, 4 656 M), S2 reappro 50% (32 207 t, 14 922 M), S3 reappro 100% (37 988 t, 17 304 M), S4 baisse prix (41 067 t, 17 168 M)
  - 4 196 forecasts detailles generes (produit × agence × mois × scenario)
- Phase EXECUTE:
  - Excel forecast multi-feuilles (9 feuilles, 231 KB)
  - Notebook Jupyter commente (5 sections PACE)
  - 13 graphiques de visualisation
  - 5 PDFs: resume executif (2p), proposition projet (7p), matrice RACI (3p), document strategique PACE (15p), guide methodologique (30+p)

Stage Summary:
- 9 livrables produits totalisant ~1.4 MB
- Forecast Q4 2026 scenario reference (S3): 37 988 t, 17 304 M FCFA
- Sensibilite critique au reappro soja: ecart S1-S3 = 12 648 M FCFA (73% du CA)
- Bundle ratio 2.4:1 atteint en aout (vs 3.5:1 juillet) - a maintenir Q4
- Top 5 agences CA Q4: FAMLA 3 420 M, NDOBO 1 870 M, DJELENG 1 230 M, MESSASSI 1 145 M, MBOUDA 870 M
- Recommandation principale: reappro soja urgent avant 15/09/2026 (80 000 sacs min)

---
Task ID: aout-update-29
Agent: main
Task: Mettre a jour analyse zero achat, bundle soja-concentres et soja_aout avec la nouvelle extraction au 28/08/2026. User mentionne aussi que les livrables du projet forecast doivent pouvoir etre mis a jour au fur et a mesure.

Work Log:
- Identifie la nouvelle extraction: `NJS GROUP ERP - Lignes de commandes + multicompany (24).xlsx` (modifiee 29/08, 10 515 lignes Livree).
- Couverture: 01/08 au 28/08. Le 28/08 a 420 lignes Livree (journee complete). Donnees utilisees: 01/08 au 28/08 (24 jours ouvres lun-sam, 92% du mois).
- Metriques August MTD (au 28/08, 24j):
  - TOURTEAUX: 3 625 t (moy 151,0 t/j, proj 3 927 t, obj 3 850 -> 102% ajuste, +77 t d'avance)
  - CONCENTRES: 1 532 t (moy 63,8 t/j, proj 1 660 t, obj 1 534 -> 108% ajuste, +125 t d'avance)
  - Bundle ratio: 2,4:1 (stable vs 26/08)
  - Cross-sell: 95% bundle (1 847/1 939 cmds soja), 92 soja-only (5%)
  - Dist ratios: 74% <= 3:1 (+4 pts vs 26/08), 25% 3-5:1, 1% >5:1
  - Stock BEKOKO net: 71 230 sacs (3 561 t), rupture probable 18/09/2026 (21 jours, stable)
  - Churned: 248 clients S1 sans achat aout a 24j (vs 264 au 26/08, 993 au 04/08)
  - Nouveaux clients: 188 (vs 171 au 26/08)
  - Cross-sell: 555 clients soja S1 avec conc en aout (69%)
- Performance CONCENTRES par region (au 28/08):
  - Ouest: 688 / 605 = 114% (+83 t)
  - Centre: 510 / 491 = 104% (+18 t)
  - Littoral: 462 / 438 = 106% (+24 t)
  - TOTAL: 1 660 / 1 535 = 108% (+125 t) — 8/14 agences au-dessus (amelioration vs 26/08: 7/14)
  - DJELENG 171% (+80 t), PK11 145% (+22 t), MBOUDA 132% (+31 t) — top performers
  - FAMLA 93% (-28 t), NDOBO 89% (-22 t), MESSASSI 95% (-10 t) — sous l'objectif

Stage Summary:
- analyse_soja_aout.xlsx (5 feuilles, 16 Ko): au 28/08, moy/j soja 151 t, rupture probable 18/09/2026
- analyse_zero_achat.pdf (22 pages, 338 Ko): section 5.2 refraichie au 28/08
- analyse_bundle_soja_concentres.pdf (17 pages, 90 Ko): section 8 refraichie au 28/08
- Note user: les livrables du projet forecast doivent pouvoir etre mis a jour au fur et a mesure des mises a jour des extractions.
- Outputs JSON: aout_mtd_29.json, soja_aout_29.json
- Conclusion generale: Au 28/08 (92% du mois), les CONCENTRES depassent l'objectif de 8% (108% ajuste, +125 t d'avance) avec 8/14 agences au-dessus. La distribution des ratios bundle s'est amelioree (74% <= 3:1 vs 72% au 26/08). Le stock soja reste critique (rupture probable 18/09, 21 jours de stock). Le churn continue de se resorber (248 vs 264 au 26/08). Avec 2 jours ouvrables restants (29-30/08), le mois d'aout s'annonce comme le 2e mois consecutif au-dessus de l'objectif CONCENTRES.

---
Task ID: aout-update-30
Agent: main
Task: Mettre a jour analyse zero achat, bundle soja-concentres, soja_aout et forecast pipeline avec extraction au 29/08/2026.

Work Log:
- Extraction: NJS GROUP ERP (25).xlsx, 10 821 lignes Livree, couverture 01-29/08.
- 25 jours ouvres (lun-sam), 96% du mois. Quasi définitif (plus que 1 jour ouvrable: lundi 31/08).
- Metriques August MTD (au 29/08, 25j):
  - TOURTEAUX: 3 672 t (moy 146.9 t/j, proj 3 818 t, obj 3 850 -> 99% ajuste, -32 t)
  - CONCENTRES: 1 560 t (moy 62.4 t/j, proj 1 623 t, obj 1 534 -> 106% ajuste, +88 t d'avance)
  - Bundle ratio: 2,4:1 (stable)
  - Cross-sell: 95% (1 891/1 987), 96 soja-only (5%)
  - Dist ratios: 74% <= 3:1, 24% 3-5:1, 1% >5:1
  - Stock: rupture probable 20/09/2026 (22 jours)
  - Churned: 247, Nouveaux: 194, Cross-sell soja→conc: 557 (69%)
- Performance CONCENTRES par region:
  - Ouest: 672 / 605 = 111% (+66 t)
  - Centre: 494 / 491 = 101% (+2 t)
  - Littoral: 457 / 438 = 104% (+19 t)
  - TOTAL: 1 623 / 1 535 = 106% (+88 t) — 8/14 agences au-dessus
- Forecast pipeline execute (2.4 min): dataset, forecast S3, Excel, graphiques tous mis a jour.
- Tous les livrables sont a jour avec l'extraction du 29/08.

Stage Summary:
- analyse_soja_aout.xlsx (5 feuilles): au 29/08, moy/j soja 147 t, rupture probable 20/09/2026
- analyse_zero_achat.pdf (22 pages): section 5.2 refraichie au 29/08 (25j, 96%)
- analyse_bundle_soja_concentres.pdf (17 pages): section 8 refraichie au 29/08
- forecast_q4_2026_S3_volumes_valeurs.xlsx (8 feuilles): mis a jour via pipeline automatise
- Conclusion: Août 2026 = 2e mois consecutif au-dessus de l'objectif CONCENTRES (106% ajuste, +88 t). Le TOURTEAUX termine juste sous l'objectif (99%, -32 t). Bundle ratio record a 2.4:1. Stock soja critique (rupture 20/09).

---
Task ID: forecast-2027-pdfs
Agent: main
Task: Régénérer tous les PDFs Q4 2026 avec chiffres définitifs août + générer les 5 PDFs PACE pour le forecast 2027.

Work Log:
- Q4 2026 PDFs régénérés avec chiffres définitifs (extraction 27, 01-31/08):
  - 01_resume_executif.pdf, 02_proposition_projet.pdf, 03_matrice_raci.pdf, 04_document_strategique_pace.pdf, 05_guide_methodologique.pdf
  - Tous mis à jour: Volume Q4 = 40 921 t, CA Q4 = 18 798 M FCFA, 115 592 records, 7 224 lignes
- Forecast 2027 PDFs générés (nouveaux):
  - 01_resume_executif.pdf (50 KB) - Vue d'ensemble 2027
  - 02_proposition_projet.pdf (50 KB) - Cadre formel 2027
  - 03_matrice_raci.pdf (49 KB) - Rôles et responsabilités 2027
  - 04_document_strategique_pace.pdf (51 KB) - Stratégie PACE 2027
  - 05_guide_methodologique.pdf (51 KB) - Méthodes Prophet + désaisonnalisation
- Tous les PDFs utilisent les mêmes chiffres officiels:
  - 2027: Volume 109 288 t, CA 58 550 M FCFA, 69 produits, 14 agences, 7 224 lignes
  - Q4 2026: Volume 40 921 t, CA 18 798 M FCFA (mis à jour avec extraction 31/08)
- Innovations 2027 documentées dans tous les PDFs:
  - Désaisonnalisation effet soja (cap moyenne S1)
  - En cours + Validées incluses
  - Prix soja actualisé 25 000 FCFA
  - Forecast 12 mois complet

---
Task ID: forecast-2027-v2-complement-alimentaire
Agent: main
Task: Mettre à jour le forecast 2027 selon les instructions utilisateur: (1) utiliser les années 2024-2026 uniquement (laisser 2021-2023), (2) utiliser les tendances BELGOKILL comme proxy pour toute la famille COMPLEMENT ALIMENTAIRE, (3) ajouter la famille COMPLEMENT_ALIMENTAIRE au forecast 2027 (1L=1kg).

Work Log:
- Phase 1: Construction du nouveau dataset 2024-2026
  - Sources: LY_24 (Jul-Dec 2024, 52 215 records), 2025 (66 206 records), S1 2026 (37 745 records), Juillet 2026 (7 537 records), Août 2026 (6 857 records)
  - Exclusion des années 2021-2023 (focus 2024-2026, 32 mois d'historique)
  - Ajout de la famille COMPLEMENT_ALIMENTAIRE: 10 produits liquides (V300/V305 BELGOKILL, CA001-CA008 BELGO xxx)
  - Conversion 1L = 1kg (qte = kg, tonnes = qte/1000), V305 = 200 kg/qte
  - Exclusion MAIS (M1051/M1052) et DIVERS
  - Total: 170 560 records, 78 produits, 14 agences
  - Script: scripts/build_dataset_2024_2026.py
  - Output: scripts/dataset_2024_2026.csv

- Phase 2: Mise à jour du forecast 2027
  - 7 familles au lieu de 6 (ajout COMPLEMENT_ALIMENTAIRE)
  - Prophet pour 5 familles (TOURTEAUX, CONCENTRES, INGREDIENTS, ALIMENT_COMPLET, COMPLEMENT_ALIMENTAIRE)
  - Extrapolation pour 2 familles (MATERIEL_ELEVAGE, PREMIX)
  - Désaisonnalisation soja Jul-Août 2026 maintenue (cap moyenne S1 2026)
  - En cours + Validées août 2026 inclus
  - Prix soja 25 000 FCFA/sac maintenu
  - Prix COMPLEMENT_ALIMENTAIRE par litre: 2 500-15 000 FCFA/L selon produit
  - BELGOKILL (V300) utilisé comme proxy tendance (38% du CA famille)
  - Script: scripts/forecast_2027_v3.py
  - Output: scripts/forecast_2027_S3.csv (8 868 lignes)

- Phase 3: Régénération Excel
  - 8 feuilles (vs 7 avant): synthèse, par famille × mois, par région × mois, par agence, par produit, détail complet, hypothèses, saisonnalité
  - 7 familles dans synthèse (COMPLEMENT_ALIMENTAIRE surligné en jaune)
  - Sheet 8 nouvelle: coefficients saisonniers 2025 vs 2027
  - Script: scripts/forecast_2027_excel_v2.py
  - Output: download/forecast_2027_S3_volumes_valeurs.xlsx (390 KB)

- Phase 4: Régénération des 5 PDFs PACE
  - Tous les chiffres actualisés: 118 608 t, 64 114 M FCFA (vs 75 112 t, 41 219 M FCFA avant)
  - NOUVEAU: Section dédiée COMPLEMENT_ALIMENTAIRE dans guide méthodologique (section 7)
  - NOUVEAU: Tableau CA par produit BELGOxxx avec parts famille
  - NOUVEAU: Lignes surlignées en jaune pour COMPLEMENT_ALIMENTAIRE dans tableaux
  - NOUVEAU: Historique 2024-2026 vs 2027 (au lieu de YTD 2026 + Q4 fcst)
  - Script: scripts/forecast_2027_pdfs_v2.py
  - Outputs: 5 PDFs dans download/forecast_2027/ (270 KB total)

Stage Summary:
- dataset_2024_2026.csv: 170 560 records (vs 115 086 avant), 78 produits (vs 69), 32 mois (vs 20)
- forecast_2027_S3.csv: 8 868 lignes (vs 12 600 — moins car COMPLEMENT_ALIM n'a pas toutes les agences)
- forecast_2027_S3_volumes_valeurs.xlsx: 390 KB, 8 feuilles, 7 familles
- 5 PDFs forecast 2027 régénérés:
  - 01_resume_executif.pdf (58 KB)
  - 02_proposition_projet.pdf (51 KB)
  - 03_matrice_raci.pdf (49 KB)
  - 04_document_strategique_pace.pdf (57 KB)
  - 05_guide_methodologique.pdf (55 KB)
- Forecast 2027 v2: 118 608 t, 64 114 M FCFA (+58% volume vs 2026 annualisé)
- 7 familles: TOURTEAUX (92 517 t, 46 261 M), CONCENTRES (24 680 t, 16 504 M), ALIMENT_COMPLET (890 t, 688 M), MATERIEL_ELEVAGE (0 t, 377 M), INGREDIENTS (518 t, 107 M), PREMIX (0 t, 160 M), COMPLEMENT_ALIMENTAIRE (3 t, 17 M)
- Top 5 agences 2027: FAMLA 18 085 M, NDOBO 9 476 M, MESSASSI 5 848 M, DJELENG 5 166 M, VILLAGE 3 749 M
- Conclusion: v2 intègre COMPLEMENT_ALIMENTAIRE (proxy BELGOKILL, 1L=1kg) et utilise 32 mois d'historique (2024-2026). La progression de +58% vs 2026 annualisé reflète la tendance haussière 2024-2026 amplifiée par l'effet prix soja 25 000 FCFA. COMPLEMENT_ALIMENTAIRE reste marginal en volume (3 t) mais contribue à 17 M FCFA de CA grâce à la forte valeur unitaire des liquides.

---
Task ID: forecast-v3-2023-integration
Agent: main
Task: Intégrer l'année 2023 au forecast, refaire Q4 2026 et 2027, MATERIEL_ELEVAGE tonnes=0, BELGOKILL V300 1L only (V305 200L exclu).

Work Log:
- Phase 1: Construction du dataset 2023-2026
  - Source LY_21_24 (filtre 2023): 6 016 records (Jan-Dec 2023)
  - Combiné avec dataset_2024_2026 (170 560 records) → 176 576 records au total
  - V305 (BELGOKILL 200L) EXCLU du dataset selon demande utilisateur
  - MATERIEL_ELEVAGE tonnes=0 appliqué (CA only, non exprimable en volume)
  - COMPLEMENT_ALIMENTAIRE: 9 produits liquides 1L (V300 + CA001-CA008), conversion 1L=1kg
  - Script: scripts/build_dataset_2023_2026.py
  - Output: scripts/dataset_2023_2026.csv (24 381 KB)
  - Vérification: V305 absent, MATERIEL_ELEVAGE tonnes=0, COMPLEMENT_ALIMENTAIRE 18.8 t (2023-2026)

- Phase 2: Forecast Q4 2026 (v2)
  - Script: scripts/pace_11_forecast_S3_v2.py
  - 7 familles au lieu de 6 (ajout COMPLEMENT_ALIMENTAIRE)
  - 5 familles Prophet (incluant COMPLEMENT_ALIMENTAIRE)
  - 2 familles extrapolation (MATERIEL_ELEVAGE tonnes=0, PREMIX tonnes=0)
  - Désaisonnalisation soja maintenue (cap moyenne S1 2026)
  - V305 EXCLU — seul V300 1L conservé
  - Prix Q4: 25 000 FCFA/sac soja, prix par litre pour COMPLEMENT (2 500-15 000 FCFA/L)
  - Output: scripts/forecast_q4_2026_S3.csv (3 884 lignes)

- Phase 3: Forecast 2027 (v3)
  - Script: scripts/forecast_2027_v3.py (mis à jour)
  - Même méthodologie que Q4 2026 mais 12 mois (Jan-Déc 2027)
  - Output: scripts/forecast_2027_S3.csv (11 652 lignes)

- Phase 4: Régénération Excels
  - Q4 2026 Excel: scripts/pace_12_excel_S3_v2.py
    - 8 feuilles, 7 familles, COMPLEMENT_ALIMENTAIRE surligné en jaune
    - Output: download/forecast_q4_2026_S3_volumes_valeurs.xlsx (188 KB)
  - 2027 Excel: scripts/forecast_2027_excel_v2.py (mis à jour)
    - 8 feuilles, 7 familles, sheet "Saisonnalité" ajoutée
    - Output: download/forecast_2027_S3_volumes_valeurs.xlsx (516 KB)

- Phase 5: Régénération PDFs (10 au total)
  - Script unique: scripts/forecast_pdfs_v3.py (génère Q4 2026 + 2027)
  - Q4 2026: 5 PDFs dans download/forecast_q4_2026/ (51-56 KB chacun)
  - 2027: 5 PDFs dans download/forecast_2027/ (49-55 KB chacun)
  - Tous les PDFs mentionnent: données 2023-2026 (44 mois), V305 exclu, MATERIEL_ELEVAGE tonnes=0

Stage Summary:
- Dataset 2023-2026: 176 576 records (vs 170 560 v2), 115 produits (vs 78 v2), 44 mois (vs 32 v2)
- Q4 2026 v3 forecast: 25 410 t, 13 924 M FCFA (vs 40 921 t, 18 798 v1)
  - TOURTEAUX: 18 321 t (65,8% CA)
  - CONCENTRÉS: 6 610 t (31,6% CA)
  - COMPLEMENT_ALIMENTAIRE: 1 t, 7 M FCFA (V300 1L only)
  - MATERIEL_ELEVAGE: 0 t, 114 M FCFA (CA only)
  - Pic octobre: 7 628 t (29,5% CA Q4)
  - Top agences: FAMLA 4 028 M (29%), NDOBO 1 845 M, MESSASSI 1 335 M

- 2027 v3 forecast: 81 582 t, 44 402 M FCFA (vs 118 608 t, 64 114 v2)
  - TOURTEAUX: 60 673 t (68,3% CA)
  - CONCENTRÉS: 19 658 t (29,5% CA)
  - COMPLEMENT_ALIMENTAIRE: 3 t, 18 M FCFA (V300 1L only)
  - MATERIEL_ELEVAGE: 0 t, 250 M FCFA (CA only)
  - Q4 2027: 22 533 t (27,6% CA)
  - Top agences: FAMLA 13 761 M, NDOBO 5 516 M, MESSASSI 4 098 M
  - Progression 2027 vs 2026 annualisé: +4,4% (vs +51,8% en v2 — baisse car historique plus large)

- 10 PDFs générés (5 Q4 2026 + 5 2027), tous avec mention V305 EXCLU et données 2023-2026
- Conclusion: v3 intègre l'année 2023 (44 mois d'historique), exclut V305 (BELGOKILL 200L), maintient MATERIEL_ELEVAGE à 0 tonne. Les volumes forecast sont plus conservateurs qu'en v2 (Q4 2026: 25 410 t vs 40 921; 2027: 81 582 t vs 118 608) car l'historique 2023 (année post-COVID avec volumes modérés) apporte une saisonnalité plus équilibrée et l'absence de V305 réduit le CA famille COMPLEMENT_ALIMENTAIRE.
