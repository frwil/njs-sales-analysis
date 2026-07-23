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
