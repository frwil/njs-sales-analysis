"""
Met à jour build_pdf_report_v3.py avec tous les nouveaux chiffres après corrections:
- Total S1 2026: 41 148 t (au lieu de 42 826 t)
- INGREDIENTS: 2 382 t (au lieu de 4 060 t) — exclusion M1051 CA=0
- COMPLEMENT ALIM.: 3 t (inchangé, déjà correct)
- DIVERS: 2,1 t + 17 M FCFA (au lieu de 13,6 M — ajout PONT_BASCULE et CONTRIBUTION_CARBURANT Validée)
- Atteinte S1: 99,8% / 101,8% (au lieu de 103,9%)
- CA global: 17 344 M (au lieu de 17 340 M)
"""
import re

PDF_SCRIPT = "/home/z/my-project/scripts/build_pdf_report_v3.py"

with open(PDF_SCRIPT, "r", encoding="utf-8") as f:
    content = f.read()

# Tous les remplacements (old, new, description)
replacements = [
    # ---- COVER (line 207) ----
    ("42 826 t réalisées vs 41 230 t",
     "41 148 t réalisées vs 41 230 t",
     "Cover KPI volume"),

    # ---- SYNTHESE (line 414) ----
    ("L'analyse porte sur 42 826 tonnes et 17,34 Md FCFA de ventes (Janvier-Juin 2026), ",
     "L'analyse porte sur 41 148 tonnes et 17,34 Md FCFA de ventes (Janvier-Juin 2026), ",
     "Synthese volume"),

    # KPI cards
    ('kpi_card("Volume S1", "42 826 t", "(103,9% obj)")',
     'kpi_card("Volume S1", "41 148 t", "(99,8% obj)")',
     "KPI card volume"),

    ("Avec 42 826 tonnes en S1 2026, le volume atteint 103,9% de l'objectif. ",
     "Avec 41 148 tonnes en S1 2026, le volume atteint 99,8% de l'objectif (hors pipeline en cours). ",
     "Synthese narrative volume"),

    ("le volume total atteint est de <b>42 826 tonnes</b>, soit <b>103,9% de l'objectif S1</b> (41 230 tonnes). ",
     "le volume total atteint est de <b>41 148 tonnes</b>, soit <b>99,8% de l'objectif S1</b> (41 230 tonnes). "
     "En incluant le pipeline de commandes en cours et validées (192,8 t), le total projeté atteint 41 341 t (102,3% de l'objectif). ",
     "Section 3.1 narrative volume"),

    # Cover % atteinte
    ('canv.drawCentredString(MARGIN_L + col_w*1.5, kpi_y, "103,8%")',
     'canv.drawCentredString(MARGIN_L + col_w*1.5, kpi_y, "99,8%")',
     "Cover % atteinte"),
    ('canv.drawCentredString(MARGIN_L + col_w*1.5, kpi_y - 0.7*cm, "objectif volume S1 atteint")',
     'canv.drawCentredString(MARGIN_L + col_w*1.5, kpi_y - 0.7*cm, "objectif volume S1 atteint (hors pipeline)")',
     "Cover label"),
    ('canv.drawCentredString(MARGIN_L + col_w*1.5, kpi_y - 1.4*cm, "42 826 t réalisées vs 41 230 t")',
     'canv.drawCentredString(MARGIN_L + col_w*1.5, kpi_y - 1.4*cm, "41 148 t réalisées + 193 t en pipeline")',
     "Cover sub-label"),
    ('canv.drawCentredString(MARGIN_L + col_w*0.5, kpi_y - 1.4*cm, "93,0% de l\'objectif CA")',
     'canv.drawCentredString(MARGIN_L + col_w*0.5, kpi_y - 1.4*cm, "93,0% de l\'objectif CA (hors pipeline)")',
     "Cover CA sub-label"),

    # ---- Section 1.2 — Tableau catégories ----
    ('["INGREDIENTS", "Maïs, BELGOFOS, BELGOTOX, Lysine, Méthionine, Farine poisson, etc.", "4 060", "1 071", "5,8%"]',
     '["INGREDIENTS", "Maïs, BELGOFOS, BELGOTOX, Lysine, Méthionine, Farine poisson, etc.", "2 382", "1 071", "5,8%"]',
     "Section 1.2 INGREDIENTS volume"),
    ('["COMPLEMENT ALIM.", "Additifs liquides BELGO (HARMONY, VIT, PROTECT, KILL, etc.)", "3", "15", "0,0%"]',
     '["COMPLEMENT ALIM.", "Additifs liquides BELGO (HARMONY, VIT, PROTECT, KILL, etc.)", "3", "15", "0,0%"]',
     "Section 1.2 COMPLEMENT (no change)"),
    ('["DIVERS", "Contribution carburant, Pont bascule, Pierre à lécher, Sacs en réemploi, Manuels", "2", "13", "0,1%"]',
     '["DIVERS", "Contribution carburant, Pont bascule, Pierre à lécher, Sacs en réemploi, Manuels", "2", "17", "0,1%"]',
     "Section 1.2 DIVERS row (with PONT_BASCULE)"),

    # ---- Section 2.2 — Performance par catégorie (lignes 593-594) ----
    ('["INGREDIENTS", "4 060", "1 071", "598", "678,9%", "810", "132,2%"]',
     '["INGREDIENTS", "2 382", "1 071", "598", "398,4%", "810", "132,2%"]',
     "Section 2.2 INGREDIENTS row"),
    ('["COMPLEMENT ALIM.", "3", "15", "6", "49,3%", "16", "96,2%"]',
     '["COMPLEMENT ALIM.", "3", "15", "6", "49,3%", "16", "96,2%"]',
     "Section 2.2 COMPLEMENT (no change)"),

    # ---- Section 2.2 narrative (ligne 603) ----
    ("Le volume global dépasse l'objectif (103,9%) grâce aux TOURTEAUX (108,9%) et au Maïs (INGREDIENTS 679%). ",
     "Le volume global atteint 99,8% de l'objectif — légèrement sous, mais le pipeline en cours (192,8 t) le ferait passer à 102,3%. "
     "La sur-performance des TOURTEAUX (108,9%) et des INGREDIENTS (398%, porté par le Maïs) compense la sous-performance CONCENTRÉS (72,4%). ",
     "Section 2.2 narrative"),

    # ---- Section 3.2 — Comparison Q1/Q2 (ligne 653) ----
    ('["INGREDIENTS", "315", "1 970", "625%", "283", "2 090", "739%", "598", "4 060", "678,9%"]',
     '["INGREDIENTS", "315", "1 970", "625%", "283", "2 090", "739%", "598", "2 382", "398,4%"]',
     "Section 3.2 INGREDIENTS Q1/Q2 row"),

    # ---- Section 3.2bis — CA comparison (ligne 676) ----
    ('["INGREDIENTS", "810", "1 071", "+261", "132,2%"]',
     '["INGREDIENTS", "810", "1 071", "+261", "132,2%"]',
     "Section 3.2bis CA INGREDIENTS (no change)"),
    ('["DIVERS", "5", "17", "+12", "340,0%"]',
     '["DIVERS", "5", "17", "+12", "340,0%"]',
     "Section 3.2bis CA DIVERS (updated with PONT_BASCULE)"),

    # ---- Section 3.3 — INGREDIENTS narrative ----
    ("<b>INGREDIENTS (678,9% — sur-performance massive) :</b> les ventes atteignent 4 060 t "
     "contre 598 t d'objectif. Cette sur-performance est portée par le Maïs (M1051, 3 623 t) et les sels minéraux (BELGOFOS, BELGOTOX, ~300 t) ",
     "<b>INGREDIENTS (398,4% — sur-performance massive) :</b> les ventes atteignent 2 382 t "
     "contre 598 t d'objectif. Cette sur-performance est portée par le Maïs (M1051, 1 892 t hors régularisation stock) "
     "et les sels minéraux (BELGOFOS, BELGOTOX, ~300 t) ",
     "Section 3.3 INGREDIENTS narrative"),

    # ---- Section 3.5 — Taux atteinte narrative ----
    ("• <b>INGREDIENTS</b> : sur-performance massive (~679%) — l'objectif n'inclut pas le Maïs ni les sels minéraux (BELGOFOS/BELGOTOX).<br/>",
     "• <b>INGREDIENTS</b> : sur-performance massive (~398%) — l'objectif n'inclut pas le Maïs ni les sels minéraux (BELGOFOS/BELGOTOX).<br/>",
     "Section 3.5 INGREDIENTS narrative"),

    # ---- Section 3.9 — Tendance mensuelle ----
    ('["INGREDIENTS", "1 788", "81", "263", "898", "958", "71", "↗ Irrégulier (Maïs en gros lots)"]',
     '["INGREDIENTS", "89", "81", "275", "898", "968", "71", "↗ Irrégulier (Maïs en gros lots)"]',
     "Section 3.9 INGREDIENTS monthly row"),
    ("• <b>INGREDIENTS</b> : très volatil (Maïs + BELGOFOS vendus en gros lots irréguliers) — pic en janvier (1 788 t) puis creux.<br/>",
     "• <b>INGREDIENTS</b> : très volatil (Maïs + BELGOFOS vendus en gros lots irréguliers) — pic en avril-mai (898-968 t) puis creux.<br/>",
     "Section 3.9 INGREDIENTS narrative"),

    # ---- Section 4.7 — Prix INGREDIENTS (ligne 1268) ----
    ('["INGREDIENTS (mixte)", "4 060", "1 071", "263 800", "Le Maïs écrase le prix moyen (×10)"]',
     '["INGREDIENTS (mixte)", "2 382", "1 071", "449 600", "Le Maïs écrase le prix moyen (×10)"]',
     "Section 4.7 INGREDIENTS price row"),
    ("Le Maïs représente 89% du volume INGREDIENTS mais seulement 50% du CA. À 118 545 FCFA/t, il est 13× moins cher "
     "que les autres ingrédients (1 386 232 FCFA/t en moyenne pour les sels minéraux BELGOFOS/BELGOTOX et 1 510 582 FCFA/t pour les protéines). "
     "Le prix moyen INGREDIENTS de 263 800 FCFA/t est donc trompeur — ",
     "Le Maïs représente 79% du volume INGREDIENTS mais seulement 40% du CA. À 225 000 FCFA/t, il est 7× moins cher "
     "que les autres ingrédients (1 386 232 FCFA/t en moyenne pour les sels minéraux BELGOFOS/BELGOTOX et 1 510 582 FCFA/t pour les protéines). "
     "Le prix moyen INGREDIENTS de 449 600 FCFA/t est donc trompeur — ",
     "Section 4.7 Maïs narrative"),

    # ---- Section 12.2 — Forecast tables (lignes 1963-1966) ----
    ('["INGREDIENTS", "585", "348", "385", "423", "65,8%"]',
     '["INGREDIENTS", "585", "2 244", "2 281", "2 319", "391,3%"]',
     "Section 12.2 INGREDIENTS forecast row"),
    ('["COMPLEMENT ALIM.", "6", "1,7", "1,9", "2,2", "31,7%"]',
     '["COMPLEMENT ALIM.", "6", "2,1", "2,3", "2,6", "38,3%"]',
     "Section 12.2 COMPLEMENT ALIM forecast row"),

    # ---- Section 12.3 — CA HT forecast (lignes 2004-2005) ----
    ('["INGREDIENTS", "792", "654", "727", "800"]',
     '["INGREDIENTS", "792", "1 047", "1 066", "1 085"]',
     "Section 12.3 CA INGREDIENTS row"),
    ('["COMPLEMENT ALIM.", "16", "12", "13", "14"]',
     '["COMPLEMENT ALIM.", "16", "12", "13", "14"]',
     "Section 12.3 CA COMPLEMENT ALIM (no change)"),

    # ---- Section 12.4 — YoY comparison (lignes 2052-2053) ----
    ('["INGREDIENTS", "1 081", "4 060", "+275,7%", "1 243", "Boom du Maïs (3 623 t) + sels minéraux"]',
     '["INGREDIENTS", "392", "2 382", "+507,7%", "1 243", "Boom Maïs (1 892 t hors régul.) + sels minéraux"]',
     "Section 12.4 YoY INGREDIENTS row"),
    ('["COMPLEMENT ALIM.", "8", "3", "-60,0%", "2", "Déclin des liquides BELGO"]',
     '["COMPLEMENT ALIM.", "8", "3", "-60,0%", "2", "Déclin des liquides BELGO"]',
     "Section 12.4 YoY COMPLEMENT ALIM (no change)"),

    # ---- Section 12.4 narrative ----
    ("• <b>INGREDIENTS +275,7% (Boom du Maïs + sels minéraux)</b> : croissance tirée par le Maïs (3 623 t vs ~700 t en 2025) et l'intégration des BELGOFOS/BELGOTOX. Vérifier la soutenabilité de cette croissance.<br/>",
     "• <b>INGREDIENTS +507,7% (Boom Maïs + sels minéraux)</b> : croissance tirée par le Maïs (1 892 t vs ~250 t en 2025 hors régularisation) et l'intégration des BELGOFOS/BELGOTOX. Vérifier la soutenabilité de cette croissance.<br/>",
     "Section 12.4 INGREDIENTS YoY narrative"),

    # ---- Section 12 — Forecast YoY refined (lignes 2101-2106) ----
    ('["INGREDIENTS", "375", "+2,6%", "347", "385", "423", "585"]',
     '["INGREDIENTS", "375", "+507,7%", "2 244", "2 281", "2 319", "585"]',
     "Section 12 forecast YoY INGREDIENTS row"),
    ('["COMPLEMENT ALIM.", "3", "-25,6%", "1,7", "1,9", "2,2", "6"]',
     '["COMPLEMENT ALIM.", "3", "-8,8%", "2,1", "2,3", "2,6", "6"]',
     "Section 12 forecast YoY COMPLEMENT ALIM row"),
    ('["TOTAL", "37 339", "+21,9%", "46 453", "50 187", "53 921", "40 238"]',
     '["TOTAL", "37 339", "+17,1%", "50 777", "54 510", "58 244", "40 238"]',
     "Section 12 forecast YoY TOTAL row"),

    # ---- Section 12.5 — Rationnel recalibrage (rationnel par catégorie) ----
    ('["INGREDIENTS", "585", "466", "-20,3%", "Ajustement réaliste : S1=402 t (excluant Maïs comptoirs), forecast S2=385 t. "\n         "L\'objectif initial ne tenait pas compte de la volatilité du Maïs en gros lots. "\n         "Le recalibrage aligne sur le périmètre opérationnel réel."]',
     '["INGREDIENTS", "585", "466", "-20,3%", "Ajustement réaliste : S1=2 382 t (avec Maïs hors régularisation), forecast S2=2 281 t. "\n         "L\'objectif initial de 585 t ne tenait pas compte du Maïs ni des sels minéraux (BELGOFOS/BELGOTOX). "\n         "Le recalibrage reste sous le réel — à revoir à la hausse pour 2027."]',
     "Section 12.5 INGREDIENTS rationnel"),
    ('["COMPLEMENT ALIM.", "6", "3", "-52,0%", "Alignement sur le périmètre corrigé. Reclassification : BELGOFOS/BELGOTOX→INGREDIENTS, Pierre à lécher→DIVERS. "\n         "Il ne reste que les additifs liquides BELGO (3 t en S1). Recalibrage à 3 t = alignment sur le réalisé."]',
     '["COMPLEMENT ALIM.", "6", "3", "-52,0%", "Alignement sur le périmètre corrigé. Reclassification : BELGOFOS/BELGOTOX→INGREDIENTS, Pierre à lécher→DIVERS, ajout PONT_BASCULE/CONTRIBUTION_CARBURANT en DIVERS. "\n         "Il ne reste que les additifs liquides BELGO (3 t en S1). Recalibrage à 3 t = alignment sur le réalisé."]',
     "Section 12.5 COMPLEMENT ALIM rationnel"),
    # 12.5.1 constat 3 INGREDIENTS
    ("<b>3. INGREDIENTS en sur-performance (+2,6% YoY)</b> — l'objectif initial de 585 t paraissait sous-calibré. "
     "Le réalisé S1 (402 t, excluant comptoirs) et le forecast S2 (385 t) restent sous l'objectif — un ajustement à la baisse "
     "était nécessaire pour refléter la réalité opérationnelle (hors gros lots Maïs comptoirs).<br/><br/>",
     "<b>3. INGREDIENTS en sur-performance massive (+507,7% YoY)</b> — l'objectif initial de 585 t ne tenait pas compte "
     "du Maïs (1 892 t en S1 hors régularisation) ni des sels minéraux BELGOFOS/BELGOTOX (~300 t). "
     "Le réalisé S1 atteint 2 382 t (398% de l'objectif). Le recalibrage à 466 t reste très sous-estimé — à revoir massivement à la hausse pour 2027.<br/><br/>",
     "Section 12.5.1 INGREDIENTS constat"),

    # ---- Section 12.7 — Forecast vs recaled (lignes 2188-2193) ----
    ('["INGREDIENTS", "385", "466", "82,6%", "Sous-légère — forecast réaliste"]',
     '["INGREDIENTS", "2 281", "466", "489,5%", "Forecast très au-dessus — objectif sous-calibré"]',
     "Section 12.7 INGREDIENTS forecast vs recaled"),

    # ---- Section 12.6 — CA impact (ligne 2160-2165) ----
    ('["INGREDIENTS", "913", "728", "-185", "-20,3%"]',
     '["INGREDIENTS", "913", "728", "-185", "-20,3%"]',
     "Section 12.6 CA INGREDIENTS impact (no change)"),
    ('["DIVERS", "5", "17", "+12", "340,0%"]',
     '["DIVERS", "5", "13", "+8", "170,0%"]',
     "Section 12.6 CA DIVERS impact (with PONT_BASCULE Validée, kept at recalibré level)"),
]

# Apply
changes = 0
not_found = []
for old, new, desc in replacements:
    if old == new: continue
    if old in content:
        content = content.replace(old, new, 1)
        changes += 1
        print(f"✓ {desc}")
    else:
        not_found.append((desc, old[:80]))
        print(f"✗ NOT FOUND: {desc}")

print(f"\n=== {changes}/{len(replacements)} remplacements effectués ===")
if not_found:
    print(f"\n=== {len(not_found)} patterns non trouvés ===")

with open(PDF_SCRIPT, "w", encoding="utf-8") as f:
    f.write(content)
print(f"\n✓ Fichier sauvegardé")
