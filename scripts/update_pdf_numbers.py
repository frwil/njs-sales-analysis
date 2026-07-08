"""
Update build_pdf_report_v3.py with corrected numbers after recategorization.
- BELGOFOS/BELGOTOX → INGREDIENTS (was COMPLEMENT ALIMENTAIRE)
- PL102 → DIVERS (was COMPLEMENT ALIMENTAIRE)
- Liquids 1L → 1kg conversion (already handled by regex)
"""
import re

PDF_SCRIPT = "/home/z/my-project/scripts/build_pdf_report_v3.py"

with open(PDF_SCRIPT, "r", encoding="utf-8") as f:
    content = f.read()

# ===== REPLACEMENTS =====
# Each tuple: (old_str, new_str, description)
replacements = [
    # ---- COVER PAGE (line 207) ----
    ("42 816 t réalisées vs 41 230 t",
     "42 826 t réalisées vs 41 230 t",
     "Cover KPI volume"),
    
    # ---- SYNTHESE EXECUTIVE ----
    ("L'analyse porte sur 42 816 tonnes et 17,34 Md FCFA de ventes (Janvier-Juin 2026), ",
     "L'analyse porte sur 42 826 tonnes et 17,34 Md FCFA de ventes (Janvier-Juin 2026), ",
     "Synthese volume"),
    
    # Cover KPI cards (lines 432, 447, 580, 641)
    ('kpi_card("Volume S1", "42 816 t", "(103,8% obj)")',
     'kpi_card("Volume S1", "42 826 t", "(103,9% obj)")',
     "KPI card volume"),
    
    ("Avec 42 816 tonnes en S1 2026, le volume atteint 103,8% de l'objectif. ",
     "Avec 42 826 tonnes en S1 2026, le volume atteint 103,9% de l'objectif. ",
     "Synthese narrative volume"),
    
    ("le volume total atteint est de <b>42 816 tonnes</b>, soit <b>103,8% de l'objectif S1</b> (41 230 tonnes). ",
     "le volume total atteint est de <b>42 826 tonnes</b>, soit <b>103,9% de l'objectif S1</b> (41 230 tonnes). ",
     "Section 3.1 narrative volume"),
    
    # ---- SECTION 1.2 — Catégories de produits (lines 527-529, 540) ----
    ('["INGREDIENTS", "Maïs (M1051), Lysine, Méthionine, Farine poisson, etc.", "3 759", "678", "3,9%"]',
     '["INGREDIENTS", "Maïs, BELGOFOS, BELGOTOX, Lysine, Méthionine, Farine poisson, etc.", "4 060", "1 071", "5,8%"]',
     "Section 1.2 INGREDIENTS row"),
    
    ('["COMPLEMENT ALIM.", "BELGOFOS, BELGOTOX, BELGOKILL, compléments liquides", "306", "410", "2,4%"]',
     '["COMPLEMENT ALIM.", "Additifs liquides BELGO (HARMONY, VIT, PROTECT, KILL, etc.)", "3", "15", "0,0%"]',
     "Section 1.2 COMPLEMENT ALIM row"),
    
    # ---- SECTION 2.2 — Performance par catégorie (lines 593-594) ----
    ('["INGREDIENTS", "3 759", "678", "598", "628,5%", "810", "83,7%"]',
     '["INGREDIENTS", "4 060", "1 071", "598", "678,9%", "810", "132,2%"]',
     "Section 2.2 INGREDIENTS row"),
    
    ('["COMPLEMENT ALIM.", "306", "410", "6", "5104%", "16", "2600%"]',
     '["COMPLEMENT ALIM.", "3", "15", "6", "49,3%", "16", "96,2%"]',
     "Section 2.2 COMPLEMENT ALIM row"),
    
    # ---- SECTION 2.2 narrative (line 603) ----
    ("Le volume global dépasse l'objectif (103,8%) grâce aux TOURTEAUX (108,9%) et au Maïs (INGREDIENTS 628%). ",
     "Le volume global dépasse l'objectif (103,9%) grâce aux TOURTEAUX (108,9%) et au Maïs (INGREDIENTS 679%). ",
     "Section 2.2 narrative"),
    
    # ---- SECTION 2.3 / 3.1 — narrative (line 696) ----
    ("Le COMPLEMENT ALIMENTAIRE sur-performe massivement (2600%) — l'objectif CA de 16 M est largement sous-estimé ",
     "Le COMPLEMENT ALIMENTAIRE atteint 96% de son objectif CA (15,2 M vs 15,8 M) — l'objectif est bien calibré pour la gamme liquides. ",
     "Section 2.3 narrative COMPLEMENT ALIM"),
    
    # ---- SECTION 3.2 — Comparison Q1/Q2 (lines 653-654) ----
    ('["INGREDIENTS", "315", "1 970", "625%", "283", "1 789", "632%", "598", "3 759", "628,5%"]',
     '["INGREDIENTS", "315", "1 970", "625%", "283", "2 090", "739%", "598", "4 060", "678,9%"]',
     "Section 3.2 INGREDIENTS Q1/Q2 row"),
    
    ('["COMPLEMENT ALIM.", "3", "165", "5494%", "3", "141", "4714%", "6", "306", "5104%"]',
     '["COMPLEMENT ALIM.", "3", "0,4", "13%", "3", "2,6", "87%", "6", "3,0", "49,3%"]',
     "Section 3.2 COMPLEMENT ALIM Q1/Q2 row"),
    
    # ---- SECTION 3.2bis — CA comparison (lines 676-677) ----
    ('["INGREDIENTS", "810", "678", "-132", "83,7%"]',
     '["INGREDIENTS", "810", "1 071", "+261", "132,2%"]',
     "Section 3.2bis CA INGREDIENTS row"),
    
    ('["COMPLEMENT ALIM.", "16", "410", "+394", "2600%"]',
     '["COMPLEMENT ALIM.", "16", "15", "-1", "96,2%"]',
     "Section 3.2bis CA COMPLEMENT ALIM row"),
    
    # ---- SECTION 3.3 — Lecture par catégorie (lines 738-741) ----
    ("<b>INGREDIENTS (628,5% — sur-performance massive) :</b> les ventes atteignent 3 759 t "
     "contre 598 t d'objectif. Cette sur-performance est portée par le Maïs (M1051, 3 623 t) ",
     "<b>INGREDIENTS (678,9% — sur-performance massive) :</b> les ventes atteignent 4 060 t "
     "contre 598 t d'objectif. Cette sur-performance est portée par le Maïs (M1051, 3 623 t) ",
     "Section 3.3 INGREDIENTS narrative"),
    
    ("Les objectifs INGREDIENTS doivent être révisés massivement à la hausse pour intégrer le Maïs. ",
     "Les objectifs INGREDIENTS doivent être révisés massivement à la hausse pour intégrer le Maïs et les sels minéraux (BELGOFOS, BELGOTOX). ",
     "Section 3.3 INGREDIENTS objectives narrative"),
    
    # ---- SECTION 4 / 3.9 — Tendance mensuelle (lines 919, 921) ----
    ('["INGREDIENTS", "1 731", "28", "211", "850", "912", "27", "↗ Irrégulier (Maïs en grappe)"]',
     '["INGREDIENTS", "1 788", "81", "263", "898", "958", "71", "↗ Irrégulier (Maïs en gros lots)"]',
     "Section 3.9 INGREDIENTS monthly row"),
    
    ('["COMPLEMENT ALIM.", "58", "54", "53", "49", "47", "45", "↓ Léger déclin"]',
     '["COMPLEMENT ALIM.", "0,4", "0,3", "0,2", "0,6", "0,6", "0,8", "→ Stable (liquides)"]',
     "Section 3.9 COMPLEMENT ALIM monthly row"),
    
    # ---- SECTION 3.9 narrative (lines 930-932) ----
    ("• <b>INGREDIENTS</b> : très volatil (Maïs vendu en gros lots irréguliers) — pic en janvier (1 731 t) puis creux.<br/>",
     "• <b>INGREDIENTS</b> : très volatil (Maïs + BELGOFOS vendus en gros lots irréguliers) — pic en janvier (1 788 t) puis creux.<br/>",
     "Section 3.9 INGREDIENTS narrative"),
    
    ("• <b>COMPLEMENT ALIMENTAIRE</b> : léger déclin continu (58 → 45 t) — surveiller.",
     "• <b>COMPLEMENT ALIMENTAIRE</b> : stable à 0,4-0,8 t/mois (gamme liquides BELGO).",
     "Section 3.9 COMPLEMENT ALIM narrative"),
    
    # ---- SECTION 3.5 — Taux atteinte (line 806) ----
    ("• <b>INGREDIENTS</b> : stagnation critique (~22%), problème structurel.<br/>",
     "• <b>INGREDIENTS</b> : sur-performance massive (~679%) — l'objectif n'inclut pas le Maïs ni les sels minéraux (BELGOFOS/BELGOTOX).<br/>",
     "Section 3.5 INGREDIENTS narrative"),
    
    # ---- SECTION 4.7 — Prix INGREDIENTS (lines 1268, 1278-1281) ----
    ('["INGREDIENTS (mixte)", "3 759", "678", "180 337", "Le Maïs écrase le prix moyen (×8)"]',
     '["INGREDIENTS (mixte)", "4 060", "1 071", "263 800", "Le Maïs écrase le prix moyen (×10)"]',
     "Section 4.7 INGREDIENTS price row"),
    
    ("Le Maïs représente 96% du volume INGREDIENTS mais seulement 63% du CA. À 118 545 FCFA/t, il est 13× moins cher "
     "que les autres ingrédients (1 510 582 FCFA/t). Le prix moyen INGREDIENTS de 180 337 FCFA/t est donc trompeur — ",
     "Le Maïs représente 89% du volume INGREDIENTS mais seulement 50% du CA. À 118 545 FCFA/t, il est 13× moins cher "
     "que les autres ingrédients (1 510 582 FCFA/t pour les protéines et 1 386 232 FCFA/t pour les sels minéraux BELGOFOS/BELGOTOX). Le prix moyen INGREDIENTS de 263 800 FCFA/t est donc trompeur — ",
     "Section 4.7 Maïs narrative"),
    
    ("L'objectif INGREDIENTS (598 t) ne semble pas inclure le Maïs — d'où l'écart massif (628% d'atteinte). ",
     "L'objectif INGREDIENTS (598 t) ne semble pas inclure le Maïs ni les sels minéraux (BELGOFOS/BELGOTOX, ~300 t) — d'où l'écart massif (679% d'atteinte). ",
     "Section 4.7 INGREDIENTS objectives narrative"),
    
    # ---- SECTION 12.2 — Forecast tables (lines 1963-1966) ----
    ('["INGREDIENTS", "585", "74", "132", "220", "22,6%"]',
     '["INGREDIENTS", "585", "348", "385", "423", "65,8%"]',
     "Section 12.2 INGREDIENTS forecast row"),
    
    ('["COMPLEMENT ALIM.", "6", "264", "270", "276", "4505%"]',
     '["COMPLEMENT ALIM.", "6", "1,7", "1,9", "2,2", "31,7%"]',
     "Section 12.2 COMPLEMENT ALIM forecast row"),
    
    # ---- SECTION 12.3 — CA HT forecast (lines 2004-2005) ----
    ('["INGREDIENTS", "792", "587", "663", "740"]',
     '["INGREDIENTS", "792", "654", "727", "800"]',
     "Section 12.3 CA INGREDIENTS row"),
    
    ('["COMPLEMENT ALIM.", "16", "369", "410", "451"]',
     '["COMPLEMENT ALIM.", "16", "12", "13", "14"]',
     "Section 12.3 CA COMPLEMENT ALIM row"),
    
    # ---- SECTION 12.4 — YoY comparison (lines 2052-2053, 2069, 2071) ----
    ('["INGREDIENTS", "509", "3 759", "+637,7%", "572", "Boom du Maïs (3 623 t)"]',
     '["INGREDIENTS", "1 081", "4 060", "+275,7%", "1 243", "Boom du Maïs (3 623 t) + sels minéraux"]',
     "Section 12.4 YoY INGREDIENTS row"),
    
    ('["COMPLEMENT ALIM.", "4", "306", "+7360%", "3", "Boom (objectif sous-estimé)"]',
     '["COMPLEMENT ALIM.", "8", "3", "-60,0%", "2", "Déclin des liquides BELGO"]',
     "Section 12.4 YoY COMPLEMENT ALIM row"),
    
    ("• <b>INGREDIENTS +637,7% (grâce au Maïs 3 623 t non comptabilisé auparavant)</b> : effondrement de 257 tonnes. Vérifier si c'est une perte de marché ou un changement de stratégie produit.<br/>",
     "• <b>INGREDIENTS +275,7% (Boom du Maïs + sels minéraux)</b> : croissance tirée par le Maïs (3 623 t vs ~700 t en 2025) et l'intégration des BELGOFOS/BELGOTOX. Vérifier la soutenabilité de cette croissance.<br/>",
     "Section 12.4 INGREDIENTS YoY narrative"),
    
    ("• <b>COMPLEMENT ALIMENTAIRE +8225%</b> : explosion liée à la gamme BELGOFOS/BELGOTOX — objectif 2026 est à revoir à la hausse.",
     "• <b>COMPLEMENT ALIMENTAIRE -60%</b> : déclin des additifs liquides BELGO (3 t vs 8 t en 2025). L'objectif (6 t) reste ambitieux — à surveiller.",
     "Section 12.4 COMPLEMENT ALIM YoY narrative"),
    
    # ---- SECTION 12 — Forecast YoY refined (lines 2101-2106) ----
    ('["INGREDIENTS", "375", "-65,6%", "92", "129", "167", "585"]',
     '["INGREDIENTS", "375", "+2,6%", "347", "385", "423", "585"]',
     "Section 12 forecast YoY INGREDIENTS row"),
    
    ('["COMPLEMENT ALIM.", "3", "+8225%", "214", "214", "214", "6"]',
     '["COMPLEMENT ALIM.", "3", "-25,6%", "1,7", "1,9", "2,2", "6"]',
     "Section 12 forecast YoY COMPLEMENT ALIM row"),
    
    ('["TOTAL", "37 339", "+31,2%", "46 407", "50 141", "53 875", "40 238"]',
     '["TOTAL", "37 339", "+21,9%", "46 453", "50 187", "53 921", "40 238"]',
     "Section 12 forecast YoY TOTAL row"),
    
    # ---- SECTION 12.5 — Recalibrage table (lines 2134-2139) ----
    ('["INGREDIENTS", "585", "466", "-119", "-20,3%", "Baisse — ajustement réaliste"]',
     '["INGREDIENTS", "585", "466", "-119", "-20,3%", "Ajustement réaliste (S1=402 t, forecast=385 t)"]',
     "Section 12.5 INGREDIENTS recaled row"),
    
    ('["COMPLEMENT ALIM.", "6", "3", "-3", "-52,0%", "Toujours sous-estimé (270 t en S1)"]',
     '["COMPLEMENT ALIM.", "6", "3", "-3", "-52,0%", "Aligné sur réel S1 (3 t — additifs liquides)"]',
     "Section 12.5 COMPLEMENT ALIM recaled row"),
    
    # ---- SECTION 12.6 — CA impact (lines 2160, 2163) ----
    ('["INGREDIENTS", "913", "728", "-185", "-20,3%"]',
     '["INGREDIENTS", "913", "728", "-185", "-20,3%"]',
     "Section 12.6 CA INGREDIENTS impact (no change — CA objective unchanged)"),
    
    ('["COMPLEMENT ALIM.", "8", "4", "-4", "-52,0%"]',
     '["COMPLEMENT ALIM.", "16", "8", "-8", "-50,0%"]',
     "Section 12.6 CA COMPLEMENT ALIM impact"),
    
    # ---- SECTION 12.7 — Forecast vs recaled (lines 2188-2193) ----
    ('["INGREDIENTS", "129", "466", "27,7%", "Très sous — problème structurel"]',
     '["INGREDIENTS", "385", "466", "82,6%", "Sous-légère — forecast réaliste"]',
     "Section 12.7 INGREDIENTS forecast vs recaled"),
    
    ('["COMPLEMENT ALIM.", "214", "3", "7432%", "Objectif aberrant — à corriger"]',
     '["COMPLEMENT ALIM.", "1,9", "3", "63,3%", "Sous-légère — déclin liquides"]',
     "Section 12.7 COMPLEMENT ALIM forecast vs recaled"),
    
    # ---- SECTION 12.7 narrative (lines 2199-2203) ----
    ("• Les CONCENTRÉS ont un objectif ambitieux (11 650 t) que le forecast réaliste (9 160 t) n'atteint qu'à 78,6%. ",
     "• Les CONCENTRÉS ont un objectif ambitieux (11 650 t) que le forecast réaliste (9 160 t) n'atteint qu'à 78,6%. ",
     "Section 12.7 narrative (no change)"),
    
    ("• <b>Risque</b> : ce dépassement est entièrement porté par les TOURTEAUX. Si la rupture s'arrête, le total retomberait à ~37 000 t (92% de l'objectif recalibré).",
     "• <b>Risque</b> : ce dépassement est entièrement porté par les TOURTEAUX. Si la rupture s'arrête, le total retomberait à ~37 000 t (92% de l'objectif recalibré).",
     "Section 12.7 risk narrative (no change)"),
]

# Apply replacements
changes = 0
not_found = []
for old, new, desc in replacements:
    if old == new:
        continue  # skip no-op
    if old in content:
        content = content.replace(old, new, 1)
        changes += 1
        print(f"✓ {desc}")
    else:
        not_found.append((desc, old[:80]))
        print(f"✗ NOT FOUND: {desc}")
        print(f"   Pattern: {old[:80]}...")

print(f"\n=== {changes}/{len(replacements)} remplacements effectués ===")
if not_found:
    print(f"\n=== {len(not_found)} patterns non trouvés ===")

# Save
with open(PDF_SCRIPT, "w", encoding="utf-8") as f:
    f.write(content)
print(f"\n✓ Fichier sauvegardé: {PDF_SCRIPT}")
