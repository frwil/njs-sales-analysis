"""
Map each product reference from the sales file to one of the 10 objective categories.

Categories from objectives file:
  - ALIMENT COMPLET
  - INGREDIENTS
  - COMPLEMENT ALIMENTAIRE
  - MATERIEL ELEVAGE
  - PREMIX
  - CONCENTRES
  - TOURTEAUX
  - ALVEOLE
  - DIVERS
  - Innovations
"""
import re
import json
from openpyxl import load_workbook
from collections import defaultdict

SRC = "/home/z/my-project/download/ventes_livrees.xlsx"

# Product categories mapping based on BELGOCAM product knowledge
# CB100, CB101 = CHICK BOOSTER -> ALIMENT COMPLET (aliment starter)
# CB200, CB201 = PIGLET BOOSTER -> ALIMENT COMPLET
# T102, T1021, T1023, T1024 = TOURTEAUX DE SOJA -> TOURTEAUX
# C101, C102, C1022, C103, C104, C1042, C1043, C1044, C105, C1053, C1054, C1055 = BELGO concentrés -> CONCENTRES
# P102N2, P104N2, P109 = PREMIX -> PREMIX
# I105, I1051, I106, I1061, I107, I1071 = INGREDIENTS (sulfate fer, methionine, lysine)
# F114, F1142, F1145, F1146, F1147 = FARINE DE POISSON -> INGREDIENTS
# B100, B1001 = BICARBONATE -> INGREDIENTS
# CF101, CF1012 = CARBONATE DE CALCIUM -> INGREDIENTS
# M1051 = MAIS -> INGREDIENTS
# S101 = SEL -> INGREDIENTS
# S0020 = SAC -> DIVERS
# ALAP25 = BELGO RABBIT -> CONCENTRES (it's a concentrate for rabbits)
# APCL2, APCL25, APCL3, APCL4.5, APCL4.55, APCL6, APCL65, APCL8, APCL85 = BELGO FISH CLARIA -> ALIMENT COMPLET (fish feed)
# CA001.1, CA002.1, CA003.1, CA004.1, CA005.1, CA006.1, CA007.1, CA008.1 = BELGO liquids -> COMPLEMENT ALIMENTAIRE
# E101, E1011, E1014 = BELGOTOX -> COMPLEMENT ALIMENTAIRE
# V300 = BELGOKILL -> COMPLEMENT ALIMENTAIRE
# P105, P1051, P1053 = BELGOFOS -> COMPLEMENT ALIMENTAIRE
# PL102 = PIERRE A LECHER -> COMPLEMENT ALIMENTAIRE
# MA100, MA101, MA90 = MANUELS -> DIVERS
# MAT001, MAT0023, MAT003-MAT073, MAT096 = MATERIEL ELEVAGE
# CT100 = CARTON A OEUFS -> ALVEOLE
# MAT017 = ALVEOLE BELGO -> ALVEOLE
# MAT073 = CAISSES A OEUFS -> ALVEOLE
# C108 = BELGO RUMINANT -> CONCENTRES
# CONTRIBUTION_CARBURANT = DIVERS

PRODUCT_CATEGORY = {
    # TOURTEAUX
    "T102": "TOURTEAUX", "T1021": "TOURTEAUX", "T1023": "TOURTEAUX", "T1024": "TOURTEAUX",

    # CONCENTRES (BELGO 10% + BELGO 5% + BELGO RABBIT + BELGO RUMINANT)
    "C101": "CONCENTRES", "C102": "CONCENTRES", "C1022": "CONCENTRES",
    "C103": "CONCENTRES", "C104": "CONCENTRES", "C1042": "CONCENTRES", "C1043": "CONCENTRES", "C1044": "CONCENTRES",
    "C105": "CONCENTRES", "C1053": "CONCENTRES", "C1054": "CONCENTRES", "C1055": "CONCENTRES",
    "C108": "CONCENTRES",  # BELGO RUMINANT
    "ALAP25": "CONCENTRES",  # BELGO RABBIT

    # ALIMENT COMPLET (Booster + Fish)
    "CB100": "ALIMENT COMPLET", "CB101": "ALIMENT COMPLET",  # Chick Booster
    "CB200": "ALIMENT COMPLET", "CB201": "ALIMENT COMPLET",  # Piglet Booster
    "APCL2": "ALIMENT COMPLET", "APCL25": "ALIMENT COMPLET", "APCL3": "ALIMENT COMPLET",
    "APCL4.5": "ALIMENT COMPLET", "APCL4.55": "ALIMENT COMPLET",
    "APCL6": "ALIMENT COMPLET", "APCL65": "ALIMENT COMPLET",
    "APCL8": "ALIMENT COMPLET", "APCL85": "ALIMENT COMPLET",

    # PREMIX
    "P102N2": "PREMIX", "P104N2": "PREMIX", "P109": "PREMIX",

    # INGREDIENTS (matières premières pour aliment composé)
    "I105": "INGREDIENTS", "I1051": "INGREDIENTS",  # Sulfate de fer
    "I106": "INGREDIENTS", "I1061": "INGREDIENTS",  # Methionine
    "I107": "INGREDIENTS", "I1071": "INGREDIENTS",  # Lysine
    "F114": "INGREDIENTS", "F1142": "INGREDIENTS", "F1145": "INGREDIENTS",
    "F1146": "INGREDIENTS", "F1147": "INGREDIENTS",  # Farine de poisson
    "B100": "INGREDIENTS", "B1001": "INGREDIENTS",  # Bicarbonate
    "CF101": "INGREDIENTS", "CF1012": "INGREDIENTS",  # Carbonate de calcium
    "M1051": "INGREDIENTS",  # Mais
    "S101": "INGREDIENTS",  # Sel

    # COMPLEMENT ALIMENTAIRE (additifs liquides + BELGOFOS + BELGOTOX)
    "CA001.1": "COMPLEMENT ALIMENTAIRE", "CA002.1": "COMPLEMENT ALIMENTAIRE",
    "CA003.1": "COMPLEMENT ALIMENTAIRE", "CA004.1": "COMPLEMENT ALIMENTAIRE",
    "CA005.1": "COMPLEMENT ALIMENTAIRE", "CA006.1": "COMPLEMENT ALIMENTAIRE",
    "CA007.1": "COMPLEMENT ALIMENTAIRE", "CA008.1": "COMPLEMENT ALIMENTAIRE",
    "E101": "COMPLEMENT ALIMENTAIRE", "E1011": "COMPLEMENT ALIMENTAIRE", "E1014": "COMPLEMENT ALIMENTAIRE",  # BELGOTOX
    "V300": "COMPLEMENT ALIMENTAIRE",  # BELGOKILL
    "P105": "COMPLEMENT ALIMENTAIRE", "P1051": "COMPLEMENT ALIMENTAIRE", "P1053": "COMPLEMENT ALIMENTAIRE",  # BELGOFOS
    "PL102": "COMPLEMENT ALIMENTAIRE",  # Pierre à lécher

    # MATERIEL ELEVAGE
    "MAT001": "MATERIEL ELEVAGE", "MAT0023": "MATERIEL ELEVAGE",
    "MAT003": "MATERIEL ELEVAGE", "MAT004": "MATERIEL ELEVAGE", "MAT005": "MATERIEL ELEVAGE",
    "MAT006": "MATERIEL ELEVAGE", "MAT007": "MATERIEL ELEVAGE", "MAT008": "MATERIEL ELEVAGE",
    "MAT009": "MATERIEL ELEVAGE", "MAT017": "MATERIEL ELEVAGE",  # actually ALVEOLE BELGO - reclassify below
    "MAT020": "MATERIEL ELEVAGE", "MAT030": "MATERIEL ELEVAGE", "MAT033": "MATERIEL ELEVAGE",
    "MAT039": "MATERIEL ELEVAGE", "MAT040": "MATERIEL ELEVAGE", "MAT042": "MATERIEL ELEVAGE",
    "MAT049": "MATERIEL ELEVAGE", "MAT050": "MATERIEL ELEVAGE", "MAT053": "MATERIEL ELEVAGE",
    "MAT054": "MATERIEL ELEVAGE", "MAT055": "MATERIEL ELEVAGE", "MAT060": "MATERIEL ELEVAGE",
    "MAT073": "MATERIEL ELEVAGE",  # actually caisses a oeufs
    "MAT096": "MATERIEL ELEVAGE",

    # ALVEOLE (cartons/caisses pour oeufs)
    "CT100": "ALVEOLE",  # Carton a oeufs
    # MAT017 ALVEOLE BELGO and MAT073 CAISSES A OEUFS → reclassify
    "MAT017": "ALVEOLE",
    "MAT073": "ALVEOLE",

    # DIVERS
    "S0020": "DIVERS",  # Sac en réemploi
    "CONTRIBUTION_CARBURANT": "DIVERS",
    "MA100": "DIVERS", "MA101": "DIVERS", "MA90": "DIVERS",  # Manuels
}

# Save mapping for reuse
with open("/home/z/my-project/scripts/product_category_map.json", "w", encoding="utf-8") as f:
    json.dump(PRODUCT_CATEGORY, f, ensure_ascii=False, indent=2)

# Now read sales file to verify all products are mapped
print("Reading sales file to verify mapping...")
wb = load_workbook(SRC, read_only=True, data_only=True)
all_products = set()
for sn in wb.sheetnames:
    ws = wb[sn]
    for row in ws.iter_rows(min_row=3, values_only=True):
        ref_prod = row[0] if len(row) > 0 else None
        if ref_prod is not None:
            all_products.add(str(ref_prod).strip())
wb.close()

print(f"\nTotal products in sales file: {len(all_products)}")
mapped = set(PRODUCT_CATEGORY.keys())
unmapped = all_products - mapped
extra = mapped - all_products
print(f"Mapped: {len(mapped & all_products)}")
print(f"Unmapped (in sales, not in mapping): {len(unmapped)}")
if unmapped:
    print(f"  Unmapped refs: {sorted(unmapped)}")
print(f"Extra (in mapping, not in sales): {len(extra)}")
if extra:
    print(f"  Extra refs: {sorted(extra)}")
