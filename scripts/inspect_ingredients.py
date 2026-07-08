import openpyxl
import json
import re
from collections import defaultdict

# Charger le mapping produit→catégorie
with open('/home/z/my-project/scripts/product_category_map.json') as f:
    cat_map = json.load(f)

# Charger ventes livrées sur les 6 mois
wb = openpyxl.load_workbook('/home/z/my-project/upload/ventes janv a juin 2026.xlsx', data_only=True)

# Parser de poids
WEIGHT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(KG|Kg|kg|KG|GRAMMES?|G|L)\b", re.IGNORECASE)
def parse_weight_kg(desc):
    if not desc: return 0.0
    matches = WEIGHT_RE.findall(str(desc))
    if not matches: return 0.0
    val_str, unit = matches[-1]
    val = float(val_str)
    u = unit.upper()
    if u == "KG": return val
    if u in ("G", "GRAMME", "GRAMMES"): return val/1000.0
    if u == "L": return val
    return 0.0

# Cas spécial: Maïs = 50kg (pas de poids dans la description)
MAIS_WEIGHT = 50.0

ingredients_products = defaultdict(lambda: {"vol_kg": 0.0, "ca_ht": 0.0, "descs": set(), "lignes": 0})

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    if ws.cell(2, 1).value != 'Réf. produit':
        continue
    for row in range(3, ws.max_row + 1):
        ref = ws.cell(row, 1).value
        desc = ws.cell(row, 2).value
        qte = ws.cell(row, 3).value
        ca = ws.cell(row, 9).value
        etat = ws.cell(row, 16).value
        if not ref or etat != 'Livrée':
            continue
        ref_str = str(ref).strip()
        if ref_str in cat_map and cat_map[ref_str] == "INGREDIENTS":
            try:
                q = float(qte) if qte else 0
                c = float(ca) if ca else 0
            except:
                continue
            # Poids
            if ref_str == "M1051":
                weight = MAIS_WEIGHT
            else:
                weight = parse_weight_kg(desc)
            vol_kg = q * weight
            
            ingredients_products[ref_str]["vol_kg"] += vol_kg
            ingredients_products[ref_str]["ca_ht"] += c
            ingredients_products[ref_str]["descs"].add(str(desc) if desc else '')
            ingredients_products[ref_str]["lignes"] += 1

print("=" * 110)
print("PRODUITS DANS LA CATÉGORIE INGREDIENTS (ventes livrées S1 2026, TOUS CLIENTS)")
print("=" * 110)
print(f"{'Réf':<10} {'Description':<55} {'Poids (kg)':>10} {'Sacs':>8} {'Tonnes':>10} {'CA HT (M FCFA)':>16} {'Lignes':>8}")
print("-" * 110)

total_sacs = 0
total_tonnes = 0
total_ca = 0
total_lignes = 0

for ref in sorted(ingredients_products.keys()):
    p = ingredients_products[ref]
    descs = list(p["descs"])
    desc_str = descs[0] if descs else ''
    if len(descs) > 1:
        desc_str = f"{descs[0]} (+{len(descs)-1} variantes)"
    
    # Extraire poids
    if ref == "M1051":
        poids = MAIS_WEIGHT
        poids_str = f"{poids} (manuel)"
    else:
        weights = set()
        for d in descs:
            m = WEIGHT_RE.search(d)
            if m:
                val_str, unit = m.group(1), m.group(2).upper()
                val = float(val_str)
                if unit == "KG":
                    weights.add(val)
                elif unit in ("G", "GRAMME", "GRAMMES"):
                    weights.add(val/1000.0)
                elif unit == "L":
                    weights.add(val)
        poids = max(weights) if weights else 0
        poids_str = str(weights) if weights else "?"
    
    sacs = sum(1 for _ in range(int(p["lignes"])))  # approx
    # Better: compute sacs from qte sums
    tonnes = p["vol_kg"] / 1000.0
    ca_m = p["ca_ht"] / 1e6
    print(f"{ref:<10} {desc_str[:55]:<55} {poids_str:>10} {'':>8} {tonnes:>10.2f} {ca_m:>16.2f} {p['lignes']:>8}")
    total_tonnes += tonnes
    total_ca += p["ca_ht"]
    total_lignes += p["lignes"]

print("-" * 110)
print(f"{'TOTAL':<10} {'':<55} {'':>10} {'':>8} {total_tonnes:>10.2f} {total_ca/1e6:>16.2f} {total_lignes:>8}")
print()
print(f"Total tonnes (S1 2026, tous clients): {total_tonnes:.1f} t")
print(f"Total CA HT (S1 2026, tous clients): {total_ca/1e6:.1f} M FCFA")
print(f"Prix moyen: {total_ca/total_tonnes/1000:.0f} FCFA/t" if total_tonnes > 0 else "")
print()
print(f"Nombre de produits INGREDIENTS: {len(ingredients_products)}")

# Vérifier les reclassifications
print("\n" + "=" * 110)
print("VÉRIFICATION DES RECLASSIFICATIONS")
print("=" * 110)
print(f"\nBELGOFOS (P105, P1051, P1053) — doivent être INGREDIENTS:")
for ref in ["P105", "P1051", "P1053"]:
    cat = cat_map.get(ref, "?")
    in_list = ref in ingredients_products
    print(f"  {ref}: catégorie={cat}  présent dans ventes S1={in_list}")

print(f"\nBELGOTOX (E101, E1011, E1014) — doivent être INGREDIENTS:")
for ref in ["E101", "E1011", "E1014"]:
    cat = cat_map.get(ref, "?")
    in_list = ref in ingredients_products
    print(f"  {ref}: catégorie={cat}  présent dans ventes S1={in_list}")

print(f"\nAutres INGREDIENTS (devraient tous être présents):")
other_ings = ["I105", "I1051", "I106", "I1061", "I107", "I1071", "F114", "F1142", "F1145", "F1146", "F1147", "B100", "B1001", "CF101", "CF1012", "M1051", "S101"]
for ref in other_ings:
    cat = cat_map.get(ref, "?")
    in_list = ref in ingredients_products
    if in_list:
        p = ingredients_products[ref]
        print(f"  {ref}: catégorie={cat}  ventes S1: {p['vol_kg']/1000:.1f} t, {p['ca_ht']/1e6:.1f} M")
    else:
        print(f"  {ref}: catégorie={cat}  ⚠️ AUCUNE VENTE LIVRÉE en S1 2026")

