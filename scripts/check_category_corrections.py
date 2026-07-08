import json
from collections import defaultdict

with open('/home/z/my-project/scripts/product_category_map.json') as f:
    cat_map = json.load(f)

# Lister toutes les catégories actuelles
cats = defaultdict(list)
for p, c in cat_map.items():
    cats[c].append(p)
print("Catégories actuelles:")
for c in sorted(cats.keys()):
    print(f"  {c}: {len(cats[c])} produits")

# Afficher ce qui est dans DIVERS actuellement (s'il existe)
if 'DIVERS' in cats:
    print("\nDIVERS contient déjà:", cats['DIVERS'])
