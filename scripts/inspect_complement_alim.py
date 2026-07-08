import openpyxl
from collections import defaultdict
import re

# Charger le mapping produit→catégorie
import json
with open('/home/z/my-project/scripts/product_category_map.json') as f:
    cat_map = json.load(f)

# Charger ventes livrées sur les 6 mois (concat feuilles)
wb = openpyxl.load_workbook('/home/z/my-project/upload/ventes janv a juin 2026.xlsx', data_only=True)
complement_products = defaultdict(lambda: {"volume_sacs": 0, "ca_ht": 0, "descs": set(), "lignes": 0})

# Également collecter les descriptions pour vérifier le poids
all_complement_descs = defaultdict(lambda: {"count": 0, "sacs": 0, "ca": 0})

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    # Vérifier si cette feuille a le bon header
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
        if ref_str in cat_map and 'COMPLEMENT' in cat_map[ref_str].upper():
            try:
                q = float(qte) if qte else 0
                c = float(ca) if ca else 0
            except:
                continue
            complement_products[ref_str]["volume_sacs"] += q
            complement_products[ref_str]["ca_ht"] += c
            complement_products[ref_str]["descs"].add(str(desc) if desc else '')
            complement_products[ref_str]["lignes"] += 1
            
            if desc:
                all_complement_descs[str(desc)]["count"] += 1
                all_complement_descs[str(desc)]["sacs"] += q
                all_complement_descs[str(desc)]["ca"] += c

print("=" * 100)
print("PRODUITS DANS LA CATÉGORIE COMPLEMENT ALIMENTAIRE (ventes livrées S1 2026)")
print("=" * 100)
print(f"{'Réf':<10} {'Description':<50} {'Sacs':>10} {'CA HT (FCFA)':>15} {'Lignes':>8}")
print("-" * 100)
total_sacs = 0
total_ca = 0
for ref in sorted(complement_products.keys()):
    p = complement_products[ref]
    descs = list(p["descs"])
    desc_str = descs[0] if descs else ''
    if len(descs) > 1:
        desc_str = f"{descs[0]} (+{len(descs)-1} autres)"
    print(f"{ref:<10} {desc_str[:50]:<50} {p['volume_sacs']:>10.0f} {p['ca_ht']:>15,.0f} {p['lignes']:>8}")
    total_sacs += p["volume_sacs"]
    total_ca += p["ca_ht"]
print("-" * 100)
print(f"{'TOTAL':<10} {'':<50} {total_sacs:>10.0f} {total_ca:>15,.0f}")

print("\n")
print("=" * 100)
print("DÉTAIL DES DESCRIPTIONS (pour extraction du poids)")
print("=" * 100)
print(f"{'Description':<60} {'Lignes':>8} {'Sacs':>10} {'CA HT':>15}")
print("-" * 100)
for desc in sorted(all_complement_descs.keys(), key=lambda x: -all_complement_descs[x]["ca"]):
    d = all_complement_descs[desc]
    print(f"{desc[:60]:<60} {d['count']:>8} {d['sacs']:>10.0f} {d['ca']:>15,.0f}")

# Vérifier le fichier de poids
print("\n")
print("=" * 100)
print("CONVERSION EN TONNES")
print("=" * 100)

# Extraire poids des descriptions
import re
weights_per_ref = defaultdict(set)
for ref in complement_products:
    for desc in complement_products[ref]["descs"]:
        # Chercher patterns: "50Kg", "25Kg", "5Kg", "1 Kg", etc.
        m = re.search(r'(\d+(?:\.\d+)?)\s*Kg', desc, re.IGNORECASE)
        if m:
            weights_per_ref[ref].add(float(m.group(1)))

print(f"{'Réf':<10} {'Poids extraits (kg)':<30} {'Poids retenu':<15} {'Sacs':>10} {'Tonnes':>10}")
print("-" * 80)
total_tonnes = 0
for ref in sorted(complement_products.keys()):
    sacs = complement_products[ref]["volume_sacs"]
    poids = weights_per_ref[ref]
    poids_str = str(sorted(poids))
    poids_retenu = max(poids) if poids else None
    tonnes = (sacs * poids_retenu / 1000) if poids_retenu else None
    total_tonnes += tonnes if tonnes else 0
    print(f"{ref:<10} {poids_str:<30} {str(poids_retenu):<15} {sacs:>10.0f} {tonnes if tonnes else 'N/A':>10}")
print("-" * 80)
print(f"{'TOTAL':<10} {'':<30} {'':<15} {total_sacs:>10.0f} {total_tonnes:>10.2f}")
