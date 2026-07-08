"""
Applique toutes les corrections:
1. CF101 → poids 1 kg (confirmé par ratio prix avec CF1012)
2. S101 (Sel) → poids 1 kg (sachets de 1 kg)
3. M1051 (Maïs) → exclure lignes à CA=0 (régularisation stock SPC)
4. PONT_BASCULE → ajouter dans product_category_map.json (catégorie DIVERS)

Output: table de vérification INGREDIENTS avec clés de conversion
"""
import openpyxl
import json
import re
from collections import defaultdict

# Charger mapping
with open('/home/z/my-project/scripts/product_category_map.json') as f:
    cat_map = json.load(f)

# Ajouter PONT_BASCULE en DIVERS
cat_map["PONT_BASCULE"] = "DIVERS"
with open('/home/z/my-project/scripts/product_category_map.json', 'w', encoding='utf-8') as f:
    json.dump(cat_map, f, ensure_ascii=False, indent=2)
print("✓ PONT_BASCULE ajouté dans product_category_map.json (catégorie DIVERS)")

# === POIDS MANUELS (overrides pour produits mal configurés) ===
MANUAL_WEIGHTS = {
    "M1051": 50.0,    # Maïs en sacs de 50 kg (poids non indiqué dans la description)
    "CF101": 1.0,     # Carbonate de calcium 1 kg (confirmé par ratio prix avec CF1012)
    "S101": 1.0,      # Sel en sachets de 1 kg
}

# === PARSER DE POIDS (avec overrides manuels) ===
WEIGHT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(KG|Kg|kg|KG|GRAMMES?|G|L)\b", re.IGNORECASE)
def parse_weight_kg(ref, desc):
    """Retourne le poids en kg, avec overrides manuels pour produits mal configurés."""
    ref_str = str(ref).strip() if ref else ""
    if ref_str in MANUAL_WEIGHTS:
        return MANUAL_WEIGHTS[ref_str]
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

# === LIRE VENTES 2026 (S1) ===
wb = openpyxl.load_workbook('/home/z/my-project/upload/ventes janv a juin 2026.xlsx', read_only=True, data_only=True)

# Compter PONT_BASCULE et autres DIVERS
divers_products = defaultdict(lambda: {"qte": 0, "ca": 0, "lignes": 0, "descs": set()})

# Compter INGREDIENTS
ingredients_products = defaultdict(lambda: {"qte": 0, "ca": 0, "lignes": 0, "descs": set(), "vol_kg": 0, "lignes_prix_zero": 0, "qte_prix_zero": 0})

# Stats Maïs
mais_stats = {
    "total_lignes": 0,
    "total_qte": 0,
    "total_ca": 0,
    "lignes_prix_zero": 0,
    "qte_prix_zero": 0,
    "ca_prix_zero": 0,
}

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    if ws.cell(2, 1).value != 'Réf. produit':
        continue
    for row in ws.iter_rows(min_row=3, values_only=True):
        ref = row[0] if len(row) > 0 else None
        desc = row[1] if len(row) > 1 else None
        qte = row[2] if len(row) > 2 else 0
        ca = row[8] if len(row) > 8 else 0
        etat = row[15] if len(row) > 15 else None
        if not ref or etat != 'Livrée':
            continue
        ref_str = str(ref).strip()
        try:
            q = float(qte) if qte else 0
            c = float(ca) if ca else 0
        except:
            continue
        
        # Suivi Maïs
        if ref_str == "M1051":
            mais_stats["total_lignes"] += 1
            mais_stats["total_qte"] += q
            mais_stats["total_ca"] += c
            if c == 0:
                mais_stats["lignes_prix_zero"] += 1
                mais_stats["qte_prix_zero"] += q
                mais_stats["ca_prix_zero"] += c
                continue  # exclure cette ligne
        
        # Catégorie
        cat = cat_map.get(ref_str, "DIVERS")
        weight = parse_weight_kg(ref_str, desc)
        vol_kg = q * weight
        
        if cat == "INGREDIENTS":
            ingredients_products[ref_str]["qte"] += q
            ingredients_products[ref_str]["ca"] += c
            ingredients_products[ref_str]["lignes"] += 1
            ingredients_products[ref_str]["descs"].add(str(desc) if desc else '')
            ingredients_products[ref_str]["vol_kg"] += vol_kg
        elif cat == "DIVERS":
            divers_products[ref_str]["qte"] += q
            divers_products[ref_str]["ca"] += c
            divers_products[ref_str]["lignes"] += 1
            divers_products[ref_str]["descs"].add(str(desc) if desc else '')

wb.close()

# === AFFICHAGE INGREDIENTS — LISTE COMPLÈTE AVEC CLÉS DE CONVERSION ===
print("\n" + "=" * 120)
print("LISTE COMPLÈTE DES PRODUITS INGREDIENTS — S1 2026 (CORRIGÉ)")
print("=" * 120)
print(f"{'Réf':<10} {'Description':<50} {'Clé conv.':<15} {'Poids':>8} {'Sacs':>10} {'Tonnes':>10} {'CA HT (M)':>12} {'Prix (k/t)':>11} {'Lignes':>8}")
print("-" * 120)

total_sacs = 0
total_tonnes = 0
total_ca = 0
total_lignes = 0

# Trier par volume décroissant
for ref in sorted(ingredients_products.keys(), key=lambda r: -ingredients_products[r]["vol_kg"]):
    p = ingredients_products[ref]
    descs = list(p["descs"])
    desc_str = descs[0] if descs else ''
    if len(descs) > 1:
        desc_str = f"{descs[0]} (+{len(descs)-1} var.)"
    
    # Déterminer clé de conversion
    if ref in MANUAL_WEIGHTS:
        cle = "MANUEL"
        poids = MANUAL_WEIGHTS[ref]
    else:
        weights_found = set()
        for d in descs:
            m = WEIGHT_RE.search(d)
            if m:
                val_str, unit = m.group(1), m.group(2).upper()
                val = float(val_str)
                if unit == "KG":
                    weights_found.add(val)
                elif unit in ("G", "GRAMME", "GRAMMES"):
                    weights_found.add(val/1000.0)
                elif unit == "L":
                    weights_found.add(val)
        if weights_found:
            cle = "REGEX"
            poids = max(weights_found)
        else:
            cle = "❌ AUCUN"
            poids = 0
    
    tonnes = p["vol_kg"] / 1000.0
    ca_m = p["ca"] / 1e6
    prix_k_t = (p["ca"] / p["vol_kg"]) if p["vol_kg"] > 0 else 0
    
    print(f"{ref:<10} {desc_str[:50]:<50} {cle:<15} {poids:>8.2f} {p['qte']:>10.0f} {tonnes:>10.2f} {ca_m:>12.2f} {prix_k_t:>11.0f} {p['lignes']:>8}")
    
    total_sacs += p["qte"]
    total_tonnes += tonnes
    total_ca += p["ca"]
    total_lignes += p["lignes"]

print("-" * 120)
print(f"{'TOTAL':<10} {'':<50} {'':<15} {'':>8} {total_sacs:>10.0f} {total_tonnes:>10.2f} {total_ca/1e6:>12.2f} {'':>11} {total_lignes:>8}")

print(f"\nVolume total INGREDIENTS (S1 2026, CORRIGÉ): {total_tonnes:.1f} t")
print(f"CA HT total: {total_ca/1e6:.1f} M FCFA")
print(f"Prix moyen: {total_ca/total_tonnes/1000:.0f} FCFA/t" if total_tonnes > 0 else "")

# === MAÏS — DÉTAIL DES CORRECTIONS ===
print("\n" + "=" * 120)
print("MAÏS (M1051) — DÉTAIL DES CORRECTIONS")
print("=" * 120)
print(f"Avant correction:")
print(f"  Lignes totales: {mais_stats['total_lignes']}")
print(f"  Qté totale: {mais_stats['total_qte']:.0f} sacs")
print(f"  Volume total: {mais_stats['total_qte']*50/1000:.1f} t")
print(f"  CA total: {mais_stats['total_ca']:,.0f} FCFA")
print(f"\nLignes à CA=0 (régularisation stock — EXCLUES):")
print(f"  Lignes: {mais_stats['lignes_prix_zero']}")
print(f"  Qté: {mais_stats['qte_prix_zero']:.0f} sacs")
print(f"  Volume: {mais_stats['qte_prix_zero']*50/1000:.1f} t")
print(f"\nAprès correction:")
print(f"  Lignes conservées: {mais_stats['total_lignes'] - mais_stats['lignes_prix_zero']}")
print(f"  Qté conservée: {mais_stats['total_qte'] - mais_stats['qte_prix_zero']:.0f} sacs")
print(f"  Volume conservé: {(mais_stats['total_qte'] - mais_stats['qte_prix_zero'])*50/1000:.1f} t")
print(f"  CA conservé: {mais_stats['total_ca']:,.0f} FCFA (inchangé)")

# === DIVERS — NOUVEAU PONT_BASCULE ===
print("\n" + "=" * 120)
print("DIVERS — INCLUT MAINTENANT PONT_BASCULE")
print("=" * 120)
print(f"{'Réf':<25} {'Description':<35} {'CA (M FCFA)':>12} {'Lignes':>8}")
print("-" * 80)
total_ca_divers = 0
for ref in sorted(divers_products.keys()):
    p = divers_products[ref]
    descs = list(p["descs"])
    desc_str = descs[0] if descs else ''
    print(f"{ref:<25} {desc_str[:35]:<35} {p['ca']/1e6:>12.2f} {p['lignes']:>8}")
    total_ca_divers += p["ca"]
print("-" * 80)
print(f"{'TOTAL DIVERS':<25} {'':<35} {total_ca_divers/1e6:>12.2f} M FCFA")

# === SAUVEGARDER LES POIDS MANUELS POUR RÉUTILISATION ===
with open('/home/z/my-project/scripts/manual_weights.json', 'w', encoding='utf-8') as f:
    json.dump(MANUAL_WEIGHTS, f, ensure_ascii=False, indent=2)
print(f"\n✓ Poids manuels sauvegardés dans manual_weights.json")
print(json.dumps(MANUAL_WEIGHTS, indent=2))
