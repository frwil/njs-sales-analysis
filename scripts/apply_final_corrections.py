"""
Applique toutes les corrections finales:
1. CF101 → poids 1 kg (manuel)
2. S101 → poids 1 kg (manuel)
3. M1051 → exclure lignes à CA=0 (régularisation stock)
4. PONT_BASCULE → catégorie DIVERS, inclure état "Validée"
5. CONTRIBUTION_CARBURANT → inclure état "Validée" en plus de "Livrée"
"""
import openpyxl
import json
import re
from collections import defaultdict

# Charger mapping
with open('/home/z/my-project/scripts/product_category_map.json') as f:
    cat_map = json.load(f)

# === POIDS MANUELS ===
MANUAL_WEIGHTS = {
    "M1051": 50.0,    # Maïs en sacs de 50 kg
    "CF101": 1.0,     # Carbonate de calcium 1 kg (confirmé par ratio prix)
    "S101": 1.0,      # Sel en sachets de 1 kg
}

# === SERVICES AVEC ÉTAT "VALIDÉE" ===
# Ces "services" ne sont jamais "Livrés" (pas de livraison physique)
# mais seulement "Validés". Pour les inclure dans le CA, on élargit le filtre.
SERVICES_VALIDEE = {"PONT_BASCULE", "CONTRIBUTION_CARBURANT"}

# === PARSER DE POIDS ===
WEIGHT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(KG|Kg|kg|KG|GRAMMES?|G|L)\b", re.IGNORECASE)
def parse_weight_kg(ref, desc):
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

def should_include(ref, etat, ca):
    """Filtre pour déterminer si une ligne doit être incluse."""
    ref_str = str(ref).strip() if ref else ""
    etat_str = str(etat) if etat else ""
    
    # Pour M1051 (Maïs) — exclure lignes à CA=0 (régularisation stock)
    if ref_str == "M1051" and (ca is None or float(ca) == 0):
        return False
    
    # État standard = "Livrée"
    if etat_str == "Livrée":
        return True
    
    # Pour les services (PONT_BASCULE, CONTRIBUTION_CARBURANT) — inclure "Validée"
    if ref_str in SERVICES_VALIDEE and etat_str == "Validée":
        return True
    
    # Tout le reste (Annulée, Brouillon, En cours, vide) — exclure
    return False

# === LIRE VENTES 2026 (S1) ===
wb = openpyxl.load_workbook('/home/z/my-project/upload/ventes janv a juin 2026.xlsx', read_only=True, data_only=True)

# Stats par catégorie
cat_stats = defaultdict(lambda: {"vol_kg": 0, "ca": 0, "lignes": 0})
product_stats = defaultdict(lambda: {"vol_kg": 0, "ca": 0, "lignes": 0, "descs": set(), "etats": set()})

# Stats Maïs (avant/après)
mais_stats = {
    "total_lignes": 0,
    "excluded_lignes": 0,
    "included_lignes": 0,
    "total_qte": 0,
    "excluded_qte": 0,
}

# Stats par état (vérification)
etats_count = defaultdict(lambda: {"ca": 0, "lignes": 0})

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    if ws.cell(2, 1).value != 'Réf. produit':
        continue
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row or len(row) < 16: continue
        ref = row[0]
        desc = row[1]
        qte = row[2]
        ca = row[8]
        etat = row[15]
        if not ref: continue
        ref_str = str(ref).strip()
        
        # Stats Maïs
        if ref_str == "M1051":
            mais_stats["total_lignes"] += 1
            mais_stats["total_qte"] += float(qte) if qte else 0
        
        # Filtre
        if not should_include(ref_str, etat, ca):
            if ref_str == "M1051":
                mais_stats["excluded_lignes"] += 1
                mais_stats["excluded_qte"] += float(qte) if qte else 0
            continue
        
        try:
            q = float(qte) if qte else 0
            c = float(ca) if ca else 0
        except:
            continue
        
        # Stats par état (post-filtre)
        etats_count[str(etat)]["ca"] += c
        etats_count[str(etat)]["lignes"] += 1
        
        # Catégorie et poids
        cat = cat_map.get(ref_str, "DIVERS")
        weight = parse_weight_kg(ref_str, desc)
        vol_kg = q * weight
        
        cat_stats[cat]["vol_kg"] += vol_kg
        cat_stats[cat]["ca"] += c
        cat_stats[cat]["lignes"] += 1
        
        product_stats[ref_str]["vol_kg"] += vol_kg
        product_stats[ref_str]["ca"] += c
        product_stats[ref_str]["lignes"] += 1
        product_stats[ref_str]["descs"].add(str(desc) if desc else '')
        product_stats[ref_str]["etats"].add(str(etat) if etat else '')

wb.close()

# === RÉCAPITULATIF PAR ÉTAT (post-filtre) ===
print("=" * 80)
print("1. LIGNES INCLUSES PAR ÉTAT (post-filtre corrigé)")
print("=" * 80)
print(f"{'État':<15} {'Lignes':>8} {'CA (M FCFA)':>15}")
print("-" * 40)
total_lignes_etat = 0
total_ca_etat = 0
for etat, d in sorted(etats_count.items()):
    print(f"{etat:<15} {d['lignes']:>8} {d['ca']/1e6:>15.2f}")
    total_lignes_etat += d['lignes']
    total_ca_etat += d['ca']
print("-" * 40)
print(f"{'TOTAL':<15} {total_lignes_etat:>8} {total_ca_etat/1e6:>15.2f}")

# === MAÏS — DÉTAIL ===
print("\n" + "=" * 80)
print("2. MAÏS (M1051) — CORRECTION")
print("=" * 80)
mais_stats["included_lignes"] = mais_stats["total_lignes"] - mais_stats["excluded_lignes"]
print(f"  Lignes totales (tous états): {mais_stats['total_lignes']}")
print(f"  Lignes exclues (CA=0, régularisation): {mais_stats['excluded_lignes']}")
print(f"  Lignes incluses: {mais_stats['included_lignes']}")
print(f"  Qté totale: {mais_stats['total_qte']:.0f} sacs")
print(f"  Qté exclue: {mais_stats['excluded_qte']:.0f} sacs")
print(f"  Qté incluse: {mais_stats['total_qte'] - mais_stats['excluded_qte']:.0f} sacs")
print(f"  Volume inclus: {(mais_stats['total_qte'] - mais_stats['excluded_qte'])*50/1000:.1f} t")

# === TOTAL GLOBAL ===
print("\n" + "=" * 80)
print("3. TOTAL GLOBAL S1 2026 (APRÈS CORRECTIONS)")
print("=" * 80)
total_vol = 0
total_ca = 0
print(f"{'Catégorie':<28} {'Volume (t)':>12} {'CA (M FCFA)':>14} {'Lignes':>10}")
print("-" * 70)
for cat in sorted(cat_stats.keys()):
    d = cat_stats[cat]
    vol_t = d["vol_kg"] / 1000.0
    ca_m = d["ca"] / 1e6
    print(f"{cat:<28} {vol_t:>12.1f} {ca_m:>14.2f} {d['lignes']:>10}")
    total_vol += vol_t
    total_ca += d["ca"]
print("-" * 70)
print(f"{'TOTAL':<28} {total_vol:>12.1f} {total_ca/1e6:>14.2f}")

# === DIVERS — DÉTAIL ===
print("\n" + "=" * 80)
print("4. DIVERS — DÉTAIL POST-CORRECTION")
print("=" * 80)
print(f"{'Réf':<28} {'Description':<35} {'États':<25} {'CA (M)':>10} {'Lignes':>8}")
print("-" * 110)
total_ca_divers = 0
for ref in sorted([r for r in product_stats.keys() if cat_map.get(r) == "DIVERS"]):
    p = product_stats[ref]
    descs = list(p["descs"])
    desc_str = descs[0] if descs else ''
    etats_str = ", ".join(sorted(p["etats"]))
    print(f"{ref:<28} {desc_str[:35]:<35} {etats_str[:25]:<25} {p['ca']/1e6:>10.2f} {p['lignes']:>8}")
    total_ca_divers += p["ca"]
print("-" * 110)
print(f"{'TOTAL DIVERS':<28} {'':<35} {'':<25} {total_ca_divers/1e6:>10.2f} M FCFA")

# === COMPARAISON AVANT/APRÈS ===
print("\n" + "=" * 80)
print("5. IMPACT DES CORRECTIONS (vs analyse précédente)")
print("=" * 80)
print(f"{'Catégorie':<28} {'Avant (t)':>12} {'Après (t)':>12} {'Écart (t)':>12} {'Avant (M)':>12} {'Après (M)':>12}")
print("-" * 90)

# Avant = valeurs précédemment utilisées dans le PDF
avant = {
    "ALIMENT COMPLET": (545.2, 470.6),
    "ALVEOLE": (0.0, 13.8),
    "COMPLEMENT ALIMENTAIRE": (3.0, 15.2),
    "CONCENTRES": (9075.4, 6064.4),
    "DIVERS": (2.1, 13.6),
    "INGREDIENTS": (4059.7, 1070.6),  # avant exclusion ligne CA=0
    "MATERIEL ELEVAGE": (10.3, 128.8),
    "PREMIX": (68.8, 126.0),
    "TOURTEAUX": (29061.5, 9437.4),
}

for cat in sorted(cat_stats.keys()):
    d = cat_stats[cat]
    vol_t = d["vol_kg"] / 1000.0
    ca_m = d["ca"] / 1e6
    av_vol, av_ca = avant.get(cat, (0, 0))
    diff_vol = vol_t - av_vol
    diff_ca = ca_m - av_ca
    print(f"{cat:<28} {av_vol:>12.1f} {vol_t:>12.1f} {diff_vol:>+12.1f} {av_ca:>12.1f} {ca_m:>12.1f}")

# Sauvegarder la configuration des services
with open('/home/z/my-project/scripts/services_validee.json', 'w', encoding='utf-8') as f:
    json.dump(list(SERVICES_VALIDEE), f, ensure_ascii=False, indent=2)
print(f"\n✓ Services avec état 'Validée' sauvegardés: {SERVICES_VALIDEE}")

# === IMPACT SUR INGREDIENTS — nouvelle liste ===
print("\n" + "=" * 100)
print("6. NOUVELLE LISTE INGREDIENTS AVEC CLÉS DE CONVERSION (CORRIGÉE)")
print("=" * 100)
print(f"{'Réf':<10} {'Description':<45} {'Clé':<10} {'Poids':>6} {'Sacs':>10} {'Tonnes':>10} {'CA (M)':>10} {'Prix (k/t)':>11}")
print("-" * 110)

ingredients = {ref: p for ref, p in product_stats.items() if cat_map.get(ref) == "INGREDIENTS"}
total_sacs = 0
total_t = 0
total_ca_i = 0
for ref in sorted(ingredients.keys(), key=lambda r: -ingredients[r]["vol_kg"]):
    p = ingredients[ref]
    descs = list(p["descs"])
    desc_str = descs[0] if descs else ''
    if len(descs) > 1:
        desc_str = f"{descs[0]} (+{len(descs)-1})"
    
    if ref in MANUAL_WEIGHTS:
        cle = "MANUEL"
        poids = MANUAL_WEIGHTS[ref]
    else:
        cle = "REGEX"
        weights = set()
        for d in descs:
            m = WEIGHT_RE.search(d)
            if m:
                v, u = m.group(1), m.group(2).upper()
                v = float(v)
                if u == "KG": weights.add(v)
                elif u in ("G", "GRAMME", "GRAMMES"): weights.add(v/1000)
                elif u == "L": weights.add(v)
        poids = max(weights) if weights else 0
    
    t = p["vol_kg"] / 1000
    ca_m = p["ca"] / 1e6
    prix = (p["ca"] / p["vol_kg"]) if p["vol_kg"] > 0 else 0
    print(f"{ref:<10} {desc_str[:45]:<45} {cle:<10} {poids:>6.2f} {p['lignes']:>10} {t:>10.2f} {ca_m:>10.2f} {prix:>11.0f}")
    total_sacs += p["lignes"]
    total_t += t
    total_ca_i += p["ca"]
print("-" * 110)
print(f"{'TOTAL':<10} {'':<45} {'':<10} {'':>6} {total_sacs:>10} {total_t:>10.2f} {total_ca_i/1e6:>10.2f}")
