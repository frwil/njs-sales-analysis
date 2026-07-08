import openpyxl
import json
import re
from collections import defaultdict

wb = openpyxl.load_workbook('/home/z/my-project/upload/ventes janv a juin 2026.xlsx', data_only=True)

# ============================================================
# 1. COMPARAISON CF101 vs CF1012 — PRIX
# ============================================================
print("=" * 100)
print("1. COMPARAISON CF101 vs CF1012 — PRIX UNITAIRE")
print("=" * 100)
print()

cf101_data = {"qte": 0, "ca": 0, "lignes": 0, "descs": set()}
cf1012_data = {"qte": 0, "ca": 0, "lignes": 0, "descs": set()}

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
        try:
            q = float(qte) if qte else 0
            c = float(ca) if ca else 0
        except:
            continue
        if ref_str == "CF101":
            cf101_data["qte"] += q
            cf101_data["ca"] += c
            cf101_data["lignes"] += 1
            cf101_data["descs"].add(str(desc) if desc else '')
        elif ref_str == "CF1012":
            cf1012_data["qte"] += q
            cf1012_data["ca"] += c
            cf1012_data["lignes"] += 1
            cf1012_data["descs"].add(str(desc) if desc else '')

print(f"CF101 (Carbonate de calcium):")
print(f"  Description: {cf101_data['descs']}")
print(f"  Quantité totale: {cf101_data['qte']:.0f} sacs")
print(f"  CA HT total: {cf101_data['ca']:,.0f} FCFA")
print(f"  Prix unitaire moyen: {cf101_data['ca']/cf101_data['qte']:.0f} FCFA/sac" if cf101_data['qte'] > 0 else "  N/A")
print(f"  Lignes: {cf101_data['lignes']}")

print(f"\nCF1012 (Carbonate de calcium 50 KG):")
print(f"  Description: {cf1012_data['descs']}")
print(f"  Quantité totale: {cf1012_data['qte']:.0f} sacs")
print(f"  CA HT total: {cf1012_data['ca']:,.0f} FCFA")
print(f"  Prix unitaire moyen: {cf1012_data['ca']/cf1012_data['qte']:.0f} FCFA/sac" if cf1012_data['qte'] > 0 else "  N/A")
print(f"  Lignes: {cf1012_data['lignes']}")

if cf101_data['qte'] > 0 and cf1012_data['qte'] > 0:
    p_cf101 = cf101_data['ca']/cf101_data['qte']
    p_cf1012 = cf1012_data['ca']/cf1012_data['qte']
    ratio = p_cf101 / p_cf1012
    print(f"\nRatio CF101 / CF1012: {ratio:.3f}")
    print(f"Si CF1012 = 50 kg → CF101 = {50*ratio:.1f} kg")
    if abs(ratio - 0.02) < 0.01:
        print(f"✓ Le ratio correspond à 1kg vs 50kg (ratio = 0,02) → CF101 est bien en 1 Kg")
        poids_cf101 = 1.0
    elif abs(ratio - 0.5) < 0.05:
        print(f"✓ Le ratio correspond à 25kg vs 50kg → CF101 est en 25 Kg")
        poids_cf101 = 25.0
    elif abs(ratio - 1) < 0.05:
        print(f"⚠️ Prix identiques — pourrait être un même conditionnement")
        poids_cf101 = 1.0
    else:
        print(f"Ratio atypique — utiliser 1 kg par défaut")
        poids_cf101 = 1.0
else:
    poids_cf101 = 1.0
    print(f"\n→ Utilisation de 1 kg par défaut pour CF101")

# ============================================================
# 2. MAÏS — LIGNES À PRIX 0 (régularisation stock)
# ============================================================
print("\n" + "=" * 100)
print("2. MAÏS (M1051) — LIGNES À PRIX 0 (RÉGULARISATION STOCK)")
print("=" * 100)
print()

mais_data = {
    "total": {"qte": 0, "ca": 0, "lignes": 0},
    "prix_zero": {"qte": 0, "ca": 0, "lignes": 0},
    "prix_normal": {"qte": 0, "ca": 0, "lignes": 0},
    "exemples_prix_zero": []
}

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    if ws.cell(2, 1).value != 'Réf. produit':
        continue
    for row in range(3, ws.max_row + 1):
        ref = ws.cell(row, 1).value
        qte = ws.cell(row, 3).value
        ca = ws.cell(row, 9).value
        etat = ws.cell(row, 16).value
        tiers = ws.cell(row, 6).value
        if not ref or etat != 'Livrée':
            continue
        ref_str = str(ref).strip()
        if ref_str != "M1051":
            continue
        try:
            q = float(qte) if qte else 0
            c = float(ca) if ca else 0
        except:
            continue
        mais_data["total"]["qte"] += q
        mais_data["total"]["ca"] += c
        mais_data["total"]["lignes"] += 1
        if c == 0:
            mais_data["prix_zero"]["qte"] += q
            mais_data["prix_zero"]["ca"] += c
            mais_data["prix_zero"]["lignes"] += 1
            if len(mais_data["exemples_prix_zero"]) < 5:
                mais_data["exemples_prix_zero"].append({
                    "tiers": str(tiers)[:40] if tiers else "",
                    "qte": q,
                    "ca": c
                })
        else:
            mais_data["prix_normal"]["qte"] += q
            mais_data["prix_normal"]["ca"] += c
            mais_data["prix_normal"]["lignes"] += 1

print(f"MAÏS (M1051) — Total S1 2026:")
print(f"  Toutes lignes: {mais_data['total']['lignes']} | Qté: {mais_data['total']['qte']:.0f} sacs | CA: {mais_data['total']['ca']:,.0f} FCFA")
print(f"  → Volume total: {mais_data['total']['qte']*50/1000:.1f} t")
print()
print(f"Lignes à PRIX 0 (régularisation stock — À RETIRER):")
print(f"  Lignes: {mais_data['prix_zero']['lignes']}")
print(f"  Qté: {mais_data['prix_zero']['qte']:.0f} sacs")
print(f"  Volume: {mais_data['prix_zero']['qte']*50/1000:.1f} t")
print(f"  CA: {mais_data['prix_zero']['ca']:,.0f} FCFA")
if mais_data["exemples_prix_zero"]:
    print(f"\n  Exemples:")
    for ex in mais_data["exemples_prix_zero"]:
        print(f"    - {ex['tiers']}: {ex['qte']:.0f} sacs, CA={ex['ca']:.0f} FCFA")

print()
print(f"Lignes à PRIX NORMAL (à conserver):")
print(f"  Lignes: {mais_data['prix_normal']['lignes']}")
print(f"  Qté: {mais_data['prix_normal']['qte']:.0f} sacs")
print(f"  Volume: {mais_data['prix_normal']['qte']*50/1000:.1f} t")
print(f"  CA: {mais_data['prix_normal']['ca']:,.0f} FCFA")
if mais_data["prix_normal"]["qte"] > 0:
    print(f"  Prix moyen: {mais_data['prix_normal']['ca']/mais_data['prix_normal']['qte']/50:.0f} FCFA/kg = {mais_data['prix_normal']['ca']/mais_data['prix_normal']['qte']*20:.0f} FCFA/t")

print(f"\n→ Impact de la correction:")
print(f"  Avant: {mais_data['total']['qte']*50/1000:.1f} t")
print(f"  Après: {mais_data['prix_normal']['qte']*50/1000:.1f} t")
print(f"  Différence: -{(mais_data['total']['qte']-mais_data['prix_normal']['qte'])*50/1000:.1f} t")

# ============================================================
# 3. CONTRIBUTION_CARBURANT ET PONT_BASCULE
# ============================================================
print("\n" + "=" * 100)
print("3. CONTRIBUTION_CARBURANT ET PONT_BASCULE — INTÉGRATION AU CA GLOBAL")
print("=" * 100)
print()

# Chercher toutes les refs qui ressemblent à des frais
frais_keywords = ["CONTRIBUTION", "CARBURANT", "PONT", "BASCULE", "S0020", "MA100", "MA101", "MA90"]
frais_data = defaultdict(lambda: {"qte": 0, "ca": 0, "lignes": 0, "descs": set()})

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
        desc_str = str(desc) if desc else ""
        # Match par ref ou par description
        match = False
        for kw in frais_keywords:
            if kw in ref_str.upper() or kw in desc_str.upper():
                match = True
                break
        if match:
            try:
                q = float(qte) if qte else 0
                c = float(ca) if ca else 0
            except:
                continue
            frais_data[ref_str]["qte"] += q
            frais_data[ref_str]["ca"] += c
            frais_data[ref_str]["lignes"] += 1
            frais_data[ref_str]["descs"].add(desc_str[:60])

# Vérifier ce qui est dans le mapping actuel
with open('/home/z/my-project/scripts/product_category_map.json') as f:
    cat_map = json.load(f)

print(f"{'Réf':<30} {'Description':<50} {'Catégorie':<22} {'CA (M)':>10} {'Lignes':>8}")
print("-" * 130)
total_ca_frais = 0
for ref in sorted(frais_data.keys()):
    p = frais_data[ref]
    cat = cat_map.get(ref, "❌ ABSENT")
    descs = list(p["descs"])
    desc_str = descs[0] if descs else ""
    if len(descs) > 1:
        desc_str = f"{descs[0]} (+{len(descs)-1})"
    print(f"{ref:<30} {desc_str[:50]:<50} {cat:<22} {p['ca']/1e6:>10.2f} {p['lignes']:>8}")
    total_ca_frais += p["ca"]

print("-" * 130)
print(f"{'TOTAL':<30} {'':<50} {'':<22} {total_ca_frais/1e6:>10.2f} M FCFA")

# Vérifier PONT_BASCULE spécifiquement
print(f"\nRecherche spécifique 'PONT_BASCULE' ou 'PONT BASCULE':")
pont_bascule_trouve = False
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    if ws.cell(2, 1).value != 'Réf. produit':
        continue
    for row in range(3, ws.max_row + 1):
        ref = ws.cell(row, 1).value
        desc = ws.cell(row, 2).value
        if not ref and not desc:
            continue
        ref_str = str(ref).strip() if ref else ""
        desc_str = str(desc) if desc else ""
        if "PONT" in ref_str.upper() or "BASCULE" in ref_str.upper() or "PONT" in desc_str.upper() or "BASCULE" in desc_str.upper():
            print(f"  Feuille {sheet_name} - Réf: {ref_str}, Desc: {desc_str[:60]}")
            pont_bascule_trouve = True
if not pont_bascule_trouve:
    print("  ❌ Aucune ligne avec PONT_BASCULE ou PONT BASCULE trouvée dans les ventes S1 2026")

# Vérifier le CA global et la part de frais
print("\n" + "=" * 100)
print("IMPACT SUR LE CA GLOBAL")
print("=" * 100)

# Charger ventes totales
total_ca_all = 0
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    if ws.cell(2, 1).value != 'Réf. produit':
        continue
    for row in range(3, ws.max_row + 1):
        ca = ws.cell(row, 9).value
        etat = ws.cell(row, 16).value
        if etat != 'Livrée':
            continue
        try:
            c = float(ca) if ca else 0
        except:
            continue
        total_ca_all += c

print(f"CA HT total S1 2026 (toutes refs, tous clients): {total_ca_all/1e6:,.1f} M FCFA")
print(f"CA des frais (CONTRIBUTION_CARBURANT + autres DIVERS): {total_ca_frais/1e6:,.1f} M FCFA")
print(f"Part des frais dans le CA global: {total_ca_frais/total_ca_all*100:.2f}%")
