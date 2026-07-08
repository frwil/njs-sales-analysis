"""
Annexe 20/80 par type de concentrés (Chair, Porc, Ponte) avec objectifs S2.
Attribution des objectifs S2 basée sur le poids du client dans la consommation globale du type.
"""
import openpyxl
import json
import re
from collections import defaultdict

# Charger mapping
with open('/home/z/my-project/scripts/product_category_map.json') as f:
    cat_map = json.load(f)

# Charger clients exclus (comptoirs, etc.)
with open('/home/z/my-project/scripts/excluded_clients.json') as f:
    excluded_tiers = set(json.load(f))

# Mapping produit → type de concentré
CONCENTRE_TYPE_MAP = {
    # Chair 10%
    "C104": "Chair", "C1042": "Chair", "C1043": "Chair", "C1044": "Chair",
    # Chair 5%
    "C105": "Chair", "C1053": "Chair", "C1054": "Chair", "C1055": "Chair",
    # Ponte
    "C102": "Ponte", "C1022": "Ponte",
    "C101": "Ponte",
    # Porc
    "C103": "Porc",
    # Rabbit (mais c'est maintenant ALIMENT COMPLET, on l'exclut)
    # "ALAP25": "Rabbit",
}

# Poids manuels
MANUAL_WEIGHTS = {"M1051": 50.0, "CF101": 1.0, "S101": 1.0}
SERVICES_VALIDEE = {"PONT_BASCULE", "CONTRIBUTION_CARBURANT"}

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

def normalize_client(tiers_str):
    """Extraire code + nom du client."""
    if not tiers_str: return ("", "")
    s = str(tiers_str).strip()
    # Format: "CU2601-14325 -  TCHONBEUA AUGUSTIN"
    parts = s.split(" - ", 1)
    if len(parts) == 2:
        code = parts[0].strip()
        nom = parts[1].strip()
        return (code, nom)
    return (s, s)

# === LIRE VENTES S1 2026 (Livrée + services Validée) ===
wb = openpyxl.load_workbook('/home/z/my-project/upload/ventes janv a juin 2026.xlsx', read_only=True, data_only=True)

# Stats par client et par type de concentré
client_type_stats = defaultdict(lambda: defaultdict(lambda: {"vol_kg": 0, "ca": 0, "qte_sacs": 0}))
# Stats globales par type (tous clients)
type_global = defaultdict(lambda: {"vol_kg": 0, "ca": 0, "qte_sacs": 0, "clients": set()})
# Stats par client (toutes catégories pour identification 20/80)
client_global = defaultdict(lambda: {"ca": 0, "nom": ""})

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
        tiers = row[5]
        if not ref: continue
        ref_str = str(ref).strip()
        etat_str = str(etat) if etat else ""
        
        # Filtre: Livrée ou services Validée
        include = False
        if etat_str == "Livrée":
            include = True
        elif ref_str in SERVICES_VALIDEE and etat_str == "Validée":
            include = True
        if not include: continue
        
        # Exclure M1051 CA=0
        if ref_str == "M1051" and (ca is None or float(ca) == 0):
            continue
        
        # Exclure clients comptoirs
        if tiers in excluded_tiers:
            continue
        
        # Si c'est un concentré, classifier par type
        if ref_str in CONCENTRE_TYPE_MAP:
            type_conc = CONCENTRE_TYPE_MAP[ref_str]
            try:
                q = float(qte) if qte else 0
                c = float(ca) if ca else 0
            except:
                continue
            weight = parse_weight_kg(ref_str, desc)
            vol_kg = q * weight
            
            code, nom = normalize_client(tiers)
            client_key = code if code else nom
            
            client_type_stats[client_key][type_conc]["vol_kg"] += vol_kg
            client_type_stats[client_key][type_conc]["ca"] += c
            client_type_stats[client_key][type_conc]["qte_sacs"] += q
            client_type_stats[client_key][type_conc]["nom"] = nom
            
            type_global[type_conc]["vol_kg"] += vol_kg
            type_global[type_conc]["ca"] += c
            type_global[type_conc]["qte_sacs"] += q
            type_global[type_conc]["clients"].add(client_key)
        
        # Stats globales par client
        try:
            c = float(ca) if ca else 0
        except:
            continue
        code, nom = normalize_client(tiers)
        client_key = code if code else nom
        if client_key:
            client_global[client_key]["ca"] += c
            client_global[client_key]["nom"] = nom

wb.close()

# === AFFICHAGE GLOBAL PAR TYPE ===
print("=" * 110)
print("STATISTIQUES GLOBALES PAR TYPE DE CONCENTRÉ (S1 2026, clients hors comptoirs)")
print("=" * 110)
print(f"{'Type':<10} {'Clients':>10} {'Volume (t)':>14} {'CA (M FCFA)':>14} {'Prix moyen (k/t)':>18}")
print("-" * 70)

obj_s2_by_type = {}  # objectifs S2 par type (basés sur obj global CONCENTRES 11 650t réparti)
# Répartition S2 2025: Chair ~75%, Ponte ~20%, Porc ~5%
# S2 2026 recalibré total CONCENTRES = 11 650 t
# Appliquons la même répartition que S1 2026
type_share = {}
total_vol_all = sum(t["vol_kg"] for t in type_global.values())
for t in ["Chair", "Ponte", "Porc"]:
    if t in type_global:
        share = type_global[t]["vol_kg"] / total_vol_all if total_vol_all > 0 else 0
        type_share[t] = share
        obj_s2_by_type[t] = 11650 * share  # tonnes

for t in ["Chair", "Ponte", "Porc"]:
    if t not in type_global: continue
    d = type_global[t]
    vol_t = d["vol_kg"] / 1000
    ca_m = d["ca"] / 1e6
    prix = d["ca"]/d["vol_kg"] if d["vol_kg"] > 0 else 0
    print(f"{t:<10} {len(d['clients']):>10} {vol_t:>14.1f} {ca_m:>14.2f} {prix/1000:>18.0f}")
print("-" * 70)
total_clients = len(set().union(*[t["clients"] for t in type_global.values()]))
total_vol = sum(t["vol_kg"] for t in type_global.values()) / 1000
total_ca = sum(t["ca"] for t in type_global.values()) / 1e6
print(f"{'TOTAL':<10} {total_clients:>10} {total_vol:>14.1f} {total_ca:>14.2f}")

print(f"\nObjectifs S2 2026 par type (répartition proportionnelle de l'obj. CONCENTRES = 11 650 t):")
for t in ["Chair", "Ponte", "Porc"]:
    print(f"  {t}: {obj_s2_by_type[t]:.0f} t (share {type_share[t]*100:.1f}%)")

# === ANNEXE 20/80 PAR TYPE ===
print("\n" + "=" * 110)
print("ANNEXE 20/80 PAR TYPE DE CONCENTRÉ")
print("=" * 110)

annexe_data = {}  # pour sauvegarde JSON

for type_conc in ["Chair", "Ponte", "Porc"]:
    # Récupérer clients de ce type, triés par volume décroissant
    clients_type = []
    for client_key, types in client_type_stats.items():
        if type_conc in types:
            d = types[type_conc]
            nom = d.get("nom", "")
            clients_type.append({
                "code": client_key,
                "nom": nom,
                "vol_t": d["vol_kg"]/1000,
                "ca_m": d["ca"]/1e6,
                "sacs": d["qte_sacs"]
            })
    clients_type.sort(key=lambda x: -x["vol_t"])
    
    total_vol_type = sum(c["vol_t"] for c in clients_type)
    total_ca_type = sum(c["ca_m"] for c in clients_type)
    
    # Pareto 20/80 — calculer le top 20% des clients qui font 80% du volume
    n_clients = len(clients_type)
    top_20_pct_count = max(1, int(n_clients * 0.2))
    
    # Trouver le cutoff à 80% du volume
    cumul = 0
    cutoff_idx = n_clients
    for i, c in enumerate(clients_type):
        cumul += c["vol_t"]
        if cumul >= 0.8 * total_vol_type:
            cutoff_idx = i + 1
            break
    
    top_20_80 = clients_type[:cutoff_idx]
    
    print(f"\n--- {type_conc.upper()} — Top {len(top_20_80)} clients 20/80 (sur {n_clients} total) ---")
    print(f"Volume total {type_conc}: {total_vol_type:.1f} t | CA: {total_ca_type:.1f} M FCFA | Prix moyen: {total_ca_type*1e6/(total_vol_type*1000):.0f} FCFA/t")
    print(f"Objectif S2 {type_conc}: {obj_s2_by_type[type_conc]:.0f} t (vs S1={total_vol_type:.0f} t, croissance nécessaire: +{(obj_s2_by_type[type_conc]/total_vol_type-1)*100:.1f}%)")
    print()
    print(f"{'#':<4} {'Code':<18} {'Nom':<35} {'Vol S1 (t)':>12} {'CA (M)':>10} {'% vol':>8} {'Obj S2 (t)':>12} {'CA obj (M)':>12}")
    print("-" * 115)
    
    annexe_data[type_conc] = {
        "total_clients": n_clients,
        "total_vol_t": total_vol_type,
        "total_ca_m": total_ca_type,
        "obj_s2_t": obj_s2_by_type[type_conc],
        "top_20_80_count": len(top_20_80),
        "clients": []
    }
    
    obj_total_t = obj_s2_by_type[type_conc]
    
    for i, c in enumerate(top_20_80, 1):
        pct = c["vol_t"] / total_vol_type * 100 if total_vol_type > 0 else 0
        # Objectif S2 = même poids relatif × objectif global du type
        obj_t = c["vol_t"] / total_vol_type * obj_total_t if total_vol_type > 0 else 0
        # CA objectif = obj volume × prix moyen actuel
        prix_moyen = c["ca_m"] / c["vol_t"] * 1000 if c["vol_t"] > 0 else 0  # k/t
        obj_ca = obj_t * prix_moyen / 1000  # M FCFA
        print(f"{i:<4} {c['code'][:17]:<18} {c['nom'][:34]:<35} {c['vol_t']:>12.2f} {c['ca_m']:>10.2f} {pct:>7.1f}% {obj_t:>12.1f} {obj_ca:>12.1f}")
        
        annexe_data[type_conc]["clients"].append({
            "rang": i,
            "code": c["code"],
            "nom": c["nom"],
            "vol_s1_t": round(c["vol_t"], 2),
            "ca_s1_m": round(c["ca_m"], 2),
            "pct_vol": round(pct, 1),
            "obj_s2_t": round(obj_t, 1),
            "obj_s2_ca_m": round(obj_ca, 1),
            "prix_moyen_k_t": round(prix_moyen, 0)
        })
    
    top_vol = sum(c["vol_t"] for c in top_20_80)
    top_ca = sum(c["ca_m"] for c in top_20_80)
    top_obj = sum(c["vol_t"] / total_vol_type * obj_total_t for c in top_20_80)
    print("-" * 115)
    print(f"{'':<4} {'SOUS-TOTAL 20/80':<18} {'':<35} {top_vol:>12.1f} {top_ca:>10.1f} {top_vol/total_vol_type*100:>7.1f}% {top_obj:>12.1f}")
    print(f"{'':<4} {'RESTE (80/20)':<18} {'':<35} {total_vol_type-top_vol:>12.1f} {total_ca_type-top_ca:>10.1f} {(total_vol_type-top_vol)/total_vol_type*100:>7.1f}% {obj_total_t-top_obj:>12.1f}")
    print(f"{'':<4} {'TOTAL':<18} {'':<35} {total_vol_type:>12.1f} {total_ca_type:>10.1f} {'100,0%':>8} {obj_total_t:>12.1f}")

# Sauvegarder
with open('/home/z/my-project/scripts/annexe_20_80_par_type.json', 'w', encoding='utf-8') as f:
    json.dump(annexe_data, f, ensure_ascii=False, indent=2)
print(f"\n✓ Annexe sauvegardée dans annexe_20_80_par_type.json")

# === STATS POUR OBJECTIFS INDIVIDUELS ===
print("\n" + "=" * 110)
print("SYNTHÈSE — OBJECTIFS S2 PAR TYPE DE CONCENTRÉ")
print("=" * 110)
print(f"{'Type':<10} {'Clients S1':>10} {'20/80':>8} {'Vol S1 (t)':>12} {'Obj S2 (t)':>12} {'Croissance':>12} {'CA S1 (M)':>12} {'CA obj S2 (M)':>14}")
print("-" * 92)
for t in ["Chair", "Ponte", "Porc"]:
    if t not in annexe_data: continue
    d = annexe_data[t]
    croissance = (d["obj_s2_t"]/d["total_vol_t"]-1)*100 if d["total_vol_t"] > 0 else 0
    # CA objectif = obj volume × prix moyen S1
    prix_moyen = d["total_ca_m"] / d["total_vol_t"] * 1000 if d["total_vol_t"] > 0 else 0  # k/t
    ca_obj = d["obj_s2_t"] * prix_moyen / 1000
    print(f"{t:<10} {d['total_clients']:>10} {d['top_20_80_count']:>8} {d['total_vol_t']:>12.1f} {d['obj_s2_t']:>12.1f} {croissance:>+11.1f}% {d['total_ca_m']:>12.1f} {ca_obj:>14.1f}")
print("-" * 92)
total_vol_s1 = sum(annexe_data[t]["total_vol_t"] for t in annexe_data)
total_obj_s2 = sum(annexe_data[t]["obj_s2_t"] for t in annexe_data)
total_ca_s1 = sum(annexe_data[t]["total_ca_m"] for t in annexe_data)
prix_moyen_global = total_ca_s1 / total_vol_s1 * 1000
total_ca_obj = total_obj_s2 * prix_moyen_global / 1000
print(f"{'TOTAL':<10} {sum(annexe_data[t]['total_clients'] for t in annexe_data):>10} {sum(annexe_data[t]['top_20_80_count'] for t in annexe_data):>8} {total_vol_s1:>12.1f} {total_obj_s2:>12.1f} {(total_obj_s2/total_vol_s1-1)*100:>+11.1f}% {total_ca_s1:>12.1f} {total_ca_obj:>14.1f}")
