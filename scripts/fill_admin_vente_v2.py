"""
Version corrigée : applique le MÊME mapping cat_map aux deux années (2025 et 2026)
pour assurer la comparabilité.
"""
import openpyxl
import json
import re
from collections import defaultdict
from datetime import datetime

with open('/home/z/my-project/scripts/product_category_map.json') as f:
    cat_map = json.load(f)

MANUAL_WEIGHTS = {"M1051": 50.0, "CF101": 1.0, "S101": 1.0, "P109": 25.0}
SERVICES_VALIDEE = {"PONT_BASCULE", "CONTRIBUTION_CARBURANT"}
WEIGHT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(KG|Kg|kg|KG|GRAMMES?|G|L)\b", re.IGNORECASE)

def parse_weight_kg(ref, desc):
    ref_str = str(ref).strip() if ref else ""
    if ref_str in MANUAL_WEIGHTS: return MANUAL_WEIGHTS[ref_str]
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

def normalize_agence(agence_str):
    if not agence_str: return ""
    s = str(agence_str).strip()
    if s.upper().startswith("SPC") or s.upper().startswith("PDC"): return ""
    if "AGENCE " in s.upper(): s = s[7:]
    if s.upper().startswith("DE "): s = s[3:]
    if "BAMENDA" in s.upper() and "MBOUDA" in s.upper(): return "Mbouda"
    if "BERTOUA" in s.upper(): return "Bertoua"
    if s: s = s[0].upper() + s[1:].lower()
    return s

def extract_data(wb_path, year, has_native_cat=False):
    """
    Extrait CA et volume par agence, catégorie, produit.
    - has_native_cat=True: utilise la catégorie native du fichier (2025)
    - has_native_cat=False: utilise cat_map.json (2026)
    Pour comparaison cohérente, on utilise TOUJOURS cat_map.
    """
    wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
    
    agence_ca = defaultdict(float)
    cat_ca = defaultdict(float)
    cat_vol = defaultdict(float)
    prod_ca = defaultdict(float)
    prod_vol = defaultdict(float)
    
    if year == 2025:
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 14: continue
            ref = row[0]
            desc = row[1]
            date_cmd = row[5]
            ca = row[6]
            etat = row[9]
            agence = row[12]
            
            if not ref: continue
            if str(etat) != "Livrée": continue
            
            # Filtre S1 (Jan-Jun)
            month = None
            if date_cmd:
                if hasattr(date_cmd, 'month'): month = date_cmd.month
                else:
                    try:
                        s = str(date_cmd)
                        parts = s.split("/")
                        if len(parts) == 3: month = int(parts[1])
                        elif len(s) >= 7: month = int(s[5:7])
                    except: pass
            if month is None or month > 6: continue
            
            try:
                c = float(ca) if ca else 0
            except: continue
            
            ref_str = str(ref).strip()
            # Appliquer cat_map (pas la catégorie native du fichier)
            cat = cat_map.get(ref_str, "DIVERS")
            
            # Pour 2025, le volume est déjà en tonnes (col 11)
            # Mais pour Maïs (M1051), le fichier a vol=0 — on recalcule avec qte×50kg
            qte = row[2] if len(row) > 2 else 0
            vol_t = 0
            if ref_str == "M1051":
                try:
                    q = float(qte) if qte else 0
                    vol_t = q * 50 / 1000
                except: vol_t = 0
            else:
                vol_t_native = row[11] if len(row) > 11 else 0
                try:
                    vol_t = float(vol_t_native) if vol_t_native else 0
                except: vol_t = 0
            
            agence_norm = normalize_agence(agence)
            if agence_norm:
                agence_ca[agence_norm] += c
            
            cat_ca[cat] += c
            cat_vol[cat] += vol_t
            prod_ca[ref_str] += c
            prod_vol[ref_str] += vol_t
    else:  # 2026
        sheet_to_month = {"Sheet 1": 1, "Feuil1": 2, "Feuil2": 3, "Feuil3": 4, "Feuil4": 5, "Feuil5": 6}
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            if ws.cell(2, 1).value != 'Réf. produit': continue
            for row in ws.iter_rows(min_row=3, values_only=True):
                if not row or len(row) < 16: continue
                ref = row[0]
                desc = row[1]
                qte = row[2]
                ca = row[8]
                etat = row[15]
                agence = row[17] if len(row) > 17 else None
                
                if not ref: continue
                ref_str = str(ref).strip()
                etat_str = str(etat) if etat else ""
                
                if etat_str != "Livrée":
                    if not (ref_str in SERVICES_VALIDEE and etat_str == "Validée"):
                        continue
                
                try:
                    q = float(qte) if qte else 0
                    c = float(ca) if ca else 0
                except: continue
                
                if ref_str == "M1051" and c == 0: continue
                
                weight = parse_weight_kg(ref_str, desc)
                vol_t = q * weight / 1000.0
                
                cat = cat_map.get(ref_str, "DIVERS")
                agence_norm = normalize_agence(agence)
                
                if agence_norm:
                    agence_ca[agence_norm] += c
                
                cat_ca[cat] += c
                cat_vol[cat] += vol_t
                prod_ca[ref_str] += c
                prod_vol[ref_str] += vol_t
    
    wb.close()
    return dict(agence_ca), dict(cat_ca), dict(cat_vol), dict(prod_ca), dict(prod_vol)

# Extraction
print("Lecture 2025 S1...")
ag25, cat_ca25, cat_vol25, prod_ca25, prod_vol25 = extract_data(
    '/home/z/my-project/upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx', 2025)
print(f"  CA total: {sum(ag25.values())/1e6:.1f} M FCFA")
print(f"  Vol total: {sum(cat_vol25.values()):.1f} t")

print("\nLecture 2026 S1...")
ag26, cat_ca26, cat_vol26, prod_ca26, prod_vol26 = extract_data(
    '/home/z/my-project/upload/ventes janv a juin 2026.xlsx', 2026)
print(f"  CA total: {sum(ag26.values())/1e6:.1f} M FCFA")
print(f"  Vol total: {sum(cat_vol26.values()):.1f} t")

# Vérification par catégorie
print("\n=== CATÉGORIES (mapping unifié) ===")
print(f"{'Catégorie':<28} {'CA 2025 (M)':>12} {'CA 2026 (M)':>12} {'Vol 2025':>10} {'Vol 2026':>10}")
print("-" * 75)
for cat in sorted(set(list(cat_ca25.keys()) + list(cat_ca26.keys()))):
    print(f"{cat:<28} {cat_ca25.get(cat,0)/1e6:>12.1f} {cat_ca26.get(cat,0)/1e6:>12.1f} {cat_vol25.get(cat,0):>10.1f} {cat_vol26.get(cat,0):>10.1f}")

# ===== AGRÉGATION SELON LES RUBRIQUES DEMANDÉES =====
def get_rubrique_data(cat_ca, cat_vol, prod_ca, prod_vol):
    r = {}
    r['TOURTEAU SOJA'] = {'ca': cat_ca.get('TOURTEAUX', 0), 'vol': cat_vol.get('TOURTEAUX', 0)}
    r['INGREDIENTS'] = {
        'ca': cat_ca.get('INGREDIENTS', 0) - prod_ca.get('M1051', 0),
        'vol': cat_vol.get('INGREDIENTS', 0) - prod_vol.get('M1051', 0)
    }
    r['MAIS'] = {'ca': prod_ca.get('M1051', 0), 'vol': prod_vol.get('M1051', 0)}
    r['MATERIEL ELEVAGE'] = {'ca': cat_ca.get('MATERIEL ELEVAGE', 0), 'vol': cat_vol.get('MATERIEL ELEVAGE', 0)}
    r['ALIMENTS COMPLETS'] = {'ca': cat_ca.get('ALIMENT COMPLET', 0), 'vol': cat_vol.get('ALIMENT COMPLET', 0)}
    r['CONCENTRES'] = {'ca': cat_ca.get('CONCENTRES', 0), 'vol': cat_vol.get('CONCENTRES', 0)}
    # VENTES DIVERSES = DIVERS strict (hors PONT_BASCULE, CONTRIBUTION_CARBURANT) + COMPLEMENT ALIMENTAIRE
    divers_strict_ca = cat_ca.get('DIVERS', 0) - prod_ca.get('PONT_BASCULE', 0) - prod_ca.get('CONTRIBUTION_CARBURANT', 0)
    divers_strict_vol = cat_vol.get('DIVERS', 0) - prod_vol.get('PONT_BASCULE', 0) - prod_vol.get('CONTRIBUTION_CARBURANT', 0)
    r['DIVERSES'] = {
        'ca': divers_strict_ca + cat_ca.get('COMPLEMENT ALIMENTAIRE', 0),
        'vol': divers_strict_vol + cat_vol.get('COMPLEMENT ALIMENTAIRE', 0)
    }
    # PRODUITS ACCESSOIRES = PONT_BASCULE + CONTRIBUTION_CARBURANT
    r['PRODUITS ACCESSOIRES'] = {
        'ca': prod_ca.get('PONT_BASCULE', 0) + prod_ca.get('CONTRIBUTION_CARBURANT', 0),
        'vol': prod_vol.get('PONT_BASCULE', 0) + prod_vol.get('CONTRIBUTION_CARBURANT', 0)
    }
    r['ALVEOLES'] = {'ca': cat_ca.get('ALVEOLE', 0), 'vol': cat_vol.get('ALVEOLE', 0)}
    return r

rubriques_2025 = get_rubrique_data(cat_ca25, cat_vol25, prod_ca25, prod_vol25)
rubriques_2026 = get_rubrique_data(cat_ca26, cat_vol26, prod_ca26, prod_vol26)

print("\n=== RUBRIQUES FINALES (S1 2025 vs S1 2026) ===")
print(f"{'Rubrique':<35} {'CA 2025 (M)':>12} {'CA 2026 (M)':>12} {'Vol 2025 (t)':>14} {'Vol 2026 (t)':>14}")
print("-" * 90)
total_ca_25 = total_ca_26 = total_vol_25 = total_vol_26 = 0
for r in ['TOURTEAU SOJA', 'INGREDIENTS', 'MAIS', 'MATERIEL ELEVAGE', 'ALIMENTS COMPLETS', 'CONCENTRES', 'DIVERSES', 'PRODUITS ACCESSOIRES', 'ALVEOLES']:
    ca25 = rubriques_2025[r]['ca']
    ca26 = rubriques_2026[r]['ca']
    v25 = rubriques_2025[r]['vol']
    v26 = rubriques_2026[r]['vol']
    print(f"{r:<35} {ca25/1e6:>12.2f} {ca26/1e6:>12.2f} {v25:>14.2f} {v26:>14.2f}")
    total_ca_25 += ca25
    total_ca_26 += ca26
    total_vol_25 += v25
    total_vol_26 += v26
print("-" * 90)
print(f"{'TOTAL':<35} {total_ca_25/1e6:>12.2f} {total_ca_26/1e6:>12.2f} {total_vol_25:>14.2f} {total_vol_26:>14.2f}")

# Sauvegarder
data_to_save = {
    'rubriques_2025': rubriques_2025,
    'rubriques_2026': rubriques_2026,
    'agence_ca_2025': ag25,
    'agence_ca_2026': ag26,
    'agences_ordre': [
        {'label': 'YAOUNDE-MESSASI', 'agence': 'Messassi'},
        {'label': 'YAOUNDE-AHALA', 'agence': 'Ahala'},
        {'label': 'YAOUNDE-NKOABANG', 'agence': 'Nkoabang'},
        {'label': 'YAOUNDE-NKOLBISSON', 'agence': 'Nkolbisson'},
        {'label': 'BERTOUA', 'agence': 'Bertoua'},
        {'label': 'NGAOUNDERE', 'agence': 'Ngaoundere'},
        {'label': 'BAFOUSSAM-FAMLA', 'agence': 'Famla'},
        {'label': 'BAFOUSSAM-DJELENG', 'agence': 'Djeleng'},
        {'label': 'BAMENDA', 'agence': 'Mbouda'},
        {'label': 'NKONGSAMBA', 'agence': 'Nkongsamba'},
        {'label': 'BUEA', 'agence': 'Buea'},
        {'label': 'DOUALA-AXE LOURD', 'agence': 'Village'},
        {'label': 'DOUALA-SIEGE', 'agence': 'Ndobo'},
        {'label': 'DOUALA-PK11', 'agence': 'Pk11'},
    ],
}
with open('/home/z/my-project/scripts/admin_vente_data.json', 'w', encoding='utf-8') as f:
    json.dump(data_to_save, f, ensure_ascii=False, indent=2, default=str)
print("\n✓ Données sauvegardées")
