"""
Remplit le fichier ADMINISTRATEUR DE VENTE.xlsx avec les données S1 2026 et S1 2025.

Mapping rubriques:
- VENTES TOURTEAU SOJA = TOURTEAUX
- VENTES INGREDIENTS = INGREDIENTS - M1051 (Maïs)
- MAIS = M1051
- VENTES MATERIEL ELEVAGE = MATERIEL ELEVAGE
- ALIMENTS POISSON & CHICK & PIGLET BOOSTER = ALIMENT COMPLET
- VENTES CONCENTRES = CONCENTRES
- VENTES DIVERSES = DIVERS strict (Manuels, Sacs, Pierre à lécher) + COMPLEMENT ALIMENTAIRE (liquides)
- PRODUITS ACCESSOIRES = PONT_BASCULE + CONTRIBUTION_CARBURANT (+ autres services)
- ALVEOLES = ALVEOLE

Mapping agences (depuis fichier, R81-R94):
- YAOUNDE-MESSASI → Messassi
- YAOUNDE-AHALA → Ahala
- YAOUNDE-NKOABANG → Nkoabang
- YAOUNDE-NKOLBISSON → Nkolbisson
- BERTOUA → Bertoua
- NGAOUNDERE → Ngaoundere
- BAFOUSSAM-FAMLA → Famla
- BAFOUSSAM-DJELENG → Djeleng
- BAMENDA → Mbouda
- NKONGSAMBA → Nkongsamba
- BUEA → Buea
- DOUALA-AXE LOURD → Village
- DOUALA-SIEGE → Ndobo
- DOUALA-PK11 → Pk11
"""
import openpyxl
import json
import re
from collections import defaultdict
from datetime import datetime
from copy import copy

# Charger mapping
with open('/home/z/my-project/scripts/product_category_map.json') as f:
    cat_map = json.load(f)

MANUAL_WEIGHTS = {"M1051": 50.0, "CF101": 1.0, "S101": 1.0, "P109": 25.0}
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

# ===== EXTRACTION 2025 (S1 = Jan-Jun) =====
print("Lecture fichier 2025 (S1)...")
wb25 = openpyxl.load_workbook('/home/z/my-project/upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx', read_only=True, data_only=True)
ws25 = wb25.active

# Pour 2025 : colonnes natives
# Col 0: ref, Col 1: desc, Col 5: date, Col 6: CA, Col 9: état, Col 10: catégorie, Col 11: volume (t), Col 12: agence, Col 13: région

# Données par agence (CA) — S1 2025
agence_ca_2025 = defaultdict(float)
# Données par catégorie et par produit — S1 2025
cat_ca_2025 = defaultdict(float)
cat_vol_2025 = defaultdict(float)
prod_ca_2025 = defaultdict(float)
prod_vol_2025 = defaultdict(float)

for row in ws25.iter_rows(min_row=2, values_only=True):
    if not row or len(row) < 14: continue
    ref = row[0]
    desc = row[1]
    date_cmd = row[5]
    ca = row[6]
    etat = row[9]
    cat = row[10]
    vol_t = row[11]
    agence = row[12]
    
    if not ref: continue
    etat_str = str(etat) if etat else ""
    if etat_str != "Livrée": continue
    
    # Filtre S1 (Jan-Jun)
    month = None
    if date_cmd:
        if hasattr(date_cmd, 'month'):
            month = date_cmd.month
        else:
            try:
                s = str(date_cmd)
                parts = s.split("/")
                if len(parts) == 3: month = int(parts[1])
                elif len(s) >= 7: month = int(s[5:7])
            except: pass
    if month is None or month > 6: continue
    
    try:
        v = float(vol_t) if vol_t else 0
        c = float(ca) if ca else 0
    except: continue
    
    ref_str = str(ref).strip()
    cat_str = str(cat).strip() if cat else "DIVERS"
    if cat_str == "DIVERS2": cat_str = "DIVERS"
    
    agence_norm = normalize_agence(agence)
    if agence_norm:
        agence_ca_2025[agence_norm] += c
    
    cat_ca_2025[cat_str] += c
    cat_vol_2025[cat_str] += v
    prod_ca_2025[ref_str] += c
    prod_vol_2025[ref_str] += v

wb25.close()
print(f"  2025 S1 — CA total: {sum(agence_ca_2025.values())/1e6:.1f} M FCFA")
print(f"  2025 S1 — Vol total: {sum(cat_vol_2025.values()):.1f} t")

# ===== EXTRACTION 2026 (S1) =====
print("\nLecture fichier 2026 (S1)...")
wb26 = openpyxl.load_workbook('/home/z/my-project/upload/ventes janv a juin 2026.xlsx', read_only=True, data_only=True)

agence_ca_2026 = defaultdict(float)
cat_ca_2026 = defaultdict(float)
cat_vol_2026 = defaultdict(float)
prod_ca_2026 = defaultdict(float)
prod_vol_2026 = defaultdict(float)

sheet_to_month = {"Sheet 1": 1, "Feuil1": 2, "Feuil2": 3, "Feuil3": 4, "Feuil4": 5, "Feuil5": 6}

for sheet_name in wb26.sheetnames:
    ws = wb26[sheet_name]
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
            agence_ca_2026[agence_norm] += c
        
        cat_ca_2026[cat] += c
        cat_vol_2026[cat] += vol_t
        prod_ca_2026[ref_str] += c
        prod_vol_2026[ref_str] += vol_t

wb26.close()
print(f"  2026 S1 — CA total: {sum(agence_ca_2026.values())/1e6:.1f} M FCFA")
print(f"  2026 S1 — Vol total: {sum(cat_vol_2026.values()):.1f} t")

# ===== AGRÉGATION SELON LES RUBRIQUES DEMANDÉES =====
def get_rubrique_data(year_data_ca, year_data_vol, prod_ca, prod_vol):
    """Construit les valeurs par rubrique selon le mapping défini."""
    r = {}
    # VENTES TOURTEAU SOJA
    r['TOURTEAU SOJA'] = {
        'ca': year_data_ca.get('TOURTEAUX', 0),
        'vol': year_data_vol.get('TOURTEAUX', 0)
    }
    # VENTES INGREDIENTS = INGREDIENTS - M1051 (Maïs)
    r['INGREDIENTS'] = {
        'ca': year_data_ca.get('INGREDIENTS', 0) - prod_ca.get('M1051', 0),
        'vol': year_data_vol.get('INGREDIENTS', 0) - prod_vol.get('M1051', 0)
    }
    # MAIS = M1051
    r['MAIS'] = {
        'ca': prod_ca.get('M1051', 0),
        'vol': prod_vol.get('M1051', 0)
    }
    # VENTES MATERIEL ELEVAGE
    r['MATERIEL ELEVAGE'] = {
        'ca': year_data_ca.get('MATERIEL ELEVAGE', 0),
        'vol': year_data_vol.get('MATERIEL ELEVAGE', 0)
    }
    # ALIMENTS POISSON & CHICK & PIGLET BOOSTER = ALIMENT COMPLET
    r['ALIMENTS COMPLETS'] = {
        'ca': year_data_ca.get('ALIMENT COMPLET', 0),
        'vol': year_data_vol.get('ALIMENT COMPLET', 0)
    }
    # VENTES CONCENTRES
    r['CONCENTRES'] = {
        'ca': year_data_ca.get('CONCENTRES', 0),
        'vol': year_data_vol.get('CONCENTRES', 0)
    }
    # VENTES DIVERSES = DIVERS strict (hors services) + COMPLEMENT ALIMENTAIRE
    divers_strict_ca = year_data_ca.get('DIVERS', 0) - prod_ca.get('PONT_BASCULE', 0) - prod_ca.get('CONTRIBUTION_CARBURANT', 0)
    divers_strict_vol = year_data_vol.get('DIVERS', 0) - prod_vol.get('PONT_BASCULE', 0) - prod_vol.get('CONTRIBUTION_CARBURANT', 0)
    r['DIVERSES'] = {
        'ca': divers_strict_ca + year_data_ca.get('COMPLEMENT ALIMENTAIRE', 0),
        'vol': divers_strict_vol + year_data_vol.get('COMPLEMENT ALIMENTAIRE', 0)
    }
    # PRODUITS ACCESSOIRES = PONT_BASCULE + CONTRIBUTION_CARBURANT
    r['PRODUITS ACCESSOIRES'] = {
        'ca': prod_ca.get('PONT_BASCULE', 0) + prod_ca.get('CONTRIBUTION_CARBURANT', 0),
        'vol': prod_vol.get('PONT_BASCULE', 0) + prod_vol.get('CONTRIBUTION_CARBURANT', 0)
    }
    # ALVEOLES
    r['ALVEOLES'] = {
        'ca': year_data_ca.get('ALVEOLE', 0),
        'vol': year_data_vol.get('ALVEOLE', 0)
    }
    return r

rubriques_2025 = get_rubrique_data(cat_ca_2025, cat_vol_2025, prod_ca_2025, prod_vol_2025)
rubriques_2026 = get_rubrique_data(cat_ca_2026, cat_vol_2026, prod_ca_2026, prod_vol_2026)

print("\n=== RUBRIQUES (S1 2025 vs S1 2026) ===")
print(f"{'Rubrique':<35} {'CA 2025 (M)':>12} {'CA 2026 (M)':>12} {'Vol 2025 (t)':>14} {'Vol 2026 (t)':>14}")
print("-" * 90)
total_ca_25 = 0
total_ca_26 = 0
total_vol_25 = 0
total_vol_26 = 0
for r in ['TOURTEAU SOJA', 'INGREDIENTS', 'MAIS', 'MATERIEL ELEVAGE', 'ALIMENTS COMPLETS', 'CONCENTRES', 'DIVERSES', 'PRODUITS ACCESSOIRES', 'ALVEOLES']:
    ca25 = rubriques_2025[r]['ca']
    ca26 = rubriques_2026[r]['ca']
    v25 = rubriques_2025[r]['vol']
    v26 = rubriques_2026[r]['vol']
    print(f"{r:<35} {ca25/1e6:>12.1f} {ca26/1e6:>12.1f} {v25:>14.1f} {v26:>14.1f}")
    total_ca_25 += ca25
    total_ca_26 += ca26
    total_vol_25 += v25
    total_vol_26 += v26
print("-" * 90)
print(f"{'TOTAL':<35} {total_ca_25/1e6:>12.1f} {total_ca_26/1e6:>12.1f} {total_vol_25:>14.1f} {total_vol_26:>14.1f}")

# ===== DONNÉES PAR AGENCE =====
print("\n=== CA PAR AGENCE (S1 2025 vs S1 2026) ===")
print(f"{'Agence':<15} {'CA 2025 (M)':>12} {'CA 2026 (M)':>12}")
print("-" * 50)
# Mapping ordre du fichier (R117-R130)
agences_ordre = [
    ('YAOUNDE-MESSASI', 'Messassi'),
    ('YAOUNDE-AHALA', 'Ahala'),
    ('YAOUNDE-NKOABANG', 'Nkoabang'),
    ('YAOUNDE-NKOLBISSON', 'Nkolbisson'),
    ('BERTOUA', 'Bertoua'),
    ('NGAOUNDERE', 'Ngaoundere'),
    ('BAFOUSSAM-FAMLA', 'Famla'),
    ('BAFOUSSAM-DJELENG', 'Djeleng'),
    ('BAMENDA', 'Mbouda'),
    ('NKONGSAMBA', 'Nkongsamba'),
    ('BUEA', 'Buea'),
    ('DOUALA-AXE LOURD', 'Village'),
    ('DOUALA-SIEGE', 'Ndobo'),
    ('DOUALA-PK11', 'Pk11'),
]
for label, agence in agences_ordre:
    ca25 = agence_ca_2025.get(agence, 0)
    ca26 = agence_ca_2026.get(agence, 0)
    print(f"{label:<15} {ca25/1e6:>12.1f} {ca26/1e6:>12.1f}")
total_ca_ag_25 = sum(agence_ca_2025.values())
total_ca_ag_26 = sum(agence_ca_2026.values())
print(f"{'TOTAL':<15} {total_ca_ag_25/1e6:>12.1f} {total_ca_ag_26/1e6:>12.1f}")

# ===== SAUVEGARDER POUR UTILISATION ULTÉRIEURE =====
data_to_save = {
    'rubriques_2025': {k: {'ca': v['ca'], 'vol': v['vol']} for k, v in rubriques_2025.items()},
    'rubriques_2026': {k: {'ca': v['ca'], 'vol': v['vol']} for k, v in rubriques_2026.items()},
    'agence_ca_2025': dict(agence_ca_2025),
    'agence_ca_2026': dict(agence_ca_2026),
    'agences_ordre': [{'label': l, 'agence': a} for l, a in agences_ordre],
}
with open('/home/z/my-project/scripts/admin_vente_data.json', 'w', encoding='utf-8') as f:
    json.dump(data_to_save, f, ensure_ascii=False, indent=2, default=str)
print("\n✓ Données sauvegardées dans admin_vente_data.json")
