"""
Met à jour ADMINISTRATEUR_DE_VENTE_rempli.xlsx avec les données Janvier → Septembre (2025 vs 2026).

Méthodologie identique au remplissage S1 final (fill_admin_vente_v2.py) :
- état « Livrée » uniquement (services PONT_BASCULE / CONTRIBUTION_CARBURANT acceptés en « Validée » côté 2026)
- Montant HT (convention du fichier, col 6 en 2025 / col 8 en 2026)
- catégorie = cat_map.json pour LES DEUX années (comparabilité)
- 2025 (Jan-Sep) : upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx
    (ref=0, qté=2, date=5, HT=6, état=9, vol natif t=11, agence=12) — volume M1051 recalculé qté×50kg
- 2026 Jan-Juin : upload/ventes janv a juin 2026.xlsx — 6 feuilles, 17 colonnes
    (ref=0, desc=1, qté=2, HT=8, état=15, agence=17)
- 2026 Juillet : ERP (9), Août : ERP (27), Septembre : ERP (51) — 16 colonnes
    (ref=0, desc=1, qté=2, HT=8, état=13, agence=15)

Écrit dans download/ADMINISTRATEUR_DE_VENTE_rempli.xlsx (structure et formules conservées) :
- Feuille 1 « CAHT par agence A » : B117-B130 (2026), D117-D130 (2025)
- Feuille 2 « Ventes par produits A » : L46-L55 (2026), N46-N55 (2025), ligne PREMIX R51
- Feuille 3 « Vente en volume A » : D24-D29 (2026), E24-E29 (2025)
"""
import openpyxl
import json
import re
from collections import defaultdict

# ===== Outils repris de fill_admin_vente.py =====
with open('scripts/product_category_map.json', encoding='utf-8') as f:
    cat_map = json.load(f)

MANUAL_WEIGHTS = {"M1051": 50.0, "CF101": 1.0, "S101": 1.0, "P109": 25.0}
SERVICES_VALIDEE = {"PONT_BASCULE", "CONTRIBUTION_CARBURANT"}
WEIGHT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(KG|Kg|kg|KG|GRAMMES?|G|L)\b", re.IGNORECASE)


def parse_weight_kg(ref, desc):
    ref_str = str(ref).strip() if ref else ""
    if ref_str in MANUAL_WEIGHTS:
        return MANUAL_WEIGHTS[ref_str]
    if not desc:
        return 0.0
    matches = WEIGHT_RE.findall(str(desc))
    if not matches:
        return 0.0
    val_str, unit = matches[-1]
    val = float(val_str)
    u = unit.upper()
    if u == "KG":
        return val
    if u in ("G", "GRAMME", "GRAMMES"):
        return val / 1000.0
    if u == "L":
        return val
    return 0.0


def normalize_agence(agence_str):
    if not agence_str:
        return ""
    s = str(agence_str).strip()
    if s.upper().startswith("SPC") or s.upper().startswith("PDC"):
        return ""
    if "AGENCE " in s.upper():
        s = s[7:]
    if s.upper().startswith("DE "):
        s = s[3:]
    if "BAMENDA" in s.upper() and "MBOUDA" in s.upper():
        return "Mbouda"
    if "BERTOUA" in s.upper():
        return "Bertoua"
    if s:
        s = s[0].upper() + s[1:].lower()
    return s


def month_of(d):
    if d is None:
        return None
    if hasattr(d, 'month'):
        return d.month
    try:
        s = str(d)
        parts = s.split("/")
        if len(parts) == 3:
            return int(parts[1])
        if len(s) >= 7:
            return int(s[5:7])
    except Exception:
        pass
    return None


# ===== EXTRACTION 2025 (Jan-Sep) =====
print("Lecture fichier 2025 (Jan-Sep)...")
wb25 = openpyxl.load_workbook('upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx', read_only=True, data_only=True)
ws25 = wb25.active

agence_ca_2025 = defaultdict(float)
cat_ca_2025 = defaultdict(float)
cat_vol_2025 = defaultdict(float)
prod_ca_2025 = defaultdict(float)
prod_vol_2025 = defaultdict(float)

for row in ws25.iter_rows(min_row=2, values_only=True):
    if not row or len(row) < 14:
        continue
    ref, qte, date_cmd, ca, etat, agence = row[0], row[2], row[5], row[6], row[9], row[12]
    if not ref:
        continue
    if (str(etat) if etat else "") != "Livrée":
        continue
    m = month_of(date_cmd)
    if m is None or m > 9:
        continue
    try:
        c = float(ca) if ca else 0
    except Exception:
        continue
    ref_str = str(ref).strip()
    # catégorie = cat_map (même mapping que 2026, pas la catégorie native du fichier)
    cat = cat_map.get(ref_str, "DIVERS")
    # volume : M1051 recalculé qté×50kg (vol natif = 0 dans le fichier), sinon vol natif col 11
    if ref_str == "M1051":
        try:
            q = float(qte) if qte else 0
            vol_t = q * 50 / 1000
        except Exception:
            vol_t = 0
    else:
        try:
            vol_t = float(row[11]) if row[11] else 0
        except Exception:
            vol_t = 0
    agence_norm = normalize_agence(agence)
    if agence_norm:
        agence_ca_2025[agence_norm] += c
    cat_ca_2025[cat] += c
    cat_vol_2025[cat] += vol_t
    prod_ca_2025[ref_str] += c
    prod_vol_2025[ref_str] += vol_t
wb25.close()
print(f"  2025 Jan-Sep — CA total: {sum(agence_ca_2025.values())/1e6:.1f} M FCFA | Vol: {sum(cat_vol_2025.values()):.1f} t")


def read_2026_rows(wb, sheet_names, col_ca, col_etat, col_agence, expected_month=None):
    """Parcourt les lignes ERP 2026 et agrège (CA par agence/catégorie/produit)."""
    for sheet_name in sheet_names:
        ws = wb[sheet_name]
        if ws.cell(2, 1).value != 'Réf. produit':
            continue
        for row in ws.iter_rows(min_row=3, values_only=True):
            if not row or len(row) <= max(col_ca, col_etat, col_agence):
                continue
            ref, desc, qte = row[0], row[1], row[2]
            ca, etat, agence = row[col_ca], row[col_etat], row[col_agence]
            if not ref:
                continue
            ref_str = str(ref).strip()
            etat_str = str(etat) if etat else ""
            if etat_str != "Livrée":
                if not (ref_str in SERVICES_VALIDEE and etat_str == "Validée"):
                    continue
            if expected_month is not None and month_of(row[6]) != expected_month:
                continue
            try:
                q = float(qte) if qte else 0
                c = float(ca) if ca else 0
            except Exception:
                continue
            if ref_str == "M1051" and c == 0:
                continue
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


agence_ca_2026 = defaultdict(float)
cat_ca_2026 = defaultdict(float)
cat_vol_2026 = defaultdict(float)
prod_ca_2026 = defaultdict(float)
prod_vol_2026 = defaultdict(float)

# Janvier-Juin : classeur S1 (17 colonnes : HT=8, état=15, agence=17)
print("Lecture fichier 2026 Jan-Juin (ventes janv a juin 2026.xlsx)...")
wb26 = openpyxl.load_workbook('upload/ventes janv a juin 2026.xlsx', read_only=True, data_only=True)
read_2026_rows(wb26, wb26.sheetnames, col_ca=8, col_etat=15, col_agence=17)
wb26.close()

# Juillet / Août / Septembre : ERP (9), (27), (51) — 16 colonnes (HT=8, état=13, agence=15)
ERP_MENSUELS = [
    ('upload/NJS GROUP ERP - Lignes de commandes + multicompany (9).xlsx', 7),
    ('upload/NJS GROUP ERP - Lignes de commandes + multicompany (27).xlsx', 8),
    ('upload/NJS GROUP ERP - Lignes de commandes + multicompany (51).xlsx', 9),
]
for path, mois in ERP_MENSUELS:
    print(f"Lecture fichier 2026 mois {mois} ({path.split('/')[-1]})...")
    wbm = openpyxl.load_workbook(path, read_only=True, data_only=True)
    read_2026_rows(wbm, ['Sheet 1'], col_ca=8, col_etat=13, col_agence=15, expected_month=mois)
    wbm.close()
print(f"  2026 Jan-Sep — CA total: {sum(agence_ca_2026.values())/1e6:.1f} M FCFA | Vol: {sum(cat_vol_2026.values()):.1f} t")

# ===== AGRÉGATION PAR RUBRIQUE (mapping identique à fill_admin_vente.py) =====
def get_rubrique_data(ca, vol, prod_ca, prod_vol):
    r = {}
    r['TOURTEAU SOJA'] = {'ca': ca.get('TOURTEAUX', 0), 'vol': vol.get('TOURTEAUX', 0)}
    r['INGREDIENTS'] = {
        'ca': ca.get('INGREDIENTS', 0) - prod_ca.get('M1051', 0),
        'vol': vol.get('INGREDIENTS', 0) - prod_vol.get('M1051', 0)}
    r['MAIS'] = {'ca': prod_ca.get('M1051', 0), 'vol': prod_vol.get('M1051', 0)}
    r['MATERIEL ELEVAGE'] = {'ca': ca.get('MATERIEL ELEVAGE', 0), 'vol': vol.get('MATERIEL ELEVAGE', 0)}
    r['ALIMENTS COMPLETS'] = {'ca': ca.get('ALIMENT COMPLET', 0), 'vol': vol.get('ALIMENT COMPLET', 0)}
    r['CONCENTRES'] = {'ca': ca.get('CONCENTRES', 0), 'vol': vol.get('CONCENTRES', 0)}
    divers_strict_ca = ca.get('DIVERS', 0) - prod_ca.get('PONT_BASCULE', 0) - prod_ca.get('CONTRIBUTION_CARBURANT', 0)
    divers_strict_vol = vol.get('DIVERS', 0) - prod_vol.get('PONT_BASCULE', 0) - prod_vol.get('CONTRIBUTION_CARBURANT', 0)
    r['DIVERSES'] = {
        'ca': divers_strict_ca + ca.get('COMPLEMENT ALIMENTAIRE', 0),
        'vol': divers_strict_vol + vol.get('COMPLEMENT ALIMENTAIRE', 0)}
    r['PRODUITS ACCESSOIRES'] = {
        'ca': prod_ca.get('PONT_BASCULE', 0) + prod_ca.get('CONTRIBUTION_CARBURANT', 0),
        'vol': prod_vol.get('PONT_BASCULE', 0) + prod_vol.get('CONTRIBUTION_CARBURANT', 0)}
    r['ALVEOLES'] = {'ca': ca.get('ALVEOLE', 0), 'vol': vol.get('ALVEOLE', 0)}
    r['PREMIX'] = {'ca': ca.get('PREMIX', 0), 'vol': vol.get('PREMIX', 0)}
    return r

rubriques_2025 = get_rubrique_data(cat_ca_2025, cat_vol_2025, prod_ca_2025, prod_vol_2025)
rubriques_2026 = get_rubrique_data(cat_ca_2026, cat_vol_2026, prod_ca_2026, prod_vol_2026)

print("\n=== RUBRIQUES (Jan-Sep 2025 vs Jan-Sep 2026) ===")
print(f"{'Rubrique':<24} {'CA 2025 (M)':>13} {'CA 2026 (M)':>13} {'Vol 2025 (t)':>13} {'Vol 2026 (t)':>13}")
print("-" * 80)
ORDRE_RUBRIQUES = ['TOURTEAU SOJA', 'INGREDIENTS', 'MAIS', 'MATERIEL ELEVAGE', 'ALIMENTS COMPLETS',
                   'CONCENTRES', 'PREMIX', 'DIVERSES', 'PRODUITS ACCESSOIRES', 'ALVEOLES']
for r in ORDRE_RUBRIQUES:
    print(f"{r:<24} {rubriques_2025[r]['ca']/1e6:>13.1f} {rubriques_2026[r]['ca']/1e6:>13.1f} "
          f"{rubriques_2025[r]['vol']:>13.1f} {rubriques_2026[r]['vol']:>13.1f}")

# ===== AGRÉGATION PAR AGENCE =====
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
print("\n=== CA PAR AGENCE (Jan-Sep 2025 vs Jan-Sep 2026) ===")
print(f"{'Agence':<16} {'CA 2025 (M)':>13} {'CA 2026 (M)':>13}")
print("-" * 45)
for label, agence in agences_ordre:
    print(f"{label:<16} {agence_ca_2025.get(agence, 0)/1e6:>13.1f} {agence_ca_2026.get(agence, 0)/1e6:>13.1f}")

# ===== CONTRÔLE : cohérence avec le S1 livré (Jan-Sep doit être >= S1) =====
with open('scripts/admin_vente_data.json', encoding='utf-8') as f:
    s1_data = json.load(f)
print("\n=== CONTRÔLE vs S1 livré (deltas = contribution Juil-Sep) ===")
ok = True
for r in ['TOURTEAU SOJA', 'INGREDIENTS', 'MAIS', 'MATERIEL ELEVAGE', 'ALIMENTS COMPLETS',
          'CONCENTRES', 'DIVERSES', 'PRODUITS ACCESSOIRES', 'ALVEOLES']:
    d_ca = rubriques_2026[r]['ca'] - s1_data['rubriques_2026'][r]['ca']
    d_vol = rubriques_2026[r]['vol'] - s1_data['rubriques_2026'][r]['vol']
    flag = "" if (d_ca >= -1 and d_vol >= -1) else "  <-- ANOMALIE"
    if flag:
        ok = False
    print(f"  {r:<22} ΔCA 2026 {d_ca/1e6:>9.1f} M | ΔVol 2026 {d_vol:>9.1f} t{flag}")
print(f"  -> cohérence S1: {'OK' if ok else 'PROBLEME'}")

# ===== SAUVEGARDE JSON =====
data_to_save = {
    'rubriques_2025': {k: {'ca': v['ca'], 'vol': v['vol']} for k, v in rubriques_2025.items()},
    'rubriques_2026': {k: {'ca': v['ca'], 'vol': v['vol']} for k, v in rubriques_2026.items()},
    'agence_ca_2025': dict(agence_ca_2025),
    'agence_ca_2026': dict(agence_ca_2026),
    'agences_ordre': [{'label': l, 'agence': a} for l, a in agences_ordre],
}
with open('scripts/admin_vente_data_janv_sept.json', 'w', encoding='utf-8') as f:
    json.dump(data_to_save, f, ensure_ascii=False, indent=2, default=str)
print("\n✓ Données sauvegardées dans scripts/admin_vente_data_janv_sept.json")

# ===== ÉCRITURE DANS LE FICHIER REMPLI (structure conservée) =====
SRC_OUT = 'download/ADMINISTRATEUR_DE_VENTE_rempli.xlsx'
wb = openpyxl.load_workbook(SRC_OUT)

# Feuille 1 : CAHT par agence A — B117-B130 (2026), D117-D130 (2025)
ws1 = wb['CAHT par agence A']
for i, item in enumerate(agences_ordre):
    row = 117 + i
    ws1.cell(row=row, column=2, value=round(agence_ca_2026.get(item[1], 0), 2))
    ws1.cell(row=row, column=4, value=round(agence_ca_2025.get(item[1], 0), 2))

# Feuille 2 : Ventes par produits A — L46-L55 (2026), N46-N55 (2025)
# Structure actuelle du fichier : R46 TOURTEAU SOJA ... R51 PREMIX ... R55 MAIS
ws2 = wb['Ventes par produits A']
rows2 = [
    (46, 'TOURTEAU SOJA'),
    (47, 'INGREDIENTS'),
    (48, 'MATERIEL ELEVAGE'),
    (49, 'ALIMENTS COMPLETS'),
    (50, 'CONCENTRES'),
    (51, 'PREMIX'),
    (52, 'DIVERSES'),
    (53, 'PRODUITS ACCESSOIRES'),
    (54, 'ALVEOLES'),
    (55, 'MAIS'),
]
for row, key in rows2:
    ws2.cell(row=row, column=12, value=round(rubriques_2026[key]['ca'], 2))
    ws2.cell(row=row, column=14, value=round(rubriques_2025[key]['ca'], 2))

# Feuille 3 : Vente en volume A — D24-D29 (2026), E24-E29 (2025)
ws3 = wb['Vente en volume A']
rows3 = [
    (24, 'TOURTEAU SOJA'),
    (25, 'INGREDIENTS'),
    (26, 'MAIS'),
    (27, 'ALIMENTS COMPLETS'),
    (28, 'CONCENTRES'),
    (29, 'DIVERSES'),
]
for row, key in rows3:
    ws3.cell(row=row, column=4, value=round(rubriques_2026[key]['vol'], 2))
    ws3.cell(row=row, column=5, value=round(rubriques_2025[key]['vol'], 2))

wb.save(SRC_OUT)
print(f"✓ Fichier mis à jour: {SRC_OUT}")

# Mise à jour de premix_ca.json (utilisé par add_premix_row.py)
with open('scripts/premix_ca.json', 'w', encoding='utf-8') as f:
    json.dump({'ca_2025': rubriques_2025['PREMIX']['ca'],
               'ca_2026': rubriques_2026['PREMIX']['ca']}, f, ensure_ascii=False, indent=2)
print("✓ scripts/premix_ca.json mis à jour (Jan-Sep)")
