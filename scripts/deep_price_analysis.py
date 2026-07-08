"""
Deep price analysis:
1. Price by CONCENTRES sub-category (10% Chair vs 5% Chair vs 10% Porc vs 5% Ponte vs 10% Ponte)
2. Price by format for C104 (50Kg vs 25Kg vs 5Kg vs 1Kg)
3. Maïs vs other ingredients price separation
4. Price by agency and region
"""
import json, re
from openpyxl import load_workbook
from collections import defaultdict

SRC_2026 = "download/ventes_livrees.xlsx"
SRC_2025 = "upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx"

with open("scripts/product_category_map.json", "r") as f:
    CAT_MAP = json.load(f)
with open("scripts/zero_weight_products.json", "r") as f:
    ZERO_WEIGHT = set(json.load(f))

WEIGHT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(KG|Kg|kg|KG|GRAMMES?|G|L)\b", re.IGNORECASE)
def parse_weight(desc, ref=""):
    if ref in ZERO_WEIGHT: return 0.0
    if ref == "M1051": return 50.0
    if ref == "P109": return 25.0
    if not desc: return 0.0
    m = WEIGHT_RE.findall(str(desc))
    if not m: return 0.0
    val, unit = m[-1]
    val = float(val)
    u = unit.upper()
    if u == "KG": return val
    if u in ("G","GRAMME","GRAMMES"): return val/1000
    if u == "L": return val
    return 0.0

# Region mapping
AGENCE_REGION = {
    "Ahala": "Centre", "Bertoua": "Centre", "Buea": "Littoral",
    "Djeleng": "Ouest", "Famla": "Ouest", "Mbouda": "Ouest",
    "Messassi": "Centre", "Ndobo": "Littoral", "Ngaoundere": "Centre",
    "Nkoabang": "Centre", "Nkolbisson": "Centre", "Nkongsamba": "Littoral",
    "PK11": "Littoral", "Village": "Littoral",
    "SPC BAF-CHEFFERIE": "Ouest", "SPC BUEA": "Centre", "SPC KYE-OSSI": "Centre",
    "SPC VILLAGE": "Littoral", "SPC Village": "Littoral",
    "SPC-DLA-BERI": "Littoral", "SPC-DSCHANG": "Centre",
    "SPC-NDERE": "Centre", "SPC-TPO": "Centre", "SPC-YASSA": "Littoral",
    "PDC Emana": "Centre",
}

def norm_ag(s):
    if not s: return ""
    s = str(s).strip()
    if s.upper().startswith("AGENCE "): s = s[7:]
    if s.upper().startswith("DE "): s = s[3:]
    if "BAMENDA" in s.upper() and "MBOUDA" in s.upper(): return "Mbouda"
    if s: s = s[0].upper() + s[1:].lower()
    return s

# Read 2026
wb = load_workbook(SRC_2026, read_only=True, data_only=True)
pw = {}

# Per-product aggregates
product_data = defaultdict(lambda: {"vol_t": 0, "ca": 0, "rows": 0, "desc": ""})
# Per-agency-per-category
agency_cat = defaultdict(lambda: defaultdict(lambda: {"vol_t": 0, "ca": 0}))
# Per-region
region_data = defaultdict(lambda: {"vol_t": 0, "ca": 0})

for sn in wb.sheetnames:
    ws = wb[sn]
    for row in ws.iter_rows(min_row=3, values_only=True):
        tiers = row[5] if len(row)>5 else None
        if not tiers or not str(tiers).strip(): continue
        ref = str(row[0] or "").strip()
        if not ref: continue
        desc = str(row[1] or "")
        qte = row[2] if len(row)>2 else 0
        ca = row[8] if len(row)>8 else 0
        agence = row[17] if len(row)>17 else None

        if ref not in pw: pw[ref] = parse_weight(desc, ref)
        w = pw[ref]
        try: qf = float(qte) if qte else 0
        except: qf = 0
        try: cf = float(ca) if ca else 0
        except: cf = 0
        vol_t = qf * w / 1000
        cat = CAT_MAP.get(ref, "DIVERS")

        # Per product
        product_data[ref]["vol_t"] += vol_t
        product_data[ref]["ca"] += cf
        product_data[ref]["rows"] += 1
        if not product_data[ref]["desc"]:
            product_data[ref]["desc"] = desc

        # Per agency per category
        ag_norm = norm_ag(agence)
        if ag_norm:
            agency_cat[ag_norm][cat]["vol_t"] += vol_t
            agency_cat[ag_norm][cat]["ca"] += cf

            # Per region
            region = AGENCE_REGION.get(ag_norm, "Autre")
            region_data[region]["vol_t"] += vol_t
            region_data[region]["ca"] += cf

wb.close()

# ===== 1. PRICE BY CONCENTRES SUB-CATEGORY =====
CONC_REFS = {
    "C104": "BELGO 10% Chair 50Kg", "C1042": "BELGO 10% Chair 1Kg",
    "C1043": "BELGO 10% Chair 5Kg", "C1044": "BELGO 10% Chair 25Kg",
    "C105": "BELGO 10% Porc 50Kg", "C1053": "BELGO 10% Porc 1Kg",
    "C1054": "BELGO 10% Porc 5Kg", "C1055": "BELGO 10% Porc 25Kg",
    "C102": "BELGO 10% Ponte 50Kg", "C1022": "BELGO 10% Ponte 5Kg",
    "C103": "BELGO 5% Chair 50Kg", "C101": "BELGO 5% Ponte 50Kg",
}

# Group by sub-category
subcat_data = defaultdict(lambda: {"vol_t": 0, "ca": 0})
for ref, name in CONC_REFS.items():
    p = product_data.get(ref, {"vol_t": 0, "ca": 0})
    # Determine sub-category
    if "10%" in name and "Chair" in name: sc = "BELGO 10% Chair"
    elif "10%" in name and "Porc" in name: sc = "BELGO 10% Porc"
    elif "10%" in name and "Ponte" in name: sc = "BELGO 10% Ponte"
    elif "5%" in name and "Chair" in name: sc = "BELGO 5% Chair"
    elif "5%" in name and "Ponte" in name: sc = "BELGO 5% Ponte"
    else: sc = "Autres"
    subcat_data[sc]["vol_t"] += p["vol_t"]
    subcat_data[sc]["ca"] += p["ca"]

print("=" * 100)
print("1. PRIX PAR SOUS-CATÉGORIE DE CONCENTRÉS (S1 2026)")
print("=" * 100)
print(f"{'Sous-catégorie':<22}{'Vol (t)':>10}{'CA (M FCFA)':>14}{'Prix (FCFA/t)':>16}{'Prix (USD/t)':>14}")
print("-" * 76)
for sc in ["BELGO 10% Chair", "BELGO 5% Chair", "BELGO 10% Porc", "BELGO 5% Ponte", "BELGO 10% Ponte"]:
    d = subcat_data[sc]
    price = d["ca"] / d["vol_t"] if d["vol_t"] > 0 else 0
    usd = price / 600 if price > 0 else 0
    print(f"{sc:<22}{d['vol_t']:>10.1f}{d['ca']/1e6:>14.1f}{price:>16,.0f}{usd:>14,.0f}")

# ===== 2. PRICE BY FORMAT FOR C104 =====
print(f"\n{'=' * 100}")
print("2. PRIX PAR FORMAT POUR C104 (BELGO 10% Chair)")
print("=" * 100)
print(f"{'Réf':<8}{'Description':<30}{'Vol (t)':>10}{'CA (M FCFA)':>14}{'Prix (FCFA/t)':>16}{'Format':>10}")
print("-" * 88)
for ref in ["C104", "C1044", "C1043", "C1042"]:
    p = product_data.get(ref, {"vol_t": 0, "ca": 0, "desc": ""})
    price = p["ca"] / p["vol_t"] if p["vol_t"] > 0 else 0
    fmt = ref.replace("C104", "50Kg") if ref == "C104" else ref.replace("C1044", "25Kg").replace("C1043", "5Kg").replace("C1042", "1Kg")
    print(f"{ref:<8}{p['desc'][:28]:<30}{p['vol_t']:>10.1f}{p['ca']/1e6:>14.1f}{price:>16,.0f}{fmt:>10}")

# ===== 3. MAÏS vs OTHER INGREDIENTS =====
print(f"\n{'=' * 100}")
print("3. SÉPARATION MAÏS / AUTRES INGRÉDIENTS (S1 2026)")
print("=" * 100)
mais = product_data.get("M1051", {"vol_t": 0, "ca": 0, "desc": "MAIS"})
other_ing_ca = 0
other_ing_vol = 0
for ref, p in product_data.items():
    cat = CAT_MAP.get(ref, "")
    if cat == "INGREDIENTS" and ref != "M1051":
        other_ing_ca += p["ca"]
        other_ing_vol += p["vol_t"]

mais_price = mais["ca"] / mais["vol_t"] if mais["vol_t"] > 0 else 0
other_price = other_ing_ca / other_ing_vol if other_ing_vol > 0 else 0
mixed_price = (mais["ca"] + other_ing_ca) / (mais["vol_t"] + other_ing_vol) if (mais["vol_t"] + other_ing_vol) > 0 else 0

print(f"{'Composant':<22}{'Vol (t)':>10}{'CA (M FCFA)':>14}{'Prix (FCFA/t)':>16}")
print("-" * 62)
print(f"{'Maïs (M1051)':<22}{mais['vol_t']:>10.1f}{mais['ca']/1e6:>14.1f}{mais_price:>16,.0f}")
print(f"{'Autres ingrédients':<22}{other_ing_vol:>10.1f}{other_ing_ca/1e6:>14.1f}{other_price:>16,.0f}")
print(f"{'INGREDIENTS (mixte)':<22}{mais['vol_t']+other_ing_vol:>10.1f}{(mais['ca']+other_ing_ca)/1e6:>14.1f}{mixed_price:>16,.0f}")

# ===== 4. PRICE BY AGENCY AND REGION =====
print(f"\n{'=' * 100}")
print("4. PRIX MOYEN PAR AGENCE (toutes catégories, S1 2026)")
print("=" * 100)
print(f"{'Agence':<16}{'Région':<12}{'Vol (t)':>10}{'CA (M FCFA)':>14}{'Prix (FCFA/t)':>16}")
print("-" * 68)
for ag in sorted(agency_cat.keys(), key=lambda a: -sum(agency_cat[a][c]["ca"] for c in agency_cat[a])):
    vol = sum(agency_cat[ag][c]["vol_t"] for c in agency_cat[ag])
    ca = sum(agency_cat[ag][c]["ca"] for c in agency_cat[ag])
    price = ca / vol if vol > 0 else 0
    region = AGENCE_REGION.get(ag, "?")
    if ca > 0:
        print(f"{ag:<16}{region:<12}{vol:>10.1f}{ca/1e6:>14.1f}{price:>16,.0f}")

print(f"\n{'=' * 100}")
print("5. PRIX MOYEN PAR RÉGION (toutes catégories, S1 2026)")
print("=" * 100)
print(f"{'Région':<14}{'Vol (t)':>10}{'CA (M FCFA)':>14}{'Prix (FCFA/t)':>16}{'% CA':>8}")
print("-" * 62)
total_ca = sum(r["ca"] for r in region_data.values())
for region in sorted(region_data.keys(), key=lambda r: -region_data[r]["ca"]):
    d = region_data[region]
    price = d["ca"] / d["vol_t"] if d["vol_t"] > 0 else 0
    pct = d["ca"] / total_ca * 100 if total_ca > 0 else 0
    print(f"{region:<14}{d['vol_t']:>10.1f}{d['ca']/1e6:>14.1f}{price:>16,.0f}{pct:>7.1f}%")

# CONCENTRES by region
print(f"\n{'=' * 100}")
print("6. CONCENTRÉS PAR RÉGION (S1 2026)")
print("=" * 100)
region_conc = defaultdict(lambda: {"vol_t": 0, "ca": 0})
for ag in agency_cat:
    region = AGENCE_REGION.get(ag, "?")
    d = agency_cat[ag].get("CONCENTRES", {"vol_t": 0, "ca": 0})
    region_conc[region]["vol_t"] += d["vol_t"]
    region_conc[region]["ca"] += d["ca"]

print(f"{'Région':<14}{'Vol conc (t)':>14}{'CA conc (M)':>14}{'Prix conc':>16}{'% vol':>8}")
print("-" * 66)
total_conc_vol = sum(r["vol_t"] for r in region_conc.values())
for region in sorted(region_conc.keys(), key=lambda r: -region_conc[r]["vol_t"]):
    d = region_conc[region]
    price = d["ca"] / d["vol_t"] if d["vol_t"] > 0 else 0
    pct = d["vol_t"] / total_conc_vol * 100 if total_conc_vol > 0 else 0
    print(f"{region:<14}{d['vol_t']:>14.1f}{d['ca']/1e6:>14.1f}{price:>16,.0f}{pct:>7.1f}%")

# Save
output = {
    "conc_subcat_price": {sc: {"vol_t": subcat_data[sc]["vol_t"], "ca": subcat_data[sc]["ca"],
                                "price": subcat_data[sc]["ca"]/subcat_data[sc]["vol_t"] if subcat_data[sc]["vol_t"]>0 else 0}
                          for sc in subcat_data},
    "c104_formats": {ref: {"vol_t": product_data[ref]["vol_t"], "ca": product_data[ref]["ca"],
                           "price": product_data[ref]["ca"]/product_data[ref]["vol_t"] if product_data[ref]["vol_t"]>0 else 0,
                           "desc": product_data[ref]["desc"]}
                     for ref in ["C104", "C1044", "C1043", "C1042"]},
    "mais_vs_other": {"mais": {"vol_t": mais["vol_t"], "ca": mais["ca"], "price": mais_price},
                      "other": {"vol_t": other_ing_vol, "ca": other_ing_ca, "price": other_price},
                      "mixed_price": mixed_price},
    "region_data": {r: region_data[r] for r in region_data},
    "region_conc": {r: region_conc[r] for r in region_conc},
    "agency_region": AGENCE_REGION,
}
with open("scripts/deep_price_analysis.json", "w") as f:
    json.dump(output, f, ensure_ascii=False, indent=2, default=str)
