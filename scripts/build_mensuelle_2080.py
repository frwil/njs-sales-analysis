"""
Génère un fichier Excel séparé avec:
- Feuille 1: Ventes mensuelles Déc 2025 → Juin 2026 par agence vs objectifs (avec %)
- Feuilles 2-4: Clients 20/80 par région (1 feuille par région)
"""
import openpyxl
import json
import re
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter

# ===== CONFIG =====
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

AGENCE_REGION = {
    "Ahala": "Centre", "Bertoua": "Centre", "Buea": "Littoral",
    "Djeleng": "Ouest", "Famla": "Ouest", "Mbouda": "Ouest",
    "Messassi": "Centre", "Ndobo": "Littoral", "Ngaoundere": "Centre",
    "Nkoabang": "Centre", "Nkolbisson": "Centre", "Nkongsamba": "Littoral",
    "Pk11": "Littoral", "Village": "Littoral",
}

AGENCES_ORDER = ["Famla", "Ndobo", "Messassi", "Djeleng", "Mbouda", "Village",
                 "Bertoua", "Nkongsamba", "Nkoabang", "Ngaoundere", "Buea",
                 "Ahala", "Nkolbisson", "Pk11"]

MONTHS_LABELS = ["Déc 2025", "Jan 2026", "Fév 2026", "Mar 2026", "Avr 2026", "Mai 2026", "Juin 2026"]

# ===== LIRE DÉCEMBRE 2025 =====
print("Lecture décembre 2025...")
wb25 = openpyxl.load_workbook('/home/z/my-project/upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx', read_only=True, data_only=True)
ws25 = wb25.active

agence_month_vol = defaultdict(lambda: defaultdict(float))  # agence → month_key → vol

for row in ws25.iter_rows(min_row=2, values_only=True):
    if not row or len(row) < 14: continue
    ref = row[0]
    desc = row[1]
    ca = row[6]
    etat = row[9]
    date_cmd = row[5]
    agence = row[12]
    
    if not ref: continue
    if str(etat) != "Livrée": continue
    
    # Only December (month=12)
    if not date_cmd: continue
    if hasattr(date_cmd, 'month'):
        if date_cmd.month != 12: continue
    else:
        try:
            s = str(date_cmd)
            parts = s.split("/")
            if len(parts) == 3 and int(parts[1]) != 12: continue
        except: continue
    
    ref_str = str(ref).strip()
    cat = cat_map.get(ref_str, "DIVERS")
    if cat == "MATERIEL ELEVAGE": continue  # pas de volume pour matériel
    
    try: c = float(ca) if ca else 0
    except: continue
    
    # Volume (recalcul Maïs)
    if ref_str == "M1051":
        try:
            q = float(row[2]) if row[2] else 0
            vol_t = q * 50 / 1000
        except: vol_t = 0
    else:
        try: vol_t = float(row[11]) if row[11] else 0
        except: vol_t = 0
    
    agence_norm = normalize_agence(agence)
    if agence_norm:
        agence_month_vol[agence_norm]["Dec2025"] += vol_t

wb25.close()
print("  Décembre 2025 chargé.")

# ===== LIRE JANVIER-JUIN 2026 =====
print("Lecture Janvier-Juin 2026...")
wb26 = openpyxl.load_workbook('/home/z/my-project/upload/ventes janv a juin 2026.xlsx', read_only=True, data_only=True)
sheet_to_month = {"Sheet 1": "Jan2026", "Feuil1": "Fev2026", "Feuil2": "Mar2026",
                  "Feuil3": "Avr2026", "Feuil4": "Mai2026", "Feuil5": "Juin2026"}

for sheet_name in wb26.sheetnames:
    ws = wb26[sheet_name]
    if ws.cell(2, 1).value != 'Réf. produit': continue
    month_key = sheet_to_month.get(sheet_name)
    if month_key is None: continue
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
        cat = cat_map.get(ref_str, "DIVERS")
        if cat == "MATERIEL ELEVAGE": continue
        weight = parse_weight_kg(ref_str, desc)
        vol_t = q * weight / 1000.0
        agence_norm = normalize_agence(agence)
        if agence_norm:
            agence_month_vol[agence_norm][month_key] += vol_t

wb26.close()
print("  Janvier-Juin 2026 chargé.")

# ===== LIRE OBJECTIFS MENSUELS PAR AGENCE =====
print("Lecture objectifs mensuels par agence...")
with open('/home/z/my-project/scripts/objectives_comparison.json') as f:
    obj_data = json.load(f)

agence_month_obj = defaultdict(lambda: defaultdict(float))
for agence, cats in obj_data.get('agence_objectives', {}).items():
    agence_norm = normalize_agence(agence)
    if not agence_norm: continue
    for cat, months in cats.items():
        if cat == "MATERIEL ELEVAGE": continue
        for m_str, vol in months.items():
            m = int(m_str)
            if 1 <= m <= 6:  # S1 2026
                month_key = ["Jan2026", "Fev2026", "Mar2026", "Avr2026", "Mai2026", "Juin2026"][m-1]
                agence_month_obj[agence_norm][month_key] += float(vol)

# Pas d'objectif pour Déc 2025 (c'est du réalisé)
print("  Objectifs chargés.")

# ===== CRÉER FICHIER EXCEL =====
print("\nCréation du fichier Excel...")
OUT = '/home/z/my-project/download/analyse_mensuelle_2080.xlsx'
wb_out = openpyxl.Workbook()

# Styles
NAVY = "1F4E78"
GOLD = "FFD966"
WHITE = "FFFFFF"
GRAY = "808080"
LIGHT_GRAY = "F2F2F2"
GREEN = "C6EFCE"
RED = "FFC7CE"

HEADER_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
HEADER_FONT = Font(bold=True, color=WHITE, size=10)
TOTAL_FILL = PatternFill(start_color=GOLD, end_color=GOLD, fill_type="solid")
TOTAL_FONT = Font(bold=True, size=10)
BAND_FILL = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
thin_border = Border(
    left=Side(style='thin', color=GRAY),
    right=Side(style='thin', color=GRAY),
    top=Side(style='thin', color=GRAY),
    bottom=Side(style='thin', color=GRAY)
)

# ===== FEUILLE 1: VENTES MENSUELLES PAR AGENCE =====
ws1 = wb_out.active
ws1.title = "Ventes mensuelles par agence"

# Titre
ws1.cell(1, 1, "VENTES MENSUELLES PAR AGENCE — Décembre 2025 à Juin 2026 (vs objectifs)")
ws1.cell(1, 1).font = Font(bold=True, size=14, color=NAVY)
ws1.merge_cells('A1:T1')

# Header row 3
headers = ["Agence", "Région"]
for ml in MONTHS_LABELS:
    headers.append(f"{ml} obj")
    headers.append(f"{ml} réel")
    headers.append("%")
headers.append("S1 2026 obj")
headers.append("S1 2026 réel")
headers.append("S1 2026 %")

for col, h in enumerate(headers, 1):
    c = ws1.cell(3, col, h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = thin_border

# Données
row_idx = 4
for ag in AGENCES_ORDER:
    region = AGENCE_REGION.get(ag, "")
    ws1.cell(row_idx, 1, ag).font = Font(bold=True)
    ws1.cell(row_idx, 2, region)
    
    s1_obj = 0
    s1_real = 0
    col = 3
    for ml_idx, month_key in enumerate(["Dec2025", "Jan2026", "Fev2026", "Mar2026", "Avr2026", "Mai2026", "Juin2026"]):
        obj = agence_month_obj.get(ag, {}).get(month_key, 0)
        real = agence_month_vol.get(ag, {}).get(month_key, 0)
        
        # Déc 2025 n'a pas d'objectif
        if month_key == "Dec2025":
            ws1.cell(row_idx, col, "—")
            ws1.cell(row_idx, col+1, round(real, 0))
            ws1.cell(row_idx, col+2, "—")
        else:
            ws1.cell(row_idx, col, round(obj, 0))
            ws1.cell(row_idx, col+1, round(real, 0))
            pct = real/obj*100 if obj > 0 else 0
            pct_cell = ws1.cell(row_idx, col+2, f"{pct:.0f}%")
            if pct >= 100:
                pct_cell.fill = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
            elif pct < 80:
                pct_cell.fill = PatternFill(start_color=RED, end_color=RED, fill_type="solid")
            s1_obj += obj
            s1_real += real
        
        col += 3
    
    # S1 totals (Jan-Juin, pas Dec)
    ws1.cell(row_idx, col, round(s1_obj, 0))
    ws1.cell(row_idx, col+1, round(s1_real, 0))
    pct_s1 = s1_real/s1_obj*100 if s1_obj > 0 else 0
    pct_cell = ws1.cell(row_idx, col+2, f"{pct_s1:.0f}%")
    pct_cell.font = Font(bold=True)
    if pct_s1 >= 100:
        pct_cell.fill = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
    elif pct_s1 < 80:
        pct_cell.fill = PatternFill(start_color=RED, end_color=RED, fill_type="solid")
    
    if row_idx % 2 == 0:
        for c in range(1, col+3):
            if not ws1.cell(row_idx, c).fill or ws1.cell(row_idx, c).fill.start_color.rgb in [None, "00000000"]:
                ws1.cell(row_idx, c).fill = BAND_FILL
    
    for c in range(1, col+3):
        ws1.cell(row_idx, c).border = thin_border
    
    row_idx += 1

# Ligne TOTAL
ws1.cell(row_idx, 1, "TOTAL").font = TOTAL_FONT
ws1.cell(row_idx, 1).fill = TOTAL_FILL
ws1.cell(row_idx, 2, "").fill = TOTAL_FILL

total_s1_obj = 0
total_s1_real = 0
col = 3
for month_key in ["Dec2025", "Jan2026", "Fev2026", "Mar2026", "Avr2026", "Mai2026", "Juin2026"]:
    total_obj = sum(agence_month_obj.get(ag, {}).get(month_key, 0) for ag in AGENCES_ORDER)
    total_real = sum(agence_month_vol.get(ag, {}).get(month_key, 0) for ag in AGENCES_ORDER)
    
    if month_key == "Dec2025":
        ws1.cell(row_idx, col, "—")
        ws1.cell(row_idx, col+1, round(total_real, 0))
        ws1.cell(row_idx, col+2, "—")
    else:
        ws1.cell(row_idx, col, round(total_obj, 0))
        ws1.cell(row_idx, col+1, round(total_real, 0))
        pct = total_real/total_obj*100 if total_obj > 0 else 0
        ws1.cell(row_idx, col+2, f"{pct:.0f}%")
        total_s1_obj += total_obj
        total_s1_real += total_real
    
    for c in range(col, col+3):
        ws1.cell(row_idx, c).fill = TOTAL_FILL
        ws1.cell(row_idx, c).font = TOTAL_FONT
    col += 3

ws1.cell(row_idx, col, round(total_s1_obj, 0))
ws1.cell(row_idx, col+1, round(total_s1_real, 0))
pct_total = total_s1_real/total_s1_obj*100 if total_s1_obj > 0 else 0
ws1.cell(row_idx, col+2, f"{pct_total:.0f}%")
for c in range(col, col+3):
    ws1.cell(row_idx, c).fill = TOTAL_FILL
    ws1.cell(row_idx, c).font = TOTAL_FONT

for c in range(1, col+3):
    ws1.cell(row_idx, c).border = thin_border

# Largeurs colonnes
ws1.column_dimensions['A'].width = 14
ws1.column_dimensions['B'].width = 10
for c in range(3, col+3):
    ws1.column_dimensions[get_column_letter(c)].width = 11

ws1.freeze_panes = "C4"
print(f"  Feuille 1 créée: {row_idx-3} agences + total")

# ===== FEUILLES 2-4: CLIENTS 20/80 PAR RÉGION =====
# Charger clients exclus
with open('/home/z/my-project/scripts/excluded_clients.json') as f:
    excluded_tiers = set(json.load(f))

COMPTOIR_RE = re.compile(r"CLIENTS?\s+COMPTOIR", re.IGNORECASE)

print("\nCalcul des clients 20/80...")

# Lire ventes 2026 S1 par client (avec agence)
wb26 = openpyxl.load_workbook('/home/z/my-project/upload/ventes janv a juin 2026.xlsx', read_only=True, data_only=True)

client_ca = defaultdict(float)  # tiers → CA total
client_agence = {}  # tiers → agence
client_nom = {}  # tiers → nom

for sheet_name in wb26.sheetnames:
    ws = wb26[sheet_name]
    if ws.cell(2, 1).value != 'Réf. produit': continue
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row or len(row) < 16: continue
        ref = row[0]
        ca = row[8]
        etat = row[15]
        tiers = row[5]
        agence = row[17] if len(row) > 17 else None
        if not ref or not tiers: continue
        ref_str = str(ref).strip()
        etat_str = str(etat) if etat else ""
        if etat_str != "Livrée":
            if not (ref_str in SERVICES_VALIDEE and etat_str == "Validée"):
                continue
        try: c = float(ca) if ca else 0
        except: continue
        if ref_str == "M1051" and c == 0: continue
        
        tiers_str = str(tiers).strip()
        if tiers_str in excluded_tiers: continue
        if COMPTOIR_RE.search(tiers_str): continue
        
        agence_norm = normalize_agence(agence)
        if not agence_norm: continue
        
        client_ca[tiers_str] += c
        client_agence[tiers_str] = agence_norm
        
        # Extract name
        parts = tiers_str.split(" - ", 1)
        if len(parts) == 2:
            client_nom[tiers_str] = parts[1].strip()
        else:
            client_nom[tiers_str] = tiers_str

wb26.close()

# Identifier 20/80 (top 20% des clients = 80% du CA)
total_ca = sum(client_ca.values())
sorted_clients = sorted(client_ca.items(), key=lambda x: -x[1])

cumul = 0
clients_2080 = []
for tiers, ca in sorted_clients:
    cumul += ca
    if cumul >= total_ca * 0.80:
        clients_2080.append(tiers)
        break
    clients_2080.append(tiers)

print(f"  Total clients: {len(sorted_clients)}")
print(f"  Clients 20/80: {len(clients_2080)} ({len(clients_2080)/len(sorted_clients)*100:.1f}%)")
print(f"  CA 20/80: {sum(client_ca[t] for t in clients_2080)/1e6:.0f} M ({sum(client_ca[t] for t in clients_2080)/total_ca*100:.1f}%)")

# Grouper par région
for region in ["Ouest", "Centre", "Littoral"]:
    ws = wb_out.create_sheet(f"20-80 {region}")
    
    # Titre
    ws.cell(1, 1, f"CLIENTS 20/80 — RÉGION {region.upper()}")
    ws.cell(1, 1).font = Font(bold=True, size=14, color=NAVY)
    ws.merge_cells('A1:F1')
    
    headers = ["Rang", "Code client", "Nom client", "Agence", "CA S1 2026 (FCFA)", "CA S1 2026 (M)"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(3, col, h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = thin_border
    
    # Filtrer clients 20/80 par région
    region_clients = []
    for tiers in clients_2080:
        ag = client_agence.get(tiers, "")
        if AGENCE_REGION.get(ag, "") == region:
            region_clients.append((tiers, client_ca[tiers], ag, client_nom.get(tiers, "")))
    
    # Trier par CA décroissant
    region_clients.sort(key=lambda x: -x[1])
    
    row_idx = 4
    for rank, (tiers, ca, ag, nom) in enumerate(region_clients, 1):
        ws.cell(row_idx, 1, rank)
        ws.cell(row_idx, 2, tiers[:20])
        ws.cell(row_idx, 3, nom[:50])
        ws.cell(row_idx, 4, ag)
        ws.cell(row_idx, 5, round(ca, 0))
        ws.cell(row_idx, 6, round(ca/1e6, 1))
        
        for c in range(1, 7):
            ws.cell(row_idx, c).border = thin_border
        
        if row_idx % 2 == 0:
            for c in range(1, 7):
                ws.cell(row_idx, c).fill = BAND_FILL
        
        row_idx += 1
    
    # Total
    ws.cell(row_idx, 1, "").fill = TOTAL_FILL
    ws.cell(row_idx, 3, "TOTAL").font = TOTAL_FONT
    ws.cell(row_idx, 3).fill = TOTAL_FILL
    total_region = sum(ca for _, ca, _, _ in region_clients)
    ws.cell(row_idx, 5, round(total_region, 0)).font = TOTAL_FONT
    ws.cell(row_idx, 5).fill = TOTAL_FILL
    ws.cell(row_idx, 6, round(total_region/1e6, 1)).font = TOTAL_FONT
    ws.cell(row_idx, 6).fill = TOTAL_FILL
    for c in range(1, 7):
        ws.cell(row_idx, c).border = thin_border
    
    # Largeurs
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 12
    
    ws.freeze_panes = "A4"
    print(f"  Feuille '{region}' créée: {len(region_clients)} clients")

# Sauvegarder
wb_out.save(OUT)
import os
print(f"\n✓ Fichier sauvegardé: {OUT}")
print(f"  Taille: {os.path.getsize(OUT):,} octets")
print(f"  Feuilles: {wb_out.sheetnames}")
