"""
Génère un fichier Excel: ventes vs objectifs de chaque variété de concentré
(regroupant tous les formats) de Janvier 2025 à Juin 2026, mois par mois.
"""
import openpyxl
import json
import re
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ===== CONFIG =====
MANUAL_WEIGHTS = {"M1051": 50.0, "CF101": 1.0, "S101": 1.0, "P109": 25.0}

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

# Mapping variété → refs
VARIETES = {
    "BELGO 10% Chair": {"C104", "C1042", "C1043", "C1044"},
    "BELGO 5% Chair": {"C103"},
    "BELGO 10% Ponte": {"C102", "C1022"},
    "BELGO 5% Ponte": {"C101"},
    "BELGO 10% Porc": {"C105", "C1053", "C1054", "C1055"},
    "BELGO Ruminant": {"C108"},
}

# Refs inversé: ref → variété
REF_TO_VARIETY = {}
for v, refs in VARIETES.items():
    for r in refs:
        REF_TO_VARIETY[r] = v

# ===== LIRE 2025 (12 mois) =====
print("Lecture 2025 (12 mois)...")
wb25 = openpyxl.load_workbook('/home/z/my-project/upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx', read_only=True, data_only=True)
ws25 = wb25.active

# vol par variété × mois (2025)
vol_2025 = defaultdict(lambda: defaultdict(float))  # variety → month → vol_t
ca_2025 = defaultdict(lambda: defaultdict(float))

for row in ws25.iter_rows(min_row=2, values_only=True):
    if not row or len(row) < 14: continue
    ref = row[0]
    desc = row[1]
    ca = row[6]
    etat = row[9]
    date_cmd = row[5]
    vol_t_native = row[11]
    
    if not ref: continue
    if str(etat) != "Livrée": continue
    
    ref_str = str(ref).strip()
    if ref_str not in REF_TO_VARIETY: continue
    
    # Mois
    month = None
    if date_cmd:
        if hasattr(date_cmd, 'month'): month = date_cmd.month
        else:
            try:
                s = str(date_cmd)
                parts = s.split("/")
                if len(parts) == 3: month = int(parts[1])
            except: pass
    if month is None: continue
    
    variety = REF_TO_VARIETY[ref_str]
    
    # Volume (déjà en tonnes dans le fichier 2025)
    try: v = float(vol_t_native) if vol_t_native else 0
    except: v = 0
    try: c = float(ca) if ca else 0
    except: c = 0
    
    vol_2025[variety][month] += v
    ca_2025[variety][month] += c

wb25.close()
print("  2025 chargé.")

# ===== LIRE 2026 (Jan-Juin) =====
print("Lecture 2026 (Jan-Juin)...")
wb26 = openpyxl.load_workbook('/home/z/my-project/upload/ventes janv a juin 2026.xlsx', read_only=True, data_only=True)
sheet_to_month = {"Sheet 1": 1, "Feuil1": 2, "Feuil2": 3, "Feuil3": 4, "Feuil4": 5, "Feuil5": 6}

vol_2026 = defaultdict(lambda: defaultdict(float))
ca_2026 = defaultdict(lambda: defaultdict(float))

for sheet_name in wb26.sheetnames:
    ws = wb26[sheet_name]
    if ws.cell(2, 1).value != 'Réf. produit': continue
    month = sheet_to_month.get(sheet_name)
    if month is None: continue
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row or len(row) < 16: continue
        ref = row[0]
        desc = row[1]
        qte = row[2]
        ca = row[8]
        etat = row[15]
        if not ref: continue
        ref_str = str(ref).strip()
        if ref_str not in REF_TO_VARIETY: continue
        etat_str = str(etat) if etat else ""
        if etat_str != "Livrée": continue
        
        variety = REF_TO_VARIETY[ref_str]
        weight = parse_weight_kg(ref_str, desc)
        try: q = float(qte) if qte else 0
        except: q = 0
        try: c = float(ca) if ca else 0
        except: c = 0
        vol_t = q * weight / 1000.0
        
        vol_2026[variety][month] += vol_t
        ca_2026[variety][month] += c

wb26.close()
print("  2026 chargé.")

# ===== LIRE OBJECTIFS 2026 (par variété, par mois) =====
print("Lecture objectifs 2026...")
with open('/home/z/my-project/scripts/objectives_comparison.json') as f:
    obj_data = json.load(f)

# Les objectifs sont par catégorie (CONCENTRES), pas par variété
# On doit estimer les objectifs par variété à partir des objectifs agence×catégorie
# Pour simplifier: on utilise les objectifs globaux CONCENTRES et on les répartit
# selon le poids de chaque variété dans le réalisé S1 2026

# Calculer le poids de chaque variété dans S1 2026
total_conc_s1_2026 = sum(sum(vol_2026[v].values()) for v in VARIETES)
variety_share = {}
for v in VARIETES:
    vol_v = sum(vol_2026[v].values())
    variety_share[v] = vol_v / total_conc_s1_2026 if total_conc_s1_2026 > 0 else 0

print(f"  Poids des variétés dans S1 2026:")
for v in VARIETES:
    print(f"    {v}: {variety_share[v]*100:.1f}%")

# Objectifs CONCENTRES par mois (global)
conc_obj_monthly = {}
for m_str, vol in obj_data['global_objectives'].get('CONCENTRES', {}).items():
    conc_obj_monthly[int(m_str)] = float(vol)

# Répartir par variété
vol_obj = defaultdict(dict)  # variety → month → vol_obj
for v in VARIETES:
    for m in range(1, 13):
        if m in conc_obj_monthly:
            vol_obj[v][m] = conc_obj_monthly[m] * variety_share[v]
        else:
            vol_obj[v][m] = 0

print("  Objectifs répartis.")

# ===== CRÉER FICHIER EXCEL =====
print("\nCréation du fichier Excel...")
OUT = '/home/z/my-project/download/concentres_varietes_mensuel.xlsx'
wb_out = openpyxl.Workbook()

# Styles
NAVY = "1F4E78"
HEADER_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
TOTAL_FILL = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
TOTAL_FONT = Font(bold=True, size=10)
BAND_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
thin_border = Border(
    left=Side(style='thin', color="808080"),
    right=Side(style='thin', color="808080"),
    top=Side(style='thin', color="808080"),
    bottom=Side(style='thin', color="808080")
)

ws = wb_out.active
ws.title = "Concentrés par variété"

# Titre
ws.cell(1, 1, "VENTES VS OBJECTIFS PAR VARIÉTÉ DE CONCENTRÉS — Janvier 2025 à Juin 2026 (mois par mois)")
ws.cell(1, 1).font = Font(bold=True, size=14, color=NAVY)
ws.merge_cells('A1:AC1')

# Préparer les labels de mois
month_labels_2025 = [f"{m:02d}/25" for m in range(1, 13)]
month_labels_2026 = [f"{m:02d}/26" for m in range(1, 7)]
all_months = month_labels_2025 + month_labels_2026  # 18 mois

# Header
headers = ["Variété"]
for ml in all_months:
    headers.append(f"{ml} (t)")
headers.append("Total 2025 (t)")
headers.append("Total S1 2026 (t)")
headers.append("Obj S1 2026 (t)")
headers.append("% S1 2026")

for col, h in enumerate(headers, 1):
    c = ws.cell(3, col, h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = thin_border

# Données par variété
row_idx = 4
varieties_order = ["BELGO 10% Chair", "BELGO 5% Chair", "BELGO 10% Ponte", "BELGO 5% Ponte", "BELGO 10% Porc", "BELGO Ruminant"]

total_2025_all = defaultdict(float)
total_2026_all = defaultdict(float)
total_obj_all = 0

for variety in varieties_order:
    ws.cell(row_idx, 1, variety).font = Font(bold=True)
    col = 2
    
    total_2025 = 0
    total_2026 = 0
    total_obj = 0
    
    # 2025 (12 mois)
    for m in range(1, 13):
        v = vol_2025.get(variety, {}).get(m, 0)
        ws.cell(row_idx, col, round(v, 1))
        total_2025 += v
        total_2025_all[m] += v
        col += 1
    
    # 2026 S1 (6 mois)
    for m in range(1, 7):
        v = vol_2026.get(variety, {}).get(m, 0)
        obj = vol_obj.get(variety, {}).get(m, 0)
        # Afficher "réalisé / objectif"
        ws.cell(row_idx, col, f"{round(v, 1)} / {round(obj, 1)}")
        total_2026 += v
        total_obj += obj
        total_2026_all[m] += v
        col += 1
    
    # Totaux
    ws.cell(row_idx, col, round(total_2025, 1))
    ws.cell(row_idx, col+1, round(total_2026, 1))
    ws.cell(row_idx, col+2, round(total_obj, 1))
    pct = total_2026/total_obj*100 if total_obj > 0 else 0
    ws.cell(row_idx, col+3, f"{pct:.0f}%")
    total_obj_all += total_obj
    
    if row_idx % 2 == 0:
        for c in range(1, col+4):
            if not ws.cell(row_idx, c).fill or ws.cell(row_idx, c).fill.start_color.rgb in [None, "00000000"]:
                ws.cell(row_idx, c).fill = BAND_FILL
    
    for c in range(1, col+4):
        ws.cell(row_idx, c).border = thin_border
    
    row_idx += 1

# Ligne TOTAL
ws.cell(row_idx, 1, "TOTAL").font = TOTAL_FONT
ws.cell(row_idx, 1).fill = TOTAL_FILL
col = 2

total_2025_grand = 0
total_2026_grand = 0

for m in range(1, 13):
    v = total_2025_all[m]
    ws.cell(row_idx, col, round(v, 1))
    ws.cell(row_idx, col).fill = TOTAL_FILL
    ws.cell(row_idx, col).font = TOTAL_FONT
    total_2025_grand += v
    col += 1

for m in range(1, 7):
    v = total_2026_all[m]
    obj = sum(vol_obj.get(variety, {}).get(m, 0) for variety in varieties_order)
    ws.cell(row_idx, col, f"{round(v, 1)} / {round(obj, 1)}")
    ws.cell(row_idx, col).fill = TOTAL_FILL
    ws.cell(row_idx, col).font = TOTAL_FONT
    total_2026_grand += v
    col += 1

ws.cell(row_idx, col, round(total_2025_grand, 1))
ws.cell(row_idx, col+1, round(total_2026_grand, 1))
ws.cell(row_idx, col+2, round(total_obj_all, 1))
pct_total = total_2026_grand/total_obj_all*100 if total_obj_all > 0 else 0
ws.cell(row_idx, col+3, f"{pct_total:.0f}%")

for c in range(1, col+4):
    ws.cell(row_idx, c).fill = TOTAL_FILL
    ws.cell(row_idx, c).font = TOTAL_FONT
    ws.cell(row_idx, c).border = thin_border

# Largeurs colonnes
ws.column_dimensions['A'].width = 22
for c in range(2, col+4):
    ws.column_dimensions[get_column_letter(c)].width = 14

ws.freeze_panes = "B4"

# ===== FEUILLE 2: DÉTAIL CA =====
ws2 = wb_out.create_sheet("CA par variété")

ws2.cell(1, 1, "CHIFFRE D'AFFAIRES PAR VARIÉTÉ DE CONCENTRÉS — Janvier 2025 à Juin 2026 (M FCFA)")
ws2.cell(1, 1).font = Font(bold=True, size=14, color=NAVY)
ws2.merge_cells('A1:AC1')

headers2 = ["Variété"]
for ml in all_months:
    headers2.append(f"{ml} (M)")
headers2.append("Total 2025 (M)")
headers2.append("Total S1 2026 (M)")

for col, h in enumerate(headers2, 1):
    c = ws2.cell(3, col, h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = thin_border

row_idx = 4
for variety in varieties_order:
    ws2.cell(row_idx, 1, variety).font = Font(bold=True)
    col = 2
    total_2025 = 0
    total_2026 = 0
    
    for m in range(1, 13):
        c = ca_2025.get(variety, {}).get(m, 0) / 1e6
        ws2.cell(row_idx, col, round(c, 1))
        total_2025 += c
        col += 1
    
    for m in range(1, 7):
        c = ca_2026.get(variety, {}).get(m, 0) / 1e6
        ws2.cell(row_idx, col, round(c, 1))
        total_2026 += c
        col += 1
    
    ws2.cell(row_idx, col, round(total_2025, 1))
    ws2.cell(row_idx, col+1, round(total_2026, 1))
    
    if row_idx % 2 == 0:
        for c in range(1, col+2):
            if not ws2.cell(row_idx, c).fill or ws2.cell(row_idx, c).fill.start_color.rgb in [None, "00000000"]:
                ws2.cell(row_idx, c).fill = BAND_FILL
    
    for c in range(1, col+2):
        ws2.cell(row_idx, c).border = thin_border
    
    row_idx += 1

ws2.column_dimensions['A'].width = 22
for c in range(2, col+2):
    ws2.column_dimensions[get_column_letter(c)].width = 12
ws2.freeze_panes = "B4"

# Sauvegarder
wb_out.save(OUT)
import os
print(f"\n✓ Fichier sauvegardé: {OUT}")
print(f"  Taille: {os.path.getsize(OUT):,} octets")
print(f"  Feuilles: {wb_out.sheetnames}")
