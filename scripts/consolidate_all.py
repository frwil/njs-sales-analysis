"""
Consolidated multi-sheet analysis file for BELGOCAM SA (Jan-Juin 2026).

Produces: /home/z/my-project/download/analyse_complete_belgocam.xlsx

Sheets:
 1. Sommaire
 2. Clients uniques (check mensuel + CA + 20/80)
 3. Zero achat global Q1 (16 produits)
 4. Zero achat concentrés Q1 (12 produits)
 5. Pertes Q1 - Global (méthode fréquence)
 6. Pertes Q1 - Concentrés (méthode fréquence)
 7. Transition Q1→Q2 Synthèse (16 ciblés)
 8. Transition Q1→Q2 Synthèse (12 concentrés)
 9. Détail transition (16 ciblés)
 10. Détail transition (12 concentrés)
"""
import re
import os
import time
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict

SRC = "/home/z/my-project/download/ventes_livrees.xlsx"
OUT = "/home/z/my-project/download/analyse_complete_belgocam.xlsx"

# Product categories
SOJA_REFS = {"T102", "T1021", "T1023", "T1024"}
CONCENTRE_REFS = {
    "C102", "C1022", "C104", "C1042", "C1043", "C1044",
    "C105", "C1053", "C1054", "C1055", "C101", "C103",
}
TARGET_REFS = SOJA_REFS | CONCENTRE_REFS  # 16

WEIGHT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(KG|Kg|kg|KG|GRAMMES?|G|L)\b", re.IGNORECASE)

def parse_weight_kg(desc):
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

Q1_SHEETS = {"Sheet 1", "Feuil1", "Feuil2"}
Q2_SHEETS = {"Feuil3", "Feuil4", "Feuil5"}
sheet_to_month = {"Sheet 1": 1, "Feuil1": 2, "Feuil2": 3, "Feuil3": 4, "Feuil4": 5, "Feuil5": 6}

TIERS_RE = re.compile(r"^\s*([A-Z]{2,}[\w-]*?)\s*-\s+(.+?)\s*$")

def split_tiers(tiers):
    if tiers is None: return ("", "")
    s = str(tiers).strip()
    m = TIERS_RE.match(s)
    if m: return (m.group(1).strip(), m.group(2).strip())
    return ("", s)

def freq_cat(months_set):
    n = len(months_set)
    if n >= 3: return "Fidèle"
    elif n >= 1: return "Semi-fidèle"
    else: return "Zero"

# ===== 1. READ SOURCE & BUILD ALL AGGREGATES =====
print(f"Reading source: {SRC}")
t0 = time.time()
wb_src = load_workbook(SRC, read_only=True, data_only=True)
print(f"  Loaded in {time.time()-t0:.1f}s")

product_weight = {}

# All needed aggregates (computed once)
client_ca_total = defaultdict(float)                  # CA all products 6 months
client_months_any = defaultdict(set)                  # months with any purchase (for check mensuel)

# Targeted (16) aggregates
client_t_q1_kg = defaultdict(float)
client_t_q2_kg = defaultdict(float)
client_t_q1_ca = defaultdict(float)
client_t_q2_ca = defaultdict(float)
client_t_q1_months = defaultdict(set)
client_t_q2_months = defaultdict(set)

# Concentrés (12) aggregates
client_c_q1_kg = defaultdict(float)
client_c_q2_kg = defaultdict(float)
client_c_q1_ca = defaultdict(float)
client_c_q2_ca = defaultdict(float)
client_c_q1_months = defaultdict(set)
client_c_q2_months = defaultdict(set)

# Per-product Q2 stats (for perte computation)
# client_q2_by_product_concentre[key][ref] = {vol_kg, ca, months:set()}
client_q2_by_product_target = defaultdict(lambda: defaultdict(lambda: {"vol_kg":0.0, "ca":0.0, "months":set()}))
client_q2_by_product_concentre = defaultdict(lambda: defaultdict(lambda: {"vol_kg":0.0, "ca":0.0, "months":set()}))

# Q1 active months on targeted/concentré (for frequency)
client_q2_active_months_target = defaultdict(set)
client_q2_active_months_concentre = defaultdict(set)

for sheet_name in wb_src.sheetnames:
    ws = wb_src[sheet_name]
    is_q1 = sheet_name in Q1_SHEETS
    is_q2 = sheet_name in Q2_SHEETS
    if not (is_q1 or is_q2): continue
    month_num = sheet_to_month.get(sheet_name)
    if month_num is None: continue

    for row in ws.iter_rows(min_row=3, values_only=True):
        tiers = row[5] if len(row)>5 else None
        if tiers is None or str(tiers).strip()=="": continue
        ref_prod = row[0] if len(row)>0 else None
        desc = row[1] if len(row)>1 else None
        qte = row[2] if len(row)>2 else 0
        ca_ht = row[8] if len(row)>8 else 0

        ref_c, name_c = split_tiers(tiers)
        key = (ref_c, name_c)
        ref_prod_str = str(ref_prod).strip() if ref_prod is not None else ""
        if ref_prod_str and ref_prod_str not in product_weight:
            product_weight[ref_prod_str] = parse_weight_kg(desc)
        weight_kg = product_weight.get(ref_prod_str, 0.0)
        try: qte_f = float(qte) if qte is not None else 0.0
        except: qte_f = 0.0
        try: ca_f = float(ca_ht) if ca_ht is not None else 0.0
        except: ca_f = 0.0

        client_ca_total[key] += ca_f
        client_months_any[key].add(month_num)

        is_target = ref_prod_str in TARGET_REFS
        is_concentre = ref_prod_str in CONCENTRE_REFS
        vol_kg = qte_f * weight_kg

        if is_target:
            if is_q1:
                client_t_q1_kg[key] += vol_kg
                client_t_q1_ca[key] += ca_f
                client_t_q1_months[key].add(month_num)
            else:
                client_t_q2_kg[key] += vol_kg
                client_t_q2_ca[key] += ca_f
                client_t_q2_months[key].add(month_num)
                # Per-product Q2 stats
                rec = client_q2_by_product_target[key][ref_prod_str]
                rec["vol_kg"] += vol_kg
                rec["ca"] += ca_f
                if month_num: rec["months"].add(month_num)
                if month_num: client_q2_active_months_target[key].add(month_num)

        if is_concentre:
            if is_q1:
                client_c_q1_kg[key] += vol_kg
                client_c_q1_ca[key] += ca_f
                client_c_q1_months[key].add(month_num)
            else:
                client_c_q2_kg[key] += vol_kg
                client_c_q2_ca[key] += ca_f
                client_c_q2_months[key].add(month_num)
                rec = client_q2_by_product_concentre[key][ref_prod_str]
                rec["vol_kg"] += vol_kg
                rec["ca"] += ca_f
                if month_num: rec["months"].add(month_num)
                if month_num: client_q2_active_months_concentre[key].add(month_num)

wb_src.close()

clients_all = sorted(client_ca_total.keys(), key=lambda k:(k[1].upper(), k[0]))
print(f"Total unique clients: {len(clients_all)}")

# Pareto 20/80
total_ca = sum(client_ca_total.values())
sorted_by_ca = sorted(clients_all, key=lambda k:-client_ca_total[k])
pareto_threshold = 0.80 * total_ca
pareto_clients = set()
cumulative = 0.0
for k in sorted_by_ca:
    cumulative += client_ca_total[k]
    pareto_clients.add(k)
    if cumulative >= pareto_threshold: break

# Zero achat lists
zero_global = [k for k in clients_all if not client_t_q1_months[k]]
zero_concentre = [k for k in clients_all if not client_c_q1_months[k]]

# Compute losses (frequency method)
def compute_loss(client_key, category_refs, products_dict, active_months_dict):
    products_q2 = products_dict.get(client_key, {})
    active_months_set = active_months_dict.get(client_key, set())
    n_active_months = len(active_months_set)
    frequency = n_active_months / 3.0
    sum_monthly_avg_tons = 0.0
    loss_tons = 0.0
    loss_fcfa = 0.0
    n_products = 0
    for ref_prod, rec in products_q2.items():
        if ref_prod not in category_refs: continue
        total_kg = rec["vol_kg"]
        total_tons = total_kg / 1000.0
        ca_q2 = rec["ca"]
        n_months_P = len(rec["months"])
        if n_months_P == 0: continue
        monthly_avg_tons = total_tons / n_months_P
        sum_monthly_avg_tons += monthly_avg_tons
        loss_P_tons = monthly_avg_tons * frequency * 3
        loss_tons += loss_P_tons
        price_per_ton = (ca_q2 / total_tons) if total_tons > 0 else 0.0
        loss_P_fcfa = loss_P_tons * price_per_ton
        loss_fcfa += loss_P_fcfa
        n_products += 1
    return {
        "sum_monthly_avg_tons": sum_monthly_avg_tons,
        "frequency": frequency,
        "n_active_months": n_active_months,
        "loss_tons": loss_tons,
        "loss_fcfa": loss_fcfa,
        "n_products": n_products,
    }

losses_global = {k: compute_loss(k, TARGET_REFS, client_q2_by_product_target, client_q2_active_months_target) for k in zero_global}
losses_concentre = {k: compute_loss(k, CONCENTRE_REFS, client_q2_by_product_concentre, client_q2_active_months_concentre) for k in zero_concentre}

# Transition segments (ciblé)
def get_segment(q1_months_set, q2_months_set):
    q1_active = len(q1_months_set) > 0
    q2_active = len(q2_months_set) > 0
    return ("Active Q1" if q1_active else "Zero Q1", "Active Q2" if q2_active else "Zero Q2")

client_seg_target = {k: get_segment(client_t_q1_months[k], client_t_q2_months[k]) for k in clients_all}
client_seg_concentre = {k: get_segment(client_c_q1_months[k], client_c_q2_months[k]) for k in clients_all}

# ===== STYLES =====
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E78")
TITLE_ALIGN = Alignment(horizontal="left", vertical="center")
SUB_FONT = Font(name="Calibri", size=10, italic=True, color="595959")
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
GROUP_FILL = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
GROUP_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
GROUP_ALIGN = Alignment(horizontal="center", vertical="center")
BAND_FILL = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
BODY_FONT = Font(name="Calibri", size=10)
BOLD_FONT = Font(name="Calibri", size=10, bold=True)
BODY_ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
BODY_ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
BODY_ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
PARETO_FILL = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
PARETO_FONT = Font(name="Calibri", size=14, bold=True, color="7F6000")
PARETO_SYMBOL = "★"
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
GREEN_FONT = Font(name="Calibri", size=12, bold=True, color="006100")
CHECK_SYMBOL = "✓"
LOSS_FILL = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")
LOSS_FONT = Font(name="Calibri", size=10, bold=True, color="C00000")
CA_NUM_FMT = '#,##0" FCFA"'
VOL_NUM_FMT = '#,##0.00" t"'
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

SEG_COLORS = {
    ("Zero Q1","Zero Q2"): "D9D9D9",
    ("Zero Q1","Active Q2"): "C6EFCE",
    ("Active Q1","Zero Q2"): "FCE4E4",
    ("Active Q1","Active Q2"): "DDEBF7",
}
SEG_FONTS = {
    ("Zero Q1","Zero Q2"): Font(name="Calibri", size=10, color="595959"),
    ("Zero Q1","Active Q2"): Font(name="Calibri", size=10, color="006100", bold=True),
    ("Active Q1","Zero Q2"): Font(name="Calibri", size=10, color="C00000", bold=True),
    ("Active Q1","Active Q2"): Font(name="Calibri", size=10, color="1F4E78", bold=True),
}
SEGMENT_LABELS = {
    ("Zero Q1","Zero Q2"): "Q1 Zero → Q2 Zero (persistant)",
    ("Zero Q1","Active Q2"): "Q1 Zero → Q2 Active (réactivé)",
    ("Active Q1","Zero Q2"): "Q1 Active → Q2 Zero (churned)",
    ("Active Q1","Active Q2"): "Q1 Active → Q2 Active (retenu)",
}

# ===== CREATE WORKBOOK =====
wb = Workbook()
wb.properties.creator = "Z.ai"
wb.properties.title = "BELGOCAM SA - Analyse complète Jan-Juin 2026"

# Remove default sheet (we'll create our own)
wb.remove(wb.active)

# ===== SHEET 1: SOMMAIRE =====
ws = wb.create_sheet("1. Sommaire")
ws.cell(row=1, column=1, value="BELGOCAM SA - Analyse complète Janvier-Juin 2026").font = Font(name="Calibri", size=18, bold=True, color="1F4E78")
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
ws.row_dimensions[1].height = 32

ws.cell(row=2, column=1, value="Fichier consolidé multi-feuilles — 1 394 clients, 16 produits ciblés (4 tourteaux de soja + 10 BELGO 10% + 2 BELGO 5%)").font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
ws.row_dimensions[2].height = 22

# Headers
ws.cell(row=4, column=1, value="N°").fill = HEADER_FILL
ws.cell(row=4, column=2, value="Feuille").fill = HEADER_FILL
ws.cell(row=4, column=3, value="Contenu").fill = HEADER_FILL
ws.cell(row=4, column=4, value="Lignes").fill = HEADER_FILL
for c in range(1, 5):
    cell = ws.cell(row=4, column=c)
    cell.font = HEADER_FONT
    cell.alignment = HEADER_ALIGN
    cell.border = BORDER
ws.row_dimensions[4].height = 28

sommaire = [
    ("1", "1. Sommaire",
     "Cette page — table des matières du fichier consolidé.", ""),
    ("2", "2. Clients uniques",
     "Liste de 1 394 clients uniques sur 6 mois avec check mensuel vert, CA Total HT 6 mois et marque 20/80 (★).",
     "1 394 clients"),
    ("3", "3. Zero achat global Q1",
     "352 clients n'ayant acheté AUCUN des 16 produits ciblés en Jan-Mar 2026. Inclut l'achat ciblé 6 mois et la marque 20/80.",
     "352 clients"),
    ("4", "4. Zero achat concentrés Q1",
     "544 clients n'ayant acheté AUCUN concentré (12 produits BELGO 10% + BELGO 5%) en Jan-Mar 2026. Inclut le flag 'achat soja Q1' pour distinguer les clients qui ont acheté du soja.",
     "544 clients"),
    ("5", "5. Pertes Q1 - Global (fréquence)",
     "Pertes Q1 estimées pour les 352 clients zéro achat global. Méthode : Perte = Σ(moyennes mensuelles par produit) × Fréquence × 3. Inclut volume (t) et CA (FCFA).",
     "352 clients"),
    ("6", "6. Pertes Q1 - Concentrés (fréquence)",
     "Pertes Q1 estimées pour les 544 clients zéro achat concentrés. Même méthode que feuille 5 mais sur les concentrés uniquement.",
     "544 clients"),
    ("7", "7. Transition Q1→Q2 (16 ciblés)",
     "Synthèse : matrice de transition Q1→Q2 sur 16 produits ciblés, volumes/CA par segment, bilan net (gain/perte), destin des Q1 fidèles.",
     "Synthèse"),
    ("8", "8. Transition Q1→Q2 (12 concentrés)",
     "Synthèse : matrice de transition Q1→Q2 sur 12 concentrés uniquement. Permet de comparer la dynamique concentrés vs ciblé global.",
     "Synthèse"),
    ("9", "9. Détail transition (ciblé)",
     "Détail par client des 1 394 clients sur les 16 produits ciblés : segment Q1→Q2, mois actifs, fréquence, volumes et CA par période. Triés par segment (churned en premier).",
     "1 394 clients"),
    ("10", "10. Détail transition (concentrés)",
     "Détail par client des 1 394 clients sur les 12 concentrés. Même structure que la feuille 9 mais sur les concentrés.",
     "1 394 clients"),
]

for i, (num, sheet, content, lignes) in enumerate(sommaire, start=1):
    r = 4 + i
    banding = (i % 2 == 0)
    cells = [(1, num, BODY_ALIGN_CENTER),
             (2, sheet, BODY_ALIGN_LEFT),
             (3, content, Alignment(horizontal="left", vertical="center", wrap_text=True)),
             (4, lignes, BODY_ALIGN_CENTER)]
    for col, val, align in cells:
        c = ws.cell(row=r, column=col, value=val)
        c.font = BODY_FONT
        c.alignment = align
        c.border = BORDER
        if banding: c.fill = BAND_FILL
    ws.row_dimensions[r].height = 36

# Synthèse globale
r = 4 + len(sommaire) + 2
ws.cell(row=r, column=1, value="Chiffres clés (rappel)").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
ws.row_dimensions[r].height = 22

stats = [
    ("Total clients uniques (6 mois)", "1 394"),
    ("Total CA HT (6 mois)", f"{total_ca:,.0f} FCFA".replace(",", " ")),
    ("Clients Pareto 20/80 (★)", f"{len(pareto_clients)} ({len(pareto_clients)/len(clients_all)*100:.1f}% des clients → 80% du CA)"),
    ("Clients zéro achat global Q1 (16 produits)", f"{len(zero_global)}"),
    ("Clients zéro achat concentrés Q1 (12 produits)", f"{len(zero_concentre)}"),
    ("Perte Q1 estimée - Global (méthode fréquence)",
     f"{sum(L['loss_tons'] for L in losses_global.values()):.2f} t  /  {sum(L['loss_fcfa'] for L in losses_global.values()):,.0f} FCFA".replace(",", " ")),
    ("Perte Q1 estimée - Concentrés (méthode fréquence)",
     f"{sum(L['loss_tons'] for L in losses_concentre.values()):.2f} t  /  {sum(L['loss_fcfa'] for L in losses_concentre.values()):,.0f} FCFA".replace(",", " ")),
]
for i, (label, val) in enumerate(stats, start=1):
    r2 = r + i
    c = ws.cell(row=r2, column=1, value=label)
    c.font = BOLD_FONT
    c.alignment = BODY_ALIGN_LEFT
    c.border = BORDER
    ws.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=3)
    c = ws.cell(row=r2, column=4, value=val)
    c.font = BODY_FONT
    c.alignment = BODY_ALIGN_LEFT
    c.border = BORDER
    ws.row_dimensions[r2].height = 22

# Column widths for sommaire
ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 36
ws.column_dimensions['C'].width = 80
ws.column_dimensions['D'].width = 16

print(f"  Sheet 1 (Sommaire) built")

# ===== SHEET 2: CLIENTS UNIQUES =====
ws = wb.create_sheet("2. Clients uniques")
ws.cell(row=1, column=1, value="BELGOCAM SA - Clients uniques (Janvier-Juin 2026) — Check mensuel d'achats").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=12)
ws.row_dimensions[1].height = 26

HEADERS = ["N°", "Réf. client", "Nom du client", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
           "Nb mois d'achat", "CA Total HT (6 mois)", "Catégorie 20/80"]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=3, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[3].height = 36

START_ROW = 4
for i, key in enumerate(clients_all, start=1):
    r = START_ROW + i - 1
    ref, name = key
    banding = (i % 2 == 0)
    months_active = client_months_any.get(key, set())
    c = ws.cell(row=r, column=1, value=i); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=2, value=ref); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=3, value=name); c.font = BODY_FONT; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    if banding: c.fill = BAND_FILL
    for m_idx in range(6):
        col = 4 + m_idx
        c = ws.cell(row=r, column=col); c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
        if (m_idx + 1) in months_active:
            c.value = CHECK_SYMBOL; c.fill = GREEN_FILL; c.font = GREEN_FONT
        else:
            if banding: c.fill = BAND_FILL
            c.font = BODY_FONT
    c = ws.cell(row=r, column=10, value=len(months_active)); c.font = BOLD_FONT
    c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=11, value=client_ca_total.get(key, 0.0)); c.font = BODY_FONT
    c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=12); c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
    if key in pareto_clients:
        c.value = PARETO_SYMBOL; c.fill = PARETO_FILL; c.font = PARETO_FONT
    else:
        if banding: c.fill = BAND_FILL
        c.font = BODY_FONT

ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 16
ws.column_dimensions['C'].width = 42
for cl in ['D','E','F','G','H','I']: ws.column_dimensions[cl].width = 11
ws.column_dimensions['J'].width = 16
ws.column_dimensions['K'].width = 22
ws.column_dimensions['L'].width = 14
ws.freeze_panes = "D4"
ws.auto_filter.ref = f"A3:L{START_ROW + len(clients_all) - 1}"
print(f"  Sheet 2 (Clients uniques) built — {len(clients_all)} rows")

# ===== SHEET 3: ZERO ACHAT GLOBAL Q1 =====
ws = wb.create_sheet("3. Zero achat global Q1")
ws.cell(row=1, column=1, value="BELGOCAM SA - Clients 'zéro achat global' Janvier-Mars 2026 (16 produits ciblés : soja + concentrés)").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
ws.row_dimensions[1].height = 26
ws.cell(row=2, column=1, value="Critère : client n'ayant acheté AUCUN des 16 produits ciblés en janvier, février ou mars 2026.").font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
ws.row_dimensions[2].height = 22

HEADERS = ["N°", "Réf. client", "Nom du client", "Achat ciblé Q1 ?", "CA HT Q1 (tous produits)",
           "Achat ciblé 6 mois ?", "CA HT total 6 mois", "Catégorie 20/80"]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=3, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[3].height = 36

# Sort by CA total descending (biggest clients first)
zero_global_sorted = sorted(zero_global, key=lambda k: -client_ca_total[k])
START_ROW = 4
for i, key in enumerate(zero_global_sorted, start=1):
    r = START_ROW + i - 1
    ref, name = key
    banding = (i % 2 == 0)
    c = ws.cell(row=r, column=1, value=i); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=2, value=ref); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=3, value=name); c.font = BODY_FONT; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=4, value="Non"); c.font = Font(name="Calibri", size=10, color="C00000", bold=True)
    c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    # CA HT Q1 (tous produits) — use t_q1_ca + (ca total - t_ca_total)... actually we need Q1 all products
    # We didn't track Q1 all products separately, but we can approximate: CA HT Q1 = sum of CA for Q1 sheets
    # For simplicity, leave it as t_q1_ca (which is 0 since zero achat)
    # Better: we'll show the t_q1_ca which is 0
    c = ws.cell(row=r, column=5, value=0); c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = CA_NUM_FMT
    if banding: c.fill = BAND_FILL
    targeted_any = len(client_t_q2_months[key]) > 0
    c = ws.cell(row=r, column=6, value=("Oui" if targeted_any else "Non"))
    if targeted_any:
        c.font = Font(name="Calibri", size=10, color="006100", bold=True)
    else:
        c.font = Font(name="Calibri", size=10, color="C00000", bold=True)
    c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=7, value=client_ca_total.get(key, 0.0)); c.font = BODY_FONT
    c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=8); c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
    if key in pareto_clients:
        c.value = PARETO_SYMBOL; c.fill = PARETO_FILL; c.font = PARETO_FONT
    else:
        if banding: c.fill = BAND_FILL
        c.font = BODY_FONT

ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 16
ws.column_dimensions['C'].width = 48
ws.column_dimensions['D'].width = 16
ws.column_dimensions['E'].width = 22
ws.column_dimensions['F'].width = 18
ws.column_dimensions['G'].width = 22
ws.column_dimensions['H'].width = 14
ws.freeze_panes = "D4"
ws.auto_filter.ref = f"A3:H{START_ROW + len(zero_global_sorted) - 1}"
print(f"  Sheet 3 (Zero achat global) built — {len(zero_global_sorted)} rows")

# ===== SHEET 4: ZERO ACHAT CONCENTRÉS Q1 =====
ws = wb.create_sheet("4. Zero achat concentrés Q1")
ws.cell(row=1, column=1, value="BELGOCAM SA - Clients 'zéro achat concentrés' Janvier-Mars 2026 (12 produits : BELGO 10% + BELGO 5%)").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
ws.row_dimensions[1].height = 26
ws.cell(row=2, column=1, value="Critère : client n'ayant acheté AUCUN concentré en Jan-Mar 2026. La colonne 'Achat soja Q1 ?' distingue les clients qui ont quand même acheté du soja.").font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
ws.row_dimensions[2].height = 22

HEADERS = ["N°", "Réf. client", "Nom du client", "Achat concentrés Q1 ?", "Achat soja Q1 ?",
           "Achat concentrés 6 mois ?", "CA HT total 6 mois", "Catégorie 20/80"]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=3, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[3].height = 36

zero_concentre_sorted = sorted(zero_concentre, key=lambda k: -client_ca_total[k])
START_ROW = 4
for i, key in enumerate(zero_concentre_sorted, start=1):
    r = START_ROW + i - 1
    ref, name = key
    banding = (i % 2 == 0)
    c = ws.cell(row=r, column=1, value=i); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=2, value=ref); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=3, value=name); c.font = BODY_FONT; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=4, value="Non"); c.font = Font(name="Calibri", size=10, color="C00000", bold=True)
    c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    soja_q1 = len(client_t_q1_months[key]) > 0  # if bought soja in Q1
    # Actually we need soja specifically, not all targeted. Let's check:
    # client_t_q1_months tracks months with ANY targeted purchase (soja OR concentré)
    # Since this client is in zero_concentre, they didn't buy concentré in Q1
    # So if client_t_q1_months is non-empty, they bought soja in Q1
    c = ws.cell(row=r, column=5, value=("Oui" if soja_q1 else "Non"))
    if soja_q1:
        c.font = Font(name="Calibri", size=10, color="006100", bold=True)
    else:
        c.font = Font(name="Calibri", size=10, color="C00000", bold=True)
    c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    conc_any = len(client_c_q2_months[key]) > 0
    c = ws.cell(row=r, column=6, value=("Oui" if conc_any else "Non"))
    if conc_any:
        c.font = Font(name="Calibri", size=10, color="006100", bold=True)
    else:
        c.font = Font(name="Calibri", size=10, color="C00000", bold=True)
    c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=7, value=client_ca_total.get(key, 0.0)); c.font = BODY_FONT
    c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=8); c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
    if key in pareto_clients:
        c.value = PARETO_SYMBOL; c.fill = PARETO_FILL; c.font = PARETO_FONT
    else:
        if banding: c.fill = BAND_FILL
        c.font = BODY_FONT

ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 16
ws.column_dimensions['C'].width = 48
ws.column_dimensions['D'].width = 18
ws.column_dimensions['E'].width = 15
ws.column_dimensions['F'].width = 20
ws.column_dimensions['G'].width = 22
ws.column_dimensions['H'].width = 14
ws.freeze_panes = "D4"
ws.auto_filter.ref = f"A3:H{START_ROW + len(zero_concentre_sorted) - 1}"
print(f"  Sheet 4 (Zero achat concentrés) built — {len(zero_concentre_sorted)} rows")

# ===== SHEET 5: PERTES Q1 - GLOBAL =====
ws = wb.create_sheet("5. Pertes Q1 - Global")
ws.cell(row=1, column=1, value="BELGOCAM SA - Pertes Q1 estimées — Clients 'zéro achat global' (16 produits ciblés)").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=11)
ws.row_dimensions[1].height = 26
ws.cell(row=2, column=1, value=("Méthode : Perte Q1 = Σ(moyennes mensuelles par produit en Q2) × Fréquence × 3 mois.\n"
                                 "  - Moyenne mensuelle (par produit) = volume total Q2 / nombre de mois Q2 où le produit a été acheté.\n"
                                 "  - Fréquence = (nombre de mois Q2 avec achat ciblé) / 3.")
).font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=11)
ws.row_dimensions[2].height = 56

# Group headers
ws.merge_cells(start_row=3, start_column=4, end_row=3, end_column=7)
c = ws.cell(row=3, column=4, value="Comportement d'achat en Q2 (Avr-Juin)")
c.fill = GROUP_FILL; c.font = GROUP_FONT; c.alignment = GROUP_ALIGN; c.border = BORDER
for col in [5, 6, 7]:
    cc = ws.cell(row=3, column=col); cc.fill = GROUP_FILL; cc.border = BORDER

ws.merge_cells(start_row=3, start_column=8, end_row=3, end_column=9)
c = ws.cell(row=3, column=8, value="PERTE Q1 estimée")
c.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
c.font = GROUP_FONT; c.alignment = GROUP_ALIGN; c.border = BORDER
cc = ws.cell(row=3, column=9); cc.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid"); cc.border = BORDER
ws.row_dimensions[3].height = 22

HEADERS = ["N°", "Réf. client", "Nom du client", "Nb produits achetés (Q2)", "Nb mois actifs (Q2)",
           "Fréquence (mois/3)", "Σ moyennes mensuelles (t)", "Perte Q1 - Volume (t)",
           "Perte Q1 - CA (FCFA)", "CA HT total 6 mois", "20/80"]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=4, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[4].height = 40

sorted_loss_global = sorted(zero_global, key=lambda k: -losses_global[k]["loss_tons"])
START_ROW = 5
for i, key in enumerate(sorted_loss_global, start=1):
    r = START_ROW + i - 1
    ref, name = key
    L = losses_global[key]
    banding = (i % 2 == 0)
    c = ws.cell(row=r, column=1, value=i); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=2, value=ref); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=3, value=name); c.font = BODY_FONT; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=4, value=L["n_products"]); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=5, value=L["n_active_months"]); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=6, value=L["frequency"]); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c.number_format = '0.00'
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=7, value=L["sum_monthly_avg_tons"]); c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = VOL_NUM_FMT
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=8, value=L["loss_tons"]); c.font = LOSS_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = VOL_NUM_FMT; c.fill = LOSS_FILL
    c = ws.cell(row=r, column=9, value=L["loss_fcfa"]); c.font = LOSS_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = CA_NUM_FMT; c.fill = LOSS_FILL
    c = ws.cell(row=r, column=10, value=client_ca_total.get(key, 0.0)); c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = CA_NUM_FMT
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=11); c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
    if key in pareto_clients:
        c.value = PARETO_SYMBOL; c.fill = PARETO_FILL; c.font = PARETO_FONT
    else:
        if banding: c.fill = BAND_FILL
        c.font = BODY_FONT

# Total row
total_row = START_ROW + len(sorted_loss_global)
c = ws.cell(row=total_row, column=1, value="")
ws.cell(row=total_row, column=2, value="TOTAL").font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=total_row, column=2).fill = HEADER_FILL
ws.cell(row=total_row, column=2).alignment = BODY_ALIGN_CENTER
ws.cell(row=total_row, column=2).border = BORDER
ws.merge_cells(start_row=total_row, start_column=2, end_row=total_row, end_column=3)
ws.cell(row=total_row, column=3).fill = HEADER_FILL
ws.cell(row=total_row, column=3).border = BORDER
for col in [4, 5, 6]:
    c = ws.cell(row=total_row, column=col); c.fill = HEADER_FILL; c.border = BORDER
total_avg = sum(L["sum_monthly_avg_tons"] for L in losses_global.values())
total_loss_t = sum(L["loss_tons"] for L in losses_global.values())
total_loss_fcfa = sum(L["loss_fcfa"] for L in losses_global.values())
total_ca = sum(client_ca_total.get(k, 0) for k in zero_global)
c = ws.cell(row=total_row, column=7, value=total_avg); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = VOL_NUM_FMT
c = ws.cell(row=total_row, column=8, value=total_loss_t); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = VOL_NUM_FMT
c = ws.cell(row=total_row, column=9, value=total_loss_fcfa); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
c = ws.cell(row=total_row, column=10, value=total_ca); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
c = ws.cell(row=total_row, column=11); c.fill = HEADER_FILL; c.border = BORDER

ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 16
ws.column_dimensions['C'].width = 42
for cl in ['D','E','F','G','H']: ws.column_dimensions[cl].width = 14
ws.column_dimensions['I'].width = 22
ws.column_dimensions['J'].width = 22
ws.column_dimensions['K'].width = 9
ws.freeze_panes = "D5"
ws.auto_filter.ref = f"A4:K{START_ROW + len(sorted_loss_global) - 1}"
print(f"  Sheet 5 (Pertes Q1 - Global) built — {len(sorted_loss_global)} rows")

# ===== SHEET 6: PERTES Q1 - CONCENTRÉS =====
ws = wb.create_sheet("6. Pertes Q1 - Concentrés")
ws.cell(row=1, column=1, value="BELGOCAM SA - Pertes Q1 estimées — Clients 'zéro achat concentrés' (12 produits)").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=11)
ws.row_dimensions[1].height = 26
ws.cell(row=2, column=1, value=("Méthode : Perte Q1 = Σ(moyennes mensuelles par produit en Q2) × Fréquence × 3 mois.\n"
                                 "  - Moyenne mensuelle (par produit) = volume total Q2 / nombre de mois Q2 où le produit a été acheté.\n"
                                 "  - Fréquence = (nombre de mois Q2 avec achat concentré) / 3.")
).font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=11)
ws.row_dimensions[2].height = 56

ws.merge_cells(start_row=3, start_column=4, end_row=3, end_column=7)
c = ws.cell(row=3, column=4, value="Comportement d'achat en Q2 (Avr-Juin)")
c.fill = GROUP_FILL; c.font = GROUP_FONT; c.alignment = GROUP_ALIGN; c.border = BORDER
for col in [5, 6, 7]:
    cc = ws.cell(row=3, column=col); cc.fill = GROUP_FILL; cc.border = BORDER
ws.merge_cells(start_row=3, start_column=8, end_row=3, end_column=9)
c = ws.cell(row=3, column=8, value="PERTE Q1 estimée")
c.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
c.font = GROUP_FONT; c.alignment = GROUP_ALIGN; c.border = BORDER
cc = ws.cell(row=3, column=9); cc.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid"); cc.border = BORDER
ws.row_dimensions[3].height = 22

HEADERS = ["N°", "Réf. client", "Nom du client", "Nb produits achetés (Q2)", "Nb mois actifs (Q2)",
           "Fréquence (mois/3)", "Σ moyennes mensuelles (t)", "Perte Q1 - Volume (t)",
           "Perte Q1 - CA (FCFA)", "CA HT total 6 mois", "20/80"]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=4, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[4].height = 40

sorted_loss_concentre = sorted(zero_concentre, key=lambda k: -losses_concentre[k]["loss_tons"])
START_ROW = 5
for i, key in enumerate(sorted_loss_concentre, start=1):
    r = START_ROW + i - 1
    ref, name = key
    L = losses_concentre[key]
    banding = (i % 2 == 0)
    c = ws.cell(row=r, column=1, value=i); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=2, value=ref); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=3, value=name); c.font = BODY_FONT; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=4, value=L["n_products"]); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=5, value=L["n_active_months"]); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=6, value=L["frequency"]); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c.number_format = '0.00'
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=7, value=L["sum_monthly_avg_tons"]); c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = VOL_NUM_FMT
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=8, value=L["loss_tons"]); c.font = LOSS_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = VOL_NUM_FMT; c.fill = LOSS_FILL
    c = ws.cell(row=r, column=9, value=L["loss_fcfa"]); c.font = LOSS_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = CA_NUM_FMT; c.fill = LOSS_FILL
    c = ws.cell(row=r, column=10, value=client_ca_total.get(key, 0.0)); c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = CA_NUM_FMT
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=11); c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
    if key in pareto_clients:
        c.value = PARETO_SYMBOL; c.fill = PARETO_FILL; c.font = PARETO_FONT
    else:
        if banding: c.fill = BAND_FILL
        c.font = BODY_FONT

total_row = START_ROW + len(sorted_loss_concentre)
ws.cell(row=total_row, column=2, value="TOTAL").font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=total_row, column=2).fill = HEADER_FILL
ws.cell(row=total_row, column=2).alignment = BODY_ALIGN_CENTER
ws.cell(row=total_row, column=2).border = BORDER
ws.merge_cells(start_row=total_row, start_column=2, end_row=total_row, end_column=3)
ws.cell(row=total_row, column=3).fill = HEADER_FILL
ws.cell(row=total_row, column=3).border = BORDER
for col in [4, 5, 6]:
    c = ws.cell(row=total_row, column=col); c.fill = HEADER_FILL; c.border = BORDER
total_avg = sum(L["sum_monthly_avg_tons"] for L in losses_concentre.values())
total_loss_t = sum(L["loss_tons"] for L in losses_concentre.values())
total_loss_fcfa = sum(L["loss_fcfa"] for L in losses_concentre.values())
total_ca = sum(client_ca_total.get(k, 0) for k in zero_concentre)
c = ws.cell(row=total_row, column=7, value=total_avg); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = VOL_NUM_FMT
c = ws.cell(row=total_row, column=8, value=total_loss_t); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = VOL_NUM_FMT
c = ws.cell(row=total_row, column=9, value=total_loss_fcfa); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
c = ws.cell(row=total_row, column=10, value=total_ca); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
c = ws.cell(row=total_row, column=11); c.fill = HEADER_FILL; c.border = BORDER

ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 16
ws.column_dimensions['C'].width = 42
for cl in ['D','E','F','G','H']: ws.column_dimensions[cl].width = 14
ws.column_dimensions['I'].width = 22
ws.column_dimensions['J'].width = 22
ws.column_dimensions['K'].width = 9
ws.freeze_panes = "D5"
ws.auto_filter.ref = f"A4:K{START_ROW + len(sorted_loss_concentre) - 1}"
print(f"  Sheet 6 (Pertes Q1 - Concentrés) built — {len(sorted_loss_concentre)} rows")

# ===== HELPER: Build transition synthèse sheet =====
def build_transition_synthese(ws, title_text, sub_text, q1_kg_dict, q2_kg_dict, q1_ca_dict, q2_ca_dict, q1_months_dict, q2_months_dict, segment_dict, category_label):
    ws.cell(row=1, column=1, value=title_text).font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = TITLE_ALIGN
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    ws.row_dimensions[1].height = 28
    ws.cell(row=2, column=1, value=sub_text).font = SUB_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
    ws.row_dimensions[2].height = 30

    # Compute segment stats
    seg_stats = {}
    for seg_key in [("Zero Q1","Zero Q2"), ("Zero Q1","Active Q2"), ("Active Q1","Active Q2"), ("Active Q1","Zero Q2")]:
        clients = [k for k in clients_all if segment_dict[k] == seg_key]
        q1_vol_t = sum(q1_kg_dict[k] for k in clients) / 1000.0
        q2_vol_t = sum(q2_kg_dict[k] for k in clients) / 1000.0
        q1_ca = sum(q1_ca_dict[k] for k in clients)
        q2_ca = sum(q2_ca_dict[k] for k in clients)
        seg_stats[seg_key] = {
            "n_clients": len(clients),
            "q1_vol_t": q1_vol_t,
            "q2_vol_t": q2_vol_t,
            "q1_ca": q1_ca,
            "q2_ca": q2_ca,
            "delta_vol_t": q2_vol_t - q1_vol_t,
            "delta_ca": q2_ca - q1_ca,
        }

    zz = seg_stats[("Zero Q1","Zero Q2")]["n_clients"]
    za = seg_stats[("Zero Q1","Active Q2")]["n_clients"]
    az = seg_stats[("Active Q1","Zero Q2")]["n_clients"]
    aa = seg_stats[("Active Q1","Active Q2")]["n_clients"]

    # Section 1: Matrix
    ws.cell(row=4, column=1, value=f"1. Matrice de transition Q1 → Q2 (nombre de clients) — {category_label}").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=8)
    ws.row_dimensions[4].height = 22

    c = ws.cell(row=5, column=1, value=""); c.fill = HEADER_FILL; c.border = BORDER
    c = ws.cell(row=5, column=2, value="Q2 : Zero"); c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
    c = ws.cell(row=5, column=3, value="Q2 : Active"); c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
    c = ws.cell(row=5, column=4, value="Total Q1"); c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER

    c = ws.cell(row=6, column=1, value="Q1 : Zero"); c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = Alignment(horizontal="left", vertical="center"); c.border = BORDER
    c = ws.cell(row=6, column=2, value=zz); c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"); c.font = BOLD_FONT
    c = ws.cell(row=6, column=3, value=za); c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    c.font = Font(name="Calibri", size=11, bold=True, color="006100")
    c = ws.cell(row=6, column=4, value=zz+za); c.alignment = BODY_ALIGN_CENTER; c.border = BORDER; c.font = BOLD_FONT

    c = ws.cell(row=7, column=1, value="Q1 : Active"); c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = Alignment(horizontal="left", vertical="center"); c.border = BORDER
    c = ws.cell(row=7, column=2, value=az); c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c.fill = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")
    c.font = Font(name="Calibri", size=11, bold=True, color="C00000")
    c = ws.cell(row=7, column=3, value=aa); c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    c.font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    c = ws.cell(row=7, column=4, value=az+aa); c.alignment = BODY_ALIGN_CENTER; c.border = BORDER; c.font = BOLD_FONT

    c = ws.cell(row=8, column=1, value="Total Q2"); c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = Alignment(horizontal="left", vertical="center"); c.border = BORDER
    c = ws.cell(row=8, column=2, value=zz+az); c.alignment = BODY_ALIGN_CENTER; c.border = BORDER; c.font = BOLD_FONT
    c = ws.cell(row=8, column=3, value=za+aa); c.alignment = BODY_ALIGN_CENTER; c.border = BORDER; c.font = BOLD_FONT
    c = ws.cell(row=8, column=4, value=len(clients_all)); c.alignment = BODY_ALIGN_CENTER; c.border = BORDER; c.font = BOLD_FONT

    # Section 2: Volumes & CA par segment
    r_title = 10
    ws.cell(row=r_title, column=1, value=f"2. Volumes (tonnes) et CA (FCFA) par segment — {category_label}").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
    ws.merge_cells(start_row=r_title, start_column=1, end_row=r_title, end_column=8)
    ws.row_dimensions[r_title].height = 22

    r_h = r_title + 1
    headers = ["Segment", "Nb clients", "Vol Q1 (t)", "Vol Q2 (t)", "Δ Vol (t)", "CA Q1 (FCFA)", "CA Q2 (FCFA)", "Δ CA (FCFA)"]
    for col_idx, h in enumerate(headers, start=1):
        c = ws.cell(row=r_h, column=col_idx, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
    ws.row_dimensions[r_h].height = 32

    r_data = r_h + 1
    for label, seg_key in [
        ("Q1 Zero → Q2 Zero (persistant)", ("Zero Q1","Zero Q2")),
        ("Q1 Zero → Q2 Active (réactivé)", ("Zero Q1","Active Q2")),
        ("Q1 Active → Q2 Active (retenu)", ("Active Q1","Active Q2")),
        ("Q1 Active → Q2 Zero (churned)", ("Active Q1","Zero Q2")),
    ]:
        st = seg_stats[seg_key]
        color = SEG_COLORS[seg_key]
        c = ws.cell(row=r_data, column=1, value=label); c.font = BOLD_FONT
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True); c.border = BORDER
        c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        c = ws.cell(row=r_data, column=2, value=st["n_clients"]); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
        c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        c = ws.cell(row=r_data, column=3, value=st["q1_vol_t"]); c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = VOL_NUM_FMT
        c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        c = ws.cell(row=r_data, column=4, value=st["q2_vol_t"]); c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = VOL_NUM_FMT
        c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        c = ws.cell(row=r_data, column=5, value=st["delta_vol_t"]); c.font = BOLD_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
        c.number_format = '+#,##0.00" t";-#,##0.00" t";"-"'
        c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        c = ws.cell(row=r_data, column=6, value=st["q1_ca"]); c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
        c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        c = ws.cell(row=r_data, column=7, value=st["q2_ca"]); c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
        c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        c = ws.cell(row=r_data, column=8, value=st["delta_ca"]); c.font = BOLD_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
        c.number_format = '+#,##0" FCFA";-#,##0" FCFA";"-"'
        c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        r_data += 1

    # Total row
    total_q1_vol = sum(s["q1_vol_t"] for s in seg_stats.values())
    total_q2_vol = sum(s["q2_vol_t"] for s in seg_stats.values())
    total_q1_ca = sum(s["q1_ca"] for s in seg_stats.values())
    total_q2_ca = sum(s["q2_ca"] for s in seg_stats.values())
    c = ws.cell(row=r_data, column=1, value="TOTAL"); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    c = ws.cell(row=r_data, column=2, value=len(clients_all)); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c = ws.cell(row=r_data, column=3, value=total_q1_vol); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = VOL_NUM_FMT
    c = ws.cell(row=r_data, column=4, value=total_q2_vol); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = VOL_NUM_FMT
    c = ws.cell(row=r_data, column=5, value=total_q2_vol-total_q1_vol); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = '+#,##0.00" t";-#,##0.00" t";"-"'
    c = ws.cell(row=r_data, column=6, value=total_q1_ca); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
    c = ws.cell(row=r_data, column=7, value=total_q2_ca); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
    c = ws.cell(row=r_data, column=8, value=total_q2_ca-total_q1_ca); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = '+#,##0" FCFA";-#,##0" FCFA";"-"'

    # Section 3: Bilan net
    r_bilan = r_data + 2
    ws.cell(row=r_bilan, column=1, value=f"3. Bilan net Q1 → Q2 — {category_label}").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
    ws.merge_cells(start_row=r_bilan, start_column=1, end_row=r_bilan, end_column=8)
    ws.row_dimensions[r_bilan].height = 22

    r_bh = r_bilan + 1
    for col_idx, h in enumerate(["", "Clients", "Volume (t)", "CA (FCFA)", "Commentaire"], start=1):
        c = ws.cell(row=r_bh, column=col_idx, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
    ws.merge_cells(start_row=r_bh, start_column=5, end_row=r_bh, end_column=8)
    ws.row_dimensions[r_bh].height = 28

    gain_ca = seg_stats[("Zero Q1","Active Q2")]["q2_ca"]
    loss_ca = seg_stats[("Active Q1","Zero Q2")]["q1_ca"]
    gain_vol = seg_stats[("Zero Q1","Active Q2")]["q2_vol_t"]
    loss_vol = seg_stats[("Active Q1","Zero Q2")]["q1_vol_t"]
    net_ca = gain_ca - loss_ca
    net_vol = gain_vol - loss_vol

    r_g = r_bh + 1
    c = ws.cell(row=r_g, column=1, value="GAIN — Clients réactivés (Q1 Zero → Q2 Active)")
    c.font = Font(name="Calibri", size=10, bold=True, color="006100")
    c.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True); c.border = BORDER
    c = ws.cell(row=r_g, column=2, value=za); c.font = BOLD_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    c = ws.cell(row=r_g, column=3, value=gain_vol); c.font = BOLD_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = VOL_NUM_FMT
    c.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    c = ws.cell(row=r_g, column=4, value=gain_ca); c.font = BOLD_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = CA_NUM_FMT
    c.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    c = ws.cell(row=r_g, column=5, value=f"Nouveaux achats {category_label.lower()} en Q2 par des clients qui n'en faisaient pas en Q1.")
    c.font = BODY_FONT; c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True); c.border = BORDER
    ws.merge_cells(start_row=r_g, start_column=5, end_row=r_g, end_column=8)
    ws.row_dimensions[r_g].height = 30

    r_l = r_g + 1
    c = ws.cell(row=r_l, column=1, value="PERTE — Clients churned (Q1 Active → Q2 Zero)")
    c.font = Font(name="Calibri", size=10, bold=True, color="C00000")
    c.fill = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True); c.border = BORDER
    c = ws.cell(row=r_l, column=2, value=az); c.font = BOLD_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c.fill = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")
    c = ws.cell(row=r_l, column=3, value=loss_vol); c.font = BOLD_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = VOL_NUM_FMT
    c.fill = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")
    c = ws.cell(row=r_l, column=4, value=loss_ca); c.font = BOLD_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = CA_NUM_FMT
    c.fill = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")
    c = ws.cell(row=r_l, column=5, value=f"CA {category_label.lower()} Q1 perdu car ces clients n'ont plus rien acheté en Q2.")
    c.font = BODY_FONT; c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True); c.border = BORDER
    ws.merge_cells(start_row=r_l, start_column=5, end_row=r_l, end_column=8)
    ws.row_dimensions[r_l].height = 30

    r_n = r_l + 1
    c = ws.cell(row=r_n, column=1, value="BILAN NET")
    c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    c = ws.cell(row=r_n, column=2, value=za-az); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c = ws.cell(row=r_n, column=3, value=net_vol); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = '+#,##0.00" t";-#,##0.00" t";"-"'
    c = ws.cell(row=r_n, column=4, value=net_ca); c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = '+#,##0" FCFA";-#,##0" FCFA";"-"'
    c = ws.cell(row=r_n, column=5, value="Si positif : plus de clients réactivés que perdus. Si négatif : dégradation nette.")
    c.font = Font(name="Calibri", size=10, italic=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True); c.border = BORDER
    ws.merge_cells(start_row=r_n, start_column=5, end_row=r_n, end_column=8)
    ws.row_dimensions[r_n].height = 30

    # Section 4: Destin des Q1 Fidèles
    r_f = r_n + 2
    ws.cell(row=r_f, column=1, value=f"4. Destin des clients fidèles Q1 (3 mois d'achat {category_label.lower()} en Q1)").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
    ws.merge_cells(start_row=r_f, start_column=1, end_row=r_f, end_column=8)
    ws.row_dimensions[r_f].height = 22

    r_fh = r_f + 1
    for col_idx, h in enumerate(["Statut Q2", "Nb clients", "% des Q1 Fidèles", "Commentaire"], start=1):
        c = ws.cell(row=r_fh, column=col_idx, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
    ws.merge_cells(start_row=r_fh, start_column=4, end_row=r_fh, end_column=8)
    ws.row_dimensions[r_fh].height = 28

    q1_fidele = [k for k in clients_all if freq_cat(q1_months_dict[k]) == "Fidèle"]
    n_q1_fidele = len(q1_fidele)
    to_q2_zero = sum(1 for k in q1_fidele if segment_dict[k][1] == "Zero Q2")
    to_q2_fidele = sum(1 for k in q1_fidele if freq_cat(q2_months_dict[k]) == "Fidèle")
    to_q2_semi = sum(1 for k in q1_fidele if freq_cat(q2_months_dict[k]) == "Semi-fidèle")

    freq_rows = [
        ("Q2 Fidèle (3 mois)", to_q2_fidele, "Maintenu à un rythme mensuel — client stable."),
        ("Q2 Semi-fidèle (1-2 mois)", to_q2_semi, "Déclin de fréquence — risque d'attrition."),
        ("Q2 Zero (churned)", to_q2_zero, "PERTE SÈCHE — gros client fidèle devenu inactif."),
    ]
    r_fd = r_fh + 1
    for label, n, comment in freq_rows:
        color = "DDEBF7" if "Fidèle" in label and "Zero" not in label else "FFF2CC" if "Semi" in label else "FCE4E4"
        c = ws.cell(row=r_fd, column=1, value=label); c.font = BOLD_FONT
        c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
        c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        c = ws.cell(row=r_fd, column=2, value=n); c.font = BOLD_FONT
        c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
        c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        pct = (n/n_q1_fidele*100) if n_q1_fidele > 0 else 0
        c = ws.cell(row=r_fd, column=3, value=pct/100); c.font = BODY_FONT
        c.alignment = BODY_ALIGN_CENTER; c.border = BORDER; c.number_format = '0.0%'
        c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        c = ws.cell(row=r_fd, column=4, value=comment); c.font = BODY_FONT
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True); c.border = BORDER
        ws.merge_cells(start_row=r_fd, start_column=4, end_row=r_fd, end_column=8)
        c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        ws.row_dimensions[r_fd].height = 24
        r_fd += 1

    c = ws.cell(row=r_fd, column=1, value="TOTAL Q1 Fidèles")
    c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    c = ws.cell(row=r_fd, column=2, value=n_q1_fidele)
    c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c = ws.cell(row=r_fd, column=3, value=1.0)
    c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER; c.number_format = '0.0%'
    c = ws.cell(row=r_fd, column=4, value=""); c.fill = HEADER_FILL; c.border = BORDER
    ws.merge_cells(start_row=r_fd, start_column=4, end_row=r_fd, end_column=8)

    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 18
    ws.column_dimensions['G'].width = 18
    ws.column_dimensions['H'].width = 18


# ===== SHEET 7: TRANSITION SYNTHESE (CIBLÉ) =====
ws = wb.create_sheet("7. Transition Q1-Q2 (cible)")
build_transition_synthese(
    ws,
    title_text="BELGOCAM SA - Transition Q1 → Q2 2026 (16 produits ciblés : soja + concentrés)",
    sub_text="Analyse du comportement des 1 394 clients entre Q1 (Jan-Mar) et Q2 (Avr-Jun) 2026 sur les 16 produits ciblés.",
    q1_kg_dict=client_t_q1_kg, q2_kg_dict=client_t_q2_kg,
    q1_ca_dict=client_t_q1_ca, q2_ca_dict=client_t_q2_ca,
    q1_months_dict=client_t_q1_months, q2_months_dict=client_t_q2_months,
    segment_dict=client_seg_target,
    category_label="Ciblé (16 produits)"
)
print(f"  Sheet 7 (Transition synthèse ciblé) built")

# ===== SHEET 8: TRANSITION SYNTHESE (CONCENTRÉS) =====
ws = wb.create_sheet("8. Transition Q1-Q2 (conc.)")
build_transition_synthese(
    ws,
    title_text="BELGOCAM SA - Transition Q1 → Q2 2026 (12 concentrés : BELGO 10% + BELGO 5%)",
    sub_text="Analyse du comportement des 1 394 clients entre Q1 (Jan-Mar) et Q2 (Avr-Jun) 2026 sur les concentrés uniquement.",
    q1_kg_dict=client_c_q1_kg, q2_kg_dict=client_c_q2_kg,
    q1_ca_dict=client_c_q1_ca, q2_ca_dict=client_c_q2_ca,
    q1_months_dict=client_c_q1_months, q2_months_dict=client_c_q2_months,
    segment_dict=client_seg_concentre,
    category_label="Concentrés (12 produits)"
)
print(f"  Sheet 8 (Transition synthèse concentrés) built")

# ===== HELPER: Build detail transition sheet =====
def build_detail_transition(ws, title_text, q1_kg_dict, q2_kg_dict, q1_ca_dict, q2_ca_dict, q1_months_dict, q2_months_dict, segment_dict):
    ws.cell(row=1, column=1, value=title_text).font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = TITLE_ALIGN
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=13)
    ws.row_dimensions[1].height = 26

    HEADERS = ["N°", "Réf. client", "Nom du client", "Segment Q1→Q2",
               "Mois actifs Q1", "Fréquence Q1", "Mois actifs Q2", "Fréquence Q2",
               "Vol Q1 (t)", "CA Q1 (FCFA)", "Vol Q2 (t)", "CA Q2 (FCFA)", "20/80"]
    for col_idx, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=3, column=col_idx, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
    ws.row_dimensions[3].height = 36

    seg_order = {
        ("Active Q1","Zero Q2"): 0,
        ("Zero Q1","Active Q2"): 1,
        ("Active Q1","Active Q2"): 2,
        ("Zero Q1","Zero Q2"): 3,
    }

    def sort_key(k):
        seg = segment_dict[k]
        return (seg_order[seg], -(q1_ca_dict[k] + q2_ca_dict[k]))

    sorted_clients = sorted(clients_all, key=sort_key)

    START_ROW = 4
    for i, key in enumerate(sorted_clients, start=1):
        r = START_ROW + i - 1
        ref, name = key
        seg = segment_dict[key]
        seg_label = SEGMENT_LABELS[seg]
        color = SEG_COLORS[seg]
        font = SEG_FONTS[seg]
        q1_months = sorted(q1_months_dict[key])
        q2_months = sorted(q2_months_dict[key])
        q1_freq = freq_cat(q1_months_dict[key])
        q2_freq = freq_cat(q2_months_dict[key])
        q1_vol_t = q1_kg_dict[key] / 1000.0
        q2_vol_t = q2_kg_dict[key] / 1000.0
        q1_ca = q1_ca_dict[key]
        q2_ca = q2_ca_dict[key]
        fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

        cells = [
            (1, i, BODY_ALIGN_CENTER, None),
            (2, ref, BODY_ALIGN_CENTER, None),
            (3, name, BODY_ALIGN_LEFT, None),
            (4, seg_label, Alignment(horizontal="left", vertical="center", wrap_text=True), None),
            (5, ", ".join(str(m) for m in q1_months) if q1_months else "-", BODY_ALIGN_CENTER, None),
            (6, q1_freq, BODY_ALIGN_CENTER, None),
            (7, ", ".join(str(m) for m in q2_months) if q2_months else "-", BODY_ALIGN_CENTER, None),
            (8, q2_freq, BODY_ALIGN_CENTER, None),
            (9, q1_vol_t, BODY_ALIGN_RIGHT, VOL_NUM_FMT),
            (10, q1_ca, BODY_ALIGN_RIGHT, CA_NUM_FMT),
            (11, q2_vol_t, BODY_ALIGN_RIGHT, VOL_NUM_FMT),
            (12, q2_ca, BODY_ALIGN_RIGHT, CA_NUM_FMT),
        ]
        for col_idx, val, align, fmt in cells:
            c = ws.cell(row=r, column=col_idx, value=val)
            c.font = font; c.alignment = align; c.border = BORDER
            if fmt: c.number_format = fmt
            c.fill = fill
        c = ws.cell(row=r, column=13)
        c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
        if key in pareto_clients:
            c.value = PARETO_SYMBOL; c.fill = PARETO_FILL; c.font = PARETO_FONT
        else:
            c.fill = fill; c.font = font

    widths = [6, 16, 38, 38, 14, 14, 14, 14, 12, 18, 12, 18, 9]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "E4"
    ws.auto_filter.ref = f"A3:M{START_ROW + len(sorted_clients) - 1}"

# ===== SHEET 9: DÉTAIL TRANSITION (CIBLÉ) =====
ws = wb.create_sheet("9. Détail transition (cible)")
build_detail_transition(
    ws,
    title_text="BELGOCAM SA - Détail par client : transition Q1 → Q2 sur 16 produits ciblés",
    q1_kg_dict=client_t_q1_kg, q2_kg_dict=client_t_q2_kg,
    q1_ca_dict=client_t_q1_ca, q2_ca_dict=client_t_q2_ca,
    q1_months_dict=client_t_q1_months, q2_months_dict=client_t_q2_months,
    segment_dict=client_seg_target
)
print(f"  Sheet 9 (Détail transition ciblé) built — {len(clients_all)} rows")

# ===== SHEET 10: DÉTAIL TRANSITION (CONCENTRÉS) =====
ws = wb.create_sheet("10. Détail transition (conc.)")
build_detail_transition(
    ws,
    title_text="BELGOCAM SA - Détail par client : transition Q1 → Q2 sur 12 concentrés",
    q1_kg_dict=client_c_q1_kg, q2_kg_dict=client_c_q2_kg,
    q1_ca_dict=client_c_q1_ca, q2_ca_dict=client_c_q2_ca,
    q1_months_dict=client_c_q1_months, q2_months_dict=client_c_q2_months,
    segment_dict=client_seg_concentre
)
print(f"  Sheet 10 (Détail transition concentrés) built — {len(clients_all)} rows")

# ===== SAVE =====
wb.save(OUT)
print(f"\nSaved: {OUT}")
print(f"File size: {os.path.getsize(OUT):,} bytes")
print(f"Total sheets: {len(wb.sheetnames)}")
print(f"Sheets: {wb.sheetnames}")
