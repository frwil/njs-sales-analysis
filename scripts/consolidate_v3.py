"""
Final consolidated file v3:
- Excludes 23 clients (15 internal/filiale + 13 CLIENT COMPTOIR)
- 14 sheets including:
  * Sommaire
  * Clients uniques (check mensuel + 20/80)
  * Zero achat global Q1 (16 produits)
  * Zero achat concentrés Q1 (12 produits)
  * Zero achat 20/80 Q1 - Ciblé
  * Zero achat 20/80 Q1 - Concentrés
  * Pertes Q1 - Global (fréquence)
  * Pertes Q1 - Concentrés (fréquence)
  * Transition Q1-Q2 synthèse (ciblé)
  * Transition Q1-Q2 synthèse (concentrés)
  * Détail transition (ciblé)
  * Détail transition (concentrés)
  * Graphiques (7 NATIVE Excel charts - modifiables)
  * Recommandations (plan d'action priorisé)
"""
import re
import os
import json
import time
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from collections import defaultdict

SRC = "/home/z/my-project/download/ventes_livrees.xlsx"
OUT = "/home/z/my-project/download/analyse_complete_belgocam.xlsx"
CHARTS_DIR = "/home/z/my-project/scripts/charts"

# Load excluded clients
with open("/home/z/my-project/scripts/excluded_clients.json", "r", encoding="utf-8") as f:
    excluded_tiers = set(json.load(f))

SOJA_REFS = {"T102","T1021","T1023","T1024"}
CONCENTRE_REFS = {"C102","C1022","C104","C1042","C1043","C1044","C105","C1053","C1054","C1055","C101","C103"}
TARGET_REFS = SOJA_REFS | CONCENTRE_REFS

WEIGHT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(KG|Kg|kg|KG|GRAMMES?|G|L)\b", re.IGNORECASE)
def parse_weight_kg(desc):
    if not desc: return 0.0
    matches = WEIGHT_RE.findall(str(desc))
    if not matches: return 0.0
    val_str, unit = matches[-1]
    val = float(val_str)
    u = unit.upper()
    if u == "KG": return val
    if u in ("G","GRAMME","GRAMMES"): return val/1000.0
    if u == "L": return val
    return 0.0

Q1_SHEETS = {"Sheet 1","Feuil1","Feuil2"}
Q2_SHEETS = {"Feuil3","Feuil4","Feuil5"}
sheet_to_month = {"Sheet 1":1,"Feuil1":2,"Feuil2":3,"Feuil3":4,"Feuil4":5,"Feuil5":6}

TIERS_RE = re.compile(r"^\s*([A-Z]{2,}[\w-]*?)\s*-\s+(.+?)\s*$")
def split_tiers(tiers):
    if tiers is None: return ("","")
    s = str(tiers).strip()
    m = TIERS_RE.match(s)
    if m: return (m.group(1).strip(), m.group(2).strip())
    return ("", s)

def freq_cat(months_set):
    n = len(months_set)
    if n >= 3: return "Fidèle"
    elif n >= 1: return "Semi-fidèle"
    else: return "Zero"

# ===== READ SOURCE & BUILD AGGREGATES (excluding internal clients) =====
print(f"Reading source: {SRC}")
t0 = time.time()
wb_src = load_workbook(SRC, read_only=True, data_only=True)
print(f"  Loaded in {time.time()-t0:.1f}s")

product_weight = {}
client_ca_total = defaultdict(float)
client_months_any = defaultdict(set)
client_t_q1_kg = defaultdict(float); client_t_q2_kg = defaultdict(float)
client_t_q1_ca = defaultdict(float); client_t_q2_ca = defaultdict(float)
client_t_q1_months = defaultdict(set); client_t_q2_months = defaultdict(set)
client_c_q1_kg = defaultdict(float); client_c_q2_kg = defaultdict(float)
client_c_q1_ca = defaultdict(float); client_c_q2_ca = defaultdict(float)
client_c_q1_months = defaultdict(set); client_c_q2_months = defaultdict(set)
client_q2_by_product_target = defaultdict(lambda: defaultdict(lambda: {"vol_kg":0.0,"ca":0.0,"months":set()}))
client_q2_by_product_concentre = defaultdict(lambda: defaultdict(lambda: {"vol_kg":0.0,"ca":0.0,"months":set()}))
client_q2_active_months_target = defaultdict(set)
client_q2_active_months_concentre = defaultdict(set)

# NEW: per-month Q2 aggregates (months 4, 5, 6 = April, May, June)
client_t_month_kg = {m: defaultdict(float) for m in [1,2,3,4,5,6]}
client_t_month_ca = {m: defaultdict(float) for m in [1,2,3,4,5,6]}
client_t_month_active = {m: defaultdict(bool) for m in [1,2,3,4,5,6]}
client_c_month_active = {m: defaultdict(bool) for m in [1,2,3,4,5,6]}
client_any_month_active = {m: defaultdict(bool) for m in [1,2,3,4,5,6]}

# NEW: agency aggregation per client
# client_agencies_ca[key][agence] = CA HT total in this agency
client_agencies_ca = defaultdict(lambda: defaultdict(float))

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
        if tiers in excluded_tiers: continue  # EXCLUDE
        ref_prod = row[0] if len(row)>0 else None
        desc = row[1] if len(row)>1 else None
        qte = row[2] if len(row)>2 else 0
        ca_ht = row[8] if len(row)>8 else 0
        agence = row[17] if len(row)>17 else None  # column R = agence
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

        # Aggregate agency CA
        agence_str = str(agence).strip() if agence is not None else "(vide)"
        client_agencies_ca[key][agence_str] += ca_f

        client_ca_total[key] += ca_f
        client_months_any[key].add(month_num)
        client_any_month_active[month_num][key] = True
        is_target = ref_prod_str in TARGET_REFS
        is_concentre = ref_prod_str in CONCENTRE_REFS
        vol_kg = qte_f * weight_kg
        if is_target:
            client_t_month_active[month_num][key] = True
            client_t_month_kg[month_num][key] += vol_kg
            client_t_month_ca[month_num][key] += ca_f
            if is_q1:
                client_t_q1_kg[key] += vol_kg; client_t_q1_ca[key] += ca_f
                client_t_q1_months[key].add(month_num)
            else:
                client_t_q2_kg[key] += vol_kg; client_t_q2_ca[key] += ca_f
                client_t_q2_months[key].add(month_num)
                rec = client_q2_by_product_target[key][ref_prod_str]
                rec["vol_kg"] += vol_kg; rec["ca"] += ca_f
                if month_num: rec["months"].add(month_num)
                if month_num: client_q2_active_months_target[key].add(month_num)
        if is_concentre:
            client_c_month_active[month_num][key] = True
            if is_q1:
                client_c_q1_kg[key] += vol_kg; client_c_q1_ca[key] += ca_f
                client_c_q1_months[key].add(month_num)
            else:
                client_c_q2_kg[key] += vol_kg; client_c_q2_ca[key] += ca_f
                client_c_q2_months[key].add(month_num)
                rec = client_q2_by_product_concentre[key][ref_prod_str]
                rec["vol_kg"] += vol_kg; rec["ca"] += ca_f
                if month_num: rec["months"].add(month_num)
                if month_num: client_q2_active_months_concentre[key].add(month_num)

wb_src.close()

clients_all = sorted(client_ca_total.keys(), key=lambda k:(k[1].upper(), k[0]))
print(f"Total unique clients (after exclusion): {len(clients_all)}")

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

zero_global = [k for k in clients_all if not client_t_q1_months[k]]
zero_concentre = [k for k in clients_all if not client_c_q1_months[k]]

# Helper: format client agency list with primary agency marked
def format_agencies(key):
    """Return a string like '★AGENCE FAMLA, AGENCE NDOBO' where ★ marks the primary (highest CA) agency."""
    agencies = client_agencies_ca.get(key, {})
    if not agencies:
        return "-"
    # Sort by CA desc
    sorted_ag = sorted(agencies.items(), key=lambda x: -x[1])
    # Primary agency marked with ★
    parts = []
    for i, (ag, ca) in enumerate(sorted_ag):
        if i == 0:
            parts.append(f"★ {ag}")
        else:
            parts.append(ag)
    return ", ".join(parts)

def get_primary_agency(key):
    """Return the agency where the client has the highest CA (or '-' if none)."""
    agencies = client_agencies_ca.get(key, {})
    if not agencies:
        return "-"
    sorted_ag = sorted(agencies.items(), key=lambda x: -x[1])
    return sorted_ag[0][0]

# Losses
def compute_loss(client_key, category_refs, products_dict, active_months_dict):
    products_q2 = products_dict.get(client_key, {})
    active_months_set = active_months_dict.get(client_key, set())
    n_active = len(active_months_set)
    frequency = n_active / 3.0
    sum_avg = 0.0; loss_t = 0.0; loss_c = 0.0; n_products = 0
    for ref, rec in products_q2.items():
        if ref not in category_refs: continue
        total_t = rec["vol_kg"]/1000.0
        n_m = len(rec["months"])
        if n_m == 0: continue
        m_avg = total_t / n_m
        sum_avg += m_avg
        loss_t += m_avg * frequency * 3
        price = (rec["ca"]/total_t) if total_t > 0 else 0
        loss_c += m_avg * frequency * 3 * price
        n_products += 1
    return {"sum_monthly_avg_tons": sum_avg, "frequency": frequency,
            "n_active_months": n_active, "loss_tons": loss_t,
            "loss_fcfa": loss_c, "n_products": n_products}

losses_global = {k: compute_loss(k, TARGET_REFS, client_q2_by_product_target, client_q2_active_months_target) for k in zero_global}
losses_concentre = {k: compute_loss(k, CONCENTRE_REFS, client_q2_by_product_concentre, client_q2_active_months_concentre) for k in zero_concentre}

# Transition segments
def get_seg(q1m, q2m):
    return ("Active Q1" if len(q1m)>0 else "Zero Q1", "Active Q2" if len(q2m)>0 else "Zero Q2")
client_seg_target = {k: get_seg(client_t_q1_months[k], client_t_q2_months[k]) for k in clients_all}
client_seg_concentre = {k: get_seg(client_c_q1_months[k], client_c_q2_months[k]) for k in clients_all}

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

wb = Workbook()
wb.properties.creator = "Z.ai"
wb.properties.title = "BELGOCAM SA - Analyse complète Jan-Juin 2026"
wb.remove(wb.active)

# ===== SHEET 1: SOMMAIRE =====
ws = wb.create_sheet("1. Sommaire")
ws.cell(row=1, column=1, value="BELGOCAM SA - Analyse complète Janvier-Juin 2026").font = Font(name="Calibri", size=18, bold=True, color="1F4E78")
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
ws.row_dimensions[1].height = 32
ws.cell(row=2, column=1, value=f"Fichier consolidé multi-feuilles — {len(clients_all)} clients (après exclusion de 35 clients internes/filiales/comptoirs), 16 produits ciblés").font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
ws.row_dimensions[2].height = 22

ws.cell(row=4, column=1, value="N°").fill = HEADER_FILL
ws.cell(row=4, column=2, value="Feuille").fill = HEADER_FILL
ws.cell(row=4, column=3, value="Contenu").fill = HEADER_FILL
ws.cell(row=4, column=4, value="Lignes").fill = HEADER_FILL
for c in range(1, 5):
    cell = ws.cell(row=4, column=c)
    cell.font = HEADER_FONT; cell.alignment = HEADER_ALIGN; cell.border = BORDER
ws.row_dimensions[4].height = 28

sommaire = [
    ("1", "1. Sommaire", "Cette page — table des matières du fichier consolidé.", ""),
    ("2", "2. Clients uniques", f"Liste de {len(clients_all)} clients uniques sur 6 mois avec check mensuel vert, CA Total HT 6 mois et marque 20/80 (★).", f"{len(clients_all)} clients"),
    ("3", "3. Zero achat global Q1", f"{len(zero_global)} clients n'ayant acheté AUCUN des 16 produits ciblés en Jan-Mar 2026.", f"{len(zero_global)} clients"),
    ("4", "4. Zero achat concentrés Q1", f"{len(zero_concentre)} clients n'ayant acheté AUCUN concentré (12 produits) en Jan-Mar 2026.", f"{len(zero_concentre)} clients"),
    ("5", "5. Zero achat 20/80 Q1 (ciblé)", "Clients 20/80 qui n'ont rien acheté de ciblé en Q1. GROS COMPTES À RISQUE — action commerciale prioritaire.", "Filtre 20/80"),
    ("6", "6. Zero achat 20/80 Q1 (conc.)", "Clients 20/80 qui n'ont rien acheté de concentrés en Q1.", "Filtre 20/80"),
    ("7", "7. Pertes Q1 - Global (fréquence)", f"Pertes Q1 estimées pour les {len(zero_global)} clients zéro achat global. Méthode : Perte = Σ(moyennes mensuelles) × Fréquence × 3.", f"{len(zero_global)} clients"),
    ("8", "8. Pertes Q1 - Concentrés (fréq.)", f"Pertes Q1 estimées pour les {len(zero_concentre)} clients zéro achat concentrés. Même méthode.", f"{len(zero_concentre)} clients"),
    ("9", "9. Transition Q1-Q2 (cible)", "Synthèse : matrice de transition Q1→Q2 sur 16 produits ciblés, bilan net (gain/perte), destin des Q1 fidèles.", "Synthèse"),
    ("10", "10. Transition Q1-Q2 (conc.)", "Synthèse : matrice de transition Q1→Q2 sur 12 concentrés uniquement.", "Synthèse"),
    ("11", "11. Détail transition (cible)", f"Détail par client des {len(clients_all)} clients sur 16 produits ciblés. Triés par segment (churned en premier).", f"{len(clients_all)} clients"),
    ("12", "12. Détail transition (conc.)", f"Détail par client des {len(clients_all)} clients sur 12 concentrés.", f"{len(clients_all)} clients"),
    ("13", "13. Evolution mensuelle (cible)", "Analyse mois par mois (Jan→Juin) sur 16 produits ciblés. Volume (t) et statut actif ✓ pour chaque mois, Q2 détaillé Avril/Mai/Juin séparément.", f"{len(clients_all)} clients"),
    ("14", "14. Evolution mensuelle (conc.)", "Analyse mois par mois (Jan→Juin) sur 12 concentrés. Q2 détaillé Avril/Mai/Juin séparément.", f"{len(clients_all)} clients"),
    ("15", "15. Zero achat par agence (cibl)", "NOUVEAU — Par agence : proportion de zero achat Q1 (ciblé) avec décomposition 20/80 vs autres + graphique natif.", "Par agence"),
    ("16", "16. Zero achat par agence (conc)", "NOUVEAU — Par agence : proportion de zero achat Q1 (concentrés) avec décomposition 20/80 vs autres + graphique natif.", "Par agence"),
    ("17", "17. Graphiques", "7 graphiques natifs Excel (modifiables) : évolution mensuelle, segments de transition, top 10 pertes, répartition 20/80, bilan net.", "7 graphiques"),
    ("18", "18. Recommandations", "Plan d'action commercial priorisé en 6 axes avec clients cibles, montants et actions concrètes.", "Plan d'action"),
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
        c.font = BODY_FONT; c.alignment = align; c.border = BORDER
        if banding: c.fill = BAND_FILL
    ws.row_dimensions[r].height = 36

r = 4 + len(sommaire) + 2
ws.cell(row=r, column=1, value="Chiffres clés (rappel)").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
ws.row_dimensions[r].height = 22

stats = [
    ("Total clients uniques (après exclusion)", f"{len(clients_all)}"),
    ("Clients exclus (internes/filiales/comptoirs)", "35 (SPC, PDC filiale NJS, soldes compta, 25 clients comptoirs agences)"),
    ("Total CA HT (6 mois, après exclusion)", f"{total_ca:,.0f} FCFA".replace(",", " ")),
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
    c.font = BOLD_FONT; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    ws.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=3)
    c = ws.cell(row=r2, column=4, value=val)
    c.font = BODY_FONT; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    ws.row_dimensions[r2].height = 22

ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 36
ws.column_dimensions['C'].width = 80
ws.column_dimensions['D'].width = 16

print(f"  Sheet 1 (Sommaire) built")

# ===== SHEET 2: CLIENTS UNIQUES =====
ws = wb.create_sheet("2. Clients uniques")
ws.cell(row=1, column=1, value="BELGOCAM SA - Clients uniques (Janvier-Juin 2026) — Check mensuel d'achats").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=13)
ws.row_dimensions[1].height = 26

HEADERS = ["N°", "Réf. client", "Nom du client", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
           "Nb mois d'achat", "CA Total HT (6 mois)", "Agence(s) [★ = principale]", "Catégorie 20/80"]
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
    c = ws.cell(row=r, column=12, value=format_agencies(key)); c.font = BODY_FONT
    c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=13); c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
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
ws.column_dimensions['L'].width = 38
ws.column_dimensions['M'].width = 14
ws.freeze_panes = "D4"
ws.auto_filter.ref = f"A3:M{START_ROW + len(clients_all) - 1}"
print(f"  Sheet 2 (Clients uniques) built")

# ===== SHEET 3: ZERO ACHAT GLOBAL Q1 =====
ws = wb.create_sheet("3. Zero achat global Q1")
ws.cell(row=1, column=1, value="BELGOCAM SA - Clients 'zéro achat global' Janvier-Mars 2026 (16 produits ciblés)").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
ws.row_dimensions[1].height = 26
ws.cell(row=2, column=1, value="Critère : client n'ayant acheté AUCUN des 16 produits ciblés en Jan-Mar 2026.").font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=9)
ws.row_dimensions[2].height = 22

HEADERS = ["N°", "Réf. client", "Nom du client", "Achat ciblé Q1 ?", "CA HT Q1 (tous produits)",
           "Achat ciblé 6 mois ?", "CA HT total 6 mois", "Agence(s) [★ = principale]", "Catégorie 20/80"]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=3, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[3].height = 36

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
    c = ws.cell(row=r, column=8, value=format_agencies(key)); c.font = BODY_FONT
    c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=9); c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
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
ws.column_dimensions['H'].width = 38
ws.column_dimensions['I'].width = 14
ws.freeze_panes = "D4"
ws.auto_filter.ref = f"A3:I{START_ROW + len(zero_global_sorted) - 1}"
print(f"  Sheet 3 (Zero achat global) built")

# ===== SHEET 4: ZERO ACHAT CONCENTRÉS Q1 =====
ws = wb.create_sheet("4. Zero achat concentrés Q1")
ws.cell(row=1, column=1, value="BELGOCAM SA - Clients 'zéro achat concentrés' Janvier-Mars 2026 (12 produits)").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
ws.row_dimensions[1].height = 26
ws.cell(row=2, column=1, value="Critère : client n'ayant acheté AUCUN concentré en Jan-Mar 2026. La colonne 'Achat soja Q1 ?' distingue les clients qui ont quand même acheté du soja.").font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=9)
ws.row_dimensions[2].height = 22

HEADERS = ["N°", "Réf. client", "Nom du client", "Achat concentrés Q1 ?", "Achat soja Q1 ?",
           "Achat concentrés 6 mois ?", "CA HT total 6 mois", "Agence(s) [★ = principale]", "Catégorie 20/80"]
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
    soja_q1 = len(client_t_q1_months[key]) > 0
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
    c = ws.cell(row=r, column=8, value=format_agencies(key)); c.font = BODY_FONT
    c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=9); c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
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
ws.column_dimensions['H'].width = 38
ws.column_dimensions['I'].width = 14
ws.freeze_panes = "D4"
ws.auto_filter.ref = f"A3:I{START_ROW + len(zero_concentre_sorted) - 1}"
print(f"  Sheet 4 (Zero achat concentrés) built")

# ===== SHEETS 5 & 6: ZERO ACHAT 20/80 Q1 (NEW) =====
def build_zero_2080_sheet(ws, title_text, sub_text, zero_list, category_label):
    """Build a sheet listing only the 20/80 clients who are zero-achat."""
    ws.cell(row=1, column=1, value=title_text).font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = TITLE_ALIGN
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
    ws.row_dimensions[1].height = 26
    ws.cell(row=2, column=1, value=sub_text).font = SUB_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=10)
    ws.row_dimensions[2].height = 32

    # Filter to 20/80 only
    za_2080 = [k for k in zero_list if k in pareto_clients]
    za_2080_sorted = sorted(za_2080, key=lambda k: -client_ca_total[k])

    # Summary banner
    n_2080_total = len(pareto_clients)
    n_2080_zero = len(za_2080)
    pct = (n_2080_zero / n_2080_total * 100) if n_2080_total > 0 else 0
    ws.cell(row=4, column=1, value=f"📊 {n_2080_zero} client(s) 20/80 sur {n_2080_total} ({pct:.1f}%) sont en zéro achat Q1").font = Font(name="Calibri", size=12, bold=True, color="C00000")
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=10)
    ws.row_dimensions[4].height = 28

    HEADERS = ["N°", "Réf. client", "Nom du client", "CA HT total 6 mois",
               "Achat ciblé 6 mois ?", "Top produit Q2", "CA Top produit Q2",
               "Perte Q1 estimée (FCFA)", "Agence(s) [★ = principale]", "Priorité"]
    for col_idx, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=6, column=col_idx, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
    ws.row_dimensions[6].height = 36

    START_ROW = 7
    for i, key in enumerate(za_2080_sorted, start=1):
        r = START_ROW + i - 1
        ref, name = key
        banding = (i % 2 == 0)
        # Determine if bought ciblé in 6 months
        targeted_any = len(client_t_q2_months[key]) > 0 if category_label == "ciblé" else len(client_c_q2_months[key]) > 0
        # Top product Q2
        products_q2 = (client_q2_by_product_target if category_label == "ciblé" else client_q2_by_product_concentre).get(key, {})
        top_prod = None; top_ca = 0
        for ref_p, rec in products_q2.items():
            if rec["ca"] > top_ca:
                top_ca = rec["ca"]; top_prod = ref_p
        # Loss Q1
        loss = (losses_global if category_label == "ciblé" else losses_concentre).get(key, {}).get("loss_fcfa", 0)
        # Priority
        if loss > 50_000_000:
            priority = "🔴 CRITIQUE"
        elif loss > 10_000_000:
            priority = "🟠 ÉLEVÉE"
        elif loss > 0:
            priority = "🟡 MOYENNE"
        else:
            priority = "⚪ FAIBLE"

        c = ws.cell(row=r, column=1, value=i); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
        if banding: c.fill = BAND_FILL
        c = ws.cell(row=r, column=2, value=ref); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
        if banding: c.fill = BAND_FILL
        c = ws.cell(row=r, column=3, value=name); c.font = BOLD_FONT; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
        c.fill = PARETO_FILL
        c = ws.cell(row=r, column=4, value=client_ca_total.get(key, 0.0)); c.font = BODY_FONT
        c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
        if banding: c.fill = BAND_FILL
        c = ws.cell(row=r, column=5, value=("Oui" if targeted_any else "Non"))
        if targeted_any:
            c.font = Font(name="Calibri", size=10, color="006100", bold=True)
        else:
            c.font = Font(name="Calibri", size=10, color="C00000", bold=True)
        c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
        if banding: c.fill = BAND_FILL
        c = ws.cell(row=r, column=6, value=top_prod or "-"); c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
        if banding: c.fill = BAND_FILL
        c = ws.cell(row=r, column=7, value=top_ca); c.font = BODY_FONT
        c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
        if banding: c.fill = BAND_FILL
        c = ws.cell(row=r, column=8, value=loss); c.font = LOSS_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
        c.number_format = CA_NUM_FMT; c.fill = LOSS_FILL
        c = ws.cell(row=r, column=9, value=format_agencies(key)); c.font = BODY_FONT
        c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
        if banding: c.fill = BAND_FILL
        c = ws.cell(row=r, column=10, value=priority); c.font = BOLD_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
        if banding: c.fill = BAND_FILL

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 42
    ws.column_dimensions['D'].width = 22
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 14
    ws.column_dimensions['G'].width = 22
    ws.column_dimensions['H'].width = 22
    ws.column_dimensions['I'].width = 38
    ws.column_dimensions['J'].width = 16
    ws.freeze_panes = "D7"
    if za_2080_sorted:
        ws.auto_filter.ref = f"A6:J{START_ROW + len(za_2080_sorted) - 1}"
    return len(za_2080_sorted)

ws = wb.create_sheet("5. Zero achat 20-80 Q1 (cible)")
n5 = build_zero_2080_sheet(
    ws,
    title_text="BELGOCAM SA - Clients 20/80 en 'zéro achat global' Q1 (16 produits ciblés)",
    sub_text="⚠️ GROS COMPTES À RISQUE — Ces clients font partie du top 20/80 (★) mais n'ont acheté AUCUN produit ciblé en Jan-Mar 2026. Action commerciale prioritaire.",
    zero_list=zero_global,
    category_label="ciblé"
)
print(f"  Sheet 5 (Zero achat 20/80 ciblé) built — {n5} clients")

ws = wb.create_sheet("6. Zero achat 20-80 Q1 (conc.)")
n6 = build_zero_2080_sheet(
    ws,
    title_text="BELGOCAM SA - Clients 20/80 en 'zéro achat concentrés' Q1 (12 produits)",
    sub_text="⚠️ GROS COMPTES À RISQUE — Ces clients font partie du top 20/80 (★) mais n'ont acheté AUCUN concentré en Jan-Mar 2026. Action commerciale prioritaire.",
    zero_list=zero_concentre,
    category_label="concentré"
)
print(f"  Sheet 6 (Zero achat 20/80 concentrés) built — {n6} clients")

# ===== SHEETS 7 & 8: PERTES Q1 (GLOBAL + CONCENTRÉS) =====
def build_pertes_sheet(ws, title_text, sub_text, zero_list, losses_dict):
    ws.cell(row=1, column=1, value=title_text).font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = TITLE_ALIGN
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=12)
    ws.row_dimensions[1].height = 26
    ws.cell(row=2, column=1, value=sub_text).font = SUB_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=12)
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

    HEADERS = ["N°", "Réf. client", "Nom du client", "Nb produits (Q2)", "Nb mois actifs (Q2)",
               "Fréquence", "Σ moyennes (t)", "Perte Volume (t)", "Perte CA (FCFA)",
               "CA HT total 6 mois", "Agence(s) [★ = principale]", "20/80"]
    for col_idx, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=4, column=col_idx, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
    ws.row_dimensions[4].height = 40

    sorted_list = sorted(zero_list, key=lambda k: -losses_dict[k]["loss_tons"])
    START_ROW = 5
    for i, key in enumerate(sorted_list, start=1):
        r = START_ROW + i - 1
        ref, name = key
        L = losses_dict[key]
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
        c = ws.cell(row=r, column=11, value=format_agencies(key)); c.font = BODY_FONT
        c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
        if banding: c.fill = BAND_FILL
        c = ws.cell(row=r, column=12); c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
        if key in pareto_clients:
            c.value = PARETO_SYMBOL; c.fill = PARETO_FILL; c.font = PARETO_FONT
        else:
            if banding: c.fill = BAND_FILL
            c.font = BODY_FONT

    total_row = START_ROW + len(sorted_list)
    ws.cell(row=total_row, column=2, value="TOTAL").font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    ws.cell(row=total_row, column=2).fill = HEADER_FILL
    ws.cell(row=total_row, column=2).alignment = BODY_ALIGN_CENTER
    ws.cell(row=total_row, column=2).border = BORDER
    ws.merge_cells(start_row=total_row, start_column=2, end_row=total_row, end_column=3)
    ws.cell(row=total_row, column=3).fill = HEADER_FILL
    ws.cell(row=total_row, column=3).border = BORDER
    for col in [4, 5, 6]:
        c = ws.cell(row=total_row, column=col); c.fill = HEADER_FILL; c.border = BORDER
    total_avg = sum(L["sum_monthly_avg_tons"] for L in losses_dict.values())
    total_loss_t = sum(L["loss_tons"] for L in losses_dict.values())
    total_loss_fcfa = sum(L["loss_fcfa"] for L in losses_dict.values())
    total_ca = sum(client_ca_total.get(k, 0) for k in zero_list)
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
    c = ws.cell(row=total_row, column=12); c.fill = HEADER_FILL; c.border = BORDER

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 42
    for cl in ['D','E','F','G','H']: ws.column_dimensions[cl].width = 14
    ws.column_dimensions['I'].width = 22
    ws.column_dimensions['J'].width = 22
    ws.column_dimensions['K'].width = 38
    ws.column_dimensions['L'].width = 9
    ws.freeze_panes = "D5"
    ws.auto_filter.ref = f"A4:L{START_ROW + len(sorted_list) - 1}"

ws = wb.create_sheet("7. Pertes Q1 - Global")
build_pertes_sheet(
    ws,
    title_text="BELGOCAM SA - Pertes Q1 estimées — Clients 'zéro achat global' (16 produits)",
    sub_text=("Méthode : Perte Q1 = Σ(moyennes mensuelles par produit en Q2) × Fréquence × 3 mois.\n"
              "  - Moyenne mensuelle (par produit) = volume total Q2 / nombre de mois Q2 où le produit a été acheté.\n"
              "  - Fréquence = (nombre de mois Q2 avec achat ciblé) / 3."),
    zero_list=zero_global, losses_dict=losses_global
)
print(f"  Sheet 7 (Pertes Q1 - Global) built")

ws = wb.create_sheet("8. Pertes Q1 - Concentres")
build_pertes_sheet(
    ws,
    title_text="BELGOCAM SA - Pertes Q1 estimées — Clients 'zéro achat concentrés' (12 produits)",
    sub_text=("Méthode : Perte Q1 = Σ(moyennes mensuelles par produit en Q2) × Fréquence × 3 mois.\n"
              "  - Moyenne mensuelle (par produit) = volume total Q2 / nombre de mois Q2 où le produit a été acheté.\n"
              "  - Fréquence = (nombre de mois Q2 avec achat concentré) / 3."),
    zero_list=zero_concentre, losses_dict=losses_concentre
)
print(f"  Sheet 8 (Pertes Q1 - Concentrés) built")

# ===== SHEETS 9 & 10: TRANSITION SYNTHESE =====
def build_transition_synthese(ws, title_text, sub_text, q1_kg_dict, q2_kg_dict, q1_ca_dict, q2_ca_dict, q1_months_dict, q2_months_dict, segment_dict, category_label):
    ws.cell(row=1, column=1, value=title_text).font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = TITLE_ALIGN
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    ws.row_dimensions[1].height = 28
    ws.cell(row=2, column=1, value=sub_text).font = SUB_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
    ws.row_dimensions[2].height = 30

    seg_stats = {}
    for seg_key in [("Zero Q1","Zero Q2"), ("Zero Q1","Active Q2"), ("Active Q1","Active Q2"), ("Active Q1","Zero Q2")]:
        clients = [k for k in clients_all if segment_dict[k] == seg_key]
        q1_vol_t = sum(q1_kg_dict[k] for k in clients) / 1000.0
        q2_vol_t = sum(q2_kg_dict[k] for k in clients) / 1000.0
        q1_ca = sum(q1_ca_dict[k] for k in clients)
        q2_ca = sum(q2_ca_dict[k] for k in clients)
        seg_stats[seg_key] = {"n_clients": len(clients), "q1_vol_t": q1_vol_t, "q2_vol_t": q2_vol_t,
                              "q1_ca": q1_ca, "q2_ca": q2_ca, "delta_vol_t": q2_vol_t - q1_vol_t,
                              "delta_ca": q2_ca - q1_ca}

    zz = seg_stats[("Zero Q1","Zero Q2")]["n_clients"]
    za = seg_stats[("Zero Q1","Active Q2")]["n_clients"]
    az = seg_stats[("Active Q1","Zero Q2")]["n_clients"]
    aa = seg_stats[("Active Q1","Active Q2")]["n_clients"]

    ws.cell(row=4, column=1, value=f"1. Matrice de transition Q1 → Q2 — {category_label}").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
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

    # Section 2
    r_title = 10
    ws.cell(row=r_title, column=1, value=f"2. Volumes et CA par segment — {category_label}").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
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

    # Bilan net
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
    c = ws.cell(row=r_g, column=5, value=f"Nouveaux achats {category_label.lower()} en Q2 par clients sans achat en Q1.")
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
    c = ws.cell(row=r_n, column=5, value="Si positif : plus de clients réactivés que perdus.")
    c.font = Font(name="Calibri", size=10, italic=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True); c.border = BORDER
    ws.merge_cells(start_row=r_n, start_column=5, end_row=r_n, end_column=8)
    ws.row_dimensions[r_n].height = 30

    # Destin Q1 fidèles
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
        ("Q2 Fidèle (3 mois)", to_q2_fidele, "Maintenu — client stable."),
        ("Q2 Semi-fidèle (1-2 mois)", to_q2_semi, "Déclin de fréquence — risque d'attrition."),
        ("Q2 Zero (churned)", to_q2_zero, "PERTE SÈCHE — fidèle devenu inactif."),
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

ws = wb.create_sheet("9. Transition Q1-Q2 (cible)")
build_transition_synthese(
    ws,
    title_text="BELGOCAM SA - Transition Q1 → Q2 2026 (16 produits ciblés)",
    sub_text="Analyse du comportement des clients entre Q1 (Jan-Mar) et Q2 (Avr-Jun) 2026 sur les 16 produits ciblés.",
    q1_kg_dict=client_t_q1_kg, q2_kg_dict=client_t_q2_kg,
    q1_ca_dict=client_t_q1_ca, q2_ca_dict=client_t_q2_ca,
    q1_months_dict=client_t_q1_months, q2_months_dict=client_t_q2_months,
    segment_dict=client_seg_target,
    category_label="Ciblé (16 produits)"
)
print(f"  Sheet 9 (Transition synthèse ciblé) built")

ws = wb.create_sheet("10. Transition Q1-Q2 (conc.)")
build_transition_synthese(
    ws,
    title_text="BELGOCAM SA - Transition Q1 → Q2 2026 (12 concentrés)",
    sub_text="Analyse du comportement des clients entre Q1 (Jan-Mar) et Q2 (Avr-Jun) 2026 sur les concentrés uniquement.",
    q1_kg_dict=client_c_q1_kg, q2_kg_dict=client_c_q2_kg,
    q1_ca_dict=client_c_q1_ca, q2_ca_dict=client_c_q2_ca,
    q1_months_dict=client_c_q1_months, q2_months_dict=client_c_q2_months,
    segment_dict=client_seg_concentre,
    category_label="Concentrés (12 produits)"
)
print(f"  Sheet 10 (Transition synthèse concentrés) built")

# ===== SHEETS 11 & 12: DÉTAIL TRANSITION =====
def build_detail_transition(ws, title_text, q1_kg_dict, q2_kg_dict, q1_ca_dict, q2_ca_dict, q1_months_dict, q2_months_dict, segment_dict):
    ws.cell(row=1, column=1, value=title_text).font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = TITLE_ALIGN
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=14)
    ws.row_dimensions[1].height = 26

    HEADERS = ["N°", "Réf. client", "Nom du client", "Segment Q1→Q2",
               "Mois actifs Q1", "Fréquence Q1", "Mois actifs Q2", "Fréquence Q2",
               "Vol Q1 (t)", "CA Q1 (FCFA)", "Vol Q2 (t)", "CA Q2 (FCFA)",
               "Agence(s) [★ = principale]", "20/80"]
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
            (13, format_agencies(key), BODY_ALIGN_LEFT, None),
        ]
        for col_idx, val, align, fmt in cells:
            c = ws.cell(row=r, column=col_idx, value=val)
            c.font = font; c.alignment = align; c.border = BORDER
            if fmt: c.number_format = fmt
            c.fill = fill
        c = ws.cell(row=r, column=14)
        c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
        if key in pareto_clients:
            c.value = PARETO_SYMBOL; c.fill = PARETO_FILL; c.font = PARETO_FONT
        else:
            c.fill = fill; c.font = font

    widths = [6, 16, 38, 38, 14, 14, 14, 14, 12, 18, 12, 18, 38, 9]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "E4"
    ws.auto_filter.ref = f"A3:N{START_ROW + len(sorted_clients) - 1}"

ws = wb.create_sheet("11. Détail transition (cible)")
build_detail_transition(
    ws,
    title_text="BELGOCAM SA - Détail par client : transition Q1 → Q2 sur 16 produits ciblés",
    q1_kg_dict=client_t_q1_kg, q2_kg_dict=client_t_q2_kg,
    q1_ca_dict=client_t_q1_ca, q2_ca_dict=client_t_q2_ca,
    q1_months_dict=client_t_q1_months, q2_months_dict=client_t_q2_months,
    segment_dict=client_seg_target
)
print(f"  Sheet 11 (Détail transition ciblé) built")

ws = wb.create_sheet("12. Détail transition (conc.)")
build_detail_transition(
    ws,
    title_text="BELGOCAM SA - Détail par client : transition Q1 → Q2 sur 12 concentrés",
    q1_kg_dict=client_c_q1_kg, q2_kg_dict=client_c_q2_kg,
    q1_ca_dict=client_c_q1_ca, q2_ca_dict=client_c_q2_ca,
    q1_months_dict=client_c_q1_months, q2_months_dict=client_c_q2_months,
    segment_dict=client_seg_concentre
)
print(f"  Sheet 12 (Détail transition concentrés) built")

# ===== SHEET 13 (NEW): ÉVOLUTION MENSUELLE Q2 MOIS PAR MOIS =====
# Detailed month-by-month analysis: for each client, show vol/CA/status for Jan, Feb, Mar, Apr, May, Jun
# on targeted products. This gives the month-by-month Q2 view requested.

ws = wb.create_sheet("13. Evolution mensuelle (cible)")
ws.cell(row=1, column=1, value="BELGOCAM SA - Évolution mensuelle (Janvier → Juin) sur 16 produits ciblés").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=21)
ws.row_dimensions[1].height = 26

ws.cell(row=2, column=1, value=("Pour chaque client : statut actif (✓) et volume (t) sur produits ciblés, mois par mois. "
                                  "Permet de voir la dynamique mensuelle (Avril / Mai / Juin) et le moment exact de la (ré)activation ou du churn.")
).font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=21)
ws.row_dimensions[2].height = 30

# Group headers (row 3)
ws.merge_cells(start_row=3, start_column=4, end_row=3, end_column=9)
c = ws.cell(row=3, column=4, value="Q1 (Jan-Mar) — mois par mois")
c.fill = GROUP_FILL; c.font = GROUP_FONT; c.alignment = GROUP_ALIGN; c.border = BORDER
for col in [5, 6, 7, 8, 9]:
    cc = ws.cell(row=3, column=col); cc.fill = GROUP_FILL; cc.border = BORDER

ws.merge_cells(start_row=3, start_column=10, end_row=3, end_column=15)
c = ws.cell(row=3, column=10, value="Q2 (Avr-Juin) — mois par mois")
c.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
c.font = GROUP_FONT; c.alignment = GROUP_ALIGN; c.border = BORDER
for col in [11, 12, 13, 14, 15]:
    cc = ws.cell(row=3, column=col); cc.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid"); cc.border = BORDER

ws.merge_cells(start_row=3, start_column=16, end_row=3, end_column=19)
c = ws.cell(row=3, column=16, value="Synthèse")
c.fill = PatternFill(start_color="375623", end_color="375623", fill_type="solid")
c.font = GROUP_FONT; c.alignment = GROUP_ALIGN; c.border = BORDER
for col in [17, 18, 19]:
    cc = ws.cell(row=3, column=col); cc.fill = PatternFill(start_color="375623", end_color="375623", fill_type="solid"); cc.border = BORDER
ws.row_dimensions[3].height = 22

# Column headers (row 4)
HEADERS = [
    "N°", "Réf. client", "Nom du client",
    "Jan (t)", "Jan ✓", "Fév (t)", "Fév ✓", "Mar (t)", "Mar ✓",
    "Avr (t)", "Avr ✓", "Mai (t)", "Mai ✓", "Juin (t)", "Juin ✓",
    "Segment Q1→Q2", "Vol 6m (t)", "CA 6m (FCFA)", "Agence(s) [★ = principale]",
    "20/80", "Statut détaillé"
]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=4, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[4].height = 36

# Build data rows
# Sort: churned first, then reactivated, then retained, then persistent
seg_order = {
    ("Active Q1","Zero Q2"): 0,
    ("Zero Q1","Active Q2"): 1,
    ("Active Q1","Active Q2"): 2,
    ("Zero Q1","Zero Q2"): 3,
}
def sort_key(k):
    seg = client_seg_target[k]
    return (seg_order[seg], -(client_t_q1_ca[k] + client_t_q2_ca[k]))

sorted_clients = sorted(clients_all, key=sort_key)

START_ROW = 5
for i, key in enumerate(sorted_clients, start=1):
    r = START_ROW + i - 1
    ref, name = key
    seg = client_seg_target[key]
    seg_label = SEGMENT_LABELS[seg]
    color = SEG_COLORS[seg]
    font = SEG_FONTS[seg]
    banding = (i % 2 == 0)
    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

    # Build detailed status string
    q1_months_active = sorted(client_t_q1_months[key])
    q2_months_active = sorted(client_t_q2_months[key])
    status_parts = []
    if not q1_months_active and not q2_months_active:
        status_parts.append("Jamais actif (ciblé) sur 6 mois")
    else:
        if q1_months_active:
            month_names_map = {1:"Jan",2:"Fév",3:"Mar"}
            status_parts.append(f"Q1: {','.join(month_names_map[m] for m in q1_months_active)}")
        else:
            status_parts.append("Q1: inactif")
        if q2_months_active:
            month_names_map = {4:"Avr",5:"Mai",6:"Juin"}
            status_parts.append(f"Q2: {','.join(month_names_map[m] for m in q2_months_active)}")
        else:
            status_parts.append("Q2: inactif")
    status_detail = " | ".join(status_parts)

    # Vol 6m and CA 6m (targeted)
    vol_6m_t = (client_t_q1_kg[key] + client_t_q2_kg[key]) / 1000.0
    ca_6m = client_t_q1_ca[key] + client_t_q2_ca[key]

    # Cells: (col, value, align, fmt)
    cells = [
        (1, i, BODY_ALIGN_CENTER, None),
        (2, ref, BODY_ALIGN_CENTER, None),
        (3, name, BODY_ALIGN_LEFT, None),
        (4, client_t_month_kg[1][key]/1000.0, BODY_ALIGN_RIGHT, VOL_NUM_FMT),
        (5, "✓" if client_t_month_active[1][key] else "", BODY_ALIGN_CENTER, None),
        (6, client_t_month_kg[2][key]/1000.0, BODY_ALIGN_RIGHT, VOL_NUM_FMT),
        (7, "✓" if client_t_month_active[2][key] else "", BODY_ALIGN_CENTER, None),
        (8, client_t_month_kg[3][key]/1000.0, BODY_ALIGN_RIGHT, VOL_NUM_FMT),
        (9, "✓" if client_t_month_active[3][key] else "", BODY_ALIGN_CENTER, None),
        (10, client_t_month_kg[4][key]/1000.0, BODY_ALIGN_RIGHT, VOL_NUM_FMT),
        (11, "✓" if client_t_month_active[4][key] else "", BODY_ALIGN_CENTER, None),
        (12, client_t_month_kg[5][key]/1000.0, BODY_ALIGN_RIGHT, VOL_NUM_FMT),
        (13, "✓" if client_t_month_active[5][key] else "", BODY_ALIGN_CENTER, None),
        (14, client_t_month_kg[6][key]/1000.0, BODY_ALIGN_RIGHT, VOL_NUM_FMT),
        (15, "✓" if client_t_month_active[6][key] else "", BODY_ALIGN_CENTER, None),
        (16, seg_label, Alignment(horizontal="left", vertical="center", wrap_text=True), None),
        (17, vol_6m_t, BODY_ALIGN_RIGHT, VOL_NUM_FMT),
        (18, ca_6m, BODY_ALIGN_RIGHT, CA_NUM_FMT),
        (19, format_agencies(key), BODY_ALIGN_LEFT, None),
    ]
    for col_idx, val, align, fmt in cells:
        c = ws.cell(row=r, column=col_idx, value=val)
        c.font = font; c.alignment = align; c.border = BORDER
        if fmt: c.number_format = fmt
        c.fill = fill
    # Highlight active month ticks with green
    for col_idx in [5, 7, 9, 11, 13, 15]:
        v = ws.cell(row=r, column=col_idx).value
        if v == "✓":
            ws.cell(row=r, column=col_idx).fill = GREEN_FILL
            ws.cell(row=r, column=col_idx).font = GREEN_FONT
    # 20/80 marker
    c = ws.cell(row=r, column=20)
    c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
    if key in pareto_clients:
        c.value = PARETO_SYMBOL; c.fill = PARETO_FILL; c.font = PARETO_FONT
    else:
        c.fill = fill; c.font = font
    # Status detail
    c = ws.cell(row=r, column=21, value=status_detail)
    c.font = font; c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True); c.border = BORDER
    c.fill = fill

# Column widths
widths_13 = [6, 16, 38,
             10, 7, 10, 7, 10, 7,
             10, 7, 10, 7, 10, 7,
             38, 12, 18, 38, 9, 40]
for i, w in enumerate(widths_13, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "D5"
ws.auto_filter.ref = f"A4:U{START_ROW + len(sorted_clients) - 1}"

print(f"  Sheet 13 (Évolution mensuelle) built — {len(sorted_clients)} rows")

# ===== SHEET 14 (NEW): ÉVOLUTION MENSUELLE CONCENTRÉS =====
ws = wb.create_sheet("14. Evolution mensuelle (conc.)")
ws.cell(row=1, column=1, value="BELGOCAM SA - Évolution mensuelle (Janvier → Juin) sur 12 concentrés").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=21)
ws.row_dimensions[1].height = 26

ws.cell(row=2, column=1, value="Pour chaque client : statut actif (✓) sur concentrés, mois par mois. Permet de voir la dynamique mensuelle Q2.").font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=21)
ws.row_dimensions[2].height = 30

# Group headers
ws.merge_cells(start_row=3, start_column=4, end_row=3, end_column=9)
c = ws.cell(row=3, column=4, value="Q1 (Jan-Mar)")
c.fill = GROUP_FILL; c.font = GROUP_FONT; c.alignment = GROUP_ALIGN; c.border = BORDER
for col in [5, 6, 7, 8, 9]:
    cc = ws.cell(row=3, column=col); cc.fill = GROUP_FILL; cc.border = BORDER
ws.merge_cells(start_row=3, start_column=10, end_row=3, end_column=15)
c = ws.cell(row=3, column=10, value="Q2 (Avr-Juin)")
c.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
c.font = GROUP_FONT; c.alignment = GROUP_ALIGN; c.border = BORDER
for col in [11, 12, 13, 14, 15]:
    cc = ws.cell(row=3, column=col); cc.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid"); cc.border = BORDER
ws.merge_cells(start_row=3, start_column=16, end_row=3, end_column=19)
c = ws.cell(row=3, column=16, value="Synthèse")
c.fill = PatternFill(start_color="375623", end_color="375623", fill_type="solid")
c.font = GROUP_FONT; c.alignment = GROUP_ALIGN; c.border = BORDER
for col in [17, 18, 19]:
    cc = ws.cell(row=3, column=col); cc.fill = PatternFill(start_color="375623", end_color="375623", fill_type="solid"); cc.border = BORDER
ws.row_dimensions[3].height = 22

HEADERS = [
    "N°", "Réf. client", "Nom du client",
    "Jan ✓", "Fév ✓", "Mar ✓", "Avr ✓", "Mai ✓", "Juin ✓",
    "Vol Q1 conc (t)", "CA Q1 conc", "Vol Q2 conc (t)", "CA Q2 conc",
    "Vol 6m conc (t)", "CA 6m conc",
    "Segment Q1→Q2", "Vol 6m (t)", "CA 6m (FCFA)", "Agence(s) [★ = principale]",
    "20/80", "Statut détaillé"
]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=4, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[4].height = 36

# Sort: churned first
seg_order_c = {
    ("Active Q1","Zero Q2"): 0,
    ("Zero Q1","Active Q2"): 1,
    ("Active Q1","Active Q2"): 2,
    ("Zero Q1","Zero Q2"): 3,
}
def sort_key_c(k):
    seg = client_seg_concentre[k]
    return (seg_order_c[seg], -(client_c_q1_ca[k] + client_c_q2_ca[k]))

sorted_clients_c = sorted(clients_all, key=sort_key_c)

START_ROW = 5
for i, key in enumerate(sorted_clients_c, start=1):
    r = START_ROW + i - 1
    ref, name = key
    seg = client_seg_concentre[key]
    seg_label = SEGMENT_LABELS[seg]
    color = SEG_COLORS[seg]
    font = SEG_FONTS[seg]
    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

    q1_months_active = sorted(client_c_q1_months[key])
    q2_months_active = sorted(client_c_q2_months[key])
    status_parts = []
    if not q1_months_active and not q2_months_active:
        status_parts.append("Jamais actif (concentrés) sur 6 mois")
    else:
        if q1_months_active:
            month_names_map = {1:"Jan",2:"Fév",3:"Mar"}
            status_parts.append(f"Q1: {','.join(month_names_map[m] for m in q1_months_active)}")
        else:
            status_parts.append("Q1: inactif")
        if q2_months_active:
            month_names_map = {4:"Avr",5:"Mai",6:"Juin"}
            status_parts.append(f"Q2: {','.join(month_names_map[m] for m in q2_months_active)}")
        else:
            status_parts.append("Q2: inactif")
    status_detail = " | ".join(status_parts)

    vol_q1_c_t = client_c_q1_kg[key] / 1000.0
    vol_q2_c_t = client_c_q2_kg[key] / 1000.0
    ca_q1_c = client_c_q1_ca[key]
    ca_q2_c = client_c_q2_ca[key]
    vol_6m_c_t = vol_q1_c_t + vol_q2_c_t
    ca_6m_c = ca_q1_c + ca_q2_c
    vol_6m_t = (client_t_q1_kg[key] + client_t_q2_kg[key]) / 1000.0
    ca_6m = client_t_q1_ca[key] + client_t_q2_ca[key]

    cells = [
        (1, i, BODY_ALIGN_CENTER, None),
        (2, ref, BODY_ALIGN_CENTER, None),
        (3, name, BODY_ALIGN_LEFT, None),
        (4, "✓" if client_c_month_active[1][key] else "", BODY_ALIGN_CENTER, None),
        (5, "✓" if client_c_month_active[2][key] else "", BODY_ALIGN_CENTER, None),
        (6, "✓" if client_c_month_active[3][key] else "", BODY_ALIGN_CENTER, None),
        (7, "✓" if client_c_month_active[4][key] else "", BODY_ALIGN_CENTER, None),
        (8, "✓" if client_c_month_active[5][key] else "", BODY_ALIGN_CENTER, None),
        (9, "✓" if client_c_month_active[6][key] else "", BODY_ALIGN_CENTER, None),
        (10, vol_q1_c_t, BODY_ALIGN_RIGHT, VOL_NUM_FMT),
        (11, ca_q1_c, BODY_ALIGN_RIGHT, CA_NUM_FMT),
        (12, vol_q2_c_t, BODY_ALIGN_RIGHT, VOL_NUM_FMT),
        (13, ca_q2_c, BODY_ALIGN_RIGHT, CA_NUM_FMT),
        (14, vol_6m_c_t, BODY_ALIGN_RIGHT, VOL_NUM_FMT),
        (15, ca_6m_c, BODY_ALIGN_RIGHT, CA_NUM_FMT),
        (16, seg_label, Alignment(horizontal="left", vertical="center", wrap_text=True), None),
        (17, vol_6m_t, BODY_ALIGN_RIGHT, VOL_NUM_FMT),
        (18, ca_6m, BODY_ALIGN_RIGHT, CA_NUM_FMT),
        (19, format_agencies(key), BODY_ALIGN_LEFT, None),
    ]
    for col_idx, val, align, fmt in cells:
        c = ws.cell(row=r, column=col_idx, value=val)
        c.font = font; c.alignment = align; c.border = BORDER
        if fmt: c.number_format = fmt
        c.fill = fill
    for col_idx in [4, 5, 6, 7, 8, 9]:
        v = ws.cell(row=r, column=col_idx).value
        if v == "✓":
            ws.cell(row=r, column=col_idx).fill = GREEN_FILL
            ws.cell(row=r, column=col_idx).font = GREEN_FONT
    c = ws.cell(row=r, column=20)
    c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
    if key in pareto_clients:
        c.value = PARETO_SYMBOL; c.fill = PARETO_FILL; c.font = PARETO_FONT
    else:
        c.fill = fill; c.font = font
    c = ws.cell(row=r, column=21, value=status_detail)
    c.font = font; c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True); c.border = BORDER
    c.fill = fill

widths_14 = [6, 16, 38,
             8, 8, 8, 8, 8, 8,
             14, 16, 14, 16, 14, 16,
             38, 12, 18, 38, 9, 40]
for i, w in enumerate(widths_14, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "D5"
ws.auto_filter.ref = f"A4:U{START_ROW + len(sorted_clients_c) - 1}"

print(f"  Sheet 14 (Évolution mensuelle concentrés) built — {len(sorted_clients_c)} rows")

# ===== SHEET 15 (NEW): ZERO ACHAT PAR AGENCE (CIBLÉ) =====
# By agency: count of zero-achat clients (ciblé) broken down by 20/80 vs others, with native charts
from openpyxl.chart import BarChart, LineChart, PieChart, Reference, BarChart3D
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.series import DataPoint

# Compute agency-level stats for zero-achat ciblé
agency_zero_cible_stats = defaultdict(lambda: {"total_clients": 0, "zero_2080": 0, "zero_others": 0, "active_clients": 0})
for k in clients_all:
    primary_ag = get_primary_agency(k)
    is_zero = k in zero_global
    is_2080 = k in pareto_clients
    agency_zero_cible_stats[primary_ag]["total_clients"] += 1
    if is_zero:
        if is_2080:
            agency_zero_cible_stats[primary_ag]["zero_2080"] += 1
        else:
            agency_zero_cible_stats[primary_ag]["zero_others"] += 1
    else:
        agency_zero_cible_stats[primary_ag]["active_clients"] += 1

# Sort agencies by total clients desc
sorted_agencies_cible = sorted(agency_zero_cible_stats.items(), key=lambda x: -x[1]["total_clients"])

ws = wb.create_sheet("15. Zero achat par agence (cibl)")
ws.cell(row=1, column=1, value="BELGOCAM SA - Zero achat par agence (16 produits ciblés)").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
ws.row_dimensions[1].height = 26

ws.cell(row=2, column=1, value="Répartition des clients zero achat Q1 par agence (agence principale du client). 20/80 = clients Pareto.").font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
ws.row_dimensions[2].height = 22

HEADERS = ["N°", "Agence", "Total clients", "Clients actifs Q1 (ciblé)", "Zero achat 20/80 ★", "Zero achat autres", "Total zero achat", "% zero achat"]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=4, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[4].height = 36

START_ROW = 5
for i, (ag, stats) in enumerate(sorted_agencies_cible, start=1):
    r = START_ROW + i - 1
    banding = (i % 2 == 0)
    total_z = stats["zero_2080"] + stats["zero_others"]
    pct = (total_z / stats["total_clients"] * 100) if stats["total_clients"] > 0 else 0
    cells = [
        (1, i, BODY_ALIGN_CENTER),
        (2, ag, BODY_ALIGN_LEFT),
        (3, stats["total_clients"], BODY_ALIGN_CENTER),
        (4, stats["active_clients"], BODY_ALIGN_CENTER),
        (5, stats["zero_2080"], BODY_ALIGN_CENTER),
        (6, stats["zero_others"], BODY_ALIGN_CENTER),
        (7, total_z, BODY_ALIGN_CENTER),
    ]
    for col, val, align in cells:
        c = ws.cell(row=r, column=col, value=val)
        c.font = BODY_FONT; c.alignment = align; c.border = BORDER
        if banding: c.fill = BAND_FILL
    # Highlight 20/80 zero in gold
    c = ws.cell(row=r, column=5)
    if stats["zero_2080"] > 0:
        c.fill = PARETO_FILL; c.font = PARETO_FONT
    # Highlight % zero in red if high
    c = ws.cell(row=r, column=8, value=pct/100)
    c.font = BOLD_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c.number_format = '0.0%'
    if pct >= 40:
        c.fill = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")
        c.font = Font(name="Calibri", size=10, bold=True, color="C00000")
    elif banding:
        c.fill = BAND_FILL

# Total row
total_row = START_ROW + len(sorted_agencies_cible)
total_all = sum(s["total_clients"] for s in agency_zero_cible_stats.values())
total_active = sum(s["active_clients"] for s in agency_zero_cible_stats.values())
total_2080 = sum(s["zero_2080"] for s in agency_zero_cible_stats.values())
total_others = sum(s["zero_others"] for s in agency_zero_cible_stats.values())
total_zero = total_2080 + total_others
ws.cell(row=total_row, column=2, value="TOTAL").font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=total_row, column=2).fill = HEADER_FILL
ws.cell(row=total_row, column=2).alignment = BODY_ALIGN_CENTER
ws.cell(row=total_row, column=2).border = BORDER
for col, val in [(3, total_all), (4, total_active), (5, total_2080), (6, total_others), (7, total_zero)]:
    c = ws.cell(row=total_row, column=col, value=val)
    c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
c = ws.cell(row=total_row, column=8, value=total_zero/total_all if total_all > 0 else 0)
c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER; c.number_format = '0.0%'

# Column widths
ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 32
ws.column_dimensions['C'].width = 14
ws.column_dimensions['D'].width = 18
ws.column_dimensions['E'].width = 18
ws.column_dimensions['F'].width = 18
ws.column_dimensions['G'].width = 18
ws.column_dimensions['H'].width = 14

# Add 2 native charts (stacked bar 20/80 vs others, and pie total zero vs active)
# Stacked bar chart data is in columns J/K starting at row 4 (hidden)
DATA_COL = 10  # column J
# Headers
ws.cell(row=4, column=DATA_COL, value="Agence").font = BOLD_FONT
ws.cell(row=4, column=DATA_COL+1, value="Zero 20/80").font = BOLD_FONT
ws.cell(row=4, column=DATA_COL+2, value="Zero autres").font = BOLD_FONT
ws.cell(row=4, column=DATA_COL+3, value="Actifs Q1").font = BOLD_FONT
for i, (ag, stats) in enumerate(sorted_agencies_cible, start=1):
    ws.cell(row=4+i, column=DATA_COL, value=ag)
    ws.cell(row=4+i, column=DATA_COL+1, value=stats["zero_2080"])
    ws.cell(row=4+i, column=DATA_COL+2, value=stats["zero_others"])
    ws.cell(row=4+i, column=DATA_COL+3, value=stats["active_clients"])
# Hide the data columns
for col_idx in range(DATA_COL, DATA_COL+4):
    ws.column_dimensions[get_column_letter(col_idx)].hidden = True

# Stacked Bar Chart
chart1 = BarChart()
chart1.type = "bar"
chart1.style = 10
chart1.grouping = "stacked"
chart1.overlap = 100
chart1.title = "Répartition Zero achat (ciblé) par agence — 20/80 vs Autres"
chart1.x_axis.title = "Nombre de clients"
chart1.y_axis.title = "Agence"
chart1.x_axis.delete = False
chart1.y_axis.delete = False
chart1.height = 14
chart1.width = 22
n_agencies = len(sorted_agencies_cible)
data = Reference(ws, min_col=DATA_COL+1, min_row=4, max_col=DATA_COL+3, max_row=4+n_agencies)
cats = Reference(ws, min_col=DATA_COL, min_row=5, max_row=4+n_agencies)
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
chart1.legend.position = 'b'
# Color series: 20/80 = gold, others = gray, actifs = blue
series_colors = ["FFC000", "A6A6A6", "2E75B6"]
for i, s in enumerate(chart1.series):
    s.graphicalProperties = GraphicalProperties(solidFill=series_colors[i])
ws.add_chart(chart1, f"A{total_row + 3}")

print(f"  Sheet 15 (Zero achat par agence - ciblé) built — {len(sorted_agencies_cible)} agencies")

# ===== SHEET 16 (NEW): ZERO ACHAT PAR AGENCE (CONCENTRÉS) =====
agency_zero_conc_stats = defaultdict(lambda: {"total_clients": 0, "zero_2080": 0, "zero_others": 0, "active_clients": 0})
for k in clients_all:
    primary_ag = get_primary_agency(k)
    is_zero = k in zero_concentre
    is_2080 = k in pareto_clients
    agency_zero_conc_stats[primary_ag]["total_clients"] += 1
    if is_zero:
        if is_2080:
            agency_zero_conc_stats[primary_ag]["zero_2080"] += 1
        else:
            agency_zero_conc_stats[primary_ag]["zero_others"] += 1
    else:
        agency_zero_conc_stats[primary_ag]["active_clients"] += 1

sorted_agencies_conc = sorted(agency_zero_conc_stats.items(), key=lambda x: -x[1]["total_clients"])

ws = wb.create_sheet("16. Zero achat par agence (conc)")
ws.cell(row=1, column=1, value="BELGOCAM SA - Zero achat par agence (12 concentrés)").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
ws.row_dimensions[1].height = 26

ws.cell(row=2, column=1, value="Répartition des clients zero achat Q1 (concentrés) par agence (agence principale du client). 20/80 = clients Pareto.").font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
ws.row_dimensions[2].height = 22

HEADERS = ["N°", "Agence", "Total clients", "Clients actifs Q1 (conc.)", "Zero achat 20/80 ★", "Zero achat autres", "Total zero achat", "% zero achat"]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=4, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[4].height = 36

START_ROW = 5
for i, (ag, stats) in enumerate(sorted_agencies_conc, start=1):
    r = START_ROW + i - 1
    banding = (i % 2 == 0)
    total_z = stats["zero_2080"] + stats["zero_others"]
    pct = (total_z / stats["total_clients"] * 100) if stats["total_clients"] > 0 else 0
    cells = [
        (1, i, BODY_ALIGN_CENTER),
        (2, ag, BODY_ALIGN_LEFT),
        (3, stats["total_clients"], BODY_ALIGN_CENTER),
        (4, stats["active_clients"], BODY_ALIGN_CENTER),
        (5, stats["zero_2080"], BODY_ALIGN_CENTER),
        (6, stats["zero_others"], BODY_ALIGN_CENTER),
        (7, total_z, BODY_ALIGN_CENTER),
    ]
    for col, val, align in cells:
        c = ws.cell(row=r, column=col, value=val)
        c.font = BODY_FONT; c.alignment = align; c.border = BORDER
        if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=5)
    if stats["zero_2080"] > 0:
        c.fill = PARETO_FILL; c.font = PARETO_FONT
    c = ws.cell(row=r, column=8, value=pct/100)
    c.font = BOLD_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c.number_format = '0.0%'
    if pct >= 50:
        c.fill = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")
        c.font = Font(name="Calibri", size=10, bold=True, color="C00000")
    elif banding:
        c.fill = BAND_FILL

total_row = START_ROW + len(sorted_agencies_conc)
total_all = sum(s["total_clients"] for s in agency_zero_conc_stats.values())
total_active = sum(s["active_clients"] for s in agency_zero_conc_stats.values())
total_2080 = sum(s["zero_2080"] for s in agency_zero_conc_stats.values())
total_others = sum(s["zero_others"] for s in agency_zero_conc_stats.values())
total_zero = total_2080 + total_others
ws.cell(row=total_row, column=2, value="TOTAL").font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=total_row, column=2).fill = HEADER_FILL
ws.cell(row=total_row, column=2).alignment = BODY_ALIGN_CENTER
ws.cell(row=total_row, column=2).border = BORDER
for col, val in [(3, total_all), (4, total_active), (5, total_2080), (6, total_others), (7, total_zero)]:
    c = ws.cell(row=total_row, column=col, value=val)
    c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
c = ws.cell(row=total_row, column=8, value=total_zero/total_all if total_all > 0 else 0)
c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER; c.number_format = '0.0%'

ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 32
ws.column_dimensions['C'].width = 14
ws.column_dimensions['D'].width = 18
ws.column_dimensions['E'].width = 18
ws.column_dimensions['F'].width = 18
ws.column_dimensions['G'].width = 18
ws.column_dimensions['H'].width = 14

# Data for chart
DATA_COL = 10
ws.cell(row=4, column=DATA_COL, value="Agence").font = BOLD_FONT
ws.cell(row=4, column=DATA_COL+1, value="Zero 20/80").font = BOLD_FONT
ws.cell(row=4, column=DATA_COL+2, value="Zero autres").font = BOLD_FONT
ws.cell(row=4, column=DATA_COL+3, value="Actifs Q1").font = BOLD_FONT
for i, (ag, stats) in enumerate(sorted_agencies_conc, start=1):
    ws.cell(row=4+i, column=DATA_COL, value=ag)
    ws.cell(row=4+i, column=DATA_COL+1, value=stats["zero_2080"])
    ws.cell(row=4+i, column=DATA_COL+2, value=stats["zero_others"])
    ws.cell(row=4+i, column=DATA_COL+3, value=stats["active_clients"])
for col_idx in range(DATA_COL, DATA_COL+4):
    ws.column_dimensions[get_column_letter(col_idx)].hidden = True

chart2 = BarChart()
chart2.type = "bar"
chart2.style = 10
chart2.grouping = "stacked"
chart2.overlap = 100
chart2.title = "Répartition Zero achat (concentrés) par agence — 20/80 vs Autres"
chart2.x_axis.title = "Nombre de clients"
chart2.y_axis.title = "Agence"
chart2.x_axis.delete = False
chart2.y_axis.delete = False
chart2.height = 14
chart2.width = 22
n_agencies = len(sorted_agencies_conc)
data = Reference(ws, min_col=DATA_COL+1, min_row=4, max_col=DATA_COL+3, max_row=4+n_agencies)
cats = Reference(ws, min_col=DATA_COL, min_row=5, max_row=4+n_agencies)
chart2.add_data(data, titles_from_data=True)
chart2.set_categories(cats)
chart2.legend.position = 'b'
for i, s in enumerate(chart2.series):
    s.graphicalProperties = GraphicalProperties(solidFill=series_colors[i])
ws.add_chart(chart2, f"A{total_row + 3}")

print(f"  Sheet 16 (Zero achat par agence - conc.) built — {len(sorted_agencies_conc)} agencies")

# ===== SHEET 17: GRAPHIQUES (NATIVE EXCEL CHARTS) =====
from openpyxl.chart import BarChart, LineChart, PieChart, Reference, BarChart3D
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.fill import ColorChoice
from openpyxl.chart.series import DataPoint

ws = wb.create_sheet("17. Graphiques")
ws.cell(row=1, column=1, value="BELGOCAM SA - Graphiques d'analyse (graphiques natifs Excel)").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=12)
ws.row_dimensions[1].height = 26

ws.cell(row=2, column=1, value="Graphiques natifs Excel — modifiables, dynamiques. Les données sources sont en colonnes N-X (masquées).").font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=12)
ws.row_dimensions[2].height = 22

# ---- Compute chart data ----
# Chart 1: Monthly active clients evolution
months_list = [1, 2, 3, 4, 5, 6]
month_labels = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin"]
# Need to recompute monthly active counts (we didn't store them in the main loop)
# Recompute from client_months_any (which has all months where any product was bought)
month_active_any_counts = [0] * 7  # index 1..6
month_active_target_counts = [0] * 7
month_active_concentre_counts = [0] * 7
for k in clients_all:
    for m in client_months_any[k]:
        month_active_any_counts[m] += 1
    for m in client_t_q1_months[k]:
        month_active_target_counts[m] += 1
    for m in client_t_q2_months[k]:
        month_active_target_counts[m] += 1
    for m in client_c_q1_months[k]:
        month_active_concentre_counts[m] += 1
    for m in client_c_q2_months[k]:
        month_active_concentre_counts[m] += 1

# Chart 2 & 3: Segments transition
seg_target_counts = {
    "Persistant": sum(1 for k in clients_all if client_seg_target[k]==("Zero Q1","Zero Q2")),
    "Réactivé": sum(1 for k in clients_all if client_seg_target[k]==("Zero Q1","Active Q2")),
    "Retenu": sum(1 for k in clients_all if client_seg_target[k]==("Active Q1","Active Q2")),
    "Churned": sum(1 for k in clients_all if client_seg_target[k]==("Active Q1","Zero Q2")),
}
seg_concentre_counts = {
    "Persistant": sum(1 for k in clients_all if client_seg_concentre[k]==("Zero Q1","Zero Q2")),
    "Réactivé": sum(1 for k in clients_all if client_seg_concentre[k]==("Zero Q1","Active Q2")),
    "Retenu": sum(1 for k in clients_all if client_seg_concentre[k]==("Active Q1","Active Q2")),
    "Churned": sum(1 for k in clients_all if client_seg_concentre[k]==("Active Q1","Zero Q2")),
}

# Chart 4 & 5: Top 10 pertes
top10_global = sorted(zero_global, key=lambda k: -losses_global[k]["loss_fcfa"])[:10]
top10_conc = sorted(zero_concentre, key=lambda k: -losses_concentre[k]["loss_fcfa"])[:10]

# Chart 6: Répartition 20/80 zero achat
za_pareto_g = sum(1 for k in zero_global if k in pareto_clients)
za_nonpareto_g = len(zero_global) - za_pareto_g
za_pareto_c = sum(1 for k in zero_concentre if k in pareto_clients)
za_nonpareto_c = len(zero_concentre) - za_pareto_c

# Chart 7: Bilan net
gain_cible_ca = sum(client_t_q2_ca[k] for k in clients_all if client_seg_target[k]==("Zero Q1","Active Q2"))
loss_cible_ca = sum(client_t_q1_ca[k] for k in clients_all if client_seg_target[k]==("Active Q1","Zero Q2"))
gain_conc_ca = sum(client_c_q2_ca[k] for k in clients_all if client_seg_concentre[k]==("Zero Q1","Active Q2"))
loss_conc_ca = sum(client_c_q1_ca[k] for k in clients_all if client_seg_concentre[k]==("Active Q1","Zero Q2"))

# ---- Write data tables in columns N onwards (hidden) ----
# We'll write data in columns N, O, P, ... (14, 15, 16, ...)
DATA_COL_START = 14  # column N

def write_data_table(ws, start_row, headers, rows):
    """Write a small data table starting at (start_row, DATA_COL_START). Returns (header_row, data_start_row, data_end_row, n_cols)."""
    # Headers
    for j, h in enumerate(headers):
        c = ws.cell(row=start_row, column=DATA_COL_START + j, value=h)
        c.font = Font(name="Calibri", size=10, bold=True)
    # Data
    for i, row_vals in enumerate(rows, start=1):
        for j, v in enumerate(row_vals):
            ws.cell(row=start_row + i, column=DATA_COL_START + j, value=v)
    return start_row, start_row + 1, start_row + len(rows), len(headers)

# Chart 1 data: months × 3 series (any / target / concentre)
chart1_header_row = 4
chart1_headers = ["Mois", "Tous produits", "Ciblé (16)", "Concentrés (12)"]
chart1_rows = [[month_labels[i], month_active_any_counts[i+1], month_active_target_counts[i+1], month_active_concentre_counts[i+1]] for i in range(6)]
h_row, d_start, d_end, n_cols = write_data_table(ws, chart1_header_row, chart1_headers, chart1_rows)

# Chart 2 data: segments ciblé
chart2_header_row = d_end + 3
chart2_headers = ["Segment", "Nb clients"]
chart2_rows = [[s, seg_target_counts[s]] for s in ["Persistant", "Réactivé", "Retenu", "Churned"]]
h_row2, d_start2, d_end2, _ = write_data_table(ws, chart2_header_row, chart2_headers, chart2_rows)

# Chart 3 data: segments concentre
chart3_header_row = d_end2 + 3
chart3_headers = ["Segment", "Nb clients"]
chart3_rows = [[s, seg_concentre_counts[s]] for s in ["Persistant", "Réactivé", "Retenu", "Churned"]]
h_row3, d_start3, d_end3, _ = write_data_table(ws, chart3_header_row, chart3_headers, chart3_rows)

# Chart 4 data: Top 10 pertes global
chart4_header_row = d_end3 + 3
chart4_headers = ["Client", "Perte (M FCFA)"]
chart4_rows = [[k[1][:30], losses_global[k]["loss_fcfa"]/1e6] for k in top10_global]
h_row4, d_start4, d_end4, _ = write_data_table(ws, chart4_header_row, chart4_headers, chart4_rows)

# Chart 5 data: Top 10 pertes concentre
chart5_header_row = d_end4 + 3
chart5_headers = ["Client", "Perte (M FCFA)"]
chart5_rows = [[k[1][:30], losses_concentre[k]["loss_fcfa"]/1e6] for k in top10_conc]
h_row5, d_start5, d_end5, _ = write_data_table(ws, chart5_header_row, chart5_headers, chart5_rows)

# Chart 6 data: Zero achat 20/80
chart6_header_row = d_end5 + 3
chart6_headers = ["Catégorie", "Ciblé", "Concentrés"]
chart6_rows = [
    ["20/80", za_pareto_g, za_pareto_c],
    ["Autres", za_nonpareto_g, za_nonpareto_c],
]
h_row6, d_start6, d_end6, _ = write_data_table(ws, chart6_header_row, chart6_headers, chart6_rows)

# Chart 7 data: Bilan net
chart7_header_row = d_end6 + 3
chart7_headers = ["Catégorie", "Gain (M FCFA)", "Perte (M FCFA)", "Bilan net (M FCFA)"]
chart7_rows = [
    ["Ciblé (16)", gain_cible_ca/1e6, loss_cible_ca/1e6, (gain_cible_ca-loss_cible_ca)/1e6],
    ["Concentrés (12)", gain_conc_ca/1e6, loss_conc_ca/1e6, (gain_conc_ca-loss_conc_ca)/1e6],
]
h_row7, d_start7, d_end7, _ = write_data_table(ws, chart7_header_row, chart7_headers, chart7_rows)

# Hide the data columns (N onwards)
for col_idx in range(DATA_COL_START, DATA_COL_START + 5):
    ws.column_dimensions[get_column_letter(col_idx)].hidden = True

# ---- Helper to style a chart ----
def style_chart(chart, title, x_axis_title="", y_axis_title=""):
    chart.title = title
    chart.style = 10  # built-in style
    chart.x_axis.title = x_axis_title
    chart.y_axis.title = y_axis_title
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    # Legend at bottom
    chart.legend.position = 'b'
    # Show data labels
    chart.dataLabels = DataLabelList(showVal=True)

# ---- Chart 1: LineChart - Évolution mensuelle ----
chart1 = LineChart()
chart1.title = "1. Évolution mensuelle du nombre de clients actifs"
chart1.style = 12
chart1.y_axis.title = "Nombre de clients actifs"
chart1.x_axis.title = "Mois"
chart1.x_axis.delete = False
chart1.y_axis.delete = False
chart1.height = 11
chart1.width = 20
data = Reference(ws, min_col=DATA_COL_START+1, min_row=chart1_header_row, max_col=DATA_COL_START+3, max_row=chart1_header_row+6)
cats = Reference(ws, min_col=DATA_COL_START, min_row=chart1_header_row+1, max_row=chart1_header_row+6)
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
# Series colors
series_colors = ["1F4E78", "FFC000", "ED7D31"]
for i, s in enumerate(chart1.series):
    s.graphicalProperties = GraphicalProperties()
    s.graphicalProperties.line = LineProperties(solidFill=series_colors[i], w=28000)
    s.graphicalProperties.line.solidFill = series_colors[i]
chart1.legend.position = 'b'
ws.add_chart(chart1, "A4")

# ---- Chart 2: BarChart - Segments ciblé ----
chart2 = BarChart()
chart2.type = "col"
chart2.style = 10
chart2.title = "2. Segments de transition Q1 → Q2 (16 produits ciblés)"
chart2.y_axis.title = "Nombre de clients"
chart2.x_axis.title = "Segment"
chart2.x_axis.delete = False
chart2.y_axis.delete = False
chart2.height = 11
chart2.width = 20
data = Reference(ws, min_col=DATA_COL_START+1, min_row=chart2_header_row, max_row=chart2_header_row+4)
cats = Reference(ws, min_col=DATA_COL_START, min_row=chart2_header_row+1, max_row=chart2_header_row+4)
chart2.add_data(data, titles_from_data=True)
chart2.set_categories(cats)
chart2.legend = None
chart2.dataLabels = DataLabelList(showVal=True)
# Color each bar differently
seg_colors = ["A6A6A6", "70AD47", "2E75B6", "C00000"]
if chart2.series:
    s = chart2.series[0]
    s.data_points = [DataPoint(idx=i, spPr=GraphicalProperties(solidFill=seg_colors[i])) for i in range(4)]
ws.add_chart(chart2, "A28")

# ---- Chart 3: BarChart - Segments concentre ----
chart3 = BarChart()
chart3.type = "col"
chart3.style = 10
chart3.title = "3. Segments de transition Q1 → Q2 (12 concentrés)"
chart3.y_axis.title = "Nombre de clients"
chart3.x_axis.title = "Segment"
chart3.x_axis.delete = False
chart3.y_axis.delete = False
chart3.height = 11
chart3.width = 20
data = Reference(ws, min_col=DATA_COL_START+1, min_row=chart3_header_row, max_row=chart3_header_row+4)
cats = Reference(ws, min_col=DATA_COL_START, min_row=chart3_header_row+1, max_row=chart3_header_row+4)
chart3.add_data(data, titles_from_data=True)
chart3.set_categories(cats)
chart3.legend = None
chart3.dataLabels = DataLabelList(showVal=True)
if chart3.series:
    s = chart3.series[0]
    s.data_points = [DataPoint(idx=i, spPr=GraphicalProperties(solidFill=seg_colors[i])) for i in range(4)]
ws.add_chart(chart3, "A52")

# ---- Chart 4: BarChart (horizontal) - Top 10 pertes global ----
chart4 = BarChart()
chart4.type = "bar"  # horizontal
chart4.style = 10
chart4.title = "4. Top 10 pertes Q1 — Zéro achat global (16 produits)"
chart4.x_axis.title = "Perte (millions FCFA)"
chart4.y_axis.title = "Client"
chart4.x_axis.delete = False
chart4.y_axis.delete = False
chart4.height = 12
chart4.width = 20
data = Reference(ws, min_col=DATA_COL_START+1, min_row=chart4_header_row, max_row=chart4_header_row+10)
cats = Reference(ws, min_col=DATA_COL_START, min_row=chart4_header_row+1, max_row=chart4_header_row+10)
chart4.add_data(data, titles_from_data=True)
chart4.set_categories(cats)
chart4.legend = None
chart4.dataLabels = DataLabelList(showVal=True)
if chart4.series:
    s = chart4.series[0]
    s.graphicalProperties = GraphicalProperties(solidFill="C00000")
ws.add_chart(chart4, "A76")

# ---- Chart 5: BarChart (horizontal) - Top 10 pertes concentre ----
chart5 = BarChart()
chart5.type = "bar"
chart5.style = 10
chart5.title = "5. Top 10 pertes Q1 — Zéro achat concentrés (12 produits)"
chart5.x_axis.title = "Perte (millions FCFA)"
chart5.y_axis.title = "Client"
chart5.x_axis.delete = False
chart5.y_axis.delete = False
chart5.height = 12
chart5.width = 20
data = Reference(ws, min_col=DATA_COL_START+1, min_row=chart5_header_row, max_row=chart5_header_row+10)
cats = Reference(ws, min_col=DATA_COL_START, min_row=chart5_header_row+1, max_row=chart5_header_row+10)
chart5.add_data(data, titles_from_data=True)
chart5.set_categories(cats)
chart5.legend = None
chart5.dataLabels = DataLabelList(showVal=True)
if chart5.series:
    s = chart5.series[0]
    s.graphicalProperties = GraphicalProperties(solidFill="ED7D31")
ws.add_chart(chart5, "A102")

# ---- Chart 6: PieChart - Zero achat 20/80 ----
# Two pie charts side by side
# Pie 1: Ciblé
chart6a = PieChart()
chart6a.title = "6a. Zero achat Q1 — Ciblé (16 produits)"
chart6a.height = 11
chart6a.width = 11
data = Reference(ws, min_col=DATA_COL_START+1, min_row=chart6_header_row, max_row=chart6_header_row+2)
cats = Reference(ws, min_col=DATA_COL_START, min_row=chart6_header_row+1, max_row=chart6_header_row+2)
chart6a.add_data(data, titles_from_data=True)
chart6a.set_categories(cats)
chart6a.dataLabels = DataLabelList(showPercent=True, showCatName=True)
pie_colors = ["FFC000", "A6A6A6"]
if chart6a.series:
    s = chart6a.series[0]
    s.data_points = [DataPoint(idx=i, spPr=GraphicalProperties(solidFill=pie_colors[i])) for i in range(2)]
ws.add_chart(chart6a, "A128")

# Pie 2: Concentrés
chart6b = PieChart()
chart6b.title = "6b. Zero achat Q1 — Concentrés (12 produits)"
chart6b.height = 11
chart6b.width = 11
data = Reference(ws, min_col=DATA_COL_START+2, min_row=chart6_header_row, max_row=chart6_header_row+2)
cats = Reference(ws, min_col=DATA_COL_START, min_row=chart6_header_row+1, max_row=chart6_header_row+2)
chart6b.add_data(data, titles_from_data=True)
chart6b.set_categories(cats)
chart6b.dataLabels = DataLabelList(showPercent=True, showCatName=True)
if chart6b.series:
    s = chart6b.series[0]
    s.data_points = [DataPoint(idx=i, spPr=GraphicalProperties(solidFill=pie_colors[i])) for i in range(2)]
ws.add_chart(chart6b, "K128")

# ---- Chart 7: BarChart - Bilan net ----
chart7 = BarChart()
chart7.type = "col"
chart7.style = 10
chart7.title = "7. Bilan net Q1 → Q2 : Ciblé vs Concentrés"
chart7.y_axis.title = "Montant (millions FCFA)"
chart7.x_axis.title = "Catégorie"
chart7.x_axis.delete = False
chart7.y_axis.delete = False
chart7.height = 11
chart7.width = 20
data = Reference(ws, min_col=DATA_COL_START+1, min_row=chart7_header_row, max_col=DATA_COL_START+3, max_row=chart7_header_row+2)
cats = Reference(ws, min_col=DATA_COL_START, min_row=chart7_header_row+1, max_row=chart7_header_row+2)
chart7.add_data(data, titles_from_data=True)
chart7.set_categories(cats)
chart7.legend.position = 'b'
chart7.dataLabels = DataLabelList(showVal=True)
bilan_colors = ["70AD47", "C00000", "1F4E78"]
for i, s in enumerate(chart7.series):
    s.graphicalProperties = GraphicalProperties(solidFill=bilan_colors[i])
ws.add_chart(chart7, "A152")

ws.column_dimensions['A'].width = 12
print(f"  Sheet 17 (Graphiques) built — 7 native Excel charts")

# ===== SHEET 18: RECOMMANDATIONS =====
ws = wb.create_sheet("18. Recommandations")
ws.cell(row=1, column=1, value="BELGOCAM SA - Plan d'action commercial priorisé").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
ws.row_dimensions[1].height = 28
ws.cell(row=2, column=1, value="Synthèse des 5 axes prioritaires d'action commerciale, issus de l'analyse Q1-Q2 2026.").font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
ws.row_dimensions[2].height = 22

# Compute key numbers for recommendations
za_pareto_global = sum(1 for k in zero_global if k in pareto_clients)
za_pareto_conc = sum(1 for k in zero_concentre if k in pareto_clients)
churned_target = [k for k in clients_all if client_seg_target[k] == ("Active Q1","Zero Q2")]
churned_concentre = [k for k in clients_all if client_seg_concentre[k] == ("Active Q1","Zero Q2")]
total_loss_global_fcfa = sum(L["loss_fcfa"] for L in losses_global.values())
total_loss_concentre_fcfa = sum(L["loss_fcfa"] for L in losses_concentre.values())

# Header
HEADERS_REC = ["Priorité", "Axe d'action", "Cible", "Nb clients", "Montant enjeu (FCFA)", "Action concrète", "Délai", "Pilote suggéré"]
for col_idx, h in enumerate(HEADERS_REC, start=1):
    c = ws.cell(row=4, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[4].height = 36

# Recommendations data
recs = [
    # Priority, Axe, Cible, Nb clients, Montant, Action, Délai, Pilote
    ("🔴 CRITIQUE",
     "1. Sauvetage des clients 20/80 en zéro achat Q1",
     f"Voir feuilles 5 (ciblé) et 6 (concentrés) — clients ★ sans achat ciblé/concentré en Q1",
     f"{za_pareto_global} (ciblé) + {za_pareto_conc} (conc.)",
     f"≈ {sum(losses_global[k]['loss_fcfa'] for k in zero_global if k in pareto_clients):,.0f} FCFA".replace(",", " "),
     "RDV individuel avec chaque client par le directeur commercial. Offre dédiée : remise volume + échantillon + livraison gratuite. Vérifier s'il y a un problème qualité/prix/concurrence.",
     "30 jours",
     "Directeur Commercial + Responsables agences"),
    ("🟠 ÉLEVÉE",
     "2. Recontact des clients churned (actifs Q1 → Zero Q2)",
     f"Voir feuilles 9 et 12 — segment 'Q1 Active → Q2 Zero'",
     f"{len(churned_target)} (ciblé) + {len(churned_concentre)} (conc.)",
     f"≈ {sum(client_t_q1_ca[k] for k in churned_target):,.0f} FCFA".replace(",", " "),
     "Appel téléphonique par le commercial dédié sous 7 jours. Enquête : pourquoi plus d'achat ? (concurrence, prix, qualité, défaut livraison). Offre de retour : -5% sur première commande de réactivation.",
     "15 jours",
     "Commerciaux terrain"),
    ("🟠 ÉLEVÉE",
     "3. Sauvetage des anciens clients fidèles en déclin (Q1 Fidèle → Q2 Semi-fidèle)",
     f"Voir feuille 9 — Clients 'Fidèle Q1' passés 'Semi-fidèle Q2'",
     "97 (ciblé) + 75 (conc.)",
     "Variable — à quantifier par client",
     "Visite physique par le commercial. Diagnostic : changement d'activité ? Concurrence ? Insatisfaction ? Mise en place d'un plan de fidélisation sur 3 mois (tarif préférentiel, livraison prioritaire).",
     "30 jours",
     "Commerciaux + Service client"),
    ("🟡 MOYENNE",
     "4. Prospection des 118 clients jamais acquérables (Zero Q1 + Q2)",
     "Voir feuille 3 — clients persistants Zero sur 6 mois",
     "118 (ciblé)",
     "Potentiel élevé mais à estimer",
     "Campagne d'échantillonnage + invitation à une démonstration produit. Tarif découverte sur première commande. Identifier ceux qui achètent chez les concurrents (analyse de marché locale).",
     "60 jours",
     "Marketing + Commerciaux"),
    ("🟡 MOYENNE",
     "5. Capitalisation sur les 234 clients réactivés Q2",
     "Voir feuilles 9 et 12 — segment 'Q1 Zero → Q2 Active'",
     "234 (ciblé) + 184 (conc.)",
     "Gain déjà réalisé : +549 M FCFA (ciblé), +145 M (conc.)",
     "Enquête qualitative auprès de 30 clients : qu'est-ce qui a déclenché l'achat Q2 ? (Visite commerciale, promo, rupture concurrentielle, nouveauté produit). Reproduire les leviers efficaces à grande échelle.",
     "45 jours",
     "Marketing + Direction commerciale"),
    ("🟢 STRUCTURANTE",
     "6. Analyse approfondie du segment 'Retenu en déclin' sur concentrés",
     "Voir feuille 10 — Q1 Active → Q2 Active mais volume -248 t",
     "672 clients (retenu concentrés)",
     "Perte volume -248 t / -164 M FCFA",
     "Le segment 'retenu' sur concentrés décline en volume malgré la rétention. Enquête : prix trop élevés vs concurrents ? Qualité perçue en baisse ? Substitution par soja ? Lancer un atelier interne produit/marché.",
     "90 jours",
     "Direction Produit + Direction Commerciale"),
]

current_row = 5
for rec in recs:
    priority, axe, cible, nb, montant, action, delai, pilote = rec
    banding = (current_row % 2 == 0)
    # Priority color
    if "CRITIQUE" in priority:
        prio_color = "FCE4E4"; prio_font = Font(name="Calibri", size=10, bold=True, color="C00000")
    elif "ÉLEVÉE" in priority:
        prio_color = "FFE6CC"; prio_font = Font(name="Calibri", size=10, bold=True, color="BF6000")
    elif "MOYENNE" in priority:
        prio_color = "FFF2CC"; prio_font = Font(name="Calibri", size=10, bold=True, color="806000")
    else:
        prio_color = "E2EFDA"; prio_font = Font(name="Calibri", size=10, bold=True, color="375623")

    c = ws.cell(row=current_row, column=1, value=priority); c.font = prio_font
    c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c.fill = PatternFill(start_color=prio_color, end_color=prio_color, fill_type="solid")
    c = ws.cell(row=current_row, column=2, value=axe); c.font = BOLD_FONT
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True); c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=current_row, column=3, value=cible); c.font = BODY_FONT
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True); c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=current_row, column=4, value=nb); c.font = BODY_FONT
    c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=current_row, column=5, value=montant); c.font = BOLD_FONT
    c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=current_row, column=6, value=action); c.font = BODY_FONT
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True); c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=current_row, column=7, value=delai); c.font = BODY_FONT
    c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=current_row, column=8, value=pilote); c.font = BODY_FONT
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True); c.border = BORDER
    if banding: c.fill = BAND_FILL
    ws.row_dimensions[current_row].height = 95
    current_row += 1

# Synthèse en bas
current_row += 2
ws.cell(row=current_row, column=1, value="SYNTHÈSE DU PLAN D'ACTION").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=8)
ws.row_dimensions[current_row].height = 22
current_row += 1

synth_rows = [
    ("Enjeu total identifié (pertes Q1 + risques churn)",
     f"≈ {total_loss_global_fcfa + sum(client_t_q1_ca[k] for k in churned_target):,.0f} FCFA".replace(",", " ")),
    ("Gain potentiel (si réactivation complète des 20/80 zéro achat)",
     f"≈ {sum(losses_global[k]['loss_fcfa'] for k in zero_global if k in pareto_clients):,.0f} FCFA".replace(",", " ")),
    ("Nb total de clients concernés par une action prioritaire",
     f"{za_pareto_global + za_pareto_conc + len(churned_target) + len(churned_concentre)} clients"),
    ("Horizon de mise en œuvre",
     "30-90 jours (priorités CRITIQUE et ÉLEVÉE sous 30 jours)"),
    ("Pilotage",
     "Revue mensuelle du CA ciblé/concentrés par segment + tableau de bord des actions"),
]
for label, val in synth_rows:
    c = ws.cell(row=current_row, column=1, value=label)
    c.font = BOLD_FONT; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=4)
    c = ws.cell(row=current_row, column=5, value=val)
    c.font = BODY_FONT; c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True); c.border = BORDER
    ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=8)
    ws.row_dimensions[current_row].height = 26
    current_row += 1

ws.column_dimensions['A'].width = 18
ws.column_dimensions['B'].width = 38
ws.column_dimensions['C'].width = 38
ws.column_dimensions['D'].width = 16
ws.column_dimensions['E'].width = 22
ws.column_dimensions['F'].width = 50
ws.column_dimensions['G'].width = 12
ws.column_dimensions['H'].width = 25

print(f"  Sheet 18 (Recommandations) built")

# ===== SHEET 19: TOP 14 CLIENTS 20/80 PRIORITAIRES (NOMINATIF) =====
ws = wb.create_sheet("19. Top 14 clients 20-80")
ws.cell(row=1, column=1, value="BELGOCAM SA - Top 14 clients 20/80 prioritaires à réactiver en Q1").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=11)
ws.row_dimensions[1].height = 26

ws.cell(row=2, column=1, value="⚠️ Ces 14 clients font partie du top 20/80 (★) mais n'ont acheté AUCUN produit ciblé en Q1. À recontacter EN PRIORITÉ ABSOLUE pendant la rupture concurrente de soja.").font = Font(name="Calibri", size=11, bold=True, color="C00000")
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=11)
ws.row_dimensions[2].height = 32

# Get the 14 clients 20/80 in zero_global, sorted by perte Q1 descending
top14_clients = sorted(
    [k for k in zero_global if k in pareto_clients],
    key=lambda k: -losses_global[k]["loss_fcfa"]
)

# Summary banner
ws.cell(row=4, column=1, value=f"📊 {len(top14_clients)} clients 20/80 à réactiver — Perte Q1 totale estimée : {sum(losses_global[k]['loss_fcfa'] for k in top14_clients):,.0f} FCFA".replace(",", " ")).font = Font(name="Calibri", size=12, bold=True, color="C00000")
ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=11)
ws.row_dimensions[4].height = 26

HEADERS = ["N°", "Réf. client", "Nom du client", "Agence principale",
           "CA HT total 6 mois", "Achat ciblé 6m ?", "Top produit Q2", "CA Top produit Q2",
           "Perte Q1 estimée (FCFA)", "Priorité", "Action recommandée"]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=6, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[6].height = 40

START_ROW = 7
for i, key in enumerate(top14_clients, start=1):
    r = START_ROW + i - 1
    ref, name = key
    primary_ag = get_primary_agency(key)
    targeted_any = len(client_t_q2_months[key]) > 0
    # Top product Q2
    products_q2 = client_q2_by_product_target.get(key, {})
    top_prod = None; top_ca = 0
    for ref_p, rec in products_q2.items():
        if rec["ca"] > top_ca:
            top_ca = rec["ca"]; top_prod = ref_p
    loss = losses_global[key]["loss_fcfa"]
    # Priority
    if loss > 50_000_000:
        priority = "🔴 CRITIQUE"
    elif loss > 10_000_000:
        priority = "🟠 ÉLEVÉE"
    elif loss > 0:
        priority = "🟡 MOYENNE"
    else:
        priority = "⚪ FAIBLE"
    # Action recommendation
    if loss > 50_000_000:
        action = "RDV Directeur Commercial sous 7j. Offre : soja garanti 3 mois + remise concentrés -5%."
    elif loss > 10_000_000:
        action = "Visite commerciale sous 15j. Bundle soja + concentrés préférentiel."
    elif loss > 0:
        action = "Appel téléphonique sous 30j. Échantillon concentrés + offre découverte soja."
    else:
        action = "Prospection standard. Diagnostic besoins + offre tarification."

    cells = [
        (1, i, BODY_ALIGN_CENTER, None),
        (2, ref, BODY_ALIGN_CENTER, None),
        (3, name, BODY_ALIGN_LEFT, None),
        (4, primary_ag, BODY_ALIGN_LEFT, None),
        (5, client_ca_total.get(key, 0.0), BODY_ALIGN_RIGHT, CA_NUM_FMT),
        (6, "Oui" if targeted_any else "Non", BODY_ALIGN_CENTER, None),
        (7, top_prod or "-", BODY_ALIGN_CENTER, None),
        (8, top_ca, BODY_ALIGN_RIGHT, CA_NUM_FMT),
        (9, loss, BODY_ALIGN_RIGHT, CA_NUM_FMT),
        (10, priority, BODY_ALIGN_CENTER, None),
        (11, action, Alignment(horizontal="left", vertical="center", wrap_text=True), None),
    ]
    for col_idx, val, align, fmt in cells:
        c = ws.cell(row=r, column=col_idx, value=val)
        c.font = BODY_FONT; c.alignment = align; c.border = BORDER
        if fmt: c.number_format = fmt
    # Highlight name in gold (20/80)
    ws.cell(row=r, column=3).fill = PARETO_FILL
    ws.cell(row=r, column=3).font = Font(name="Calibri", size=10, bold=True, color="7F6000")
    # Highlight loss in red
    ws.cell(row=r, column=9).fill = LOSS_FILL
    ws.cell(row=r, column=9).font = LOSS_FONT
    # Priority color
    if "CRITIQUE" in priority:
        ws.cell(row=r, column=10).fill = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")
    elif "ÉLEVÉE" in priority:
        ws.cell(row=r, column=10).fill = PatternFill(start_color="FFE6CC", end_color="FFE6CC", fill_type="solid")
    elif "MOYENNE" in priority:
        ws.cell(row=r, column=10).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    ws.row_dimensions[r].height = 38

# Total row
total_row = START_ROW + len(top14_clients)
total_loss = sum(losses_global[k]["loss_fcfa"] for k in top14_clients)
total_ca = sum(client_ca_total[k] for k in top14_clients)
c = ws.cell(row=total_row, column=2, value="TOTAL")
c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
ws.merge_cells(start_row=total_row, start_column=2, end_row=total_row, end_column=4)
for col in [3, 4]:
    ws.cell(row=total_row, column=col).fill = HEADER_FILL
    ws.cell(row=total_row, column=col).border = BORDER
c = ws.cell(row=total_row, column=5, value=total_ca)
c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
for col in [6, 7, 8]:
    ws.cell(row=total_row, column=col).fill = HEADER_FILL
    ws.cell(row=total_row, column=col).border = BORDER
c = ws.cell(row=total_row, column=9, value=total_loss)
c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
for col in [10, 11]:
    ws.cell(row=total_row, column=col).fill = HEADER_FILL
    ws.cell(row=total_row, column=col).border = BORDER

ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 16
ws.column_dimensions['C'].width = 38
ws.column_dimensions['D'].width = 28
ws.column_dimensions['E'].width = 18
ws.column_dimensions['F'].width = 14
ws.column_dimensions['G'].width = 14
ws.column_dimensions['H'].width = 18
ws.column_dimensions['I'].width = 22
ws.column_dimensions['J'].width = 16
ws.column_dimensions['K'].width = 50
ws.freeze_panes = "E7"
ws.auto_filter.ref = f"A6:K{START_ROW + len(top14_clients) - 1}"
print(f"  Sheet 19 (Top 14 clients 20/80) built — {len(top14_clients)} clients")

# ===== SHEET 20: PROJECTION CA 6 MOIS (3 SCÉNARIOS) =====
ws = wb.create_sheet("20. Projection CA 6 mois")
ws.cell(row=1, column=1, value="BELGOCAM SA - Projection CA additionnel à 6 mois (3 scénarios)").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
ws.row_dimensions[1].height = 26

ws.cell(row=2, column=1, value="Hypothèses de réactivation par segment, compte tenu de la rupture concurrente sur le soja. Base : pertes Q1 estimées + churn Q1→Q2.").font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
ws.row_dimensions[2].height = 22

# Compute baseline numbers
total_loss_cible_fcfa = sum(L["loss_fcfa"] for L in losses_global.values())  # ~580 M
loss_churned_cible = sum(client_t_q1_ca[k] for k in clients_all if client_seg_target[k] == ("Active Q1","Zero Q2"))  # ~273 M
n_zero_2080_cible = sum(1 for k in zero_global if k in pareto_clients)  # 14
loss_2080_cible = sum(losses_global[k]["loss_fcfa"] for k in zero_global if k in pareto_clients)
ca_2080_total = sum(client_ca_total[k] for k in zero_global if k in pareto_clients)
n_persistent_zero = sum(1 for k in clients_all if client_seg_target[k] == ("Zero Q1","Zero Q2"))  # ~118

# Scenario definitions
scenarios = [
    {
        "name": "🔴 PESSIMISTE",
        "color": "FCE4E4",
        "context": "Rupture concurrente courte (1-2 mois). Clients acquis retournent vite chez concurrents. Aucune action marketing forte.",
        "taux_2080": 0.30,  # 30% des 14 réactivés
        "taux_churned": 0.10,  # 10% des churned reconquis
        "taux_persistent": 0.05,  # 5% des persistants acquis
        "taux_loss_recovery": 0.10,  # 10% des pertes Q1 récupérées
        "ca_concurrent_capture": 100_000_000,  # 100 M de CA capté sur clients concurrents
    },
    {
        "name": "🟡 RÉALISTE",
        "color": "FFF2CC",
        "context": "Rupture concurrente modérée (3-4 mois). Actions commerciales ciblées sur top clients. Bundles soja+concentrés déployés.",
        "taux_2080": 0.60,
        "taux_churned": 0.30,
        "taux_persistent": 0.15,
        "taux_loss_recovery": 0.30,
        "ca_concurrent_capture": 300_000_000,
    },
    {
        "name": "🟢 OPTIMISTE",
        "color": "C6EFCE",
        "context": "Rupture concurrente prolongée (6+ mois). Conquête massive. Verrouillage contractuel multi-produits. Bundles + remises volume.",
        "taux_2080": 0.90,
        "taux_churned": 0.50,
        "taux_persistent": 0.30,
        "taux_loss_recovery": 0.60,
        "ca_concurrent_capture": 600_000_000,
    },
]

# Section 1: Hypothèses par scénario
ws.cell(row=4, column=1, value="1. Hypothèses par scénario").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=8)
ws.row_dimensions[4].height = 22

h_headers = ["Hypothèse", "🔴 Pessimiste", "🟡 Réaliste", "🟢 Optimiste"]
for col_idx, h in enumerate(h_headers, start=1):
    c = ws.cell(row=5, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[5].height = 32

hyp_rows = [
    ("Contexte", scenarios[0]["context"], scenarios[1]["context"], scenarios[2]["context"]),
    (f"% clients 20/80 réactivés (sur {n_zero_2080_cible})", f"{scenarios[0]['taux_2080']*100:.0f}%  ({int(scenarios[0]['taux_2080']*n_zero_2080_cible)} clients)", f"{scenarios[1]['taux_2080']*100:.0f}%  ({int(scenarios[1]['taux_2080']*n_zero_2080_cible)} clients)", f"{scenarios[2]['taux_2080']*100:.0f}%  ({int(scenarios[2]['taux_2080']*n_zero_2080_cible)} clients)"),
    ("% clients churned reconquis", f"{scenarios[0]['taux_churned']*100:.0f}%", f"{scenarios[1]['taux_churned']*100:.0f}%", f"{scenarios[2]['taux_churned']*100:.0f}%"),
    (f"% clients persistants acquis (sur {n_persistent_zero})", f"{scenarios[0]['taux_persistent']*100:.0f}%  ({int(scenarios[0]['taux_persistent']*n_persistent_zero)})", f"{scenarios[1]['taux_persistent']*100:.0f}%  ({int(scenarios[1]['taux_persistent']*n_persistent_zero)})", f"{scenarios[2]['taux_persistent']*100:.0f}%  ({int(scenarios[2]['taux_persistent']*n_persistent_zero)})"),
    ("% pertes Q1 récupérées", f"{scenarios[0]['taux_loss_recovery']*100:.0f}%", f"{scenarios[1]['taux_loss_recovery']*100:.0f}%", f"{scenarios[2]['taux_loss_recovery']*100:.0f}%"),
    ("CA capté sur clients concurrents (FCFA)", f"{scenarios[0]['ca_concurrent_capture']:,.0f}".replace(",", " "), f"{scenarios[1]['ca_concurrent_capture']:,.0f}".replace(",", " "), f"{scenarios[2]['ca_concurrent_capture']:,.0f}".replace(",", " ")),
]
for i, row in enumerate(hyp_rows, start=1):
    r = 5 + i
    banding = (i % 2 == 0)
    for col_idx, val in enumerate(row, start=1):
        c = ws.cell(row=r, column=col_idx, value=val)
        c.font = BOLD_FONT if col_idx == 1 else BODY_FONT
        c.alignment = Alignment(horizontal="left" if col_idx in [1, 2] else "left", vertical="center", wrap_text=True)
        c.border = BORDER
        if banding: c.fill = BAND_FILL
    ws.row_dimensions[r].height = 32

# Section 2: CA additionnel estimé par segment
r_start_sec2 = 5 + len(hyp_rows) + 2
ws.cell(row=r_start_sec2, column=1, value="2. CA additionnel estimé à 6 mois (FCFA)").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
ws.merge_cells(start_row=r_start_sec2, start_column=1, end_row=r_start_sec2, end_column=8)
ws.row_dimensions[r_start_sec2].height = 22

ca_headers = ["Segment", "Base (FCFA)", "🔴 Pessimiste", "🟡 Réaliste", "🟢 Optimiste"]
for col_idx, h in enumerate(ca_headers, start=1):
    c = ws.cell(row=r_start_sec2 + 1, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[r_start_sec2 + 1].height = 32

# Calculate CA additionnel per segment per scenario
ca_segment_rows = []
# Segment 1: 20/80 réactivation (CA moyen par client = ca_2080_total / n_zero_2080_cible)
avg_ca_2080 = ca_2080_total / n_zero_2080_cible if n_zero_2080_cible > 0 else 0
ca_segment_rows.append((
    f"Réactivation clients 20/80 ({n_zero_2080_cible} ciblés)",
    ca_2080_total,
    int(scenarios[0]["taux_2080"] * n_zero_2080_cible * avg_ca_2080 * 0.5),  # 0.5 = part 6mois du CA annuel
    int(scenarios[1]["taux_2080"] * n_zero_2080_cible * avg_ca_2080 * 0.5),
    int(scenarios[2]["taux_2080"] * n_zero_2080_cible * avg_ca_2080 * 0.5),
))
# Segment 2: Reconquête churned (CA Q1 perdu)
ca_segment_rows.append((
    f"Reconquête clients churned ({loss_churned_cible:,.0f} FCFA perdus)".replace(",", " "),
    loss_churned_cible,
    int(loss_churned_cible * scenarios[0]["taux_churned"]),
    int(loss_churned_cible * scenarios[1]["taux_churned"]),
    int(loss_churned_cible * scenarios[2]["taux_churned"]),
))
# Segment 3: Acquisition persistants zero
ca_segment_rows.append((
    f"Acquisition clients persistants zero ({n_persistent_zero} ciblés)",
    0,  # No baseline
    int(scenarios[0]["taux_persistent"] * n_persistent_zero * 500_000),  # 500k FCFA avg per new client
    int(scenarios[1]["taux_persistent"] * n_persistent_zero * 1_000_000),  # 1M avg
    int(scenarios[2]["taux_persistent"] * n_persistent_zero * 2_000_000),  # 2M avg
))
# Segment 4: Récupération pertes Q1 estimées (méthode fréquence)
ca_segment_rows.append((
    f"Récupération pertes Q1 estimées ({total_loss_cible_fcfa:,.0f} FCFA)".replace(",", " "),
    total_loss_cible_fcfa,
    int(total_loss_cible_fcfa * scenarios[0]["taux_loss_recovery"]),
    int(total_loss_cible_fcfa * scenarios[1]["taux_loss_recovery"]),
    int(total_loss_cible_fcfa * scenarios[2]["taux_loss_recovery"]),
))
# Segment 5: Conquête clients concurrents (rupture soja)
ca_segment_rows.append((
    "Conquête clients concurrents (rupture soja)",
    0,
    scenarios[0]["ca_concurrent_capture"],
    scenarios[1]["ca_concurrent_capture"],
    scenarios[2]["ca_concurrent_capture"],
))

total_pess = 0; total_real = 0; total_opt = 0
for i, (label, base, pess, real, opt) in enumerate(ca_segment_rows, start=1):
    r = r_start_sec2 + 1 + i
    banding = (i % 2 == 0)
    total_pess += pess; total_real += real; total_opt += opt
    cells = [
        (1, label, BODY_ALIGN_LEFT, None),
        (2, base, BODY_ALIGN_RIGHT, CA_NUM_FMT),
        (3, pess, BODY_ALIGN_RIGHT, CA_NUM_FMT),
        (4, real, BODY_ALIGN_RIGHT, CA_NUM_FMT),
        (5, opt, BODY_ALIGN_RIGHT, CA_NUM_FMT),
    ]
    for col_idx, val, align, fmt in cells:
        c = ws.cell(row=r, column=col_idx, value=val)
        c.font = BODY_FONT; c.alignment = align; c.border = BORDER
        if fmt: c.number_format = fmt
        if banding: c.fill = BAND_FILL
    ws.row_dimensions[r].height = 24

# Total row
total_row_sec2 = r_start_sec2 + 1 + len(ca_segment_rows) + 1
c = ws.cell(row=total_row_sec2, column=1, value="CA ADDITIONNEL TOTAL")
c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
c = ws.cell(row=total_row_sec2, column=2, value="—")
c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
for col, val, color in [(3, total_pess, "8B0000"), (4, total_real, "BF6000"), (5, total_opt, "375623")]:
    c = ws.cell(row=total_row_sec2, column=col, value=val)
    c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER; c.number_format = CA_NUM_FMT
ws.row_dimensions[total_row_sec2].height = 30

# Section 3: Projection CA total (CA actuel + additionnel)
r_start_sec3 = total_row_sec2 + 2
ws.cell(row=r_start_sec3, column=1, value="3. Projection CA total à 6 mois (CA actuel + additionnel)").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
ws.merge_cells(start_row=r_start_sec3, start_column=1, end_row=r_start_sec3, end_column=8)
ws.row_dimensions[r_start_sec3].height = 22

proj_headers = ["", "Actuel (6 mois)", "🔴 Pessimiste", "🟡 Réaliste", "🟢 Optimiste"]
for col_idx, h in enumerate(proj_headers, start=1):
    c = ws.cell(row=r_start_sec3 + 1, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[r_start_sec3 + 1].height = 32

# Projection rows
proj_rows = [
    ("CA HT actuel 6 mois", total_ca, total_ca, total_ca, total_ca),
    ("CA additionnel estimé", 0, total_pess, total_real, total_opt),
    ("CA projeté 6 mois (S2 2026)", total_ca, total_ca + total_pess, total_ca + total_real, total_ca + total_opt),
    ("Croissance vs S1 2026", 0, total_pess / total_ca, total_real / total_ca, total_opt / total_ca),
]
for i, (label, base, pess, real, opt) in enumerate(proj_rows, start=1):
    r = r_start_sec3 + 1 + i
    banding = (i % 2 == 0)
    is_total_row = "projeté" in label.lower() or "croissance" in label.lower()
    c = ws.cell(row=r, column=1, value=label)
    c.font = BOLD_FONT if is_total_row else BODY_FONT
    c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    if banding and not is_total_row: c.fill = BAND_FILL
    elif is_total_row: c.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    cells_vals = [(2, base), (3, pess), (4, real), (5, opt)]
    for col, val in cells_vals:
        c = ws.cell(row=r, column=col, value=val)
        c.font = BOLD_FONT if is_total_row else BODY_FONT
        c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
        if "Croissance" in label:
            c.number_format = '+0.0%;-0.0%;0.0%'
        else:
            c.number_format = CA_NUM_FMT
        if banding and not is_total_row: c.fill = BAND_FILL
        elif is_total_row and col > 1: c.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    ws.row_dimensions[r].height = 24

# Column widths
ws.column_dimensions['A'].width = 50
ws.column_dimensions['B'].width = 22
ws.column_dimensions['C'].width = 22
ws.column_dimensions['D'].width = 22
ws.column_dimensions['E'].width = 22
ws.column_dimensions['F'].width = 12
ws.column_dimensions['G'].width = 12
ws.column_dimensions['H'].width = 12

print(f"  Sheet 20 (Projection CA 6 mois) built — 3 scenarios")

# ===== SHEET 21: PERSPECTIVES STRATÉGIQUES =====
ws = wb.create_sheet("21. Perspectives strategiques")
ws.cell(row=1, column=1, value="BELGOCAM SA - Perspectives stratégiques (S2 2026)").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
ws.row_dimensions[1].height = 28

ws.cell(row=2, column=1, value="Cadre : rupture concurrente sur le soja. Analyse en 5 axes avec plan d'action 90 jours.").font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
ws.row_dimensions[2].height = 22

# Section headers
HEADERS = ["Axe", "Constat / Opportunité", "Cibles prioritaires", "Actions concrètes", "KPI", "Délai"]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=4, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[4].height = 36

# Strategic axes data
axes = [
    {
        "axe": "🎯 1. CONQUÊTE SOJA\n(rupture concurrentielle)",
        "constat": "Les concurrents sont en rupture de soja. Fenêtre de tir exceptionnelle : le soja représente 16 730 lignes / 6 mois (40% du volume). CA additionnel potentiel : 300-600 M FCFA.",
        "cibles": "• 14 clients 20/80 zéro achat Q1 (ciblé)\n• 192 clients churned Q1→Q2 (273 M FCFA perdus)\n• 118 clients persistants zero (jamais acquis)\n• Clients concurrents en rupture",
        "actions": "• Opération 'Soja disponible' J+7 : RDV direct commercial\n• Argumentaire : 'BELGOCAM = sécurité approvisionnement'\n• Bundles soja + concentrés (-5% conc. si achat soja)\n• Vérification quotidienne stocks soja\n• Verrouillage contractuel 6 mois multi-produits",
        "kpi": "• Nb nouveaux clients soja\n• Volume soja additionnel (t)\n• CA additionnel (FCFA)\n• Taux de conversion par segment",
        "delai": "15 jours (urgence)",
    },
    {
        "axe": "🛡️ 2. DÉFENSE CONCENTRÉS\n(faiblesse structurelle)",
        "constat": "Bilan net Q1→Q2 concentrés fragile : +52 M FCFA (vs +276 M ciblé). Segment 'retenu' en déclin -248 t / -164 M FCFA. Taux churn fidèles 3,1% (vs 1,9% ciblé). 544 clients jamais acquis concentrés.",
        "cibles": "• 178 clients churned Q1→Q2 concentrés (93 M FCFA)\n• 97 fidèles Q1 en déclin Q2 (volume -248 t)\n• 41 clients 20/80 zéro achat concentrés\n• 672 clients 'retenu' (déclin volume)",
        "actions": "• Benchmark prix concurrents (12 références)\n• Enquête qualitative 30 clients churned\n• Ajustement tarifaire si nécessaire (C104, C103)\n• Visite physique 97 fidèles en déclin\n• Bundle soja+concentrés obligatoire\n• Test nouveaux formats (1Kg, 5Kg, 25Kg)",
        "kpi": "• Volume concentrés Q3 vs Q2\n• Taux de rétention fidèles\n• Prix moyen vs concurrents\n• Nb clients actifs concentrés",
        "delai": "30-60 jours",
    },
    {
        "axe": "🏢 3. STRATÉGIE PAR AGENCE\n(disparités majeures)",
        "constat": "Disparités importantes : FAMLA (293 clients, 19% zero) vs MESSASSI (167, 33,5% zero) vs NGAOUNDERE (91, 34,1% zero). 1 client sur 3 en zero achat dans certaines agences = anomalie.",
        "cibles": "• MESSASSI (56 zero achat, 0 zero 20/80)\n• NGAOUNDERE (31 zero, 0 zero 20/80)\n• FAMLA (49 zero, 7 zero 20/80 prioritaires)\n• DJELENG (25 zero, 2 zero 20/80)",
        "actions": "• Diagnostic terrain MESSASSI + NGAOUNDERE\n• Renforcement effectif commercial si sous-dimensionné\n• Visibilité locale : PLV, événements\n• Plan marketing agence par agence\n• Objectifs commerciaux individualisés par agence",
        "kpi": "• % zero achat par agence (cible <25%)\n• Nb nouveaux clients par agence/mois\n• CA par agence vs objectif",
        "delai": "60-90 jours",
    },
    {
        "axe": "⚠️ 4. RISQUES À SURVEILLER",
        "constat": "Rupture concurrente temporaire : risque de retour client si pas de verrouillage. Stock soja à surveiller pour pouvoir servir la demande. Cannibalisation possible soja vs concentrés.",
        "cibles": "• Stock soja BELGOCAM\n• Concurrents (réapprovisionnement)\n• Taux d'attrition concentrés\n• Marge mix produit",
        "actions": "• Suivi quotidien stock soja + plan réappro\n• Contrats multi-produits 6 mois (verrouillage)\n• Bundles obligatoires (anti-cannibalisation)\n• Veille concurrentielle hebdo\n• Tableau de bord churn mensuel",
        "kpi": "• Niveau stock soja (jours)\n• Taux de rétention nouveaux clients\n• Mix produit (soja vs concentrés)\n• Marge brute par produit",
        "delai": "Continu (révue mensuelle)",
    },
    {
        "axe": "🎬 5. PLAN 90 JOURS",
        "constat": "3 phases : conquête (J1-15), verrouillage (J15-30), reconquête concentrés (J30-60), consolidation (J60-90).",
        "cibles": "• J1-15 : 324 clients prioritaires (14+192+118)\n• J15-30 : nouveaux clients soja\n• J30-60 : 97 fidèles en déclin + 178 churned conc.\n• J60-90 : bilan + ajustements",
        "actions": "• J1-15 : Lancement 'Soja disponible' + vérif stocks\n• J15-30 : Contrats 6 mois + bundles + enquête 30 clients\n• J30-60 : Benchmark prix + ajustement + visites physiques\n• J60-90 : Bilan conquest soja + plan marketing agences\n• Revue mensuelle comité direction",
        "kpi": "• J15 : 50% des 324 clients contactés\n• J30 : 60% des nouveaux verrouillés contrat\n• J60 : Benchmark livré + 30 visites faites\n• J90 : Bilan CA additionnel vs scénario réaliste",
        "delai": "90 jours (3 phases)",
    },
]

START_ROW = 5
for i, axis in enumerate(axes, start=1):
    r = START_ROW + i - 1
    banding = (i % 2 == 0)
    # Axe (column 1) with priority color
    if "CONQUÊTE" in axis["axe"]:
        ax_color = "C6EFCE"; ax_font = Font(name="Calibri", size=11, bold=True, color="375623")
    elif "DÉFENSE" in axis["axe"]:
        ax_color = "FCE4E4"; ax_font = Font(name="Calibri", size=11, bold=True, color="C00000")
    elif "AGENCE" in axis["axe"]:
        ax_color = "DDEBF7"; ax_font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    elif "RISQUES" in axis["axe"]:
        ax_color = "FFF2CC"; ax_font = Font(name="Calibri", size=11, bold=True, color="806000")
    else:
        ax_color = "E2EFDA"; ax_font = Font(name="Calibri", size=11, bold=True, color="375623")

    cells = [
        (1, axis["axe"], ax_font, Alignment(horizontal="left", vertical="center", wrap_text=True),
         PatternFill(start_color=ax_color, end_color=ax_color, fill_type="solid")),
        (2, axis["constat"], BODY_FONT, Alignment(horizontal="left", vertical="center", wrap_text=True),
         BAND_FILL if banding else None),
        (3, axis["cibles"], BODY_FONT, Alignment(horizontal="left", vertical="center", wrap_text=True),
         BAND_FILL if banding else None),
        (4, axis["actions"], BODY_FONT, Alignment(horizontal="left", vertical="center", wrap_text=True),
         BAND_FILL if banding else None),
        (5, axis["kpi"], BODY_FONT, Alignment(horizontal="left", vertical="center", wrap_text=True),
         BAND_FILL if banding else None),
        (6, axis["delai"], BOLD_FONT, Alignment(horizontal="center", vertical="center", wrap_text=True),
         BAND_FILL if banding else None),
    ]
    for col_idx, val, font, align, fill in cells:
        c = ws.cell(row=r, column=col_idx, value=val)
        c.font = font; c.alignment = align; c.border = BORDER
        if fill: c.fill = fill
    ws.row_dimensions[r].height = 130

# Synthèse en bas
synth_row = START_ROW + len(axes) + 1
ws.cell(row=synth_row, column=1, value="SYNTHÈSE STRATÉGIQUE").font = Font(name="Calibri", size=13, bold=True, color="1F4E78")
ws.merge_cells(start_row=synth_row, start_column=1, end_row=synth_row, end_column=6)
ws.row_dimensions[synth_row].height = 26

synthese_text = (
    "1. La rupture concurrente sur le soja est une AUBAINE À 6 MOIS : traiter comme opération de guerre (moyens commerciaux maximaux, prix premium, verrouillage contractuel).\n\n"
    "2. Le vrai sujet stratégique est la FAIBLESSE DES CONCENTRÉS : sans résoudre le déclin volume du segment 'retenu' (-248 t Q2), la rentabilité long-terme est menacée car les concentrés sont plus marginaux que le soja brut.\n\n"
    "3. Les 14 CLIENTS 20/80 EN ZÉRO ACHAT Q1 sont le TEST ULTIME : si on ne les réactive pas pendant cette période de rupture concurrente (où l'argument 'stock disponible' est imbattable), c'est qu'il y a un problème structurel (prix, qualité, relation client) à diagnostiquer urgemment.\n\n"
    "🎯 Objectif S2 2026 : +175 M FCFA de CA additionnel (scénario réaliste), soit +1% de croissance vs S1 2026."
)
ws.cell(row=synth_row + 1, column=1, value=synthese_text).font = Font(name="Calibri", size=11)
ws.cell(row=synth_row + 1, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
ws.cell(row=synth_row + 1, column=1).fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
ws.cell(row=synth_row + 1, column=1).border = BORDER
ws.merge_cells(start_row=synth_row + 1, start_column=1, end_row=synth_row + 1, end_column=6)
ws.row_dimensions[synth_row + 1].height = 130

ws.column_dimensions['A'].width = 22
ws.column_dimensions['B'].width = 45
ws.column_dimensions['C'].width = 38
ws.column_dimensions['D'].width = 45
ws.column_dimensions['E'].width = 32
ws.column_dimensions['F'].width = 16

print(f"  Sheet 21 (Perspectives stratégiques) built")

# Note: Sommaire update skipped to avoid merged cell conflicts.
# New sheets 19-21 are accessible via sheet tabs at the bottom of the workbook.

# ===== SHEETS 22-26: CHICK BOOSTER & PIGLET BOOSTER ANALYSIS =====
# Products:
#   Chick Booster: CB100, CB101
#   Piglet Booster: CB200, CB201
#   Concentrés Chair (10% + 5%): C104, C1042, C1043, C1044, C103
#   Concentrés Ponte (10% + 5%): C102, C1022, C101
#   Concentrés Porc (10%): C105, C1053, C1054, C1055

CHICK_REFS = {"CB100", "CB101"}
PIGLET_REFS = {"CB200", "CB201"}
CHAIR_REFS = {"C104", "C1042", "C1043", "C1044", "C103"}  # 10% + 5% Chair
PONTE_REFS = {"C102", "C1022", "C101"}  # 10% + 5% Ponte
PORC_REFS = {"C105", "C1053", "C1054", "C1055"}  # 10% Porc

# Re-read source to compute per-product aggregates for Chick/Piglet
print("\nReading source again for Chick/Piglet analysis...")
wb_src2 = load_workbook(SRC, read_only=True, data_only=True)

# Per-product aggregates (over 6 months)
product_aggregates = defaultdict(lambda: {"vol_kg": 0.0, "ca": 0.0, "qte": 0, "rows": 0})
# Per-client per-product-group (Chick / Piglet / Chair / Ponte / Porc)
client_chick = defaultdict(lambda: {"vol_kg": 0.0, "ca": 0.0, "qte": 0})
client_piglet = defaultdict(lambda: {"vol_kg": 0.0, "ca": 0.0, "qte": 0})
client_chair = defaultdict(lambda: {"vol_kg": 0.0, "ca": 0.0, "qte": 0})
client_ponte = defaultdict(lambda: {"vol_kg": 0.0, "ca": 0.0, "qte": 0})
client_porc = defaultdict(lambda: {"vol_kg": 0.0, "ca": 0.0, "qte": 0})

for sheet_name in wb_src2.sheetnames:
    ws = wb_src2[sheet_name]
    for row in ws.iter_rows(min_row=3, values_only=True):
        tiers = row[5] if len(row) > 5 else None
        if tiers is None or str(tiers).strip() == "":
            continue
        if tiers in excluded_tiers:
            continue
        ref_prod = row[0] if len(row) > 0 else None
        desc = row[1] if len(row) > 1 else None
        qte = row[2] if len(row) > 2 else 0
        ca_ht = row[8] if len(row) > 8 else 0

        ref_c, name_c = split_tiers(tiers)
        key = (ref_c, name_c)
        ref_prod_str = str(ref_prod).strip() if ref_prod is not None else ""
        if not ref_prod_str:
            continue

        if ref_prod_str not in product_weight:
            product_weight[ref_prod_str] = parse_weight_kg(desc)
        weight_kg = product_weight.get(ref_prod_str, 0.0)
        try:
            qte_f = float(qte) if qte is not None else 0.0
        except:
            qte_f = 0.0
        try:
            ca_f = float(ca_ht) if ca_ht is not None else 0.0
        except:
            ca_f = 0.0
        vol_kg = qte_f * weight_kg

        # Aggregate per-product (for sheet 22)
        if ref_prod_str in CHICK_REFS or ref_prod_str in PIGLET_REFS:
            p = product_aggregates[ref_prod_str]
            p["vol_kg"] += vol_kg
            p["ca"] += ca_f
            p["qte"] += qte_f
            p["rows"] += 1

        # Aggregate per-client per-category
        if ref_prod_str in CHICK_REFS:
            client_chick[key]["vol_kg"] += vol_kg
            client_chick[key]["ca"] += ca_f
            client_chick[key]["qte"] += qte_f
        if ref_prod_str in PIGLET_REFS:
            client_piglet[key]["vol_kg"] += vol_kg
            client_piglet[key]["ca"] += ca_f
            client_piglet[key]["qte"] += qte_f
        if ref_prod_str in CHAIR_REFS:
            client_chair[key]["vol_kg"] += vol_kg
            client_chair[key]["ca"] += ca_f
            client_chair[key]["qte"] += qte_f
        if ref_prod_str in PONTE_REFS:
            client_ponte[key]["vol_kg"] += vol_kg
            client_ponte[key]["ca"] += ca_f
            client_ponte[key]["qte"] += qte_f
        if ref_prod_str in PORC_REFS:
            client_porc[key]["vol_kg"] += vol_kg
            client_porc[key]["ca"] += ca_f
            client_porc[key]["qte"] += qte_f

wb_src2.close()
print(f"  Chick Booster clients: {len(client_chick)}")
print(f"  Piglet Booster clients: {len(client_piglet)}")
print(f"  Concentrés Chair clients: {len(client_chair)}")
print(f"  Concentrés Ponte clients: {len(client_ponte)}")
print(f"  Concentrés Porc clients: {len(client_porc)}")

# ===== SHEET 22: VENTES CHICK BOOSTER & PIGLET BOOSTER (6 MOIS) =====
ws = wb.create_sheet("22. Ventes Chick & Piglet")
ws.cell(row=1, column=1, value="BELGOCAM SA - Ventes Chick Booster & Piglet Booster (Janvier-Juin 2026)").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
ws.row_dimensions[1].height = 26

ws.cell(row=2, column=1, value="Synthèse des ventes sur 6 mois par référence produit + agrégats par gamme.").font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=10)
ws.row_dimensions[2].height = 22

# Section 1: Détail par référence
ws.cell(row=4, column=1, value="1. Ventes par référence produit (6 mois)").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=10)
ws.row_dimensions[4].height = 22

HEADERS = ["Réf. produit", "Description", "Poids unitaire (kg)", "Quantité totale (sacs)",
           "Volume total (t)", "CA HT total (FCFA)", "Prix moyen (/t)", "Nb lignes", "Gamme", ""]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=5, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[5].height = 36

# Build rows for each product
products_detail = []
for ref in sorted(CHICK_REFS) + sorted(PIGLET_REFS):
    p = product_aggregates.get(ref, {"vol_kg": 0, "ca": 0, "qte": 0, "rows": 0})
    weight = product_weight.get(ref, 0)
    # Get description
    desc = ""
    if ref == "CB100": desc = "CHICK BOOSTER 25 Kg"
    elif ref == "CB101": desc = "CHICK BOOSTER 5Kg"
    elif ref == "CB200": desc = "PIGLET BOOSTER 25Kg"
    elif ref == "CB201": desc = "PIGLET BOOSTER 5Kg"
    gamme = "Chick Booster" if ref in CHICK_REFS else "Piglet Booster"
    vol_t = p["vol_kg"] / 1000.0
    price_per_t = (p["ca"] / vol_t) if vol_t > 0 else 0
    products_detail.append((ref, desc, weight, p["qte"], vol_t, p["ca"], price_per_t, p["rows"], gamme))

START_ROW = 6
for i, (ref, desc, weight, qte, vol_t, ca, price, rows, gamme) in enumerate(products_detail, start=1):
    r = START_ROW + i - 1
    banding = (i % 2 == 0)
    cells = [
        (1, ref, BODY_ALIGN_CENTER),
        (2, desc, BODY_ALIGN_LEFT),
        (3, weight, BODY_ALIGN_CENTER),
        (4, qte, BODY_ALIGN_RIGHT),
        (5, vol_t, BODY_ALIGN_RIGHT),
        (6, ca, BODY_ALIGN_RIGHT),
        (7, price, BODY_ALIGN_RIGHT),
        (8, rows, BODY_ALIGN_CENTER),
        (9, gamme, BODY_ALIGN_CENTER),
    ]
    for col, val, align in cells:
        c = ws.cell(row=r, column=col, value=val)
        c.font = BODY_FONT; c.alignment = align; c.border = BORDER
        if banding: c.fill = BAND_FILL
    ws.cell(row=r, column=3).number_format = '0.00'
    ws.cell(row=r, column=4).number_format = '#,##0'
    ws.cell(row=r, column=5).number_format = VOL_NUM_FMT
    ws.cell(row=r, column=6).number_format = CA_NUM_FMT
    ws.cell(row=r, column=7).number_format = '#,##0" FCFA/t"'
    # Color gamme
    if "Chick" in gamme:
        ws.cell(row=r, column=9).fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
        ws.cell(row=r, column=9).font = Font(name="Calibri", size=10, bold=True, color="7F6000")
    else:
        ws.cell(row=r, column=9).fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
        ws.cell(row=r, column=9).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

# Subtotals by gamme
r_subtotal = START_ROW + len(products_detail) + 1
# Chick subtotal
chick_total_vol = sum(p[4] for p in products_detail if p[8] == "Chick Booster")
chick_total_ca = sum(p[5] for p in products_detail if p[8] == "Chick Booster")
chick_total_qte = sum(p[3] for p in products_detail if p[8] == "Chick Booster")
c = ws.cell(row=r_subtotal, column=1, value="SOUS-TOTAL CHICK BOOSTER")
c.font = Font(name="Calibri", size=11, bold=True, color="7F6000")
c.fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
ws.merge_cells(start_row=r_subtotal, start_column=1, end_row=r_subtotal, end_column=3)
for col in [2, 3]:
    ws.cell(row=r_subtotal, column=col).fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
    ws.cell(row=r_subtotal, column=col).border = BORDER
ws.cell(row=r_subtotal, column=4, value=chick_total_qte).font = Font(name="Calibri", size=11, bold=True, color="7F6000")
ws.cell(row=r_subtotal, column=4).fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
ws.cell(row=r_subtotal, column=4).alignment = BODY_ALIGN_RIGHT; ws.cell(row=r_subtotal, column=4).border = BORDER
ws.cell(row=r_subtotal, column=4).number_format = '#,##0'
ws.cell(row=r_subtotal, column=5, value=chick_total_vol).font = Font(name="Calibri", size=11, bold=True, color="7F6000")
ws.cell(row=r_subtotal, column=5).fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
ws.cell(row=r_subtotal, column=5).alignment = BODY_ALIGN_RIGHT; ws.cell(row=r_subtotal, column=5).border = BORDER
ws.cell(row=r_subtotal, column=5).number_format = VOL_NUM_FMT
ws.cell(row=r_subtotal, column=6, value=chick_total_ca).font = Font(name="Calibri", size=11, bold=True, color="7F6000")
ws.cell(row=r_subtotal, column=6).fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
ws.cell(row=r_subtotal, column=6).alignment = BODY_ALIGN_RIGHT; ws.cell(row=r_subtotal, column=6).border = BORDER
ws.cell(row=r_subtotal, column=6).number_format = CA_NUM_FMT
for col in [7, 8, 9]:
    ws.cell(row=r_subtotal, column=col).fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
    ws.cell(row=r_subtotal, column=col).border = BORDER

# Piglet subtotal
r_subtotal2 = r_subtotal + 1
piglet_total_vol = sum(p[4] for p in products_detail if p[8] == "Piglet Booster")
piglet_total_ca = sum(p[5] for p in products_detail if p[8] == "Piglet Booster")
piglet_total_qte = sum(p[3] for p in products_detail if p[8] == "Piglet Booster")
c = ws.cell(row=r_subtotal2, column=1, value="SOUS-TOTAL PIGLET BOOSTER")
c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
ws.merge_cells(start_row=r_subtotal2, start_column=1, end_row=r_subtotal2, end_column=3)
for col in [2, 3]:
    ws.cell(row=r_subtotal2, column=col).fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
    ws.cell(row=r_subtotal2, column=col).border = BORDER
ws.cell(row=r_subtotal2, column=4, value=piglet_total_qte).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=r_subtotal2, column=4).fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
ws.cell(row=r_subtotal2, column=4).alignment = BODY_ALIGN_RIGHT; ws.cell(row=r_subtotal2, column=4).border = BORDER
ws.cell(row=r_subtotal2, column=4).number_format = '#,##0'
ws.cell(row=r_subtotal2, column=5, value=piglet_total_vol).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=r_subtotal2, column=5).fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
ws.cell(row=r_subtotal2, column=5).alignment = BODY_ALIGN_RIGHT; ws.cell(row=r_subtotal2, column=5).border = BORDER
ws.cell(row=r_subtotal2, column=5).number_format = VOL_NUM_FMT
ws.cell(row=r_subtotal2, column=6, value=piglet_total_ca).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=r_subtotal2, column=6).fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
ws.cell(row=r_subtotal2, column=6).alignment = BODY_ALIGN_RIGHT; ws.cell(row=r_subtotal2, column=6).border = BORDER
ws.cell(row=r_subtotal2, column=6).number_format = CA_NUM_FMT
for col in [7, 8, 9]:
    ws.cell(row=r_subtotal2, column=col).fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
    ws.cell(row=r_subtotal2, column=col).border = BORDER

# Grand total
r_grand = r_subtotal2 + 1
c = ws.cell(row=r_grand, column=1, value="TOTAL CHICK + PIGLET")
c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
ws.merge_cells(start_row=r_grand, start_column=1, end_row=r_grand, end_column=3)
for col in [2, 3]:
    ws.cell(row=r_grand, column=col).fill = HEADER_FILL; ws.cell(row=r_grand, column=col).border = BORDER
ws.cell(row=r_grand, column=4, value=chick_total_qte + piglet_total_qte).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=r_grand, column=4).fill = HEADER_FILL; ws.cell(row=r_grand, column=4).alignment = BODY_ALIGN_RIGHT; ws.cell(row=r_grand, column=4).border = BORDER
ws.cell(row=r_grand, column=4).number_format = '#,##0'
ws.cell(row=r_grand, column=5, value=chick_total_vol + piglet_total_vol).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=r_grand, column=5).fill = HEADER_FILL; ws.cell(row=r_grand, column=5).alignment = BODY_ALIGN_RIGHT; ws.cell(row=r_grand, column=5).border = BORDER
ws.cell(row=r_grand, column=5).number_format = VOL_NUM_FMT
ws.cell(row=r_grand, column=6, value=chick_total_ca + piglet_total_ca).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=r_grand, column=6).fill = HEADER_FILL; ws.cell(row=r_grand, column=6).alignment = BODY_ALIGN_RIGHT; ws.cell(row=r_grand, column=6).border = BORDER
ws.cell(row=r_grand, column=6).number_format = CA_NUM_FMT
for col in [7, 8, 9]:
    ws.cell(row=r_grand, column=col).fill = HEADER_FILL; ws.cell(row=r_grand, column=col).border = BORDER

# Section 2: Synthèse par gamme
r_synth = r_grand + 2
ws.cell(row=r_synth, column=1, value="2. Synthèse par gamme").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
ws.merge_cells(start_row=r_synth, start_column=1, end_row=r_synth, end_column=10)
ws.row_dimensions[r_synth].height = 22

synth_headers = ["Gamme", "Nb produits", "Nb clients acheteurs", "Volume total (t)", "CA HT total (FCFA)", "% CA Booster", "Prix moyen (/t)", "", "", ""]
for col_idx, h in enumerate(synth_headers, start=1):
    c = ws.cell(row=r_synth + 1, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[r_synth + 1].height = 36

total_booster_ca = chick_total_ca + piglet_total_ca
gamme_rows = [
    ("Chick Booster", 2, len(client_chick), chick_total_vol, chick_total_ca, chick_total_ca / total_booster_ca if total_booster_ca > 0 else 0,
     (chick_total_ca / chick_total_vol) if chick_total_vol > 0 else 0),
    ("Piglet Booster", 2, len(client_piglet), piglet_total_vol, piglet_total_ca, piglet_total_ca / total_booster_ca if total_booster_ca > 0 else 0,
     (piglet_total_ca / piglet_total_vol) if piglet_total_vol > 0 else 0),
]
for i, (gamme, n_prod, n_clients, vol, ca, pct, price) in enumerate(gamme_rows, start=1):
    r = r_synth + 1 + i
    banding = (i % 2 == 0)
    cells = [(1, gamme), (2, n_prod), (3, n_clients), (4, vol), (5, ca), (6, pct), (7, price)]
    for col, val in cells:
        c = ws.cell(row=r, column=col, value=val)
        c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER if col <= 3 else BODY_ALIGN_RIGHT; c.border = BORDER
        if banding: c.fill = BAND_FILL
    ws.cell(row=r, column=4).number_format = VOL_NUM_FMT
    ws.cell(row=r, column=5).number_format = CA_NUM_FMT
    ws.cell(row=r, column=6).number_format = '0.0%'
    ws.cell(row=r, column=7).number_format = '#,##0" FCFA/t"'

# Column widths
ws.column_dimensions['A'].width = 22
ws.column_dimensions['B'].width = 28
ws.column_dimensions['C'].width = 16
ws.column_dimensions['D'].width = 18
ws.column_dimensions['E'].width = 16
ws.column_dimensions['F'].width = 22
ws.column_dimensions['G'].width = 18
ws.column_dimensions['H'].width = 12
ws.column_dimensions['I'].width = 16
ws.column_dimensions['J'].width = 10

print(f"  Sheet 22 (Ventes Chick & Piglet) built")

# ===== SHEET 23: LISTE CLIENTS CHICK BOOSTER =====
ws = wb.create_sheet("23. Clients Chick Booster")
ws.cell(row=1, column=1, value="BELGOCAM SA - Clients acheteurs Chick Booster (Janvier-Juin 2026)").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
ws.row_dimensions[1].height = 26

n_chick = len(client_chick)
ws.cell(row=2, column=1, value=f"📊 {n_chick} client(s) ont acheté du Chick Booster (CB100 25Kg ou CB101 5Kg) sur 6 mois.").font = Font(name="Calibri", size=11, bold=True, color="7F6000")
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=10)
ws.row_dimensions[2].height = 22

HEADERS = ["N°", "Réf. client", "Nom du client", "Agence principale", "CA HT 6 mois (tous produits)",
           "Vol Chick (t)", "CA Chick (FCFA)", "Qté Chick (sacs)", "% Chick dans CA total", "20/80"]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=4, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[4].height = 36

chick_sorted = sorted(client_chick.keys(), key=lambda k: -client_chick[k]["ca"])
START_ROW = 5
for i, key in enumerate(chick_sorted, start=1):
    r = START_ROW + i - 1
    ref, name = key
    banding = (i % 2 == 0)
    chick_data = client_chick[key]
    ca_total = client_ca_total.get(key, 0)
    pct = (chick_data["ca"] / ca_total) if ca_total > 0 else 0
    cells = [
        (1, i, BODY_ALIGN_CENTER),
        (2, ref, BODY_ALIGN_CENTER),
        (3, name, BODY_ALIGN_LEFT),
        (4, get_primary_agency(key), BODY_ALIGN_LEFT),
        (5, ca_total, BODY_ALIGN_RIGHT),
        (6, chick_data["vol_kg"] / 1000.0, BODY_ALIGN_RIGHT),
        (7, chick_data["ca"], BODY_ALIGN_RIGHT),
        (8, chick_data["qte"], BODY_ALIGN_RIGHT),
        (9, pct, BODY_ALIGN_CENTER),
    ]
    for col, val, align in cells:
        c = ws.cell(row=r, column=col, value=val)
        c.font = BODY_FONT; c.alignment = align; c.border = BORDER
        if banding: c.fill = BAND_FILL
    ws.cell(row=r, column=5).number_format = CA_NUM_FMT
    ws.cell(row=r, column=6).number_format = VOL_NUM_FMT
    ws.cell(row=r, column=7).number_format = CA_NUM_FMT
    ws.cell(row=r, column=8).number_format = '#,##0'
    ws.cell(row=r, column=9).number_format = '0.0%'
    # 20/80 marker
    c = ws.cell(row=r, column=10)
    c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
    if key in pareto_clients:
        c.value = PARETO_SYMBOL; c.fill = PARETO_FILL; c.font = PARETO_FONT
    else:
        if banding: c.fill = BAND_FILL

# Total row
total_row = START_ROW + len(chick_sorted)
total_ca_chick = sum(client_chick[k]["ca"] for k in chick_sorted)
total_vol_chick = sum(client_chick[k]["vol_kg"] for k in chick_sorted) / 1000.0
total_qte_chick = sum(client_chick[k]["qte"] for k in chick_sorted)
ws.cell(row=total_row, column=2, value="TOTAL").font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=total_row, column=2).fill = HEADER_FILL
ws.cell(row=total_row, column=2).alignment = BODY_ALIGN_CENTER; ws.cell(row=total_row, column=2).border = BORDER
ws.merge_cells(start_row=total_row, start_column=2, end_row=total_row, end_column=4)
for col in [3, 4]:
    ws.cell(row=total_row, column=col).fill = HEADER_FILL; ws.cell(row=total_row, column=col).border = BORDER
for col in [5]:
    ws.cell(row=total_row, column=col).fill = HEADER_FILL; ws.cell(row=total_row, column=col).border = BORDER
ws.cell(row=total_row, column=6, value=total_vol_chick).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=total_row, column=6).fill = PatternFill(start_color="BF6000", end_color="BF6000", fill_type="solid")
ws.cell(row=total_row, column=6).alignment = BODY_ALIGN_RIGHT; ws.cell(row=total_row, column=6).border = BORDER
ws.cell(row=total_row, column=6).number_format = VOL_NUM_FMT
ws.cell(row=total_row, column=7, value=total_ca_chick).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=total_row, column=7).fill = PatternFill(start_color="BF6000", end_color="BF6000", fill_type="solid")
ws.cell(row=total_row, column=7).alignment = BODY_ALIGN_RIGHT; ws.cell(row=total_row, column=7).border = BORDER
ws.cell(row=total_row, column=7).number_format = CA_NUM_FMT
ws.cell(row=total_row, column=8, value=total_qte_chick).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=total_row, column=8).fill = PatternFill(start_color="BF6000", end_color="BF6000", fill_type="solid")
ws.cell(row=total_row, column=8).alignment = BODY_ALIGN_RIGHT; ws.cell(row=total_row, column=8).border = BORDER
ws.cell(row=total_row, column=8).number_format = '#,##0'
for col in [9, 10]:
    ws.cell(row=total_row, column=col).fill = HEADER_FILL; ws.cell(row=total_row, column=col).border = BORDER

ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 16
ws.column_dimensions['C'].width = 42
ws.column_dimensions['D'].width = 25
ws.column_dimensions['E'].width = 22
ws.column_dimensions['F'].width = 14
ws.column_dimensions['G'].width = 22
ws.column_dimensions['H'].width = 16
ws.column_dimensions['I'].width = 16
ws.column_dimensions['J'].width = 9
ws.freeze_panes = "E5"
ws.auto_filter.ref = f"A4:J{START_ROW + len(chick_sorted) - 1}"
print(f"  Sheet 23 (Clients Chick Booster) built — {n_chick} clients")

# ===== SHEET 24: LISTE CLIENTS PIGLET BOOSTER =====
ws = wb.create_sheet("24. Clients Piglet Booster")
ws.cell(row=1, column=1, value="BELGOCAM SA - Clients acheteurs Piglet Booster (Janvier-Juin 2026)").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
ws.row_dimensions[1].height = 26

n_piglet = len(client_piglet)
ws.cell(row=2, column=1, value=f"📊 {n_piglet} client(s) ont acheté du Piglet Booster (CB200 25Kg ou CB201 5Kg) sur 6 mois.").font = Font(name="Calibri", size=11, bold=True, color="833C00")
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=10)
ws.row_dimensions[2].height = 22

HEADERS = ["N°", "Réf. client", "Nom du client", "Agence principale", "CA HT 6 mois (tous produits)",
           "Vol Piglet (t)", "CA Piglet (FCFA)", "Qté Piglet (sacs)", "% Piglet dans CA total", "20/80"]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=4, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[4].height = 36

piglet_sorted = sorted(client_piglet.keys(), key=lambda k: -client_piglet[k]["ca"])
START_ROW = 5
for i, key in enumerate(piglet_sorted, start=1):
    r = START_ROW + i - 1
    ref, name = key
    banding = (i % 2 == 0)
    piglet_data = client_piglet[key]
    ca_total = client_ca_total.get(key, 0)
    pct = (piglet_data["ca"] / ca_total) if ca_total > 0 else 0
    cells = [
        (1, i, BODY_ALIGN_CENTER),
        (2, ref, BODY_ALIGN_CENTER),
        (3, name, BODY_ALIGN_LEFT),
        (4, get_primary_agency(key), BODY_ALIGN_LEFT),
        (5, ca_total, BODY_ALIGN_RIGHT),
        (6, piglet_data["vol_kg"] / 1000.0, BODY_ALIGN_RIGHT),
        (7, piglet_data["ca"], BODY_ALIGN_RIGHT),
        (8, piglet_data["qte"], BODY_ALIGN_RIGHT),
        (9, pct, BODY_ALIGN_CENTER),
    ]
    for col, val, align in cells:
        c = ws.cell(row=r, column=col, value=val)
        c.font = BODY_FONT; c.alignment = align; c.border = BORDER
        if banding: c.fill = BAND_FILL
    ws.cell(row=r, column=5).number_format = CA_NUM_FMT
    ws.cell(row=r, column=6).number_format = VOL_NUM_FMT
    ws.cell(row=r, column=7).number_format = CA_NUM_FMT
    ws.cell(row=r, column=8).number_format = '#,##0'
    ws.cell(row=r, column=9).number_format = '0.0%'
    c = ws.cell(row=r, column=10)
    c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
    if key in pareto_clients:
        c.value = PARETO_SYMBOL; c.fill = PARETO_FILL; c.font = PARETO_FONT
    else:
        if banding: c.fill = BAND_FILL

total_row = START_ROW + len(piglet_sorted)
total_ca_piglet = sum(client_piglet[k]["ca"] for k in piglet_sorted)
total_vol_piglet = sum(client_piglet[k]["vol_kg"] for k in piglet_sorted) / 1000.0
total_qte_piglet = sum(client_piglet[k]["qte"] for k in piglet_sorted)
ws.cell(row=total_row, column=2, value="TOTAL").font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=total_row, column=2).fill = HEADER_FILL
ws.cell(row=total_row, column=2).alignment = BODY_ALIGN_CENTER; ws.cell(row=total_row, column=2).border = BORDER
ws.merge_cells(start_row=total_row, start_column=2, end_row=total_row, end_column=4)
for col in [3, 4, 5]:
    ws.cell(row=total_row, column=col).fill = HEADER_FILL; ws.cell(row=total_row, column=col).border = BORDER
ws.cell(row=total_row, column=6, value=total_vol_piglet).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=total_row, column=6).fill = PatternFill(start_color="833C00", end_color="833C00", fill_type="solid")
ws.cell(row=total_row, column=6).alignment = BODY_ALIGN_RIGHT; ws.cell(row=total_row, column=6).border = BORDER
ws.cell(row=total_row, column=6).number_format = VOL_NUM_FMT
ws.cell(row=total_row, column=7, value=total_ca_piglet).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=total_row, column=7).fill = PatternFill(start_color="833C00", end_color="833C00", fill_type="solid")
ws.cell(row=total_row, column=7).alignment = BODY_ALIGN_RIGHT; ws.cell(row=total_row, column=7).border = BORDER
ws.cell(row=total_row, column=7).number_format = CA_NUM_FMT
ws.cell(row=total_row, column=8, value=total_qte_piglet).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
ws.cell(row=total_row, column=8).fill = PatternFill(start_color="833C00", end_color="833C00", fill_type="solid")
ws.cell(row=total_row, column=8).alignment = BODY_ALIGN_RIGHT; ws.cell(row=total_row, column=8).border = BORDER
ws.cell(row=total_row, column=8).number_format = '#,##0'
for col in [9, 10]:
    ws.cell(row=total_row, column=col).fill = HEADER_FILL; ws.cell(row=total_row, column=col).border = BORDER

ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 16
ws.column_dimensions['C'].width = 42
ws.column_dimensions['D'].width = 25
ws.column_dimensions['E'].width = 22
ws.column_dimensions['F'].width = 14
ws.column_dimensions['G'].width = 22
ws.column_dimensions['H'].width = 16
ws.column_dimensions['I'].width = 16
ws.column_dimensions['J'].width = 9
ws.freeze_panes = "E5"
ws.auto_filter.ref = f"A4:J{START_ROW + len(piglet_sorted) - 1}"
print(f"  Sheet 24 (Clients Piglet Booster) built — {n_piglet} clients")

# ===== SHEET 25: SYNERGIE CHICK BOOSTER × CONCENTRÉS CHAIR/PONTE =====
ws = wb.create_sheet("25. Synergie Chick x Conc Chair-Ponte")
ws.cell(row=1, column=1, value="BELGOCAM SA - Synergie Chick Booster × Concentrés Chair/Ponte").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
ws.row_dimensions[1].height = 26

# Compute synergy stats
chick_clients_set = set(client_chick.keys())
chair_clients_set = set(client_chair.keys())
ponte_clients_set = set(client_ponte.keys())
chair_or_ponte_clients = chair_clients_set | ponte_clients_set

# Venn diagram: Chick buyers who also buy Chair or Ponte
chick_and_chair = chick_clients_set & chair_clients_set
chick_and_ponte = chick_clients_set & ponte_clients_set
chick_and_chair_or_ponte = chick_clients_set & chair_or_ponte_clients
chick_only = chick_clients_set - chair_or_ponte_clients

n_chick_and_chair_or_ponte = len(chick_and_chair_or_ponte)
pct_synergy_chick = (n_chick_and_chair_or_ponte / n_chick * 100) if n_chick > 0 else 0

ws.cell(row=2, column=1, value=f"📊 Sur {n_chick} clients Chick Booster, {n_chick_and_chair_or_ponte} ({pct_synergy_chick:.1f}%) achètent AUSSI des concentrés Chair ou Ponte. {len(chick_only)} ({100-pct_synergy_chick:.1f}%) n'achètent QUE du Chick (opportunité cross-sell).").font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=10)
ws.row_dimensions[2].height = 36

# Section 1: Venn / cross-tab summary
ws.cell(row=4, column=1, value="1. Synthèse de la synergie").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=10)
ws.row_dimensions[4].height = 22

synth_headers = ["Catégorie", "Nb clients", "% des clients Chick", "Commentaire", "", "", "", "", "", ""]
for col_idx, h in enumerate(synth_headers, start=1):
    c = ws.cell(row=5, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[5].height = 32

synth_rows = [
    (f"Total clients Chick Booster", n_chick, 1.0, "Base 100%"),
    (f"Clients Chick + Concentrés Chair (C104, C1042, C1043, C1044, C103)", len(chick_and_chair), len(chick_and_chair) / n_chick if n_chick > 0 else 0, "Cross-sell sur Chair"),
    (f"Clients Chick + Concentrés Ponte (C102, C1022, C101)", len(chick_and_ponte), len(chick_and_ponte) / n_chick if n_chick > 0 else 0, "Cross-sell sur Ponte"),
    (f"Clients Chick + Chair OU Ponte (cross-sell réussi)", n_chick_and_chair_or_ponte, n_chick_and_chair_or_ponte / n_chick if n_chick > 0 else 0, "✅ Cross-sell sur au moins 1 concentré volaille"),
    (f"Clients Chick ONLY (pas de concentrés Chair/Ponte)", len(chick_only), len(chick_only) / n_chick if n_chick > 0 else 0, "⚠️ Opportunité cross-sell énorme"),
]
for i, (cat, n, pct, comment) in enumerate(synth_rows, start=1):
    r = 5 + i
    banding = (i % 2 == 0)
    c = ws.cell(row=r, column=1, value=cat)
    c.font = BOLD_FONT; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=2, value=n)
    c.font = BOLD_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=3, value=pct)
    c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c.number_format = '0.0%'
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=4, value=comment)
    c.font = BODY_FONT; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=10)
    if banding: c.fill = BAND_FILL
    # Highlight specific rows
    if "cross-sell réussi" in cat.lower() or "chair ou ponte" in cat.lower():
        ws.cell(row=r, column=2).fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        ws.cell(row=r, column=2).font = Font(name="Calibri", size=10, bold=True, color="006100")
    elif "chick only" in cat.lower():
        ws.cell(row=r, column=2).fill = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")
        ws.cell(row=r, column=2).font = Font(name="Calibri", size=10, bold=True, color="C00000")

# Section 2: Liste détaillée
r_detail_start = 5 + len(synth_rows) + 2
ws.cell(row=r_detail_start, column=1, value="2. Liste détaillée des clients Chick Booster avec statut cross-sell").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
ws.merge_cells(start_row=r_detail_start, start_column=1, end_row=r_detail_start, end_column=10)
ws.row_dimensions[r_detail_start].height = 22

HEADERS = ["N°", "Réf. client", "Nom du client", "Agence", "CA Chick (FCFA)",
           "CA Concentrés Chair (FCFA)", "CA Concentrés Ponte (FCFA)", "Statut cross-sell", "20/80", ""]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=r_detail_start + 1, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[r_detail_start + 1].height = 36

START_ROW = r_detail_start + 2
chick_sorted2 = sorted(client_chick.keys(), key=lambda k: -client_chick[k]["ca"])
for i, key in enumerate(chick_sorted2, start=1):
    r = START_ROW + i - 1
    ref, name = key
    banding = (i % 2 == 0)
    chick_ca = client_chick[key]["ca"]
    chair_ca = client_chair.get(key, {"ca": 0})["ca"]
    ponte_ca = client_ponte.get(key, {"ca": 0})["ca"]
    if key in chick_and_chair_or_ponte:
        status = "✅ Cross-sell"
        status_color = "C6EFCE"
        status_font = Font(name="Calibri", size=10, bold=True, color="006100")
    else:
        status = "⚠️ Chick only"
        status_color = "FCE4E4"
        status_font = Font(name="Calibri", size=10, bold=True, color="C00000")

    cells = [
        (1, i, BODY_ALIGN_CENTER),
        (2, ref, BODY_ALIGN_CENTER),
        (3, name, BODY_ALIGN_LEFT),
        (4, get_primary_agency(key), BODY_ALIGN_LEFT),
        (5, chick_ca, BODY_ALIGN_RIGHT),
        (6, chair_ca, BODY_ALIGN_RIGHT),
        (7, ponte_ca, BODY_ALIGN_RIGHT),
    ]
    for col, val, align in cells:
        c = ws.cell(row=r, column=col, value=val)
        c.font = BODY_FONT; c.alignment = align; c.border = BORDER
        if banding: c.fill = BAND_FILL
    for col in [5, 6, 7]:
        ws.cell(row=r, column=col).number_format = CA_NUM_FMT
    # Status with color
    c = ws.cell(row=r, column=8, value=status)
    c.font = status_font; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c.fill = PatternFill(start_color=status_color, end_color=status_color, fill_type="solid")
    # 20/80 marker
    c = ws.cell(row=r, column=9)
    c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
    if key in pareto_clients:
        c.value = PARETO_SYMBOL; c.fill = PARETO_FILL; c.font = PARETO_FONT
    else:
        if banding: c.fill = BAND_FILL

# Column widths
ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 16
ws.column_dimensions['C'].width = 38
ws.column_dimensions['D'].width = 22
ws.column_dimensions['E'].width = 18
ws.column_dimensions['F'].width = 22
ws.column_dimensions['G'].width = 22
ws.column_dimensions['H'].width = 16
ws.column_dimensions['I'].width = 9
ws.column_dimensions['J'].width = 8
ws.freeze_panes = f"E{START_ROW}"
ws.auto_filter.ref = f"A{r_detail_start + 1}:I{START_ROW + len(chick_sorted2) - 1}"

# Native pie chart for synergy
# Write chart data in column L (hidden)
DATA_COL = 12
ws.cell(row=5, column=DATA_COL, value="Catégorie").font = BOLD_FONT
ws.cell(row=5, column=DATA_COL+1, value="Nb clients").font = BOLD_FONT
chart_data = [
    ("Cross-sell (Chair ou Ponte)", n_chick_and_chair_or_ponte),
    ("Chick only (opportunité)", len(chick_only)),
]
for i, (cat, n) in enumerate(chart_data, start=1):
    ws.cell(row=5+i, column=DATA_COL, value=cat)
    ws.cell(row=5+i, column=DATA_COL+1, value=n)
ws.column_dimensions[get_column_letter(DATA_COL)].hidden = True
ws.column_dimensions[get_column_letter(DATA_COL+1)].hidden = True

from openpyxl.chart import PieChart
chart_pie = PieChart()
chart_pie.title = f"Synergie Chick Booster × Concentrés Chair/Ponte (sur {n_chick} clients Chick)"
chart_pie.height = 10
chart_pie.width = 14
data = Reference(ws, min_col=DATA_COL+1, min_row=5, max_row=7)
cats = Reference(ws, min_col=DATA_COL, min_row=6, max_row=7)
chart_pie.add_data(data, titles_from_data=True)
chart_pie.set_categories(cats)
chart_pie.dataLabels = DataLabelList(showPercent=True, showCatName=True)
pie_colors_2 = ["70AD47", "C00000"]
if chart_pie.series:
    s = chart_pie.series[0]
    s.data_points = [DataPoint(idx=i, spPr=GraphicalProperties(solidFill=pie_colors_2[i])) for i in range(2)]
# Anchor chart to the right of the synthesis section
ws.add_chart(chart_pie, f"J5")

print(f"  Sheet 25 (Synergie Chick x Conc Chair/Ponte) built — {n_chick} clients Chick, {n_chick_and_chair_or_ponte} cross-sell ({pct_synergy_chick:.1f}%)")

# ===== SHEET 26: SYNERGIE PIGLET BOOSTER × CONCENTRÉS PORC =====
ws = wb.create_sheet("26. Synergie Piglet x Conc Porc")
ws.cell(row=1, column=1, value="BELGOCAM SA - Synergie Piglet Booster × Concentrés Porc").font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
ws.row_dimensions[1].height = 26

# Compute synergy stats
piglet_clients_set = set(client_piglet.keys())
porc_clients_set = set(client_porc.keys())
piglet_and_porc = piglet_clients_set & porc_clients_set
piglet_only = piglet_clients_set - porc_clients_set

n_piglet_and_porc = len(piglet_and_porc)
pct_synergy_piglet = (n_piglet_and_porc / n_piglet * 100) if n_piglet > 0 else 0

ws.cell(row=2, column=1, value=f"📊 Sur {n_piglet} clients Piglet Booster, {n_piglet_and_porc} ({pct_synergy_piglet:.1f}%) achètent AUSSI des concentrés Porc (C105, C1053, C1054, C1055). {len(piglet_only)} ({100-pct_synergy_piglet:.1f}%) n'achètent QUE du Piglet (opportunité cross-sell).").font = Font(name="Calibri", size=11, bold=True, color="833C00")
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=10)
ws.row_dimensions[2].height = 36

# Section 1: Synthèse
ws.cell(row=4, column=1, value="1. Synthèse de la synergie").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=10)
ws.row_dimensions[4].height = 22

synth_headers = ["Catégorie", "Nb clients", "% des clients Piglet", "Commentaire", "", "", "", "", "", ""]
for col_idx, h in enumerate(synth_headers, start=1):
    c = ws.cell(row=5, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[5].height = 32

synth_rows = [
    ("Total clients Piglet Booster", n_piglet, 1.0, "Base 100%"),
    (f"Clients Piglet + Concentrés Porc (C105, C1053, C1054, C1055)", n_piglet_and_porc, n_piglet_and_porc / n_piglet if n_piglet > 0 else 0, "✅ Cross-sell sur Porc"),
    (f"Clients Piglet ONLY (pas de concentrés Porc)", len(piglet_only), len(piglet_only) / n_piglet if n_piglet > 0 else 0, "⚠️ Opportunité cross-sell énorme"),
]
for i, (cat, n, pct, comment) in enumerate(synth_rows, start=1):
    r = 5 + i
    banding = (i % 2 == 0)
    c = ws.cell(row=r, column=1, value=cat)
    c.font = BOLD_FONT; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=2, value=n)
    c.font = BOLD_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=3, value=pct)
    c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c.number_format = '0.0%'
    if banding: c.fill = BAND_FILL
    c = ws.cell(row=r, column=4, value=comment)
    c.font = BODY_FONT; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=10)
    if banding: c.fill = BAND_FILL
    if "cross-sell" in cat.lower() and "porc" in cat.lower():
        ws.cell(row=r, column=2).fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        ws.cell(row=r, column=2).font = Font(name="Calibri", size=10, bold=True, color="006100")
    elif "piglet only" in cat.lower():
        ws.cell(row=r, column=2).fill = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")
        ws.cell(row=r, column=2).font = Font(name="Calibri", size=10, bold=True, color="C00000")

# Section 2: Liste détaillée
r_detail_start = 5 + len(synth_rows) + 2
ws.cell(row=r_detail_start, column=1, value="2. Liste détaillée des clients Piglet Booster avec statut cross-sell").font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
ws.merge_cells(start_row=r_detail_start, start_column=1, end_row=r_detail_start, end_column=10)
ws.row_dimensions[r_detail_start].height = 22

HEADERS = ["N°", "Réf. client", "Nom du client", "Agence", "CA Piglet (FCFA)",
           "CA Concentrés Porc (FCFA)", "Total Booster+Porc (FCFA)", "Statut cross-sell", "20/80", ""]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=r_detail_start + 1, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[r_detail_start + 1].height = 36

START_ROW = r_detail_start + 2
piglet_sorted2 = sorted(client_piglet.keys(), key=lambda k: -client_piglet[k]["ca"])
for i, key in enumerate(piglet_sorted2, start=1):
    r = START_ROW + i - 1
    ref, name = key
    banding = (i % 2 == 0)
    piglet_ca = client_piglet[key]["ca"]
    porc_ca = client_porc.get(key, {"ca": 0})["ca"]
    if key in piglet_and_porc:
        status = "✅ Cross-sell"
        status_color = "C6EFCE"
        status_font = Font(name="Calibri", size=10, bold=True, color="006100")
    else:
        status = "⚠️ Piglet only"
        status_color = "FCE4E4"
        status_font = Font(name="Calibri", size=10, bold=True, color="C00000")

    cells = [
        (1, i, BODY_ALIGN_CENTER),
        (2, ref, BODY_ALIGN_CENTER),
        (3, name, BODY_ALIGN_LEFT),
        (4, get_primary_agency(key), BODY_ALIGN_LEFT),
        (5, piglet_ca, BODY_ALIGN_RIGHT),
        (6, porc_ca, BODY_ALIGN_RIGHT),
        (7, piglet_ca + porc_ca, BODY_ALIGN_RIGHT),
    ]
    for col, val, align in cells:
        c = ws.cell(row=r, column=col, value=val)
        c.font = BODY_FONT; c.alignment = align; c.border = BORDER
        if banding: c.fill = BAND_FILL
    for col in [5, 6, 7]:
        ws.cell(row=r, column=col).number_format = CA_NUM_FMT
    c = ws.cell(row=r, column=8, value=status)
    c.font = status_font; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    c.fill = PatternFill(start_color=status_color, end_color=status_color, fill_type="solid")
    c = ws.cell(row=r, column=9)
    c.border = BORDER; c.alignment = BODY_ALIGN_CENTER
    if key in pareto_clients:
        c.value = PARETO_SYMBOL; c.fill = PARETO_FILL; c.font = PARETO_FONT
    else:
        if banding: c.fill = BAND_FILL

ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 16
ws.column_dimensions['C'].width = 38
ws.column_dimensions['D'].width = 22
ws.column_dimensions['E'].width = 18
ws.column_dimensions['F'].width = 22
ws.column_dimensions['G'].width = 22
ws.column_dimensions['H'].width = 16
ws.column_dimensions['I'].width = 9
ws.column_dimensions['J'].width = 8
ws.freeze_panes = f"E{START_ROW}"
ws.auto_filter.ref = f"A{r_detail_start + 1}:I{START_ROW + len(piglet_sorted2) - 1}"

# Native pie chart
DATA_COL = 12
ws.cell(row=5, column=DATA_COL, value="Catégorie").font = BOLD_FONT
ws.cell(row=5, column=DATA_COL+1, value="Nb clients").font = BOLD_FONT
chart_data = [
    ("Cross-sell Porc", n_piglet_and_porc),
    ("Piglet only (opportunité)", len(piglet_only)),
]
for i, (cat, n) in enumerate(chart_data, start=1):
    ws.cell(row=5+i, column=DATA_COL, value=cat)
    ws.cell(row=5+i, column=DATA_COL+1, value=n)
ws.column_dimensions[get_column_letter(DATA_COL)].hidden = True
ws.column_dimensions[get_column_letter(DATA_COL+1)].hidden = True

chart_pie2 = PieChart()
chart_pie2.title = f"Synergie Piglet Booster × Concentrés Porc (sur {n_piglet} clients Piglet)"
chart_pie2.height = 10
chart_pie2.width = 14
data = Reference(ws, min_col=DATA_COL+1, min_row=5, max_row=7)
cats = Reference(ws, min_col=DATA_COL, min_row=6, max_row=7)
chart_pie2.add_data(data, titles_from_data=True)
chart_pie2.set_categories(cats)
chart_pie2.dataLabels = DataLabelList(showPercent=True, showCatName=True)
if chart_pie2.series:
    s = chart_pie2.series[0]
    s.data_points = [DataPoint(idx=i, spPr=GraphicalProperties(solidFill=pie_colors_2[i])) for i in range(2)]
ws.add_chart(chart_pie2, f"J5")

print(f"  Sheet 26 (Synergie Piglet x Conc Porc) built — {n_piglet} clients Piglet, {n_piglet_and_porc} cross-sell ({pct_synergy_piglet:.1f}%)")

# ===== SHEET 27-30: ANNEXE 20/80 PAR TYPE DE CONCENTRÉS =====
import json as _json
with open('/home/z/my-project/scripts/annexe_20_80_par_type.json', 'r', encoding='utf-8') as _f:
    _annexe = _json.load(_f)

# Sheet 27: Synthèse annexe
ws = wb.create_sheet("27. Annexe 20-80 Synthèse")
headers = ["Type", "Clients S1", "Top 20/80", "Vol S1 (t)", "CA S1 (M)", "Obj S2 (t)", "Obj S2 CA (M)", "Croissance"]
ws.append(headers)
for col in range(1, len(headers)+1):
    c = ws.cell(row=1, column=col)
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border = thin_border

row_idx = 2
for t in ["Chair", "Ponte", "Porc"]:
    if t not in _annexe: continue
    d = _annexe[t]
    croissance = (d["obj_s2_t"]/d["total_vol_t"]-1)*100 if d["total_vol_t"] > 0 else 0
    ws.cell(row=row_idx, column=1, value=t)
    ws.cell(row=row_idx, column=2, value=d["total_clients"])
    ws.cell(row=row_idx, column=3, value=d["top_20_80_count"])
    ws.cell(row=row_idx, column=4, value=round(d["total_vol_t"], 1))
    ws.cell(row=row_idx, column=5, value=round(d["total_ca_m"], 1))
    ws.cell(row=row_idx, column=6, value=round(d["obj_s2_t"], 0))
    # CA obj S2 = obj_t * prix_moyen_S1
    prix_moyen = d["total_ca_m"] / d["total_vol_t"] * 1000 if d["total_vol_t"] > 0 else 0
    ws.cell(row=row_idx, column=7, value=round(d["obj_s2_t"] * prix_moyen / 1000, 1))
    ws.cell(row=row_idx, column=8, value=f"+{croissance:.1f}%")
    row_idx += 1

# Total
total_clients = sum(_annexe[t]["total_clients"] for t in _annexe)
total_2080 = sum(_annexe[t]["top_20_80_count"] for t in _annexe)
total_vol = sum(_annexe[t]["total_vol_t"] for t in _annexe)
total_ca = sum(_annexe[t]["total_ca_m"] for t in _annexe)
total_obj = sum(_annexe[t]["obj_s2_t"] for t in _annexe)
prix_moyen_global = total_ca / total_vol * 1000 if total_vol > 0 else 0
total_obj_ca = total_obj * prix_moyen_global / 1000
ws.cell(row=row_idx, column=1, value="TOTAL")
ws.cell(row=row_idx, column=2, value=total_clients)
ws.cell(row=row_idx, column=3, value=total_2080)
ws.cell(row=row_idx, column=4, value=round(total_vol, 1))
ws.cell(row=row_idx, column=5, value=round(total_ca, 1))
ws.cell(row=row_idx, column=6, value=round(total_obj, 0))
ws.cell(row=row_idx, column=7, value=round(total_obj_ca, 1))
ws.cell(row=row_idx, column=8, value=f"+{(total_obj/total_vol-1)*100:.1f}%")
for col in range(1, 9):
    c = ws.cell(row=row_idx, column=col)
    c.font = Font(bold=True)
    c.fill = PARETO_FILL

# Column widths
for col_letter, w in zip("ABCDEFGH", [10, 12, 12, 12, 12, 12, 14, 12]):
    ws.column_dimensions[col_letter].width = w

print(f"  Sheet 27 (Annexe 20/80 Synthèse) built")

# Sheets 28-30: Top 20/80 par type
sheet_names_map = {"Chair": "28. Annexe 20-80 Chair", "Ponte": "29. Annexe 20-80 Ponte", "Porc": "30. Annexe 20-80 Porc"}
for type_conc, sheet_name in sheet_names_map.items():
    ws = wb.create_sheet(sheet_name)
    headers = ["Rang", "Code client", "Nom client", "Vol S1 (t)", "CA S1 (M FCFA)", "% vol type", "Obj S2 (t)", "CA obj S2 (M)", "Prix moyen (k/t)"]
    ws.append(headers)
    for col in range(1, len(headers)+1):
        c = ws.cell(row=1, column=col)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = thin_border

    if type_conc in _annexe:
        for c_data in _annexe[type_conc]["clients"]:
            ws.append([
                c_data["rang"],
                c_data["code"],
                c_data["nom"],
                c_data["vol_s1_t"],
                c_data["ca_s1_m"],
                c_data["pct_vol"],
                c_data["obj_s2_t"],
                c_data["obj_s2_ca_m"],
                c_data["prix_moyen_k_t"]
            ])
    
    # Column widths
    for col_letter, w in zip("ABCDEFGHI", [6, 18, 40, 12, 14, 10, 12, 14, 14]):
        ws.column_dimensions[col_letter].width = w
    
    # Freeze first row
    ws.freeze_panes = "A2"
    
    print(f"  Sheet {sheet_name.split('.')[0]} ({type_conc}) built — {len(_annexe.get(type_conc, {}).get('clients', []))} clients")

# ===== SAVE =====
wb.save(OUT)
print(f"\nSaved: {OUT}")
print(f"File size: {os.path.getsize(OUT):,} bytes")
print(f"Total sheets: {len(wb.sheetnames)}")
print(f"Sheets: {wb.sheetnames}")
