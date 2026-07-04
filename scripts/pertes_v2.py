"""
Compute volume/CA losses for zero-achat clients using the user's methodology:

  Perte Q1 (par produit) = Moyenne mensuelle (sur mois actifs) × Fréquence × 3 mois

Where:
- "Moyenne mensuelle (sur mois actifs)" for product P
    = total volume of P in Q2 / number of months in Q2 where P was bought
- "Fréquence" (client-level, on the targeted category)
    = number of months in Q2 where client bought ANY targeted product / 3
- "3 mois" = number of months in Q1 (Jan-Mar) where the client was a "zero achat"

So: Perte Q1 (total) = Σ_P (moyenne_mensuelle_P) × Fréquence × 3
                   = Σ_P (total_Q2_P / n_months_P) × (n_active_months / 3) × 3

This makes frequency matter:
- A client active every month in Q2 (freq=1) has full loss
- A client active 1 month out of 3 in Q2 (freq=1/3) has 1/3 of the loss

For CA loss, we use the client's own average price (CA/volume) for each product
when available, else 0.

Output: /home/z/my-project/download/pertes_zero_achat.xlsx with 2 sheets:
- "Zero global (16)"  : 352 clients, loss on ciblé (soja + concentrés)
- "Zero concentrés"   : 544 clients, loss on concentrés only
"""
import re
import os
import time
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict

# ---- Paths ----
SRC = "/home/z/my-project/download/ventes_livrees.xlsx"
OUT = "/home/z/my-project/download/pertes_zero_achat.xlsx"

# ---- Product categories ----
SOJA_REFS = {"T102", "T1021", "T1023", "T1024"}
CONCENTRE_REFS = {
    "C102", "C1022", "C104", "C1042", "C1043", "C1044",
    "C105", "C1053", "C1054", "C1055",
    "C101", "C103",
}
TARGET_REFS = SOJA_REFS | CONCENTRE_REFS  # 16

# ---- Unit weight (kg) parser --------------------------------------------
WEIGHT_RE = re.compile(
    r"(\d+(?:\.\d+)?)\s*(KG|Kg|kg|KG|GRAMMES?|G|L)\b",
    re.IGNORECASE
)

def parse_weight_kg(desc):
    if not desc:
        return 0.0
    matches = WEIGHT_RE.findall(str(desc))
    if not matches:
        return 0.0
    val_str, unit = matches[-1]
    val = float(val_str)
    unit_upper = unit.upper()
    if unit_upper == "KG":
        return val
    if unit_upper in ("G", "GRAMME", "GRAMMES"):
        return val / 1000.0
    if unit_upper == "L":
        return val
    return 0.0

# ---- Months --------------------------------------------------------------
Q2_MONTHS = [
    ("Feuil3", "Avril",  3),
    ("Feuil4", "Mai",    4),
    ("Feuil5", "Juin",   5),
]
Q1_MONTHS_SHEETS = {"Sheet 1", "Feuil1", "Feuil2"}

TIERS_RE = re.compile(r"^\s*([A-Z]{2,}[\w-]*?)\s*-\s+(.+?)\s*$")

def split_tiers(tiers):
    if tiers is None:
        return ("", "")
    s = str(tiers).strip()
    m = TIERS_RE.match(s)
    if m:
        return (m.group(1).strip(), m.group(2).strip())
    return ("", s)

# ---- 1. Read source & aggregate ------------------------------------------
print(f"Reading source: {SRC}")
t0 = time.time()
wb_src = load_workbook(SRC, read_only=True, data_only=True)
print(f"  Loaded in {time.time()-t0:.1f}s")

# Product weight map (ref -> kg per unit)
product_weight = {}

# Per-client aggregates:
# - client_q2_by_product[key][ref_prod] = {"vol_kg": float, "ca": float, "months": set()}
#   Accumulates Q2 purchases per product per client.
# - client_q1_targeted_flag[key] = True if bought any targeted product in Q1 (zero-achat filter)
# - client_q1_concentre_flag[key] = True if bought any concentré in Q1
# - client_ca_total[key] = total CA HT over 6 months
# - client_q2_active_months[key] = set of months where client bought ANY targeted product in Q2
# - client_q2_active_months_concentre[key] = set of months where client bought ANY concentré in Q2

client_q2_by_product = defaultdict(lambda: defaultdict(lambda: {"vol_kg": 0.0, "ca": 0.0, "months": set(), "occasions": set()}))
client_q1_targeted_flag = defaultdict(bool)
client_q1_concentre_flag = defaultdict(bool)
client_ca_total = defaultdict(float)
client_q2_active_months_targeted = defaultdict(set)
client_q2_active_months_concentre = defaultdict(set)
client_q2_active_months_any = defaultdict(set)  # any product (for general activity)

for sheet_name in wb_src.sheetnames:
    ws = wb_src[sheet_name]
    is_q1 = sheet_name in Q1_MONTHS_SHEETS
    is_q2 = not is_q1
    n = 0
    for row in ws.iter_rows(min_row=3, values_only=True):
        tiers = row[5] if len(row) > 5 else None
        if tiers is None or str(tiers).strip() == "":
            continue
        n += 1
        ref_prod = row[0] if len(row) > 0 else None
        desc     = row[1] if len(row) > 1 else None
        qte      = row[2] if len(row) > 2 else 0
        date_cmd = row[6] if len(row) > 6 else None   # G: Date de commande (e.g. "02/01/2026")
        ca_ht    = row[8] if len(row) > 8 else 0

        ref_c, name_c = split_tiers(tiers)
        key = (ref_c, name_c)

        ref_prod_str = str(ref_prod).strip() if ref_prod is not None else ""
        if ref_prod_str and ref_prod_str not in product_weight:
            product_weight[ref_prod_str] = parse_weight_kg(desc)
        weight_kg = product_weight.get(ref_prod_str, 0.0)

        try:
            qte_f = float(qte) if qte is not None else 0.0
        except (ValueError, TypeError):
            qte_f = 0.0
        try:
            ca_f = float(ca_ht) if ca_ht is not None else 0.0
        except (ValueError, TypeError):
            ca_f = 0.0

        client_ca_total[key] += ca_f

        is_targeted = ref_prod_str in TARGET_REFS
        is_concentre = ref_prod_str in CONCENTRE_REFS

        # Extract month from date_cmd (format "DD/MM/YYYY")
        month = None
        if date_cmd:
            parts = str(date_cmd).split("/")
            if len(parts) == 3:
                try:
                    month = int(parts[1])
                except ValueError:
                    pass

        if is_q1:
            if is_targeted:
                client_q1_targeted_flag[key] = True
            if is_concentre:
                client_q1_concentre_flag[key] = True
        else:  # Q2
            # Accumulate per-product Q2 stats (only for targeted products to save memory)
            if is_targeted:
                rec = client_q2_by_product[key][ref_prod_str]
                rec["vol_kg"] += qte_f * weight_kg
                rec["ca"] += ca_f
                if month:
                    rec["months"].add(month)
                # Use date_cmd as occasion identifier
                if date_cmd:
                    rec["occasions"].add(str(date_cmd))

                # Track active months on targeted category
                if month:
                    client_q2_active_months_targeted[key].add(month)
                if is_concentre and month:
                    client_q2_active_months_concentre[key].add(month)

            # Track active months on ANY product (general activity)
            if month:
                client_q2_active_months_any[key].add(month)

    print(f"  {sheet_name}: {n} client occurrences")

wb_src.close()

clients_all = sorted(client_ca_total.keys(), key=lambda k: (k[1].upper(), k[0]))
print(f"\nTotal unique clients: {len(clients_all)}")

# ---- 2. Pareto 20/80 ----------------------------------------------------
total_ca = sum(client_ca_total.values())
sorted_by_ca = sorted(clients_all, key=lambda k: -client_ca_total[k])
pareto_threshold = 0.80 * total_ca
pareto_clients = set()
cumulative = 0.0
for key in sorted_by_ca:
    cumulative += client_ca_total[key]
    pareto_clients.add(key)
    if cumulative >= pareto_threshold:
        break

# ---- 3. Build zero-achat lists ------------------------------------------
zero_global = [k for k in clients_all if not client_q1_targeted_flag[k]]
zero_concentre = [k for k in clients_all if not client_q1_concentre_flag[k]]

print(f"Zero achat global (16 produits)     : {len(zero_global)} clients")
print(f"Zero achat concentrés (12 produits) : {len(zero_concentre)} clients")

# ---- 4. Compute losses per client per category ---------------------------
def compute_loss(client_key, category_refs):
    """
    Compute Q1 loss for a client on a given category (set of product refs).

    Returns dict with:
      - sum_monthly_avg_tons   : Σ_P (total_Q2_P / n_months_P)  [tons]
      - frequency              : n_active_months_targeted_Q2 / 3
      - n_active_months        : number of months in Q2 with any targeted purchase
      - loss_tons              : sum_monthly_avg_tons × frequency × 3
      - loss_fcfa              : Σ_P (loss_P_tons × client_price_per_ton_P)
      - n_products             : number of distinct targeted products bought in Q2
      - per_product_detail     : list of (ref_prod, total_tons, n_months, monthly_avg_tons,
                                          loss_tons, ca_q2, price_per_ton, loss_fcfa)
    """
    products_q2 = client_q2_by_product.get(client_key, {})
    active_months_set = client_q2_active_months_targeted.get(client_key, set())
    n_active_months = len(active_months_set)
    frequency = n_active_months / 3.0  # Q2 = 3 months

    sum_monthly_avg_tons = 0.0
    loss_tons = 0.0
    loss_fcfa = 0.0
    per_product = []

    for ref_prod, rec in products_q2.items():
        if ref_prod not in category_refs:
            continue
        total_kg = rec["vol_kg"]
        total_tons = total_kg / 1000.0
        ca_q2 = rec["ca"]
        n_months_P = len(rec["months"])
        if n_months_P == 0:
            continue  # can't compute average without month info
        monthly_avg_tons = total_tons / n_months_P
        sum_monthly_avg_tons += monthly_avg_tons

        # Loss for this product = monthly_avg × frequency × 3
        loss_P_tons = monthly_avg_tons * frequency * 3
        loss_tons += loss_P_tons

        # Price per ton (client's own price in Q2)
        price_per_ton = (ca_q2 / total_tons) if total_tons > 0 else 0.0
        loss_P_fcfa = loss_P_tons * price_per_ton
        loss_fcfa += loss_P_fcfa

        per_product.append({
            "ref_prod": ref_prod,
            "total_tons": total_tons,
            "n_months": n_months_P,
            "monthly_avg_tons": monthly_avg_tons,
            "loss_tons": loss_P_tons,
            "ca_q2": ca_q2,
            "price_per_ton": price_per_ton,
            "loss_fcfa": loss_P_fcfa,
        })

    return {
        "sum_monthly_avg_tons": sum_monthly_avg_tons,
        "frequency": frequency,
        "n_active_months": n_active_months,
        "loss_tons": loss_tons,
        "loss_fcfa": loss_fcfa,
        "n_products": len(per_product),
        "per_product_detail": per_product,
    }


# Compute losses for both lists
print("\nComputing losses for zero-achat global clients...")
losses_global = {}
for k in zero_global:
    losses_global[k] = compute_loss(k, TARGET_REFS)

print("Computing losses for zero-achat concentrés clients...")
losses_concentre = {}
for k in zero_concentre:
    losses_concentre[k] = compute_loss(k, CONCENTRE_REFS)

# Quick top 5 preview
print("\nTop 5 pertes (zero global) — by loss_tons:")
for k in sorted(zero_global, key=lambda k: -losses_global[k]["loss_tons"])[:5]:
    L = losses_global[k]
    print(f"  {k[1][:40]:<40}  perte={L['loss_tons']:>8.2f} t  "
          f"freq={L['frequency']:.2f}  moy_mens={L['sum_monthly_avg_tons']:>6.2f} t  "
          f"CA_perte={L['loss_fcfa']:>14,.0f} FCFA")

print("\nTop 5 pertes (zero concentrés) — by loss_tons:")
for k in sorted(zero_concentre, key=lambda k: -losses_concentre[k]["loss_tons"])[:5]:
    L = losses_concentre[k]
    print(f"  {k[1][:40]:<40}  perte={L['loss_tons']:>8.2f} t  "
          f"freq={L['frequency']:.2f}  moy_mens={L['sum_monthly_avg_tons']:>6.2f} t  "
          f"CA_perte={L['loss_fcfa']:>14,.0f} FCFA")

# ---- 5. Build output workbook --------------------------------------------
print(f"\n{'='*70}")
print(f"Building output: {OUT}")
print(f"{'='*70}")

# ---- Styles ----
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E78")
TITLE_ALIGN = Alignment(horizontal="left", vertical="center")
SUB_FONT = Font(name="Calibri", size=10, italic=True, color="595959")
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
GROUP_FILL = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
GROUP_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
GROUP_ALIGN = Alignment(horizontal="center", vertical="center")
LOSS_GROUP_FILL = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
BAND_FILL = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
BODY_FONT = Font(name="Calibri", size=10)
BODY_ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
BODY_ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
BODY_ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
PARETO_FILL = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
PARETO_FONT = Font(name="Calibri", size=14, bold=True, color="7F6000")
PARETO_SYMBOL = "★"
LOSS_HIGHLIGHT_FILL = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")
LOSS_FONT = Font(name="Calibri", size=10, bold=True, color="C00000")
CA_NUM_FMT = '#,##0" FCFA"'
VOL_NUM_FMT = '#,##0.00" t"'
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def build_sheet(ws, title_text, sub_text, client_list, losses_dict):
    """Build one analysis sheet."""
    # Sort clients by loss_tons descending (biggest losses first)
    sorted_clients = sorted(
        client_list,
        key=lambda k: -losses_dict[k]["loss_tons"]
    )
    n_clients = len(sorted_clients)

    # Column layout (1-based):
    # A: N°
    # B: Réf. client
    # C: Nom du client
    # D: Nb produits achetés (Q2)
    # E: Nb mois actifs (Q2)
    # F: Fréquence (mois/3)
    # G: Σ moyennes mensuelles (t)
    # H: PERTE Q1 - Volume (t)  = G × F × 3
    # I: PERTE Q1 - CA (FCFA)
    # J: CA HT total 6 mois
    # K: 20/80

    # Row 1: Title
    ws.cell(row=1, column=1, value=title_text).font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = TITLE_ALIGN
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=11)
    ws.row_dimensions[1].height = 26

    # Row 2: Subtitle (criteria + methodology)
    ws.cell(row=2, column=1, value=sub_text).font = SUB_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=11)
    ws.row_dimensions[2].height = 56

    # Row 3: Group headers
    ws.merge_cells(start_row=3, start_column=4, end_row=3, end_column=7)
    c = ws.cell(row=3, column=4, value="Comportement d'achat en Q2 (Avr-Juin)")
    c.fill = GROUP_FILL; c.font = GROUP_FONT; c.alignment = GROUP_ALIGN; c.border = BORDER
    for col in [5, 6, 7]:
        cc = ws.cell(row=3, column=col); cc.fill = GROUP_FILL; cc.border = BORDER

    ws.merge_cells(start_row=3, start_column=8, end_row=3, end_column=9)
    c = ws.cell(row=3, column=8, value="PERTE Q1 estimée (Jan-Mar)")
    c.fill = LOSS_GROUP_FILL; c.font = GROUP_FONT; c.alignment = GROUP_ALIGN; c.border = BORDER
    for col in [9]:
        cc = ws.cell(row=3, column=col); cc.fill = LOSS_GROUP_FILL; cc.border = BORDER

    ws.row_dimensions[3].height = 22

    # Row 4: Column headers
    HEADERS = [
        "N°", "Réf. client", "Nom du client",
        "Nb produits achetés (Q2)",
        "Nb mois actifs (Q2)",
        "Fréquence (mois/3)",
        "Σ moyennes mensuelles (t)",
        "Perte Q1 - Volume (t)",
        "Perte Q1 - CA (FCFA)",
        "CA HT total 6 mois",
        "20/80",
    ]
    for col_idx, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=4, column=col_idx, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT
        c.alignment = HEADER_ALIGN; c.border = BORDER
    ws.row_dimensions[4].height = 40

    # Data rows
    START_ROW = 5
    for i, key in enumerate(sorted_clients, start=1):
        r = START_ROW + i - 1
        ref, name = key
        L = losses_dict[key]
        banding = (i % 2 == 0)

        # N°
        c = ws.cell(row=r, column=1, value=i)
        c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
        if banding: c.fill = BAND_FILL

        # Réf
        c = ws.cell(row=r, column=2, value=ref)
        c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
        if banding: c.fill = BAND_FILL

        # Nom
        c = ws.cell(row=r, column=3, value=name)
        c.font = BODY_FONT; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
        if banding: c.fill = BAND_FILL

        # Nb produits achetés (Q2)
        c = ws.cell(row=r, column=4, value=L["n_products"])
        c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
        if banding: c.fill = BAND_FILL

        # Nb mois actifs (Q2)
        c = ws.cell(row=r, column=5, value=L["n_active_months"])
        c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
        if banding: c.fill = BAND_FILL

        # Fréquence
        c = ws.cell(row=r, column=6, value=L["frequency"])
        c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
        c.number_format = '0.00'
        if banding: c.fill = BAND_FILL

        # Σ moyennes mensuelles (t)
        c = ws.cell(row=r, column=7, value=L["sum_monthly_avg_tons"])
        c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
        c.number_format = VOL_NUM_FMT
        if banding: c.fill = BAND_FILL

        # Perte Q1 - Volume (t) - HIGHLIGHTED
        c = ws.cell(row=r, column=8, value=L["loss_tons"])
        c.font = LOSS_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
        c.number_format = VOL_NUM_FMT
        c.fill = LOSS_HIGHLIGHT_FILL

        # Perte Q1 - CA (FCFA) - HIGHLIGHTED
        c = ws.cell(row=r, column=9, value=L["loss_fcfa"])
        c.font = LOSS_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
        c.number_format = CA_NUM_FMT
        c.fill = LOSS_HIGHLIGHT_FILL

        # CA HT total 6 mois
        c = ws.cell(row=r, column=10, value=client_ca_total.get(key, 0.0))
        c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
        c.number_format = CA_NUM_FMT
        if banding: c.fill = BAND_FILL

        # 20/80 marker
        c = ws.cell(row=r, column=11)
        c.border = BORDER
        c.alignment = BODY_ALIGN_CENTER
        if key in pareto_clients:
            c.value = PARETO_SYMBOL
            c.fill = PARETO_FILL
            c.font = PARETO_FONT
        else:
            if banding: c.fill = BAND_FILL
            c.font = BODY_FONT

    # Column widths
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 42
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 13
    ws.column_dimensions['F'].width = 13
    ws.column_dimensions['G'].width = 17
    ws.column_dimensions['H'].width = 17
    ws.column_dimensions['I'].width = 22
    ws.column_dimensions['J'].width = 22
    ws.column_dimensions['K'].width = 9

    ws.freeze_panes = "D5"
    ws.auto_filter.ref = f"A4:K{START_ROW + n_clients - 1}"

    # Add a TOTAL row at the top (right after the headers)
    # Actually, let's add it at the bottom for visibility
    total_row = START_ROW + n_clients
    c = ws.cell(row=total_row, column=1, value="")
    ws.cell(row=total_row, column=2, value="TOTAL").font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    ws.cell(row=total_row, column=2).fill = HEADER_FILL
    ws.cell(row=total_row, column=2).alignment = BODY_ALIGN_CENTER
    ws.cell(row=total_row, column=2).border = BORDER
    # merge B+C
    ws.merge_cells(start_row=total_row, start_column=2, end_row=total_row, end_column=3)
    ws.cell(row=total_row, column=3).fill = HEADER_FILL
    ws.cell(row=total_row, column=3).border = BORDER

    total_n_products = sum(L["n_products"] for L in losses_dict.values())
    total_loss_tons = sum(L["loss_tons"] for L in losses_dict.values())
    total_loss_fcfa = sum(L["loss_fcfa"] for L in losses_dict.values())
    total_ca = sum(client_ca_total.get(k, 0) for k in client_list)
    total_avg = sum(L["sum_monthly_avg_tons"] for L in losses_dict.values())

    # Nb produits - leave blank (sum doesn't make sense)
    c = ws.cell(row=total_row, column=4); c.fill = HEADER_FILL; c.border = BORDER
    c = ws.cell(row=total_row, column=5); c.fill = HEADER_FILL; c.border = BORDER
    c = ws.cell(row=total_row, column=6); c.fill = HEADER_FILL; c.border = BORDER

    # Σ moyennes mensuelles (t)
    c = ws.cell(row=total_row, column=7, value=total_avg)
    c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = VOL_NUM_FMT

    # Perte Q1 - Volume
    c = ws.cell(row=total_row, column=8, value=total_loss_tons)
    c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
    c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = VOL_NUM_FMT

    # Perte Q1 - CA
    c = ws.cell(row=total_row, column=9, value=total_loss_fcfa)
    c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
    c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = CA_NUM_FMT

    # CA total 6 mois
    c = ws.cell(row=total_row, column=10, value=total_ca)
    c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = HEADER_FILL; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = CA_NUM_FMT

    c = ws.cell(row=total_row, column=11); c.fill = HEADER_FILL; c.border = BORDER

    return n_clients, total_loss_tons, total_loss_fcfa


# ---- Build workbook with 2 sheets ----
wb = Workbook()
wb.properties.creator = "Z.ai"

# Sheet 1: Zero achat global (16 produits) — loss on ciblé
ws1 = wb.active
ws1.title = "Zero global (16)"
n1, loss_tons_1, loss_fcfa_1 = build_sheet(
    ws1,
    title_text="BELGOCAM SA - Pertes Q1 estimées — Clients 'zéro achat global' (16 produits ciblés)",
    sub_text=("Critère : client n'ayant acheté AUCUN des 16 produits ciblés (4 tourteaux de soja + "
              "10 BELGO 10% + 2 BELGO 5%) en janvier, février ou mars 2026.\n"
              "Méthode : Perte Q1 = Σ(moyennes mensuelles par produit en Q2) × Fréquence × 3 mois.\n"
              "  - Moyenne mensuelle (par produit) = volume total Q2 / nombre de mois Q2 où le produit a été acheté.\n"
              "  - Fréquence = (nombre de mois Q2 avec achat ciblé) / 3."),
    client_list=zero_global,
    losses_dict=losses_global,
)

# Sheet 2: Zero achat concentrés (12 produits)
ws2 = wb.create_sheet("Zero concentrés (12)")
n2, loss_tons_2, loss_fcfa_2 = build_sheet(
    ws2,
    title_text="BELGOCAM SA - Pertes Q1 estimées — Clients 'zéro achat concentrés' (12 produits)",
    sub_text=("Critère : client n'ayant acheté AUCUN concentré (10 BELGO 10% + 2 BELGO 5%) en "
              "janvier, février ou mars 2026.\n"
              "Méthode : Perte Q1 = Σ(moyennes mensuelles par produit en Q2) × Fréquence × 3 mois.\n"
              "  - Moyenne mensuelle (par produit) = volume total Q2 / nombre de mois Q2 où le produit a été acheté.\n"
              "  - Fréquence = (nombre de mois Q2 avec achat concentré) / 3."),
    client_list=zero_concentre,
    losses_dict=losses_concentre,
)

wb.save(OUT)
print(f"  Saved: {OUT}")
print(f"  File size: {os.path.getsize(OUT):,} bytes")

# ---- Summary -------------------------------------------------------------
print(f"\n{'='*70}")
print(f"SUMMARY")
print(f"{'='*70}")
print(f"Zero achat global ({n1} clients):")
print(f"  Total perte Q1 (volume) : {loss_tons_1:>10.2f} t")
print(f"  Total perte Q1 (CA)     : {loss_fcfa_1:>15,.0f} FCFA")
print(f"Zero achat concentrés ({n2} clients):")
print(f"  Total perte Q1 (volume) : {loss_tons_2:>10.2f} t")
print(f"  Total perte Q1 (CA)     : {loss_fcfa_2:>15,.0f} FCFA")
