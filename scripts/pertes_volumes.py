"""
Compute volume losses (in tons) for each zero-achat client.

Volumes are computed in 3 categories:
- All products (tous produits)
- Targeted (soja + concentrés, 16 products)
- Concentrés only (12 products: BELGO 10% + BELGO 5%)

And 3 periods:
- Q1 (Jan-Mar) - the "loss" period for zero-achat clients
- Q2 (Apr-Jun)
- 6 months total

Output: ONE file with 2 sheets:
1. "Zero achat global (16)"  -> 352 clients (no targeted product in Q1)
2. "Zero achat concentrés"   -> 544 clients (no concentré in Q1)

Each sheet shows volumes (tons) by category and period, sorted by
total 6-month volume (all products) descending so biggest losses appear first.
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
OUT = "/home/z/my-project/download/pertes_volumes_zero_achat.xlsx"

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
        return val  # 1L ≈ 1 kg
    return 0.0

# ---- Months --------------------------------------------------------------
MONTHS = [
    ("Sheet 1", "Janvier", 0),
    ("Feuil1",  "Février", 1),
    ("Feuil2",  "Mars",    2),
    ("Feuil3",  "Avril",   3),
    ("Feuil4",  "Mai",     4),
    ("Feuil5",  "Juin",    5),
]
Q1_SHEETS = {"Sheet 1", "Feuil1", "Feuil2"}

TIERS_RE = re.compile(r"^\s*([A-Z]{2,}[\w-]*?)\s*-\s+(.+?)\s*$")

def split_tiers(tiers):
    if tiers is None:
        return ("", "")
    s = str(tiers).strip()
    m = TIERS_RE.match(s)
    if m:
        return (m.group(1).strip(), m.group(2).strip())
    return ("", s)

# ---- 1. Read source & build per-client aggregates ------------------------
print(f"Reading source: {SRC}")
t0 = time.time()
wb_src = load_workbook(SRC, read_only=True, data_only=True)
print(f"  Loaded in {time.time()-t0:.1f}s")

# Build product weight map (ref -> kg per unit)
product_weight = {}

# Per-client aggregates (all in KG; we'll convert to tons at output time)
client_vol_total_q1     = defaultdict(float)  # all products, Q1
client_vol_total_q2     = defaultdict(float)  # all products, Q2
client_vol_targeted_q1  = defaultdict(float)
client_vol_targeted_q2  = defaultdict(float)
client_vol_concentre_q1 = defaultdict(float)
client_vol_concentre_q2 = defaultdict(float)
client_ca_total         = defaultdict(float)
client_targeted_q1_flag = defaultdict(bool)   # bought any targeted product in Q1
client_concentre_q1_flag= defaultdict(bool)  # bought any concentré in Q1

for sheet_name, month_label, m_idx in MONTHS:
    ws = wb_src[sheet_name]
    n = 0
    for row in ws.iter_rows(min_row=3, values_only=True):
        tiers = row[5] if len(row) > 5 else None
        if tiers is None or str(tiers).strip() == "":
            continue
        n += 1
        ref_prod = row[0] if len(row) > 0 else None
        desc     = row[1] if len(row) > 1 else None
        qte      = row[2] if len(row) > 2 else 0       # Qté commandée (in bags/units)
        ca_ht    = row[8] if len(row) > 8 else 0

        ref_c, name_c = split_tiers(tiers)
        key = (ref_c, name_c)

        # Build/lookup product weight
        ref_prod_str = str(ref_prod).strip() if ref_prod is not None else ""
        if ref_prod_str and ref_prod_str not in product_weight:
            product_weight[ref_prod_str] = parse_weight_kg(desc)
        weight_kg = product_weight.get(ref_prod_str, 0.0)

        # Quantity
        try:
            qte_f = float(qte) if qte is not None else 0.0
        except (ValueError, TypeError):
            qte_f = 0.0

        vol_kg = qte_f * weight_kg

        # CA
        try:
            ca_f = float(ca_ht) if ca_ht is not None else 0.0
        except (ValueError, TypeError):
            ca_f = 0.0

        # Aggregations
        if sheet_name in Q1_SHEETS:
            client_vol_total_q1[key] += vol_kg
            if ref_prod_str in TARGET_REFS:
                client_vol_targeted_q1[key] += vol_kg
                client_targeted_q1_flag[key] = True
            if ref_prod_str in CONCENTRE_REFS:
                client_vol_concentre_q1[key] += vol_kg
                client_concentre_q1_flag[key] = True
        else:
            client_vol_total_q2[key] += vol_kg
            if ref_prod_str in TARGET_REFS:
                client_vol_targeted_q2[key] += vol_kg
            if ref_prod_str in CONCENTRE_REFS:
                client_vol_concentre_q2[key] += vol_kg

        client_ca_total[key] += ca_f

    print(f"  {month_label:<10} ({sheet_name}): {n} client occurrences")

wb_src.close()

clients_all = sorted(client_ca_total.keys(), key=lambda k: (k[1].upper(), k[0]))
print(f"\nTotal unique clients: {len(clients_all)}")

# ---- 2. Pareto 20/80 (recompute for self-containment) --------------------
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

# ---- 3. Build zero-achat lists -------------------------------------------
zero_global = [k for k in clients_all if not client_targeted_q1_flag[k]]
zero_concentre = [k for k in clients_all if not client_concentre_q1_flag[k]]

print(f"Zero achat global (16 produits)     : {len(zero_global)} clients")
print(f"Zero achat concentrés (12 produits) : {len(zero_concentre)} clients")

# ---- 4. Build output workbook --------------------------------------------
print(f"\nBuilding output: {OUT}")

# ---- Styles ----
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E78")
TITLE_ALIGN = Alignment(horizontal="left", vertical="center")
SUB_FONT = Font(name="Calibri", size=10, italic=True, color="595959")
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Group header (above column groups)
GROUP_FILL = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
GROUP_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
GROUP_ALIGN = Alignment(horizontal="center", vertical="center")

BAND_FILL = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
BODY_FONT = Font(name="Calibri", size=10)
BODY_ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
BODY_ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
BODY_ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
PARETO_FILL = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
PARETO_FONT = Font(name="Calibri", size=14, bold=True, color="7F6000")
PARETO_SYMBOL = "★"

# Highlight Q1 losses in red-ish
LOSS_FILL = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")

CA_NUM_FMT = '#,##0" FCFA"'
VOL_NUM_FMT = '#,##0.00" t"'
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def build_sheet(ws, title_text, sub_text, client_list):
    """Build one analysis sheet for a given zero-achat client list."""
    # Sort clients by total 6-month volume (all products) descending
    sorted_clients = sorted(
        client_list,
        key=lambda k: -(client_vol_total_q1[k] + client_vol_total_q2[k])
    )
    n_clients = len(sorted_clients)

    # Columns layout (1-based):
    # A: N°
    # B: Réf. client
    # C: Nom du client
    # D-F: Vol Q1 (tous / ciblé / concentrés)  -> "Perte Q1"
    # G-I: Vol Q2 (tous / ciblé / concentrés)
    # J-L: Vol 6 mois (tous / ciblé / concentrés)
    # M: CA HT 6 mois
    # N: 20/80

    # Row 1: Title (merge A1:N1)
    ws.cell(row=1, column=1, value=title_text).font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = TITLE_ALIGN
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=14)
    ws.row_dimensions[1].height = 26

    # Row 2: Subtitle / criteria
    ws.cell(row=2, column=1, value=sub_text).font = SUB_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=14)
    ws.row_dimensions[2].height = 30

    # Row 3: Group headers (merged across each group of 3 columns)
    # Q1 group
    ws.merge_cells(start_row=3, start_column=4, end_row=3, end_column=6)
    c = ws.cell(row=3, column=4, value="PERTE Q1 (Jan-Mar) — volumes en tonnes")
    c.fill = GROUP_FILL; c.font = GROUP_FONT; c.alignment = GROUP_ALIGN
    c.border = BORDER
    # Q2 group
    ws.merge_cells(start_row=3, start_column=7, end_row=3, end_column=9)
    c = ws.cell(row=3, column=7, value="Q2 (Avr-Juin) — volumes en tonnes")
    c.fill = GROUP_FILL; c.font = GROUP_FONT; c.alignment = GROUP_ALIGN
    c.border = BORDER
    # 6 months group
    ws.merge_cells(start_row=3, start_column=10, end_row=3, end_column=12)
    c = ws.cell(row=3, column=10, value="6 mois (Jan-Juin) — volumes en tonnes")
    c.fill = GROUP_FILL; c.font = GROUP_FONT; c.alignment = GROUP_ALIGN
    c.border = BORDER

    # Also fill the empty cells in the merged ranges with the same fill
    for col in [5, 6, 8, 9, 11, 12]:
        cc = ws.cell(row=3, column=col)
        cc.fill = GROUP_FILL
        cc.border = BORDER

    # Columns A, B, C, M, N on row 3: leave empty (merged with row 4 below)
    ws.row_dimensions[3].height = 22

    # Row 4: Column headers
    HEADERS = [
        "N°", "Réf. client", "Nom du client",
        "Tous produits", "Ciblé (soja+conc.)", "Concentrés",
        "Tous produits", "Ciblé (soja+conc.)", "Concentrés",
        "Tous produits", "Ciblé (soja+conc.)", "Concentrés",
        "CA HT 6 mois", "20/80",
    ]
    for col_idx, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=4, column=col_idx, value=h)
        c.fill = HEADER_FILL; c.font = HEADER_FONT
        c.alignment = HEADER_ALIGN; c.border = BORDER
    ws.row_dimensions[4].height = 36

    # Data rows
    START_ROW = 5
    for i, key in enumerate(sorted_clients, start=1):
        r = START_ROW + i - 1
        ref, name = key
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

        # Q1 volumes (cols 4-6) — convert kg to tons (÷1000)
        vol_q1_all = client_vol_total_q1[key] / 1000.0
        vol_q1_tar = client_vol_targeted_q1[key] / 1000.0
        vol_q1_con = client_vol_concentre_q1[key] / 1000.0

        for col_idx, val in [(4, vol_q1_all), (5, vol_q1_tar), (6, vol_q1_con)]:
            c = ws.cell(row=r, column=col_idx, value=val)
            c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
            c.number_format = VOL_NUM_FMT
            # Highlight loss cells (they should be 0 for the filtered category)
            c.fill = LOSS_FILL

        # Q2 volumes (cols 7-9)
        vol_q2_all = client_vol_total_q2[key] / 1000.0
        vol_q2_tar = client_vol_targeted_q2[key] / 1000.0
        vol_q2_con = client_vol_concentre_q2[key] / 1000.0
        for col_idx, val in [(7, vol_q2_all), (8, vol_q2_tar), (9, vol_q2_con)]:
            c = ws.cell(row=r, column=col_idx, value=val)
            c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
            c.number_format = VOL_NUM_FMT
            if banding: c.fill = BAND_FILL

        # 6 months volumes (cols 10-12)
        vol_6_all = vol_q1_all + vol_q2_all
        vol_6_tar = vol_q1_tar + vol_q2_tar
        vol_6_con = vol_q1_con + vol_q2_con
        for col_idx, val in [(10, vol_6_all), (11, vol_6_tar), (12, vol_6_con)]:
            c = ws.cell(row=r, column=col_idx, value=val)
            c.font = Font(name="Calibri", size=10, bold=True)
            c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
            c.number_format = VOL_NUM_FMT
            if banding: c.fill = BAND_FILL

        # CA HT 6 mois
        c = ws.cell(row=r, column=13, value=client_ca_total.get(key, 0.0))
        c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
        c.number_format = CA_NUM_FMT
        if banding: c.fill = BAND_FILL

        # 20/80 marker
        c = ws.cell(row=r, column=14)
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
    for col_letter in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
        ws.column_dimensions[col_letter].width = 16
    ws.column_dimensions['M'].width = 20
    ws.column_dimensions['N'].width = 9

    # Freeze panes (keep N°, Réf, Nom and headers visible)
    ws.freeze_panes = "D5"

    # Auto filter on the column header row
    ws.auto_filter.ref = f"A4:N{START_ROW + n_clients - 1}"

    return n_clients, sorted_clients


# ---- Build workbook with 2 sheets ----
wb = Workbook()
wb.properties.creator = "Z.ai"

# Sheet 1: Zero achat global (16 produits)
ws1 = wb.active
ws1.title = "Zero global (16)"
n1, sorted_global = build_sheet(
    ws1,
    title_text="BELGOCAM SA - Pertes en volumes (tonnes) — Clients 'zéro achat global' Janvier-Mars 2026",
    sub_text=("Critère : client n'ayant acheté AUCUN des 16 produits ciblés (4 tourteaux de soja + "
              "10 BELGO 10% + 2 BELGO 5%) en janvier, février ou mars 2026. "
              "Les cases roses 'PERTE Q1' devraient toutes être à 0 pour les colonnes ciblé et concentrés "
              "(par définition du filtre). La colonne 'Tous produits Q1' peut être > 0 si le client a "
              "acheté d'autres produits."),
    client_list=zero_global,
)

# Sheet 2: Zero achat concentrés (12 produits)
ws2 = wb.create_sheet("Zero concentrés (12)")
n2, sorted_concentre = build_sheet(
    ws2,
    title_text="BELGOCAM SA - Pertes en volumes (tonnes) — Clients 'zéro achat concentrés' Janvier-Mars 2026",
    sub_text=("Critère : client n'ayant acheté AUCUN concentré (10 BELGO 10% + 2 BELGO 5%) en "
              "janvier, février ou mars 2026. La colonne 'Concentrés Q1' (en rose) doit être à 0 "
              "(par définition du filtre). La colonne 'Ciblé Q1' peut être > 0 si le client a acheté "
              "du soja en Q1."),
    client_list=zero_concentre,
)

wb.save(OUT)
print(f"  Saved: {OUT}")
print(f"  File size: {os.path.getsize(OUT):,} bytes")

# ---- 5. Summary stats ----------------------------------------------------
print(f"\n{'='*70}")
print(f"SUMMARY")
print(f"{'='*70}")

# Top 5 pertes (Q2 volume) in zero_global — proxy for "what they could have bought in Q1"
print(f"\nTop 5 'pertes Q1' (proxy = volume Q2) — Zero achat global:")
print(f"  {'Client':<50} {'Vol Q2 tous (t)':>15} {'Vol Q2 ciblé (t)':>17} {'20/80':>6}")
for key in sorted_global[:5]:
    name = key[1][:48]
    v_q2_all = client_vol_total_q2[key] / 1000.0
    v_q2_tar = client_vol_targeted_q2[key] / 1000.0
    p = "★" if key in pareto_clients else ""
    print(f"  {name:<50} {v_q2_all:>15.2f} {v_q2_tar:>17.2f} {p:>6}")

print(f"\nTop 5 'pertes Q1' (proxy = volume Q2) — Zero achat concentrés:")
print(f"  {'Client':<50} {'Vol Q2 tous (t)':>15} {'Vol Q2 conc (t)':>17} {'20/80':>6}")
for key in sorted_concentre[:5]:
    name = key[1][:48]
    v_q2_all = client_vol_total_q2[key] / 1000.0
    v_q2_con = client_vol_concentre_q2[key] / 1000.0
    p = "★" if key in pareto_clients else ""
    print(f"  {name:<50} {v_q2_all:>15.2f} {v_q2_con:>17.2f} {p:>6}")

# Totals
total_vol_q2_global_all = sum(client_vol_total_q2[k] for k in zero_global) / 1000.0
total_vol_q2_global_tar = sum(client_vol_targeted_q2[k] for k in zero_global) / 1000.0
total_vol_q2_conc_all = sum(client_vol_total_q2[k] for k in zero_concentre) / 1000.0
total_vol_q2_conc_con = sum(client_vol_concentre_q2[k] for k in zero_concentre) / 1000.0

print(f"\nTotaux des 'pertes Q1' (proxy = volume Q2 acheté après la perte):")
print(f"  Zero achat global ({n1} clients):")
print(f"    Volume Q2 tous produits    : {total_vol_q2_global_all:>10.2f} t")
print(f"    Volume Q2 ciblé (soja+conc): {total_vol_q2_global_tar:>10.2f} t")
print(f"  Zero achat concentrés ({n2} clients):")
print(f"    Volume Q2 tous produits    : {total_vol_q2_conc_all:>10.2f} t")
print(f"    Volume Q2 concentrés       : {total_vol_q2_conc_con:>10.2f} t")
