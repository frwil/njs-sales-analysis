"""
Zero-achat on CONCENTRÉS only (BELGO 10% + BELGO 5%, 12 products, no soja).

Concentré products (12):
- BELGO 10% (10): C102, C1022, C104, C1042, C1043, C1044, C105, C1053, C1054, C1055
- BELGO 5%  (2):  C101, C103

Soja products (4) — for cross-reference:
- T102, T1021, T1023, T1024

Output: /home/z/my-project/download/zero_achat_concentres.xlsx
Columns:
- N°
- Réf. client
- Nom du client
- Achat concentrés Q1 ?    (always "Non" — that's the filter)
- Achat soja Q1 ?          ("Oui" if client bought soja in Q1, "Non" otherwise)
- Achat ciblé 6 mois ?     (soja OR concentrés, any month)
- CA HT Q1 (tous produits)
- CA HT total 6 mois
- Catégorie 20/80
"""
import re
import os
import time
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import defaultdict

# ---- Paths ----
SRC = "/home/z/my-project/download/ventes_janv_juin_2026_livrees.xlsx"
OUT = "/home/z/my-project/download/zero_achat_concentres.xlsx"

# ---- Product categories ----
CONCENTRE_REFS = {
    "C102", "C1022", "C104", "C1042", "C1043", "C1044",   # BELGO 10% (6)
    "C105", "C1053", "C1054", "C1055",                    # BELGO 10% (4)
    "C101", "C103",                                         # BELGO 5% (2)
}
assert len(CONCENTRE_REFS) == 12

SOJA_REFS = {"T102", "T1021", "T1023", "T1024"}

MONTHS = [
    ("Sheet 1", "Janvier", 0),
    ("Feuil1",  "Février", 1),
    ("Feuil2",  "Mars",    2),
    ("Feuil3",  "Avril",   3),
    ("Feuil4",  "Mai",     4),
    ("Feuil5",  "Juin",    5),
]
FIRST_QUARTER_SHEETS = {"Sheet 1", "Feuil1", "Feuil2"}

TIERS_RE = re.compile(r"^\s*([A-Z]{2,}[\w-]*?)\s*-\s+(.+?)\s*$")

# Règle : l'analyse zéro achat exclut les clients internes (COMPTOIR, SPC, PDC, EMANA)
INTERNAL_PATTERNS = ('COMPTOIR', 'SPC', 'PDC', 'EMANA')

def is_internal(tiers_str):
    if tiers_str is None:
        return False
    s = str(tiers_str).upper()
    return any(p in s for p in INTERNAL_PATTERNS)

def split_tiers(tiers):
    if tiers is None:
        return ("", "")
    s = str(tiers).strip()
    m = TIERS_RE.match(s)
    if m:
        return (m.group(1).strip(), m.group(2).strip())
    return ("", s)

# ---- 1. Read source & aggregate per-client data ---------------------------
print(f"Reading source: {SRC}")
t0 = time.time()
wb_src = load_workbook(SRC, read_only=True, data_only=True)
print(f"  Loaded in {time.time()-t0:.1f}s")

client_ca_total        = defaultdict(float)
client_ca_q1           = defaultdict(float)
client_concentre_q1    = defaultdict(bool)  # bought any concentré in Q1
client_concentre_any   = defaultdict(bool)  # bought any concentré in any month
client_soja_q1         = defaultdict(bool)  # bought any soja in Q1

for sheet_name, month_label, m_idx in MONTHS:
    ws = wb_src[sheet_name]
    n = 0
    for row in ws.iter_rows(min_row=3, values_only=True):
        tiers = row[5] if len(row) > 5 else None
        if tiers is None or str(tiers).strip() == "":
            continue
        if is_internal(tiers):
            continue  # clients internes exclus de l'analyse zéro achat
        n += 1
        ref_prod = row[0] if len(row) > 0 else None
        ca_ht    = row[8] if len(row) > 8 else 0
        ref_c, name_c = split_tiers(tiers)
        key = (ref_c, name_c)

        try:
            ca_ht_f = float(ca_ht) if ca_ht is not None else 0.0
        except (ValueError, TypeError):
            ca_ht_f = 0.0

        client_ca_total[key] += ca_ht_f

        is_concentre = (ref_prod is not None and str(ref_prod).strip() in CONCENTRE_REFS)
        is_soja      = (ref_prod is not None and str(ref_prod).strip() in SOJA_REFS)

        if is_concentre:
            client_concentre_any[key] = True
            if sheet_name in FIRST_QUARTER_SHEETS:
                client_concentre_q1[key] = True
        if is_soja and sheet_name in FIRST_QUARTER_SHEETS:
            client_soja_q1[key] = True
        if sheet_name in FIRST_QUARTER_SHEETS:
            client_ca_q1[key] += ca_ht_f
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

print(f"Pareto clients: {len(pareto_clients)} ({len(pareto_clients)/len(clients_all)*100:.1f}% of clients → 80% of CA)")

# ---- 3. Build zero-achat concentrés list ---------------------------------
# A client is "zero achat concentrés Q1" if client_concentre_q1[key] is False
zero_concentre_clients = [k for k in clients_all if not client_concentre_q1[k]]
zero_concentre_clients.sort(key=lambda k: (k[1].upper(), k[0]))

# Cross-tab with soja purchase in Q1
bought_soja_only = sum(1 for k in zero_concentre_clients if client_soja_q1[k])
bought_nothing_targeted = sum(1 for k in zero_concentre_clients if not client_soja_q1[k])
acquired_concentre_after_q1 = sum(1 for k in zero_concentre_clients if client_concentre_any[k])
pareto_in_zero = sum(1 for k in zero_concentre_clients if k in pareto_clients)

print(f"\n{'='*70}")
print(f"ZERO ACHAT CONCENTRÉS (Jan-Mars)")
print(f"{'='*70}")
print(f"  Total unique clients                       : {len(clients_all)}")
print(f"  Clients WITH concentré purchase in Q1      : {len(clients_all) - len(zero_concentre_clients)}")
print(f"  Clients 'zero achat concentrés' Q1         : {len(zero_concentre_clients)}")
print(f"    dont ont acheté du SOJA en Q1            : {bought_soja_only}")
print(f"    dont n'ont rien acheté de ciblé en Q1    : {bought_nothing_targeted}")
print(f"    dont acquisition concentrés APRÈS mars   : {acquired_concentre_after_q1}")
print(f"    dont font partie du 20/80                : {pareto_in_zero}")

# ---- 4. Build output workbook -------------------------------------------
print(f"\n{'='*70}")
print(f"Building output: {OUT}")
print(f"{'='*70}")

wb = Workbook()
wb.properties.creator = "Z.ai"
ws = wb.active
ws.title = "Zero concentrés Q1"

# ---- Styles ----
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E78")
TITLE_ALIGN = Alignment(horizontal="left", vertical="center")
SUB_FONT = Font(name="Calibri", size=10, italic=True, color="595959")
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
BAND_FILL = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
BODY_FONT = Font(name="Calibri", size=10)
BODY_ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
BODY_ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
BODY_ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
PARETO_FILL = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
PARETO_FONT = Font(name="Calibri", size=14, bold=True, color="7F6000")
PARETO_SYMBOL = "★"
CA_NUM_FMT = '#,##0" FCFA"'
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Row 1: Title
ws.cell(row=1, column=1,
        value=("BELGOCAM SA - Clients 'zéro achat concentrés' Janvier-Mars 2026 "
               "(BELGO 10% + BELGO 5%, 12 produits)")
).font = TITLE_FONT
ws.cell(row=1, column=1).alignment = TITLE_ALIGN
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
ws.row_dimensions[1].height = 26

# Row 2: Criteria explanation
ws.cell(row=2, column=1,
        value=("Critère : client n'ayant acheté AUCUN concentré (10 BELGO 10% + 2 BELGO 5%) "
               "en janvier, février ou mars 2026. La colonne 'Achat soja Q1 ?' permet de "
               "distinguer ceux qui ont quand même acheté du soja de ceux qui n'ont rien acheté "
               "de ciblé.")
).font = SUB_FONT
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=9)
ws.row_dimensions[2].height = 32

# Row 3: Headers
HEADERS = ["N°", "Réf. client", "Nom du client",
           "Achat concentrés Q1 ?", "Achat soja Q1 ?",
           "Achat concentrés 6 mois ?",
           "CA HT Q1 (tous produits)", "CA HT total 6 mois",
           "Catégorie 20/80"]
for col_idx, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=3, column=col_idx, value=h)
    c.fill = HEADER_FILL; c.font = HEADER_FONT
    c.alignment = HEADER_ALIGN; c.border = BORDER
ws.row_dimensions[3].height = 40

# Data rows
for i, key in enumerate(zero_concentre_clients, start=1):
    r = 3 + i
    ref, name = key
    banding = (i % 2 == 0)

    # N°
    c = ws.cell(row=r, column=1, value=i)
    c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL

    # Réf. client
    c = ws.cell(row=r, column=2, value=ref)
    c.font = BODY_FONT; c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL

    # Nom du client
    c = ws.cell(row=r, column=3, value=name)
    c.font = BODY_FONT; c.alignment = BODY_ALIGN_LEFT; c.border = BORDER
    if banding: c.fill = BAND_FILL

    # Achat concentrés Q1 ? — always "Non" (filter)
    c = ws.cell(row=r, column=4, value="Non")
    c.font = Font(name="Calibri", size=10, color="C00000", bold=True)
    c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL

    # Achat soja Q1 ? — Oui / Non
    soja_q1 = client_soja_q1.get(key, False)
    c = ws.cell(row=r, column=5, value=("Oui" if soja_q1 else "Non"))
    if soja_q1:
        c.font = Font(name="Calibri", size=10, color="006100", bold=True)
    else:
        c.font = Font(name="Calibri", size=10, color="C00000", bold=True)
    c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL

    # Achat concentrés 6 mois ?
    conc_any = client_concentre_any.get(key, False)
    c = ws.cell(row=r, column=6, value=("Oui" if conc_any else "Non"))
    if conc_any:
        c.font = Font(name="Calibri", size=10, color="006100", bold=True)
    else:
        c.font = Font(name="Calibri", size=10, color="C00000", bold=True)
    c.alignment = BODY_ALIGN_CENTER; c.border = BORDER
    if banding: c.fill = BAND_FILL

    # CA HT Q1
    c = ws.cell(row=r, column=7, value=client_ca_q1.get(key, 0.0))
    c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = CA_NUM_FMT
    if banding: c.fill = BAND_FILL

    # CA HT total 6 mois
    c = ws.cell(row=r, column=8, value=client_ca_total.get(key, 0.0))
    c.font = BODY_FONT; c.alignment = BODY_ALIGN_RIGHT; c.border = BORDER
    c.number_format = CA_NUM_FMT
    if banding: c.fill = BAND_FILL

    # 20/80 marker
    c = ws.cell(row=r, column=9)
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
ws.column_dimensions['C'].width = 45
ws.column_dimensions['D'].width = 18
ws.column_dimensions['E'].width = 15
ws.column_dimensions['F'].width = 20
ws.column_dimensions['G'].width = 22
ws.column_dimensions['H'].width = 22
ws.column_dimensions['I'].width = 14

ws.freeze_panes = "D4"
ws.auto_filter.ref = f"A3:I{3 + len(zero_concentre_clients)}"

wb.save(OUT)
print(f"  Saved: {OUT}")
print(f"  File size: {os.path.getsize(OUT):,} bytes")
print(f"  Rows: {len(zero_concentre_clients)} clients")
