"""Generate 20/80 Pareto client list for OUEST region — April to today (mid-Sept 2026).
Combines historical dataset (Apr-Aug 2026) + September extraction (36).xlsx.

Definition: "20/80" = top clients contributing ~80% of total CA (Pareto principle).
Output: Excel file with full client list + top 20/80 marked.
"""
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import defaultdict
from datetime import datetime, date
import os

# === Config ===
DATASET = '/home/z/my-project/scripts/dataset_2023_2026.csv'
SEPT_EXTRACTION = '/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (36).xlsx'
OUT_XLSX = '/home/z/my-project/download/clients_20_80_ouest_avril_septembre_2026.xlsx'

# Agences Ouest (from AGENCE_MAP)
OUEST_AGENCES_ERP = ['AGENCE FAMLA', 'AGENCE DJELENG', 'AGENCE DE BAMENDA - DEPOT MBOUDA', 'SPC BAF-CHEFFERIE', 'SPC-DSCHANG']
OUEST_AGENCES_DS = ['Famla', 'Djeleng', 'Mbouda', 'Baf-Chefferie', 'Dschang']

# Internal clients to exclude
INTERNAL = ['SPC', 'PDC', 'EMANA']  # COMPTOIR inclus (clients comptoir à compter)

def is_internal(c):
    if not c: return False
    s = str(c).upper()
    return any(p in s for p in INTERNAL)

# === 1. Load historical dataset (April-August 2026) ===
print("Loading historical dataset...")
df = pd.read_csv(DATASET, parse_dates=['date'], low_memory=False)
print(f"  Total records: {len(df)}")

# Filter April 2026 onwards
df_apr = df[(df['date'].dt.year == 2026) & (df['date'].dt.month >= 4)].copy()
print(f"  April-August 2026 records: {len(df_apr)}")

# Filter Ouest region
df_ouest = df_apr[df_apr['region'] == 'Ouest'].copy()
print(f"  Ouest records: {len(df_ouest)}")

# Combined CA (TTC fallback to HT)
df_ouest['ca_combined'] = df_ouest['montant_ttc'].where(df_ouest['montant_ttc'] > 0, df_ouest['montant_ht'])

# Filter out internal clients
df_ouest_ext = df_ouest[~df_ouest['client'].apply(is_internal)].copy()
print(f"  Ouest records (external only): {len(df_ouest_ext)}")

# === 2. Load September extraction ===
print("\nLoading September extraction...")
wb = openpyxl.load_workbook(SEPT_EXTRACTION, read_only=True, data_only=True)
ws = wb['Sheet 1']

sept_records = []
for r in ws.iter_rows(min_row=3, values_only=True):
    if not r or r[0] == 'Total': continue
    if r[13] != 'Livrée': continue
    date_str = str(r[6])[:10] if r[6] else ''
    if '/09/2026' not in date_str: continue
    
    ag = r[15] if len(r) > 15 else ''
    if ag not in OUEST_AGENCES_ERP: continue
    
    client = r[5] if len(r) > 5 else ''
    if is_internal(client): continue
    
    ref = r[0]
    qte = float(r[2]) if r[2] else 0
    ca_ttc = float(r[9]) if r[9] else 0
    ca_ht = float(r[8]) if r[8] else 0
    ca = ca_ttc if ca_ttc > 0 else ca_ht
    
    sept_records.append({
        'client': client,
        'agence': ag.replace('AGENCE ', '').replace('DEPOT ', ''),
        'ref': ref,
        'qte': qte,
        'ca_combined': ca,
        'date_str': date_str,
    })

print(f"  September Ouest Livrées: {len(sept_records)}")

# Convert to DataFrame
df_sept = pd.DataFrame(sept_records)

# Normalize agence names to match dataset
AGENCE_NORMALIZE = {
    'FAMLA': 'Famla',
    'DJELENG': 'Djeleng',
    'BAMENDA - MBOUDA': 'Mbouda',
    'BAMENDA - DEPOT MBOUDA': 'Mbouda',
    'BAF-CHEFFERIE': 'Baf-Chefferie',
    'DSCHANG': 'Dschang',
}
df_sept['agence'] = df_sept['agence'].apply(lambda x: AGENCE_NORMALIZE.get(x, x))

# === 3. Aggregate by client ===
print("\nAggregating by client...")

# From historical
hist_agg = df_ouest_ext.groupby('client').agg(
    ca_combined=('ca_combined', 'sum'),
    tonnes=('tonnes', 'sum'),
    n_orders=('ref', 'count'),
    agences=('agence', lambda x: ', '.join(sorted(set(x)))),
    first_date=('date', 'min'),
    last_date=('date', 'max'),
).reset_index()

# From September
sept_agg = df_sept.groupby('client').agg(
    ca_combined=('ca_combined', 'sum'),
    n_orders=('ref', 'count'),
    agences=('agence', lambda x: ', '.join(sorted(set(x)))),
).reset_index()
sept_agg = sept_agg.rename(columns={'ca_combined': 'ca_sept', 'n_orders': 'n_orders_sept', 'agences': 'agences_sept'})

# Merge
merged = pd.merge(hist_agg, sept_agg, on='client', how='outer').fillna(0)
merged['ca_total'] = merged['ca_combined'] + merged['ca_sept']
merged['n_orders_total'] = merged['n_orders'] + merged['n_orders_sept']

# Use agences from hist if present, else from sept
merged['agences_final'] = merged.apply(lambda r: r['agences'] if r['agences'] != '' else r['agences_sept'], axis=1)

# Sort by CA descending
merged = merged.sort_values('ca_total', ascending=False).reset_index(drop=True)

# === 4. Apply Pareto 20/80 ===
print("\nApplying Pareto 20/80...")
total_ca = merged['ca_total'].sum()
merged['ca_cumul'] = merged['ca_total'].cumsum()
merged['pct_ca'] = merged['ca_total'] / total_ca * 100
merged['pct_ca_cumul'] = merged['ca_cumul'] / total_ca * 100

# Mark top clients contributing to 80% of CA
merged['is_20_80'] = merged['pct_ca_cumul'] <= 80
# Always include the first client that crosses 80%
first_cross_80 = merged[~merged['is_20_80']].index.min()
if not pd.isna(first_cross_80):
    merged.loc[first_cross_80, 'is_20_80'] = True

n_clients = len(merged)
n_20_80 = merged['is_20_80'].sum()
pct_clients_20_80 = n_20_80 / n_clients * 100
ca_20_80 = merged[merged['is_20_80']]['ca_total'].sum()
pct_ca_20_80 = ca_20_80 / total_ca * 100

print(f"  Total clients: {n_clients}")
print(f"  Total CA: {total_ca:,.0f} FCFA ({total_ca/1e6:.1f} M)")
print(f"  Top 20/80 clients: {n_20_80} ({pct_clients_20_80:.1f}% of clients)")
print(f"  CA from 20/80 clients: {ca_20_80:,.0f} FCFA ({pct_ca_20_80:.1f}% of CA)")

# === 5. Build Excel ===
print("\nBuilding Excel file...")

HEAD_FILL = PatternFill('solid', fgColor='1F4E78')
HEAD_FONT = Font(bold=True, color='FFFFFF', size=11)
TOTAL_FILL = PatternFill('solid', fgColor='FFF2CC')
TOTAL_FONT = Font(bold=True, size=11)
STAR_FILL = PatternFill('solid', fgColor='FFE699')
THIN = Side(border_style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb_out = openpyxl.Workbook()
wb_out.remove(wb_out.active)

# Sheet 1: Synthèse
ws = wb_out.create_sheet("1. Synthèse")
ws['A1'] = 'BELGOCAM SA - Analyse 20/80 Pareto — Région Ouest'
ws['A1'].font = Font(bold=True, size=14, color='1F4E78')
ws['A2'] = f'Période : Avril 2026 - Septembre 2026 (au {date.today().strftime("%d/%m/%Y")})'
ws['A2'].font = Font(italic=True, size=10, color='595959')
ws['A3'] = 'Région : Ouest (FAMLA, DJELENG, MBOUDA, SPC BAF-CHEFFERIE, SPC DSCHANG)'
ws['A3'].font = Font(italic=True, size=10, color='595959')

row = 5
ws.cell(row=row, column=1, value='SYNTHÈSE PARETO 20/80').font = Font(bold=True, color='1F4E78', size=12)
row += 1
headers = ['Indicateur', 'Valeur']
for i, h in enumerate(headers, 1): ws.cell(row=row, column=i, value=h)
for c in range(1, 3):
    cell = ws.cell(row=row, column=c)
    cell.fill = HEAD_FILL; cell.font = HEAD_FONT
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = BORDER
row += 1

synth_data = [
    ('Période analysée', 'Avril 2026 → Septembre 2026'),
    ('Région', 'Ouest'),
    ('Agences incluses', 'FAMLA, DJELENG, MBOUDA, SPC BAF-CHEFFERIE, SPC DSCHANG'),
    ('Clients internes exclus', 'SPC, PDC, EMANA (COMPTOIR inclus)'),
    ('Nombre total de clients', n_clients),
    ('CA total cumulé (FCFA)', f'{total_ca:,.0f}'.replace(',', ' ')),
    ('CA total cumulé (M FCFA)', f'{total_ca/1e6:.1f}'),
    ('Volume total cumulé (t)', f'{merged["tonnes"].sum():,.0f}'.replace(',', ' ')),
    ('', ''),
    ('Clients 20/80 (top)', n_20_80),
    ('% clients (sur total)', f'{pct_clients_20_80:.1f}%'),
    ('CA des 20/80 (M FCFA)', f'{ca_20_80/1e6:.1f}'),
    ('% CA total', f'{pct_ca_20_80:.1f}%'),
    ('', ''),
    ('Règle Pareto', 'Top clients contribuant à ~80% du CA'),
    ('Marquage Excel', 'Colonne "20/80" = OUI pour les top clients'),
]
for label, val in synth_data:
    ws.cell(row=row, column=1, value=label).font = Font(bold=True)
    ws.cell(row=row, column=2, value=val)
    ws.cell(row=row, column=1).border = BORDER
    ws.cell(row=row, column=2).border = BORDER
    row += 1

ws.column_dimensions['A'].width = 35
ws.column_dimensions['B'].width = 55

# Sheet 2: Liste complète (all clients)
ws2 = wb_out.create_sheet("2. Liste complète")
ws2['A1'] = 'BELGOCAM SA - Liste complète des clients Ouest (Avril-Septembre 2026)'
ws2['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers = ['Rang', '20/80', 'Client', 'CA total (FCFA)', 'CA (M FCFA)', '% CA', '% CA cumul', 'Tonnes', 'Nb cmdes', 'Agences', 'Premier achat', 'Dernier achat']
for i, h in enumerate(headers, 1): ws2.cell(row=row, column=i, value=h)
for c in range(1, len(headers)+1):
    cell = ws2.cell(row=row, column=c)
    cell.fill = HEAD_FILL; cell.font = HEAD_FONT
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = BORDER
row += 1

for idx, r in merged.iterrows():
    is_20 = bool(r['is_20_80'])
    ws2.cell(row=row, column=1, value=idx + 1)
    ws2.cell(row=row, column=2, value='★ OUI' if is_20 else '')
    ws2.cell(row=row, column=3, value=r['client'])
    ws2.cell(row=row, column=4, value=round(r['ca_total']))
    ws2.cell(row=row, column=5, value=round(r['ca_total']/1e6, 2))
    ws2.cell(row=row, column=6, value=round(r['pct_ca'], 2))
    ws2.cell(row=row, column=7, value=round(r['pct_ca_cumul'], 2))
    ws2.cell(row=row, column=8, value=round(r['tonnes'], 1) if r['tonnes'] > 0 else 0)
    ws2.cell(row=row, column=9, value=int(r['n_orders_total']))
    ws2.cell(row=row, column=10, value=r['agences_final'])
    first_date = r['first_date']
    if hasattr(first_date, 'strftime'):
        ws2.cell(row=row, column=11, value=first_date.strftime('%d/%m/%Y'))
    else:
        ws2.cell(row=row, column=11, value='')
    ws2.cell(row=row, column=12, value=r['date_str'] if 'date_str' in r and pd.notna(r.get('date_str')) else 'Sept 2026')
    
    # Color 20/80 rows
    for c in range(1, len(headers)+1):
        cell = ws2.cell(row=row, column=c)
        cell.border = BORDER
        if is_20:
            cell.fill = STAR_FILL
            if c == 2:
                cell.font = Font(bold=True, color='C00000')
    row += 1

# Total row
ws2.cell(row=row, column=1, value='TOTAL')
ws2.cell(row=row, column=2, value=f'{n_20_80} clients')
ws2.cell(row=row, column=4, value=round(total_ca))
ws2.cell(row=row, column=5, value=round(total_ca/1e6, 1))
ws2.cell(row=row, column=6, value=100.0)
ws2.cell(row=row, column=8, value=round(merged['tonnes'].sum(), 1))
ws2.cell(row=row, column=9, value=int(merged['n_orders_total'].sum()))
for c in range(1, len(headers)+1):
    cell = ws2.cell(row=row, column=c)
    cell.fill = TOTAL_FILL; cell.font = TOTAL_FONT; cell.border = BORDER

# Column widths
widths = [6, 8, 38, 16, 12, 10, 12, 10, 10, 30, 14, 14]
for i, w in enumerate(widths, 1):
    ws2.column_dimensions[chr(64+i)].width = w
ws2.freeze_panes = 'D4'

# Sheet 3: Top 20/80 only
ws3 = wb_out.create_sheet("3. Top 20-80")
ws3['A1'] = f'Top {n_20_80} clients 20/80 Ouest — {pct_ca_20_80:.1f}% du CA'
ws3['A1'].font = Font(bold=True, size=14, color='1F4E78')

row = 3
headers = ['Rang', 'Client', 'CA (M FCFA)', '% CA', '% CA cumul', 'Tonnes', 'Nb cmdes', 'Agences', 'Premier achat']
for i, h in enumerate(headers, 1): ws3.cell(row=row, column=i, value=h)
for c in range(1, len(headers)+1):
    cell = ws3.cell(row=row, column=c)
    cell.fill = HEAD_FILL; cell.font = HEAD_FONT
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = BORDER
row += 1

rank = 0
for idx, r in merged[merged['is_20_80']].iterrows():
    rank += 1
    ws3.cell(row=row, column=1, value=rank)
    ws3.cell(row=row, column=2, value=r['client'])
    ws3.cell(row=row, column=3, value=round(r['ca_total']/1e6, 2))
    ws3.cell(row=row, column=4, value=round(r['pct_ca'], 2))
    ws3.cell(row=row, column=5, value=round(r['pct_ca_cumul'], 2))
    ws3.cell(row=row, column=6, value=round(r['tonnes'], 1) if r['tonnes'] > 0 else 0)
    ws3.cell(row=row, column=7, value=int(r['n_orders_total']))
    ws3.cell(row=row, column=8, value=r['agences_final'])
    first_date = r['first_date']
    if hasattr(first_date, 'strftime'):
        ws3.cell(row=row, column=9, value=first_date.strftime('%d/%m/%Y'))
    
    for c in range(1, len(headers)+1):
        cell = ws3.cell(row=row, column=c)
        cell.border = BORDER
        cell.fill = STAR_FILL
    row += 1

# Total
ws3.cell(row=row, column=1, value='TOTAL')
ws3.cell(row=row, column=2, value=f'{n_20_80} clients')
ws3.cell(row=row, column=3, value=round(ca_20_80/1e6, 1))
ws3.cell(row=row, column=4, value=round(pct_ca_20_80, 1))
ws3.cell(row=row, column=5, value=100.0)
ws3.cell(row=row, column=6, value=round(merged[merged['is_20_80']]['tonnes'].sum(), 1))
ws3.cell(row=row, column=7, value=int(merged[merged['is_20_80']]['n_orders_total'].sum()))
for c in range(1, len(headers)+1):
    cell = ws3.cell(row=row, column=c)
    cell.fill = TOTAL_FILL; cell.font = TOTAL_FONT; cell.border = BORDER

widths3 = [6, 38, 14, 10, 12, 10, 10, 30, 14]
for i, w in enumerate(widths3, 1):
    ws3.column_dimensions[chr(64+i)].width = w
ws3.freeze_panes = 'C4'

wb_out.save(OUT_XLSX)
print(f"\n=== EXCEL SAVED ===")
print(f"Path: {OUT_XLSX}")
print(f"Size: {os.path.getsize(OUT_XLSX)/1024:.0f} KB")
print(f"Sheets: {wb_out.sheetnames}")
