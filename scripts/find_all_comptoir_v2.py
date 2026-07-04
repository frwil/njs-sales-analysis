"""
Refined: capture both 'CLIENT COMPTOIR' (singular) and 'CLIENTS COMPTOIR' (plural)
plus any other variants.
"""
from openpyxl import load_workbook
import re
import json

SRC = "/home/z/my-project/download/ventes_livrees.xlsx"
wb = load_workbook(SRC, read_only=True, data_only=True)

# Broader pattern: CLIENT(S) COMPTOIR (any case, any number of S)
# Also capture bare "COMPTOIR" + (agence name) patterns
COMPTOIR_RE = re.compile(r"CLIENTS?\s+COMPTOIR", re.IGNORECASE)

clients = {}
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    for row in ws.iter_rows(min_row=3, values_only=True):
        tiers = row[5] if len(row) > 5 else None
        if tiers is None or str(tiers).strip() == "":
            continue
        if tiers not in clients:
            clients[tiers] = 0
        ca = row[8] if len(row) > 8 else 0
        try:
            clients[tiers] += float(ca) if ca is not None else 0
        except:
            pass

wb.close()

# Find all matches
comptoir_clients = []
print(f"Total unique clients: {len(clients)}")
print(f"\nAll CLIENT(S) COMPTOIR variants:")
print(f"{'Tiers':<80} {'CA HT 6 mois':>15}")
print("-" * 100)
for tiers, ca in sorted(clients.items(), key=lambda x: -x[1]):
    if COMPTOIR_RE.search(str(tiers)):
        comptoir_clients.append(tiers)
        print(f"{str(tiers)[:78]:<80} {ca:>15,.0f}")

print(f"\nTotal: {len(comptoir_clients)}")

# Merge with existing
with open("/home/z/my-project/scripts/excluded_clients.json", "r", encoding="utf-8") as f:
    existing = set(json.load(f))

new = [t for t in comptoir_clients if t not in existing]
print(f"\nAlready excluded: {len(comptoir_clients) - len(new)}")
print(f"NEW to add: {len(new)}")
for t in new:
    print(f"  + {t}")

merged = sorted(set(existing) | set(comptoir_clients))
with open("/home/z/my-project/scripts/excluded_clients.json", "w", encoding="utf-8") as f:
    json.dump(merged, f, ensure_ascii=False, indent=2)
print(f"\nTotal excluded list: {len(merged)}")
