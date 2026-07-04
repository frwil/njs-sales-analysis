"""
Find all CLIENT COMPTOIR variants (agences, suffixes) to exclude.
Look for any tier containing 'COMPTOIR' as a client name pattern.
"""
from openpyxl import load_workbook
import re
import json

SRC = "/home/z/my-project/download/ventes_livrees.xlsx"
wb = load_workbook(SRC, read_only=True, data_only=True)

# Pattern: any tier with "CLIENT COMPTOIR" anywhere, OR "COMPTOIR" as standalone name
# This catches: "CLIENT COMPTOIR", "CLIENT COMPTOIR YAOUNDE", "CLIENT COMPTOIR SIEGE", etc.
COMPTOIR_RE = re.compile(r"CLIENT\s+COMPTOIR", re.IGNORECASE)

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

# Find all CLIENT COMPTOIR variants
comptoir_clients = []
print(f"Total unique clients: {len(clients)}")
print(f"\nAll CLIENT COMPTOIR variants found:")
print(f"{'Tiers':<80} {'CA HT 6 mois':>15}")
print("-" * 100)
for tiers, ca in sorted(clients.items(), key=lambda x: -x[1]):
    if COMPTOIR_RE.search(str(tiers)):
        comptoir_clients.append(tiers)
        print(f"{str(tiers)[:78]:<80} {ca:>15,.0f}")

print(f"\nTotal CLIENT COMPTOIR variants: {len(comptoir_clients)}")

# Merge with existing excluded list
with open("/home/z/my-project/scripts/excluded_clients.json", "r", encoding="utf-8") as f:
    existing_excluded = set(json.load(f))

# Print what's already in the list vs new
already_in = [t for t in comptoir_clients if t in existing_excluded]
new_to_add = [t for t in comptoir_clients if t not in existing_excluded]
print(f"\nAlready in excluded list: {len(already_in)}")
print(f"NEW to add: {len(new_to_add)}")
if new_to_add:
    print("\nNew CLIENT COMPTOIR variants to add:")
    for t in new_to_add:
        print(f"  - {t}")

# Save merged list
new_excluded = sorted(set(existing_excluded) | set(comptoir_clients))
with open("/home/z/my-project/scripts/excluded_clients.json", "w", encoding="utf-8") as f:
    json.dump(new_excluded, f, ensure_ascii=False, indent=2)
print(f"\nTotal excluded list size: {len(new_excluded)} (was {len(existing_excluded)})")
print(f"Saved to: /home/z/my-project/scripts/excluded_clients.json")
