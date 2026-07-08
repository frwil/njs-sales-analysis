import openpyxl

# Recherche exhaustive de PONT BASCULE dans tous les fichiers Excel
files = [
    '/home/z/my-project/upload/ventes janv a juin 2026.xlsx',
    '/home/z/my-project/upload/86d96135-9db7-45bc-bba6-a69efa2c5ee3.xlsx',  # 2025
    '/home/z/my-project/upload/objectifs 2026 - Takou.xlsx',
    '/home/z/my-project/upload/DOC-20260706-WA0017.xlsx',
]

for path in files:
    print(f"\n=== {path.split('/')[-1]} ===")
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            count = 0
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    if cell and "BASCULE" in str(cell).upper():
                        count += 1
                        if count <= 3:
                            print(f"  Feuille '{sheet_name}' - Ligne: {row[:5]}")
                            break
                if count > 3:
                    break
            if count > 0:
                print(f"  → {count} occurrence(s) 'BASCULE' dans '{sheet_name}'")
        wb.close()
    except Exception as e:
        print(f"  Erreur: {e}")
