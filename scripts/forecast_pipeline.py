#!/usr/bin/env python3
"""
Pipeline automatisé - Forecast Q4 2026 (Scénario S3)
Usage: python scripts/forecast_pipeline.py [chemin_extraction_aout]

Ce script:
1. Reconstruit le dataset consolidé (2025 + S1 2026 + juillet + août)
2. Recalcule les prix moyens par produit
3. Lance la modélisation Prophet + scénario S3
4. Génère l'Excel forecast multi-feuilles
5. Met à jour les graphiques de visualisation

Durée totale: ~5 minutes

Si aucun chemin n'est fourni, utilise le fichier le plus récent dans /home/z/my-project/upload/
"""
import os
import sys
import glob
import subprocess
import time
from datetime import datetime

UPLOAD_DIR = "/home/z/my-project/upload"
SCRIPTS_DIR = "/home/z/my-project/scripts"
DOWNLOAD_DIR = "/home/z/my-project/download"

def find_latest_extraction():
    """Find the most recent NJS GROUP ERP extraction in upload dir."""
    pattern = os.path.join(UPLOAD_DIR, "NJS GROUP ERP - Lignes de commandes + multicompany*.xlsx")
    files = glob.glob(pattern)
    if not files:
        print("❌ Aucune extraction trouvée dans", UPLOAD_DIR)
        sys.exit(1)
    # Sort by modification time
    files.sort(key=os.path.getmtime, reverse=True)
    return files[0]

def update_file_path(script_path, old_pattern, new_path):
    """Update a file path in a Python script."""
    with open(script_path, 'r') as f:
        content = f.read()
    
    # Replace any AOUT_SRC or SRC line that matches the pattern
    import re
    # Match patterns like: AOUT_SRC = "..." or SRC = '...'
    content = re.sub(
        r'(AOUT_SRC\s*=\s*["\'])[^"\']*["\']',
        f'AOUT_SRC = "{new_path}"',
        content
    )
    content = re.sub(
        r'(\nSRC\s*=\s*["\'])NJS GROUP ERP[^"\']*["\']',
        f'\nSRC = \'{new_path}\'',
        content
    )
    
    with open(script_path, 'w') as f:
        f.write(content)

def run_script(script_name, description):
    """Run a Python script and report status."""
    script_path = os.path.join(SCRIPTS_DIR, script_name)
    print(f"\n{'='*60}")
    print(f"▶ {description}")
    print(f"  Script: {script_name}")
    print(f"{'='*60}")
    
    start = time.time()
    result = subprocess.run(
        [sys.executable, script_path],
        capture_output=True, text=True, timeout=600,
        cwd="/home/z/my-project"
    )
    elapsed = time.time() - start
    
    if result.returncode == 0:
        # Show last 5 lines of output
        lines = result.stdout.strip().split('\n')
        for line in lines[-5:]:
            print(f"  {line}")
        print(f"  ✅ Terminé en {elapsed:.0f}s")
    else:
        print(f"  ❌ ERREUR (code {result.returncode}):")
        print(f"  {result.stderr[-500:]}")
        return False
    return True

def main():
    print("=" * 60)
    print("  FORECAST Q4 2026 - PIPELINE AUTOMATISÉ")
    print("  Scénario S3 | Sans Maïs | Avec Matériel Élevage + Premix")
    print("=" * 60)
    print(f"  Démarrage: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    
    # === Step 0: Find extraction file ===
    if len(sys.argv) > 1:
        aout_file = sys.argv[1]
        if not os.path.isabs(aout_file):
            aout_file = os.path.join(UPLOAD_DIR, aout_file)
    else:
        aout_file = find_latest_extraction()
    
    print(f"\n📄 Extraction août: {os.path.basename(aout_file)}")
    print(f"   Chemin: {aout_file}")
    print(f"   Taille: {os.path.getsize(aout_file) / 1024:.0f} KB")
    
    # === Step 1: Update file paths in scripts ===
    print(f"\n🔧 Mise à jour des chemins de fichiers...")
    
    # Update pace_01 (dataset preparation) - uses AOUT_SRC for the august file
    # The august file is referenced via FILE_AOUT in pace_01
    pace_01 = os.path.join(SCRIPTS_DIR, "pace_01_prepare_dataset.py")
    with open(pace_01, 'r') as f:
        content = f.read()
    import re
    content = re.sub(
        r'FILE_AOUT\s*=\s*["\'][^"\']*["\']',
        f'FILE_AOUT = "{aout_file}"',
        content
    )
    with open(pace_01, 'r') as f:
        content_v2 = f.read()
    # Also update pace_10 (v2 dataset with material)
    pace_10 = os.path.join(SCRIPTS_DIR, "pace_10_dataset_v2.py")
    with open(pace_10, 'r') as f:
        content = f.read()
    content = re.sub(
        r'FILE_AOUT\s*=\s*["\'][^"\']*["\']',
        f'FILE_AOUT = "{aout_file}"',
        content
    )
    with open(pace_10, 'w') as f:
        f.write(content)
    print(f"  ✅ pace_10_dataset_v2.py → FILE_AOUT mis à jour")
    
    # Update pace_11 (forecast S3) - uses AOUT_SRC
    pace_11 = os.path.join(SCRIPTS_DIR, "pace_11_forecast_S3.py")
    with open(pace_11, 'r') as f:
        content = f.read()
    content = re.sub(
        r'AOUT_SRC\s*=\s*["\'][^"\']*["\']',
        f'AOUT_SRC = "{aout_file}"',
        content
    )
    with open(pace_11, 'w') as f:
        f.write(content)
    print(f"  ✅ pace_11_forecast_S3.py → AOUT_SRC mis à jour")
    
    # === Step 2: Run pipeline ===
    total_start = time.time()
    
    # Step 2a: Build dataset v2 (with material, no mais)
    if not run_script("pace_10_dataset_v2.py", "Phase PREPARE - Construction dataset v2"):
        print("\n❌ Pipeline interrompu: échec dataset v2")
        sys.exit(1)
    
    # Step 2b: Compute prices (Option A)
    # Note: pace_02 uses the original dataset, we need to adapt
    # For now, prices are embedded in pace_11 via prix_forecast.json
    # Skip if prix_forecast.json already exists
    if not os.path.exists(os.path.join(SCRIPTS_DIR, "prix_forecast.json")):
        if not run_script("pace_02_prix_moyens.py", "Phase PREPARE - Calcul prix moyens"):
            print("\n❌ Pipeline interrompu: échec calcul prix")
            sys.exit(1)
    else:
        print("\n⏭️  Prix déjà calculés (prix_forecast.json existe) - étape ignorée")
    
    # Step 2c: Run AED (optional - generates charts)
    if not run_script("pace_03_aed.py", "Phase ANALYZE - AED et graphiques"):
        print("\n⚠️ AED échoué (non bloquant) - continuation")
    
    # Step 2d: Run Prophet forecast S3
    if not run_script("pace_11_forecast_S3.py", "Phase CONSTRUCT - Forecast Prophet S3"):
        print("\n❌ Pipeline interrompu: échec forecast")
        sys.exit(1)
    
    # Step 2e: Generate Excel
    if not run_script("pace_12_excel_S3.py", "Phase EXECUTE - Excel forecast S3"):
        print("\n❌ Pipeline interrompu: échec Excel")
        sys.exit(1)
    
    # Step 2f: Generate charts
    if not run_script("pace_06_charts.py", "Phase EXECUTE - Graphiques de visualisation"):
        print("\n⚠️ Graphiques échoués (non bloquant) - continuation")
    
    total_elapsed = time.time() - total_start
    
    # === Step 3: Summary ===
    print(f"\n{'='*60}")
    print(f"  ✅ PIPELINE TERMINÉ AVEC SUCCÈS")
    print(f"{'='*60}")
    print(f"  Durée totale: {total_elapsed:.0f}s ({total_elapsed/60:.1f} min)")
    print(f"  Date: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"\n  📁 Livrables générés:")
    
    deliverables = [
        ("Dataset consolidé v2", f"{SCRIPTS_DIR}/dataset_consolide_v2.csv"),
        ("Forecast S3 (CSV)", f"{SCRIPTS_DIR}/forecast_q4_2026_S3.csv"),
        ("Excel forecast S3", f"{DOWNLOAD_DIR}/forecast_q4_2026_S3_volumes_valeurs.xlsx"),
        ("Graphiques", f"{DOWNLOAD_DIR}/forecast_q4_2026/charts/"),
    ]
    
    for name, path in deliverables:
        if os.path.exists(path):
            if os.path.isfile(path):
                size = os.path.getsize(path) / 1024
                print(f"    ✅ {name}: {path} ({size:.0f} KB)")
            else:
                n_files = len(os.listdir(path))
                print(f"    ✅ {name}: {path} ({n_files} fichiers)")
        else:
            print(f"    ❌ {name}: MANQUANT")
    
    print(f"\n  💡 Pour mettre à jour les PDFs du projet PACE:")
    print(f"     python scripts/pace_07_pdf_batch1.py")
    print(f"     python scripts/pace_08_pdf_batch2.py")
    print(f"     python scripts/pace_09_guide_methodo.py")
    print(f"\n  💡 Pour mettre à jour les analyses mensuelles:")
    print(f"     python scripts/update_soja_aout.py")
    print(f"     python scripts/compute_aout_mtd_metrics.py")
    print(f"     python scripts/build_zero_achat_pdf.py")
    print(f"     python scripts/build_bundle_analysis_pdf.py")

if __name__ == '__main__':
    main()
