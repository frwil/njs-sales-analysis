"""
Génère les livrables mensuels de performance pour une plage de mois (défaut : janvier → septembre 2026).

Chaque livrable est autonome : le YTD de chaque fichier est arrêté au mois du fichier
(ex. fichier de mars = YTD Jan → Mars ; fichier de janvier = YTD janvier seul).

Pour chaque mois : compute_monthly_performance.py <mois> (JSON + graphiques)
puis build_monthly_performance_pdf.py <mois> (PDF) — dans cet ordre, car les
graphiques partagés sont écrasés à chaque mois et doivent être embarqués aussitôt.

Usage :
  python generate_all_monthly_livrables.py [mois_debut [mois_fin]]
"""
import os
import subprocess
import sys

BASE = "/home/z/my-project"
START = int(sys.argv[1]) if len(sys.argv) > 1 else 1
END = int(sys.argv[2]) if len(sys.argv) > 2 else 9

env = dict(os.environ, PYTHONIOENCODING='utf-8')
failed = []

for m in range(START, END + 1):
    print("\n" + "=" * 70)
    print(f"  MOIS {m:02d}/2026 — calcul + PDF")
    print("=" * 70)
    for script in ('scripts/compute_monthly_performance.py',
                   'scripts/build_monthly_performance_pdf.py'):
        r = subprocess.run([sys.executable, os.path.join(BASE, script), str(m)],
                           env=env, cwd=BASE)
        if r.returncode != 0:
            print(f"!!! ERREUR mois {m:02d} sur {script}")
            failed.append((m, script))

if failed:
    print("\nECHECS :", failed)
    sys.exit(1)
print("\nTous les livrables mensuels sont générés.")
