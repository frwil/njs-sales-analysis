# -*- coding: utf-8 -*-
"""Vérification qualité données : conversions tonnes + valeurs aberrantes."""
import pandas as pd
import numpy as np

pd.set_option('display.width', 200)
print('=' * 80)
print('1) CONVERSIONS TONNES — dataset_2023_2026.csv')
print('=' * 80)
d = pd.read_csv('/home/z/my-project/scripts/dataset_2023_2026.csv', low_memory=False)
d['qte'] = pd.to_numeric(d['qte'], errors='coerce')
d['weight_kg'] = pd.to_numeric(d['weight_kg'], errors='coerce')
d['kg'] = pd.to_numeric(d['kg'], errors='coerce')
d['tonnes'] = pd.to_numeric(d['tonnes'], errors='coerce')
d['sacs_50'] = pd.to_numeric(d['sacs_50'], errors='coerce')
d['montant_ttc'] = pd.to_numeric(d['montant_ttc'], errors='coerce')
d['montant_ht'] = pd.to_numeric(d['montant_ht'], errors='coerce')

# kg = qte * weight_kg ?
ok_kg = (d['kg'] - d['qte'] * d['weight_kg']).abs()
print(f'kg = qte*weight_kg : {len(d) - (ok_kg > 0.5).sum():,}/{len(d):,} lignes OK ({100*(len(d)-(ok_kg>0.5).sum())/len(d):.2f}%)')
bad = d[ok_kg > 0.5]
if len(bad):
    print('  Lignes kg incohérentes (échantillon):')
    print(bad[['date', 'ref', 'qte', 'weight_kg', 'kg']].head(5).to_string(index=False))

# tonnes = kg / 1000 ?
ok_t = (d['tonnes'] - d['kg'] / 1000).abs()
print(f'tonnes = kg/1000 : {len(d) - (ok_t > 0.01).sum():,}/{len(d):,} lignes OK')
bad = d[ok_t > 0.01]
if len(bad):
    print('  Lignes tonnes incohérentes (échantillon):')
    print(bad[['date', 'ref', 'kg', 'tonnes']].head(5).to_string(index=False))

# refs 50 kg : sacs_50 doit = qte
d50 = d[d['weight_kg'] == 50]
ok_s = (d50['sacs_50'] - d50['qte']).abs()
print(f'refs 50kg : sacs_50 = qte : {len(d50) - (ok_s > 0.01).sum():,}/{len(d50):,} lignes OK')
bad = d50[ok_s > 0.01]
if len(bad):
    print(bad[['date', 'ref', 'qte', 'sacs_50']].head(5).to_string(index=False))

print()
print('=' * 80)
print('2) CONVERSIONS TONNES — CSV forecasts Q4 et 2027')
print('=' * 80)
for f in ['forecast_q4_2026_S3.csv', 'forecast_2027_S3.csv']:
    q = pd.read_csv(f'/home/z/my-project/scripts/{f}')
    fam50 = q[q['family'].isin(['TOURTEAUX', 'CONCENTRES'])]
    ecart = (q['tonnes'] - q['sacs_50'] * 50 / 1000).abs()
    print(f'{f}: tonnes = sacs*50/1000 : {len(q) - (ecart > 0.01).sum():,}/{len(q):,} lignes OK; '
          f'écart max {ecart.max():.2f} t')
    # autres familles : prix/unité plausibilité
    autres = q[~q['family'].isin(['TOURTEAUX', 'CONCENTRES'])]
    p = autres[autres['sacs_50'] > 0].copy()
    p['prix_unit'] = p['ca_m_fcfa'] * 1e6 / p['sacs_50']
    print(f'  autres familles prix/unité : min {p["prix_unit"].min():.0f}, max {p["prix_unit"].max():.0f}, '
          f'médiane {p["prix_unit"].median():.0f}')

print()
print('=' * 80)
print('3) VALEURS ABERRANTES — prix moyen par sac (dataset, par famille x mois)')
print('=' * 80)
d2 = d[(d['sacs_50'] > 0) & (d['family'] == 'TOURTEAUX')].copy()
d2['prix_sac'] = d2['montant_ttc'] / d2['sacs_50']
d2['ym'] = d2['date'].astype(str).str[:7]
px = d2.groupby('ym').agg(sacs=('sacs_50', 'sum'), prix=('prix_sac', 'mean'),
                          prix_min=('prix_sac', 'min'), prix_max=('prix_sac', 'max'))
px = px[(px['sacs'] > 100)]
print('Soja prix/sac moyen par mois (derniers 15 mois):')
print(px.tail(15).round(0).to_string())

# lignes aberrantes : prix/sac < 5000 ou > 60000
ab = d2[(d2['prix_sac'] < 5000) | (d2['prix_sac'] > 60000)]
print(f'\nSoja lignes prix/sac hors [5 000, 60 000] : {len(ab):,} / {len(d2):,}')
if len(ab):
    print(ab[['date', 'ref', 'client', 'qte', 'montant_ttc', 'prix_sac']].head(10).to_string(index=False))

print()
print('=' * 80)
print('4) VALEURS ABERRANTES — autres familles, prix/unité dataset')
print('=' * 80)
d3 = d[(d['sacs_50'] > 0)].copy()
d3['prix_unit'] = d3['montant_ttc'] / d3['sacs_50']
for fam in sorted(d3['family'].unique()):
    sub = d3[d3['family'] == fam]
    print(f'{fam:25s} lignes {len(sub):>7,} | prix/unité médian {sub["prix_unit"].median():>9,.0f} | '
          f'min {sub["prix_unit"].min():>9,.0f} | max {sub["prix_unit"].max():>9,.0f}')

print()
print('=' * 80)
print('5) VALEURS ABERRANTES — montants par ligne et panier moyen')
print('=' * 80)
d['ligne_ca'] = d['montant_ttc']
big = d[d['ligne_ca'] > 100e6]
print(f'Lignes CA > 100 M FCFA : {len(big)}')
if len(big):
    print(big[['date', 'ref', 'client', 'qte', 'montant_ttc']].head(10).to_string(index=False))
neg = d[(d['qte'] <= 0) | (d['montant_ttc'] < 0)]
print(f'Lignes qte<=0 ou CA<0 : {len(neg)}')
if len(neg):
    print(neg[['date', 'ref', 'client', 'qte', 'montant_ttc']].head(10).to_string(index=False))

# panier moyen par commande (ERP septembre)
print()
print('=' * 80)
print('6) PANIER MOYEN — ERP septembre (Livrée)')
print('=' * 80)
from openpyxl import load_workbook
wb = load_workbook('/home/z/my-project/upload/NJS GROUP ERP - Lignes de commandes + multicompany (51).xlsx', read_only=True)
ws = wb.active
cmds = {}
for r in ws.iter_rows(min_row=3, values_only=True):
    if not r or not r[0] or r[0] == 'Total':
        continue
    etat = str(r[13]).strip() if r[13] else ''
    if etat != 'Livrée':
        continue
    ds = str(r[6])[:10] if r[6] else ''
    if '/09/2026' not in ds:
        continue
    c = cmds.setdefault(r[3], {'qte': 0, 'ttc': 0.0, 'client': r[5]})
    c['qte'] += r[2] or 0
    c['ttc'] += r[9] or 0
dfc = pd.DataFrame(cmds.values())
dfc['panier'] = dfc['ttc'] / np.maximum(dfc['qte'].map(lambda x: max(x, 1)), 1)
print(f'Commandes Livrée sept : {len(dfc):,}')
print(f'Panier moyen/commande : {dfc["ttc"].mean():,.0f} FCFA | médian {dfc["ttc"].median():,.0f} | max {dfc["ttc"].max():,.0f}')
print(f'Quantité moyenne/commande : {dfc["qte"].mean():.1f} | max {dfc["qte"].max():.0f}')
bigc = dfc[dfc['ttc'] > 50e6]
print(f'Commandes > 50 M FCFA : {len(bigc)}')
if len(bigc):
    print(bigc.head(10).to_string(index=False))
top = dfc.nlargest(5, 'ttc')
print('Top 5 commandes (CA):')
print(top.to_string(index=False))
print()
print('VERIFICATION TERMINEE')
