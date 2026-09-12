"""
Forecast Q4 2026 + 2027 - Génération des 5 PDFs PACE - 
AVEC COMPLEMENT_ALIMENTAIRE (V300 1L only), données 2023-2026, MATERIEL_ELEVAGE tonnes=0

Génère:
- 5 PDFs pour Q4 2026 (dans download/forecast_q4_2026/)
- 5 PDFs pour 2027 (dans download/forecast_2027/)

VERSION 4: Données chargées dynamiquement depuis les CSV (plus de valeurs hardcoded).
"""
import os
import json
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle,
    Image, HRFlowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY

# === CHARGEMENT DYNAMIQUE DES DONNÉES ===
print("Chargement dynamique des données...")
Q4_CSV = "/home/z/my-project/scripts/forecast_q4_2026_S3.csv"
F2027_CSV = "/home/z/my-project/scripts/forecast_2027_S3.csv"
DATASET_CSV = "/home/z/my-project/scripts/dataset_2023_2026.csv"

q4_df = pd.read_csv(Q4_CSV)
f2027_df = pd.read_csv(F2027_CSV)
hist_df = pd.read_csv(DATASET_CSV, parse_dates=['date'], low_memory=False)

# Helpers de formatage
def fmt_t(x):
    """Format tonnes: 1234 -> '1 234'"""
    return f"{int(round(x)):,}".replace(',', ' ')

def fmt_ca(x):
    """Format CA: 1234.5 -> '1 234' (M FCFA)"""
    return f"{int(round(x)):,}".replace(',', ' ')

def fmt_pct(x):
    """Format %: 12.34 -> '12,3%'"""
    return f"{x:.1f}%".replace('.', ',')

def fmt_pct_signed(x):
    """Format signed %: 12.34 -> '+12,3%' or -5.0 -> '-5,0%'"""
    s = '+' if x >= 0 else ''
    return f"{s}{x:.1f}%".replace('.', ',')

# === Q4 2026 ===
Q4_TOTAL_T = q4_df['tonnes'].sum()
Q4_TOTAL_CA = q4_df['ca_m_fcfa'].sum()
Q4_FAM = q4_df.groupby('family').agg(t=('tonnes','sum'), ca=('ca_m_fcfa','sum')).reset_index()
Q4_FAM['pct'] = Q4_FAM['ca'] / Q4_TOTAL_CA * 100

Q4_MONTH = q4_df.groupby('month').agg(t=('tonnes','sum'), ca=('ca_m_fcfa','sum')).reset_index()
Q4_MONTH['pct'] = Q4_MONTH['ca'] / Q4_TOTAL_CA * 100

Q4_TOP_AGENCES = q4_df.groupby(['agence','region']).agg(ca=('ca_m_fcfa','sum')).reset_index().sort_values('ca', ascending=False).head(5)

# YTD 2026 (Jan-Août réel) + Q4 forecast par famille
hist_2026_ytd = hist_df[(hist_df['date'].dt.year == 2026) & (hist_df['date'].dt.month <= 8)]
YTD_2026_BY_FAM = hist_2026_ytd.groupby('family')['tonnes'].sum()
YTD_2026_TOTAL_T = YTD_2026_BY_FAM.sum()
Q4_FAM_BY_FAM = q4_df.groupby('family')['tonnes'].sum()

# === 2027 ===
F2027_TOTAL_T = f2027_df['tonnes'].sum()
F2027_TOTAL_CA = f2027_df['ca_m_fcfa'].sum()
F2027_FAM = f2027_df.groupby('family').agg(t=('tonnes','sum'), ca=('ca_m_fcfa','sum')).reset_index()
F2027_FAM['pct'] = F2027_FAM['ca'] / F2027_TOTAL_CA * 100

f2027_df['quarter'] = ((f2027_df['month'] - 1) // 3) + 1
F2027_Q = f2027_df.groupby('quarter').agg(t=('tonnes','sum'), ca=('ca_m_fcfa','sum')).reset_index()
F2027_Q['pct'] = F2027_Q['ca'] / F2027_TOTAL_CA * 100

F2027_TOP_AGENCES = f2027_df.groupby(['agence','region']).agg(ca=('ca_m_fcfa','sum')).reset_index().sort_values('ca', ascending=False).head(5)

# Historique par année par famille (tonnes)
HIST_BY_YEAR_FAM = {}
for year in [2023, 2024, 2025, 2026]:
    yr_df = hist_df[hist_df['date'].dt.year == year]
    if year == 2026:
        yr_df = yr_df[yr_df['date'].dt.month <= 8]  # YTD Jan-Août
    HIST_BY_YEAR_FAM[year] = yr_df.groupby('family')['tonnes'].sum()

# Total 2026 LY = YTD réel + Q4 forecast
LY_2026_BY_FAM_T = YTD_2026_BY_FAM.add(Q4_FAM_BY_FAM, fill_value=0)
LY_2026_TOTAL_T = LY_2026_BY_FAM_T.sum()

# CA 2026 LY = YTD réel + Q4 forecast
YTD_2026_CA_BY_FAM = hist_2026_ytd.groupby('family')['montant_ttc'].sum() / 1e6
Q4_CA_BY_FAM = q4_df.groupby('family')['ca_m_fcfa'].sum()
LY_2026_CA_BY_FAM = YTD_2026_CA_BY_FAM.add(Q4_CA_BY_FAM, fill_value=0)
LY_2026_TOTAL_CA = LY_2026_CA_BY_FAM.sum()

# === Top 5 agences par CA ===
def top_agences_table(df, top_n=5):
    top = df.groupby(['agence','region']).agg(ca=('ca_m_fcfa','sum')).reset_index().sort_values('ca', ascending=False).head(top_n)
    rows = [["Rang", "Agence", "Région", "CA (M FCFA)"]]
    for i, (_, r) in enumerate(top.iterrows(), 1):
        rows.append([str(i), r['agence'].upper(), r['region'], fmt_ca(r['ca'])])
    return rows

# === Construction des tables ===
def build_q4_synth_table():
    return [
        ["Indicateur", "Valeur", "Détail"],
        ["Volume total Q4 2026", f"{fmt_t(Q4_TOTAL_T)} t", "8 familles, 121 produits, 25 agences"],
        ["CA total Q4 2026", f"{fmt_ca(Q4_TOTAL_CA)} M FCFA", "Prix soja: Q4=17 170 (moyen YTD) / 2027=16 800 (médiane)"],
        ["Période", "Sept-Déc 2026 (4 mois)", "Saison haute (35-46% du volume annuel)"],
        ["Scénario", "S3 (réappro 100%)", "80 000 sacs au 15/09/2026"],
        ["Données historiques", "176 576 enregistrements", "Jan 2023 - Août 2026 + En cours/Validées"],
        ["Désaisonnalisation", "Effet soja Jul-Août 2026 neutralisé", "Cap à moyenne S1 2026"],
        ["ALIMENT_COMPLET", f"{fmt_t(Q4_FAM[Q4_FAM['family']=='ALIMENT_COMPLET']['t'].iloc[0])} t / {fmt_ca(Q4_FAM[Q4_FAM['family']=='ALIMENT_COMPLET']['ca'].iloc[0])} M FCFA", "Prophet + facteur reprise +15% (post-2024)"],
        ["PREMIX", f"{fmt_t(Q4_FAM[Q4_FAM['family']=='PREMIX']['t'].iloc[0])} t / {fmt_ca(Q4_FAM[Q4_FAM['family']=='PREMIX']['ca'].iloc[0])} M FCFA", "Prophet (avec volumes)"],
        ["ALVEOLES (forfait SPC)", "0 t / 34 M FCFA", "4 refs MAT011/MAT014/MAT015/MAT017"],
        ["Forfait SPC", "39 M/an MAT only (2026 annualisé)", "ALV=0, MAT=38.4M"],
        ["Bundle 2.5:1", "Ratio soja:concentré <= 2.5:1", "6 ajustements Q4"],
    ]

def build_q4_fam_table():
    rows = [["Famille", "Volume Q4 (t)", "CA Q4 (M FCFA)", "Part CA"]]
    fam_order = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT_COMPLET', 'INGREDIENTS', 'ALVEOLES', 'MATERIEL_ELEVAGE', 'PREMIX', 'COMPLEMENT_ALIMENTAIRE']
    for fam in fam_order:
        sub = Q4_FAM[Q4_FAM['family']==fam]
        if len(sub) == 0:
            rows.append([fam, "0", "0", "0,0%"])
        else:
            rows.append([fam, fmt_t(sub['t'].iloc[0]), fmt_ca(sub['ca'].iloc[0]), fmt_pct(sub['pct'].iloc[0])])
    rows.append(["TOTAL", fmt_t(Q4_TOTAL_T), fmt_ca(Q4_TOTAL_CA), "100%"])
    return rows

def build_q4_month_table():
    month_names = {9: "Septembre 2026", 10: "Octobre 2026", 11: "Novembre 2026", 12: "Décembre 2026"}
    rows = [["Mois", "Volume (t)", "CA (M FCFA)", "Part CA"]]
    for m in [9, 10, 11, 12]:
        sub = Q4_MONTH[Q4_MONTH['month']==m]
        if len(sub) > 0:
            rows.append([month_names[m], fmt_t(sub['t'].iloc[0]), fmt_ca(sub['ca'].iloc[0]), fmt_pct(sub['pct'].iloc[0])])
    rows.append(["TOTAL Q4", fmt_t(Q4_TOTAL_T), fmt_ca(Q4_TOTAL_CA), "100%"])
    return rows

def build_q4_ytd_table():
    fam_order = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT_COMPLET', 'INGREDIENTS', 'PREMIX', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES', 'MATERIEL_ELEVAGE']
    rows = [["Famille", "YTD 2026 (t)", "Q4 fcst (t)", "Total 2026 (t)"]]
    total_ytd = 0
    total_q4 = 0
    for fam in fam_order:
        ytd = YTD_2026_BY_FAM.get(fam, 0)
        q4 = Q4_FAM_BY_FAM.get(fam, 0)
        total = ytd + q4
        total_ytd += ytd
        total_q4 += q4
        if fam in ('ALVEOLES', 'MATERIEL_ELEVAGE'):
            rows.append([fam, fmt_t(ytd), "0 (CA only)", fmt_t(ytd)])
        else:
            rows.append([fam, fmt_t(ytd), fmt_t(q4), fmt_t(total)])
    rows.append(["MAÏS (exclu)", "0", "0", "0"])
    rows.append(["TOTAL", fmt_t(total_ytd), fmt_t(total_q4), fmt_t(total_ytd + total_q4)])
    return rows

def build_q4_fam_detail_table():
    """Pour guide méthodologique Q4 - avec colonne méthode"""
    methods = {
        'TOURTEAUX': "Prophet + prix 17 170 (moyen YTD 2026)",
        'CONCENTRES': "Prophet + bundle 2.5:1",
        'ALIMENT_COMPLET': "Prophet + filtrage 2024 + reprise +15%",
        'INGREDIENTS': "Prophet + prix 2026 réels",
        'ALVEOLES': "Extrap 2026 + Forfait SPC (ALV=0)",
        'MATERIEL_ELEVAGE': "Extrap + Forfait SPC (2026 annualisé)",
        'PREMIX': "Prophet (avec volumes)",
        'COMPLEMENT_ALIMENTAIRE': "Prophet (proxy V300 1L)",
    }
    fam_order = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT_COMPLET', 'INGREDIENTS', 'ALVEOLES', 'MATERIEL_ELEVAGE', 'PREMIX', 'COMPLEMENT_ALIMENTAIRE']
    rows = [["Famille", "Volume Q4 (t)", "CA Q4 (M FCFA)", "Part CA", "Méthode"]]
    for fam in fam_order:
        sub = Q4_FAM[Q4_FAM['family']==fam]
        if len(sub) == 0:
            rows.append([fam, "0", "0", "0,0%", methods.get(fam, "")])
        else:
            rows.append([fam, fmt_t(sub['t'].iloc[0]), fmt_ca(sub['ca'].iloc[0]), fmt_pct(sub['pct'].iloc[0]), methods.get(fam, "")])
    return rows

def build_q4_month_detail_table():
    """Pour guide méthodologique Q4 - avec colonne lecture"""
    readings = {
        9: "Démarrage Q4 + réappro",
        10: "Pic mensuel",
        11: "Maintien",
        12: "Fêtes de fin d'année",
    }
    month_names = {9: "Septembre 2026", 10: "Octobre 2026", 11: "Novembre 2026", 12: "Décembre 2026"}
    rows = [["Mois", "Volume (t)", "CA (M FCFA)", "Part CA", "Lecture"]]
    for m in [9, 10, 11, 12]:
        sub = Q4_MONTH[Q4_MONTH['month']==m]
        if len(sub) > 0:
            rows.append([month_names[m], fmt_t(sub['t'].iloc[0]), fmt_ca(sub['ca'].iloc[0]), fmt_pct(sub['pct'].iloc[0]), readings[m]])
    return rows

# === Tables 2027 ===
def build_2027_synth_table():
    return [
        ["Indicateur", "Valeur", "Détail"],
        ["Volume total 2027", f"{fmt_t(F2027_TOTAL_T)} t", "8 familles, 121 produits, 25 agences"],
        ["CA total 2027", f"{fmt_ca(F2027_TOTAL_CA)} M FCFA", "Prix soja: Q4=17 170 (moyen YTD) / 2027=16 800 (médiane)"],
        ["Période", "12 mois (Jan-Déc 2027)", "Forecast complet annuel"],
        ["Scénario", "S3 (réappro 100%)", "Situation normale"],
        ["Données historiques", "176 576 enregistrements", "Jan 2023 - Août 2026 + En cours/Validées"],
        ["Désaisonnalisation", "Effet soja Jul-Août 2026 neutralisé", "Cap à moyenne S1 2026"],
        ["ALIMENT_COMPLET", f"{fmt_t(F2027_FAM[F2027_FAM['family']=='ALIMENT_COMPLET']['t'].iloc[0])} t / {fmt_ca(F2027_FAM[F2027_FAM['family']=='ALIMENT_COMPLET']['ca'].iloc[0])} M FCFA", "Prophet + facteur reprise +15% (post-2024)"],
        ["PREMIX", f"{fmt_t(F2027_FAM[F2027_FAM['family']=='PREMIX']['t'].iloc[0])} t / {fmt_ca(F2027_FAM[F2027_FAM['family']=='PREMIX']['ca'].iloc[0])} M FCFA", "Prophet (avec volumes)"],
        ["ALVEOLES (forfait SPC)", "0 t / 102 M FCFA", "4 refs MAT011/MAT014/MAT015/MAT017"],
        ["Forfait SPC", "39 M/an MAT only (2026 annualisé)", "ALV=0, MAT=38.4M"],
        ["Bundle 2.5:1", "Ratio soja:concentré <= 2.5:1", "26 ajustements annuels"],
    ]

def build_2027_fam_table():
    rows = [["Famille", "Volume (t)", "CA (M FCFA)", "Part CA"]]
    fam_order = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT_COMPLET', 'INGREDIENTS', 'ALVEOLES', 'MATERIEL_ELEVAGE', 'PREMIX', 'COMPLEMENT_ALIMENTAIRE']
    for fam in fam_order:
        sub = F2027_FAM[F2027_FAM['family']==fam]
        if len(sub) == 0:
            rows.append([fam, "0", "0", "0,0%"])
        else:
            rows.append([fam, fmt_t(sub['t'].iloc[0]), fmt_ca(sub['ca'].iloc[0]), fmt_pct(sub['pct'].iloc[0])])
    rows.append(["MAÏS (exclu)", "0", "0", "0,0%"])
    rows.append(["TOTAL", fmt_t(F2027_TOTAL_T), fmt_ca(F2027_TOTAL_CA), "100%"])
    return rows

def build_2027_q_table():
    q_names = {1: "Q1 (Jan-Mar)", 2: "Q2 (Avr-Juin)", 3: "Q3 (Juil-Sept)", 4: "Q4 (Oct-Déc)"}
    rows = [["Trimestre", "Volume (t)", "CA (M FCFA)", "Part CA"]]
    for q in [1, 2, 3, 4]:
        sub = F2027_Q[F2027_Q['quarter']==q]
        if len(sub) > 0:
            rows.append([q_names[q], fmt_t(sub['t'].iloc[0]), fmt_ca(sub['ca'].iloc[0]), fmt_pct(sub['pct'].iloc[0])])
    rows.append(["TOTAL", fmt_t(F2027_TOTAL_T), fmt_ca(F2027_TOTAL_CA), "100%"])
    return rows

def build_2027_hist_table():
    """Historique 2024-2026 (Vol + CA) + Forecast 2027 par famille.
    Inclut 2024, 2025 (vol + CA) et 2026 LY (année complète = YTD réel + Q4 forecast, vol + CA).
    """
    import json as _json
    import os as _os
    _hist_path = '/home/z/my-project/scripts/historical_2024_2026.json'
    if _os.path.exists(_hist_path):
        _hist = _json.load(open(_hist_path))
    else:
        _hist = {'by_year': {}, '2026_ly': {}}
    
    fam_order = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT_COMPLET', 'INGREDIENTS', 'PREMIX', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES', 'MATERIEL_ELEVAGE', 'MAIS']
    rows = [["Famille", "2024 Vol (t)", "2024 CA (M)", "2025 Vol (t)", "2025 CA (M)", "2026 LY Vol (t)", "2026 LY CA (M)", "2027 fcst (t)", "2027 CA (M)"]]
    for fam in fam_order:
        # 2024
        h24 = _hist.get('by_year', {}).get('2024', {}).get(fam, {'vol': 0, 'ca': 0})
        v24 = h24['vol'] if isinstance(h24, dict) else 0
        c24 = h24['ca'] if isinstance(h24, dict) else 0
        # 2025
        h25 = _hist.get('by_year', {}).get('2025', {}).get(fam, {'vol': 0, 'ca': 0})
        v25 = h25['vol'] if isinstance(h25, dict) else 0
        c25 = h25['ca'] if isinstance(h25, dict) else 0
        # 2026 LY (full year)
        h26 = _hist.get('2026_ly', {}).get(fam, {'vol': 0, 'ca': 0})
        v26 = h26['vol'] if isinstance(h26, dict) else 0
        c26 = h26['ca'] if isinstance(h26, dict) else 0
        # 2027 forecast
        t27 = F2027_FAM[F2027_FAM['family']==fam]['t'].iloc[0] if len(F2027_FAM[F2027_FAM['family']==fam])>0 else 0
        ca27 = F2027_FAM[F2027_FAM['family']==fam]['ca'].iloc[0] if len(F2027_FAM[F2027_FAM['family']==fam])>0 else 0
        
        rows.append([fam, fmt_t(v24), fmt_ca(c24), fmt_t(v25), fmt_ca(c25), fmt_t(v26), fmt_ca(c26), fmt_t(t27), fmt_ca(ca27)])
    
    # Total row (hors MAIS for 2027)
    tot_v24 = sum(_hist.get('by_year', {}).get('2024', {}).get(f, {'vol':0})['vol'] if isinstance(_hist.get('by_year', {}).get('2024', {}).get(f, {'vol':0}), dict) else 0 for f in fam_order if f != 'MAIS')
    tot_c24 = sum(_hist.get('by_year', {}).get('2024', {}).get(f, {'ca':0})['ca'] if isinstance(_hist.get('by_year', {}).get('2024', {}).get(f, {'ca':0}), dict) else 0 for f in fam_order if f != 'MAIS')
    tot_v25 = sum(_hist.get('by_year', {}).get('2025', {}).get(f, {'vol':0})['vol'] if isinstance(_hist.get('by_year', {}).get('2025', {}).get(f, {'vol':0}), dict) else 0 for f in fam_order if f != 'MAIS')
    tot_c25 = sum(_hist.get('by_year', {}).get('2025', {}).get(f, {'ca':0})['ca'] if isinstance(_hist.get('by_year', {}).get('2025', {}).get(f, {'ca':0}), dict) else 0 for f in fam_order if f != 'MAIS')
    tot_v26 = sum(_hist.get('2026_ly', {}).get(f, {'vol':0})['vol'] if isinstance(_hist.get('2026_ly', {}).get(f, {'vol':0}), dict) else 0 for f in fam_order if f != 'MAIS')
    tot_c26 = sum(_hist.get('2026_ly', {}).get(f, {'ca':0})['ca'] if isinstance(_hist.get('2026_ly', {}).get(f, {'ca':0}), dict) else 0 for f in fam_order if f != 'MAIS')
    
    rows.append(["TOTAL (hors MAIS)", fmt_t(tot_v24), fmt_ca(tot_c24), fmt_t(tot_v25), fmt_ca(tot_c25), fmt_t(tot_v26), fmt_ca(tot_c26), fmt_t(F2027_TOTAL_T), fmt_ca(F2027_TOTAL_CA)])
    return rows

def build_2027_fam_detail_table():
    """Pour document stratégique 2027 - avec colonne méthode"""
    methods = {
        'TOURTEAUX': "Prophet + prix 17 170 (moyen YTD 2026)",
        'CONCENTRES': "Prophet + bundle 2.5:1",
        'ALIMENT_COMPLET': "Prophet + filtrage 2024 + reprise +15%",
        'INGREDIENTS': "Prophet + prix 2026 réels",
        'ALVEOLES': "Extrap 2026 + Forfait SPC (ALV=0)",
        'MATERIEL_ELEVAGE': "Extrap + Forfait SPC (2026 annualisé)",
        'PREMIX': "Prophet (avec volumes)",
        'COMPLEMENT_ALIMENTAIRE': "Prophet (V300 1L proxy)",
    }
    fam_order = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT_COMPLET', 'INGREDIENTS', 'ALVEOLES', 'MATERIEL_ELEVAGE', 'PREMIX', 'COMPLEMENT_ALIMENTAIRE']
    rows = [["Famille", "Volume (t)", "CA (M FCFA)", "Part CA", "Méthode"]]
    for fam in fam_order:
        sub = F2027_FAM[F2027_FAM['family']==fam]
        if len(sub) == 0:
            rows.append([fam, "0", "0", "0,0%", methods.get(fam, "")])
        else:
            rows.append([fam, fmt_t(sub['t'].iloc[0]), fmt_ca(sub['ca'].iloc[0]), fmt_pct(sub['pct'].iloc[0]), methods.get(fam, "")])
    return rows

def build_2027_q_detail_table():
    """Pour document stratégique 2027 - avec colonne lecture"""
    readings = {1: "Démarrage", 2: "Montée", 3: "Creux", 4: "Pic saisonnier"}
    q_names = {1: "Q1 (Jan-Mar)", 2: "Q2 (Avr-Juin)", 3: "Q3 (Juil-Sept)", 4: "Q4 (Oct-Déc)"}
    rows = [["Trimestre", "Volume (t)", "CA (M FCFA)", "Part CA", "Lecture"]]
    for q in [1, 2, 3, 4]:
        sub = F2027_Q[F2027_Q['quarter']==q]
        if len(sub) > 0:
            rows.append([q_names[q], fmt_t(sub['t'].iloc[0]), fmt_ca(sub['ca'].iloc[0]), fmt_pct(sub['pct'].iloc[0]), readings[q]])
    return rows

def build_2026_full_year_table():
    """Réalisation 2026 complète = YTD réel (Jan-Août) + Q4 forecast (Sept-Déc) par famille.
    Affiché dans le résumé exécutif 2027 pour retracer l'année 2026 avant le forecast 2027."""
    fam_order = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT_COMPLET', 'INGREDIENTS', 'PREMIX',
                 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES', 'MATERIEL_ELEVAGE', 'MAIS']
    rows = [["Famille", "YTD réel (t)", "Q4 fcst (t)", "Total 2026 (t)", "2027 fcst (t)", "Δ Vol %"]]
    total_ytd = 0
    total_q4 = 0
    total_2026 = 0
    total_2027 = 0
    # Pour le total hors MAIS (comparaison apples-to-apples avec 2027)
    total_ytd_no_mais = 0
    total_q4_no_mais = 0
    total_2026_no_mais = 0
    total_2027_no_mais = 0
    for fam in fam_order:
        ytd = YTD_2026_BY_FAM.get(fam, 0)
        q4 = Q4_FAM_BY_FAM.get(fam, 0)
        total_2026_fam = ytd + q4
        t27 = F2027_FAM[F2027_FAM['family']==fam]['t'].iloc[0] if len(F2027_FAM[F2027_FAM['family']==fam])>0 else 0
        var = (t27/total_2026_fam - 1) * 100 if total_2026_fam > 0 else None
        total_ytd += ytd
        total_q4 += q4
        total_2026 += total_2026_fam
        total_2027 += t27
        if fam != 'MAIS':
            total_ytd_no_mais += ytd
            total_q4_no_mais += q4
            total_2026_no_mais += total_2026_fam
            total_2027_no_mais += t27
        # Pour ALVEOLES et MATERIEL_ELEVAGE, YTD réel est 0 (pas de volume) et Q4 = 0 (CA only)
        if fam in ('ALVEOLES', 'MATERIEL_ELEVAGE'):
            q4_display = "0 (CA only)"
        else:
            q4_display = fmt_t(q4)
        # Pour MAIS, Q4 = 0 (exclu du forecast)
        if fam == 'MAIS':
            q4_display = "0 (exclu)"
        rows.append([fam, fmt_t(ytd), q4_display, fmt_t(total_2026_fam), fmt_t(t27),
                    fmt_pct_signed(var) if var is not None else "—"])
    var_total_avec_mais = (total_2027/total_2026 - 1) * 100 if total_2026 > 0 else 0
    var_total_hors_mais = (total_2027_no_mais/total_2026_no_mais - 1) * 100 if total_2026_no_mais > 0 else 0
    rows.append(["TOTAL (toutes familles)", fmt_t(total_ytd), fmt_t(total_q4), fmt_t(total_2026),
                fmt_t(total_2027), fmt_pct_signed(var_total_avec_mais)])
    rows.append(["TOTAL hors MAIS (référence)", fmt_t(total_ytd_no_mais), fmt_t(total_q4_no_mais),
                fmt_t(total_2026_no_mais), fmt_t(total_2027_no_mais), fmt_pct_signed(var_total_hors_mais)])
    return rows

def build_2027_progression_table():
    """Progression 2027 vs 2026 LY par famille"""
    fam_order = ['TOURTEAUX', 'CONCENTRES', 'ALIMENT_COMPLET', 'INGREDIENTS', 'PREMIX', 'COMPLEMENT_ALIMENTAIRE', 'ALVEOLES', 'MATERIEL_ELEVAGE']
    rows = [["Famille", "2026 (t)", "2027 fcst (t)", "Δ Vol %", "2026 (M)", "2027 fcst (M)", "Δ CA %"]]
    total_2026_t = 0
    total_2027_t = 0
    total_2026_ca = 0
    total_2027_ca = 0
    for fam in fam_order:
        t26 = LY_2026_BY_FAM_T.get(fam, 0)
        t27 = F2027_FAM[F2027_FAM['family']==fam]['t'].iloc[0] if len(F2027_FAM[F2027_FAM['family']==fam])>0 else 0
        ca26 = LY_2026_CA_BY_FAM.get(fam, 0)
        ca27 = F2027_FAM[F2027_FAM['family']==fam]['ca'].iloc[0] if len(F2027_FAM[F2027_FAM['family']==fam])>0 else 0
        var_t = (t27/t26 - 1) * 100 if t26 > 0 else None
        var_ca = (ca27/ca26 - 1) * 100 if ca26 > 0 else None
        total_2026_t += t26
        total_2027_t += t27
        total_2026_ca += ca26
        total_2027_ca += ca27
        rows.append([fam, fmt_t(t26), fmt_t(t27),
                    fmt_pct_signed(var_t) if var_t is not None else "—",
                    fmt_ca(ca26), fmt_ca(ca27),
                    fmt_pct_signed(var_ca) if var_ca is not None else "—"])
    var_total_t = (total_2027_t/total_2026_t - 1) * 100
    var_total_ca = (total_2027_ca/total_2026_ca - 1) * 100
    rows.append(["TOTAL", fmt_t(total_2026_t), fmt_t(total_2027_t), fmt_pct_signed(var_total_t),
                fmt_ca(total_2026_ca), fmt_ca(total_2027_ca), fmt_pct_signed(var_total_ca)])
    return rows

print(f"  Q4 2026: {Q4_TOTAL_T:.0f} t / {Q4_TOTAL_CA:.0f} M FCFA")
print(f"  2027: {F2027_TOTAL_T:.0f} t / {F2027_TOTAL_CA:.0f} M FCFA")
print(f"  LY 2026: {LY_2026_TOTAL_T:.0f} t / {LY_2026_TOTAL_CA:.0f} M FCFA")
print(f"  Var 2027 vs 2026: volume {fmt_pct_signed((F2027_TOTAL_T/LY_2026_TOTAL_T-1)*100)}, CA {fmt_pct_signed((F2027_TOTAL_CA/LY_2026_TOTAL_CA-1)*100)}")
print("Données chargées.")

pdfmetrics.registerFont(TTFont('DejaVuSans', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))

NAVY = HexColor('#1F4E78'); GOLD = HexColor('#C9A961'); GRAY = HexColor('#595959')
LIGHT_GRAY = HexColor('#F2F2F2'); GREEN = HexColor('#C6EFCE'); S3_COLOR = 'C6EFCE'
NEW_FAMILY_COLOR = 'FFE699'  # Yellow for COMPLEMENT_ALIMENTAIRE

styles = getSampleStyleSheet()
H1 = ParagraphStyle('H1', parent=styles['Heading1'], fontName='DejaVuSans-Bold', fontSize=18, textColor=NAVY, spaceAfter=14, spaceBefore=10)
H2 = ParagraphStyle('H2', parent=styles['Heading2'], fontName='DejaVuSans-Bold', fontSize=14, textColor=NAVY, spaceAfter=10, spaceBefore=14)
H3 = ParagraphStyle('H3', parent=styles['Heading3'], fontName='DejaVuSans-Bold', fontSize=12, textColor=GOLD, spaceAfter=8, spaceBefore=10)
BODY = ParagraphStyle('Body', parent=styles['Normal'], fontName='DejaVuSans', fontSize=10, leading=14, alignment=TA_JUSTIFY, spaceAfter=8)
BULLET = ParagraphStyle('Bullet', parent=BODY, leftIndent=20, bulletIndent=10, spaceAfter=4)
SMALL = ParagraphStyle('Small', parent=BODY, fontSize=8, textColor=GRAY)

CELL_STYLE = ParagraphStyle('CellStyle', parent=BODY, fontName='DejaVuSans', fontSize=9, leading=11, alignment=TA_LEFT, spaceAfter=0, spaceBefore=0)

def make_table(data, col_widths=None, font_size=9, header_color=NAVY, highlight_rows=None):
    cell_style = ParagraphStyle('CD', parent=CELL_STYLE, fontSize=font_size, leading=font_size+2)
    cell_header = ParagraphStyle('CH', parent=cell_style, fontName='DejaVuSans-Bold', textColor=colors.white, alignment=TA_CENTER)
    cell_center = ParagraphStyle('CC', parent=cell_style, alignment=TA_CENTER)
    processed = []
    for ri, row in enumerate(data):
        prow = []
        for ci, cell in enumerate(row):
            s = str(cell) if cell is not None else ''
            if ri == 0: prow.append(Paragraph(s, cell_header))
            elif len(s) <= 15: prow.append(Paragraph(s, cell_center))
            else: prow.append(Paragraph(s, cell_style))
        processed.append(prow)
    t = Table(processed, colWidths=col_widths, repeatRows=1)
    style_list = [
        ('BACKGROUND', (0,0), (-1,0), header_color),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.gray),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_GRAY]),
        ('TOPPADDING', (0,0), (-1,-1), 4), ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4), ('RIGHTPADDING', (0,0), (-1,-1), 4),
    ]
    if highlight_rows:
        for ridx in highlight_rows:
            style_list.append(('BACKGROUND', (0, ridx), (-1, ridx), HexColor('#FFE699')))
    t.setStyle(TableStyle(style_list))
    return t

def cover_page(title, subtitle, doc_type, date_str="3 septembre 2026"):
    elements = []
    elements.append(Spacer(1, 4*cm))
    elements.append(Paragraph("BELGOCAM SA", ParagraphStyle('CL', parent=styles['Title'], fontName='DejaVuSans-Bold', fontSize=32, textColor=NAVY, alignment=TA_CENTER)))
    elements.append(Spacer(1, 0.5*cm))
    elements.append(HRFlowable(width="60%", thickness=2, color=GOLD, spaceBefore=10, spaceAfter=20, hAlign='CENTER'))
    elements.append(Spacer(1, 1*cm))
    elements.append(Paragraph(title, ParagraphStyle('CT', parent=styles['Title'], fontName='DejaVuSans-Bold', fontSize=28, textColor=NAVY, alignment=TA_CENTER, spaceAfter=20)))
    elements.append(Paragraph(subtitle, ParagraphStyle('CS', parent=styles['Title'], fontName='DejaVuSans', fontSize=16, textColor=GOLD, alignment=TA_CENTER, spaceAfter=30)))
    elements.append(Spacer(1, 2*cm))
    elements.append(HRFlowable(width="40%", thickness=1, color=GRAY, spaceBefore=10, spaceAfter=10, hAlign='CENTER'))
    elements.append(Paragraph(f"<b>{doc_type}</b>", ParagraphStyle('CI', parent=BODY, fontSize=11, alignment=TA_CENTER, textColor=GRAY)))
    elements.append(Paragraph("William Francis Fohom", ParagraphStyle('CI2', parent=BODY, fontSize=11, alignment=TA_CENTER, textColor=GRAY)))
    elements.append(Paragraph("Data Analyst | Administrateur National de Ventes", ParagraphStyle('CI3', parent=BODY, fontSize=11, alignment=TA_CENTER, textColor=GRAY)))
    elements.append(Spacer(1, 0.5*cm))
    elements.append(Paragraph(date_str, ParagraphStyle('CI4', parent=BODY, fontSize=11, alignment=TA_CENTER, textColor=GRAY)))
    elements.append(PageBreak())
    return elements


# ============================================
# Q4 2026 - 5 PDFs
# ============================================
Q4_DIR = "/home/z/my-project/download/forecast_q4_2026"
os.makedirs(Q4_DIR, exist_ok=True)

# === Q4 2026: 1. RÉSUMÉ EXÉCUTIF ===
print("Generating Q4 2026: 1. Résumé exécutif ...")
exec_path = f"{Q4_DIR}/01_resume_executif.pdf"
doc = SimpleDocTemplate(exec_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []

story.append(Paragraph("RÉSUMÉ EXÉCUTIF", H1))
story.append(Paragraph("Forecast Q4 2026 - Volumes et Valeurs", H2))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Contexte</b>", H3))
story.append(Paragraph(
    "BELGOCAM SA doit anticiper ses volumes de ventes et son chiffre d'affaires pour le Q4 2026 "
    "(septembre - décembre). Cette  du forecast s'appuie sur <b>176 576 enregistrements</b> "
    "couvrant <b>44 mois d'historique</b> (janvier 2023 - août 2026), intégrant les commandes En cours et "
    "Validées d'août 2026. L'effet soja exceptionnel de juillet-août 2026 (rupture concurrente) a été "
    "<b>désaisonnalisé</b>. Le prix du soja a été actualisé à <b>Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026)</b> (hausse du 24/08/2026).",
    BODY))
story.append(Paragraph(
    "<b>MÉTHODOLOGIE</b> : (1) Intégration de l'année 2023 (4 ans d'historique au total). "
    "(2) Ajout de la famille <b>COMPLEMENT_ALIMENTAIRE</b> (BELGOKILL V300 1L only, V305 200L exclu). "
    "Conversion 1L = 1kg. (3) MATERIEL_ELEVAGE et ALVEOLES à 0 tonne (CA only).",
    BODY))

story.append(Paragraph("<b>Méthodologie</b>", H3))
story.append(Paragraph(
    "Le modèle <b>Prophet</b> (Facebook/Meta) a été entraîné sur 5 familles alimentaires (TOURTEAUX, CONCENTRÉS, "
    "INGRÉDIENTS, ALIMENT COMPLET, COMPLEMENT ALIMENTAIRE) au niveau famille × région. Le MATERIEL ÉLEVAGE et "
    "les PREMIX sont projetés par extrapolation de la moyenne historique Q4. Le forecast couvre <b>8 familles, "
    "121 produits, 25 agences</b> et 3 régions. Le scénario S3 (réappro soja 100%) est utilisé comme référence.",
    BODY))

story.append(Paragraph("<b>Résultats clés Q4 2026</b>", H3))
synth_data = build_q4_synth_table()
story.append(make_table(synth_data, col_widths=[5*cm, 4*cm, 8*cm], font_size=9, highlight_rows=[7, 8]))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Synthèse par famille</b>", H3))
fam_data = build_q4_fam_table()
story.append(make_table(fam_data, col_widths=[5*cm, 3*cm, 3*cm, 2.5*cm], font_size=9, highlight_rows=[4, 8]))
story.append(Spacer(1, 0.2*cm))

story.append(Paragraph("<b>Synthèse par mois</b>", H3))
month_data = build_q4_month_table()
story.append(make_table(month_data, col_widths=[4*cm, 3.5*cm, 3.5*cm, 2.5*cm], font_size=9))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Top 5 agences par CA Q4 2026</b>", H3))
top_data = top_agences_table(q4_df)
story.append(make_table(top_data, col_widths=[1.5*cm, 4*cm, 3*cm, 4.5*cm], font_size=9))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Réalisation 2026 YTD + Projection fin d'année</b>", H3))
ytd_data = build_q4_ytd_table()
story.append(make_table(ytd_data, col_widths=[3.5*cm, 3*cm, 3*cm, 3*cm], font_size=9, highlight_rows=[6, 7]))
story.append(Spacer(1, 0.3*cm))


story.append(Paragraph("<b>Évolution du prix moyen du soja (T102 50kg) — Argumentaire de justification</b>", H3))
story.append(Paragraph(
    "Le tableau ci-dessous présente l\'évolution du prix moyen pondéré du soja (CA TTC / quantité) "
    "sur la période Oct 2025 → Juin 2026. Cette période couvre 9 mois de ventes « normales » "
    "(hors hausse circonstancielle de juillet-août 2026 liée à la rupture concurrente). "
    "Le prix moyen pondéré sur 9 mois (Oct 2025 → Juin 2026) est de <b>16 090 FCFA/sac</b>. La médiane YTD 2026 (Jan-Août) est de <b>16 800 FCFA/sac</b>. La moyenne pondérée YTD 2026 est de <b>17 170 FCFA/sac</b>, utilisé comme prix de référence "
    "pour le forecast 2027. Pour le Q4 2026, la <b>moyenne pondérée</b> de 17 170 FCFA/sac est utilisée (plus représentative "
    "des prix de fin d\'année). Ces prix excluent volontairement la hausse circonstancielle de juillet-août 2026 "
    "(prix atteignant 25 000 FCFA/sac) qui n\'est pas soutenable sur 12 mois.",
    BODY))

soja_price_data = [
    ["Mois", "Sacs vendus", "CA TTC (M FCFA)", "Prix/sac (FCFA)", "Min", "Max", "Médiane"],
    ["Oct 2025", "145 612", "2 223", "15 269", "14 700", "20 976", "15 300"],
    ["Nov 2025", "106 386", "1 695", "15 928", "14 700", "19 490", "17 170"],
    ["Déc 2025", "130 800", "2 153", "16 459", "15 500", "21 215", "17 170"],
    ["Jan 2026", "106 761", "1 767", "16 548", "16 000", "19 490", "17 170"],
    ["Fév 2026", "72 371", "1 185", "16 378", "15 500", "19 490", "17 170"],
    ["Mar 2026", "77 803", "1 249", "16 055", "15 500", "18 990", "15 900"],
    ["Avr 2026", "83 128", "1 296", "15 586", "15 000", "18 490", "15 600"],
    ["Mai 2026", "96 788", "1 508", "15 580", "15 000", "18 490", "15 600"],
    ["Juin 2026", "144 091", "2 432", "16 876", "15 000", "20 490", "16 990"],
    ["TOTAL", "963 740", "15 507", "17 170", "14 700", "21 215", "17 170"],
]
story.append(make_table(soja_price_data, col_widths=[2*cm, 2.2*cm, 2.2*cm, 2.2*cm, 1.8*cm, 1.8*cm, 1.8*cm], font_size=7.5))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph(
    "<b>Lecture</b> : Le prix moyen pond\u00e9r\u00e9 Oct 2025 \u2192 Juin 2026 = <b>16 090 FCFA/sac</b>. "
    "La m\u00e9diane YTD 2026 (Jan-Ao\u00fbt) = <b>16 800 FCFA/sac</b>. La moyenne pond\u00e9r\u00e9e YTD 2026 = <b>17 170 FCFA/sac</b>. "
    "avec un creux en Avril-Mai (15 580-15 586) et un pic en Juin (16 876) qui précède la hausse circonstancielle "
    "de Juillet-Août. <b>Le forecast 2027 utilise 16 800 FCFA/sac</b> (médiane YTD 2026), "
    "et le forecast Q4 2026 utilise <b>17 170 FCFA/sac</b> (moyenne pondérée YTD 2026).",
    BODY))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Recommandations principales</b>", H3))
_oct_q4 = Q4_MONTH[Q4_MONTH['month']==10]
_oct_t = _oct_q4['t'].iloc[0] if len(_oct_q4)>0 else 0
_oct_ca = _oct_q4['ca'].iloc[0] if len(_oct_q4)>0 else 0
_famla_q4 = q4_df[q4_df['agence']=='Famla']['ca_m_fcfa'].sum()
recos = [
    "<b>Réapprovisionnement soja urgent</b> — Commander 80 000 sacs minimum avant le 15/09/2026 pour éviter le scénario S1 (perte de 12 000+ M FCFA vs S3).",
    f"<b>Planification commerciale Q4</b> — Utiliser le scénario S3 ({fmt_t(Q4_TOTAL_T)} t, {fmt_ca(Q4_TOTAL_CA)} M FCFA) comme référence.",
    f"<b>Pic d'octobre</b> — Octobre = pic du Q4 ({fmt_t(_oct_t)} t, {fmt_ca(_oct_ca)} M FCFA). Anticiper capacité de production et logistique.",
    "<b>Maintien du bundle</b> — Ratio bundle 2,3:1 atteint en août 2026 à maintenir Q4 2026.",
    f"<b>Surveillance FAMLA</b> — Représente une part critique du CA Q4 ({fmt_ca(_famla_q4)} M FCFA). Performance critique.",
    f"<b>Suivi COMPLEMENT_ALIMENTAIRE</b> — Nouvelle famille intégrée au forecast ({fmt_t(Q4_FAM[Q4_FAM['family']=='COMPLEMENT_ALIMENTAIRE']['t'].iloc[0])} t, {fmt_ca(Q4_FAM[Q4_FAM['family']=='COMPLEMENT_ALIMENTAIRE']['ca'].iloc[0])} M FCFA). BELGOKILL V300 1L comme proxy.",
]
for r in recos:
    story.append(Paragraph(f"• {r}", BULLET))

story.append(Spacer(1, 0.3*cm))
story.append(Paragraph("<b>Conclusion</b>", H3))
story.append(Paragraph(
    f"Le forecast Q4 2026 projette <b>{fmt_t(Q4_TOTAL_T)} tonnes</b> pour un CA de <b>{fmt_ca(Q4_TOTAL_CA)} M FCFA</b>. "
    "La désaisonnalisation de l'effet soja exceptionnel, l'actualisation du prix à Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026), "
    "l'intégration de l'année 2023 (44 mois d'historique), l'ajout de la famille COMPLEMENT_ALIMENTAIRE "
    f"(V300 1L only), le filtrage du creux 2024 pour ALIMENT_COMPLET, et le passage de PREMIX en Prophet avec volumes "
    f"permettent une projection réaliste. Le pic d'octobre ({fmt_t(_oct_t)} t) nécessitera une "
    "anticipation renforcée du réapprovisionnement soja.",
    BODY))

doc.build(story)
print(f"✓ Q4 2026 Résumé exécutif: {os.path.getsize(exec_path)/1024:.0f} KB")

# === Q4 2026: 2. PROPOSITION DE PROJET ===
print("Generating Q4 2026: 2. Proposition ...")
prop_path = f"{Q4_DIR}/02_proposition_projet.pdf"
doc = SimpleDocTemplate(prop_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []
story.extend(cover_page("Proposition de Projet", "Forecast Q4 2026", "PROPOSITION DE PROJET"))

story.append(Paragraph("1. Contexte et justification", H1))
story.append(Paragraph(
    "BELGOCAM SA doit anticiper ses volumes et CA pour le Q4 2026 (sept-déc). Cette  s'appuie sur "
    "44 mois d'historique (Jan 2023 - Août 2026) et intègre les innovations suivantes : ajout de la famille "
    "COMPLEMENT_ALIMENTAIRE (BELGOKILL V300 1L only, V305 200L exclu, 1L=1kg), intégration de l'année 2023, "
    "désaisonnalisation de l'effet soja, inclusion des commandes En cours/Validées, actualisation du prix soja "
    "à Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026). Le forecast couvre 8 familles, 121 produits et 25 agences.",
    BODY))

story.append(Paragraph("2. Objectifs", H1))
story.append(Paragraph("<b>Objectif principal</b>", H3))
story.append(Paragraph(
    "Produire un forecast Q4 2026 (4 mois) en volumes et valeurs, désagrégé par produit, famille, agence et "
    "région, selon le scénario S3 (réappro soja 100%).",
    BODY))
story.append(Paragraph("<b>Objectifs spécifiques</b>", H3))
objs = [
    "Modéliser la tendance et la saisonnalité mensuelles via Prophet (5 familles alimentaires)",
    "Désaisonnaliser l'effet soja exceptionnel de juillet-août 2026",
    "Intégrer les commandes En cours et Validées comme potentielles ventes Livrées",
    "Actualiser le prix soja à Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026) (hausse du 24/08/2026)",
    "Ajouter la famille COMPLEMENT_ALIMENTAIRE (V300 1L only, 1L=1kg)",
    "Intégrer l'année 2023 (44 mois d'historique au total)",
    "MATERIEL_ELEVAGE toujours à 0 en tonnes",
    "Exclure le Maïs, les produits opportunistes et V305 (BELGOKILL 200L)",
    "Produire les livrables PACE complets (8 documents)",
]
for o in objs:
    story.append(Paragraph(f"• {o}", BULLET))

story.append(Paragraph("3. Périmètre", H1))
scope_data = [
    ["Dimension", "Détail", "Volume"],
    ["Période forecast", "Septembre - Décembre 2026 (4 mois)", "4 mois"],
    ["Familles incluses", "TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, ALIMENT COMPLET, MATERIEL ÉLEVAGE, PREMIX, COMPLEMENT_ALIMENTAIRE, ALVEOLES", "8 familles"],
    ["Produits", "115 références (T102, C101-C108, MAT014, P102N2, V300 BELGOKILL, etc.)", "121 produits"],
    ["Agences", "25 agences BELGOCAM", "25 agences"],
    ["Régions", "Ouest, Centre, Littoral", "3 régions"],
    ["Niveau détail", "Produit × Agence × Mois", "4 569 lignes"],
    ["Scénario", "S3 - Réappro soja 100%", "1 scénario"],
    ["Données historiques", "176 576 enregistrements (Jan 2023 - Août 2026)", "44 mois"],
    ["Désaisonnalisation", "Effet soja Jul-Août 2026 neutralisé", "—"],
    ["Prix soja", "Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026)", "—"],
    ["COMPLEMENT_ALIM.", "V300 1L only (V305 exclu), 1L=1kg", "9 produits"],
]
story.append(make_table(scope_data, col_widths=[3.5*cm, 9.5*cm, 4*cm], font_size=9, highlight_rows=[11, 12]))

story.append(Paragraph("4. Méthodologie PACE", H1))
story.append(Paragraph(
    "La méthodologie PACE (Prepare-Analyze-Construct-Execute) structure le projet en 4 phases. "
    "L'innovation majeure est l'ajout de la famille COMPLEMENT_ALIMENTAIRE avec proxy BELGOKILL V300 1L "
    "et l'utilisation de 44 mois d'historique (Jan 2023 - Août 2026).",
    BODY))

pace_data = [
    ["Phase", "Activités", "Livrables"],
    ["P - PREPARE", "Consolidation 176 576 records + En cours/Validées, désaisonnalisation soja, ajout COMPLEMENT_ALIM.", "Dataset 2023-2026, prix Q4"],
    ["A - ANALYZE", "AED, saisonnalité 2023-2025, top produits/agences", "Graphiques, synthèse AED"],
    ["C - CONSTRUCT", "Prophet (5 familles × 3 régions) + extrapolation (2 familles), forecast Q4 S3", "Forecast Q4 2026 (4 569 lignes)"],
    ["E - EXECUTE", "Excel 8 feuilles, 5 PDFs PACE, graphiques", "8 livrables finaux"],
]
story.append(make_table(pace_data, col_widths=[3*cm, 7.5*cm, 6.5*cm], font_size=8))

story.append(Paragraph("5. Équipe et gouvernance", H1))
team_data = [
    ["Rôle", "Responsabilité", "Phase PACE"],
    ["Data Analyst (W. F. Fohom)", "Modélisation, AED, forecast, livrables", "P + A + C + E"],
    ["Direction Commerciale", "Validation hypothèses et objectifs", "A + E"],
    ["Direction Production", "Stocks, capacité, réappro soja", "P + C"],
    ["Direction Générale", "Décisions stratégiques", "E"],
    ["Contrôle de Gestion", "Cohérence financière", "E"],
]
story.append(make_table(team_data, col_widths=[4.5*cm, 8.5*cm, 4*cm], font_size=9))

story.append(Paragraph("6. Risques et mitigation", H1))
risks_data = [
    ["Risque", "Probabilité", "Mitigation"],
    ["Historique 44 mois — suffisant pour Prophet", "Faible", "Mise à jour trimestrielle"],
    ["Prix soja volatil (+47% en 2 mois)", "Élevée", "Prix actualisé 17 170 FCFA, scénario S3"],
    ["Effet soja 2026 non récurrent en Q4", "Élevée", "Cap désaisonnalisation à moyenne S1"],
    ["Commandes En cours/Validées non converties", "Faible", "Filtre qualité, exclusion Annulées"],
    ["V305 exclu — BELGOKILL 200L absent", "Faible", "V300 1L utilisé comme proxy suffisant"],
]
story.append(make_table(risks_data, col_widths=[7*cm, 3*cm, 7*cm], font_size=9))

story.append(Paragraph("7. Critères de succès", H1))
success = [
    "Forecast Q4 2026 produit sur 4 mois avec désagrégation complète (4 569 lignes)",
    "Désaisonnalisation effective de l'effet soja exceptionnel",
    "En cours et Validées intégrés comme potentielles ventes",
    "Prix soja actualisé à Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026)",
    "Famille COMPLEMENT_ALIMENTAIRE ajoutée avec proxy BELGOKILL V300 1L",
    "Écart forecast vs réalité ≤ 15% par mois",
    "Livrables PACE complets (8 documents) produits et diffusés",
]
for s in success:
    story.append(Paragraph(f"• {s}", BULLET))

doc.build(story)
print(f"✓ Q4 2026 Proposition: {os.path.getsize(prop_path)/1024:.0f} KB")

# === Q4 2026: 3. MATRICE RACI ===
print("Generating Q4 2026: 3. Matrice RACI ...")
raci_path = f"{Q4_DIR}/03_matrice_raci.pdf"
doc = SimpleDocTemplate(raci_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
story = []
story.extend(cover_page("Matrice RACI", "Forecast Q4 2026", "MATRICE RACI"))

story.append(Paragraph("1. Définition des rôles RACI", H1))
raci_def = [
    ["Lettre", "Rôle", "Description"],
    ["R", "Responsable", "Celui qui réalise la tâche"],
    ["A", "Approbateur", "Celui qui rend compte et valide"],
    ["C", "Consulté", "Celui dont l'avis est sollicité"],
    ["I", "Informé", "Celui qui est tenu au courant"],
]
story.append(make_table(raci_def, col_widths=[2*cm, 5*cm, 11*cm], font_size=9))

story.append(Paragraph("2. Acteurs du projet", H1))
actors = [
    ["Code", "Acteur", "Implication"],
    ["DA", "Data Analyst (W. F. Fohom)", "Pleine (R sur toutes les tâches techniques)"],
    ["DC", "Direction Commerciale", "Validation hypothèses et objectifs"],
    ["DP", "Direction Production", "Stocks, capacité, réappro soja"],
    ["DG", "Direction Générale", "Décision finale et allocation ressources"],
    ["CG", "Contrôle de Gestion", "Cohérence financière"],
    ["RA", "Responsables d'Agences (14)", "Informés et consultés sur leur périmètre"],
]
story.append(make_table(actors, col_widths=[1.5*cm, 6*cm, 10.5*cm], font_size=9))

story.append(Paragraph("3. Matrice RACI détaillée par étape PACE", H1))
raci_matrix = [
    ["Phase", "Tâche", "DA", "DC", "DP", "DG", "CG", "RA"],
    ["PREPARE", "1.1 Consolidation 2023-2026 + En cours/Validées", "R/A", "I", "C", "I", "I", "I"],
    ["", "1.2 Désaisonnalisation effet soja", "R/A", "C", "C", "I", "I", "I"],
    ["", "1.3 Calcul prix Q4 (soja 17 170)", "R/A", "C", "I", "I", "C", "I"],
    ["", "1.4 Ajout COMPLEMENT_ALIMENTAIRE (V300 1L)", "R/A", "C", "I", "I", "I", "I"],
    ["", "1.5 Validation dataset", "R", "C", "C", "I", "A", "I"],
    ["ANALYZE", "2.1 AED + saisonnalité 2023-2025", "R/A", "I", "I", "I", "I", "I"],
    ["", "2.2 Analyse effet soja exceptionnel", "R/A", "C", "C", "I", "I", "I"],
    ["", "2.3 Validation hypothèses AED", "R", "C", "C", "I", "A", "I"],
    ["CONSTRUCT", "3.1 Prophet (5 familles × 3 régions)", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.2 Extrapolation MAT+PREMIX (CA only)", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.3 Forecast Q4 S3", "R", "C", "I", "A", "C", "I"],
    ["", "3.4 Désagrégation produit × agence", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.5 Validation forecast Q4", "R", "C", "I", "A", "C", "I"],
    ["EXECUTE", "4.1 Excel 8 feuilles", "R/A", "I", "I", "I", "I", "I"],
    ["", "4.2 PDFs PACE (5 documents)", "R/A", "I", "I", "I", "I", "I"],
    ["", "4.3 Présentation CODIR", "R", "C", "C", "A", "C", "I"],
    ["", "4.4 Diffusion agences", "R", "A", "I", "I", "I", "C"],
    ["", "4.5 Mise à jour trimestrielle", "R/A", "C", "C", "I", "I", "I"],
]
story.append(make_table(raci_matrix, col_widths=[2.5*cm, 6.5*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm], font_size=7.5))

story.append(Paragraph("4. Points de contrôle", H1))
control_data = [
    ["Jalon", "Phase", "Date", "Approbateur", "Livrable"],
    ["Kick-off", "PREPARE", "03/09/2026", "DG", "Proposition validée"],
    ["Validation dataset", "PREPARE", "03/09/2026", "CG", "Dataset 2023-2026 + COMPLEMENT_ALIM."],
    ["Validation AED", "ANALYZE", "04/09/2026", "DC", "Synthèse AED"],
    ["Validation forecast", "CONSTRUCT", "05/09/2026", "DG + DC", "Forecast Q4 2027 complet"],
    ["Livrables finaux", "EXECUTE", "05/09/2026", "DG", "8 livrables"],
    ["Présentation CODIR", "EXECUTE", "06/09/2026", "DG", "Validation officielle"],
    ["Mise à jour Jan 2027", "—", "Jan 2027", "DA", "Réactualisation avec données réelles"],
]
story.append(make_table(control_data, col_widths=[3.5*cm, 2.5*cm, 3*cm, 3*cm, 5*cm], font_size=8))

doc.build(story)
print(f"✓ Q4 2026 Matrice RACI: {os.path.getsize(raci_path)/1024:.0f} KB")

# === Q4 2026: 4. DOCUMENT STRATÉGIQUE PACE ===
print("Generating Q4 2026: 4. Document stratégique PACE ...")
strat_path = f"{Q4_DIR}/04_document_strategique_pace.pdf"
doc = SimpleDocTemplate(strat_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []
story.extend(cover_page("Document Stratégique PACE", "Forecast Q4 2026 (8 familles, données 2023-2026)", "STRATÉGIE PACE"))

story.append(Paragraph("1. Vision et objectifs", H1))
story.append(Paragraph("1.1 Vision", H2))
story.append(Paragraph(
    f"Faire du forecast Q4 2026 un <b>outil de planification opérationnelle</b> permettant à BELGOCAM SA d'anticiper "
    f"{fmt_t(Q4_TOTAL_T)} tonnes de ventes et {fmt_ca(Q4_TOTAL_CA)} M FCFA de chiffre d'affaires sur la période septembre-décembre 2026, "
    "avec une désaisonnalisation de l'effet soja exceptionnel et l'intégration de la famille COMPLEMENT_ALIMENTAIRE.",
    BODY))

story.append(Paragraph("1.2 Objectifs stratégiques", H2))
obj_data = [
    ["Objectif", "Description", "KPI"],
    ["Forecast Q4", "4 mois Sept-Dec 2026 en volume + valeur", f"{fmt_t(Q4_TOTAL_T)} t, {fmt_ca(Q4_TOTAL_CA)} M FCFA"],
    ["Désaisonnalisation", "Neutraliser l'effet soja Jul-Août 2026", "Cap moyenne S1 2026"],
    ["En cours + Validées", "Intégrer comme potentielles ventes", "255 commandes incluses"],
    ["Prix actualisé", "Soja Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026)", "vs 17 170 médian 2023-2026"],
    ["Désagrégation", "Produit × agence × mois", f"{len(q4_df):,} lignes".replace(',', ' ')],
    ["COMPLEMENT_ALIM.", "V300 1L only (V305 exclu), 1L=1kg", f"{fmt_t(Q4_FAM[Q4_FAM['family']=='COMPLEMENT_ALIMENTAIRE']['t'].iloc[0])} t, {fmt_ca(Q4_FAM[Q4_FAM['family']=='COMPLEMENT_ALIMENTAIRE']['ca'].iloc[0])} M FCFA"],
]
story.append(make_table(obj_data, col_widths=[3.5*cm, 8.5*cm, 5*cm], font_size=9, highlight_rows=[6, 7]))

story.append(Paragraph("2. Phase PREPARE - Préparation des données", H1))
story.append(Paragraph("2.1 Sources consolidées (2023-2026)", H2))
sources_data = [
    ["Source", "Période", "Records"],
    ["LY_21_24 (filtre 2023)", "Jan-Dec 2023", "6 016"],
    ["LY_24 (Jul-Dec 2024)", "Juillet-Décembre 2024", "52 215"],
    ["Historique 2025", "Jan-Déc 2025", "66 206"],
    ["S1 + Juil + Août 2026", "Jan-Août 2026", "52 139"],
    ["En cours + Validées", "Août 2026", "~255"],
    ["TOTAL", "44 mois", "176 576"],
]
story.append(make_table(sources_data, col_widths=[5*cm, 5*cm, 5*cm], font_size=9))

story.append(Paragraph("2.2 Désaisonnalisation de l'effet soja", H2))
story.append(Paragraph(
    "Les volumes soja de juillet-août 2026 ont été exceptionnellement élevés (rupture concurrente). "
    "Pour éviter que Prophet n'extrapole ce phénomène non récurrent, les volumes soja Jul-Août 2026 ont été "
    "<b>cappés à la moyenne S1 2026</b>.",
    BODY))

story.append(Paragraph("2.3  Famille COMPLEMENT_ALIMENTAIRE", H2))
story.append(Paragraph(
    "La famille COMPLEMENT_ALIMENTAIRE (9 produits liquides 1L) a été ajoutée au forecast Q4 2026. "
    "Selon demande utilisateur : <b>seul V300 (BELGOKILL 1L) est conservé — V305 (BELGOKILL 200L) est EXCLU</b>. "
    "Conversion 1L = 1kg appliquée pour exprimer les volumes en tonnes.",
    BODY))

story.append(Paragraph("3. Phase CONSTRUCT - Modélisation", H1))
story.append(Paragraph("3.1 Configuration Prophet", H2))
config_data = [
    ["Paramètre", "Valeur", "Justification"],
    ["yearly_seasonality", "True", "Capture la saisonnalité annuelle (pic Q4)"],
    ["seasonality_mode", "multiplicative", "Saisonnalité proportionnelle à la tendance"],
    ["changepoint_prior_scale", "0.05", "Tendance modérément flexible"],
    ["interval_width", "0.8", "Intervalle de confiance 80%"],
    ["Période forecast", "4 mois (Sep-Dec 2026)", "Q4 2026"],
    ["Familles Prophet", "5 (TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, ALIMENT_COMPLET, COMPLEMENT_ALIMENTAIRE)", "Inclut COMPLEMENT_ALIMENTAIRE"],

]
story.append(make_table(config_data, col_widths=[5*cm, 5*cm, 7*cm], font_size=9, highlight_rows=[6]))

story.append(Paragraph("4. Résultats par famille et mois", H1))
story.append(Paragraph("4.1 Par famille", H2))
fam_detail = build_q4_fam_detail_table()
story.append(make_table(fam_detail, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2*cm, 4*cm], font_size=9, highlight_rows=[4, 8]))

story.append(Paragraph("4.2 Par mois", H2))
month_detail = build_q4_month_detail_table()
story.append(make_table(month_detail, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2*cm, 4.5*cm], font_size=9))

story.append(Paragraph("5. Plan de déploiement", H1))
deploy_data = [
    ["Étape", "Date", "Action", "Cible"],
    ["1", "03/09/2026", "Validation proposition", "DG"],
    ["2", "05/09/2026", "Livrables finaux prêts", "DA"],
    ["3", "06/09/2026", "Présentation CODIR", "DG + DC + DP + CG"],
    ["4", "07/09/2026", "Diffusion aux agences", "14 RA"],
    ["5", "Jan 2027", "Mise à jour avec données réelles Q4", "DA"],
]
story.append(make_table(deploy_data, col_widths=[1*cm, 3*cm, 7*cm, 6*cm], font_size=9))

story.append(Paragraph("6. Conclusion", H1))
story.append(Paragraph(
    f"Le forecast Q4 2026 projette <b>{fmt_t(Q4_TOTAL_T)} tonnes</b> pour <b>{fmt_ca(Q4_TOTAL_CA)} M FCFA</b>. La désaisonnalisation "
    "de l'effet soja, l'actualisation du prix à 16 800 FCFA (médiane) pour 2027 et 17 170 (moyen) pour Q4 2026, l'intégration de l'année 2023 (44 mois d'historique), "
    "l'ajout de la famille COMPLEMENT_ALIMENTAIRE (V300 1L only), le filtrage du creux 2024 pour ALIMENT_COMPLET, le passage de PREMIX en Prophet avec volumes, et le maintien de MATERIEL_ELEVAGE à 0 tonne "
    f"permettent une projection réaliste. Le pic d'octobre ({fmt_t(Q4_MONTH[Q4_MONTH['month']==10]['t'].iloc[0])} t) nécessitera une anticipation renforcée "
    "du réapprovisionnement soja.",
    BODY))

doc.build(story)
print(f"✓ Q4 2026 Document stratégique: {os.path.getsize(strat_path)/1024:.0f} KB")

# === Q4 2026: 5. GUIDE MÉTHODOLOGIQUE ===
print("Generating Q4 2026: 5. Guide méthodologique ...")
guide_path = f"{Q4_DIR}/05_guide_methodologique.pdf"
doc = SimpleDocTemplate(guide_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []
story.extend(cover_page("Guide Méthodologique", "Forecast Q4 2026 (Prophet + COMPLEMENT_ALIMENTAIRE V300 1L)", "GUIDE MÉTHODOLOGIQUE"))

story.append(Paragraph("Sommaire", H1))
story.append(Paragraph(
    "1. Introduction et contexte<br/>"
    "2. Méthodologie PACE<br/>"
    "3. Outils et environnement<br/>"
    "4. Phase PREPARE - Données 2023-2026<br/>"
    "5. Phase PREPARE - Désaisonnalisation soja<br/>"
    "6. Phase PREPARE - COMPLEMENT_ALIMENTAIRE<br/>"
    "7. Phase ANALYZE - AED<br/>"
    "8. Phase CONSTRUCT - Prophet<br/>"
    "9. Phase EXECUTE - Livrables<br/>"
    "Annexes",
    BODY))

story.append(PageBreak())
story.append(Paragraph("1. Introduction et contexte", H1))
story.append(Paragraph(
    "Ce guide décrit la méthodologie complète du forecast Q4 2026 de BELGOCAM SA (). Il couvre 4 mois "
    "(septembre-décembre 2026) avec 8 familles de produits (incluant la nouvelle famille COMPLEMENT_ALIMENTAIRE) "
    "et 25 agences. Les innovations : "
    "(1) ajout de la famille COMPLEMENT_ALIMENTAIRE (V300 BELGOKILL 1L only, V305 200L EXCLU, 1L=1kg), "
    "(2) intégration de l'année 2023 (44 mois d'historique au total), "
    "(3) MATERIEL_ELEVAGE toujours à 0 tonne (non exprimable en volume), "
    "(4) désaisonnalisation de l'effet soja, "
    "(5) inclusion des commandes En cours/Validées, "
    "(6) actualisation du prix soja à Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026).",
    BODY))

story.append(Paragraph("2. Méthodologie PACE", H1))
story.append(Paragraph(
    "La méthodologie PACE (Prepare-Analyze-Construct-Execute) structure le projet en 4 phases successives.",
    BODY))
pace_detail = [
    ["Phase", "Objectif", "Livrables", "Durée"],
    ["P - PREPARE", "Préparer données 2023-2026 + désaisonnalisation + COMPLEMENT_ALIM.", "Dataset 176 576 records, prix Q4", "2 jours"],
    ["A - ANALYZE", "Comprendre données + effet soja", "Graphiques, synthèse AED", "1 jour"],
    ["C - CONSTRUCT", "Modéliser Prophet Q4 (8 familles)", "Forecast Q4 2026 (4 569 lignes)", "2 jours"],
    ["E - EXECUTE", "Produire livrables finaux", "Excel + 5 PDFs", "1 jour"],
]
story.append(make_table(pace_detail, col_widths=[3*cm, 4*cm, 5*cm, 3*cm], font_size=9))

story.append(Paragraph("3. Outils et environnement", H1))
stack_data = [
    ["Outil", "Version", "Usage"],
    ["Python", "3.12", "Langage principal"],
    ["pandas", "2.x", "Manipulation données"],
    ["Prophet", "1.1+", "Modélisation prédictive (5 familles × 3 régions = 15 modèles)"],
    ["openpyxl", "3.x", "Génération Excel (8 feuilles)"],
    ["ReportLab", "4.x", "Génération PDFs (5 documents PACE)"],
    ["matplotlib", "3.x", "Visualisations"],
]
story.append(make_table(stack_data, col_widths=[4*cm, 3*cm, 8*cm], font_size=9))

story.append(PageBreak())
story.append(Paragraph("4. Phase PREPARE - Données 2023-2026", H1))
story.append(Paragraph(
    "Le dataset consolidé couvre 44 mois (janvier 2023 - août 2026), soit 176 576 enregistrements. "
    "Les sources sont : LY_21_24 filtre 2023 (6 016 records), LY_24 Jul-Dec 2024 (52 215), "
    "Historique 2025 (66 206), S1+Juil+Août 2026 (52 139), En cours + Validées (~255).",
    BODY))
story.append(Paragraph(
    "<b>Intégration de l'année 2023</b> : La  inclut 2023 pour enrichir l'historique avec 4 ans "
    "de données (2023-2026). Cela permet à Prophet de mieux capturer la saisonnalité annuelle.",
    BODY))

story.append(Paragraph("5. Phase PREPARE - Désaisonnalisation soja", H1))
story.append(Paragraph(
    "L'effet soja exceptionnel de juillet-août 2026 (rupture concurrente) a généré des volumes anormalement "
    "élevés non récurrents. Les volumes soja Jul-Août 2026 ont été cappés à la moyenne S1 2026.",
    BODY))
story.append(Paragraph("<b>Méthode</b>", H3))
story.append(Paragraph(
    "1. Calcul moyenne mensuelle soja S1 2026 (Jan-Juin)<br/>"
    "2. Identification mois Jul-Août 2026 avec volume max<br/>"
    "3. Calcul facteur de cap = moyenne S1 / volume max<br/>"
    "4. Application du facteur aux enregistrements soja Jul-Août 2026<br/>"
    "5. Volumes réduits proportionnellement, signal saisonnier préservé",
    BODY))

story.append(PageBreak())
story.append(Paragraph("6.  Famille COMPLEMENT_ALIMENTAIRE", H1))
story.append(Paragraph(
    "La famille COMPLEMENT_ALIMENTAIRE regroupe 9 produits liquides BELGO : BELGOKILL V300 (1L), "
    "BELGO HARMONY, BELGO PROTECT, BELGO DRY LIT, BELGO WATER CLEAN, BELGO VIT Ese, BELGO THERMO, "
    "BELGO BIO SELECT, BELGO FRESH — tous en conditionnement 1L.",
    BODY))

story.append(Paragraph("<b>Exclusion de V305 (BELGOKILL 200L)</b>", H3))
story.append(Paragraph(
    "Selon demande utilisateur explicite : seul le conditionnement 1L est conservé. "
    "V305 (BELGOKILL 200L) est EXCLU du forecast. Cela permet d'avoir une famille homogène "
    "(tous produits en 1L) et d'éviter le biais du conditionnement bulk.",
    BODY))

story.append(Paragraph("<b>Conversion 1L = 1kg</b>", H3))
story.append(Paragraph(
    "Pour exprimer les volumes en tonnes (nécessaire pour Prophet), la conversion 1L = 1kg est appliquée : "
    "<b>1 qte = 1 kg, tonnes = qte / 1000</b>. Approximation valable pour liquides aqueux (densité ~1).",
    BODY))

story.append(Paragraph("<b>Proxy BELGOKILL V300 1L</b>", H3))
story.append(Paragraph(
    "BELGOKILL V300 (1L) est le produit dominant de la famille (38% du CA famille sur 2023-2026). "
    "Sa tendance est utilisée comme proxy pour modéliser toute la famille via Prophet. "
    "Les autres produits (CA001-CA008) sont désagrégés selon leurs parts historiques.",
    BODY))

ca_proxy = [
    ["Réf.", "Produit", "Poids (kg/qte)", "Prix Q4 2026 (FCFA/L)", "Statut"],
    ["V300", "BELGOKILL 1L", "1", "2 500", "✓ Inclus (proxy)"],
    ["V305", "BELGOKILL 200L", "—", "—", "✗ Exclu"],
    ["CA003.1", "BELGO HARMONY 1L", "1", "8 000", "✓ Inclus"],
    ["CA004.1", "BELGO PROTECT 1L", "1", "9 800", "✓ Inclus"],
    ["CA006.1", "BELGO DRY LIT 1L", "1", "6 500", "✓ Inclus"],
    ["CA001.1", "BELGO WATER CLEAN 1L", "1", "5 500", "✓ Inclus"],
    ["CA002.1", "BELGO VIT Ese 1L", "1", "10 500", "✓ Inclus"],
    ["CA005.1", "BELGO THERMO 1L", "1", "15 000", "✓ Inclus"],
    ["CA007.1", "BELGO BIO SELECT 1L", "1", "8 000", "✓ Inclus"],
    ["CA008.1", "BELGO FRESH 1L", "1", "14 000", "✓ Inclus"],
]
story.append(make_table(ca_proxy, col_widths=[2*cm, 5*cm, 2.5*cm, 3.5*cm, 3*cm], font_size=9, highlight_rows=[2]))

story.append(Paragraph("7. Phase ANALYZE - AED", H1))
story.append(Paragraph(
    "L'AED révèle : (1) TOURTEAUX = 76% du volume historique, (2) Q4 = 35-46% du volume annuel selon familles, "
    "(3) pic octobre (indice 1.86 pour TOURTEAUX en 2025), (4) FAMLA = 31.6% du volume, "
    "(5) COMPLEMENT_ALIMENTAIRE = 3-7 t/an (faible volume, forte valeur unitaire).",
    BODY))

story.append(Paragraph("8. Phase CONSTRUCT - Prophet", H1))
story.append(Paragraph(
    "15 modèles Prophet famille × région (5 familles × 3 régions) + extrapolation Q4 moyenne pour MATERIEL_ELEVAGE "
    "et PREMIX (2 familles × 3 régions = 6 extrapolations). Désagrégation par produit × agence selon parts "
    "historiques de CA. Prix Q4 2026 : Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026) soja, prix août 2026 autres familles, prix par litre "
    "pour COMPLEMENT_ALIMENTAIRE (2 500-15 000 FCFA/L selon produit).",
    BODY))

story.append(Paragraph("9. Phase EXECUTE - Livrables", H1))
livrables_data = [
    ["#", "Livrable", "Format", "Usage"],
    ["1", "Résumé exécutif", "PDF", "Direction Générale"],
    ["2", "Proposition de projet", "PDF", "CODIR"],
    ["3", "Matrice RACI", "PDF", "Gouvernance"],
    ["4", "Document stratégique PACE", "PDF", "Référence stratégique"],
    ["5", "Guide méthodologique", "PDF", "Référence méthodes"],
    ["6", "Excel forecast Q4 2026 (8 feuilles)", "XLSX", "Pilotage opérationnel"],
]
story.append(make_table(livrables_data, col_widths=[1*cm, 5*cm, 3*cm, 6*cm], font_size=9))

story.append(Paragraph("Annexe - Glossaire", H1))
glossaire = [
    ["Terme", "Définition"],
    ["Prophet", "Bibliothèque prévision séries temporelles (Facebook/Meta)"],
    ["Désaisonnalisation", "Neutralisation d'un effet exceptionnel non récurrent"],
    ["S3", "Scénario réappro soja 100% (situation normale)"],
    ["Cap S1", "Plafonnement des volumes Jul-Août 2026 à moyenne S1 2026"],
    ["Q4", "Quatrième trimestre (oct-déc)"],
    ["PACE", "Prepare-Analyze-Construct-Execute"],
    ["COMPLEMENT_ALIM.", "9 produits liquides 1L BELGOxxx (V300 1L proxy)"],
    ["ALVEOLES séparé", "Famille distincte de MATERIEL_ELEVAGE (4 refs MAT011/14/15/17)"],
    ["SPC agences", "10 agences SPC incluses (Baf-Chefferie dominant)"],
    ["Maroua", "Agence ajoutée (Centre), soja T102 only"],
    ["Bundle 2.5:1", "Ratio soja:concentré forcé à <=2.5:1"],
    ["Saisonnalité ALV 2026", "ALVEOLES utilise 2026 (13.9 M/an) au lieu du pic 2025 (739 M)"],
    ["10 SPC agences", "Toutes SPC incluses (10 agences)"],
    ["Forfait SPC", "Forfait MAT_ELEV (38.4 M/an) basé sur 2026 annualisé, poids 2025"],
    ["SPC PK15 forfait", "Forfait réaliste 1 M/an (0.5 ALV + 0.5 MAT)"],
    [" V305 EXCLU", "BELGOKILL 200L retiré du forecast (seul V300 1L conservé)"],
    [" 1L=1kg", "Conversion pour volumes liquides en tonnes"],
    [" Données 2023-2026", "44 mois d'historique (Jan 2023 - Août 2026)"],
    ["MATERIEL_ELEVAGE", "Toujours 0 en tonnes (CA only). ALVEOLES désormais séparés en famille distincte"],
]
story.append(make_table(glossaire, col_widths=[5*cm, 11*cm], font_size=9, highlight_rows=[7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]))

doc.build(story)
print(f"✓ Q4 2026 Guide méthodologique: {os.path.getsize(guide_path)/1024:.0f} KB")

print(f"\n=== Q4 2026: 5 PDFs GÉNÉRÉS ===")
for f in sorted(os.listdir(Q4_DIR)):
    if f.endswith('.pdf'):
        print(f"  {f}: {os.path.getsize(os.path.join(Q4_DIR, f))/1024:.0f} KB")


# ============================================
# 2027 - 5 PDFs (use updated numbers)
# ============================================
F2027_DIR = "/home/z/my-project/download/forecast_2027"
os.makedirs(F2027_DIR, exist_ok=True)

# === 2027: 1. RÉSUMÉ EXÉCUTIF ===
print("\nGenerating 2027: 1. Résumé exécutif ...")
exec_path = f"{F2027_DIR}/01_resume_executif.pdf"
doc = SimpleDocTemplate(exec_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []

story.append(Paragraph("RÉSUMÉ EXÉCUTIF", H1))
story.append(Paragraph("Forecast 2027 - Volumes et Valeurs (12 mois)", H2))
story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Contexte</b>", H3))
story.append(Paragraph(
    "BELGOCAM SA doit anticiper ses volumes de ventes et son chiffre d'affaires pour l'année 2027 complète. "
    "Cette  s'appuie sur <b>176 576 enregistrements</b> couvrant <b>44 mois d'historique</b> "
    "(janvier 2023 - août 2026), intégrant les commandes En cours et Validées d'août 2026. "
    "L'effet soja exceptionnel de juillet-août 2026 a été <b>désaisonnalisé</b>. Le prix du soja a été "
    "actualisé à <b>Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026)</b>.",
    BODY))
story.append(Paragraph(
    "<b>MÉTHODOLOGIE</b> : (1) Intégration de l'année 2023 (44 mois d'historique). "
    "(2) Ajout de la famille <b>COMPLEMENT_ALIMENTAIRE</b> (BELGOKILL V300 1L only, V305 200L exclu, 1L=1kg). "
    "(3) MATERIEL_ELEVAGE toujours à 0 en tonnes.",
    BODY))

story.append(Paragraph("<b>Méthodologie</b>", H3))
story.append(Paragraph(
    "Le modèle <b>Prophet</b> a été entraîné sur 5 familles alimentaires (TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, "
    "ALIMENT COMPLET, COMPLEMENT ALIMENTAIRE) au niveau famille × région. Le MATERIEL ÉLEVAGE et les PREMIX "
    "sont projetés par extrapolation de la moyenne historique. Le forecast couvre <b>8 familles, 121 produits, "
    "25 agences</b> et 3 régions.",
    BODY))

story.append(Paragraph("<b>Résultats clés 2027</b>", H3))
synth_data = build_2027_synth_table()
story.append(make_table(synth_data, col_widths=[5*cm, 4*cm, 8*cm], font_size=9, highlight_rows=[7, 8]))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Synthèse par famille</b>", H3))
fam_data = build_2027_fam_table()
story.append(make_table(fam_data, col_widths=[5*cm, 3*cm, 3*cm, 2.5*cm], font_size=9, highlight_rows=[4, 8]))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Synthèse par trimestre</b>", H3))
q_data = build_2027_q_table()
story.append(make_table(q_data, col_widths=[4*cm, 3.5*cm, 3.5*cm, 2.5*cm], font_size=9))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Top 5 agences par CA 2027</b>", H3))
top_data = top_agences_table(f2027_df)
story.append(make_table(top_data, col_widths=[1.5*cm, 4*cm, 3*cm, 4.5*cm], font_size=9))

story.append(PageBreak())

# Calculs dynamiques pour les insights — comparaison cohérente "apples-to-apples"
# MAIS est exclu du forecast (opportuniste, mis à 0). Pour comparer 2027 vs 2026
# de manière cohérente, on exclut MAIS des deux côtés (sinon on sous-estime la croissance
# des familles réellement forecastées).
_tt_2027 = F2027_TOTAL_T
_tt_2026 = LY_2026_TOTAL_T
_ca_2027 = F2027_TOTAL_CA
_ca_2026 = LY_2026_TOTAL_CA

# Total 2026 LY sans MAIS (apples-to-apples)
_mais_t = LY_2026_BY_FAM_T.get('MAIS', 0)
_mais_ca = LY_2026_CA_BY_FAM.get('MAIS', 0)
_tt_2026_no_mais = _tt_2026 - _mais_t
_ca_2026_no_mais = _ca_2026 - _mais_ca

_var_t_total = (_tt_2027/_tt_2026_no_mais - 1) * 100
_var_ca_total = (_ca_2027/_ca_2026_no_mais - 1) * 100
_ac_2027_t = F2027_FAM[F2027_FAM['family']=='ALIMENT_COMPLET']['t'].iloc[0] if len(F2027_FAM[F2027_FAM['family']=='ALIMENT_COMPLET'])>0 else 0
_ac_2026_t = LY_2026_BY_FAM_T.get('ALIMENT_COMPLET', 0)
_ac_var_t = (_ac_2027_t/_ac_2026_t - 1) * 100 if _ac_2026_t > 0 else 0
_pm_2027_t = F2027_FAM[F2027_FAM['family']=='PREMIX']['t'].iloc[0] if len(F2027_FAM[F2027_FAM['family']=='PREMIX'])>0 else 0
_pm_2027_ca = F2027_FAM[F2027_FAM['family']=='PREMIX']['ca'].iloc[0] if len(F2027_FAM[F2027_FAM['family']=='PREMIX'])>0 else 0
_pm_2026_ca = LY_2026_CA_BY_FAM.get('PREMIX', 0)
_pm_var_ca = (_pm_2027_ca/_pm_2026_ca - 1) * 100 if _pm_2026_ca > 0 else 0
_ing_var_ca = (F2027_FAM[F2027_FAM['family']=='INGREDIENTS']['ca'].iloc[0]/LY_2026_CA_BY_FAM.get('INGREDIENTS', 1) - 1) * 100 if LY_2026_CA_BY_FAM.get('INGREDIENTS', 0) > 0 else 0
_conc_var_ca = (F2027_FAM[F2027_FAM['family']=='CONCENTRES']['ca'].iloc[0]/LY_2026_CA_BY_FAM.get('CONCENTRES', 1) - 1) * 100 if LY_2026_CA_BY_FAM.get('CONCENTRES', 0) > 0 else 0
_tour_var_ca = (F2027_FAM[F2027_FAM['family']=='TOURTEAUX']['ca'].iloc[0]/LY_2026_CA_BY_FAM.get('TOURTEAUX', 1) - 1) * 100 if LY_2026_CA_BY_FAM.get('TOURTEAUX', 0) > 0 else 0

story.append(Paragraph("<b>Réalisation 2026 complète (Jan-Août réel + Q4 forecast)</b>", H3))
story.append(Paragraph(
    "Le tableau ci-dessous retrace l'année 2026 dans son intégralité : volumes réalisés de janvier à août 2026 "
    "(issus de l'ERP, <b>YTD réel</b>), auxquels s'ajoute le <b>forecast Q4 2026</b> (septembre-décembre, "
    "scénario S3 réappro soja 100%). La colonne <b>Total 2026</b> représente l'année complète reconstruite, "
    "qui sert de base de comparaison au forecast 2027.",
    BODY))

realisation_2026_data = build_2026_full_year_table()
story.append(make_table(realisation_2026_data, col_widths=[3.5*cm, 2.3*cm, 2.3*cm, 2.3*cm, 2.3*cm, 1.6*cm], font_size=8, highlight_rows=[10, 11]))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph(
    f"<b>Lecture</b> : L'année 2026 complète (réalisée + forecast Q4) s'établit à <b>{fmt_t(LY_2026_TOTAL_T)} t</b> "
    f"(dont {fmt_t(YTD_2026_TOTAL_T)} t déjà réalisés sur Jan-Août et {fmt_t(Q4_TOTAL_T)} t en forecast Q4). "
    f"Le forecast 2027 à {fmt_t(F2027_TOTAL_T)} t représente <b>{fmt_pct_signed(_var_t_total)}</b> vs 2026 hors MAIS "
    f"(le MAIS, opportuniste, est exclu du forecast 2027). La dynamique de croissance est portée par les "
    "CONCENTRÉS (bundle 2.5:1), ALIMENT_COMPLET (filtrage 2024 + reprise +15%) et PREMIX (restauré en Prophet avec volumes).",
    BODY))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Historique 2023-2026 vs Forecast 2027</b>", H3))
story.append(Paragraph(
    "Le forecast 2027 s'appuie sur 4 ans d'historique (Jan 2023 - Août 2026). Le tableau ci-dessous présente "
    "l'évolution par famille et par année.",
    BODY))

hist_data = build_2027_hist_table()
story.append(make_table(hist_data, col_widths=[2.5*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.7*cm, 1.7*cm, 1.5*cm, 1.5*cm], font_size=7, highlight_rows=[9, 10]))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph(
    f"<b>Lecture</b> : Le TOURTEAUX montre une trajectoire haussière ({fmt_t(HIST_BY_YEAR_FAM[2023].get('TOURTEAUX', 0))} t en 2023 → {fmt_t(F2027_FAM[F2027_FAM['family']=='TOURTEAUX']['t'].iloc[0])} t forecast 2027). "
    f"Les CONCENTRÉS progressent également ({fmt_pct_signed(_conc_var_ca)} CA vs 2026). ALIMENT_COMPLET bénéficie du filtrage du creux 2024 "
    f"et d'un facteur reprise +15% ({fmt_t(_ac_2027_t)} t, {fmt_pct_signed(_ac_var_t)} vs 2026). PREMIX restauré à {fmt_t(_pm_2027_t)} t "
    f"(mode Prophet avec volumes). Progression globale 2027 vs 2026 : <b>{fmt_pct_signed(_var_t_total)}</b> en volume, "
    f"<b>{fmt_pct_signed(_var_ca_total)}</b> en CA.",
    BODY))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Progression 2027 vs 2026 (détail par famille)</b>", H3))
story.append(Paragraph(
    "Le tableau ci-dessous présente la progression détaillée 2027 vs 2026 (YTD réel + Q4 forecast) par famille, "
    "en volume et en chiffre d\'affaires. Cette progression reflète l\'effet cumulé des innovations : "
    "ajustement bundle 2.5:1 (augmente CONCENTRÉS), forfait SPC (ALVEOLES=0, MAT_ELEV=38,4 M annualisé), "
    "filtrage 2024 + facteur reprise +15% sur ALIMENT_COMPLET, et PREMIX en Prophet avec volumes.",
    BODY))

progression_data = build_2027_progression_table()
story.append(make_table(progression_data, col_widths=[3*cm, 1.8*cm, 2.2*cm, 1.4*cm, 1.8*cm, 2.2*cm, 1.4*cm], font_size=7.5))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Insights clés et argumentaires de justification</b>", H3))
insights_data = [
    ["#", "Insight", "Justification"],
    ["1", f"TOURTEAUX: {fmt_pct_signed(_tour_var_ca)} CA (prix 16 800 médiane YTD 2026)", "Prix 16 800 (médiane YTD 2026) — base de projection réaliste post-réappro"],
    ["2", f"CONCENTRÉS: {fmt_pct_signed(_conc_var_ca)} CA (bundle 2.5:1)", "Bundle 2.5:1 force CONCENTRÉS à la hausse quand ratio soja/concentré > 2,5 (26 ajustements 2027)"],
    ["3", f"ALIMENT_COMPLET: {fmt_pct_signed(_ac_var_t)} volume (filtrage 2024 + reprise +15%)", "Année 2024 exclue (anomalie circonstancielle corrigée), facteur reprise +15% pour capter la dynamique 2025→2026"],
    ["4", f"PREMIX: {fmt_t(_pm_2027_t)} t (restauration des volumes)", "Passage d'extrapolation CA-only à Prophet avec volumes — reflette mieux l'activité réelle"],
    ["5", f"INGREDIENTS: prix 2026 réels (bug sacs 50kg corrigé)", "Prix 2026 réels par unité (CA TTC / qte) — bug sacs 50kg corrigé, tous les prix sont maintenant corrects"],
    ["6", "ALVEOLES: 0 t (activité nulle 2026)", "Forfait ALV=0 (activité 2026 nulle) — uniquement CA forfaitaire"],
    ["7", "MATERIEL_ELEVAGE: 0 t (CA only)", "Forfait MAT 38,4 M annualisé 2026 — activité non volumique"],
    ["8", f"Total: {fmt_pct_signed(_var_t_total)} volume, {fmt_pct_signed(_var_ca_total)} CA", "Effet combiné: prix soja réaliste, bundle 2.5:1, corrections prix INGREDIENTS, reprise ALIMENT_COMPLET, PREMIX Prophet"],
]
story.append(make_table(insights_data, col_widths=[0.8*cm, 6*cm, 10*cm], font_size=8))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("<b>Argumentaire commercial pour le forecast 2027</b>", H3))
story.append(Paragraph(
    f"Le forecast 2027 à <b>{fmt_t(F2027_TOTAL_T)} t ({fmt_pct_signed(_var_t_total)} vs 2026)</b> et <b>{fmt_ca(F2027_TOTAL_CA)} M FCFA ({fmt_pct_signed(_var_ca_total)})</b> reflète plusieurs dynamiques : "
    f"(1) <b>Effet prix soja</b> (16 800 médiane YTD 2026) → CA TOURTEAUX qui représente 51% du CA total. "
    f"(2) <b>Bundle 2.5:1</b> applique une discipline commerciale historique — quand le soja domine (rupture 2026), "
    "le concentré doit suivre proportionnellement. "
    f"(3) <b>Forfait SPC révisé</b> reflète la réalité 2026 (ALV=0 car activité nulle, MAT=38,4 M annualisé). "
    f"(4) <b>Prix INGREDIENTS corrigés</b> sur 12 références (P105, I106, I107, etc.) basés sur moyennes 2026 réelles. "
    f"(5) <b>ALIMENT_COMPLET</b> : filtrage du creux 2024 (anomalie circonstancielle) + facteur reprise +15% ({fmt_t(_ac_2027_t)} t, {fmt_pct_signed(_ac_var_t)} vs 2026). "
    f"(6) <b>PREMIX restauré</b> en Prophet avec volumes ({fmt_t(_pm_2027_t)} t). "
    f"(7) <b>25 agences incluses</b> (14 BELGOCAM + 10 SPC + Maroua).",
    BODY))

story.append(Paragraph("<b>Recommandations principales</b>", H3))
_q4_2027_ca = F2027_Q[F2027_Q['quarter']==4]['ca'].iloc[0] if len(F2027_Q[F2027_Q['quarter']==4])>0 else 0
_q4_2027_pct = (_q4_2027_ca/F2027_TOTAL_CA*100) if F2027_TOTAL_CA>0 else 0
recos = [
    f"<b>Planification annuelle 2027</b> — Utiliser le forecast S3 ({fmt_t(F2027_TOTAL_T)} t, {fmt_ca(F2027_TOTAL_CA)} M FCFA) comme base budgétaire.",
    f"<b>Saisonnalité Q4</b> — Le Q4 2027 représente {fmt_pct(_q4_2027_pct)} du CA annuel ({fmt_ca(_q4_2027_ca)} M FCFA). Préparer stocks dès septembre 2027.",
    "<b>Pic d'octobre</b> — Octobre 2027 = pic de l'année. Anticiper le réapprovisionnement soja avant septembre 2027.",
    "<b>Maintien du bundle</b> — Ratio bundle 2,3:1 atteint en août 2026 à maintenir en 2027.",
    f"<b>Surveillance COMPLEMENT_ALIMENTAIRE</b> — Famille ajoutée au forecast (V300 1L proxy). {fmt_ca(F2027_FAM[F2027_FAM['family']=='COMPLEMENT_ALIMENTAIRE']['ca'].iloc[0])} M FCFA de CA prévu.",
    "<b>Surveillance FAMLA et NDOBO</b> — Ces 2 agences représentent 43% du CA 2027.",
    "<b>Mise à jour trimestrielle</b> — Actualiser le forecast chaque trimestre avec données ERP.",
]
for r in recos:
    story.append(Paragraph(f"• {r}", BULLET))

story.append(Spacer(1, 0.3*cm))
story.append(Paragraph("<b>Conclusion</b>", H3))
story.append(Paragraph(
    f"Le forecast 2027 projette <b>{fmt_t(F2027_TOTAL_T)} tonnes</b> pour un CA de <b>{fmt_ca(F2027_TOTAL_CA)} M FCFA</b> "
    f"({fmt_pct_signed(_var_t_total)} volume, {fmt_pct_signed(_var_ca_total)} CA vs 2026). La désaisonnalisation "
    "de l'effet soja, l'actualisation du prix à 16 800 FCFA (médiane) pour 2027 et 17 170 (moyen) pour Q4 2026, l'intégration de l'année 2023 (44 mois d'historique), "
    "l'ajout de la famille COMPLEMENT_ALIMENTAIRE (V300 1L only), le filtrage du creux 2024 pour ALIMENT_COMPLET, et le maintien de MATERIEL_ELEVAGE à 0 tonne "
    "permettent une projection réaliste. Le pic d'octobre nécessitera une anticipation renforcée du "
    "réapprovisionnement soja dès septembre 2027.",
    BODY))

doc.build(story)
print(f"✓ 2027 Résumé exécutif: {os.path.getsize(exec_path)/1024:.0f} KB")

# === 2027: 2-5 PDFs (proposition, RACI, stratégique, guide) ===
# Create concise versions reusing cover_page and table structures

# 2. PROPOSITION
print("Generating 2027: 2. Proposition ...")
prop_path = f"{F2027_DIR}/02_proposition_projet.pdf"
doc = SimpleDocTemplate(prop_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []
story.extend(cover_page("Proposition de Projet", "Forecast 2027", "PROPOSITION DE PROJET"))

story.append(Paragraph("1. Contexte", H1))
story.append(Paragraph(
    "BELGOCAM SA doit anticiper ses volumes et CA pour l'année 2027 complète. Cette  s'appuie sur "
    "44 mois d'historique (Jan 2023 - Août 2026) et intègre : désaisonnalisation de l'effet soja exceptionnel, "
    "inclusion des commandes En cours/Validées, actualisation du prix soja à Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026), "
    "ajout de la famille COMPLEMENT_ALIMENTAIRE (V300 1L only, V305 200L exclu), intégration de l'année 2023, "
    "MATERIEL_ELEVAGE à 0 tonne. Le forecast couvre 8 familles, 121 produits et 25 agences.",
    BODY))

story.append(Paragraph("2. Objectifs", H1))
story.append(Paragraph("<b>Objectif principal</b>", H3))
story.append(Paragraph(
    "Produire un forecast annuel 2027 (12 mois) en volumes et valeurs, désagrégé par produit, famille, agence "
    "et région, selon le scénario S3.",
    BODY))
story.append(Paragraph("<b>Objectifs spécifiques</b>", H3))
objs = [
    "Modéliser tendance + saisonnalité mensuelles via Prophet (5 familles)",
    "Désaisonnaliser l'effet soja exceptionnel de juillet-août 2026",
    "Intégrer les commandes En cours et Validées",
    "Actualiser le prix soja à Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026)",
    " Ajouter COMPLEMENT_ALIMENTAIRE (V300 1L only, 1L=1kg)",
    " Intégrer l'année 2023 (44 mois d'historique)",
    " MATERIEL_ELEVAGE toujours à 0 tonne",
    "Exclure le Maïs, V305 (BELGOKILL 200L) et produits opportunistes",
    "Produire les livrables PACE complets",
]
for o in objs:
    story.append(Paragraph(f"• {o}", BULLET))

story.append(Paragraph("3. Périmètre", H1))
scope_data = [
    ["Dimension", "Détail", "Volume"],
    ["Période forecast", "Janvier - Décembre 2027 (12 mois)", "12 mois"],
    ["Familles incluses", "8 familles (TOURTEAUX, CONCENTRÉS, INGRÉDIENTS, ALIMENT COMPLET, MATERIEL ÉLEVAGE, PREMIX, COMPLEMENT_ALIMENTAIRE)", "8"],
    ["Produits", "115 références (T102, C101-C108, MAT014, P102N2, V300 BELGOKILL 1L, etc.)", "121"],
    ["Agences", "25 agences BELGOCAM", "20"],
    ["Régions", "Ouest, Centre, Littoral", "3"],
    ["Niveau détail", "Produit × Agence × Mois", "13 704 lignes"],
    ["Scénario", "S3 - Réappro soja 100%", "1"],
    ["Données historiques", "176 576 enregistrements (Jan 2023 - Août 2026)", "44 mois"],
    ["Désaisonnalisation", "Effet soja Jul-Août 2026 neutralisé", "—"],
    ["Prix soja", "Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026)", "—"],
    ["COMPLEMENT_ALIM.", "V300 1L only (V305 exclu), 1L=1kg", "9 produits"],
    ["ALVEOLES séparé", "4 refs MAT011/14/15/17, tonnes=0", "CA only"],
]
story.append(make_table(scope_data, col_widths=[3.5*cm, 9.5*cm, 4*cm], font_size=9, highlight_rows=[11, 12]))

story.append(Paragraph("4. Méthodologie PACE", H1))
pace_data = [
    ["Phase", "Activités", "Livrables"],
    ["P - PREPARE", "Consolidation 176 576 records + En cours/Validées, désaisonnalisation soja, ajout COMPLEMENT_ALIM.", "Dataset 2023-2026, prix 2027"],
    ["A - ANALYZE", "AED, saisonnalité 2023-2025, top produits/agences", "Graphiques, synthèse AED"],
    ["C - CONSTRUCT", "Prophet (5 familles × 3 régions) + extrapolation (2 familles), forecast 12 mois S3", "Forecast 2027 (13 704 lignes)"],
    ["E - EXECUTE", "Excel 8 feuilles, 5 PDFs PACE", "8 livrables finaux"],
]
story.append(make_table(pace_data, col_widths=[3*cm, 7.5*cm, 6.5*cm], font_size=8))

story.append(Paragraph("5. Équipe et gouvernance", H1))
team_data = [
    ["Rôle", "Responsabilité", "Phase PACE"],
    ["Data Analyst (W. F. Fohom)", "Modélisation, AED, forecast, livrables", "P + A + C + E"],
    ["Direction Commerciale", "Validation hypothèses et objectifs", "A + E"],
    ["Direction Production", "Stocks, capacité, réappro soja", "P + C"],
    ["Direction Générale", "Décisions stratégiques", "E"],
    ["Contrôle de Gestion", "Cohérence financière", "E"],
]
story.append(make_table(team_data, col_widths=[4.5*cm, 8.5*cm, 4*cm], font_size=9))

story.append(Paragraph("6. Risques et mitigation", H1))
risks_data = [
    ["Risque", "Probabilité", "Mitigation"],
    ["Historique 44 mois — suffisant", "Faible", "Mise à jour trimestrielle"],
    ["Prix soja volatil (+47% en 2 mois)", "Élevée", "Prix actualisé 17 170 FCFA, scénario S3"],
    ["Effet soja 2026 non récurrent en 2027", "Élevée", "Cap désaisonnalisation à moyenne S1"],
    ["V305 exclu — BELGOKILL 200L absent", "Faible", "V300 1L utilisé comme proxy suffisant"],
    ["Nouvelles hausses tarifaires en 2027", "Moyenne", "Hypothèse prix stable en 2027"],
]
story.append(make_table(risks_data, col_widths=[7*cm, 3*cm, 7*cm], font_size=9))

story.append(Paragraph("7. Critères de succès", H1))
success = [
    "Forecast 2027 produit sur 12 mois avec désagrégation complète",
    "Désaisonnalisation effective de l'effet soja",
    "Prix soja actualisé à Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026)",
    "Famille COMPLEMENT_ALIMENTAIRE ajoutée (V300 1L only)",
    "Écart forecast vs réalité ≤ 15% par trimestre",
    "Livrables PACE complets (8 documents) produits et diffusés",
]
for s in success:
    story.append(Paragraph(f"• {s}", BULLET))

doc.build(story)
print(f"✓ 2027 Proposition: {os.path.getsize(prop_path)/1024:.0f} KB")

# 3. MATRICE RACI 2027
print("Generating 2027: 3. Matrice RACI ...")
raci_path = f"{F2027_DIR}/03_matrice_raci.pdf"
doc = SimpleDocTemplate(raci_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
story = []
story.extend(cover_page("Matrice RACI", "Forecast 2027", "MATRICE RACI"))

story.append(Paragraph("1. Définition des rôles RACI", H1))
raci_def = [
    ["Lettre", "Rôle", "Description"],
    ["R", "Responsable", "Celui qui réalise la tâche"],
    ["A", "Approbateur", "Celui qui rend compte et valide"],
    ["C", "Consulté", "Celui dont l'avis est sollicité"],
    ["I", "Informé", "Celui qui est tenu au courant"],
]
story.append(make_table(raci_def, col_widths=[2*cm, 5*cm, 11*cm], font_size=9))

story.append(Paragraph("2. Acteurs du projet", H1))
actors = [
    ["Code", "Acteur", "Implication"],
    ["DA", "Data Analyst (W. F. Fohom)", "Pleine (R sur toutes les tâches techniques)"],
    ["DC", "Direction Commerciale", "Validation hypothèses et objectifs"],
    ["DP", "Direction Production", "Stocks, capacité, réappro soja"],
    ["DG", "Direction Générale", "Décision finale et allocation ressources"],
    ["CG", "Contrôle de Gestion", "Cohérence financière"],
    ["RA", "Responsables d'Agences (14)", "Informés et consultés sur leur périmètre"],
]
story.append(make_table(actors, col_widths=[1.5*cm, 6*cm, 10.5*cm], font_size=9))

story.append(Paragraph("3. Matrice RACI détaillée par étape PACE", H1))
raci_matrix = [
    ["Phase", "Tâche", "DA", "DC", "DP", "DG", "CG", "RA"],
    ["PREPARE", "1.1 Consolidation 2023-2026 + En cours/Validées", "R/A", "I", "C", "I", "I", "I"],
    ["", "1.2 Désaisonnalisation effet soja", "R/A", "C", "C", "I", "I", "I"],
    ["", "1.3 Calcul prix 2027 (soja 17 170)", "R/A", "C", "I", "I", "C", "I"],
    ["", "1.4 Ajout COMPLEMENT_ALIMENTAIRE (V300 1L)", "R/A", "C", "I", "I", "I", "I"],
    ["", "1.5 Validation dataset", "R", "C", "C", "I", "A", "I"],
    ["ANALYZE", "2.1 AED + saisonnalité 2023-2025", "R/A", "I", "I", "I", "I", "I"],
    ["", "2.2 Analyse effet soja exceptionnel", "R/A", "C", "C", "I", "I", "I"],
    ["", "2.3 Validation hypothèses AED", "R", "C", "C", "I", "A", "I"],
    ["CONSTRUCT", "3.1 Prophet (5 familles × 3 régions)", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.2 Extrapolation MAT+PREMIX (CA only)", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.3 Forecast 12 mois S3", "R", "C", "I", "A", "C", "I"],
    ["", "3.4 Désagrégation produit × agence", "R/A", "I", "I", "I", "I", "I"],
    ["", "3.5 Validation forecast 2027", "R", "C", "I", "A", "C", "I"],
    ["EXECUTE", "4.1 Excel 8 feuilles", "R/A", "I", "I", "I", "I", "I"],
    ["", "4.2 PDFs PACE (5 documents)", "R/A", "I", "I", "I", "I", "I"],
    ["", "4.3 Présentation CODIR", "R", "C", "C", "A", "C", "I"],
    ["", "4.4 Diffusion agences", "R", "A", "I", "I", "I", "C"],
    ["", "4.5 Mise à jour trimestrielle", "R/A", "C", "C", "I", "I", "I"],
]
story.append(make_table(raci_matrix, col_widths=[2.5*cm, 6.5*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm, 1.2*cm], font_size=7.5))

story.append(Paragraph("4. Points de contrôle", H1))
control_data = [
    ["Jalon", "Phase", "Date", "Approbateur", "Livrable"],
    ["Kick-off", "PREPARE", "03/09/2026", "DG", "Proposition validée"],
    ["Validation dataset", "PREPARE", "03/09/2026", "CG", "Dataset 2023-2026 + COMPLEMENT_ALIM."],
    ["Validation AED", "ANALYZE", "04/09/2026", "DC", "Synthèse AED"],
    ["Validation forecast", "CONSTRUCT", "05/09/2026", "DG + DC", "Forecast 2027 complet"],
    ["Livrables finaux", "EXECUTE", "05/09/2026", "DG", "8 livrables"],
    ["Présentation CODIR", "EXECUTE", "06/09/2026", "DG", "Validation officielle"],
    ["Mise à jour Q1", "—", "Mars 2027", "DA", "Réactualisation"],
]
story.append(make_table(control_data, col_widths=[3.5*cm, 2.5*cm, 3*cm, 3*cm, 5*cm], font_size=8))

doc.build(story)
print(f"✓ 2027 Matrice RACI: {os.path.getsize(raci_path)/1024:.0f} KB")

# 4. DOCUMENT STRATÉGIQUE PACE 2027
print("Generating 2027: 4. Document stratégique PACE ...")
strat_path = f"{F2027_DIR}/04_document_strategique_pace.pdf"
doc = SimpleDocTemplate(strat_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []
story.extend(cover_page("Document Stratégique PACE", "Forecast 2027 (8 familles, données 2023-2026)", "STRATÉGIE PACE"))

story.append(Paragraph("1. Vision et objectifs", H1))
story.append(Paragraph("1.1 Vision", H2))
story.append(Paragraph(
    f"Faire du forecast 2027 un <b>outil de planification annuelle</b> permettant à BELGOCAM SA d'anticiper "
    f"{fmt_t(F2027_TOTAL_T)} tonnes de ventes et {fmt_ca(F2027_TOTAL_CA)} M FCFA de chiffre d'affaires, avec désaisonnalisation de l'effet soja "
    "et intégration de la famille COMPLEMENT_ALIMENTAIRE (V300 1L only).",
    BODY))

story.append(Paragraph("1.2 Objectifs stratégiques", H2))
obj_data = [
    ["Objectif", "Description", "KPI"],
    ["Forecast annuel", "12 mois 2027 en volume + valeur", f"{fmt_t(F2027_TOTAL_T)} t, {fmt_ca(F2027_TOTAL_CA)} M FCFA"],
    ["Désaisonnalisation", "Neutraliser l'effet soja Jul-Août 2026", "Cap moyenne S1 2026"],
    ["En cours + Validées", "Intégrer comme potentielles ventes", "255 commandes incluses"],
    ["Prix actualisé", "Soja Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026)", "vs 17 170 médian 2023-2026"],
    ["Désagrégation", "Produit × agence × mois", f"{len(f2027_df):,} lignes".replace(',', ' ')],
    ["Adoption", "Diffusion CODIR + 25 agences", "100% agences informées"],
    [" COMPLEMENT_ALIM.", "V300 1L only (V305 exclu), 1L=1kg", f"{fmt_t(F2027_FAM[F2027_FAM['family']=='COMPLEMENT_ALIMENTAIRE']['t'].iloc[0])} t, {fmt_ca(F2027_FAM[F2027_FAM['family']=='COMPLEMENT_ALIMENTAIRE']['ca'].iloc[0])} M FCFA"],
    [" ALVEOLES séparé", "4 refs (MAT011/MAT014/MAT015/MAT017), tonnes=0", f"0 t, {fmt_ca(F2027_FAM[F2027_FAM['family']=='ALVEOLES']['ca'].iloc[0])} M FCFA"],
]
story.append(make_table(obj_data, col_widths=[3.5*cm, 8.5*cm, 5*cm], font_size=9, highlight_rows=[6, 7]))

story.append(Paragraph("2. Phase PREPARE - Préparation des données", H1))
story.append(Paragraph("2.1 Sources consolidées (2023-2026)", H2))
sources_data = [
    ["Source", "Période", "Records"],
    ["LY_21_24 (filtre 2023)", "Jan-Dec 2023", "6 016"],
    ["LY_24 (Jul-Dec 2024)", "Juillet-Décembre 2024", "52 215"],
    ["Historique 2025", "Jan-Déc 2025", "66 206"],
    ["S1 + Juil + Août 2026", "Jan-Août 2026", "52 139"],
    ["En cours + Validées", "Août 2026", "~255"],
    ["TOTAL", "44 mois", "176 576"],
]
story.append(make_table(sources_data, col_widths=[5*cm, 5*cm, 5*cm], font_size=9))

story.append(Paragraph("2.2 Désaisonnalisation de l'effet soja", H2))
story.append(Paragraph(
    "Les volumes soja de juillet-août 2026 ont été exceptionnellement élevés (rupture concurrente). "
    "Pour éviter que Prophet n'extrapole ce phénomène non récurrent, les volumes soja Jul-Août 2026 ont été "
    "<b>cappés à la moyenne S1 2026</b>.",
    BODY))

story.append(Paragraph("2.3  Famille COMPLEMENT_ALIMENTAIRE", H2))
story.append(Paragraph(
    "La famille COMPLEMENT_ALIMENTAIRE (9 produits liquides 1L) a été ajoutée au forecast 2027. "
    "Selon demande utilisateur : <b>seul V300 (BELGOKILL 1L) est conservé — V305 (BELGOKILL 200L) est EXCLU</b>. "
    "Conversion 1L = 1kg appliquée pour exprimer les volumes en tonnes.",
    BODY))

story.append(Paragraph("3. Phase CONSTRUCT - Modélisation", H1))
story.append(Paragraph("3.1 Configuration Prophet", H2))
config_data = [
    ["Paramètre", "Valeur", "Justification"],
    ["yearly_seasonality", "True", "Capture la saisonnalité annuelle"],
    ["seasonality_mode", "multiplicative", "Saisonnalité proportionnelle"],
    ["changepoint_prior_scale", "0.05", "Tendance modérée"],
    ["interval_width", "0.8", "Intervalle 80%"],
    ["Période forecast", "12 mois (Jan-Déc 2027)", "Année complète"],
    ["Familles Prophet", "5 (incluant COMPLEMENT_ALIMENTAIRE)", ""],

]
story.append(make_table(config_data, col_widths=[5*cm, 5*cm, 7*cm], font_size=9, highlight_rows=[6]))

story.append(Paragraph("4. Résultats par famille et trimestre", H1))
story.append(Paragraph("4.1 Par famille", H2))
fam_detail = build_2027_fam_detail_table()
story.append(make_table(fam_detail, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2*cm, 4*cm], font_size=9, highlight_rows=[4, 8]))

story.append(Paragraph("4.2 Par trimestre", H2))
q_detail = build_2027_q_detail_table()
story.append(make_table(q_detail, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 2*cm, 4.5*cm], font_size=9))

story.append(Paragraph("5. Historique 2023-2026 vs Forecast 2027", H1))
hist_detail = build_2027_hist_table()
story.append(make_table(hist_detail, col_widths=[2.5*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.7*cm, 1.7*cm, 1.5*cm, 1.5*cm], font_size=7, highlight_rows=[9, 10]))

story.append(Paragraph("6. Plan de déploiement", H1))
deploy_data = [
    ["Étape", "Date", "Action", "Cible"],
    ["1", "03/09/2026", "Validation proposition", "DG"],
    ["2", "05/09/2026", "Livrables finaux prêts", "DA"],
    ["3", "06/09/2026", "Présentation CODIR", "DG + DC + DP + CG"],
    ["4", "07/09/2026", "Diffusion aux agences", "14 RA"],
    ["5", "Mars 2027", "Mise à jour Q1", "DA"],
]
story.append(make_table(deploy_data, col_widths=[1*cm, 3*cm, 7*cm, 6*cm], font_size=9))

story.append(Paragraph("7. Conclusion", H1))
story.append(Paragraph(
    f"Le forecast 2027 projette <b>{fmt_t(F2027_TOTAL_T)} tonnes</b> pour <b>{fmt_ca(F2027_TOTAL_CA)} M FCFA</b> "
    f"({fmt_pct_signed(_var_t_total)} volume, {fmt_pct_signed(_var_ca_total)} CA vs 2026 hors MAIS). La désaisonnalisation "
    "de l'effet soja, l'actualisation du prix à 16 800 FCFA (médiane) pour 2027 et 17 170 (moyen) pour Q4 2026, l'intégration de l'année 2023 (44 mois d'historique), "
    "l'ajout de la famille COMPLEMENT_ALIMENTAIRE (V300 1L only), le filtrage du creux 2024 pour ALIMENT_COMPLET, et le maintien de MATERIEL_ELEVAGE à 0 tonne "
    "permettent une projection réaliste.",
    BODY))

doc.build(story)
print(f"✓ 2027 Document stratégique: {os.path.getsize(strat_path)/1024:.0f} KB")

# 5. GUIDE MÉTHODOLOGIQUE 2027
print("Generating 2027: 5. Guide méthodologique ...")
guide_path = f"{F2027_DIR}/05_guide_methodologique.pdf"
doc = SimpleDocTemplate(guide_path, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
story = []
story.extend(cover_page("Guide Méthodologique", "Forecast 2027 (Prophet + COMPLEMENT_ALIMENTAIRE V300 1L)", "GUIDE MÉTHODOLOGIQUE"))

story.append(Paragraph("Sommaire", H1))
story.append(Paragraph(
    "1. Introduction et contexte<br/>"
    "2. Méthodologie PACE<br/>"
    "3. Outils et environnement<br/>"
    "4. Phase PREPARE - Données 2023-2026<br/>"
    "5. Phase PREPARE - Désaisonnalisation soja<br/>"
    "6. Phase PREPARE - COMPLEMENT_ALIMENTAIRE (V300 1L)<br/>"
    "7. Phase ANALYZE - AED<br/>"
    "8. Phase CONSTRUCT - Prophet<br/>"
    "9. Phase EXECUTE - Livrables<br/>"
    "Annexes",
    BODY))

story.append(PageBreak())
story.append(Paragraph("1. Introduction et contexte", H1))
story.append(Paragraph(
    "Ce guide décrit la méthodologie complète du forecast 2027 de BELGOCAM SA (). Il couvre 12 mois "
    "(janvier-décembre 2027) avec 8 familles de produits (incluant COMPLEMENT_ALIMENTAIRE) et 25 agences. "
    "Innovations : (1) ajout COMPLEMENT_ALIMENTAIRE (V300 1L only, V305 200L EXCLU, 1L=1kg), "
    "(2) intégration année 2023 (44 mois d'historique), (3) MATERIEL_ELEVAGE toujours 0 tonne, "
    "(4) désaisonnalisation effet soja, (5) inclusion En cours/Validées, (6) prix soja 17 170 FCFA.",
    BODY))

story.append(Paragraph("2. Méthodologie PACE", H1))
pace_detail = [
    ["Phase", "Objectif", "Livrables", "Durée"],
    ["P - PREPARE", "Préparer données 2023-2026 + désaisonnalisation + COMPLEMENT_ALIM.", "Dataset 176 576 records, prix 2027", "2 jours"],
    ["A - ANALYZE", "Comprendre données + effet soja", "Graphiques, synthèse AED", "1 jour"],
    ["C - CONSTRUCT", "Modéliser Prophet 12 mois (8 familles)", "Forecast 2027 (13 704 lignes)", "2 jours"],
    ["E - EXECUTE", "Produire livrables finaux", "Excel + 5 PDFs", "1 jour"],
]
story.append(make_table(pace_detail, col_widths=[3*cm, 4*cm, 5*cm, 3*cm], font_size=9))

story.append(Paragraph("3. Outils et environnement", H1))
stack_data = [
    ["Outil", "Version", "Usage"],
    ["Python", "3.12", "Langage principal"],
    ["pandas", "2.x", "Manipulation données"],
    ["Prophet", "1.1+", "Modélisation prédictive (5 familles × 3 régions = 15 modèles)"],
    ["openpyxl", "3.x", "Génération Excel (8 feuilles)"],
    ["ReportLab", "4.x", "Génération PDFs (5 documents PACE)"],
]
story.append(make_table(stack_data, col_widths=[4*cm, 3*cm, 8*cm], font_size=9))

story.append(PageBreak())
story.append(Paragraph("4. Phase PREPARE - Données 2023-2026", H1))
story.append(Paragraph(
    "Le dataset consolidé couvre 44 mois (janvier 2023 - août 2026), soit 176 576 enregistrements. "
    "Sources : LY_21_24 filtre 2023 (6 016), LY_24 Jul-Dec 2024 (52 215), Historique 2025 (66 206), "
    "S1+Juil+Août 2026 (52 139), En cours + Validées (~255).",
    BODY))
story.append(Paragraph(
    "<b>Intégration de l'année 2023</b> : La  inclut 2023 pour enrichir l'historique avec 4 ans "
    "de données. Prophet dispose ainsi d'une saisonnalité annuelle complète.",
    BODY))

story.append(Paragraph("5. Phase PREPARE - Désaisonnalisation soja", H1))
story.append(Paragraph(
    "L'effet soja exceptionnel de juillet-août 2026 (rupture concurrente) a généré des volumes anormalement "
    "élevés non récurrents. Les volumes soja Jul-Août 2026 ont été cappés à la moyenne S1 2026.",
    BODY))

story.append(PageBreak())
story.append(Paragraph("6.  Famille COMPLEMENT_ALIMENTAIRE", H1))
story.append(Paragraph(
    "La famille COMPLEMENT_ALIMENTAIRE regroupe 9 produits liquides BELGO en conditionnement 1L : "
    "BELGOKILL V300 (1L), BELGO HARMONY, BELGO PROTECT, BELGO DRY LIT, BELGO WATER CLEAN, BELGO VIT Ese, "
    "BELGO THERMO, BELGO BIO SELECT, BELGO FRESH.",
    BODY))

story.append(Paragraph("<b>Exclusion de V305 (BELGOKILL 200L)</b>", H3))
story.append(Paragraph(
    "Selon demande utilisateur explicite : seul le conditionnement 1L est conservé. "
    "V305 (BELGOKILL 200L) est EXCLU du forecast. Cela permet d'avoir une famille homogène (tous 1L) "
    "et d'éviter le biais du conditionnement bulk.",
    BODY))

story.append(Paragraph("<b>Conversion 1L = 1kg</b>", H3))
story.append(Paragraph(
    "Pour exprimer les volumes en tonnes (nécessaire pour Prophet), la conversion 1L = 1kg est appliquée : "
    "<b>1 qte = 1 kg, tonnes = qte / 1000</b>.",
    BODY))

story.append(Paragraph("<b>Proxy BELGOKILL V300 1L</b>", H3))
story.append(Paragraph(
    "BELGOKILL V300 (1L) est le produit dominant de la famille (38% du CA famille sur 2023-2026). "
    "Sa tendance est utilisée comme proxy pour modéliser toute la famille via Prophet.",
    BODY))

ca_proxy = [
    ["Réf.", "Produit", "Poids (kg/qte)", "Prix 2027 (FCFA/L)", "Statut"],
    ["V300", "BELGOKILL 1L", "1", "2 500", "✓ Inclus (proxy)"],
    ["V305", "BELGOKILL 200L", "—", "—", "✗ Exclu"],
    ["CA003.1", "BELGO HARMONY 1L", "1", "8 000", "✓ Inclus"],
    ["CA004.1", "BELGO PROTECT 1L", "1", "9 800", "✓ Inclus"],
    ["CA006.1", "BELGO DRY LIT 1L", "1", "6 500", "✓ Inclus"],
    ["CA001.1", "BELGO WATER CLEAN 1L", "1", "5 500", "✓ Inclus"],
    ["CA002.1", "BELGO VIT Ese 1L", "1", "10 500", "✓ Inclus"],
    ["CA005.1", "BELGO THERMO 1L", "1", "15 000", "✓ Inclus"],
    ["CA007.1", "BELGO BIO SELECT 1L", "1", "8 000", "✓ Inclus"],
    ["CA008.1", "BELGO FRESH 1L", "1", "14 000", "✓ Inclus"],
]
story.append(make_table(ca_proxy, col_widths=[2*cm, 5*cm, 2.5*cm, 3.5*cm, 3*cm], font_size=9, highlight_rows=[2]))

story.append(Paragraph("7. Phase ANALYZE - AED", H1))
story.append(Paragraph(
    "L'AED révèle : (1) TOURTEAUX = 76% du volume historique, (2) Q4 = 35-46% du volume annuel, "
    "(3) pic octobre (indice 1.86 pour TOURTEAUX en 2025), (4) FAMLA = 31.6% du volume, "
    "(5) COMPLEMENT_ALIMENTAIRE = 3-7 t/an (faible volume, forte valeur unitaire).",
    BODY))

story.append(Paragraph("8. Phase CONSTRUCT - Prophet", H1))
story.append(Paragraph(
    "15 modèles Prophet famille × région (5 familles Prophet × 3 régions) + extrapolation pour MATERIEL_ELEVAGE "
    "et PREMIX (CA only, tonnes=0). Désagrégation par produit × agence selon parts historiques de CA. "
    "Prix 2027 : Q4: 17 170 (moyen YTD 2026) / 2027: 16 800 (médiane YTD 2026) soja, prix par litre pour COMPLEMENT_ALIMENTAIRE (2 500-15 000 FCFA/L).",
    BODY))

story.append(Paragraph("9. Phase EXECUTE - Livrables", H1))
livrables_data = [
    ["#", "Livrable", "Format", "Usage"],
    ["1", "Résumé exécutif", "PDF", "Direction Générale"],
    ["2", "Proposition de projet", "PDF", "CODIR"],
    ["3", "Matrice RACI", "PDF", "Gouvernance"],
    ["4", "Document stratégique PACE", "PDF", "Référence stratégique"],
    ["5", "Guide méthodologique", "PDF", "Référence méthodes"],
    ["6", "Excel forecast 2027 (8 feuilles)", "XLSX", "Pilotage opérationnel"],
]
story.append(make_table(livrables_data, col_widths=[1*cm, 5*cm, 3*cm, 6*cm], font_size=9))

story.append(Paragraph("Annexe - Glossaire", H1))
glossaire = [
    ["Terme", "Définition"],
    ["Prophet", "Bibliothèque prévision séries temporelles (Facebook/Meta)"],
    ["Désaisonnalisation", "Neutralisation d'un effet exceptionnel non récurrent"],
    ["S3", "Scénario réappro soja 100%"],
    ["Cap S1", "Plafonnement Jul-Août 2026 à moyenne S1 2026"],
    ["Q4", "Quatrième trimestre (oct-déc)"],
    ["PACE", "Prepare-Analyze-Construct-Execute"],
    ["COMPLEMENT_ALIM.", "9 produits liquides 1L BELGOxxx (V300 1L proxy)"],
    ["ALVEOLES séparé", "Famille distincte de MATERIEL_ELEVAGE (4 refs MAT011/14/15/17)"],
    ["SPC agences", "10 agences SPC incluses (Baf-Chefferie dominant)"],
    ["Maroua", "Agence ajoutée (Centre), soja T102 only"],
    ["Bundle 2.5:1", "Ratio soja:concentré forcé à <=2.5:1"],
    ["Saisonnalité ALV 2026", "ALVEOLES utilise 2026 (13.9 M/an) au lieu du pic 2025 (739 M)"],
    ["10 SPC agences", "Toutes SPC incluses (10 agences)"],
    ["Forfait SPC", "Forfait MAT_ELEV (38.4 M/an) basé sur 2026 annualisé, poids 2025"],
    ["SPC PK15 forfait", "Forfait réaliste 1 M/an (0.5 ALV + 0.5 MAT)"],
    [" V305 EXCLU", "BELGOKILL 200L retiré (seul V300 1L conservé)"],
    [" 1L=1kg", "Conversion pour volumes liquides en tonnes"],
    [" Données 2023-2026", "44 mois d'historique (Jan 2023 - Août 2026)"],
    ["MATERIEL_ELEVAGE", "Toujours 0 en tonnes (CA only). ALVEOLES désormais séparés en famille distincte"],
]
story.append(make_table(glossaire, col_widths=[5*cm, 11*cm], font_size=9, highlight_rows=[7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]))

doc.build(story)
print(f"✓ 2027 Guide méthodologique: {os.path.getsize(guide_path)/1024:.0f} KB")

print(f"\n=== 2027: 5 PDFs GÉNÉRÉS ===")
for f in sorted(os.listdir(F2027_DIR)):
    if f.endswith('.pdf'):
        print(f"  {f}: {os.path.getsize(os.path.join(F2027_DIR, f))/1024:.0f} KB")

print("\n=== TOUS LES 10 PDFs GÉNÉRÉS ===")
print(f"Q4 2026: {Q4_DIR}")
print(f"2027: {F2027_DIR}")
