"""
Ajoute une section 'Réalisation 2026 YTD + Projection fin d'année' dans les PDFs Q4 2026 et 2027.
Modifie les scripts source et régénère les PDFs.
"""
import os
import re

# === Définition de la section à insérer ===
# Cette section sera insérée avant "Conclusion" ou "Recommandations principales" dans chaque PDF

YTD_SECTION_RESUME = '''
story.append(Paragraph("<b>Réalisation 2026 YTD + Projection fin d\\'année</b>", H3))

story.append(Paragraph(
    "Le tableau ci-dessous présente la réalisation YTD (Jan-Août 2026, tous clients) "
    "vs objectifs, et la projection de fin d\\'année (YTD réel + Q4 forecast S3) vs objectif annuel.",
    BODY))

ytd_data = [
    ["Famille", "YTD réel (t)", "Obj YTD (t)", "% YTD", "Q4 fcst (t)", "Total 2026 (t)", "Obj annuel (t)", "% Annuel", "Statut"],
    ["TOURTEAUX", "39 718", "34 359", "116%", "33 853", "73 571", "54 196", "136%", "\\u2705"],
    ["CONCENTR\\u00c9S", "12 066", "15 792", "76%", "6 393", "18 459", "24 192", "76%", "\\u274c"],
    ["ALIMENT COMPLET", "525", "663", "79%", "368", "893", "1 046", "85%", "\\u274c"],
    ["INGR\\u00c9DIENTS", "548", "728", "75%", "307", "855", "1 064", "80%", "\\u274c"],
    ["PREMIX", "61", "80", "76%", "0", "61", "130", "47%", "\\u274c"],
    ["TOTAL", "52 918", "51 622", "103%", "40 921", "93 839", "80 628", "116%", "\\u2705"],
]
story.append(make_table(ytd_data, col_widths=[2.8*cm, 1.8*cm, 1.8*cm, 1.2*cm, 1.8*cm, 2*cm, 2*cm, 1.2*cm, 0.8*cm], font_size=7.5))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph(
    "<b>Lecture cl\\u00e9</b> : Le TOURTEAUX surperforme massivement (136% de l\\'objectif annuel) gr\\u00e2ce \\u00e0 la rupture "
    "concurrente de juillet-ao\\u00fbt. Les CONCENTR\\u00c9S restent \\u00e0 76% malgr\\u00e9 un ao\\u00fbt \\u00e0 113%. "
    "Au niveau global, l\\'objectif annuel est d\\u00e9pass\\u00e9 de 116% (+13 211 t), mais cette performance est enti\\u00e8rement "
    "tir\\u00e9e par le TOURTEAUX (effet exceptionnel non r\\u00e9current).",
    BODY))
story.append(Spacer(1, 0.3*cm))
'''

YTD_SECTION_STRATEGIE = '''
story.append(PageBreak())
story.append(Paragraph("6.5 R\\u00e9alisation 2026 YTD + Projection fin d\\'ann\\u00e9e", H1))
story.append(Paragraph(
    "Cette section pr\\u00e9sente la r\\u00e9alisation YTD (Jan-Ao\\u00fbt 2026, tous clients) compar\\u00e9e aux objectifs, "
    "ainsi que la projection de fin d\\'ann\\u00e9e (YTD r\\u00e9el + Q4 forecast S3) vs objectif annuel. "
    "Cette analyse permet de contextualiser le forecast Q4 dans la trajectoire annuelle compl\\u00e8te.",
    BODY))

story.append(Paragraph("6.5.1 Volumes YTD vs objectifs", H2))
ytd_vol_data = [
    ["Famille", "YTD r\\u00e9el (t)", "Obj YTD (t)", "% YTD", "\\u00c9cart (t)", "Statut"],
    ["TOURTEAUX", "39 718", "34 359", "116%", "+5 359", "\\u2705"],
    ["CONCENTR\\u00c9S", "12 066", "15 792", "76%", "-3 726", "\\u274c"],
    ["ALIMENT COMPLET", "525", "663", "79%", "-138", "\\u274c"],
    ["INGR\\u00c9DIENTS", "548", "728", "75%", "-180", "\\u274c"],
    ["PREMIX", "61", "80", "76%", "-19", "\\u274c"],
    ["TOTAL", "52 918", "51 622", "103%", "+1 296", "\\u2705"],
]
story.append(make_table(ytd_vol_data, col_widths=[3.5*cm, 2.5*cm, 2.5*cm, 1.5*cm, 2.5*cm, 1.5*cm], font_size=9))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("6.5.2 CA YTD vs objectifs", H2))
ytd_ca_data = [
    ["Famille", "CA YTD r\\u00e9el (M FCFA)", "CA YTD obj (M FCFA)", "% YTD"],
    ["TOURTEAUX", "13 478", "11 563", "117%"],
    ["CONCENTR\\u00c9S", "8 069", "10 970", "74%"],
    ["ALIMENT COMPLET", "427", "576", "74%"],
    ["INGR\\u00c9DIENTS", "777", "1 074", "72%"],
    ["PREMIX", "115", "149", "77%"],
    ["TOTAL", "22 866", "24 332", "94%"],
]
story.append(make_table(ytd_ca_data, col_widths=[4*cm, 4*cm, 4*cm, 2.5*cm], font_size=9))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph("6.5.3 Projection fin d\\'ann\\u00e9e (YTD + Q4 forecast S3)", H2))
proj_data = [
    ["Famille", "YTD r\\u00e9el (t)", "Q4 forecast (t)", "Total 2026 (t)", "Obj annuel (t)", "% Annuel", "\\u00c9cart (t)"],
    ["TOURTEAUX", "39 718", "33 853", "73 571", "54 196", "136%", "+19 375"],
    ["CONCENTR\\u00c9S", "12 066", "6 393", "18 459", "24 192", "76%", "-5 733"],
    ["ALIMENT COMPLET", "525", "368", "893", "1 046", "85%", "-153"],
    ["INGR\\u00c9DIENTS", "548", "307", "855", "1 064", "80%", "-209"],
    ["PREMIX", "61", "0", "61", "130", "47%", "-69"],
    ["TOTAL", "52 918", "40 921", "93 839", "80 628", "116%", "+13 211"],
]
story.append(make_table(proj_data, col_widths=[3*cm, 2*cm, 2.5*cm, 2.5*cm, 2.5*cm, 1.5*cm, 2.5*cm], font_size=8))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph(
    "<b>Analyse</b> : Le TOURTEAUX surperforme massivement (136% de l\\'objectif annuel, +19 375 t) "
    "gr\\u00e2ce \\u00e0 la rupture concurrente de juillet-ao\\u00fbt 2026. Cette surperformance est exceptionnelle et non r\\u00e9currente. "
    "Les CONCENTR\\u00c9S restent \\u00e0 76% de l\\'objectif annuel (-5 733 t) : malgr\\u00e9 un mois d\\'ao\\u00fbt \\u00e0 113%, "
    "le retard accumul\\u00e9 sur S1 (-3 726 t) n\\'est pas rattrapable en Q4. "
    "Au niveau global, l\\'objectif annuel est d\\u00e9pass\\u00e9 \\u00e0 116% (+13 211 t), mais cette performance est enti\\u00e8rement tir\\u00e9e "
    "par le TOURTEAUX. Sans cet effet exceptionnel, l\\'objectif serait manqu\\u00e9. "
    "Cette analyse justifie la d\\u00e9saisonnalisation de l\\'effet soja dans le forecast 2027, pour \\u00e9viter de projeter "
    "un ph\\u00e9nom\\u00e8ne non r\\u00e9current.",
    BODY))
'''

print("Section definitions created")

# === 1. Insert into pace_07_pdf_batch1.py (Résumé exécutif Q4 2026) ===
# Insert before "Recommandations principales"
print("\n1. Inserting YTD section in résumé exécutif Q4 2026...")
with open('scripts/pace_07_pdf_batch1.py', 'r') as f:
    content = f.read()

# Insert before 'story.append(Paragraph("<b>Recommandations principales</b>", H3))'
insertion_point = 'story.append(Paragraph("<b>Recommandations principales</b>", H3))'
ytd_section_exec = YTD_SECTION_RESUME.replace('\\u2705', '\u2705').replace('\\u274c', '\u274c').replace('\\u00e9', '\u00e9').replace('\\u00fb', '\u00fb').replace('\\u00e8', '\u00e8').replace("\\'", "'")
# Fix the escaped quotes
ytd_section_exec = ytd_section_exec.replace("\\'", "'")

content = content.replace(insertion_point, ytd_section_exec + '\n' + insertion_point)

with open('scripts/pace_07_pdf_batch1.py', 'w') as f:
    f.write(content)
print("  ✅ Résumé exécutif Q4 2026 updated")

# === 2. Insert into pace_08_pdf_batch2.py (Document stratégique Q4 2026) ===
print("\n2. Inserting YTD section in document stratégique Q4 2026...")
with open('scripts/pace_08_pdf_batch2.py', 'r') as f:
    content = f.read()

# Insert before "7. Plan de déploiement et adoption"
insertion_point = 'story.append(Paragraph("7. Plan de déploiement et adoption", H1))'
ytd_section_strat = YTD_SECTION_STRATEGIE.replace('\\u2705', '\u2705').replace('\\u274c', '\u274c').replace('\\u00e9', '\u00e9').replace('\\u00fb', '\u00fb').replace('\\u00e8', '\u00e8').replace("\\'", "'")

content = content.replace(insertion_point, ytd_section_strat + '\n' + insertion_point)

with open('scripts/pace_08_pdf_batch2.py', 'w') as f:
    f.write(content)
print("  ✅ Document stratégique Q4 2026 updated")

# === 3. Insert into forecast_2027_pdfs.py (Résumé exécutif 2027) ===
print("\n3. Inserting YTD section in résumé exécutif 2027...")
with open('scripts/forecast_2027_pdfs.py', 'r') as f:
    content = f.read()

insertion_point = 'story.append(Paragraph("<b>Recommandations principales</b>", H3))'
# For 2027, add context about the 2026 YTD as baseline
ytd_section_2027 = '''
story.append(Paragraph("<b>Réalisation 2026 YTD + Projection fin d'année (contexte)</b>", H3))

story.append(Paragraph(
    "Le forecast 2027 s'appuie sur la réalisation 2026. Le tableau ci-dessous présente la trajectoire 2026 "
    "(YTD réel Jan-Août + Q4 forecast S3) vs objectif annuel, qui sert de base de comparaison.",
    BODY))

ytd_data = [
    ["Famille", "YTD 2026 (t)", "Q4 fcst (t)", "Total 2026 (t)", "Obj annuel (t)", "% Annuel", "Statut"],
    ["TOURTEAUX", "39 718", "33 853", "73 571", "54 196", "136%", "\u2705"],
    ["CONCENTR\u00c9S", "12 066", "6 393", "18 459", "24 192", "76%", "\u274c"],
    ["ALIMENT COMPLET", "525", "368", "893", "1 046", "85%", "\u274c"],
    ["INGR\u00c9DIENTS", "548", "307", "855", "1 064", "80%", "\u274c"],
    ["PREMIX", "61", "0", "61", "130", "47%", "\u274c"],
    ["TOTAL", "52 918", "40 921", "93 839", "80 628", "116%", "\u2705"],
]
story.append(make_table(ytd_data, col_widths=[2.8*cm, 2*cm, 1.8*cm, 2.2*cm, 2.5*cm, 1.5*cm, 1*cm], font_size=8))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph(
    "<b>Lecture</b> : En 2026, le TOURTEAUX surperforme à 136% (effet rupture concurrente exceptionnel) "
    "tandis que les CONCENTR\u00c9S restent à 76% de l'objectif. Le forecast 2027 projette une normalisation : "
    "109 288 t (vs 93 839 t en 2026, +16%) et 58 550 M FCFA. La désaisonnalisation de l'effet soja 2026 "
    "évite de répliquer ce phénomène non récurrent dans les prédictions 2027.",
    BODY))
story.append(Spacer(1, 0.3*cm))
'''

content = content.replace(insertion_point, ytd_section_2027 + '\n' + insertion_point)

with open('scripts/forecast_2027_pdfs.py', 'w') as f:
    f.write(content)
print("  ✅ Résumé exécutif 2027 updated")

# === 4. Insert into forecast_2027_pdfs.py (Document stratégique 2027) ===
print("\n4. Inserting YTD section in document stratégique 2027...")
with open('scripts/forecast_2027_pdfs.py', 'r') as f:
    content = f.read()

# Find the "5. Plan de déploiement" section in the 2027 stratégique
insertion_point = 'story.append(Paragraph("5. Plan de déploiement", H1))'
ytd_section_strat_2027 = '''
story.append(PageBreak())
story.append(Paragraph("4.5 Réalisation 2026 YTD + Projection fin d'année (contexte)", H1))
story.append(Paragraph(
    "Cette section présente la trajectoire 2026 (YTD réel Jan-Août + Q4 forecast S3) qui sert de base "
    "de comparaison au forecast 2027. L'analyse révèle une surperformance exceptionnelle du TOURTEAUX "
    "(136% de l'objectif annuel) due à la rupture concurrente, et un déficit structurel des CONCENTRÉS (76%).",
    BODY))

proj_data = [
    ["Famille", "YTD réel (t)", "Q4 fcst (t)", "Total 2026 (t)", "Obj annuel (t)", "% Annuel", "Écart (t)"],
    ["TOURTEAUX", "39 718", "33 853", "73 571", "54 196", "136%", "+19 375"],
    ["CONCENTRÉS", "12 066", "6 393", "18 459", "24 192", "76%", "-5 733"],
    ["ALIMENT COMPLET", "525", "368", "893", "1 046", "85%", "-153"],
    ["INGRÉDIENTS", "548", "307", "855", "1 064", "80%", "-209"],
    ["PREMIX", "61", "0", "61", "130", "47%", "-69"],
    ["TOTAL", "52 918", "40 921", "93 839", "80 628", "116%", "+13 211"],
]
story.append(make_table(proj_data, col_widths=[3*cm, 2*cm, 2.5*cm, 2.5*cm, 2.5*cm, 1.5*cm, 2.5*cm], font_size=8))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph(
    "<b>Analyse</b> : Le TOURTEAUX surperforme massivement (+19 375 t) grâce à la rupture concurrente "
    "de juillet-août 2026. Cette surperformance est exceptionnelle et non récurrente. Les CONCENTRÉS restent "
    "à 76% de l'objectif annuel (-5 733 t) : malgré un mois d'août à 113%, le retard accumulé sur S1 n'est pas "
    "rattrapable en Q4. Au niveau global, l'objectif annuel est dépassé à 116% (+13 211 t), mais cette performance "
    "est entièrement tirée par le TOURTEAUX. Cette analyse justifie la désaisonnalisation de l'effet soja "
    "dans le forecast 2027, pour éviter de projeter un phénomène non récurrent.",
    BODY))

'''
content = content.replace(insertion_point, ytd_section_strat_2027 + '\n' + insertion_point)

with open('scripts/forecast_2027_pdfs.py', 'w') as f:
    f.write(content)
print("  ✅ Document stratégique 2027 updated")

print("\n=== All 4 scripts updated. Now regenerating PDFs... ===")
