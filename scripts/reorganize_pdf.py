"""
Reorganize the PDF script:
1. Move sections 9 + 9bis (Ventes vs Objectifs + Focus CONCENTRÉS) to position 3-4
2. Shift old sections 3-8 to 5-10
3. Renumber all sections
4. Replace "réel" when referring to objectives
5. Update sommaire
"""
import re

with open('/home/z/my-project/scripts/build_pdf_report_v2.py', 'r', encoding='utf-8') as f:
    content = f.read()
lines = content.split('\n')

# Section boundaries (0-indexed line numbers)
# Format: (start_0indexed, end_0indexed_exclusive)
SECTIONS = {
    'header': (0, 384),           # Everything before Sommaire (imports, styles, cover, etc.)
    'sommaire': (384, 424),       # Sommaire
    'synthese': (424, 479),       # Synthèse exécutive
    'sec1': (479, 561),           # 1. État des lieux
    'sec2': (561, 618),           # 2. Chiffres clés
    'sec3_zero': (618, 717),      # 3. Zéro Achat (old)
    'sec4_pertes': (717, 828),    # 4. Pertes Q1 (old)
    'sec5_trans': (828, 958),     # 5. Transition Q1→Q2 (old)
    'sec6_chick': (958, 1055),    # 6. Chick & Piglet (old)
    'sec7_persp': (1055, 1161),   # 7. Perspectives (old)
    'sec8_proj': (1161, 1213),    # 8. Projection CA (old)
    'sec9_ventes': (1213, 1547),  # 9. Ventes vs Objectifs (old)
    'sec9bis_conc': (1547, 1687), # 9bis. Focus CONCENTRÉS (old)
    'sec10_forecast': (1687, 1772), # 10. Forecast (old)
    'sec10bis_yoy': (1772, 1990),   # 10bis. YoY (old)
    'sec11_reco': (1990, 2113),    # 11. Recommandations (old)
    'sec12_concl': (2113, len(lines)), # 12. Conclusion
}

# Extract each section's text
def get_section(name):
    s, e = SECTIONS[name]
    return '\n'.join(lines[s:e])

# Build new order


def renumber_section(text, old_num, new_num):
    """Renumber a section: replace section numbers in titles and references."""
    # Replace section titles like "3. Title" → "5. Title"
    # But be careful not to replace numbers inside data or random text
    # Strategy: only replace patterns like "X.Y" or "X." at start of section titles
    
    # Replace H1 titles: "old. Title" → "new. Title"
    text = re.sub(rf'"{old_num}\. ', f'"{new_num}. ', text)
    text = re.sub(rf'"{old_num}bis\. ', f'"{new_num}. ', text)
    
    # Replace H2 titles: "old.X Title" → "new.X Title"  
    text = re.sub(rf'"{old_num}\.(\d)', f'"{new_num}.\\1', text)
    text = re.sub(rf'"{old_num}bis\.(\d)', f'"{new_num}.\\1', text)
    
    # Replace section references in text: "section X" → "section Y"
    text = re.sub(rf'section {old_num}\b', f'section {new_num}', text, flags=re.IGNORECASE)
    text = re.sub(rf'section {old_num}bis\b', f'section {new_num}', text, flags=re.IGNORECASE)
    
    # Replace "figure X" references — these don't change, keep as is
    
    return text

# Actually, let me redo this properly with a function defined before use
# Re-read the file
with open('/home/z/my-project/scripts/build_pdf_report_v2.py', 'r', encoding='utf-8') as f:
    content = f.read()
lines = content.split('\n')

def get_section_lines(name):
    s, e = SECTIONS[name]
    return lines[s:e]

def renumber_section_lines(section_lines, old_num, new_num):
    """Renumber section references in a list of lines."""
    result = []
    for line in section_lines:
        # Replace H1/H2 titles: "old. Title" → "new. Title"
        line = re.sub(rf'"{old_num}\. ', f'"{new_num}. ', line)
        line = re.sub(rf'"{old_num}bis\. ', f'"{new_num}. ', line)
        # Replace H2 sub-titles: "old.X Title" → "new.X Title"
        line = re.sub(rf'"{old_num}\.(\d)', f'"{new_num}.\\1', line)
        line = re.sub(rf'"{old_num}bis\.(\d)', f'"{new_num}.\\1', line)
        # Replace section references in body text
        line = re.sub(rf'section {old_num}\b', f'section {new_num}', line, flags=re.IGNORECASE)
        line = re.sub(rf'section {old_num}bis\b', f'section {new_num}', line, flags=re.IGNORECASE)
        # Replace "la section X" patterns
        line = re.sub(rf'(la |de la |dans la |voir la |cf\. la ){old_num}\b', rf'\g<1>{new_num}', line, flags=re.IGNORECASE)
        line = re.sub(rf'(la |de la |dans la |voir la |cf\. la ){old_num}bis\b', rf'\g<1>{new_num}', line, flags=re.IGNORECASE)
        result.append(line)
    return result

# Build new sommaire
new_sommaire = '''    # ---------- PAGE 2: SOMMAIRE ----------
    story.append(Paragraph("Sommaire", H1))
    story.append(section_divider())
    story.append(Spacer(1, 0.3*cm))

    toc_data = [
        ["Section", "Titre", "Page"],
        ["", "Synthèse exécutive", "3"],
        ["1", "État des lieux et méthodologie", "4"],
        ["2", "Chiffres clés et segmentation clients", "5"],
        ["3", "Ventes vs Objectifs 2026", "7"],
        ["4", "Focus CONCENTRÉS — Cœur de marge", "12"],
        ["5", "Analyse Zéro Achat Q1", "16"],
        ["6", "Pertes Q1 estimées (méthode fréquence)", "18"],
        ["7", "Transition Q1→Q2 et dynamique mensuelle", "20"],
        ["8", "Analyse Chick Booster & Piglet Booster", "22"],
        ["9", "Perspectives stratégiques — 5 axes", "24"],
        ["10", "Projection CA 6 mois (3 scénarios)", "26"],
        ["11", "Forecast S2 2026 (3 scénarios)", "27"],
        ["12", "Analyse YoY 2025-2026 et forecast affiné", "29"],
        ["13", "Recommandations et plan d'action priorisé", "33"],
        ["14", "Conclusion et prochaines étapes", "35"],
    ]
    t = make_table(toc_data, col_widths=[1.5*cm, 12*cm, 2.5*cm], header_row=True)
    story.append(t)
    story.append(Spacer(1, 0.6*cm))
    story.append(Paragraph(
        "<i>Ce rapport présente l'analyse complète des ventes BELGOCAM SA sur la période Janvier-Juin 2026, "
        "après exclusion de 35 clients internes (filiales NJS, SPC, comptoirs d'agences et soldes comptables). "
        "L'analyse porte sur 1 359 clients actifs et 16 produits ciblés (4 tourteaux de soja + 10 BELGO 10% + 2 BELGO 5%).</i>",
        BODY_ITALIC
    ))
    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph(
        "<b>Glossaire des abréviations :</b> Q1 = Trimestre 1 (Janvier-Mars 2026) ; Q2 = Trimestre 2 (Avril-Juin 2026) ; "
        "S1 = Semestre 1 (Janvier-Juin 2026) ; S2 = Semestre 2 (Juillet-Décembre 2026) ; CA = Chiffre d'Affaires ; HT = Hors Taxes ; FCFA = Franc CFA.",
        SMALL
    ))'''

# Build new content
new_lines = []

# Header (everything before sommaire)
new_lines.extend(get_section_lines('header'))

# New sommaire
new_lines.extend(new_sommaire.split('\n'))

# Synthèse exécutive (unchanged numbering)
new_lines.extend(get_section_lines('synthese'))

# 1. État des lieux (unchanged)
new_lines.extend(get_section_lines('sec1'))

# 2. Chiffres clés (unchanged)
new_lines.extend(get_section_lines('sec2'))

# 3. Ventes vs Objectifs (was 9)
new_lines.extend(renumber_section_lines(get_section_lines('sec9_ventes'), '9', '3'))

# 4. Focus CONCENTRÉS (was 9bis)
new_lines.extend(renumber_section_lines(get_section_lines('sec9bis_conc'), '9bis', '4'))

# 5. Zéro Achat (was 3)
new_lines.extend(renumber_section_lines(get_section_lines('sec3_zero'), '3', '5'))

# 6. Pertes Q1 (was 4)
new_lines.extend(renumber_section_lines(get_section_lines('sec4_pertes'), '4', '6'))

# 7. Transition (was 5)
new_lines.extend(renumber_section_lines(get_section_lines('sec5_trans'), '5', '7'))

# 8. Chick & Piglet (was 6)
new_lines.extend(renumber_section_lines(get_section_lines('sec6_chick'), '6', '8'))

# 9. Perspectives (was 7)
new_lines.extend(renumber_section_lines(get_section_lines('sec7_persp'), '7', '9'))

# 10. Projection CA (was 8)
new_lines.extend(renumber_section_lines(get_section_lines('sec8_proj'), '8', '10'))

# 11. Forecast (was 10)
new_lines.extend(renumber_section_lines(get_section_lines('sec10_forecast'), '10', '11'))

# 12. YoY (was 10bis)
new_lines.extend(renumber_section_lines(get_section_lines('sec10bis_yoy'), '10bis', '12'))

# 13. Recommandations (was 11)
new_lines.extend(renumber_section_lines(get_section_lines('sec11_reco'), '11', '13'))

# 14. Conclusion (was 12)
new_lines.extend(renumber_section_lines(get_section_lines('sec12_concl'), '12', '14'))

# Join
new_content = '\n'.join(new_lines)

# Now replace "réel" when referring to objectives
# "CA objectif réel" → "CA objectif"
# "objectifs CA réels" → "objectifs CA"
# "CA obj réel" → "CA obj"
# But keep "CA réel" (actual sales) and "S1 réel" etc.
new_content = new_content.replace("CA objectif réel", "CA objectif")
new_content = new_content.replace("objectifs CA réels", "objectifs CA")
new_content = new_content.replace("objectif CA réel", "objectif CA")
new_content = new_content.replace("CA obj réel", "CA obj")
new_content = new_content.replace("CA obj. réel", "CA obj.")
new_content = new_content.replace("Prix obj réel", "Prix objectif")
new_content = new_content.replace("CA objectifs réels", "CA objectifs")
new_content = new_content.replace("réels × taux", "× taux")
new_content = new_content.replace("objectifs réels", "objectifs")
new_content = new_content.replace("Objectifs S2 recalibrés", "Objectifs S2 recalibrés")  # keep this one

# Also fix references to "section 3" that now point to "section 5" (Zero Achat)
# These were already handled by renumber_section_lines, but let's check cross-references
# "diagnostic de la section 3" in old text → should now be "section 5"
# The renumber function handles this for the section being renumbered,
# but cross-references FROM other sections need manual fixing
# Old section 3 (Zero Achat) → new section 5
# Old section 9 (Ventes) → new section 3
# Check for "section 3 (zéro achat" which should now be "section 5"
new_content = new_content.replace("section 3 (zéro achat", "section 5 (zéro achat")
new_content = new_content.replace("section 3 (Zéro", "section 5 (Zéro")

# Write
with open('/home/z/my-project/scripts/build_pdf_report_v3.py', 'w', encoding='utf-8') as f:
    f.write(new_content)

print(f"New script written: build_pdf_report_v3.py ({len(new_lines)} lines)")
print(f"Old script: {len(lines)} lines")
