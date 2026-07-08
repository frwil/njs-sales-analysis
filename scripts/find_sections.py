"""
Reorganize the PDF script: move Ventes vs Objectifs (9) and Focus CONCENTRÉS (9bis)
BEFORE Zéro Achat (3), Pertes (4), Transition (5).

New structure:
1. État des lieux
2. Chiffres clés
3. Ventes vs Objectifs (was 9, 9.7-9.10, 9.2bis)
4. Focus CONCENTRÉS (was 9bis)
5. Zéro Achat (was 3)
6. Pertes Q1 (was 4)
7. Transition Q1→Q2 (was 5)
8. Chick & Piglet (was 6)
9. Perspectives (was 7)
10. Projection CA (was 8)
11. Forecast S2 + YoY + recalibrage (was 10, 10bis)
12. Recommandations (was 11)
13. Conclusion (was 12)

Also: replace "réel" with "" when referring to objectives ("CA objectif réel" → "CA objectif")
"""
import re

with open('/home/z/my-project/scripts/build_pdf_report_v2.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find line ranges for each section
# Sections are delimited by "# ----------" comments and H1/H2 titles
sections = {}

# We need to find the exact line ranges
# Strategy: find all H1 titles and their line numbers
h1_titles = []
for i, line in enumerate(lines):
    if 'story.append(Paragraph("' in line and ', H1))' in line:
        # Extract the title
        m = re.search(r'Paragraph\("([^"]+)", H1\)', line)
        if m:
            h1_titles.append((i, m.group(1)))

print("H1 sections found:")
for i, (ln, title) in enumerate(h1_titles):
    end = h1_titles[i+1][0] if i+1 < len(h1_titles) else len(lines)
    print(f"  Line {ln+1}: {title[:60]} (ends at {end})")

# Also find the sommaire section
sommaire_start = None
for i, line in enumerate(lines):
    if 'toc_data = [' in line:
        sommaire_start = i
        break

# Find key boundaries
# Build a map of section_name -> (start_line, end_line)
section_map = {}
for i, (ln, title) in enumerate(h1_titles):
    start = ln
    # Find the PageBreak before this section (go back to find the comment block)
    for j in range(ln, max(ln-5, 0), -1):
        if '# ----------' in lines[j]:
            start = j
            break
    end = h1_titles[i+1][0] if i+1 < len(h1_titles) else len(lines)
    # Find PageBreak before next section
    for j in range(end, min(end+5, len(lines))):
        if 'story.append(PageBreak())' in lines[j]:
            end = j + 1
            break
    section_map[title] = (start, end)

# Print section ranges
print("\nSection ranges:")
for title, (s, e) in section_map.items():
    print(f"  {title[:50]:<52} lines {s+1}-{e}")
