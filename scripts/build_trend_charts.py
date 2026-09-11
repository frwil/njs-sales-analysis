"""Generate charts for the trend analysis (4 periods: W-3, W-2, W-1, P1).
Saves PNG files that will be embedded in the PDF.
"""
import json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
import os

# === Fonts ===
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf')
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# === Colors (BELGOCAM corporate) ===
NAVY = '#1F3A5F'
GOLD = '#C9A961'
RED = '#C00000'
GREEN = '#548235'
ORANGE = '#ED7D31'
GRAY = '#595959'
LIGHT_BLUE = '#5B9BD5'

# === Load data ===
DATA = json.load(open('/home/z/my-project/scripts/sept_trend_analysis.json'))
PERIODS_ORDER = ['W-3', 'W-2', 'W-1', 'P1']
PERIOD_LABELS = ['W-3\n(15-20/08)', 'W-2\n(22-27/08)', 'W-1\n(29/08-03/09)', 'P1\n(05-10/09)']

OUT_DIR = '/home/z/my-project/scripts/charts'
os.makedirs(OUT_DIR, exist_ok=True)

# === CHART 1: Global volumes (bar chart) - Soja vs Conc per period ===
fig, ax = plt.subplots(figsize=(10, 5.5), constrained_layout=True)
x = np.arange(len(PERIODS_ORDER))
width = 0.35

soja_vals = [DATA['global'][p]['soja_t'] for p in PERIODS_ORDER]
conc_vals = [DATA['global'][p]['conc_t'] for p in PERIODS_ORDER]

bars1 = ax.bar(x - width/2, soja_vals, width, label='Soja (TOURTEAUX)', color=NAVY)
bars2 = ax.bar(x + width/2, conc_vals, width, label='Concentrés', color=GOLD)

# Add value labels on bars
for bar in bars1:
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2, h + 15, f'{h:.0f}', 
            ha='center', va='bottom', fontsize=9, fontweight='bold', color=NAVY)
for bar in bars2:
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2, h + 15, f'{h:.0f}', 
            ha='center', va='bottom', fontsize=9, fontweight='bold', color=GOLD)

ax.set_xticks(x)
ax.set_xticklabels(PERIOD_LABELS, fontsize=9)
ax.set_ylabel('Volume (tonnes)', fontsize=10)
ax.set_title('Évolution des volumes Soja vs Concentrés sur 4 semaines\n(chaque période = 5 jours ouvrés lun-sam)', 
             fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.legend(loc='upper right', fontsize=9)
ax.grid(True, alpha=0.3, axis='y')
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
# Highlight P1 with arrow
ax.annotate('Hausse prix 04/09\n(27 000 FCFA/sac)', xy=(3, 480), xytext=(2.2, 850),
            fontsize=8, color=RED, ha='center',
            arrowprops=dict(arrowstyle='->', color=RED, lw=1.5))
# Set y-limit to leave space for labels
ax.set_ylim(0, max(max(soja_vals), max(conc_vals)) * 1.2)

plt.savefig(f'{OUT_DIR}/chart1_global_volumes.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart 1: Global volumes')

# === CHART 2: Daily average trend (line chart) - Soja and Conc moy/j ===
fig, ax = plt.subplots(figsize=(10, 5.5), constrained_layout=True)

soja_moy = [DATA['global'][p]['soja_moy_j'] for p in PERIODS_ORDER]
conc_moy = [DATA['global'][p]['conc_moy_j'] for p in PERIODS_ORDER]

ax.plot(x, soja_moy, marker='o', linewidth=2.5, markersize=10, color=NAVY, label='Soja moy/j (t)')
ax.plot(x, conc_moy, marker='s', linewidth=2.5, markersize=10, color=GOLD, label='Conc moy/j (t)')

# Add value labels
for i, (s, c) in enumerate(zip(soja_moy, conc_moy)):
    ax.annotate(f'{s:.1f}', xy=(i, s), xytext=(0, 12), textcoords='offset points',
                ha='center', fontsize=9, fontweight='bold', color=NAVY)
    ax.annotate(f'{c:.1f}', xy=(i, c), xytext=(0, -18), textcoords='offset points',
                ha='center', fontsize=9, fontweight='bold', color=GOLD)

# Add vertical line between W-1 and P1 (price hike)
ax.axvline(x=2.5, color=RED, linestyle='--', alpha=0.7, linewidth=1.5)
ax.text(2.5, max(soja_moy)*0.95, ' ↑ Hausse prix\n   04/09', color=RED, fontsize=8, ha='left', va='top')

ax.set_xticks(x)
ax.set_xticklabels(PERIOD_LABELS, fontsize=9)
ax.set_ylabel('Cadence quotidienne (t/jour)', fontsize=10)
ax.set_title('Tendance cadence quotidienne — Soja vs Concentrés\n(moyenne sur 5 jours ouvrés par période)', 
             fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.legend(loc='upper right', fontsize=9)
ax.grid(True, alpha=0.3)
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_ylim(0, max(soja_moy) * 1.25)

plt.savefig(f'{OUT_DIR}/chart2_daily_trend.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart 2: Daily trend')

# === CHART 3: Stacked bar — Total per region over 4 periods ===
fig, ax = plt.subplots(figsize=(10, 5.5), constrained_layout=True)

regions = ['Ouest', 'Centre', 'Littoral']
colors_reg = [NAVY, GOLD, ORANGE]
x_reg = np.arange(len(PERIODS_ORDER))
width_reg = 0.25

for i, (region, color) in enumerate(zip(regions, colors_reg)):
    soja_vals = [DATA['regions'][region][p]['soja_t'] for p in PERIODS_ORDER]
    conc_vals = [DATA['regions'][region][p]['conc_t'] for p in PERIODS_ORDER]
    totals = [s + c for s, c in zip(soja_vals, conc_vals)]
    
    bars = ax.bar(x_reg + (i-1)*width_reg, totals, width_reg, label=region, color=color)
    # Add value labels
    for j, bar in enumerate(bars):
        h = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, h + 15, f'{h:.0f}',
                ha='center', va='bottom', fontsize=8, color=color, fontweight='bold')

ax.set_xticks(x_reg)
ax.set_xticklabels(PERIOD_LABELS, fontsize=9)
ax.set_ylabel('Volume total (tonnes)', fontsize=10)
ax.set_title('Volume total (Soja + Conc) par région sur 4 semaines\n(chaque période = 5 jours ouvrés)', 
             fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.legend(loc='upper right', fontsize=9)
ax.grid(True, alpha=0.3, axis='y')
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_ylim(0, 700)

plt.savefig(f'{OUT_DIR}/chart3_regional_stacked.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart 3: Regional stacked')

# === CHART 4: Per-region evolution (small multiples) - 3 subplots ===
fig, axes = plt.subplots(1, 3, figsize=(13, 4.5), constrained_layout=True, sharey=False)

for i, (region, color) in enumerate(zip(regions, colors_reg)):
    ax = axes[i]
    soja_vals = [DATA['regions'][region][p]['soja_t'] for p in PERIODS_ORDER]
    conc_vals = [DATA['regions'][region][p]['conc_t'] for p in PERIODS_ORDER]
    
    x_local = np.arange(len(PERIODS_ORDER))
    bars1 = ax.bar(x_local - 0.2, soja_vals, 0.4, label='Soja', color=NAVY)
    bars2 = ax.bar(x_local + 0.2, conc_vals, 0.4, label='Conc', color=GOLD)
    
    ax.set_xticks(x_local)
    ax.set_xticklabels(['W-3', 'W-2', 'W-1', 'P1'], fontsize=9)
    ax.set_title(region, fontsize=11, fontweight='bold', color=color)
    ax.grid(True, alpha=0.3, axis='y')
    ax.set_axisbelow(True)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    if i == 0:
        ax.set_ylabel('Volume (t)', fontsize=10)
        ax.legend(fontsize=8, loc='upper right')
    
    # Annotate max value
    max_val = max(max(soja_vals), max(conc_vals))
    ax.set_ylim(0, max_val * 1.25)
    
    # Highlight Littoral anomaly W-2
    if region == 'Littoral':
        ax.annotate('Pic NDOBO\n(488 t)', xy=(1, soja_vals[1]), xytext=(1.5, soja_vals[1]*0.85),
                    fontsize=7, color=RED, ha='center', fontweight='bold',
                    arrowprops=dict(arrowstyle='->', color=RED, lw=1.2))

fig.suptitle('Évolution par région — Soja vs Concentrés sur 4 semaines', 
             fontsize=13, fontweight='bold', color=NAVY, y=1.02)

plt.savefig(f'{OUT_DIR}/chart4_region_multiples.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart 4: Regional multiples')

# === CHART 5: Top 10 agences - grouped bar (Soja P1 vs avg W-1/W-2/W-3) ===
fig, ax = plt.subplots(figsize=(12, 6), constrained_layout=True)

# Get top 10 agences by P1 soja
top_agences = list(DATA['agences'].items())[:10]
ag_names = [ag[0] for ag in top_agences]
p1_soja = [ag[1]['P1']['soja_t'] for ag in top_agences]
w1_soja = [ag[1]['W-1']['soja_t'] for ag in top_agences]
w2_soja = [ag[1]['W-2']['soja_t'] for ag in top_agences]
w3_soja = [ag[1]['W-3']['soja_t'] for ag in top_agences]

x_ag = np.arange(len(ag_names))
width_ag = 0.2

ax.bar(x_ag - 1.5*width_ag, w3_soja, width_ag, label='W-3 (15-20/08)', color='#A8C5E2')
ax.bar(x_ag - 0.5*width_ag, w2_soja, width_ag, label='W-2 (22-27/08)', color=LIGHT_BLUE)
ax.bar(x_ag + 0.5*width_ag, w1_soja, width_ag, label='W-1 (29/08-03/09)', color=NAVY)
ax.bar(x_ag + 1.5*width_ag, p1_soja, width_ag, label='P1 (05-10/09)', color=GOLD)

ax.set_xticks(x_ag)
ax.set_xticklabels(ag_names, fontsize=9, rotation=30, ha='right')
ax.set_ylabel('Volume Soja (tonnes)', fontsize=10)
ax.set_title('Top 10 agences — Évolution du volume SOJA sur 4 semaines', 
             fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.legend(loc='upper right', fontsize=8, ncol=2)
ax.grid(True, alpha=0.3, axis='y')
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# Annotate NDOBO anomaly
ndobo_idx = ag_names.index('Ndobo') if 'Ndobo' in ag_names else -1
if ndobo_idx >= 0:
    ax.annotate(f'Pic W-2: {w2_soja[ndobo_idx]:.0f} t', 
                xy=(ndobo_idx - 0.5*width_ag, w2_soja[ndobo_idx]),
                xytext=(ndobo_idx, w2_soja[ndobo_idx] * 1.15),
                fontsize=8, color=RED, ha='center', fontweight='bold',
                arrowprops=dict(arrowstyle='->', color=RED, lw=1.2))

plt.savefig(f'{OUT_DIR}/chart5_top_agences.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart 5: Top agences')

# === CHART 6: Ratio soja/conc evolution (line chart) ===
fig, ax = plt.subplots(figsize=(10, 5.5), constrained_layout=True)

ratios_global = [DATA['global'][p]['ratio'] for p in PERIODS_ORDER]
ratios_ouest = [DATA['regions']['Ouest'][p]['soja_t']/DATA['regions']['Ouest'][p]['conc_t'] if DATA['regions']['Ouest'][p]['conc_t'] > 0 else 0 for p in PERIODS_ORDER]
ratios_centre = [DATA['regions']['Centre'][p]['soja_t']/DATA['regions']['Centre'][p]['conc_t'] if DATA['regions']['Centre'][p]['conc_t'] > 0 else 0 for p in PERIODS_ORDER]
ratios_littoral = [DATA['regions']['Littoral'][p]['soja_t']/DATA['regions']['Littoral'][p]['conc_t'] if DATA['regions']['Littoral'][p]['conc_t'] > 0 else 0 for p in PERIODS_ORDER]

ax.plot(x, ratios_global, marker='o', linewidth=2.5, markersize=10, color=NAVY, label='National')
ax.plot(x, ratios_ouest, marker='s', linewidth=2, markersize=8, color=GOLD, label='Ouest')
ax.plot(x, ratios_centre, marker='^', linewidth=2, markersize=8, color=GREEN, label='Centre')
ax.plot(x, ratios_littoral, marker='D', linewidth=2, markersize=8, color=ORANGE, label='Littoral')

# Objective line
ax.axhline(y=2.5, color=RED, linestyle='--', alpha=0.6, linewidth=1.2, label='Objectif ≤ 2,5:1')

# Add value labels for national
for i, r in enumerate(ratios_global):
    ax.annotate(f'{r:.2f}', xy=(i, r), xytext=(0, 12), textcoords='offset points',
                ha='center', fontsize=9, fontweight='bold', color=NAVY)

# Highlight Littoral anomaly
ax.annotate(f'Anomalie W-2\nLittoral: {ratios_littoral[1]:.2f}:1', xy=(1, ratios_littoral[1]), 
            xytext=(0.5, ratios_littoral[1] - 1.5), fontsize=8, color=RED, ha='center',
            arrowprops=dict(arrowstyle='->', color=RED, lw=1.2))

ax.set_xticks(x)
ax.set_xticklabels(PERIOD_LABELS, fontsize=9)
ax.set_ylabel('Ratio soja/concentrés', fontsize=10)
ax.set_title('Évolution du ratio bundle soja/concentrés\n(objectif ≤ 2,5:1)', 
             fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.legend(loc='upper right', fontsize=8)
ax.grid(True, alpha=0.3)
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_ylim(0, max(ratios_littoral) * 1.15)

plt.savefig(f'{OUT_DIR}/chart6_ratio_evolution.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart 6: Ratio evolution')

print(f"\nAll 6 charts saved in {OUT_DIR}/")
