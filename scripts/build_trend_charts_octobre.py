"""Generate charts for the October trend analysis (5 periods: W-3, W-2, W-1, P1, P2).
Régénère les 12 graphiques (chart1-6, chartA-F) avec les annotations
spécifiques à octobre : baisse de prix du 22/09 puis du 07/10, reprise soja,
tassement concentrés, dégradation du ratio.

Périodes (6 jours lun-sam chacune) :
P1   : 05-10/10/2026 (post-baisse 07/10 : 20 500 → 17 600 F/sac)
P2   : 12-17/10/2026 (première semaine complète post-baisse)
W-1 : 28/09-03/10/2026
W-2 : 21-26/09/2026
W-3 : 14-19/09/2026
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
DATA = json.load(open('/home/z/my-project/scripts/oct_trend_analysis.json'))
EXT = json.load(open('/home/z/my-project/scripts/oct_trend_extended.json'))
PERIODS_ORDER = ['W-3', 'W-2', 'W-1', 'P1', 'P2']
PERIOD_LABELS = ['W-3\n(14-19/09)', 'W-2\n(21-26/09)', 'W-1\n(28/09-03/10)', 'P1\n(05-10/10)', 'P2\n(12-17/10)']
WORK_DAYS = 6

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
ax.set_title('Évolution des volumes Soja vs Concentrés sur 5 semaines\n(chaque période = 6 jours lun-sam)',
             fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.legend(loc='upper right', fontsize=9)
ax.grid(True, alpha=0.3, axis='y')
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
# Annotate price cuts
ax.annotate('Baisse prix 22/09\n(→ ~21 000 F/sac)', xy=(1, soja_vals[1]), xytext=(0.45, max(soja_vals) * 1.22),
            fontsize=8, color=GREEN, ha='center',
            arrowprops=dict(arrowstyle='->', color=GREEN, lw=1.5))
ax.annotate('Baisse prix 07/10\n(20 500 → 17 600 F/sac)', xy=(3, soja_vals[3]), xytext=(2.7, max(soja_vals) * 1.22),
            fontsize=8, color=RED, ha='center',
            arrowprops=dict(arrowstyle='->', color=RED, lw=1.5))
ax.set_ylim(0, max(max(soja_vals), max(conc_vals)) * 1.35)

plt.savefig(f'{OUT_DIR}/chart1_global_volumes.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart 1: Global volumes')

# === CHART 2: Daily average trend (line chart) ===
fig, ax = plt.subplots(figsize=(10, 5.5), constrained_layout=True)

soja_moy = [DATA['global'][p]['soja_moy_j'] for p in PERIODS_ORDER]
conc_moy = [DATA['global'][p]['conc_moy_j'] for p in PERIODS_ORDER]

ax.plot(x, soja_moy, marker='o', linewidth=2.5, markersize=10, color=NAVY, label='Soja moy/j (t)')
ax.plot(x, conc_moy, marker='s', linewidth=2.5, markersize=10, color=GOLD, label='Conc moy/j (t)')

for i, (s, c) in enumerate(zip(soja_moy, conc_moy)):
    ax.annotate(f'{s:.1f}', xy=(i, s), xytext=(0, 12), textcoords='offset points',
                ha='center', fontsize=9, fontweight='bold', color=NAVY)
    ax.annotate(f'{c:.1f}', xy=(i, c), xytext=(0, -18), textcoords='offset points',
                ha='center', fontsize=9, fontweight='bold', color=GOLD)

ax.axvline(x=0.5, color=GREEN, linestyle='--', alpha=0.7, linewidth=1.5)
ax.text(0.5, max(soja_moy)*0.97, ' ↓ Baisse prix\n   22/09', color=GREEN, fontsize=8, ha='left', va='top')
ax.axvline(x=2.5, color=RED, linestyle='--', alpha=0.7, linewidth=1.5)
ax.text(2.5, max(soja_moy)*0.97, ' ↓ Baisse prix\n   07/10 (17 600 F)', color=RED, fontsize=8, ha='left', va='top')

ax.set_xticks(x)
ax.set_xticklabels(PERIOD_LABELS, fontsize=9)
ax.set_ylabel('Cadence quotidienne (t/jour)', fontsize=10)
ax.set_title('Tendance cadence quotidienne — Soja vs Concentrés\n(moyenne sur 6 jours lun-sam par période)',
             fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.legend(loc='upper right', fontsize=9)
ax.grid(True, alpha=0.3)
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_ylim(0, max(soja_moy) * 1.3)

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
    for j, bar in enumerate(bars):
        h = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, h + 15, f'{h:.0f}',
                ha='center', va='bottom', fontsize=8, color=color, fontweight='bold')

ax.set_xticks(x_reg)
ax.set_xticklabels(PERIOD_LABELS, fontsize=9)
ax.set_ylabel('Volume total (tonnes)', fontsize=10)
ax.set_title('Volume total (Soja + Conc) par région sur 5 semaines\n(chaque période = 6 jours lun-sam)',
             fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.legend(loc='upper right', fontsize=9)
ax.grid(True, alpha=0.3, axis='y')
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
max_total = max(DATA['regions'][r][p]['soja_t'] + DATA['regions'][r][p]['conc_t']
                for r in regions for p in PERIODS_ORDER)
ax.set_ylim(0, max_total * 1.25)

plt.savefig(f'{OUT_DIR}/chart3_regional_stacked.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart 3: Regional stacked')

# === CHART 4: Per-region evolution (small multiples) ===
fig, axes = plt.subplots(1, 3, figsize=(13, 4.5), constrained_layout=True, sharey=False)

for i, (region, color) in enumerate(zip(regions, colors_reg)):
    ax = axes[i]
    soja_vals = [DATA['regions'][region][p]['soja_t'] for p in PERIODS_ORDER]
    conc_vals = [DATA['regions'][region][p]['conc_t'] for p in PERIODS_ORDER]

    x_local = np.arange(len(PERIODS_ORDER))
    bars1 = ax.bar(x_local - 0.2, soja_vals, 0.4, label='Soja', color=NAVY)
    bars2 = ax.bar(x_local + 0.2, conc_vals, 0.4, label='Conc', color=GOLD)

    ax.set_xticks(x_local)
    ax.set_xticklabels(['W-3', 'W-2', 'W-1', 'P1', 'P2'], fontsize=9)
    ax.set_title(region, fontsize=11, fontweight='bold', color=color)
    ax.grid(True, alpha=0.3, axis='y')
    ax.set_axisbelow(True)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    if i == 0:
        ax.set_ylabel('Volume (t)', fontsize=10)
        ax.legend(fontsize=8, loc='upper right')

    max_val = max(max(soja_vals), max(conc_vals))
    ax.set_ylim(0, max_val * 1.25)

    if region == 'Ouest':
        ax.annotate(f'Pic Famla P1\n({soja_vals[3]:.0f} t soja)', xy=(3, soja_vals[3]), xytext=(2.2, soja_vals[3]*1.02),
                    fontsize=7, color=RED, ha='center', fontweight='bold',
                    arrowprops=dict(arrowstyle='->', color=RED, lw=1.2))
    if region == 'Littoral':
        ax.annotate('Repli P1\naprès pic W-1', xy=(3, soja_vals[3]), xytext=(2.0, soja_vals[1]*0.95),
                    fontsize=7, color=RED, ha='center', fontweight='bold',
                    arrowprops=dict(arrowstyle='->', color=RED, lw=1.2))

fig.suptitle('Évolution par région — Soja vs Concentrés sur 5 semaines (octobre)',
             fontsize=13, fontweight='bold', color=NAVY, y=1.02)

plt.savefig(f'{OUT_DIR}/chart4_region_multiples.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart 4: Regional multiples')

# === CHART 5: Top 10 agences - grouped bar (Soja) ===
fig, ax = plt.subplots(figsize=(12, 6), constrained_layout=True)

top_agences = list(DATA['agences'].items())[:10]
ag_names = [ag[0] for ag in top_agences]
p2_soja = [ag[1]['P2']['soja_t'] for ag in top_agences]
p1_soja = [ag[1]['P1']['soja_t'] for ag in top_agences]
w1_soja = [ag[1]['W-1']['soja_t'] for ag in top_agences]
w2_soja = [ag[1]['W-2']['soja_t'] for ag in top_agences]
w3_soja = [ag[1]['W-3']['soja_t'] for ag in top_agences]

x_ag = np.arange(len(ag_names))
width_ag = 0.16

ax.bar(x_ag - 2*width_ag, w3_soja, width_ag, label='W-3 (14-19/09)', color='#A8C5E2')
ax.bar(x_ag - 1*width_ag, w2_soja, width_ag, label='W-2 (21-26/09)', color=LIGHT_BLUE)
ax.bar(x_ag, w1_soja, width_ag, label='W-1 (28/09-03/10)', color=NAVY)
ax.bar(x_ag + 1*width_ag, p1_soja, width_ag, label='P1 (05-10/10)', color=GOLD)
ax.bar(x_ag + 2*width_ag, p2_soja, width_ag, label='P2 (12-17/10)', color=RED)

ax.set_xticks(x_ag)
ax.set_xticklabels(ag_names, fontsize=9, rotation=30, ha='right')
ax.set_ylabel('Volume Soja (tonnes)', fontsize=10)
ax.set_title('Top 10 agences — Évolution du volume SOJA sur 5 semaines',
             fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.legend(loc='upper right', fontsize=8, ncol=2)
ax.grid(True, alpha=0.3, axis='y')
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

famla_idx = ag_names.index('Famla') if 'Famla' in ag_names else -1
if famla_idx >= 0:
    ax.annotate(f'Famla: {w3_soja[famla_idx]:.0f} → {p2_soja[famla_idx]:.0f} t\n({(p2_soja[famla_idx]/w3_soja[famla_idx]-1)*100:+.0f}%)',
                xy=(famla_idx + 2*width_ag, p2_soja[famla_idx]),
                xytext=(famla_idx, p2_soja[famla_idx] * 1.08),
                fontsize=8, color=RED, ha='center', fontweight='bold',
                arrowprops=dict(arrowstyle='->', color=RED, lw=1.2))

plt.savefig(f'{OUT_DIR}/chart5_top_agences.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart 5: Top agences')

# === CHART 6: Ratio soja/conc evolution ===
fig, ax = plt.subplots(figsize=(10, 5.5), constrained_layout=True)

ratios_global = [DATA['global'][p]['ratio'] for p in PERIODS_ORDER]
ratios_ouest = [DATA['regions']['Ouest'][p]['soja_t']/DATA['regions']['Ouest'][p]['conc_t'] if DATA['regions']['Ouest'][p]['conc_t'] > 0 else 0 for p in PERIODS_ORDER]
ratios_centre = [DATA['regions']['Centre'][p]['soja_t']/DATA['regions']['Centre'][p]['conc_t'] if DATA['regions']['Centre'][p]['conc_t'] > 0 else 0 for p in PERIODS_ORDER]
ratios_littoral = [DATA['regions']['Littoral'][p]['soja_t']/DATA['regions']['Littoral'][p]['conc_t'] if DATA['regions']['Littoral'][p]['conc_t'] > 0 else 0 for p in PERIODS_ORDER]

ax.plot(x, ratios_global, marker='o', linewidth=2.5, markersize=10, color=NAVY, label='National')
ax.plot(x, ratios_ouest, marker='s', linewidth=2, markersize=8, color=GOLD, label='Ouest')
ax.plot(x, ratios_centre, marker='^', linewidth=2, markersize=8, color=GREEN, label='Centre')
ax.plot(x, ratios_littoral, marker='D', linewidth=2, markersize=8, color=ORANGE, label='Littoral')

ax.axhline(y=2.5, color=RED, linestyle='--', alpha=0.6, linewidth=1.2, label='Objectif ≤ 2,5:1')

for i, r in enumerate(ratios_global):
    ax.annotate(f'{r:.2f}', xy=(i, r), xytext=(0, 12), textcoords='offset points',
                ha='center', fontsize=9, fontweight='bold', color=NAVY)

ax.annotate('Dégradation du ratio :\nle soja repart, pas les conc', xy=(4, ratios_global[4]),
            xytext=(1.8, max(ratios_littoral) * 1.08), fontsize=8, color=RED, ha='center',
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
ax.set_ylim(0, max(ratios_littoral) * 1.2)

plt.savefig(f'{OUT_DIR}/chart6_ratio_evolution.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart 6: Ratio evolution')

# =====================================================================
# CHART A: CONCENTRÉS focus — reprise modérée puis plateau
# =====================================================================
fig, ax = plt.subplots(figsize=(10, 5.5), constrained_layout=True)

conc_trend = [DATA['global'][p]['conc_t'] for p in PERIODS_ORDER]
bars = ax.bar(x, conc_trend, 0.5, color=GOLD, edgecolor=NAVY, linewidth=1.2)

for i, bar in enumerate(bars):
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2, h + 8, f'{h:.1f} t',
            ha='center', va='bottom', fontsize=10, fontweight='bold', color=NAVY)

for i in range(1, len(conc_trend)):
    prev_h = conc_trend[i-1]
    curr_h = conc_trend[i]
    mid_x = (i-1 + i) / 2
    mid_y = max(prev_h, curr_h) + 30
    color = RED if curr_h < prev_h else GREEN
    ax.annotate('', xy=(i, curr_h + 5), xytext=(i-1, prev_h + 5),
                arrowprops=dict(arrowstyle='->', color=color, lw=1.5, alpha=0.7))
    var = (curr_h/prev_h - 1)*100
    ax.text(mid_x, mid_y + 15, f'{var:+.1f}%', ha='center', fontsize=10,
            fontweight='bold', color=color,
            bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor=color, alpha=0.9))

total_var = (conc_trend[-1]/conc_trend[0] - 1)*100
soja_var_cumul = (DATA['global']['P2']['soja_t']/DATA['global']['W-3']['soja_t'] - 1)*100
ax.text(0.02, 0.95, f'Reprise cumulée\nW-3 → P2: +{total_var:.1f}%\n(soja : +{soja_var_cumul:.0f} %)',
        transform=ax.transAxes, fontsize=10, fontweight='bold', color=GREEN,
        bbox=dict(boxstyle='round,pad=0.5', facecolor='#EAF3E0', edgecolor=GREEN, alpha=0.9),
        verticalalignment='top')

ax.set_xticks(x)
ax.set_xticklabels(PERIOD_LABELS, fontsize=9)
ax.set_ylabel('Volume Concentrés (tonnes)', fontsize=10)
ax.set_title(f'CONCENTRÉS — Reprise modérée (+{total_var:.0f} %) puis plateau depuis W-2\n(le concentré ne suit pas la reprise du soja)',
             fontsize=12, fontweight='bold', color=RED, pad=10)
ax.grid(True, alpha=0.3, axis='y')
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_ylim(0, max(conc_trend) * 1.35)

plt.savefig(f'{OUT_DIR}/chartA_conc_decline.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart A: CONC reprene moderee')

# === CHART B: Comparison SOJA vs CONC — divergence (index W-3 = 100) ===
fig, ax = plt.subplots(figsize=(11, 6), constrained_layout=True)

soja_moy_trend = [DATA['global'][p]['soja_moy_j'] for p in PERIODS_ORDER]
conc_moy_trend = [DATA['global'][p]['conc_moy_j'] for p in PERIODS_ORDER]
soja_index = [v/soja_moy_trend[0]*100 for v in soja_moy_trend]
conc_index = [v/conc_moy_trend[0]*100 for v in conc_moy_trend]

ax.plot(x, soja_index, marker='o', linewidth=3, markersize=12, color=NAVY, label='Soja (TOURTEAUX) moy/j')
ax.plot(x, conc_index, marker='s', linewidth=3, markersize=12, color=GOLD, label='Concentrés moy/j', linestyle='--')

ax.axhline(y=100, color=GRAY, linestyle=':', alpha=0.6, linewidth=1)
ax.text(4.1, 100, 'Base W-3 = 100', fontsize=8, color=GRAY, va='center')

for i, (s, c) in enumerate(zip(soja_index, conc_index)):
    ax.annotate(f'{s:.0f}', xy=(i, s), xytext=(0, 12), textcoords='offset points',
                ha='center', fontsize=10, fontweight='bold', color=NAVY)
    ax.annotate(f'{c:.0f}', xy=(i, c), xytext=(0, -20), textcoords='offset points',
                ha='center', fontsize=10, fontweight='bold', color=GOLD)

ax.fill_between(x, soja_index, conc_index, where=[s >= c for s, c in zip(soja_index, conc_index)],
                color=RED, alpha=0.1, label='Divergence (Soja > Conc)')

soja_var_b = (soja_moy_trend[-1]/soja_moy_trend[0] - 1)*100
conc_var_b = (conc_moy_trend[-1]/conc_moy_trend[0] - 1)*100
ax.annotate(f'Soja: +{soja_var_b:.0f} % cumulé\n(2 baisses de prix)',
            xy=(4, soja_index[4]), xytext=(2.3, soja_index[4] * 1.1),
            fontsize=9, color=NAVY, ha='center',
            arrowprops=dict(arrowstyle='->', color=NAVY, lw=1.2))
ax.annotate(f'Conc: +{conc_var_b:.0f} % seulement\n(paniers en baisse)',
            xy=(4, conc_index[4]), xytext=(2.6, conc_index[2] * 0.8),
            fontsize=9, color=RED, ha='center', fontweight='bold',
            arrowprops=dict(arrowstyle='->', color=RED, lw=1.5))

ax.set_xticks(x)
ax.set_xticklabels(PERIOD_LABELS, fontsize=9)
ax.set_ylabel('Indice (base W-3 = 100)', fontsize=10)
ax.set_title(f'Dynamiques divergentes : SOJA (+{soja_var_b:.0f} %) vs CONCENTRÉS (+{conc_var_b:.0f} %)\nCadence quotidienne indexée à 100 en W-3',
             fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.legend(loc='upper right', fontsize=9)
ax.grid(True, alpha=0.3)
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_ylim(min(min(soja_index), min(conc_index), 60) * 0.8,
            max(max(soja_index), max(conc_index)) * 1.25)

plt.savefig(f'{OUT_DIR}/chartB_divergence.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart B: Divergence SOJA vs CONC')

# === CHART C: CONC by region — reprise inégale ===
fig, ax = plt.subplots(figsize=(10, 5.5), constrained_layout=True)

regions = ['Ouest', 'Centre', 'Littoral']
colors_reg = [NAVY, GOLD, ORANGE]
markers = ['o', 's', 'D']

for i, (region, color, marker) in enumerate(zip(regions, colors_reg, markers)):
    conc_vals = [DATA['regions'][region][p]['conc_t'] for p in PERIODS_ORDER]
    conc_moy = [v/WORK_DAYS for v in conc_vals]

    ax.plot(x, conc_moy, marker=marker, linewidth=2.5, markersize=10, color=color, label=region)

    ax.annotate(f'{conc_moy[-1]:.1f}', xy=(4, conc_moy[-1]), xytext=(8, 0),
                textcoords='offset points', fontsize=9, color=color, fontweight='bold')
    ax.annotate(f'{conc_moy[0]:.1f}', xy=(0, conc_moy[0]), xytext=(-15, 0),
                textcoords='offset points', fontsize=9, color=color, fontweight='bold',
                ha='right')

ax.set_xticks(x)
ax.set_xticklabels(PERIOD_LABELS, fontsize=9)
ax.set_ylabel('Cadence Conc moy/j (tonnes)', fontsize=10)
ax.set_title('CONCENTRÉS par région — Reprise inégale\n(pic P1 post-baisse non prolongé : repli généralisé en P2)',
             fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.legend(loc='upper right', fontsize=9)
ax.grid(True, alpha=0.3)
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
max_conc_moy = max(max(DATA['regions'][r][p]['conc_t'] for p in PERIODS_ORDER) / WORK_DAYS for r in regions)
ax.set_ylim(0, max_conc_moy * 1.3)

plt.savefig(f'{OUT_DIR}/chartC_conc_regional.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart C: CONC regional')

# === CHART D: CONC by top 10 agences ===
fig, ax = plt.subplots(figsize=(12, 6), constrained_layout=True)

top_agences = list(DATA['agences'].items())[:10]
ag_names = [ag[0] for ag in top_agences]
p2_conc = [ag[1]['P2']['conc_t'] for ag in top_agences]
p1_conc = [ag[1]['P1']['conc_t'] for ag in top_agences]
w1_conc = [ag[1]['W-1']['conc_t'] for ag in top_agences]
w2_conc = [ag[1]['W-2']['conc_t'] for ag in top_agences]
w3_conc = [ag[1]['W-3']['conc_t'] for ag in top_agences]

x_ag = np.arange(len(ag_names))
width_ag = 0.16

ax.bar(x_ag - 2*width_ag, w3_conc, width_ag, label='W-3 (14-19/09)', color='#A8C5E2')
ax.bar(x_ag - 1*width_ag, w2_conc, width_ag, label='W-2 (21-26/09)', color=LIGHT_BLUE)
ax.bar(x_ag, w1_conc, width_ag, label='W-1 (28/09-03/10)', color=NAVY)
ax.bar(x_ag + 1*width_ag, p1_conc, width_ag, label='P1 (05-10/10)', color=GOLD)
ax.bar(x_ag + 2*width_ag, p2_conc, width_ag, label='P2 (12-17/10)', color=RED)

ax.set_xticks(x_ag)
ax.set_xticklabels(ag_names, fontsize=9, rotation=30, ha='right')
ax.set_ylabel('Volume Concentrés (tonnes)', fontsize=10)
ax.set_title('Top 10 agences — Évolution du volume CONCENTRÉS sur 5 semaines\n(le concentré ne suit pas la reprise soja, surtout à Famla)',
             fontsize=12, fontweight='bold', color=RED, pad=10)
ax.legend(loc='upper right', fontsize=8, ncol=2)
ax.grid(True, alpha=0.3, axis='y')
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

famla_idx = ag_names.index('Famla') if 'Famla' in ag_names else -1
if famla_idx >= 0:
    ax.annotate(f'FAMLA: soja x{1 + (p2_soja[famla_idx]/w3_soja[famla_idx] - 1):+.0%}\nmais conc {w3_conc[famla_idx]:.0f} → {p2_conc[famla_idx]:.0f} t ({(p2_conc[famla_idx]/w3_conc[famla_idx]-1)*100:+.0f}%)',
                xy=(famla_idx + 2*width_ag, p2_conc[famla_idx]),
                xytext=(famla_idx, max(p2_conc)*0.85),
                fontsize=8, color=RED, ha='center', fontweight='bold',
                arrowprops=dict(arrowstyle='->', color=RED, lw=1.2))

plt.savefig(f'{OUT_DIR}/chartD_conc_top_agences.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart D: CONC top agences')

# === CHART E: 4 metrics comparison (2x2 grid) ===
fig, axes = plt.subplots(2, 2, figsize=(13, 8.5), constrained_layout=True)
x = np.arange(len(PERIODS_ORDER))

ax = axes[0, 0]
soja_vol = [EXT['soja'][p]['volume_t'] for p in PERIODS_ORDER]
conc_vol = [EXT['conc'][p]['volume_t'] for p in PERIODS_ORDER]
ax.plot(x, soja_vol, marker='o', linewidth=2.5, markersize=10, color=NAVY, label='Soja')
ax.plot(x, conc_vol, marker='s', linewidth=2.5, markersize=10, color=GOLD, label='Conc')
for i, (s, c) in enumerate(zip(soja_vol, conc_vol)):
    ax.annotate(f'{s:.0f}', xy=(i, s), xytext=(0, 10), textcoords='offset points',
                ha='center', fontsize=8, color=NAVY, fontweight='bold')
    ax.annotate(f'{c:.0f}', xy=(i, c), xytext=(0, -16), textcoords='offset points',
                ha='center', fontsize=8, color=GOLD, fontweight='bold')
ax.set_xticks(x); ax.set_xticklabels(PERIOD_LABELS, fontsize=8)
ax.set_title('Volume (tonnes)', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(fontsize=8, loc='upper right')
ax.grid(True, alpha=0.3); ax.set_axisbelow(True)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)

ax = axes[0, 1]
soja_cli = [EXT['soja'][p]['n_clients'] for p in PERIODS_ORDER]
conc_cli = [EXT['conc'][p]['n_clients'] for p in PERIODS_ORDER]
ax.plot(x, soja_cli, marker='o', linewidth=2.5, markersize=10, color=NAVY, label='Soja clients')
ax.plot(x, conc_cli, marker='s', linewidth=2.5, markersize=10, color=GOLD, label='Conc clients')
for i, (s, c) in enumerate(zip(soja_cli, conc_cli)):
    ax.annotate(f'{s}', xy=(i, s), xytext=(0, 10), textcoords='offset points',
                ha='center', fontsize=8, color=NAVY, fontweight='bold')
    ax.annotate(f'{c}', xy=(i, c), xytext=(0, -16), textcoords='offset points',
                ha='center', fontsize=8, color=GOLD, fontweight='bold')
ax.set_xticks(x); ax.set_xticklabels(PERIOD_LABELS, fontsize=8)
ax.set_title('Nombre de clients uniques', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(fontsize=8, loc='upper right')
ax.grid(True, alpha=0.3); ax.set_axisbelow(True)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)

ax = axes[1, 0]
soja_cmd = [EXT['soja'][p]['n_orders'] for p in PERIODS_ORDER]
conc_cmd = [EXT['conc'][p]['n_orders'] for p in PERIODS_ORDER]
ax.plot(x, soja_cmd, marker='o', linewidth=2.5, markersize=10, color=NAVY, label='Soja cmds')
ax.plot(x, conc_cmd, marker='s', linewidth=2.5, markersize=10, color=GOLD, label='Conc cmds')
for i, (s, c) in enumerate(zip(soja_cmd, conc_cmd)):
    ax.annotate(f'{s}', xy=(i, s), xytext=(0, 10), textcoords='offset points',
                ha='center', fontsize=8, color=NAVY, fontweight='bold')
    ax.annotate(f'{c}', xy=(i, c), xytext=(0, -16), textcoords='offset points',
                ha='center', fontsize=8, color=GOLD, fontweight='bold')
ax.set_xticks(x); ax.set_xticklabels(PERIOD_LABELS, fontsize=8)
ax.set_title('Nombre de commandes', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(fontsize=8, loc='upper right')
ax.grid(True, alpha=0.3); ax.set_axisbelow(True)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)

ax = axes[1, 1]
soja_pc = [EXT['soja'][p]['panier_client_kg'] for p in PERIODS_ORDER]
conc_pc = [EXT['conc'][p]['panier_client_kg'] for p in PERIODS_ORDER]
ax.plot(x, soja_pc, marker='o', linewidth=2.5, markersize=10, color=NAVY, label='Soja panier/cli')
ax.plot(x, conc_pc, marker='s', linewidth=2.5, markersize=10, color=GOLD, label='Conc panier/cli')
for i, (s, c) in enumerate(zip(soja_pc, conc_pc)):
    ax.annotate(f'{s:.0f}', xy=(i, s), xytext=(0, 10), textcoords='offset points',
                ha='center', fontsize=8, color=NAVY, fontweight='bold')
    ax.annotate(f'{c:.0f}', xy=(i, c), xytext=(0, -16), textcoords='offset points',
                ha='center', fontsize=8, color=GOLD, fontweight='bold')
ax.set_xticks(x); ax.set_xticklabels(PERIOD_LABELS, fontsize=8)
ax.set_title('Panier moyen par client (kg)', fontsize=11, fontweight='bold', color=NAVY)
ax.legend(fontsize=8, loc='upper right')
ax.grid(True, alpha=0.3); ax.set_axisbelow(True)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)

fig.suptitle('Comparaison SOJA vs CONCENTRÉS — 4 métriques sur 5 semaines (octobre)\n(Cumul W-3 → P2 : soja = + clients ET + paniers ; conc = + clients mais paniers en baisse)',
             fontsize=13, fontweight='bold', color=NAVY, y=1.02)

plt.savefig(f'{OUT_DIR}/chartE_4metrics.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart E: 4 metrics comparison')

# === CHART F: Cumul variations W-3 → P2 — by metric and family ===
fig, ax = plt.subplots(figsize=(11, 6), constrained_layout=True)

metrics_list = ['volume', 'clients', 'orders', 'panier_client', 'panier_order']
metric_labels = ['Volume', 'Nb clients', 'Nb commandes', 'Panier/client', 'Panier/cmd']

soja_cumul = [EXT['cumul_w3_to_p2']['soja'][m] for m in metrics_list]
conc_cumul = [EXT['cumul_w3_to_p2']['conc'][m] for m in metrics_list]

x_pos = np.arange(len(metrics_list))
width = 0.35

bars1 = ax.bar(x_pos - width/2, soja_cumul, width, label='SOJA', color=NAVY)
bars2 = ax.bar(x_pos + width/2, conc_cumul, width, label='CONCENTRÉS', color=GOLD)

for bar in bars1:
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2, h + (1 if h >= 0 else -3),
            f'{h:+.1f}%', ha='center', va='bottom' if h >= 0 else 'top',
            fontsize=9, fontweight='bold', color=NAVY)
for bar in bars2:
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2, h + (1 if h >= 0 else -3),
            f'{h:+.1f}%', ha='center', va='bottom' if h >= 0 else 'top',
            fontsize=9, fontweight='bold', color=GOLD)

ax.axhline(y=0, color='black', linewidth=0.8)
ax.set_xticks(x_pos); ax.set_xticklabels(metric_labels, fontsize=10)
ax.set_ylabel('Variation cumulée W-3 → P2 (%)', fontsize=10)
ax.set_title('Variation cumulée W-3 → P2 par métrique\n(SOJA vs CONCENTRÉS — octobre)',
             fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.legend(loc='lower right', fontsize=10)
ax.grid(True, alpha=0.3, axis='y')
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)

cum_soja_cli = EXT['cumul_w3_to_p2']['soja']['clients']
cum_soja_pc = EXT['cumul_w3_to_p2']['soja']['panier_client']
cum_conc_vol = EXT['cumul_w3_to_p2']['conc']['volume']
cum_conc_po = EXT['cumul_w3_to_p2']['conc']['panier_order']
ax.text(0.02, 0.95, f'Soja: tout en hausse\n(+{cum_soja_cli:+.0f} % clients, +{cum_soja_pc:+.0f} % paniers)\nConc: volumes +{cum_conc_vol:+.0f} % mais\npaniers en baisse ({cum_conc_po:+.0f} %)',
        transform=ax.transAxes, fontsize=9, color=RED, fontweight='bold',
        bbox=dict(boxstyle='round,pad=0.5', facecolor='#FCE4EC', edgecolor=RED, alpha=0.9),
        verticalalignment='top')

plt.savefig(f'{OUT_DIR}/chartF_cumul_variations.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart F: Cumul variations')

print("\nAll 12 charts regenerated (October data).")
