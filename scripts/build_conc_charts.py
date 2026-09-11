"""Generate charts focused on CONCENTRÉS decline trend.
The CONCENTRÉS is on a clear downward trend (76.3 → 51.6 → 48.0 → 44.9 t/j),
contrasting with SOJA which oscillates (no clear trend).
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

# === Colors ===
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

# === Compute CONC trend indicators ===
conc_trend = [DATA['global'][p]['conc_t'] for p in PERIODS_ORDER]
conc_moy_trend = [DATA['global'][p]['conc_moy_j'] for p in PERIODS_ORDER]
soja_moy_trend = [DATA['global'][p]['soja_moy_j'] for p in PERIODS_ORDER]

# Variations week-over-week for CONC
conc_vars = []
for i in range(1, len(conc_moy_trend)):
    var = (conc_moy_trend[i]/conc_moy_trend[i-1] - 1)*100
    conc_vars.append(var)

# Total variation W-3 → P1
total_var_conc = (conc_moy_trend[-1]/conc_moy_trend[0] - 1)*100
total_var_soja = (soja_moy_trend[-1]/soja_moy_trend[0] - 1)*100

print(f"CONC trend (t/j): {conc_moy_trend}")
print(f"CONC weekly variations: {conc_vars}")
print(f"CONC total W-3 → P1: {total_var_conc:.1f}%")
print(f"SOJA total W-3 → P1: {total_var_soja:.1f}%")

# === CHART A: CONCENTRÉS focus — Decline trend ===
fig, ax = plt.subplots(figsize=(10, 5.5), constrained_layout=True)
x = np.arange(len(PERIODS_ORDER))

# Bar chart for CONC volumes
bars = ax.bar(x, conc_trend, 0.5, color=GOLD, edgecolor=NAVY, linewidth=1.2)

# Add value labels
for i, bar in enumerate(bars):
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2, h + 8, f'{h:.1f} t',
            ha='center', va='bottom', fontsize=10, fontweight='bold', color=NAVY)

# Add variation arrows between bars
for i in range(1, len(conc_trend)):
    prev_h = conc_trend[i-1]
    curr_h = conc_trend[i]
    mid_x = (i-1 + i) / 2
    mid_y = max(prev_h, curr_h) + 30
    
    # Arrow from prev to curr
    color = RED if curr_h < prev_h else GREEN
    ax.annotate('', xy=(i, curr_h + 5), xytext=(i-1, prev_h + 5),
                arrowprops=dict(arrowstyle='->', color=color, lw=1.5, alpha=0.7))
    
    # Variation label
    var = (curr_h/prev_h - 1)*100
    ax.text(mid_x, mid_y + 15, f'{var:+.1f}%', ha='center', fontsize=10, 
            fontweight='bold', color=color,
            bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor=color, alpha=0.9))

# Total variation W-3 → P1
total_var = (conc_trend[-1]/conc_trend[0] - 1)*100
ax.text(0.02, 0.95, f'Baisse cumulative\nW-3 → P1: {total_var:.1f}%', 
        transform=ax.transAxes, fontsize=10, fontweight='bold', color=RED,
        bbox=dict(boxstyle='round,pad=0.5', facecolor='#FCE4EC', edgecolor=RED, alpha=0.9),
        verticalalignment='top')

ax.set_xticks(x)
ax.set_xticklabels(PERIOD_LABELS, fontsize=9)
ax.set_ylabel('Volume Concentrés (tonnes)', fontsize=10)
ax.set_title('CONCENTRÉS — Tendance baissière continue sur 4 semaines\n(baisse cumulée -41,1% de W-3 à P1)', 
             fontsize=12, fontweight='bold', color=RED, pad=10)
ax.grid(True, alpha=0.3, axis='y')
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_ylim(0, max(conc_trend) * 1.3)

plt.savefig(f'{OUT_DIR}/chartA_conc_decline.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart A: CONCENTRÉS decline')

# === CHART B: Comparison SOJA vs CONC — Different dynamics ===
fig, ax = plt.subplots(figsize=(11, 6), constrained_layout=True)

# Normalize to W-3 = 100 (index)
soja_index = [v/soja_moy_trend[0]*100 for v in soja_moy_trend]
conc_index = [v/conc_moy_trend[0]*100 for v in conc_moy_trend]

ax.plot(x, soja_index, marker='o', linewidth=3, markersize=12, color=NAVY, label='Soja (TOURTEAUX) moy/j')
ax.plot(x, conc_index, marker='s', linewidth=3, markersize=12, color=GOLD, label='Concentrés moy/j', linestyle='--')

# Reference line at 100 (W-3 baseline)
ax.axhline(y=100, color=GRAY, linestyle=':', alpha=0.6, linewidth=1)
ax.text(3.1, 100, 'Base W-3 = 100', fontsize=8, color=GRAY, va='center')

# Add value labels
for i, (s, c) in enumerate(zip(soja_index, conc_index)):
    ax.annotate(f'{s:.0f}', xy=(i, s), xytext=(0, 12), textcoords='offset points',
                ha='center', fontsize=10, fontweight='bold', color=NAVY)
    ax.annotate(f'{c:.0f}', xy=(i, c), xytext=(0, -20), textcoords='offset points',
                ha='center', fontsize=10, fontweight='bold', color=GOLD)

# Highlight divergence
ax.fill_between(x, soja_index, conc_index, where=[s >= c for s, c in zip(soja_index, conc_index)],
                color=RED, alpha=0.1, label='Divergence (Soja > Conc)')

# Annotations for key insights
ax.annotate('Soja: oscillations\n(pas de tendance claire)', 
            xy=(2, soja_index[2]), xytext=(1.5, 130),
            fontsize=9, color=NAVY, ha='center',
            arrowprops=dict(arrowstyle='->', color=NAVY, lw=1.2))
ax.annotate('Conc: baisse continue\n(-41% cumulé)', 
            xy=(3, conc_index[3]), xytext=(2.5, 60),
            fontsize=9, color=RED, ha='center', fontweight='bold',
            arrowprops=dict(arrowstyle='->', color=RED, lw=1.5))

ax.set_xticks(x)
ax.set_xticklabels(PERIOD_LABELS, fontsize=9)
ax.set_ylabel('Indice (base W-3 = 100)', fontsize=10)
ax.set_title('Dynamiques divergentes : SOJA (oscillations) vs CONCENTRÉS (baisse continue)\nCadence quotidienne indexée à 100 en W-3', 
             fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.legend(loc='upper right', fontsize=9)
ax.grid(True, alpha=0.3)
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_ylim(40, 140)

plt.savefig(f'{OUT_DIR}/chartB_divergence.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart B: Divergence SOJA vs CONC')

# === CHART C: CONC by region — Decline per region ===
fig, ax = plt.subplots(figsize=(10, 5.5), constrained_layout=True)

regions = ['Ouest', 'Centre', 'Littoral']
colors_reg = [NAVY, GOLD, ORANGE]
markers = ['o', 's', 'D']

for i, (region, color, marker) in enumerate(zip(regions, colors_reg, markers)):
    conc_vals = [DATA['regions'][region][p]['conc_t'] for p in PERIODS_ORDER]
    work_days = 5
    conc_moy = [v/work_days for v in conc_vals]
    
    ax.plot(x, conc_moy, marker=marker, linewidth=2.5, markersize=10, color=color, label=region)
    
    # Add value labels for last point (P1)
    ax.annotate(f'{conc_moy[-1]:.1f}', xy=(3, conc_moy[-1]), xytext=(8, 0),
                textcoords='offset points', fontsize=9, color=color, fontweight='bold')
    # First point (W-3)
    ax.annotate(f'{conc_moy[0]:.1f}', xy=(0, conc_moy[0]), xytext=(-15, 0),
                textcoords='offset points', fontsize=9, color=color, fontweight='bold',
                ha='right')

ax.set_xticks(x)
ax.set_xticklabels(PERIOD_LABELS, fontsize=9)
ax.set_ylabel('Cadence Conc moy/j (tonnes)', fontsize=10)
ax.set_title('CONCENTRÉS par région — Baisse généralisée sur 4 semaines\n(Ouest, Centre, Littoral tous en recul)', 
             fontsize=12, fontweight='bold', color=RED, pad=10)
ax.legend(loc='upper right', fontsize=9)
ax.grid(True, alpha=0.3)
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.set_ylim(0, 40)

plt.savefig(f'{OUT_DIR}/chartC_conc_regional.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart C: CONC regional')

# === CHART D: CONC by top 10 agences ===
fig, ax = plt.subplots(figsize=(12, 6), constrained_layout=True)

top_agences = list(DATA['agences'].items())[:10]
ag_names = [ag[0] for ag in top_agences]
p1_conc = [ag[1]['P1']['conc_t'] for ag in top_agences]
w1_conc = [ag[1]['W-1']['conc_t'] for ag in top_agences]
w2_conc = [ag[1]['W-2']['conc_t'] for ag in top_agences]
w3_conc = [ag[1]['W-3']['conc_t'] for ag in top_agences]

x_ag = np.arange(len(ag_names))
width_ag = 0.2

ax.bar(x_ag - 1.5*width_ag, w3_conc, width_ag, label='W-3 (15-20/08)', color='#A8C5E2')
ax.bar(x_ag - 0.5*width_ag, w2_conc, width_ag, label='W-2 (22-27/08)', color=LIGHT_BLUE)
ax.bar(x_ag + 0.5*width_ag, w1_conc, width_ag, label='W-1 (29/08-03/09)', color=NAVY)
ax.bar(x_ag + 1.5*width_ag, p1_conc, width_ag, label='P1 (05-10/09)', color=GOLD)

ax.set_xticks(x_ag)
ax.set_xticklabels(ag_names, fontsize=9, rotation=30, ha='right')
ax.set_ylabel('Volume Concentrés (tonnes)', fontsize=10)
ax.set_title('Top 10 agences — Évolution du volume CONCENTRÉS sur 4 semaines\n(baisse généralisée, particulièrement FAMLA)', 
             fontsize=12, fontweight='bold', color=RED, pad=10)
ax.legend(loc='upper right', fontsize=8, ncol=2)
ax.grid(True, alpha=0.3, axis='y')
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# Highlight FAMLA decline
famla_idx = ag_names.index('Famla')
if famla_idx >= 0:
    ax.annotate(f'FAMLA: {w3_conc[famla_idx]:.0f} → {p1_conc[famla_idx]:.0f} t\n({(p1_conc[famla_idx]/w3_conc[famla_idx]-1)*100:+.0f}%)', 
                xy=(famla_idx + 1.5*width_ag, p1_conc[famla_idx]),
                xytext=(famla_idx, max(p1_conc)*0.85),
                fontsize=8, color=RED, ha='center', fontweight='bold',
                arrowprops=dict(arrowstyle='->', color=RED, lw=1.2))

plt.savefig(f'{OUT_DIR}/chartD_conc_top_agences.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart D: CONC top agences')

print(f"\nAll 4 CONC-focused charts saved in {OUT_DIR}/")
print(f"\n=== KEY INSIGHT ===")
print(f"SOJA total W-3 → P1: {total_var_soja:+.1f}% (oscillations, pas de tendance)")
print(f"CONC total W-3 → P1: {total_var_conc:+.1f}% (BAISSE CONTINUE)")
