"""Generate charts showing clients, orders, and average basket trends.
Answers the user's question: do these metrics confirm that SOJA is "stable"?
"""
import json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
import os

fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf')
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

NAVY = '#1F3A5F'
GOLD = '#C9A961'
RED = '#C00000'
GREEN = '#548235'
ORANGE = '#ED7D31'
GRAY = '#595959'
LIGHT_BLUE = '#5B9BD5'

# Load extended data
DATA = json.load(open('/home/z/my-project/scripts/sept_trend_extended.json'))
# Periods in chronological order (W-3, W-2, W-1, P1)
PERIODS_CHR = ['W-3', 'W-2', 'W-1', 'P1']
PERIOD_LABELS = ['W-3\n(15-20/08)', 'W-2\n(22-27/08)', 'W-1\n(29/08-03/09)', 'P1\n(05-10/09)']

OUT_DIR = '/home/z/my-project/scripts/charts'
os.makedirs(OUT_DIR, exist_ok=True)

# === CHART E: 4 metrics comparison (2x2 grid) ===
fig, axes = plt.subplots(2, 2, figsize=(13, 8.5), constrained_layout=True)
x = np.arange(len(PERIODS_CHR))

# Subplot 1: Volume
ax = axes[0, 0]
soja_vol = [DATA['soja'][p]['volume_t'] for p in PERIODS_CHR]
conc_vol = [DATA['conc'][p]['volume_t'] for p in PERIODS_CHR]
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

# Subplot 2: Number of clients
ax = axes[0, 1]
soja_cli = [DATA['soja'][p]['n_clients'] for p in PERIODS_CHR]
conc_cli = [DATA['conc'][p]['n_clients'] for p in PERIODS_CHR]
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

# Subplot 3: Number of orders
ax = axes[1, 0]
soja_cmd = [DATA['soja'][p]['n_orders'] for p in PERIODS_CHR]
conc_cmd = [DATA['conc'][p]['n_orders'] for p in PERIODS_CHR]
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

# Subplot 4: Average basket per client (kg)
ax = axes[1, 1]
soja_pc = [DATA['soja'][p]['panier_client_kg'] for p in PERIODS_CHR]
conc_pc = [DATA['conc'][p]['panier_client_kg'] for p in PERIODS_CHR]
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

fig.suptitle('Comparaison SOJA vs CONCENTRÉS — 4 métriques sur 4 semaines\n(Les nouvelles métriques confirment-elles la "stabilité" du SOJA ?)', 
             fontsize=13, fontweight='bold', color=NAVY, y=1.02)

plt.savefig(f'{OUT_DIR}/chartE_4metrics.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart E: 4 metrics comparison')

# === CHART F: Cumul variations W-3 → P1 — by metric and family ===
fig, ax = plt.subplots(figsize=(11, 6), constrained_layout=True)

metrics_list = ['volume', 'clients', 'orders', 'panier_client', 'panier_order']
metric_labels = ['Volume', 'Nb clients', 'Nb commandes', 'Panier/client', 'Panier/cmd']

soja_cumul = [DATA['cumul_w3_to_p1']['soja'][m] for m in metrics_list]
conc_cumul = [DATA['cumul_w3_to_p1']['conc'][m] for m in metrics_list]

x_pos = np.arange(len(metrics_list))
width = 0.35

bars1 = ax.bar(x_pos - width/2, soja_cumul, width, label='SOJA', color=NAVY)
bars2 = ax.bar(x_pos + width/2, conc_cumul, width, label='CONCENTRÉS', color=GOLD)

# Add value labels
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
ax.set_ylabel('Variation cumulée W-3 → P1 (%)', fontsize=10)
ax.set_title('Variation cumulée W-3 → P1 par métrique\n(SOJA vs CONCENTRÉS)', 
             fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.legend(loc='lower right', fontsize=10)
ax.grid(True, alpha=0.3, axis='y')
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)

# Annotation
ax.text(0.02, 0.95, 'Constat: SOJA est AUSSI en baisse\nsur toutes les métriques\n(le "stable" était faux)', 
        transform=ax.transAxes, fontsize=9, color=RED, fontweight='bold',
        bbox=dict(boxstyle='round,pad=0.5', facecolor='#FCE4EC', edgecolor=RED, alpha=0.9),
        verticalalignment='top')

plt.savefig(f'{OUT_DIR}/chartF_cumul_variations.png', dpi=120, bbox_inches='tight')
plt.close()
print('✓ Chart F: Cumul variations')

print("\n=== KEY FINDING ===")
print("With corrected data (excluding internal clients),")
print(f"SOJA cumulative W-3 → P1: Volume {DATA['cumul_w3_to_p1']['soja']['volume']:+.1f}%")
print(f"SOJA cumulative W-3 → P1: Clients {DATA['cumul_w3_to_p1']['soja']['clients']:+.1f}%")
print(f"SOJA cumulative W-3 → P1: Commandes {DATA['cumul_w3_to_p1']['soja']['orders']:+.1f}%")
print(f"SOJA cumulative W-3 → P1: Panier/client {DATA['cumul_w3_to_p1']['soja']['panier_client']:+.1f}%")
print()
print(f"CONC cumulative W-3 → P1: Volume {DATA['cumul_w3_to_p1']['conc']['volume']:+.1f}%")
print(f"CONC cumulative W-3 → P1: Clients {DATA['cumul_w3_to_p1']['conc']['clients']:+.1f}%")
print(f"CONC cumulative W-3 → P1: Commandes {DATA['cumul_w3_to_p1']['conc']['orders']:+.1f}%")
print(f"CONC cumulative W-3 → P1: Panier/client {DATA['cumul_w3_to_p1']['conc']['panier_client']:+.1f}%")
