"""Compact, readable study-workflow flowchart (Figure 3)."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Polygon

TEAL, DTEAL, ORANGE, RED, PURPLE, GREY = "#007B7B", "#006263", "#E87722", "#C0504D", "#6B3FA0", "#7F7F7F"
plt.rcParams.update({"font.family": "Carlito"})

def box(ax, x, y, w, h, title, body, fc="#E6F2F2", ec=TEAL, tc=DTEAL):
    ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h, boxstyle="round,pad=0.3,rounding_size=1", fc=fc, ec=ec, lw=1.6, zorder=2))
    ax.text(x, y + h / 2 - 2.0, title, ha="center", va="center", fontsize=12.5, weight="bold", color=tc, zorder=3)
    ax.text(x, y - 1.3, body, ha="center", va="center", fontsize=10.5, color="#262626", zorder=3, linespacing=1.3)

def src(ax, x, y, text):
    ax.add_patch(FancyBboxPatch((x - 12, y - 3.6), 24, 7.2, boxstyle="round,pad=0.3,rounding_size=1", fc="#F2F2F2", ec="#A6A6A6", lw=1.1, zorder=2))
    ax.text(x, y, text, ha="center", va="center", fontsize=10, color="#262626", zorder=3, linespacing=1.25)

def arrow(ax, x1, y1, x2, y2, c=DTEAL, ls="-", lw=1.8):
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1), zorder=1,
                arrowprops=dict(arrowstyle="-|>", color=c, lw=lw, ls=ls, mutation_scale=15))

fig, ax = plt.subplots(figsize=(8.2, 9.0))
ax.set_xlim(0, 100); ax.set_ylim(0, 112); ax.axis("off")
cx, w, h = 42, 50, 9.6
steps = [
    (104, "1. Define the project", "100 MW wind + 50 MW PV + 30 MW/60 MWh BESS\nin the CWO REZ; 70% PPA + 30% spot market"),
    (91, "2. Technical analysis", "Weibull wind yield, PV performance ratio,\nbattery shifting, losses, MLF, degradation"),
    (78, "3. Financial analysis", "CAPEX, OPEX, revenue → 25-year cash flow;\nNPV, IRR, LCOE, payback, break-even PPA"),
    (65, "4. Environmental analysis", "CO₂ avoided = energy × NGA factor;\nland, biodiversity, visual and noise impacts"),
    (52, "5. Risk assessment", "Risk register: 12 risks in 6 categories;\nprobability × impact, mitigation"),
    (39, "6. Scenario and uncertainty analysis", "Price +20%, CAPEX −15%, rate +2 pp; tornado;\n5,000-run Monte Carlo; sizing study"),
]
for y, t, b in steps:
    box(ax, cx, y, w, h, t, b)
for (y1, *_), (y2, *_) in zip(steps[:-1], steps[1:]):
    arrow(ax, cx, y1 - h / 2 - 0.4, cx, y2 + h / 2 + 0.4)

# verification diamond
dy = 23.5
ax.add_patch(Polygon([(cx, dy + 7), (cx + 17, dy), (cx, dy - 7), (cx - 17, dy)], closed=True, fc="#FFF4E6", ec=ORANGE, lw=1.8, zorder=2))
ax.text(cx, dy + 2.2, "7. Verification", ha="center", va="center", fontsize=12.5, weight="bold", color="#9C4A00", zorder=3)
ax.text(cx, dy - 2.2, "Python = Excel? Hand checks?\nSources plausible?", ha="center", va="center", fontsize=10, zorder=3, linespacing=1.2)
arrow(ax, cx, 39 - h / 2 - 0.4, cx, dy + 7.3)

# loop back on failure
ax.plot([cx - 17, 7, 7], [dy, dy, 91], color=RED, lw=1.6, zorder=1)
arrow(ax, 7, 91, cx - w / 2 - 0.6, 91, c=RED, lw=1.6)
ax.text(5.2, 57, "No: correct inputs or model\n(wind CF 29.7% rejected; Excel error fixed)", rotation=90, ha="center", va="center", fontsize=9.5, color=RED, linespacing=1.2)

# decision
box(ax, cx, 6.5, w, 9.0, "8. Integrated decision", "Combine all results → INVESTIGATE FURTHER\n(proceed to development stage)", fc="#E6F2F2", ec=DTEAL)
arrow(ax, cx, dy - 7.3, cx, 6.5 + 4.5 + 0.4)
ax.text(cx + 1.2, 13.6, "Yes", fontsize=10.5, color=DTEAL, weight="bold")

# data sources
sx = 86
ax.text(sx, 110.3, "Data sources", ha="center", fontsize=11.5, weight="bold", color=GREY)
for y, t, ty in [(102.5, "BoM, Global Wind Atlas\n(resource)", 91), (91, "AEMO MLF, NER S5.2.5\n(grid and GPS)", 91), (78, "CSIRO GenCost 2025-26\n(costs)", 78),
                 (65, "DCCEEW NGA 2025\n(emission factor)", 65), (52, "AEMO ISP, NEM prices,\nCIS/LTESA", 52)]:
    src(ax, sx, y, t)
    arrow(ax, sx - 12.4, y, cx + w / 2 + 0.6, ty if abs(ty - y) < 1 else ty + 2, c="#8C8C8C", ls="--", lw=1.3)

# AI support
ax.add_patch(FancyBboxPatch((74, 14), 24, 18, boxstyle="round,pad=0.3,rounding_size=1", fc="#F6F0FB", ec=PURPLE, lw=1.3, ls="--", zorder=2))
ax.text(86, 29.6, "AI support (Claude)", ha="center", fontsize=11.5, weight="bold", color=PURPLE)
ax.text(86, 21.6, "Drafts model code\nProposes scenarios\nSearches benchmarks\nDrafts report text\nAuthors verify all (step 7)", ha="center", va="center", fontsize=10, linespacing=1.3)
arrow(ax, 73.6, 23.5, cx + 17.4, dy, c=PURPLE, ls="--", lw=1.4)

fig.tight_layout(pad=0.3)
fig.savefig("fig_workflow.png", dpi=220)
print("saved")
