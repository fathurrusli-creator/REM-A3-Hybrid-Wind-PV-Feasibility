"""Draws the system single-line diagram and the study workflow flowchart for the report."""
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Rectangle, FancyBboxPatch, Polygon

NAVY, BLUE, ORANGE, GREEN, GREY, RED = "#006263", "#007B7B", "#E87722", "#3C9D9B", "#595959", "#C0504D"
plt.rcParams.update({"font.family": "calibri", "font.size": 9.5})


# ------------------------------------------------------------------ symbols
def line(ax, x1, y1, x2, y2, c="k", lw=1.6, ls="-", z=1):
    ax.plot([x1, x2], [y1, y2], color=c, lw=lw, ls=ls, zorder=z, solid_capstyle="butt")


def breaker(ax, x, y, s=1.6):
    ax.add_patch(Rectangle((x - s / 2, y - s / 2), s, s, fc="white", ec="k", lw=1.4, zorder=3))


def transformer(ax, x, y, r=2.0, lab=None, side="right"):
    ax.add_patch(Circle((x, y + r * 0.62), r, fc="none", ec="k", lw=1.4, zorder=3))
    ax.add_patch(Circle((x, y - r * 0.62), r, fc="none", ec="k", lw=1.4, zorder=3))
    if lab:
        dx = r + 1.0 if side == "right" else -(r + 1.0)
        ax.text(x + dx, y, lab, ha="left" if side == "right" else "right", va="center", fontsize=7.5, color=GREY)


def inverter(ax, x, y, s=6, lab="", fc="white"):
    ax.add_patch(Rectangle((x - s / 2, y - s / 2), s, s, fc=fc, ec="k", lw=1.4, zorder=3))
    line(ax, x - s / 2, y - s / 2, x + s / 2, y + s / 2, lw=1, z=4)
    ax.text(x - s / 4.2, y + s / 4.6, "~", ha="center", va="center", fontsize=11, zorder=5)
    ax.text(x + s / 4.2, y - s / 4.0, "=", ha="center", va="center", fontsize=11, zorder=5)
    if lab:
        ax.text(x + s / 2 + 1, y, lab, ha="left", va="center", fontsize=7.5, color=GREY)


def pv_array(ax, x, y, w=9, h=6):
    ax.add_patch(Rectangle((x - w / 2, y - h / 2), w, h, fc="#FDF1D6", ec="k", lw=1.4, zorder=3))
    for i in range(1, 3):
        line(ax, x - w / 2 + i * w / 3, y - h / 2, x - w / 2 + i * w / 3, y + h / 2, c=GREY, lw=0.6, z=4)
    line(ax, x - w / 2, y, x + w / 2, y, c=GREY, lw=0.6, z=4)
    # sunlight arrows
    for k in (-1.8, 0.6):
        ax.annotate("", xy=(x + k, y + h / 2 + 0.3), xytext=(x + k - 2.2, y + h / 2 + 3.2),
                    arrowprops=dict(arrowstyle="-|>", color=ORANGE, lw=1.3), zorder=5)


def battery(ax, x, y):
    for i, (w, lw) in enumerate([(6, 1.6), (3.4, 3.2), (6, 1.6), (3.4, 3.2)]):
        yy = y + 2.2 - i * 1.45
        line(ax, x - w / 2, yy, x + w / 2, yy, lw=lw, z=4)
    ax.text(x + 3.8, y + 2.5, "+", fontsize=9)
    ax.text(x + 3.8, y - 2.4, "−", fontsize=9)


def turbine(ax, x, y, h=11):
    line(ax, x, y, x, y + h, lw=2.2, c=GREY)
    hub = (x, y + h)
    for ang in (90, 210, 330):
        a = np.radians(ang)
        tip = (hub[0] + 6.0 * np.cos(a), hub[1] + 6.0 * np.sin(a))
        perp = np.array([-np.sin(a), np.cos(a)]) * 0.75
        ax.add_patch(Polygon([hub, (hub[0] + 1.8 * np.cos(a) + perp[0], hub[1] + 1.8 * np.sin(a) + perp[1]), tip],
                             closed=True, fc=BLUE, ec=NAVY, lw=0.8, zorder=4))
    ax.add_patch(Circle(hub, 0.75, fc=NAVY, ec=NAVY, zorder=5))
    ax.add_patch(Circle((x, y), 0.9, fc="none", ec="k", lw=1.3, zorder=3))
    ax.text(x, y, "G", ha="center", va="center", fontsize=6.5, zorder=6)


def meter(ax, x, y, r=1.7):
    ax.add_patch(Circle((x, y), r, fc="white", ec="k", lw=1.4, zorder=3))
    ax.text(x, y, "Wh", ha="center", va="center", fontsize=6.5, zorder=4)


def tower(ax, x, y, h=7):
    pts = [(x - 2.2, y), (x - 0.5, y + h), (x + 0.5, y + h), (x + 2.2, y)]
    ax.plot(*zip(*pts), color=GREY, lw=1.2)
    for f in (0.35, 0.65):
        line(ax, x - 2.2 + 1.7 * f, y + h * f, x + 2.2 - 1.7 * f, y + h * f, c=GREY, lw=0.9)
    line(ax, x - 2.8, y + h - 0.6, x + 2.8, y + h - 0.6, c=GREY, lw=1.2)


# ------------------------------------------------------------------ system diagram
def system_diagram(path="fig_system.png"):
    fig, ax = plt.subplots(figsize=(11, 8.6))
    ax.set_xlim(0, 112)
    ax.set_ylim(-9, 96)
    ax.axis("off")

    # grid
    line(ax, 30, 84, 80, 84, lw=5, c=NAVY)
    tower(ax, 36, 84.5)
    tower(ax, 74, 84.5)
    ax.text(55, 94, "NEM – Central-West Orana REZ 330 kV network (EnergyCo / Transgrid)", ha="center", fontsize=9.5, weight="bold", color=NAVY)
    ax.text(82, 84, "330 kV", va="center", fontsize=8, color=NAVY)

    # connection point
    x0 = 55
    line(ax, x0, 84, x0, 47)
    meter(ax, x0, 77.5)
    ax.text(x0 + 2.6, 77.5, "Connection point & revenue meter\n(NEM settlement × MLF 0.93; 150 MW access right)", va="center", fontsize=7.5, color=GREY)
    breaker(ax, x0, 71.5)
    ax.text(x0 - 2, 71.5, "CB", ha="right", va="center", fontsize=7.5, color=GREY)
    transformer(ax, x0, 63, r=2.6, lab="Main transformer\n330/33 kV, 180 MVA", side="left")
    breaker(ax, x0, 54.5)

    # 33 kV bus
    line(ax, 12, 47, 100, 47, lw=5, c="k")
    ax.text(101, 47, "33 kV\ncollector bus", va="center", fontsize=8)

    feeders = {"wind": 22, "pv": 55, "bess": 88}
    for x in feeders.values():
        line(ax, x, 47, x, 26)
        breaker(ax, x, 42)
        transformer(ax, x, 34.5, r=2.0)

    # wind feeder
    xw = feeders["wind"]
    ax.text(xw + 3, 34.5, "0.69/33 kV\nunit transformers", va="center", fontsize=7.5, color=GREY)
    # collector cable routed around the turbines, connecting at each WTG base
    line(ax, xw, 26, xw - 14, 26)
    line(ax, xw - 14, 26, xw - 14, 4)
    line(ax, xw - 14, 4, xw + 9, 4)
    for dx in (-9, 0, 9):
        line(ax, xw + dx, 4, xw + dx, 8.1)
        turbine(ax, xw + dx, 9, h=10)
    ax.text(xw, -0.5, "WIND FARM\n20 × 5.0 MW WTG = 100 MW\n(4 collector feeders; 3 WTG shown)", ha="center", va="top", fontsize=8.2, weight="bold", color=NAVY)

    # PV feeder
    xp = feeders["pv"]
    ax.text(xp + 3, 34.5, "0.69/33 kV\ninverter stations", va="center", fontsize=7.5, color=GREY)
    line(ax, xp, 26, xp, 22)
    inverter(ax, xp, 19, lab="DC/AC inverters\n50 MWac")
    line(ax, xp, 16, xp, 9.5)
    pv_array(ax, xp, 6.5)
    ax.text(xp, -0.5, "SOLAR PV\n62.5 MWp DC / 50 MWac\nsingle-axis tracking (DC/AC 1.25)", ha="center", va="top", fontsize=8.2, weight="bold", color="#B85A12")

    # BESS feeder
    xb = feeders["bess"]
    ax.text(xb + 3, 34.5, "0.69/33 kV\nPCS transformer", va="center", fontsize=7.5, color=GREY)
    line(ax, xb, 26, xb, 22)
    inverter(ax, xb, 19, lab="Grid-forming\nPCS 30 MW", fc="#E6F2F2")
    line(ax, xb, 16, xb, 9.5)
    battery(ax, xb, 6.5)
    ax.text(xb, -0.5, "BESS (LFP)\n30 MW / 60 MWh, 2 h\nRTE 85%", ha="center", va="top", fontsize=8.2, weight="bold", color=GREEN)

    # PPC
    ax.add_patch(FancyBboxPatch((83, 62), 22, 8, boxstyle="round,pad=0.4", fc="#F2F2F2", ec=GREY, lw=1, ls="--"))
    ax.text(94, 66, "Power Plant Controller\n(GPS, voltage & P/Q control,\nAEMO dispatch / SCADA)", ha="center", va="center", fontsize=7.3)
    for x in feeders.values():
        line(ax, 94, 62, 94, 55, c=GREY, lw=0.9, ls=":")
    line(ax, feeders["wind"] + 6, 55, feeders["bess"] + 6, 55, c=GREY, lw=0.9, ls=":")
    for x in feeders.values():
        line(ax, x + 6, 55, x + 6, 22 if x != feeders["wind"] else 30, c=GREY, lw=0.9, ls=":")

    # legend
    lx, ly = 3, 80
    ax.add_patch(Rectangle((lx - 1.5, ly - 23), 23, 26.5, fc="white", ec="#BFBFBF", lw=0.8))
    ax.text(lx, ly + 1.5, "Symbols", weight="bold", fontsize=8)
    breaker(ax, lx + 1.5, ly - 2.5, s=1.4); ax.text(lx + 4, ly - 2.5, "Circuit breaker", va="center", fontsize=7.5)
    transformer(ax, lx + 1.5, ly - 7.5, r=1.3); ax.text(lx + 4, ly - 7.5, "Transformer", va="center", fontsize=7.5)
    inverter(ax, lx + 1.5, ly - 12.8, s=3.2); ax.text(lx + 4, ly - 12.8, "Inverter (DC/AC)", va="center", fontsize=7.5)
    meter(ax, lx + 1.5, ly - 17.5, r=1.4); ax.text(lx + 4, ly - 17.5, "Energy meter", va="center", fontsize=7.5)
    line(ax, lx, ly - 21.5, lx + 3, ly - 21.5, c=GREY, lw=0.9, ls=":"); ax.text(lx + 4, ly - 21.5, "Control / comms", va="center", fontsize=7.5)

    fig.tight_layout()
    fig.savefig(path, dpi=200, bbox_inches="tight")
    plt.close(fig)


# ------------------------------------------------------------------ flowchart
def box(ax, x, y, w, h, text, fc, ec=NAVY, bold_first=True, fs=8.3):
    ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h, boxstyle="round,pad=0.35,rounding_size=0.8", fc=fc, ec=ec, lw=1.3, zorder=2))
    lines = text.split("\n")
    if bold_first:
        ax.text(x, y + (len(lines) - 1) * 1.05, lines[0], ha="center", va="center", fontsize=fs + 0.6, weight="bold", zorder=3, color=NAVY)
        if len(lines) > 1:
            ax.text(x, y - 1.0, "\n".join(lines[1:]), ha="center", va="center", fontsize=fs - 0.4, zorder=3, linespacing=1.25)
    else:
        ax.text(x, y, text, ha="center", va="center", fontsize=fs, zorder=3)


def arrow(ax, x1, y1, x2, y2, c=NAVY, ls="-", rad=0.0, lab=None):
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1), zorder=1,
                arrowprops=dict(arrowstyle="-|>", color=c, lw=1.4, ls=ls, connectionstyle=f"arc3,rad={rad}", mutation_scale=13))
    if lab:
        ax.text((x1 + x2) / 2 + 0.8, (y1 + y2) / 2, lab, fontsize=7.5, color=c, va="center")


def flowchart(path="fig_workflow.png"):
    fig, ax = plt.subplots(figsize=(10, 12.5))
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 125)
    ax.axis("off")
    cx, w = 50, 40
    steps = [
        (117, "1. Define the project\nSite (CWO REZ), sizing 100 MW wind + 50 MW PV\n+ 30 MW/60 MWh BESS, PPA/merchant structure", "#E6F2F2"),
        (102, "2. Technical analysis\nWeibull wind yield, PV performance ratio,\nBESS shifting, losses, MLF, degradation", "#E6F2F2"),
        (87, "3. Financial analysis\nCAPEX, OPEX, revenue → 25-yr cash flow\nNPV, IRR, LCOE, payback, break-even PPA", "#E6F2F2"),
        (72, "4. Environmental analysis\nCO₂ avoided = energy × NGA factor;\nland, biodiversity, visual & noise impacts", "#E6F2F2"),
        (57, "5. Risk assessment\nRisk register: 12 risks × 6 categories,\nprobability × impact, mitigation", "#E6F2F2"),
        (42, "6. Scenario & uncertainty analysis\nPrice +20%, CAPEX −15%, rate +2 pp;\ntornado sensitivity; 5,000-run Monte Carlo", "#E6F2F2"),
    ]
    for y, t, fc in steps:
        box(ax, cx, y, w, 10, t, fc)
    for (y1, *_), (y2, *_) in zip(steps[:-1], steps[1:]):
        arrow(ax, cx, y1 - 5.4, cx, y2 + 5.4)

    # decision diamond
    dy = 26
    ax.add_patch(Polygon([(cx, dy + 7), (cx + 15, dy), (cx, dy - 7), (cx - 15, dy)], closed=True, fc="#FFF4D6", ec=ORANGE, lw=1.5, zorder=2))
    ax.text(cx, dy + 1.2, "7. Verification", ha="center", va="center", weight="bold", fontsize=8.6, color="#7A5200")
    ax.text(cx, dy - 2.2, "Python = Excel? Hand checks\nand sources plausible?", ha="center", va="center", fontsize=7.4)
    arrow(ax, cx, 42 - 5.4, cx, dy + 7.2)
    # loop back
    ax.plot([cx - 15, 8, 8], [dy, dy, 102], color=RED, lw=1.4, zorder=1)
    arrow(ax, 8, 102, cx - w / 2 - 0.6, 102, c=RED)
    ax.text(9, 64, "No → correct inputs / model\n(e.g. wind CF 29.7% rejected,\nExcel #NUM! error fixed)", fontsize=7.4, color=RED, rotation=90, va="center", ha="left")

    # recommendation
    box(ax, cx, 9, w, 9.5, "8. Integrated decision\nCombine technical, financial, environmental,\nrisk & scenario results → INVESTIGATE FURTHER", "#E6F2F2", ec=NAVY)
    arrow(ax, cx, dy - 7.2, cx, 9 + 5.2, lab="Yes")

    # side: data sources
    sx = 88
    ax.text(sx, 121.5, "Data sources", ha="center", weight="bold", fontsize=8.6, color=GREY)
    srcs = [(114, "BoM, Global Wind Atlas\n(resource)", 102), (101, "AEMO MLF, NER S5.2.5\n(grid & GPS)", 102), (88, "CSIRO GenCost\n2025-26 (costs)", 87),
            (75, "DCCEEW NGA 2025\n(emission factor)", 72), (62, "AEMO ISP, NEM\nprices, CIS/LTESA", 57)]
    for y, t, ty in srcs:
        box(ax, sx, y, 20, 7, t, "#F2F2F2", ec="#A6A6A6", bold_first=False, fs=7.4)
        arrow(ax, sx - 10.4, y, cx + w / 2 + 0.6, ty, c="#8C8C8C", ls="--")

    # side: AI support
    ax.add_patch(FancyBboxPatch((76, 18), 22, 26, boxstyle="round,pad=0.4", fc="#F7F0FA", ec="#7B4F9E", lw=1.1, ls="--"))
    ax.text(87, 41.5, "AI support (Claude)", ha="center", weight="bold", fontsize=8.4, color="#5B2C83")
    ax.text(87, 30, "• Drafts model code\n• Proposes scenarios &\n  distributions\n• Web search for\n  benchmarks\n• Drafts report text\n\nAuthor checks every\noutput (step 7)",
            ha="center", va="center", fontsize=7.4, linespacing=1.25)
    arrow(ax, 75.6, 30, cx + 15.3, dy + 0.5, c="#7B4F9E", ls="--")

    fig.tight_layout()
    fig.savefig(path, dpi=200, bbox_inches="tight")
    plt.close(fig)


if __name__ == "__main__":
    system_diagram()
    flowchart()
    print("saved")
