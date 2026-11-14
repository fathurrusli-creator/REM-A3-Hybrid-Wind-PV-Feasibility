"""Adds numbered callouts and a legend to the AI-generated conceptual illustration."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Circle
from matplotlib.lines import Line2D
from PIL import Image

img = Image.open("in/gemini_system.jpg")
W, H = img.size
NAVY = "#006263"

# (circle position, target point)
callouts = {
    1: ((330, 95), (222, 350)),     # wind turbine nacelle
    2: ((150, 1060), (480, 1037)),   # PV tracker rows
    3: ((1436, 800), (1185, 955)),  # PV inverter station
    4: ((2470, 955), (2230, 860)),  # BESS containers
    5: ((1560, 420), (1700, 590)),  # substation transformer
    6: ((2130, 120), (2262, 290)),  # 330 kV tower
}
legend = [
    ("1", "Wind farm: 20 × 5.0 MW turbines (100 MW)"),
    ("2", "Solar PV: 62.5 MWp / 50 MWac single-axis trackers"),
    ("3", "PV inverter / MV transformer stations"),
    ("4", "Grid-forming BESS: 30 MW / 60 MWh (LFP)"),
    ("5", "330/33 kV substation, control room & PPC"),
    ("6", "330 kV transmission line to the NEM (CWO REZ)"),
]

fig = plt.figure(figsize=(12, 8.0), dpi=220)
ax = fig.add_axes([0, 0.155, 1, 0.845])
ax.imshow(img)
ax.set_xlim(0, W)
ax.set_ylim(H, 0)
ax.axis("off")
for n, ((cx, cy), (tx, ty)) in callouts.items():
    ax.plot([cx, tx], [cy, ty], color="white", lw=3.2, zorder=3, solid_capstyle="round")
    ax.plot([cx, tx], [cy, ty], color=NAVY, lw=1.4, zorder=4)
    ax.add_patch(Circle((tx, ty), 9, fc=NAVY, ec="white", lw=1.2, zorder=5))
    ax.add_patch(Circle((cx, cy), 44, fc="white", ec=NAVY, lw=2.2, zorder=6))
    ax.text(cx, cy + 2, str(n), ha="center", va="center", fontsize=13, weight="bold", family="Carlito", color=NAVY, zorder=7)

# legend panel
lg = fig.add_axes([0.0, 0.0, 1, 0.15])
lg.axis("off")
lg.set_xlim(0, 3)
lg.set_ylim(0, 3.2)
for i, (n, t) in enumerate(legend):
    col, row = i // 3, i % 3
    x, y = 0.06 + col * 1.5, 2.6 - row * 0.95
    lg.plot([x], [y], marker="o", ms=15, mfc="white", mec=NAVY, mew=1.6)
    lg.text(x, y, n, ha="center", va="center", fontsize=8.5, weight="bold", color=NAVY)
    lg.text(x + 0.11, y, t, va="center", fontsize=9.5, family="Carlito")
fig.text(0.995, 0.005, "Cable routes: orange = PV array cables; blue = wind collector cables; green = BESS feeder. Conceptual, not to scale.",
         ha="right", va="bottom", fontsize=7.2, color="#595959", style="italic")
fig.savefig("fig_illustration.png", dpi=220, facecolor="white")
print("saved")
