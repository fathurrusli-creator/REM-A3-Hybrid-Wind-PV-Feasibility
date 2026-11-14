"""Figure: CAPEX breakdown and year-1 revenue breakdown (base case)."""
import json, matplotlib
matplotlib.use("Agg"); import matplotlib.pyplot as plt
J = json.load(open("results.json")); C = J["capex"]; y = J["y1"]
plt.rcParams.update({"font.family": "Carlito", "font.size": 10.5, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.titlecolor": "#007B7B", "axes.edgecolor": "#7F7F7F"})
fig, axs = plt.subplots(1, 2, figsize=(10, 3.3))
cap = [("Wind turbines + BoP", C["wind"]), ("Solar PV", C["pv"]), ("Grid-forming BESS", C["bess"]), ("Connection & substation", C["conn"]), ("Development & owner's", C["dev"])]
rev = [("PPA (70% of output)", y["PPA revenue"]), ("Merchant energy", y["Merchant energy revenue"]), ("BESS arbitrage uplift", y["BESS arbitrage uplift"]),
       ("FCAS & system services", y["FCAS & system services"]), ("Certificates", y["Certificates"])]
for ax, data, title, tot in ((axs[0], cap, f"CAPEX: A${C['total']/1e6:,.1f} M", C["total"]), (axs[1], rev, f"Revenue, year 1: A${y['Total revenue']/1e6:,.2f} M", y["Total revenue"])):
    labels = [d[0] for d in data][::-1]; vals = [d[1] / 1e6 for d in data][::-1]
    colors = ["#007B7B" if v == max(vals) else "#66B2B2" for v in vals]
    ax.barh(labels, vals, color=colors)
    for i, v in enumerate(vals):
        ax.text(v + max(vals) * 0.01, i, f"{v:,.1f} ({v * 1e6 / tot:.0%})", va="center", fontsize=9.5)
    ax.set_xlim(0, max(vals) * 1.32); ax.set_xlabel("A$ million"); ax.set_title(title, loc="left", fontsize=11)
    ax.grid(axis="x", color="#E5E5E5"); ax.set_axisbelow(True)
fig.tight_layout(); fig.savefig("fig_breakdown.png", dpi=200)
print("ok")
