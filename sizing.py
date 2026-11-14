"""Sizing study: how NPV, IRR, LCOE and break-even PPA change with wind / PV / BESS size.
Technology-specific PPA prices are calibrated so the base design (100 MW wind + 50 MWac PV) blends to the
A$95/MWh used in the main model, leaving all base results unchanged."""
import importlib.util, json
import numpy as np, pandas as pd
spec = importlib.util.spec_from_file_location("m", "rem_hybrid_model.py"); m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
B = m.BASE
T0 = m.technical(B)
w_share = T0["wind_e1"] / (T0["wind_e1"] + T0["pv_e1"])
x = B["p_ppa"] - (w_share * 105 + (1 - w_share) * 60)           # firming premium so base blend = A$95
P_W, P_S = 105 + x, 60 + x

def config(wind, pv, bess_mwh=60, bess_mw=30):
    p = dict(B)
    wind, pv = max(wind, 1e-6), max(pv, 1e-6)
    p.update(wind_mw=wind, pv_mwac=pv, pv_mwdc=pv * 1.25, bess_mwh=bess_mwh, bess_mw=bess_mw)
    tot = wind + pv
    p["c_conn"] = B["c_conn"] * tot / 150                         # connection scaled to nameplate
    p["inv_cost"] = B["inv_cost"] * pv / 50
    p["decom"] = B["decom"] * tot / 150
    T = m.technical(p)
    ew, es = T["wind_e1"], T["pv_e1"]
    p["p_ppa"] = (ew * P_W + es * P_S) / (ew + es) if ew + es > 0 else P_S
    return p

def metrics(p, **lev):
    R = m.run(p, **lev)
    return dict(npv=R["npv"], irr=R["irr"], lcoe=R["lcoe"], be=R["be_ppa"], ppa=p["p_ppa"], capex=R["capex"],
                gen=R["gen1"], co2=R["co2_life"], pb=R["pb"])

out = {"P_W": P_W, "P_S": P_S}
# check base unchanged
b = metrics(config(100, 50)); out["base_check"] = b
# 1) wind sweep, PV 50
cases = {"Base prices": {}, "Price +20%": dict(price_f=1.2), "Price +20% & CAPEX -15%": dict(price_f=1.2, capex_f=0.85)}
ws = list(range(0, 201, 10))
out["wind_sweep"] = {c: [metrics(config(w, 50), **lev) | {"wind": w} for w in ws] for c, lev in cases.items()}
# 2) PV sweep, wind 100
pvs = list(range(0, 151, 10))
out["pv_sweep"] = {c: [metrics(config(100, s), **lev) | {"pv": s} for s in pvs] for c, lev in cases.items()}
# 3) configuration comparison
cfgs = [("Base design: 100 W + 50 PV + BESS", 100, 50, 60, 30), ("Same, without BESS", 100, 50, 0, 0),
        ("Wind-only 150 MW + BESS", 150, 0, 60, 30), ("Solar-only 150 MW + BESS", 0, 150, 60, 30),
        ("Wind-led 125 W + 25 PV + BESS", 125, 25, 60, 30), ("Balanced 75 W + 75 PV + BESS", 75, 75, 60, 30),
        ("Solar-led 50 W + 100 PV + BESS", 50, 100, 60, 30), ("Smaller 50 W + 50 PV + BESS", 50, 50, 60, 30)]
rows = []
for name, w, s, e, pw in cfgs:
    p = config(w, s, e, pw)
    r0, r1, r2 = metrics(p), metrics(p, price_f=1.2), metrics(p, price_f=1.2, capex_f=0.85)
    rows.append(dict(config=name, wind=w, pv=s, bess=f"{pw}/{e}", capex=r0["capex"], gen=r0["gen"], ppa=r0["ppa"], npv=r0["npv"], irr=r0["irr"],
                     lcoe=r0["lcoe"], be=r0["be"], co2=r0["co2"], npv_p20=r1["npv"], npv_comb=r2["npv"]))
out["configs"] = rows
# 4) marginal NPV per MW
out["marg_wind"] = (metrics(config(110, 50))["npv"] - metrics(config(100, 50))["npv"]) / 10
out["marg_pv"] = (metrics(config(100, 60))["npv"] - metrics(config(100, 50))["npv"]) / 10
# wind-only / pv-only standalone break-even (no BESS)
out["be_wind_only"] = metrics(config(150, 0, 0, 0))["be"]; out["be_pv_only"] = metrics(config(0, 150, 0, 0))["be"]
json.dump(out, open("sizing.json", "w"), indent=1, default=float)
print("P_W %.2f P_S %.2f | base NPV %.3f M (main model -97.515)" % (P_W, P_S, b["npv"] / 1e6))
print("marginal NPV per MW: wind %.2f M, PV %.2f M" % (out["marg_wind"] / 1e6, out["marg_pv"] / 1e6))
print("standalone BE (blended contract price): wind-only %.1f, PV-only %.1f" % (out["be_wind_only"], out["be_pv_only"]))
df = pd.DataFrame(rows); pd.set_option("display.width", 250)
print(df.assign(capex=df.capex / 1e6, npv=df.npv / 1e6, npv_p20=df.npv_p20 / 1e6, npv_comb=df.npv_comb / 1e6, co2=df.co2 / 1e6).round(2).to_string())
for c in cases:
    print(c, [(r["wind"], round(r["npv"] / 1e6, 1)) for r in out["wind_sweep"][c]][::4])

# ---------------- figure ----------------
import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
plt.rcParams.update({"font.family": "calibri", "font.size": 10.5, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.grid": True, "grid.color": "#E5E5E5", "axes.titlecolor": "#007B7B", "axes.edgecolor": "#7F7F7F"})
cols = {"Base prices": "#007B7B", "Price +20%": "#E87722", "Price +20% & CAPEX -15%": "#7F7F7F"}
fig, axs = plt.subplots(1, 2, figsize=(10, 3.9), sharey=True)
for ax, key, xk, xl, base_x in ((axs[0], "wind_sweep", "wind", "Wind capacity (MW), PV fixed at 50 MWac", 100),
                                (axs[1], "pv_sweep", "pv", "PV capacity (MWac), wind fixed at 100 MW", 50)):
    for c, v in out[key].items():
        ax.plot([r[xk] for r in v], [r["npv"] / 1e6 for r in v], lw=2.2, color=cols[c], label=c)
    ax.axhline(0, color="#404040", lw=0.9)
    ax.axvline(base_x, color="#404040", lw=0.8, ls="--")
    ax.text(base_x + 2, ax.get_ylim()[1] * 0.92 if ax.get_ylim()[1] > 0 else 5, "base design", fontsize=9, color="#404040")
    ax.set_xlabel(xl)
axs[0].set_ylabel("NPV at 8% (A$ million)")
axs[0].set_title("Varying the wind farm size", loc="left", fontsize=11)
axs[1].set_title("Varying the solar farm size", loc="left", fontsize=11)
axs[0].legend(frameon=False, loc="upper left", fontsize=9)
fig.tight_layout(); fig.savefig("fig_sizing.png", dpi=200); plt.close(fig)

# ---------------- Excel sheet ----------------
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
wb = load_workbook(m.OUT_XLSX)
if "Sizing_Study" in wb.sheetnames: del wb["Sizing_Study"]
ws = wb.create_sheet("Sizing_Study")
A = lambda **k: Font(name="Arial", size=10, **k)
ws["A1"] = "Sizing study (values computed by sizing.py using the same model as the CF sheets)"; ws["A1"].font = Font(name="Arial", size=13, bold=True)
ws["A2"] = (f"Technology-specific PPA prices: wind A${P_W:.2f}/MWh, solar A${P_S:.2f}/MWh (calibrated so the base 100/50 MW design blends to A$95/MWh). "
            "Connection and decommissioning costs scale with total nameplate; PV inverter replacement scales with PV size.")
ws["A2"].font = A(italic=True)
hdr = ["Configuration", "Wind (MW)", "PV (MWac)", "BESS (MW/MWh)", "CAPEX (A$ M)", "Net export yr 1 (MWh)", "Contract price (A$/MWh)",
       "NPV base (A$ M)", "IRR base", "LCOE (A$/MWh)", "Break-even PPA (A$/MWh)", "Gap to break-even", "NPV price +20% (A$ M)",
       "NPV price +20% & CAPEX -15% (A$ M)", "Lifetime CO2 avoided (Mt)"]
for j, h in enumerate(hdr, 1):
    c = ws.cell(row=4, column=j, value=h); c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="007B7B"); c.alignment = Alignment(wrap_text=True, horizontal="center")
for i, r in enumerate(rows, 5):
    vals = [r["config"], r["wind"], r["pv"], r["bess"], r["capex"] / 1e6, r["gen"], r["ppa"], r["npv"] / 1e6, r["irr"], r["lcoe"], r["be"],
            r["be"] / r["ppa"] - 1, r["npv_p20"] / 1e6, r["npv_comb"] / 1e6, r["co2"] / 1e6]
    fmts = [None, "0", "0", None, "#,##0.0", "#,##0", "0.00", "#,##0.0;(#,##0.0)", "0.00%", "0.0", "0.0", "0.0%", "#,##0.0;(#,##0.0)", "#,##0.0;(#,##0.0)", "0.00"]
    for j, (v, f) in enumerate(zip(vals, fmts), 1):
        c = ws.cell(row=i, column=j, value=v); c.font = A()
        if f: c.number_format = f
r0 = 5 + len(rows) + 2
ws.cell(row=r0, column=1, value="Wind sweep (PV 50 MWac, BESS 30/60): NPV (A$ M)").font = A(bold=True)
for j, h in enumerate(["Wind (MW)"] + list(out["wind_sweep"].keys()), 1):
    ws.cell(row=r0 + 1, column=j, value=h).font = A(bold=True)
for k, wv in enumerate(ws_ := list(range(0, 201, 10))):
    ws.cell(row=r0 + 2 + k, column=1, value=wv).font = A()
    for j, c in enumerate(out["wind_sweep"].keys(), 2):
        cell = ws.cell(row=r0 + 2 + k, column=j, value=out["wind_sweep"][c][k]["npv"] / 1e6); cell.number_format = "#,##0.0;(#,##0.0)"; cell.font = A()
ws.column_dimensions["A"].width = 36
for col in "BCDEFGHIJKLMNO": ws.column_dimensions[col].width = 14
ws.row_dimensions[4].height = 42
wb.save(m.OUT_XLSX)
print("figure + Excel sheet saved")
