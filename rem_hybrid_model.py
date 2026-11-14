"""
Renewable Energy Management - Class Assignment #3
Hybrid Solar PV-Wind + Grid-Forming BESS Feasibility Model
Central-West Orana REZ, NSW, Australia

What this script does
  1. Computes annual generation (Weibull/power-curve wind, PR-based PV), BESS energy
     shifting, 25-year nominal pre-tax unlevered cash flows and NPV, IRR, LCOE,
     simple and discounted payback, and CO2 avoided for the Base Case and 3 scenarios.
  2. Runs a one-at-a-time sensitivity (tornado) and a Monte Carlo uncertainty analysis.
  3. Writes a formula-driven Excel workbook (inputs in blue, formulas in black) so that
     every result can be recalculated and audited in Excel, and saves report figures.

Run:   python rem_hybrid_model.py
Needs: numpy, pandas, openpyxl, matplotlib (no other dependencies)
"""
import math
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

OUT_XLSX = "REM_A3_Hybrid_PV_Wind_Financial_Model.xlsx"

# ----------------------------------------------------------------------------------
# 1. ASSUMPTIONS  (key, value, unit, description, source)
# ----------------------------------------------------------------------------------
A = [
    # --- project / general
    ("life", 25, "years", "Operating lifetime", "Assignment brief; typical NEM asset life"),
    ("r", 0.08, "-", "Discount rate (nominal, pre-tax WACC)", "Assumed; within NEM project-finance range"),
    ("esc", 0.025, "-", "Price and cost escalation (CPI)", "RBA 2-3% inflation target band (midpoint)"),
    # --- wind
    ("wind_mw", 100, "MW", "Wind capacity (20 x 5.0 MW WTG)", "Project definition"),
    ("wtg_mw", 5.0, "MW", "Turbine rated power", "Generic 5 MW class, ~160 m rotor"),
    ("v_mean", 7.5, "m/s", "Mean wind speed at 150 m hub height", "Assumed CWO ridge-line value; verify with Global Wind Atlas + met mast (BoM surface data for validation)"),
    ("k", 2.1, "-", "Weibull shape factor", "Typical inland NSW value"),
    ("v_in", 3.0, "m/s", "Cut-in speed", "Generic power curve"),
    ("v_rated", 10.5, "m/s", "Rated speed", "Generic low-specific-power curve (~250 W/m2)"),
    ("v_out", 25.0, "m/s", "Cut-out speed", "Generic power curve"),
    ("wl_avail", 0.97, "-", "Wind availability factor", "Industry benchmark"),
    ("wl_wake", 0.93, "-", "Wake efficiency", "Industry benchmark (7% wake loss)"),
    ("wl_elec", 0.98, "-", "Electrical efficiency (collection + transformer)", "Industry benchmark"),
    ("wl_perf", 0.98, "-", "Turbine performance / environmental", "Industry benchmark"),
    ("wl_curt", 0.97, "-", "Network + economic curtailment factor", "Assumed for CWO REZ (3%)"),
    ("wind_deg", 0.003, "/yr", "Wind performance degradation", "Staffell & Green (2014) order of magnitude"),
    # --- PV
    ("pv_mwac", 50, "MWac", "PV inverter (AC) capacity", "Project definition"),
    ("pv_mwdc", 62.5, "MWp", "PV DC capacity (DC/AC ratio 1.25)", "Project definition"),
    ("ghi", 2000, "kWh/m2/yr", "Global horizontal irradiation (Dubbo/Orana region)", "BoM gridded solar exposure, ~5.5 kWh/m2/day"),
    ("trk_gain", 1.20, "-", "Single-axis tracker gain (POA/GHI)", "Typical for latitude ~32 S"),
    ("pr_temp", 0.93, "-", "Temperature loss factor", "Assumed for hot inland summer"),
    ("pr_soil", 0.98, "-", "Soiling factor", "Assumed"),
    ("pr_mism", 0.97, "-", "Mismatch + DC wiring factor", "Assumed"),
    ("pr_inv", 0.98, "-", "Inverter efficiency", "Manufacturer CEC-weighted"),
    ("pr_clip", 0.99, "-", "Inverter clipping factor", "DC/AC 1.25"),
    ("pr_ac", 0.985, "-", "AC / transformer / auxiliaries", "Assumed"),
    ("pr_avail", 0.99, "-", "PV availability", "Industry benchmark"),
    ("pr_curt", 0.96, "-", "Economic curtailment (negative prices)", "Assumed for NSW midday (4%)"),
    ("pv_deg", 0.005, "/yr", "PV module degradation", "Manufacturer warranty / IEA-PVPS"),
    # --- BESS
    ("bess_mw", 30, "MW", "BESS power (grid-forming)", "Project definition"),
    ("bess_mwh", 60, "MWh", "BESS energy (2 h)", "Project definition"),
    ("dod", 0.90, "-", "Usable depth of discharge", "Assumed"),
    ("cycles", 300, "cycles/yr", "Equivalent full cycles", "Assumed (1/day, ~300 days)"),
    ("rte", 0.85, "-", "Round-trip efficiency (AC-AC)", "GenCost / manufacturer data"),
    ("bess_fade", 0.02, "/yr", "BESS capacity fade", "Assumed; restored by augmentation in year 12"),
    ("aug_year", 12, "year", "BESS augmentation / PV inverter replacement year", "Assumed"),
    # --- market / revenue
    ("mlf", 0.93, "-", "Marginal Loss Factor (CWO connection point)", "Assumed from AEMO MLF reports for Central West NSW (0.90-0.96)"),
    ("ppa_share", 0.70, "-", "Share of net export sold under corporate PPA", "Assumed contracting strategy"),
    ("p_ppa", 95, "A$/MWh", "PPA price, bundled energy + certificates (2026 A$, CPI-indexed)", "Assumed; consistent with recent NSW wind/hybrid PPAs"),
    ("p_wind", 85, "A$/MWh", "Merchant wind capture price (2026 A$)", "Assumed from AEMO NSW1 spot history"),
    ("p_solar", 45, "A$/MWh", "Merchant solar capture price (2026 A$)", "Assumed; reflects daytime price cannibalisation"),
    ("p_peak", 150, "A$/MWh", "Evening peak price for BESS discharge (2026 A$)", "Assumed from NSW1 evening peak"),
    ("p_cert", 5, "A$/MWh", "Certificate value on merchant volume (LGC/REGO), flat", "Assumed; LGC prices falling to 2030"),
    ("fcas", 50000, "A$/MW-yr", "FCAS + system-strength/inertia services (2026 A$)", "Assumed; Transgrid system-strength contracts"),
    # --- CAPEX
    ("c_wind", 3248, "A$/kW", "Onshore wind capital cost", "CSIRO GenCost 2025-26 (draft)"),
    ("c_pv", 1621, "A$/kW", "Large-scale solar PV capital cost (per kWac)", "CSIRO GenCost 2025-26 (draft)"),
    ("c_bess", 525, "A$/kWh", "2-h battery capital cost", "CSIRO GenCost 2025-26 (draft)"),
    ("gfm_prem", 0.10, "-", "Grid-forming inverter premium on BESS", "Assumed"),
    ("c_conn", 25e6, "A$", "Grid connection, substation, system-strength remediation", "Assumed"),
    ("c_dev", 0.05, "-", "Development & owner's costs (% of above)", "Assumed"),
    # --- OPEX (2026 A$, escalated)
    ("o_wind", 45, "A$/kW-yr", "Wind fixed O&M incl. land lease and insurance", "GenCost FOM uplifted for leases/insurance"),
    ("o_pv", 20, "A$/kW-yr", "PV fixed O&M", "GenCost-based"),
    ("o_bess", 8, "A$/kWh-yr", "BESS fixed O&M incl. warranty", "Assumed"),
    ("o_other", 1.5e6, "A$/yr", "REZ access fee, network charges, asset management", "Assumed"),
    ("aug_cost", 0.35, "-", "BESS augmentation cost (% of BESS CAPEX)", "Assumed"),
    ("inv_cost", 3e6, "A$", "PV inverter replacement", "Assumed (~A$60/kWac)"),
    ("decom", 5e6, "A$", "Net decommissioning cost in final year (after salvage)", "Assumed"),
    # --- environment
    ("ef", 0.64, "tCO2-e/MWh", "NSW grid emission factor (scope 2, location-based)", "DCCEEW NGA Factors 2025"),
    ("ef_decl", 0.05, "/yr", "Grid decarbonisation rate (sensitivity only)", "Assumed, consistent with AEMO ISP Step Change direction"),
]
BASE = {k: v for k, v, *_ in A}

SCENARIOS = [
    ("Base Case", dict(price_f=1.00, capex_f=1.00, dr_add=0.00)),
    ("S1 Price +20%", dict(price_f=1.20, capex_f=1.00, dr_add=0.00)),
    ("S2 CAPEX -15%", dict(price_f=1.00, capex_f=0.85, dr_add=0.00)),
    ("S3 Rate +2 pp", dict(price_f=1.00, capex_f=1.00, dr_add=0.02)),
]


# ----------------------------------------------------------------------------------
# 2. TECHNICAL MODEL
# ----------------------------------------------------------------------------------
def weibull_cdf(v, k, c):
    return 1.0 - np.exp(-(np.asarray(v) / c) ** k)


def power_curve(v, p):
    """Generic WTG power curve (MW): cubic between cut-in and rated."""
    v = np.asarray(v, dtype=float)
    out = np.where((v >= p["v_in"]) & (v < p["v_rated"]),
                   p["wtg_mw"] * (v ** 3 - p["v_in"] ** 3) / (p["v_rated"] ** 3 - p["v_in"] ** 3), 0.0)
    out = np.where((v >= p["v_rated"]) & (v < p["v_out"]), p["wtg_mw"], out)
    return out


def wind_bins(p):
    lo = np.arange(0, 30, 0.5)
    hi = lo + 0.5
    mid = lo + 0.25
    c = p["v_mean"] / math.gamma(1 + 1 / p["k"])
    prob = weibull_cdf(hi, p["k"], c) - weibull_cdf(lo, p["k"], c)
    return lo, hi, mid, prob, power_curve(mid, p), c


def technical(p):
    lo, hi, mid, prob, pw, c = wind_bins(p)
    n_wtg = p["wind_mw"] / p["wtg_mw"]
    gross_wind = 8760 * float(np.sum(prob * pw)) * n_wtg               # MWh/yr
    wind_loss = p["wl_avail"] * p["wl_wake"] * p["wl_elec"] * p["wl_perf"] * p["wl_curt"]
    wind_e1 = gross_wind * wind_loss
    pr = (p["pr_temp"] * p["pr_soil"] * p["pr_mism"] * p["pr_inv"] * p["pr_clip"]
          * p["pr_ac"] * p["pr_avail"] * p["pr_curt"])
    poa = p["ghi"] * p["trk_gain"]
    pv_e1 = p["pv_mwdc"] * poa * pr                                     # MWp * kWh/m2 * PR = MWh
    bess_d1 = p["bess_mwh"] * p["dod"] * p["cycles"]                    # MWh discharged, year 1
    return dict(c=c, gross_wind=gross_wind, wind_loss=wind_loss, wind_e1=wind_e1,
                cf_wind_gross=gross_wind / (p["wind_mw"] * 8760), cf_wind=wind_e1 / (p["wind_mw"] * 8760),
                pr=pr, poa=poa, pv_e1=pv_e1, cf_pv=pv_e1 / (p["pv_mwac"] * 8760), bess_d1=bess_d1)


# ----------------------------------------------------------------------------------
# 3. FINANCIAL MODEL
# ----------------------------------------------------------------------------------
def capex_items(p):
    wind = p["wind_mw"] * 1000 * p["c_wind"]
    pv = p["pv_mwac"] * 1000 * p["c_pv"]
    bess = p["bess_mwh"] * 1000 * p["c_bess"] * (1 + p["gfm_prem"])
    conn = p["c_conn"]
    dev = p["c_dev"] * (wind + pv + bess + conn)
    return dict(wind=wind, pv=pv, bess=bess, conn=conn, dev=dev, total=wind + pv + bess + conn + dev)


def irr(cf, lo=-0.99, hi=1.0):
    f = lambda r: np.sum(cf / (1 + r) ** np.arange(len(cf)))
    if f(lo) * f(hi) > 0:
        return float("nan")
    for _ in range(200):
        mid = (lo + hi) / 2
        if f(lo) * f(mid) <= 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


def payback(cum, cf):
    for t in range(1, len(cum)):
        if cum[t - 1] < 0 <= cum[t]:
            return (t - 1) + (-cum[t - 1]) / cf[t]
    return float("nan")


def run(p, price_f=1.0, capex_f=1.0, dr_add=0.0):
    T = technical(p)
    L = int(p["life"])
    t = np.arange(1, L + 1)
    r = p["r"] + dr_add
    wind = T["wind_e1"] * (1 - p["wind_deg"]) ** (t - 1)
    pv = T["pv_e1"] * (1 - p["pv_deg"]) ** (t - 1)
    age = np.where(t <= p["aug_year"], t - 1, t - p["aug_year"] - 1)
    disch = T["bess_d1"] * (1 - p["bess_fade"]) ** age
    loss = disch / p["rte"] - disch
    net = wind + pv - loss
    g = (1 + p["esc"]) ** (t - 1)
    s = p["ppa_share"]
    rev_ppa = s * (wind + pv) * p["p_ppa"] * p["mlf"] * g * price_f
    rev_mer = ((1 - s) * wind * p["p_wind"] + ((1 - s) * pv - loss) * p["p_solar"]) * p["mlf"] * g * price_f
    rev_arb = disch * (p["p_peak"] - p["p_solar"]) * p["mlf"] * g * price_f
    rev_fcas = p["fcas"] * p["bess_mw"] * g
    rev_cert = ((1 - s) * (wind + pv) - loss) * p["p_cert"] * p["mlf"]
    rev = rev_ppa + rev_mer + rev_arb + rev_fcas + rev_cert
    cx = capex_items(p)
    capex = cx["total"] * capex_f
    o_fix = (p["wind_mw"] * 1000 * p["o_wind"] + p["pv_mwac"] * 1000 * p["o_pv"]
             + p["bess_mwh"] * 1000 * p["o_bess"] + p["o_other"]) * g
    major = np.where(t == p["aug_year"], (p["aug_cost"] * cx["bess"] * capex_f + p["inv_cost"]) * g, 0.0)
    major = major + np.where(t == L, p["decom"] * g, 0.0)
    opex = o_fix + major
    cf = np.concatenate([[-capex], rev - opex])
    yrs = np.arange(0, L + 1)
    df = 1 / (1 + r) ** yrs
    dcf = cf * df
    cum, dcum = np.cumsum(cf), np.cumsum(dcf)
    npv = float(dcf.sum())
    lcoe = (capex + float(np.sum(opex * df[1:]))) / float(np.sum(net * df[1:]))
    p_eff = p["p_ppa"] * price_f                      # effective PPA price in this scenario
    dnpv_dppa = float(np.sum(rev_ppa / p_eff * df[1:]))
    be_ppa = p_eff - npv / dnpv_dppa                  # NPV is linear in the PPA price
    co2 = net * p["ef"]
    co2_decl = net * p["ef"] * (1 - p["ef_decl"]) ** (t - 1)
    table = pd.DataFrame({
        "Year": yrs, "Wind MWh": np.r_[0, wind], "PV MWh": np.r_[0, pv], "BESS discharge MWh": np.r_[0, disch],
        "BESS losses MWh": np.r_[0, loss], "Net export MWh": np.r_[0, net],
        "PPA revenue": np.r_[0, rev_ppa], "Merchant energy revenue": np.r_[0, rev_mer],
        "BESS arbitrage uplift": np.r_[0, rev_arb], "FCAS & system services": np.r_[0, rev_fcas],
        "Certificates": np.r_[0, rev_cert], "Total revenue": np.r_[0, rev],
        "Fixed O&M": np.r_[0, o_fix], "Major maint. & decom.": np.r_[0, major], "Total OPEX": np.r_[0, opex],
        "CAPEX": np.r_[capex, np.zeros(L)], "Net cash flow": cf, "Discount factor": df, "Discounted CF": dcf,
        "Cumulative CF": cum, "Cumulative DCF": dcum, "CO2 avoided t": np.r_[0, co2]})
    return dict(table=table, T=T, capex=capex, npv=npv, irr=irr(cf), lcoe=lcoe,
                pb=payback(cum, cf), dpb=payback(dcum, dcf), r=r, be_ppa=be_ppa,
                gen1=float(net[0]), gen_avg=float(net.mean()), rev1=float(rev[0]), rev_avg=float(rev.mean()),
                opex1=float(opex[0]), co2_1=float(co2[0]), co2_life=float(co2.sum()),
                co2_life_decl=float(co2_decl.sum()), gen_life=float(net.sum()))


# ----------------------------------------------------------------------------------
# 4. SENSITIVITY + MONTE CARLO
# ----------------------------------------------------------------------------------
def tornado(p):
    base = run(p)["npv"]
    rows = []
    tests = [("Electricity prices (all energy prices)", "price", 0.8, 1.2),
             ("CAPEX", "capex", 1.15, 0.85),
             ("Discount rate (+/-2 pp)", "dr", 0.02, -0.02),
             ("Mean wind speed (+/-7%)", "v_mean", 0.93, 1.07),
             ("Marginal Loss Factor (0.88 / 0.96)", "mlf_abs", 0.88, 0.96),
             ("PPA price (+/-20%)", "p_ppa", 0.8, 1.2),
             ("Fixed OPEX (+/-20%)", "opex", 1.2, 0.8),
             ("Solar irradiation (+/-5%)", "ghi", 0.95, 1.05)]
    for name, key, lo, hi in tests:
        vals = []
        for f in (lo, hi):
            q = dict(p)
            kw = {}
            if key == "price":
                kw["price_f"] = f
            elif key == "capex":
                kw["capex_f"] = f
            elif key == "dr":
                kw["dr_add"] = f
            elif key == "mlf_abs":
                q["mlf"] = f
            elif key == "opex":
                for o in ("o_wind", "o_pv", "o_bess", "o_other"):
                    q[o] = p[o] * f
            else:
                q[key] = p[key] * f
            vals.append(run(q, **kw)["npv"])
        rows.append((name, vals[0], vals[1], max(vals) - min(vals)))
    df = pd.DataFrame(rows, columns=["Variable", "NPV adverse (A$)", "NPV favourable (A$)", "Swing (A$)"])
    return base, df.sort_values("Swing (A$)", ascending=False).reset_index(drop=True)


def monte_carlo(p, n=5000, seed=42):
    rng = np.random.default_rng(seed)
    out = []
    for _ in range(n):
        q = dict(p)
        q["v_mean"] = p["v_mean"] * rng.normal(1, 0.05)
        q["ghi"] = p["ghi"] * rng.normal(1, 0.03)
        q["mlf"] = rng.triangular(0.86, p["mlf"], 0.97)
        pf = rng.triangular(0.75, 1.0, 1.20)
        cf_ = rng.triangular(0.90, 1.0, 1.20)
        dr = rng.normal(0, 0.01)
        for o in ("o_wind", "o_pv", "o_bess", "o_other"):
            q[o] = p[o] * rng.triangular(0.9, 1.0, 1.25)
        R = run(q, price_f=pf, capex_f=cf_, dr_add=dr)
        out.append((R["npv"], R["irr"], R["lcoe"]))
    return pd.DataFrame(out, columns=["NPV", "IRR", "LCOE"])


# ----------------------------------------------------------------------------------
# 5. EXCEL WORKBOOK (formula-driven)
# ----------------------------------------------------------------------------------
BLUE = Font(name="Arial", size=10, color="0000FF")
BLACK = Font(name="Arial", size=10, color="000000")
GREEN = Font(name="Arial", size=10, color="008000")
BOLD = Font(name="Arial", size=10, bold=True)
HFONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
TITLE = Font(name="Arial", size=13, bold=True)
HFILL = PatternFill("solid", fgColor="1F4E78")
YFILL = PatternFill("solid", fgColor="FFFF00")
SFILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Border(bottom=Side(style="thin", color="BFBFBF"))
MONEY = '#,##0;(#,##0);"-"'
MWH = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'


def hdr(ws, row, cols):
    for j, h in enumerate(cols, 1):
        c = ws.cell(row=row, column=j, value=h)
        c.font, c.fill = HFONT, HFILL
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")


def build_workbook(base_results, scen_results, tor_df, mc_df):
    wb = Workbook()
    # ---------------- Assumptions
    wa = wb.active
    wa.title = "Assumptions"
    wa["A1"] = "Inputs and assumptions (blue = hard-coded input; edit here and every sheet recalculates)"
    wa["A1"].font = TITLE
    hdr(wa, 3, ["Key", "Value", "Unit", "Description", "Source / basis"])
    ref = {}
    for i, (k, v, u, d, s) in enumerate(A, start=4):
        wa.cell(row=i, column=1, value=k).font = BLACK
        c = wa.cell(row=i, column=2, value=v)
        c.font = BLUE
        if k in ("r", "esc", "ppa_share", "mlf", "v_mean", "p_ppa", "c_wind", "ef"):
            c.fill = YFILL
        if isinstance(v, float) and v < 1 and u in ("-", "/yr"):
            c.number_format = "0.0%" if k not in ("mlf", "k", "trk_gain", "rte", "dod") and not k.startswith(("wl_", "pr_")) else "0.000"
        elif isinstance(v, (int, float)) and abs(v) >= 1000:
            c.number_format = "#,##0"
        wa.cell(row=i, column=3, value=u).font = BLACK
        wa.cell(row=i, column=4, value=d).font = BLACK
        wa.cell(row=i, column=5, value=s).font = BLACK
        ref[k] = f"Assumptions!$B${i}"
    for col, w in zip("ABCDE", (12, 14, 12, 58, 70)):
        wa.column_dimensions[col].width = w
    last = 4 + len(A)
    wa.cell(row=last + 1, column=1, value="Yellow fill = key assumptions that drive the investment decision.").font = BOLD

    # ---------------- Technical
    wt = wb.create_sheet("Technical")
    wt["A1"] = "Technical calculations: wind (Weibull x power curve) and PV (performance ratio)"
    wt["A1"].font = TITLE
    rows = [
        ("Weibull scale factor c (m/s)", f"={ref['v_mean']}/EXP(GAMMALN(1+1/{ref['k']}))", "c = v_mean / Gamma(1+1/k)"),
        ("Number of turbines", f"={ref['wind_mw']}/{ref['wtg_mw']}", "N = P_wind / P_WTG"),
        ("Gross energy per WTG (MWh/yr)", "=8760*SUMPRODUCT(Technical!$E$32:$E$91,Technical!$F$32:$F$91)", "E = 8760 x sum(p_i x P(v_i))"),
        ("Gross wind energy (MWh/yr)", "=B5*B6", "E_gross = N x E_WTG"),
        ("Gross wind capacity factor", f"=B7/({ref['wind_mw']}*8760)", "CF = E / (P x 8760)"),
        ("Wind loss factor (product)", f"={ref['wl_avail']}*{ref['wl_wake']}*{ref['wl_elec']}*{ref['wl_perf']}*{ref['wl_curt']}", "eta_w = product of loss factors"),
        ("Net wind energy, year 1 (MWh)", "=B7*B9", "E_wind = E_gross x eta_w"),
        ("Net wind capacity factor", f"=B10/({ref['wind_mw']}*8760)", ""),
        ("Plane-of-array irradiation (kWh/m2/yr)", f"={ref['ghi']}*{ref['trk_gain']}", "H_POA = GHI x tracker gain"),
        ("PV performance ratio (product)", f"={ref['pr_temp']}*{ref['pr_soil']}*{ref['pr_mism']}*{ref['pr_inv']}*{ref['pr_clip']}*{ref['pr_ac']}*{ref['pr_avail']}*{ref['pr_curt']}", "PR = product of loss factors"),
        ("Net PV energy, year 1 (MWh)", f"={ref['pv_mwdc']}*B12*B13", "E_PV = P_DC x H_POA x PR"),
        ("PV capacity factor (AC basis)", f"=B14/({ref['pv_mwac']}*8760)", ""),
        ("Specific yield (kWh/kWp)", f"=B14/{ref['pv_mwdc']}", ""),
        ("BESS discharge, year 1 (MWh)", f"={ref['bess_mwh']}*{ref['dod']}*{ref['cycles']}", "E_dis = E_BESS x DoD x cycles"),
        ("BESS losses, year 1 (MWh)", f"=B17/{ref['rte']}-B17", "E_loss = E_dis/RTE - E_dis"),
        ("Hybrid net export, year 1 (MWh)", "=B10+B14-B18", "E_net = E_wind + E_PV - E_loss"),
        ("Hybrid capacity factor (on 150 MW)", f"=B19/(({ref['wind_mw']}+{ref['pv_mwac']})*8760)", ""),
    ]
    hdr(wt, 3, ["Quantity", "Value", "Equation"])
    for i, (lab, f, eq) in enumerate(rows, start=4):
        wt.cell(row=i, column=1, value=lab).font = BLACK
        c = wt.cell(row=i, column=2, value=f)
        c.font = BLACK
        c.number_format = "0.0%" if ("factor" in lab.lower() and "Weibull" not in lab) or "ratio" in lab.lower() else "#,##0.00" if i == 4 else "#,##0"
        wt.cell(row=i, column=3, value=eq).font = BLACK
    wt["A30"] = "Wind-speed distribution and generic 5 MW power curve"
    wt["A30"].font = BOLD
    hdr(wt, 31, ["v low (m/s)", "v high (m/s)", "v mid (m/s)", "", "Probability", "WTG power (MW)"])
    for j in range(60):
        r_ = 32 + j
        wt.cell(row=r_, column=1, value=j * 0.5).font = BLUE
        wt.cell(row=r_, column=2, value=f"=A{r_}+0.5").font = BLACK
        wt.cell(row=r_, column=3, value=f"=A{r_}+0.25").font = BLACK
        wt.cell(row=r_, column=5, value=f"=(1-EXP(-((B{r_}/$B$4)^{ref['k']})))-(1-EXP(-((A{r_}/$B$4)^{ref['k']})))").number_format = "0.0000"
        wt.cell(row=r_, column=6, value=(
            f"=IF(AND(C{r_}>={ref['v_in']},C{r_}<{ref['v_rated']}),{ref['wtg_mw']}*(C{r_}^3-{ref['v_in']}^3)/({ref['v_rated']}^3-{ref['v_in']}^3),"
            f"IF(AND(C{r_}>={ref['v_rated']},C{r_}<{ref['v_out']}),{ref['wtg_mw']},0))")).number_format = "0.000"
    wt.column_dimensions["A"].width = 40
    wt.column_dimensions["B"].width = 16
    wt.column_dimensions["C"].width = 42
    wt.column_dimensions["E"].width = 13
    wt.column_dimensions["F"].width = 15

    # ---------------- CAPEX
    wc = wb.create_sheet("CAPEX")
    wc["A1"] = "Capital cost build-up (2026 A$)"
    wc["A1"].font = TITLE
    hdr(wc, 3, ["Item", "A$", "A$/kW (on 150 MW)", "Basis"])
    cap_rows = [
        ("Wind turbines + BoP", f"={ref['wind_mw']}*1000*{ref['c_wind']}", "100 MW x GenCost A$/kW"),
        ("Solar PV (single-axis tracking)", f"={ref['pv_mwac']}*1000*{ref['c_pv']}", "50 MWac x GenCost A$/kW"),
        ("Grid-forming BESS", f"={ref['bess_mwh']}*1000*{ref['c_bess']}*(1+{ref['gfm_prem']})", "60 MWh x A$/kWh x (1+GFM premium)"),
        ("Connection, substation, system strength", f"={ref['c_conn']}", "Assumed lump sum"),
        ("Development & owner's costs", f"={ref['c_dev']}*SUM(B4:B7)", "% of items above"),
        ("TOTAL CAPEX", "=SUM(B4:B8)", ""),
    ]
    for i, (lab, f, b) in enumerate(cap_rows, start=4):
        wc.cell(row=i, column=1, value=lab).font = BOLD if "TOTAL" in lab else BLACK
        c = wc.cell(row=i, column=2, value=f)
        c.number_format = MONEY
        c.font = BOLD if "TOTAL" in lab else BLACK
        c2 = wc.cell(row=i, column=3, value=f"=B{i}/(({ref['wind_mw']}+{ref['pv_mwac']})*1000)")
        c2.number_format = "#,##0"
        wc.cell(row=i, column=4, value=b).font = BLACK
    wc.column_dimensions["A"].width = 40
    wc.column_dimensions["B"].width = 18
    wc.column_dimensions["C"].width = 20
    wc.column_dimensions["D"].width = 38

    # ---------------- Scenario cash-flow sheets
    cols = ["Year", "Wind MWh", "PV MWh", "BESS discharge MWh", "BESS losses MWh", "Net export MWh",
            "Price index", "PPA revenue", "Merchant energy revenue", "BESS arbitrage uplift",
            "FCAS & system services", "Certificates", "Total revenue", "Fixed O&M", "Major maint. & decom.",
            "Total OPEX", "CAPEX", "Net cash flow", "Discount factor", "Discounted CF", "Cumulative CF",
            "Cumulative DCF", "CO2 avoided (t)", "Payback helper", "Disc. payback helper"]
    metric_cells = {}
    sheet_names = ["CF_Base", "CF_S1_Price+20", "CF_S2_CAPEX-15", "CF_S3_Rate+2pp"]
    for (sname, lev), wsname in zip(SCENARIOS, sheet_names):
        ws = wb.create_sheet(wsname)
        ws["A1"] = f"25-year cash flow - {sname} (nominal A$, pre-tax, unlevered)"
        ws["A1"].font = TITLE
        ws["A2"], ws["A3"], ws["A4"] = "Price factor", "CAPEX factor", "Discount-rate adder"
        ws["B2"], ws["B3"], ws["B4"] = lev["price_f"], lev["capex_f"], lev["dr_add"]
        for c in ("B2", "B3", "B4"):
            ws[c].font, ws[c].fill = BLUE, YFILL
        ws["B4"].number_format = "0.0%"
        ws["D2"], ws["E2"] = "Discount rate used", f"={ref['r']}+B4"
        ws["E2"].number_format = "0.0%"
        ws["D3"], ws["E3"] = "CAPEX used (A$)", "=CAPEX!$B$9*B3"
        ws["E3"].number_format = MONEY
        H = 7
        hdr(ws, H, cols)
        r0 = H + 1
        rN = r0 + int(BASE["life"])
        for y in range(0, int(BASE["life"]) + 1):
            R = r0 + y
            ws.cell(row=R, column=1, value=y).font = BLACK
            if y == 0:
                vals = {17: "=$E$3", 18: f"=-Q{R}", 19: f"=1/(1+$E$2)^A{R}", 20: f"=R{R}*S{R}", 21: f"=R{R}", 22: f"=T{R}"}
                for col in range(2, 24):
                    ws.cell(row=R, column=col, value=vals.get(col, 0))
            else:
                age = f"IF(A{R}<={ref['aug_year']},A{R}-1,A{R}-{ref['aug_year']}-1)"
                f = {
                    2: f"=Technical!$B$10*(1-{ref['wind_deg']})^(A{R}-1)",
                    3: f"=Technical!$B$14*(1-{ref['pv_deg']})^(A{R}-1)",
                    4: f"=Technical!$B$17*(1-{ref['bess_fade']})^{age}",
                    5: f"=D{R}/{ref['rte']}-D{R}",
                    6: f"=B{R}+C{R}-E{R}",
                    7: f"=(1+{ref['esc']})^(A{R}-1)",
                    8: f"={ref['ppa_share']}*(B{R}+C{R})*{ref['p_ppa']}*{ref['mlf']}*G{R}*$B$2",
                    9: f"=((1-{ref['ppa_share']})*B{R}*{ref['p_wind']}+((1-{ref['ppa_share']})*C{R}-E{R})*{ref['p_solar']})*{ref['mlf']}*G{R}*$B$2",
                    10: f"=D{R}*({ref['p_peak']}-{ref['p_solar']})*{ref['mlf']}*G{R}*$B$2",
                    11: f"={ref['fcas']}*{ref['bess_mw']}*G{R}",
                    12: f"=((1-{ref['ppa_share']})*(B{R}+C{R})-E{R})*{ref['p_cert']}*{ref['mlf']}",
                    13: f"=SUM(H{R}:L{R})",
                    14: f"=({ref['wind_mw']}*1000*{ref['o_wind']}+{ref['pv_mwac']}*1000*{ref['o_pv']}+{ref['bess_mwh']}*1000*{ref['o_bess']}+{ref['o_other']})*G{R}",
                    15: f"=IF(A{R}={ref['aug_year']},({ref['aug_cost']}*CAPEX!$B$6*$B$3+{ref['inv_cost']})*G{R},0)+IF(A{R}={ref['life']},{ref['decom']}*G{R},0)",
                    16: f"=N{R}+O{R}",
                    17: 0,
                    18: f"=M{R}-P{R}-Q{R}",
                    19: f"=1/(1+$E$2)^A{R}",
                    20: f"=R{R}*S{R}",
                    21: f"=U{R-1}+R{R}",
                    22: f"=V{R-1}+T{R}",
                    23: f"=F{R}*{ref['ef']}",
                    24: f'=IF(AND(U{R-1}<0,U{R}>=0),A{R-1}+(-U{R-1})/R{R},"")',
                    25: f'=IF(AND(V{R-1}<0,V{R}>=0),A{R-1}+(-V{R-1})/T{R},"")',
                }
                for col, val in f.items():
                    ws.cell(row=R, column=col, value=val)
            for col in range(2, 26):
                c = ws.cell(row=R, column=col)
                c.font = BLACK
                c.number_format = "0.0000" if col in (7, 19) else "0.00" if col in (24, 25) else MONEY
        # totals & metrics
        T0 = rN + 2
        ws.cell(row=T0, column=1, value="Lifetime total").font = BOLD
        for col in (2, 3, 6, 13, 16, 18, 23):
            L_ = get_column_letter(col)
            c = ws.cell(row=T0, column=col, value=f"=SUM({L_}{r0}:{L_}{rN})")
            c.font, c.number_format = BOLD, MONEY
        mrow = T0 + 2
        metrics = [
            ("Annual net generation, year 1 (MWh)", f"=F{r0+1}", MWH),
            ("Average annual net generation (MWh)", f"=AVERAGE(F{r0+1}:F{rN})", MWH),
            ("Annual revenue, year 1 (A$)", f"=M{r0+1}", MONEY),
            ("Average annual revenue (A$)", f"=AVERAGE(M{r0+1}:M{rN})", MONEY),
            ("NPV (A$)", f"=SUM(T{r0}:T{rN})", MONEY),
            ("IRR", f"=IRR(R{r0}:R{rN})", "0.00%"),
            ("LCOE (A$/MWh)", f"=(Q{r0}+SUMPRODUCT(P{r0+1}:P{rN},S{r0+1}:S{rN}))/SUMPRODUCT(F{r0+1}:F{rN},S{r0+1}:S{rN})", "0.00"),
            ("Simple payback (years)", f'=IF(COUNT(X{r0+1}:X{rN})=0,"> 25",MIN(X{r0+1}:X{rN}))', "0.0"),
            ("Discounted payback (years)", f'=IF(COUNT(Y{r0+1}:Y{rN})=0,"> 25",MIN(Y{r0+1}:Y{rN}))', "0.0"),
            ("CO2 avoided, year 1 (t)", f"=W{r0+1}", MWH),
            ("Lifetime CO2 avoided (t)", f"=SUM(W{r0+1}:W{rN})", MWH),
            ("Break-even PPA price (A$/MWh, 2026 A$)", f"={ref['p_ppa']}*$B$2-B{mrow+4}/(SUMPRODUCT(H{r0+1}:H{rN},S{r0+1}:S{rN})/({ref['p_ppa']}*$B$2))", "0.00"),
        ]
        metric_cells[wsname] = {}
        for i, (lab, f, nf) in enumerate(metrics):
            ws.cell(row=mrow + i, column=1, value=lab).font = BOLD
            c = ws.cell(row=mrow + i, column=2, value=f)
            c.number_format, c.font, c.fill = nf, BOLD, SFILL
            metric_cells[wsname][lab] = f"'{wsname}'!$B${mrow + i}"
        ws.column_dimensions["A"].width = 36
        for col in range(2, 26):
            ws.column_dimensions[get_column_letter(col)].width = 15
        ws.freeze_panes = ws.cell(row=H + 1, column=2)

    # ---------------- Scenario summary
    wsum = wb.create_sheet("Scenario_Summary", 1)
    wsum["A1"] = "Scenario comparison (all values link to the cash-flow sheets)"
    wsum["A1"].font = TITLE
    hdr(wsum, 3, ["Metric"] + [s for s, _ in SCENARIOS])
    labels = list(metric_cells["CF_Base"].keys())
    for i, lab in enumerate(labels, start=4):
        wsum.cell(row=i, column=1, value=lab).font = BOLD
        for j, wsname in enumerate(sheet_names, start=2):
            c = wsum.cell(row=i, column=j, value=f"={metric_cells[wsname][lab]}")
            c.font = GREEN
            c.number_format = "0.00%" if lab == "IRR" else "0.00" if ("LCOE" in lab or "PPA" in lab) else "0.0" if "payback" in lab else MONEY
    wsum.column_dimensions["A"].width = 40
    for col in "BCDE":
        wsum.column_dimensions[col].width = 18
    n = 4 + len(labels) + 1
    wsum.cell(row=n, column=1, value="Change in NPV vs Base (A$)").font = BOLD
    for j in range(2, 6):
        L_ = get_column_letter(j)
        c = wsum.cell(row=n, column=j, value=f"={L_}8-$B$8")
        c.number_format = MONEY
    wsum.cell(row=n + 2, column=1, value="Scenario levers are the yellow cells at the top of each CF_ sheet.").font = BLACK

    # ---------------- Sensitivity + Monte Carlo (python-computed values)
    wsn = wb.create_sheet("Sensitivity_Tornado")
    wsn["A1"] = "One-at-a-time NPV sensitivity (values computed by rem_hybrid_model.py)"
    wsn["A1"].font = TITLE
    hdr(wsn, 3, list(tor_df.columns))
    for i, row in enumerate(tor_df.itertuples(index=False), start=4):
        for j, v in enumerate(row, start=1):
            c = wsn.cell(row=i, column=j, value=v)
            c.font = BLACK
            if j > 1:
                c.number_format = MONEY
    wsn.column_dimensions["A"].width = 42
    for col in "BCD":
        wsn.column_dimensions[col].width = 20

    wm = wb.create_sheet("Monte_Carlo")
    wm["A1"] = "Monte Carlo uncertainty analysis, 5,000 runs (values computed by rem_hybrid_model.py)"
    wm["A1"].font = TITLE
    stats = [("Mean NPV (A$)", mc_df.NPV.mean(), MONEY), ("P10 NPV (A$)", mc_df.NPV.quantile(0.10), MONEY),
             ("P50 NPV (A$)", mc_df.NPV.quantile(0.50), MONEY), ("P90 NPV (A$)", mc_df.NPV.quantile(0.90), MONEY),
             ("Probability NPV > 0", (mc_df.NPV > 0).mean(), "0.0%"), ("Mean IRR", mc_df.IRR.mean(), "0.00%"),
             ("P50 LCOE (A$/MWh)", mc_df.LCOE.quantile(0.5), "0.00")]
    for i, (lab, v, nf) in enumerate(stats, start=3):
        wm.cell(row=i, column=1, value=lab).font = BOLD
        c = wm.cell(row=i, column=2, value=float(v))
        c.number_format = nf
    wm["D3"] = "Distributions: wind speed N(1,5%), GHI N(1,3%), MLF Tri(0.86,0.93,0.97), price Tri(0.75,1,1.2), CAPEX Tri(0.9,1,1.2), discount rate +N(0,1 pp), OPEX Tri(0.9,1,1.25)"
    hdr(wm, 12, ["Run", "NPV (A$)", "IRR", "LCOE (A$/MWh)"])
    for i, row in enumerate(mc_df.head(1000).itertuples(), start=13):
        wm.cell(row=i, column=1, value=row.Index + 1)
        wm.cell(row=i, column=2, value=float(row.NPV)).number_format = MONEY
        wm.cell(row=i, column=3, value=float(row.IRR)).number_format = "0.00%"
        wm.cell(row=i, column=4, value=float(row.LCOE)).number_format = "0.00"
    wm.cell(row=11, column=1, value="First 1,000 of 5,000 runs listed below.").font = BLACK
    wm.column_dimensions["A"].width = 26
    for col in "BCD":
        wm.column_dimensions[col].width = 18

    # fonts everywhere to Arial
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.font is None or c.font.name != "Arial":
                    c.font = Font(name="Arial", size=10, bold=c.font.bold if c.font else False,
                                  color=c.font.color if c.font else None)
    wb.save(OUT_XLSX)
    return OUT_XLSX


# ----------------------------------------------------------------------------------
# 6. FIGURES
# ----------------------------------------------------------------------------------
def figures(scen, tor_base, tor_df, mc_df):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams.update({"font.family": "calibri", "font.size": 10.5, "axes.spines.top": False,
                         "axes.spines.right": False, "axes.grid": True, "grid.color": "#E5E5E5", "axes.titlecolor": "#007B7B", "axes.edgecolor": "#7F7F7F"})
    COLS = ["#007B7B", "#E87722", "#66B2B2", "#7F7F7F"]

    # Fig 1: energy by source over life (base)
    tb = scen["Base Case"]["table"].iloc[1:]
    fig, ax = plt.subplots(figsize=(7.5, 3.6))
    ax.bar(tb.Year, tb["Wind MWh"] / 1e3, color="#007B7B", label="Wind")
    ax.bar(tb.Year, tb["PV MWh"] / 1e3, bottom=tb["Wind MWh"] / 1e3, color="#E87722", label="Solar PV")
    ax.plot(tb.Year, tb["Net export MWh"] / 1e3, color="#222", lw=1.5, label="Net export (after BESS losses)")
    ax.set_xlabel("Operating year")
    ax.set_ylabel("GWh/year")
    ax.set_title("Annual generation by source, base case", loc="left", fontsize=11)
    ax.set_ylim(0, 520)
    ax.legend(frameon=False, ncol=3, loc="upper center")
    fig.tight_layout()
    fig.savefig("fig1_generation.png", dpi=200)
    plt.close(fig)

    # Fig 2: cumulative cash flow by scenario
    fig, ax = plt.subplots(figsize=(7.5, 3.8))
    for (name, R), c in zip(scen.items(), COLS):
        ax.plot(R["table"].Year, R["table"]["Cumulative DCF"] / 1e6, color=c, lw=2, label=name)
    ax.axhline(0, color="#555", lw=0.8)
    ax.set_xlabel("Year")
    ax.set_ylabel("Cumulative discounted cash flow (A$ million)")
    ax.set_title("Cumulative discounted cash flow by scenario", loc="left", fontsize=11)
    ax.legend(frameon=False, loc="upper left")
    fig.tight_layout()
    fig.savefig("fig2_cumulative_dcf.png", dpi=200)
    plt.close(fig)

    # Fig 3: tornado
    d = tor_df.iloc[::-1]
    fig, ax = plt.subplots(figsize=(7.5, 3.9))
    y = np.arange(len(d))
    lo = np.minimum(d["NPV adverse (A$)"], d["NPV favourable (A$)"]) / 1e6
    hi = np.maximum(d["NPV adverse (A$)"], d["NPV favourable (A$)"]) / 1e6
    b = tor_base / 1e6
    ax.barh(y, b - lo, left=lo, color="#E87722", label="Adverse")
    ax.barh(y, hi - b, left=b, color="#007B7B", label="Favourable")
    ax.axvline(b, color="#222", lw=1)
    ax.axvline(0, color="#888", lw=0.8, ls="--")
    ax.set_yticks(y)
    ax.set_yticklabels(d["Variable"])
    ax.set_xlabel("NPV (A$ million)")
    ax.set_title(f"NPV sensitivity (base NPV = -A${abs(b):,.1f} M)", loc="left", fontsize=11)
    ax.legend(frameon=False, loc="lower right")
    ax.grid(axis="y", visible=False)
    fig.tight_layout()
    fig.savefig("fig3_tornado.png", dpi=200)
    plt.close(fig)

    # Fig 4: Monte Carlo histogram
    fig, ax = plt.subplots(figsize=(7.5, 3.6))
    ax.hist(mc_df.NPV / 1e6, bins=60, color="#007B7B", alpha=0.85)
    ax.axvline(0, color="#E87722", lw=1.6, label="NPV = 0")
    for q, ls in ((0.1, ":"), (0.5, "--"), (0.9, ":")):
        ax.axvline(mc_df.NPV.quantile(q) / 1e6, color="#222", ls=ls, lw=1)
    ax.set_xlabel("NPV (A$ million)")
    ax.set_ylabel("Runs")
    ax.set_title(f"Monte Carlo NPV distribution (5,000 runs; P(NPV>0) = {(mc_df.NPV > 0).mean():.0%})",
                 loc="left", fontsize=11)
    ax.legend(frameon=False)
    fig.tight_layout()
    fig.savefig("fig4_montecarlo.png", dpi=200)
    plt.close(fig)


# ----------------------------------------------------------------------------------
if __name__ == "__main__":
    scen = {name: run(BASE, **lev) for name, lev in SCENARIOS}
    T = scen["Base Case"]["T"]
    cx = capex_items(BASE)
    print("=== TECHNICAL ===")
    for k in ("c", "gross_wind", "cf_wind_gross", "wind_loss", "wind_e1", "cf_wind", "poa", "pr", "pv_e1", "cf_pv", "bess_d1"):
        print(f"{k:15s} {T[k]:,.4f}")
    print("=== CAPEX ===")
    for k, v in cx.items():
        print(f"{k:8s} {v:,.0f}")
    print("=== SCENARIOS ===")
    summ = pd.DataFrame({n: {"Gen y1 MWh": R["gen1"], "Gen avg MWh": R["gen_avg"], "Rev y1": R["rev1"],
                             "Rev avg": R["rev_avg"], "OPEX y1": R["opex1"], "CAPEX": R["capex"], "NPV": R["npv"],
                             "IRR": R["irr"], "LCOE": R["lcoe"], "Payback": R["pb"], "Disc payback": R["dpb"],
                             "CO2 y1": R["co2_1"], "CO2 life": R["co2_life"], "CO2 life (decl EF)": R["co2_life_decl"],
                             "Gen life": R["gen_life"], "BE PPA": R["be_ppa"]} for n, R in scen.items()})
    pd.set_option("display.float_format", lambda x: f"{x:,.4f}")
    print(summ)
    tb, tor = tornado(BASE)
    print(tor)
    mc = monte_carlo(BASE)
    print(mc.describe(percentiles=[0.1, 0.5, 0.9]))
    print("P(NPV>0) =", (mc.NPV > 0).mean())
    print(scen["Base Case"]["table"].iloc[:3].T)
    build_workbook(scen["Base Case"], scen, tor, mc)
    figures(scen, tb, tor, mc)
    summ.to_csv("scenario_summary_python.csv")
    import json
    comb = run(BASE, price_f=1.2, capex_f=0.85)
    js = dict(T={k: float(v) for k, v in T.items()}, capex={k: float(v) for k, v in cx.items()},
              scen={n: {k: (float(v) if not isinstance(v, (pd.DataFrame, dict)) else None) for k, v in R.items() if k not in ("table", "T")} for n, R in scen.items()},
              comb=dict(npv=comb["npv"], irr=comb["irr"], pb=comb["pb"], dpb=comb["dpb"], lcoe=comb["lcoe"]),
              y1=scen["Base Case"]["table"].iloc[1].to_dict(), y12=scen["Base Case"]["table"].iloc[12].to_dict(),
              y25=scen["Base Case"]["table"].iloc[25].to_dict(),
              mc=dict(mean=mc.NPV.mean(), p10=mc.NPV.quantile(.1), p50=mc.NPV.quantile(.5), p90=mc.NPV.quantile(.9),
                      pos=(mc.NPV > 0).mean(), irr=mc.IRR.mean(), lcoe50=mc.LCOE.quantile(.5)),
              tornado=tor.to_dict(orient="records"),
              cf=scen["Base Case"]["table"][["Year", "Net export MWh", "Total revenue", "Total OPEX", "CAPEX", "Net cash flow", "Discounted CF", "Cumulative CF", "Cumulative DCF"]].to_dict(orient="records"))
    json.dump(js, open("results.json", "w"), indent=1, default=float)
    tor.to_csv("tornado_python.csv", index=False)
    print("Saved", OUT_XLSX)
