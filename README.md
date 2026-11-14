# REM-A3-Hybrid-Wind-PV-Feasibility
# REM Assignment 3 – Hybrid Wind–PV + BESS Feasibility Study
150 MW hybrid (100 MW wind + 50 MWac PV) + 30 MW/60 MWh BESS, Central-West Orana REZ, NSW.

**Authors:** Fathurrahman Rusli Yakub (202522810), Tri Susilo Wirawan (202522670)

## Run order
1. `rem_hybrid_model.py` – main model: Excel file, 4 scenarios, tornado, Monte Carlo, Figures
2. `sizing.py` – sizing study (8 wind/PV/BESS configurations)
3. `fig_breakdown.py` – CAPEX and revenue breakdown chart
4. `make_diagrams.py` – single-line diagram
5. `fig_workflow_v2.py` – study workflow flowchart
6. `label_illustration.py` – labels the system illustration (needs `in/gemini_system.jpg`)

Requires: numpy, pandas, matplotlib, openpyxl, Pillow

Base case: NPV −A$97.5 M, IRR 5.75%, LCOE A$134.45/MWh, break-even PPA A$122.1/MWh.
