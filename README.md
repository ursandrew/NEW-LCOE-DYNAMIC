# energy-optimizer-pro

Standalone LCOE sizing tool for Solar PV + Wind, replicating the detailed
CAPEX/OPEX anchor-point methodology of Solar_3.xlsm-style Excel workbooks
directly in Python.

**Scope (v1):** PV + Wind only. Not integrated with the main Energy Modeling
Optimizer — this is a separate tool. BESS and Hydro detailed cost breakdowns
are planned; once added, this becomes the intended successor to the main EMO
tool.

## Files
- `cost_engine.py` — anchor-point interpolation + category CAPEX/OPEX engine,
  validated exact-match against Solar_3.xlsm's computation table.
- `lcoe_engine.py` — growing-annuity LCOE formula, reverse-engineered and
  validated to machine precision against the workbook's own LCOE column.
  NOT the HOMER Pro methodology used by the main EMO tool — this tool
  intentionally replicates the Excel, not HOMER.
- `excel_loader.py` — loads CAPEX/OPEX category tables directly from an
  uploaded `.xlsm` matching this structure; run directly (`python
  excel_loader.py <path>`) to re-validate against a workbook.
  Also runs `python lcoe_engine.py` to validate the LCOE formula standalone.
- `app.py` — Streamlit UI: editable CAPEX/OPEX category tables (Civil,
  Mechanical, Electrical, Project Mgmt, Misc) for Solar and Wind, a
  target-energy grid search across capacity ranges, and Solar / Wind /
  Hybrid LCOE output for direct comparison against the source Excel.

## Validation status
- CAPEX/OPEX: exact match (< $0.01 diff) at 50/60/70/100 MWp solar.
- LCOE: exact match (diff ~1e-13) at the same points.
- Wind: CAPEX matches the workbook's own summary totals at 100/200/500 MW;
  full anchor-point LCOE cross-check pending (Wind sheet's computation table
  in the source file wasn't populated with run data at time of writing).

## Run
```
pip install -r requirements.txt
streamlit run app.py
```

## Known differences from the Excel (intentional improvements)
- Denominator/zero guard on rate calculations (source Excel Flag #2).
- Base EPC Cost composition is explicit per category rather than assumed —
  confirmed Solar excludes Mechanical + Misc, Wind excludes only Misc
  (source Excel Flag #3).
- Arbitrary anchor point counts/spacing supported, not fixed to the
  original workbook's 6 points.

Author: SJ | 2026
