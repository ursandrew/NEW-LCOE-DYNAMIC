# energy-optimizer-pro

Standalone LCOE sizing tool for Solar PV (Wind cost breakdown included,
Wind sizing to follow), replicating the detailed CAPEX/OPEX anchor-point
methodology of Solar_3.xlsm-style Excel workbooks directly in Python.

**Scope (v1):** Solar sizing + LCOE is fully wired. Wind CAPEX/OPEX tables
exist but are not yet part of the sizing search. BESS and Hydro detailed
cost breakdowns are planned; once added, this becomes the intended
successor to the main EMO tool.

## Files
- `cost_engine.py` — anchor-point interpolation + category CAPEX/OPEX engine,
  validated exact-match against Solar_3.xlsm's computation table.
- `lcoe_engine.py` — growing-annuity LCOE formula, reverse-engineered and
  validated to machine precision against the workbook's own LCOE column.
  NOT the HOMER Pro methodology used by the main EMO tool — this tool
  intentionally replicates the Excel, not HOMER. Run `python lcoe_engine.py`
  to re-validate the formula standalone.
- `excel_loader.py` — loads CAPEX/OPEX category tables directly from an
  uploaded `.xlsm` matching this structure; run `python excel_loader.py
  <path>` to re-validate against a workbook.
- `dispatch.py` — profile-based Solar sizing: scales an hourly PV profile
  by capacity (same convention as the main EMO tool — profile is a
  reference-capacity curve, scaled by `capacity / baseline`), computes
  hourly unmet load against an hourly Load profile, and grid-searches for
  the smallest PV capacity meeting a target unmet-load %.
- `app.py` — Streamlit UI: editable CAPEX/OPEX category tables for Solar
  and Wind, and a Results tab with two sizing modes:
  - **Profile-based** — upload Load + PV CSVs, set a target unmet-load %,
    get the smallest feasible capacity and its LCOE.
  - **Excel validation mode** — no profiles needed; reproduces the Excel's
    own anchor-point energy and LCOE exactly, for checking the cost/LCOE
    math in isolation before trusting profile-based results.

## Validation status
- CAPEX/OPEX: exact match (< $0.01 diff) at 50/60/70/100 MWp solar.
- LCOE: exact match (diff ~1e-13) at the same points.
- Wind: CAPEX matches the workbook's own summary totals at 100/200/500 MW;
  full anchor-point LCOE cross-check pending (Wind sheet's computation table
  in the source file wasn't populated with run data at time of writing).
  Wind is not yet part of the sizing search.

## A note on solar-only sizing
Without BESS, unmet load can never drop below roughly the night-time share
of annual load, no matter how large the PV plant is — zero generation at
night is a physical floor, not a sizing problem. Low unmet-load targets
(e.g. 4%) will correctly come back infeasible until BESS or Wind are added
to the dispatch. Use a loose target (50-60%) to validate the search
mechanics in the meantime; the default is set to 55% for this reason.

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
