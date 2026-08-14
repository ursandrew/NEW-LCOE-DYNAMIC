"""
lcoe_engine.py — Excel-matching LCOE calculation (growing annuity method).

This tool is a standalone replication of Solar_3.xlsm's own financial logic
in Python — NOT a HOMER Pro cross-check like the main EMO tool. The formula
below is reverse-engineered and validated exact-match against the workbook's
own COMPUTATION TABLE at every anchor point (run this file directly to see
the validation output).

Excel's methodology, confirmed formula-by-formula:
    NPV_CAPEX  = CAPEX (Year-0 lump sum, not discounted)
    NPV_OPEX   = growing annuity PV of Year-1 OPEX, escalating at inflation,
                 discounted at the nominal discount rate
    NPV_Energy = growing annuity PV of Year-1 energy, "growing" at
                 (-degradation rate), discounted at the same nominal rate
    LCOE       = (NPV_CAPEX + NPV_OPEX) / NPV_Energy

    Growing annuity closed form:
        PV = CF1 / (r - g) * [1 - ((1+g)/(1+r))^n]          (r != g)
        PV = CF1 * n / (1+r)                                 (r == g, limit case)

Single nominal discount rate throughout — no real-rate conversion, no CRF,
no replacement/salvage (the Excel doesn't model component replacement at
all, consistent with Flags #5/#6 from the earlier workbook audit).

Author: SJ | 2026
"""

from dataclasses import dataclass


def pv_growing_annuity(cf1: float, growth_rate: float, discount_rate: float, n: int) -> float:
    """PV of a cash flow starting at cf1 in year 1, growing at growth_rate
    each year, discounted at discount_rate, for n years."""
    if abs(discount_rate - growth_rate) < 1e-12:
        return cf1 * n / (1 + discount_rate)
    return cf1 / (discount_rate - growth_rate) * (1 - ((1 + growth_rate) / (1 + discount_rate)) ** n)


@dataclass
class LCOEResult:
    technology: str
    capex: float
    opex_year1: float
    energy_year1_kwh: float
    npv_capex: float
    npv_opex: float
    npv_energy_kwh: float
    lcoe_per_kwh: float
    lcoe_per_mwh: float


def _component_lcoe(technology: str, capex: float, opex_year1: float,
                     energy_year1_kwh: float, discount_rate: float,
                     inflation_rate: float, degradation_rate: float,
                     project_lifetime: int) -> LCOEResult:
    npv_capex = capex
    npv_opex = pv_growing_annuity(opex_year1, inflation_rate, discount_rate, project_lifetime)
    npv_energy = pv_growing_annuity(energy_year1_kwh, -degradation_rate, discount_rate, project_lifetime)
    lcoe_per_kwh = (npv_capex + npv_opex) / npv_energy if npv_energy else 0.0

    return LCOEResult(
        technology=technology, capex=capex, opex_year1=opex_year1,
        energy_year1_kwh=energy_year1_kwh,
        npv_capex=npv_capex, npv_opex=npv_opex, npv_energy_kwh=npv_energy,
        lcoe_per_kwh=lcoe_per_kwh, lcoe_per_mwh=lcoe_per_kwh * 1000,
    )


def solar_lcoe(capex, opex_year1, energy_year1_kwh, discount_rate, inflation_rate,
               degradation_rate, project_lifetime) -> LCOEResult:
    return _component_lcoe('Solar PV', capex, opex_year1, energy_year1_kwh,
                            discount_rate, inflation_rate, degradation_rate, project_lifetime)


def wind_lcoe(capex, opex_year1, energy_year1_kwh, discount_rate, inflation_rate,
              degradation_rate, project_lifetime) -> LCOEResult:
    return _component_lcoe('Wind', capex, opex_year1, energy_year1_kwh,
                            discount_rate, inflation_rate, degradation_rate, project_lifetime)


def hybrid_lcoe(solar_result: LCOEResult, wind_result: LCOEResult) -> LCOEResult:
    """Sum NPV(CAPEX+OPEX) and NPV(Energy) across both, then divide — not a
    simple average of the two LCOEs (that would over-weight the smaller one)."""
    total_npv_capex = solar_result.npv_capex + wind_result.npv_capex
    total_npv_opex = solar_result.npv_opex + wind_result.npv_opex
    total_npv_energy = solar_result.npv_energy_kwh + wind_result.npv_energy_kwh
    lcoe_per_kwh = (total_npv_capex + total_npv_opex) / total_npv_energy if total_npv_energy else 0.0

    return LCOEResult(
        technology='Hybrid (Solar + Wind)',
        capex=solar_result.capex + wind_result.capex,
        opex_year1=solar_result.opex_year1 + wind_result.opex_year1,
        energy_year1_kwh=solar_result.energy_year1_kwh + wind_result.energy_year1_kwh,
        npv_capex=total_npv_capex, npv_opex=total_npv_opex, npv_energy_kwh=total_npv_energy,
        lcoe_per_kwh=lcoe_per_kwh, lcoe_per_mwh=lcoe_per_kwh * 1000,
    )


if __name__ == '__main__':
    ground_truth = {
        50:  {'capex': 25366171.3,  'opex_y1': 917394.0556,  'energy_y1': 85000000.2499999,  'lcoe': 0.04410606664284536},
        60:  {'capex': 30398101.3,  'opex_y1': 1077777.2156, 'energy_y1': 102000000.29999988, 'lcoe': 0.043659862782946174},
        70:  {'capex': 35264831.3,  'opex_y1': 1236177.9755999998, 'energy_y1': 119000000.34999986, 'lcoe': 0.04319717211403093},
        100: {'capex': 49538271.79, 'opex_y1': 1715459.26148, 'energy_y1': 170000000.4999998, 'lcoe': 0.04225084216453368},
    }
    print("=" * 70)
    print("VALIDATION: lcoe_engine vs. Solar_3.xlsm LCOE column")
    print("=" * 70)
    for cap, truth in ground_truth.items():
        result = solar_lcoe(truth['capex'], truth['opex_y1'], truth['energy_y1'],
                             discount_rate=0.06, inflation_rate=0.055, degradation_rate=0.005,
                             project_lifetime=25)
        diff = result.lcoe_per_kwh - truth['lcoe']
        status = "OK" if abs(diff) < 1e-6 else "MISMATCH"
        print(f"{cap} MWp [{status}]  engine={result.lcoe_per_kwh:.10f}  excel={truth['lcoe']:.10f}  diff={diff:.2e}")
