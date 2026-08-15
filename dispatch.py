"""
dispatch.py — Profile-based Solar PV sizing via target unmet load %.

Mirrors the main Energy Modeling Optimizer's exact methodology (see
optimize_gridsearch_hydro_WITH_DEGRADATION.py):
    - The 'Output_kW' profile is a per-1-kW normalized specific-yield curve
      (solar_config['baseline_kw'] = 1.0 in the main tool) — values are
      output per kW of installed capacity, scaled to any candidate capacity
      by a direct multiply (capacity_kW * profile[h]), not a baseline ratio.
    - Unmet load per hour = max(0, load - generation).
    - unmet_% = sum(unmet) / sum(load) * 100.
    - A capacity is "feasible" if unmet_% <= target_unmet_%.
    - Grid search picks the SMALLEST feasible capacity (cheapest system that
      still meets the reliability target — larger systems cost more even
      though their LCOE is lower, per the anchor-point cost curve).

This intentionally replaces the earlier "target annual energy (GWh)"
constraint, which had no basis in either the Excel workbook (which has no
optimization objective at all — it just reports LCOE for a given capacity)
or the main EMO tool's actual sizing framework.

Solar-only for now. Wind to follow once this is validated.

Author: SJ | 2026
"""

import numpy as np
import pandas as pd
from dataclasses import dataclass


def load_hourly_profile(csv_path_or_buffer) -> np.ndarray:
    """
    Load an 8760-hour profile from CSV, matching the main EMO tool's
    convention: data is read from the SECOND column (iloc[:, 1]).
    """
    df = pd.read_csv(csv_path_or_buffer)
    values = df.iloc[:, 1].to_numpy(dtype=float)
    return values


@dataclass
class DispatchResult:
    capacity_mw: float
    hourly_generation_kwh: np.ndarray
    hourly_unmet_kwh: np.ndarray
    total_generation_kwh: float
    total_load_kwh: float
    total_unmet_kwh: float
    unmet_percent: float


def simulate_solar_only_dispatch(load_profile_kwh: np.ndarray, pv_profile_per_kw: np.ndarray,
                                  candidate_capacity_mw: float) -> DispatchResult:
    """
    Scale the PV profile to candidate_capacity_mw and compute unmet load
    against the load profile. Matches the main EMO tool's exact convention
    (optimize_gridsearch_hydro_WITH_DEGRADATION.py, solar_config['baseline_kw']
    = 1.0): the 'Output_kW' profile is a per-1-kW normalized specific-yield
    curve (values ~0-1, representing output per kW of installed capacity),
    NOT a full generation curve for some reference-size plant. Scaling is
    therefore a direct multiply by capacity in kW — no baseline ratio needed,
    since the baseline is always 1 kW.

    Solar-only: no wind/hydro/BESS in this dispatch yet, so unmet load is
    simply max(0, load - pv_generation) each hour.
    """
    candidate_capacity_kw = candidate_capacity_mw * 1000
    pv_gen = pv_profile_per_kw * candidate_capacity_kw
    unmet = np.maximum(0.0, load_profile_kwh - pv_gen)

    total_gen = float(pv_gen.sum())
    total_load = float(load_profile_kwh.sum())
    total_unmet = float(unmet.sum())
    unmet_pct = (total_unmet / total_load * 100) if total_load > 0 else 0.0

    return DispatchResult(
        capacity_mw=candidate_capacity_mw,
        hourly_generation_kwh=pv_gen,
        hourly_unmet_kwh=unmet,
        total_generation_kwh=total_gen,
        total_load_kwh=total_load,
        total_unmet_kwh=total_unmet,
        unmet_percent=unmet_pct,
    )


def find_min_capacity_meeting_target(load_profile_kwh: np.ndarray, pv_profile_per_kw: np.ndarray,
                                      target_unmet_percent: float,
                                      capacity_min_mw: float, capacity_max_mw: float,
                                      capacity_step_mw: float) -> dict:
    """
    Grid search over candidate capacities; returns the smallest capacity
    whose unmet_% <= target, plus the full scan for charting/inspection.
    """
    n_steps = int(round((capacity_max_mw - capacity_min_mw) / capacity_step_mw)) + 1
    candidates_mw = [capacity_min_mw + i * capacity_step_mw for i in range(n_steps)]

    scan = []
    best = None
    for cap in candidates_mw:
        result = simulate_solar_only_dispatch(load_profile_kwh, pv_profile_per_kw, cap)
        scan.append(result)
        if result.unmet_percent <= target_unmet_percent and best is None:
            best = result  # candidates_mw ascending -> first feasible is smallest

    return {
        'best': best,          # None if no capacity in range meets the target
        'scan': scan,
        'feasible': best is not None,
    }
