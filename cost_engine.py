"""
cost_engine.py — Detailed CAPEX/OPEX calculation engine for energy-optimizer-pro

Reproduces the Solar_3.xlsm anchor-point CAPEX/OPEX methodology as native
Python, for both Solar (PV) and Wind. Unlike the Excel workbook, this module:

  - Uses piecewise-linear interpolation between whatever anchor capacities are
    supplied (any number >= 2), not a fixed 6-point table.
  - Computes "% of Base EPC Cost" items (Project Mgmt category) AFTER Base EPC
    (Civil + Mechanical + Electrical + Misc) is resolved, matching the Excel's
    actual order of operations.
  - Adds a denominator guard so a zero/negative (discount - inflation) never
    produces a divide-by-zero (this was Excel Flag #2 from the workbook audit).
  - Returns a full category breakdown, not just a single CAPEX/OPEX number, so
    UI and downstream LCOE code can show line-item detail.

Author: SJ | 2026
"""

from dataclasses import dataclass, field
from typing import List, Dict


# ==============================================================================
# DATA MODEL
# ==============================================================================

@dataclass
class CostItem:
    """A single CAPEX or OPEX line item, priced across N anchor capacities."""
    name: str
    uom: str                 # '$/Wp', '$/kW', 'Lump Sum (USD)', '% of Base EPC Cost',
                              # '% of CAPEX/year', '$/Wp/year', '$/kW/year', 'Lump Sum/year'
    values: List[float]      # one value per anchor capacity, same order as anchors_mw

    def interpolated_value(self, capacity_mw: float, anchors_mw: List[float]) -> float:
        return _piecewise_linear_interp(capacity_mw, anchors_mw, self.values)


@dataclass
class CostCategory:
    """
    A named group of line items, e.g. 'Civil & Construction'.

    counts_toward_base_epc: whether this category's subtotal feeds into
        "Base EPC Cost" (the base that % categories, e.g. Project Mgmt fees,
        are computed against). This is workbook-specific — Solar_3.xlsm
        excludes Mechanical AND Misc from Base EPC; the Wind sheet excludes
        only Misc. Set explicitly per category rather than assumed, since
        this was Excel Flag #3 from the source-workbook audit and silently
        assuming a rule here would reproduce the ambiguity, not fix it.
    """
    label: str
    items: List[CostItem] = field(default_factory=list)
    counts_toward_base_epc: bool = True


@dataclass
class DetailedCostConfig:
    """Full detailed cost definition for one technology (Solar or Wind)."""
    anchors_mw: List[float]              # anchor capacities, ascending, e.g. [50,55,60,65,70,130]
    capex_unit_basis: str                # '$/Wp' or '$/kW' — controls unit scaling
    capex_categories: List[CostCategory] # order matters: last category assumed = Project Mgmt
                                          # if it contains '% of Base EPC Cost' items
    opex_items: List[CostItem]           # flat list, OPEX Year-1 items
    project_mgmt_category_label: str = "Project Management"


# ==============================================================================
# INTERPOLATION
# ==============================================================================

def _piecewise_linear_interp(x: float, xs: List[float], ys: List[float]) -> float:
    """
    Piecewise-linear interpolation across arbitrary anchor points.
    Clamps (does not extrapolate) outside [min(xs), max(xs)] — flat-line
    beyond the range, with a caller-facing flag to warn on this (see
    compute_detailed_capex_opex's 'extrapolated' return field).
    """
    n = len(xs)
    if n == 1:
        return ys[0]
    if x <= xs[0]:
        return ys[0]
    if x >= xs[-1]:
        return ys[-1]
    for i in range(n - 1):
        x0, x1 = xs[i], xs[i + 1]
        if x0 <= x <= x1:
            if x1 == x0:
                return ys[i]
            frac = (x - x0) / (x1 - x0)
            return ys[i] + frac * (ys[i + 1] - ys[i])
    return ys[-1]  # unreachable, safety fallback


# ==============================================================================
# CORE CALCULATION
# ==============================================================================

def _resolve_item_dollar_value(item: CostItem, capacity_mw: float,
                                capacity_base_unit: float, anchors_mw: List[float],
                                base_epc_cost: float = None) -> float:
    """
    Convert one interpolated rate into a dollar amount for the given capacity.

    capacity_base_unit: capacity expressed in the item's UOM base unit
                         (Wp for '$/Wp' items, kW for '$/kW' items).
    base_epc_cost: required only for '% of Base EPC Cost' / '% of CAPEX/year' items.
    """
    rate = item.interpolated_value(capacity_mw, anchors_mw)
    uom = item.uom.strip().lower()

    if uom in ('$/wp', '$/wp/year'):
        return rate * capacity_base_unit
    if uom in ('$/kw', '$/kw/year', '$/kw-yr'):
        return rate * capacity_base_unit
    if uom.startswith('lump sum'):
        return rate
    if uom.startswith('% of base epc') or uom.startswith('% of capex'):
        if base_epc_cost is None:
            raise ValueError(
                f"Item '{item.name}' needs base_epc_cost to resolve a percentage rate, "
                f"but none was supplied."
            )
        return rate * base_epc_cost
    raise ValueError(f"Unrecognized UOM '{item.uom}' on item '{item.name}'")


def compute_detailed_capex(config: DetailedCostConfig, capacity_mw: float) -> Dict:
    """
    Compute full CAPEX breakdown for a given capacity (MW / MWp).

    Returns dict with per-category subtotals, grand total, per-item detail,
    and an 'extrapolated' flag if capacity fell outside the anchor range.
    """
    anchors = config.anchors_mw
    extrapolated = capacity_mw < anchors[0] or capacity_mw > anchors[-1]

    # capacity in the base unit used by $/Wp or $/kW rates
    if config.capex_unit_basis == '$/Wp':
        capacity_base_unit = capacity_mw * 1_000_000   # MW -> Wp
    elif config.capex_unit_basis == '$/kW':
        capacity_base_unit = capacity_mw * 1_000       # MW -> kW
    else:
        raise ValueError(f"Unsupported capex_unit_basis: {config.capex_unit_basis}")

    # --- Pass 1: resolve every non-percentage category (Base EPC components) ---
    category_results = {}
    base_epc_cost = 0.0
    pct_categories = []  # deferred to pass 2

    for cat in config.capex_categories:
        has_pct_items = any(i.uom.strip().lower().startswith('% of base epc')
                             for i in cat.items)
        if has_pct_items:
            pct_categories.append(cat)
            continue

        item_values = {}
        subtotal = 0.0
        for item in cat.items:
            val = _resolve_item_dollar_value(item, capacity_mw, capacity_base_unit, anchors)
            item_values[item.name] = val
            subtotal += val

        category_results[cat.label] = {'items': item_values, 'subtotal': subtotal}
        if cat.counts_toward_base_epc:
            base_epc_cost += subtotal

    # --- Pass 2: resolve percentage-of-Base-EPC categories (e.g. Project Mgmt) ---
    for cat in pct_categories:
        item_values = {}
        subtotal = 0.0
        for item in cat.items:
            val = _resolve_item_dollar_value(item, capacity_mw, capacity_base_unit,
                                              anchors, base_epc_cost=base_epc_cost)
            item_values[item.name] = val
            subtotal += val
        category_results[cat.label] = {'items': item_values, 'subtotal': subtotal}

    grand_total = sum(c['subtotal'] for c in category_results.values())

    return {
        'capacity_mw': capacity_mw,
        'base_epc_cost': base_epc_cost,
        'categories': category_results,
        'grand_total_capex': grand_total,
        'capex_per_base_unit': grand_total / capacity_base_unit if capacity_base_unit else 0.0,
        'extrapolated': extrapolated,
    }


def compute_detailed_opex(config: DetailedCostConfig, capacity_mw: float,
                           grand_total_capex: float) -> Dict:
    """
    Compute Year-1 OPEX breakdown. '% of CAPEX/year' items need the CAPEX
    grand total already computed (pass it in from compute_detailed_capex).
    """
    anchors = config.anchors_mw
    if config.capex_unit_basis == '$/Wp':
        capacity_base_unit = capacity_mw * 1_000_000
    else:
        capacity_base_unit = capacity_mw * 1_000

    item_values = {}
    total = 0.0
    for item in config.opex_items:
        val = _resolve_item_dollar_value(item, capacity_mw, capacity_base_unit,
                                          anchors, base_epc_cost=grand_total_capex)
        item_values[item.name] = val
        total += val

    return {'capacity_mw': capacity_mw, 'items': item_values, 'grand_total_opex_year1': total}


def compute_detailed_capex_opex(config: DetailedCostConfig, capacity_mw: float) -> Dict:
    """
    Convenience wrapper: returns (capital_$, om_annual_$, full_detail) —
    this is the exact interface the existing NPC/LCOE engines expect
    (calculate_npc_homer_style / manager's DF cash-flow function), so this
    function is the single integration point into your existing LCOE code.
    """
    capex = compute_detailed_capex(config, capacity_mw)
    opex = compute_detailed_opex(config, capacity_mw, capex['grand_total_capex'])
    return {
        'capital': capex['grand_total_capex'],
        'om_annual': opex['grand_total_opex_year1'],
        'capex_detail': capex,
        'opex_detail': opex,
    }


# ==============================================================================
# DENOMINATOR GUARD (Excel Flag #2 fix) — reused by LCOE calcs downstream
# ==============================================================================

def safe_real_discount_rate(nominal_rate: float, inflation_rate: float,
                             min_gap: float = 1e-6) -> float:
    """
    Fisher equation with a guard: if (1+nominal)/(1+inflation) collapses toward
    1.0 (i.e. nominal ≈ inflation), the real rate approaches 0, which is fine.
    The actual failure mode in the Excel was inflation > discount rate causing
    a negative or zero denominator elsewhere in its NPV formulas with no check.
    This helper doesn't hide that condition — it surfaces it explicitly.
    """
    real_rate = (1 + nominal_rate) / (1 + inflation_rate) - 1
    if abs(1 + real_rate) < min_gap:
        raise ValueError(
            f"Real discount rate ({real_rate:.6f}) is degenerate (~-100%). "
            f"Check nominal_rate={nominal_rate} and inflation_rate={inflation_rate}."
        )
    return real_rate
