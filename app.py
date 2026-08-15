"""
energy-optimizer-pro — Detailed CAPEX/OPEX LCOE tool (Solar + Wind)

Companion to the main Energy Modeling Optimizer. This tool takes the same
granular CAPEX/OPEX line-item breakdown used in Solar_3.xlsm-style bid
costing (Civil / Mechanical / Electrical / Project Mgmt / Misc, anchor-point
interpolated) and uses it to size Solar + Wind and report Solar LCOE, Wind
LCOE, and blended Hybrid LCOE — directly cross-checkable against the source
Excel workbook at its own anchor capacities.

Scope (v1): PV + Wind only. BESS and Hydro to follow; once added, this tool
is intended to succeed the main EMO tool.

Author: SJ | 2026
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from io import BytesIO

from cost_engine import CostItem, CostCategory, DetailedCostConfig, compute_detailed_capex_opex
from lcoe_engine import solar_lcoe, wind_lcoe, hybrid_lcoe
from dispatch import load_hourly_profile, find_min_capacity_meeting_target
import excel_loader


# ==============================================================================
# PAGE CONFIG — same branding as Energy Modeling Optimizer
# ==============================================================================

st.set_page_config(page_title="Energy Optimizer Pro", page_icon="⚡",
                    layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<div style="display:flex;align-items:center;justify-content:center;gap:14px;margin-bottom:4px">
    <svg width="52" height="52" viewBox="0 0 52 52" xmlns="http://www.w3.org/2000/svg">
        <rect width="52" height="52" rx="0" fill="#0047AB"/>
        <text x="18" y="38" font-family="Arial,sans-serif" font-size="32"
              font-weight="bold" fill="white" text-anchor="middle">S</text>
        <text x="36" y="37" font-family="Arial,sans-serif" font-size="18"
              font-weight="bold" fill="white" text-anchor="middle">J</text>
        <circle cx="38" cy="18" r="4" fill="#E63946"/>
    </svg>
    <p style="font-size:2.5rem;font-weight:bold;color:#1f77b4;margin:0">
        Energy Optimizer Pro
    </p>
</div>
""", unsafe_allow_html=True)
st.markdown("**Detailed CAPEX/OPEX LCOE Sizing: Solar PV + Wind**")
st.caption("Python replication of Excel-style anchor-point cost workbooks (e.g. Solar_3.xlsm) — "
           "validated to machine precision against the source workbook's own LCOE column")
st.markdown("---")


# ==============================================================================
# SHARED CHART STYLING — explicit tick/gridline colors so axes stay readable
# regardless of the surrounding Streamlit theme (light or dark).
# ==============================================================================

def style_chart(fig, height=400):
    axis_style = dict(
        tickfont=dict(size=13, color='#1a1a1a'),
        title_font=dict(size=14, color='#1a1a1a'),
        gridcolor='#d5d5d5', zerolinecolor='#b0b0b0',
        showline=True, linecolor='#333333', linewidth=1,
    )
    fig.update_layout(
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(color='#1a1a1a', size=13),
        xaxis=axis_style, yaxis=axis_style,
        height=height, margin=dict(t=50, b=40, l=60, r=30),
    )
    return fig


SJ_PALETTE = ['#FDB462', '#80B1D3', '#8DD3C7', '#FB8072', '#BEBADA']


# ==============================================================================
# DEFAULT ITEM SCHEMAS — mirrors Solar_3.xlsm exactly so out-of-the-box output
# is directly comparable to the source workbook before any edits are made.
# ==============================================================================

def default_solar_config() -> DetailedCostConfig:
    anchors = [50, 55, 60, 65, 70, 130]
    civil = [
        CostItem('Site Clearance, Grading & Leveling', '$/Wp', [0.015]*6),
        CostItem('Module Mounting Structure (MMS) — Foundation', '$/Wp', [0.03]*6),
        CostItem('Land Cost', 'Lump Sum (USD)', [500000, 600000, 620000, 640000, 660000, 680000]),
        CostItem('Boundary Fencing & Access Gates', 'Lump Sum', [80000, 81000, 82000, 83000, 84000, 85000]),
        CostItem('Internal Roads & Drainage', 'Lump Sum (USD)', [120000, 130000, 140000, 150000, 160000, 170000]),
        CostItem('Control Room / Site Office — Civil Works', 'Lump Sum (USD)', [60000]*6),
        CostItem('Miscellaneous / Others', 'Lump Sum (USD)', [20, 40, 20, 20, 20, 35]),
    ]
    mechanical = [
        CostItem('Item1', '$/Wp', [0.015]*6),
        CostItem('Item2', '$/Wp', [0.03]*6),
        CostItem('Item3', 'Lump Sum (USD)', [500000, 600000, 620000, 640000, 660000, 680000]),
        CostItem('Item4', 'Lump Sum', [80000, 81000, 82000, 83000, 84000, 85000]),
        CostItem('Item5', 'Lump Sum (USD)', [120000, 130000, 140000, 150000, 160000, 170000]),
    ]
    electrical = [
        CostItem('PV Modules', '$/Wp', [0.18]*6),
        CostItem('Inverters', '$/Wp', [0.05]*6),
        CostItem('Module Mounting Structure (Steel)', '$/Wp', [0.06]*6),
        CostItem('DC Cabling', '$/Wp', [0.015]*6),
        CostItem('AC Cabling (LV/MV)', '$/Wp', [0.02]*6),
        CostItem('String Combiner Boxes (SCB)', '$/Wp', [0.008]*6),
        CostItem('Inverter Duty Transformer', '$/Wp', [0.015]*6),
        CostItem('Earthing & Lightning Protection', '$/Wp', [0.005]*6),
        CostItem('SCADA & Monitoring System', 'Lump Sum', [50000, 60000, 50000, 60000, 50000, 60000]),
        CostItem('Miscellaneous / Others', '$/Wp', [0]*6),
    ]
    project_mgmt = [
        CostItem('EPC / Project Management Fee', '% of Base EPC Cost', [0.03]*6),
        CostItem('Project Insurance (Construction Phase)', '% of Base EPC Cost', [0.005]*6),
        CostItem('Permits, Approvals & Statutory Fees', 'Lump Sum', [40000, 50000, 40000, 50000, 40000, 50000]),
        CostItem('Contingency', '% of Base EPC Cost', [0.03]*6),
        CostItem('Miscellaneous / Others', '$/Wp', [0]*6),
    ]
    misc = [
        CostItem('Transportation & Logistics', '$/Wp', [0.005]*6),
        CostItem('Construction Power & Water', 'Lump Sum', [20000, 30000, 20000, 30000, 20000, 30000]),
        CostItem('Security During Construction', 'Lump Sum', [15000, 25000, 15000, 25000, 15000, 25000]),
        CostItem('Testing & Commissioning', 'Lump Sum', [25000, 30000, 25000, 30000, 25000, 30000]),
        CostItem('Miscellaneous / Others', 'Lump Sum', [10000, 15000, 10000, 15000, 10000, 15000]),
    ]
    opex = [
        CostItem('O&M Contract (Fixed + Variable)', '$/Wp/year', [0.008]*6),
        CostItem('Insurance (Operational Phase)', '% of CAPEX/year', [0.004]*6),
        CostItem('Module Cleaning', '$/Wp/year', [0.002]*6),
        CostItem('Asset Management Fee', '% of CAPEX/year', [0.005]*6),
        CostItem('Spare Parts / Inverter Replacement Reserve', '% of CAPEX/year', [0.003]*6),
        CostItem('Land Lease', 'Lump Sum/year', [60000, 65000, 60000, 65000, 60000, 65000]),
        CostItem('Security', 'Lump Sum/year', [20000, 25000, 20000, 25000, 20000, 25000]),
        CostItem('Monitoring & Communication', 'Lump Sum/year', [8000, 9000, 8000, 9000, 8000, 9000]),
        CostItem('Vegetation Management', 'Lump Sum/year', [15000, 18000, 15000, 18000, 15000, 18000]),
        CostItem('Miscellaneous OPEX', 'Lump Sum (USD)/year', [10000, 12000, 10000, 12000, 10000, 12000]),
    ]
    return DetailedCostConfig(
        anchors_mw=anchors, capex_unit_basis='$/Wp',
        capex_categories=[
            CostCategory('Civil & Construction', civil, counts_toward_base_epc=True),
            CostCategory('Mechanical', mechanical, counts_toward_base_epc=False),
            CostCategory('Electrical', electrical, counts_toward_base_epc=True),
            CostCategory('Miscellaneous', misc, counts_toward_base_epc=False),
            CostCategory('Project Management', project_mgmt),
        ],
        opex_items=opex,
    )


def default_wind_config() -> DetailedCostConfig:
    anchors = [100, 150, 200, 300, 400, 500]
    turbine = [
        CostItem('Turbine Supply (Ex-Works)', '$/kW', [900, 900, 890, 880, 870, 860]),
        CostItem('Transportation & Logistics', '$/kW', [60, 60, 58, 56, 54, 52]),
        CostItem('Erection & Commissioning', '$/kW', [50, 50, 49, 48, 47, 46]),
        CostItem('Crane & Heavy Lift Equipment', 'Lump Sum', [800000, 850000, 900000, 1000000, 1100000, 1200000]),
        CostItem('Turbine Warranty / Extended Service', '$/kW', [30]*6),
    ]
    civil_bop = [
        CostItem('Foundation & Civil Works', '$/kW', [120, 118, 116, 114, 112, 110]),
        CostItem('Internal Roads & Access', 'Lump Sum', [1500000, 1800000, 2100000, 2600000, 3100000, 3600000]),
        CostItem('Site Clearance & Grading', 'Lump Sum', [500000, 600000, 700000, 850000, 1000000, 1150000]),
        CostItem('Array Cabling (Collector System)', '$/kW', [40, 40, 39, 38, 37, 36]),
        CostItem('Control Building / O&M Facility', 'Lump Sum', [600000, 600000, 650000, 700000, 750000, 800000]),
    ]
    electrical_grid = [
        CostItem('Substation Equipment', '$/kW', [70, 70, 68, 66, 64, 62]),
        CostItem('Transmission Line / Interconnection', 'Lump Sum', [2000000, 2200000, 2400000, 2800000, 3200000, 3600000]),
        CostItem('Metering & Protection', 'Lump Sum', [200000, 200000, 210000, 220000, 230000, 240000]),
        CostItem('SCADA & Communication System', 'Lump Sum', [300000, 300000, 310000, 320000, 330000, 340000]),
    ]
    project_mgmt = [
        CostItem('EPC / Project Management Fee', '% of Base EPC Cost', [0.03]*6),
        CostItem('Insurance (Construction Phase)', '% of Base EPC Cost', [0.005]*6),
        CostItem('Permits & Environmental Studies', 'Lump Sum', [400000, 420000, 440000, 470000, 500000, 530000]),
        CostItem('Contingency', '% of Base EPC Cost', [0.05]*6),
        CostItem('Miscellaneous / Others', '$/kW', [10]*6),
    ]
    misc = [
        CostItem('Transportation & Logistics', '$/kW', [0.005]*6),
        CostItem('Construction Power & Water', 'Lump Sum', [20000, 30000, 20000, 30000, 20000, 30000]),
        CostItem('Security During Construction', 'Lump Sum', [15000, 25000, 15000, 25000, 15000, 25000]),
        CostItem('Testing & Commissioning', 'Lump Sum', [25000, 30000, 25000, 30000, 25000, 30000]),
        CostItem('Miscellaneous / Others', 'Lump Sum', [10000, 15000, 10000, 15000, 10000, 15000]),
    ]
    opex = [
        CostItem('O&M Contract (Fixed + Variable)', '$/kW/year', [25, 25, 24, 23, 22, 21]),
        CostItem('Insurance (Operational Phase)', '% of CAPEX/year', [0.004]*6),
        CostItem('Land Lease', 'Lump Sum/year', [150000, 170000, 190000, 230000, 270000, 310000]),
        CostItem('Asset Management Fee', '% of CAPEX/year', [0.005]*6),
        CostItem('Spare Parts / Component Reserve', '% of CAPEX/year', [0.003]*6),
        CostItem('Grid / Transmission Charges', 'Lump Sum/year', [100000, 110000, 120000, 140000, 160000, 180000]),
        CostItem('Monitoring & SCADA', 'Lump Sum/year', [50000, 50000, 55000, 60000, 65000, 70000]),
        CostItem('Miscellaneous OPEX', 'Lump Sum/year', [40000, 42000, 44000, 48000, 52000, 56000]),
    ]
    return DetailedCostConfig(
        anchors_mw=anchors, capex_unit_basis='$/kW',
        capex_categories=[
            CostCategory('Turbine Supply & Installation', turbine, counts_toward_base_epc=True),
            CostCategory('Balance of Plant (Civil)', civil_bop, counts_toward_base_epc=True),
            CostCategory('Grid Connection & Substation', electrical_grid, counts_toward_base_epc=True),
            CostCategory('Miscellaneous', misc, counts_toward_base_epc=False),
            CostCategory('Project Management & Soft Costs', project_mgmt),
        ],
        opex_items=opex,
    )


# Default annual generation anchors (kWh/yr), from Solar_3.xlsm rows 15 —
# used to interpolate energy output at any capacity, same as cost items.
SOLAR_ENERGY_ANCHORS_KWH = [85000000.25, 93500000.27, 102000000.30, 110500000.32, 119000000.35, 221000000.65]
WIND_ENERGY_ANCHORS_KWH = [306600000, 459900000, 613200000, 919800000, 1226400000, 1533000000]


def interp_energy(capacity_mw, anchors_mw, energy_anchors_kwh):
    from cost_engine import _piecewise_linear_interp
    return _piecewise_linear_interp(capacity_mw, anchors_mw, energy_anchors_kwh)


# ==============================================================================
# SESSION STATE — load defaults once
# ==============================================================================

if 'solar_cfg' not in st.session_state:
    st.session_state.solar_cfg = default_solar_config()
if 'wind_cfg' not in st.session_state:
    st.session_state.wind_cfg = default_wind_config()


# ==============================================================================
# SIDEBAR — cost category editors + financial + search settings
# ==============================================================================

with st.sidebar:
    st.header("⚙️ Configuration")

    uploaded_xlsm = st.file_uploader("Load anchor costs from Excel (.xlsm)", type=['xlsm'])
    if uploaded_xlsm is not None and st.button("Load from Excel"):
        with open('/tmp/_uploaded.xlsm', 'wb') as f:
            f.write(uploaded_xlsm.read())
        st.session_state.solar_cfg = excel_loader.load_solar_config('/tmp/_uploaded.xlsm')
        st.session_state.wind_cfg = excel_loader.load_wind_config('/tmp/_uploaded.xlsm')
        st.success("Loaded anchor cost tables from workbook.")

    st.markdown("---")
    st.subheader("💰 Financial Parameters")
    st.caption("Matches the Excel's own methodology: single nominal discount rate, "
               "OPEX escalated by inflation, energy declines by degradation — no "
               "component replacement/salvage modeling (the Excel doesn't do this either).")
    project_lifetime = st.number_input("Project Lifetime (years)", value=25, min_value=1, max_value=50, step=1)
    col1, col2 = st.columns(2)
    with col1:
        solar_discount_pct = st.number_input("Solar Discount Rate (%)", value=6.0, min_value=0.0, max_value=20.0, step=0.1)
        solar_inflation_pct = st.number_input("Solar Inflation Rate (%)", value=5.5, min_value=0.0, max_value=20.0, step=0.1)
        solar_degradation_pct = st.number_input("Solar Degradation Rate (%/yr)", value=0.5, min_value=0.0, max_value=5.0, step=0.05)
    with col2:
        wind_discount_pct = st.number_input("Wind Discount Rate (%)", value=7.0, min_value=0.0, max_value=20.0, step=0.1)
        wind_inflation_pct = st.number_input("Wind Inflation Rate (%)", value=5.5, min_value=0.0, max_value=20.0, step=0.1)
        wind_degradation_pct = st.number_input("Wind Degradation Rate (%/yr)", value=0.3, min_value=0.0, max_value=5.0, step=0.05)

    st.markdown("---")
    st.subheader("🔎 Solar Sizing — Target Unmet Load")
    st.caption("Same framework as the main EMO tool: profiles scaled by capacity, "
               "sized to the smallest PV capacity meeting a target unmet-load %. "
               "Wind sizing will be added once this is validated.")

    load_file = st.file_uploader("Load Profile (CSV, 8760 hrs, value in 2nd column)", type=['csv'])
    pv_file = st.file_uploader("Solar PV Profile (CSV, 8760 hrs, value in 2nd column)", type=['csv'])
    st.caption("PV profile convention matches the main EMO tool: 'Output_kW' is a per-1-kW "
               "normalized specific-yield curve (values ~0-1) — generation scales as "
               "`profile[h] × capacity_kW` directly. No reference-capacity input needed.")
    target_unmet_pct = st.number_input(
        "Target Unmet Load (%)", value=55.0, min_value=0.0, max_value=100.0, step=1.0,
        help="Solar-only, no BESS: unmet load can never drop below roughly the "
             "night-time share of annual load, no matter how large the PV plant is "
             "(zero generation at night is a hard physical floor, not a sizing problem). "
             "Use a loose target like 50-60% for now to validate the search mechanics. "
             "Low targets (e.g. 4%) will correctly come back infeasible until BESS/Wind are added."
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        solar_min = st.number_input("Solar Search Min (MWp)", value=50.0, step=5.0)
    with col2:
        solar_max = st.number_input("Solar Search Max (MWp)", value=130.0, step=5.0)
    with col3:
        solar_step = st.number_input("Solar Search Step (MWp)", value=5.0, step=1.0)

    st.markdown("---")
    st.subheader("✅ Cross-Validation vs. Excel")
    energy_source = st.radio(
        "Annual energy used in the LCOE calculation:",
        ["From uploaded PV profile (actual dispatch)", "From Excel anchor interpolation (validation mode)"],
        help="Use profile-based energy for real sizing work. Switch to Excel anchor mode to reproduce "
             "Solar_3.xlsm's LCOE exactly at its own anchor capacities, for cross-checking the cost/LCOE math."
    )


# ==============================================================================
# MAIN — cost category editors (Solar / Wind tabs)
# ==============================================================================

def render_category_editor(cfg: DetailedCostConfig, tech_label: str, state_key: str):
    st.subheader(f"{tech_label} — CAPEX Categories")
    anchor_cols = [f"@{a:g}{'MWp' if cfg.capex_unit_basis=='$/Wp' else 'MW'}" for a in cfg.anchors_mw]

    for cat in cfg.capex_categories:
        with st.expander(f"{cat.label}  ({'Base EPC' if cat.counts_toward_base_epc else 'excl. Base EPC'})"):
            rows = [{'Item': i.name, 'UOM': i.uom, **dict(zip(anchor_cols, i.values))} for i in cat.items]
            df = pd.DataFrame(rows)
            edited = st.data_editor(df, key=f"{state_key}_{cat.label}", num_rows="fixed", use_container_width=True)
            for idx, item in enumerate(cat.items):
                item.values = [float(edited.iloc[idx][c]) for c in anchor_cols]

    st.subheader(f"{tech_label} — OPEX (Year 1) Items")
    rows = [{'Item': i.name, 'UOM': i.uom, **dict(zip(anchor_cols, i.values))} for i in cfg.opex_items]
    df = pd.DataFrame(rows)
    edited = st.data_editor(df, key=f"{state_key}_opex", num_rows="fixed", use_container_width=True)
    for idx, item in enumerate(cfg.opex_items):
        item.values = [float(edited.iloc[idx][c]) for c in anchor_cols]


tab_solar, tab_wind, tab_results = st.tabs(["☀️ Solar Costs", "💨 Wind Costs", "📊 Results"])

with tab_solar:
    render_category_editor(st.session_state.solar_cfg, "Solar PV", "solar")

with tab_wind:
    render_category_editor(st.session_state.wind_cfg, "Wind", "wind")

with tab_results:
    if st.button("🚀 Run Sizing Search", type="primary"):
        solar_cfg = st.session_state.solar_cfg

        use_profile = energy_source.startswith("From uploaded")

        if use_profile and (load_file is None or pv_file is None):
            st.error("Upload both a Load Profile and a Solar PV Profile CSV, or switch to "
                     "'From Excel anchor interpolation' mode to size without profiles.")
            st.stop()

        if use_profile:
            load_kwh = load_hourly_profile(load_file)
            pv_kwh = load_hourly_profile(pv_file)
            if len(load_kwh) != 8760 or len(pv_kwh) != 8760:
                st.warning(f"Expected 8760 hourly rows; got Load={len(load_kwh)}, PV={len(pv_kwh)}. "
                           f"Proceeding, but check the profiles if this wasn't intentional.")

            search = find_min_capacity_meeting_target(
                load_kwh, pv_kwh, target_unmet_pct,
                solar_min, solar_max, solar_step
            )
            if not search['feasible']:
                st.error(f"No capacity up to {solar_max:.0f} MWp meets {target_unmet_pct:.1f}% unmet load.\n\n"
                         f"For solar-only (no BESS), unmet load plateaus at roughly the night-time share of "
                         f"annual load — widening the capacity range further usually won't help. Try a "
                         f"looser target (e.g. 50-60%). The scan below shows where your search plateaued.")

                scan_df = pd.DataFrame([{'Capacity (MWp)': r.capacity_mw, 'Unmet Load (%)': r.unmet_percent}
                                         for r in search['scan']])
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=scan_df['Capacity (MWp)'], y=scan_df['Unmet Load (%)'],
                                          mode='lines+markers', line=dict(color='#E63946', width=3),
                                          marker=dict(size=8, color='#E63946')))
                fig.add_hline(y=target_unmet_pct, line_dash='dash', line_color='#1976D2', line_width=2,
                              annotation_text=f"Target: {target_unmet_pct:.1f}%",
                              annotation_font=dict(size=13, color='#1976D2'))
                style_chart(fig, height=380)
                fig.update_layout(xaxis_title="Solar Capacity (MWp)", yaxis_title="Unmet Load (%)")
                st.plotly_chart(fig, use_container_width=True)
                st.stop()

            best_dispatch = search['best']
            optimal_capacity = best_dispatch.capacity_mw
            annual_energy_kwh = best_dispatch.total_generation_kwh
            unmet_pct_achieved = best_dispatch.unmet_percent
        else:
            # Validation mode: no dispatch, just pick the smallest anchor capacity
            # (or let the user pick — for now, default to the smallest anchor point
            # so results are directly comparable to Excel's first row).
            optimal_capacity = solar_cfg.anchors_mw[0]
            annual_energy_kwh = interp_energy(optimal_capacity, solar_cfg.anchors_mw, SOLAR_ENERGY_ANCHORS_KWH)
            unmet_pct_achieved = None

        s_cost = compute_detailed_capex_opex(solar_cfg, optimal_capacity)
        s_lcoe = solar_lcoe(s_cost['capital'], s_cost['om_annual'], annual_energy_kwh,
                             solar_discount_pct/100, solar_inflation_pct/100,
                             solar_degradation_pct/100, project_lifetime)

        if use_profile:
            st.success(f"Smallest Solar capacity meeting {target_unmet_pct:.1f}% unmet load target: "
                       f"**{optimal_capacity:.0f} MWp** (achieved {unmet_pct_achieved:.2f}% unmet)")
        else:
            st.info(f"Validation mode: showing Solar LCOE at anchor capacity **{optimal_capacity:.0f} MWp** "
                    f"using the Excel's own anchor-interpolated annual energy.")

        k1, k2, k3 = st.columns(3)
        k1.metric("Solar Capacity", f"{optimal_capacity:.0f} MWp")
        k2.metric("Annual Energy (Year 1)", f"{annual_energy_kwh/1e6:.1f} GWh")
        k3.metric("Solar LCOE", f"${s_lcoe.lcoe_per_kwh*1000:.2f}/MWh")

        if use_profile:
            excel_energy_kwh = interp_energy(optimal_capacity, solar_cfg.anchors_mw, SOLAR_ENERGY_ANCHORS_KWH)
            diff_pct = (annual_energy_kwh - excel_energy_kwh) / excel_energy_kwh * 100 if excel_energy_kwh else 0
            st.caption(f"📊 For reference: Excel's anchor-interpolated energy at {optimal_capacity:.0f} MWp "
                       f"would be {excel_energy_kwh/1e6:.1f} GWh ({diff_pct:+.1f}% vs. profile-based). "
                       f"A large gap usually means the uploaded PV profile's yield differs from the Excel's "
                       f"generation assumption — worth reconciling before trusting profile-based sizing.")
        else:
            st.caption("✅ This LCOE should match Solar_3.xlsm's LCOE column at this exact anchor capacity "
                       "to within rounding — use this mode to confirm the cost/LCOE math before switching "
                       "to profile-based sizing.")

        st.markdown("---")
        st.subheader("☀️ Solar Cost Breakdown")
        s_capex_detail = s_cost['capex_detail']
        c1, c2 = st.columns(2)
        with c1:
            for label, cat in s_capex_detail['categories'].items():
                st.write(f"**{label}**: ${cat['subtotal']:,.0f}")
        with c2:
            st.write(f"**Base EPC Cost**: ${s_capex_detail['base_epc_cost']:,.0f}")
            st.write(f"**Grand Total CAPEX**: ${s_capex_detail['grand_total_capex']:,.0f}")
            st.write(f"**Year-1 OPEX**: ${s_cost['om_annual']:,.0f}")
        if s_capex_detail['extrapolated']:
            st.warning("Solar capacity is outside the anchor range — CAPEX/OPEX values are extrapolated (flat beyond range).")

        # Cost breakdown chart
        cat_labels = list(s_capex_detail['categories'].keys())
        cat_values_m = [s_capex_detail['categories'][l]['subtotal'] / 1e6 for l in cat_labels]
        fig_cost = go.Figure(data=[go.Bar(
            x=cat_labels, y=cat_values_m, marker_color=SJ_PALETTE,
            text=[f'${v:.2f}M' for v in cat_values_m], textposition='outside')])
        style_chart(fig_cost, height=380)
        fig_cost.update_layout(title='Solar CAPEX by Category', xaxis_title='Category',
                                yaxis_title='Cost ($M)', showlegend=False)
        st.plotly_chart(fig_cost, use_container_width=True)

        if use_profile:
            st.markdown("---")
            st.subheader("📈 Unmet Load vs. Capacity (search scan)")
            scan_df = pd.DataFrame([{'Capacity (MWp)': r.capacity_mw, 'Unmet Load (%)': r.unmet_percent}
                                     for r in search['scan']])
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=scan_df['Capacity (MWp)'], y=scan_df['Unmet Load (%)'],
                                      mode='lines+markers', line=dict(color='#E63946', width=3),
                                      marker=dict(size=8, color='#E63946')))
            fig.add_hline(y=target_unmet_pct, line_dash='dash', line_color='#1976D2', line_width=2,
                          annotation_text=f"Target: {target_unmet_pct:.1f}%",
                          annotation_font=dict(size=13, color='#1976D2'))
            style_chart(fig, height=380)
            fig.update_layout(xaxis_title="Solar Capacity (MWp)", yaxis_title="Unmet Load (%)")
            st.plotly_chart(fig, use_container_width=True)

            st.markdown("---")
            st.subheader("🕐 Representative Day Dispatch")
            st.caption("A single day picked at the median daily PV output across the year — "
                       "not the best or worst case, a typical one.")
            pv_hourly = best_dispatch.hourly_generation_kwh
            n_days = len(pv_hourly) // 24
            daily_pv = np.array([pv_hourly[d*24:(d+1)*24].sum() for d in range(n_days)])
            median_day = int(np.argsort(daily_pv)[len(daily_pv) // 2])
            start, end = median_day * 24, median_day * 24 + 24
            hours_of_day = np.arange(24)

            fig_day = go.Figure()
            fig_day.add_trace(go.Scatter(x=hours_of_day, y=load_kwh[start:end] / 1000,
                                          name='Load', mode='lines', line=dict(color='#1976D2', width=3)))
            fig_day.add_trace(go.Scatter(x=hours_of_day, y=pv_hourly[start:end] / 1000,
                                          name='Solar PV', mode='lines', fill='tozeroy',
                                          line=dict(color='#FDB462', width=2)))
            style_chart(fig_day, height=380)
            fig_day.update_layout(title=f'Representative Day (Day {median_day + 1} of year)',
                                   xaxis_title='Hour of Day', yaxis_title='Power (MW)',
                                   legend=dict(font=dict(color='#1a1a1a')))
            st.plotly_chart(fig_day, use_container_width=True)

        st.markdown("---")
        st.subheader("📥 Export Results")

        excel_output = BytesIO()
        with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
            summary_rows = [
                ['Parameter', 'Value'],
                ['Sizing Mode', 'Profile-based (unmet load target)' if use_profile else 'Excel validation mode'],
                ['Target Unmet Load (%)', target_unmet_pct if use_profile else ''],
                ['', ''],
                ['Solar Discount Rate (%)', solar_discount_pct],
                ['Solar Inflation Rate (%)', solar_inflation_pct],
                ['Solar Degradation Rate (%/yr)', solar_degradation_pct],
                ['Project Lifetime (years)', project_lifetime],
                ['', ''],
                ['Optimal Solar Capacity (MWp)', optimal_capacity],
                ['Annual Energy Year 1 (kWh)', annual_energy_kwh],
                ['Achieved Unmet Load (%)', unmet_pct_achieved if use_profile else ''],
                ['Solar LCOE ($/kWh)', s_lcoe.lcoe_per_kwh],
                ['Solar LCOE ($/MWh)', s_lcoe.lcoe_per_mwh],
                ['Grand Total CAPEX ($)', s_capex_detail['grand_total_capex']],
                ['Year-1 OPEX ($)', s_cost['om_annual']],
            ]
            pd.DataFrame(summary_rows).to_excel(writer, sheet_name='Summary', index=False, header=False)

            cost_rows = [{'Category': l, 'Subtotal ($)': c['subtotal']}
                         for l, c in s_capex_detail['categories'].items()]
            cost_rows += [
                {'Category': 'Base EPC Cost', 'Subtotal ($)': s_capex_detail['base_epc_cost']},
                {'Category': 'Grand Total CAPEX', 'Subtotal ($)': s_capex_detail['grand_total_capex']},
                {'Category': 'Year-1 OPEX', 'Subtotal ($)': s_cost['om_annual']},
            ]
            pd.DataFrame(cost_rows).to_excel(writer, sheet_name='Cost_Breakdown', index=False)

            if use_profile:
                combos_df = pd.DataFrame([{
                    'Capacity_MWp': r.capacity_mw, 'Unmet_Load_pct': r.unmet_percent,
                    'Annual_Generation_GWh': r.total_generation_kwh / 1e6,
                    'Meets_Target': r.unmet_percent <= target_unmet_pct,
                } for r in search['scan']])
                combos_df.to_excel(writer, sheet_name='All_Combinations', index=False)

                hourly_df = pd.DataFrame({
                    'Hour': np.arange(len(load_kwh)),
                    'Load_kW': load_kwh,
                    'Solar_PV_kW': best_dispatch.hourly_generation_kwh,
                    'Unmet_kW': best_dispatch.hourly_unmet_kwh,
                })
                hourly_df.to_excel(writer, sheet_name='Year_1_Hourly', index=False)

        excel_output.seek(0)
        st.download_button(
            label="📥 Download Full Results (Excel)",
            data=excel_output,
            file_name=f"energy_optimizer_pro_solar_{optimal_capacity:.0f}MWp.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        if use_profile:
            st.caption("Includes every capacity the grid search tried (**All_Combinations**) and the "
                       "full 8760-hour Year-1 dispatch for the winning capacity (**Year_1_Hourly**) — "
                       "same structure as the main EMO tool's export.")
    else:
        st.info("Configure Solar cost categories in the tab above, upload Load and PV profiles "
                 "(or switch to validation mode), then click **Run Sizing Search**.")

st.markdown("---")
st.markdown("<p style='text-align:center;color:gray'>Developed by SJ | 2026 | energy-optimizer-pro v1.0</p>",
            unsafe_allow_html=True)
