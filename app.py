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
import plotly.graph_objects as go

from cost_engine import CostItem, CostCategory, DetailedCostConfig, compute_detailed_capex_opex
from lcoe_engine import solar_lcoe, wind_lcoe, hybrid_lcoe
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
    st.subheader("🔎 Sizing Search")
    st.caption("No hourly dispatch — sizes on annual energy vs. target, matching the Excel workbook's own scope.")
    target_energy_gwh = st.number_input("Target Annual Energy (GWh/yr)", value=150.0, min_value=1.0, step=10.0)
    col1, col2 = st.columns(2)
    with col1:
        solar_min = st.number_input("Solar Min (MWp)", value=50.0, step=5.0)
        solar_max = st.number_input("Solar Max (MWp)", value=130.0, step=5.0)
        solar_step = st.number_input("Solar Step (MWp)", value=5.0, step=1.0)
    with col2:
        wind_min = st.number_input("Wind Min (MW)", value=100.0, step=10.0)
        wind_max = st.number_input("Wind Max (MW)", value=500.0, step=10.0)
        wind_step = st.number_input("Wind Step (MW)", value=25.0, step=5.0)


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
        wind_cfg = st.session_state.wind_cfg

        solar_caps = [solar_min + i*solar_step for i in range(int((solar_max-solar_min)//solar_step)+1)]
        wind_caps = [wind_min + i*wind_step for i in range(int((wind_max-wind_min)//wind_step)+1)]

        candidates = []
        for sc in solar_caps:
            solar_energy_kwh = interp_energy(sc, solar_cfg.anchors_mw, SOLAR_ENERGY_ANCHORS_KWH)
            for wc in wind_caps:
                wind_energy_kwh = interp_energy(wc, wind_cfg.anchors_mw, WIND_ENERGY_ANCHORS_KWH)
                total_gwh = (solar_energy_kwh + wind_energy_kwh) / 1e6
                if total_gwh >= target_energy_gwh:
                    candidates.append((sc, wc, solar_energy_kwh, wind_energy_kwh, total_gwh))

        if not candidates:
            st.error("No capacity combination in the search range meets the target energy. Widen the ranges.")
        else:
            results = []
            for sc, wc, se_kwh, we_kwh, total_gwh in candidates:
                s_cost = compute_detailed_capex_opex(solar_cfg, sc)
                w_cost = compute_detailed_capex_opex(wind_cfg, wc)

                s_lcoe = solar_lcoe(s_cost['capital'], s_cost['om_annual'], se_kwh,
                                     solar_discount_pct/100, solar_inflation_pct/100,
                                     solar_degradation_pct/100, project_lifetime)
                w_lcoe = wind_lcoe(w_cost['capital'], w_cost['om_annual'], we_kwh,
                                    wind_discount_pct/100, wind_inflation_pct/100,
                                    wind_degradation_pct/100, project_lifetime)
                h_lcoe = hybrid_lcoe(s_lcoe, w_lcoe)

                results.append({
                    'solar_mwp': sc, 'wind_mw': wc, 'total_gwh': total_gwh,
                    'solar_capex': s_cost['capital'], 'wind_capex': w_cost['capital'],
                    'total_npv_cost': h_lcoe.npv_capex + h_lcoe.npv_opex,
                    'solar_lcoe': s_lcoe.lcoe_per_kwh, 'wind_lcoe': w_lcoe.lcoe_per_kwh,
                    'hybrid_lcoe': h_lcoe.lcoe_per_kwh,
                    's_lcoe_obj': s_lcoe, 'w_lcoe_obj': w_lcoe, 'h_lcoe_obj': h_lcoe,
                    's_cost_detail': s_cost, 'w_cost_detail': w_cost,
                })

            results_df = pd.DataFrame(results)
            best = results_df.loc[results_df['total_npv_cost'].idxmin()]

            st.success(f"Optimal sizing (min NPV cost meeting {target_energy_gwh:.0f} GWh/yr target): "
                       f"**Solar {best['solar_mwp']:.0f} MWp + Wind {best['wind_mw']:.0f} MW**")

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Solar LCOE", f"${best['solar_lcoe']*1000:.2f}/MWh")
            k2.metric("Wind LCOE", f"${best['wind_lcoe']*1000:.2f}/MWh")
            k3.metric("Hybrid LCOE", f"${best['hybrid_lcoe']*1000:.2f}/MWh")
            k4.metric("NPV of Cost", f"${best['total_npv_cost']/1e6:.2f}M")

            st.caption("✅ Solar LCOE and Wind LCOE use the same growing-annuity NPV formula as "
                       "Solar_3.xlsm — verify by entering this tool's Solar/Wind capacity at one of "
                       "the workbook's own anchor points and comparing the LCOE column directly.")

            st.markdown("---")
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("☀️ Solar Cost Breakdown")
                for label, cat in best['s_cost_detail']['categories'].items():
                    st.write(f"**{label}**: ${cat['subtotal']:,.0f}")
                st.write(f"**Base EPC Cost**: ${best['s_cost_detail']['base_epc_cost']:,.0f}")
                st.write(f"**Grand Total CAPEX**: ${best['s_cost_detail']['grand_total_capex']:,.0f}")
                if best['s_cost_detail']['extrapolated']:
                    st.warning("Solar capacity is outside the anchor range — values are extrapolated (flat beyond range).")
            with c2:
                st.subheader("💨 Wind Cost Breakdown")
                for label, cat in best['w_cost_detail']['categories'].items():
                    st.write(f"**{label}**: ${cat['subtotal']:,.0f}")
                st.write(f"**Base EPC Cost**: ${best['w_cost_detail']['base_epc_cost']:,.0f}")
                st.write(f"**Grand Total CAPEX**: ${best['w_cost_detail']['grand_total_capex']:,.0f}")
                if best['w_cost_detail']['extrapolated']:
                    st.warning("Wind capacity is outside the anchor range — values are extrapolated (flat beyond range).")

            st.markdown("---")
            st.subheader("📈 LCOE vs. Capacity Mix (all candidates meeting target)")
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=results_df['solar_mwp'], y=results_df['hybrid_lcoe']*1000,
                                      mode='markers', marker=dict(size=8, color=results_df['wind_mw'],
                                      colorscale='Viridis', showscale=True, colorbar=dict(title="Wind MW")),
                                      text=results_df['wind_mw'], name='Hybrid LCOE'))
            fig.update_layout(xaxis_title="Solar Capacity (MWp)", yaxis_title="Hybrid LCOE ($/MWh)",
                               plot_bgcolor='white', paper_bgcolor='white', font=dict(color='#333333'), height=420)
            st.plotly_chart(fig, use_container_width=True)

            st.markdown("---")
            st.subheader("All Candidates")
            display_df = results_df[['solar_mwp', 'wind_mw', 'total_gwh', 'solar_lcoe', 'wind_lcoe', 'hybrid_lcoe', 'total_npv_cost']].copy()
            display_df.columns = ['Solar (MWp)', 'Wind (MW)', 'Energy (GWh/yr)', 'Solar LCOE ($/kWh)',
                                   'Wind LCOE ($/kWh)', 'Hybrid LCOE ($/kWh)', 'NPV of Cost ($)']
            st.dataframe(display_df.sort_values('NPV of Cost ($)'), use_container_width=True)
    else:
        st.info("Configure Solar and Wind cost categories in the tabs above, then click **Run Sizing Search**.")

st.markdown("---")
st.markdown("<p style='text-align:center;color:gray'>Developed by SJ | 2026 | energy-optimizer-pro v1.0</p>",
            unsafe_allow_html=True)
