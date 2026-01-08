import pandas as pd
import streamlit as st


def _clamp_pct(value):
    if value is None or pd.isna(value):
        return 0.0
    return max(0.0, min(float(value), 1.0))


def _non_negative(value):
    if value is None or pd.isna(value):
        return 0.0
    return max(0.0, float(value))


def format_currency(value):
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, str):
        return value
    return f"{value / 1_000_000:,.2f} m EUR"


def format_pct(value):
    if value is None or pd.isna(value) or value == "":
        return ""
    return f"{value:.1%}"


def build_revenue_model_outputs(assumptions_state, scenario):
    revenue_state = assumptions_state["revenue_model"]
    revenue_final_by_year = []
    components_by_year = []
    reference_value = _non_negative(
        revenue_state["reference_revenue_eur"].get(scenario, 0.0)
    )

    for year_index in range(5):
        fte = revenue_state["consulting_fte"][scenario][year_index]
        workdays = revenue_state["workdays_per_year"][scenario][year_index]
        utilization = revenue_state["utilization_rate"][scenario][year_index]
        base_rate = revenue_state["day_rate_eur"][scenario][year_index]
        rate_growth = revenue_state["day_rate_growth_pct"][scenario][year_index]
        revenue_growth = revenue_state["revenue_growth_pct"][scenario][year_index]

        day_rate = base_rate * ((1 + rate_growth) ** year_index)
        capacity_revenue = fte * workdays * utilization * day_rate
        growth_adjusted = capacity_revenue * (1 + revenue_growth)

        guarantee_pct = revenue_state["guarantee_pct_by_year"][scenario][year_index]
        guaranteed_floor = reference_value * guarantee_pct

        final_total = max(guaranteed_floor, growth_adjusted)
        share_guaranteed = (
            guaranteed_floor / final_total if final_total else 0.0
        )

        revenue_final_by_year.append(final_total)
        components_by_year.append(
            {
                "consulting_fte": fte,
                "capacity_revenue": capacity_revenue,
                "growth_adjusted_revenue": growth_adjusted,
                "guaranteed_floor": guaranteed_floor,
                "final_total": final_total,
                "share_guaranteed": share_guaranteed,
            }
        )

    return revenue_final_by_year, components_by_year


def render_revenue_model_assumptions(input_model):
    st.header("Revenue Model")
    st.write("Detailed revenue planning (5-year view).")

    assumptions_state = st.session_state["assumptions"]
    revenue_state = assumptions_state["revenue_model"]
    scenario = st.session_state.get("assumptions.scenario", "Base")
    year_columns = [f"Year {i}" for i in range(5)]

    st.markdown("### Revenue Drivers")
    driver_rows = [
        ("Consulting FTE", "FTE"),
        ("Workdays per Year", "Days"),
        ("Utilization %", "%"),
        ("Day Rate (EUR)", "EUR"),
        ("Day Rate Growth (% p.a.)", "%"),
        ("Revenue Growth (% p.a.)", "%"),
    ]
    driver_values = {
        "Consulting FTE": revenue_state["consulting_fte"][scenario],
        "Workdays per Year": revenue_state["workdays_per_year"][scenario],
        "Utilization %": revenue_state["utilization_rate"][scenario],
        "Day Rate (EUR)": revenue_state["day_rate_eur"][scenario],
        "Day Rate Growth (% p.a.)": revenue_state["day_rate_growth_pct"][scenario],
        "Revenue Growth (% p.a.)": revenue_state["revenue_growth_pct"][scenario],
    }
    driver_table = {
        "Parameter": [row[0] for row in driver_rows],
        "Unit": [row[1] for row in driver_rows],
    }
    for idx, col in enumerate(year_columns):
        driver_table[col] = [
            driver_values[param][idx] for param, _ in driver_rows
        ]
    driver_df = pd.DataFrame(driver_table)
    driver_edit = st.data_editor(
        driver_df,
        hide_index=True,
        key="revenue_model.drivers",
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
        },
        use_container_width=True,
    )
    for year_index in range(5):
        revenue_state["consulting_fte"][scenario][year_index] = _non_negative(
            driver_edit.loc[0, year_columns[year_index]]
        )
        revenue_state["workdays_per_year"][scenario][year_index] = _non_negative(
            driver_edit.loc[1, year_columns[year_index]]
        )
        revenue_state["utilization_rate"][scenario][year_index] = _clamp_pct(
            driver_edit.loc[2, year_columns[year_index]]
        )
        revenue_state["day_rate_eur"][scenario][year_index] = _non_negative(
            driver_edit.loc[3, year_columns[year_index]]
        )
        revenue_state["day_rate_growth_pct"][scenario][year_index] = _clamp_pct(
            driver_edit.loc[4, year_columns[year_index]]
        )
        revenue_state["revenue_growth_pct"][scenario][year_index] = _clamp_pct(
            driver_edit.loc[5, year_columns[year_index]]
        )

    st.markdown("### Group Revenue Guarantee (Floor)")
    reference_df = pd.DataFrame(
        {
            "Parameter": ["Reference Revenue"],
            "Unit": ["EUR"],
            **{
                col: [revenue_state["reference_revenue_eur"][scenario]]
                for col in year_columns
            },
        }
    )
    reference_edit = st.data_editor(
        reference_df,
        hide_index=True,
        key="revenue_model.reference",
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
        },
        use_container_width=True,
    )
    revenue_state["reference_revenue_eur"][scenario] = _non_negative(
        reference_edit.loc[0, "Year 0"]
    )

    guarantee_df = pd.DataFrame(
        {
            "Parameter": ["Guarantee %"],
            "Unit": ["%"],
            **{
                col: [revenue_state["guarantee_pct_by_year"][scenario][idx]]
                for idx, col in enumerate(year_columns)
            },
        }
    )
    guarantee_edit = st.data_editor(
        guarantee_df,
        hide_index=True,
        key="revenue_model.guarantees",
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
        },
        use_container_width=True,
    )
    for year_index in range(5):
        revenue_state["guarantee_pct_by_year"][scenario][year_index] = _clamp_pct(
            guarantee_edit.loc[0, year_columns[year_index]]
        )

    st.markdown("### Revenue Bridge / Summary")
    bridge_rows = []
    for year_index in range(5):
        reference_value = revenue_state["reference_revenue_eur"][scenario]
        fte = revenue_state["consulting_fte"][scenario][year_index]
        workdays = revenue_state["workdays_per_year"][scenario][year_index]
        utilization = revenue_state["utilization_rate"][scenario][year_index]
        base_rate = revenue_state["day_rate_eur"][scenario][year_index]
        rate_growth = revenue_state["day_rate_growth_pct"][scenario][year_index]
        revenue_growth = revenue_state["revenue_growth_pct"][scenario][year_index]
        day_rate = base_rate * ((1 + rate_growth) ** year_index)
        capacity_revenue = fte * workdays * utilization * day_rate
        growth_adjusted = capacity_revenue * (1 + revenue_growth)
        guarantee_pct = revenue_state["guarantee_pct_by_year"][scenario][year_index]
        guaranteed_floor = reference_value * guarantee_pct
        final_revenue = max(guaranteed_floor, growth_adjusted)
        guaranteed_share = guaranteed_floor / final_revenue if final_revenue else 0.0
        bridge_rows.append(
            {
                "Year": f"Year {year_index}",
                "Capacity Revenue": capacity_revenue,
                "Growth Adj.": growth_adjusted,
                "Guaranteed Floor": guaranteed_floor,
                "Final Revenue": final_revenue,
                "Guaranteed %": guaranteed_share,
            }
        )
    bridge_df = pd.DataFrame(bridge_rows)
    bridge_units = {
        "Capacity Revenue": "EUR",
        "Growth Adj.": "EUR",
        "Guaranteed Floor": "EUR",
        "Final Revenue": "EUR",
        "Guaranteed %": "%",
    }
    bridge_df.insert(1, "Unit", ["EUR"] * len(bridge_df))
    for col in [
        "Capacity Revenue",
        "Growth Adj.",
        "Guaranteed Floor",
        "Final Revenue",
    ]:
        bridge_df[col] = bridge_df[col].apply(format_currency)
    bridge_df["Guaranteed %"] = bridge_df["Guaranteed %"].apply(format_pct)
    st.data_editor(
        bridge_df,
        hide_index=True,
        key="revenue_model.bridge",
        disabled=True,
        use_container_width=True,
    )
