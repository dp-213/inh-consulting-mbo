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
    reference_value = _non_negative(revenue_state["reference"][0].get(scenario, 0.0))
    revenue_final_by_year = []
    components_by_year = []
    for year_index in range(5):
        guarantee_pct = _clamp_pct(
            revenue_state["guarantees"][year_index].get(scenario, 0.0)
        )
        in_group = _non_negative(
            revenue_state["in_group"][year_index].get(scenario, 0.0)
        )
        external = _non_negative(
            revenue_state["external"][year_index].get(scenario, 0.0)
        )
        guaranteed_floor = reference_value * guarantee_pct
        modeled_total = in_group + external
        final_total = max(guaranteed_floor, modeled_total)
        share_guaranteed = (
            guaranteed_floor / final_total if final_total else 0.0
        )
        revenue_final_by_year.append(final_total)
        components_by_year.append(
            {
                "guaranteed_floor": guaranteed_floor,
                "modeled_in_group": in_group,
                "modeled_external": external,
                "modeled_total": modeled_total,
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

    st.markdown("### Inputs")
    reference_value = _non_negative(revenue_state["reference"][0].get(scenario, 0.0))
    reference_row = pd.DataFrame(
        [{"Parameter": "Reference Revenue (EUR)", **{c: reference_value for c in year_columns}}]
    )
    reference_edit = st.data_editor(
        reference_row,
        hide_index=True,
        key="revenue_model.reference",
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Year 1": st.column_config.NumberColumn(disabled=True),
            "Year 2": st.column_config.NumberColumn(disabled=True),
            "Year 3": st.column_config.NumberColumn(disabled=True),
            "Year 4": st.column_config.NumberColumn(disabled=True),
        },
        use_container_width=True,
    )
    revenue_state["reference"][0][scenario] = _non_negative(
        reference_edit.loc[0, "Year 0"]
    )

    def _year_table(label, rows_key):
        data = {col: [] for col in year_columns}
        data["Parameter"] = [label]
        for year_index, col in enumerate(year_columns):
            data[col].append(
                revenue_state[rows_key][year_index].get(scenario, 0.0)
            )
        df = pd.DataFrame(data)
        edit = st.data_editor(
            df,
            hide_index=True,
            key=f"revenue_model.{rows_key}",
            column_config={
                "Parameter": st.column_config.TextColumn(disabled=True),
            },
            use_container_width=True,
        )
        for year_index in range(5):
            revenue_state[rows_key][year_index][scenario] = _non_negative(
                edit.loc[0, year_columns[year_index]]
            )

    _year_table("Guarantee %", "guarantees")
    _year_table("In-Group Revenue (EUR)", "in_group")
    _year_table("External Revenue (EUR)", "external")

    st.markdown("### Summary (Revenue Bridge)")
    reference_value = _non_negative(revenue_state["reference"][0].get(scenario, 0.0))
    bridge_rows = []
    for year_index in range(5):
        guarantee_pct = _clamp_pct(
            revenue_state["guarantees"][year_index].get(scenario, 0.0)
        )
        in_group = _non_negative(
            revenue_state["in_group"][year_index].get(scenario, 0.0)
        )
        external = _non_negative(
            revenue_state["external"][year_index].get(scenario, 0.0)
        )
        guaranteed_floor = reference_value * guarantee_pct
        modeled_total = in_group + external
        final_total = max(guaranteed_floor, modeled_total)
        share_guaranteed = (
            guaranteed_floor / final_total if final_total else 0.0
        )
        bridge_rows.append(
            {
                "Parameter": f"Year {year_index}",
                "Guaranteed Floor": guaranteed_floor,
                "In-Group": in_group,
                "External": external,
                "Modeled Total": modeled_total,
                "Final Revenue": final_total,
                "Guaranteed Share %": share_guaranteed,
            }
        )
    bridge_df = pd.DataFrame(bridge_rows)
    for col in [
        "Guaranteed Floor",
        "In-Group",
        "External",
        "Modeled Total",
        "Final Revenue",
    ]:
        bridge_df[col] = bridge_df[col].apply(format_currency)
    bridge_df["Guaranteed Share %"] = bridge_df["Guaranteed Share %"].apply(format_pct)
    st.dataframe(bridge_df, use_container_width=True)
