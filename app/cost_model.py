import pandas as pd
import streamlit as st


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


def build_cost_model_outputs(assumptions_state, revenue_final_by_year):
    cost_state = assumptions_state["cost_model"]
    apply_inflation = bool(cost_state["inflation"].get("apply", False))
    inflation_rate = cost_state["inflation"].get("rate_pct", 0.0)
    cost_totals_by_year = []
    for year_index in range(5):
        personnel_row = cost_state["personnel"][year_index]
        fixed_row = cost_state["fixed_overhead"][year_index]
        variable_row = cost_state["variable_costs"][year_index]
        inflation_factor = (1 + inflation_rate) ** year_index if apply_inflation else 1.0

        consultant_total = (
            _non_negative(personnel_row["Consultant FTE"])
            * _non_negative(personnel_row["Consultant Loaded Cost (EUR)"])
            * inflation_factor
        )
        backoffice_total = (
            _non_negative(personnel_row["Backoffice FTE"])
            * _non_negative(personnel_row["Backoffice Loaded Cost (EUR)"])
            * inflation_factor
        )
        management_total = _non_negative(
            personnel_row["Management Cost (EUR)"]
        ) * inflation_factor
        personnel_total = consultant_total + backoffice_total + management_total

        fixed_total = sum(
            _non_negative(fixed_row[col])
            for col in [
                "Advisory",
                "Legal",
                "IT & Software",
                "Office Rent",
                "Services",
                "Other Services",
            ]
        ) * inflation_factor

        revenue = revenue_final_by_year[year_index]
        variable_total = 0.0
        for prefix in ["Training", "Travel", "Communication"]:
            cost_type = variable_row[f"{prefix} Type"]
            value = _non_negative(variable_row[f"{prefix} Value"])
            if cost_type == "%":
                variable_total += revenue * value
            else:
                variable_total += value * inflation_factor

        overhead_total = fixed_total + variable_total
        cost_totals_by_year.append(
            {
                "consultant_costs": consultant_total,
                "backoffice_costs": backoffice_total,
                "management_costs": management_total,
                "personnel_costs": personnel_total,
                "overhead_and_variable_costs": overhead_total,
                "total_operating_costs": personnel_total + overhead_total,
            }
        )
    return cost_totals_by_year


def render_cost_model_assumptions(input_model):
    st.header("Cost Model")
    st.write("Detailed annual cost planning (5-year view).")

    assumptions_state = st.session_state["assumptions"]
    cost_state = assumptions_state["cost_model"]
    year_columns = [f"Year {i}" for i in range(5)]

    st.markdown("### Inflation")
    inflation_cols = st.columns([1, 1])
    cost_state["inflation"]["apply"] = inflation_cols[0].toggle(
        "Apply inflation to costs",
        value=bool(cost_state["inflation"].get("apply", False)),
    )
    cost_state["inflation"]["rate_pct"] = inflation_cols[1].number_input(
        "Inflation Rate (% p.a.)",
        min_value=0.0,
        max_value=0.2,
        step=0.005,
        value=float(cost_state["inflation"].get("rate_pct", 0.0)),
        format="%.3f",
    )

    st.markdown("### Consultant Costs")
    consultant_table = {
        "Parameter": ["Consultant FTE", "Consultant Loaded Cost (EUR)"],
    }
    for year_index, col in enumerate(year_columns):
        consultant_table[col] = [
            cost_state["personnel"][year_index]["Consultant FTE"],
            cost_state["personnel"][year_index]["Consultant Loaded Cost (EUR)"],
        ]
    consultant_df = pd.DataFrame(consultant_table)
    consultant_edit = st.data_editor(
        consultant_df,
        hide_index=True,
        key="cost_model.consultant_table",
        column_config={"Parameter": st.column_config.TextColumn(disabled=True)},
        use_container_width=True,
    )
    for year_index in range(5):
        cost_state["personnel"][year_index]["Consultant FTE"] = _non_negative(
            consultant_edit.loc[0, year_columns[year_index]]
        )
        cost_state["personnel"][year_index]["Consultant Loaded Cost (EUR)"] = _non_negative(
            consultant_edit.loc[1, year_columns[year_index]]
        )

    st.markdown("### Backoffice Costs")
    backoffice_table = {
        "Parameter": ["Backoffice FTE", "Backoffice Loaded Cost (EUR)"],
    }
    for year_index, col in enumerate(year_columns):
        backoffice_table[col] = [
            cost_state["personnel"][year_index]["Backoffice FTE"],
            cost_state["personnel"][year_index]["Backoffice Loaded Cost (EUR)"],
        ]
    backoffice_df = pd.DataFrame(backoffice_table)
    backoffice_edit = st.data_editor(
        backoffice_df,
        hide_index=True,
        key="cost_model.backoffice_table",
        column_config={"Parameter": st.column_config.TextColumn(disabled=True)},
        use_container_width=True,
    )
    for year_index in range(5):
        cost_state["personnel"][year_index]["Backoffice FTE"] = _non_negative(
            backoffice_edit.loc[0, year_columns[year_index]]
        )
        cost_state["personnel"][year_index]["Backoffice Loaded Cost (EUR)"] = _non_negative(
            backoffice_edit.loc[1, year_columns[year_index]]
        )

    st.markdown("### Management")
    management_table = {
        "Parameter": ["Management Cost (EUR)"],
    }
    for year_index, col in enumerate(year_columns):
        management_table[col] = [
            cost_state["personnel"][year_index]["Management Cost (EUR)"]
        ]
    management_df = pd.DataFrame(management_table)
    management_edit = st.data_editor(
        management_df,
        hide_index=True,
        key="cost_model.management_table",
        column_config={"Parameter": st.column_config.TextColumn(disabled=True)},
        use_container_width=True,
    )
    for year_index in range(5):
        cost_state["personnel"][year_index]["Management Cost (EUR)"] = _non_negative(
            management_edit.loc[0, year_columns[year_index]]
        )

    st.markdown("### Fixed Overhead")
    fixed_table = {
        "Parameter": [
            "Advisory",
            "Legal",
            "IT & Software",
            "Office Rent",
            "Services",
            "Other Services",
        ],
    }
    for year_index, col in enumerate(year_columns):
        fixed_table[col] = [
            cost_state["fixed_overhead"][year_index]["Advisory"],
            cost_state["fixed_overhead"][year_index]["Legal"],
            cost_state["fixed_overhead"][year_index]["IT & Software"],
            cost_state["fixed_overhead"][year_index]["Office Rent"],
            cost_state["fixed_overhead"][year_index]["Services"],
            cost_state["fixed_overhead"][year_index]["Other Services"],
        ]
    fixed_df = pd.DataFrame(fixed_table)
    fixed_edit = st.data_editor(
        fixed_df,
        hide_index=True,
        key="cost_model.fixed_table",
        column_config={"Parameter": st.column_config.TextColumn(disabled=True)},
        use_container_width=True,
    )
    for year_index in range(5):
        cost_state["fixed_overhead"][year_index]["Advisory"] = _non_negative(
            fixed_edit.loc[0, year_columns[year_index]]
        )
        cost_state["fixed_overhead"][year_index]["Legal"] = _non_negative(
            fixed_edit.loc[1, year_columns[year_index]]
        )
        cost_state["fixed_overhead"][year_index]["IT & Software"] = _non_negative(
            fixed_edit.loc[2, year_columns[year_index]]
        )
        cost_state["fixed_overhead"][year_index]["Office Rent"] = _non_negative(
            fixed_edit.loc[3, year_columns[year_index]]
        )
        cost_state["fixed_overhead"][year_index]["Services"] = _non_negative(
            fixed_edit.loc[4, year_columns[year_index]]
        )
        cost_state["fixed_overhead"][year_index]["Other Services"] = _non_negative(
            fixed_edit.loc[5, year_columns[year_index]]
        )

    st.markdown("### Variable Costs")
    type_table = {"Parameter": ["Training", "Travel", "Communication"]}
    value_table = {"Parameter": ["Training", "Travel", "Communication"]}
    for year_index, col in enumerate(year_columns):
        type_table[col] = [
            cost_state["variable_costs"][year_index]["Training Type"],
            cost_state["variable_costs"][year_index]["Travel Type"],
            cost_state["variable_costs"][year_index]["Communication Type"],
        ]
        value_table[col] = [
            cost_state["variable_costs"][year_index]["Training Value"],
            cost_state["variable_costs"][year_index]["Travel Value"],
            cost_state["variable_costs"][year_index]["Communication Value"],
        ]
    type_df = pd.DataFrame(type_table)
    value_df = pd.DataFrame(value_table)
    type_edit = st.data_editor(
        type_df,
        hide_index=True,
        key="cost_model.variable_type",
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Year 0": st.column_config.SelectboxColumn(options=["EUR", "%"]),
            "Year 1": st.column_config.SelectboxColumn(options=["EUR", "%"]),
            "Year 2": st.column_config.SelectboxColumn(options=["EUR", "%"]),
            "Year 3": st.column_config.SelectboxColumn(options=["EUR", "%"]),
            "Year 4": st.column_config.SelectboxColumn(options=["EUR", "%"]),
        },
        use_container_width=True,
    )
    value_edit = st.data_editor(
        value_df,
        hide_index=True,
        key="cost_model.variable_value",
        column_config={"Parameter": st.column_config.TextColumn(disabled=True)},
        use_container_width=True,
    )
    for year_index in range(5):
        cost_state["variable_costs"][year_index]["Training Type"] = type_edit.loc[
            0, year_columns[year_index]
        ]
        cost_state["variable_costs"][year_index]["Travel Type"] = type_edit.loc[
            1, year_columns[year_index]
        ]
        cost_state["variable_costs"][year_index]["Communication Type"] = type_edit.loc[
            2, year_columns[year_index]
        ]
        cost_state["variable_costs"][year_index]["Training Value"] = _non_negative(
            value_edit.loc[0, year_columns[year_index]]
        )
        cost_state["variable_costs"][year_index]["Travel Value"] = _non_negative(
            value_edit.loc[1, year_columns[year_index]]
        )
        cost_state["variable_costs"][year_index]["Communication Value"] = _non_negative(
            value_edit.loc[2, year_columns[year_index]]
        )

    st.markdown("### Cost Summary")
    scenario = st.session_state.get("assumptions.scenario", "Base")
    from revenue_model import build_revenue_model_outputs

    revenue_final_by_year, _ = build_revenue_model_outputs(
        assumptions_state, scenario
    )
    cost_totals = build_cost_model_outputs(
        assumptions_state, revenue_final_by_year
    )
    summary_rows = []
    for year_index in range(5):
        summary_rows.append(
            {
                "Year": f"Year {year_index}",
                "Consultant": cost_totals[year_index]["consultant_costs"],
                "Backoffice": cost_totals[year_index]["backoffice_costs"],
                "Management": cost_totals[year_index]["management_costs"],
                "Total Personnel": cost_totals[year_index]["personnel_costs"],
                "Fixed OH": (
                    cost_totals[year_index]["overhead_and_variable_costs"]
                    - sum(
                        _non_negative(cost_state["variable_costs"][year_index][f"{prefix} Value"])
                        if cost_state["variable_costs"][year_index][f"{prefix} Type"] == "EUR"
                        else revenue_final_by_year[year_index]
                        * _non_negative(cost_state["variable_costs"][year_index][f"{prefix} Value"])
                        for prefix in ["Training", "Travel", "Communication"]
                    )
                ),
                "Variable": (
                    cost_totals[year_index]["overhead_and_variable_costs"]
                    - sum(
                        _non_negative(cost_state["fixed_overhead"][year_index][col])
                        for col in [
                            "Advisory",
                            "Legal",
                            "IT & Software",
                            "Office Rent",
                            "Services",
                            "Other Services",
                        ]
                    )
                ),
                "Total Operating Costs": cost_totals[year_index]["total_operating_costs"],
            }
        )
    summary_df = pd.DataFrame(summary_rows)
    for col in [
        "Consultant",
        "Backoffice",
        "Management",
        "Total Personnel",
        "Fixed OH",
        "Variable",
        "Total Operating Costs",
    ]:
        summary_df[col] = summary_df[col].apply(format_currency)
    st.dataframe(summary_df, use_container_width=True)
