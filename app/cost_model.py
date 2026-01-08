import pandas as pd
import streamlit as st


def _non_negative(value):
    if value is None or pd.isna(value):
        return 0.0
    return max(0.0, float(value))


def _percent_to_display(value):
    if value is None or pd.isna(value):
        return value
    return float(value) * 100


def _percent_from_display(value):
    if value is None or pd.isna(value):
        return value
    return float(value) / 100


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
    inflation_df = pd.DataFrame(
        {
            "Parameter": ["Apply Inflation", "Inflation Rate (% p.a.)"],
            "Unit": ["", "%"],
            "Value": [
                bool(cost_state["inflation"].get("apply", False)),
                _percent_to_display(cost_state["inflation"].get("rate_pct", 0.0)),
            ],
        }
    )
    inflation_edit = st.data_editor(
        inflation_df,
        hide_index=True,
        key="cost_model.inflation",
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            "Value": st.column_config.NumberColumn(format=",.2f"),
        },
        use_container_width=True,
    )
    cost_state["inflation"]["apply"] = bool(inflation_edit.loc[0, "Value"])
    cost_state["inflation"]["rate_pct"] = _percent_from_display(
        _non_negative(inflation_edit.loc[1, "Value"])
    )

    st.markdown("### Consultant Costs")
    consultant_table = {
        "Parameter": ["Consultant FTE", "Consultant Loaded Cost"],
        "Unit": ["FTE", "EUR"],
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
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            **{col: st.column_config.NumberColumn(format=",.2f") for col in year_columns},
        },
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
        "Parameter": ["Backoffice FTE", "Backoffice Loaded Cost"],
        "Unit": ["FTE", "EUR"],
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
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            **{col: st.column_config.NumberColumn(format=",.2f") for col in year_columns},
        },
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
        "Parameter": ["Management Cost"],
        "Unit": ["EUR"],
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
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            **{col: st.column_config.NumberColumn(format=",.2f") for col in year_columns},
        },
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
        "Unit": ["EUR"] * 6,
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
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            **{col: st.column_config.NumberColumn(format=",.2f") for col in year_columns},
        },
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
    type_table = {"Parameter": ["Training", "Travel", "Communication"], "Unit": ["Type"] * 3}
    value_table = {"Parameter": ["Training", "Travel", "Communication"], "Unit": ["EUR / %"] * 3}
    for year_index, col in enumerate(year_columns):
        type_table[col] = [
            cost_state["variable_costs"][year_index]["Training Type"],
            cost_state["variable_costs"][year_index]["Travel Type"],
            cost_state["variable_costs"][year_index]["Communication Type"],
        ]
        value_table[col] = [
            _percent_to_display(
                cost_state["variable_costs"][year_index]["Training Value"]
            )
            if cost_state["variable_costs"][year_index]["Training Type"] == "%"
            else cost_state["variable_costs"][year_index]["Training Value"],
            _percent_to_display(
                cost_state["variable_costs"][year_index]["Travel Value"]
            )
            if cost_state["variable_costs"][year_index]["Travel Type"] == "%"
            else cost_state["variable_costs"][year_index]["Travel Value"],
            _percent_to_display(
                cost_state["variable_costs"][year_index]["Communication Value"]
            )
            if cost_state["variable_costs"][year_index]["Communication Type"] == "%"
            else cost_state["variable_costs"][year_index]["Communication Value"],
        ]
    type_df = pd.DataFrame(type_table)
    value_df = pd.DataFrame(value_table)
    type_edit = st.data_editor(
        type_df,
        hide_index=True,
        key="cost_model.variable_type",
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
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
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            **{col: st.column_config.NumberColumn(format=",.2f") for col in year_columns},
        },
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
        training_value = _non_negative(
            value_edit.loc[0, year_columns[year_index]]
        )
        travel_value = _non_negative(
            value_edit.loc[1, year_columns[year_index]]
        )
        communication_value = _non_negative(
            value_edit.loc[2, year_columns[year_index]]
        )
        if cost_state["variable_costs"][year_index]["Training Type"] == "%":
            training_value = _percent_from_display(training_value)
        if cost_state["variable_costs"][year_index]["Travel Type"] == "%":
            travel_value = _percent_from_display(travel_value)
        if cost_state["variable_costs"][year_index]["Communication Type"] == "%":
            communication_value = _percent_from_display(communication_value)
        cost_state["variable_costs"][year_index]["Training Value"] = training_value
        cost_state["variable_costs"][year_index]["Travel Value"] = travel_value
        cost_state["variable_costs"][year_index]["Communication Value"] = communication_value

    st.markdown("### Cost Summary")
    scenario = st.session_state.get("assumptions.scenario", "Base")
    from revenue_model import build_revenue_model_outputs

    revenue_final_by_year, _ = build_revenue_model_outputs(
        assumptions_state, scenario
    )
    cost_totals = build_cost_model_outputs(
        assumptions_state, revenue_final_by_year
    )
    summary_rows = {
        "Consultant": [],
        "Backoffice": [],
        "Management": [],
        "Total Personnel": [],
        "Fixed OH": [],
        "Variable": [],
        "Total Operating Costs": [],
    }
    for year_index in range(5):
        fixed_total = sum(
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
        variable_total = sum(
            _non_negative(cost_state["variable_costs"][year_index][f"{prefix} Value"])
            if cost_state["variable_costs"][year_index][f"{prefix} Type"] == "EUR"
            else revenue_final_by_year[year_index]
            * _non_negative(cost_state["variable_costs"][year_index][f"{prefix} Value"])
            for prefix in ["Training", "Travel", "Communication"]
        )
        summary_rows["Consultant"].append(
            cost_totals[year_index]["consultant_costs"]
        )
        summary_rows["Backoffice"].append(
            cost_totals[year_index]["backoffice_costs"]
        )
        summary_rows["Management"].append(
            cost_totals[year_index]["management_costs"]
        )
        summary_rows["Total Personnel"].append(
            cost_totals[year_index]["personnel_costs"]
        )
        summary_rows["Fixed OH"].append(fixed_total)
        summary_rows["Variable"].append(variable_total)
        summary_rows["Total Operating Costs"].append(
            cost_totals[year_index]["total_operating_costs"]
        )

    summary_table = {
        "Parameter": list(summary_rows.keys()),
        "Unit": ["EUR"] * len(summary_rows),
    }
    for year_index, col in enumerate(year_columns):
        summary_table[col] = [
            summary_rows[row_key][year_index] for row_key in summary_rows
        ]
    summary_df = pd.DataFrame(summary_table)
    st.data_editor(
        summary_df,
        hide_index=True,
        key="cost_model.summary",
        disabled=True,
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            **{col: st.column_config.NumberColumn(format=",.2f") for col in year_columns},
        },
        use_container_width=True,
    )
