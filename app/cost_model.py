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


def _format_number_display(value, decimals=0):
    if value is None or pd.isna(value):
        return ""
    try:
        number = float(str(value).replace(",", ""))
    except ValueError:
        return value
    if decimals == 0:
        return f"{number:,.0f}"
    return f"{number:,.{decimals}f}"


def _parse_number_display(value):
    if value is None or pd.isna(value) or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return float(str(value).replace(",", ""))


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

    explain_cost = st.toggle("Explain cost logic & assumptions")
    if explain_cost:
        st.markdown(
            """
            <div style="background:#f3f4f6;padding:12px 14px;border-radius:6px;font-size:0.9rem;">
              <strong>A. Consultant costs</strong>
              <ul>
                <li>Consultant FTEs are planned in the Cost Model.</li>
                <li>Average base cost grows with inflation.</li>
                <li>Loaded cost includes employer-side burdens.</li>
                <li>No seniority mix is modeled (intentional simplification).</li>
              </ul>
              <strong>B. Backoffice costs</strong>
              <ul>
                <li>Planned independently of revenue.</li>
                <li>Fully fixed and inflation-driven.</li>
              </ul>
              <strong>C. Fixed overhead</strong>
              <ul>
                <li>Sticky cost base.</li>
                <li>Inflated annually.</li>
                <li>No assumed operating leverage.</li>
              </ul>
              <strong>D. Variable costs</strong>
              <ul>
                <li>Conservative modeling via % of revenue or absolute values.</li>
                <li>Zero in the base case is intentional.</li>
              </ul>
              <strong>E. Explicit exclusions</strong>
              <ul>
                <li>No cost synergies.</li>
                <li>No restructuring effects.</li>
                <li>No bonus pools tied to upside.</li>
              </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("### Inflation")
    inflation_df = pd.DataFrame(
        {
            "Parameter": ["Apply Inflation", "Inflation Rate (% p.a.)"],
            "Unit": ["", "%"],
            "Value": [
                bool(cost_state["inflation"].get("apply", False)),
                _format_number_display(
                    _percent_to_display(cost_state["inflation"].get("rate_pct", 0.0)),
                    1,
                ),
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
            "Value": st.column_config.TextColumn(),
        },
        use_container_width=True,
    )
    cost_state["inflation"]["apply"] = bool(inflation_edit.loc[0, "Value"])
    cost_state["inflation"]["rate_pct"] = _percent_from_display(
        _non_negative(_parse_number_display(inflation_edit.loc[1, "Value"]))
    )

    st.markdown("### Consultant Costs")
    consultant_table = {
        "Parameter": ["Consultant FTE", "Consultant Loaded Cost"],
        "Unit": ["FTE", "EUR"],
    }
    for year_index, col in enumerate(year_columns):
        consultant_table[col] = [
            _format_number_display(
                cost_state["personnel"][year_index]["Consultant FTE"], 0
            ),
            _format_number_display(
                cost_state["personnel"][year_index]["Consultant Loaded Cost (EUR)"],
                0,
            ),
        ]
    consultant_df = pd.DataFrame(consultant_table)
    consultant_edit = st.data_editor(
        consultant_df,
        hide_index=True,
        key="cost_model.consultant_table",
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            **{col: st.column_config.TextColumn() for col in year_columns},
        },
        use_container_width=True,
    )
    for year_index in range(5):
        cost_state["personnel"][year_index]["Consultant FTE"] = _non_negative(
            _parse_number_display(consultant_edit.loc[0, year_columns[year_index]])
        )
        cost_state["personnel"][year_index]["Consultant Loaded Cost (EUR)"] = _non_negative(
            _parse_number_display(consultant_edit.loc[1, year_columns[year_index]])
        )

    st.markdown("### Backoffice Costs")
    backoffice_table = {
        "Parameter": ["Backoffice FTE", "Backoffice Loaded Cost"],
        "Unit": ["FTE", "EUR"],
    }
    for year_index, col in enumerate(year_columns):
        backoffice_table[col] = [
            _format_number_display(
                cost_state["personnel"][year_index]["Backoffice FTE"], 0
            ),
            _format_number_display(
                cost_state["personnel"][year_index]["Backoffice Loaded Cost (EUR)"],
                0,
            ),
        ]
    backoffice_df = pd.DataFrame(backoffice_table)
    backoffice_edit = st.data_editor(
        backoffice_df,
        hide_index=True,
        key="cost_model.backoffice_table",
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            **{col: st.column_config.TextColumn() for col in year_columns},
        },
        use_container_width=True,
    )
    for year_index in range(5):
        cost_state["personnel"][year_index]["Backoffice FTE"] = _non_negative(
            _parse_number_display(backoffice_edit.loc[0, year_columns[year_index]])
        )
        cost_state["personnel"][year_index]["Backoffice Loaded Cost (EUR)"] = _non_negative(
            _parse_number_display(backoffice_edit.loc[1, year_columns[year_index]])
        )

    st.markdown("### Management")
    management_table = {
        "Parameter": ["Management Cost"],
        "Unit": ["EUR"],
    }
    for year_index, col in enumerate(year_columns):
        management_table[col] = [
            _format_number_display(
                cost_state["personnel"][year_index]["Management Cost (EUR)"], 0
            )
        ]
    management_df = pd.DataFrame(management_table)
    management_edit = st.data_editor(
        management_df,
        hide_index=True,
        key="cost_model.management_table",
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            **{col: st.column_config.TextColumn() for col in year_columns},
        },
        use_container_width=True,
    )
    for year_index in range(5):
        cost_state["personnel"][year_index]["Management Cost (EUR)"] = _non_negative(
            _parse_number_display(management_edit.loc[0, year_columns[year_index]])
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
            _format_number_display(
                cost_state["fixed_overhead"][year_index]["Advisory"], 0
            ),
            _format_number_display(
                cost_state["fixed_overhead"][year_index]["Legal"], 0
            ),
            _format_number_display(
                cost_state["fixed_overhead"][year_index]["IT & Software"], 0
            ),
            _format_number_display(
                cost_state["fixed_overhead"][year_index]["Office Rent"], 0
            ),
            _format_number_display(
                cost_state["fixed_overhead"][year_index]["Services"], 0
            ),
            _format_number_display(
                cost_state["fixed_overhead"][year_index]["Other Services"], 0
            ),
        ]
    fixed_df = pd.DataFrame(fixed_table)
    fixed_edit = st.data_editor(
        fixed_df,
        hide_index=True,
        key="cost_model.fixed_table",
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            **{col: st.column_config.TextColumn() for col in year_columns},
        },
        use_container_width=True,
    )
    for year_index in range(5):
        cost_state["fixed_overhead"][year_index]["Advisory"] = _non_negative(
            _parse_number_display(fixed_edit.loc[0, year_columns[year_index]])
        )
        cost_state["fixed_overhead"][year_index]["Legal"] = _non_negative(
            _parse_number_display(fixed_edit.loc[1, year_columns[year_index]])
        )
        cost_state["fixed_overhead"][year_index]["IT & Software"] = _non_negative(
            _parse_number_display(fixed_edit.loc[2, year_columns[year_index]])
        )
        cost_state["fixed_overhead"][year_index]["Office Rent"] = _non_negative(
            _parse_number_display(fixed_edit.loc[3, year_columns[year_index]])
        )
        cost_state["fixed_overhead"][year_index]["Services"] = _non_negative(
            _parse_number_display(fixed_edit.loc[4, year_columns[year_index]])
        )
        cost_state["fixed_overhead"][year_index]["Other Services"] = _non_negative(
            _parse_number_display(fixed_edit.loc[5, year_columns[year_index]])
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
            _format_number_display(
                _percent_to_display(
                    cost_state["variable_costs"][year_index]["Training Value"]
                ),
                1,
            )
            if cost_state["variable_costs"][year_index]["Training Type"] == "%"
            else _format_number_display(
                cost_state["variable_costs"][year_index]["Training Value"], 0
            ),
            _format_number_display(
                _percent_to_display(
                    cost_state["variable_costs"][year_index]["Travel Value"]
                ),
                1,
            )
            if cost_state["variable_costs"][year_index]["Travel Type"] == "%"
            else _format_number_display(
                cost_state["variable_costs"][year_index]["Travel Value"], 0
            ),
            _format_number_display(
                _percent_to_display(
                    cost_state["variable_costs"][year_index]["Communication Value"]
                ),
                1,
            )
            if cost_state["variable_costs"][year_index]["Communication Type"] == "%"
            else _format_number_display(
                cost_state["variable_costs"][year_index]["Communication Value"], 0
            ),
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
            **{col: st.column_config.TextColumn() for col in year_columns},
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
            _parse_number_display(value_edit.loc[0, year_columns[year_index]])
        )
        travel_value = _non_negative(
            _parse_number_display(value_edit.loc[1, year_columns[year_index]])
        )
        communication_value = _non_negative(
            _parse_number_display(value_edit.loc[2, year_columns[year_index]])
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
            _format_number_display(summary_rows[row_key][year_index], 0)
            for row_key in summary_rows
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
            **{col: st.column_config.TextColumn() for col in year_columns},
        },
        use_container_width=True,
    )
