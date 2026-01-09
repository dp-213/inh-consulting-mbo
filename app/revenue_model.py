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


def format_pct(value):
    if value is None or pd.isna(value) or value == "":
        return ""
    return f"{value:.1%}"


def build_revenue_model_outputs(assumptions_state, scenario):
    revenue_state = assumptions_state["revenue_model"]
    cost_state = assumptions_state["cost_model"]
    revenue_final_by_year = []
    components_by_year = []
    reference_value = _non_negative(
        revenue_state["reference_revenue_eur"].get(scenario, 0.0)
    )

    for year_index in range(5):
        fte = _non_negative(
            cost_state["personnel"][year_index]["Consultant FTE"]
        )
        workdays = revenue_state["workdays_per_year"][scenario][year_index]
        utilization = revenue_state["utilization_rate"][scenario][year_index]
        group_rate = revenue_state["group_day_rate_eur"][scenario][year_index]
        external_rate = revenue_state["external_day_rate_eur"][scenario][year_index]
        rate_growth = revenue_state["day_rate_growth_pct"][scenario][year_index]
        revenue_growth = revenue_state["revenue_growth_pct"][scenario][year_index]
        group_share = revenue_state["group_capacity_share_pct"][scenario][year_index]
        external_share = revenue_state["external_capacity_share_pct"][scenario][
            year_index
        ]
        total_share = group_share + external_share
        if total_share > 0:
            group_share = group_share / total_share
            external_share = external_share / total_share

        group_rate = group_rate * ((1 + rate_growth) ** year_index)
        external_rate = external_rate * ((1 + rate_growth) ** year_index)
        capacity_days = fte * workdays * utilization
        adjusted_capacity_days = capacity_days * (1 + revenue_growth)
        modeled_group_revenue = (
            adjusted_capacity_days * group_share * group_rate
        )
        modeled_external_revenue = (
            adjusted_capacity_days * external_share * external_rate
        )
        modeled_total_revenue = modeled_group_revenue + modeled_external_revenue

        guarantee_pct = revenue_state["guarantee_pct_by_year"][scenario][year_index]
        guaranteed_floor = reference_value * guarantee_pct
        guaranteed_group_revenue = max(
            modeled_group_revenue, guaranteed_floor
        )
        final_total = guaranteed_group_revenue + modeled_external_revenue
        share_guaranteed = (
            guaranteed_group_revenue / final_total if final_total else 0.0
        )

        revenue_final_by_year.append(final_total)
        components_by_year.append(
            {
                "consulting_fte": fte,
                "capacity_days": capacity_days,
                "adjusted_capacity_days": adjusted_capacity_days,
                "group_share_pct": group_share,
                "external_share_pct": external_share,
                "modeled_group_revenue": modeled_group_revenue,
                "modeled_external_revenue": modeled_external_revenue,
                "modeled_total_revenue": modeled_total_revenue,
                "guaranteed_floor": guaranteed_floor,
                "guaranteed_group_revenue": guaranteed_group_revenue,
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
    cost_state = assumptions_state["cost_model"]
    scenario = st.session_state.get("assumptions.scenario", "Base")
    year_columns = [f"Year {i}" for i in range(5)]

    explain_revenue = st.toggle("Explain revenue logic & assumptions")
    if explain_revenue:
        st.markdown(
            """
            <div style="background:#f3f4f6;padding:12px 14px;border-radius:6px;font-size:0.9rem;">
              <strong>A. Capacity logic</strong>
              <ul>
                <li>Consultant FTEs are derived from the Cost Model and are the single source of truth.</li>
                <li>Billable days and utilization define billable capacity in days.</li>
              </ul>
              <strong>B. Revenue split logic</strong>
              <ul>
                <li>Capacity is allocated between Group and External consulting each year.</li>
                <li>External revenue uses the same capacity pool and does not add capacity.</li>
              </ul>
              <strong>C. Pricing logic</strong>
              <ul>
                <li>Group day rates reflect conservative contractual pricing.</li>
                <li>External day rates represent optional upside with market risk.</li>
              </ul>
              <strong>D. Guarantee logic</strong>
              <ul>
                <li>Guarantee applies only to Group revenue.</li>
                <li>Guaranteed Group Revenue = max(Modeled Group Revenue, Reference × Guarantee %).</li>
                <li>Final Revenue = Guaranteed Group Revenue + External Revenue.</li>
              </ul>
              <strong>E. Not modeled</strong>
              <ul>
                <li>No explicit presales deduction (covered by utilization).</li>
                <li>No seniority mix (single blended rate).</li>
                <li>No multi-layer ramp-up beyond basic assumptions.</li>
              </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("### Consultant Capacity (Derived)")
    derived_fte_table = {
        "Parameter": ["Consultant FTE (Derived from Cost Model)"],
        "Unit": ["FTE"],
    }
    for year_index, col in enumerate(year_columns):
        derived_fte_table[col] = [
            _format_number_display(
                cost_state["personnel"][year_index]["Consultant FTE"], 0
            )
        ]
    derived_fte_df = pd.DataFrame(derived_fte_table)
    st.data_editor(
        derived_fte_df,
        hide_index=True,
        key="revenue_model.derived_fte",
        disabled=True,
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            **{col: st.column_config.TextColumn() for col in year_columns},
        },
        use_container_width=True,
    )

    st.markdown("### Revenue Drivers")
    driver_rows = [
        ("Workdays per Year", "Days"),
        ("Utilization %", "%"),
        ("Day Rate Growth (% p.a.)", "%"),
        ("Revenue Growth (% p.a.)", "%"),
    ]
    driver_values = {
        "Workdays per Year": revenue_state["workdays_per_year"][scenario],
        "Utilization %": revenue_state["utilization_rate"][scenario],
        "Day Rate Growth (% p.a.)": revenue_state["day_rate_growth_pct"][scenario],
        "Revenue Growth (% p.a.)": revenue_state["revenue_growth_pct"][scenario],
    }
    driver_table = {
        "Parameter": [row[0] for row in driver_rows],
        "Unit": [row[1] for row in driver_rows],
    }
    for idx, col in enumerate(year_columns):
        driver_table[col] = [
            (
                _format_number_display(
                    _percent_to_display(driver_values[param][idx]), 1
                )
                if param in {"Utilization %", "Day Rate Growth (% p.a.)", "Revenue Growth (% p.a.)"}
                else _format_number_display(driver_values[param][idx], 0)
            )
            for param, _ in driver_rows
        ]
    driver_df = pd.DataFrame(driver_table)
    driver_edit = st.data_editor(
        driver_df,
        hide_index=True,
        key="revenue_model.drivers",
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            **{col: st.column_config.TextColumn() for col in year_columns},
        },
        use_container_width=True,
    )
    for year_index in range(5):
        revenue_state["workdays_per_year"][scenario][year_index] = _non_negative(
            _parse_number_display(driver_edit.loc[0, year_columns[year_index]])
        )
        revenue_state["utilization_rate"][scenario][year_index] = _clamp_pct(
            _percent_from_display(
                _parse_number_display(driver_edit.loc[1, year_columns[year_index]])
            )
        )
        revenue_state["day_rate_growth_pct"][scenario][year_index] = _clamp_pct(
            _percent_from_display(
                _parse_number_display(driver_edit.loc[2, year_columns[year_index]])
            )
        )
        revenue_state["revenue_growth_pct"][scenario][year_index] = _clamp_pct(
            _percent_from_display(
                _parse_number_display(driver_edit.loc[3, year_columns[year_index]])
            )
        )

    st.markdown("### Capacity Allocation")
    allocation_table = {
        "Parameter": ["Group Capacity Share %", "External Capacity Share %"],
        "Unit": ["%", "%"],
    }
    for year_index, col in enumerate(year_columns):
        allocation_table[col] = [
            _format_number_display(
                _percent_to_display(
                    revenue_state["group_capacity_share_pct"][scenario][
                        year_index
                    ]
                ),
                1,
            ),
            _format_number_display(
                _percent_to_display(
                    revenue_state["external_capacity_share_pct"][scenario][
                        year_index
                    ]
                ),
                1,
            ),
        ]
    allocation_df = pd.DataFrame(allocation_table)
    allocation_edit = st.data_editor(
        allocation_df,
        hide_index=True,
        key="revenue_model.capacity_allocation",
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            **{col: st.column_config.TextColumn() for col in year_columns},
        },
        use_container_width=True,
    )
    for year_index in range(5):
        revenue_state["group_capacity_share_pct"][scenario][year_index] = _clamp_pct(
            _percent_from_display(
                _parse_number_display(
                    allocation_edit.loc[0, year_columns[year_index]]
                )
            )
        )
        revenue_state["external_capacity_share_pct"][scenario][
            year_index
        ] = _clamp_pct(
            _percent_from_display(
                _parse_number_display(
                    allocation_edit.loc[1, year_columns[year_index]]
                )
            )
        )

    st.markdown("### Pricing Assumptions")
    pricing_table = {
        "Parameter": ["Group Day Rate (EUR)", "External Day Rate (EUR)"],
        "Unit": ["EUR", "EUR"],
    }
    for year_index, col in enumerate(year_columns):
        pricing_table[col] = [
            _format_number_display(
                revenue_state["group_day_rate_eur"][scenario][year_index], 0
            ),
            _format_number_display(
                revenue_state["external_day_rate_eur"][scenario][year_index], 0
            ),
        ]
    pricing_df = pd.DataFrame(pricing_table)
    pricing_edit = st.data_editor(
        pricing_df,
        hide_index=True,
        key="revenue_model.pricing",
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            **{col: st.column_config.TextColumn() for col in year_columns},
        },
        use_container_width=True,
    )
    for year_index in range(5):
        revenue_state["group_day_rate_eur"][scenario][year_index] = _non_negative(
            _parse_number_display(pricing_edit.loc[0, year_columns[year_index]])
        )
        revenue_state["external_day_rate_eur"][scenario][year_index] = _non_negative(
            _parse_number_display(pricing_edit.loc[1, year_columns[year_index]])
        )

    st.markdown("### Group Revenue Guarantee (Floor)")
    reference_df = pd.DataFrame(
        {
            "Parameter": ["Reference Revenue"],
            "Unit": ["EUR"],
            **{
                col: [
                    _format_number_display(
                        revenue_state["reference_revenue_eur"][scenario], 0
                    )
                ]
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
            **{col: st.column_config.TextColumn() for col in year_columns},
        },
        use_container_width=True,
    )
    revenue_state["reference_revenue_eur"][scenario] = _non_negative(
        _parse_number_display(reference_edit.loc[0, "Year 0"])
    )

    guarantee_df = pd.DataFrame(
        {
            "Parameter": ["Guarantee %"],
            "Unit": ["%"],
            **{
                col: [
                    _format_number_display(
                        _percent_to_display(
                            revenue_state["guarantee_pct_by_year"][scenario][idx]
                        ),
                        1,
                    )
                ]
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
            **{col: st.column_config.TextColumn() for col in year_columns},
        },
        use_container_width=True,
    )
    for year_index in range(5):
        revenue_state["guarantee_pct_by_year"][scenario][year_index] = _clamp_pct(
            _percent_from_display(
                _parse_number_display(
                    guarantee_edit.loc[0, year_columns[year_index]]
                )
            )
        )

    st.markdown("### Revenue Bridge / Summary")
    bridge_rows = {
        "Capacity Revenue": [],
        "Modeled Group Revenue": [],
        "Modeled External Revenue": [],
        "Guaranteed Group Revenue": [],
        "Guaranteed Floor": [],
        "Final Revenue": [],
        "Guaranteed %": [],
    }
    for year_index in range(5):
        reference_value = revenue_state["reference_revenue_eur"][scenario]
        fte = _non_negative(
            cost_state["personnel"][year_index]["Consultant FTE"]
        )
        workdays = revenue_state["workdays_per_year"][scenario][year_index]
        utilization = revenue_state["utilization_rate"][scenario][year_index]
        group_rate = revenue_state["group_day_rate_eur"][scenario][year_index]
        external_rate = revenue_state["external_day_rate_eur"][scenario][year_index]
        rate_growth = revenue_state["day_rate_growth_pct"][scenario][year_index]
        revenue_growth = revenue_state["revenue_growth_pct"][scenario][year_index]
        group_rate = group_rate * ((1 + rate_growth) ** year_index)
        external_rate = external_rate * ((1 + rate_growth) ** year_index)
        capacity_days = fte * workdays * utilization
        adjusted_capacity_days = capacity_days * (1 + revenue_growth)
        group_share = revenue_state["group_capacity_share_pct"][scenario][year_index]
        external_share = revenue_state["external_capacity_share_pct"][scenario][year_index]
        total_share = group_share + external_share
        if total_share > 0:
            group_share = group_share / total_share
            external_share = external_share / total_share
        modeled_group_revenue = (
            adjusted_capacity_days * group_share * group_rate
        )
        modeled_external_revenue = (
            adjusted_capacity_days * external_share * external_rate
        )
        modeled_total = modeled_group_revenue + modeled_external_revenue
        guarantee_pct = revenue_state["guarantee_pct_by_year"][scenario][year_index]
        guaranteed_floor = reference_value * guarantee_pct
        guaranteed_group_revenue = max(
            modeled_group_revenue, guaranteed_floor
        )
        final_revenue = guaranteed_group_revenue + modeled_external_revenue
        guaranteed_share = (
            guaranteed_group_revenue / final_revenue if final_revenue else 0.0
        )
        bridge_rows["Capacity Revenue"].append(modeled_total)
        bridge_rows["Guaranteed Floor"].append(guaranteed_floor)
        bridge_rows["Modeled Group Revenue"].append(modeled_group_revenue)
        bridge_rows["Modeled External Revenue"].append(modeled_external_revenue)
        bridge_rows["Guaranteed Group Revenue"].append(guaranteed_group_revenue)
        bridge_rows["Final Revenue"].append(final_revenue)
        bridge_rows["Guaranteed %"].append(_percent_to_display(guaranteed_share))

    bridge_table = {
        "Parameter": list(bridge_rows.keys()),
        "Unit": [
            "EUR",
            "EUR",
            "EUR",
            "EUR",
            "EUR",
            "EUR",
            "%",
        ],
    }
    for year_index, col in enumerate(year_columns):
        bridge_table[col] = [
            _format_number_display(
                bridge_rows[row_key][year_index],
                1 if row_key == "Guaranteed %" else 0,
            )
            for row_key in bridge_rows
        ]
    bridge_df = pd.DataFrame(bridge_table)
    st.data_editor(
        bridge_df,
        hide_index=True,
        key="revenue_model.bridge",
        disabled=True,
        column_config={
            "Parameter": st.column_config.TextColumn(disabled=True),
            "Unit": st.column_config.TextColumn(disabled=True),
            **{col: st.column_config.TextColumn() for col in year_columns},
        },
        use_container_width=True,
    )
