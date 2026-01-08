import pandas as pd
import streamlit as st

from data_model import InputModel, create_demo_input_model
import run_model as run_model


def _pnl_dict_to_list(pnl_dict):
    # Convert P&L dict to a list with explicit year for downstream models.
    pnl_list = []
    for year_label, year_data in pnl_dict.items():
        try:
            year_index = int(str(year_label).split()[-1])
        except (ValueError, IndexError):
            year_index = year_label
        pnl_list.append({"year": year_index, **year_data})
    return pnl_list


def _format_section_title(section_key):
    return section_key.replace("_", " ").title()


def _render_input_field(
    input_field, widget_key, scenario_tag=None, scenario_help=None
):
    label = input_field.description
    if scenario_tag:
        label = f"{label} {scenario_tag}"
    help_text = f"Excel: {input_field.excel_ref}"
    if scenario_help:
        help_text = f"{help_text} | {scenario_help}"
    value = input_field.value

    if isinstance(value, bool):
        return st.checkbox(
            label,
            value=value,
            help=help_text,
            key=widget_key,
            disabled=not input_field.editable,
        )
    if isinstance(value, (int, float)):
        step = 1.0 if isinstance(value, int) and not isinstance(value, bool) else 0.01
        return st.number_input(
            label,
            value=float(value),
            step=step,
            help=help_text,
            key=widget_key,
            disabled=not input_field.editable,
        )
    if value is None:
        text_value = st.text_input(
            label,
            value="",
            help=help_text,
            key=widget_key,
            disabled=not input_field.editable,
        )
        return None if text_value == "" else text_value

    return st.text_input(
        label,
        value=str(value),
        help=help_text,
        key=widget_key,
        disabled=not input_field.editable,
    )


def _render_section(
    section_data, section_key, selected_scenario=None, is_scenario_section=False
):
    edited_values = {}
    for key, value in section_data.items():
        field_key = f"{section_key}.{key}"
        if hasattr(value, "value") and hasattr(value, "editable"):
            edited_values[key] = _render_input_field(value, field_key)
        elif isinstance(value, dict):
            st.markdown(f"**{_format_section_title(key)}**")
            if is_scenario_section:
                edited_values[key] = {}
                for scenario_key, scenario_field in value.items():
                    scenario_tag = (
                        "(Scenario-driven)"
                        if scenario_key == selected_scenario
                        else "(Scenario parameter)"
                    )
                    scenario_help = (
                        "Active scenario input"
                        if scenario_key == selected_scenario
                        else "Not active for selected scenario"
                    )
                    scenario_field_key = (
                        f"{field_key}.{scenario_key}"
                    )
                    edited_values[key][scenario_key] = _render_input_field(
                        scenario_field,
                        scenario_field_key,
                        scenario_tag=scenario_tag,
                        scenario_help=scenario_help,
                    )
            else:
                edited_values[key] = _render_section(
                    value,
                    field_key,
                    selected_scenario=selected_scenario,
                    is_scenario_section=is_scenario_section,
                )
    return edited_values


def _apply_section_values(section_data, edited_values):
    for key, value in edited_values.items():
        if hasattr(section_data.get(key), "value"):
            section_data[key].value = value
        elif isinstance(section_data.get(key), dict):
            _apply_section_values(section_data[key], value)


def run_app():
    st.title("Financial Model")

    # Build input model and collect editable values from the sidebar.
    base_model = create_demo_input_model()
    edited_values = {}

    with st.sidebar:
        st.markdown("## Assumptions")
        st.write("Sidebar active")

        scenario_options = ["Base", "Best", "Worst"]
        scenario_default = base_model.scenario_selection[
            "selected_scenario"
        ].value
        scenario_index = (
            scenario_options.index(scenario_default)
            if scenario_default in scenario_options
            else 0
        )
        selected_scenario = st.selectbox(
            "Scenario (controls scenario-driven fields)",
            scenario_options,
            index=scenario_index,
        )
        auto_sync = st.checkbox("Auto-update scenario inputs", value=True)

        previous_scenario = st.session_state.get(
            "selected_scenario", selected_scenario
        )
        if auto_sync and selected_scenario != previous_scenario:
            scenario_key = selected_scenario.lower()
            for metric_key, scenario_map in (
                base_model.scenario_parameters.items()
            ):
                if scenario_key in scenario_map:
                    widget_key = (
                        f"scenario_parameters.{metric_key}.{scenario_key}"
                    )
                    st.session_state[widget_key] = scenario_map[
                        scenario_key
                    ].value
        st.session_state["selected_scenario"] = selected_scenario

        for section_key, section_data in base_model.__dict__.items():
            if not isinstance(section_data, dict):
                continue
            section_title = _format_section_title(section_key)
            with st.expander(section_title, expanded=False):
                edited_values[section_key] = _render_section(
                    section_data,
                    section_key,
                    selected_scenario=selected_scenario.lower(),
                    is_scenario_section=section_key == "scenario_parameters",
                )

    input_model = create_demo_input_model()
    if "scenario_selection" in edited_values:
        edited_values["scenario_selection"]["selected_scenario"] = (
            selected_scenario
        )
    for section_key, section_values in edited_values.items():
        _apply_section_values(
            getattr(input_model, section_key), section_values
        )

    # Run model calculations in the standard order.
    pnl_result = run_model.calculate_pnl(input_model)
    pnl_list = _pnl_dict_to_list(pnl_result)
    cashflow_result = run_model.calculate_cashflow(input_model, pnl_list)
    debt_schedule = run_model.calculate_debt_schedule(
        input_model, cashflow_result
    )
    balance_sheet = run_model.calculate_balance_sheet(
        input_model, cashflow_result, debt_schedule
    )
    investment_result = run_model.calculate_investment(
        input_model, cashflow_result, pnl_result
    )

    (
        tab_overview,
        tab_pnl,
        tab_cashflow,
        tab_debt,
        tab_equity,
    ) = st.tabs(
        [
            "Overview",
            "Operating Model (P&L)",
            "Cashflow Details",
            "Debt Schedule",
            "Equity Case",
        ]
    )

    with tab_overview:
        pnl_table = pd.DataFrame.from_dict(pnl_result, orient="index")
        cashflow_table = pd.DataFrame(cashflow_result)

        total_revenue_avg = pnl_table["revenue"].mean()
        ebitda_margin = (
            pnl_table["ebitda"].sum() / pnl_table["revenue"].sum()
            if pnl_table["revenue"].sum() != 0
            else 0
        )
        ebit_avg = pnl_table["ebit"].mean()
        min_cash_balance = cashflow_table["cash_balance"].min()
        irr = investment_result["irr"]

        kpi_col_1, kpi_col_2, kpi_col_3, kpi_col_4, kpi_col_5 = st.columns(5)
        kpi_col_1.metric("Avg Revenue", f"{total_revenue_avg:,.0f} EUR")
        kpi_col_2.metric("EBITDA Margin", f"{ebitda_margin:.1%}")
        kpi_col_3.metric("Avg EBIT", f"{ebit_avg:,.0f} EUR")
        kpi_col_4.metric("Minimum Cash", f"{min_cash_balance:,.0f} EUR")
        kpi_col_5.metric("IRR", f"{irr:.1%}")

        scenario_label = input_model.scenario_selection[
            "selected_scenario"
        ].value
        st.markdown(f"**Scenario:** {scenario_label}")

        st.markdown("### Operating Performance")
        st.write(
            "Revenue, EBITDA, and EBIT reflect the selected scenario and "
            "operating assumptions."
        )

        st.markdown("### Financing Overview")
        st.write(
            "Debt service and cash balance trends summarize financing capacity."
        )

        st.markdown("### Equity Case")
        st.write("IRR and equity cashflows summarize investor outcomes.")

    with tab_pnl:
        pnl_table = pd.DataFrame.from_dict(pnl_result, orient="index")

        total_revenue_avg = pnl_table["revenue"].mean()
        ebitda_margin = (
            pnl_table["ebitda"].sum() / pnl_table["revenue"].sum()
            if pnl_table["revenue"].sum() != 0
            else 0
        )
        ebit_avg = pnl_table["ebit"].mean()
        net_income_avg = pnl_table["net_income"].mean()

        kpi_col_1, kpi_col_2, kpi_col_3, kpi_col_4 = st.columns(4)
        kpi_col_1.metric("Avg Revenue", f"{total_revenue_avg:,.0f} EUR")
        kpi_col_2.metric("EBITDA Margin", f"{ebitda_margin:.1%}")
        kpi_col_3.metric("Avg EBIT", f"{ebit_avg:,.0f} EUR")
        kpi_col_4.metric("Avg Net Income", f"{net_income_avg:,.0f} EUR")

        pnl_display = pnl_table.copy()
        year_labels = []
        for year_label in pnl_display.index:
            try:
                year_index = int(str(year_label).split()[-1])
                year_labels.append(f"Year {year_index}")
            except (ValueError, IndexError):
                year_labels.append(str(year_label))
        pnl_display.insert(0, "year", year_labels)

        pnl_display.rename(
            columns={
                "year": "Year",
                "revenue": "Umsatz (EUR)",
                "personnel_costs": "Personalkosten (EUR)",
                "overhead_and_variable_costs": "Overhead & Variable Kosten (EUR)",
                "ebitda": "EBITDA (EUR)",
                "depreciation": "Abschreibungen (EUR)",
                "ebit": "EBIT (EUR)",
                "taxes": "Steuern (EUR)",
                "net_income": "Jahresueberschuss (EUR)",
            },
            inplace=True,
        )

        pnl_money_columns = [
            "Umsatz (EUR)",
            "Personalkosten (EUR)",
            "Overhead & Variable Kosten (EUR)",
            "EBITDA (EUR)",
            "Abschreibungen (EUR)",
            "EBIT (EUR)",
            "Steuern (EUR)",
            "Jahresueberschuss (EUR)",
        ]
        for col in pnl_money_columns:
            if col in pnl_display.columns:
                pnl_display[col] = pnl_display[col].map(
                    lambda x: f"{x:,.0f}" if pd.notna(x) else ""
                )

        st.table(pnl_display)

    with tab_cashflow:
        cashflow_table = pd.DataFrame(cashflow_result)
        min_cash_balance = cashflow_table["cash_balance"].min()
        avg_operating_cf = cashflow_table["operating_cf"].mean()
        cumulative_cashflow = cashflow_table["net_cashflow"].sum()

        kpi_col_1, kpi_col_2, kpi_col_3 = st.columns(3)
        kpi_col_1.metric("Minimum Cash", f"{min_cash_balance:,.0f} EUR")
        kpi_col_2.metric("Avg Operating CF", f"{avg_operating_cf:,.0f} EUR")
        kpi_col_3.metric("Cumulative CF", f"{cumulative_cashflow:,.0f} EUR")

        cashflow_display = cashflow_table.copy()
        cashflow_display["year"] = cashflow_display["year"].map(
            lambda x: f"Year {int(x)}" if pd.notna(x) else ""
        )
        cashflow_display.rename(
            columns={
                "year": "Year",
                "operating_cf": "Operating CF (EUR)",
                "investing_cf": "Investing CF (EUR)",
                "financing_cf": "Financing CF (EUR)",
                "net_cashflow": "Net Cashflow (EUR)",
                "cash_balance": "Cash Bestand (EUR)",
            },
            inplace=True,
        )
        cashflow_money_columns = [
            "Operating CF (EUR)",
            "Investing CF (EUR)",
            "Financing CF (EUR)",
            "Net Cashflow (EUR)",
            "Cash Bestand (EUR)",
        ]
        for col in cashflow_money_columns:
            if col in cashflow_display.columns:
                cashflow_display[col] = cashflow_display[col].map(
                    lambda x: f"{x:,.0f}" if pd.notna(x) else ""
                )
        st.table(cashflow_display)

    with tab_debt:
        debt_table = pd.DataFrame(debt_schedule)
        initial_debt = (
            debt_table["outstanding_principal"].iloc[0]
            + debt_table["principal_payment"].iloc[0]
        )
        min_dscr = (
            debt_table["dscr"].min() if "dscr" in debt_table.columns else 0
        )

        fully_repaid_year = None
        for _, row in debt_table.iterrows():
            if row["outstanding_principal"] <= 0:
                fully_repaid_year = f"Year {int(row['year'])}"
                break
        debt_repaid_label = (
            f"Yes ({fully_repaid_year})" if fully_repaid_year else "No"
        )

        kpi_col_1, kpi_col_2, kpi_col_3 = st.columns(3)
        kpi_col_1.metric("Initial Debt", f"{initial_debt:,.0f} EUR")
        kpi_col_2.metric("Minimum DSCR", f"{min_dscr:.2f}x")
        kpi_col_3.metric("Debt Fully Repaid", debt_repaid_label)

        debt_display = debt_table.copy()
        debt_display["year"] = debt_display["year"].map(
            lambda x: f"Year {int(x)}" if pd.notna(x) else ""
        )
        debt_display.rename(
            columns={
                "year": "Year",
                "interest_expense": "Zinsaufwand (EUR)",
                "principal_payment": "Tilgung (EUR)",
                "debt_service": "Schuldendienst (EUR)",
                "outstanding_principal": "Restschuld (EUR)",
                "dscr": "DSCR (x)",
            },
            inplace=True,
        )
        debt_money_columns = [
            "Zinsaufwand (EUR)",
            "Tilgung (EUR)",
            "Schuldendienst (EUR)",
            "Restschuld (EUR)",
        ]
        for col in debt_money_columns:
            if col in debt_display.columns:
                debt_display[col] = debt_display[col].map(
                    lambda x: f"{x:,.0f}" if pd.notna(x) else ""
                )
        if "DSCR (x)" in debt_display.columns:
            debt_display["DSCR (x)"] = debt_display["DSCR (x)"].map(
                lambda x: f"{x:.2f}" if pd.notna(x) else ""
            )
        st.table(debt_display)

    with tab_equity:
        summary = {
            "initial_equity": investment_result["initial_equity"],
            "exit_value": investment_result["exit_value"],
            "irr": investment_result["irr"],
        }
        summary_table = pd.DataFrame([summary])

        equity_cashflows = investment_result["equity_cashflows"]
        total_equity_invested = investment_result["initial_equity"]
        total_distributions = sum(
            cf for cf in equity_cashflows if cf > 0
        )
        cash_on_cash_multiple = (
            total_distributions / abs(total_equity_invested)
            if total_equity_invested
            else 0
        )

        kpi_col_1, kpi_col_2, kpi_col_3 = st.columns(3)
        kpi_col_1.metric(
            "Total Equity Invested",
            f"{total_equity_invested:,.0f} EUR",
        )
        kpi_col_2.metric(
            "IRR",
            f"{investment_result['irr']:.1%}",
        )
        kpi_col_3.metric(
            "Cash-on-Cash Multiple",
            f"{cash_on_cash_multiple:.2f}x",
        )

        summary_display = summary_table.copy()
        summary_display.rename(
            columns={
                "initial_equity": "Eigenkapital (Start, EUR)",
                "exit_value": "Exit Value (EUR)",
                "irr": "IRR (%)",
            },
            inplace=True,
        )
        summary_money_columns = [
            "Eigenkapital (Start, EUR)",
            "Exit Value (EUR)",
        ]
        for col in summary_money_columns:
            if col in summary_display.columns:
                summary_display[col] = summary_display[col].map(
                    lambda x: f"{x:,.0f}" if pd.notna(x) else ""
                )
        if "IRR (%)" in summary_display.columns:
            summary_display["IRR (%)"] = summary_display["IRR (%)"].map(
                lambda x: f"{x:.1%}" if pd.notna(x) else ""
            )

        cashflows_table = pd.DataFrame(
            {"equity_cashflows": investment_result["equity_cashflows"]}
        )
        cashflows_display = cashflows_table.copy()
        cashflows_display.insert(
            0, "year", [f"Year {i}" for i in range(len(cashflows_display))]
        )
        cashflows_display.rename(
            columns={
                "year": "Year",
                "equity_cashflows": "Equity Cashflows (EUR)",
            },
            inplace=True,
        )
        cashflows_display["Equity Cashflows (EUR)"] = cashflows_display[
            "Equity Cashflows (EUR)"
        ].map(lambda x: f"{x:,.0f}" if pd.notna(x) else "")

        st.table(summary_display)
        st.table(cashflows_display)


if __name__ == "__main__":
    run_app()
