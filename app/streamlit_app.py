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


def format_currency(value):
    if value is None or pd.isna(value):
        return ""
    return f"{value / 1000:,.0f} kEUR"


def format_pct(value):
    if value is None or pd.isna(value):
        return ""
    return f"{value:.1%}"


def format_int(value):
    if value is None or pd.isna(value):
        return ""
    return f"{int(round(value)):,}"


def _style_totals(df, columns_to_bold):
    def style_row(_):
        return [
            "font-weight: 600;" if col in columns_to_bold else ""
            for col in df.columns
        ]

    return df.style.apply(style_row, axis=1)


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


def _collect_values_from_session(section_data, section_key):
    values = {}
    for key, value in section_data.items():
        field_key = f"{section_key}.{key}"
        if hasattr(value, "value"):
            if field_key in st.session_state:
                values[key] = st.session_state[field_key]
            else:
                values[key] = value.value
        elif isinstance(value, dict):
            values[key] = _collect_values_from_session(value, field_key)
    return values


def _get_field_by_path(base_model, path_parts):
    current = base_model
    for part in path_parts:
        if isinstance(current, dict):
            current = current.get(part)
        else:
            return None
    if hasattr(current, "value"):
        return current
    return None


def _set_field_value(field_key, value):
    st.session_state[field_key] = value


def _render_inline_controls(title, controls, columns=3):
    st.subheader(title)
    cols = st.columns(columns)
    for index, control in enumerate(controls):
        col = cols[index % columns]
        with col:
            widget_key = f"inline.{title}.{control['field_key']}"
            if control["type"] == "select":
                selection = st.selectbox(
                    control["label"],
                    control["options"],
                    index=control["index"],
                    key=widget_key,
                )
                _set_field_value(control["field_key"], selection)
            else:
                if control["type"] == "pct":
                    ui_value = st.number_input(
                        control["label"],
                        value=float(control["value"] * 100),
                        step=0.1,
                        format="%.1f",
                        key=widget_key,
                    )
                    _set_field_value(control["field_key"], ui_value / 100)
                elif control["type"] == "int":
                    value = st.number_input(
                        control["label"],
                        value=float(control["value"]),
                        step=1.0,
                        format="%.0f",
                        key=widget_key,
                    )
                    _set_field_value(control["field_key"], value)
                else:
                    value = st.number_input(
                        control["label"],
                        value=float(control["value"]),
                        step=control.get("step", 1.0),
                        format=control.get("format", "%.0f"),
                        key=widget_key,
                    )
                    _set_field_value(control["field_key"], value)


def run_app():
    st.title("Financial Model")

    # Navigation for question-driven layout.
    with st.sidebar:
        st.markdown("## Navigation")
        page = st.radio(
            "Go to",
            [
                "Overview",
                "Operating Model (P&L)",
                "Cashflow & Liquidity",
                "Balance Sheet",
                "Financing & Debt",
                "Valuation & Purchase Price",
                "Equity Case",
                "Assumptions (Advanced)",
            ],
            key="nav_page",
        )

    # Build input model and collect editable values from the assumptions page.
    base_model = create_demo_input_model()
    selected_scenario = st.session_state.get(
        "scenario_selection.selected_scenario",
        base_model.scenario_selection["selected_scenario"].value,
    )
    scenario_key = selected_scenario.lower()

    if page == "Assumptions (Advanced)":
        st.header("Assumptions (Advanced)")
        st.write(
            "Review and adjust all input assumptions from the Excel sheet."
        )

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
        _set_field_value("scenario_selection.selected_scenario", selected_scenario)
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

        scenario_key = selected_scenario.lower()
        utilization_field = _get_field_by_path(
            base_model.__dict__,
            ["scenario_parameters", "utilization_rate", scenario_key],
        )
        day_rate_field = _get_field_by_path(
            base_model.__dict__,
            ["scenario_parameters", "day_rate_eur", scenario_key],
        )
        purchase_price_field = _get_field_by_path(
            base_model.__dict__,
            ["transaction_and_financing", "purchase_price_eur"],
        )
        equity_field = _get_field_by_path(
            base_model.__dict__,
            ["transaction_and_financing", "equity_contribution_eur"],
        )
        debt_field = _get_field_by_path(
            base_model.__dict__,
            ["transaction_and_financing", "senior_term_loan_start_eur"],
        )
        interest_field = _get_field_by_path(
            base_model.__dict__,
            ["transaction_and_financing", "senior_interest_rate_pct"],
        )

        advanced_controls = [
            {
                "type": "select",
                "label": "Scenario",
                "options": scenario_options,
                "index": scenario_index,
                "field_key": "scenario_selection.selected_scenario",
            },
            {
                "type": "pct",
                "label": "Utilization (%)",
                "field_key": f"scenario_parameters.utilization_rate.{scenario_key}",
                "value": utilization_field.value,
            },
            {
                "type": "number",
                "label": "Day Rate (EUR)",
                "field_key": f"scenario_parameters.day_rate_eur.{scenario_key}",
                "value": day_rate_field.value,
                "step": 100.0,
                "format": "%.0f",
            },
            {
                "type": "number",
                "label": "Purchase Price (EUR)",
                "field_key": "transaction_and_financing.purchase_price_eur",
                "value": purchase_price_field.value,
                "step": 100000.0,
                "format": "%.0f",
            },
            {
                "type": "number",
                "label": "Equity Contribution (EUR)",
                "field_key": "transaction_and_financing.equity_contribution_eur",
                "value": equity_field.value,
                "step": 100000.0,
                "format": "%.0f",
            },
            {
                "type": "number",
                "label": "Debt Amount (EUR)",
                "field_key": "transaction_and_financing.senior_term_loan_start_eur",
                "value": debt_field.value,
                "step": 100000.0,
                "format": "%.0f",
            },
            {
                "type": "pct",
                "label": "Interest Rate (%)",
                "field_key": "transaction_and_financing.senior_interest_rate_pct",
                "value": interest_field.value,
            },
        ]
        _render_inline_controls("Key Inputs", advanced_controls, columns=3)

        section_order = [
            "scenario_parameters",
            "operating_assumptions",
            "personnel_cost_assumptions",
            "overhead_and_variable_costs",
            "capex_and_working_capital",
            "transaction_and_financing",
            "tax_and_distributions",
            "valuation_assumptions",
        ]
        section_labels = {
            "scenario_parameters": "Scenario Parameters",
            "operating_assumptions": "Operating Assumptions",
            "personnel_cost_assumptions": "Personnel Costs",
            "overhead_and_variable_costs": "Overhead & Variable Costs",
            "capex_and_working_capital": "Capex & Working Capital",
            "transaction_and_financing": "Transaction & Financing",
            "tax_and_distributions": "Tax & Valuation",
            "valuation_assumptions": "Tax & Valuation",
        }
        section_help = {
            "scenario_parameters": "Scenario-specific utilization and day-rate inputs.",
            "operating_assumptions": "Headcount and delivery capacity assumptions.",
            "personnel_cost_assumptions": "Compensation, bonus, and payroll drivers.",
            "overhead_and_variable_costs": "Fixed overhead and revenue-linked costs.",
            "capex_and_working_capital": "Capex, depreciation, and liquidity settings.",
            "transaction_and_financing": "Purchase price and debt structure inputs.",
            "tax_and_distributions": "Tax rates and payout assumptions.",
            "valuation_assumptions": "Buyer/seller valuation and DCF parameters.",
        }

        edited_values = {}
        for section_key in section_order:
            if section_key not in base_model.__dict__:
                continue
            section_data = base_model.__dict__[section_key]
            if not isinstance(section_data, dict):
                continue
            section_title = section_labels.get(
                section_key, _format_section_title(section_key)
            )
            with st.expander(section_title, expanded=False):
                st.caption(section_help.get(section_key, ""))
                edited_values[section_key] = _render_section(
                    section_data,
                    section_key,
                    selected_scenario=selected_scenario.lower(),
                    is_scenario_section=section_key == "scenario_parameters",
                )
        st.session_state["edited_values"] = edited_values

    edited_values = {}
    for section_key, section_data in base_model.__dict__.items():
        if not isinstance(section_data, dict):
            continue
        edited_values[section_key] = _collect_values_from_session(
            section_data, section_key
        )

    input_model = create_demo_input_model()
    if "scenario_selection" in edited_values:
        edited_values["scenario_selection"]["selected_scenario"] = (
            st.session_state.get(
                "scenario_selection.selected_scenario",
                selected_scenario,
            )
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

    if page == "Overview":
        st.header("Overview")
        st.write(
            "Top-level view of operating performance, liquidity, and "
            "equity outcomes."
        )
        scenario_options = ["Base", "Best", "Worst"]
        selected_scenario = st.session_state.get(
            "scenario_selection.selected_scenario",
            input_model.scenario_selection["selected_scenario"].value,
        )
        scenario_index = (
            scenario_options.index(selected_scenario)
            if selected_scenario in scenario_options
            else 0
        )
        scenario_key = selected_scenario.lower()
        utilization_field = _get_field_by_path(
            input_model.__dict__,
            ["scenario_parameters", "utilization_rate", scenario_key],
        )
        day_rate_field = _get_field_by_path(
            input_model.__dict__,
            ["scenario_parameters", "day_rate_eur", scenario_key],
        )
        purchase_price_field = _get_field_by_path(
            input_model.__dict__,
            ["transaction_and_financing", "purchase_price_eur"],
        )
        equity_field = _get_field_by_path(
            input_model.__dict__,
            ["transaction_and_financing", "equity_contribution_eur"],
        )
        debt_field = _get_field_by_path(
            input_model.__dict__,
            ["transaction_and_financing", "senior_term_loan_start_eur"],
        )
        interest_field = _get_field_by_path(
            input_model.__dict__,
            ["transaction_and_financing", "senior_interest_rate_pct"],
        )

        overview_controls = [
            {
                "type": "select",
                "label": "Scenario",
                "options": scenario_options,
                "index": scenario_index,
                "field_key": "scenario_selection.selected_scenario",
            },
            {
                "type": "pct",
                "label": "Utilization (%)",
                "field_key": f"scenario_parameters.utilization_rate.{scenario_key}",
                "value": utilization_field.value,
            },
            {
                "type": "number",
                "label": "Day Rate (EUR)",
                "field_key": f"scenario_parameters.day_rate_eur.{scenario_key}",
                "value": day_rate_field.value,
                "step": 100.0,
                "format": "%.0f",
            },
            {
                "type": "number",
                "label": "Purchase Price (EUR)",
                "field_key": "transaction_and_financing.purchase_price_eur",
                "value": purchase_price_field.value,
                "step": 100000.0,
                "format": "%.0f",
            },
            {
                "type": "number",
                "label": "Equity Contribution (EUR)",
                "field_key": "transaction_and_financing.equity_contribution_eur",
                "value": equity_field.value,
                "step": 100000.0,
                "format": "%.0f",
            },
            {
                "type": "number",
                "label": "Debt Amount (EUR)",
                "field_key": "transaction_and_financing.senior_term_loan_start_eur",
                "value": debt_field.value,
                "step": 100000.0,
                "format": "%.0f",
            },
            {
                "type": "pct",
                "label": "Interest Rate (%)",
                "field_key": "transaction_and_financing.senior_interest_rate_pct",
                "value": interest_field.value,
            },
        ]
        _render_inline_controls("Key Inputs", overview_controls, columns=3)
        pnl_table = pd.DataFrame.from_dict(pnl_result, orient="index")
        cashflow_table = pd.DataFrame(cashflow_result)

        purchase_price = input_model.transaction_and_financing[
            "purchase_price_eur"
        ].value
        equity_contribution = input_model.transaction_and_financing[
            "equity_contribution_eur"
        ].value
        initial_debt = input_model.transaction_and_financing[
            "senior_term_loan_start_eur"
        ].value
        opening_cash = (
            cashflow_table["cash_balance"].iloc[0]
            if not cashflow_table.empty
            else 0
        )
        net_debt = initial_debt - opening_cash
        avg_ebitda = pnl_table["ebitda"].mean()
        avg_free_cashflow = (
            cashflow_table["operating_cf"].mean()
            + cashflow_table["investing_cf"].mean()
        )
        min_cash_balance = cashflow_table["cash_balance"].min()
        irr = investment_result["irr"]

        kpi_row_1 = st.columns(3)
        kpi_row_2 = st.columns(3)
        kpi_row_3 = st.columns(1)

        kpi_row_1[0].metric(
            "Purchase Price",
            format_currency(purchase_price),
            help="Transaction assumption: Purchase price input (EUR).",
        )
        kpi_row_1[0].caption("Headline transaction value.")
        kpi_row_1[1].metric(
            "Equity Contribution",
            format_currency(equity_contribution),
            help="Transaction assumption: Equity contribution input (EUR).",
        )
        kpi_row_1[1].caption("Sponsor equity invested at close.")
        kpi_row_1[2].metric(
            "Net Debt",
            format_currency(net_debt),
            help="Calculated as initial senior debt minus opening cash balance.",
        )
        kpi_row_1[2].caption("Leverage after closing cash position.")

        kpi_row_2[0].metric(
            "Avg EBITDA",
            format_currency(avg_ebitda),
            help="Average EBITDA across the 5-year plan.",
        )
        kpi_row_2[0].caption("Operating performance proxy.")
        kpi_row_2[1].metric(
            "Avg Free Cash Flow",
            format_currency(avg_free_cashflow),
            help="Average of Operating CF plus Investing CF across the plan.",
        )
        kpi_row_2[1].caption("Cash generation after capex.")
        kpi_row_2[2].metric(
            "Minimum Cash Balance",
            format_currency(min_cash_balance),
            help="Minimum cash balance observed across all years.",
        )
        kpi_row_2[2].caption("Liquidity low point.")

        kpi_row_3[0].metric(
            "Levered Equity IRR",
            format_pct(irr),
            help="IRR based on equity cashflows including exit value.",
        )
        kpi_row_3[0].caption("Return to equity after leverage.")

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

    if page == "Valuation & Purchase Price":
        st.header("Valuation & Purchase Price")
        st.write(
            "Valuation perspective, purchase price context, and high-level "
            "deal view."
        )
        scenario_options = ["Base", "Best", "Worst"]
        selected_scenario = st.session_state.get(
            "scenario_selection.selected_scenario",
            input_model.scenario_selection["selected_scenario"].value,
        )
        scenario_index = (
            scenario_options.index(selected_scenario)
            if selected_scenario in scenario_options
            else 0
        )
        seller_multiple_field = _get_field_by_path(
            input_model.__dict__,
            ["valuation_assumptions", "multiple_valuation", "seller_multiple"],
        )
        buyer_multiple_field = _get_field_by_path(
            input_model.__dict__,
            ["valuation_assumptions", "multiple_valuation", "buyer_multiple"],
        )
        wacc_field = _get_field_by_path(
            input_model.__dict__,
            ["valuation_assumptions", "dcf_valuation", "discount_rate_wacc"],
        )
        terminal_growth_field = _get_field_by_path(
            input_model.__dict__,
            ["valuation_assumptions", "dcf_valuation", "terminal_growth_rate"],
        )
        forecast_years_field = _get_field_by_path(
            input_model.__dict__,
            ["valuation_assumptions", "dcf_valuation", "explicit_forecast_years"],
        )
        purchase_price_field = _get_field_by_path(
            input_model.__dict__,
            ["transaction_and_financing", "purchase_price_eur"],
        )

        valuation_controls = [
            {
                "type": "select",
                "label": "Scenario",
                "options": scenario_options,
                "index": scenario_index,
                "field_key": "scenario_selection.selected_scenario",
            },
            {
                "type": "number",
                "label": "Seller Multiple (x)",
                "field_key": "valuation_assumptions.multiple_valuation.seller_multiple",
                "value": seller_multiple_field.value or 0,
                "step": 0.1,
                "format": "%.2f",
            },
            {
                "type": "number",
                "label": "Buyer Multiple (x)",
                "field_key": "valuation_assumptions.multiple_valuation.buyer_multiple",
                "value": buyer_multiple_field.value or 0,
                "step": 0.1,
                "format": "%.2f",
            },
            {
                "type": "pct",
                "label": "WACC (%)",
                "field_key": "valuation_assumptions.dcf_valuation.discount_rate_wacc",
                "value": wacc_field.value or 0,
            },
            {
                "type": "pct",
                "label": "Terminal Growth (%)",
                "field_key": "valuation_assumptions.dcf_valuation.terminal_growth_rate",
                "value": terminal_growth_field.value or 0,
            },
            {
                "type": "int",
                "label": "Forecast Years",
                "field_key": "valuation_assumptions.dcf_valuation.explicit_forecast_years",
                "value": forecast_years_field.value or 5,
            },
            {
                "type": "number",
                "label": "Purchase Price (EUR)",
                "field_key": "transaction_and_financing.purchase_price_eur",
                "value": purchase_price_field.value,
                "step": 100000.0,
                "format": "%.0f",
            },
        ]
        _render_inline_controls("Key Inputs", valuation_controls, columns=3)
        st.markdown("### Seller View")
        st.write(
            "Seller view reflects upside valuation using a multiple and a "
            "no-exit DCF based on plan free cash flow."
        )

        seller_multiple = input_model.valuation_assumptions[
            "multiple_valuation"
        ]["seller_multiple"].value
        seller_multiple = 0 if seller_multiple is None else seller_multiple

        avg_ebit = (
            pd.DataFrame.from_dict(pnl_result, orient="index")["ebit"].mean()
        )
        seller_multiple_value = avg_ebit * seller_multiple

        wacc = input_model.valuation_assumptions["dcf_valuation"][
            "discount_rate_wacc"
        ].value
        wacc = 0 if wacc is None else wacc

        free_cashflows = []
        for year_data in cashflow_result[:5]:
            free_cashflows.append(
                year_data["operating_cf"] + year_data["investing_cf"]
            )
        dcf_value = 0
        for i, cashflow in enumerate(free_cashflows, start=1):
            dcf_value += cashflow / ((1 + wacc) ** i) if wacc != 0 else cashflow

        seller_kpi_col_1, seller_kpi_col_2 = st.columns(2)
        seller_kpi_col_1.metric(
            "EBIT Multiple Valuation",
            format_currency(seller_multiple_value),
            help="Average EBIT multiplied by seller multiple assumption.",
        )
        seller_kpi_col_2.metric(
            "DCF Valuation (No Exit)",
            format_currency(dcf_value),
            help="5-year free cash flow discounted at WACC.",
        )

        st.markdown("### Buyer View")
        st.write(
            "Buyer view reflects affordability based on liquidity, "
            "bankability, and equity return thresholds."
        )

        min_cash_balance = (
            pd.DataFrame(cashflow_result)["cash_balance"].min()
        )
        min_dscr = (
            pd.DataFrame(debt_schedule)["dscr"].min()
            if debt_schedule
            else 0
        )
        target_irr = input_model.valuation_assumptions["dcf_valuation"][
            "discount_rate_wacc"
        ].value
        target_irr = 0 if target_irr is None else target_irr
        actual_irr = investment_result["irr"]

        cash_ok = 1 if min_cash_balance >= 0 else 0
        dscr_ratio = min_dscr / 1.3 if 1.3 != 0 else 0
        irr_ratio = (
            actual_irr / target_irr if target_irr not in (0, None) else 1
        )
        affordability_factor = min(cash_ok, dscr_ratio, irr_ratio, 1)

        purchase_price = input_model.transaction_and_financing[
            "purchase_price_eur"
        ].value
        buyer_value = purchase_price * affordability_factor

        buyer_kpi_col_1, buyer_kpi_col_2, buyer_kpi_col_3 = st.columns(3)
        buyer_kpi_col_1.metric(
            "Max Affordable Price",
            format_currency(buyer_value),
            help=(
                "Scaled purchase price based on min cash >= 0, "
                "DSCR >= 1.3, and equity IRR vs target."
            ),
        )
        buyer_kpi_col_2.metric(
            "Minimum DSCR",
            f"{min_dscr:.2f}x",
            help="Lowest DSCR observed across the debt schedule.",
        )
        buyer_kpi_col_3.metric(
            "Equity IRR vs Target",
            f"{actual_irr:.1%} vs {target_irr:.1%}",
            help="Actual equity IRR compared to target hurdle rate.",
        )

        st.markdown("### Valuation Gap")
        st.write(
            "Difference between seller expectations and buyer affordability."
        )
        seller_value = max(seller_multiple_value, dcf_value)
        valuation_gap = seller_value - buyer_value
        gap_col_1, gap_col_2 = st.columns(2)
        gap_col_1.metric(
            "Seller Value",
            format_currency(seller_value),
            help="Higher of seller multiple and DCF valuation.",
        )
        gap_col_2.metric(
            "Buyer Value",
            format_currency(buyer_value),
            help="Affordability-based buyer price ceiling.",
        )
        st.caption(f"Valuation gap: {format_currency(valuation_gap)}")

    if page == "Operating Model (P&L)":
        st.header("Operating Model (P&L)")
        st.write(
            "Detailed revenue and cost build-up based on operating inputs."
        )
        scenario_options = ["Base", "Best", "Worst"]
        selected_scenario = st.session_state.get(
            "scenario_selection.selected_scenario",
            input_model.scenario_selection["selected_scenario"].value,
        )
        scenario_index = (
            scenario_options.index(selected_scenario)
            if selected_scenario in scenario_options
            else 0
        )
        scenario_key = selected_scenario.lower()
        utilization_field = _get_field_by_path(
            input_model.__dict__,
            ["scenario_parameters", "utilization_rate", scenario_key],
        )
        day_rate_field = _get_field_by_path(
            input_model.__dict__,
            ["scenario_parameters", "day_rate_eur", scenario_key],
        )
        fte_field = _get_field_by_path(
            input_model.__dict__,
            ["operating_assumptions", "consulting_fte_start"],
        )
        work_days_field = _get_field_by_path(
            input_model.__dict__,
            ["operating_assumptions", "work_days_per_year"],
        )
        wage_inflation_field = _get_field_by_path(
            input_model.__dict__,
            ["personnel_cost_assumptions", "wage_inflation_pct"],
        )
        operating_controls = [
            {
                "type": "select",
                "label": "Scenario",
                "options": scenario_options,
                "index": scenario_index,
                "field_key": "scenario_selection.selected_scenario",
            },
            {
                "type": "pct",
                "label": "Utilization (%)",
                "field_key": f"scenario_parameters.utilization_rate.{scenario_key}",
                "value": utilization_field.value,
            },
            {
                "type": "number",
                "label": "Day Rate (EUR)",
                "field_key": f"scenario_parameters.day_rate_eur.{scenario_key}",
                "value": day_rate_field.value,
                "step": 100.0,
                "format": "%.0f",
            },
            {
                "type": "int",
                "label": "Consulting FTE",
                "field_key": "operating_assumptions.consulting_fte_start",
                "value": fte_field.value,
            },
            {
                "type": "int",
                "label": "Working Days per Year",
                "field_key": "operating_assumptions.work_days_per_year",
                "value": work_days_field.value,
            },
            {
                "type": "pct",
                "label": "Wage Inflation (%)",
                "field_key": "personnel_cost_assumptions.wage_inflation_pct",
                "value": wage_inflation_field.value,
            },
        ]
        _render_inline_controls("Key Inputs", operating_controls, columns=3)
        explain_calculations = st.toggle("Explain calculations")
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
        kpi_col_1.metric("Avg Revenue", format_currency(total_revenue_avg))
        kpi_col_2.metric("EBITDA Margin", format_pct(ebitda_margin))
        kpi_col_3.metric("Avg EBIT", format_currency(ebit_avg))
        kpi_col_4.metric("Avg Net Income", format_currency(net_income_avg))

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
                "revenue": "Umsatz (kEUR)",
                "personnel_costs": "Personalkosten (kEUR)",
                "overhead_and_variable_costs": "Overhead & Variable Kosten (kEUR)",
                "ebitda": "EBITDA (kEUR)",
                "depreciation": "Abschreibungen (kEUR)",
                "ebit": "EBIT (kEUR)",
                "taxes": "Steuern (kEUR)",
                "net_income": "Jahresueberschuss (kEUR)",
            },
            inplace=True,
        )

        pnl_format_map = {
            "Umsatz (kEUR)": format_currency,
            "Personalkosten (kEUR)": format_currency,
            "Overhead & Variable Kosten (kEUR)": format_currency,
            "EBITDA (kEUR)": format_currency,
            "Abschreibungen (kEUR)": format_currency,
            "EBIT (kEUR)": format_currency,
            "Steuern (kEUR)": format_currency,
            "Jahresueberschuss (kEUR)": format_currency,
        }
        pnl_totals = [
            "Umsatz (kEUR)",
            "EBITDA (kEUR)",
            "EBIT (kEUR)",
            "Jahresueberschuss (kEUR)",
        ]
        pnl_styled = _style_totals(pnl_display, pnl_totals).format(
            pnl_format_map
        )
        st.dataframe(pnl_styled, use_container_width=True)

        if explain_calculations:
            st.markdown("**Revenue build-up**")
            st.write(
                "Revenue = Utilization × Working days × Headcount × Day rate"
            )
            st.markdown("**Personnel cost logic**")
            st.write(
                "Personnel costs = Consulting FTE costs (base + bonus + payroll) "
                "+ Backoffice costs, with wage inflation applied annually."
            )
            st.markdown("**Overhead growth logic**")
            st.write(
                "Fixed overhead grows by overhead inflation; variable overhead "
                "is calculated as a percentage of revenue."
            )

    if page == "Cashflow & Liquidity":
        st.header("Cashflow & Liquidity")
        st.write(
            "Cash generation, investment outflows, and liquidity runway."
        )
        scenario_options = ["Base", "Best", "Worst"]
        selected_scenario = st.session_state.get(
            "scenario_selection.selected_scenario",
            input_model.scenario_selection["selected_scenario"].value,
        )
        scenario_index = (
            scenario_options.index(selected_scenario)
            if selected_scenario in scenario_options
            else 0
        )
        scenario_key = selected_scenario.lower()
        capex_field = _get_field_by_path(
            input_model.__dict__,
            ["capex_and_working_capital", "capex_eur_per_year"],
        )
        depreciation_field = _get_field_by_path(
            input_model.__dict__,
            ["capex_and_working_capital", "depreciation_eur_per_year"],
        )
        dso_field = _get_field_by_path(
            input_model.__dict__,
            ["capex_and_working_capital", "dso_days"],
        )
        min_cash_field = _get_field_by_path(
            input_model.__dict__,
            ["capex_and_working_capital", "minimum_cash_balance_eur"],
        )
        tax_field = _get_field_by_path(
            input_model.__dict__,
            ["tax_and_distributions", "tax_rate_pct"],
        )

        cashflow_controls = [
            {
                "type": "select",
                "label": "Scenario",
                "options": scenario_options,
                "index": scenario_index,
                "field_key": "scenario_selection.selected_scenario",
            },
            {
                "type": "number",
                "label": "Capex (EUR)",
                "field_key": "capex_and_working_capital.capex_eur_per_year",
                "value": capex_field.value,
                "step": 10000.0,
                "format": "%.0f",
            },
            {
                "type": "number",
                "label": "Depreciation (EUR)",
                "field_key": "capex_and_working_capital.depreciation_eur_per_year",
                "value": depreciation_field.value,
                "step": 10000.0,
                "format": "%.0f",
            },
            {
                "type": "int",
                "label": "DSO (days)",
                "field_key": "capex_and_working_capital.dso_days",
                "value": dso_field.value,
            },
            {
                "type": "number",
                "label": "Minimum Cash (EUR)",
                "field_key": "capex_and_working_capital.minimum_cash_balance_eur",
                "value": min_cash_field.value,
                "step": 50000.0,
                "format": "%.0f",
            },
            {
                "type": "pct",
                "label": "Tax Rate (%)",
                "field_key": "tax_and_distributions.tax_rate_pct",
                "value": tax_field.value,
            },
        ]
        _render_inline_controls("Key Inputs", cashflow_controls, columns=3)

        cashflow_table = pd.DataFrame(cashflow_result)
        min_cash_balance = cashflow_table["cash_balance"].min()
        avg_operating_cf = cashflow_table["operating_cf"].mean()
        cumulative_cashflow = cashflow_table["net_cashflow"].sum()

        kpi_col_1, kpi_col_2, kpi_col_3 = st.columns(3)
        kpi_col_1.metric("Minimum Cash", format_currency(min_cash_balance))
        kpi_col_2.metric("Avg Operating CF", format_currency(avg_operating_cf))
        kpi_col_3.metric(
            "Cumulative CF", format_currency(cumulative_cashflow)
        )

        cashflow_display = cashflow_table.copy()
        cashflow_display["year"] = cashflow_display["year"].map(
            lambda x: f"Year {int(x)}" if pd.notna(x) else ""
        )
        cashflow_display.rename(
            columns={
                "year": "Year",
                "operating_cf": "Operating CF (kEUR)",
                "investing_cf": "Investing CF (kEUR)",
                "financing_cf": "Financing CF (kEUR)",
                "net_cashflow": "Net Cashflow (kEUR)",
                "cash_balance": "Cash Balance (kEUR)",
            },
            inplace=True,
        )
        cashflow_format_map = {
            "Operating CF (kEUR)": format_currency,
            "Investing CF (kEUR)": format_currency,
            "Financing CF (kEUR)": format_currency,
            "Net Cashflow (kEUR)": format_currency,
            "Cash Balance (kEUR)": format_currency,
        }
        cashflow_totals = ["Net Cashflow (kEUR)", "Cash Balance (kEUR)"]
        cashflow_styled = _style_totals(
            cashflow_display, cashflow_totals
        ).format(cashflow_format_map)
        st.dataframe(cashflow_styled, use_container_width=True)

    if page == "Balance Sheet":
        st.header("Balance Sheet")
        st.write(
            "Assets, liabilities, and equity position by year."
        )
        scenario_options = ["Base", "Best", "Worst"]
        selected_scenario = st.session_state.get(
            "scenario_selection.selected_scenario",
            input_model.scenario_selection["selected_scenario"].value,
        )
        scenario_index = (
            scenario_options.index(selected_scenario)
            if selected_scenario in scenario_options
            else 0
        )
        min_cash_field = _get_field_by_path(
            input_model.__dict__,
            ["capex_and_working_capital", "minimum_cash_balance_eur"],
        )
        dso_field = _get_field_by_path(
            input_model.__dict__,
            ["capex_and_working_capital", "dso_days"],
        )
        capex_field = _get_field_by_path(
            input_model.__dict__,
            ["capex_and_working_capital", "capex_eur_per_year"],
        )
        depreciation_field = _get_field_by_path(
            input_model.__dict__,
            ["capex_and_working_capital", "depreciation_eur_per_year"],
        )
        tax_field = _get_field_by_path(
            input_model.__dict__,
            ["tax_and_distributions", "tax_rate_pct"],
        )

        balance_controls = [
            {
                "type": "select",
                "label": "Scenario",
                "options": scenario_options,
                "index": scenario_index,
                "field_key": "scenario_selection.selected_scenario",
            },
            {
                "type": "number",
                "label": "Minimum Cash (EUR)",
                "field_key": "capex_and_working_capital.minimum_cash_balance_eur",
                "value": min_cash_field.value,
                "step": 50000.0,
                "format": "%.0f",
            },
            {
                "type": "int",
                "label": "DSO (days)",
                "field_key": "capex_and_working_capital.dso_days",
                "value": dso_field.value,
            },
            {
                "type": "number",
                "label": "Capex (EUR)",
                "field_key": "capex_and_working_capital.capex_eur_per_year",
                "value": capex_field.value,
                "step": 10000.0,
                "format": "%.0f",
            },
            {
                "type": "number",
                "label": "Depreciation (EUR)",
                "field_key": "capex_and_working_capital.depreciation_eur_per_year",
                "value": depreciation_field.value,
                "step": 10000.0,
                "format": "%.0f",
            },
            {
                "type": "pct",
                "label": "Tax Rate (%)",
                "field_key": "tax_and_distributions.tax_rate_pct",
                "value": tax_field.value,
            },
        ]
        _render_inline_controls("Key Inputs", balance_controls, columns=3)

        balance_table = pd.DataFrame(balance_sheet)
        avg_assets = balance_table["assets"].mean()
        avg_liabilities = balance_table["liabilities"].mean()
        avg_equity = balance_table["equity"].mean()

        kpi_col_1, kpi_col_2, kpi_col_3 = st.columns(3)
        kpi_col_1.metric("Avg Assets", format_currency(avg_assets))
        kpi_col_2.metric("Avg Liabilities", format_currency(avg_liabilities))
        kpi_col_3.metric("Avg Equity", format_currency(avg_equity))

        balance_display = balance_table.copy()
        balance_display["year"] = balance_display["year"].map(
            lambda x: f"Year {int(x)}" if pd.notna(x) else ""
        )
        balance_display.rename(
            columns={
                "year": "Year",
                "assets": "Assets (kEUR)",
                "liabilities": "Liabilities (kEUR)",
                "equity": "Equity (kEUR)",
                "working_capital": "Working Capital (kEUR)",
                "retained_earnings": "Retained Earnings (kEUR)",
            },
            inplace=True,
        )
        balance_format_map = {
            "Assets (kEUR)": format_currency,
            "Liabilities (kEUR)": format_currency,
            "Equity (kEUR)": format_currency,
            "Working Capital (kEUR)": format_currency,
            "Retained Earnings (kEUR)": format_currency,
        }
        balance_totals = ["Assets (kEUR)", "Equity (kEUR)"]
        balance_styled = _style_totals(
            balance_display, balance_totals
        ).format(balance_format_map)
        st.dataframe(balance_styled, use_container_width=True)

    if page == "Financing & Debt":
        st.header("Financing & Debt")
        st.write(
            "Debt structure, service capacity, and covenant compliance."
        )
        scenario_options = ["Base", "Best", "Worst"]
        selected_scenario = st.session_state.get(
            "scenario_selection.selected_scenario",
            input_model.scenario_selection["selected_scenario"].value,
        )
        scenario_index = (
            scenario_options.index(selected_scenario)
            if selected_scenario in scenario_options
            else 0
        )
        purchase_price_field = _get_field_by_path(
            input_model.__dict__,
            ["transaction_and_financing", "purchase_price_eur"],
        )
        equity_field = _get_field_by_path(
            input_model.__dict__,
            ["transaction_and_financing", "equity_contribution_eur"],
        )
        debt_field = _get_field_by_path(
            input_model.__dict__,
            ["transaction_and_financing", "senior_term_loan_start_eur"],
        )
        interest_field = _get_field_by_path(
            input_model.__dict__,
            ["transaction_and_financing", "senior_interest_rate_pct"],
        )
        repayment_field = _get_field_by_path(
            input_model.__dict__,
            ["transaction_and_financing", "senior_repayment_per_year_eur"],
        )
        revolver_field = _get_field_by_path(
            input_model.__dict__,
            ["transaction_and_financing", "revolver_limit_eur"],
        )

        financing_controls = [
            {
                "type": "select",
                "label": "Scenario",
                "options": scenario_options,
                "index": scenario_index,
                "field_key": "scenario_selection.selected_scenario",
            },
            {
                "type": "number",
                "label": "Purchase Price (EUR)",
                "field_key": "transaction_and_financing.purchase_price_eur",
                "value": purchase_price_field.value,
                "step": 100000.0,
                "format": "%.0f",
            },
            {
                "type": "number",
                "label": "Equity Contribution (EUR)",
                "field_key": "transaction_and_financing.equity_contribution_eur",
                "value": equity_field.value,
                "step": 100000.0,
                "format": "%.0f",
            },
            {
                "type": "number",
                "label": "Debt Amount (EUR)",
                "field_key": "transaction_and_financing.senior_term_loan_start_eur",
                "value": debt_field.value,
                "step": 100000.0,
                "format": "%.0f",
            },
            {
                "type": "pct",
                "label": "Interest Rate (%)",
                "field_key": "transaction_and_financing.senior_interest_rate_pct",
                "value": interest_field.value,
            },
            {
                "type": "number",
                "label": "Annual Repayment (EUR)",
                "field_key": "transaction_and_financing.senior_repayment_per_year_eur",
                "value": repayment_field.value,
                "step": 100000.0,
                "format": "%.0f",
            },
            {
                "type": "number",
                "label": "Revolver Limit (EUR)",
                "field_key": "transaction_and_financing.revolver_limit_eur",
                "value": revolver_field.value,
                "step": 50000.0,
                "format": "%.0f",
            },
        ]
        _render_inline_controls("Key Inputs", financing_controls, columns=3)
        cashflow_table = pd.DataFrame(cashflow_result)
        min_cash_balance = cashflow_table["cash_balance"].min()
        avg_operating_cf = cashflow_table["operating_cf"].mean()
        cumulative_cashflow = cashflow_table["net_cashflow"].sum()

        kpi_col_1, kpi_col_2, kpi_col_3 = st.columns(3)
        kpi_col_1.metric("Minimum Cash", format_currency(min_cash_balance))
        kpi_col_2.metric("Avg Operating CF", format_currency(avg_operating_cf))
        kpi_col_3.metric(
            "Cumulative CF", format_currency(cumulative_cashflow)
        )
        st.markdown("### Bankability Table")
        st.write(
            "Focus on DSCR compliance and cash headroom by year."
        )

        cashflow_display = cashflow_table.copy()
        cashflow_display["year"] = cashflow_display["year"].map(
            lambda x: f"Year {int(x)}" if pd.notna(x) else ""
        )
        cashflow_display.rename(
            columns={
                "year": "Year",
                "cash_balance": "Cash Balance (kEUR)",
            },
            inplace=True,
        )
        cashflow_display = cashflow_display[
            ["Year", "Cash Balance (kEUR)"]
        ]

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
        kpi_col_1.metric("Initial Debt", format_currency(initial_debt))
        kpi_col_2.metric("Minimum DSCR", f"{min_dscr:.2f}x")
        kpi_col_3.metric("Debt Fully Repaid", debt_repaid_label)

        debt_display = debt_table.copy()
        debt_display["year"] = debt_display["year"].map(
            lambda x: f"Year {int(x)}" if pd.notna(x) else ""
        )
        debt_display.rename(
            columns={
                "year": "Year",
                "interest_expense": "Interest Expense (kEUR)",
                "principal_payment": "Debt Repayment (kEUR)",
                "debt_service": "Debt Service (kEUR)",
                "outstanding_principal": "Debt Outstanding (kEUR)",
                "dscr": "DSCR (x)",
            },
            inplace=True,
        )
        debt_display = debt_display[
            [
                "Year",
                "Debt Outstanding (kEUR)",
                "Interest Expense (kEUR)",
                "Debt Service (kEUR)",
                "DSCR (x)",
            ]
        ]

        bankability_table = debt_display.merge(
            cashflow_display, on="Year", how="left"
        )

        def _highlight_bankability(row):
            styles = []
            for col in bankability_table.columns:
                if col == "DSCR (x)" and pd.notna(row[col]) and row[col] < 1.3:
                    styles.append("background-color: #fdecea;")
                elif (
                    col == "Cash Balance (kEUR)"
                    and pd.notna(row[col])
                    and row[col] < 0
                ):
                    styles.append("background-color: #fdecea;")
                else:
                    styles.append("")
            return styles

        bankability_format_map = {
            "Debt Outstanding (kEUR)": format_currency,
            "Interest Expense (kEUR)": format_currency,
            "Debt Service (kEUR)": format_currency,
            "DSCR (x)": lambda x: f"{x:.2f}" if pd.notna(x) else "",
            "Cash Balance (kEUR)": format_currency,
        }

        bankability_styled = bankability_table.style.apply(
            _highlight_bankability, axis=1
        ).format(bankability_format_map)
        st.dataframe(bankability_styled, use_container_width=True)

        st.markdown("### Bank Commentary")
        st.write(
            "Red rows indicate years with DSCR below 1.3x or negative cash. "
            "These years typically trigger tighter covenants or pricing."
        )

    if page == "Equity Case":
        st.header("Equity Case")
        st.write(
            "Investor returns and exit value based on current assumptions."
        )
        scenario_options = ["Base", "Best", "Worst"]
        selected_scenario = st.session_state.get(
            "scenario_selection.selected_scenario",
            input_model.scenario_selection["selected_scenario"].value,
        )
        scenario_index = (
            scenario_options.index(selected_scenario)
            if selected_scenario in scenario_options
            else 0
        )
        equity_field = _get_field_by_path(
            input_model.__dict__,
            ["transaction_and_financing", "equity_contribution_eur"],
        )
        debt_field = _get_field_by_path(
            input_model.__dict__,
            ["transaction_and_financing", "senior_term_loan_start_eur"],
        )
        seller_multiple_field = _get_field_by_path(
            input_model.__dict__,
            ["valuation_assumptions", "multiple_valuation", "seller_multiple"],
        )
        wacc_field = _get_field_by_path(
            input_model.__dict__,
            ["valuation_assumptions", "dcf_valuation", "discount_rate_wacc"],
        )
        dividend_field = _get_field_by_path(
            input_model.__dict__,
            ["tax_and_distributions", "dividend_payout_ratio_pct"],
        )

        equity_controls = [
            {
                "type": "select",
                "label": "Scenario",
                "options": scenario_options,
                "index": scenario_index,
                "field_key": "scenario_selection.selected_scenario",
            },
            {
                "type": "number",
                "label": "Equity Contribution (EUR)",
                "field_key": "transaction_and_financing.equity_contribution_eur",
                "value": equity_field.value,
                "step": 100000.0,
                "format": "%.0f",
            },
            {
                "type": "number",
                "label": "Debt Amount (EUR)",
                "field_key": "transaction_and_financing.senior_term_loan_start_eur",
                "value": debt_field.value,
                "step": 100000.0,
                "format": "%.0f",
            },
            {
                "type": "number",
                "label": "Seller Multiple (x)",
                "field_key": "valuation_assumptions.multiple_valuation.seller_multiple",
                "value": seller_multiple_field.value or 0,
                "step": 0.1,
                "format": "%.2f",
            },
            {
                "type": "pct",
                "label": "WACC / Target IRR (%)",
                "field_key": "valuation_assumptions.dcf_valuation.discount_rate_wacc",
                "value": wacc_field.value or 0,
            },
            {
                "type": "pct",
                "label": "Dividend Payout (%)",
                "field_key": "tax_and_distributions.dividend_payout_ratio_pct",
                "value": dividend_field.value,
            },
        ]
        _render_inline_controls("Key Inputs", equity_controls, columns=3)
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
            format_currency(total_equity_invested),
        )
        kpi_col_2.metric("IRR", format_pct(summary["irr"]))
        kpi_col_3.metric(
            "Cash-on-Cash Multiple", f"{cash_on_cash_multiple:.2f}x"
        )
        summary_display = summary_table.copy()
        summary_display.rename(
            columns={
                "initial_equity": "Eigenkapital (Start, kEUR)",
                "exit_value": "Exit Value (kEUR)",
                "irr": "IRR (%)",
            },
            inplace=True,
        )
        summary_format_map = {
            "Eigenkapital (Start, kEUR)": format_currency,
            "Exit Value (kEUR)": format_currency,
            "IRR (%)": format_pct,
        }
        summary_totals = ["Eigenkapital (Start, kEUR)", "Exit Value (kEUR)"]
        summary_styled = _style_totals(
            summary_display, summary_totals
        ).format(summary_format_map)

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
                "equity_cashflows": "Equity Cashflows (kEUR)",
            },
            inplace=True,
        )
        cashflows_format_map = {
            "Equity Cashflows (kEUR)": format_currency
        }
        cashflows_styled = _style_totals(
            cashflows_display, ["Equity Cashflows (kEUR)"]
        ).format(cashflows_format_map)

        st.dataframe(summary_styled, use_container_width=True)
        st.dataframe(cashflows_styled, use_container_width=True)


if __name__ == "__main__":
    run_app()
