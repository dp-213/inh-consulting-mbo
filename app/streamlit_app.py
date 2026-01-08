import io
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

st.set_page_config(layout="wide")

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
    if isinstance(value, str):
        return value
    return f"{value / 1_000_000:,.2f} m EUR"


def format_pct(value):
    if value is None or pd.isna(value) or value == "":
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


def _default_cashflow_assumptions():
    return {
        "tax_cash_rate_pct": 0.30,
        "tax_payment_lag_years": 1,
        "capex_pct_revenue": 0.01,
        "working_capital_pct_revenue": 0.02,
        "opening_cash_balance_eur": 250000,
    }


def _default_balance_sheet_assumptions(input_model):
    opening_equity = input_model.transaction_and_financing[
        "equity_contribution_eur"
    ].value
    minimum_cash = input_model.capex_and_working_capital[
        "minimum_cash_balance_eur"
    ].value
    return {
        "opening_equity_eur": opening_equity,
        "depreciation_rate_pct": 0.20,
        "minimum_cash_balance_eur": minimum_cash,
    }


def _default_financing_assumptions(input_model):
    return {
        "initial_debt_eur": input_model.transaction_and_financing[
            "senior_term_loan_start_eur"
        ].value,
        "interest_rate_pct": input_model.transaction_and_financing[
            "senior_interest_rate_pct"
        ].value,
        "amortization_type": "Linear",
        "amortization_period_years": 5,
        "grace_period_years": 0,
        "special_repayment_year": None,
        "special_repayment_amount_eur": 0.0,
        "minimum_dscr": 1.3,
        "minimum_cash_balance_eur": input_model.capex_and_working_capital[
            "minimum_cash_balance_eur"
        ].value,
        "target_irr": 0.25,
        "max_equity_contribution_eur": None,
        "min_cash_yield": 0.08,
    }


def _default_valuation_assumptions(input_model):
    return {
        "seller_ebit_multiple": input_model.valuation_assumptions[
            "multiple_valuation"
        ]["seller_multiple"].value
        or 0.0,
        "reference_year": 1,
        "buyer_discount_rate": input_model.valuation_assumptions["dcf_valuation"][
            "discount_rate_wacc"
        ].value
        or 0.10,
        "valuation_start_year": 0,
        "debt_at_close_eur": input_model.transaction_and_financing[
            "senior_term_loan_start_eur"
        ].value,
        "transaction_cost_pct": 0.01,
        "include_terminal_value": False,
    }


def _seed_session_defaults(input_model):
    def _seed_section(section_data, prefix=""):
        for key, value in section_data.items():
            full_key = f"{prefix}.{key}" if prefix else key
            if hasattr(value, "value"):
                st.session_state.setdefault(full_key, value.value)
            elif isinstance(value, dict):
                _seed_section(value, full_key)

    for section_key, section_value in input_model.__dict__.items():
        if isinstance(section_value, dict):
            _seed_section(section_value, section_key)

    selected_scenario = st.session_state.get(
        "scenario_selection.selected_scenario",
        input_model.scenario_selection["selected_scenario"].value,
    )
    scenario_key = selected_scenario.lower()
    base_utilization = input_model.scenario_parameters["utilization_rate"][
        scenario_key
    ].value
    st.session_state.setdefault(
        "utilization_by_year", [base_utilization] * 5
    )
    for year_index in range(5):
        st.session_state.setdefault(
            f"utilization_by_year.{year_index}",
            st.session_state["utilization_by_year"][year_index],
        )

    for key, value in _default_cashflow_assumptions().items():
        st.session_state.setdefault(f"cashflow.{key}", value)

    for key, value in _default_balance_sheet_assumptions(input_model).items():
        st.session_state.setdefault(f"balance_sheet.{key}", value)

    for key, value in _default_financing_assumptions(input_model).items():
        st.session_state.setdefault(f"financing.{key}", value)
    st.session_state.setdefault("financing.preset", "Custom")

    for key, value in _default_valuation_assumptions(input_model).items():
        st.session_state.setdefault(f"valuation.{key}", value)


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


def _get_current_value(field_key, fallback):
    return st.session_state.get(field_key, fallback)


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
                if isinstance(control["options"][0], str) and control["options"] == [
                    "Off",
                    "On",
                ]:
                    _set_field_value(
                        control["field_key"], True if selection == "On" else False
                    )
                elif isinstance(control["options"][0], str) and any(
                    option.startswith("Year ") for option in control["options"]
                ):
                    if selection == "None":
                        _set_field_value(control["field_key"], None)
                    else:
                        _set_field_value(
                            control["field_key"], int(selection.split()[-1])
                        )
                else:
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


def _render_pnl_html(pnl_statement, section_rows, bold_rows):
    def escape(text):
        return (
            str(text)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

    columns = ["Line Item", "Year 0", "Year 1", "Year 2", "Year 3", "Year 4"]
    header_cells = "".join(f"<th>{escape(col)}</th>" for col in columns)

    body_rows = []
    for _, row in pnl_statement.iterrows():
        label = row["Line Item"]
        row_class = ""
        if label in section_rows:
            row_class = "section-row"
        elif label in bold_rows:
            row_class = "total-row"
        cells = []
        for col in columns:
            value = row[col]
            if col != "Line Item":
                if label in {
                    "EBITDA Margin",
                    "EBIT Margin",
                    "Personnel Cost Ratio",
                    "Guaranteed Revenue %",
                    "Non-Guaranteed Revenue %",
                    "Net Margin",
                    "Opex Ratio",
                }:
                    value = format_pct(value)
                else:
                    value = format_currency(value)
            cell_value = "&nbsp;" if value in ("", None) else escape(value)
            cells.append(f"<td>{cell_value}</td>")
        body_rows.append(f"<tr class=\"{row_class}\">{''.join(cells)}</tr>")

    css = """
    <style>
      .pnl-table { width: 100%; border-collapse: collapse; table-layout: fixed; }
      .pnl-table col.line-item { width: 42%; }
      .pnl-table col.year { width: 11.6%; }
      .pnl-table th, .pnl-table td {
        padding: 2px 6px;
        white-space: nowrap;
        line-height: 1.0;
        border: 0;
        font-size: 0.9rem;
      }
      .pnl-table th { text-align: right; font-weight: 600; }
      .pnl-table th:first-child { text-align: left; }
      .pnl-table td { text-align: right; }
      .pnl-table td:first-child { text-align: left; }
      .pnl-table .section-row td {
        font-weight: 700;
        background: #f9fafb;
      }
      .pnl-table .total-row td {
        font-weight: 700;
        background: #f3f4f6;
        border-top: 1px solid #c7c7c7;
      }
    </style>
    """
    colgroup = (
        "<colgroup>"
        "<col class=\"line-item\"/>"
        "<col class=\"year\"/><col class=\"year\"/><col class=\"year\"/>"
        "<col class=\"year\"/><col class=\"year\"/>"
        "</colgroup>"
    )
    table_html = (
        f"{css}<table class=\"pnl-table\">{colgroup}"
        f"<thead><tr>{header_cells}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody></table>"
    )
    st.markdown(table_html, unsafe_allow_html=True)


def _render_cashflow_html(cashflow_statement, section_rows, bold_rows):
    def escape(text):
        return (
            str(text)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

    columns = ["Line Item", "Year 0", "Year 1", "Year 2", "Year 3", "Year 4"]
    header_cells = "".join(f"<th>{escape(col)}</th>" for col in columns)

    body_rows = []
    for _, row in cashflow_statement.iterrows():
        label = row["Line Item"]
        row_class = ""
        if label in section_rows:
            row_class = "section-row"
        elif label in bold_rows:
            row_class = "total-row"
        cells = []
        for col in columns:
            value = row[col]
            cell_class = ""
            if col != "Line Item":
                value = format_currency(value)
                try:
                    if float(row[col]) < 0:
                        cell_class = "negative"
                except (TypeError, ValueError):
                    cell_class = ""
            cell_value = "&nbsp;" if value in ("", None) else escape(value)
            cells.append(f"<td class=\"{cell_class}\">{cell_value}</td>")
        body_rows.append(f"<tr class=\"{row_class}\">{''.join(cells)}</tr>")

    css = """
    <style>
      .cashflow-table { width: 100%; border-collapse: collapse; table-layout: fixed; }
      .cashflow-table col.line-item { width: 42%; }
      .cashflow-table col.year { width: 11.6%; }
      .cashflow-table th, .cashflow-table td {
        padding: 2px 6px;
        white-space: nowrap;
        line-height: 1.0;
        border: 0;
        font-size: 0.9rem;
      }
      .cashflow-table th { text-align: right; font-weight: 600; }
      .cashflow-table th:first-child { text-align: left; }
      .cashflow-table td { text-align: right; }
      .cashflow-table td:first-child { text-align: left; }
      .cashflow-table .section-row td {
        font-weight: 700;
        background: #f9fafb;
      }
      .cashflow-table .total-row td {
        font-weight: 700;
        background: #f3f4f6;
        border-top: 1px solid #c7c7c7;
      }
      .cashflow-table td.negative { color: #b45309; }
    </style>
    """
    colgroup = (
        "<colgroup>"
        "<col class=\"line-item\"/>"
        "<col class=\"year\"/><col class=\"year\"/><col class=\"year\"/>"
        "<col class=\"year\"/><col class=\"year\"/>"
        "</colgroup>"
    )
    table_html = (
        f"{css}<table class=\"cashflow-table\">{colgroup}"
        f"<thead><tr>{header_cells}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody></table>"
    )
    st.markdown(table_html, unsafe_allow_html=True)


def _render_balance_sheet_html(balance_statement, section_rows, bold_rows):
    def escape(text):
        return (
            str(text)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

    columns = ["Line Item", "Year 0", "Year 1", "Year 2", "Year 3", "Year 4"]
    header_cells = "".join(f"<th>{escape(col)}</th>" for col in columns)

    body_rows = []
    for _, row in balance_statement.iterrows():
        label = row["Line Item"]
        row_class = ""
        if label in section_rows:
            row_class = "section-row"
        elif label in bold_rows:
            row_class = "total-row"
        cells = []
        for col in columns:
            value = row[col]
            cell_class = ""
            if col != "Line Item":
                value = format_currency(value)
                try:
                    if float(row[col]) < 0:
                        cell_class = "negative"
                except (TypeError, ValueError):
                    cell_class = ""
            cell_value = "&nbsp;" if value in ("", None) else escape(value)
            cells.append(f"<td class=\"{cell_class}\">{cell_value}</td>")
        body_rows.append(f"<tr class=\"{row_class}\">{''.join(cells)}</tr>")

    css = """
    <style>
      .balance-table { width: 100%; border-collapse: collapse; table-layout: fixed; }
      .balance-table col.line-item { width: 42%; }
      .balance-table col.year { width: 11.6%; }
      .balance-table th, .balance-table td {
        padding: 2px 6px;
        white-space: nowrap;
        line-height: 1.0;
        border: 0;
        font-size: 0.9rem;
      }
      .balance-table th { text-align: right; font-weight: 600; }
      .balance-table th:first-child { text-align: left; }
      .balance-table td { text-align: right; }
      .balance-table td:first-child { text-align: left; }
      .balance-table .section-row td {
        font-weight: 700;
        background: #f9fafb;
      }
      .balance-table .total-row td {
        font-weight: 700;
        background: #f3f4f6;
        border-top: 1px solid #c7c7c7;
      }
      .balance-table td.negative { color: #b45309; }
    </style>
    """
    colgroup = (
        "<colgroup>"
        "<col class=\"line-item\"/>"
        "<col class=\"year\"/><col class=\"year\"/><col class=\"year\"/>"
        "<col class=\"year\"/><col class=\"year\"/>"
        "</colgroup>"
    )
    table_html = (
        f"{css}<table class=\"balance-table\">{colgroup}"
        f"<thead><tr>{header_cells}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody></table>"
    )
    st.markdown(table_html, unsafe_allow_html=True)


def _render_custom_table_html(
    statement, section_rows, bold_rows, row_formatters=None
):
    def escape(text):
        return (
            str(text)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

    columns = ["Line Item", "Year 0", "Year 1", "Year 2", "Year 3", "Year 4"]
    header_cells = "".join(f"<th>{escape(col)}</th>" for col in columns)
    row_formatters = row_formatters or {}

    body_rows = []
    for _, row in statement.iterrows():
        label = row["Line Item"]
        row_class = ""
        if label in section_rows:
            row_class = "section-row"
        elif label in bold_rows:
            row_class = "total-row"
        cells = []
        for col in columns:
            value = row[col]
            cell_class = ""
            if col != "Line Item":
                formatter = row_formatters.get(label, format_currency)
                if value is None or value == "" or pd.isna(value):
                    value = ""
                else:
                    value = formatter(value)
                try:
                    if float(row[col]) < 0:
                        cell_class = "negative"
                except (TypeError, ValueError):
                    if label == "Covenant Breach" and value == "YES":
                        cell_class = "negative"
                    else:
                        cell_class = ""
            cell_value = "&nbsp;" if value in ("", None) else escape(value)
            cells.append(f"<td class=\"{cell_class}\">{cell_value}</td>")
        body_rows.append(f"<tr class=\"{row_class}\">{''.join(cells)}</tr>")

    css = """
    <style>
      .custom-table { width: 100%; border-collapse: collapse; table-layout: fixed; }
      .custom-table col.line-item { width: 42%; }
      .custom-table col.year { width: 11.6%; }
      .custom-table th, .custom-table td {
        padding: 2px 6px;
        white-space: nowrap;
        line-height: 1.0;
        border: 0;
        font-size: 0.9rem;
      }
      .custom-table th { text-align: right; font-weight: 600; }
      .custom-table th:first-child { text-align: left; }
      .custom-table td { text-align: right; }
      .custom-table td:first-child { text-align: left; }
      .custom-table .section-row td {
        font-weight: 700;
        background: #f9fafb;
      }
      .custom-table .total-row td {
        font-weight: 700;
        background: #f3f4f6;
        border-top: 1px solid #c7c7c7;
      }
      .custom-table td.negative { color: #b45309; }
    </style>
    """
    colgroup = (
        "<colgroup>"
        "<col class=\"line-item\"/>"
        "<col class=\"year\"/><col class=\"year\"/><col class=\"year\"/>"
        "<col class=\"year\"/><col class=\"year\"/>"
        "</colgroup>"
    )
    table_html = (
        f"{css}<table class=\"custom-table\">{colgroup}"
        f"<thead><tr>{header_cells}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody></table>"
    )
    st.markdown(table_html, unsafe_allow_html=True)


def _build_pnl_excel(input_model):
    scenario = input_model.scenario_selection["selected_scenario"].value
    scenario_key = scenario.lower()
    cashflow_assumptions = getattr(
        input_model, "cashflow_assumptions", _default_cashflow_assumptions()
    )
    balance_assumptions = getattr(
        input_model,
        "balance_sheet_assumptions",
        _default_balance_sheet_assumptions(input_model),
    )
    valuation_runtime = getattr(
        input_model,
        "valuation_runtime",
        _default_valuation_assumptions(input_model),
    )
    financing_assumptions = getattr(
        input_model,
        "financing_assumptions",
        _default_financing_assumptions(input_model),
    )

    assumptions = [
        ("Consulting FTE", input_model.operating_assumptions["consulting_fte_start"].value),
        ("FTE Growth %", input_model.operating_assumptions["consulting_fte_growth_pct"].value),
        ("Workdays per Year", input_model.operating_assumptions["work_days_per_year"].value),
        ("Utilization %", input_model.scenario_parameters["utilization_rate"][scenario_key].value),
        ("Day Rate (EUR)", input_model.scenario_parameters["day_rate_eur"][scenario_key].value),
        ("Day Rate Growth %", input_model.operating_assumptions["day_rate_growth_pct"].value),
        ("Guarantee % Year 1", input_model.operating_assumptions["revenue_guarantee_pct_year_1"].value),
        ("Guarantee % Year 2", input_model.operating_assumptions["revenue_guarantee_pct_year_2"].value),
        ("Guarantee % Year 3", input_model.operating_assumptions["revenue_guarantee_pct_year_3"].value),
        ("Consultant Base Cost (EUR)", input_model.personnel_cost_assumptions["avg_consultant_base_cost_eur_per_year"].value),
        ("Bonus %", input_model.personnel_cost_assumptions["bonus_pct_of_base"].value),
        ("Payroll Burden %", input_model.personnel_cost_assumptions["payroll_burden_pct_of_comp"].value),
        ("Wage Inflation %", input_model.personnel_cost_assumptions["wage_inflation_pct"].value),
        ("Backoffice FTE", input_model.operating_assumptions["backoffice_fte_start"].value),
        ("Backoffice Growth %", input_model.operating_assumptions["backoffice_fte_growth_pct"].value),
        ("Backoffice Salary (EUR)", input_model.operating_assumptions["avg_backoffice_salary_eur_per_year"].value),
        ("Overhead Inflation %", input_model.overhead_and_variable_costs["overhead_inflation_pct"].value),
        ("External Advisors (EUR)", input_model.overhead_and_variable_costs["legal_audit_eur_per_year"].value),
        ("IT (EUR)", input_model.overhead_and_variable_costs["it_and_software_eur_per_year"].value),
        ("Office (EUR)", input_model.overhead_and_variable_costs["rent_eur_per_year"].value),
        ("Insurance (EUR)", input_model.overhead_and_variable_costs["insurance_eur_per_year"].value),
        ("Other Services (EUR)", input_model.overhead_and_variable_costs["other_overhead_eur_per_year"].value),
        ("Depreciation (EUR)", input_model.capex_and_working_capital["depreciation_eur_per_year"].value),
        ("Purchase Price (EUR)", input_model.transaction_and_financing["purchase_price_eur"].value),
        ("Equity Contribution (EUR)", input_model.transaction_and_financing["equity_contribution_eur"].value),
        ("Debt Amount (EUR)", input_model.transaction_and_financing["senior_term_loan_start_eur"].value),
        ("Interest Rate %", input_model.transaction_and_financing["senior_interest_rate_pct"].value),
        ("Annual Debt Repayment (EUR)", input_model.transaction_and_financing["senior_repayment_per_year_eur"].value),
        ("Tax Rate %", input_model.tax_and_distributions["tax_rate_pct"].value),
        ("Tax Cash Rate (%)", cashflow_assumptions["tax_cash_rate_pct"]),
        ("Tax Payment Lag (Years)", cashflow_assumptions["tax_payment_lag_years"]),
        ("Capex (% of Revenue)", cashflow_assumptions["capex_pct_revenue"]),
        ("Working Capital Adjustment (% of Revenue)", cashflow_assumptions["working_capital_pct_revenue"]),
        ("Opening Cash Balance (EUR)", cashflow_assumptions["opening_cash_balance_eur"]),
    ]
    assumptions_df = pd.DataFrame(assumptions, columns=["Item", "Value"])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        assumptions_df.to_excel(writer, sheet_name="Assumptions", index=False)
        wb = writer.book
        ws_pnl = wb.create_sheet("P&L")
        ws_kpi = wb.create_sheet("KPIs")
        ws_cashflow = wb.create_sheet("Cashflow")
        ws_balance = wb.create_sheet("Balance Sheet")
        ws_valuation = wb.create_sheet("Valuation")
        ws_financing = wb.create_sheet("Financing & Debt")
        ws_financing_notes = wb.create_sheet("Financing Notes")

        assumption_cells = {
            name: f"Assumptions!B{idx + 2}" for idx, (name, _) in enumerate(assumptions)
        }

        def year_col(col_index):
            return get_column_letter(col_index)

        year_headers = ["Year 0", "Year 1", "Year 2", "Year 3", "Year 4"]
        ws_pnl["A1"] = "Line Item"
        for idx, header in enumerate(year_headers, start=2):
            ws_pnl.cell(row=1, column=idx, value=header)

        line_items = [
            "Revenue",
            "Guaranteed Revenue",
            "Non-Guaranteed Revenue",
            "Total Revenue",
            "Personnel Costs",
            "Consultant Compensation",
            "Backoffice Compensation",
            "Management / MD Compensation",
            "Total Personnel Costs",
            "Operating Expenses",
            "External Consulting / Advisors",
            "IT",
            "Office",
            "Other Services",
            "Total Operating Expenses",
            "EBITDA",
            "Depreciation",
            "EBIT",
            "Interest Expense",
            "EBT",
            "Taxes",
            "Net Income (Jahresüberschuss)",
        ]

        for row_idx, item in enumerate(line_items, start=2):
            ws_pnl.cell(row=row_idx, column=1, value=item)

        for year_index in range(5):
            col = year_col(2 + year_index)
            fte = f"({assumption_cells['Consulting FTE']}*(1+{assumption_cells['FTE Growth %']})^{year_index})"
            workdays = assumption_cells["Workdays per Year"]
            utilization = assumption_cells["Utilization %"]
            day_rate = f"({assumption_cells['Day Rate (EUR)']}*(1+{assumption_cells['Day Rate Growth %']})^{year_index})"
            guarantee_pct = "0"
            if year_index == 0:
                guarantee_pct = assumption_cells["Guarantee % Year 1"]
            elif year_index == 1:
                guarantee_pct = assumption_cells["Guarantee % Year 2"]
            elif year_index == 2:
                guarantee_pct = assumption_cells["Guarantee % Year 3"]

            guaranteed = f"={fte}*{workdays}*{day_rate}*{guarantee_pct}"
            non_guaranteed = f"={fte}*{workdays}*{day_rate}*MAX({utilization}-{guarantee_pct},0)"
            total_revenue = f"={col}3+{col}4"

            consultant_cost = (
                f"={fte}*{assumption_cells['Consultant Base Cost (EUR)']}*"
                f"(1+{assumption_cells['Bonus %']}+{assumption_cells['Payroll Burden %']})*"
                f"(1+{assumption_cells['Wage Inflation %']})^{year_index}"
            )
            backoffice_fte = f"({assumption_cells['Backoffice FTE']}*(1+{assumption_cells['Backoffice Growth %']})^{year_index})"
            backoffice_cost = (
                f"={backoffice_fte}*{assumption_cells['Backoffice Salary (EUR)']}*"
                f"(1+{assumption_cells['Payroll Burden %']})*"
                f"(1+{assumption_cells['Wage Inflation %']})^{year_index}"
            )
            management_cost = "=0"
            total_personnel = f"={col}7+{col}8+{col}9"

            external_advisors = f"={assumption_cells['External Advisors (EUR)']}*(1+{assumption_cells['Overhead Inflation %']})^{year_index}"
            it_cost = f"={assumption_cells['IT (EUR)']}*(1+{assumption_cells['Overhead Inflation %']})^{year_index}"
            office_cost = f"={assumption_cells['Office (EUR)']}*(1+{assumption_cells['Overhead Inflation %']})^{year_index}"
            other_services = f"=({assumption_cells['Insurance (EUR)']}+{assumption_cells['Other Services (EUR)']})*(1+{assumption_cells['Overhead Inflation %']})^{year_index}"
            total_opex = f"={col}12+{col}13+{col}14+{col}15"

            ebitda = f"={col}5-{col}10-{col}16"
            depreciation = f"={assumption_cells['Depreciation (EUR)']}"
            ebit = f"={col}17-{col}18"
            interest = f"={assumption_cells['Debt Amount (EUR)']}*{assumption_cells['Interest Rate %']}"
            ebt = f"={col}19-{col}20"
            taxes = f"=MAX({col}21,0)*{assumption_cells['Tax Rate %']}"
            net_income = f"={col}21-{col}22"

            ws_pnl[f"{col}3"] = guaranteed
            ws_pnl[f"{col}4"] = non_guaranteed
            ws_pnl[f"{col}5"] = total_revenue
            ws_pnl[f"{col}7"] = consultant_cost
            ws_pnl[f"{col}8"] = backoffice_cost
            ws_pnl[f"{col}9"] = management_cost
            ws_pnl[f"{col}10"] = total_personnel
            ws_pnl[f"{col}12"] = external_advisors
            ws_pnl[f"{col}13"] = it_cost
            ws_pnl[f"{col}14"] = office_cost
            ws_pnl[f"{col}15"] = other_services
            ws_pnl[f"{col}16"] = total_opex
            ws_pnl[f"{col}17"] = ebitda
            ws_pnl[f"{col}18"] = depreciation
            ws_pnl[f"{col}19"] = ebit
            ws_pnl[f"{col}20"] = interest
            ws_pnl[f"{col}21"] = ebt
            ws_pnl[f"{col}22"] = taxes
            ws_pnl[f"{col}23"] = net_income

        ws_kpi["A1"] = "KPI"
        for idx, header in enumerate(year_headers, start=2):
            ws_kpi.cell(row=1, column=idx, value=header)

        kpis = [
            "Revenue per Consultant",
            "EBITDA Margin",
            "EBIT Margin",
            "Personnel Cost Ratio",
            "Opex Ratio",
            "Net Margin",
            "Guaranteed Revenue %",
        ]
        for row_idx, kpi in enumerate(kpis, start=2):
            ws_kpi.cell(row=row_idx, column=1, value=kpi)

        for year_index in range(5):
            col = year_col(2 + year_index)
            fte = f"({assumption_cells['Consulting FTE']}*(1+{assumption_cells['FTE Growth %']})^{year_index})"
            ws_kpi[f"{col}2"] = f"='P&L'!{col}5/{fte}"
            ws_kpi[f"{col}3"] = f"='P&L'!{col}17/'P&L'!{col}5"
            ws_kpi[f"{col}4"] = f"='P&L'!{col}19/'P&L'!{col}5"
            ws_kpi[f"{col}5"] = f"='P&L'!{col}10/'P&L'!{col}5"
            ws_kpi[f"{col}6"] = f"='P&L'!{col}16/'P&L'!{col}5"
            ws_kpi[f"{col}7"] = f"='P&L'!{col}23/'P&L'!{col}5"
            ws_kpi[f"{col}8"] = f"='P&L'!{col}3/'P&L'!{col}5"

        ws_cashflow["A1"] = "Cashflow Assumptions"
        cashflow_assumption_rows = [
            ("Tax Cash Rate (%)", "Tax Cash Rate (%)"),
            ("Tax Payment Lag (Years)", "Tax Payment Lag (Years)"),
            ("Capex (% of Revenue)", "Capex (% of Revenue)"),
            ("Working Capital Adjustment (% of Revenue)", "Working Capital Adjustment (% of Revenue)"),
            ("Opening Cash Balance (EUR)", "Opening Cash Balance (EUR)"),
        ]
        for idx, (label, key) in enumerate(cashflow_assumption_rows, start=2):
            ws_cashflow.cell(row=idx, column=1, value=label)
            ws_cashflow.cell(
                row=idx, column=2, value=f"={assumption_cells[key]}"
            )

        cashflow_table_start = 8
        ws_cashflow.cell(row=cashflow_table_start, column=1, value="Line Item")
        for idx, header in enumerate(year_headers, start=2):
            ws_cashflow.cell(row=cashflow_table_start, column=idx, value=header)

        cashflow_items = [
            "OPERATING CASHFLOW",
            "EBITDA",
            "Taxes Paid",
            "Working Capital Change",
            "Operating Cashflow",
            "INVESTING CASHFLOW",
            "Capex",
            "Free Cashflow",
            "FINANCING CASHFLOW",
            "Debt Drawdown",
            "Interest Paid",
            "Debt Repayment",
            "Net Cashflow",
            "LIQUIDITY",
            "Opening Cash",
            "Net Cashflow",
            "Closing Cash",
        ]
        for row_idx, item in enumerate(
            cashflow_items, start=cashflow_table_start + 1
        ):
            ws_cashflow.cell(row=row_idx, column=1, value=item)

        tax_cash_cell = assumption_cells["Tax Cash Rate (%)"]
        tax_lag_cell = assumption_cells["Tax Payment Lag (Years)"]
        capex_pct_cell = assumption_cells["Capex (% of Revenue)"]
        wc_pct_cell = assumption_cells["Working Capital Adjustment (% of Revenue)"]
        opening_cash_cell = assumption_cells["Opening Cash Balance (EUR)"]
        debt_amount_cell = assumption_cells["Debt Amount (EUR)"]
        interest_rate_cell = assumption_cells["Interest Rate %"]
        repayment_cell = assumption_cells["Annual Debt Repayment (EUR)"]

        cashflow_row_map = {
            "EBITDA": cashflow_table_start + 2,
            "Taxes Paid": cashflow_table_start + 3,
            "Working Capital Change": cashflow_table_start + 4,
            "Operating Cashflow": cashflow_table_start + 5,
            "Capex": cashflow_table_start + 7,
            "Free Cashflow": cashflow_table_start + 8,
            "Debt Drawdown": cashflow_table_start + 10,
            "Interest Paid": cashflow_table_start + 11,
            "Debt Repayment": cashflow_table_start + 12,
            "Net Cashflow (Financing)": cashflow_table_start + 13,
            "Opening Cash": cashflow_table_start + 15,
            "Net Cashflow (Liquidity)": cashflow_table_start + 16,
            "Closing Cash": cashflow_table_start + 17,
        }

        for year_index in range(5):
            col = year_col(2 + year_index)
            prev_col = year_col(1 + year_index) if year_index > 0 else None
            ws_cashflow[f"{col}{cashflow_row_map['EBITDA']}"] = f"='P&L'!{col}17"

            taxes_due = f"MAX('P&L'!{col}21,0)*{tax_cash_cell}"
            if year_index == 0:
                taxes_prev = "0"
            else:
                taxes_prev = f"MAX('P&L'!{prev_col}21,0)*{tax_cash_cell}"
            ws_cashflow[f"{col}{cashflow_row_map['Taxes Paid']}"] = (
                f"=IF({tax_lag_cell}=0,{taxes_due},IF({tax_lag_cell}=1,{taxes_prev},0))"
            )

            ws_cashflow[f"{col}{cashflow_row_map['Working Capital Change']}"] = (
                f"='P&L'!{col}5*{wc_pct_cell}"
            )
            ws_cashflow[f"{col}{cashflow_row_map['Operating Cashflow']}"] = (
                f"={col}{cashflow_row_map['EBITDA']}"
                f"-{col}{cashflow_row_map['Taxes Paid']}"
                f"-{col}{cashflow_row_map['Working Capital Change']}"
            )
            ws_cashflow[f"{col}{cashflow_row_map['Capex']}"] = (
                f"='P&L'!{col}5*{capex_pct_cell}"
            )
            ws_cashflow[f"{col}{cashflow_row_map['Free Cashflow']}"] = (
                f"={col}{cashflow_row_map['Operating Cashflow']}"
                f"-{col}{cashflow_row_map['Capex']}"
            )

            if year_index == 0:
                debt_drawdown = f"={debt_amount_cell}"
            else:
                debt_drawdown = "=0"
            ws_cashflow[f"{col}{cashflow_row_map['Debt Drawdown']}"] = debt_drawdown

            outstanding = f"MAX({debt_amount_cell}-{repayment_cell}*{year_index},0)"
            ws_cashflow[f"{col}{cashflow_row_map['Interest Paid']}"] = (
                f"={outstanding}*{interest_rate_cell}"
            )
            ws_cashflow[f"{col}{cashflow_row_map['Debt Repayment']}"] = (
                f"=MIN({repayment_cell},{outstanding})"
            )
            net_cashflow_formula = (
                f"={col}{cashflow_row_map['Free Cashflow']}"
                f"+{col}{cashflow_row_map['Debt Drawdown']}"
                f"-{col}{cashflow_row_map['Interest Paid']}"
                f"-{col}{cashflow_row_map['Debt Repayment']}"
            )
            ws_cashflow[
                f"{col}{cashflow_row_map['Net Cashflow (Financing)']}"
            ] = net_cashflow_formula
            ws_cashflow[
                f"{col}{cashflow_row_map['Net Cashflow (Liquidity)']}"
            ] = net_cashflow_formula

            if year_index == 0:
                ws_cashflow[f"{col}{cashflow_row_map['Opening Cash']}"] = (
                    f"={opening_cash_cell}"
                )
            else:
                ws_cashflow[f"{col}{cashflow_row_map['Opening Cash']}"] = (
                    f"={prev_col}{cashflow_row_map['Closing Cash']}"
                )
            ws_cashflow[f"{col}{cashflow_row_map['Closing Cash']}"] = (
                f"={col}{cashflow_row_map['Opening Cash']}"
                f"+{col}{cashflow_row_map['Net Cashflow (Liquidity)']}"
            )

        notes_row = cashflow_table_start + len(cashflow_items) + 2
        ws_cashflow.cell(row=notes_row, column=1, value="Notes")
        ws_cashflow.cell(
            row=notes_row + 1,
            column=1,
            value="Operating CF = EBITDA - Taxes Paid - Working Capital Change.",
        )
        ws_cashflow.cell(
            row=notes_row + 2,
            column=1,
            value="Capex and Working Capital are modeled as % of revenue.",
        )
        ws_cashflow.cell(
            row=notes_row + 3,
            column=1,
            value="Closing Cash = Opening Cash + Net Cashflow.",
        )

        ws_balance["A1"] = "Balance Sheet Assumptions"
        balance_assumption_rows = [
            ("Opening Equity (EUR)", balance_assumptions["opening_equity_eur"]),
            ("Depreciation Rate (%)", balance_assumptions["depreciation_rate_pct"]),
            ("Minimum Cash Balance (EUR)", balance_assumptions["minimum_cash_balance_eur"]),
            ("Opening Fixed Assets (EUR)", 0),
        ]
        for idx, (label, value) in enumerate(balance_assumption_rows, start=2):
            ws_balance.cell(row=idx, column=1, value=label)
            ws_balance.cell(row=idx, column=2, value=value)

        balance_table_start = 8
        ws_balance.cell(row=balance_table_start, column=1, value="Line Item")
        for idx, header in enumerate(year_headers, start=2):
            ws_balance.cell(row=balance_table_start, column=idx, value=header)

        balance_items = [
            "ASSETS",
            "Cash",
            "Fixed Assets (Net)",
            "Total Assets",
            "LIABILITIES",
            "Financial Debt",
            "Total Liabilities",
            "EQUITY",
            "Equity at Start of Year",
            "Net Income",
            "Dividends",
            "Equity at End of Year",
            "CHECK",
            "Total Assets",
            "Total Liabilities + Equity",
        ]
        for row_idx, item in enumerate(
            balance_items, start=balance_table_start + 1
        ):
            ws_balance.cell(row=row_idx, column=1, value=item)

        balance_row_map = {
            "Cash": balance_table_start + 2,
            "Fixed Assets (Net)": balance_table_start + 3,
            "Total Assets (Assets)": balance_table_start + 4,
            "Financial Debt": balance_table_start + 6,
            "Total Liabilities": balance_table_start + 7,
            "Equity at Start of Year": balance_table_start + 9,
            "Net Income": balance_table_start + 10,
            "Dividends": balance_table_start + 11,
            "Equity at End of Year": balance_table_start + 12,
            "Total Assets (Check)": balance_table_start + 14,
            "Total Liabilities + Equity": balance_table_start + 15,
        }

        opening_equity_cell = "B2"
        depreciation_rate_cell = "B3"
        opening_fixed_assets_cell = "B5"
        debt_amount_cell = assumption_cells["Debt Amount (EUR)"]
        repayment_cell = assumption_cells["Annual Debt Repayment (EUR)"]

        for year_index in range(5):
            col = year_col(2 + year_index)
            prev_col = year_col(1 + year_index) if year_index > 0 else None
            ws_balance[f"{col}{balance_row_map['Cash']}"] = (
                f"=Cashflow!{col}{cashflow_row_map['Closing Cash']}"
            )

            if year_index == 0:
                ws_balance[f"{col}{balance_row_map['Fixed Assets (Net)']}"] = (
                    f"=MAX({opening_fixed_assets_cell}"
                    f"+Cashflow!{col}{cashflow_row_map['Capex']}"
                    f"-({opening_fixed_assets_cell}"
                    f"+Cashflow!{col}{cashflow_row_map['Capex']})*{depreciation_rate_cell},0)"
                )
                ws_balance[f"{col}{balance_row_map['Equity at Start of Year']}"] = (
                    f"={opening_equity_cell}"
                )
            else:
                ws_balance[f"{col}{balance_row_map['Fixed Assets (Net)']}"] = (
                    f"=MAX({prev_col}{balance_row_map['Fixed Assets (Net)']}"
                    f"+Cashflow!{col}{cashflow_row_map['Capex']}"
                    f"-({prev_col}{balance_row_map['Fixed Assets (Net)']}"
                    f"+Cashflow!{col}{cashflow_row_map['Capex']})*{depreciation_rate_cell},0)"
                )
                ws_balance[f"{col}{balance_row_map['Equity at Start of Year']}"] = (
                    f"={prev_col}{balance_row_map['Equity at End of Year']}"
                )

            ws_balance[f"{col}{balance_row_map['Total Assets (Assets)']}"] = (
                f"={col}{balance_row_map['Cash']}+{col}{balance_row_map['Fixed Assets (Net)']}"
            )
            ws_balance[f"{col}{balance_row_map['Financial Debt']}"] = (
                f"=MAX({debt_amount_cell}-{repayment_cell}*{year_index + 1},0)"
            )
            ws_balance[f"{col}{balance_row_map['Total Liabilities']}"] = (
                f"={col}{balance_row_map['Financial Debt']}"
            )
            ws_balance[f"{col}{balance_row_map['Net Income']}"] = (
                f"='P&L'!{col}23"
            )
            ws_balance[f"{col}{balance_row_map['Dividends']}"] = "=0"
            ws_balance[f"{col}{balance_row_map['Equity at End of Year']}"] = (
                f"={col}{balance_row_map['Equity at Start of Year']}"
                f"+{col}{balance_row_map['Net Income']}"
                f"-{col}{balance_row_map['Dividends']}"
            )
            ws_balance[f"{col}{balance_row_map['Total Assets (Check)']}"] = (
                f"={col}{balance_row_map['Total Assets (Assets)']}"
            )
            ws_balance[f"{col}{balance_row_map['Total Liabilities + Equity']}"] = (
                f"={col}{balance_row_map['Total Liabilities']}"
                f"+{col}{balance_row_map['Equity at End of Year']}"
            )

        balance_notes_row = balance_table_start + len(balance_items) + 2
        ws_balance.cell(row=balance_notes_row, column=1, value="Notes")
        ws_balance.cell(
            row=balance_notes_row + 1,
            column=1,
            value="Fixed Assets end = prior assets + capex - depreciation.",
        )
        ws_balance.cell(
            row=balance_notes_row + 2,
            column=1,
            value="Equity end = equity start + net income - dividends.",
        )
        ws_balance.cell(
            row=balance_notes_row + 3,
            column=1,
            value="Total Assets should equal Total Liabilities + Equity.",
        )

        ws_valuation["A1"] = "Valuation Assumptions"
        valuation_assumption_rows = [
            ("Seller EBIT Multiple (x)", valuation_runtime["seller_ebit_multiple"]),
            ("Reference Year (0-4)", valuation_runtime["reference_year"]),
            ("Discount Rate (WACC)", valuation_runtime["buyer_discount_rate"]),
            ("Valuation Start Year (0-4)", valuation_runtime["valuation_start_year"]),
            ("Debt at Close (EUR)", valuation_runtime["debt_at_close_eur"]),
            ("Transaction Costs (% of EV)", valuation_runtime["transaction_cost_pct"]),
            ("Include Terminal Value (1=On)", 1 if valuation_runtime["include_terminal_value"] else 0),
        ]
        for idx, (label, value) in enumerate(valuation_assumption_rows, start=2):
            ws_valuation.cell(row=idx, column=1, value=label)
            ws_valuation.cell(row=idx, column=2, value=value)

        seller_table_start = 10
        ws_valuation.cell(row=seller_table_start, column=1, value="Seller Valuation (Multiple-Based)")
        ws_valuation.cell(row=seller_table_start + 1, column=1, value="Line Item")
        for idx, header in enumerate(year_headers, start=2):
            ws_valuation.cell(row=seller_table_start + 1, column=idx, value=header)

        seller_items = [
            "Reference Year EBIT",
            "Applied EBIT Multiple",
            "Enterprise Value (EV)",
            "Net Debt at Close",
            "Equity Value (Seller View)",
        ]
        for row_idx, item in enumerate(seller_items, start=seller_table_start + 2):
            ws_valuation.cell(row=row_idx, column=1, value=item)

        seller_row_map = {
            "Reference Year EBIT": seller_table_start + 2,
            "Applied EBIT Multiple": seller_table_start + 3,
            "Enterprise Value (EV)": seller_table_start + 4,
            "Net Debt at Close": seller_table_start + 5,
            "Equity Value (Seller View)": seller_table_start + 6,
        }
        seller_multiple_cell = "B2"
        reference_year_cell = "B3"

        for year_index in range(5):
            col = year_col(2 + year_index)
            is_ref_year = f"={reference_year_cell}={year_index}"
            ebit_cell = f"=INDEX('P&L'!B19:F19,1,{reference_year_cell}+1)"
            ws_valuation[f"{col}{seller_row_map['Reference Year EBIT']}"] = (
                f"=IF({is_ref_year},{ebit_cell},\"\")"
            )
            ws_valuation[f"{col}{seller_row_map['Applied EBIT Multiple']}"] = (
                f"=IF({is_ref_year},{seller_multiple_cell},\"\")"
            )
            ws_valuation[f"{col}{seller_row_map['Enterprise Value (EV)']}"] = (
                f"=IF({is_ref_year},{ebit_cell}*{seller_multiple_cell},\"\")"
            )
            ws_valuation[f"{col}{seller_row_map['Net Debt at Close']}"] = (
                f"=IF({year_index}=0,'Balance Sheet'!{col}14-'Balance Sheet'!{col}10,\"\")"
            )
            ws_valuation[f"{col}{seller_row_map['Equity Value (Seller View)']}"] = (
                f"=IF({is_ref_year},{col}{seller_row_map['Enterprise Value (EV)']}-'Balance Sheet'!B14+'Balance Sheet'!B10,\"\")"
            )

        buyer_table_start = seller_table_start + len(seller_items) + 4
        ws_valuation.cell(row=buyer_table_start, column=1, value="Buyer Valuation (DCF)")
        ws_valuation.cell(row=buyer_table_start + 1, column=1, value="Line Item")
        for idx, header in enumerate(year_headers, start=2):
            ws_valuation.cell(row=buyer_table_start + 1, column=idx, value=header)

        buyer_items = [
            "Free Cashflow",
            "Discount Factor",
            "Present Value of FCF",
            "Cumulative PV of FCF",
            "Terminal Value",
            "Enterprise Value (DCF)",
            "Net Debt at Close",
            "Transaction Costs",
            "Equity Value (Buyer View)",
        ]
        for row_idx, item in enumerate(buyer_items, start=buyer_table_start + 2):
            ws_valuation.cell(row=row_idx, column=1, value=item)

        buyer_row_map = {
            "Free Cashflow": buyer_table_start + 2,
            "Discount Factor": buyer_table_start + 3,
            "Present Value of FCF": buyer_table_start + 4,
            "Cumulative PV of FCF": buyer_table_start + 5,
            "Terminal Value": buyer_table_start + 6,
            "Enterprise Value (DCF)": buyer_table_start + 7,
            "Net Debt at Close": buyer_table_start + 8,
            "Transaction Costs": buyer_table_start + 9,
            "Equity Value (Buyer View)": buyer_table_start + 10,
        }

        discount_rate_cell = "B4"
        start_year_cell = "B5"
        debt_close_cell = "B6"
        tx_cost_cell = "B7"
        include_terminal_cell = "B8"

        for year_index in range(5):
            col = year_col(2 + year_index)
            prev_col = year_col(1 + year_index) if year_index > 0 else None
            ws_valuation[f"{col}{buyer_row_map['Free Cashflow']}"] = (
                f"=Cashflow!{col}{cashflow_row_map['Free Cashflow']}"
            )
            ws_valuation[f"{col}{buyer_row_map['Discount Factor']}"] = (
                f"=IF({year_index}>={start_year_cell},1/(1+{discount_rate_cell})^({year_index}-{start_year_cell}+1),0)"
            )
            ws_valuation[f"{col}{buyer_row_map['Present Value of FCF']}"] = (
                f"={col}{buyer_row_map['Free Cashflow']}*{col}{buyer_row_map['Discount Factor']}"
            )
            if prev_col:
                ws_valuation[f"{col}{buyer_row_map['Cumulative PV of FCF']}"] = (
                    f"={prev_col}{buyer_row_map['Cumulative PV of FCF']}+{col}{buyer_row_map['Present Value of FCF']}"
                )
            else:
                ws_valuation[f"{col}{buyer_row_map['Cumulative PV of FCF']}"] = (
                    f"={col}{buyer_row_map['Present Value of FCF']}"
                )
            if year_index == 4:
                ws_valuation[f"{col}{buyer_row_map['Terminal Value']}"] = (
                    f"=IF({include_terminal_cell}=1,{col}{buyer_row_map['Free Cashflow']}/{discount_rate_cell},\"\")"
                )
                ws_valuation[f"{col}{buyer_row_map['Enterprise Value (DCF)']}"] = (
                    f"={col}{buyer_row_map['Cumulative PV of FCF']}+IF({include_terminal_cell}=1,{col}{buyer_row_map['Terminal Value']}*{col}{buyer_row_map['Discount Factor']},0)"
                )
                ws_valuation[f"{col}{buyer_row_map['Net Debt at Close']}"] = (
                    f"={debt_close_cell}"
                )
                ws_valuation[f"{col}{buyer_row_map['Transaction Costs']}"] = (
                    f"={col}{buyer_row_map['Enterprise Value (DCF)']}*{tx_cost_cell}"
                )
                ws_valuation[f"{col}{buyer_row_map['Equity Value (Buyer View)']}"] = (
                    f"={col}{buyer_row_map['Enterprise Value (DCF)']}-{col}{buyer_row_map['Net Debt at Close']}-{col}{buyer_row_map['Transaction Costs']}"
                )

        bridge_start = buyer_table_start + len(buyer_items) + 4
        ws_valuation.cell(row=bridge_start, column=1, value="Purchase Price Bridge")
        ws_valuation.cell(row=bridge_start + 1, column=1, value="Line Item")
        ws_valuation.cell(row=bridge_start + 1, column=2, value="Year 0")

        bridge_items = [
            "Seller Equity Value",
            "Buyer Equity Value",
            "Valuation Gap (EUR)",
            "Valuation Gap (%)",
        ]
        for row_idx, item in enumerate(bridge_items, start=bridge_start + 2):
            ws_valuation.cell(row=row_idx, column=1, value=item)

        ws_valuation[f"B{bridge_start + 2}"] = (
            f"=INDEX(B{seller_row_map['Equity Value (Seller View)']}:F{seller_row_map['Equity Value (Seller View)']},1,{reference_year_cell}+1)"
        )
        ws_valuation[f"B{bridge_start + 3}"] = (
            f"=F{buyer_row_map['Equity Value (Buyer View)']}"
        )
        ws_valuation[f"B{bridge_start + 4}"] = (
            f"=B{bridge_start + 3}-B{bridge_start + 2}"
        )
        ws_valuation[f"B{bridge_start + 5}"] = (
            f"=IF(B{bridge_start + 2}=0,0,B{bridge_start + 4}/B{bridge_start + 2})"
        )

        valuation_notes_row = bridge_start + len(bridge_items) + 2
        ws_valuation.cell(row=valuation_notes_row, column=1, value="Notes")
        ws_valuation.cell(
            row=valuation_notes_row + 1,
            column=1,
            value="Seller EV = EBIT (reference year) × multiple.",
        )
        ws_valuation.cell(
            row=valuation_notes_row + 2,
            column=1,
            value="DCF uses free cashflow discounted at the buyer rate.",
        )
        ws_valuation.cell(
            row=valuation_notes_row + 3,
            column=1,
            value="Equity value = EV - net debt - transaction costs.",
        )

        ws_financing["A1"] = "Financing Assumptions"
        financing_rows = [
            ("Initial Debt Amount (EUR)", financing_assumptions["initial_debt_eur"]),
            ("Interest Rate", financing_assumptions["interest_rate_pct"]),
            ("Amortisation Type", financing_assumptions["amortization_type"]),
            ("Amortisation Period (Years)", financing_assumptions["amortization_period_years"]),
            ("Grace Period (Years)", financing_assumptions["grace_period_years"]),
            ("Special Repayment Year", financing_assumptions["special_repayment_year"] if financing_assumptions["special_repayment_year"] is not None else -1),
            ("Special Repayment Amount (EUR)", financing_assumptions["special_repayment_amount_eur"]),
            ("Minimum DSCR", financing_assumptions["minimum_dscr"]),
            ("Minimum Cash Balance (EUR)", financing_assumptions["minimum_cash_balance_eur"]),
            ("Target IRR", financing_assumptions["target_irr"]),
            ("Max Equity Contribution (EUR)", financing_assumptions["max_equity_contribution_eur"] or 0),
            ("Minimum Cash Yield", financing_assumptions["min_cash_yield"]),
        ]
        for idx, (label, value) in enumerate(financing_rows, start=2):
            ws_financing.cell(row=idx, column=1, value=label)
            ws_financing.cell(row=idx, column=2, value=value)

        sources_start = 12
        ws_financing.cell(row=sources_start, column=1, value="Sources & Uses")
        ws_financing.cell(row=sources_start + 1, column=1, value="Line Item")
        ws_financing.cell(row=sources_start + 1, column=2, value="Amount")

        sources_items = [
            "USES",
            "Purchase Price",
            "Transaction Fees",
            "Refinancing",
            "Minimum Cash at Close",
            "Total Uses",
            "SOURCES",
            "Senior Debt",
            "Equity Contribution",
            "Total Sources",
            "Sources - Uses",
        ]
        for row_idx, item in enumerate(sources_items, start=sources_start + 2):
            ws_financing.cell(row=row_idx, column=1, value=item)

        sources_row_map = {
            item: sources_start + 2 + idx
            for idx, item in enumerate(sources_items)
        }
        purchase_price_cell = assumption_cells["Purchase Price (EUR)"]
        equity_cell = assumption_cells["Equity Contribution (EUR)"]
        tx_cost_pct_cell = "Valuation!B7"
        initial_debt_cell = "B2"
        min_cash_cell = "B10"

        ws_financing[f"B{sources_row_map['Purchase Price']}"] = (
            f"={purchase_price_cell}"
        )
        ws_financing[f"B{sources_row_map['Transaction Fees']}"] = (
            f"={purchase_price_cell}*{tx_cost_pct_cell}"
        )
        ws_financing[f"B{sources_row_map['Refinancing']}"] = "=0"
        ws_financing[f"B{sources_row_map['Minimum Cash at Close']}"] = (
            f"={min_cash_cell}"
        )
        ws_financing[f"B{sources_row_map['Total Uses']}"] = (
            f"=B{sources_row_map['Purchase Price']}"
            f"+B{sources_row_map['Transaction Fees']}"
            f"+B{sources_row_map['Refinancing']}"
            f"+B{sources_row_map['Minimum Cash at Close']}"
        )
        ws_financing[f"B{sources_row_map['Senior Debt']}"] = (
            f"={initial_debt_cell}"
        )
        ws_financing[f"B{sources_row_map['Equity Contribution']}"] = (
            f"={equity_cell}"
        )
        ws_financing[f"B{sources_row_map['Total Sources']}"] = (
            f"=B{sources_row_map['Senior Debt']}+B{sources_row_map['Equity Contribution']}"
        )
        ws_financing[f"B{sources_row_map['Sources - Uses']}"] = (
            f"=B{sources_row_map['Total Sources']}-B{sources_row_map['Total Uses']}"
        )

        debt_table_start = sources_start + len(sources_items) + 3
        ws_financing.cell(row=debt_table_start, column=1, value="Debt Schedule")
        ws_financing.cell(row=debt_table_start + 1, column=1, value="Line Item")
        for idx, header in enumerate(year_headers, start=2):
            ws_financing.cell(row=debt_table_start + 1, column=idx, value=header)

        debt_items = [
            "Opening Debt",
            "Debt Drawdown",
            "Scheduled Repayment",
            "Special Repayment",
            "Total Repayment",
            "Closing Debt",
            "Interest Expense",
        ]
        for row_idx, item in enumerate(debt_items, start=debt_table_start + 2):
            ws_financing.cell(row=row_idx, column=1, value=item)

        debt_row_map = {
            "Opening Debt": debt_table_start + 2,
            "Debt Drawdown": debt_table_start + 3,
            "Scheduled Repayment": debt_table_start + 4,
            "Special Repayment": debt_table_start + 5,
            "Total Repayment": debt_table_start + 6,
            "Closing Debt": debt_table_start + 7,
            "Interest Expense": debt_table_start + 8,
        }

        interest_rate_cell = "B3"
        amort_type_cell = "B4"
        amort_period_cell = "B5"
        grace_period_cell = "B6"
        special_year_cell = "B7"
        special_amount_cell = "B8"

        for year_index in range(5):
            col = year_col(2 + year_index)
            prev_col = year_col(1 + year_index) if year_index > 0 else None
            ws_financing[f"{col}{debt_row_map['Opening Debt']}"] = (
                f"={initial_debt_cell}" if year_index == 0 else f"={prev_col}{debt_row_map['Closing Debt']}"
            )
            ws_financing[f"{col}{debt_row_map['Debt Drawdown']}"] = (
                f"=IF({year_index}=0,{initial_debt_cell},0)"
            )
            ws_financing[f"{col}{debt_row_map['Scheduled Repayment']}"] = (
                f"=IF({amort_type_cell}=\"Bullet\","
                f"IF({year_index}={amort_period_cell}-1,{initial_debt_cell},0),"
                f"IF({year_index}<{grace_period_cell},0,"
                f"IF({year_index}<{amort_period_cell},{initial_debt_cell}/{amort_period_cell},0)))"
            )
            ws_financing[f"{col}{debt_row_map['Special Repayment']}"] = (
                f"=IF({special_year_cell}={year_index},{special_amount_cell},0)"
            )
            ws_financing[f"{col}{debt_row_map['Total Repayment']}"] = (
                f"=MIN({col}{debt_row_map['Opening Debt']},{col}{debt_row_map['Scheduled Repayment']}+{col}{debt_row_map['Special Repayment']})"
            )
            ws_financing[f"{col}{debt_row_map['Closing Debt']}"] = (
                f"={col}{debt_row_map['Opening Debt']}-{col}{debt_row_map['Total Repayment']}"
            )
            ws_financing[f"{col}{debt_row_map['Interest Expense']}"] = (
                f"={col}{debt_row_map['Opening Debt']}*{interest_rate_cell}"
            )

        service_table_start = debt_table_start + len(debt_items) + 3
        ws_financing.cell(row=service_table_start, column=1, value="Debt Service & Covenants")
        ws_financing.cell(row=service_table_start + 1, column=1, value="Line Item")
        for idx, header in enumerate(year_headers, start=2):
            ws_financing.cell(row=service_table_start + 1, column=idx, value=header)

        service_items = [
            "EBITDA",
            "Operating Cashflow",
            "CFADS",
            "Debt Service",
            "DSCR",
            "Minimum Required DSCR",
            "Covenant Breach",
        ]
        for row_idx, item in enumerate(service_items, start=service_table_start + 2):
            ws_financing.cell(row=row_idx, column=1, value=item)

        service_row_map = {
            "EBITDA": service_table_start + 2,
            "Operating Cashflow": service_table_start + 3,
            "CFADS": service_table_start + 4,
            "Debt Service": service_table_start + 5,
            "DSCR": service_table_start + 6,
            "Minimum Required DSCR": service_table_start + 7,
            "Covenant Breach": service_table_start + 8,
        }
        min_dscr_cell = "B9"

        for year_index in range(5):
            col = year_col(2 + year_index)
            ws_financing[f"{col}{service_row_map['EBITDA']}"] = f"='P&L'!{col}17"
            ws_financing[f"{col}{service_row_map['Operating Cashflow']}"] = (
                f"=Cashflow!{col}{cashflow_row_map['Operating Cashflow']}"
            )
            ws_financing[f"{col}{service_row_map['CFADS']}"] = (
                f"=Cashflow!{col}{cashflow_row_map['Operating Cashflow']}"
                f"-Cashflow!{col}{cashflow_row_map['Capex']}"
            )
            ws_financing[f"{col}{service_row_map['Debt Service']}"] = (
                f"={col}{debt_row_map['Interest Expense']}+{col}{debt_row_map['Total Repayment']}"
            )
            ws_financing[f"{col}{service_row_map['DSCR']}"] = (
                f"=IF({col}{service_row_map['Debt Service']}=0,0,{col}{service_row_map['CFADS']}/{col}{service_row_map['Debt Service']})"
            )
            ws_financing[f"{col}{service_row_map['Minimum Required DSCR']}"] = (
                f"={min_dscr_cell}"
            )
            ws_financing[f"{col}{service_row_map['Covenant Breach']}"] = (
                f"=IF({col}{service_row_map['DSCR']}<{min_dscr_cell},\"YES\",\"NO\")"
            )

        investor_table_start = service_table_start + len(service_items) + 3
        ws_financing.cell(
            row=investor_table_start,
            column=1,
            value="Investor Cashflow Bridge",
        )
        ws_financing.cell(
            row=investor_table_start + 1,
            column=1,
            value="Line Item",
        )
        for idx, header in enumerate(year_headers, start=2):
            ws_financing.cell(
                row=investor_table_start + 1, column=idx, value=header
            )

        investor_items = [
            "Free Cashflow",
            "Debt Service",
            "Mandatory Cash Retention",
            "Cash to Equity",
        ]
        for row_idx, item in enumerate(
            investor_items, start=investor_table_start + 2
        ):
            ws_financing.cell(row=row_idx, column=1, value=item)

        investor_row_map = {
            "Free Cashflow": investor_table_start + 2,
            "Debt Service": investor_table_start + 3,
            "Mandatory Cash Retention": investor_table_start + 4,
            "Cash to Equity": investor_table_start + 5,
        }

        for year_index in range(5):
            col = year_col(2 + year_index)
            ws_financing[f"{col}{investor_row_map['Free Cashflow']}"] = (
                f"=Cashflow!{col}{cashflow_row_map['Free Cashflow']}"
            )
            ws_financing[f"{col}{investor_row_map['Debt Service']}"] = (
                f"={col}{debt_row_map['Interest Expense']}"
                f"+{col}{debt_row_map['Total Repayment']}"
            )
            ws_financing[
                f"{col}{investor_row_map['Cash to Equity']}"
            ] = f"=Cashflow!{col}{cashflow_row_map['Net Cashflow (Liquidity)']}"
            ws_financing[
                f"{col}{investor_row_map['Mandatory Cash Retention']}"
            ] = (
                f"={col}{investor_row_map['Free Cashflow']}"
                f"-{col}{investor_row_map['Debt Service']}"
                f"-{col}{investor_row_map['Cash to Equity']}"
            )

        investor_kpi_start = investor_table_start + len(investor_items) + 3
        ws_financing.cell(
            row=investor_kpi_start, column=1, value="Investor KPIs"
        )
        ws_financing.cell(
            row=investor_kpi_start + 1, column=1, value="Metric"
        )
        ws_financing.cell(
            row=investor_kpi_start + 1, column=2, value="Value"
        )

        investor_kpis = [
            "Equity Contribution",
            "Target IRR",
            "Achieved IRR",
            "Cash-on-Cash",
            "Average Cash Yield",
        ]
        for row_idx, item in enumerate(
            investor_kpis, start=investor_kpi_start + 2
        ):
            ws_financing.cell(row=row_idx, column=1, value=item)

        irr_cashflows_start = investor_kpi_start + len(investor_kpis) + 3
        ws_financing.cell(
            row=irr_cashflows_start, column=1, value="IRR Cashflows"
        )
        irr_cashflow_rows = [
            "Initial Equity",
            "Year 0 Cash to Equity",
            "Year 1 Cash to Equity",
            "Year 2 Cash to Equity",
            "Year 3 Cash to Equity",
            "Year 4 Cash to Equity",
        ]
        for row_idx, item in enumerate(
            irr_cashflow_rows, start=irr_cashflows_start + 1
        ):
            ws_financing.cell(row=row_idx, column=1, value=item)

        ws_financing[f"B{irr_cashflows_start + 1}"] = f"=-{equity_cell}"
        for year_index in range(5):
            ws_financing[f"B{irr_cashflows_start + 2 + year_index}"] = (
                f"=Cashflow!{year_col(2 + year_index)}{cashflow_row_map['Net Cashflow (Liquidity)']}"
            )

        target_irr_cell = "B11"
        ws_financing[f"B{investor_kpi_start + 2}"] = f"={equity_cell}"
        ws_financing[f"B{investor_kpi_start + 3}"] = f"={target_irr_cell}"
        ws_financing[f"B{investor_kpi_start + 4}"] = (
            f"=IRR(B{irr_cashflows_start + 1}:B{irr_cashflows_start + 6})"
        )
        ws_financing[f"B{investor_kpi_start + 5}"] = (
            f"=IF({equity_cell}=0,0,"
            f"SUMIF(B{irr_cashflows_start + 2}:B{irr_cashflows_start + 6},\">0\")/"
            f"{equity_cell})"
        )
        ws_financing[f"B{investor_kpi_start + 6}"] = (
            f"=IF({equity_cell}=0,0,"
            f"(SUMIF(B{irr_cashflows_start + 2}:B{irr_cashflows_start + 6},\">0\")/5)/"
            f"{equity_cell})"
        )

        financing_notes_row = irr_cashflows_start + len(irr_cashflow_rows) + 2
        ws_financing.cell(row=financing_notes_row, column=1, value="Notes")
        ws_financing.cell(
            row=financing_notes_row + 1,
            column=1,
            value="Sources & uses must reconcile: Total Sources - Total Uses = 0.",
        )
        ws_financing.cell(
            row=financing_notes_row + 2,
            column=1,
            value="Interest = Opening Debt × Interest Rate.",
        )
        ws_financing.cell(
            row=financing_notes_row + 3,
            column=1,
            value="DSCR = CFADS / Debt Service.",
        )
        ws_financing.cell(
            row=financing_notes_row + 4,
            column=1,
            value="CFADS = Operating Cashflow - Maintenance Capex.",
        )

        ws_financing_notes["A1"] = "Financing Notes"
        ws_financing_notes["A2"] = (
            "Sources & uses reconcile purchase price, fees, and minimum cash at close."
        )
        ws_financing_notes["A3"] = (
            "Debt service is based on opening debt, scheduled amortisation, and interest."
        )
        ws_financing_notes["A4"] = (
            "CFADS measures cash available for debt service after maintenance capex."
        )
        ws_financing_notes["A5"] = (
            "Investor cashflows reflect remaining cash after debt service."
        )

        writer.close()

    output.seek(0)
    return output


def run_app():
    st.title("Financial Model")

    base_model = create_demo_input_model()
    st.session_state.setdefault("edit_pnl_assumptions", True)
    st.session_state.setdefault("edit_cashflow_assumptions", False)
    st.session_state.setdefault("edit_balance_sheet_assumptions", False)
    st.session_state.setdefault("edit_valuation_assumptions", False)
    st.session_state.setdefault("edit_financing_assumptions", False)
    if not st.session_state.get("defaults_initialized"):
        _seed_session_defaults(base_model)
        st.session_state["defaults_initialized"] = True

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
        if page == "Operating Model (P&L)":
            st.markdown("## Quick Assumptions")
            st.session_state["edit_pnl_assumptions"] = True
            if st.session_state.get("edit_pnl_assumptions"):
                scenario_options = ["Base", "Best", "Worst"]
                selected_scenario = st.session_state.get(
                    "scenario_selection.selected_scenario",
                    base_model.scenario_selection["selected_scenario"].value,
                )
                scenario_index = (
                    scenario_options.index(selected_scenario)
                    if selected_scenario in scenario_options
                    else 0
                )
                scenario_key = selected_scenario.lower()
                utilization_field = _get_field_by_path(
                    base_model.__dict__,
                    ["scenario_parameters", "utilization_rate", scenario_key],
                )
                day_rate_field = _get_field_by_path(
                    base_model.__dict__,
                    ["scenario_parameters", "day_rate_eur", scenario_key],
                )
                fte_field = _get_field_by_path(
                    base_model.__dict__,
                    ["operating_assumptions", "consulting_fte_start"],
                )
                work_days_field = _get_field_by_path(
                    base_model.__dict__,
                    ["operating_assumptions", "work_days_per_year"],
                )
                day_rate_growth_field = _get_field_by_path(
                    base_model.__dict__,
                    ["operating_assumptions", "day_rate_growth_pct"],
                )
                guarantee_y1_field = _get_field_by_path(
                    base_model.__dict__,
                    ["operating_assumptions", "revenue_guarantee_pct_year_1"],
                )
                guarantee_y2_field = _get_field_by_path(
                    base_model.__dict__,
                    ["operating_assumptions", "revenue_guarantee_pct_year_2"],
                )
                guarantee_y3_field = _get_field_by_path(
                    base_model.__dict__,
                    ["operating_assumptions", "revenue_guarantee_pct_year_3"],
                )
                base_cost_field = _get_field_by_path(
                    base_model.__dict__,
                    ["personnel_cost_assumptions", "avg_consultant_base_cost_eur_per_year"],
                )
                bonus_field = _get_field_by_path(
                    base_model.__dict__,
                    ["personnel_cost_assumptions", "bonus_pct_of_base"],
                )
                payroll_field = _get_field_by_path(
                    base_model.__dict__,
                    ["personnel_cost_assumptions", "payroll_burden_pct_of_comp"],
                )
                wage_inflation_field = _get_field_by_path(
                    base_model.__dict__,
                    ["personnel_cost_assumptions", "wage_inflation_pct"],
                )
                legal_field = _get_field_by_path(
                    base_model.__dict__,
                    ["overhead_and_variable_costs", "legal_audit_eur_per_year"],
                )
                it_field = _get_field_by_path(
                    base_model.__dict__,
                    ["overhead_and_variable_costs", "it_and_software_eur_per_year"],
                )
                rent_field = _get_field_by_path(
                    base_model.__dict__,
                    ["overhead_and_variable_costs", "rent_eur_per_year"],
                )
                other_field = _get_field_by_path(
                    base_model.__dict__,
                    ["overhead_and_variable_costs", "other_overhead_eur_per_year"],
                )
                insurance_field = _get_field_by_path(
                    base_model.__dict__,
                    ["overhead_and_variable_costs", "insurance_eur_per_year"],
                )
                overhead_inflation_field = _get_field_by_path(
                    base_model.__dict__,
                    ["overhead_and_variable_costs", "overhead_inflation_pct"],
                )
                utilization_defaults = st.session_state.get(
                    "utilization_by_year",
                    [utilization_field.value] * 5,
                )
                pnl_controls = [
                    {
                        "type": "select",
                        "label": "Scenario",
                        "options": scenario_options,
                        "index": scenario_index,
                        "field_key": "scenario_selection.selected_scenario",
                    },
                    {
                        "type": "int",
                        "label": "Consulting FTE",
                        "field_key": "operating_assumptions.consulting_fte_start",
                        "value": _get_current_value(
                            "operating_assumptions.consulting_fte_start",
                            fte_field.value,
                        ),
                    },
                    {
                        "type": "int",
                        "label": "Workdays per Year",
                        "field_key": "operating_assumptions.work_days_per_year",
                        "value": _get_current_value(
                            "operating_assumptions.work_days_per_year",
                            work_days_field.value,
                        ),
                    },
                    {
                        "type": "pct",
                        "label": "Utilization Year 0 (%)",
                        "field_key": "utilization_by_year.0",
                        "value": _get_current_value(
                            "utilization_by_year.0",
                            utilization_defaults[0],
                        ),
                    },
                    {
                        "type": "pct",
                        "label": "Utilization Year 1 (%)",
                        "field_key": "utilization_by_year.1",
                        "value": _get_current_value(
                            "utilization_by_year.1",
                            utilization_defaults[1],
                        ),
                    },
                    {
                        "type": "pct",
                        "label": "Utilization Year 2 (%)",
                        "field_key": "utilization_by_year.2",
                        "value": _get_current_value(
                            "utilization_by_year.2",
                            utilization_defaults[2],
                        ),
                    },
                    {
                        "type": "pct",
                        "label": "Utilization Year 3 (%)",
                        "field_key": "utilization_by_year.3",
                        "value": _get_current_value(
                            "utilization_by_year.3",
                            utilization_defaults[3],
                        ),
                    },
                    {
                        "type": "pct",
                        "label": "Utilization Year 4 (%)",
                        "field_key": "utilization_by_year.4",
                        "value": _get_current_value(
                            "utilization_by_year.4",
                            utilization_defaults[4],
                        ),
                    },
                    {
                        "type": "number",
                        "label": "Day Rate (EUR)",
                        "field_key": f"scenario_parameters.day_rate_eur.{scenario_key}",
                        "value": _get_current_value(
                            f"scenario_parameters.day_rate_eur.{scenario_key}",
                            day_rate_field.value,
                        ),
                        "step": 100.0,
                        "format": "%.0f",
                    },
                    {
                        "type": "pct",
                        "label": "Day Rate Growth (%)",
                        "field_key": "operating_assumptions.day_rate_growth_pct",
                        "value": _get_current_value(
                            "operating_assumptions.day_rate_growth_pct",
                            day_rate_growth_field.value,
                        ),
                    },
                    {
                        "type": "pct",
                        "label": "Guarantee Year 1 (%)",
                        "field_key": "operating_assumptions.revenue_guarantee_pct_year_1",
                        "value": _get_current_value(
                            "operating_assumptions.revenue_guarantee_pct_year_1",
                            guarantee_y1_field.value,
                        ),
                    },
                    {
                        "type": "pct",
                        "label": "Guarantee Year 2 (%)",
                        "field_key": "operating_assumptions.revenue_guarantee_pct_year_2",
                        "value": _get_current_value(
                            "operating_assumptions.revenue_guarantee_pct_year_2",
                            guarantee_y2_field.value,
                        ),
                    },
                    {
                        "type": "pct",
                        "label": "Guarantee Year 3 (%)",
                        "field_key": "operating_assumptions.revenue_guarantee_pct_year_3",
                        "value": _get_current_value(
                            "operating_assumptions.revenue_guarantee_pct_year_3",
                            guarantee_y3_field.value,
                        ),
                    },
                    {
                        "type": "number",
                        "label": "Consultant Base Cost (EUR)",
                        "field_key": "personnel_cost_assumptions.avg_consultant_base_cost_eur_per_year",
                        "value": _get_current_value(
                            "personnel_cost_assumptions.avg_consultant_base_cost_eur_per_year",
                            base_cost_field.value,
                        ),
                        "step": 1000.0,
                        "format": "%.0f",
                    },
                    {
                        "type": "pct",
                        "label": "Bonus (%)",
                        "field_key": "personnel_cost_assumptions.bonus_pct_of_base",
                        "value": _get_current_value(
                            "personnel_cost_assumptions.bonus_pct_of_base",
                            bonus_field.value,
                        ),
                    },
                    {
                        "type": "pct",
                        "label": "Payroll Burden (%)",
                        "field_key": "personnel_cost_assumptions.payroll_burden_pct_of_comp",
                        "value": _get_current_value(
                            "personnel_cost_assumptions.payroll_burden_pct_of_comp",
                            payroll_field.value,
                        ),
                    },
                    {
                        "type": "pct",
                        "label": "Wage Inflation (%)",
                        "field_key": "personnel_cost_assumptions.wage_inflation_pct",
                        "value": _get_current_value(
                            "personnel_cost_assumptions.wage_inflation_pct",
                            wage_inflation_field.value,
                        ),
                    },
                    {
                        "type": "number",
                        "label": "External Advisors (EUR)",
                        "field_key": "overhead_and_variable_costs.legal_audit_eur_per_year",
                        "value": _get_current_value(
                            "overhead_and_variable_costs.legal_audit_eur_per_year",
                            legal_field.value,
                        ),
                        "step": 10000.0,
                        "format": "%.0f",
                    },
                    {
                        "type": "number",
                        "label": "IT (EUR)",
                        "field_key": "overhead_and_variable_costs.it_and_software_eur_per_year",
                        "value": _get_current_value(
                            "overhead_and_variable_costs.it_and_software_eur_per_year",
                            it_field.value,
                        ),
                        "step": 10000.0,
                        "format": "%.0f",
                    },
                    {
                        "type": "number",
                        "label": "Office (EUR)",
                        "field_key": "overhead_and_variable_costs.rent_eur_per_year",
                        "value": _get_current_value(
                            "overhead_and_variable_costs.rent_eur_per_year",
                            rent_field.value,
                        ),
                        "step": 10000.0,
                        "format": "%.0f",
                    },
                    {
                        "type": "number",
                        "label": "Other Services (EUR)",
                        "field_key": "overhead_and_variable_costs.other_overhead_eur_per_year",
                        "value": _get_current_value(
                            "overhead_and_variable_costs.other_overhead_eur_per_year",
                            other_field.value,
                        ),
                        "step": 10000.0,
                        "format": "%.0f",
                    },
                    {
                        "type": "number",
                        "label": "Insurance (EUR)",
                        "field_key": "overhead_and_variable_costs.insurance_eur_per_year",
                        "value": _get_current_value(
                            "overhead_and_variable_costs.insurance_eur_per_year",
                            insurance_field.value,
                        ),
                        "step": 10000.0,
                        "format": "%.0f",
                    },
                    {
                        "type": "pct",
                        "label": "Overhead Inflation (%)",
                        "field_key": "overhead_and_variable_costs.overhead_inflation_pct",
                        "value": _get_current_value(
                            "overhead_and_variable_costs.overhead_inflation_pct",
                            overhead_inflation_field.value,
                        ),
                    },
                ]
                _render_inline_controls("P&L Drivers", pnl_controls, columns=1)
                st.session_state["utilization_by_year"] = [
                    st.session_state.get(
                        f"utilization_by_year.{year_index}",
                        utilization_defaults[year_index],
                    )
                    for year_index in range(5)
                ]
        if page == "Cashflow & Liquidity" and st.session_state.get(
            "edit_cashflow_assumptions"
        ):
            st.markdown("## Cashflow Assumptions")
            cashflow_defaults = _default_cashflow_assumptions()
            cashflow_controls = [
                {
                    "type": "pct",
                    "label": "Tax Cash Rate (%)",
                    "field_key": "cashflow.tax_cash_rate_pct",
                    "value": _get_current_value(
                        "cashflow.tax_cash_rate_pct",
                        cashflow_defaults["tax_cash_rate_pct"],
                    ),
                },
                {
                    "type": "select",
                    "label": "Tax Payment Lag (Years)",
                    "options": [0, 1],
                    "index": [0, 1].index(
                        _get_current_value(
                            "cashflow.tax_payment_lag_years",
                            cashflow_defaults["tax_payment_lag_years"],
                        )
                    ),
                    "field_key": "cashflow.tax_payment_lag_years",
                },
                {
                    "type": "pct",
                    "label": "Capex (% of Revenue)",
                    "field_key": "cashflow.capex_pct_revenue",
                    "value": _get_current_value(
                        "cashflow.capex_pct_revenue",
                        cashflow_defaults["capex_pct_revenue"],
                    ),
                },
                {
                    "type": "pct",
                    "label": "Working Capital Adjustment (% of Revenue)",
                    "field_key": "cashflow.working_capital_pct_revenue",
                    "value": _get_current_value(
                        "cashflow.working_capital_pct_revenue",
                        cashflow_defaults["working_capital_pct_revenue"],
                    ),
                },
                {
                    "type": "number",
                    "label": "Opening Cash Balance (EUR)",
                    "field_key": "cashflow.opening_cash_balance_eur",
                    "value": _get_current_value(
                        "cashflow.opening_cash_balance_eur",
                        cashflow_defaults["opening_cash_balance_eur"],
                    ),
                    "step": 50000.0,
                    "format": "%.0f",
                },
            ]
            _render_inline_controls("Cashflow Drivers", cashflow_controls, columns=1)
        if page == "Balance Sheet" and st.session_state.get(
            "edit_balance_sheet_assumptions"
        ):
            st.markdown("## Balance Sheet Assumptions")
            balance_defaults = _default_balance_sheet_assumptions(base_model)
            balance_controls = [
                {
                    "type": "number",
                    "label": "Opening Equity (EUR)",
                    "field_key": "balance_sheet.opening_equity_eur",
                    "value": _get_current_value(
                        "balance_sheet.opening_equity_eur",
                        balance_defaults["opening_equity_eur"],
                    ),
                    "step": 100000.0,
                    "format": "%.0f",
                },
                {
                    "type": "pct",
                    "label": "Depreciation Rate (%)",
                    "field_key": "balance_sheet.depreciation_rate_pct",
                    "value": _get_current_value(
                        "balance_sheet.depreciation_rate_pct",
                        balance_defaults["depreciation_rate_pct"],
                    ),
                },
                {
                    "type": "number",
                    "label": "Minimum Cash Balance (EUR)",
                    "field_key": "balance_sheet.minimum_cash_balance_eur",
                    "value": _get_current_value(
                        "balance_sheet.minimum_cash_balance_eur",
                        balance_defaults["minimum_cash_balance_eur"],
                    ),
                    "step": 50000.0,
                    "format": "%.0f",
                },
            ]
            _render_inline_controls("Balance Sheet Drivers", balance_controls, columns=1)
        if page == "Valuation & Purchase Price" and st.session_state.get(
            "edit_valuation_assumptions"
        ):
            st.markdown("## Valuation Assumptions")
            valuation_defaults = _default_valuation_assumptions(base_model)
            valuation_controls = [
                {
                    "type": "number",
                    "label": "EBIT Multiple (x)",
                    "field_key": "valuation.seller_ebit_multiple",
                    "value": _get_current_value(
                        "valuation.seller_ebit_multiple",
                        valuation_defaults["seller_ebit_multiple"],
                    ),
                    "step": 0.1,
                    "format": "%.2f",
                },
                {
                    "type": "select",
                    "label": "Reference Year for Multiple",
                    "options": ["Year 0", "Year 1", "Year 2", "Year 3", "Year 4"],
                    "index": _get_current_value(
                        "valuation.reference_year",
                        valuation_defaults["reference_year"],
                    ),
                    "field_key": "valuation.reference_year",
                },
                {
                    "type": "pct",
                    "label": "Discount Rate (WACC)",
                    "field_key": "valuation.buyer_discount_rate",
                    "value": _get_current_value(
                        "valuation.buyer_discount_rate",
                        valuation_defaults["buyer_discount_rate"],
                    ),
                },
                {
                    "type": "select",
                    "label": "Valuation Start Year",
                    "options": ["Year 0", "Year 1", "Year 2", "Year 3", "Year 4"],
                    "index": _get_current_value(
                        "valuation.valuation_start_year",
                        valuation_defaults["valuation_start_year"],
                    ),
                    "field_key": "valuation.valuation_start_year",
                },
                {
                    "type": "number",
                    "label": "Debt at Close (EUR)",
                    "field_key": "valuation.debt_at_close_eur",
                    "value": _get_current_value(
                        "valuation.debt_at_close_eur",
                        valuation_defaults["debt_at_close_eur"],
                    ),
                    "step": 100000.0,
                    "format": "%.0f",
                },
                {
                    "type": "pct",
                    "label": "Transaction Costs (% of EV)",
                    "field_key": "valuation.transaction_cost_pct",
                    "value": _get_current_value(
                        "valuation.transaction_cost_pct",
                        valuation_defaults["transaction_cost_pct"],
                    ),
                },
                {
                    "type": "select",
                    "label": "Include Terminal Value",
                    "options": ["Off", "On"],
                    "index": 1
                    if _get_current_value(
                        "valuation.include_terminal_value",
                        valuation_defaults["include_terminal_value"],
                    )
                    else 0,
                    "field_key": "valuation.include_terminal_value",
                },
            ]
            _render_inline_controls("Valuation Drivers", valuation_controls, columns=1)
        if page == "Financing & Debt" and st.session_state.get(
            "edit_financing_assumptions"
        ):
            st.markdown("## Financing Assumptions")
            financing_defaults = _default_financing_assumptions(base_model)
            preset_options = ["Custom", "Aggressive", "Base (Bankable)", "Conservative"]
            selected_preset = st.selectbox(
                "Preset",
                preset_options,
                index=preset_options.index(
                    st.session_state.get("financing.preset", "Custom")
                ),
            )
            if selected_preset != st.session_state.get("financing.preset"):
                st.session_state["financing.preset"] = selected_preset
                if selected_preset == "Aggressive":
                    st.session_state["financing.initial_debt_eur"] = (
                        financing_defaults["initial_debt_eur"] * 1.2
                    )
                    st.session_state["financing.amortization_type"] = "Bullet"
                    st.session_state["financing.amortization_period_years"] = 5
                    st.session_state["financing.grace_period_years"] = 2
                elif selected_preset == "Base (Bankable)":
                    st.session_state["financing.initial_debt_eur"] = (
                        financing_defaults["initial_debt_eur"]
                    )
                    st.session_state["financing.amortization_type"] = "Linear"
                    st.session_state["financing.amortization_period_years"] = 5
                    st.session_state["financing.grace_period_years"] = 1
                elif selected_preset == "Conservative":
                    st.session_state["financing.initial_debt_eur"] = (
                        financing_defaults["initial_debt_eur"] * 0.75
                    )
                    st.session_state["financing.amortization_type"] = "Linear"
                    st.session_state["financing.amortization_period_years"] = 4
                    st.session_state["financing.grace_period_years"] = 0

            financing_controls = [
                {
                    "type": "number",
                    "label": "Initial Debt Amount at Close (EUR)",
                    "field_key": "financing.initial_debt_eur",
                    "value": _get_current_value(
                        "financing.initial_debt_eur",
                        financing_defaults["initial_debt_eur"],
                    ),
                    "step": 100000.0,
                    "format": "%.0f",
                },
                {
                    "type": "pct",
                    "label": "Interest Rate (% fixed)",
                    "field_key": "financing.interest_rate_pct",
                    "value": _get_current_value(
                        "financing.interest_rate_pct",
                        financing_defaults["interest_rate_pct"],
                    ),
                },
                {
                    "type": "select",
                    "label": "Amortisation Type",
                    "options": ["Linear", "Bullet"],
                    "index": 0
                    if _get_current_value(
                        "financing.amortization_type",
                        financing_defaults["amortization_type"],
                    )
                    == "Linear"
                    else 1,
                    "field_key": "financing.amortization_type",
                },
                {
                    "type": "int",
                    "label": "Amortisation Period (years)",
                    "field_key": "financing.amortization_period_years",
                    "value": _get_current_value(
                        "financing.amortization_period_years",
                        financing_defaults["amortization_period_years"],
                    ),
                },
                {
                    "type": "int",
                    "label": "Grace Period (years)",
                    "field_key": "financing.grace_period_years",
                    "value": _get_current_value(
                        "financing.grace_period_years",
                        financing_defaults["grace_period_years"],
                    ),
                },
                {
                    "type": "select",
                    "label": "Special Repayment Year",
                    "options": [
                        "None",
                        "Year 0",
                        "Year 1",
                        "Year 2",
                        "Year 3",
                        "Year 4",
                    ],
                    "index": [
                        "None",
                        "Year 0",
                        "Year 1",
                        "Year 2",
                        "Year 3",
                        "Year 4",
                    ].index(
                        "None"
                        if _get_current_value(
                            "financing.special_repayment_year",
                            financing_defaults["special_repayment_year"],
                        )
                        is None
                        else f"Year {_get_current_value('financing.special_repayment_year', financing_defaults['special_repayment_year'])}"
                    ),
                    "field_key": "financing.special_repayment_year",
                },
                {
                    "type": "number",
                    "label": "Special Repayment Amount (EUR)",
                    "field_key": "financing.special_repayment_amount_eur",
                    "value": _get_current_value(
                        "financing.special_repayment_amount_eur",
                        financing_defaults["special_repayment_amount_eur"],
                    ),
                    "step": 100000.0,
                    "format": "%.0f",
                },
                {
                    "type": "number",
                    "label": "Minimum DSCR",
                    "field_key": "financing.minimum_dscr",
                    "value": _get_current_value(
                        "financing.minimum_dscr",
                        financing_defaults["minimum_dscr"],
                    ),
                    "step": 0.05,
                    "format": "%.2f",
                },
                {
                    "type": "number",
                    "label": "Minimum Cash Balance (EUR)",
                    "field_key": "financing.minimum_cash_balance_eur",
                    "value": _get_current_value(
                        "financing.minimum_cash_balance_eur",
                        financing_defaults["minimum_cash_balance_eur"],
                    ),
                    "step": 50000.0,
                    "format": "%.0f",
                },
                {
                    "type": "pct",
                    "label": "Target IRR",
                    "field_key": "financing.target_irr",
                    "value": _get_current_value(
                        "financing.target_irr",
                        financing_defaults["target_irr"],
                    ),
                },
                {
                    "type": "number",
                    "label": "Max Equity Contribution (EUR)",
                    "field_key": "financing.max_equity_contribution_eur",
                    "value": _get_current_value(
                        "financing.max_equity_contribution_eur",
                        financing_defaults["max_equity_contribution_eur"],
                    )
                    or 0,
                    "step": 100000.0,
                    "format": "%.0f",
                },
                {
                    "type": "pct",
                    "label": "Minimum Cash Yield",
                    "field_key": "financing.min_cash_yield",
                    "value": _get_current_value(
                        "financing.min_cash_yield",
                        financing_defaults["min_cash_yield"],
                    ),
                },
            ]
            _render_inline_controls("Financing Drivers", financing_controls, columns=1)

    # Build input model and collect editable values from the assumptions page.
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

    utilization_by_year = st.session_state.get("utilization_by_year")
    if not isinstance(utilization_by_year, list) or len(utilization_by_year) != 5:
        scenario_utilization = input_model.scenario_parameters[
            "utilization_rate"
        ][scenario_key].value
        utilization_by_year = [scenario_utilization] * 5
        st.session_state["utilization_by_year"] = utilization_by_year
    input_model.utilization_by_year = utilization_by_year

    cashflow_defaults = _default_cashflow_assumptions()
    input_model.cashflow_assumptions = {
        "tax_cash_rate_pct": st.session_state.get(
            "cashflow.tax_cash_rate_pct",
            cashflow_defaults["tax_cash_rate_pct"],
        ),
        "tax_payment_lag_years": st.session_state.get(
            "cashflow.tax_payment_lag_years",
            cashflow_defaults["tax_payment_lag_years"],
        ),
        "capex_pct_revenue": st.session_state.get(
            "cashflow.capex_pct_revenue",
            cashflow_defaults["capex_pct_revenue"],
        ),
        "working_capital_pct_revenue": st.session_state.get(
            "cashflow.working_capital_pct_revenue",
            cashflow_defaults["working_capital_pct_revenue"],
        ),
        "opening_cash_balance_eur": st.session_state.get(
            "cashflow.opening_cash_balance_eur",
            cashflow_defaults["opening_cash_balance_eur"],
        ),
    }

    balance_defaults = _default_balance_sheet_assumptions(input_model)
    input_model.balance_sheet_assumptions = {
        "opening_equity_eur": st.session_state.get(
            "balance_sheet.opening_equity_eur",
            balance_defaults["opening_equity_eur"],
        ),
        "depreciation_rate_pct": st.session_state.get(
            "balance_sheet.depreciation_rate_pct",
            balance_defaults["depreciation_rate_pct"],
        ),
        "minimum_cash_balance_eur": st.session_state.get(
            "balance_sheet.minimum_cash_balance_eur",
            balance_defaults["minimum_cash_balance_eur"],
        ),
    }

    financing_defaults = _default_financing_assumptions(input_model)
    input_model.financing_assumptions = {
        "initial_debt_eur": st.session_state.get(
            "financing.initial_debt_eur",
            financing_defaults["initial_debt_eur"],
        ),
        "interest_rate_pct": st.session_state.get(
            "financing.interest_rate_pct",
            financing_defaults["interest_rate_pct"],
        ),
        "amortization_type": st.session_state.get(
            "financing.amortization_type",
            financing_defaults["amortization_type"],
        ),
        "amortization_period_years": st.session_state.get(
            "financing.amortization_period_years",
            financing_defaults["amortization_period_years"],
        ),
        "grace_period_years": st.session_state.get(
            "financing.grace_period_years",
            financing_defaults["grace_period_years"],
        ),
        "special_repayment_year": st.session_state.get(
            "financing.special_repayment_year",
            financing_defaults["special_repayment_year"],
        ),
        "special_repayment_amount_eur": st.session_state.get(
            "financing.special_repayment_amount_eur",
            financing_defaults["special_repayment_amount_eur"],
        ),
        "minimum_dscr": st.session_state.get(
            "financing.minimum_dscr",
            financing_defaults["minimum_dscr"],
        ),
        "minimum_cash_balance_eur": st.session_state.get(
            "financing.minimum_cash_balance_eur",
            financing_defaults["minimum_cash_balance_eur"],
        ),
        "target_irr": st.session_state.get(
            "financing.target_irr",
            financing_defaults["target_irr"],
        ),
        "max_equity_contribution_eur": st.session_state.get(
            "financing.max_equity_contribution_eur",
            financing_defaults["max_equity_contribution_eur"],
        ),
        "min_cash_yield": st.session_state.get(
            "financing.min_cash_yield",
            financing_defaults["min_cash_yield"],
        ),
    }

    valuation_defaults = _default_valuation_assumptions(input_model)
    input_model.valuation_runtime = {
        "seller_ebit_multiple": st.session_state.get(
            "valuation.seller_ebit_multiple",
            valuation_defaults["seller_ebit_multiple"],
        ),
        "reference_year": st.session_state.get(
            "valuation.reference_year",
            valuation_defaults["reference_year"],
        ),
        "buyer_discount_rate": st.session_state.get(
            "valuation.buyer_discount_rate",
            valuation_defaults["buyer_discount_rate"],
        ),
        "valuation_start_year": st.session_state.get(
            "valuation.valuation_start_year",
            valuation_defaults["valuation_start_year"],
        ),
        "debt_at_close_eur": st.session_state.get(
            "valuation.debt_at_close_eur",
            valuation_defaults["debt_at_close_eur"],
        ),
        "transaction_cost_pct": st.session_state.get(
            "valuation.transaction_cost_pct",
            valuation_defaults["transaction_cost_pct"],
        ),
        "include_terminal_value": st.session_state.get(
            "valuation.include_terminal_value",
            valuation_defaults["include_terminal_value"],
        ),
    }

    # Run model calculations in the standard order.
    pnl_result = run_model.calculate_pnl(input_model)
    pnl_list = _pnl_dict_to_list(pnl_result)
    cashflow_result = run_model.calculate_cashflow(input_model, pnl_list)
    debt_schedule = run_model.calculate_debt_schedule(
        input_model, cashflow_result
    )
    balance_sheet = run_model.calculate_balance_sheet(
        input_model, cashflow_result, debt_schedule, pnl_result
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
        utilization_by_year = getattr(
            input_model, "utilization_by_year", [utilization_field.value] * 5
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
        st.write("Seller vs. buyer view (5-year plan)")
        if st.button(
            "Edit Valuation Assumptions",
            key="edit_valuation_assumptions_button",
            help="Open valuation assumptions in the sidebar",
        ):
            st.session_state["edit_valuation_assumptions"] = True

        valuation_assumptions = _default_valuation_assumptions(input_model)
        seller_multiple = st.session_state.get(
            "valuation.seller_ebit_multiple",
            valuation_assumptions["seller_ebit_multiple"],
        )
        reference_year = st.session_state.get(
            "valuation.reference_year",
            valuation_assumptions["reference_year"],
        )
        buyer_discount_rate = st.session_state.get(
            "valuation.buyer_discount_rate",
            valuation_assumptions["buyer_discount_rate"],
        )
        valuation_start_year = st.session_state.get(
            "valuation.valuation_start_year",
            valuation_assumptions["valuation_start_year"],
        )
        debt_at_close = st.session_state.get(
            "valuation.debt_at_close_eur",
            valuation_assumptions["debt_at_close_eur"],
        )
        transaction_cost_pct = st.session_state.get(
            "valuation.transaction_cost_pct",
            valuation_assumptions["transaction_cost_pct"],
        )
        include_terminal_value = st.session_state.get(
            "valuation.include_terminal_value",
            valuation_assumptions["include_terminal_value"],
        )

        pnl_table = pd.DataFrame.from_dict(pnl_result, orient="index")
        ebit_ref = pnl_table.loc[f"Year {reference_year}", "ebit"]
        seller_ev = ebit_ref * seller_multiple
        balance_table = pd.DataFrame(balance_sheet)
        net_debt_close = (
            balance_table.loc[0, "financial_debt"]
            - balance_table.loc[0, "cash"]
        )
        seller_equity_value = seller_ev - net_debt_close

        seller_rows = {
            "Reference Year EBIT": {},
            "Applied EBIT Multiple": {},
            "Enterprise Value (EV)": {},
            "Net Debt at Close": {},
            "Equity Value (Seller View)": {},
        }
        for year_index in range(5):
            year_label = f"Year {year_index}"
            seller_rows["Reference Year EBIT"][year_label] = (
                ebit_ref if year_index == reference_year else ""
            )
            seller_rows["Applied EBIT Multiple"][year_label] = (
                seller_multiple if year_index == reference_year else ""
            )
            seller_rows["Enterprise Value (EV)"][year_label] = (
                seller_ev if year_index == reference_year else ""
            )
            seller_rows["Net Debt at Close"][year_label] = (
                net_debt_close if year_index == 0 else ""
            )
            seller_rows["Equity Value (Seller View)"][year_label] = (
                seller_equity_value if year_index == reference_year else ""
            )

        st.markdown("### Seller Valuation (Multiple-Based)")
        seller_table = pd.DataFrame.from_dict(seller_rows, orient="index")
        seller_table = seller_table[
            [f"Year {i}" for i in range(5)]
        ].reset_index()
        seller_table.rename(columns={"index": "Line Item"}, inplace=True)
        seller_total_rows = {"Enterprise Value (EV)", "Equity Value (Seller View)"}
        seller_formatters = {
            "Applied EBIT Multiple": lambda value: f"{value:.2f}x"
            if value not in ("", None)
            else "",
        }
        _render_custom_table_html(
            seller_table, set(), seller_total_rows, seller_formatters
        )

        st.markdown("### Buyer Valuation (DCF)")
        cashflow_table = pd.DataFrame(cashflow_result)
        free_cashflows = cashflow_table["free_cashflow"].tolist()

        dcf_rows = {
            "Free Cashflow": {},
            "Discount Factor": {},
            "Present Value of FCF": {},
            "Cumulative PV of FCF": {},
            "Terminal Value": {},
            "Enterprise Value (DCF)": {},
            "Net Debt at Close": {},
            "Transaction Costs": {},
            "Equity Value (Buyer View)": {},
        }
        cumulative_pv = 0.0
        for year_index, fcf in enumerate(free_cashflows):
            year_label = f"Year {year_index}"
            if year_index >= valuation_start_year:
                exponent = year_index - valuation_start_year + 1
                discount_factor = (
                    1 / ((1 + buyer_discount_rate) ** exponent)
                    if buyer_discount_rate
                    else 1.0
                )
            else:
                discount_factor = 0.0
            pv_fcf = fcf * discount_factor
            cumulative_pv += pv_fcf

            dcf_rows["Free Cashflow"][year_label] = fcf
            dcf_rows["Discount Factor"][year_label] = discount_factor
            dcf_rows["Present Value of FCF"][year_label] = pv_fcf
            dcf_rows["Cumulative PV of FCF"][year_label] = cumulative_pv

        terminal_value = 0.0
        terminal_pv = 0.0
        if include_terminal_value and buyer_discount_rate:
            terminal_value = free_cashflows[-1] / buyer_discount_rate
            last_exponent = max(1, len(free_cashflows) - valuation_start_year)
            terminal_pv = terminal_value / (
                (1 + buyer_discount_rate) ** last_exponent
            )
            dcf_rows["Terminal Value"]["Year 4"] = terminal_value

        enterprise_value_dcf = cumulative_pv + terminal_pv
        transaction_costs = enterprise_value_dcf * transaction_cost_pct
        buyer_equity_value = (
            enterprise_value_dcf - debt_at_close - transaction_costs
        )

        for year_index in range(5):
            year_label = f"Year {year_index}"
            dcf_rows["Enterprise Value (DCF)"][year_label] = (
                enterprise_value_dcf if year_index == 4 else ""
            )
            dcf_rows["Net Debt at Close"][year_label] = (
                debt_at_close if year_index == 4 else ""
            )
            dcf_rows["Transaction Costs"][year_label] = (
                transaction_costs if year_index == 4 else ""
            )
            dcf_rows["Equity Value (Buyer View)"][year_label] = (
                buyer_equity_value if year_index == 4 else ""
            )

        dcf_table = pd.DataFrame.from_dict(dcf_rows, orient="index")
        dcf_table = dcf_table[[f"Year {i}" for i in range(5)]].reset_index()
        dcf_table.rename(columns={"index": "Line Item"}, inplace=True)
        dcf_total_rows = {"Enterprise Value (DCF)", "Equity Value (Buyer View)"}
        dcf_formatters = {
            "Discount Factor": lambda value: f"{value:.2f}"
            if value not in ("", None)
            else "",
        }
        _render_custom_table_html(
            dcf_table, set(), dcf_total_rows, dcf_formatters
        )

        st.markdown("### Purchase Price Bridge")
        valuation_gap = buyer_equity_value - seller_equity_value
        valuation_gap_pct = (
            valuation_gap / seller_equity_value
            if seller_equity_value
            else 0
        )
        bridge_rows = {
            "Seller Equity Value": {"Year 0": seller_equity_value},
            "Buyer Equity Value": {"Year 0": buyer_equity_value},
            "Valuation Gap (EUR)": {"Year 0": valuation_gap},
            "Valuation Gap (%)": {"Year 0": valuation_gap_pct},
        }
        for year_index in range(1, 5):
            year_label = f"Year {year_index}"
            for key in bridge_rows:
                bridge_rows[key][year_label] = ""
        bridge_table = pd.DataFrame.from_dict(bridge_rows, orient="index")
        bridge_table = bridge_table[[f"Year {i}" for i in range(5)]].reset_index()
        bridge_table.rename(columns={"index": "Line Item"}, inplace=True)
        bridge_formatters = {
            "Valuation Gap (%)": format_pct,
        }
        _render_custom_table_html(
            bridge_table, set(), {"Valuation Gap (EUR)"}, bridge_formatters
        )
        gap_label = "Buyer > Seller" if valuation_gap >= 0 else "Buyer < Seller"
        st.caption(f"Valuation gap indicator: {gap_label}.")

        st.markdown("### KPIs")
        purchase_price = input_model.transaction_and_financing[
            "purchase_price_eur"
        ].value
        year0_revenue = pnl_table.loc["Year 0", "revenue"]
        implied_ev_multiple = (
            seller_ev / ebit_ref if ebit_ref else 0
        )
        purchase_price_pct_revenue = (
            purchase_price / year0_revenue if year0_revenue else 0
        )
        kpi_table = pd.DataFrame(
            [
                {
                    "KPI": "Implied EV / EBIT (Seller)",
                    "Value": f"{implied_ev_multiple:.2f}x",
                },
                {
                    "KPI": "Implied Equity IRR (Buyer)",
                    "Value": format_pct(investment_result["irr"]),
                },
                {
                    "KPI": "Max Affordable Purchase Price (Buyer)",
                    "Value": format_currency(buyer_equity_value),
                },
                {
                    "KPI": "Headroom vs Seller Ask",
                    "Value": format_currency(valuation_gap),
                },
                {
                    "KPI": "Purchase Price as % of Revenue",
                    "Value": format_pct(purchase_price_pct_revenue),
                },
            ]
        )
        st.dataframe(kpi_table, use_container_width=True)

        explain_valuation = st.toggle("Explain Valuation Logic")
        if explain_valuation:
            st.markdown("### Seller Perspective")
            st.write(
                "The seller view uses an EBIT multiple on the selected reference year "
                "to anchor enterprise value."
            )
            st.caption(
                f"Reference Year EBIT (Year {reference_year}) = {format_currency(ebit_ref)}."
            )
            st.caption(
                f"Enterprise Value = EBIT × Multiple = {format_currency(ebit_ref)} "
                f"× {seller_multiple:.2f}x = {format_currency(seller_ev)}."
            )
            st.caption(
                f"Equity Value = EV - Net Debt = {format_currency(seller_ev)} "
                f"- {format_currency(net_debt_close)} = {format_currency(seller_equity_value)}."
            )

            st.markdown("### Buyer Perspective")
            st.write(
                "The buyer view discounts free cashflows from the valuation start year "
                "at the required return."
            )
            st.caption(
                f"Discount Rate = {format_pct(buyer_discount_rate)}; "
                f"Valuation Start Year = Year {valuation_start_year}."
            )
            st.caption(
                "Present Value of FCF = FCF × Discount Factor; "
                "Enterprise Value is the sum of PVs."
            )
            if include_terminal_value:
                st.caption(
                    f"Terminal Value = Final Year FCF / Discount Rate = "
                    f"{format_currency(free_cashflows[-1])} / "
                    f"{format_pct(buyer_discount_rate)}."
                )
            st.caption(
                f"Equity Value (Buyer) = EV - Debt at Close - Transaction Costs "
                f"= {format_currency(enterprise_value_dcf)} - "
                f"{format_currency(debt_at_close)} - "
                f"{format_currency(transaction_costs)}."
            )
            st.markdown("### Buyer vs. Seller Gap")
            st.write(
                "The valuation gap highlights the difference between seller expectations "
                "and buyer affordability after financing and transaction costs."
            )

    if page == "Operating Model (P&L)":
        st.header("Operating Model (P&L)")
        st.write("Consolidated income statement (5-year plan)")
        scenario_options = ["Base", "Best", "Worst"]
        selected_scenario = st.session_state.get(
            "scenario_selection.selected_scenario",
            input_model.scenario_selection["selected_scenario"].value,
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
        fte_growth_field = _get_field_by_path(
            input_model.__dict__,
            ["operating_assumptions", "consulting_fte_growth_pct"],
        )
        work_days_field = _get_field_by_path(
            input_model.__dict__,
            ["operating_assumptions", "work_days_per_year"],
        )
        day_rate_growth_field = _get_field_by_path(
            input_model.__dict__,
            ["operating_assumptions", "day_rate_growth_pct"],
        )
        guarantee_y1_field = _get_field_by_path(
            input_model.__dict__,
            ["operating_assumptions", "revenue_guarantee_pct_year_1"],
        )
        guarantee_y2_field = _get_field_by_path(
            input_model.__dict__,
            ["operating_assumptions", "revenue_guarantee_pct_year_2"],
        )
        guarantee_y3_field = _get_field_by_path(
            input_model.__dict__,
            ["operating_assumptions", "revenue_guarantee_pct_year_3"],
        )

        st.markdown("### P&L (GuV)")
        if st.button(
            "Edit P&L Assumptions",
            key="edit_pnl_assumptions_button",
            help="Open relevant P&L assumptions in the sidebar",
        ):
            st.session_state["edit_pnl_assumptions"] = True


        pnl_table = pd.DataFrame.from_dict(pnl_result, orient="index")
        year_indexes = list(range(len(pnl_table)))

        wage_inflation = input_model.personnel_cost_assumptions[
            "wage_inflation_pct"
        ].value
        consultant_base_cost = input_model.personnel_cost_assumptions[
            "avg_consultant_base_cost_eur_per_year"
        ].value
        bonus_pct = input_model.personnel_cost_assumptions[
            "bonus_pct_of_base"
        ].value
        payroll_pct = input_model.personnel_cost_assumptions[
            "payroll_burden_pct_of_comp"
        ].value
        backoffice_fte_start = input_model.operating_assumptions[
            "backoffice_fte_start"
        ].value
        backoffice_growth = input_model.operating_assumptions[
            "backoffice_fte_growth_pct"
        ].value
        backoffice_salary = input_model.operating_assumptions[
            "avg_backoffice_salary_eur_per_year"
        ].value
        overhead_inflation = input_model.overhead_and_variable_costs[
            "overhead_inflation_pct"
        ].value
        depreciation = input_model.capex_and_working_capital[
            "depreciation_eur_per_year"
        ].value

        interest_by_year = {
            row["year"]: row["interest_expense"]
            for row in debt_schedule
        }

        line_items = {}
        def _set_line_value(name, year_label, value):
            if name not in line_items:
                line_items[name] = {
                    "Line Item": name,
                    "Year 0": "",
                    "Year 1": "",
                    "Year 2": "",
                    "Year 3": "",
                    "Year 4": "",
                }
            line_items[name][year_label] = value

        for year_index in year_indexes:
            consultants_fte = fte_field.value * (
                (1 + fte_growth_field.value) ** year_index
            )
            billable_days = work_days_field.value
            utilization = utilization_by_year[year_index]
            day_rate = day_rate_field.value * (
                (1 + day_rate_growth_field.value) ** year_index
            )

            if year_index == 0:
                guarantee_pct = guarantee_y1_field.value
            elif year_index == 1:
                guarantee_pct = guarantee_y2_field.value
            elif year_index == 2:
                guarantee_pct = guarantee_y3_field.value
            else:
                guarantee_pct = 0

            guaranteed_revenue = (
                consultants_fte * billable_days * day_rate * guarantee_pct
            )
            non_guaranteed_revenue = (
                consultants_fte
                * billable_days
                * day_rate
                * max(utilization - guarantee_pct, 0)
            )
            total_revenue = guaranteed_revenue + non_guaranteed_revenue

            consultant_cost_per_fte = consultant_base_cost * (
                (1 + bonus_pct) + payroll_pct
            )
            consultant_cost_per_fte *= (1 + wage_inflation) ** year_index
            consultant_comp = consultant_cost_per_fte * consultants_fte

            backoffice_fte = backoffice_fte_start * (
                (1 + backoffice_growth) ** year_index
            )
            backoffice_cost_per_fte = backoffice_salary * (1 + payroll_pct)
            backoffice_cost_per_fte *= (1 + wage_inflation) ** year_index
            backoffice_comp = backoffice_cost_per_fte * backoffice_fte

            management_comp = 0

            external_advisors = input_model.overhead_and_variable_costs[
                "legal_audit_eur_per_year"
            ].value * ((1 + overhead_inflation) ** year_index)
            it_cost = input_model.overhead_and_variable_costs[
                "it_and_software_eur_per_year"
            ].value * ((1 + overhead_inflation) ** year_index)
            office_cost = input_model.overhead_and_variable_costs[
                "rent_eur_per_year"
            ].value * ((1 + overhead_inflation) ** year_index)
            other_services = (
                input_model.overhead_and_variable_costs[
                    "insurance_eur_per_year"
                ].value
                + input_model.overhead_and_variable_costs[
                    "other_overhead_eur_per_year"
                ].value
            ) * ((1 + overhead_inflation) ** year_index)

            total_personnel = consultant_comp + backoffice_comp + management_comp
            total_operating = (
                external_advisors + it_cost + office_cost + other_services
            )
            ebitda = total_revenue - total_personnel - total_operating
            ebit = ebitda - depreciation
            interest = interest_by_year.get(year_index, 0)
            ebt = ebit - interest
            taxes = pnl_table.iloc[year_index]["taxes"]
            net_income = pnl_table.iloc[year_index]["net_income"]

            year_label = f"Year {year_index}"
            _set_line_value("Guaranteed Revenue", year_label, guaranteed_revenue)
            _set_line_value(
                "Non-Guaranteed Revenue", year_label, non_guaranteed_revenue
            )
            _set_line_value("Total Revenue", year_label, total_revenue)
            _set_line_value(
                "Consultant Compensation", year_label, consultant_comp
            )
            _set_line_value(
                "Backoffice Compensation", year_label, backoffice_comp
            )
            _set_line_value(
                "Management / MD Compensation", year_label, management_comp
            )
            _set_line_value(
                "Total Personnel Costs", year_label, total_personnel
            )
            _set_line_value(
                "External Consulting / Advisors",
                year_label,
                external_advisors,
            )
            _set_line_value("IT", year_label, it_cost)
            _set_line_value("Office", year_label, office_cost)
            _set_line_value("Other Services", year_label, other_services)
            _set_line_value(
                "Total Operating Expenses", year_label, total_operating
            )
            _set_line_value("EBITDA", year_label, ebitda)
            _set_line_value("Depreciation", year_label, depreciation)
            _set_line_value("EBIT", year_label, ebit)
            _set_line_value("Interest Expense", year_label, interest)
            _set_line_value("EBT", year_label, ebt)
            _set_line_value("Taxes", year_label, taxes)
            _set_line_value("Net Income (Jahresueberschuss)", year_label, net_income)

            revenue_per_consultant = (
                total_revenue / consultants_fte if consultants_fte else 0
            )
            ebitda_margin = ebitda / total_revenue if total_revenue else 0
            ebit_margin = ebit / total_revenue if total_revenue else 0
            personnel_cost_ratio = (
                total_personnel / total_revenue if total_revenue else 0
            )
            guaranteed_pct = (
                guaranteed_revenue / total_revenue if total_revenue else 0
            )
            non_guaranteed_pct = (
                non_guaranteed_revenue / total_revenue if total_revenue else 0
            )
            net_margin = net_income / total_revenue if total_revenue else 0
            opex_ratio = total_operating / total_revenue if total_revenue else 0

            _set_line_value(
                "Revenue per Consultant",
                year_label,
                revenue_per_consultant,
            )
            _set_line_value("EBITDA Margin", year_label, ebitda_margin)
            _set_line_value("EBIT Margin", year_label, ebit_margin)
            _set_line_value("Personnel Cost Ratio", year_label, personnel_cost_ratio)
            _set_line_value("Guaranteed Revenue %", year_label, guaranteed_pct)
            _set_line_value(
                "Non-Guaranteed Revenue %", year_label, non_guaranteed_pct
            )
            _set_line_value("Net Margin", year_label, net_margin)
            _set_line_value("Opex Ratio", year_label, opex_ratio)

        row_order = [
            "Revenue",
            "Guaranteed Revenue",
            "Non-Guaranteed Revenue",
            "Total Revenue",
            "Personnel Costs",
            "Consultant Compensation",
            "Backoffice Compensation",
            "Management / MD Compensation",
            "Total Personnel Costs",
            "Operating Expenses",
            "External Consulting / Advisors",
            "IT",
            "Office",
            "Other Services",
            "Total Operating Expenses",
            "EBITDA",
            "Depreciation",
            "EBIT",
            "Interest Expense",
            "EBT",
            "Taxes",
            "Net Income (Jahresueberschuss)",
            "KPI",
            "Revenue per Consultant",
            "EBITDA Margin",
            "EBIT Margin",
            "Personnel Cost Ratio",
            "Guaranteed Revenue %",
            "Non-Guaranteed Revenue %",
            "Net Margin",
            "Opex Ratio",
        ]

        label_rows = []
        for label in row_order:
            if label in ("Revenue", "Personnel Costs", "Operating Expenses", "KPI"):
                label_rows.append(
                    {
                        "Line Item": label,
                        "Year 0": "",
                        "Year 1": "",
                        "Year 2": "",
                        "Year 3": "",
                        "Year 4": "",
                    }
                )
            else:
                row = line_items.get(label)
                if row:
                    label_rows.append(row)
                else:
                    label_rows.append(
                        {
                            "Line Item": label,
                            "Year 0": "",
                            "Year 1": "",
                            "Year 2": "",
                            "Year 3": "",
                            "Year 4": "",
                        }
                    )

        pnl_statement = pd.DataFrame(label_rows)
        format_map = {
            "Year 0": format_currency,
            "Year 1": format_currency,
            "Year 2": format_currency,
            "Year 3": format_currency,
            "Year 4": format_currency,
        }
        bold_rows = {
            "Total Revenue",
            "Total Personnel Costs",
            "Total Operating Expenses",
            "EBITDA",
            "EBIT",
            "EBT",
            "Net Income (Jahresueberschuss)",
        }
        section_rows = {"Revenue", "Personnel Costs", "Operating Expenses", "KPI"}

        def _style_pnl(df):
            styles = pd.DataFrame("", index=df.index, columns=df.columns)
            for idx, row in df.iterrows():
                for col in df.columns:
                    align = "text-align: left;" if col == "Line Item" else "text-align: right;"
                    styles.at[idx, col] += align
                if row["Line Item"] in section_rows:
                    styles.loc[idx, :] += "font-weight: 700; font-size: 1.05rem; background-color: #f9fafb;"
                if row["Line Item"] in bold_rows:
                    styles.loc[idx, :] += "font-weight: 700; background-color: #f3f4f6; border-top: 1px solid #c7c7c7;"
            return styles

        _render_pnl_html(pnl_statement, section_rows, bold_rows)

        # Table rendered via HTML for full-width layout.

        explain_pnl = st.toggle("Explain P&L logic")
        if explain_pnl:
            def _format_currency_expl(value):
                formatted = format_currency(value)
                return formatted if formatted else "—"

            def _format_pct_expl(value):
                formatted = format_pct(value)
                return formatted if formatted else "—"

            def _format_int_expl(value):
                formatted = format_int(value)
                return formatted if formatted else "—"

            def _safe_calc(values, func):
                if any(value is None for value in values):
                    return None
                try:
                    return func(*values)
                except Exception:
                    return None

            year_labels = [f"Year {year_index}" for year_index in year_indexes]
            utilization_by_year = getattr(
                input_model, "utilization_by_year", [utilization_field.value] * 5
            )

            st.markdown("### Revenue Logic")
            st.write(
                "Revenue is built from delivery capacity and pricing. Day rate "
                "grows by the annual day-rate growth assumption. Utilization is "
                "set per year and guarantees create a floor for revenue in years 1–3."
            )

            st.caption(
                f"Day Rate_y = {_format_int_expl(day_rate_field.value)} EUR "
                f"× (1 + {_format_pct_expl(day_rate_growth_field.value)})^y"
            )

            driver_metrics = {
                "Consulting FTE": {},
                "Workdays per Year": {},
                "Utilization %": {},
                "Day Rate (EUR)": {},
                "Guarantee %": {},
            }
            revenue_metrics = {}
            missing_revenue_inputs = [
                name
                for name, value in [
                    ("Consulting FTE", fte_field.value),
                    ("FTE Growth %", fte_growth_field.value),
                    ("Workdays per Year", work_days_field.value),
                    ("Day Rate (EUR)", day_rate_field.value),
                    ("Day Rate Growth %", day_rate_growth_field.value),
                    ("Utilization (per year)", utilization_by_year),
                ]
                if value is None
            ]

            for year_index, year_label in enumerate(year_labels):
                consultants_fte = _safe_calc(
                    [fte_field.value, fte_growth_field.value],
                    lambda fte, growth: fte * ((1 + growth) ** year_index),
                )
                day_rate_year = _safe_calc(
                    [day_rate_field.value, day_rate_growth_field.value],
                    lambda rate, growth: rate * ((1 + growth) ** year_index),
                )
                utilization = (
                    utilization_by_year[year_index]
                    if isinstance(utilization_by_year, list)
                    else None
                )
                guarantee_pct = 0
                if year_index == 0:
                    guarantee_pct = guarantee_y1_field.value
                elif year_index == 1:
                    guarantee_pct = guarantee_y2_field.value
                elif year_index == 2:
                    guarantee_pct = guarantee_y3_field.value

                theoretical_revenue = _safe_calc(
                    [
                        consultants_fte,
                        work_days_field.value,
                        utilization,
                        day_rate_year,
                    ],
                    lambda fte, days, util, rate: fte * days * util * rate,
                )
                guaranteed_revenue = _safe_calc(
                    [
                        consultants_fte,
                        work_days_field.value,
                        day_rate_year,
                        guarantee_pct,
                    ],
                    lambda fte, days, rate, guarantee: fte * days * rate * guarantee,
                )
                non_guaranteed_revenue = _safe_calc(
                    [
                        consultants_fte,
                        work_days_field.value,
                        day_rate_year,
                        utilization,
                        guarantee_pct,
                    ],
                    lambda fte, days, rate, util, guarantee: fte
                    * days
                    * rate
                    * max(util - guarantee, 0),
                )

                driver_metrics["Consulting FTE"][year_label] = consultants_fte
                driver_metrics["Workdays per Year"][year_label] = (
                    work_days_field.value
                )
                driver_metrics["Utilization %"][year_label] = utilization
                driver_metrics["Day Rate (EUR)"][year_label] = day_rate_year
                driver_metrics["Guarantee %"][year_label] = guarantee_pct

                for metric, value in (
                    ("Theoretical Revenue", theoretical_revenue),
                    ("Guaranteed Revenue", guaranteed_revenue),
                    ("Non-Guaranteed Revenue", non_guaranteed_revenue),
                ):
                    if metric not in revenue_metrics:
                        revenue_metrics[metric] = {
                            year: None for year in year_labels
                        }
                    revenue_metrics[metric][year_label] = value

            driver_table = pd.DataFrame.from_dict(
                driver_metrics, orient="index"
            )
            driver_table = driver_table[year_labels]
            for metric in driver_table.index:
                if metric in {"Utilization %", "Guarantee %"}:
                    driver_table.loc[metric] = driver_table.loc[metric].apply(
                        _format_pct_expl
                    )
                elif metric in {"Consulting FTE", "Workdays per Year", "Day Rate (EUR)"}:
                    driver_table.loc[metric] = driver_table.loc[metric].apply(
                        _format_int_expl
                    )
                else:
                    driver_table.loc[metric] = driver_table.loc[metric].apply(
                        _format_currency_expl
                    )
            st.dataframe(driver_table, use_container_width=True)

            revenue_table = pd.DataFrame.from_dict(
                revenue_metrics, orient="index"
            )
            revenue_table = revenue_table[year_labels].applymap(
                _format_currency_expl
            )
            st.dataframe(revenue_table, use_container_width=True)
            st.caption(
                "Formula: FTE × Workdays × Utilization_y × DayRate_y, where "
                "DayRate_y = DayRate × (1 + Day Rate Growth)^y."
            )
            if missing_revenue_inputs:
                st.caption(
                    f"Missing inputs: {', '.join(missing_revenue_inputs)}."
                )
            year0_revenue = line_items.get("Total Revenue", {}).get("Year 0")
            if year0_revenue is not None:
                target_revenue = 20_000_000
                delta = year0_revenue - target_revenue
                status = "OK" if abs(delta) <= 10_000 else "Check"
                st.caption(
                    f"Debug: Year 0 Total Revenue = {_format_currency_expl(year0_revenue)} "
                    f"(target 20.00 m EUR, {status})."
                )

            st.markdown("### Personnel Costs Logic")
            st.write(
                "Consultant compensation is driven by base cost per consultant, "
                "bonus and payroll burden, with wage inflation applied annually. "
                "Backoffice costs follow the same inflation logic."
            )

            personnel_metrics = {}
            missing_personnel_inputs = [
                name
                for name, value in [
                    ("Consulting FTE", fte_field.value),
                    ("FTE Growth %", fte_growth_field.value),
                    ("Consultant Base Cost", consultant_base_cost),
                    ("Bonus %", bonus_pct),
                    ("Payroll Burden %", payroll_pct),
                    ("Wage Inflation %", wage_inflation),
                    ("Backoffice FTE", backoffice_fte_start),
                    ("Backoffice Growth %", backoffice_growth),
                    ("Backoffice Salary", backoffice_salary),
                ]
                if value is None
            ]

            for year_index, year_label in enumerate(year_labels):
                consultants_fte = _safe_calc(
                    [fte_field.value, fte_growth_field.value],
                    lambda fte, growth: fte * ((1 + growth) ** year_index),
                )
                consultant_cost_per_fte = _safe_calc(
                    [consultant_base_cost, bonus_pct, payroll_pct, wage_inflation],
                    lambda base, bonus, payroll, inflation: base
                    * ((1 + bonus) + payroll)
                    * ((1 + inflation) ** year_index),
                )
                consultant_comp = _safe_calc(
                    [consultants_fte, consultant_cost_per_fte],
                    lambda fte, cost: fte * cost,
                )
                backoffice_fte = _safe_calc(
                    [backoffice_fte_start, backoffice_growth],
                    lambda fte, growth: fte * ((1 + growth) ** year_index),
                )
                backoffice_cost_per_fte = _safe_calc(
                    [backoffice_salary, payroll_pct, wage_inflation],
                    lambda salary, payroll, inflation: salary
                    * (1 + payroll)
                    * ((1 + inflation) ** year_index),
                )
                backoffice_comp = _safe_calc(
                    [backoffice_fte, backoffice_cost_per_fte],
                    lambda fte, cost: fte * cost,
                )
                management_comp = 0
                total_personnel = _safe_calc(
                    [consultant_comp, backoffice_comp, management_comp],
                    lambda a, b, c: a + b + c,
                )

                for metric, value in (
                    ("Consultant Compensation", consultant_comp),
                    ("Backoffice Compensation", backoffice_comp),
                    ("Management / MD Compensation", management_comp),
                    ("Total Personnel Costs", total_personnel),
                ):
                    if metric not in personnel_metrics:
                        personnel_metrics[metric] = {
                            year: None for year in year_labels
                        }
                    personnel_metrics[metric][year_label] = value

            personnel_table = pd.DataFrame.from_dict(
                personnel_metrics, orient="index"
            )
            personnel_table = personnel_table[year_labels].applymap(
                _format_currency_expl
            )
            st.dataframe(personnel_table, use_container_width=True)
            personnel_ratio_table = pd.DataFrame.from_dict(
                {"Personnel Cost Ratio": line_items["Personnel Cost Ratio"]},
                orient="index",
            )
            personnel_ratio_table = personnel_ratio_table[year_labels]
            personnel_ratio_table = personnel_ratio_table.applymap(
                _format_pct_expl
            )
            st.dataframe(personnel_ratio_table, use_container_width=True)
            if missing_personnel_inputs:
                st.caption(
                    f"Missing inputs: {', '.join(missing_personnel_inputs)}."
                )

            st.markdown("### Operating Expenses Logic")
            st.write(
                "Operating expenses are built from fixed annual costs inflated by "
                "the overhead inflation assumption."
            )

            opex_metrics = {}
            missing_opex_inputs = [
                name
                for name, value in [
                    ("External Advisors", input_model.overhead_and_variable_costs["legal_audit_eur_per_year"].value),
                    ("IT", input_model.overhead_and_variable_costs["it_and_software_eur_per_year"].value),
                    ("Office", input_model.overhead_and_variable_costs["rent_eur_per_year"].value),
                    ("Insurance", input_model.overhead_and_variable_costs["insurance_eur_per_year"].value),
                    ("Other Services", input_model.overhead_and_variable_costs["other_overhead_eur_per_year"].value),
                    ("Overhead Inflation %", overhead_inflation),
                ]
                if value is None
            ]

            for year_index, year_label in enumerate(year_labels):
                external_advisors = _safe_calc(
                    [
                        input_model.overhead_and_variable_costs[
                            "legal_audit_eur_per_year"
                        ].value,
                        overhead_inflation,
                    ],
                    lambda base, inflation: base * ((1 + inflation) ** year_index),
                )
                it_cost = _safe_calc(
                    [
                        input_model.overhead_and_variable_costs[
                            "it_and_software_eur_per_year"
                        ].value,
                        overhead_inflation,
                    ],
                    lambda base, inflation: base * ((1 + inflation) ** year_index),
                )
                office_cost = _safe_calc(
                    [
                        input_model.overhead_and_variable_costs[
                            "rent_eur_per_year"
                        ].value,
                        overhead_inflation,
                    ],
                    lambda base, inflation: base * ((1 + inflation) ** year_index),
                )
                other_services = _safe_calc(
                    [
                        input_model.overhead_and_variable_costs[
                            "insurance_eur_per_year"
                        ].value,
                        input_model.overhead_and_variable_costs[
                            "other_overhead_eur_per_year"
                        ].value,
                        overhead_inflation,
                    ],
                    lambda insurance, other, inflation: (insurance + other)
                    * ((1 + inflation) ** year_index),
                )
                total_opex = _safe_calc(
                    [external_advisors, it_cost, office_cost, other_services],
                    lambda a, b, c, d: a + b + c + d,
                )

                for metric, value in (
                    ("External Consulting / Advisors", external_advisors),
                    ("IT", it_cost),
                    ("Office", office_cost),
                    ("Other Services", other_services),
                    ("Total Operating Expenses", total_opex),
                ):
                    if metric not in opex_metrics:
                        opex_metrics[metric] = {
                            year: None for year in year_labels
                        }
                    opex_metrics[metric][year_label] = value

            opex_table = pd.DataFrame.from_dict(opex_metrics, orient="index")
            opex_table = opex_table[year_labels].applymap(
                _format_currency_expl
            )
            st.dataframe(opex_table, use_container_width=True)
            opex_ratio_table = pd.DataFrame.from_dict(
                {"Opex Ratio": line_items["Opex Ratio"]},
                orient="index",
            )
            opex_ratio_table = opex_ratio_table[year_labels]
            opex_ratio_table = opex_ratio_table.applymap(_format_pct_expl)
            st.dataframe(opex_ratio_table, use_container_width=True)
            if missing_opex_inputs:
                st.caption(
                    f"Missing inputs: {', '.join(missing_opex_inputs)}."
                )

            st.markdown("### Earnings Bridge")
            st.write(
                "EBITDA bridges revenue to operating costs. EBIT subtracts "
                "depreciation, then interest and taxes produce net income."
            )

            earnings_metrics = {}
            for year_index, year_label in enumerate(year_labels):
                revenue = line_items["Total Revenue"][year_label]
                total_personnel = line_items["Total Personnel Costs"][year_label]
                total_operating = line_items["Total Operating Expenses"][year_label]
                ebitda = line_items["EBITDA"][year_label]
                interest = line_items["Interest Expense"][year_label]
                taxes = line_items["Taxes"][year_label]
                net_income = line_items["Net Income (Jahresueberschuss)"][
                    year_label
                ]
                ebit = line_items["EBIT"][year_label]
                ebt = line_items["EBT"][year_label]

                for metric, value in (
                    ("Total Revenue", revenue),
                    ("Total Personnel Costs", total_personnel),
                    ("Total Operating Expenses", total_operating),
                    ("EBITDA", ebitda),
                    ("Depreciation", depreciation),
                    ("EBIT", ebit),
                    ("Interest Expense", interest),
                    ("EBT", ebt),
                    ("Taxes", taxes),
                    ("Net Income (Jahresueberschuss)", net_income),
                ):
                    if metric not in earnings_metrics:
                        earnings_metrics[metric] = {
                            year: None for year in year_labels
                        }
                    earnings_metrics[metric][year_label] = value

            earnings_table = pd.DataFrame.from_dict(
                earnings_metrics, orient="index"
            )
            earnings_table = earnings_table[year_labels].applymap(
                _format_currency_expl
            )
            st.dataframe(earnings_table, use_container_width=True)

            st.markdown("### KPI Definitions")
            st.write(
                "KPIs summarize profitability and efficiency using the P&L "
                "line items for each year."
            )

            kpi_metrics = {
                "Revenue per Consultant": {},
                "EBITDA Margin": {},
                "EBIT Margin": {},
                "Personnel Cost Ratio": {},
                "Guaranteed Revenue %": {},
                "Non-Guaranteed Revenue %": {},
                "Net Margin": {},
                "Opex Ratio": {},
            }
            for year_label in year_labels:
                kpi_metrics["Revenue per Consultant"][year_label] = line_items[
                    "Revenue per Consultant"
                ][year_label]
                kpi_metrics["EBITDA Margin"][year_label] = line_items[
                    "EBITDA Margin"
                ][year_label]
                kpi_metrics["EBIT Margin"][year_label] = line_items[
                    "EBIT Margin"
                ][year_label]
                kpi_metrics["Personnel Cost Ratio"][year_label] = line_items[
                    "Personnel Cost Ratio"
                ][year_label]
                kpi_metrics["Guaranteed Revenue %"][year_label] = line_items[
                    "Guaranteed Revenue %"
                ][year_label]
                kpi_metrics["Non-Guaranteed Revenue %"][year_label] = line_items[
                    "Non-Guaranteed Revenue %"
                ][year_label]
                kpi_metrics["Net Margin"][year_label] = line_items[
                    "Net Margin"
                ][year_label]
                kpi_metrics["Opex Ratio"][year_label] = line_items[
                    "Opex Ratio"
                ][year_label]

            kpi_table = pd.DataFrame.from_dict(kpi_metrics, orient="index")
            kpi_table = kpi_table[year_labels]
            percent_kpis = {
                "EBITDA Margin",
                "EBIT Margin",
                "Personnel Cost Ratio",
                "Guaranteed Revenue %",
                "Non-Guaranteed Revenue %",
                "Net Margin",
                "Opex Ratio",
            }
            for metric in kpi_table.index:
                formatter = (
                    _format_pct_expl
                    if metric in percent_kpis
                    else _format_currency_expl
                )
                kpi_table.loc[metric] = kpi_table.loc[metric].apply(formatter)
            st.dataframe(kpi_table, use_container_width=True)

        pnl_excel = _build_pnl_excel(input_model)
        st.download_button(
            "Download P&L as Excel",
            data=pnl_excel.getvalue(),
            file_name="Financial_Model_PnL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if page == "Cashflow & Liquidity":
        st.header("Cashflow & Liquidity")
        st.write("Consolidated cashflow statement (5-year plan)")
        if st.button(
            "Edit Cashflow Assumptions",
            key="edit_cashflow_assumptions_button",
            help="Open cashflow assumptions in the sidebar",
        ):
            st.session_state["edit_cashflow_assumptions"] = True

        cashflow_line_items = {}

        def _set_cashflow_value(name, year_label, value):
            if name not in cashflow_line_items:
                cashflow_line_items[name] = {
                    "Line Item": name,
                    "Year 0": "",
                    "Year 1": "",
                    "Year 2": "",
                    "Year 3": "",
                    "Year 4": "",
                }
            cashflow_line_items[name][year_label] = value

        for row in cashflow_result:
            year_label = f"Year {row['year']}"
            _set_cashflow_value("EBITDA", year_label, row["ebitda"])
            _set_cashflow_value("Taxes Paid", year_label, row["taxes_paid"])
            _set_cashflow_value(
                "Working Capital Change",
                year_label,
                row["working_capital_change"],
            )
            _set_cashflow_value(
                "Operating Cashflow", year_label, row["operating_cf"]
            )
            _set_cashflow_value("Capex", year_label, row["capex"])
            _set_cashflow_value(
                "Free Cashflow", year_label, row["free_cashflow"]
            )
            _set_cashflow_value(
                "Debt Drawdown", year_label, row["debt_drawdown"]
            )
            _set_cashflow_value(
                "Interest Paid", year_label, row["interest_paid"]
            )
            _set_cashflow_value(
                "Debt Repayment", year_label, row["debt_repayment"]
            )
            _set_cashflow_value(
                "Net Cashflow", year_label, row["net_cashflow"]
            )
            _set_cashflow_value(
                "Opening Cash", year_label, row["opening_cash"]
            )
            _set_cashflow_value(
                "Closing Cash", year_label, row["cash_balance"]
            )

        cashflow_row_order = [
            "OPERATING CASHFLOW",
            "EBITDA",
            "Taxes Paid",
            "Working Capital Change",
            "Operating Cashflow",
            "INVESTING CASHFLOW",
            "Capex",
            "Free Cashflow",
            "FINANCING CASHFLOW",
            "Debt Drawdown",
            "Interest Paid",
            "Debt Repayment",
            "Net Cashflow",
            "LIQUIDITY",
            "Opening Cash",
            "Net Cashflow",
            "Closing Cash",
        ]

        cashflow_rows = []
        for label in cashflow_row_order:
            if label in (
                "OPERATING CASHFLOW",
                "INVESTING CASHFLOW",
                "FINANCING CASHFLOW",
                "LIQUIDITY",
            ):
                cashflow_rows.append(
                    {
                        "Line Item": label,
                        "Year 0": "",
                        "Year 1": "",
                        "Year 2": "",
                        "Year 3": "",
                        "Year 4": "",
                    }
                )
            else:
                cashflow_rows.append(cashflow_line_items.get(label))

        cashflow_statement = pd.DataFrame(cashflow_rows)
        cashflow_section_rows = {
            "OPERATING CASHFLOW",
            "INVESTING CASHFLOW",
            "FINANCING CASHFLOW",
            "LIQUIDITY",
        }
        cashflow_total_rows = {
            "Operating Cashflow",
            "Free Cashflow",
            "Net Cashflow",
            "Closing Cash",
        }
        _render_cashflow_html(
            cashflow_statement, cashflow_section_rows, cashflow_total_rows
        )

        cashflow_years = [f"Year {i}" for i in range(5)]
        kpi_metrics = {
            "Operating Cashflow": {},
            "Free Cashflow": {},
            "Cash Conversion": {},
        }
        for year_label in cashflow_years:
            ebitda = cashflow_line_items["EBITDA"][year_label]
            free_cf = cashflow_line_items["Free Cashflow"][year_label]
            operating_cf = cashflow_line_items["Operating Cashflow"][year_label]
            kpi_metrics["Operating Cashflow"][year_label] = operating_cf
            kpi_metrics["Free Cashflow"][year_label] = free_cf
            kpi_metrics["Cash Conversion"][year_label] = (
                free_cf / ebitda if ebitda else 0
            )

        kpi_table = pd.DataFrame.from_dict(kpi_metrics, orient="index")
        kpi_table = kpi_table[cashflow_years]
        for metric in kpi_table.index:
            formatter = (
                format_pct if metric == "Cash Conversion" else format_currency
            )
            kpi_table.loc[metric] = kpi_table.loc[metric].apply(formatter)
        st.markdown("### KPI Summary")
        st.dataframe(kpi_table, use_container_width=True)

        cash_balances = [
            row["cash_balance"] for row in cashflow_result
        ]
        min_cash = min(cash_balances) if cash_balances else 0
        max_cash = max(cash_balances) if cash_balances else 0
        negative_years = [
            f"Year {row['year']}"
            for row in cashflow_result
            if row["cash_balance"] < 0
        ]
        summary_table = pd.DataFrame(
            [
                {
                    "Metric": "Minimum Cash Balance",
                    "Value": format_currency(min_cash),
                },
                {
                    "Metric": "Cash Volatility (Max - Min)",
                    "Value": format_currency(max_cash - min_cash),
                },
                {
                    "Metric": "Years with Negative Cash",
                    "Value": ", ".join(negative_years) if negative_years else "None",
                },
            ]
        )
        st.dataframe(summary_table, use_container_width=True)

        explain_cashflow = st.toggle("Explain Cashflow Logic")
        if explain_cashflow:
            cashflow_assumptions = input_model.cashflow_assumptions
            st.markdown("### Operating Cashflow Logic")
            st.write(
                "Operating cashflow starts from EBITDA and adjusts for cash "
                "taxes and the working capital timing proxy."
            )
            st.write(
                "Cash taxes differ from tax expense because they are based on "
                "EBT and can be paid with a lag."
            )
            operating_table = pd.DataFrame.from_dict(
                {
                    "EBITDA": cashflow_line_items["EBITDA"],
                    "Taxes Paid": cashflow_line_items["Taxes Paid"],
                    "Working Capital Change": cashflow_line_items[
                        "Working Capital Change"
                    ],
                    "Operating Cashflow": cashflow_line_items[
                        "Operating Cashflow"
                    ],
                },
                orient="index",
            )
            operating_table = operating_table[cashflow_years].applymap(
                format_currency
            )
            st.dataframe(operating_table, use_container_width=True)
            st.caption(
                "Operating Cashflow = EBITDA - Taxes Paid - Working Capital Change."
            )
            st.caption(
                "Taxes Paid = max(EBT, 0) × "
                f"{format_pct(cashflow_assumptions['tax_cash_rate_pct'])} "
                f"with a {cashflow_assumptions['tax_payment_lag_years']}-year lag."
            )
            st.caption(
                "Working Capital Change = Revenue × "
                f"{format_pct(cashflow_assumptions['working_capital_pct_revenue'])}."
            )

            st.markdown("### Investing Cashflow Logic")
            st.write(
                "Capex is modeled as a stable percentage of revenue, which "
                "is typical for consulting businesses with limited fixed assets."
            )
            investing_table = pd.DataFrame.from_dict(
                {
                    "Capex": cashflow_line_items["Capex"],
                    "Free Cashflow": cashflow_line_items["Free Cashflow"],
                },
                orient="index",
            )
            investing_table = investing_table[cashflow_years].applymap(
                format_currency
            )
            st.dataframe(investing_table, use_container_width=True)
            st.caption("Free Cashflow = Operating Cashflow - Capex.")
            st.caption(
                "Capex = Revenue × "
                f"{format_pct(cashflow_assumptions['capex_pct_revenue'])}."
            )

            st.markdown("### Financing Cashflow Logic")
            st.write(
                "Financing cashflow reflects initial debt funding and annual "
                "debt service. Interest paid is based on outstanding principal."
            )
            financing_table = pd.DataFrame.from_dict(
                {
                    "Debt Drawdown": cashflow_line_items["Debt Drawdown"],
                    "Interest Paid": cashflow_line_items["Interest Paid"],
                    "Debt Repayment": cashflow_line_items["Debt Repayment"],
                    "Net Cashflow": cashflow_line_items["Net Cashflow"],
                },
                orient="index",
            )
            financing_table = financing_table[cashflow_years].applymap(
                format_currency
            )
            st.dataframe(financing_table, use_container_width=True)
            st.caption(
                "Net Cashflow = Free Cashflow + Financing Cashflow."
            )

            st.markdown("### Liquidity Logic")
            st.write(
                "Closing cash is the opening balance plus net cashflow, "
                "highlighting years with potential liquidity pressure."
            )
            liquidity_table = pd.DataFrame.from_dict(
                {
                    "Opening Cash": cashflow_line_items["Opening Cash"],
                    "Net Cashflow": cashflow_line_items["Net Cashflow"],
                    "Closing Cash": cashflow_line_items["Closing Cash"],
                },
                orient="index",
            )
            liquidity_table = liquidity_table[cashflow_years].applymap(
                format_currency
            )
            st.dataframe(liquidity_table, use_container_width=True)

    if page == "Balance Sheet":
        st.header("Balance Sheet")
        st.write("Simplified balance sheet (5-year plan)")
        if st.button(
            "Edit Balance Sheet Assumptions",
            key="edit_balance_sheet_assumptions_button",
            help="Open balance sheet assumptions in the sidebar",
        ):
            st.session_state["edit_balance_sheet_assumptions"] = True

        balance_line_items = {}

        def _set_balance_value(name, year_label, value):
            if name not in balance_line_items:
                balance_line_items[name] = {
                    "Line Item": name,
                    "Year 0": "",
                    "Year 1": "",
                    "Year 2": "",
                    "Year 3": "",
                    "Year 4": "",
                }
            balance_line_items[name][year_label] = value

        for row in balance_sheet:
            year_label = f"Year {row['year']}"
            _set_balance_value("Cash", year_label, row["cash"])
            _set_balance_value(
                "Fixed Assets (Net)", year_label, row["fixed_assets"]
            )
            _set_balance_value(
                "Total Assets", year_label, row["total_assets"]
            )
            _set_balance_value(
                "Financial Debt", year_label, row["financial_debt"]
            )
            _set_balance_value(
                "Total Liabilities", year_label, row["total_liabilities"]
            )
            _set_balance_value(
                "Equity at Start of Year", year_label, row["equity_start"]
            )
            _set_balance_value("Net Income", year_label, row["net_income"])
            _set_balance_value("Dividends", year_label, row["dividends"])
            _set_balance_value(
                "Equity at End of Year", year_label, row["equity_end"]
            )
            _set_balance_value(
                "Total Liabilities + Equity",
                year_label,
                row["total_liabilities_equity"],
            )
            _set_balance_value(
                "Total Assets",
                year_label,
                row["total_assets"],
            )

        balance_row_order = [
            "ASSETS",
            "Cash",
            "Fixed Assets (Net)",
            "Total Assets",
            "LIABILITIES",
            "Financial Debt",
            "Total Liabilities",
            "EQUITY",
            "Equity at Start of Year",
            "Net Income",
            "Dividends",
            "Equity at End of Year",
            "CHECK",
            "Total Assets",
            "Total Liabilities + Equity",
        ]

        balance_rows = []
        for label in balance_row_order:
            if label in ("ASSETS", "LIABILITIES", "EQUITY", "CHECK"):
                balance_rows.append(
                    {
                        "Line Item": label,
                        "Year 0": "",
                        "Year 1": "",
                        "Year 2": "",
                        "Year 3": "",
                        "Year 4": "",
                    }
                )
            else:
                row = balance_line_items.get(label)
                if row:
                    balance_rows.append(row)
                else:
                    balance_rows.append(
                        {
                            "Line Item": label,
                            "Year 0": "",
                            "Year 1": "",
                            "Year 2": "",
                            "Year 3": "",
                            "Year 4": "",
                        }
                    )

        balance_statement = pd.DataFrame(balance_rows)
        balance_section_rows = {"ASSETS", "LIABILITIES", "EQUITY", "CHECK"}
        balance_total_rows = {
            "Total Assets",
            "Total Liabilities",
            "Equity at End of Year",
            "Total Liabilities + Equity",
        }
        _render_balance_sheet_html(
            balance_statement, balance_section_rows, balance_total_rows
        )

        reconciliation_issues = [
            f"Year {row['year']}"
            for row in balance_sheet
            if abs(row["balance_check"]) > 1e-2
        ]
        if reconciliation_issues:
            st.warning(
                "Balance sheet does not reconcile in "
                f"{', '.join(reconciliation_issues)}."
            )

        cashflow_years = [f"Year {i}" for i in range(5)]
        ebitda_by_year = pd.DataFrame.from_dict(
            pnl_result, orient="index"
        )["ebitda"].to_dict()
        kpi_metrics = {
            "Net Debt": {},
            "Equity Ratio": {},
            "Net Debt / EBITDA": {},
            "Minimum Cash Headroom": {},
        }
        min_cash = input_model.balance_sheet_assumptions[
            "minimum_cash_balance_eur"
        ]

        for row in balance_sheet:
            year_label = f"Year {row['year']}"
            net_debt = row["financial_debt"] - row["cash"]
            equity_ratio = (
                row["equity_end"] / row["total_assets"]
                if row["total_assets"]
                else 0
            )
            ebitda = ebitda_by_year.get(f"Year {row['year']}", 0)
            net_debt_to_ebitda = (
                net_debt / ebitda if ebitda else 0
            )
            cash_headroom = row["cash"] - min_cash

            kpi_metrics["Net Debt"][year_label] = net_debt
            kpi_metrics["Equity Ratio"][year_label] = equity_ratio
            kpi_metrics["Net Debt / EBITDA"][year_label] = net_debt_to_ebitda
            kpi_metrics["Minimum Cash Headroom"][year_label] = cash_headroom

        kpi_table = pd.DataFrame.from_dict(kpi_metrics, orient="index")
        kpi_table = kpi_table[cashflow_years]
        for metric in kpi_table.index:
            if metric in {"Equity Ratio"}:
                kpi_table.loc[metric] = kpi_table.loc[metric].apply(
                    format_pct
                )
            elif metric == "Net Debt / EBITDA":
                kpi_table.loc[metric] = kpi_table.loc[metric].apply(
                    lambda value: f"{value:.2f}x"
                )
            else:
                kpi_table.loc[metric] = kpi_table.loc[metric].apply(
                    format_currency
                )
        st.markdown("### KPI Summary")
        st.dataframe(kpi_table, use_container_width=True)

        explain_balance = st.toggle("Explain Balance Sheet Logic")
        if explain_balance:
            balance_assumptions = input_model.balance_sheet_assumptions
            st.markdown("### Balance Sheet Scope")
            st.write(
                "This balance sheet is intentionally simplified for a consulting "
                "carve-out. It focuses on cash, fixed assets, debt, and equity "
                "to support decision-making and bank discussions."
            )
            st.write(
                "Receivables, payables, and inventory are excluded because the "
                "model uses a working-capital proxy in cashflow."
            )

            st.markdown("### Asset Logic")
            asset_table = pd.DataFrame.from_dict(
                {
                    "Cash": balance_line_items["Cash"],
                    "Fixed Assets (Net)": balance_line_items[
                        "Fixed Assets (Net)"
                    ],
                    "Total Assets": balance_line_items["Total Assets"],
                },
                orient="index",
            )
            asset_table = asset_table[cashflow_years].applymap(
                format_currency
            )
            st.dataframe(asset_table, use_container_width=True)
            st.caption(
                "Cash is taken directly from the cashflow closing cash balance."
            )
            st.caption(
                "Fixed Assets end = Prior Fixed Assets + Capex - Depreciation."
            )
            st.caption(
                "Depreciation = (Fixed Assets + Capex) × "
                f"{format_pct(balance_assumptions['depreciation_rate_pct'])}."
            )

            st.markdown("### Debt Logic")
            debt_table = pd.DataFrame.from_dict(
                {
                    "Financial Debt": balance_line_items["Financial Debt"],
                    "Total Liabilities": balance_line_items["Total Liabilities"],
                },
                orient="index",
            )
            debt_table = debt_table[cashflow_years].applymap(format_currency)
            st.dataframe(debt_table, use_container_width=True)
            st.caption(
                "Financial debt follows the debt schedule from the financing model."
            )

            st.markdown("### Equity Logic")
            equity_table = pd.DataFrame.from_dict(
                {
                    "Equity at Start of Year": balance_line_items[
                        "Equity at Start of Year"
                    ],
                    "Net Income": balance_line_items["Net Income"],
                    "Dividends": balance_line_items["Dividends"],
                    "Equity at End of Year": balance_line_items[
                        "Equity at End of Year"
                    ],
                },
                orient="index",
            )
            equity_table = equity_table[cashflow_years].applymap(
                format_currency
            )
            st.dataframe(equity_table, use_container_width=True)
            st.caption(
                "Equity end = Equity start + Net Income - Dividends (assumed 0)."
            )

            st.markdown("### Reconciliation Check")
            check_table = pd.DataFrame.from_dict(
                {
                    "Total Assets": balance_line_items["Total Assets"],
                    "Total Liabilities + Equity": balance_line_items[
                        "Total Liabilities + Equity"
                    ],
                },
                orient="index",
            )
            check_table = check_table[cashflow_years].applymap(
                format_currency
            )
            st.dataframe(check_table, use_container_width=True)

    if page == "Financing & Debt":
        st.header("Financing & Debt")
        st.write("Debt structure, service and bankability (5-year plan)")
        if st.button(
            "Edit Financing Assumptions",
            key="edit_financing_assumptions_button",
            help="Open financing assumptions in the sidebar",
        ):
            st.session_state["edit_financing_assumptions"] = True

        financing_assumptions = input_model.financing_assumptions
        debt_line_items = {}
        cashflow_by_year = {
            row["year"]: row for row in cashflow_result
        }
        valuation_runtime = getattr(
            input_model,
            "valuation_runtime",
            _default_valuation_assumptions(input_model),
        )

        purchase_price = input_model.transaction_and_financing[
            "purchase_price_eur"
        ].value
        equity_contribution = input_model.transaction_and_financing[
            "equity_contribution_eur"
        ].value
        transaction_fee_pct = valuation_runtime.get(
            "transaction_cost_pct", 0.0
        )
        transaction_fees = purchase_price * transaction_fee_pct
        refinancing_amount = 0.0
        minimum_cash_close = financing_assumptions[
            "minimum_cash_balance_eur"
        ]
        total_uses = (
            purchase_price
            + transaction_fees
            + refinancing_amount
            + minimum_cash_close
        )
        senior_debt = financing_assumptions["initial_debt_eur"]
        total_sources = senior_debt + equity_contribution

        sources_uses_rows = [
            {
                "Line Item": "USES",
                "Year 0": "",
                "Year 1": "",
                "Year 2": "",
                "Year 3": "",
                "Year 4": "",
            },
            {
                "Line Item": "Purchase Price",
                "Year 0": purchase_price,
                "Year 1": "",
                "Year 2": "",
                "Year 3": "",
                "Year 4": "",
            },
            {
                "Line Item": "Transaction Fees",
                "Year 0": transaction_fees,
                "Year 1": "",
                "Year 2": "",
                "Year 3": "",
                "Year 4": "",
            },
            {
                "Line Item": "Refinancing",
                "Year 0": refinancing_amount,
                "Year 1": "",
                "Year 2": "",
                "Year 3": "",
                "Year 4": "",
            },
            {
                "Line Item": "Minimum Cash at Close",
                "Year 0": minimum_cash_close,
                "Year 1": "",
                "Year 2": "",
                "Year 3": "",
                "Year 4": "",
            },
            {
                "Line Item": "Total Uses",
                "Year 0": total_uses,
                "Year 1": "",
                "Year 2": "",
                "Year 3": "",
                "Year 4": "",
            },
            {
                "Line Item": "SOURCES",
                "Year 0": "",
                "Year 1": "",
                "Year 2": "",
                "Year 3": "",
                "Year 4": "",
            },
            {
                "Line Item": "Senior Debt",
                "Year 0": senior_debt,
                "Year 1": "",
                "Year 2": "",
                "Year 3": "",
                "Year 4": "",
            },
            {
                "Line Item": "Equity Contribution",
                "Year 0": equity_contribution,
                "Year 1": "",
                "Year 2": "",
                "Year 3": "",
                "Year 4": "",
            },
            {
                "Line Item": "Total Sources",
                "Year 0": total_sources,
                "Year 1": "",
                "Year 2": "",
                "Year 3": "",
                "Year 4": "",
            },
        ]
        sources_uses_statement = pd.DataFrame(sources_uses_rows)

        st.markdown("### Sources & Uses")
        _render_custom_table_html(
            sources_uses_statement,
            {"USES", "SOURCES"},
            {"Total Uses", "Total Sources"},
            {},
        )
        if abs(total_sources - total_uses) > 1:
            st.warning(
                "Sources and uses do not reconcile. "
                f"Gap: {format_currency(total_sources - total_uses)}"
            )

        initial_debt = debt_schedule[0]["opening_debt"]
        peak_debt = max(row["opening_debt"] for row in debt_schedule)
        avg_dscr = sum(row["dscr"] for row in debt_schedule) / len(
            debt_schedule
        )
        min_dscr_value = min(row["dscr"] for row in debt_schedule)
        breach_years = [
            f"Year {row['year']}"
            for row in debt_schedule
            if row["covenant_breach"]
        ]
        credit_conclusion = (
            "Not bankable as structured"
            if breach_years
            else "Bankable under base assumptions"
        )

        st.markdown("### Executive Credit Summary")
        credit_summary = pd.DataFrame(
            [
                {"Metric": "Initial Debt (EUR)", "Value": format_currency(initial_debt)},
                {"Metric": "Peak Debt (EUR)", "Value": format_currency(peak_debt)},
                {"Metric": "Average DSCR", "Value": f"{avg_dscr:.2f}x"},
                {"Metric": "Minimum DSCR", "Value": f"{min_dscr_value:.2f}x"},
                {
                    "Metric": "Covenant Breaches",
                    "Value": f"{len(breach_years)} ({', '.join(breach_years)})"
                    if breach_years
                    else "0 (None)",
                },
                {"Metric": "Credit Conclusion", "Value": credit_conclusion},
            ]
        )
        st.dataframe(credit_summary, use_container_width=True)

        def _set_debt_value(name, year_label, value):
            if name not in debt_line_items:
                debt_line_items[name] = {
                    "Line Item": name,
                    "Year 0": "",
                    "Year 1": "",
                    "Year 2": "",
                    "Year 3": "",
                    "Year 4": "",
                }
            debt_line_items[name][year_label] = value

        for row in debt_schedule:
            year_label = f"Year {row['year']}"
            _set_debt_value("Opening Debt", year_label, row["opening_debt"])
            _set_debt_value("Debt Drawdown", year_label, row["debt_drawdown"])
            _set_debt_value(
                "Scheduled Repayment", year_label, row["scheduled_repayment"]
            )
            _set_debt_value(
                "Special Repayment", year_label, row["special_repayment"]
            )
            _set_debt_value(
                "Total Repayment", year_label, row["total_repayment"]
            )
            _set_debt_value("Closing Debt", year_label, row["closing_debt"])
            _set_debt_value(
                "Interest Expense", year_label, row["interest_expense"]
            )

        debt_row_order = [
            "Opening Debt",
            "Debt Drawdown",
            "Scheduled Repayment",
            "Special Repayment",
            "Total Repayment",
            "Closing Debt",
            "Interest Expense",
        ]
        debt_rows = [debt_line_items.get(label) for label in debt_row_order]
        debt_statement = pd.DataFrame(debt_rows)
        st.markdown("### Debt Schedule")
        _render_custom_table_html(
            debt_statement, set(), {"Total Repayment", "Closing Debt"}, {}
        )
        st.caption(
            f"Repayment profile: {financing_assumptions['amortization_type']} "
            f"over {financing_assumptions['amortization_period_years']} years "
            f"with {financing_assumptions['grace_period_years']} years grace. "
            "Interest is calculated on opening debt."
        )

        service_rows = []
        for row in debt_schedule:
            year_label = f"Year {row['year']}"
            ebitda = pnl_result[f"Year {row['year']}"]["ebitda"]
            operating_cf = cashflow_by_year[row["year"]]["operating_cf"]
            capex = cashflow_by_year[row["year"]]["capex"]
            cfads = operating_cf - capex
            debt_service = row["debt_service"]
            dscr = row["dscr"]
            min_dscr = row["minimum_dscr"]
            breach = "YES" if row["covenant_breach"] else "NO"

            service_rows.append(
                {
                    "Line Item": "EBITDA",
                    year_label: ebitda,
                }
            )
            service_rows.append(
                {
                    "Line Item": "Operating Cashflow",
                    year_label: operating_cf,
                }
            )
            service_rows.append(
                {
                    "Line Item": "CFADS",
                    year_label: cfads,
                }
            )
            service_rows.append(
                {
                    "Line Item": "Debt Service",
                    year_label: debt_service,
                }
            )
            service_rows.append(
                {
                    "Line Item": "DSCR",
                    year_label: dscr,
                }
            )
            service_rows.append(
                {
                    "Line Item": "Minimum Required DSCR",
                    year_label: min_dscr,
                }
            )
            service_rows.append(
                {
                    "Line Item": "Covenant Breach",
                    year_label: breach,
                }
            )

        service_metrics = {}
        for entry in service_rows:
            label = entry["Line Item"]
            service_metrics.setdefault(label, {})
            for year_index in range(5):
                year_label = f"Year {year_index}"
                if year_label in entry:
                    service_metrics[label][year_label] = entry[year_label]

        service_table = pd.DataFrame.from_dict(service_metrics, orient="index")
        service_table = service_table[
            [f"Year {i}" for i in range(5)]
        ].reset_index()
        service_table.rename(columns={"index": "Line Item"}, inplace=True)
        service_formatters = {
            "DSCR": lambda value: f"{value:.2f}x"
            if value not in ("", None)
            else "",
            "Minimum Required DSCR": lambda value: f"{value:.2f}x"
            if value not in ("", None)
            else "",
            "Covenant Breach": lambda value: value if value else "",
        }
        st.markdown("### Debt Service & Covenants")
        _render_custom_table_html(
            service_table, set(), {"Debt Service"}, service_formatters
        )

        avg_dscr = sum(row["dscr"] for row in debt_schedule) / len(
            debt_schedule
        )
        min_dscr_value = min(row["dscr"] for row in debt_schedule)
        peak_debt = max(row["opening_debt"] for row in debt_schedule)
        peak_year = max(
            debt_schedule, key=lambda row: row["opening_debt"]
        )["year"]
        ebitda_year0 = pnl_result["Year 0"]["ebitda"]
        debt_at_close = debt_schedule[0]["opening_debt"]
        net_debt = debt_at_close - balance_sheet[0]["cash"]
        debt_to_ebitda = debt_at_close / ebitda_year0 if ebitda_year0 else 0
        net_debt_to_ebitda = (
            net_debt / ebitda_year0 if ebitda_year0 else 0
        )

        st.markdown("### Interpretation")
        if breach_years:
            st.write(
                f"DSCR falls below the minimum in {', '.join(breach_years)}, "
                "driven by repayment intensity and available CFADS."
            )
        else:
            st.write(
                "DSCR remains above the minimum in all years, supported by "
                "stable CFADS and the chosen amortisation profile."
            )
        leverage_ratio = (
            initial_debt / pnl_result["Year 0"]["ebitda"]
            if pnl_result["Year 0"]["ebitda"]
            else 0
        )
        st.caption(
            f"Leverage at close: {leverage_ratio:.2f}x; "
            f"amortisation type: {financing_assumptions['amortization_type']}."
        )
        st.caption(f"Bank view conclusion: {credit_conclusion}.")

        kpi_table = pd.DataFrame(
            [
                {"KPI": "Average DSCR", "Value": f"{avg_dscr:.2f}x"},
                {"KPI": "Minimum DSCR", "Value": f"{min_dscr_value:.2f}x"},
                {
                    "KPI": "Peak Debt",
                    "Value": f"{format_currency(peak_debt)} (Year {peak_year})",
                },
                {
                    "KPI": "Debt / EBITDA (at close)",
                    "Value": f"{debt_to_ebitda:.2f}x",
                },
                {
                    "KPI": "Net Debt / EBITDA",
                    "Value": f"{net_debt_to_ebitda:.2f}x",
                },
            ]
        )
        st.markdown("### KPIs")
        st.dataframe(kpi_table, use_container_width=True)

        equity_cashflows = investment_result["equity_cashflows"]
        total_equity_invested = investment_result["initial_equity"]
        total_distributions = sum(cf for cf in equity_cashflows if cf > 0)
        cash_on_cash = (
            total_distributions / abs(total_equity_invested)
            if total_equity_invested
            else 0
        )

        st.markdown("### Investor Financing View")
        ebitda_year0 = pnl_result["Year 0"]["ebitda"]
        entry_ev = purchase_price + transaction_fees
        entry_multiple = (
            entry_ev / ebitda_year0 if ebitda_year0 else 0
        )
        ownership_pct = (
            equity_contribution / (equity_contribution + senior_debt)
            if equity_contribution + senior_debt
            else 0
        )
        net_debt_at_close = senior_debt - balance_sheet[0]["cash"]
        investor_summary = pd.DataFrame(
            [
                {
                    "Metric": "Equity Contribution",
                    "Value": format_currency(equity_contribution),
                },
                {
                    "Metric": "Ownership %",
                    "Value": format_pct(ownership_pct),
                },
                {
                    "Metric": "Entry Multiple (EV / EBITDA)",
                    "Value": f"{entry_multiple:.2f}x",
                },
                {
                    "Metric": "Debt at Close",
                    "Value": format_currency(senior_debt),
                },
                {
                    "Metric": "Net Debt at Close",
                    "Value": format_currency(net_debt_at_close),
                },
            ]
        )
        st.dataframe(investor_summary, use_container_width=True)
        st.caption(
            "Investor provides the residual capital after bank funding "
            "and minimum liquidity needs are satisfied."
        )

        bridge_rows = []
        cash_to_equity = equity_cashflows[1:]
        for year_index in range(5):
            year_label = f"Year {year_index}"
            free_cf = cashflow_by_year[year_index]["free_cashflow"]
            debt_service = (
                debt_schedule[year_index]["interest_expense"]
                + debt_schedule[year_index]["total_repayment"]
            )
            equity_cf = (
                cash_to_equity[year_index]
                if year_index < len(cash_to_equity)
                else 0
            )
            mandatory_retention = free_cf - debt_service - equity_cf
            bridge_rows.append(
                {
                    "Line Item": "Free Cashflow",
                    year_label: free_cf,
                }
            )
            bridge_rows.append(
                {
                    "Line Item": "Debt Service",
                    year_label: debt_service,
                }
            )
            bridge_rows.append(
                {
                    "Line Item": "Mandatory Cash Retention",
                    year_label: mandatory_retention,
                }
            )
            bridge_rows.append(
                {
                    "Line Item": "Cash to Equity",
                    year_label: equity_cf,
                }
            )

        bridge_metrics = {}
        for entry in bridge_rows:
            label = entry["Line Item"]
            bridge_metrics.setdefault(label, {})
            for year_index in range(5):
                year_label = f"Year {year_index}"
                if year_label in entry:
                    bridge_metrics[label][year_label] = entry[year_label]

        bridge_table = pd.DataFrame.from_dict(bridge_metrics, orient="index")
        bridge_table = bridge_table[
            [f"Year {i}" for i in range(5)]
        ].reset_index()
        bridge_table.rename(columns={"index": "Line Item"}, inplace=True)

        st.markdown("### Cashflow Available to Equity")
        _render_custom_table_html(
            bridge_table,
            set(),
            {"Cash to Equity"},
            {},
        )

        target_irr = financing_assumptions["target_irr"]
        max_equity = financing_assumptions["max_equity_contribution_eur"]
        min_cash_yield = financing_assumptions["min_cash_yield"]
        average_cash_distribution = (
            sum(cf for cf in cash_to_equity if cf > 0) / 5
            if cash_to_equity
            else 0
        )
        cash_yield = (
            average_cash_distribution / equity_contribution
            if equity_contribution
            else 0
        )
        equity_stress_years = [
            f"Year {i}"
            for i, value in enumerate(cash_to_equity)
            if value < 0
        ]
        meets_target_irr = (
            investment_result["irr"] >= target_irr
            if target_irr is not None
            else True
        )
        meets_max_equity = (
            equity_contribution <= max_equity
            if max_equity
            else True
        )
        meets_cash_yield = (
            cash_yield >= min_cash_yield
            if min_cash_yield is not None
            else True
        )

        st.markdown("### Investor Constraints")
        investor_constraints = pd.DataFrame(
            [
                {
                    "Metric": "Target IRR",
                    "Value": format_pct(target_irr),
                    "Status": "YES" if meets_target_irr else "NO",
                },
                {
                    "Metric": "Max Equity Contribution",
                    "Value": format_currency(max_equity)
                    if max_equity
                    else "N/A",
                    "Status": "YES" if meets_max_equity else "NO",
                },
                {
                    "Metric": "Min Cash Yield",
                    "Value": format_pct(min_cash_yield),
                    "Status": "YES" if meets_cash_yield else "NO",
                },
                {
                    "Metric": "Equity Stress Year(s)",
                    "Value": ", ".join(equity_stress_years)
                    if equity_stress_years
                    else "None",
                    "Status": "",
                },
            ]
        )
        st.dataframe(investor_constraints, use_container_width=True)

        st.markdown("### Leverage vs Investor Returns")
        st.write(
            "Higher leverage can lift IRR but reduces DSCR headroom; "
            f"current minimum DSCR is {min_dscr_value:.2f}x versus a "
            f"{financing_assumptions['minimum_dscr']:.2f}x covenant. "
            f"Equity IRR is {format_pct(investment_result['irr'])}."
        )

        explain_financing = st.toggle("Explain Financing Logic")
        if explain_financing:
            st.markdown("### 1. Transaction Funding (Sources & Uses)")
            st.write(
                "Purchase price, fees, and required cash at close are funded "
                "through senior debt and equity. Sources must match uses."
            )
            st.markdown("### 2. Bank Debt Mechanics & Covenants")
            st.write(
                f"Initial debt of {format_currency(initial_debt)} is amortised "
                f"{financing_assumptions['amortization_type'].lower()} over "
                f"{financing_assumptions['amortization_period_years']} years "
                f"with {financing_assumptions['grace_period_years']} years grace. "
                f"Minimum DSCR is {financing_assumptions['minimum_dscr']:.2f}x."
            )
            st.markdown("### 3. Residual Equity Funding")
            st.write(
                "Equity funds the residual capital after debt capacity and "
                "liquidity requirements are satisfied."
            )
            st.markdown("### 4. Cashflow Split: Bank vs Investor")
            st.write(
                "Operating cashflow less capex defines CFADS; debt service "
                "is paid first, and remaining cash is available to equity."
            )
            st.markdown("### 5. Resulting Investor Returns")
            st.write(
                f"Equity IRR is {format_pct(investment_result['irr'])}, with "
                "returns sensitive to leverage and covenant headroom."
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
                "initial_equity": "Eigenkapital (Start, m EUR)",
                "exit_value": "Exit Value (m EUR)",
                "irr": "IRR (%)",
            },
            inplace=True,
        )
        summary_format_map = {
            "Eigenkapital (Start, m EUR)": format_currency,
            "Exit Value (m EUR)": format_currency,
            "IRR (%)": format_pct,
        }
        summary_totals = ["Eigenkapital (Start, m EUR)", "Exit Value (m EUR)"]
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
                "equity_cashflows": "Equity Cashflows (m EUR)",
            },
            inplace=True,
        )
        cashflows_format_map = {
            "Equity Cashflows (m EUR)": format_currency
        }
        cashflows_styled = _style_totals(
            cashflows_display, ["Equity Cashflows (m EUR)"]
        ).format(cashflows_format_map)

        st.dataframe(summary_styled, use_container_width=True)
        st.dataframe(cashflows_styled, use_container_width=True)


if __name__ == "__main__":
    run_app()
