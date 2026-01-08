import io
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

st.set_page_config(layout="wide")

from data_model import InputModel, create_demo_input_model
import run_model as run_model
from calculations.investment import _calculate_irr


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
    cashflow_defaults = _default_cashflow_assumptions()
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
        "maintenance_capex_pct_revenue": cashflow_defaults[
            "capex_pct_revenue"
        ],
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


def _default_equity_assumptions(input_model):
    total_equity = input_model.transaction_and_financing[
        "equity_contribution_eur"
    ].value
    default_multiple = (
        input_model.valuation_assumptions["multiple_valuation"][
            "seller_multiple"
        ].value
        or 0.0
    )
    return {
        "sponsor_equity_eur": total_equity * 0.5,
        "investor_equity_eur": total_equity * 0.5,
        "exit_year": 4,
        "exit_method": "Exit Multiple",
        "exit_multiple": default_multiple,
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

    for key, value in _default_valuation_assumptions(input_model).items():
        st.session_state.setdefault(f"valuation.{key}", value)

    for key, value in _default_equity_assumptions(input_model).items():
        st.session_state.setdefault(f"equity.{key}", value)


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
    statement,
    section_rows,
    bold_rows,
    row_formatters=None,
    year_labels=None,
):
    def escape(text):
        return (
            str(text)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

    if year_labels is None:
        year_labels = ["Year 0", "Year 1", "Year 2", "Year 3", "Year 4"]
    columns = ["Line Item"] + year_labels
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

    line_item_width = 35
    year_width = (100 - line_item_width) / max(len(year_labels), 1)
    css = f"""
    <style>
      .custom-table {{ width: 100%; border-collapse: collapse; table-layout: fixed; }}
      .custom-table col.line-item {{ width: {line_item_width}%; }}
      .custom-table col.year {{ width: {year_width:.2f}%; }}
      .custom-table th, .custom-table td {{
        padding: 2px 6px;
        white-space: nowrap;
        line-height: 1.0;
        border: 0;
        font-size: 0.9rem;
      }}
      .custom-table th {{ text-align: right; font-weight: 600; }}
      .custom-table th:first-child {{ text-align: left; }}
      .custom-table td {{ text-align: right; }}
      .custom-table td:first-child {{ text-align: left; }}
      .custom-table .section-row td {{
        font-weight: 700;
        background: #f9fafb;
      }}
      .custom-table .total-row td {{
        font-weight: 700;
        background: #f3f4f6;
        border-top: 1px solid #c7c7c7;
      }}
      .custom-table td.negative {{ color: #b45309; }}
    </style>
    """
    colgroup = "<colgroup><col class=\"line-item\"/>"
    colgroup += "".join("<col class=\"year\"/>" for _ in year_labels)
    colgroup += "</colgroup>"
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
    equity_assumptions = getattr(
        input_model,
        "equity_assumptions",
        _default_equity_assumptions(input_model),
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
        ("Maintenance Capex (% of Revenue)", financing_assumptions["maintenance_capex_pct_revenue"]),
        ("Minimum DSCR", financing_assumptions["minimum_dscr"]),
        ("Amortisation Period (Years)", financing_assumptions["amortization_period_years"]),
        ("Working Capital Adjustment (% of Revenue)", cashflow_assumptions["working_capital_pct_revenue"]),
        ("Opening Cash Balance (EUR)", cashflow_assumptions["opening_cash_balance_eur"]),
        ("Sponsor Equity (EUR)", equity_assumptions["sponsor_equity_eur"]),
        ("Investor Equity (EUR)", equity_assumptions["investor_equity_eur"]),
        ("Exit Year", equity_assumptions["exit_year"]),
        ("Exit Method", equity_assumptions["exit_method"]),
        ("Exit Multiple (x)", equity_assumptions["exit_multiple"]),
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
        ws_equity = wb.create_sheet("Equity Case")

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
            ("Initial Debt Amount (EUR)", f"={assumption_cells['Debt Amount (EUR)']}"),
            ("Interest Rate", f"={assumption_cells['Interest Rate %']}"),
            ("Amortisation Type", "Linear"),
            ("Amortisation Period (Years)", f"={assumption_cells['Amortisation Period (Years)']}"),
            ("Minimum DSCR", f"={assumption_cells['Minimum DSCR']}"),
            ("Maintenance Capex (% of Revenue)", f"={assumption_cells['Maintenance Capex (% of Revenue)']}"),
            ("Working Capital Change (% of Revenue)", "=0"),
            ("Tax Cash Rate (%)", f"={assumption_cells['Tax Cash Rate (%)']}"),
            ("Tax Payment Lag (Years)", f"={assumption_cells['Tax Payment Lag (Years)']}"),
        ]
        for idx, (label, value) in enumerate(financing_rows, start=2):
            ws_financing.cell(row=idx, column=1, value=label)
            ws_financing.cell(row=idx, column=2, value=value)

        debt_table_start = len(financing_rows) + 4
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
        initial_debt_cell = "B2"

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
                f"=IF({year_index}<{amort_period_cell},{initial_debt_cell}/{amort_period_cell},0)"
            )
            ws_financing[f"{col}{debt_row_map['Special Repayment']}"] = (
                "=0"
            )
            ws_financing[f"{col}{debt_row_map['Total Repayment']}"] = (
                f"=MIN({col}{debt_row_map['Opening Debt']},{col}{debt_row_map['Scheduled Repayment']})"
            )
            ws_financing[f"{col}{debt_row_map['Closing Debt']}"] = (
                f"={col}{debt_row_map['Opening Debt']}-{col}{debt_row_map['Total Repayment']}"
            )
            ws_financing[f"{col}{debt_row_map['Interest Expense']}"] = (
                f"={col}{debt_row_map['Opening Debt']}*{interest_rate_cell}"
            )

        bank_table_start = debt_table_start + len(debt_items) + 3
        ws_financing.cell(row=bank_table_start, column=1, value="Bank View")
        ws_financing.cell(row=bank_table_start + 1, column=1, value="Line Item")
        for idx, header in enumerate(year_headers, start=2):
            ws_financing.cell(row=bank_table_start + 1, column=idx, value=header)

        bank_items = [
            "EBITDA",
            "Cash Taxes",
            "Capex (Maintenance)",
            "CFADS",
            "Interest Expense",
            "Scheduled Repayment",
            "Debt Service",
            "DSCR",
            "Minimum Required DSCR",
            "Covenant Breach",
        ]
        for row_idx, item in enumerate(bank_items, start=bank_table_start + 2):
            ws_financing.cell(row=row_idx, column=1, value=item)

        bank_row_map = {
            "EBITDA": bank_table_start + 2,
            "Cash Taxes": bank_table_start + 3,
            "Capex (Maintenance)": bank_table_start + 4,
            "CFADS": bank_table_start + 5,
            "Interest Expense": bank_table_start + 6,
            "Scheduled Repayment": bank_table_start + 7,
            "Debt Service": bank_table_start + 8,
            "DSCR": bank_table_start + 9,
            "Minimum Required DSCR": bank_table_start + 10,
            "Covenant Breach": bank_table_start + 11,
        }

        min_dscr_cell = "B6"
        maintenance_capex_cell = "B7"
        wc_change_cell = "B8"

        for year_index in range(5):
            col = year_col(2 + year_index)
            ws_financing[f"{col}{bank_row_map['EBITDA']}"] = f"='P&L'!{col}17"
            ws_financing[f"{col}{bank_row_map['Cash Taxes']}"] = (
                f"=Cashflow!{col}{cashflow_row_map['Taxes Paid']}"
            )
            ws_financing[f"{col}{bank_row_map['Capex (Maintenance)']}"] = (
                f"='P&L'!{col}5*{maintenance_capex_cell}"
            )
            ws_financing[f"{col}{bank_row_map['CFADS']}"] = (
                f"={col}{bank_row_map['EBITDA']}"
                f"-{col}{bank_row_map['Cash Taxes']}"
                f"-{col}{bank_row_map['Capex (Maintenance)']}"
                f"+('P&L'!{col}5*{wc_change_cell})"
            )
            ws_financing[f"{col}{bank_row_map['Interest Expense']}"] = (
                f"={col}{debt_row_map['Interest Expense']}"
            )
            ws_financing[f"{col}{bank_row_map['Scheduled Repayment']}"] = (
                f"={col}{debt_row_map['Scheduled Repayment']}"
            )
            ws_financing[f"{col}{bank_row_map['Debt Service']}"] = (
                f"={col}{bank_row_map['Interest Expense']}"
                f"+{col}{bank_row_map['Scheduled Repayment']}"
            )
            ws_financing[f"{col}{bank_row_map['DSCR']}"] = (
                f"=IF({col}{bank_row_map['Debt Service']}=0,0,"
                f"{col}{bank_row_map['CFADS']}/{col}{bank_row_map['Debt Service']})"
            )
            ws_financing[f"{col}{bank_row_map['Minimum Required DSCR']}"] = (
                f"={min_dscr_cell}"
            )
            ws_financing[f"{col}{bank_row_map['Covenant Breach']}"] = (
                f"=IF({col}{bank_row_map['DSCR']}<{min_dscr_cell},\"YES\",\"NO\")"
            )

        financing_notes_row = bank_table_start + len(bank_items) + 2
        ws_financing.cell(row=financing_notes_row, column=1, value="Notes")
        ws_financing.cell(
            row=financing_notes_row + 1,
            column=1,
            value="CFADS = EBITDA - Cash Taxes - Maintenance Capex ± Working Capital Change.",
        )
        ws_financing.cell(
            row=financing_notes_row + 2,
            column=1,
            value="Debt Service = Interest Expense + Scheduled Repayment.",
        )
        ws_financing.cell(
            row=financing_notes_row + 3,
            column=1,
            value="DSCR = CFADS / Debt Service.",
        )

        ws_equity["A1"] = "Equity Assumptions"
        equity_rows = [
            ("Sponsor Equity (EUR)", "Sponsor Equity (EUR)"),
            ("Investor Equity (EUR)", "Investor Equity (EUR)"),
            ("Exit Year", "Exit Year"),
            ("Exit Method", "Exit Method"),
            ("Exit Multiple (x)", "Exit Multiple (x)"),
        ]
        for idx, (label, key) in enumerate(equity_rows, start=2):
            ws_equity.cell(row=idx, column=1, value=label)
            ws_equity.cell(
                row=idx, column=2, value=f"={assumption_cells[key]}"
            )

        equity_calc_start = len(equity_rows) + 4
        ws_equity.cell(row=equity_calc_start, column=1, value="Equity Summary")
        ws_equity.cell(row=equity_calc_start + 1, column=1, value="Metric")
        ws_equity.cell(row=equity_calc_start + 1, column=2, value="Value")

        sponsor_cell = "B2"
        investor_cell = "B3"
        exit_year_cell = "B4"
        exit_method_cell = "B5"
        exit_multiple_cell = "B6"

        ws_equity.cell(row=equity_calc_start + 2, column=1, value="Total Equity")
        ws_equity.cell(
            row=equity_calc_start + 2, column=2, value=f"={sponsor_cell}+{investor_cell}"
        )
        ws_equity.cell(row=equity_calc_start + 3, column=1, value="Sponsor %")
        ws_equity.cell(
            row=equity_calc_start + 3,
            column=2,
            value=f"=IF(B{equity_calc_start + 2}=0,0,{sponsor_cell}/B{equity_calc_start + 2})",
        )
        ws_equity.cell(row=equity_calc_start + 4, column=1, value="Investor %")
        ws_equity.cell(
            row=equity_calc_start + 4,
            column=2,
            value=f"=IF(B{equity_calc_start + 2}=0,0,{investor_cell}/B{equity_calc_start + 2})",
        )
        ws_equity.cell(row=equity_calc_start + 5, column=1, value="Exit EBITDA")
        ws_equity.cell(
            row=equity_calc_start + 5,
            column=2,
            value=f"=INDEX('P&L'!B17:F17,1,MIN({exit_year_cell},4)+1)",
        )
        ws_equity.cell(row=equity_calc_start + 6, column=1, value="Net Debt at Exit")
        ws_equity.cell(
            row=equity_calc_start + 6,
            column=2,
            value=f"=INDEX('Balance Sheet'!B14:F14,1,MIN({exit_year_cell},4)+1)-INDEX('Balance Sheet'!B10:F10,1,MIN({exit_year_cell},4)+1)",
        )
        ws_equity.cell(row=equity_calc_start + 7, column=1, value="Enterprise Value Exit")
        ws_equity.cell(
            row=equity_calc_start + 7,
            column=2,
            value=f"={exit_multiple_cell}*B{equity_calc_start + 5}",
        )
        ws_equity.cell(row=equity_calc_start + 8, column=1, value="Equity Value Exit")
        ws_equity.cell(
            row=equity_calc_start + 8,
            column=2,
            value=f"=B{equity_calc_start + 7}-B{equity_calc_start + 6}",
        )

        cashflow_start = equity_calc_start + 11
        ws_equity.cell(row=cashflow_start, column=1, value="Equity Cashflows")
        ws_equity.cell(row=cashflow_start + 1, column=1, value="Line Item")
        equity_year_headers = [
            "Year 0",
            "Year 1",
            "Year 2",
            "Year 3",
            "Year 4",
            "Year 5",
            "Year 6",
            "Year 7",
        ]
        for idx, header in enumerate(equity_year_headers, start=2):
            ws_equity.cell(row=cashflow_start + 1, column=idx, value=header)

        ws_equity.cell(row=cashflow_start + 2, column=1, value="Sponsor Cashflow")
        ws_equity.cell(row=cashflow_start + 3, column=1, value="Investor Cashflow")
        ws_equity.cell(row=cashflow_start + 4, column=1, value="Sponsor Residual Equity Value")

        sponsor_pct_cell = f"B{equity_calc_start + 3}"
        investor_pct_cell = f"B{equity_calc_start + 4}"
        equity_value_exit_cell = f"B{equity_calc_start + 8}"

        for year_index in range(8):
            col = year_col(2 + year_index)
            ws_equity[f"{col}{cashflow_start + 2}"] = (
                f"=IF({year_index}=0,-{sponsor_cell},"
                f"IF({year_index}={exit_year_cell},{equity_value_exit_cell},0))"
            )
            ws_equity[f"{col}{cashflow_start + 3}"] = (
                f"=IF({year_index}=0,-{investor_cell},"
                f"IF({year_index}={exit_year_cell},{equity_value_exit_cell}*{investor_pct_cell},0))"
            )
            ws_equity[f"{col}{cashflow_start + 4}"] = (
                f"=IF({year_index}={exit_year_cell},{equity_value_exit_cell},0)"
            )

        kpi_start = cashflow_start + 6
        ws_equity.cell(row=kpi_start, column=1, value="Equity KPIs")
        ws_equity.cell(row=kpi_start + 1, column=1, value="Investor")
        ws_equity.cell(row=kpi_start + 1, column=2, value="Invested Equity")
        ws_equity.cell(row=kpi_start + 1, column=3, value="Exit Proceeds")
        ws_equity.cell(row=kpi_start + 1, column=4, value="MOIC")
        ws_equity.cell(row=kpi_start + 1, column=5, value="IRR")

        ws_equity.cell(row=kpi_start + 2, column=1, value="Sponsor")
        ws_equity.cell(row=kpi_start + 3, column=1, value="Investor")

        ws_equity.cell(row=kpi_start + 2, column=2, value=f"={sponsor_cell}")
        ws_equity.cell(row=kpi_start + 3, column=2, value=f"={investor_cell}")
        ws_equity.cell(
            row=kpi_start + 2,
            column=3,
            value=f"={equity_value_exit_cell}",
        )
        ws_equity.cell(
            row=kpi_start + 3,
            column=3,
            value=f"={equity_value_exit_cell}*{investor_pct_cell}",
        )
        ws_equity.cell(
            row=kpi_start + 2,
            column=4,
            value="=\"—\"",
        )
        ws_equity.cell(
            row=kpi_start + 3,
            column=4,
            value=f"=IF({investor_cell}=0,0,C{kpi_start + 3}/{investor_cell})",
        )
        ws_equity.cell(
            row=kpi_start + 2,
            column=5,
            value=f"=IRR(B{cashflow_start + 2}:I{cashflow_start + 2})",
        )
        ws_equity.cell(
            row=kpi_start + 3,
            column=5,
            value=f"=IRR(B{cashflow_start + 3}:I{cashflow_start + 3})",
        )

        writer.close()

    output.seek(0)
    return output


def run_app():
    st.title("Financial Model")

    base_model = create_demo_input_model()
    if not st.session_state.get("defaults_initialized"):
        _seed_session_defaults(base_model)
        st.session_state["defaults_initialized"] = True

    # Navigation for question-driven layout.
    st.session_state.setdefault("current_page", "Overview")
    with st.sidebar:
        nav_css = """
        <style>
          .nav-section {
            font-size: 0.7rem;
            letter-spacing: 0.12em;
            text-transform: uppercase;
            color: #6b7280;
            margin-top: 0.9rem;
            margin-bottom: 0.35rem;
          }
          .nav-item {
            padding: 0.45rem 0.6rem;
            border-radius: 6px;
            margin-bottom: 0.25rem;
            font-weight: 500;
            color: #111827;
          }
          .nav-item.active {
            background: #eef2ff;
            border: 1px solid #c7d2fe;
          }
          div[data-testid="stSidebar"] button {
            width: 100%;
            justify-content: flex-start;
            border-radius: 6px;
            padding: 0.45rem 0.6rem;
          }
          div[data-testid="stSidebar"] button:hover {
            background: #f3f4f6;
            border-color: #e5e7eb;
          }
        </style>
        """
        st.markdown(nav_css, unsafe_allow_html=True)
        editor_css = """
        <style>
          .rdg-cell[aria-readonly="true"] {
            background: #f3f4f6;
            color: #6b7280;
            font-style: italic;
          }
          .rdg-cell[aria-readonly="false"] {
            background: #ffffff;
            border: 1px solid #93c5fd;
          }
        </style>
        """
        st.markdown(editor_css, unsafe_allow_html=True)

        def _nav_item(label):
            if st.session_state["current_page"] == label:
                st.markdown(
                    f"<div class=\"nav-item active\">{label}</div>",
                    unsafe_allow_html=True,
                )
            else:
                if st.button(
                    label,
                    key=f"nav_{label}",
                    use_container_width=True,
                ):
                    st.session_state["current_page"] = label

        st.markdown("<div class=\"nav-section\">OVERVIEW</div>", unsafe_allow_html=True)
        _nav_item("Overview")

        st.markdown("<div class=\"nav-section\">OPERATING MODEL</div>", unsafe_allow_html=True)
        _nav_item("Operating Model (P&L)")
        _nav_item("Cashflow & Liquidity")
        _nav_item("Balance Sheet")

        st.markdown("<div class=\"nav-section\">FINANCING</div>", unsafe_allow_html=True)
        _nav_item("Financing & Debt")
        _nav_item("Equity Case")

        st.markdown("<div class=\"nav-section\">VALUATION</div>", unsafe_allow_html=True)
        _nav_item("Valuation & Purchase Price")

        st.markdown("<div class=\"nav-section\">SETTINGS</div>", unsafe_allow_html=True)
        _nav_item("Assumptions (Advanced)")
        _nav_item("Settings")

        page = st.session_state["current_page"]
        assumptions_state = st.session_state["assumptions"]

        def _sidebar_editor(title, key, df, column_config):
            st.markdown(f"### {title}")
            edited = st.data_editor(
                df,
                hide_index=True,
                key=key,
                column_config=column_config,
                use_container_width=True,
            )
            return edited

        if page == "Operating Model (P&L)":
            revenue_df = pd.DataFrame(assumptions_state["revenue_drivers"])
            auto_sync = st.session_state.get("assumptions.auto_sync", True)
            base_label = (
                "Base (Active)" if selected_scenario == "Base" else "Base"
            )
            best_label = (
                "Best (Active)" if selected_scenario == "Best" else "Best"
            )
            worst_label = (
                "Worst (Active)" if selected_scenario == "Worst" else "Worst"
            )
            revenue_df = revenue_df.rename(
                columns={
                    "Base": base_label,
                    "Best": best_label,
                    "Worst": worst_label,
                }
            )
            edited_revenue = _sidebar_editor(
                "Revenue Drivers",
                "sidebar.revenue_drivers",
                revenue_df,
                {
                    "Parameter": st.column_config.TextColumn(disabled=True),
                    "Unit": st.column_config.TextColumn(disabled=True),
                    "Description": st.column_config.TextColumn(disabled=True),
                    base_label: st.column_config.NumberColumn(
                        disabled=auto_sync and selected_scenario != "Base"
                    ),
                    best_label: st.column_config.NumberColumn(
                        disabled=auto_sync and selected_scenario != "Best"
                    ),
                    worst_label: st.column_config.NumberColumn(
                        disabled=auto_sync and selected_scenario != "Worst"
                    ),
                },
            )
            edited_revenue = edited_revenue.rename(
                columns={
                    base_label: "Base",
                    best_label: "Best",
                    worst_label: "Worst",
                }
            )
            assumptions_state["revenue_drivers"] = edited_revenue.to_dict("records")

            guarantees_df = pd.DataFrame(assumptions_state["revenue_guarantees"])
            edited_guarantees = _sidebar_editor(
                "Revenue Guarantees",
                "sidebar.revenue_guarantees",
                guarantees_df,
                {
                    "Year": st.column_config.TextColumn(disabled=True),
                    "Description": st.column_config.TextColumn(disabled=True),
                },
            )
            assumptions_state["revenue_guarantees"] = edited_guarantees.to_dict("records")
            _apply_assumptions_state()

        if page == "Cashflow & Liquidity":
            cashflow_df = pd.DataFrame(assumptions_state["cashflow"])
            edited_cashflow = _sidebar_editor(
                "Cashflow Assumptions",
                "sidebar.cashflow",
                cashflow_df,
                {
                    "Parameter": st.column_config.TextColumn(disabled=True),
                    "Unit": st.column_config.TextColumn(disabled=True),
                    "Notes": st.column_config.TextColumn(disabled=True),
                },
            )
            assumptions_state["cashflow"] = edited_cashflow.to_dict("records")
            _apply_assumptions_state()

        if page == "Balance Sheet":
            balance_df = pd.DataFrame(assumptions_state["balance_sheet"])
            edited_balance = _sidebar_editor(
                "Balance Sheet Assumptions",
                "sidebar.balance_sheet",
                balance_df,
                {
                    "Parameter": st.column_config.TextColumn(disabled=True),
                    "Unit": st.column_config.TextColumn(disabled=True),
                    "Notes": st.column_config.TextColumn(disabled=True),
                },
            )
            assumptions_state["balance_sheet"] = edited_balance.to_dict("records")
            _apply_assumptions_state()

        if page == "Financing & Debt":
            financing_df = pd.DataFrame(assumptions_state["financing"])
            edited_financing = _sidebar_editor(
                "Financing Assumptions",
                "sidebar.financing",
                financing_df,
                {
                    "Parameter": st.column_config.TextColumn(disabled=True),
                    "Unit": st.column_config.TextColumn(disabled=True),
                    "Notes": st.column_config.TextColumn(disabled=True),
                },
            )
            assumptions_state["financing"] = edited_financing.to_dict("records")
            _apply_assumptions_state()

        if page == "Valuation & Purchase Price":
            valuation_df = pd.DataFrame(assumptions_state["valuation"])
            edited_valuation = _sidebar_editor(
                "Valuation Assumptions",
                "sidebar.valuation",
                valuation_df,
                {
                    "Parameter": st.column_config.TextColumn(disabled=True),
                    "Unit": st.column_config.TextColumn(disabled=True),
                    "Notes": st.column_config.TextColumn(disabled=True),
                },
            )
            assumptions_state["valuation"] = edited_valuation.to_dict("records")
            _apply_assumptions_state()

        if page == "Equity Case":
            equity_df = pd.DataFrame(assumptions_state["equity"])
            edited_equity = _sidebar_editor(
                "Equity Assumptions",
                "sidebar.equity",
                equity_df,
                {
                    "Parameter": st.column_config.TextColumn(disabled=True),
                    "Unit": st.column_config.TextColumn(disabled=True),
                    "Notes": st.column_config.TextColumn(disabled=True),
                },
            )
            assumptions_state["equity"] = edited_equity.to_dict("records")
            _apply_assumptions_state()

    # Build input model and collect editable values from the assumptions page.
    selected_scenario = st.session_state.get(
        "scenario_selection.selected_scenario",
        base_model.scenario_selection["selected_scenario"].value,
    )
    scenario_key = selected_scenario.lower()

    def _seed_assumptions_state():
        return {
            "revenue_drivers": [
                {
                    "Parameter": "Consulting FTE",
                    "Unit": "FTE",
                    "Base": base_model.operating_assumptions["consulting_fte_start"].value,
                    "Best": base_model.operating_assumptions["consulting_fte_start"].value,
                    "Worst": base_model.operating_assumptions["consulting_fte_start"].value,
                    "Description": "Starting consulting headcount.",
                },
                {
                    "Parameter": "Workdays per Year",
                    "Unit": "Days",
                    "Base": base_model.operating_assumptions["work_days_per_year"].value,
                    "Best": base_model.operating_assumptions["work_days_per_year"].value,
                    "Worst": base_model.operating_assumptions["work_days_per_year"].value,
                    "Description": "Standard workdays per consultant.",
                },
                {
                    "Parameter": "Utilization (%)",
                    "Unit": "%",
                    "Base": base_model.scenario_parameters["utilization_rate"]["base"].value,
                    "Best": base_model.scenario_parameters["utilization_rate"]["best"].value,
                    "Worst": base_model.scenario_parameters["utilization_rate"]["worst"].value,
                    "Description": "Billable utilization rate.",
                },
                {
                    "Parameter": "Day Rate (EUR)",
                    "Unit": "EUR",
                    "Base": base_model.scenario_parameters["day_rate_eur"]["base"].value,
                    "Best": base_model.scenario_parameters["day_rate_eur"]["best"].value,
                    "Worst": base_model.scenario_parameters["day_rate_eur"]["worst"].value,
                    "Description": "Base daily billing rate.",
                },
                {
                    "Parameter": "Day Rate Growth (% p.a.)",
                    "Unit": "%",
                    "Base": base_model.operating_assumptions["day_rate_growth_pct"].value,
                    "Best": base_model.operating_assumptions["day_rate_growth_pct"].value,
                    "Worst": base_model.operating_assumptions["day_rate_growth_pct"].value,
                    "Description": "Annual pricing growth.",
                },
            ],
            "revenue_guarantees": [
                {"Year": "Year 1", "Guarantee %": base_model.operating_assumptions["revenue_guarantee_pct_year_1"].value, "Description": "Guaranteed share of revenue in Year 1."},
                {"Year": "Year 2", "Guarantee %": base_model.operating_assumptions["revenue_guarantee_pct_year_2"].value, "Description": "Guaranteed share of revenue in Year 2."},
                {"Year": "Year 3", "Guarantee %": base_model.operating_assumptions["revenue_guarantee_pct_year_3"].value, "Description": "Guaranteed share of revenue in Year 3."},
            ],
            "personnel_costs": [
                {"Role": "Consultant Base Salary", "Cost Type": "Fixed", "Base Value (EUR)": base_model.personnel_cost_assumptions["avg_consultant_base_cost_eur_per_year"].value, "Growth (%)": base_model.personnel_cost_assumptions["wage_inflation_pct"].value, "Notes": "Base salary per consultant."},
                {"Role": "Consultant Variable (% Revenue)", "Cost Type": "Percent of Base", "Base Value (EUR)": base_model.personnel_cost_assumptions["bonus_pct_of_base"].value, "Growth (%)": "", "Notes": "Bonus as % of base salary."},
                {"Role": "Backoffice Cost per FTE", "Cost Type": "Fixed", "Base Value (EUR)": base_model.operating_assumptions["avg_backoffice_salary_eur_per_year"].value, "Growth (%)": base_model.personnel_cost_assumptions["wage_inflation_pct"].value, "Notes": "Average backoffice salary."},
                {"Role": "Management / MD Cost", "Cost Type": "Fixed", "Base Value (EUR)": 0.0, "Growth (%)": "", "Notes": "Not modeled in v1."},
            ],
            "opex": [
                {"Category": "External Consulting", "Cost Type": "Fixed", "Value": base_model.overhead_and_variable_costs["legal_audit_eur_per_year"].value, "Unit": "EUR", "Notes": "External advisors."},
                {"Category": "IT", "Cost Type": "Fixed", "Value": base_model.overhead_and_variable_costs["it_and_software_eur_per_year"].value, "Unit": "EUR", "Notes": "IT and software."},
                {"Category": "Office", "Cost Type": "Fixed", "Value": base_model.overhead_and_variable_costs["rent_eur_per_year"].value, "Unit": "EUR", "Notes": "Office rent."},
                {"Category": "Other Services", "Cost Type": "Fixed", "Value": base_model.overhead_and_variable_costs["other_overhead_eur_per_year"].value, "Unit": "EUR", "Notes": "Other services (excludes insurance)."},
            ],
            "financing": [
                {"Parameter": "Senior Debt Amount", "Value": base_model.transaction_and_financing["senior_term_loan_start_eur"].value, "Unit": "EUR", "Notes": "Opening senior term loan."},
                {"Parameter": "Interest Rate", "Value": base_model.transaction_and_financing["senior_interest_rate_pct"].value, "Unit": "%", "Notes": "Fixed interest rate."},
                {"Parameter": "Amortisation Years", "Value": _default_financing_assumptions(base_model)["amortization_period_years"], "Unit": "Years", "Notes": "Linear amortisation period."},
                {"Parameter": "Transaction Fees (%)", "Value": _default_valuation_assumptions(base_model)["transaction_cost_pct"], "Unit": "%", "Notes": "Fees as % of EV."},
            ],
            "equity": [
                {"Parameter": "Sponsor Equity Contribution", "Value": _default_equity_assumptions(base_model)["sponsor_equity_eur"], "Unit": "EUR", "Notes": "Management equity contribution."},
                {"Parameter": "Investor Equity Contribution", "Value": _default_equity_assumptions(base_model)["investor_equity_eur"], "Unit": "EUR", "Notes": "External investor contribution."},
                {"Parameter": "Investor Exit Year", "Value": _default_equity_assumptions(base_model)["exit_year"], "Unit": "Year", "Notes": "Exit year for investor."},
                {"Parameter": "Exit Multiple (x EBITDA)", "Value": _default_equity_assumptions(base_model)["exit_multiple"], "Unit": "x", "Notes": "Exit multiple on EBITDA."},
                {"Parameter": "Distribution Rule", "Value": "Pro-rata", "Unit": "", "Notes": "Fixed distribution rule."},
            ],
            "cashflow": [
                {"Parameter": "Tax Cash Rate", "Value": _default_cashflow_assumptions()["tax_cash_rate_pct"], "Unit": "%", "Notes": "Cash tax rate on EBT."},
                {"Parameter": "Tax Payment Lag", "Value": _default_cashflow_assumptions()["tax_payment_lag_years"], "Unit": "Years", "Notes": "Timing lag for cash taxes."},
                {"Parameter": "Capex (% of Revenue)", "Value": _default_cashflow_assumptions()["capex_pct_revenue"], "Unit": "%", "Notes": "Capex as % of revenue."},
                {"Parameter": "Working Capital (% of Revenue)", "Value": _default_cashflow_assumptions()["working_capital_pct_revenue"], "Unit": "%", "Notes": "Working capital adjustment."},
                {"Parameter": "Opening Cash Balance", "Value": _default_cashflow_assumptions()["opening_cash_balance_eur"], "Unit": "EUR", "Notes": "Opening cash balance."},
            ],
            "balance_sheet": [
                {"Parameter": "Opening Equity", "Value": _default_balance_sheet_assumptions(base_model)["opening_equity_eur"], "Unit": "EUR", "Notes": "Opening equity value."},
                {"Parameter": "Depreciation Rate", "Value": _default_balance_sheet_assumptions(base_model)["depreciation_rate_pct"], "Unit": "%", "Notes": "Fixed asset depreciation rate."},
                {"Parameter": "Minimum Cash Balance", "Value": _default_balance_sheet_assumptions(base_model)["minimum_cash_balance_eur"], "Unit": "EUR", "Notes": "Minimum cash balance."},
            ],
            "valuation": [
                {"Parameter": "Seller EBIT Multiple", "Value": _default_valuation_assumptions(base_model)["seller_ebit_multiple"], "Unit": "x", "Notes": "EBIT multiple for seller view."},
                {"Parameter": "Reference Year", "Value": _default_valuation_assumptions(base_model)["reference_year"], "Unit": "Year", "Notes": "Reference year for multiple."},
                {"Parameter": "Discount Rate (WACC)", "Value": _default_valuation_assumptions(base_model)["buyer_discount_rate"], "Unit": "%", "Notes": "DCF discount rate."},
                {"Parameter": "Valuation Start Year", "Value": _default_valuation_assumptions(base_model)["valuation_start_year"], "Unit": "Year", "Notes": "DCF start year."},
                {"Parameter": "Debt at Close", "Value": _default_valuation_assumptions(base_model)["debt_at_close_eur"], "Unit": "EUR", "Notes": "Net debt at close."},
                {"Parameter": "Transaction Costs (%)", "Value": _default_valuation_assumptions(base_model)["transaction_cost_pct"], "Unit": "%", "Notes": "Fees as % of EV."},
            ],
        }

    st.session_state.setdefault("assumptions", _seed_assumptions_state())
    st.session_state.setdefault("assumptions.auto_sync", True)

    def _apply_assumptions_state():
        state = st.session_state["assumptions"]
        for row in state["revenue_drivers"]:
            param = row["Parameter"]
            if param == "Utilization (%)":
                st.session_state["scenario_parameters.utilization_rate.base"] = _clamp_pct(row["Base"])
                st.session_state["scenario_parameters.utilization_rate.best"] = _clamp_pct(row["Best"])
                st.session_state["scenario_parameters.utilization_rate.worst"] = _clamp_pct(row["Worst"])
            elif param == "Day Rate (EUR)":
                st.session_state["scenario_parameters.day_rate_eur.base"] = _non_negative(row["Base"])
                st.session_state["scenario_parameters.day_rate_eur.best"] = _non_negative(row["Best"])
                st.session_state["scenario_parameters.day_rate_eur.worst"] = _non_negative(row["Worst"])
            elif param == "Consulting FTE":
                st.session_state["operating_assumptions.consulting_fte_start"] = _non_negative(row["Base"])
            elif param == "Workdays per Year":
                st.session_state["operating_assumptions.work_days_per_year"] = _non_negative(row["Base"])
            elif param == "Day Rate Growth (% p.a.)":
                st.session_state["operating_assumptions.day_rate_growth_pct"] = _clamp_pct(row["Base"])

        guarantee_map = {
            "Year 1": "operating_assumptions.revenue_guarantee_pct_year_1",
            "Year 2": "operating_assumptions.revenue_guarantee_pct_year_2",
            "Year 3": "operating_assumptions.revenue_guarantee_pct_year_3",
        }
        for row in state["revenue_guarantees"]:
            key = guarantee_map.get(row["Year"])
            if key:
                st.session_state[key] = _clamp_pct(row["Guarantee %"])

        for row in state["personnel_costs"]:
            role = row["Role"]
            if role == "Consultant Base Salary":
                st.session_state[
                    "personnel_cost_assumptions.avg_consultant_base_cost_eur_per_year"
                ] = _non_negative(row["Base Value (EUR)"])
                st.session_state["personnel_cost_assumptions.wage_inflation_pct"] = _clamp_pct(row["Growth (%)"])
            elif role == "Consultant Variable (% Revenue)":
                st.session_state["personnel_cost_assumptions.bonus_pct_of_base"] = _clamp_pct(row["Base Value (EUR)"])
            elif role == "Backoffice Cost per FTE":
                st.session_state[
                    "operating_assumptions.avg_backoffice_salary_eur_per_year"
                ] = _non_negative(row["Base Value (EUR)"])
                st.session_state["personnel_cost_assumptions.wage_inflation_pct"] = _clamp_pct(row["Growth (%)"])

        for row in state["opex"]:
            category = row["Category"]
            if category == "External Consulting":
                st.session_state["overhead_and_variable_costs.legal_audit_eur_per_year"] = _non_negative(row["Value"])
            elif category == "IT":
                st.session_state["overhead_and_variable_costs.it_and_software_eur_per_year"] = _non_negative(row["Value"])
            elif category == "Office":
                st.session_state["overhead_and_variable_costs.rent_eur_per_year"] = _non_negative(row["Value"])
            elif category == "Other Services":
                st.session_state["overhead_and_variable_costs.other_overhead_eur_per_year"] = _non_negative(row["Value"])

        for row in state["financing"]:
            parameter = row["Parameter"]
            if parameter == "Senior Debt Amount":
                st.session_state["transaction_and_financing.senior_term_loan_start_eur"] = _non_negative(row["Value"])
            elif parameter == "Interest Rate":
                st.session_state["transaction_and_financing.senior_interest_rate_pct"] = _clamp_pct(row["Value"])
            elif parameter == "Amortisation Years":
                st.session_state["financing.amortization_period_years"] = int(max(1, row["Value"]))
            elif parameter == "Transaction Fees (%)":
                st.session_state["valuation.transaction_cost_pct"] = _clamp_pct(row["Value"])

        for row in state["equity"]:
            parameter = row["Parameter"]
            if parameter == "Sponsor Equity Contribution":
                st.session_state["equity.sponsor_equity_eur"] = _non_negative(row["Value"])
            elif parameter == "Investor Equity Contribution":
                st.session_state["equity.investor_equity_eur"] = _non_negative(row["Value"])
            elif parameter == "Investor Exit Year":
                try:
                    exit_val = int(float(row["Value"]))
                except (TypeError, ValueError):
                    exit_val = _default_equity_assumptions(base_model)["exit_year"]
                st.session_state["equity.exit_year"] = int(max(3, min(7, exit_val)))
            elif parameter == "Exit Multiple (x EBITDA)":
                st.session_state["equity.exit_multiple"] = float(row["Value"])

        for row in state["cashflow"]:
            parameter = row["Parameter"]
            if parameter == "Tax Cash Rate":
                st.session_state["cashflow.tax_cash_rate_pct"] = _clamp_pct(row["Value"])
            elif parameter == "Tax Payment Lag":
                st.session_state["cashflow.tax_payment_lag_years"] = int(max(0, min(1, row["Value"])))
            elif parameter == "Capex (% of Revenue)":
                st.session_state["cashflow.capex_pct_revenue"] = _clamp_pct(row["Value"])
            elif parameter == "Working Capital (% of Revenue)":
                st.session_state["cashflow.working_capital_pct_revenue"] = _clamp_pct(row["Value"])
            elif parameter == "Opening Cash Balance":
                st.session_state["cashflow.opening_cash_balance_eur"] = _non_negative(row["Value"])

        for row in state["balance_sheet"]:
            parameter = row["Parameter"]
            if parameter == "Opening Equity":
                st.session_state["balance_sheet.opening_equity_eur"] = _non_negative(row["Value"])
            elif parameter == "Depreciation Rate":
                st.session_state["balance_sheet.depreciation_rate_pct"] = _clamp_pct(row["Value"])
            elif parameter == "Minimum Cash Balance":
                st.session_state["balance_sheet.minimum_cash_balance_eur"] = _non_negative(row["Value"])

        for row in state["valuation"]:
            parameter = row["Parameter"]
            if parameter == "Seller EBIT Multiple":
                st.session_state["valuation.seller_ebit_multiple"] = float(row["Value"])
            elif parameter == "Reference Year":
                st.session_state["valuation.reference_year"] = int(max(0, min(4, row["Value"])))
            elif parameter == "Discount Rate (WACC)":
                st.session_state["valuation.buyer_discount_rate"] = _clamp_pct(row["Value"])
            elif parameter == "Valuation Start Year":
                st.session_state["valuation.valuation_start_year"] = int(max(0, min(4, row["Value"])))
            elif parameter == "Debt at Close":
                st.session_state["valuation.debt_at_close_eur"] = _non_negative(row["Value"])
            elif parameter == "Transaction Costs (%)":
                st.session_state["valuation.transaction_cost_pct"] = _clamp_pct(row["Value"])

    if page == "Assumptions (Advanced)":
        st.header("Assumptions (Advanced)")
        st.write("Master input sheet – all model assumptions in one place")

        scenario_options = ["Base", "Best", "Worst"]
        scenario_default = base_model.scenario_selection[
            "selected_scenario"
        ].value
        scenario_index = (
            scenario_options.index(scenario_default)
            if scenario_default in scenario_options
            else 0
        )
        info_cols = st.columns([1, 1])
        selected_scenario = info_cols[0].selectbox(
            "Scenario",
            scenario_options,
            index=scenario_index,
            key="assumptions.scenario",
        )
        auto_sync = info_cols[1].toggle(
            "Auto-apply scenario values",
            value=st.session_state.get("assumptions.auto_sync", True),
            key="assumptions.auto_sync",
        )
        st.info("Changes here affect all pages instantly.")
        st.session_state["scenario_selection.selected_scenario"] = (
            selected_scenario
        )
        scenario_key = selected_scenario.lower()

        def _clamp_pct(value):
            if value is None or pd.isna(value):
                return 0.0
            return max(0.0, min(float(value), 1.0))

        def _non_negative(value):
            if value is None or pd.isna(value):
                return 0.0
            return max(0.0, float(value))

        if auto_sync:
            util_value = st.session_state.get(
                f"scenario_parameters.utilization_rate.{scenario_key}",
                base_model.scenario_parameters["utilization_rate"][
                    scenario_key
                ].value,
            )
            st.session_state["utilization_by_year"] = [util_value] * 5
            for i in range(5):
                st.session_state[f"utilization_by_year.{i}"] = util_value

        assumptions_state = st.session_state["assumptions"]
        revenue_df = pd.DataFrame(assumptions_state["revenue_drivers"])
        active_label = selected_scenario
        base_label = "Base (Active)" if active_label == "Base" else "Base"
        best_label = "Best (Active)" if active_label == "Best" else "Best"
        worst_label = "Worst (Active)" if active_label == "Worst" else "Worst"
        revenue_df = revenue_df.rename(
            columns={"Base": base_label, "Best": best_label, "Worst": worst_label}
        )
        st.markdown("### Revenue Drivers")
        revenue_edit = st.data_editor(
            revenue_df,
            hide_index=True,
            key="assumptions.revenue_drivers",
            column_config={
                "Parameter": st.column_config.TextColumn(disabled=True),
                "Unit": st.column_config.TextColumn(disabled=True),
                "Description": st.column_config.TextColumn(disabled=True),
                base_label: st.column_config.NumberColumn(
                    disabled=auto_sync and active_label != "Base"
                ),
                best_label: st.column_config.NumberColumn(
                    disabled=auto_sync and active_label != "Best"
                ),
                worst_label: st.column_config.NumberColumn(
                    disabled=auto_sync and active_label != "Worst"
                ),
            },
            use_container_width=True,
        )
        revenue_edit = revenue_edit.rename(
            columns={base_label: "Base", best_label: "Best", worst_label: "Worst"}
        )
        assumptions_state["revenue_drivers"] = revenue_edit.to_dict("records")
        for _, row in revenue_edit.iterrows():
            param = row["Parameter"]
            if param == "Utilization (%)":
                st.session_state["scenario_parameters.utilization_rate.base"] = _clamp_pct(row["Base"])
                st.session_state["scenario_parameters.utilization_rate.best"] = _clamp_pct(row["Best"])
                st.session_state["scenario_parameters.utilization_rate.worst"] = _clamp_pct(row["Worst"])
            elif param == "Day Rate (EUR)":
                st.session_state["scenario_parameters.day_rate_eur.base"] = _non_negative(row["Base"])
                st.session_state["scenario_parameters.day_rate_eur.best"] = _non_negative(row["Best"])
                st.session_state["scenario_parameters.day_rate_eur.worst"] = _non_negative(row["Worst"])
            elif param == "Consulting FTE":
                st.session_state["operating_assumptions.consulting_fte_start"] = _non_negative(row["Base"])
            elif param == "Workdays per Year":
                st.session_state["operating_assumptions.work_days_per_year"] = _non_negative(row["Base"])
            elif param == "Day Rate Growth (% p.a.)":
                st.session_state["operating_assumptions.day_rate_growth_pct"] = _clamp_pct(row["Base"])

        st.markdown("### Revenue Guarantees")
        guarantee_df = pd.DataFrame(assumptions_state["revenue_guarantees"])
        guarantee_edit = st.data_editor(
            guarantee_df,
            hide_index=True,
            key="assumptions.revenue_guarantees",
            column_config={
                "Year": st.column_config.TextColumn(disabled=True),
                "Description": st.column_config.TextColumn(disabled=True),
            },
            use_container_width=True,
        )
        assumptions_state["revenue_guarantees"] = guarantee_edit.to_dict("records")
        guarantee_map = {
            "Year 1": "operating_assumptions.revenue_guarantee_pct_year_1",
            "Year 2": "operating_assumptions.revenue_guarantee_pct_year_2",
            "Year 3": "operating_assumptions.revenue_guarantee_pct_year_3",
        }
        for _, row in guarantee_edit.iterrows():
            key = guarantee_map.get(row["Year"])
            if key:
                st.session_state[key] = _clamp_pct(row["Guarantee %"])

        st.markdown("### Personnel Costs")
        personnel_df = pd.DataFrame(assumptions_state["personnel_costs"])
        personnel_edit = st.data_editor(
            personnel_df,
            hide_index=True,
            key="assumptions.personnel_costs",
            column_config={
                "Role": st.column_config.TextColumn(disabled=True),
                "Cost Type": st.column_config.TextColumn(disabled=True),
                "Notes": st.column_config.TextColumn(disabled=True),
            },
            use_container_width=True,
        )
        assumptions_state["personnel_costs"] = personnel_edit.to_dict("records")
        for _, row in personnel_edit.iterrows():
            role = row["Role"]
            if role == "Consultant Base Salary":
                st.session_state[
                    "personnel_cost_assumptions.avg_consultant_base_cost_eur_per_year"
                ] = _non_negative(row["Base Value (EUR)"])
                st.session_state["personnel_cost_assumptions.wage_inflation_pct"] = _clamp_pct(row["Growth (%)"])
            elif role == "Consultant Variable (% Revenue)":
                st.session_state["personnel_cost_assumptions.bonus_pct_of_base"] = _clamp_pct(row["Base Value (EUR)"])
            elif role == "Backoffice Cost per FTE":
                st.session_state[
                    "operating_assumptions.avg_backoffice_salary_eur_per_year"
                ] = _non_negative(row["Base Value (EUR)"])
                st.session_state["personnel_cost_assumptions.wage_inflation_pct"] = _clamp_pct(row["Growth (%)"])
            elif role == "Management / MD Cost":
                st.session_state["assumptions.management_md_cost_eur"] = _non_negative(row["Base Value (EUR)"])

        st.markdown("### Operating Expenses (Opex)")
        opex_df = pd.DataFrame(assumptions_state["opex"])
        opex_edit = st.data_editor(
            opex_df,
            hide_index=True,
            key="assumptions.opex",
            column_config={
                "Category": st.column_config.TextColumn(disabled=True),
                "Cost Type": st.column_config.TextColumn(disabled=True),
                "Unit": st.column_config.TextColumn(disabled=True),
                "Notes": st.column_config.TextColumn(disabled=True),
            },
            use_container_width=True,
        )
        assumptions_state["opex"] = opex_edit.to_dict("records")
        for _, row in opex_edit.iterrows():
            category = row["Category"]
            if category == "External Consulting":
                st.session_state["overhead_and_variable_costs.legal_audit_eur_per_year"] = _non_negative(row["Value"])
            elif category == "IT":
                st.session_state["overhead_and_variable_costs.it_and_software_eur_per_year"] = _non_negative(row["Value"])
            elif category == "Office":
                st.session_state["overhead_and_variable_costs.rent_eur_per_year"] = _non_negative(row["Value"])
            elif category == "Other Services":
                st.session_state["overhead_and_variable_costs.other_overhead_eur_per_year"] = _non_negative(row["Value"])

        st.markdown("### Financing Assumptions")
        financing_df = pd.DataFrame(assumptions_state["financing"])
        financing_edit = st.data_editor(
            financing_df,
            hide_index=True,
            key="assumptions.financing",
            column_config={
                "Parameter": st.column_config.TextColumn(disabled=True),
                "Unit": st.column_config.TextColumn(disabled=True),
                "Notes": st.column_config.TextColumn(disabled=True),
            },
            use_container_width=True,
        )
        assumptions_state["financing"] = financing_edit.to_dict("records")
        for _, row in financing_edit.iterrows():
            parameter = row["Parameter"]
            if parameter == "Senior Debt Amount":
                st.session_state["transaction_and_financing.senior_term_loan_start_eur"] = _non_negative(row["Value"])
            elif parameter == "Interest Rate":
                st.session_state["transaction_and_financing.senior_interest_rate_pct"] = _clamp_pct(row["Value"])
            elif parameter == "Amortisation Years":
                st.session_state["financing.amortization_period_years"] = int(max(1, row["Value"]))
            elif parameter == "Transaction Fees (%)":
                st.session_state["valuation.transaction_cost_pct"] = _clamp_pct(row["Value"])

        st.markdown("### Equity & Investor Assumptions")
        equity_df = pd.DataFrame(assumptions_state["equity"])
        equity_edit = st.data_editor(
            equity_df,
            hide_index=True,
            key="assumptions.equity",
            column_config={
                "Parameter": st.column_config.TextColumn(disabled=True),
                "Unit": st.column_config.TextColumn(disabled=True),
                "Notes": st.column_config.TextColumn(disabled=True),
            },
            use_container_width=True,
        )
        assumptions_state["equity"] = equity_edit.to_dict("records")
        for _, row in equity_edit.iterrows():
            parameter = row["Parameter"]
            if parameter == "Sponsor Equity Contribution":
                st.session_state["equity.sponsor_equity_eur"] = _non_negative(row["Value"])
            elif parameter == "Investor Equity Contribution":
                st.session_state["equity.investor_equity_eur"] = _non_negative(row["Value"])
            elif parameter == "Investor Exit Year":
                try:
                    exit_val = int(float(row["Value"]))
                except (TypeError, ValueError):
                    exit_val = _default_equity_assumptions(base_model)[
                        "exit_year"
                    ]
                st.session_state["equity.exit_year"] = int(
                    max(3, min(7, exit_val))
                )
            elif parameter == "Exit Multiple (x EBITDA)":
                st.session_state["equity.exit_multiple"] = float(row["Value"])

        st.markdown("### Cashflow Assumptions")
        cashflow_df = pd.DataFrame(assumptions_state["cashflow"])
        cashflow_edit = st.data_editor(
            cashflow_df,
            hide_index=True,
            key="assumptions.cashflow",
            column_config={
                "Parameter": st.column_config.TextColumn(disabled=True),
                "Unit": st.column_config.TextColumn(disabled=True),
                "Notes": st.column_config.TextColumn(disabled=True),
            },
            use_container_width=True,
        )
        assumptions_state["cashflow"] = cashflow_edit.to_dict("records")
        for _, row in cashflow_edit.iterrows():
            parameter = row["Parameter"]
            if parameter == "Tax Cash Rate":
                st.session_state["cashflow.tax_cash_rate_pct"] = _clamp_pct(row["Value"])
            elif parameter == "Tax Payment Lag":
                st.session_state["cashflow.tax_payment_lag_years"] = int(max(0, min(1, row["Value"])))
            elif parameter == "Capex (% of Revenue)":
                st.session_state["cashflow.capex_pct_revenue"] = _clamp_pct(row["Value"])
            elif parameter == "Working Capital (% of Revenue)":
                st.session_state["cashflow.working_capital_pct_revenue"] = _clamp_pct(row["Value"])
            elif parameter == "Opening Cash Balance":
                st.session_state["cashflow.opening_cash_balance_eur"] = _non_negative(row["Value"])

        st.markdown("### Balance Sheet Assumptions")
        balance_df = pd.DataFrame(assumptions_state["balance_sheet"])
        balance_edit = st.data_editor(
            balance_df,
            hide_index=True,
            key="assumptions.balance_sheet",
            column_config={
                "Parameter": st.column_config.TextColumn(disabled=True),
                "Unit": st.column_config.TextColumn(disabled=True),
                "Notes": st.column_config.TextColumn(disabled=True),
            },
            use_container_width=True,
        )
        assumptions_state["balance_sheet"] = balance_edit.to_dict("records")
        for _, row in balance_edit.iterrows():
            parameter = row["Parameter"]
            if parameter == "Opening Equity":
                st.session_state["balance_sheet.opening_equity_eur"] = _non_negative(row["Value"])
            elif parameter == "Depreciation Rate":
                st.session_state["balance_sheet.depreciation_rate_pct"] = _clamp_pct(row["Value"])
            elif parameter == "Minimum Cash Balance":
                st.session_state["balance_sheet.minimum_cash_balance_eur"] = _non_negative(row["Value"])

        st.markdown("### Valuation Assumptions")
        valuation_df = pd.DataFrame(assumptions_state["valuation"])
        valuation_edit = st.data_editor(
            valuation_df,
            hide_index=True,
            key="assumptions.valuation",
            column_config={
                "Parameter": st.column_config.TextColumn(disabled=True),
                "Unit": st.column_config.TextColumn(disabled=True),
                "Notes": st.column_config.TextColumn(disabled=True),
            },
            use_container_width=True,
        )
        assumptions_state["valuation"] = valuation_edit.to_dict("records")
        for _, row in valuation_edit.iterrows():
            parameter = row["Parameter"]
            if parameter == "Seller EBIT Multiple":
                st.session_state["valuation.seller_ebit_multiple"] = float(row["Value"])
            elif parameter == "Reference Year":
                st.session_state["valuation.reference_year"] = int(max(0, min(4, row["Value"])))
            elif parameter == "Discount Rate (WACC)":
                st.session_state["valuation.buyer_discount_rate"] = _clamp_pct(row["Value"])
            elif parameter == "Valuation Start Year":
                st.session_state["valuation.valuation_start_year"] = int(max(0, min(4, row["Value"])))
            elif parameter == "Debt at Close":
                st.session_state["valuation.debt_at_close_eur"] = _non_negative(row["Value"])
            elif parameter == "Transaction Costs (%)":
                st.session_state["valuation.transaction_cost_pct"] = _clamp_pct(row["Value"])

        _apply_assumptions_state()

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
        "amortization_type": "Linear",
        "amortization_period_years": st.session_state.get(
            "financing.amortization_period_years",
            financing_defaults["amortization_period_years"],
        ),
        "grace_period_years": 0,
        "special_repayment_year": None,
        "special_repayment_amount_eur": 0.0,
        "minimum_dscr": st.session_state.get(
            "financing.minimum_dscr",
            financing_defaults["minimum_dscr"],
        ),
        "minimum_cash_balance_eur": financing_defaults[
            "minimum_cash_balance_eur"
        ],
        "maintenance_capex_pct_revenue": st.session_state.get(
            "financing.maintenance_capex_pct_revenue",
            financing_defaults["maintenance_capex_pct_revenue"],
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
    equity_defaults = _default_equity_assumptions(input_model)
    input_model.equity_assumptions = {
        "sponsor_equity_eur": st.session_state.get(
            "equity.sponsor_equity_eur",
            equity_defaults["sponsor_equity_eur"],
        ),
        "investor_equity_eur": st.session_state.get(
            "equity.investor_equity_eur",
            equity_defaults["investor_equity_eur"],
        ),
        "exit_year": st.session_state.get(
            "equity.exit_year",
            equity_defaults["exit_year"],
        ),
        "exit_method": st.session_state.get(
            "equity.exit_method",
            equity_defaults["exit_method"],
        ),
        "exit_multiple": st.session_state.get(
            "equity.exit_multiple",
            equity_defaults["exit_multiple"],
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

    def _format_value(value, formatter):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return "—"
        return formatter(value)

    def _build_model_snapshot_text():
        scenario = input_model.scenario_selection["selected_scenario"].value
        scenario_key = scenario.lower()
        utilization_by_year = getattr(
            input_model,
            "utilization_by_year",
            [
                input_model.scenario_parameters["utilization_rate"][
                    scenario_key
                ].value
            ]
            * 5,
        )
        day_rate_base = input_model.scenario_parameters["day_rate_eur"][
            scenario_key
        ].value
        day_rate_growth = input_model.operating_assumptions[
            "day_rate_growth_pct"
        ].value
        guarantees = [
            input_model.operating_assumptions[
                "revenue_guarantee_pct_year_1"
            ].value,
            input_model.operating_assumptions[
                "revenue_guarantee_pct_year_2"
            ].value,
            input_model.operating_assumptions[
                "revenue_guarantee_pct_year_3"
            ].value,
        ]
        pnl_lines = []
        fcf_lines = []
        dscr_lines = []
        for year_index in range(5):
            year_label = f"Year {year_index}"
            pnl_data = pnl_result.get(year_label, {})
            revenue = pnl_data.get("revenue")
            ebitda = pnl_data.get("ebitda")
            net_income = pnl_data.get("net_income")
            pnl_lines.append(
                f"{year_label}: Revenue {_format_value(revenue, format_currency)}, "
                f"EBITDA {_format_value(ebitda, format_currency)}, "
                f"Net Income {_format_value(net_income, format_currency)}"
            )
            fcf_value = cashflow_result[year_index].get("free_cashflow")
            fcf_lines.append(
                f"{year_label}: {_format_value(fcf_value, format_currency)}"
            )
            dscr_value = debt_schedule[year_index].get("dscr")
            dscr_lines.append(
                f"{year_label}: {_format_value(dscr_value, lambda v: f'{v:.2f}x')}"
            )

        text_lines = [
            "1) Model Overview",
            "Model type: Consulting MBO",
            "Planning horizon: 5 years (Year 0–Year 4)",
            "Scenarios available: Base, Best, Worst",
            "",
            "2) Assumptions",
            "Revenue drivers:",
            f"- Consulting FTE: {_format_value(input_model.operating_assumptions['consulting_fte_start'].value, format_int)}",
            f"- Workdays per Year: {_format_value(input_model.operating_assumptions['work_days_per_year'].value, format_int)}",
            "- Utilization per year: "
            + ", ".join(
                f"Year {idx} {_format_value(value, format_pct)}"
                for idx, value in enumerate(utilization_by_year)
            ),
            f"- Day Rate (Base): {_format_value(day_rate_base, format_int)} EUR",
            f"- Day Rate Growth: {_format_value(day_rate_growth, format_pct)}",
            "Revenue guarantees:",
            f"- Year 1: {_format_value(guarantees[0], format_pct)}",
            f"- Year 2: {_format_value(guarantees[1], format_pct)}",
            f"- Year 3: {_format_value(guarantees[2], format_pct)}",
            "Personnel costs:",
            f"- Consultant Base Cost: {_format_value(input_model.personnel_cost_assumptions['avg_consultant_base_cost_eur_per_year'].value, format_int)} EUR",
            f"- Consultant Bonus: {_format_value(input_model.personnel_cost_assumptions['bonus_pct_of_base'].value, format_pct)}",
            f"- Backoffice Cost per FTE: {_format_value(input_model.operating_assumptions['avg_backoffice_salary_eur_per_year'].value, format_int)} EUR",
            "Opex:",
            f"- External Consulting: {_format_value(input_model.overhead_and_variable_costs['legal_audit_eur_per_year'].value, format_currency)}",
            f"- IT: {_format_value(input_model.overhead_and_variable_costs['it_and_software_eur_per_year'].value, format_currency)}",
            f"- Office: {_format_value(input_model.overhead_and_variable_costs['rent_eur_per_year'].value, format_currency)}",
            f"- Other Services: {_format_value(input_model.overhead_and_variable_costs['other_overhead_eur_per_year'].value, format_currency)}",
            "Financing assumptions:",
            f"- Purchase Price: {_format_value(input_model.transaction_and_financing['purchase_price_eur'].value, format_currency)}",
            f"- Debt Amount: {_format_value(input_model.transaction_and_financing['senior_term_loan_start_eur'].value, format_currency)}",
            f"- Interest Rate: {_format_value(input_model.transaction_and_financing['senior_interest_rate_pct'].value, format_pct)}",
            f"- Amortisation Years: {_format_value(input_model.financing_assumptions['amortization_period_years'], format_int)}",
            f"- Transaction Fees (%): {_format_value(st.session_state.get('valuation.transaction_cost_pct', 0.0), format_pct)}",
            "Equity & investor assumptions:",
            f"- Sponsor Equity: {_format_value(st.session_state.get('equity.sponsor_equity_eur'), format_currency)}",
            f"- Investor Equity: {_format_value(st.session_state.get('equity.investor_equity_eur'), format_currency)}",
            f"- Investor Exit Year: Year {st.session_state.get('equity.exit_year')}",
            f"- Exit Multiple: {_format_value(st.session_state.get('equity.exit_multiple'), lambda v: f'{v:.2f}x')}",
            "- Distribution Rule: Pro-rata",
            "",
            "3) Calculation Logic (Plain English)",
            "Revenue build-up: FTE × Workdays × Utilization × Day Rate.",
            "EBITDA: Revenue minus Personnel Costs and Operating Expenses.",
            "Cashflow: EBITDA minus Taxes, Working Capital Change, and Capex.",
            "Debt service: Interest on opening debt plus scheduled repayment.",
            "Equity cashflows: Investor exits in the exit year; sponsor retains residual equity.",
            "Exit value: Exit EBITDA × Exit Multiple, less net debt.",
            "",
            "4) KPI Definitions",
            "EBITDA Margin = EBITDA / Revenue",
            "DSCR = Operating Cashflow / Debt Service",
            "IRR = Discount rate where NPV of cashflows = 0",
            "MOIC = Total proceeds / Invested equity",
            "",
            "5) Notes",
            "Distributions: None modeled during the hold period.",
            "Exit: Investor exits fully at the selected exit year.",
            "Simplifications: Single debt tranche, no working capital schedule.",
        ]
        return "\n".join(text_lines)

    if page == "Overview":
        st.header("Overview")
        st.write("Decision-oriented executive summary.")

        revenue_values = [pnl_result[f"Year {i}"]["revenue"] for i in range(5)]
        ebitda_values = [pnl_result[f"Year {i}"]["ebitda"] for i in range(5)]
        ebit_values = [pnl_result[f"Year {i}"]["ebit"] for i in range(5)]
        net_income_values = [
            pnl_result[f"Year {i}"]["net_income"] for i in range(5)
        ]
        free_cashflow_values = [
            cashflow_result[i]["free_cashflow"] for i in range(5)
        ]
        cash_balances = [cashflow_result[i]["cash_balance"] for i in range(5)]
        min_cash_balance = min(cash_balances)
        avg_revenue = sum(revenue_values) / len(revenue_values)
        avg_ebitda = sum(ebitda_values) / len(ebitda_values)
        avg_ebitda_margin = (
            avg_ebitda / avg_revenue if avg_revenue else 0
        )
        avg_free_cashflow = sum(free_cashflow_values) / len(
            free_cashflow_values
        )

        purchase_price = input_model.transaction_and_financing[
            "purchase_price_eur"
        ].value
        equity_contribution = input_model.transaction_and_financing[
            "equity_contribution_eur"
        ].value
        debt_at_close = debt_schedule[0]["opening_debt"]
        leverage_irr = st.session_state.get(
            "equity.sponsor_irr", investment_result["irr"]
        )

        st.markdown("### A. Deal Snapshot")
        deal_snapshot = pd.DataFrame(
            [
                {
                    "Metric": "Purchase Price (EUR)",
                    "Value": format_currency(purchase_price),
                    "Explanation": "Enterprise value paid at close.",
                },
                {
                    "Metric": "Equity Contribution (EUR)",
                    "Value": format_currency(equity_contribution),
                    "Explanation": "Equity funded by management / sponsor.",
                },
                {
                    "Metric": "Debt at Close (EUR)",
                    "Value": format_currency(debt_at_close),
                    "Explanation": "Senior debt drawn at close.",
                },
                {
                    "Metric": "Avg Revenue (EUR)",
                    "Value": format_currency(avg_revenue),
                    "Explanation": "Average annual revenue over plan.",
                },
                {
                    "Metric": "Avg EBITDA & Margin",
                    "Value": f"{format_currency(avg_ebitda)} / {format_pct(avg_ebitda_margin)}",
                    "Explanation": "Average EBITDA and margin over plan.",
                },
                {
                    "Metric": "Avg Free Cash Flow",
                    "Value": format_currency(avg_free_cashflow),
                    "Explanation": "Average annual free cashflow.",
                },
                {
                    "Metric": "Levered Equity IRR",
                    "Value": format_pct(leverage_irr),
                    "Explanation": "Equity return from levered cashflows.",
                },
                {
                    "Metric": "Minimum Cash Balance",
                    "Value": format_currency(min_cash_balance),
                    "Explanation": "Lowest cash balance over plan.",
                },
            ]
        )
        st.dataframe(deal_snapshot, use_container_width=True)

        st.markdown("### B. Business Performance Summary")
        performance_table = pd.DataFrame(
            {
                "Revenue": revenue_values,
                "EBITDA": ebitda_values,
                "EBIT": ebit_values,
                "Net Income": net_income_values,
            },
            index=[f"Year {i}" for i in range(5)],
        )
        performance_table = performance_table.applymap(format_currency)
        st.dataframe(performance_table, use_container_width=True)

        st.markdown("### C. Financing & Risk Snapshot")
        min_dscr_value = min(row["dscr"] for row in debt_schedule)
        dscr_threshold = input_model.financing_assumptions["minimum_dscr"]
        stress_years = [
            f"Year {row['year']}"
            for row in debt_schedule
            if row["dscr"] < dscr_threshold
        ]
        financing_snapshot = pd.DataFrame(
            [
                {
                    "Metric": "Minimum DSCR",
                    "Value": f"{min_dscr_value:.2f}x",
                },
                {
                    "Metric": "Covenant Headroom",
                    "Value": "NO" if stress_years else "YES",
                },
                {
                    "Metric": "Peak Debt",
                    "Value": format_currency(
                        max(row["opening_debt"] for row in debt_schedule)
                    ),
                },
                {
                    "Metric": "Years with Covenant Stress",
                    "Value": ", ".join(stress_years) if stress_years else "None",
                },
            ]
        )
        st.dataframe(financing_snapshot, use_container_width=True)

        st.markdown("### D. Equity Story")
        st.info(
            f"At a purchase price of {format_currency(purchase_price)} and equity of "
            f"{format_currency(equity_contribution)}, the deal generates a levered "
            f"IRR of {format_pct(leverage_irr)}, driven by operating cashflow and "
            "deleveraging."
        )

    if page == "Settings":
        st.header("Settings")
        st.write("Model transparency, export, and technical controls")

        st.markdown("### Model Snapshot")
        if st.button("Copy Full Model Snapshot", key="copy_model_snapshot"):
            st.session_state["model_snapshot_text"] = _build_model_snapshot_text()
        if st.session_state.get("model_snapshot_text"):
            st.text_area(
                "Snapshot",
                value=st.session_state["model_snapshot_text"],
                height=420,
            )

        st.markdown("### Calculation Logic")
        st.markdown("**Revenue Logic**")
        st.write(
            "Revenue is calculated as consulting FTE multiplied by workdays, "
            "utilization, and the applicable day rate (including growth)."
        )
        st.markdown("**Cost Logic**")
        st.write(
            "Personnel costs use base salary, bonus, and payroll burden with "
            "wage inflation. Operating expenses include fixed overhead items."
        )
        st.markdown("**Financing Logic**")
        st.write(
            "Debt service equals interest on opening debt plus scheduled "
            "repayment over the amortisation period."
        )
        st.markdown("**Equity Logic**")
        st.write(
            "Investor exits in the selected year at the exit multiple. "
            "Sponsor retains residual equity thereafter."
        )

        st.markdown("### Data Persistence")
        st.write(
            "Inputs are stored in Streamlit session_state during the session."
        )
        st.write("Values do not persist across a full browser refresh.")
        st.write("All assumptions are exported to the Excel model.")

        st.markdown("### Export Status")
        st.write("Excel export: Enabled (Beta)")
        st.write(
            "Sheets included: Assumptions, P&L, Cashflow, Balance Sheet, "
            "Valuation, Financing, Equity"
        )

        with st.expander("Danger Zone", expanded=False):
            reset_confirm = st.checkbox(
                "I understand this will reset model state",
                key="reset_confirm",
            )
            if st.button("Reset Model", disabled=not reset_confirm):
                st.session_state.clear()
                st.rerun()
            if st.button("Reset Scenario", disabled=not reset_confirm):
                st.session_state["scenario_selection.selected_scenario"] = (
                    base_model.scenario_selection["selected_scenario"].value
                )
                st.rerun()
            if st.button("Clear Session State", disabled=not reset_confirm):
                st.session_state.clear()
                st.rerun()

    if page == "Valuation & Purchase Price":
        st.header("Valuation & Purchase Price")
        st.write("Seller vs. buyer view (5-year plan)")

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
        financing_assumptions = input_model.financing_assumptions
        cashflow_by_year = {row["year"]: row for row in cashflow_result}
        maintenance_capex_pct = financing_assumptions[
            "maintenance_capex_pct_revenue"
        ]

        bank_rows = []
        for year_index in range(5):
            year_label = f"Year {year_index}"
            ebitda = pnl_result[year_label]["ebitda"]
            cash_taxes = cashflow_by_year[year_index]["taxes_paid"]
            revenue = pnl_result[year_label]["revenue"]
            maintenance_capex = revenue * maintenance_capex_pct
            working_capital = 0.0
            cfads = ebitda - cash_taxes - maintenance_capex + working_capital
            interest = debt_schedule[year_index]["interest_expense"]
            scheduled_repayment = debt_schedule[year_index][
                "scheduled_repayment"
            ]
            debt_service = interest + scheduled_repayment
            dscr = cfads / debt_service if debt_service else 0
            min_dscr = financing_assumptions["minimum_dscr"]
            dscr_display = -abs(dscr) if dscr < min_dscr else dscr
            breach = "YES" if dscr < min_dscr else "NO"

            bank_rows.append({"Line Item": "EBITDA", year_label: ebitda})
            bank_rows.append(
                {"Line Item": "Cash Taxes", year_label: cash_taxes}
            )
            bank_rows.append(
                {"Line Item": "Capex (Maintenance)", year_label: maintenance_capex}
            )
            bank_rows.append({"Line Item": "CFADS", year_label: cfads})
            bank_rows.append(
                {"Line Item": "Interest Expense", year_label: interest}
            )
            bank_rows.append(
                {
                    "Line Item": "Scheduled Repayment",
                    year_label: scheduled_repayment,
                }
            )
            bank_rows.append(
                {"Line Item": "Debt Service", year_label: debt_service}
            )
            bank_rows.append({"Line Item": "DSCR", year_label: dscr_display})
            bank_rows.append(
                {
                    "Line Item": "Minimum Required DSCR",
                    year_label: min_dscr,
                }
            )
            bank_rows.append(
                {"Line Item": "Covenant Breach", year_label: breach}
            )

        bank_metrics = {}
        for entry in bank_rows:
            label = entry["Line Item"]
            bank_metrics.setdefault(label, {})
            for year_index in range(5):
                year_label = f"Year {year_index}"
                if year_label in entry:
                    bank_metrics[label][year_label] = entry[year_label]

        bank_table = pd.DataFrame.from_dict(bank_metrics, orient="index")
        bank_table = bank_table[
            [f"Year {i}" for i in range(5)]
        ].reset_index()
        bank_table.rename(columns={"index": "Line Item"}, inplace=True)
        bank_formatters = {
            "DSCR": lambda value: f"{abs(value):.2f}x"
            if value not in ("", None)
            else "",
            "Minimum Required DSCR": lambda value: f"{value:.2f}x"
            if value not in ("", None)
            else "",
            "Covenant Breach": lambda value: value if value else "",
        }
        st.markdown("### Bank View")
        _render_custom_table_html(
            bank_table,
            set(),
            {"CFADS", "Debt Service"},
            bank_formatters,
        )

        dscr_values = [
            abs(value)
            for value in bank_metrics.get("DSCR", {}).values()
            if isinstance(value, (int, float))
        ]
        avg_dscr = sum(dscr_values) / len(dscr_values) if dscr_values else 0
        min_dscr_value = min(dscr_values) if dscr_values else 0
        peak_debt = max(row["opening_debt"] for row in debt_schedule)
        debt_at_close = debt_schedule[0]["opening_debt"]
        ebitda_year0 = pnl_result["Year 0"]["ebitda"]
        debt_to_ebitda = debt_at_close / ebitda_year0 if ebitda_year0 else 0

        st.markdown("### KPIs")
        kpi_table = pd.DataFrame(
            [
                {"KPI": "Average DSCR", "Value": f"{avg_dscr:.2f}x"},
                {"KPI": "Minimum DSCR", "Value": f"{min_dscr_value:.2f}x"},
                {"KPI": "Peak Debt", "Value": format_currency(peak_debt)},
                {
                    "KPI": "Debt / EBITDA (at close)",
                    "Value": f"{debt_to_ebitda:.2f}x",
                },
            ]
        )
        st.dataframe(kpi_table, use_container_width=True)

        explain_financing = st.toggle("Explain Financing Logic")
        if explain_financing:
            st.markdown("### Explanation")
            st.write(
                "CFADS is the cash the business generates to service debt "
                "after cash taxes and maintenance capex."
            )
            critical_years = []
            for year_index in range(5):
                year_label = f"Year {year_index}"
                cfads = bank_metrics["CFADS"].get(year_label, 0)
                debt_service = bank_metrics["Debt Service"].get(year_label, 0)
                dscr_value = abs(bank_metrics["DSCR"].get(year_label, 0))
                min_dscr = bank_metrics["Minimum Required DSCR"].get(
                    year_label, 0
                )
                status = (
                    "above"
                    if dscr_value >= min_dscr
                    else "below"
                )
                st.write(
                    f"In {year_label}, the business generates "
                    f"{format_currency(cfads)} of cash available for debt service. "
                    f"Required debt service is {format_currency(debt_service)}, "
                    f"resulting in a DSCR of {dscr_value:.2f}x. "
                    f"This is {status} the required covenant of {min_dscr:.2f}x."
                )
                if dscr_value < min_dscr:
                    critical_years.append(year_label)

            if critical_years:
                st.write(
                    f"Critical years: {', '.join(critical_years)}. "
                    "The structure would not be acceptable to a senior bank "
                    "without changes."
                )
            else:
                st.write(
                    "All years meet the covenant threshold. "
                    "The structure is bankable under current assumptions."
                )

    if page == "Equity Case":
        st.header("Equity Case")
        st.write("Sponsor and investor economics with explicit entry and exit.")

        equity_defaults = _default_equity_assumptions(input_model)
        sponsor_equity = st.session_state.get(
            "equity.sponsor_equity_eur",
            equity_defaults["sponsor_equity_eur"],
        )
        investor_equity = st.session_state.get(
            "equity.investor_equity_eur",
            equity_defaults["investor_equity_eur"],
        )
        total_equity = sponsor_equity + investor_equity
        sponsor_pct = sponsor_equity / total_equity if total_equity else 0.0
        investor_pct = investor_equity / total_equity if total_equity else 0.0
        exit_year = st.session_state.get(
            "equity.exit_year",
            equity_defaults["exit_year"],
        )
        exit_multiple = st.session_state.get(
            "equity.exit_multiple",
            equity_defaults["exit_multiple"],
        )
        exit_year = min(max(exit_year, 3), 7)
        exit_year_index = min(exit_year, 4)
        ebitda_exit = pnl_result[f"Year {exit_year_index}"]["ebitda"]
        net_debt_exit = (
            balance_sheet[exit_year_index]["financial_debt"]
            - balance_sheet[exit_year_index]["cash"]
        )
        enterprise_value_exit = ebitda_exit * exit_multiple
        equity_value_exit = enterprise_value_exit - net_debt_exit

        sponsor_cashflows = []
        investor_cashflows = []
        sponsor_residual_value = equity_value_exit
        investor_exit_proceeds = equity_value_exit * investor_pct
        for year_index in range(8):
            if year_index == 0:
                sponsor_cf = -sponsor_equity
                investor_cf = -investor_equity
            elif year_index == exit_year:
                sponsor_cf = sponsor_residual_value
                investor_cf = investor_exit_proceeds
            else:
                sponsor_cf = 0.0
                investor_cf = 0.0
            sponsor_cashflows.append(sponsor_cf)
            investor_cashflows.append(investor_cf)

        sponsor_irr = _calculate_irr(sponsor_cashflows)
        investor_irr = _calculate_irr(investor_cashflows)
        st.session_state["equity.sponsor_irr"] = sponsor_irr
        st.session_state["equity.investor_irr"] = investor_irr

        st.markdown("### Headline Metrics")
        investor_cols = st.columns(4)
        sponsor_cols = st.columns(4)
        investor_cols[0].metric(
            "Investor Equity (EUR)", format_currency(investor_equity)
        )
        investor_cols[1].metric("Exit Year", f"Year {exit_year}")
        investor_cols[2].metric(
            "Exit Equity Value (EUR)", format_currency(equity_value_exit)
        )
        investor_cols[3].metric("Investor IRR", format_pct(investor_irr))
        sponsor_cols[0].metric(
            "Sponsor Equity (EUR)", format_currency(sponsor_equity)
        )
        sponsor_cols[1].metric(
            "Ownership at Entry", format_pct(sponsor_pct)
        )
        sponsor_cols[2].metric("Ownership Post Exit", "100%")
        sponsor_cols[3].metric("Sponsor IRR", format_pct(sponsor_irr))

        st.markdown("### Sources & Uses")
        purchase_price = input_model.transaction_and_financing[
            "purchase_price_eur"
        ].value
        transaction_cost_pct = st.session_state.get(
            "valuation.transaction_cost_pct",
            _default_valuation_assumptions(input_model)["transaction_cost_pct"],
        )
        transaction_fees = purchase_price * transaction_cost_pct
        opening_cash = cashflow_result[0]["cash_balance"]
        debt_at_close = debt_schedule[0]["opening_debt"]
        uses_total = purchase_price + transaction_fees + opening_cash
        sources_total = debt_at_close + sponsor_equity + investor_equity

        sources_uses = pd.DataFrame(
            [
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
                    "Year 0": debt_at_close,
                    "Year 1": "",
                    "Year 2": "",
                    "Year 3": "",
                    "Year 4": "",
                },
                {
                    "Line Item": "Sponsor Equity",
                    "Year 0": sponsor_equity,
                    "Year 1": "",
                    "Year 2": "",
                    "Year 3": "",
                    "Year 4": "",
                },
                {
                    "Line Item": "Investor Equity",
                    "Year 0": investor_equity,
                    "Year 1": "",
                    "Year 2": "",
                    "Year 3": "",
                    "Year 4": "",
                },
                {
                    "Line Item": "Total Sources",
                    "Year 0": sources_total,
                    "Year 1": "",
                    "Year 2": "",
                    "Year 3": "",
                    "Year 4": "",
                },
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
                    "Line Item": "Cash to Balance Sheet",
                    "Year 0": opening_cash,
                    "Year 1": "",
                    "Year 2": "",
                    "Year 3": "",
                    "Year 4": "",
                },
                {
                    "Line Item": "Total Uses",
                    "Year 0": uses_total,
                    "Year 1": "",
                    "Year 2": "",
                    "Year 3": "",
                    "Year 4": "",
                },
            ]
        )
        _render_custom_table_html(
            sources_uses,
            {"SOURCES", "USES"},
            {"Total Sources", "Total Uses"},
            {},
        )

        st.markdown("### Equity Ownership")
        ownership_table = pd.DataFrame(
            [
                {
                    "Line Item": "Sponsor Equity",
                    "Equity (EUR)": sponsor_equity,
                    "Ownership (%)": format_pct(sponsor_pct),
                },
                {
                    "Line Item": "Investor Equity",
                    "Equity (EUR)": investor_equity,
                    "Ownership (%)": format_pct(investor_pct),
                },
                {
                    "Line Item": "Total Equity",
                    "Equity (EUR)": total_equity,
                    "Ownership (%)": format_pct(1.0),
                },
            ]
        )
        _render_custom_table_html(
            ownership_table,
            set(),
            {"Total Equity"},
            {
                "Sponsor Equity": lambda value: format_currency(value)
                if isinstance(value, (int, float))
                else value,
                "Investor Equity": lambda value: format_currency(value)
                if isinstance(value, (int, float))
                else value,
                "Total Equity": lambda value: format_currency(value)
                if isinstance(value, (int, float))
                else value,
            },
            year_labels=["Equity (EUR)", "Ownership (%)"],
        )

        st.markdown("### Equity Cashflows")
        cashflow_rows = [
            {
                "Line Item": "Sponsor Cashflow",
                **{f"Year {i}": sponsor_cashflows[i] for i in range(8)},
            },
            {
                "Line Item": "Investor Cashflow",
                **{f"Year {i}": investor_cashflows[i] for i in range(8)},
            },
            {
                "Line Item": "Sponsor Residual Equity Value",
                **{
                    f"Year {i}": sponsor_residual_value if i == exit_year else 0
                    for i in range(8)
                },
            },
        ]
        cashflow_table = pd.DataFrame(cashflow_rows)
        year_labels = [f"Year {i}" for i in range(8)]
        cashflow_table = cashflow_table[["Line Item"] + year_labels]
        _render_custom_table_html(
            cashflow_table,
            set(),
            set(),
            {},
            year_labels=year_labels,
        )

        sponsor_residual = sponsor_residual_value
        investor_moic = (
            (investor_exit_proceeds / abs(investor_equity))
            if investor_equity
            else 0
        )

        st.markdown("### Equity KPIs")
        kpi_table = pd.DataFrame(
            [
                {
                    "Investor": "Sponsor",
                    "Invested Equity": format_currency(sponsor_equity),
                    "Exit Proceeds": format_currency(sponsor_residual),
                    "MOIC": "—",
                    "IRR": format_pct(sponsor_irr),
                },
                {
                    "Investor": "Investor",
                    "Invested Equity": format_currency(investor_equity),
                    "Exit Proceeds": format_currency(investor_exit_proceeds),
                    "MOIC": f"{investor_moic:.2f}x",
                    "IRR": format_pct(investor_irr),
                },
            ]
        )
        st.dataframe(kpi_table, use_container_width=True)

        explain_equity = st.toggle("Explain Equity Logic")
        if explain_equity:
            st.markdown("### Explanation")
            st.write(
                f"Management contributes {format_currency(sponsor_equity)} of equity, "
                "which is insufficient to fund the transaction alone. "
                f"An external investor contributes {format_currency(investor_equity)}, "
                f"resulting in an initial ownership split of "
                f"{format_pct(sponsor_pct)} sponsor and {format_pct(investor_pct)} investor."
            )
            st.write(
                f"The purchase price of {format_currency(purchase_price)} is "
                f"funded by {format_currency(debt_at_close)} of senior debt and "
                f"{format_currency(total_equity)} of equity."
            )
            st.write(
                f"The investor exits in year {exit_year} using an EBITDA multiple "
                f"of {exit_multiple:.2f}x. This implies an equity value of "
                f"{format_currency(equity_value_exit)}, and the investor receives "
                f"{format_currency(investor_exit_proceeds)} on exit."
            )
            st.write(
                "During the holding period, no interim distributions are modeled. "
                "The investor exits fully in the exit year, while management "
                "retains 100% ownership thereafter."
            )
            st.write(
                "Sponsor returns are driven by the residual equity value after the "
                "investor exits, as well as leverage and exit valuation."
            )


if __name__ == "__main__":
    run_app()
