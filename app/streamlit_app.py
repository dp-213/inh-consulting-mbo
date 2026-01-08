import pandas as pd
import streamlit as st

from app.data_model import InputModel, create_demo_input_model
import app.run_model as run_model


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


def run_app():
    st.title("Financial Model")

    # Build input model and apply sidebar overrides.
    use_demo = st.sidebar.checkbox("Use demo values", value=True)
    input_model = create_demo_input_model() if use_demo else InputModel()

    scenario = st.sidebar.selectbox("Scenario", ["Base", "Best", "Worst"])
    input_model.scenario_selection["selected_scenario"].value = scenario

    utilization_default = input_model.scenario_parameters["utilization_rate"][
        scenario.lower()
    ].value
    day_rate_default = input_model.scenario_parameters["day_rate_eur"][
        scenario.lower()
    ].value

    utilization_override = st.sidebar.number_input(
        "Utilization override",
        value=float(utilization_default),
        step=0.01,
        format="%.2f",
    )
    day_rate_override = st.sidebar.number_input(
        "Day rate override (EUR)",
        value=float(day_rate_default),
        step=100.0,
        format="%.0f",
    )

    input_model.scenario_parameters["utilization_rate"][
        scenario.lower()
    ].value = utilization_override
    input_model.scenario_parameters["day_rate_eur"][
        scenario.lower()
    ].value = day_rate_override

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
        input_model, cashflow_result
    )

    tab_pnl, tab_cashflow, tab_debt, tab_equity = st.tabs(
        ["P&L", "Cashflow", "Debt", "Equity"]
    )

    with tab_pnl:
        pnl_table = pd.DataFrame.from_dict(pnl_result, orient="index")
        st.table(pnl_table)

    with tab_cashflow:
        cashflow_table = pd.DataFrame(cashflow_result)
        st.table(cashflow_table)

    with tab_debt:
        debt_table = pd.DataFrame(debt_schedule)
        st.table(debt_table)

    with tab_equity:
        summary = {
            "initial_equity": investment_result["initial_equity"],
            "exit_value": investment_result["exit_value"],
            "irr": investment_result["irr"],
        }
        summary_table = pd.DataFrame([summary])
        cashflows_table = pd.DataFrame(
            {"equity_cashflows": investment_result["equity_cashflows"]}
        )
        st.table(summary_table)
        st.table(cashflows_table)


if __name__ == "__main__":
    run_app()
