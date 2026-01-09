from data_model import InputModel
from calculations.pnl import calculate_pnl
from calculations.cashflow import calculate_cashflow
from calculations.debt import calculate_debt_schedule
from calculations.balance_sheet import calculate_balance_sheet
from calculations.investment import calculate_investment
from revenue_model import build_revenue_model_outputs
from cost_model import build_cost_model_outputs


def run_model(assumptions_state=None, scenario="Base", input_model=None):
    # Create the central input model.
    if input_model is None:
        input_model = InputModel()
    if assumptions_state is None:
        raise ValueError("assumptions_state is required for revenue and cost models.")

    # Revenue and cost models are calculated once in their dedicated modules.
    revenue_final_by_year, revenue_components_by_year = build_revenue_model_outputs(
        assumptions_state, scenario
    )
    cost_model_totals = build_cost_model_outputs(
        assumptions_state, revenue_final_by_year
    )
    input_model.revenue_final_by_year = revenue_final_by_year
    input_model.revenue_components_by_year = revenue_components_by_year
    input_model.cost_model_totals_by_year = cost_model_totals

    # Run the integrated model in the required order.
    debt_schedule = calculate_debt_schedule(input_model)
    pnl_base = calculate_pnl(
        input_model,
        revenue_final_by_year=revenue_final_by_year,
        cost_totals_by_year=cost_model_totals,
        debt_schedule=debt_schedule,
    )
    cashflow_result = calculate_cashflow(input_model, pnl_base, debt_schedule)
    depreciation_by_year = {
        row["year"]: row.get("depreciation", 0.0) for row in cashflow_result
    }
    pnl_result = calculate_pnl(
        input_model,
        depreciation_by_year,
        revenue_final_by_year=revenue_final_by_year,
        cost_totals_by_year=cost_model_totals,
        debt_schedule=debt_schedule,
    )
    debt_schedule = calculate_debt_schedule(input_model, cashflow_result)
    balance_sheet = calculate_balance_sheet(
        input_model, cashflow_result, debt_schedule, pnl_result
    )
    investment_result = calculate_investment(
        input_model, cashflow_result, pnl_result, balance_sheet
    )

    # Collect all outputs in one dictionary.
    model_results = {
        "pnl": pnl_result,
        "cashflow": cashflow_result,
        "debt_schedule": debt_schedule,
        "balance_sheet": balance_sheet,
        "investment": investment_result,
    }

    return model_results
