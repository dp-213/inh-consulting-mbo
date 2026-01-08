from data_model import InputModel
from calculations.pnl import calculate_pnl
from calculations.cashflow import calculate_cashflow
from calculations.debt import calculate_debt_schedule
from calculations.balance_sheet import calculate_balance_sheet
from calculations.investment import calculate_investment


def run_model():
    # Create the central input model.
    input_model = InputModel()

    # Run the integrated model in the required order.
    pnl_result = calculate_pnl(input_model)
    cashflow_result = calculate_cashflow(input_model, pnl_result)
    debt_schedule = calculate_debt_schedule(input_model, cashflow_result)
    balance_sheet = calculate_balance_sheet(
        input_model, cashflow_result, debt_schedule
    )
    investment_result = calculate_investment(
        input_model, cashflow_result, pnl_result
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
