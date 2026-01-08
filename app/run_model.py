from app.data_model import InputModel
from app.calculations.pnl import calculate_pnl
from app.calculations.cashflow import calculate_cashflow
from app.calculations.debt import calculate_debt_schedule
from app.calculations.balance_sheet import calculate_balance_sheet
from app.calculations.investment import calculate_investment


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
    investment_result = calculate_investment(input_model, cashflow_result)

    # Collect all outputs in one dictionary.
    model_results = {
        "pnl": pnl_result,
        "cashflow": cashflow_result,
        "debt_schedule": debt_schedule,
        "balance_sheet": balance_sheet,
        "investment": investment_result,
    }

    return model_results
