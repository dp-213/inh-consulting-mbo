from data_model import InputModel
from calculations.pnl import calculate_pnl
from calculations.cashflow import calculate_cashflow
from calculations.debt import calculate_debt_schedule
from calculations.balance_sheet import calculate_balance_sheet
from calculations.investment import calculate_investment


def run_model():
    # Create the central input model.
    input_model = InputModel()

    # Build revenue by year from the revenue model inputs.
    revenue_model = getattr(input_model, "revenue_model", {})
    reference_revenue = revenue_model.get(
        "reference_revenue_eur"
    ).value
    revenue_by_year = []
    for year_index in range(5):
        guarantee_pct = revenue_model.get(
            f"guarantee_pct_year_{year_index}"
        ).value
        in_group = revenue_model.get(
            f"in_group_revenue_year_{year_index}"
        ).value
        external = revenue_model.get(
            f"external_revenue_year_{year_index}"
        ).value
        modeled_revenue = in_group + external
        guaranteed = min(
            modeled_revenue, guarantee_pct * reference_revenue
        )
        final_revenue = max(guaranteed, modeled_revenue)
        revenue_by_year.append(final_revenue)
    input_model.revenue_by_year = revenue_by_year

    # Run the integrated model in the required order.
    pnl_base = calculate_pnl(input_model)
    debt_schedule = calculate_debt_schedule(input_model)
    cashflow_result = calculate_cashflow(input_model, pnl_base, debt_schedule)
    depreciation_by_year = {
        row["year"]: row.get("depreciation", 0.0) for row in cashflow_result
    }
    pnl_result = calculate_pnl(input_model, depreciation_by_year)
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
