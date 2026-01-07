def calculate_balance_sheet(input_model, cashflow_result, debt_schedule):
    """
    Build a simplified multi-year balance sheet.
    Returns a list of yearly dictionaries.
    """
    # Use a simple static working capital placeholder each year.
    working_capital = input_model.operations["working_capital_change"]

    balance_sheet = []

    for year_data, debt_data in zip(cashflow_result, debt_schedule):
        year = year_data["year"]
        cash_balance = year_data["cash_balance"]
        outstanding_principal = debt_data["outstanding_principal"]

        # Assets are limited to cash and working capital in this simplified model.
        assets = cash_balance + working_capital

        # Liabilities are the remaining debt balance.
        liabilities = outstanding_principal

        # Equity is the residual needed to make the balance sheet balance.
        equity = assets - liabilities

        balance_sheet.append(
            {
                "year": year,
                "assets": assets,
                "liabilities": liabilities,
                "equity": equity,
            }
        )

    return balance_sheet
