def calculate_balance_sheet(input_model, cashflow_result, debt_schedule, pnl_result=None):
    """
    Build a simplified multi-year balance sheet.
    Returns a list of yearly dictionaries.
    """
    balance_sheet = []
    working_capital_balance = 0.0
    retained_earnings = 0.0

    # Build a simple net income lookup if P&L results are provided.
    net_income_by_year = {}
    if pnl_result is not None:
        if isinstance(pnl_result, dict):
            for year_label, year_data in pnl_result.items():
                try:
                    year_key = int(str(year_label).split()[-1])
                except (ValueError, IndexError):
                    year_key = year_label
                net_income_by_year[year_key] = year_data.get("net_income", 0)
        else:
            for year_data in pnl_result:
                net_income_by_year[year_data.get("year")] = year_data.get(
                    "net_income", 0
                )

    for year_data, debt_data in zip(cashflow_result, debt_schedule):
        year = year_data["year"]
        cash_balance = year_data["cash_balance"]
        outstanding_principal = debt_data["outstanding_principal"]

        # Working capital change is derived from cashflow and net income when available.
        net_income = net_income_by_year.get(year, 0)
        working_capital_change = net_income - year_data["operating_cf"]
        working_capital_balance += working_capital_change

        # Retained earnings accumulate net income from the P&L.
        retained_earnings += net_income

        # Assets are limited to cash and working capital in this simplified model.
        assets = cash_balance + working_capital_balance

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
                "working_capital": working_capital_balance,
                "retained_earnings": retained_earnings,
            }
        )

    return balance_sheet
