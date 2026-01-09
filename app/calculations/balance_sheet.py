def calculate_balance_sheet(
    input_model, cashflow_result, debt_schedule, pnl_result=None
):
    """
    Build a simplified multi-year balance sheet for a consulting model.
    Returns a list of yearly dictionaries.
    """
    balance_sheet = []

    balance_assumptions = getattr(input_model, "balance_sheet_assumptions", {})
    opening_equity = balance_assumptions.get("opening_equity_eur", 0.0)
    equity_contribution = input_model.transaction_and_financing[
        "equity_contribution_eur"
    ].value
    purchase_price = input_model.transaction_and_financing[
        "purchase_price_eur"
    ].value

    # Build a net income and tax lookup from P&L results.
    net_income_by_year = {}
    taxes_by_year = {}
    if pnl_result is not None:
        for year_data in pnl_result:
            net_income_by_year[year_data.get("year")] = year_data.get(
                "net_income", 0
            )
            taxes_by_year[year_data.get("year")] = year_data.get("taxes", 0.0)

    debt_by_year = {
        debt_data["year"]: debt_data["closing_debt"]
        for debt_data in debt_schedule
    }

    equity_start = opening_equity
    fixed_assets = 0.0
    acquisition_intangible = purchase_price
    working_capital_balance = 0.0
    tax_payable_balance = 0.0

    for year_data in cashflow_result:
        year = year_data["year"]
        cash_balance = year_data["cash_balance"]
        capex = year_data.get("capex", 0.0)
        depreciation = year_data.get("depreciation", 0.0)
        working_capital_balance = year_data.get(
            "working_capital_balance", working_capital_balance
        )

        # Fixed assets follow cashflow-derived capex and depreciation.
        fixed_assets = max(fixed_assets + capex - depreciation, 0.0)

        financial_debt = debt_by_year.get(year, 0.0)
        net_income = net_income_by_year.get(year, 0.0)
        dividends = 0.0
        # Model the upfront equity contribution at close in year 0 only.
        equity_injection = equity_contribution if year == 0 else 0.0
        equity_buyback = 0.0
        equity_end = (
            equity_start
            + net_income
            - dividends
            + equity_injection
            - equity_buyback
        )

        taxes_due = taxes_by_year.get(year, 0.0)
        taxes_paid = year_data.get("taxes_paid", 0.0)
        tax_payable_balance += taxes_due - taxes_paid

        total_assets = (
            cash_balance
            + fixed_assets
            + working_capital_balance
            + acquisition_intangible
        )
        total_liabilities = financial_debt + tax_payable_balance
        total_liabilities_equity = total_liabilities + equity_end
        balance_check = total_assets - total_liabilities_equity
        if abs(balance_check) > 1.0:
            raise ValueError(
                f"Balance sheet out of balance in year {year}: {balance_check}"
            )

        balance_sheet.append(
            {
                "year": year,
                "cash": cash_balance,
                "fixed_assets": fixed_assets,
                "acquisition_intangible": acquisition_intangible,
                "working_capital": working_capital_balance,
                "total_assets": total_assets,
                "financial_debt": financial_debt,
                "tax_payable": tax_payable_balance,
                "total_liabilities": total_liabilities,
                "equity_start": equity_start,
                "net_income": net_income,
                "dividends": dividends,
                "equity_injection": equity_injection,
                "equity_buyback": equity_buyback,
                "equity_end": equity_end,
                "total_liabilities_equity": total_liabilities_equity,
                "balance_check": balance_check,
            }
        )

        equity_start = equity_end

    return balance_sheet
