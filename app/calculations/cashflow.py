def calculate_cashflow(input_model, pnl_result, debt_schedule):
    """
    Build a simple multi-year cash flow statement from P&L results.
    Returns a list of yearly dictionaries with a running cash balance.
    """
    financing_assumptions = getattr(input_model, "financing_assumptions", {})

    # Map legacy financing fields to Excel-equivalent transaction inputs.
    debt_amount = financing_assumptions.get(
        "initial_debt_eur",
        input_model.transaction_and_financing[
            "senior_term_loan_start_eur"
        ].value,
    )

    cashflow_assumptions = getattr(input_model, "cashflow_assumptions", {})
    tax_cash_rate_pct = cashflow_assumptions.get(
        "tax_cash_rate_pct",
        input_model.tax_and_distributions["tax_rate_pct"].value,
    )
    tax_payment_lag_years = cashflow_assumptions.get(
        "tax_payment_lag_years", 0
    )
    capex_pct_revenue = cashflow_assumptions.get(
        "capex_pct_revenue", 0.0
    )
    working_capital_pct_revenue = cashflow_assumptions.get(
        "working_capital_pct_revenue", 0.0
    )
    opening_cash_balance = cashflow_assumptions.get(
        "opening_cash_balance_eur", 0.0
    )

    cashflow = []
    cash_balance = opening_cash_balance
    taxes_due_by_year = []
    debt_by_year = {row["year"]: row for row in debt_schedule}
    depreciation_rate = getattr(
        input_model, "balance_sheet_assumptions", {}
    ).get("depreciation_rate_pct", 0.0)
    fixed_assets = 0.0
    working_capital_balance = 0.0
    # Step through each P&L year and derive cash flow lines.
    for i, year_data in enumerate(pnl_result):
        year = year_data["year"]
        revenue = year_data.get("revenue", 0)
        ebitda = year_data.get("ebitda", 0)

        # Debt schedule is the single source of truth for interest and repayment.
        debt_row = debt_by_year.get(year, {})
        interest = debt_row.get("interest_expense", 0.0)
        principal_repayment = debt_row.get("total_repayment", 0.0)
        debt_drawdown = debt_row.get("debt_drawdown", 0.0)

        # Working capital is a balance; cash impact is the year-over-year change.
        working_capital_current = revenue * working_capital_pct_revenue
        working_capital_change = (
            working_capital_current - working_capital_balance
        )
        working_capital_balance = working_capital_current
        capex = revenue * capex_pct_revenue

        # Depreciation is derived from capex and fixed asset roll-forward.
        depreciation = (fixed_assets + capex) * depreciation_rate
        fixed_assets = max(fixed_assets + capex - depreciation, 0.0)

        # EBIT uses EBITDA and cashflow-derived depreciation.
        ebit = ebitda - depreciation

        # Taxes are cash-based on EBT, with an optional payment lag.
        ebt = ebit - interest
        taxes_due = max(ebt, 0) * tax_cash_rate_pct
        taxes_due_by_year.append(taxes_due)
        if tax_payment_lag_years == 0:
            taxes_paid = taxes_due
        elif tax_payment_lag_years == 1:
            taxes_paid = taxes_due_by_year[i - 1] if i > 0 else 0.0
        else:
            taxes_paid = 0.0

        # Operating cash flow starts from EBITDA and adjusts for taxes and working capital.
        operating_cf = ebitda - taxes_paid - working_capital_change

        # Investing cash flow is primarily capital expenditures.
        investing_cf = -capex
        free_cashflow = operating_cf + investing_cf

        # Financing cash flow includes only debt drawdown and debt service.
        financing_cf = (
            debt_drawdown - interest - principal_repayment
            if i == 0
            else -(interest + principal_repayment)
        )

        net_cashflow = free_cashflow + financing_cf
        opening_cash = cash_balance
        cash_balance += net_cashflow

        cashflow.append(
            {
                "year": year,
                "ebitda": ebitda,
                "depreciation": depreciation,
                "taxes_paid": taxes_paid,
                "working_capital_change": working_capital_change,
                "working_capital_balance": working_capital_balance,
                "operating_cf": operating_cf,
                "capex": capex,
                "free_cashflow": free_cashflow,
                "debt_drawdown": debt_drawdown,
                "interest_paid": interest,
                "debt_repayment": principal_repayment,
                "investing_cf": investing_cf,
                "financing_cf": financing_cf,
                "net_cashflow": net_cashflow,
                "opening_cash": opening_cash,
                "cash_balance": cash_balance,
            }
        )

    return cashflow
