def calculate_cashflow(input_model, pnl_result):
    """
    Build a simple multi-year cash flow statement from P&L results.
    Returns a list of yearly dictionaries with a running cash balance.
    """
    # Map legacy financing fields to Excel-equivalent transaction inputs.
    debt_amount = input_model.transaction_and_financing[
        "senior_term_loan_start_eur"
    ].value
    interest_rate = input_model.transaction_and_financing[
        "senior_interest_rate_pct"
    ].value
    annual_repayment = input_model.transaction_and_financing[
        "senior_repayment_per_year_eur"
    ].value
    equity_amount = input_model.transaction_and_financing[
        "equity_contribution_eur"
    ].value

    # Map operating inputs to Excel-equivalent capex and working capital fields.
    capex = input_model.capex_and_working_capital["capex_eur_per_year"].value
    # No explicit working capital change input in v1; keep as zero for now.
    working_capital_change = 0.0

    cashflow = []
    cash_balance = 0.0

    # Step through each P&L year and derive cash flow lines.
    for i, year_data in enumerate(pnl_result):
        year = year_data["year"]
        net_income = year_data["net_income"]

        # Operating cash flow starts from net income and adjusts for working capital.
        operating_cf = net_income - working_capital_change

        # Investing cash flow is primarily capital expenditures.
        investing_cf = -capex

        # Financing cash flow includes interest, debt repayment, and initial funding.
        interest = debt_amount * interest_rate
        principal_repayment = annual_repayment
        financing_cf = -(interest + principal_repayment)

        # In the first year, add initial debt and equity funding.
        if i == 0:
            financing_cf += debt_amount + equity_amount

        net_cashflow = operating_cf + investing_cf + financing_cf
        cash_balance += net_cashflow

        cashflow.append(
            {
                "year": year,
                "operating_cf": operating_cf,
                "investing_cf": investing_cf,
                "financing_cf": financing_cf,
                "net_cashflow": net_cashflow,
                "cash_balance": cash_balance,
            }
        )

    return cashflow
