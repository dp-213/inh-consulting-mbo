def calculate_debt_schedule(input_model, cashflow_result):
    """
    Build a simple debt schedule using cash flow results.
    Returns a list of yearly dictionaries.
    """
    # Map legacy financing fields to Excel-equivalent transaction inputs.
    initial_debt = input_model.transaction_and_financing[
        "senior_term_loan_start_eur"
    ].value
    interest_rate = input_model.transaction_and_financing[
        "senior_interest_rate_pct"
    ].value
    annual_repayment = input_model.transaction_and_financing[
        "senior_repayment_per_year_eur"
    ].value

    schedule = []
    outstanding_principal = initial_debt

    # Calculate yearly interest, amortization, and DSCR.
    for i, year_data in enumerate(cashflow_result):
        year = year_data["year"]

        # Interest is based on opening principal.
        interest_expense = outstanding_principal * interest_rate
        principal_payment = annual_repayment
        debt_service = interest_expense + principal_payment

        # DSCR uses operating cash flow divided by total debt service.
        operating_cf = year_data["operating_cf"]
        dscr = operating_cf / debt_service if debt_service != 0 else 0

        # Reduce principal after payment.
        outstanding_principal -= principal_payment
        if outstanding_principal < 0:
            outstanding_principal = 0

        schedule.append(
            {
                "year": year,
                "interest_expense": interest_expense,
                "principal_payment": principal_payment,
                "debt_service": debt_service,
                "outstanding_principal": outstanding_principal,
                "dscr": dscr,
            }
        )

    return schedule
