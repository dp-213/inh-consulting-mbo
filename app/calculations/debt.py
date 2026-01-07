def calculate_debt_schedule(input_model, cashflow_result):
    """
    Build a simple debt schedule using cash flow results.
    Returns a list of yearly dictionaries.
    """
    initial_debt = input_model.financing["debt_amount"]
    interest_rate = input_model.financing["interest_rate"]
    amortization_rate = input_model.financing["amortization_rate"]

    schedule = []
    outstanding_principal = initial_debt

    # Calculate yearly interest, amortization, and DSCR.
    for i, year_data in enumerate(cashflow_result):
        year = year_data["year"]

        # Interest is based on opening principal.
        interest_expense = outstanding_principal * interest_rate
        principal_payment = outstanding_principal * amortization_rate
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
