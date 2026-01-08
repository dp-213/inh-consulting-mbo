def calculate_debt_schedule(input_model, cashflow_result):
    """
    Build a simple debt schedule using cash flow results.
    Returns a list of yearly dictionaries.
    """
    # Map legacy financing fields to Excel-equivalent transaction inputs.
    financing_assumptions = getattr(input_model, "financing_assumptions", {})
    initial_debt = financing_assumptions.get(
        "initial_debt_eur",
        input_model.transaction_and_financing[
            "senior_term_loan_start_eur"
        ].value,
    )
    interest_rate = financing_assumptions.get(
        "interest_rate_pct",
        input_model.transaction_and_financing[
            "senior_interest_rate_pct"
        ].value,
    )
    amort_type = financing_assumptions.get("amortization_type", "Linear")
    amort_period = financing_assumptions.get("amortization_period_years", 5)
    grace_period = financing_assumptions.get("grace_period_years", 0)
    special_year = financing_assumptions.get("special_repayment_year", None)
    special_amount = financing_assumptions.get("special_repayment_amount_eur", 0.0)
    min_dscr = financing_assumptions.get("minimum_dscr", 1.3)

    schedule = []
    outstanding_principal = initial_debt

    # Calculate yearly interest, amortization, and DSCR.
    for i, year_data in enumerate(cashflow_result):
        year = year_data["year"]
        opening_debt = outstanding_principal

        # Interest is based on opening principal.
        interest_expense = opening_debt * interest_rate
        if amort_type == "Bullet":
            scheduled_repayment = (
                opening_debt if i == max(amort_period - 1, 0) else 0.0
            )
        else:
            scheduled_repayment = (
                0.0
                if i < grace_period
                else (
                    initial_debt / amort_period
                    if i < amort_period
                    else 0.0
                )
            )
        special_repayment = (
            special_amount if special_year == i else 0.0
        )
        total_repayment = min(
            opening_debt, scheduled_repayment + special_repayment
        )
        debt_service = interest_expense + total_repayment

        # DSCR uses operating cash flow divided by total debt service.
        operating_cf = year_data["operating_cf"]
        cfads = operating_cf - year_data.get("capex", 0.0)
        dscr = cfads / debt_service if debt_service != 0 else 0

        # Reduce principal after payment.
        outstanding_principal = max(opening_debt - total_repayment, 0.0)

        schedule.append(
            {
                "year": year,
                "opening_debt": opening_debt,
                "debt_drawdown": initial_debt if i == 0 else 0.0,
                "scheduled_repayment": scheduled_repayment,
                "special_repayment": special_repayment,
                "total_repayment": total_repayment,
                "closing_debt": outstanding_principal,
                "interest_expense": interest_expense,
                "principal_payment": total_repayment,
                "debt_service": debt_service,
                "outstanding_principal": outstanding_principal,
                "dscr": dscr,
                "cfads": cfads,
                "minimum_dscr": min_dscr,
                "covenant_breach": dscr < min_dscr,
            }
        )

    return schedule
