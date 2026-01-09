def calculate_debt_schedule(input_model, cashflow_result=None):
    """
    Build a simple debt schedule using cash flow results.
    Returns a list of yearly dictionaries.
    """
    # Map legacy financing fields to Excel-equivalent transaction inputs.
    financing_assumptions = getattr(input_model, "financing_assumptions", {})
    if "senior_debt_amount" not in financing_assumptions:
        raise ValueError("Senior debt amount missing from financing assumptions.")
    initial_debt = financing_assumptions["senior_debt_amount"]
    interest_rate = financing_assumptions.get("interest_rate_pct")
    if interest_rate is None:
        raise ValueError("Interest rate missing from financing assumptions.")
    amort_type = financing_assumptions.get("amortization_type", "Linear")
    amort_period = financing_assumptions.get("amortization_period_years", 5)
    grace_period = financing_assumptions.get("grace_period_years", 0)
    special_year = financing_assumptions.get("special_repayment_year", None)
    special_amount = financing_assumptions.get("special_repayment_amount_eur", 0.0)
    min_dscr = financing_assumptions.get("minimum_dscr", 1.3)

    schedule = []
    outstanding_principal = 0.0

    # Calculate yearly interest and amortization.
    for i in range(5):
        year = i
        debt_drawdown = initial_debt if i == 0 else 0.0
        opening_debt = outstanding_principal + debt_drawdown

        # Interest is based on opening principal (includes Year 0 drawdown).
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
                    opening_debt / amort_period
                    if i < amort_period
                    else 0.0
                )
            )
        if (
            amort_type != "Bullet"
            and i >= grace_period
            and opening_debt > 0
            and amort_period
        ):
            expected_repayment = opening_debt / amort_period
            if abs(scheduled_repayment - expected_repayment) > 1e-6:
                raise ValueError(
                    "Scheduled repayment does not scale with senior debt amount."
                )
        special_repayment = (
            special_amount if special_year == i else 0.0
        )
        pre_repayment_balance = opening_debt
        total_repayment = min(
            pre_repayment_balance, scheduled_repayment + special_repayment
        )
        debt_service = interest_expense + total_repayment
        if initial_debt > 0 and opening_debt > 0:
            if interest_expense == 0:
                raise ValueError(
                    "Interest expense is zero with positive debt balance."
                )
            if scheduled_repayment == 0:
                raise ValueError(
                    "Scheduled repayment is zero with positive debt balance."
                )
            if debt_service == 0:
                raise ValueError(
                    "Debt service is zero with positive debt balance."
                )

        # DSCR will be added when cashflow data is available.
        cfads = None
        dscr = None
        covenant_breach = None

        # Reduce principal after payment.
        outstanding_principal = max(
            pre_repayment_balance - total_repayment, 0.0
        )

        reconciliation_gap = (
            opening_debt - total_repayment - outstanding_principal
        )
        if abs(reconciliation_gap) > 1e-6:
            raise ValueError(
                f"Debt reconciliation failed in year {year}: {reconciliation_gap}"
            )

        schedule.append(
            {
                "year": year,
                "opening_debt": opening_debt,
                "debt_drawdown": debt_drawdown,
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
                "covenant_breach": covenant_breach,
            }
        )

    if cashflow_result is not None:
        cashflow_by_year = {row["year"]: row for row in cashflow_result}
        for row in schedule:
            year = row["year"]
            year_data = cashflow_by_year.get(year, {})
            operating_cf = year_data.get("operating_cf", 0.0)
            cfads = operating_cf - year_data.get("capex", 0.0)
            debt_service = row["debt_service"]
            dscr = cfads / debt_service if debt_service != 0 else 0
            row["cfads"] = cfads
            row["dscr"] = dscr
            row["covenant_breach"] = dscr < min_dscr

    return schedule
