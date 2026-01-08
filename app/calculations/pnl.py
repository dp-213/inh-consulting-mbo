def calculate_pnl(input_model, depreciation_by_year=None):
    """
    Calculate a 5-year plan P&L based strictly on InputModel inputs.
    Returns a list of yearly dictionaries with integer year indices.
    """
    planning_horizon_years = 5

    selected_scenario = input_model.scenario_selection["selected_scenario"].value
    utilization_rate = input_model.scenario_parameters["utilization_rate"][
        selected_scenario.lower()
    ].value
    utilization_by_year = getattr(input_model, "utilization_by_year", None)
    daily_rate = input_model.scenario_parameters["day_rate_eur"][
        selected_scenario.lower()
    ].value

    # Operating assumptions.
    consultants_fte_start = input_model.operating_assumptions[
        "consulting_fte_start"
    ].value
    consultants_fte_growth_pct = input_model.operating_assumptions[
        "consulting_fte_growth_pct"
    ].value
    work_days_per_year = input_model.operating_assumptions[
        "work_days_per_year"
    ].value
    day_rate_growth_pct = input_model.operating_assumptions[
        "day_rate_growth_pct"
    ].value
    guarantee_year_1 = input_model.operating_assumptions[
        "revenue_guarantee_pct_year_1"
    ].value
    guarantee_year_2 = input_model.operating_assumptions[
        "revenue_guarantee_pct_year_2"
    ].value
    guarantee_year_3 = input_model.operating_assumptions[
        "revenue_guarantee_pct_year_3"
    ].value

    backoffice_fte_start = input_model.operating_assumptions[
        "backoffice_fte_start"
    ].value
    backoffice_fte_growth_pct = input_model.operating_assumptions[
        "backoffice_fte_growth_pct"
    ].value
    avg_backoffice_salary_eur_per_year = input_model.operating_assumptions[
        "avg_backoffice_salary_eur_per_year"
    ].value

    # Personnel cost assumptions.
    avg_consultant_base_cost_eur_per_year = input_model.personnel_cost_assumptions[
        "avg_consultant_base_cost_eur_per_year"
    ].value
    wage_inflation_pct = input_model.personnel_cost_assumptions[
        "wage_inflation_pct"
    ].value
    management_md_cost = getattr(
        input_model, "management_md_cost_eur_per_year", 0.0
    )
    management_md_growth = getattr(
        input_model, "management_md_cost_growth_pct", 0.0
    )

    # Overhead and variable costs.
    rent_eur_per_year = input_model.overhead_and_variable_costs[
        "rent_eur_per_year"
    ].value
    it_and_software_eur_per_year = input_model.overhead_and_variable_costs[
        "it_and_software_eur_per_year"
    ].value
    overhead_inflation_pct = input_model.overhead_and_variable_costs[
        "overhead_inflation_pct"
    ].value
    insurance_eur_per_year = input_model.overhead_and_variable_costs[
        "insurance_eur_per_year"
    ].value
    legal_audit_eur_per_year = input_model.overhead_and_variable_costs[
        "legal_audit_eur_per_year"
    ].value
    other_overhead_eur_per_year = input_model.overhead_and_variable_costs[
        "other_overhead_eur_per_year"
    ].value
    travel_pct_of_revenue = input_model.overhead_and_variable_costs[
        "travel_pct_of_revenue"
    ].value
    recruiting_pct_of_revenue = input_model.overhead_and_variable_costs[
        "recruiting_pct_of_revenue"
    ].value
    training_pct_of_revenue = input_model.overhead_and_variable_costs[
        "training_pct_of_revenue"
    ].value
    marketing_pct_of_revenue = input_model.overhead_and_variable_costs[
        "marketing_pct_of_revenue"
    ].value

    # Tax assumptions.
    tax_rate_pct = input_model.tax_and_distributions["tax_rate_pct"].value

    pnl_by_year = []

    reference_volume_eur = 20_000_000
    if isinstance(revenue_model, dict):
        reference_volume_eur = revenue_model[
            "reference_revenue_eur"
        ].value

    revenue_by_year = getattr(input_model, "revenue_by_year", None)
    revenue_model = getattr(input_model, "revenue_model", None)

    for year_index in range(planning_horizon_years):
        # Grow FTEs and rates by their respective growth assumptions.
        consultants_fte = consultants_fte_start * (
            (1 + consultants_fte_growth_pct) ** year_index
        )
        backoffice_fte = backoffice_fte_start * (
            (1 + backoffice_fte_growth_pct) ** year_index
        )
        current_daily_rate = daily_rate * (
            (1 + day_rate_growth_pct) ** year_index
        )

        current_utilization = (
            utilization_by_year[year_index]
            if isinstance(utilization_by_year, list)
            and len(utilization_by_year) == planning_horizon_years
            else utilization_rate
        )

        # Revenue comes from the dedicated revenue model if present.
        if isinstance(revenue_by_year, list) and len(revenue_by_year) > year_index:
            total_revenue = revenue_by_year[year_index]
        else:
            total_revenue = (
                consultants_fte
                * current_utilization
                * work_days_per_year
                * current_daily_rate
            )
        guarantee_pct = 0
        if isinstance(revenue_model, dict):
            guarantee_pct = revenue_model[
                f"guarantee_pct_year_{year_index}"
            ].value
        else:
            if year_index == 0:
                guarantee_pct = guarantee_year_1
            elif year_index == 1:
                guarantee_pct = guarantee_year_2
            elif year_index == 2:
                guarantee_pct = guarantee_year_3

        # Revenue guarantees classify revenue; they do not change total revenue.
        guaranteed_revenue = min(
            total_revenue, guarantee_pct * reference_volume_eur
        )
        non_guaranteed_revenue = total_revenue - guaranteed_revenue
        revenue = total_revenue

        # Consultant personnel cost: base + bonus + payroll burden, inflated.
        # Consultant all-in cost per FTE drives compensation directly.
        consultant_cost_per_fte = avg_consultant_base_cost_eur_per_year
        consultant_cost_per_fte *= (1 + wage_inflation_pct) ** year_index
        consultant_personnel_cost = consultants_fte * consultant_cost_per_fte

        # Backoffice personnel cost: all-in cost per FTE with inflation.
        backoffice_cost_per_fte = avg_backoffice_salary_eur_per_year
        backoffice_cost_per_fte *= (1 + wage_inflation_pct) ** year_index
        backoffice_personnel_cost = backoffice_fte * backoffice_cost_per_fte

        # Management / MD cost from assumptions with annual growth.
        managing_directors_cost = management_md_cost * (
            (1 + management_md_growth) ** year_index
        )

        total_personnel_costs = (
            consultant_personnel_cost
            + backoffice_personnel_cost
            + managing_directors_cost
        )

        # Fixed overhead costs with inflation.
        fixed_overhead = (
            rent_eur_per_year
            + it_and_software_eur_per_year
            + insurance_eur_per_year
            + legal_audit_eur_per_year
            + other_overhead_eur_per_year
        )
        fixed_overhead *= (1 + overhead_inflation_pct) ** year_index

        # Variable overhead costs based on revenue.
        variable_overhead = revenue * (
            travel_pct_of_revenue
            + recruiting_pct_of_revenue
            + training_pct_of_revenue
            + marketing_pct_of_revenue
        )

        overhead_and_variable_costs = fixed_overhead + variable_overhead

        # EBITDA and EBIT.
        ebitda = revenue - total_personnel_costs - overhead_and_variable_costs
        if isinstance(depreciation_by_year, dict):
            depreciation = depreciation_by_year.get(year_index, 0.0)
        elif (
            isinstance(depreciation_by_year, list)
            and len(depreciation_by_year) > year_index
        ):
            depreciation = depreciation_by_year[year_index]
        else:
            depreciation = 0.0
        ebit = ebitda - depreciation

        # Taxes apply to positive EBIT only.
        taxable_income = ebit if ebit > 0 else 0
        taxes = taxable_income * tax_rate_pct
        net_income = ebit - taxes

        pnl_by_year.append(
            {
                "year": year_index,
                "revenue": revenue,
                "personnel_costs": total_personnel_costs,
                "overhead_and_variable_costs": overhead_and_variable_costs,
                "ebitda": ebitda,
                "depreciation": depreciation,
                "ebit": ebit,
                "taxes": taxes,
                "net_income": net_income,
            }
        )

    return pnl_by_year
