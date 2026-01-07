def calculate_pnl(input_model):
    """
    Calculate a 5-year plan P&L based strictly on InputModel inputs.
    Returns a dictionary keyed by year index (0-4).
    """
    planning_horizon_years = 5

    selected_scenario = input_model.scenario_selection["selected_scenario"].value
    utilization_rate = input_model.scenario_parameters["utilization_rate"][
        selected_scenario.lower()
    ].value
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
    billable_days_per_year = input_model.operating_assumptions[
        "billable_days_per_year"
    ].value
    day_rate_growth_pct = input_model.operating_assumptions[
        "day_rate_growth_pct"
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
    bonus_pct_of_base = input_model.personnel_cost_assumptions[
        "bonus_pct_of_base"
    ].value
    payroll_burden_pct_of_comp = input_model.personnel_cost_assumptions[
        "payroll_burden_pct_of_comp"
    ].value
    wage_inflation_pct = input_model.personnel_cost_assumptions[
        "wage_inflation_pct"
    ].value

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

    # Capex and working capital assumptions.
    depreciation_eur_per_year = input_model.capex_and_working_capital[
        "depreciation_eur_per_year"
    ].value

    # Tax assumptions.
    tax_rate_pct = input_model.tax_and_distributions["tax_rate_pct"].value

    pnl_by_year = {}

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

        # Revenue calculation based on scenario utilization and billable days.
        revenue = (
            consultants_fte
            * utilization_rate
            * billable_days_per_year
            * current_daily_rate
        )

        # Consultant personnel cost: base + bonus + payroll burden, inflated.
        consultant_cost_per_fte = avg_consultant_base_cost_eur_per_year * (
            (1 + bonus_pct_of_base) + payroll_burden_pct_of_comp
        )
        consultant_cost_per_fte *= (1 + wage_inflation_pct) ** year_index
        consultant_personnel_cost = consultants_fte * consultant_cost_per_fte

        # Backoffice personnel cost: salary plus payroll burden, inflated.
        backoffice_cost_per_fte = (
            avg_backoffice_salary_eur_per_year
            * (1 + payroll_burden_pct_of_comp)
        )
        backoffice_cost_per_fte *= (1 + wage_inflation_pct) ** year_index
        backoffice_personnel_cost = backoffice_fte * backoffice_cost_per_fte

        # Managing Directors: not defined in inputs; set to zero for v1.
        managing_directors_cost = 0

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
        ebit = ebitda - depreciation_eur_per_year

        # Taxes apply to positive EBIT only.
        taxable_income = ebit if ebit > 0 else 0
        taxes = taxable_income * tax_rate_pct
        net_income = ebit - taxes

        pnl_by_year[f"Year {year_index}"] = {
            "revenue": revenue,
            "personnel_costs": total_personnel_costs,
            "overhead_and_variable_costs": overhead_and_variable_costs,
            "ebitda": ebitda,
            "depreciation": depreciation_eur_per_year,
            "ebit": ebit,
            "taxes": taxes,
            "net_income": net_income,
        }

    return pnl_by_year
