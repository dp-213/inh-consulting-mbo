def calculate_pnl(input_model, depreciation_by_year=None, revenue_final_by_year=None, cost_totals_by_year=None):
    """
    Calculate a 5-year plan P&L based strictly on InputModel inputs.
    Returns a list of yearly dictionaries with integer year indices.
    """
    planning_horizon_years = 5

    # Tax assumptions.
    tax_rate_pct = input_model.tax_and_distributions["tax_rate_pct"].value

    pnl_by_year = []

    if not isinstance(revenue_final_by_year, list) or len(revenue_final_by_year) != planning_horizon_years:
        raise ValueError("revenue_final_by_year must be a 5-year list.")
    if not isinstance(cost_totals_by_year, list) or len(cost_totals_by_year) != planning_horizon_years:
        raise ValueError("cost_totals_by_year must be a 5-year list.")

    for year_index in range(planning_horizon_years):
        revenue = revenue_final_by_year[year_index]
        year_costs = cost_totals_by_year[year_index]
        total_personnel_costs = year_costs.get("personnel_costs", 0.0)
        overhead_and_variable_costs = year_costs.get("overhead_and_variable_costs", 0.0)

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
