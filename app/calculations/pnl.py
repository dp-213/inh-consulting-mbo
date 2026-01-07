def calculate_pnl(input_model):
    """
    Build a simple multi-year P&L from InputModel inputs.
    Returns a list of yearly dictionaries.
    """
    start_year = input_model.general["start_year"]
    years = input_model.general["projection_years"]
    base_revenue = input_model.operations["revenue"]
    base_opex = input_model.operations["operating_expenses"]
    growth_rate = input_model.assumptions["growth_rate"]
    tax_rate = input_model.assumptions["tax_rate"]

    pnl = []

    # Step through each projection year and compute basic P&L lines.
    for i in range(years):
        year = start_year + i
        revenue = base_revenue * ((1 + growth_rate) ** i)
        opex = base_opex * ((1 + growth_rate) ** i)

        # EBITDA is revenue minus operating expenses.
        ebitda = revenue - opex
        # EBIT is treated as EBITDA in this simple model.
        ebit = ebitda

        # Taxes apply to positive EBIT only.
        taxable_income = ebit if ebit > 0 else 0
        taxes = taxable_income * tax_rate
        net_income = ebit - taxes

        pnl.append(
            {
                "year": year,
                "revenue": revenue,
                "opex": opex,
                "EBITDA": ebitda,
                "EBIT": ebit,
                "taxes": taxes,
                "net_income": net_income,
            }
        )

    return pnl
