from data_model import InputModel
from calculations.pnl import calculate_pnl
from calculations.cashflow import calculate_cashflow
from calculations.debt import calculate_debt_schedule
from calculations.balance_sheet import calculate_balance_sheet
from calculations.investment import calculate_investment


def run_model():
    # Create the central input model.
    input_model = InputModel()

    # Build revenue by year from the revenue model inputs.
    revenue_model = getattr(input_model, "revenue_model", {})
    reference_revenue = revenue_model.get(
        "reference_revenue_eur"
    ).value
    revenue_by_year = []
    for year_index in range(5):
        guarantee_pct = revenue_model.get(
            f"guarantee_pct_year_{year_index}"
        ).value
        in_group = revenue_model.get(
            f"in_group_revenue_year_{year_index}"
        ).value
        external = revenue_model.get(
            f"external_revenue_year_{year_index}"
        ).value
        modeled_revenue = in_group + external
        guaranteed = min(
            modeled_revenue, guarantee_pct * reference_revenue
        )
        final_revenue = max(guaranteed, modeled_revenue)
        revenue_by_year.append(final_revenue)
    input_model.revenue_by_year = revenue_by_year

    # Build cost model totals (authoritative cost inputs).
    cost_model = getattr(input_model, "cost_model", {})
    wage_inflation = input_model.personnel_cost_assumptions[
        "wage_inflation_pct"
    ].value
    bonus_pct_default = input_model.personnel_cost_assumptions[
        "bonus_pct_of_base"
    ].value
    payroll_pct_default = input_model.personnel_cost_assumptions[
        "payroll_burden_pct_of_comp"
    ].value
    management_cost = getattr(
        input_model, "management_md_cost_eur_per_year", 0.0
    )
    management_growth = getattr(
        input_model, "management_md_cost_growth_pct", 0.0
    )
    cost_model_totals = []
    for year_index in range(5):
        revenue = revenue_by_year[year_index]
        consultant_fte = cost_model.get(
            f"consultant_fte_year_{year_index}"
        ).value
        consultant_base = cost_model.get(
            f"consultant_base_cost_eur_year_{year_index}"
        ).value * ((1 + wage_inflation) ** year_index)
        consultant_bonus = cost_model.get(
            f"consultant_bonus_pct_year_{year_index}"
        ).value if cost_model else bonus_pct_default
        consultant_payroll = cost_model.get(
            f"consultant_payroll_pct_year_{year_index}"
        ).value if cost_model else payroll_pct_default
        consultant_loaded = consultant_base * (
            1 + consultant_bonus + consultant_payroll
        )
        consultant_total = consultant_fte * consultant_loaded

        backoffice_fte = cost_model.get(
            f"backoffice_fte_year_{year_index}"
        ).value
        backoffice_base = cost_model.get(
            f"backoffice_base_cost_eur_year_{year_index}"
        ).value * ((1 + wage_inflation) ** year_index)
        backoffice_payroll = cost_model.get(
            f"backoffice_payroll_pct_year_{year_index}"
        ).value if cost_model else payroll_pct_default
        backoffice_loaded = backoffice_base * (1 + backoffice_payroll)
        backoffice_total = backoffice_fte * backoffice_loaded

        managing_directors_cost = management_cost * (
            (1 + management_growth) ** year_index
        )

        fixed_overhead = (
            cost_model.get(f"fixed_overhead_advisory_year_{year_index}").value
            + cost_model.get(f"fixed_overhead_legal_year_{year_index}").value
            + cost_model.get(f"fixed_overhead_it_year_{year_index}").value
            + cost_model.get(f"fixed_overhead_office_year_{year_index}").value
            + cost_model.get(f"fixed_overhead_services_year_{year_index}").value
        )

        def _variable_cost(prefix):
            pct = cost_model.get(f"variable_{prefix}_pct_year_{year_index}").value
            eur = cost_model.get(f"variable_{prefix}_eur_year_{year_index}").value
            return eur if eur > 0 else revenue * pct

        variable_total = (
            _variable_cost("training")
            + _variable_cost("travel")
            + _variable_cost("communication")
        )
        personnel_total = consultant_total + backoffice_total + managing_directors_cost
        overhead_total = fixed_overhead + variable_total
        cost_model_totals.append(
            {
                "consultant_costs": consultant_total,
                "backoffice_costs": backoffice_total,
                "management_costs": managing_directors_cost,
                "personnel_costs": personnel_total,
                "overhead_and_variable_costs": overhead_total,
                "total_operating_costs": personnel_total + overhead_total,
            }
        )
    input_model.cost_model_totals_by_year = cost_model_totals

    # Run the integrated model in the required order.
    pnl_base = calculate_pnl(input_model)
    debt_schedule = calculate_debt_schedule(input_model)
    cashflow_result = calculate_cashflow(input_model, pnl_base, debt_schedule)
    depreciation_by_year = {
        row["year"]: row.get("depreciation", 0.0) for row in cashflow_result
    }
    pnl_result = calculate_pnl(input_model, depreciation_by_year)
    debt_schedule = calculate_debt_schedule(input_model, cashflow_result)
    balance_sheet = calculate_balance_sheet(
        input_model, cashflow_result, debt_schedule, pnl_result
    )
    investment_result = calculate_investment(
        input_model, cashflow_result, pnl_result, balance_sheet
    )

    # Collect all outputs in one dictionary.
    model_results = {
        "pnl": pnl_result,
        "cashflow": cashflow_result,
        "debt_schedule": debt_schedule,
        "balance_sheet": balance_sheet,
        "investment": investment_result,
    }

    return model_results
