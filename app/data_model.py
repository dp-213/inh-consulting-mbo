class InputField:
    def __init__(self, value, description, excel_ref, editable):
        self.value = value
        self.description = description
        self.excel_ref = excel_ref
        self.editable = editable


class InputModel:
    def __init__(self):
        # Scenario selection: choose which scenario drives the model.
        self.scenario_selection = {
            "selected_scenario": InputField(
                value="Base",
                description="Scenario selector for Base/Best/Worst cases",
                excel_ref="00_Inputs_Assumptions!B3",
                editable=True,
            )
        }

        # Scenario parameters: key revenue drivers by scenario.
        self.scenario_parameters = {
            "utilization_rate": {
                "base": InputField(
                    value=0.70,
                    description="Utilization rate for Base scenario",
                    excel_ref="00_Inputs_Assumptions!B7",
                    editable=True,
                ),
                "best": InputField(
                    value=0.80,
                    description="Utilization rate for Best scenario",
                    excel_ref="00_Inputs_Assumptions!C7",
                    editable=True,
                ),
                "worst": InputField(
                    value=0.55,
                    description="Utilization rate for Worst scenario",
                    excel_ref="00_Inputs_Assumptions!D7",
                    editable=True,
                ),
            },
            "day_rate_eur": {
                "base": InputField(
                    value=2750,
                    description="Day rate (EUR) for Base scenario",
                    excel_ref="00_Inputs_Assumptions!B8",
                    editable=True,
                ),
                "best": InputField(
                    value=3500,
                    description="Day rate (EUR) for Best scenario",
                    excel_ref="00_Inputs_Assumptions!C8",
                    editable=True,
                ),
                "worst": InputField(
                    value=1800,
                    description="Day rate (EUR) for Worst scenario",
                    excel_ref="00_Inputs_Assumptions!D8",
                    editable=True,
                ),
            },
        }

        # Operating assumptions: staffing and delivery capacity.
        self.operating_assumptions = {
            "consulting_fte_start": InputField(
                value=60,
                description="Starting consulting FTEs",
                excel_ref="00_Inputs_Assumptions!B11",
                editable=True,
            ),
            "consulting_fte_growth_pct": InputField(
                value=0.00,
                description="Annual growth rate for consulting FTEs",
                excel_ref="00_Inputs_Assumptions!B12",
                editable=True,
            ),
            "work_days_per_year": InputField(
                value=220,
                description="Working days per year",
                excel_ref="00_Inputs_Assumptions!B13",
                editable=True,
            ),
            "training_internal_days": InputField(
                value=15,
                description="Training and internal days per year",
                excel_ref="00_Inputs_Assumptions!B14",
                editable=True,
            ),
            "sick_days": InputField(
                value=10,
                description="Sick days per year",
                excel_ref="00_Inputs_Assumptions!B15",
                editable=True,
            ),
            "unpaid_vacation_days": InputField(
                value=10,
                description="Unpaid vacation days per year",
                excel_ref="00_Inputs_Assumptions!B16",
                editable=True,
            ),
            "billable_days_per_year": InputField(
                value=185,
                description=(
                    "Calculated in Excel: Work days minus Training/Internal minus Sick"
                    " minus Unpaid vacation"
                ),
                excel_ref="00_Inputs_Assumptions!B17",
                editable=False,
            ),
            "day_rate_growth_pct": InputField(
                value=0.00,
                description="Annual growth rate for day rate",
                excel_ref="00_Inputs_Assumptions!B18",
                editable=True,
            ),
            "revenue_guarantee_pct_year_1": InputField(
                value=0.00,
                description="Revenue guarantee as percent of capacity for Year 1",
                excel_ref="N/A",
                editable=True,
            ),
            "revenue_guarantee_pct_year_2": InputField(
                value=0.00,
                description="Revenue guarantee as percent of capacity for Year 2",
                excel_ref="N/A",
                editable=True,
            ),
            "revenue_guarantee_pct_year_3": InputField(
                value=0.00,
                description="Revenue guarantee as percent of capacity for Year 3",
                excel_ref="N/A",
                editable=True,
            ),
            "new_hire_ramp_up_factor_fy1": InputField(
                value=0.08,
                description="Ramp-up factor for new hires in FY1",
                excel_ref="00_Inputs_Assumptions!B19",
                editable=True,
            ),
            "backoffice_fte_start": InputField(
                value=10,
                description="Starting backoffice FTEs",
                excel_ref="00_Inputs_Assumptions!B20",
                editable=True,
            ),
            "backoffice_fte_growth_pct": InputField(
                value=0.00,
                description="Annual growth rate for backoffice FTEs",
                excel_ref="00_Inputs_Assumptions!B21",
                editable=True,
            ),
            "avg_backoffice_salary_eur_per_year": InputField(
                value=None,
                description=(
                    "Excel cell empty; must be set for full cost model TODO"
                ),
                excel_ref="00_Inputs_Assumptions!B22",
                editable=True,
            ),
        }

        # Personnel cost assumptions: compensation and payroll factors.
        self.personnel_cost_assumptions = {
            "avg_consultant_base_cost_eur_per_year": InputField(
                value=214285.7142857143,
                description=(
                    "Calculated in Excel (formula present); treat as fixed default in v1"
                ),
                excel_ref="00_Inputs_Assumptions!B25",
                editable=False,
            ),
            "bonus_pct_of_base": InputField(
                value=0.00,
                description="Bonus as percent of base compensation",
                excel_ref="00_Inputs_Assumptions!B26",
                editable=True,
            ),
            "payroll_burden_pct_of_comp": InputField(
                value=0.00,
                description="Payroll burden as percent of compensation",
                excel_ref="00_Inputs_Assumptions!B27",
                editable=True,
            ),
            "wage_inflation_pct": InputField(
                value=0.02,
                description="Annual wage inflation rate",
                excel_ref="00_Inputs_Assumptions!B28",
                editable=True,
            ),
        }

        # Overhead and variable costs: fixed overhead plus revenue-linked costs.
        self.overhead_and_variable_costs = {
            "rent_eur_per_year": InputField(
                value=400000,
                description="Annual rent costs (EUR)",
                excel_ref="00_Inputs_Assumptions!B31",
                editable=True,
            ),
            "it_and_software_eur_per_year": InputField(
                value=300000,
                description="Annual IT and software costs (EUR)",
                excel_ref="00_Inputs_Assumptions!B32",
                editable=True,
            ),
            "overhead_inflation_pct": InputField(
                value=0.02,
                description="Annual overhead inflation rate",
                excel_ref="00_Inputs_Assumptions!B33",
                editable=True,
            ),
            "insurance_eur_per_year": InputField(
                value=50000,
                description="Annual insurance costs (EUR)",
                excel_ref="00_Inputs_Assumptions!B34",
                editable=True,
            ),
            "legal_audit_eur_per_year": InputField(
                value=150000,
                description="Annual legal and audit costs (EUR)",
                excel_ref="00_Inputs_Assumptions!B35",
                editable=True,
            ),
            "other_overhead_eur_per_year": InputField(
                value=50000,
                description="Other annual overhead costs (EUR)",
                excel_ref="00_Inputs_Assumptions!B36",
                editable=True,
            ),
            "travel_pct_of_revenue": InputField(
                value=0.05,
                description="Travel costs as percent of revenue",
                excel_ref="00_Inputs_Assumptions!B37",
                editable=True,
            ),
            "recruiting_pct_of_revenue": InputField(
                value=0.03,
                description="Recruiting costs as percent of revenue",
                excel_ref="00_Inputs_Assumptions!B38",
                editable=True,
            ),
            "training_pct_of_revenue": InputField(
                value=0.01,
                description="Training costs as percent of revenue",
                excel_ref="00_Inputs_Assumptions!B39",
                editable=True,
            ),
            "marketing_pct_of_revenue": InputField(
                value=0.01,
                description="Marketing costs as percent of revenue",
                excel_ref="00_Inputs_Assumptions!B40",
                editable=True,
            ),
        }

        # Capex and working capital: capital spend, depreciation, and liquidity needs.
        self.capex_and_working_capital = {
            "depreciation_eur_per_year": InputField(
                value=150000,
                description="Annual depreciation (EUR)",
                excel_ref="00_Inputs_Assumptions!B43",
                editable=True,
            ),
            "capex_eur_per_year": InputField(
                value=200000,
                description="Annual capital expenditures (EUR)",
                excel_ref="00_Inputs_Assumptions!B44",
                editable=True,
            ),
            "dso_days": InputField(
                value=60,
                description="Days sales outstanding",
                excel_ref="00_Inputs_Assumptions!B45",
                editable=True,
            ),
            "minimum_cash_balance_eur": InputField(
                value=250000,
                description="Minimum cash balance (EUR)",
                excel_ref="00_Inputs_Assumptions!B46",
                editable=True,
            ),
        }

        # Transaction and financing: deal structure and debt terms.
        self.transaction_and_financing = {
            "purchase_price_eur": InputField(
                value=16000000,
                description="Purchase price (EUR)",
                excel_ref="00_Inputs_Assumptions!B49",
                editable=True,
            ),
            "equity_contribution_eur": InputField(
                value=2000000,
                description="Equity contribution (EUR)",
                excel_ref="00_Inputs_Assumptions!B50",
                editable=True,
            ),
            "senior_term_loan_start_eur": InputField(
                value=11000000,
                description="Senior term loan opening balance (EUR)",
                excel_ref="00_Inputs_Assumptions!B51",
                editable=True,
            ),
            "senior_interest_rate_pct": InputField(
                value=0.06,
                description="Senior term loan interest rate",
                excel_ref="00_Inputs_Assumptions!B52",
                editable=True,
            ),
            "senior_repayment_per_year_eur": InputField(
                value=1000000,
                description="Annual senior debt repayment (EUR)",
                excel_ref="00_Inputs_Assumptions!B53",
                editable=True,
            ),
            "revolver_limit_eur": InputField(
                value=1500000,
                description="Revolver credit limit (EUR)",
                excel_ref="00_Inputs_Assumptions!B54",
                editable=True,
            ),
            "revolver_interest_rate_pct": InputField(
                value=0.07,
                description="Revolver interest rate",
                excel_ref="00_Inputs_Assumptions!B55",
                editable=True,
            ),
            "special_repayment_amount_eur": InputField(
                value=0,
                description="Optional; set to 0 in v1",
                excel_ref="00_Inputs_Assumptions!B56",
                editable=True,
            ),
            "special_repayment_year": InputField(
                value=None,
                description="Optional; None in v1",
                excel_ref="00_Inputs_Assumptions!B57",
                editable=True,
            ),
        }

        # Tax and distributions: tax rate and shareholder payouts.
        self.tax_and_distributions = {
            "tax_rate_pct": InputField(
                value=0.30,
                description="Corporate tax rate",
                excel_ref="00_Inputs_Assumptions!B58",
                editable=True,
            ),
            "dividend_payout_ratio_pct": InputField(
                value=0.00,
                description="Dividend payout ratio",
                excel_ref="00_Inputs_Assumptions!B59",
                editable=True,
            ),
            "dividends_allowed_starting_fy": InputField(
                value=4,
                description="First fiscal year when dividends are allowed",
                excel_ref="00_Inputs_Assumptions!B60",
                editable=True,
            ),
        }

        # Valuation assumptions: buyer vs seller inputs for multiple and DCF views.
        self.valuation_assumptions = {
            "general_valuation_context": {
                "valuation_reference_metric": InputField(
                    value="EBIT",
                    description=(
                        "Metric used as valuation basis (e.g. EBIT or EBITDA)"
                    ),
                    excel_ref="N/A",
                    editable=True,
                )
            },
            # Multiple-based valuation: buyer is conservative; seller is optimistic.
            "multiple_valuation": {
                "buyer_multiple": InputField(
                    value=None,
                    description=(
                        "Multiple assumed from buyer perspective"
                        " (downside / conservative)"
                    ),
                    excel_ref="N/A",
                    editable=True,
                ),
                "seller_multiple": InputField(
                    value=None,
                    description=(
                        "Multiple assumed from seller perspective"
                        " (upside / true value view)"
                    ),
                    excel_ref="N/A",
                    editable=True,
                ),
            },
            # DCF-based valuation: buyer uses downward adjustments; seller uses upward.
            "dcf_valuation": {
                "discount_rate_wacc": InputField(
                    value=None,
                    description="Weighted Average Cost of Capital used for DCF",
                    excel_ref="N/A",
                    editable=True,
                ),
                "explicit_forecast_years": InputField(
                    value=None,
                    description="Number of explicit forecast years for DCF",
                    excel_ref="N/A",
                    editable=True,
                ),
                "terminal_growth_rate": InputField(
                    value=None,
                    description=(
                        "Long-term growth rate after explicit forecast period"
                    ),
                    excel_ref="N/A",
                    editable=True,
                ),
                "buyer_adjustment_factor": InputField(
                    value=None,
                    description=(
                        "Optional downward adjustment from buyer perspective"
                    ),
                    excel_ref="N/A",
                    editable=True,
                ),
                "seller_adjustment_factor": InputField(
                    value=None,
                    description=(
                        "Optional upward adjustment from seller perspective"
                    ),
                    excel_ref="N/A",
                    editable=True,
                ),
            },
            "valuation_display_mode": InputField(
                value=None,
                description=(
                    "Defines how valuation should be presented"
                    " (e.g. range, midpoint, comparison)"
                ),
                excel_ref="N/A",
                editable=True,
            ),
        }


def create_demo_input_model():
    """
    Create an InputModel populated with reasonable demo values.
    """
    input_model = InputModel()

    # Fill previously empty or optional fields with demo values.
    input_model.operating_assumptions[
        "avg_backoffice_salary_eur_per_year"
    ].value = 80000
    input_model.operating_assumptions["consulting_fte_start"].value = 63
    input_model.operating_assumptions["consulting_fte_growth_pct"].value = 0.0
    input_model.operating_assumptions["work_days_per_year"].value = 220
    input_model.operating_assumptions["day_rate_growth_pct"].value = 0.0
    input_model.operating_assumptions["backoffice_fte_start"].value = 18
    input_model.operating_assumptions["backoffice_fte_growth_pct"].value = 0.0

    input_model.scenario_parameters["utilization_rate"]["base"].value = 0.68
    input_model.scenario_parameters["utilization_rate"]["best"].value = 0.68
    input_model.scenario_parameters["utilization_rate"]["worst"].value = 0.68
    input_model.scenario_parameters["day_rate_eur"]["base"].value = 2125
    input_model.scenario_parameters["day_rate_eur"]["best"].value = 2125
    input_model.scenario_parameters["day_rate_eur"]["worst"].value = 2125
    input_model.operating_assumptions[
        "revenue_guarantee_pct_year_1"
    ].value = 0.80
    input_model.operating_assumptions[
        "revenue_guarantee_pct_year_2"
    ].value = 0.60
    input_model.operating_assumptions[
        "revenue_guarantee_pct_year_3"
    ].value = 0.60

    input_model.personnel_cost_assumptions[
        "avg_consultant_base_cost_eur_per_year"
    ].value = 150000
    input_model.personnel_cost_assumptions["bonus_pct_of_base"].value = 0.0
    input_model.personnel_cost_assumptions[
        "payroll_burden_pct_of_comp"
    ].value = 0.0
    input_model.personnel_cost_assumptions["wage_inflation_pct"].value = 0.02

    input_model.overhead_and_variable_costs[
        "legal_audit_eur_per_year"
    ].value = 320000
    input_model.overhead_and_variable_costs[
        "it_and_software_eur_per_year"
    ].value = 440000
    input_model.overhead_and_variable_costs[
        "rent_eur_per_year"
    ].value = 1730000
    input_model.overhead_and_variable_costs[
        "insurance_eur_per_year"
    ].value = 400000
    input_model.overhead_and_variable_costs[
        "other_overhead_eur_per_year"
    ].value = 700000
    input_model.overhead_and_variable_costs[
        "travel_pct_of_revenue"
    ].value = 0.0
    input_model.overhead_and_variable_costs[
        "recruiting_pct_of_revenue"
    ].value = 0.0
    input_model.overhead_and_variable_costs[
        "training_pct_of_revenue"
    ].value = 0.0
    input_model.overhead_and_variable_costs[
        "marketing_pct_of_revenue"
    ].value = 0.0

    input_model.transaction_and_financing[
        "special_repayment_amount_eur"
    ].value = 0
    input_model.transaction_and_financing[
        "special_repayment_year"
    ].value = 3

    input_model.valuation_assumptions["multiple_valuation"][
        "buyer_multiple"
    ].value = 5.0
    input_model.valuation_assumptions["multiple_valuation"][
        "seller_multiple"
    ].value = 7.0

    input_model.valuation_assumptions["dcf_valuation"][
        "discount_rate_wacc"
    ].value = 0.10
    input_model.valuation_assumptions["dcf_valuation"][
        "explicit_forecast_years"
    ].value = 5
    input_model.valuation_assumptions["dcf_valuation"][
        "terminal_growth_rate"
    ].value = 0.02
    input_model.valuation_assumptions["dcf_valuation"][
        "buyer_adjustment_factor"
    ].value = 0.95
    input_model.valuation_assumptions["dcf_valuation"][
        "seller_adjustment_factor"
    ].value = 1.05

    input_model.valuation_assumptions["valuation_display_mode"].value = "range"

    return input_model
