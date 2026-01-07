class InputModel:
    def __init__(self):
        # General business identifiers and scope settings.
        self.general = {
            "company_name": "",
            "business_unit": "",
            "industry": "",
            "currency": "",
            "start_year": 0,
            "projection_years": 0,
        }

        # Transaction structure and deal-level terms.
        self.transaction = {
            "purchase_price": 0.0,
            "purchase_multiple": 0.0,
            "deal_fees": 0.0,
            "closing_date": "",
            "seller": "",
        }

        # Operating performance drivers and cost structure.
        self.operations = {
            "revenue": 0.0,
            "ebit": 0.0,
            "ebit_margin": 0.0,
            "operating_expenses": 0.0,
            "capex": 0.0,
            "working_capital_change": 0.0,
        }

        # Financing sources, terms, and capital structure.
        self.financing = {
            "debt_amount": 0.0,
            "interest_rate": 0.0,
            "equity_amount": 0.0,
            "debt_term_years": 0,
            "amortization_rate": 0.0,
        }

        # Cross-cutting modeling assumptions and policy inputs.
        self.assumptions = {
            "tax_rate": 0.0,
            "inflation_rate": 0.0,
            "growth_rate": 0.0,
            "exit_multiple": 0.0,
            "discount_rate": 0.0,
        }
