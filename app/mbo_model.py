class MBOModel:
    def __init__(
        self,
        revenue,
        ebit,
        purchase_multiple,
        debt_amount,
        interest_rate,
        tax_rate,
    ):
        # Store inputs so methods can use them for calculations.
        self.revenue = revenue
        self.ebit = ebit
        self.purchase_multiple = purchase_multiple
        self.debt_amount = debt_amount
        self.interest_rate = interest_rate
        self.tax_rate = tax_rate

    def enterprise_value(self):
        # Enterprise value is a multiple of EBIT.
        return self.purchase_multiple * self.ebit

    def equity_required(self):
        # Equity required is enterprise value minus debt.
        return self.enterprise_value() - self.debt_amount

    def annual_interest_cost(self):
        # Annual interest cost from debt.
        return self.debt_amount * self.interest_rate

    def net_income_estimate(self):
        # Start with EBIT and subtract interest.
        profit_after_interest = self.ebit - self.annual_interest_cost()
        # Taxes are applied to profit after interest.
        taxes = profit_after_interest * self.tax_rate
        return profit_after_interest - taxes
