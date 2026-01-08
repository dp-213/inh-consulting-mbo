def calculate_investment(input_model, cashflow_result, pnl_result=None):
    """
    Calculate basic equity investment performance metrics.
    Returns a dictionary with initial equity, cashflows, exit value, and IRR.
    """
    # Map legacy financing fields to Excel-equivalent transaction inputs.
    equity_amount = input_model.transaction_and_financing[
        "equity_contribution_eur"
    ].value
    exit_multiple = input_model.valuation_assumptions["multiple_valuation"][
        "seller_multiple"
    ].value
    exit_multiple = 0 if exit_multiple is None else exit_multiple

    # Estimate exit value from the final year EBIT and exit multiple.
    final_year_ebit = 0
    if pnl_result:
        final_year_ebit = pnl_result[-1].get("ebit", 0)

    exit_value = final_year_ebit * exit_multiple

    # Build equity cashflows: initial outflow, then yearly net cashflows.
    equity_cashflows = [-equity_amount]
    for i, year_data in enumerate(cashflow_result):
        cashflow = year_data["net_cashflow"]

        # Add exit value in the final year.
        if i == len(cashflow_result) - 1:
            cashflow += exit_value

        equity_cashflows.append(cashflow)

    irr = _calculate_irr(equity_cashflows)

    return {
        "initial_equity": equity_amount,
        "equity_cashflows": equity_cashflows,
        "exit_value": exit_value,
        "irr": irr,
    }


def _calculate_irr(cashflows, max_iterations=100):
    """
    Estimate IRR using a simple bisection method.
    Returns 0.0 if no sign change is found.
    """
    def npv(rate):
        return sum(cf / ((1 + rate) ** i) for i, cf in enumerate(cashflows))

    low = -0.9
    high = 1.0
    npv_low = npv(low)
    npv_high = npv(high)

    # Expand the high rate until we find a sign change or hit a limit.
    while npv_low * npv_high > 0 and high < 10:
        high *= 2
        npv_high = npv(high)

    if npv_low * npv_high > 0:
        return 0.0

    for _ in range(max_iterations):
        mid = (low + high) / 2
        npv_mid = npv(mid)

        if abs(npv_mid) < 1e-8:
            return mid

        if npv_low * npv_mid < 0:
            high = mid
            npv_high = npv_mid
        else:
            low = mid
            npv_low = npv_mid

    return (low + high) / 2
