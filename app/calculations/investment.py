def calculate_investment(
    input_model, cashflow_result, pnl_result=None, balance_sheet=None
):
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

    enterprise_value = final_year_ebit * exit_multiple
    net_debt_exit = 0.0
    excess_cash = 0.0
    if balance_sheet:
        last_year = balance_sheet[-1]
        net_debt_exit = last_year.get("financial_debt", 0.0)
        excess_cash = last_year.get("cash", 0.0)
    exit_value = enterprise_value - net_debt_exit + excess_cash

    # Build equity cashflows: initial outflow, dividends (if any), exit proceeds.
    equity_cashflows = [-equity_amount]
    for i in range(len(cashflow_result)):
        dividend = 0.0
        if i == len(cashflow_result) - 1:
            equity_cashflows.append(dividend + exit_value)
        else:
            equity_cashflows.append(dividend)

    irr = _calculate_irr(equity_cashflows)

    return {
        "initial_equity": equity_amount,
        "equity_cashflows": equity_cashflows,
        "exit_value": exit_value,
        "enterprise_value": enterprise_value,
        "net_debt_exit": net_debt_exit,
        "excess_cash_exit": excess_cash,
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
