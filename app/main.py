import streamlit as st

from streamlit_app import run_app

pages = {
    "OPERATING MODEL": [
        st.Page(lambda: run_app("Operating Model (P&L)"), title="Operating Model (P&L)"),
        st.Page(lambda: run_app("Cashflow & Liquidity"), title="Cashflow & Liquidity"),
        st.Page(lambda: run_app("Balance Sheet"), title="Balance Sheet"),
    ],
    "PLANNING": [
        st.Page(lambda: run_app("Revenue Model"), title="Revenue Model"),
        st.Page(lambda: run_app("Cost Model"), title="Cost Model"),
        st.Page(lambda: run_app("Other Assumptions"), title="Other Assumptions"),
    ],
    "FINANCING": [
        st.Page(lambda: run_app("Financing & Debt"), title="Financing & Debt"),
        st.Page(lambda: run_app("Equity Case"), title="Equity Case"),
    ],
    "VALUATION": [
        st.Page(lambda: run_app("Valuation & Purchase Price"), title="Valuation & Purchase Price"),
    ],
    "SETTINGS": [
        st.Page(lambda: run_app("Model Settings"), title="Model Settings"),
    ],
}

nav = st.navigation(pages)
nav.run()
