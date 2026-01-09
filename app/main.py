import streamlit as st

from streamlit_app import run_app


def main():
    st.set_page_config(page_title="MBO Financial Model", layout="wide")
    st.session_state.setdefault("page_key", "Operating Model (P&L)")

    with st.sidebar:
        st.markdown("MBO Financial Model")

        st.markdown("### OPERATING MODEL")
        if st.sidebar.button("Operating Model (P&L)", use_container_width=True):
            st.session_state["page_key"] = "Operating Model (P&L)"
        if st.sidebar.button("Cashflow & Liquidity", use_container_width=True):
            st.session_state["page_key"] = "Cashflow & Liquidity"
        if st.sidebar.button("Balance Sheet", use_container_width=True):
            st.session_state["page_key"] = "Balance Sheet"

        st.markdown("### PLANNING")
        if st.sidebar.button("Revenue Model", use_container_width=True):
            st.session_state["page_key"] = "Revenue Model"
        if st.sidebar.button("Cost Model", use_container_width=True):
            st.session_state["page_key"] = "Cost Model"
        if st.sidebar.button("Other Assumptions", use_container_width=True):
            st.session_state["page_key"] = "Other Assumptions"

        st.markdown("### FINANCING")
        if st.sidebar.button("Financing & Debt", use_container_width=True):
            st.session_state["page_key"] = "Financing & Debt"
        if st.sidebar.button("Equity Case", use_container_width=True):
            st.session_state["page_key"] = "Equity Case"

        st.markdown("### VALUATION")
        if st.sidebar.button("Valuation & Purchase Price", use_container_width=True):
            st.session_state["page_key"] = "Valuation & Purchase Price"

        st.markdown("### SETTINGS")
        if st.sidebar.button("Model Settings", use_container_width=True):
            st.session_state["page_key"] = "Model Settings"

    run_app(st.session_state["page_key"])


if __name__ == "__main__":
    main()
