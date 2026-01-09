import streamlit as st

from streamlit_app import run_app


def main():
    st.set_page_config(page_title="MBO Financial Model", layout="wide")
    st.session_state.setdefault("page_key", "Operating Model (P&L)")

    with st.sidebar:
        st.markdown("MBO Financial Model")

        operating_options = [
            "Operating Model (P&L)",
            "Cashflow & Liquidity",
            "Balance Sheet",
        ]
        planning_options = [
            "Revenue Model",
            "Cost Model",
            "Other Assumptions",
        ]
        financing_options = [
            "Financing & Debt",
            "Equity Case",
        ]
        valuation_options = ["Valuation & Purchase Price"]
        settings_options = ["Model Settings"]

        st.markdown("### OPERATING MODEL")
        operating_index = (
            operating_options.index(st.session_state["page_key"])
            if st.session_state["page_key"] in operating_options
            else 0
        )
        selected_operating = st.sidebar.selectbox(
            "",
            operating_options,
            index=operating_index,
            key="nav_operating",
        )
        if selected_operating in operating_options:
            st.session_state["page_key"] = selected_operating

        st.markdown("### PLANNING")
        planning_index = (
            planning_options.index(st.session_state["page_key"])
            if st.session_state["page_key"] in planning_options
            else 0
        )
        selected_planning = st.sidebar.selectbox(
            "",
            planning_options,
            index=planning_index,
            key="nav_planning",
        )
        if selected_planning in planning_options:
            st.session_state["page_key"] = selected_planning

        st.markdown("### FINANCING")
        financing_index = (
            financing_options.index(st.session_state["page_key"])
            if st.session_state["page_key"] in financing_options
            else 0
        )
        selected_financing = st.sidebar.selectbox(
            "",
            financing_options,
            index=financing_index,
            key="nav_financing",
        )
        if selected_financing in financing_options:
            st.session_state["page_key"] = selected_financing

        st.markdown("### VALUATION")
        valuation_index = (
            valuation_options.index(st.session_state["page_key"])
            if st.session_state["page_key"] in valuation_options
            else 0
        )
        selected_valuation = st.sidebar.selectbox(
            "",
            valuation_options,
            index=valuation_index,
            key="nav_valuation",
        )
        if selected_valuation in valuation_options:
            st.session_state["page_key"] = selected_valuation

        st.markdown("### SETTINGS")
        settings_index = (
            settings_options.index(st.session_state["page_key"])
            if st.session_state["page_key"] in settings_options
            else 0
        )
        selected_settings = st.sidebar.selectbox(
            "",
            settings_options,
            index=settings_index,
            key="nav_settings",
        )
        if selected_settings in settings_options:
            st.session_state["page_key"] = selected_settings

    run_app(st.session_state["page_key"])


if __name__ == "__main__":
    main()
