import streamlit as st

from streamlit_app import run_app


def main():
    st.set_page_config(page_title="MBO Financial Model", layout="wide")
    st.session_state.setdefault("page_key", "Operating Model (P&L)")

    sections = [
        (
            "OPERATING MODEL",
            "nav_operating",
            [
                "Operating Model (P&L)",
                "Cashflow & Liquidity",
                "Balance Sheet",
            ],
        ),
        (
            "PLANNING",
            "nav_planning",
            [
                "Revenue Model",
                "Cost Model",
                "Other Assumptions",
            ],
        ),
        (
            "FINANCING",
            "nav_financing",
            [
                "Financing & Debt",
                "Equity Case",
            ],
        ),
        ("VALUATION", "nav_valuation", ["Valuation & Purchase Price"]),
        ("SETTINGS", "nav_settings", ["Model Settings"]),
    ]

    current_page = st.session_state["page_key"]
    for _, key, options in sections:
        if current_page not in options:
            st.session_state.pop(key, None)

    with st.sidebar:
        st.markdown(
            """
            <style>
              [data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child {
                display: none;
              }
              [data-testid="stSidebar"] div[role="radiogroup"] > label {
                padding: 0.2rem 0.45rem;
                border-radius: 0.4rem;
              }
              [data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) {
                background-color: #e5e7eb;
                font-weight: 600;
              }
            </style>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("MBO Financial Model")
        for section_title, key, options in sections:
            st.markdown(f"### {section_title}")
            index = options.index(current_page) if current_page in options else None
            selection = st.radio(
                section_title,
                options,
                index=index,
                key=key,
                label_visibility="collapsed",
            )
            if selection is not None and selection != current_page:
                st.session_state["page_key"] = selection
                current_page = selection

    run_app(st.session_state["page_key"])


if __name__ == "__main__":
    main()
