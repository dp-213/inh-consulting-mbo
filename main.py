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
        def _on_nav_change(key_name):
            st.session_state["page_key"] = st.session_state.get(key_name)
            for _, other_key, _ in sections:
                if other_key != key_name:
                    st.session_state.pop(other_key, None)

        for section_title, key, options in sections:
            st.markdown(f"### {section_title}")
            index = options.index(current_page) if current_page in options else None
            st.radio(
                section_title,
                options,
                index=index,
                key=key,
                label_visibility="collapsed",
                on_change=_on_nav_change,
                args=(key,),
            )

    run_app(st.session_state["page_key"])


if __name__ == "__main__":
    main()
