import streamlit as st
from streamlit_app import run_app


def main():
    page = st.session_state.get("page", "Operating Model (P&L)")
    run_app(page)


if __name__ == "__main__":
    main()
