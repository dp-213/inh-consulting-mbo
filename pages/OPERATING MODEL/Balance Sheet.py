import streamlit as st

from app.streamlit_app import run_app

st.set_page_config(page_title="Balance Sheet", layout="wide")
run_app("Balance Sheet")
