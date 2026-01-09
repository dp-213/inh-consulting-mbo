import streamlit as st

from app.streamlit_app import run_app

st.set_page_config(page_title="Financing & Debt", layout="wide")
run_app("Financing & Debt")
