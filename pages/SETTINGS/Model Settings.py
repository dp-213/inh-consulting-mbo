import streamlit as st

from app.streamlit_app import run_app

st.set_page_config(page_title="Model Settings", layout="wide")
run_app("Model Settings")
