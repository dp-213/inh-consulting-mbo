Integrated MBO Financial Model

What this is
An integrated Management Buyout (MBO) financial model implemented in Python. It produces a connected P&L, cashflow, debt schedule, balance sheet, and equity performance view from a single input model.

What it replaces
A spreadsheet-driven (Excel) model. This codebase mirrors the input sheet structure and provides a reproducible, version-controlled alternative for scenario-based analysis.

Architecture overview
- app/data_model.py: central InputModel that mirrors the Excel input sheet and stores all assumptions with metadata.
- app/calculations/: pure calculation modules (P&L, cashflow, debt, balance sheet, investment).
- app/run_model.py: orchestration layer that runs the full model in the correct order.
- app/streamlit_app.py: minimal UI to interact with scenarios and view results as tables.

Run the model locally
python -c "from app.run_model import run_model; print(run_model())"

Start the Streamlit app
streamlit run app/streamlit_app.py

Not included yet
- Deployment or hosting
- Authentication or user management
- Exit scenarios or advanced valuation logic
