# 🧾 EIB Assistant – RSU Tax Formatter

A Streamlit web app that reads your RSU output spreadsheet and automatically generates rows in the exact format required by the **Payroll Off cycle Payment EIB** (134-column Workday layout).

## Features
- Upload your RSU output Excel file
- Filter by Employee ID and/or Transaction ID
- Configurable EIB settings (Batch ID, dates, Run Category, etc.)
- Auto-maps state abbreviations → Payroll Authority Tax Codes
- Generates TVRSU, NVRSU, W_SWW + sub-calc lines per state
- Skips zero-value state rows automatically
- Single-employee or bulk export
- Download output as Excel ready to paste into your EIB

## Deploy on Streamlit Cloud
1. Fork or upload this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo
4. Set `app.py` as the main file
5. Deploy!
