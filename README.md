# 🧾 EIB Assistant – RSU Tax Formatter

A Streamlit app that reformats RSU payroll output data into the exact column layout required by the **Workday Payroll Off-cycle Payment EIB** (Enterprise Interface Builder).

---

## Features

- 📂 Upload your output spreadsheet (Original / Deloitte sheet)
- 🔍 Filter by Employee ID and Transaction ID
- ⚙️ Configure Batch ID, Payment Date, Run Category, and other EIB fields
- 🎨 Colour-coded preview table (TVRSU, NVRSU, W_SWW, sub-calc lines)
- ⬇️ Download EIB-formatted Excel file (134 columns, correct layout)
- ⚡ Bulk mode — generate all employees at once
- 🗺️ Auto-converts state abbreviations (e.g. `NY`) to Payroll Authority Tax Codes (e.g. `36`)
- 🚫 Skips zero-value state rows automatically

---

## Getting Started

### Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open [http://localhost:8501](http://localhost:8501) in your browser.

### Dev Container (VS Code)

1. Open this repo in VS Code
2. Click **Reopen in Container** when prompted
3. Run `streamlit run app.py` in the terminal

---

## EIB Output Format

The exported Excel file matches the **Payroll Off-cycle Payment** EIB template with 134 columns. Key populated columns:

| Column | Field | Value |
|--------|-------|-------|
| B | Spreadsheet Key* | Transaction ID |
| D | Batch ID | Configurable |
| E | Payment ID* | `="Trailing2State"&D&F` formula |
| F | Employee* | Employee ID |
| G | Payment Date* | Configurable |
| H | Period Date* | Configurable |
| I | Payment Priority* | 3 |
| J | Run Category | RSU_USA |
| L | Result Type* | OnDemandPayment |
| N | Reason* | Stock |
| AN | Row ID* | 1, 2, 3… |
| AO | Earning* | TVRSU / NVRSU |
| AP | Deduction* | W_SWW |
| AR | Amount | From source |
| BF | State Authority | Payroll Tax Code |
| CD | Sub Row ID | 1 / 2 / 3 |
| CE | Related Calculation | W_CWCGW / W_TTWG / W_TXWG |
| CF | Value | Income for Region |

---

## State Code Mapping

State abbreviations are automatically converted to Payroll Authority Tax Codes per the official mapping (e.g. `CT` → `09`, `NY` → `36`, `CA` → `06`). Leading zeros are always preserved.

---

## File Structure

```
.
├── .devcontainer/
│   └── devcontainer.json   # VS Code Dev Container config
├── app.py                  # Main Streamlit application
├── index.html              # Optional landing page
├── logo.png                # App logo (replace with your own)
├── requirements.txt        # Python dependencies
├── ADD_LOGO_HERE.txt       # Logo replacement instructions
└── README.md               # This file
```

---

## Requirements

```
streamlit>=1.32.0
openpyxl>=3.1.0
pandas>=2.0.0
xlsxwriter>=3.1.0
```
