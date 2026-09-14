# 🧾 EIB Assistant – RSU Tax Formatter

Formats RSU payroll output data into the exact column layout required by the **Workday Payroll Off-cycle Payment EIB**.

Available in two formats:
- **`index.html`** — Standalone offline app, open directly in any browser
- **`app.py`** — Streamlit web app, deployable to Streamlit Community Cloud

---

## Features

- 📂 Upload output spreadsheet (Original / Deloitte sheet) via drag & drop
- 🔍 Filter by Employee ID and Transaction ID
- ⚙️ Configure Batch ID, Payment Date, Run Category, and other EIB fields
- 🎨 Colour-coded preview table (TVRSU, NVRSU, W_SWW, sub-calc lines)
- 📋 Copy rows to clipboard — paste directly into EIB (134 columns)
- ⬇️ Export EIB-formatted Excel file (134 columns, correct layout)
- ⚡ Bulk mode — generate all employees at once
- 🗺️ Auto-converts state abbreviations (e.g. `NY`) to Payroll Authority Tax Codes (e.g. `36`)
- 🚫 Skips zero-value state rows automatically

---

## Option 1: HTML (No Install)

1. Download or clone this repo
2. Open `index.html` in Chrome or Edge
3. Done — fully offline, no server needed

---

## Option 2: Streamlit App

### Run Locally
```bash
pip install -r requirements.txt
streamlit run app.py
```
Open [http://localhost:8501](http://localhost:8501)

### Deploy to Streamlit Community Cloud
1. Push this repo to **GitHub** (Streamlit Cloud requires GitHub)
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**
3. Select your repo, branch `main`, main file `app.py`
4. Click **Deploy**

### Dev Container (VS Code)
1. Open repo in VS Code
2. Click **Reopen in Container** when prompted
3. Run `streamlit run app.py` in the terminal

---

## EIB Output Format

Exported Excel matches the **Payroll Off-cycle Payment** EIB template (134 columns).

| Column | Field | Value |
|--------|-------|-------|
| B | Spreadsheet Key* | Transaction ID |
| D | Batch ID | Configurable |
| E | Payment ID* | `="Trailing2State"&D&F` |
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

State abbreviations auto-convert to Payroll Authority Tax Codes
(e.g. `CT` → `09`, `NY` → `36`, `CA` → `06`). Leading zeros preserved.

---

## File Structure

```
.
├── .devcontainer/
│   └── devcontainer.json   # VS Code Dev Container config
├── app.py                  # Streamlit app
├── index.html              # Standalone offline HTML app
├── logo.png                # App logo (replace with your own)
├── requirements.txt        # Python dependencies
├── ADD_LOGO_HERE.txt       # Logo instructions
└── README.md               # This file
```

---

## Requirements (Streamlit only)

```
streamlit>=1.32.0
openpyxl>=3.1.0
pandas>=2.0.0
xlsxwriter>=3.1.0
```
