# 🧾 EIB Assistant – RSU Tax Formatter

A fully self-contained, offline HTML tool that reformats RSU payroll output data into the exact column layout required by the **Workday Payroll Off-cycle Payment EIB** (Enterprise Interface Builder).

No server, no install, no internet required. Just open `index.html` in any browser.

---

## Features

- 📂 Upload your output spreadsheet (Original / Deloitte sheet) via drag & drop or browse
- 🔍 Filter by Employee ID and Transaction ID
- ⚙️ Configure Batch ID, Payment Date, Run Category, and other EIB fields
- 🎨 Colour-coded preview table (TVRSU, NVRSU, W_SWW, sub-calc lines)
- 📋 Copy rows to clipboard (134 columns, paste directly into EIB)
- ⬇️ Export EIB-formatted Excel file (134 columns, correct layout)
- ⚡ Bulk mode — generate all employees at once
- 🗺️ Auto-converts state abbreviations (e.g. `NY`) to Payroll Authority Tax Codes (e.g. `36`)
- 🚫 Skips zero-value state rows automatically
- 📦 Fully offline — XLSX library bundled inside the HTML file

---

## How to Use

1. Open `index.html` in Chrome or Edge
2. **Step 1** — Upload your output spreadsheet
3. **Step 2** — Select Employee ID and Transaction ID
4. **Step 3** — Confirm or update EIB settings (Batch ID, dates, etc.)
5. **Step 4** — Preview the EIB rows and click **Export to Excel** or **Copy Rows**
6. Paste or import directly into your EIB template

---

## EIB Output Format

The exported Excel file matches the **Payroll Off-cycle Payment** EIB template with 134 columns.
Key populated columns:

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

State abbreviations are automatically converted to Payroll Authority Tax Codes
(e.g. `CT` → `09`, `NY` → `36`, `CA` → `06`). Leading zeros are always preserved.

---

## File Structure

```
.
├── .devcontainer/
│   └── devcontainer.json   # VS Code Dev Container config
├── index.html              # Complete offline app (HTML + CSS + JS bundled)
├── logo.png                # App logo (replace with your own)
├── ADD_LOGO_HERE.txt       # Logo replacement instructions
└── README.md               # This file
```

---

## Sharing

Since the entire app is a single HTML file, you can share it by:
- Sending `index.html` directly via email or Teams
- Hosting on any web server or SharePoint
- Pushing to GitLab Pages for a permanent team URL

### GitLab Pages
Push this repo to GitLab and enable **Pages** in Settings → Pages.
The app will be available at `https://your-group.gitlab.io/your-repo/`.
