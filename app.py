import streamlit as st
import pandas as pd
import openpyxl
import io
from copy import deepcopy

# ── Page config ──────────────────────────────────────────────
st.set_page_config(
    page_title="EIB Assistant – RSU Tax Formatter",
    page_icon="🧾",
    layout="wide",
)

# ── Custom CSS ────────────────────────────────────────────────
st.markdown("""
<style>
  .main .block-container { padding-top: 1.5rem; }
  .stDataFrame { font-size: 0.82rem; }
  div[data-testid="stMetric"] { background:#f0f4ff; border-radius:8px; padding:10px 16px; }
  .row-tvrsu { background-color: #e6f4ea; }
  .row-nvrsu { background-color: #e8f0fe; }
  .row-wsww  { background-color: #fce8b2; }
  .row-sub   { background-color: #fdf0e0; }
</style>
""", unsafe_allow_html=True)

# ── State abbreviation → Payroll Authority Tax Code ──────────
STATE_MAP = {
    'AL':'01','AK':'02','AZ':'04','AR':'05','CA':'06','CO':'08','CT':'09',
    'DE':'10','DC':'11','FL':'12','GA':'13','GU':'66','HI':'15','ID':'16',
    'IL':'17','IN':'18','IA':'19','KS':'20','KY':'21','LA':'22','ME':'23',
    'MD':'24','MA':'25','MI':'26','MN':'27','MS':'28','MO':'29','MT':'30',
    'NE':'31','NV':'32','NH':'33','NJ':'34','NM':'35','NY':'36','NC':'37',
    'ND':'38','OH':'39','OK':'40','OR':'41','PA':'42','PR':'72','RI':'44',
    'SC':'45','SD':'46','TN':'47','TX':'48','UT':'49','VT':'50','VA':'51',
    'VI':'78','WA':'53','WV':'54','WI':'55','WY':'56',
}

def resolve_state_code(raw):
    """Convert state abbreviation or numeric code to Payroll Authority Tax Code."""
    if raw is None or str(raw).strip() == '':
        return ''
    s = str(raw).strip().upper()
    if s in STATE_MAP:
        return STATE_MAP[s]
    try:
        return str(int(s)).zfill(2)
    except ValueError:
        return s

# ── EIB column layout (0-indexed, 134 total columns) ─────────
TOTAL_COLS = 134
C = dict(
    spreadsheetKey=1, batchId=3, paymentId=4, employee=5,
    paymentDate=6, periodDate=7, payPriority=8, runCategory=9,
    resultType=11, reason=13, rowId=39, earning=40, deduction=41,
    amount=43, stateAuthority=57, subRowId=85, relatedCalc=86, value=87,
)

EIB_HEADERS = [
    'Fields','Spreadsheet Key*','Payroll Off-cycle Payment','Batch ID','Payment ID*',
    'Employee*','Payment Date*','Period Date*','Payment Priority*','Run Category',
    'Pay Group','Result Type*','Replacement','Reason*','Use Supplemental Tax Rate',
    'Override Payment to Check','Take Additional Withholding',
    'Include Retro Differences in Payment','Load or Refresh Input','Row ID*',
    'Tax Frequency Value','Tax Frequency Period','Third Party Sick Pay','Net Amount',
    'Check Number','Bank Account','Company','Region','Location','Cost Center',
    'Job Profile','State  Work ','State  Resident ','County  Work ','County  Resident ',
    'City  Work ','City  Resident ','School District  Resident ',
    'Payroll Reference Number','Row ID*','Earning*','Deduction*','Position','Amount',
    'Hours','Rate','Adjustment','Reference Date','Currency','Location','Region',
    'Job Profile','Cost Center','Project','Project Phase','Project Task',
    'Withholding Order Case','State Authority','County Authority','City Authority',
    'School District Authority','Custom Worktag 01','Custom Worktag 02',
    'Custom Worktag 03','Custom Worktag 04','Custom Worktag 05','Fund','Grant',
    'Gift','Program','Business Unit','Object Class','Custom Organization+',
    'Custom Worktag 06','Custom Worktag 07','Custom Worktag 08','Custom Worktag 09',
    'Custom Worktag 10','Custom Worktag 11','Custom Worktag 12','Custom Worktag 13',
    'Custom Worktag 14','Custom Worktag 15','Local Other Tax Authority','NI Category',
    'Row ID*','Related Calculation','Value','Company',
]

COL_SEARCH = {
    'empId':        ['employee id'],
    'txnId':        ['transaction id'],
    'name':         ['name'],
    'stateCode':    ['payroll authority tax code','state code','region code'],
    'totalDist':    ['total distribution/exercise gross gain','reportable income'],
    'totalWith':    ['total employee withholding (actual & hypo taxes)','country total actual tax ee liability'],
    'regionTax':    ['region tax employee liability'],
    'incomeRegion': ['income for region tax pre gross up'],
}

def find_col(headers, keys):
    for key in keys:
        for i, h in enumerate(headers):
            if h and key in str(h).lower().replace('\n', ' '):
                return i
    return -1

def to_num(v):
    try:
        return float(v) if v not in (None, '') else None
    except (ValueError, TypeError):
        return None

# ── Parse sheet into structured records ──────────────────────
def parse_sheet(ws):
    rows = list(ws.iter_rows(values_only=True))
    header_row_idx = -1
    headers = []
    for i, row in enumerate(rows[:5]):
        joined = ' '.join(str(c).lower() for c in row if c).lower()
        if 'employee' in joined or 'transaction' in joined:
            header_row_idx = i
            headers = [str(c) if c else '' for c in row]
            break
    if header_row_idx < 0:
        return []

    cE = find_col(headers, COL_SEARCH['empId'])
    cT = find_col(headers, COL_SEARCH['txnId'])
    cN = find_col(headers, COL_SEARCH['name'])
    cS = find_col(headers, COL_SEARCH['stateCode'])
    cD = find_col(headers, COL_SEARCH['totalDist'])
    cW = find_col(headers, COL_SEARCH['totalWith'])
    cR = find_col(headers, COL_SEARCH['regionTax'])
    cI = find_col(headers, COL_SEARCH['incomeRegion'])

    if cE < 0 or cT < 0:
        return []

    groups, order = {}, []
    last_emp, last_txn = None, None

    for row in rows[header_row_idx + 1:]:
        emp_id = str(row[cE]).strip() if cE >= 0 and row[cE] else None
        txn_id = str(row[cT]).strip() if cT >= 0 and row[cT] else None
        if not emp_id and not txn_id and all(v is None for v in row):
            continue
        if not emp_id: emp_id = last_emp
        if not txn_id: txn_id = last_txn
        if not emp_id or not txn_id: continue
        last_emp, last_txn = emp_id, txn_id

        key = f"{emp_id}|{txn_id}"
        if key not in groups:
            groups[key] = {'empId': emp_id, 'txnId': str(txn_id), 'name': '', 'rows': []}
            order.append(key)
        if cN >= 0 and row[cN] and not groups[key]['name']:
            groups[key]['name'] = str(row[cN])

        sc = row[cS] if cS >= 0 else None
        groups[key]['rows'].append({
            'stateCode':    resolve_state_code(sc),
            'totalDist':    to_num(row[cD] if cD >= 0 else None),
            'totalWith':    to_num(row[cW] if cW >= 0 else None),
            'regionTax':    to_num(row[cR] if cR >= 0 else None),
            'incomeRegion': to_num(row[cI] if cI >= 0 else None),
        })

    return [groups[k] for k in order]

# ── EIB row generation ────────────────────────────────────────
def make_empty_row():
    return [None] * TOTAL_COLS

def build_base(settings, spreadsheet_key, emp_id):
    r = make_empty_row()
    r[C['spreadsheetKey']] = spreadsheet_key
    r[C['batchId']]        = settings['batchId']
    r[C['paymentId']]      = f'="Trailing2State"&D&F'  # replaced with real formula on export
    r[C['employee']]       = emp_id
    r[C['paymentDate']]    = settings['paymentDate']
    r[C['periodDate']]     = settings['periodDate']
    r[C['payPriority']]    = int(settings['payPriority'])
    r[C['runCategory']]    = settings['runCategory']
    r[C['resultType']]     = settings['resultType']
    r[C['reason']]         = settings['reason']
    return r

def generate_eib_rows(records, settings):
    eib = []
    for rec in records:
        primary = next((r for r in rec['rows'] if r['totalDist'] is not None), rec['rows'][0])
        key = rec['txnId']
        line = 1

        # TVRSU
        r = build_base(settings, key, rec['empId'])
        r[C['rowId']]  = line; line += 1
        r[C['earning']] = 'TVRSU'
        r[C['amount']]  = primary['totalDist']
        eib.append({'type': 'TVRSU', 'empId': rec['empId'], 'txnId': rec['txnId'], 'data': r})

        # NVRSU
        r = build_base(settings, key, rec['empId'])
        r[C['rowId']]  = line; line += 1
        r[C['earning']] = 'NVRSU'
        r[C['amount']]  = primary['totalWith']
        eib.append({'type': 'NVRSU', 'empId': rec['empId'], 'txnId': rec['txnId'], 'data': r})

        # W_SWW per state
        for sr in rec['rows']:
            # Skip rows where both regionTax and incomeRegion are 0 / None
            all_zero = (sr['regionTax'] in (0, None)) and (sr['incomeRegion'] in (0, None))
            if all_zero:
                continue
            ln = line; line += 1

            # W_SWW + W_CWCGW
            r = build_base(settings, key, rec['empId'])
            r[C['rowId']]         = ln
            r[C['deduction']]     = 'W_SWW'
            r[C['amount']]        = sr['regionTax']
            r[C['stateAuthority']]= sr['stateCode']
            r[C['subRowId']]      = 1
            r[C['relatedCalc']]   = 'W_CWCGW'
            r[C['value']]         = sr['incomeRegion']
            eib.append({'type': 'W_SWW', 'empId': rec['empId'], 'txnId': rec['txnId'],
                        'stateCode': sr['stateCode'], 'data': r})

            # W_TTWG
            r = build_base(settings, key, rec['empId'])
            r[C['rowId']]      = ln
            r[C['subRowId']]   = 2
            r[C['relatedCalc']]= 'W_TTWG'
            r[C['value']]      = sr['incomeRegion']
            eib.append({'type': 'sub', 'empId': rec['empId'], 'txnId': rec['txnId'], 'data': r})

            # W_TXWG
            r = build_base(settings, key, rec['empId'])
            r[C['rowId']]      = ln
            r[C['subRowId']]   = 3
            r[C['relatedCalc']]= 'W_TXWG'
            r[C['value']]      = sr['incomeRegion']
            eib.append({'type': 'sub', 'empId': rec['empId'], 'txnId': rec['txnId'], 'data': r})

    return eib

# ── Display columns for preview table ────────────────────────
DISPLAY = [
    ('Spr. Key',    C['spreadsheetKey']),
    ('Batch ID',    C['batchId']),
    ('Payment ID',  C['paymentId']),
    ('Employee',    C['employee']),
    ('Pay Date',    C['paymentDate']),
    ('Period Date', C['periodDate']),
    ('Priority',    C['payPriority']),
    ('Run Cat.',    C['runCategory']),
    ('Result Type', C['resultType']),
    ('Reason',      C['reason']),
    ('Row ID',      C['rowId']),
    ('Earning',     C['earning']),
    ('Deduction',   C['deduction']),
    ('Amount',      C['amount']),
    ('State Auth',  C['stateAuthority']),
    ('Sub Row ID',  C['subRowId']),
    ('Rel. Calc',   C['relatedCalc']),
    ('Value',       C['value']),
]

def eib_to_df(eib_rows):
    """Convert EIB rows to a display DataFrame."""
    records = []
    for row in eib_rows:
        records.append({label: row['data'][idx] for label, idx in DISPLAY})
    return pd.DataFrame(records)

def eib_to_excel_bytes(eib_rows, settings):
    """Export all 134 columns to Excel bytes for download."""
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Payroll Off cycle Payment'

    # Header row
    header = make_empty_row()
    for i, h in enumerate(EIB_HEADERS):
        header[i] = h
    ws.append(header)

    # Data rows
    for i, row in enumerate(eib_rows):
        d = list(row['data'])
        excel_row = i + 2
        d[C['paymentId']] = f'=CONCATENATE("Trailing2State",D{excel_row},F{excel_row})'
        ws.append(d)

    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ══════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ══════════════════════════════════════════════════════════════

st.markdown("""<h1 style='margin-bottom:0'>🧾 EIB Assistant</h1>
<p style='color:#666;margin-top:4px;font-size:1rem'>RSU Tax Formatter — Payroll Off-cycle Payment</p>
""", unsafe_allow_html=True)
st.divider()

# ── Step 1: Upload ────────────────────────────────────────────
st.subheader("Step 1 · Load Output File")
uploaded = st.file_uploader(
    "Upload your output spreadsheet (Original / Deloitte sheet)",
    type=['xlsx', 'xls'],
    label_visibility='collapsed'
)

if not uploaded:
    st.info("Upload your output file to get started.")
    st.stop()

# Load workbook
wb_src = openpyxl.load_workbook(uploaded, data_only=True)
st.success(f"✅ Loaded: **{uploaded.name}**  |  Sheets: {', '.join(wb_src.sheetnames)}", icon="📂")

# ── Step 2: Select Sheet / Employee / Transaction ─────────────
st.subheader("Step 2 · Select Employee / Transaction")

col1, col2, col3 = st.columns([1.2, 1.2, 1.2])

preferred = ['Original', 'Deloitte', 'QUESTIONS_']
ordered_sheets = [s for s in preferred if s in wb_src.sheetnames] + \
                 [s for s in wb_src.sheetnames if s not in preferred]

with col1:
    sheet_name = st.selectbox("Source Sheet", ordered_sheets)

ws_src = wb_src[sheet_name]
parsed = parse_sheet(ws_src)

if not parsed:
    st.error("Could not find expected columns in this sheet. Try a different sheet.")
    st.stop()

all_emps = sorted(set(r['empId'] for r in parsed))

with col2:
    emp_id = st.selectbox("Employee ID", ['— choose —'] + all_emps)

if emp_id == '— choose —':
    st.info("Select an Employee ID to continue.")
    st.stop()

all_txns = [r['txnId'] for r in parsed if r['empId'] == emp_id]
with col3:
    txn_choice = st.selectbox("Transaction ID", ['All'] + all_txns)

# Input preview
recs = [r for r in parsed if r['empId'] == emp_id and
        (txn_choice == 'All' or r['txnId'] == txn_choice)]

with st.expander("📋 Source data preview", expanded=False):
    preview_rows = []
    for rec in recs:
        for sr in rec['rows']:
            preview_rows.append({
                'Employee ID': rec['empId'],
                'Transaction ID': rec['txnId'],
                'Name': rec['name'],
                'State Code': sr['stateCode'],
                'Total Dist / Gross Gain': sr['totalDist'],
                'EE Withholding': sr['totalWith'],
                'Region Tax': sr['regionTax'],
                'Income for Region': sr['incomeRegion'],
            })
    st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

# ── Step 3: EIB Settings ──────────────────────────────────────
st.subheader("Step 3 · EIB Settings")

c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
with c1: batch_id      = st.text_input("Batch ID",        value="RSU May 26")
with c2: payment_date  = st.text_input("Payment Date",    value="2026-05-01")
with c3: period_date   = st.text_input("Period Date",     value="2026-05-01")
with c4: pay_priority  = st.text_input("Payment Priority",value="3")
with c5: run_category  = st.text_input("Run Category",    value="RSU_USA")
with c6: result_type   = st.text_input("Result Type",     value="OnDemandPayment")
with c7: reason        = st.text_input("Reason",          value="Stock")

settings = dict(
    batchId=batch_id, paymentDate=payment_date, periodDate=period_date,
    payPriority=pay_priority, runCategory=run_category,
    resultType=result_type, reason=reason,
)

# ── Step 4: Generate & Preview ───────────────────────────────
st.subheader("Step 4 · EIB Output")

eib_rows = generate_eib_rows(recs, settings)
df = eib_to_df(eib_rows)

col_m1, col_m2, col_m3 = st.columns(3)
with col_m1: st.metric("Total EIB Rows", len(eib_rows))
with col_m2: st.metric("Transactions", len(recs))
with col_m3: st.metric("States Processed",
    sum(1 for r in eib_rows if r['type'] == 'W_SWW'))

# Colour-coded preview
def colour_row(row):
    styles = ['' for _ in row]
    earning = row.get('Earning', '')
    deduction = row.get('Deduction', '')
    rel_calc = row.get('Rel. Calc', '')
    if earning == 'TVRSU':
        return ['background-color: #e6f4ea'] * len(row)
    elif earning == 'NVRSU':
        return ['background-color: #e8f0fe'] * len(row)
    elif deduction == 'W_SWW':
        return ['background-color: #fce8b2'] * len(row)
    elif rel_calc in ('W_TTWG', 'W_TXWG'):
        return ['background-color: #fdf0e0'] * len(row)
    return styles

st.dataframe(
    df.style.apply(colour_row, axis=1),
    use_container_width=True,
    hide_index=True,
    height=450,
)

# Legend
st.markdown("""
<div style='display:flex;gap:20px;flex-wrap:wrap;margin-top:6px;font-size:0.8rem'>
  <span><span style='background:#c6e8ce;padding:2px 10px;border-radius:4px'>■</span> TVRSU – Total Distribution</span>
  <span><span style='background:#c2d8fb;padding:2px 10px;border-radius:4px'>■</span> NVRSU – EE Withholding</span>
  <span><span style='background:#fad67d;padding:2px 10px;border-radius:4px'>■</span> W_SWW – State Tax</span>
  <span><span style='background:#fce4c3;padding:2px 10px;border-radius:4px'>■</span> Sub-calc lines</span>
</div>
""", unsafe_allow_html=True)

st.divider()

# ── Download ──────────────────────────────────────────────────
col_dl1, col_dl2 = st.columns([1, 3])
with col_dl1:
    st.download_button(
        label="⬇️ Download EIB Excel",
        data=eib_to_excel_bytes(eib_rows, settings),
        file_name="eib_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

# ── Bulk Section ──────────────────────────────────────────────
st.subheader("⚡ Bulk — All Employees")
st.caption("Generate and download EIB rows for every employee in the selected sheet at once.")

if st.button("Generate All EIB Rows", type="secondary"):
    all_eib = generate_eib_rows(parsed, settings)
    df_all = eib_to_df(all_eib)
    st.success(f"✅ {len(all_eib)} EIB rows generated for {len(parsed)} transactions.")
    st.dataframe(df_all, use_container_width=True, hide_index=True, height=400)
    st.download_button(
        label="⬇️ Download Bulk EIB Excel",
        data=eib_to_excel_bytes(all_eib, settings),
        file_name="eib_bulk_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
