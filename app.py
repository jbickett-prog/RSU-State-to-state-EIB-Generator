import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
import io
from datetime import date

# ── Page config ───────────────────────────────────────────
st.set_page_config(
    page_title="EIB Assistant – RSU Tax Formatter",
    page_icon="🧾",
    layout="wide",
)

# ── State abbrev → Payroll Authority Tax Code ─────────────
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

TOTAL_COLS = 134  # EIB has 134 columns

# EIB column indices (0-based)
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

# ── Helpers ───────────────────────────────────────────────
def resolve_state(raw):
    if raw is None or str(raw).strip() == '':
        return ''
    s = str(raw).strip().upper()
    if s in STATE_MAP:
        return STATE_MAP[s]
    try:
        return str(int(s)).zfill(2)
    except:
        return s

def to_num(v):
    if v is None or str(v).strip() == '':
        return None
    try:
        return float(v)
    except:
        return None

def find_col(headers, keys):
    for key in keys:
        for i, h in enumerate(headers):
            if h and key.lower() in str(h).lower().replace('\n', ' '):
                return i
    return -1

# ── Parse uploaded Excel ──────────────────────────────────
def parse_sheet(ws_data):
    """ws_data: list of lists (rows of values)"""
    header_row_idx = -1
    headers = []
    for i, row in enumerate(ws_data[:5]):
        joined = ' '.join(str(c) for c in row if c).lower()
        if 'employee' in joined or 'transaction' in joined:
            header_row_idx = i
            headers = [str(c) if c else '' for c in row]
            break
    if header_row_idx < 0:
        return []

    c_emp   = find_col(headers, ['employee id'])
    c_txn   = find_col(headers, ['transaction id'])
    c_name  = find_col(headers, ['name'])
    c_state = find_col(headers, ['payroll authority tax code', 'state code', 'region code'])
    c_tdist = find_col(headers, ['total distribution/exercise gross gain', 'reportable income'])
    c_twith = find_col(headers, ['total employee withholding (actual & hypo taxes)', 'country total actual tax ee liability'])
    c_rtax  = find_col(headers, ['region tax employee liability'])
    c_inc   = find_col(headers, ['income for region tax pre gross up'])

    if c_emp < 0 or c_txn < 0:
        return []

    groups = {}
    order = []
    last_emp, last_txn = None, None

    for row in ws_data[header_row_idx + 1:]:
        def get(i): return row[i] if i >= 0 and i < len(row) else None

        emp_id = str(get(c_emp)).strip() if get(c_emp) else None
        txn_id = str(get(c_txn)).strip() if get(c_txn) else None

        if not emp_id and not txn_id and all(v is None or str(v).strip() == '' for v in row):
            continue
        emp_id = emp_id or last_emp
        txn_id = txn_id or last_txn
        if not emp_id or not txn_id:
            continue

        last_emp, last_txn = emp_id, txn_id
        key = f"{emp_id}|{txn_id}"
        if key not in groups:
            groups[key] = {'empId': emp_id, 'txnId': txn_id, 'name': '', 'rows': []}
            order.append(key)

        if c_name >= 0 and get(c_name) and not groups[key]['name']:
            groups[key]['name'] = str(get(c_name))

        groups[key]['rows'].append({
            'stateCode':    resolve_state(get(c_state)),
            'totalDist':    to_num(get(c_tdist)),
            'totalWith':    to_num(get(c_twith)),
            'regionTax':    to_num(get(c_rtax)),
            'incomeRegion': to_num(get(c_inc)),
        })

    return [groups[k] for k in order]

# ── Build one blank EIB row ───────────────────────────────
def blank_row():
    return [None] * TOTAL_COLS

def base_row(txn_id, emp_id, settings):
    r = blank_row()
    r[C['spreadsheetKey']] = txn_id
    r[C['batchId']]        = settings['batch_id']
    r[C['paymentId']]      = f'="Trailing2State"&D&F'   # placeholder; real formula set on export
    r[C['employee']]       = emp_id
    r[C['paymentDate']]    = settings['payment_date']
    r[C['periodDate']]     = settings['period_date']
    r[C['payPriority']]    = int(settings['priority'])
    r[C['runCategory']]    = settings['run_category']
    r[C['resultType']]     = settings['result_type']
    r[C['reason']]         = settings['reason']
    return r

# ── Generate EIB rows ─────────────────────────────────────
def generate_eib(records, settings):
    eib_rows = []
    for rec in records:
        emp_id = rec['empId']
        txn_id = rec['txnId']
        primary = next((r for r in rec['rows'] if r['totalDist'] is not None), rec['rows'][0])
        line = 1

        # TVRSU
        r = base_row(txn_id, emp_id, settings)
        r[C['rowId']]   = line; line += 1
        r[C['earning']] = 'TVRSU'
        r[C['amount']]  = primary['totalDist']
        eib_rows.append({'type': 'TVRSU', 'emp': emp_id, 'txn': txn_id, 'data': r})

        # NVRSU
        r = base_row(txn_id, emp_id, settings)
        r[C['rowId']]   = line; line += 1
        r[C['earning']] = 'NVRSU'
        r[C['amount']]  = primary['totalWith']
        eib_rows.append({'type': 'NVRSU', 'emp': emp_id, 'txn': txn_id, 'data': r})

        # W_SWW per state
        for sr in rec['rows']:
            if (sr['regionTax'] or 0) == 0 and (sr['incomeRegion'] or 0) == 0:
                continue  # skip zero rows
            ln = line; line += 1

            # W_SWW + W_CWCGW
            r = base_row(txn_id, emp_id, settings)
            r[C['rowId']]         = ln
            r[C['deduction']]     = 'W_SWW'
            r[C['amount']]        = sr['regionTax']
            r[C['stateAuthority']]= sr['stateCode']
            r[C['subRowId']]      = 1
            r[C['relatedCalc']]   = 'W_CWCGW'
            r[C['value']]         = sr['incomeRegion']
            eib_rows.append({'type': 'W_SWW', 'emp': emp_id, 'txn': txn_id, 'state': sr['stateCode'], 'data': r})

            # W_TTWG
            r = base_row(txn_id, emp_id, settings)
            r[C['rowId']]       = ln
            r[C['subRowId']]    = 2
            r[C['relatedCalc']] = 'W_TTWG'
            r[C['value']]       = sr['incomeRegion']
            eib_rows.append({'type': 'sub', 'emp': emp_id, 'txn': txn_id, 'data': r})

            # W_TXWG
            r = base_row(txn_id, emp_id, settings)
            r[C['rowId']]       = ln
            r[C['subRowId']]    = 3
            r[C['relatedCalc']] = 'W_TXWG'
            r[C['value']]       = sr['incomeRegion']
            eib_rows.append({'type': 'sub', 'emp': emp_id, 'txn': txn_id, 'data': r})

    return eib_rows

# ── Build Excel output ────────────────────────────────────
def build_excel(eib_rows):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Payroll Off cycle Payment'

    # Header row
    header_row_out = [None] * TOTAL_COLS
    for i, h in enumerate(EIB_HEADERS):
        header_row_out[i] = h
    ws.append(header_row_out)

    # Style header
    hdr_fill = PatternFill('solid', fgColor='F6F8FA')
    for cell in ws[1]:
        cell.fill = hdr_fill
        cell.font = Font(bold=True, size=9)
        cell.alignment = Alignment(wrap_text=False)

    # Fill colours
    fills = {
        'TVRSU': PatternFill('solid', fgColor='C6E8CE'),
        'NVRSU': PatternFill('solid', fgColor='C2D8FB'),
        'W_SWW': PatternFill('solid', fgColor='FAD67D'),
        'sub':   PatternFill('solid', fgColor='FCE4C3'),
    }

    for i, row in enumerate(eib_rows, start=2):
        data = list(row['data'])
        # Inject real Payment ID formula
        col_d = openpyxl.utils.get_column_letter(C['batchId'] + 1)
        col_f = openpyxl.utils.get_column_letter(C['employee'] + 1)
        data[C['paymentId']] = f'="Trailing2State"&{col_d}{i}&{col_f}{i}'
        ws.append(data)
        fill = fills.get(row['type'])
        if fill:
            for cell in ws[i]:
                cell.fill = fill

    # Column widths for used columns
    used = {
        C['spreadsheetKey']: 14, C['batchId']: 14, C['paymentId']: 30,
        C['employee']: 14, C['paymentDate']: 13, C['periodDate']: 13,
        C['payPriority']: 10, C['runCategory']: 12, C['resultType']: 18,
        C['reason']: 10, C['rowId']: 8, C['earning']: 10, C['deduction']: 12,
        C['amount']: 12, C['stateAuthority']: 13, C['subRowId']: 10,
        C['relatedCalc']: 14, C['value']: 12,
    }
    for col_idx, width in used.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx + 1)].width = width

    ws.freeze_panes = 'A2'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Display helpers ───────────────────────────────────────
ROW_COLORS = {'TVRSU': '#c6e8ce', 'NVRSU': '#c2d8fb', 'W_SWW': '#fad67d', 'sub': '#fce4c3'}

def display_preview(eib_rows):
    display_cols = [
        ('Spr. Key',    'spreadsheetKey'),
        ('Batch ID',    'batchId'),
        ('Employee',    'employee'),
        ('Pay Date',    'paymentDate'),
        ('Row ID',      'rowId'),
        ('Earning',     'earning'),
        ('Deduction',   'deduction'),
        ('Amount',      'amount'),
        ('State Auth',  'stateAuthority'),
        ('Sub Row',     'subRowId'),
        ('Rel. Calc',   'relatedCalc'),
        ('Value',       'value'),
    ]

    rows_out = []
    for row in eib_rows:
        d = row['data']
        r = {label: d[C[key]] for label, key in display_cols}
        r['_type'] = row['type']
        rows_out.append(r)

    df = pd.DataFrame(rows_out)
    df = df.rename(columns={c: c for c in df.columns if c != '_type'})

    def style_row(row):
        color = ROW_COLORS.get(row['_type'], '#ffffff')
        return [f'background-color: {color}'] * len(row)

    display_df = df.drop(columns=['_type'])
    styled = display_df.style.apply(
        lambda row: [f'background-color: {ROW_COLORS.get(df.loc[row.name, "_type"], "#ffffff")}'] * len(row),
        axis=1
    )
    st.dataframe(styled, use_container_width=True, height=400)

# ══════════════════════════════════════════════════════════
#  MAIN APP
# ══════════════════════════════════════════════════════════

st.markdown("## 🧾 EIB Assistant – RSU Tax Formatter")
st.caption("Generates Workday-ready EIB rows from your RSU output spreadsheet")

# ── STEP 1: Upload ────────────────────────────────────────
st.markdown("### Step 1 · Upload Output File")
uploaded = st.file_uploader("Drop your output .xlsx file here", type=['xlsx', 'xls'])

if not uploaded:
    st.info("Upload your RSU output spreadsheet to get started.")
    st.stop()

wb_in = openpyxl.load_workbook(uploaded, data_only=True)
preferred = ['Original', 'Deloitte', 'QUESTIONS_']
sheets = wb_in.sheetnames
ordered_sheets = [s for s in preferred if s in sheets] + [s for s in sheets if s not in preferred]

# ── STEP 2: Select ────────────────────────────────────────
st.markdown("### Step 2 · Select Employee / Transaction")
col1, col2, col3 = st.columns(3)

with col1:
    sheet_name = st.selectbox("Source Sheet", ordered_sheets)

ws_data = [[cell.value for cell in row] for row in wb_in[sheet_name].iter_rows()]
records = parse_sheet(ws_data)

if not records:
    st.error("Could not parse the selected sheet. Make sure it has Employee ID and Transaction ID columns.")
    st.stop()

emp_ids = ['— All —'] + sorted(set(r['empId'] for r in records))
with col2:
    emp_sel = st.selectbox("Employee ID", emp_ids)

txn_options = ['— All —']
if emp_sel != '— All —':
    txn_options += [r['txnId'] for r in records if r['empId'] == emp_sel]
with col3:
    txn_sel = st.selectbox("Transaction ID", txn_options)

# ── STEP 3: EIB Settings ──────────────────────────────────
st.markdown("### Step 3 · EIB Settings")
c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
batch_id     = c1.text_input("Batch ID",        value="RSU May 26")
payment_date = c2.text_input("Payment Date",    value="2026-05-01")
period_date  = c3.text_input("Period Date",     value="2026-05-01")
priority     = c4.text_input("Priority",        value="3")
run_cat      = c5.text_input("Run Category",    value="RSU_USA")
result_type  = c6.text_input("Result Type",     value="OnDemandPayment")
reason       = c7.text_input("Reason",          value="Stock")

settings = dict(
    batch_id=batch_id, payment_date=payment_date, period_date=period_date,
    priority=priority, run_category=run_cat, result_type=result_type, reason=reason,
)

# ── STEP 4: Generate ─────────────────────────────────────
st.markdown("### Step 4 · Generate EIB Output")

tab1, tab2 = st.tabs(["Single Employee", "Bulk — All Employees"])

with tab1:
    filtered = [r for r in records
                if (emp_sel == '— All —' or r['empId'] == emp_sel)
                and (txn_sel == '— All —' or r['txnId'] == txn_sel)]

    if not filtered:
        st.warning("No records match the current selection.")
    else:
        eib_rows = generate_eib(filtered, settings)
        st.success(f"✅ {len(eib_rows)} EIB rows generated for {len(filtered)} transaction(s)")

        # Legend
        lc1, lc2, lc3, lc4 = st.columns(4)
        lc1.markdown("🟢 TVRSU – Total Dist.")
        lc2.markdown("🔵 NVRSU – EE Withholding")
        lc3.markdown("🟡 W_SWW – State Tax")
        lc4.markdown("🟠 Sub-calc lines")

        display_preview(eib_rows)

        excel_buf = build_excel(eib_rows)
        st.download_button(
            label="⬇️ Download EIB Excel",
            data=excel_buf,
            file_name=f"eib_output_{emp_sel.replace('— All —','all')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with tab2:
    all_eib = generate_eib(records, settings)
    st.success(f"✅ {len(all_eib)} EIB rows generated for {len(records)} total transactions")
    display_preview(all_eib)

    bulk_buf = build_excel(all_eib)
    st.download_button(
        label="⬇️ Download Bulk EIB Excel",
        data=bulk_buf,
        file_name="eib_bulk_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
