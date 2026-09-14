
import streamlit as st
import pandas as pd
import openpyxl
import io

# ── Page config ────────────────────────────────────────────
st.set_page_config(
    page_title="EIB Assistant – RSU Tax Formatter",
    page_icon="🧾",
    layout="wide",
)

# ── CSS: match the HTML app exactly ───────────────────────
st.markdown("""
<style>
/* Hide default streamlit chrome */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 0 !important; max-width: 100% !important; }

/* ── App shell ── */
.eib-app { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
           background: #f0f2f5; color: #1a1a2e; min-height: 100vh; }

/* ── Header ── */
.eib-header {
  background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
  color: white; padding: 18px 32px; display: flex; align-items: baseline;
  gap: 14px; box-shadow: 0 2px 12px rgba(0,0,0,0.3); margin-bottom: 24px;
}
.eib-logo   { font-size: 1.6rem; font-weight: 700; letter-spacing: -0.5px; }
.eib-sub    { font-size: 0.9rem; color: #94b4e4; font-weight: 400; }

/* ── Cards ── */
.eib-card {
  background: white; border-radius: 14px; padding: 28px 32px;
  margin: 0 24px 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); position: relative;
}
.eib-step-badge {
  position: absolute; top: 24px; right: 28px;
  background: #e8f0fe; color: #1967d2; font-size: 0.72rem; font-weight: 700;
  padding: 3px 10px; border-radius: 20px; text-transform: uppercase; letter-spacing: 0.5px;
}
.eib-card h2 { font-size: 1.15rem; font-weight: 700; margin-bottom: 6px; color: #1a1a2e; }
.eib-card p  { color: #555; font-size: 0.9rem; margin-bottom: 14px; }

/* ── Legend ── */
.eib-legend { display: flex; gap: 18px; flex-wrap: wrap; margin-bottom: 14px; }
.eib-legend-item { display: flex; align-items: center; gap: 7px; font-size: 0.78rem; color: #555; }
.eib-dot { width: 12px; height: 12px; border-radius: 3px; flex-shrink: 0; display: inline-block; }
.dot-tvrsu { background:#c6e8ce; border:1px solid #a8d5b5; }
.dot-nvrsu { background:#c2d8fb; border:1px solid #a0c0f5; }
.dot-wsww  { background:#fad67d; border:1px solid #f0bc40; }
.dot-sub   { background:#fce4c3; border:1px solid #f0c890; }

/* ── Table ── */
.eib-table-wrap { overflow-x: auto; }
.eib-table { border-collapse: collapse; width: 100%; font-size: 0.84rem; }
.eib-table th {
  background: #f0f2f5; color: #444; font-weight: 700; padding: 9px 12px;
  text-align: left; border-bottom: 2px solid #d1d5db; white-space: nowrap;
  font-size: 0.78rem; text-transform: uppercase; letter-spacing: 0.3px;
}
.eib-table td { padding: 8px 12px; border-bottom: 1px solid #f0f2f5; white-space: nowrap; }
.eib-table tr.row-tvrsu td { background: #e6f4ea; }
.eib-table tr.row-nvrsu td { background: #e8f0fe; }
.eib-table tr.row-wsww  td { background: #fce8b2; }
.eib-table tr.row-sub   td { background: #fdf0e0; }
.eib-table tr.row-div   td { border-top: 2.5px solid #c0cfe8; padding: 0; height: 3px; }
.eib-table td.num       { text-align: right; font-variant-numeric: tabular-nums; }

/* ── Row count badge ── */
.row-count-badge {
  display: inline-block; background: #f0f4ff; color: #1967d2;
  font-size: 0.82rem; font-weight: 700; padding: 4px 12px;
  border-radius: 20px; border: 1px solid #c8d8f8; margin-left: 8px;
}

/* ── Export success msg ── */
.export-ok {
  margin-top: 12px; background: #e6f4ea; color: #137333;
  padding: 10px 16px; border-radius: 8px; font-size: 0.88rem; font-weight: 600;
}

/* ── Streamlit widget overrides ── */
div[data-testid="stFileUploader"] > label { display: none; }
div[data-testid="stFileUploader"] section {
  border: 2px dashed #c5cdd8 !important; border-radius: 10px !important;
  background: #fafbfc !important; padding: 24px !important;
}
div[data-testid="stFileUploader"] section:hover {
  border-color: #1967d2 !important; background: #e8f0fe !important;
}
div[data-testid="stSelectbox"] label, div[data-testid="stTextInput"] label {
  font-size: 0.78rem !important; font-weight: 600 !important;
  color: #666 !important; text-transform: uppercase; letter-spacing: 0.4px;
}
div[data-testid="stDownloadButton"] button, div[data-testid="stButton"] button {
  border-radius: 8px !important; font-weight: 600 !important; font-size: 0.9rem !important;
}
/* Primary blue button */
div[data-testid="stDownloadButton"] button {
  background-color: #1967d2 !important; color: white !important;
  border: none !important;
}
div[data-testid="stDownloadButton"] button:hover { background-color: #1558b8 !important; }
</style>
""", unsafe_allow_html=True)

# ── State map ──────────────────────────────────────────────
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

def resolve_state(raw):
    if not raw and raw != 0: return ''
    s = str(raw).strip().upper()
    if s in STATE_MAP: return STATE_MAP[s]
    try: return str(int(s)).zfill(2)
    except: return s

# ── EIB constants ──────────────────────────────────────────
TOTAL_COLS = 134
C = dict(spreadsheetKey=1,batchId=3,paymentId=4,employee=5,paymentDate=6,
         periodDate=7,payPriority=8,runCategory=9,resultType=11,reason=13,
         rowId=39,earning=40,deduction=41,amount=43,stateAuthority=57,
         subRowId=85,relatedCalc=86,value=87)

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

DISPLAY = [
    ('Spr. Key',   C['spreadsheetKey']),('Batch ID',    C['batchId']),
    ('Payment ID', C['paymentId']),     ('Employee',    C['employee']),
    ('Pay Date',   C['paymentDate']),   ('Period Date', C['periodDate']),
    ('Priority',   C['payPriority']),   ('Run Cat.',    C['runCategory']),
    ('Result Type',C['resultType']),    ('Reason',      C['reason']),
    ('Row ID',     C['rowId']),         ('Earning',     C['earning']),
    ('Deduction',  C['deduction']),     ('Amount',      C['amount']),
    ('State Auth', C['stateAuthority']),('Sub Row ID',  C['subRowId']),
    ('Rel. Calc',  C['relatedCalc']),   ('Value',       C['value']),
]

COL_KEYS = dict(
    empId        = ['employee id'],
    txnId        = ['transaction id'],
    name         = ['name'],
    stateCode    = ['payroll authority tax code','state code','region code'],
    totalDist    = ['total distribution/exercise gross gain','reportable income'],
    totalWith    = ['total employee withholding (actual & hypo taxes)','country total actual tax ee liability'],
    regionTax    = ['region tax employee liability'],
    incomeRegion = ['income for region tax pre gross up'],
)

def find_col(headers, keys):
    for key in keys:
        for i, h in enumerate(headers):
            if h and key in str(h).lower().replace('\n',' '):
                return i
    return -1

def to_num(v):
    try: return float(v) if v not in (None,'') else None
    except: return None

def parse_sheet(ws):
    rows = list(ws.iter_rows(values_only=True))
    header_idx, headers = -1, []
    for i, row in enumerate(rows[:5]):
        joined = ' '.join(str(c).lower() for c in row if c)
        if 'employee' in joined or 'transaction' in joined:
            header_idx = i
            headers = [str(c) if c else '' for c in row]
            break
    if header_idx < 0: return []

    cE=find_col(headers,COL_KEYS['empId']);   cT=find_col(headers,COL_KEYS['txnId'])
    cN=find_col(headers,COL_KEYS['name']);    cS=find_col(headers,COL_KEYS['stateCode'])
    cD=find_col(headers,COL_KEYS['totalDist']);cW=find_col(headers,COL_KEYS['totalWith'])
    cR=find_col(headers,COL_KEYS['regionTax']);cI=find_col(headers,COL_KEYS['incomeRegion'])
    if cE<0 or cT<0: return []

    groups, order, last_e, last_t = {}, [], None, None
    for row in rows[header_idx+1:]:
        e = str(row[cE]).strip() if cE>=0 and row[cE] else None
        t = str(row[cT]).strip() if cT>=0 and row[cT] else None
        if not e and not t and all(v is None for v in row): continue
        if not e: e = last_e
        if not t: t = last_t
        if not e or not t: continue
        last_e, last_t = e, t
        key = f"{e}|{t}"
        if key not in groups:
            groups[key] = {'empId':e,'txnId':str(t),'name':'','rows':[]}
            order.append(key)
        if cN>=0 and row[cN] and not groups[key]['name']:
            groups[key]['name'] = str(row[cN])
        sc = row[cS] if cS>=0 else None
        groups[key]['rows'].append({
            'stateCode':    resolve_state(sc),
            'totalDist':    to_num(row[cD] if cD>=0 else None),
            'totalWith':    to_num(row[cW] if cW>=0 else None),
            'regionTax':    to_num(row[cR] if cR>=0 else None),
            'incomeRegion': to_num(row[cI] if cI>=0 else None),
        })
    return [groups[k] for k in order]

def empty_row(): return [None]*TOTAL_COLS

def base_row(s, skey, emp):
    r = empty_row()
    r[C['spreadsheetKey']] = skey
    r[C['batchId']]        = s['batchId']
    r[C['paymentId']]      = '="Trailing2State"&D&F'
    r[C['employee']]       = emp
    r[C['paymentDate']]    = s['paymentDate']
    r[C['periodDate']]     = s['periodDate']
    r[C['payPriority']]    = int(s['payPriority'])
    r[C['runCategory']]    = s['runCategory']
    r[C['resultType']]     = s['resultType']
    r[C['reason']]         = s['reason']
    return r

def gen_eib(records, s):
    out = []
    for rec in records:
        primary = next((r for r in rec['rows'] if r['totalDist'] is not None), rec['rows'][0])
        key = rec['txnId']; line = 1

        r = base_row(s, key, rec['empId'])
        r[C['rowId']]=line; line+=1; r[C['earning']]='TVRSU'; r[C['amount']]=primary['totalDist']
        out.append({'type':'TVRSU','empId':rec['empId'],'txnId':rec['txnId'],'data':r})

        r = base_row(s, key, rec['empId'])
        r[C['rowId']]=line; line+=1; r[C['earning']]='NVRSU'; r[C['amount']]=primary['totalWith']
        out.append({'type':'NVRSU','empId':rec['empId'],'txnId':rec['txnId'],'data':r})

        for sr in rec['rows']:
            if (sr['regionTax'] in (0,None)) and (sr['incomeRegion'] in (0,None)): continue
            ln = line; line += 1

            r = base_row(s, key, rec['empId'])
            r[C['rowId']]=ln; r[C['deduction']]='W_SWW'; r[C['amount']]=sr['regionTax']
            r[C['stateAuthority']]=sr['stateCode']; r[C['subRowId']]=1
            r[C['relatedCalc']]='W_CWCGW'; r[C['value']]=sr['incomeRegion']
            out.append({'type':'W_SWW','empId':rec['empId'],'txnId':rec['txnId'],'data':r})

            r = base_row(s, key, rec['empId'])
            r[C['rowId']]=ln; r[C['subRowId']]=2; r[C['relatedCalc']]='W_TTWG'; r[C['value']]=sr['incomeRegion']
            out.append({'type':'sub','empId':rec['empId'],'txnId':rec['txnId'],'data':r})

            r = base_row(s, key, rec['empId'])
            r[C['rowId']]=ln; r[C['subRowId']]=3; r[C['relatedCalc']]='W_TXWG'; r[C['value']]=sr['incomeRegion']
            out.append({'type':'sub','empId':rec['empId'],'txnId':rec['txnId'],'data':r})
    return out

def fmt(v):
    if v is None: return ''
    if isinstance(v, float): return f"{v:,.2f}"
    return str(v)

def render_table(eib_rows):
    cols = DISPLAY
    last_txn = None
    rows_html = ''
    for row in eib_rows:
        cls = {'TVRSU':'row-tvrsu','NVRSU':'row-nvrsu','W_SWW':'row-wsww','sub':'row-sub'}.get(row['type'],'')
        if row['txnId'] != last_txn and last_txn is not None:
            rows_html += f'<tr class="row-div"><td colspan="{len(cols)}"></td></tr>'
        last_txn = row['txnId']
        cells = ''
        for label, idx in cols:
            v = row['data'][idx]
            num_cls = ' class="num"' if isinstance(v, (int,float)) and not isinstance(v, bool) else ''
            cells += f'<td{num_cls}>{fmt(v)}</td>'
        rows_html += f'<tr class="{cls}">{cells}</tr>'

    headers_html = ''.join(f'<th>{label}</th>' for label, _ in cols)
    return f'''
    <div class="eib-table-wrap">
      <table class="eib-table">
        <thead><tr>{headers_html}</tr></thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>'''

def to_excel(eib_rows, s):
    out = io.BytesIO()
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = 'Payroll Off cycle Payment'
    hdr = empty_row()
    for i, h in enumerate(EIB_HEADERS): hdr[i] = h
    ws.append(hdr)
    for i, row in enumerate(eib_rows):
        d = list(row['data'])
        n = i + 2
        d[C['paymentId']] = f'=CONCATENATE("Trailing2State",D{n},F{n})'
        ws.append(d)
    wb.save(out); out.seek(0)
    return out.getvalue()

# ══════════════════════════════════════════════════════════
#  UI
# ══════════════════════════════════════════════════════════

st.markdown("""
<div class="eib-app">
  <div class="eib-header">
    <div class="eib-logo">🧾 EIB Assistant</div>
    <div class="eib-sub">RSU Tax Formatter</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Step 1: Upload ─────────────────────────────────────────
st.markdown("""
<div class="eib-card">
  <div class="eib-step-badge">Step 1</div>
  <h2>Load Output File</h2>
  <p>Upload your output spreadsheet (the May 26 style file with the <strong>Original</strong> or <strong>Deloitte</strong> sheet).</p>
</div>
""", unsafe_allow_html=True)

with st.container():
    st.markdown('<div style="margin: -12px 24px 20px;">', unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload file", type=['xlsx','xls'], label_visibility='collapsed')
    st.markdown('</div>', unsafe_allow_html=True)

if not uploaded:
    st.stop()

wb_src = openpyxl.load_workbook(uploaded, data_only=True)
st.markdown(f'<div style="margin: -10px 24px 16px;"><span style="color:#188038;font-weight:600;font-size:0.9rem">✅ Loaded: {uploaded.name}</span></div>', unsafe_allow_html=True)

# ── Step 2: Select ─────────────────────────────────────────
preferred = ['Original','Deloitte','QUESTIONS_']
sheet_list = [s for s in preferred if s in wb_src.sheetnames] + \
             [s for s in wb_src.sheetnames if s not in preferred]

st.markdown("""
<div class="eib-card" style="margin-top:0">
  <div class="eib-step-badge">Step 2</div>
  <h2>Select Employee / Transaction</h2>
</div>
""", unsafe_allow_html=True)

with st.container():
    st.markdown('<div style="margin:-12px 24px 0">', unsafe_allow_html=True)
    c1,c2,c3,c4 = st.columns([1,1,1,0.6])
    with c1: sheet_name = st.selectbox("Source Sheet", sheet_list)
    ws_src = wb_src[sheet_name]
    parsed = parse_sheet(ws_src)
    if not parsed:
        st.error("Could not find expected columns. Try a different sheet.")
        st.stop()
    all_emps = sorted(set(r['empId'] for r in parsed))
    with c2: emp_id = st.selectbox("Employee ID", ['— choose —'] + all_emps)
    if emp_id == '— choose —':
        st.markdown('</div>', unsafe_allow_html=True)
        st.stop()
    all_txns = [r['txnId'] for r in parsed if r['empId'] == emp_id]
    with c3: txn_choice = st.selectbox("Transaction ID", ['All'] + all_txns)
    with c4:
        st.markdown('<div style="padding-top:24px">', unsafe_allow_html=True)
        load_clicked = st.button("Load Data", type="primary", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

recs = [r for r in parsed if r['empId']==emp_id and (txn_choice=='All' or r['txnId']==txn_choice)]

if load_clicked:
    # Source data preview
    preview = []
    for rec in recs:
        for sr in rec['rows']:
            preview.append({'Employee ID':rec['empId'],'Transaction ID':rec['txnId'],
                'Name':rec['name'],'State Code':sr['stateCode'],
                'Total Dist':sr['totalDist'],'EE Withholding':sr['totalWith'],
                'Region Tax':sr['regionTax'],'Income for Region':sr['incomeRegion']})
    rows_html = ''.join(
        f"<tr><td>{r['Employee ID']}</td><td>{r['Transaction ID']}</td><td>{r['Name']}</td>"
        f"<td><strong>{r['State Code']}</strong></td>"
        f"<td class='num'>{fmt(r['Total Dist'])}</td><td class='num'>{fmt(r['EE Withholding'])}</td>"
        f"<td class='num'>{fmt(r['Region Tax'])}</td><td class='num'>{fmt(r['Income for Region'])}</td></tr>"
        for r in preview)
    st.markdown(f"""
    <div class="eib-card" style="margin-top:0">
      <div class="eib-table-wrap">
        <table class="eib-table"><thead><tr>
          <th>Employee ID</th><th>Transaction ID</th><th>Name</th><th>State Code</th>
          <th>Total Dist</th><th>EE Withholding</th><th>Region Tax</th><th>Income for Region</th>
        </tr></thead><tbody>{rows_html}</tbody></table>
      </div>
    </div>""", unsafe_allow_html=True)
    st.session_state['recs_loaded'] = True
    st.session_state['recs'] = recs

if not st.session_state.get('recs_loaded'):
    st.stop()

recs = st.session_state.get('recs', recs)

# ── Step 3: Settings ───────────────────────────────────────
st.markdown("""
<div class="eib-card">
  <div class="eib-step-badge">Step 3</div>
  <h2>EIB Settings</h2>
  <p>These values are stamped on every row to match your EIB template.</p>
</div>
""", unsafe_allow_html=True)

with st.container():
    st.markdown('<div style="margin:-12px 24px 0">', unsafe_allow_html=True)
    s1,s2,s3,s4,s5,s6,s7 = st.columns(7)
    with s1: batch_id     = st.text_input("Batch ID",         value="RSU May 26")
    with s2: payment_date = st.text_input("Payment Date",     value="2026-05-01")
    with s3: period_date  = st.text_input("Period Date",      value="2026-05-01")
    with s4: pay_priority = st.text_input("Payment Priority", value="3")
    with s5: run_cat      = st.text_input("Run Category",     value="RSU_USA")
    with s6: result_type  = st.text_input("Result Type",      value="OnDemandPayment")
    with s7: reason       = st.text_input("Reason",           value="Stock")
    st.markdown('</div>', unsafe_allow_html=True)

settings = dict(batchId=batch_id, paymentDate=payment_date, periodDate=period_date,
                payPriority=pay_priority, runCategory=run_cat, resultType=result_type, reason=reason)

# ── Step 4: Output ─────────────────────────────────────────
eib_rows = gen_eib(recs, settings)
table_html = render_table(eib_rows)

st.markdown(f"""
<div class="eib-card">
  <div class="eib-step-badge">Step 4</div>
  <h2>EIB Formatted Output
    <span class="row-count-badge">{len(eib_rows)} EIB rows</span>
  </h2>
  <div class="eib-legend">
    <div class="eib-legend-item"><span class="eib-dot dot-tvrsu"></span> TVRSU – Total Distribution</div>
    <div class="eib-legend-item"><span class="eib-dot dot-nvrsu"></span> NVRSU – EE Withholding</div>
    <div class="eib-legend-item"><span class="eib-dot dot-wsww"></span> W_SWW – State Tax</div>
    <div class="eib-legend-item"><span class="eib-dot dot-sub"></span> Sub-calc lines</div>
  </div>
  {table_html}
</div>
""", unsafe_allow_html=True)

with st.container():
    st.markdown('<div style="margin:-12px 24px 20px;display:flex;gap:12px">', unsafe_allow_html=True)
    dl1, dl2, _ = st.columns([1.2, 1.6, 4])
    with dl1:
        st.download_button("⬇️ Export to Excel (EIB format)",
            data=to_excel(eib_rows, settings), file_name="eib_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ── Bulk ───────────────────────────────────────────────────
st.markdown("""
<div class="eib-card">
  <div class="eib-step-badge">Bulk</div>
  <h2>Bulk Process All Employees</h2>
  <p>Generate EIB rows for <strong>every</strong> employee in the selected sheet at once.</p>
</div>
""", unsafe_allow_html=True)

with st.container():
    st.markdown('<div style="margin:-12px 24px 20px">', unsafe_allow_html=True)
    if st.button("⚡ Generate All EIB Rows", type="secondary"):
        all_eib = gen_eib(parsed, settings)
        bulk_table = render_table(all_eib)
        st.markdown(f"""
        <div class="eib-card" style="margin-top:12px">
          <p style="color:#137333;font-weight:600;margin-bottom:12px">
            ✅ {len(all_eib)} EIB rows generated for {len(parsed)} transactions.
          </p>
          {bulk_table}
        </div>""", unsafe_allow_html=True)
        dl_b1, _ = st.columns([1.5, 5])
        with dl_b1:
            st.download_button("⬇️ Export All to Excel (EIB format)",
                data=to_excel(all_eib, settings), file_name="eib_bulk_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
