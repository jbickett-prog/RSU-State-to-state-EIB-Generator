
// ============================================================
//  EIB Assistant - RSU Tax Formatter
//  Output matches "Payroll Off cycle Payment" EIB column layout exactly.
//
//  EIB has 134 columns. Used columns (0-indexed):
//    [1]  Spreadsheet Key*      → source transaction seq number
//    [3]  Batch ID              → user-configurable
//    [4]  Payment ID*           → formula: ="Trailing2State"&[BatchID]&[EmpID]
//    [5]  Employee*             → Employee ID
//    [6]  Payment Date*         → user-configurable
//    [7]  Period Date*          → user-configurable
//    [8]  Payment Priority*     → 3
//    [9]  Run Category          → RSU_USA
//    [11] Result Type*          → OnDemandPayment
//    [13] Reason*               → Stock
//    [39] Row ID* (input line)  → 1, 2, 3, 4…
//    [40] Earning*              → TVRSU / NVRSU
//    [41] Deduction*            → W_SWW
//    [43] Amount                → value
//    [57] State Authority       → state tax code (W_SWW rows only)
//    [85] Row ID* (related calc)→ 1, 2, 3
//    [86] Related Calculation   → W_CWCGW / W_TTWG / W_TXWG
//    [87] Value                 → income for region
// ============================================================

const TOTAL_COLS = 134;

// EIB column indices (0-based)
const C = {
  spreadsheetKey:   1,
  batchId:          3,
  paymentId:        4,
  employee:         5,
  paymentDate:      6,
  periodDate:       7,
  payPriority:      8,
  runCategory:      9,
  resultType:       11,
  reason:           13,
  rowId:            39,
  earning:          40,
  deduction:        41,
  amount:           43,
  stateAuthority:   57,
  subRowId:         85,
  relatedCalc:      86,
  value:            87,
};

// EIB header rows (rows 1-5 of the template)
// We only include the columns that matter; rest are empty
const EIB_ROW1 = makeEmptyRow(); EIB_ROW1[0] = 'Payroll Off cycle Payment - v25.2';
const EIB_ROW2 = makeEmptyRow(); EIB_ROW2[0]='Area'; EIB_ROW2[1]='All'; EIB_ROW2[3]='Payroll Off cycle Payment Data'; EIB_ROW2[14]='On Demand Payment Data'; EIB_ROW2[19]='Tax Frequency Override+'; EIB_ROW2[22]='Manual Payment Data'; EIB_ROW2[26]='Result Worktag Overrides Data'; EIB_ROW2[39]='Off cycle Input Data+'; EIB_ROW2[49]='Payroll Worktags Data'; EIB_ROW2[85]='Input Line Data+'; EIB_ROW2[89]='Off cycle Input Data+';
const EIB_ROW3 = makeEmptyRow(); ['Required','Required','Optional','Optional','Required','Required','Required','Required','Required','Optional','Optional','Required','Optional','Required'].forEach((v,i)=>EIB_ROW3[i]=v);
const EIB_ROW4_HEADERS = ['Fields','Spreadsheet Key*','Payroll Off-cycle Payment','Batch ID','Payment ID*','Employee*','Payment Date*','Period Date*','Payment Priority*','Run Category','Pay Group','Result Type*','Replacement','Reason*','Use Supplemental Tax Rate','Override Payment to Check','Take Additional Withholding','Include Retro Differences in Payment','Load or Refresh Input','Row ID*','Tax Frequency Value','Tax Frequency Period','Third Party Sick Pay','Net Amount','Check Number','Bank Account','Company','Region','Location','Cost Center','Job Profile','State  Work ','State  Resident ','County  Work ','County  Resident ','City  Work ','City  Resident ','School District  Resident ','Payroll Reference Number','Row ID*','Earning*','Deduction*','Position','Amount','Hours','Rate','Adjustment','Reference Date','Currency','Location','Region','Job Profile','Cost Center','Project','Project Phase','Project Task','Withholding Order Case','State Authority','County Authority','City Authority','School District Authority','Custom Worktag 01','Custom Worktag 02','Custom Worktag 03','Custom Worktag 04','Custom Worktag 05','Fund','Grant','Gift','Program','Business Unit','Object Class','Custom Organization+','Custom Worktag 06','Custom Worktag 07','Custom Worktag 08','Custom Worktag 09','Custom Worktag 10','Custom Worktag 11','Custom Worktag 12','Custom Worktag 13','Custom Worktag 14','Custom Worktag 15','Local Other Tax Authority','NI Category','Row ID*','Related Calculation','Value','Company'];

function makeEmptyRow() { return Array(TOTAL_COLS).fill(null); }

// ── State abbreviation → Payroll Authority Tax Code map ──
// Format preserved exactly as-is (e.g. '01', '06', '36')
const STATE_ABB_TO_CODE = {
  'AL':'01','AK':'02','AZ':'04','AR':'05','CA':'06','CO':'08','CT':'09',
  'DE':'10','DC':'11','FL':'12','GA':'13','GU':'66','HI':'15','ID':'16',
  'IL':'17','IN':'18','IA':'19','KS':'20','KY':'21','LA':'22','ME':'23',
  'MD':'24','MA':'25','MI':'26','MN':'27','MS':'28','MO':'29','MT':'30',
  'NE':'31','NV':'32','NH':'33','NJ':'34','NM':'35','NY':'36','NC':'37',
  'ND':'38','OH':'39','OK':'40','OR':'41','PA':'42','PR':'72','RI':'44',
  'SC':'45','SD':'46','TN':'47','TX':'48','UT':'49','VT':'50','VA':'51',
  'VI':'78','WA':'53','WV':'54','WI':'55','WY':'56'
};

// Resolve a raw state value to its tax code.
// Accepts: abbreviation ('NY'), numeric code ('36' or 36), or already-padded code ('09').
function resolveStateCode(raw) {
  if (raw === null || raw === undefined || raw === '') return '';
  const s = String(raw).trim().toUpperCase();
  // If it's a known abbreviation, return the mapped code
  if (STATE_ABB_TO_CODE[s]) return STATE_ABB_TO_CODE[s];
  // Otherwise treat as a numeric code and preserve leading-zero padding
  const n = parseInt(s, 10);
  if (!isNaN(n)) return String(n).padStart(2, '0');
  return s;
}

// ── State ─────────────────────────────────────────────────
let workbook = null;
let parsedData = [];
let currentEibRows = [];
let allEibRows = [];

// ── Column finder ─────────────────────────────────────────
const COL_KEYS = {
  empId:        ['employee id'],
  txnId:        ['transaction id'],
  name:         ['name'],
  stateCode:    ['payroll authority tax code','state code','region code'],
  totalDist:    ['total distribution/exercise gross gain','reportable income'],
  totalWith:    ['total employee withholding (actual & hypo taxes)','country total actual tax ee liability'],
  regionTax:    ['region tax employee liability'],
  incomeRegion: ['income for region tax pre gross up'],
};

function findCol(headers, keys) {
  for (const key of keys) {
    const idx = headers.findIndex(h => h && h.toLowerCase().replace(/\n/g,' ').includes(key));
    if (idx >= 0) return idx;
  }
  return -1;
}
function num(v) { if (v===null||v===undefined||v==='') return null; const n=parseFloat(v); return isNaN(n)?null:n; }
function fmtNum(v) { if (v===null||v===undefined) return ''; return typeof v==='number'?v.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}):String(v); }

// ── File upload ────────────────────────────────────────────
const dropZone = document.getElementById('dropZone1');
const fileInput = document.getElementById('fileInput');
dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', e => { e.preventDefault(); dropZone.classList.remove('dragover'); handleFile(e.dataTransfer.files[0]); });
fileInput.addEventListener('change', () => handleFile(fileInput.files[0]));

function handleFile(file) {
  if (!file) return;
  document.getElementById('fileStatus').textContent = `Loading ${file.name}…`;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      workbook = XLSX.read(e.target.result, { type: 'binary' });
      document.getElementById('fileStatus').textContent = `✅ ${file.name}`;
      populateSheetSelect();
      document.getElementById('step2').classList.remove('hidden');
      document.getElementById('step3').classList.remove('hidden');
      document.getElementById('bulkBtn').classList.remove('hidden');
      document.getElementById('bulkBtn').disabled = false;
    } catch(err) {
      document.getElementById('fileStatus').textContent = `❌ ${err.message}`;
    }
  };
  reader.readAsBinaryString(file);
}

// ── Sheet selector ─────────────────────────────────────────
function populateSheetSelect() {
  const sel = document.getElementById('sheetSelect');
  sel.innerHTML = '';
  const preferred = ['Original','Deloitte','QUESTIONS_'];
  const sheets = workbook.SheetNames;
  const ordered = [...preferred.filter(s => sheets.includes(s)), ...sheets.filter(s => !preferred.includes(s))];
  ordered.forEach(name => { const o=document.createElement('option'); o.value=name; o.textContent=name; sel.appendChild(o); });
  sel.addEventListener('change', () => loadSheet(sel.value));
  loadSheet(ordered[0]);
}

function loadSheet(sheetName) {
  const ws = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
  parsedData = parseSheet(raw);
  populateEmpSelect();
}

// ── Parse sheet ───────────────────────────────────────────
function parseSheet(raw) {
  let headerRow = -1, headers = [];
  for (let i = 0; i < Math.min(5, raw.length); i++) {
    const row = raw[i] || [];
    const joined = row.filter(Boolean).join(' ').toLowerCase();
    if (joined.includes('employee') || joined.includes('transaction')) {
      headerRow = i; headers = row.map(c => c ? String(c) : ''); break;
    }
  }
  if (headerRow < 0) return [];

  const cEmpId  = findCol(headers, COL_KEYS.empId);
  const cTxnId  = findCol(headers, COL_KEYS.txnId);
  const cName   = findCol(headers, COL_KEYS.name);
  const cState  = findCol(headers, COL_KEYS.stateCode);
  const cTDist  = findCol(headers, COL_KEYS.totalDist);
  const cTWith  = findCol(headers, COL_KEYS.totalWith);
  const cRTax   = findCol(headers, COL_KEYS.regionTax);
  const cIncome = findCol(headers, COL_KEYS.incomeRegion);

  if (cEmpId < 0 || cTxnId < 0) return [];

  const groups = {}, order = [];
  let lastEmp = null, lastTxn = null;

  for (let i = headerRow + 1; i < raw.length; i++) {
    const row = raw[i] || [];
    let empId = row[cEmpId] ? String(row[cEmpId]).trim() : null;
    let txnId = row[cTxnId] ? String(row[cTxnId]).trim() : null;

    if (!empId && !txnId && !row.some(v => v !== null)) continue;
    if (!empId) empId = lastEmp;
    if (!txnId) txnId = lastTxn;
    if (!empId || !txnId) continue;

    lastEmp = empId; lastTxn = txnId;
    const key = `${empId}|${txnId}`;
    if (!groups[key]) { groups[key] = { empId, txnId, name: '', rows: [] }; order.push(key); }
    if (cName >= 0 && row[cName] && !groups[key].name) groups[key].name = String(row[cName]);

    const sc = cState >= 0 ? row[cState] : null;
    groups[key].rows.push({
      stateCode:    resolveStateCode(sc),
      totalDist:    num(cTDist  >= 0 ? row[cTDist]  : null),
      totalWith:    num(cTWith  >= 0 ? row[cTWith]  : null),
      regionTax:    num(cRTax   >= 0 ? row[cRTax]   : null),
      incomeRegion: num(cIncome >= 0 ? row[cIncome] : null),
    });
  }
  return order.map(k => groups[k]);
}

// ── Employee / Transaction selectors ─────────────────────
function populateEmpSelect() {
  const empSel = document.getElementById('empSelect');
  empSel.innerHTML = '<option value="">— choose —</option>';
  [...new Set(parsedData.map(d => d.empId))].sort().forEach(e => {
    const o = document.createElement('option'); o.value = e; o.textContent = e; empSel.appendChild(o);
  });
  empSel.onchange = () => populateTxnSelect(empSel.value);
  document.getElementById('txnSelect').innerHTML = '<option value="">— all —</option>';
}

function populateTxnSelect(empId) {
  const txnSel = document.getElementById('txnSelect');
  txnSel.innerHTML = '<option value="">— all —</option>';
  if (!empId) return;
  parsedData.filter(d => d.empId === empId).forEach(d => {
    const o = document.createElement('option'); o.value = d.txnId; o.textContent = d.txnId; txnSel.appendChild(o);
  });
}

document.getElementById('loadBtn').addEventListener('click', () => {
  const empId = document.getElementById('empSelect').value;
  const txnId = document.getElementById('txnSelect').value;
  if (!empId) { alert('Please select an Employee ID.'); return; }
  const records = parsedData.filter(d => d.empId === empId && (!txnId || d.txnId === txnId));
  if (!records.length) { alert('No data found.'); return; }
  renderInputPreview(records);
  currentEibRows = generateEibRows(records);
  renderOutputTable(currentEibRows, 'outputPreview');
  document.getElementById('rowCount').textContent = `${currentEibRows.length} EIB rows`;
  document.getElementById('step4').classList.remove('hidden');
  document.getElementById('step4').scrollIntoView({ behavior:'smooth', block:'start' });
});

function renderInputPreview(records) {
  let html = '<table><thead><tr><th>Emp ID</th><th>Txn ID</th><th>Name</th><th>State</th>'
    + '<th>Total Dist</th><th>EE Withholding</th><th>Region Tax</th><th>Income for Region</th></tr></thead><tbody>';
  records.forEach(rec => {
    rec.rows.forEach(r => {
      html += `<tr><td>${rec.empId}</td><td>${rec.txnId}</td><td>${rec.name}</td>
        <td><strong>${r.stateCode}</strong></td>
        <td class="num-cell">${fmtNum(r.totalDist)}</td>
        <td class="num-cell">${fmtNum(r.totalWith)}</td>
        <td class="num-cell">${fmtNum(r.regionTax)}</td>
        <td class="num-cell">${fmtNum(r.incomeRegion)}</td></tr>`;
    });
  });
  html += '</tbody></table>';
  document.getElementById('inputPreview').innerHTML = html;
}

// ── EIB Row Generation ────────────────────────────────────
// Universal pattern (matches actual EIB exactly):
// For each transaction:
//   Row 1: TVRSU earning, amount = totalDist (from primary state row)
//   Row 2: NVRSU earning, amount = totalWith (from primary state row)
//   For each state row (all rows):
//     Row N:   W_SWW deduction, amount = regionTax, stateAuth = stateCode,
//              subRowId=1, relCalc=W_CWCGW, value=incomeRegion
//     Row N:   (blank deduction), subRowId=2, relCalc=W_TTWG, value=incomeRegion
//     Row N:   (blank deduction), subRowId=3, relCalc=W_TXWG, value=incomeRegion
//
// The "primary" state row = the one with a non-null totalDist value.

function getSettings() {
  return {
    batchId:      document.getElementById('batchId').value.trim()      || 'RSU May 26',
    paymentDate:  document.getElementById('paymentDate').value.trim()  || '2026-05-01',
    periodDate:   document.getElementById('periodDate').value.trim()   || '2026-05-01',
    payPriority:  document.getElementById('payPriority').value.trim()  || '3',
    runCategory:  document.getElementById('runCategory').value.trim()  || 'RSU_USA',
    resultType:   document.getElementById('resultType').value.trim()   || 'OnDemandPayment',
    reason:       document.getElementById('reason').value.trim()       || 'Stock',
  };
}

function buildBaseRow(s, spreadsheetKey, empId) {
  const r = makeEmptyRow();
  r[C.spreadsheetKey] = spreadsheetKey;
  r[C.batchId]        = s.batchId;
  r[C.paymentId]      = `="Trailing2State"&D&F`;  // formula placeholder — displayed as text; actual formula set on export
  r[C.employee]       = empId;
  r[C.paymentDate]    = s.paymentDate;
  r[C.periodDate]     = s.periodDate;
  r[C.payPriority]    = Number(s.payPriority);
  r[C.runCategory]    = s.runCategory;
  r[C.resultType]     = s.resultType;
  r[C.reason]         = s.reason;
  return r;
}

function generateEibRows(records) {
  const s = getSettings();
  const rows = [];     // each entry = { meta, data[] } for display; for export = flat 134-col array

  // flat EIB data rows only (no header)
  const eibFlat = [];

  records.forEach(rec => {
    const primary = rec.rows.find(r => r.totalDist !== null) || rec.rows[0];
    const spreadsheetKey = rec.txnId;  // The "Spreadsheet Key" = transaction seq number from source
    let lineNum = 1;

    // TVRSU row
    const tvRow = buildBaseRow(s, spreadsheetKey, rec.empId);
    tvRow[C.rowId]   = lineNum++;
    tvRow[C.earning] = 'TVRSU';
    tvRow[C.amount]  = primary.totalDist;
    eibFlat.push({ type:'TVRSU', empId:rec.empId, txnId:rec.txnId, data: tvRow });

    // NVRSU row
    const nvRow = buildBaseRow(s, spreadsheetKey, rec.empId);
    nvRow[C.rowId]   = lineNum++;
    nvRow[C.earning] = 'NVRSU';
    nvRow[C.amount]  = primary.totalWith;
    eibFlat.push({ type:'NVRSU', empId:rec.empId, txnId:rec.txnId, data: nvRow });

    // W_SWW + sub-calc rows (3 per state)
    // Skip entirely if regionTax and incomeRegion are both 0 (nothing to report)
    rec.rows.forEach(sr => {
      const allZero = (sr.regionTax === 0 || sr.regionTax === null) &&
                      (sr.incomeRegion === 0 || sr.incomeRegion === null);
      if (allZero) return;

      const ln = lineNum++;

      // W_SWW main row (sub-line 1: W_CWCGW)
      const wRow = buildBaseRow(s, spreadsheetKey, rec.empId);
      wRow[C.rowId]         = ln;
      wRow[C.deduction]     = 'W_SWW';
      wRow[C.amount]        = sr.regionTax;
      wRow[C.stateAuthority]= sr.stateCode;
      wRow[C.subRowId]      = 1;
      wRow[C.relatedCalc]   = 'W_CWCGW';
      wRow[C.value]         = sr.incomeRegion;
      eibFlat.push({ type:'W_SWW', empId:rec.empId, txnId:rec.txnId, stateCode:sr.stateCode, data: wRow });

      // W_TTWG sub-line
      const t2Row = buildBaseRow(s, spreadsheetKey, rec.empId);
      t2Row[C.rowId]      = ln;
      t2Row[C.subRowId]   = 2;
      t2Row[C.relatedCalc]= 'W_TTWG';
      t2Row[C.value]      = sr.incomeRegion;
      eibFlat.push({ type:'sub', subCode:'W_TTWG', empId:rec.empId, txnId:rec.txnId, data: t2Row });

      // W_TXWG sub-line
      const t3Row = buildBaseRow(s, spreadsheetKey, rec.empId);
      t3Row[C.rowId]      = ln;
      t3Row[C.subRowId]   = 3;
      t3Row[C.relatedCalc]= 'W_TXWG';
      t3Row[C.value]      = sr.incomeRegion;
      eibFlat.push({ type:'sub', subCode:'W_TXWG', empId:rec.empId, txnId:rec.txnId, data: t3Row });
    });
  });

  return eibFlat;
}

// ── Render the preview table (show only the populated columns) ──
// Display columns for readability:
const DISPLAY_COLS = [
  { idx: C.spreadsheetKey, label: 'Spr. Key'   },
  { idx: C.batchId,        label: 'Batch ID'   },
  { idx: C.paymentId,      label: 'Payment ID' },
  { idx: C.employee,       label: 'Employee'   },
  { idx: C.paymentDate,    label: 'Pay Date'   },
  { idx: C.periodDate,     label: 'Period Date'},
  { idx: C.payPriority,    label: 'Priority'   },
  { idx: C.runCategory,    label: 'Run Cat.'   },
  { idx: C.resultType,     label: 'Result Type'},
  { idx: C.reason,         label: 'Reason'     },
  { idx: C.rowId,          label: 'Row ID'     },
  { idx: C.earning,        label: 'Earning'    },
  { idx: C.deduction,      label: 'Deduction'  },
  { idx: C.amount,         label: 'Amount'     },
  { idx: C.stateAuthority, label: 'State Auth' },
  { idx: C.subRowId,       label: 'Sub Row ID' },
  { idx: C.relatedCalc,    label: 'Rel. Calc'  },
  { idx: C.value,          label: 'Value'      },
];

function renderOutputTable(rows, containerId) {
  let lastTxn = null;
  let html = '<table><thead><tr>'
    + DISPLAY_COLS.map(c => `<th>${c.label}</th>`).join('')
    + '</tr></thead><tbody>';

  rows.forEach(row => {
    const cls = { TVRSU:'row-tvrsu', NVRSU:'row-nvrsu', W_SWW:'row-wsww', sub:'row-sub' }[row.type] || '';
    if (row.txnId !== lastTxn && lastTxn !== null) {
      html += `<tr class="row-divider"><td colspan="${DISPLAY_COLS.length}" style="padding:0;height:3px;"></td></tr>`;
    }
    lastTxn = row.txnId;

    html += `<tr class="${cls}">`;
    DISPLAY_COLS.forEach(col => {
      const v = row.data[col.idx];
      const isNum = typeof v === 'number';
      html += `<td${isNum ? ' class="num-cell"' : ''}>${v !== null && v !== undefined ? (isNum ? fmtNum(v) : v) : ''}</td>`;
    });
    html += '</tr>';
  });

  html += '</tbody></table>';
  document.getElementById(containerId).innerHTML = html;
}

// ── Copy to Clipboard (tab-separated, all 134 cols) ──────
document.getElementById('copyBtn').addEventListener('click', () => {
  const tsv = currentEibRows.map(row => row.data.map(v => v !== null && v !== undefined ? v : '').join('\t')).join('\n');
  navigator.clipboard.writeText(tsv).then(() => {
    const btn = document.getElementById('copyBtn');
    const orig = btn.textContent;
    btn.textContent = '✅ Copied!';
    setTimeout(() => btn.textContent = orig, 2000);
  });
});

// ── Export to Excel (EIB format) ─────────────────────────
document.getElementById('exportBtn').addEventListener('click', () => {
  exportEib(currentEibRows, 'eib_output.xlsx');
  document.getElementById('exportMsg').classList.remove('hidden');
  setTimeout(() => document.getElementById('exportMsg').classList.add('hidden'), 4000);
});

function exportEib(rows, filename) {
  const s = getSettings();

  // Build the header rows exactly like the EIB template
  const headerRow5 = makeEmptyRow();
  EIB_ROW4_HEADERS.forEach((h, i) => { headerRow5[i] = h; });

  const aoa = [headerRow5]; // just one header row for clean pasting

  rows.forEach((row, i) => {
    // Clone and inject the real Payment ID formula
    const d = [...row.data];
    const excelRow = i + 2; // data starts at row 2 if header is row 1
    const colD = colLetter(C.batchId + 1);   // D
    const colF = colLetter(C.employee + 1);  // F
    d[C.paymentId] = `="Trailing2State"&${colD}${excelRow}&${colF}${excelRow}`;
    aoa.push(d);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Set column widths for the populated columns
  const wscols = Array(TOTAL_COLS).fill({ wch: 3 });
  DISPLAY_COLS.forEach(c => { wscols[c.idx] = { wch: Math.max(c.label.length + 2, 12) }; });
  ws['!cols'] = wscols;

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Payroll Off cycle Payment');
  XLSX.writeFile(wb, filename);
}

function colLetter(n) {
  let s = '';
  while (n > 0) { let r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ── Bulk ─────────────────────────────────────────────────
document.getElementById('bulkBtn').addEventListener('click', () => {
  if (!parsedData.length) return;
  allEibRows = generateEibRows(parsedData);
  renderOutputTable(allEibRows, 'bulkPreview');
  const nonSub = allEibRows.length;
  document.getElementById('bulkPreview').insertAdjacentHTML('afterbegin',
    `<p style="margin-bottom:12px;color:#137333;font-weight:600;">✅ ${nonSub} EIB rows generated for ${parsedData.length} transactions.</p>`);
  document.getElementById('bulkExportRow').classList.remove('hidden');
  document.getElementById('step5').scrollIntoView({ behavior:'smooth', block:'start' });
});

document.getElementById('bulkExportBtn').addEventListener('click', () => {
  exportEib(allEibRows, 'eib_bulk_output.xlsx');
});
