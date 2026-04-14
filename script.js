'use strict';

/* ══════════════════════════════════════════════════════════════════
   BRANCH COLOR MAP
══════════════════════════════════════════════════════════════════ */
const BRANCH_COLORS = ['#2563eb', '#059669', '#d97706', '#dc2626', '#7c3aed', '#ea580c', '#0d9488', '#be185d', '#0891b2', '#65a30d'];
let branchColorMap = {};
function getBranchColor(branch) {
  if (!branch) return '#8b93a5';
  if (!branchColorMap[branch]) {
    const idx = Object.keys(branchColorMap).length % BRANCH_COLORS.length;
    branchColorMap[branch] = BRANCH_COLORS[idx];
  }
  return branchColorMap[branch];
}

/* ══════════════════════════════════════════════════════════════════
   STATE
══════════════════════════════════════════════════════════════════ */
let rawParsed = [];   // After parseSheet() — pre-clean item rows
let allData = [];   // After cleaning pipeline
let filtered = [];   // After applyFilters()
let errorLog = [];
let currentSort = { col: 'date', dir: 'asc' };
let cancelledCount = 0;
let loadedWorkbooks = []; // Array of { name: string, wb: Workbook, sheets: string[] }
let selectedSheets = new Set(); // Set of unique keys: "filename » sheetname"
let outputMode = 'merge';  // 'merge' | 'separate' | 'both'
let dashFilter = null;
let detectedColMap = {};  // { uniqueKey: { colMap } }
let rawHeaders = {};      // { uniqueKey: [headers] }
let sheetRowCounts = {};  // { uniqueKey: number }
let cleanStats = { dupes: 0, grossFixed: 0, grossWrong: 0, missingFields: 0 };

/* ══════════════════════════════════════════════════════════════════
   NAVIGATION
══════════════════════════════════════════════════════════════════ */
function switchPage(id) {
  // Hide all pages
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));

  const navEl = document.getElementById('nav-' + id);
  if (navEl) navEl.classList.add('active');

  if (id === 'upload') {
    document.getElementById('page-upload').style.display = 'flex';
    document.getElementById('topbarTitle').textContent = 'Upload File';
    document.getElementById('exportBtn').style.display = 'none';
    document.getElementById('recordPill').style.display = 'none';
  } else {
    document.getElementById('page-upload').style.display = 'none';
    const page = document.getElementById('page-' + id);
    if (page) page.classList.add('active');

    const titles = {
      configure: 'Configure', dashboard: 'Dashboard', preview: 'Preview',
      records: 'Records', errors: 'Error Log', analytics: 'Analytics', clean: 'Clean Log'
    };
    document.getElementById('topbarTitle').textContent = titles[id] || id;
    document.getElementById('exportBtn').style.display = allData.length ? 'flex' : 'none';
    document.getElementById('recordPill').style.display = allData.length ? 'flex' : 'none';
    document.getElementById('recordPill').textContent = allData.length + ' records';

    if (id === 'analytics') renderAnalytics();
    if (id === 'records') applyFilters();
    if (id === 'errors') renderErrorLog();
    if (id === 'dashboard') updateDashboard();
  }

  // Close mobile sidebar
  if (window.innerWidth <= 768) {
    document.querySelector('.sidebar').classList.remove('open');
    document.getElementById('sidebarOverlay').classList.remove('open');
  }
}

switchPage('upload');

/* ══════════════════════════════════════════════════════════════════
   MOBILE SIDEBAR
══════════════════════════════════════════════════════════════════ */
function toggleSidebar() {
  document.querySelector('.sidebar').classList.toggle('open');
  document.getElementById('sidebarOverlay').classList.toggle('open');
}

/* ══════════════════════════════════════════════════════════════════
   FILE HANDLING
══════════════════════════════════════════════════════════════════ */
const dropZone = document.getElementById('dropZone');
if (dropZone) {
  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault(); dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length) handleFiles({ target: { files: e.dataTransfer.files } });
  });
}

/**
 * handleFiles() — entry point for multiple files
 */
async function handleFiles(e) {
  const files = Array.from(e.target.files);
  if (!files.length) return;

  setProgress(true, `Loading ${files.length} file(s)…`, 10);

  for (let i = 0; i < files.length; i++) {
    const pct = 10 + Math.round((i / files.length) * 80);
    setProgress(true, `Reading ${files[i].name}…`, pct);
    await processFile(files[i]);
  }

  setProgress(false);
  e.target.value = ''; // reset
  renderSheetSelector();
  renderLoadedFiles();
  updateUIForLoadedState();
  showToast(`✅ Loaded ${files.length} file(s). Select sheets to process.`);
}

function handleFile(e) {
  // Legacy single file support
  if (e.target.files[0]) handleFiles(e);
}

/**
 * parseFile() — reads the uploaded Excel with FileReader
 */
function parseFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
        resolve(wb);
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error('FileReader failed'));
    reader.readAsBinaryString(file);
  });
}

/**
 * processFile() — core logic to parse one file and add to state
 */
async function processFile(file) {
  try {
    const wb = await parseFile(file);
    const sheets = wb.SheetNames;
    
    // Check if file already loaded to avoid duplicates
    if (loadedWorkbooks.some(w => w.name === file.name)) return;

    loadedWorkbooks.push({ name: file.name, wb, sheets });

    // Pre-scan sheets
    sheets.forEach(sheetName => {
      const uKey = `${file.name} » ${sheetName}`;
      const ws = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
      rawHeaders[uKey] = rows[0] || [];
      const { records, colMap } = parseSheetWithMap(rows, sheetName, uKey);
      sheetRowCounts[uKey] = records.length;
      detectedColMap[uKey] = colMap;
      
      // Auto-select by default if it has records
      if (records.length > 0) selectedSheets.add(uKey);
    });

    const bp = document.getElementById('btnProcess');
    if (bp) bp.disabled = false;
    updateColumnDetectionUI();
  } catch (err) {
    console.error(err);
    showToast('❌ Error reading ' + file.name + ': ' + err.message);
  }
}

function updateUIForLoadedState() {
  const hasData = loadedWorkbooks.length > 0;
  document.getElementById('dropZone').style.display = hasData ? 'none' : 'block';
  document.getElementById('loadedFilesPanel').style.display = hasData ? 'block' : 'none';
  document.getElementById('uploadControlsTop').style.display = hasData ? 'flex' : 'none';
  
  if (hasData) {
    const totalSheets = loadedWorkbooks.reduce((acc, w) => acc + w.sheets.length, 0);
    document.getElementById('fileNameDisplay').textContent = `${loadedWorkbooks.length} files loaded`;
    document.getElementById('fileStats').textContent = `${totalSheets} total sheets detected`;
  } else {
    document.getElementById('fileNameDisplay').textContent = 'No file loaded';
    document.getElementById('fileStats').textContent = 'Upload an Excel file to begin';
  }
}

function renderLoadedFiles() {
  const list = document.getElementById('loadedFilesList');
  if (!list) return;
  list.innerHTML = loadedWorkbooks.map((w, idx) => `
    <div class="file-item">
      <div class="file-icon">📊</div>
      <div class="file-name" title="${esc(w.name)}">${esc(w.name)}</div>
      <div class="file-sheets">${w.sheets.length} sheets</div>
      <button class="btn-remove-file" onclick="removeFile(${idx})" title="Remove file">✕</button>
    </div>
  `).join('');
}

function removeFile(idx) {
  const file = loadedWorkbooks[idx];
  if (!file) return;
  
  // Remove associated sheets from selection
  selectedSheets.forEach(key => {
    if (key.startsWith(file.name + ' » ')) selectedSheets.delete(key);
  });
  
  // Clean up maps
  file.sheets.forEach(s => {
    const uKey = `${file.name} » ${s}`;
    delete rawHeaders[uKey];
    delete sheetRowCounts[uKey];
    delete detectedColMap[uKey];
  });

  loadedWorkbooks.splice(idx, 1);
  renderLoadedFiles();
  renderSheetSelector();
  updateUIForLoadedState();
  updateColumnDetectionUI();
  if (loadedWorkbooks.length === 0) clearAllData();
}

function clearAllData() {
  loadedWorkbooks = [];
  selectedSheets.clear();
  rawParsed = [];
  allData = [];
  filtered = [];
  errorLog = [];
  detectedColMap = {};
  rawHeaders = {};
  sheetRowCounts = {};
  cleanStats = { dupes: 0, grossFixed: 0, grossWrong: 0, missingFields: 0 };
  
  updateUIForLoadedState();
  renderSheetSelector();
  renderLoadedFiles();
  
  const bp = document.getElementById('btnProcess');
  if (bp) bp.disabled = true;
  
  document.getElementById('page-upload').style.display = 'flex';
  document.getElementById('sheetSelector').style.display = 'none';
  document.getElementById('btnNextConfigure').style.display = 'none';
  
  showToast('🧹 All data cleared.');
}

function setProgress(show, label = '', pct = 0) {
  const wrap = document.getElementById('progressWrap');
  if (!wrap) return;
  wrap.style.display = show ? 'block' : 'none';
  document.getElementById('progressLabel').textContent = label;
  const inner = document.getElementById('progressInner');
  inner.style.width = pct + '%';
}

/* ══════════════════════════════════════════════════════════════════
   SHEET SELECTOR
══════════════════════════════════════════════════════════════════ */
function renderSheetSelector() {
  const panel = document.getElementById('sheetSelector');
  const list = document.getElementById('sheetList');
  const ssCount = document.getElementById('ssCount');
  if (!panel || !list) return;

  const allSheetKeys = [];
  loadedWorkbooks.forEach(w => {
    w.sheets.forEach(s => allSheetKeys.push(`${w.name} » ${s}`));
  });

  if (allSheetKeys.length === 0) {
    panel.style.display = 'none';
    return;
  }

  panel.style.display = 'block';
  document.getElementById('btnNextConfigure').style.display = '';

  if (ssCount) ssCount.textContent = allSheetKeys.length + ' sheets total';

  let html = '';
  loadedWorkbooks.forEach(w => {
    html += `<div class="sheet-group-label" style="font-size:11px;font-weight:700;margin:12px 0 6px;color:var(--muted)">📁 ${esc(w.name)}</div>`;
    w.sheets.forEach(name => {
      const uKey = `${w.name} » ${name}`;
      const c = getBranchColor(name);
      const rows = sheetRowCounts[uKey] || 0;
      const chk = selectedSheets.has(uKey);
      const safeId = 'ss-' + uKey.replace(/[^a-zA-Z0-9]/g, '_');
      html += `<label class="sheet-item${chk ? '' : ' disabled'}" id="si-${safeId}" for="${safeId}">
        <input type="checkbox" id="${safeId}" ${chk ? 'checked' : ''} onchange="toggleSheet('${esc(uKey)}', this.checked)">
        <div class="sheet-item-dot" style="background:${c}"></div>
        <div class="sheet-item-name">${esc(name)}</div>
        <span class="sheet-item-rows">${rows} rows</span>
      </label>`;
    });
  });

  list.innerHTML = html;
  updateSheetMergeInfo();
}

function toggleSheet(uKey, checked) {
  if (checked) selectedSheets.add(uKey);
  else selectedSheets.delete(uKey);
  const safeId = 'ss-' + uKey.replace(/[^a-zA-Z0-9]/g, '_');
  const el = document.getElementById('si-' + safeId);
  if (el) el.classList.toggle('disabled', !checked);
  updateSheetMergeInfo();
}

function sheetSelectAll() {
  const allSheetKeys = [];
  loadedWorkbooks.forEach(w => {
    w.sheets.forEach(s => {
      const uKey = `${w.name} » ${s}`;
      selectedSheets.add(uKey);
      const safeId = 'ss-' + uKey.replace(/[^a-zA-Z0-9]/g, '_');
      const cb = document.getElementById(safeId);
      if (cb) cb.checked = true;
      const el = document.getElementById('si-' + safeId);
      if (el) el.classList.remove('disabled');
    });
  });
  updateSheetMergeInfo();
}

function sheetDeselectAll() {
  selectedSheets.clear();
  loadedWorkbooks.forEach(w => {
    w.sheets.forEach(s => {
      const uKey = `${w.name} » ${s}`;
      const safeId = 'ss-' + uKey.replace(/[^a-zA-Z0-9]/g, '_');
      const cb = document.getElementById(safeId);
      if (cb) cb.checked = false;
      const el = document.getElementById('si-' + safeId);
      if (el) el.classList.add('disabled');
    });
  });
  updateSheetMergeInfo();
}

function sheetInvert() {
  loadedWorkbooks.forEach(w => {
    w.sheets.forEach(s => {
      const uKey = `${w.name} » ${s}`;
      if (selectedSheets.has(uKey)) selectedSheets.delete(uKey);
      else selectedSheets.add(uKey);
      
      const safeId = 'ss-' + uKey.replace(/[^a-zA-Z0-9]/g, '_');
      const cb = document.getElementById(safeId);
      if (cb) cb.checked = selectedSheets.has(uKey);
      const el = document.getElementById('si-' + safeId);
      if (el) el.classList.toggle('disabled', !selectedSheets.has(uKey));
    });
  });
  updateSheetMergeInfo();
}

function updateSheetMergeInfo() {
  const info = document.getElementById('sheetMergeInfo');
  if (!info) return;
  const count = selectedSheets.size;
  const total = [...selectedSheets].reduce((s, uKey) => s + (sheetRowCounts[uKey] || 0), 0);
  if (count === 0) {
    info.textContent = '⚠️ No sheets selected — select at least one sheet to proceed.';
    info.style.color = 'var(--red)';
  } else {
    const fileCount = new Set([...selectedSheets].map(k => k.split(' » ')[0])).size;
    info.textContent = `✅ ${count} sheets from ${fileCount} files selected · ~${total} rows will be processed`;
    info.style.color = 'var(--green)';
  }
}

/* ══════════════════════════════════════════════════════════════════
   OUTPUT MODE
══════════════════════════════════════════════════════════════════ */
function setOutputMode(mode) {
  outputMode = mode;
  ['merge', 'separate', 'both'].forEach(m => {
    const el = document.getElementById('om-' + m);
    if (el) el.classList.toggle('selected', m === mode);
  });
  // Show "add branch column" toggle only for merge/both modes
  const bw = document.getElementById('addBranchToggleWrap');
  if (bw) bw.style.display = (mode === 'separate') ? 'none' : '';
}

/* ══════════════════════════════════════════════════════════════════
   COLUMN DETECTION — detectColumns(rows)
   Inspects header row for known column names.
   Returns a mapping: { date, customer, creditNote, narration, qty, rate, gross, value }
══════════════════════════════════════════════════════════════════ */
function detectColumns(rows) {
  const colMap = { date: -1, customer: -1, creditNote: -1, narration: -1, qty: -1, rate: -1, gross: -1, value: -1, vtype: -1, voucher: -1 };
  let headerIdx = -1;

  // Scan first 25 rows for a header containing 'date' AND 'voucher'
  for (let i = 0; i < Math.min(25, rows.length); i++) {
    const low = rows[i].map(c => String(c || '').trim().toLowerCase());
    if (low.some(c => c === 'date') && low.some(c => c.includes('voucher'))) {
      headerIdx = i;
      low.forEach((h, idx) => {
        if (h === 'date') colMap.date = idx;
        if (h === 'particulars' || h === 'party name' || h === 'customer') colMap.customer = idx;
        if (h === 'voucher type') colMap.vtype = idx;
        if (h.includes('voucher no') || h.includes('voucher number')) colMap.voucher = idx;
        if (h === 'narration' || h === 'description') colMap.narration = idx;
        if (h === 'quantity' || h === 'qty') colMap.qty = idx;
        if (h === 'rate' || h === 'unit rate') colMap.rate = idx;
        if (h === 'value' || h === 'net rate' || h === 'net value') colMap.value = idx;
        if (h.includes('gross') || h === 'amount') colMap.gross = idx;
      });
      break;
    }
  }

  // Fallback: detect CN/ pattern to locate voucher column and derive header row
  if (headerIdx < 0) {
    outer: for (let i = 0; i < Math.min(35, rows.length); i++) {
      for (let c = 0; c < (rows[i]?.length || 0); c++) {
        if (/^(CN|CNG|CNS)\//i.test(String(rows[i][c] || '').trim())) {
          colMap.voucher = c;
          headerIdx = Math.max(0, i - 1);
          break outer;
        }
      }
    }
  }

  return { colMap, headerIdx };
}

/**
 * mapColumns() — show manual mapping UI if auto-detect fails for critical columns
 */
function updateColumnDetectionUI() {
  const statusEl = document.getElementById('colDetectionStatus');
  const mapperEl = document.getElementById('columnMapper');
  if (!statusEl || !mapperEl) return;

  const activeSheetKeys = [...selectedSheets];
  if (!activeSheetKeys.length) {
    statusEl.textContent = '⚠️ No sheets selected.';
    statusEl.className = 'col-detection-status warning';
    mapperEl.style.display = 'none';
    return;
  }

  let allDetected = true;
  const mapperItems = [];

  activeSheetKeys.forEach(uKey => {
    const cm = detectedColMap[uKey] || {};
    const missing = [];
    if (cm.voucher < 0) missing.push('Credit Note Number');
    if (cm.qty < 0) missing.push('Quantity');
    if (cm.rate < 0) missing.push('Unit Rate');
    if (missing.length) allDetected = false;

    const headers = rawHeaders[uKey] || [];
    const optionsHtml = headers.map((h, i) => `<option value="${i}">${esc(String(h || '(col ' + (i + 1) + ')'))}</option>`).join('');

    const fields = [
      { key: 'date', label: 'Date', val: cm.date },
      { key: 'customer', label: 'Customer', val: cm.customer },
      { key: 'voucher', label: 'Credit Note Number', val: cm.voucher },
      { key: 'narration', label: 'Narration', val: cm.narration },
      { key: 'qty', label: 'Quantity', val: cm.qty },
      { key: 'rate', label: 'Unit Rate', val: cm.rate },
      { key: 'gross', label: 'Gross', val: cm.gross },
    ];

    const [fName, sName] = uKey.split(' » ');

    mapperItems.push(`<div class="mapper-sheet-group">
      <div class="mapper-sheet-label" style="color:${getBranchColor(sName)}">▸ ${esc(fName)} / ${esc(sName)}</div>
      <div class="mapper-fields-grid">
        ${fields.map(f => {
      const detected = f.val >= 0;
      return `<div class="mapper-item">
            <label>${f.label}</label>
            ${detected
          ? `<div style="font-size:12px;font-weight:500;color:var(--ink);margin-top:2px">${esc(String(rawHeaders[uKey]?.[f.val] || '—'))}</div>
                 <div class="auto-detected">✓ Auto-detected (col ${f.val + 1})</div>`
          : `<select class="col-manual-select" data-ukey="${esc(uKey)}" data-field="${f.key}" onchange="manualMapColumn(this)">
                   <option value="-1">— Not detected —</option>${optionsHtml}
                 </select>`
        }
          </div>`;
    }).join('')}
      </div>
    </div>`);
  });

  if (allDetected) {
    statusEl.innerHTML = '<span class="col-status-icon">✅</span> All required columns auto-detected successfully.';
    statusEl.className = 'col-detection-status success';
  } else {
    statusEl.innerHTML = '<span class="col-status-icon">⚠️</span> Some columns were not detected. Please map them manually below.';
    statusEl.className = 'col-detection-status warning';
  }

  mapperEl.innerHTML = `<style>
    .mapper-sheet-group { grid-column: 1 / -1; margin-bottom: 8px; border-bottom: 1px solid var(--line2); padding-bottom: 12px; }
    .mapper-sheet-label { font-size: 11px; font-weight: 700; margin-bottom: 10px; }
    .mapper-fields-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); gap: 8px; }
  </style>` + mapperItems.join('');
  mapperEl.style.display = 'block';
}

/**
 * manualMapColumn() — called when user selects a column from dropdown
 */
function manualMapColumn(select) {
  const uKey = select.dataset.ukey;
  const field = select.dataset.field;
  const colIdx = parseInt(select.value, 10);
  if (!detectedColMap[uKey]) detectedColMap[uKey] = {};
  detectedColMap[uKey][field] = colIdx;
  if (field === 'voucher') detectedColMap[uKey].creditNote = colIdx;
}

/* ══════════════════════════════════════════════════════════════════
   SHEET PARSER — parseSheetWithMap(rows, branchName)
   Returns { records, colMap }
══════════════════════════════════════════════════════════════════ */
function parseSheetWithMap(rows, branchName, uKey) {
  const { colMap, headerIdx } = detectColumns(rows);
  const records = parseSheetRows(rows, branchName, colMap, headerIdx);
  return { records, colMap: { ...colMap, _headerIdx: headerIdx } };
}

/**
 * parseSheetRows() — extract flat item-level records from a sheet
 */
function parseSheetRows(rows, branchName, colMap, headerIdx) {
  const CN_PATTERN = /^(CN|CNG|CNS|CR)\//i;
  const CANCELLED = /\(cancelled/i;
  const records = [];

  // Numeric coercion helper
  const toNum = v => {
    if (v === null || v === undefined || String(v).trim() === '') return 0;
    const s = String(v).replace(/[^0-9.-]/g, '');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  };

  // Date formatter
  const fmtDate = v => {
    if (!v) return '';
    if (v instanceof Date) {
      const d = v.getDate(), m = v.getMonth() + 1, y = v.getFullYear();
      return `${String(d).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}`;
    }
    return String(v).trim();
  };

  let i = headerIdx >= 0 ? headerIdx + 1 : 0;
  let currentVoucher = '';
  let currentCustomer = '';
  let currentDate = '';
  let currentNarration = '';

  while (i < rows.length) {
    const r = rows[i];
    if (!r || r.length === 0) { i++; continue; }

    const voucherVal = String(r[colMap.voucher] || '').trim();
    const dateVal = colMap.date >= 0 ? fmtDate(r[colMap.date]) : '';
    const custVal = colMap.customer >= 0 ? String(r[colMap.customer] || '').trim() : '';
    const narrVal = colMap.narration >= 0 ? String(r[colMap.narration] || '').trim() : '';
    const qtyVal = toNum(r[colMap.qty]);
    const rateVal = toNum(r[colMap.rate]);
    const valueVal = toNum(r[colMap.value >= 0 ? colMap.value : -1]);
    const grossVal = toNum(r[colMap.gross >= 0 ? colMap.gross : -1]);

    // Skip Grand Total and separator rows
    const rowText = r.map(c => String(c || '').trim().toLowerCase()).join(' ');
    if (rowText.includes('grand total') || rowText.includes('grandtotal')) { i++; continue; }
    if (r.every(c => String(c || '').trim() === '')) { i++; continue; }

    // Skip cancelled vouchers
    if (CANCELLED.test(voucherVal) || CANCELLED.test(narrVal)) {
      cancelledCount++;
      i++; continue;
    }

    // When a new voucher number appears, update context
    if (voucherVal && CN_PATTERN.test(voucherVal)) {
      currentVoucher = voucherVal;
      currentDate = dateVal || currentDate;
      currentCustomer = custVal || currentCustomer;
      currentNarration = narrVal || currentNarration;
    }

    // Emit record if there's a meaningful item line (has qty or gross or rate)
    if (currentVoucher && (qtyVal !== 0 || rateVal !== 0 || grossVal !== 0 || valueVal !== 0)) {
      records.push({
        date: currentDate,
        customer: currentCustomer || custVal,
        creditNote: currentVoucher,
        narration: narrVal || currentNarration,
        qty: qtyVal,
        unitRate: rateVal,
        netRate: valueVal,
        gross: grossVal || valueVal,
        branch: branchName,
        _flags: [],
      });
    } else if (!currentVoucher && custVal && dateVal) {
      // Non-CN row with content — emit if it has a value
      if (grossVal || valueVal) {
        records.push({
          date: dateVal,
          customer: custVal,
          creditNote: voucherVal,
          narration: narrVal,
          qty: qtyVal,
          unitRate: rateVal,
          netRate: valueVal,
          gross: grossVal || valueVal,
          branch: branchName,
          _flags: [],
        });
      }
    }

    i++;
  }

  return records;
}

/* ══════════════════════════════════════════════════════════════════
   BUILD RAW PARSED — merges selected sheets
══════════════════════════════════════════════════════════════════ */
function buildRawParsed() {
  cancelledCount = 0;
  branchColorMap = {};
  const allRecords = [];

  const activeSheetKeys = [...selectedSheets];
  activeSheetKeys.forEach(uKey => {
    const [fileName, sheetName] = uKey.split(' » ');
    const workbookObj = loadedWorkbooks.find(w => w.name === fileName);
    if (!workbookObj) return;

    getBranchColor(sheetName); // assign color
    const ws = workbookObj.wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
    const cm = detectedColMap[uKey] || {};
    const records = parseSheetRows(rows, sheetName, cm, cm._headerIdx ?? -1);
    allRecords.push(...records);
  });

  rawParsed = allRecords.map(r => ({ ...r, _flags: [] }));
}

/* ══════════════════════════════════════════════════════════════════
   PROCESS DATA — main entry point from "Process & Clean" button
══════════════════════════════════════════════════════════════════ */
function processData() {
  if (loadedWorkbooks.length === 0) { showToast('⚠️ Please upload at least one file first.'); return; }
  if (selectedSheets.size === 0) { showToast('⚠️ No sheets selected — go to Upload and select at least one.'); return; }

  const btn = document.getElementById('btnProcess');
  const status = document.getElementById('processStatus');
  if (btn) { btn.disabled = true; btn.innerHTML = '⏳ Processing…'; }
  if (status) status.textContent = '';

  showToast(`🔄 Merging ${selectedSheets.size} sheet(s)…`);

  setTimeout(() => {
    try {
      buildRawParsed();

      if (!rawParsed.length) {
        if (btn) { btn.disabled = false; btn.innerHTML = '🔄 Process & Clean'; }
        showToast('⚠️ No records found in selected sheets.');
        return;
      }

      runCleaningPipeline();

      // Update UI
      renderRawPreview();
      renderCleanedPreview();
      renderSummaryCards();
      updateCleanLog();
      populateFilters();

      // Update nav badges
      const nbR = document.getElementById('nb-records');
      const nbE = document.getElementById('nb-errors');
      if (nbR) nbR.textContent = allData.length;
      if (nbE) nbE.textContent = errorLog.length;

      if (btn) { btn.disabled = false; btn.innerHTML = '✅ Processed'; }
      const reBtn = document.getElementById('btnReprocess');
      if (reBtn) reBtn.style.display = '';
      if (status) status.textContent = `✅ ${allData.length} rows cleaned · ${cleanStats.dupes} dupes removed · ${cleanStats.grossFixed} gross fixed`;

      showToast(`✅ Done! ${allData.length} rows · ${errorLog.length} issues found.`);

      // Switch to preview
      switchPage('preview');
      document.getElementById('summaryCards').style.display = '';
      document.getElementById('previewActions').style.display = '';
    } catch (err) {
      if (btn) { btn.disabled = false; btn.innerHTML = '🔄 Process & Clean'; }
      console.error(err);
      showToast('❌ Error during processing: ' + err.message);
    }
  }, 80);
}

/* ══════════════════════════════════════════════════════════════════
   CLEANING PIPELINE — runCleaningPipeline()
══════════════════════════════════════════════════════════════════ */
function runCleaningPipeline() {
  errorLog = [];
  cleanStats = { dupes: 0, grossFixed: 0, grossWrong: 0, missingFields: 0 };

  let data = rawParsed.map(r => ({ ...r, _flags: [] }));

  // Read toggle states
  const doTrim = document.getElementById('tog-trim')?.checked !== false;
  const doGross = document.getElementById('tog-gross')?.checked !== false;
  const doDupes = document.getElementById('tog-dupes')?.checked !== false;
  const doValidate = document.getElementById('tog-validate')?.checked !== false;
  const doSort = document.getElementById('tog-sort')?.checked !== false;
  const doExtract = document.getElementById('tog-extract')?.checked === true;

  if (doTrim) data = trimSpaces(data);
  if (doGross) data = calculateGross(data);
  if (doValidate) data = validateData(data);
  if (doDupes) data = detectDuplicates(data);
  if (doExtract) data = extractProductFields(data);
  if (doSort) data.sort((a, b) => dp(a.date) - dp(b.date));

  allData = data;
  filtered = [...allData];
}

/* ── Cleaning Modules ───────────────────────────────────────────── */

/** trimSpaces() — strip leading/trailing whitespace */
function trimSpaces(data) {
  return data.map(r => ({
    ...r,
    customer: r.customer.trim(),
    creditNote: r.creditNote.trim(),
    narration: r.narration.trim(),
    branch: r.branch.trim(),
    date: r.date.trim(),
  }));
}

/**
 * calculateGross() — Fix Gross = Qty × Unit Rate
 * Allows adjustment rows where qty or rate = 0
 */
function calculateGross(data) {
  return data.map(r => {
    const row = { ...r };
    // Only validate when both qty and rate are present and non-zero
    if (row.qty > 0 && row.unitRate > 0) {
      const expected = Math.round(row.qty * row.unitRate * 100) / 100;
      if (!row.gross || row.gross === 0) {
        // Missing gross — compute it
        row._grossOriginal = 0;
        row.gross = expected;
        row._flags = [...row._flags, 'gross-computed'];
        cleanStats.grossFixed++;
        pushError('info', row.creditNote, row.customer, row.date, row.branch,
          `Gross was missing — computed as Qty×Rate = ₹${fmt2(expected)}`);
      } else {
        const diff = Math.abs(row.gross - expected);
        if (diff > 1) {
          // Incorrect gross — correct it
          row._grossOriginal = row.gross;
          row.gross = expected;
          row._flags = [...row._flags, 'gross-corrected'];
          cleanStats.grossFixed++;
          cleanStats.grossWrong++;
          pushError('warn', row.creditNote, row.customer, row.date, row.branch,
            `Gross was ₹${fmt2(row._grossOriginal)} but Qty×Rate = ₹${fmt2(expected)} — corrected`);
        }
      }
    }
    return row;
  });
}

/**
 * validateData() — flag rows missing required fields
 */
function validateData(data) {
  return data.map(r => {
    const row = { ...r };
    const missing = [];
    if (!row.date) missing.push('Date');
    if (!row.customer) missing.push('Customer');
    if (!row.creditNote) missing.push('Credit Note Number');
    if (missing.length) {
      row._flags = [...row._flags, 'error'];
      cleanStats.missingFields++;
      pushError('err', row.creditNote || '?', row.customer || '?', row.date || '?', row.branch,
        `Missing required field(s): ${missing.join(', ')}`);
    }
    return row;
  });
}

/**
 * detectDuplicates() — remove exact duplicates on CN + Narration + Qty + Rate
 * (matches spec exactly — NOT customer/branch based to avoid cross-branch false positives)
 */
function detectDuplicates(data) {
  const seen = new Set();
  const output = [];
  data.forEach(r => {
    // Key per spec: Credit Note Number + Narration + Quantity + Unit Rate
    const key = [r.creditNote.toLowerCase(), r.narration.toLowerCase(), r.qty, r.unitRate].join('||');
    if (seen.has(key)) {
      cleanStats.dupes++;
      pushError('warn', r.creditNote, r.customer, r.date, r.branch,
        'Duplicate row removed (same CN + Narration + Qty + Rate)');
    } else {
      seen.add(key);
      output.push(r);
    }
  });
  return output;
}

/**
 * extractProductFields() — optional: parse product name/code from narration
 */
function extractProductFields(data) {
  return data.map(r => {
    const row = { ...r };
    // Simple heuristic: extract codes like "ABC-123" or "P/12345" from narration
    const codeMatch = row.narration.match(/\b([A-Z]{1,5}[-\/]\d{3,})\b/);
    if (codeMatch) {
      row._productCode = codeMatch[1];
      row._flags = [...row._flags, 'product-extracted'];
    }
    return row;
  });
}

/* ══════════════════════════════════════════════════════════════════
   SUMMARY CARDS
══════════════════════════════════════════════════════════════════ */
function renderSummaryCards() {
  const el = id => document.getElementById(id);
  if (el('sum-total')) el('sum-total').textContent = allData.length.toLocaleString('en-IN');
  if (el('sum-dupes')) el('sum-dupes').textContent = cleanStats.dupes;
  if (el('sum-gross')) el('sum-gross').textContent = cleanStats.grossFixed;
  if (el('sum-errors')) el('sum-errors').textContent = errorLog.length;
}

/* ══════════════════════════════════════════════════════════════════
   generateSummary() — returns an object of summary stats
══════════════════════════════════════════════════════════════════ */
function generateSummary() {
  const totGross = allData.reduce((s, d) => s + d.gross, 0);
  const unique = new Set(allData.map(d => d.customer.trim().toLowerCase())).size;
  const uniqueV = new Set(allData.map(d => d.creditNote)).size;
  const branches = [...new Set(allData.map(d => d.branch))];
  return { total: allData.length, totGross, unique, uniqueV, branches, ...cleanStats };
}

/* ══════════════════════════════════════════════════════════════════
   PREVIEW RENDERING — generatePreview()
══════════════════════════════════════════════════════════════════ */

/**
 * generatePreview(data, limit) — renders first N rows into an HTML string
 */
function generatePreview(data, limit = 20) {
  const headers = ['Date', 'Customer', 'Credit Note No.', 'Narration', 'Qty', 'Unit Rate', 'Net Rate', 'Gross', 'Branch', 'Status'];
  let html = `<thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>`;

  data.slice(0, limit).forEach(d => {
    // Highlight logic
    let rowStyle = '';
    if (d._flags.includes('gross-corrected')) rowStyle = 'style="background:#fffbeb"';
    else if (d._flags.includes('error')) rowStyle = 'style="background:#fff5f5"';

    let statusBadge = '<span class="preview-badge" style="background:var(--green-light);color:var(--green)">OK</span>';
    if (d._flags.includes('error')) statusBadge = '<span class="preview-badge" style="background:var(--red-light);color:var(--red)">Error</span>';
    else if (d._flags.includes('gross-corrected')) statusBadge = '<span class="preview-badge" style="background:#fefce8;color:#854d0e">⚠ Gross Fixed</span>';
    else if (d._flags.includes('gross-computed')) statusBadge = '<span class="preview-badge" style="background:var(--green-light);color:var(--green)">➕ Gross Added</span>';

    const bc = getBranchColor(d.branch);
    // Yellow highlight for corrected gross cell
    const grossCell = d._flags.includes('gross-corrected')
      ? `<td style="text-align:right;font-family:'DM Mono',monospace;font-weight:600;background:#fef08a">${fmt2(d.gross)}<span title="Was: ₹${fmt2(d._grossOriginal || 0)}" style="cursor:help;margin-left:4px">✏️</span></td>`
      : `<td style="text-align:right;font-family:'DM Mono',monospace;font-weight:600">${d.gross ? fmt2(d.gross) : '—'}</td>`;

    html += `<tr ${rowStyle}>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(d.date)}</td>
      <td style="max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${esc(d.customer)}">${esc(d.customer)}</td>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(d.creditNote)}</td>
      <td style="max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${esc(d.narration)}">${esc(d.narration) || '—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${d.qty || '—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${d.unitRate ? fmt2(d.unitRate) : '—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${d.netRate ? fmt2(d.netRate) : '—'}</td>
      ${grossCell}
      <td><span class="badge badge-branch" style="background:${bc}">${esc(d.branch)}</span></td>
      <td>${statusBadge}</td>
    </tr>`;
  });

  if (data.length > limit) {
    html += `<tr><td colspan="10" style="text-align:center;color:var(--muted);padding:12px;font-size:12px">…and ${data.length - limit} more rows (showing first ${limit})</td></tr>`;
  }
  html += '</tbody>';
  return html;
}

function renderCleanedPreview() {
  const table = document.getElementById('previewCleanedTable');
  if (!table) return;
  if (!allData.length) {
    table.innerHTML = '<thead><tr><th>No cleaned data yet — press Process</th></tr></thead><tbody></tbody>';
    return;
  }
  table.innerHTML = generatePreview(allData, 20);
}

function renderRawPreview() {
  const table = document.getElementById('previewRawTable');
  if (!table) return;
  if (!rawParsed.length) {
    table.innerHTML = '<thead><tr><th>No raw data yet</th></tr></thead><tbody></tbody>';
    return;
  }
  // Raw preview uses simplified columns
  const headers = ['Date', 'Customer', 'Credit Note No.', 'Narration', 'Qty', 'Unit Rate', 'Net Rate', 'Gross', 'Branch'];
  let html = `<thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>`;
  rawParsed.slice(0, 20).forEach(d => {
    const bc = getBranchColor(d.branch);
    html += `<tr>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(d.date)}</td>
      <td style="max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(d.customer)}</td>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(d.creditNote)}</td>
      <td style="max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(d.narration) || '—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${d.qty || '—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${d.unitRate ? fmt2(d.unitRate) : '—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${d.netRate ? fmt2(d.netRate) : '—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:600">${d.gross ? fmt2(d.gross) : '—'}</td>
      <td><span class="badge badge-branch" style="background:${bc}">${esc(d.branch)}</span></td>
    </tr>`;
  });
  if (rawParsed.length > 20) html += `<tr><td colspan="9" style="text-align:center;color:var(--muted);padding:12px;font-size:12px">…and ${rawParsed.length - 20} more rows</td></tr>`;
  html += '</tbody>';
  table.innerHTML = html;
}

function switchPreviewTab(btn, tab) {
  document.querySelectorAll('#page-preview .tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('#page-preview .tab-pane').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('preview-' + tab).classList.add('active');
}

/* ══════════════════════════════════════════════════════════════════
   ERROR LOG
══════════════════════════════════════════════════════════════════ */
function pushError(sev, voucher, customer, date, branch, msg) {
  errorLog.push({ sev, voucher, customer, date, branch, msg });
}

function renderErrorLog() {
  const list = document.getElementById('errorList');
  const tbody = document.getElementById('errorTableBody');
  const card = document.getElementById('errorTableCard');
  const count = document.getElementById('errorTotalCount');
  if (count) count.textContent = errorLog.length + ' issue' + (errorLog.length !== 1 ? 's' : '');

  if (!errorLog.length) {
    if (list) list.innerHTML = '<div class="no-errors">✅ No data issues detected.</div>';
    if (card) card.style.display = 'none';
    return;
  }

  if (list) {
    list.innerHTML = errorLog.slice(0, 50).map(e => `
      <div class="error-item">
        <span class="error-sev sev-${e.sev}">${e.sev.toUpperCase()}</span>
        <div>
          <div class="error-msg">${esc(e.msg)}</div>
          <div class="error-ref">${esc(e.voucher)} · ${esc(e.customer)} · ${esc(e.branch)}</div>
        </div>
      </div>`).join('') +
      (errorLog.length > 50 ? `<div style="padding:10px 16px;font-size:12px;color:var(--muted)">…and ${errorLog.length - 50} more issues</div>` : '');
  }

  if (card) card.style.display = '';
  if (tbody) {
    tbody.innerHTML = errorLog.map(e => `<tr>
      <td><span class="error-sev sev-${e.sev}">${e.sev.toUpperCase()}</span></td>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(e.voucher)}</td>
      <td>${esc(e.customer)}</td>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(e.date)}</td>
      <td><span class="badge badge-branch" style="background:${getBranchColor(e.branch)}">${esc(e.branch)}</span></td>
      <td style="font-size:12px">${esc(e.msg)}</td>
    </tr>`).join('');
  }
}

/* ══════════════════════════════════════════════════════════════════
   FILTERS & SORT
══════════════════════════════════════════════════════════════════ */
function populateFilters() {
  const branches = [...new Set(allData.map(d => d.branch))].sort();
  const fb = document.getElementById('filterBranch');
  if (fb) fb.innerHTML = '<option value="">All branches</option>' + branches.map(b => `<option>${esc(b)}</option>`).join('');
}

function applyFilters() {
  const q = (document.getElementById('searchBox')?.value || '').toLowerCase();
  const br = document.getElementById('filterBranch')?.value || '';
  const status = document.getElementById('filterStatus')?.value || '';
  const sort = document.getElementById('filterSort')?.value || 'date-asc';

  filtered = allData.filter(d => {
    const mq = !q || (d.customer.toLowerCase().includes(q) || d.creditNote.toLowerCase().includes(q) || d.narration.toLowerCase().includes(q) || d.branch.toLowerCase().includes(q));
    const mb = !br || d.branch === br;
    let ms = true;
    if (status === 'clean') ms = !d._flags.length;
    if (status === 'grossfixed') ms = d._flags.includes('gross-computed') || d._flags.includes('gross-corrected');
    if (status === 'error') ms = d._flags.includes('error');
    return mq && mb && ms;
  });

  const [col, dir] = sort.split('-');
  currentSort = { col: col || 'date', dir: dir || 'asc' };
  sortFiltered();
  renderTable();
}

function clearFilters() {
  const sb = document.getElementById('searchBox');
  const fb = document.getElementById('filterBranch');
  const fs = document.getElementById('filterStatus');
  const fso = document.getElementById('filterSort');
  if (sb) sb.value = '';
  if (fb) fb.value = '';
  if (fs) fs.value = '';
  if (fso) fso.value = 'date-asc';
  applyFilters();
}

function sortBy(col) {
  if (currentSort.col === col) currentSort.dir = currentSort.dir === 'asc' ? 'desc' : 'asc';
  else currentSort = { col, dir: 'asc' };
  const fs = document.getElementById('filterSort');
  if (fs) fs.value = '';
  sortFiltered();
  updateSortIcons();
  renderTable();
}

function sortFiltered() {
  const { col, dir } = currentSort;
  const m = dir === 'asc' ? 1 : -1;
  filtered.sort((a, b) => {
    if (col === 'date') return m * (dp(a.date) - dp(b.date));
    if (col === 'customer') return m * a.customer.localeCompare(b.customer);
    if (col === 'branch') return m * a.branch.localeCompare(b.branch);
    return m * ((a[col] || 0) - (b[col] || 0));
  });
}

function updateSortIcons() {
  ['date', 'customer', 'gross', 'branch'].forEach(c => {
    const el = document.getElementById('sa-' + c);
    if (!el) return;
    el.className = 'sort-arrows' + (currentSort.col === c ? ' ' + currentSort.dir : '');
  });
}

/* ══════════════════════════════════════════════════════════════════
   TABLE RENDER
══════════════════════════════════════════════════════════════════ */
function renderTable() {
  const tbody = document.getElementById('tableBody');
  if (!tbody) return;
  const fc = document.getElementById('filterCount');
  if (fc) fc.textContent = filtered.length + ' rows';

  if (!filtered.length) {
    tbody.innerHTML = `<tr><td colspan="10"><div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-title">No records match</div></div></td></tr>`;
    const pi = document.getElementById('pageInfo');
    if (pi) pi.textContent = '0 rows';
    return;
  }

  const totGross = filtered.reduce((s, d) => s + d.gross, 0);
  let html = '';

  filtered.forEach((d, i) => {
    let rowClass = '';
    if (d._flags.includes('error')) rowClass = 'row-error';
    else if (d._flags.includes('gross-corrected')) rowClass = 'row-gross-wrong';
    else if (d._flags.includes('gross-computed')) rowClass = 'row-gross-fixed';

    const narr = d.narration.length > 35 ? d.narration.slice(0, 32) + '…' : d.narration;
    const bc = getBranchColor(d.branch);

    let statusBadges = '';
    if (!d._flags.length) statusBadges = '<span class="badge badge-ok">OK</span>';
    else {
      if (d._flags.includes('error')) statusBadges += '<span class="badge badge-error">Error</span> ';
      if (d._flags.includes('gross-computed')) statusBadges += '<span class="badge badge-fixed">Gross+</span> ';
      if (d._flags.includes('gross-corrected')) statusBadges += '<span class="badge badge-gross-wrong">Gross✏️</span> ';
    }

    html += `<tr class="${rowClass}" onclick="showDetail(${i})" style="cursor:pointer">
      <td class="cell-date">${esc(d.date)}</td>
      <td class="cell-customer"><div class="cell-customer-inner" title="${esc(d.customer)}">${esc(d.customer)}</div></td>
      <td class="cell-voucher">${esc(d.creditNote)}</td>
      <td class="cell-narr" title="${esc(d.narration)}">${esc(narr) || '—'}</td>
      <td class="cell-num">${fmtQty(d.qty)}</td>
      <td class="cell-num">${d.unitRate > 0 ? fmt(d.unitRate) : '—'}</td>
      <td class="cell-num">${d.netRate > 0 ? fmt(d.netRate) : '—'}</td>
      <td class="cell-num cell-gross">${fmt(d.gross)}${d._grossOriginal ? `<span title="Was: ₹${fmt2(d._grossOriginal)}" style="cursor:help;margin-left:3px;color:var(--amber)">✏️</span>` : ''}</td>
      <td><span class="badge badge-branch" style="background:${bc}">${esc(d.branch)}</span></td>
      <td style="white-space:nowrap">${statusBadges}</td>
    </tr>`;
  });

  const uv = new Set(filtered.map(d => d.creditNote)).size;
  html += `<tr class="totals-row">
    <td colspan="7">Total — ${filtered.length} rows / ${uv} credit notes</td>
    <td class="cell-num">${fmt(totGross)}</td>
    <td colspan="2"></td>
  </tr>`;

  tbody.innerHTML = html;
  const pi = document.getElementById('pageInfo');
  if (pi) pi.textContent = `${filtered.length} rows · ${uv} CNs (of ${allData.length} total)`;
  const ft = document.getElementById('filteredTotal');
  if (ft) ft.textContent = filtered.length < allData.length ? 'Filtered gross: ' + fmtK(totGross) : '';
}

/* ══════════════════════════════════════════════════════════════════
   DETAIL MODAL
══════════════════════════════════════════════════════════════════ */
function showDetail(idx) {
  const d = filtered[idx];
  if (!d) return;
  let flagHtml = '';
  if (d._flags.length) {
    const msgs = [];
    if (d._flags.includes('error')) msgs.push('❌ Missing required field(s)');
    if (d._flags.includes('gross-computed')) msgs.push('🔧 Gross was missing — computed as Qty × Rate');
    if (d._flags.includes('gross-corrected')) msgs.push(`✏️ Gross corrected from ₹${fmt2(d._grossOriginal || 0)} → ₹${fmt2(d.gross)}`);
    flagHtml = `<div class="flag-box">${msgs.join('<br>')}</div>`;
  }
  const bc = getBranchColor(d.branch);
  document.getElementById('modalVoucher').textContent = d.creditNote;
  document.getElementById('modalDate').textContent = d.date + ' · ' + d.branch;
  document.getElementById('modalFlagBox').innerHTML = flagHtml;
  document.getElementById('detailGrid').innerHTML = `
    <div><div class="detail-label">Customer</div><div class="detail-value">${esc(d.customer)}</div></div>
    <div><div class="detail-label">Branch</div><div class="detail-value"><span class="badge badge-branch" style="background:${bc}">${esc(d.branch)}</span></div></div>
    <div><div class="detail-label">Quantity</div><div class="detail-value mono">${fmtQty(d.qty) || '—'}</div></div>
    <div><div class="detail-label">Unit Rate</div><div class="detail-value mono">${d.unitRate > 0 ? fmt(d.unitRate) : '—'}</div></div>
    <div><div class="detail-label">Net Rate</div><div class="detail-value mono">${fmt(d.netRate)}</div></div>
    <div><div class="detail-label">Gross</div><div class="detail-value big">${fmt(d.gross)}</div></div>`;
  document.getElementById('narrationSection').innerHTML = d.narration
    ? `<div style="margin-bottom:12px"><div class="detail-label" style="margin-bottom:5px">Narration</div><div class="narr-box">${esc(d.narration)}</div></div>` : '';
  document.getElementById('detailOverlay').classList.add('open');
}

function closeModal() {
  document.getElementById('detailOverlay').classList.remove('open');
}

document.addEventListener('keydown', e => {
  if (e.key === 'Escape') {
    closeModal();
    closeDownloadModal();
  }
});

/* ══════════════════════════════════════════════════════════════════
   DOWNLOAD CONFIRMATION MODAL
══════════════════════════════════════════════════════════════════ */
function confirmAndDownload() {
  if (!allData.length) { showToast('⚠️ No data to export — process a file first.'); return; }

  const summary = generateSummary();
  const modal = document.getElementById('downloadOverlay');
  const fileList = document.getElementById('dlFileList');
  const dlSum = document.getElementById('dlSummary');
  const dlSub = document.getElementById('dlModalSub');

  const addBranch = document.getElementById('addBranchCol')?.checked !== false;

  // Build file list based on output mode
  let files = [];
  if (outputMode === 'merge' || outputMode === 'both') {
    files.push({ icon: '📊', name: 'Merged_Cleaned.xlsx', desc: `${allData.length} rows · all ${summary.branches.length} branches merged` + (addBranch ? ' · Branch column included' : '') });
  }
  if (outputMode === 'separate' || outputMode === 'both') {
    const activeSheetKeys = [...selectedSheets];
    activeSheetKeys.forEach(uKey => {
      const [fName, sName] = uKey.split(' » ');
      const cnt = allData.filter(d => d.branch === sName).length;
      files.push({ icon: '📋', name: `${sName}_cleaned.xlsx`, desc: `${cnt} rows · from ${fName}` });
    });
  }
  if (errorLog.length) {
    files.push({ icon: '⚠️', name: 'Error_Log (included in each file)', desc: `${errorLog.length} issues logged` });
  }

  fileList.innerHTML = files.map(f => `
    <div class="dl-file-item">
      <div class="dl-file-icon">${f.icon}</div>
      <div>
        <div class="dl-file-name">${esc(f.name)}</div>
        <div class="dl-file-desc">${esc(f.desc)}</div>
      </div>
    </div>`).join('');

  dlSum.innerHTML = `
    ✅ <strong>${allData.length}</strong> rows ready to export<br>
    🚫 <strong>${cleanStats.dupes}</strong> duplicates removed<br>
    🔧 <strong>${cleanStats.grossFixed}</strong> gross values fixed<br>
    ⚠️ <strong>${errorLog.length}</strong> issues logged
  `;

  const modeLabel = { merge: 'Merge into One File', separate: 'Separate Files', both: 'Both (Merged + Separate)' };
  dlSub.textContent = 'Mode: ' + (modeLabel[outputMode] || outputMode);

  modal.classList.add('open');
}

function closeDownloadModal() {
  document.getElementById('downloadOverlay').classList.remove('open');
}

/* ══════════════════════════════════════════════════════════════════
   EXPORT — exportExcel()
   Handles Merge / Separate / Both modes
══════════════════════════════════════════════════════════════════ */
function executeDownload() {
  closeDownloadModal();
  exportExcel();
}

function exportExcel() {
  if (!allData.length) { showToast('⚠️ No data to export'); return; }

  const addBranch = document.getElementById('addBranchCol')?.checked !== false;

  if (outputMode === 'merge' || outputMode === 'both') {
    exportMerged(addBranch);
  }
  if (outputMode === 'separate' || outputMode === 'both') {
    exportSeparate();
  }

  showToast(`📥 Downloading in "${outputMode}" mode…`);
}

/**
 * exportMerged() — one combined file: Merged_Cleaned.xlsx
 */
function exportMerged(includeBranch = true) {
  const headers = ['Date', 'Customer', 'Credit Note Number', 'Narration', 'Quantity', 'Unit Rate', 'Net Rate', 'Gross'];
  if (includeBranch) headers.push('Branch');
  headers.push('Status Flags');

  const wsData = [
    headers,
    ...allData.map(d => {
      const row = [d.date, d.customer, d.creditNote, d.narration, d.qty, d.unitRate, d.netRate, d.gross];
      if (includeBranch) row.push(d.branch);
      row.push(d._flags.join(', '));
      return row;
    })
  ];

  const tot = allData.reduce((s, d) => ({ netRate: s.netRate + d.netRate, gross: s.gross + d.gross }), { netRate: 0, gross: 0 });
  const totalRow = ['TOTAL', '', '', '', '', '', Math.round(tot.netRate * 100) / 100, Math.round(tot.gross * 100) / 100];
  if (includeBranch) totalRow.push('');
  totalRow.push('');
  wsData.push(totalRow);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  const colWidths = [12, 42, 18, 50, 10, 12, 14, 16];
  if (includeBranch) colWidths.push(14);
  colWidths.push(18);
  ws['!cols'] = colWidths.map(w => ({ wch: w }));

  // Apply yellow highlight to gross-corrected rows using cell styles
  XLSX.utils.book_append_sheet(wb, ws, 'Final_Data');
  appendErrorSheet(wb);

  const fn = `Merged_Cleaned_${today()}.xlsx`;
  XLSX.writeFile(wb, fn);
}

/**
 * exportSeparate() — one file per sheet: BranchName_cleaned.xlsx
 */
function exportSeparate() {
  const activeSheetKeys = [...selectedSheets];
  activeSheetKeys.forEach(uKey => {
    const [fileName, sheetName] = uKey.split(' » ');
    const branchData = allData.filter(d => d.branch === sheetName);
    if (!branchData.length) return;

    const wsData = [
      ['Date', 'Customer', 'Credit Note Number', 'Narration', 'Quantity', 'Unit Rate', 'Net Rate', 'Gross', 'Status Flags'],
      ...branchData.map(d => [d.date, d.customer, d.creditNote, d.narration, d.qty, d.unitRate, d.netRate, d.gross, d._flags.join(', ')])
    ];

    const tot = branchData.reduce((s, d) => ({ netRate: s.netRate + d.netRate, gross: s.gross + d.gross }), { netRate: 0, gross: 0 });
    wsData.push(['TOTAL', '', '', '', '', '', Math.round(tot.netRate * 100) / 100, Math.round(tot.gross * 100) / 100, '']);

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!cols'] = [12, 42, 18, 50, 10, 12, 14, 16, 18].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));

    // Include relevant errors for this branch
    const branchErrors = errorLog.filter(e => e.branch === sheetName);
    if (branchErrors.length) {
      const wsE = XLSX.utils.aoa_to_sheet([
        ['Severity', 'Credit Note No.', 'Customer', 'Date', 'Issue'],
        ...branchErrors.map(e => [e.sev.toUpperCase(), e.voucher, e.customer, e.date, e.msg])
      ]);
      wsE['!cols'] = [10, 18, 38, 12, 60].map(w => ({ wch: w }));
      XLSX.utils.book_append_sheet(wb, wsE, 'Error_Log');
    }

    const safeName = sheetName.replace(/[\/\\*?[\]:]/g, '_');
    XLSX.writeFile(wb, `${safeName}_cleaned_${today()}.xlsx`);
  });
}

/**
 * appendErrorSheet() — adds Error_Log sheet to a workbook
 */
function appendErrorSheet(wb) {
  if (!errorLog.length) return;
  const wsE = XLSX.utils.aoa_to_sheet([
    ['Severity', 'Credit Note No.', 'Customer', 'Date', 'Branch', 'Issue'],
    ...errorLog.map(e => [e.sev.toUpperCase(), e.voucher, e.customer, e.date, e.branch, e.msg])
  ]);
  wsE['!cols'] = [10, 18, 38, 12, 14, 60].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsE, 'Error_Log');
}

/* ══════════════════════════════════════════════════════════════════
   DASHBOARD
══════════════════════════════════════════════════════════════════ */
function getDashFiltered() {
  if (!dashFilter) return allData;
  return allData.filter(d => {
    if (dashFilter.type === 'branch') return d.branch === dashFilter.value;
    if (dashFilter.type === 'month') {
      const p = d.date.split('/');
      return p.length >= 3 && (p[1] + '/' + p[2]) === dashFilter.value;
    }
    if (dashFilter.type === 'customer') return d.customer === dashFilter.value;
    return true;
  });
}

function updateDashboard() {
  if (!allData.length) return;
  const src = getDashFiltered();
  const totGross = src.reduce((s, d) => s + d.gross, 0);
  const unique = new Set(src.map(d => d.customer.trim().toLowerCase())).size;
  const uniqueV = new Set(src.map(d => d.creditNote)).size;
  const cleanRows = src.filter(d => !d._flags.length).length;

  const el = id => document.getElementById(id);
  if (el('s-count')) el('s-count').textContent = src.length.toLocaleString('en-IN');
  if (el('s-count-sub')) el('s-count-sub').textContent = `${uniqueV} credit notes · ${unique} customers`;
  if (el('s-clean')) el('s-clean').textContent = cleanRows.toLocaleString('en-IN');
  if (el('s-clean-sub')) el('s-clean-sub').textContent = src.length ? Math.round(cleanRows / src.length * 100) + '% clean' : '—';
  if (el('s-branches')) el('s-branches').textContent = [...new Set(src.map(d => d.branch))].length;
  if (el('s-branches-sub')) el('s-branches-sub').textContent = [...new Set(src.map(d => d.branch))].join(', ');
  if (el('s-grossfixed')) el('s-grossfixed').textContent = src.filter(d => d._flags.includes('gross-computed') || d._flags.includes('gross-corrected')).length;
  if (el('s-grossfixed-sub')) el('s-grossfixed-sub').textContent = src.filter(d => d._flags.includes('gross-corrected')).length + ' corrected';
  const errCnt = src.filter(d => d._flags.includes('error')).length;
  if (el('s-errors')) el('s-errors').textContent = errCnt;
  if (el('s-errors-sub')) el('s-errors-sub').textContent = `${errCnt} rows with issues`;

  // Top 8 customers
  const cMap = {};
  src.forEach(d => {
    if (!cMap[d.customer]) cMap[d.customer] = { gross: 0, count: 0, branches: new Set() };
    cMap[d.customer].gross += d.gross;
    cMap[d.customer].count++;
    cMap[d.customer].branches.add(d.branch);
  });
  const topC = Object.entries(cMap).sort((a, b) => b[1].gross - a[1].gross).slice(0, 8);
  const cMax = topC[0] ? topC[0][1].gross : 1;
  const tcc = el('topCustomersChart');
  if (tcc) {
    tcc.innerHTML = topC.length ? topC.map(([name, v]) => `
      <div class="bar-row" onclick="drillCustomer('${esc(name.replace(/'/g, "\\'"))}')">
        <div class="bar-tooltip">${esc(name)} — ${fmt(v.gross)} · ${v.count} CNs · ${[...v.branches].join(', ')}</div>
        <div class="bar-label" title="${esc(name)}">${esc(name.slice(0, 22))}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${(v.gross / cMax * 100).toFixed(1)}%"></div></div>
        <div class="bar-val">${fmtK(v.gross)}</div>
      </div>`).join('')
      : '<div style="color:var(--muted);font-size:13px;padding:12px 0">No data available.</div>';
  }

  // Branch donut
  renderBranchDonut(src);

  // Monthly grid
  const m = {};
  src.forEach(d => {
    const p = d.date.split('/');
    if (p.length < 3) return;
    const k = p[1] + '/' + p[2];
    if (!m[k]) m[k] = { count: 0, gross: 0, key: k, label: new Date(+p[2], +p[1] - 1).toLocaleString('default', { month: 'short', year: '2-digit' }) };
    m[k].count++; m[k].gross += d.gross;
  });
  const sorted = Object.entries(m).sort((a, b) => new Date('01/' + a[0]) - new Date('01/' + b[0])).slice(-12);
  const mg = el('monthlyGrid');
  if (mg) {
    mg.innerHTML = sorted.map(([k, v]) => `
      <div class="month-item${dashFilter?.type === 'month' && dashFilter.value === k ? ' selected' : ''}" onclick="dashMonthClick('${k}')">
        <div class="month-name">${v.label}</div>
        <div class="month-count">${v.count}</div>
        <div class="month-val">${fmtK(v.gross)}</div>
      </div>`).join('');
  }

  // Branch breakdown
  const bMap = {};
  src.forEach(d => {
    if (!bMap[d.branch]) bMap[d.branch] = { count: 0, gross: 0, customers: new Set() };
    bMap[d.branch].count++; bMap[d.branch].gross += d.gross;
    bMap[d.branch].customers.add(d.customer);
  });
  const bb = el('branchBreakdown');
  if (bb) {
    bb.innerHTML = Object.entries(bMap).map(([name, v]) => `
      <div class="type-row${dashFilter?.type === 'branch' && dashFilter.value === name ? ' selected' : ''}" onclick="dashBranchClick('${esc(name)}')">
        <div class="type-left">
          <div class="type-dot" style="background:${getBranchColor(name)}"></div>
          <div class="type-name">${esc(name)}</div>
        </div>
        <div class="type-right">
          <div class="type-count">${v.count} rows · ${v.customers.size} customers</div>
          <div class="type-amount">${fmtK(v.gross)}</div>
        </div>
      </div>`).join('');
  }

  // Filter banner
  const banner = el('dashFilterBanner');
  if (banner) {
    if (dashFilter) {
      banner.style.display = 'flex';
      let txt = '';
      if (dashFilter.type === 'branch') txt = `🏢 Showing: ${dashFilter.value} branch only`;
      if (dashFilter.type === 'month') txt = `📅 Showing: ${dashFilter.value} only`;
      if (dashFilter.type === 'customer') txt = `👤 Showing: ${dashFilter.value} only`;
      el('dashFilterText').textContent = txt;
    } else {
      banner.style.display = 'none';
    }
  }
}

function renderBranchDonut(src) {
  const canvas = document.getElementById('branchDonut');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const bMap = {};
  src.forEach(d => { bMap[d.branch] = (bMap[d.branch] || 0) + d.gross; });
  const entries = Object.entries(bMap).sort((a, b) => b[1] - a[1]);
  const total = entries.reduce((s, [, v]) => s + v, 0);
  if (!total) return;

  ctx.clearRect(0, 0, 160, 160);
  let angle = -Math.PI / 2;
  entries.forEach(([name, val]) => {
    const slice = (val / total) * Math.PI * 2;
    ctx.beginPath();
    ctx.moveTo(80, 80);
    ctx.arc(80, 80, 70, angle, angle + slice);
    ctx.closePath();
    ctx.fillStyle = getBranchColor(name);
    ctx.fill();
    angle += slice;
  });
  // Donut hole
  ctx.beginPath();
  ctx.arc(80, 80, 40, 0, Math.PI * 2);
  ctx.fillStyle = '#fff';
  ctx.fill();
  // Center text
  ctx.fillStyle = '#0d0f12';
  ctx.font = 'bold 14px DM Sans, sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText(entries.length, 80, 77);
  ctx.font = '11px DM Sans, sans-serif';
  ctx.fillStyle = '#8b93a5';
  ctx.fillText('branches', 80, 93);
}

function dashBranchClick(name) {
  dashFilter = dashFilter?.type === 'branch' && dashFilter.value === name ? null : { type: 'branch', value: name };
  updateDashboard();
}
function dashMonthClick(key) {
  dashFilter = dashFilter?.type === 'month' && dashFilter.value === key ? null : { type: 'month', value: key };
  updateDashboard();
}
function drillCustomer(name) {
  dashFilter = dashFilter?.type === 'customer' && dashFilter.value === name ? null : { type: 'customer', value: name };
  updateDashboard();
}
function clearDashFilter() { dashFilter = null; updateDashboard(); }

/* ══════════════════════════════════════════════════════════════════
   ANALYTICS
══════════════════════════════════════════════════════════════════ */
function renderAnalytics() {
  if (!allData.length) return;

  const totGross = allData.reduce((s, d) => s + d.gross, 0);
  const avgGross = totGross / allData.length;
  const byVoucher = {};
  allData.forEach(d => { byVoucher[d.creditNote] = (byVoucher[d.creditNote] || 0) + d.gross; });
  const maxEntry = Object.entries(byVoucher).sort((a, b) => b[1] - a[1])[0];

  const dates = allData.map(d => dp(d.date)).filter(Boolean);
  const minD = new Date(Math.min(...dates));
  const maxD = new Date(Math.max(...dates));
  const fmtD = d => d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });

  const el = id => document.getElementById(id);
  if (el('a-avg')) el('a-avg').textContent = fmtK(avgGross);
  if (el('a-max')) el('a-max').textContent = maxEntry ? fmtK(maxEntry[1]) : '—';
  if (el('a-max-cust')) el('a-max-cust').textContent = maxEntry ? maxEntry[0] : '—';
  if (el('a-range')) el('a-range').textContent = dates.length ? `${fmtD(minD)} → ${fmtD(maxD)}` : '—';
  if (el('a-branches')) el('a-branches').textContent = [...new Set(allData.map(d => d.branch))].length;
  if (el('a-total')) el('a-total').textContent = fmtK(totGross);

  // Top 15 customers
  const cMap = {};
  allData.forEach(d => { cMap[d.customer] = (cMap[d.customer] || 0) + d.gross; });
  const topC = Object.entries(cMap).sort((a, b) => b[1] - a[1]).slice(0, 15);
  const cMax = topC[0] ? topC[0][1] : 1;
  const bcc = el('bigCustomerChart');
  if (bcc) {
    bcc.innerHTML = topC.map(([name, val]) => `
      <div class="bar-row">
        <div class="bar-label" title="${esc(name)}">${esc(name.slice(0, 22))}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${(val / cMax * 100).toFixed(1)}%"></div></div>
        <div class="bar-val">${fmtK(val)}</div>
      </div>`).join('');
  }

  // Monthly gross
  const mo = {};
  allData.forEach(d => {
    const p = d.date.split('/');
    if (p.length < 3) return;
    const k = p[1] + '/' + p[2];
    if (!mo[k]) mo[k] = { gross: 0, label: new Date(+p[2], +p[1] - 1).toLocaleString('default', { month: 'short', year: '2-digit' }) };
    mo[k].gross += d.gross;
  });
  const moS = Object.entries(mo).sort((a, b) => new Date('01/' + a[0]) - new Date('01/' + b[0]));
  const mMax = Math.max(...moS.map(([, v]) => v.gross), 1);
  const mbc = el('monthBarChart');
  if (mbc) {
    mbc.innerHTML = moS.map(([, v]) => `
      <div class="bar-row">
        <div class="bar-label">${v.label}</div>
        <div class="bar-track"><div class="bar-fill green" style="width:${(v.gross / mMax * 100).toFixed(1)}%"></div></div>
        <div class="bar-val">${fmtK(v.gross)}</div>
      </div>`).join('');
  }

  // Branch gross
  const bMap = {};
  allData.forEach(d => { bMap[d.branch] = (bMap[d.branch] || 0) + d.gross; });
  const bEntries = Object.entries(bMap).sort((a, b) => b[1] - a[1]);
  const bMax = bEntries[0] ? bEntries[0][1] : 1;
  const bbc = el('branchBarChart');
  if (bbc) {
    bbc.innerHTML = bEntries.map(([name, val]) => `
      <div class="bar-row">
        <div class="bar-label">${esc(name)}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${(val / bMax * 100).toFixed(1)}%;background:${getBranchColor(name)}"></div></div>
        <div class="bar-val">${fmtK(val)}</div>
      </div>`).join('');
  }
}

/* ══════════════════════════════════════════════════════════════════
   CLEAN LOG
══════════════════════════════════════════════════════════════════ */
function updateCleanLog() {
  const totGross = allData.reduce((s, d) => s + d.gross, 0);
  const unique = new Set(allData.map(d => d.customer.trim().toLowerCase())).size;
  const uniqueV = new Set(allData.map(d => d.creditNote)).size;
  const branches = [...new Set(allData.map(d => d.branch))];

  const cst = document.getElementById('cleanSummaryTable');
  if (cst) {
    cst.innerHTML = `
      <table class="issues-table">
        <tr><th>Metric</th><th>Value</th></tr>
        <tr><td>Files loaded</td><td><strong>${loadedWorkbooks.length}</strong></td></tr>
        <tr><td>Sheets selected</td><td><strong>${selectedSheets.size}</strong></td></tr>
        <tr><td>Output mode</td><td><strong>${outputMode}</strong></td></tr>
        <tr><td>Total rows before cleaning</td><td><strong>${rawParsed.length}</strong></td></tr>
        <tr><td>Output rows after cleaning</td><td><strong>${allData.length}</strong></td></tr>
        ${[...selectedSheets].map(uKey => {
      const [fName, sName] = uKey.split(' » ');
      const cnt = allData.filter(d => d.branch === sName).length;
      return `<tr><td>↳ ${esc(sName)} <span style="font-size:10px;color:var(--muted)">(${esc(fName)})</span></td><td>${cnt} rows</td></tr>`;
    }).join('')}
        <tr><td>Unique credit notes</td><td>${uniqueV}</td></tr>
        <tr><td>Cancelled vouchers excluded</td><td>${cancelledCount}</td></tr>
        <tr><td>Duplicate rows removed</td><td>${cleanStats.dupes}</td></tr>
        <tr><td>Gross values fixed / computed</td><td>${cleanStats.grossFixed}</td></tr>
        <tr><td>Gross values corrected (were wrong)</td><td>${cleanStats.grossWrong}</td></tr>
        <tr><td>Errors found</td><td>${errorLog.length}</td></tr>
        <tr><td>Unique customers</td><td>${unique}</td></tr>
        <tr><td>Total gross value</td><td><strong>${fmt(totGross)}</strong></td></tr>
        <tr><td>Branches</td><td>${branches.map(b => `<span class="badge badge-branch" style="background:${getBranchColor(b)}">${esc(b)}</span>`).join(' ')}</td></tr>
      </table>`;
  }

  const flags = [];
  const missingRate = allData.filter(d => d.unitRate === 0).length;
  const missingGross = allData.filter(d => d.gross === 0).length;
  if (missingRate) flags.push({ label: 'Unit Rate = 0', count: missingRate, sev: 'info' });
  if (missingGross) flags.push({ label: 'Gross = ₹0', count: missingGross, sev: 'warn' });
  if (cleanStats.grossWrong) flags.push({ label: 'Gross values corrected', count: cleanStats.grossWrong, sev: 'warn' });
  if (cleanStats.dupes) flags.push({ label: 'Duplicate rows removed', count: cleanStats.dupes, sev: 'warn' });

  const qf = document.getElementById('qualityFlags');
  if (qf) {
    qf.innerHTML = flags.length
      ? `<table class="issues-table">
          <tr><th>Issue</th><th>Count</th><th>Status</th></tr>
          ${flags.map(f => `<tr><td>${f.label}</td><td>${f.count}</td><td><span class="issue-badge ib-${f.sev === 'err' ? 'err' : f.sev === 'warn' ? 'warn' : 'ok'}">${f.sev === 'err' ? 'Review' : f.sev === 'warn' ? 'Check' : 'Info'}</span></td></tr>`).join('')}
        </table>`
      : '<div style="color:var(--green);font-size:13px;font-weight:500">✅ No quality issues found</div>';
  }
}

/* ══════════════════════════════════════════════════════════════════
   UTILITY FUNCTIONS
══════════════════════════════════════════════════════════════════ */
function fmt(n, sym = '₹') {
  if (!n || n === 0) return '—';
  return sym + n.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmt2(n) {
  return (n || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtQty(n) {
  if (!n || n === 0) return '—';
  return n.toLocaleString('en-IN', { maximumFractionDigits: 4 });
}
function fmtK(n) {
  if (!n) return '₹0';
  if (n >= 10000000) return '₹' + (n / 10000000).toFixed(2) + 'Cr';
  if (n >= 100000) return '₹' + (n / 100000).toFixed(2) + 'L';
  return '₹' + Math.round(n).toLocaleString('en-IN');
}
function dp(s) {
  try {
    const [d, mo, y] = s.split('/');
    return new Date(+y, +mo - 1, +d).getTime();
  } catch { return 0; }
}
function today() {
  return new Date().toISOString().slice(0, 10);
}
function esc(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

let _tt;
function showToast(msg) {
  clearTimeout(_tt);
  const toast = document.getElementById('toast');
  const toastMsg = document.getElementById('toastMsg');
  if (toastMsg) toastMsg.textContent = msg;
  if (toast) {
    toast.classList.add('show');
    _tt = setTimeout(() => toast.classList.remove('show'), 3500);
  }
}