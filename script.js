'use strict';

/* ══════════════════════════════════════════════════════════════════
   BRANCH COLOR MAP
═══════════════════════════════════════════════════════════════════ */
const BRANCH_COLORS = ['#2563eb','#059669','#d97706','#dc2626','#7c3aed','#ea580c','#0d9488','#be185d'];
let branchColorMap = {};
function getBranchColor(branch) {
  if (!branchColorMap[branch]) {
    const idx = Object.keys(branchColorMap).length % BRANCH_COLORS.length;
    branchColorMap[branch] = BRANCH_COLORS[idx];
  }
  return branchColorMap[branch];
}

/* ══════════════════════════════════════════════════════════════════
   MOBILE SIDEBAR
═══════════════════════════════════════════════════════════════════ */
function toggleSidebar() {
  document.querySelector('.sidebar').classList.toggle('open');
  document.getElementById('sidebarOverlay').classList.toggle('open');
}
document.querySelectorAll('.nav-item').forEach(item => {
  item.addEventListener('click', () => {
    if (window.innerWidth <= 768) {
      document.querySelector('.sidebar').classList.remove('open');
      document.getElementById('sidebarOverlay').classList.remove('open');
    }
  });
});

/* ══════════════════════════════════════════════════════════════════
   STATE
═══════════════════════════════════════════════════════════════════ */
let rawParsed     = [];   // after parseData() — unprocessed item rows
let allData       = [];   // after cleaning pipeline
let filtered      = [];   // after applyFilters()
let errorLog      = [];
let currentSort   = { col:'date', dir:'asc' };
let cancelledCount = 0;
let currentFile   = '';
let sheetNames    = [];
let sheetRowCounts = {};
let cleanStats    = { dupes:0, grossFixed:0, grossWrong:0, missingFields:0 };
let storedWorkbook = null;     // keep the parsed workbook so we can re-merge selected sheets
let selectedSheets = new Set(); // which sheets the user has ticked

/* ══════════════════════════════════════════════════════════════════
   NAVIGATION
═══════════════════════════════════════════════════════════════════ */
function switchPage(id) {
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
    if(page) page.classList.add('active');
    
    const titles = { dashboard:'Dashboard', preview:'Data Preview', records:'Records',
      errors:'Error Log', analytics:'Analytics', clean:'Clean Log' };
    document.getElementById('topbarTitle').textContent = titles[id] || id;
    document.getElementById('exportBtn').style.display = allData.length ? 'flex' : 'none';
    document.getElementById('recordPill').style.display = allData.length ? 'flex' : 'none';
    document.getElementById('recordPill').textContent = allData.length + ' records';
    
    if (id === 'preview' && storedWorkbook) {
      document.getElementById('btnProcessMerge').disabled = false;
      document.getElementById('btnProcessMerge').innerHTML = '🔄 Process & Merge';
    }
    if (id === 'analytics') renderAnalytics();
    if (id === 'records') applyFilters();
    if (id === 'errors') renderErrorLog();
  }
}
switchPage('upload');

/* ══════════════════════════════════════════════════════════════════
   FILE HANDLING — MULTI-SHEET
═══════════════════════════════════════════════════════════════════ */
const dropZone = document.getElementById('dropZone');
if (dropZone) {
  dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('dragover'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault(); dropZone.classList.remove('dragover');
    if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]);
  });
}

function handleFile(e) {
  if (e.target.files[0]) processFile(e.target.files[0]);
}

function processFile(file) {
  currentFile = file.name;
  branchColorMap = {};
  cancelledCount = 0;
  setProgress(true, 'Reading file…', 20);

  const reader = new FileReader();
  reader.onload = e => {
    setProgress(true, 'Parsing sheets…', 40);
    try {
      const wb = XLSX.read(e.target.result, { type:'binary', cellDates:true });
      storedWorkbook = wb;
      sheetNames = wb.SheetNames;
      sheetRowCounts = {};

      setProgress(true, 'Scanning ' + sheetNames.length + ' sheets…', 60);

      setTimeout(() => {
        /* Pre-parse every sheet to get row counts for the selector */
        sheetNames.forEach(sheetName => {
          const ws = wb.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(ws, { header:1, raw:false, defval:'' });
          const records = parseSheet(rows, sheetName);
          sheetRowCounts[sheetName] = records.length;
        });

        /* Select all sheets by default */
        selectedSheets = new Set(sheetNames);

        /* Render the sheet selector UI */
        renderSheetSelector();
        document.getElementById('btnNextPreview').style.display = '';

        setProgress(false);
        document.getElementById('fileNameDisplay').textContent = currentFile;
        showToast('📂 Found ' + sheetNames.length + ' sheets — select which ones to merge, then go to Preview.');
      }, 50);
    } catch (err) {
      setProgress(false);
      alert('Error reading file: ' + err.message);
    }
  };
  reader.readAsBinaryString(file);
}

function setProgress(show, label='', pct=0) {
  const pWrap = document.getElementById('progressWrap');
  if (pWrap) {
    pWrap.style.display = show ? 'block' : 'none';
    document.getElementById('progressLabel').textContent = label;
    document.getElementById('progressInner').style.width = pct + '%';
  }
}

/* ══════════════════════════════════════════════════════════════════
   SHEET SELECTOR — render checkboxes, toggle controls
═══════════════════════════════════════════════════════════════════ */

/** Render the sheet selector panel in the upload card */
function renderSheetSelector() {
  const panel = document.getElementById('sheetSelector');
  const list  = document.getElementById('sheetList');
  if (panel) panel.classList.add('show');
  if (!list) return;

  const ssCount = document.getElementById('ssCount');
  if (ssCount) ssCount.textContent = sheetNames.length + ' sheets';

  list.innerHTML = sheetNames.map(name => {
    const c = getBranchColor(name);
    const rows = sheetRowCounts[name] || 0;
    const checked = selectedSheets.has(name);
    const id = 'ss-' + name.replace(/\s+/g, '_');
    return `<label class="sheet-item${checked ? '' : ' disabled'}" id="si-${name.replace(/\s+/g,'_')}" for="${id}">
      <input type="checkbox" id="${id}" ${checked ? 'checked' : ''} onchange="toggleSheet('${esc(name.replace(/'/g,"\\'"))}', this.checked)">
      <div class="sheet-item-dot" style="background:${c}"></div>
      <div class="sheet-item-name">${esc(name)}</div>
      <div class="sheet-item-meta">
        <span class="sheet-item-rows">${rows} rows</span>
      </div>
    </label>`;
  }).join('');

  updateSheetMergeInfo();
}

/** Toggle a single sheet on/off */
function toggleSheet(name, checked) {
  if (checked) selectedSheets.add(name);
  else selectedSheets.delete(name);

  /* Update visual state */
  const el = document.getElementById('si-' + name.replace(/\s+/g, '_'));
  if (el) el.classList.toggle('disabled', !checked);
  updateSheetMergeInfo();
}

/** Select all sheets */
function sheetSelectAll() {
  selectedSheets = new Set(sheetNames);
  sheetNames.forEach(n => {
    const cb = document.getElementById('ss-' + n.replace(/\s+/g, '_'));
    if (cb) cb.checked = true;
    const el = document.getElementById('si-' + n.replace(/\s+/g, '_'));
    if (el) el.classList.remove('disabled');
  });
  updateSheetMergeInfo();
}

/** Deselect all sheets */
function sheetDeselectAll() {
  selectedSheets.clear();
  sheetNames.forEach(n => {
    const cb = document.getElementById('ss-' + n.replace(/\s+/g, '_'));
    if (cb) cb.checked = false;
    const el = document.getElementById('si-' + n.replace(/\s+/g, '_'));
    if (el) el.classList.add('disabled');
  });
  updateSheetMergeInfo();
}

/** Invert selection */
function sheetInvert() {
  sheetNames.forEach(n => {
    if (selectedSheets.has(n)) selectedSheets.delete(n);
    else selectedSheets.add(n);
    const cb = document.getElementById('ss-' + n.replace(/\s+/g, '_'));
    if (cb) cb.checked = selectedSheets.has(n);
    const el = document.getElementById('si-' + n.replace(/\s+/g, '_'));
    if (el) el.classList.toggle('disabled', !selectedSheets.has(n));
  });
  updateSheetMergeInfo();
}

/** Update the merge info text below the sheet list */
function updateSheetMergeInfo() {
  const info = document.getElementById('sheetMergeInfo');
  if (!info) return;
  const count = selectedSheets.size;
  const totalRows = [...selectedSheets].reduce((s, n) => s + (sheetRowCounts[n]||0), 0);
  if (count === 0) {
    info.textContent = '⚠️ No sheets selected — select at least one sheet to proceed.';
    info.style.color = 'var(--red)';
  } else {
    info.textContent = '✅ ' + count + ' of ' + sheetNames.length + ' sheets selected · ~' + totalRows + ' rows will be merged';
    info.style.color = 'var(--green)';
  }
}

/* ══════════════════════════════════════════════════════════════════
   MERGE SELECTED SHEETS — buildRawParsed()
   Re-parses only the selected sheets from the stored workbook
═══════════════════════════════════════════════════════════════════ */
function buildRawParsed() {
  if (!storedWorkbook) return;
  cancelledCount = 0;
  branchColorMap = {};
  const allRecords = [];
  
  const sheetsToMerge = sheetNames.filter(n => selectedSheets.has(n));

  sheetsToMerge.forEach(sheetName => {
    getBranchColor(sheetName); // assign color
    const ws = storedWorkbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header:1, raw:false, defval:'' });
    const records = parseSheet(rows, sheetName);
    allRecords.push(...records);
  });

  rawParsed = allRecords.map(r => ({ ...r, _flags:[] }));
}

/** Process & Merge button — merges selected sheets, then cleans */
function processAndMerge() {
  if (selectedSheets.size === 0) {
    showToast('⚠️ No sheets selected — go back to Upload and select at least one sheet.');
    return;
  }

  const btn = document.getElementById('btnProcessMerge');
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '⏳ Processing…';
  }
  showToast('🔄 Merging ' + selectedSheets.size + ' sheets…');

  setTimeout(() => {
    /* Step 1: Re-parse only selected sheets */
    buildRawParsed();

    if (!rawParsed.length) {
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = '🔄 Process & Merge';
      }
      showToast('⚠️ No records found in selected sheets.');
      return;
    }

    /* Step 2: Run cleaning pipeline */
    runCleaningPipeline();

    /* Step 3: Render previews */
    renderRawPreview();
    renderCleanedPreview();
    renderColumnMapper();

    if (btn) btn.innerHTML = '✅ Processed & Merged';
    document.getElementById('btnDownloadCleaned').disabled = false;
    document.getElementById('rerunBtn').disabled = false;

    /* Switch to cleaned tab */
    document.querySelectorAll('#page-preview .tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('#page-preview .tab-pane').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('#page-preview .tab-btn')[1].classList.add('active');
    document.getElementById('preview-cleaned').classList.add('active');

    const mergedNames = [...selectedSheets].join(', ');
    showToast('✅ Merged ' + selectedSheets.size + ' sheets (' + mergedNames + ') → ' + allData.length + ' rows. Ready to download!');
  }, 80);
}

/* ══════════════════════════════════════════════════════════════════
   SHEET PARSER — parseSheet(rows, branchName)
   Parses a single sheet's rows into flat item-level records.
   Adds the Branch column from the sheet name.
═══════════════════════════════════════════════════════════════════ */
function parseSheet(rows, branchName) {
  const CN_RE    = /^(CN|CNG|CNS)\//i;
  const CANCELLED = /\(cancelled/i;
  const records  = [];

  const cn = v => {
    if (v === null || v === undefined || String(v).trim() === '') return 0;
    const s = String(v).replace(/[^0-9.-]/g, '');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  };

  /* Find header row to determine column indices dynamically */
  let headerIdx = -1;
  let colMap = { date:0, particulars:1, vtype:2, voucher:3, narration:4, qty:5, rate:6, value:7, gross:8 };

  for (let i = 0; i < Math.min(20, rows.length); i++) {
    const rowLower = rows[i].map(c => String(c||'').trim().toLowerCase());
    if (rowLower.some(c => c === 'date') && rowLower.some(c => c.includes('voucher'))) {
      headerIdx = i;
      /* Map columns dynamically from header text */
      rowLower.forEach((h, idx) => {
        if (h === 'date') colMap.date = idx;
        if (h === 'particulars') colMap.particulars = idx;
        if (h === 'voucher type') colMap.vtype = idx;
        if (h.includes('voucher no')) colMap.voucher = idx;
        if (h === 'narration') colMap.narration = idx;
        if (h === 'quantity') colMap.qty = idx;
        if (h === 'rate') colMap.rate = idx;
        if (h === 'value') colMap.value = idx;
        if (h.includes('gross')) colMap.gross = idx;
      });
      break;
    }
  }

  /* If no header found, try a fallback scan for CN/ pattern rows */
  if (headerIdx < 0) {
    for (let i = 0; i < Math.min(30, rows.length); i++) {
      const r = rows[i];
      for (let c = 0; c < (r ? r.length : 0); c++) {
        if (/^(CN|CNG|CNS)\//i.test(String(r[c]||'').trim())) {
          colMap.voucher = c;
          headerIdx = Math.max(0, i - 1);
          break;
        }
      }
      if (headerIdx >= 0) break;
    }
  }

  let i = headerIdx >= 0 ? headerIdx + 1 : 0;

  while (i < rows.length) {
    const r       = rows[i];
    const voucher = String(r[colMap.voucher] || '').trim();
    const dateRaw = String(r[colMap.date] || '').trim();

    if (!CN_RE.test(voucher) || !dateRaw || dateRaw.toLowerCase() === 'date') { i++; continue; }

    const particulars = String(r[colMap.particulars] || '').trim();
    if (CANCELLED.test(particulars)) { cancelledCount++; i++; continue; }

    /* Date normalisation */
    let dateStr = '';
    try {
      const d = new Date(dateRaw);
      dateStr = isNaN(d.getTime()) ? dateRaw : d.toLocaleDateString('en-GB');
    } catch { dateStr = dateRaw; }

    const customer   = particulars;
    const creditNote = voucher;
    const narration  = String(r[colMap.narration] || '').replace(/_x000D_\\n/g,' ').replace(/_x000D_\n/g,' ').trim();
    const mainValue  = cn(r[colMap.value]);
    let   mainGross  = cn(r[colMap.gross]);
    if (mainGross === 0 && mainValue > 0) mainGross = mainValue;

    /* Collect item sub-rows */
    let j = i + 1;
    const rawItems = [];
    while (j < rows.length) {
      const nr = rows[j];
      const nv = String(nr[colMap.voucher] || '').trim();
      const nd = String(nr[colMap.date] || '').trim();
      const np = String(nr[colMap.particulars] || '').trim().toLowerCase();
      if (CN_RE.test(nv) && nd) break;
      if (np === 'grand total') break;
      const iName = String(nr[colMap.particulars] || '').trim();
      if (iName && np !== 'grand total') {
        rawItems.push({ name: iName, qty: cn(nr[colMap.qty]), rate: cn(nr[colMap.rate]), value: cn(nr[colMap.value]) });
      }
      j++;
    }

    /* Emit one record per item or one fallback */
    if (rawItems.length > 0) {
      const totalItemVal = rawItems.reduce((s, it) => s + it.value, 0);
      rawItems.forEach(it => {
        let itemGross = 0;
        if (totalItemVal > 0 && it.value > 0) {
          itemGross = Math.round(it.value / totalItemVal * mainGross * 100) / 100;
        } else if (rawItems.length === 1) {
          itemGross = mainGross;
        } else if (totalItemVal === 0) {
          itemGross = Math.round(mainGross / rawItems.length * 100) / 100;
        }
        records.push({
          date: dateStr, customer, creditNote, narration,
          qty: it.qty, unitRate: it.rate, netRate: it.value, gross: itemGross,
          branch: branchName, _flags: []
        });
      });
    } else {
      records.push({
        date: dateStr, customer, creditNote, narration,
        qty: cn(r[colMap.qty]), unitRate: cn(r[colMap.rate]), netRate: mainValue, gross: mainGross,
        branch: branchName, _flags: []
      });
    }
    i = j;
  }
  return records;
}

/* ══════════════════════════════════════════════════════════════════
   CLEANING PIPELINE
═══════════════════════════════════════════════════════════════════ */
function runCleaningPipeline() {
  const opts = {
    dedup   : document.getElementById('opt-dedup').checked,
    gross   : document.getElementById('opt-gross').checked,
    trim    : document.getElementById('opt-trim').checked,
    validate: document.getElementById('opt-validate').checked,
  };

  cleanStats = { dupes:0, grossFixed:0, grossWrong:0, missingFields:0 };
  errorLog   = [];

  let data = rawParsed.map(r => ({ ...r, _flags:[] }));

  /* 1. Trim */
  if (opts.trim) data = trimSpaces(data);
  /* 2. Gross fix */
  if (opts.gross) data = calculateGross(data);
  /* 3. Validate */
  if (opts.validate) data = validateData(data);
  /* 4. Dedup */
  if (opts.dedup) data = detectDuplicates(data);
  /* 5. Sort by Date ascending */
  data.sort((a, b) => dp(a.date) - dp(b.date));

  allData  = data;
  filtered = [...allData];

  populateFilters();
  applyFilters();
  updateDashboard();
  updateCleanLog();

  const fs = document.getElementById('fileStats');
  if (fs) fs.textContent = allData.length + ' rows · ' + sheetNames.length + ' sheets';
  const nr = document.getElementById('nb-records');
  if (nr) nr.textContent = allData.length;
  const ne = document.getElementById('nb-errors');
  if (ne) ne.textContent = errorLog.length;
}

function rerunCleaning() {
  if (!rawParsed.length) return;
  showToast('🔄 Re-running cleaning…');
  setTimeout(() => {
    runCleaningPipeline();
    renderCleanedPreview();
    showToast('✅ Done — ' + allData.length + ' rows');
  }, 50);
}

/* ─── CLEANING MODULES ───────────────────────────────────────── */

function trimSpaces(data) {
  return data.map(r => ({
    ...r,
    customer  : r.customer.trim(),
    creditNote: r.creditNote.trim(),
    narration : r.narration.trim(),
    branch    : r.branch.trim(),
    date      : r.date.trim(),
  }));
}

function calculateGross(data) {
  return data.map(r => {
    const row = { ...r };
    if (row.qty > 0 && row.unitRate > 0) {
      const expected = Math.round(row.qty * row.unitRate * 100) / 100;
      if (!row.gross || row.gross === 0) {
        row._grossOriginal = 0;
        row.gross = expected;
        row._flags = [...row._flags, 'gross-computed'];
        cleanStats.grossFixed++;
        pushError('info', row.creditNote, row.customer, row.date, row.branch,
          'Gross was missing — computed as Qty×Rate = ₹' + fmt2(expected));
      } else {
        const diff = Math.abs(row.gross - expected);
        if (diff > 1) {
          row._grossOriginal = row.gross;
          row.gross = expected;
          row._flags = [...row._flags, 'gross-corrected'];
          cleanStats.grossFixed++;
          cleanStats.grossWrong++;
          pushError('warn', row.creditNote, row.customer, row.date, row.branch,
            'Gross was ₹' + fmt2(row._grossOriginal) + ' but Qty×Rate = ₹' + fmt2(expected) + ' — corrected');
        }
      }
    }
    return row;
  });
}

function validateData(data) {
  return data.map(r => {
    const row = { ...r };
    const missing = [];
    if (!row.date)       missing.push('Date');
    if (!row.customer)   missing.push('Customer');
    if (!row.creditNote) missing.push('Credit Note Number');
    if (missing.length > 0) {
      row._flags = [...row._flags, 'error'];
      cleanStats.missingFields++;
      pushError('err', row.creditNote || '?', row.customer || '?', row.date || '?', row.branch,
        'Missing: ' + missing.join(', '));
    }
    return row;
  });
}

function detectDuplicates(data) {
  const seen = new Set();
  const output = [];
  data.forEach(r => {
    const key = [r.creditNote, r.customer.toLowerCase(), r.qty, r.unitRate, r.branch].join('||');
    if (seen.has(key)) {
      cleanStats.dupes++;
      pushError('warn', r.creditNote, r.customer, r.date, r.branch,
        'Duplicate row detected (same CN + Customer + Qty + Rate + Branch)');
    } else {
      seen.add(key);
      output.push(r);
    }
  });
  return output;
}

/* ══════════════════════════════════════════════════════════════════
   ERROR LOG
═══════════════════════════════════════════════════════════════════ */
function pushError(sev, voucher, customer, date, branch, msg) {
  errorLog.push({ sev, voucher, customer, date, branch, msg });
}

function renderErrorLog() {
  const list  = document.getElementById('errorList');
  const tbody = document.getElementById('errorTableBody');
  const card  = document.getElementById('errorTableCard');
  const count = document.getElementById('errorTotalCount');
  if (count) count.textContent = errorLog.length + ' issue' + (errorLog.length !== 1 ? 's' : '');

  if (!errorLog.length) {
    if (list) list.innerHTML = '<div class="no-errors">✅ No data issues detected.</div>';
    if (card) card.style.display = 'none';
    return;
  }

  if (list) {
    list.innerHTML = errorLog.slice(0, 40).map(e => `
      <div class="error-item">
        <span class="error-sev sev-${e.sev}">${e.sev.toUpperCase()}</span>
        <div>
          <div class="error-msg">${esc(e.msg)}</div>
          <div class="error-ref">${esc(e.voucher)} · ${esc(e.customer)} · ${esc(e.branch)}</div>
        </div>
      </div>`).join('') +
      (errorLog.length > 40 ? `<div style="padding:10px 16px;font-size:12px;color:var(--muted)">…and ${errorLog.length - 40} more</div>` : '');
  }

  if (card) card.style.display = '';
  if (tbody) {
    tbody.innerHTML = errorLog.map(e => `<tr>
      <td><span class="error-sev sev-${e.sev}">${e.sev.toUpperCase()}</span></td>
      <td style="font-family:'DM Mono',monospace;font-size:11.5px">${esc(e.voucher)}</td>
      <td>${esc(e.customer)}</td>
      <td style="font-family:'DM Mono',monospace;font-size:11.5px">${esc(e.date)}</td>
      <td><span class="badge badge-branch" style="background:${getBranchColor(e.branch)}">${esc(e.branch)}</span></td>
      <td style="font-size:12px">${esc(e.msg)}</td>
    </tr>`).join('');
  }
}

/* ══════════════════════════════════════════════════════════════════
   PREVIEW RENDERING
═══════════════════════════════════════════════════════════════════ */

function renderColumnMapper() {
  const section = document.getElementById('columnMapperSection');
  if (section) section.style.display = '';
  const cm = document.getElementById('columnMapper');
  if (cm) {
    const activeSheets = sheetNames.filter(n => selectedSheets.has(n));
    cm.innerHTML = activeSheets.map(name => {
      const c = getBranchColor(name);
      return `<div class="mapper-item">
        <label>Sheet</label>
        <div style="font-size:13px;font-weight:500;color:var(--ink);margin-top:2px;display:flex;align-items:center;gap:6px">
          <span style="width:8px;height:8px;border-radius:50%;background:${c};display:inline-block"></span>${esc(name)}
        </div>
        <div class="auto-detected">✓ ${sheetRowCounts[name] || 0} rows detected</div>
      </div>`;
    }).join('');
    
    const ps = document.getElementById('previewSubtitle');
    if (ps) {
      ps.textContent = activeSheets.length + ' of ' + sheetNames.length + ' sheets merged. Showing first 50 rows with Branch column.';
    }
  }
}

function renderRawPreview() {
  const table = document.getElementById('previewRawTable');
  if (!table) return;
  if (!rawParsed.length) {
    table.innerHTML = '<thead><tr><th>No data</th></tr></thead><tbody></tbody>';
    return;
  }
  const headers = ['Date','Customer','Credit Note No.','Narration','Qty','Unit Rate','Net Rate','Gross','Branch'];
  let html = '<thead><tr>' + headers.map(h => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';
  rawParsed.slice(0, 50).forEach(d => {
    const bc = getBranchColor(d.branch);
    html += `<tr>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(d.date)}</td>
      <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(d.customer)}</td>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(d.creditNote)}</td>
      <td style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(d.narration)||'—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${d.qty||'—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${d.unitRate?fmt2(d.unitRate):'—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${d.netRate?fmt2(d.netRate):'—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:600">${d.gross?fmt2(d.gross):'—'}</td>
      <td><span class="badge badge-branch" style="background:${bc}">${esc(d.branch)}</span></td>
    </tr>`;
  });
  if (rawParsed.length > 50) html += `<tr><td colspan="9" style="text-align:center;color:var(--muted);padding:12px">… and ${rawParsed.length - 50} more rows</td></tr>`;
  html += '</tbody>';
  table.innerHTML = html;
}

function renderCleanedPreview() {
  const table = document.getElementById('previewCleanedTable');
  if (!table) return;
  if (!allData.length) {
    table.innerHTML = '<thead><tr><th>No cleaned data yet</th></tr></thead><tbody></tbody>';
    return;
  }
  const headers = ['Date','Customer','Credit Note No.','Narration','Qty','Unit Rate','Net Rate','Gross','Branch','Status'];
  let html = '<thead><tr>' + headers.map(h => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';
  allData.slice(0, 50).forEach(d => {
    let statusHtml = '<span class="preview-badge" style="background:var(--green-light);color:var(--green)">OK</span>';
    if (d._flags.includes('error')) statusHtml = '<span class="preview-badge" style="background:var(--red-light);color:var(--red)">Error</span>';
    else if (d._flags.includes('gross-corrected')) statusHtml = '<span class="preview-badge" style="background:#fefce8;color:#854d0e">Gross Fixed</span>';
    else if (d._flags.includes('gross-computed')) statusHtml = '<span class="preview-badge" style="background:var(--green-light);color:var(--green)">Gross Added</span>';
    const bc = getBranchColor(d.branch);
    html += `<tr>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(d.date)}</td>
      <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(d.customer)}</td>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(d.creditNote)}</td>
      <td style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(d.narration)||'—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${d.qty||'—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${d.unitRate?fmt2(d.unitRate):'—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${d.netRate?fmt2(d.netRate):'—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:600">${d.gross?fmt2(d.gross):'—'}</td>
      <td><span class="badge badge-branch" style="background:${bc}">${esc(d.branch)}</span></td>
      <td>${statusHtml}</td>
    </tr>`;
  });
  if (allData.length > 50) html += `<tr><td colspan="10" style="text-align:center;color:var(--muted);padding:12px">… and ${allData.length - 50} more rows</td></tr>`;
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
   FILTERS & SORT
═══════════════════════════════════════════════════════════════════ */
function populateFilters() {
  const branches = [...new Set(allData.map(d => d.branch))].sort();
  const fb = document.getElementById('filterBranch');
  if (fb) fb.innerHTML = '<option value="">All branches</option>' + branches.map(b=>`<option>${b}</option>`).join('');
}

function applyFilters() {
  const searchBox = document.getElementById('searchBox');
  if (!searchBox) return;
  const q       = searchBox.value.toLowerCase();
  const br      = document.getElementById('filterBranch').value;
  const status  = document.getElementById('filterStatus').value;
  const sortSel = document.getElementById('filterSort').value;

  filtered = allData.filter(d => {
    const mq = !q || (d.customer.toLowerCase().includes(q) || d.creditNote.toLowerCase().includes(q) || d.narration.toLowerCase().includes(q) || d.branch.toLowerCase().includes(q));
    const mb = !br || d.branch === br;
    let ms = true;
    if (status === 'clean')      ms = !d._flags.length;
    if (status === 'grossfixed') ms = d._flags.includes('gross-computed') || d._flags.includes('gross-corrected');
    if (status === 'error')      ms = d._flags.includes('error');
    return mq && mb && ms;
  });

  const [col, dir] = sortSel.split('-');
  currentSort = { col: col || 'date', dir: dir || 'asc' };
  sortFiltered();
  renderTable();
}

function clearFilters() {
  document.getElementById('searchBox').value = '';
  document.getElementById('filterBranch').value = '';
  document.getElementById('filterStatus').value = '';
  document.getElementById('filterSort').value = 'date-asc';
  applyFilters();
}

function sortBy(col) {
  if (currentSort.col === col) currentSort.dir = currentSort.dir === 'asc' ? 'desc' : 'asc';
  else currentSort = { col, dir:'asc' };
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
    if (col === 'date')     return m * (dp(a.date) - dp(b.date));
    if (col === 'customer') return m * a.customer.localeCompare(b.customer);
    if (col === 'branch')   return m * a.branch.localeCompare(b.branch);
    return m * ((a[col] || 0) - (b[col] || 0));
  });
}

function updateSortIcons() {
  ['date','customer','qty','gross','branch'].forEach(c => {
    const el = document.getElementById('sa-' + c);
    if (!el) return;
    el.className = 'sort-arrows' + (currentSort.col === c ? ' ' + currentSort.dir : '');
  });
}

/* ══════════════════════════════════════════════════════════════════
   TABLE RENDER
═══════════════════════════════════════════════════════════════════ */
function renderTable() {
  const tbody = document.getElementById('tableBody');
  if (!tbody) return;
  const fc = document.getElementById('filterCount');
  if (fc) fc.textContent = filtered.length + ' rows';

  if (!filtered.length) {
    tbody.innerHTML = `<tr><td colspan="10"><div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-title">No records match</div></div></td></tr>`;
    const pi = document.getElementById('pageInfo');
    if (pi) pi.textContent = '0 rows';
    const ft = document.getElementById('filteredTotal');
    if (ft) ft.textContent = '';
    return;
  }

  const totGross = filtered.reduce((s, d) => s + d.gross, 0);
  let html = '';

  filtered.forEach((d, i) => {
    let rowClass = '';
    if (d._flags.includes('error'))               rowClass = 'row-error';
    else if (d._flags.includes('gross-corrected')) rowClass = 'row-gross-wrong';
    else if (d._flags.includes('gross-computed'))  rowClass = 'row-gross-fixed';

    const narr = d.narration.length > 35 ? d.narration.slice(0,32) + '…' : d.narration;
    const bc   = getBranchColor(d.branch);

    let statusBadges = '';
    if (!d._flags.length) {
      statusBadges = '<span class="badge badge-ok">OK</span>';
    } else {
      if (d._flags.includes('error'))           statusBadges += '<span class="badge badge-error">Error</span> ';
      if (d._flags.includes('gross-computed'))  statusBadges += '<span class="badge badge-fixed">Gross added</span> ';
      if (d._flags.includes('gross-corrected')) statusBadges += '<span class="badge badge-gross-wrong">Gross fixed</span> ';
    }

    html += `<tr class="${rowClass}" onclick="showDetail(${i})">
      <td class="cell-date">${d.date}</td>
      <td class="cell-customer"><div class="cell-customer-inner" title="${esc(d.customer)}">${esc(d.customer)}</div></td>
      <td class="cell-voucher">${esc(d.creditNote)}</td>
      <td class="cell-narr" title="${esc(d.narration)}">${esc(narr)||'—'}</td>
      <td class="cell-num">${fmtQty(d.qty)}</td>
      <td class="cell-num">${d.unitRate > 0 ? fmt(d.unitRate) : '—'}</td>
      <td class="cell-num">${d.netRate > 0 ? fmt(d.netRate) : '—'}</td>
      <td class="cell-num cell-gross">${fmt(d.gross)}${d._grossOriginal ? '<span title="Was: ₹'+fmt2(d._grossOriginal)+'" style="cursor:help;margin-left:3px;color:var(--amber)">✏️</span>' : ''}</td>
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
═══════════════════════════════════════════════════════════════════ */
function showDetail(idx) {
  const d = filtered[idx];
  let flagHtml = '';
  if (d._flags.length) {
    const msgs = [];
    if (d._flags.includes('error'))           msgs.push('❌ Missing required field(s)');
    if (d._flags.includes('gross-computed'))  msgs.push('🔧 Gross was missing — computed as Qty × Rate');
    if (d._flags.includes('gross-corrected')) msgs.push('✏️ Gross corrected from ₹' + fmt2(d._grossOriginal||0) + ' → ₹' + fmt2(d.gross));
    flagHtml = `<div class="flag-box">${msgs.join('<br>')}</div>`;
  }
  const bc = getBranchColor(d.branch);
  document.getElementById('modalVoucher').textContent = d.creditNote;
  document.getElementById('modalDate').textContent    = d.date + ' · ' + d.branch;
  document.getElementById('modalFlagBox').innerHTML   = flagHtml;
  document.getElementById('detailGrid').innerHTML = `
    <div><div class="detail-label">Customer</div><div class="detail-value">${esc(d.customer)}</div></div>
    <div><div class="detail-label">Branch</div><div class="detail-value"><span class="badge badge-branch" style="background:${bc}">${esc(d.branch)}</span></div></div>
    <div><div class="detail-label">Quantity</div><div class="detail-value mono">${fmtQty(d.qty)||'—'}</div></div>
    <div><div class="detail-label">Unit Rate</div><div class="detail-value mono">${d.unitRate > 0 ? fmt(d.unitRate) : '—'}</div></div>
    <div><div class="detail-label">Net Rate</div><div class="detail-value mono">${fmt(d.netRate)}</div></div>
    <div><div class="detail-label">Gross</div><div class="detail-value big">${fmt(d.gross)}</div></div>`;
  document.getElementById('narrationSection').innerHTML = d.narration
    ? `<div style="margin-bottom:12px"><div class="detail-label" style="margin-bottom:5px">Narration</div><div class="narr-box">${esc(d.narration)}</div></div>` : '';
  document.getElementById('detailOverlay').classList.add('open');
}
function closeModal() { document.getElementById('detailOverlay').classList.remove('open'); }
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

/* ══════════════════════════════════════════════════════════════════
   INTERACTIVE DASHBOARD
═══════════════════════════════════════════════════════════════════ */
let dashFilter = null; // { type:'branch', value:'MUMBAI' } etc.

function updateDashboard() {
  const src     = dashFilter ? getDashFiltered() : allData;
  const totGross  = src.reduce((s, d) => s + d.gross, 0);
  const unique    = new Set(src.map(d => d.customer.trim().toLowerCase())).size;
  const uniqueV   = new Set(src.map(d => d.creditNote)).size;
  const cleanRows = src.filter(d => !d._flags.length).length;

  const sc = document.getElementById('s-count');
  if (sc) sc.textContent = src.length.toLocaleString('en-IN');
  const scs = document.getElementById('s-count-sub');
  if (scs) scs.textContent = uniqueV + ' credit notes · ' + unique + ' customers';
  const scl = document.getElementById('s-clean');
  if (scl) scl.textContent = cleanRows.toLocaleString('en-IN');
  const scls = document.getElementById('s-clean-sub');
  if (scls) scls.textContent = src.length ? Math.round(cleanRows / src.length * 100) + '% clean' : '—';
  const sbr = document.getElementById('s-branches');
  if (sbr) sbr.textContent = [...new Set(src.map(d => d.branch))].length;
  const sbrs = document.getElementById('s-branches-sub');
  if (sbrs) sbrs.textContent = [...new Set(src.map(d => d.branch))].join(', ');
  const sgf = document.getElementById('s-grossfixed');
  if (sgf) sgf.textContent = src.filter(d => d._flags.includes('gross-computed')||d._flags.includes('gross-corrected')).length;
  const sgfs = document.getElementById('s-grossfixed-sub');
  if (sgfs) sgfs.textContent = src.filter(d => d._flags.includes('gross-corrected')).length + ' values corrected';
  const errSrc = src.filter(d => d._flags.includes('error')).length;
  const ser = document.getElementById('s-errors');
  if (ser) ser.textContent = errSrc;
  const sers = document.getElementById('s-errors-sub');
  if (sers) sers.textContent = errSrc + ' rows with issues';

  /* Sparklines — monthly distribution for total & clean */
  renderSparkline('spark-total', src);
  renderSparkline('spark-clean', src.filter(d => !d._flags.length));

  /* Top 8 Customers — with hover tooltips & click drill-down */
  const cMap = {};
  src.forEach(d => {
    if (!cMap[d.customer]) cMap[d.customer] = { gross:0, count:0, branches:new Set() };
    cMap[d.customer].gross += d.gross;
    cMap[d.customer].count++;
    cMap[d.customer].branches.add(d.branch);
  });
  const topC = Object.entries(cMap).sort((a,b) => b[1].gross-a[1].gross).slice(0,8);
  const cMax = topC[0] ? topC[0][1].gross : 1;
  const tcc = document.getElementById('topCustomersChart');
  if (tcc) {
    tcc.innerHTML = topC.map(([name,v]) => `
      <div class="bar-row" onclick="drillCustomer('${esc(name.replace(/'/g,"\\'"))}')">
        <div class="bar-tooltip">${esc(name)} — ${fmt(v.gross)} · ${v.count} CNs · ${[...v.branches].join(', ')}</div>
        <div class="bar-label" title="${esc(name)}">${esc(name.slice(0,22))}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${(v.gross/cMax*100).toFixed(1)}%"></div></div>
        <div class="bar-val">${fmtK(v.gross)}</div>
      </div>`).join('');
  }

  /* Branch Donut Chart */
  renderBranchDonut(src);

  /* Monthly grid — clickable */
  const m = {};
  src.forEach(d => {
    const p = d.date.split('/'); if (p.length < 3) return;
    const k = p[1]+'/'+p[2];
    if (!m[k]) m[k] = { count:0, gross:0, key:k, label: new Date(+p[2],+p[1]-1).toLocaleString('default',{month:'short',year:'2-digit'}) };
    m[k].count++; m[k].gross += d.gross;
  });
  const sorted = Object.entries(m).sort((a,b)=>new Date('01/'+a[0])-new Date('01/'+b[0])).slice(-12);
  const mg = document.getElementById('monthlyGrid');
  if (mg) {
    mg.innerHTML = sorted.map(([k,v]) => `
      <div class="month-item${dashFilter && dashFilter.type==='month' && dashFilter.value===k?' selected':''}" onclick="dashMonthClick('${k}')">
        <div class="month-name">${v.label}</div>
        <div class="month-count">${v.count}</div>
        <div class="month-val">${fmtK(v.gross)}</div>
      </div>`).join('');
  }

  /* Branch breakdown — clickable */
  const bMap = {};
  src.forEach(d => {
    if (!bMap[d.branch]) bMap[d.branch] = { count:0, gross:0, customers:new Set() };
    bMap[d.branch].count++; bMap[d.branch].gross += d.gross;
    bMap[d.branch].customers.add(d.customer);
  });
  const bb = document.getElementById('branchBreakdown');
  if (bb) {
    bb.innerHTML = Object.entries(bMap).map(([name,v]) => `
      <div class="type-row${dashFilter && dashFilter.type==='branch' && dashFilter.value===name?' selected':''}" onclick="dashBranchClick('${esc(name)}')">
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

  /* Filter banner */
  const banner = document.getElementById('dashFilterBanner');
  if (banner) {
    if (dashFilter) {
      banner.classList.add('show');
      let txt = '';
      if (dashFilter.type === 'branch') txt = '🏢 Showing: ' + dashFilter.value + ' branch only';
      else if (dashFilter.type === 'month') txt = '📅 Showing: ' + dashFilter.label + ' only';
      else if (dashFilter.type === 'customer') txt = '👤 Showing: ' + dashFilter.value;
      document.getElementById('dashFilterText').textContent = txt;
    } else {
      banner.classList.remove('show');
    }
  }
}

/** Get filtered data based on current dashboard filter */
function getDashFiltered() {
  if (!dashFilter) return allData;
  if (dashFilter.type === 'branch') return allData.filter(d => d.branch === dashFilter.value);
  if (dashFilter.type === 'month') {
    return allData.filter(d => {
      const p = d.date.split('/');
      return p.length >= 3 && (p[1]+'/'+p[2]) === dashFilter.value;
    });
  }
  if (dashFilter.type === 'customer') return allData.filter(d => d.customer === dashFilter.value);
  return allData;
}

/** Clear dashboard filter */
function clearDashFilter() {
  dashFilter = null;
  closeDrill();
  updateDashboard();
}

/** Sparkline renderer — mini bar chart inside stat cards */
function renderSparkline(containerId, data) {
  const el = document.getElementById(containerId);
  if (!el) return;
  const m = {};
  data.forEach(d => {
    const p = d.date.split('/'); if (p.length < 3) return;
    const k = p[1]+'/'+p[2];
    m[k] = (m[k]||0) + 1;
  });
  const vals = Object.entries(m).sort((a,b)=>new Date('01/'+a[0])-new Date('01/'+b[0])).map(([,v])=>v);
  if (!vals.length) { el.innerHTML = ''; return; }
  const mx = Math.max(...vals);
  el.innerHTML = vals.map(v => `<div class="spark-bar" style="height:${Math.max(8,v/mx*100)}%"></div>`).join('');
}

/** Donut chart for branch distribution */
function renderBranchDonut(data) {
  const wrap = document.getElementById('branchDonut');
  if (!wrap) return;
  const bMap = {};
  data.forEach(d => { bMap[d.branch] = (bMap[d.branch]||0) + d.gross; });
  const entries = Object.entries(bMap).sort((a,b) => b[1]-a[1]);
  const total = entries.reduce((s,e)=>s+e[1],0);
  if (!total) { wrap.innerHTML = '<div style="color:var(--muted);padding:20px">No data</div>'; return; }

  /* Build SVG donut */
  const cx=70, cy=70, r=54, sw=18;
  let angle = -90;
  let paths = '';
  entries.forEach(([name, val]) => {
    const pct = val / total;
    const sweep = pct * 360;
    const large = sweep > 180 ? 1 : 0;
    const a1 = angle * Math.PI / 180;
    const a2 = (angle + sweep) * Math.PI / 180;
    const x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1);
    const x2 = cx + r * Math.cos(a2), y2 = cy + r * Math.sin(a2);
    paths += `<path d="M${x1},${y1} A${r},${r} 0 ${large},1 ${x2},${y2}"
      fill="none" stroke="${getBranchColor(name)}" stroke-width="${sw}"
      style="cursor:pointer;transition:stroke-width .15s,opacity .15s"
      onmouseover="this.style.strokeWidth='${sw+4}'" onmouseout="this.style.strokeWidth='${sw}'"
      onclick="dashBranchClick('${esc(name)}')" />`;
    angle += sweep;
  });

  const svg = `<svg class="donut-svg" viewBox="0 0 140 140">
    ${paths}
    <text x="70" y="68" text-anchor="middle" class="donut-center">${fmtK(total)}</text>
    <text x="70" y="82" text-anchor="middle" class="donut-center-sub">TOTAL GROSS</text>
  </svg>`;

  const legend = entries.map(([name, val]) => `
    <div class="donut-legend-item" onclick="dashBranchClick('${esc(name)}')">
      <div class="donut-legend-dot" style="background:${getBranchColor(name)}"></div>
      <div class="donut-legend-name">${esc(name)}</div>
      <div class="donut-legend-val">${fmtK(val)} <span style="color:var(--muted);font-size:10px">(${(val/total*100).toFixed(1)}%)</span></div>
    </div>`).join('');

  wrap.innerHTML = svg + `<div class="donut-legend">${legend}</div>`;
}

/* ── Dashboard Click Handlers ────────────────────────────── */

/** Click on a stat card → open drill panel with relevant info */
function dashStatClick(type) {
  /* Highlight card */
  document.querySelectorAll('.stat-card').forEach(c => c.classList.remove('active-filter'));
  const card = document.getElementById('kpi-' + type);
  if (card) card.classList.add('active-filter');

  let title = '', rows = [], headers = [], stats = [];

  if (type === 'total') {
    title = '📄 All Records Summary';
    const bMap = {};
    allData.forEach(d => {
      if (!bMap[d.branch]) bMap[d.branch] = { count:0, gross:0, customers:new Set() };
      bMap[d.branch].count++;
      bMap[d.branch].gross += d.gross;
      bMap[d.branch].customers.add(d.customer);
    });
    headers = '<tr><th>Branch</th><th>Records</th><th>Customers</th><th class="dt-num">Gross</th></tr>';
    rows = Object.entries(bMap).map(([n,v]) => `<tr>
      <td><span class="badge badge-branch" style="background:${getBranchColor(n)}">${esc(n)}</span></td>
      <td>${v.count}</td><td>${v.customers.size}</td>
      <td class="dt-num">${fmt(v.gross)}</td></tr>`);
    stats = [
      { l:'Total rows', v:allData.length },
      { l:'Credit Notes', v:new Set(allData.map(d=>d.creditNote)).size },
      { l:'Customers', v:new Set(allData.map(d=>d.customer.toLowerCase())).size },
      { l:'Total Gross', v:fmtK(allData.reduce((s,d)=>s+d.gross,0)) }
    ];
  }
  else if (type === 'clean') {
    title = '✅ Clean Rows Breakdown';
    const clean = allData.filter(d => !d._flags.length);
    const flagged = allData.filter(d => d._flags.length);
    headers = '<tr><th>Status</th><th>Count</th><th class="dt-num">Gross</th><th>% of Total</th></tr>';
    rows = [
      `<tr><td><span class="badge badge-ok">Clean</span></td><td>${clean.length}</td><td class="dt-num">${fmtK(clean.reduce((s,d)=>s+d.gross,0))}</td><td>${(clean.length/allData.length*100).toFixed(1)}%</td></tr>`,
      `<tr><td><span class="badge badge-error">Flagged</span></td><td>${flagged.length}</td><td class="dt-num">${fmtK(flagged.reduce((s,d)=>s+d.gross,0))}</td><td>${(flagged.length/allData.length*100).toFixed(1)}%</td></tr>`,
    ];
    /* Per branch clean % */
    const branches = [...new Set(allData.map(d=>d.branch))];
    branches.forEach(b => {
      const bAll = allData.filter(d=>d.branch===b);
      const bClean = bAll.filter(d=>!d._flags.length);
      rows.push(`<tr><td><span class="badge badge-branch" style="background:${getBranchColor(b)}">${esc(b)}</span></td><td>${bClean.length} / ${bAll.length}</td><td class="dt-num">${fmtK(bClean.reduce((s,d)=>s+d.gross,0))}</td><td>${(bClean.length/bAll.length*100).toFixed(1)}%</td></tr>`);
    });
    stats = [
      { l:'Clean', v:clean.length },
      { l:'Flagged', v:flagged.length },
      { l:'Clean %', v:(clean.length/allData.length*100).toFixed(1)+'%' },
      { l:'Clean Gross', v:fmtK(clean.reduce((s,d)=>s+d.gross,0)) }
    ];
  }
  else if (type === 'branches') {
    title = '🏢 Branch Comparison';
    const bMap = {};
    allData.forEach(d => {
      if (!bMap[d.branch]) bMap[d.branch] = { count:0, gross:0, customers:new Set(), cns:new Set(), avgGross:0 };
      bMap[d.branch].count++;
      bMap[d.branch].gross += d.gross;
      bMap[d.branch].customers.add(d.customer);
      bMap[d.branch].cns.add(d.creditNote);
    });
    headers = '<tr><th>Branch</th><th>Records</th><th>CNs</th><th>Customers</th><th class="dt-num">Gross</th><th class="dt-num">Avg/CN</th></tr>';
    rows = Object.entries(bMap).sort((a,b)=>b[1].gross-a[1].gross).map(([n,v]) => {
      const avg = v.cns.size ? v.gross/v.cns.size : 0;
      return `<tr onclick="dashBranchClick('${esc(n)}')" style="cursor:pointer">
        <td><span class="badge badge-branch" style="background:${getBranchColor(n)}">${esc(n)}</span></td>
        <td>${v.count}</td><td>${v.cns.size}</td><td>${v.customers.size}</td>
        <td class="dt-num">${fmt(v.gross)}</td><td class="dt-num">${fmt(avg)}</td></tr>`;
    });
    stats = Object.entries(bMap).map(([n,v]) => ({ l:n, v:fmtK(v.gross) })).slice(0,4);
  }
  else if (type === 'grossfixed') {
    title = '🔧 Gross Corrections Detail';
    const fixed = allData.filter(d => d._flags.includes('gross-computed')||d._flags.includes('gross-corrected'));
    headers = '<tr><th>Credit Note</th><th>Customer</th><th>Branch</th><th>Issue</th><th class="dt-num">Gross</th></tr>';
    rows = fixed.slice(0,30).map(d => `<tr>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(d.creditNote)}</td>
      <td style="max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(d.customer)}</td>
      <td><span class="badge badge-branch" style="background:${getBranchColor(d.branch)}">${esc(d.branch)}</span></td>
      <td>${d._flags.includes('gross-corrected') ? '<span class="badge badge-gross-wrong">Corrected</span>' : '<span class="badge badge-fixed">Computed</span>'}</td>
      <td class="dt-num">${fmt(d.gross)}</td></tr>`);
    if (fixed.length > 30) rows.push(`<tr><td colspan="5" style="text-align:center;color:var(--muted)">…and ${fixed.length-30} more</td></tr>`);
    stats = [
      { l:'Total Fixed', v:fixed.length },
      { l:'Computed', v:fixed.filter(d=>d._flags.includes('gross-computed')).length },
      { l:'Corrected', v:fixed.filter(d=>d._flags.includes('gross-corrected')).length },
      { l:'Fixed Gross', v:fmtK(fixed.reduce((s,d)=>s+d.gross,0)) }
    ];
  }
  else if (type === 'errors') {
    title = '⚠️ Error Details';
    const errs = allData.filter(d => d._flags.includes('error'));
    headers = '<tr><th>Credit Note</th><th>Customer</th><th>Branch</th><th>Date</th><th class="dt-num">Gross</th></tr>';
    rows = errs.slice(0,30).map(d => `<tr>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(d.creditNote)}</td>
      <td style="max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(d.customer)}</td>
      <td><span class="badge badge-branch" style="background:${getBranchColor(d.branch)}">${esc(d.branch)}</span></td>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(d.date)}</td>
      <td class="dt-num">${fmt(d.gross)}</td></tr>`);
    if (errs.length > 30) rows.push(`<tr><td colspan="5" style="text-align:center;color:var(--muted)">…and ${errs.length-30} more</td></tr>`);
    stats = [
      { l:'Error Rows', v:errs.length },
      { l:'Error Gross', v:fmtK(errs.reduce((s,d)=>s+d.gross,0)) },
      { l:'Branches', v:[...new Set(errs.map(d=>d.branch))].length },
      { l:'% of Total', v:allData.length?(errs.length/allData.length*100).toFixed(1)+'%':'0%' }
    ];
  }

  openDrill(title, headers, rows, stats);
}

/** Click on a customer bar → drill into that customer */
function drillCustomer(name) {
  const cRows = allData.filter(d => d.customer === name);
  if (!cRows.length) return;

  const title = '👤 ' + name;
  const headers = '<tr><th>Date</th><th>Credit Note</th><th>Branch</th><th>Narration</th><th class="dt-num">Qty</th><th class="dt-num">Gross</th></tr>';
  const rows = cRows.sort((a,b)=>dp(b.date)-dp(a.date)).slice(0,30).map(d => `<tr>
    <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(d.date)}</td>
    <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(d.creditNote)}</td>
    <td><span class="badge badge-branch" style="background:${getBranchColor(d.branch)}">${esc(d.branch)}</span></td>
    <td style="max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(d.narration)||'—'}</td>
    <td class="dt-num">${fmtQty(d.qty)}</td>
    <td class="dt-num">${fmt(d.gross)}</td></tr>`);
  if (cRows.length > 30) rows.push(`<tr><td colspan="6" style="text-align:center;color:var(--muted)">…and ${cRows.length-30} more</td></tr>`);

  const totG = cRows.reduce((s,d)=>s+d.gross,0);
  const stats = [
    { l:'Credit Notes', v:new Set(cRows.map(d=>d.creditNote)).size },
    { l:'Total Rows', v:cRows.length },
    { l:'Branches', v:[...new Set(cRows.map(d=>d.branch))].join(', ') },
    { l:'Total Gross', v:fmtK(totG) }
  ];
  openDrill(title, headers, rows, stats);

  /* Also set dashboard filter to this customer */
  dashFilter = { type:'customer', value:name };
  updateDashboard();
}

/** Click on a branch → filter dashboard to that branch */
function dashBranchClick(name) {
  if (dashFilter && dashFilter.type === 'branch' && dashFilter.value === name) {
    clearDashFilter(); return;
  }
  dashFilter = { type:'branch', value:name };
  closeDrill();
  updateDashboard();
}

/** Click on a month card → filter dashboard to that month */
function dashMonthClick(monthKey) {
  if (dashFilter && dashFilter.type === 'month' && dashFilter.value === monthKey) {
    clearDashFilter(); return;
  }
  const parts = monthKey.split('/');
  const label = new Date(+parts[1], +parts[0]-1).toLocaleString('default',{month:'long',year:'numeric'});
  dashFilter = { type:'month', value:monthKey, label:label };
  closeDrill();
  updateDashboard();
}

/** Open drill-down panel */
function openDrill(title, headers, rows, stats) {
  const dt = document.getElementById('drillTitle');
  const th = document.getElementById('drillThead');
  const tb = document.getElementById('drillTbody');
  const ds = document.getElementById('drillStats');
  const dp = document.getElementById('drillPanel');
  
  if (dt) dt.innerHTML = title;
  if (th) th.innerHTML = headers;
  if (tb) tb.innerHTML = rows.join('');
  if (ds) ds.innerHTML = stats.map(s =>
    `<div class="drill-mini"><div class="drill-mini-label">${s.l}</div><div class="drill-mini-value">${s.v}</div></div>`
  ).join('');
  if (dp) dp.classList.add('open');
}

/** Close drill-down panel */
function closeDrill() {
  const dp = document.getElementById('drillPanel');
  if (dp) dp.classList.remove('open');
  document.querySelectorAll('.stat-card').forEach(c => c.classList.remove('active-filter'));
}

/* ══════════════════════════════════════════════════════════════════
   ANALYTICS
═══════════════════════════════════════════════════════════════════ */
function renderAnalytics() {
  if (!allData.length) return;
  const totGross = allData.reduce((s,d) => s+d.gross, 0);
  const avg      = totGross / allData.length;
  const maxRec   = allData.reduce((a,b) => b.gross>a.gross?b:a, allData[0]);
  const dates    = allData.map(d => dp(d.date)).filter(Boolean);
  const minD = new Date(Math.min(...dates)), maxD = new Date(Math.max(...dates));
  const branches = [...new Set(allData.map(d => d.branch))];

  const aAvg = document.getElementById('a-avg');
  if (aAvg) aAvg.textContent = fmtK(avg);
  const aMax = document.getElementById('a-max');
  if (aMax) aMax.textContent = fmt(maxRec.gross);
  const aMaxCust = document.getElementById('a-max-cust');
  if (aMaxCust) aMaxCust.textContent = maxRec.customer.slice(0,32) + ' (' + maxRec.branch + ')';
  const aRange = document.getElementById('a-range');
  if (aRange) aRange.textContent = minD.toLocaleDateString('en-GB') + ' – ' + maxD.toLocaleDateString('en-GB');
  const aBranches = document.getElementById('a-branches');
  if (aBranches) aBranches.textContent = branches.length;
  const aTotal = document.getElementById('a-total');
  if (aTotal) aTotal.textContent = fmtK(totGross);

  /* Top 15 customers */
  const cMap = {};
  allData.forEach(d => { cMap[d.customer] = (cMap[d.customer]||0) + d.gross; });
  const topC = Object.entries(cMap).sort((a,b) => b[1]-a[1]).slice(0,15);
  const cMax = topC[0] ? topC[0][1] : 1;
  const bcc = document.getElementById('bigCustomerChart');
  if (bcc) {
    bcc.innerHTML = topC.map(([name,val]) => `
      <div class="bar-row">
        <div class="bar-label" title="${esc(name)}">${esc(name.slice(0,22))}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${(val/cMax*100).toFixed(1)}%"></div></div>
        <div class="bar-val">${fmtK(val)}</div>
      </div>`).join('');
  }

  /* Monthly gross bar */
  const mo = {};
  allData.forEach(d => {
    const p = d.date.split('/'); if (p.length < 3) return;
    const k = p[1]+'/'+p[2];
    if (!mo[k]) mo[k] = { count:0, gross:0, label: new Date(+p[2],+p[1]-1).toLocaleString('default',{month:'short',year:'2-digit'}) };
    mo[k].count++; mo[k].gross += d.gross;
  });
  const moS = Object.entries(mo).sort((a,b)=>new Date('01/'+a[0])-new Date('01/'+b[0]));
  const mMax = Math.max(...moS.map(([,v])=>v.gross),1);
  const mbc = document.getElementById('monthBarChart');
  if (mbc) {
    mbc.innerHTML = moS.map(([,v]) => `
      <div class="bar-row">
        <div class="bar-label">${v.label}</div>
        <div class="bar-track"><div class="bar-fill green" style="width:${(v.gross/mMax*100).toFixed(1)}%"></div></div>
        <div class="bar-val">${fmtK(v.gross)}</div>
      </div>`).join('');
  }

  /* Branch bar chart */
  const bMap = {};
  allData.forEach(d => { bMap[d.branch] = (bMap[d.branch]||0) + d.gross; });
  const bEntries = Object.entries(bMap).sort((a,b) => b[1]-a[1]);
  const bMax = bEntries[0] ? bEntries[0][1] : 1;
  const bbc = document.getElementById('branchBarChart');
  if (bbc) {
    bbc.innerHTML = bEntries.map(([name,val]) => `
      <div class="bar-row">
        <div class="bar-label">${esc(name)}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${(val/bMax*100).toFixed(1)}%;background:${getBranchColor(name)}"></div></div>
        <div class="bar-val">${fmtK(val)}</div>
      </div>`).join('');
  }
}

/* ══════════════════════════════════════════════════════════════════
   CLEAN LOG
═══════════════════════════════════════════════════════════════════ */
function updateCleanLog() {
  const totGross  = allData.reduce((s,d) => s+d.gross, 0);
  const unique    = new Set(allData.map(d => d.customer.trim().toLowerCase())).size;
  const uniqueV   = new Set(allData.map(d => d.creditNote)).size;
  const branches  = [...new Set(allData.map(d => d.branch))];

  const cst = document.getElementById('cleanSummaryTable');
  if (cst) {
    cst.innerHTML = `
      <table class="issues-table">
        <tr><th>Metric</th><th>Value</th></tr>
        <tr><td>File name</td><td>${esc(currentFile)}</td></tr>
        <tr><td>Sheets detected</td><td><strong>${sheetNames.join(', ')}</strong></td></tr>
        <tr><td>Total rows before cleaning</td><td><strong>${rawParsed.length}</strong></td></tr>
        <tr><td>Output rows after cleaning</td><td><strong>${allData.length}</strong></td></tr>
        ${sheetNames.map(n => `<tr><td>↳ ${esc(n)}</td><td>${allData.filter(d=>d.branch===n).length} rows</td></tr>`).join('')}
        <tr><td>Unique credit notes</td><td>${uniqueV}</td></tr>
        <tr><td>Cancelled vouchers excluded</td><td>${cancelledCount}</td></tr>
        <tr><td>Duplicate rows removed</td><td>${cleanStats.dupes}</td></tr>
        <tr><td>Gross values fixed / computed</td><td>${cleanStats.grossFixed}</td></tr>
        <tr><td>Gross values corrected (were wrong)</td><td>${cleanStats.grossWrong}</td></tr>
        <tr><td>Errors found</td><td>${errorLog.length}</td></tr>
        <tr><td>Unique customers</td><td>${unique}</td></tr>
        <tr><td>Total gross value</td><td><strong>${fmt(totGross)}</strong></td></tr>
        <tr><td>Branches</td><td>${branches.map(b => '<span class="badge badge-branch" style="background:'+getBranchColor(b)+'">'+esc(b)+'</span>').join(' ')}</td></tr>
      </table>`;
  }

  const flags = [];
  const missingRate  = allData.filter(d => d.unitRate === 0).length;
  const missingVal   = allData.filter(d => d.netRate === 0).length;
  const missingGross = allData.filter(d => d.gross === 0).length;
  if (missingRate)   flags.push({ label:'Unit Rate = 0', count:missingRate, sev:'info' });
  if (missingVal)    flags.push({ label:'Net Rate = ₹0', count:missingVal, sev:'err' });
  if (missingGross)  flags.push({ label:'Gross = ₹0', count:missingGross, sev:'warn' });
  if (cleanStats.grossWrong) flags.push({ label:'Gross values corrected', count:cleanStats.grossWrong, sev:'warn' });
  if (cleanStats.dupes)      flags.push({ label:'Duplicate rows removed', count:cleanStats.dupes, sev:'warn' });

  const qf = document.getElementById('qualityFlags');
  if (qf) {
    qf.innerHTML = flags.length
      ? `<table class="issues-table">
          <tr><th>Issue</th><th>Count</th><th>Status</th></tr>
          ${flags.map(f=>`<tr><td>${f.label}</td><td>${f.count}</td><td><span class="issue-badge ib-${f.sev==='err'?'err':f.sev==='warn'?'warn':'ok'}">${f.sev==='err'?'Review':f.sev==='warn'?'Check':'Info'}</span></td></tr>`).join('')}
        </table>`
      : '<div style="color:var(--green);font-size:13px;font-weight:500">✅ No quality issues found</div>';
  }
}

/* ══════════════════════════════════════════════════════════════════
   EXPORT — Single Excel with "Final_Data" sheet
═══════════════════════════════════════════════════════════════════ */
function exportExcel() {
  if (!allData.length) { showToast('⚠️ No data to export'); return; }

  const wsData = [
    ['Date','Customer','Credit Note Number','Narration','Quantity','Unit Rate','Net Rate','Gross','Branch','Status Flags'],
    ...allData.map(d => [
      d.date, d.customer, d.creditNote, d.narration,
      d.qty, d.unitRate, d.netRate, d.gross, d.branch,
      d._flags.join(', ')
    ])
  ];

  const tot = allData.reduce((s,d) => ({ netRate:s.netRate+d.netRate, gross:s.gross+d.gross }), { netRate:0, gross:0 });
  wsData.push(['TOTAL','','','','','', Math.round(tot.netRate*100)/100, Math.round(tot.gross*100)/100, '', '']);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [{wch:12},{wch:42},{wch:18},{wch:50},{wch:10},{wch:12},{wch:14},{wch:16},{wch:14},{wch:18}];
  XLSX.utils.book_append_sheet(wb, ws, 'Final_Data');

  /* Sheet 2 — Error Log */
  if (errorLog.length) {
    const wsErrors = [
      ['Severity','Credit Note No.','Customer','Date','Branch','Issue'],
      ...errorLog.map(e => [e.sev.toUpperCase(), e.voucher, e.customer, e.date, e.branch, e.msg])
    ];
    const ws2 = XLSX.utils.aoa_to_sheet(wsErrors);
    ws2['!cols'] = [{wch:10},{wch:18},{wch:38},{wch:12},{wch:14},{wch:60}];
    XLSX.utils.book_append_sheet(wb, ws2, 'Error_Log');
  }

  const fn = `credit_notes_cleaned_${new Date().toISOString().slice(0,10)}.xlsx`;
  XLSX.writeFile(wb, fn);
  showToast('📥 Downloaded ' + fn);
}

/* ══════════════════════════════════════════════════════════════════
   UTILITY FUNCTIONS
═══════════════════════════════════════════════════════════════════ */
function fmt(n, sym='₹') {
  if (!n || n === 0) return '—';
  return sym + n.toLocaleString('en-IN', { minimumFractionDigits:2, maximumFractionDigits:2 });
}
function fmt2(n) { return n.toLocaleString('en-IN', { minimumFractionDigits:2, maximumFractionDigits:2 }); }
function fmtQty(n) { if (!n || n === 0) return '—'; return n.toLocaleString('en-IN', { maximumFractionDigits:4 }); }
function fmtK(n) {
  if (n >= 10000000) return '₹' + (n/10000000).toFixed(2) + 'Cr';
  if (n >= 100000)   return '₹' + (n/100000).toFixed(2) + 'L';
  return '₹' + Math.round(n).toLocaleString('en-IN');
}
function dp(s) {
  try { const [d,mo,y] = s.split('/'); return new Date(+y, +mo-1, +d).getTime(); } catch { return 0; }
}
function esc(s) {
  if (!s) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

let _tt;
function showToast(msg) {
  clearTimeout(_tt);
  const toast = document.getElementById('toast');
  const toastMsg = document.getElementById('toastMsg');
  if (toastMsg) toastMsg.textContent = msg;
  if (toast) {
    toast.classList.add('show');
    _tt = setTimeout(() => toast.classList.remove('show'), 3200);
  }
}
