'use strict';

/* ═══════════════════════════════════════════════════════════════
   MOBILE SIDEBAR TOGGLE
═══════════════════════════════════════════════════════════════ */
function toggleSidebar() {
  const sidebar = document.querySelector('.sidebar');
  const overlay = document.getElementById('sidebarOverlay');
  sidebar.classList.toggle('open');
  overlay.classList.toggle('open');
}
/** Auto-close sidebar on mobile when a nav item is clicked */
document.querySelectorAll('.nav-item').forEach(item => {
  item.addEventListener('click', () => {
    if (window.innerWidth <= 768) {
      document.querySelector('.sidebar').classList.remove('open');
      document.getElementById('sidebarOverlay').classList.remove('open');
    }
  });
});
/** Close sidebar on export nav clicks too */
document.querySelectorAll('.nav .nav-item[onclick*="export"]').forEach(item => {
  item.addEventListener('click', () => {
    if (window.innerWidth <= 768) {
      document.querySelector('.sidebar').classList.remove('open');
      document.getElementById('sidebarOverlay').classList.remove('open');
    }
  });
});

/* ═══════════════════════════════════════════════════════════════
   STATE — All app data lives here
═══════════════════════════════════════════════════════════════ */
let rawSheetRows = [];  // raw rows from SheetJS (array of arrays)
let rawParsed    = [];  // after parseData() — unprocessed item rows
let allData      = [];  // after cleaning pipeline
let filtered     = [];  // after applyFilters()
let errorLog     = [];  // structured error entries
let currentSort  = { col:'date', dir:'asc' };
let cancelledCount = 0;
let currentFile  = '';
let detectedHeaders = []; // auto-detected column headers
let cleanStats   = { dupes:0, grossFixed:0, grossWrong:0, missingFields:0, janRepeats:0 };

/* ═══════════════════════════════════════════════════════════════
   NAVIGATION — switchPage()
   Handles sidebar navigation and page visibility
═══════════════════════════════════════════════════════════════ */
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
    document.getElementById('page-' + id).classList.add('active');
    const titles = {
      dashboard:'Dashboard', preview:'Data Preview', records:'Records',
      errors:'Error Log', analytics:'Analytics', products:'Product Cleaner', clean:'Clean Log'
    };
    document.getElementById('topbarTitle').textContent = titles[id] || id;
    document.getElementById('exportBtn').style.display = allData.length ? 'flex' : 'none';
    document.getElementById('recordPill').style.display = allData.length ? 'flex' : 'none';
    document.getElementById('recordPill').textContent = allData.length + ' records';
    if (id === 'analytics') renderAnalytics();
    if (id === 'records') applyFilters();
    if (id === 'errors') renderErrorLog();
  }
}
switchPage('upload');

/* ═══════════════════════════════════════════════════════════════
   FILE HANDLING — parseFile()
   Uses FileReader API and SheetJS to read .xlsx / .csv files
═══════════════════════════════════════════════════════════════ */
const dropZone = document.getElementById('dropZone');
dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', ()  => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.classList.remove('dragover');
  const f = e.dataTransfer.files[0];
  if (f) processFile(f);
});

/** Handle file input change event */
function handleFile(e) {
  const f = e.target.files[0];
  if (f) processFile(f);
}

/** Main file entry point — reads file with FileReader, passes to parser */
function processFile(file) {
  currentFile = file.name;
  setProgress(true, 'Reading file…', 25);

  const reader = new FileReader();
  reader.onload = e => {
    setProgress(true, 'Parsing rows…', 60);
    try {
      const wb  = XLSX.read(e.target.result, { type:'binary', cellDates:true });
      const ws  = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header:1, raw:false, defval:'' });

      /* Store raw rows for preview */
      rawSheetRows = raw;

      setProgress(true, 'Detecting columns…', 70);
      setTimeout(() => {
        /* Auto-detect headers */
        autoDetectHeaders(raw);
        /* Render original data preview */
        renderOriginalPreview(raw);

        setProgress(true, 'Cleaning data…', 80);
        setTimeout(() => {
          parseData(raw);
          setProgress(true, 'Removing duplicates…', 95);
          setTimeout(() => {
            /* Render cleaned preview */
            renderCleanedPreview();
            setProgress(false);
          }, 100);
        }, 50);
      }, 50);
    } catch (err) {
      setProgress(false);
      alert('Error reading file: ' + err.message);
    }
  };
  reader.readAsBinaryString(file);
}

/** Update progress bar UI */
function setProgress(show, label='', pct=0) {
  const wrap  = document.getElementById('progressWrap');
  const inner = document.getElementById('progressInner');
  const lbl   = document.getElementById('progressLabel');
  wrap.style.display  = show ? 'block' : 'none';
  lbl.textContent     = label;
  inner.style.width   = pct + '%';
}

/* ═══════════════════════════════════════════════════════════════
   AUTO-DETECT COLUMN HEADERS
   Scans first 20 rows to find the header row and map columns
═══════════════════════════════════════════════════════════════ */
function autoDetectHeaders(rows) {
  /* Known header patterns for credit note files */
  const headerKeywords = ['date','particulars','buyer','voucher type','voucher no',
    'narration','quantity','rate','value','gross total','credit note'];

  let headerRow = -1;
  detectedHeaders = [];

  /* Scan first 20 rows to find the header row */
  for (let i = 0; i < Math.min(20, rows.length); i++) {
    const row = rows[i];
    if (!row || !row.length) continue;
    const rowLower = row.map(c => String(c || '').trim().toLowerCase());
    /* Count how many header keywords appear in this row */
    const matchCount = headerKeywords.filter(kw =>
      rowLower.some(cell => cell.includes(kw))
    ).length;

    if (matchCount >= 3) {
      headerRow = i;
      detectedHeaders = row.map(c => String(c || '').trim());
      break;
    }
  }

  /* Render column mapper */
  const mapper = document.getElementById('columnMapper');
  const section = document.getElementById('columnMapperSection');

  if (headerRow >= 0 && detectedHeaders.length > 0) {
    section.style.display = '';
    const cols = detectedHeaders.filter(Boolean).slice(0, 12);
    mapper.innerHTML = cols.map((col, i) => `
      <div class="mapper-item">
        <label>Column ${i + 1}</label>
        <div style="font-size:13px;font-weight:500;color:var(--ink);margin-top:2px">${esc(col)}</div>
        <div class="auto-detected">✓ Auto-detected</div>
      </div>
    `).join('');
  } else {
    section.style.display = 'none';
  }
}

/* ═══════════════════════════════════════════════════════════════
   DATA PREVIEW — Render original and cleaned tables
═══════════════════════════════════════════════════════════════ */

/** Render the raw original data preview (first 50 rows) */
function renderOriginalPreview(rows) {
  const table = document.getElementById('previewOriginalTable');
  if (!rows || !rows.length) {
    table.innerHTML = '<thead><tr><th>No data</th></tr></thead><tbody></tbody>';
    return;
  }

  /* Find data start (skip company header rows) */
  let start = 0;
  const CN_RE = /^(CN|CNG|CNS)\//i;
  for (let i = 0; i < Math.min(20, rows.length); i++) {
    const rowLower = rows[i].map(c => String(c || '').trim().toLowerCase());
    if (rowLower.some(c => c.includes('date') || c.includes('voucher'))) {
      start = i;
      break;
    }
  }

  const previewRows = rows.slice(start, start + 51);
  if (!previewRows.length) return;

  /* Build header from first row of preview */
  const headers = previewRows[0].slice(0, 12).map(h => String(h || '').trim());
  let html = '<thead><tr>' + headers.map(h => `<th>${esc(h || '—')}</th>`).join('') + '</tr></thead><tbody>';

  previewRows.slice(1, 51).forEach(row => {
    html += '<tr>' + row.slice(0, 12).map(cell => {
      const val = String(cell || '').trim();
      return `<td>${esc(val.length > 60 ? val.slice(0, 57) + '…' : val) || '<span style="color:var(--muted)">—</span>'}</td>`;
    }).join('') + '</tr>';
  });

  html += '</tbody>';
  table.innerHTML = html;
}

/** Render cleaned data preview (first 50 rows of allData) */
function renderCleanedPreview() {
  const table = document.getElementById('previewCleanedTable');
  if (!allData || !allData.length) {
    table.innerHTML = '<thead><tr><th>No cleaned data yet</th></tr></thead><tbody></tbody>';
    return;
  }

  const headers = ['Date','Customer','Voucher','Item','Qty','Rate','Value','Gross','Status'];
  let html = '<thead><tr>' + headers.map(h => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';

  allData.slice(0, 50).forEach(d => {
    /* Determine row status badge */
    let statusHtml = '<span class="preview-badge" style="background:var(--green-light);color:var(--green)">OK</span>';
    if (d._flags.includes('error')) statusHtml = '<span class="preview-badge" style="background:var(--red-light);color:var(--red)">Error</span>';
    else if (d._flags.includes('gross-corrected')) statusHtml = '<span class="preview-badge" style="background:#fefce8;color:#854d0e">Gross Fixed</span>';
    else if (d._flags.includes('gross-computed')) statusHtml = '<span class="preview-badge" style="background:var(--green-light);color:var(--green)">Gross Added</span>';

    html += `<tr>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(d.date)}</td>
      <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${esc(d.customer)}">${esc(d.customer)}</td>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${esc(d.voucher)}</td>
      <td style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${esc(d.itemName)}">${esc(d.itemName) || '—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${d.qty || '—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${d.rate ? fmt2(d.rate) : '—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${d.value ? fmt2(d.value) : '—'}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:600">${d.gross ? fmt2(d.gross) : '—'}</td>
      <td>${statusHtml}</td>
    </tr>`;
  });

  if (allData.length > 50) {
    html += `<tr><td colspan="9" style="text-align:center;color:var(--muted);padding:12px">… and ${allData.length - 50} more rows</td></tr>`;
  }

  html += '</tbody>';
  table.innerHTML = html;
}

/** Switch between Original / Cleaned preview tabs */
function switchPreviewTab(tab) {
  document.querySelectorAll('#page-preview .tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('#page-preview .tab-pane').forEach(p => p.classList.remove('active'));
  event.target.classList.add('active');
  document.getElementById('preview-' + tab).classList.add('active');
}

/* ═══════════════════════════════════════════════════════════════
   PARSER — parseData()
   Converts raw sheet rows → item-level flat records (rawParsed).
   Then feeds into the cleaning pipeline.
═══════════════════════════════════════════════════════════════ */
function parseData(rows) {
  const CN_RE    = /^(CN|CNG|CNS)\//i;
  const CANCELLED = /\(cancelled/i;
  const records  = [];
  cancelledCount = 0;

  /** Safe number extractor — handles "23.0000 NOS", "290.00/NOS" etc. */
  const cn = v => {
    if (v === null || v === undefined || String(v).trim() === '') return 0;
    const s = String(v).replace(/[^0-9.-]/g, '');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  };

  /* ── Parse the ENTIRE file ── */
  let i = 0;
  while (i < rows.length) {
    const r       = rows[i];
    const voucher = String(r[4] || '').trim();
    const dateRaw = String(r[0] || '').trim();

    /* Must be a voucher main row */
    if (!CN_RE.test(voucher) || !dateRaw || dateRaw.toLowerCase() === 'date') { i++; continue; }
    if (isNaN(new Date(dateRaw).getTime()) && !/\d{4}/.test(dateRaw))          { i++; continue; }

    const particulars = String(r[1] || '').trim();
    const buyer       = String(r[2] || '').trim();
    if (CANCELLED.test(particulars) || CANCELLED.test(buyer)) { cancelledCount++; i++; continue; }

    /* Date normalisation */
    let dateStr = '';
    try {
      const d = new Date(dateRaw);
      dateStr = isNaN(d.getTime()) ? dateRaw : d.toLocaleDateString('en-GB');
    } catch { dateStr = dateRaw; }

    /* Header fields */
    const customer  = buyer.length >= particulars.length ? buyer : particulars;
    const vtype     = String(r[3] || '').trim();
    const narration = String(r[5] || '').replace(/_x000D_\\n/g,' ').replace(/_x000D_\n/g,' ').trim();
    const mainValue = cn(r[8]);
    let   mainGross = cn(r[9]);
    if (mainGross === 0 && mainValue > 0) mainGross = mainValue;

    /* Collect item sub-rows */
    let j = i + 1;
    const rawItems = [];
    while (j < rows.length) {
      const nr = rows[j];
      const nv = String(nr[4] || '').trim();
      const nd = String(nr[0] || '').trim();
      const np = String(nr[1] || '').trim().toLowerCase();
      if (CN_RE.test(nv) && nd && !isNaN(new Date(nd).getTime())) break;
      if (np === 'grand total') break;
      const iName = String(nr[1] || '').trim();
      if (iName && np !== 'grand total') {
        rawItems.push({ name: iName, qty: cn(nr[6]), rate: cn(nr[7]), value: cn(nr[8]) });
      }
      j++;
    }

    /* Emit one record per item — or one fallback record if no detail rows */
    if (rawItems.length > 0) {
      const totalItemVal = rawItems.reduce((s, it) => s + it.value, 0);
      rawItems.forEach((it, idx) => {
        let itemGross = 0;
        if (totalItemVal > 0 && it.value > 0) {
          itemGross = Math.round(it.value / totalItemVal * mainGross * 100) / 100;
        } else if (rawItems.length === 1) {
          itemGross = mainGross;
        } else if (totalItemVal === 0) {
          itemGross = Math.round(mainGross / rawItems.length * 100) / 100;
        }
        records.push({ date:dateStr, customer, vtype, voucher, narration,
          itemName:it.name, qty:it.qty, rate:it.rate, value:it.value, gross:itemGross,
          _flags:[] });
      });
    } else {
      records.push({ date:dateStr, customer, vtype, voucher, narration,
        itemName:'', qty:cn(r[6]), rate:0, value:mainValue, gross:mainGross,
        _flags:[] });
    }
    i = j;
  }

  if (!records.length) { showToast('⚠️ No records found — check file format'); return; }

  rawParsed = records.map(r => ({ ...r, _flags:[] }));
  runCleaningPipeline();

  document.getElementById('fileNameDisplay').textContent = currentFile;
  document.getElementById('rerunBtn').disabled = false;
  switchPage('dashboard');
  showToast('✅ Loaded ' + allData.length + ' item-level rows');
}

/* ═══════════════════════════════════════════════════════════════
   CLEANING PIPELINE — cleanDataPipeline()
   Orchestrates all cleaning modules in order
═══════════════════════════════════════════════════════════════ */
function runCleaningPipeline() {
  /* Read control toggles */
  const opts = {
    dedup   : document.getElementById('opt-dedup').checked,
    gross   : document.getElementById('opt-gross').checked,
    product : document.getElementById('opt-product').checked,
    trim    : document.getElementById('opt-trim').checked,
    validate: document.getElementById('opt-validate').checked,
    jancheck: document.getElementById('opt-jancheck').checked,
  };

  /* Reset stats */
  cleanStats = { dupes:0, grossFixed:0, grossWrong:0, missingFields:0, janRepeats:0 };
  errorLog   = [];

  /* 1. Start from raw parsed rows (fresh copy each run) */
  let data = rawParsed.map(r => ({ ...r, _flags:[] }));

  /* 2. Trim spaces */
  if (opts.trim) data = trimSpaces(data);

  /* 3. Gross fix & validation */
  if (opts.gross) data = calculateGross(data);

  /* 4. Validate required fields */
  if (opts.validate) data = validateData(data);

  /* 5. Jan repeat detection */
  if (opts.jancheck) data = detectJanRepeats(data);

  /* 6. Duplicate detection & removal */
  if (opts.dedup) data = detectDuplicates(data);

  /* 7. Sort by date ascending */
  data.sort((a, b) => dp(a.date) - dp(b.date));

  allData  = data;
  filtered = [...allData];

  populateFilters();
  applyFilters();
  updateDashboard();
  updateCleanLog();
  renderProductTable();

  document.getElementById('fileStats').textContent = allData.length + ' item rows';
  document.getElementById('nb-records').textContent = allData.length;
  document.getElementById('nb-errors').textContent  = errorLog.length;
  document.getElementById('multiItemAlert').classList.add('show');
}

/** Re-run with current toggle states */
function rerunCleaning() {
  if (!rawParsed.length) return;
  showToast('🔄 Re-running cleaning pipeline…');
  setTimeout(() => {
    runCleaningPipeline();
    renderCleanedPreview();
    showToast('✅ Cleaning complete — ' + allData.length + ' rows');
  }, 50);
}

/* ═══════════════════════════════════════════════════════════════
   MODULE: trimSpaces()
   Removes leading/trailing whitespace from all text fields
═══════════════════════════════════════════════════════════════ */
function trimSpaces(data) {
  return data.map(r => ({
    ...r,
    customer  : r.customer.trim(),
    voucher   : r.voucher.trim(),
    narration : r.narration.trim(),
    itemName  : r.itemName.trim(),
    vtype     : r.vtype.trim(),
    date      : r.date.trim(),
  }));
}

/* ═══════════════════════════════════════════════════════════════
   MODULE: calculateGross()
   Rule: Gross = Qty × Rate (if both non-zero).
   — If gross is 0/missing → compute and tag "gross-computed"
   — If gross exists but deviates >₹1 from Qty×Rate → correct it
═══════════════════════════════════════════════════════════════ */
function calculateGross(data) {
  return data.map(r => {
    const row = { ...r };
    if (row.qty > 0 && row.rate > 0) {
      const expected = Math.round(row.qty * row.rate * 100) / 100;

      if (!row.gross || row.gross === 0) {
        /* Missing gross — compute it */
        row._grossOriginal = 0;
        row.gross = expected;
        row._flags = [...row._flags, 'gross-computed'];
        cleanStats.grossFixed++;
        pushError('info', row.voucher, row.customer, row.date,
          `Gross was missing — computed as Qty×Rate = ₹${fmt2(expected)}`, 'Gross');

      } else {
        const diff = Math.abs(row.gross - expected);
        if (diff > 1) {
          /* Wrong gross — correct it */
          row._grossOriginal = row.gross;
          row.gross = expected;
          row._flags = [...row._flags, 'gross-corrected'];
          cleanStats.grossFixed++;
          cleanStats.grossWrong++;
          pushError('warn', row.voucher, row.customer, row.date,
            `Gross was ₹${fmt2(row._grossOriginal)} but Qty×Rate = ₹${fmt2(expected)} (diff ₹${fmt2(diff)}) — corrected`, 'Gross');
        }
      }
    }
    return row;
  });
}

/* ═══════════════════════════════════════════════════════════════
   MODULE: validateData()
   Flags rows missing Date, Customer, or Credit Note Number.
   Also flags value = 0 (possible data entry gap).
═══════════════════════════════════════════════════════════════ */
function validateData(data) {
  return data.map(r => {
    const row = { ...r };
    const missing = [];
    if (!row.date)     missing.push('Date');
    if (!row.customer) missing.push('Customer');
    if (!row.voucher)  missing.push('Credit Note Number');

    if (missing.length > 0) {
      row._flags = [...row._flags, 'error'];
      cleanStats.missingFields++;
      pushError('err', row.voucher || '?', row.customer || '?', row.date || '?',
        'Missing required field(s): ' + missing.join(', '), missing.join(', '));
    }
    if (row.value === 0) {
      pushError('warn', row.voucher, row.customer, row.date,
        'Value = ₹0 — possible missing entry', 'Value');
    }
    return row;
  });
}

/* ═══════════════════════════════════════════════════════════════
   MODULE: detectJanRepeats()
   Flags rows whose voucher appears in BOTH a non-January month
   AND January — the Jan entries are carry-overs from a re-export.
   
   KEY: We count unique voucher numbers, not rows. A 3-item
   voucher = 1 unique voucher, so multi-item CNs never falsely flagged.
═══════════════════════════════════════════════════════════════ */
function detectJanRepeats(data) {
  /* Step 1: For each unique voucher, collect the SET of months */
  const voucherMonths = {};
  data.forEach(r => {
    const d = dp(r.date);
    if (!d) return;
    const mo = new Date(d).getMonth(); // 0=Jan
    if (!voucherMonths[r.voucher]) voucherMonths[r.voucher] = new Set();
    voucherMonths[r.voucher].add(mo);
  });

  /* Step 2: A voucher is a Jan repeat if it appears in January (mo=0)
     AND also in at least one other month */
  const janRepeatVouchers = new Set();
  Object.entries(voucherMonths).forEach(([voucher, months]) => {
    if (months.has(0) && months.size > 1) {
      janRepeatVouchers.add(voucher);
    }
  });

  return data.map(r => {
    const row = { ...r };
    if (!janRepeatVouchers.has(r.voucher)) return row;

    /* Only flag the JANUARY rows of these vouchers */
    const d = dp(r.date);
    if (d && new Date(d).getMonth() === 0) {
      if (!row._flags.includes('jan-repeat')) {
        row._flags = [...row._flags, 'jan-repeat'];
        cleanStats.janRepeats++;
        pushError('warn', row.voucher, row.customer, row.date,
          'Jan duplicate: this voucher also exists in a later month — Jan copy removed', 'Date');
      }
    }
    return row;
  });
}

/* ═══════════════════════════════════════════════════════════════
   MODULE: detectDuplicates()
   
   BUSINESS LOGIC (CRITICAL):
   ✔ One Credit Note can have MULTIPLE items → VALID
   ❌ Same Credit Note + SAME Item + SAME Qty + SAME Rate → DUPLICATE
   
   Key = Credit Note Number + Item/Narration + Quantity + Unit Rate
   First occurrence kept; subsequent duplicates removed.
   Jan-repeat rows are ALSO removed here.
═══════════════════════════════════════════════════════════════ */
function detectDuplicates(data) {
  const seen   = new Set();
  const output = [];

  /* First pass: remove jan-repeat rows outright */
  let janRemoved = 0;
  const noJan = data.filter(r => {
    if (r._flags.includes('jan-repeat')) { janRemoved++; return false; }
    return true;
  });
  cleanStats.janRepeats = janRemoved;

  /* Second pass: dedup by composite key on remaining rows */
  noJan.forEach(r => {
    const key = [
      r.voucher,
      r.itemName.toLowerCase().trim(),
      r.qty,
      r.rate
    ].join('||');

    if (seen.has(key)) {
      cleanStats.dupes++;
      pushError('warn', r.voucher, r.customer, r.date,
        'Duplicate row detected (same Credit Note + Item + Qty + Rate)', 'All fields');
    } else {
      seen.add(key);
      output.push(r);
    }
  });
  return output;
}

/* ═══════════════════════════════════════════════════════════════
   MODULE: cleanProductName()
   Extracts structured fields from a raw product name string:
   — Clean_Name  : remaining text after removing brackets & part no.
   — Part_Number : last alphanumeric word before brackets
   — Make        : content of LAST ( ) group
   — Tracker     : content of LAST [ ] group
═══════════════════════════════════════════════════════════════ */
function cleanProductName(raw) {
  if (!raw) return { cleanName:'', partNumber:'', make:'', tracker:'' };

  let str = raw.trim();

  /* extractTracker() — Extract LAST [ ] → tracker */
  let tracker = '';
  const sqMatches = [...str.matchAll(/\[([^\]]*)\]/g)];
  if (sqMatches.length) {
    tracker = sqMatches[sqMatches.length - 1][1].trim();
    str = str.replace(sqMatches[sqMatches.length - 1][0], '').trim();
  }

  /* Extract LAST ( ) → make/brand */
  let make = '';
  const rnMatches = [...str.matchAll(/\(([^)]*)\)/g)];
  if (rnMatches.length) {
    make = rnMatches[rnMatches.length - 1][1].trim();
    str  = str.replace(rnMatches[rnMatches.length - 1][0], '').trim();
  }

  /* Part number = last "word" that contains digits and is alphanumeric */
  let partNumber = '';
  const words = str.split(/\s+/).filter(Boolean);
  for (let i = words.length - 1; i >= 0; i--) {
    if (/^[A-Z0-9][A-Z0-9\/\-.*+#]{1,}$/i.test(words[i]) && /[0-9]/.test(words[i])) {
      partNumber = words[i];
      words.splice(i, 1);
      break;
    }
  }

  /* Clean name = what remains */
  const cleanName = words.join(' ').replace(/\s{2,}/g, ' ').trim();

  return { cleanName, partNumber, make, tracker };
}

/** extractTracker() — Standalone function for tracker extraction */
function extractTracker(raw) {
  if (!raw) return '';
  const sqMatches = [...raw.matchAll(/\[([^\]]*)\]/g)];
  if (sqMatches.length) return sqMatches[sqMatches.length - 1][1].trim();
  return '';
}

/* Live single-input parser for the Product Cleaner page */
function liveParseProduct() {
  const raw = document.getElementById('productInput').value;
  const res = cleanProductName(raw);
  document.getElementById('pr-name').textContent    = res.cleanName    || '—';
  document.getElementById('pr-part').textContent    = res.partNumber   || '—';
  document.getElementById('pr-make').textContent    = res.make         || '—';
  document.getElementById('pr-tracker').textContent = res.tracker      || '—';
}

/* Render product table from allData item names */
function renderProductTable() {
  const seen = new Set();
  const rows = [];
  allData.forEach(d => {
    if (d.itemName && !seen.has(d.itemName)) {
      seen.add(d.itemName);
      rows.push(d.itemName);
    }
  });
  rows.sort();

  const note = document.getElementById('productFileNote');
  const tbody = document.getElementById('productTableBody');

  if (!rows.length) {
    note.textContent = 'No items found in file.';
    tbody.innerHTML  = '<tr><td colspan="5" style="text-align:center;color:var(--muted);padding:20px">No data.</td></tr>';
    return;
  }

  note.textContent = rows.length + ' unique item names parsed from file.';
  tbody.innerHTML  = rows.map(name => {
    const p = cleanProductName(name);
    return `<tr>
      <td title="${esc(name)}">${esc(name.length > 45 ? name.slice(0,42)+'…' : name)}</td>
      <td>${esc(p.cleanName)   || '<span style="color:var(--muted)">—</span>'}</td>
      <td>${esc(p.partNumber)  || '<span style="color:var(--muted)">—</span>'}</td>
      <td>${esc(p.make)        || '<span style="color:var(--muted)">—</span>'}</td>
      <td>${esc(p.tracker)     || '<span style="color:var(--muted)">—</span>'}</td>
    </tr>`;
  }).join('');
}

/* ═══════════════════════════════════════════════════════════════
   ERROR LOG — pushError() & renderErrorLog()
═══════════════════════════════════════════════════════════════ */
function pushError(sev, voucher, customer, date, msg, field) {
  errorLog.push({ sev, voucher, customer, date, msg, field });
}

function renderErrorLog() {
  const list  = document.getElementById('errorList');
  const tbody = document.getElementById('errorTableBody');
  const card  = document.getElementById('errorTableCard');
  const count = document.getElementById('errorTotalCount');

  count.textContent = errorLog.length + ' issue' + (errorLog.length !== 1 ? 's' : '');

  if (!errorLog.length) {
    list.innerHTML = '<div class="no-errors">✅ No data issues detected.</div>';
    card.style.display = 'none';
    return;
  }

  /* Summary list */
  list.innerHTML = errorLog.slice(0, 40).map(e => `
    <div class="error-item">
      <span class="error-sev sev-${e.sev}">${e.sev.toUpperCase()}</span>
      <div>
        <div class="error-msg">${esc(e.msg)}</div>
        <div class="error-ref">${esc(e.voucher)} · ${esc(e.customer)} · ${esc(e.date)}</div>
      </div>
    </div>`).join('') +
    (errorLog.length > 40 ? `<div style="padding:10px 16px;font-size:12px;color:var(--muted)">…and ${errorLog.length - 40} more. Download Excel to see full Error Log sheet.</div>` : '');

  /* Detailed table */
  card.style.display = '';
  tbody.innerHTML = errorLog.map(e => `<tr>
    <td><span class="error-sev sev-${e.sev}">${e.sev.toUpperCase()}</span></td>
    <td style="font-family:'DM Mono',monospace;font-size:11.5px">${esc(e.voucher)}</td>
    <td>${esc(e.customer)}</td>
    <td style="font-family:'DM Mono',monospace;font-size:11.5px">${esc(e.date)}</td>
    <td style="font-size:12px">${esc(e.msg)}</td>
    <td style="font-size:12px;color:var(--muted)">${esc(e.field)}</td>
  </tr>`).join('');
}

/* ═══════════════════════════════════════════════════════════════
   FILTERS & SORT — applyFilters(), sortBy()
═══════════════════════════════════════════════════════════════ */
function populateFilters() {
  const vtypes = [...new Set(allData.map(d => d.vtype).filter(Boolean))].sort();
  const years  = [...new Set(allData.map(d => d.date.split('/').pop()).filter(Boolean))].sort().reverse();
  document.getElementById('filterVType').innerHTML = '<option value="">All types</option>' + vtypes.map(v=>`<option>${v}</option>`).join('');
  document.getElementById('filterYear').innerHTML  = '<option value="">All years</option>'  + years.map(y=>`<option>${y}</option>`).join('');
}

function applyFilters() {
  const q       = document.getElementById('searchBox').value.toLowerCase();
  const vt      = document.getElementById('filterVType').value;
  const yr      = document.getElementById('filterYear').value;
  const status  = document.getElementById('filterStatus').value;
  const sortSel = document.getElementById('filterSort').value;

  filtered = allData.filter(d => {
    const mq = !q || (
      d.customer.toLowerCase().includes(q) ||
      d.voucher.toLowerCase().includes(q)  ||
      d.narration.toLowerCase().includes(q)||
      d.itemName.toLowerCase().includes(q)
    );
    const mv = !vt || d.vtype === vt;
    const my = !yr || d.date.endsWith(yr);
    let ms = true;
    if (status === 'clean')      ms = !d._flags.length;
    if (status === 'dup')        ms = d._flags.includes('dup');
    if (status === 'grossfixed') ms = d._flags.includes('gross-computed') || d._flags.includes('gross-corrected');
    if (status === 'error')      ms = d._flags.includes('error');
    return mq && mv && my && ms;
  });

  const [col, dir] = sortSel.split('-');
  currentSort = { col: col || 'date', dir: dir || 'asc' };
  sortFiltered();
  renderTable();
}

function clearFilters() {
  document.getElementById('searchBox').value   = '';
  document.getElementById('filterVType').value = '';
  document.getElementById('filterYear').value  = '';
  document.getElementById('filterStatus').value = '';
  document.getElementById('filterSort').value  = 'date-asc';
  applyFilters();
}

function sortBy(col) {
  if (currentSort.col === col) currentSort.dir = currentSort.dir === 'asc' ? 'desc' : 'asc';
  else currentSort = { col, dir:'asc' };
  document.getElementById('filterSort').value = '';
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
    if (col === 'voucher')  return m * a.voucher.localeCompare(b.voucher);
    return m * ((a[col] || 0) - (b[col] || 0));
  });
}

function updateSortIcons() {
  ['date','customer','voucher','qty','value','gross'].forEach(c => {
    const el = document.getElementById('sa-' + c);
    if (!el) return;
    el.className = 'sort-arrows' + (currentSort.col === c ? ' ' + currentSort.dir : '');
  });
}

/* ═══════════════════════════════════════════════════════════════
   TABLE RENDER — renderTable()
═══════════════════════════════════════════════════════════════ */
function renderTable() {
  const tbody = document.getElementById('tableBody');
  document.getElementById('filterCount').textContent = filtered.length + ' rows';

  if (!filtered.length) {
    tbody.innerHTML = `<tr><td colspan="11"><div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-title">No records match</div></div></td></tr>`;
    document.getElementById('pageInfo').textContent = '0 rows';
    document.getElementById('filteredTotal').textContent = '';
    return;
  }

  const totVal   = filtered.reduce((s, d) => s + d.value, 0);
  const totGross = filtered.reduce((s, d) => s + d.gross, 0);

  let html = '';
  filtered.forEach((d, i) => {
    /* Row CSS class based on flags */
    let rowClass = '';
    if (d._flags.includes('error'))               rowClass = 'row-error';
    else if (d._flags.includes('dup') || d._flags.includes('jan-repeat')) rowClass = 'row-dup';
    else if (d._flags.includes('gross-corrected')) rowClass = 'row-gross-wrong';
    else if (d._flags.includes('gross-computed'))  rowClass = 'row-gross-fixed';

    const narr     = d.narration.length > 40 ? d.narration.slice(0,37) + '…' : d.narration;
    const itemDisp = d.itemName.length > 28  ? d.itemName.slice(0,25) + '…' : d.itemName;
    const rateCell = d.rate > 0 ? fmt(d.rate) : '<span style="color:var(--muted);font-size:11px">—</span>';

    /* Status badge(s) */
    let statusBadges = '';
    if (!d._flags.length) {
      statusBadges = '<span class="badge badge-ok">OK</span>';
    } else {
      if (d._flags.includes('error'))           statusBadges += '<span class="badge badge-error">Error</span> ';
      if (d._flags.includes('dup'))             statusBadges += '<span class="badge badge-dup">Duplicate</span> ';
      if (d._flags.includes('jan-repeat'))      statusBadges += '<span class="badge badge-dup">Jan repeat</span> ';
      if (d._flags.includes('gross-computed'))  statusBadges += '<span class="badge badge-fixed">Gross added</span> ';
      if (d._flags.includes('gross-corrected')) statusBadges += '<span class="badge badge-gross-wrong">Gross fixed</span> ';
    }

    html += `<tr class="${rowClass}" onclick="showDetail(${i})">
      <td class="cell-date">${d.date}</td>
      <td class="cell-customer"><div class="cell-customer-inner" title="${esc(d.customer)}">${esc(d.customer)}</div></td>
      <td>${getBadge(d.voucher)}</td>
      <td class="cell-voucher">${esc(d.voucher)}</td>
      <td class="cell-narr" title="${esc(d.narration)}">${esc(narr) || '—'}</td>
      <td class="cell-items" title="${esc(d.itemName)}">${esc(itemDisp) || '—'}</td>
      <td class="cell-num">${fmtQty(d.qty)}</td>
      <td class="cell-num">${rateCell}</td>
      <td class="cell-num">${fmt(d.value)}</td>
      <td class="cell-num cell-gross">${fmt(d.gross)}${d._grossOriginal ? `<span title="Was: ₹${fmt2(d._grossOriginal)}" style="cursor:help;margin-left:3px;color:var(--amber)">✏️</span>` : ''}</td>
      <td style="white-space:nowrap">${statusBadges}</td>
    </tr>`;
  });

  const uv = new Set(filtered.map(d => d.voucher)).size;
  html += `<tr class="totals-row">
    <td colspan="7">Total — ${filtered.length} rows / ${uv} vouchers</td>
    <td></td>
    <td class="cell-num">${fmt(totVal)}</td>
    <td class="cell-num">${fmt(totGross)}</td>
    <td></td>
  </tr>`;

  tbody.innerHTML = html;
  document.getElementById('pageInfo').textContent = `${filtered.length} rows · ${uv} vouchers (of ${allData.length} total)`;
  document.getElementById('filteredTotal').textContent = filtered.length < allData.length ? 'Filtered gross: ' + fmtK(totGross) : '';
}

/* ═══════════════════════════════════════════════════════════════
   DETAIL MODAL — showDetail()
═══════════════════════════════════════════════════════════════ */
function showDetail(idx) {
  const d = filtered[idx];

  /* Build flag warning box */
  let flagHtml = '';
  if (d._flags.length) {
    const msgs = [];
    if (d._flags.includes('error'))           msgs.push('❌ Missing required field(s)');
    if (d._flags.includes('dup'))             msgs.push('🔁 Possible duplicate row');
    if (d._flags.includes('jan-repeat'))      msgs.push('📅 Repeated January entry');
    if (d._flags.includes('gross-computed'))  msgs.push('🔧 Gross was missing — computed as Qty × Rate');
    if (d._flags.includes('gross-corrected')) msgs.push(`✏️ Gross corrected from ₹${fmt2(d._grossOriginal || 0)} → ₹${fmt2(d.gross)}`);
    flagHtml = `<div class="flag-box">${msgs.join('<br>')}</div>`;
  }

  /* Product parse for item */
  const pp = cleanProductName(d.itemName);
  const parsedHtml = d.itemName ? `
    <div style="background:var(--line2);border-radius:7px;padding:10px 12px;font-size:11.5px;margin-top:10px">
      <div style="font-size:10.5px;font-weight:600;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:8px">Parsed product fields</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px">
        <div><span style="color:var(--muted)">Clean name:</span> ${esc(pp.cleanName)||'—'}</div>
        <div><span style="color:var(--muted)">Part #:</span> <strong>${esc(pp.partNumber)||'—'}</strong></div>
        <div><span style="color:var(--muted)">Make:</span> ${esc(pp.make)||'—'}</div>
        <div><span style="color:var(--muted)">Tracker:</span> ${esc(pp.tracker)||'—'}</div>
      </div>
    </div>` : '';

  document.getElementById('modalVoucher').textContent = d.voucher;
  document.getElementById('modalDate').textContent    = d.date + ' · ' + d.vtype;
  document.getElementById('modalFlagBox').innerHTML   = flagHtml;
  document.getElementById('detailGrid').innerHTML = `
    <div class="detail-item"><div class="detail-label">Customer</div><div class="detail-value">${esc(d.customer)}</div></div>
    <div class="detail-item"><div class="detail-label">Series</div><div class="detail-value">${getBadge(d.voucher)}</div></div>
    <div class="detail-item" style="grid-column:span 2"><div class="detail-label">Item / Part Number</div><div class="detail-value" style="font-size:14px">${esc(d.itemName)||'—'}</div></div>
    <div class="detail-item"><div class="detail-label">Quantity</div><div class="detail-value mono">${fmtQty(d.qty)||'—'}</div></div>
    <div class="detail-item"><div class="detail-label">Rate</div><div class="detail-value mono">${d.rate > 0 ? fmt(d.rate) : '—'}</div></div>
    <div class="detail-item"><div class="detail-label">Value</div><div class="detail-value mono">${fmt(d.value)}</div></div>
    <div class="detail-item"><div class="detail-label">Gross (allocated)</div><div class="detail-value big">${fmt(d.gross)}</div></div>`;
  document.getElementById('narrationSection').innerHTML = d.narration
    ? `<div style="margin-bottom:12px"><div class="detail-label" style="margin-bottom:5px">Narration</div><div class="narr-box">${esc(d.narration)}</div></div>` : '';
  document.getElementById('itemsSection').innerHTML = parsedHtml;
  document.getElementById('detailOverlay').classList.add('open');
}
function closeModal() { document.getElementById('detailOverlay').classList.remove('open'); }
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

/* ═══════════════════════════════════════════════════════════════
   DASHBOARD — updateDashboard()
═══════════════════════════════════════════════════════════════ */
function updateDashboard() {
  const totVal   = allData.reduce((s, d) => s + d.value, 0);
  const totGross = allData.reduce((s, d) => s + d.gross, 0);
  const unique   = new Set(allData.map(d => d.customer.trim().toLowerCase())).size;
  const uniqueV  = new Set(allData.map(d => d.voucher)).size;
  const cleanRows= allData.filter(d => !d._flags.length).length;

  document.getElementById('s-count').textContent         = allData.length.toLocaleString('en-IN');
  document.getElementById('s-count-sub').textContent     = `${uniqueV} vouchers · ${cancelledCount} cancelled excluded`;
  document.getElementById('s-clean').textContent         = cleanRows.toLocaleString('en-IN');
  document.getElementById('s-clean-sub').textContent     = `${unique} unique customers`;
  document.getElementById('s-dupes').textContent         = cleanStats.dupes + cleanStats.janRepeats;
  document.getElementById('s-dupes-sub').textContent     = `${cleanStats.janRepeats} Jan re-export dupes · ${cleanStats.dupes} exact dupes`;
  document.getElementById('s-grossfixed').textContent    = cleanStats.grossFixed;
  document.getElementById('s-grossfixed-sub').textContent= `${cleanStats.grossWrong} values corrected`;
  document.getElementById('s-errors').textContent        = errorLog.length;
  document.getElementById('s-errors-sub').textContent    = `${cleanStats.missingFields} missing fields`;

  renderTopCustomersChart('topCustomersChart', 8, false);
  renderTypeBreakdown('typeBreakdown');
  renderMonthlyGrid();
}

function renderTopCustomersChart(elId, n, green) {
  const m = {};
  allData.forEach(d => { const k = d.customer.trim(); m[k] = (m[k]||0) + d.gross; });
  const top = Object.entries(m).sort((a,b) => b[1]-a[1]).slice(0,n);
  const max = top[0] ? top[0][1] : 1;
  document.getElementById(elId).innerHTML = top.map(([name,val]) => `
    <div class="bar-row">
      <div class="bar-label" title="${esc(name)}">${esc(name)}</div>
      <div class="bar-track"><div class="bar-fill${green?' green':''}" style="width:${(val/max*100).toFixed(1)}%"></div></div>
      <div class="bar-val">${fmtK(val)}</div>
    </div>`).join('');
}

function renderTypeBreakdown(elId) {
  const m = {};
  const colors = ['var(--accent)','var(--green)','var(--amber)','var(--purple)','var(--orange)','var(--red)','var(--teal)'];
  allData.forEach(d => {
    const k = d.vtype || '(no type)';
    if (!m[k]) m[k] = { count:0, gross:0 };
    m[k].count++; m[k].gross += d.gross;
  });
  const sorted = Object.entries(m).sort((a,b) => b[1].gross - a[1].gross);
  document.getElementById(elId).innerHTML = sorted.map(([name,v], i) => `
    <div class="type-row">
      <div class="type-left">
        <div class="type-dot" style="background:${colors[i % colors.length]}"></div>
        <div class="type-name">${esc(name)}</div>
      </div>
      <div class="type-right">
        <div class="type-count">${v.count} rows</div>
        <div class="type-amount">${fmtK(v.gross)}</div>
      </div>
    </div>`).join('');
}

function renderMonthlyGrid() {
  const m = {};
  allData.forEach(d => {
    const parts = d.date.split('/');
    if (parts.length < 3) return;
    const [, mm, yy] = parts; const k = `${mm}/${yy}`;
    if (!m[k]) m[k] = { count:0, gross:0, label:new Date(+yy,+mm-1).toLocaleString('default',{month:'short',year:'2-digit'}) };
    m[k].count++; m[k].gross += d.gross;
  });
  const sorted = Object.entries(m).sort((a,b)=>new Date('01/'+a[0])-new Date('01/'+b[0])).slice(-12);
  document.getElementById('monthlyGrid').innerHTML = sorted.map(([,v]) => `
    <div class="month-item">
      <div class="month-name">${v.label}</div>
      <div class="month-count">${v.count}</div>
      <div class="month-val">${fmtK(v.gross)}</div>
    </div>`).join('');
}

/* ═══════════════════════════════════════════════════════════════
   ANALYTICS — renderAnalytics()
═══════════════════════════════════════════════════════════════ */
function renderAnalytics() {
  if (!allData.length) return;
  const totGross  = allData.reduce((s,d) => s+d.gross, 0);
  const avg       = totGross / allData.length;
  const maxRec    = allData.reduce((a,b) => b.gross>a.gross?b:a, allData[0]);
  const dates     = allData.map(d => dp(d.date)).filter(Boolean);
  const minD      = new Date(Math.min(...dates)), maxD = new Date(Math.max(...dates));
  const uItems    = new Set(allData.map(d => d.itemName.trim().toLowerCase()).filter(Boolean)).size;

  document.getElementById('a-avg').textContent      = fmtK(avg);
  document.getElementById('a-max').textContent      = fmt(maxRec.gross);
  document.getElementById('a-max-cust').textContent = maxRec.customer.slice(0,32);
  document.getElementById('a-range').textContent    = minD.toLocaleDateString('en-GB') + ' – ' + maxD.toLocaleDateString('en-GB');
  document.getElementById('a-multi').textContent    = uItems.toLocaleString('en-IN');
  document.getElementById('a-total').textContent    = fmtK(totGross);

  renderTopCustomersChart('bigCustomerChart', 15, true);

  /* Monthly gross bar */
  const mo = {};
  allData.forEach(d => {
    const p = d.date.split('/'); if (p.length < 3) return;
    const k = `${p[1]}/${p[2]}`;
    if (!mo[k]) mo[k] = { count:0, gross:0, label:new Date(+p[2],+p[1]-1).toLocaleString('default',{month:'short',year:'2-digit'}) };
    mo[k].count++; mo[k].gross += d.gross;
  });
  const moS = Object.entries(mo).sort((a,b)=>new Date('01/'+a[0])-new Date('01/'+b[0]));
  const mMax = Math.max(...moS.map(([,v]) => v.gross), 1);
  document.getElementById('monthBarChart').innerHTML = moS.map(([,v]) => `
    <div class="bar-row">
      <div class="bar-label">${v.label}</div>
      <div class="bar-track"><div class="bar-fill green" style="width:${(v.gross/mMax*100).toFixed(1)}%"></div></div>
      <div class="bar-val">${fmtK(v.gross)}</div>
    </div>`).join('');

  /* Top items */
  const im = {};
  allData.forEach(d => {
    if (!d.itemName) return;
    if (!im[d.itemName]) im[d.itemName] = { gross:0, qty:0 };
    im[d.itemName].gross += d.gross; im[d.itemName].qty += d.qty;
  });
  const topI = Object.entries(im).sort((a,b) => b[1].gross-a[1].gross).slice(0,12);
  const iMax = topI[0] ? topI[0][1].gross : 1;
  document.getElementById('vtypeBreak').innerHTML = topI.map(([name,v]) => `
    <div class="bar-row">
      <div class="bar-label" title="${esc(name)}" style="width:175px">${esc(name.slice(0,26))}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${(v.gross/iMax*100).toFixed(1)}%"></div></div>
      <div class="bar-val">${fmtK(v.gross)}</div>
    </div>`).join('');
}

/* ═══════════════════════════════════════════════════════════════
   CLEAN LOG (Quality Report) — updateCleanLog()
═══════════════════════════════════════════════════════════════ */
function updateCleanLog() {
  const totGross    = allData.reduce((s,d) => s+d.gross, 0);
  const unique      = new Set(allData.map(d => d.customer.trim().toLowerCase())).size;
  const uniqueV     = new Set(allData.map(d => d.voucher)).size;
  const uniqueItems = new Set(allData.map(d => d.itemName.trim().toLowerCase()).filter(Boolean)).size;
  const vtypeSeries = [...new Set(allData.map(d => d.voucher.split('/')[0]))].join(', ');

  document.getElementById('cleanSummaryTable').innerHTML = `
    <table class="issues-table">
      <tr><th>Metric</th><th>Value</th></tr>
      <tr><td>File name</td><td>${esc(currentFile)}</td></tr>
      <tr><td>Total rows before cleaning</td><td><strong>${rawParsed.length}</strong></td></tr>
      <tr><td>Output rows (item-level) after cleaning</td><td><strong>${allData.length}</strong></td></tr>
      <tr><td>Unique vouchers</td><td>${uniqueV}</td></tr>
      <tr><td>Cancelled vouchers excluded</td><td>${cancelledCount}</td></tr>
      <tr><td>Duplicate rows removed</td><td>${cleanStats.dupes}</td></tr>
      <tr><td>January repeats removed</td><td>${cleanStats.janRepeats}</td></tr>
      <tr><td>Gross values fixed / computed</td><td>${cleanStats.grossFixed}</td></tr>
      <tr><td>Gross values corrected (were wrong)</td><td>${cleanStats.grossWrong}</td></tr>
      <tr><td>Errors found</td><td>${errorLog.length}</td></tr>
      <tr><td>Unique item/part names</td><td>${uniqueItems}</td></tr>
      <tr><td>Unique customers</td><td>${unique}</td></tr>
      <tr><td>Total gross value</td><td><strong>${fmt(totGross)}</strong></td></tr>
      <tr><td>Voucher series detected</td><td>${esc(vtypeSeries)}</td></tr>
    </table>`;

  const flags = [];
  const missingName  = allData.filter(d => !d.itemName).length;
  const missingRate  = allData.filter(d => d.rate === 0).length;
  const missingVal   = allData.filter(d => d.value === 0).length;
  const missingGross = allData.filter(d => d.gross === 0).length;

  if (missingName)   flags.push({ label:'Item name missing (voucher with no detail rows)', count:missingName, sev:'warn' });
  if (missingRate)   flags.push({ label:'Rate = 0 (service items or freight charges)', count:missingRate, sev:'info' });
  if (missingVal)    flags.push({ label:'Value = ₹0', count:missingVal, sev:'err' });
  if (missingGross)  flags.push({ label:'Gross = ₹0 (may be SEZ/export)', count:missingGross, sev:'warn' });
  if (cleanStats.grossWrong) flags.push({ label:'Gross values corrected (Qty × Rate mismatch)', count:cleanStats.grossWrong, sev:'warn' });
  if (cleanStats.dupes)      flags.push({ label:'Duplicate rows removed', count:cleanStats.dupes, sev:'warn' });
  if (cleanStats.janRepeats) flags.push({ label:'January repeat entries removed', count:cleanStats.janRepeats, sev:'warn' });

  document.getElementById('qualityFlags').innerHTML = flags.length
    ? `<table class="issues-table">
        <tr><th>Issue</th><th>Count</th><th>Status</th></tr>
        ${flags.map(f=>`<tr>
          <td>${f.label}</td>
          <td>${f.count}</td>
          <td><span class="issue-badge ib-${f.sev==='err'?'err':f.sev==='warn'?'warn':'ok'}">${f.sev==='err'?'Review':f.sev==='warn'?'Check':'Info'}</span></td>
        </tr>`).join('')}
      </table>`
    : '<div style="color:var(--green);font-size:13px;font-weight:500">✅ No quality issues found</div>';
}

/* ═══════════════════════════════════════════════════════════════
   EXPORT — exportExcel()
   Sheet 1: Cleaned Data
   Sheet 2: Error Log
   Sheet 3: Item Summary
   Sheet 4: Product Names (parsed)
═══════════════════════════════════════════════════════════════ */
function exportExcel() {
  if (!filtered.length) { showToast('⚠️ No data to export'); return; }

  /* Sheet 1 — Cleaned Data (Item-Level Detail) */
  const wsData = [
    ['Date','Customer','Voucher Type','Credit Note No.','Narration','Item / Part Number','Quantity','Rate','Value','Gross Total','Status Flags'],
    ...filtered.map(d => [
      d.date, d.customer, d.vtype, d.voucher, d.narration,
      d.itemName, d.qty, d.rate, d.value, d.gross,
      d._flags.join(', ')
    ])
  ];
  const tot = filtered.reduce((s,d) => ({ value:s.value+d.value, gross:s.gross+d.gross }), { value:0, gross:0 });
  wsData.push(['TOTAL','','','','','','','', Math.round(tot.value*100)/100, Math.round(tot.gross*100)/100, '']);

  /* Sheet 2 — Error Log */
  const wsErrors = [
    ['Severity','Credit Note No.','Customer','Date','Issue','Field'],
    ...errorLog.map(e => [e.sev.toUpperCase(), e.voucher, e.customer, e.date, e.msg, e.field])
  ];

  /* Sheet 3 — Item Summary */
  const iMap = {};
  filtered.forEach(d => {
    const k = d.itemName || '(no item name)';
    if (!iMap[k]) iMap[k] = { itemName:k, totalQty:0, totalValue:0, totalGross:0, txnCount:0 };
    iMap[k].totalQty   += d.qty;
    iMap[k].totalValue += d.value;
    iMap[k].totalGross += d.gross;
    iMap[k].txnCount++;
  });
  const itemSummary = Object.values(iMap).sort((a,b) => b.totalGross - a.totalGross);
  const wsSummary = [
    ['Item / Part Number','No. of Transactions','Total Qty Returned','Total Value','Total Gross'],
    ...itemSummary.map(r => [r.itemName, r.txnCount, r.totalQty, r.totalValue, r.totalGross])
  ];

  /* Sheet 4 — Product Names Parsed */
  const seenP = new Set();
  const productRows = [['Original Name','Clean Name','Part Number','Make','Tracker']];
  filtered.forEach(d => {
    if (d.itemName && !seenP.has(d.itemName)) {
      seenP.add(d.itemName);
      const p = cleanProductName(d.itemName);
      productRows.push([d.itemName, p.cleanName, p.partNumber, p.make, p.tracker]);
    }
  });

  /* Build workbook */
  const wb = XLSX.utils.book_new();

  const ws1 = XLSX.utils.aoa_to_sheet(wsData);
  ws1['!cols'] = [{wch:12},{wch:42},{wch:22},{wch:18},{wch:50},{wch:45},{wch:10},{wch:12},{wch:14},{wch:18},{wch:20}];
  XLSX.utils.book_append_sheet(wb, ws1, 'Cleaned Data');

  const ws2 = XLSX.utils.aoa_to_sheet(wsErrors);
  ws2['!cols'] = [{wch:10},{wch:18},{wch:38},{wch:12},{wch:70},{wch:20}];
  XLSX.utils.book_append_sheet(wb, ws2, 'Error Log');

  const ws3 = XLSX.utils.aoa_to_sheet(wsSummary);
  ws3['!cols'] = [{wch:45},{wch:18},{wch:18},{wch:16},{wch:16}];
  XLSX.utils.book_append_sheet(wb, ws3, 'Item Summary');

  const ws4 = XLSX.utils.aoa_to_sheet(productRows);
  ws4['!cols'] = [{wch:50},{wch:40},{wch:20},{wch:22},{wch:14}];
  XLSX.utils.book_append_sheet(wb, ws4, 'Product Names');

  const fn = `credit_notes_cleaned_${new Date().toISOString().slice(0,10)}.xlsx`;
  XLSX.writeFile(wb, fn);
  showToast('📥 Downloaded ' + fn + ' (4 sheets)');
}

/** CSV export of current filtered view */
function exportCSV() {
  if (!filtered.length) { showToast('⚠️ No data to export'); return; }
  const cols = ['Date','Customer','Voucher Type','Credit Note No.','Narration','Item / Part Number','Quantity','Rate','Value','Gross Total','Status'];
  const rows = filtered.map(d => [
    d.date,
    `"${d.customer.replace(/"/g,'""')}"`,
    `"${d.vtype.replace(/"/g,'""')}"`,
    d.voucher,
    `"${d.narration.replace(/"/g,'""')}"`,
    `"${d.itemName.replace(/"/g,'""')}"`,
    d.qty, d.rate, d.value, d.gross,
    `"${d._flags.join(', ')}"`
  ].join(','));
  const csv = [cols.join(','), ...rows].join('\n');
  const a = document.createElement('a');
  a.href = 'data:text/csv;charset=utf-8,\uFEFF' + encodeURIComponent(csv);
  a.download = `credit_notes_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  showToast('📄 CSV downloaded');
}

/* ═══════════════════════════════════════════════════════════════
   UTILITY FUNCTIONS
═══════════════════════════════════════════════════════════════ */

/** Format currency in Indian locale */
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

/** Parse a dd/mm/yyyy date string → timestamp */
function dp(s) {
  try { const [d,mo,y] = s.split('/'); return new Date(+y, +mo-1, +d).getTime(); } catch { return 0; }
}

/** Voucher series badge */
function getBadge(v) {
  if (/^CNG/i.test(v)) return '<span class="badge badge-2425">CNG 24-25</span>';
  if (/^CNS/i.test(v)) return '<span class="badge badge-surat">CNS Surat</span>';
  return '<span class="badge badge-2526">CN 25-26</span>';
}

/** HTML-escape a string to prevent XSS in innerHTML */
function esc(s) {
  if (!s) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/* ── TOAST ─────────────────────────────────────────────────── */
let _tt;
function showToast(msg) {
  clearTimeout(_tt);
  document.getElementById('toastMsg').textContent = msg;
  document.getElementById('toast').classList.add('show');
  _tt = setTimeout(() => document.getElementById('toast').classList.remove('show'), 3200);
}
