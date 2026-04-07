/* ── STATE ── */
let allData=[], filtered=[], currentSort={col:'date',dir:'desc'}, cancelledCount=0, currentFile='';

/* ── NAV ── */
function switchPage(id){
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  document.getElementById('nav-'+id).classList.add('active');
  if(id==='upload'){
    document.getElementById('page-upload').style.display='flex';
    document.getElementById('topbarTitle').textContent='Upload File';
    document.getElementById('exportBtn').style.display='none';
    document.getElementById('recordPill').style.display='none';
  } else {
    document.getElementById('page-upload').style.display='none';
    document.getElementById('page-'+id).classList.add('active');
    const titles={dashboard:'Dashboard',records:'Records',analytics:'Analytics',clean:'Clean Log'};
    document.getElementById('topbarTitle').textContent=titles[id]||id;
    document.getElementById('exportBtn').style.display=allData.length?'flex':'none';
    document.getElementById('recordPill').style.display=allData.length?'flex':'none';
    document.getElementById('recordPill').textContent=allData.length+' records';
    if(id==='analytics') renderAnalytics();
  }
}
switchPage('upload');

/* ── FILE ── */
const dropZone=document.getElementById('dropZone');
dropZone.addEventListener('dragover',e=>{e.preventDefault();dropZone.classList.add('dragover')});
dropZone.addEventListener('dragleave',()=>dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop',e=>{e.preventDefault();dropZone.classList.remove('dragover');const f=e.dataTransfer.files[0];if(f)processFile(f)});
function handleFile(e){const f=e.target.files[0];if(f)processFile(f)}
function processFile(file){
  currentFile=file.name;
  const bar=document.getElementById('processingBar'),inner=document.getElementById('processingBarInner');
  bar.style.display='block'; inner.style.width='30%';
  setTimeout(()=>inner.style.width='70%',200);
  const reader=new FileReader();
  reader.onload=e=>{
    inner.style.width='90%';
    try{
      const wb=XLSX.read(e.target.result,{type:'binary',cellDates:true});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const raw=XLSX.utils.sheet_to_json(ws,{header:1,raw:false,defval:''});
      parseData(raw);
      inner.style.width='100%';
      setTimeout(()=>{bar.style.display='none';inner.style.width='0%'},400);
    }catch(err){alert('Error reading file: '+err.message);bar.style.display='none';}
  };
  reader.readAsBinaryString(file);
}

/* ── PARSER — ITEM-LEVEL (1 item = 1 row) ── */
function parseData(rows){
  const CN_RE=/^(CN|CNG|CNS)\//i;
  const CANCELLED=/\(cancelled/i;
  const records=[];       // item-level output
  cancelledCount=0;

  /* ── numeric extractor: handles "23.0000 NOS", "290.00/NOS", plain numbers ── */
  const cn=v=>{
    if(v===null||v===undefined||String(v).trim()==='') return 0;
    const s=String(v).replace(/[^0-9.-]/g,'');
    const n=parseFloat(s);
    return isNaN(n)?0:n;
  };

  /* ── find duplicate-section start: first repeated voucher number ── */
  const seenVouchers=new Set();
  let dupStart=rows.length;
  for(let i=0;i<rows.length;i++){
    const v=String(rows[i][4]||'').trim();
    if(CN_RE.test(v)){
      if(seenVouchers.has(v)){dupStart=i;break;}
      seenVouchers.add(v);
    }
  }

  let i=0;
  while(i<dupStart){
    const r=rows[i];
    const voucher=String(r[4]||'').trim();
    const dateRaw=String(r[0]||'').trim();

    /* ── must be a voucher main row ── */
    if(!CN_RE.test(voucher)||!dateRaw||dateRaw.toLowerCase()==='date'){i++;continue;}
    if(isNaN(new Date(dateRaw).getTime())&&!/\d{4}/.test(dateRaw)){i++;continue;}

    const particulars=String(r[1]||'').trim();
    const buyer=String(r[2]||'').trim();
    if(CANCELLED.test(particulars)||CANCELLED.test(buyer)){cancelledCount++;i++;continue;}

    /* ── header fields shared by every item in this voucher ── */
    let dateStr='';
    try{const d=new Date(dateRaw);dateStr=isNaN(d.getTime())?dateRaw:d.toLocaleDateString('en-GB');}catch{dateStr=dateRaw;}

    const customer=buyer.length>=particulars.length?buyer:particulars;
    const vtype=String(r[3]||'').trim();
    const narration=String(r[5]||'').replace(/_x000D_\\n/g,' ').replace(/_x000D_\n/g,' ').trim();

    /* ── totals on main row (always correct) ── */
    const mainValue=cn(r[8]);
    let mainGross=cn(r[9]);
    if(mainGross===0&&mainValue>0) mainGross=mainValue;  // SEZ/export: 0 GST → gross = value

    /* ── collect item sub-rows until next voucher main row or grand total ── */
    let j=i+1;
    const rawItems=[];
    while(j<dupStart){
      const nr=rows[j];
      const nv=String(nr[4]||'').trim();
      const nd=String(nr[0]||'').trim();
      const np=String(nr[1]||'').trim().toLowerCase();
      // Stop: next main voucher row
      if(CN_RE.test(nv)&&nd&&!isNaN(new Date(nd).getTime())) break;
      // Stop: grand total row
      if(np==='grand total') break;
      // Collect item row (any row with a non-empty col1)
      const iName=String(nr[1]||'').trim();
      if(iName&&np!=='grand total'){
        rawItems.push({
          name  : iName,
          qty   : cn(nr[6]),   // item qty  (col6 of detail row)
          rate  : cn(nr[7]),   // item rate (col7 of detail row — NEVER on main row)
          value : cn(nr[8])    // item value (col8 of detail row)
        });
      }
      j++;
    }

    /* ── EMIT one output record per item ── */
    if(rawItems.length>0){
      // Gross allocation: distribute main gross proportionally by item value
      const totalItemVal=rawItems.reduce((s,it)=>s+it.value,0);

      rawItems.forEach(it=>{
        let itemGross=0;
        if(totalItemVal>0&&it.value>0){
          itemGross=Math.round(it.value/totalItemVal*mainGross*100)/100;
        } else if(rawItems.length===1){
          itemGross=mainGross;  // single item with no value → full gross
        }

        records.push({
          date     : dateStr,
          customer : customer,
          vtype    : vtype,
          voucher  : voucher,
          narration: narration,
          itemName : it.name,
          qty      : it.qty,
          rate     : it.rate,
          value    : it.value,
          gross    : itemGross
        });
      });

    } else {
      /* ── voucher with NO item sub-rows → emit one row using main row data ── */
      records.push({
        date     : dateStr,
        customer : customer,
        vtype    : vtype,
        voucher  : voucher,
        narration: narration,
        itemName : '',
        qty      : cn(r[6]),
        rate     : 0,
        value    : mainValue,
        gross    : mainGross
      });
    }

    i=j;  // advance past all consumed sub-rows
  }

  if(!records.length){showToast('⚠️ No records found — check file format');return;}

  allData=[...records];
  filtered=[...records];
  populateFilters();
  applyFilters();
  updateDashboard();
  updateCleanLog();
  document.getElementById('fileNameDisplay').textContent=currentFile;
  document.getElementById('fileStats').textContent=records.length+' item rows parsed';
  switchPage('dashboard');
  showToast('✅ Loaded '+records.length+' item-level rows');
}

/* ── FILTERS ── */
function populateFilters(){
  const vtypes=[...new Set(allData.map(d=>d.vtype).filter(Boolean))].sort();
  const years=[...new Set(allData.map(d=>d.date.split('/').pop()).filter(Boolean))].sort().reverse();
  document.getElementById('filterVType').innerHTML='<option value="">All types</option>'+vtypes.map(v=>`<option>${v}</option>`).join('');
  document.getElementById('filterYear').innerHTML='<option value="">All years</option>'+years.map(y=>`<option>${y}</option>`).join('');
}

function applyFilters(){
  const q=document.getElementById('searchBox').value.toLowerCase();
  const vt=document.getElementById('filterVType').value;
  const yr=document.getElementById('filterYear').value;
  const fi=document.getElementById('filterItems').value;
  const sortSel=document.getElementById('filterSort').value;

  filtered=allData.filter(d=>{
    // search across customer, voucher, narration AND item name
    const mq=!q||(d.customer.toLowerCase().includes(q)||d.voucher.toLowerCase().includes(q)||d.narration.toLowerCase().includes(q)||d.itemName.toLowerCase().includes(q));
    const mv=!vt||d.vtype===vt;
    const my=!yr||d.date.endsWith(yr);
    // "single": items with rate > 0 (physical goods); "multi": service/no-rate items
    const mf=!fi||(fi==='single'&&d.rate>0)||(fi==='multi'&&d.rate===0);
    return mq&&mv&&my&&mf;
  });

  const [col,dir]=sortSel.split('-');
  const colMap={date:'date',gross:'gross',customer:'customer'};
  currentSort={col:colMap[col]||'date',dir};
  sortFiltered();
  renderTable();

  // hide multi-item alert — no longer needed at item level
  document.getElementById('multiItemAlert').classList.remove('show');
}

function clearFilters(){
  document.getElementById('searchBox').value='';
  document.getElementById('filterVType').value='';
  document.getElementById('filterYear').value='';
  document.getElementById('filterItems').value='';
  document.getElementById('filterSort').value='date-desc';
  applyFilters();
}


/* ── SORT ── */
function sortBy(col){
  if(currentSort.col===col) currentSort.dir=currentSort.dir==='asc'?'desc':'asc';
  else currentSort={col,dir:'asc'};
  document.getElementById('filterSort').value='';
  sortFiltered();
  updateSortIcons();
  renderTable();
}
function sortFiltered(){
  const{col,dir}=currentSort;const m=dir==='asc'?1:-1;
  filtered.sort((a,b)=>{
    if(col==='date') return m*(dp(a.date)-dp(b.date));
    if(col==='customer') return m*a.customer.localeCompare(b.customer);
    if(col==='voucher') return m*a.voucher.localeCompare(b.voucher);
    return m*(a[col]-b[col]);
  });
}
function dp(s){try{const[d,mo,y]=s.split('/');return new Date(+y,+mo-1,+d).getTime();}catch{return 0;}}
function updateSortIcons(){
  ['date','customer','voucher','qty','rate','value','gross'].forEach(c=>{
    const el=document.getElementById('sa-'+c);
    if(!el)return;
    el.className='sort-arrows'+(currentSort.col===c?' '+currentSort.dir:'');
  });
}

/* ── FORMAT ── */
function fmt(n,sym='₹'){if(!n||n===0)return'—';return sym+n.toLocaleString('en-IN',{minimumFractionDigits:2,maximumFractionDigits:2});}
function fmtQty(n){if(!n||n===0)return'—';return n.toLocaleString('en-IN',{maximumFractionDigits:4});}
function fmtK(n){if(n>=10000000)return'₹'+(n/10000000).toFixed(2)+'Cr';if(n>=100000)return'₹'+(n/100000).toFixed(2)+'L';return'₹'+Math.round(n).toLocaleString('en-IN');}
function getBadge(v){if(/^CNG/i.test(v))return'<span class="badge badge-2425">CNG 24-25</span>';if(/^CNS/i.test(v))return'<span class="badge badge-surat">CNS Surat</span>';return'<span class="badge badge-2526">CN 25-26</span>';}

/* ── TABLE ── */
function renderTable(){
  const tbody=document.getElementById('tableBody');
  document.getElementById('filterCount').textContent=filtered.length+' rows';

  if(!filtered.length){
    tbody.innerHTML=`<tr><td colspan="10"><div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-title">No records match</div></div></td></tr>`;
    document.getElementById('pageInfo').textContent='0 rows';
    document.getElementById('filteredTotal').textContent='';
    return;
  }

  const totVal=filtered.reduce((s,d)=>s+d.value,0);
  const totGross=filtered.reduce((s,d)=>s+d.gross,0);

  let html='';
  filtered.forEach((d,i)=>{
    const narr=d.narration.length>45?d.narration.slice(0,42)+'…':d.narration;
    const itemDisp=d.itemName.length>32?d.itemName.slice(0,29)+'…':d.itemName;
    const rateCell=d.rate>0?fmt(d.rate):'<span style="color:var(--muted);font-size:11px">—</span>';
    html+=`<tr onclick="showDetail(${i})">
      <td class="cell-date">${d.date}</td>
      <td class="cell-customer"><div class="cell-customer-inner" title="${d.customer}">${d.customer}</div></td>
      <td>${getBadge(d.voucher)}</td>
      <td class="cell-voucher">${d.voucher}</td>
      <td class="cell-narr" title="${d.narration}">${narr||'—'}</td>
      <td class="cell-items" title="${d.itemName}" style="font-weight:500;color:var(--ink2)">${itemDisp||'—'}</td>
      <td class="cell-num">${fmtQty(d.qty)}</td>
      <td class="cell-num">${rateCell}</td>
      <td class="cell-num">${fmt(d.value)}</td>
      <td class="cell-num cell-gross">${fmt(d.gross)}</td>
    </tr>`;
  });

  html+=`<tr class="totals-row">
    <td colspan="6">Total — ${filtered.length} item row${filtered.length!==1?'s':''}</td>
    <td></td><td></td>
    <td class="cell-num">${fmt(totVal)}</td>
    <td class="cell-num">${fmt(totGross)}</td>
  </tr>`;

  tbody.innerHTML=html;
  // Count unique vouchers in filtered set
  const uniqueVouchers=new Set(filtered.map(d=>d.voucher)).size;
  document.getElementById('pageInfo').textContent=`${filtered.length} item rows across ${uniqueVouchers} vouchers (of ${allData.length} total rows)`;
  document.getElementById('filteredTotal').textContent=filtered.length<allData.length?'Filtered gross: '+fmtK(totGross):'';
}

/* ── MODAL ── */
function showDetail(idx){
  const d=filtered[idx];
  document.getElementById('modalVoucher').textContent=d.voucher;
  document.getElementById('modalDate').textContent=d.date+' · '+d.vtype;
  document.getElementById('detailGrid').innerHTML=`
    <div class="detail-item"><div class="detail-label">Customer</div><div class="detail-value">${d.customer}</div></div>
    <div class="detail-item"><div class="detail-label">Series</div><div class="detail-value">${getBadge(d.voucher)}</div></div>
    <div class="detail-item" style="grid-column:span 2"><div class="detail-label">Item / Part Number</div><div class="detail-value" style="font-size:15px">${d.itemName||'—'}</div></div>
    <div class="detail-item"><div class="detail-label">Quantity</div><div class="detail-value mono">${fmtQty(d.qty)||'—'}</div></div>
    <div class="detail-item"><div class="detail-label">Rate</div><div class="detail-value mono">${d.rate>0?fmt(d.rate):'—'}</div></div>
    <div class="detail-item"><div class="detail-label">Value</div><div class="detail-value mono">${fmt(d.value)}</div></div>
    <div class="detail-item"><div class="detail-label">Gross (allocated)</div><div class="detail-value big">${fmt(d.gross)}</div></div>
  `;
  document.getElementById('narrationSection').innerHTML=d.narration
    ?`<div style="margin-bottom:14px"><div class="detail-label" style="margin-bottom:6px">Narration</div><div class="narr-box">${d.narration}</div></div>`:'';
  document.getElementById('itemsSection').innerHTML='';
  document.getElementById('detailOverlay').classList.add('open');
}
function closeModal(){document.getElementById('detailOverlay').classList.remove('open');}
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeModal();});

/* ── DASHBOARD ── */
function updateDashboard(){
  const totVal=allData.reduce((s,d)=>s+d.value,0);
  const totGross=allData.reduce((s,d)=>s+d.gross,0);
  const unique=new Set(allData.map(d=>d.customer.trim().toLowerCase())).size;
  const uniqueVouchers=new Set(allData.map(d=>d.voucher)).size;
  document.getElementById('s-count').textContent=allData.length.toLocaleString('en-IN');
  document.getElementById('s-count-sub').textContent=uniqueVouchers+' vouchers · '+cancelledCount+' cancelled excluded';
  document.getElementById('s-value').textContent=fmtK(totVal);
  document.getElementById('s-value-sub').textContent='sum of all item values';
  document.getElementById('s-gross').textContent=fmtK(totGross);
  document.getElementById('s-gross-sub').textContent='incl. GST';
  document.getElementById('s-customers').textContent=unique.toLocaleString('en-IN');
  document.getElementById('s-customers-sub').textContent='unique buyers';
  renderTopCustomersChart('topCustomersChart',8,false);
  renderTypeBreakdown('typeBreakdown');
  renderMonthlyGrid();
}

function renderTopCustomersChart(elId,n,green){
  const m={};allData.forEach(d=>{const k=d.customer.trim();m[k]=(m[k]||0)+d.gross;});
  const top=Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,n);
  const max=top[0]?top[0][1]:1;
  document.getElementById(elId).innerHTML=top.map(([name,val])=>`
    <div class="bar-row">
      <div class="bar-label" title="${name}">${name}</div>
      <div class="bar-track"><div class="bar-fill${green?' green':''}" style="width:${(val/max*100).toFixed(1)}%"></div></div>
      <div class="bar-val">${fmtK(val)}</div>
    </div>`).join('');
}
function renderTypeBreakdown(elId){
  const colors=['#2563eb','#059669','#d97706','#7c3aed','#dc2626'];
  const m={};allData.forEach(d=>{if(!m[d.vtype])m[d.vtype]={count:0,gross:0};m[d.vtype].count++;m[d.vtype].gross+=d.gross;});
  const list=Object.entries(m).sort((a,b)=>b[1].gross-a[1].gross);
  document.getElementById(elId).innerHTML=list.map(([name,v],i)=>`
    <div class="type-row">
      <div class="type-left"><div class="type-dot" style="background:${colors[i%colors.length]}"></div><div class="type-name">${name}</div></div>
      <div class="type-right"><div class="type-count">${v.count} records</div><div class="type-amount">${fmtK(v.gross)}</div></div>
    </div>`).join('');
}
function renderMonthlyGrid(){
  const m={};
  allData.forEach(d=>{
    const[dd,mm,yy]=d.date.split('/');const k=`${mm}/${yy}`;
    if(!m[k])m[k]={count:0,gross:0,label:new Date(+yy,+mm-1).toLocaleString('default',{month:'short',year:'2-digit'})};
    m[k].count++;m[k].gross+=d.gross;
  });
  const sorted=Object.entries(m).sort((a,b)=>new Date('01/'+a[0])-new Date('01/'+b[0])).slice(-12);
  document.getElementById('monthlyGrid').innerHTML=sorted.map(([,v])=>`
    <div class="month-item"><div class="month-name">${v.label}</div><div class="month-count">${v.count}</div><div class="month-val">${fmtK(v.gross)}</div></div>`).join('');
}

/* ── ANALYTICS ── */
function renderAnalytics(){
  if(!allData.length)return;
  const totGross=allData.reduce((s,d)=>s+d.gross,0);
  const avg=totGross/allData.length;
  const maxRec=allData.reduce((a,b)=>b.gross>a.gross?b:a,allData[0]);
  const dates=allData.map(d=>dp(d.date)).filter(Boolean);
  const minD=new Date(Math.min(...dates)),maxD=new Date(Math.max(...dates));
  const uniqueItems=new Set(allData.map(d=>d.itemName.trim().toLowerCase()).filter(Boolean)).size;

  document.getElementById('a-avg').textContent=fmtK(avg);
  document.getElementById('a-max').textContent=fmt(maxRec.gross);
  document.getElementById('a-max-cust').textContent=maxRec.customer.slice(0,32);
  document.getElementById('a-range').textContent=minD.toLocaleDateString('en-GB')+' – '+maxD.toLocaleDateString('en-GB');
  document.getElementById('a-multi').textContent=uniqueItems.toLocaleString('en-IN');

  // Top customers by gross
  renderTopCustomersChart('bigCustomerChart',15,true);

  // Monthly gross
  const m={};
  allData.forEach(d=>{
    const[dd,mm,yy]=d.date.split('/');const k=`${mm}/${yy}`;
    if(!m[k])m[k]={count:0,gross:0,label:new Date(+yy,+mm-1).toLocaleString('default',{month:'short',year:'2-digit'})};
    m[k].count++;m[k].gross+=d.gross;
  });
  const mSorted=Object.entries(m).sort((a,b)=>new Date('01/'+a[0])-new Date('01/'+b[0]));
  const mMax=Math.max(...mSorted.map(([,v])=>v.gross),1);
  document.getElementById('monthBarChart').innerHTML=mSorted.map(([,v])=>`
    <div class="bar-row"><div class="bar-label">${v.label}</div><div class="bar-track"><div class="bar-fill green" style="width:${(v.gross/mMax*100).toFixed(1)}%"></div></div><div class="bar-val">${fmtK(v.gross)}</div></div>`).join('');

  // Top items by gross (new — only possible at item level)
  const itemMap={};
  allData.forEach(d=>{
    if(!d.itemName) return;
    if(!itemMap[d.itemName])itemMap[d.itemName]={gross:0,qty:0,count:0};
    itemMap[d.itemName].gross+=d.gross;
    itemMap[d.itemName].qty+=d.qty;
    itemMap[d.itemName].count++;
  });
  const topItems=Object.entries(itemMap).sort((a,b)=>b[1].gross-a[1].gross).slice(0,12);
  const iMax=topItems[0]?topItems[0][1].gross:1;
  document.getElementById('vtypeBreak').innerHTML=topItems.map(([name,v])=>`
    <div class="bar-row">
      <div class="bar-label" title="${name}" style="width:180px">${name.slice(0,28)}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${(v.gross/iMax*100).toFixed(1)}%"></div></div>
      <div class="bar-val">${fmtK(v.gross)}</div>
    </div>`).join('');
}

/* ── CLEAN LOG ── */
function updateCleanLog(){
  const totGross=allData.reduce((s,d)=>s+d.gross,0);
  const unique=new Set(allData.map(d=>d.customer.trim().toLowerCase())).size;
  const uniqueVouchers=new Set(allData.map(d=>d.voucher)).size;
  const uniqueItems=new Set(allData.map(d=>d.itemName.trim().toLowerCase()).filter(Boolean)).size;
  const missingGross=allData.filter(d=>d.gross===0).length;
  const missingVal=allData.filter(d=>d.value===0).length;
  const missingRate=allData.filter(d=>d.rate===0).length;
  const missingName=allData.filter(d=>!d.itemName).length;
  const vtypeSeries=[...new Set(allData.map(d=>d.voucher.split('/')[0]))].join(', ');

  document.getElementById('cleanSummaryTable').innerHTML=`
    <table class="issues-table">
      <tr><th>Metric</th><th>Value</th></tr>
      <tr><td>File name</td><td>${currentFile}</td></tr>
      <tr><td>Output rows (item-level)</td><td><strong>${allData.length}</strong></td></tr>
      <tr><td>Unique vouchers</td><td>${uniqueVouchers}</td></tr>
      <tr><td>Cancelled vouchers excluded</td><td>${cancelledCount}</td></tr>
      <tr><td>Unique item/part names</td><td>${uniqueItems}</td></tr>
      <tr><td>Unique customers</td><td>${unique}</td></tr>
      <tr><td>Total gross value</td><td><strong>${fmt(totGross)}</strong></td></tr>
      <tr><td>Voucher series detected</td><td>${vtypeSeries}</td></tr>
    </table>`;

  const flags=[];
  if(missingName)   flags.push({label:'Item name missing (voucher with no detail rows)',count:missingName,sev:'warn'});
  if(missingRate)   flags.push({label:'Rate = 0 (service items or freight charges)',count:missingRate,sev:'info'});
  if(missingVal)    flags.push({label:'Value = ₹0',count:missingVal,sev:'err'});
  if(missingGross)  flags.push({label:'Gross = ₹0 (may be SEZ/export)',count:missingGross,sev:'warn'});

  document.getElementById('qualityFlags').innerHTML=flags.length
    ?`<table class="issues-table">
        <tr><th>Issue</th><th>Count</th><th>Status</th></tr>
        ${flags.map(f=>`<tr><td>${f.label}</td><td>${f.count}</td><td><span class="issue-badge ib-${f.sev==='err'?'err':f.sev==='warn'?'warn':'ok'}">${f.sev==='err'?'Review':f.sev==='warn'?'Check':'Info'}</span></td></tr>`).join('')}
      </table>`
    :'<div style="color:var(--green);font-size:13px;font-weight:500">✅ No quality issues found</div>';
}

/* ── EXPORT ── */
function exportExcel(){
  if(!filtered.length){showToast('⚠️ No data to export');return;}

  // Sheet 1: Item-level detail (main output)
  const wsData=[
    ['Date','Customer','Voucher Type','Voucher No.','Narration','Item / Part Number','Quantity','Rate','Value','Gross Total (allocated)'],
    ...filtered.map(d=>[d.date,d.customer,d.vtype,d.voucher,d.narration,d.itemName,d.qty,d.rate,d.value,d.gross])
  ];
  const tot=filtered.reduce((s,d)=>({value:s.value+d.value,gross:s.gross+d.gross}),{value:0,gross:0});
  wsData.push(['TOTAL','','','','','','','',tot.value,tot.gross]);

  // Sheet 2: Item summary — unique items with total qty + total gross (for annual product analysis)
  const itemSummaryMap={};
  filtered.forEach(d=>{
    const k=d.itemName||'(no item name)';
    if(!itemSummaryMap[k])itemSummaryMap[k]={itemName:k,totalQty:0,totalValue:0,totalGross:0,txnCount:0};
    itemSummaryMap[k].totalQty+=d.qty;
    itemSummaryMap[k].totalValue+=d.value;
    itemSummaryMap[k].totalGross+=d.gross;
    itemSummaryMap[k].txnCount++;
  });
  const itemSummary=Object.values(itemSummaryMap).sort((a,b)=>b.totalGross-a.totalGross);
  const wsSummaryData=[
    ['Item / Part Number','No. of Transactions','Total Qty Returned','Total Value','Total Gross'],
    ...itemSummary.map(r=>[r.itemName,r.txnCount,Math.round(r.totalQty*10000)/10000,Math.round(r.totalValue*100)/100,Math.round(r.totalGross*100)/100])
  ];

  const wb=XLSX.utils.book_new();

  const ws=XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols']=[{wch:12},{wch:42},{wch:22},{wch:18},{wch:50},{wch:45},{wch:10},{wch:12},{wch:14},{wch:18}];
  XLSX.utils.book_append_sheet(wb,ws,'Item-Level Detail');

  const ws2=XLSX.utils.aoa_to_sheet(wsSummaryData);
  ws2['!cols']=[{wch:45},{wch:18},{wch:18},{wch:16},{wch:16}];
  XLSX.utils.book_append_sheet(wb,ws2,'Item Summary');

  const fn=`credit_notes_items_${new Date().toISOString().slice(0,10)}.xlsx`;
  XLSX.writeFile(wb,fn);
  showToast('📥 Downloaded '+fn+' (2 sheets)');
}

function exportCSV(){
  if(!filtered.length){showToast('⚠️ No data to export');return;}
  const cols=['Date','Customer','Voucher Type','Voucher No.','Narration','Item / Part Number','Quantity','Rate','Value','Gross Total'];
  const rows=filtered.map(d=>[
    d.date,
    `"${d.customer.replace(/"/g,'""')}"`,
    `"${d.vtype.replace(/"/g,'""')}"`,
    d.voucher,
    `"${d.narration.replace(/"/g,'""')}"`,
    `"${d.itemName.replace(/"/g,'""')}"`,
    d.qty,d.rate,d.value,d.gross
  ].join(','));
  const csv=[cols.join(','),...rows].join('\n');
  const a=document.createElement('a');
  a.href='data:text/csv;charset=utf-8,\uFEFF'+encodeURIComponent(csv);
  a.download=`credit_notes_items_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  showToast('📄 CSV downloaded');
}

/* ── TOAST ── */
let tt;
function showToast(msg){
  clearTimeout(tt);
  document.getElementById('toastMsg').textContent=msg;
  document.getElementById('toast').classList.add('show');
  tt=setTimeout(()=>document.getElementById('toast').classList.remove('show'),3000);
}
