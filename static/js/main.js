/**
 * static/js/main.js - Interactive client logic for REPOTIC Automation
 */

// ── Drag & Drop Configuration ────────────────────────────────────────────────
const ZONES = [
  { dz: 'dz-repotic',  inp: 'inp-repotic',  pill: 'pill-repotic',  name: 'name-repotic',  size: 'size-repotic'  },
  { dz: 'dz-ledger',   inp: 'inp-ledger',   pill: 'pill-ledger',   name: 'name-ledger',   size: 'size-ledger'   },
  { dz: 'dz-template', inp: 'inp-template', pill: 'pill-template', name: 'name-template', size: 'size-template' }
];

function formatFileSize(bytes) {
  if (!bytes) return '';
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

ZONES.forEach(({ dz, inp, pill, name, size }) => {
  const zone = document.getElementById(dz);
  const input = document.getElementById(inp);

  input.addEventListener('change', () => {
    if (input.files[0]) markFile(zone, pill, name, size, input.files[0]);
  });

  zone.addEventListener('dragover', (e) => {
    e.preventDefault();
    zone.classList.add('drag-over');
  });

  zone.addEventListener('dragleave', () => {
    zone.classList.remove('drag-over');
  });

  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (!file) return;

    const dt = new DataTransfer();
    dt.items.add(file);
    input.files = dt.files;
    markFile(zone, pill, name, size, file);
  });
});

function markFile(zone, pill, nameEl, sizeEl, file) {
  zone.classList.add('loaded');
  document.getElementById(nameEl).textContent = file.name;
  document.getElementById(sizeEl).textContent = formatFileSize(file.size);
  checkReady();
}

function checkReady() {
  const ready = ZONES.every(({ inp }) => document.getElementById(inp).files.length > 0);
  const btn = document.getElementById('runBtn');
  const hint = document.getElementById('actionHint');

  btn.disabled = !ready;
  if (ready) {
    hint.textContent = 'All files loaded. Click above to process and generate output.';
  } else {
    hint.textContent = 'Upload all three Excel files above to start processing.';
  }
}

// ── Accounting Period Change Handler ─────────────────────────────────────────
function onPeriodChange() {
  const month = parseInt(document.getElementById('sel-month').value, 10);
  const year = parseInt(document.getElementById('sel-year').value, 10);

  // Financial Year (starts April)
  let fy;
  if (month >= 4) {
    const s = String(year % 100).padStart(2, '0');
    const e = String((year + 1) % 100).padStart(2, '0');
    fy = `${s}-${e}`;
  } else {
    const s = String((year - 1) % 100).padStart(2, '0');
    const e = String(year % 100).padStart(2, '0');
    fy = `${s}-${e}`;
  }

  const mm = String(month).padStart(2, '0');
  const prefix = `${mm}/${fy}/`;

  // Last day of month
  const lastDay = new Date(year, month, 0).getDate();
  const dd = String(lastDay).padStart(2, '0');
  const invDate = `${dd}-${mm}-${year}`;

  document.getElementById('inp-inv-prefix').value = prefix;
  document.getElementById('inp-inv-date').value = invDate;
}

// ── Execution Handler ────────────────────────────────────────────────────────
async function runAutomation() {
  const btn = document.getElementById('runBtn');
  const label = document.getElementById('runLabel');
  const errBox = document.getElementById('errorBox');

  btn.classList.add('loading');
  btn.disabled = true;
  errBox.style.display = 'none';

  const fd = new FormData();
  fd.append('repotic',  document.getElementById('inp-repotic').files[0]);
  fd.append('ledger',   document.getElementById('inp-ledger').files[0]);
  fd.append('template', document.getElementById('inp-template').files[0]);

  const prefixEl = document.getElementById('inp-inv-prefix');
  const dateEl = document.getElementById('inp-inv-date');
  if (prefixEl) fd.append('inv_prefix', prefixEl.value.trim());
  if (dateEl)   fd.append('inv_date', dateEl.value.trim());

  try {
    const res = await fetch('/run', { method: 'POST', body: fd });
    const data = await res.json();

    if (!res.ok || data.error) {
      showError(data.error || res.statusText);
      return;
    }

    renderResults(data);
  } catch (err) {
    showError(String(err));
  } finally {
    btn.classList.remove('loading');
    btn.disabled = false;
    label.textContent = 'Re-process Data';
  }
}

function showError(msg) {
  const box = document.getElementById('errorBox');
  box.textContent = 'Error: ' + msg;
  box.style.display = 'block';
  box.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

// ── Render Results ───────────────────────────────────────────────────────────
function fmt(n) {
  if (n === null || n === undefined || (typeof n === 'number' && isNaN(n))) return '–';
  return Number(n).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function getMarketplaceTag(name) {
  if (!name) return '';
  const cls = { Flipkart: 'tag-flipkart', Meesho: 'tag-meesho', Snapdeal: 'tag-snapdeal' }[name] || '';
  return `<span class="mkt-tag ${cls}">${name}</span>`;
}

function renderResults(data) {
  document.getElementById('resultSummaryText').textContent =
    `${data.total_rows} rows extracted, mapped, and validated across ${data.summary.length} marketplaces.`;

  // Configure instant download button from base64 payload (stateless serverless compatible)
  const dlBtn = document.getElementById('downloadBtn');
  if (data.file_base64) {
    try {
      const byteChars = atob(data.file_base64);
      const byteNumbers = new Uint8Array(byteChars.length);
      for (let i = 0; i < byteChars.length; i++) {
        byteNumbers[i] = byteChars.charCodeAt(i);
      }
      const blob = new Blob([byteNumbers], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      const blobUrl = URL.createObjectURL(blob);
      dlBtn.href = blobUrl;
      dlBtn.download = 'Final_Output_Tamilnadu.xlsx';
    } catch (e) {
      console.error('Error creating download blob', e);
      dlBtn.href = '/download';
      dlBtn.download = 'Final_Output_Tamilnadu.xlsx';
    }
  } else {
    dlBtn.href = '/download';
    dlBtn.download = 'Final_Output_Tamilnadu.xlsx';
  }

  // Summary Cards
  const cardsContainer = document.getElementById('statCards');
  cardsContainer.innerHTML = '';
  data.summary.forEach(s => {
    const isOk = s.status === 'OK';
    cardsContainer.innerHTML += `
      <div class="stat-card">
        <div class="stat-top">
          <span class="stat-name">${s.sheet}</span>
          <span class="${isOk ? 'badge-ok' : 'badge-fail'}">${isOk ? 'Validated' : 'Mismatch'}</span>
        </div>
        <div>
          <div class="stat-count">${s.rows}</div>
          <div class="stat-count-sub">Rows Extracted</div>
        </div>
        <div class="stat-finances">
          <div>Taxable: <strong>₹${fmt(s.TaxableAmt)}</strong></div>
          <div>GST: ₹${fmt(s.IGSTAmt)} (I) | ₹${fmt(s.CGSTAmt)} (C) | ₹${fmt(s.SGSTAmt)} (S)</div>
          <div>Net Total: <strong>₹${fmt(s.Net_Amt)}</strong></div>
        </div>
      </div>
    `;
  });

  // Unmatched States
  const unmatchedBox = document.getElementById('unmatchedBox');
  const unmatchedTags = document.getElementById('unmatchedTags');
  if (data.unmatched && data.unmatched.length > 0) {
    unmatchedTags.innerHTML = data.unmatched.map(u => `
      <div class="unmatched-pill">
        ${getMarketplaceTag(u.sheet)}
        <span>${u.state}</span>
      </div>
    `).join('');
    unmatchedBox.style.display = 'block';
  } else {
    unmatchedBox.style.display = 'none';
  }

  // Preview Table
  document.getElementById('previewTitle').textContent =
    `Data Preview (First 10 of ${data.total_rows} Rows)`;

  const tbody = document.getElementById('previewBody');
  tbody.innerHTML = '';
  data.preview.forEach((row, idx) => {
    const salesLedger = row['Sales Ledger'] || row.Sales_Ledger;
    const billOfSupply = row['Bill of Supply'] || row.Bill_of_Supply || row.StateOfSupply;
    const blank = (v) => (v === null || v === undefined) ? '<span class="dim">–</span>' : v;

    tbody.innerHTML += `
      <tr>
        <td class="num dim">${idx + 1}</td>
        <td><strong style="color:#93c5fd; font-family:'JetBrains Mono', monospace; font-size:0.75rem;">${blank(row.InvNo)}</strong></td>
        <td><span style="color:#94a3b8; font-family:'JetBrains Mono', monospace; font-size:0.72rem;">${blank(row.Inv_Dt)}</span></td>
        <td>${getMarketplaceTag(row._marketplace)}</td>
        <td><span style="font-size:0.75rem; color:#cbd5e1; font-weight:600;">${blank(row.Vch_Type || 'Auto Sales')}</span></td>
        <td><strong>${blank(row.StateOfSupply)}</strong></td>
        <td><strong>${blank(billOfSupply)}</strong></td>
        <td class="dim">${blank(row.HSNCode)}</td>
        <td class="num">${row.Qty != null ? Number(row.Qty).toLocaleString('en-IN') : '–'}</td>
        <td class="num">${row.TaxPer != null ? row.TaxPer + '%' : '–'}</td>
        <td class="num">₹${fmt(row.TaxableAmt)}</td>
        <td class="num">₹${fmt(row.IGSTAmt)}</td>
        <td class="num">₹${fmt(row.SGSTAmt)}</td>
        <td class="num">₹${fmt(row.CGSTAmt)}</td>
        <td class="num"><strong>₹${fmt(row.Net_Amt)}</strong></td>
        <td class="cell-truncate" title="${row.Pty_Name || ''}">${blank(row.Pty_Name)}</td>
        <td class="cell-truncate" title="${salesLedger || ''}">${blank(salesLedger)}</td>
      </tr>
    `;
  });

  // Process Log
  const logBox = document.getElementById('logBox');
  logBox.innerHTML = data.logs.map(line => {
    if (line.startsWith('─') || line.startsWith('=')) {
      return `<div class="l-head">${line}</div>`;
    }
    if (/\bOK\b/.test(line) && !line.includes('[')) {
      return `<div class="l-ok">${line}</div>`;
    }
    if (line.includes('MISMATCH') || line.includes('[ERROR]')) {
      return `<div class="l-err">${line}</div>`;
    }
    if (line.includes('[UNMATCHED]') || line.includes('[WARNING]')) {
      return `<div class="l-warn">${line}</div>`;
    }
    if (line.includes('[SKIP]')) {
      return `<div class="l-mute">${line}</div>`;
    }
    return `<div>${line}</div>`;
  }).join('');

  document.getElementById('results').style.display = 'block';
  document.getElementById('results').scrollIntoView({ behavior: 'smooth', block: 'start' });
}
