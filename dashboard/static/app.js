/**
 * app.js  –  FinFormat Pro PDF-Converter Style UI Logic
 */

'use strict';

/* ──────────────────────────────────────
   State
────────────────────────────────────── */
const state = {
  sessionId:  null,
  outputFile: null,
  company:    null,
  files:      [], // array of { name, size, type }
  preview:    null,
  charts:     [],
};

/* ──────────────────────────────────────
   DOM helpers
────────────────────────────────────── */
const $ = id => document.getElementById(id);
window.$ = $;
const show = id => { const el = $(id); if (el) el.style.display = ''; };
const hide = id => { const el = $(id); if (el) el.style.display = 'none'; };

/* ──────────────────────────────────────
   Status bar
────────────────────────────────────── */
function setStatus(text, color = '#4fc3f7') {
  const t = document.querySelector('.status-text');
  const d = document.querySelector('.status-dot');
  if (t) t.textContent = text;
  if (d) d.style.background = color;
}

/* ──────────────────────────────────────
   Dropzone & File Input Initialization
────────────────────────────────────── */
function initDropzone() {
  const dz = $('dropzone');
  const fileInput = $('fileInput');

  // Trigger file dialog on dropzone click – but NOT when clicking the select button
  // (the button has its own onclick handler to avoid double-triggering)
  dz.addEventListener('click', (e) => {
    // Walk up to see if click originated from the select button
    if (!e.target.closest('.select-files-btn')) {
      fileInput.click();
    }
  });

  fileInput.addEventListener('change', () => {
    if (fileInput.files.length > 0) {
      uploadFiles([...fileInput.files]);
      fileInput.value = ''; // reset so same file can be re-selected
    }
  });

  // Drag and drop events
  dz.addEventListener('dragover', (e) => {
    e.preventDefault();
    dz.classList.add('drag-over');
  });

  dz.addEventListener('dragleave', () => {
    dz.classList.remove('drag-over');
  });

  dz.addEventListener('drop', (e) => {
    e.preventDefault();
    dz.classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) {
      uploadFiles([...e.dataTransfer.files]);
    }
  });
}

/* ──────────────────────────────────────
   Upload Files to Queue
────────────────────────────────────── */
async function uploadFiles(filesList) {
  const csvFiles = filesList.filter(f => f.name.toLowerCase().endsWith('.csv'));
  if (csvFiles.length === 0) {
    alert('Please select valid CSV files. Only .csv files are accepted.');
    return;
  }

  const formData = new FormData();
  if (state.sessionId) {
    formData.append('session_id', state.sessionId);
  }
  csvFiles.forEach(f => formData.append('files', f));

  showLoading(`Uploading ${csvFiles.length} file(s)…`);
  setStatus('Uploading…', '#ffa726');

  try {
    const resp = await fetch('/upload', {
      method: 'POST',
      body: formData,
      credentials: 'same-origin',
    });
    hideLoading();

    if (!resp.ok) {
      let errMsg = 'Upload failed';
      try { const err = await resp.json(); errMsg = err.error || errMsg; } catch(_) {}
      throw new Error(errMsg);
    }

    const data = await resp.json();
    state.sessionId = data.session_id;
    state.files     = data.files || [];

    renderFileQueue();
    setStatus(`${state.files.length} file(s) ready ✓`, '#66bb6a');
  } catch (err) {
    hideLoading();
    setStatus('Upload error', '#ef5350');
    alert('Upload error: ' + err.message);
    console.error('Upload error:', err);
  }
}

/* ──────────────────────────────────────
   Remove File from Queue
────────────────────────────────────── */
async function removeFile(filename) {
  if (!state.sessionId) return;

  showLoading('Updating file queue…');
  try {
    const resp = await fetch('/remove-file', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ session_id: state.sessionId, filename: filename })
    });
    hideLoading();

    if (resp.ok) {
      const data = await resp.json();
      state.files = data.files || [];
      renderFileQueue();
    }
  } catch (err) {
    hideLoading();
    console.error('Failed to remove file:', err);
  }
}

/* ──────────────────────────────────────
   Render File Queue List
────────────────────────────────────── */
function renderFileQueue() {
  const list    = $('fileQueueList');
  const countEl = $('queueCount');
  const btn     = $('convertBtn');

  if (state.files.length === 0) {
    hide('fileQueueSection');   // pass string ID, not element
    btn.disabled = true;
    return;
  }

  show('fileQueueSection');     // pass string ID, not element
  countEl.textContent = state.files.length;
  btn.disabled = false;

  list.innerHTML = '';
  state.files.forEach(file => {
    const card = document.createElement('div');
    card.className = 'file-card';
    card.innerHTML = `
      <div class="file-card-left">
        <div class="file-details">
          <div class="file-name" title="${file.name}">${file.name}</div>
          <div class="file-meta">
            <span class="file-size">${file.size}</span>
            <span class="file-type-badge">${file.type}</span>
          </div>
        </div>
      </div>
      <button class="file-remove-btn" title="Remove file" onclick="removeFile('${file.name.replace(/'/g, "\\'")}')">✕</button>
    `;
    list.appendChild(card);
  });
}

/* ──────────────────────────────────────
   AI Mode Toggle
────────────────────────────────────── */
function toggleAiMode() {
  const isAi = $('aiModeToggle').checked;
  if (isAi) {
    show('aiConfigPanel');
  } else {
    hide('aiConfigPanel');
    hide('agentLogWindow');
  }
}

/* ──────────────────────────────────────
   Run Conversion Process
────────────────────────────────────── */
async function runConversion() {
  if (state.files.length === 0) {
    alert('Please add at least one CSV file.');
    return;
  }

  const company = $('companyName').value.trim() || 'Company';
  state.company = company;
  
  const isAi = $('aiModeToggle').checked;
  const apiKey = $('apiKey').value.trim();

  if (isAi && !apiKey) {
    alert('Please enter your Gemini API Key to use the AI Agent.');
    return;
  }

  const endpoint = isAi ? '/agent-process' : '/process';
  const payload = { session_id: state.sessionId, company_name: company };
  if (isAi) payload.api_key = apiKey;

  setStatus('Converting…', '#ffa726');
  
  if (isAi) {
    show('agentLogWindow');
    $('agentLogContent').innerHTML = '<div>Connecting to AI Agent...</div>';
    
    // Start SSE listener
    const evtSource = new EventSource(`/agent-process-stream/${state.sessionId}`);
    evtSource.onmessage = function(e) {
      const data = JSON.parse(e.data);
      if (data.message) {
        const logContent = $('agentLogContent');
        logContent.innerHTML += `<div>> ${data.message}</div>`;
        logContent.scrollTop = logContent.scrollHeight;
      }
    };
    evtSource.onerror = function() {
      evtSource.close();
    };
  } else {
    showLoading('Parsing CSV files and generating model…');
  }

  try {
    const resp = await fetch(endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      credentials: 'same-origin',
      body: JSON.stringify(payload),
    });

    hideLoading();

    if (!resp.ok) {
      let errMsg = 'Conversion failed';
      try { const err = await resp.json(); errMsg = err.error || errMsg; } catch(_) {}
      throw new Error(errMsg);
    }

    const result = await resp.json();
    state.outputFile = result.output_file;
    state.preview    = result.preview;

    renderResults(result);

    show('resultsCard');
    $('resultsCard').scrollIntoView({ behavior: 'smooth' });
    setStatus('Converted ✓', '#66bb6a');
  } catch (err) {
    hideLoading();
    setStatus('Error', '#ef5350');
    if (isAi) {
      const logContent = $('agentLogContent');
      logContent.innerHTML += `<div style="color:#ef5350;">> Error: ${err.message}</div>`;
      logContent.scrollTop = logContent.scrollHeight;
    }
    alert('Conversion error: ' + err.message);
    console.error('Conversion error:', err);
  }
}

/* ──────────────────────────────────────
   Render Results & Tables
────────────────────────────────────── */
function renderResults(result) {
  $('resultsTitle').textContent = `${result.company} – Model Generated`;
  $('downloadInfo').textContent =
    `${result.periods.length} periods parsed from ${state.files.length} file(s)`;
  $('downloadBtn').dataset.sessionId = result.session_id;
  $('downloadBtn').dataset.file      = result.output_file;

  const p = result.preview;

  // Tables
  if (p.quarterly) renderTable('table-quarterly', p.quarterly, true);
  if (p.pl)        renderTable('table-pl',        p.pl,        false);
  if (p.bs)        renderTable('table-bs',        p.bs,        false);
  if (p.cf)        renderTable('table-cf',        p.cf,        false);
  if (p.ratios)    renderTable('table-ratios',    p.ratios,    false);

  // Charts
  setTimeout(() => renderCharts(p.chart_data, p.annual_periods), 100);
}

function renderTable(containerId, tableData, isQuarterly) {
  const container = $(containerId);
  if (!container || !tableData) return;

  const { headers, rows } = tableData;
  const annualCols = new Set();

  if (!isQuarterly) {
    headers.forEach((h, i) => { if (i > 0) annualCols.add(i); });
  }

  let html = '<table><thead><tr>';
  headers.forEach((h, i) => {
    const cls = annualCols.has(i) ? ' class="annual-col"' : '';
    html += `<th${cls}>${h}</th>`;
  });
  html += '</tr></thead><tbody>';

  rows.forEach(row => {
    html += '<tr>';
    row.forEach((cell, i) => {
      const cls = annualCols.has(i) ? ' class="annual-col"' : '';
      let valCls = '';
      if (i > 0 && typeof cell === 'string') {
        const num = parseFloat(cell.replace(/,/g, '').replace('%',''));
        if (!isNaN(num)) valCls = num >= 0 ? ' val-pos' : ' val-neg';
      }
      html += `<td${cls}><span class="${valCls.trim()}">${cell}</span></td>`;
    });
    html += '</tr>';
  });

  html += '</tbody></table>';
  container.innerHTML = html;
}

/* ──────────────────────────────────────
   Render Charts
────────────────────────────────────── */
function renderCharts(chartData, periods) {
  state.charts.forEach(c => c.destroy());
  state.charts = [];

  const PALETTE = {
    cyan:   { border: '#4fc3f7', bg: 'rgba(79,195,247,.15)' },
    gold:   { border: '#e8b86d', bg: 'rgba(232,184,109,.12)' },
    green:  { border: '#66bb6a', bg: 'rgba(102,187,106,.12)' },
    orange: { border: '#ffa726', bg: 'rgba(255,167,38,.12)' },
    purple: { border: '#ab47bc', bg: 'rgba(171,71,188,.12)' },
  };

  const defaultOptions = {
    responsive: true,
    plugins: {
      legend: { labels: { color: '#8ba3c0', font: { family: 'Inter', size: 11 } } },
      tooltip: {
        backgroundColor: '#0d1420',
        borderColor: '#2d4060',
        borderWidth: 1,
        titleColor: '#e8edf5',
        bodyColor: '#8ba3c0',
      },
    },
    scales: {
      x: { ticks: { color: '#4a6480', font: { size: 10 } }, grid: { color: '#1e2d45' } },
      y: { ticks: { color: '#4a6480', font: { size: 10 } }, grid: { color: '#1e2d45' } },
    },
  };

  function makeChart(canvasId, config) {
    const canvas = $(canvasId);
    if (!canvas) return null;
    const ctx = canvas.getContext('2d');
    const chart = new Chart(ctx, config);
    state.charts.push(chart);
    return chart;
  }

  // Revenue chart
  const rev = chartData['gross_sales'];
  if (rev) {
    makeChart('chartRevenue', {
      type: 'bar',
      data: {
        labels: rev.periods,
        datasets: [{
          label: 'Gross Sales (Rs. Cr.)',
          data: rev.values,
          backgroundColor: PALETTE.cyan.bg,
          borderColor: PALETTE.cyan.border,
          borderWidth: 2,
          borderRadius: 4,
        }],
      },
      options: { ...defaultOptions },
    });
  }

  // EBITDA + PAT chart
  const ebitda = chartData['ebitda'];
  const pat    = chartData['pat'];
  if (ebitda || pat) {
    const datasets = [];
    if (ebitda) datasets.push({
      label: 'EBITDA', data: ebitda.values,
      borderColor: PALETTE.gold.border, backgroundColor: PALETTE.gold.bg,
      borderWidth: 2, fill: true, tension: .4,
    });
    if (pat) datasets.push({
      label: 'PAT', data: pat.values,
      borderColor: PALETTE.green.border, backgroundColor: PALETTE.green.bg,
      borderWidth: 2, fill: true, tension: .4,
    });
    makeChart('chartProfits', {
      type: 'line',
      data: { labels: (ebitda || pat).periods, datasets },
      options: { ...defaultOptions },
    });
  }

  // EBITDA Margin
  const margin = chartData['ebitda_margin'];
  if (margin) {
    makeChart('chartMargin', {
      type: 'line',
      data: {
        labels: margin.periods,
        datasets: [{
          label: 'EBITDA Margin (%)',
          data: margin.values?.map(v => v !== null ? +(v * 100).toFixed(1) : null),
          borderColor: PALETTE.orange.border, backgroundColor: PALETTE.orange.bg,
          borderWidth: 2, fill: true, tension: .4, pointRadius: 3,
        }],
      },
      options: {
        ...defaultOptions,
        scales: {
          ...defaultOptions.scales,
          y: { ...defaultOptions.scales.y, ticks: { ...defaultOptions.scales.y.ticks, callback: v => v + '%' } },
        },
      },
    });
  }

  // PAT chart
  if (pat) {
    makeChart('chartEPS', {
      type: 'bar',
      data: {
        labels: pat.periods,
        datasets: [{
          label: 'PAT (Rs. Cr.)',
          data: pat.values,
          backgroundColor: PALETTE.purple.bg,
          borderColor: PALETTE.purple.border,
          borderWidth: 2,
          borderRadius: 4,
        }],
      },
      options: { ...defaultOptions },
    });
  }
}

/* ──────────────────────────────────────
   Download Excel
────────────────────────────────────── */
function downloadExcel() {
  const btn = $('downloadBtn');
  const sid  = btn.dataset.sessionId || state.sessionId;
  const file = btn.dataset.file      || state.outputFile;
  if (!sid || !file) { alert('No file to download.'); return; }
  window.location.href = `/download/${sid}/${file}`;
}

/* ──────────────────────────────────────
   Tab Switching
────────────────────────────────────── */
function initTabs() {
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      const panelName = tab.dataset.tab;
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      const panel = $(`tab-${panelName}`);
      if (panel) panel.classList.add('active');

      if (panelName === 'charts' && state.preview) {
        renderCharts(state.preview.chart_data, state.preview.annual_periods);
      }
    });
  });
}

/* ──────────────────────────────────────
   Loading & Reset
────────────────────────────────────── */
function showLoading(text) {
  $('spinnerText').textContent = text || 'Processing…';
  show('loadingOverlay');
}
function hideLoading() { hide('loadingOverlay'); }

function resetConverter() {
  state.files = [];
  state.sessionId = null;
  state.outputFile = null;
  renderFileQueue();
  hide('resultsCard');
  $('converterBox').scrollIntoView({ behavior: 'smooth' });
  setStatus('Ready', '#66bb6a');
}

/* ──────────────────────────────────────
   Initialization
────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  initDropzone();
  initTabs();
  renderFileQueue();
});
