let selectedFile = null;
let analysisData = null;
let charts = {};
let previewPage = 0;
let previewPageSize = 10;
let allPreviewRows = [];
let previewCols = [];
let uploadResponse = null;  // holds upload response with columns and suggestions

function showTab(name, btn) {
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('nav button').forEach(b => b.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  if (btn) btn.classList.add('active');
}

function handleFileSelect(e) {
  selectedFile = e.target.files[0];
  if (selectedFile) {
    document.getElementById('file-name').textContent = `Selected: ${selectedFile.name} (${(selectedFile.size / 1024).toFixed(1)} KB)`;
    document.getElementById('upload-btn').disabled = false;
    document.getElementById('upload-status').innerHTML = '';
  }
}

// Allow drag & drop
const zone = document.querySelector('.upload-zone');
zone.addEventListener('dragover', e => { e.preventDefault(); zone.style.background = 'var(--blue-light)'; });
zone.addEventListener('dragleave', () => { zone.style.background = ''; });
zone.addEventListener('drop', e => {
  e.preventDefault();
  zone.style.background = '';
  const file = e.dataTransfer.files[0];
  if (file) {
    selectedFile = file;
    document.getElementById('file-name').textContent = `Selected: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
    document.getElementById('upload-btn').disabled = false;
  }
});

async function uploadFile() {
  if (!selectedFile) return;
  const spinner = document.getElementById('spinner');
  const status = document.getElementById('upload-status');
  spinner.style.display = 'block';
  status.innerHTML = '';
  document.getElementById('upload-btn').disabled = true;

  const formData = new FormData();
  formData.append('file', selectedFile);

  try {
    const res = await fetch('/upload', { method: 'POST', body: formData });
    const data = await res.json();
    spinner.style.display = 'none';
    document.getElementById('upload-btn').disabled = false;

    if (data.error) {
      status.innerHTML = `<div class="alert alert-warn">⚠ ${data.error}</div>`;
      return;
    }

    uploadResponse = data;

    if (data.already_mapped) {
      status.innerHTML = `<div class="alert alert-success">✓ Columns matched automatically (${data.n} rows). Running analysis...</div>`;
      await runAnalysis(buildIdentityMapping(data.expected_columns));
    } else {
      status.innerHTML = `<div class="alert alert-info">ℹ ${data.n} rows loaded. Please map your columns before running analysis.</div>`;
      buildMappingForm(data);
      // switch to mapping tab
      document.querySelectorAll('nav button')[1].click();
    }
  } catch (err) {
    spinner.style.display = 'none';
    document.getElementById('upload-btn').disabled = false;
    status.innerHTML = `<div class="alert alert-warn">⚠ Upload failed: ${err.message}</div>`;
  }
}

function buildIdentityMapping(expectedCols) {
  const m = {};
  expectedCols.forEach(c => { m[c] = c; });
  return m;
}

const DIM_OF = col => {
  if (col.startsWith('SD')) return 'SD';
  if (col.startsWith('TQ')) return 'TQ';
  if (col.startsWith('PP')) return 'PP';
  if (col.startsWith('COM')) return 'COM';
  if (col.startsWith('SAT')) return 'SAT';
  return 'NPS';
};

function buildMappingForm(data) {
  const { file_columns, expected_columns, column_labels, suggestions } = data;
  const alertEl = document.getElementById('mapping-alert');
  const form = document.getElementById('mapping-form');

  const autoCount = Object.values(suggestions).filter(v => v !== null).length;

  // File columns reference panel
  const colPills = file_columns.map(c =>
    `<span class="col-pill" id="pill-${CSS.escape(c)}" title="${c}">${c}</span>`
  ).join('');

  alertEl.innerHTML = `
    <div class="alert alert-info" style="margin-bottom:10px;">
      <strong>${autoCount} of ${expected_columns.length}</strong> columns auto-matched.
      Use the dropdowns below to assign your file's columns to each survey variable. NPS is optional.
    </div>
    <div class="col-ref-box">
      <div class="col-ref-title">
        📋 Your file has <strong>${file_columns.length}</strong> columns — click any pill to copy the name:
      </div>
      <div class="col-ref-pills">${colPills}</div>
    </div>`;

  // Add pill click-to-copy behaviour
  setTimeout(() => {
    document.querySelectorAll('.col-pill').forEach(pill => {
      pill.addEventListener('click', () => {
        navigator.clipboard?.writeText(pill.title).catch(() => {});
        pill.classList.add('copied');
        setTimeout(() => pill.classList.remove('copied'), 1200);
      });
    });
  }, 0);

  const optionsHtml = '<option value="">— Skip —</option>' +
    file_columns.map(c => `<option value="${c}">${c}</option>`).join('');

  let html = '<div class="mapping-grid">';
  html += '<div class="mapping-header">Survey Variable</div>' +
          '<div class="mapping-header">Questionnaire Item</div>' +
          '<div class="mapping-header">Your File Column ↓ select to map</div>';

  expected_columns.forEach(col => {
    const dim = DIM_OF(col);
    const suggested = suggestions[col] || '';
    const isOptional = col === 'NPS';
    const matched = suggested ? 'auto-matched' : '';
    const desc = (column_labels[col] || '').split('–').slice(1).join('–').trim();
    html += `<div class="mapping-row">
      <div><span class="dim-label dim-${dim}">${dim}</span><strong>${col}</strong>
        ${isOptional ? '<em style="color:var(--muted);font-size:0.75rem"> (optional)</em>' : ''}
      </div>
      <div style="color:var(--muted);font-size:0.83rem;line-height:1.4">${desc}</div>
      <div>
        <select id="map-${col}" class="${matched}" onchange="onMapChange(this,'${col}')">
          ${optionsHtml}
        </select>
        <div id="map-hint-${col}" class="map-hint"></div>
      </div>
    </div>`;
  });

  html += '</div>';
  form.innerHTML = html;

  // Set suggested values & wire change events
  expected_columns.forEach(col => {
    const sel = document.getElementById(`map-${col}`);
    if (suggestions[col]) {
      sel.value = suggestions[col];
      markPillUsed(suggestions[col]);
    }
  });

  updateUnmappedCount();
}

function onMapChange(sel, col) {
  sel.classList.toggle('auto-matched', sel.value !== '');
  // show a small hint with the selected column name
  const hint = document.getElementById(`map-hint-${col}`);
  if (hint) hint.textContent = sel.value ? `✓ mapped to: ${sel.value}` : '';
  updateUnmappedCount();
}

function markPillUsed(colName) {
  const pill = document.getElementById(`pill-${CSS.escape(colName)}`);
  if (pill) pill.classList.add('used');
}

function updateUnmappedCount() {
  if (!uploadResponse) return;
  const required = uploadResponse.expected_columns.filter(c => c !== 'NPS');
  const unmapped = required.filter(c => {
    const sel = document.getElementById(`map-${c}`);
    return !sel || !sel.value;
  });
  const statusEl = document.getElementById('mapping-status');
  if (unmapped.length === 0) {
    statusEl.innerHTML = '<div class="alert alert-success">✓ All required columns mapped. Ready to run analysis.</div>';
  } else {
    statusEl.innerHTML = `<div class="alert alert-info">ℹ ${unmapped.length} required column(s) still need mapping: <strong>${unmapped.join(', ')}</strong></div>`;
  }
}

async function confirmMapping() {
  if (!uploadResponse) return;
  const mapping = {};
  uploadResponse.expected_columns.forEach(col => {
    const sel = document.getElementById(`map-${col}`);
    if (sel && sel.value) mapping[col] = sel.value;
  });

  const required = uploadResponse.expected_columns.filter(c => c !== 'NPS');
  const missing = required.filter(c => !mapping[c]);
  if (missing.length > 0) {
    document.getElementById('mapping-status').innerHTML =
      `<div class="alert alert-warn">⚠ Still need to map: <strong>${missing.join(', ')}</strong><br>
       Select a column from the dropdown for each item above.</div>`;
    return;
  }

  document.getElementById('mapping-status').innerHTML = '';
  await runAnalysis(mapping);
}

async function runAnalysis(mapping) {
  const spinner = document.getElementById('mapping-spinner');
  const status = document.getElementById('mapping-status');
  if (spinner) spinner.style.display = 'block';

  try {
    const res = await fetch('/analyze', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ mapping }),
    });
    const data = await res.json();
    if (spinner) spinner.style.display = 'none';

    if (data.error) {
      if (status) status.innerHTML = `<div class="alert alert-warn">⚠ ${data.error}</div>`;
      document.getElementById('upload-status').innerHTML = `<div class="alert alert-warn">⚠ ${data.error}</div>`;
      return;
    }

    analysisData = data;
    if (data.errors && data.errors.length > 0) {
      const msg = data.errors.map(e => `<div class="alert alert-warn">⚠ ${e}</div>`).join('');
      if (status) status.innerHTML = msg;
    } else {
      const msg = `<div class="alert alert-success">✓ Analysis complete. ${data.n} respondents processed.</div>`;
      if (status) status.innerHTML = msg;
      document.getElementById('upload-status').innerHTML = msg;
    }

    renderAll(data);
    // Navigate to analysis tab
    document.querySelectorAll('nav button')[3].click();
  } catch (err) {
    if (spinner) spinner.style.display = 'none';
    const msg = `<div class="alert alert-warn">⚠ Analysis failed: ${err.message}</div>`;
    if (status) status.innerHTML = msg;
  }
}

function renderAll(data) {
  renderPreview(data);
  renderKPIs(data);
  renderInsights(data);
  renderDescriptive(data);
  renderCSI(data);
  renderReliability(data);
  renderRegression(data);
  renderCorrelation(data);
  renderNPS(data);
  renderCharts(data);
}
function makeTable(headers, rows, cellFn) {
  let html = '<table><tr>' + headers.map(h => `<th>${h}</th>`).join('') + '</tr>';
  rows.forEach(row => {
    html += '<tr>' + row.map((cell, i) => `<td>${cellFn ? cellFn(cell, i, row) : cell}</td>`).join('') + '</tr>';
  });
  return html + '</table>';
}

function renderPreview(data) {
  previewCols = data.columns;
  allPreviewRows = data.preview.map(r => previewCols.map(c => r[c] !== undefined ? r[c] : ''));
  previewPage = 0;
  document.getElementById('preview-info').textContent =
    `${data.n} rows × ${previewCols.length} columns`;
  renderPreviewPage();
}

function renderPreviewPage() {
  const start = previewPage * previewPageSize;
  const end = Math.min(start + previewPageSize, allPreviewRows.length);
  const pageRows = allPreviewRows.slice(start, end);
  document.getElementById('preview-table').innerHTML = makeTable(previewCols, pageRows);
  const totalPages = Math.ceil(allPreviewRows.length / previewPageSize);
  document.getElementById('page-info').textContent =
    `Page ${previewPage + 1} of ${totalPages} (rows ${start + 1}–${end} of ${allPreviewRows.length})`;
  document.getElementById('prev-page').disabled = previewPage === 0;
  document.getElementById('next-page').disabled = end >= allPreviewRows.length;
}

function changePage(dir) {
  const totalPages = Math.ceil(allPreviewRows.length / previewPageSize);
  previewPage = Math.max(0, Math.min(previewPage + dir, totalPages - 1));
  renderPreviewPage();
}

function changePageSize(val) {
  previewPageSize = parseInt(val);
  previewPage = 0;
  renderPreviewPage();
}

function renderKPIs(data) {
  const grid = document.getElementById('kpi-grid');
  const csiClass = data.overall_csi >= 80 ? 'green' : data.overall_csi >= 60 ? '' : 'orange';
  grid.innerHTML = `
    <div class="kpi"><div class="val">${data.n}</div><div class="lbl">Respondents</div></div>
    <div class="kpi ${csiClass}"><div class="val">${data.overall_csi}%</div><div class="lbl">Overall CSI</div></div>
    <div class="kpi"><div class="val">${data.overall_interp}</div><div class="lbl">Satisfaction Level</div></div>
    ${data.regression ? `<div class="kpi"><div class="val">${(data.regression.r_squared * 100).toFixed(1)}%</div><div class="lbl">Variance Explained (R²)</div></div>` : ''}
    ${data.nps ? `<div class="kpi"><div class="val">${data.nps.nps_score}</div><div class="lbl">NPS Score</div></div>` : ''}
  `;
}

function renderInsights(data) {
  const card = document.getElementById('insights-card');
  const list = document.getElementById('insights-list');
  if (data.insights && data.insights.length > 0) {
    card.style.display = '';
    list.innerHTML = data.insights.map(i => `<li>${i}</li>`).join('');
  }
}

function renderDescriptive(data) {
  const dimRows = data.descriptive_dims.map(r => [r.Dimension, r.Mean, r['Std Dev']]);
  document.getElementById('desc-dims-table').innerHTML =
    makeTable(['Dimension', 'Mean', 'Std Dev'], dimRows);

  const itemRows = data.descriptive_items.map(r => [r.Item, r.Description, r.Mean, r['Std Dev'], r.Min, r.Max]);
  document.getElementById('desc-items-table').innerHTML =
    makeTable(['Item', 'Description', 'Mean', 'Std Dev', 'Min', 'Max'], itemRows);
}

function csiColor(val) {
  if (val >= 80) return '<span class="badge badge-green">' + val + '</span>';
  if (val >= 60) return '<span class="badge badge-blue">' + val + '</span>';
  if (val >= 40) return '<span class="badge badge-orange">' + val + '</span>';
  return '<span class="badge badge-red">' + val + '</span>';
}

function renderCSI(data) {
  const rows = data.csi.map(r => [r.Dimension, r['Mean Score'], r['CSI (%)'], r.Interpretation]);
  document.getElementById('csi-table').innerHTML = makeTable(
    ['Dimension', 'Mean Score', 'CSI (%)', 'Interpretation'],
    rows,
    (cell, i) => i === 2 ? csiColor(cell) : cell
  );
}

function alphaColor(val) {
  if (val >= 0.9) return '<span class="badge badge-green">' + val + '</span>';
  if (val >= 0.8) return '<span class="badge badge-green">' + val + '</span>';
  if (val >= 0.7) return '<span class="badge badge-blue">' + val + '</span>';
  return '<span class="badge badge-orange">' + val + '</span>';
}

function renderReliability(data) {
  const rows = data.reliability.map(r => [r.Dimension, r['Cronbach Alpha'], r.Interpretation]);
  document.getElementById('reliability-table').innerHTML = makeTable(
    ['Dimension', 'Cronbach Alpha', 'Interpretation'],
    rows,
    (cell, i) => i === 1 ? alphaColor(cell) : cell
  );
}

function renderRegression(data) {
  const reg = data.regression;
  if (!reg) {
    document.getElementById('regression-summary').textContent = 'Insufficient data for regression.';
    return;
  }
  document.getElementById('regression-summary').innerHTML =
    `N = ${reg.n} &nbsp;|&nbsp; R² = ${reg.r_squared} &nbsp;|&nbsp; Adj. R² = ${reg.adj_r_squared} &nbsp;|&nbsp; F = ${reg.f_statistic} &nbsp;|&nbsp; p = ${reg.f_pvalue}`;

  const rows = reg.coefficients.map(r => [r.Variable, r.Beta, r['Std Error'], r['t-value'], r['p-value'], r.Significant]);
  document.getElementById('regression-table').innerHTML = makeTable(
    ['Variable', 'Beta', 'Std Error', 't-value', 'p-value', 'Significant'],
    rows,
    (cell, i) => i === 5
      ? `<span class="${cell === 'Yes' ? 'sig-yes' : 'sig-no'}">${cell}</span>`
      : cell
  );
}

function renderCorrelation(data) {
  const corr = data.correlation;
  const labels = Object.keys(corr);
  let html = '<table><tr><th></th>' + labels.map(l => `<th>${l}</th>`).join('') + '</tr>';
  labels.forEach(row => {
    html += `<tr><th>${row}</th>`;
    labels.forEach(col => {
      const val = corr[row][col];
      const abs = Math.abs(val);
      const bg = abs > 0.7 ? '#bbdefb' : abs > 0.4 ? '#e3f2fd' : '';
      html += `<td style="background:${bg};text-align:center">${val}</td>`;
    });
    html += '</tr>';
  });
  html += '</table>';
  document.getElementById('correlation-table').innerHTML = html;
}

function renderNPS(data) {
  if (!data.nps) return;
  const nps = data.nps;
  const card = document.getElementById('nps-card');
  card.style.display = '';
  card.querySelector('#nps-content').innerHTML = `
    <div class="kpi-grid">
      <div class="kpi"><div class="val">${nps.nps_score}</div><div class="lbl">NPS Score</div></div>
      <div class="kpi green"><div class="val">${nps.promoters}</div><div class="lbl">Promoters (9–10)</div></div>
      <div class="kpi"><div class="val">${nps.passives}</div><div class="lbl">Passives (7–8)</div></div>
      <div class="kpi orange"><div class="val">${nps.detractors}</div><div class="lbl">Detractors (0–6)</div></div>
    </div>
  `;
}

function destroyChart(id) {
  if (charts[id]) { charts[id].destroy(); delete charts[id]; }
}

// Chart configs stored so we can re-render on type change
const chartConfigs = {};

const COLORS = ['#1a6fa8','#2e7d32','#e65100','#6a1b9a','#00838f'];

// Types supported per chart (heatmap is always scatter, excluded from switcher)
const CHART_TYPES = {
  mean:  ['bar', 'line', 'radar', 'polarArea', 'pie', 'doughnut'],
  csi:   ['bar', 'line', 'radar', 'polarArea', 'pie', 'doughnut'],
  reg:   ['bar', 'line'],
  alpha: ['bar', 'line', 'radar', 'polarArea'],
};

function buildDataset(type, label, data, colors) {
  const isCircular = ['pie','doughnut','polarArea'].includes(type);
  return {
    label,
    data,
    backgroundColor: isCircular ? colors : colors,
    borderColor: isCircular ? colors : colors.map(c => c),
    borderWidth: type === 'line' ? 2 : 1,
    borderRadius: type === 'bar' ? 6 : 0,
    fill: type === 'line' ? false : undefined,
    pointRadius: type === 'line' ? 5 : undefined,
    tension: type === 'line' ? 0.3 : undefined,
  };
}

function buildOptions(type, extraOpts) {
  const isCircular = ['pie','doughnut','polarArea'].includes(type);
  const isRadar = type === 'radar';
  const base = {
    responsive: true,
    plugins: { legend: { display: isCircular || isRadar } },
  };
  if (!isCircular && !isRadar) {
    base.scales = extraOpts?.scales || {};
  }
  return base;
}

function renderSwitchableChart(id, type, label, labels, data, colors, extraOpts) {
  destroyChart(id);
  const dataset = buildDataset(type, label, data, colors);
  const options = buildOptions(type, extraOpts);
  charts[id] = new Chart(document.getElementById('chart-' + id), {
    type,
    data: { labels, datasets: [dataset] },
    options,
  });
}

function changeChartType(id, newType) {
  const cfg = chartConfigs[id];
  if (!cfg) return;
  cfg.type = newType;
  renderSwitchableChart(id, newType, cfg.label, cfg.labels, cfg.data, cfg.colors, cfg.extraOpts);
}

function renderCharts(data) {
  const dimLabels = data.descriptive_dims.map(d => d.Dimension);
  const dimMeans  = data.descriptive_dims.map(d => d.Mean);
  const csiVals   = data.csi.map(d => d['CSI (%)']);
  const csiLabels = data.csi.map(d => d.Dimension);

  // Store configs for re-render on type switch
  chartConfigs['mean'] = {
    type: 'bar', label: 'Mean Score', labels: dimLabels, data: dimMeans, colors: COLORS,
    extraOpts: { scales: { y: { min: 0, max: 5, ticks: { stepSize: 1 } } } }
  };
  chartConfigs['csi'] = {
    type: 'bar', label: 'CSI (%)', labels: csiLabels, data: csiVals, colors: COLORS,
    extraOpts: { scales: { y: { min: 0, max: 100 } } }
  };

  if (data.regression) {
    const regLabels = data.regression.coefficients.map(c => c.Variable);
    const regBetas  = data.regression.coefficients.map(c => c.Beta);
    const regColors = data.regression.coefficients.map(c => c.Significant === 'Yes' ? '#2e7d32' : '#aaa');
    chartConfigs['reg'] = {
      type: 'bar', label: 'Beta Coefficient', labels: regLabels, data: regBetas, colors: regColors,
      extraOpts: {}
    };
  }

  const alphaLabels = data.reliability.map(r => r.Dimension);
  const alphaVals   = data.reliability.map(r => r['Cronbach Alpha']);
  chartConfigs['alpha'] = {
    type: 'bar', label: "Cronbach's Alpha", labels: alphaLabels, data: alphaVals, colors: COLORS,
    extraOpts: { scales: { y: { min: 0, max: 1 } } }
  };

  // Render all switchable charts
  ['mean','csi','alpha'].forEach(id => {
    const c = chartConfigs[id];
    renderSwitchableChart(id, c.type, c.label, c.labels, c.data, c.colors, c.extraOpts);
  });
  if (chartConfigs['reg']) {
    const c = chartConfigs['reg'];
    renderSwitchableChart('reg', c.type, c.label, c.labels, c.data, c.colors, c.extraOpts);
  }

  // Populate dropdowns with available types
  Object.keys(CHART_TYPES).forEach(id => {
    const sel = document.getElementById('type-' + id);
    if (!sel) return;
    sel.innerHTML = CHART_TYPES[id].map(t =>
      `<option value="${t}">${t.charAt(0).toUpperCase() + t.slice(1)}</option>`
    ).join('');
    sel.value = 'bar';
  });

  // Correlation heatmap — fixed scatter, no switcher
  renderCorrHeatmap(data);
}

function renderCorrHeatmap(data) {
  destroyChart('corr');
  const corrKeys = Object.keys(data.correlation);
  const n = corrKeys.length;
  const corrDatasets = [];
  corrKeys.forEach((row, ri) => {
    corrKeys.forEach((col, ci) => {
      corrDatasets.push({ x: ci, y: ri, v: data.correlation[row][col] });
    });
  });

  charts['corr'] = new Chart(document.getElementById('chart-corr'), {
    type: 'scatter',
    data: {
      datasets: [{
        data: corrDatasets.map(d => ({ x: d.x, y: d.y })),
        backgroundColor: corrDatasets.map(d => {
          const v = d.v;
          if (v >= 0.7)  return 'rgba(26,111,168,0.85)';
          if (v >= 0.4)  return 'rgba(26,111,168,0.45)';
          if (v >= 0)    return 'rgba(26,111,168,0.15)';
          return 'rgba(230,81,0,0.4)';
        }),
        pointRadius: corrDatasets.map(() => 28),
        pointStyle: 'rect',
      }]
    },
    options: {
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => {
              const d = corrDatasets[ctx.dataIndex];
              return `${corrKeys[d.x]} vs ${corrKeys[d.y]}: ${d.v}`;
            }
          }
        }
      },
      scales: {
        x: { min: -0.5, max: n - 0.5, ticks: { callback: v => corrKeys[v] || '' }, grid: { display: false } },
        y: { min: -0.5, max: n - 0.5, ticks: { callback: v => corrKeys[v] || '' }, grid: { display: false } }
      }
    }
  });
}