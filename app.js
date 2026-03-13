/* ========================================================
   DataViz Pro — app.js
   Pure JS: SheetJS + Chart.js, no build tools required
   ======================================================== */

'use strict';

// ── State ──────────────────────────────────────────────────────────────────
const pqaState = {
  data: [],
  columns: [],
  sheetName: null,
};

const state = {
  workbook: null,       // XLSX workbook
  sheetName: null,      // active sheet
  rawData: [],          // original parsed rows (array of objects)
  filteredData: [],     // after applying filters
  columns: [],          // [{ name, type }]
  filters: {},          // { colName: value }
  sortCol: null,
  sortDir: 'asc',
  page: 0,
  pageSize: 50,
  searchQuery: '',
  charts: {},           // { status, platform }
  fileHandle: null,     // FileSystemFileHandle for live-sync
  lastModified: null,   // last known file modification timestamp
  syncInterval: null,   // setInterval ID
};

// ── DOM refs ───────────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);

const ui = {
  uploadScreen: $('upload-screen'),
  dashboard: $('dashboard'),
  loading: $('loading'),
  dropZone: $('drop-zone'),
  fileInput: $('file-input'),
  fileNameDisplay: $('file-name-display'),
  sheetTabs: $('sheet-tabs'),
  fieldList: $('field-list'),
  filterList: $('filter-list'),
  kpiRow: $('kpi-row'),
  tableHead: $('table-head'),
  tableBody: $('table-body'),
  tableSearch: $('table-search'),
  rowCount: $('row-count'),
  pageInfo: $('page-info'),
  btnPrev: $('btn-prev'),
  btnNext: $('btn-next'),
  btnNewFile: $('btn-new-file'),
  btnExport: $('btn-export'),
  btnReset: $('btn-reset-filters'),
  statusCol: $('status-col'),
  platformCol: $('platform-col'),
  pqaStatusCol:   $('pqa-status-col'),
  pqaPlatformCol: $('pqa-platform-col'),
};

// ── Colour palette ─────────────────────────────────────────────────────────
const PALETTE = [
  '#F2C811','#6366f1','#10b981','#f59e0b','#ef4444',
  '#3b82f6','#8b5cf6','#ec4899','#14b8a6','#f97316',
  '#22c55e','#06b6d4','#a855f7','#eab308','#64748b',
];

const CHART_DEFAULTS = {
  responsive: true,
  maintainAspectRatio: false,
  plugins: { legend: { labels: { font: { size: 11 }, boxWidth: 12 } } },
};

// ── Utility: format numbers (Indian style) ─────────────────────────────────
function formatNumber(n) {
  if (n == null || isNaN(n)) return '—';
  const abs = Math.abs(n);
  if (abs >= 1e7) return (n / 1e7).toFixed(2) + ' Cr';
  if (abs >= 1e5) return (n / 1e5).toFixed(2) + ' L';
  if (abs >= 1e3) return (n / 1e3).toFixed(1) + 'K';
  return Number.isInteger(n) ? n.toString() : n.toFixed(2);
}

function rawNumber(n) {
  if (n == null || n === '') return NaN;
  const parsed = parseFloat(String(n).replace(/,/g, ''));
  return isNaN(parsed) ? NaN : parsed;
}

// ── Type detection ─────────────────────────────────────────────────────────
function detectType(values) {
  const sample = values.filter(v => v != null && v !== '').slice(0, 50);
  if (!sample.length) return 'text';
  let numCount = 0, dateCount = 0;
  for (const v of sample) {
    if (!isNaN(rawNumber(v))) numCount++;
    else if (!isNaN(Date.parse(String(v)))) dateCount++;
  }
  if (numCount / sample.length > 0.7) return 'number';
  if (dateCount / sample.length > 0.6) return 'date';
  return 'text';
}

function detectColumns(data) {
  if (!data.length) return [];
  const keys = Object.keys(data[0]);
  return keys.map(name => {
    const vals = data.map(r => r[name]);
    return { name, type: detectType(vals) };
  });
}

// ── File handling ──────────────────────────────────────────────────────────
function showLoading() { ui.loading.classList.remove('hidden'); }
function hideLoading() { ui.loading.classList.add('hidden'); }

// Core loader — accepts ArrayBuffer + display name
// silent=true → preserve current sheet/filters/search/sort (used by live-sync)
function loadFromBuffer(buffer, fileName, silent = false) {
  const data = new Uint8Array(buffer);
  const wb = XLSX.read(data, { type: 'array', cellDates: true });
  state.workbook = wb;
  ui.fileNameDisplay.textContent = fileName;
  buildSheetTabs(wb);
  loadPQASection(wb);

  if (silent) {
    // Keep everything the user had — just refresh data
    const sheet = (state.sheetName && wb.SheetNames.includes(state.sheetName))
      ? state.sheetName : wb.SheetNames[0];
    loadSheet(sheet);
  } else {
    state.filters = {};
    state.searchQuery = '';
    state.page = 0;
    state.sortCol = null;
    state.sortDir = 'asc';
    if (ui.tableSearch) ui.tableSearch.value = '';
    loadSheet(wb.SheetNames[0]);
    showDashboard();
  }
}

// Called when user manually picks a file
function handleFile(file) {
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx','xls','csv'].includes(ext)) {
    alert('Unsupported file type. Please upload .xlsx, .xls, or .csv');
    return;
  }
  showLoading();
  const reader = new FileReader();
  reader.onload = e => {
    try {
      loadFromBuffer(e.target.result, file.name);
    } catch (err) {
      console.error(err);
      alert('Error reading file: ' + err.message);
    } finally {
      hideLoading();
    }
  };
  reader.readAsArrayBuffer(file);
}

// Auto-load: tries to fetch 'data.xlsx' from the same folder
// Works when opened via VS Code Live Server (not raw file://)
async function autoLoadDefaultFile() {
  const candidates = ['data/data.xlsx', 'data/data.xls', 'data/data.csv', 'data.xlsx', 'data.xls', 'data.csv'];
  for (const name of candidates) {
    try {
      const res = await fetch(name);
      if (!res.ok) continue;
      const buffer = await res.arrayBuffer();
      loadFromBuffer(buffer, name + ' (auto-loaded)');
      return; // loaded successfully, stop trying
    } catch (_) {
      // fetch blocked (file:// protocol) or file not found — try next
    }
  }
  // No default file found — upload screen stays visible (do nothing)
}

// ── Sheet management ───────────────────────────────────────────────────────
function buildSheetTabs(wb) {
  ui.sheetTabs.innerHTML = '';
  wb.SheetNames.forEach(name => {
    const btn = document.createElement('button');
    btn.className = 'sheet-tab';
    btn.textContent = name;
    btn.onclick = () => {
      state.filters = {};
      state.searchQuery = '';
      state.page = 0;
      ui.tableSearch.value = '';
      loadSheet(name);
    };
    ui.sheetTabs.appendChild(btn);
  });
}

function loadSheet(name) {
  state.sheetName = name;
  document.querySelectorAll('.sheet-tab').forEach(t => {
    t.classList.toggle('active', t.textContent === name);
  });
  const qaLabel = $('qa-sheet-label');
  if (qaLabel) qaLabel.textContent = name;
  const ws = state.workbook.Sheets[name];
  const raw = XLSX.utils.sheet_to_json(ws, { defval: '' });
  state.rawData = raw;
  state.columns = detectColumns(raw);
  applyFilters();
  buildFilters();
  populateChartSelects();
  renderKPIs();
  renderCharts();
  renderTable();
}

// ── Show/hide screens ──────────────────────────────────────────────────────
function showDashboard() {
  ui.uploadScreen.classList.add('hidden');
  ui.dashboard.classList.remove('hidden');
}
function showUpload() {
  stopAutoSync();
  state.fileHandle = null;
  state.lastModified = null;
  ui.dashboard.classList.add('hidden');
  ui.uploadScreen.classList.remove('hidden');
  state.workbook = null;
  state.rawData = [];
  state.filteredData = [];
  Object.values(state.charts).forEach(c => c && c.destroy());
  state.charts = {};
  // Reset PQA
  pqaState.data = [];
  pqaState.columns = [];
  pqaState.sheetName = null;
  const pqaSec = $('pqa-section');
  if (pqaSec) pqaSec.classList.add('hidden');
  const sbList = $('sb-game-list');
  if (sbList) sbList.innerHTML = '<div class="sb-game-empty">Load data to see games</div>';
  const sbCount = $('sb-game-count');
  if (sbCount) sbCount.textContent = '';
  const sbSearch = $('sb-game-search');
  if (sbSearch) sbSearch.value = '';

  // Reset PQA search section
  const pqaSearchSec = $('pqa-search-section');
  if (pqaSearchSec) pqaSearchSec.classList.add('hidden');
  const sbPqaList = $('sb-pqa-game-list');
  if (sbPqaList) sbPqaList.innerHTML = '<div class="sb-game-empty">No PQA data loaded</div>';
  const sbPqaCount = $('sb-pqa-game-count');
  if (sbPqaCount) sbPqaCount.textContent = '';
  const sbPqaSearch = $('sb-pqa-game-search');
  if (sbPqaSearch) sbPqaSearch.value = '';
}

// ── Filters ────────────────────────────────────────────────────────────────
function applyFilters() {
  let data = state.rawData;

  // dropdown filters
  for (const [col, val] of Object.entries(state.filters)) {
    if (val === '' || val === '__all__') continue;
    data = data.filter(row => String(row[col]) === val);
  }

  // search filter
  if (state.searchQuery) {
    const q = state.searchQuery.toLowerCase();
    data = data.filter(row =>
      Object.values(row).some(v => String(v).toLowerCase().includes(q))
    );
  }

  state.filteredData = data;
}

function buildFilters() {
  ui.filterList.innerHTML = '';
  // Show filters for text/date columns only (up to 6)
  const filterCols = state.columns
    .filter(c => c.type !== 'number')
    .slice(0, 6);

  filterCols.forEach(col => {
    const uniqueVals = [...new Set(state.rawData.map(r => String(r[col.name])))].sort();
    if (uniqueVals.length > 100) return; // skip high-cardinality

    const div = document.createElement('div');
    div.className = 'filter-item';
    div.innerHTML = `<label>${col.name}</label>`;

    const sel = document.createElement('select');
    sel.className = 'filter-select';
    sel.innerHTML = `<option value="__all__">All</option>` +
      uniqueVals.map(v => `<option value="${escHtml(v)}">${escHtml(v)}</option>`).join('');
    sel.value = state.filters[col.name] || '__all__';
    sel.onchange = () => {
      state.filters[col.name] = sel.value;
      state.page = 0;
      applyFilters();
      renderKPIs();
      renderCharts();
      renderTable();
    };
    div.appendChild(sel);
    ui.filterList.appendChild(div);
  });
}

function resetFilters() {
  state.filters = {};
  state.searchQuery = '';
  state.page = 0;
  state.sortCol = null;
  state.sortDir = 'asc';
  if (ui.tableSearch) ui.tableSearch.value = '';
  loadSheet(state.sheetName);
}

// ── Field list ─────────────────────────────────────────────────────────────
function buildFieldList() {
  ui.fieldList.innerHTML = '';
  state.columns.forEach(col => {
    const div = document.createElement('div');
    div.className = 'field-item';
    const badge = col.type === 'number' ? 'badge-num' : col.type === 'date' ? 'badge-date' : 'badge-text';
    const label = col.type === 'number' ? 'NUM' : col.type === 'date' ? 'DATE' : 'TEXT';
    div.innerHTML = `<span class="badge ${badge}">${label}</span><span class="field-name" title="${escHtml(col.name)}">${escHtml(col.name)}</span>`;
    ui.fieldList.appendChild(div);
  });
}

// ── Chart select population ────────────────────────────────────────────────
function populateChartSelects() {
  const allCols = state.columns.map(c => c.name);
  const catCols = state.columns.filter(c => c.type !== 'number').map(c => c.name);
  const cols = catCols.length ? catCols : allCols;

  const opts = cols.map(o => `<option value="${escHtml(o)}">${escHtml(o)}</option>`).join('');
  ui.statusCol.innerHTML = opts;
  ui.platformCol.innerHTML = opts;

  // Auto-detect: pick column whose name contains a known keyword
  function detect(keywords) {
    return cols.find(c => keywords.some(k => c.toLowerCase().includes(k))) || cols[0] || '';
  }
  const statusDef   = detect(['status', 'state', 'result', 'pass', 'fail', 'qa']);
  const platformDef = detect(['platform', 'device', 'type', 'sp', 'stb', 'jp']);

  if (statusDef) ui.statusCol.value = statusDef;
  // Platform: pick a different column from status if possible
  const platformAlt = cols.find(c => c !== ui.statusCol.value) || '';
  ui.platformCol.value = (platformDef && platformDef !== ui.statusCol.value)
    ? platformDef : (platformAlt || platformDef);
}

// ── Hero Total Count ────────────────────────────────────────────────────────
function renderKPIs() {
  const total    = state.filteredData.length;
  const rawTotal = state.rawData.length;
  $('hero-total').textContent = total.toLocaleString('en-IN');
  $('hero-sub').textContent   = total < rawTotal
    ? `${total.toLocaleString('en-IN')} filtered  /  ${rawTotal.toLocaleString('en-IN')} total`
    : `${rawTotal.toLocaleString('en-IN')} total records`;

  // Platform counts — reuse aggregate() with the same column as the Platform chart
  const platCol = ui.platformCol ? ui.platformCol.value : '';
  if (platCol) {
    const counts = aggregate(state.filteredData, platCol, '__count__', 'count', 'none');
    const find = tag => {
      const t = tag.toLowerCase();
      const hit = counts.find(e => e.key.toLowerCase().includes(t));
      return hit ? hit.value.toLocaleString('en-IN') : '0';
    };
    $('kpi-sp-val').textContent  = find('sp');
    $('kpi-stb-val').textContent = find('stb');
    $('kpi-jp-val').textContent  = find('jp');
  }
}

// ── Aggregation helpers ────────────────────────────────────────────────────
function aggregate(data, groupCol, valCol, aggFn, sortMode) {
  // sortMode: 'value_desc' | 'label_asc' | 'none'
  const groups = {};
  const insertOrder = [];
  data.forEach(row => {
    const key = String(row[groupCol] ?? '(blank)');
    if (!groups[key]) { groups[key] = []; insertOrder.push(key); }
    const v = valCol === '__count__' ? 1 : rawNumber(row[valCol]);
    if (!isNaN(v)) groups[key].push(v);
    else if (valCol === '__count__') groups[key].push(1);
  });

  let entries = insertOrder.map(k => {
    const vals = groups[k];
    let agg;
    if (aggFn === 'sum' || aggFn === 'count') agg = vals.reduce((a, b) => a + b, 0);
    else if (aggFn === 'avg') agg = vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
    else if (aggFn === 'max') agg = Math.max(...vals);
    else agg = vals.reduce((a, b) => a + b, 0);
    return { key: k, value: agg };
  });

  if (sortMode === 'value_desc') {
    entries.sort((a, b) => b.value - a.value);
  } else if (sortMode === 'label_asc') {
    entries.sort((a, b) => a.key.localeCompare(b.key, undefined, { numeric: true }));
  }
  // 'none' keeps original data order

  return entries;
}

// ── Chart rendering ────────────────────────────────────────────────────────
function destroyChart(key) {
  if (state.charts[key]) { state.charts[key].destroy(); state.charts[key] = null; }
}

function renderCharts() {
  renderStatus();
  renderPlatform();
}

// Status distribution — donut chart
function renderStatus() {
  destroyChart('status');
  const col = ui.statusCol.value;
  if (!col) return;

  const entries = aggregate(state.filteredData, col, '__count__', 'count', 'value_desc').slice(0, 12);
  const ctx = $('status-chart').getContext('2d');
  state.charts.status = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: entries.map(e => e.key),
      datasets: [{
        data: entries.map(e => e.value),
        backgroundColor: entries.map((_, i) => PALETTE[i % PALETTE.length] + 'e0'),
        borderColor: '#fff',
        borderWidth: 2,
        hoverOffset: 10,
      }]
    },
    options: {
      ...CHART_DEFAULTS,
      cutout: '60%',
      plugins: {
        legend: { position: 'right', labels: { font: { size: 11 }, boxWidth: 12, padding: 10 } },
        tooltip: {
          callbacks: {
            label: ctx => {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct   = total ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
              return ` ${ctx.label}: ${ctx.parsed}  (${pct}%)`;
            }
          }
        }
      }
    }
  });
}

// Platform distribution — horizontal bar chart
function renderPlatform() {
  destroyChart('platform');
  const col = ui.platformCol.value;
  if (!col) return;

  const entries = aggregate(state.filteredData, col, '__count__', 'count', 'value_desc').slice(0, 15);
  const ctx = $('platform-chart').getContext('2d');
  state.charts.platform = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: entries.map(e => e.key),
      datasets: [{
        label: 'Games',
        data: entries.map(e => e.value),
        backgroundColor: entries.map((_, i) => PALETTE[i % PALETTE.length] + 'cc'),
        borderColor:     entries.map((_, i) => PALETTE[i % PALETTE.length]),
        borderWidth: 1,
        borderRadius: 5,
      }]
    },
    options: {
      ...CHART_DEFAULTS,
      indexAxis: 'y',
      plugins: {
        ...CHART_DEFAULTS.plugins,
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ` Games: ${ctx.parsed.x}` } }
      },
      scales: {
        y: { ticks: { font: { size: 11 } } },
        x: { beginAtZero: true, ticks: { stepSize: 1, font: { size: 10 } } }
      }
    }
  });
}

// ── Table ──────────────────────────────────────────────────────────────────
function getTableData() {
  let data = [...state.filteredData];

  if (state.sortCol) {
    const col = state.columns.find(c => c.name === state.sortCol);
    const isNum = col && col.type === 'number';
    data.sort((a, b) => {
      const av = isNum ? (rawNumber(a[state.sortCol]) || 0) : String(a[state.sortCol] ?? '');
      const bv = isNum ? (rawNumber(b[state.sortCol]) || 0) : String(b[state.sortCol] ?? '');
      let cmp = isNum ? av - bv : av.localeCompare(bv, undefined, { numeric: true });
      return state.sortDir === 'desc' ? -cmp : cmp;
    });
  }

  const total = data.length;
  const pages = Math.max(1, Math.ceil(total / state.pageSize));
  state.page = Math.min(state.page, pages - 1);
  const start = state.page * state.pageSize;
  const page = data.slice(start, start + state.pageSize);
  return { page, total, pages, start };
}

function renderTable() {
  const cols = state.columns;
  if (!cols.length) return;

  // Header
  ui.tableHead.innerHTML = '<tr>' + cols.map(col => {
    let cls = '';
    if (state.sortCol === col.name) cls = 'sort-' + state.sortDir;
    return `<th class="${cls}" data-col="${escHtml(col.name)}">${escHtml(col.name)}<span class="sort-icon"></span></th>`;
  }).join('') + '</tr>';

  // Sort click
  ui.tableHead.querySelectorAll('th').forEach(th => {
    th.onclick = () => {
      const col = th.dataset.col;
      if (state.sortCol === col) {
        state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
      } else {
        state.sortCol = col;
        state.sortDir = 'asc';
      }
      renderTable();
    };
  });

  const { page, total, pages, start } = getTableData();

  // Body
  ui.tableBody.innerHTML = page.map(row =>
    '<tr class="clickable-row">' + cols.map(col => `<td title="${escHtml(String(row[col.name] ?? ''))}">${escHtml(String(row[col.name] ?? ''))}</td>`).join('') + '</tr>'
  ).join('');

  // Click handler — show detail modal
  ui.tableBody.querySelectorAll('.clickable-row').forEach((tr, idx) => {
    tr.addEventListener('click', () => showGameDetail(page[idx]));
  });

  // Footer
  ui.rowCount.textContent = `${total.toLocaleString('en-IN')} rows`;
  ui.pageInfo.textContent = `Page ${state.page + 1} of ${pages}`;
  ui.btnPrev.disabled = state.page === 0;
  ui.btnNext.disabled = state.page >= pages - 1;

  renderSidebarGameList();
}

// ── Export PDF — visuals on page 1 exactly as on screen, table on page 2 ────
async function exportAllVisuals() {
  if (!state.filteredData.length) { alert('Pehle data load karo.'); return; }

  const loadingMsg = document.querySelector('#loading p');
  const fileName   = (ui.fileNameDisplay.textContent || 'dashboard')
    .replace(/[^a-z0-9]/gi, '_').replace(/_+/g, '_');

  loadingMsg.textContent = 'PDF bana raha hai...';
  showLoading();

  // A4 Landscape (mm)
  const PW = 297, PH = 210, M = 8;
  const UW  = PW - M * 2;          // usable width
  const HDR = 12;                   // header bar height
  const UH  = PH - HDR - M * 1.5;  // usable height (content area)
  const GAP = 5;                    // gap between hero and charts
  const SCALE = 2;

  const dateStr = new Date().toLocaleDateString('en-IN',
    { day: '2-digit', month: 'short', year: 'numeric' });

  // Capture an element: returns { data, mmH } scaled to full usable width
  const snap = async (el, bg = '#f0f2f5') => {
    const c = await html2canvas(el, {
      backgroundColor: bg, scale: SCALE,
      useCORS: true, logging: false, scrollX: 0, scrollY: 0,
    });
    const pw = c.width / SCALE, ph = c.height / SCALE;
    return { data: c.toDataURL('image/jpeg', 0.93), mmH: ph * (UW / pw) };
  };

  try {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    let pg = 1;

    const drawHdr = () => {
      pdf.setFillColor(242, 200, 17);
      pdf.rect(0, 0, PW, HDR, 'F');
      pdf.setFont('helvetica', 'bold'); pdf.setFontSize(11); pdf.setTextColor(26, 26, 46);
      pdf.text('DataViz Pro — Dashboard', M, 8.5);
      pdf.setFont('helvetica', 'normal'); pdf.setFontSize(9); pdf.setTextColor(50, 50, 70);
      pdf.text(`${dateStr}  |  Page ${pg}`, PW - M, 8.5, { align: 'right' });
    };

    // ── PAGE 1: Hero + Charts grid captured as ONE image each ────────────────
    // Capturing .charts-grid as a single element preserves the 2-column
    // side-by-side layout — exactly as it appears on screen.
    const hero   = await snap(document.querySelector('.hero-section'));
    const charts = await snap(document.querySelector('.charts-grid'));

    const totalH = hero.mmH + GAP + charts.mmH;
    // If combined height exceeds the page, scale everything down proportionally
    const fit    = totalH > UH ? UH / totalH : 1;
    const W      = UW * fit;
    const offX   = M + (UW - W) / 2;   // center horizontally when scaled

    // ── PAGE 1: QA Slide ─────────────────────────────────────────────────────
    drawHdr();
    let y = HDR + M;
    pdf.addImage(hero.data,   'JPEG', offX, y, W, hero.mmH   * fit);
    y += hero.mmH * fit + GAP;
    pdf.addImage(charts.data, 'JPEG', offX, y, W, charts.mmH * fit);

    // ── PAGE 2: PQA Slide (only if section is visible) ───────────────────────
    const pqaSection = $('pqa-section');
    if (pqaSection && !pqaSection.classList.contains('hidden') && pqaState.data.length) {
      pdf.addPage();
      pg++;

      const pqaHeroSnap   = await snap(pqaSection.querySelector('.hero-section'));
      const pqaChartsSnap = await snap(pqaSection.querySelector('.charts-grid'));

      const totalH2 = pqaHeroSnap.mmH + GAP + pqaChartsSnap.mmH;
      const fit2    = totalH2 > UH ? UH / totalH2 : 1;
      const W2      = UW * fit2;
      const offX2   = M + (UW - W2) / 2;

      drawHdr();
      let y2 = HDR + M;
      pdf.addImage(pqaHeroSnap.data,   'JPEG', offX2, y2, W2, pqaHeroSnap.mmH   * fit2);
      y2 += pqaHeroSnap.mmH * fit2 + GAP;
      pdf.addImage(pqaChartsSnap.data, 'JPEG', offX2, y2, W2, pqaChartsSnap.mmH * fit2);
    }

    pdf.save(fileName + '_dashboard.pdf');

  } catch (err) {
    console.error(err);
    alert('PDF export fail hua: ' + err.message);
  } finally {
    loadingMsg.textContent = 'Processing file...';
    hideLoading();
  }
}

// ── PQA Slide ──────────────────────────────────────────────────────────────

function loadPQASection(wb) {
  const pqaSheet = wb.SheetNames.find(n => n.toLowerCase().includes('pqa'));
  const pqaSection = $('pqa-section');

  if (!pqaSheet || !pqaSection) {
    if (pqaSection) pqaSection.classList.add('hidden');
    return;
  }

  pqaState.sheetName = pqaSheet;
  pqaState.data    = XLSX.utils.sheet_to_json(wb.Sheets[pqaSheet], { defval: '' });
  pqaState.columns = detectColumns(pqaState.data);

  const sheetLabel = $('pqa-sheet-label');
  if (sheetLabel) sheetLabel.textContent = pqaSheet;

  // Populate column selects
  const catCols = pqaState.columns.filter(c => c.type !== 'number').map(c => c.name);
  const choices  = catCols.length ? catCols : pqaState.columns.map(c => c.name);
  const opts = choices.map(o => `<option value="${escHtml(o)}">${escHtml(o)}</option>`).join('');
  if (ui.pqaStatusCol)   ui.pqaStatusCol.innerHTML   = opts;
  if (ui.pqaPlatformCol) ui.pqaPlatformCol.innerHTML = opts;

  function detect(keywords) {
    return choices.find(c => keywords.some(k => c.toLowerCase().includes(k))) || choices[0] || '';
  }
  const sDef = detect(['status', 'state', 'result', 'pass', 'fail', 'pqa', 'qa']);
  const pDef = detect(['platform', 'device', 'type', 'sp', 'stb', 'jp']);

  if (ui.pqaStatusCol && sDef)   ui.pqaStatusCol.value   = sDef;
  if (ui.pqaPlatformCol) {
    const useP = pDef && pDef !== ui.pqaStatusCol?.value
      ? pDef
      : (choices.find(c => c !== (ui.pqaStatusCol?.value || '')) || pDef);
    if (useP) ui.pqaPlatformCol.value = useP;
  }

  pqaSection.classList.remove('hidden');
  renderPQAKPIs();
  renderPQACharts();

  // Show PQA sidebar search section and populate it
  const pqaSearchSec = $('pqa-search-section');
  if (pqaSearchSec) pqaSearchSec.classList.remove('hidden');
  renderPQASidebarGameList();
}

function renderPQAKPIs() {
  const total = pqaState.data.length;
  $('pqa-hero-total').textContent = total.toLocaleString('en-IN');
  $('pqa-hero-sub').textContent   = `${total.toLocaleString('en-IN')} total records`;

  const platCol = ui.pqaPlatformCol ? ui.pqaPlatformCol.value : '';
  if (platCol) {
    const counts = aggregate(pqaState.data, platCol, '__count__', 'count', 'none');
    const find = tag => {
      const t = tag.toLowerCase();
      const hit = counts.find(e => e.key.toLowerCase().includes(t));
      return hit ? hit.value.toLocaleString('en-IN') : '0';
    };
    $('pqa-kpi-sp-val').textContent  = find('sp');
    $('pqa-kpi-stb-val').textContent = find('stb');
    $('pqa-kpi-jp-val').textContent  = find('jp');
  }
}

function renderPQACharts() {
  renderPQAStatus();
  renderPQAPlatform();
}

function renderPQAStatus() {
  destroyChart('pqaStatus');
  const col = ui.pqaStatusCol ? ui.pqaStatusCol.value : '';
  if (!col || !pqaState.data.length) return;

  const entries = aggregate(pqaState.data, col, '__count__', 'count', 'value_desc').slice(0, 12);
  const ctx = $('pqa-status-chart').getContext('2d');
  state.charts.pqaStatus = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: entries.map(e => e.key),
      datasets: [{
        data: entries.map(e => e.value),
        backgroundColor: entries.map((_, i) => PALETTE[i % PALETTE.length] + 'e0'),
        borderColor: '#fff',
        borderWidth: 2,
        hoverOffset: 10,
      }]
    },
    options: {
      ...CHART_DEFAULTS,
      cutout: '60%',
      plugins: {
        legend: { position: 'right', labels: { font: { size: 11 }, boxWidth: 12, padding: 10 } },
        tooltip: {
          callbacks: {
            label: ctx => {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct   = total ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
              return ` ${ctx.label}: ${ctx.parsed}  (${pct}%)`;
            }
          }
        }
      }
    }
  });
}

function renderPQAPlatform() {
  destroyChart('pqaPlatform');
  const col = ui.pqaPlatformCol ? ui.pqaPlatformCol.value : '';
  if (!col || !pqaState.data.length) return;

  const entries = aggregate(pqaState.data, col, '__count__', 'count', 'value_desc').slice(0, 15);
  const ctx = $('pqa-platform-chart').getContext('2d');
  state.charts.pqaPlatform = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: entries.map(e => e.key),
      datasets: [{
        label: 'PQA Games',
        data: entries.map(e => e.value),
        backgroundColor: entries.map((_, i) => PALETTE[i % PALETTE.length] + 'cc'),
        borderColor:     entries.map((_, i) => PALETTE[i % PALETTE.length]),
        borderWidth: 1,
        borderRadius: 5,
      }]
    },
    options: {
      ...CHART_DEFAULTS,
      indexAxis: 'y',
      plugins: {
        ...CHART_DEFAULTS.plugins,
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ` PQA Games: ${ctx.parsed.x}` } }
      },
      scales: {
        y: { ticks: { font: { size: 11 } } },
        x: { beginAtZero: true, ticks: { stepSize: 1, font: { size: 10 } } }
      }
    }
  });
}

// ── Helper: escape HTML ────────────────────────────────────────────────────
function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── Sidebar Game Search ────────────────────────────────────────────────────

function getGameNameCol() {
  // Match camelCase "GameName", "game name", "game_name", "Game Title", etc.
  return (
    state.columns.find(c => /game.?name/i.test(c.name))  ||
    state.columns.find(c => /game.?title/i.test(c.name)) ||
    state.columns.find(c => /\bname\b/i.test(c.name))    ||
    state.columns.find(c => /\btitle\b/i.test(c.name))   ||
    state.columns.find(c => c.type === 'text')            ||
    state.columns[0]                                       ||
    null
  );
}

function renderSidebarGameList() {
  const listEl  = $('sb-game-list');
  const countEl = $('sb-game-count');
  if (!listEl) return;

  const nameCol = getGameNameCol();
  if (!nameCol || !state.filteredData.length) {
    listEl.innerHTML  = '<div class="sb-game-empty">Load data to see games</div>';
    if (countEl) countEl.textContent = '';
    return;
  }

  const query       = ($('sb-game-search')?.value || '').trim().toLowerCase();
  const platColName = ui.platformCol?.value || '';

  // Show ALL rows — no deduplication, so count matches total QA games
  const items = state.filteredData
    .map(row => ({
      name: String(row[nameCol.name] || '').trim(),
      plat: platColName ? String(row[platColName] || '').trim() : '',
      row,
    }))
    .filter(item => item.name && (!query || item.name.toLowerCase().includes(query)));

  if (countEl) countEl.textContent = items.length + ' game' + (items.length !== 1 ? 's' : '');

  if (!items.length) {
    listEl.innerHTML = '<div class="sb-game-empty">No games found</div>';
    return;
  }

  listEl.innerHTML = items
    .map((item, i) => {
      const sub = item.plat
        ? `<span class="sb-game-sub">${escHtml(item.plat)}</span>`
        : '';
      return `<button class="sb-game-item" data-idx="${i}" title="${escHtml(item.name)}">
        <span class="sb-game-name">${escHtml(item.name)}</span>${sub}
      </button>`;
    })
    .join('');

  listEl.querySelectorAll('.sb-game-item').forEach((btn, i) => {
    btn.addEventListener('click', () => showGameDetail(items[i].row));
  });
}

// ── Game Detail Modal ──────────────────────────────────────────────────────

function getStatusBadgeClass(value) {
  const v = String(value).toLowerCase().trim();
  const pass    = ['pass', 'passed', 'done', 'completed', 'approved', 'yes', 'ok', 'cleared', 'fixed'];
  const fail    = ['fail', 'failed', 'error', 'rejected', 'no', 'crash', 'blocked', 'abort'];
  const pending = ['pending', 'in progress', 'wip', 'hold', 'review', 'open', 'todo', 'inprogress', 'ongoing'];
  if (pass.some(k => v.includes(k)))    return 'badge-status-pass';
  if (fail.some(k => v.includes(k)))    return 'badge-status-fail';
  if (pending.some(k => v.includes(k))) return 'badge-status-pending';
  return 'badge-status-default';
}

// cols   — column definitions to use (state.columns for QA, pqaState.columns for PQA)
// isPQA  — true → show purple PQA header accent
function showGameDetail(row, cols, isPQA) {
  cols = cols || state.columns;
  if (!cols.length) return;

  const titleCol =
    cols.find(c => /game.?name/i.test(c.name))  ||
    cols.find(c => /game.?title/i.test(c.name)) ||
    cols.find(c => /\bname\b/i.test(c.name))    ||
    cols.find(c => /\btitle\b/i.test(c.name))   ||
    cols.find(c => c.type === 'text')            ||
    cols[0];

  const title = String(row[titleCol.name] || '—');
  $('game-modal-title').textContent = title;

  // Header colour: purple for PQA, default dark for QA
  const modalHeader = document.querySelector('.game-modal-header');
  if (modalHeader) {
    modalHeader.style.background = isPQA ? '#2d1b4e' : '';
  }
  const modalLabel = document.querySelector('.game-modal-label');
  if (modalLabel) {
    modalLabel.textContent = isPQA ? 'PQA Game Details' : 'Game Details';
    modalLabel.style.color = isPQA ? '#a855f7' : '';
  }
  const modalIcon = document.querySelector('.game-modal-icon');
  if (modalIcon) {
    modalIcon.style.background = isPQA ? 'rgba(168,85,247,0.15)' : '';
    modalIcon.style.color      = isPQA ? '#a855f7' : '';
  }

  const statusKeywords = ['status', 'state', 'result', 'pass', 'fail', 'qa', 'pqa', 'approval', 'review'];

  const body = $('game-modal-body');
  body.innerHTML = cols.map(col => {
    const raw = row[col.name];
    const val = (raw == null || raw === '') ? '' : String(raw);
    const isTitle     = col.name === titleCol.name;
    const isStatusCol = statusKeywords.some(k => col.name.toLowerCase().includes(k));

    let valueHtml;
    if (!val) {
      valueHtml = `<span class="game-detail-field-value val-empty">—</span>`;
    } else if (isStatusCol) {
      valueHtml = `<span class="game-detail-badge ${getStatusBadgeClass(val)}">${escHtml(val)}</span>`;
    } else {
      valueHtml = `<span class="game-detail-field-value">${escHtml(val)}</span>`;
    }

    return `<div class="game-detail-field${isTitle ? ' game-detail-field-full' : ''}">
      <div class="game-detail-field-label">${escHtml(col.name)}</div>
      <div>${valueHtml}</div>
    </div>`;
  }).join('');

  $('game-detail-modal').classList.remove('hidden');
  document.body.style.overflow = 'hidden';
}

// ── PQA Sidebar Game Search ────────────────────────────────────────────────

function getPQAGameNameCol() {
  const cols = pqaState.columns;
  return (
    cols.find(c => /game.?name/i.test(c.name))  ||
    cols.find(c => /game.?title/i.test(c.name)) ||
    cols.find(c => /\bname\b/i.test(c.name))    ||
    cols.find(c => /\btitle\b/i.test(c.name))   ||
    cols.find(c => c.type === 'text')            ||
    cols[0]                                       ||
    null
  );
}

function renderPQASidebarGameList() {
  const listEl  = $('sb-pqa-game-list');
  const countEl = $('sb-pqa-game-count');
  if (!listEl) return;

  const nameCol = getPQAGameNameCol();
  if (!nameCol || !pqaState.data.length) {
    listEl.innerHTML = '<div class="sb-game-empty">No PQA data loaded</div>';
    if (countEl) countEl.textContent = '';
    return;
  }

  const query       = ($('sb-pqa-game-search')?.value || '').trim().toLowerCase();
  const platColName = ui.pqaPlatformCol?.value || '';

  // Show ALL rows — no deduplication, so count matches total PQA games
  const items = pqaState.data
    .map(row => ({
      name: String(row[nameCol.name] || '').trim(),
      plat: platColName ? String(row[platColName] || '').trim() : '',
      row,
    }))
    .filter(item => item.name && (!query || item.name.toLowerCase().includes(query)));

  if (countEl) countEl.textContent = items.length + ' game' + (items.length !== 1 ? 's' : '');

  if (!items.length) {
    listEl.innerHTML = '<div class="sb-game-empty">No PQA games found</div>';
    return;
  }

  listEl.innerHTML = items
    .map((item, i) => {
      const sub = item.plat
        ? `<span class="sb-game-sub">${escHtml(item.plat)}</span>`
        : '';
      return `<button class="sb-game-item sb-pqa-game-item" data-idx="${i}" title="${escHtml(item.name)}">
        <span class="sb-game-name">${escHtml(item.name)}</span>${sub}
      </button>`;
    })
    .join('');

  listEl.querySelectorAll('.sb-pqa-game-item').forEach((btn, i) => {
    btn.addEventListener('click', () => showGameDetail(items[i].row, pqaState.columns, true));
  });
}

function hideGameDetail() {
  $('game-detail-modal').classList.add('hidden');
  document.body.style.overflow = '';
}

// ── Live-sync (File System Access API) ─────────────────────────────────────

// Open file picker — gets a persistent FileSystemFileHandle for live-sync
// Falls back to regular <input> on browsers without the API (Firefox)
async function openFilePicker() {
  if (!window.showOpenFilePicker) {
    ui.fileInput.click();   // fallback — no live-sync, one-time load
    return;
  }
  try {
    const [handle] = await window.showOpenFilePicker({
      types: [{
        description: 'Spreadsheet',
        accept: {
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
          'application/vnd.ms-excel': ['.xls'],
          'text/csv': ['.csv'],
        },
      }],
      multiple: false,
    });
    await loadFromHandle(handle);
  } catch (err) {
    if (err.name !== 'AbortError') console.error(err);
  }
}

// Load file from a FileSystemFileHandle (first load or re-read)
async function loadFromHandle(handle, silent = false) {
  try {
    const file = await handle.getFile();
    const buffer = await file.arrayBuffer();
    state.fileHandle = handle;
    state.lastModified = file.lastModified;
    loadFromBuffer(buffer, file.name, silent);
    if (!silent) startAutoSync();
  } catch (err) {
    console.error('Error reading file handle:', err);
  }
}

// Start polling — check every 2.5 s if the file changed on disk
function startAutoSync() {
  stopAutoSync();
  state.syncInterval = setInterval(checkForUpdates, 2500);
  updateSyncUI(true);
}

function stopAutoSync() {
  if (state.syncInterval) {
    clearInterval(state.syncInterval);
    state.syncInterval = null;
  }
  updateSyncUI(false);
}

// Poll: compare lastModified; silent-reload on change
async function checkForUpdates() {
  if (!state.fileHandle) return;
  try {
    const file = await state.fileHandle.getFile();
    if (file.lastModified !== state.lastModified) {
      state.lastModified = file.lastModified;
      const buffer = await file.arrayBuffer();
      loadFromBuffer(buffer, file.name, true);   // silent — keep filters
      flashSyncBadge();
    }
  } catch (_) {
    // File may be locked while Excel is saving — skip this cycle
  }
}

// UI helpers for the sync badge in the header
function updateSyncUI(active) {
  const badge = $('sync-status');
  if (!badge) return;
  if (active) {
    badge.classList.remove('hidden');
    $('sync-label').textContent = 'Live sync';
    badge.classList.remove('flash');
  } else {
    badge.classList.add('hidden');
  }
}

function flashSyncBadge() {
  const badge = $('sync-status');
  if (!badge) return;
  badge.classList.add('flash');
  $('sync-label').textContent = 'Updated!';
  setTimeout(() => {
    badge.classList.remove('flash');
    $('sync-label').textContent = 'Live sync';
  }, 1800);
}

// ── Event wiring ───────────────────────────────────────────────────────────
function init() {
  // Sidebar toggle (mobile)
  const btnSidebarToggle = $('btn-sidebar-toggle');
  const sidebarEl        = $('sidebar');
  const sidebarOverlay   = $('sidebar-overlay');

  function openSidebar()  { sidebarEl.classList.add('open');    sidebarOverlay.classList.add('active'); }
  function closeSidebar() { sidebarEl.classList.remove('open'); sidebarOverlay.classList.remove('active'); }

  if (btnSidebarToggle) btnSidebarToggle.addEventListener('click', () => {
    sidebarEl.classList.contains('open') ? closeSidebar() : openSidebar();
  });
  if (sidebarOverlay) sidebarOverlay.addEventListener('click', closeSidebar);
  // Close sidebar on filter change (mobile UX)
  if (sidebarEl) sidebarEl.addEventListener('change', () => {
    if (window.innerWidth <= 640) closeSidebar();
  });
  // Drag & drop — try to grab a FileSystemFileHandle for live-sync
  ui.dropZone.addEventListener('dragover', e => { e.preventDefault(); ui.dropZone.classList.add('drag-over'); });
  ui.dropZone.addEventListener('dragleave', () => ui.dropZone.classList.remove('drag-over'));
  ui.dropZone.addEventListener('drop', async e => {
    e.preventDefault();
    ui.dropZone.classList.remove('drag-over');
    if (window.showOpenFilePicker && e.dataTransfer.items?.[0]?.getAsFileSystemHandle) {
      try {
        const handle = await e.dataTransfer.items[0].getAsFileSystemHandle();
        if (handle?.kind === 'file') { await loadFromHandle(handle); return; }
      } catch (_) {}
    }
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });

  // Browse Files button (upload screen)
  $('btn-browse').addEventListener('click', e => { e.stopPropagation(); openFilePicker(); });
  ui.dropZone.addEventListener('click', openFilePicker);

  // File input fallback (used when showOpenFilePicker is unavailable)
  ui.fileInput.addEventListener('change', e => {
    const file = e.target.files[0];
    if (file) handleFile(file);
    e.target.value = '';
  });


  // Game Detail Modal close
  $('game-modal-close').addEventListener('click', hideGameDetail);
  $('game-detail-modal').addEventListener('click', e => {
    if (e.target === $('game-detail-modal')) hideGameDetail();
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') hideGameDetail();
  });

  // Buttons
  ui.btnNewFile.addEventListener('click', showUpload);
  ui.btnExport.addEventListener('click', exportAllVisuals);
  ui.btnReset.addEventListener('click', resetFilters);

  // Pagination
  ui.btnPrev.addEventListener('click', () => { if (state.page > 0) { state.page--; renderTable(); } });
  ui.btnNext.addEventListener('click', () => {
    const pages = Math.ceil(state.filteredData.length / state.pageSize);
    if (state.page < pages - 1) { state.page++; renderTable(); }
  });

  // Search
  let searchTimer;
  ui.tableSearch.addEventListener('input', e => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      state.searchQuery = e.target.value.trim();
      state.page = 0;
      applyFilters();
      renderKPIs();
      renderCharts();
      renderTable();
    }, 280);
  });

  // Sidebar QA game name search
  const sbGameSearch = $('sb-game-search');
  if (sbGameSearch) sbGameSearch.addEventListener('input', renderSidebarGameList);

  // Sidebar PQA game name search
  const sbPqaGameSearch = $('sb-pqa-game-search');
  if (sbPqaGameSearch) sbPqaGameSearch.addEventListener('input', renderPQASidebarGameList);

  // Chart control changes
  ui.statusCol.addEventListener('change', renderStatus);
  ui.platformCol.addEventListener('change', renderPlatform);

  // PQA chart control changes
  if (ui.pqaStatusCol)   ui.pqaStatusCol.addEventListener('change',   renderPQAStatus);
  if (ui.pqaPlatformCol) ui.pqaPlatformCol.addEventListener('change', () => { renderPQAKPIs(); renderPQAPlatform(); });
}

// Boot
init();
// Show hint on file:// protocol (Live Server not running)
if (location.protocol === 'file:') {
  const hint = document.getElementById('auto-load-hint');
  if (hint) hint.style.display = 'block';
} else {
  autoLoadDefaultFile();
}
