/* ========================================================
   DataViz Pro — app.js
   Pure JS: SheetJS + Chart.js, no build tools required
   ======================================================== */

'use strict';

// ── Microsoft Graph / SharePoint config ────────────────────────────────────
// PASTE YOUR AZURE AD CLIENT ID BELOW (from portal.azure.com)
const GRAPH_CONFIG = {
  clientId: 'YOUR_CLIENT_ID_HERE',          // ← replace with your Azure AD App Client ID
  tenantId: 'common',                        // use 'common' for any org account
  redirectUri: window.location.origin + window.location.pathname,
  scopes: ['User.Read', 'Files.Read'],
  targetFile: 'dataforautomation.xlsx',      // file name to search in OneDrive
};

let msalInstance = null;

function initMSAL() {
  if (!window.msal) return; // MSAL library not loaded
  if (GRAPH_CONFIG.clientId === 'YOUR_CLIENT_ID_HERE') return; // not configured yet
  try {
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: GRAPH_CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${GRAPH_CONFIG.tenantId}`,
        redirectUri: GRAPH_CONFIG.redirectUri,
      },
      cache: { cacheLocation: 'sessionStorage' },
    });
    msalInstance.handleRedirectPromise().then(resp => {
      if (resp) onMSALLogin(resp.account);
      else {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length) onMSALLogin(accounts[0]);
      }
    }).catch(err => console.error('MSAL redirect error:', err));
  } catch (e) {
    console.error('MSAL init error:', e);
  }
}

function onMSALLogin(account) {
  const signinArea  = document.getElementById('sp-signin-area');
  const signedArea  = document.getElementById('sp-signed-area');
  const nameEl      = document.getElementById('sp-user-name');
  if (signinArea) signinArea.classList.add('hidden');
  if (signedArea) signedArea.classList.remove('hidden');
  if (nameEl) nameEl.textContent = account.name || account.username;
}

async function msSignIn() {
  if (!msalInstance) {
    alert('Microsoft sign-in is not configured yet.\n\nPaste your Azure AD Client ID into app.js (GRAPH_CONFIG.clientId).\nSee console for instructions.');
    return;
  }
  try {
    const resp = await msalInstance.loginPopup({ scopes: GRAPH_CONFIG.scopes });
    onMSALLogin(resp.account);
  } catch (err) {
    if (err.errorCode !== 'user_cancelled') alert('Sign-in failed: ' + err.message);
  }
}

function msSignOut() {
  if (!msalInstance) return;
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length) msalInstance.logoutPopup({ account: accounts[0] });
  const signinArea = document.getElementById('sp-signin-area');
  const signedArea = document.getElementById('sp-signed-area');
  if (signinArea) signinArea.classList.remove('hidden');
  if (signedArea) signedArea.classList.add('hidden');
}

async function getGraphToken() {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) throw new Error('Not signed in');
  const resp = await msalInstance.acquireTokenSilent({
    scopes: GRAPH_CONFIG.scopes,
    account: accounts[0],
  });
  return resp.accessToken;
}

async function loadFromSharePoint() {
  const statusEl = document.getElementById('sp-load-status');
  function setStatus(msg, isErr = false) {
    if (!statusEl) return;
    statusEl.textContent = msg;
    statusEl.style.color = isErr ? '#ef4444' : '#10b981';
  }

  if (!msalInstance) { alert('MSAL not initialised.'); return; }
  showLoading();
  setStatus('Connecting to OneDrive…');
  try {
    const token = await getGraphToken();
    const headers = { Authorization: `Bearer ${token}` };

    // Search for the file by name in the user's OneDrive
    setStatus('Searching for ' + GRAPH_CONFIG.targetFile + '…');
    const searchRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/root/search(q='${encodeURIComponent(GRAPH_CONFIG.targetFile)}')`,
      { headers }
    );
    if (!searchRes.ok) throw new Error(`Graph search failed: ${searchRes.status}`);
    const searchJson = await searchRes.json();
    const match = (searchJson.value || []).find(
      f => f.name.toLowerCase() === GRAPH_CONFIG.targetFile.toLowerCase()
    );
    if (!match) throw new Error(`File "${GRAPH_CONFIG.targetFile}" not found in your OneDrive.`);

    // Download the file content
    setStatus('Downloading file…');
    const dlRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/items/${match.id}/content`,
      { headers }
    );
    if (!dlRes.ok) throw new Error(`Download failed: ${dlRes.status}`);
    const buffer = await dlRes.arrayBuffer();
    setStatus('Loaded successfully!');
    loadFromBuffer(buffer, match.name);
  } catch (err) {
    console.error('SharePoint load error:', err);
    setStatus('Error: ' + err.message, true);
    hideLoading();
  }
}

// ── Load from pasted URL (no sign-in required) ─────────────────────────────
function toDownloadUrl(url) {
  // SharePoint / OneDrive share links → add download=1 to force binary download
  if (
    (url.includes('sharepoint.com') || url.includes('1drv.ms') || url.includes('onedrive.live.com')) &&
    !url.includes('download=1')
  ) {
    return url.includes('?') ? url + '&download=1' : url + '?download=1';
  }
  return url;
}

function isExcelBuffer(buffer) {
  // XLSX/ZIP magic bytes: PK (0x50 0x4B)
  const b = new Uint8Array(buffer.slice(0, 4));
  return b[0] === 0x50 && b[1] === 0x4B;
}

async function loadFromUrl() {
  const input    = document.getElementById('url-paste-input');
  const statusEl = document.getElementById('url-load-status');
  const url      = (input ? input.value : '').trim();

  function setStatus(html, isErr = false) {
    if (!statusEl) return;
    statusEl.innerHTML = html;
    statusEl.style.color = isErr ? '#ef4444' : '#10b981';
  }

  function showDownloadFallback(rawUrl) {
    setStatus(
      `<b>SharePoint ne fetch block kar diya.</b><br><br>` +
      `<b style="color:#f59e0b">Fix karo — 2 steps:</b><br>` +
      `&nbsp;1. <a href="${rawUrl}" target="_blank" rel="noopener"
            style="color:#3b82f6;text-decoration:underline;font-weight:600;">
            Yahan click karo — file download hogi
          </a><br>` +
      `&nbsp;2. Us file ko <b>drag &amp; drop</b> karo ya <b>Browse</b> se upload karo`,
      true
    );
  }

  if (!url) { setStatus('Pehle URL paste karo.', true); return; }

  showLoading();
  const downloadUrl = toDownloadUrl(url);
  const fileName = url.split('/').pop().split('?')[0] || 'loaded-file.xlsx';

  async function tryFetch(label, fetchFn) {
    setStatus(`Trying ${label}…`);
    try {
      const res = await fetchFn();
      if (!res.ok) return false;
      const buffer = await res.arrayBuffer();
      if (!isExcelBuffer(buffer)) return 'not_excel';
      setStatus(`Loaded via ${label}!`);
      loadFromBuffer(buffer, fileName);
      return true;
    } catch (_) { return false; }
  }

  // Attempt 1 — direct (with browser session cookies, works if already logged in to SP)
  const r1 = await tryFetch('Direct', () => fetch(downloadUrl, { credentials: 'include' }));
  if (r1 === true) return;
  if (r1 === 'not_excel') { hideLoading(); showDownloadFallback(url); return; }

  // Attempt 2 — corsproxy.io (most reliable free CORS proxy right now)
  const r2 = await tryFetch('corsproxy.io', () => fetch('https://corsproxy.io/?' + encodeURIComponent(downloadUrl)));
  if (r2 === true) return;
  if (r2 === 'not_excel') { hideLoading(); showDownloadFallback(url); return; }

  // Attempt 3 — api.allorigins.win
  const r3 = await tryFetch('allorigins', () => fetch('https://api.allorigins.win/raw?url=' + encodeURIComponent(downloadUrl)));
  if (r3 === true) return;
  if (r3 === 'not_excel') { hideLoading(); showDownloadFallback(url); return; }

  // Attempt 4 — thingproxy
  const r4 = await tryFetch('thingproxy', () => fetch('https://thingproxy.freeboard.io/fetch/' + downloadUrl));
  if (r4 === true) return;

  // All attempts failed — show fallback
  hideLoading();
  showDownloadFallback(url);
}

// ── State ──────────────────────────────────────────────────────────────────
const pqaState = {
  data: [],
  columns: [],
  sheetName: null,
};

const liveState = {
  data: [],
  columns: [],
  sheetName: null,
};

const premiumState = {
  data: [],
  columns: [],
  sheetName: null,
};

const gamesnacksState = {
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
  pqaStatusCol:    $('pqa-status-col'),
  pqaPlatformCol:  $('pqa-platform-col'),
  livePlatformCol: $('live-platform-col'),
  liveVCountCol:   $('live-vcount-col'),
  liveDevCol:      $('live-dev-col'),
  liveDayCol:      $('live-day-col'),
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
  loadLiveSection(wb);
  loadPremiumSection(wb);
  loadGameSnacksSection(wb);

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
  const candidates = [
    'data/data.xlsx', 'data/data.xls', 'data/data.csv',
    'data.xlsx', 'data.xls', 'data.csv',
    'data/dataforautomation.xlsx', 'dataforautomation.xlsx',
  ];
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

// Load file from a user-supplied URL (Method 2)
async function loadFromURL() {
  const urlInput = $('url-input');
  const url = urlInput ? urlInput.value.trim() : '';
  if (!url) { alert('Please paste a URL first.'); return; }

  showLoading();
  try {
    const res = await fetch(url, { credentials: 'include' });
    if (!res.ok) throw new Error(`Server returned ${res.status} ${res.statusText}`);
    const buffer = await res.arrayBuffer();
    // Derive a clean filename from the URL
    const filename = decodeURIComponent(url.split('/').pop().split('?')[0]) || 'file.xlsx';
    loadFromBuffer(buffer, filename);
    if (urlInput) urlInput.value = '';
  } catch (err) {
    console.error('URL load failed:', err);
    alert(
      'Could not load file from that URL.\n\n' +
      'Common reasons:\n' +
      '• The link is not a direct download link (SharePoint viewer URLs won\'t work)\n' +
      '• CORS policy blocked the request\n' +
      '• You\'re not logged in to the file\'s host\n\n' +
      'For SharePoint / OneDrive:\n' +
      '  1. Open the file in SharePoint\n' +
      '  2. Click "Download" — copy the download URL from the browser address bar\n' +
      '  3. Paste that URL here\n\n' +
      'Error: ' + err.message
    );
  } finally {
    hideLoading();
  }
}

// ── Sheet management ───────────────────────────────────────────────────────
function getSectionTarget(sheetName) {
  const n = sheetName.toLowerCase();
  if (/pqa/.test(n))           return 'pqa-section';
  if (/live/.test(n))          return 'live-section';
  if (/premium/.test(n))       return 'premium-section';
  if (/game\s*snack/.test(n))  return 'gamesnacks-section';
  return null; // QA / main sheet
}

function scrollToSection(el) {
  const content = $('content');
  if (!content || !el) return;
  const top = el.getBoundingClientRect().top - content.getBoundingClientRect().top + content.scrollTop - 10;
  content.scrollTo({ top, behavior: 'smooth' });
}

function buildSheetTabs(wb) {
  ui.sheetTabs.innerHTML = '';
  wb.SheetNames.forEach(name => {
    const sectionId = getSectionTarget(name);
    const btn = document.createElement('button');
    btn.className = 'sheet-tab';
    btn.textContent = name;

    if (sectionId) {
      // Dedicated section — scroll to it, don't reload QA data
      btn.onclick = () => {
        const el = $(sectionId);
        if (!el || el.classList.contains('hidden')) return;
        scrollToSection(el);
        document.querySelectorAll('.sheet-tab').forEach(t => t.classList.remove('active'));
        btn.classList.add('active');
      };
    } else {
      // QA / main sheet — load data
      btn.onclick = () => {
        state.filters = {};
        state.searchQuery = '';
        state.page = 0;
        ui.tableSearch.value = '';
        loadSheet(name);
        const content = $('content');
        if (content) content.scrollTo({ top: 0, behavior: 'smooth' });
      };
    }
    ui.sheetTabs.appendChild(btn);
  });
}

function initScrollActiveTabs() {
  const content = $('content');
  if (!content) return;
  const sections = [
    { id: 'gamesnacks-section', pattern: /game\s*snack/i },
    { id: 'premium-section',   pattern: /premium/i },
    { id: 'live-section',      pattern: /live/i },
    { id: 'pqa-section',       pattern: /pqa/i },
  ];
  content.addEventListener('scroll', () => {
    const contentTop = content.getBoundingClientRect().top;
    let matchPattern = null;
    for (const s of sections) {
      const el = $(s.id);
      if (!el || el.classList.contains('hidden')) continue;
      if (el.getBoundingClientRect().top - contentTop <= 80) { matchPattern = s.pattern; break; }
    }
    document.querySelectorAll('.sheet-tab').forEach(tab => {
      const isActive = matchPattern
        ? matchPattern.test(tab.textContent)
        : tab.textContent === state.sheetName;
      tab.classList.toggle('active', isActive);
    });
  }, { passive: true });
}

// Map sheet names to display metadata
function getSheetMeta(name) {
  const n = name.toLowerCase();
  if (/pqa/.test(n))                        return { title: 'PQA Games — 2025 / 2026',     sub: 'Pre-QA testing overview',  hero: 'Total PQA Games'    };
  if (/live/.test(n))                       return { title: 'Live Games',    sub: 'Live games overview', hero: 'Total Live Games' };
  if (/premium/.test(n))                    return { title: 'Premium Games', sub: 'Premium games overview',   hero: 'Total Premium Games' };
  if (/game\s*snack/.test(n))               return { title: 'GameSnacks ',    sub: 'GameSnacks overview',      hero: 'Total GameSnacks'   };
  return                                           { title: 'QA Games — 2025 / 2026',      sub: 'QA testing overview',      hero: 'Total QA Games'     };
}

function loadSheet(name) {
  state.sheetName = name;
  document.querySelectorAll('.sheet-tab').forEach(t => {
    t.classList.toggle('active', t.textContent === name);
  });

  // Update slide header text dynamically
  const meta = getSheetMeta(name);
  const titleEl = $('qa-slide-title');
  const subEl   = $('qa-slide-sub-text');
  const heroLbl = $('hero-card-label');
  if (titleEl) titleEl.textContent = meta.title;
  if (subEl)   subEl.textContent   = meta.sub;
  if (heroLbl) heroLbl.textContent = meta.hero;

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

  // Reset Live
  liveState.data = [];
  liveState.columns = [];
  liveState.sheetName = null;
  const liveSec = $('live-section');
  if (liveSec) liveSec.classList.add('hidden');

  // Reset Premium
  premiumState.data = [];
  premiumState.columns = [];
  premiumState.sheetName = null;
  const premiumSec = $('premium-section');
  if (premiumSec) premiumSec.classList.add('hidden');

  // Reset GameSnacks
  gamesnacksState.data = [];
  gamesnacksState.columns = [];
  gamesnacksState.sheetName = null;
  const gsSec = $('gamesnacks-section');
  if (gsSec) gsSec.classList.add('hidden');
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
  // Skip game-name and partner columns — those have dedicated UI controls
  const skipPattern = /game.?name|game.?title|\bname\b|\btitle\b|partner/i;
  const filterCols = state.columns
    .filter(c => c.type !== 'number' && !skipPattern.test(c.name))
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
const CHART_SKIP = /game.?name|game.?title|pqa.?game|\bname\b|\btitle\b|partner/i;

function populateChartSelects() {
  const allCols = state.columns.map(c => c.name);
  const catCols = state.columns.filter(c => c.type !== 'number').map(c => c.name);
  const cols = (catCols.length ? catCols : allCols).filter(c => !CHART_SKIP.test(c));

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

  const platCol   = ui.platformCol ? ui.platformCol.value : '';
  const statusCol = ui.statusCol   ? ui.statusCol.value   : '';
  if (platCol) {
    const counts = aggregate(state.filteredData, platCol, '__count__', 'count', 'none');
    const find = (...tags) => {
      const total = counts
        .filter(e => tags.some(t => e.key.toLowerCase().includes(t.toLowerCase())))
        .reduce((sum, e) => sum + e.value, 0);
      return total ? total.toLocaleString('en-IN') : '0';
    };
    $('kpi-sp-val').textContent  = find('sp', 'mobile', 'smartphone');
    $('kpi-stb-val').textContent = find('stb');
    $('kpi-jp-val').textContent  = find('jp', 'jiophone', 'jio phone', 'candy');

    // Status breakdown per platform
    if (statusCol) {
      const platforms = {
        sp:  ['sp', 'mobile', 'smartphone'],
        stb: ['stb'],
        jp:  ['jp', 'jiophone', 'jio phone', 'candy'],
      };
      const statuses = {
        live:   ['live'],
        r2g:    ['r2g', 'ready to go'],
        rework: ['rework', 're-work'],
        hold:   ['hold'],
      };
      Object.entries(platforms).forEach(([platKey, platTags]) => {
        const platRows = state.filteredData.filter(row =>
          platTags.some(t => String(row[platCol] || '').toLowerCase().includes(t.toLowerCase()))
        );
        Object.entries(statuses).forEach(([statusKey, statusTags]) => {
          const count = platRows.filter(row =>
            statusTags.some(t => String(row[statusCol] || '').toLowerCase().includes(t.toLowerCase()))
          ).length;
          const el = $(`kpi-${platKey}-${statusKey}`);
          if (el) el.textContent = count.toLocaleString('en-IN');
        });
      });
    }
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
        legend: {
          position: 'right',
          labels: {
            font: { size: 11 },
            boxWidth: 12,
            padding: 10,
            generateLabels: chart => {
              const data = chart.data;
              return data.labels.map((label, i) => ({
                text: `${label}  (${data.datasets[0].data[i]})`,
                fillStyle: data.datasets[0].backgroundColor[i],
                strokeStyle: '#fff',
                lineWidth: 2,
                hidden: false,
                index: i
              }));
            }
          }
        },
        tooltip: {
          callbacks: {
            label: ctx => {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct   = total ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
              return ` ${ctx.label}: ${ctx.parsed}  (${pct}%)`;
            }
          }
        },
        datalabels: {
          display: ctx => ctx.dataset.data[ctx.dataIndex] > 0,
          color: '#fff',
          font: { size: 12, weight: 'bold' },
          formatter: (value) => value,
          textStrokeColor: 'rgba(0,0,0,0.4)',
          textStrokeWidth: 2,
        }
      }
    },
    plugins: [ChartDataLabels]
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
      layout: { padding: { right: 36 } },
      plugins: {
        ...CHART_DEFAULTS.plugins,
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ` Games: ${ctx.parsed.x}` } },
        datalabels: {
          anchor: 'end',
          align: 'right',
          color: '#94a3b8',
          font: { size: 11, weight: 'bold' },
          formatter: v => v,
        }
      },
      scales: {
        y: { ticks: { font: { size: 11 } } },
        x: { beginAtZero: true, ticks: { stepSize: 1, font: { size: 10 } } }
      }
    },
    plugins: [ChartDataLabels]
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

  loadingMsg.textContent = 'PDF generating...';
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
      pdf.text('Dashboard', M, 8.5);
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

    // ── PAGE 3: Live Slide (only if section is visible) ───────────────────────
    const liveSection = $('live-section');
    if (liveSection && !liveSection.classList.contains('hidden') && liveState.data.length) {
      pdf.addPage();
      pg++;

      const liveHeroSnap = await snap(liveSection.querySelector('.hero-section'));
      const liveGridSnap = await snap(liveSection.querySelector('.charts-grid'));

      const totalH3 = liveHeroSnap.mmH + GAP + liveGridSnap.mmH;
      const fit3    = totalH3 > UH ? UH / totalH3 : 1;
      const W3      = UW * fit3;
      const offX3   = M + (UW - W3) / 2;

      drawHdr();
      let y3 = HDR + M;
      pdf.addImage(liveHeroSnap.data, 'JPEG', offX3, y3, W3, liveHeroSnap.mmH * fit3);
      y3 += liveHeroSnap.mmH * fit3 + GAP;
      pdf.addImage(liveGridSnap.data, 'JPEG', offX3, y3, W3, liveGridSnap.mmH * fit3);
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
  const choices  = (catCols.length ? catCols : pqaState.columns.map(c => c.name)).filter(c => !CHART_SKIP.test(c));
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
        },
        datalabels: {
          display: ctx => ctx.dataset.data[ctx.dataIndex] > 0,
          color: '#fff',
          font: { size: 12, weight: 'bold' },
          formatter: v => v,
          textStrokeColor: 'rgba(0,0,0,0.4)',
          textStrokeWidth: 2,
        }
      }
    },
    plugins: [ChartDataLabels]
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
      layout: { padding: { right: 36 } },
      plugins: {
        ...CHART_DEFAULTS.plugins,
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ` PQA Games: ${ctx.parsed.x}` } },
        datalabels: {
          anchor: 'end',
          align: 'right',
          color: '#94a3b8',
          font: { size: 11, weight: 'bold' },
          formatter: v => v,
        }
      },
      scales: {
        y: { ticks: { font: { size: 11 } } },
        x: { beginAtZero: true, ticks: { stepSize: 1, font: { size: 10 } } }
      }
    },
    plugins: [ChartDataLabels]
  });
}

// ── Premium Slide ───────────────────────────────────────────────────────────

function loadPremiumSection(wb) {
  const premSheet   = wb.SheetNames.find(n => n.toLowerCase().includes('premium'));
  const premSection = $('premium-section');

  if (!premSheet || !premSection) {
    if (premSection) premSection.classList.add('hidden');
    return;
  }

  premiumState.sheetName = premSheet;
  premiumState.data    = XLSX.utils.sheet_to_json(wb.Sheets[premSheet], { defval: '' });
  premiumState.columns = detectColumns(premiumState.data);

  const sheetLabel = $('premium-sheet-label');
  if (sheetLabel) sheetLabel.textContent = premSheet;

  const total = premiumState.data.length;
  $('premium-hero-total').textContent = total.toLocaleString('en-IN');
  $('premium-hero-sub').textContent   = `${total.toLocaleString('en-IN')} total records`;

  premSection.classList.remove('hidden');
  renderPremiumList();
}

function renderPremiumList() {
  const listEl  = $('premium-game-list');
  const countEl = $('premium-search-count');
  if (!listEl) return;

  const cols = premiumState.columns;
  if (!cols.length || !premiumState.data.length) {
    listEl.innerHTML = '<div class="premium-game-empty">No premium data loaded</div>';
    return;
  }

  // Detect game name column
  const nameCol = (
    cols.find(c => /game.?name/i.test(c.name))  ||
    cols.find(c => /game.?title/i.test(c.name)) ||
    cols.find(c => /\bname\b/i.test(c.name))    ||
    cols.find(c => /\btitle\b/i.test(c.name))   ||
    cols.find(c => c.type === 'text')            ||
    cols[0]
  );

  // Detect partner column
  const partnerCol = (
    cols.find(c => /partner/i.test(c.name))     ||
    cols.find(c => /publisher/i.test(c.name))   ||
    cols.find(c => /developer/i.test(c.name))   ||
    cols.find(c => /studio/i.test(c.name))      ||
    cols.find(c => /vendor/i.test(c.name))      ||
    null
  );

  const query = ($('premium-search')?.value || '').trim().toLowerCase();

  let rows = premiumState.data;
  if (query) {
    rows = rows.filter(row =>
      Object.values(row).some(v => String(v).toLowerCase().includes(query))
    );
  }

  if (countEl) countEl.textContent = rows.length + ' / ' + premiumState.data.length;

  if (!rows.length) {
    listEl.innerHTML = '<div class="premium-game-empty">No games found</div>';
    return;
  }

  listEl.innerHTML = rows.map((row, i) => {
    const name    = escHtml(String(row[nameCol.name]           || '—').trim());
    const partner = partnerCol ? escHtml(String(row[partnerCol.name] || '').trim()) : '';
    return `<button class="premium-game-item" data-idx="${i}">
      <span class="premium-game-num">${i + 1}.</span>
      <span class="premium-game-info">
        <span class="premium-game-name">${name}</span>
        ${partner ? `<span class="premium-game-partner">${partner}</span>` : ''}
      </span>
      <span class="premium-game-arrow">›</span>
    </button>`;
  }).join('');

  listEl.querySelectorAll('.premium-game-item').forEach((btn, i) => {
    btn.addEventListener('click', () => showGameDetail(rows[i], premiumState.columns, 'premium'));
  });
}

// ── GameSnacks Slide ────────────────────────────────────────────────────────

function loadGameSnacksSection(wb) {
  const gsSheet   = wb.SheetNames.find(n => /game\s*snacks/i.test(n));
  const gsSection = $('gamesnacks-section');

  if (!gsSheet || !gsSection) {
    if (gsSection) gsSection.classList.add('hidden');
    return;
  }

  gamesnacksState.sheetName = gsSheet;
  gamesnacksState.data    = XLSX.utils.sheet_to_json(wb.Sheets[gsSheet], { defval: '' });
  gamesnacksState.columns = detectColumns(gamesnacksState.data);

  const sheetLabel = $('gs-sheet-label');
  if (sheetLabel) sheetLabel.textContent = gsSheet;

  const total = gamesnacksState.data.length;
  $('gs-hero-total').textContent = total.toLocaleString('en-IN');
  $('gs-hero-sub').textContent   = `${total.toLocaleString('en-IN')} total records`;

  gsSection.classList.remove('hidden');
  renderGameSnacksList();
}

function renderGameSnacksList() {
  const listEl  = $('gs-game-list');
  const countEl = $('gs-search-count');
  if (!listEl) return;

  const cols = gamesnacksState.columns;
  if (!cols.length || !gamesnacksState.data.length) {
    listEl.innerHTML = '<div class="gs-game-empty">No GameSnacks data loaded</div>';
    return;
  }

  const nameCol = (
    cols.find(c => /game.?name/i.test(c.name))  ||
    cols.find(c => /game.?title/i.test(c.name)) ||
    cols.find(c => /\bname\b/i.test(c.name))    ||
    cols.find(c => /\btitle\b/i.test(c.name))   ||
    cols.find(c => c.type === 'text')            ||
    cols[0]
  );

  const partnerCol = (
    cols.find(c => /partner/i.test(c.name))   ||
    cols.find(c => /publisher/i.test(c.name)) ||
    cols.find(c => /developer/i.test(c.name)) ||
    cols.find(c => /studio/i.test(c.name))    ||
    cols.find(c => /vendor/i.test(c.name))    ||
    null
  );

  const query = ($('gs-search')?.value || '').trim().toLowerCase();

  let rows = gamesnacksState.data;
  if (query) {
    rows = rows.filter(row =>
      Object.values(row).some(v => String(v).toLowerCase().includes(query))
    );
  }

  if (countEl) countEl.textContent = rows.length + ' / ' + gamesnacksState.data.length;

  if (!rows.length) {
    listEl.innerHTML = '<div class="gs-game-empty">No games found</div>';
    return;
  }

  listEl.innerHTML = rows.map((row, i) => {
    const name    = escHtml(String(row[nameCol.name]                || '—').trim());
    const partner = partnerCol ? escHtml(String(row[partnerCol.name] || '').trim()) : '';
    return `<button class="gs-game-item" data-idx="${i}">
      <span class="gs-game-num">${i + 1}.</span>
      <span class="gs-game-info">
        <span class="gs-game-name">${name}</span>
        ${partner ? `<span class="gs-game-partner">${partner}</span>` : ''}
      </span>
      <span class="gs-game-arrow">›</span>
    </button>`;
  }).join('');

  listEl.querySelectorAll('.gs-game-item').forEach((btn, i) => {
    btn.addEventListener('click', () => showGameDetail(rows[i], gamesnacksState.columns, 'gamesnacks'));
  });
}

function renderLiveSerialList() {
  const listEl  = $('live-serial-list');
  const countEl = $('live-serial-count');
  if (!listEl) return;

  const platCol = ui.livePlatformCol ? ui.livePlatformCol.value : '';

  // Detect game name column
  const cols = liveState.columns;
  const nameCol = (
    cols.find(c => /game.?name/i.test(c.name))  ||
    cols.find(c => /game.?title/i.test(c.name)) ||
    cols.find(c => /\bname\b/i.test(c.name))    ||
    cols.find(c => /\btitle\b/i.test(c.name))   ||
    cols.find(c => c.type === 'text')            ||
    cols[0] || null
  );

  if (!nameCol || !liveState.data.length) {
    listEl.innerHTML = '<div class="live-serial-empty">No data loaded</div>';
    return;
  }

  // Exclude StoreFront rows
  let rows = platCol
    ? liveState.data.filter(row => !/store\s*front/i.test(String(row[platCol] || '')))
    : liveState.data;

  // Apply search filter
  const query = ($('live-serial-search')?.value || '').trim().toLowerCase();
  if (query) {
    rows = rows.filter(row => String(row[nameCol.name] || '').toLowerCase().includes(query));
  }

  if (countEl) countEl.textContent = rows.length + ' games';

  if (!rows.length) {
    listEl.innerHTML = '<div class="live-serial-empty">No games found</div>';
    return;
  }

  listEl.innerHTML = rows.map((row, i) => {
    const name = escHtml(String(row[nameCol.name] || '—').trim());
    const plat = platCol ? escHtml(String(row[platCol] || '').trim()) : '';
    return `<button class="live-serial-item" data-idx="${i}" title="Click to view details">
      <span class="live-serial-num">${i + 1}.</span>
      <span class="live-serial-name">${name}</span>
      ${plat ? `<span class="live-serial-plat">${plat}</span>` : ''}
      <span class="live-serial-arrow">›</span>
    </button>`;
  }).join('');

  // Click → show full detail modal
  listEl.querySelectorAll('.live-serial-item').forEach((btn, i) => {
    btn.addEventListener('click', () => showGameDetail(rows[i], liveState.columns, 'live'));
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
// theme  — 'pqa' → purple, 'live' → green, else default; also accepts true for legacy PQA calls
function showGameDetail(row, cols, theme) {
  cols = cols || state.columns;
  if (!cols.length) return;

  const isPQA        = theme === 'pqa' || theme === true;
  const isLive       = theme === 'live';
  const isPremium    = theme === 'premium';
  const isGameSnacks = theme === 'gamesnacks';

  const titleCol =
    cols.find(c => /game.?name/i.test(c.name))  ||
    cols.find(c => /game.?title/i.test(c.name)) ||
    cols.find(c => /\bname\b/i.test(c.name))    ||
    cols.find(c => /\btitle\b/i.test(c.name))   ||
    cols.find(c => c.type === 'text')            ||
    cols[0];

  const title = String(row[titleCol.name] || '—');
  $('game-modal-title').textContent = title;

  // Header colour: purple for PQA, green for Live, default dark for QA
  const modalHeader = document.querySelector('.game-modal-header');
  if (modalHeader) {
    modalHeader.style.background = isPQA ? '#2d1b4e' : isLive ? '#0d3322' : isPremium ? '#431407' : isGameSnacks ? '#0c2340' : '';
  }
  const modalLabel = document.querySelector('.game-modal-label');
  if (modalLabel) {
    modalLabel.textContent = isPQA ? 'PQA Game Details' : isLive ? 'Live Game Details' : isPremium ? 'Premium Game Details' : isGameSnacks ? 'GameSnacks Details' : 'Game Details';
    modalLabel.style.color = isPQA ? '#a855f7' : isLive ? '#10b981' : isPremium ? '#f97316' : isGameSnacks ? '#06b6d4' : '';
  }
  const modalIcon = document.querySelector('.game-modal-icon');
  if (modalIcon) {
    modalIcon.style.background = isPQA ? 'rgba(168,85,247,0.15)' : isLive ? 'rgba(16,185,129,0.15)' : isPremium ? 'rgba(249,115,22,0.15)' : isGameSnacks ? 'rgba(6,182,212,0.15)' : '';
    modalIcon.style.color      = isPQA ? '#a855f7' : isLive ? '#10b981' : isPremium ? '#f97316' : isGameSnacks ? '#06b6d4' : '';
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

// ── Live Slide ──────────────────────────────────────────────────────────────

function loadLiveSection(wb) {
  const liveSheet  = wb.SheetNames.find(n => n.toLowerCase().includes('live'));
  const liveSection = $('live-section');

  if (!liveSheet || !liveSection) {
    if (liveSection) liveSection.classList.add('hidden');
    return;
  }

  liveState.sheetName = liveSheet;
  liveState.data    = XLSX.utils.sheet_to_json(wb.Sheets[liveSheet], { defval: '' });
  liveState.columns = detectColumns(liveState.data);

  const sheetLabel = $('live-sheet-label');
  if (sheetLabel) sheetLabel.textContent = liveSheet;

  // Populate column selects
  const catCols = liveState.columns.filter(c => c.type !== 'number').map(c => c.name);
  const choices  = (catCols.length ? catCols : liveState.columns.map(c => c.name)).filter(c => !CHART_SKIP.test(c));
  const allCols  = liveState.columns.map(c => c.name);
  const opts = choices.map(o => `<option value="${escHtml(o)}">${escHtml(o)}</option>`).join('');
  const allOpts = allCols.map(o => `<option value="${escHtml(o)}">${escHtml(o)}</option>`).join('');
  if (ui.livePlatformCol) ui.livePlatformCol.innerHTML = opts;
  if (ui.liveDevCol) ui.liveDevCol.innerHTML = allOpts;
  if (ui.liveDayCol) ui.liveDayCol.innerHTML = allOpts;

  // Populate V Count select with numeric columns
  const numCols = liveState.columns.filter(c => c.type === 'number').map(c => c.name);
  if (ui.liveVCountCol) {
    ui.liveVCountCol.innerHTML = '<option value="">(none)</option>' +
      numCols.map(o => `<option value="${escHtml(o)}">${escHtml(o)}</option>`).join('');
    const vcDef = numCols.find(c => /v[\s_-]?count|viewer|vcount/i.test(c));
    if (vcDef) ui.liveVCountCol.value = vcDef;
  }

  function detect(keywords) {
    return allCols.find(c => keywords.some(k => c.toLowerCase().includes(k))) || allCols[0] || '';
  }
  const pDef   = detect(['platform', 'device', 'type', 'sp', 'stb', 'jp']);
  const devDef = detect(['developer', 'dev', 'studio', 'publisher']);
  const dayDef = detect(['live day', 'liveday', 'day', 'date', 'go live']);

  if (ui.livePlatformCol && pDef)   ui.livePlatformCol.value = pDef;
  if (ui.liveDevCol     && devDef)  ui.liveDevCol.value      = devDef;
  if (ui.liveDayCol     && dayDef)  ui.liveDayCol.value      = dayDef;

  liveSection.classList.remove('hidden');
  renderLiveKPIs();
  renderLiveCharts();
  renderLiveSerialList();
}

function renderLiveKPIs() {
  const platCol = ui.livePlatformCol ? ui.livePlatformCol.value : '';

  // Exclude StoreFront rows from all Live KPI calculations
  const liveData = platCol
    ? liveState.data.filter(row => !/store\s*front/i.test(String(row[platCol] || '')))
    : liveState.data;

  const total = liveData.length;
  $('live-hero-total').textContent = total.toLocaleString('en-IN');
  $('live-hero-sub').textContent   = `${total.toLocaleString('en-IN')} total records`;

  if (platCol) {
    const counts = aggregate(liveData, platCol, '__count__', 'count', 'none');

    // Sum ALL entries whose key matches any of the given keywords
    const find = (...tags) => {
      const total = counts
        .filter(e => tags.some(t => e.key.toLowerCase().includes(t.toLowerCase())))
        .reduce((sum, e) => sum + e.value, 0);
      return total ? total.toLocaleString('en-IN') : '0';
    };

    $('live-kpi-sp-val').textContent  = find('sp', 'mobile', 'smartphone');
    $('live-kpi-stb-val').textContent = find('stb');
    $('live-kpi-jp-val').textContent  = find('jp', 'jiophone', 'jio phone', 'candy');
  }
}

function renderLiveCharts() {
  renderLiveDetailsChart();
  renderLivePlatform();
}

function renderLiveStatus() {
  destroyChart('liveStatus');
  const col = ui.liveStatusCol ? ui.liveStatusCol.value : '';
  if (!col || !liveState.data.length) return;

  const entries = aggregate(liveState.data, col, '__count__', 'count', 'value_desc')
    .filter(e => !/store\s*front/i.test(e.key))
    .slice(0, 12);
  const ctx = $('live-status-chart').getContext('2d');
  state.charts.liveStatus = new Chart(ctx, {
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

function renderLiveDetailsChart() {
  destroyChart('liveDetails');
  const groupCol = ui.liveDevCol?.value  || '';
  const dayCol   = ui.liveDayCol?.value  || '';
  const platCol  = ui.livePlatformCol?.value || '';
  if (!groupCol || !liveState.data.length) return;

  let data = platCol
    ? liveState.data.filter(row => !/store\s*front/i.test(String(row[platCol] || '')))
    : liveState.data;

  const entries = aggregate(data, groupCol, '__count__', 'count', 'value_desc').slice(0, 15);

  const ctx = $('live-details-chart');
  if (!ctx) return;
  state.charts.liveDetails = new Chart(ctx.getContext('2d'), {
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
      cutout: '55%',
      plugins: {
        legend: {
          position: 'right',
          labels: {
            font: { size: 11 }, boxWidth: 12, padding: 10,
            generateLabels: chart => {
              const d = chart.data;
              return d.labels.map((label, i) => ({
                text: `${label}  (${d.datasets[0].data[i]})`,
                fillStyle: d.datasets[0].backgroundColor[i],
                strokeStyle: '#fff',
                lineWidth: 2,
                hidden: false,
                index: i,
              }));
            }
          }
        },
        tooltip: {
          callbacks: {
            label: ctx => {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct   = total ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
              return ` ${ctx.label}: ${ctx.parsed}  (${pct}%)`;
            }
          }
        },
        datalabels: {
          display: ctx => ctx.dataset.data[ctx.dataIndex] > 0,
          color: '#fff',
          font: { size: 11, weight: 'bold' },
          formatter: v => v,
          textStrokeColor: 'rgba(0,0,0,0.4)',
          textStrokeWidth: 2,
        }
      }
    },
    plugins: [ChartDataLabels]
  });
}

function renderLivePlatform() {
  destroyChart('livePlatform');
  const col   = ui.livePlatformCol ? ui.livePlatformCol.value : '';
  const vcCol = ui.liveVCountCol   ? ui.liveVCountCol.value   : '';
  if (!col || !liveState.data.length) return;

  const entries = aggregate(liveState.data, col, '__count__', 'count', 'value_desc')
    .filter(e => !/store\s*front/i.test(e.key))
    .slice(0, 15);

  const datasets = [{
    label: 'Live Games',
    data: entries.map(e => e.value),
    backgroundColor: entries.map((_, i) => PALETTE[i % PALETTE.length] + 'cc'),
    borderColor:     entries.map((_, i) => PALETTE[i % PALETTE.length]),
    borderWidth: 1,
    borderRadius: 5,
    xAxisID: 'x',
  }];

  if (vcCol) {
    const vcMap = Object.fromEntries(
      aggregate(liveState.data, col, vcCol, 'sum', 'none').map(e => [e.key, e.value])
    );
    datasets.push({
      label: vcCol,
      data: entries.map(e => vcMap[e.key] ?? 0),
      backgroundColor: '#6366f155',
      borderColor: '#6366f1',
      borderWidth: 1,
      borderRadius: 5,
      xAxisID: 'x2',
    });
  }

  const ctx = $('live-platform-chart').getContext('2d');
  state.charts.livePlatform = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: entries.map(e => e.key),
      datasets,
    },
    options: {
      ...CHART_DEFAULTS,
      indexAxis: 'y',
      layout: { padding: { right: 36 } },
      plugins: {
        ...CHART_DEFAULTS.plugins,
        legend: { display: !!vcCol, labels: { font: { size: 11 }, boxWidth: 12 } },
        tooltip: {
          callbacks: {
            label: ctx => {
              const v = ctx.parsed.x;
              return ` ${ctx.dataset.label}: ${formatNumber(v)}`;
            }
          }
        },
        datalabels: {
          anchor: 'end',
          align: 'right',
          color: '#94a3b8',
          font: { size: 11, weight: 'bold' },
          formatter: (v, ctx) => ctx.datasetIndex === 0 ? v : formatNumber(v),
        }
      },
      scales: {
        y: { ticks: { font: { size: 11 } } },
        x: {
          beginAtZero: true,
          ticks: { stepSize: 1, font: { size: 10 } },
          title: { display: !!vcCol, text: 'Games Count', font: { size: 10 } },
        },
        ...(vcCol ? {
          x2: {
            type: 'linear',
            position: 'top',
            beginAtZero: true,
            ticks: { font: { size: 10 }, callback: v => formatNumber(v) },
            title: { display: true, text: vcCol, font: { size: 10 }, color: '#6366f1' },
            grid: { drawOnChartArea: false },
          }
        } : {}),
      }
    },
    plugins: [ChartDataLabels]
  });
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

  // Microsoft Sign-in / SharePoint buttons
  const btnMsSignin  = $('btn-ms-signin');
  const btnMsSignout = $('btn-ms-signout');
  const btnSpLoad    = $('btn-sp-load');
  if (btnMsSignin)  btnMsSignin.addEventListener('click', msSignIn);
  if (btnMsSignout) btnMsSignout.addEventListener('click', msSignOut);
  if (btnSpLoad)    btnSpLoad.addEventListener('click', loadFromSharePoint);

  // URL paste load (no sign-in)
  const btnUrlLoad = $('btn-url-load');
  if (btnUrlLoad) btnUrlLoad.addEventListener('click', loadFromUrl);
  const urlPasteInput = $('url-paste-input');
  if (urlPasteInput) urlPasteInput.addEventListener('keydown', e => { if (e.key === 'Enter') loadFromUrl(); });


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

  // Live chart control changes
  if (ui.livePlatformCol) ui.livePlatformCol.addEventListener('change', () => { renderLiveKPIs(); renderLivePlatform(); renderLiveDetailsChart(); });
  if (ui.liveVCountCol)   ui.liveVCountCol.addEventListener('change', renderLivePlatform);
  if (ui.liveDevCol)      ui.liveDevCol.addEventListener('change', renderLiveDetailsChart);
  if (ui.liveDayCol)      ui.liveDayCol.addEventListener('change', renderLiveDetailsChart);

  // Live serial list search
  const liveSerialSearch = $('live-serial-search');
  if (liveSerialSearch) liveSerialSearch.addEventListener('input', renderLiveSerialList);

  // Premium game search
  const premiumSearch = $('premium-search');
  if (premiumSearch) premiumSearch.addEventListener('input', renderPremiumList);

  // GameSnacks search
  const gsSearch = $('gs-search');
  if (gsSearch) gsSearch.addEventListener('input', renderGameSnacksList);

  // Chart control changes
  ui.statusCol.addEventListener('change', () => { renderStatus(); renderKPIs(); });
  ui.platformCol.addEventListener('change', () => { renderPlatform(); renderKPIs(); });

  // PQA chart control changes
  if (ui.pqaStatusCol)   ui.pqaStatusCol.addEventListener('change',   renderPQAStatus);
  if (ui.pqaPlatformCol) ui.pqaPlatformCol.addEventListener('change', () => { renderPQAKPIs(); renderPQAPlatform(); });

  // Scroll-based active tab
  initScrollActiveTabs();
}

// Boot
init();
initMSAL();
// Show hint on file:// protocol (Live Server not running)
if (location.protocol === 'file:') {
  const hint = document.getElementById('auto-load-hint');
  if (hint) hint.style.display = 'block';
} else {
  autoLoadDefaultFile();
}
