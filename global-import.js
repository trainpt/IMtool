// ═══════════════════════════════════════════
// Global Import + Sheet Management
// Uses data-store.js for persistence
// ═══════════════════════════════════════════

let giParsedSheets = [];
let giFileName = '';
let currentPreviewSheet = null;

// ─── Init ───
function initGlobalImport() {
  // Import button
  document.getElementById('btnGlobalImport').addEventListener('click', showImportModal);
  document.getElementById('importCloseBtn').addEventListener('click', closeImportModal);
  document.getElementById('importBackBtn').addEventListener('click', importGoBack);

  // Dropzone
  const dz = document.getElementById('importDropzone');
  dz.addEventListener('click', () => document.getElementById('importFileInput').click());
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
  dz.addEventListener('drop', e => {
    e.preventDefault(); dz.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) handleImportFile(e.dataTransfer.files[0]);
  });
  document.getElementById('importFileInput').addEventListener('change', e => {
    if (e.target.files[0]) handleImportFile(e.target.files[0]);
    e.target.value = '';
  });

  // Sheet preview actions
  document.getElementById('sheetPreviewFilter').addEventListener('input', e => filterPreviewTable(e.target.value));
  document.getElementById('sheetSendMerge').addEventListener('click', sendSheetToMerge);
  document.getElementById('sheetOpenReview').addEventListener('click', () => sendSheetToTool('review'));
  document.getElementById('sheetOpenEditor').addEventListener('click', () => sendSheetToTool('coleditor'));
  document.getElementById('sheetDeleteBtn').addEventListener('click', deleteCurrentSheet);

  // Save/Load
  document.getElementById('btnSave').addEventListener('click', saveToFile);
  document.getElementById('loadFileInput').addEventListener('change', loadFromFile);

  renderSheetList();
}

// ─── Import Modal ───
function showImportModal() {
  giParsedSheets = [];
  giFileName = '';
  document.getElementById('importStep1').style.display = '';
  document.getElementById('importStep2').style.display = 'none';
  document.getElementById('importBackBtn').style.display = 'none';
  document.getElementById('importSuccess').style.display = 'none';
  document.getElementById('globalImportModal').classList.add('show');
}

function closeImportModal() {
  document.getElementById('globalImportModal').classList.remove('show');
}

function importGoBack() {
  document.getElementById('importStep1').style.display = '';
  document.getElementById('importStep2').style.display = 'none';
  document.getElementById('importBackBtn').style.display = 'none';
  document.getElementById('importSuccess').style.display = 'none';
}

// ─── File Handling ───
async function handleImportFile(file) {
  giFileName = file.name;
  try {
    const result = await readUploadedFile(file);
    giParsedSheets = result.sheets;
    if (giParsedSheets.length === 0) { alert('No usable sheets found.'); return; }
    showImportStep2();
  } catch(e) {
    alert('Error reading file: ' + e.message);
  }
}

function showImportStep2() {
  document.getElementById('importStep1').style.display = 'none';
  document.getElementById('importStep2').style.display = '';
  document.getElementById('importBackBtn').style.display = '';
  document.getElementById('importSuccess').style.display = 'none';

  document.getElementById('importFileName').textContent = giFileName + ' — ' + giParsedSheets.length + ' sheet' + (giParsedSheets.length > 1 ? 's' : '');

  // Sheet list with checkboxes
  const listEl = document.getElementById('importSheetList');
  let html = '';
  giParsedSheets.forEach((sheet, idx) => {
    const exists = getSheet(sheet.name);
    const badge = exists ? ' <span class="text-muted small">(will replace)</span>' : '';
    html += '<label class="import-sheet-row"><input type="checkbox" checked data-idx="' + idx + '">' +
      '<span class="name">' + escHtml(sheet.name) + badge + '</span>' +
      '<span class="info">' + sheet.rows.length + ' rows</span></label>';
  });
  html += '<button class="btn btn-primary" style="width:100%;margin-top:10px;" id="importGoBtn">Import Selected</button>';
  listEl.innerHTML = html;
  document.getElementById('importGoBtn').addEventListener('click', doImport);

  // Org select
  const select = document.getElementById('importOrgSelect');
  select.innerHTML = '<option value="">-- Select existing --</option>';
  getAllOrgNames().forEach(name => {
    const opt = document.createElement('option');
    opt.value = name; opt.textContent = name;
    select.appendChild(opt);
  });
  document.getElementById('importNewOrg').value = '';
}

function doImport() {
  const orgSelect = document.getElementById('importOrgSelect').value;
  const newOrg = document.getElementById('importNewOrg').value.trim();
  const orgName = newOrg || orgSelect;
  if (!orgName) { alert('Select or enter an organization name.'); return; }

  // Ensure org exists
  createOrg(orgName);

  const checkboxes = document.querySelectorAll('#importSheetList input[type="checkbox"]:checked');
  if (checkboxes.length === 0) { alert('Select at least one sheet.'); return; }

  let count = 0;
  checkboxes.forEach(cb => {
    const sheet = giParsedSheets[parseInt(cb.dataset.idx)];
    if (sheet) { saveSheet(sheet.name, sheet.headers, sheet.rows, giFileName, orgName); count++; }
  });

  document.getElementById('importSuccess').innerHTML = '<strong>' + count + ' sheet(s)</strong> imported to <strong>' + escHtml(orgName) + '</strong>';
  document.getElementById('importSuccess').style.display = '';
  renderSheetList();
}

// ─── Sidebar Sheet List ───
function renderSheetList() {
  const container = document.getElementById('sheetList');
  if (!container) return;
  const store = getStore();
  const names = Object.keys(store.sheets || {}).sort();

  if (names.length === 0) {
    container.innerHTML = '<div class="sidebar-empty">No data imported</div>';
    return;
  }

  // Group by org
  const byOrg = {};
  names.forEach(name => {
    const org = store.sheets[name].orgName || 'Unassigned';
    if (!byOrg[org]) byOrg[org] = [];
    byOrg[org].push({ name, ...store.sheets[name] });
  });

  let html = '';
  Object.keys(byOrg).sort().forEach(orgName => {
    const sheets = byOrg[orgName];
    const isOpen = localStorage.getItem('sheetOrg_' + orgName) !== 'closed';
    html += '<div class="sb-org-group">';
    html += '<div class="sb-org-group-header" onclick="toggleSheetOrg(\'' + escJs(orgName) + '\')">';
    html += '<span class="arrow' + (isOpen ? ' open' : '') + '" id="sheetArrow_' + escHtml(orgName) + '">&#9654;</span>';
    html += '<span class="name">' + escHtml(orgName) + '</span>';
    html += '<span class="count">' + sheets.length + '</span>';
    html += '</div>';
    html += '<div class="sb-org-group-list' + (isOpen ? ' open' : '') + '" id="sheetList_' + escHtml(orgName) + '">';
    sheets.forEach(sheet => {
      const active = currentPreviewSheet === sheet.name ? ' active' : '';
      html += '<div class="sb-item' + active + '" onclick="previewSheet(\'' + escJs(sheet.name) + '\')" title="' + escHtml(sheet.name) + '">';
      html += '<span class="name">' + escHtml(sheet.name) + '</span>';
      html += '<span class="count">' + sheet.rowCount + '</span>';
      html += '<span class="del" onclick="event.stopPropagation();confirmDeleteSheet(\'' + escJs(sheet.name) + '\')" title="Remove">&#10005;</span>';
      html += '</div>';
    });
    html += '</div></div>';
  });
  container.innerHTML = html;
}

function toggleSheetOrg(orgName) {
  const list = document.getElementById('sheetList_' + orgName);
  const arrow = document.getElementById('sheetArrow_' + orgName);
  if (list) list.classList.toggle('open');
  if (arrow) arrow.classList.toggle('open');
  localStorage.setItem('sheetOrg_' + orgName, list?.classList.contains('open') ? 'open' : 'closed');
}

function confirmDeleteSheet(name) {
  if (!confirm('Remove "' + name + '" from imported sheets?')) return;
  deleteSheet(name);
  renderSheetList();
  if (currentPreviewSheet === name) { currentPreviewSheet = null; switchPage('merge'); }
}

// ─── Sheet Preview ───
function previewSheet(name) {
  const sheet = getSheet(name);
  if (!sheet) return;
  currentPreviewSheet = name;
  switchPage('sheetpreview');
  renderSheetList(); // update active state

  document.getElementById('sheetPreviewTitle').textContent = name + ' (' + sheet.rows.length + ' rows' + (sheet.orgName ? ' — ' + sheet.orgName : '') + ')';
  document.getElementById('sheetPreviewFilter').value = '';

  let html = '<table class="data-table" id="sheetPreviewTable"><thead><tr>';
  sheet.headers.forEach(h => { html += '<th>' + escHtml(h) + '</th>'; });
  html += '</tr></thead><tbody>';
  sheet.rows.forEach(row => {
    html += '<tr>';
    row.forEach(cell => { html += '<td>' + escHtml(String(cell)) + '</td>'; });
    for (let i = row.length; i < sheet.headers.length; i++) html += '<td></td>';
    html += '</tr>';
  });
  html += '</tbody></table>';
  document.getElementById('sheetPreviewContent').innerHTML = html;

  // Click to copy on cells
  document.querySelectorAll('#sheetPreviewTable td').forEach(td => {
    td.addEventListener('click', () => {
      const text = td.textContent.trim();
      if (!text) return;
      navigator.clipboard.writeText(text);
      const prev = document.querySelector('#sheetPreviewTable .cell-clicked');
      if (prev) prev.classList.remove('cell-clicked');
      td.classList.add('cell-clicked');
    });
  });

  // Sort on headers
  document.querySelectorAll('#sheetPreviewTable th').forEach((th, ci) => {
    th.dataset.sortDir = '';
    th.addEventListener('click', () => {
      const tbody = document.querySelector('#sheetPreviewTable tbody');
      const rows = Array.from(tbody.querySelectorAll('tr'));
      const dir = th.dataset.sortDir === 'asc' ? 'desc' : 'asc';
      document.querySelectorAll('#sheetPreviewTable th').forEach(h => { h.dataset.sortDir = ''; h.classList.remove('sort-active'); });
      th.dataset.sortDir = dir; th.classList.add('sort-active');
      rows.sort((a, b) => {
        const av = (a.children[ci]?.textContent || '').trim();
        const bv = (b.children[ci]?.textContent || '').trim();
        const an = parseFloat(av), bn = parseFloat(bv);
        if (!isNaN(an) && !isNaN(bn)) return dir === 'asc' ? an - bn : bn - an;
        return dir === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av);
      });
      rows.forEach(r => tbody.appendChild(r));
    });
  });
}

function filterPreviewTable(query) {
  const q = query.toLowerCase().trim();
  document.querySelectorAll('#sheetPreviewTable tbody tr').forEach(tr => {
    tr.style.display = (!q || tr.textContent.toLowerCase().includes(q)) ? '' : 'none';
  });
}

function sendSheetToMerge() {
  if (!currentPreviewSheet) return;
  const sheet = getSheet(currentPreviewSheet);
  if (!sheet) return;
  switchPage('merge');
  const sheetRows = [sheet.headers, ...sheet.rows];
  if (typeof processSheet === 'function') processSheet(sheetRows);
}

function sendSheetToTool(tool) {
  if (!currentPreviewSheet) return;
  const sheet = getSheet(currentPreviewSheet);
  if (!sheet) return;
  switchPage(tool);

  if (tool === 'review' && typeof rvLoadSheetData === 'function') {
    rvLoadSheetData(sheet.headers, sheet.rows);
  } else if (tool === 'coleditor' && typeof ceLoadSheetData === 'function') {
    ceLoadSheetData(sheet.headers, sheet.rows);
  }
}

function deleteCurrentSheet() {
  if (!currentPreviewSheet) return;
  if (!confirm('Delete "' + currentPreviewSheet + '"?')) return;
  deleteSheet(currentPreviewSheet);
  renderSheetList();
  currentPreviewSheet = null;
  switchPage('merge');
}

// ─── Save / Load (full app state) ───
let fileHandle = null;

async function saveToFile() {
  const data = {
    store: getStore(),
    savedAt: new Date().toISOString()
  };

  if (window.showSaveFilePicker) {
    try {
      if (!fileHandle) {
        fileHandle = await window.showSaveFilePicker({
          suggestedName: 'im-tools-data.json',
          types: [{ description: 'JSON', accept: { 'application/json': ['.json'] } }]
        });
      }
      const writable = await fileHandle.createWritable();
      await writable.write(JSON.stringify(data, null, 2));
      await writable.close();
      document.getElementById('saveStatus').textContent = 'Saved ' + new Date().toLocaleTimeString();
      document.getElementById('saveStatus').className = 'save-status saved';
      return;
    } catch(e) {
      if (e.name === 'AbortError') return;
    }
  }

  // Fallback download
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = 'im-tools-data.json'; a.click();
  URL.revokeObjectURL(a.href);
  document.getElementById('saveStatus').textContent = 'Downloaded';
  document.getElementById('saveStatus').className = 'save-status saved';
}

function loadFromFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = evt => {
    try {
      const loaded = JSON.parse(evt.target.result);
      if (loaded.store) {
        localStorage.setItem('imToolsData_v2', JSON.stringify(loaded.store));
      } else if (loaded.data) {
        // Old format compatibility
        localStorage.setItem('imToolsData_v2', JSON.stringify(loaded.data));
      }
      alert('Data loaded. Refreshing...');
      location.reload();
    } catch(err) { alert('Failed: ' + err.message); }
  };
  reader.readAsText(file);
  e.target.value = '';
}
