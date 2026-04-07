const HEADERS = ['Location','Sites','Crop','Variety','Location Type','Planted Date','Acreage','Plant Count'];
const RANCH_COL = 1; // index of Sites column
const ACREAGE_COL = 6;
const PLANTCOUNT_COL = 7;
const SKIP_FOR_DONE = new Set([RANCH_COL]);
const DATA_KEY = 'tableAppData';

let currentOrg = null;
let currentRanch = null;
let copiedCells = {};
let moveRowIndex = null;
let sortCol = -1;
let sortAsc = true;
const NUMERIC_COLS = new Set([ACREAGE_COL, PLANTCOUNT_COL]);

// --- Data store ---
// Structure: { orgs: { "OrgName": { ranches: { "RanchName": { rows: [], progress: {} } } } } }
function getData() {
  try { return JSON.parse(localStorage.getItem(DATA_KEY)) || { orgs: {} }; } catch(e) { return { orgs: {} }; }
}
function saveData(data) { localStorage.setItem(DATA_KEY, JSON.stringify(data)); }

function getOrg(data, org) { return data.orgs[org]; }
function getRanch(data, org, ranch) { return data.orgs[org]?.ranches?.[ranch]; }

// --- Archive store ---
const ARCHIVE_KEY = 'tableAppArchive';
function getArchive() {
  try { return JSON.parse(localStorage.getItem(ARCHIVE_KEY)) || []; } catch(e) { return []; }
}
function saveArchive(archive) { localStorage.setItem(ARCHIVE_KEY, JSON.stringify(archive)); }

function archiveItem(type, name, content, parentOrg) {
  const archive = getArchive();
  archive.push({
    type: type, // 'org', 'ranch', 'rows', 'row'
    name: name,
    parentOrg: parentOrg || null,
    content: content,
    archivedAt: new Date().toISOString()
  });
  saveArchive(archive);
}

function restoreArchiveItem(index) {
  const archive = getArchive();
  const item = archive[index];
  if (!item) return;
  const data = getData();

  if (item.type === 'org') {
    if (data.orgs[item.name]) {
      // Merge ranches back
      const restored = item.content;
      for (const r of Object.keys(restored.ranches || {})) {
        if (!data.orgs[item.name].ranches[r]) data.orgs[item.name].ranches[r] = restored.ranches[r];
      }
    } else {
      data.orgs[item.name] = item.content;
    }
  } else if (item.type === 'ranch') {
    if (!data.orgs[item.parentOrg]) data.orgs[item.parentOrg] = { ranches: {} };
    if (data.orgs[item.parentOrg].ranches[item.name]) {
      // Append rows
      const existing = data.orgs[item.parentOrg].ranches[item.name];
      const restored = item.content;
      (restored.rows || []).forEach((row, i) => {
        const ni = existing.rows.length;
        existing.rows.push(row);
        existing.progress[ni] = (restored.progress && restored.progress[i]) || [];
      });
    } else {
      data.orgs[item.parentOrg].ranches[item.name] = item.content;
    }
  } else if (item.type === 'rows') {
    // Cleared ranch data
    if (!data.orgs[item.parentOrg]) data.orgs[item.parentOrg] = { ranches: {} };
    if (!data.orgs[item.parentOrg].ranches[item.name]) {
      data.orgs[item.parentOrg].ranches[item.name] = item.content;
    } else {
      const existing = data.orgs[item.parentOrg].ranches[item.name];
      (item.content.rows || []).forEach((row, i) => {
        const ni = existing.rows.length;
        existing.rows.push(row);
        existing.progress[ni] = (item.content.progress && item.content.progress[i]) || [];
      });
    }
  } else if (item.type === 'row') {
    if (data.orgs[item.parentOrg]?.ranches?.[item.name]) {
      const ranch = data.orgs[item.parentOrg].ranches[item.name];
      const ni = ranch.rows.length;
      ranch.rows.push(item.content.row);
      ranch.progress[ni] = item.content.progress || [];
    }
  }

  saveData(data);
  archive.splice(index, 1);
  saveArchive(archive);
  renderSidebar();
  renderMain();
}

function permanentDeleteArchive(index) {
  if (!confirm('Permanently delete this item? This cannot be undone.')) return;
  const archive = getArchive();
  archive.splice(index, 1);
  saveArchive(archive);
  renderSidebar();
}

function renderArchiveSection() {
  const container = document.getElementById('archiveSection');
  const archive = getArchive();
  if (archive.length === 0) { container.innerHTML = ''; return; }

  let html = '<div class="archive-header" onclick="toggleArchiveList()"><span class="arrow" id="archiveArrow">&#9654;</span>Archived<span class="archive-count">(' + archive.length + ')</span></div>';
  html += '<div class="archive-list" id="archiveList">';

  archive.forEach((item, i) => {
    const label = item.type === 'org' ? item.name : (item.parentOrg + ' > ' + item.name);
    html += '<div class="archive-item">';
    html += '<span class="archive-name" title="' + escHtml(label) + '">' + escHtml(label) + '</span>';
    html += '<span class="archive-type">' + item.type + '</span>';
    html += '<span class="archive-actions">';
    html += '<span class="restore-btn" title="Restore" onclick="restoreArchiveItem(' + i + ')">Restore</span>';
    html += '<span class="permadelete-btn" title="Delete permanently" onclick="permanentDeleteArchive(' + i + ')">Delete</span>';
    html += '</span></div>';
  });

  html += '</div>';
  container.innerHTML = html;
}

function toggleArchiveList() {
  const list = document.getElementById('archiveList');
  const arrow = document.getElementById('archiveArrow');
  if (list) list.classList.toggle('open');
  if (arrow) arrow.classList.toggle('open');
}

// --- Migration from old format ---
function migrateOldData() {
  // Fix: rename any "Hillcrest Ranch" org to "Bee Sweet Citrus"
  const data = getData();
  if (data.orgs['Hillcrest Ranch']) {
    if (!data.orgs['Bee Sweet Citrus']) data.orgs['Bee Sweet Citrus'] = { ranches: {} };
    // Merge ranches into Bee Sweet Citrus
    for (const r of Object.keys(data.orgs['Hillcrest Ranch'].ranches)) {
      if (!data.orgs['Bee Sweet Citrus'].ranches[r]) {
        data.orgs['Bee Sweet Citrus'].ranches[r] = data.orgs['Hillcrest Ranch'].ranches[r];
      }
    }
    delete data.orgs['Hillcrest Ranch'];
    saveData(data);
  }

  const oldProjects = localStorage.getItem('tableProjects');
  const oldProgress = localStorage.getItem('mergedTableProgress');
  if (!oldProjects && !oldProgress) return;

  if (oldProjects) {
    try {
      const projects = JSON.parse(oldProjects);
      // All old projects migrate under "Bee Sweet Citrus" org, grouped by ranch column
      const orgName = 'Bee Sweet Citrus';
      if (!data.orgs[orgName]) data.orgs[orgName] = { ranches: {} };
      for (const name of Object.keys(projects)) {
        const proj = projects[name];
        if (!proj.rows || proj.rows.length === 0) continue;
        const grouped = {};
        proj.rows.forEach((row, i) => {
          const ranch = row[RANCH_COL] || 'Unknown Ranch';
          if (!grouped[ranch]) grouped[ranch] = { rows: [], progress: {} };
          const ni = grouped[ranch].rows.length;
          grouped[ranch].rows.push(row);
          grouped[ranch].progress[ni] = (proj.progress && proj.progress[i]) ? proj.progress[i] : [];
        });
        for (const r of Object.keys(grouped)) {
          if (!data.orgs[orgName].ranches[r]) data.orgs[orgName].ranches[r] = grouped[r];
          else {
            const existing = data.orgs[orgName].ranches[r];
            grouped[r].rows.forEach((row, i) => {
              const ni = existing.rows.length;
              existing.rows.push(row);
              existing.progress[ni] = grouped[r].progress[i] || [];
            });
          }
        }
      }
      localStorage.removeItem('tableProjects');
    } catch(e) {}
  }

  if (oldProgress) {
    // Migrate hardcoded rows
    const oldRows = [
      ["13035 (BK #5 POWELL NAVELS)","BEESHCR(HILLCREST RANCH)","Oranges","Navels","Block","7/19/22 20:44","19","2764"],
      ["13037 (BK #7 FN NAVELS)","BEESHCR(HILLCREST RANCH)","Oranges","Navels","Block","7/19/22 20:44","18","2242"],
      ["13038 (BK #8 FN NAVELS)","BEESHCR(HILLCREST RANCH)","Oranges","Navels","Block","7/19/22 20:44","18","2236"],
      ["13039 (BK #9 BARN NAVEL)","BEESHCR(HILLCREST RANCH)","Oranges","Navels","Block","7/19/22 20:44","11","1612"],
      ["13040 (BK #10 LEMON)","BEESHCR(HILLCREST RANCH)","Lemons","Lemons","Block","7/19/22 20:44","11.2","1426"],
      ["13041 (BK #11 VALS)","BEESHCR(HILLCREST RANCH)","Oranges","Valencia","Block","7/19/22 20:44","19","3416"],
      ["13042 (BK #12 STR GRF)","BEESHCR(HILLCREST RANCH)","Grapefruit","Star Ruby","Block","7/19/22 20:44","6.8","992"],
      ["13044 (BK #15 LANE NAVEL)","BEESHCR(HILLCREST RANCH)","Oranges","Navels","Block","7/19/22 20:44","7.9","1416"],
      ["13045 (BK #16 VALS)","BEESHCR(HILLCREST RANCH)","Oranges","Valencia","Block","7/19/22 20:44","11.8","2125"],
      ["13047 (BK #18 FUKO)","BEESHCR(HILLCREST RANCH)","Oranges","Navels","Block","7/19/22 20:44","19.6","3534"],
      ["13048 (BK #19 BARN NAVEL)","BEESHCR(HILLCREST RANCH)","Oranges","Navels","Block","7/19/22 20:44","8.6","1247"],
      ["13050 (BK #21 FUKO)","BEESHCR(HILLCREST RANCH)","Oranges","Navels","Block","7/19/22 20:44","20","3596"],
      ["13051 (BK #17 TANGO)","BEESHCR(HILLCREST RANCH)","Mandarins","Murcott","Block","7/19/22 20:44","18.4","3306"],
      ["13052 (BK #20 PAGE)","BEESHCR(HILLCREST RANCH)","Mandarins","Page","Block","7/19/22 20:44","9.7","1740"],
      ["13053 (BK #1 TANGO)","BEESHCR(HILLCREST RANCH)","Mandarins","Murcott","Block","7/19/22 20:44","18.6","4504"],
      ["13054 (BK #2 TANGO)","BEESHCR(HILLCREST RANCH)","Mandarins","Murcott","Block","7/19/22 20:44","18.6","4505"],
    ];
    if (!data.orgs['Bee Sweet Citrus']) data.orgs['Bee Sweet Citrus'] = { ranches: {} };
    if (!data.orgs['Bee Sweet Citrus'].ranches['BEESHCR(HILLCREST RANCH)']) {
      data.orgs['Bee Sweet Citrus'].ranches['BEESHCR(HILLCREST RANCH)'] = { rows: oldRows, progress: {} };
    }
    localStorage.removeItem('mergedTableProgress');
  }

  saveData(data);
}

// --- Org / Ranch management ---
function createOrg() {
  const input = document.getElementById('newOrgName');
  const name = input.value.trim();
  if (!name) return;
  const data = getData();
  if (data.orgs[name]) { alert('Organization already exists.'); return; }
  data.orgs[name] = { ranches: {} };
  saveData(data);
  input.value = '';
  currentOrg = name;
  renderSidebar();
}

function deleteOrg(name, e) {
  e.stopPropagation();
  if (!confirm('Archive organization "' + name + '" and all its ranches?')) return;
  const data = getData();
  archiveItem('org', name, JSON.parse(JSON.stringify(data.orgs[name])));
  delete data.orgs[name];
  saveData(data);
  if (currentOrg === name) { currentOrg = null; currentRanch = null; }
  renderSidebar();
  renderMain();
}

function createRanch(orgName) {
  const input = document.querySelector('.add-ranch-input[data-org="' + CSS.escape(orgName) + '"]');
  if (!input) return;
  const name = input.value.trim();
  if (!name) return;
  const data = getData();
  if (!data.orgs[orgName]) return;
  if (data.orgs[orgName].ranches[name]) { alert('Ranch already exists.'); return; }
  data.orgs[orgName].ranches[name] = { rows: [], progress: {} };
  saveData(data);
  input.value = '';
  currentOrg = orgName;
  currentRanch = name;
  renderSidebar();
  renderMain();
}

function deleteRanch(orgName, ranchName, e) {
  e.stopPropagation();
  if (!confirm('Archive ranch "' + ranchName + '"?')) return;
  const data = getData();
  if (data.orgs[orgName]?.ranches?.[ranchName]) {
    archiveItem('ranch', ranchName, JSON.parse(JSON.stringify(data.orgs[orgName].ranches[ranchName])), orgName);
    delete data.orgs[orgName].ranches[ranchName];
  }
  saveData(data);
  if (currentOrg === orgName && currentRanch === ranchName) currentRanch = null;
  renderSidebar();
  renderMain();
}

function selectRanch(orgName, ranchName) {
  currentOrg = orgName;
  currentRanch = ranchName;
  sortCol = -1;
  sortAsc = true;
  renderSidebar();
  renderMain();
}

function toggleOrg(orgName) {
  const list = document.querySelector('.ranch-list[data-org="' + CSS.escape(orgName) + '"]');
  const arrow = document.querySelector('.arrow[data-org="' + CSS.escape(orgName) + '"]');
  if (list) { list.classList.toggle('open'); }
  if (arrow) { arrow.classList.toggle('open'); }
}

function isRanchComplete(ranchData) {
  if (!ranchData || ranchData.rows.length === 0) return false;
  const progress = ranchData.progress || {};
  return ranchData.rows.every((row, ri) => {
    const required = row.reduce((cols, cell, ci) => {
      if (cell !== '' && !SKIP_FOR_DONE.has(ci)) cols.push(ci);
      return cols;
    }, []);
    const done = new Set(progress[ri] || []);
    return required.every(ci => done.has(ci));
  });
}

// --- Sidebar ---
function renderSidebar() {
  const container = document.getElementById('orgList');
  const data = getData();
  container.innerHTML = '';
  renderArchiveSection();

  for (const orgName of Object.keys(data.orgs).sort()) {
    const org = data.orgs[orgName];
    const section = document.createElement('div');
    section.className = 'org-section';

    // Org header
    const header = document.createElement('div');
    header.className = 'org-header';
    const isOpen = currentOrg === orgName;
    header.innerHTML = '<span class="arrow' + (isOpen ? ' open' : '') + '" data-org="' + escHtml(orgName) + '">&#9654;</span>' +
      '<span class="org-name">' + escHtml(orgName) + '</span>' +
      '<span class="org-actions"><span title="Delete org" data-action="delorg">x</span></span>';
    header.addEventListener('click', (e) => {
      if (e.target.dataset.action === 'delorg') { deleteOrg(orgName, e); return; }
      toggleOrg(orgName);
    });
    section.appendChild(header);

    // Ranch list
    const ranchList = document.createElement('div');
    ranchList.className = 'ranch-list' + (isOpen ? ' open' : '');
    ranchList.dataset.org = orgName;

    const ranches = Object.keys(org.ranches || {}).sort((a, b) => {
      const aDone = isRanchComplete(org.ranches[a]);
      const bDone = isRanchComplete(org.ranches[b]);
      if (aDone !== bDone) return aDone ? 1 : -1; // completed go to bottom
      return a.localeCompare(b);
    });
    for (const ranchName of ranches) {
      const ranchData = org.ranches[ranchName];
      const done = isRanchComplete(ranchData);
      const btn = document.createElement('button');
      btn.className = 'ranch-btn' + (currentOrg === orgName && currentRanch === ranchName ? ' active' : '') + (done ? ' ranch-done' : '');
      const count = ranchData.rows.length;
      const doneLabel = done ? ' \u2713' : '';
      btn.innerHTML = '<span class="ranch-name" title="' + escHtml(ranchName) + '">' + escHtml(ranchName) + ' (' + count + ')' + doneLabel + '</span><span class="ranch-del" title="Delete ranch">x</span>';
      btn.addEventListener('click', (e) => {
        if (e.target.classList.contains('ranch-del')) { deleteRanch(orgName, ranchName, e); return; }
        selectRanch(orgName, ranchName);
      });
      btn.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        e.stopPropagation();
        showRanchCtxMenu(e.clientX, e.clientY, orgName, ranchName);
      });
      ranchList.appendChild(btn);
    }

    // Add ranch input
    const addRow = document.createElement('div');
    addRow.className = 'add-row';
    addRow.innerHTML = '<input class="add-ranch-input" data-org="' + escHtml(orgName) + '" placeholder="New ranch..." onkeydown="if(event.key===\'Enter\')createRanch(\'' + escJs(orgName) + '\')">' +
      '<button onclick="createRanch(\'' + escJs(orgName) + '\')">+</button>';
    ranchList.appendChild(addRow);

    section.appendChild(ranchList);
    container.appendChild(section);
  }
}

// --- Main table ---
function renderStats() {
  const bar = document.getElementById('statsBar');
  if (!currentOrg || !currentRanch) { bar.style.display = 'none'; return; }
  const data = getData();
  const ranch = getRanch(data, currentOrg, currentRanch);
  if (!ranch || ranch.rows.length === 0) { bar.style.display = 'none'; return; }

  const total = ranch.rows.length;
  const progress = ranch.progress || {};
  let doneCount = 0;

  ranch.rows.forEach((row, ri) => {
    const required = row.reduce((cols, cell, ci) => {
      if (cell !== '' && !SKIP_FOR_DONE.has(ci)) cols.push(ci);
      return cols;
    }, []);
    const done = new Set(progress[ri] || []);
    if (required.every(ci => done.has(ci))) doneCount++;
  });

  const remaining = total - doneCount;
  const pct = total > 0 ? Math.round((doneCount / total) * 100) : 0;

  bar.style.display = 'block';
  bar.innerHTML =
    '<span class="stat"><span class="stat-value">' + total + '</span> total</span>' +
    '<span class="stat"><span class="stat-value stat-done">' + doneCount + '</span> done</span>' +
    '<span class="stat"><span class="stat-value stat-remaining">' + remaining + '</span> left</span>' +
    '<span class="stat"><span class="stat-value stat-pct">' + pct + '%</span>' +
    '<span class="stat-progress"><span class="stat-progress-fill" style="width:' + pct + '%;"></span></span></span>';
}

function renderHeaders(thead) {
  HEADERS.forEach((h, ci) => {
    const th = document.createElement('th');
    let arrow = '';
    if (sortCol === ci) {
      arrow = '<span class="sort-arrow">' + (sortAsc ? '\u25B2' : '\u25BC') + '</span>';
      th.className = 'sort-active';
    } else {
      arrow = '<span class="sort-arrow">\u25B2</span>';
    }
    th.innerHTML = escHtml(h) + arrow;
    th.addEventListener('click', () => {
      if (sortCol === ci) {
        sortAsc = !sortAsc;
      } else {
        sortCol = ci;
        sortAsc = true;
      }
      renderMain();
    });
    thead.appendChild(th);
  });
}

function renderMain() {
  const thead = document.getElementById('thead');
  const tbody = document.getElementById('tbody');
  const emptyMsg = document.getElementById('emptyMsg');
  const breadcrumb = document.getElementById('breadcrumb');
  const uploadLabel = document.getElementById('uploadLabel');
  const headerBtn = document.getElementById('headerBtn');
  const pasteBtn = document.getElementById('pasteBtn');
  const doneAllBtn = document.getElementById('doneAllBtn');
  const resetBtn = document.getElementById('resetBtn');
  const clearBtn = document.getElementById('clearBtn');

  thead.innerHTML = '';
  tbody.innerHTML = '';
  copiedCells = {};
  document.getElementById('pasteBox').style.display = 'none';
  document.getElementById('headerSetupBox').style.display = 'none';
  document.getElementById('columnMapper').style.display = 'none';

  const hasMapping = currentOrg && getOrgMapping(currentOrg);
  updateMappingStatus();

  if (!currentOrg || !currentRanch) {
    breadcrumb.innerHTML = currentOrg ? '<span class="org-part">' + escHtml(currentOrg) + '</span> — select a ranch' : 'Select a project';
    emptyMsg.style.display = '';
    emptyMsg.textContent = !currentOrg ? 'Create an organization, then add ranches and upload Excel files.' : 'Select a ranch from the sidebar.';
    uploadLabel.style.display = currentOrg ? '' : 'none';
    headerBtn.style.display = currentOrg ? '' : 'none';
    pasteBtn.style.display = (currentOrg && hasMapping) ? '' : 'none';
    doneAllBtn.style.display = 'none';
    resetBtn.style.display = 'none';
    clearBtn.style.display = 'none';
    return;
  }

  breadcrumb.innerHTML = '<span class="org-part">' + escHtml(currentOrg) + ' &rsaquo; </span><span class="ranch-part">' + escHtml(currentRanch) + '</span>';
  uploadLabel.style.display = '';
  headerBtn.style.display = '';
  pasteBtn.style.display = hasMapping ? '' : 'none';
  doneAllBtn.style.display = '';
  resetBtn.style.display = '';
  clearBtn.style.display = '';

  const data = getData();
  const ranch = getRanch(data, currentOrg, currentRanch);
  if (!ranch || ranch.rows.length === 0) {
    emptyMsg.style.display = '';
    emptyMsg.textContent = 'No data yet. Upload an Excel file to add rows.';
    renderHeaders(thead);
    return;
  }

  emptyMsg.style.display = 'none';
  renderHeaders(thead);

  const progress = ranch.progress || {};

  // Auto-mark "None" and "0" cells as done
  let autoMarked = false;
  ranch.rows.forEach((row, ri) => {
    const prog = new Set(progress[ri] || []);
    row.forEach((cell, ci) => {
      if ((cell === 'None' || cell === '0' || cell === '') && !prog.has(ci)) {
        prog.add(ci);
        autoMarked = true;
      }
    });
    progress[ri] = [...prog];
  });
  if (autoMarked) {
    ranch.progress = progress;
    saveData(data);
  }

  // Build sorted index array
  const indices = ranch.rows.map((_, i) => i);
  if (sortCol >= 0) {
    indices.sort((a, b) => {
      let va = ranch.rows[a][sortCol] || '';
      let vb = ranch.rows[b][sortCol] || '';
      if (NUMERIC_COLS.has(sortCol)) {
        va = parseFloat(va) || 0;
        vb = parseFloat(vb) || 0;
        return sortAsc ? va - vb : vb - va;
      }
      va = va.toLowerCase();
      vb = vb.toLowerCase();
      if (va < vb) return sortAsc ? -1 : 1;
      if (va > vb) return sortAsc ? 1 : -1;
      return 0;
    });
  }

  indices.forEach(ri => {
    const row = ranch.rows[ri];
    copiedCells[ri] = new Set(progress[ri] || []);
    const nonEmpty = row.filter((c, i) => c !== "" && !SKIP_FOR_DONE.has(i)).length;

    const tr = document.createElement('tr');
    tr.dataset.ri = ri;

    row.forEach((cell, ci) => {
      const td = document.createElement('td');
      td.textContent = cell;
      td.dataset.ci = ci;
      if (copiedCells[ri].has(ci)) td.classList.add('cell-highlighted');

      let clickCount = 0;
      td.addEventListener('click', () => {
        if (cell === "") return;
        clickCount++;
        if (clickCount === 1) {
          setTimeout(() => {
            if (clickCount === 1) {
              // Single click: highlight & copy
              navigator.clipboard.writeText(cell).then(() => {
                // Remove previous "just clicked" indicator
                const prev = document.querySelector('.cell-just-clicked');
                if (prev) prev.classList.remove('cell-just-clicked');
                td.classList.add('cell-highlighted');
                td.classList.add('cell-just-clicked');
                copiedCells[ri].add(ci);
                saveCurrentProgress();
                checkRowComplete(tr, ri, row, nonEmpty);
              });
            } else {
              // Double click: unhighlight
              td.classList.remove('cell-highlighted');
              td.classList.remove('cell-just-clicked');
              copiedCells[ri].delete(ci);
              saveCurrentProgress();
              checkRowComplete(tr, ri, row, nonEmpty);
            }
            clickCount = 0;
          }, 200);
        }
      });

      tr.appendChild(td);
    });

    // Right-click to move row
    tr.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      showCtxMenu(e.clientX, e.clientY, ri);
    });

    checkRowComplete(tr, ri, row, nonEmpty);
    tbody.appendChild(tr);
  });
  renderStats();
}

function checkRowComplete(tr, ri, row, nonEmpty) {
  const copiedNonEmpty = [...copiedCells[ri]].filter(i => row[i] !== "" && !SKIP_FOR_DONE.has(i)).length;
  const blockCell = tr.querySelectorAll('td')[0];
  if (copiedNonEmpty >= nonEmpty) {
    tr.classList.add('copied');
    if (!blockCell.querySelector('.badge')) {
      const badge = document.createElement('span');
      badge.className = 'badge';
      badge.textContent = '\u2713 Done';
      blockCell.appendChild(badge);
    }
  } else {
    tr.classList.remove('copied');
    const badge = blockCell.querySelector('.badge');
    if (badge) badge.remove();
  }
}

function saveCurrentProgress() {
  if (!currentOrg || !currentRanch) return;
  const data = getData();
  const ranch = getRanch(data, currentOrg, currentRanch);
  if (!ranch) return;
  const obj = {};
  for (const key in copiedCells) obj[key] = [...copiedCells[key]];
  ranch.progress = obj;
  saveData(data);
  renderSidebar();
  renderStats();
}

function markAllDone() {
  if (!currentOrg || !currentRanch) return;
  const data = getData();
  const ranch = getRanch(data, currentOrg, currentRanch);
  if (!ranch || ranch.rows.length === 0) return;
  ranch.rows.forEach((row, ri) => {
    const allCols = [];
    row.forEach((cell, ci) => { if (cell !== '') allCols.push(ci); });
    ranch.progress[ri] = allCols;
  });
  saveData(data);
  renderMain();
}

function resetProgress() {
  if (!currentRanch || !confirm('Reset all highlight progress for "' + currentRanch + '"?')) return;
  const data = getData();
  const ranch = getRanch(data, currentOrg, currentRanch);
  if (!ranch) return;
  ranch.progress = {};
  saveData(data);
  renderMain();
}

function clearRanchData() {
  if (!currentRanch || !confirm('Archive all data from "' + currentRanch + '"?')) return;
  const data = getData();
  const ranch = getRanch(data, currentOrg, currentRanch);
  if (!ranch || ranch.rows.length === 0) return;
  archiveItem('rows', currentRanch, JSON.parse(JSON.stringify(ranch)), currentOrg);
  ranch.rows = [];
  ranch.progress = {};
  saveData(data);
  renderSidebar();
  renderMain();
}

// --- Context menu & move ---
function showCtxMenu(x, y, rowIndex) {
  const menu = document.getElementById('ctxMenu');
  menu.innerHTML = '<div data-action="move">Move to another ranch...</div><div data-action="delete">Delete row</div>';
  menu.style.display = 'block';
  menu.style.left = x + 'px';
  menu.style.top = y + 'px';
  menu.onclick = (e) => {
    const action = e.target.dataset.action;
    menu.style.display = 'none';
    if (action === 'move') openMoveModal(rowIndex);
    else if (action === 'delete') deleteRow(rowIndex);
  };
}

document.addEventListener('click', () => { document.getElementById('ctxMenu').style.display = 'none'; });

function showRanchCtxMenu(x, y, orgName, ranchName) {
  const menu = document.getElementById('ctxMenu');
  menu.innerHTML = '<div data-action="rename">Rename ranch</div><div data-action="merge">Merge into another ranch</div>';
  menu.style.display = 'block';
  menu.style.left = x + 'px';
  menu.style.top = y + 'px';
  menu.onclick = (e) => {
    const action = e.target.dataset.action;
    menu.style.display = 'none';
    if (action === 'rename') renameRanch(orgName, ranchName);
    else if (action === 'merge') mergeRanch(orgName, ranchName);
  };
}

function renameRanch(orgName, ranchName) {
  const newName = prompt('Rename "' + ranchName + '" to:', ranchName);
  if (!newName || newName.trim() === '' || newName.trim() === ranchName) return;
  const trimmed = newName.trim();
  const data = getData();
  const org = data.orgs[orgName];
  if (!org || !org.ranches[ranchName]) return;
  if (org.ranches[trimmed]) { alert('A ranch named "' + trimmed + '" already exists. Use Merge instead.'); return; }
  // Rename: copy data to new key, delete old
  org.ranches[trimmed] = org.ranches[ranchName];
  delete org.ranches[ranchName];
  // Update Sites column in all rows
  org.ranches[trimmed].rows.forEach(row => { row[RANCH_COL] = trimmed; });
  saveData(data);
  if (currentOrg === orgName && currentRanch === ranchName) currentRanch = trimmed;
  renderSidebar();
  renderMain();
}

function mergeRanch(orgName, srcRanchName) {
  const data = getData();
  const org = data.orgs[orgName];
  if (!org) return;
  const otherRanches = Object.keys(org.ranches).filter(r => r !== srcRanchName).sort();
  if (otherRanches.length === 0) { alert('No other ranches to merge into.'); return; }

  // Use the move modal to pick target
  const select = document.getElementById('moveTarget');
  select.innerHTML = '';
  otherRanches.forEach(r => {
    const opt = document.createElement('option');
    opt.value = orgName + '|||' + r;
    opt.textContent = r;
    select.appendChild(opt);
  });

  document.getElementById('moveModal').querySelector('h3').textContent = 'Merge "' + srcRanchName + '" into:';
  document.getElementById('moveModal').classList.add('show');

  // Override confirm to do merge
  const confirmBtn = document.getElementById('moveModal').querySelector('.confirm-btn');
  const newConfirm = confirmBtn.cloneNode(true);
  newConfirm.textContent = 'Merge';
  confirmBtn.parentNode.replaceChild(newConfirm, confirmBtn);
  newConfirm.addEventListener('click', () => {
    const val = select.value;
    if (!val) return;
    const [tOrg, tRanch] = val.split('|||');
    const freshData = getData();
    const src = freshData.orgs[tOrg]?.ranches?.[srcRanchName];
    const dst = freshData.orgs[tOrg]?.ranches?.[tRanch];
    if (!src || !dst) return;

    // Move all rows from src to dst
    src.rows.forEach((row, i) => {
      row[RANCH_COL] = tRanch;
      const newIdx = dst.rows.length;
      dst.rows.push(row);
      dst.progress[newIdx] = src.progress[i] || [];
    });

    // Delete source ranch
    delete freshData.orgs[tOrg].ranches[srcRanchName];
    saveData(freshData);

    if (currentOrg === orgName && currentRanch === srcRanchName) currentRanch = tRanch;
    document.getElementById('moveModal').classList.remove('show');
    // Restore modal title
    document.getElementById('moveModal').querySelector('h3').textContent = 'Move row to another ranch';
    renderSidebar();
    renderMain();
  });
}

function deleteRow(ri) {
  if (!confirm('Archive this row?')) return;
  const data = getData();
  const ranch = getRanch(data, currentOrg, currentRanch);
  if (!ranch) return;
  const removedRow = ranch.rows.splice(ri, 1)[0];
  const removedProgress = ranch.progress[ri] || [];
  archiveItem('row', currentRanch, { row: removedRow, progress: removedProgress }, currentOrg);
  // Rebuild progress indices
  const newProgress = {};
  Object.keys(ranch.progress).forEach(k => {
    const ki = parseInt(k);
    if (ki < ri) newProgress[ki] = ranch.progress[ki];
    else if (ki > ri) newProgress[ki - 1] = ranch.progress[ki];
  });
  ranch.progress = newProgress;
  saveData(data);
  renderSidebar();
  renderMain();
}

function openMoveModal(rowIndex) {
  moveRowIndex = rowIndex;
  const data = getData();
  const select = document.getElementById('moveTarget');
  select.innerHTML = '';

  // List all ranches across all orgs
  for (const orgName of Object.keys(data.orgs).sort()) {
    const org = data.orgs[orgName];
    for (const ranchName of Object.keys(org.ranches).sort()) {
      if (orgName === currentOrg && ranchName === currentRanch) continue;
      const opt = document.createElement('option');
      opt.value = orgName + '|||' + ranchName;
      opt.textContent = orgName + ' > ' + ranchName;
      select.appendChild(opt);
    }
  }

  if (select.options.length === 0) {
    alert('No other ranches to move to. Create one first.');
    return;
  }
  document.getElementById('moveModal').classList.add('show');
}

function closeMoveModal() {
  document.getElementById('moveModal').classList.remove('show');
  moveRowIndex = null;
}

function confirmMoveRow() {
  if (moveRowIndex === null) return;
  const val = document.getElementById('moveTarget').value;
  if (!val) return;
  const [targetOrg, targetRanch] = val.split('|||');

  const data = getData();
  const srcRanch = getRanch(data, currentOrg, currentRanch);
  const dstRanch = getRanch(data, targetOrg, targetRanch);
  if (!srcRanch || !dstRanch) return;

  const row = srcRanch.rows.splice(moveRowIndex, 1)[0];
  // Update ranch column to target ranch name
  row[RANCH_COL] = targetRanch;
  const newIndex = dstRanch.rows.length;
  dstRanch.rows.push(row);
  dstRanch.progress[newIndex] = srcRanch.progress[moveRowIndex] || [];

  // Rebuild source progress
  const newProgress = {};
  Object.keys(srcRanch.progress).forEach(k => {
    const ki = parseInt(k);
    if (ki < moveRowIndex) newProgress[ki] = srcRanch.progress[ki];
    else if (ki > moveRowIndex) newProgress[ki - 1] = srcRanch.progress[ki];
  });
  srcRanch.progress = newProgress;

  saveData(data);
  closeMoveModal();
  renderSidebar();
  renderMain();
}

// --- Excel upload (uses saved mapping if available, otherwise shows mapper) ---
document.getElementById('fileInput').addEventListener('change', function(e) {
  const file = e.target.files[0];
  if (!file || !currentOrg) return;
  const reader = new FileReader();
  reader.onload = function(evt) {
    let sheetRows;
    if (file.name.endsWith('.csv')) {
      const text = evt.target.result;
      sheetRows = text.split('\n').map(line => line.split(',').map(c => c.replace(/^"|"$/g, '').trim()));
    } else {
      const d = new Uint8Array(evt.target.result);
      const wb = XLSX.read(d, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      sheetRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    }
    sheetRows = sheetRows.filter(r => r.some(c => String(c).trim() !== ''));
    if (sheetRows.length < 2) { alert('File has no data rows.'); return; }

    const mapping = getOrgMapping(currentOrg);
    if (mapping) {
      // Use saved mapping, skip header row
      importWithMapping(sheetRows.slice(1), mapping);
    } else {
      // No mapping saved yet — auto-detect from file headers and show mapper
      const fileHeaders = sheetRows[0].map(h => String(h).trim());
      showExcelMapper(sheetRows, fileHeaders);
    }
  };
  if (file.name.endsWith('.csv')) reader.readAsText(file);
  else reader.readAsArrayBuffer(file);
  e.target.value = '';
});

function showExcelMapper(sheetRows, fileHeaders) {
  const mapper = document.getElementById('columnMapper');
  mapper.style.display = '';

  let html = '<div class="column-map"><h4>Map Excel columns to table columns</h4>';
  html += '<p style="color:#666;font-size:11px;margin-bottom:8px;">File headers: ' + fileHeaders.join(', ') + '</p>';

  HEADERS.forEach((h, i) => {
    html += '<div class="map-row"><label>' + h + ':</label><select data-target="' + i + '">';
    html += '<option value="-1">(skip / leave empty)</option>';
    fileHeaders.forEach((fh, fi) => {
      const selected = autoMatch(h, fh) ? ' selected' : '';
      html += '<option value="' + fi + '"' + selected + '>' + fh + '</option>';
    });
    html += '</select></div>';
  });

  html += '<div class="map-actions"><button class="confirm-btn" id="confirmExcelMap">Save Mapping & Import</button><button class="cancel-btn" id="cancelExcelMap">Cancel</button></div></div>';
  mapper.innerHTML = html;

  document.getElementById('confirmExcelMap').addEventListener('click', () => {
    const mapping = [];
    mapper.querySelectorAll('select[data-target]').forEach(sel => mapping.push(parseInt(sel.value)));
    saveOrgMapping(currentOrg, mapping, fileHeaders);
    importWithMapping(sheetRows.slice(1), mapping);
    mapper.style.display = 'none';
    mapper.innerHTML = '';
  });
  document.getElementById('cancelExcelMap').addEventListener('click', () => { mapper.style.display = 'none'; mapper.innerHTML = ''; });
}

function autoMatch(target, source) {
  const t = target.toLowerCase().replace(/[^a-z]/g, '');
  const s = source.toLowerCase().replace(/[^a-z]/g, '');
  if (t === s) return true;
  const aliases = {
    'location': ['location','blockname','block','name','blockid'],
    'sites': ['sites','site','ranch','ranchname','farm'],
    'crop': ['crop','croptype','commodity'],
    'variety': ['variety','varietyname'],
    'locationtype': ['locationtype','type','blocktype','loctype'],
    'planteddate': ['planteddate','planted','dateplanted','plantdate'],
    'acreage': ['acreage','acres','area'],
    'plantcount': ['plantcount','treecount','trees','count','numberoftrees','plants'],
  };
  if (aliases[t]) return aliases[t].includes(s);
  return false;
}

// --- Org column mapping (saved per org) ---
function getOrgMapping(orgName) {
  const data = getData();
  return data.orgs[orgName]?.columnMapping || null;
}

function saveOrgMapping(orgName, mapping, fileHeaders) {
  const data = getData();
  if (!data.orgs[orgName]) return;
  data.orgs[orgName].columnMapping = mapping;
  data.orgs[orgName].fileHeaders = fileHeaders;
  saveData(data);
}

function updateMappingStatus() {
  const el = document.getElementById('mappingStatus');
  if (!currentOrg) { el.innerHTML = ''; return; }
  const data = getData();
  const org = data.orgs[currentOrg];
  if (org?.columnMapping) {
    const fh = org.fileHeaders || [];
    const short = fh.filter(h => h).slice(0, 6).join(', ');
    const extra = fh.filter(h => h).length > 6 ? ' +' + (fh.filter(h=>h).length - 6) + ' more' : '';
    el.innerHTML = 'Mapped: ' + short + extra + ' <a href="#" onclick="clearOrgMapping();return false;">Reset</a>';
  } else {
    el.innerHTML = '<span style="color:var(--amber);">No headers set</span>';
  }
}

function clearOrgMapping() {
  if (!currentOrg || !confirm('Clear saved header mapping for "' + currentOrg + '"?')) return;
  const data = getData();
  if (data.orgs[currentOrg]) {
    delete data.orgs[currentOrg].columnMapping;
    delete data.orgs[currentOrg].fileHeaders;
    saveData(data);
  }
  renderMain();
}

// --- Paste parsing ---
function parsePastedText(text) {
  const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
  if (lines.length === 0) return null;
  return lines.map(line => line.split('\t').map(c => c.trim()));
}

function parsePastedHtml(html) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');
  const rows = doc.querySelectorAll('tr');
  if (rows.length === 0) return null;
  const result = [];
  rows.forEach(tr => {
    const cells = [];
    tr.querySelectorAll('td, th').forEach(cell => cells.push(cell.textContent.trim()));
    if (cells.length > 0) result.push(cells);
  });
  return result.length > 0 ? result : null;
}

function parseClipboard(e) {
  let parsed = null;
  const htmlData = e.clipboardData.getData('text/html');
  if (htmlData) parsed = parsePastedHtml(htmlData);
  if (!parsed) {
    const textData = e.clipboardData.getData('text/plain');
    if (textData) parsed = parsePastedText(textData);
  }
  return parsed;
}

// --- Step 1: Header setup ---
function toggleHeaderSetup() {
  const box = document.getElementById('headerSetupBox');
  box.style.display = box.style.display === 'none' ? '' : 'none';
  if (box.style.display !== 'none') document.getElementById('headerArea').focus();
}

function processHeaderPaste() {
  const text = document.getElementById('headerArea').value;
  if (!text.trim()) { alert('Paste the header row first.'); return; }
  if (!currentOrg) { alert('Select an organization first.'); return; }
  const parsed = parsePastedText(text);
  if (!parsed || parsed.length === 0) { alert('Could not parse headers.'); return; }
  // Take only the first row as headers
  const fileHeaders = parsed[0].map(h => String(h).trim());
  document.getElementById('headerSetupBox').style.display = 'none';
  document.getElementById('headerArea').value = '';
  showHeaderMapper(fileHeaders);
}

function showHeaderMapper(fileHeaders) {
  const mapper = document.getElementById('columnMapper');
  mapper.style.display = '';

  let html = '<div class="column-map"><h4>Map your spreadsheet columns to table columns</h4>';
  html += '<p style="color:#666;font-size:11px;margin-bottom:8px;">Your headers: ' + fileHeaders.join(', ') + '</p>';

  HEADERS.forEach((h, i) => {
    html += '<div class="map-row"><label>' + h + ':</label><select data-target="' + i + '">';
    html += '<option value="-1">(skip / leave empty)</option>';
    fileHeaders.forEach((fh, fi) => {
      const selected = autoMatch(h, fh) ? ' selected' : '';
      html += '<option value="' + fi + '"' + selected + '>' + fh + '</option>';
    });
    html += '</select></div>';
  });

  html += '<div class="map-actions"><button class="confirm-btn" id="confirmHeaderMap">Save Mapping</button><button class="cancel-btn" id="cancelHeaderMap">Cancel</button></div></div>';
  mapper.innerHTML = html;

  document.getElementById('confirmHeaderMap').addEventListener('click', () => {
    const mapping = [];
    mapper.querySelectorAll('select[data-target]').forEach(sel => mapping.push(parseInt(sel.value)));
    saveOrgMapping(currentOrg, mapping, fileHeaders);
    mapper.style.display = 'none';
    mapper.innerHTML = '';
    renderMain();
    alert('Header mapping saved for "' + currentOrg + '". You can now paste data rows.');
  });
  document.getElementById('cancelHeaderMap').addEventListener('click', () => { mapper.style.display = 'none'; mapper.innerHTML = ''; });
}

// --- Step 2: Data paste ---
function togglePasteBox() {
  const box = document.getElementById('pasteBox');
  box.style.display = box.style.display === 'none' ? '' : 'none';
  if (box.style.display !== 'none') document.getElementById('pasteArea').focus();
}

function processPasteBox() {
  const text = document.getElementById('pasteArea').value;
  if (!text.trim()) { alert('Nothing to import. Paste data first.'); return; }
  if (!currentOrg) { alert('Select an organization first.'); return; }
  const mapping = getOrgMapping(currentOrg);
  if (!mapping) { alert('Set up the header mapping first (click "Set Headers").'); return; }
  const parsed = parsePastedText(text);
  if (!parsed || parsed.length === 0) { alert('Could not parse pasted data.'); return; }
  document.getElementById('pasteBox').style.display = 'none';
  document.getElementById('pasteArea').value = '';
  importWithMapping(parsed, mapping);
}

let pendingImportRows = null;

function importWithMapping(dataRows, mapping) {
  const newRows = dataRows.map(sr => {
    return HEADERS.map((_, hi) => {
      const fi = mapping[hi];
      if (fi < 0 || fi >= sr.length) return '';
      return String(sr[fi]).trim();
    });
  }).filter(r => r.some(c => c !== ''));

  if (newRows.length === 0) { alert('No rows to import.'); return; }

  // Safety check: Acreage must always be less than Plant Count. Auto-swap if flipped.
  let swapCount = 0;
  newRows.forEach(row => {
    const acreage = parseFloat(row[ACREAGE_COL]);
    const plantCount = parseFloat(row[PLANTCOUNT_COL]);
    if (!isNaN(acreage) && !isNaN(plantCount) && acreage > 0 && plantCount > 0 && acreage > plantCount) {
      const tmp = row[ACREAGE_COL];
      row[ACREAGE_COL] = row[PLANTCOUNT_COL];
      row[PLANTCOUNT_COL] = tmp;
      swapCount++;
    }
  });
  if (swapCount > 0) {
    alert('Warning: ' + swapCount + ' row(s) had Acreage greater than Plant Count. These were auto-swapped.');
  }

  // Group by ranch column
  const grouped = {};
  newRows.forEach(row => {
    const ranchName = row[RANCH_COL] || 'Unknown Ranch';
    if (!grouped[ranchName]) grouped[ranchName] = [];
    grouped[ranchName].push(row);
  });

  pendingImportRows = grouped;
  showImportDestModal(grouped);
}

function showImportDestModal(grouped) {
  const data = getData();
  const org = data.orgs[currentOrg];
  const existingRanches = Object.keys(org?.ranches || {}).sort();
  const incomingRanches = Object.keys(grouped).sort();
  // Combine all unique ranch names (existing + incoming)
  const allTargets = [...new Set([...existingRanches, ...incomingRanches])].sort();
  const list = document.getElementById('importDestList');
  list.innerHTML = '';

  for (const srcName of incomingRanches) {
    const row = document.createElement('div');
    row.className = 'import-dest-row';

    const label = document.createElement('div');
    label.innerHTML = '<div class="src-name">' + escHtml(srcName) + '</div><div class="src-count">' + grouped[srcName].length + ' row(s)</div>';

    const select = document.createElement('select');
    select.dataset.src = srcName;

    // Option: keep as its own ranch (with its original name)
    const keepOpt = document.createElement('option');
    keepOpt.value = '__keep__';
    keepOpt.textContent = 'Keep as "' + srcName + '"';
    select.appendChild(keepOpt);

    // Options: merge into any existing or other incoming ranch
    // Add new (incoming) options first, highlighted
    allTargets.forEach(r => {
      if (r === srcName) return;
      const isExisting = existingRanches.includes(r);
      if (isExisting) return;
      const opt = document.createElement('option');
      opt.value = r;
      opt.textContent = 'Merge into "' + r + '" (new)';
      opt.className = 'option-new';
      opt.style.backgroundColor = '#e8f5e9';
      opt.style.fontWeight = 'bold';
      select.appendChild(opt);
    });
    // Then existing options
    allTargets.forEach(r => {
      if (r === srcName) return;
      const isExisting = existingRanches.includes(r);
      if (!isExisting) return;
      const opt = document.createElement('option');
      opt.value = r;
      opt.textContent = 'Merge into "' + r + '"';
      select.appendChild(opt);
    });

    row.appendChild(label);
    row.appendChild(select);
    list.appendChild(row);
  }

  document.getElementById('importDestModal').classList.add('show');
}

function cancelImportDest() {
  document.getElementById('importDestModal').classList.remove('show');
  pendingImportRows = null;
}

function confirmImportDest() {
  if (!pendingImportRows) return;
  const data = getData();
  if (!data.orgs[currentOrg]) return;

  const selects = document.querySelectorAll('#importDestList select');
  let totalImported = 0;
  let firstNewRanch = null;

  selects.forEach(sel => {
    const srcName = sel.dataset.src;
    const target = sel.value;
    const rows = pendingImportRows[srcName];
    if (!rows || rows.length === 0) return;

    const destName = (target === '__keep__') ? srcName : target;
    const isMerge = (target !== '__keep__'); // rows redirected to a different ranch

    // Update Sites column to destination name
    rows.forEach(row => { row[RANCH_COL] = destName; });

    if (!data.orgs[currentOrg].ranches[destName]) {
      data.orgs[currentOrg].ranches[destName] = { rows: [], progress: {} };
    }
    const ranch = data.orgs[currentOrg].ranches[destName];

    const startIdx = ranch.rows.length;
    rows.forEach((row, i) => {
      ranch.rows.push(row);
      if (isMerge) {
        // Merged from another source — needs work, start fresh
        ranch.progress[startIdx + i] = [];
      } else {
        // Kept as own ranch — mark as done
        const allCols = [];
        row.forEach((cell, ci) => { if (cell !== '') allCols.push(ci); });
        ranch.progress[startIdx + i] = allCols;
      }
    });
    totalImported += rows.length;
    // Track the first ranch that received new (non-done) rows
    if (!firstNewRanch && isMerge) firstNewRanch = destName;
  });

  saveData(data);
  pendingImportRows = null;
  document.getElementById('importDestModal').classList.remove('show');

  // Navigate to the first ranch with new rows
  if (firstNewRanch) {
    currentRanch = firstNewRanch;
  } else if (!currentRanch) {
    const ranches = Object.keys(data.orgs[currentOrg].ranches).sort();
    if (ranches.length > 0) currentRanch = ranches[0];
  }
  sortCol = -1;
  sortAsc = true;
  renderSidebar();
  renderMain();
  alert('Imported ' + totalImported + ' rows.');
}

// Ctrl+V paste listener — imports data rows directly using saved mapping
document.addEventListener('paste', function(e) {
  const tag = (e.target.tagName || '').toLowerCase();
  if (tag === 'input' || tag === 'textarea' || tag === 'select') return;
  if (!currentOrg) return;

  const mapping = getOrgMapping(currentOrg);
  if (!mapping) { return; } // No mapping set, ignore Ctrl+V on page

  e.preventDefault();
  const parsed = parseClipboard(e);
  if (!parsed || parsed.length === 0) {
    alert('Could not parse pasted data.');
    return;
  }
  importWithMapping(parsed, mapping);
});

// --- Helpers ---
function escHtml(s) { return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function escJs(s) { return s.replace(/\\/g,'\\\\').replace(/'/g,"\\'"); }

// --- Fix existing data: swap acreage/plant count if flipped ---
function fixExistingData() {
  const data = getData();
  let changed = false;
  for (const orgName of Object.keys(data.orgs)) {
    const org = data.orgs[orgName];
    for (const ranchName of Object.keys(org.ranches || {})) {
      const ranch = org.ranches[ranchName];
      (ranch.rows || []).forEach(row => {
        // Strip trailing/leading spaces from all cells
        for (let i = 0; i < row.length; i++) {
          const trimmed = String(row[i]).trim();
          if (trimmed !== row[i]) { row[i] = trimmed; changed = true; }
        }
        // Swap acreage/plant count if flipped
        const acreage = parseFloat(row[ACREAGE_COL]);
        const plantCount = parseFloat(row[PLANTCOUNT_COL]);
        if (!isNaN(acreage) && !isNaN(plantCount) && acreage > 0 && plantCount > 0 && acreage > plantCount) {
          const tmp = row[ACREAGE_COL];
          row[ACREAGE_COL] = row[PLANTCOUNT_COL];
          row[PLANTCOUNT_COL] = tmp;
          changed = true;
        }
      });
    }
  }
  if (changed) saveData(data);
}

// --- Sidebar resize ---
(function() {
  const sidebar = document.getElementById('sidebar');
  const handle = document.getElementById('sidebarResize');
  let dragging = false;

  // Restore saved width
  const savedWidth = localStorage.getItem('sidebarWidth');
  if (savedWidth) sidebar.style.width = savedWidth + 'px';

  handle.addEventListener('mousedown', (e) => {
    e.preventDefault();
    dragging = true;
    handle.classList.add('dragging');
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
  });

  document.addEventListener('mousemove', (e) => {
    if (!dragging) return;
    const newWidth = Math.min(500, Math.max(160, e.clientX));
    sidebar.style.width = newWidth + 'px';
  });

  document.addEventListener('mouseup', () => {
    if (!dragging) return;
    dragging = false;
    handle.classList.remove('dragging');
    document.body.style.cursor = '';
    document.body.style.userSelect = '';
    // Save width
    localStorage.setItem('sidebarWidth', parseInt(sidebar.style.width));
  });
})();

// --- File save/load with autosave ---
let fileHandle = null;
let autosaveInterval = null;

function updateSaveStatus(msg, type) {
  const el = document.getElementById('saveStatus');
  el.textContent = msg;
  el.className = 'save-status' + (type ? ' ' + type : '');
}

async function manualSave() {
  try {
    if (!fileHandle) {
      // First time: prompt user to pick save location
      fileHandle = await window.showSaveFilePicker({
        suggestedName: 'table-progress.json',
        startIn: 'downloads',
        types: [{ description: 'JSON', accept: { 'application/json': ['.json'] } }]
      });
      startAutosave();
    }
    await writeToFile();
  } catch (e) {
    if (e.name !== 'AbortError') {
      updateSaveStatus('Save failed: ' + e.message, 'error');
      // Fallback: download as file
      fallbackDownload();
    }
  }
}

async function writeToFile() {
  const saveData = {
    data: getData(),
    archive: getArchive(),
    savedAt: new Date().toISOString()
  };
  if (fileHandle) {
    try {
      const writable = await fileHandle.createWritable();
      await writable.write(JSON.stringify(saveData, null, 2));
      await writable.close();
      const time = new Date().toLocaleTimeString();
      updateSaveStatus('Saved ' + time, 'saved');
    } catch (e) {
      updateSaveStatus('Autosave failed', 'error');
    }
  } else {
    fallbackDownload();
  }
}

function fallbackDownload() {
  const saveData = {
    data: getData(),
    archive: getArchive(),
    savedAt: new Date().toISOString()
  };
  const blob = new Blob([JSON.stringify(saveData, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'table-progress.json';
  a.click();
  URL.revokeObjectURL(a.href);
  const time = new Date().toLocaleTimeString();
  updateSaveStatus('Downloaded ' + time, 'saved');
}

function startAutosave() {
  if (autosaveInterval) clearInterval(autosaveInterval);
  autosaveInterval = setInterval(() => {
    writeToFile();
  }, 30000); // Autosave every 30 seconds
  updateSaveStatus('Autosave active', 'saved');
}

// Load from file
document.getElementById('loadFileInput').addEventListener('change', function(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(evt) {
    try {
      const loaded = JSON.parse(evt.target.result);
      if (loaded.data) {
        localStorage.setItem(DATA_KEY, JSON.stringify(loaded.data));
      }
      if (loaded.archive) {
        localStorage.setItem(ARCHIVE_KEY, JSON.stringify(loaded.archive));
      }
      updateSaveStatus('Loaded from file', 'saved');
      currentOrg = null;
      currentRanch = null;
      // Re-init
      const initData = getData();
      const orgNames = Object.keys(initData.orgs).sort();
      if (orgNames.length > 0) {
        currentOrg = orgNames[0];
        const ranches = Object.keys(initData.orgs[initData.orgs[currentOrg] ? currentOrg : orgNames[0]].ranches).sort();
        if (ranches.length > 0) currentRanch = ranches[0];
      }
      renderSidebar();
      renderMain();
      alert('Data loaded successfully.');
    } catch (err) {
      alert('Failed to load file: ' + err.message);
    }
  };
  reader.readAsText(file);
  e.target.value = '';
});

// --- Init ---
migrateOldData();
fixExistingData();
renderSidebar();
// Auto-select first org/ranch
const initData = getData();
const orgNames = Object.keys(initData.orgs).sort();
if (orgNames.length > 0 && !currentOrg) {
  currentOrg = orgNames[0];
  const ranches = Object.keys(initData.orgs[currentOrg].ranches).sort();
  if (ranches.length > 0) currentRanch = ranches[0];
  renderSidebar();
}
renderMain();
