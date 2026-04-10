// ═══════════════════════════════════════════
// Unified Data Store
// All app data in one localStorage key
// ═══════════════════════════════════════════

const STORE_KEY = 'imToolsData_v2';

function getStore() {
  try {
    return JSON.parse(localStorage.getItem(STORE_KEY)) || createEmptyStore();
  } catch(e) {
    return createEmptyStore();
  }
}

function saveStore(data) {
  localStorage.setItem(STORE_KEY, JSON.stringify(data));
}

function createEmptyStore() {
  return {
    orgs: {},        // { "OrgName": { ranches: {}, columnMapping: null, fileHeaders: null } }
    sheets: {},      // { "SheetName": { headers, rows, orgName, fileName, rowCount, importedAt } }
    ocrSaves: {},    // { "OrgName": [{ name, source, headers, rows, doneCells, archived, savedAt }] }
    archive: [],     // [{ type, name, content, parentOrg, archivedAt }]
    settings: {}     // App settings
  };
}

// ─── Migration from old stores ───
function migrateOldStores() {
  const store = getStore();
  let migrated = false;

  // Migrate from tableAppData
  const oldMerge = localStorage.getItem('tableAppData');
  if (oldMerge && !store._mergedTableApp) {
    try {
      const old = JSON.parse(oldMerge);
      if (old.orgs) {
        Object.keys(old.orgs).forEach(orgName => {
          if (!store.orgs[orgName]) store.orgs[orgName] = old.orgs[orgName];
        });
      }
      store._mergedTableApp = true;
      migrated = true;
    } catch(e) {}
  }

  // Migrate from imGlobalSheets
  const oldSheets = localStorage.getItem('imGlobalSheets');
  if (oldSheets && !store._mergedGlobalSheets) {
    try {
      const old = JSON.parse(oldSheets);
      if (old.sheets) {
        Object.keys(old.sheets).forEach(name => {
          if (!store.sheets[name]) store.sheets[name] = old.sheets[name];
        });
      }
      store._mergedGlobalSheets = true;
      migrated = true;
    } catch(e) {}
  }

  // Migrate from orgDataStore
  const oldOCR = localStorage.getItem('orgDataStore');
  if (oldOCR && !store._mergedOcrSaves) {
    try {
      const old = JSON.parse(oldOCR);
      Object.keys(old).forEach(orgName => {
        if (!store.ocrSaves[orgName]) store.ocrSaves[orgName] = old[orgName];
      });
      store._mergedOcrSaves = true;
      migrated = true;
    } catch(e) {}
  }

  // Migrate archive
  const oldArchive = localStorage.getItem('tableAppArchive');
  if (oldArchive && !store._mergedArchive) {
    try {
      store.archive = JSON.parse(oldArchive);
      store._mergedArchive = true;
      migrated = true;
    } catch(e) {}
  }

  if (migrated) saveStore(store);
}

// ─── Convenience accessors ───

// Orgs
function getOrgs() { return getStore().orgs; }
function getOrg(name) { return getStore().orgs[name]; }
function getRanch(orgName, ranchName) { return getStore().orgs[orgName]?.ranches?.[ranchName]; }

function getAllOrgNames() {
  const store = getStore();
  const names = new Set();
  Object.keys(store.orgs || {}).forEach(n => names.add(n));
  Object.keys(store.ocrSaves || {}).forEach(n => names.add(n));
  Object.values(store.sheets || {}).forEach(s => { if (s.orgName) names.add(s.orgName); });
  return [...names].sort();
}

function createOrg(name) {
  const store = getStore();
  if (store.orgs[name]) return false;
  store.orgs[name] = { ranches: {} };
  saveStore(store);
  return true;
}

// Sheets (imported)
function getSheet(name) { return getStore().sheets[name]; }
function getSheetNames() { return Object.keys(getStore().sheets).sort(); }

function saveSheet(name, headers, rows, fileName, orgName) {
  const store = getStore();
  store.sheets[name] = {
    headers, rows: rows.map(r => r.map(c => String(c).trim())),
    fileName, orgName: orgName || '', rowCount: rows.length,
    importedAt: new Date().toISOString()
  };
  // Ensure org exists
  if (orgName && !store.orgs[orgName]) store.orgs[orgName] = { ranches: {} };
  saveStore(store);
}

function deleteSheet(name) {
  const store = getStore();
  delete store.sheets[name];
  saveStore(store);
}

function getSheetsByOrg(orgName) {
  const store = getStore();
  return Object.keys(store.sheets).filter(n => store.sheets[n].orgName === orgName).sort();
}

// OCR saves
function getOcrSaves(orgName) { return getStore().ocrSaves[orgName] || []; }

function saveOcrData(orgName, item) {
  const store = getStore();
  if (!store.ocrSaves[orgName]) store.ocrSaves[orgName] = [];
  // Check for replacement
  const existingIdx = store.ocrSaves[orgName].findIndex(i => i.name === item.name);
  if (existingIdx >= 0) {
    // Transfer done progress for matching rows
    const existing = store.ocrSaves[orgName][existingIdx];
    const oldStatus = {};
    (existing.rows || []).forEach((r, ri) => {
      oldStatus[r.join('\t')] = { doneCells: (existing.doneCells?.[ri]) || [], archived: (existing.archived?.[ri]) || [] };
    });
    item.doneCells = item.rows.map((r, ri) => {
      const key = r.join('\t');
      return oldStatus[key] ? oldStatus[key].doneCells : (item.doneCells?.[ri] || []);
    });
    item.archived = item.rows.map((r, ri) => {
      const key = r.join('\t');
      return oldStatus[key] ? oldStatus[key].archived : (item.archived?.[ri] || []);
    });
    store.ocrSaves[orgName][existingIdx] = item;
  } else {
    store.ocrSaves[orgName].push(item);
  }
  saveStore(store);
}

// Archive
function getArchive() { return getStore().archive || []; }

function archiveItem(type, name, content, parentOrg) {
  const store = getStore();
  store.archive.push({ type, name, content: JSON.parse(JSON.stringify(content)), parentOrg, archivedAt: new Date().toISOString() });
  saveStore(store);
}

// Init
migrateOldStores();
