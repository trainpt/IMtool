// ═══════════════════════════════════════════
// Session preferences — IN-MEMORY ONLY
//
// Nothing this app holds is written to the browser: no localStorage, no
// IndexedDB, no cookies. Everything lives for the life of the page and is gone
// on reload. Each tool takes its own file upload and keeps its own data —
// there is no shared store, no imported-sheet library and no projects.
//
// All that's left is a Web Storage-shaped shim for small per-session
// preferences (OCR engine, Vision API key, active tab, per-tool toggles). It
// keeps the same method signatures as localStorage so the modules that used to
// call it kept their existing call sites and try/catch fallbacks.
// ═══════════════════════════════════════════

const MEMORY_PREFS = new Map();

const sessionPrefs = {
  getItem(k) { return MEMORY_PREFS.has(k) ? MEMORY_PREFS.get(k) : null; },
  setItem(k, v) { MEMORY_PREFS.set(String(k), String(v)); },
  removeItem(k) { MEMORY_PREFS.delete(k); }
};
