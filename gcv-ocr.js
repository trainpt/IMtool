// ═══ Google Cloud Vision OCR adapter ═══
// Returns the same shape as Tesseract.recognize() so callers don't branch on shape:
//   { data: { text, words: [{text, confidence, bbox:{x0,y0,x1,y1}}], lines, confidence } }
// Free tier: 1000 DOCUMENT_TEXT_DETECTION calls/month. Key sits in localStorage and
// rides the URL of every request — the user MUST restrict the key in Google Cloud
// Console to (a) Cloud Vision API only, (b) referrer matching the file's origin.

(function () {
  'use strict';

  // ─── Key storage ───
  function gcvGetApiKey() {
    return (localStorage.getItem('gcvApiKey') || '').trim();
  }
  function gcvSetApiKey(k) {
    if (k && k.trim()) localStorage.setItem('gcvApiKey', k.trim());
    else localStorage.removeItem('gcvApiKey');
  }
  function gcvClearApiKey() { localStorage.removeItem('gcvApiKey'); }

  // ─── Monthly usage counter (local-only — Google's bill is authoritative) ───
  function _currentMonth() {
    const d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  }
  function gcvGetMonthlyCount() {
    try {
      const raw = localStorage.getItem('gcvUsage');
      if (!raw) return 0;
      const obj = JSON.parse(raw);
      if (!obj || obj.month !== _currentMonth()) return 0;
      return obj.count || 0;
    } catch (e) { return 0; }
  }
  function gcvBumpCount() {
    const month = _currentMonth();
    let count = gcvGetMonthlyCount() + 1;
    localStorage.setItem('gcvUsage', JSON.stringify({ month: month, count: count }));
    _refreshUsageUI();
    return count;
  }
  function _refreshUsageUI() {
    const el = document.getElementById('gcvUsage');
    if (el) el.textContent = 'GCV calls this month: ' + gcvGetMonthlyCount() + ' / 1000';
  }

  // ─── Save-button handler ───
  function gcvSaveKey() {
    const input = document.getElementById('gcvApiKeyInput');
    if (!input) return;
    gcvSetApiKey(input.value);
    input.value = gcvGetApiKey() ? '••••••••' + gcvGetApiKey().slice(-4) : '';
    const status = document.getElementById('gcvKeyStatus');
    if (status) {
      status.textContent = gcvGetApiKey() ? 'Saved.' : 'Cleared.';
      setTimeout(() => { if (status) status.textContent = ''; }, 2000);
    }
    _refreshUsageUI();
  }

  function gcvHydrateKeyInput() {
    const input = document.getElementById('gcvApiKeyInput');
    if (!input) return;
    const k = gcvGetApiKey();
    if (k) input.value = '••••••••' + k.slice(-4);
    _refreshUsageUI();
  }

  // ─── Main adapter ───
  async function gcvRecognize(dataUrl, onProgress) {
    const key = gcvGetApiKey();
    if (!key) throw new Error('NO_GCV_KEY');
    if (onProgress) onProgress({ status: 'encoding', progress: 0.1 });

    const b64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');
    if (b64.length > 7_000_000) {
      throw new Error('Image too large for GCV (>7MB base64). Crop or downscale first.');
    }

    if (onProgress) onProgress({ status: 'uploading', progress: 0.3 });

    let resp;
    try {
      resp = await fetch(
        'https://vision.googleapis.com/v1/images:annotate?key=' + encodeURIComponent(key),
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            requests: [{
              image: { content: b64 },
              features: [{ type: 'DOCUMENT_TEXT_DETECTION' }]
            }]
          })
        }
      );
    } catch (netErr) {
      throw new Error('GCV network/CORS error: ' + (netErr && netErr.message ? netErr.message : netErr));
    }

    if (!resp.ok) {
      let body = '';
      try { body = await resp.text(); } catch (e) { /* ignore */ }
      throw new Error('GCV HTTP ' + resp.status + ': ' + body.slice(0, 400));
    }

    if (onProgress) onProgress({ status: 'parsing', progress: 0.85 });

    const json = await resp.json();
    if (json.responses && json.responses[0] && json.responses[0].error) {
      const e = json.responses[0].error;
      throw new Error('GCV API error: ' + (e.message || JSON.stringify(e)));
    }

    gcvBumpCount();

    const fta = (json.responses && json.responses[0] && json.responses[0].fullTextAnnotation) || {};
    const text = fta.text || '';

    const words = [];
    const pages = fta.pages || [];
    for (let p = 0; p < pages.length; p++) {
      const blocks = pages[p].blocks || [];
      for (let b = 0; b < blocks.length; b++) {
        const paras = blocks[b].paragraphs || [];
        for (let pa = 0; pa < paras.length; pa++) {
          const ws = paras[pa].words || [];
          for (let w = 0; w < ws.length; w++) {
            const word = ws[w];
            const syms = word.symbols || [];
            let t = '';
            for (let s = 0; s < syms.length; s++) {
              t += (syms[s].text || '');
              // GCV puts the detected break (space, newline) on the symbol.
              const brk = syms[s].property && syms[s].property.detectedBreak;
              if (brk && (brk.type === 'SPACE' || brk.type === 'SURE_SPACE')) {
                // Trailing space is implicit in word separation; we keep words atomic.
                // No-op.
              }
            }
            const vs = (word.boundingBox && word.boundingBox.vertices) || [];
            if (!vs.length || !t) continue;
            const xs = vs.map(v => v.x || 0);
            const ys = vs.map(v => v.y || 0);
            const conf = (typeof word.confidence === 'number') ? word.confidence : 0.95;
            words.push({
              text: t,
              confidence: conf * 100,
              bbox: {
                x0: Math.min.apply(null, xs),
                y0: Math.min.apply(null, ys),
                x1: Math.max.apply(null, xs),
                y1: Math.max.apply(null, ys)
              }
            });
          }
        }
      }
    }

    const avgConf = words.length
      ? words.reduce((a, w) => a + w.confidence, 0) / words.length
      : 0;

    if (onProgress) onProgress({ status: 'done', progress: 1.0 });

    return { data: { text: text, words: words, lines: [], confidence: avgConf } };
  }

  // ─── Expose on window ───
  window.gcvGetApiKey = gcvGetApiKey;
  window.gcvSetApiKey = gcvSetApiKey;
  window.gcvClearApiKey = gcvClearApiKey;
  window.gcvGetMonthlyCount = gcvGetMonthlyCount;
  window.gcvBumpCount = gcvBumpCount;
  window.gcvSaveKey = gcvSaveKey;
  window.gcvHydrateKeyInput = gcvHydrateKeyInput;
  window.gcvRecognize = gcvRecognize;
})();
