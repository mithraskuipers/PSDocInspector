let state = {
  results: [],
  skipped: [],
  sourcePath: '',
  scannedFiles: 0,
  totalFiles: 0,
  expandedRow: null,
  ocrSelected: new Set(), // fullPaths of skipped, OCR-eligible PDFs currently checked
  ocrRunning: false,
  ocrAvailable: null, // null = not checked yet, true/false once /api/ocr-status responds
  ocrUnavailableReason: '',
  // The keywords/options that actually produced the current results, saved
  // at scan time (or restored from an imported file) so exports can record
  // exactly what was searched for, independent of whatever is currently
  // sitting in the form fields.
  searchSettings: { keywords: [], recursive: true, caseSensitive: false, wholeWord: false, useRegex: false }
};

const $ = (id) => document.getElementById(id);

function setStatus(text, isError) {
  const el = $('status');
  el.textContent = text || '';
  el.classList.toggle('error', !!isError);
}

function parseKeywords(raw) {
  return raw.split(/[\n,]/).map(s => s.trim()).filter(Boolean);
}

// ---------- Stop server ----------

$('stopServerBtn').addEventListener('click', async () => {
  if (!confirm('Stop the DocInspector server? This will close the connection and the app will stop working until it is started again.')) return;
  setStatus('Stopping server...');
  $('stopServerBtn').disabled = true;
  try {
    await fetch('/api/shutdown', { method: 'POST' });
  } catch (e) {
    // The server closes the connection as part of shutting down, so a
    // failed/aborted fetch here is the expected outcome, not an error.
  }
  setStatus('Server stopped. You can close this tab.');
  document.body.style.opacity = '0.5';
  document.body.style.pointerEvents = 'none';
});

// ---------- Browse ----------

$('browseBtn').addEventListener('click', async () => {
  setStatus('Waiting for folder selection...');
  try {
    const res = await fetch('/api/browse');
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Server returned an error');
    if (data.path) {
      $('folderPath').value = data.path;
      setStatus('');
    } else {
      setStatus('No folder selected.');
    }
  } catch (e) {
    setStatus('Could not open folder dialog: ' + e.message, true);
  }
});

// ---------- Scan ----------

// ---------- Progress bar ----------

function showProgress(visible) {
  $('progressPanel').style.display = visible ? 'block' : 'none';
}

function setProgress(processed, total, currentFile, findingsCount, skippedCount) {
  const pct = total > 0 ? Math.round((processed / total) * 100) : 0;
  $('progressBarFill').style.width = pct + '%';
  $('progressPercent').textContent = pct + '%';
  $('progressLabel').textContent = total > 0
    ? `${processed} of ${total} file(s) scanned${currentFile ? ' \u2014 ' + currentFile : ''}`
    : 'Preparing scan...';
  const findingsEl = $('progressFindings');
  findingsEl.textContent = typeof findingsCount === 'number'
    ? `${findingsCount} finding${findingsCount === 1 ? '' : 's'} so far`
    : '';
  const skippedEl = $('progressSkipped');
  skippedEl.textContent = typeof skippedCount === 'number' && skippedCount > 0
    ? `${skippedCount} skipped/failed`
    : '';
}

// Clears every trace of a previous scan (results, skipped files, filters,
// and their panels) so starting a new scan doesn't leave stale rows visible
// underneath the progress bar while the new one runs.
function resetScanUI() {
  state.results = [];
  state.skipped = [];
  state.sourcePath = '';
  state.scannedFiles = 0;
  state.totalFiles = 0;
  state.expandedRow = null;
  state.ocrSelected = new Set();
  state.searchSettings = { keywords: [], recursive: true, caseSensitive: false, wholeWord: false, useRegex: false };

  $('filterKeyword').value = '';
  $('filterDir').value = '';
  $('filterFile').value = '';
  $('filterPanel').style.display = 'none';

  $('resultsPanel').style.display = 'none';
  $('resultsBody').innerHTML = '';
  $('emptyState').style.display = 'none';

  $('skippedPanel').style.display = 'none';
  $('skippedBody').innerHTML = '';

  document.querySelectorAll('.detail-row').forEach(el => el.remove());

  setProgress(0, 0, '', 0, 0);
}

$('scanBtn').addEventListener('click', async () => {
  const path = $('folderPath').value.trim();
  const keywords = parseKeywords($('keywords').value);

  if (!path) { setStatus('Enter or browse to a folder first.', true); return; }
  if (keywords.length === 0) { setStatus('Enter at least one keyword.', true); return; }

  const recursive = $('recursive').checked;
  const caseSensitive = $('caseSensitive').checked;
  const wholeWord = $('wholeWord').checked;
  const useRegex = $('useRegex').checked;

  resetScanUI();
  state.searchSettings = { keywords, recursive, caseSensitive, wholeWord, useRegex };
  $('scanBtn').disabled = true;
  showProgress(true);
  setStatus('Scanning...');

  try {
    const res = await fetch('/api/scan', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ path, recursive, keywords, caseSensitive, wholeWord, useRegex })
    });

    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error || 'Scan failed');
    }

    // The server streams newline-delimited JSON (NDJSON): a 'start' line,
    // one 'progress' line per file processed (including running findings and
    // skipped counts), periodic 'partial' lines carrying newly-found rows
    // (so results/skipped files show up live instead of only at the end -
    // useful since a full scan can take a while), and a final 'done' line
    // carrying the authoritative full result set.
    const data = await readNdjsonScanResponse(
      res,
      (processed, total, currentFile, findingsCount, skippedCount) => {
        setProgress(processed, total, currentFile, findingsCount, skippedCount);
      },
      (newResults, newSkipped) => {
        if (newResults.length) state.results.push(...newResults);
        if (newSkipped.length) state.skipped.push(...newSkipped);
        if (newResults.length || newSkipped.length) {
          populateFilterOptions();
          renderResults();
          renderSkipped();
        }
      }
    );

    if (!data) throw new Error('Scan did not complete - no results received');

    // 'done' carries the authoritative full arrays - replace (not append)
    // whatever the partial updates already rendered, so counts can't drift.
    state.results = Array.isArray(data.results) ? data.results : (data.results ? [data.results] : []);
    state.skipped = Array.isArray(data.skippedFiles) ? data.skippedFiles : (data.skippedFiles ? [data.skippedFiles] : []);
    state.sourcePath = data.sourcePath || path;
    state.scannedFiles = data.scannedFiles || 0;
    state.totalFiles = data.totalFiles || 0;
    state.ocrSelected = new Set(state.skipped.filter(isOcrCandidate).map(s => s.fullPath));

    let msg = `Scanned ${data.scannedFiles} of ${data.totalFiles} file(s), ${state.results.length} match row(s).`;
    if (state.skipped.length) msg += ` ${state.skipped.length} file(s) skipped/failed - see below.`;
    setStatus(msg);

    populateFilterOptions();
    renderResults();
    renderSkipped();
  } catch (e) {
    setStatus('Error: ' + e.message, true);
  } finally {
    $('scanBtn').disabled = false;
    showProgress(false);
  }
});

// Reads a streamed NDJSON response from /api/scan, invoking onProgress for
// each 'progress' line and onPartial for each 'partial' line (newly-found
// result/skipped rows, so the UI can render them live), and returns the
// payload of the final 'done' line.
async function readNdjsonScanResponse(res, onProgress, onPartial) {
  if (!res.body || !res.body.getReader) {
    // Fallback for environments without streaming fetch support: read the
    // whole body at once and parse it the same way.
    const text = await res.text();
    return parseNdjsonText(text, onProgress, onPartial);
  }

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  let done = null;

  while (true) {
    const { done: streamDone, value } = await reader.read();
    if (streamDone) break;
    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split('\n');
    buffer = lines.pop();
    for (const line of lines) {
      const msg = parseNdjsonLine(line);
      if (!msg) continue;
      if (msg.type === 'progress') onProgress(msg.processed, msg.total, msg.currentFile, msg.findingsCount, msg.skipped);
      else if (msg.type === 'partial') onPartial && onPartial(msg.newResults || [], msg.newSkipped || []);
      else if (msg.type === 'done') done = msg;
    }
  }
  const trailing = parseNdjsonLine(buffer);
  if (trailing && trailing.type === 'done') done = trailing;

  return done;
}

function parseNdjsonText(text, onProgress, onPartial) {
  let done = null;
  text.split('\n').forEach(line => {
    const msg = parseNdjsonLine(line);
    if (!msg) return;
    if (msg.type === 'progress') onProgress(msg.processed, msg.total, msg.currentFile, msg.findingsCount, msg.skipped);
    else if (msg.type === 'partial') onPartial && onPartial(msg.newResults || [], msg.newSkipped || []);
    else if (msg.type === 'done') done = msg;
  });
  return done;
}

function parseNdjsonLine(line) {
  const trimmed = (line || '').trim();
  if (!trimmed) return null;
  try {
    return JSON.parse(trimmed);
  } catch {
    return null;
  }
}

// ---------- OCR availability ----------

// Checked once on load so the button/label reflect reality immediately
// instead of only failing when the user clicks "Run OCR".
async function checkOcrAvailability() {
  try {
    const res = await fetch('/api/ocr-status');
    const data = await res.json();
    state.ocrAvailable = !!data.available;
    state.ocrUnavailableReason = data.reason || '';
  } catch (e) {
    state.ocrAvailable = false;
    state.ocrUnavailableReason = 'Could not reach the server to check OCR availability.';
  }
  updateOcrButton();
}
checkOcrAvailability();

// ---------- OCR rescan of skipped PDFs ----------

function showOcrProgress(visible) {
  $('ocrProgressPanel').style.display = visible ? 'block' : 'none';
}

function setOcrProgress(processed, total, currentFile, findingsCount, skippedCount) {
  const pct = total > 0 ? Math.round((processed / total) * 100) : 0;
  $('ocrProgressBarFill').style.width = pct + '%';
  $('ocrProgressPercent').textContent = pct + '%';
  $('ocrProgressLabel').textContent = total > 0
    ? `${processed} of ${total} PDF(s) OCR'd${currentFile ? ' \u2014 ' + currentFile : ''}`
    : 'Preparing OCR...';
  const findingsEl = $('ocrProgressFindings');
  findingsEl.textContent = typeof findingsCount === 'number'
    ? `${findingsCount} finding${findingsCount === 1 ? '' : 's'} so far`
    : '';
  const skippedEl = $('ocrProgressSkipped');
  skippedEl.textContent = typeof skippedCount === 'number' && skippedCount > 0
    ? `${skippedCount} still unreadable`
    : '';
}

$('ocrBtn').addEventListener('click', async () => {
  const items = state.skipped.filter(s => state.ocrSelected.has(s.fullPath));
  if (!items.length) return;

  const keywords = parseKeywords($('keywords').value);
  if (keywords.length === 0) { setStatus('Enter at least one keyword before running OCR.', true); return; }

  const proceed = confirm(
    `Run OCR on ${items.length} original PDF file${items.length === 1 ? '' : 's'}?\n\n` +
    `This opens and renders each PDF directly to recognize text, which can take a while for large or multi-page files.`
  );
  if (!proceed) return;

  state.ocrRunning = true;
  updateOcrButton();
  setOcrProgress(0, 0, '', 0, 0);
  showOcrProgress(true);
  setStatus(`Running OCR on ${items.length} PDF(s)...`);

  try {
    const res = await fetch('/api/ocr', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        items: items.map(s => ({ fullPath: s.fullPath, fileName: s.fileName, directory: s.directory, extension: s.extension })),
        keywords,
        caseSensitive: $('caseSensitive').checked,
        wholeWord: $('wholeWord').checked,
        useRegex: $('useRegex').checked
      })
    });

    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error || 'OCR failed');
    }

    const data = await readNdjsonScanResponse(res, (processed, total, currentFile, findingsCount, skippedCount) => {
      setOcrProgress(processed, total, currentFile, findingsCount, skippedCount);
    });

    if (!data) throw new Error('OCR did not complete - no results received');

    const recovered = new Set(data.recoveredPaths || []);
    const newResults = Array.isArray(data.results) ? data.results : (data.results ? [data.results] : []);
    const stillSkipped = Array.isArray(data.skippedFiles) ? data.skippedFiles : (data.skippedFiles ? [data.skippedFiles] : []);

    // Drop the OCR'd files from the skipped list (recovered ones are gone
    // entirely; still-unreadable ones are replaced with the updated
    // ocrAttempted record from the server) and merge in any new matches.
    const ocrdPaths = new Set(items.map(s => s.fullPath));
    state.skipped = state.skipped.filter(s => !ocrdPaths.has(s.fullPath)).concat(stillSkipped);
    state.results = state.results.concat(newResults);
    state.scannedFiles = (state.scannedFiles || 0) + recovered.size;
    items.forEach(s => state.ocrSelected.delete(s.fullPath));

    let msg = `OCR complete: recovered ${recovered.size} of ${items.length} PDF(s), ${newResults.length} new match row(s).`;
    if (stillSkipped.length) msg += ` ${stillSkipped.length} still unreadable after OCR.`;
    setStatus(msg);

    populateFilterOptions();
    renderResults();
    renderSkipped();
  } catch (e) {
    setStatus('OCR error: ' + e.message, true);
  } finally {
    state.ocrRunning = false;
    updateOcrButton();
    showOcrProgress(false);
  }
});

// ---------- Filtering ----------

function populateFilterOptions() {
  const sel = $('filterKeyword');
  const unique = [...new Set(state.results.map(r => r.keyword))].sort();
  sel.innerHTML = '<option value="">All keywords</option>' + unique.map(k => `<option value="${escapeHtml(k)}">${escapeHtml(k)}</option>`).join('');
  $('filterPanel').style.display = state.results.length ? 'block' : 'none';
}

function getFilteredResults() {
  const kw = $('filterKeyword').value.toLowerCase();
  const dir = $('filterDir').value.toLowerCase();
  const file = $('filterFile').value.toLowerCase();
  const snippet = $('filterSnippet').value.toLowerCase();

  return state.results.filter(r => {
    if (kw && r.keyword.toLowerCase() !== kw) return false;
    if (dir && !(r.directory || '').toLowerCase().includes(dir)) return false;
    if (file && !(r.fileName || '').toLowerCase().includes(file)) return false;
    if (snippet && !toSnippetArray(r.snippets).some(s => s.toLowerCase().includes(snippet))) return false;
    return true;
  });
}

['filterKeyword', 'filterDir', 'filterFile', 'filterSnippet'].forEach(id => {
  $(id).addEventListener('input', renderResults);
  $(id).addEventListener('change', renderResults);
});

$('clearFilters').addEventListener('click', () => {
  $('filterKeyword').value = '';
  $('filterDir').value = '';
  $('filterFile').value = '';
  $('filterSnippet').value = '';
  renderResults();
});

// ---------- Render ----------

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function escapeRegExp(s) {
  return String(s ?? '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// Normalizes a result's `snippets` field to an array. Needed because
// PowerShell's ConvertTo-Json collapses a single-item pipeline result to a
// bare scalar instead of a 1-element array, so a result with exactly one
// match can come back (or be re-imported from a previously-exported file)
// as a plain string rather than an array of one string.
function toSnippetArray(snippets) {
  if (Array.isArray(snippets)) return snippets;
  if (snippets) return [snippets];
  return [];
}

// Builds non-overlapping highlight ranges for one or more search terms
// against raw (unescaped) snippet text. Earlier entries in `patterns` win
// when two terms overlap in the same spot.
function buildHighlightRanges(text, patterns) {
  const ranges = [];
  patterns.forEach(p => {
    if (!p.term) return;
    const re = new RegExp(escapeRegExp(p.term), 'gi');
    let m;
    while ((m = re.exec(text)) !== null) {
      ranges.push({ start: m.index, end: m.index + m[0].length, cls: p.cls });
      if (m[0].length === 0) re.lastIndex++;
    }
  });
  ranges.sort((a, b) => a.start - b.start || (b.end - b.start) - (a.end - a.start));
  const resolved = [];
  let lastEnd = -1;
  ranges.forEach(r => {
    if (r.start >= lastEnd) {
      resolved.push(r);
      lastEnd = r.end;
    }
  });
  return resolved;
}

// Renders a snippet with the matched keyword highlighted in red, and (if a
// "search within results" term is active) that term highlighted separately.
function highlightSnippet(text, keyword, searchTerm) {
  const raw = String(text ?? '');
  const ranges = buildHighlightRanges(raw, [
    { term: keyword, cls: 'kw-match' },
    { term: searchTerm, cls: 'term-match' }
  ]);
  if (!ranges.length) return escapeHtml(raw);
  let out = '';
  let pos = 0;
  ranges.forEach(r => {
    out += escapeHtml(raw.slice(pos, r.start));
    out += `<mark class="${r.cls}">${escapeHtml(raw.slice(r.start, r.end))}</mark>`;
    pos = r.end;
  });
  out += escapeHtml(raw.slice(pos));
  return out;
}

let lastFilteredResults = [];
// Explicit "expand all" intent, set only via the Expand all/Collapse all
// button. Deliberately NOT inferred from the live DOM before each render:
// when a filter (e.g. "search within results") temporarily narrows the
// results to zero rows, there's nothing left in the DOM to introspect, so
// a DOM-based check would silently forget that expand-all was active and
// fail to restore it once matching rows come back.
let expandAllIntent = false;

function renderResults() {
  const filtered = getFilteredResults();
  lastFilteredResults = filtered;
  const snippetTerm = $('filterSnippet').value.trim();
  const body = $('resultsBody');
  body.innerHTML = '';
  $('resultsPanel').style.display = state.results.length ? 'block' : 'none';
  $('resultsCount').textContent = `${filtered.length} of ${state.results.length} row(s)`;
  $('emptyState').style.display = filtered.length ? 'none' : 'block';

  filtered.forEach((r) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="keyword">${escapeHtml(r.keyword)}</td>
      <td><span class="filename-link" title="Open ${escapeHtml(r.fullPath)}">${escapeHtml(r.fileName)}</span>${r.viaOcr ? ' <span class="badge viaocr">via OCR</span>' : ''}</td>
      <td>${escapeHtml(r.extension)}</td>
      <td class="count">${r.count}</td>
      <td class="dir" title="${escapeHtml(r.fullPath)}">${escapeHtml(r.directory)}</td>
    `;
    tr.addEventListener('click', () => toggleDetail(tr, r, snippetTerm));
    const link = tr.querySelector('.filename-link');
    link.addEventListener('click', (e) => {
      e.stopPropagation();
      openFile(r.fullPath);
    });
    body.appendChild(tr);
  });

  if (expandAllIntent && filtered.length) {
    setAllExpanded(true); // re-expands the freshly built rows, updates the label itself
  } else {
    updateExpandAllLabel();
  }
}

// ---------- Expand / collapse all ----------

function areAllRowsExpanded() {
  const rows = document.querySelectorAll('#resultsBody > tr:not(.detail-row)');
  return rows.length > 0 && [...rows].every(tr => tr.nextElementSibling && tr.nextElementSibling.classList.contains('detail-row'));
}

function updateExpandAllLabel() {
  const btn = $('expandAllBtn');
  if (!btn) return;
  const rows = document.querySelectorAll('#resultsBody > tr:not(.detail-row)');
  btn.disabled = rows.length === 0;
  btn.textContent = areAllRowsExpanded() ? 'Collapse all' : 'Expand all';
}

function setAllExpanded(expand) {
  expandAllIntent = expand; // remember the user's explicit choice, independent of DOM state
  const term = $('filterSnippet').value.trim();
  const rows = document.querySelectorAll('#resultsBody > tr:not(.detail-row)');
  rows.forEach((tr, i) => {
    const existing = tr.nextElementSibling;
    const hasDetail = existing && existing.classList.contains('detail-row');
    if (expand && !hasDetail) {
      tr.after(buildDetailRow(lastFilteredResults[i], term));
    } else if (!expand && hasDetail) {
      existing.remove();
    }
  });
  updateExpandAllLabel();
}

$('expandAllBtn').addEventListener('click', () => setAllExpanded(!areAllRowsExpanded()));

// Opens a result's original file in its default application via the server
// (the browser has no way to launch a native app for a local path itself).
async function openFile(fullPath) {
  setStatus(`Opening ${fullPath}...`);
  try {
    const res = await fetch('/api/open-file', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fullPath })
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || 'Could not open file');
    setStatus('');
  } catch (e) {
    setStatus('Could not open file: ' + e.message, true);
  }
}

// A skipped PDF is a candidate for OCR if the server flagged it as
// ocrEligible (no extractable text layer - likely scanned/image-based) and
// OCR hasn't already been attempted against it.
function isOcrCandidate(s) {
  return !!s.ocrEligible && !s.ocrAttempted;
}

function renderSkipped() {
  const panel = $('skippedPanel');
  const body = $('skippedBody');
  body.innerHTML = '';

  if (!state.skipped.length) {
    panel.style.display = 'none';
    updateOcrButton();
    return;
  }
  panel.style.display = 'block';

  const candidates = state.skipped.filter(isOcrCandidate);
  const eligibleCount = state.skipped.filter(s => s.ocrEligible).length;
  let msg = `${state.skipped.length} file(s) skipped`;
  if (eligibleCount) msg += ` \u2014 ${eligibleCount} PDF(s) skipped due to no OCR (no extractable text layer)`;
  $('skippedCount').textContent = msg;

  state.skipped.forEach(s => {
    const tr = document.createElement('tr');
    const candidate = isOcrCandidate(s);
    const checkboxCell = candidate
      ? `<td class="checkcol"><input type="checkbox" class="ocrRowCheck" data-path="${escapeHtml(s.fullPath)}" ${state.ocrSelected.has(s.fullPath) ? 'checked' : ''}></td>`
      : `<td class="checkcol"></td>`;
    let badge = '';
    if (s.viaOcr === false && s.ocrAttempted) {
      badge = '<span class="badge attempted">OCR attempted</span>';
    } else if (candidate) {
      badge = '<span class="badge eligible">OCR eligible</span>';
    } else if (s.ocrAttempted) {
      badge = '<span class="badge attempted">OCR attempted</span>';
    }
    tr.innerHTML = `
      ${checkboxCell}
      <td title="${escapeHtml(s.fullPath)}">${escapeHtml(s.fileName)}</td>
      <td>${escapeHtml(s.extension)}</td>
      <td class="dir" title="${escapeHtml(s.fullPath)}">${escapeHtml(s.directory)}</td>
      <td class="reason">${escapeHtml(s.reason)}</td>
      <td>${badge}</td>
    `;
    body.appendChild(tr);
  });

  body.querySelectorAll('.ocrRowCheck').forEach(cb => {
    cb.addEventListener('change', () => {
      if (cb.checked) state.ocrSelected.add(cb.dataset.path);
      else state.ocrSelected.delete(cb.dataset.path);
      updateOcrButton();
      syncOcrSelectAll();
    });
  });

  syncOcrSelectAll();
  updateOcrButton();
}

function syncOcrSelectAll() {
  const candidates = state.skipped.filter(isOcrCandidate);
  const selectAll = $('ocrSelectAll');
  if (!candidates.length) {
    selectAll.checked = false;
    selectAll.indeterminate = false;
    selectAll.disabled = true;
    return;
  }
  selectAll.disabled = false;
  const selectedCount = candidates.filter(s => state.ocrSelected.has(s.fullPath)).length;
  selectAll.checked = selectedCount === candidates.length;
  selectAll.indeterminate = selectedCount > 0 && selectedCount < candidates.length;
}

function updateOcrButton() {
  const btn = $('ocrBtn');
  const note = $('ocrStatusNote');
  const count = state.ocrSelected.size;

  if (state.ocrAvailable === false) {
    btn.textContent = 'OCR unavailable';
    btn.disabled = true;
    btn.title = state.ocrUnavailableReason || 'OCR is not available on this machine.';
    if (note) note.textContent = state.ocrUnavailableReason ? `OCR unavailable \u2014 ${state.ocrUnavailableReason}` : 'OCR unavailable on this machine.';
    return;
  }

  btn.title = '';
  if (note) note.textContent = '';
  btn.textContent = count ? `Run OCR on ${count} selected PDF${count === 1 ? '' : 's'}` : 'Run OCR on selected PDFs';
  btn.disabled = state.ocrRunning || count === 0;
}

$('ocrSelectAll').addEventListener('change', () => {
  const candidates = state.skipped.filter(isOcrCandidate);
  if ($('ocrSelectAll').checked) {
    candidates.forEach(s => state.ocrSelected.add(s.fullPath));
  } else {
    candidates.forEach(s => state.ocrSelected.delete(s.fullPath));
  }
  renderSkipped();
});

function buildDetailRow(r, snippetTerm) {
  const detail = document.createElement('tr');
  detail.className = 'detail-row';
  // When "search within results" is active, only show snippets that
  // actually match it (still highlighted), rather than all of them.
  const term = (snippetTerm || '').trim();
  const allSnippets = toSnippetArray(r.snippets);
  const matching = term
    ? allSnippets.filter(s => s.toLowerCase().includes(term.toLowerCase()))
    : allSnippets;
  const snippets = matching.map(s => `<div class="snippet">${highlightSnippet(s, r.keyword, term)}</div>`).join('');
  const noneMsg = term ? 'No snippet matches your search within results' : 'No snippet available';
  detail.innerHTML = `
    <td colspan="5">
      <div class="full-path">${escapeHtml(r.fullPath)}</div>
      ${snippets || `<div class="snippet">${noneMsg}</div>`}
    </td>
  `;
  return detail;
}

function toggleDetail(tr, r, snippetTerm) {
  const existing = tr.nextElementSibling;
  if (existing && existing.classList.contains('detail-row')) {
    existing.remove();
  } else {
    tr.after(buildDetailRow(r, snippetTerm));
  }
  updateExpandAllLabel();
}

// ---------- Export ----------

function downloadFile(filename, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

$('exportJson').addEventListener('click', () => {
  const payload = {
    exportedAt: new Date().toISOString(),
    sourcePath: state.sourcePath,
    scannedFiles: state.scannedFiles,
    totalFiles: state.totalFiles,
    searchSettings: state.searchSettings,
    results: state.results,
    skippedFiles: state.skipped
  };
  downloadFile('docinspector-results.json', JSON.stringify(payload, null, 2), 'application/json');
});

$('exportTxt').addEventListener('click', () => {
  const lines = [];
  const s = state.searchSettings || {};
  lines.push(`DocInspector Results - Exported ${new Date().toISOString()}`);
  lines.push(`Source: ${state.sourcePath}`);
  lines.push(`Keywords: ${(s.keywords || []).join(', ')}`);
  lines.push(`Recursive: ${s.recursive ? 'yes' : 'no'}`);
  lines.push(`Case sensitive: ${s.caseSensitive ? 'yes' : 'no'}`);
  lines.push(`Whole word: ${s.wholeWord ? 'yes' : 'no'}`);
  lines.push(`Regex: ${s.useRegex ? 'yes' : 'no'}`);
  lines.push('================================');
  lines.push('');
  state.results.forEach(r => {
    lines.push(`Keyword: ${r.keyword}`);
    lines.push(`File: ${r.fileName}`);
    lines.push(`Path: ${r.fullPath}`);
    lines.push(`Count: ${r.count}`);
    if (r.viaOcr) lines.push('Via OCR: yes');
    lines.push('Snippets:');
    toSnippetArray(r.snippets).forEach(s => lines.push(`  - ${s}`));
    lines.push('');
  });

  if (state.skipped.length) {
    lines.push('================================');
    lines.push('SKIPPED FILES - REQUIRE MANUAL INSPECTION');
    lines.push('================================');
    lines.push('');
    state.skipped.forEach(s => {
      lines.push(`File: ${s.fileName}`);
      lines.push(`Path: ${s.fullPath}`);
      lines.push(`Reason: ${s.reason}`);
      if (s.ocrEligible) lines.push(`OCR eligible: ${!s.ocrAttempted ? 'yes' : 'no (already attempted)'}`);
      if (s.ocrAttempted) lines.push('OCR attempted: yes');
      lines.push('');
    });
  }

  downloadFile('docinspector-results.txt', lines.join('\n'), 'text/plain');
});

// ---------- Import ----------

$('importBtn').addEventListener('click', () => $('importFile').click());

$('importFile').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const text = await file.text();
  try {
    const parsed = parseImportedContent(text);
    state.results = parsed.results;
    state.skipped = parsed.skipped || [];
    state.sourcePath = parsed.sourcePath || '';
    state.scannedFiles = parsed.scannedFiles || 0;
    state.totalFiles = parsed.totalFiles || 0;
    state.searchSettings = parsed.searchSettings || { keywords: [], recursive: true, caseSensitive: false, wholeWord: false, useRegex: false };
    state.ocrSelected = new Set(state.skipped.filter(isOcrCandidate).map(s => s.fullPath));

    // Reflect the imported search back into the form so it's visible and
    // ready to re-run, rather than leaving stale/unrelated field values.
    const s = state.searchSettings;
    if (state.sourcePath) $('folderPath').value = state.sourcePath;
    $('keywords').value = (s.keywords || []).join('\n');
    $('recursive').checked = !!s.recursive;
    $('caseSensitive').checked = !!s.caseSensitive;
    $('wholeWord').checked = !!s.wholeWord;
    $('useRegex').checked = !!s.useRegex;

    populateFilterOptions();
    renderResults();
    renderSkipped();
    let msg = `Imported ${state.results.length} row(s) from ${file.name}.`;
    if (state.skipped.length) msg += ` ${state.skipped.length} skipped file(s) included.`;
    setStatus(msg);
  } catch (err) {
    setStatus('Could not parse imported file: ' + err.message, true);
  }
  e.target.value = '';
});

function parseImportedContent(text) {
  const trimmed = text.trim();
  if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
    const parsed = JSON.parse(trimmed);
    if (Array.isArray(parsed)) return { results: parsed };
    return {
      results: parsed.results || [],
      skipped: parsed.skippedFiles || [],
      sourcePath: parsed.sourcePath,
      scannedFiles: parsed.scannedFiles,
      totalFiles: parsed.totalFiles,
      searchSettings: parsed.searchSettings
    };
  }
  const [resultsSection, skippedSection] = text.split('SKIPPED FILES - REQUIRE MANUAL INSPECTION');
  return {
    results: parseTxtExport(resultsSection),
    skipped: skippedSection ? parseTxtSkipped(skippedSection) : [],
    sourcePath: extractTxtField(text, 'Source'),
    searchSettings: {
      keywords: extractTxtField(text, 'Keywords').split(',').map(k => k.trim()).filter(Boolean),
      recursive: extractTxtField(text, 'Recursive').toLowerCase() === 'yes',
      caseSensitive: extractTxtField(text, 'Case sensitive').toLowerCase() === 'yes',
      wholeWord: extractTxtField(text, 'Whole word').toLowerCase() === 'yes',
      useRegex: extractTxtField(text, 'Regex').toLowerCase() === 'yes'
    }
  };
}

function parseTxtSkipped(text) {
  const blocks = text.split(/\n\s*\n/).map(b => b.trim()).filter(b => b.startsWith('File:'));
  return blocks.map(block => {
    const lines = block.split('\n');
    const get = (label) => {
      const line = lines.find(l => l.startsWith(label + ':'));
      return line ? line.slice(label.length + 1).trim() : '';
    };
    const fullPath = get('Path');
    const fileName = fullPath ? fullPath.split(/[\\/]/).pop() : get('File');
    const directory = fullPath ? fullPath.slice(0, fullPath.length - fileName.length - 1) : '';
    const ocrEligibleRaw = get('OCR eligible').toLowerCase();
    return {
      fileName,
      directory,
      fullPath,
      extension: fileName.includes('.') ? fileName.slice(fileName.lastIndexOf('.')) : '',
      reason: get('Reason'),
      ocrEligible: ocrEligibleRaw.startsWith('yes'),
      ocrAttempted: get('OCR attempted').toLowerCase() === 'yes'
    };
  });
}

function extractTxtField(text, label) {
  const m = text.match(new RegExp(`^${label}:\\s*(.*)$`, 'm'));
  return m ? m[1].trim() : '';
}

function parseTxtExport(text) {
  const blocks = text.split(/\n\s*\n/).map(b => b.trim()).filter(b => b.startsWith('Keyword:'));
  return blocks.map(block => {
    const lines = block.split('\n');
    const get = (label) => {
      const line = lines.find(l => l.startsWith(label + ':'));
      return line ? line.slice(label.length + 1).trim() : '';
    };
    const fullPath = get('Path');
    const fileName = fullPath ? fullPath.split(/[\\/]/).pop() : '';
    const directory = fullPath ? fullPath.slice(0, fullPath.length - fileName.length - 1) : '';
    const snippets = lines.filter(l => l.trim().startsWith('- ')).map(l => l.trim().slice(2));
    return {
      keyword: get('Keyword'),
      fileName,
      directory,
      fullPath,
      extension: fileName.includes('.') ? fileName.slice(fileName.lastIndexOf('.')) : '',
      count: parseInt(get('Count'), 10) || snippets.length,
      snippets,
      viaOcr: get('Via OCR').toLowerCase() === 'yes'
    };
  });
}
