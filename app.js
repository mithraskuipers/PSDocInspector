let state = {
  results: [],
  skipped: [],
  sourcePath: '',
  scannedFiles: 0,
  totalFiles: 0,
  expandedRow: null
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

function setProgress(processed, total, currentFile) {
  const pct = total > 0 ? Math.round((processed / total) * 100) : 0;
  $('progressBarFill').style.width = pct + '%';
  $('progressPercent').textContent = pct + '%';
  $('progressLabel').textContent = total > 0
    ? `${processed} of ${total} file(s) scanned${currentFile ? ' \u2014 ' + currentFile : ''}`
    : 'Preparing scan...';
}

$('scanBtn').addEventListener('click', async () => {
  const path = $('folderPath').value.trim();
  const keywords = parseKeywords($('keywords').value);

  if (!path) { setStatus('Enter or browse to a folder first.', true); return; }
  if (keywords.length === 0) { setStatus('Enter at least one keyword.', true); return; }

  $('scanBtn').disabled = true;
  setProgress(0, 0, '');
  showProgress(true);
  setStatus('Scanning...');

  try {
    const res = await fetch('/api/scan', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        path,
        recursive: $('recursive').checked,
        keywords,
        caseSensitive: $('caseSensitive').checked,
        wholeWord: $('wholeWord').checked,
        useRegex: $('useRegex').checked
      })
    });

    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error || 'Scan failed');
    }

    // The server streams newline-delimited JSON (NDJSON): a 'start' line,
    // one 'progress' line per file processed, then a final 'done' line
    // carrying the full result set. This is what powers the live progress bar.
    const data = await readNdjsonScanResponse(res, (processed, total, currentFile) => {
      setProgress(processed, total, currentFile);
    });

    if (!data) throw new Error('Scan did not complete - no results received');

    state.results = Array.isArray(data.results) ? data.results : (data.results ? [data.results] : []);
    state.skipped = Array.isArray(data.skippedFiles) ? data.skippedFiles : (data.skippedFiles ? [data.skippedFiles] : []);
    state.sourcePath = data.sourcePath || path;
    state.scannedFiles = data.scannedFiles || 0;
    state.totalFiles = data.totalFiles || 0;

    let msg = `Scanned ${data.scannedFiles} of ${data.totalFiles} file(s), ${state.results.length} match row(s).`;
    if (state.skipped.length) msg += ` ${state.skipped.length} file(s) skipped - see below.`;
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
// each 'progress' line, and returns the payload of the final 'done' line.
async function readNdjsonScanResponse(res, onProgress) {
  if (!res.body || !res.body.getReader) {
    // Fallback for environments without streaming fetch support: read the
    // whole body at once and parse it the same way.
    const text = await res.text();
    return parseNdjsonText(text, onProgress);
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
      if (msg.type === 'progress') onProgress(msg.processed, msg.total, msg.currentFile);
      else if (msg.type === 'done') done = msg;
    }
  }
  const trailing = parseNdjsonLine(buffer);
  if (trailing && trailing.type === 'done') done = trailing;

  return done;
}

function parseNdjsonText(text, onProgress) {
  let done = null;
  text.split('\n').forEach(line => {
    const msg = parseNdjsonLine(line);
    if (!msg) return;
    if (msg.type === 'progress') onProgress(msg.processed, msg.total, msg.currentFile);
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

  return state.results.filter(r => {
    if (kw && r.keyword.toLowerCase() !== kw) return false;
    if (dir && !(r.directory || '').toLowerCase().includes(dir)) return false;
    if (file && !(r.fileName || '').toLowerCase().includes(file)) return false;
    return true;
  });
}

['filterKeyword', 'filterDir', 'filterFile'].forEach(id => {
  $(id).addEventListener('input', renderResults);
  $(id).addEventListener('change', renderResults);
});

$('clearFilters').addEventListener('click', () => {
  $('filterKeyword').value = '';
  $('filterDir').value = '';
  $('filterFile').value = '';
  renderResults();
});

// ---------- Render ----------

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function renderResults() {
  const filtered = getFilteredResults();
  const body = $('resultsBody');
  body.innerHTML = '';
  $('resultsPanel').style.display = state.results.length ? 'block' : 'none';
  $('resultsCount').textContent = `${filtered.length} of ${state.results.length} row(s)`;
  $('emptyState').style.display = filtered.length ? 'none' : 'block';

  filtered.forEach((r, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="keyword">${escapeHtml(r.keyword)}</td>
      <td>${escapeHtml(r.fileName)}</td>
      <td>${escapeHtml(r.extension)}</td>
      <td class="count">${r.count}</td>
      <td class="dir" title="${escapeHtml(r.fullPath)}">${escapeHtml(r.directory)}</td>
    `;
    tr.addEventListener('click', () => toggleDetail(tr, r, idx));
    body.appendChild(tr);
  });
}

function renderSkipped() {
  const panel = $('skippedPanel');
  const body = $('skippedBody');
  body.innerHTML = '';

  if (!state.skipped.length) {
    panel.style.display = 'none';
    return;
  }
  panel.style.display = 'block';
  $('skippedCount').textContent = `${state.skipped.length} file(s) skipped`;

  state.skipped.forEach(s => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td title="${escapeHtml(s.fullPath)}">${escapeHtml(s.fileName)}</td>
      <td>${escapeHtml(s.extension)}</td>
      <td class="dir" title="${escapeHtml(s.fullPath)}">${escapeHtml(s.directory)}</td>
      <td class="reason">${escapeHtml(s.reason)}</td>
    `;
    body.appendChild(tr);
  });
}

function toggleDetail(tr, r, idx) {
  const existing = tr.nextElementSibling;
  if (existing && existing.classList.contains('detail-row')) {
    existing.remove();
    return;
  }
  document.querySelectorAll('.detail-row').forEach(el => el.remove());

  const detail = document.createElement('tr');
  detail.className = 'detail-row';
  const snippets = (r.snippets || []).map(s => `<div class="snippet">${escapeHtml(s)}</div>`).join('');
  detail.innerHTML = `
    <td colspan="5">
      <div class="full-path">${escapeHtml(r.fullPath)}</div>
      ${snippets || '<div class="snippet">No snippet available</div>'}
    </td>
  `;
  tr.after(detail);
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
    results: state.results,
    skippedFiles: state.skipped
  };
  downloadFile('docinspector-results.json', JSON.stringify(payload, null, 2), 'application/json');
});

$('exportTxt').addEventListener('click', () => {
  const lines = [];
  lines.push(`DocInspector Results - Exported ${new Date().toISOString()}`);
  lines.push(`Source: ${state.sourcePath}`);
  lines.push('================================');
  lines.push('');
  state.results.forEach(r => {
    lines.push(`Keyword: ${r.keyword}`);
    lines.push(`File: ${r.fileName}`);
    lines.push(`Path: ${r.fullPath}`);
    lines.push(`Count: ${r.count}`);
    lines.push('Snippets:');
    (r.snippets || []).forEach(s => lines.push(`  - ${s}`));
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
      totalFiles: parsed.totalFiles
    };
  }
  const [resultsSection, skippedSection] = text.split('SKIPPED FILES - REQUIRE MANUAL INSPECTION');
  return {
    results: parseTxtExport(resultsSection),
    skipped: skippedSection ? parseTxtSkipped(skippedSection) : [],
    sourcePath: extractTxtField(text, 'Source')
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
    return {
      fileName,
      directory,
      fullPath,
      extension: fileName.includes('.') ? fileName.slice(fileName.lastIndexOf('.')) : '',
      reason: get('Reason')
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
      snippets
    };
  });
}
