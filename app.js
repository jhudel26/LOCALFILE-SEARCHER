// app.js
// Modern dashboard, tabs + 2-column layout.
// SweetAlert2 for errors/success/warnings only.
// Uses pdf.js and SheetJS (CDN loaded in index.html).

const templateInput = document.getElementById('templateFile');
const dropArea = document.getElementById('dropArea');
const templateInfo = document.getElementById('templateInfo');
const previewWrap = document.getElementById('previewWrap');
const previewList = document.getElementById('previewList');
const clearTemplateBtn = document.getElementById('clearTemplate');
const downloadSampleBtn = document.getElementById('downloadSample');
const exportTemplateBtn = document.getElementById('exportTemplate');

const chooseSourceBtn = document.getElementById('chooseSource');
const chooseDestBtn = document.getElementById('chooseDest');
const sourceInfo = document.getElementById('sourceInfo');
const destInfo = document.getElementById('destInfo');

const startBtn = document.getElementById('startBtn');
const cancelBtn = document.getElementById('cancelBtn');
const copyFilesToggle = document.getElementById('copyFilesToggle');
const caseInsensitiveCheckbox = document.getElementById('caseInsensitive');

const statusText = document.getElementById('status');
const progressBar = document.getElementById('progressBar');
const progressText = document.getElementById('progressText');

const logEl = document.getElementById('log');
const logCount = document.getElementById('logCount');
const clearLogBtn = document.getElementById('clearLog');
const downloadReportBtn = document.getElementById('downloadReportBtn');

const templateTableWrap = document.getElementById('templateTableWrap');
const reportPreview = document.getElementById('reportPreview');

const tabButtons = document.querySelectorAll('.tab');
const tabPanels = document.querySelectorAll('.tabpanel');

// Search state
let searchResults = [];
let filesScanned = 0;
let matchesFound = 0;
let duplicatesFound = 0;
let fileHashes = new Map();

let sourceHandle = null;
let destHandle = null;
let templateWords = [];
let cancelRequested = false;
let lastReportBlobUrl = null;
let lastReportName = null;

// Setup pdf.js worker
if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://unpkg.com/pdfjs-dist@3.10.111/build/pdf.worker.min.js';
}

// Tabs behavior
tabButtons.forEach(btn => {
  btn.addEventListener('click', () => {
    tabButtons.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    const tab = btn.dataset.tab;
    tabPanels.forEach(p => {
      p.classList.toggle('hidden', p.id !== tab);
    });
  });
});

// Utility functions
async function calculateFileHash(file) {
  const buffer = await file.arrayBuffer();
  const hashBuffer = await crypto.subtle.digest('SHA-256', buffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Logging & UI
function log(msg) {
  const time = new Date().toLocaleTimeString();
  const line = `[${time}] ${msg}`;
  logEl.textContent = line + '\n' + logEl.textContent;
  logCount.textContent = `${logEl.textContent.split('\n').filter(Boolean).length} log entries`;
}

function clearLog() {
  logEl.textContent = '';
  logCount.textContent = '0 log entries';
}

function setStatus(s) {
  statusText.textContent = s;
}

function setProgress(done, total) {
  const pct = total === 0 ? 0 : Math.round((done / total) * 100);
  progressBar.style.width = pct + '%';
  progressText.textContent = `${done} / ${total}`;
}

// Template load (skip header)
async function loadXlsxFile(file) {
  try {
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const dataRows = rows.slice(1); // skip header
    templateWords = dataRows.map(r => (r[0] || '').toString().trim()).filter(Boolean);
    templateInfo.textContent = `Loaded ${templateWords.length} searchwords from "${file.name}" (header ignored)`;
    renderPreview();
    exportTemplateBtn.disabled = !templateWords.length;
    clearTemplateBtn.classList.remove('hidden');
    renderTemplateTable();
    Swal.fire({ icon: 'success', title: 'Template loaded', text: `${templateWords.length} words loaded.` });
    log(`Template loaded: ${file.name} (${templateWords.length} words)`);
  } catch (err) {
    console.error(err);
    templateInfo.textContent = 'Failed to read template';
    log('Error loading template: ' + (err.message || err));
    Swal.fire({ icon: 'error', title: 'Failed to load template', text: `${err.message || err}` });
  }
}

function renderPreview() {
  if (!templateWords.length) {
    previewWrap.classList.add('hidden');
    return;
  }
  previewWrap.classList.remove('hidden');
  previewList.innerHTML = '';
  const items = templateWords.slice(0, 8);
  for (const t of items) {
    const li = document.createElement('li');
    li.textContent = t;
    previewList.appendChild(li);
  }
  if (templateWords.length > 8) {
    const more = document.createElement('li');
    more.textContent = `... and ${templateWords.length - 8} more`;
    previewList.appendChild(more);
  }
}

function renderTemplateTable() {
  templateTableWrap.innerHTML = '';
  if (!templateWords.length) {
    templateTableWrap.textContent = 'No template loaded';
    return;
  }
  const table = document.createElement('table');
  table.className = 'w-full text-sm font-mono';
  const thead = document.createElement('thead');
  thead.innerHTML = `<tr class="text-xs text-slate-500"><th class="text-left">#</th><th class="text-left">searchword</th></tr>`;
  table.appendChild(thead);
  const tbody = document.createElement('tbody');
  templateWords.forEach((w, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td class="pr-3">${i+1}</td><td>${w}</td>`;
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  templateTableWrap.appendChild(table);
}

// drag & drop
dropArea.addEventListener('click', () => templateInput.click());
dropArea.addEventListener('dragover', (e) => { e.preventDefault(); dropArea.classList.add('ring-2', 'ring-indigo-400'); });
dropArea.addEventListener('dragleave', () => dropArea.classList.remove('ring-2', 'ring-indigo-400'));
dropArea.addEventListener('drop', (e) => {
  e.preventDefault();
  dropArea.classList.remove('ring-2', 'ring-indigo-400');
  const f = e.dataTransfer.files[0];
  if (f) {
    templateInput.files = e.dataTransfer.files;
    loadXlsxFile(f);
  }
});
templateInput.addEventListener('change', (e) => { const f = e.target.files[0]; if (f) loadXlsxFile(f); });

// clear template
clearTemplateBtn.addEventListener('click', () => {
  templateInput.value = '';
  templateWords = [];
  templateInfo.textContent = 'No template loaded';
  previewWrap.classList.add('hidden');
  exportTemplateBtn.disabled = true;
  clearTemplateBtn.classList.add('hidden');
  templateTableWrap.innerHTML = '';
  Swal.fire({ icon: 'success', title: 'Template cleared' });
  log('Template cleared');
});

// sample template download
downloadSampleBtn.addEventListener('click', () => {
  const rows = [['searchword'], [''], [''], ['']];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'template');
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([out], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'template_sample.xlsx';
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  Swal.fire({ icon: 'success', title: 'Sample downloaded' });
  log('Sample template downloaded');
});

// export loaded template
exportTemplateBtn.addEventListener('click', () => {
  if (!templateWords.length) return;
  const rows = [['searchword'], ...templateWords.map(w => [w])];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'template');
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([out], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `template_export_${new Date().toISOString().slice(0,10)}.xlsx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  Swal.fire({ icon: 'success', title: 'Template exported' });
  log('Template exported (download)');
});

// folder pickers
chooseSourceBtn.addEventListener('click', async () => {
  try { sourceHandle = await window.showDirectoryPicker(); sourceInfo.textContent = `Source: ${sourceHandle.name}`; log(`Source selected: ${sourceHandle.name}`); } 
  catch (err) { log('Source selection canceled'); }
});
chooseDestBtn.addEventListener('click', async () => {
  try { destHandle = await window.showDirectoryPicker(); destInfo.textContent = `Destination: ${destHandle.name}`; log(`Destination selected: ${destHandle.name}`); } 
  catch (err) { log('Destination selection canceled'); }
});

// cancel
cancelBtn.addEventListener('click', () => { cancelRequested = true; setStatus('Cancelling...'); Swal.fire({ icon: 'warning', title: 'Cancelling', text: 'Operation will stop shortly.' }); log('Cancel requested'); });

// download report
downloadReportBtn.addEventListener('click', () => {
  if (!lastReportBlobUrl) { Swal.fire({ icon: 'warning', title: 'No report available', text: 'Run a search first.' }); return; }
  const a = document.createElement('a'); a.href = lastReportBlobUrl; a.download = lastReportName || 'search_report.xlsx'; document.body.appendChild(a); a.click(); a.remove();
  log('Report downloaded by user (fallback)');
});

// helper: ensure nested path exists
async function ensurePathAndGetHandle(rootDirHandle, filePath) {
  const parts = filePath.split('/');
  const fileName = parts.pop();
  let curr = rootDirHandle;
  for (const p of parts) curr = await curr.getDirectoryHandle(p, { create: true });
  return await curr.getFileHandle(fileName, { create: true });
}

// core search logic
startBtn.addEventListener('click', async () => {
  if (!templateWords.length) { Swal.fire({ icon: 'error', title: 'No template', text: 'Upload a template first.' }); return; }
  if (!sourceHandle) { Swal.fire({ icon: 'error', title: 'No source folder', text: 'Please choose a source folder.' }); return; }
  if (!destHandle) { Swal.fire({ icon: 'error', title: 'No destination folder', text: 'Please choose a destination folder.' }); return; }

  cancelRequested = false;
  setStatus('Preparing...');
  log('Search started');
  lastReportBlobUrl && URL.revokeObjectURL(lastReportBlobUrl);
  lastReportBlobUrl = null;
  downloadReportBtn.disabled = true;
  reportPreview.innerHTML = '';
  const foundMap = {};
  for (const w of templateWords) foundMap[w] = [];

  // collect files
  const fileHandles = [];
  async function traverse(dirHandle, basePath = '') {
    for await (const [name, handle] of dirHandle.entries()) {
      if (cancelRequested) return;
      const entryPath = basePath ? `${basePath}/${name}` : name;
      if (handle.kind === 'file') fileHandles.push({ handle, path: entryPath });
      else if (handle.kind === 'directory') await traverse(handle, entryPath);
    }
  }
  await traverse(sourceHandle);
  log(`${fileHandles.length} files discovered`);
  setStatus('Scanning files...');
  setProgress(0, fileHandles.length);

  const caseInsensitive = caseInsensitiveCheckbox.checked;
  const norm = s => caseInsensitive ? s.toLowerCase() : s;
  const copyMatched = copyFilesToggle.checked;
  let copiedCount = 0;

  for (let i = 0; i < fileHandles.length; i++) {
    if (cancelRequested) break;
    const { handle, path } = fileHandles[i];
    try {
      const file = await handle.getFile();
      const nameNorm = norm(file.name);
      const fileHash = await calculateFileHash(file);
      if (fileHashes.has(fileHash)) duplicatesFound++;
      else fileHashes.set(fileHash, path);

      // filename match
      for (const w of templateWords) if (nameNorm.includes(norm(w))) foundMap[w].push(`FILENAME: ${sourceHandle.name}/${path}`);

      // PDF content
      if (file.name.toLowerCase().endsWith('.pdf')) {
        try {
          const ab = await file.arrayBuffer();
          const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
          for (let p = 1; p <= pdf.numPages; p++) {
            const page = await pdf.getPage(p);
            const txt = await page.getTextContent();
            const pageText = txt.items.map(i => i.str).join(' ');
            for (const w of templateWords) if (norm(pageText).includes(norm(w))) foundMap[w].push(`PDF_PAGE${p}: ${sourceHandle.name}/${path}`);
          }
        } catch (err) { log(`PDF parse error (${path}): ${err.message}`); }
      }

      // copy
      if (copyMatched && templateWords.some(w => foundMap[w].some(loc => loc.endsWith(path)))) {
        try {
          const destFileHandle = await ensurePathAndGetHandle(destHandle, path);
          const writable = await destFileHandle.createWritable();
          await writable.write(await file.arrayBuffer());
          await writable.close();
          copiedCount++;
        } catch (err) { log(`Copy failed for ${path}: ${err.message}`); }
      }

      filesScanned++;
      matchesFound = Object.values(foundMap).flat().length;
      setProgress(i + 1, fileHandles.length);
      document.getElementById('filesScanned').textContent = filesScanned;
      document.getElementById('matchesFound').textContent = matchesFound;
      document.getElementById('duplicatesFound').textContent = duplicatesFound;
      log(`Scanned: ${file.name}`);
    } catch (err) { log(`Error processing ${path}: ${err.message}`); }
  }

  setStatus(cancelRequested ? 'Cancelled' : 'Completed');
  if (cancelRequested) Swal.fire({ icon: 'warning', title: 'Cancelled', text: 'Operation stopped by user.' });
  else Swal.fire({ icon: 'success', title: 'Search finished', text: 'Search completed.' });

  // report
  const rows = [['searchword', 'result', 'where_found']];
  reportPreview.innerHTML = '';
  for (const w of templateWords) {
    const locations = [...new Set(foundMap[w])];
    const result = locations.length ? 'FOUND' : 'NOT FOUND';
    rows.push([w, result, locations.join(' ; ')]);
    const div = document.createElement('div');
    div.className = 'py-1 border-b last:border-b-0';
    div.innerHTML = `<div class="text-xs"><strong>${w}</strong> — <span class="text-slate-500">${result}</span></div>
                     <div class="text-xs font-mono text-slate-600">${locations.join(' ; ') || '-'}</div>`;
    reportPreview.appendChild(div);
  }

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'report');
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  lastReportBlobUrl && URL.revokeObjectURL(lastReportBlobUrl);
  lastReportBlobUrl = URL.createObjectURL(blob);
  lastReportName = `search_report_${new Date().toISOString().slice(0,10)}.xlsx`;
  downloadReportBtn.disabled = false;

  try {
    const fileHandle = await destHandle.getFileHandle('search_report.xlsx', { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(wbout);
    await writable.close();
    log('Report saved to destination folder');
  } catch { log('Could not save report to folder, fallback download available'); }

  log(`Search completed. Files checked: ${fileHandles.length}, Files copied: ${copiedCount}`);
});
