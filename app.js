// app.js
// Modern dashboard logic with 2-column layout and card-style results
// Works with Tailwind 3 + SweetAlert2 + Lucide + pdf.js + SheetJS

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

// search state
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

// ------------------- Utility Functions -------------------

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

function log(msg) {
  const time = new Date().toLocaleTimeString();
  const line = `[${time}] ${msg}`;
  logEl.textContent += line + '\n';
  logEl.scrollTop = logEl.scrollHeight;
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

// ------------------- Template Handling -------------------

async function loadXlsxFile(file) {
  try {
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    templateWords = rows.slice(1).map(r => (r[0] || '').toString().trim()).filter(Boolean);

    templateInfo.textContent = `Loaded ${templateWords.length} words from "${file.name}"`;
    previewWrap.classList.remove('hidden');
    renderPreview();
    exportTemplateBtn.disabled = !templateWords.length;
    clearTemplateBtn.classList.remove('hidden');
    renderTemplateTable();

    Swal.fire({ icon: 'success', title: 'Template loaded', text: `${templateWords.length} words loaded.` });
    log(`Template loaded: ${file.name} (${templateWords.length} words)`);
  } catch (err) {
    console.error(err);
    templateInfo.textContent = 'Failed to read template';
    Swal.fire({ icon: 'error', title: 'Failed to load template', text: err.message || err });
    log('Error loading template: ' + (err.message || err));
  }
}

function renderPreview() {
  previewList.innerHTML = '';
  const items = templateWords.slice(0, 8);
  items.forEach(t => {
    const li = document.createElement('li');
    li.textContent = t;
    li.className = 'px-2 py-1 bg-indigo-100 dark:bg-indigo-900/30 rounded text-indigo-800 dark:text-indigo-200 text-xs';
    previewList.appendChild(li);
  });
  if (templateWords.length > 8) {
    const more = document.createElement('li');
    more.textContent = `... and ${templateWords.length - 8} more`;
    more.className = 'text-slate-500 dark:text-slate-400 text-xs';
    previewList.appendChild(more);
  }
}

function renderTemplateTable() {
  templateTableWrap.innerHTML = '';
  if (!templateWords.length) return;
  const table = document.createElement('table');
  table.className = 'w-full text-sm font-mono border border-slate-300 dark:border-slate-700 rounded';
  const thead = document.createElement('thead');
  thead.innerHTML = `<tr class="bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300"><th class="p-1 text-left">#</th><th class="p-1 text-left">Word</th></tr>`;
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  templateWords.forEach((w, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td class="p-1 border-t border-slate-200 dark:border-slate-700">${i+1}</td><td class="p-1 border-t border-slate-200 dark:border-slate-700">${w}</td>`;
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  templateTableWrap.appendChild(table);
}

// ------------------- Folder Selection -------------------

chooseSourceBtn.addEventListener('click', async () => {
  try { sourceHandle = await window.showDirectoryPicker(); sourceInfo.textContent = `Source: ${sourceHandle.name}`; log(`Source selected: ${sourceHandle.name}`); } 
  catch { log('Source selection canceled'); }
});

chooseDestBtn.addEventListener('click', async () => {
  try { destHandle = await window.showDirectoryPicker(); destInfo.textContent = `Destination: ${destHandle.name}`; log(`Destination selected: ${destHandle.name}`); } 
  catch { log('Destination selection canceled'); }
});

// ------------------- Start Search -------------------

startBtn.addEventListener('click', async () => {
  if (!templateWords.length) return Swal.fire({ icon:'error', title:'No template', text:'Upload a template first.' });
  if (!sourceHandle) return Swal.fire({ icon:'error', title:'No source', text:'Choose a source folder.' });
  if (!destHandle) return Swal.fire({ icon:'error', title:'No dest', text:'Choose a destination folder.' });

  cancelRequested = false;
  setStatus('Preparing...');
  log('Search started');
  lastReportBlobUrl && URL.revokeObjectURL(lastReportBlobUrl);
  lastReportBlobUrl = null;
  downloadReportBtn.disabled = true;
  reportPreview.innerHTML = '';
  searchResults = [];

  // Traverse files
  const fileHandles = [];
  async function traverse(dir, base='') {
    for await (const [name, handle] of dir.entries()) {
      if (cancelRequested) return;
      const path = base ? `${base}/${name}` : name;
      if (handle.kind === 'file') fileHandles.push({handle, path});
      else if (handle.kind === 'directory') await traverse(handle, path);
    }
  }
  await traverse(sourceHandle);
  setStatus('Scanning files...');
  setProgress(0, fileHandles.length);
  log(`${fileHandles.length} files discovered`);

  const caseInsensitive = caseInsensitiveCheckbox.checked;
  const copyMatched = copyFilesToggle.checked;

  // initialize found map
  const foundMap = {};
  templateWords.forEach(w => foundMap[w] = []);

  let checked = 0;
  let copiedCount = 0;

  for (const f of fileHandles) {
    if (cancelRequested) break;
    checked++;
    try {
      const nameNorm = caseInsensitive ? f.handle.name.toLowerCase() : f.handle.name;
      templateWords.forEach(w => {
        const term = caseInsensitive ? w.toLowerCase() : w;
        if (nameNorm.includes(term)) foundMap[w].push(`${sourceHandle.name}/${f.path}`);
      });

      // PDF first page check
      if (f.handle.name.toLowerCase().endsWith('.pdf')) {
        try {
          const file = await f.handle.getFile();
          const ab = await file.arrayBuffer();
          const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
          if (pdf.numPages >= 1) {
            const page = await pdf.getPage(1);
            const txt = await page.getTextContent();
            const pageText = txt.items.map(i => i.str).join(' ');
            templateWords.forEach(w => {
              const term = caseInsensitive ? w.toLowerCase() : w;
              if ((caseInsensitive ? pageText.toLowerCase() : pageText).includes(term)) foundMap[w].push(`${sourceHandle.name}/${f.path}`);
            });
          }
        } catch { /* ignore PDF parse errors */ }
      }

      // copy matched files
      if (copyMatched) {
        const matched = templateWords.some(w => foundMap[w].includes(`${sourceHandle.name}/${f.path}`));
        if (matched) {
          try {
            const targetHandle = await ensurePathAndGetHandle(destHandle, f.path);
            const arrayBuffer = await (await f.handle.getFile()).arrayBuffer();
            const writable = await targetHandle.createWritable();
            await writable.write(arrayBuffer); await writable.close();
            copiedCount++;
          } catch {}
        }
      }

    } catch (err) { log(`Error scanning ${f.path}: ${err.message || err}`); }
    if (checked % 5 === 0 || checked === fileHandles.length) setProgress(checked, fileHandles.length);
  }

  if (cancelRequested) { setStatus('Cancelled'); log('Operation cancelled'); Swal.fire({icon:'warning',title:'Cancelled'}); }
  else { setStatus('Completed'); log(`Search finished. Files copied: ${copiedCount}`); Swal.fire({icon:'success',title:'Search Finished'}); }

  // Generate report
  const rows = [['searchword','result','where_found']];
  reportPreview.innerHTML = '';
  templateWords.forEach(w => {
    const locs = [...new Set(foundMap[w])];
    rows.push([w, locs.length ? 'FOUND':'NOT FOUND', locs.join(' ; ')]);
    const div = document.createElement('div');
    div.className = 'p-2 border-b border-slate-200 dark:border-slate-700';
    div.innerHTML = `<div class="font-bold">${w}</div><div class="text-xs text-slate-600 dark:text-slate-400">${locs.join(' ; ') || '-'}</div>`;
    reportPreview.appendChild(div);
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'report');
  const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});

  const blob = new Blob([wbout], {type:'application/octet-stream'});
  lastReportBlobUrl && URL.revokeObjectURL(lastReportBlobUrl);
  lastReportBlobUrl = URL.createObjectURL(blob);
  lastReportName = `search_report_${new Date().toISOString().slice(0,10)}.xlsx`;
  downloadReportBtn.disabled = false;

  // save to destination
  try {
    const fileHandle = await destHandle.getFileHandle('search_report.xlsx', {create:true});
    const writable = await fileHandle.createWritable();
    await writable.write(wbout); await writable.close();
    log('Report saved to destination folder');
  } catch { log('Could not save report to destination; use Download Report'); }
  setProgress(fileHandles.length, fileHandles.length);
});

// ------------------- Helper Functions -------------------

async function ensurePathAndGetHandle(rootDirHandle, filePath) {
  const parts = filePath.split('/');
  const fileName = parts.pop();
  let curr = rootDirHandle;
  for (const p of parts) curr = await curr.getDirectoryHandle(p, {create:true});
  return await curr.getFileHandle(fileName, {create:true});
}

// ------------------- UI Event Listeners -------------------

clearLogBtn.addEventListener('click', clearLog);

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

downloadSampleBtn.addEventListener('click', () => {
  const rows = [['searchword'],[''],[''],['']];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'template');
  const out = XLSX.write(wb, {bookType:'xlsx', type:'array'});
  const blob = new Blob([out], {type:'application/octet-stream'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'template_sample.xlsx';
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
  Swal.fire({icon:'success',title:'Sample downloaded'});
  log('Sample template downloaded');
});

exportTemplateBtn.addEventListener('click', () => {
  if (!templateWords.length) return;
  const rows = [['searchword'], ...templateWords.map(w=>[w])];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'template');
  const out = XLSX.write(wb, {bookType:'xlsx', type:'array'});
  const blob = new Blob([out], {type:'application/octet-stream'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = `template_export_${new Date().toISOString().slice(0,10)}.xlsx`;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
  Swal.fire({icon:'success',title:'Template exported'});
  log('Template exported');
});

cancelBtn.addEventListener('click', () => { cancelRequested=true; setStatus('Cancelling...'); Swal.fire({icon:'warning',title:'Cancelling'}); log('Cancel requested'); });
downloadReportBtn.addEventListener('click', () => {
  if (!lastReportBlobUrl) return Swal.fire({icon:'warning',title:'No report available'});
  const a = document.createElement('a');
  a.href = lastReportBlobUrl; a.download = lastReportName; document.body.appendChild(a); a.click(); a.remove();
  log('Report downloaded by user');
});
