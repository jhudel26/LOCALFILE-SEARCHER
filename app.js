// app.js (module)
const templateInput = document.getElementById('templateFile');
const dropArea = document.getElementById('dropArea');
const templateLabel = document.getElementById('templateLabel');
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

let sourceHandle = null;
let destHandle = null;
let templateWords = [];
let cancelRequested = false;
let lastReportBlobUrl = null;
let lastReportName = null;

// pdf.js worker
if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://unpkg.com/pdfjs-dist@3.10.111/build/pdf.worker.min.js';
}

// --- UI Helpers ---
function log(msg) {
  const time = new Date().toLocaleTimeString();
  const line = `[${time}] ${msg}`;
  logEl.textContent += line + '\n';
  logEl.scrollTop = logEl.scrollHeight;
  logCount.textContent = `${logEl.textContent.split('\n').filter(Boolean).length} entries`;
}

function clearLog() {
  logEl.textContent = '';
  logCount.textContent = '0 entries';
}

function setStatus(s) {
  statusText.textContent = s;
}

function setProgress(done, total) {
  const pct = total === 0 ? 0 : Math.round((done / total) * 100);
  progressBar.style.width = pct + '%';
  progressText.textContent = `${done} / ${total}`;
}

// --- Template handling ---
async function loadXlsxFile(file) {
  try {
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab, { type: 'array' });
    const firstSheet = wb.SheetNames[0];
    const ws = wb.Sheets[firstSheet];
    // read as array of arrays
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    templateWords = rows.map(r => (r[0] || '').toString().trim()).filter(Boolean);
    templateInfo.textContent = `Loaded ${templateWords.length} search words from "${file.name}"`;
    renderPreview();
    exportTemplateBtn.disabled = !templateWords.length;
    clearTemplateBtn.classList.remove('hidden');
    log(`Template loaded: ${file.name} (${templateWords.length} words)`);
  } catch (err) {
    console.error(err);
    templateInfo.textContent = 'Failed to read template';
    log('Error loading template: ' + (err.message || err));
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

// drag & drop
dropArea.addEventListener('click', () => templateInput.click());
dropArea.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropArea.classList.add('ring-2', 'ring-indigo-400');
});
dropArea.addEventListener('dragleave', () => {
  dropArea.classList.remove('ring-2', 'ring-indigo-400');
});
dropArea.addEventListener('drop', (e) => {
  e.preventDefault();
  dropArea.classList.remove('ring-2', 'ring-indigo-400');
  const f = e.dataTransfer.files[0];
  if (f) {
    templateInput.files = e.dataTransfer.files;
    loadXlsxFile(f);
  }
});

templateInput.addEventListener('change', (e) => {
  const f = e.target.files[0];
  if (f) loadXlsxFile(f);
});

clearTemplateBtn.addEventListener('click', () => {
  templateInput.value = '';
  templateWords = [];
  templateInfo.textContent = 'No template loaded';
  previewWrap.classList.add('hidden');
  exportTemplateBtn.disabled = true;
  clearTemplateBtn.classList.add('hidden');
  log('Template cleared by user');
});

// Sample template generator
downloadSampleBtn.addEventListener('click', () => {
  const rows = [
    ['searchword'],
    ['AMI000001'],
    ['AMI000002'],
    ['AMI000100'],
  ];
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
  log('Sample template downloaded');
});

// Export loaded template (re-create xlsx from templateWords)
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
  log('Template exported (download)');
});

// --- Folder pickers (File System Access API) ---
chooseSourceBtn.addEventListener('click', async () => {
  try {
    sourceHandle = await window.showDirectoryPicker();
    sourceInfo.textContent = `Source: ${sourceHandle.name}`;
    log(`Source folder selected: ${sourceHandle.name}`);
  } catch (err) {
    log('Source selection canceled.');
  }
});

chooseDestBtn.addEventListener('click', async () => {
  try {
    destHandle = await window.showDirectoryPicker();
    destInfo.textContent = `Destination: ${destHandle.name}`;
    log(`Destination folder selected: ${destHandle.name}`);
  } catch (err) {
    log('Destination selection canceled.');
  }
});

// --- Control buttons ---
cancelBtn.addEventListener('click', () => {
  cancelRequested = true;
  setStatus('Cancelling...');
  log('Cancel requested');
});

clearLogBtn.addEventListener('click', clearLog);

// Download report button fallback
downloadReportBtn.addEventListener('click', () => {
  if (!lastReportBlobUrl) {
    alert('No report available to download.');
    return;
  }
  const a = document.createElement('a');
  a.href = lastReportBlobUrl;
  a.download = lastReportName || 'search_report.xlsx';
  document.body.appendChild(a);
  a.click();
  a.remove();
  log('Report downloaded (fallback)');
});

// --- Core search logic ---
startBtn.addEventListener('click', async () => {
  if (!templateWords.length) {
    alert('Please upload an Excel template (.xlsx) with searchwords in the first column.');
    return;
  }
  if (!sourceHandle) {
    alert('Please choose a source folder.');
    return;
  }
  if (!destHandle) {
    alert('Please choose a destination folder.');
    return;
  }

  cancelRequested = false;
  setStatus('Preparing...');
  log('Search started');

  // Build found map
  const foundMap = {};
  for (const w of templateWords) foundMap[w] = [];

  // Collect file handles
  const fileHandles = [];
  async function traverse(dirHandle, basePath = '') {
    for await (const [name, handle] of dirHandle.entries()) {
      if (cancelRequested) return;
      const entryPath = basePath ? `${basePath}/${name}` : name;
      if (handle.kind === 'file') {
        fileHandles.push({ handle, path: entryPath });
      } else if (handle.kind === 'directory') {
        await traverse(handle, entryPath);
      }
    }
  }

  await traverse(sourceHandle);
  log(`${fileHandles.length} files discovered`);

  let checked = 0;
  setProgress(checked, fileHandles.length);
  setStatus('Scanning files...');
  const caseInsensitive = caseInsensitiveCheckbox.checked;
  const norm = s => caseInsensitive ? s.toLowerCase() : s;

  const copyMatched = copyFilesToggle.checked;
  let copiedCount = 0;

  for (const fileEntry of fileHandles) {
    if (cancelRequested) break;
    const { handle, path } = fileEntry;
    checked++;
    try {
      const name = handle.name;
      const nameToTest = norm(name);

      for (const w of templateWords) {
        const testWord = norm(w);
        if (nameToTest.includes(testWord)) {
          foundMap[w].push(`FILENAME: ${path}`);
        }
      }

      if (name.toLowerCase().endsWith('.pdf')) {
        try {
          const file = await handle.getFile();
          const ab = await file.arrayBuffer();
          const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
          if (pdf.numPages >= 1) {
            const page = await pdf.getPage(1);
            const txt = await page.getTextContent();
            const pageText = txt.items.map(i => i.str).join(' ');
            const textToTest = norm(pageText);
            for (const w of templateWords) {
              const testWord = norm(w);
              if (textToTest.includes(testWord)) {
                foundMap[w].push(`PDF_FIRST_PAGE: ${path}`);
              }
            }
            page.cleanup && page.cleanup();
          }
          pdf.cleanup && pdf.cleanup();
        } catch (pdfErr) {
          log(`PDF parse error (${path}): ${pdfErr.message || pdfErr}`);
        }
      }

      // Optional copy matched files to destination
      if (copyMatched) {
        // If any word matched in filename or PDF first page (foundMap... contains entries)
        const matched = templateWords.some(w => foundMap[w].some(loc => loc.endsWith(path)));
        if (matched) {
          // copy file into dest preserving folder name
          try {
            const targetHandle = await ensurePathAndGetHandle(destHandle, path);
            const srcFile = await handle.getFile();
            const writable = await targetHandle.createWritable();
            await writable.write(await srcFile.arrayBuffer());
            await writable.close();
            copiedCount++;
            log(`Copied matched file: ${path}`);
          } catch (copyErr) {
            log(`Failed copying ${path}: ${copyErr.message || copyErr}`);
          }
        }
      }
    } catch (err) {
      log(`Error scanning ${path}: ${err.message || err}`);
    }

    if (checked % 5 === 0 || checked === fileHandles.length) {
      setProgress(checked, fileHandles.length);
    }
  }

  if (cancelRequested) {
    setStatus('Cancelled');
    log('Operation cancelled by user');
  } else {
    setStatus('Completed');
    log('Scanning finished');
  }

  // Build rows for report
  const rows = [];
  rows.push(['searchword', 'result', 'where_found']);
  for (const w of templateWords) {
    const locations = [...new Set(foundMap[w])];
    const result = locations.length ? 'FOUND' : 'NOT FOUND';
    const where = locations.join(' ; ');
    rows.push([w, result, where]);
  }

  // Create workbook
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'report');
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

  // Try to save to dest folder; fallback to download
  try {
    const fileHandle = await destHandle.getFileHandle('search_report.xlsx', { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(wbout);
    await writable.close();
    log('Report saved to destination: search_report.xlsx');
    setStatus('Report saved');
    downloadReportBtn.disabled = false;
    lastReportBlobUrl = null;
  } catch (saveErr) {
    log('Saving to destination failed; creating fallback download.');
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    if (lastReportBlobUrl) URL.revokeObjectURL(lastReportBlobUrl);
    lastReportBlobUrl = URL.createObjectURL(blob);
    lastReportName = 'search_report.xlsx';
    downloadReportBtn.disabled = false;
    setStatus('Report ready to download');
    log('Use "Download Report" to save the file.');
  }

  // Final progress
  setProgress(fileHandles.length, fileHandles.length);
  log(`Search completed. Files checked: ${fileHandles.length}. Matches copied: ${copiedCount}`);
});

// helper: ensure nested path exists and return file handle for writing that file
async function ensurePathAndGetHandle(rootDirHandle, filePath) {
  // filePath could be like "a/b/c.pdf" or just "c.pdf"
  const parts = filePath.split('/');
  const fileName = parts.pop();
  let curr = rootDirHandle;
  for (const p of parts) {
    curr = await curr.getDirectoryHandle(p, { create: true });
  }
  return await curr.getFileHandle(fileName, { create: true });
}
