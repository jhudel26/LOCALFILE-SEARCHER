// app.js (module)
// DOM Elements
const templateInput = document.getElementById('templateFile');
const dropArea = document.getElementById('dropArea');
const templateLabel = document.getElementById('templateLabel');
const templateInfo = document.getElementById('templateInfo');
const previewWrap = document.getElementById('previewWrap');
const previewList = document.getElementById('previewList');
const termCount = document.getElementById('termCount');
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
const statusIndicator = document.getElementById('statusIndicator');
const statusDot = document.querySelector('.status-dot');
const progressBar = document.getElementById('progressBar');
const progressText = document.getElementById('progressText');
const progressInfo = document.getElementById('progressInfo');

const logEl = document.getElementById('log');
const logCount = document.getElementById('logCount');
const clearLogBtn = document.getElementById('clearLog');
const downloadReportBtn = document.getElementById('downloadReportBtn');

// App state
let sourceHandle = null;
let destHandle = null;
let templateWords = [];
let cancelRequested = false;
let lastReportBlobUrl = null;
let lastReportName = null;
let isProcessing = false;

// Initialize PDF.js worker
if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://unpkg.com/pdfjs-dist@3.10.111/build/pdf.worker.min.js';
}

// --- UI Helpers ---
function log(msg, type = 'info') {
  const time = new Date().toLocaleTimeString();
  const line = document.createElement('div');
  line.className = `log-entry ${type}`;
  
  // Add icon based on message type
  let icon = 'ℹ️';
  if (type === 'success') icon = '✅';
  else if (type === 'error') icon = '❌';
  else if (type === 'warning') icon = '⚠️';
  
  line.textContent = `[${time}] ${icon} ${msg}`;
  logEl.appendChild(line);
  logEl.scrollTop = logEl.scrollHeight;
  
  // Update log count
  const entryCount = logEl.children.length;
  logCount.textContent = `${entryCount} ${entryCount === 1 ? 'entry' : 'entries'}`;
  
  // Add visual feedback for errors
  if (type === 'error') {
    line.classList.add('text-red-500', 'dark:text-red-400');
  } else if (type === 'success') {
    line.classList.add('text-green-600', 'dark:text-green-400');
  } else if (type === 'warning') {
    line.classList.add('text-amber-600', 'dark:text-amber-400');
  }
}

async function clearLog() {
  const result = await Swal.fire({
    title: 'Clear Activity Log',
    text: 'Are you sure you want to clear the activity log? This action cannot be undone.',
    icon: 'warning',
    showCancelButton: true,
    confirmButtonColor: '#4f46e5',
    cancelButtonColor: '#6b7280',
    confirmButtonText: 'Yes, clear it!',
    cancelButtonText: 'Cancel',
    reverseButtons: true
  });

  if (result.isConfirmed) {
    logEl.innerHTML = '';
    logCount.textContent = '0 entries';
    
    await Swal.fire(
      'Cleared!',
      'The activity log has been cleared.',
      'success'
    );
    
    log('Log cleared', 'info');
  }
}

function setStatus(status, type = 'info') {
  statusText.textContent = status;
  
  // Update status dot
  statusDot.className = 'status-dot';
  if (type === 'processing') {
    statusDot.classList.add('processing');
    statusText.className = 'text-amber-500';
  } else if (type === 'success') {
    statusDot.classList.add('success');
    statusText.className = 'text-green-500';
  } else if (type === 'error') {
    statusDot.classList.add('error');
    statusText.className = 'text-red-500';
  } else {
    statusDot.classList.add('idle');
    statusText.className = 'text-slate-500 dark:text-slate-400';
  }
  
  // Update progress info
  if (progressInfo) {
    if (type === 'processing') {
      progressInfo.textContent = 'Processing files...';
    } else if (type === 'success') {
      progressInfo.textContent = 'Search completed successfully';
    } else if (type === 'error') {
      progressInfo.textContent = 'An error occurred';
    } else {
      progressInfo.textContent = 'Ready to start searching';
    }
  }
}

function setProgress(done, total) {
  const pct = total === 0 ? 0 : Math.round((done / total) * 100);
  progressBar.style.width = `${pct}%`;
  progressText.textContent = `${done} / ${total} files`;
  
  // Update progress info with more details
  if (progressInfo && total > 0) {
    const remaining = total - done;
    progressInfo.textContent = `Processed ${done} of ${total} files (${pct}%)`;
    
    // Add estimated time remaining if we have enough data
    if (done > 5) {
      const timePerFile = performance.now() / (done * 1000); // in seconds
      const remainingTime = Math.round(timePerFile * remaining);
      if (remainingTime > 0) {
        const mins = Math.floor(remainingTime / 60);
        const secs = Math.round(remainingTime % 60);
        progressInfo.textContent += ` • ${mins > 0 ? `${mins}m ` : ''}${secs}s remaining`;
      }
    }
  }
}

function showNotification(message, type = 'info') {
  // Create notification element
  const notification = document.createElement('div');
  notification.className = `fixed top-4 right-4 p-4 rounded-lg shadow-lg z-50 transform transition-all duration-300 translate-x-0 opacity-100 ${
    type === 'success' ? 'bg-green-100 text-green-800 dark:bg-green-900/80 dark:text-green-200' :
    type === 'error' ? 'bg-red-100 text-red-800 dark:bg-red-900/80 dark:text-red-200' :
    'bg-blue-100 text-blue-800 dark:bg-blue-900/80 dark:text-blue-200'
  }`;
  
  notification.innerHTML = `
    <div class="flex items-center">
      <i class="${
        type === 'success' ? 'fas fa-check-circle' :
        type === 'error' ? 'fas fa-exclamation-circle' :
        'fas fa-info-circle'
      } mr-2"></i>
      <span>${message}</span>
      <button class="ml-4 text-current opacity-70 hover:opacity-100">
        <i class="fas fa-times"></i>
      </button>
    </div>
  `;
  
  // Add click handler to dismiss
  notification.querySelector('button').addEventListener('click', () => {
    notification.classList.add('translate-x-full', 'opacity-0');
    setTimeout(() => notification.remove(), 300);
  });
  
  // Auto-dismiss after 5 seconds
  document.body.appendChild(notification);
  setTimeout(() => {
    notification.classList.add('translate-x-full', 'opacity-0');
    setTimeout(() => notification.remove(), 300);
  }, 5000);
}

function setProcessingState(processing) {
  isProcessing = processing;
  
  // Update UI based on processing state
  startBtn.disabled = processing;
  chooseSourceBtn.disabled = processing;
  chooseDestBtn.disabled = processing;
  downloadSampleBtn.disabled = processing;
  exportTemplateBtn.disabled = processing || templateWords.length === 0;
  
  // Show/hide cancel button
  cancelBtn.classList.toggle('hidden', !processing);
  
  // Update status
  if (processing) {
    setStatus('Processing...', 'processing');
    log('Starting search operation...', 'info');
  } else if (cancelRequested) {
    setStatus('Cancelled', 'error');
    log('Operation cancelled by user', 'warning');
  } else {
    setStatus('Ready', 'idle');
  }
}

// --- Template handling ---
async function loadXlsxFile(file) {
  try {
    setProcessingState(true);
    
    // Show loading state
    templateInfo.textContent = 'Loading template...';
    templateLabel.textContent = file.name;
    
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab, { type: 'array' });
    const firstSheet = wb.SheetNames[0];
    const ws = wb.Sheets[firstSheet];
    
    // Read as array of arrays and process
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    templateWords = [...new Set(rows.map(r => (r[0] || '').toString().trim()).filter(Boolean))];
    
    if (templateWords.length === 0) {
      throw new Error('No valid search terms found in the first column');
    }
    
    // Update UI
    templateInfo.textContent = `Loaded ${templateWords.length} unique search terms`;
    templateLabel.textContent = file.name;
    templateLabel.title = file.name; // Add tooltip for long filenames
    
    renderPreview();
    exportTemplateBtn.disabled = false;
    clearTemplateBtn.classList.remove('hidden');
    
    log(`Template loaded: "${file.name}" with ${templateWords.length} unique search terms`, 'success');
    showNotification(`Template loaded with ${templateWords.length} search terms`, 'success');
    
  } catch (err) {
    console.error('Error loading template:', err);
    templateInfo.textContent = 'Failed to read template';
    log(`Error loading template: ${err.message || 'Invalid file format'}`, 'error');
    showNotification('Failed to load template. Please check the file format.', 'error');
    
    // Reset template input
    templateInput.value = '';
    templateWords = [];
    renderPreview();
  } finally {
    setProcessingState(false);
  }
}

function renderPreview() {
  if (!templateWords || templateWords.length === 0) {
    previewWrap.classList.add('hidden');
    termCount.textContent = '0';
    return;
  }
  
  previewWrap.classList.remove('hidden');
  previewList.innerHTML = '';
  
  // Update term count
  termCount.textContent = templateWords.length;
  
  // Show first 5 terms with a "show more" button if there are more
  const maxVisible = 5;
  const itemsToShow = templateWords.slice(0, maxVisible);
  
  itemsToShow.forEach((term, index) => {
    const li = document.createElement('div');
    li.className = 'flex items-center py-1 px-2 hover:bg-slate-100 dark:hover:bg-slate-700/50 rounded';
    li.innerHTML = `
      <span class="inline-block w-5 text-xs text-slate-400">${index + 1}.</span>
      <span class="font-mono text-sm truncate" title="${term}">${term}</span>
    `;
    previewList.appendChild(li);
  });
  
  // Add "show more" button if there are more terms
  if (templateWords.length > maxVisible) {
    const moreCount = templateWords.length - maxVisible;
    const moreBtn = document.createElement('button');
    moreBtn.className = 'text-xs text-indigo-600 dark:text-indigo-400 hover:underline mt-1';
    moreBtn.textContent = `+ ${moreCount} more terms...`;
    moreBtn.addEventListener('click', (e) => {
      e.preventDefault();
      showAllTerms();
    });
    
    const container = document.createElement('div');
    container.className = 'text-center pt-1 border-t border-slate-100 dark:border-slate-700 mt-2';
    container.appendChild(moreBtn);
    previewList.appendChild(container);
  }
}

function showAllTerms() {
  if (!templateWords.length) return;
  
  // Create a modal to show all terms
  const modal = document.createElement('div');
  modal.className = 'fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50 p-4';
  
  modal.innerHTML = `
    <div class="bg-white dark:bg-slate-800 rounded-xl shadow-2xl w-full max-w-md max-h-[80vh] flex flex-col">
      <div class="p-4 border-b border-slate-200 dark:border-slate-700 flex justify-between items-center">
        <h3 class="text-lg font-semibold">All Search Terms (${templateWords.length})</h3>
        <button class="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200">
          <i class="fas fa-times"></i>
        </button>
      </div>
      <div class="p-4 overflow-y-auto flex-1">
        <div class="space-y-1">
          ${templateWords.map((term, i) => `
            <div class="flex items-center py-1.5 px-2 hover:bg-slate-100 dark:hover:bg-slate-700/50 rounded">
              <span class="w-8 text-sm text-slate-400">${i + 1}.</span>
              <span class="font-mono text-sm break-all">${term}</span>
            </div>
          `).join('')}
        </div>
      </div>
      <div class="p-4 border-t border-slate-200 dark:border-slate-700 flex justify-end">
        <button class="btn btn-ghost text-sm">Close</button>
      </div>
    </div>
  `;
  
  // Add close handlers
  const closeBtn = modal.querySelector('button');
  const closeModal = () => {
    modal.classList.add('opacity-0');
    setTimeout(() => modal.remove(), 200);
  };
  
  closeBtn.addEventListener('click', closeModal);
  modal.addEventListener('click', (e) => {
    if (e.target === modal) closeModal();
  });
  
  // Add escape key handler
  const handleKeyDown = (e) => {
    if (e.key === 'Escape') closeModal();
  };
  
  document.addEventListener('keydown', handleKeyDown);
  
  // Clean up event listener when modal is closed
  modal.addEventListener('animationend', function handler() {
    if (modal.classList.contains('opacity-0')) {
      document.removeEventListener('keydown', handleKeyDown);
      modal.removeEventListener('animationend', handler);
    }
  });
  
  // Add to DOM with animation
  document.body.appendChild(modal);
  requestAnimationFrame(() => {
    modal.classList.add('opacity-100');
  });
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
  const pathSeparator = '/'; // Use forward slash for consistency
  
  async function traverse(dirHandle, basePath = '') {
    for await (const [name, handle] of dirHandle.entries()) {
      if (cancelRequested) return;
      const entryPath = basePath ? `${basePath}${pathSeparator}${name}` : name;
      
      if (handle.kind === 'file') {
        fileHandles.push({ 
          handle, 
          path: entryPath,
          name: name
        });
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
          const pdf = await pdfjsLib.getDocument({ 
            data: ab,
            // Add these options for better error handling
            disableWorker: false,
            disableAutoFetch: true,
            disableStream: true
          }).promise;
          
          if (pdf.numPages >= 1) {
            const page = await pdf.getPage(1);
            try {
              const txt = await page.getTextContent();
              // Add null check for txt.items
              const pageText = txt.items ? txt.items.map(i => i.str || '').join(' ') : '';
              const textToTest = norm(pageText);
              for (const w of templateWords) {
                const testWord = norm(w);
                if (textToTest.includes(testWord)) {
                  foundMap[w].push(`PDF_FIRST_PAGE: ${path}`);
                }
              }
            } finally {
              // Ensure page resources are always cleaned up
              if (page.cleanup) {
                try {
                  await page.cleanup();
                } catch (cleanupErr) {
                  console.warn('Page cleanup error:', cleanupErr);
                }
              }
            }
          }
          // Ensure PDF resources are always cleaned up
          if (pdf.cleanup) {
            try {
              await pdf.cleanup();
            } catch (cleanupErr) {
              console.warn('PDF cleanup error:', cleanupErr);
            }
          }
        } catch (pdfErr) {
          // Safely handle the error to prevent uncaught exceptions
          try {
            const errorMessage = pdfErr && (pdfErr.message || String(pdfErr));
            log(`❌ PDF error (${path}): ${errorMessage}`, 'error');
            console.error('PDF processing error:', {
              error: pdfErr,
              name: pdfErr && pdfErr.name,
              message: errorMessage,
              stack: pdfErr && pdfErr.stack
            });
          } catch (logErr) {
            // Last resort error handling if even the error logging fails
            console.error('Critical error in PDF error handler:', logErr);
          }
        }
      }

      // Optional copy matched files to destination
      if (copyMatched) {
        // Check if this file was matched in any way
        const matched = templateWords.some(w => 
          foundMap[w].some(loc => {
            // Match either the full path or just the filename
            const locPath = loc.split(':').pop().trim();
            return path.endsWith(locPath) || path.includes(locPath);
          })
        );
        
        if (matched) {
          try {
            log(`⏳ Preparing to copy: ${path}`);
            const targetHandle = await ensurePathAndGetHandle(destHandle, path);
            const srcFile = await handle.getFile();
            
            if (srcFile) {
              log(`📄 Source file size: ${(srcFile.size / 1024).toFixed(2)} KB`);
              const writable = await targetHandle.createWritable();
              const fileContent = await srcFile.arrayBuffer();
              await writable.write(fileContent);
              await writable.close();
              copiedCount++;
              log(`✅ Successfully copied: ${path}`);
            } else {
              log(`❌ Source file not accessible: ${path}`, 'error');
            }
          } catch (copyErr) {
            log(`❌ Failed to copy ${path}: ${copyErr.message || copyErr}`, 'error');
            console.error('Copy error details:', {
              error: copyErr,
              name: copyErr.name,
              stack: copyErr.stack
            });
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
  
  // Track all unique matched files
  const allMatchedFiles = new Set();
  
  for (const w of templateWords) {
    const locations = [...new Set(foundMap[w])];
    const result = locations.length ? 'FOUND' : 'NOT FOUND';
    const where = locations.join(' ; ');
    rows.push([w, result, where]);
    
    // Extract file paths from locations
    locations.forEach(loc => {
      const match = loc.match(/:\s*(.*)$/);
      if (match && match[1]) {
        allMatchedFiles.add(match[1]);
      }
    });
  }
  
  // Log summary of matched files
  if (allMatchedFiles.size > 0) {
    log('\n📋 Matched Files Summary:');
    log('======================');
    Array.from(allMatchedFiles).sort().forEach((file, index) => {
      log(`${index + 1}. ${file}`);
    });
    log(`\nTotal matched files: ${allMatchedFiles.size}`);
    
    if (copyMatched) {
      log(`\n📦 Copied ${copiedCount} file(s) to destination folder`);
      if (copiedCount < allMatchedFiles.size) {
        log(`ℹ️  Some files might not have been copied due to errors. Check the log above for details.`, 'warning');
      }
    } else {
      log('\nℹ️  File copying is disabled. Enable "Copy matched files" to copy files to destination.');
    }
  } else {
    log('\nℹ️  No files matched the search criteria.');
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
