// app.js - Optimized for Vercel Deployment

// Function to switch between tabs
function switchToTab(tabId) {
  try {
    // Hide all tab panels
    const tabPanels = document.querySelectorAll('.tabpanel');
    if (tabPanels && tabPanels.length > 0) {
      tabPanels.forEach(panel => {
        if (panel && panel.classList) {
          panel.classList.add('hidden');
        }
      });
    }

    // Show the selected tab panel
    const targetPanel = document.getElementById(tabId);
    if (targetPanel && targetPanel.classList) {
      targetPanel.classList.remove('hidden');
    }

    // Update tab styling
    const tabs = document.querySelectorAll('.tab');
    if (tabs && tabs.length > 0) {
      tabs.forEach(tab => {
        if (tab && tab.classList) {
          if (tab.dataset.tab === tabId) {
            tab.classList.remove('border-transparent', 'text-gray-500', 'hover:text-gray-700', 'hover:border-gray-300', 'dark:text-gray-400', 'dark:hover:text-gray-300');
            tab.classList.add('border-primary-500', 'text-primary-600', 'dark:text-primary-400');
          } else {
            tab.classList.add('border-transparent', 'text-gray-500', 'hover:text-gray-700', 'hover:border-gray-300', 'dark:text-gray-400', 'dark:hover:text-gray-300');
            tab.classList.remove('border-primary-500', 'text-primary-600', 'dark:text-primary-400');
          }
        }
      });
    }
  } catch (err) {
    console.error('Error switching tabs:', err);
  }
}

const templateInput = document.getElementById('templateFile');
const dropArea = document.getElementById('dropArea');
const templateInfo = document.getElementById('templateInfo');
const previewWrap = document.getElementById('previewWrap');
const previewList = document.getElementById('previewList');
const clearTemplateBtn = document.getElementById('clearTemplate');
const downloadSampleBtn = document.getElementById('downloadSampleBtn');
const exportTemplateBtn = document.getElementById('exportTemplateBtn');

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

// PDF.js worker
if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://unpkg.com/pdfjs-dist@3.10.111/build/pdf.worker.min.js';
}

// Utility
async function calculateFileHash(file) {
  const buffer = await file.arrayBuffer();
  const hashBuffer = await crypto.subtle.digest('SHA-256', buffer);
  return Array.from(new Uint8Array(hashBuffer)).map(b => b.toString(16).padStart(2,'0')).join('');
}

function formatFileSize(bytes) {
  if (!bytes) return '0 Bytes';
  const sizes = ['Bytes','KB','MB','GB'];
  const i = Math.floor(Math.log(bytes)/Math.log(1024));
  return `${(bytes / Math.pow(1024, i)).toFixed(2)} ${sizes[i]}`;
}

// Logging
function log(msg){
  const time = new Date().toLocaleTimeString();
  logEl.textContent = `[${time}] ${msg}\n` + logEl.textContent;
  logCount.textContent = logEl.textContent.split('\n').filter(Boolean).length + ' log entries';
}
function clearLog(){ logEl.textContent=''; logCount.textContent='0 log entries'; }

// Status & progress
function setStatus(s){ statusText.textContent = s; }
function setProgress(done,total){
  const pct = total ? Math.round((done/total)*100) : 0;
  progressBar.style.width = pct + '%';
  progressText.textContent = `${done} / ${total}`;
}

// Drag & Drop + Click Fallback
dropArea.addEventListener('dragover', e => { e.preventDefault(); dropArea.classList.add('ring-2','ring-indigo-400'); e.dataTransfer.dropEffect='copy'; });
dropArea.addEventListener('dragleave', e => { e.preventDefault(); dropArea.classList.remove('ring-2','ring-indigo-400'); });
dropArea.addEventListener('drop', e => {
  e.preventDefault();
  dropArea.classList.remove('ring-2','ring-indigo-400');
  if(!e.dataTransfer.files.length) return;
  const f = e.dataTransfer.files[0];
  templateInput.files = e.dataTransfer.files;
  loadXlsxFile(f);
});
dropArea.addEventListener('click', () => templateInput.click());
templateInput.addEventListener('change', e => { if(e.target.files[0]) loadXlsxFile(e.target.files[0]); });

// Load Excel template (skip header)
async function loadXlsxFile(file) {
  try {
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab, {type: 'array'});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, {header: 1, defval: ''});
    templateWords = rows.slice(1).map(r => r[0]?.toString().trim()).filter(Boolean);
    templateInfo.textContent = `Loaded ${templateWords.length} words from "${file.name}" (header ignored)`;
    renderPreview(); 
    renderTemplateTable();
    exportTemplateBtn.disabled = !templateWords.length;
    clearTemplateBtn.classList.remove('hidden');
    Swal.fire({icon: 'success', title: 'Template loaded', text: `${templateWords.length} words loaded.`});
    log(`Template loaded: ${file.name} (${templateWords.length} words)`);
    
    // Only try to switch tab if the element exists
    if (document.getElementById('sourceDestTab')) {
      // Use setTimeout to ensure DOM is ready
      setTimeout(() => switchToTab('sourceDestTab'), 100);
    }
  } catch(err) {
    console.error(err);
    templateInfo.textContent = 'Failed to read template';
    Swal.fire({icon: 'error', title: 'Failed to load template', text: err.message || err});
    log('Error loading template: ' + (err.message || err));
  }
}

// Preview & table
function renderPreview() {
  // Get or create the previewWrap element
  let previewWrap = document.getElementById('previewWrap');
  
  if (!previewWrap) {
    // If previewWrap doesn't exist, create it
    previewWrap = document.createElement('div');
    previewWrap.id = 'previewWrap';
    previewWrap.className = 'mt-4';
    
    // Find the template info element and insert after it
    const templateInfo = document.getElementById('templateInfo');
    if (templateInfo && templateInfo.parentNode) {
      templateInfo.parentNode.insertBefore(previewWrap, templateInfo.nextSibling);
    }
    
    // Update the previewWrap reference
    window.previewWrap = previewWrap;
  }

  if (!templateWords.length) {
    if (previewWrap) {
      previewWrap.classList.add('hidden');
    }
    return;
  }

  if (previewWrap) {
    previewWrap.classList.remove('hidden');
    previewWrap.innerHTML = ''; // Clear previous content
    
    // Create the preview content
    const previewContent = document.createElement('div');
    previewContent.className = 'bg-white dark:bg-gray-800 shadow rounded-lg p-4';
    
    const title = document.createElement('h3');
    title.className = 'text-sm font-medium text-gray-900 dark:text-white mb-2';
    title.textContent = 'Template Preview';
    
    const previewList = document.createElement('ul');
    previewList.className = 'list-disc pl-6 mt-2';
    
    templateWords.slice(0, 8).forEach(t => {
      const li = document.createElement('li');
      li.textContent = t;
      li.className = 'text-sm text-gray-700 dark:text-gray-300';
      previewList.appendChild(li);
    });
    
    if (templateWords.length > 8) {
      const li = document.createElement('li');
      li.textContent = `... and ${templateWords.length - 8} more`;
      li.className = 'text-sm text-gray-500 dark:text-gray-400 italic';
      previewList.appendChild(li);
    }
    
    // Assemble the preview
    previewContent.appendChild(title);
    previewContent.appendChild(previewList);
    previewWrap.appendChild(previewContent);
  }
}
function renderTemplateTable(){
  templateTableWrap.innerHTML='';
  if(!templateWords.length){ templateTableWrap.textContent='No template loaded'; return; }
  const table=document.createElement('table'); table.className='w-full text-sm font-mono';
  table.innerHTML='<thead class="text-xs text-slate-500"><tr><th>#</th><th>searchword</th></tr></thead>';
  const tbody=document.createElement('tbody');
  templateWords.forEach((w,i)=>{ const tr=document.createElement('tr'); tr.innerHTML=`<td>${i+1}</td><td>${w}</td>`; tbody.appendChild(tr); });
  table.appendChild(tbody); templateTableWrap.appendChild(table);
}

// Clear template
clearTemplateBtn.addEventListener('click',()=>{
  templateInput.value=''; templateWords=[];
  templateInfo.textContent='No template loaded';
  previewWrap.classList.add('hidden');
  exportTemplateBtn.disabled=true;
  clearTemplateBtn.classList.add('hidden');
  templateTableWrap.innerHTML='';
  Swal.fire({icon:'success',title:'Template cleared'});
  log('Template cleared');
});

// Sample download
downloadSampleBtn.addEventListener('click',()=>{
  const rows=[['searchword'],[''],[''],['']];
  const wb=XLSX.utils.book_new();
  const ws=XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb,ws,'template');
  const out=XLSX.write(wb,{bookType:'xlsx',type:'array'});
  const blob=new Blob([out],{type:'application/octet-stream'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='template_sample.xlsx'; a.click(); a.remove();
  Swal.fire({icon:'success',title:'Sample downloaded'}); log('Sample template downloaded');
});

// Export template
exportTemplateBtn.addEventListener('click',()=>{
  if(!templateWords.length) return;
  const rows=[['searchword'], ...templateWords.map(w=>[w])];
  const wb=XLSX.utils.book_new();
  const ws=XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb,ws,'template');
  const out=XLSX.write(wb,{bookType:'xlsx',type:'array'});
  const blob=new Blob([out],{type:'application/octet-stream'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=`template_export_${new Date().toISOString().slice(0,10)}.xlsx`; a.click(); a.remove();
  Swal.fire({icon:'success',title:'Template exported'}); log('Template exported');
});

// Folder pickers
chooseSourceBtn.addEventListener('click', async () => {
  try { 
    sourceHandle = await window.showDirectoryPicker(); 
    sourceInfo.textContent = `Source: ${sourceHandle.name}`; 
    log(`Source selected: ${sourceHandle.name}`);
    
    // If destination is also selected, go to Search tab
    if (destHandle) {
      switchToTab('searchTab');
    }
  } catch(err) { 
    log('Source selection canceled'); 
  }
});

chooseDestBtn.addEventListener('click', async () => {
  try { 
    destHandle = await window.showDirectoryPicker(); 
    destInfo.textContent = `Destination: ${destHandle.name}`; 
    log(`Destination selected: ${destHandle.name}`);
    
    // If source is also selected, go to Search tab
    if (sourceHandle) {
      switchToTab('searchTab');
    }
  } catch(err) { 
    log('Destination selection canceled'); 
  }
});

// Start search
startBtn.addEventListener('click', async () => {
  if (!templateWords.length) {
    Swal.fire({icon: 'error', title: 'No Template', text: 'Please upload a template first.'});
    return switchToTab('templateTab');
  }
  
  if (!sourceHandle) {
    Swal.fire({icon: 'error', title: 'No Source', text: 'Please select a source directory.'});
    return switchToTab('sourceDestTab');
  }
  
  if (!destHandle) {
    Swal.fire({icon: 'error', title: 'No Destination', text: 'Please select a destination directory.'});
    return switchToTab('sourceDestTab');
  }
  
  try {
    // Reset state
    searchResults = [];
    filesScanned = 0;
    matchesFound = 0;
    duplicatesFound = 0;
    fileHashes.clear();
    cancelRequested = false;
    
    // Update UI
    setStatus('Starting search...');
    setProgress(0, 1);
    startBtn.disabled = true;
    cancelBtn.disabled = false;
    
    // Switch to logs tab
    switchToTab('logsTab');
    
    // Start the search
    await searchFiles();
    
    // Generate report when done
    if (!cancelRequested) {
      await generateReport();
      switchToTab('reportTab');
    }
  } catch (err) {
    console.error('Search error:', err);
    setStatus('Error during search');
    log(`Error: ${err.message || err}`);
    Swal.fire({icon: 'error', title: 'Search Error', text: err.message || 'An error occurred during search'});
  } finally {
    startBtn.disabled = false;
    cancelBtn.disabled = true;
  }
});

// Cancel
cancelBtn.addEventListener('click', () => {
  cancelRequested = true;
  setStatus('Cancelling...');
  Swal.fire({icon: 'warning', title: 'Cancelling', text: 'Operation will stop shortly.'});
  log('Cancel requested');
});

// Download report fallback
downloadReportBtn.addEventListener('click', ()=>{
  if(!lastReportBlobUrl){ Swal.fire({icon:'warning',title:'No report',text:'Run a search first.'}); return; }
  const a=document.createElement('a'); a.href=lastReportBlobUrl; a.download=lastReportName||'search_report.xlsx'; a.click(); a.remove();
  log('Report downloaded (fallback)');
});
// Search through files
async function searchFiles() {
  if (!sourceHandle) {
    throw new Error('No source directory selected');
  }

  try {
    // Reset counters
    filesScanned = 0;
    matchesFound = 0;
    searchResults = [];
    
    // Get all files recursively
    const files = await getFilesRecursively(sourceHandle);
    const totalFiles = files.length;
    
    if (totalFiles === 0) {
      setStatus('No files found in the selected directory');
      return;
    }

    setStatus(`Searching in ${totalFiles} files...`);
    log(`Starting search in ${totalFiles} files...`);

    // Process each file
    for (let i = 0; i < files.length; i++) {
      if (cancelRequested) {
        setStatus('Search cancelled');
        log('Search cancelled by user');
        return;
      }

      const file = files[i];
      filesScanned++;
      setProgress(i + 1, totalFiles);
      setStatus(`Searching... (${i + 1}/${totalFiles})`);

      try {
        const content = await readFileContent(file);
        const matches = findMatches(content);
        
        if (matches.length > 0) {
          matchesFound += matches.length;
          searchResults.push({
            path: file.relativePath,
            name: file.name,
            matches: matches
          });
          log(`Found ${matches.length} match(es) in ${file.name}`);
        }
      } catch (err) {
        console.error(`Error processing ${file.name}:`, err);
        log(`Error processing ${file.name}: ${err.message}`);
      }
    }

    setStatus(`Search complete. Found ${matchesFound} matches in ${searchResults.length} files.`);
    log(`Search complete. Scanned ${filesScanned} files, found ${matchesFound} matches in ${searchResults.length} files.`);
    
    // Generate report
    await generateReport();
    
  } catch (err) {
    console.error('Search error:', err);
    setStatus('Error during search');
    log(`Search error: ${err.message}`);
    throw err;
  }
}

// Helper function to get all files recursively
async function getFilesRecursively(directoryHandle, path = '') {
  const files = [];
  
  try {
    // Get all entries in the current directory
    const entries = [];
    for await (const entry of directoryHandle.values()) {
      entries.push(entry);
    }
    
    // Process each entry
    for (const entry of entries) {
      try {
        const entryPath = path ? `${path}/${entry.name}` : entry.name;
        
        if (entry.kind === 'file') {
          // Only process certain file types
          const ext = entry.name.split('.').pop().toLowerCase();
          if (['txt', 'pdf', 'doc', 'docx', 'xls', 'xlsx', 'csv'].includes(ext)) {
            files.push({
              handle: entry,
              name: entry.name,
              relativePath: entryPath
            });
          }
        } else if (entry.kind === 'directory') {
          // Recursively process subdirectories
          const dirFiles = await getFilesRecursively(entry, entryPath);
          files.push(...dirFiles);
        }
      } catch (err) {
        console.warn(`Error processing ${entry?.name || 'unknown entry'}:`, err);
        log(`Warning: Could not process ${entry?.name || 'a file/directory'}. It may be inaccessible.`);
      }
    }
  } catch (err) {
    console.error(`Error reading directory ${directoryHandle?.name || 'unknown'}:`, err);
    log(`Error: Could not read directory ${directoryHandle?.name || 'unknown'}. It may be inaccessible.`);
  }
  
  return files;
}

// Helper function to read file content
async function readFileContent(file) {
  try {
    const fileHandle = file.handle;
    const fileData = await fileHandle.getFile();
    
    // Handle different file types
    const ext = file.name.split('.').pop().toLowerCase();
    
    if (['txt', 'csv'].includes(ext)) {
      return await fileData.text();
    } else if (ext === 'pdf') {
      // Check if PDF.js is loaded
      if (typeof pdfjsLib === 'undefined') {
        console.warn('PDF.js not loaded, skipping PDF text extraction');
        return ''; // Skip PDF processing if PDF.js is not available
      }
      
      try {
        const arrayBuffer = await fileData.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        let text = '';
        
        // Process each page
        for (let i = 1; i <= pdf.numPages; i++) {
          try {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            const pageText = content.items
              .filter(item => item.str && item.str.trim() !== '')
              .map(item => item.str)
              .join(' ');
            
            if (pageText) {
              text += pageText + '\n';
            }
          } catch (pageError) {
            console.warn(`Error processing page ${i} of ${file.name}:`, pageError);
            continue; // Skip to next page on error
          }
        }
        
        return text;
      } catch (pdfError) {
        console.error(`Error processing PDF ${file.name}:`, pdfError);
        return ''; // Return empty string if PDF processing fails
      }
    } else if (['doc', 'docx', 'xls', 'xlsx'].includes(ext)) {
      // For Office files, we'll just read as text for now
      // In a real app, you'd want to use a library like mammoth or docx
      return await fileData.text();
    }
    
    return ''; // Return empty string for unsupported file types
  } catch (error) {
    console.error(`Error reading file ${file?.name || 'unknown'}:`, error);
    return ''; // Return empty string on any error
  }
}

// Helper function to find matches in content
function findMatches(content) {
  if (!content || !templateWords.length) return [];
  
  const matches = [];
  const caseInsensitive = caseInsensitiveCheckbox?.checked;
  
  templateWords.forEach(term => {
    const regex = new RegExp(
      caseInsensitive ? term : `\\b${term}\\b`, 
      caseInsensitive ? 'gi' : 'g'
    );
    
    let match;
    while ((match = regex.exec(content)) !== null) {
      matches.push({
        term: term,
        index: match.index,
        context: getMatchContext(content, match.index, term.length)
      });
    }
  });
  
  return matches;
}

// Helper function to get context around a match
function getMatchContext(content, index, length, contextLength = 50) {
  const start = Math.max(0, index - contextLength);
  const end = Math.min(content.length, index + length + contextLength);
  let context = content.substring(start, end);
  
  if (start > 0) context = '...' + context;
  if (end < content.length) context = context + '...';
  
  return context;
}

// Generate search report
async function generateReport() {
  try {
    if (searchResults.length === 0) {
      reportPreview.innerHTML = '<p class="text-gray-500 dark:text-gray-400">No matches found.</p>';
      return;
    }
    
    // Create report HTML
    let reportHTML = `
      <div class="mb-4">
        <h3 class="text-lg font-medium text-gray-900 dark:text-white mb-2">Search Results</h3>
        <p class="text-sm text-gray-600 dark:text-gray-300">
          Found ${matchesFound} matches in ${searchResults.length} files.
        </p>
      </div>
    `;
    
    // Add results
    searchResults.forEach(result => {
      reportHTML += `
        <div class="mb-6 border-b border-gray-200 dark:border-gray-700 pb-4">
          <h4 class="font-medium text-gray-900 dark:text-white">${result.path}</h4>
          <div class="mt-2 space-y-2">
      `;
      
      result.matches.forEach(match => {
        reportHTML += `
          <div class="p-3 bg-gray-50 dark:bg-gray-800 rounded-md">
            <span class="text-sm text-gray-600 dark:text-gray-400">Found: </span>
            <span class="font-mono bg-yellow-100 dark:bg-yellow-900 px-1 rounded">${escapeHtml(match.term)}</span>
            <div class="mt-1 text-sm text-gray-700 dark:text-gray-300">${highlightMatch(escapeHtml(match.context), match.term)}</div>
          </div>
        `;
      });
      
      reportHTML += `</div></div>`;
    });
    
    // Update the report preview
    reportPreview.innerHTML = reportHTML;
    
    // Enable download button
    downloadReportBtn.disabled = false;
    
  } catch (err) {
    console.error('Error generating report:', err);
    log(`Error generating report: ${err.message}`);
    reportPreview.innerHTML = `<p class="text-red-500">Error generating report: ${escapeHtml(err.message)}</p>`;
  }
}

// Helper function to escape HTML
function escapeHtml(unsafe) {
  return unsafe
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

// Helper function to highlight matches in text
function highlightMatch(text, term) {
  if (!term) return text;
  const regex = new RegExp(`(${escapeRegExp(term)})`, 'gi');
  return text.replace(regex, '<span class="bg-yellow-200 dark:bg-yellow-800 font-medium">$1</span>');
}

// Helper function to escape regex special characters
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}