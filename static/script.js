// Core UI State Management
let currentMode = 'pdf2docx'; // Default Mode
let queue = [];
let historyList = [];

// DOM Elements
const togglePdf2Docx = document.getElementById('toggle-pdf2docx');
const toggleDocx2Pdf = document.getElementById('toggle-docx2pdf');
const dropZone = document.getElementById('drop-zone');
const fileSelector = document.getElementById('file-selector');
const browseLink = document.querySelector('.browse-link');
const formatHint = document.getElementById('format-hint');

const queueSection = document.getElementById('queue-section');
const queueListContainer = document.getElementById('queue-list-container');
const queueCountBadge = document.getElementById('queue-count');
const convertAllBtn = document.getElementById('convert-all-btn');
const clearCompletedBtn = document.getElementById('clear-completed-btn');

const historySection = document.getElementById('history-section');
const historyListContainer = document.getElementById('history-list-container');
const historyCountBadge = document.getElementById('history-count');

const toast = document.getElementById('toast');
const toastMessage = document.querySelector('.toast-message');

// --- 1. Mode Switching Interface ---
function setConversionMode(mode) {
    currentMode = mode;
    if (mode === 'pdf2docx') {
        togglePdf2Docx.classList.add('active');
        toggleDocx2Pdf.classList.remove('active');
        formatHint.textContent = 'Supports PDF documents up to 50MB';
    } else {
        toggleDocx2Pdf.classList.add('active');
        togglePdf2Docx.classList.remove('active');
        formatHint.textContent = 'Supports Word documents (.docx/.doc) up to 50MB';
    }
}

togglePdf2Docx.addEventListener('click', () => setConversionMode('pdf2docx'));
toggleDocx2Pdf.addEventListener('click', () => setConversionMode('docx2pdf'));

// --- 2. File Selection & Drag-and-Drop Handlers ---
browseLink.addEventListener('click', (e) => {
    e.stopPropagation();
    fileSelector.click();
});

dropZone.addEventListener('click', () => {
    fileSelector.click();
});

fileSelector.addEventListener('change', (e) => {
    handleSelectedFiles(e.target.files);
    fileSelector.value = ''; // Reset input so same file can be re-selected
});

// Drag over styles
['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('dragover');
    }, false);
});

['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('dragover');
    }, false);
});

dropZone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    handleSelectedFiles(dt.files);
});

// --- 3. Queue Management Logic ---
function handleSelectedFiles(filesList) {
    const addedFiles = Array.from(filesList);
    let rejectedCount = 0;
    
    addedFiles.forEach(file => {
        const ext = file.name.split('.').pop().toLowerCase();
        let fileMode = currentMode;
        
        // Intelligent auto-routing based on file extension
        if (ext === 'pdf') {
            fileMode = 'pdf2docx';
        } else if (['docx', 'doc'].includes(ext)) {
            fileMode = 'docx2pdf';
        } else {
            rejectedCount++;
            return; // Reject unsupported extensions
        }

        // Check file size (50MB Limit)
        if (file.size > 50 * 1024 * 1024) {
            showToast(`File "${file.name}" exceeds the 50MB limit`, 'error');
            return;
        }

        // Prevent duplicate file staging in active queue
        const isDuplicate = queue.some(item => item.file.name === file.name && item.file.size === file.size);
        if (isDuplicate) return;

        // Stage into queue
        const queueItem = {
            id: 'qi_' + Math.random().toString(36).substr(2, 9),
            file: file,
            mode: fileMode,
            status: 'queued',
            progress: 0,
            fileId: null,
            outputFilename: null,
            error: null
        };
        queue.push(queueItem);
    });

    if (rejectedCount > 0) {
        showToast(`Skipped ${rejectedCount} unsupported file(s). Use PDFs or Word docs.`, 'error');
    }

    if (queue.length > 0) {
        renderQueue();
    }
}

function removeQueueItem(itemId) {
    // Prevent removal of actively converting items
    const item = queue.find(i => i.id === itemId);
    if (item && item.status === 'converting') {
        showToast("Cannot remove a file while it is converting", "error");
        return;
    }
    
    queue = queue.filter(item => item.id !== itemId);
    renderQueue();
}

function clearCompleted() {
    queue = queue.filter(item => item.status !== 'completed' && item.status !== 'error');
    renderQueue();
}

clearCompletedBtn.addEventListener('click', clearCompleted);

// --- 4. Queue Conversion Loop ---
async function startQueueConversion() {
    const itemsToConvert = queue.filter(item => item.status === 'queued' || item.status === 'error');
    
    if (itemsToConvert.length === 0) {
        showToast("No files ready in the queue to convert", "error");
        return;
    }

    // Disable conversion triggers during active runs
    convertAllBtn.disabled = true;
    convertAllBtn.innerHTML = `
        <span>Converting...</span>
        <svg class="spinner-icon animate-spin" viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="3">
            <circle cx="12" cy="12" r="10" stroke-opacity="0.25"></circle>
            <path d="M4 12a8 8 0 0 1 8-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 0 1 4 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
        </svg>
    `;

    for (const item of itemsToConvert) {
        await convertSingleItem(item);
    }

    // Re-enable interface buttons when completed
    convertAllBtn.disabled = false;
    convertAllBtn.innerHTML = `
        <span>Convert Queue</span>
        <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <polygon points="5 3 19 12 5 21 5 3"></polygon>
        </svg>
    `;
    
    // Toggle completed clear buttons visibility
    const completedItems = queue.filter(item => item.status === 'completed' || item.status === 'error');
    if (completedItems.length > 0) {
        clearCompletedBtn.classList.remove('hidden');
    }
}

convertAllBtn.addEventListener('click', startQueueConversion);

// Convert single file item via local API
async function convertSingleItem(item) {
    item.status = 'converting';
    item.progress = 5;
    renderQueue();

    // Progress bar smooth staging simulation
    let progressInterval = setInterval(() => {
        if (item.progress < 85) {
            item.progress += Math.floor(Math.random() * 8) + 2;
            updateProgressDOM(item.id, item.progress);
        }
    }, 450);

    const formData = new FormData();
    formData.append('file', item.file);
    formData.append('direction', item.mode);

    try {
        const response = await fetch('/api/convert', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        clearInterval(progressInterval);
        
        if (response.ok && data.success) {
            item.status = 'completed';
            item.progress = 100;
            item.fileId = data.fileId;
            item.outputFilename = data.filename;
            
            // Add entry into session history log
            addToHistory(data.filename, data.fileId, item.mode);
            showToast(`Converted: ${item.file.name}`, 'success');
        } else {
            item.status = 'error';
            item.error = data.error || "Conversion failed";
            showToast(item.error, 'error');
        }
    } catch (err) {
        clearInterval(progressInterval);
        item.status = 'error';
        item.error = "Connection lost or local process aborted";
        showToast(item.error, 'error');
    }

    renderQueue();
}

function addToHistory(filename, fileId, mode) {
    historyList.unshift({
        id: 'hist_' + Date.now(),
        filename: filename,
        fileId: fileId,
        mode: mode,
        timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
    });
    renderHistory();
}

// --- 5. DOM Rendering Systems ---

function renderQueue() {
    if (queue.length === 0) {
        queueSection.classList.add('hidden');
        clearCompletedBtn.classList.add('hidden');
        return;
    }

    queueSection.classList.remove('hidden');
    queueCountBadge.textContent = `${queue.length} File${queue.length > 1 ? 's' : ''}`;
    
    // Check if we should show standard action btn or clear btn
    const processing = queue.some(item => item.status === 'converting');
    convertAllBtn.style.display = processing ? 'none' : 'flex';
    
    // Clear completed is shown if some conversions exist and no active runs
    const completedCount = queue.filter(item => item.status === 'completed' || item.status === 'error').length;
    if (completedCount > 0 && !processing) {
        clearCompletedBtn.classList.remove('hidden');
    } else {
        clearCompletedBtn.classList.add('hidden');
    }

    queueListContainer.innerHTML = queue.map(item => {
        const isPDF = item.mode === 'pdf2docx';
        const fileIconClass = isPDF ? 'icon-pdf' : 'icon-docx';
        const targetType = isPDF ? 'DOCX' : 'PDF';
        const fileExt = isPDF ? 'pdf' : 'docx';
        
        let statusBadgeHTML = '';
        let actionHTML = '';
        
        if (item.status === 'queued') {
            statusBadgeHTML = `<span class="status-badge status-queued">Queued</span>`;
            actionHTML = `
                <button class="remove-btn" onclick="removeQueueItem('${item.id}')" title="Remove file">
                    <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18"></line>
                        <line x1="6" y1="6" x2="18" y2="18"></line>
                    </svg>
                </button>
            `;
        } else if (item.status === 'converting') {
            statusBadgeHTML = `<span class="status-badge status-converting">Converting</span>`;
            actionHTML = `
                <svg class="animate-spin" viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="var(--secondary)" stroke-width="3">
                    <circle cx="12" cy="12" r="10" stroke-opacity="0.25"></circle>
                    <path d="M4 12a8 8 0 0 1 8-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 0 1 4 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                </svg>
            `;
        } else if (item.status === 'completed') {
            statusBadgeHTML = `<span class="status-badge status-completed">Completed</span>`;
            actionHTML = `
                <a href="/api/download/${item.fileId}" class="download-btn" title="Download converted file" download>
                    <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                        <polyline points="7 10 12 15 17 10"></polyline>
                        <line x1="12" y1="15" x2="12" y2="3"></line>
                    </svg>
                </a>
            `;
        } else if (item.status === 'error') {
            statusBadgeHTML = `<span class="status-badge status-error" title="${item.error}">Error</span>`;
            actionHTML = `
                <button class="remove-btn" onclick="removeQueueItem('${item.id}')" title="Clear error">
                    <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18"></line>
                        <line x1="6" y1="6" x2="18" y2="18"></line>
                    </svg>
                </button>
            `;
        }

        return `
            <div class="queue-item" id="item_${item.id}">
                <div class="item-meta">
                    <div class="file-info">
                        <div class="file-icon ${fileIconClass}">
                            <span style="font-size: 0.65rem; font-weight: 800; font-family: var(--font-heading);">${fileExt.toUpperCase()}</span>
                        </div>
                        <div class="file-details">
                            <span class="file-name" title="${item.file.name}">${item.file.name}</span>
                            <span class="file-size">${formatBytes(item.file.size)} &bull; to ${targetType}</span>
                        </div>
                    </div>
                    <div class="status-container">
                        ${statusBadgeHTML}
                        ${actionHTML}
                    </div>
                </div>
                ${item.status === 'converting' || item.status === 'completed' ? `
                    <div class="progress-container">
                        <div class="progress-bar" id="progress_${item.id}" style="width: ${item.progress}%"></div>
                    </div>
                ` : ''}
                ${item.status === 'error' ? `
                    <span style="color: #fca5a5; font-size: 0.75rem; margin-top: 0.25rem;">${item.error}</span>
                ` : ''}
            </div>
        `;
    }).join('');
}

function updateProgressDOM(id, val) {
    const bar = document.getElementById(`progress_${id}`);
    if (bar) {
        bar.style.width = `${val}%`;
    }
}

function renderHistory() {
    if (historyList.length === 0) {
        historySection.classList.add('hidden');
        return;
    }

    historySection.classList.remove('hidden');
    historyCountBadge.textContent = `${historyList.length} Converted`;

    historyListContainer.innerHTML = historyList.map(item => {
        const isPDF = item.mode === 'pdf2docx';
        const fileIconClass = isPDF ? 'icon-docx' : 'icon-pdf'; // Icon of the OUT file
        const fileExt = isPDF ? 'docx' : 'pdf';

        return `
            <div class="history-item">
                <div class="item-meta">
                    <div class="file-info">
                        <div class="file-icon ${fileIconClass}">
                            <span style="font-size: 0.65rem; font-weight: 800; font-family: var(--font-heading);">${fileExt.toUpperCase()}</span>
                        </div>
                        <div class="file-details">
                            <span class="file-name" title="${item.filename}">${item.filename}</span>
                            <span class="file-size">Completed at ${item.timestamp}</span>
                        </div>
                    </div>
                    <a href="/api/download/${item.fileId}" class="download-btn" title="Download converted file" download>
                        <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                            <polyline points="7 10 12 15 17 10"></polyline>
                            <line x1="12" y1="15" x2="12" y2="3"></line>
                        </svg>
                    </a>
                </div>
            </div>
        `;
    }).join('');
}

// --- 6. Interface Utilities ---

function formatBytes(bytes, decimals = 2) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

let toastTimeout;
function showToast(message, type = 'success') {
    clearTimeout(toastTimeout);
    
    toastMessage.textContent = message;
    toast.className = 'toast show';
    
    if (type === 'error') {
        toast.classList.add('error');
    } else {
        toast.classList.add('success');
    }

    toastTimeout = setTimeout(() => {
        toast.classList.remove('show');
    }, 4500);
}
