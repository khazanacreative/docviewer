/**
 * Excel Document Viewer - Logic Script
 * Developed for premium UI/UX, instant responses, and 100% client-side execution.
 */

// State Management
let appData = {
  filename: '',
  documents: []
};
let filteredDocuments = [];
let activeIndex = -1;

// IndexedDB Database Utilities
const DB_NAME = 'ExcelViewerDB';
const DB_VERSION = 1;
const STORE_NAME = 'excel_files';

function initDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: 'name' });
      }
    };
    request.onsuccess = (e) => resolve(e.target.result);
    request.onerror = (e) => reject(e.target.error);
  });
}

function saveFileToStorage(name, arrayBuffer, documents) {
  return initDB().then(db => {
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, 'readwrite');
      const store = transaction.objectStore(STORE_NAME);
      const dataRecord = {
        name: name,
        data: arrayBuffer,
        documents: documents,
        updatedAt: Date.now()
      };
      const request = store.put(dataRecord);
      request.onsuccess = () => resolve();
      request.onerror = (e) => reject(e.target.error);
    });
  });
}

function getFileFromStorage(name) {
  return initDB().then(db => {
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, 'readonly');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.get(name);
      request.onsuccess = (e) => resolve(e.target.result);
      request.onerror = (e) => reject(e.target.error);
    });
  });
}

function getAllFilesFromStorage() {
  return initDB().then(db => {
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, 'readonly');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.getAll();
      request.onsuccess = (e) => {
        const files = e.target.result || [];
        files.sort((a, b) => b.updatedAt - a.updatedAt);
        resolve(files);
      };
      request.onerror = (e) => reject(e.target.error);
    });
  });
}

function deleteFileFromStorage(name) {
  return initDB().then(db => {
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE_NAME, 'readwrite');
      const store = transaction.objectStore(STORE_NAME);
      const request = store.delete(name);
      request.onsuccess = () => resolve();
      request.onerror = (e) => reject(e.target.error);
    });
  });
}

// DOM Elements
const landingContainer = document.getElementById('landing-container');
const appContainer = document.getElementById('app-container');
const layoutWrapper = document.getElementById('layout-wrapper');
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const loadDemoBtn = document.getElementById('load-demo-btn');
const changeFileBtn = document.getElementById('change-file-btn');
const activeFilename = document.getElementById('active-filename');
const dragOverlay = document.getElementById('drag-overlay');
const savedFilesSection = document.getElementById('saved-files-section');
const savedFilesList = document.getElementById('saved-files-list');

const searchInput = document.getElementById('search-input');
const clearSearchBtn = document.getElementById('clear-search-btn');
const documentList = document.getElementById('document-list');
const noSearchResults = document.getElementById('no-search-results');

const contentHeader = document.getElementById('content-header');
const contentBodyWrapper = document.getElementById('content-body-wrapper');
const contentBody = document.getElementById('content-body');
const emptyState = document.getElementById('empty-state');
const backBtn = document.getElementById('back-btn');
const copyBtn = document.getElementById('copy-btn');
const copyBtnText = document.getElementById('copy-btn-text');
const docTitle = document.getElementById('doc-title');
const docMeta = document.getElementById('doc-meta');

const toast = document.getElementById('toast');
const toastMessage = document.getElementById('toast-message');

/* ==========================================================================
   INITIALIZATION & PERSISTENCE
   ========================================================================== */

function showLandingPage() {
  landingContainer.classList.remove('hidden');
  appContainer.classList.add('hidden');
  updateSavedFilesList();
  lucide.createIcons();
}

function updateSavedFilesList() {
  if (!savedFilesSection || !savedFilesList) return;
  getAllFilesFromStorage()
    .then(files => {
      if (files.length > 0) {
        savedFilesSection.classList.remove('hidden');
        savedFilesList.innerHTML = '';
        
        files.forEach(file => {
          const item = document.createElement('div');
          item.className = 'saved-file-item';
          
          const formattedDate = new Date(file.updatedAt).toLocaleString('id-ID', {
            day: 'numeric',
            month: 'short',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
          });
          
          item.innerHTML = `
            <div class="saved-file-info">
              <i data-lucide="file-spreadsheet"></i>
              <div class="saved-file-text">
                <span class="saved-file-name" title="${file.name}">${file.name}</span>
                <span class="saved-file-date">Diperbarui: ${formattedDate}</span>
              </div>
            </div>
            <div class="saved-file-actions">
              <button class="saved-file-btn delete-btn" title="Hapus berkas dari browser">
                <i data-lucide="trash-2"></i>
              </button>
            </div>
          `;
          
          // Click to load the file
          item.querySelector('.saved-file-info').addEventListener('click', () => {
            appData.filename = file.name;
            appData.documents = file.documents;
            localStorage.setItem('active_excel_file', file.name);
            setupApplication();
            showToast(`Berkas "${file.name}" berhasil dimuat!`);
          });
          
          // Click to delete the file
          item.querySelector('.delete-btn').addEventListener('click', (e) => {
            e.stopPropagation();
            if (confirm(`Apakah Anda yakin ingin menghapus berkas "${file.name}"?`)) {
              deleteFileFromStorage(file.name)
                .then(() => {
                  showToast(`Berkas "${file.name}" dihapus.`);
                  if (localStorage.getItem('active_excel_file') === file.name) {
                    localStorage.removeItem('active_excel_file');
                  }
                  updateSavedFilesList();
                })
                .catch(err => {
                  showToast(`Gagal menghapus berkas: ${err.message}`, true);
                });
            }
          });
          
          savedFilesList.appendChild(item);
        });
        
        lucide.createIcons();
      } else {
        savedFilesSection.classList.add('hidden');
        savedFilesList.innerHTML = '';
      }
    })
    .catch(err => {
      console.error('Gagal mengambil daftar berkas:', err);
    });
}

document.addEventListener('DOMContentLoaded', () => {
  // Check for active file in localStorage
  const activeFile = localStorage.getItem('active_excel_file');
  if (activeFile) {
    getFileFromStorage(activeFile)
      .then(fileRecord => {
        if (fileRecord && fileRecord.documents && fileRecord.documents.length > 0) {
          appData.filename = fileRecord.name;
          appData.documents = fileRecord.documents;
          setupApplication();
        } else {
          showLandingPage();
        }
      })
      .catch(e => {
        console.error('Gagal memuat file aktif dari IndexedDB:', e);
        showLandingPage();
      });
  } else {
    showLandingPage();
  }
});

/* ==========================================================================
   FILE UPLOAD & DRAG/DROP EVENTS
   ========================================================================== */

// Change File Button (sidebar header)
if (changeFileBtn) {
  changeFileBtn.addEventListener('click', () => {
    // Reset state
    localStorage.removeItem('active_excel_file');
    appData = { filename: '', documents: [] };
    filteredDocuments = [];
    activeIndex = -1;
    
    // Reset UI
    searchInput.value = '';
    clearSearchBtn.classList.add('hidden');
    documentList.innerHTML = '';
    noSearchResults.classList.add('hidden');
    
    contentHeader.classList.add('hidden');
    contentBodyWrapper.classList.add('hidden');
    emptyState.classList.remove('hidden');
    layoutWrapper.classList.remove('show-detail');
    
    showLandingPage();
  });
}

// Landing Page Upload Trigger
if (fileInput) {
  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFile(e.target.files[0]);
    }
  });
}

// Full-screen Drag & Drop Handling
let dragCounter = 0;

window.addEventListener('dragenter', (e) => {
  e.preventDefault();
  e.stopPropagation();
  dragCounter++;
  if (dragOverlay) {
    dragOverlay.classList.remove('hidden');
  }
});

window.addEventListener('dragover', (e) => {
  e.preventDefault();
  e.stopPropagation();
});

window.addEventListener('dragleave', (e) => {
  e.preventDefault();
  e.stopPropagation();
  dragCounter--;
  if (dragCounter === 0 && dragOverlay) {
    dragOverlay.classList.add('hidden');
  }
});

window.addEventListener('drop', (e) => {
  e.preventDefault();
  e.stopPropagation();
  dragCounter = 0;
  if (dragOverlay) {
    dragOverlay.classList.add('hidden');
  }
  
  const dt = e.dataTransfer;
  const files = dt.files;
  if (files.length > 0) {
    handleFile(files[0]);
  }
});

// Load Demo File
if (loadDemoBtn) {
  loadDemoBtn.addEventListener('click', () => {
    loadDemoBtn.disabled = true;
    loadDemoBtn.innerHTML = '<i data-lucide="loader-2" class="btn-icon animate-spin"></i> Memuat berkas demo...';
    lucide.createIcons();
    
    fetch('demo.csv')
      .then(response => {
        if (!response.ok) throw new Error('File demo tidak ditemukan');
        return response.arrayBuffer();
      })
      .then(buffer => {
        processExcelBuffer(buffer, 'demo.csv');
      })
      .catch(error => {
        showToast('Gagal memuat file demo: ' + error.message, true);
        loadDemoBtn.disabled = false;
        loadDemoBtn.innerHTML = '<i data-lucide="play" class="btn-icon"></i> Coba dengan File Demo';
        lucide.createIcons();
      });
  });
}

/* ==========================================================================
   EXCEL PARSING WITH SHEETJS
   ========================================================================== */
function handleFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    processExcelBuffer(e.target.result, file.name);
  };
  reader.onerror = () => {
    showToast('Gagal membaca file', true);
  };
  reader.readAsArrayBuffer(file);
}

function processExcelBuffer(buffer, filename, isQuiet = false) {
  try {
    const data = new Uint8Array(buffer);
    const workbook = XLSX.read(data, { type: 'array' });
    
    if (workbook.SheetNames.length === 0) {
      throw new Error('File Excel kosong atau tidak valid.');
    }
    
    // Parse first sheet
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    
    // Read raw grid to ensure precise column layout
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (rows.length === 0) {
      throw new Error('Sheet pertama kosong.');
    }
    
    // Check if the first row contains column header names to skip them
    let startIdx = 0;
    const firstRow = rows[0];
    const isHeader = firstRow.some(cell => {
      if (cell === null || cell === undefined) return false;
      const s = String(cell).toLowerCase().trim();
      return s === 'judul' || s === 'keterangan' || s === 'konten' || s === 'title' || s === 'description' || s === 'content';
    });
    
    if (isHeader) {
      startIdx = 1;
    }
    
    // Parse documents matching: Col 0 -> Judul, Col 1 -> Keterangan, Col 2 -> Konten
    const docs = [];
    for (let i = startIdx; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;
      
      const judul = row[0] !== undefined ? String(row[0]).trim() : '';
      const keterangan = row[1] !== undefined ? String(row[1]).trim() : '';
      const konten = row[2] !== undefined ? String(row[2]) : '';
      
      // Push if at least one column has data
      if (judul || keterangan || konten) {
        docs.push({ judul, keterangan, konten });
      }
    }
    
    if (docs.length === 0) {
      throw new Error('Tidak ditemukan data dokumen pada file Excel. Pastikan kolom A, B, atau C berisi data.');
    }
    
    // Save to State
    appData.filename = filename;
    appData.documents = docs;
    
    // Save to IndexedDB
    saveFileToStorage(filename, buffer, docs)
      .then(() => {
        localStorage.setItem('active_excel_file', filename);
        updateSavedFilesList();
      })
      .catch(err => {
        console.error('Gagal menyimpan berkas ke IndexedDB:', err);
      });
    
    // Setup and show application
    setupApplication();
    if (!isQuiet) {
      showToast('File berhasil dimuat!');
    }
    
  } catch (error) {
    showToast(error.message, true);
    // Reset demo button if active
    if (loadDemoBtn) {
      loadDemoBtn.disabled = false;
      loadDemoBtn.innerHTML = '<i data-lucide="play" class="btn-icon"></i> Coba dengan File Demo';
      lucide.createIcons();
    }
  }
}


/* ==========================================================================
   APPLICATION SCREEN SETUP & LOGIC
   ========================================================================== */
function setupApplication() {
  activeFilename.textContent = appData.filename;
  filteredDocuments = [...appData.documents];
  
  // Render sidebar document list
  renderDocumentList();
  
  // Transition views
  landingContainer.classList.add('hidden');
  appContainer.classList.remove('hidden');
  
  // Reset demo button loading state
  loadDemoBtn.disabled = false;
  loadDemoBtn.innerHTML = '<i data-lucide="play" class="btn-icon"></i> Coba dengan File Demo';
  
  // Setup Lucide icons in application
  lucide.createIcons();
}

// Render Document List
function renderDocumentList() {
  documentList.innerHTML = '';
  
  if (filteredDocuments.length === 0) {
    noSearchResults.classList.remove('hidden');
    return;
  }
  
  noSearchResults.classList.add('hidden');
  
  filteredDocuments.forEach((doc, idx) => {
    // Find original index from appData.documents to maintain active state properly
    const originalIdx = appData.documents.findIndex(
      originalDoc => originalDoc.judul === doc.judul && 
                     originalDoc.keterangan === doc.keterangan && 
                     originalDoc.konten === doc.konten
    );
    
    const card = document.createElement('button');
    card.type = 'button';
    card.className = `doc-item-card ${originalIdx === activeIndex ? 'active' : ''}`;
    card.id = `doc-card-${originalIdx}`;
    
    // Fallbacks if empty
    const displayTitle = doc.judul || 'Tanpa Judul';
    const displayDesc = doc.keterangan || '(Tidak ada keterangan)';
    
    card.innerHTML = `
      <h3 class="doc-item-title">${displayTitle}</h3>
      <p class="doc-item-desc">${displayDesc}</p>
    `;
    
    card.addEventListener('click', () => selectDocument(originalIdx));
    documentList.appendChild(card);
  });
}

// Select Document
function selectDocument(originalIdx) {
  activeIndex = originalIdx;
  
  // Update selected class in sidebar cards
  const allCards = document.querySelectorAll('.doc-item-card');
  allCards.forEach(card => card.classList.remove('active'));
  
  const activeCard = document.getElementById(`doc-card-${originalIdx}`);
  if (activeCard) {
    activeCard.classList.add('active');
  }
  
  // Load content
  const doc = appData.documents[originalIdx];
  docTitle.textContent = doc.judul || 'Tanpa Judul';
  docMeta.textContent = doc.keterangan || '(Tidak ada keterangan)';
  
  // Render text with WhatsApp formatting
  contentBody.innerHTML = formatWhatsAppText(doc.konten);
  
  // Trigger fade-in animation
  contentBody.classList.remove('fade-in');
  void contentBody.offsetWidth; // force reflow
  contentBody.classList.add('fade-in');
  
  // Hide empty state and show reading area
  emptyState.classList.add('hidden');
  contentHeader.classList.remove('hidden');
  contentBodyWrapper.classList.remove('hidden');
  
  // Scroll content to top
  contentBodyWrapper.scrollTop = 0;
  
  // Mobile UI sliding transition: show detail panel
  layoutWrapper.classList.add('show-detail');
}

/* ==========================================================================
   WHATSAPP TEXT FORMATTING UTILITY
   ========================================================================== */
function formatWhatsAppText(text) {
  if (!text) return '<span class="text-muted">(Tidak ada konten teks)</span>';
  
  // 1. Escape HTML to prevent XSS
  let formatted = text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
  
  // 2. Format URLs: match http:// or https:// and format as clickable link
  const urlRegex = /(https?:\/\/[^\s]+)/g;
  formatted = formatted.replace(urlRegex, (url) => {
    return `<a href="${url}" target="_blank" rel="noopener noreferrer">${url}</a>`;
  });
  
  // 3. Format Bold: *teks*
  formatted = formatted.replace(/\*([^*]+)\*/g, '<strong>$1</strong>');
  
  // 4. Format Italic: _teks_
  formatted = formatted.replace(/_([^_]+)_/g, '<em>$1</em>');
  
  // 5. Format Strikethrough: ~teks~
  formatted = formatted.replace(/~([^~]+)~/g, '<del>$1</del>');
  
  return formatted;
}

/* ==========================================================================
   SEARCH & FILTERS
   ========================================================================== */
searchInput.addEventListener('input', (e) => {
  const query = e.target.value.toLowerCase().trim();
  
  if (query.length > 0) {
    clearSearchBtn.classList.remove('hidden');
    
    filteredDocuments = appData.documents.filter(doc => {
      const titleMatch = doc.judul.toLowerCase().includes(query);
      const descMatch = doc.keterangan.toLowerCase().includes(query);
      // We can also allow searching in content text
      const contentMatch = doc.konten.toLowerCase().includes(query);
      return titleMatch || descMatch || contentMatch;
    });
  } else {
    clearSearchBtn.classList.add('hidden');
    filteredDocuments = [...appData.documents];
  }
  
  renderDocumentList();
});

// Clear Search Input
clearSearchBtn.addEventListener('click', () => {
  searchInput.value = '';
  clearSearchBtn.classList.add('hidden');
  filteredDocuments = [...appData.documents];
  renderDocumentList();
  searchInput.focus();
});

/* ==========================================================================
   COPY TO CLIPBOARD & MOBILE NAVIGATION
   ========================================================================== */

// Copy text to clipboard
copyBtn.addEventListener('click', () => {
  if (activeIndex === -1) return;
  
  const doc = appData.documents[activeIndex];
  const textToCopy = doc.konten || '';
  
  navigator.clipboard.writeText(textToCopy)
    .then(() => {
      // Visual feedback on copy button
      copyBtn.classList.add('copy-success');
      copyBtnText.textContent = 'Tersalin!';
      
      // Update icon dynamically
      const copyIcon = copyBtn.querySelector('i');
      if (copyIcon) {
        copyIcon.setAttribute('data-lucide', 'check');
        lucide.createIcons();
      }
      
      // Show Toast Notification
      showToast('Konten dokumen berhasil disalin ke clipboard!');
      
      // Restore state after 2 seconds
      setTimeout(() => {
        copyBtn.classList.remove('copy-success');
        copyBtnText.textContent = 'Salin Teks';
        if (copyIcon) {
          copyIcon.setAttribute('data-lucide', 'copy');
          lucide.createIcons();
        }
      }, 2000);
    })
    .catch(err => {
      console.error('Gagal menyalin teks: ', err);
      showToast('Gagal menyalin teks ke clipboard.', true);
    });
});

// Mobile Back Button: Slide detail panel out
backBtn.addEventListener('click', () => {
  layoutWrapper.classList.remove('show-detail');
});

/* ==========================================================================
   TOAST SYSTEM
   ========================================================================== */
let toastTimeout;
function showToast(message, isError = false) {
  clearTimeout(toastTimeout);
  
  toastMessage.textContent = message;
  const icon = toast.querySelector('.toast-icon');
  
  if (isError) {
    toast.style.borderLeft = '4px solid #ef4444';
    if (icon) {
      icon.setAttribute('data-lucide', 'alert-circle');
      icon.style.color = '#ef4444';
    }
  } else {
    toast.style.borderLeft = 'none';
    if (icon) {
      icon.setAttribute('data-lucide', 'check-circle');
      icon.style.color = '#10b981';
    }
  }
  
  lucide.createIcons();
  toast.classList.remove('hidden');
  
  toastTimeout = setTimeout(() => {
    toast.classList.add('hidden');
  }, 3000);
}
