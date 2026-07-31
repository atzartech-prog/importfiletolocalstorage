// StorageLearn App JavaScript Logic

// Global states
let selectedFileType = 'json';
let selectedFile = null;
let currentWorkbook = null;
let renderMode = 'rendered'; // 'rendered' or 'raw'

// Code snippets database
const codeSnippets = {
    'code-json': `// 1. Membaca & Menyimpan File JSON ke LocalStorage
function handleJSONUpload(file, storageKey) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            // Membaca konten string dan menguji validitas format JSON
            const rawContent = e.target.result;
            const parsedData = JSON.parse(rawContent);
            
            // Konversi kembali ke string terkompresi sebelum masuk LocalStorage
            const stringifiedData = JSON.stringify(parsedData);
            localStorage.setItem(storageKey, stringifiedData);
            
            alert('Sukses menyimpan JSON ke LocalStorage!');
            updateStorageMonitor();
        } catch (err) {
            console.error('Format JSON tidak valid!', err);
        }
    };
    
    // Membaca file sebagai teks biasa
    reader.readAsText(file);
}`,
    'code-excel': `// 2. Membaca & Menyimpan Excel (.xlsx/.xls) Menggunakan SheetJS
function handleExcelUpload(file, storageKey, selectedSheetName) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        // 1. Membaca file Excel sebagai buffer biner mentah (ArrayBuffer)
        const data = new Uint8Array(e.target.result);
        
        // 2. Parse menggunakan pustaka SheetJS (XLSX)
        const workbook = XLSX.read(data, { type: 'array' });
        
        // 3. Mengambil Sheet tertentu berdasarkan pilihan user
        const targetSheet = workbook.Sheets[selectedSheetName];
        
        // 4. Konversi baris-kolom Excel ke format Array of Objects JSON
        const jsonData = XLSX.utils.sheet_to_json(targetSheet);
        
        // 5. Stringify data dan simpan ke LocalStorage
        localStorage.setItem(storageKey, JSON.stringify(jsonData));
        
        alert(\`Sukses menyimpan data Sheet "\${selectedSheetName}" ke LocalStorage!\`);
        updateStorageMonitor();
    };
    
    // Excel adalah file biner, wajib dibaca sebagai ArrayBuffer
    reader.readAsArrayBuffer(file);
}`,
    'code-image': `// 3. Mengubah Gambar ke String Base64 untuk LocalStorage
function handleImageUpload(file, storageKey) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        // Hasil e.target.result berupa string panjang berformat:
        // "data:image/png;base64,iVBORw0KGgoAAA..."
        const base64String = e.target.result;
        
        // Simpan langsung ke LocalStorage
        localStorage.setItem(storageKey, base64String);
        
        alert('Gambar berhasil dikonversi ke Base64 & disimpan!');
        updateStorageMonitor();
    };
    
    // Membaca file sebagai DataURL (Base64 encoded string)
    reader.readAsDataURL(file);
}`,
    'code-csv': `// 4. Parser File CSV ke JSON & Simpan ke LocalStorage
function handleCSVUpload(file, storageKey) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        const text = e.target.result;
        
        // Parser CSV Sederhana (Split baris dan kolom)
        const lines = text.split(/\\r?\\n/).filter(line => line.trim() !== "");
        if (lines.length === 0) return;
        
        // Ambil header kolom di baris pertama
        const headers = lines[0].split(',').map(h => h.trim());
        const jsonData = [];
        
        // Iterasi baris data berikutnya
        for (let i = 1; i < lines.length; i++) {
            const cols = lines[i].split(',').map(c => c.trim());
            if (cols.length === headers.length) {
                const rowObj = {};
                headers.forEach((header, index) => {
                    rowObj[header] = cols[index];
                });
                jsonData.push(rowObj);
            }
        }
        
        // Simpan dalam format String JSON ke LocalStorage
        localStorage.setItem(storageKey, JSON.stringify(jsonData));
        updateStorageMonitor();
    };
    
    reader.readAsText(file);
}`
};

// Map file types to accepted formats, default keys, and label descriptions
const fileTypeConfigs = {
    'json': {
        accept: '.json',
        defaultKey: 'user_data_json',
        hint: 'Menerima format file: .json',
        codeSnippetKey: 'code-json'
    },
    'excel': {
        accept: '.xlsx,.xls',
        defaultKey: 'excel_inventory_data',
        hint: 'Menerima format file: .xlsx, .xls',
        codeSnippetKey: 'code-excel'
    },
    'csv': {
        accept: '.csv',
        defaultKey: 'sales_report_csv',
        hint: 'Menerima format file: .csv',
        codeSnippetKey: 'code-csv'
    },
    'image': {
        accept: 'image/*',
        defaultKey: 'app_logo_base64',
        hint: 'Menerima file gambar: .png, .jpg, .jpeg, .webp, .gif',
        codeSnippetKey: 'code-image'
    },
    'text': {
        accept: '.txt',
        defaultKey: 'config_notes_txt',
        hint: 'Menerima format file teks: .txt',
        codeSnippetKey: 'code-json' // fallback to basic text snippet
    }
};

// DOM Elements
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('file-input');
const fileInfoBox = document.getElementById('file-info-box');
const infoFilename = document.getElementById('info-filename');
const infoFilesize = document.getElementById('info-filesize');
const btnCancelFile = document.getElementById('btn-cancel-file');
const btnProcessFile = document.getElementById('btn-process-file');
const storageKeyInput = document.getElementById('storage-key');
const btnSuggestKey = document.getElementById('btn-suggest-key');
const excelSheetSelectorWrapper = document.getElementById('excel-sheet-selector-wrapper');
const excelSheetSelect = document.getElementById('excel-sheet-select');

const storageProgressBar = document.getElementById('storage-progress-bar');
const storagePercentageText = document.getElementById('storage-percentage-text');
const storageUsedText = document.getElementById('storage-used-text');
const localstorageKeysList = document.getElementById('localstorage-keys-list');
const btnClearStorage = document.getElementById('btn-clear-storage');
const btnResetApp = document.getElementById('btn-reset-app');

const selectPreviewKey = document.getElementById('select-preview-key');
const visualizerOutput = document.getElementById('visualizer-output');

// Modal Elements
const detailModal = document.getElementById('detail-modal');
const modalKeyName = document.getElementById('modal-key-name');
const modalKeySize = document.getElementById('modal-key-size');
const modalJsonContent = document.getElementById('modal-json-content');
const btnCloseModal = document.getElementById('btn-close-modal');
const btnCloseModalConfirm = document.getElementById('btn-close-modal-confirm');
const btnCopyModal = document.getElementById('btn-copy-modal');

// Init application
document.addEventListener('DOMContentLoaded', () => {
    lucide.createIcons();
    initEventListeners();
    setupDefaultState();
    loadDemoData(false); // Only populate if empty
    updateStorageMonitor();
    updateCodeDisplay('code-json');
});

// Setup default UI state
function setupDefaultState() {
    selectedFileType = 'json';
    updateUploaderUI();
}

// Event Listeners Registration
function initEventListeners() {
    // Type button selector click
    document.querySelectorAll('.type-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const btnEl = e.currentTarget;
            document.querySelectorAll('.type-btn').forEach(b => b.classList.remove('active'));
            btnEl.classList.add('active');
            selectedFileType = btnEl.getAttribute('data-type');
            
            resetFileSelection();
            updateUploaderUI();
            
            // Auto switch snippet tab if user clicks type
            const config = fileTypeConfigs[selectedFileType];
            if (config) {
                updateCodeDisplay(config.codeSnippetKey);
                document.getElementById('select-snippet-type').value = config.codeSnippetKey;
            }
        });
    });

    // Dropzone drag-drop behaviors
    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragleave', () => {
        dropzone.classList.remove('dragover');
    });

    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            handleFileSelect(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileSelect(e.target.files[0]);
        }
    });

    btnCancelFile.addEventListener('click', () => {
        resetFileSelection();
    });

    btnSuggestKey.addEventListener('click', () => {
        const config = fileTypeConfigs[selectedFileType];
        if (config) {
            const timeSuffix = Math.floor(Date.now() / 1000) % 10000;
            storageKeyInput.value = `${config.defaultKey}_${timeSuffix}`;
        }
    });

    btnProcessFile.addEventListener('click', () => {
        processAndSave();
    });

    // Storage monitor actions
    btnClearStorage.addEventListener('click', () => {
        if (confirm('Apakah Anda yakin ingin menghapus semua data LocalStorage untuk origin ini?')) {
            localStorage.clear();
            updateStorageMonitor();
            resetVisualizer();
            alert('LocalStorage berhasil dikosongkan!');
        }
    });

    btnResetApp.addEventListener('click', () => {
        if (confirm('Kembalikan aplikasi ke data demo bawaan? Ini akan membersihkan LocalStorage saat ini.')) {
            loadDemoData(true); // force reload demo data
            updateStorageMonitor();
            resetVisualizer();
            alert('Aplikasi direset ke data demo!');
        }
    });

    // Education tabs logic
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const targetTab = e.currentTarget.getAttribute('data-tab');
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            
            e.currentTarget.classList.add('active');
            document.getElementById(`tab-${targetTab}`).classList.add('active');
        });
    });

    // Snippet select change
    document.getElementById('select-snippet-type').addEventListener('change', (e) => {
        updateCodeDisplay(e.target.value);
    });

    // Visualizer selection and toggles
    selectPreviewKey.addEventListener('change', () => {
        renderSelectedData();
    });

    document.querySelectorAll('.btn-toggle').forEach(btn => {
        btn.addEventListener('click', (e) => {
            document.querySelectorAll('.btn-toggle').forEach(b => b.classList.remove('active'));
            e.currentTarget.classList.add('active');
            renderMode = e.currentTarget.getAttribute('data-mode');
            renderSelectedData();
        });
    });

    // Modal behavior
    const closeModal = () => {
        detailModal.classList.add('hidden');
    };

    btnCloseModal.addEventListener('click', closeModal);
    btnCloseModalConfirm.addEventListener('click', closeModal);
    detailModal.addEventListener('click', (e) => {
        if (e.target === detailModal) closeModal();
    });

    btnCopyModal.addEventListener('click', () => {
        navigator.clipboard.writeText(modalJsonContent.textContent)
            .then(() => {
                alert('Teks berhasil disalin ke papan klip!');
            })
            .catch(err => {
                console.error('Gagal menyalin text', err);
            });
    });
}

// Update file uploader dropzone configuration dynamically
function updateUploaderUI() {
    const config = fileTypeConfigs[selectedFileType];
    if (config) {
        fileInput.setAttribute('accept', config.accept);
        document.getElementById('file-accepted-formats').textContent = config.hint;
        storageKeyInput.value = config.defaultKey;
    }
}

// File Selection Handler
function handleFileSelect(file) {
    // Validate simple extensions if needed
    const config = fileTypeConfigs[selectedFileType];
    if (selectedFileType === 'json' && !file.name.endsWith('.json')) {
        alert('File bukan format JSON (.json)!');
        return;
    }
    if (selectedFileType === 'excel' && !file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        alert('File bukan format Excel (.xlsx, .xls)!');
        return;
    }
    if (selectedFileType === 'csv' && !file.name.endsWith('.csv')) {
        alert('File bukan format CSV (.csv)!');
        return;
    }
    if (selectedFileType === 'image' && !file.type.startsWith('image/')) {
        alert('File yang dipilih bukan file gambar!');
        return;
    }
    if (selectedFileType === 'text' && !file.name.endsWith('.txt')) {
        alert('Silakan pilih file teks berekstensi .txt!');
        return;
    }

    selectedFile = file;
    infoFilename.textContent = file.name;
    infoFilesize.textContent = formatBytes(file.size);
    fileInfoBox.classList.remove('hidden');
    btnProcessFile.removeAttribute('disabled');

    // If it's an Excel file, pre-read the sheets using SheetJS so the sheet selector can be filled
    if (selectedFileType === 'excel') {
        preParseExcelSheets(file);
    }
}

// Reset selected files
function resetFileSelection() {
    selectedFile = null;
    currentWorkbook = null;
    fileInput.value = '';
    fileInfoBox.classList.add('hidden');
    btnProcessFile.setAttribute('disabled', 'true');
    excelSheetSelectorWrapper.style.display = 'none';
    excelSheetSelect.innerHTML = '<option value="">-- Upload file Excel dulu --</option>';
}

// Pre-read workbook sheet names
function preParseExcelSheets(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            currentWorkbook = workbook;
            
            // Populate select
            excelSheetSelect.innerHTML = '';
            workbook.SheetNames.forEach(name => {
                const opt = document.createElement('option');
                opt.value = name;
                opt.textContent = name;
                excelSheetSelect.appendChild(opt);
            });
            
            excelSheetSelectorWrapper.style.display = 'block';
        } catch (err) {
            alert('Gagal mendeteksi sheets Excel! Pastikan file Excel tidak rusak.');
            console.error(err);
            resetFileSelection();
        }
    };
    reader.readAsArrayBuffer(file);
}

// Main save handler
function processAndSave() {
    if (!selectedFile) return;

    const storageKey = storageKeyInput.value.trim();
    if (!storageKey) {
        alert('Nama Key LocalStorage tidak boleh kosong!');
        return;
    }

    const reader = new FileReader();

    if (selectedFileType === 'json') {
        reader.onload = function(e) {
            try {
                const content = e.target.result;
                // Validate JSON syntax
                const parsed = JSON.parse(content);
                localStorage.setItem(storageKey, JSON.stringify(parsed));
                onSaveSuccess(storageKey);
            } catch (err) {
                alert('Error parsing JSON! Pastikan struktur file JSON Anda valid. Detail: ' + err.message);
            }
        };
        reader.readAsText(selectedFile);
    } 
    else if (selectedFileType === 'excel') {
        const sheetName = excelSheetSelect.value;
        if (!sheetName || !currentWorkbook) {
            alert('Silakan pilih sheet Excel terlebih dahulu!');
            return;
        }

        try {
            const worksheet = currentWorkbook.Sheets[sheetName];
            // SheetJS converts table to JSON Array of Objects
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            localStorage.setItem(storageKey, JSON.stringify(jsonData));
            onSaveSuccess(storageKey);
        } catch (err) {
            alert('Gagal memproses sheet Excel ke JSON: ' + err.message);
        }
    } 
    else if (selectedFileType === 'csv') {
        reader.onload = function(e) {
            try {
                const text = e.target.result;
                const parsedData = parseCSV(text);
                localStorage.setItem(storageKey, JSON.stringify(parsedData));
                onSaveSuccess(storageKey);
            } catch (err) {
                alert('Gagal memproses file CSV: ' + err.message);
            }
        };
        reader.readAsText(selectedFile);
    } 
    else if (selectedFileType === 'image') {
        reader.onload = function(e) {
            try {
                // base64 DataURL
                const base64String = e.target.result;
                localStorage.setItem(storageKey, base64String);
                onSaveSuccess(storageKey);
            } catch (err) {
                alert('Gagal merubah gambar ke Base64: ' + err.message);
            }
        };
        reader.readAsDataURL(selectedFile);
    } 
    else if (selectedFileType === 'text') {
        reader.onload = function(e) {
            try {
                const content = e.target.result;
                localStorage.setItem(storageKey, content);
                onSaveSuccess(storageKey);
            } catch (err) {
                alert('Gagal menyimpan file teks: ' + err.message);
            }
        };
        reader.readAsText(selectedFile);
    }
}

function onSaveSuccess(key) {
    alert(`Sukses! Data disimpan ke LocalStorage dengan key: "${key}"`);
    resetFileSelection();
    updateStorageMonitor();
    
    // Select the newly uploaded key in preview dropdown
    selectPreviewKey.value = key;
    renderSelectedData();
}

// Simple Robust CSV Parser
function parseCSV(text) {
    const lines = text.split(/\r?\n/).filter(line => line.trim() !== "");
    if (lines.length === 0) return [];
    
    // Header parsing
    const headers = splitCSVLine(lines[0]);
    const result = [];
    
    for (let i = 1; i < lines.length; i++) {
        const currentLine = splitCSVLine(lines[i]);
        if (currentLine.length === headers.length) {
            const obj = {};
            for (let j = 0; j < headers.length; j++) {
                obj[headers[j]] = currentLine[j];
            }
            result.push(obj);
        }
    }
    return result;
}

// Helper to handle commas inside quotes in CSV
function splitCSVLine(line) {
    const result = [];
    let insideQuote = false;
    let entry = '';
    
    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        if (char === '"') {
            insideQuote = !insideQuote;
        } else if (char === ',' && !insideQuote) {
            result.push(entry.trim().replace(/^["']|["']$/g, ''));
            entry = '';
        } else {
            entry += char;
        }
    }
    result.push(entry.trim().replace(/^["']|["']$/g, ''));
    return result;
}

// Real-time local storage monitor update
function updateStorageMonitor() {
    const keys = Object.keys(localStorage);
    localstorageKeysList.innerHTML = '';
    
    // Clear dropdown selection
    const prevSelected = selectPreviewKey.value;
    selectPreviewKey.innerHTML = '<option value="">-- Pilih Key dari LocalStorage --</option>';

    if (keys.length === 0) {
        localstorageKeysList.innerHTML = `
            <div class="empty-state">
                <i data-lucide="inbox"></i>
                <p>LocalStorage Anda kosong. Upload file untuk mengisi.</p>
            </div>
        `;
        storageProgressBar.style.width = '0%';
        storagePercentageText.textContent = '0.00%';
        storageUsedText.textContent = '0.00 KB';
        lucide.createIcons();
        return;
    }

    let totalBytes = 0;
    
    // Sort keys alphabetically
    keys.sort().forEach(key => {
        const value = localStorage.getItem(key);
        // Approximate byte size: JS strings are UTF-16, so 2 bytes per char usually.
        // In localstorage calculations, browsers usually count UTF-16 or UTF-8 characters. 
        // We will calculate exact string length as UTF-8 bytes for gauge simulation.
        const sizeBytes = encodeURIComponent(value).replace(/%[0-9A-F]{2}/g, '').length;
        totalBytes += sizeBytes;

        // Determine type of storage key for badges
        let typeBadge = 'text';
        let badgeClass = 'badge-text';
        
        if (value.startsWith('data:image/')) {
            typeBadge = 'image';
            badgeClass = 'badge-image';
        } else if (key.includes('excel') || key.includes('sheet') || key.includes('inventory')) {
            typeBadge = 'excel';
            badgeClass = 'badge-excel';
        } else if (key.includes('csv')) {
            typeBadge = 'csv';
            badgeClass = 'badge-csv';
        } else {
            try {
                JSON.parse(value);
                typeBadge = 'json';
                badgeClass = 'badge-json';
            } catch(e) {
                // fall back to default text
            }
        }

        // Add to dropdown preview
        const opt = document.createElement('option');
        opt.value = key;
        opt.textContent = `${key} (${formatBytes(sizeBytes)})`;
        selectPreviewKey.appendChild(opt);

        // Add row to monitor list
        const row = document.createElement('div');
        row.className = 'storage-item';
        row.innerHTML = `
            <div class="item-info">
                <span class="item-key" title="${key}">${key}</span>
                <div class="item-meta">
                    <span class="item-size">${formatBytes(sizeBytes)}</span>
                    <span class="item-type-badge ${badgeClass}">${typeBadge}</span>
                </div>
            </div>
            <div class="item-actions">
                <button type="button" class="btn-icon view-btn" data-key="${key}" title="Lihat Isi JSON">
                    <i data-lucide="eye"></i>
                </button>
                <button type="button" class="btn-icon delete-btn" data-key="${key}" title="Hapus Key ini">
                    <i data-lucide="trash-2"></i>
                </button>
            </div>
        `;
        localstorageKeysList.appendChild(row);
    });

    // Re-select previously selected preview key if it still exists
    if (keys.includes(prevSelected)) {
        selectPreviewKey.value = prevSelected;
    }

    // Storage capacity calculations (Browser standard: ~5MB)
    const limitBytes = 5 * 1024 * 1024; // 5MB
    const percentage = (totalBytes / limitBytes) * 100;
    
    storageProgressBar.style.width = `${Math.min(percentage, 100)}%`;
    storagePercentageText.textContent = `${percentage.toFixed(2)}%`;
    storageUsedText.textContent = formatBytes(totalBytes);
    
    // Add warning color if storage is > 80%
    if (percentage > 80) {
        storageProgressBar.style.background = 'linear-gradient(to right, var(--warning), var(--danger))';
    } else {
        storageProgressBar.style.background = 'linear-gradient(to right, var(--secondary), var(--primary), var(--accent))';
    }

    // Attach actions to dynamically created rows
    document.querySelectorAll('.delete-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const keyToDelete = e.currentTarget.getAttribute('data-key');
            if (confirm(`Hapus key "${keyToDelete}" dari LocalStorage?`)) {
                localStorage.removeItem(keyToDelete);
                updateStorageMonitor();
                if (selectPreviewKey.value === keyToDelete) {
                    resetVisualizer();
                }
            }
        });
    });

    document.querySelectorAll('.view-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const keyToView = e.currentTarget.getAttribute('data-key');
            showModalDetail(keyToView);
        });
    });

    lucide.createIcons();
}

// Data details Modal Viewer
function showModalDetail(key) {
    const value = localStorage.getItem(key);
    modalKeyName.textContent = key;
    
    // exact byte estimation
    const sizeBytes = encodeURIComponent(value).replace(/%[0-9A-F]{2}/g, '').length;
    modalKeySize.textContent = formatBytes(sizeBytes);
    
    try {
        // Pretty print JSON
        const parsed = JSON.parse(value);
        modalJsonContent.textContent = JSON.stringify(parsed, null, 4);
    } catch(e) {
        // Raw text if not JSON
        modalJsonContent.textContent = value;
    }
    
    detailModal.classList.remove('hidden');
}

// Render Board Visualizer
function renderSelectedData() {
    const key = selectPreviewKey.value;
    if (!key) {
        resetVisualizer();
        return;
    }

    const value = localStorage.getItem(key);
    if (value === null) {
        resetVisualizer();
        return;
    }

    visualizerOutput.innerHTML = '';

    // RAW JSON STRING MODE
    if (renderMode === 'raw') {
        const rawPre = document.createElement('pre');
        rawPre.className = 'raw-code-view';
        try {
            const parsed = JSON.parse(value);
            rawPre.textContent = JSON.stringify(parsed, null, 4);
        } catch (e) {
            rawPre.textContent = value; // Fallback plain text
        }
        visualizerOutput.appendChild(rawPre);
        return;
    }

    // VISUALISASI UI MODE (depends on content type)
    
    // 1. Image Base64 check
    if (value.startsWith('data:image/')) {
        const imgContainer = document.createElement('div');
        imgContainer.className = 'image-render-container';
        imgContainer.innerHTML = `
            <div class="image-preview-frame">
                <img src="${value}" alt="Rendered Base64 LocalStorage Image">
            </div>
            <div class="alert-box alert-warning">
                <strong>Catatan Base64:</strong> Gambar di atas di-render langsung dari string Base64 LocalStorage. Ukuran string adalah <code>${formatBytes(value.length)}</code>.
            </div>
        `;
        visualizerOutput.appendChild(imgContainer);
        return;
    }

    // 2. Tabular Data (JSON Array of Objects from JSON/Excel/CSV)
    try {
        const parsed = JSON.parse(value);
        
        if (Array.isArray(parsed) && parsed.length > 0) {
            // Render beautiful HTML table
            const tableWrapper = document.createElement('div');
            tableWrapper.className = 'visualizer-table-wrapper';
            
            // Extract unique keys for headers
            const headers = Object.keys(parsed[0]);
            
            let tableHTML = `<table class="visualizer-table"><thead><tr>`;
            headers.forEach(h => {
                tableHTML += `<th>${h}</th>`;
            });
            tableHTML += `</tr></thead><tbody>`;
            
            parsed.forEach(row => {
                tableHTML += `<tr>`;
                headers.forEach(h => {
                    const cellVal = row[h] !== undefined ? row[h] : '';
                    tableHTML += `<td>${cellVal}</td>`;
                });
                tableHTML += `</tr>`;
            });
            
            tableHTML += `</tbody></table>`;
            tableWrapper.innerHTML = tableHTML;
            visualizerOutput.appendChild(tableWrapper);
        } 
        else if (typeof parsed === 'object' && parsed !== null) {
            // It's a single JSON object, render as key-value card
            const grid = document.createElement('div');
            grid.className = 'card-render-grid';
            
            const card = document.createElement('div');
            card.className = 'data-card';
            card.innerHTML = `
                <div class="data-card-header">
                    <span class="data-card-title">JSON Object Data</span>
                    <span class="data-card-index">{}</span>
                </div>
                <div class="data-card-body">
                    ${Object.entries(parsed).map(([k, v]) => `
                        <div class="data-row-item">
                            <span class="data-row-key">${k}</span>
                            <span class="data-row-val" title="${typeof v === 'object' ? JSON.stringify(v) : v}">
                                ${typeof v === 'object' ? 'Object/Array' : v}
                            </span>
                        </div>
                    `).join('')}
                </div>
            `;
            grid.appendChild(card);
            visualizerOutput.appendChild(grid);
        } else {
            // Simple array or value parsed as JSON (e.g. number/boolean)
            renderPlainTextView(value);
        }
    } catch (e) {
        // 3. Fallback: Plain Text (for .txt uploaded files)
        renderPlainTextView(value);
    }
}

function renderPlainTextView(text) {
    const txtView = document.createElement('div');
    txtView.className = 'text-plain-view';
    txtView.textContent = text;
    visualizerOutput.appendChild(txtView);
}

function resetVisualizer() {
    selectPreviewKey.value = '';
    visualizerOutput.innerHTML = `
        <div class="empty-state">
            <i data-lucide="eye-off"></i>
            <p>Silakan pilih Key LocalStorage di atas untuk memvisualisasikan data.</p>
        </div>
    `;
    lucide.createIcons();
}

// Code snippets viewer helper
function updateCodeDisplay(snippetKey) {
    const display = document.getElementById('code-display');
    const title = document.getElementById('snippet-title');
    
    if (codeSnippets[snippetKey]) {
        display.textContent = codeSnippets[snippetKey];
        title.textContent = snippetKey === 'code-json' ? 'json-parser.js' : 
                          snippetKey === 'code-excel' ? 'excel-sheetjs-parser.js' : 
                          snippetKey === 'code-image' ? 'image-base64-encoder.js' : 
                          'csv-parser.js';
    }
}

// Help utility to format file size readable
function formatBytes(bytes, decimals = 2) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

// Pre-fill local storage with demo learning assets
function loadDemoData(force = false) {
    const demoJson = 'demo_users_json';
    const demoCsv = 'demo_sales_csv';
    const demoImg = 'demo_avatar_base64';
    const demoTxt = 'demo_policy_txt';
    const demoExcel = 'demo_sheetjs_parsed';

    if (!force && (localStorage.getItem(demoJson) || localStorage.getItem(demoExcel))) {
        return; // Already populated
    }

    localStorage.clear();

    // 1. JSON Demo Data
    const users = [
        { "ID": "U001", "Nama": "Ahmad Dani", "Email": "dani@web.com", "Peran": "Administrator", "Status": "Aktif" },
        { "ID": "U002", "Nama": "Siti Aminah", "Email": "siti.amin@web.com", "Peran": "Editor", "Status": "Aktif" },
        { "ID": "U003", "Nama": "Budi Santoso", "Email": "budi.s@web.com", "Peran": "Pelanggan", "Status": "Nonaktif" }
    ];
    localStorage.setItem(demoJson, JSON.stringify(users));

    // 2. CSV Demo Data (parsed to structured JSON array)
    const sales = [
        { "Bulan": "Januari", "Target": "Rp 50jt", "Realisasi": "Rp 48.5jt", "Persentase": "97%" },
        { "Bulan": "Februari", "Target": "Rp 50jt", "Realisasi": "Rp 52.1jt", "Persentase": "104%" },
        { "Bulan": "Maret", "Target": "Rp 60jt", "Realisasi": "Rp 61.2jt", "Persentase": "102%" }
    ];
    localStorage.setItem(demoCsv, JSON.stringify(sales));

    // 3. Image Demo Data (Data URL SVG representation to save space but show image rendering)
    const svgDemo = `data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="300" height="200" viewBox="0 0 300 200"><rect width="100%" height="100%" fill="%231e1b4b"/><defs><linearGradient id="g" x1="0%" y1="0%" x2="100%" y2="100%"><stop offset="0%" stop-color="%238b5cf6"/><stop offset="100%" stop-color="%23d946ef"/></linearGradient></defs><circle cx="150" cy="85" r="45" fill="url(%23g)"/><text x="150" y="160" fill="%23e5e7eb" font-family="sans-serif" font-weight="bold" font-size="16" text-anchor="middle">StorageLearn Demo</text><text x="150" y="180" fill="%239ca3af" font-family="sans-serif" font-size="11" text-anchor="middle">Image Rendered from LocalStorage</text></svg>`;
    localStorage.setItem(demoImg, svgDemo);

    // 4. Text Demo Data
    const notes = `DOKUMEN INTEGRITAS BELAJAR LOCALSTORAGE:
1. LocalStorage bersifat synchronous dan memblokir thread utama.
2. Hindari menyimpan objek biner berukuran > 1MB.
3. Selalu pergunakan try-catch saat memanggil JSON.parse() dari LocalStorage.
4. Gunakan keys yang unik agar tidak menimpa data sistem lainnya.`;
    localStorage.setItem(demoTxt, notes);

    // 5. Excel parsed Demo Data
    const products = [
        { "Kode": "P-001", "Nama Barang": "Kopi Arabika Toraja 250g", "Kategori": "Minuman", "Stok": "45", "Harga": "Rp 65,000" },
        { "Kode": "P-002", "Nama Barang": "Teh Hijau Organic 100g", "Kategori": "Minuman", "Stok": "120", "Harga": "Rp 32,500" },
        { "Kode": "P-003", "Nama Barang": "Madu Hutan Murni 500ml", "Kategori": "Kesehatan", "Stok": "18", "Harga": "Rp 120,000" }
    ];
    localStorage.setItem(demoExcel, JSON.stringify(products));
}
