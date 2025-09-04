// Global variables
let tw2Data = null;
let excelData = null;

// Initialize on page load
$(document).ready(function() {
    setupDropzones();
    setupFileInputs();
});

// Setup drag and drop zones
function setupDropzones() {
    // TW2 Dropzone
    const tw2Dropzone = document.getElementById('tw2-dropzone');
    const tw2FileInput = document.getElementById('tw2-file-input');
    
    tw2Dropzone.addEventListener('click', () => tw2FileInput.click());
    
    tw2Dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        tw2Dropzone.classList.add('dragover');
    });
    
    tw2Dropzone.addEventListener('dragleave', () => {
        tw2Dropzone.classList.remove('dragover');
    });
    
    tw2Dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        tw2Dropzone.classList.remove('dragover');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleTW2File(files[0]);
        }
    });
    
    // Excel Dropzone
    const excelDropzone = document.getElementById('excel-dropzone');
    const excelFileInput = document.getElementById('excel-file-input');
    
    excelDropzone.addEventListener('click', () => excelFileInput.click());
    
    excelDropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        excelDropzone.classList.add('dragover');
    });
    
    excelDropzone.addEventListener('dragleave', () => {
        excelDropzone.classList.remove('dragover');
    });
    
    excelDropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        excelDropzone.classList.remove('dragover');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleExcelFile(files[0]);
        }
    });
}

// Setup file input handlers
function setupFileInputs() {
    document.getElementById('tw2-file-input').addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleTW2File(e.target.files[0]);
        }
    });
    
    document.getElementById('excel-file-input').addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleExcelFile(e.target.files[0]);
        }
    });
}

// Handle TW2 file upload
function handleTW2File(file) {
    if (!file.name.endsWith('.tw2') && !file.name.endsWith('.mdb')) {
        showToast('Please select a valid .tw2 or .mdb file', 'error');
        return;
    }
    
    const formData = new FormData();
    formData.append('file', file);
    
    showToast('Uploading TW2 file...', 'info');
    
    // Use XMLHttpRequest for better control
    const xhr = new XMLHttpRequest();
    
    xhr.open('POST', '/upload_tw2', true);
    
    xhr.onload = function() {
        if (xhr.status === 200) {
            try {
                const data = JSON.parse(xhr.responseText);
                if (data.success) {
                    tw2Data = data;
                    document.getElementById('tw2-filename').textContent = file.name;
                    document.getElementById('tw2-records').textContent = data.row_count || '0';
                    document.getElementById('tw2-file-info').style.display = 'block';
                    document.getElementById('tw2-dropzone').classList.add('has-file');
                    showToast('TW2 file loaded successfully', 'success');
                } else {
                    showToast('Error: ' + (data.error || 'Unknown error'), 'error');
                }
            } catch (e) {
                console.error('Parse error:', e);
                showToast('Error parsing response', 'error');
            }
        } else {
            try {
                const errorData = JSON.parse(xhr.responseText);
                showToast('Error: ' + (errorData.error || 'Upload failed'), 'error');
            } catch (e) {
                showToast('Error uploading file', 'error');
            }
        }
    };
    
    xhr.onerror = function() {
        showToast('Network error during upload', 'error');
    };
    
    xhr.send(formData);
}

// Handle Excel file upload
function handleExcelFile(file) {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        showToast('Please select a valid Excel file', 'error');
        return;
    }
    
    const formData = new FormData();
    formData.append('file', file);
    
    showToast('Uploading Excel file...', 'info');
    
    const xhr = new XMLHttpRequest();
    
    xhr.open('POST', '/upload_excel', true);
    
    xhr.onload = function() {
        if (xhr.status === 200) {
            try {
                const data = JSON.parse(xhr.responseText);
                if (data.success) {
                    excelData = data;
                    document.getElementById('excel-filename').textContent = file.name;
                    document.getElementById('excel-records').textContent = data.row_count || '0';
                    document.getElementById('excel-file-info').style.display = 'block';
                    document.getElementById('excel-dropzone').classList.add('has-file');
                    showToast('Excel file loaded successfully', 'success');
                } else {
                    showToast('Error: ' + (data.error || 'Unknown error'), 'error');
                }
            } catch (e) {
                console.error('Parse error:', e);
                showToast('Error parsing response', 'error');
            }
        } else {
            showToast('Error uploading file', 'error');
        }
    };
    
    xhr.onerror = function() {
        showToast('Network error during upload', 'error');
    };
    
    xhr.send(formData);
}

// Show toast message
function showToast(message, type = 'info') {
    const container = document.getElementById('status-messages');
    
    const toast = document.createElement('div');
    toast.className = `toast ${type} show`;
    toast.setAttribute('role', 'alert');
    toast.innerHTML = `
        <div class="toast-header">
            <strong class="me-auto">${type.charAt(0).toUpperCase() + type.slice(1)}</strong>
            <button type="button" class="btn-close" data-bs-dismiss="toast"></button>
        </div>
        <div class="toast-body">
            ${message}
        </div>
    `;
    
    container.appendChild(toast);
    
    // Auto remove after 5 seconds
    setTimeout(() => {
        toast.remove();
    }, 5000);
}

// Stub functions for now
function viewTW2Data() {
    showToast('TW2 data preview coming soon', 'info');
}

function viewExcelData() {
    showToast('Excel data preview coming soon', 'info');
}

function previewMapping() {
    showToast('Mapping preview coming soon', 'info');
}

function applyMapping() {
    showToast('Apply mapping coming soon', 'info');
}

function resetMapping() {
    showToast('Reset mapping coming soon', 'info');
}