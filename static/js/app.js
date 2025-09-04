// Global variables
let tw2Data = null;
let excelData = null;
let mappingFields = null;
let currentMappings = {};

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
    
    fetch('/upload_tw2', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            tw2Data = data;
            document.getElementById('tw2-filename').textContent = file.name;
            document.getElementById('tw2-records').textContent = data.row_count;
            document.getElementById('tw2-file-info').style.display = 'block';
            document.getElementById('tw2-dropzone').classList.add('has-file');
            showToast('TW2 file loaded successfully', 'success');
            checkBothFilesLoaded();
        } else {
            showToast('Error loading TW2 file: ' + data.error, 'error');
        }
    })
    .catch(error => {
        showToast('Error uploading file: ' + error, 'error');
    });
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
    
    fetch('/upload_excel', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            excelData = data;
            document.getElementById('excel-filename').textContent = file.name;
            document.getElementById('excel-records').textContent = data.row_count;
            document.getElementById('excel-file-info').style.display = 'block';
            document.getElementById('excel-dropzone').classList.add('has-file');
            showToast('Excel file loaded successfully', 'success');
            checkBothFilesLoaded();
        } else {
            showToast('Error loading Excel file: ' + data.error, 'error');
        }
    })
    .catch(error => {
        showToast('Error uploading file: ' + error, 'error');
    });
}

// Check if both files are loaded
function checkBothFilesLoaded() {
    if (tw2Data && excelData) {
        loadMappingFields();
        document.getElementById('mapping-section').style.display = 'block';
    }
}

// Load mapping fields
function loadMappingFields() {
    fetch('/get_mapping_fields')
    .then(response => response.json())
    .then(data => {
        mappingFields = data;
        buildMappingTable();
    })
    .catch(error => {
        showToast('Error loading mapping fields: ' + error, 'error');
    });
}

// Build the mapping table
function buildMappingTable() {
    const tbody = document.getElementById('mapping-table');
    tbody.innerHTML = '';
    
    mappingFields.target_fields.forEach(field => {
        const row = document.createElement('tr');
        
        // Target field column
        const targetCell = document.createElement('td');
        targetCell.innerHTML = `<strong>${field}</strong>`;
        row.appendChild(targetCell);
        
        // Source field dropdown
        const sourceCell = document.createElement('td');
        const select = document.createElement('select');
        select.className = 'form-select form-select-sm';
        select.id = `mapping-${field}`;
        
        // Add empty option
        const emptyOption = document.createElement('option');
        emptyOption.value = '';
        emptyOption.textContent = '-- Select Excel Column --';
        select.appendChild(emptyOption);
        
        // Add Excel columns as options
        mappingFields.excel_fields.forEach(excelField => {
            const option = document.createElement('option');
            option.value = excelField;
            option.textContent = excelField;
            
            // Auto-select suggested mapping
            if (mappingFields.suggested_mappings[field] === excelField) {
                option.selected = true;
                currentMappings[field] = excelField;
            }
            
            select.appendChild(option);
        });
        
        select.addEventListener('change', (e) => {
            if (e.target.value) {
                currentMappings[field] = e.target.value;
            } else {
                delete currentMappings[field];
            }
        });
        
        sourceCell.appendChild(select);
        row.appendChild(sourceCell);
        
        // Actions column
        const actionsCell = document.createElement('td');
        if (field === 'Tag') {
            actionsCell.innerHTML = '<span class="badge bg-danger">Required</span>';
        } else {
            actionsCell.innerHTML = '<span class="badge bg-secondary">Optional</span>';
        }
        row.appendChild(actionsCell);
        
        tbody.appendChild(row);
    });
}

// View TW2 Data
function viewTW2Data() {
    if (!tw2Data) return;
    
    document.getElementById('data-preview-section').style.display = 'block';
    
    // Switch to TW2 tab
    document.getElementById('tw2-tab').click();
    
    // Build table
    const table = document.getElementById('tw2-table');
    
    // Create header
    let headerHtml = '<thead><tr>';
    const columns = tw2Data.columns.slice(0, 10); // Show first 10 columns
    columns.forEach(col => {
        headerHtml += `<th>${col.name}</th>`;
    });
    headerHtml += '</tr></thead>';
    
    // Create body
    let bodyHtml = '<tbody>';
    tw2Data.data.slice(0, 50).forEach(row => { // Show first 50 rows
        bodyHtml += '<tr>';
        columns.forEach(col => {
            let value = row[col.name];
            if (value === null || value === undefined) value = '';
            bodyHtml += `<td>${value}</td>`;
        });
        bodyHtml += '</tr>';
    });
    bodyHtml += '</tbody>';
    
    table.innerHTML = headerHtml + bodyHtml;
}

// View Excel Data
function viewExcelData() {
    if (!excelData) return;
    
    document.getElementById('data-preview-section').style.display = 'block';
    
    // Switch to Excel tab
    document.getElementById('excel-tab').click();
    
    // Build table
    const table = document.getElementById('excel-table');
    
    // Create header
    let headerHtml = '<thead><tr>';
    excelData.columns.forEach(col => {
        headerHtml += `<th>${col}</th>`;
    });
    headerHtml += '</tr></thead>';
    
    // Create body
    let bodyHtml = '<tbody>';
    excelData.data.forEach(row => {
        bodyHtml += '<tr>';
        excelData.columns.forEach(col => {
            let value = row[col];
            if (value === null || value === undefined) value = '';
            bodyHtml += `<td>${value}</td>`;
        });
        bodyHtml += '</tr>';
    });
    bodyHtml += '</tbody>';
    
    table.innerHTML = headerHtml + bodyHtml;
}

// Preview mapping
function previewMapping() {
    if (!currentMappings.Tag) {
        showToast('Please map the Tag field first (required for matching records)', 'error');
        return;
    }
    
    showToast('Generating preview...', 'info');
    
    fetch('/preview_mapping', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ mappings: currentMappings })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            displayPreview(data.preview);
            showToast('Preview generated successfully', 'success');
        } else {
            showToast('Error generating preview: ' + data.error, 'error');
        }
    })
    .catch(error => {
        showToast('Error: ' + error, 'error');
    });
}

// Display preview
function displayPreview(previewData) {
    document.getElementById('data-preview-section').style.display = 'block';
    document.getElementById('preview-tab').click();
    
    const container = document.getElementById('preview-content');
    container.innerHTML = '';
    
    if (previewData.length === 0) {
        container.innerHTML = '<div class="alert alert-warning">No matching records found</div>';
        return;
    }
    
    previewData.forEach(item => {
        const div = document.createElement('div');
        div.className = 'preview-item';
        
        let html = `<h6>Tag: ${item.tag}</h6>`;
        
        for (const [field, change] of Object.entries(item.changes)) {
            html += `
                <div class="change-item">
                    <span class="field-name">${field}:</span>
                    <span class="old-value">${change.old || 'empty'}</span>
                    <span>→</span>
                    <span class="new-value">${change.new || 'empty'}</span>
                </div>
            `;
        }
        
        div.innerHTML = html;
        container.appendChild(div);
    });
}

// Apply mapping
function applyMapping() {
    if (!currentMappings.Tag) {
        showToast('Please map the Tag field first (required for matching records)', 'error');
        return;
    }
    
    if (!confirm('Are you sure you want to apply these mappings? A backup will be created.')) {
        return;
    }
    
    showToast('Applying mappings...', 'info');
    
    fetch('/apply_mapping', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ mappings: currentMappings })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            let message = `Successfully updated ${data.updated_records} records. `;
            message += `Backup saved as: ${data.backup_file}`;
            showToast(message, 'success');
            
            if (data.errors && data.errors.length > 0) {
                console.error('Errors during update:', data.errors);
                showToast(`Warning: ${data.errors.length} errors occurred. Check console for details.`, 'error');
            }
        } else {
            showToast('Error applying mappings: ' + data.error, 'error');
        }
    })
    .catch(error => {
        showToast('Error: ' + error, 'error');
    });
}

// Reset mapping
function resetMapping() {
    currentMappings = {};
    buildMappingTable();
    showToast('Mappings reset', 'info');
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