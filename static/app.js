// Global variables
        let tw2Data = null;
        let excelData = null;
        let updatedTw2Data = null;
        let mappingFields = null;
        let currentMappings = {};

        // TW2 field descriptions for tooltips
        const fieldDescriptions = {
            'Tag': 'Unit identifier - Must match Excel Unit_No (e.g., V-1-1 → V-1-01)',
            'UnitSize': 'VAV unit size designation (e.g., 6, 8, 10, 14, 24x16)',
            'InletSize': 'Air inlet size in inches (e.g., 6", 8", 10", 14")',
            'CFMDesign': 'Design air flow rate in CFM - Maximum airflow capacity',
            'CFMMinPrime': 'Minimum primary airflow in CFM when heating/cooling',
            'HeatingPrime': 'Primary airflow during heating mode in CFM',
            'HWCFM': 'Hot water coil airflow in CFM - Airflow through heating coil',
            'HWGPM': 'Hot water flow rate in GPM (gallons per minute)'
        };

        // Initialize on page load
        $(document).ready(function() {
            setupDropzones();
            setupFileInputs();
        });

        // Setup drag and drop zones
        function setupDropzones() {
            // TW2 Dropzone
            setupDropzone('tw2-dropzone', 'tw2-file-input', handleTW2File);
            // Excel Dropzone
            setupDropzone('excel-dropzone', 'excel-file-input', handleExcelFile);
            // Updated TW2 Dropzone
            setupDropzone('updated-tw2-dropzone', 'updated-tw2-file-input', handleUpdatedTW2File);
        }

        function setupDropzone(dropzoneId, inputId, handler) {
            const dropzone = document.getElementById(dropzoneId);
            const fileInput = document.getElementById(inputId);
            
            dropzone.addEventListener('click', () => fileInput.click());
            
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
                
                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    handler(files[0]);
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
            
            document.getElementById('updated-tw2-file-input').addEventListener('change', (e) => {
                if (e.target.files.length > 0) {
                    handleUpdatedTW2File(e.target.files[0]);
                }
            });
        }

        // Format numbers to two decimals when possible
        function formatNumber(value) {
            if (value === null || value === undefined) return "";
            if (typeof value === "number" && isFinite(value)) return value.toFixed(2);
            if (typeof value === "string") {
                const num = Number(value);
                if (!Number.isNaN(num) && isFinite(num)) return num.toFixed(2);
            }
            return value;
        }

        // Only round MBH/LAT/WPD/APD columns; show others as-is
        const ROUND_KEYS = ['mbh', 'lat', 'wpd', 'apd'];

        function normalizeColumnName(name) {
            return String(name || "").replace(/[\s]/g, " ").replace(/[ _]+/g, "_").replace(/_/g, "").toLowerCase();
        }

        function formatByColumn(columnName, value) {
            const norm = normalizeColumnName(columnName);
            const shouldRound = ROUND_KEYS.some(k => norm.includes(k));
            if (shouldRound) {
                return formatNumber(value);
            }
            return value === null || value === undefined ? '' : value;
        }



        // Navigate to the Performance Comparison tab/view
        function showComparisonView() {
            try {
                // Ensure the preview section is visible
                const preview = document.getElementById('data-preview-section');
                if (preview) preview.style.display = 'block';

                const tab = document.getElementById('comparison-tab');
                if (tab) tab.click();

                const panel = document.getElementById('comparison-data');
                if (panel) panel.scrollIntoView({ behavior: 'smooth', block: 'start' });

                // If both required datasets are loaded, run the comparison immediately
                const hasExcel = !!excelData || (window.sessionStorage && sessionStorage.getItem('excel_loaded') === '1');
                const hasUpdated = !!updatedTw2Data || (window.sessionStorage && sessionStorage.getItem('updated_tw2_loaded') === '1');

                if (hasExcel && hasUpdated) {
                    runPerformanceComparison();
                } else {
                    // Guide the user to load missing files
                    if (!hasExcel && !hasUpdated) {
                        showToast('Please upload Excel and Updated TW2 files first', 'error');
                    } else if (!hasExcel) {
                        showToast('Please upload the Excel file first', 'error');
                        const excelTab = document.getElementById('excel-tab');
                        if (excelTab) excelTab.click();
                    } else {
                        showToast('Please upload the Updated TW2 file first', 'error');
                        const tw2Tab = document.getElementById('tw2-tab');
                        if (tw2Tab) tw2Tab.click();
                    }
                }
            } catch (e) {
                console.warn('NAV: Unable to switch to comparison view', e);
            }
        }

        // Sanitize local Windows path: remove surrounding quotes (Copy as path)
        function sanitizeLocalPath(path) {
            if (!path) return path;
            path = path.trim();
            if ((path.startsWith('"') && path.endsWith('"')) || (path.startsWith("'") && path.endsWith("'"))) {
                return path.slice(1, -1).trim();
            }
            return path;
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
            
            uploadFile('/upload_tw2', formData, (data) => {
                tw2Data = data;
                document.getElementById('tw2-filename').textContent = file.name;
                document.getElementById('tw2-records').textContent = data.row_count;
                
                // Show some column names
                const columnsDiv = document.getElementById('tw2-columns');
                const keyColumns = ['Tag', 'UnitSize', 'InletSize', 'CFMDesign', 'HWCFM', 'HWGPM'];
                const availableKeyColumns = keyColumns.filter(col => data.columns.includes(col));
                columnsDiv.innerHTML = `<small><strong>Key columns:</strong> ${availableKeyColumns.join(', ')}</small>`;
                
                document.getElementById('tw2-file-info').style.display = 'block';
                document.getElementById('tw2-dropzone').classList.add('has-file');
                showToast('TW2 file loaded successfully', 'success');
                checkBothFilesLoaded();
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
            
            // Add configuration parameters
            const dataStartRow = document.getElementById('data-start-row').value;
            const headerRows = document.getElementById('header-rows').value;
            const skipTitleRow = document.getElementById('skip-title-row').checked;
            
            formData.append('data_start_row', dataStartRow);
            formData.append('header_rows', headerRows);
            formData.append('skip_title_row', skipTitleRow);
            
            showToast('Uploading Excel file...', 'info');
            
            uploadFile('/upload_excel', formData, (data) => {
                excelData = data;
                try { if (window.sessionStorage) sessionStorage.setItem('excel_loaded','1'); } catch (e) {}
                document.getElementById('excel-filename').textContent = file.name;
                document.getElementById('excel-records').textContent = data.row_count;
                
                // Show column names and header mapping info if available
                const columnsDiv = document.getElementById('excel-columns');
                let columnInfo = `<small><strong>Columns (${data.columns.length}):</strong> ${data.columns.join(', ')}</small>`;
                
                // Add header detection info if available
                if (data.header_info) {
                    columnInfo += '<br><details class="mt-2">';
                    columnInfo += '<summary><small><strong>Header Detection Details</strong> (click to expand)</small></summary>';
                    columnInfo += '<div class="mt-1 p-2 bg-light border rounded">';
                    columnInfo += `<small><strong>Original Headers:</strong> ${data.header_info.original_headers ? data.header_info.original_headers.join(', ') : 'N/A'}</small><br>`;
                    columnInfo += `<small><strong>Combined Headers:</strong> ${data.header_info.combined_headers ? data.header_info.combined_headers.join(', ') : 'N/A'}</small><br>`;
                    columnInfo += `<small><strong>Mapped Headers:</strong> ${data.columns.join(', ')}</small>`;
                    columnInfo += '</div></details>';
                }
                
                columnsDiv.innerHTML = columnInfo;
                
                document.getElementById('excel-file-info').style.display = 'block';
                document.getElementById('excel-dropzone').classList.add('has-file');
                showToast('Excel file loaded successfully', 'success');
                checkBothFilesLoaded();
            });
        }

        // Generic file upload function
        function uploadFile(endpoint, formData, successCallback) {
            fetch(endpoint, {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    successCallback(data);
                } else {
                    showToast('Error: ' + (data.error || 'Upload failed'), 'error');
                }
            })
            .catch(error => {
                console.error('Upload error:', error);
                showToast('Network error during upload', 'error');
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
                
                // Target field column with tooltip
                const targetCell = document.createElement('td');
                const description = fieldDescriptions[field] || 'No description available';
                targetCell.innerHTML = `<strong data-bs-toggle="tooltip" data-bs-placement="right" title="${description}" style="cursor: help; border-bottom: 1px dotted #999;">${field}</strong>`;
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
                    updateApplyButton();
                });
                
                sourceCell.appendChild(select);
                row.appendChild(sourceCell);
                
                // Status column
                const statusCell = document.createElement('td');
                if (field === 'Tag') {
                    statusCell.innerHTML = '<span class="badge bg-danger">Required</span>';
                } else {
                    statusCell.innerHTML = '<span class="badge bg-secondary">Optional</span>';
                }
                row.appendChild(statusCell);
                
                tbody.appendChild(row);
            });
            
            // Initialize Bootstrap tooltips
            var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
            var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {
                return new bootstrap.Tooltip(tooltipTriggerEl);
            });
            
            updateApplyButton();
        }

        function updateApplyButton() {
            const applyBtn = document.getElementById('apply-mapping-btn');
            // Enable button if Tag is mapped
            if (currentMappings.Tag) {
                applyBtn.disabled = false;
            } else {
                applyBtn.disabled = true;
            }
        }

        // View TW2 Data
        function viewTW2Data() {
            if (!tw2Data) return;
            
            document.getElementById('data-preview-section').style.display = 'block';
            document.getElementById('tw2-tab').click();
            
            const table = document.getElementById('tw2-table');
            
            // Show ALL available columns from TW2 data
            const availableColumns = tw2Data.columns;
            
            // Create header - show ALL columns
            let headerHtml = '<thead><tr>';
            availableColumns.forEach(col => {
                headerHtml += `<th>${col}</th>`;
            });
            headerHtml += '</tr></thead>';
            
            // Create body - show ALL records with scrolling
            let bodyHtml = '<tbody>';
            tw2Data.data.forEach((row, index) => {
                bodyHtml += '<tr>';
                availableColumns.forEach(col => {
                    let value = row[col];
                    if (value === null || value === undefined) value = '';
                    bodyHtml += `<td>${formatByColumn(col, value)}</td>`;
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
            document.getElementById('excel-tab').click();
            
            const table = document.getElementById('excel-table');
            
            // Create header - show ALL columns
            let headerHtml = '<thead><tr>';
            excelData.columns.forEach(col => {
                headerHtml += `<th>${col}</th>`;
            });
            headerHtml += '</tr></thead>';
            
            // Create body - show ALL records
            let bodyHtml = '<tbody>';
            excelData.data.forEach((row, index) => {
                bodyHtml += `<tr>`;
                excelData.columns.forEach(col => {  // Show ALL columns including Unit_No
                    let value = row[col];
                    if (value === null || value === undefined) value = '';
                    bodyHtml += `<td>${formatByColumn(col, value)}</td>`;
                });
                bodyHtml += '</tr>';
            });
            bodyHtml += '</tbody>';
            
            table.innerHTML = headerHtml + bodyHtml;
        }

        // Apply mapping
        function applyMapping() {
            if (!currentMappings.Tag) {
                showToast('Please map the Tag field first (required for matching records)', 'error');
                return;
            }
            
            if (!confirm('Are you sure you want to apply these mappings? This will update your TW2 database. A backup will be created automatically.')) {
                return;
            }
            
            showToast('Applying mappings and updating database...', 'info');
            
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
                    let message = `Successfully updated ${data.updated_records} records.`;
                    if (data.backup_file) {
                        message += ` Backup saved as: ${data.backup_file.split('\\').pop()}`;
                    }
                    
                    document.getElementById('results-content').innerHTML = `
                        <h5>Update Complete!</h5>
                        <p>${message}</p>
                        ${data.errors ? `<p><strong>Warnings:</strong> ${data.errors.length} issues occurred. Check console for details.</p>` : ''}
                    `;
                    document.getElementById('results-section').style.display = 'block';
                    
                    showToast('Database updated successfully!', 'success');
                    downloadMergedFile();
                    
                    if (data.errors && data.errors.length > 0) {
                        console.error('Update errors:', data.errors);
                    }
                } else {
                    showToast('Error applying mappings: ' + data.error, 'error');
                }
            })
            .catch(error => {
                showToast('Error: ' + error, 'error');
            });
        }

        function downloadMergedFile() {
            const link = document.createElement('a');
            link.href = '/download_merged_tw2';
            link.download = '';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            showToast('Merged TW2 file download started', 'info');
        }

        // Reset mapping
        function resetMapping() {
            currentMappings = {};
            if (mappingFields) {
                buildMappingTable();
            }
            document.getElementById('results-section').style.display = 'none';
            showToast('Mappings reset', 'info');
        }

        // Show toast message
        function showToast(message, type = 'info') {
            const container = document.getElementById('status-messages');
            
            const toastId = 'toast_' + Date.now();
            const toast = document.createElement('div');
            toast.id = toastId;
            toast.className = `toast ${type} show`;
            toast.setAttribute('role', 'alert');
            toast.innerHTML = `
                <div class="toast-header">
                    <strong class="me-auto">${type.charAt(0).toUpperCase() + type.slice(1)}</strong>
                    <button type="button" class="btn-close" onclick="document.getElementById('${toastId}').remove()"></button>
                </div>
                <div class="toast-body">
                    ${message}
                </div>
            `;
            
            container.appendChild(toast);
            
            // Auto remove after 2.5 seconds for success/info, 4 seconds for errors
            const timeout = type === 'error' ? 4000 : 2500;
            setTimeout(() => {
                if (document.getElementById(toastId)) {
                    document.getElementById(toastId).remove();
                }
            }, timeout);
        }

        // Handle Updated TW2 file upload
        function handleUpdatedTW2File(file) {
            if (!file.name.endsWith('.tw2') && !file.name.endsWith('.mdb')) {
                showToast('Please select a TW2 or MDB file', 'error');
                return;
            }

            const formData = new FormData();
            formData.append('file', file);
            
            // Add original path if provided
            const originalPath = sanitizeLocalPath(document.getElementById('original-tw2-path').value);
            if (originalPath) {
                formData.append('original_path', originalPath);
                console.log('TW2 UPLOAD: Including original path:', originalPath);
            }

            showToast('Uploading updated TW2 file...', 'info');

            fetch('/upload_updated_tw2', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    updatedTw2Data = data;
                    try { if (window.sessionStorage) sessionStorage.setItem('updated_tw2_loaded','1'); } catch (e) {}
                    showToast(`Updated TW2 file loaded: ${data.records} records`, 'success');
                    
                    // Update UI
                    document.getElementById('updated-tw2-dropzone').classList.add('has-file');
                    document.getElementById('updated-tw2-filename').textContent = data.filename;
                    document.getElementById('updated-tw2-records').textContent = data.records;
                    document.getElementById('updated-tw2-columns').innerHTML = `<small class="text-muted">${data.column_count} columns available</small>`;
                    document.getElementById('updated-tw2-file-info').style.display = 'block';
                    
                    // Enable comparison button if all data is loaded
                    updateComparisonButtonState();
                } else {
                    showToast(`Error: ${data.error}`, 'error');
                }
            })
            .catch(error => {
                console.error('Upload error:', error);
                showToast('Upload failed', 'error');
            });
        }

        // View Updated TW2 Data
        function viewUpdatedTW2Data() {
            if (!updatedTw2Data) {
                showToast('No updated TW2 data loaded', 'error');
                return;
            }
            
            // Switch to TW2 data tab and load the updated data
            const tw2Tab = document.getElementById('tw2-tab');
            const tw2TabPane = document.getElementById('tw2-data');
            
            // Activate the tab
            document.querySelectorAll('.nav-link').forEach(link => link.classList.remove('active'));
            document.querySelectorAll('.tab-pane').forEach(pane => {
                pane.classList.remove('show', 'active');
            });
            
            tw2Tab.classList.add('active');
            tw2TabPane.classList.add('show', 'active');
            
            // Load updated TW2 data into the TW2 table
            displayUpdatedTW2Data();
            
            // Show the preview section
            document.getElementById('data-preview-section').style.display = 'block';
            
            showToast('Showing updated TW2 data in TW2 tab', 'success');
        }

        // Display Updated TW2 Data in the TW2 table
        function displayUpdatedTW2Data() {
            fetch('/get_updated_tw2_data')
            .then(response => {
                if (!response.ok) {
                    if (response.status === 400) {
                        // Expected when no TW2 data is uploaded yet
                        return { error: 'No updated TW2 data uploaded yet', expected: true };
                    }
                    throw new Error(`HTTP ${response.status}`);
                }
                return response.json();
            })
            .then(data => {
                if (data.success) {
                    const table = document.getElementById('tw2-table');
                    
                    // Show ALL available columns from updated TW2 data
                    const availableColumns = data.columns;
                    
                    // Create header with indication this is updated data
                    let headerHtml = '<thead><tr><th colspan="' + availableColumns.length + '" class="bg-warning text-dark text-center">Updated TW2 Data (Post-Titus Teams)</th></tr><tr>';
                    availableColumns.forEach(col => {
                        headerHtml += `<th>${col}</th>`;
                    });
                    headerHtml += '</tr></thead>';
                    
                    // Create body - show ALL records with scrolling
                    let bodyHtml = '<tbody>';
                    data.data.forEach((row, index) => {
                        bodyHtml += '<tr>';
                        availableColumns.forEach(col => {
                            let value = row[col];
                            if (value === null || value === undefined) value = '';
                            bodyHtml += `<td>${formatByColumn(col, value)}</td>`;
                        });
                        bodyHtml += '</tr>';
                    });
                    bodyHtml += '</tbody>';
                    
                    table.innerHTML = headerHtml + bodyHtml;
                } else if (data.error && data.expected) {
                    // Expected error - no TW2 data uploaded yet
                    const tableContainer = document.getElementById('tw2-table');
                    tableContainer.innerHTML = `
                        <thead>
                            <tr><th colspan="100%" class="bg-warning text-dark">Updated TW2 Data Preview</th></tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td colspan="100%" class="text-center p-4">
                                    <div class="alert alert-secondary">
                                        <h6><i class="bi bi-upload"></i> Upload Updated TW2 File</h6>
                                        <p>Upload your updated TW2 file from Titus Teams to view the data here.</p>
                                        <small class="text-muted">This section will show your updated TW2 data after processing in Titus Teams.</small>
                                    </div>
                                </td>
                            </tr>
                        </tbody>
                    `;
                } else {
                    // Unexpected error or fallback
                    const tableContainer = document.getElementById('tw2-table');
                    tableContainer.innerHTML = `
                        <thead>
                            <tr><th colspan="100%" class="bg-warning text-dark">Updated TW2 Data Preview</th></tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td colspan="100%" class="text-center p-4">
                                    <div class="alert alert-warning">
                                        <h6>Unable to Load TW2 Data</h6>
                                        <p>Error: ${data.error || 'Unknown error'}</p>
                                        <small class="text-muted">Please try uploading your TW2 file again.</small>
                                    </div>
                                </td>
                            </tr>
                        </tbody>
                    `;
                }
            })
            .catch(error => {
                console.log('TW2 data display error (expected if no file uploaded):', error);
                // Gracefully handle the case when no updated TW2 data is available
                const tableContainer = document.getElementById('tw2-table');
                tableContainer.innerHTML = `
                    <thead>
                        <tr><th colspan="100%" class="bg-warning text-dark">Updated TW2 Data Preview</th></tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td colspan="100%" class="text-center p-4">
                                <div class="alert alert-secondary">
                                    <h6><i class="bi bi-upload"></i> Upload Updated TW2 File</h6>
                                    <p>Upload your updated TW2 file from Titus Teams to view the data here.</p>
                                    <small class="text-muted">This section will show your updated TW2 data after processing in Titus Teams.</small>
                                </div>
                            </td>
                        </tr>
                    </tbody>
                `;
            });
        }

        // Update comparison button state
        function updateComparisonButtonState() {
            const compareBtn = document.getElementById('compare-btn');
            if (excelData && updatedTw2Data) {
                compareBtn.disabled = false;
            } else {
                compareBtn.disabled = true;
            }
        }

        // Run Performance Comparison
        function runPerformanceComparison() {
            if (!excelData || !updatedTw2Data) {
                showToast('Please upload both Excel and updated TW2 files first', 'error');
                return;
            }

            const mbhLatLowerMargin = parseFloat(document.getElementById('mbh-lat-lower-margin').value) || 15;
            const mbhLatUpperMargin = parseFloat(document.getElementById('mbh-lat-upper-margin').value) || 25;
            const wpdThreshold = parseFloat(document.getElementById('wpd-threshold').value) || 5;
            const apdThreshold = parseFloat(document.getElementById('apd-threshold').value) || 0.25;

            showToast('Running performance comparison...', 'info');

            fetch('/compare_performance', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    mbh_lat_lower_margin: mbhLatLowerMargin,
                    mbh_lat_upper_margin: mbhLatUpperMargin,
                    wpd_threshold: wpdThreshold,
                    apd_threshold: apdThreshold
                })
            })
            .then(response => {
                if (!response.ok) {
                    return response.json().then(errorData => {
                        throw new Error(errorData.error || `HTTP ${response.status}`);
                    });
                }
                return response.json();
            })
            .then(data => {
                if (data.success) {
                    const payload = data.data || {};
                    displayComparisonResults(payload.results, payload.summary);

                    const summary = payload.summary || {};
                    const totalUnits = typeof summary.total !== 'undefined' ? summary.total : null;
                    const baseMessage = totalUnits !== null
                        ? `Comparison completed: ${totalUnits} units analyzed`
                        : 'Comparison completed';

                    let sourceLabel = '';
                    if (payload.tw2_source === 'original') {
                        sourceLabel = 'Original TW2 file';
                    } else if (payload.tw2_source === 'local') {
                        sourceLabel = 'Local TW2 copy';
                    } else if (payload.tw2_source) {
                        sourceLabel = payload.tw2_source;
                    }

                    const pathDetails = payload.tw2_path
                        ? `${sourceLabel ? sourceLabel + ': ' : ''}${payload.tw2_path}`
                        : '';

                    const finalMessage = pathDetails ? `${baseMessage} (${pathDetails})` : baseMessage;
                    showToast(finalMessage, 'success');
                } else {
                    showToast(`Comparison failed: ${data.error}`, 'error');
                }
            })
            .catch(error => {
                console.error('Comparison error:', error);
                if (error.message.includes('Excel data not loaded') || error.message.includes('Updated TW2 data not loaded')) {
                    showToast('Please upload both Excel and updated TW2 files first', 'error');
                } else {
                    showToast(`Comparison failed: ${error.message}`, 'error');
                }
            });
        }

        // Setup HW Rows editing functionality
        function setupHWRowsEditing() {
            const hwRowsSelects = document.querySelectorAll('.hw-rows-select');

            hwRowsSelects.forEach(select => {
                select.addEventListener('change', function() {
                    const originalValue = this.getAttribute('data-original');
                    const currentValue = this.value;

                    if (currentValue !== originalValue) {
                        this.style.backgroundColor = '#fff3cd';
                        this.classList.add('hw-rows-modified');
                    } else {
                        this.style.backgroundColor = '';
                        this.classList.remove('hw-rows-modified');
                    }

                    updateHWRowsSaveButtonState();
                });
            });

            let buttonContainer = document.getElementById('hw-rows-buttons');
            if (!buttonContainer) {
                buttonContainer = document.createElement('div');
                buttonContainer.id = 'hw-rows-buttons';
                buttonContainer.className = 'mt-3 text-center';
                buttonContainer.innerHTML = `
                    <button id="save-hw-rows-btn" class="btn btn-success btn-sm me-2" onclick="saveHWRows()" disabled>
                        <i class="bi bi-check-circle"></i> Save HW Rows Changes
                    </button>
                    <button id="reset-hw-rows-btn" class="btn btn-secondary btn-sm" onclick="resetHWRows()" disabled>
                        <i class="bi bi-arrow-clockwise"></i> Reset Changes
                    </button>
                    <div id="hw-rows-status" class="text-muted mt-2" style="font-size: 0.875rem;"></div>
                `;

                const comparisonTable = document.getElementById('comparison-table');
                if (comparisonTable && comparisonTable.parentNode) {
                    comparisonTable.parentNode.insertBefore(buttonContainer, comparisonTable.nextSibling);
                }
            }

            updateHWRowsSaveButtonState();
        }

        function updateHWRowsSaveButtonState() {
            const modifiedSelects = document.querySelectorAll('.hw-rows-select.hw-rows-modified');
            const saveBtn = document.getElementById('save-hw-rows-btn');
            const resetBtn = document.getElementById('reset-hw-rows-btn');
            const statusDiv = document.getElementById('hw-rows-status');

            if (modifiedSelects.length > 0) {
                if (saveBtn) saveBtn.disabled = false;
                if (resetBtn) resetBtn.disabled = false;
                if (statusDiv) statusDiv.textContent = `${modifiedSelects.length} unit${modifiedSelects.length === 1 ? '' : 's'} modified`;
            } else {
                if (saveBtn) saveBtn.disabled = true;
                if (resetBtn) resetBtn.disabled = true;
                if (statusDiv) statusDiv.textContent = '';
            }
        }
        window.updateHWRowsSaveButtonState = updateHWRowsSaveButtonState;

        function saveHWRows() {
            const modifiedSelects = document.querySelectorAll('.hw-rows-select.hw-rows-modified');
            if (modifiedSelects.length === 0) {
                showToast('No changes to save', 'info');
                return;
            }

            const edits = Array.from(modifiedSelects).map(select => ({
                unit_tag: select.getAttribute('data-unit-tag'),
                hw_rows: parseInt(select.value, 10),
            }));

            if (!confirm(`Save HW Rows changes for ${edits.length} unit${edits.length === 1 ? '' : 's'}?`)) {
                return;
            }

            showToast('Saving HW Rows changes...', 'info');

            const pathInput = document.getElementById('original-tw2-path');
            const originalPath = pathInput ? sanitizeLocalPath(pathInput.value) : '';
            const payload = { edits, original_path: originalPath || '' };

            fetch('/save_hw_rows', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(payload),
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    const baseMessage = `Successfully updated ${data.updated_count} unit${data.updated_count === 1 ? '' : 's'}`;
                    const targetPath = data.target_path || '';
                    const successMessage = targetPath ? `${baseMessage}. Saved to: ${targetPath}` : baseMessage;
                    showToast(successMessage, 'success');

                    modifiedSelects.forEach(select => {
                        select.setAttribute('data-original', select.value);
                        select.style.backgroundColor = '';
                        select.classList.remove('hw-rows-modified');
                    });

                    updateHWRowsSaveButtonState();
                } else {
                    showToast(`Error saving changes: ${data.error}`, 'error');
                }
            })
            .catch(error => {
                showToast(`Error saving changes: ${error.message}`, 'error');
            });
        }
        window.saveHWRows = saveHWRows;

        function resetHWRows() {
            const modifiedSelects = document.querySelectorAll('.hw-rows-select.hw-rows-modified');
            if (modifiedSelects.length === 0) return;

            modifiedSelects.forEach(select => {
                select.value = select.getAttribute('data-original');
                select.style.backgroundColor = '';
                select.classList.remove('hw-rows-modified');
            });

            updateHWRowsSaveButtonState();
            showToast('Changes reset', 'info');
        }
        window.resetHWRows = resetHWRows;


// Display comparison results
        function displayComparisonResults(results, summary) {
            console.log('DISPLAY: displayComparisonResults called with:', results?.length, 'results');
            console.log('DISPLAY: Summary:', summary);
            const tableContainer = document.getElementById('comparison-table');
            console.log('DISPLAY: Table container found:', tableContainer);
            // Make sure the container sections are visible and the correct tab is active
            try {
                const preview = document.getElementById('data-preview-section');
                if (preview) preview.style.display = 'block';
                const tab = document.getElementById('comparison-tab');
                if (tab) tab.click();
            } catch (e) {}
            
            let headerHtml = `
                <thead>
                    <tr>
                        <th>Unit Tag</th>
                        <th>Status</th>
                        <th>Excel MBH</th>
                        <th>TW2 MBH</th>
                        <th>MBH Diff</th>
                        <th>Excel LAT</th>
                        <th>TW2 LAT</th>
                        <th>LAT Diff</th>
                        <th>WPD</th>
                        <th>APD</th>
                        <th>HW Rows</th>
                    </tr>
                </thead>
            `;
            
            let bodyHtml = '<tbody>';
            results.forEach((result, index) => {
                const statusClass = result.status === 'Pass' ? 'comparison-pass' : 
                                  result.status === 'Warning' ? 'comparison-warning' : 
                                  result.status === 'Fail' ? 'comparison-fail' : '';
                
                bodyHtml += `
                    <tr class="${statusClass}">
                        <td class="unit-tag">${result.unit_tag}</td>
                        <td class="status-${result.status.toLowerCase()}">${result.status}</td>
                        <td class="comparison-value">${formatNumber(result.excel_mbh) || 'N/A'}</td>
                        <td class="comparison-value">${formatNumber(result.tw2_mbh) || 'N/A'}</td>
                        <td class="percentage-diff">${result.mbh_diff}</td>
                        <td class="comparison-value">${formatNumber(result.excel_lat) || 'N/A'}</td>
                        <td class="comparison-value">${formatNumber(result.tw2_lat) || 'N/A'}</td>
                        <td class="percentage-diff">${result.lat_diff}</td>
                        <td class="comparison-value">${formatNumber(result.tw2_wpd) || 'N/A'}</td>
                        <td class="comparison-value">${formatNumber(result.tw2_apd) || 'N/A'}</td>
                        <td class="hw-rows-cell">
                            ${result.status !== 'Not Found' ?
                                `<select class="hw-rows-select"
                                        data-unit-tag="${result.unit_tag}"
                                        data-original="${result.tw2_hw_rows || 1}">
                                    <option value="1" ${(result.tw2_hw_rows || 1) == 1 ? 'selected' : ''}>1</option>
                                    <option value="2" ${(result.tw2_hw_rows || 1) == 2 ? 'selected' : ''}>2</option>
                                    <option value="3" ${(result.tw2_hw_rows || 1) == 3 ? 'selected' : ''}>3</option>
                                    <option value="4" ${(result.tw2_hw_rows || 1) == 4 ? 'selected' : ''}>4</option>
                                </select>`
                                : 'N/A'
                            }
                        </td>
                    </tr>
                `;
            });
            bodyHtml += '</tbody>';
            
            tableContainer.innerHTML = headerHtml + bodyHtml;
            
            setupHWRowsEditing();

            // Show results section
            document.getElementById('comparison-results').style.display = 'block';
            // Update summary badges if present
            try {
                const s = summary || {};
                const setText = (id, text) => { const el = document.getElementById(id); if (el) el.textContent = text; };
                if (typeof s.pass !== 'undefined') setText('summary-pass', `Pass: ${s.pass}`);
                if (typeof s.warning !== 'undefined') setText('summary-warning', `Warn: ${s.warning}`);
                if (typeof s.fail !== 'undefined') setText('summary-fail', `Fail: ${s.fail}`);
                if (typeof s.not_found !== 'undefined') setText('summary-notfound', `Not Found: ${s.not_found}`);
            } catch (e) {
                console.warn('DISPLAY: unable to update summary badges', e);
            }
            
            // Show acceptable ranges and summary
            const rangeInfo = document.createElement('div');
            rangeInfo.className = 'alert alert-info mb-3';
            rangeInfo.innerHTML = `
                <strong>Acceptable Ranges:</strong> 
                MBH/LAT: -${document.getElementById('mbh-lat-lower-margin').value}% to +${document.getElementById('mbh-lat-upper-margin').value}% | 
                WPD: ≤${document.getElementById('wpd-threshold').value} | 
                APD: ≤${document.getElementById('apd-threshold').value}
            `;
            
            const resultsContainer = document.getElementById('comparison-results');
            const existingAlert = resultsContainer.querySelector('.alert-info');
            if (existingAlert) existingAlert.remove();
            resultsContainer.insertBefore(rangeInfo, resultsContainer.firstChild);
            
            // Update summary in a toast
            showToast(`Summary: ${summary.pass} Pass, ${summary.warning} Warning, ${summary.fail} Fail, ${summary.not_found} Not Found`, 'info');
        }

        // Export comparison results
        function exportComparisonResults() {
            showToast('Export feature coming soon', 'info');
        }

        // Refresh and compare function
        function refreshAndCompare() {
            console.log('REFRESH: Starting refresh and compare operation');
            
            const refreshBtn = document.getElementById('refresh-btn');
            const originalText = refreshBtn.innerHTML;
            
            // Disable button and show loading state
            refreshBtn.disabled = true;
            refreshBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2" role="status"></span>Refreshing...';
            
            // Get current configuration values
            const config = {
                mbh_lat_lower_margin: parseFloat(document.getElementById('mbh-lat-lower-margin').value) || 15,
                mbh_lat_upper_margin: parseFloat(document.getElementById('mbh-lat-upper-margin').value) || 25,
                wpd_threshold: parseFloat(document.getElementById('wpd-threshold').value) || 5,
                apd_threshold: parseFloat(document.getElementById('apd-threshold').value) || 0.25,
                // Pass through the original path (if user provided it)
                original_path: sanitizeLocalPath((document.getElementById('original-tw2-path').value || ''))
            };
            
            console.log('REFRESH: Configuration:', config);
            console.log('REFRESH: Sending request to /refresh_and_compare');
            
            showToast('Refreshing TW2 file and running comparison...', 'info');
            
            fetch('/refresh_and_compare', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(config)
            })
            .then(response => {
                console.log('REFRESH: Response status:', response.status);
                console.log('REFRESH: Response headers:', [...response.headers.entries()]);
                return response.json();
            })
            .then(data => {
                console.log('REFRESH: Response data:', data);
                
                if (data.success) {
                    console.log('REFRESH: ✅ Success response received');
                    const payload = data.data || {};
                    if (payload.comparison_available) {
                        console.log('REFRESH: Comparison data available, displaying results');
                        console.log('REFRESH: About to call displayComparisonResults with data:', payload.results);
                        console.log('REFRESH: Summary data:', payload.summary);
                        try {
                            displayComparisonResults(payload.results, payload.summary);
                            console.log('REFRESH: ✅ displayComparisonResults completed successfully');
                        } catch (error) {
                            console.error('REFRESH: ❌ Error in displayComparisonResults:', error);
                        }
                        // Update path source indicator
                        try {
                            const el = document.getElementById('path-source-indicator');
                            if (el && payload.path_source) {
                                el.textContent = `Source: ${payload.path_source === 'original' ? 'Original TW2 file' : 'Local copy'}`;
                            }
                        } catch (e) {}
                        showToast('TW2 data refreshed and comparison completed successfully', 'success');
                    } else {
                        console.log('REFRESH: No comparison available:', payload.message);
                        showToast(payload.message, 'info');
                    }
                } else {
                    console.log('REFRESH: ❌ Error response:', data.error);
                    const debugInfo = (data.data && data.data.debug) ? data.data.debug : data.debug;
                    if (debugInfo) {
                        console.log('REFRESH: Debug info:', debugInfo);
                    }
                    showToast(`Refresh failed: ${data.error}`, 'error');
                }
            })
            .catch(error => {
                console.error('REFRESH: ❌ Fetch error:', error);
                showToast('Failed to refresh TW2 data', 'error');
            })
            .finally(() => {
                // Restore button state
                refreshBtn.disabled = false;
                refreshBtn.innerHTML = originalText;
            });
        }

        // Path validation function
        function validateTW2Path() {
            const pathInput = document.getElementById('original-tw2-path');
            const validateBtn = document.getElementById('validate-path-btn');
            const feedback = document.getElementById('path-validation-feedback');
            const path = sanitizeLocalPath(pathInput.value);
            
            if (!path) {
                showFeedback(feedback, 'Please enter a file path', 'warning');
                return;
            }
            
            // Update button state
            validateBtn.disabled = true;
            validateBtn.innerHTML = '<i class="bi bi-hourglass-split"></i> Validating...';
            
            console.log('PATH VALIDATION: Validating path:', path);
            
            fetch('/validate_tw2_path', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ path: path })
            })
            .then(response => response.json())
            .then(data => {
                if (data.valid) {
                    showFeedback(feedback, `✓ Valid TW2 file (${data.records} records, ${data.columns} columns)`, 'success');
                    pathInput.classList.remove('is-invalid');
                    pathInput.classList.add('is-valid');
                } else {
                    showFeedback(feedback, `✗ ${data.error}: ${data.details || ''}`, 'error');
                    pathInput.classList.remove('is-valid');
                    pathInput.classList.add('is-invalid');
                }
            })
            .catch(error => {
                console.error('PATH VALIDATION: Error:', error);
                showFeedback(feedback, '✗ Validation failed: Network error', 'error');
                pathInput.classList.remove('is-valid');
                pathInput.classList.add('is-invalid');
            })
            .finally(() => {
                validateBtn.disabled = false;
                validateBtn.innerHTML = '<i class="bi bi-check-circle"></i> Validate';
            });
        }
        
        // Helper function to show feedback
        function showFeedback(element, message, type) {
            element.style.display = 'block';
            element.className = `mt-1 small text-${type === 'success' ? 'success' : type === 'warning' ? 'warning' : 'danger'}`;
            element.textContent = message;
        }

        // Update existing Excel handler to enable comparison button
        const originalHandleExcelFile = handleExcelFile;
        handleExcelFile = function(file) {
            originalHandleExcelFile(file);
            // Update comparison button state after Excel is loaded
            setTimeout(updateComparisonButtonState, 1000);
        };
