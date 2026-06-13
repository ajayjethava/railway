// Add More flow JavaScript for Step 2 Locations
(function() {
    var unsavedTerminalData = {};
    var requestedJunctions = {{ junction_count }};
    var startJunctionNumber = {{ start_junction_number }};
    var stationIdSource = "{{ station_id_source if station_id_source else '' }}";
    var sheetName = "{{ sheet_name }}";
    var currentStep = {{ step }};
    var configModal, cableModal, successModal, errorModal, terminalModal, headerModal, groupModal, chokeModal, resistorModal;
    var currentJunctionIndex = null;
    var currentCables = [];
    var currentRowsConfig = [];
    var unsavedCableData = {};
    var currentCableIndex = null;
    var currentHeaders = [];
    var currentGroups = [];
    var currentChokes = [];
    var currentResistors = [];
   
    var savedCableIds = new Set();
    var isSavingAllCables = false;
    var currentJunctionConfig = null;
    var junctionCablesConfigured = {};
    var unsavedHeaderData = {};
    var unsavedGroupData = {};
    var unsavedChokeData = {};
    var unsavedResistorData = {};
    
    // Store current draft data
    var currentDraftData = null;
    var isDraftMode = false;
    var draftVersion = 1;

    // ============================================================================
    // 📋 HELPER FUNCTION TO PARSE CABLE NAME - UPDATED
    // ============================================================================
    function parseCableName(cableName) {
        // Parse cable name format like "A T1-12" to get terminal start and end
        const match = cableName.match(/([A-Z])\s+T(\d+)-(\d+)/);
        if (match) {
            return {
                terminalStart: parseInt(match[2], 10),
                terminalEnd: parseInt(match[3], 10)
            };
        }
        // Fallback to default values if parsing fails
        return { terminalStart: 1, terminalEnd: 12 };
    }


    // ============================================================================
    // 📋 HELPER FUNCTION TO CHECK IF CABLE HAS DRAFT TERMINALS
    // ============================================================================
    function hasDraftTerminalsForCable(cableId) {
        if (!cableId) return false;
        
        // Check local unsaved data
        if (unsavedTerminalData[cableId] && unsavedTerminalData[cableId].length > 0) {
            return true;
        }
        
        // Check database for draft data
        // Note: This would require a backend endpoint to check for draft terminals
        // For now, we'll rely on the local unsavedTerminalData
        
        return false;
}


    // ============================================================================
    // 📋 HELPER FUNCTION TO GET TERMINAL INFO FROM CABLE - NEW
    // ============================================================================
    function getTerminalInfoFromCable(cable) {
        // First try to get from cable properties
        if (cable.start_no && cable.terminal) {
            const terminalStart = parseInt(cable.start_no) || 1;
            const terminalEnd = terminalStart + (parseInt(cable.terminal) || 12) - 1;
            return { terminalStart, terminalEnd };
        }
        
        // Fallback to parsing cable name
        return parseCableName(cable.cable_name);
    }

    // ============================================================================
    // 🎉 COMPACT SUCCESS/ERROR POPUP NOTIFICATIONS
    // ============================================================================
    function showCompactNotification(message, type = 'success', duration = 2000) {
        var existingNotification = document.querySelector('.compact-notification');
        if (existingNotification) {
            existingNotification.remove();
        }
       
        var notificationDiv = document.createElement('div');
        notificationDiv.className = 'compact-notification';
       
        var bgColor = type === 'success' ? '#10b981' : '#ef4444';
        var icon = type === 'success' ? '✓' : '✕';
       
        notificationDiv.style.cssText = `
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background-color: ${bgColor};
            color: white;
            padding: 16px 32px;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 500;
            z-index: 10000;
            box-shadow: 0 4px 12px rgba(0,0,0,0.3);
            display: flex;
            align-items: center;
            gap: 12px;
            animation: popIn 0.3s ease-in-out;
            min-width: 250px;
        `;
       
        notificationDiv.innerHTML = `
            <span style="font-size: 22px; font-weight: bold;">${icon}</span>
            <span>${message}</span>
        `;
       
        document.body.appendChild(notificationDiv);
       
        setTimeout(() => {
            notificationDiv.style.animation = 'popOut 0.3s ease-in-out';
            setTimeout(() => {
                notificationDiv.remove();
            }, 300);
        }, duration);
    }

    // ============================================================================
    // 📋 CABLE CONFIGURATION SAVE FUNCTION
    // ============================================================================
    function saveCableConfigurationToDatabase(configRows, junctionName) {
        var projectId = {{ current_project.id }};
        
        // Get the junction ID from the current configuration
        // We need to extract it from the DOM or pass it from the backend
        var junctionId = getJunctionIdForCurrentConfig();
        
        var configData = {
            project_id: projectId,
            junction_box_id: junctionId,  // Use junction ID instead of name
            config_rows: configRows,
            is_draft: false,
            draft_version: 0
        };
        
        console.log('Saving cable config with junction_box_id:', junctionId, 'for junction:', junctionName);
        
        return fetch('/save_cable_configuration', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(configData)
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .then(result => {
            if (result.success) {
                showCompactNotification('✓ Cable configuration saved successfully!', 'success', 2000);
                return result;
            } else {
                throw new Error(result.message || 'Failed to save cable configuration');
            }
        })
        .catch(error => {
            console.error('Error saving cable configuration:', error);
            showCompactNotification('✕ Error saving configuration: ' + error.message, 'error', 3000);
            throw error;
        });
    }

    // ============================================================================
    // 📋 HELPER FUNCTION TO GET JUNCTION ID
    // ============================================================================
    function getJunctionIdForCurrentConfig() {
        if (!currentJunctionConfig) {
            console.error('No current junction configuration found');
            return null;
        }
        
        // Try to get junction ID from the DOM
        var junctionIndex = currentJunctionConfig.junctionIndex;
        
        // Look for hidden input containing junction_id
        var junctionIdInput = document.querySelector(`input[name="junctions[${junctionIndex}][junction_id]"]`);
        
        if (junctionIdInput && junctionIdInput.value) {
            return junctionIdInput.value;
        }
        
        // If not found, try to extract from junction name pattern
        var junctionName = currentJunctionConfig.junctionName;
        var match = junctionName.match(/^LOC(\d+)/);
        if (match) {
            return match[1]; // Return the numeric part
        }
        
        // Last resort: use junction index
        return ({{ start_junction_number }} + junctionIndex).toString();
    }

    // ============================================================================
    // 📋 DRAFT MANAGEMENT FUNCTIONS
    // ============================================================================
    function saveCableConfigDraft() {
    
    // Collect data from the table
    var configRows = [];
    var isValid = true;
    
    document.querySelectorAll('#configTableBody tr').forEach((row, index) => {
        var rowName = row.querySelector('.config-row-name').value.trim();
        var cableType = row.querySelector('.config-cable-type').value;
        var cableCount = parseInt(row.querySelector('.config-cable-count').value);
        
        if (!rowName || isNaN(cableCount) || cableCount < 1) {
            isValid = false;
            return;
        }
        
        configRows.push({
            row_number: index + 1,
            location_row_name: rowName,
            cable_type: cableType,
            number_of_cables: cableCount
        });
    });
    
   
    
    var draftData = {
        project_id: {{ current_project.id }},
        junction_box_id: currentJunctionConfig.junctionName,
        config_rows: configRows,
        is_draft: true,
        draft_version: draftVersion
    };
    
    var saveDraftBtn = document.getElementById('saveDraftBtn');
    var originalText = saveDraftBtn.innerHTML;
    saveDraftBtn.disabled = true;
    saveDraftBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving...';
    
    fetch('/save_cable_config_draft', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(draftData)
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        return response.json();
    })
    .then(result => {
        saveDraftBtn.disabled = false;
        saveDraftBtn.innerHTML = originalText;
        
        if (result.success) {
            showCompactNotification('✓ Draft saved successfully!', 'success', 2000);
            currentDraftData = draftData;
            isDraftMode = true;
            
            if (result.draft_version) {
                draftVersion = result.draft_version;
            }
        } else {
            showCompactNotification('✕ ' + (result.message || 'Failed to save draft.'), 'error', 3000);
        }
    })
    .catch(error => {
        saveDraftBtn.disabled = false;
        saveDraftBtn.innerHTML = originalText;
        //showCompactNotification('✕ Error saving draft: ' + error.message, 'error', 3000);
        console.error('Error saving draft:', error);
    });
}
function loadCableConfigDraft(junctionBoxId) {
    var projectId = {{ current_project.id }};
    
    fetch(`/get_cable_config_draft?project_id=${projectId}&junction_box_id=${encodeURIComponent(junctionBoxId)}`)
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .then(result => {
            if (result.success && result.config_rows && result.config_rows.length > 0) {
                // Store draft data
                isDraftMode = true;
                draftVersion = result.draft_version || 1;
                
                // Fill the table with draft data
                var tableBody = document.getElementById('configTableBody');
                if (tableBody && result.config_rows.length > 0) {
                    // Clear existing rows
                    tableBody.innerHTML = '';
                    
                    // Create rows from draft data
                    result.config_rows.forEach((rowConfig, index) => {
                        var row = document.createElement('tr');
                        row.innerHTML = `
                            <td><strong>${rowConfig.row_number || index + 1}</strong></td>
                            <td>
                                <input type="text" class="form-control config-row-name" value="${rowConfig.location_row_name || String.fromCharCode(65 + index)}" placeholder="Enter row name (e.g., A, B, C...)" required>
                            </td>
                            <td>
                                <select class="form-select config-cable-type" required>
                                    <option value="cable" ${rowConfig.cable_type === 'cable' ? 'selected' : ''}>Cable</option>
                                    <option value="relay_box" ${rowConfig.cable_type === 'relay_box' ? 'selected' : ''}>Relay Box</option>
                                </select>
                            </td>
                            <td>
                                <input type="number" class="form-control config-cable-count" value="${rowConfig.number_of_cables || 1}" min="1" required>
                            </td>
                        `;
                        tableBody.appendChild(row);
                    });
                    
                    showCompactNotification('✓ Loaded draft configuration', 'success', 2000);
                }
            }
        })
        .catch(error => {
            console.error('Error loading draft:', error);
            // No draft found or error - continue with default rows
        });
}
   
    function clearCableConfigDraft(junctionBoxId) {
        var projectId = {{ current_project.id }};
        
        fetch(`/clear_cable_config_draft?project_id=${projectId}&junction_box_id=${encodeURIComponent(junctionBoxId)}`, {
            method: 'POST'
        })
        .then(response => response.json())
        .then(result => {
            if (result.success) {
                currentDraftData = null;
                isDraftMode = false;
                draftVersion = 1;
                showCompactNotification('✓ Draft cleared successfully', 'success', 2000);
            }
        })
        .catch(error => {
            console.error('Error clearing draft:', error);
        });
    }
    // ============================================================================
    // 📋 CABLE TABLE DRAFT MANAGEMENT FUNCTIONS - UPDATED
    // ============================================================================
    function saveCableTableDraft() {
        if (!currentJunctionConfig) {
            showCompactNotification('✕ No junction configuration found.', 'error', 3000);
            return;
        }
        
        const junctionBoxId = currentJunctionConfig.junctionName;
        const junctionBoxName = junctionBoxId;
        
        // Collect cable data in the format expected by the backend
        const cableData = [];
        
        // Collect data from each cable table
        document.querySelectorAll('[id^="cableTableBody-"]').forEach((tableBody, tableIndex) => {
            const rowName = currentRowsConfig[tableIndex]?.rowName || `Row ${tableIndex + 1}`;
            
            tableBody.querySelectorAll('tr').forEach((row, cableIndex) => {
                // Get values from the table row
                const positionInput = row.querySelector('td:nth-child(1) input');
                const locationBoxInput = row.querySelector('td:nth-child(5) input');
                const cableIdInput = row.querySelector('td:nth-child(6) input');
                const terminalInput = row.querySelector('td:nth-child(7) input');
                const startNoInput = row.querySelector('td:nth-child(8) input');
                const cableNameInput = row.querySelector('td:nth-child(9) input');
                
                if (positionInput && locationBoxInput && cableIdInput) {
                    const cable = {
                        row: rowName,
                        position: positionInput.value || '',
                        terminal: terminalInput ? terminalInput.value : '',
                        start_no: startNoInput ? startNoInput.value : '',
                        cable_id: cableIdInput.value || `${junctionBoxId}-${rowName}-${positionInput.value}`,
                        cable_name: cableNameInput ? cableNameInput.value : `Cable ${rowName}${positionInput.value}`
                    };
                    cableData.push(cable);
                }
            });
        });
        
        if (cableData.length === 0) {
            showCompactNotification('✕ No cable data to save as draft.', 'error', 3000);
            return;
        }
        
        const draftData = {
            junction_box_id: junctionBoxId,
            junction_box_name: junctionBoxName,
            cable_data: cableData
        };
        
        console.log('Saving cable draft:', draftData);
        
        const saveDraftBtn = document.getElementById('saveCableTableDraftBtn');
        const originalText = saveDraftBtn.innerHTML;
        saveDraftBtn.disabled = true;
        saveDraftBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving...';
        
        // Use the correct endpoint
        fetch('/save_cable_table_draft', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(draftData)
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .then(result => {
            saveDraftBtn.disabled = false;
            saveDraftBtn.innerHTML = originalText;
            
            if (result.success) {
                showCompactNotification('✓ Cable table draft saved successfully!', 'success', 2000);
                console.log('Draft saved:', result);
                
                // Update summary display
                updateCableSummaryDisplay();
                
                // Update cable statuses if needed
                updateCableStatuses();
            } else {
                showCompactNotification('✕ ' + (result.message || 'Failed to save draft.'), 'error', 3000);
            }
        })
        .catch(error => {
            saveDraftBtn.disabled = false;
            saveDraftBtn.innerHTML = originalText;
            showCompactNotification('✕ Error saving cable table draft: ' + error.message, 'error', 3000);
            console.error('Error saving cable table draft:', error);
        });
    }

    function loadCableTableDraft(junctionBoxId) {
        if (!junctionBoxId) {
            console.log('No junction box ID provided for loading cable table draft');
            return;
        }
        
        fetch(`/get_cable_table_draft?junction_box_id=${encodeURIComponent(junctionBoxId)}`)
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                return response.json();
            })
            .then(result => {
                if (result.success && result.cable_data && result.cable_data.length > 0) {
                    console.log(`Loaded cable table draft with ${result.cable_data.length} cables for junction ${junctionBoxId}`);
                    
                    // Store the draft data for population
                    populateCableTablesFromDraft(result.cable_data);
                    showCompactNotification('✓ Loaded cable table draft', 'success', 2000);
                } else {
                    console.log('No cable table draft found for this junction');
                }
            })
            .catch(error => {
                console.error('Error loading cable table draft:', error);
                // No draft found or error - continue with default
            });
    }

    function populateCableTablesFromDraft(cableData) {
        console.log('Populating cable tables from draft:', cableData);
        
        // Clear current cables array and repopulate
        currentCables = [];
        
        cableData.forEach(draftCable => {
            // Find the corresponding row table
            const rowIndex = currentRowsConfig.findIndex(row => row.rowName === draftCable.row);
            
            if (rowIndex !== -1) {
                const tableBody = document.getElementById(`cableTableBody-${rowIndex}`);
                if (tableBody) {
                    // Find or create row for this position
                    let foundRow = null;
                    const rows = tableBody.querySelectorAll('tr');
                    
                    // Check if a row with this position already exists
                    rows.forEach(row => {
                        const positionInput = row.querySelector('td:nth-child(1) input');
                        if (positionInput && positionInput.value === draftCable.position) {
                            foundRow = row;
                        }
                    });
                    
                    if (foundRow) {
                        // Update existing row
                        const locationBoxInput = foundRow.querySelector('td:nth-child(5) input');
                        const cableIdInput = foundRow.querySelector('td:nth-child(6) input');
                        const terminalInput = foundRow.querySelector('td:nth-child(7) input');
                        const startNoInput = foundRow.querySelector('td:nth-child(8) input');
                        const cableNameInput = foundRow.querySelector('td:nth-child(9) input');
                        
                        if (locationBoxInput) locationBoxInput.value = draftCable.junction_box || '';
                        if (cableIdInput) cableIdInput.value = draftCable.cable_id;
                        if (terminalInput) terminalInput.value = draftCable.terminal || '';
                        if (startNoInput) startNoInput.value = draftCable.start_no || '';
                        if (cableNameInput) cableNameInput.value = draftCable.cable_name || '';
                    }
                    
                    // Add to currentCables array
                    const cable = {
                        position: draftCable.position,
                        junction_name: currentJunctionConfig?.junctionName || '',
                        junction_size: currentJunctionConfig?.junctionSize || '',
                        row: draftCable.row,
                        junction_box: draftCable.junction_box || '',
                        cable_id: draftCable.cable_id,
                        terminal: draftCable.terminal || '',
                        start_no: draftCable.start_no || '',
                        cable_name: draftCable.cable_name || '',
                        cable_type: draftCable.cable_type || 'cable'
                    };
                    
                    currentCables.push(cable);
                }
            }
        });
        
        // Update cable statuses
        updateCableStatuses();
    }

    function clearCableTableDraft(junctionBoxId) {
        if (!confirm('Are you sure you want to clear the cable table draft? This action cannot be undone.')) {
            return;
        }
        
        fetch('/clear_cable_table_draft', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({
                project_id: {{ current_project.id }},
                junction_box_id: junctionBoxId
            })
        })
        .then(response => response.json())
        .then(result => {
            if (result.success) {
                showCompactNotification('✓ Cable table draft cleared successfully', 'success', 2000);
                
                // Update summary display
                updateCableSummaryDisplay();
            } else {
                showCompactNotification('✕ ' + (result.message || 'Failed to clear draft.'), 'error', 3000);
            }
        })
        .catch(error => {
            console.error('Error clearing cable table draft:', error);
            showCompactNotification('✕ Error clearing cable table draft', 'error', 3000);
        });
    }

    // Function to update cable summary display
    function updateCableSummaryDisplay() {
        fetch('/get_cable_summary')
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                return response.json();
            })
            .then(result => {
                if (result.success) {
                    // Update summary display elements if they exist
                    const totalCablesEl = document.getElementById('totalCablesCount');
                    const totalRowsEl = document.getElementById('totalRowsCount');
                    const totalJunctionsEl = document.getElementById('totalJunctionsCount');
                    
                    if (totalCablesEl) totalCablesEl.textContent = result.total_cables || 0;
                    if (totalRowsEl) totalRowsEl.textContent = result.total_rows || 0;
                    if (totalJunctionsEl) totalJunctionsEl.textContent = result.total_junctions || 0;
                }
            })
            .catch(error => {
                console.error('Error updating cable summary:', error);
            });
    }

    // ============================================================================
    // 📋 UPDATED: SAVE ALL CABLES FUNCTION
    // ============================================================================
    function saveAllCables() {
        if (!currentJunctionConfig) {
            showCompactNotification('✕ No junction configuration found.', 'error', 3000);
            return;
        }
        
        const junctionBoxId = currentJunctionConfig.junctionName;
        const junctionBoxName = junctionBoxId;
        
        // First save as draft
        saveCableTableDraft();
        
        // Then finalize the draft
        setTimeout(() => {
            const saveAllBtn = document.getElementById('saveAllCablesBtn');
            const originalText = saveAllBtn.innerHTML;
            saveAllBtn.disabled = true;
            saveAllBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Finalizing...';
            
            fetch('/finalize_cable_table', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({
                    junction_box_id: junctionBoxId
                })
            })
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                return response.json();
            })
            .then(result => {
                saveAllBtn.disabled = false;
                saveAllBtn.innerHTML = originalText;
                
                if (result.success) {
                    showCompactNotification('✓ All cables saved and finalized successfully!', 'success', 2000);
                    
                    // Update summary display
                    updateCableSummaryDisplay();
                    
                    // Close modal
                    if (cableModal) {
                        cableModal.hide();
                    }
                    
                    // Optional: Reload or redirect
                    setTimeout(() => {
                        window.location.reload();
                    }, 1500);
                } else {
                    showCompactNotification('✕ ' + (result.message || 'Failed to finalize cables.'), 'error', 3000);
                }
            })
            .catch(error => {
                saveAllBtn.disabled = false;
                saveAllBtn.innerHTML = originalText;
                showCompactNotification('✕ Error finalizing cables: ' + error.message, 'error', 3000);
                console.error('Error finalizing cables:', error);
            });
        }, 500); // Small delay to ensure draft is saved first
    }

    // ============================================================================
    // 📋 UPDATED: OPEN CABLE TABLE MODAL - LOAD DRAFT
    // ============================================================================
    function openCableTableModal(junctionIndex) {
        var junction = {};
        var formData = new FormData(document.getElementById('addMoreJunctionGridForm'));
        formData.forEach((value, key) => {
            if (key.startsWith(`junctions[${junctionIndex}]`)) {
                var field = key.match(/junctions\[\d+\]\[(.*)\]/)[1];
                junction[field] = value;
            }
        });
        
        var junctionName = junction.junction_name || junction.junction_id || 'LOC' + (junction.junction_id || (startJunctionNumber + junctionIndex));
        var junctionSize = junction.junction_size || 'Not Set';
        
        // Check if junctionName already has (F) or (H) suffix
        // If not, add it based on junctionSize
        var junctionNameWithSize = junctionName;
        
        // Remove any existing (F) or (H) suffix first to check
        var cleanName = junctionName.replace(/\s*\([FH]\)$/, '');
        
        // Only add suffix if it's not already there
        if (!junctionName.endsWith(' (F)') && !junctionName.endsWith(' (H)')) {
            if (junctionSize === 'Full') {
                junctionNameWithSize = cleanName + ' (F)';
            } else if (junctionSize === 'Half') {
                junctionNameWithSize = cleanName + ' (H)';
            }
        }
        
        // Get Location ID for junction_box
        var locationBox = parseInt(junction.junction_id) || (startJunctionNumber + junctionIndex);
        
        document.getElementById('cableModalJunctionName').textContent = junctionNameWithSize;
        var rowInfoElement = document.getElementById('cableModalRowInfo');
        rowInfoElement.innerHTML = '';
        
        currentRowsConfig.forEach((rowConfig, index) => {
            var rowInfo = document.createElement('div');
            rowInfo.className = 'fw-bold';
            rowInfo.innerHTML = `Row: ${rowConfig.rowName} | Type: ${rowConfig.cableType} | Cables: ${rowConfig.cableCount}`;
            rowInfoElement.appendChild(rowInfo);
        });
        
        var container = document.getElementById('cableTablesContainer');
        container.innerHTML = '';
        currentCables = [];
        
        fetch('/get_existing_cables')
            .then(response => {
                if (!response.ok) {
                    return [];
                }
                return response.json();
            })
            .then(existingCables => {
                // Convert old cable_box to relay_box
                if (existingCables && existingCables.length > 0 && !existingCables.error) {
                    existingCables.forEach(cable => {
                        if (cable.cable_type === 'cable_box') {
                            cable.cable_type = 'relay_box';
                        }
                    });
                }
                
                var maxCableId = 0;
                if (existingCables && existingCables.length > 0 && !existingCables.error) {
                    existingCables.forEach(cable => {
                        if (cable.cable_id) {
                            var id = parseInt(cable.cable_id);
                            if (!isNaN(id) && id > maxCableId) {
                                maxCableId = id;
                            }
                        }
                    });
                }
                var nextCableId = maxCableId + 1;
            
                currentRowsConfig.forEach((rowConfig, rowIndex) => {
                    var rowTable = document.createElement('div');
                    rowTable.className = 'card mb-4';
                    rowTable.innerHTML = `
                        <div class="card-header bg-light">
                            <h6 class="mb-0">Row: ${rowConfig.rowName} | Type: ${rowConfig.cableType} | Cables: ${rowConfig.cableCount}</h6>
                        </div>
                        <div class="card-body">
                            <div class="table-responsive">
                                <table class="table table-bordered table-hover">
                                    <thead class="table-light">
                                        <tr>
                                            <th>Cable Position</th>
                                            <th>Location Name</th>
                                            <th>Location Size</th>
                                            <th>Row</th>
                                            <th>Location Box</th>
                                            <th>Cable Id</th>
                                            <th>Terminal</th>
                                            <th>Start No</th>
                                            <th>Cable Name</th>
                                            <th>Cable Type</th>
                                            <th>Status</th>
                                            <th>Actions</th>
                                        </tr>
                                    </thead>
                                    <tbody id="cableTableBody-${rowIndex}">
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    `;
                    container.appendChild(rowTable);
                    var rowTableBody = document.getElementById(`cableTableBody-${rowIndex}`);
                    var currentStartNo = 1;
                    var rowPosition = 1;
                
                    for (var pos = 1; pos <= rowConfig.cableCount; pos++) {
                        var existingCable = null;
                        if (existingCables && existingCables.length > 0) {
                            existingCable = existingCables.find(cable =>
                                (cable.junction_name === junctionName || cable.junction_name === junctionNameWithSize) &&
                                cable.row === rowConfig.rowName &&
                                parseInt(cable.junction_box) === locationBox &&
                                cable.cable_type === rowConfig.cableType &&
                                parseInt(cable.position) === rowPosition
                            );
                        }
                        var cableId, startNo, terminal;
                        if (existingCable) {
                            cableId = parseInt(existingCable.cable_id);
                            startNo = parseInt(existingCable.start_no) || 1;
                            terminal = parseInt(existingCable.terminal) || 12;
                            currentStartNo = startNo + terminal;
                        } else {
                            cableId = nextCableId;
                            if (rowConfig.cableType === 'relay_box') {
                                startNo = pos;
                            } else {
                                startNo = currentStartNo;
                            }
                            terminal = 12;
                            if (rowConfig.cableType !== 'relay_box') {
                                currentStartNo = startNo + terminal;
                            }
                            nextCableId++;
                        }
                        var cableName;
                        if (rowConfig.cableType === 'relay_box') {
                            cableName = existingCable ? existingCable.cable_name : '';
                        } else {
                            var endNo = startNo + terminal - 1;
                            cableName = rowConfig.rowName + ' T' + startNo + '-' + endNo;
                        }
                        
                        var cable = {
                            position: rowPosition,
                            junction_name: junctionNameWithSize,  // Use the properly formatted name here
                            junction_size: junctionSize,
                            row: rowConfig.rowName,
                            junction_box: locationBox, // Use Location ID instead of pos
                            cable_id: cableId,
                            terminal: terminal,
                            start_no: startNo,
                            cable_name: cableName,
                            cable_type: rowConfig.cableType
                        };
                        currentCables.push(cable);
                        var cableRow = document.createElement('tr');
                        var startNoDisplay = startNo;
                        var cableNameReadonly = '';
                    
                        if (rowConfig.cableType === 'relay_box') {
                            startNoDisplay = startNo.toString().padStart(2, '0');
                            cableNameReadonly = '';
                        } else {
                            cableNameReadonly = 'readonly';
                        }

                        var buttonHtml = rowConfig.cableType !== 'relay_box' ? `
                            <button type="button" class="btn btn-sm btn-info config-terminals-btn" data-index="${currentCables.length-1}">
                                <i class="bi bi-gear"></i> Config Term
                            </button>
                        ` : '';
                    
                        cableRow.innerHTML = `
                            <td>
                                <input type="text" class="form-control" value="${rowPosition}" readonly>
                            </td>
                            <td>
                                <input type="text" class="form-control" value="${junctionNameWithSize}" readonly>
                            </td>
                            <td>
                                <input type="text" class="form-control" value="${junctionSize}" readonly>
                            </td>
                            <td>
                                <input type="text" class="form-control cable-row-input" data-index="${currentCables.length-1}" value="${rowConfig.rowName}" readonly>
                            </td>
                            <td>
                                <input type="text" class="form-control cable-junction-box-input" data-index="${currentCables.length-1}" value="${locationBox}" readonly>
                            </td>
                            <td>
                                <input type="text" class="form-control cable-id-input" data-index="${currentCables.length-1}" value="${cable.cable_id}" readonly>
                            </td>
                            <td>
                                <input type="number" class="form-control cable-terminal-input" data-index="${currentCables.length-1}" value="${cable.terminal}">
                            </td>
                            <td>
                                <input type="number" class="form-control cable-start-no-input" data-index="${currentCables.length-1}" value="${startNoDisplay}">
                            </td>
                            <td>
                                <input type="text" class="form-control cable-name-input" data-index="${currentCables.length-1}" value="${cableName}" ${cableNameReadonly}>
                            </td>
                            <td>
                                <input type="text" class="form-control cable-type-input" data-index="${currentCables.length-1}" value="${rowConfig.cableType}" readonly>
                            </td>
                            <td>
                                <span id="status-${currentCables.length-1}" class="cable-status-indicator">
                                    <i class="bi bi-x-circle-fill text-danger" title="Terminals not configured"></i>
                                </span>
                            </td>
                            <td>
                                ${buttonHtml}
                            </td>
                        `;
                        rowTableBody.appendChild(cableRow);
                        rowPosition++;
                    }
                });
                
                // After creating the tables, load draft data if exists
                loadCableTableDraft(junctionName);
                
                updateCableStatuses();
                
                document.querySelectorAll('.cable-terminal-input').forEach(input => {
                    input.addEventListener('input', function() {
                        var index = parseInt(this.dataset.index);
                        currentCables[index].terminal = this.value;
                        if (currentCables[index].cable_type !== 'relay_box') {
                            updateCableName(index);
                        }
                        updateSubsequentStartNos(index);
                    });
                });
                document.querySelectorAll('.cable-start-no-input').forEach(input => {
                    input.addEventListener('input', function() {
                        var index = parseInt(this.dataset.index);
                        var cable = currentCables[index];
                        var newStartNo = parseInt(this.value) || 1;
                    
                        if (cable.cable_type === 'relay_box') {
                            this.value = newStartNo.toString().padStart(2, '0');
                            cable.start_no = newStartNo;
                        } else {
                            cable.start_no = newStartNo;
                            updateCableName(index);
                        }
                        updateSubsequentStartNos(index);
                    });
                });
                document.querySelectorAll('.cable-name-input').forEach(input => {
                    input.addEventListener('input', function() {
                        var index = parseInt(this.dataset.index);
                        var cable = currentCables[index];
                        if (cable.cable_type === 'relay_box') {
                            cable.cable_name = this.value;
                        }
                    });
                });
                
                function updateSubsequentStartNos(changedIndex) {
                    var changedCable = currentCables[changedIndex];
                    if (changedCable.cable_type === 'relay_box') {
                        return;
                    }
                
                    var rowName = changedCable.row;
                    var currentStartNo = parseInt(changedCable.start_no) || 1;
                    var terminalCount = parseInt(changedCable.terminal) || 12;
                
                    var subsequentCables = currentCables.filter((cable, idx) =>
                        cable.row === rowName &&
                        cable.cable_type !== 'relay_box' &&
                        cable.position > changedCable.position
                    ).sort((a, b) => a.position - b.position);
                
                    var nextStartNo = currentStartNo + terminalCount;
                    subsequentCables.forEach(cable => {
                        cable.start_no = nextStartNo;
                    
                        var startNoInput = document.querySelector(`.cable-start-no-input[data-index="${currentCables.indexOf(cable)}"]`);
                        if (startNoInput) {
                            startNoInput.value = nextStartNo;
                        }
                    
                        if (cable.cable_type !== 'relay_box') {
                            updateCableName(currentCables.indexOf(cable));
                        }
                    
                        nextStartNo += parseInt(cable.terminal) || 12;
                    });
                }
                
                setTimeout(function() {
                    document.querySelectorAll('.config-terminals-btn').forEach(button => {
                        button.replaceWith(button.cloneNode(true));
                    });
                
                    document.querySelectorAll('.config-terminals-btn').forEach(button => {
                        button.addEventListener('click', function() {
                            var index = parseInt(this.dataset.index);
                            currentCableIndex = index;
                            var cable = currentCables[index];
                            var terminalCount = parseInt(document.querySelector(`.cable-terminal-input[data-index="${index}"]`).value) || 0;
                            if (terminalCount < 1) {
                                showCompactNotification('✕ Terminal count must be at least 1 to configure.', 'error', 3000);
                                return;
                            }
                            if (savedCableIds.has(cable.cable_id)) {
                                openTerminalModal(cable, terminalCount);
                            } else {
                                saveCableAndOpenTerminals(cable, terminalCount, button);
                            }
                        });
                    });
                }, 100);
                
                cableModal.show();
            })
            .catch(error => {
                console.error('Error fetching existing cables:', error);
            });
    }

    // Initialize summary display on page load
    document.addEventListener('DOMContentLoaded', function() {
        updateCableSummaryDisplay();
    });    
    // ============================================================================
    // 📋 TERMINAL DRAFT MANAGEMENT FUNCTIONS
    // ============================================================================
    function saveTerminalDraft() {
        if (!currentCableIndex && currentCableIndex !== 0) {
            showCompactNotification('✕ No cable selected for saving draft.', 'error', 3000);
            return;
        }
        
        var cable = currentCables[currentCableIndex];
        var cableId = cable.cable_id;
        
        if (!cableId) {
            showCompactNotification('✕ No cable ID available to save draft.', 'error', 3000);
            return;
        }
        
        // Collect terminal data from the table
        var terminalData = [];
        document.querySelectorAll('#terminalTableBody tr').forEach((row, idx) => {
            terminalData.push({
                cable_id: cableId,
                terminal_id: row.querySelector('.terminal-id-input').value,
                terminal_no: row.querySelector('.terminal-name-input').value,
                symbol: row.querySelector('.symbol-input').value,
                input_left: row.querySelector('.input-left-input').value,
                input_right: row.querySelector('.input-right-input').value,
                spare: row.querySelector('.spare-input').value,
                input_connected: row.querySelector('.input-connected-input').value,
                output_connected: row.querySelector('.output-connected-input').value,
                input_connected_extra: row.querySelector('.input-connected-extra-input').value,
                output_connected_extra: row.querySelector('.output-connected-extra-input').value,
                output_left: row.querySelector('.output-left-input').value,
                output_right: row.querySelector('.output-right-input').value
            });
        });
        
        if (terminalData.length === 0) {
            showCompactNotification('✕ No terminal data to save as draft.', 'error', 3000);
            return;
        }
        
        var draftData = {
            project_id: {{ current_project.id }},
            junction_box_id: currentJunctionConfig?.junctionName || '',
            cable_id: cableId,
            cable_name: cable.cable_name,
            terminal_data: terminalData,
            is_draft: true,
            draft_type: 'terminal_config'
        };
        
        var saveDraftBtn = document.getElementById('saveTerminalDraftBtn');
        var originalText = saveDraftBtn.innerHTML;
        saveDraftBtn.disabled = true;
        saveDraftBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving...';
        
        fetch('/save_terminal_draft', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(draftData)
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .then(result => {
            saveDraftBtn.disabled = false;
            saveDraftBtn.innerHTML = originalText;
            
            if (result.success) {
                showCompactNotification('✓ Terminal draft saved successfully!', 'success', 2000);
                // Store unsaved data locally as well
                unsavedTerminalData[cableId] = terminalData;

                // NEW: Update the cable status in the cable table
                var cable = currentCables[currentCableIndex];
                if (cable) {
                    updateCableStatusForSingleCable(cable.cable_id, currentCableIndex);
                }
            } else {
                showCompactNotification('✕ ' + (result.message || 'Failed to save draft.'), 'error', 3000);
            }
        })
        .catch(error => {
            saveDraftBtn.disabled = false;
            saveDraftBtn.innerHTML = originalText;
            showCompactNotification('✕ Error saving terminal draft: ' + error.message, 'error', 3000);
            console.error('Error saving terminal draft:', error);
        });
    }

    function loadTerminalDraft(cableId) {
        var projectId = {{ current_project.id }};
        
        if (!cableId) {
            console.log('No cable ID provided for loading terminal draft');
            return;
        }
        
        fetch(`/get_terminal_draft?project_id=${projectId}&cable_id=${encodeURIComponent(cableId)}`)
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                return response.json();
            })
            .then(result => {
                if (result.success && result.terminal_data && result.terminal_data.length > 0) {
                    console.log(`Loaded terminal draft with ${result.terminal_data.length} terminals for cable ${cableId}`);
                    
                    // Store the draft data in unsavedTerminalData for use in openTerminalModal
                    unsavedTerminalData[cableId] = result.terminal_data;
                    
                    // If terminal modal is already open, populate it
                    if (terminalModal && terminalModal._isShown) {
                        populateTerminalTableFromDraft(result.terminal_data);
                    }
                }
            })
            .catch(error => {
                console.error('Error loading terminal draft:', error);
                // No draft found or error - continue with default
            });
    }

    function populateTerminalTableFromDraft(terminalData) {
        var tableBody = document.getElementById('terminalTableBody');
        if (!tableBody) return;
        
        var rows = tableBody.querySelectorAll('tr');
        
        terminalData.forEach((terminal, idx) => {
            if (idx < rows.length) {
                var row = rows[idx];
                
                // Populate each field
                if (row.querySelector('.terminal-id-input')) {
                    row.querySelector('.terminal-id-input').value = terminal.terminal_id || (idx + 1);
                }
                if (row.querySelector('.terminal-name-input')) {
                    row.querySelector('.terminal-name-input').value = terminal.terminal_no || '';
                }
                if (row.querySelector('.symbol-input')) {
                    row.querySelector('.symbol-input').value = terminal.symbol || 'ara/wago';
                }
                if (row.querySelector('.input-left-input')) {
                    row.querySelector('.input-left-input').value = terminal.input_left || '';
                }
                if (row.querySelector('.input-right-input')) {
                    row.querySelector('.input-right-input').value = terminal.input_right || '';
                }
                if (row.querySelector('.spare-input')) {
                    row.querySelector('.spare-input').value = terminal.spare || 'N';
                }
                if (row.querySelector('.input-connected-input')) {
                    row.querySelector('.input-connected-input').value = terminal.input_connected || 'Y';
                }
                if (row.querySelector('.output-connected-input')) {
                    row.querySelector('.output-connected-input').value = terminal.output_connected || 'Y';
                }
                if (row.querySelector('.input-connected-extra-input')) {
                    row.querySelector('.input-connected-extra-input').value = terminal.input_connected_extra || '';
                }
                if (row.querySelector('.output-connected-extra-input')) {
                    row.querySelector('.output-connected-extra-input').value = terminal.output_connected_extra || '';
                }
                if (row.querySelector('.output-left-input')) {
                    row.querySelector('.output-left-input').value = terminal.output_left || '';
                }
                if (row.querySelector('.output-right-input')) {
                    row.querySelector('.output-right-input').value = terminal.output_right || '';
                }
            }
        });
        
        // Re-add event listeners for Y/N fields
        addTerminalYNEventListeners();
        
        // NEW: Update the cable status since we have loaded draft terminals
        var cable = currentCables[currentCableIndex];
        if (cable) {
            updateCableStatusImmediately(currentCableIndex);
        }
    }

    function clearTerminalDraft(cableId) {
        if (!confirm('Are you sure you want to clear the terminal draft? This action cannot be undone.')) {
            return;
        }
        
        var projectId = {{ current_project.id }};
        
        fetch('/clear_terminal_draft', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({
                project_id: projectId,
                cable_id: cableId
            })
        })
        .then(response => response.json())
        .then(result => {
            if (result.success) {
                // Clear local unsaved data
                if (unsavedTerminalData[cableId]) {
                    delete unsavedTerminalData[cableId];
                }
                showCompactNotification('✓ Terminal draft cleared successfully', 'success', 2000);
            } else {
                showCompactNotification('✕ ' + (result.message || 'Failed to clear draft.'), 'error', 3000);
            }
        })
        .catch(error => {
            console.error('Error clearing terminal draft:', error);
            showCompactNotification('✕ Error clearing terminal draft', 'error', 3000);
        });
    }

    // ============================================================================
    // 📋 DOM CONTENT LOADED - MODAL INITIALIZATION
    // ============================================================================
    document.addEventListener('DOMContentLoaded', function() {
        configModal = new bootstrap.Modal(document.getElementById('cableConfigModal'));
        cableModal = new bootstrap.Modal(document.getElementById('addCableModal'));
        successModal = new bootstrap.Modal(document.getElementById('successFlashModal'));
        errorModal = new bootstrap.Modal(document.getElementById('errorModal'));
        terminalModal = new bootstrap.Modal(document.getElementById('terminalConfigModal'));
        headerModal = new bootstrap.Modal(document.getElementById('headerConfigModal'));
        groupModal = new bootstrap.Modal(document.getElementById('groupConfigModal'));
        chokeModal = new bootstrap.Modal(document.getElementById('chokeConfigModal'));
        resistorModal = new bootstrap.Modal(document.getElementById('resistorConfigModal'));
        
        // ============================================================================
        // 📋 ADD ALL DRAFT BUTTON EVENT LISTENERS
        // ============================================================================
        
        // Add event listener for Save Draft button
        document.getElementById('saveDraftBtn')?.addEventListener('click', function(e) {
            e.preventDefault();
            saveCableConfigDraft();
        });
        
        // Add event listener for Save Cable Table Draft button
        document.getElementById('saveCableTableDraftBtn')?.addEventListener('click', function(e) {
            e.preventDefault();
            saveCableTableDraft();
        });
        
        // Add event listener for Save Terminal Draft button
        document.getElementById('saveTerminalDraftBtn')?.addEventListener('click', function(e) {
            e.preventDefault();
            saveTerminalDraft();
        });
        
        // Add event listener for Save Terminal Header Draft button
        document.getElementById('saveTerminalHeaderDraftBtn')?.addEventListener('click', function(e) {
            e.preventDefault();
            saveTerminalHeaderDraft();
        });
        
        // Add event listener for Save Group Table Draft button
        document.getElementById('saveGroupTableDraftBtn')?.addEventListener('click', function(e) {
            e.preventDefault();
            saveGroupTableDraft();
        });
        
        // Add event listener for Save Choke Table Draft button
        document.getElementById('saveChokeTableDraftBtn')?.addEventListener('click', function(e) {
            e.preventDefault();
            saveChokeTableDraft();
        });
        
        // Add event listener for Clear Draft button
        document.getElementById('clearDraftBtn')?.addEventListener('click', function(e) {
            e.preventDefault();
            if (currentJunctionConfig && currentJunctionConfig.junctionName) {
                clearCableConfigDraft(currentJunctionConfig.junctionName);
            }
        });
        
        setTimeout(function() {
            for (var i = 0; i < requestedJunctions; i++) {
                validateRow(i);
            }
        }, 100);
        
        var previousToConfigBtn = document.getElementById('previousToConfigBtn');
        if (previousToConfigBtn) {
            previousToConfigBtn.addEventListener('click', function() {
                // Close cable modal first
                cableModal.hide();
                
                // Store the current configuration before navigating back
                var currentConfig = [...currentRowsConfig];
                
                // Clear the current cables array to prevent duplication
                currentCables = [];
                
                if (currentJunctionConfig) {
                    // Give a small delay to ensure modal is hidden before showing the next one
                    setTimeout(() => {
                        openCableConfigModal(
                            currentJunctionConfig.junctionIndex,
                            currentJunctionConfig.junctionName,
                            currentJunctionConfig.junctionSize,
                            currentJunctionConfig.rowCount,
                            currentConfig  // Pass the saved configuration
                        );
                    }, 300);
                }
            });
        }
        
        document.querySelectorAll('.config-cables-btn').forEach(button => {
            button.addEventListener('click', function() {
                currentJunctionIndex = parseInt(this.dataset.junctionIndex);
                var junctionRow = this.closest('tr').querySelector('.junction-row-input').value;
                var junctionName = this.closest('tr').querySelector('.junction-name-input').value;
                var junctionSize = this.closest('tr').querySelector('.junction-size-select').value;
               
                currentJunctionConfig = {
                    junctionIndex: currentJunctionIndex,
                    junctionName: junctionName,
                    junctionSize: junctionSize,
                    rowCount: parseInt(junctionRow)
                };
               
                openCableConfigModal(currentJunctionIndex, junctionName, junctionSize, parseInt(junctionRow));
            });
        });
        
        document.querySelectorAll('.edit-cables-btn').forEach(button => {
            button.addEventListener('click', function(e) {
                e.preventDefault();
                currentJunctionIndex = parseInt(this.dataset.junctionIndex);
                var junctionRow = document.querySelector(`input[name="junctions[${currentJunctionIndex}][junction_row]"]`).value;
                var junctionName = document.querySelector(`input[name="junctions[${currentJunctionIndex}][junction_name]"]`).value;
                var junctionSize = document.querySelector(`select[name="junctions[${currentJunctionIndex}][junction_size]"]`).value;
               
                currentJunctionConfig = {
                    junctionIndex: currentJunctionIndex,
                    junctionName: junctionName,
                    junctionSize: junctionSize,
                    rowCount: parseInt(junctionRow)
                };
               
                openCableConfigModal(currentJunctionIndex, junctionName, junctionSize, parseInt(junctionRow));
            });
        });
        
        document.querySelectorAll('.junction-name-input, .junction-size-select, .junction-row-input').forEach(input => {
            input.addEventListener('input', function() {
                var rowIndex = parseInt(this.name.match(/junctions\[(\d+)\]/)[1]);
                validateRow(rowIndex);
            });
            input.addEventListener('change', function() {
                var rowIndex = parseInt(this.name.match(/junctions\[(\d+)\]/)[1]);
                validateRow(rowIndex);
            });
            input.addEventListener('blur', function() {
                var rowIndex = parseInt(this.name.match(/junctions\[(\d+)\]/)[1]);
                validateRow(rowIndex);
            });
        });
        
        document.querySelectorAll('.junction-size-select').forEach(select => {
            select.addEventListener('change', function() {
                var rowIndex = parseInt(this.name.match(/junctions\[(\d+)\]/)[1]);
                validateRow(rowIndex);
            });
        });
        
        document.getElementById('addMoreJunctionGridForm')?.addEventListener('input', function() {
            for (var i = 0; i < requestedJunctions; i++) {
                validateRow(i);
            }
        });
        
        // ============================================================================
        // 📋 UPDATED: CABLE CONFIG FORM SUBMISSION - NOW SAVES TO DATABASE
        // ============================================================================
        document.getElementById('cableConfigForm').addEventListener('submit', function(e) {
            e.preventDefault();
            
            // CLEAR EXISTING CONFIG BEFORE PROCESSING NEW ONE
            currentRowsConfig = [];
            var configRows = [];
            var isValid = true;
            
            document.querySelectorAll('#configTableBody tr').forEach((row, index) => {
                var rowName = row.querySelector('.config-row-name').value.trim();
                var cableType = row.querySelector('.config-cable-type').value;
                var cableCount = parseInt(row.querySelector('.config-cable-count').value);
                
                if (!rowName || isNaN(cableCount) || cableCount < 1) {
                    isValid = false;
                    return;
                }
                
                // Store in currentRowsConfig
                currentRowsConfig.push({
                    rowName: rowName,
                    cableType: cableType,
                    cableCount: cableCount
                });
                
                // Prepare data for database
                configRows.push({
                    row_number: index + 1,
                    location_row_name: rowName,
                    cable_type: cableType,
                    number_of_cables: cableCount
                });
            });
            
            if (!isValid) {
                showCompactNotification('✕ Please fill all fields correctly', 'error', 3000);
                return;
            }
            
            // Show loading state
            var submitBtn = document.querySelector('#cableConfigForm button[type="submit"]');
            var originalText = submitBtn.innerHTML;
            submitBtn.disabled = true;
            submitBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving...';
            
            // Save configuration to database before proceeding
            saveCableConfigurationToDatabase(configRows, currentJunctionConfig.junctionName)
                .then(result => {
                    submitBtn.disabled = false;
                    submitBtn.innerHTML = originalText;
                    
                    // Clear any draft data since we're saving the final configuration
                    clearCableConfigDraft(currentJunctionConfig.junctionName);
                    
                    configModal.hide();
                    openCableTableModal(currentJunctionIndex);
                })
                .catch(error => {
                    submitBtn.disabled = false;
                    submitBtn.innerHTML = originalText;
                    // Error is already handled in the function
                });
        });
        
        document.getElementById('saveAllCablesBtn').addEventListener('click', saveAllCables);
        
        // REMOVED: saveAllTerminalsBtn event listener
        
        document.getElementById('configureHeadersBtn').addEventListener('click', function() {
            saveTerminalDataToUnsaved();
            saveTerminalsBeforeHeaders();
        });
        
        document.getElementById('headerPreviousBtn')?.addEventListener('click', function() {
            var cable = currentCables[currentCableIndex];
            var cableKey = cable.cable_id;
            unsavedHeaderData[cableKey] = currentHeaders;
           
            headerModal.hide();
            setTimeout(() => {
                terminalModal.show();
            }, 300);
        });
        
        // UPDATED: Configure Choke button event listener
        document.getElementById('configureChokeBtn').addEventListener('click', function() {
            // First save group data to unsaved storage
            var cable = currentCables[currentCableIndex];
            var cableKey = cable.cable_id;
            unsavedGroupData[cableKey] = currentGroups;
           
            // Save draft before opening choke config
            saveGroupTableDraftBeforeChoke();
        });
        
        document.getElementById('configureResistorBtn').addEventListener('click', function() {
            openResistorConfigModal();
        });
        
        document.getElementById('addHeaderRowBtn').addEventListener('click', function() {
            var cable = currentCables[currentCableIndex];
            var terminalInfo = getTerminalInfoFromCable(cable);
            addHeaderRow(terminalInfo.terminalStart, terminalInfo.terminalEnd);
        });
        
        document.getElementById('addGroupRowBtn').addEventListener('click', function() {
            addGroupRow();
        });
        
        document.getElementById('addChokeRowBtn').addEventListener('click', function() {
            addChokeRow();
        });
        
        document.getElementById('addResistorRowBtn').addEventListener('click', function() {
            addResistorRow();
        });
        
        // REMOVED: saveAllHeadersBtn event listener
        // document.getElementById('saveAllHeadersBtn').addEventListener('click', saveAllHeaders);
        
        // REMOVED: saveAllGroupsBtn event listener
        // document.getElementById('saveAllGroupsBtn').addEventListener('click', saveAllGroups);
        
        // REMOVED: saveAllChokesBtn event listener
        // document.getElementById('saveAllChokesBtn').addEventListener('click', saveAllChokes);
        
        document.getElementById('saveAllResistorsBtn').addEventListener('click', saveAllResistors);
        
        document.getElementById('terminalPreviousBtn')?.addEventListener('click', function() {
            saveTerminalDataToUnsaved();
            terminalModal.hide();
            setTimeout(() => {
                cableModal.show();
            }, 300);
        });
        
        document.getElementById('groupPreviousBtn').addEventListener('click', function() {
            var cable = currentCables[currentCableIndex];
            var cableKey = cable.cable_id;
            unsavedGroupData[cableKey] = currentGroups;
           
            groupModal.hide();
            terminalModal.show();
        });
        
        // ============================================================================
        // 📋 UPDATED: Configure Groups Button Functionality
        // ============================================================================
        document.getElementById('configureGroupsBtn').addEventListener('click', function() {
            // First save header data to unsaved storage
            var cable = currentCables[currentCableIndex];
            var cableKey = cable.cable_id;
            unsavedHeaderData[cableKey] = currentHeaders;
            
            // Save draft before opening groups
            saveTerminalHeaderDraftBeforeGroups();
        });
        
        document.getElementById('resistorPreviousBtn').addEventListener('click', function() {
            var cable = currentCables[currentCableIndex];
            var cableKey = cable.cable_id;
            unsavedResistorData[cableKey] = currentResistors;
           
            resistorModal.hide();
            terminalModal.show();
        });
        
        document.getElementById('chokePreviousBtn')?.addEventListener('click', function() {
            var cable = currentCables[currentCableIndex];
            var cableKey = cable.cable_id;
            unsavedChokeData[cableKey] = currentChokes;
           
            chokeModal.hide();
            groupModal.show();
        });
        
        document.getElementById('errorModal')?.addEventListener('hidden.bs.modal', function() {
            setTimeout(function() {
                forceRevalidateAllRows();
            }, 100);
        });
    });



    function cleanupCableConfigState() {
        // Clear the cable tables container
        var container = document.getElementById('cableTablesContainer');
        if (container) {
            container.innerHTML = '';
        }
        
        // Clear cable arrays
        currentCables = [];
        
        // Clear unsaved data
        unsavedTerminalData = {};
        unsavedHeaderData = {};
        unsavedGroupData = {};
        unsavedChokeData = {};
        unsavedResistorData = {};
        
        // Reset indices
        currentCableIndex = null;
        
        // Clear saved cable IDs for this session
        savedCableIds.clear();
    }

    // ============================================================================
    // 📋 VALIDATION & UTILITY FUNCTIONS
    // ============================================================================
    function validateRow(rowIndex) {
        var junctionNameInput = document.querySelector(`input[name="junctions[${rowIndex}][junction_name]"]`);
        var junctionSizeSelect = document.querySelector(`select[name="junctions[${rowIndex}][junction_size]"]`);
        var junctionRowInput = document.querySelector(`input[name="junctions[${rowIndex}][junction_row]"]`);
        var configButton = document.querySelector(`.config-cables-btn[data-junction-index="${rowIndex}"]`);
        
        if (!junctionNameInput || !junctionSizeSelect || !junctionRowInput || !configButton) {
            return false;
        }
        
        var junctionName = junctionNameInput.value;
        var junctionSize = junctionSizeSelect.value;
        var junctionRow = junctionRowInput.value;
        var isValid = junctionName && junctionName.trim() !== '' &&
                      ['Full', 'Half'].includes(junctionSize) &&
                      junctionRow && !isNaN(junctionRow) && parseInt(junctionRow) > 0;
        
        configButton.disabled = !isValid;
       
        if (isValid) {
            configButton.classList.remove('btn-secondary');
            configButton.classList.add('btn-primary');
        } else {
            configButton.classList.remove('btn-primary');
            configButton.classList.add('btn-secondary');
        }
       
        return isValid;
    }
    
    function hasUnsavedCables() {
        var visibleConfigButtons = document.querySelectorAll('.config-cables-btn:not([style*="display: none"])');
        return visibleConfigButtons.length > 0;
    }
    
    function hasOpenCableModalWithUnsavedCables() {
        if (cableModal && cableModal._element && cableModal._element.style.display !== 'none') {
            var unsavedCables = currentCables.filter(cable => !savedCableIds.has(cable.cable_id));
            return unsavedCables.length > 0;
        }
        return false;
    }
    
    function areAllFormFieldsValid() {
        for (var i = 0; i < requestedJunctions; i++) {
            var junctionNameInput = document.querySelector(`input[name="junctions[${i}][junction_name]"]`);
            var junctionSizeSelect = document.querySelector(`select[name="junctions[${i}][junction_size]"]`);
            var junctionRowInput = document.querySelector(`input[name="junctions[${i}][junction_row]"]`);
            
            if (!junctionNameInput || !junctionSizeSelect || !junctionRowInput) {
                continue;
            }
            
            var junctionName = junctionNameInput.value;
            var junctionSize = junctionSizeSelect.value;
            var junctionRow = junctionRowInput.value;
            var isValid = junctionName && junctionName.trim() !== '' &&
                          ['Full', 'Half'].includes(junctionSize) &&
                          junctionRow && !isNaN(junctionRow) && parseInt(junctionRow) > 0;
            
            if ((junctionName || junctionSize || junctionRow) && !isValid) {
                return false;
            }
        }
        return true;
    }
    
    function forceRevalidateAllRows() {
        for (var i = 0; i < requestedJunctions; i++) {
            validateRow(i);
        }
    }

    // ============================================================================
    // 📋 CABLE UNSAVED DATA HELPER FUNCTIONS
    // ============================================================================
    function saveUnsavedCableData() {
        if (currentJunctionIndex === null) return;
        var junctionKey = currentJunctionIndex;
        unsavedCableData[junctionKey] = {
            cables: JSON.parse(JSON.stringify(currentCables)),
            rowsConfig: JSON.parse(JSON.stringify(currentRowsConfig))
        };
    }
    
    function loadUnsavedCableData() {
        if (currentJunctionIndex === null) return null;
        var junctionKey = currentJunctionIndex;
        if (unsavedCableData[junctionKey]) {
            return unsavedCableData[junctionKey];
        }
        return null;
    }
    
    // ============================================================================
    // 📋 CABLE CONFIGURATION MODAL FUNCTIONS
    // ============================================================================
    // ============================================================================
    // 📋 CABLE CONFIGURATION MODAL FUNCTIONS
    // ============================================================================
    function openCableConfigModal(junctionIndex, junctionName, junctionSize, rowCount, existingRowsConfig = null) {
        // First, try to get the junction ID
        var junctionIdInput = document.querySelector(`input[name="junctions[${junctionIndex}][junction_id]"]`);
        var junctionId = junctionIdInput ? junctionIdInput.value : null;
        
        // Check if junctionName already has (F) or (H) suffix
        // If not, add it based on junctionSize
        var displayJunctionName = junctionName;
        
        // Remove any existing (F) or (H) suffix first to check
        var cleanName = junctionName.replace(/\s*\([FH]\)$/, '');
        
        // Only add suffix if it's not already there
        if (!junctionName.endsWith(' (F)') && !junctionName.endsWith(' (H)')) {
            if (junctionSize === 'Full') {
                displayJunctionName = cleanName + ' (F)';
            } else if (junctionSize === 'Half') {
                displayJunctionName = cleanName + ' (H)';
            }
        }
        
        currentJunctionConfig = {
            junctionIndex: junctionIndex,
            junctionName: displayJunctionName,  // Use the properly formatted name
            junctionId: junctionId,
            junctionSize: junctionSize,
            rowCount: rowCount
        };
        
        var tableBody = document.getElementById('configTableBody');
        tableBody.innerHTML = '';
        
        // RESET CURRENT ROWS CONFIG - This is the key fix
        currentRowsConfig = [];
        
        // Reset draft mode
        isDraftMode = false;
        currentDraftData = null;
        
        // Create table rows based on rowCount
        for (var i = 1; i <= rowCount; i++) {
            var row = document.createElement('tr');
            
            var rowName = String.fromCharCode(64 + i);
            // Convert old cable_box to relay_box if needed
            var cableType = existingRowsConfig && existingRowsConfig[i-1] ? 
                        (existingRowsConfig[i-1].cableType === 'cable_box' ? 'relay_box' : existingRowsConfig[i-1].cableType) : 
                        'cable';
            var cableCount = existingRowsConfig && existingRowsConfig[i-1] ? existingRowsConfig[i-1].cableCount : 1;
              
            row.innerHTML = `
                <td><strong>${i}</strong></td>
                <td>
                    <input type="text" class="form-control config-row-name" value="${rowName}" placeholder="Enter row name (e.g., A, B, C...)" required>
                </td>
                <td>
                    <select class="form-select config-cable-type" required>
                        <option value="cable" ${cableType === 'cable' ? 'selected' : ''}>Cable</option>
                        <option value="relay_box" ${cableType === 'relay_box' ? 'selected' : ''}>Relay Box</option>
                    </select>
                </td>
                <td>
                    <input type="number" class="form-control config-cable-count" value="${cableCount}" min="1" required>
                </td>
            `;
            tableBody.appendChild(row);
        }
        
        // Check for existing draft data
        loadCableConfigDraft(displayJunctionName);
        
        // Override with existingRowsConfig if provided (from unsaved data)
        if (existingRowsConfig && existingRowsConfig.length > 0) {
            existingRowsConfig.forEach((rowConfig, index) => {
                if (index < rowCount) {
                    var row = document.querySelectorAll('#configTableBody tr')[index];
                    if (row) {
                        row.querySelector('.config-row-name').value = rowConfig.rowName;
                        // Convert cable_box to relay_box for dropdown
                        var cableTypeValue = rowConfig.cableType === 'cable_box' ? 'relay_box' : rowConfig.cableType;
                        row.querySelector('.config-cable-type').value = cableTypeValue;
                        row.querySelector('.config-cable-count').value = rowConfig.cableCount;
                    }
                }
            });
        }
        
        if (displayJunctionName) {
            document.getElementById('cableConfigModalLabel').innerHTML =
                `<i class="bi bi-gear"></i> Cables Configuration Setup - ${displayJunctionName}`;
        }
        
        const modalOpenedEvent = new CustomEvent('cableModalOpened', {
            detail: { junctionIndex: junctionIndex }
        });
        document.dispatchEvent(modalOpenedEvent);
        configModal.show();
    }
    
    function updateCableName(index) {
        var cable = currentCables[index];
        if (!cable) return;
       
        if (cable.cable_type === 'relay_box') {
            return;
        }
       
        var row = cable.row;
        var startNo = parseInt(cable.start_no) || 1;
        var terminal = parseInt(cable.terminal) || 12;
        var endNo = startNo + terminal - 1;
        var cableName = row + ' T' + startNo + '-' + endNo;
       
        cable.cable_name = cableName;
       
        var cableNameInput = document.querySelector(`.cable-name-input[data-index="${index}"]`);
        if (cableNameInput) {
            cableNameInput.value = cableName;
        }
    }
    
    function checkCableExists(cableId, callback) {
        fetch(`/check_cable_exists?cable_id=${cableId}`)
            .then(response => response.json())
            .then(data => {
                callback(data.exists || false);
            })
            .catch(error => {
                console.error('Error checking cable:', error);
                callback(false);
            });
    }
    
    function saveCableAndOpenTerminals(cable, terminalCount, button) {
        button.disabled = true;
        button.innerHTML = '<span class="spinner-border spinner-border-sm"></span>';
        fetch('/add_cable_ajax', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(cable)
        })
        .then(response => {
            if (!response.ok) {
                return response.text().then(text => {
                    throw new Error(`Error: ${response.status} - ${text}`);
                });
            }
            return response.json();
        })
        .then(result => {
            button.disabled = false;
            button.innerHTML = '<i class="bi bi-gear"></i> Config Term';
            if (result.success) {
                savedCableIds.add(cable.cable_id);
                openTerminalModal(cable, terminalCount);
            } else {
                showCompactNotification('✕ ' + (result.message || 'Failed to save cable.'), 'error', 3000);
            }
        })
        .catch(error => {
            button.disabled = false;
            button.innerHTML = '<i class="bi bi-gear"></i> Config Term';
            showCompactNotification('✕ ' + (error.message || 'An error occurred while saving the cable.'), 'error', 3000);
        });
    }

    // ============================================================================
    // 📋 TERMINAL MODAL & SAVE FUNCTIONS - UPDATED TO UPDATE CABLE PROPERTIES
    // ============================================================================
    function openTerminalModal(cable, terminalCount) {
        document.getElementById('terminalConfigModalLabel').innerHTML =
            `<i class="bi bi-gear"></i> Configure Terminals for ${cable.cable_type === 'relay_box' ? 'Relay Box' : 'Cable'} - ${cable.junction_name}`;
    
        var cableInfoElement = document.getElementById('terminalCableInfo');
        if (cableInfoElement) {
            cableInfoElement.innerHTML = `
                <div class="alert alert-info">
                    <h6 class="alert-heading">Configure all terminals for this ${cable.cable_type === 'relay_box' ? 'relay box' : 'cable'} in the table below.</h6>
                    <hr>
                    <div class="row">
                        <div class="col-md-3"><strong>Location:</strong> ${cable.junction_name}</div>
                        <div class="col-md-3"><strong>Row:</strong> ${cable.row}</div>
                        <div class="col-md-3"><strong>Cable Position:</strong> ${cable.position}</div>
                        <div class="col-md-3"><strong>Location Box:</strong> ${cable.junction_box}</div>
                    </div>
                    <div class="row mt-2">
                        <div class="col-md-3"><strong>Cable ID:</strong> ${cable.cable_id}</div>
                        <div class="col-md-3"><strong>Cable Name:</strong> ${cable.cable_name}</div>
                        <div class="col-md-3"><strong>Start No:</strong> ${cable.start_no}</div>
                        <div class="col-md-3"><strong>Terminals:</strong> ${terminalCount}</div>
                    </div>
                    <div class="row mt-2">
                        <div class="col-md-3"><strong>Type:</strong> ${cable.cable_type === 'relay_box' ? 'Relay Box' : 'Cable'}</div>
                    </div>
                </div>
            `;
        }
        var tableBody = document.getElementById('terminalTableBody');
        tableBody.innerHTML = '';
    
        var startNo = parseInt(cable.start_no) || 1;
        var cableKey = cable.cable_id;
        var hasUnsavedData = unsavedTerminalData[cableKey];
        
        // NEW: Check if cable has draft terminal data
        var hasDraftData = false;
        
        // Load draft data if exists
        loadTerminalDraft(cable.cable_id);
        
        // Check if we have draft data after loading
        if (unsavedTerminalData[cableKey] && unsavedTerminalData[cableKey].length > 0) {
            hasDraftData = true;
            hasUnsavedData = unsavedTerminalData[cableKey];
        }
    
        for (var i = 1; i <= terminalCount; i++) {
            var terminalName = startNo + (i - 1);
            var unsavedRow = hasUnsavedData && hasUnsavedData[i-1] ? hasUnsavedData[i-1] : null;
        
            var inputLeft = unsavedRow ? unsavedRow.input_left : "";
            var inputRight = unsavedRow ? unsavedRow.input_right : "";
            var symbol = unsavedRow ? unsavedRow.symbol : "ara/wago";
            var spare = unsavedRow ? unsavedRow.spare : "N";
            var inputConnected = unsavedRow ? unsavedRow.input_connected : "Y";
            var outputConnected = unsavedRow ? unsavedRow.output_connected : "Y";
            var inputConnectedExtra = unsavedRow ? unsavedRow.input_connected_extra : "";
            var outputConnectedExtra = unsavedRow ? unsavedRow.output_connected_extra : "";
            var outputLeft = unsavedRow ? unsavedRow.output_left : "";
            var outputRight = unsavedRow ? unsavedRow.output_right : "";
        
            var row = document.createElement('tr');
            row.innerHTML = `
                <td><input type="text" class="form-control cable-id-input" value="${cable.cable_id}" readonly></td>
                <td><input type="text" class="form-control terminal-id-input" value="${i}" placeholder="Enter terminal id"></td>
                <td><input type="text" class="form-control terminal-name-input" value="${terminalName}" placeholder="Enter terminal number"></td>
                <td>
                    <select class="form-select symbol-input">
                        <option value="ara/wago" ${symbol === 'ara/wago' ? 'selected' : ''}>Ara/Wago</option>
                        <option value="single_fuse" ${symbol === 'single_fuse' ? 'selected' : ''}>Single Fuse</option>
                        <option value="dual_fuse" ${symbol === 'dual_fuse' ? 'selected' : ''}>Dual Fuse</option>
                    </select>
                </td>
                <td><input type="text" class="form-control input-left-input" value="${inputLeft}" placeholder="Enter input left"></td>
                <td><input type="text" class="form-control input-right-input" value="${inputRight}" placeholder="Enter input right"></td>
                <td>
                    <div class="d-flex align-items-center justify-content-center">
                        <input type="text" class="form-control spare-input text-center" value="${spare}" style="max-width: 60px;">
                    </div>
                </td>
                <td>
                    <div class="d-flex align-items-center justify-content-center">
                        <input type="text" class="form-control input-connected-input text-center" value="${inputConnected}" style="max-width: 60px;">
                    </div>
                </td>
                <td>
                    <div class="d-flex align-items-center justify-content-center">
                        <input type="text" class="form-control output-connected-input text-center" value="${outputConnected}" style="max-width: 60px;">
                    </div>
                </td>
                <td><input type="text" class="form-control input-connected-extra-input" value="${inputConnectedExtra}" placeholder="Enter input connected extra"></td>
                <td><input type="text" class="form-control output-connected-extra-input" value="${outputConnectedExtra}" placeholder="Enter output connected extra"></td>
                <td><input type="text" class="form-control output-left-input" value="${outputLeft}" placeholder="Enter output left"></td>
                <td><input type="text" class="form-control output-right-input" value="${outputRight}" placeholder="Enter output right"></td>
            `;
            tableBody.appendChild(row);
        }
    
        addTerminalYNEventListeners();
        
        // NEW: Update status based on whether we have draft data
        setTimeout(() => {
            if (hasDraftData) {
                var cableIndex = currentCableIndex;
                if (cable && cable.cable_type === 'relay_box') {
                    updateCableStatusImmediately(cableIndex);
                } else if (cable) {
                    updateCableStatusForSingleCable(cable.cable_id, cableIndex);
                }
            }
        }, 500);
    
        if (terminalModal) {
            terminalModal.show();
        } else {
            terminalModal = new bootstrap.Modal(document.getElementById('terminalConfigModal'));
            terminalModal.show();
        }
    }
    
    // ============================================================================
    // 📋 UPDATED: SAVE TERMINAL DATA TO UNSAVED - ALSO UPDATE CABLE PROPERTIES
    // ============================================================================
    function saveTerminalDataToUnsaved() {
        var cable = currentCables[currentCableIndex];
        var cableId = cable.cable_id;
        var terminals = [];
        
        // Get current terminal count from the table
        var terminalCount = document.querySelectorAll('#terminalTableBody tr').length;
        
        document.querySelectorAll('#terminalTableBody tr').forEach((row, idx) => {
            terminals.push({
                cable_id: cableId,
                terminal_id: row.querySelector('.terminal-id-input').value,
                terminal_no: row.querySelector('.terminal-name-input').value,
                symbol: row.querySelector('.symbol-input').value,
                input_left: row.querySelector('.input-left-input').value,
                input_right: row.querySelector('.input-right-input').value,
                spare: row.querySelector('.spare-input').value,
                input_connected: row.querySelector('.input-connected-input').value,
                output_connected: row.querySelector('.output-connected-input').value,
                input_connected_extra: row.querySelector('.input-connected-extra-input').value,
                output_connected_extra: row.querySelector('.output-connected-extra-input').value,
                output_left: row.querySelector('.output-left-input').value,
                output_right: row.querySelector('.output-right-input').value
            });
        });
        
        // Update cable properties if we have terminals
        if (terminals.length > 0) {
            // Get start_no from the first terminal (if available)
            var firstTerminalNo = parseInt(terminals[0].terminal_no) || parseInt(cable.start_no) || 1;
            
            // Update cable object with current values
            cable.terminal = terminalCount;
            cable.start_no = firstTerminalNo;
            
            // Update the cable name if it's not a relay box
            if (cable.cable_type !== 'relay_box') {
                var endNo = firstTerminalNo + terminalCount - 1;
                cable.cable_name = cable.row + ' T' + firstTerminalNo + '-' + endNo;
                
                // Also update the input field in the cable table
                var cableNameInput = document.querySelector(`.cable-name-input[data-index="${currentCableIndex}"]`);
                if (cableNameInput) {
                    cableNameInput.value = cable.cable_name;
                }
            }
            
            console.log('Updated cable properties:', cable);
        }
        
        unsavedTerminalData[cableId] = terminals;
    }
    
    function addTerminalYNEventListeners() {
        document.querySelectorAll('.spare-input, .input-connected-input, .output-connected-input').forEach(input => {
            input.defaultValue = input.value;
           
            input.addEventListener('click', function() {
                this.select();
            });
           
            input.addEventListener('keydown', function(event) {
                const allowedKeys = ['Y', 'N', 'y', 'n', 'Backspace', 'Delete', 'Tab', 'ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'];
               
                if (allowedKeys.includes(event.key)) {
                    if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
                        this.value = this.value === 'Y' ? 'N' : 'Y';
                        this.defaultValue = this.value;
                        event.preventDefault();
                    }
                    else if (event.key === 'Y' || event.key === 'y') {
                        this.value = 'Y';
                        this.defaultValue = 'Y';
                        event.preventDefault();
                    }
                    else if (event.key === 'N' || event.key === 'n') {
                        this.value = 'N';
                        this.defaultValue = 'N';
                        event.preventDefault();
                    }
                } else {
                    event.preventDefault();
                }
            });
           
            input.addEventListener('input', function() {
                const value = this.value.toUpperCase();
                if (value === 'Y' || value === 'N') {
                    this.value = value;
                    this.defaultValue = value;
                } else if (value.length > 1) {
                    if (value.includes('Y')) {
                        this.value = 'Y';
                        this.defaultValue = 'Y';
                    } else if (value.includes('N')) {
                        this.value = 'N';
                        this.defaultValue = 'N';
                    } else {
                        this.value = this.defaultValue;
                    }
                }
            });
           
            input.addEventListener('blur', function() {
                const value = this.value.toUpperCase();
                if (value !== 'Y' && value !== 'N') {
                    this.value = this.defaultValue;
                } else {
                    this.value = value;
                    this.defaultValue = value;
                }
            });
        });
    }
    
    function saveTerminalsBeforeHeaders() {
        var cable = currentCables[currentCableIndex];
        if (cable && !savedCableIds.has(cable.cable_id)) {
            showCompactNotification('✕ Please save terminals before configuring headers.', 'error', 3000);
            return false;
        }
        openHeaderConfigModal();
        return true;
    }

    // ============================================================================
    // 📋 SAVE ALL FUNCTIONS WITH COMPACT NOTIFICATIONS - UPDATED
    // ============================================================================
    function saveAllCables() {
        if (isSavingAllCables) {
            return;
        }
       
        var saveAllBtn = document.getElementById('saveAllCablesBtn');
        isSavingAllCables = true;
        saveAllBtn.disabled = true;
        saveAllBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving All...';
        var cablesToSave = currentCables.filter(cable => !savedCableIds.has(cable.cable_id));
       
        if (cablesToSave.length === 0) {
            saveAllBtn.disabled = false;
            saveAllBtn.innerHTML = '<i class="bi bi-check-circle"></i> Save All Cables';
            isSavingAllCables = false;
           
            if (currentJunctionIndex !== null) {
                junctionCablesConfigured[currentJunctionIndex] = true;
               
                // 🎯 NEW: Dispatch event to notify main grid that cables are configured
                const cablesSavedEvent = new CustomEvent('cablesSaved', {
                    detail: {
                        junctionIndex: currentJunctionIndex,
                        junctionName: currentJunctionConfig?.junctionName || 'Unknown',
                        cableCount: currentCables.length
                    }
                });
                document.dispatchEvent(cablesSavedEvent);
            }
           
            cableModal.hide();
            showCompactNotification('✓ All cables saved successfully!', 'success', 2000);
            return;
        }
        var promises = cablesToSave.map(cable => {
            return fetch('/add_cable_ajax', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify(cable)
            })
            .then(response => {
                if (!response.ok) {
                    throw new Error(`Error: ${response.status}`);
                }
                return response.json();
            });
        });
        Promise.all(promises)
            .then(results => {
                saveAllBtn.disabled = false;
                saveAllBtn.innerHTML = '<i class="bi bi-check-circle"></i> Save All Cables';
                isSavingAllCables = false;
                var allSuccess = results.every(result => result.success);
               
                if (allSuccess) {
                    cablesToSave.forEach(cable => {
                        savedCableIds.add(cable.cable_id);
                    });
                   
                    if (currentJunctionIndex !== null) {
                        junctionCablesConfigured[currentJunctionIndex] = true;
                       
                        // 🎯 NEW: Dispatch event to notify main grid that cables are configured
                        const cablesSavedEvent = new CustomEvent('cablesSaved', {
                            detail: {
                                junctionIndex: currentJunctionIndex,
                                junctionName: currentJunctionConfig?.junctionName || 'Unknown',
                                cableCount: cablesToSave.length
                            }
                        });
                        document.dispatchEvent(cablesSavedEvent);
                    }
                   
                    cableModal.hide();
                    showCompactNotification('✓ All cables saved successfully!', 'success', 2000);
                } else {
                    showCompactNotification('✕ Some cables failed to save. Please check and try again.', 'error', 3000);
                    console.log(results);
                }
            })
            .catch(error => {
                saveAllBtn.disabled = false;
                saveAllBtn.innerHTML = '<i class="bi bi-check-circle"></i> Save All Cables';
                isSavingAllCables = false;
                showCompactNotification('✕ ' + (error.message || 'An error occurred while saving cables.'), 'error', 3000);
                console.log(error);
            });
    }
    
    function saveAllResistors() {
        var saveBtn = document.getElementById('saveAllResistorsBtn');
        saveBtn.disabled = true;
        saveBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving...';
        var promises = currentResistors.map(resistor => {
            return fetch('/add_resistor_ajax', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify(resistor)
            }).then(response => {
                if (!response.ok) {
                    return response.text().then(text => {
                        throw new Error(`Failed to save resistor: ${response.status} - ${text}`);
                    });
                }
                return response.json();
            });
        });
        Promise.all(promises)
            .then(results => {
                saveBtn.disabled = false;
                saveBtn.innerHTML = '<i class="bi bi-check-circle"></i> Save All Resistors';
                if (results.every(r => r.success)) {
                    var cable = currentCables[currentCableIndex];
                    delete unsavedResistorData[cable.cable_id];
                    resistorModal.hide();
                    showCompactNotification(' All resistors saved successfully!', 'success', 2000);
                } else {
                    showCompactNotification('✕ Some resistors failed to save.', 'error', 3000);
                }
            })
            .catch(error => {
                saveBtn.disabled = false;
                saveBtn.innerHTML = '<i class="bi bi-check-circle"></i> Save All Resistors';
                showCompactNotification('✕ ' + (error.message || 'An error occurred while saving resistors.'), 'error', 3000);
            });
    }

    // ============================================================================
    // 📋 CABLE STATUS FUNCTIONS
    // ============================================================================
    function updateCableStatus() {
        var configButton = document.querySelector(`.config-cables-btn[data-junction-index="${currentJunctionIndex}"]`);
        var cableStatus = document.querySelector(`.cable-status[data-junction-index="${currentJunctionIndex}"]`);
        if (configButton && cableStatus) {
            configButton.style.display = 'none';
            cableStatus.style.display = 'inline';
        }
    }
    
    function updateCableStatusImmediately(index) {
        var statusElement = document.getElementById(`status-${index}`);
        if (statusElement) {
            var cable = currentCables[index];
            if (cable.cable_type === 'relay_box') {
                statusElement.innerHTML = '<i class="bi bi-check-circle-fill text-success" title="No terminals required"></i>';
            } else {
                // Check if we have draft terminal data for this cable
                var cableKey = cable.cable_id;
                if (unsavedTerminalData[cableKey] && unsavedTerminalData[cableKey].length > 0) {
                    // We have draft terminal data, show green
                    statusElement.innerHTML = '<i class="bi bi-check-circle-fill text-success" title="Terminals configured (draft)"></i>';
                } else {
                    // Check database for saved terminals
                    fetch(`/check_terminals_for_cable?cable_id=${cable.cable_id}`)
                        .then(response => {
                            if (!response.ok) {
                                return;
                            }
                            return response.json();
                        })
                        .then(data => {
                            if (!data) return;
                            
                            if (statusElement) {
                                if (data.exists) {
                                    statusElement.innerHTML = '<i class="bi bi-check-circle-fill text-success" title="Terminals configured"></i>';
                                } else {
                                    statusElement.innerHTML = '<i class="bi bi-x-circle-fill text-danger" title="Terminals not configured"></i>';
                                }
                            }
                        })
                        .catch(error => {
                            console.error('Error checking terminal status for cable:', cable.cable_id, error);
                            statusElement.innerHTML = '<i class="bi bi-x-circle-fill text-danger" title="Terminals not configured"></i>';
                        });
                }
            }
        }
    }
    
    function updateCableStatusForSingleCable(cableId, index) {
        var cable = currentCables.find(c => parseInt(c.cable_id) === parseInt(cableId));
        var statusElement = document.getElementById(`status-${index}`);
        if (statusElement) {
            if (cable && cable.cable_type === 'relay_box') {
                statusElement.innerHTML = '<i class="bi bi-check-circle-fill text-success" title="No terminals required"></i>';
                return;
            }
        }
        fetch(`/check_terminals_for_cable?cable_id=${cableId}`)
            .then(response => {
                if (!response.ok) {
                    updateCableStatusImmediately(index);
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                return response.json();
            })
            .then(data => {
                var statusElement = document.getElementById(`status-${index}`);
                if (statusElement) {
                    if (data.exists) {
                        statusElement.innerHTML = '<i class="bi bi-check-circle-fill text-success" title="Terminals configured"></i>';
                    } else {
                        statusElement.innerHTML = '<i class="bi bi-x-circle-fill text-danger" title="Terminals not configured"></i>';
                    }
                }
            })
            .catch(error => {
                console.error('Error checking terminal status for cable:', cableId, error);
            });
    }
    
    function updateCableStatuses() {
        currentCables.forEach((cable, index) => {
            var statusElement = document.getElementById(`status-${index}`);
            if (!statusElement) return;
            if (cable.cable_type === 'relay_box') {
                statusElement.innerHTML = '<i class="bi bi-check-circle-fill text-success" title="No terminals required"></i>';
                return;
            }
            fetch(`/check_terminals_for_cable?cable_id=${cable.cable_id}`)
                .then(response => {
                    if (!response.ok) {
                        return;
                    }
                    return response.json();
                })
                .then(data => {
                    if (!data) return;
                   
                    var statusElement = document.getElementById(`status-${index}`);
                    if (statusElement) {
                        if (data.exists) {
                            statusElement.innerHTML = '<i class="bi bi-check-circle-fill text-success" title="Terminals configured"></i>';
                        } else {
                            statusElement.innerHTML = '<i class="bi bi-x-circle-fill text-danger" title="Terminals not configured"></i>';
                        }
                    }
                })
                .catch(error => {
                    console.error('Error checking terminal status for cable:', cable.cable_id, error);
                });
        });
    }

    // ============================================================================
    // 📋 HEADER CONFIGURATION FUNCTIONS - UPDATED TO USE CABLE PROPERTIES
    // ============================================================================
    function openHeaderConfigModal() {
        var cable = currentCables[currentCableIndex];
        document.getElementById('headerCableId').value = cable.cable_id;
        
        var tableBody = document.getElementById('headerTableBody');
        tableBody.innerHTML = '';
        
        var cableKey = cable.cable_id;
        if (unsavedHeaderData[cableKey]) {
            currentHeaders = unsavedHeaderData[cableKey];
            refreshHeaderTable();
            headerModal.show();
            return;
        } else {
            currentHeaders = [];
            
            // Get terminal info from cable properties
            var terminalInfo = getTerminalInfoFromCable(cable);
            var terminalStart = terminalInfo.terminalStart;
            var terminalEnd = terminalInfo.terminalEnd;
            
            // Debug log
            console.log('Opening header modal for cable:', cable);
            console.log('Terminal info from cable:', terminalInfo);
            console.log('Cable properties - start_no:', cable.start_no, 'terminal:', cable.terminal);
            console.log('Cable name:', cable.cable_name);
            
            tableBody.innerHTML = '<tr><td colspan="6" class="text-center"><div class="spinner-border spinner-border-sm" role="status"></div> Loading header data...</td></tr>';
            
            // First, load draft data if exists
            loadTerminalHeaderDraft(cable.cable_id);
            
            // Then check for existing final headers
            fetch(`/get_headers_for_cable?cable_id=${cable.cable_id}`)
                .then(response => response.json())
                .then(existingHeaders => {
                    tableBody.innerHTML = '';
                    
                    let nextHeaderId = 1;
                    if (existingHeaders && existingHeaders.length > 0) {
                        existingHeaders.forEach(header => {
                            // Only add if not already in currentHeaders (from draft)
                            const exists = currentHeaders.some(h => 
                                h.header_type === header.header_type && 
                                h.terminal_start === header.terminal_start && 
                                h.terminal_end === header.terminal_end
                            );
                            if (!exists) {
                                currentHeaders.push({
                                    cable_id: header.cable_id,
                                    header_type: header.header_type,
                                    terminal_start: header.terminal_start,
                                    terminal_end: header.terminal_end,
                                    input_output: header.input_output,
                                    text: header.text
                                });
                            }
                        });
                        
                        if (currentHeaders.length > 0) {
                            refreshHeaderTable();
                        } else {
                            // Use terminal info from cable properties
                            addHeaderRowWithTerminals(terminalStart, terminalEnd);
                        }
                    } else {
                        // Use terminal info from cable properties
                        addHeaderRowWithTerminals(terminalStart, terminalEnd);
                    }
                    
                    headerModal.show();
                })
                .catch(error => {
                    console.error('Error fetching existing headers:', error);
                    tableBody.innerHTML = '';
                    // Use terminal info from cable properties
                    addHeaderRowWithTerminals(terminalStart, terminalEnd);
                    headerModal.show();
                });
        }
    }
    
    // ============================================================================
    // 📋 UPDATED: ADD HEADER ROW WITH TERMINALS - IMPROVED TOOLTIP
    // ============================================================================
    function addHeaderRowWithTerminals(terminalStart, terminalEnd) {
        var tableBody = document.getElementById('headerTableBody');
        var cableId = document.getElementById('headerCableId').value;
        var cable = currentCables[currentCableIndex];
        
        var header = {
            cable_id: cableId,
            header_type: '',
            terminal_start: terminalStart,
            terminal_end: terminalEnd,
            input_output: '',
            text: ''
        };
        
        var rowIndex = currentHeaders.length;
        currentHeaders.push(header);
        
        var cableKey = cable.cable_id;
        unsavedHeaderData[cableKey] = currentHeaders;
        
        // Get the actual terminal range from cable properties
        var actualTerminalStart = parseInt(cable.start_no) || terminalStart;
        var actualTerminalEnd = actualTerminalStart + (parseInt(cable.terminal) || (terminalEnd - terminalStart + 1)) - 1;
        
        var row = document.createElement('tr');
        row.innerHTML = `
            <td>
                <select class="form-select header-type-input" data-index="${rowIndex}">
                    <option value="">Select Header Type</option>
                    <option value="WIREFROM">WIREFROM</option>
                    <option value="WIRETO">WIRETO</option>
                    <option value="RELAY">RELAY</option>
                    <option value="RELAY_BOX">RELAY BOX</option>
                    <option value="RELAY_CONTACT_BOX">RELAY CONTACT BOX</option>
                </select>
            </td>
            <td>
                <input type="number" class="form-control header-terminal-start-input" data-index="${rowIndex}" value="${actualTerminalStart}" placeholder="Start" min="1">
                <small class="form-text text-muted">Actual terminal range: ${actualTerminalStart}-${actualTerminalEnd}</small>
            </td>
            <td>
                <input type="number" class="form-control header-terminal-end-input" data-index="${rowIndex}" value="${actualTerminalEnd}" placeholder="End" min="1">
                <small class="form-text text-muted">Based on ${cable.terminal || '?'} terminals starting at ${actualTerminalStart}</small>
            </td>
            <td>
                <select class="form-select header-input-output-input" data-index="${rowIndex}">
                    <option value="">Select I/O</option>
                    <option value="Input">Input</option>
                    <option value="Output">Output</option>
                </select>
            </td>
            <td>
                <input type="text" class="form-control header-text-input" data-index="${rowIndex}" placeholder="Header text">
            </td>
            <td>
                <button type="button" class="btn btn-sm btn-danger remove-header-row-btn" data-index="${rowIndex}">
                    <i class="bi bi-trash"></i>
                </button>
            </td>
        `;
        tableBody.appendChild(row);
        
        addHeaderRowEventListeners(rowIndex);
    }
    
    // ============================================================================
    // 📋 UPDATED: ADD HEADER ROW - IMPROVED TOOLTIP
    // ============================================================================
    function addHeaderRow(terminalStart, terminalEnd) {
        var tableBody = document.getElementById('headerTableBody');
        var cable = currentCables[currentCableIndex];
        
        // If terminalStart and terminalEnd are not provided, get from cable properties
        if (terminalStart === undefined || terminalEnd === undefined) {
            var terminalInfo = getTerminalInfoFromCable(cable);
            terminalStart = terminalInfo.terminalStart;
            terminalEnd = terminalInfo.terminalEnd;
        }
        
        var header = {
            cable_id: cable.cable_id,
            header_type: '',
            terminal_start: terminalStart,
            terminal_end: terminalEnd,
            input_output: '',
            text: ''
        };
        
        var rowIndex = currentHeaders.length;
        currentHeaders.push(header);
        
        var cableKey = cable.cable_id;
        unsavedHeaderData[cableKey] = currentHeaders;
        
        // Get the actual terminal range from cable properties
        var actualTerminalStart = parseInt(cable.start_no) || terminalStart;
        var actualTerminalEnd = actualTerminalStart + (parseInt(cable.terminal) || (terminalEnd - terminalStart + 1)) - 1;
        
        var row = document.createElement('tr');
        row.innerHTML = `
            <td>
                <select class="form-select header-type-input" data-index="${rowIndex}">
                    <option value="">Select Header Type</option>
                    <option value="WIREFROM">WIREFROM</option>
                    <option value="WIRETO">WIRETO</option>
                    <option value="RELAY">RELAY</option>
                    <option value="RELAY_BOX">RELAY BOX</option>
                    <option value="RELAY_CONTACT_BOX">RELAY CONTACT BOX</option>
                </select>
            </td>
            <td>
                <input type="number" class="form-control header-terminal-start-input" data-index="${rowIndex}" value="${actualTerminalStart}" placeholder="Start" min="1">
                <small class="form-text text-muted">Actual: ${actualTerminalStart}-${actualTerminalEnd}</small>
            </td>
            <td>
                <input type="number" class="form-control header-terminal-end-input" data-index="${rowIndex}" value="${actualTerminalEnd}" placeholder="End" min="1">
                <small class="form-text text-muted">Cable: ${cable.cable_name} (${cable.terminal || '?'} terminals)</small>
            </td>
            <td>
                <select class="form-select header-input-output-input" data-index="${rowIndex}">
                    <option value="">Select I/O</option>
                    <option value="Input">Input</option>
                    <option value="Output">Output</option>
                </select>
            </td>
            <td>
                <input type="text" class="form-control header-text-input" data-index="${rowIndex}" placeholder="Header text">
            </td>
            <td>
                <button type="button" class="btn btn-sm btn-danger remove-header-row-btn" data-index="${rowIndex}">
                    <i class="bi bi-trash"></i>
                </button>
            </td>
        `;
        tableBody.appendChild(row);
        
        addHeaderRowEventListeners(rowIndex);
    }
    
    function addHeaderRowEventListeners(rowIndex) {
        var cable = currentCables[currentCableIndex];
        var cableKey = cable.cable_id;
        
        document.querySelector(`.header-type-input[data-index="${rowIndex}"]`).addEventListener('change', function() {
            currentHeaders[rowIndex].header_type = this.value;
            unsavedHeaderData[cableKey] = currentHeaders;
        });
        
        document.querySelector(`.header-terminal-start-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentHeaders[rowIndex].terminal_start = parseInt(this.value) || 1;
            unsavedHeaderData[cableKey] = currentHeaders;
        });
        
        document.querySelector(`.header-terminal-end-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentHeaders[rowIndex].terminal_end = parseInt(this.value) || 1;
            unsavedHeaderData[cableKey] = currentHeaders;
        });
        
        document.querySelector(`.header-input-output-input[data-index="${rowIndex}"]`).addEventListener('change', function() {
            currentHeaders[rowIndex].input_output = this.value;
            unsavedHeaderData[cableKey] = currentHeaders;
        });
        
        document.querySelector(`.header-text-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentHeaders[rowIndex].text = this.value;
            unsavedHeaderData[cableKey] = currentHeaders;
        });
        
        document.querySelector(`.remove-header-row-btn[data-index="${rowIndex}"]`).addEventListener('click', function() {
            removeHeaderRow(rowIndex);
        });
    }
    
    function removeHeaderRow(rowIndex) {
        var cable = currentCables[currentCableIndex];
        var cableKey = cable.cable_id;
        
        currentHeaders.splice(rowIndex, 1);
        unsavedHeaderData[cableKey] = currentHeaders;
        refreshHeaderTable();
    }
    
    // ============================================================================
    // 📋 UPDATED: REFRESH HEADER TABLE - IMPROVED TOOLTIP
    // ============================================================================
    function refreshHeaderTable() {
        var tableBody = document.getElementById('headerTableBody');
        tableBody.innerHTML = '';
        
        var cable = currentCables[currentCableIndex];
        
        currentHeaders.forEach((header, index) => {
            // Get actual terminal range from cable properties for tooltip
            var actualTerminalStart = parseInt(cable.start_no) || header.terminal_start || 1;
            var actualTerminalEnd = actualTerminalStart + (parseInt(cable.terminal) || (header.terminal_end - header.terminal_start + 1) || 12) - 1;
            
            var row = document.createElement('tr');
            row.innerHTML = `
                <td>
                    <select class="form-select header-type-input" data-index="${index}">
                        <option value="">Select Header Type</option>
                        <option value="WIREFROM" ${header.header_type === 'WIREFROM' ? 'selected' : ''}>WIREFROM</option>
                        <option value="WIRETO" ${header.header_type === 'WIRETO' ? 'selected' : ''}>WIRETO</option>
                        <option value="RELAY" ${header.header_type === 'RELAY' ? 'selected' : ''}>RELAY</option>
                        <option value="RELAY_BOX" ${header.header_type === 'RELAY_BOX' ? 'selected' : ''}>RELAY BOX</option>
                        <option value="RELAY_CONTACT_BOX" ${header.header_type === 'RELAY_CONTACT_BOX' ? 'selected' : ''}>RELAY CONTACT BOX</option>
                    </select>
                </td>
                <td>
                    <input type="number" class="form-control header-terminal-start-input" data-index="${index}" value="${actualTerminalStart || ''}" placeholder="Start" min="1">
                    <small class="form-text text-muted">Cable start: ${actualTerminalStart}</small>
                </td>
                <td>
                    <input type="number" class="form-control header-terminal-end-input" data-index="${index}" value="${actualTerminalEnd|| ''}" placeholder="End" min="1">
                    <small class="form-text text-muted">Cable end: ${actualTerminalEnd}</small>
                </td>
                <td>
                    <select class="form-select header-input-output-input" data-index="${index}">
                        <option value="">Select I/O</option>
                        <option value="Input" ${header.input_output === 'Input' ? 'selected' : ''}>Input</option>
                        <option value="Output" ${header.input_output === 'Output' ? 'selected' : ''}>Output</option>
                    </select>
                </td>
                <td>
                    <input type="text" class="form-control header-text-input" data-index="${index}" value="${header.text || ''}" placeholder="Header text">
                </td>
                <td>
                    <button type="button" class="btn btn-sm btn-danger remove-header-row-btn" data-index="${index}">
                        <i class="bi bi-trash"></i>
                    </button>
                </td>
            `;
            tableBody.appendChild(row);
            addHeaderRowEventListeners(index);
        });
    }

    // ============================================================================
    // 📋 GROUP CONFIGURATION FUNCTIONS
    // ============================================================================
    function openGroupConfigModal() {
        var cable = currentCables[currentCableIndex];
        document.getElementById('groupCableId').value = cable.cable_id;
       
        var tableBody = document.getElementById('groupTableBody');
        tableBody.innerHTML = '';
       
        var cableKey = cable.cable_id;
        if (unsavedGroupData[cableKey]) {
            currentGroups = unsavedGroupData[cableKey];
            refreshGroupTable();
            groupModal.show();
            return;
        } else {
            currentGroups = [];
           
            // First, load draft data if exists
            loadGroupTableDraft(cable.cable_id);
            
            // Then check for existing final groups
            fetch(`/get_groups_for_cable?cable_id=${cable.cable_id}`)
                .then(response => response.json())
                .then(existingGroups => {
                    let nextGroupId = 1;
                    if (existingGroups && existingGroups.length > 0) {
                        const maxId = Math.max(...existingGroups.map(group => parseInt(group.group_id.replace('GR', '')) || 0));
                        nextGroupId = maxId + 1;
                       
                        existingGroups.forEach(group => {
                            // Only add if not already in currentGroups (from draft)
                            const exists = currentGroups.some(g => g.group_id === group.group_id);
                            if (!exists) {
                                currentGroups.push({
                                    cable_id: group.cable_id,
                                    group_id: group.group_id,
                                    terminal_no: group.terminal_no || '',
                                    input_output: group.input_output || '',
                                    text: group.text || ''
                                });
                            }
                        });
                    }
                   
                    window.nextGroupId = nextGroupId;
                   
                    if (currentGroups.length === 0) {
                        addGroupRow();
                    } else {
                        refreshGroupTable();
                    }
                    groupModal.show();
                })
                .catch(error => {
                    console.error('Error fetching existing groups:', error);
                    window.nextGroupId = 1;
                    if (currentGroups.length === 0) {
                        addGroupRow();
                    }
                    groupModal.show();
                });
        }
    }
    
    function addGroupRow() {
        var cable = currentCables[currentCableIndex];
        var groupId = window.nextGroupId || 1;
       
        var group = {
            cable_id: cable.cable_id,
            group_id: groupId.toString(),
            terminal_no: '',
            input_output: '',
            text: ''
        };
       
        var rowIndex = currentGroups.length;
        currentGroups.push(group);
       
        var cableKey = cable.cable_id;
        unsavedGroupData[cableKey] = currentGroups;
       
        refreshGroupTable();
       
        window.nextGroupId = groupId + 1;
    }
    
    function refreshGroupTable() {
        var tableBody = document.getElementById('groupTableBody');
        tableBody.innerHTML = '';
       
        currentGroups.forEach((group, index) => {
            var row = document.createElement('tr');
            row.innerHTML = `
                <td>
                    <input type="text" class="form-control group-id-input" data-index="${index}" value="${group.group_id || ''}" readonly>
                </td>
                <td>
                    <input type="text" class="form-control group-terminal-no-input" data-index="${index}" value="${group.terminal_no || ''}" placeholder="Terminal numbers (e.g., 1,2,4,54)">
                </td>
                <td>
                    <select class="form-select group-input-output-input" data-index="${index}">
                        <option value="">Select I/O</option>
                        <option value="Input" ${group.input_output === 'Input' ? 'selected' : ''}>Input</option>
                        <option value="Output" ${group.input_output === 'Output' ? 'selected' : ''}>Output</option>
                    </select>
                </td>
                <td>
                    <input type="text" class="form-control group-text-input" data-index="${index}" value="${group.text || ''}" placeholder="Group text">
                </td>
                <td>
                    <button type="button" class="btn btn-sm btn-danger remove-group-row-btn" data-index="${index}">
                        <i class="bi bi-trash"></i>
                    </button>
                </td>
            `;
            tableBody.appendChild(row);
            addGroupRowEventListeners(index);
        });
    }
    
    function addGroupRowEventListeners(rowIndex) {
        var cable = currentCables[currentCableIndex];
        var cableKey = cable.cable_id;
       
        document.querySelector(`.group-terminal-no-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentGroups[rowIndex].terminal_no = this.value;
            unsavedGroupData[cableKey] = currentGroups;
        });
       
        document.querySelector(`.group-input-output-input[data-index="${rowIndex}"]`).addEventListener('change', function() {
            currentGroups[rowIndex].input_output = this.value;
            unsavedGroupData[cableKey] = currentGroups;
        });
       
        document.querySelector(`.group-text-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentGroups[rowIndex].text = this.value;
            unsavedGroupData[cableKey] = currentGroups;
        });
       
        document.querySelector(`.remove-group-row-btn[data-index="${rowIndex}"]`).addEventListener('click', function() {
            removeGroupRow(rowIndex);
        });
    }
    
    function removeGroupRow(rowIndex) {
        var cable = currentCables[currentCableIndex];
        var cableKey = cable.cable_id;
       
        currentGroups.splice(rowIndex, 1);
        unsavedGroupData[cableKey] = currentGroups;
        refreshGroupTable();
    }

    // ============================================================================
    // 📋 CHOKE CONFIGURATION FUNCTIONS
    // ============================================================================
    function openChokeConfigModal() {
        var cable = currentCables[currentCableIndex];
        document.getElementById('chokeCableId').value = cable.cable_id;
       
        var tableBody = document.getElementById('chokeTableBody');
        tableBody.innerHTML = '';
       
        var cableKey = cable.cable_id;
        if (unsavedChokeData[cableKey]) {
            currentChokes = unsavedChokeData[cableKey];
            refreshChokeTable();
            chokeModal.show();
            return;
        } else {
            currentChokes = [];
           
            // First, load draft data if exists
            loadChokeTableDraft(cable.cable_id);
            
            // Then check for existing final chokes
            fetch(`/get_chokes_for_cable?cable_id=${cable.cable_id}`)
                .then(response => response.json())
                .then(existingChokes => {
                    let nextChokeId = 1;
                    if (existingChokes && existingChokes.length > 0) {
                        const maxId = Math.max(...existingChokes.map(choke => parseInt(choke.choke_id.replace('CH', '')) || 0));
                        nextChokeId = maxId + 1;
                       
                        existingChokes.forEach(choke => {
                            // Only add if not already in currentChokes (from draft)
                            const exists = currentChokes.some(c => c.choke_id === choke.choke_id);
                            if (!exists) {
                                currentChokes.push({
                                    cable_id: choke.cable_id,
                                    choke_id: choke.choke_id,
                                    input_terminal: choke.input_terminal || '',
                                    output_terminal: choke.output_terminal || '',
                                    terminal_name: choke.terminal_name || 'CHOKE',
                                    output_type: choke.output_type || '',
                                    output_text: choke.output_text || '',
                                    output_connected: choke.output_connected || ''
                                });
                            }
                        });
                    }
                   
                    window.nextChokeId = nextChokeId;
                   
                    if (currentChokes.length === 0) {
                        addChokeRow();
                    } else {
                        refreshChokeTable();
                    }
                    chokeModal.show();
                })
                .catch(error => {
                    console.error('Error fetching existing chokes:', error);
                    window.nextChokeId = 1;
                    if (currentChokes.length === 0) {
                        addChokeRow();
                    }
                    chokeModal.show();
                });
        }
    }
    
    function addChokeRow() {
        var cable = currentCables[currentCableIndex];
        var chokeId = window.nextChokeId || 1;
       
        var choke = {
            cable_id: cable.cable_id,
            choke_id: chokeId.toString(),
            input_terminal: '',
            output_terminal: '',
            terminal_name: 'CHOKE',
            output_type: '',
            output_text: '',
            output_connected: ''
        };
       
        var rowIndex = currentChokes.length;
        currentChokes.push(choke);
       
        var cableKey = cable.cable_id;
        unsavedChokeData[cableKey] = currentChokes;
       
        refreshChokeTable();
       
        window.nextChokeId = chokeId + 1;
    }
    
    function refreshChokeTable() {
        var tableBody = document.getElementById('chokeTableBody');
        tableBody.innerHTML = '';
       
        currentChokes.forEach((choke, index) => {
            var row = document.createElement('tr');
            row.innerHTML = `
                <td>
                    <input type="text" class="form-control choke-cable-id-input" data-index="${index}" value="${choke.cable_id || ''}" readonly>
                </td>
                <td>
                    <input type="text" class="form-control choke-id-input" data-index="${index}" value="${choke.choke_id || ''}" readonly>
                </td>
                <td>
                    <input type="text" class="form-control choke-input-terminal-input" data-index="${index}" value="${choke.input_terminal || ''}" placeholder="Input terminal">
                </td>
                <td>
                    <input type="text" class="form-control choke-output-terminal-input" data-index="${index}" value="${choke.output_terminal || ''}" placeholder="Output terminal">
                </td>
                <td>
                    <input type="text" class="form-control choke-terminal-name-input" data-index="${index}" value="${choke.terminal_name || 'CHOKE'}" placeholder="Terminal name">
                </td>
                <td>
                    <select class="form-select choke-output-type-input" data-index="${index}" style="min-width: 120px;">
                        <option value="">Select Output Type</option>
                        <option value="relay" ${choke.output_type === 'relay' ? 'selected' : ''}>Relay</option>
                        <option value="relay_box" ${choke.output_type === 'relay_box' ? 'selected' : ''}>Relay Box</option>
                    </select>
                </td>
                <td>
                    <input type="text" class="form-control choke-output-text-input" data-index="${index}" value="${choke.output_text || ''}" placeholder="Output text">
                </td>
                <td>
                    <input type="text" class="form-control choke-output-connected-input" data-index="${index}" value="${choke.output_connected || ''}" placeholder="Output connected">
                </td>
                <td>
                    <button type="button" class="btn btn-sm btn-danger remove-choke-row-btn" data-index="${index}">
                        <i class="bi bi-trash"></i>
                    </button>
                </td>
            `;
            tableBody.appendChild(row);
            addChokeRowEventListeners(index);
        });
    }
    
    function addChokeRowEventListeners(rowIndex) {
        var cable = currentCables[currentCableIndex];
        var cableKey = cable.cable_id;
       
        document.querySelector(`.choke-input-terminal-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentChokes[rowIndex].input_terminal = this.value;
            unsavedChokeData[cableKey] = currentChokes;
        });
       
        document.querySelector(`.choke-output-terminal-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentChokes[rowIndex].output_terminal = this.value;
            unsavedChokeData[cableKey] = currentChokes;
        });
       
        document.querySelector(`.choke-terminal-name-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentChokes[rowIndex].terminal_name = this.value;
            unsavedChokeData[cableKey] = currentChokes;
        });
       
        document.querySelector(`.choke-output-type-input[data-index="${rowIndex}"]`).addEventListener('change', function() {
            currentChokes[rowIndex].output_type = this.value;
            unsavedChokeData[cableKey] = currentChokes;
        });
       
        document.querySelector(`.choke-output-text-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentChokes[rowIndex].output_text = this.value;
            unsavedChokeData[cableKey] = currentChokes;
        });
       
        document.querySelector(`.choke-output-connected-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentChokes[rowIndex].output_connected = this.value;
            unsavedChokeData[cableKey] = currentChokes;
        });
       
        document.querySelector(`.remove-choke-row-btn[data-index="${rowIndex}"]`).addEventListener('click', function() {
            removeChokeRow(rowIndex);
        });
    }
    
    function removeChokeRow(rowIndex) {
        var cable = currentCables[currentCableIndex];
        var cableKey = cable.cable_id;
       
        currentChokes.splice(rowIndex, 1);
        unsavedChokeData[cableKey] = currentChokes;
        refreshChokeTable();
    }

    // ============================================================================
    // 📋 RESISTOR CONFIGURATION FUNCTIONS
    // ============================================================================
    function openResistorConfigModal() {
        var cable = currentCables[currentCableIndex];
        document.getElementById('resistorCableId').value = cable.cable_id;
       
        var tableBody = document.getElementById('resistorTableBody');
        tableBody.innerHTML = '';
       
        var cableKey = cable.cable_id;
        if (unsavedResistorData[cableKey]) {
            currentResistors = unsavedResistorData[cableKey];
            refreshResistorTable();
            resistorModal.show();
            return;
        } else {
            currentResistors = [];
           
            fetch(`/get_resistors_for_cable?cable_id=${cable.cable_id}`)
                .then(response => response.json())
                .then(existingResistors => {
                    let nextResistorId = 1;
                    if (existingResistors && existingResistors.length > 0) {
                        const maxId = Math.max(...existingResistors.map(resistor => parseInt(resistor.resistor_id.replace('R', '')) || 0));
                        nextResistorId = maxId + 1;
                       
                        existingResistors.forEach(resistor => {
                            currentResistors.push({
                                cable_id: resistor.cable_id,
                                resistor_id: resistor.resistor_id,
                                input_terminal: resistor.input_terminal || '',
                                output_terminal: resistor.output_terminal || '',
                                resistor_name: resistor.resistor_name || 'R'
                            });
                        });
                    }
                   
                    window.nextResistorId = nextResistorId;
                   
                    if (currentResistors.length === 0) {
                        addResistorRow();
                    } else {
                        refreshResistorTable();
                    }
                    resistorModal.show();
                })
                .catch(error => {
                    console.error('Error fetching existing resistors:', error);
                    window.nextResistorId = 1;
                    if (currentResistors.length === 0) {
                        addResistorRow();
                    }
                    resistorModal.show();
                });
        }
    }
    
    function addResistorRow() {
        var cable = currentCables[currentCableIndex];
        var resistorId = window.nextResistorId || 1;
       
        var resistor = {
            cable_id: cable.cable_id,
            resistor_id: resistorId.toString(),
            input_terminal: '',
            output_terminal: '',
            resistor_name: 'R'
        };
       
        var rowIndex = currentResistors.length;
        currentResistors.push(resistor);
       
        var cableKey = cable.cable_id;
        unsavedResistorData[cableKey] = currentResistors;
       
        refreshResistorTable();
       
        window.nextResistorId = resistorId + 1;
    }
    
    function refreshResistorTable() {
        var tableBody = document.getElementById('resistorTableBody');
        tableBody.innerHTML = '';
       
        currentResistors.forEach((resistor, index) => {
            var row = document.createElement('tr');
            row.innerHTML = `
                <td>
                    <input type="text" class="form-control resistor-cable-id-input" data-index="${index}" value="${resistor.cable_id || ''}" readonly>
                </td>
                <td>
                    <input type="text" class="form-control resistor-id-input" data-index="${index}" value="${resistor.resistor_id || ''}" readonly>
                </td>
                <td>
                    <input type="text" class="form-control resistor-input-terminal-input" data-index="${index}" value="${resistor.input_terminal || ''}" placeholder="Input terminal">
                </td>
                <td>
                    <input type="text" class="form-control resistor-output-terminal-input" data-index="${index}" value="${resistor.output_terminal || ''}" placeholder="Output terminal">
                </td>
                <td>
                    <input type="text" class="form-control resistor-resistor-name-input" data-index="${index}" value="${resistor.resistor_name || 'R'}" placeholder="Resistor name">
                </td>
                <td>
                    <button type="button" class="btn btn-sm btn-danger remove-resistor-row-btn" data-index="${index}">
                        <i class="bi bi-trash"></i>
                    </button>
                </td>
            `;
            tableBody.appendChild(row);
            addResistorRowEventListeners(index);
        });
    }
    
    function addResistorRowEventListeners(rowIndex) {
        var cable = currentCables[currentCableIndex];
        var cableKey = cable.cable_id;
       
        document.querySelector(`.resistor-input-terminal-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentResistors[rowIndex].input_terminal = this.value;
            unsavedResistorData[cableKey] = currentResistors;
        });
       
        document.querySelector(`.resistor-output-terminal-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentResistors[rowIndex].output_terminal = this.value;
            unsavedResistorData[cableKey] = currentResistors;
        });
       
        document.querySelector(`.resistor-resistor-name-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentResistors[rowIndex].resistor_name = this.value;
            unsavedResistorData[cableKey] = currentResistors;
        });
       
        document.querySelector(`.remove-resistor-row-btn[data-index="${rowIndex}"]`).addEventListener('click', function() {
            removeResistorRow(rowIndex);
        });
    }
    
    function removeResistorRow(rowIndex) {
        var cable = currentCables[currentCableIndex];
        var cableKey = cable.cable_id;
       
        currentResistors.splice(rowIndex, 1);
        unsavedResistorData[cableKey] = currentResistors;
        refreshResistorTable();
    }
    // ============================================================================
    // 📋 TERMINAL HEADER DRAFT MANAGEMENT FUNCTIONS
    // ============================================================================
    function saveTerminalHeaderDraft() {
        const cableId = document.getElementById('headerCableId').value;
        
        if (!cableId) {
            showCompactNotification('✕ No cable ID found for saving terminal header draft.', 'error', 3000);
            return;
        }
        
        // Collect header data from the table
        const headerData = collectTerminalHeaderDataFromTable();
        
        if (headerData.length === 0) {
            showCompactNotification('✕ No header data to save as draft.', 'error', 3000);
            return;
        }
        
        const draftData = {
            cable_id: cableId,
            header_data: headerData
        };
        
        console.log('Saving terminal header draft:', draftData);
        
        const saveDraftBtn = document.getElementById('saveTerminalHeaderDraftBtn');
        const originalText = saveDraftBtn.innerHTML;
        saveDraftBtn.disabled = true;
        saveDraftBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving...';
        
        fetch('/save_terminal_header_draft', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(draftData)
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .then(result => {
            saveDraftBtn.disabled = false;
            saveDraftBtn.innerHTML = originalText;
            
            if (result.success) {
                showCompactNotification('✓ Terminal header draft saved successfully!', 'success', 2000);
                console.log('Terminal header draft saved:', result);
                
                // Update terminal header summary display
                updateTerminalHeaderSummaryDisplay();
            } else {
                showCompactNotification('✕ ' + (result.message || 'Failed to save draft.'), 'error', 3000);
            }
        })
        .catch(error => {
            saveDraftBtn.disabled = false;
            saveDraftBtn.innerHTML = originalText;
            showCompactNotification('✕ Error saving terminal header draft: ' + error.message, 'error', 3000);
            console.error('Error saving terminal header draft:', error);
        });
    }

    function loadTerminalHeaderDraft(cableId) {
        if (!cableId) {
            console.log('No cable ID provided for loading terminal header draft');
            return;
        }
        
        fetch(`/get_terminal_header_draft?cable_id=${encodeURIComponent(cableId)}`)
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                return response.json();
            })
            .then(result => {
                if (result.success && result.header_data && result.header_data.length > 0) {
                    console.log(`Loaded terminal header draft with ${result.header_data.length} headers for cable ${cableId}`);
                    
                    // Store the draft data for population
                    populateTerminalHeaderTablesFromDraft(result.header_data);
                    showCompactNotification('✓ Loaded terminal header draft', 'success', 2000);
                } else {
                    console.log('No terminal header draft found for this cable');
                }
            })
            .catch(error => {
                console.error('Error loading terminal header draft:', error);
                // No draft found or error - continue with default
            });
    }

    function collectTerminalHeaderDataFromTable() {
        const headerData = [];
        
        // Collect data from header table
        const tableBody = document.getElementById('headerTableBody');
        if (!tableBody) return headerData;
        
        tableBody.querySelectorAll('tr').forEach((row, index) => {
            const headerTypeInput = row.querySelector('.header-type-input');
            const terminalStartInput = row.querySelector('.header-terminal-start-input');
            const terminalEndInput = row.querySelector('.header-terminal-end-input');
            const inputOutputInput = row.querySelector('.header-input-output-input');
            const textInput = row.querySelector('.header-text-input');
            
            if (headerTypeInput && terminalStartInput && terminalEndInput) {
                const header = {
                    cable_id: document.getElementById('headerCableId').value,
                    header_type: headerTypeInput.value || '',
                    terminal_start: terminalStartInput.value || '',
                    terminal_end: terminalEndInput.value || '',
                    input_output: inputOutputInput ? inputOutputInput.value : '',
                    text: textInput ? textInput.value : ''
                };
                headerData.push(header);
            }
        });
        
        console.log(`Collected ${headerData.length} terminal headers from table`);
        return headerData;
    }

    function populateTerminalHeaderTablesFromDraft(headerData) {
        console.log('Populating terminal header tables from draft:', headerData);
        
        // Clear current headers array
        currentHeaders = [];
        
        // Clear table first
        const tableBody = document.getElementById('headerTableBody');
        tableBody.innerHTML = '';
        
        // Recreate rows from draft data
        headerData.forEach((draftHeader, index) => {
            // Create new row
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>
                    <select class="form-select header-type-input" data-index="${index}">
                        <option value="">Select Header Type</option>
                        <option value="WIREFROM" ${draftHeader.header_type === 'WIREFROM' ? 'selected' : ''}>WIREFROM</option>
                        <option value="WIRETO" ${draftHeader.header_type === 'WIRETO' ? 'selected' : ''}>WIRETO</option>
                        <option value="RELAY" ${draftHeader.header_type === 'RELAY' ? 'selected' : ''}>RELAY</option>
                        <option value="RELAY_BOX" ${draftHeader.header_type === 'RELAY_BOX' ? 'selected' : ''}>RELAY BOX</option>
                        <option value="RELAY_CONTACT_BOX" ${draftHeader.header_type === 'RELAY_CONTACT_BOX' ? 'selected' : ''}>RELAY CONTACT BOX</option>
                    </select>
                </td>
                <td>
                    <input type="number" class="form-control header-terminal-start-input" data-index="${index}" value="${draftHeader.terminal_start || ''}" placeholder="Start" min="1">
                    <small class="form-text text-muted">Based on cable name: ${draftHeader.terminal_start || '1'}</small>
                </td>
                <td>
                    <input type="number" class="form-control header-terminal-end-input" data-index="${index}" value="${draftHeader.terminal_end || ''}" placeholder="End" min="1">
                    <small class="form-text text-muted">Based on cable name: ${draftHeader.terminal_end || '12'}</small>
                </td>
                <td>
                    <select class="form-select header-input-output-input" data-index="${index}">
                        <option value="">Select I/O</option>
                        <option value="Input" ${draftHeader.input_output === 'Input' ? 'selected' : ''}>Input</option>
                        <option value="Output" ${draftHeader.input_output === 'Output' ? 'selected' : ''}>Output</option>
                    </select>
                </td>
                <td>
                    <input type="text" class="form-control header-text-input" data-index="${index}" value="${draftHeader.text || ''}" placeholder="Header text">
                </td>
                <td>
                    <button type="button" class="btn btn-sm btn-danger remove-header-row-btn" data-index="${index}">
                        <i class="bi bi-trash"></i>
                    </button>
                </td>
            `;
            tableBody.appendChild(row);
            
            // Add to currentHeaders array
            currentHeaders.push({
                cable_id: draftHeader.cable_id,
                header_type: draftHeader.header_type,
                terminal_start: draftHeader.terminal_start,
                terminal_end: draftHeader.terminal_end,
                input_output: draftHeader.input_output,
                text: draftHeader.text
            });
            
            // Add event listeners
            addHeaderRowEventListeners(index);
        });
    }

    function updateTerminalHeaderSummaryDisplay() {
        fetch('/get_terminal_header_summary')
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                return response.json();
            })
            .then(result => {
                if (result.success) {
                    // Update summary display element if it exists
                    const totalHeadersEl = document.getElementById('totalTerminalHeadersCount');
                    if (totalHeadersEl) {
                        totalHeadersEl.textContent = result.total_terminal_headers || 0;
                    }
                }
            })
            .catch(error => {
                console.error('Error updating terminal header summary:', error);
            });
    }

    // ============================================================================
    // 📋 GROUP TABLE DRAFT MANAGEMENT FUNCTIONS
    // ============================================================================
    function saveGroupTableDraft() {
        const cableId = document.getElementById('groupCableId').value;
        
        if (!cableId) {
            showCompactNotification('✕ No cable ID found for saving group draft.', 'error', 3000);
            return;
        }
        
        // Collect group data from the table
        const groupData = collectGroupDataFromTable();
        
        if (groupData.length === 0) {
            showCompactNotification('✕ No group data to save as draft.', 'error', 3000);
            return;
        }
        
        const draftData = {
            cable_id: cableId,
            group_data: groupData
        };
        
        console.log('Saving group draft:', draftData);
        
        const saveDraftBtn = document.getElementById('saveGroupTableDraftBtn');
        const originalText = saveDraftBtn.innerHTML;
        saveDraftBtn.disabled = true;
        saveDraftBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving...';
        
        fetch('/save_group_table_draft', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(draftData)
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .then(result => {
            saveDraftBtn.disabled = false;
            saveDraftBtn.innerHTML = originalText;
            
            if (result.success) {
                showCompactNotification('✓ Group table draft saved successfully!', 'success', 2000);
                console.log('Group draft saved:', result);
                
                // Update group summary display
                updateGroupSummaryDisplay();
            } else {
                showCompactNotification('✕ ' + (result.message || 'Failed to save draft.'), 'error', 3000);
            }
        })
        .catch(error => {
            saveDraftBtn.disabled = false;
            saveDraftBtn.innerHTML = originalText;
            showCompactNotification('✕ Error saving group table draft: ' + error.message, 'error', 3000);
            console.error('Error saving group table draft:', error);
        });
    }

    function loadGroupTableDraft(cableId) {
        if (!cableId) {
            console.log('No cable ID provided for loading group table draft');
            return;
        }
        
        fetch(`/get_group_table_draft?cable_id=${encodeURIComponent(cableId)}`)
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                return response.json();
            })
            .then(result => {
                if (result.success && result.group_data && result.group_data.length > 0) {
                    console.log(`Loaded group table draft with ${result.group_data.length} groups for cable ${cableId}`);
                    
                    // Store the draft data for population
                    populateGroupTablesFromDraft(result.group_data);
                    showCompactNotification('✓ Loaded group table draft', 'success', 2000);
                } else {
                    console.log('No group table draft found for this cable');
                }
            })
            .catch(error => {
                console.error('Error loading group table draft:', error);
                // No draft found or error - continue with default
            });
    }

    function collectGroupDataFromTable() {
        const groupData = [];
        
        // Collect data from group table
        const tableBody = document.getElementById('groupTableBody');
        if (!tableBody) return groupData;
        
        tableBody.querySelectorAll('tr').forEach((row, index) => {
            const groupIdInput = row.querySelector('.group-id-input');
            const terminalNoInput = row.querySelector('.group-terminal-no-input');
            const inputOutputInput = row.querySelector('.group-input-output-input');
            const textInput = row.querySelector('.group-text-input');
            
            if (groupIdInput) {
                // Get group ID - remove DRAFT- prefix if it exists
                let groupId = groupIdInput.value;
                if (groupId.startsWith('DRAFT-')) {
                    groupId = groupId.substring(6); // Remove 'DRAFT-' prefix
                }
                
                const group = {
                    cable_id: document.getElementById('groupCableId').value,
                    group_id: groupId || `GR${index + 1}`,
                    terminal_no: terminalNoInput ? terminalNoInput.value : '',
                    input_output: inputOutputInput ? inputOutputInput.value : '',
                    text: textInput ? textInput.value : ''
                };
                groupData.push(group);
            }
        });
        
        console.log(`Collected ${groupData.length} groups from table`);
        return groupData;
    }

    function populateGroupTablesFromDraft(groupData) {
        console.log('Populating group tables from draft:', groupData);
        
        // Clear current groups array
        currentGroups = [];
        
        // Clear table first
        const tableBody = document.getElementById('groupTableBody');
        tableBody.innerHTML = '';
        
        // Recreate rows from draft data
        groupData.forEach((draftGroup, index) => {
            // Create new row
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>
                    <input type="text" class="form-control group-id-input" data-index="${index}" value="${draftGroup.group_id || ''}" readonly>
                </td>
                <td>
                    <input type="text" class="form-control group-terminal-no-input" data-index="${index}" value="${draftGroup.terminal_no || ''}" placeholder="Terminal numbers (e.g., 1,2,4,54)">
                </td>
                <td>
                    <select class="form-select group-input-output-input" data-index="${index}">
                        <option value="">Select I/O</option>
                        <option value="Input" ${draftGroup.input_output === 'Input' ? 'selected' : ''}>Input</option>
                        <option value="Output" ${draftGroup.input_output === 'Output' ? 'selected' : ''}>Output</option>
                    </select>
                </td>
                <td>
                    <input type="text" class="form-control group-text-input" data-index="${index}" value="${draftGroup.text || ''}" placeholder="Group text">
                </td>
                <td>
                    <button type="button" class="btn btn-sm btn-danger remove-group-row-btn" data-index="${index}">
                        <i class="bi bi-trash"></i>
                    </button>
                </td>
            `;
            tableBody.appendChild(row);
            
            // Add to currentGroups array
            currentGroups.push({
                cable_id: draftGroup.cable_id,
                group_id: draftGroup.group_id,
                terminal_no: draftGroup.terminal_no,
                input_output: draftGroup.input_output,
                text: draftGroup.text
            });
            
            // Add event listeners
            addGroupRowEventListeners(index);
        });
    }

    function updateGroupSummaryDisplay() {
        fetch('/get_group_summary')
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                return response.json();
            })
            .then(result => {
                if (result.success) {
                    // Update summary display element if it exists
                    const totalGroupsEl = document.getElementById('totalGroupsCount');
                    if (totalGroupsEl) {
                        totalGroupsEl.textContent = result.total_groups || 0;
                    }
                }
            })
            .catch(error => {
                console.error('Error updating group summary:', error);
            });
    }

    // ============================================================================
    // 📋 CHOKE TABLE DRAFT MANAGEMENT FUNCTIONS
    // ============================================================================
    function saveChokeTableDraft() {
        const cableId = document.getElementById('chokeCableId').value;
        
        if (!cableId) {
            showCompactNotification('✕ No cable ID found for saving choke draft.', 'error', 3000);
            return;
        }
        
        // Collect choke data from the table
        const chokeData = collectChokeDataFromTable();
        
        if (chokeData.length === 0) {
            showCompactNotification('✕ No choke data to save as draft.', 'error', 3000);
            return;
        }
        
        const draftData = {
            cable_id: cableId,
            choke_data: chokeData
        };
        
        console.log('Saving choke draft:', draftData);
        
        const saveDraftBtn = document.getElementById('saveChokeTableDraftBtn');
        const originalText = saveDraftBtn.innerHTML;
        saveDraftBtn.disabled = true;
        saveDraftBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving...';
        
        fetch('/save_choke_table_draft', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(draftData)
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .then(result => {
            saveDraftBtn.disabled = false;
            saveDraftBtn.innerHTML = originalText;
            
            if (result.success) {
                showCompactNotification('✓ Choke table draft saved successfully!', 'success', 2000);
                console.log('Choke draft saved:', result);
                
                // Update choke summary display
                updateChokeSummaryDisplay();
            } else {
                showCompactNotification('✕ ' + (result.message || 'Failed to save draft.'), 'error', 3000);
            }
        })
        .catch(error => {
            saveDraftBtn.disabled = false;
            saveDraftBtn.innerHTML = originalText;
            showCompactNotification('✕ Error saving choke table draft: ' + error.message, 'error', 3000);
            console.error('Error saving choke table draft:', error);
        });
    }

    function loadChokeTableDraft(cableId) {
        if (!cableId) {
            console.log('No cable ID provided for loading choke table draft');
            return;
        }
        
        fetch(`/get_choke_table_draft?cable_id=${encodeURIComponent(cableId)}`)
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                return response.json();
            })
            .then(result => {
                if (result.success && result.choke_data && result.choke_data.length > 0) {
                    console.log(`Loaded choke table draft with ${result.choke_data.length} chokes for cable ${cableId}`);
                    
                    // Store the draft data for population
                    populateChokeTablesFromDraft(result.choke_data);
                    showCompactNotification('✓ Loaded choke table draft', 'success', 2000);
                } else {
                    console.log('No choke table draft found for this cable');
                }
            })
            .catch(error => {
                console.error('Error loading choke table draft:', error);
                // No draft found or error - continue with default
            });
    }

    function collectChokeDataFromTable() {
        const chokeData = [];
        
        // Collect data from choke table
        const tableBody = document.getElementById('chokeTableBody');
        if (!tableBody) return chokeData;
        
        tableBody.querySelectorAll('tr').forEach((row, index) => {
            const cableIdInput = row.querySelector('.choke-cable-id-input');
            const chokeIdInput = row.querySelector('.choke-id-input');
            const inputTerminalInput = row.querySelector('.choke-input-terminal-input');
            const outputTerminalInput = row.querySelector('.choke-output-terminal-input');
            const terminalNameInput = row.querySelector('.choke-terminal-name-input');
            const outputTypeInput = row.querySelector('.choke-output-type-input');
            const outputTextInput = row.querySelector('.choke-output-text-input');
            const outputConnectedInput = row.querySelector('.choke-output-connected-input');
            
            if (cableIdInput && chokeIdInput) {
                // Get choke ID - remove DRAFT- prefix if it exists
                let chokeId = chokeIdInput.value;
                if (chokeId.startsWith('DRAFT-')) {
                    chokeId = chokeId.substring(6); // Remove 'DRAFT-' prefix
                }
                
                const choke = {
                    cable_id: cableIdInput.value,
                    choke_id: chokeId || `CH${index + 1}`,
                    input_terminal: inputTerminalInput ? inputTerminalInput.value : '',
                    output_terminal: outputTerminalInput ? outputTerminalInput.value : '',
                    terminal_name: terminalNameInput ? terminalNameInput.value : 'CHOKE',
                    output_type: outputTypeInput ? outputTypeInput.value : '',
                    output_text: outputTextInput ? outputTextInput.value : '',
                    output_connected: outputConnectedInput ? outputConnectedInput.value : ''
                };
                chokeData.push(choke);
            }
        });
        
        console.log(`Collected ${chokeData.length} chokes from table`);
        return chokeData;
    }

    function populateChokeTablesFromDraft(chokeData) {
        console.log('Populating choke tables from draft:', chokeData);
        
        // Clear current chokes array
        currentChokes = [];
        
        // Clear table first
        const tableBody = document.getElementById('chokeTableBody');
        tableBody.innerHTML = '';
        
        // Recreate rows from draft data
        chokeData.forEach((draftChoke, index) => {
            // Create new row
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>
                    <input type="text" class="form-control choke-cable-id-input" data-index="${index}" value="${draftChoke.cable_id || ''}" readonly>
                </td>
                <td>
                    <input type="text" class="form-control choke-id-input" data-index="${index}" value="${draftChoke.choke_id || ''}" readonly>
                </td>
                <td>
                    <input type="text" class="form-control choke-input-terminal-input" data-index="${index}" value="${draftChoke.input_terminal || ''}" placeholder="Input terminal">
                </td>
                <td>
                    <input type="text" class="form-control choke-output-terminal-input" data-index="${index}" value="${draftChoke.output_terminal || ''}" placeholder="Output terminal">
                </td>
                <td>
                    <input type="text" class="form-control choke-terminal-name-input" data-index="${index}" value="${draftChoke.terminal_name || 'CHOKE'}" placeholder="Terminal name">
                </td>
                <td>
                    <select class="form-select choke-output-type-input" data-index="${index}" style="min-width: 120px;">
                        <option value="">Select Output Type</option>
                        <option value="relay" ${draftChoke.output_type === 'relay' ? 'selected' : ''}>Relay</option>
                        <option value="relay_box" ${draftChoke.output_type === 'relay_box' ? 'selected' : ''}>Relay Box</option>
                    </select>
                </td>
                <td>
                    <input type="text" class="form-control choke-output-text-input" data-index="${index}" value="${draftChoke.output_text || ''}" placeholder="Output text">
                </td>
                <td>
                    <input type="text" class="form-control choke-output-connected-input" data-index="${index}" value="${draftChoke.output_connected || ''}" placeholder="Output connected">
                </td>
                <td>
                    <button type="button" class="btn btn-sm btn-danger remove-choke-row-btn" data-index="${index}">
                        <i class="bi bi-trash"></i>
                    </button>
                </td>
            `;
            tableBody.appendChild(row);
            
            // Add to currentChokes array
            currentChokes.push({
                cable_id: draftChoke.cable_id,
                choke_id: draftChoke.choke_id,
                input_terminal: draftChoke.input_terminal,
                output_terminal: draftChoke.output_terminal,
                terminal_name: draftChoke.terminal_name,
                output_type: draftChoke.output_type,
                output_text: draftChoke.output_text,
                output_connected: draftChoke.output_connected
            });
            
            // Add event listeners
            addChokeRowEventListeners(index);
        });
    }

    function updateChokeSummaryDisplay() {
        fetch('/get_choke_summary')
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                return response.json();
            })
            .then(result => {
                if (result.success) {
                    // Update summary display element if it exists
                    const totalChokesEl = document.getElementById('totalChokesCount');
                    if (totalChokesEl) {
                        totalChokesEl.textContent = result.total_chokes || 0;
                    }
                }
            })
            .catch(error => {
                console.error('Error updating choke summary:', error);
            });
    }

    // ============================================================================
    // 📋 FORM SUBMISSION & NAVIGATION
    // ============================================================================
    document.getElementById('addMoreJunctionGridForm')?.addEventListener('submit', function(e) {
        e.preventDefault();
        forceRevalidateAllRows();
        if (hasOpenCableModalWithUnsavedCables()) {
            showCompactNotification('✕ Please save all cables in the cable configuration modal before saving locations. Click "Save All Cables" first.', 'error', 3000);
            return;
        }
        if (hasUnsavedCables()) {
            showCompactNotification('✕ Please configure and save all cables for all locations before saving. Click the "Config" button for each location to configure cables.', 'error', 3000);
            return;
        }
        if (!areAllFormFieldsValid()) {
            showCompactNotification('✕ Please fill all required fields (Location Name, Location Size, and Location Row) for all locations.', 'error', 3000);
            return;
        }
        var formData = new FormData(this);
        var junctions = [];
        var hasValidData = false;
        for (var i = 0; i < requestedJunctions; i++) {
            var junction = {};
            var isValidRow = true;
            formData.forEach((value, key) => {
                if (key.startsWith(`junctions[${i}]`)) {
                    var field = key.match(/junctions\[\d+\]\[(.*)\]/)[1];
                    junction[field] = value;
                }
            });
            if (!junction.junction_name || junction.junction_name.trim() === '') {
                isValidRow = false;
            }
            if (!['Full', 'Half'].includes(junction.junction_size)) {
                isValidRow = false;
            }
            try {
                var rowCount = parseInt(junction.junction_row);
                if (!rowCount || rowCount <= 0) {
                    isValidRow = false;
                }
            } catch (e) {
                isValidRow = false;
            }
            if (Object.keys(junction).length > 0 && isValidRow) {
                junctions.push(junction);
                hasValidData = true;
            }
        }
        if (!hasValidData) {
            showCompactNotification('✕ Please enter a location name, select a valid size, and enter a Location Row count for at least one location.', 'error', 3000);
            return;
        }
        var submitBtn = document.getElementById('saveMoreJunctionsBtn');
        submitBtn.disabled = true;
        submitBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving...';
        fetch('/add_junctions_ajax', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({ junctions: junctions })
        })
        .then(response => {
            if (!response.ok) {
                return response.text().then(text => {
                    throw new Error(`Error: ${response.status} - ${text}`);
                });
            }
            return response.json();
        })
        .then(result => {
            submitBtn.disabled = false;
            submitBtn.innerHTML = '<i class="bi bi-save"></i> Save Additional Locations';
            if (result.success && result.junctions && result.junctions.length > 0) {
                document.querySelector('.card.border-primary').style.display = 'none';
                showCompactNotification(' Locations saved successfully!', 'success', 2000);
                setTimeout(() => {
                    window.location.href = '/workflow/step/5?show_header_modal=true';
                }, 1500);
            } else {
                showCompactNotification('✕ ' + (result.message || 'Failed to add locations.'), 'error', 3000);
            }
        })
        .catch(error => {
            submitBtn.disabled = false;
            submitBtn.innerHTML = '<i class="bi bi-save"></i> Save Additional Locations';
            showCompactNotification('✕ ' + (error.message || 'An error occurred.'), 'error', 3000);
        });
    });

    // ============================================================================
    // 📋 UPDATED: SAVE TERMINAL HEADER DRAFT BEFORE GROUPS FUNCTION
    // ============================================================================
    function saveTerminalHeaderDraftBeforeGroups() {
        const cableId = document.getElementById('headerCableId').value;
        
        if (!cableId) {
            showCompactNotification('✕ No cable ID found for saving header draft.', 'error', 3000);
            openGroupConfigModal(); // Still open groups even if no cable ID
            return;
        }
        
        // Collect header data from the table
        const headerData = collectTerminalHeaderDataFromTable();
        
        if (headerData.length === 0) {
            showCompactNotification('✕ No header data to save as draft.', 'error', 3000);
            openGroupConfigModal(); // Still open groups even if no data
            return;
        }
        
        const draftData = {
            cable_id: cableId,
            header_data: headerData
        };
        
        console.log('Saving terminal header draft before groups:', draftData);
        
        const configureGroupsBtn = document.getElementById('configureGroupsBtn');
        const originalText = configureGroupsBtn.innerHTML;
        configureGroupsBtn.disabled = true;
        configureGroupsBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving...';
        
        fetch('/save_terminal_header_draft', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(draftData)
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .then(result => {
            configureGroupsBtn.disabled = false;
            configureGroupsBtn.innerHTML = originalText;
            
            if (result.success) {
                showCompactNotification('✓ Header draft saved successfully!', 'success', 2000);
                console.log('Header draft saved before groups:', result);
                
                // Update terminal header summary display
                updateTerminalHeaderSummaryDisplay();
                
                // Now open the groups modal
                setTimeout(() => {
                    openGroupConfigModal();
                }, 500);
            } else {
                showCompactNotification('✕ ' + (result.message || 'Failed to save draft.'), 'error', 3000);
                // Still open groups modal even if draft save fails
                setTimeout(() => {
                    openGroupConfigModal();
                }, 500);
            }
        })
        .catch(error => {
            configureGroupsBtn.disabled = false;
            configureGroupsBtn.innerHTML = originalText;
            showCompactNotification('✕ Error saving header draft: ' + error.message, 'error', 3000);
            console.error('Error saving header draft:', error);
            // Still open groups modal even if error
            setTimeout(() => {
                openGroupConfigModal();
            }, 500);
        });
    }

    // ============================================================================
    // 📋 UPDATED: SAVE GROUP TABLE DRAFT BEFORE CHOKE FUNCTION
    // ============================================================================
    function saveGroupTableDraftBeforeChoke() {
        const cableId = document.getElementById('groupCableId').value;
        
        if (!cableId) {
            showCompactNotification('✕ No cable ID found for saving group draft.', 'error', 3000);
            openChokeConfigModal(); // Still open choke even if no cable ID
            return;
        }
        
        // Collect group data from the table
        const groupData = collectGroupDataFromTable();
        
        if (groupData.length === 0) {
            showCompactNotification('✕ No group data to save as draft.', 'error', 3000);
            openChokeConfigModal(); // Still open choke even if no data
            return;
        }
        
        const draftData = {
            cable_id: cableId,
            group_data: groupData
        };
        
        console.log('Saving group table draft before choke:', draftData);
        
        const configureChokeBtn = document.getElementById('configureChokeBtn');
        const originalText = configureChokeBtn.innerHTML;
        configureChokeBtn.disabled = true;
        configureChokeBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving...';
        
        fetch('/save_group_table_draft', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(draftData)
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .then(result => {
            configureChokeBtn.disabled = false;
            configureChokeBtn.innerHTML = originalText;
            
            if (result.success) {
                showCompactNotification('✓ Group draft saved successfully!', 'success', 2000);
                console.log('Group draft saved before choke:', result);
                
                // Update group summary display
                updateGroupSummaryDisplay();
                
                // Now open the choke modal
                setTimeout(() => {
                    openChokeConfigModal();
                }, 500);
            } else {
                showCompactNotification('✕ ' + (result.message || 'Failed to save draft.'), 'error', 3000);
                // Still open choke modal even if draft save fails
                setTimeout(() => {
                    openChokeConfigModal();
                }, 500);
            }
        })
        .catch(error => {
            configureChokeBtn.disabled = false;
            configureChokeBtn.innerHTML = originalText;
            showCompactNotification('✕ Error saving group draft: ' + error.message, 'error', 3000);
            console.error('Error saving group draft:', error);
            // Still open choke modal even if error
            setTimeout(() => {
                openChokeConfigModal();
            }, 500);
        });
    }
})();