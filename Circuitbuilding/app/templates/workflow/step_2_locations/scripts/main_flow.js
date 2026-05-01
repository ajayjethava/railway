// Main flow JavaScript for Step 2 Locations
(function() {
    var requestedJunctions = {{ junction_count }};
    var existingJunctions = {{ rows|length if rows else 0 }};
    var freshNavigation = {{ 'true' if fresh_navigation else 'false' }};
    
    var remainingJunctions;
    var actualExisting = existingJunctions;
    
    if (freshNavigation || existingJunctions === 0) {
        remainingJunctions = requestedJunctions;
        actualExisting = 0;
    } else {
        remainingJunctions = requestedJunctions - existingJunctions;
        if (remainingJunctions < 0) remainingJunctions = 0;
    }
    
    var startJunctionNumber = {{ start_junction_number }};
    var stationIdSource = "{{ station_id_source if station_id_source else '' }}";
    
    var configModal, cableModal, successModal, errorModal, terminalModal, headerModal, groupModal;
    var currentJunctionIndex = null;
    var currentCables = [];
    var currentRowsConfig = [];
    var currentCableIndex = null;
    var currentHeaders = [];
    var currentGroups = [];

    var sheetName = "{{ sheet_name }}";
    var currentStep = {{ step }};

    // Function to format location name based on size
    function formatLocationName(baseName, size) {
        if (!baseName) return baseName;
        
        // Remove existing suffixes if any
        baseName = baseName.replace(/\s*\([FH]\)$/, '');
        
        if (size === 'Full') {
            return baseName + ' (F)';
        } else if (size === 'Half') {
            return baseName + ' (H)';
        }
        return baseName;
    }

    // Function to update location name in real-time
    function updateLocationName(junctionIndex) {
        const nameInput = document.querySelector(`input[name="junctions[${junctionIndex}][junction_name]"]`);
        const sizeSelect = document.querySelector(`select[name="junctions[${junctionIndex}][junction_size]"]`);
        
        if (!nameInput || !sizeSelect) return;
        
        let baseName = nameInput.value.trim();
        const size = sizeSelect.value;
        
        // Store original name without suffix for editing
        if (!nameInput.dataset.originalName && baseName) {
            nameInput.dataset.originalName = baseName.replace(/\s*\([FH]\)$/, '');
        }
        
        const formattedName = formatLocationName(baseName, size);
        
        // Update preview if exists
        const previewElement = document.querySelector(`.location-row[data-index="${junctionIndex}"] .location-name-preview`);
        const previewText = document.querySelector(`.location-row[data-index="${junctionIndex}"] .location-name-preview .preview-text`);
        
        if (previewElement && previewText) {
            if (baseName && size) {
                previewText.textContent = formattedName;
                previewElement.style.display = 'block';
            } else {
                previewElement.style.display = 'none';
            }
        }
        
        return formattedName;
    }

    document.addEventListener('DOMContentLoaded', function() {
        configModal = new bootstrap.Modal(document.getElementById('cableConfigModal'));
        cableModal = new bootstrap.Modal(document.getElementById('addCableModal'));
        successModal = new bootstrap.Modal(document.getElementById('successFlashModal'));
        errorModal = new bootstrap.Modal(document.getElementById('errorModal'));
        terminalModal = new bootstrap.Modal(document.getElementById('terminalConfigModal'));
        headerModal = new bootstrap.Modal(document.getElementById('headerConfigModal'));
        groupModal = new bootstrap.Modal(document.getElementById('groupConfigModal'));

        // Enable Config buttons when valid data is entered
        document.querySelectorAll('.config-cables-btn').forEach(button => {
            button.addEventListener('click', function() {
                currentJunctionIndex = parseInt(this.dataset.junctionIndex);
                openCableConfigModal(currentJunctionIndex);
            });
        });

        // Handle Edit Cables buttons
        document.querySelectorAll('.edit-cables-btn').forEach(button => {
            button.addEventListener('click', function(e) {
                e.preventDefault();
                currentJunctionIndex = parseInt(this.dataset.junctionIndex);
                openCableConfigModal(currentJunctionIndex);
            });
        });

        // Enable Config button when all required fields are valid
        document.querySelectorAll('.junction-name-input, .junction-size-select, .junction-row-input').forEach(input => {
            input.addEventListener('input', function() {
                var rowIndex = parseInt(this.name.match(/junctions\[(\d+)\]/)[1]);
                validateRow(rowIndex);
            });
        });

        // For select elements, also listen for change events
        document.querySelectorAll('.junction-size-select').forEach(select => {
            select.addEventListener('change', function() {
                var rowIndex = parseInt(this.name.match(/junctions\[(\d+)\]/)[1]);
                const formattedName = updateLocationName(rowIndex);
                
                // Update the input value with formatted name
                const nameInput = document.querySelector(`input[name="junctions[${rowIndex}][junction_name]"]`);
                if (nameInput && formattedName) {
                    nameInput.value = formattedName;
                }
                
                validateRow(rowIndex);
            });
        });

        // Auto-format location name when name input loses focus
        document.querySelectorAll('.junction-name-input').forEach(input => {
            input.addEventListener('blur', function() {
                var rowIndex = parseInt(this.name.match(/junctions\[(\d+)\]/)[1]);
                const sizeSelect = document.querySelector(`select[name="junctions[${rowIndex}][junction_size]"]`);
                
                if (sizeSelect && sizeSelect.value) {
                    const formattedName = updateLocationName(rowIndex);
                    this.value = formattedName;
                }
                
                validateRow(rowIndex);
            });
        });

        // Auto-open cable config when Location Row is entered
        document.querySelectorAll('.junction-row-input').forEach(input => {
            input.addEventListener('blur', function() {
                var rowIndex = parseInt(this.dataset.junctionIndex);
                var rowValue = this.value.trim();
                
                if (rowValue && !isNaN(rowValue) && rowValue >= 1) {
                    if (validateRow(rowIndex)) {
                        currentJunctionIndex = rowIndex;
                        setTimeout(() => {
                            openCableConfigModal(rowIndex);
                        }, 300);
                    }
                }
            });
        });

        // Cable Config Form submission
        document.getElementById('cableConfigForm').addEventListener('submit', function(e) {
            e.preventDefault();
            
            // Collect all row configurations
            currentRowsConfig = [];
            var isValid = true;
            
            document.querySelectorAll('#configTableBody tr').forEach((row, index) => {
                var rowName = row.querySelector('.config-row-name').value.trim();
                var cableCount = parseInt(row.querySelector('.config-cable-count').value);
                
                if (!rowName || isNaN(cableCount) || cableCount < 1) {
                    isValid = false;
                    return;
                }
                
                currentRowsConfig.push({
                    rowName: rowName,
                    cableCount: cableCount
                });
            });
            
            if (!isValid) {
                document.getElementById('errorMessage').textContent = 'Please enter valid row names and cable counts for all rows.';
                errorModal.show();
                return;
            }

            configModal.hide();
            openCableTableModal(currentJunctionIndex);
        });

        // Save All Cables button event
        document.getElementById('saveAllCablesBtn').addEventListener('click', saveAllCables);

        // Save All Terminals button event
        document.getElementById('saveAllTerminalsBtn').addEventListener('click', saveAllTerminals);

        // Header Configuration Button event
        document.getElementById('configureHeadersBtn').addEventListener('click', function() {
            openHeaderConfigModal();
        });

        // Configure Groups Button event
        document.getElementById('configureGroupsBtn').addEventListener('click', function() {
            openGroupConfigModal();
        });

        // Add Header Row button event
        document.getElementById('addHeaderRowBtn').addEventListener('click', function() {
            addHeaderRow();
        });

        // Add Group Row button event
        document.getElementById('addGroupRowBtn').addEventListener('click', function() {
            addGroupRow();
        });

        // Save All Headers button event
        document.getElementById('saveAllHeadersBtn').addEventListener('click', saveAllHeaders);

        // Save All Groups button event
        document.getElementById('saveAllGroupsBtn').addEventListener('click', saveAllGroups);

        if (existingJunctions > 0) {
            // If there are existing locations, enable the next step button
            document.getElementById('nextStepBtn').disabled = false;
        } else if (requestedJunctions === 0) {
            showFinalCompletion();
        }
    });

    function validateRow(rowIndex) {
        var junctionName = document.querySelector(`input[name="junctions[${rowIndex}][junction_name]"]`).value;
        var junctionSize = document.querySelector(`select[name="junctions[${rowIndex}][junction_size]"]`).value;
        var junctionRow = document.querySelector(`input[name="junctions[${rowIndex}][junction_row]"]`).value;
        var configButton = document.querySelector(`.config-cables-btn[data-junction-index="${rowIndex}"]`);

        var isValid = junctionName && junctionName.trim() !== '' && 
                      ['Full', 'Half'].includes(junctionSize) && 
                      junctionRow && parseInt(junctionRow) > 0;

        configButton.disabled = !isValid;
        return isValid;
    }

    function openCableConfigModal(junctionIndex) {
        var junctionRowValue = document.querySelector(`input[name="junctions[${junctionIndex}][junction_row]"]`).value;
        var numRows = parseInt(junctionRowValue);
        
        // Generate configuration table rows
        var tableBody = document.getElementById('configTableBody');
        tableBody.innerHTML = '';
        
        for (var i = 1; i <= numRows; i++) {
            var row = document.createElement('tr');
            row.innerHTML = `
                <td>
                    <strong>${i}</strong>
                </td>
                <td>
                    <input type="text" class="form-control config-row-name" value="${String.fromCharCode(64 + i)}" placeholder="Enter row name (e.g., A, B, C...)" required>
                </td>
                <td>
                    <input type="number" class="form-control config-cable-count" value="1" min="1" required>
                </td>
            `;
            tableBody.appendChild(row);
        }
        
        configModal.show();
    }

    function openCableTableModal(junctionIndex) {
        var junction = {};
        var formData = new FormData(document.getElementById('junctionGridForm'));

        formData.forEach((value, key) => {
            if (key.startsWith(`junctions[${junctionIndex}]`)) {
                var field = key.match(/junctions\[\d+\]\[(.*)\]/)[1];
                junction[field] = value;
            }
        });

        var junctionName = junction.junction_name || junction.junction_id || 'LOC' + (junction.junction_id || (startJunctionNumber + junctionIndex));
        var junctionSize = junction.junction_size || 'Not Set';

        // Set modal title
        document.getElementById('cableModalJunctionName').textContent = junctionName;

        // Generate cable tables for each row
        var container = document.getElementById('cableTablesContainer');
        container.innerHTML = '';

        currentCables = [];
        var overallPosition = 1;

        // FIRST: Fetch existing cables to check what already exists
        fetch('/get_existing_cables')
            .then(response => {
                if (!response.ok) {
                    console.log('No existing cables found or error fetching cables, starting from ID 1');
                    return [];
                }
                return response.json();
            })
            .then(existingCables => {
                console.log('Existing cables:', existingCables);
                
                if (existingCables && existingCables.error) {
                    console.log('Error from server:', existingCables.error);
                    existingCables = [];
                }
                
                if (!existingCables || existingCables.length === 0) {
                    console.log('No existing cables found, starting from cable ID 1');
                    existingCables = [];
                }

                // Find the maximum cable_id from ALL existing cables
                var maxCableId = 0;
                if (existingCables && existingCables.length > 0) {
                    existingCables.forEach(cable => {
                        var id = parseInt(cable.cable_id);
                        if (!isNaN(id) && id > maxCableId) {
                            maxCableId = id;
                        }
                    });
                }
                console.log('Maximum existing cable ID:', maxCableId);

                // Start from the next available ID
                var nextCableId = maxCableId + 1;
                
                currentRowsConfig.forEach((rowConfig, rowIndex) => {
                    // Create a table for each row
                    var rowTable = document.createElement('div');
                    rowTable.className = 'card mb-4';
                    rowTable.innerHTML = `
                        <div class="card-header bg-light">
                            <h6 class="mb-0">Row: ${rowConfig.rowName} | Cables: ${rowConfig.cableCount}</h6>
                        </div>
                        <div class="card-body">
                            <div class="table-responsive">
                                <table class="table table-bordered table-hover">
                                    <thead class="table-light">
                                        <tr>
                                            <th>Position</th>
                                            <th>Location Name</th>
                                            <th>Location Size</th>
                                            <th>Row</th>
                                            <th>Location Box</th>
                                            <th>Cable Id</th>
                                            <th>Terminal</th>
                                            <th>Start No</th>
                                            <th>Cable Name</th>
                                            <th>Status</th>
                                            <th>Actions</th>
                                        </tr>
                                    </thead>
                                    <tbody id="cableTableBody-${rowIndex}">
                                        <!-- Cable rows for this specific row will be generated here -->
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    `;
                    container.appendChild(rowTable);

                    // Generate cable rows for this specific row
                    var rowTableBody = document.getElementById(`cableTableBody-${rowIndex}`);
                    
                    // Start No calculation based on previous cable's terminal
                    var currentStartNo = 1; // Start from 1 for the first cable in row
                    
                    for (var pos = 1; pos <= rowConfig.cableCount; pos++) {
                        // Check if this cable already exists for this location and row
                        var existingCable = null;
                        if (existingCables && existingCables.length > 0) {
                            existingCable = existingCables.find(cable => 
                                cable.junction_name === junctionName && 
                                cable.row === rowConfig.rowName && 
                                parseInt(cable.junction_box) === pos
                            );
                        }

                        var cableId, startNo, terminal;

                        if (existingCable) {
                            // Use existing cable ID and start_no
                            cableId = parseInt(existingCable.cable_id);
                            startNo = parseInt(existingCable.start_no) || 1;
                            terminal = parseInt(existingCable.terminal) || 12;

                            // Update currentStartNo for the next cable based on this cable's terminal count
                            currentStartNo = startNo + terminal;
                        } else {
                            // Calculate new cable ID and start_no
                            cableId = nextCableId;
                            startNo = currentStartNo;
                            terminal = 12; // Default terminal count

                            // Update currentStartNo for the next cable
                            currentStartNo = startNo + terminal;

                            // Increment cable ID for next cable
                            nextCableId++;
                        }

                        // Calculate cable name
                        var endNo = startNo + terminal - 1;
                        var cableName = rowConfig.rowName + ' T' + startNo + '-' + endNo;

                        var cable = {
                            position: overallPosition,
                            junction_name: junctionName,
                            junction_size: junctionSize,
                            row: rowConfig.rowName,
                            junction_box: pos,
                            cable_id: cableId,
                            terminal: terminal,
                            start_no: startNo,
                            cable_name: cableName
                        };

                        currentCables.push(cable);

                        var cableRow = document.createElement('tr');
                        cableRow.innerHTML = `
                            <td>
                                <input type="text" class="form-control" value="${overallPosition}" readonly>
                            </td>
                            <td>
                                <input type="text" class="form-control" value="${junctionName}" readonly>
                            </td>
                            <td>
                                <input type="text" class="form-control" value="${junctionSize}" readonly>
                            </td>
                            <td>
                                <input type="text" class="form-control cable-row-input" data-index="${overallPosition-1}" value="${rowConfig.rowName}" readonly>
                            </td>
                            <td>
                                <input type="text" class="form-control cable-junction-box-input" data-index="${overallPosition-1}" value="${cable.junction_box}">
                            </td>
                            <td>
                                <input type="text" class="form-control cable-id-input" data-index="${overallPosition-1}" value="${cable.cable_id}" readonly>
                            </td>
                            <td>
                                <input type="number" class="form-control cable-terminal-input" data-index="${overallPosition-1}" value="${cable.terminal}">
                            </td>
                            <td>
                                <input type="number" class="form-control cable-start-no-input" data-index="${overallPosition-1}" value="${cable.start_no}">
                            </td>
                            <td>
                                <input type="text" class="form-control cable-name-input" data-index="${overallPosition-1}" value="${cableName}" readonly>
                            </td>
                            <td>
                                <span id="status-${overallPosition-1}" class="cable-status-indicator">
                                    <i class="bi bi-x-circle-fill text-danger" title="Terminals not configured"></i>
                                </span>
                            </td>
                            <td>
                                <button type="button" class="btn btn-sm btn-info config-terminals-btn" data-index="${overallPosition-1}">
                                    <i class="bi bi-gear"></i> Config Term
                                </button>
                            </td>
                        `;
                        rowTableBody.appendChild(cableRow);
                        overallPosition++;
                    }
                });

                // Check terminal status for all cables
                updateCableStatuses();

                // Add event listeners to inputs
                document.querySelectorAll('.cable-junction-box-input').forEach(input => {
                    input.addEventListener('input', function() {
                        var index = parseInt(this.dataset.index);
                        currentCables[index].junction_box = this.value;
                    });
                });

                document.querySelectorAll('.cable-terminal-input').forEach(input => {
                    input.addEventListener('input', function() {
                        var index = parseInt(this.dataset.index);
                        currentCables[index].terminal = this.value;
                        // Update cable name when terminal changes
                        updateCableName(index);
                        
                        // Update subsequent Start No values in the same row when terminal changes
                        updateSubsequentStartNos(index);
                    });
                });

                document.querySelectorAll('.cable-start-no-input').forEach(input => {
                    input.addEventListener('input', function() {
                        var index = parseInt(this.dataset.index);
                        currentCables[index].start_no = parseInt(this.value) || 1;
                        // Update cable name when start no changes
                        updateCableName(index);
                        
                        // Update subsequent Start No values in the same row
                        updateSubsequentStartNos(index);
                    });
                });

                // Function to update subsequent Start No values when a Start No or Terminal is changed
                function updateSubsequentStartNos(changedIndex) {
                    var changedCable = currentCables[changedIndex];
                    var rowName = changedCable.row;
                    var currentStartNo = parseInt(changedCable.start_no) || 1;
                    var terminalCount = parseInt(changedCable.terminal) || 12;
                    
                    // Find all cables in the same row that come after the changed cable
                    var subsequentCables = currentCables.filter((cable, idx) => 
                        cable.row === rowName && 
                        parseInt(cable.junction_box) > parseInt(changedCable.junction_box)
                    ).sort((a, b) => parseInt(a.junction_box) - parseInt(b.junction_box));
                    
                    // Update subsequent cables in the same row
                    var nextStartNo = currentStartNo + terminalCount;
                    subsequentCables.forEach(cable => {
                        cable.start_no = nextStartNo;
                        
                        // Update the input field
                        var startNoInput = document.querySelector(`.cable-start-no-input[data-index="${currentCables.indexOf(cable)}"]`);
                        if (startNoInput) {
                            startNoInput.value = nextStartNo;
                        }
                        
                        // Update cable name
                        updateCableName(currentCables.indexOf(cable));
                        
                        // Calculate next start number based on this cable's terminal count
                        nextStartNo += parseInt(cable.terminal) || 12;
                    });
                }

                // Add event listeners to config terminals buttons
                document.querySelectorAll('.config-terminals-btn').forEach(button => {
                    button.addEventListener('click', function() {
                        var index = parseInt(this.dataset.index);
                        currentCableIndex = index;
                        var cable = currentCables[index];
                        var terminalCount = parseInt(document.querySelector(`.cable-terminal-input[data-index="${index}"]`).value) || 0;
                        if (terminalCount < 1) {
                            document.getElementById('errorMessage').textContent = 'Terminal count must be at least 1 to configure.';
                            errorModal.show();
                            return;
                        }

                        // Save the cable first
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
                                // Open the terminal modal
                                document.getElementById('terminalModalCableId').textContent = cable.cable_id;
                                document.getElementById('terminalCountSpan').textContent = terminalCount;
                                var tableBody = document.getElementById('terminalTableBody');
                                tableBody.innerHTML = '';
                                for (var i = 1; i <= terminalCount; i++) {
                                    // Use the actual start_no from the cable and calculate terminal names
                                    var terminalName = parseInt(cable.start_no) + i - 1;
                                    var inputLeft = ((i-1) % 4 < 2) ? "120ATPR" : "120BTPR";
                                    var row = document.createElement('tr');
                                    row.innerHTML = `
                                        <td><input type="text" class="form-control cable-id-input" value="${cable.cable_id}" readonly></td>
                                        <td><input type="text" class="form-control terminal-id-input" value="${i}" placeholder="Enter terminal id"></td>
                                        <td><input type="text" class="form-control terminal-name-input" value="${terminalName}" placeholder="Enter terminal number"></td>
                                        <td>
                                            <select class="form-select symbol-input">
                                                <option value="ara/wago">Ara/Wago</option>
                                                <option value="single_fuse">Single Fuse</option>
                                                <option value="dual_fuse">Dual Fuse</option>
                                            </select>
                                        </td>
                                        <td><input type="text" class="form-control input-left-input" value="${inputLeft}" placeholder="Enter input left"></td>
                                        <td><input type="text" class="form-control input-right-input" value="" placeholder="Enter input right"></td>
                                        <!-- Spare field with arrow controls -->
                                        <td>
                                            <div class="d-flex align-items-center justify-content-center">
                                                <input type="hidden" class="form-control spare-input" value="N">
                                                <button type="button" class="btn btn-sm btn-outline-primary spare-up-btn" style="border-radius: 4px 0 0 4px;">
                                                    <i class="bi bi-arrow-up"></i>
                                                </button>
                                                <span class="spare-display mx-2" style="min-width: 20px; text-align: center;">N</span>
                                                <button type="button" class="btn btn-sm btn-outline-primary spare-down-btn" style="border-radius: 0 4px 4px 0;">
                                                    <i class="bi bi-arrow-down"></i>
                                                </button>
                                            </div>
                                        </td>
                                        <!-- Input Connected field with arrow controls -->
                                        <td>
                                            <div class="d-flex align-items-center justify-content-center">
                                                <input type="hidden" class="form-control input-connected-input" value="Y">
                                                <button type="button" class="btn btn-sm btn-outline-primary input-connected-up-btn" style="border-radius: 4px 0 0 4px;">
                                                    <i class="bi bi-arrow-up"></i>
                                                </button>
                                                <span class="input-connected-display mx-2" style="min-width: 20px; text-align: center;">Y</span>
                                                <button type="button" class="btn btn-sm btn-outline-primary input-connected-down-btn" style="border-radius: 0 4px 4px 0;">
                                                    <i class="bi bi-arrow-down"></i>
                                                </button>
                                            </div>
                                        </td>
                                        <!-- Output Connected field with arrow controls -->
                                        <td>
                                            <div class="d-flex align-items-center justify-content-center">
                                                <input type="hidden" class="form-control output-connected-input" value="Y">
                                                <button type="button" class="btn btn-sm btn-outline-primary output-connected-up-btn" style="border-radius: 4px 0 0 4px;">
                                                    <i class="bi bi-arrow-up"></i>
                                                </button>
                                                <span class="output-connected-display mx-2" style="min-width: 20px; text-align: center;">Y</span>
                                                <button type="button" class="btn btn-sm btn-outline-primary output-connected-down-btn" style="border-radius: 0 4px 4px 0;">
                                                    <i class="bi bi-arrow-down"></i>
                                                </button>
                                            </div>
                                        </td>
                                        <td><input type="text" class="form-control output-left-input" value="" placeholder="Enter output left"></td>
                                        <td><input type="text" class="form-control output-right-input" value="" placeholder="Enter output right"></td>
                                    `;
                                    tableBody.appendChild(row);
                                }
                                
                                // Add event listeners for arrow buttons
                                addTerminalArrowEventListeners();
                                
                                terminalModal.show();
                            } else {
                                document.getElementById('errorMessage').textContent = result.message || 'Failed to save cable.';
                                errorModal.show();
                            }
                        })
                        .catch(error => {
                            button.disabled = false;
                            button.innerHTML = '<i class="bi bi-gear"></i> Config Term';
                            document.getElementById('errorMessage').textContent = error.message || 'An error occurred while saving the cable.';
                            console.log(error);
                            errorModal.show();
                        });
                    });
                });
                cableModal.show();
            })
            .catch(error => {
                console.error('Error fetching existing cables:', error);
                console.log('Using default cable ID starting from 1');
            });
    }

    // Function to update cable name based on row, start_no and terminal
    function updateCableName(index) {
        var cable = currentCables[index];
        if (!cable) return;
        
        var row = cable.row;
        var startNo = parseInt(cable.start_no) || 1;
        var terminal = parseInt(cable.terminal) || 12;
        var endNo = startNo + terminal - 1;
        var cableName = row + ' T' + startNo + '-' + endNo;
        
        // Update cable object
        cable.cable_name = cableName;
        
        // Update the input field
        var cableNameInput = document.querySelector(`.cable-name-input[data-index="${index}"]`);
        if (cableNameInput) {
            cableNameInput.value = cableName;
        }
    }

    // Function to add event listeners for terminal arrow buttons
    function addTerminalArrowEventListeners() {
        // Spare field
        document.querySelectorAll('.spare-up-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                var row = this.closest('tr');
                var hiddenInput = row.querySelector('.spare-input');
                var displaySpan = row.querySelector('.spare-display');
                hiddenInput.value = 'Y';
                displaySpan.textContent = 'Y';
                displaySpan.className = 'spare-display mx-2 text-success fw-bold';
            });
        });
        
        document.querySelectorAll('.spare-down-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                var row = this.closest('tr');
                var hiddenInput = row.querySelector('.spare-input');
                var displaySpan = row.querySelector('.spare-display');
                hiddenInput.value = 'N';
                displaySpan.textContent = 'N';
                displaySpan.className = 'spare-display mx-2 text-danger fw-bold';
            });
        });
        
        // Input Connected field
        document.querySelectorAll('.input-connected-up-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                var row = this.closest('tr');
                var hiddenInput = row.querySelector('.input-connected-input');
                var displaySpan = row.querySelector('.input-connected-display');
                hiddenInput.value = 'Y';
                displaySpan.textContent = 'Y';
                displaySpan.className = 'input-connected-display mx-2 text-success fw-bold';
            });
        });
        
        document.querySelectorAll('.input-connected-down-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                var row = this.closest('tr');
                var hiddenInput = row.querySelector('.input-connected-input');
                var displaySpan = row.querySelector('.input-connected-display');
                hiddenInput.value = 'N';
                displaySpan.textContent = 'N';
                displaySpan.className = 'input-connected-display mx-2 text-danger fw-bold';
            });
        });
        
        // Output Connected field
        document.querySelectorAll('.output-connected-up-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                var row = this.closest('tr');
                var hiddenInput = row.querySelector('.output-connected-input');
                var displaySpan = row.querySelector('.output-connected-display');
                hiddenInput.value = 'Y';
                displaySpan.textContent = 'Y';
                displaySpan.className = 'output-connected-display mx-2 text-success fw-bold';
            });
        });
        
        document.querySelectorAll('.output-connected-down-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                var row = this.closest('tr');
                var hiddenInput = row.querySelector('.output-connected-input');
                var displaySpan = row.querySelector('.output-connected-display');
                hiddenInput.value = 'N';
                displaySpan.textContent = 'N';
                displaySpan.className = 'output-connected-display mx-2 text-danger fw-bold';
            });
        });
    }

    function saveAllCables() {
        var saveAllBtn = document.getElementById('saveAllCablesBtn');

        saveAllBtn.disabled = true;
        saveAllBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving All...';

        var promises = currentCables.map(cable => {
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

                var allSuccess = results.every(result => result.success);
                
                if (allSuccess) {
                    cableModal.hide();
                    
                    // Update the config button to show completed status
                    var configButton = document.querySelector(`.config-cables-btn[data-junction-index="${currentJunctionIndex}"]`);
                    var cableStatus = document.querySelector(`.cable-status[data-junction-index="${currentJunctionIndex}"]`);
                    if (configButton && cableStatus) {
                        configButton.style.display = 'none';
                        cableStatus.style.display = 'inline';
                    }

                    // Enable next step button
                    document.getElementById('nextStepBtn').disabled = false;

                    successModal.show();
                    setTimeout(() => { successModal.hide(); }, 1500);
                } else {
                    document.getElementById('errorMessage').textContent = 'Some cables failed to save. Please check and try again.';
                    console.log(results); // Log the results for debugging
                    errorModal.show();
                }
            })
            .catch(error => {
                saveAllBtn.disabled = false;
                saveAllBtn.innerHTML = '<i class="bi bi-check-circle"></i> Save All Cables';
                document.getElementById('errorMessage').textContent = error.message || 'An error occurred while saving cables.';
                console.log(error); // Log the error for debugging
                errorModal.show();
            });
    }

    function saveAllTerminals() {
        var saveBtn = document.getElementById('saveAllTerminalsBtn');
        saveBtn.disabled = true;
        saveBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving...';

        var cable = currentCables[currentCableIndex];
        var cableId = cable.cable_id;
        var terminals = [];

        document.querySelectorAll('#terminalTableBody tr').forEach((row, idx) => {
            terminals.push({
                cable_id: cableId,
                terminal_id: row.querySelector('.terminal-id-input').value,
                terminal_no: row.querySelector('.terminal-name-input').value, // CHANGED: terminal_name → terminal_no
                symbol: row.querySelector('.symbol-input').value,
                input_left: row.querySelector('.input-left-input').value,
                input_right: row.querySelector('.input-right-input').value,
                spare: row.querySelector('.spare-input').value,
                input_connected: row.querySelector('.input-connected-input').value,
                output_connected: row.querySelector('.output-connected-input').value,
                output_left: row.querySelector('.output-left-input').value,
                output_right: row.querySelector('.output-right-input').value
            });
        });

        var promises = terminals.map(terminal => {
            return fetch('/add_terminal_ajax', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify(terminal)
            }).then(response => {
                if (!response.ok) {
                    return response.text().then(text => {
                        throw new Error(`Failed to save terminal: ${response.status} - ${text}`);
                    });
                }
                return response.json();
            });
        });

        Promise.all(promises)
            .then(results => {
                saveBtn.disabled = false;
                saveBtn.innerHTML = '<i class="bi bi-check-circle"></i> Save All Terminals';
                if (results.every(r => r.success)) {
                    // Update cable status immediately to green check mark
                    updateCableStatusImmediately(currentCableIndex);
                    
                    // Also update by server check (as a backup)
                    updateCableStatusForSingleCable(cableId, currentCableIndex);
                    
                    terminalModal.hide();
                    successModal.show();
                    setTimeout(() => successModal.hide(), 1500);
                } else {
                    document.getElementById('errorMessage').textContent = 'Some terminals failed to save.';
                    console.log(results); // Log the results for debugging
                    errorModal.show();
                }
            })
            .catch(error => {
                saveBtn.disabled = false;
                saveBtn.innerHTML = '<i class="bi bi-check-circle"></i> Save All Terminals';
                document.getElementById('errorMessage').textContent = error.message || 'An error occurred.';
                console.log(error); // Log the error for debugging
                errorModal.show();
            });
    }

    // Function to immediately update cable status to green
    function updateCableStatusImmediately(index) {
        var statusElement = document.getElementById(`status-${index}`);
        if (statusElement) {
            statusElement.innerHTML = '<i class="bi bi-check-circle-fill text-success" title="Terminals configured"></i>';
        }
    }

    // Function to update cable status for a single cable by checking the server
    function updateCableStatusForSingleCable(cableId, index) {
        fetch(`/check_terminals_for_cable?cable_id=${cableId}`)
            .then(response => {
                if (!response.ok) {
                    // If endpoint returns 404, assume terminals exist (since we just saved them)
                    console.log(`Server check failed for cable ${cableId}, assuming terminals exist`);
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
                // Even if there's an error, we've already updated the status immediately
            });
    }

    // Function to check terminal status for all cables
    function updateCableStatuses() {
        currentCables.forEach((cable, index) => {
            fetch(`/check_terminals_for_cable?cable_id=${cable.cable_id}`)
                .then(response => {
                    if (!response.ok) {
                        // If endpoint returns 404, skip this cable
                        console.log(`Server check failed for cable ${cable.cable_id}, skipping`);
                        return;
                    }
                    return response.json();
                })
                .then(data => {
                    if (!data) return; // Skip if no data due to error
                    
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
                    // Don't update status on error - keep existing status
                });
        });
    }

    // Header Configuration Functions
    function openHeaderConfigModal() {
        var cable = currentCables[currentCableIndex];
        document.getElementById('headerCableId').value = cable.cable_id;
        
        // Clear existing header rows
        currentHeaders = [];
        var tableBody = document.getElementById('headerTableBody');
        tableBody.innerHTML = '';
        
        // Show loading state
        tableBody.innerHTML = '<tr><td colspan="6" class="text-center"><div class="spinner-border spinner-border-sm" role="status"></div> Loading terminal data...</td></tr>';
        
        // Fetch terminal data for this cable to get actual terminal names
        fetch(`/get_terminals_for_cable?cable_id=${cable.cable_id}`)
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                return response.json();
            })
            .then(terminals => {
                console.log('Fetched terminals for cable:', terminals);
                
                // Clear loading state
                tableBody.innerHTML = '';
                
                if (terminals && terminals.length > 0) {
                    // Extract terminal names and ensure they are numbers
                    var terminalNames = terminals
                        .map(terminal => {
                            // Use terminal_name if available, otherwise fall back to terminal_id
                            const name = terminal.terminal_name;
                            if (name && !isNaN(name)) {
                                return parseInt(name);
                            }
                            return null;
                        })
                        .filter(name => name !== null && !isNaN(name));
                    
                    console.log('Processed terminal names:', terminalNames);
                    
                    if (terminalNames.length > 0) {
                        var minTerminal = Math.min(...terminalNames);
                        var maxTerminal = Math.max(...terminalNames);
                        console.log(`Calculated terminal range: ${minTerminal} to ${maxTerminal}`);
                        
                        // Add one initial header row with actual terminal range
                        addHeaderRowWithTerminals(minTerminal, maxTerminal);
                    } else {
                        console.log('No valid terminal numbers found, using manual entry');
                        // Add empty row for manual entry
                        addHeaderRow();
                    }
                } else {
                    console.log('No terminals found for this cable, using manual entry');
                    // Add empty row for manual entry
                    addHeaderRow();
                }
            })
            .catch(error => {
                console.error('Error fetching terminals:', error);
                // Clear loading state and show manual entry
                tableBody.innerHTML = '';
                console.log('Using manual entry due to fetch error');
                addHeaderRow();
            });
        
        // Show the header modal
        headerModal.show();
    }

    function addHeaderRowWithTerminals(terminalStart, terminalEnd) {
        var tableBody = document.getElementById('headerTableBody');
        var cable = currentCables[currentCableIndex];
        
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
        
        var row = document.createElement('tr');
        row.innerHTML = `
            <td>
                <select class="form-select header-type-input" data-index="${rowIndex}">
                    <option value="">Select Header Type</option>
                    <option value="WIREFROM">WIREFROM</option>
                    <option value="WIRETO">WIRETO</option>
                    <option value="RELAY">RELAY</option>
                </select>
            </td>
            <td>
                <input type="number" class="form-control header-terminal-start-input" data-index="${rowIndex}" value="${terminalStart}" placeholder="Start" min="1">
                <small class="form-text text-muted">Based on terminal data: ${terminalStart}-${terminalEnd}</small>
            </td>
            <td>
                <input type="number" class="form-control header-terminal-end-input" data-index="${rowIndex}" value="${terminalEnd}" placeholder="End" min="1">
                <small class="form-text text-muted">Based on terminal data: ${terminalStart}-${terminalEnd}</small>
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
        
        // Add event listeners for the new row
        addHeaderRowEventListeners(rowIndex);
    }

    function addHeaderRow() {
        var tableBody = document.getElementById('headerTableBody');
        var cable = currentCables[currentCableIndex];
        
        // Try to calculate from cable data as fallback
        var terminalStart = 1;
        var terminalEnd = 12;
        
        // If we have start_no and terminal count, use that
        if (cable.start_no && cable.terminal) {
            terminalStart = parseInt(cable.start_no);
            terminalEnd = terminalStart + parseInt(cable.terminal) - 1;
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
        
        var row = document.createElement('tr');
        row.innerHTML = `
            <td>
                <select class="form-select header-type-input" data-index="${rowIndex}">
                    <option value="">Select Header Type</option>
                    <option value="WIREFROM">WIREFROM</option>
                    <option value="WIRETO">WIRETO</option>
                    <option value="RELAY">RELAY</option>
                </select>
            </td>
            <td>
                <input type="number" class="form-control header-terminal-start-input" data-index="${rowIndex}" value="${terminalStart}" placeholder="Start" min="1">
                <small class="form-text text-muted">Enter start terminal number</small>
            </td>
            <td>
                <input type="number" class="form-control header-terminal-end-input" data-index="${rowIndex}" value="${terminalEnd}" placeholder="End" min="1">
                <small class="form-text text-muted">Enter end terminal number</small>
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
        
        // Add event listeners for the new row
        addHeaderRowEventListeners(rowIndex);
    }

    function addHeaderRowEventListeners(rowIndex) {
        document.querySelector(`.header-type-input[data-index="${rowIndex}"]`).addEventListener('change', function() {
            currentHeaders[rowIndex].header_type = this.value;
        });
        
        document.querySelector(`.header-terminal-start-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentHeaders[rowIndex].terminal_start = this.value;
        });
        
        document.querySelector(`.header-terminal-end-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentHeaders[rowIndex].terminal_end = this.value;
        });
        
        document.querySelector(`.header-input-output-input[data-index="${rowIndex}"]`).addEventListener('change', function() {
            currentHeaders[rowIndex].input_output = this.value;
        });
        
        document.querySelector(`.header-text-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentHeaders[rowIndex].text = this.value;
        });
        
        document.querySelector(`.remove-header-row-btn[data-index="${rowIndex}"]`).addEventListener('click', function() {
            removeHeaderRow(rowIndex);
        });
    }

    function removeHeaderRow(rowIndex) {
        currentHeaders.splice(rowIndex, 1);
        refreshHeaderTable();
    }

    function refreshHeaderTable() {
        var tableBody = document.getElementById('headerTableBody');
        tableBody.innerHTML = '';
        
        currentHeaders.forEach((header, index) => {
            var row = document.createElement('tr');
            row.innerHTML = `
                <td>
                    <select class="form-select header-type-input" data-index="${index}">
                        <option value="">Select Header Type</option>
                        <option value="WIREFROM" ${header.header_type === 'WIREFROM' ? 'selected' : ''}>WIREFROM</option>
                        <option value="WIRETO" ${header.header_type === 'WIRETO' ? 'selected' : ''}>WIRETO</option>
                        <option value="RELAY" ${header.header_type === 'RELAY' ? 'selected' : ''}>RELAY</option>
                    </select>
                </td>
                <td>
                    <input type="number" class="form-control header-terminal-start-input" data-index="${index}" value="${header.terminal_start || ''}" placeholder="Start" min="1">
                </td>
                <td>
                    <input type="number" class="form-control header-terminal-end-input" data-index="${index}" value="${header.terminal_end || ''}" placeholder="End" min="1">
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

    function saveAllHeaders() {
        var saveBtn = document.getElementById('saveAllHeadersBtn');
        saveBtn.disabled = true;
        saveBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving...';

        var promises = currentHeaders.map(header => {
            return fetch('/add_header_ajax', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify(header)
            }).then(response => {
                if (!response.ok) {
                    return response.text().then(text => {
                        throw new Error(`Failed to save header: ${response.status} - ${text}`);
                    });
                }
                return response.json();
            });
        });

        Promise.all(promises)
            .then(results => {
                saveBtn.disabled = false;
                saveBtn.innerHTML = '<i class="bi bi-check-circle"></i> Save All Headers';
                if (results.every(r => r.success)) {
                    headerModal.hide();
                    successModal.show();
                    setTimeout(() => successModal.hide(), 1500);
                } else {
                    document.getElementById('errorMessage').textContent = 'Some headers failed to save.';
                    errorModal.show();
                }
            })
            .catch(error => {
                saveBtn.disabled = false;
                saveBtn.innerHTML = '<i class="bi bi-check-circle"></i> Save All Headers';
                document.getElementById('errorMessage').textContent = error.message || 'An error occurred while saving headers.';
                errorModal.show();
            });
    }

    // Group Configuration Functions
    function openGroupConfigModal() {
        var cable = currentCables[currentCableIndex];
        document.getElementById('groupCableId').value = cable.cable_id;
        
        // Clear existing group rows
        currentGroups = [];
        var tableBody = document.getElementById('groupTableBody');
        tableBody.innerHTML = '';
        
        // Fetch existing groups for this cable to determine next group ID
        fetch(`/get_groups_for_cable?cable_id=${cable.cable_id}`)
            .then(response => response.json())
            .then(existingGroups => {
                let nextGroupId = 1;
                if (existingGroups && existingGroups.length > 0) {
                    // Find the highest existing group ID and add 1
                    const maxId = Math.max(...existingGroups.map(group => parseInt(group.group_id) || 0));
                    nextGroupId = maxId + 1;
                }
                
                // Set the initial group ID for new rows
                window.nextGroupId = nextGroupId;
                
                // Add one initial group row
                addGroupRow();
            })
            .catch(error => {
                console.error('Error fetching existing groups:', error);
                // Start from 1 if there's an error
                window.nextGroupId = 1;
                addGroupRow();
            });
        
        // Show the group modal
        groupModal.show();
    }

    function addGroupRow() {
        var tableBody = document.getElementById('groupTableBody');
        var cable = currentCables[currentCableIndex];
        
        var group = {
            cable_id: cable.cable_id,
            group_id: '',
            terminal_no: '',
            input_output: '',
            text: ''
        };
        
        var rowIndex = currentGroups.length;
        currentGroups.push(group);
        
        var row = document.createElement('tr');
        row.innerHTML = `
            <td>
                <input type="text" class="form-control group-id-input" data-index="${rowIndex}" placeholder="Group ID">
            </td>
            <td>
                <input type="text" class="form-control group-terminal-no-input" data-index="${rowIndex}" placeholder="(e.g., 1,2,4,54)">
            </td>
            <td>
                <select class="form-select group-input-output-input" data-index="${rowIndex}">
                    <option value="">Select I/O</option>
                    <option value="Input">Input</option>
                    <option value="Output">Output</option>
                </select>
            </td>
            <td>
                <input type="text" class="form-control group-text-input" data-index="${rowIndex}" placeholder="Group text">
            </td>
            <td>
                <button type="button" class="btn btn-sm btn-danger remove-group-row-btn" data-index="${rowIndex}">
                    <i class="bi bi-trash"></i>
                </button>
            </td>
        `;
        tableBody.appendChild(row);
        
        // Add event listeners for the new row
        addGroupRowEventListeners(rowIndex);
    }

    function addGroupRowEventListeners(rowIndex) {
        document.querySelector(`.group-id-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentGroups[rowIndex].group_id = this.value;
        });
        
        document.querySelector(`.group-terminal-no-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentGroups[rowIndex].terminal_no = this.value;
        });
        
        document.querySelector(`.group-input-output-input[data-index="${rowIndex}"]`).addEventListener('change', function() {
            currentGroups[rowIndex].input_output = this.value;
        });
        
        document.querySelector(`.group-text-input[data-index="${rowIndex}"]`).addEventListener('input', function() {
            currentGroups[rowIndex].text = this.value;
        });
        
        document.querySelector(`.remove-group-row-btn[data-index="${rowIndex}"]`).addEventListener('click', function() {
            removeGroupRow(rowIndex);
        });
    }

    function removeGroupRow(rowIndex) {
        currentGroups.splice(rowIndex, 1);
        refreshGroupTable();
    }

    function refreshGroupTable() {
        var tableBody = document.getElementById('groupTableBody');
        tableBody.innerHTML = '';
        
        currentGroups.forEach((group, index) => {
            var row = document.createElement('tr');
            row.innerHTML = `
                <td>
                    <input type="text" class="form-control group-id-input" data-index="${index}" value="${group.group_id || ''}" placeholder="Group ID">
                </td>
                <td>
                    <input type="text" class="form-control group-terminal-no-input" data-index="${index}" value="${group.terminal_no || ''}" placeholder="(e.g., 1,2,4,54)">
                </td>
                <td>
                    <select class="form-select group-input-output-input" data-index="${index}">
                        <option value="">Select I/O</option>
                        <option value="Input" ${group.input_output === 'Input' ? 'selected' : ''}>Input</option>
                        <option value="Output" ${group.input_output === 'Output' ? 'selected' : ''}>Output</option>
                        <option value="Both" ${group.input_output === 'Both' ? 'selected' : ''}>Both</option>
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

    function saveAllGroups() {
        var saveBtn = document.getElementById('saveAllGroupsBtn');
        saveBtn.disabled = true;
        saveBtn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Saving...';

        var promises = currentGroups.map(group => {
            return fetch('/add_group_ajax', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify(group)
            }).then(response => {
                if (!response.ok) {
                    return response.text().then(text => {
                        throw new Error(`Failed to save group: ${response.status} - ${text}`);
                    });
                }
                return response.json();
            });
        });

        Promise.all(promises)
            .then(results => {
                saveBtn.disabled = false;
                saveBtn.innerHTML = '<i class="bi bi-check-circle"></i> Save All Groups';
                if (results.every(r => r.success)) {
                    groupModal.hide();
                    successModal.show();
                    setTimeout(() => successModal.hide(), 1500);
                } else {
                    document.getElementById('errorMessage').textContent = 'Some groups failed to save.';
                    errorModal.show();
                }
            })
            .catch(error => {
                saveBtn.disabled = false;
                saveBtn.innerHTML = '<i class="bi bi-check-circle"></i> Save All Groups';
                document.getElementById('errorMessage').textContent = error.message || 'An error occurred while saving groups.';
                errorModal.show();
            });
    }

    document.getElementById('junctionGridForm')?.addEventListener('submit', function(e) {
        if (e.submitter?.value !== 'add_junctions') return;

        e.preventDefault();

        // Format all location names before submission
        for (var i = 0; i < requestedJunctions; i++) {
            var nameInput = document.querySelector(`input[name="junctions[${i}][junction_name]"]`);
            var sizeSelect = document.querySelector(`select[name="junctions[${i}][junction_size]"]`);
            if (nameInput && sizeSelect) {
                var baseName = nameInput.value;
                var size = sizeSelect.value;
                var formattedName = formatLocationName(baseName, size);
                nameInput.value = formattedName;
            }
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

            // Check location name
            if (!junction.junction_name || junction.junction_name.trim() === '') {
                isValidRow = false;
            }

            // Check location size
            if (!['Full', 'Half'].includes(junction.junction_size)) {
                isValidRow = false;
            }

            // Check location row
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
            document.getElementById('errorMessage').textContent = 'Please enter a location name, select a valid size, and enter a Location Row count for at least one location.';
            errorModal.show();
            return;
        }

        var submitBtn = document.getElementById('saveJunctionsBtn');
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
            submitBtn.innerHTML = 'Save Locations';

            if (result.success && result.junctions && result.junctions.length > 0) {
                document.querySelector('.table-responsive').style.display = 'none';
                successModal.show();
                setTimeout(() => { 
                    successModal.hide();
                    // NAVIGATE TO HEADERS STEP (STEP 5) AFTER SAVING LOCATIONS
                    window.location.href = '{{ url_for("main.workflow_step", step=5) }}?show_header_modal=true';
                }, 1000);
            } else {
                document.getElementById('errorMessage').textContent = result.message || 'Failed to add locations.';
                errorModal.show();
            }
        })
        .catch(error => {
            submitBtn.disabled = false;
            submitBtn.innerHTML = 'Save Locations';
            document.getElementById('errorMessage').textContent = error.message || 'An error occurred while adding locations.';
            errorModal.show();
        });
    });

    function showFinalCompletion() {
        document.getElementById('allCompletionMessage').style.display = 'block';
        document.getElementById('addMoreAfterCompletion').style.display = 'block';
        document.getElementById('junctionGridForm').style.display = 'none';
    }
})();