/**
 * KUHL (V19) Processing Logic
 * Validates Fabric/Yarn rows in Buyer CBD files
 *
 * Validation Rules:
 * - Scan Column A for "Fabric/Yarn"
 * - When Column A contains "Fabric/Yarn", check if Column K is 5%
 * - Column A displayed as "Type"
 * - Column K displayed as "Consumption"
 */

class KuhlProcessor {
    constructor() {
        this.bcbdResults = [];
        this.validationRules = {
            consumption: {
                column: 'K',
                columnIndex: 10,
                label: 'Consumption',
                expectedValue: 0.05,
                expectedDisplay: '5%'
            }
        };
    }

    /**
     * Initialize - Display validation rules in the OB drop zone
     */
    initialize() {
        this.displayValidationRules();
        console.log('KUHL Processor initialized');
    }

    /**
     * Display validation rules in the OB drop zone
     */
    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v19');
        if (!obDropZone) return;

        let contentHTML = `
            <div class="burton-cost-container">
                <div class="burton-cost-header">

                </div>
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">Fabric/Yarn Section:</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column K - Consumption: <strong>5%</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">Trim Section:</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column K - Consumption: <strong>3%</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column E - Supplier: <strong>Contains "Local" or "Nominated"</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column H - C.I.F.VS FOB %: <strong>0.012%</strong> (if Local) or <strong>15%</strong> (if Nominated)</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">Labelling Section:</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column K - Consumption: <strong>3%</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column E - Supplier: <strong>Contains "Local" or "Nominated"</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column H - C.I.F.VS FOB %: <strong>0.012%</strong> (if Local) or <strong>15%</strong> (if Nominated)</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">Profit Margin:</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column M - Value: <strong>0.60 to 0.95</strong></div>
                    </div>
                </div>
            </div>
        `;

        obDropZone.innerHTML = contentHTML;
    }

    /**
     * Process files and generate results
     */
    async processFiles(bcbdFiles) {
        this.bcbdResults = [];

        try {
            if (!bcbdFiles || bcbdFiles.length === 0) {
                return this.generateErrorHTML('Please upload Buyer CBD files');
            }

            // Process each BCBD file
            for (const file of bcbdFiles) {
                const fileData = await this.parseBuyerCBDFile(file);
                const validationResults = this.validateFile(fileData);
                this.bcbdResults.push({
                    fileName: file.name,
                    results: validationResults
                });
            }

            return this.generateResultsHTML(this.bcbdResults);

        } catch (error) {
            console.error('Error processing files:', error);
            return this.generateErrorHTML(error.message);
        }
    }

    /**
     * Parse Buyer CBD Excel file
     */
    async parseBuyerCBDFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    // Get the first sheet
                    const firstSheetName = workbook.SheetNames[0];
                    const sheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

                    resolve(jsonData);
                } catch (error) {
                    reject(error);
                }
            };

            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Validate a percentage value against expected
     */
    validatePercentage(actualValue, expectedValue) {
        if (actualValue === undefined || actualValue === null || actualValue === '') {
            return { isValid: false, displayValue: 'Empty', isEmpty: true };
        }

        let numericValue;
        let displayValue = actualValue.toString();

        if (typeof actualValue === 'number') {
            numericValue = actualValue;
        } else {
            const strValue = actualValue.toString().trim();
            if (strValue.endsWith('%')) {
                numericValue = parseFloat(strValue) / 100;
            } else {
                numericValue = parseFloat(strValue);
            }
        }

        if (isNaN(numericValue)) {
            return { isValid: false, displayValue: displayValue, isEmpty: false };
        }

        // Check if value matches expected (with tolerance)
        const isValid = Math.abs(numericValue - expectedValue) < 0.0001;

        // Display as percentage
        if (numericValue < 1) {
            // For small percentages like 0.012%, show more decimals
            if (numericValue < 0.01) {
                displayValue = (numericValue * 100).toFixed(3) + '%';
            } else {
                displayValue = (numericValue * 100).toFixed(0) + '%';
            }
        } else {
            displayValue = numericValue.toFixed(0) + '%';
        }

        return { isValid, displayValue, isEmpty: false };
    }

    /**
     * Validate a consumption value (5% for Fabric/Yarn)
     */
    validateConsumption(actualValue) {
        return this.validatePercentage(actualValue, 0.05);
    }

    /**
     * Validate a consumption value (3% for Trim)
     */
    validateTrimConsumption(actualValue) {
        return this.validatePercentage(actualValue, 0.03);
    }

    /**
     * Validate C.I.F.VS FOB % (0.012% for Local, 0.15 for Nominated)
     */
    validateCifVsFob(actualValue, isNominated = false) {
        const expectedValue = isNominated ? 0.15 : 0.00012; // 0.15 (15%) for Nominated, 0.012% for Local
        return this.validatePercentage(actualValue, expectedValue);
    }

    /**
     * Validate Supplier contains "Local" or "Nominated"
     */
    validateSupplierLocal(actualValue) {
        if (actualValue === undefined || actualValue === null || actualValue === '') {
            return { isValid: false, displayValue: 'Empty', isEmpty: true, isNominated: false };
        }

        const strValue = actualValue.toString().trim();
        const isLocal = strValue.toLowerCase().includes('local');
        const isNominated = strValue.toLowerCase().includes('nominated');
        const isValid = isLocal || isNominated;

        return { isValid, displayValue: strValue, isEmpty: false, isNominated };
    }

    /**
     * Check if a value matches "Fabric/Yarn" pattern
     * Handles variations like "Fabric/Yarn", "Fabric / Yarn", "fabric/yarn", etc.
     */
    isFabricYarn(value) {
        if (!value) return false;

        const normalized = value.toString().toLowerCase().trim();

        // Check for exact or near-exact matches
        if (normalized.includes('fabric/yarn')) return true;
        if (normalized.includes('fabric / yarn')) return true;
        if (normalized.includes('fabric/ yarn')) return true;
        if (normalized.includes('fabric /yarn')) return true;

        // Check if it contains both "fabric" and "yarn" (in any format)
        if (normalized.includes('fabric') && normalized.includes('yarn')) return true;

        return false;
    }

    /**
     * Check if a value matches "Trim" pattern
     */
    isTrim(value) {
        if (!value) return false;

        const normalized = value.toString().toLowerCase().trim();

        if (normalized.includes('trim')) return true;

        return false;
    }

    /**
     * Check if a value matches "Labelling" pattern
     */
    isLabelling(value) {
        if (!value) return false;

        const normalized = value.toString().toLowerCase().trim();

        if (normalized.includes('label')) return true;

        return false;
    }

    /**
     * Check if a value matches "PROFIT MARGIN(%)" pattern
     */
    isProfitMargin(value) {
        if (!value) return false;

        const normalized = value.toString().toLowerCase().trim();

        if (normalized.includes('profit margin')) return true;
        if (normalized.includes('profit margin(%)')) return true;
        if (normalized.includes('profit margin (%)')) return true;

        return false;
    }

    /**
     * Validate Profit Margin value (should be between 0.60 and 0.95)
     */
    validateProfitMargin(actualValue) {
        if (actualValue === undefined || actualValue === null || actualValue === '') {
            return { isValid: false, displayValue: 'Empty', isEmpty: true };
        }

        let numericValue;
        let displayValue = actualValue.toString();

        if (typeof actualValue === 'number') {
            numericValue = actualValue;
        } else {
            const strValue = actualValue.toString().trim();
            // Remove % sign if present
            if (strValue.endsWith('%')) {
                numericValue = parseFloat(strValue) / 100;
            } else {
                numericValue = parseFloat(strValue);
            }
        }

        if (isNaN(numericValue)) {
            return { isValid: false, displayValue: displayValue, isEmpty: false };
        }

        // Check if value is between 0.60 and 0.95
        const isValid = numericValue >= 0.60 && numericValue <= 0.95;

        // Display as decimal (e.g., 0.75)
        displayValue = numericValue.toFixed(2);

        return { isValid, displayValue, isEmpty: false };
    }

    /**
     * Check if a value indicates a new section (ends Fabric/Yarn section)
     */
    isFabricYarnEndSection(value) {
        if (!value) return false;

        const normalized = value.toString().toLowerCase().trim();

        // Sections that would end the Fabric/Yarn section
        const sectionHeaders = [
            'trim', 'trims', 'accessories', 'accessory', 'packaging',
            'label', 'labels', 'labelling', 'thread', 'sundries', 'total', 'subtotal',
            'sub-total', 'sub total', 'hardware', 'zipper', 'button'
        ];

        for (const header of sectionHeaders) {
            if (normalized.includes(header)) return true;
        }

        return false;
    }

    /**
     * Check if a value indicates a new section (ends Trim section)
     */
    isTrimEndSection(value) {
        if (!value) return false;

        const normalized = value.toString().toLowerCase().trim();

        // Sections that would end the Trim section
        const sectionHeaders = [
            'packaging', 'total', 'subtotal', 'sub-total', 'sub total',
            'labor', 'labour', 'cm', 'cutting', 'sewing', 'finishing',
            'label', 'labelling'
        ];

        for (const header of sectionHeaders) {
            if (normalized.includes(header)) return true;
        }

        return false;
    }

    /**
     * Check if a value indicates a new section (ends Labelling section)
     */
    isLabellingEndSection(value) {
        if (!value) return false;

        const normalized = value.toString().toLowerCase().trim();

        // Sections that would end the Labelling section
        const sectionHeaders = [
            'packaging', 'total', 'subtotal', 'sub-total', 'sub total',
            'labor', 'labour', 'cm', 'cutting', 'sewing', 'finishing'
        ];

        for (const header of sectionHeaders) {
            if (normalized.includes(header)) return true;
        }

        return false;
    }

    /**
     * Validate file - find all Fabric/Yarn, Trim, and Labelling rows
     * Handles merged cells by tracking when we're inside sections
     * Returns cell references for valid and invalid cells
     */
    validateFile(jsonData) {
        const results = {
            // Fabric/Yarn results (Column K = 5%)
            fabricYarn: {
                validCells: [],
                invalidCells: []
            },
            // Trim results
            trim: {
                consumption: { validCells: [], invalidCells: [] },
                supplier: { validCells: [], invalidCells: [] },
                cifVsFob: { validCells: [], invalidCells: [] }
            },
            // Labelling results (same rules as Trim)
            labelling: {
                consumption: { validCells: [], invalidCells: [] },
                supplier: { validCells: [], invalidCells: [] },
                cifVsFob: { validCells: [], invalidCells: [] }
            },
            // Profit Margin results (Column M = 0.60-0.95)
            profitMargin: {
                validCells: [],
                invalidCells: []
            }
        };

        const colA = 0;  // Column A index
        const colE = 4;  // Column E index (Supplier)
        const colH = 7;  // Column H index (C.I.F.VS FOB %)
        const colK = 10; // Column K index (Consumption)
        const colM = 12; // Column M index (Profit Margin)

        console.log(`Scanning ${jsonData.length} rows...`);

        let inFabricYarnSection = false;
        let inTrimSection = false;
        let inLabellingSection = false;

        // Scan through all rows
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;

            const typeValue = row[colA] ? row[colA].toString().trim() : '';

            // Check if we're entering a Fabric/Yarn section
            if (this.isFabricYarn(typeValue)) {
                inFabricYarnSection = true;
                inTrimSection = false;
                inLabellingSection = false;
                console.log(`Row ${i + 1}: Entering Fabric/Yarn section`);

                this.validateFabricYarnRow(row, i, colK, results.fabricYarn);
                continue;
            }

            // Check if we're entering a Trim section
            if (this.isTrim(typeValue) && !this.isFabricYarn(typeValue) && !this.isLabelling(typeValue)) {
                inTrimSection = true;
                inFabricYarnSection = false;
                inLabellingSection = false;
                console.log(`Row ${i + 1}: Entering Trim section`);

                this.validateTrimRow(row, i, colE, colH, colK, results.trim);
                continue;
            }

            // Check if we're entering a Labelling section
            if (this.isLabelling(typeValue) && !this.isFabricYarn(typeValue)) {
                inLabellingSection = true;
                inFabricYarnSection = false;
                inTrimSection = false;
                console.log(`Row ${i + 1}: Entering Labelling section`);

                this.validateTrimRow(row, i, colE, colH, colK, results.labelling);
                continue;
            }

            // Check if we're exiting the Fabric/Yarn section
            if (inFabricYarnSection && typeValue && this.isFabricYarnEndSection(typeValue)) {
                inFabricYarnSection = false;
                console.log(`Row ${i + 1}: Exiting Fabric/Yarn section (found: "${typeValue}")`);

                // Check if this is also start of another section
                if (this.isTrim(typeValue) && !this.isLabelling(typeValue)) {
                    inTrimSection = true;
                    console.log(`Row ${i + 1}: Entering Trim section`);
                    this.validateTrimRow(row, i, colE, colH, colK, results.trim);
                } else if (this.isLabelling(typeValue)) {
                    inLabellingSection = true;
                    console.log(`Row ${i + 1}: Entering Labelling section`);
                    this.validateTrimRow(row, i, colE, colH, colK, results.labelling);
                }
                continue;
            }

            // Check if we're exiting the Trim section
            if (inTrimSection && typeValue && this.isTrimEndSection(typeValue)) {
                inTrimSection = false;
                console.log(`Row ${i + 1}: Exiting Trim section (found: "${typeValue}")`);

                // Check if entering Labelling section
                if (this.isLabelling(typeValue)) {
                    inLabellingSection = true;
                    console.log(`Row ${i + 1}: Entering Labelling section`);
                    this.validateTrimRow(row, i, colE, colH, colK, results.labelling);
                }
                continue;
            }

            // Check if we're exiting the Labelling section
            if (inLabellingSection && typeValue && this.isLabellingEndSection(typeValue)) {
                inLabellingSection = false;
                console.log(`Row ${i + 1}: Exiting Labelling section (found: "${typeValue}")`);
                continue;
            }

            // Validate rows in their respective sections
            if (inFabricYarnSection) {
                this.validateFabricYarnRow(row, i, colK, results.fabricYarn);
            }
            if (inTrimSection) {
                this.validateTrimRow(row, i, colE, colH, colK, results.trim);
            }
            if (inLabellingSection) {
                this.validateTrimRow(row, i, colE, colH, colK, results.labelling);
            }

            // Check for PROFIT MARGIN(%) - single row lookup
            if (this.isProfitMargin(typeValue)) {
                console.log(`Row ${i + 1}: Found PROFIT MARGIN(%)`);
                this.validateProfitMarginRow(row, i, colM, results.profitMargin);
            }
        }

        console.log(`Fabric/Yarn: ${results.fabricYarn.validCells.length} valid, ${results.fabricYarn.invalidCells.length} invalid`);
        console.log(`Trim: ${results.trim.consumption.validCells.length + results.trim.consumption.invalidCells.length} cells`);
        console.log(`Labelling: ${results.labelling.consumption.validCells.length + results.labelling.consumption.invalidCells.length} cells`);
        console.log(`Profit Margin: ${results.profitMargin.validCells.length} valid, ${results.profitMargin.invalidCells.length} invalid`);

        return results;
    }

    /**
     * Validate a Fabric/Yarn row (Column K = 5%)
     */
    validateFabricYarnRow(row, rowIndex, colK, fabricYarnResults) {
        const consumptionValue = row[colK];

        if (consumptionValue === undefined || consumptionValue === null || consumptionValue === '') {
            return;
        }

        const validation = this.validateConsumption(consumptionValue);
        const cellRef = `K${rowIndex + 1}`;

        console.log(`Fabric/Yarn ${cellRef}: Value="${consumptionValue}" -> ${validation.isValid ? 'VALID' : 'INVALID'}`);

        if (validation.isValid) {
            fabricYarnResults.validCells.push(cellRef);
        } else {
            fabricYarnResults.invalidCells.push({
                cell: cellRef,
                value: validation.displayValue,
                expected: '5%'
            });
        }
    }

    /**
     * Validate a Trim row (Column K = 3%, Column E contains "Local" or "Nominated", Column H = 0.012% for Local or 0.15 for Nominated)
     */
    validateTrimRow(row, rowIndex, colE, colH, colK, trimResults) {
        // Validate Column K (Consumption = 3%)
        const consumptionValue = row[colK];
        if (consumptionValue !== undefined && consumptionValue !== null && consumptionValue !== '') {
            const validation = this.validateTrimConsumption(consumptionValue);
            const cellRef = `K${rowIndex + 1}`;

            console.log(`Trim Consumption ${cellRef}: Value="${consumptionValue}" -> ${validation.isValid ? 'VALID' : 'INVALID'}`);

            if (validation.isValid) {
                trimResults.consumption.validCells.push(cellRef);
            } else {
                trimResults.consumption.invalidCells.push({
                    cell: cellRef,
                    value: validation.displayValue,
                    expected: '3%'
                });
            }
        }

        // Validate Column E (Supplier contains "Local" or "Nominated")
        const supplierValue = row[colE];
        let isNominated = false;
        if (supplierValue !== undefined && supplierValue !== null && supplierValue !== '') {
            const validation = this.validateSupplierLocal(supplierValue);
            const cellRef = `E${rowIndex + 1}`;
            isNominated = validation.isNominated;

            console.log(`Trim Supplier ${cellRef}: Value="${supplierValue}" -> ${validation.isValid ? 'VALID' : 'INVALID'} (isNominated: ${isNominated})`);

            if (validation.isValid) {
                trimResults.supplier.validCells.push(cellRef);
            } else {
                trimResults.supplier.invalidCells.push({
                    cell: cellRef,
                    value: validation.displayValue,
                    expected: 'Contains "Local" or "Nominated"'
                });
            }
        }

        // Validate Column H (C.I.F.VS FOB % = 0.012% for Local, 0.15 for Nominated)
        const cifVsFobValue = row[colH];
        if (cifVsFobValue !== undefined && cifVsFobValue !== null && cifVsFobValue !== '') {
            const validation = this.validateCifVsFob(cifVsFobValue, isNominated);
            const cellRef = `H${rowIndex + 1}`;
            const expectedValue = isNominated ? '15%' : '0.012%';

            console.log(`Trim CIF vs FOB ${cellRef}: Value="${cifVsFobValue}" -> ${validation.isValid ? 'VALID' : 'INVALID'} (expected: ${expectedValue})`);

            if (validation.isValid) {
                trimResults.cifVsFob.validCells.push(cellRef);
            } else {
                trimResults.cifVsFob.invalidCells.push({
                    cell: cellRef,
                    value: validation.displayValue,
                    expected: expectedValue
                });
            }
        }
    }

    /**
     * Validate a Profit Margin row (Column M = 0.60-0.95)
     */
    validateProfitMarginRow(row, rowIndex, colM, profitMarginResults) {
        const profitMarginValue = row[colM];

        if (profitMarginValue === undefined || profitMarginValue === null || profitMarginValue === '') {
            return;
        }

        const validation = this.validateProfitMargin(profitMarginValue);
        const cellRef = `M${rowIndex + 1}`;

        console.log(`Profit Margin ${cellRef}: Value="${profitMarginValue}" -> ${validation.isValid ? 'VALID' : 'INVALID'}`);

        if (validation.isValid) {
            profitMarginResults.validCells.push(cellRef);
        } else {
            profitMarginResults.invalidCells.push({
                cell: cellRef,
                value: validation.displayValue,
                expected: '0.60-0.95'
            });
        }
    }

    /**
     * Format cells display - valid cells in green, invalid cells in red with expected value
     */
    formatAllCells(validCells, invalidCells) {
        let html = '';

        // Add valid cells (green) - inline comma separated
        if (validCells && validCells.length > 0) {
            const validParts = validCells.map(cell =>
                `<span style="color: #065f46; font-weight: 600;">${cell}</span>`
            );
            html += validParts.join(', ');
        }

        // Add invalid cells (red) - each on its own line with details
        if (invalidCells && invalidCells.length > 0) {
            if (html) html += '<br>';
            const invalidParts = invalidCells.map(c =>
                `<span style="color: #991b1b; font-weight: 600;">${c.cell}</span> <span style="font-size: 0.85em; color: #849bba;">(Actual: ${c.value}, Expected: ${c.expected})</span>`
            );
            html += invalidParts.join('<br>');
        }

        if (!html) return '-';
        return html;
    }

    /**
     * Generate HTML for results display
     * Shows rows for Fabric/Yarn, Trim, and Labelling with cell references
     */
    generateResultsHTML(results) {
        if (!results || results.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">KUHL Validation Ready</p>
                    <p>Upload Buyer CBD files to validate.</p>
                </div>
            `;
        }

        let html = '';

        // Add search bar and export button at the top
        html += `
            <div style="margin-bottom: 15px; display: flex; justify-content: space-between; align-items: center; gap: 12px;">
                <div class="search-container">
                    <div class="search-icon">
                        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path stroke-linecap="round" stroke-linejoin="round" d="M21 21l-5.197-5.197m0 0A7.5 7.5 0 105.196 5.196a7.5 7.5 0 0010.607 10.607z" />
                        </svg>
                    </div>
                    <input
                        type="text"
                        class="search-input-expandable"
                        placeholder="Search by filename..."
                        oninput="window.kuhlProcessor.searchByFilename(this.value)"
                    />
                </div>
                <button onclick="window.kuhlProcessor.exportToPDF()" class="export-btn">
                    Export
                </button>
            </div>
        `;

        for (const fileResult of results) {
            html += `<div class="file-result-group">`;

            const fabricYarn = fileResult.results.fabricYarn;
            const trim = fileResult.results.trim;
            const labelling = fileResult.results.labelling;
            const profitMargin = fileResult.results.profitMargin;

            // Calculate totals for summary
            const totalValid = fabricYarn.validCells.length +
                trim.consumption.validCells.length + trim.supplier.validCells.length + trim.cifVsFob.validCells.length +
                labelling.consumption.validCells.length + labelling.supplier.validCells.length + labelling.cifVsFob.validCells.length +
                profitMargin.validCells.length;

            const totalInvalid = fabricYarn.invalidCells.length +
                trim.consumption.invalidCells.length + trim.supplier.invalidCells.length + trim.cifVsFob.invalidCells.length +
                labelling.consumption.invalidCells.length + labelling.supplier.invalidCells.length + labelling.cifVsFob.invalidCells.length +
                profitMargin.invalidCells.length;

            html += `
                <div class="file-summary-box">
                    <strong>File:</strong> ${fileResult.fileName}<br>
                    <strong>Total Validation:</strong> ${totalValid} passed, ${totalInvalid} failed
                </div>
            `;

            html += `
                <table id="v19ResultsTable" class="results-table">
                    <thead>
                        <tr class="header-labels-row">
                            <th>Type</th>
                            <th>Column</th>
                            <th>Cells</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">Fabric/Yarn</td>
                            <td style="padding: 0.875rem 1rem;">K (Consumption = 5%)</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatAllCells(fabricYarn.validCells, fabricYarn.invalidCells)}</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;" rowspan="3">Trim</td>
                            <td style="padding: 0.875rem 1rem;">K (Consumption = 3%)</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatAllCells(trim.consumption.validCells, trim.consumption.invalidCells)}</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem;">E (Supplier = "Local" or "Nominated")</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatAllCells(trim.supplier.validCells, trim.supplier.invalidCells)}</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem;">H (C.I.F.VS FOB = 0.012% or 15%)</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatAllCells(trim.cifVsFob.validCells, trim.cifVsFob.invalidCells)}</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;" rowspan="3">Labelling</td>
                            <td style="padding: 0.875rem 1rem;">K (Consumption = 3%)</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatAllCells(labelling.consumption.validCells, labelling.consumption.invalidCells)}</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem;">E (Supplier = "Local" or "Nominated")</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatAllCells(labelling.supplier.validCells, labelling.supplier.invalidCells)}</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem;">H (C.I.F.VS FOB = 0.012% or 15%)</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatAllCells(labelling.cifVsFob.validCells, labelling.cifVsFob.invalidCells)}</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">Profit Margin</td>
                            <td style="padding: 0.875rem 1rem;">M (Value = 0.60-0.95)</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatAllCells(profitMargin.validCells, profitMargin.invalidCells)}</td>
                        </tr>
                    </tbody>
                </table>
            </div>`;
        }

        return html;
    }

    /**
     * Generate error HTML
     */
    generateErrorHTML(errorMessage) {
        return `
            <div style="background: #fee; border-left: 4px solid #dc3545; padding: 1.5rem; border-radius: 8px;">
                <p style="color: #dc3545; font-weight: 600; margin-bottom: 0.5rem;">
                    Error Processing Files
                </p>
                <p style="color: #721c24; font-size: 0.95rem;">
                    ${errorMessage}
                </p>
            </div>
        `;
    }

    /**
     * Export results to PDF
     */
    async exportToPDF() {
        if (!window.pdfExporter) {
            console.error('PDF Exporter not loaded');
            alert('PDF export module not available. Please refresh the page.');
            return;
        }

        if (!this.bcbdResults || this.bcbdResults.length === 0) {
            alert('No results to export. Please generate results first.');
            return;
        }

        const config = window.pdfExporter.createKuhlConfig(this.bcbdResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }

    /**
     * Search by filename
     */
    searchByFilename(searchTerm) {
        const fileGroups = document.querySelectorAll('#tab-v19 .file-result-group');

        if (!fileGroups || fileGroups.length === 0) {
            return;
        }

        const searchLower = searchTerm.toLowerCase().trim();

        if (searchLower === '') {
            fileGroups.forEach(group => {
                group.style.display = '';
            });
            return;
        }

        fileGroups.forEach(group => {
            const summaryBox = group.querySelector('.file-summary-box');
            if (!summaryBox) return;

            const fullText = summaryBox.textContent || summaryBox.innerText;
            const lines = fullText.split(/\r?\n/).map(line => line.trim());
            let filename = '';

            for (const line of lines) {
                if (line.toLowerCase().startsWith('file:')) {
                    filename = line.substring(5).trim().toLowerCase();
                    break;
                }
            }

            if (filename && filename.includes(searchLower)) {
                group.style.display = '';
            } else {
                group.style.display = 'none';
            }
        });
    }
}

// Initialize the processor
window.kuhlProcessor = new KuhlProcessor();

// Auto-initialize when V19 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    const v19Tabs = document.querySelectorAll('[data-tab="v19"]');
    v19Tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            window.kuhlProcessor.initialize();
        });
    });

    const v19TabContent = document.getElementById('tab-v19');
    if (v19TabContent && v19TabContent.classList.contains('active')) {
        window.kuhlProcessor.initialize();
    }
});
