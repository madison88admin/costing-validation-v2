/**
 * Foot Asylum (V18) Processing Logic
 * Validates Fabrics, Trims, and Packaging sections in Buyer CBD files
 *
 * Validation Rules (Fabrics section):
 * - Column D: Must be TRUE (Main Material) - Fixed column
 * - Column P: Must be "USD" (Supplier Currency) - Fixed column
 * - Wastage %: Must be exactly 5% - Column detected dynamically from Row 2
 * - Overhead Cost: Must be exactly 0.5 - Column detected dynamically from Row 2
 * - Testing Cost: Must be exactly 0.1 - Column detected dynamically from Row 2
 * - Profit %: Must be exactly 10% - Column detected dynamically from Row 2
 *
 * Validation Rules (Trims section):
 * - Column D: Must be FALSE (Main Material) - Fixed column
 * - Column P: Must be "USD" (Supplier Currency) - Fixed column
 * - Wastage %: Must be exactly 3% - Column detected dynamically from Row 2
 *
 * Validation Rules (Packaging section):
 * - Column D: Must be FALSE (Main Material) - Fixed column
 * - Column P: Must be "USD" (Supplier Currency) - Fixed column
 * - Wastage %: Must be exactly 3% - Column detected dynamically from Row 2
 *
 * Dynamic Column Detection:
 * The column for Wastage % is detected by scanning Row 2 of each Excel file.
 */

class FootAsylumProcessor {
    constructor() {
        this.bcbdResults = [];
        this.validationRules = {
            mainMaterial: {
                column: 'D',
                columnIndex: 3,
                label: 'Main Material',
                shortLabel: 'D',
                expectedValue: true,
                expectedDisplay: 'TRUE'
            },
            supplierCurrency: {
                column: 'P',
                columnIndex: 15,
                label: 'Supplier Currency',
                shortLabel: 'P',
                expectedValue: 'USD',
                expectedDisplay: 'USD'
            },
            wastage: {
                column: 'W',
                columnIndex: 22,
                label: 'Wastage %',
                shortLabel: 'W',
                expectedValue: 0.05,
                expectedDisplay: '5%'
            },
            overheadCost: {
                column: 'AG',
                columnIndex: 32,
                label: 'Overhead Cost',
                shortLabel: 'AG',
                expectedValue: 0.5,
                expectedDisplay: '0.5'
            },
            testingCost: {
                column: 'AK',
                columnIndex: 36,
                label: 'Testing Cost',
                shortLabel: 'AK',
                expectedValue: 0.1,
                expectedDisplay: '0.1'
            },
            profitFOB: {
                column: 'AL',
                columnIndex: 37,
                label: 'Profit %',
                shortLabel: 'AL',
                expectedValue: 0.10,
                expectedDisplay: '10%'
            }
        };

        // Trims section validation rules
        this.trimsValidationRules = {
            mainMaterial: {
                column: 'D',
                columnIndex: 3,
                label: 'Main Material',
                shortLabel: 'D',
                expectedValue: false,
                expectedDisplay: 'FALSE'
            },
            supplierCurrency: {
                column: 'P',
                columnIndex: 15,
                label: 'Supplier Currency',
                shortLabel: 'P',
                expectedValue: 'USD',
                expectedDisplay: 'USD'
            },
            wastage: {
                column: 'W',
                columnIndex: 22,
                label: 'Wastage %',
                shortLabel: 'W',
                expectedValue: 0.03,
                expectedDisplay: '3%'
            }
        };

        // Packaging section validation rules (same as Trims)
        this.packagingValidationRules = {
            mainMaterial: {
                column: 'D',
                columnIndex: 3,
                label: 'Main Material',
                shortLabel: 'D',
                expectedValue: false,
                expectedDisplay: 'FALSE'
            },
            supplierCurrency: {
                column: 'P',
                columnIndex: 15,
                label: 'Supplier Currency',
                shortLabel: 'P',
                expectedValue: 'USD',
                expectedDisplay: 'USD'
            },
            wastage: {
                column: 'W',
                columnIndex: 22,
                label: 'Wastage %',
                shortLabel: 'W',
                expectedValue: 0.03,
                expectedDisplay: '3%'
            }
        };
    }

    /**
     * Initialize - Display validation rules in the OB drop zone
     */
    initialize() {
        this.displayValidationRules();
        console.log('Foot Asylum Processor initialized');
    }

    /**
     * Display validation rules in the OB drop zone
     */
    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v18');
        if (!obDropZone) return;

        let contentHTML = `
            <div class="burton-cost-container">
                <div class="burton-cost-header">

                </div>
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>FABRICS SECTION</strong> (between "Fabrics (...)" and "Trims (")</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line">Column D - Main Material: TRUE</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line">Column P - Supplier Currency: USD</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line">Wastage %: 5% (column detected from Row 2)</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line">Overhead Cost: 0.5 (column detected from Row 2)</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line">Testing Cost: 0.1 (column detected from Row 2)</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line">Profit % / Total FOB: 10% (column detected from Row 2)</div>
                    </div>
                    <div class="burton-cost-item" style="margin-top: 12px; border-top: 1px solid #cbd5e1; padding-top: 12px;">
                        <div class="burton-item-line"><strong>TRIMS SECTION</strong> (between "Trims (...)" and "Packaging")</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line">Column D - Main Material: FALSE</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line">Column P - Supplier Currency: USD</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line">Wastage %: 3% (column detected from Row 2)</div>
                    </div>
                    <div class="burton-cost-item" style="margin-top: 12px; border-top: 1px solid #cbd5e1; padding-top: 12px;">
                        <div class="burton-item-line"><strong>PACKAGING SECTION</strong> (between "Packaging (...)" and "Graphics")</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line">Column D - Main Material: FALSE</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line">Column P - Supplier Currency: USD</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line">Wastage %: 3% (column detected from Row 2)</div>
                    </div>
                </div>
            </div>
        `;

        obDropZone.innerHTML = contentHTML;
    }

    /**
     * Convert column letter to index (A=0, B=1, ..., Z=25, AA=26, etc.)
     */
    columnToIndex(column) {
        let index = 0;
        for (let i = 0; i < column.length; i++) {
            index = index * 26 + column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
        }
        return index - 1;
    }

    /**
     * Convert column index to letter (0=A, 1=B, ..., 25=Z, 26=AA, etc.)
     */
    indexToColumn(index) {
        let column = '';
        let temp = index + 1;
        while (temp > 0) {
            let remainder = (temp - 1) % 26;
            column = String.fromCharCode(65 + remainder) + column;
            temp = Math.floor((temp - 1) / 26);
        }
        return column;
    }

    /**
     * Scan Row 2 to find dynamic column positions for Wastage %, Overhead Cost, Testing Cost, and Profit %
     * @param {Array} jsonData - Parsed Excel data
     * @returns {Object} - Object with detected column positions
     */
    detectDynamicColumns(jsonData) {
        const row2 = jsonData[1]; // Row 2 (0-indexed)
        const dynamicColumns = {
            wastage: { column: null, columnIndex: null, found: false },
            overheadCost: { column: null, columnIndex: null, found: false },
            testingCost: { column: null, columnIndex: null, found: false },
            profitFOB: { column: null, columnIndex: null, found: false }
        };

        if (!row2) {
            console.warn('Row 2 not found in file');
            return dynamicColumns;
        }

        // Scan each cell in Row 2
        for (let i = 0; i < row2.length; i++) {
            const cellValue = row2[i];
            if (cellValue === undefined || cellValue === null || cellValue === '') continue;

            const cellText = cellValue.toString().trim().toLowerCase();

            // Check for Wastage %
            if (cellText === 'wastage %' || cellText === 'wastage%') {
                dynamicColumns.wastage.column = this.indexToColumn(i);
                dynamicColumns.wastage.columnIndex = i;
                dynamicColumns.wastage.found = true;
            }
            // Check for Overhead Cost
            else if (cellText === 'overhead cost' || cellText === 'overhead') {
                dynamicColumns.overheadCost.column = this.indexToColumn(i);
                dynamicColumns.overheadCost.columnIndex = i;
                dynamicColumns.overheadCost.found = true;
            }
            // Check for Testing Cost
            else if (cellText === 'testing cost' || cellText === 'testing') {
                dynamicColumns.testingCost.column = this.indexToColumn(i);
                dynamicColumns.testingCost.columnIndex = i;
                dynamicColumns.testingCost.found = true;
            }
            // Check for Profit %
            else if (cellText === 'profit %' || cellText === 'profit%' || cellText === 'profit % / total fob' || cellText === 'profit %/ total fob') {
                dynamicColumns.profitFOB.column = this.indexToColumn(i);
                dynamicColumns.profitFOB.columnIndex = i;
                dynamicColumns.profitFOB.found = true;
            }
        }

        console.log('Detected dynamic columns:', dynamicColumns);
        return dynamicColumns;
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
     * Find the Fabrics section boundaries
     * @param {Array} jsonData - Parsed Excel data
     * @returns {Object} - { startRow, endRow, sectionFound }
     */
    findFabricsSection(jsonData) {
        const colA = 0; // Column A index
        let startRow = -1;
        let endRow = -1;

        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || !row[colA]) continue;

            const cellValue = row[colA].toString().trim();

            // Find start: "Fabrics (" followed by any value and ")"
            if (startRow === -1) {
                const fabricsMatch = cellValue.match(/^Fabrics\s*\([^)]+\)/i);
                if (fabricsMatch) {
                    startRow = i + 1; // Start from the next row
                    continue;
                }
            }

            // Find end: "Trims ("
            if (startRow !== -1 && cellValue.match(/^Trims\s*\(/i)) {
                endRow = i;
                break;
            }
        }

        return {
            startRow,
            endRow,
            sectionFound: startRow !== -1 && endRow !== -1
        };
    }

    /**
     * Find the Trims section boundaries
     * @param {Array} jsonData - Parsed Excel data
     * @returns {Object} - { startRow, endRow, sectionFound }
     */
    findTrimsSection(jsonData) {
        const colA = 0; // Column A index
        let startRow = -1;
        let endRow = -1;

        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || !row[colA]) continue;

            const cellValue = row[colA].toString().trim();

            // Find start: "Trims (" followed by any value and ")"
            if (startRow === -1) {
                const trimsMatch = cellValue.match(/^Trims\s*\([^)]+\)/i);
                if (trimsMatch) {
                    startRow = i + 1; // Start from the next row
                    continue;
                }
            }

            // Find end: "Packaging"
            if (startRow !== -1 && cellValue.match(/^Packaging/i)) {
                endRow = i;
                break;
            }
        }

        return {
            startRow,
            endRow,
            sectionFound: startRow !== -1 && endRow !== -1
        };
    }

    /**
     * Find the Packaging section boundaries
     * @param {Array} jsonData - Parsed Excel data
     * @returns {Object} - { startRow, endRow, sectionFound }
     */
    findPackagingSection(jsonData) {
        const colA = 0; // Column A index
        let startRow = -1;
        let endRow = -1;

        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || !row[colA]) continue;

            const cellValue = row[colA].toString().trim();

            // Find start: "Packaging (" followed by any value and ")"
            if (startRow === -1) {
                const packagingMatch = cellValue.match(/^Packaging\s*\([^)]+\)/i);
                if (packagingMatch) {
                    startRow = i + 1; // Start from the next row
                    continue;
                }
            }

            // Find end: "Graphics"
            if (startRow !== -1 && cellValue.match(/^Graphics/i)) {
                endRow = i;
                break;
            }
        }

        return {
            startRow,
            endRow,
            sectionFound: startRow !== -1 && endRow !== -1
        };
    }

    /**
     * Validate a single value against expected value
     */
    validateValue(actualValue, rule) {
        if (actualValue === undefined || actualValue === null || actualValue === '') {
            return { isValid: false, displayValue: 'Empty', isEmpty: true };
        }

        let isValid = false;
        let displayValue = actualValue.toString();

        // Handle boolean TRUE
        if (rule.expectedValue === true) {
            if (typeof actualValue === 'boolean') {
                isValid = actualValue === true;
                displayValue = actualValue ? 'TRUE' : 'FALSE';
            } else {
                const strVal = actualValue.toString().toUpperCase().trim();
                isValid = strVal === 'TRUE';
                displayValue = strVal;
            }
        }
        // Handle boolean FALSE
        else if (rule.expectedValue === false) {
            if (typeof actualValue === 'boolean') {
                isValid = actualValue === false;
                displayValue = actualValue ? 'TRUE' : 'FALSE';
            } else {
                const strVal = actualValue.toString().toUpperCase().trim();
                isValid = strVal === 'FALSE';
                displayValue = strVal;
            }
        }
        // Handle string comparison (USD)
        else if (typeof rule.expectedValue === 'string') {
            const strVal = actualValue.toString().trim().toUpperCase();
            isValid = strVal === rule.expectedValue.toUpperCase();
            displayValue = actualValue.toString().trim();
        }
        // Handle percentage values (5%, 10%)
        else if (rule.expectedDisplay.includes('%')) {
            let numericValue;
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

            if (!isNaN(numericValue)) {
                // Normalize the value: if >= 1, it's stored as percentage (e.g., 10 for 10%)
                // Convert to decimal for comparison (divide by 100)
                let normalizedValue = numericValue;
                if (numericValue >= 1) {
                    normalizedValue = numericValue / 100;
                }

                isValid = Math.abs(normalizedValue - rule.expectedValue) < 0.0001;

                // Display as percentage
                if (numericValue < 1) {
                    displayValue = (numericValue * 100).toFixed(0) + '%';
                } else {
                    displayValue = numericValue.toFixed(0) + '%';
                }
            }
        }
        // Handle numeric values (0.5, 0.1)
        else if (typeof rule.expectedValue === 'number') {
            let numericValue;
            if (typeof actualValue === 'number') {
                numericValue = actualValue;
            } else {
                numericValue = parseFloat(actualValue.toString().trim());
            }

            if (!isNaN(numericValue)) {
                isValid = Math.abs(numericValue - rule.expectedValue) < 0.0001;
                displayValue = numericValue.toString();
            }
        }

        return { isValid, displayValue, isEmpty: false };
    }

    /**
     * Validate file against rules - returns row-based results for Fabrics, Trims, and Packaging sections
     */
    validateFile(jsonData) {
        const results = {
            detectedColumns: {}, // Store detected column positions for display
            fabrics: {
                sectionFound: false,
                startRow: -1,
                endRow: -1,
                rows: [],
                activeRules: null
            },
            trims: {
                sectionFound: false,
                startRow: -1,
                endRow: -1,
                rows: [],
                activeRules: null
            },
            packaging: {
                sectionFound: false,
                startRow: -1,
                endRow: -1,
                rows: [],
                activeRules: null
            }
        };

        // First, detect dynamic columns from Row 2
        const dynamicColumns = this.detectDynamicColumns(jsonData);
        results.detectedColumns = dynamicColumns;

        // =====================
        // FABRICS SECTION
        // =====================
        const fabricsActiveRules = JSON.parse(JSON.stringify(this.validationRules));

        // Update Wastage column if detected
        if (dynamicColumns.wastage.found) {
            fabricsActiveRules.wastage.column = dynamicColumns.wastage.column;
            fabricsActiveRules.wastage.columnIndex = dynamicColumns.wastage.columnIndex;
            fabricsActiveRules.wastage.shortLabel = dynamicColumns.wastage.column;
        }

        // Update Overhead Cost column if detected
        if (dynamicColumns.overheadCost.found) {
            fabricsActiveRules.overheadCost.column = dynamicColumns.overheadCost.column;
            fabricsActiveRules.overheadCost.columnIndex = dynamicColumns.overheadCost.columnIndex;
            fabricsActiveRules.overheadCost.shortLabel = dynamicColumns.overheadCost.column;
        }

        // Update Testing Cost column if detected
        if (dynamicColumns.testingCost.found) {
            fabricsActiveRules.testingCost.column = dynamicColumns.testingCost.column;
            fabricsActiveRules.testingCost.columnIndex = dynamicColumns.testingCost.columnIndex;
            fabricsActiveRules.testingCost.shortLabel = dynamicColumns.testingCost.column;
        }

        // Update Profit % column if detected
        if (dynamicColumns.profitFOB.found) {
            fabricsActiveRules.profitFOB.column = dynamicColumns.profitFOB.column;
            fabricsActiveRules.profitFOB.columnIndex = dynamicColumns.profitFOB.columnIndex;
            fabricsActiveRules.profitFOB.shortLabel = dynamicColumns.profitFOB.column;
        }

        results.fabrics.activeRules = fabricsActiveRules;

        // Find Fabrics section
        const fabricsSection = this.findFabricsSection(jsonData);
        results.fabrics.sectionFound = fabricsSection.sectionFound;
        results.fabrics.startRow = fabricsSection.startRow;
        results.fabrics.endRow = fabricsSection.endRow;

        if (fabricsSection.sectionFound) {
            // Validate each row in the Fabrics section
            for (let i = fabricsSection.startRow; i < fabricsSection.endRow; i++) {
                const row = jsonData[i];
                if (!row) continue;

                const hasData = Object.values(fabricsActiveRules).some(rule => {
                    const cellValue = row[rule.columnIndex];
                    return cellValue !== undefined && cellValue !== null && cellValue !== '';
                });
                if (!hasData) continue;

                const rowResult = {
                    rowNumber: i + 1,
                    columns: {}
                };

                let rowHasAnyData = false;

                for (const [key, rule] of Object.entries(fabricsActiveRules)) {
                    const cellValue = row[rule.columnIndex];
                    const validation = this.validateValue(cellValue, rule);

                    rowResult.columns[key] = {
                        value: validation.displayValue,
                        isValid: validation.isValid,
                        isEmpty: validation.isEmpty,
                        expected: rule.expectedDisplay,
                        column: rule.column
                    };

                    if (!validation.isEmpty) {
                        rowHasAnyData = true;
                    }
                }

                if (rowHasAnyData) {
                    results.fabrics.rows.push(rowResult);
                }
            }
        }

        // =====================
        // TRIMS SECTION
        // =====================
        const trimsActiveRules = JSON.parse(JSON.stringify(this.trimsValidationRules));

        // Update Wastage column if detected (same column as Fabrics, but different expected value)
        if (dynamicColumns.wastage.found) {
            trimsActiveRules.wastage.column = dynamicColumns.wastage.column;
            trimsActiveRules.wastage.columnIndex = dynamicColumns.wastage.columnIndex;
            trimsActiveRules.wastage.shortLabel = dynamicColumns.wastage.column;
        }

        results.trims.activeRules = trimsActiveRules;

        // Find Trims section
        const trimsSection = this.findTrimsSection(jsonData);
        results.trims.sectionFound = trimsSection.sectionFound;
        results.trims.startRow = trimsSection.startRow;
        results.trims.endRow = trimsSection.endRow;

        if (trimsSection.sectionFound) {
            // Validate each row in the Trims section
            for (let i = trimsSection.startRow; i < trimsSection.endRow; i++) {
                const row = jsonData[i];
                if (!row) continue;

                const hasData = Object.values(trimsActiveRules).some(rule => {
                    const cellValue = row[rule.columnIndex];
                    return cellValue !== undefined && cellValue !== null && cellValue !== '';
                });
                if (!hasData) continue;

                const rowResult = {
                    rowNumber: i + 1,
                    columns: {}
                };

                let rowHasAnyData = false;

                for (const [key, rule] of Object.entries(trimsActiveRules)) {
                    const cellValue = row[rule.columnIndex];
                    const validation = this.validateValue(cellValue, rule);

                    rowResult.columns[key] = {
                        value: validation.displayValue,
                        isValid: validation.isValid,
                        isEmpty: validation.isEmpty,
                        expected: rule.expectedDisplay,
                        column: rule.column
                    };

                    if (!validation.isEmpty) {
                        rowHasAnyData = true;
                    }
                }

                if (rowHasAnyData) {
                    results.trims.rows.push(rowResult);
                }
            }
        }

        // =====================
        // PACKAGING SECTION
        // =====================
        const packagingActiveRules = JSON.parse(JSON.stringify(this.packagingValidationRules));

        // Update Wastage column if detected
        if (dynamicColumns.wastage.found) {
            packagingActiveRules.wastage.column = dynamicColumns.wastage.column;
            packagingActiveRules.wastage.columnIndex = dynamicColumns.wastage.columnIndex;
            packagingActiveRules.wastage.shortLabel = dynamicColumns.wastage.column;
        }

        results.packaging.activeRules = packagingActiveRules;

        // Find Packaging section
        const packagingSection = this.findPackagingSection(jsonData);
        results.packaging.sectionFound = packagingSection.sectionFound;
        results.packaging.startRow = packagingSection.startRow;
        results.packaging.endRow = packagingSection.endRow;

        if (packagingSection.sectionFound) {
            // Validate each row in the Packaging section
            for (let i = packagingSection.startRow; i < packagingSection.endRow; i++) {
                const row = jsonData[i];
                if (!row) continue;

                const hasData = Object.values(packagingActiveRules).some(rule => {
                    const cellValue = row[rule.columnIndex];
                    return cellValue !== undefined && cellValue !== null && cellValue !== '';
                });
                if (!hasData) continue;

                const rowResult = {
                    rowNumber: i + 1,
                    columns: {}
                };

                let rowHasAnyData = false;

                for (const [key, rule] of Object.entries(packagingActiveRules)) {
                    const cellValue = row[rule.columnIndex];
                    const validation = this.validateValue(cellValue, rule);

                    rowResult.columns[key] = {
                        value: validation.displayValue,
                        isValid: validation.isValid,
                        isEmpty: validation.isEmpty,
                        expected: rule.expectedDisplay,
                        column: rule.column
                    };

                    if (!validation.isEmpty) {
                        rowHasAnyData = true;
                    }
                }

                if (rowHasAnyData) {
                    results.packaging.rows.push(rowResult);
                }
            }
        }

        return results;
    }

    /**
     * Format cell value with color coding
     */
    formatCellValue(cellData) {
        if (cellData.isEmpty) {
            return `<span style="color: #64748b;">-</span>`;
        }

        if (cellData.isValid) {
            return `<span style="color: #065f46; font-weight: 600;">${cellData.value}</span>`;
        } else {
            return `<span style="color: #991b1b; font-weight: 600;">${cellData.value}</span><br><span style="font-size: 0.75em; color: #849bba;">Expected: ${cellData.expected}</span>`;
        }
    }

    /**
     * Generate HTML for results display
     */
    generateResultsHTML(results) {
        if (!results || results.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">Foot Asylum Validation Ready</p>
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
                        oninput="window.footAsylumProcessor.searchByFilename(this.value)"
                    />
                </div>
                <button onclick="window.footAsylumProcessor.exportToPDF()" class="export-btn">
                    Export
                </button>
            </div>
        `;

        for (const fileResult of results) {
            html += `<div class="file-result-group">`;

            // Calculate validation counts for all sections
            let totalValid = 0;
            let totalInvalid = 0;

            // Count Fabrics validations
            for (const rowResult of fileResult.results.fabrics.rows) {
                for (const [key, cellData] of Object.entries(rowResult.columns)) {
                    if (!cellData.isEmpty) {
                        if (cellData.isValid) totalValid++;
                        else totalInvalid++;
                    }
                }
            }

            // Count Trims validations
            for (const rowResult of fileResult.results.trims.rows) {
                for (const [key, cellData] of Object.entries(rowResult.columns)) {
                    if (!cellData.isEmpty) {
                        if (cellData.isValid) totalValid++;
                        else totalInvalid++;
                    }
                }
            }

            // Count Packaging validations
            for (const rowResult of fileResult.results.packaging.rows) {
                for (const [key, cellData] of Object.entries(rowResult.columns)) {
                    if (!cellData.isEmpty) {
                        if (cellData.isValid) totalValid++;
                        else totalInvalid++;
                    }
                }
            }

            const fabricsSectionStatus = fileResult.results.fabrics.sectionFound
                ? `Found (rows ${fileResult.results.fabrics.startRow + 1}-${fileResult.results.fabrics.endRow})`
                : `Not found`;

            const trimsSectionStatus = fileResult.results.trims.sectionFound
                ? `Found (rows ${fileResult.results.trims.startRow + 1}-${fileResult.results.trims.endRow})`
                : `Not found`;

            const packagingSectionStatus = fileResult.results.packaging.sectionFound
                ? `Found (rows ${fileResult.results.packaging.startRow + 1}-${fileResult.results.packaging.endRow})`
                : `Not found`;

            // Build detected columns info
            const detectedCols = fileResult.results.detectedColumns;
            let detectedColsText = '';
            if (detectedCols) {
                const colInfo = [];
                if (detectedCols.wastage.found) colInfo.push(`Wastage: ${detectedCols.wastage.column}`);
                if (detectedCols.overheadCost.found) colInfo.push(`Overhead: ${detectedCols.overheadCost.column}`);
                if (detectedCols.testingCost.found) colInfo.push(`Testing: ${detectedCols.testingCost.column}`);
                if (detectedCols.profitFOB.found) colInfo.push(`Profit: ${detectedCols.profitFOB.column}`);
                if (colInfo.length > 0) {
                    detectedColsText = colInfo.join(', ');
                }
            }

            html += `
                <div class="file-summary-box">
                    <strong>File:</strong> ${fileResult.fileName}<br>
                    <strong>Fabrics:</strong> ${fabricsSectionStatus} | <strong>Trims:</strong> ${trimsSectionStatus} | <strong>Packaging:</strong> ${packagingSectionStatus}<br>
                    <strong>Detected Columns:</strong> ${detectedColsText}<br>
                    <strong>Summary:</strong> ${totalValid} passed, ${totalInvalid} failed
                </div>
            `;

            // Check if we have any data to display
            const hasFabricsData = fileResult.results.fabrics.sectionFound && fileResult.results.fabrics.rows.length > 0;
            const hasTrimsData = fileResult.results.trims.sectionFound && fileResult.results.trims.rows.length > 0;
            const hasPackagingData = fileResult.results.packaging.sectionFound && fileResult.results.packaging.rows.length > 0;

            if (hasFabricsData || hasTrimsData || hasPackagingData) {
                const fabricsRules = fileResult.results.fabrics.activeRules || this.validationRules;
                const fabricsRuleKeys = Object.keys(fabricsRules);

                html += `
                    <table class="results-table">
                        <thead>
                            <tr class="header-labels-row">
                                <th style="width: 80px;">Section</th>
                                <th style="width: 60px;">Row</th>
                `;

                // Use Fabrics headers (more columns)
                for (const key of fabricsRuleKeys) {
                    const rule = fabricsRules[key];
                    html += `<th>${rule.label}<br><span style="font-size: 0.75em; font-weight: normal; color: #64748b;">(${rule.column}) ${rule.expectedDisplay}</span></th>`;
                }

                html += `
                            </tr>
                        </thead>
                        <tbody>
                `;

                // Add Fabrics rows
                if (hasFabricsData) {
                    for (const rowResult of fileResult.results.fabrics.rows) {
                        html += `
                            <tr style="border-bottom: 1px solid #e0e8f0;">
                                <td style="padding: 0.875rem 1rem; font-weight: 600;">Fabrics</td>
                                <td style="padding: 0.875rem 1rem; font-weight: 600;">${rowResult.rowNumber}</td>
                        `;

                        for (const key of fabricsRuleKeys) {
                            const cellData = rowResult.columns[key];
                            html += `<td style="padding: 0.875rem 1rem;">${this.formatCellValue(cellData)}</td>`;
                        }

                        html += `</tr>`;
                    }
                }

                // Add Trims rows
                if (hasTrimsData) {
                    const trimsRules = fileResult.results.trims.activeRules || this.trimsValidationRules;

                    for (const rowResult of fileResult.results.trims.rows) {
                        html += `
                            <tr style="border-bottom: 1px solid #e0e8f0;">
                                <td style="padding: 0.875rem 1rem; font-weight: 600;">Trims</td>
                                <td style="padding: 0.875rem 1rem; font-weight: 600;">${rowResult.rowNumber}</td>
                        `;

                        // For Trims, we only have 3 columns, so fill the rest with dashes
                        for (const key of fabricsRuleKeys) {
                            if (trimsRules[key]) {
                                const cellData = rowResult.columns[key];
                                html += `<td style="padding: 0.875rem 1rem;">${this.formatCellValue(cellData)}</td>`;
                            } else {
                                html += `<td style="padding: 0.875rem 1rem; color: #64748b;">-</td>`;
                            }
                        }

                        html += `</tr>`;
                    }
                }

                // Add Packaging rows
                if (hasPackagingData) {
                    const packagingRules = fileResult.results.packaging.activeRules || this.packagingValidationRules;

                    for (const rowResult of fileResult.results.packaging.rows) {
                        html += `
                            <tr style="border-bottom: 1px solid #e0e8f0;">
                                <td style="padding: 0.875rem 1rem; font-weight: 600;">Packaging</td>
                                <td style="padding: 0.875rem 1rem; font-weight: 600;">${rowResult.rowNumber}</td>
                        `;

                        // For Packaging, we only have 3 columns, so fill the rest with dashes
                        for (const key of fabricsRuleKeys) {
                            if (packagingRules[key]) {
                                const cellData = rowResult.columns[key];
                                html += `<td style="padding: 0.875rem 1rem;">${this.formatCellValue(cellData)}</td>`;
                            } else {
                                html += `<td style="padding: 0.875rem 1rem; color: #64748b;">-</td>`;
                            }
                        }

                        html += `</tr>`;
                    }
                }

                html += `
                        </tbody>
                    </table>
                `;
            }

            // Show error messages if sections not found
            if (!fileResult.results.fabrics.sectionFound) {
                html += `
                    <div style="padding: 1rem; background: #fee; border-radius: 8px; margin-top: 1rem;">
                        <p style="color: #991b1b; font-weight: 600;">Fabrics section not found in file.</p>
                    </div>
                `;
            }

            if (!fileResult.results.trims.sectionFound) {
                html += `
                    <div style="padding: 1rem; background: #fee; border-radius: 8px; margin-top: 1rem;">
                        <p style="color: #991b1b; font-weight: 600;">Trims section not found in file.</p>
                    </div>
                `;
            }

            if (!fileResult.results.packaging.sectionFound) {
                html += `
                    <div style="padding: 1rem; background: #fee; border-radius: 8px; margin-top: 1rem;">
                        <p style="color: #991b1b; font-weight: 600;">Packaging section not found in file.</p>
                    </div>
                `;
            }

            html += `</div>`; // Close file-result-group
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

        const config = window.pdfExporter.createFootAsylumConfig(this.bcbdResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }

    /**
     * Search by filename
     */
    searchByFilename(searchTerm) {
        const fileGroups = document.querySelectorAll('#tab-v18 .file-result-group');

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
window.footAsylumProcessor = new FootAsylumProcessor();

// Auto-initialize when V18 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    const v18Tabs = document.querySelectorAll('[data-tab="v18"]');
    v18Tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            window.footAsylumProcessor.initialize();
        });
    });

    const v18TabContent = document.getElementById('tab-v18');
    if (v18TabContent && v18TabContent.classList.contains('active')) {
        window.footAsylumProcessor.initialize();
    }
});
