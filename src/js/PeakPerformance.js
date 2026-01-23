/**
 * Peak Performance Processing Logic (V10)
 * Validates Buyer CBD files against Peak Performance criteria
 *
 * Validation Rules:
 * - Find "Fabric/Yarn" in Column A
 * - Check Column J for "5%" from that row until "Fabric Subtotal" is found in Column A
 * - Validate standard items from CSV against uploaded Excel files
 */

class PeakPerformanceProcessor {
    constructor() {
        this.bcbdFiles = [];
        this.bcbdResults = [];
        this.standardItems = []; // Loaded from CSV
        this.csvLoaded = false;
        this.validationRules = {
            // Fabric/Yarn wastage section
            fabricWastage: {
                startKeyword: 'FABRIC/YARN',
                stopKeyword: 'FABRIC SUBTOTAL',
                searchCol: 0,           // Column A (index 0)
                wastageCol: 9,          // Column J (index 9)
                expectedValue: '5%',
                label: 'Fabric/Yarn Wastage'
            },
            // Standard items column mapping
            columnMapping: {
                supplier: 0,           // Column A
                supplierItem: 1,       // Column B
                garmentPart: 2,        // Column C
                materialDesc: 3,       // Column D
                yield: 8,              // Column I
                wastage: 9,            // Column J
                fobUnitCost: 10,       // Column K
                cifUnitCost: 11        // Column L
            }
        };
    }

    /**
     * Initialize V10 - Load CSV and display validation rules
     */
    async initialize() {
        await this.loadCSVData();
        this.displayValidationRules();
        console.log('Peak Performance Processor initialized');
    }

    /**
     * Load standard items from CSV file
     */
    async loadCSVData() {
        try {
            const response = await fetch('assets/data/PeakPerformance.csv');
            const csvText = await response.text();
            this.parseCSV(csvText);
            this.csvLoaded = true;
            console.log('Peak Performance CSV loaded:', this.standardItems.length, 'items');
        } catch (error) {
            console.error('Error loading Peak Performance CSV:', error);
            this.csvLoaded = false;
        }
    }

    /**
     * Parse CSV text into standard items array
     */
    parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        this.standardItems = [];

        // Skip header row (first line)
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;

            // Parse CSV line (handle commas in quoted fields)
            const values = this.parseCSVLine(line);

            if (values.length >= 8) {
                this.standardItems.push({
                    supplier: values[0] || '-',
                    supplierItem: values[1] || '-',
                    garmentPart: values[2] || '-',
                    materialDesc: values[3] || '-',
                    yield: values[4] || '-',
                    wastage: values[5] || '-',
                    fobUnitCost: values[6] || '-',
                    cifUnitCost: values[7] || '-'
                });
            }
        }
    }

    /**
     * Parse a single CSV line handling quoted fields
     */
    parseCSVLine(line) {
        const result = [];
        let current = '';
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
            const char = line[i];

            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === ',' && !inQuotes) {
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        result.push(current.trim());

        return result;
    }

    /**
     * Display validation rules in the OB drop zone (Burton-style)
     */
    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v10');
        if (!obDropZone) return;

        const fw = this.validationRules.fabricWastage;

        let html = `
            <div class="burton-cost-container">
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Fabric/Yarn Wastage Check</strong></div>
                        <div class="burton-item-line"><strong>Start:</strong> Find "${fw.startKeyword}" in Column A</div>
                        <div class="burton-item-line"><strong>Check:</strong> Column J must contain "${fw.expectedValue}"</div>
                        <div class="burton-item-line"><strong>Stop:</strong> When "${fw.stopKeyword}" is found in Column A</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Standard Items Check</strong></div>
                        <div class="burton-item-line"><strong>Items loaded:</strong> ${this.standardItems.length}</div>
                        <div class="burton-item-line"><strong>Columns:</strong> A (Supplier), B (Item#), C (Part), D (Desc), I (Yield), J (Wastage), K (FOB), L (CIF)</div>
                    </div>
                </div>
            </div>
        `;

        obDropZone.innerHTML = html;
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

            // Ensure CSV is loaded
            if (!this.csvLoaded) {
                await this.loadCSVData();
            }

            // Process each BCBD file
            for (const file of bcbdFiles) {
                const validationResult = await this.validateFile(file);
                this.bcbdResults.push({
                    fileName: file.name,
                    results: validationResult
                });
            }

            return this.generateResultsHTML(this.bcbdResults);

        } catch (error) {
            console.error('Error processing files:', error);
            return this.generateErrorHTML(error.message);
        }
    }

    /**
     * Validate a single file
     */
    async validateFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    // Get the first sheet
                    const firstSheetName = workbook.SheetNames[0];
                    const sheet = workbook.Sheets[firstSheetName];

                    // Convert to JSON for easier processing
                    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

                    // Run all validations
                    const fabricWastageCheck = this.checkFabricWastage(jsonData);
                    const standardItemsCheck = this.checkStandardItems(jsonData);

                    resolve({
                        fabricWastageCheck: fabricWastageCheck,
                        standardItemsCheck: standardItemsCheck
                    });
                } catch (error) {
                    reject(error);
                }
            };

            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Check Fabric/Yarn wastage values
     * Looks for "Fabric/Yarn" in Column A, then checks Column J for "5%" until "Fabric Subtotal" is found
     */
    checkFabricWastage(jsonData) {
        const fw = this.validationRules.fabricWastage;
        const searchCol = fw.searchCol;     // Column A (index 0)
        const wastageCol = fw.wastageCol;   // Column J (index 9)

        // Find the start row (Fabric/Yarn)
        let startRowIndex = -1;
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const colA = row[searchCol] ? String(row[searchCol]).trim().toUpperCase() : '';

            if (colA.includes(fw.startKeyword)) {
                startRowIndex = i;
                console.log(`Found "${fw.startKeyword}" at row ${i + 1}`);
                break;
            }
        }

        if (startRowIndex === -1) {
            return {
                found: false,
                message: `"${fw.startKeyword}" not found in Column A`
            };
        }

        // Skip the header row (the row after Fabric/Yarn contains column headers like "Handling/Wastage")
        const dataStartRow = startRowIndex + 2; // Skip Fabric/Yarn row and header row

        // Find the stop row (Fabric Subtotal)
        let stopRowIndex = -1;
        for (let i = dataStartRow; i < jsonData.length; i++) {
            const row = jsonData[i];
            const colA = row[searchCol] ? String(row[searchCol]).trim().toUpperCase() : '';

            if (colA.includes(fw.stopKeyword)) {
                stopRowIndex = i;
                console.log(`Found "${fw.stopKeyword}" at row ${i + 1}`);
                break;
            }
        }

        if (stopRowIndex === -1) {
            return {
                found: false,
                message: `"${fw.stopKeyword}" not found in Column A after "${fw.startKeyword}"`
            };
        }

        // Check Column J for "5%" from dataStartRow to stopRow (exclusive)
        const validCells = [];
        const invalidCells = [];

        for (let i = dataStartRow; i < stopRowIndex; i++) {
            const row = jsonData[i];
            const cellValue = row[wastageCol];

            // Skip empty cells
            if (cellValue === null || cellValue === undefined || cellValue === '') {
                continue;
            }

            const cellAddress = `J${i + 1}`;
            const cellValueStr = String(cellValue).trim();

            // Check if the value contains "5%" or equals 5% or 0.05
            const isValid = this.isValidWastage(cellValue, 5);

            if (isValid) {
                validCells.push({
                    rowNumber: i + 1,
                    cellAddress: cellAddress,
                    value: cellValueStr
                });
            } else {
                invalidCells.push({
                    rowNumber: i + 1,
                    cellAddress: cellAddress,
                    value: cellValueStr
                });
            }
        }

        return {
            found: true,
            startRow: startRowIndex + 1,
            stopRow: stopRowIndex + 1,
            validCells: validCells,
            invalidCells: invalidCells,
            isValid: invalidCells.length === 0,
            summary: `${validCells.length} valid, ${invalidCells.length} invalid`
        };
    }

    /**
     * Check standard items from CSV against Excel data
     */
    checkStandardItems(jsonData) {
        if (!this.standardItems || this.standardItems.length === 0) {
            return {
                found: false,
                message: 'No standard items loaded from CSV'
            };
        }

        const colMap = this.validationRules.columnMapping;
        const results = [];

        // For each standard item from CSV
        for (const stdItem of this.standardItems) {
            const itemResult = {
                standardItem: stdItem,
                found: false,
                rowNumber: -1,
                checks: [],
                isValid: true
            };

            // Find matching row in Excel by Material Description (Column D)
            // Use exact case-sensitive matching to avoid "Hangtag" matching "SMS Hangtag"
            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                const materialDesc = row[colMap.materialDesc] ? String(row[colMap.materialDesc]).trim() : '';

                // Match by exact material description (case-sensitive)
                if (materialDesc && stdItem.materialDesc !== '-' &&
                    materialDesc === stdItem.materialDesc) {

                    itemResult.found = true;
                    itemResult.rowNumber = i + 1;

                    // Check each column
                    itemResult.checks = this.validateItemColumns(row, stdItem, i + 1, colMap);
                    itemResult.isValid = itemResult.checks.every(c => c.isValid);
                    break;
                }
            }

            results.push(itemResult);
        }

        const foundCount = results.filter(r => r.found).length;
        const validCount = results.filter(r => r.found && r.isValid).length;

        return {
            found: true,
            items: results,
            foundCount: foundCount,
            validCount: validCount,
            totalCount: this.standardItems.length,
            isValid: validCount === foundCount && foundCount > 0
        };
    }

    /**
     * Validate individual columns for a standard item
     */
    validateItemColumns(row, stdItem, rowNum, colMap) {
        const checks = [];

        // Check Supplier (Column A)
        if (stdItem.supplier !== '-') {
            const actual = row[colMap.supplier] ? String(row[colMap.supplier]).trim() : '';
            checks.push({
                column: 'A',
                label: 'Supplier',
                expected: stdItem.supplier,
                actual: actual,
                cellAddress: `A${rowNum}`,
                isValid: actual.toUpperCase().includes(stdItem.supplier.toUpperCase())
            });
        }

        // Check Supplier Item # (Column B)
        if (stdItem.supplierItem !== '-') {
            const actual = row[colMap.supplierItem] ? String(row[colMap.supplierItem]).trim() : '';
            checks.push({
                column: 'B',
                label: 'Supplier Item #',
                expected: stdItem.supplierItem,
                actual: actual,
                cellAddress: `B${rowNum}`,
                isValid: actual === stdItem.supplierItem
            });
        }

        // Check Garment Part (Column C)
        if (stdItem.garmentPart !== '-') {
            const actual = row[colMap.garmentPart] ? String(row[colMap.garmentPart]).trim() : '';
            checks.push({
                column: 'C',
                label: 'Garment Part',
                expected: stdItem.garmentPart,
                actual: actual,
                cellAddress: `C${rowNum}`,
                isValid: actual.toUpperCase().includes(stdItem.garmentPart.toUpperCase())
            });
        }

        // Check Yield (Column I)
        if (stdItem.yield !== '-') {
            const actual = row[colMap.yield];
            const actualStr = actual !== null && actual !== undefined ? String(actual).trim() : '';
            const expectedNum = parseFloat(stdItem.yield);
            const actualNum = parseFloat(actualStr);
            checks.push({
                column: 'I',
                label: 'Yield',
                expected: stdItem.yield,
                actual: actualStr,
                cellAddress: `I${rowNum}`,
                isValid: !isNaN(expectedNum) && !isNaN(actualNum) && Math.abs(expectedNum - actualNum) < 0.01
            });
        }

        // Check Handling/Wastage (Column J)
        if (stdItem.wastage !== '-') {
            const actual = row[colMap.wastage];
            const actualStr = actual !== null && actual !== undefined ? String(actual).trim() : '';
            checks.push({
                column: 'J',
                label: 'Wastage',
                expected: stdItem.wastage,
                actual: actualStr,
                cellAddress: `J${rowNum}`,
                isValid: this.compareWastage(stdItem.wastage, actualStr)
            });
        }

        // Check FOB Unit Cost (Column K)
        if (stdItem.fobUnitCost !== '-') {
            const actual = row[colMap.fobUnitCost];
            const actualStr = actual !== null && actual !== undefined ? String(actual).trim() : '';
            checks.push({
                column: 'K',
                label: 'FOB Unit Cost',
                expected: stdItem.fobUnitCost,
                actual: actualStr,
                cellAddress: `K${rowNum}`,
                isValid: this.compareCost(stdItem.fobUnitCost, actualStr)
            });
        }

        // Check CIF Unit Cost (Column L)
        if (stdItem.cifUnitCost !== '-') {
            const actual = row[colMap.cifUnitCost];
            const actualStr = actual !== null && actual !== undefined ? String(actual).trim() : '';
            checks.push({
                column: 'L',
                label: 'CIF Unit Cost',
                expected: stdItem.cifUnitCost,
                actual: actualStr,
                cellAddress: `L${rowNum}`,
                isValid: this.compareCost(stdItem.cifUnitCost, actualStr)
            });
        }

        return checks;
    }

    /**
     * Compare wastage values (handles % format)
     */
    compareWastage(expected, actual) {
        if (!expected || !actual) return false;

        // Parse expected
        let expectedNum = parseFloat(String(expected).replace(/[%,\s]/g, ''));
        // Parse actual
        let actualNum = parseFloat(String(actual).replace(/[%,\s]/g, ''));

        if (isNaN(expectedNum) || isNaN(actualNum)) return false;

        // If actual is decimal (0.03) and expected is percentage (3), convert
        if (actualNum < 1 && expectedNum >= 1) {
            actualNum = actualNum * 100;
        }

        return Math.abs(expectedNum - actualNum) < 0.5;
    }

    /**
     * Compare cost values (handles $ format)
     */
    compareCost(expected, actual) {
        if (!expected || !actual) return false;

        // Parse expected
        let expectedNum = parseFloat(String(expected).replace(/[$,\s]/g, ''));
        // Parse actual
        let actualNum = parseFloat(String(actual).replace(/[$,\s]/g, ''));

        if (isNaN(expectedNum) || isNaN(actualNum)) return false;

        return Math.abs(expectedNum - actualNum) < 0.001;
    }

    /**
     * Check if a value represents valid wastage percentage
     */
    isValidWastage(value, expectedPercent) {
        if (value === null || value === undefined || value === '') {
            return false;
        }

        const strValue = String(value).trim().toUpperCase();

        // Check for percentage string
        if (strValue === `${expectedPercent}%` || strValue.includes(`${expectedPercent}%`)) {
            return true;
        }

        // Check for numeric value
        const numValue = parseFloat(strValue.replace(/[%,\s]/g, ''));
        if (!isNaN(numValue)) {
            // Accept both 5 and 0.05 for 5%
            const decimalValue = expectedPercent / 100;
            if (Math.abs(numValue - expectedPercent) < 0.001 || Math.abs(numValue - decimalValue) < 0.001) {
                return true;
            }
        }

        return false;
    }

    /**
     * Generate HTML for results display
     */
    generateResultsHTML(results) {
        if (!results || results.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">Peak Performance Validation Ready</p>
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
                        oninput="window.peakPerformanceProcessor.searchByFilename(this.value)"
                    />
                </div>
                <button onclick="window.peakPerformanceProcessor.exportToPDF()" class="export-btn">
                    Export
                </button>
            </div>
        `;

        for (const fileResult of results) {
            const fabricWastage = fileResult.results.fabricWastageCheck;
            const standardItems = fileResult.results.standardItemsCheck;

            // Count valid checks
            let validCount = 0;
            let totalChecks = 1; // Fabric wastage

            if (standardItems.found && standardItems.items) {
                totalChecks += standardItems.items.filter(i => i.found).length;
            }

            if (fabricWastage.found && fabricWastage.isValid) validCount++;
            if (standardItems.found && standardItems.items) {
                validCount += standardItems.items.filter(i => i.found && i.isValid).length;
            }

            html += `<div class="file-result-group">`;

            // File summary
            html += `
                <div class="file-summary-box">
                    <strong>File:</strong> ${fileResult.fileName}<br>
                    <strong>Summary:</strong> ${validCount} out of ${totalChecks} checks passed
                </div>
            `;

            // Validation results table
            html += `
                <table class="results-table" style="table-layout: fixed; width: 100%;">
                    <thead>
                        <tr class="header-labels-row">
                            <th style="width: 50%;">Validation Check</th>
                            <th style="width: 50%;">Value</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            // Fabric/Yarn Wastage row
            if (fabricWastage.found) {
                const allCells = [];
                fabricWastage.validCells.forEach(cell => {
                    allCells.push(`<span style="color: #065f46; font-weight: 600;">${cell.value}</span>`);
                });
                fabricWastage.invalidCells.forEach(cell => {
                    allCells.push(`<span style="color: #991b1b; font-weight: 600;">${cell.value}</span> <span style="font-size: 0.85em; color: #849bba;">(Expected: 5%)</span>`);
                });

                if (allCells.length > 0) {
                    html += `
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">Fabric/Yarn Wastage (5%)</td>
                            <td style="padding: 0.875rem 1rem; text-align: center;">
                                ${allCells.join(', ')}
                            </td>
                        </tr>
                    `;
                } else {
                    html += `
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">Fabric/Yarn Wastage (5%)</td>
                            <td style="padding: 0.875rem 1rem; text-align: center;">
                                <span style="color: #6b7280; font-weight: 600;">No data in range</span>
                            </td>
                        </tr>
                    `;
                }
            } else {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">Fabric/Yarn Wastage (5%)</td>
                        <td style="padding: 0.875rem 1rem; text-align: center; color: #991b1b; font-weight: 600;">${fabricWastage.message || 'Not found'}</td>
                    </tr>
                `;
            }

            // Standard Items section with separate columns
            if (standardItems.found && standardItems.items) {
                // Close the previous table and start a new one for standard items
                html += `
                    </tbody>
                </table>

                <h3 style="margin: 1.5rem 0 0.75rem 0; font-size: 1.1rem; color: #2b4a6c;">Standard Items Check</h3>
                <table class="results-table" style="table-layout: fixed; width: 100%;">
                    <thead>
                        <tr class="header-labels-row">
                            <th style="width: 18%;">Material Description</th>
                            <th style="width: 12%;">Supplier</th>
                            <th style="width: 10%;">Item #</th>
                            <th style="width: 12%;">Garment Part</th>
                            <th style="width: 10%;">Yield</th>
                            <th style="width: 10%;">Wastage</th>
                            <th style="width: 14%;">FOB Cost</th>
                            <th style="width: 14%;">CIF Cost</th>
                        </tr>
                    </thead>
                    <tbody>
                `;

                for (const item of standardItems.items) {
                    const itemName = item.standardItem.materialDesc;

                    if (!item.found) {
                        html += `
                            <tr style="border-bottom: 1px solid #e0e8f0;">
                                <td style="padding: 0.5rem; font-weight: 600; font-size: 0.85rem;">${itemName}</td>
                                <td colspan="7" style="padding: 0.5rem; text-align: center; color: #991b1b; font-weight: 600;">Not found in file</td>
                            </tr>
                        `;
                    } else {
                        // Build cell values for each column
                        const getCheckByLabel = (label) => item.checks.find(c => c.label === label);

                        const supplierCheck = getCheckByLabel('Supplier');
                        const itemNumCheck = getCheckByLabel('Supplier Item #');
                        const partCheck = getCheckByLabel('Garment Part');
                        const yieldCheck = getCheckByLabel('Yield');
                        const wastageCheck = getCheckByLabel('Wastage');
                        const fobCheck = getCheckByLabel('FOB Unit Cost');
                        const cifCheck = getCheckByLabel('CIF Unit Cost');

                        const formatCell = (check) => {
                            if (!check) return '<span style="color: #6b7280;">-</span>';
                            const displayValue = check.actual || '-';
                            if (check.isValid) {
                                return `<span style="color: #065f46; font-weight: 600;">${displayValue}</span>`;
                            } else {
                                return `<span style="color: #991b1b; font-weight: 600;">${displayValue}</span><br><span style="font-size: 0.75em; color: #849bba;">(Expected: ${check.expected})</span>`;
                            }
                        };

                        html += `
                            <tr style="border-bottom: 1px solid #e0e8f0;">
                                <td style="padding: 0.5rem; font-weight: 600; font-size: 0.85rem;">${itemName}</td>
                                <td style="padding: 0.5rem; text-align: center; font-size: 0.85rem;">${formatCell(supplierCheck)}</td>
                                <td style="padding: 0.5rem; text-align: center; font-size: 0.85rem;">${formatCell(itemNumCheck)}</td>
                                <td style="padding: 0.5rem; text-align: center; font-size: 0.85rem;">${formatCell(partCheck)}</td>
                                <td style="padding: 0.5rem; text-align: center; font-size: 0.85rem;">${formatCell(yieldCheck)}</td>
                                <td style="padding: 0.5rem; text-align: center; font-size: 0.85rem;">${formatCell(wastageCheck)}</td>
                                <td style="padding: 0.5rem; text-align: center; font-size: 0.85rem;">${formatCell(fobCheck)}</td>
                                <td style="padding: 0.5rem; text-align: center; font-size: 0.85rem;">${formatCell(cifCheck)}</td>
                            </tr>
                        `;
                    }
                }
            }

            html += `
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
     * Search by filename - filters file result groups based on filename
     */
    searchByFilename(searchTerm) {
        const resultsContainer = document.getElementById('results-v10');
        if (!resultsContainer) return;

        const fileGroups = resultsContainer.querySelectorAll('.file-result-group');

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

    /**
     * Export results to PDF using the unified Export.js module
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

        const config = this.createPeakPerformanceConfig(this.bcbdResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }

    /**
     * Create PDF export configuration for Peak Performance
     */
    createPeakPerformanceConfig(results) {
        return {
            title: 'Peak Performance Validation Report',
            brandName: 'Peak Performance',
            files: results.map(fileResult => {
                const checks = [];
                const fabricWastage = fileResult.results.fabricWastageCheck;
                const standardItems = fileResult.results.standardItemsCheck;

                // Add fabric wastage check
                if (fabricWastage.found) {
                    checks.push({
                        label: 'Fabric/Yarn Wastage (5%)',
                        isValid: fabricWastage.isValid,
                        details: `${fabricWastage.validCells.length} valid, ${fabricWastage.invalidCells.length} invalid`,
                        validCells: fabricWastage.validCells.map(c => c.cellAddress),
                        invalidCells: fabricWastage.invalidCells.map(c => `${c.cellAddress} (${c.value})`)
                    });
                } else {
                    checks.push({
                        label: 'Fabric/Yarn Wastage (5%)',
                        isValid: false,
                        details: fabricWastage.message || 'Not found'
                    });
                }

                // Add standard items checks
                if (standardItems.found && standardItems.items) {
                    for (const item of standardItems.items) {
                        if (item.found) {
                            checks.push({
                                label: item.standardItem.materialDesc,
                                isValid: item.isValid,
                                details: item.checks.map(c => `${c.label}: ${c.isValid ? 'OK' : c.actual}`).join(', ')
                            });
                        }
                    }
                }

                return {
                    fileName: fileResult.fileName,
                    checks: checks,
                    validCount: checks.filter(c => c.isValid).length,
                    totalCount: checks.length
                };
            })
        };
    }
}

// Initialize the processor
window.peakPerformanceProcessor = new PeakPerformanceProcessor();

// Initialize when V10 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    const v10Tab = document.querySelector('[data-tab="v10"]');
    if (v10Tab) {
        v10Tab.addEventListener('click', () => {
            window.peakPerformanceProcessor.initialize();
        });
    }

    // If V10 tab is already active on load, initialize immediately
    const v10TabContent = document.getElementById('tab-v10');
    if (v10TabContent && v10TabContent.classList.contains('active')) {
        window.peakPerformanceProcessor.initialize();
    }
});
