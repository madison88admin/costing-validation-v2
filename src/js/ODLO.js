/**
 * ODLO Cost Breakdown Processor
 * Validates BCBD files for ODLO brand requirements
 *
 * Validation Logic:
 * 1. Check column A for the word "Category" → Column B should be "Accessories"
 * 2. Check column A for the word "Garment Maker" → Column B should be "Madison 88 (USD)"
 * 3. Check for TRIMS section in column A, find THD-10005 rows and validate columns B, E, F, G
 * 4. Column B validations (check Column G for expected values):
 *    - Labour Sewing minutes → 0.1
 *    - Labour Heat transfer pressing → 0
 *    - Labour Seam sealing → 0.4
 *    - Knitting minutes → 0.05
 *    - Linking minutes → 0.15
 *    - Overhead in % → 10% to 15%
 *    - Profit in % → 4% to 8%
 */

class ODLOProcessor {
    constructor() {
        this.validationRules = [
            {
                name: 'Category',
                markerColumn: 0, // Column A
                marker: 'category',
                checkColumn: 1, // Column B
                expected: 'Accessories',
                exactMatch: false // Case-insensitive comparison
            },
            {
                name: 'Garment Maker',
                markerColumn: 0, // Column A
                marker: 'garment maker',
                checkColumn: 1, // Column B
                expected: 'Madison 88 (USD)',
                exactMatch: false // Case-insensitive comparison
            }
        ];
    }

    async initialize() {
        this.displayValidationRules();
    }

    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v22');
        if (!obDropZone) return;

        let html = `
            <div class="burton-cost-container">
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>ODLO Validation Rules:</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Rule 1 - Category Check:</strong></div>
                        <div class="burton-item-line">• Find "Category" in Column A → Column B = <strong>Accessories</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Rule 2 - Garment Maker Check:</strong></div>
                        <div class="burton-item-line">• Find "Garment Maker" in Column A → Column B = <strong>Madison 88 (USD)</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Rule 3 - Trims Check:</strong></div>
                        <div class="burton-item-line">• Find "TRIMS" in Column A, then find "THD-10005" rows</div>
                        <div class="burton-item-line">• Validate: Col B, E, F, G values</div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Rule 4 - Column B → G Checks:</strong></div>
                        <div class="burton-item-line">• Labour Sewing minutes → <strong>0.1</strong></div>
                        <div class="burton-item-line">• Labour Heat transfer pressing → <strong>0</strong></div>
                        <div class="burton-item-line">• Labour Seam sealing → <strong>0.4</strong></div>
                        <div class="burton-item-line">• Knitting minutes → <strong>0.05</strong></div>
                        <div class="burton-item-line">• Linking minutes → <strong>0.15</strong></div>
                        <div class="burton-item-line">• Overhead in % → <strong>10% - 15%</strong></div>
                        <div class="burton-item-line">• Profit in % → <strong>4% - 8%</strong></div>
                    </div>
                </div>
            </div>
        `;

        obDropZone.innerHTML = html;
    }

    async processFiles(bcbdFiles) {
        const results = [];

        for (const file of bcbdFiles) {
            try {
                const fileResult = await this.processFile(file);
                results.push(fileResult);
            } catch (error) {
                console.error(`Error processing file ${file.name}:`, error);
                results.push({
                    fileName: file.name,
                    error: error.message,
                    checks: []
                });
            }
        }

        // Store results for export
        this.fileResults = results;

        return this.generateResultsHTML(results);
    }

    async processFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    // Process first sheet (or we can process all sheets)
                    const firstSheetName = workbook.SheetNames[0];
                    const sheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                    const checks = this.validateSheet(jsonData);

                    resolve({
                        fileName: file.name,
                        sheetName: firstSheetName,
                        checks: checks,
                        error: null
                    });
                } catch (error) {
                    reject(error);
                }
            };

            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsArrayBuffer(file);
        });
    }

    validateSheet(jsonData) {
        const results = [];

        for (const rule of this.validationRules) {
            const result = {
                name: rule.name,
                expected: rule.expected,
                found: false,
                rowNumber: -1,
                actual: '',
                isValid: false,
                markerColumn: this.getColumnLetter(rule.markerColumn),
                checkColumn: this.getColumnLetter(rule.checkColumn)
            };

            // Scan through all rows looking for the marker
            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                const markerCell = row[rule.markerColumn] ? String(row[rule.markerColumn]).trim() : '';
                const markerCellLower = markerCell.toLowerCase();

                // Check if this row contains the marker
                if (markerCellLower === rule.marker.toLowerCase()) {
                    result.found = true;
                    result.rowNumber = i + 1; // 1-indexed for display

                    // Get the value in the check column
                    const actualValue = row[rule.checkColumn] ? String(row[rule.checkColumn]).trim() : '';
                    result.actual = actualValue || 'Empty';

                    // Validate the value
                    if (rule.exactMatch) {
                        result.isValid = actualValue === rule.expected;
                    } else {
                        result.isValid = actualValue.toLowerCase() === rule.expected.toLowerCase();
                    }

                    break; // Found the first occurrence, stop searching
                }
            }

            results.push(result);
        }

        // Add Trims validation (can return multiple results)
        const trimsResults = this.validateTrimsSection(jsonData);
        for (const trimsResult of trimsResults) {
            results.push(trimsResult);
        }

        // Add Column B → Column G validations
        const colBValidations = this.validateColumnBRules(jsonData);
        for (const colBResult of colBValidations) {
            results.push(colBResult);
        }

        return results;
    }

    validateColumnBRules(jsonData) {
        const rules = [
            { name: 'Labour Sewing minutes', marker: 'labour sewing minutes', expected: '0.1', type: 'exact' },
            { name: 'Labour Heat transfer pressing', marker: 'labour heat transfer pressing', expected: '0', type: 'exact' },
            { name: 'Labour Seam sealing', marker: 'labour seam sealing', expected: '0.4', type: 'exact' },
            { name: 'Knitting minutes', marker: 'knitting minutes', expected: '0.05', type: 'exact' },
            { name: 'Linking minutes', marker: 'linking minutes', expected: '0.15', type: 'exact' },
            { name: 'Overhead in %', marker: 'overhead in %', expected: '10% - 15%', type: 'range', min: 0.10, max: 0.15 },
            { name: 'Profit in %', marker: 'profit in %', expected: '4% - 8%', type: 'range', min: 0.04, max: 0.08 }
        ];

        const results = [];

        for (const rule of rules) {
            const result = {
                name: rule.name,
                expected: rule.expected,
                found: false,
                rowNumber: -1,
                actual: '',
                isValid: false,
                markerColumn: 'B',
                checkColumn: 'G',
                isColumnBRule: true // Flag to identify these rules for table display
            };

            // Scan column B for the marker
            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                const cellB = row[1] ? String(row[1]).trim().toLowerCase() : '';

                if (cellB === rule.marker) {
                    result.found = true;
                    result.rowNumber = i + 1;

                    // Get column G value
                    const cellG = row[6] !== undefined && row[6] !== null ? String(row[6]).trim() : '';

                    // Validate based on type
                    if (rule.type === 'exact') {
                        result.actual = cellG || 'Empty';
                        result.isValid = cellG === rule.expected;
                    } else if (rule.type === 'range') {
                        // Parse the value as a number (handle percentage format)
                        let numValue = parseFloat(cellG);
                        if (isNaN(numValue)) {
                            result.actual = cellG || 'Empty';
                            result.isValid = false;
                        } else {
                            // If value is <= 1, it's a decimal (e.g., 0.10 for 10%)
                            // Multiply by 100 for display as percentage
                            if (numValue <= 1) {
                                result.actual = (numValue * 100).toFixed(2) + '%';
                                result.isValid = numValue >= rule.min && numValue <= rule.max;
                            } else {
                                // Already a percentage (e.g., 10 for 10%)
                                result.actual = numValue.toFixed(2) + '%';
                                result.isValid = (numValue / 100) >= rule.min && (numValue / 100) <= rule.max;
                            }
                        }
                    }

                    break;
                }
            }

            results.push(result);
        }

        return results;
    }

    validateTrimsSection(jsonData) {
        const results = [];

        // Step 1: Find TRIMS in column A
        let trimsRowIndex = -1;
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const cellA = row[0] ? String(row[0]).trim().toLowerCase() : '';
            if (cellA === 'trims') {
                trimsRowIndex = i;
                break;
            }
        }

        if (trimsRowIndex === -1) {
            results.push({
                name: 'Trims',
                expected: 'THD-10005 | Sewing thread dye to match | p | 1 | 0.008',
                found: false,
                rowNumber: -1,
                actual: 'TRIMS section not found',
                isValid: false,
                markerColumn: 'A',
                checkColumn: 'B, E, F, G'
            });
            return results;
        }

        // Step 2: Scan rows below TRIMS until we find Total in column G
        // Collect ALL THD-10005 entries
        let trimsCount = 0;
        for (let i = trimsRowIndex + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            const cellG = row[6] ? String(row[6]).trim().toLowerCase() : '';

            // Stop if we hit Total in column G
            if (cellG === 'total') {
                break;
            }

            const cellA = row[0] ? String(row[0]).trim() : '';

            // Check if this row has THD-10005 in column A
            if (cellA.toUpperCase() === 'THD-10005') {
                trimsCount++;

                // Get values from columns B, E, F, G
                const cellB = row[1] ? String(row[1]).trim() : '';
                const cellE = row[4] ? String(row[4]).trim() : '';
                const cellF = row[5] ? String(row[5]).trim() : '';
                const cellGVal = row[6] ? String(row[6]).trim() : '';

                // Validate each column individually
                const cellValidations = [
                    { value: cellA, isValid: true, expected: 'THD-10005' },
                    { value: cellB, isValid: cellB.toLowerCase() === 'sewing thread dye to match', expected: 'Sewing thread dye to match' },
                    { value: cellE, isValid: cellE.toLowerCase() === 'p', expected: 'p' },
                    { value: cellF, isValid: cellF === '1', expected: '1' },
                    { value: cellGVal, isValid: cellGVal === '0.008', expected: '0.008' }
                ];

                const allValid = cellValidations.every(cv => cv.isValid);

                results.push({
                    name: `Trims (${trimsCount})`,
                    expected: 'THD-10005 | Sewing thread dye to match | p | 1 | 0.008',
                    found: true,
                    rowNumber: i + 1, // 1-indexed for display
                    actual: `${cellA} | ${cellB} | ${cellE} | ${cellF} | ${cellGVal}`,
                    isValid: allValid,
                    markerColumn: 'A',
                    checkColumn: 'B, E, F, G',
                    cellValidations: cellValidations // Store individual cell validations
                });
            }
        }

        // If no THD-10005 found in the section
        if (results.length === 0) {
            results.push({
                name: 'Trims',
                expected: 'THD-10005 | Sewing thread dye to match | p | 1 | 0.008',
                found: false,
                rowNumber: -1,
                actual: 'THD-10005 not found in TRIMS section',
                isValid: false,
                markerColumn: 'A',
                checkColumn: 'B, E, F, G'
            });
        }

        return results;
    }

    getColumnLetter(index) {
        const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
        if (index < 26) {
            return letters[index];
        }
        // Handle columns beyond Z (AA, AB, etc.)
        return letters[Math.floor(index / 26) - 1] + letters[index % 26];
    }

    generateResultsHTML(results) {
        let html = '';

        // Add export button at the top
        html += `
            <div style="margin-bottom: 15px; display: flex; justify-content: flex-end; align-items: center;">
                <button onclick="window.odloProcessor.exportToPDF()" class="export-btn">
                    Export
                </button>
            </div>
        `;

        for (const fileResult of results) {
            html += `<div class="file-result-group">`;

            if (fileResult.error) {
                html += `
                    <div class="file-summary-box">
                        <strong>File:</strong> ${fileResult.fileName}<br>
                        <span style="color: #991b1b;">Error: ${fileResult.error}</span>
                    </div>
                `;
            } else {
                // Calculate summary
                const totalChecks = fileResult.checks.length;
                const validChecks = fileResult.checks.filter(c => c.isValid).length;
                const foundChecks = fileResult.checks.filter(c => c.found).length;

                html += `
                    <div class="file-summary-box">
                        <strong>File:</strong> ${fileResult.fileName}<br>
                        <strong>Sheet:</strong> ${fileResult.sheetName}<br>
                        <strong>Summary:</strong> ${validChecks} out of ${foundChecks} found checks are valid (${totalChecks} total rules)
                    </div>
                `;

                // Create results table with 3 columns
                html += `
                    <table class="results-table" style="table-layout: fixed; width: 100%;">
                        <thead>
                            <tr class="header-labels-row">
                                <th style="width: 250px;">Check Name</th>
                                <th>Value</th>
                                <th style="width: 200px;">Expected</th>
                            </tr>
                        </thead>
                        <tbody>
                `;

                for (const check of fileResult.checks) {
                    let valueHTML = '';
                    let expectedHTML = '';

                    if (!check.found) {
                        valueHTML = `<span style="color: #b45309; font-weight: 600;">Not Found</span>`;
                        expectedHTML = `<span style="color: #849bba;">${check.expected}</span>`;
                    } else if (check.cellValidations) {
                        // Handle Trims with individual cell coloring
                        const coloredValues = check.cellValidations.map(cv => {
                            const color = cv.isValid ? '#065f46' : '#991b1b';
                            return `<span style="color: ${color}; font-weight: 600;">${cv.value || 'Empty'}</span>`;
                        }).join(' | ');
                        valueHTML = coloredValues;

                        // Show expected values for invalid cells only
                        if (!check.isValid) {
                            const invalidCells = check.cellValidations
                                .filter(cv => !cv.isValid)
                                .map(cv => `${cv.expected}`)
                                .join(', ');
                            expectedHTML = `<span style="color: #849bba;">${invalidCells}</span>`;
                        } else {
                            expectedHTML = `<span style="color: #065f46;">✓</span>`;
                        }
                    } else if (check.isValid) {
                        valueHTML = `<span style="color: #065f46; font-weight: 600;">${check.actual}</span>`;
                        expectedHTML = `<span style="color: #065f46;">✓</span>`;
                    } else {
                        valueHTML = `<span style="color: #991b1b; font-weight: 600;">${check.actual}</span>`;
                        expectedHTML = `<span style="color: #849bba;">${check.expected}</span>`;
                    }

                    html += `
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">${check.name}</td>
                            <td style="padding: 0.875rem 1rem;">${valueHTML}</td>
                            <td style="padding: 0.875rem 1rem;">${expectedHTML}</td>
                        </tr>
                    `;
                }

                html += `
                        </tbody>
                    </table>
                `;
            }

            html += '</div>';
        }

        return html;
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

        if (!this.fileResults || this.fileResults.length === 0) {
            alert('No results to export. Please generate results first.');
            return;
        }

        const config = window.pdfExporter.createODLOConfig(this.fileResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }
}

// Initialize processor
window.odloProcessor = new ODLOProcessor();
