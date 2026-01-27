/**
 * Foot Asylum (V18) Processing Logic
 * Validates Fabrics section in Buyer CBD files
 *
 * Validation Rules (within Fabrics section):
 * - Column D: Must be TRUE
 * - Column P: Must be "USD"
 * - Column W: Must be exactly 5%
 * - Column AG: Must be exactly 0.5
 * - Column AK: Must be exactly 0.1
 * - Column AL: Must be exactly 10%
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
                    <h3>Foot Asylum Validation Rules</h3>
                    <p class="cost-subtitle">Fabrics Section Validation</p>
                </div>
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Section:</strong> Fabrics (between "Fabrics (...)" and "Trims (")</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Column D - Main Material:</strong> TRUE</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Column P - Supplier Currency:</strong> USD</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Column W - Wastage %:</strong> 5%</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Column AG - Overhead Cost:</strong> 0.5</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Column AK - Testing Cost:</strong> 0.1</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Column AL - Profit % / Total FOB:</strong> 10%</div>
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
     * Validate a single value against expected value
     */
    validateValue(actualValue, rule) {
        if (actualValue === undefined || actualValue === null || actualValue === '') {
            return { isValid: false, displayValue: 'Empty', isEmpty: true };
        }

        let isValid = false;
        let displayValue = actualValue.toString();

        // Handle boolean TRUE/FALSE
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
                isValid = Math.abs(numericValue - rule.expectedValue) < 0.0001;
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
     * Validate file against rules - returns row-based results
     */
    validateFile(jsonData) {
        const results = {
            sectionFound: false,
            startRow: -1,
            endRow: -1,
            rows: [] // Array of row validation results
        };

        // Find Fabrics section
        const section = this.findFabricsSection(jsonData);
        results.sectionFound = section.sectionFound;
        results.startRow = section.startRow;
        results.endRow = section.endRow;

        if (!section.sectionFound) {
            return results;
        }

        // Validate each row in the Fabrics section
        for (let i = section.startRow; i < section.endRow; i++) {
            const row = jsonData[i];
            if (!row) continue;

            // Skip completely empty rows - check if any of the validation columns have data
            const hasData = Object.values(this.validationRules).some(rule => {
                const cellValue = row[rule.columnIndex];
                return cellValue !== undefined && cellValue !== null && cellValue !== '';
            });
            if (!hasData) continue;

            // Build row result
            const rowResult = {
                rowNumber: i + 1, // 1-indexed for display
                columns: {}
            };

            let rowHasAnyData = false;

            // Validate each column rule for this row
            for (const [key, rule] of Object.entries(this.validationRules)) {
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
                results.rows.push(rowResult);
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

            // Calculate validation counts
            let totalValid = 0;
            let totalInvalid = 0;

            for (const rowResult of fileResult.results.rows) {
                for (const [key, cellData] of Object.entries(rowResult.columns)) {
                    if (!cellData.isEmpty) {
                        if (cellData.isValid) {
                            totalValid++;
                        } else {
                            totalInvalid++;
                        }
                    }
                }
            }

            const sectionStatus = fileResult.results.sectionFound
                ? `<span style="color: #065f46;">Found (rows ${fileResult.results.startRow + 1}-${fileResult.results.endRow})</span>`
                : `<span style="color: #991b1b;">Not found</span>`;

            html += `
                <div class="file-summary-box">
                    <strong>File:</strong> ${fileResult.fileName}<br>
                    <strong>Fabrics Section:</strong> ${sectionStatus}<br>
                    <strong>Validations:</strong> ${totalValid} passed, ${totalInvalid} failed
                </div>
            `;

            if (!fileResult.results.sectionFound) {
                html += `
                    <div style="padding: 1rem; background: #fee; border-radius: 8px; margin-top: 1rem;">
                        <p style="color: #991b1b; font-weight: 600;">Fabrics section not found in file.</p>
                        <p style="color: #721c24; font-size: 0.9em;">Looking for "Fabrics (...)" in Column A to start and "Trims (" to end.</p>
                    </div>
                </div>`;
                continue;
            }

            if (fileResult.results.rows.length === 0) {
                html += `
                    <div style="padding: 1rem; background: #fef3c7; border-radius: 8px; margin-top: 1rem;">
                        <p style="color: #92400e; font-weight: 600;">No data rows found in Fabrics section.</p>
                    </div>
                </div>`;
                continue;
            }

            // Build header row with column labels
            const ruleKeys = Object.keys(this.validationRules);

            html += `
                <table id="v18ResultsTable" class="results-table">
                    <thead>
                        <tr class="header-labels-row">
                            <th style="width: 60px;">Row</th>
            `;

            // Add column headers
            for (const key of ruleKeys) {
                const rule = this.validationRules[key];
                html += `<th>${rule.label}<br><span style="font-size: 0.75em; font-weight: normal; color: #64748b;">(${rule.column}) ${rule.expectedDisplay}</span></th>`;
            }

            html += `
                        </tr>
                    </thead>
                    <tbody>
            `;

            // Add data rows
            for (const rowResult of fileResult.results.rows) {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">${rowResult.rowNumber}</td>
                `;

                for (const key of ruleKeys) {
                    const cellData = rowResult.columns[key];
                    html += `<td style="padding: 0.875rem 1rem;">${this.formatCellValue(cellData)}</td>`;
                }

                html += `</tr>`;
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
