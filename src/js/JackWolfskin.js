/**
 * Jack Wolfskin Processing Logic
 * Validates specific cells in Buyer CBD files against expected values
 */

class JackWolfskinProcessor {
    constructor() {
        this.bcbdResults = [];
        this.validationRules = {
            supplier: {
                label: 'Supplier',
                searchColumn: 'B',
                searchValue: 'Supplier',
                valueColumn: 'C',
                expectedValue: 'Madison88'
            },
            overheadCost: {
                label: 'Overhead Cost',
                searchColumn: 'J',
                searchValue: 'Overhead Cost',
                valueColumn: 'L',
                expectedValue: 0.35
            },
            profit: {
                label: 'Profit',
                searchColumn: 'J',
                searchValue: 'Profit',
                valueColumn: 'L',
                expectedRange: { min: 0.15, max: 0.25 }
            }
        };
    }

    /**
     * Initialize - Display validation rules in the OB drop zone
     */
    initialize() {
        this.displayValidationRules();
        console.log('Jack Wolfskin Processor initialized');
    }

    /**
     * Display validation rules in the OB drop zone
     */
    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v15');
        if (!obDropZone) return;

        let contentHTML = `
            <div class="burton-cost-container">
                <div class="burton-cost-header">
                    <h3>Jack Wolfskin Validation Rules</h3>
                    <p class="cost-subtitle">Expected values to validate</p>
                </div>
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Supplier:</strong> Madison88</div>
                        <div class="burton-item-line" style="font-size: 0.85em; color: #7a92ab;">Column B "Supplier" → Column C value</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Overhead Cost:</strong> 0.35</div>
                        <div class="burton-item-line" style="font-size: 0.85em; color: #7a92ab;">Column J "Overhead Cost" → Column L value</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Profit:</strong> 0.15 - 0.25</div>
                        <div class="burton-item-line" style="font-size: 0.85em; color: #7a92ab;">Column J "Profit" → Column L value (range)</div>
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

                    // Get the last sheet
                    const lastSheetName = workbook.SheetNames[workbook.SheetNames.length - 1];
                    const sheet = workbook.Sheets[lastSheetName];
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
     * Convert column letter to index (A=0, B=1, C=2, etc.)
     */
    columnToIndex(column) {
        let index = 0;
        for (let i = 0; i < column.length; i++) {
            index = index * 26 + column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
        }
        return index - 1;
    }

    /**
     * Validate file against rules
     */
    validateFile(jsonData) {
        const results = [];

        // Validate Supplier (Column B → Column C)
        const supplierResult = this.findAndValidate(
            jsonData,
            this.columnToIndex('B'),
            'Supplier',
            this.columnToIndex('C'),
            'Madison88',
            false
        );
        results.push({
            label: 'Supplier',
            ...supplierResult
        });

        // Validate Overhead Cost (Column J → Column L)
        const overheadResult = this.findAndValidate(
            jsonData,
            this.columnToIndex('J'),
            'Overhead Cost',
            this.columnToIndex('L'),
            0.35,
            true
        );
        results.push({
            label: 'Overhead Cost',
            ...overheadResult
        });

        // Validate Profit (Column J → Column L, range 0.15-0.25)
        const profitResult = this.findAndValidateRange(
            jsonData,
            this.columnToIndex('J'),
            'Profit',
            this.columnToIndex('L'),
            0.15,
            0.25
        );
        results.push({
            label: 'Profit',
            ...profitResult
        });

        return results;
    }

    /**
     * Find a value in a column and validate the corresponding value in another column
     */
    findAndValidate(jsonData, searchColIndex, searchValue, valueColIndex, expectedValue, isNumeric) {
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const cellValue = row[searchColIndex] ? row[searchColIndex].toString().trim() : '';

            // Check if the cell contains or starts with the search value
            if (cellValue.toLowerCase().includes(searchValue.toLowerCase()) ||
                cellValue.toLowerCase().startsWith(searchValue.toLowerCase())) {

                const actualValue = row[valueColIndex] ? row[valueColIndex].toString().trim() : '';

                let isValid = false;
                if (isNumeric) {
                    const actualNum = parseFloat(actualValue);
                    const expectedNum = parseFloat(expectedValue);
                    isValid = !isNaN(actualNum) && Math.abs(actualNum - expectedNum) < 0.001;
                } else {
                    isValid = actualValue.toLowerCase() === expectedValue.toString().toLowerCase();
                }

                return {
                    found: true,
                    row: i + 1,
                    actualValue: actualValue,
                    expectedValue: expectedValue.toString(),
                    isValid: isValid
                };
            }
        }

        return {
            found: false,
            row: null,
            actualValue: null,
            expectedValue: expectedValue.toString(),
            isValid: false
        };
    }

    /**
     * Find a value and validate against a range
     */
    findAndValidateRange(jsonData, searchColIndex, searchValue, valueColIndex, minValue, maxValue) {
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const cellValue = row[searchColIndex] ? row[searchColIndex].toString().trim() : '';

            // Check if the cell contains or equals the search value (exact match for "Profit" to avoid matching "Overhead Cost")
            if (cellValue.toLowerCase() === searchValue.toLowerCase() ||
                (cellValue.toLowerCase().includes(searchValue.toLowerCase()) &&
                 !cellValue.toLowerCase().includes('overhead'))) {

                const actualValue = row[valueColIndex] ? row[valueColIndex].toString().trim() : '';
                const actualNum = parseFloat(actualValue);

                const isValid = !isNaN(actualNum) && actualNum >= minValue && actualNum <= maxValue;

                return {
                    found: true,
                    row: i + 1,
                    actualValue: actualValue,
                    expectedValue: `${minValue} - ${maxValue}`,
                    isValid: isValid,
                    isRange: true
                };
            }
        }

        return {
            found: false,
            row: null,
            actualValue: null,
            expectedValue: `${minValue} - ${maxValue}`,
            isValid: false,
            isRange: true
        };
    }

    /**
     * Format field value with color coding and expected value display
     */
    formatFieldValue(result) {
        if (!result.found) {
            return `<span style="color: #991b1b; font-weight: 600;">Not Found</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${result.expectedValue}</span>`;
        }

        if (result.actualValue === '' || result.actualValue === null) {
            return `<span style="color: #991b1b; font-weight: 600;">Empty</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${result.expectedValue}</span>`;
        }

        if (result.isValid) {
            return `<span style="color: #065f46; font-weight: 600;">${result.actualValue}</span>`;
        } else {
            return `<span style="color: #991b1b; font-weight: 600;">${result.actualValue}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${result.expectedValue}</span>`;
        }
    }

    /**
     * Generate HTML for results display
     */
    generateResultsHTML(results) {
        if (!results || results.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">Jack Wolfskin Validation Ready</p>
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
                        oninput="window.jackWolfskinProcessor.searchByFilename(this.value)"
                    />
                </div>
                <button onclick="window.jackWolfskinProcessor.exportToPDF()" class="export-btn">
                    Export
                </button>
            </div>
        `;

        for (const fileResult of results) {
            html += `<div class="file-result-group">`;

            // Add summary
            const totalItems = fileResult.results.length;
            const validItems = fileResult.results.filter(r => r.isValid).length;

            html += `
                <div class="file-summary-box">
                    <strong>File:</strong> ${fileResult.fileName}<br>
                    <strong>Summary:</strong> ${validItems} out of ${totalItems} validations passed
                </div>
            `;

            // Create comparison table
            html += `
                <table id="v15ResultsTable" class="results-table">
                    <thead>
                        <tr class="header-labels-row">
                            <th>Field</th>
                            <th>BCBD Value</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            for (const item of fileResult.results) {
                const statusIcon = item.isValid ? '✓' : '✗';
                const statusColor = item.isValid ? '#065f46' : '#991b1b';
                const statusText = item.isValid ? 'VALID' : 'INVALID';

                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">${item.label}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(item)}</td>
                        <td style="padding: 0.875rem 1rem;">
                            <span style="color: ${statusColor}; font-weight: 600;">${statusIcon} ${statusText}</span>
                        </td>
                    </tr>
                `;
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

        const config = window.pdfExporter.createJackWolfskinConfig(this.bcbdResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }

    /**
     * Search by filename
     */
    searchByFilename(searchTerm) {
        const fileGroups = document.querySelectorAll('#tab-v15 .file-result-group');

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
window.jackWolfskinProcessor = new JackWolfskinProcessor();

// Auto-initialize when V15 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    const v15Tabs = document.querySelectorAll('[data-tab="v15"]');
    v15Tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            window.jackWolfskinProcessor.initialize();
        });
    });

    const v15TabContent = document.getElementById('tab-v15');
    if (v15TabContent && v15TabContent.classList.contains('active')) {
        window.jackWolfskinProcessor.initialize();
    }
});
