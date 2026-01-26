/**
 * 511 Processing Logic
 * Validates specific cells in Buyer CBD files against expected values
 */

class Processor511 {
    constructor() {
        this.bcbdResults = [];
        this.validationRules = {
            factory: {
                label: 'FACTORY',
                labelCell: 'D7',
                valueCell: 'E7',
                expectedValue: 'PT Ujump Indonesia'
            },
            coo: {
                label: 'COO:',
                labelCell: 'D8',
                valueCell: 'E8',
                expectedValue: 'Indonesia'
            },
            remarks: {
                label: 'Remarks',
                valueCell: 'E15',
                expectedValue: 'if under style minimum, we will request upcharge, profit includes OH cost'
            }
        };
        this.sectionRules = {
            fabrics: {
                label: 'Fabrics',
                startKeyword: 'a.', // More specific - starts with A.
                startKeyword2: 'fabrics',
                endKeyword: 'total fabric cost',
                wastageColumn: 'J',
                expectedWastage: 0.05, // 5%
                expectedDisplay: '5%'
            },
            trims: {
                label: 'Trims',
                startKeyword: 'b.', // More specific - starts with B.
                startKeyword2: 'trims',
                endKeyword: 'total trims cost',
                wastageColumn: 'J',
                expectedWastage: 0.03, // 3%
                expectedDisplay: '3%'
            },
            packaging: {
                label: 'Packaging',
                startKeyword: 'c.', // More specific - starts with C.
                startKeyword2: 'labels',
                endKeyword: 'total packing cost',
                wastageColumn: 'J',
                expectedWastage: 0.03, // 3%
                expectedDisplay: '3%'
            }
        };
    }

    /**
     * Initialize - Display validation rules in the OB drop zone
     */
    initialize() {
        this.displayValidationRules();
        console.log('511 Processor initialized');
    }

    /**
     * Display validation rules in the OB drop zone
     */
    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v16');
        if (!obDropZone) return;

        let contentHTML = `
            <div class="burton-cost-container">
                <div class="burton-cost-header">
                    <h3>511 Validation Rules</h3>
                    <p class="cost-subtitle">Expected values to validate</p>
                </div>
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>FACTORY (D7):</strong> PT Ujump Indonesia (E7)</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>COO: (D8):</strong> Indonesia (E8)</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Remarks (E15):</strong></div>
                        <div class="burton-item-line" style="font-size: 0.85em; color: #7a92ab;">if under style minimum, we will request upcharge, profit includes OH cost</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Fabrics Wastage% (Column J):</strong> 5%</div>
                        <div class="burton-item-line" style="font-size: 0.85em; color: #7a92ab;">Between "A. FABRICS..." and "TOTAL FABRIC COST"</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Trims Wastage% (Column J):</strong> 3%</div>
                        <div class="burton-item-line" style="font-size: 0.85em; color: #7a92ab;">Between "B. TRIMS..." and "TOTAL TRIMS COST"</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Packaging Wastage% (Column J):</strong> 3%</div>
                        <div class="burton-item-line" style="font-size: 0.85em; color: #7a92ab;">Between "C. LABELS..." and "TOTAL PACKING COST"</div>
                    </div>
                </div>
            </div>
        `;

        obDropZone.innerHTML = contentHTML;
    }

    /**
     * Convert column letter to index (A=0, B=1, etc.)
     */
    columnToIndex(column) {
        let index = 0;
        for (let i = 0; i < column.length; i++) {
            index = index * 26 + column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
        }
        return index - 1;
    }

    /**
     * Convert row index to cell reference (e.g., row 6, col G -> G7)
     */
    indicesToCell(rowIndex, colLetter) {
        return `${colLetter}${rowIndex + 1}`;
    }

    /**
     * Convert cell reference to row and column indices (e.g., "D7" -> {row: 6, col: 3})
     */
    cellToIndices(cellRef) {
        const match = cellRef.match(/^([A-Z]+)(\d+)$/);
        if (!match) return null;

        const colStr = match[1];
        const rowNum = parseInt(match[2], 10);

        let colIndex = 0;
        for (let i = 0; i < colStr.length; i++) {
            colIndex = colIndex * 26 + colStr.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
        }

        return {
            row: rowNum - 1, // 0-indexed
            col: colIndex - 1 // 0-indexed
        };
    }

    /**
     * Get cell value from parsed sheet data
     */
    getCellValue(jsonData, cellRef) {
        const indices = this.cellToIndices(cellRef);
        if (!indices) return null;

        if (indices.row < jsonData.length && indices.col < (jsonData[indices.row]?.length || 0)) {
            const value = jsonData[indices.row][indices.col];
            return value !== undefined && value !== null ? value.toString().trim() : '';
        }
        return '';
    }

    /**
     * Validate wastage% in a section between start and end markers
     * @param {Array} jsonData - Parsed Excel data
     * @param {Object} sectionRule - Section validation rules
     * @param {number} searchFromRow - Row to start searching from (to prevent overlap)
     * @returns {Object} - Validation result with endRow for next section
     */
    validateSectionWastage(jsonData, sectionRule, searchFromRow = 0) {
        const wastageColIndex = this.columnToIndex(sectionRule.wastageColumn);
        const colA = this.columnToIndex('A');
        const colB = this.columnToIndex('B');

        // Find start and end rows by scanning ONLY columns A and B for section headers
        // Start searching from searchFromRow to prevent overlap with previous sections
        let startRow = -1;
        let endRow = -1;

        for (let i = searchFromRow; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;

            // Only check columns A and B for section markers
            const cellA = row[colA] ? row[colA].toString().trim().toLowerCase() : '';
            const cellB = row[colB] ? row[colB].toString().trim().toLowerCase() : '';
            const combinedText = cellA + ' ' + cellB;

            // For start marker, need both keywords (e.g., "a." AND "fabrics")
            if (startRow === -1) {
                if (combinedText.includes(sectionRule.startKeyword) &&
                    combinedText.includes(sectionRule.startKeyword2)) {
                    startRow = i;
                }
            }

            // For end marker, just need the end keyword
            if (startRow !== -1 && combinedText.includes(sectionRule.endKeyword)) {
                endRow = i;
                break;
            }
        }

        const validCells = [];
        const invalidCells = [];

        if (startRow === -1 || endRow === -1) {
            return {
                label: sectionRule.label,
                sectionFound: false,
                validCells: validCells,
                invalidCells: invalidCells,
                expectedValue: sectionRule.expectedDisplay,
                isValid: false,
                endRow: searchFromRow // Return same row so next section can continue searching
            };
        }

        // Check each row between start and end (exclusive of markers)
        for (let i = startRow + 1; i < endRow; i++) {
            const row = jsonData[i];
            if (!row) continue;

            // Get wastage column value
            const cellValue = row[wastageColIndex];
            if (cellValue === undefined || cellValue === null || cellValue === '') continue;

            const cellRef = this.indicesToCell(i, sectionRule.wastageColumn);
            let numericValue;

            // Parse the value - could be percentage string or decimal
            if (typeof cellValue === 'number') {
                numericValue = cellValue;
            } else {
                const strValue = cellValue.toString().trim();
                if (strValue.endsWith('%')) {
                    numericValue = parseFloat(strValue) / 100;
                } else {
                    numericValue = parseFloat(strValue);
                }
            }

            if (isNaN(numericValue)) continue;

            // Compare with expected wastage (allow small tolerance)
            const isMatch = Math.abs(numericValue - sectionRule.expectedWastage) < 0.001;

            // Format actual value for display
            let displayValue;
            if (numericValue < 1) {
                displayValue = (numericValue * 100).toFixed(0) + '%';
            } else {
                displayValue = numericValue.toFixed(0) + '%';
            }

            if (isMatch) {
                validCells.push({ cell: cellRef, value: displayValue });
            } else {
                invalidCells.push({ cell: cellRef, value: displayValue });
            }
        }

        return {
            label: sectionRule.label,
            sectionFound: true,
            validCells: validCells,
            invalidCells: invalidCells,
            expectedValue: sectionRule.expectedDisplay,
            isValid: invalidCells.length === 0 && validCells.length > 0,
            endRow: endRow // Return endRow so next section starts after this one
        };
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
                    results: validationResults.cellResults,
                    sectionResults: validationResults.sectionResults
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
     * Validate file against rules
     */
    validateFile(jsonData) {
        const cellResults = [];
        const sectionResults = [];

        // Validate FACTORY (D7 label, E7 value)
        const factoryLabel = this.getCellValue(jsonData, 'D7');
        const factoryValue = this.getCellValue(jsonData, 'E7');
        const factoryExpected = this.validationRules.factory.expectedValue;
        const factoryLabelMatch = factoryLabel.toLowerCase().includes('factory');
        const factoryValid = factoryValue.toLowerCase() === factoryExpected.toLowerCase();

        cellResults.push({
            label: 'FACTORY',
            labelCell: 'D7',
            valueCell: 'E7',
            actualLabel: factoryLabel,
            actualValue: factoryValue,
            expectedValue: factoryExpected,
            labelFound: factoryLabelMatch,
            isValid: factoryValid
        });

        // Validate COO (D8 label, E8 value)
        const cooLabel = this.getCellValue(jsonData, 'D8');
        const cooValue = this.getCellValue(jsonData, 'E8');
        const cooExpected = this.validationRules.coo.expectedValue;
        const cooLabelMatch = cooLabel.toLowerCase().includes('coo');
        const cooValid = cooValue.toLowerCase() === cooExpected.toLowerCase();

        cellResults.push({
            label: 'COO:',
            labelCell: 'D8',
            valueCell: 'E8',
            actualLabel: cooLabel,
            actualValue: cooValue,
            expectedValue: cooExpected,
            labelFound: cooLabelMatch,
            isValid: cooValid
        });

        // Validate Remarks (E15 value)
        const remarksValue = this.getCellValue(jsonData, 'E15');
        const remarksExpected = this.validationRules.remarks.expectedValue;
        const remarksValid = remarksValue.toLowerCase() === remarksExpected.toLowerCase();

        cellResults.push({
            label: 'Remarks',
            labelCell: null,
            valueCell: 'E15',
            actualLabel: null,
            actualValue: remarksValue,
            expectedValue: remarksExpected,
            labelFound: true,
            isValid: remarksValid
        });

        // Validate section wastages - process in sequence to prevent overlap
        // Each section starts searching from the end of the previous section
        const fabricsResult = this.validateSectionWastage(jsonData, this.sectionRules.fabrics, 0);
        sectionResults.push(fabricsResult);

        // Trims starts after Fabrics ends
        const trimsResult = this.validateSectionWastage(jsonData, this.sectionRules.trims, fabricsResult.endRow + 1);
        sectionResults.push(trimsResult);

        // Packaging starts after Trims ends
        const packagingResult = this.validateSectionWastage(jsonData, this.sectionRules.packaging, trimsResult.endRow + 1);
        sectionResults.push(packagingResult);

        return { cellResults, sectionResults };
    }

    /**
     * Format field value with color coding and expected value display
     */
    formatFieldValue(result) {
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
     * Format wastage cells for display (Prana-style)
     */
    formatWastageCells(section) {
        if (!section.sectionFound) {
            return `<span style="color: #64748b;">Section not found in file</span>`;
        }

        const validItems = section.validCells;
        const invalidItems = section.invalidCells;

        let html = '';

        // Show valid cells in green (just cell addresses)
        if (validItems.length > 0) {
            const validCells = validItems.map(item => item.cell).join(', ');
            html += `<span style="color: #065f46; font-weight: 600;">${validCells}</span>`;
        }

        // Show invalid cells in red (cell address with value)
        if (invalidItems.length > 0) {
            if (validItems.length > 0) {
                html += '<br>';
            }
            const invalidCells = invalidItems.map(item => {
                return `<span style="color: #991b1b; font-weight: 600;">${item.cell}: ${item.value}</span>`;
            }).join(', ');
            html += invalidCells;
        }

        // If no items found
        if (validItems.length === 0 && invalidItems.length === 0) {
            html = `<span style="color: #64748b;">No items found in section</span>`;
        }

        return html;
    }

    /**
     * Generate HTML for results display
     */
    generateResultsHTML(results) {
        if (!results || results.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">511 Validation Ready</p>
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
                        oninput="window.processor511.searchByFilename(this.value)"
                    />
                </div>
                <button onclick="window.processor511.exportToPDF()" class="export-btn">
                    Export
                </button>
            </div>
        `;

        for (const fileResult of results) {
            html += `<div class="file-result-group">`;

            // Calculate total validations
            const cellValidCount = fileResult.results.filter(r => r.isValid).length;
            const cellTotalCount = fileResult.results.length;
            const sectionValidCount = fileResult.sectionResults.filter(r => r.isValid).length;
            const sectionTotalCount = fileResult.sectionResults.length;

            html += `
                <div class="file-summary-box">
                    <strong>File:</strong> ${fileResult.fileName}<br>
                    <strong>Cell Validations:</strong> ${cellValidCount} out of ${cellTotalCount} passed<br>
                    <strong>Section Validations:</strong> ${sectionValidCount} out of ${sectionTotalCount} passed
                </div>
            `;

            // Cell validation table
            html += `
                <table id="v16ResultsTable" class="results-table">
                    <thead>
                        <tr class="header-labels-row">
                            <th>Field</th>
                            <th>Cell</th>
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
                        <td style="padding: 0.875rem 1rem;">${item.valueCell}</td>
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
            `;

            // Section validation results - table format like Prana
            html += `
                <div style="margin-top: 1.5rem; margin-bottom: 1.5rem;">
                <table class="results-table" style="table-layout: fixed; width: 100%;">
                    <thead>
                        <tr class="header-labels-row">
                            <th style="width: 200px;">Section</th>
                            <th>Wastage% (Column J)</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            for (const section of fileResult.sectionResults) {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">${section.label}<br><span style="font-size: 0.85em; color: #64748b;">Expected: ${section.expectedValue}</span></td>
                        <td style="padding: 0.875rem 1rem;">
                            ${this.formatWastageCells(section)}
                        </td>
                    </tr>
                `;
            }

            html += `
                    </tbody>
                </table>
                </div>
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

        const config = window.pdfExporter.create511Config(this.bcbdResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }

    /**
     * Search by filename
     */
    searchByFilename(searchTerm) {
        const fileGroups = document.querySelectorAll('#tab-v16 .file-result-group');

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
window.processor511 = new Processor511();

// Auto-initialize when V16 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    const v16Tabs = document.querySelectorAll('[data-tab="v16"]');
    v16Tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            window.processor511.initialize();
        });
    });

    const v16TabContent = document.getElementById('tab-v16');
    if (v16TabContent && v16TabContent.classList.contains('active')) {
        window.processor511.initialize();
    }
});
