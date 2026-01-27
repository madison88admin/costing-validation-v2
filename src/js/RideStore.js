/**
 * Ride Store Cost Breakdown Processor
 * Validates BCBD files against predefined validation rules
 */

class RideStoreProcessor {
    constructor() {
        this.validationRules = {
            A13: { field: 'supplier', expected: 'Madison 88 Ltd.' },
            A14: { field: 'factory', expected: 'Madison 88 Ltd.' },
            A15: { field: 'countryOfOrigin', expected: 'Indonesia' },
            B13: { field: 'supplierValue', expected: 'Madison 88 Ltd.' },
            B14: { field: 'factoryValue', expected: 'Madison 88 Ltd.' },
            B15: { field: 'countryValue', expected: 'Indonesia' }
        };
        this.fileResults = [];
    }

    async initialize() {
        this.displayValidationRules();
    }

    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v17');
        if (!obDropZone) return;

        let html = `
            <div class="burton-cost-container">
                <h2 class="drop-title">Ride Store Validation Criteria</h2>
                <p class="drop-subtitle">Expected values to check</p>
                <div class="burton-cost-items">
        `;

        // Display validation rules
        html += `
            <div class="burton-cost-item">
                <div class="burton-item-line"><strong>Cell A13:</strong> supplier</div>
                <div class="burton-item-line"><strong>Cell A14:</strong> Factory</div>
                <div class="burton-item-line"><strong>Cell A15:</strong> Country of origin</div>
            </div>
            <div class="burton-cost-item">
                <div class="burton-item-line"><strong>Cell B13:</strong> Madison 88 Ltd.</div>
                <div class="burton-item-line"><strong>Cell B14:</strong> Madison 88 Ltd.</div>
                <div class="burton-item-line"><strong>Cell B15:</strong> Indonesia</div>
            </div>
        `;

        html += `
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
                    cellResults: [],
                    sectionResults: []
                });
            }
        }

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
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

                    const validationResults = this.validateData(jsonData);

                    resolve({
                        fileName: file.name,
                        cellResults: validationResults.cellResults,
                        sectionResults: validationResults.sectionResults,
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

    validateData(jsonData) {
        const cellResults = [];
        const sectionResults = [];

        // Check A13-B13 (Supplier - merged)
        const a13Value = jsonData[12] && jsonData[12][0] ? jsonData[12][0].toString().trim() : '';
        const b13Value = jsonData[12] && jsonData[12][1] ? jsonData[12][1].toString().trim() : '';
        const isA13Valid = a13Value.toLowerCase() === 'supplier:';
        const isB13Valid = b13Value === 'Madison 88 Ltd.';
        cellResults.push({
            label: 'Supplier',
            cell: 'A13-B13',
            expected: 'supplier: | Madison 88 Ltd.',
            actual: `${a13Value} | ${b13Value}`,
            isValid: isA13Valid && isB13Valid
        });

        // Check A14-B14 (Factory - merged)
        const a14Value = jsonData[13] && jsonData[13][0] ? jsonData[13][0].toString().trim() : '';
        const b14Value = jsonData[13] && jsonData[13][1] ? jsonData[13][1].toString().trim() : '';
        const isA14Valid = a14Value.toLowerCase() === 'factory:';
        const isB14Valid = b14Value === 'Madison 88 Ltd.';
        cellResults.push({
            label: 'Factory',
            cell: 'A14-B14',
            expected: 'Factory: | Madison 88 Ltd.',
            actual: `${a14Value} | ${b14Value}`,
            isValid: isA14Valid && isB14Valid
        });

        // Check A15-B15 (Country - merged)
        const a15Value = jsonData[14] && jsonData[14][0] ? jsonData[14][0].toString().trim() : '';
        const b15Value = jsonData[14] && jsonData[14][1] ? jsonData[14][1].toString().trim() : '';
        const isA15Valid = a15Value.toLowerCase() === 'country of origin:';
        const isB15Valid = b15Value === 'Indonesia';
        cellResults.push({
            label: 'Country of Origin',
            cell: 'A15-B15',
            expected: 'Country of origin: | Indonesia',
            actual: `${a15Value} | ${b15Value}`,
            isValid: isA15Valid && isB15Valid
        });

        // Validate sections
        const fabricSection = this.validateSection(jsonData, 'FABRIC/Main Material', 'TOTAL FABRIC', '5%');
        sectionResults.push(fabricSection);

        const trimsSection = this.validateSection(jsonData, 'TRIMS & ACCESSORIES', 'TOTAL TRIMS & ACCESSORIES', '3%');
        sectionResults.push(trimsSection);

        const packagingResult = this.validateLabelsAndPackaging(jsonData);
        sectionResults.push(packagingResult.section);

        // Validate Overhead
        const overheadResult = this.validateOverhead(jsonData);
        if (overheadResult) {
            cellResults.push(overheadResult);
        }

        // Validate Profit
        const profitResults = this.validateProfit(jsonData);
        cellResults.push(...profitResults);

        return { cellResults, sectionResults };
    }

    /**
     * Validate a section with a specific wastage percentage
     * @param {Array} jsonData - Excel data
     * @param {string} startMarker - Text in Column A that marks the start
     * @param {string} endMarker - Text in Column F that marks the end
     * @param {string} expectedWastage - Expected value for Column H
     * @returns {Object} Section result with valid and invalid cells
     */
    validateSection(jsonData, startMarker, endMarker, expectedWastage) {
        let inSection = false;
        const validCells = [];
        const invalidCells = [];

        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const colA = row[0] ? row[0].toString().trim() : '';
            const colF = row[5] ? row[5].toString().trim() : ''; // Column F (index 5)
            const colH = row[7] ? row[7].toString().trim() : ''; // Column H (index 7)

            // Check if we're starting the section
            if (colA.includes(startMarker)) {
                inSection = true;
                continue; // Skip the header row itself
            }

            // If we're in the section, validate each row
            if (inSection) {
                // Check if we've reached the end marker
                if (colF.includes(endMarker)) {
                    break; // Stop checking this section
                }

                // Skip empty rows
                if (!colH || colH === '') continue;

                // Validate Column H for this row
                const normalizedColH = this.normalizePercentage(colH);
                const normalizedExpected = this.normalizePercentage(expectedWastage);
                const isValid = normalizedColH === normalizedExpected;

                const cellRef = `H${i + 1}`;

                if (isValid) {
                    validCells.push({ cell: cellRef, value: colH });
                } else {
                    invalidCells.push({ cell: cellRef, value: colH });
                }
            }
        }

        return {
            label: startMarker,
            expectedValue: expectedWastage,
            validCells: validCells,
            invalidCells: invalidCells,
            isValid: invalidCells.length === 0 && validCells.length > 0,
            sectionFound: validCells.length > 0 || invalidCells.length > 0
        };
    }

    /**
     * Validate LABELS & PACKAGING section with special rules
     * @returns {Object} Object containing section result with General Packaging sub-validations
     */
    validateLabelsAndPackaging(jsonData) {
        let inSection = false;
        const validCells = [];
        const invalidCells = [];

        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const colA = row[0] ? row[0].toString().trim() : '';
            const colB = row[1] ? row[1].toString().trim() : ''; // Column B (index 1)
            const colF = row[5] ? row[5].toString().trim() : ''; // Column F (index 5)
            const colG = row[6] ? row[6].toString().trim() : ''; // Column G (index 6)
            const colH = row[7] ? row[7].toString().trim() : ''; // Column H (index 7)

            // Check if we're starting the section
            if (colA.includes('LABELS & PACKAGING')) {
                inSection = true;
                continue; // Skip the header row itself
            }

            // If we're in the section, validate each row
            if (inSection) {
                // Check if we've reached the end marker
                if (colF.includes('TOTAL TRIMS')) {
                    break; // Stop checking this section
                }

                // Skip empty rows
                if (!colH || colH === '') continue;

                // Validate Column H = 3%
                const normalizedColH = this.normalizePercentage(colH);
                const normalizedExpected = this.normalizePercentage('3%');
                const isValidWastage = normalizedColH === normalizedExpected;

                const cellRef = `H${i + 1}`;
                let gpDetails = null;

                // Additional checks for "General Packaging" rows
                if (colA.includes('General Packaging')) {
                    // Column B must have a supplier (not empty)
                    const isValidSupplier = colB !== '';

                    // Column F must be "Pcs"
                    const isValidUnit = colF.toLowerCase() === 'pcs';

                    // Column G must be 1
                    const normalizedColG = colG.toString().trim();
                    const isValidQty = normalizedColG === '1' || normalizedColG === '1.0' || normalizedColG === '1.00' || normalizedColG === '1.000';

                    gpDetails = {
                        supplier: { value: colB || 'Empty', isValid: isValidSupplier },
                        unit: { value: colF, isValid: isValidUnit },
                        quantity: { value: colG, isValid: isValidQty },
                        allValid: isValidSupplier && isValidUnit && isValidQty
                    };
                }

                if (isValidWastage) {
                    validCells.push({ cell: cellRef, value: colH, gpDetails });
                } else {
                    invalidCells.push({ cell: cellRef, value: colH, gpDetails });
                }
            }
        }

        return {
            section: {
                label: 'LABELS & PACKAGING',
                expectedValue: '3%',
                validCells: validCells,
                invalidCells: invalidCells,
                isValid: invalidCells.length === 0 && validCells.length > 0 &&
                         validCells.every(cell => !cell.gpDetails || cell.gpDetails.allValid) &&
                         invalidCells.every(cell => !cell.gpDetails || cell.gpDetails.allValid),
                sectionFound: validCells.length > 0 || invalidCells.length > 0
            }
        };
    }

    /**
     * Validate Overhead row
     * @param {Array} jsonData - Excel data
     * @returns {Object|null} Validation result or null if not found
     */
    validateOverhead(jsonData) {
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const colA = row[0] ? row[0].toString().trim() : '';

            if (colA.includes('Overhead: Rent, electricity, transport etc.')) {
                const colJ = row[9] ? row[9].toString().trim() : ''; // Column J (index 9)
                const normalizedColJ = parseFloat(colJ);
                const isValid = normalizedColJ === 0.60 || normalizedColJ === 0.6;

                return {
                    label: 'Overhead',
                    cell: `J${i + 1}`,
                    expected: '0.60',
                    actual: colJ || 'Empty',
                    isValid: isValid
                };
            }
        }
        return null;
    }

    /**
     * Validate Profit row
     * @param {Array} jsonData - Excel data
     * @returns {Array} Array of validation results
     */
    validateProfit(jsonData) {
        const results = [];

        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const colA = row[0] ? row[0].toString().trim() : '';

            if (colA.toLowerCase() === 'profit') {
                // Check Column D
                const colD = row[3] ? row[3].toString().trim() : ''; // Column D (index 3)
                const isDValid = colD === '% of FOB';
                results.push({
                    label: 'Profit - Unit',
                    cell: `D${i + 1}`,
                    expected: '% of FOB',
                    actual: colD || 'Empty',
                    isValid: isDValid
                });

                // Check Column H
                const colH = row[7] ? row[7].toString().trim() : ''; // Column H (index 7)
                const normalizedColH = this.normalizePercentage(colH);
                const normalizedExpected = this.normalizePercentage('8.43%');
                const isHValid = normalizedColH === normalizedExpected;
                results.push({
                    label: 'Profit - Percentage',
                    cell: `H${i + 1}`,
                    expected: '8.43%',
                    actual: colH || 'Empty',
                    isValid: isHValid
                });

                break; // Only check the first "Profit" row
            }
        }

        return results;
    }

    /**
     * Normalize percentage values for comparison
     * Handles formats like "5%", "5.0%", "0.05", etc.
     */
    normalizePercentage(value) {
        if (!value) return '';

        const str = value.toString().trim().toLowerCase();

        // If it's already a percentage string like "5%"
        if (str.includes('%')) {
            return str;
        }

        // If it's a decimal like 0.05, convert to percentage
        const num = parseFloat(str);
        if (!isNaN(num)) {
            if (num < 1) {
                // Assume it's a decimal (0.05 = 5%)
                return `${num * 100}%`;
            } else {
                // Assume it's already a percentage value (5 = 5%)
                return `${num}%`;
            }
        }

        return str;
    }

    /**
     * Format field value with color coding (like 511)
     */
    formatFieldValue(result) {
        if (!result.actual || result.actual === '') {
            return `<span style="color: #991b1b; font-weight: 600;">Empty</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${result.expected}</span>`;
        }

        if (result.isValid) {
            return `<span style="color: #065f46; font-weight: 600;">${result.actual}</span>`;
        } else {
            return `<span style="color: #991b1b; font-weight: 600;">${result.actual}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${result.expected}</span>`;
        }
    }

    /**
     * Format wastage cells for section display (like 511)
     */
    formatWastageCells(section) {
        if (!section.sectionFound) {
            return `<span style="color: #64748b;">Section not found in file</span>`;
        }

        const validItems = section.validCells;
        const invalidItems = section.invalidCells;

        let html = '';

        // Show valid cells in green with GP details if applicable
        if (validItems.length > 0) {
            const cellsHtml = validItems.map(item => {
                let cellText = item.cell;

                if (item.gpDetails) {
                    // Add General Packaging details inline
                    const supplierColor = item.gpDetails.supplier.isValid ? '#065f46' : '#991b1b';
                    const unitColor = item.gpDetails.unit.isValid ? '#065f46' : '#991b1b';
                    const qtyColor = item.gpDetails.quantity.isValid ? '#065f46' : '#991b1b';
                    const gpLabelColor = item.gpDetails.allValid ? '#065f46' : '#991b1b';

                    cellText += ` <span style="font-size: 0.85em; color: #ffffff;">(</span>` +
                        `<span style="font-size: 0.85em; color: ${gpLabelColor};">General Packaging:</span>` +
                        `<span style="font-size: 0.85em;"> <span style="color: ${supplierColor};">${item.gpDetails.supplier.value}</span>` +
                        `<span style="color: #ffffff;">, </span>` +
                        `<span style="color: ${unitColor};">${item.gpDetails.unit.value}</span>` +
                        `<span style="color: #ffffff;">, </span>` +
                        `<span style="color: ${qtyColor};">${item.gpDetails.quantity.value}</span></span>` +
                        `<span style="font-size: 0.85em; color: #ffffff;">)</span>`;
                }

                return cellText;
            }).join(', ');

            html += `<span style="color: #065f46; font-weight: 600;">${cellsHtml}</span>`;
        }

        // Show invalid cells in red with GP details if applicable
        if (invalidItems.length > 0) {
            if (validItems.length > 0) {
                html += '<br>';
            }

            const cellsHtml = invalidItems.map(item => {
                let cellText = `${item.cell}: ${item.value}`;

                if (item.gpDetails) {
                    // Add General Packaging details inline
                    const supplierColor = item.gpDetails.supplier.isValid ? '#065f46' : '#991b1b';
                    const unitColor = item.gpDetails.unit.isValid ? '#065f46' : '#991b1b';
                    const qtyColor = item.gpDetails.quantity.isValid ? '#065f46' : '#991b1b';
                    const gpLabelColor = item.gpDetails.allValid ? '#065f46' : '#991b1b';

                    cellText += ` <span style="font-size: 0.85em; color: #ffffff;">(</span>` +
                        `<span style="font-size: 0.85em; color: ${gpLabelColor};">General Packaging:</span>` +
                        `<span style="font-size: 0.85em;"> <span style="color: ${supplierColor};">${item.gpDetails.supplier.value}</span>` +
                        `<span style="color: #ffffff;">, </span>` +
                        `<span style="color: ${unitColor};">${item.gpDetails.unit.value}</span>` +
                        `<span style="color: #ffffff;">, </span>` +
                        `<span style="color: ${qtyColor};">${item.gpDetails.quantity.value}</span></span>` +
                        `<span style="font-size: 0.85em; color: #ffffff;">)</span>`;
                }

                return `<span style="color: #991b1b; font-weight: 600;">${cellText}</span>`;
            }).join(', ');

            html += cellsHtml;
        }

        // If no items found
        if (validItems.length === 0 && invalidItems.length === 0) {
            html = `<span style="color: #64748b;">No items found in section</span>`;
        }

        return html;
    }

    generateResultsHTML(results) {
        if (!results || results.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">Ride Store Validation Ready</p>
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
                        oninput="window.rideStoreProcessor.searchByFilename(this.value)"
                    />
                </div>
                <button onclick="window.rideStoreProcessor.exportToPDF()" class="export-btn">
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
                        <strong style="color: #991b1b;">Error:</strong> ${fileResult.error}
                    </div>
                `;
                html += `</div>`;
                continue;
            }

            // Calculate totals
            const cellValidCount = fileResult.cellResults.filter(r => r.isValid).length;
            const cellTotalCount = fileResult.cellResults.length;
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
                <table class="results-table">
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

            for (const item of fileResult.cellResults) {
                const statusIcon = item.isValid ? '✓' : '✗';
                const statusColor = item.isValid ? '#065f46' : '#991b1b';
                const statusText = item.isValid ? 'VALID' : 'INVALID';

                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">${item.label}</td>
                        <td style="padding: 0.875rem 1rem;">${item.cell}</td>
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

            // Section validation results
            html += `
                <div style="margin-top: 1.5rem; margin-bottom: 1.5rem;">
                <table class="results-table" style="table-layout: fixed; width: 100%;">
                    <thead>
                        <tr class="header-labels-row">
                            <th style="width: 250px;">Section</th>
                            <th>Wastage% (Column H)</th>
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

        const config = window.pdfExporter.createRideStoreConfig(this.fileResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }

    searchByFilename(searchTerm) {
        const fileGroups = document.querySelectorAll('.file-result-group');

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
window.rideStoreProcessor = new RideStoreProcessor();

// Auto-initialize when V17 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    const v17Tabs = document.querySelectorAll('[data-tab="v17"]');
    v17Tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            window.rideStoreProcessor.initialize();
        });
    });

    const v17TabContent = document.getElementById('tab-v17');
    if (v17TabContent && v17TabContent.classList.contains('active')) {
        window.rideStoreProcessor.initialize();
    }
});
