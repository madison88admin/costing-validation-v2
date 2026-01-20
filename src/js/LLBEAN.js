/**
 * LLBEAN Processing Logic (V6)
 * Validates Buyer CBD files against LLBEAN criteria
 */

class LLBEANProcessor {
    constructor() {
        this.bcbdFiles = [];
        this.bcbdResults = [];
        this.validationRules = {
            b5Keywords: ['beanie', 'socks', 'scarf'],
            trimsBox: {
                item: 'Box',
                supplier: 'Local',
                consumption: 1.00,
                unitPrice: 0.06,
                totalCost: 0.06
            },
            totalFinancialCost: {
                beanie: 0.20,
                socks: 0.50,
                scarf: 0.40
            }
        };
    }

    /**
     * Initialize V6 - Display validation rules in the OB drop zone
     */
    initialize() {
        this.displayValidationRules();
        console.log('LLBEAN Processor initialized');
    }

    /**
     * Display validation rules in the OB drop zone (Burton-style)
     */
    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v6');
        if (!obDropZone) return;

        const box = this.validationRules.trimsBox;

        let html = `
            <div class="burton-cost-container">
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Cell B5 Keywords Check</strong></div>
                        <div class="burton-item-line">Must contain: ${this.validationRules.b5Keywords.join(', ')}</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Trims Section - Box Validation</strong></div>
                        <div class="burton-item-line"><strong>Item:</strong> ${box.item}</div>
                        <div class="burton-item-line"><strong>Supplier:</strong> ${box.supplier}</div>
                        <div class="burton-item-line"><strong>Consumption:</strong> ${box.consumption.toFixed(2)}</div>
                        <div class="burton-item-line"><strong>Unit Price:</strong> $${box.unitPrice.toFixed(2)}</div>
                        <div class="burton-item-line"><strong>Total Cost:</strong> $${box.totalCost.toFixed(2)}</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Total Financial Cost (Column H)</strong></div>
                        <div class="burton-item-line"><strong>If Beanie:</strong> $${this.validationRules.totalFinancialCost.beanie.toFixed(2)}</div>
                        <div class="burton-item-line"><strong>If Socks:</strong> $${this.validationRules.totalFinancialCost.socks.toFixed(2)}</div>
                        <div class="burton-item-line"><strong>If Scarf:</strong> $${this.validationRules.totalFinancialCost.scarf.toFixed(2)}</div>
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

                    // Check cell B5
                    const b5Check = this.checkCellB5(sheet);

                    // Check Trims section for Box
                    const trimsBoxCheck = this.checkTrimsBox(jsonData);

                    // Check Total Financial Cost based on B5 keyword
                    const totalFinancialCostCheck = this.checkTotalFinancialCost(jsonData, b5Check.foundKeywords);

                    resolve({
                        b5Check: b5Check,
                        trimsBoxCheck: trimsBoxCheck,
                        totalFinancialCostCheck: totalFinancialCostCheck
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
     * Check if cell B5 contains required keywords
     */
    checkCellB5(sheet) {
        // Get cell B5 value
        const cellB5 = sheet['B5'];
        const cellValue = cellB5 ? cellB5.v : '';
        const cellValueStr = String(cellValue).toLowerCase();

        console.log(`Cell B5 value: "${cellValue}"`);

        // Check if any of the keywords are present
        const foundKeywords = [];
        for (const keyword of this.validationRules.b5Keywords) {
            if (cellValueStr.includes(keyword.toLowerCase())) {
                foundKeywords.push(keyword);
            }
        }

        return {
            cellValue: cellValue,
            isValid: foundKeywords.length > 0,
            foundKeywords: foundKeywords,
            requiredKeywords: this.validationRules.b5Keywords
        };
    }

    /**
     * Check Trims section for Box row validation
     * Looks for "Trims" in column B, then finds "Box" in column C between Trims and Total Trims cost
     */
    checkTrimsBox(jsonData) {
        let trimsStartRow = -1;
        let trimsEndRow = -1;

        // Find Trims and Total Trims cost rows in column B (index 1)
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const colB = row[1] ? String(row[1]).trim().toLowerCase() : '';

            if (colB === 'trims') {
                trimsStartRow = i;
                console.log(`Found "Trims" at row ${i + 1}`);
            }

            if (colB === 'total trims cost' || colB === 'total trims') {
                trimsEndRow = i;
                console.log(`Found "Total Trims cost" at row ${i + 1}`);
            }
        }

        // If we didn't find the section, return not found
        if (trimsStartRow === -1) {
            return {
                found: false,
                trimsFound: false,
                boxFound: false,
                message: 'Trims section not found in column B'
            };
        }

        // If no end row found, use the rest of the file
        if (trimsEndRow === -1) {
            trimsEndRow = jsonData.length - 1;
        }

        // Find Box in column C (index 2) between Trims and Total Trims cost
        let boxData = null;
        for (let i = trimsStartRow; i <= trimsEndRow; i++) {
            const row = jsonData[i];
            const colC = row[2] ? String(row[2]).trim().toLowerCase() : '';

            if (colC === 'box') {
                boxData = {
                    rowNumber: i + 1,
                    item: row[2] ? String(row[2]).trim() : '',
                    supplier: row[4] ? String(row[4]).trim() : '',        // Column E
                    consumption: row[5],                                    // Column F
                    unitPrice: row[6],                                      // Column G
                    totalCost: row[7]                                       // Column H
                };
                console.log(`Found "Box" at row ${i + 1}:`, boxData);
                break;
            }
        }

        if (!boxData) {
            return {
                found: false,
                trimsFound: true,
                boxFound: false,
                message: 'Box not found in Trims section (column C)'
            };
        }

        // Validate Box row values
        const expected = this.validationRules.trimsBox;

        const supplierValid = this.compareField(expected.supplier, boxData.supplier);
        const consumptionValid = this.compareNumericField(expected.consumption, boxData.consumption);
        const unitPriceValid = this.compareNumericField(expected.unitPrice, boxData.unitPrice);
        const totalCostValid = this.compareNumericField(expected.totalCost, boxData.totalCost);

        const allValid = supplierValid === 'VALID' &&
                         consumptionValid === 'VALID' &&
                         unitPriceValid === 'VALID' &&
                         totalCostValid === 'VALID';

        return {
            found: true,
            trimsFound: true,
            boxFound: true,
            boxData: boxData,
            expected: expected,
            validation: {
                supplier: supplierValid,
                consumption: consumptionValid,
                unitPrice: unitPriceValid,
                totalCost: totalCostValid
            },
            isValid: allValid
        };
    }

    /**
     * Check Total Financial Cost based on B5 keyword
     * Looks for "Total Financial cost" in column B and validates column H value
     */
    checkTotalFinancialCost(jsonData, foundKeywords) {
        // Find Total Financial cost row in column B (index 1)
        let totalFinancialRow = -1;
        let actualValue = null;

        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const colB = row[1] ? String(row[1]).trim().toLowerCase() : '';

            if (colB === 'total financial cost' || colB === 'total financial') {
                totalFinancialRow = i;
                actualValue = row[7]; // Column H (index 7)
                console.log(`Found "Total Financial cost" at row ${i + 1}, Column H value: ${actualValue}`);
                break;
            }
        }

        if (totalFinancialRow === -1) {
            return {
                found: false,
                message: 'Total Financial cost not found in column B'
            };
        }

        // Determine expected value based on B5 keyword
        let expectedValue = null;
        let matchedKeyword = null;

        // Check which keyword was found (prioritize in order: beanie, socks, scarf)
        for (const keyword of foundKeywords) {
            const keywordLower = keyword.toLowerCase();
            if (keywordLower === 'beanie') {
                expectedValue = this.validationRules.totalFinancialCost.beanie;
                matchedKeyword = 'Beanie';
                break;
            } else if (keywordLower === 'socks') {
                expectedValue = this.validationRules.totalFinancialCost.socks;
                matchedKeyword = 'Socks';
                break;
            } else if (keywordLower === 'scarf') {
                expectedValue = this.validationRules.totalFinancialCost.scarf;
                matchedKeyword = 'Scarf';
                break;
            }
        }

        if (expectedValue === null) {
            return {
                found: true,
                rowNumber: totalFinancialRow + 1,
                actualValue: actualValue,
                expectedValue: null,
                matchedKeyword: null,
                message: 'No valid B5 keyword found to determine expected value',
                isValid: false
            };
        }

        // Validate the value
        const validationStatus = this.compareNumericField(expectedValue, actualValue);

        return {
            found: true,
            rowNumber: totalFinancialRow + 1,
            actualValue: actualValue,
            expectedValue: expectedValue,
            matchedKeyword: matchedKeyword,
            validation: validationStatus,
            isValid: validationStatus === 'VALID'
        };
    }

    /**
     * Compare text fields (case-insensitive)
     */
    compareField(expected, actual) {
        const exp = String(expected).toLowerCase().trim();
        const act = String(actual).toLowerCase().trim();
        return exp === act ? 'VALID' : 'INVALID';
    }

    /**
     * Compare numeric fields
     */
    compareNumericField(expected, actual) {
        // Clean and parse values
        const cleanActual = String(actual).replace(/[$,\s]/g, '');
        const expNum = parseFloat(expected);
        const actNum = parseFloat(cleanActual);

        if (isNaN(actNum)) {
            return 'INVALID';
        }

        // Round to 2 decimal places for comparison
        const expRounded = parseFloat(expNum.toFixed(2));
        const actRounded = parseFloat(actNum.toFixed(2));

        return expRounded === actRounded ? 'VALID' : 'INVALID';
    }

    /**
     * Format field value with color coding
     */
    formatFieldValue(expected, actual, status, isNumeric = false) {
        const color = status === 'VALID' ? '#065f46' : '#991b1b';

        const displayExpected = isNumeric ? this.formatNumber(expected) : expected;
        const displayActual = isNumeric ? this.formatNumber(actual) : actual;

        if (!actual || actual === '') {
            return `<span style="color: #991b1b; font-weight: 600;">Empty</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${displayExpected}</span>`;
        }

        if (status === 'VALID') {
            return `<span style="color: ${color}; font-weight: 600;">${displayActual}</span>`;
        } else {
            return `<span style="color: ${color}; font-weight: 600;">${displayActual}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${displayExpected}</span>`;
        }
    }

    /**
     * Format number to 2 decimal places
     */
    formatNumber(value) {
        const num = parseFloat(String(value).replace(/[$,\s]/g, ''));
        if (isNaN(num)) return value;
        return num.toFixed(2);
    }

    /**
     * Generate HTML for results display
     */
    generateResultsHTML(results) {
        if (!results || results.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">LLBEAN Validation Ready</p>
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
                        oninput="window.llbeanProcessor.searchByFilename(this.value)"
                    />
                </div>
                <button onclick="window.llbeanProcessor.exportToPDF()" class="export-btn">
                    Export
                </button>
            </div>
        `;

        for (const fileResult of results) {
            const b5 = fileResult.results.b5Check;
            const trimsBox = fileResult.results.trimsBoxCheck;
            const totalFinancial = fileResult.results.totalFinancialCostCheck;

            // Count valid checks
            let validCount = 0;
            let totalChecks = 3; // B5 check + Trims Box check + Total Financial Cost check

            if (b5.isValid) validCount++;
            if (trimsBox.found && trimsBox.isValid) validCount++;
            if (totalFinancial.found && totalFinancial.isValid) validCount++;

            html += `<div class="file-result-group">`;

            // File summary
            html += `
                <div class="file-summary-box">
                    <strong>File:</strong> ${fileResult.fileName}<br>
                    <strong>Summary:</strong> ${validCount} out of ${totalChecks} checks passed
                </div>
            `;

            // Validation results table with fixed column widths
            html += `
                <table class="results-table" style="table-layout: fixed; width: 100%;">
                    <thead>
                        <tr class="header-labels-row">
                            <th style="width: 22%;">Validation Check</th>
                            <th style="width: 15%;">Supplier</th>
                            <th style="width: 15%;">Consumption</th>
                            <th style="width: 15%;">Unit Price</th>
                            <th style="width: 15%;">Total Cost</th>
                            <th style="width: 18%;">Status</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            // B5 Keywords row - use individual cells with text-align center
            html += `
                <tr style="border-bottom: 1px solid #e0e8f0;">
                    <td style="padding: 0.875rem 1rem; font-weight: 600;">Cell B5 Keywords<br><span style="font-size: 0.85em; color: #849bba;">Value: ${b5.cellValue || 'Empty'}</span></td>
                    <td style="padding: 0.875rem 1rem; text-align: center;">-</td>
                    <td style="padding: 0.875rem 1rem; text-align: center;">
                        ${b5.foundKeywords.length > 0
                            ? `<strong style="color: #065f46;">${b5.foundKeywords.join(', ')}</strong>`
                            : `<span style="color: #991b1b;">${b5.requiredKeywords.join(', ')}</span>`}
                    </td>
                    <td style="padding: 0.875rem 1rem; text-align: center;">-</td>
                    <td style="padding: 0.875rem 1rem; text-align: center;">-</td>
                    <td style="padding: 0.875rem 1rem; text-align: center;">
                        ${b5.isValid
                            ? '<span style="color: #065f46; font-weight: 600;">VALID</span>'
                            : '<span style="color: #991b1b; font-weight: 600;">INVALID</span>'}
                    </td>
                </tr>
            `;

            // Trims Box row
            if (trimsBox.found && trimsBox.boxData) {
                const v = trimsBox.validation;
                const expected = trimsBox.expected;
                const actual = trimsBox.boxData;

                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">Trims - Box<br><span style="font-size: 0.85em; color: #849bba;">Row ${actual.rowNumber}</span></td>
                        <td style="padding: 0.875rem 1rem; text-align: center;">${this.formatFieldValue(expected.supplier, actual.supplier, v.supplier, false)}</td>
                        <td style="padding: 0.875rem 1rem; text-align: center;">${this.formatFieldValue(expected.consumption, actual.consumption, v.consumption, true)}</td>
                        <td style="padding: 0.875rem 1rem; text-align: center;">${this.formatFieldValue(expected.unitPrice, actual.unitPrice, v.unitPrice, true)}</td>
                        <td style="padding: 0.875rem 1rem; text-align: center;">${this.formatFieldValue(expected.totalCost, actual.totalCost, v.totalCost, true)}</td>
                        <td style="padding: 0.875rem 1rem; text-align: center;">
                            ${trimsBox.isValid
                                ? '<span style="color: #065f46; font-weight: 600;">VALID</span>'
                                : '<span style="color: #991b1b; font-weight: 600;">INVALID</span>'}
                        </td>
                    </tr>
                `;
            } else {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">Trims - Box</td>
                        <td style="padding: 0.875rem 1rem; text-align: center; color: #991b1b;">-</td>
                        <td style="padding: 0.875rem 1rem; text-align: center; color: #991b1b;">${trimsBox.message || 'Not found'}</td>
                        <td style="padding: 0.875rem 1rem; text-align: center; color: #991b1b;">-</td>
                        <td style="padding: 0.875rem 1rem; text-align: center; color: #991b1b;">-</td>
                        <td style="padding: 0.875rem 1rem; text-align: center;">
                            <span style="color: #991b1b; font-weight: 600;">NOT FOUND</span>
                        </td>
                    </tr>
                `;
            }

            // Total Financial Cost row - use individual cells
            if (totalFinancial.found && totalFinancial.expectedValue !== null) {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">Total Financial Cost<br><span style="font-size: 0.85em; color: #849bba;">Row ${totalFinancial.rowNumber} (${totalFinancial.matchedKeyword})</span></td>
                        <td style="padding: 0.875rem 1rem; text-align: center;">-</td>
                        <td style="padding: 0.875rem 1rem; text-align: center;">-</td>
                        <td style="padding: 0.875rem 1rem; text-align: center;">-</td>
                        <td style="padding: 0.875rem 1rem; text-align: center;">${this.formatFieldValue(totalFinancial.expectedValue, totalFinancial.actualValue, totalFinancial.validation, true)}</td>
                        <td style="padding: 0.875rem 1rem; text-align: center;">
                            ${totalFinancial.isValid
                                ? '<span style="color: #065f46; font-weight: 600;">VALID</span>'
                                : '<span style="color: #991b1b; font-weight: 600;">INVALID</span>'}
                        </td>
                    </tr>
                `;
            } else {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">Total Financial Cost</td>
                        <td style="padding: 0.875rem 1rem; text-align: center; color: #991b1b;">-</td>
                        <td style="padding: 0.875rem 1rem; text-align: center; color: #991b1b;">${totalFinancial.message || 'Not found'}</td>
                        <td style="padding: 0.875rem 1rem; text-align: center; color: #991b1b;">-</td>
                        <td style="padding: 0.875rem 1rem; text-align: center; color: #991b1b;">-</td>
                        <td style="padding: 0.875rem 1rem; text-align: center;">
                            <span style="color: #991b1b; font-weight: 600;">NOT FOUND</span>
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
     * Search by filename - filters file result groups based on filename
     */
    searchByFilename(searchTerm) {
        const resultsContainer = document.getElementById('results-v6');
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
     * Export to PDF (placeholder for now)
     */
    async exportToPDF() {
        alert('Export functionality will be implemented later.');
    }
}

// Initialize the processor
window.llbeanProcessor = new LLBEANProcessor();

// Initialize when V6 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    const v6Tab = document.querySelector('[data-tab="v6"]');
    if (v6Tab) {
        v6Tab.addEventListener('click', () => {
            window.llbeanProcessor.initialize();
        });
    }

    // If V6 tab is already active on load, initialize immediately
    const v6TabContent = document.getElementById('tab-v6');
    if (v6TabContent && v6TabContent.classList.contains('active')) {
        window.llbeanProcessor.initialize();
    }
});
