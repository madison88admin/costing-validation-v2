/**
 * Outdoor Research Processing Logic (V8)
 * Validates Buyer CBD files against Outdoor Research criteria
 */

class OutdoorResearchProcessor {
    constructor() {
        this.bcbdFiles = [];
        this.bcbdResults = [];
        this.validationRules = {
            // General Packaging check
            // Find "General Packaging" in Column D, then verify same row values
            generalPackaging: {
                searchCol: 3,           // Column D (index 3) - search for "General Packaging"
                searchValue: 'GENERAL PACKAGING',
                checks: [
                    { colIndex: 2, colName: 'C', expectedValue: 'PACKING', label: 'Packing' },
                    { colIndex: 4, colName: 'E', expectedValue: 'FACTORY SUPPLIED', label: 'Factory Supplied' },
                    { colIndex: 6, colName: 'G', expectedValue: 1, label: 'Quantity', isNumeric: true },
                    { colIndex: 7, colName: 'H', expectedValue: 'PC', label: 'Unit' }
                ]
            },
            // Other Charges check
            // Find "Other Charges" in Column D, then verify same row values
            otherCharges: {
                searchCol: 3,           // Column D (index 3) - search for "Other Charges"
                searchValue: 'OTHER CHARGES',
                checks: [
                    { colIndex: 5, colName: 'F', expectedValue: 'OVERHEAD/PROFIT', label: 'Overhead/Profit' },
                    { colIndex: 6, colName: 'G', expectedValue: 1, label: 'Quantity', isNumeric: true },
                    { colIndex: 8, colName: 'I', expectedValue: 0.5, label: 'Value', isNumeric: true }
                ]
            }
        };
    }

    /**
     * Initialize V8 - Display validation rules in the OB drop zone
     */
    initialize() {
        this.displayValidationRules();
        console.log('Outdoor Research Processor initialized');
    }

    /**
     * Display validation rules in the OB drop zone (Burton-style)
     */
    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v8');
        if (!obDropZone) return;

        const gp = this.validationRules.generalPackaging;
        const oc = this.validationRules.otherCharges;

        let html = `
            <div class="burton-cost-container">
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>General Packaging Check</strong></div>
                        <div class="burton-item-line"><strong>Search:</strong> Column D = "${gp.searchValue}"</div>
                        ${gp.checks.map(check =>
                            `<div class="burton-item-line"><strong>Column ${check.colName}:</strong> ${check.expectedValue}</div>`
                        ).join('')}
                    </div>
                    <div class="burton-cost-item" style="margin-top: 1rem;">
                        <div class="burton-item-line"><strong>Other Charges Check</strong></div>
                        <div class="burton-item-line"><strong>Search:</strong> Column D = "${oc.searchValue}"</div>
                        ${oc.checks.map(check =>
                            `<div class="burton-item-line"><strong>Column ${check.colName}:</strong> ${check.expectedValue}</div>`
                        ).join('')}
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

                    // Run all validations
                    const generalPackagingCheck = this.checkGeneralPackaging(jsonData);
                    const otherChargesCheck = this.checkOtherCharges(jsonData);

                    resolve({
                        generalPackagingCheck: generalPackagingCheck,
                        otherChargesCheck: otherChargesCheck
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
     * Generic row validation check
     * Finds a search value in a column, then verifies values in other columns on the same row
     */
    checkRowValidation(jsonData, rule, ruleName) {
        const searchCol = rule.searchCol;
        const searchValue = rule.searchValue;

        let found = false;
        let rowNumber = -1;
        let checkResults = [];

        // Search for the value in the specified column
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const cellValue = row[searchCol] ? String(row[searchCol]).trim().toUpperCase() : '';

            if (cellValue.includes(searchValue)) {
                found = true;
                rowNumber = i + 1;
                console.log(`Found "${ruleName}" at row ${rowNumber}`);

                // Now check all the required columns on this row
                for (const check of rule.checks) {
                    const actualValue = row[check.colIndex];
                    let isValid = false;
                    let displayValue = actualValue;

                    if (check.isNumeric) {
                        // Numeric comparison
                        const numericActual = parseFloat(actualValue);
                        isValid = !isNaN(numericActual) && numericActual === check.expectedValue;
                        displayValue = isNaN(numericActual) ? actualValue : numericActual;
                    } else {
                        // String comparison (case-insensitive)
                        const actualStr = actualValue ? String(actualValue).trim().toUpperCase() : '';
                        isValid = actualStr === check.expectedValue.toUpperCase();
                        displayValue = actualValue ? String(actualValue).trim() : '';
                    }

                    checkResults.push({
                        label: check.label,
                        column: check.colName,
                        expectedValue: check.expectedValue,
                        actualValue: displayValue,
                        isValid: isValid
                    });
                }

                break;
            }
        }

        if (!found) {
            return {
                found: false,
                message: `${ruleName} not found in Column D`
            };
        }

        const allValid = checkResults.every(r => r.isValid);

        return {
            found: true,
            rowNumber: rowNumber,
            checks: checkResults,
            isValid: allValid
        };
    }

    /**
     * Check General Packaging row
     */
    checkGeneralPackaging(jsonData) {
        return this.checkRowValidation(jsonData, this.validationRules.generalPackaging, 'General Packaging');
    }

    /**
     * Check Other Charges row
     */
    checkOtherCharges(jsonData) {
        return this.checkRowValidation(jsonData, this.validationRules.otherCharges, 'Other Charges');
    }

    /**
     * Generate HTML for results display
     */
    generateResultsHTML(results) {
        if (!results || results.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">Outdoor Research Validation Ready</p>
                    <p>Upload Buyer CBD files to validate.</p>
                </div>
            `;
        }

        let html = '';

        // Add export button at the top
        html += `
            <div style="margin-bottom: 15px; display: flex; justify-content: flex-end; align-items: center;">
                <button onclick="window.outdoorResearchProcessor.exportToPDF()" class="export-btn">
                    Export
                </button>
            </div>
        `;

        for (const fileResult of results) {
            const gpCheck = fileResult.results.generalPackagingCheck;
            const ocCheck = fileResult.results.otherChargesCheck;

            // Count valid checks
            let validCount = 0;
            let totalChecks = 0;

            if (gpCheck && gpCheck.found && gpCheck.checks) {
                totalChecks++;
                if (gpCheck.isValid) validCount++;
            }
            if (ocCheck && ocCheck.found && ocCheck.checks) {
                totalChecks++;
                if (ocCheck.isValid) validCount++;
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
                            <th style="width: 30%;">Validation Check</th>
                            <th style="width: 70%;">Value</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            // General Packaging checks - all on one row
            if (gpCheck && gpCheck.found && gpCheck.checks) {
                const checkDetails = gpCheck.checks.map(check => {
                    if (check.isValid) {
                        return `<span style="color: #065f46; font-weight: 600;">${check.actualValue || 'Empty'}</span>`;
                    } else {
                        return `<span style="color: #991b1b; font-weight: 600;">${check.actualValue || 'Empty'}</span> <span style="font-size: 0.85em; color: #849bba;">(Expected: ${check.expectedValue})</span>`;
                    }
                }).join(' | ');

                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">General Packaging (Row ${gpCheck.rowNumber})</td>
                        <td style="padding: 0.875rem 1rem; text-align: left;">
                            ${checkDetails}
                        </td>
                    </tr>
                `;
            } else if (gpCheck && !gpCheck.found) {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">General Packaging</td>
                        <td style="padding: 0.875rem 1rem; text-align: left; color: #991b1b; font-weight: 600;">${gpCheck.message || 'Not found'}</td>
                    </tr>
                `;
            }

            // Other Charges checks - all on one row
            if (ocCheck && ocCheck.found && ocCheck.checks) {
                const checkDetails = ocCheck.checks.map(check => {
                    if (check.isValid) {
                        return `<span style="color: #065f46; font-weight: 600;">${check.actualValue || 'Empty'}</span>`;
                    } else {
                        return `<span style="color: #991b1b; font-weight: 600;">${check.actualValue || 'Empty'}</span> <span style="font-size: 0.85em; color: #849bba;">(Expected: ${check.expectedValue})</span>`;
                    }
                }).join(' | ');

                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">Other Charges (Row ${ocCheck.rowNumber})</td>
                        <td style="padding: 0.875rem 1rem; text-align: left;">
                            ${checkDetails}
                        </td>
                    </tr>
                `;
            } else if (ocCheck && !ocCheck.found) {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">Other Charges</td>
                        <td style="padding: 0.875rem 1rem; text-align: left; color: #991b1b; font-weight: 600;">${ocCheck.message || 'Not found'}</td>
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
        const resultsContainer = document.getElementById('results-v8');
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

        const config = window.pdfExporter.createOutdoorResearchConfig(this.bcbdResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }
}

// Initialize the processor
window.outdoorResearchProcessor = new OutdoorResearchProcessor();

// Initialize when V8 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    const v8Tab = document.querySelector('[data-tab="v8"]');
    if (v8Tab) {
        v8Tab.addEventListener('click', () => {
            window.outdoorResearchProcessor.initialize();
        });
    }

    // If V8 tab is already active on load, initialize immediately
    const v8TabContent = document.getElementById('tab-v8');
    if (v8TabContent && v8TabContent.classList.contains('active')) {
        window.outdoorResearchProcessor.initialize();
    }
});
