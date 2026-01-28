/**
 * Cotopaxi Cost Breakdown Processor
 * Validates BCBD files for Cotopaxi brand requirements
 *
 * Validation Logic:
 * 1. Find sheet named "Blank Cost Sheet"
 * 2. Check cell D7 for "VENDOR / COO" → Cell E7 should be "PT UWU JUMP INDONESIA" or "HEADS UP"
 * 3. Check cell D8 for "SUPPLIER CONTACT" → Cell E8 should be "Madison 88"
 */

class CotopaxiProcessor {
    constructor() {
        this.validationRules = [
            {
                name: 'Vendor / COO',
                row: 6, // Row 7 (0-indexed)
                markerColumn: 3, // Column D (0-indexed)
                marker: 'vendor / coo',
                checkColumn: 4, // Column E
                expected: ['PT UWU JUMP INDONESIA', 'HEADS UP'],
                type: 'multiple'
            },
            {
                name: 'Supplier Contact',
                row: 7, // Row 8 (0-indexed)
                markerColumn: 3, // Column D
                marker: 'supplier contact',
                checkColumn: 4, // Column E
                expected: 'Madison 88',
                type: 'exact'
            }
        ];
    }

    async initialize() {
        this.displayValidationRules();
    }

    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v24');
        if (!obDropZone) return;

        let html = `
            <div class="burton-cost-container">
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Cotopaxi Validation Rules:</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Sheet:</strong> Blank Cost Sheet</div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Rule 1 - Vendor / COO:</strong></div>
                        <div class="burton-item-line">Cell D7 = "VENDOR / COO" → Cell E7 = <strong>PT UWU JUMP INDONESIA</strong> or <strong>HEADS UP</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Rule 2 - Supplier Contact:</strong></div>
                        <div class="burton-item-line">Cell D8 = "SUPPLIER CONTACT" → Cell E8 = <strong>Madison 88</strong></div>
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

                    // Find the "Blank Cost Sheet" sheet
                    let targetSheetName = null;
                    for (const sheetName of workbook.SheetNames) {
                        if (sheetName.toLowerCase() === 'blank cost sheet') {
                            targetSheetName = sheetName;
                            break;
                        }
                    }

                    if (!targetSheetName) {
                        resolve({
                            fileName: file.name,
                            sheetName: 'Not Found',
                            checks: [{
                                name: 'Sheet Check',
                                expected: 'Blank Cost Sheet',
                                found: false,
                                rowNumber: -1,
                                actual: 'Sheet "Blank Cost Sheet" not found',
                                isValid: false,
                                markerColumn: '-',
                                checkColumn: '-'
                            }],
                            error: null
                        });
                        return;
                    }

                    const sheet = workbook.Sheets[targetSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                    const checks = this.validateSheet(jsonData);

                    resolve({
                        fileName: file.name,
                        sheetName: targetSheetName,
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

        // Validate each rule
        for (const rule of this.validationRules) {
            const result = {
                name: rule.name,
                expected: Array.isArray(rule.expected) ? rule.expected.join(' or ') : rule.expected,
                found: false,
                rowNumber: rule.row + 1,
                actual: '',
                isValid: false,
                markerColumn: this.getColumnLetter(rule.markerColumn),
                checkColumn: this.getColumnLetter(rule.checkColumn)
            };

            // Check if row exists
            if (jsonData.length > rule.row && jsonData[rule.row]) {
                const row = jsonData[rule.row];
                const markerCell = row[rule.markerColumn] ? String(row[rule.markerColumn]).trim() : '';
                const markerCellLower = markerCell.toLowerCase();

                // Check if marker matches
                if (markerCellLower === rule.marker.toLowerCase() || markerCellLower.includes(rule.marker.toLowerCase())) {
                    result.found = true;

                    const actualValue = row[rule.checkColumn] ? String(row[rule.checkColumn]).trim() : '';
                    result.actual = actualValue || 'Empty';

                    // Validate based on type
                    if (rule.type === 'exact') {
                        result.isValid = actualValue.toLowerCase() === rule.expected.toLowerCase();
                    } else if (rule.type === 'multiple') {
                        // Check if actual value matches any of the expected values
                        result.isValid = rule.expected.some(exp =>
                            actualValue.toLowerCase() === exp.toLowerCase()
                        );
                    }
                } else {
                    result.actual = `Marker "${rule.marker}" not found at ${result.markerColumn}${result.rowNumber}`;
                }
            } else {
                result.actual = `Row ${result.rowNumber} not found`;
            }

            results.push(result);
        }

        return results;
    }

    getColumnLetter(index) {
        const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
        if (index < 26) {
            return letters[index];
        }
        return letters[Math.floor(index / 26) - 1] + letters[index % 26];
    }

    generateResultsHTML(results) {
        let html = '';

        // Add export button at the top
        html += `
            <div style="margin-bottom: 15px; display: flex; justify-content: flex-end; align-items: center;">
                <button onclick="window.cotopaxiProcessor.exportToPDF()" class="export-btn">
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
                    } else if (check.isValid) {
                        valueHTML = `<span style="color: #065f46; font-weight: 600;">${check.actual}</span>`;
                        expectedHTML = `<span style="color: #065f46;">&#10003;</span>`;
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

        const config = window.pdfExporter.createCotopaxiConfig(this.fileResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }
}

// Initialize processor
window.cotopaxiProcessor = new CotopaxiProcessor();
