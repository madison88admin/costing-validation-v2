/**
 * ODLO Cost Breakdown Processor
 * Validates BCBD files for ODLO brand requirements
 *
 * Validation Logic:
 * 1. Check column A for the word "Category"
 * 2. On the same row, check if column B says "Accessories"
 * 3. Check column A for the word "Garment Maker"
 * 4. On the same row, check if column B says "Madison 88 (USD)"
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
                        <div class="burton-item-line">• Find "Category" in Column A</div>
                        <div class="burton-item-line">• Column B on same row should be: <strong>Accessories</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Rule 2 - Garment Maker Check:</strong></div>
                        <div class="burton-item-line">• Find "Garment Maker" in Column A</div>
                        <div class="burton-item-line">• Column B on same row should be: <strong>Madison 88 (USD)</strong></div>
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

                // Create results table
                html += `
                    <table class="results-table" style="table-layout: fixed; width: 100%;">
                        <thead>
                            <tr class="header-labels-row">
                                <th style="width: 300px;">Check Name</th>
                                <th>Value</th>
                            </tr>
                        </thead>
                        <tbody>
                `;

                for (const check of fileResult.checks) {
                    let valueHTML = '';

                    if (!check.found) {
                        valueHTML = `<span style="color: #b45309; font-weight: 600;">Not Found</span>`;
                    } else if (check.isValid) {
                        valueHTML = `<span style="color: #065f46; font-weight: 600;">${check.actual}</span>`;
                    } else {
                        valueHTML = `<span style="color: #991b1b; font-weight: 600;">${check.actual}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${check.expected}</span>`;
                    }

                    html += `
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">${check.name}</td>
                            <td style="padding: 0.875rem 1rem;">${valueHTML}</td>
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
