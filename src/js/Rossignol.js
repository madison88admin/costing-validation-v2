/**
 * Rossignol Cost Breakdown Processor
 * Validates BCBD files for Rossignol brand requirements
 *
 * Validation Logic:
 * 1. Check column D for "VENDOR NAME" → Column E should be "Madison 88"
 * 2. Check column D for "CURRENCY" → Column E should be "USD"
 * 3. Check cell H2 → Should be 5% to 10%
 * 4. Check column I for "FACTORY MARGIN" → Column J should be 0.40 to 0.70
 * 5. Check column A for "FABRIC" rows → Column B has sequence, Column L should be 5%
 */

class RossignolProcessor {
    constructor() {
        this.validationRules = [
            {
                name: 'Vendor Name',
                markerColumn: 3, // Column D (0-indexed)
                marker: 'vendor name',
                checkColumn: 4, // Column E
                expected: 'Madison 88',
                type: 'exact'
            },
            {
                name: 'Currency',
                markerColumn: 3, // Column D
                marker: 'currency',
                checkColumn: 4, // Column E
                expected: 'USD',
                type: 'exact'
            }
        ];
    }

    async initialize() {
        this.displayValidationRules();
    }

    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v23');
        if (!obDropZone) return;

        let html = `
            <div class="burton-cost-container">
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Rossignol Validation Rules:</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Rule 1 - Vendor Name:</strong></div>
                        <div class="burton-item-line">Find "VENDOR NAME" in Column D → Column E = <strong>Madison 88</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Rule 2 - Currency:</strong></div>
                        <div class="burton-item-line">Find "CURRENCY" in Column D → Column E = <strong>USD</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Rule 3 - H2 Check:</strong></div>
                        <div class="burton-item-line">Cell H2 should be <strong>5% - 10%</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Rule 4 - Factory Margin:</strong></div>
                        <div class="burton-item-line">Find "FACTORY MARGIN" in Column I → Column J = <strong>0.40 - 0.70</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Rule 5 - Material Category Rows:</strong></div>
                        <div class="burton-item-line">FABRIC → Column L = <strong>5%</strong></div>
                        <div class="burton-item-line">TRIM, Accessories, GRAPHIC, Labelling → Column L = <strong>3%</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Rule 6 - PACKAGING Rows:</strong></div>
                        <div class="burton-item-line">If Column D = "Generic Packaging" → G = <strong>m88</strong>, J = <strong>pc</strong>, L = <strong>1</strong></div>
                        <div class="burton-item-line">Otherwise → Column L = <strong>3%</strong></div>
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

                    // Process first sheet
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

        // Validate standard rules (Column D → Column E)
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

                if (markerCellLower === rule.marker.toLowerCase()) {
                    result.found = true;
                    result.rowNumber = i + 1;

                    const actualValue = row[rule.checkColumn] ? String(row[rule.checkColumn]).trim() : '';
                    result.actual = actualValue || 'Empty';

                    result.isValid = actualValue.toLowerCase() === rule.expected.toLowerCase();
                    break;
                }
            }

            results.push(result);
        }

        // Validate H2 cell (5% - 10%)
        const h2Result = this.validateH2Cell(jsonData);
        results.push(h2Result);

        // Validate Factory Margin (Column I → Column J)
        const factoryMarginResult = this.validateFactoryMargin(jsonData);
        results.push(factoryMarginResult);

        // Validate FABRIC rows (Column A → Column L should be 5%)
        const fabricResults = this.validateCategoryRows(jsonData, 'FABRIC', 0.05);
        for (const fabricResult of fabricResults) {
            results.push(fabricResult);
        }

        // Validate other category rows (Column A → Column L should be 3%)
        const categories3Percent = ['TRIM', 'ACCESSORIES', 'GRAPHIC', 'LABELLING'];
        for (const category of categories3Percent) {
            const categoryResults = this.validateCategoryRows(jsonData, category, 0.03);
            for (const categoryResult of categoryResults) {
                results.push(categoryResult);
            }
        }

        // Validate PACKAGING rows (special validation with Generic Packaging check)
        const packagingResults = this.validatePackagingRows(jsonData);
        for (const packagingResult of packagingResults) {
            results.push(packagingResult);
        }

        return results;
    }

    validateCategoryRows(jsonData, categoryName, expectedDecimal) {
        const entries = [];
        let allValid = true;
        const expectedPercent = (expectedDecimal * 100).toFixed(0) + '%';

        // Scan column A for the category
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const cellA = row[0] ? String(row[0]).trim().toUpperCase() : '';

            if (cellA === categoryName.toUpperCase()) {
                // Get column B (sequence number)
                const cellB = row[1] !== undefined && row[1] !== null ? String(row[1]).trim() : '';

                // Get column L value (index 11)
                const cellL = row[11];
                let actualValue = '';
                let isValid = false;

                if (cellL !== undefined && cellL !== null) {
                    let numValue = parseFloat(cellL);

                    if (!isNaN(numValue)) {
                        // If value is <= 1, it's a decimal (e.g., 0.05 for 5%)
                        if (numValue <= 1) {
                            actualValue = (numValue * 100).toFixed(0) + '%';
                            isValid = Math.abs(numValue - expectedDecimal) < 0.0001;
                        } else {
                            // Already a percentage (e.g., 5 for 5%)
                            actualValue = numValue.toFixed(0) + '%';
                            isValid = Math.abs(numValue - (expectedDecimal * 100)) < 0.01;
                        }
                    } else {
                        actualValue = String(cellL).trim() || 'Empty';
                    }
                } else {
                    actualValue = 'Empty';
                }

                if (!isValid) allValid = false;

                entries.push({
                    sequence: cellB,
                    value: actualValue,
                    isValid: isValid,
                    rowNumber: i + 1
                });
            }
        }

        // If no rows found for this category
        if (entries.length === 0) {
            return [{
                name: `${categoryName} Rows`,
                expected: expectedPercent,
                found: false,
                rowNumber: -1,
                actual: `No ${categoryName} rows found`,
                isValid: false,
                markerColumn: 'A',
                checkColumn: 'L'
            }];
        }

        // Return single combined result
        return [{
            name: `${categoryName} Rows`,
            expected: expectedPercent,
            found: true,
            rowNumber: -1,
            actual: '',
            isValid: allValid,
            markerColumn: 'A',
            checkColumn: 'L',
            isFabricCombined: true,
            categoryName: categoryName,
            fabricEntries: entries
        }];
    }

    validatePackagingRows(jsonData) {
        const entries = [];
        let allValid = true;

        // Scan column A for PACKAGING
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const cellA = row[0] ? String(row[0]).trim().toUpperCase() : '';

            if (cellA === 'PACKAGING') {
                // Get column B (sequence number)
                const cellB = row[1] !== undefined && row[1] !== null ? String(row[1]).trim() : '';

                // Get column D value (index 3)
                const cellD = row[3] ? String(row[3]).trim() : '';
                const hasGenericPackaging = cellD.toLowerCase().includes('generic packaging');

                if (hasGenericPackaging) {
                    // Validate Column G = "m88", Column J = "pc", Column L = 1
                    const cellG = row[6] ? String(row[6]).trim() : '';
                    const cellJ = row[9] ? String(row[9]).trim() : '';
                    const cellL = row[11];

                    // Check each value
                    const isGValid = cellG.toLowerCase() === 'm88';
                    const isJValid = cellJ.toLowerCase() === 'pc';

                    // Check Column L = 1
                    let isLValid = false;
                    let cellLValue = '';
                    if (cellL !== undefined && cellL !== null) {
                        let numValue = parseFloat(cellL);
                        if (!isNaN(numValue)) {
                            cellLValue = numValue.toString();
                            isLValid = numValue === 1;
                        } else {
                            cellLValue = String(cellL).trim() || 'Empty';
                        }
                    } else {
                        cellLValue = 'Empty';
                    }

                    const isRowValid = isGValid && isJValid && isLValid;
                    if (!isRowValid) allValid = false;

                    entries.push({
                        sequence: cellB,
                        hasGenericPackaging: true,
                        cellValidations: [
                            { column: 'D', value: cellD, isValid: true, expected: 'Generic Packaging' },
                            { column: 'G', value: cellG || 'Empty', isValid: isGValid, expected: 'm88' },
                            { column: 'J', value: cellJ || 'Empty', isValid: isJValid, expected: 'pc' },
                            { column: 'L', value: cellLValue, isValid: isLValid, expected: '1' }
                        ],
                        isValid: isRowValid,
                        rowNumber: i + 1
                    });
                } else {
                    // No Generic Packaging - validate Column L = 3%
                    const cellL = row[11];
                    let cellLValue = '';
                    let isLValid = false;

                    if (cellL !== undefined && cellL !== null) {
                        let numValue = parseFloat(cellL);
                        if (!isNaN(numValue)) {
                            // If value is <= 1, it's a decimal (e.g., 0.03 for 3%)
                            if (numValue <= 1) {
                                cellLValue = (numValue * 100).toFixed(0) + '%';
                                isLValid = Math.abs(numValue - 0.03) < 0.0001;
                            } else {
                                // Already a percentage (e.g., 3 for 3%)
                                cellLValue = numValue.toFixed(0) + '%';
                                isLValid = Math.abs(numValue - 3) < 0.01;
                            }
                        } else {
                            cellLValue = String(cellL).trim() || 'Empty';
                        }
                    } else {
                        cellLValue = 'Empty';
                    }

                    if (!isLValid) allValid = false;

                    entries.push({
                        sequence: cellB,
                        hasGenericPackaging: false,
                        value: cellLValue,
                        isValid: isLValid,
                        rowNumber: i + 1
                    });
                }
            }
        }

        // If no PACKAGING rows found
        if (entries.length === 0) {
            return [{
                name: 'PACKAGING Rows',
                expected: 'Generic Packaging: G=m88, J=pc, L=1 | Other: 3%',
                found: false,
                rowNumber: -1,
                actual: 'No PACKAGING rows found',
                isValid: false,
                markerColumn: 'A',
                checkColumn: 'D,G,J,L'
            }];
        }

        // Return single combined result
        return [{
            name: 'PACKAGING Rows',
            expected: 'Generic Packaging: G=m88, J=pc, L=1 | Other: 3%',
            found: true,
            rowNumber: -1,
            actual: '',
            isValid: allValid,
            markerColumn: 'A',
            checkColumn: 'D,G,J,L',
            isPackagingCombined: true,
            packagingEntries: entries
        }];
    }

    validateH2Cell(jsonData) {
        const result = {
            name: 'H2 Value',
            expected: '5% - 10%',
            found: false,
            rowNumber: 2,
            actual: '',
            isValid: false,
            markerColumn: 'H',
            checkColumn: 'H'
        };

        // H2 is row index 1 (0-indexed), column index 7 (H = 8th column, 0-indexed = 7)
        if (jsonData.length >= 2 && jsonData[1]) {
            const cellValue = jsonData[1][7];
            if (cellValue !== undefined && cellValue !== null) {
                result.found = true;
                let numValue = parseFloat(cellValue);

                if (!isNaN(numValue)) {
                    // If value is <= 1, it's a decimal (e.g., 0.05 for 5%)
                    if (numValue <= 1) {
                        result.actual = (numValue * 100).toFixed(2) + '%';
                        result.isValid = numValue >= 0.05 && numValue <= 0.10;
                    } else {
                        // Already a percentage (e.g., 5 for 5%)
                        result.actual = numValue.toFixed(2) + '%';
                        result.isValid = numValue >= 5 && numValue <= 10;
                    }
                } else {
                    result.actual = String(cellValue).trim() || 'Empty';
                    result.isValid = false;
                }
            } else {
                result.actual = 'Empty';
            }
        } else {
            result.actual = 'Row not found';
        }

        return result;
    }

    validateFactoryMargin(jsonData) {
        const result = {
            name: 'Factory Margin',
            expected: '0.40 - 0.70',
            found: false,
            rowNumber: -1,
            actual: '',
            isValid: false,
            markerColumn: 'I',
            checkColumn: 'J'
        };

        // Scan column I (index 8) for "FACTORY MARGIN"
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const cellI = row[8] ? String(row[8]).trim().toLowerCase() : '';

            if (cellI === 'factory margin') {
                result.found = true;
                result.rowNumber = i + 1;

                // Get column J value (index 9)
                const cellJ = row[9];
                if (cellJ !== undefined && cellJ !== null) {
                    let numValue = parseFloat(cellJ);

                    if (!isNaN(numValue)) {
                        result.actual = numValue.toFixed(2);
                        result.isValid = numValue >= 0.40 && numValue <= 0.70;
                    } else {
                        result.actual = String(cellJ).trim() || 'Empty';
                        result.isValid = false;
                    }
                } else {
                    result.actual = 'Empty';
                }

                break;
            }
        }

        if (!result.found) {
            result.actual = 'FACTORY MARGIN not found';
        }

        return result;
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
                <button onclick="window.rossignolProcessor.exportToPDF()" class="export-btn">
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

                    if (check.isPackagingCombined && check.packagingEntries) {
                        // Combined PACKAGING rows display with special validation
                        const lines = check.packagingEntries.map(entry => {
                            if (entry.hasGenericPackaging) {
                                // Show all validated cells with individual coloring
                                const cellValues = entry.cellValidations.map(cv => {
                                    const color = cv.isValid ? '#065f46' : '#991b1b';
                                    return `<span style="color: ${color}; font-weight: 600;">${cv.value}</span>`;
                                }).join(' | ');
                                return `PACKAGING ${entry.sequence}: ${cellValues}`;
                            } else {
                                // No Generic Packaging - show Column L value with 3% validation
                                const color = entry.isValid ? '#065f46' : '#991b1b';
                                return `<span style="color: ${color}; font-weight: 600;">PACKAGING ${entry.sequence} ${entry.value}</span>`;
                            }
                        });
                        valueHTML = lines.join('<br>');
                        // Show expected value in Expected column
                        expectedHTML = `<span style="color: #849bba;">${check.expected}</span>`;
                    } else if (check.isFabricCombined && check.fabricEntries) {
                        // Combined category rows display
                        const catName = check.categoryName || 'FABRIC';
                        const lines = check.fabricEntries.map(entry => {
                            if (entry.isValid) {
                                return `<span style="color: #065f46; font-weight: 600;">${catName} ${entry.sequence} ${entry.value}</span>`;
                            } else {
                                return `<span style="color: #991b1b; font-weight: 600;">${catName} ${entry.sequence} ${entry.value}</span>`;
                            }
                        });
                        valueHTML = lines.join('<br>');
                        // Show expected value in Expected column
                        expectedHTML = `<span style="color: #849bba;">${check.expected}</span>`;
                    } else if (!check.found) {
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

        const config = window.pdfExporter.createRossignolConfig(this.fileResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }
}

// Initialize processor
window.rossignolProcessor = new RossignolProcessor();
