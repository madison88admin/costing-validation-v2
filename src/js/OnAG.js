/**
 * On AG Processing Logic (V9)
 * Validates Buyer CBD files against On AG criteria
 */

class OnAGProcessor {
    constructor() {
        this.bcbdFiles = [];
        this.bcbdResults = [];
        this.validationRules = {
            // Wastage sections - similar to Mammut pattern
            // Column A = section keyword, Column J = wastage %, Column I = "Total" marks end
            wastageSections: [
                { keyword: 'MATERIAL', wastageCol: 9, wastageExpected: 5.00, label: 'Material', stopCol: 8 },   // Col J (index 9), Col I (index 8)
                { keyword: 'TRIMS', wastageCol: 9, wastageExpected: 2.00, label: 'Trims', stopCol: 8 },
                { keyword: 'PACKAGING', wastageCol: 9, wastageExpected: 2.00, label: 'Packaging', stopCol: 8 }
            ],
            // Special rule: Coats Thread within Material section has different values
            // Column B = "Coats Thread", Column I = PX/Unit (0.001), Column J = Wastage (2.00), Column L = Freight (0.0002)
            coatsThread: {
                keyword: 'COATS THREAD',
                searchCol: 1,           // Column B (index 1)
                checks: [
                    { colIndex: 8, colName: 'I', expectedValue: 0.001, label: 'PX/Unit', isNumeric: true },
                    { colIndex: 9, colName: 'J', expectedValue: 2.00, label: 'Wastage', isNumeric: true },
                    { colIndex: 11, colName: 'L', expectedValue: 0.0002, label: 'Freight', isNumeric: true }
                ]
            },
            // Process and Cost checks - Column A keyword, Column I value
            processCosts: [
                { keyword: 'KNITTING', expectedValue: 0.06, label: 'Knitting' },
                { keyword: 'SEWING', expectedValue: 0.08, label: 'Sewing' },
                { keyword: 'LABELING', expectedValue: 0.10, label: 'Labeling' },
                { keyword: 'FINISHING/STEAM/PACK', expectedValue: 0.40, label: 'Finishing/Steam/Pack' },
                { keyword: 'OVERHEAD IN %', expectedValue: 6.00, label: 'Overhead in %' },
                { keyword: 'PROFIT IN %', expectedValue: 6.00, label: 'Profit in %' },
                { keyword: 'FINANCE COST IN %', expectedValue: 5.00, label: 'Finance cost in %' },
                { keyword: 'ADDITIONAL FREIGHT IN %', expectedValue: 1.00, label: 'Additional Freight in %' },
                { keyword: 'LOGISTICS (TRANS & DOCS) IN %', expectedValue: 1.00, label: 'Logistics (Trans & Docs) in %' }
            ]
        };
    }

    /**
     * Initialize V9 - Display validation rules in the OB drop zone
     */
    initialize() {
        this.displayValidationRules();
        console.log('On AG Processor initialized');
    }

    /**
     * Display validation rules in the OB drop zone (Burton-style)
     */
    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v9');
        if (!obDropZone) return;

        let html = `
            <div class="burton-cost-container">
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>On AG Wastage Validation</strong></div>
                        <div class="burton-item-line"><strong>Search Column:</strong> A (Section Keywords)</div>
                        <div class="burton-item-line"><strong>Wastage Column:</strong> J</div>
                        <div class="burton-item-line"><strong>Stop at:</strong> "Total" in Column I</div>
                    </div>
                    ${this.validationRules.wastageSections.map(section => `
                        <div class="burton-cost-item" style="margin-top: 0.5rem;">
                            <div class="burton-item-line"><strong>${section.label} Wastage:</strong> ${section.wastageExpected}%</div>
                        </div>
                    `).join('')}
                    <div class="burton-cost-item" style="margin-top: 0.5rem;">
                        <div class="burton-item-line"><strong>Coats Thread (in Material):</strong></div>
                        <div class="burton-item-line">PX/Unit (Col I): 0.001</div>
                        <div class="burton-item-line">Wastage (Col J): 2.00%</div>
                        <div class="burton-item-line">Freight (Col L): 0.0002</div>
                    </div>
                    <div class="burton-cost-item" style="margin-top: 0.5rem;">
                        <div class="burton-item-line"><strong>Process & Cost Checks (Col A -> Col I):</strong></div>
                        ${this.validationRules.processCosts.map(item =>
                            `<div class="burton-item-line">${item.label}: ${item.expectedValue}</div>`
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

                    // Run wastage validations for each section
                    const wastageResults = this.checkAllWastageSections(jsonData);

                    // Check for Coats Thread special rule within Material section
                    const coatsThreadCheck = this.checkCoatsThread(jsonData);

                    // Check process and cost items (Knitting, Sewing, etc.)
                    const processCostsCheck = this.checkProcessCosts(jsonData);

                    resolve({
                        wastageResults: wastageResults,
                        coatsThreadCheck: coatsThreadCheck,
                        processCostsCheck: processCostsCheck
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
     * Check all wastage sections (Material, Trims, Packaging)
     */
    checkAllWastageSections(jsonData) {
        const results = [];

        for (const section of this.validationRules.wastageSections) {
            const sectionResult = this.checkWastageSection(jsonData, section);
            results.push(sectionResult);
        }

        return results;
    }

    /**
     * Check wastage for a specific section
     * Finds keyword in Column A, checks Column J for wastage %, stops at "Total" in Column I
     */
    checkWastageSection(jsonData, section) {
        const searchCol = 0;  // Column A (index 0)
        const wastageCol = section.wastageCol;  // Column J (index 9)
        const stopCol = section.stopCol;  // Column I (index 8)
        const expectedWastage = section.wastageExpected;

        let sectionFound = false;
        let startRow = -1;
        let endRow = -1;
        let validCells = [];
        let invalidCells = [];

        // First, find the section keyword in Column A
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const colA = row[searchCol] ? String(row[searchCol]).trim().toUpperCase() : '';

            if (colA.includes(section.keyword)) {
                sectionFound = true;
                startRow = i;
                console.log(`Found "${section.label}" at row ${i + 1}`);

                // Now scan from this row until we find "Total" in Column I
                for (let j = i + 1; j < jsonData.length; j++) {
                    const checkRow = jsonData[j];
                    const colI = checkRow[stopCol] ? String(checkRow[stopCol]).trim().toUpperCase() : '';

                    // Check if we've hit the "Total" marker
                    if (colI.includes('TOTAL')) {
                        endRow = j;
                        console.log(`Found "Total" at row ${j + 1} for ${section.label}`);
                        break;
                    }

                    // Skip Coats Thread row in Material section - it has its own special rule
                    if (section.keyword === 'MATERIAL') {
                        const colB = checkRow[1] ? String(checkRow[1]).trim().toUpperCase() : '';
                        if (colB.includes('COATS THREAD')) {
                            console.log(`Skipping Coats Thread row ${j + 1} in Material wastage check`);
                            continue;
                        }
                    }

                    // Check the wastage value in Column J
                    const wastageValue = checkRow[wastageCol];

                    // Skip empty cells
                    if (wastageValue === '' || wastageValue === null || wastageValue === undefined) {
                        continue;
                    }

                    // Parse the wastage value
                    let numericWastage = parseFloat(wastageValue);

                    // Handle percentage format (e.g., "5%" or "5.00%")
                    if (typeof wastageValue === 'string' && wastageValue.includes('%')) {
                        numericWastage = parseFloat(wastageValue.replace('%', ''));
                    }

                    if (!isNaN(numericWastage)) {
                        const cellAddress = `J${j + 1}`;
                        const isValid = numericWastage === expectedWastage;

                        if (isValid) {
                            validCells.push({
                                rowNumber: j + 1,
                                cellAddress: cellAddress,
                                value: wastageValue,
                                numericValue: numericWastage
                            });
                        } else {
                            invalidCells.push({
                                rowNumber: j + 1,
                                cellAddress: cellAddress,
                                value: wastageValue,
                                numericValue: numericWastage,
                                expectedValue: expectedWastage
                            });
                        }
                    }
                }

                break;
            }
        }

        if (!sectionFound) {
            return {
                section: section.label,
                found: false,
                message: `"${section.label}" not found in Column A`
            };
        }

        return {
            section: section.label,
            found: true,
            startRow: startRow + 1,
            endRow: endRow + 1,
            expectedWastage: expectedWastage,
            validCells: validCells,
            invalidCells: invalidCells,
            isValid: invalidCells.length === 0
        };
    }

    /**
     * Check Coats Thread special rule within Material section
     * Finds "Coats Thread" in Column B within Material section
     * Checks: Column I = 0.001, Column J = 2.00, Column L = 0.0002
     */
    checkCoatsThread(jsonData) {
        const ct = this.validationRules.coatsThread;
        const searchCol = ct.searchCol;  // Column B (index 1)
        const keyword = ct.keyword;

        // First find Material section boundaries
        let materialStartRow = -1;
        let materialEndRow = -1;

        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const colA = row[0] ? String(row[0]).trim().toUpperCase() : '';

            if (colA.includes('MATERIAL')) {
                materialStartRow = i;

                // Find the Total marker in Column I
                for (let j = i + 1; j < jsonData.length; j++) {
                    const checkRow = jsonData[j];
                    const colI = checkRow[8] ? String(checkRow[8]).trim().toUpperCase() : '';

                    if (colI.includes('TOTAL')) {
                        materialEndRow = j;
                        break;
                    }
                }
                break;
            }
        }

        if (materialStartRow === -1) {
            return {
                found: false,
                message: 'Material section not found'
            };
        }

        // Now search for Coats Thread within Material section
        let coatsThreadFound = false;
        let rowNumber = -1;
        let checkResults = [];

        for (let i = materialStartRow; i < (materialEndRow !== -1 ? materialEndRow : jsonData.length); i++) {
            const row = jsonData[i];
            const colB = row[searchCol] ? String(row[searchCol]).trim().toUpperCase() : '';

            if (colB.includes(keyword)) {
                coatsThreadFound = true;
                rowNumber = i + 1;
                console.log(`Found "Coats Thread" at row ${rowNumber}`);

                // Check all required columns on this row
                for (const check of ct.checks) {
                    const actualValue = row[check.colIndex];
                    let numericActual = parseFloat(actualValue);
                    let isValid = false;

                    if (!isNaN(numericActual)) {
                        // Use tolerance for floating point comparison
                        isValid = Math.abs(numericActual - check.expectedValue) < 0.00001;
                    }

                    checkResults.push({
                        label: check.label,
                        column: check.colName,
                        expectedValue: check.expectedValue,
                        actualValue: actualValue,
                        numericValue: numericActual,
                        isValid: isValid,
                        cellAddress: `${check.colName}${rowNumber}`
                    });
                }

                break;
            }
        }

        if (!coatsThreadFound) {
            return {
                found: false,
                message: 'Coats Thread not found in Material section'
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
     * Check Process and Cost items
     * Finds keywords in Column A, checks Column I for expected value
     * Uses strict matching - cell must equal the keyword exactly (not contain it)
     */
    checkProcessCosts(jsonData) {
        const processCosts = this.validationRules.processCosts;
        const results = [];

        for (const item of processCosts) {
            let found = false;
            let rowNumber = -1;
            let actualValue = null;
            let numericValue = null;
            let isValid = false;

            // Search for the keyword in Column A - STRICT match (equals, not includes)
            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                const colA = row[0] ? String(row[0]).trim().toUpperCase() : '';

                // Strict match: cell must equal the keyword exactly
                if (colA === item.keyword) {
                    found = true;
                    rowNumber = i + 1;
                    actualValue = row[8];  // Column I (index 8)

                    // Parse the value
                    numericValue = parseFloat(actualValue);

                    // Handle percentage format - if string contains %, parse the number directly
                    if (typeof actualValue === 'string' && actualValue.includes('%')) {
                        numericValue = parseFloat(actualValue.replace('%', ''));
                    } else if (!isNaN(numericValue) && numericValue < 1 && item.expectedValue >= 1) {
                        // If the value is a decimal (like 0.06) but expected is percentage (like 6),
                        // multiply by 100 to convert to percentage
                        numericValue = numericValue * 100;
                    }

                    if (!isNaN(numericValue)) {
                        // Use tolerance for floating point comparison
                        isValid = Math.abs(numericValue - item.expectedValue) < 0.01;
                    }

                    console.log(`Found "${item.label}" at row ${rowNumber}, value: ${actualValue}, parsed: ${numericValue}`);
                    break;
                }
            }

            results.push({
                label: item.label,
                keyword: item.keyword,
                found: found,
                rowNumber: rowNumber,
                expectedValue: item.expectedValue,
                actualValue: actualValue,
                numericValue: numericValue,
                isValid: isValid,
                cellAddress: found ? `I${rowNumber}` : null
            });
        }

        const allFound = results.every(r => r.found);
        const allValid = results.every(r => r.isValid);

        return {
            found: allFound,
            items: results,
            isValid: allValid
        };
    }

    /**
     * Generate HTML for results display
     */
    generateResultsHTML(results) {
        if (!results || results.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">On AG Validation Ready</p>
                    <p>Upload Buyer CBD files to validate.</p>
                </div>
            `;
        }

        let html = '';

        // Add export button at the top
        html += `
            <div style="margin-bottom: 15px; display: flex; justify-content: flex-end; align-items: center;">
                <button onclick="window.onAGProcessor.exportToPDF()" class="export-btn">
                    Export
                </button>
            </div>
        `;

        for (const fileResult of results) {
            const wastageResults = fileResult.results.wastageResults || [];
            const coatsThreadCheck = fileResult.results.coatsThreadCheck;
            const processCostsCheck = fileResult.results.processCostsCheck;

            // Count valid/invalid sections
            let validSections = 0;
            let totalSections = wastageResults.length;

            // Add Coats Thread to count if found
            if (coatsThreadCheck && coatsThreadCheck.found) {
                totalSections++;
                if (coatsThreadCheck.isValid) validSections++;
            }

            // Add Process Costs to count
            if (processCostsCheck && processCostsCheck.items) {
                totalSections += processCostsCheck.items.length;
                processCostsCheck.items.forEach(item => {
                    if (item.isValid) validSections++;
                });
            }

            for (const section of wastageResults) {
                if (section.found && section.isValid) validSections++;
            }

            html += `<div class="file-result-group">`;

            // File summary
            html += `
                <div class="file-summary-box">
                    <strong>File:</strong> ${fileResult.fileName}<br>
                    <strong>Summary:</strong> ${validSections} out of ${totalSections} sections passed
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

            // Display each wastage section
            for (const section of wastageResults) {
                if (!section.found) {
                    html += `
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">${section.section} Wastage (${section.expectedWastage || '?'}%)</td>
                            <td style="padding: 0.875rem 1rem; text-align: left; color: #991b1b; font-weight: 600;">
                                ${section.message || 'Not found'}
                            </td>
                        </tr>
                    `;
                } else {
                    // Combine valid and invalid cells on one row - show values
                    const allCells = [];

                    // Add valid cells (green) - show value
                    section.validCells.forEach(cell => {
                        allCells.push(`<span style="color: #065f46; font-weight: 600;">${cell.numericValue.toFixed(2)}</span>`);
                    });

                    // Add invalid cells (red with value and expected)
                    section.invalidCells.forEach(cell => {
                        const roundedValue = cell.numericValue.toFixed(2);
                        allCells.push(`<span style="color: #991b1b; font-weight: 600;">${roundedValue}</span> <span style="font-size: 0.85em; color: #849bba;">(Expected: ${section.expectedWastage})</span>`);
                    });

                    if (allCells.length > 0) {
                        html += `
                            <tr style="border-bottom: 1px solid #e0e8f0;">
                                <td style="padding: 0.875rem 1rem; font-weight: 600;">${section.section} Wastage (${section.expectedWastage}%)</td>
                                <td style="padding: 0.875rem 1rem; text-align: left;">
                                    ${allCells.join(', ')}
                                </td>
                            </tr>
                        `;
                    } else {
                        // No cells at all
                        html += `
                            <tr style="border-bottom: 1px solid #e0e8f0;">
                                <td style="padding: 0.875rem 1rem; font-weight: 600;">${section.section} Wastage (${section.expectedWastage}%)</td>
                                <td style="padding: 0.875rem 1rem; text-align: left;">
                                    <span style="color: #6b7280; font-weight: 600;">No data</span>
                                </td>
                            </tr>
                        `;
                    }
                }
            }

            // Display Coats Thread check - show values with labels
            if (coatsThreadCheck && coatsThreadCheck.found && coatsThreadCheck.checks) {
                const checkDetails = coatsThreadCheck.checks.map(check => {
                    const displayValue = !isNaN(check.numericValue) ? check.numericValue : check.actualValue;
                    if (check.isValid) {
                        return `<span style="color: #065f46; font-weight: 600;">${check.label}: ${displayValue}</span>`;
                    } else {
                        return `<span style="color: #991b1b; font-weight: 600;">${check.label}: ${displayValue}</span> <span style="font-size: 0.85em; color: #849bba;">(Expected: ${check.expectedValue})</span>`;
                    }
                }).join(', ');

                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">Coats Thread (Row ${coatsThreadCheck.rowNumber})</td>
                        <td style="padding: 0.875rem 1rem; text-align: left;">
                            ${checkDetails}
                        </td>
                    </tr>
                `;
            } else if (coatsThreadCheck && !coatsThreadCheck.found) {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">Coats Thread</td>
                        <td style="padding: 0.875rem 1rem; text-align: left; color: #6b7280; font-weight: 600;">
                            ${coatsThreadCheck.message || 'Not found in Material section'}
                        </td>
                    </tr>
                `;
            }

            // Display Process Costs checks - show actual values
            if (processCostsCheck && processCostsCheck.items) {
                for (const item of processCostsCheck.items) {
                    if (!item.found) {
                        html += `
                            <tr style="border-bottom: 1px solid #e0e8f0;">
                                <td style="padding: 0.875rem 1rem; font-weight: 600;">${item.label} (${item.expectedValue})</td>
                                <td style="padding: 0.875rem 1rem; text-align: left; color: #991b1b; font-weight: 600;">
                                    Not found
                                </td>
                            </tr>
                        `;
                    } else {
                        const displayValue = !isNaN(item.numericValue) ? item.numericValue : item.actualValue;
                        let cellDisplay;
                        if (item.isValid) {
                            cellDisplay = `<span style="color: #065f46; font-weight: 600;">${displayValue}</span>`;
                        } else {
                            cellDisplay = `<span style="color: #991b1b; font-weight: 600;">${displayValue}</span> <span style="font-size: 0.85em; color: #849bba;">(Expected: ${item.expectedValue})</span>`;
                        }

                        html += `
                            <tr style="border-bottom: 1px solid #e0e8f0;">
                                <td style="padding: 0.875rem 1rem; font-weight: 600;">${item.label} (${item.expectedValue})</td>
                                <td style="padding: 0.875rem 1rem; text-align: left;">
                                    ${cellDisplay}
                                </td>
                            </tr>
                        `;
                    }
                }
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
        const resultsContainer = document.getElementById('results-v9');
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

        const config = window.pdfExporter.createOnAGConfig(this.bcbdResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }
}

// Initialize the processor
window.onAGProcessor = new OnAGProcessor();

// Initialize when V9 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    const v9Tab = document.querySelector('[data-tab="v9"]');
    if (v9Tab) {
        v9Tab.addEventListener('click', () => {
            window.onAGProcessor.initialize();
        });
    }

    // If V9 tab is already active on load, initialize immediately
    const v9TabContent = document.getElementById('tab-v9');
    if (v9TabContent && v9TabContent.classList.contains('active')) {
        window.onAGProcessor.initialize();
    }
});
