/**
 * Haglofs (V21) Processing Logic
 * Validates specific cells in Buyer CBD files against expected values
 *
 * Validation Rules:
 * - Column A "Supplier" -> Column B should be "Madison 88"
 * - Column A "Material / Description" -> Column H (Allowance) should be 5% until "Total Fabric Costs"
 */

class HaglofsProcessor {
    constructor() {
        this.bcbdResults = [];
    }

    /**
     * Initialize - Display validation rules in the OB drop zone
     */
    initialize() {
        this.displayValidationRules();
        console.log('Haglofs Processor initialized');
    }

    /**
     * Display validation rules in the OB drop zone
     */
    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v21');
        if (!obDropZone) return;

        let contentHTML = `
            <div class="burton-cost-container">
                <div class="burton-cost-header">

                </div>
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">Supplier (Column A):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column B: <strong>Madison 88</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">Fabric Section (Material / Description → Total Fabric Costs):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column H (Allowance): <strong>5%</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">Trims Section (2nd Material / Description → Total Trims Costs):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column H (Allowance): <strong>3%</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">Packaging Section (3rd Material / Description → Total Packaging Costs):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column H (Allowance): <strong>3%</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">Generic Packaging (within Packaging Section):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column B: <strong>m88</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column F: <strong>pc</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column G: <strong>1</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column H: <strong>3%</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">Overhead (Column I):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column L: <strong>0.45</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">Margin (Column I):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column L: <strong>0.30 - 0.90</strong></div>
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
     * Validate file against rules
     */
    validateFile(jsonData) {
        const results = {
            supplier: null,
            fabricAllowanceRows: [],
            trimsAllowanceRows: [],
            packagingAllowanceRows: [],
            genericPackaging: null,
            overhead: null,
            margin: null
        };

        const colA = this.columnToIndex('A');
        const colB = this.columnToIndex('B');
        const colF = this.columnToIndex('F');
        const colG = this.columnToIndex('G');
        const colH = this.columnToIndex('H');
        const colI = this.columnToIndex('I');
        const colL = this.columnToIndex('L');

        let scanningSection = null; // 'fabric', 'trims', 'packaging' or null
        let materialDescCount = 0; // Track how many "Material / Description" headers we've seen

        // Scan through all rows
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;

            const colAValue = row[colA] ? row[colA].toString().trim() : '';
            const colALower = colAValue.toLowerCase();
            const colIValue = row[colI] ? row[colI].toString().trim() : '';
            const colILower = colIValue.toLowerCase();

            // Check for Supplier
            if (colALower === 'supplier') {
                const colBValue = row[colB] ? row[colB].toString().trim() : '';
                results.supplier = {
                    found: true,
                    rowIndex: i + 1,
                    expected: 'Madison 88',
                    actual: colBValue,
                    isValid: colBValue === 'Madison 88'
                };
            }

            // Check for Overhead in column I
            if (colILower === 'overhead') {
                const colLValue = row[colL] ? row[colL].toString().trim() : '';
                let isValid = false;

                // Check if it's 0.45 in various formats
                if (colLValue === '0.45' || colLValue === '45%' || colLValue === '45') {
                    isValid = true;
                } else if (!isNaN(parseFloat(colLValue))) {
                    const numValue = parseFloat(colLValue);
                    if (numValue === 0.45 || numValue === 45) {
                        isValid = true;
                    }
                }

                results.overhead = {
                    found: true,
                    rowIndex: i + 1,
                    expected: '0.45',
                    actual: colLValue,
                    isValid: isValid
                };
            }

            // Check for Margin in column I
            if (colILower === 'margin') {
                const colLValue = row[colL] ? row[colL].toString().trim() : '';
                let isValid = false;
                let numValue = null;

                // Parse the value
                if (!isNaN(parseFloat(colLValue))) {
                    numValue = parseFloat(colLValue);

                    // Check if it's in range 0.30 to 0.90 (decimal) or 30 to 90 (percentage)
                    if ((numValue >= 0.30 && numValue <= 0.90) || (numValue >= 30 && numValue <= 90)) {
                        isValid = true;
                    }
                }

                results.margin = {
                    found: true,
                    rowIndex: i + 1,
                    expected: '0.30 - 0.90',
                    actual: colLValue,
                    isValid: isValid
                };
            }

            // Check for Material / Description sections (header row)
            // More flexible matching: just need "material" and "description" in the text
            if (colALower.includes('material') && colALower.includes('description')) {
                materialDescCount++;
                console.log(`Found Material/Description #${materialDescCount} at row ${i + 1}: "${colAValue}"`);

                if (materialDescCount === 1) {
                    // First "Material / Description" → starts Fabric section (5%)
                    scanningSection = 'fabric';
                    console.log('Starting FABRIC section (5%)');
                } else if (materialDescCount === 2) {
                    // Second "Material / Description" → starts Trims section (3%)
                    scanningSection = 'trims';
                    console.log('Starting TRIMS section (3%)');
                } else if (materialDescCount === 3) {
                    // Third "Material / Description" → starts Packaging section (3%)
                    scanningSection = 'packaging';
                    console.log('Starting PACKAGING section (3%)');
                }
                continue; // Skip the header row itself
            }

            // Check if we've reached Total Fabric Costs (end of fabric section)
            if (colALower.includes('total') && colALower.includes('fabric') && colALower.includes('cost')) {
                console.log(`Found Total Fabric Costs at row ${i + 1}, stopping FABRIC section`);
                if (scanningSection === 'fabric') {
                    scanningSection = null; // Stop scanning until next "Material / Description"
                }
                continue;
            }

            // Check if we've reached Total Trims Costs (end of trims section)
            if (colALower.includes('total') && colALower.includes('trims') && colALower.includes('cost')) {
                console.log(`Found Total Trims Costs at row ${i + 1}, stopping TRIMS section`);
                if (scanningSection === 'trims') {
                    scanningSection = null; // Stop scanning
                }
                continue;
            }

            // Check if we've reached Total Packaging Costs (end of packaging section)
            if (colALower.includes('total') && colALower.includes('packaging') && colALower.includes('cost')) {
                console.log(`Found Total Packaging Costs at row ${i + 1}, stopping PACKAGING section`);
                if (scanningSection === 'packaging') {
                    scanningSection = null; // Stop scanning
                }
                continue;
            }

            // Special rule: Check for Generic Packaging in packaging section
            if (scanningSection === 'packaging' && colALower.includes('generic') && colALower.includes('packaging')) {
                console.log(`Found Generic Packaging at row ${i + 1}`);
                const colBValue = row[colB] ? row[colB].toString().trim().toLowerCase() : '';
                const colFValue = row[colF] ? row[colF].toString().trim().toLowerCase() : '';
                const colGValue = row[colG] ? row[colG].toString().trim() : '';
                const colHValue = row[colH] ? row[colH].toString().trim() : '';

                // Check all conditions: B=m88, F=pc, G=1, H=3%
                const bValid = colBValue === 'm88';
                const fValid = colFValue === 'pc';
                const gValid = colGValue === '1';

                let hValid = false;
                if (colHValue === '3%' || colHValue === '0.03' || colHValue === '3') {
                    hValid = true;
                } else if (!isNaN(parseFloat(colHValue))) {
                    const numValue = parseFloat(colHValue);
                    if (numValue === 0.03 || numValue === 3) {
                        hValid = true;
                    }
                }

                const allValid = bValid && fValid && gValid && hValid;

                results.genericPackaging = {
                    rowIndex: i + 1,
                    colB: { expected: 'm88', actual: colBValue, isValid: bValid },
                    colF: { expected: 'pc', actual: colFValue, isValid: fValid },
                    colG: { expected: '1', actual: colGValue, isValid: gValid },
                    colH: { expected: '3%', actual: colHValue, isValid: hValid },
                    isValid: allValid
                };
                continue; // Don't process this row as a normal packaging row
            }

            // If we're scanning for allowance in any section
            if (scanningSection === 'fabric' || scanningSection === 'trims' || scanningSection === 'packaging') {
                // Check column H (Allowance)
                const colHValue = row[colH] ? row[colH].toString().trim() : '';

                // Skip rows where column H is empty - don't validate empty cells
                if (!colHValue || colHValue === '') {
                    continue;
                }

                // Use column A value if present, otherwise indicate it's an empty row
                const description = colAValue !== '' ? colAValue : '(empty row)';

                let expected = '';
                let isValid = false;

                if (scanningSection === 'fabric') {
                    expected = '5%';
                    // Check if it's 5% in various formats
                    if (colHValue === '5%' || colHValue === '0.05' || colHValue === '5') {
                        isValid = true;
                    } else if (!isNaN(parseFloat(colHValue))) {
                        const numValue = parseFloat(colHValue);
                        if (numValue === 0.05 || numValue === 5) {
                            isValid = true;
                        }
                    }
                } else if (scanningSection === 'trims') {
                    expected = '3%';
                    // Check if it's 3% in various formats
                    if (colHValue === '3%' || colHValue === '0.03' || colHValue === '3') {
                        isValid = true;
                    } else if (!isNaN(parseFloat(colHValue))) {
                        const numValue = parseFloat(colHValue);
                        if (numValue === 0.03 || numValue === 3) {
                            isValid = true;
                        }
                    }
                } else if (scanningSection === 'packaging') {
                    expected = '3%';
                    // Check if it's 3% in various formats
                    if (colHValue === '3%' || colHValue === '0.03' || colHValue === '3') {
                        isValid = true;
                    } else if (!isNaN(parseFloat(colHValue))) {
                        const numValue = parseFloat(colHValue);
                        if (numValue === 0.03 || numValue === 3) {
                            isValid = true;
                        }
                    }
                }

                const rowData = {
                    rowIndex: i + 1,
                    description: description,
                    expected: expected,
                    actual: colHValue,
                    isValid: isValid
                };

                if (scanningSection === 'fabric') {
                    results.fabricAllowanceRows.push(rowData);
                } else if (scanningSection === 'trims') {
                    results.trimsAllowanceRows.push(rowData);
                } else if (scanningSection === 'packaging') {
                    results.packagingAllowanceRows.push(rowData);
                }
            }
        }

        // Handle case where Supplier wasn't found
        if (!results.supplier) {
            results.supplier = {
                found: false,
                rowIndex: null,
                expected: 'Madison 88',
                actual: 'Not Found',
                isValid: false
            };
        }

        // Handle case where Overhead wasn't found
        if (!results.overhead) {
            results.overhead = {
                found: false,
                rowIndex: null,
                expected: '0.45',
                actual: 'Not Found',
                isValid: false
            };
        }

        // Handle case where Margin wasn't found
        if (!results.margin) {
            results.margin = {
                found: false,
                rowIndex: null,
                expected: '0.30 - 0.90',
                actual: 'Not Found',
                isValid: false
            };
        }

        return results;
    }

    /**
     * Format field value with color coding (like Burton template)
     */
    formatFieldValue(expectedValue, actualValue, isValid) {
        const color = isValid ? '#065f46' : '#991b1b'; // Green if valid, red if invalid

        if (!actualValue || actualValue === '' || actualValue === 'Empty') {
            return `<span style="color: #991b1b; font-weight: 600;">Empty</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${expectedValue}</span>`;
        }

        if (isValid) {
            return `<span style="color: ${color}; font-weight: 600;">${actualValue}</span>`;
        } else {
            return `<span style="color: ${color}; font-weight: 600;">${actualValue}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${expectedValue}</span>`;
        }
    }

    /**
     * Generate HTML for results display
     */
    generateResultsHTML(fileResults) {
        if (!fileResults || fileResults.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">✓ Haglofs Validation Ready</p>
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
                        oninput="window.haglofsProcessor.searchByFilename(this.value)"
                    />
                </div>
                <button onclick="window.haglofsProcessor.exportToPDF()" class="export-btn">
                    Export
                </button>
            </div>
        `;

        for (const fileResult of fileResults) {
            html += `<div class="file-result-group">`;

            // File header
            html += `
                <div style="margin-bottom: 1rem; padding: 0.75rem; background: #f8fafc; border-left: 4px solid #3b82f6; border-radius: 4px;">
                    <strong style="font-size: 1.05em;">File:</strong> ${fileResult.fileName}
                </div>
            `;

            const supplier = fileResult.results.supplier;
            const fabricRows = fileResult.results.fabricAllowanceRows;
            const trimsRows = fileResult.results.trimsAllowanceRows;
            const packagingRows = fileResult.results.packagingAllowanceRows;
            const genericPackaging = fileResult.results.genericPackaging;
            const overhead = fileResult.results.overhead;
            const margin = fileResult.results.margin;

            // Compile all results into arrays
            const fabricValidResults = [];
            const fabricInvalidResults = [];
            const trimsValidResults = [];
            const trimsInvalidResults = [];
            const packagingValidResults = [];
            const packagingInvalidResults = [];

            for (const row of fabricRows) {
                if (row.isValid) {
                    fabricValidResults.push(`Row ${row.rowIndex}`);
                } else {
                    fabricInvalidResults.push(`Row ${row.rowIndex}: ${row.actual} (Expected: ${row.expected})`);
                }
            }

            for (const row of trimsRows) {
                if (row.isValid) {
                    trimsValidResults.push(`Row ${row.rowIndex}`);
                } else {
                    trimsInvalidResults.push(`Row ${row.rowIndex}: ${row.actual} (Expected: ${row.expected})`);
                }
            }

            for (const row of packagingRows) {
                if (row.isValid) {
                    packagingValidResults.push(`Row ${row.rowIndex}`);
                } else {
                    packagingInvalidResults.push(`Row ${row.rowIndex}: ${row.actual} (Expected: ${row.expected})`);
                }
            }

            // Build fabric results HTML
            let fabricResultsHTML = '';
            if (fabricRows.length === 0) {
                fabricResultsHTML = '<span style="color: #d97706; font-weight: 600;">No section found</span>';
            } else {
                if (fabricValidResults.length > 0) {
                    fabricResultsHTML += `<span style="color: #065f46; font-weight: 600;">${fabricValidResults.join(', ')}</span>`;
                }
                if (fabricInvalidResults.length > 0) {
                    if (fabricValidResults.length > 0) fabricResultsHTML += '<br>';
                    fabricResultsHTML += `<span style="color: #991b1b; font-weight: 600;">${fabricInvalidResults.join(' | ')}</span>`;
                }
            }

            // Build trims results HTML
            let trimsResultsHTML = '';
            if (trimsRows.length === 0) {
                trimsResultsHTML = '<span style="color: #d97706; font-weight: 600;">No section found</span>';
            } else {
                if (trimsValidResults.length > 0) {
                    trimsResultsHTML += `<span style="color: #065f46; font-weight: 600;">${trimsValidResults.join(', ')}</span>`;
                }
                if (trimsInvalidResults.length > 0) {
                    if (trimsValidResults.length > 0) trimsResultsHTML += '<br>';
                    trimsResultsHTML += `<span style="color: #991b1b; font-weight: 600;">${trimsInvalidResults.join(' | ')}</span>`;
                }
            }

            // Build packaging results HTML
            let packagingResultsHTML = '';
            if (packagingRows.length === 0) {
                packagingResultsHTML = '<span style="color: #d97706; font-weight: 600;">No section found</span>';
            } else {
                if (packagingValidResults.length > 0) {
                    packagingResultsHTML += `<span style="color: #065f46; font-weight: 600;">${packagingValidResults.join(', ')}</span>`;
                }
                if (packagingInvalidResults.length > 0) {
                    if (packagingValidResults.length > 0) packagingResultsHTML += '<br>';
                    packagingResultsHTML += `<span style="color: #991b1b; font-weight: 600;">${packagingInvalidResults.join(' | ')}</span>`;
                }
            }

            // Build Generic Packaging HTML - show all fields individually
            let genericPackagingHTML = '';
            if (genericPackaging) {
                const parts = [];

                // Column B
                const bColor = genericPackaging.colB.isValid ? '#065f46' : '#991b1b';
                const bText = genericPackaging.colB.isValid
                    ? `B: ${genericPackaging.colB.actual}`
                    : `B: ${genericPackaging.colB.actual} (Exp: ${genericPackaging.colB.expected})`;
                parts.push(`<span style="color: ${bColor}; font-weight: 600;">${bText}</span>`);

                // Column F
                const fColor = genericPackaging.colF.isValid ? '#065f46' : '#991b1b';
                const fText = genericPackaging.colF.isValid
                    ? `F: ${genericPackaging.colF.actual}`
                    : `F: ${genericPackaging.colF.actual} (Exp: ${genericPackaging.colF.expected})`;
                parts.push(`<span style="color: ${fColor}; font-weight: 600;">${fText}</span>`);

                // Column G
                const gColor = genericPackaging.colG.isValid ? '#065f46' : '#991b1b';
                const gText = genericPackaging.colG.isValid
                    ? `G: ${genericPackaging.colG.actual}`
                    : `G: ${genericPackaging.colG.actual} (Exp: ${genericPackaging.colG.expected})`;
                parts.push(`<span style="color: ${gColor}; font-weight: 600;">${gText}</span>`);

                // Column H
                const hColor = genericPackaging.colH.isValid ? '#065f46' : '#991b1b';
                const hText = genericPackaging.colH.isValid
                    ? `H: ${genericPackaging.colH.actual}`
                    : `H: ${genericPackaging.colH.actual} (Exp: ${genericPackaging.colH.expected})`;
                parts.push(`<span style="color: ${hColor}; font-weight: 600;">${hText}</span>`);

                genericPackagingHTML = `Row ${genericPackaging.rowIndex}: ${parts.join(', ')}`;
            } else {
                genericPackagingHTML = '<span style="color: #d97706; font-weight: 600;">Not found</span>';
            }

            // Single combined validation table
            html += `
                <table id="v21ValidationTable" class="results-table" style="margin-top: 1.5rem;">
                    <thead>
                        <tr class="header-labels-row">
                            <th style="color: inherit;">Validation Field</th>
                            <th style="color: inherit;">Results</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">Supplier (Row ${supplier.rowIndex || 'N/A'})</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(supplier.expected, supplier.actual, supplier.isValid)}</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">Fabric (5%)</td>
                            <td style="padding: 0.875rem 1rem;">${fabricResultsHTML}</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">Trims (3%)</td>
                            <td style="padding: 0.875rem 1rem;">${trimsResultsHTML}</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">Packaging (3%)</td>
                            <td style="padding: 0.875rem 1rem;">${packagingResultsHTML}</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">Generic Packaging</td>
                            <td style="padding: 0.875rem 1rem;">${genericPackagingHTML}</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">Overhead (Row ${overhead.rowIndex || 'N/A'})</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(overhead.expected, overhead.actual, overhead.isValid)}</td>
                        </tr>
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">Margin (Row ${margin.rowIndex || 'N/A'})</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(margin.expected, margin.actual, margin.isValid)}</td>
                        </tr>
                    </tbody>
                </table>
            `;

            html += `</div>`; // Close file-result-group
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
                    ❌ Error Processing Files
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
        if (!window.haglofsExportConfig) {
            console.error('Haglofs export configuration not loaded');
            alert('Export module not available. Please refresh the page.');
            return;
        }

        if (!this.bcbdResults || this.bcbdResults.length === 0) {
            alert('No results to export. Please generate results first.');
            return;
        }

        const config = window.haglofsExportConfig.createHaglofsConfig(this.bcbdResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }

    /**
     * Search by filename - filters file result groups based on filename
     */
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
window.haglofsProcessor = new HaglofsProcessor();
