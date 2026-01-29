/**
 * FOX (V20) Processing Logic
 * Validates specific cells in Buyer CBD files against expected values
 *
 * Validation Rules:
 * - Column C "Vendor" -> Column D should be "Madison 88 Ltd"
 * - Column C "Factory" -> Column D should be "PT UWU Jump"
 * - Column C "COO" -> Column D should be "Indonesia"
 * - Column K "OVERHEAD" -> Column L should be 0.40
 * - Column K "PROFIT & OTHERS" -> Column L should be 0.35-0.45
 * - Column A "FABRIC / UPPER / SHELL" -> Column E (Wastage%) should be 5% until Column H "SUBTOTAL"
 * - Column B "Sewing Thread" (within FABRIC/UPPER/SHELL section) -> D=1, E=3%, H=0.01, I=0.01, J=0%
 * - Column A "Standard Packaging" -> D=1, E=3%
 * - Column A "LABOR COST" -> Column B should have Knitting, Sewing, Finishing
 * - Column A "OVERHEAD COST" -> Column H should be 0.40
 * - Column A "PROFIT COST" -> Column H should be 0.35-0.45
 */

class FOXProcessor {
    constructor() {
        this.bcbdResults = [];
    }

    /**
     * Initialize - Display validation rules in the OB drop zone
     */
    initialize() {
        this.displayValidationRules();
        console.log('FOX Processor initialized');
    }

    /**
     * Display validation rules in the OB drop zone
     */
    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v20');
        if (!obDropZone) return;

        let contentHTML = `
            <div class="burton-cost-container">
                <div class="burton-cost-header">
                </div>
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">Vendor (Column C):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column D: <strong>Madison 88 Ltd</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">Factory (Column C):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column D: <strong>PT UWU Jump</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">COO (Column C):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column D: <strong>Indonesia</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">OVERHEAD (Column K):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column L: <strong>0.40</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">PROFIT & OTHERS (Column K):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column L: <strong>0.35 - 0.45</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">FABRIC / UPPER / SHELL (Column A):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column E (Wastage %): <strong>5%</strong> until SUBTOTAL</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">Sewing Thread (Column B):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">D (Usage): <strong>1</strong>, E (Wastage): <strong>3%</strong>, H (COST CIF): <strong>0.01</strong>, I (Extended Cost): <strong>0.01</strong>, J (% to Total): <strong>0%</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">Standard Packaging (Column A):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">D (Usage): <strong>1</strong>, E (Wastage): <strong>3%</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">LABOR COST (Column A):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column B: <strong>Knitting</strong>, <strong>Sewing</strong>, <strong>Finishing</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">OVERHEAD COST (Column A):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column H: <strong>0.40</strong></div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong style="color: #2b4a6c;">PROFIT COST (Column A):</strong></div>
                        <div class="burton-item-line" style="margin-left: 1rem;">Column H: <strong>0.35 - 0.45</strong></div>
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
            vendor: null,
            factory: null,
            coo: null,
            overhead: null,
            profitOthers: null,
            wastagePercent: null,
            sewingThread: null,
            standardPackaging: null,
            laborCost: null,
            overheadCost: null,
            profitCost: null
        };

        const colA = this.columnToIndex('A');
        const colB = this.columnToIndex('B');
        const colC = this.columnToIndex('C');
        const colD = this.columnToIndex('D');
        const colE = this.columnToIndex('E');
        const colH = this.columnToIndex('H');
        const colI = this.columnToIndex('I');
        const colJ = this.columnToIndex('J');
        const colK = this.columnToIndex('K');
        const colL = this.columnToIndex('L');

        // Scan through all rows to find the labels
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;

            // Check Column C for labels
            const colCValue = row[colC] ? row[colC].toString().trim() : '';
            const colCLower = colCValue.toLowerCase();

            // Check for "Vendor" in Column C
            if (colCLower === 'vendor' && !results.vendor) {
                const colDValue = row[colD] ? row[colD].toString().trim() : '';
                const expectedValue = 'Madison 88 Ltd.';
                const isValid = colDValue.toLowerCase() === expectedValue.toLowerCase();

                results.vendor = {
                    label: 'Vendor',
                    labelCell: `C${i + 1}`,
                    valueCell: `D${i + 1}`,
                    actualValue: colDValue,
                    expectedValue: expectedValue,
                    isValid: isValid
                };
            }

            // Check for "Factory" in Column C
            if (colCLower === 'factory' && !results.factory) {
                const colDValue = row[colD] ? row[colD].toString().trim() : '';
                const expectedValue = 'PT UWU Jump';
                const isValid = colDValue.toLowerCase() === expectedValue.toLowerCase();

                results.factory = {
                    label: 'Factory',
                    labelCell: `C${i + 1}`,
                    valueCell: `D${i + 1}`,
                    actualValue: colDValue,
                    expectedValue: expectedValue,
                    isValid: isValid
                };
            }

            // Check for "COO" in Column C
            if (colCLower === 'coo' && !results.coo) {
                const colDValue = row[colD] ? row[colD].toString().trim() : '';
                const expectedValue = 'Indonesia';
                const isValid = colDValue.toLowerCase() === expectedValue.toLowerCase();

                results.coo = {
                    label: 'COO',
                    labelCell: `C${i + 1}`,
                    valueCell: `D${i + 1}`,
                    actualValue: colDValue,
                    expectedValue: expectedValue,
                    isValid: isValid
                };
            }

            // Check Column K for labels
            const colKValue = row[colK] ? row[colK].toString().trim() : '';
            const colKLower = colKValue.toLowerCase();

            // Check for "OVERHEAD" in Column K
            if (colKLower === 'overhead' && !results.overhead) {
                const colLValue = row[colL];
                const expectedValue = 0.40;
                let actualNumeric = null;
                let displayValue = '';

                if (colLValue !== undefined && colLValue !== null && colLValue !== '') {
                    if (typeof colLValue === 'number') {
                        actualNumeric = colLValue;
                        displayValue = colLValue.toFixed(4);
                    } else {
                        const parsed = parseFloat(colLValue.toString().trim());
                        if (!isNaN(parsed)) {
                            actualNumeric = parsed;
                            displayValue = parsed.toFixed(4);
                        } else {
                            displayValue = colLValue.toString().trim();
                        }
                    }
                } else {
                    displayValue = 'Empty';
                }

                // STRICT comparison - must be exactly 0.40 (with tiny tolerance for floating point)
                const isValid = actualNumeric !== null && Math.abs(actualNumeric - expectedValue) < 0.0001;

                console.log(`OVERHEAD: value=${actualNumeric}, expected=${expectedValue}, isValid=${isValid}`);

                results.overhead = {
                    label: 'OVERHEAD',
                    labelCell: `K${i + 1}`,
                    valueCell: `L${i + 1}`,
                    actualValue: displayValue,
                    expectedValue: '0.40',
                    isValid: isValid
                };
            }

            // Check for "PROFIT & OTHERS" in Column K
            if (colKLower.includes('profit') && colKLower.includes('others') && !results.profitOthers) {
                const colLValue = row[colL];
                const expectedMin = 0.35;
                const expectedMax = 0.45;
                let actualNumeric = null;
                let displayValue = '';

                if (colLValue !== undefined && colLValue !== null && colLValue !== '') {
                    if (typeof colLValue === 'number') {
                        actualNumeric = colLValue;
                        displayValue = colLValue.toFixed(4);
                    } else {
                        const parsed = parseFloat(colLValue.toString().trim());
                        if (!isNaN(parsed)) {
                            actualNumeric = parsed;
                            displayValue = parsed.toFixed(4);
                        } else {
                            displayValue = colLValue.toString().trim();
                        }
                    }
                } else {
                    displayValue = 'Empty';
                }

                // STRICT comparison - must be between 0.35 and 0.45 exactly
                const isValid = actualNumeric !== null && actualNumeric >= expectedMin - 0.0001 && actualNumeric <= expectedMax + 0.0001;

                console.log(`PROFIT & OTHERS: value=${actualNumeric}, min=${expectedMin}, max=${expectedMax}, isValid=${isValid}`);

                results.profitOthers = {
                    label: 'PROFIT & OTHERS',
                    labelCell: `K${i + 1}`,
                    valueCell: `L${i + 1}`,
                    actualValue: displayValue,
                    expectedValue: '0.35 - 0.45',
                    isValid: isValid
                };
            }
        }

        // Validate Wastage % (Column E) for FABRIC / UPPER / SHELL section
        results.wastagePercent = this.validateWastagePercent(jsonData, colA, colB, colE, colH);

        // Validate Sewing Thread within FABRIC / UPPER / SHELL section
        results.sewingThread = this.validateSewingThread(jsonData, colA, colB, colD, colE, colH, colI, colJ);

        // Validate Standard Packaging rows
        results.standardPackaging = this.validateStandardPackaging(jsonData, colA, colD, colE);

        // Validate Labor Cost section
        results.laborCost = this.validateLaborCost(jsonData, colA, colB);

        // Validate OVERHEAD COST (Column A) -> Column H = 0.40
        results.overheadCost = this.validateOverheadCost(jsonData, colA, colH);

        // Validate PROFIT COST (Column A) -> Column H = 0.35-0.45
        results.profitCost = this.validateProfitCost(jsonData, colA, colH);

        // Set defaults for not found items
        if (!results.vendor) {
            results.vendor = {
                label: 'Vendor',
                labelCell: '-',
                valueCell: '-',
                actualValue: 'Not found',
                expectedValue: 'Madison 88 Ltd.',
                isValid: false
            };
        }
        if (!results.factory) {
            results.factory = {
                label: 'Factory',
                labelCell: '-',
                valueCell: '-',
                actualValue: 'Not found',
                expectedValue: 'PT UWU Jump',
                isValid: false
            };
        }
        if (!results.coo) {
            results.coo = {
                label: 'COO',
                labelCell: '-',
                valueCell: '-',
                actualValue: 'Not found',
                expectedValue: 'Indonesia',
                isValid: false
            };
        }
        if (!results.overhead) {
            results.overhead = {
                label: 'OVERHEAD',
                labelCell: '-',
                valueCell: '-',
                actualValue: 'Not found',
                expectedValue: '0.40',
                isValid: false
            };
        }
        if (!results.profitOthers) {
            results.profitOthers = {
                label: 'PROFIT & OTHERS',
                labelCell: '-',
                valueCell: '-',
                actualValue: 'Not found',
                expectedValue: '0.35 - 0.45',
                isValid: false
            };
        }

        if (!results.wastagePercent) {
            results.wastagePercent = {
                label: 'Wastage % (FABRIC/UPPER/SHELL)',
                labelCell: '-',
                valueCell: '-',
                expectedValue: '5%',
                isValid: false,
                invalidRows: [],
                validRows: []
            };
        }

        if (!results.sewingThread) {
            results.sewingThread = {
                label: 'Sewing Thread',
                labelCell: '-',
                valueCell: '-',
                expectedValue: 'D=1, E=3%, H=0.01, I=0.01, J=0%',
                isValid: false,
                notFound: true,
                validFields: [],
                invalidFields: []
            };
        }

        if (!results.standardPackaging) {
            results.standardPackaging = {
                label: 'Standard Packaging',
                labelCell: '-',
                valueCell: '-',
                expectedValue: 'D=1, E=3%',
                isValid: false,
                notFound: true,
                validFields: [],
                invalidFields: []
            };
        }

        if (!results.laborCost) {
            results.laborCost = {
                label: 'Labor Cost',
                labelCell: '-',
                valueCell: '-',
                expectedValue: 'Knitting, Sewing, Finishing',
                isValid: false,
                notFound: true,
                foundItems: [],
                missingItems: ['Knitting', 'Sewing', 'Finishing']
            };
        }

        if (!results.overheadCost) {
            results.overheadCost = {
                label: 'OVERHEAD COST',
                labelCell: '-',
                valueCell: '-',
                actualValue: 'Not found',
                expectedValue: '0.40',
                isValid: false
            };
        }

        if (!results.profitCost) {
            results.profitCost = {
                label: 'PROFIT COST',
                labelCell: '-',
                valueCell: '-',
                actualValue: 'Not found',
                expectedValue: '0.35 - 0.45',
                isValid: false
            };
        }

        return results;
    }

    /**
     * Validate Wastage % in Column E for FABRIC / UPPER / SHELL section
     * Scans Column A for "FABRIC", "UPPER", or "SHELL"
     * Then checks Column E values are 5% until Column H contains "SUBTOTAL"
     * Skips rows that contain "Sewing Thread" in Column B (those have different rules)
     */
    validateWastagePercent(jsonData, colA, colB, colE, colH) {
        let sectionStartRow = -1;
        let sectionHeaderText = '';

        // Step 1: Find the row where Column A contains FABRIC, UPPER, or SHELL
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;

            const colAValue = row[colA] ? row[colA].toString().trim().toUpperCase() : '';

            if (colAValue.includes('FABRIC') || colAValue.includes('UPPER') || colAValue.includes('SHELL')) {
                sectionStartRow = i;
                sectionHeaderText = colAValue;
                console.log(`Found section header "${sectionHeaderText}" at row ${i + 1}`);
                break;
            }
        }

        // If section not found, return not found result
        if (sectionStartRow === -1) {
            return null;
        }

        // Step 2: Check Column E values from the row after the header until SUBTOTAL in Column H
        const invalidRows = [];
        const validRows = [];
        let subtotalFound = false;

        for (let i = sectionStartRow + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;

            // Check if Column H contains SUBTOTAL - if so, stop scanning
            const colHValue = row[colH] ? row[colH].toString().trim().toUpperCase() : '';
            if (colHValue.includes('SUBTOTAL')) {
                subtotalFound = true;
                console.log(`Found SUBTOTAL at row ${i + 1}, stopping wastage scan`);
                break;
            }

            // Skip rows that contain "Sewing Thread" in Column B (they have different wastage rules)
            const colBValue = row[colB] ? row[colB].toString().trim().toUpperCase() : '';
            if (colBValue.includes('SEWING THREAD')) {
                console.log(`Skipping Sewing Thread row ${i + 1} from 5% wastage check`);
                continue;
            }

            // Get Column E value (Wastage %)
            const colEValue = row[colE];
            let numericValue = null;

            if (colEValue !== undefined && colEValue !== null && colEValue !== '') {
                if (typeof colEValue === 'number') {
                    numericValue = colEValue;
                } else {
                    const parsed = parseFloat(colEValue.toString().replace('%', '').trim());
                    if (!isNaN(parsed)) {
                        // If value is like "5" (without decimal), treat as percentage
                        numericValue = parsed > 1 ? parsed / 100 : parsed;
                    }
                }

                // Check if value is 5% (0.05) with tolerance
                const expectedValue = 0.05;
                const isValid = numericValue !== null && Math.abs(numericValue - expectedValue) < 0.0001;

                if (isValid) {
                    validRows.push({
                        cell: `E${i + 1}`
                    });
                } else {
                    const displayValue = numericValue !== null ? (numericValue * 100).toFixed(2) + '%' : colEValue.toString();
                    invalidRows.push({
                        cell: `E${i + 1}`,
                        value: displayValue
                    });
                    console.log(`Invalid wastage at E${i + 1}: ${displayValue} (expected 5%)`);
                }
            }
        }

        const allValid = invalidRows.length === 0 && validRows.length > 0;

        return {
            label: 'Wastage % (FABRIC/UPPER/SHELL)',
            labelCell: `A${sectionStartRow + 1}`,
            valueCell: 'Column E',
            expectedValue: '5%',
            isValid: allValid,
            invalidRows: invalidRows,
            validRows: validRows,
            sectionHeader: sectionHeaderText,
            subtotalFound: subtotalFound
        };
    }

    /**
     * Validate Sewing Thread within FABRIC / UPPER / SHELL section
     * Checks: D (Usage) = 1, E (Wastage) = 3%, H (COST CIF) = 0.01, I (Extended Cost) = 0.01, J (% to Total) = 0%
     * Supports multiple Sewing Thread rows in a single file
     */
    validateSewingThread(jsonData, colA, colB, colD, colE, colH, colI, colJ) {
        let sectionStartRow = -1;

        // Step 1: Find the row where Column A contains FABRIC, UPPER, or SHELL
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;

            const colAValue = row[colA] ? row[colA].toString().trim().toUpperCase() : '';

            if (colAValue.includes('FABRIC') || colAValue.includes('UPPER') || colAValue.includes('SHELL')) {
                sectionStartRow = i;
                break;
            }
        }

        // If section not found, return null
        if (sectionStartRow === -1) {
            return null;
        }

        // Step 2: Search for ALL "Sewing Thread" rows in Column B within the section
        const sewingThreadRows = [];

        for (let i = sectionStartRow + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;

            // Check if Column H contains SUBTOTAL - if so, stop scanning
            const colHValue = row[colH] ? row[colH].toString().trim().toUpperCase() : '';
            if (colHValue.includes('SUBTOTAL')) {
                break;
            }

            // Check Column B for "Sewing Thread"
            const colBValue = row[colB] ? row[colB].toString().trim().toUpperCase() : '';
            if (colBValue.includes('SEWING THREAD')) {
                sewingThreadRows.push(i);
                console.log(`Found Sewing Thread at row ${i + 1}`);
            }
        }

        // If no Sewing Thread found in section
        if (sewingThreadRows.length === 0) {
            return {
                label: 'Sewing Thread',
                labelCell: '-',
                valueCell: '-',
                expectedValue: 'D=1, E=3%, H=0.01, I=0.01, J=0%',
                isValid: false,
                notFound: true,
                validFields: [],
                invalidFields: []
            };
        }

        // Helper function to get numeric value
        const getNumeric = (value) => {
            if (value === undefined || value === null || value === '') return null;
            if (typeof value === 'number') return value;
            const parsed = parseFloat(value.toString().replace('%', '').trim());
            return isNaN(parsed) ? null : parsed;
        };

        // Step 3: Validate each column for ALL Sewing Thread rows
        const validFields = [];
        const invalidFields = [];

        for (const sewingThreadRow of sewingThreadRows) {
            const row = jsonData[sewingThreadRow];
            const rowNum = sewingThreadRow + 1;

            // Check Column D (Usage) = 1
            const colDValue = getNumeric(row[colD]);
            const colDDisplay = colDValue !== null ? colDValue.toString() : 'Empty';
            if (colDValue !== null && Math.abs(colDValue - 1) < 0.0001) {
                validFields.push({ cell: `D${rowNum}`, field: 'Usage', value: colDDisplay });
            } else {
                invalidFields.push({ cell: `D${rowNum}`, field: 'Usage', value: colDDisplay, expected: '1' });
            }

            // Check Column E (Wastage) = 3% (0.03)
            let colEValue = getNumeric(row[colE]);
            if (colEValue !== null && colEValue > 1) {
                colEValue = colEValue / 100; // Convert from percentage
            }
            const colEDisplay = colEValue !== null ? (colEValue * 100).toFixed(2) + '%' : 'Empty';
            if (colEValue !== null && Math.abs(colEValue - 0.03) < 0.0001) {
                validFields.push({ cell: `E${rowNum}`, field: 'Wastage', value: colEDisplay });
            } else {
                invalidFields.push({ cell: `E${rowNum}`, field: 'Wastage', value: colEDisplay, expected: '3%' });
            }

            // Check Column H (COST CIF) = 0.01
            const colHVal = getNumeric(row[colH]);
            const colHDisplay = colHVal !== null ? colHVal.toString() : 'Empty';
            if (colHVal !== null && Math.abs(colHVal - 0.01) < 0.0001) {
                validFields.push({ cell: `H${rowNum}`, field: 'COST CIF', value: colHDisplay });
            } else {
                invalidFields.push({ cell: `H${rowNum}`, field: 'COST CIF', value: colHDisplay, expected: '0.01' });
            }

            // Check Column I (Extended Cost) = 0.01
            const colIValue = getNumeric(row[colI]);
            const colIDisplay = colIValue !== null ? colIValue.toString() : 'Empty';
            if (colIValue !== null && Math.abs(colIValue - 0.01) < 0.0001) {
                validFields.push({ cell: `I${rowNum}`, field: 'Extended Cost', value: colIDisplay });
            } else {
                invalidFields.push({ cell: `I${rowNum}`, field: 'Extended Cost', value: colIDisplay, expected: '0.01' });
            }

            // Check Column J (% to Total) = 0% (0)
            let colJValue = getNumeric(row[colJ]);
            if (colJValue !== null && colJValue > 1) {
                colJValue = colJValue / 100; // Convert from percentage
            }
            const colJDisplay = colJValue !== null ? (colJValue * 100).toFixed(2) + '%' : 'Empty';
            if (colJValue !== null && Math.abs(colJValue) < 0.0001) {
                validFields.push({ cell: `J${rowNum}`, field: '% to Total', value: colJDisplay });
            } else {
                invalidFields.push({ cell: `J${rowNum}`, field: '% to Total', value: colJDisplay, expected: '0%' });
            }
        }

        const allValid = invalidFields.length === 0;
        const rowsLabel = sewingThreadRows.map(r => r + 1).join(', ');

        return {
            label: 'Sewing Thread',
            labelCell: sewingThreadRows.map(r => `B${r + 1}`).join(', '),
            valueCell: `Row${sewingThreadRows.length > 1 ? 's' : ''} ${rowsLabel}`,
            expectedValue: 'D=1, E=3%, H=0.01, I=0.01, J=0%',
            isValid: allValid,
            notFound: false,
            validFields: validFields,
            invalidFields: invalidFields,
            rowCount: sewingThreadRows.length
        };
    }

    /**
     * Validate Standard Packaging rows
     * Searches Column A for cells containing "Standard Packaging"
     * Checks: D (Usage) = 1, E (Wastage) = 3%
     * Supports multiple Standard Packaging rows
     */
    validateStandardPackaging(jsonData, colA, colD, colE) {
        // Helper function to get numeric value
        const getNumeric = (value) => {
            if (value === undefined || value === null || value === '') return null;
            if (typeof value === 'number') return value;
            const parsed = parseFloat(value.toString().replace('%', '').trim());
            return isNaN(parsed) ? null : parsed;
        };

        // Find all rows where Column A contains "Standard Packaging"
        const standardPackagingRows = [];

        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;

            const colAValue = row[colA] ? row[colA].toString().trim().toUpperCase() : '';
            if (colAValue.includes('STANDARD PACKAGING')) {
                standardPackagingRows.push(i);
                console.log(`Found Standard Packaging at row ${i + 1}`);
            }
        }

        // If no Standard Packaging found
        if (standardPackagingRows.length === 0) {
            return {
                label: 'Standard Packaging',
                labelCell: '-',
                valueCell: '-',
                expectedValue: 'D=1, E=3%',
                isValid: false,
                notFound: true,
                validFields: [],
                invalidFields: []
            };
        }

        // Validate each Standard Packaging row
        const validFields = [];
        const invalidFields = [];

        for (const rowIndex of standardPackagingRows) {
            const row = jsonData[rowIndex];
            const rowNum = rowIndex + 1;

            // Check Column D (Usage) = 1
            const colDValue = getNumeric(row[colD]);
            const colDDisplay = colDValue !== null ? colDValue.toString() : 'Empty';
            if (colDValue !== null && Math.abs(colDValue - 1) < 0.0001) {
                validFields.push({ cell: `D${rowNum}`, field: 'Usage', value: colDDisplay });
            } else {
                invalidFields.push({ cell: `D${rowNum}`, field: 'Usage', value: colDDisplay, expected: '1' });
            }

            // Check Column E (Wastage) = 3% (0.03)
            let colEValue = getNumeric(row[colE]);
            if (colEValue !== null && colEValue > 1) {
                colEValue = colEValue / 100; // Convert from percentage
            }
            const colEDisplay = colEValue !== null ? (colEValue * 100).toFixed(2) + '%' : 'Empty';
            if (colEValue !== null && Math.abs(colEValue - 0.03) < 0.0001) {
                validFields.push({ cell: `E${rowNum}`, field: 'Wastage', value: colEDisplay });
            } else {
                invalidFields.push({ cell: `E${rowNum}`, field: 'Wastage', value: colEDisplay, expected: '3%' });
            }
        }

        const allValid = invalidFields.length === 0;
        const rowsLabel = standardPackagingRows.map(r => r + 1).join(', ');

        return {
            label: 'Standard Packaging',
            labelCell: standardPackagingRows.map(r => `A${r + 1}`).join(', '),
            valueCell: `Row${standardPackagingRows.length > 1 ? 's' : ''} ${rowsLabel}`,
            expectedValue: 'D=1, E=3%',
            isValid: allValid,
            notFound: false,
            validFields: validFields,
            invalidFields: invalidFields,
            rowCount: standardPackagingRows.length
        };
    }

    /**
     * Validate Labor Cost section
     * Searches Column A for "LABOR COST"
     * Then checks Column B for Knitting, Sewing, and Finishing in the rows below
     */
    validateLaborCost(jsonData, colA, colB) {
        // Find the row where Column A contains "LABOR COST"
        let laborCostRow = -1;

        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;

            const colAValue = row[colA] ? row[colA].toString().trim().toUpperCase() : '';
            if (colAValue.includes('LABOR COST')) {
                laborCostRow = i;
                console.log(`Found LABOR COST at row ${i + 1}`);
                break;
            }
        }

        // If LABOR COST not found
        if (laborCostRow === -1) {
            return {
                label: 'Labor Cost',
                labelCell: '-',
                valueCell: '-',
                expectedValue: 'Knitting, Sewing, Finishing',
                isValid: false,
                notFound: true,
                foundItems: [],
                missingItems: ['Knitting', 'Sewing', 'Finishing']
            };
        }

        // Check the same row and next few rows for Knitting, Sewing, Finishing in Column B
        const requiredItems = ['Knitting', 'Sewing', 'Finishing'];
        const foundItems = [];
        const missingItems = [];

        // Search starting from the LABOR COST row itself (check up to 10 rows)
        const maxRowsToCheck = Math.min(laborCostRow + 10, jsonData.length);

        for (const item of requiredItems) {
            let found = false;
            let foundCell = '';

            // Start from laborCostRow (same row) instead of laborCostRow + 1
            for (let i = laborCostRow; i < maxRowsToCheck; i++) {
                const row = jsonData[i];
                if (!row) continue;

                const colBValue = row[colB] ? row[colB].toString().trim() : '';
                if (colBValue.toUpperCase().includes(item.toUpperCase())) {
                    found = true;
                    foundCell = `B${i + 1}`;
                    foundItems.push({ item: item, cell: foundCell, value: colBValue });
                    console.log(`Found ${item} at ${foundCell}`);
                    break;
                }
            }

            if (!found) {
                missingItems.push(item);
                console.log(`Missing ${item} in LABOR COST section`);
            }
        }

        const allValid = missingItems.length === 0;

        return {
            label: 'Labor Cost',
            labelCell: `A${laborCostRow + 1}`,
            valueCell: 'Column B',
            expectedValue: 'Knitting, Sewing, Finishing',
            isValid: allValid,
            notFound: false,
            foundItems: foundItems,
            missingItems: missingItems
        };
    }

    /**
     * Validate OVERHEAD COST
     * Searches Column A for "OVERHEAD COST"
     * Checks Column H should be 0.40
     */
    validateOverheadCost(jsonData, colA, colH) {
        // Find the row where Column A contains "OVERHEAD COST"
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;

            const colAValue = row[colA] ? row[colA].toString().trim().toUpperCase() : '';
            if (colAValue.includes('OVERHEAD COST')) {
                console.log(`Found OVERHEAD COST at row ${i + 1}`);

                const colHValue = row[colH];
                const expectedValue = 0.40;
                let actualNumeric = null;
                let displayValue = '';

                if (colHValue !== undefined && colHValue !== null && colHValue !== '') {
                    if (typeof colHValue === 'number') {
                        actualNumeric = colHValue;
                        displayValue = colHValue.toFixed(4);
                    } else {
                        const parsed = parseFloat(colHValue.toString().trim());
                        if (!isNaN(parsed)) {
                            actualNumeric = parsed;
                            displayValue = parsed.toFixed(4);
                        } else {
                            displayValue = colHValue.toString().trim();
                        }
                    }
                } else {
                    displayValue = 'Empty';
                }

                const isValid = actualNumeric !== null && Math.abs(actualNumeric - expectedValue) < 0.0001;

                console.log(`OVERHEAD COST: value=${actualNumeric}, expected=${expectedValue}, isValid=${isValid}`);

                return {
                    label: 'OVERHEAD COST',
                    labelCell: `A${i + 1}`,
                    valueCell: `H${i + 1}`,
                    actualValue: displayValue,
                    expectedValue: '0.40',
                    isValid: isValid
                };
            }
        }

        return null;
    }

    /**
     * Validate PROFIT COST
     * Searches Column A for "PROFIT COST"
     * Checks Column H should be 0.35 to 0.45
     */
    validateProfitCost(jsonData, colA, colH) {
        // Find the row where Column A contains "PROFIT COST"
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;

            const colAValue = row[colA] ? row[colA].toString().trim().toUpperCase() : '';
            if (colAValue.includes('PROFIT COST')) {
                console.log(`Found PROFIT COST at row ${i + 1}`);

                const colHValue = row[colH];
                const expectedMin = 0.35;
                const expectedMax = 0.45;
                let actualNumeric = null;
                let displayValue = '';

                if (colHValue !== undefined && colHValue !== null && colHValue !== '') {
                    if (typeof colHValue === 'number') {
                        actualNumeric = colHValue;
                        displayValue = colHValue.toFixed(4);
                    } else {
                        const parsed = parseFloat(colHValue.toString().trim());
                        if (!isNaN(parsed)) {
                            actualNumeric = parsed;
                            displayValue = parsed.toFixed(4);
                        } else {
                            displayValue = colHValue.toString().trim();
                        }
                    }
                } else {
                    displayValue = 'Empty';
                }

                const isValid = actualNumeric !== null && actualNumeric >= expectedMin - 0.0001 && actualNumeric <= expectedMax + 0.0001;

                console.log(`PROFIT COST: value=${actualNumeric}, min=${expectedMin}, max=${expectedMax}, isValid=${isValid}`);

                return {
                    label: 'PROFIT COST',
                    labelCell: `A${i + 1}`,
                    valueCell: `H${i + 1}`,
                    actualValue: displayValue,
                    expectedValue: '0.35 - 0.45',
                    isValid: isValid
                };
            }
        }

        return null;
    }

    /**
     * Format field value with color coding and expected value display
     */
    formatFieldValue(result) {
        // Special handling for wastage percent results
        if (result.label === 'Wastage % (FABRIC/UPPER/SHELL)') {
            return this.formatWastageValue(result);
        }

        // Special handling for sewing thread results
        if (result.label === 'Sewing Thread') {
            return this.formatSewingThreadValue(result);
        }

        // Special handling for standard packaging results
        if (result.label === 'Standard Packaging') {
            return this.formatStandardPackagingValue(result);
        }

        // Special handling for labor cost results
        if (result.label === 'Labor Cost') {
            return this.formatLaborCostValue(result);
        }

        if (result.actualValue === '' || result.actualValue === null || result.actualValue === 'Empty' || result.actualValue === 'Not found' || result.actualValue === 'Section not found') {
            return `<span style="color: #991b1b; font-weight: 600;">${result.actualValue || 'Empty'}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${result.expectedValue}</span>`;
        }

        if (result.isValid) {
            return `<span style="color: #065f46; font-weight: 600;">${result.actualValue}</span>`;
        } else {
            return `<span style="color: #991b1b; font-weight: 600;">${result.actualValue}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${result.expectedValue}</span>`;
        }
    }

    /**
     * Format Labor Cost validation result
     */
    formatLaborCostValue(result) {
        // Not found
        if (result.notFound) {
            return `<span style="color: #991b1b; font-weight: 600;">LABOR COST not found</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${result.expectedValue}</span>`;
        }

        let html = '<div>';

        // Show found items in green
        if (result.foundItems && result.foundItems.length > 0) {
            for (const item of result.foundItems) {
                html += `<span style="color: #065f46; font-weight: 600;">${item.cell}: ${item.value}</span><br>`;
            }
        }

        // Show missing items in red
        if (result.missingItems && result.missingItems.length > 0) {
            for (const item of result.missingItems) {
                html += `<span style="color: #991b1b; font-weight: 600;">Missing: ${item}</span><br>`;
            }
        }

        html += '</div>';

        return html;
    }

    /**
     * Format Sewing Thread validation result with cell locations and values, grouped by row
     */
    formatSewingThreadValue(result) {
        // Not found
        if (result.notFound) {
            return `<span style="color: #991b1b; font-weight: 600;">Not found in section</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${result.expectedValue}</span>`;
        }

        // Group fields by row number
        const allFields = [...(result.validFields || []), ...(result.invalidFields || [])];
        const rowGroups = {};

        for (const field of allFields) {
            // Extract row number from cell (e.g., "D18" -> "18")
            const rowNum = field.cell.replace(/[A-Z]/g, '');
            if (!rowGroups[rowNum]) {
                rowGroups[rowNum] = [];
            }
            rowGroups[rowNum].push(field);
        }

        let html = '<div>';

        // Sort row numbers and display grouped by row
        const sortedRows = Object.keys(rowGroups).sort((a, b) => parseInt(a) - parseInt(b));

        for (const rowNum of sortedRows) {
            html += `<div style="margin-bottom: 6px;"><strong style="color: #2b4a6c;">Row ${rowNum}:</strong> `;

            const fields = rowGroups[rowNum];
            const fieldParts = [];

            for (const field of fields) {
                const isValid = result.validFields?.some(f => f.cell === field.cell);
                const colLetter = field.cell.replace(/[0-9]/g, '');

                if (isValid) {
                    fieldParts.push(`<span style="color: #065f46;">${colLetter}=${field.value}</span>`);
                } else {
                    fieldParts.push(`<span style="color: #991b1b;">${colLetter}=${field.value}</span> <span style="font-size: 0.85em; color: #849bba;">(Expected: ${field.expected})</span>`);
                }
            }

            html += fieldParts.join(', ') + '</div>';
        }

        html += '</div>';

        return html;
    }

    /**
     * Format Standard Packaging validation result with cell locations and values, grouped by row
     */
    formatStandardPackagingValue(result) {
        // Not found
        if (result.notFound) {
            return `<span style="color: #991b1b; font-weight: 600;">Not found</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${result.expectedValue}</span>`;
        }

        // Group fields by row number
        const allFields = [...(result.validFields || []), ...(result.invalidFields || [])];
        const rowGroups = {};

        for (const field of allFields) {
            // Extract row number from cell (e.g., "D56" -> "56")
            const rowNum = field.cell.replace(/[A-Z]/g, '');
            if (!rowGroups[rowNum]) {
                rowGroups[rowNum] = [];
            }
            rowGroups[rowNum].push(field);
        }

        let html = '<div>';

        // Sort row numbers and display grouped by row
        const sortedRows = Object.keys(rowGroups).sort((a, b) => parseInt(a) - parseInt(b));

        for (const rowNum of sortedRows) {
            html += `<div style="margin-bottom: 6px;"><strong style="color: #2b4a6c;">Row ${rowNum}:</strong> `;

            const fields = rowGroups[rowNum];
            const fieldParts = [];

            for (const field of fields) {
                const isValid = result.validFields?.some(f => f.cell === field.cell);
                const colLetter = field.cell.replace(/[0-9]/g, '');

                if (isValid) {
                    fieldParts.push(`<span style="color: #065f46;">${colLetter}=${field.value}</span>`);
                } else {
                    fieldParts.push(`<span style="color: #991b1b;">${colLetter}=${field.value}</span> <span style="font-size: 0.85em; color: #849bba;">(Expected: ${field.expected})</span>`);
                }
            }

            html += fieldParts.join(', ') + '</div>';
        }

        html += '</div>';

        return html;
    }

    /**
     * Format wastage percent validation result with cell locations
     */
    formatWastageValue(result) {
        // Section not found
        if (result.labelCell === '-') {
            return `<span style="color: #991b1b; font-weight: 600;">Section not found</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${result.expectedValue}</span>`;
        }

        // No data rows found
        const totalRows = (result.validRows?.length || 0) + (result.invalidRows?.length || 0);
        if (totalRows === 0) {
            return `<span style="color: #991b1b; font-weight: 600;">No data rows found</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${result.expectedValue}</span>`;
        }

        let html = '<div>';

        // Show valid cells in green
        if (result.validRows && result.validRows.length > 0) {
            const validCells = result.validRows.map(r => r.cell).join(', ');
            html += `<span style="color: #065f46; font-weight: 600;">${validCells}</span>`;
        }

        // Show invalid cells in red with their actual values
        if (result.invalidRows && result.invalidRows.length > 0) {
            if (result.validRows && result.validRows.length > 0) {
                html += '<br>';
            }
            for (const row of result.invalidRows) {
                html += `<span style="color: #991b1b; font-weight: 600;">${row.cell}: ${row.value}</span> <span style="font-size: 0.85em; color: #849bba;">(Expected: 5%)</span><br>`;
            }
        }

        html += '</div>';

        return html;
    }

    /**
     * Generate HTML for results display
     */
    generateResultsHTML(results) {
        if (!results || results.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">FOX Validation Ready</p>
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
                        oninput="window.foxProcessor.searchByFilename(this.value)"
                    />
                </div>
                <button onclick="window.foxProcessor.exportToPDF()" class="export-btn">
                    Export
                </button>
            </div>
        `;

        for (const fileResult of results) {
            html += `<div class="file-result-group">`;

            const allResults = [
                fileResult.results.vendor,
                fileResult.results.factory,
                fileResult.results.coo,
                fileResult.results.overhead,
                fileResult.results.profitOthers,
                fileResult.results.wastagePercent,
                fileResult.results.sewingThread,
                fileResult.results.standardPackaging,
                fileResult.results.laborCost,
                fileResult.results.overheadCost,
                fileResult.results.profitCost
            ];

            const validCount = allResults.filter(r => r.isValid).length;
            const totalCount = allResults.length;

            html += `
                <div class="file-summary-box">
                    <strong>File:</strong> ${fileResult.fileName}<br>
                    <strong>Validation:</strong> ${validCount} out of ${totalCount} passed
                </div>
            `;

            html += `
                <table id="v20ResultsTable" class="results-table">
                    <thead>
                        <tr class="header-labels-row">
                            <th>Field</th>
                            <th>Cell</th>
                            <th>Value</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            for (const item of allResults) {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">${item.label}</td>
                        <td style="padding: 0.875rem 1rem;">${item.valueCell}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(item)}</td>
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

        const config = window.pdfExporter.createFOXConfig(this.bcbdResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }

    /**
     * Search by filename
     */
    searchByFilename(searchTerm) {
        const fileGroups = document.querySelectorAll('#tab-v20 .file-result-group');

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
window.foxProcessor = new FOXProcessor();

// Auto-initialize when V20 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    const v20Tabs = document.querySelectorAll('[data-tab="v20"]');
    v20Tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            window.foxProcessor.initialize();
        });
    });

    const v20TabContent = document.getElementById('tab-v20');
    if (v20TabContent && v20TabContent.classList.contains('active')) {
        window.foxProcessor.initialize();
    }
});
