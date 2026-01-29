/**
 * Cotopaxi Cost Breakdown Processor
 * Validates BCBD files for Cotopaxi brand requirements
 *
 * Validation Logic:
 * 1. Find sheet named "Blank Cost Sheet"
 * 2. Scan column D for "VENDOR / COO" → Check column E same row for "PT UWU JUMP INDONESIA" or "HEADS UP"
 * 3. Scan column D for "SUPPLIER CONTACT" → Check column E same row for "Madison 88"
 * 4. Scan column E for "Overhead/Margin/Profit %:" → Check column G same row for 15% - 20%
 * 5. Fabric Section (between FABRIC header and "Total Fabric Yardage"):
 *    - If VENDOR/COO = PT UWU JUMP INDONESIA:
 *      - Yarn + M88/Local in Col B → Col I = 0.15%
 *      - Yarn + NOT M88/Local in Col B → Col I = 0.5%
 *      - Fabric Freight → Col I = 0.4%
 *    - If VENDOR/COO = HEADS UP:
 *      - Fabric items → Col I = 5%
 * 6. Trims Section (between TRIMS header and "Total Trims Cost" in Col G):
 *    - If VENDOR/COO = PT UWU JUMP INDONESIA:
 *      - Col B = Local/Freight → Col I = 0.012%
 *      - Col B = Other → Col I = 0.015%
 *    - If VENDOR/COO = HEADS UP:
 *      - Trims items → Col I = 3%
 * 7. General Packaging (scan Col D for "General Packaging"):
 *    - Col F = 1
 *    - If VENDOR/COO = PT UWU JUMP INDONESIA → Col I = 0.01%
 *    - If VENDOR/COO = HEADS UP → Col I = 3%
 */

class CotopaxiProcessor {
    constructor() {
        this.validationRules = [
            {
                name: 'Vendor / COO',
                markerColumn: 3, // Column D (0-indexed)
                marker: 'vendor / coo',
                checkColumn: 4, // Column E
                expected: ['PT UWU JUMP INDONESIA', 'HEADS UP'],
                type: 'multiple'
            },
            {
                name: 'Supplier Contact',
                markerColumn: 3, // Column D
                marker: 'supplier contact',
                checkColumn: 4, // Column E
                expected: 'Madison 88',
                type: 'exact'
            },
            {
                name: 'Overhead/Margin/Profit %',
                markerColumn: 4, // Column E (0-indexed)
                marker: 'overhead/margin/profit %',
                checkColumn: 6, // Column G
                expected: '15% - 20%',
                min: 0.15,
                max: 0.20,
                type: 'range'
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
                        <div class="burton-item-line">Column D = "VENDOR / COO" → Column E = <strong>PT UWU JUMP INDONESIA</strong> or <strong>HEADS UP</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Rule 2 - Supplier Contact:</strong></div>
                        <div class="burton-item-line">Column D = "SUPPLIER CONTACT" → Column E = <strong>Madison 88</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Rule 3 - Overhead/Margin/Profit %:</strong></div>
                        <div class="burton-item-line">Column E = "Overhead/Margin/Profit %:" → Column G = <strong>15% - 20%</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.75rem; border-top: 1px solid #ccc; padding-top: 0.5rem;"><strong>Fabric Section Rules (between FABRIC and Total Fabric Yardage):</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>If VENDOR / COO = PT UWU JUMP INDONESIA:</strong></div>
                        <div class="burton-item-line">• Yarn + (M88 or Local in Col B) → Col I = <strong>0.15%</strong></div>
                        <div class="burton-item-line">• Yarn + (NOT M88/Local in Col B) → Col I = <strong>0.5%</strong></div>
                        <div class="burton-item-line">• Fabric Freight → Col I = <strong>0.4%</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>If VENDOR / COO = HEADS UP:</strong></div>
                        <div class="burton-item-line">• Fabric items → Col I = <strong>5%</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.75rem; border-top: 1px solid #ccc; padding-top: 0.5rem;"><strong>Trims Section Rules (between TRIMS and Total Trims Cost):</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>If VENDOR / COO = PT UWU JUMP INDONESIA:</strong></div>
                        <div class="burton-item-line">• Col B = Local/Freight → Col I = <strong>0.012%</strong></div>
                        <div class="burton-item-line">• Col B = Other → Col I = <strong>0.015%</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>If VENDOR / COO = HEADS UP:</strong></div>
                        <div class="burton-item-line">• Trims items → Col I = <strong>3%</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.75rem; border-top: 1px solid #ccc; padding-top: 0.5rem;"><strong>General Packaging (scan Col D):</strong></div>
                        <div class="burton-item-line">• Column F = <strong>1</strong></div>
                        <div class="burton-item-line">• If PT UWU JUMP INDONESIA → Col I = <strong>0.01%</strong></div>
                        <div class="burton-item-line">• If HEADS UP → Col I = <strong>3%</strong></div>
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
                    console.log('Available sheets:', workbook.SheetNames);
                    for (const sheetName of workbook.SheetNames) {
                        if (sheetName.trim().toLowerCase() === 'blank cost sheet') {
                            targetSheetName = sheetName;
                            console.log('Found target sheet:', targetSheetName);
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

        // First, determine the VENDOR / COO value
        let vendorCOO = '';
        for (let rowIndex = 0; rowIndex < jsonData.length; rowIndex++) {
            const row = jsonData[rowIndex];
            if (!row) continue;
            const colD = row[3] ? String(row[3]).trim().toLowerCase() : '';
            if (colD.includes('vendor / coo')) {
                vendorCOO = row[4] ? String(row[4]).trim() : '';
                break;
            }
        }

        // Validate each rule by scanning for the marker
        for (const rule of this.validationRules) {
            const result = {
                name: rule.name,
                expected: Array.isArray(rule.expected) ? rule.expected.join(' or ') : rule.expected,
                found: false,
                rowNumber: -1,
                actual: '',
                isValid: false,
                markerColumn: this.getColumnLetter(rule.markerColumn),
                checkColumn: this.getColumnLetter(rule.checkColumn)
            };

            // Scan all rows to find the marker
            for (let rowIndex = 0; rowIndex < jsonData.length; rowIndex++) {
                const row = jsonData[rowIndex];
                if (!row) continue;

                const markerCell = row[rule.markerColumn] ? String(row[rule.markerColumn]).trim() : '';
                const markerCellLower = markerCell.toLowerCase();

                // Check if this row contains the marker
                if (markerCellLower === rule.marker.toLowerCase() || markerCellLower.includes(rule.marker.toLowerCase())) {
                    result.found = true;
                    result.rowNumber = rowIndex + 1; // Convert to 1-indexed

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
                    } else if (rule.type === 'range') {
                        // Parse percentage value (handle both "20%" and 0.20 formats)
                        let numValue = parseFloat(actualValue.replace('%', ''));
                        if (actualValue.includes('%')) {
                            numValue = numValue / 100; // Convert 20% to 0.20
                        }
                        // If value is greater than 1, assume it's already a percentage (e.g., 20 means 20%)
                        if (numValue > 1) {
                            numValue = numValue / 100;
                        }
                        result.isValid = !isNaN(numValue) && numValue >= rule.min && numValue <= rule.max;
                        // Display as percentage
                        if (!isNaN(numValue)) {
                            result.actual = (numValue * 100).toFixed(1) + '%';
                        }
                    }
                    break; // Found the marker, stop scanning
                }
            }

            if (!result.found) {
                result.actual = `Marker "${rule.marker}" not found in column ${result.markerColumn}`;
            }

            results.push(result);
        }

        // Validate Fabric section based on VENDOR / COO
        const fabricResults = this.validateFabricSection(jsonData, vendorCOO);
        results.push(...fabricResults);

        // Validate Trims section based on VENDOR / COO
        const trimsResults = this.validateTrimsSection(jsonData, vendorCOO);
        results.push(...trimsResults);

        // Validate General Packaging based on VENDOR / COO
        const packagingResults = this.validateGeneralPackaging(jsonData, vendorCOO);
        results.push(...packagingResults);

        return results;
    }

    /**
     * Validate the Fabric section based on VENDOR / COO value
     * Scans from "FABRIC" header row until "Total Fabric Yardage" in column D
     */
    validateFabricSection(jsonData, vendorCOO) {
        const results = [];
        const vendorLower = vendorCOO.toLowerCase();

        // Find the FABRIC header row (column A)
        let fabricStartRow = -1;
        let fabricEndRow = jsonData.length;

        for (let rowIndex = 0; rowIndex < jsonData.length; rowIndex++) {
            const row = jsonData[rowIndex];
            if (!row) continue;

            const colA = row[0] ? String(row[0]).trim().toLowerCase() : '';
            const colD = row[3] ? String(row[3]).trim().toLowerCase() : '';

            // Find FABRIC header
            if (fabricStartRow === -1 && colA === 'fabric') {
                fabricStartRow = rowIndex;
                continue;
            }

            // Find end marker "Total Fabric Yardage" in column D
            if (fabricStartRow !== -1 && colD.includes('total fabric yardage')) {
                fabricEndRow = rowIndex;
                break;
            }
        }

        if (fabricStartRow === -1) {
            results.push({
                name: 'Fabric Section',
                expected: 'FABRIC header',
                found: false,
                rowNumber: -1,
                actual: 'FABRIC header not found in column A',
                isValid: false,
                markerColumn: 'A',
                checkColumn: '-'
            });
            return results;
        }

        // Process based on VENDOR / COO
        if (vendorLower.includes('pt uwu jump indonesia')) {
            // PT UWU JUMP INDONESIA rules
            for (let rowIndex = fabricStartRow + 1; rowIndex < fabricEndRow; rowIndex++) {
                const row = jsonData[rowIndex];
                if (!row) continue;

                const colA = row[0] ? String(row[0]).trim().toLowerCase() : '';
                const colB = row[1] ? String(row[1]).trim().toLowerCase() : '';
                const colI = row[8]; // Column I (0-indexed as 8)

                // Rule: Yarn with M88 or Local in column B → Column I should be 0.15%
                if (colA === 'yarn') {
                    const supplierValue = row[1] ? String(row[1]).trim() : '';
                    if (colB.includes('m88') || colB.includes('local')) {
                        const result = this.createFabricResult(
                            'Yarn (M88/Local)',
                            '0.15%',
                            colI,
                            0.0015,
                            rowIndex + 1,
                            supplierValue
                        );
                        results.push(result);
                    } else {
                        // Yarn without M88 or Local → Column I should be 0.5%
                        const result = this.createFabricResult(
                            'Yarn (Non-M88/Local)',
                            '0.5%',
                            colI,
                            0.005,
                            rowIndex + 1,
                            supplierValue
                        );
                        results.push(result);
                    }
                }

                // Rule: Fabric Freight → Column I should be 0.4%
                if (colA === 'fabric freight') {
                    const supplierValue = row[1] ? String(row[1]).trim() : '';
                    const result = this.createFabricResult(
                        'Fabric Freight',
                        '0.4%',
                        colI,
                        0.004,
                        rowIndex + 1,
                        supplierValue
                    );
                    results.push(result);
                }
            }
        } else if (vendorLower.includes('heads up')) {
            // HEADS UP rules - scan for Fabric in column A → Column I should be 5%
            for (let rowIndex = fabricStartRow + 1; rowIndex < fabricEndRow; rowIndex++) {
                const row = jsonData[rowIndex];
                if (!row) continue;

                const colA = row[0] ? String(row[0]).trim().toLowerCase() : '';
                const colI = row[8]; // Column I (0-indexed as 8)

                // Any row with "Fabric" in column A (but not the header)
                if (colA === 'fabric' || (colA && colA !== 'fabric freight' && colA.includes('fabric'))) {
                    // Skip if it's just the header row
                    continue;
                }

                // For HEADS UP, check non-empty rows for the 5% rule
                if (colA && colA !== '' && colA !== 'fabric') {
                    const supplierValue = row[1] ? String(row[1]).trim() : '';
                    const result = this.createFabricResult(
                        `Fabric Item (${row[0]})`,
                        '5%',
                        colI,
                        0.05,
                        rowIndex + 1,
                        supplierValue
                    );
                    results.push(result);
                }
            }
        }

        return results;
    }

    /**
     * Validate the Trims section based on VENDOR / COO value
     * Scans from "TRIMS" header row in column A
     */
    validateTrimsSection(jsonData, vendorCOO) {
        const results = [];
        const vendorLower = vendorCOO.toLowerCase();

        // Find the TRIMS header row (column A)
        let trimsStartRow = -1;
        let trimsEndRow = jsonData.length;

        for (let rowIndex = 0; rowIndex < jsonData.length; rowIndex++) {
            const row = jsonData[rowIndex];
            if (!row) continue;

            const colA = row[0] ? String(row[0]).trim().toLowerCase() : '';
            const colG = row[6] ? String(row[6]).trim().toLowerCase() : ''; // Column G (0-indexed as 6)

            // Find TRIMS header
            if (trimsStartRow === -1 && colA === 'trims') {
                trimsStartRow = rowIndex;
                continue;
            }

            // End when we find "Total Trims Cost" in column G
            if (trimsStartRow !== -1 && colG.includes('total trims cost')) {
                trimsEndRow = rowIndex;
                break;
            }
        }

        if (trimsStartRow === -1) {
            results.push({
                name: 'Trims Section',
                expected: 'TRIMS header',
                found: false,
                rowNumber: -1,
                actual: 'TRIMS header not found in column A',
                isValid: false,
                markerColumn: 'A',
                checkColumn: '-',
                supplier: ''
            });
            return results;
        }

        // Process based on VENDOR / COO
        if (vendorLower.includes('pt uwu jump indonesia')) {
            // PT UWU JUMP INDONESIA rules
            for (let rowIndex = trimsStartRow + 1; rowIndex < trimsEndRow; rowIndex++) {
                const row = jsonData[rowIndex];
                if (!row) continue;

                const colA = row[0] ? String(row[0]).trim() : '';
                const colALower = colA.toLowerCase();
                const colB = row[1] ? String(row[1]).trim().toLowerCase() : '';
                const colI = row[8]; // Column I (0-indexed as 8)

                // Skip empty rows or header-like rows
                if (!colA || colALower === 'trims') continue;

                const supplierValue = row[1] ? String(row[1]).trim() : '';

                // Rule: If column B contains "Local" or "freight" → Column I should be 0.012%
                if (colB.includes('local') || colB.includes('freight')) {
                    const result = this.createFabricResult(
                        `Trims: ${colA} (Local/Freight)`,
                        '0.012%',
                        colI,
                        0.00012,
                        rowIndex + 1,
                        supplierValue
                    );
                    results.push(result);
                } else if (colB) {
                    // Column B has other value → Column I should be 0.015%
                    const result = this.createFabricResult(
                        `Trims: ${colA}`,
                        '0.015%',
                        colI,
                        0.00015,
                        rowIndex + 1,
                        supplierValue
                    );
                    results.push(result);
                }
            }
        } else if (vendorLower.includes('heads up')) {
            // HEADS UP rules - check column A items → Column I should be 3%
            for (let rowIndex = trimsStartRow + 1; rowIndex < trimsEndRow; rowIndex++) {
                const row = jsonData[rowIndex];
                if (!row) continue;

                const colA = row[0] ? String(row[0]).trim() : '';
                const colALower = colA.toLowerCase();
                const colI = row[8]; // Column I (0-indexed as 8)

                // Skip empty rows or header-like rows
                if (!colA || colALower === 'trims') continue;

                const supplierValue = row[1] ? String(row[1]).trim() : '';

                const result = this.createFabricResult(
                    `Trims: ${colA}`,
                    '3%',
                    colI,
                    0.03,
                    rowIndex + 1,
                    supplierValue
                );
                results.push(result);
            }
        }

        return results;
    }

    /**
     * Validate General Packaging based on VENDOR / COO value
     * Scans column D for "General Packaging", checks column F = 1, column I based on vendor
     */
    validateGeneralPackaging(jsonData, vendorCOO) {
        const results = [];
        const vendorLower = vendorCOO.toLowerCase();

        // Scan column D for "General Packaging"
        for (let rowIndex = 0; rowIndex < jsonData.length; rowIndex++) {
            const row = jsonData[rowIndex];
            if (!row) continue;

            const colD = row[3] ? String(row[3]).trim().toLowerCase() : '';

            if (colD.includes('general packaging')) {
                const colF = row[5]; // Column F (0-indexed as 5)
                const colI = row[8]; // Column I (0-indexed as 8)

                // Check Column F = 1
                const colFValue = colF !== undefined && colF !== null ? String(colF).trim() : '';
                const colFNum = parseFloat(colFValue);
                const colFIsValid = colFNum === 1;

                results.push({
                    name: 'General Packaging (Qty)',
                    expected: '1',
                    found: true,
                    rowNumber: rowIndex + 1,
                    actual: colFValue || 'Empty',
                    isValid: colFIsValid,
                    markerColumn: 'D',
                    checkColumn: 'F',
                    supplier: ''
                });

                // Check Column I based on VENDOR / COO
                let expectedPercent = '';
                let expectedDecimal = 0;

                if (vendorLower.includes('pt uwu jump indonesia')) {
                    expectedPercent = '0.01%';
                    expectedDecimal = 0.0001;
                } else if (vendorLower.includes('heads up')) {
                    expectedPercent = '3%';
                    expectedDecimal = 0.03;
                }

                if (expectedPercent) {
                    const result = this.createFabricResult(
                        'General Packaging (%)',
                        expectedPercent,
                        colI,
                        expectedDecimal,
                        rowIndex + 1,
                        ''
                    );
                    results.push(result);
                }

                break; // Found General Packaging, stop scanning
            }
        }

        if (results.length === 0) {
            results.push({
                name: 'General Packaging',
                expected: 'General Packaging in column D',
                found: false,
                rowNumber: -1,
                actual: 'General Packaging not found in column D',
                isValid: false,
                markerColumn: 'D',
                checkColumn: '-',
                supplier: ''
            });
        }

        return results;
    }

    /**
     * Helper to create a fabric validation result
     */
    createFabricResult(name, expectedStr, actualValue, expectedDecimal, rowNumber, supplier = '') {
        let numValue = 0;
        let actualStr = '';

        if (actualValue !== undefined && actualValue !== null && actualValue !== '') {
            numValue = parseFloat(String(actualValue).replace('%', ''));
            // Handle percentage string format
            if (String(actualValue).includes('%')) {
                numValue = numValue / 100;
            }
            // If value is greater than 1, assume it needs conversion (e.g., 0.15 means 0.15%)
            // Actually for small percentages like 0.15%, the raw value would be 0.0015
            // Use more decimal places for very small percentages
            const percentValue = numValue * 100;
            if (percentValue < 0.1) {
                actualStr = percentValue.toFixed(3) + '%'; // Show 3 decimals for very small values
            } else if (percentValue < 1) {
                actualStr = percentValue.toFixed(2) + '%'; // Show 2 decimals for small values
            } else {
                actualStr = percentValue.toFixed(1) + '%'; // Show 1 decimal for larger values
            }
        } else {
            actualStr = 'Empty';
        }

        // For comparison, allow small tolerance
        const isValid = Math.abs(numValue - expectedDecimal) < 0.0001;

        return {
            name: name,
            expected: expectedStr,
            found: true,
            rowNumber: rowNumber,
            actual: actualStr,
            isValid: isValid,
            markerColumn: 'A',
            checkColumn: 'I',
            supplier: supplier
        };
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

                // Create results table with 4 columns (Supplier column for fabric items)
                html += `
                    <table class="results-table" style="table-layout: fixed; width: 100%;">
                        <thead>
                            <tr class="header-labels-row">
                                <th style="width: 200px;">Check Name</th>
                                <th style="width: 150px;">Supplier (Col B)</th>
                                <th>Value</th>
                                <th style="width: 150px;">Expected</th>
                            </tr>
                        </thead>
                        <tbody>
                `;

                for (const check of fileResult.checks) {
                    let valueHTML = '';
                    let expectedHTML = '';
                    let supplierHTML = check.supplier !== undefined ? check.supplier : '-';

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
                            <td style="padding: 0.875rem 1rem;">${supplierHTML}</td>
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
