/**
 * Mammut Processing Logic (V7)
 * Validates Buyer CBD files against Mammut criteria
 */

class MammutProcessor {
    constructor() {
        this.bcbdFiles = [];
        this.bcbdResults = [];
        this.validationRules = {
            // Column B -> Column C checks
            cellChecks: [
                { labelCol: 'B', labelValue: 'SUPPLIER', valueCol: 'C', expectedValue: 'Madison 88' },
                { labelCol: 'B', labelValue: 'CURRENCY', valueCol: 'C', expectedValue: 'USD' },
                { labelCol: 'B', labelValue: 'TARGET SUC', valueCol: 'C', expectedValue: 'NA' }
            ],
            // Profit Margin check: Column N for label, Column T for value (index 19)
            profitMargin: {
                labelCol: 'N',
                labelValue: 'PROFIT MARGIN:',
                valueColIndex: 19, // Column T (index 19)
                minValue: 0.30,
                maxValue: 0.50
            },
            // Wastage Cost check: Column Q (index 16)
            // Fabric section = 5%, all other sections = 3%
            wastageCost: {
                valueColIndex: 16, // Column Q (index 16)
                sections: [
                    { label: 'FABRIC TOTAL', expectedValue: 0.05 },      // 5% for fabric
                    { label: 'ZIPPER TOTAL', expectedValue: 0.03 },      // 3% for zipper
                    { label: 'TRIMS TOTAL', expectedValue: 0.03 },       // 3% for trims
                    { label: 'GRAPHIC TOTAL', expectedValue: 0.03 },     // 3% for graphic
                    { label: 'PACKING TOTAL', expectedValue: 0.03 },     // 3% for packing
                    { label: 'OTHERS TOTAL', expectedValue: 0.03 }       // 3% for others
                ]
            },
            // CMT (Cut, Make, Trim) checks - found after OTHERS TOTAL
            // Column E: label, Column H: price, Column K: exchange rate (1.00), Column L: currency (USD)
            cmtChecks: {
                labelColIndex: 4,    // Column E (index 4)
                priceColIndex: 7,    // Column H (index 7)
                exRateColIndex: 10,  // Column K (index 10)
                currencyColIndex: 11, // Column L (index 11)
                expectedExRate: 1.00,
                expectedCurrency: 'USD',
                items: [
                    { label: 'KNITTING', expectedPrice: 0.05 },
                    { label: 'LABELLING', expectedPrice: 0.05 },
                    { label: 'SEWING', expectedPrice: 0.20 },
                    { label: 'WASHING', expectedPrice: 0.17 },
                    { label: 'NEATEN/STEAMING/PACKING', expectedPrice: 0.40 }
                ]
            }
        };
    }

    /**
     * Initialize V7 - Display validation rules in the OB drop zone
     */
    initialize() {
        this.displayValidationRules();
        console.log('Mammut Processor initialized');
    }

    /**
     * Display validation rules in the OB drop zone (Burton-style)
     */
    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v7');
        if (!obDropZone) return;

        const pm = this.validationRules.profitMargin;
        const wc = this.validationRules.wastageCost;

        let html = `
            <div class="burton-cost-container">
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Cell Value Checks (Column B -> C)</strong></div>
                        ${this.validationRules.cellChecks.map(check =>
                            `<div class="burton-item-line"><strong>${check.labelValue}:</strong> ${check.expectedValue}</div>`
                        ).join('')}
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Profit Margin Check (Column N -> T)</strong></div>
                        <div class="burton-item-line"><strong>Label:</strong> ${pm.labelValue}</div>
                        <div class="burton-item-line"><strong>Valid Range:</strong> ${pm.minValue.toFixed(2)} - ${pm.maxValue.toFixed(2)}</div>
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Wastage Check (Column Q)</strong></div>
                        ${wc.sections.map(section =>
                            `<div class="burton-item-line"><strong>${section.label}:</strong> ${(section.expectedValue * 100).toFixed(0)}%</div>`
                        ).join('')}
                    </div>
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>CMT Checks (After OTHERS TOTAL)</strong></div>
                        <div class="burton-item-line"><strong>Exchange Rate (Col K):</strong> ${this.validationRules.cmtChecks.expectedExRate.toFixed(2)}</div>
                        <div class="burton-item-line"><strong>Currency (Col L):</strong> ${this.validationRules.cmtChecks.expectedCurrency}</div>
                        ${this.validationRules.cmtChecks.items.map(item =>
                            `<div class="burton-item-line"><strong>${item.label} (Col H):</strong> ${item.expectedPrice.toFixed(2)}</div>`
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
                    const cellChecks = this.checkCellValues(jsonData);
                    const profitMarginCheck = this.checkProfitMargin(jsonData);
                    const wastageCostCheck = this.checkWastageCost(jsonData);
                    const cmtCheck = this.checkCMT(jsonData);

                    resolve({
                        cellChecks: cellChecks,
                        profitMarginCheck: profitMarginCheck,
                        wastageCostCheck: wastageCostCheck,
                        cmtCheck: cmtCheck
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
     * Check cell values (Supplier, Currency, Target SUC)
     */
    checkCellValues(jsonData) {
        const results = [];

        for (const check of this.validationRules.cellChecks) {
            let found = false;
            let actualValue = '';
            let rowNumber = -1;

            // Search for the label in column B (index 1)
            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                const colB = row[1] ? String(row[1]).trim().toUpperCase() : '';

                if (colB === check.labelValue.toUpperCase() || colB.includes(check.labelValue.toUpperCase())) {
                    found = true;
                    rowNumber = i + 1;
                    // Get value from column C (index 2)
                    actualValue = row[2] ? String(row[2]).trim() : '';
                    console.log(`Found "${check.labelValue}" at row ${rowNumber}, Column C value: "${actualValue}"`);
                    break;
                }
            }

            const isValid = found && this.compareField(check.expectedValue, actualValue);

            results.push({
                label: check.labelValue,
                expectedValue: check.expectedValue,
                actualValue: actualValue,
                rowNumber: rowNumber,
                found: found,
                isValid: isValid
            });
        }

        return results;
    }

    /**
     * Check Profit Margin value
     * Looks for "PROFIT MARGIN:" in column N and checks the value in column T (same row)
     */
    checkProfitMargin(jsonData) {
        const pm = this.validationRules.profitMargin;
        let found = false;
        let actualValue = null;
        let rowNumber = -1;

        // Search for PROFIT MARGIN in column N (index 13)
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const colN = row[13] ? String(row[13]).trim().toUpperCase() : '';

            if (colN.includes('PROFIT MARGIN') || colN.includes('PROFIT MARGIN:')) {
                found = true;
                rowNumber = i + 1;
                // Get value from column T (index 19)
                actualValue = row[pm.valueColIndex];
                console.log(`Found "PROFIT MARGIN" at row ${rowNumber}, Value (col T): ${actualValue}`);
                break;
            }
        }

        if (!found) {
            return {
                found: false,
                message: 'PROFIT MARGIN not found in column N'
            };
        }

        // Parse the actual value
        let numericValue = null;
        if (actualValue !== null && actualValue !== undefined && actualValue !== '') {
            // Handle percentage format (e.g., "10%" or 0.10)
            const cleanValue = String(actualValue).replace(/[%,\s]/g, '');
            numericValue = parseFloat(cleanValue);
        }

        const isValid = numericValue !== null &&
                        numericValue >= pm.minValue &&
                        numericValue <= pm.maxValue;

        return {
            found: true,
            rowNumber: rowNumber,
            actualValue: actualValue,
            numericValue: numericValue,
            minValue: pm.minValue,
            maxValue: pm.maxValue,
            isValid: isValid
        };
    }

    /**
     * Check Wastage Cost values
     * Checks column Q for each section:
     * - Fabric section: 5%
     * - Zipper, Trims, Graphic, Packing, Others sections: 3%
     */
    checkWastageCost(jsonData) {
        const wc = this.validationRules.wastageCost;
        const colIndex = wc.valueColIndex; // Column Q (index 16)
        const sections = wc.sections;

        // Find all section TOTAL row indices
        const sectionRows = [];
        for (const section of sections) {
            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                const colB = row[1] ? String(row[1]).trim().toUpperCase() : '';

                if (colB.includes(section.label.toUpperCase())) {
                    sectionRows.push({
                        label: section.label,
                        expectedValue: section.expectedValue,
                        rowIndex: i
                    });
                    console.log(`Found "${section.label}" at row ${i + 1}`);
                    break;
                }
            }
        }

        if (sectionRows.length === 0) {
            return {
                found: false,
                message: 'No section TOTAL rows found in column B'
            };
        }

        // Sort sections by row index to process in order
        sectionRows.sort((a, b) => a.rowIndex - b.rowIndex);

        // Process each section
        const sectionResults = [];
        let startRow = 0;

        for (const section of sectionRows) {
            const endRow = section.rowIndex; // exclusive
            const expectedValue = section.expectedValue;
            let invalidCells = [];
            let validCells = [];

            // Check all rows in this section
            for (let i = startRow; i < endRow; i++) {
                const row = jsonData[i];
                const cellValue = row[colIndex];

                // Skip empty cells
                if (cellValue === null || cellValue === undefined || cellValue === '') {
                    continue;
                }

                // Parse the value
                let numericValue = null;
                const cleanValue = String(cellValue).replace(/[%,\s]/g, '');
                numericValue = parseFloat(cleanValue);

                if (isNaN(numericValue)) {
                    continue;
                }

                // Round to 2 decimal places for comparison
                const roundedValue = Math.round(numericValue * 100) / 100;
                const roundedExpected = Math.round(expectedValue * 100) / 100;

                if (roundedValue === roundedExpected) {
                    validCells.push({
                        rowNumber: i + 1,
                        cellAddress: `Q${i + 1}`,
                        value: cellValue,
                        numericValue: numericValue
                    });
                } else {
                    invalidCells.push({
                        rowNumber: i + 1,
                        cellAddress: `Q${i + 1}`,
                        value: cellValue,
                        numericValue: numericValue,
                        expectedValue: expectedValue
                    });
                }
            }

            sectionResults.push({
                label: section.label,
                totalRowNumber: section.rowIndex + 1,
                expectedValue: expectedValue,
                validCells: validCells,
                invalidCells: invalidCells,
                isValid: invalidCells.length === 0
            });

            // Next section starts after this TOTAL row
            startRow = section.rowIndex + 1;
        }

        const allValid = sectionResults.every(s => s.isValid);
        const totalValid = sectionResults.reduce((sum, s) => sum + s.validCells.length, 0);
        const totalInvalid = sectionResults.reduce((sum, s) => sum + s.invalidCells.length, 0);

        return {
            found: true,
            sections: sectionResults,
            isValid: allValid,
            summary: `${totalValid} valid, ${totalInvalid} invalid`
        };
    }

    /**
     * Check CMT (Cut, Make, Trim) items after OTHERS TOTAL
     * Validates: Column E (label), Column H (price), Column K (ex rate = 1.00), Column L (currency = USD)
     */
    checkCMT(jsonData) {
        const cmt = this.validationRules.cmtChecks;
        const labelColIndex = cmt.labelColIndex;       // Column E (index 4)
        const priceColIndex = cmt.priceColIndex;       // Column H (index 7)
        const exRateColIndex = cmt.exRateColIndex;     // Column K (index 10)
        const currencyColIndex = cmt.currencyColIndex; // Column L (index 11)

        // Find OTHERS TOTAL row first
        let othersTotalRowIndex = -1;
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const colB = row[1] ? String(row[1]).trim().toUpperCase() : '';
            if (colB.includes('OTHERS TOTAL')) {
                othersTotalRowIndex = i;
                console.log(`Found "OTHERS TOTAL" at row ${i + 1} for CMT check`);
                break;
            }
        }

        if (othersTotalRowIndex === -1) {
            return {
                found: false,
                message: 'OTHERS TOTAL row not found - cannot locate CMT section'
            };
        }

        const results = [];

        // Search for each CMT item after OTHERS TOTAL
        for (const item of cmt.items) {
            let found = false;
            let rowNumber = -1;
            let actualPrice = null;
            let actualExRate = null;
            let actualCurrency = null;

            // Search in rows after OTHERS TOTAL
            for (let i = othersTotalRowIndex + 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                const colE = row[labelColIndex] ? String(row[labelColIndex]).trim().toUpperCase() : '';

                if (colE.includes(item.label.toUpperCase())) {
                    found = true;
                    rowNumber = i + 1;
                    actualPrice = row[priceColIndex];
                    actualExRate = row[exRateColIndex];
                    actualCurrency = row[currencyColIndex] ? String(row[currencyColIndex]).trim().toUpperCase() : '';
                    console.log(`Found CMT item "${item.label}" at row ${rowNumber}`);
                    break;
                }
            }

            // Parse numeric values
            let numericPrice = null;
            let numericExRate = null;

            if (actualPrice !== null && actualPrice !== undefined && actualPrice !== '') {
                const cleanPrice = String(actualPrice).replace(/[$,\s]/g, '');
                numericPrice = parseFloat(cleanPrice);
            }

            if (actualExRate !== null && actualExRate !== undefined && actualExRate !== '') {
                const cleanExRate = String(actualExRate).replace(/[$,\s]/g, '');
                numericExRate = parseFloat(cleanExRate);
            }

            // Validate
            const priceValid = numericPrice !== null && Math.abs(numericPrice - item.expectedPrice) < 0.001;
            const exRateValid = numericExRate !== null && Math.abs(numericExRate - cmt.expectedExRate) < 0.001;
            const currencyValid = actualCurrency === cmt.expectedCurrency;

            results.push({
                label: item.label,
                found: found,
                rowNumber: rowNumber,
                expectedPrice: item.expectedPrice,
                actualPrice: actualPrice,
                numericPrice: numericPrice,
                priceValid: priceValid,
                expectedExRate: cmt.expectedExRate,
                actualExRate: actualExRate,
                numericExRate: numericExRate,
                exRateValid: exRateValid,
                expectedCurrency: cmt.expectedCurrency,
                actualCurrency: actualCurrency,
                currencyValid: currencyValid,
                isValid: found && priceValid && exRateValid && currencyValid
            });
        }

        const allValid = results.every(r => r.isValid);

        return {
            found: true,
            items: results,
            isValid: allValid
        };
    }

    /**
     * Compare text fields (case-insensitive)
     */
    compareField(expected, actual) {
        const exp = String(expected).toLowerCase().trim();
        const act = String(actual).toLowerCase().trim();
        return exp === act;
    }

    /**
     * Format number to 2 decimal places
     */
    formatNumber(value) {
        const num = parseFloat(String(value).replace(/[$,\s%]/g, ''));
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
                    <p style="font-size: 1.3em; margin-bottom: 10px;">Mammut Validation Ready</p>
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
                        oninput="window.mammutProcessor.searchByFilename(this.value)"
                    />
                </div>
                <button onclick="window.mammutProcessor.exportToPDF()" class="export-btn">
                    Export
                </button>
            </div>
        `;

        for (const fileResult of results) {
            const cellChecks = fileResult.results.cellChecks;
            const profitMargin = fileResult.results.profitMarginCheck;
            const wastageCost = fileResult.results.wastageCostCheck;
            const cmtCheck = fileResult.results.cmtCheck;

            // Count valid checks
            let validCount = 0;
            const wastageSectionCount = wastageCost.found && wastageCost.sections ? wastageCost.sections.length : 1;
            const cmtItemCount = cmtCheck && cmtCheck.found && cmtCheck.items ? cmtCheck.items.length : 0;
            let totalChecks = cellChecks.length + 1 + wastageSectionCount + cmtItemCount; // cell checks + profit margin + wastage sections + CMT items

            cellChecks.forEach(check => {
                if (check.isValid) validCount++;
            });
            if (profitMargin.found && profitMargin.isValid) validCount++;
            if (wastageCost.found && wastageCost.sections) {
                wastageCost.sections.forEach(section => {
                    if (section.isValid) validCount++;
                });
            }
            if (cmtCheck && cmtCheck.found && cmtCheck.items) {
                cmtCheck.items.forEach(item => {
                    if (item.isValid) validCount++;
                });
            }

            html += `<div class="file-result-group">`;

            // File summary
            html += `
                <div class="file-summary-box">
                    <strong>File:</strong> ${fileResult.fileName}<br>
                    <strong>Summary:</strong> ${validCount} out of ${totalChecks} checks passed
                </div>
            `;

            // Validation results table - simplified without Status column
            html += `
                <table class="results-table" style="table-layout: fixed; width: 100%;">
                    <thead>
                        <tr class="header-labels-row">
                            <th style="width: 50%;">Validation Check</th>
                            <th style="width: 50%;">Value</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            // Cell checks rows
            for (const check of cellChecks) {
                const valueColor = check.isValid ? '#065f46' : '#991b1b';
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">${check.label}</td>
                        <td style="padding: 0.875rem 1rem; text-align: center;">
                            ${check.found
                                ? `<span style="color: ${valueColor}; font-weight: 600;">${check.actualValue || 'Empty'}</span>`
                                : '<span style="color: #991b1b; font-weight: 600;">Not found</span>'}
                        </td>
                    </tr>
                `;
            }

            // Profit Margin row
            if (profitMargin.found) {
                const pmActual = profitMargin.numericValue !== null ? this.formatNumber(profitMargin.actualValue) : profitMargin.actualValue;
                const valueColor = profitMargin.isValid ? '#065f46' : '#991b1b';

                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">PROFIT MARGIN</td>
                        <td style="padding: 0.875rem 1rem; text-align: center;">
                            <span style="color: ${valueColor}; font-weight: 600;">${pmActual}</span>
                        </td>
                    </tr>
                `;
            } else {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">PROFIT MARGIN</td>
                        <td style="padding: 0.875rem 1rem; text-align: center; color: #991b1b; font-weight: 600;">${profitMargin.message || 'Not found'}</td>
                    </tr>
                `;
            }

            // Wastage rows for each section
            if (wastageCost.found && wastageCost.sections) {
                for (const section of wastageCost.sections) {
                    const sectionName = section.label.replace(' TOTAL', '');
                    const expectedPercent = (section.expectedValue * 100).toFixed(0);

                    // Combine valid and invalid cells on one row
                    const allCells = [];

                    // Add valid cells (green)
                    section.validCells.forEach(cell => {
                        allCells.push(`<span style="color: #065f46; font-weight: 600;">${cell.cellAddress}</span>`);
                    });

                    // Add invalid cells (red with value)
                    section.invalidCells.forEach(cell => {
                        const roundedValue = cell.numericValue.toFixed(2);
                        allCells.push(`<span style="color: #991b1b; font-weight: 600;">${cell.cellAddress} (${roundedValue})</span>`);
                    });

                    if (allCells.length > 0) {
                        html += `
                            <tr style="border-bottom: 1px solid #e0e8f0;">
                                <td style="padding: 0.875rem 1rem; font-weight: 600;">${sectionName} Wastage (${expectedPercent}%)</td>
                                <td style="padding: 0.875rem 1rem; text-align: center;">
                                    ${allCells.join(', ')}
                                </td>
                            </tr>
                        `;
                    } else {
                        // No cells at all
                        html += `
                            <tr style="border-bottom: 1px solid #e0e8f0;">
                                <td style="padding: 0.875rem 1rem; font-weight: 600;">${sectionName} Wastage (${expectedPercent}%)</td>
                                <td style="padding: 0.875rem 1rem; text-align: center;">
                                    <span style="color: #6b7280; font-weight: 600;">No data</span>
                                </td>
                            </tr>
                        `;
                    }
                }
            } else {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">Wastage Check</td>
                        <td style="padding: 0.875rem 1rem; text-align: center; color: #991b1b; font-weight: 600;">${wastageCost.message || 'No sections found'}</td>
                    </tr>
                `;
            }

            // CMT (Cut, Make, Trim) rows
            if (cmtCheck && cmtCheck.found && cmtCheck.items) {
                for (const item of cmtCheck.items) {
                    if (!item.found) {
                        html += `
                            <tr style="border-bottom: 1px solid #e0e8f0;">
                                <td style="padding: 0.875rem 1rem; font-weight: 600;">${item.label}</td>
                                <td style="padding: 0.875rem 1rem; text-align: center; color: #991b1b; font-weight: 600;">Not found</td>
                            </tr>
                        `;
                    } else {
                        // Build details string showing what's valid/invalid
                        const details = [];

                        // Price check
                        const priceColor = item.priceValid ? '#065f46' : '#991b1b';
                        const priceDisplay = item.numericPrice !== null ? item.numericPrice.toFixed(2) : item.actualPrice;
                        details.push(`<span style="color: ${priceColor};">Price: ${priceDisplay}</span>`);

                        // Exchange rate check
                        const exRateColor = item.exRateValid ? '#065f46' : '#991b1b';
                        const exRateDisplay = item.numericExRate !== null ? item.numericExRate.toFixed(2) : item.actualExRate;
                        details.push(`<span style="color: ${exRateColor};">Ex.Rate: ${exRateDisplay}</span>`);

                        // Currency check
                        const currencyColor = item.currencyValid ? '#065f46' : '#991b1b';
                        details.push(`<span style="color: ${currencyColor};">Currency: ${item.actualCurrency || 'N/A'}</span>`);

                        html += `
                            <tr style="border-bottom: 1px solid #e0e8f0;">
                                <td style="padding: 0.875rem 1rem; font-weight: 600;">${item.label}</td>
                                <td style="padding: 0.875rem 1rem; text-align: center; font-weight: 600;">
                                    ${details.join(' | ')}
                                </td>
                            </tr>
                        `;
                    }
                }
            } else if (cmtCheck && !cmtCheck.found) {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">Process Check</td>
                        <td style="padding: 0.875rem 1rem; text-align: center; color: #991b1b; font-weight: 600;">${cmtCheck.message || 'Not found'}</td>
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
        const resultsContainer = document.getElementById('results-v7');
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

        const config = window.pdfExporter.createMammutConfig(this.bcbdResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }
}

// Initialize the processor
window.mammutProcessor = new MammutProcessor();

// Initialize when V7 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    const v7Tab = document.querySelector('[data-tab="v7"]');
    if (v7Tab) {
        v7Tab.addEventListener('click', () => {
            window.mammutProcessor.initialize();
        });
    }

    // If V7 tab is already active on load, initialize immediately
    const v7TabContent = document.getElementById('tab-v7');
    if (v7TabContent && v7TabContent.classList.contains('active')) {
        window.mammutProcessor.initialize();
    }
});
