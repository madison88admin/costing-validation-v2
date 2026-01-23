/**
 * Prana Cost Breakdown Processor
 * Validates BCBD files - checks multiple sections for wastage in Column G
 *
 * Validation Logic:
 * 1. Check ALL sheets in the Excel file (unlike other templates)
 *
 * Section 1 - Fabrics:
 * - Find "Fabrics" in Column A
 * - Scan Column G until "Fabric Subtotal"
 * - Expected wastage: 5%
 *
 * Section 2 - Trims, Insulation:
 * - Find "Trims, Insulation" in Column A
 * - Scan Column G until "Trim, Fills Subtotal"
 * - Expected wastage: 3%
 *
 * Section 3 - Thread:
 * - Find "Thread" in Column A
 * - Scan Column G until "Thread Subtotal"
 * - Expected wastage: 3%
 * - Special: "Sewing Thread" row checks Column I (Total Yield) = 1.00, Column J (Unit Price) = 0.00949
 *
 * Section 4 - Labels / Garment Packaging:
 * - Find "LABELS / GARMENT PACKAGING" in Column A
 * - Scan Column G until "Labels/Garment Packaging Subtotal"
 * - Expected wastage: 3%
 *
 * Global Row Checks (all sheets):
 * - "Overhead" in Column A → Column K = 0.30049
 * - "Profit" in Column A → Column K ≤ 0.30
 * - "Transit/Transportation" in Column A → Column K = 0.200499
 * - "Finance" in Column A → Column J = 0, Column K = 0.15049999
 */

class PranaProcessor {
    constructor() {
        this.sections = [
            {
                name: 'Fabrics',
                startMarker: 'fabrics',
                endMarker: 'fabric subtotal',
                expectedWastage: '5%'
            },
            {
                name: 'Trims, Insulation',
                startMarker: 'trims, insulation',
                endMarker: 'trim, fills subtotal',
                expectedWastage: '3%'
            },
            {
                name: 'Thread',
                startMarker: 'thread',
                endMarker: 'thread subtotal',
                expectedWastage: '3%',
                specialItems: [
                    {
                        name: 'sewing thread',
                        checks: [
                            { column: 8, columnLetter: 'I', label: 'Total Yield', expected: '1.00' },
                            { column: 9, columnLetter: 'J', label: 'Unit Price', expected: '0.00949' }
                        ]
                    }
                ]
            },
            {
                name: 'Labels / Garment Packaging',
                startMarker: 'labels / garment packaging',
                endMarker: 'labels/garment packaging subtotal',
                expectedWastage: '3%',
                useColumnB: true, // Check Column B for item names instead of Column A
                specialItems: [
                    {
                        name: 'beanie packaging',
                        checks: [
                            { column: 8, columnLetter: 'I', label: 'Total Yield', expected: '1.00' },
                            { column: 9, columnLetter: 'J', label: 'Unit Price', expected: '0.055' }
                        ]
                    },
                    {
                        name: '25mm interior tear away label',
                        checks: [
                            { column: 8, columnLetter: 'I', label: 'Total Yield', expected: '1.00' },
                            { column: 9, columnLetter: 'J', label: 'Unit Price', expected: '0.05' }
                        ]
                    },
                    {
                        name: 'upc sticker small',
                        checks: [
                            { column: 8, columnLetter: 'I', label: 'Total Yield', expected: '1.00' },
                            { column: 9, columnLetter: 'J', label: 'Unit Price', expected: '0.008' }
                        ]
                    },
                    {
                        name: 'upc sticker small with msrp',
                        checks: [
                            { column: 8, columnLetter: 'I', label: 'Total Yield', expected: '1.00' },
                            { column: 9, columnLetter: 'J', label: 'Unit Price', expected: '0.054' }
                        ]
                    },
                    {
                        name: 'glassine tissue bag',
                        checks: [
                            { column: 8, columnLetter: 'I', label: 'Total Yield', expected: '1.00' },
                            { column: 9, columnLetter: 'J', label: 'Unit Price', expected: '0.066' }
                        ]
                    },
                    {
                        name: 'care label',
                        checks: [
                            { column: 8, columnLetter: 'I', label: 'Total Yield', expected: '1.00' },
                            { column: 9, columnLetter: 'J', label: 'Unit Price', expected: '0.03' }
                        ]
                    },
                    {
                        name: 'po label',
                        checks: [
                            { column: 8, columnLetter: 'I', label: 'Total Yield', expected: '1.00' },
                            { column: 9, columnLetter: 'J', label: 'Unit Price', expected: '0.01' }
                        ]
                    },
                    {
                        name: 'main tag upc sticker',
                        checks: [
                            { column: 8, columnLetter: 'I', label: 'Total Yield', expected: '1.00' },
                            { column: 9, columnLetter: 'J', label: 'Unit Price', expected: '0.006' }
                        ]
                    }
                ]
            }
        ];

        // Global row checks - scan Column A for keywords and validate values
        this.globalRowChecks = [
            {
                name: 'Overhead',
                marker: 'overhead',
                checks: [
                    { column: 10, columnLetter: 'K', label: 'Value', expected: '0.30049' }
                ]
            },
            {
                name: 'Profit',
                marker: 'profit',
                checks: [
                    { column: 10, columnLetter: 'K', label: 'Value', expected: '≤ 0.30', maxValue: 0.30 }
                ]
            },
            {
                name: 'Transit/Transportation',
                marker: 'transit/transportation',
                checks: [
                    { column: 10, columnLetter: 'K', label: 'Value', expected: '0.200499' }
                ]
            },
            {
                name: 'Finance',
                marker: 'finance',
                checks: [
                    { column: 9, columnLetter: 'J', label: 'Percentage', expected: '0', allowedValues: ['0', '0%'] },
                    { column: 10, columnLetter: 'K', label: 'Value', expected: '0.15049999' }
                ]
            }
        ];
    }

    async initialize() {
        this.displayValidationRules();
    }

    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v13');
        if (!obDropZone) return;

        let html = `
            <div class="burton-cost-container">
                <div class="burton-cost-items">
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>Validation Rules:</strong></div>
                        <div class="burton-item-line">Scans ALL sheets in the Excel file</div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Section 1 - Fabrics:</strong></div>
                        <div class="burton-item-line">• Column G Wastage: <strong>5%</strong></div>
                        <div class="burton-item-line">• Stops at "Fabric Subtotal"</div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Section 2 - Trims, Insulation:</strong></div>
                        <div class="burton-item-line">• Column G Wastage: <strong>3%</strong></div>
                        <div class="burton-item-line">• Stops at "Trim, Fills Subtotal"</div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Section 3 - Thread:</strong></div>
                        <div class="burton-item-line">• Column G Wastage: <strong>3%</strong></div>
                        <div class="burton-item-line">• Stops at "Thread Subtotal"</div>
                        <div class="burton-item-line">• <em>Sewing Thread:</em> Col I = <strong>1.00</strong>, Col J = <strong>0.00949</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Section 4 - Labels / Garment Packaging:</strong></div>
                        <div class="burton-item-line">• Column G Wastage: <strong>3%</strong></div>
                        <div class="burton-item-line">• Stops at "Labels/Garment Packaging Subtotal"</div>
                        <div class="burton-item-line">• <em>Special Items (Col B):</em> Col I = <strong>1.00</strong>, Col J = <strong>Unit Price</strong></div>
                        <div class="burton-item-line" style="margin-top: 0.5rem;"><strong>Global Checks (Column A):</strong></div>
                        <div class="burton-item-line">• Overhead: Col K = <strong>0.30049</strong></div>
                        <div class="burton-item-line">• Profit: Col K = <strong>≤ 0.30</strong></div>
                        <div class="burton-item-line">• Transit/Transportation: Col K = <strong>0.200499</strong></div>
                        <div class="burton-item-line">• Finance: Col J = <strong>0</strong>, Col K = <strong>0.15049999</strong></div>
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
                    sheets: []
                });
            }
        }

        return this.generateResultsHTML(results);
    }

    async processFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    // Process ALL sheets in the workbook
                    const sheetResults = [];

                    for (const sheetName of workbook.SheetNames) {
                        const sheet = workbook.Sheets[sheetName];
                        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                        const sheetValidation = this.validateSheet(jsonData, sheetName);
                        if (sheetValidation.anySectionFound) {
                            sheetResults.push(sheetValidation);
                        }
                    }

                    resolve({
                        fileName: file.name,
                        sheets: sheetResults,
                        totalSheets: workbook.SheetNames.length,
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

    validateSheet(jsonData, sheetName) {
        const result = {
            sheetName: sheetName,
            sections: [],
            globalChecks: [],
            anySectionFound: false
        };

        // Process each section type
        for (const sectionConfig of this.sections) {
            const sectionResult = this.validateSection(jsonData, sectionConfig);
            if (sectionResult.found) {
                result.sections.push(sectionResult);
                result.anySectionFound = true;
            }
        }

        // Process global row checks
        const globalResults = this.validateGlobalRowChecks(jsonData);
        if (globalResults.length > 0) {
            result.globalChecks = globalResults;
            result.anySectionFound = true;
        }

        return result;
    }

    validateGlobalRowChecks(jsonData) {
        const results = [];

        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const cellA = row[0] ? String(row[0]).trim() : '';
            const cellALower = cellA.toLowerCase();
            const rowNum = i + 1;

            // Check each global row check config
            for (const checkConfig of this.globalRowChecks) {
                if (cellALower === checkConfig.marker.toLowerCase()) {
                    const checkResult = {
                        name: checkConfig.name,
                        rowNumber: rowNum,
                        checks: []
                    };

                    for (const check of checkConfig.checks) {
                        const actualValue = row[check.column];
                        const result = this.validateGlobalValue(actualValue, check.expected, check.label, rowNum, check.columnLetter, check.allowedValues, check.maxValue);
                        checkResult.checks.push(result);
                    }

                    results.push(checkResult);
                }
            }
        }

        return results;
    }

    validateGlobalValue(actual, expected, label, rowNum, colLetter, allowedValues, maxValue) {
        let actualStr = '';
        let actualNum = NaN;

        if (actual !== undefined && actual !== null && actual !== '') {
            actualNum = parseFloat(actual);
            if (!isNaN(actualNum)) {
                // Format to match expected precision (but handle ≤ prefix)
                const expectedClean = expected.replace(/[≤<>=]/g, '').trim();
                const expectedDecimals = (expectedClean.split('.')[1] || '').length;
                actualStr = actualNum.toFixed(expectedDecimals);
            } else {
                actualStr = String(actual).trim();
            }
        } else {
            actualStr = 'Empty';
        }

        // Check if valid
        let isValid = false;

        // If maxValue is specified, check if actual is <= maxValue
        if (maxValue !== undefined) {
            isValid = !isNaN(actualNum) && actualNum <= maxValue;
        } else {
            // Otherwise check exact match or allowedValues
            isValid = actualStr === expected;
            if (!isValid && allowedValues) {
                isValid = allowedValues.includes(actualStr);
            }
        }

        return {
            label: label,
            expected: expected,
            actual: actualStr,
            isValid: isValid,
            cellAddress: `${colLetter}${rowNum}`
        };
    }

    validateSection(jsonData, sectionConfig) {
        const result = {
            name: sectionConfig.name,
            found: false,
            rowStart: -1,
            rowEnd: -1,
            expectedWastage: sectionConfig.expectedWastage,
            items: [],
            specialItemResults: [],
            allValid: true
        };

        let inSection = false;

        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            const cellA = row[0] ? String(row[0]).trim() : '';
            const cellALower = cellA.toLowerCase();

            // Check for section start marker
            if (!inSection && cellALower === sectionConfig.startMarker) {
                inSection = true;
                result.found = true;
                result.rowStart = i + 1; // 1-indexed for display
                continue;
            }

            // Check for section end marker
            if (inSection && cellALower === sectionConfig.endMarker) {
                result.rowEnd = i + 1; // 1-indexed for display
                break;
            }

            // If we're in the section, check Column G
            if (inSection) {
                const cellG = row[6]; // Column G (index 6)
                const rowNum = i + 1; // 1-indexed for display

                // Skip empty rows
                if (!cellA && !cellG) continue;

                const itemResult = this.validateWastage(cellG, rowNum, cellA, sectionConfig.expectedWastage);
                result.items.push(itemResult);

                if (!itemResult.isValid) {
                    result.allValid = false;
                }

                // Check for special items (like "sewing thread" in Thread section)
                if (sectionConfig.specialItems) {
                    // Use Column B if specified, otherwise Column A
                    const itemNameColumn = sectionConfig.useColumnB ? 1 : 0;
                    const cellForMatching = row[itemNameColumn] ? String(row[itemNameColumn]).trim() : '';
                    const cellForMatchingLower = cellForMatching.toLowerCase();

                    for (const specialItem of sectionConfig.specialItems) {
                        if (cellForMatchingLower === specialItem.name.toLowerCase()) {
                            // Get Column A value for display
                            const columnAValue = row[0] ? String(row[0]).trim() : '';

                            const specialResult = {
                                name: cellForMatching,
                                columnA: columnAValue,
                                rowNumber: rowNum,
                                checks: []
                            };

                            for (const check of specialItem.checks) {
                                const actualValue = row[check.column];
                                const checkResult = this.validateSpecialValue(actualValue, check.expected, check.label, rowNum, check.columnLetter);
                                specialResult.checks.push(checkResult);

                                if (!checkResult.isValid) {
                                    result.allValid = false;
                                }
                            }

                            result.specialItemResults.push(specialResult);
                        }
                    }
                }
            }
        }

        return result;
    }

    validateSpecialValue(actual, expected, label, rowNum, colLetter) {
        let actualStr = '';

        if (actual !== undefined && actual !== null && actual !== '') {
            const actualNum = parseFloat(actual);
            if (!isNaN(actualNum)) {
                // Format to match expected precision
                const expectedDecimals = (expected.split('.')[1] || '').length;
                actualStr = actualNum.toFixed(expectedDecimals);
            } else {
                actualStr = String(actual).trim();
            }
        } else {
            actualStr = 'Empty';
        }

        const isValid = actualStr === expected;

        return {
            label: label,
            expected: expected,
            actual: actualStr,
            isValid: isValid,
            cellAddress: `${colLetter}${rowNum}`
        };
    }

    validateWastage(actual, rowNum, itemName, expectedWastage) {
        // Handle the actual value
        let actualStr = '';

        if (actual !== undefined && actual !== null && actual !== '') {
            // Check if it's a number (Excel might store 5% as 0.05)
            const actualNum = parseFloat(actual);
            if (!isNaN(actualNum)) {
                if (actualNum < 1) {
                    // Convert decimal to percentage (0.05 -> 5%)
                    actualStr = (actualNum * 100).toFixed(0) + '%';
                } else {
                    actualStr = actualNum.toFixed(0) + '%';
                }
            } else {
                actualStr = String(actual).trim();
            }
        } else {
            actualStr = 'Empty';
        }

        // Normalize for comparison
        const normalizeValue = (val) => {
            if (!val || val === 'Empty') return '';
            // Remove % and whitespace, get number
            let normalized = String(val).replace(/%/g, '').trim();
            const num = parseFloat(normalized);
            if (!isNaN(num)) {
                return num.toString();
            }
            return normalized.toLowerCase();
        };

        const normalizedActual = normalizeValue(actualStr);
        const normalizedExpected = normalizeValue(expectedWastage);

        const isValid = normalizedActual === normalizedExpected;

        return {
            rowNumber: rowNum,
            itemName: itemName || `Row ${rowNum}`,
            expected: expectedWastage,
            actual: actualStr,
            isValid: isValid,
            cellAddress: `G${rowNum}`
        };
    }

    generateResultsHTML(results) {
        let html = '';

        for (const fileResult of results) {
            html += `<div class="file-result-group">`;

            if (fileResult.error) {
                html += `
                    <div class="file-summary-box">
                        <strong>File:</strong> ${fileResult.fileName}<br>
                        <span style="color: #991b1b;">Error: ${fileResult.error}</span>
                    </div>
                `;
            } else if (fileResult.sheets.length === 0) {
                html += `
                    <div class="file-summary-box">
                        <strong>File:</strong> ${fileResult.fileName}<br>
                        <strong>Sheets scanned:</strong> ${fileResult.totalSheets}<br>
                        <span style="color: #b45309;">No validation sections found in any sheet</span>
                    </div>
                `;
            } else {
                // Calculate overall summary across all sheets and sections
                let totalItems = 0;
                let validItems = 0;

                for (const sheet of fileResult.sheets) {
                    for (const section of sheet.sections) {
                        totalItems += section.items.length;
                        validItems += section.items.filter(i => i.isValid).length;
                    }
                }

                const sheetsWithSections = fileResult.sheets.length;

                html += `
                    <div class="file-summary-box">
                        <strong>File:</strong> ${fileResult.fileName}<br>
                        <strong>Sheets scanned:</strong> ${fileResult.totalSheets} | <strong>Sheets with sections:</strong> ${sheetsWithSections}<br>
                        <strong>Summary:</strong> ${validItems} out of ${totalItems} wastage values are correct
                    </div>
                `;

                // Show results for each sheet
                for (const sheetResult of fileResult.sheets) {
                    html += this.generateSheetTable(sheetResult);
                }
            }

            html += '</div>';
        }

        return html;
    }

    generateSheetTable(sheetResult) {
        let html = `
            <div style="margin-top: 1.5rem; margin-bottom: 1.5rem;">
            <table class="results-table" style="table-layout: fixed; width: 100%;">
                <thead>
                    <tr class="header-labels-row">
                        <th style="width: 200px;">Sheet Name</th>
                        <th style="width: 150px;">Section</th>
                        <th>Wastage (Column G)</th>
                    </tr>
                </thead>
                <tbody>
        `;

        // Add a row for each section found in this sheet
        for (let i = 0; i < sheetResult.sections.length; i++) {
            const section = sheetResult.sections[i];
            const validItems = section.items.filter(item => item.isValid);
            const invalidItems = section.items.filter(item => !item.isValid);

            html += `
                <tr style="border-bottom: 1px solid #e0e8f0;">
                    ${i === 0 ? `<td style="padding: 0.875rem 1rem; font-weight: 600; vertical-align: top; max-width: 200px; word-wrap: break-word;" rowspan="${sheetResult.sections.length}">${sheetResult.sheetName}</td>` : ''}
                    <td style="padding: 0.875rem 1rem; font-weight: 600;">${section.name}<br><span style="font-size: 0.85em; color: #64748b;">Expected: ${section.expectedWastage}</span></td>
                    <td style="padding: 0.875rem 1rem;">
                        ${this.formatWastageCells(validItems, invalidItems)}
                        ${this.formatSpecialItems(section.specialItemResults)}
                    </td>
                </tr>
            `;
        }

        // Add global checks if any
        if (sheetResult.globalChecks && sheetResult.globalChecks.length > 0) {
            const totalRows = sheetResult.sections.length + 1; // +1 for global checks row

            // Update rowspan for sheet name cell if sections exist
            html = html.replace(
                `rowspan="${sheetResult.sections.length}"`,
                `rowspan="${totalRows}"`
            );

            html += `
                <tr style="border-bottom: 1px solid #e0e8f0;">
                    ${sheetResult.sections.length === 0 ? `<td style="padding: 0.875rem 1rem; font-weight: 600; vertical-align: top; max-width: 200px; word-wrap: break-word;">${sheetResult.sheetName}</td>` : ''}
                    <td style="padding: 0.875rem 1rem; font-weight: 600;">Global Checks</td>
                    <td style="padding: 0.875rem 1rem;">
                        ${this.formatGlobalChecks(sheetResult.globalChecks)}
                    </td>
                </tr>
            `;
        }

        html += `
                </tbody>
            </table>
            </div>
        `;

        return html;
    }

    formatGlobalChecks(globalChecks) {
        let html = '';

        for (const check of globalChecks) {
            html += `<div style="margin-bottom: 0.5rem;">`;
            html += `<span style="font-weight: 600; color: #64748b; font-size: 0.85em;">${check.name} (Row ${check.rowNumber}):</span> `;

            const checkResults = check.checks.map(c => {
                if (c.isValid) {
                    return `<span style="color: #065f46; font-weight: 600;">${c.label}: ${c.actual}</span>`;
                } else {
                    return `<span style="color: #991b1b; font-weight: 600;">${c.label}: ${c.actual}</span> <span style="font-size: 0.85em; color: #849bba;">(Expected: ${c.expected})</span>`;
                }
            }).join(', ');
            html += checkResults;

            html += `</div>`;
        }

        return html;
    }

    formatSpecialItems(specialItemResults) {
        if (!specialItemResults || specialItemResults.length === 0) {
            return '';
        }

        let html = '';

        for (const specialItem of specialItemResults) {
            html += `<div style="margin-top: 0.5rem; padding-top: 0.5rem; border-top: 1px dashed #64748b;">`;
            // Show Column A if available, then item name
            const columnADisplay = specialItem.columnA ? `${specialItem.columnA} - ` : '';
            html += `<span style="font-weight: 600; color: #64748b; font-size: 0.85em;">${columnADisplay}${specialItem.name} (Row ${specialItem.rowNumber}):</span><br>`;

            // Show each check with actual value in green (valid) or red with expected (invalid)
            const checkResults = specialItem.checks.map(c => {
                if (c.isValid) {
                    return `<span style="color: #065f46; font-weight: 600;">${c.label}: ${c.actual}</span>`;
                } else {
                    return `<span style="color: #991b1b; font-weight: 600;">${c.label}: ${c.actual}</span> <span style="font-size: 0.85em; color: #849bba;">(Expected: ${c.expected})</span>`;
                }
            }).join(', ');
            html += checkResults;

            html += `</div>`;
        }

        return html;
    }

    formatWastageCells(validItems, invalidItems) {
        let html = '';

        // Show valid cells in green (just cell addresses)
        if (validItems.length > 0) {
            const validCells = validItems.map(item => item.cellAddress).join(', ');
            html += `<span style="color: #065f46; font-weight: 600;">${validCells}</span>`;
        }

        // Show invalid cells in red (cell address with value)
        if (invalidItems.length > 0) {
            if (validItems.length > 0) {
                html += '<br>';
            }
            const invalidCells = invalidItems.map(item => {
                return `<span style="color: #991b1b; font-weight: 600;">${item.cellAddress}: ${item.actual}</span>`;
            }).join(', ');
            html += invalidCells;
        }

        // If no items found
        if (validItems.length === 0 && invalidItems.length === 0) {
            html = `<span style="color: #64748b;">No items found in section</span>`;
        }

        return html;
    }
}

// Initialize processor
window.pranaProcessor = new PranaProcessor();
