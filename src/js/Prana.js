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
                expectedWastage: '3%'
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

        return result;
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
                    for (const specialItem of sectionConfig.specialItems) {
                        if (cellALower === specialItem.name.toLowerCase()) {
                            const specialResult = {
                                name: cellA,
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

        html += `
                </tbody>
            </table>
            </div>
        `;

        return html;
    }

    formatSpecialItems(specialItemResults) {
        if (!specialItemResults || specialItemResults.length === 0) {
            return '';
        }

        let html = '';

        for (const specialItem of specialItemResults) {
            html += `<div style="margin-top: 0.5rem; padding-top: 0.5rem; border-top: 1px dashed #64748b;">`;
            html += `<span style="font-weight: 600; color: #64748b; font-size: 0.85em;">${specialItem.name} (Row ${specialItem.rowNumber}):</span><br>`;

            const validChecks = specialItem.checks.filter(c => c.isValid);
            const invalidChecks = specialItem.checks.filter(c => !c.isValid);

            if (validChecks.length > 0) {
                const validCells = validChecks.map(c => c.cellAddress).join(', ');
                html += `<span style="color: #065f46; font-weight: 600;">${validCells}</span>`;
            }

            if (invalidChecks.length > 0) {
                if (validChecks.length > 0) html += '<br>';
                const invalidCells = invalidChecks.map(c => {
                    return `<span style="color: #991b1b; font-weight: 600;">${c.cellAddress}: ${c.actual}</span> <span style="font-size: 0.85em; color: #849bba;">(Expected: ${c.expected})</span>`;
                }).join(', ');
                html += invalidCells;
            }

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
