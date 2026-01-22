/**
 * Vuori Cost Breakdown Processor
 * Validates BCBD files against standard cost breakdown data from CSV
 *
 * CSV Columns mapping to Excel:
 * 1st Column (CSV) = Column D (Excel) - Cost Material Description
 * 2nd Column (CSV) = Column E (Excel) - Cost Material Code
 * 3rd Column (CSV) = Column F (Excel) - Cost Material Subtype
 * 4th Column (CSV) = Column G (Excel) - Cost Construction
 * 5th Column (CSV) = Column S (Excel) - Supplier Unit Cost
 * 6th Column (CSV) = Column W (Excel) - Material Wastage %
 */

class VuoriProcessor {
    constructor() {
        this.csvData = [];
        this.validationItems = [];
        this.columnMapping = {
            materialDesc: 3,    // Column D (index 3)
            materialCode: 4,    // Column E (index 4)
            materialSubtype: 5, // Column F (index 5)
            construction: 6,    // Column G (index 6)
            supplierCost: 18,   // Column S (index 18)
            wastage: 22         // Column W (index 22)
        };
    }

    async initialize() {
        await this.loadCSVData();
        this.displayValidationRules();
    }

    async loadCSVData() {
        try {
            const response = await fetch('assets/data/Vuori_CostBreakdown.csv');
            const csvText = await response.text();
            this.parseCSV(csvText);
            console.log('Vuori CSV data loaded:', this.validationItems);
        } catch (error) {
            console.error('Error loading Vuori CSV:', error);
        }
    }

    parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        this.validationItems = [];

        for (const line of lines) {
            if (!line.trim()) continue;

            // Parse CSV line (handling potential commas in values)
            const values = this.parseCSVLine(line);

            if (values.length >= 6) {
                const materialDesc = values[0].trim();
                const materialCode = values[1].trim();
                const materialSubtype = values[2].trim();
                const construction = values[3].trim();
                const supplierCost = values[4].trim();
                const wastage = values[5].trim();

                this.validationItems.push({
                    materialDesc: materialDesc,
                    materialCode: materialCode,
                    materialSubtype: materialSubtype,
                    construction: construction,
                    supplierCost: supplierCost,
                    wastage: wastage
                });
            }
        }
    }

    parseCSVLine(line) {
        const values = [];
        let current = '';
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === ',' && !inQuotes) {
                values.push(current);
                current = '';
            } else {
                current += char;
            }
        }
        values.push(current);
        return values;
    }

    displayValidationRules() {
        const obDropZone = document.getElementById('obDropZone-v12');
        if (!obDropZone) return;

        let html = `
            <div class="burton-cost-container">
                <div class="burton-cost-items">
        `;

        // Display each item from the CSV
        for (const item of this.validationItems) {
            html += `
                <div class="burton-cost-item">
                    <div class="burton-item-line"><strong>Material Desc (D):</strong> ${item.materialDesc}</div>
                    <div class="burton-item-line"><strong>Material Code (E):</strong> ${item.materialCode}</div>
                    <div class="burton-item-line"><strong>Material Subtype (F):</strong> ${item.materialSubtype}</div>
                    <div class="burton-item-line"><strong>Construction (G):</strong> ${item.construction}</div>
                    <div class="burton-item-line"><strong>Supplier Cost (S):</strong> ${item.supplierCost}</div>
                    <div class="burton-item-line"><strong>Wastage (W):</strong> ${item.wastage}</div>
                </div>
            `;
        }

        html += `
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
                    items: []
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
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

                    const validationResults = this.validateData(jsonData);

                    resolve({
                        fileName: file.name,
                        items: validationResults,
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

    validateData(jsonData) {
        const results = [];
        const colMap = this.columnMapping;

        for (const validationItem of this.validationItems) {
            // If Material Description is '-', find ALL matching rows by Material Subtype
            // Otherwise, find single match by Material Subtype + Material Code
            const isMultiMatch = validationItem.materialDesc === '-';
            let foundAny = false;

            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                const cellF = row[colMap.materialSubtype] ? String(row[colMap.materialSubtype]).trim() : '';
                const cellE = row[colMap.materialCode] ? String(row[colMap.materialCode]).trim() : '';

                // Match by Material Subtype (case-insensitive)
                let isMatch = cellF.toLowerCase() === validationItem.materialSubtype.toLowerCase();

                // If Material Code in CSV is not '-', also check Material Code matches
                if (isMatch && validationItem.materialCode !== '-') {
                    isMatch = cellE.toLowerCase() === validationItem.materialCode.toLowerCase();
                }

                if (isMatch) {
                    foundAny = true;

                    const itemResult = {
                        validationItem: validationItem,
                        found: true,
                        rowNumber: i + 1,
                        checks: [],
                        isValid: true
                    };

                    // Check Material Description (Column D)
                    const actualMaterialDesc = row[colMap.materialDesc];
                    const materialDescCheck = this.validateValue(
                        actualMaterialDesc,
                        validationItem.materialDesc,
                        'Material Desc',
                        i + 1,
                        'D'
                    );
                    itemResult.checks.push(materialDescCheck);

                    // Check Material Code (Column E)
                    const actualMaterialCode = row[colMap.materialCode];
                    const materialCodeCheck = this.validateValue(
                        actualMaterialCode,
                        validationItem.materialCode,
                        'Material Code',
                        i + 1,
                        'E'
                    );
                    itemResult.checks.push(materialCodeCheck);

                    // Check Material Subtype (Column F)
                    const materialSubtypeCheck = this.validateValue(
                        cellF,
                        validationItem.materialSubtype,
                        'Material Subtype',
                        i + 1,
                        'F'
                    );
                    itemResult.checks.push(materialSubtypeCheck);

                    // Check Construction (Column G)
                    const actualConstruction = row[colMap.construction];
                    const constructionCheck = this.validateValue(
                        actualConstruction,
                        validationItem.construction,
                        'Construction',
                        i + 1,
                        'G'
                    );
                    itemResult.checks.push(constructionCheck);

                    // Check Supplier Cost (Column S)
                    const actualSupplierCost = row[colMap.supplierCost];
                    const supplierCostCheck = this.validateValue(
                        actualSupplierCost,
                        validationItem.supplierCost,
                        'Supplier Cost',
                        i + 1,
                        'S'
                    );
                    itemResult.checks.push(supplierCostCheck);

                    // Check Wastage (Column W)
                    const actualWastage = row[colMap.wastage];
                    const wastageCheck = this.validateValue(
                        actualWastage,
                        validationItem.wastage,
                        'Wastage',
                        i + 1,
                        'W'
                    );
                    itemResult.checks.push(wastageCheck);

                    itemResult.isValid = itemResult.checks.every(c => c.isValid);
                    results.push(itemResult);

                    // If not multi-match mode, break after first match
                    if (!isMultiMatch) {
                        break;
                    }
                }
            }

            // If no matches found, add a "not found" result
            if (!foundAny) {
                results.push({
                    validationItem: validationItem,
                    found: false,
                    rowNumber: -1,
                    checks: [],
                    isValid: false
                });
            }
        }

        return results;
    }

    validateValue(actual, expected, label, rowNum, colLetter) {
        // Handle empty/null values
        let actualStr = actual !== undefined && actual !== null && actual !== ''
            ? String(actual).trim()
            : '-';
        const expectedStr = expected !== undefined && expected !== null && expected !== ''
            ? String(expected).trim()
            : '-';

        // If expected is '-', skip validation (always valid)
        if (expectedStr === '-') {
            return {
                label: label,
                expected: expectedStr,
                actual: actualStr,
                isValid: true,
                skipped: true,
                cellAddress: `${colLetter}${rowNum}`
            };
        }

        // Handle percentage conversion for Wastage column
        // Excel stores 5% as 0.05, so multiply by 100 for display and comparison
        if (label === 'Wastage' && expectedStr.includes('%')) {
            const actualNum = parseFloat(actualStr);
            if (!isNaN(actualNum) && actualNum < 1) {
                // Convert decimal to percentage (0.05 -> 5.00%)
                actualStr = (actualNum * 100).toFixed(2) + '%';
            }
        }

        // Normalize values for comparison
        const normalizeValue = (val) => {
            if (val === '-') return '-';
            // Remove $, %, and whitespace, convert to lowercase
            let normalized = String(val).replace(/[$%]/g, '').trim().toLowerCase();
            // Try to parse as number for numeric comparison
            const num = parseFloat(normalized);
            if (!isNaN(num)) {
                return num.toFixed(3);
            }
            return normalized;
        };

        const normalizedActual = normalizeValue(actualStr);
        const normalizedExpected = normalizeValue(expectedStr);

        const isValid = normalizedActual === normalizedExpected;

        return {
            label: label,
            expected: expectedStr,
            actual: actualStr,
            isValid: isValid,
            cellAddress: `${colLetter}${rowNum}`
        };
    }

    generateResultsHTML(results) {
        let html = '';

        for (const fileResult of results) {
            // Wrap each file's results in a group container (like Burton)
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
                const totalItems = fileResult.items.length;
                const validItems = fileResult.items.filter(item => item.found && item.isValid).length;

                html += `
                    <div class="file-summary-box">
                        <strong>File:</strong> ${fileResult.fileName}<br>
                        <strong>Summary:</strong> ${validItems} out of ${totalItems} items fully match
                    </div>
                `;

                html += this.generateItemsTable(fileResult.items);
            }

            html += '</div>';
        }

        return html;
    }

    /**
     * Format field value with color coding (like Burton)
     */
    formatFieldValue(expected, actual, isValid, skipped) {
        // If expected is '-', this field is skipped - just show dash
        if (expected === '-' || skipped) {
            return `<span style="color: #6b7280; font-weight: 600;">-</span>`;
        }

        const color = isValid ? '#065f46' : '#991b1b';

        if (!actual || actual === '' || actual === '-') {
            return `<span style="color: #991b1b; font-weight: 600;">Empty</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${expected}</span>`;
        }

        if (isValid) {
            return `<span style="color: ${color}; font-weight: 600;">${actual}</span>`;
        } else {
            return `<span style="color: ${color}; font-weight: 600;">${actual}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${expected}</span>`;
        }
    }

    generateItemsTable(items) {
        let html = `
            <table class="results-table">
                <thead>
                    <tr class="header-labels-row">
                        <th>Material Desc (D)</th>
                        <th>Material Code (E)</th>
                        <th>Material Subtype (F)</th>
                        <th>Construction (G)</th>
                        <th>Supplier Cost (S)</th>
                        <th>Wastage (W)</th>
                    </tr>
                </thead>
                <tbody>
        `;

        for (const item of items) {
            const expectedMaterialDesc = item.validationItem.materialDesc;
            const expectedMaterialCode = item.validationItem.materialCode;
            const expectedMaterialSubtype = item.validationItem.materialSubtype;
            const expectedConstruction = item.validationItem.construction;
            const expectedSupplierCost = item.validationItem.supplierCost;
            const expectedWastage = item.validationItem.wastage;

            if (!item.found) {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">${expectedMaterialDesc}</td>
                        <td style="padding: 0.875rem 1rem;">${expectedMaterialCode}</td>
                        <td style="padding: 0.875rem 1rem;">${expectedMaterialSubtype}</td>
                        <td colspan="3" style="text-align: center; color: #991b1b; padding: 0.875rem 1rem;">
                            Not found in Buyer CBD file
                        </td>
                    </tr>
                `;
            } else {
                const getCheckByLabel = (label) => item.checks.find(c => c.label === label);

                const materialDescCheck = getCheckByLabel('Material Desc');
                const materialCodeCheck = getCheckByLabel('Material Code');
                const materialSubtypeCheck = getCheckByLabel('Material Subtype');
                const constructionCheck = getCheckByLabel('Construction');
                const supplierCostCheck = getCheckByLabel('Supplier Cost');
                const wastageCheck = getCheckByLabel('Wastage');

                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(expectedMaterialDesc, materialDescCheck?.actual, materialDescCheck?.isValid, materialDescCheck?.skipped)}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(expectedMaterialCode, materialCodeCheck?.actual, materialCodeCheck?.isValid, materialCodeCheck?.skipped)}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(expectedMaterialSubtype, materialSubtypeCheck?.actual, materialSubtypeCheck?.isValid, materialSubtypeCheck?.skipped)}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(expectedConstruction, constructionCheck?.actual, constructionCheck?.isValid, constructionCheck?.skipped)}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(expectedSupplierCost, supplierCostCheck?.actual, supplierCostCheck?.isValid, supplierCostCheck?.skipped)}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(expectedWastage, wastageCheck?.actual, wastageCheck?.isValid, wastageCheck?.skipped)}</td>
                    </tr>
                `;
            }
        }

        html += `
                </tbody>
            </table>
        `;

        return html;
    }

    exportToExcel(fileName, items) {
        // Parse items if it's a string
        if (typeof items === 'string') {
            items = JSON.parse(items);
        }

        const exportData = [];

        // Add header row
        exportData.push([
            'Expected Material Desc (D)', 'Actual Material Desc', 'Status',
            'Expected Material Code (E)', 'Actual Material Code', 'Status',
            'Expected Material Subtype (F)', 'Actual Material Subtype', 'Status',
            'Expected Construction (G)', 'Actual Construction', 'Status',
            'Expected Supplier Cost (S)', 'Actual Supplier Cost', 'Status',
            'Expected Wastage (W)', 'Actual Wastage', 'Status',
            'Overall Status'
        ]);

        for (const item of items) {
            const getCheckByLabel = (label) => item.checks?.find(c => c.label === label);

            if (!item.found) {
                exportData.push([
                    item.validationItem.materialDesc, 'Not Found', 'N/A',
                    item.validationItem.materialCode, 'Not Found', 'N/A',
                    item.validationItem.materialSubtype, 'Not Found', 'N/A',
                    item.validationItem.construction, 'Not Found', 'N/A',
                    item.validationItem.supplierCost, 'Not Found', 'N/A',
                    item.validationItem.wastage, 'Not Found', 'N/A',
                    'Not Found'
                ]);
            } else {
                const materialDescCheck = getCheckByLabel('Material Desc');
                const materialCodeCheck = getCheckByLabel('Material Code');
                const materialSubtypeCheck = getCheckByLabel('Material Subtype');
                const constructionCheck = getCheckByLabel('Construction');
                const supplierCostCheck = getCheckByLabel('Supplier Cost');
                const wastageCheck = getCheckByLabel('Wastage');

                exportData.push([
                    item.validationItem.materialDesc, materialDescCheck?.actual || '-', materialDescCheck?.isValid ? 'Valid' : 'Invalid',
                    item.validationItem.materialCode, materialCodeCheck?.actual || '-', materialCodeCheck?.isValid ? 'Valid' : 'Invalid',
                    item.validationItem.materialSubtype, materialSubtypeCheck?.actual || '-', materialSubtypeCheck?.isValid ? 'Valid' : 'Invalid',
                    item.validationItem.construction, constructionCheck?.actual || '-', constructionCheck?.isValid ? 'Valid' : 'Invalid',
                    item.validationItem.supplierCost, supplierCostCheck?.actual || '-', supplierCostCheck?.isValid ? 'Valid' : 'Invalid',
                    item.validationItem.wastage, wastageCheck?.actual || '-', wastageCheck?.isValid ? 'Valid' : 'Invalid',
                    item.isValid ? 'Valid' : 'Invalid'
                ]);
            }
        }

        // Create workbook and worksheet
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(exportData);

        // Set column widths
        ws['!cols'] = [
            { wch: 22 }, { wch: 20 }, { wch: 10 }, // Material Desc
            { wch: 22 }, { wch: 18 }, { wch: 10 }, // Material Code
            { wch: 24 }, { wch: 20 }, { wch: 10 }, // Material Subtype
            { wch: 22 }, { wch: 18 }, { wch: 10 }, // Construction
            { wch: 22 }, { wch: 18 }, { wch: 10 }, // Supplier Cost
            { wch: 20 }, { wch: 15 }, { wch: 10 }, // Wastage
            { wch: 15 }  // Overall Status
        ];

        XLSX.utils.book_append_sheet(wb, ws, 'Vuori Validation');

        // Generate filename
        const exportFileName = `Vuori_Validation_${fileName.replace(/\.[^/.]+$/, '')}_${new Date().toISOString().slice(0, 10)}.xlsx`;

        // Download
        XLSX.writeFile(wb, exportFileName);
    }
}

// Initialize processor
window.vuoriProcessor = new VuoriProcessor();
