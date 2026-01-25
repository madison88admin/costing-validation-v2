/**
 * Skida Cost Breakdown Processor
 * Validates BCBD files against standard cost breakdown data from CSV
 *
 * CSV Columns mapping to Excel:
 * 1st Column (CSV) = Column A (Excel) - Category (Thread, Cut & Sew/Labor, Freight, Other)
 * 2nd Column (CSV) = Column B (Excel) - Description (used with "Other" to identify Overhead, Profit, Packaging)
 * 3rd Column (CSV) = Column E (Excel) - Unit Cost
 * 4th Column (CSV) = Column F (Excel) - Quantity
 */

class SkidaProcessor {
    constructor() {
        this.csvData = [];
        this.validationItems = [];
    }

    async initialize() {
        await this.loadCSVData();
        this.displayValidationRules();
    }

    async loadCSVData() {
        try {
            const response = await fetch('assets/data/Skida_Costbreakdown.csv');
            const csvText = await response.text();
            this.parseCSV(csvText);
            console.log('Skida CSV data loaded:', this.validationItems);
        } catch (error) {
            console.error('Error loading Skida CSV:', error);
        }
    }

    parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        this.validationItems = [];

        for (const line of lines) {
            if (!line.trim()) continue;

            // Parse CSV line (handling potential commas in values)
            const values = this.parseCSVLine(line);

            if (values.length >= 4) {
                const category = values[0].trim();
                const description = values[1].trim();
                const unitCost = values[2].trim();
                const quantity = values[3].trim();

                this.validationItems.push({
                    category: category,
                    description: description,
                    unitCost: unitCost,
                    quantity: quantity,
                    // For "Other" category, we need to check Column B for the specific type
                    matchKey: category === 'Other' ? description : category
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
        const obDropZone = document.getElementById('obDropZone-v11');
        if (!obDropZone) return;

        let html = `
            <div class="burton-cost-container">
                <div class="burton-cost-items">
        `;

        // Display each item from the CSV
        for (const item of this.validationItems) {
            html += `
                <div class="burton-cost-item">
                    <div class="burton-item-line"><strong>${item.category}</strong></div>
                    <div class="burton-item-line"><strong>Description:</strong> ${item.description}</div>
                    <div class="burton-item-line"><strong>Unit Cost:</strong> ${item.unitCost}</div>
                    <div class="burton-item-line"><strong>Quantity:</strong> ${item.quantity}</div>
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

        // Column mapping: A=0, B=1, E=4, F=5
        const colMap = {
            category: 0,    // Column A
            description: 1, // Column B
            unitCost: 4,    // Column E
            quantity: 5     // Column F
        };

        for (const validationItem of this.validationItems) {
            const itemResult = {
                validationItem: validationItem,
                found: false,
                rowNumber: -1,
                checks: [],
                isValid: true
            };

            // Find matching row in Excel
            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                const cellA = row[colMap.category] ? String(row[colMap.category]).trim() : '';
                const cellB = row[colMap.description] ? String(row[colMap.description]).trim() : '';

                let isMatch = false;

                if (validationItem.category === 'Other') {
                    // For "Other" category, match Column A = "Other" AND Column B = description (Overhead, Profit, Packaging)
                    if (cellA.toLowerCase() === 'other' &&
                        cellB.toLowerCase() === validationItem.description.toLowerCase()) {
                        isMatch = true;
                    }
                } else {
                    // For other categories (Thread, Cut & Sew/Labor, Freight), match Column A
                    if (cellA.toLowerCase() === validationItem.category.toLowerCase()) {
                        isMatch = true;
                    }
                }

                if (isMatch) {
                    itemResult.found = true;
                    itemResult.rowNumber = i + 1;

                    // Capture actual values from Excel for display
                    itemResult.actualCategory = cellA || '-';
                    itemResult.actualDescription = cellB || '-';

                    // Check Category (Column A)
                    const categoryCheck = this.validateValue(
                        cellA,
                        validationItem.category,
                        'Category',
                        i + 1,
                        'A'
                    );
                    itemResult.checks.push(categoryCheck);

                    // Check Description (Column B)
                    const descriptionCheck = this.validateValue(
                        cellB,
                        validationItem.description,
                        'Description',
                        i + 1,
                        'B'
                    );
                    itemResult.checks.push(descriptionCheck);

                    // Check Unit Cost (Column E)
                    const actualUnitCost = row[colMap.unitCost];
                    const unitCostCheck = this.validateValue(
                        actualUnitCost,
                        validationItem.unitCost,
                        'Unit Cost',
                        i + 1,
                        'E'
                    );
                    itemResult.checks.push(unitCostCheck);

                    // Check Quantity (Column F)
                    const actualQuantity = row[colMap.quantity];
                    const quantityCheck = this.validateValue(
                        actualQuantity,
                        validationItem.quantity,
                        'Quantity',
                        i + 1,
                        'F'
                    );
                    itemResult.checks.push(quantityCheck);

                    itemResult.isValid = itemResult.checks.every(c => c.isValid);
                    break;
                }
            }

            results.push(itemResult);
        }

        return results;
    }

    validateValue(actual, expected, label, rowNum, colLetter) {
        // Handle empty/null values
        const actualStr = actual !== undefined && actual !== null && actual !== ''
            ? String(actual).trim()
            : '-';
        const expectedStr = expected !== undefined && expected !== null && expected !== ''
            ? String(expected).trim()
            : '-';

        // Normalize values for comparison
        const normalizeValue = (val) => {
            if (val === '-') return '-';
            // Remove $ and whitespace, convert to lowercase
            let normalized = String(val).replace(/\$/g, '').trim().toLowerCase();
            // Try to parse as number for numeric comparison
            const num = parseFloat(normalized);
            if (!isNaN(num)) {
                return num.toFixed(2);
            }
            return normalized;
        };

        const normalizedActual = normalizeValue(actualStr);
        const normalizedExpected = normalizeValue(expectedStr);

        const isValid = normalizedActual === normalizedExpected ||
                       (expectedStr === '-' && actualStr === '-');

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

        // Add export button at the top
        html += `
            <div style="margin-bottom: 15px; display: flex; justify-content: flex-end; align-items: center;">
                <button onclick="window.skidaProcessor.exportToPDF()" class="export-btn">
                    Export
                </button>
            </div>
        `;

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
    formatFieldValue(expected, actual, isValid) {
        const color = isValid ? '#065f46' : '#991b1b';

        if (!actual || actual === '' || actual === '-') {
            if (expected === '-') {
                return `<span style="color: #065f46; font-weight: 600;">-</span>`;
            }
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
                        <th>Item Name</th>
                        <th>Description</th>
                        <th>Unit Cost</th>
                        <th>Quantity</th>
                    </tr>
                </thead>
                <tbody>
        `;

        for (const item of items) {
            const itemName = item.validationItem.category;
            const expectedDesc = item.validationItem.description;
            const expectedCost = item.validationItem.unitCost;
            const expectedQty = item.validationItem.quantity;

            if (!item.found) {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">${itemName}</td>
                        <td colspan="3" style="text-align: center; color: #991b1b; padding: 0.875rem 1rem;">
                            Not found in Buyer CBD file
                        </td>
                    </tr>
                `;
            } else {
                const getCheckByLabel = (label) => item.checks.find(c => c.label === label);

                const descriptionCheck = getCheckByLabel('Description');
                const unitCostCheck = getCheckByLabel('Unit Cost');
                const quantityCheck = getCheckByLabel('Quantity');

                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">${itemName}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(expectedDesc, descriptionCheck?.actual, descriptionCheck?.isValid)}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(expectedCost, unitCostCheck?.actual, unitCostCheck?.isValid)}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(expectedQty, quantityCheck?.actual, quantityCheck?.isValid)}</td>
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
        exportData.push(['Category', 'Description', 'Expected Unit Cost', 'Actual Unit Cost', 'Unit Cost Status', 'Expected Quantity', 'Actual Quantity', 'Quantity Status', 'Overall Status']);

        for (const item of items) {
            const displayName = item.validationItem.category === 'Other'
                ? `Other - ${item.validationItem.description}`
                : item.validationItem.category;

            if (!item.found) {
                exportData.push([
                    displayName,
                    item.validationItem.description,
                    item.validationItem.unitCost,
                    'Not Found',
                    'N/A',
                    item.validationItem.quantity,
                    'Not Found',
                    'N/A',
                    'Not Found'
                ]);
            } else {
                const unitCostCheck = item.checks.find(c => c.label === 'Unit Cost');
                const quantityCheck = item.checks.find(c => c.label === 'Quantity');

                exportData.push([
                    displayName,
                    item.actualDescription || item.validationItem.description,
                    item.validationItem.unitCost,
                    unitCostCheck ? unitCostCheck.actual : '-',
                    unitCostCheck ? (unitCostCheck.isValid ? 'Valid' : 'Invalid') : 'N/A',
                    item.validationItem.quantity,
                    quantityCheck ? quantityCheck.actual : '-',
                    quantityCheck ? (quantityCheck.isValid ? 'Valid' : 'Invalid') : 'N/A',
                    item.isValid ? 'Valid' : 'Invalid'
                ]);
            }
        }

        // Create workbook and worksheet
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(exportData);

        // Set column widths
        ws['!cols'] = [
            { wch: 20 }, // Category
            { wch: 40 }, // Description
            { wch: 18 }, // Expected Unit Cost
            { wch: 18 }, // Actual Unit Cost
            { wch: 15 }, // Unit Cost Status
            { wch: 18 }, // Expected Quantity
            { wch: 18 }, // Actual Quantity
            { wch: 15 }, // Quantity Status
            { wch: 15 }  // Overall Status
        ];

        XLSX.utils.book_append_sheet(wb, ws, 'Skida Validation');

        // Generate filename
        const exportFileName = `Skida_Validation_${fileName.replace(/\.[^/.]+$/, '')}_${new Date().toISOString().slice(0, 10)}.xlsx`;

        // Download
        XLSX.writeFile(wb, exportFileName);
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

        const config = window.pdfExporter.createSkidaConfig(this.fileResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }
}

// Initialize processor
window.skidaProcessor = new SkidaProcessor();
