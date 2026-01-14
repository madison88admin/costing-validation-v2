/**
 * Columbia Processing Logic
 * Automatically loads Columbia_CostBreakdown.csv from assets/data folder
 */

class ColumbiaProcessor {
    constructor() {
        this.columbiaCostData = null;
        this.bcbdResults = [];
    }

    /**
     * Initialize - Load Columbia Cost Breakdown CSV automatically
     */
    async initialize() {
        try {
            // Fetch the Columbia_CostBreakdown.csv file from assets/data folder
            const response = await fetch('assets/data/Columbia_CostBreakdown.csv');
            if (!response.ok) {
                throw new Error('Failed to load Columbia_CostBreakdown.csv');
            }

            const csvText = await response.text();
            this.columbiaCostData = this.parseCSV(csvText);

            // Display the loaded data in the OB drop zone
            this.displayColumbiaCostData();

            console.log('Columbia Cost Breakdown loaded successfully:', this.columbiaCostData);
        } catch (error) {
            console.error('Error loading Columbia Cost Breakdown:', error);
            this.displayError('Failed to load Columbia_CostBreakdown.csv from assets/data folder');
        }
    }

    /**
     * Parse CSV text into array of objects
     * Format: Description, PartNumber, UnitPrice, Quantity, Wastage
     */
    parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        const data = [];

        lines.forEach(line => {
            const values = line.split(',').map(val => val.trim());

            // First line is efficiency
            if (values[0] && values[0].toLowerCase().includes('efficiency')) {
                data.push({
                    description: values[0],
                    efficiency: values[1] || ''
                });
                return;
            }

            // Overhead line
            if (values[0] && values[0].toLowerCase().includes('overhead')) {
                data.push({
                    description: values[0],
                    overhead: values[1] || ''
                });
                return;
            }

            // Profit line
            if (values[0] && values[0].toLowerCase().includes('profit')) {
                data.push({
                    description: values[0],
                    profit: values[1] || ''
                });
                return;
            }

            // Standard format: Description, PartNumber, UnitPrice, Quantity, Wastage
            data.push({
                description: values[0] || '',
                partNumber: values[1] || '',
                unitPrice: values[2] || '',
                quantity: values[3] || '',
                wastage: values[4] || ''
            });
        });

        return data;
    }

    /**
     * Display Columbia Cost Breakdown data in the OB drop zone
     */
    displayColumbiaCostData() {
        const obDropZone = document.getElementById('obDropZone-v3');
        if (!obDropZone) return;

        let contentHTML = `
            <div class="burton-cost-container">
                <div class="burton-cost-items">
        `;

        // Display each line from the CSV
        this.columbiaCostData.forEach((item, index) => {
            // Handle efficiency line differently
            if (item.efficiency !== undefined) {
                contentHTML += `
                    <div class="burton-cost-item">
                        <div class="burton-item-line"><strong>${item.description}</strong> ${item.efficiency}</div>
                    </div>
                `;
                return;
            }

            contentHTML += `
                <div class="burton-cost-item">
                    <div class="burton-item-line"><strong>${item.description}</strong></div>
                    ${item.partNumber ? `<div class="burton-item-line"><strong>Part #:</strong> ${item.partNumber}</div>` : ''}
                    <div class="burton-item-line"><strong>Unit Price:</strong> ${this.formatToThreeDecimals(item.unitPrice)}</div>
                    <div class="burton-item-line"><strong>Qty:</strong> ${item.quantity}</div>
                    <div class="burton-item-line"><strong>Wastage:</strong> ${this.formatToThreeDecimals(item.wastage)}</div>
                </div>
            `;
        });

        contentHTML += `
                </div>
            </div>
        `;

        obDropZone.innerHTML = contentHTML;
    }

    /**
     * Display error message in the OB drop zone
     */
    displayError(errorMessage) {
        const obDropZone = document.getElementById('obDropZone-v3');
        if (!obDropZone) return;

        obDropZone.innerHTML = `
            <div class="drop-zone-content">
                <div style="background: #fee; border-left: 4px solid #dc3545; padding: 1.5rem; border-radius: 8px;">
                    <p style="color: #dc3545; font-weight: 600; margin-bottom: 0.5rem;">
                        Error Loading File
                    </p>
                    <p style="color: #721c24; font-size: 0.95rem;">
                        ${errorMessage}
                    </p>
                </div>
            </div>
        `;
    }

    /**
     * Format a numeric value to 3 decimal places
     */
    formatToThreeDecimals(value) {
        if (!value || value === '') return value;
        const cleanValue = value.toString().replace(/[$,\s]/g, '');
        const numValue = parseFloat(cleanValue);
        if (isNaN(numValue)) return value;
        return numValue.toFixed(3);
    }

    /**
     * Process files and generate results
     */
    async processFiles(bcbdFiles) {
        this.bcbdResults = [];

        try {
            if (!this.columbiaCostData || this.columbiaCostData.length === 0) {
                return this.generateErrorHTML('Columbia Cost Breakdown data not loaded');
            }

            if (!bcbdFiles || bcbdFiles.length === 0) {
                return this.generateErrorHTML('Please upload Buyer CBD files');
            }

            // Process each BCBD file
            for (const file of bcbdFiles) {
                const buyerData = await this.parseBuyerCBDFile(file);
                const comparisonResults = this.compareWithOB(buyerData);
                this.bcbdResults.push({
                    fileName: file.name,
                    results: comparisonResults
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

                    // Extract data from specific cells
                    const extractedData = this.extractColumbiaData(jsonData);
                    resolve(extractedData);
                } catch (error) {
                    reject(error);
                }
            };

            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Extract Columbia data from the Excel file
     * Efficiency% is at cell M19 (column 12, row 18 in 0-indexed)
     * Overhead is at cell O21 (column 14, row 20 in 0-indexed)
     * Profit is at cell M22 (column 12, row 21 in 0-indexed)
     * Items are found by searching Column A for keywords:
     * - Material is in Column B (index 1)
     * - FOB Cost is in Column K (index 10)
     * - Factory Usage is in Column O (index 14)
     * - Wastage is in Column Y (index 24)
     */
    extractColumbiaData(jsonData) {
        const data = {
            efficiency: '',
            overhead: '',
            profit: '',
            items: []
        };

        // Efficiency% at M19 (column M = index 12, row 19 = index 18)
        if (jsonData[18] && jsonData[18][12] !== undefined) {
            data.efficiency = jsonData[18][12].toString().trim();
        }

        // Overhead at O21 (column O = index 14, row 21 = index 20)
        if (jsonData[20] && jsonData[20][14] !== undefined) {
            data.overhead = jsonData[20][14].toString().trim();
        }

        // Profit at M22 (column M = index 12, row 22 = index 21)
        if (jsonData[21] && jsonData[21][12] !== undefined) {
            data.profit = jsonData[21][12].toString().trim();
        }

        // Get the item keywords to search for from our CSV (excluding efficiency, overhead, profit)
        const itemsToFind = this.columbiaCostData
            .filter(item => item.efficiency === undefined && item.overhead === undefined && item.profit === undefined)
            .map(item => item.description);

        // Get unique keywords
        const uniqueKeywords = [...new Set(itemsToFind)];

        // Search through all rows in Column A for matching keywords
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row[0]) continue;

            const cellA = row[0].toString().trim();

            // Check if this row matches any of our keywords
            for (const keyword of uniqueKeywords) {
                if (cellA.toLowerCase().includes(keyword.toLowerCase())) {
                    // Found a match - extract all relevant columns
                    const material = row[1] ? row[1].toString().trim() : '';           // Column B
                    const fobCost = row[10] ? row[10].toString().trim() : '';           // Column K
                    const factoryUsage = row[14] ? row[14].toString().trim() : '';      // Column O
                    const wastage = row[24] ? row[24].toString().trim() : '';           // Column Y

                    data.items.push({
                        keyword: keyword,
                        foundText: cellA,
                        material: material,
                        fobCost: fobCost,
                        factoryUsage: factoryUsage,
                        wastage: wastage,
                        rowIndex: i
                    });
                }
            }
        }

        console.log('Extracted Columbia data:', data);
        return data;
    }

    /**
     * Compare Buyer CBD data with OB data
     */
    compareWithOB(buyerData) {
        const results = [];

        // Get expected efficiency from CSV (first item with efficiency property)
        const efficiencyItem = this.columbiaCostData.find(item => item.efficiency !== undefined);
        const expectedEfficiency = efficiencyItem ? efficiencyItem.efficiency : '';

        // Compare Efficiency
        const efficiencyStatus = this.compareNumericField(expectedEfficiency, buyerData.efficiency);
        results.push({
            itemName: 'Efficiency%',
            obMaterial: expectedEfficiency,
            buyerMaterial: buyerData.efficiency,
            materialStatus: efficiencyStatus,
            obFobCost: '-',
            buyerFobCost: '-',
            fobCostStatus: 'VALID',
            obFactoryUsage: '-',
            buyerFactoryUsage: '-',
            factoryUsageStatus: 'VALID',
            obWastage: '-',
            buyerWastage: '-',
            wastageStatus: 'VALID'
        });

        // Get expected overhead from CSV
        const overheadItem = this.columbiaCostData.find(item => item.overhead !== undefined);
        const expectedOverhead = overheadItem ? overheadItem.overhead : '';

        // Compare Overhead (0.35 to 0.36 is valid)
        const overheadStatus = this.compareOverhead(expectedOverhead, buyerData.overhead);
        results.push({
            itemName: 'Overhead',
            obMaterial: expectedOverhead,
            buyerMaterial: buyerData.overhead,
            materialStatus: overheadStatus,
            obFobCost: '-',
            buyerFobCost: '-',
            fobCostStatus: 'VALID',
            obFactoryUsage: '-',
            buyerFactoryUsage: '-',
            factoryUsageStatus: 'VALID',
            obWastage: '-',
            buyerWastage: '-',
            wastageStatus: 'VALID'
        });

        // Get expected profit from CSV
        const profitItem = this.columbiaCostData.find(item => item.profit !== undefined);
        const expectedProfit = profitItem ? profitItem.profit : '';

        // Compare Profit (4% to 4.99% is valid)
        const profitStatus = this.compareProfit(expectedProfit, buyerData.profit);
        results.push({
            itemName: 'Profit',
            obMaterial: expectedProfit,
            buyerMaterial: buyerData.profit,
            materialStatus: profitStatus,
            obFobCost: '-',
            buyerFobCost: '-',
            fobCostStatus: 'VALID',
            obFactoryUsage: '-',
            buyerFactoryUsage: '-',
            factoryUsageStatus: 'VALID',
            obWastage: '-',
            buyerWastage: '-',
            wastageStatus: 'VALID'
        });

        // Get all non-efficiency, non-overhead, non-profit items from CSV
        const csvItems = this.columbiaCostData.filter(item => item.efficiency === undefined && item.overhead === undefined && item.profit === undefined);

        // Compare each CSV item with found items in BCBD
        for (const csvItem of csvItems) {
            // Find matching item in buyer data by keyword
            const buyerItem = buyerData.items.find(
                bi => bi.keyword.toLowerCase() === csvItem.description.toLowerCase() &&
                      bi.material === csvItem.partNumber
            );

            if (buyerItem) {
                // Found exact match - compare all fields
                results.push({
                    itemName: csvItem.description,
                    obMaterial: csvItem.partNumber,
                    buyerMaterial: buyerItem.material,
                    materialStatus: csvItem.partNumber === buyerItem.material ? 'VALID' : 'INVALID',
                    obFobCost: csvItem.unitPrice,
                    buyerFobCost: buyerItem.fobCost,
                    fobCostStatus: this.compareNumericField(csvItem.unitPrice, buyerItem.fobCost),
                    obFactoryUsage: csvItem.quantity,
                    buyerFactoryUsage: buyerItem.factoryUsage,
                    factoryUsageStatus: this.compareNumericField(csvItem.quantity, buyerItem.factoryUsage),
                    obWastage: csvItem.wastage,
                    buyerWastage: buyerItem.wastage,
                    wastageStatus: this.compareNumericField(csvItem.wastage, buyerItem.wastage)
                });
            } else {
                // Check if keyword exists with any material
                const keywordMatch = buyerData.items.find(
                    bi => bi.keyword.toLowerCase() === csvItem.description.toLowerCase()
                );

                if (keywordMatch) {
                    // Keyword found but material doesn't match
                    results.push({
                        itemName: csvItem.description,
                        obMaterial: csvItem.partNumber,
                        buyerMaterial: keywordMatch.material,
                        materialStatus: 'INVALID',
                        obFobCost: csvItem.unitPrice,
                        buyerFobCost: keywordMatch.fobCost,
                        fobCostStatus: this.compareNumericField(csvItem.unitPrice, keywordMatch.fobCost),
                        obFactoryUsage: csvItem.quantity,
                        buyerFactoryUsage: keywordMatch.factoryUsage,
                        factoryUsageStatus: this.compareNumericField(csvItem.quantity, keywordMatch.factoryUsage),
                        obWastage: csvItem.wastage,
                        buyerWastage: keywordMatch.wastage,
                        wastageStatus: this.compareNumericField(csvItem.wastage, keywordMatch.wastage)
                    });
                } else {
                    // Keyword not found at all
                    results.push({
                        itemName: csvItem.description,
                        obMaterial: csvItem.partNumber,
                        buyerMaterial: 'NOT FOUND',
                        materialStatus: 'INVALID',
                        obFobCost: csvItem.unitPrice,
                        buyerFobCost: 'NOT FOUND',
                        fobCostStatus: 'INVALID',
                        obFactoryUsage: csvItem.quantity,
                        buyerFactoryUsage: 'NOT FOUND',
                        factoryUsageStatus: 'INVALID',
                        obWastage: csvItem.wastage,
                        buyerWastage: 'NOT FOUND',
                        wastageStatus: 'INVALID'
                    });
                }
            }
        }

        return results;
    }

    /**
     * Compare numeric fields
     */
    compareNumericField(obValue, buyerValue) {
        const cleanOB = (obValue || '').toString().replace(/[$,\s%]/g, '');
        const cleanBuyer = (buyerValue || '').toString().replace(/[$,\s%]/g, '');

        const obNum = parseFloat(cleanOB);
        const buyerNum = parseFloat(cleanBuyer);

        if (isNaN(obNum) || isNaN(buyerNum)) {
            return 'INVALID';
        }

        // Round both values to 2 decimal places for comparison
        const obRounded = parseFloat(obNum.toFixed(2));
        const buyerRounded = parseFloat(buyerNum.toFixed(2));

        if (obRounded === buyerRounded) {
            return 'VALID';
        }

        return 'INVALID';
    }

    /**
     * Compare Overhead field (0.35 to 0.36 is valid)
     */
    compareOverhead(expectedValue, buyerValue) {
        const cleanExpected = (expectedValue || '').toString().replace(/[$,\s%]/g, '');
        const cleanBuyer = (buyerValue || '').toString().replace(/[$,\s%]/g, '');

        const expectedNum = parseFloat(cleanExpected);
        const buyerNum = parseFloat(cleanBuyer);

        if (isNaN(expectedNum) || isNaN(buyerNum)) {
            return 'INVALID';
        }

        // Valid range: 0.35 to 0.36
        if (buyerNum >= 0.35 && buyerNum <= 0.36) {
            return 'VALID';
        }

        return 'INVALID';
    }

    /**
     * Compare Profit field (4% to 4.99% is valid)
     */
    compareProfit(expectedValue, buyerValue) {
        const cleanExpected = (expectedValue || '').toString().replace(/[$,\s%]/g, '');
        const cleanBuyer = (buyerValue || '').toString().replace(/[$,\s%]/g, '');

        let expectedNum = parseFloat(cleanExpected);
        let buyerNum = parseFloat(cleanBuyer);

        if (isNaN(expectedNum) || isNaN(buyerNum)) {
            return 'INVALID';
        }

        // Convert decimal values (0.048) to percentage (4.8)
        if (buyerNum < 1) {
            buyerNum = buyerNum * 100;
        }
        if (expectedNum < 1) {
            expectedNum = expectedNum * 100;
        }

        // Valid range: 4% to 4.99%
        if (buyerNum >= 4 && buyerNum < 5) {
            return 'VALID';
        }

        return 'INVALID';
    }

    /**
     * Format field value with color coding
     */
    formatFieldValue(obValue, buyerValue, status, itemName = '') {
        const isValid = status === 'VALID';
        const color = isValid ? '#065f46' : '#991b1b';
        let displayValue = buyerValue || 'Empty';
        let expectedValue = obValue;

        // Convert decimal to percentage for Profit display
        if (itemName === 'Profit' && displayValue !== 'Empty') {
            const cleanValue = displayValue.toString().replace(/[$,\s%]/g, '');
            const numValue = parseFloat(cleanValue);
            if (!isNaN(numValue) && numValue < 1) {
                displayValue = (numValue * 100).toFixed(2) + '%';
            }
        }

        // Format Overhead expected range
        if (itemName === 'Overhead' && !isValid) {
            expectedValue = '0.35 to 0.36';
        }

        if (isValid) {
            return `<span style="color: ${color}; font-weight: 600;">${displayValue}</span>`;
        } else {
            return `<span style="color: ${color}; font-weight: 600;">${displayValue}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${expectedValue}</span>`;
        }
    }

    /**
     * Generate HTML for results display
     */
    generateResultsHTML(results) {
        if (!results || results.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">Columbia Cost Breakdown Loaded</p>
                    <p>Ready for processing. Upload Buyer CBD files to continue.</p>
                </div>
            `;
        }

        let html = '';

        for (const fileResult of results) {
            html += `<div class="file-result-group">`;

            // Count fully valid items (all fields match)
            const totalItems = fileResult.results.length;
            const validItems = fileResult.results.filter(r =>
                r.materialStatus === 'VALID' &&
                r.fobCostStatus === 'VALID' &&
                r.factoryUsageStatus === 'VALID' &&
                r.wastageStatus === 'VALID'
            ).length;

            html += `
                <div class="file-summary-box">
                    <strong>File:</strong> ${fileResult.fileName}<br>
                    <strong>Summary:</strong> ${validItems} out of ${totalItems} items fully match
                </div>
            `;

            // Create comparison table
            html += `
                <table class="results-table">
                    <thead>
                        <tr class="header-labels-row">
                            <th>Item</th>
                            <th>Material</th>
                            <th>FOB Cost</th>
                            <th>Factory Usage</th>
                            <th>Wastage</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            for (const item of fileResult.results) {
                // Skip Hangtag Package Part with material 1234
                if (item.itemName === 'Hangtag Package Part' && item.obMaterial === '1234') {
                    continue;
                }

                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">${item.itemName}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(item.obMaterial, item.buyerMaterial, item.materialStatus, item.itemName)}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(item.obFobCost, item.buyerFobCost, item.fobCostStatus, item.itemName)}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(item.obFactoryUsage, item.buyerFactoryUsage, item.factoryUsageStatus, item.itemName)}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(item.obWastage, item.buyerWastage, item.wastageStatus, item.itemName)}</td>
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
}

// Initialize the processor
window.columbiaProcessor = new ColumbiaProcessor();

// Auto-load Columbia Cost Breakdown when V3 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    // Check if we're on V3 tab and initialize
    const v3Tab = document.querySelector('[data-tab="v3"]');
    if (v3Tab) {
        v3Tab.addEventListener('click', () => {
            if (!window.columbiaProcessor.columbiaCostData) {
                window.columbiaProcessor.initialize();
            }
        });
    }

    // If V3 tab is already active on load, initialize immediately
    const v3TabContent = document.getElementById('tab-v3');
    if (v3TabContent && v3TabContent.classList.contains('active')) {
        window.columbiaProcessor.initialize();
    }
});
