/**
 * Helly Hansen Processing Logic
 * Automatically loads HellyHansen_CostBreakdown.csv from assets/data folder
 */

class HellyHansenProcessor {
    constructor() {
        this.hellyHansenCostData = null;
        this.bcbdResults = [];
    }

    /**
     * Initialize - Load Helly Hansen Cost Breakdown CSV automatically
     */
    async initialize() {
        try {
            const response = await fetch('assets/data/HellyHansen_CostBreakdown.csv');
            if (!response.ok) {
                throw new Error('Failed to load HellyHansen_CostBreakdown.csv');
            }

            const csvText = await response.text();
            this.hellyHansenCostData = this.parseCSV(csvText);

            this.displayHellyHansenCostData();

            console.log('Helly Hansen Cost Breakdown loaded:', this.hellyHansenCostData);
        } catch (error) {
            console.error('Error loading Helly Hansen Cost Breakdown:', error);
            this.displayError('Failed to load HellyHansen_CostBreakdown.csv');
        }
    }

    /**
     * Parse CSV text into array of objects
     * Format: Item, CONSM, U/P, Amount
     */
    parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        const data = [];

        lines.forEach(line => {
            const values = line.split(',').map(val => val.trim());

            data.push({
                item: values[0] || '',
                consm: values[1] || '',
                up: values[2] || '',
                amount: values[3] || ''
            });
        });

        return data;
    }

    /**
     * Display Helly Hansen Cost Breakdown data in the OB drop zone
     */
    displayHellyHansenCostData() {
        const obDropZone = document.getElementById('obDropZone-v4');
        if (!obDropZone) return;

        let contentHTML = `
            <div class="burton-cost-container">
                <div class="burton-cost-items">
        `;

        this.hellyHansenCostData.forEach((item, index) => {
            contentHTML += `
                <div class="burton-cost-item">
                    <div class="burton-item-line"><strong>${item.item}</strong></div>
                    <div class="burton-item-line"><strong>CONSM:</strong> ${item.consm}</div>
                    <div class="burton-item-line"><strong>U/P:</strong> ${item.up}</div>
                    <div class="burton-item-line"><strong>Amount:</strong> ${item.amount}</div>
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
     * Display error message
     */
    displayError(errorMessage) {
        const obDropZone = document.getElementById('obDropZone-v4');
        if (!obDropZone) return;

        obDropZone.innerHTML = `
            <div class="drop-zone-content">
                <div style="background: #fee; border-left: 4px solid #dc3545; padding: 1.5rem; border-radius: 8px;">
                    <p style="color: #dc3545; font-weight: 600;">${errorMessage}</p>
                </div>
            </div>
        `;
    }

    /**
     * Format a numeric value to 4 decimal places
     */
    formatToFourDecimals(value) {
        if (!value || value === '') return value;
        const cleanValue = value.toString().replace(/[$,\s]/g, '');
        const numValue = parseFloat(cleanValue);
        if (isNaN(numValue)) return value;
        return numValue.toFixed(4);
    }

    /**
     * Process files and generate results
     */
    async processFiles(bcbdFiles) {
        this.bcbdResults = [];

        try {
            if (!this.hellyHansenCostData || this.hellyHansenCostData.length === 0) {
                return this.generateErrorHTML('Helly Hansen Cost Breakdown data not loaded');
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
                    const extractedData = this.extractHellyHansenData(jsonData);
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
     * Extract Helly Hansen data from the Excel file
     * Item is in Column C (index 2)
     * CONSM is in Column H (index 7)
     * U/P is in Column I (index 8)
     * Amount is in Column J (index 9)
     */
    extractHellyHansenData(jsonData) {
        const data = {
            items: [],
            countryOfOrigin: ''
        };

        // Search for Country of Origin in Column C
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (row[2]) {
                const cellValue = row[2].toString().trim().toUpperCase();
                if (cellValue.includes('COUNTRY OF ORIGIN') || cellValue.includes('COUNTRY OF ORIGINAL')) {
                    if (cellValue.includes('INDO')) {
                        data.countryOfOrigin = 'INDO';
                    } else if (cellValue.includes('CHINA')) {
                        data.countryOfOrigin = 'CHINA';
                    }
                    break;
                }
            }
        }

        // Get the item keywords to search for from our CSV
        const itemsToFind = this.hellyHansenCostData.map(item => item.item);

        // Search through all rows in Column C for matching items
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row[2]) continue; // Column C

            const cellC = row[2].toString().trim();

            // Check if this row matches any of our item keywords
            for (const keyword of itemsToFind) {
                if (cellC.toLowerCase().includes(keyword.toLowerCase())) {
                    // Found a match - extract all relevant columns
                    const consm = row[7] ? row[7].toString().trim() : '';    // Column H
                    const up = row[8] ? row[8].toString().trim() : '';       // Column I
                    const amount = row[9] ? row[9].toString().trim() : '';   // Column J

                    data.items.push({
                        keyword: keyword,
                        foundText: cellC,
                        consm: consm,
                        up: up,
                        amount: amount,
                        rowIndex: i,
                        // Mark if this is a special row (only has amount, no consm/up)
                        isSpecialRow: !consm && !up && amount
                    });
                    break; // Found the keyword, move to next row
                }
            }
        }

        console.log('Extracted Helly Hansen data:', data);
        return data;
    }

    /**
     * Compare Buyer CBD data with OB data
     */
    compareWithOB(buyerData) {
        const results = [];

        // Compare each CSV item with found items in BCBD
        for (const csvItem of this.hellyHansenCostData) {
            // Find matching item in buyer data by keyword
            const buyerItem = buyerData.items.find(
                bi => bi.keyword.toLowerCase() === csvItem.item.toLowerCase()
            );

            if (buyerItem) {
                // Special handling for FINANCIAL AND OVERHEAD COST
                if (csvItem.item === 'FINANCIAL AND OVERHEAD COST') {
                    console.log('buyerData.countryOfOrigin:', buyerData.countryOfOrigin);
                    console.log('Comparison result:', buyerData.countryOfOrigin === 'INDO');
                    const expectedValue = buyerData.countryOfOrigin === 'INDO' ? '0.40' : '0.30';
                    console.log('Expected value set to:', expectedValue);
                    results.push({
                        itemName: csvItem.item,
                        obConsm: '-',
                        buyerConsm: '-',
                        consmStatus: 'N/A',
                        obUp: '-',
                        buyerUp: '-',
                        upStatus: 'N/A',
                        obAmount: expectedValue,
                        buyerAmount: buyerItem.amount,
                        amountStatus: this.compareNumericField(expectedValue, buyerItem.amount),
                        specialCase: 'FINANCIAL_OVERHEAD',
                        countryOfOrigin: buyerData.countryOfOrigin
                    });
                }
                // Special handling for MARGIN / PROFIT
                else if (csvItem.item === 'MARGIN / PROFIT') {
                    results.push({
                        itemName: csvItem.item,
                        obConsm: '-',
                        buyerConsm: '-',
                        consmStatus: 'N/A',
                        obUp: '-',
                        buyerUp: '-',
                        upStatus: 'N/A',
                        obAmount: csvItem.amount,
                        buyerAmount: buyerItem.amount,
                        amountStatus: this.validateMarginProfitRange(buyerItem.amount),
                        specialCase: 'MARGIN_PROFIT'
                    });
                }
                // Special handling for Local transportation / documentation
                else if (csvItem.item === 'Local transportation / documentation') {
                    // For this special row, the expected value is in consm field (second column in CSV)
                    const expectedValue = csvItem.consm || '0.25';
                    results.push({
                        itemName: csvItem.item,
                        obConsm: '-',
                        buyerConsm: '-',
                        consmStatus: 'N/A',
                        obUp: '-',
                        buyerUp: '-',
                        upStatus: 'N/A',
                        obAmount: expectedValue,
                        buyerAmount: buyerItem.amount,
                        amountStatus: this.compareNumericField(expectedValue, buyerItem.amount),
                        specialCase: 'LOCAL_TRANSPORT'
                    });
                }
                // Regular items
                else {
                    results.push({
                        itemName: csvItem.item,
                        obConsm: csvItem.consm,
                        buyerConsm: buyerItem.consm,
                        consmStatus: this.compareNumericField(csvItem.consm, buyerItem.consm),
                        obUp: csvItem.up,
                        buyerUp: buyerItem.up,
                        upStatus: this.compareNumericField(csvItem.up, buyerItem.up),
                        obAmount: csvItem.amount,
                        buyerAmount: buyerItem.amount,
                        amountStatus: this.compareNumericField(csvItem.amount, buyerItem.amount)
                    });
                }
            } else {
                // Item not found - check if it's a special row
                const isSpecial = csvItem.item === 'FINANCIAL AND OVERHEAD COST' ||
                                 csvItem.item === 'MARGIN / PROFIT' ||
                                 csvItem.item === 'Local transportation / documentation';

                results.push({
                    itemName: csvItem.item,
                    obConsm: isSpecial ? '-' : csvItem.consm,
                    buyerConsm: 'NOT FOUND',
                    consmStatus: isSpecial ? 'N/A' : 'INVALID',
                    obUp: isSpecial ? '-' : csvItem.up,
                    buyerUp: 'NOT FOUND',
                    upStatus: isSpecial ? 'N/A' : 'INVALID',
                    obAmount: csvItem.amount,
                    buyerAmount: 'NOT FOUND',
                    amountStatus: 'INVALID'
                });
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

        // Round both values to 4 decimal places for comparison
        const obRounded = parseFloat(obNum.toFixed(4));
        const buyerRounded = parseFloat(buyerNum.toFixed(4));

        if (obRounded === buyerRounded) {
            return 'VALID';
        }

        return 'INVALID';
    }

    /**
     * Validate MARGIN / PROFIT range (0.45 to 0.55)
     */
    validateMarginProfitRange(value) {
        const cleanValue = (value || '').toString().replace(/[$,\s%]/g, '');
        const numValue = parseFloat(cleanValue);

        if (isNaN(numValue)) {
            return 'INVALID';
        }

        // Check if value is within the range 0.45 to 0.55 (inclusive)
        if (numValue >= 0.45 && numValue <= 0.55) {
            return 'VALID';
        }

        return 'INVALID';
    }

    /**
     * Format field value with color coding
     */
    formatFieldValue(obValue, buyerValue, status, specialCase = null, item = null) {
        // Handle N/A status (for fields that don't apply to special rows)
        if (status === 'N/A') {
            return `<span style="color: #6b7280; font-weight: 500;">-</span>`;
        }

        const isValid = status === 'VALID';
        const color = isValid ? '#065f46' : '#991b1b';
        const displayValue = buyerValue || 'Empty';

        // Special handling for FINANCIAL AND OVERHEAD COST to always show country of origin
        if (specialCase === 'FINANCIAL_OVERHEAD') {
            console.log('FINANCIAL_OVERHEAD item:', item);
            if (item && item.countryOfOrigin) {
                const expectedText = `${obValue} (${item.countryOfOrigin})`;
                return `<span style="color: ${color}; font-weight: 600;">${displayValue}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${expectedText}</span>`;
            } else {
                console.log('No country of origin found in item');
                const expectedText = obValue;
                return `<span style="color: ${color}; font-weight: 600;">${displayValue}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${expectedText}</span>`;
            }
        }

        // Special handling for LOCAL_TRANSPORT to always show expected
        if (specialCase === 'LOCAL_TRANSPORT') {
            if (isValid) {
                return `<span style="color: ${color}; font-weight: 600;">${displayValue}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${obValue}</span>`;
            } else {
                return `<span style="color: ${color}; font-weight: 600;">${displayValue}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${obValue}</span>`;
            }
        }

        // Special handling for MARGIN / PROFIT to always show expected range
        if (specialCase === 'MARGIN_PROFIT') {
            const expectedText = '0.45 to 0.55';
            return `<span style="color: ${color}; font-weight: 600;">${displayValue}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${expectedText}</span>`;
        }

        if (isValid) {
            return `<span style="color: ${color}; font-weight: 600;">${displayValue}</span>`;
        } else {
            return `<span style="color: ${color}; font-weight: 600;">${displayValue}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${obValue}</span>`;
        }
    }

    /**
     * Generate HTML for results display
     */
    generateResultsHTML(results) {
        if (!results || results.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">Helly Hansen Cost Breakdown Loaded</p>
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
                r.consmStatus === 'VALID' &&
                r.upStatus === 'VALID' &&
                r.amountStatus === 'VALID'
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
                            <th>CONSM</th>
                            <th>U/P</th>
                            <th>Amount</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            for (const item of fileResult.results) {
                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">${item.itemName}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(item.obConsm, item.buyerConsm, item.consmStatus, item.specialCase, item)}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(item.obUp, item.buyerUp, item.upStatus, item.specialCase, item)}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(item.obAmount, item.buyerAmount, item.amountStatus, item.specialCase, item)}</td>
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
window.hellyHansenProcessor = new HellyHansenProcessor();

// Auto-load when V4 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    const v4Tab = document.querySelector('[data-tab="v4"]');
    if (v4Tab) {
        v4Tab.addEventListener('click', () => {
            if (!window.hellyHansenProcessor.hellyHansenCostData) {
                window.hellyHansenProcessor.initialize();
            }
        });
    }

    const v4TabContent = document.getElementById('tab-v4');
    if (v4TabContent && v4TabContent.classList.contains('active')) {
        window.hellyHansenProcessor.initialize();
    }
});
