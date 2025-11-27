/**
 * Excel V2 Processing Logic
 * Automatically loads Burton_CostBreakdown.csv from public folder
 */

class ExcelV2Processor {
    constructor() {
        this.burtonCostData = null;
        this.bcbdResults = [];
    }

    /**
     * Initialize V2 - Load Burton Cost Breakdown CSV automatically
     */
    async initialize() {
        try {
            // Fetch the Burton_CostBreakdown.csv file from public folder
            const response = await fetch('../public/Burton_CostBreakdown.csv');
            if (!response.ok) {
                throw new Error('Failed to load Burton_CostBreakdown.csv');
            }
            
            const csvText = await response.text();
            this.burtonCostData = this.parseCSV(csvText);
            
            // Display the loaded data in the OB drop zone
            this.displayBurtonCostData();
            
            console.log('Burton Cost Breakdown loaded successfully:', this.burtonCostData);
        } catch (error) {
            console.error('Error loading Burton Cost Breakdown:', error);
            this.displayError('Failed to load Burton_CostBreakdown.csv from public folder');
        }
    }

    /**
     * Parse CSV text into array of objects
     * Handles commas within the description field
     */
    parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        const data = [];
        
        lines.forEach(line => {
            // Split by comma
            const values = line.split(',').map(val => val.trim());
            
            // We expect 9 fields total
            // If we have more than 9, the description contains commas
            let description = '';
            let startIndex = 0;
            
            if (values.length > 9) {
                // Combine the first (length - 8) values as description
                const descParts = values.length - 8;
                description = values.slice(0, descParts).join(', ');
                startIndex = descParts;
            } else {
                description = values[0] || '';
                startIndex = 1;
            }
            
            data.push({
                description: description,
                details: values[startIndex] || '',
                materialName: values[startIndex + 1] || '',
                supplier: values[startIndex + 2] || '',
                quantity: values[startIndex + 3] || '',
                wastage: values[startIndex + 4] || '',
                unit: values[startIndex + 5] || '',
                unitPrice: values[startIndex + 6] || '',
                totalPrice: values[startIndex + 7] || ''
            });
        });
        
        return data;
    }

    /**
     * Display Burton Cost Breakdown data in the OB drop zone
     */
    displayBurtonCostData() {
        const obDropZone = document.getElementById('obDropZone-v2');
        if (!obDropZone) return;

        // Replace the drop zone content with the Burton Cost data display
        let contentHTML = `
            <div class="burton-cost-container">
                <div class="burton-cost-items">
        `;

        // Display each line from the CSV
        this.burtonCostData.forEach((item, index) => {
            contentHTML += `
                <div class="burton-cost-item">
                    <div class="burton-item-line"><strong>${item.description}</strong></div>
                    ${item.details ? `<div class="burton-item-line"><strong>Details:</strong> ${item.details}</div>` : ''}
                    <div class="burton-item-line"><strong>Material:</strong> ${item.materialName}</div>
                    <div class="burton-item-line"><strong>Supplier:</strong> ${item.supplier}</div>
                    <div class="burton-item-line"><strong>Qty:</strong> ${item.quantity}</div>
                    <div class="burton-item-line"><strong>Wastage:</strong> ${this.formatToThreeDecimals(item.wastage)}</div>
                    <div class="burton-item-line"><strong>Unit:</strong> ${item.unit}</div>
                    <div class="burton-item-line"><strong>Unit Price:</strong> ${this.formatToThreeDecimals(item.unitPrice)}</div>
                    <div class="burton-item-line"><strong>Total:</strong> ${this.formatToThreeDecimals(item.totalPrice)}</div>
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
        const obDropZone = document.getElementById('obDropZone-v2');
        if (!obDropZone) return;

        obDropZone.innerHTML = `
            <div class="drop-zone-content">
                <div style="background: #fee; border-left: 4px solid #dc3545; padding: 1.5rem; border-radius: 8px;">
                    <p style="color: #dc3545; font-weight: 600; margin-bottom: 0.5rem;">
                        ❌ Error Loading File
                    </p>
                    <p style="color: #721c24; font-size: 0.95rem;">
                        ${errorMessage}
                    </p>
                </div>
            </div>
        `;
    }

    /**
     * Process files and generate results
     */
    async processFiles(bcbdFiles) {
        this.bcbdResults = [];

        try {
            if (!this.burtonCostData || this.burtonCostData.length === 0) {
                return this.generateErrorHTML('Burton Cost Breakdown data not loaded');
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
                    
                    // Get the last sheet (usually contains the latest data)
                    const lastSheetName = workbook.SheetNames[workbook.SheetNames.length - 1];
                    const sheet = workbook.Sheets[lastSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
                    
                    // Parse the trims section
                    const trimsData = this.extractTrimsData(jsonData);
                    resolve(trimsData);
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Extract all data from the parsed Excel (scan entire column A)
     */
    extractTrimsData(jsonData) {
        const trimsData = [];
        
        // Scan through all rows in the Excel file
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            
            // Skip empty rows or header-like rows
            if (!row[0] || row[0].toString().trim() === '') {
                continue;
            }
            
            const cellValue = row[0].toString().trim().toUpperCase();
            
            // Skip section headers and totals
            if (cellValue.includes('BURTON') || 
                cellValue.includes('TARGET') ||
                cellValue.includes('FABRIC') ||
                cellValue.includes('TRIMS') ||
                cellValue.includes('ARTWORK') ||
                cellValue.includes('VENDOR') ||
                cellValue.includes('TOTAL') ||
                cellValue.includes('DATE') ||
                cellValue.includes('SEASON') ||
                cellValue.includes('STYLE') ||
                cellValue.includes('COLORS') ||
                cellValue.includes('FACTORY') ||
                cellValue.includes('C.O.O') ||
                cellValue.includes('REF #') ||
                cellValue === 'CM' ||
                cellValue === 'QUOTA' ||
                cellValue === 'FREIGHT' ||
                cellValue === 'DUTY' ||
                cellValue === 'OVERHEAD') {
                continue;
            }
            
            // Extract all potential trim items from column A
            trimsData.push({
                description: row[0] ? row[0].toString().trim() : '',
                details: row[1] ? row[1].toString().trim() : '',
                material: row[2] ? row[2].toString().trim() : '',
                supplier: row[3] ? row[3].toString().trim() : '',
                qty: row[4] ? row[4].toString().trim() : '',
                wastage: row[5] ? row[5].toString().trim() : '',
                unit: row[6] ? row[6].toString().trim() : '',
                unitPrice: row[7] ? row[7].toString().trim() : '',
                total: row[8] ? row[8].toString().trim() : ''
            });
        }
        
        return trimsData;
    }

    /**
     * Compare Buyer CBD data with OB data
     */
    compareWithOB(buyerData) {
        const itemsToCheck = [
            'Sewing Thread - See Vendor Guide',
            'Beanie Care Content Label',
            'White Polyester Taffeta Tracking Label, 38mm wide x 50mm long',
            'VERTICAL RFID UPC STICKER',
            'Glassine Bag',
            'Polybag Sticker',
            'EA- HSC11'
        ];
        
        const results = [];
        
        for (const itemName of itemsToCheck) {
            // Find item in OB data
            const obItem = this.findItemInOB(itemName);
            
            // Find item in Buyer data
            const buyerItem = this.findItemInBuyer(buyerData, itemName);
            
            if (!obItem) {
                results.push({
                    itemName,
                    status: 'NOT_FOUND_IN_OB',
                    comparison: null
                });
                continue;
            }
            
            if (!buyerItem) {
                results.push({
                    itemName,
                    status: 'NOT_FOUND_IN_BUYER',
                    comparison: null
                });
                continue;
            }
            
            // Compare fields
            const comparison = {
                material: this.compareField(obItem.materialName, buyerItem.material),
                supplier: this.compareField(obItem.supplier, buyerItem.supplier),
                qty: this.compareNumericField(obItem.quantity, buyerItem.qty, false),
                wastage: this.compareNumericField(obItem.wastage, buyerItem.wastage, true),
                unit: this.compareField(obItem.unit, buyerItem.unit),
                unitPrice: this.compareNumericField(obItem.unitPrice, buyerItem.unitPrice, true),
                total: this.compareNumericField(obItem.totalPrice, buyerItem.total, true)
            };
            
            // Debug logging
            console.log(`Item: ${itemName}`);
            console.log('OB Data:', obItem);
            console.log('Buyer Data:', buyerItem);
            console.log('Comparison:', comparison);
            
            results.push({
                itemName,
                status: 'FOUND',
                obData: obItem,
                buyerData: buyerItem,
                comparison
            });
        }
        
        return results;
    }

    /**
     * Find item in OB data
     */
    findItemInOB(itemName) {
        const normalizedName = this.normalizeItemName(itemName);
        return this.burtonCostData.find(item => {
            const itemDesc = this.normalizeItemName(item.description);
            return this.fuzzyMatch(normalizedName, itemDesc);
        });
    }

    /**
     * Find item in Buyer data
     */
    findItemInBuyer(buyerData, itemName) {
        const normalizedName = this.normalizeItemName(itemName);
        
        console.log(`\n=== Searching for: "${itemName}" ===`);
        console.log(`Normalized search term: "${normalizedName}"`);
        console.log('Available items in BCBD:');
        buyerData.forEach((item, index) => {
            const itemDesc = this.normalizeItemName(item.description);
            const matches = this.fuzzyMatch(normalizedName, itemDesc);
            console.log(`  [${index}] "${item.description}" -> normalized: "${itemDesc}" -> Match: ${matches}`);
        });
        
        const found = buyerData.find(item => {
            const itemDesc = this.normalizeItemName(item.description);
            return this.fuzzyMatch(normalizedName, itemDesc);
        });
        
        console.log(`Result: ${found ? 'FOUND - ' + found.description : 'NOT FOUND'}`);
        
        return found;
    }

    /**
     * Normalize item name for comparison
     */
    normalizeItemName(name) {
        return name.toLowerCase()
            .trim()
            .replace(/[,;:\-_]+/g, ' ')  // Remove commas, semicolons, colons, hyphens, underscores
            .replace(/\s+/g, ' ')         // Normalize multiple spaces to single space
            .trim();
    }

    /**
     * Fuzzy match two strings - STRICT matching for better accuracy
     */
    fuzzyMatch(str1, str2) {
        // Direct match
        if (str1 === str2) return true;
        
        // For "Sewing Thread" - must contain both "sewing" and "thread"
        if (str1.includes('sewing') && str1.includes('thread')) {
            return str2.includes('sewing') && str2.includes('thread');
        }
        
        // For other items - require at least 80% of significant keywords to match
        const keywords1 = str1.split(' ').filter(w => w.length > 3);
        const keywords2 = str2.split(' ').filter(w => w.length > 3);
        
        if (keywords1.length === 0 || keywords2.length === 0) {
            return false;
        }
        
        // Count exact keyword matches (not partial)
        const matchCount = keywords1.filter(k => keywords2.includes(k)).length;
        
        // Require at least 80% of keywords to match exactly
        return matchCount >= Math.min(keywords1.length, keywords2.length) * 0.8;
    }

    /**
     * Compare text fields
     */
    compareField(obValue, buyerValue) {
        const ob = (obValue || '').toString().toLowerCase().trim();
        const buyer = (buyerValue || '').toString().toLowerCase().trim();
        return ob === buyer ? 'VALID' : 'INVALID';
    }

    /**
     * Compare numeric fields
     * Compares values rounded to 3 decimal places
     * Returns 'VALID', 'WARNING' (minor difference of 0.001), or 'INVALID'
     */
    compareNumericField(obValue, buyerValue, checkMinorDifference = false) {
        // Remove currency symbols and convert to numbers
        const cleanOB = (obValue || '').toString().replace(/[$,\s]/g, '');
        const cleanBuyer = (buyerValue || '').toString().replace(/[$,\s]/g, '');
        
        const obNum = parseFloat(cleanOB);
        const buyerNum = parseFloat(cleanBuyer);
        
        if (isNaN(obNum) || isNaN(buyerNum)) {
            return 'INVALID';
        }
        
        // Round both values to 3 decimal places before comparing
        const obRounded = parseFloat(obNum.toFixed(3));
        const buyerRounded = parseFloat(buyerNum.toFixed(3));
        
        if (obRounded === buyerRounded) {
            return 'VALID';
        }
        
        // Check for minor difference of exactly 0.001 if enabled
        if (checkMinorDifference) {
            const difference = Math.abs(obRounded - buyerRounded);
            if (difference === 0.001) {
                return 'WARNING';
            }
        }
        
        return 'INVALID';
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
     * Format field value with color coding
     * Supports VALID (green), WARNING (yellow), and INVALID (red) statuses
     */
    formatFieldValue(obValue, buyerValue, status, isNumeric = false) {
        // Determine color based on status
        let color;
        if (status === 'VALID') {
            color = '#065f46'; // Green
        } else if (status === 'WARNING') {
            color = '#d97706'; // Yellow/Orange
        } else {
            color = '#991b1b'; // Red
        }
        
        // Format numeric values to 3 decimal places
        const displayOB = isNumeric ? this.formatToThreeDecimals(obValue) : obValue;
        const displayBuyer = isNumeric ? this.formatToThreeDecimals(buyerValue) : buyerValue;
        
        if (!buyerValue || buyerValue === '') {
            return `<span style="color: #991b1b; font-weight: 600;">Empty</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${displayOB}</span>`;
        }
        
        if (status === 'VALID') {
            return `<span style="color: ${color}; font-weight: 600;">${displayBuyer}</span>`;
        } else {
            return `<span style="color: ${color}; font-weight: 600;">${displayBuyer}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${displayOB}</span>`;
        }
    }

    /**
     * Generate HTML for results display
     */
    generateResultsHTML(results) {
        if (!results || results.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">✓ Burton Cost Breakdown Loaded</p>
                    <p>Ready for processing. Upload Buyer CBD files to continue.</p>
                    <p style="margin-top: 15px; font-size: 0.9em; color: #7a92ab;">
                        Loaded ${this.burtonCostData ? this.burtonCostData.length : 0} items from Burton_CostBreakdown.csv
                    </p>
                </div>
            `;
        }

        let html = '';
        
        for (const fileResult of results) {
            // Add summary at the top (like V1)
            const totalItems = fileResult.results.length;
            const validItems = fileResult.results.filter(r => {
                if (r.status !== 'FOUND') return false;
                const comp = r.comparison;
                return Object.values(comp).every(v => v === 'VALID');
            }).length;
            
            html += `
                <div style="margin-bottom: 20px; padding: 15px; background: #f0f7ff; border-radius: 10px; border-left: 4px solid #3b82f6;">
                    <strong>File:</strong> ${fileResult.fileName}<br>
                    <strong>Summary:</strong> ${validItems} out of ${totalItems} items fully match the OB file
                </div>
            `;
            
            // Create comparison table with V1 styling
            html += `
                <table id="v2ResultsTable" class="results-table">
                    <thead>
                        <tr class="header-labels-row">
                            <th>Item Name</th>
                            <th>Material</th>
                            <th>Supplier</th>
                            <th>Qty</th>
                            <th>Wastage</th>
                            <th>Unit</th>
                            <th>Unit Price</th>
                            <th>Total</th>
                        </tr>
                    </thead>
                    <tbody>
            `;
            
            for (const item of fileResult.results) {
                if (item.status === 'NOT_FOUND_IN_OB') {
                    html += `
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">${item.itemName}</td>
                            <td colspan="7" style="text-align: center; color: #991b1b; padding: 0.875rem 1rem;">
                                ⚠️ Not found in OB file
                            </td>
                        </tr>
                    `;
                } else if (item.status === 'NOT_FOUND_IN_BUYER') {
                    html += `
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">${item.itemName}</td>
                            <td colspan="7" style="text-align: center; color: #991b1b; padding: 0.875rem 1rem;">
                                ⚠️ Not found in Buyer CBD file
                            </td>
                        </tr>
                    `;
                } else {
                    const comp = item.comparison;
                    const obData = item.obData;
                    const buyerData = item.buyerData;
                    
                    html += `
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem; font-weight: 600;">${item.itemName}</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(obData.materialName, buyerData.material, comp.material, false)}</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(obData.supplier, buyerData.supplier, comp.supplier, false)}</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(obData.quantity, buyerData.qty, comp.qty, false)}</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(obData.wastage, buyerData.wastage, comp.wastage, true)}</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(obData.unit, buyerData.unit, comp.unit, false)}</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(obData.unitPrice, buyerData.unitPrice, comp.unitPrice, true)}</td>
                            <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(obData.totalPrice, buyerData.total, comp.total, true)}</td>
                        </tr>
                    `;
                }
            }
            
            html += `
                    </tbody>
                </table>
            `;
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
}

// Initialize the processor
window.excelV2Processor = new ExcelV2Processor();

// Auto-load Burton Cost Breakdown when V2 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    // Check if we're on V2 tab and initialize
    const v2Tab = document.querySelector('[data-tab="v2"]');
    if (v2Tab) {
        v2Tab.addEventListener('click', () => {
            if (!window.excelV2Processor.burtonCostData) {
                window.excelV2Processor.initialize();
            }
        });
    }

    // If V2 tab is already active on load, initialize immediately
    const v2TabContent = document.getElementById('tab-v2');
    if (v2TabContent && v2TabContent.classList.contains('active')) {
        window.excelV2Processor.initialize();
    }
});
