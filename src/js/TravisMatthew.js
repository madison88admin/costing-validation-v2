/**
 * Travis Matthew Processing Logic
 * Automatically loads TravisMatthew_CostBreakdown.csv from assets/data folder
 */

class TravisMatthewProcessor {
    constructor() {
        this.travisMatthewCostData = null;
        this.bcbdResults = [];
    }

    /**
     * Initialize - Load Travis Matthew Cost Breakdown CSV automatically
     */
    async initialize() {
        try {
            // Fetch the TravisMatthew_CostBreakdown.csv file from assets/data folder
            const response = await fetch('assets/data/TravisMatthew_CostBreakdown.csv');
            if (!response.ok) {
                throw new Error('Failed to load TravisMatthew_CostBreakdown.csv');
            }

            const csvText = await response.text();
            this.travisMatthewCostData = this.parseCSV(csvText);

            // Display the loaded data in the OB drop zone
            this.displayTravisMatthewCostData();

            console.log('Travis Matthew Cost Breakdown loaded successfully:', this.travisMatthewCostData);
        } catch (error) {
            console.error('Error loading Travis Matthew Cost Breakdown:', error);
            this.displayError('Failed to load TravisMatthew_CostBreakdown.csv from assets/data folder');
        }
    }

    /**
     * Parse CSV text into array of objects
     */
    parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        const data = [];

        lines.forEach(line => {
            // Split by comma
            const values = line.split(',').map(val => val.trim());

            if (values.length >= 2) {
                data.push({
                    label: values[0] || '',
                    value: values[1] || ''
                });
            }
        });

        return data;
    }

    /**
     * Display Travis Matthew Cost Breakdown data in the OB drop zone
     */
    displayTravisMatthewCostData() {
        const obDropZone = document.getElementById('obDropZone-v14');
        if (!obDropZone) return;

        // Replace the drop zone content with the Travis Matthew Cost data display
        let contentHTML = `
            <div class="burton-cost-container">
                <div class="burton-cost-header">
                </div>
                <div class="burton-cost-items">
        `;

        // Display each line from the CSV
        this.travisMatthewCostData.forEach((item, index) => {
            contentHTML += `
                <div class="burton-cost-item">
                    <div class="burton-item-line"><strong>${item.label}:</strong> ${item.value}</div>
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
        const obDropZone = document.getElementById('obDropZone-v14');
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
            if (!this.travisMatthewCostData || this.travisMatthewCostData.length === 0) {
                return this.generateErrorHTML('Travis Matthew Cost Breakdown data not loaded');
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

                    // Extract the data for validation
                    const extractedData = this.extractBuyerData(jsonData);
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
     * Extract buyer data from the parsed Excel
     * Looks for labels in Column A and values in Column B (or Column G for SHIPPING/DUTY/OVERHEAD/PROFIT)
     */
    extractBuyerData(jsonData) {
        const buyerData = {};

        console.log(`📊 Total rows in Excel: ${jsonData.length}`);

        // Define the labels we're looking for (case-insensitive)
        const labelsToFind = [
            'VENDOR',
            'FACTORY',
            'COO',
            'PORT OF EXPORT',
            'SHIPPING/DUTY/OVERHEAD/PROFIT'
        ];

        // Scan through all rows in the Excel file
        for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];

            // Skip empty rows
            if (!row[0] || row[0].toString().trim() === '') {
                continue;
            }

            const cellA = row[0].toString().trim().toUpperCase();

            // Check if this row contains one of our labels
            for (const label of labelsToFind) {
                if (cellA === label || cellA.includes(label)) {
                    // For SHIPPING/DUTY/OVERHEAD/PROFIT, value is in Column G (index 6)
                    if (label === 'SHIPPING/DUTY/OVERHEAD/PROFIT') {
                        const valueG = row[6] ? row[6].toString().trim() : '';
                        buyerData[label] = {
                            value: valueG,
                            row: i + 1, // 1-indexed for display
                            column: 'G'
                        };
                        console.log(`Found ${label} at Row ${i + 1}, Column G: "${valueG}"`);
                    } else {
                        // For other labels, value is in Column B (index 1)
                        const valueB = row[1] ? row[1].toString().trim() : '';
                        buyerData[label] = {
                            value: valueB,
                            row: i + 1, // 1-indexed for display
                            column: 'B'
                        };
                        console.log(`Found ${label} at Row ${i + 1}, Column B: "${valueB}"`);
                    }
                    break;
                }
            }
        }

        return buyerData;
    }

    /**
     * Compare Buyer CBD data with OB data
     */
    compareWithOB(buyerData) {
        const results = [];

        // Map OB labels to buyer labels
        const labelMapping = {
            'VENDOR': 'VENDOR',
            'FACTORY': 'FACTORY',
            'COO': 'COO',
            'PORT OF EXPORT': 'PORT OF EXPORT',
            'SHIPPING/DUTY/OVERHEAD/PROFIT': 'SHIPPING/DUTY/OVERHEAD/PROFIT'
        };

        for (const obItem of this.travisMatthewCostData) {
            const obLabel = obItem.label.toUpperCase();
            const obValue = obItem.value;
            const buyerLabel = labelMapping[obLabel] || obLabel;

            const buyerItem = buyerData[buyerLabel];

            if (!buyerItem) {
                results.push({
                    label: obItem.label,
                    obValue: obValue,
                    buyerValue: null,
                    status: 'NOT_FOUND',
                    location: null,
                    isValid: false
                });
                continue;
            }

            // Compare values
            const isValid = this.compareValues(obValue, buyerItem.value, obLabel);

            results.push({
                label: obItem.label,
                obValue: obValue,
                buyerValue: buyerItem.value,
                status: 'FOUND',
                location: `Row ${buyerItem.row}, Column ${buyerItem.column}`,
                isValid: isValid
            });
        }

        return results;
    }

    /**
     * Compare two values
     */
    compareValues(obValue, buyerValue, label) {
        // Normalize strings for comparison
        const obNormalized = obValue.toString().toLowerCase().trim();
        const buyerNormalized = buyerValue.toString().toLowerCase().trim();

        // For numeric values (like SHIPPING/DUTY/OVERHEAD/PROFIT), compare as numbers
        if (label === 'SHIPPING/DUTY/OVERHEAD/PROFIT') {
            const obNum = parseFloat(obValue);
            const buyerNum = parseFloat(buyerValue);

            if (!isNaN(obNum) && !isNaN(buyerNum)) {
                // Compare with 2 decimal precision
                return Math.abs(obNum - buyerNum) < 0.001;
            }
        }

        // For text values, do case-insensitive comparison
        return obNormalized === buyerNormalized;
    }

    /**
     * Format field value with color coding and expected value display
     */
    formatFieldValue(obValue, buyerValue, isValid) {
        if (buyerValue === null || buyerValue === '') {
            return `<span style="color: #991b1b; font-weight: 600;">Not Found</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${obValue}</span>`;
        }

        if (isValid) {
            return `<span style="color: #065f46; font-weight: 600;">${buyerValue}</span>`;
        } else {
            return `<span style="color: #991b1b; font-weight: 600;">${buyerValue}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${obValue}</span>`;
        }
    }

    /**
     * Generate HTML for results display
     */
    generateResultsHTML(results) {
        if (!results || results.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">✓ Travis Matthew Cost Breakdown Loaded</p>
                    <p>Ready for processing. Upload Buyer CBD files to continue.</p>
                    <p style="margin-top: 15px; font-size: 0.9em; color: #7a92ab;">
                        Loaded ${this.travisMatthewCostData ? this.travisMatthewCostData.length : 0} items from TravisMatthew_CostBreakdown.csv
                    </p>
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
                        oninput="window.travisMatthewProcessor.searchByFilename(this.value)"
                    />
                </div>
                <button onclick="window.travisMatthewProcessor.exportToPDF()" class="export-btn">
                    Export
                </button>
            </div>
        `;

        for (const fileResult of results) {
            // Wrap each file's results in a group container
            html += `<div class="file-result-group">`;

            // Add summary at the top
            const totalItems = fileResult.results.length;
            const validItems = fileResult.results.filter(r => r.isValid).length;

            html += `
                <div class="file-summary-box">
                    <strong>File:</strong> ${fileResult.fileName}<br>
                    <strong>Summary:</strong> ${validItems} out of ${totalItems} items match the OB file
                </div>
            `;

            // Create comparison table
            html += `
                <table id="v14ResultsTable" class="results-table">
                    <thead>
                        <tr class="header-labels-row">
                            <th>Field</th>
                            <th>BCBD Value</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            for (const item of fileResult.results) {
                const statusIcon = item.isValid ? '✓' : '✗';
                const statusColor = item.isValid ? '#065f46' : '#991b1b';
                const statusText = item.isValid ? 'VALID' : 'INVALID';

                html += `
                    <tr style="border-bottom: 1px solid #e0e8f0;">
                        <td style="padding: 0.875rem 1rem; font-weight: 600;">${item.label}</td>
                        <td style="padding: 0.875rem 1rem;">${this.formatFieldValue(item.obValue, item.buyerValue, item.isValid)}</td>
                        <td style="padding: 0.875rem 1rem;">
                            <span style="color: ${statusColor}; font-weight: 600;">${statusIcon} ${statusText}</span>
                        </td>
                    </tr>
                `;
            }

            html += `
                    </tbody>
                </table>
            </div>`;  // Close file-result-group
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

        const config = window.pdfExporter.createTravisMatthewConfig(this.bcbdResults);
        await window.pdfExporter.exportMultiFileToPDF(config);
    }

    /**
     * Search by filename - filters file result groups based on filename
     */
    searchByFilename(searchTerm) {
        const fileGroups = document.querySelectorAll('#tab-v14 .file-result-group');

        if (!fileGroups || fileGroups.length === 0) {
            return;
        }

        // Convert search term to lowercase for case-insensitive search
        const searchLower = searchTerm.toLowerCase().trim();

        // If search is empty, show all groups
        if (searchLower === '') {
            fileGroups.forEach(group => {
                group.style.display = '';
            });
            return;
        }

        // Filter each file group based on filename
        fileGroups.forEach(group => {
            const summaryBox = group.querySelector('.file-summary-box');
            if (!summaryBox) return;

            // Get the full text content
            const fullText = summaryBox.textContent || summaryBox.innerText;

            // Split by line breaks and find the line with "File:"
            const lines = fullText.split(/\r?\n/).map(line => line.trim());
            let filename = '';

            for (const line of lines) {
                if (line.toLowerCase().startsWith('file:')) {
                    // Extract everything after "File:"
                    filename = line.substring(5).trim().toLowerCase();
                    break;
                }
            }

            // Check if filename contains the search term
            if (filename && filename.includes(searchLower)) {
                group.style.display = '';
            } else {
                group.style.display = 'none';
            }
        });
    }
}

// Initialize the processor
window.travisMatthewProcessor = new TravisMatthewProcessor();

// Auto-load Travis Matthew Cost Breakdown when V14 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    // Add click listener to ALL elements with data-tab="v14" (both tab buttons and menu items)
    const v14Tabs = document.querySelectorAll('[data-tab="v14"]');
    v14Tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            if (!window.travisMatthewProcessor.travisMatthewCostData) {
                window.travisMatthewProcessor.initialize();
            }
        });
    });

    // If V14 tab is already active on load, initialize immediately
    const v14TabContent = document.getElementById('tab-v14');
    if (v14TabContent && v14TabContent.classList.contains('active')) {
        window.travisMatthewProcessor.initialize();
    }
});
