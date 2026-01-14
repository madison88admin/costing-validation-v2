/**
 * Excel V1 Processing Logic
 * Compares product names between OB and BCBD files with cell value validations
 */

class ExcelV1Processor {
    /**
     * Process all files and generate results
     */
    async processFiles(obFiles, bcbdFiles) {

        try {
            // Extract all product IDs and cell values from BCBD files
            const products = [];
            for (const file of bcbdFiles) {
                const productData = await this.extractProductData(file);
                if (productData.productID) {
                    products.push({
                        id: productData.productID,
                        fileName: file.name,
                        cellValues: productData.cellValues
                    });
                }
            }

            if (products.length === 0) {
                return this.generateErrorHTML('Could not find any product IDs in the BCBD files');
            }

            // Search each product in all OB files
            const allResults = [];
            for (const product of products) {
                for (const obFile of obFiles) {
                    const searchResults = await this.searchProductInWorkbook(obFile, product.id);
                    allResults.push({
                        tnfFileName: obFile.name,
                        productID: product.id,
                        productFileName: product.fileName,
                        found: searchResults.foundLocations.length > 0,
                        locations: searchResults.foundLocations,
                        cellValues: product.cellValues
                    });
                }
            }

            return this.generateResultsHTML(allResults);

        } catch (error) {
            return this.generateErrorHTML(error.message);
        }
    }

    /**
     * Extract Product ID and cell values from BCBD file
     */
    async extractProductData(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    // Get first sheet
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

                    let productID = null;

                    // First, try to extract from cell E14 (common location for Product ID)
                    if (worksheet['E14']) {
                        const e14Value = String(worksheet['E14'].v).trim();
                        const e14Match = e14Value.match(/^([A-Z]{1,2}\d[A-Z0-9]{3,8})/);
                        if (e14Match) {
                            productID = e14Match[1];
                        }
                    }

                    // If not found in E14, try to extract from filename
                    if (!productID) {
                        const fileName = file.name.replace(/\.[^/.]+$/, '');
                        const fileNameMatch = fileName.match(/^([A-Z]{1,2}\d[A-Z0-9]{3,8})/);
                        if (fileNameMatch) {
                            productID = fileNameMatch[1];
                        }
                    }

                    // If still not found, search for "Style #" or similar pattern in the sheet
                    if (!productID) {
                        for (let i = 0; i < Math.min(50, jsonData.length); i++) {
                            const row = jsonData[i];
                            for (let j = 0; j < row.length; j++) {
                                const cellValue = String(row[j]).trim();

                                // Look for "Style #" or "Style No." label
                                if ((cellValue.toLowerCase().includes('style') && cellValue.includes('#')) ||
                                    cellValue.toLowerCase().includes('style no')) {
                                    if (j + 1 < row.length && row[j + 1]) {
                                        const nextCell = String(row[j + 1]).trim();
                                        const match = nextCell.match(/^([A-Z]{1,2}\d[A-Z0-9]{3,8})/);
                                        if (match) {
                                            productID = match[1];
                                            break;
                                        }
                                    }
                                }

                                // Also check if cell matches pattern directly
                                if (/^[A-Z]{1,2}\d[A-Z0-9]{3,8}/.test(cellValue)) {
                                    const match = cellValue.match(/^([A-Z]{1,2}\d[A-Z0-9]{3,8})/);
                                    if (match) {
                                        productID = match[1];
                                    }
                                }
                            }
                            if (productID) break;
                        }
                    }

                    // Now extract cell values from specific cells in the BCBD file
                    const cellValues = {
                        standardMinuteValue: null,
                        averageEfficiency: null,
                        hourlyWages: null,
                        overheadCost: null,
                        factoryProfit: null
                    };

                    // Helper function to extract numeric value from cell
                    const extractValue = (cellRef) => {
                        if (!worksheet[cellRef]) return null;

                        let value = worksheet[cellRef].v;
                        if (typeof value === 'number') return value;

                        if (typeof value === 'string') {
                            let cleaned = value.replace(/[$,\s]/g, '');
                            let percentMatch = cleaned.match(/([\d.]+)%/);
                            if (percentMatch) return parseFloat(percentMatch[1]);
                            let numberMatch = cleaned.match(/([\d.]+)/);
                            if (numberMatch) return parseFloat(numberMatch[1]);
                        }
                        return null;
                    };

                    // Extract values from specific cells in the BCBD file
                    cellValues.standardMinuteValue = extractValue('K7');
                    cellValues.averageEfficiency = extractValue('K8');
                    cellValues.hourlyWages = extractValue('K9');
                    cellValues.overheadCost = extractValue('K11');
                    cellValues.factoryProfit = extractValue('R5');

                    resolve({ productID, cellValues });
                } catch (error) {
                    reject(new Error(`Failed to parse BCBD file: ${error.message}`));
                }
            };

            reader.onerror = () => {
                reject(new Error('Failed to read BCBD file'));
            };

            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Search for product ID across all sheets in the OB workbook and find Total SMV
     */
    async searchProductInWorkbook(file, productID) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    let foundLocations = [];

                    // Search through each sheet for the product ID and its Total SMV
                    workbook.SheetNames.forEach((sheetName) => {
                        const worksheet = workbook.Sheets[sheetName];
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

                        // Search for product ID in this sheet
                        for (let rowIndex = 0; rowIndex < jsonData.length; rowIndex++) {
                            const row = jsonData[rowIndex];
                            for (let colIndex = 0; colIndex < row.length; colIndex++) {
                                const cellValue = String(row[colIndex]).trim();

                                // Check if cell contains the product ID
                                if (cellValue === productID || cellValue.includes(productID)) {
                                    const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });

                                    // Find the SMV value for THIS specific occurrence
                                    let smvForThisOccurrence = null;

                                    // Search in the next 20 rows for "Total SMV"
                                    for (let searchRow = rowIndex; searchRow < Math.min(rowIndex + 20, jsonData.length); searchRow++) {
                                        const searchRowData = jsonData[searchRow];
                                        for (let searchCol = 0; searchCol < searchRowData.length; searchCol++) {
                                            const searchCellValue = String(searchRowData[searchCol]).trim().toLowerCase();

                                            // Look for "Total SMV" label
                                            if (searchCellValue.includes('total smv')) {
                                                // Check the next few cells in the same row for the numeric value
                                                for (let valueCol = searchCol + 1; valueCol < Math.min(searchCol + 5, searchRowData.length); valueCol++) {
                                                    const smvCellRef = XLSX.utils.encode_cell({ r: searchRow, c: valueCol });
                                                    if (worksheet[smvCellRef]) {
                                                        let smvValue = worksheet[smvCellRef].v;
                                                        if (typeof smvValue === 'number' && smvValue > 0) {
                                                            smvForThisOccurrence = smvValue;
                                                            break;
                                                        }
                                                    }
                                                }
                                                if (smvForThisOccurrence !== null) break;
                                            }
                                        }
                                        if (smvForThisOccurrence !== null) break;
                                    }

                                    foundLocations.push({
                                        sheet: sheetName,
                                        cell: cellAddress,
                                        row: rowIndex + 1,
                                        col: colIndex + 1,
                                        value: cellValue,
                                        smv: smvForThisOccurrence
                                    });
                                }
                            }
                        }
                    });

                    resolve({ foundLocations });
                } catch (error) {
                    reject(new Error(`Failed to search workbook: ${error.message}`));
                }
            };

            reader.onerror = () => {
                reject(new Error('Failed to read OB file'));
            };

            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Generate HTML for results display
     */
    generateResultsHTML(results) {
        if (results.length === 0) {
            return `
                <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                    <p style="font-size: 1.3em; margin-bottom: 10px;">❌ No Results</p>
                    <p>No BCBD or OB files to search</p>
                </div>
            `;
        }

        // Group results by product ID to check if product was found anywhere
        const productFoundStatus = {};
        results.forEach(result => {
            if (!productFoundStatus[result.productID]) {
                productFoundStatus[result.productID] = false;
            }
            if (result.found) {
                productFoundStatus[result.productID] = true;
            }
        });

        // Build table with export button
        let tableHTML = `
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
                        placeholder="Search OB or Buyer CBD files..."
                        oninput="window.excelV1Processor.searchTable(this.value)"
                    />
                </div>
                <div style="display: flex; gap: 12px;">
                    <button onclick="window.excelV1Processor.clearFilters()" class="clear-filters-btn">
                        Clear Filters
                    </button>
                    <button onclick="window.excelV1Processor.exportToPDF()" class="export-btn">
                        Export
                    </button>
                </div>
            </div>
            <table id="v1ResultsTable" class="results-table">
                <thead>
                    <tr class="header-labels-row">
                        <th>OB File/s</th>
                        <th>Buyer CBD File/s</th>
                        <th>Match Status with excel</th>
                        <th>Standard Minute Value</th>
                        <th>Average Efficiency %</th>
                        <th>Hourly Wages with Fringes</th>
                        <th>Overhead Cost Ratio to Direct Labor</th>
                        <th>Factory Profit %</th>
                    </tr>
                    <tr class="filter-row">
                        <th></th>
                        <th></th>
                        <th>
                            <select class="column-filter" data-column="2" onchange="window.excelV1Processor.filterTable()">
                                <option value="all">All</option>
                                <option value="found">Found</option>
                                <option value="not-found">Not Found</option>
                            </select>
                        </th>
                        <th>
                            <select class="column-filter" data-column="3" onchange="window.excelV1Processor.filterTable()">
                                <option value="all">All</option>
                                <option value="exact">Exact match</option>
                                <option value="close">Close match</option>
                                <option value="mismatch">Mismatch</option>
                            </select>
                        </th>
                        <th>
                            <select class="column-filter" data-column="4" onchange="window.excelV1Processor.filterTable()">
                                <option value="all">All</option>
                                <option value="valid">Valid</option>
                                <option value="invalid">Invalid</option>
                            </select>
                        </th>
                        <th>
                            <select class="column-filter" data-column="5" onchange="window.excelV1Processor.filterTable()">
                                <option value="all">All</option>
                                <option value="valid">Valid</option>
                                <option value="invalid">Invalid</option>
                            </select>
                        </th>
                        <th>
                            <select class="column-filter" data-column="6" onchange="window.excelV1Processor.filterTable()">
                                <option value="all">All</option>
                                <option value="valid">Valid</option>
                                <option value="invalid">Invalid</option>
                            </select>
                        </th>
                        <th>
                            <select class="column-filter" data-column="7" onchange="window.excelV1Processor.filterTable()">
                                <option value="all">All</option>
                                <option value="valid">Valid</option>
                                <option value="invalid">Invalid</option>
                            </select>
                        </th>
                    </tr>
                </thead>
                <tbody>
        `;

        // Helper function to format and validate cell values
        const formatCellValue = (value, expectedValue, type) => {
            if (value === null || value === undefined) {
                if (type === 'percentage') {
                    return `<span style="color: #991b1b; font-weight: 600;">Cell Empty</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${expectedValue}%</span>`;
                } else {
                    return `<span style="color: #991b1b; font-weight: 600;">Cell Empty</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${expectedValue.toFixed(3)}</span>`;
                }
            }

            let numValue = parseFloat(value);
            let displayValue;
            let isValid;

            if (type === 'percentage') {
                if (typeof value === 'string') {
                    numValue = parseFloat(value.replace('%', ''));
                }

                if (numValue < 1) {
                    numValue = numValue * 100;
                }

                displayValue = numValue.toFixed(1) + '%';
                isValid = Math.abs(numValue - expectedValue) < 0.1;
            } else {
                displayValue = numValue.toFixed(3);
                isValid = Math.abs(numValue - expectedValue) < 0.01;
            }

            const color = isValid ? '#065f46' : '#991b1b';
            const expectedDisplay = type === 'percentage' ? `${expectedValue}%` : expectedValue.toFixed(3);

            if (isValid) {
                return `<span style="color: ${color}; font-weight: 600;">${displayValue}</span>`;
            }

            return `<span style="color: ${color}; font-weight: 600;">${displayValue}</span><br><span style="font-size: 0.85em; color: #849bba;">Expected: ${expectedDisplay}</span>`;
        };

        // Helper function to compare Standard Minute Values
        const formatSMVComparison = (productSMV, tnfSMV) => {
            const truncateToThreeDecimals = (num) => {
                return Math.floor(num * 1000) / 1000;
            };

            const formatThreeDecimals = (num) => {
                const truncated = truncateToThreeDecimals(num);
                const str = truncated.toString();
                const parts = str.split('.');
                if (parts.length === 1) {
                    return str + '.000';
                } else {
                    const decimals = parts[1].padEnd(3, '0');
                    return parts[0] + '.' + decimals;
                }
            };

            if (productSMV === null || productSMV === undefined) {
                return `<span style="color: #991b1b; font-weight: 600;">Product: Empty</span>`;
            }
            if (tnfSMV === null || tnfSMV === undefined) {
                const formattedProduct = formatThreeDecimals(productSMV);
                return `<span style="color: #991b1b; font-weight: 600;">TNF: Empty</span><br><span style="font-size: 0.85em; color: #849bba;">Product: ${formattedProduct}</span>`;
            }

            const truncatedProduct = truncateToThreeDecimals(productSMV);
            const truncatedTNF = truncateToThreeDecimals(tnfSMV);

            const difference = truncateToThreeDecimals(truncatedProduct - truncatedTNF);
            const absDifference = Math.abs(difference);

            let color;
            if (absDifference < 0.001) {
                color = '#065f46';
            } else if (absDifference <= 0.01) {
                color = '#d97706';
            } else {
                color = '#991b1b';
            }

            const diffSign = difference > 0 ? '+' : '';
            const formattedProduct = formatThreeDecimals(productSMV);
            const formattedTNF = formatThreeDecimals(tnfSMV);
            const formattedDiff = formatThreeDecimals(Math.abs(difference));

            if (absDifference < 0.001) {
                return `<span style="color: ${color}; font-weight: 600;">${formattedProduct}</span>`;
            } else {
                return `<span style="color: ${color}; font-weight: 600;">BCBD: ${formattedProduct}</span><br><span style="font-size: 0.85em; color: #849bba;">OB Total SMV: ${formattedTNF} (${diffSign}${formattedDiff})</span>`;
            }
        };

        // First, show all FOUND results
        results.forEach((result) => {
            if (result.found) {
                result.locations.forEach((location) => {
                    tableHTML += `
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem;">
                                <strong style="color: #2b4a6c;">${result.tnfFileName}</strong>
                                <div style="font-size: 0.85em; color: #7a92ab; margin-top: 0.25rem;">
                                    Sheet: ${location.sheet}<br>
                                    Cell: ${location.cell} (Row ${location.row}, Col ${location.col})
                                </div>
                            </td>
                            <td style="padding: 0.875rem 1rem;">
                                <strong style="color: #2b4a6c;">${result.productID}</strong>
                                <div style="font-size: 0.85em; color: #7a92ab; margin-top: 0.25rem;">
                                    From: ${result.productFileName}
                                </div>
                            </td>
                            <td style="padding: 0.875rem 1rem;">
                                <span style="display: inline-flex; align-items: center; gap: 6px; background-color: #d1fae5; color: #065f46; padding: 0.4rem 0.8rem; border-radius: 6px; font-weight: 600; font-size: 0.8rem;">
                                    <span style="font-size: 1em;">✓</span>
                                    <span>FOUND</span>
                                </span>
                            </td>
                            <td style="padding: 0.875rem 1rem;">${formatSMVComparison(result.cellValues.standardMinuteValue, location.smv)}</td>
                            <td style="padding: 0.875rem 1rem;">${formatCellValue(result.cellValues.averageEfficiency, 50, 'percentage')}</td>
                            <td style="padding: 0.875rem 1rem;">${formatCellValue(result.cellValues.hourlyWages, 1.750, 'number')}</td>
                            <td style="padding: 0.875rem 1rem;">${formatCellValue(result.cellValues.overheadCost, 70, 'percentage')}</td>
                            <td style="padding: 0.875rem 1rem;">${formatCellValue(result.cellValues.factoryProfit, 10, 'percentage')}</td>
                        </tr>
                    `;
                });
            }
        });

        // Then, show NOT FOUND only for products that weren't found in ANY file
        const notFoundProducts = new Set();
        results.forEach((result) => {
            if (!result.found && !productFoundStatus[result.productID]) {
                if (!notFoundProducts.has(result.productID)) {
                    notFoundProducts.add(result.productID);
                    tableHTML += `
                        <tr style="border-bottom: 1px solid #e0e8f0;">
                            <td style="padding: 0.875rem 1rem;">
                                <em style="color: #7a92ab;">Searched in all files</em>
                            </td>
                            <td style="padding: 0.875rem 1rem;">
                                <strong style="color: #2b4a6c;">${result.productID}</strong>
                                <div style="font-size: 0.85em; color: #7a92ab; margin-top: 0.25rem;">
                                    From: ${result.productFileName}
                                </div>
                            </td>
                            <td style="padding: 0.875rem 1rem;">
                                <span style="display: inline-flex; align-items: center; gap: 6px; background-color: #fee2e2; color: #991b1b; padding: 0.4rem 0.8rem; border-radius: 6px; font-weight: 600; font-size: 0.8rem;">
                                    <span style="font-size: 1em;">✗</span>
                                    <span>NOT FOUND</span>
                                </span>
                            </td>
                            <td colspan="5" style="text-align: center; color: #849bba; padding: 0.875rem 1rem;">-</td>
                        </tr>
                    `;
                }
            }
        });

        tableHTML += `
                </tbody>
            </table>
        `;

        // Add summary
        const matchedResults = results.filter(r => r.found);
        const totalProducts = new Set(results.map(r => r.productID)).size;
        const matchedProducts = new Set(matchedResults.map(r => r.productID)).size;
        const notFoundCount = notFoundProducts.size;
        const totalTNFFiles = new Set(results.map(r => r.tnfFileName)).size;

        const summaryHTML = `
            <div style="margin-bottom: 20px; padding: 15px; background: #f0f7ff; border-radius: 10px; border-left: 4px solid #3b82f6;">
                <strong>Summary:</strong> Found ${matchedProducts} of ${totalProducts} products across ${totalTNFFiles} OB file(s). 
                ${notFoundCount > 0 ? `<span style="color: #991b1b;">${notFoundCount} product(s) not found in any file.</span>` : ''}
                Total matches: ${matchedResults.reduce((sum, r) => sum + r.locations.length, 0)}
            </div>
        `;

        return summaryHTML + tableHTML;
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
                <p style="color: #721c24; font-size: 0.85rem; margin-top: 0.5rem;">
                    Please make sure you have uploaded valid Excel files and that the SheetJS library is loaded.
                </p>
            </div>
        `;
    }

    /**
     * Export results to PDF using the unified Export.js module
     */
    async exportToPDF() {
        if (!window.pdfExporter) {
            alert('PDF export module not available. Please refresh the page.');
            return;
        }
        await window.pdfExporter.exportToPDF(window.pdfExporter.createTNFConfig());
    }

    /**
     * Get current filter values from dropdowns
     */
    getFilterValues() {
        const table = document.getElementById('v1ResultsTable');
        if (!table) return {};

        const filters = {};
        const filterSelects = table.querySelectorAll('.column-filter');
        filterSelects.forEach(select => {
            filters[select.getAttribute('data-column')] = select.value;
        });
        return filters;
    }

    /**
     * Check if a row passes all filter criteria
     */
    rowPassesFilters(row, filters) {
        const cells = row.querySelectorAll('td');

        for (const [columnIndex, filterValue] of Object.entries(filters)) {
            if (filterValue === 'all') continue;

            const cell = cells[columnIndex];
            if (!cell) continue;

            const cellText = cell.textContent.trim();

            // Column 2: Match Status
            if (columnIndex === '2') {
                if (filterValue === 'found' && !cellText.includes('✓ FOUND')) return false;
                if (filterValue === 'not-found' && !cellText.includes('✗ NOT FOUND')) return false;
            }

            // Column 3: Standard Minute Value
            if (columnIndex === '3' && cellText !== '-') {
                if (filterValue === 'exact' && (cellText.includes('BCBD:') || cellText.includes('Empty'))) return false;
                if (filterValue === 'close') {
                    if (!cellText.includes('BCBD:')) return false;
                    const diffMatch = cellText.match(/\([\+\-]([\d.]+)\)/);
                    if (diffMatch && parseFloat(diffMatch[1]) > 0.01) return false;
                }
                if (filterValue === 'mismatch') {
                    if (!cellText.includes('Empty') && !cellText.includes('BCBD:')) return false;
                    if (cellText.includes('BCBD:')) {
                        const diffMatch = cellText.match(/\([\+\-]([\d.]+)\)/);
                        if (diffMatch && parseFloat(diffMatch[1]) <= 0.01) return false;
                    }
                }
            }

            // Columns 4-7: Valid/Invalid filters
            if (['4', '5', '6', '7'].includes(columnIndex) && cellText !== '-') {
                if (filterValue === 'valid' && (cellText.includes('Cell Empty') || cellText.includes('Expected:'))) return false;
                if (filterValue === 'invalid' && !cellText.includes('Cell Empty') && !cellText.includes('Expected:')) return false;
            }
        }
        return true;
    }

    /**
     * Filter table based on dropdown selections
     */
    filterTable() {
        const table = document.getElementById('v1ResultsTable');
        if (!table) return;

        const filters = this.getFilterValues();
        const rows = table.querySelector('tbody').querySelectorAll('tr');

        rows.forEach(row => {
            row.style.display = this.rowPassesFilters(row, filters) ? '' : 'none';
        });
    }

    /**
     * Clear all filters and show all rows
     */
    clearFilters() {
        const table = document.getElementById('v1ResultsTable');
        if (!table) return;

        table.querySelectorAll('.column-filter').forEach(select => select.value = 'all');

        const searchInput = document.querySelector('.search-input-expandable');
        if (searchInput) searchInput.value = '';

        table.querySelector('tbody').querySelectorAll('tr').forEach(row => row.style.display = '');
    }

    /**
     * Search table rows based on OB Files and Buyer CBD Files columns
     */
    searchTable(searchTerm) {
        const table = document.getElementById('v1ResultsTable');
        if (!table) return;

        const searchLower = searchTerm.toLowerCase().trim();

        if (searchLower === '') {
            this.filterTable();
            return;
        }

        const filters = this.getFilterValues();
        const rows = table.querySelector('tbody').querySelectorAll('tr');

        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            const obFileText = cells[0] ? cells[0].textContent.toLowerCase() : '';
            const buyerCbdText = cells[1] ? cells[1].textContent.toLowerCase() : '';
            const matchFound = obFileText.includes(searchLower) || buyerCbdText.includes(searchLower);

            row.style.display = (matchFound && this.rowPassesFilters(row, filters)) ? '' : 'none';
        });
    }
}

// Initialize the processor
window.excelV1Processor = new ExcelV1Processor();