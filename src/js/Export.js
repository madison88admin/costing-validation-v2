/**
 * Unified PDF Export Module
 * Supports different templates (TNF, Burton, and future templates) with varying table formats
 */

class PDFExporter {
    constructor() {
        this.jsPDFLoaded = false;
    }

    /**
     * Load jsPDF library dynamically
     */
    async loadJsPDF() {
        return new Promise((resolve, reject) => {
            if (typeof window.jspdf !== 'undefined') {
                this.jsPDFLoaded = true;
                resolve();
                return;
            }

            const jsPDFScript = document.createElement('script');
            jsPDFScript.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
            jsPDFScript.onload = () => {
                const autoTableScript = document.createElement('script');
                autoTableScript.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js';
                autoTableScript.onload = () => {
                    console.log('jsPDF and autoTable loaded successfully');
                    this.jsPDFLoaded = true;
                    resolve();
                };
                autoTableScript.onerror = () => {
                    reject(new Error('Failed to load jsPDF autoTable plugin'));
                };
                document.head.appendChild(autoTableScript);
            };
            jsPDFScript.onerror = () => {
                reject(new Error('Failed to load jsPDF library'));
            };
            document.head.appendChild(jsPDFScript);
        });
    }

    /**
     * Export to PDF with template configuration
     * @param {Object} config - Export configuration
     * @param {string} config.title - PDF title
     * @param {string} config.tableId - HTML table ID to export
     * @param {string} config.summarySelector - CSS selector for summary element
     * @param {string} config.filenamePrefix - Prefix for the exported filename
     * @param {Array} config.columnWidths - Array of column widths in mm
     * @param {Object} config.colorRules - Color coding rules for cells
     */
    async exportToPDF(config) {
        const {
            title,
            tableId,
            summarySelector,
            filenamePrefix,
            columnWidths,
            colorRules
        } = config;

        const table = document.getElementById(tableId);

        if (!table) {
            alert('No results to export. Please generate results first.');
            return;
        }

        try {
            // Load jsPDF if not already loaded
            if (!this.jsPDFLoaded) {
                await this.loadJsPDF();
            }

            const { jsPDF } = window.jspdf;
            const doc = new jsPDF('l', 'mm', 'a4');

            // Add title
            doc.setFontSize(18);
            doc.setFont(undefined, 'bold');
            doc.text(title, 14, 15);

            // Add timestamp
            doc.setFontSize(10);
            doc.setFont(undefined, 'normal');
            const timestamp = new Date().toLocaleString();
            doc.text(`Generated: ${timestamp}`, 14, 22);

            // Get summary information
            let summaryHeight = 28;
            if (summarySelector) {
                const summaryDiv = document.querySelector(summarySelector);
                if (summaryDiv) {
                    const summaryText = summaryDiv.textContent.trim();
                    doc.setFontSize(9);
                    const lines = doc.splitTextToSize(summaryText, 260);
                    doc.text(lines, 14, 28);
                    summaryHeight = 28 + (lines.length * 4) + 5;
                }
            }

            // Prepare table data
            const tableData = this.extractTableData(table, colorRules);

            // Calculate column widths
            const colStyles = {};
            if (columnWidths && columnWidths.length > 0) {
                columnWidths.forEach((width, index) => {
                    colStyles[index] = { cellWidth: width };
                });
            }

            // Add table using autoTable plugin
            doc.autoTable({
                head: tableData.headers,
                body: tableData.rows,
                startY: summaryHeight,
                styles: {
                    fontSize: 7,
                    cellPadding: 2,
                    overflow: 'linebreak',
                    cellWidth: 'wrap'
                },
                headStyles: {
                    fillColor: [43, 74, 108],
                    textColor: [255, 255, 255],
                    fontStyle: 'bold',
                    halign: 'center'
                },
                columnStyles: colStyles,
                alternateRowStyles: {
                    fillColor: [245, 245, 245]
                },
                margin: { top: 10, right: 10, bottom: 10, left: 10 },
                didParseCell: (data) => {
                    if (data.section === 'body') {
                        this.applyCellColorRules(data, colorRules, tableData.cellStatuses);
                    }
                }
            });

            // Add page numbers
            const pageCount = doc.internal.getNumberOfPages();
            for (let i = 1; i <= pageCount; i++) {
                doc.setPage(i);
                doc.setFontSize(8);
                doc.text(
                    `Page ${i} of ${pageCount}`,
                    doc.internal.pageSize.getWidth() / 2,
                    doc.internal.pageSize.getHeight() - 10,
                    { align: 'center' }
                );
            }

            // Generate filename with date
            const now = new Date();
            const date = now.toISOString().slice(0, 10);
            const filename = `${filenamePrefix}_${date}.pdf`;

            // Save the PDF
            doc.save(filename);

            console.log('PDF exported successfully:', filename);

        } catch (error) {
            console.error('Error exporting PDF:', error);
            alert('Failed to export PDF. Please try again.');
        }
    }

    /**
     * Export multiple file results to PDF (for templates with multiple file groups)
     * @param {Object} config - Export configuration
     */
    async exportMultiFileToPDF(config) {
        const {
            title,
            fileResults,
            filenamePrefix,
            columnWidths,
            colorRules,
            headers,
            extractRowData
        } = config;

        if (!fileResults || fileResults.length === 0) {
            alert('No results to export. Please generate results first.');
            return;
        }

        try {
            // Load jsPDF if not already loaded
            if (!this.jsPDFLoaded) {
                await this.loadJsPDF();
            }

            const { jsPDF } = window.jspdf;
            const doc = new jsPDF('l', 'mm', 'a4');

            // Add title
            doc.setFontSize(18);
            doc.setFont(undefined, 'bold');
            doc.text(title, 14, 15);

            // Add timestamp
            doc.setFontSize(10);
            doc.setFont(undefined, 'normal');
            const timestamp = new Date().toLocaleString();
            doc.text(`Generated: ${timestamp}`, 14, 22);

            let currentY = 28;

            // Calculate column widths
            const colStyles = {};
            if (columnWidths && columnWidths.length > 0) {
                columnWidths.forEach((width, index) => {
                    colStyles[index] = { cellWidth: width };
                });
            }

            // Process each file result
            for (let fileIndex = 0; fileIndex < fileResults.length; fileIndex++) {
                const fileResult = fileResults[fileIndex];

                // Add page break if not the first file and not enough space
                if (fileIndex > 0) {
                    doc.addPage();
                    currentY = 15;
                }

                // Add file name
                doc.setFontSize(12);
                doc.setFont(undefined, 'bold');
                doc.text(`File: ${fileResult.fileName}`, 14, currentY);
                currentY += 6;

                // Add summary if provided
                if (fileResult.summary) {
                    doc.setFontSize(10);
                    doc.setFont(undefined, 'normal');
                    doc.text(fileResult.summary, 14, currentY);
                    currentY += 8;
                }

                // Extract table data using the provided function
                const rows = extractRowData(fileResult);
                const cellStatuses = fileResult.cellStatuses || [];

                // Add table using autoTable plugin
                doc.autoTable({
                    head: [headers],
                    body: rows,
                    startY: currentY,
                    styles: {
                        fontSize: 8,
                        cellPadding: 3,
                        overflow: 'linebreak',
                        cellWidth: 'wrap'
                    },
                    headStyles: {
                        fillColor: [43, 74, 108],
                        textColor: [255, 255, 255],
                        fontStyle: 'bold',
                        halign: 'center',
                        fontSize: 9
                    },
                    columnStyles: colStyles,
                    alternateRowStyles: {
                        fillColor: [245, 245, 245]
                    },
                    margin: { top: 10, right: 10, bottom: 10, left: 10 },
                    didParseCell: (data) => {
                        if (data.section === 'body') {
                            this.applyCellColorRules(data, colorRules, cellStatuses);
                        }
                    }
                });

                currentY = doc.lastAutoTable.finalY + 10;
            }

            // Add page numbers
            const pageCount = doc.internal.getNumberOfPages();
            for (let i = 1; i <= pageCount; i++) {
                doc.setPage(i);
                doc.setFontSize(8);
                doc.text(
                    `Page ${i} of ${pageCount}`,
                    doc.internal.pageSize.getWidth() / 2,
                    doc.internal.pageSize.getHeight() - 10,
                    { align: 'center' }
                );
            }

            // Generate filename with date
            const now = new Date();
            const date = now.toISOString().slice(0, 10);
            const filename = `${filenamePrefix}_${date}.pdf`;

            // Save the PDF
            doc.save(filename);

            console.log('PDF exported successfully:', filename);

        } catch (error) {
            console.error('Error exporting PDF:', error);
            alert('Failed to export PDF. Please try again.');
        }
    }

    /**
     * Extract table data from HTML table
     * @param {HTMLElement} table - The table element
     * @param {Object} colorRules - Color coding rules
     * @returns {Object} - Headers, rows, and cell statuses
     */
    extractTableData(table, colorRules) {
        const headers = [];
        const rows = [];
        const cellStatuses = []; // Track status for each cell [rowIndex][colIndex]

        // Extract headers (skip filter row if exists)
        const headerRows = table.querySelectorAll('thead tr');
        if (headerRows.length > 0) {
            const headerCells = headerRows[0].querySelectorAll('th');
            headerCells.forEach(cell => {
                headers.push(cell.textContent.trim());
            });
        }

        // Extract rows from tbody
        const tbody = table.querySelector('tbody');
        if (tbody) {
            const bodyRows = tbody.querySelectorAll('tr');

            bodyRows.forEach((row, rowIndex) => {
                const rowData = [];
                const rowStatuses = [];
                const cells = row.querySelectorAll('td');

                cells.forEach((cell, colIndex) => {
                    const { text, status } = this.extractCellContent(cell, colIndex, colorRules);
                    rowData.push(text);
                    rowStatuses.push(status);
                });

                rows.push(rowData);
                cellStatuses.push(rowStatuses);
            });
        }

        return {
            headers: [headers],
            rows: rows,
            cellStatuses: cellStatuses
        };
    }

    /**
     * Extract cell content and determine its status
     * @param {HTMLElement} cell - The cell element
     * @param {number} colIndex - Column index
     * @param {Object} colorRules - Color coding rules
     * @returns {Object} - Text content and status
     */
    extractCellContent(cell, colIndex, colorRules) {
        let text = '';
        let status = 'normal';

        // Check for strong elements (file names, product IDs)
        const strong = cell.querySelector('strong');
        const divs = cell.querySelectorAll('div');
        const spans = cell.querySelectorAll('span');

        if (strong) {
            text = strong.textContent.trim();
            if (divs.length > 0) {
                divs.forEach(div => {
                    text += '\n' + div.textContent.trim();
                });
            }
        } else if (spans.length > 0) {
            // Extract status from span styles
            const textParts = [];
            spans.forEach(span => {
                const spanText = span.textContent.trim();
                const style = span.getAttribute('style') || '';

                // Determine status based on color
                if (style.includes('#065f46') || style.includes('rgb(6, 95, 70)')) {
                    status = 'valid'; // Green
                } else if (style.includes('#d97706') || style.includes('rgb(217, 119, 6)')) {
                    status = 'warning'; // Yellow/Orange
                } else if (style.includes('#991b1b') || style.includes('rgb(153, 27, 27)')) {
                    status = 'invalid'; // Red
                }

                // Skip "Expected:" text for cleaner PDF
                if (!spanText.includes('Expected:')) {
                    textParts.push(spanText);
                }
            });
            text = textParts.join(' ');

            // Check for specific status indicators in text
            if (cell.textContent.includes('FOUND') && cell.textContent.includes('✓')) {
                status = 'valid';
                text = '✓ FOUND';
            } else if (cell.textContent.includes('NOT FOUND') && cell.textContent.includes('✗')) {
                status = 'invalid';
                text = '✗ NOT FOUND';
            } else if (cell.textContent.includes('Cell Empty')) {
                status = 'invalid';
            } else if (cell.textContent.includes('Expected:')) {
                // Has expected value means there's a mismatch
                if (status === 'normal') {
                    status = 'warning';
                }
            }
        } else {
            text = cell.textContent.trim();
        }

        // Clean up extra whitespace
        text = text.replace(/\s+/g, ' ').trim();

        return { text, status };
    }

    /**
     * Apply color rules to cells during PDF generation
     * @param {Object} data - autoTable cell data
     * @param {Object} colorRules - Color coding rules
     * @param {Array} cellStatuses - Cell status matrix
     */
    applyCellColorRules(data, colorRules, cellStatuses) {
        const rowIndex = data.row.index;
        const colIndex = data.column.index;
        const cellText = data.cell.text[0] || '';

        // Get status from cellStatuses if available
        let status = 'normal';
        if (cellStatuses && cellStatuses[rowIndex] && cellStatuses[rowIndex][colIndex]) {
            status = cellStatuses[rowIndex][colIndex];
        }

        // Color codes
        const GREEN = [6, 95, 70];      // #065f46 - Valid
        const YELLOW = [217, 119, 6];   // #d97706 - Warning
        const RED = [153, 27, 27];      // #991b1b - Invalid

        // Apply colors based on status
        if (status === 'valid') {
            data.cell.styles.textColor = GREEN;
            data.cell.styles.fontStyle = 'bold';
        } else if (status === 'warning') {
            data.cell.styles.textColor = YELLOW;
            data.cell.styles.fontStyle = 'bold';
        } else if (status === 'invalid') {
            data.cell.styles.textColor = RED;
            data.cell.styles.fontStyle = 'bold';
        }

        // Additional text-based color detection as fallback
        if (status === 'normal') {
            // Check for FOUND/NOT FOUND status
            if (cellText.includes('✓ FOUND') || cellText.includes('✓FOUND')) {
                data.cell.styles.textColor = GREEN;
                data.cell.styles.fontStyle = 'bold';
            } else if (cellText.includes('✗ NOT FOUND') || cellText.includes('✗NOT FOUND')) {
                data.cell.styles.textColor = RED;
                data.cell.styles.fontStyle = 'bold';
            }
            // Check for Empty or error indicators
            else if (cellText.includes('Empty') || cellText.includes('⚠️')) {
                data.cell.styles.textColor = RED;
                data.cell.styles.fontStyle = 'bold';
            }
            // Check for Expected: indicator (mismatch)
            else if (cellText.includes('Expected:')) {
                data.cell.styles.textColor = YELLOW;
                data.cell.styles.fontStyle = 'bold';
            }
            // Check for BCBD/OB comparison with difference
            else if (cellText.includes('BCBD:') && cellText.includes('OB Total SMV:')) {
                const diffMatch = cellText.match(/\([\+\-]([\d.]+)\)/);
                if (diffMatch) {
                    const difference = parseFloat(diffMatch[1]);
                    if (difference <= 0.01) {
                        data.cell.styles.textColor = YELLOW;
                    } else {
                        data.cell.styles.textColor = RED;
                    }
                    data.cell.styles.fontStyle = 'bold';
                }
            }
        }

        // Apply template-specific color rules if provided
        if (colorRules && colorRules.columns && colorRules.columns[colIndex]) {
            const colRule = colorRules.columns[colIndex];
            if (colRule.validator) {
                const validationResult = colRule.validator(cellText);
                if (validationResult === 'valid') {
                    data.cell.styles.textColor = GREEN;
                    data.cell.styles.fontStyle = 'bold';
                } else if (validationResult === 'warning') {
                    data.cell.styles.textColor = YELLOW;
                    data.cell.styles.fontStyle = 'bold';
                } else if (validationResult === 'invalid') {
                    data.cell.styles.textColor = RED;
                    data.cell.styles.fontStyle = 'bold';
                }
            }
        }
    }

    /**
     * Create TNF (V1) export configuration
     * @returns {Object} - Configuration object for TNF export
     */
    createTNFConfig() {
        return {
            title: 'Costing Validation Results - V1',
            tableId: 'v1ResultsTable',
            summarySelector: '#results-v1 div[style*="background: #f0f7ff"]',
            filenamePrefix: 'CostingValidation_V1',
            columnWidths: [35, 30, 25, 30, 25, 30, 35, 25],
            colorRules: {
                columns: {
                    2: { // Match Status column
                        validator: (text) => {
                            if (text.includes('✓ FOUND')) return 'valid';
                            if (text.includes('✗ NOT FOUND')) return 'invalid';
                            return 'normal';
                        }
                    },
                    3: { // Standard Minute Value column
                        validator: (text) => {
                            if (text.includes('Empty') || text.includes('TNF: Empty')) return 'invalid';
                            if (text.includes('BCBD:')) {
                                const diffMatch = text.match(/\([\+\-]([\d.]+)\)/);
                                if (diffMatch) {
                                    const diff = parseFloat(diffMatch[1]);
                                    if (diff <= 0.01) return 'warning';
                                    return 'invalid';
                                }
                                return 'warning';
                            }
                            if (text !== '-' && text !== '') return 'valid';
                            return 'normal';
                        }
                    },
                    4: { // Average Efficiency %
                        validator: (text) => {
                            if (text.includes('Cell Empty')) return 'invalid';
                            if (text === '-') return 'normal';
                            const match = text.match(/([\d.]+)%/);
                            if (match) {
                                const value = parseFloat(match[1]);
                                if (Math.abs(value - 50.0) < 0.1) return 'valid';
                                return 'invalid';
                            }
                            return 'normal';
                        }
                    },
                    5: { // Hourly Wages
                        validator: (text) => {
                            if (text.includes('Cell Empty')) return 'invalid';
                            if (text === '-') return 'normal';
                            const match = text.match(/([\d.]+)/);
                            if (match) {
                                const value = parseFloat(match[1]);
                                if (Math.abs(value - 1.750) < 0.01) return 'valid';
                                return 'invalid';
                            }
                            return 'normal';
                        }
                    },
                    6: { // Overhead Cost
                        validator: (text) => {
                            if (text.includes('Cell Empty')) return 'invalid';
                            if (text === '-') return 'normal';
                            const match = text.match(/([\d.]+)%/);
                            if (match) {
                                const value = parseFloat(match[1]);
                                if (Math.abs(value - 70.0) < 0.1) return 'valid';
                                return 'invalid';
                            }
                            return 'normal';
                        }
                    },
                    7: { // Factory Profit %
                        validator: (text) => {
                            if (text.includes('Cell Empty')) return 'invalid';
                            if (text === '-') return 'normal';
                            const match = text.match(/([\d.]+)%/);
                            if (match) {
                                const value = parseFloat(match[1]);
                                if (Math.abs(value - 10.0) < 0.1) return 'valid';
                                return 'warning';
                            }
                            return 'normal';
                        }
                    }
                }
            }
        };
    }

    /**
     * Create Fjall Raven (V5) export configuration
     * @param {Array} fileResults - Array of file results from Fjall Raven processor
     * @returns {Object} - Configuration object for Fjall Raven export
     */
    createFjallRavenConfig(fileResults) {
        return {
            title: 'Fjall Raven Cost Breakdown Comparison - V5',
            fileResults: fileResults.map(fileResult => {
                // Count fully valid items (excluding special items from count)
                const regularItems = fileResult.results.filter(r => !r.isSpecialItem);
                const totalItems = regularItems.length;
                const validItems = regularItems.filter(r =>
                    r.supplierMaterialCode.status !== 'INVALID' &&
                    r.bomSection.status !== 'INVALID' &&
                    r.supplier.status !== 'INVALID' &&
                    r.qty.status !== 'INVALID' &&
                    r.price.status !== 'INVALID' &&
                    r.freight.status !== 'INVALID' &&
                    r.waste.status !== 'INVALID'
                ).length;

                // Build cell statuses for coloring
                const cellStatuses = fileResult.results.map(item => {
                    if (item.isSpecialItem) {
                        return [
                            'normal', // Item (-)
                            'normal', // Supplier Mat. Code (-)
                            'normal', // BOM Section (item name)
                            'normal', // Supplier (-)
                            item.laborCost.status === 'VALID' ? 'valid' : (item.laborCost.status === 'N/A' ? 'normal' : 'invalid'),
                            item.miscellaneous.status === 'VALID' ? 'valid' : (item.miscellaneous.status === 'N/A' ? 'normal' : 'invalid'),
                            item.qty.status === 'VALID' ? 'valid' : (item.qty.status === 'N/A' ? 'normal' : 'invalid'),
                            item.firstCost.status === 'VALID' ? 'valid' : (item.firstCost.status === 'N/A' ? 'normal' : 'invalid'),
                            item.price.status === 'VALID' ? 'valid' : (item.price.status === 'N/A' ? 'normal' : 'invalid'),
                            item.freight.status === 'VALID' ? 'valid' : (item.freight.status === 'N/A' ? 'normal' : 'invalid'),
                            item.waste.status === 'VALID' ? 'valid' : (item.waste.status === 'N/A' ? 'normal' : 'invalid')
                        ];
                    } else {
                        return [
                            'normal', // Item name
                            item.supplierMaterialCode.status === 'VALID' ? 'valid' : (item.supplierMaterialCode.status === 'N/A' ? 'normal' : 'invalid'),
                            item.bomSection.status === 'VALID' ? 'valid' : (item.bomSection.status === 'N/A' ? 'normal' : 'invalid'),
                            item.supplier.status === 'VALID' ? 'valid' : (item.supplier.status === 'N/A' ? 'normal' : 'invalid'),
                            item.laborCost.status === 'VALID' ? 'valid' : (item.laborCost.status === 'N/A' ? 'normal' : 'invalid'),
                            item.miscellaneous.status === 'VALID' ? 'valid' : (item.miscellaneous.status === 'N/A' ? 'normal' : 'invalid'),
                            item.qty.status === 'VALID' ? 'valid' : (item.qty.status === 'N/A' ? 'normal' : 'invalid'),
                            item.firstCost.status === 'VALID' ? 'valid' : (item.firstCost.status === 'N/A' ? 'normal' : 'invalid'),
                            item.price.status === 'VALID' ? 'valid' : (item.price.status === 'N/A' ? 'normal' : 'invalid'),
                            item.freight.status === 'VALID' ? 'valid' : (item.freight.status === 'N/A' ? 'normal' : 'invalid'),
                            item.waste.status === 'VALID' ? 'valid' : (item.waste.status === 'N/A' ? 'normal' : 'invalid')
                        ];
                    }
                });

                return {
                    fileName: fileResult.fileName,
                    summary: `Summary: ${validItems} out of ${totalItems} items match`,
                    results: fileResult.results,
                    cellStatuses: cellStatuses
                };
            }),
            filenamePrefix: 'FjallRavenCostBreakdown_V5',
            columnWidths: [30, 25, 25, 25, 20, 20, 15, 20, 20, 20, 20],
            headers: ['Item', 'Supplier Mat. Code', 'BOM Section', 'Supplier', 'Labor Cost', 'Misc.', 'Qty', 'First Cost', 'Price', 'Freight', 'Waste'],
            colorRules: {},
            extractRowData: (fileResult) => {
                const rows = [];
                for (const item of fileResult.results) {
                    // Helper to format cell for PDF
                    const formatForPDF = (field) => {
                        if (!field) return '-';
                        if (field.status === 'N/A') return '-';

                        const displayValue = (field.buyer !== undefined && field.buyer !== null && field.buyer !== '')
                            ? field.buyer
                            : '0';

                        if (field.status === 'VALID') {
                            return displayValue;
                        } else {
                            return `${displayValue} (Expected: ${field.ob})`;
                        }
                    };

                    if (item.isSpecialItem) {
                        rows.push([
                            '-',
                            '-',
                            item.itemName,
                            '-',
                            formatForPDF(item.laborCost),
                            formatForPDF(item.miscellaneous),
                            formatForPDF(item.qty),
                            formatForPDF(item.firstCost),
                            formatForPDF(item.price),
                            formatForPDF(item.freight),
                            formatForPDF(item.waste)
                        ]);
                    } else {
                        rows.push([
                            item.itemName,
                            formatForPDF(item.supplierMaterialCode),
                            formatForPDF(item.bomSection),
                            formatForPDF(item.supplier),
                            formatForPDF(item.laborCost),
                            formatForPDF(item.miscellaneous),
                            formatForPDF(item.qty),
                            formatForPDF(item.firstCost),
                            formatForPDF(item.price),
                            formatForPDF(item.freight),
                            formatForPDF(item.waste)
                        ]);
                    }
                }
                return rows;
            }
        };
    }

    /**
     * Create Helly Hansen (V4) export configuration
     * @param {Array} fileResults - Array of file results from Helly Hansen processor
     * @param {Function} formatToFourDecimals - Helper function to format decimals
     * @returns {Object} - Configuration object for Helly Hansen export
     */
    createHellyHansenConfig(fileResults, formatToFourDecimals) {
        return {
            title: 'Helly Hansen Cost Breakdown Comparison - V4',
            fileResults: fileResults.map(fileResult => {
                const totalItems = fileResult.results.length;
                const validItems = fileResult.results.filter(r =>
                    (r.consmStatus === 'VALID' || r.consmStatus === 'N/A') &&
                    (r.upStatus === 'VALID' || r.upStatus === 'N/A') &&
                    r.amountStatus === 'VALID'
                ).length;

                // Build cell statuses for coloring
                const cellStatuses = fileResult.results.map(item => {
                    return [
                        'normal', // Item name
                        item.consmStatus === 'VALID' ? 'valid' : (item.consmStatus === 'N/A' ? 'normal' : 'invalid'),
                        item.upStatus === 'VALID' ? 'valid' : (item.upStatus === 'N/A' ? 'normal' : 'invalid'),
                        item.amountStatus === 'VALID' ? 'valid' : 'invalid'
                    ];
                });

                return {
                    fileName: fileResult.fileName,
                    summary: `Summary: ${validItems} out of ${totalItems} items fully match`,
                    results: fileResult.results,
                    cellStatuses: cellStatuses
                };
            }),
            filenamePrefix: 'HellyHansenCostBreakdown_V4',
            columnWidths: [80, 40, 40, 40],
            headers: ['Item', 'CONSM', 'U/P', 'Amount'],
            colorRules: {},
            extractRowData: (fileResult) => {
                const rows = [];
                for (const item of fileResult.results) {
                    // Helper to format cell for PDF
                    const formatForPDF = (obVal, buyerVal, status, isNumeric = true, specialCase = null) => {
                        // Handle N/A status
                        if (status === 'N/A') {
                            return '-';
                        }

                        // Special handling for MARGIN_PROFIT - check BEFORE empty check
                        if (specialCase === 'MARGIN_PROFIT') {
                            if (!buyerVal || buyerVal === '' || buyerVal === 'NOT FOUND') {
                                return `Empty (Expected: 0.45 to 0.55)`;
                            }
                            const displayBuyer = formatToFourDecimals(buyerVal);
                            return `${displayBuyer} (Expected: 0.45 to 0.55)`;
                        }

                        // Special handling for FINANCIAL_OVERHEAD - check BEFORE empty check
                        if (specialCase === 'FINANCIAL_OVERHEAD') {
                            const expectedText = item.countryOfOrigin ? `${obVal} - ${item.countryOfOrigin}` : obVal;
                            if (!buyerVal || buyerVal === '' || buyerVal === 'NOT FOUND') {
                                return `Empty (Expected: ${expectedText})`;
                            }
                            const displayBuyer = formatToFourDecimals(buyerVal);
                            return `${displayBuyer} (Expected: ${expectedText})`;
                        }

                        if (!buyerVal || buyerVal === '' || buyerVal === 'NOT FOUND') {
                            const displayOB = isNumeric && obVal ? formatToFourDecimals(obVal) : obVal;
                            return `Empty (Expected: ${displayOB})`;
                        }

                        const displayBuyer = isNumeric ? formatToFourDecimals(buyerVal) : buyerVal;
                        const displayOB = isNumeric ? formatToFourDecimals(obVal) : obVal;

                        if (status === 'VALID') {
                            return displayBuyer;
                        } else {
                            return `${displayBuyer} (Expected: ${displayOB})`;
                        }
                    };

                    rows.push([
                        item.itemName,
                        formatForPDF(item.obConsm, item.buyerConsm, item.consmStatus, true, item.specialCase),
                        formatForPDF(item.obUp, item.buyerUp, item.upStatus, true, item.specialCase),
                        formatForPDF(item.obAmount, item.buyerAmount, item.amountStatus, true, item.specialCase)
                    ]);
                }
                return rows;
            }
        };
    }

    /**
     * Create Columbia (V3) export configuration
     * @param {Array} fileResults - Array of file results from Columbia processor
     * @param {Function} formatToThreeDecimals - Helper function to format decimals
     * @returns {Object} - Configuration object for Columbia export
     */
    createColumbiaConfig(fileResults, formatToThreeDecimals) {
        return {
            title: 'Columbia Cost Breakdown Comparison - V3',
            fileResults: fileResults.map(fileResult => {
                const totalItems = fileResult.results.length;
                const validItems = fileResult.results.filter(r =>
                    r.materialStatus === 'VALID' &&
                    r.fobCostStatus === 'VALID' &&
                    r.factoryUsageStatus === 'VALID' &&
                    r.wastageStatus === 'VALID'
                ).length;

                // Build cell statuses for coloring (skip Hangtag Package Part with material 1234)
                const cellStatuses = fileResult.results
                    .filter(item => !(item.itemName === 'Hangtag Package Part' && item.obMaterial === '1234'))
                    .map(item => {
                        return [
                            'normal', // Item name
                            item.materialStatus === 'VALID' ? 'valid' : 'invalid',
                            item.fobCostStatus === 'VALID' ? 'valid' : 'invalid',
                            item.factoryUsageStatus === 'VALID' ? 'valid' : 'invalid',
                            item.wastageStatus === 'VALID' ? 'valid' : 'invalid'
                        ];
                    });

                return {
                    fileName: fileResult.fileName,
                    summary: `Summary: ${validItems} out of ${totalItems} items fully match`,
                    results: fileResult.results,
                    cellStatuses: cellStatuses
                };
            }),
            filenamePrefix: 'ColumbiaCostBreakdown_V3',
            columnWidths: [50, 50, 40, 40, 40],
            headers: ['Item', 'Material', 'FOB Cost', 'Factory Usage', 'Wastage'],
            colorRules: {},
            extractRowData: (fileResult) => {
                const rows = [];
                for (const item of fileResult.results) {
                    // Skip Hangtag Package Part with material 1234
                    if (item.itemName === 'Hangtag Package Part' && item.obMaterial === '1234') {
                        continue;
                    }

                    // Helper to format cell for PDF
                    const formatForPDF = (obVal, buyerVal, status, isNumeric = false) => {
                        if (!buyerVal || buyerVal === '' || buyerVal === 'NOT FOUND') {
                            const displayOB = isNumeric && obVal ? formatToThreeDecimals(obVal) : obVal;
                            return `Empty (Expected: ${displayOB})`;
                        }

                        const displayBuyer = isNumeric ? formatToThreeDecimals(buyerVal) : buyerVal;
                        const displayOB = isNumeric ? formatToThreeDecimals(obVal) : obVal;

                        if (status === 'VALID') {
                            return displayBuyer;
                        } else {
                            return `${displayBuyer} (Expected: ${displayOB})`;
                        }
                    };

                    rows.push([
                        item.itemName,
                        formatForPDF(item.obMaterial, item.buyerMaterial, item.materialStatus, false),
                        formatForPDF(item.obFobCost, item.buyerFobCost, item.fobCostStatus, true),
                        formatForPDF(item.obFactoryUsage, item.buyerFactoryUsage, item.factoryUsageStatus, true),
                        formatForPDF(item.obWastage, item.buyerWastage, item.wastageStatus, true)
                    ]);
                }
                return rows;
            }
        };
    }

    /**
     * Create Outdoor Research (V8) export configuration
     * @param {Array} fileResults - Array of file results from Outdoor Research processor
     * @returns {Object} - Configuration object for Outdoor Research export
     */
    createOutdoorResearchConfig(fileResults) {
        return {
            title: 'Outdoor Research Validation Results - V8',
            fileResults: fileResults.map(fileResult => {
                const gpCheck = fileResult.results.generalPackagingCheck;
                const ocCheck = fileResult.results.otherChargesCheck;

                // Count valid checks
                let validCount = 0;
                let totalChecks = 0;

                if (gpCheck && gpCheck.found && gpCheck.checks) {
                    totalChecks++;
                    if (gpCheck.isValid) validCount++;
                }
                if (ocCheck && ocCheck.found && ocCheck.checks) {
                    totalChecks++;
                    if (ocCheck.isValid) validCount++;
                }

                // Build cell statuses for coloring
                const cellStatuses = [];

                // General Packaging row
                if (gpCheck && gpCheck.found && gpCheck.checks) {
                    cellStatuses.push([
                        'normal',
                        gpCheck.isValid ? 'valid' : 'invalid'
                    ]);
                } else {
                    cellStatuses.push([
                        'normal',
                        'invalid'
                    ]);
                }

                // Other Charges row
                if (ocCheck && ocCheck.found && ocCheck.checks) {
                    cellStatuses.push([
                        'normal',
                        ocCheck.isValid ? 'valid' : 'invalid'
                    ]);
                } else {
                    cellStatuses.push([
                        'normal',
                        'invalid'
                    ]);
                }

                return {
                    fileName: fileResult.fileName,
                    summary: `Summary: ${validCount} out of ${totalChecks} checks passed`,
                    results: fileResult.results,
                    cellStatuses: cellStatuses
                };
            }),
            filenamePrefix: 'OutdoorResearchValidation_V8',
            columnWidths: [45, 105],
            headers: ['Validation Check', 'Value'],
            colorRules: {},
            extractRowData: (fileResult) => {
                const rows = [];
                const gpCheck = fileResult.results.generalPackagingCheck;
                const ocCheck = fileResult.results.otherChargesCheck;

                // General Packaging row
                if (gpCheck && gpCheck.found && gpCheck.checks) {
                    const checkDetails = gpCheck.checks.map(check => {
                        return `${check.label}: ${check.actualValue || 'Empty'}`;
                    }).join(' | ');

                    rows.push([
                        `General Packaging (Row ${gpCheck.rowNumber})`,
                        checkDetails
                    ]);
                } else if (gpCheck && !gpCheck.found) {
                    rows.push([
                        'General Packaging',
                        gpCheck.message || 'Not found'
                    ]);
                }

                // Other Charges row
                if (ocCheck && ocCheck.found && ocCheck.checks) {
                    const checkDetails = ocCheck.checks.map(check => {
                        return `${check.label}: ${check.actualValue || 'Empty'}`;
                    }).join(' | ');

                    rows.push([
                        `Other Charges (Row ${ocCheck.rowNumber})`,
                        checkDetails
                    ]);
                } else if (ocCheck && !ocCheck.found) {
                    rows.push([
                        'Other Charges',
                        ocCheck.message || 'Not found'
                    ]);
                }

                return rows;
            }
        };
    }

    /**
     * Create Mammut (V7) export configuration
     * @param {Array} fileResults - Array of file results from Mammut processor
     * @returns {Object} - Configuration object for Mammut export
     */
    createMammutConfig(fileResults) {
        return {
            title: 'Mammut Validation Results - V7',
            fileResults: fileResults.map(fileResult => {
                const cellChecks = fileResult.results.cellChecks;
                const profitMargin = fileResult.results.profitMarginCheck;
                const wastageCost = fileResult.results.wastageCostCheck;
                const cmtCheck = fileResult.results.cmtCheck;

                // Count valid checks
                let validCount = 0;
                const wastageSectionCount = wastageCost.found && wastageCost.sections ? wastageCost.sections.length : 1;
                const cmtItemCount = cmtCheck && cmtCheck.found && cmtCheck.items ? cmtCheck.items.length : 0;
                let totalChecks = cellChecks.length + 1 + wastageSectionCount + cmtItemCount;

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

                // Build cell statuses for coloring
                const cellStatuses = [];

                // Cell checks rows
                cellChecks.forEach(check => {
                    cellStatuses.push([
                        'normal',
                        check.isValid ? 'valid' : 'invalid'
                    ]);
                });

                // Profit margin row
                cellStatuses.push([
                    'normal',
                    profitMargin.found && profitMargin.isValid ? 'valid' : 'invalid'
                ]);

                // Wastage rows
                if (wastageCost.found && wastageCost.sections) {
                    wastageCost.sections.forEach(section => {
                        cellStatuses.push([
                            'normal',
                            section.isValid ? 'valid' : 'invalid'
                        ]);
                    });
                }

                // CMT rows
                if (cmtCheck && cmtCheck.found && cmtCheck.items) {
                    cmtCheck.items.forEach(item => {
                        cellStatuses.push([
                            'normal',
                            item.isValid ? 'valid' : 'invalid'
                        ]);
                    });
                }

                return {
                    fileName: fileResult.fileName,
                    summary: `Summary: ${validCount} out of ${totalChecks} checks passed`,
                    results: fileResult.results,
                    cellStatuses: cellStatuses
                };
            }),
            filenamePrefix: 'MammutValidation_V7',
            columnWidths: [45, 105],
            headers: ['Validation Check', 'Value'],
            colorRules: {},
            extractRowData: (fileResult) => {
                const rows = [];
                const cellChecks = fileResult.results.cellChecks;
                const profitMargin = fileResult.results.profitMarginCheck;
                const wastageCost = fileResult.results.wastageCostCheck;
                const cmtCheck = fileResult.results.cmtCheck;

                // Helper to format number
                const formatNumber = (value) => {
                    const num = parseFloat(String(value).replace(/[$,\s%]/g, ''));
                    if (isNaN(num)) return value;
                    return num.toFixed(2);
                };

                // Cell checks rows
                for (const check of cellChecks) {
                    rows.push([
                        check.label,
                        check.found ? (check.actualValue || 'Empty') : 'Not found'
                    ]);
                }

                // Profit Margin row
                if (profitMargin.found) {
                    const pmActual = profitMargin.numericValue !== null ? formatNumber(profitMargin.actualValue) : profitMargin.actualValue;
                    const pmRange = `${formatNumber(profitMargin.minValue)} - ${formatNumber(profitMargin.maxValue)}`;
                    rows.push([
                        'PROFIT MARGIN',
                        `${pmActual} (Expected: ${pmRange})`
                    ]);
                } else {
                    rows.push([
                        'PROFIT MARGIN',
                        profitMargin.message || 'Not found'
                    ]);
                }

                // Wastage rows for each section
                if (wastageCost.found && wastageCost.sections) {
                    for (const section of wastageCost.sections) {
                        const sectionName = section.label.replace(' TOTAL', '');
                        const expectedPercent = (section.expectedValue * 100).toFixed(0);

                        // Combine valid and invalid cells
                        const allCells = [];

                        section.validCells.forEach(cell => {
                            allCells.push(cell.cellAddress);
                        });

                        section.invalidCells.forEach(cell => {
                            const roundedValue = cell.numericValue.toFixed(2);
                            allCells.push(`${cell.cellAddress} (${roundedValue})`);
                        });

                        rows.push([
                            `${sectionName} Wastage (${expectedPercent}%)`,
                            allCells.length > 0 ? allCells.join(', ') : 'No data'
                        ]);
                    }
                } else {
                    rows.push([
                        'Wastage Check',
                        wastageCost.message || 'No sections found'
                    ]);
                }

                // CMT (Cut, Make, Trim) rows
                if (cmtCheck && cmtCheck.found && cmtCheck.items) {
                    for (const item of cmtCheck.items) {
                        if (!item.found) {
                            rows.push([
                                item.label,
                                'Not found'
                            ]);
                        } else {
                            const details = [];

                            const priceDisplay = item.numericPrice !== null ? item.numericPrice.toFixed(2) : item.actualPrice;
                            details.push(`Price: ${priceDisplay}`);

                            const exRateDisplay = item.numericExRate !== null ? item.numericExRate.toFixed(2) : item.actualExRate;
                            details.push(`Ex.Rate: ${exRateDisplay}`);

                            details.push(`Currency: ${item.actualCurrency || 'N/A'}`);

                            rows.push([
                                item.label,
                                details.join(' | ')
                            ]);
                        }
                    }
                } else if (cmtCheck && !cmtCheck.found) {
                    rows.push([
                        'Process Check',
                        cmtCheck.message || 'Not found'
                    ]);
                }

                return rows;
            }
        };
    }

    /**
     * Create LLBEAN (V6) export configuration
     * @param {Array} fileResults - Array of file results from LLBEAN processor
     * @returns {Object} - Configuration object for LLBEAN export
     */
    createLLBEANConfig(fileResults) {
        return {
            title: 'LLBEAN Validation Results - V6',
            fileResults: fileResults.map(fileResult => {
                const b5 = fileResult.results.b5Check;
                const trimsBox = fileResult.results.trimsBoxCheck;
                const totalFinancial = fileResult.results.totalFinancialCostCheck;

                // Count valid checks
                let validCount = 0;
                let totalChecks = 3;

                if (b5.isValid) validCount++;
                if (trimsBox.found && trimsBox.isValid) validCount++;
                if (totalFinancial.found && totalFinancial.isValid) validCount++;

                // Build cell statuses for coloring
                const cellStatuses = [];

                // B5 Keywords row
                cellStatuses.push([
                    'normal', // Check name
                    'normal', // Supplier
                    b5.isValid ? 'valid' : 'invalid', // Consumption (keywords)
                    'normal', // Unit Price
                    'normal', // Total Cost
                    'normal'  // Status
                ]);

                // Trims Box row
                if (trimsBox.found && trimsBox.boxData) {
                    cellStatuses.push([
                        'normal', // Check name
                        trimsBox.validation.supplier === 'VALID' ? 'valid' : 'invalid',
                        trimsBox.validation.consumption === 'VALID' ? 'valid' : 'invalid',
                        trimsBox.validation.unitPrice === 'VALID' ? 'valid' : 'invalid',
                        trimsBox.validation.totalCost === 'VALID' ? 'valid' : 'invalid',
                        'normal'  // Status
                    ]);
                } else {
                    cellStatuses.push([
                        'normal', 'invalid', 'invalid', 'invalid', 'invalid', 'normal'
                    ]);
                }

                // Total Financial Cost row
                if (totalFinancial.found && totalFinancial.expectedValue !== null) {
                    cellStatuses.push([
                        'normal', // Check name
                        'normal', // Supplier
                        'normal', // Consumption
                        'normal', // Unit Price
                        totalFinancial.validation === 'VALID' ? 'valid' : 'invalid',
                        'normal'  // Status
                    ]);
                } else {
                    cellStatuses.push([
                        'normal', 'normal', 'normal', 'normal', 'invalid', 'normal'
                    ]);
                }

                return {
                    fileName: fileResult.fileName,
                    summary: `Summary: ${validCount} out of ${totalChecks} checks passed`,
                    results: fileResult.results,
                    cellStatuses: cellStatuses
                };
            }),
            filenamePrefix: 'LLBEANValidation_V6',
            columnWidths: [35, 25, 30, 25, 25, 20],
            headers: ['Validation Check', 'Supplier', 'Consumption', 'Unit Price', 'Total Cost', 'Status'],
            colorRules: {},
            extractRowData: (fileResult) => {
                const rows = [];
                const b5 = fileResult.results.b5Check;
                const trimsBox = fileResult.results.trimsBoxCheck;
                const totalFinancial = fileResult.results.totalFinancialCostCheck;

                // Helper to format cell for PDF
                const formatForPDF = (expected, actual, status, isNumeric = false) => {
                    if (!actual || actual === '') {
                        const displayExpected = isNumeric ? parseFloat(expected).toFixed(2) : expected;
                        return `Empty (Expected: ${displayExpected})`;
                    }

                    const displayActual = isNumeric ? parseFloat(actual).toFixed(2) : actual;
                    const displayExpected = isNumeric ? parseFloat(expected).toFixed(2) : expected;

                    if (status === 'VALID') {
                        return displayActual;
                    } else {
                        return `${displayActual} (Expected: ${displayExpected})`;
                    }
                };

                // B5 Keywords row
                rows.push([
                    `Cell B5 Keywords\n(Value: ${b5.cellValue || 'Empty'})`,
                    '-',
                    b5.foundKeywords.length > 0 ? b5.foundKeywords.join(', ') : b5.requiredKeywords.join(', '),
                    '-',
                    '-',
                    b5.isValid ? 'VALID' : 'INVALID'
                ]);

                // Trims Box row
                if (trimsBox.found && trimsBox.boxData) {
                    const v = trimsBox.validation;
                    const expected = trimsBox.expected;
                    const actual = trimsBox.boxData;

                    rows.push([
                        `Trims - Box\n(Row ${actual.rowNumber})`,
                        formatForPDF(expected.supplier, actual.supplier, v.supplier, false),
                        formatForPDF(expected.consumption, actual.consumption, v.consumption, true),
                        formatForPDF(expected.unitPrice, actual.unitPrice, v.unitPrice, true),
                        formatForPDF(expected.totalCost, actual.totalCost, v.totalCost, true),
                        trimsBox.isValid ? 'VALID' : 'INVALID'
                    ]);
                } else {
                    rows.push([
                        'Trims - Box',
                        '-',
                        trimsBox.message || 'Not found',
                        '-',
                        '-',
                        'NOT FOUND'
                    ]);
                }

                // Total Financial Cost row
                if (totalFinancial.found && totalFinancial.expectedValue !== null) {
                    rows.push([
                        `Total Financial Cost\n(Row ${totalFinancial.rowNumber}, ${totalFinancial.matchedKeyword})`,
                        '-',
                        '-',
                        '-',
                        formatForPDF(totalFinancial.expectedValue, totalFinancial.actualValue, totalFinancial.validation, true),
                        totalFinancial.isValid ? 'VALID' : 'INVALID'
                    ]);
                } else {
                    rows.push([
                        'Total Financial Cost',
                        '-',
                        totalFinancial.message || 'Not found',
                        '-',
                        '-',
                        'NOT FOUND'
                    ]);
                }

                return rows;
            }
        };
    }

    /**
     * Create Burton (V2) export configuration
     * @param {Array} fileResults - Array of file results from Burton processor
     * @param {Function} formatToThreeDecimals - Helper function to format decimals
     * @returns {Object} - Configuration object for Burton export
     */
    createBurtonConfig(fileResults, formatToThreeDecimals) {
        return {
            title: 'Burton Cost Breakdown Comparison - V2',
            fileResults: fileResults.map(fileResult => {
                const totalItems = fileResult.results.length;
                const validItems = fileResult.results.filter(r => {
                    if (r.status !== 'FOUND') return false;
                    const comp = r.comparison;
                    return Object.values(comp).every(v => v === 'VALID');
                }).length;

                // Build cell statuses for coloring
                const cellStatuses = fileResult.results.map(item => {
                    if (item.status !== 'FOUND') {
                        return ['normal', 'invalid', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal'];
                    }
                    const comp = item.comparison;
                    return [
                        'normal', // Item name
                        comp.material === 'VALID' ? 'valid' : (comp.material === 'WARNING' ? 'warning' : 'invalid'),
                        comp.supplier === 'VALID' ? 'valid' : (comp.supplier === 'WARNING' ? 'warning' : 'invalid'),
                        comp.qty === 'VALID' ? 'valid' : (comp.qty === 'WARNING' ? 'warning' : 'invalid'),
                        comp.wastage === 'VALID' ? 'valid' : (comp.wastage === 'WARNING' ? 'warning' : 'invalid'),
                        comp.unit === 'VALID' ? 'valid' : (comp.unit === 'WARNING' ? 'warning' : 'invalid'),
                        comp.unitPrice === 'VALID' ? 'valid' : (comp.unitPrice === 'WARNING' ? 'warning' : 'invalid'),
                        comp.total === 'VALID' ? 'valid' : (comp.total === 'WARNING' ? 'warning' : 'invalid')
                    ];
                });

                return {
                    fileName: fileResult.fileName,
                    summary: `Summary: ${validItems} out of ${totalItems} items fully match the OB file`,
                    results: fileResult.results,
                    cellStatuses: cellStatuses
                };
            }),
            filenamePrefix: 'BurtonCostBreakdown_V2',
            columnWidths: [55, 45, 40, 15, 25, 15, 30, 30],
            headers: ['Item Name', 'Material', 'Supplier', 'Qty', 'Wastage', 'Unit', 'Unit Price', 'Total'],
            colorRules: {
                // Generic color rules for Burton
            },
            extractRowData: (fileResult) => {
                const rows = [];
                for (const item of fileResult.results) {
                    if (item.status === 'NOT_FOUND_IN_OB') {
                        rows.push([
                            item.itemName,
                            '⚠️ Not found in OB file',
                            '', '', '', '', '', ''
                        ]);
                    } else if (item.status === 'NOT_FOUND_IN_BUYER') {
                        rows.push([
                            item.itemName,
                            '⚠️ Not found in Buyer CBD file',
                            '', '', '', '', '', ''
                        ]);
                    } else {
                        const comp = item.comparison;
                        const obData = item.obData;
                        const buyerData = item.buyerData;

                        // Helper to format cell for PDF
                        const formatForPDF = (obVal, buyerVal, status, isNumeric) => {
                            if (!buyerVal || buyerVal === '') {
                                const displayOB = isNumeric ? formatToThreeDecimals(obVal) : obVal;
                                return `Empty (Expected: ${displayOB})`;
                            }

                            const displayBuyer = isNumeric ? formatToThreeDecimals(buyerVal) : buyerVal;
                            const displayOB = isNumeric ? formatToThreeDecimals(obVal) : obVal;

                            if (status === 'VALID') {
                                return displayBuyer;
                            } else {
                                return `${displayBuyer} (Expected: ${displayOB})`;
                            }
                        };

                        rows.push([
                            item.itemName,
                            formatForPDF(obData.materialName, buyerData.material, comp.material, false),
                            formatForPDF(obData.supplier, buyerData.supplier, comp.supplier, false),
                            formatForPDF(obData.quantity, buyerData.qty, comp.qty, false),
                            formatForPDF(obData.wastage, buyerData.wastage, comp.wastage, true),
                            formatForPDF(obData.unit, buyerData.unit, comp.unit, false),
                            formatForPDF(obData.unitPrice, buyerData.unitPrice, comp.unitPrice, true),
                            formatForPDF(obData.totalPrice, buyerData.total, comp.total, true)
                        ]);
                    }
                }
                return rows;
            }
        };
    }
}

// Initialize global PDF exporter instance
window.pdfExporter = new PDFExporter();
