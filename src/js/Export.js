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
     * Delegates to external configuration file
     * @returns {Object} - Configuration object for TNF export
     */
    createTNFConfig() {
        return createTNFConfig();
    }

    /**
     * Create Fjall Raven (V5) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from Fjall Raven processor
     * @returns {Object} - Configuration object for Fjall Raven export
     */
    createFjallRavenConfig(fileResults) {
        return createFjallRavenConfig(fileResults);
    }

    /**
     * Create Helly Hansen (V4) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from Helly Hansen processor
     * @param {Function} formatToFourDecimals - Helper function to format decimals
     * @returns {Object} - Configuration object for Helly Hansen export
     */
    createHellyHansenConfig(fileResults, formatToFourDecimals) {
        return createHellyHansenConfig(fileResults, formatToFourDecimals);
    }

    /**
     * Create Columbia (V3) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from Columbia processor
     * @param {Function} formatToThreeDecimals - Helper function to format decimals
     * @returns {Object} - Configuration object for Columbia export
     */
    createColumbiaConfig(fileResults, formatToThreeDecimals) {
        return createColumbiaConfig(fileResults, formatToThreeDecimals);
    }

    /**
     * Create Outdoor Research (V8) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from Outdoor Research processor
     * @returns {Object} - Configuration object for Outdoor Research export
     */
    createOutdoorResearchConfig(fileResults) {
        return createOutdoorResearchConfig(fileResults);
    }

    /**
     * Create Mammut (V7) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from Mammut processor
     * @returns {Object} - Configuration object for Mammut export
     */
    createMammutConfig(fileResults) {
        return createMammutConfig(fileResults);
    }

    /**
     * Create LLBEAN (V6) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from LLBEAN processor
     * @returns {Object} - Configuration object for LLBEAN export
     */
    createLLBEANConfig(fileResults) {
        return createLLBEANConfig(fileResults);
    }

    /**
     * Create Burton (V2) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from Burton processor
     * @param {Function} formatToThreeDecimals - Helper function to format decimals
     * @returns {Object} - Configuration object for Burton export
     */
    createBurtonConfig(fileResults, formatToThreeDecimals) {
        return createBurtonConfig(fileResults, formatToThreeDecimals);
    }

    /**
     * Create On AG (V9) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from On AG processor
     * @returns {Object} - Configuration object for On AG export
     */
    createOnAGConfig(fileResults) {
        return createOnAGConfig(fileResults);
    }

    /**
     * Create Peak Performance (V10) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from Peak Performance processor
     * @returns {Object} - Configuration object for Peak Performance export
     */
    createPeakPerformanceConfig(fileResults) {
        return createPeakPerformanceConfig(fileResults);
    }

    /**
     * Create Skida (V11) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from Skida processor
     * @returns {Object} - Configuration object for Skida export
     */
    createSkidaConfig(fileResults) {
        return createSkidaConfig(fileResults);
    }

    /**
     * Create Vuori (V12) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from Vuori processor
     * @returns {Object} - Configuration object for Vuori export
     */
    createVuoriConfig(fileResults) {
        return createVuoriConfig(fileResults);
    }

    /**
     * Create Prana (V13) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from Prana processor
     * @returns {Object} - Configuration object for Prana export
     */
    createPranaConfig(fileResults) {
        return createPranaConfig(fileResults);
    }

    /**
     * Create Travis Matthew (V14) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from Travis Matthew processor
     * @returns {Object} - Configuration object for Travis Matthew export
     */
    createTravisMatthewConfig(fileResults) {
        return createTravisMatthewConfig(fileResults);
    }

    /**
     * Create Jack Wolfskin (V15) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from Jack Wolfskin processor
     * @returns {Object} - Configuration object for Jack Wolfskin export
     */
    createJackWolfskinConfig(fileResults) {
        return createJackWolfskinConfig(fileResults);
    }

    /**
     * Create 511 (V16) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from 511 processor
     * @returns {Object} - Configuration object for 511 export
     */
    create511Config(fileResults) {
        return create511Config(fileResults);
    }

    /**
     * Create Ride Store (V17) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from Ride Store processor
     * @returns {Object} - Configuration object for Ride Store export
     */
    createRideStoreConfig(fileResults) {
        return createRideStoreConfig(fileResults);
    }

    /**
     * Create Foot Asylum (V18) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from Foot Asylum processor
     * @returns {Object} - Configuration object for Foot Asylum export
     */
    createFootAsylumConfig(fileResults) {
        return createFootAsylumConfig(fileResults);
    }

    /**
     * Create KUHL (V19) export configuration
     * Delegates to external configuration file
     * @param {Array} fileResults - Array of file results from KUHL processor
     * @returns {Object} - Configuration object for KUHL export
     */
    createKuhlConfig(fileResults) {
        return createKuhlConfig(fileResults);
    }
}

// Initialize global PDF exporter instance
window.pdfExporter = new PDFExporter();
