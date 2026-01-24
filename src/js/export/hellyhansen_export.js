/**
 * Create Helly Hansen (V4) export configuration
 * @param {Array} fileResults - Array of file results from Helly Hansen processor
 * @param {Function} formatToFourDecimals - Helper function to format decimals
 * @returns {Object} - Configuration object for Helly Hansen export
 */
function createHellyHansenConfig(fileResults, formatToFourDecimals) {
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
