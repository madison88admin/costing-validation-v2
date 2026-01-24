/**
 * Create Columbia (V3) export configuration
 * @param {Array} fileResults - Array of file results from Columbia processor
 * @param {Function} formatToThreeDecimals - Helper function to format decimals
 * @returns {Object} - Configuration object for Columbia export
 */
function createColumbiaConfig(fileResults, formatToThreeDecimals) {
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
