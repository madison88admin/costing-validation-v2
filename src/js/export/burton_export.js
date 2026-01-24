/**
 * Create Burton (V2) export configuration
 * @param {Array} fileResults - Array of file results from Burton processor
 * @param {Function} formatToThreeDecimals - Helper function to format decimals
 * @returns {Object} - Configuration object for Burton export
 */
function createBurtonConfig(fileResults, formatToThreeDecimals) {
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
