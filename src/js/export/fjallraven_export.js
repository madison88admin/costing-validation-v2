/**
 * Create Fjall Raven (V5) export configuration
 * @param {Array} fileResults - Array of file results from Fjall Raven processor
 * @returns {Object} - Configuration object for Fjall Raven export
 */
function createFjallRavenConfig(fileResults) {
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
