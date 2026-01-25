/**
 * Create Skida (V11) export configuration
 * @param {Array} fileResults - Array of file results from Skida processor
 * @returns {Object} - Configuration object for Skida export
 */
function createSkidaConfig(fileResults) {
    return {
        title: 'Skida Validation Results - V11',
        fileResults: fileResults.map(fileResult => {
            const items = fileResult.items || [];

            // Count valid checks
            const totalItems = items.length;
            const validItems = items.filter(item => item.found && item.isValid).length;

            // Build cell statuses for coloring - each item is one row with 4 columns
            const cellStatuses = [];

            for (const item of items) {
                if (!item.found) {
                    // Item not found - mark all columns as invalid
                    cellStatuses.push([
                        'normal', 'invalid', 'invalid', 'invalid'
                    ]);
                } else {
                    // Item found - check each column
                    const getCheckByLabel = (label) => item.checks.find(c => c.label === label);

                    const descriptionCheck = getCheckByLabel('Description');
                    const unitCostCheck = getCheckByLabel('Unit Cost');
                    const quantityCheck = getCheckByLabel('Quantity');

                    cellStatuses.push([
                        'normal', // Item Name column
                        descriptionCheck?.isValid ? 'valid' : 'invalid',
                        unitCostCheck?.isValid ? 'valid' : 'invalid',
                        quantityCheck?.isValid ? 'valid' : 'invalid'
                    ]);
                }
            }

            return {
                fileName: fileResult.fileName,
                summary: `Summary: ${validItems} out of ${totalItems} items fully match`,
                results: fileResult.items,
                cellStatuses: cellStatuses
            };
        }),
        filenamePrefix: 'SkidaValidation_V11',
        columnWidths: [35, 45, 35, 35],
        headers: ['Item Name', 'Description', 'Unit Cost', 'Quantity'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];
            const items = fileResult.results || [];

            for (const item of items) {
                const itemName = item.validationItem.category;
                const expectedDesc = item.validationItem.description;
                const expectedCost = item.validationItem.unitCost;
                const expectedQty = item.validationItem.quantity;

                if (!item.found) {
                    rows.push([
                        itemName,
                        'Not found in Buyer CBD file',
                        '',
                        ''
                    ]);
                } else {
                    const getCheckByLabel = (label) => item.checks.find(c => c.label === label);

                    const descriptionCheck = getCheckByLabel('Description');
                    const unitCostCheck = getCheckByLabel('Unit Cost');
                    const quantityCheck = getCheckByLabel('Quantity');

                    // Format cell value - add expected if not valid
                    const formatCell = (check, expected) => {
                        if (!check) return '-';
                        const actualValue = check.actual || '-';

                        // Handle empty values
                        if (!actualValue || actualValue === '' || actualValue === '-') {
                            if (expected === '-') {
                                return '-';
                            }
                            return `Empty (Expected: ${expected})`;
                        }

                        if (check.isValid) {
                            return actualValue;
                        } else {
                            return `${actualValue} (Expected: ${expected})`;
                        }
                    };

                    rows.push([
                        itemName,
                        formatCell(descriptionCheck, expectedDesc),
                        formatCell(unitCostCheck, expectedCost),
                        formatCell(quantityCheck, expectedQty)
                    ]);
                }
            }

            return rows;
        }
    };
}
