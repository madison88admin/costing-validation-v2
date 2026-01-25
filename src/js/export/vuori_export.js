/**
 * Create Vuori (V12) export configuration
 * @param {Array} fileResults - Array of file results from Vuori processor
 * @returns {Object} - Configuration object for Vuori export
 */
function createVuoriConfig(fileResults) {
    return {
        title: 'Vuori Validation Results - V12',
        fileResults: fileResults.map(fileResult => {
            const items = fileResult.items || [];

            // Count valid checks
            const totalItems = items.length;
            const validItems = items.filter(item => item.found && item.isValid).length;

            // Build cell statuses for coloring - each item is one row with 6 columns
            const cellStatuses = [];

            for (const item of items) {
                if (!item.found) {
                    // Item not found - mark all columns as invalid
                    cellStatuses.push([
                        'invalid', 'invalid', 'invalid', 'invalid', 'invalid', 'invalid'
                    ]);
                } else {
                    // Item found - check each column
                    const getCheckByLabel = (label) => item.checks.find(c => c.label === label);

                    const materialDescCheck = getCheckByLabel('Material Desc');
                    const materialCodeCheck = getCheckByLabel('Material Code');
                    const materialSubtypeCheck = getCheckByLabel('Material Subtype');
                    const constructionCheck = getCheckByLabel('Construction');
                    const supplierCostCheck = getCheckByLabel('Supplier Cost');
                    const wastageCheck = getCheckByLabel('Wastage');

                    // Use 'normal' for skipped fields (expected is '-')
                    cellStatuses.push([
                        materialDescCheck?.skipped ? 'normal' : (materialDescCheck?.isValid ? 'valid' : 'invalid'),
                        materialCodeCheck?.skipped ? 'normal' : (materialCodeCheck?.isValid ? 'valid' : 'invalid'),
                        materialSubtypeCheck?.skipped ? 'normal' : (materialSubtypeCheck?.isValid ? 'valid' : 'invalid'),
                        constructionCheck?.skipped ? 'normal' : (constructionCheck?.isValid ? 'valid' : 'invalid'),
                        supplierCostCheck?.skipped ? 'normal' : (supplierCostCheck?.isValid ? 'valid' : 'invalid'),
                        wastageCheck?.skipped ? 'normal' : (wastageCheck?.isValid ? 'valid' : 'invalid')
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
        filenamePrefix: 'VuoriValidation_V12',
        columnWidths: [28, 25, 28, 25, 25, 25],
        headers: ['Material Desc (D)', 'Material Code (E)', 'Material Subtype (F)', 'Construction (G)', 'Supplier Cost (S)', 'Wastage (W)'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];
            const items = fileResult.results || [];

            for (const item of items) {
                const expectedMaterialDesc = item.validationItem.materialDesc;
                const expectedMaterialCode = item.validationItem.materialCode;
                const expectedMaterialSubtype = item.validationItem.materialSubtype;
                const expectedConstruction = item.validationItem.construction;
                const expectedSupplierCost = item.validationItem.supplierCost;
                const expectedWastage = item.validationItem.wastage;

                if (!item.found) {
                    rows.push([
                        expectedMaterialDesc,
                        expectedMaterialCode,
                        expectedMaterialSubtype,
                        'Not found in Buyer CBD file',
                        '',
                        ''
                    ]);
                } else {
                    const getCheckByLabel = (label) => item.checks.find(c => c.label === label);

                    const materialDescCheck = getCheckByLabel('Material Desc');
                    const materialCodeCheck = getCheckByLabel('Material Code');
                    const materialSubtypeCheck = getCheckByLabel('Material Subtype');
                    const constructionCheck = getCheckByLabel('Construction');
                    const supplierCostCheck = getCheckByLabel('Supplier Cost');
                    const wastageCheck = getCheckByLabel('Wastage');

                    // Format cell value - add expected if not valid
                    const formatCell = (check, expected) => {
                        if (!check) return '-';

                        // If expected is '-', field is skipped
                        if (expected === '-' || check.skipped) {
                            return '-';
                        }

                        const actualValue = check.actual || '-';

                        // Handle empty values
                        if (!actualValue || actualValue === '' || actualValue === '-') {
                            return `Empty (Expected: ${expected})`;
                        }

                        if (check.isValid) {
                            return actualValue;
                        } else {
                            return `${actualValue} (Expected: ${expected})`;
                        }
                    };

                    rows.push([
                        formatCell(materialDescCheck, expectedMaterialDesc),
                        formatCell(materialCodeCheck, expectedMaterialCode),
                        formatCell(materialSubtypeCheck, expectedMaterialSubtype),
                        formatCell(constructionCheck, expectedConstruction),
                        formatCell(supplierCostCheck, expectedSupplierCost),
                        formatCell(wastageCheck, expectedWastage)
                    ]);
                }
            }

            return rows;
        }
    };
}
