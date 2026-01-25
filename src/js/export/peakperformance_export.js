/**
 * Create Peak Performance (V10) export configuration
 * @param {Array} fileResults - Array of file results from Peak Performance processor
 * @returns {Object} - Configuration object for Peak Performance export
 */
function createPeakPerformanceConfig(fileResults) {
    return {
        title: 'Peak Performance Validation Results - V10',
        fileResults: fileResults.map(fileResult => {
            const fabricWastage = fileResult.results.fabricWastageCheck;
            const standardItems = fileResult.results.standardItemsCheck;

            // Count valid checks
            let validCount = 0;
            let totalChecks = 1; // Fabric wastage

            if (standardItems.found && standardItems.items) {
                totalChecks += standardItems.items.filter(i => i.found).length;
            }

            if (fabricWastage.found && fabricWastage.isValid) validCount++;
            if (standardItems.found && standardItems.items) {
                validCount += standardItems.items.filter(i => i.found && i.isValid).length;
            }

            // Build cell statuses for coloring
            const cellStatuses = [];

            // Fabric/Yarn Wastage - one row (status for col 1, values span remaining cols)
            if (fabricWastage.found) {
                cellStatuses.push([
                    'normal',
                    fabricWastage.isValid ? 'valid' : 'invalid',
                    'normal', 'normal', 'normal', 'normal', 'normal', 'normal'
                ]);
            } else {
                cellStatuses.push([
                    'normal',
                    'invalid',
                    'normal', 'normal', 'normal', 'normal', 'normal', 'normal'
                ]);
            }

            // Standard Items - one row per item with all columns
            if (standardItems.found && standardItems.items) {
                for (const item of standardItems.items) {
                    if (!item.found) {
                        // Row with 8 columns: Material Desc + 7 check columns
                        cellStatuses.push([
                            'normal', 'invalid', 'invalid', 'invalid', 'invalid', 'invalid', 'invalid', 'invalid'
                        ]);
                    } else {
                        // Material Description + check statuses for each column
                        const rowStatus = ['normal'];

                        // Add status for each check column (Supplier, Item #, Garment Part, Yield, Wastage, FOB, CIF)
                        const checkLabels = ['Supplier', 'Supplier Item #', 'Garment Part', 'Yield', 'Wastage', 'FOB Unit Cost', 'CIF Unit Cost'];
                        for (const label of checkLabels) {
                            const check = item.checks.find(c => c.label === label);
                            if (check) {
                                rowStatus.push(check.isValid ? 'valid' : 'invalid');
                            } else {
                                rowStatus.push('normal'); // No check for this column
                            }
                        }

                        cellStatuses.push(rowStatus);
                    }
                }
            }

            return {
                fileName: fileResult.fileName,
                summary: `Summary: ${validCount} out of ${totalChecks} checks passed`,
                results: fileResult.results,
                cellStatuses: cellStatuses
            };
        }),
        filenamePrefix: 'PeakPerformanceValidation_V10',
        columnWidths: [38, 22, 18, 22, 18, 18, 20, 20],
        headers: ['Material Description', 'Supplier', 'Item #', 'Garment Part', 'Yield', 'Wastage', 'FOB Cost', 'CIF Cost'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];
            const fabricWastage = fileResult.results.fabricWastageCheck;
            const standardItems = fileResult.results.standardItemsCheck;

            // Fabric/Yarn Wastage row (first row)
            if (fabricWastage.found) {
                const allCells = [];

                // Add valid cells
                fabricWastage.validCells.forEach(cell => {
                    allCells.push(cell.value);
                });

                // Add invalid cells with expected value
                fabricWastage.invalidCells.forEach(cell => {
                    allCells.push(`${cell.value} (Expected: 5%)`);
                });

                const cellsDisplay = allCells.length > 0 ? allCells.join(', ') : 'No data in range';

                rows.push([
                    'Fabric/Yarn Wastage',
                    cellsDisplay,
                    '', '', '', '', '', ''
                ]);
            } else {
                rows.push([
                    'Fabric/Yarn Wastage',
                    fabricWastage.message || 'Not found',
                    '', '', '', '', '', ''
                ]);
            }

            // Standard Items - each item as a single row with all columns
            if (standardItems.found && standardItems.items) {
                for (const item of standardItems.items) {
                    const itemName = item.standardItem.materialDesc;

                    if (!item.found) {
                        rows.push([
                            itemName,
                            'Not found in file',
                            '',
                            '',
                            '',
                            '',
                            '',
                            ''
                        ]);
                    } else {
                        // Helper function to get check by label
                        const getCheckByLabel = (label) => item.checks.find(c => c.label === label);

                        // Get each check
                        const supplierCheck = getCheckByLabel('Supplier');
                        const itemNumCheck = getCheckByLabel('Supplier Item #');
                        const partCheck = getCheckByLabel('Garment Part');
                        const yieldCheck = getCheckByLabel('Yield');
                        const wastageCheck = getCheckByLabel('Wastage');
                        const fobCheck = getCheckByLabel('FOB Unit Cost');
                        const cifCheck = getCheckByLabel('CIF Unit Cost');

                        // Format cell value
                        const formatCell = (check) => {
                            if (!check) return '-';
                            const displayValue = check.actual || '-';
                            if (check.isValid) {
                                return displayValue;
                            } else {
                                return `${displayValue} (Expected: ${check.expected})`;
                            }
                        };

                        rows.push([
                            itemName,
                            formatCell(supplierCheck),
                            formatCell(itemNumCheck),
                            formatCell(partCheck),
                            formatCell(yieldCheck),
                            formatCell(wastageCheck),
                            formatCell(fobCheck),
                            formatCell(cifCheck)
                        ]);
                    }
                }
            }

            return rows;
        }
    };
}
