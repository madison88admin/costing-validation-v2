/**
 * Create Ride Store (V17) export configuration
 * @param {Array} fileResults - Array of file results from Ride Store processor
 * @returns {Object} - Configuration object for Ride Store export
 */
function createRideStoreConfig(fileResults) {
    return {
        title: 'Ride Store Validation Results - V17',
        fileResults: fileResults.map(fileResult => {
            if (fileResult.error) {
                return {
                    fileName: fileResult.fileName,
                    summary: `Error: ${fileResult.error}`,
                    cellResults: [],
                    sectionResults: [],
                    cellStatuses: []
                };
            }

            const cellTotalCount = fileResult.cellResults.length;
            const cellValidCount = fileResult.cellResults.filter(r => r.isValid).length;
            const sectionTotalCount = fileResult.sectionResults.length;
            const sectionValidCount = fileResult.sectionResults.filter(r => r.isValid).length;

            // Build cell statuses for coloring
            const cellStatuses = fileResult.cellResults.map(item => {
                return [
                    'normal', // Field name
                    'normal', // Cell reference
                    item.isValid ? 'valid' : 'invalid', // BCBD Value
                    item.isValid ? 'valid' : 'invalid' // Status
                ];
            });

            // Add separator row status
            cellStatuses.push(['normal', 'normal', 'normal', 'normal']);

            // Add wastage section statuses (Section column is normal, Wastage column shows validity)
            for (const section of fileResult.sectionResults) {
                cellStatuses.push([
                    'normal', // Section name
                    section.isValid ? 'valid' : 'invalid', // Wastage cells
                    'normal', // Empty column
                    'normal'  // Empty column
                ]);
            }

            return {
                fileName: fileResult.fileName,
                summary: `Cell Validations: ${cellValidCount}/${cellTotalCount} passed | Section Validations: ${sectionValidCount}/${sectionTotalCount} passed`,
                cellResults: fileResult.cellResults,
                sectionResults: fileResult.sectionResults,
                cellStatuses: cellStatuses
            };
        }),
        filenamePrefix: 'RideStore_Validation',
        columnWidths: [50, 30, 120, 40],
        headers: ['Field', 'Cell', 'BCBD Value', 'Status'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];

            if (fileResult.error) {
                rows.push(['Error', fileResult.error, '', '']);
                return rows;
            }

            // Add cell validation rows
            for (const item of fileResult.cellResults) {
                const statusIcon = item.isValid ? '✓' : '✗';
                const statusText = item.isValid ? 'VALID' : 'INVALID';

                // Format BCBD value for PDF
                let bcbdValue = '';
                if (item.actual === '' || item.actual === null || item.actual === 'Empty') {
                    bcbdValue = `Empty (Expected: ${item.expected})`;
                } else if (item.isValid) {
                    bcbdValue = item.actual;
                } else {
                    bcbdValue = `${item.actual} (Expected: ${item.expected})`;
                }

                rows.push([
                    item.label,
                    item.cell,
                    bcbdValue,
                    `${statusIcon} ${statusText}`
                ]);
            }

            // Add separator for wastage section
            rows.push(['--- Wastage% Validation ---', '', '', '']);

            // Add wastage section rows
            if (fileResult.sectionResults && fileResult.sectionResults.length > 0) {
                for (const section of fileResult.sectionResults) {
                    let wastageDetails = '';

                    if (!section.sectionFound) {
                        wastageDetails = 'Section not found in file';
                    } else {
                        // Valid cells - just cell references
                        if (section.validCells.length > 0) {
                            const validCellRefs = section.validCells.map(c => c.cell).join(', ');
                            wastageDetails += validCellRefs;
                        }

                        // Invalid cells - cell references with values
                        if (section.invalidCells.length > 0) {
                            if (section.validCells.length > 0) {
                                wastageDetails += '\n';
                            }
                            const invalidCellRefs = section.invalidCells.map(c => `${c.cell}: ${c.value}`).join(', ');
                            wastageDetails += invalidCellRefs;
                        }

                        if (section.validCells.length === 0 && section.invalidCells.length === 0) {
                            wastageDetails = 'No items found in section';
                        }
                    }

                    rows.push([
                        `${section.label}\nExpected: ${section.expectedValue}`,
                        wastageDetails,
                        '',
                        ''
                    ]);
                }
            }

            return rows;
        }
    };
}
