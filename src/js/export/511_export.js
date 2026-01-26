/**
 * Create 511 (V16) export configuration
 * @param {Array} fileResults - Array of file results from 511 processor
 * @returns {Object} - Configuration object for 511 export
 */
function create511Config(fileResults) {
    return {
        title: '511 Validation Results - V16',
        fileResults: fileResults.map(fileResult => {
            const cellTotalCount = fileResult.results.length;
            const cellValidCount = fileResult.results.filter(r => r.isValid).length;
            const sectionTotalCount = fileResult.sectionResults.length;
            const sectionValidCount = fileResult.sectionResults.filter(r => r.isValid).length;

            // Build cell statuses for coloring
            const cellStatuses = fileResult.results.map(item => {
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
                results: fileResult.results,
                sectionResults: fileResult.sectionResults,
                cellStatuses: cellStatuses
            };
        }),
        filenamePrefix: '511Validation_V16',
        columnWidths: [50, 30, 120, 40],
        headers: ['Field', 'Cell', 'BCBD Value', 'Status'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];

            // Add cell validation rows (FACTORY, COO, Remarks)
            for (const item of fileResult.results) {
                const statusIcon = item.isValid ? '✓' : '✗';
                const statusText = item.isValid ? 'VALID' : 'INVALID';

                // Format BCBD value for PDF
                let bcbdValue = '';
                if (item.actualValue === '' || item.actualValue === null) {
                    bcbdValue = `Empty (Expected: ${item.expectedValue})`;
                } else if (item.isValid) {
                    bcbdValue = item.actualValue;
                } else {
                    bcbdValue = `${item.actualValue} (Expected: ${item.expectedValue})`;
                }

                rows.push([
                    item.label,
                    item.valueCell,
                    bcbdValue,
                    `${statusIcon} ${statusText}`
                ]);
            }

            // Add separator for wastage section (matches display format)
            rows.push(['--- Wastage% Validation (Column J) ---', '', '', '']);

            // Add wastage section rows (Prana-style format)
            if (fileResult.sectionResults && fileResult.sectionResults.length > 0) {
                for (const section of fileResult.sectionResults) {
                    let wastageDetails = '';

                    if (!section.sectionFound) {
                        wastageDetails = 'Section not found in file';
                    } else {
                        // Valid cells - just cell references (shown in green in display)
                        if (section.validCells.length > 0) {
                            const validCellRefs = section.validCells.map(c => c.cell).join(', ');
                            wastageDetails += validCellRefs;
                        }

                        // Invalid cells - cell references with values (shown in red in display)
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

                    // Format matches display: Section name with expected value, then cell refs
                    rows.push([
                        `${section.label}\nExpected: ${section.expectedValue}`,
                        wastageDetails,
                        '',
                        ''
                    ]);
                }
            }

            return rows;
        },
        // Custom cell statuses that include section results
        getCellStatuses: (fileResult) => {
            const statuses = [];

            // Cell validation statuses
            for (const item of fileResult.results) {
                statuses.push([
                    'normal',
                    'normal',
                    item.isValid ? 'valid' : 'invalid',
                    item.isValid ? 'valid' : 'invalid'
                ]);
            }

            // Separator row
            statuses.push(['normal', 'normal', 'normal', 'normal']);

            // Section validation statuses
            if (fileResult.sectionResults && fileResult.sectionResults.length > 0) {
                for (const section of fileResult.sectionResults) {
                    statuses.push([
                        'normal',
                        section.isValid ? 'valid' : 'invalid',
                        'normal',
                        'normal'
                    ]);
                }
            }

            return statuses;
        }
    };
}
