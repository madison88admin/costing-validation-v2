/**
 * Create Rossignol (V23) export configuration
 * @param {Array} fileResults - Array of file results from Rossignol processor
 * @returns {Object} - Configuration object for Rossignol export
 */
function createRossignolConfig(fileResults) {
    return {
        title: 'Rossignol Validation Results - V23',
        fileResults: fileResults.map(fileResult => {
            const totalChecks = fileResult.checks?.length || 0;
            const validChecks = fileResult.checks?.filter(c => c.isValid).length || 0;
            const foundChecks = fileResult.checks?.filter(c => c.found).length || 0;

            return {
                fileName: fileResult.fileName,
                summary: `Summary: ${validChecks} out of ${foundChecks} found checks are valid (${totalChecks} total rules)`,
                results: fileResult.checks || [],
                sheetName: fileResult.sheetName
            };
        }),
        filenamePrefix: 'RossignolValidation_V23',
        columnWidths: [35, 90, 35],
        headers: ['Check Name', 'Value', 'Expected'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];
            const cellStatuses = [];
            const checks = fileResult.results || [];

            for (const check of checks) {
                let valueText = '';
                let expectedText = check.expected;
                let valueStatus = 'normal';

                if (check.isPackagingCombined && check.packagingEntries) {
                    // PACKAGING rows - combine all entries
                    const lines = check.packagingEntries.map(entry => {
                        if (entry.hasGenericPackaging) {
                            // Show all validated cells
                            const cellValues = entry.cellValidations.map(cv => cv.value).join(' | ');
                            return `PACKAGING ${entry.sequence}: ${cellValues}`;
                        } else {
                            // No Generic Packaging - show Column L value
                            return `PACKAGING ${entry.sequence} ${entry.value}`;
                        }
                    });
                    valueText = lines.join('\n');
                    valueStatus = check.isValid ? 'valid' : 'invalid';
                } else if (check.isFabricCombined && check.fabricEntries) {
                    // Category rows (FABRIC, TRIM, etc.) - combine all entries
                    const catName = check.categoryName || 'FABRIC';
                    const lines = check.fabricEntries.map(entry => {
                        return `${catName} ${entry.sequence} ${entry.value}`;
                    });
                    valueText = lines.join('\n');
                    valueStatus = check.isValid ? 'valid' : 'invalid';
                } else if (!check.found) {
                    valueText = 'Not Found';
                    valueStatus = 'warning';
                } else {
                    valueText = check.actual;
                    valueStatus = check.isValid ? 'valid' : 'invalid';
                }

                rows.push([
                    check.name,
                    valueText,
                    expectedText
                ]);

                cellStatuses.push(['normal', valueStatus, 'normal']);
            }

            return { rows, cellStatuses };
        }
    };
}
