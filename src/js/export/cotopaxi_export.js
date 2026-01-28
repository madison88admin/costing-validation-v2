/**
 * Create Cotopaxi (V24) export configuration
 * @param {Array} fileResults - Array of file results from Cotopaxi processor
 * @returns {Object} - Configuration object for Cotopaxi export
 */
function createCotopaxiConfig(fileResults) {
    return {
        title: 'Cotopaxi Validation Results - V24',
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
        filenamePrefix: 'CotopaxiValidation_V24',
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

                if (!check.found) {
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
