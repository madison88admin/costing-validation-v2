/**
 * Create ODLO (V22) export configuration
 * @param {Array} fileResults - Array of file results from ODLO processor
 * @returns {Object} - Configuration object for ODLO export
 */
function createODLOConfig(fileResults) {
    return {
        title: 'ODLO Validation Results - V22',
        fileResults: fileResults.map(fileResult => {
            const totalChecks = fileResult.checks?.length || 0;
            const validChecks = fileResult.checks?.filter(c => c.isValid).length || 0;
            const foundChecks = fileResult.checks?.filter(c => c.found).length || 0;

            // Build cell statuses for coloring
            const cellStatuses = (fileResult.checks || []).map(check => {
                if (!check.found) {
                    return ['normal', 'normal', 'normal', 'normal', 'warning']; // Not found = warning
                } else if (check.isValid) {
                    return ['normal', 'normal', 'valid', 'normal', 'valid']; // Valid
                } else {
                    return ['normal', 'normal', 'invalid', 'normal', 'invalid']; // Invalid
                }
            });

            return {
                fileName: fileResult.fileName,
                summary: `Summary: ${validChecks} out of ${foundChecks} found checks are valid (${totalChecks} total rules)`,
                results: fileResult.checks || [],
                cellStatuses: cellStatuses,
                sheetName: fileResult.sheetName
            };
        }),
        filenamePrefix: 'ODLOValidation_V22',
        columnWidths: [40, 40, 40, 20, 20],
        headers: ['Check Name', 'Expected Value', 'Actual Value', 'Location', 'Status'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];
            const checks = fileResult.results || [];

            for (const check of checks) {
                let status = '';
                if (!check.found) {
                    status = 'Not Found';
                } else if (check.isValid) {
                    status = '✓ Valid';
                } else {
                    status = '✗ Invalid';
                }

                const location = check.found ? `${check.checkColumn}${check.rowNumber}` : '-';

                rows.push([
                    check.name,
                    check.expected,
                    check.actual,
                    location,
                    status
                ]);
            }

            return rows;
        }
    };
}
