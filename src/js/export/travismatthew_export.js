/**
 * Create Travis Matthew (V14) export configuration
 * @param {Array} fileResults - Array of file results from Travis Matthew processor
 * @returns {Object} - Configuration object for Travis Matthew export
 */
function createTravisMatthewConfig(fileResults) {
    return {
        title: 'Travis Matthew Cost Breakdown Comparison - V14',
        fileResults: fileResults.map(fileResult => {
            const totalItems = fileResult.results.length;
            const validItems = fileResult.results.filter(r => r.isValid).length;

            // Build cell statuses for coloring
            const cellStatuses = fileResult.results.map(item => {
                return [
                    'normal', // Field name
                    item.isValid ? 'valid' : 'invalid', // BCBD Value
                    item.isValid ? 'valid' : 'invalid' // Status
                ];
            });

            return {
                fileName: fileResult.fileName,
                summary: `Summary: ${validItems} out of ${totalItems} items match the OB file`,
                results: fileResult.results,
                cellStatuses: cellStatuses
            };
        }),
        filenamePrefix: 'TravisMatthewCostBreakdown_V14',
        columnWidths: [60, 120, 50],
        headers: ['Field', 'BCBD Value', 'Status'],
        colorRules: {
            // Generic color rules for Travis Matthew
        },
        extractRowData: (fileResult) => {
            const rows = [];
            for (const item of fileResult.results) {
                const statusIcon = item.isValid ? '✓' : '✗';
                const statusText = item.isValid ? 'VALID' : 'INVALID';

                // Format BCBD value for PDF
                let bcbdValue = '';
                if (item.buyerValue === null || item.buyerValue === '') {
                    bcbdValue = `Not Found (Expected: ${item.obValue})`;
                } else if (item.isValid) {
                    bcbdValue = item.buyerValue;
                } else {
                    bcbdValue = `${item.buyerValue} (Expected: ${item.obValue})`;
                }

                rows.push([
                    item.label,
                    bcbdValue,
                    `${statusIcon} ${statusText}`
                ]);
            }
            return rows;
        }
    };
}
