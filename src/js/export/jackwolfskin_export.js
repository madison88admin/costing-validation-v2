/**
 * Create Jack Wolfskin (V15) export configuration
 * @param {Array} fileResults - Array of file results from Jack Wolfskin processor
 * @returns {Object} - Configuration object for Jack Wolfskin export
 */
function createJackWolfskinConfig(fileResults) {
    return {
        title: 'Jack Wolfskin Validation Results - V15',
        fileResults: fileResults.map(fileResult => {
            const totalItems = fileResult.results.length;
            const validItems = fileResult.results.filter(r => r.isValid).length;

            // Build cell statuses for coloring
            const cellStatuses = fileResult.results.map(item => {
                return [
                    'normal', // Sheet Name
                    'normal', // Field name
                    item.isValid ? 'valid' : 'invalid', // BCBD Value
                    item.isValid ? 'valid' : 'invalid' // Status
                ];
            });

            return {
                fileName: fileResult.fileName,
                sheetName: fileResult.sheetName,
                summary: `Sheet: ${fileResult.sheetName} | Summary: ${validItems} out of ${totalItems} validations passed`,
                results: fileResult.results,
                cellStatuses: cellStatuses
            };
        }),
        filenamePrefix: 'JackWolfskinValidation_V15',
        columnWidths: [70, 60, 100, 50],
        headers: ['Sheet Name', 'Field', 'BCBD Value', 'Status'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];
            for (const item of fileResult.results) {
                const statusIcon = item.isValid ? '✓' : '✗';
                const statusText = item.isValid ? 'VALID' : 'INVALID';

                // Format BCBD value for PDF
                let bcbdValue = '';
                if (!item.found) {
                    bcbdValue = `Not Found (Expected: ${item.expectedValue})`;
                } else if (item.actualValue === '' || item.actualValue === null) {
                    bcbdValue = `Empty (Expected: ${item.expectedValue})`;
                } else if (item.isValid) {
                    bcbdValue = item.actualValue;
                } else {
                    bcbdValue = `${item.actualValue} (Expected: ${item.expectedValue})`;
                }

                rows.push([
                    fileResult.sheetName,
                    item.label,
                    bcbdValue,
                    `${statusIcon} ${statusText}`
                ]);
            }
            return rows;
        }
    };
}
