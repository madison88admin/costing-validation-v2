/**
 * Create Outdoor Research (V8) export configuration
 * @param {Array} fileResults - Array of file results from Outdoor Research processor
 * @returns {Object} - Configuration object for Outdoor Research export
 */
function createOutdoorResearchConfig(fileResults) {
    return {
        title: 'Outdoor Research Validation Results - V8',
        fileResults: fileResults.map(fileResult => {
            const gpCheck = fileResult.results.generalPackagingCheck;
            const ocCheck = fileResult.results.otherChargesCheck;

            // Count valid checks
            let validCount = 0;
            let totalChecks = 0;

            if (gpCheck && gpCheck.found && gpCheck.checks) {
                totalChecks++;
                if (gpCheck.isValid) validCount++;
            }
            if (ocCheck && ocCheck.found && ocCheck.checks) {
                totalChecks++;
                if (ocCheck.isValid) validCount++;
            }

            // Build cell statuses for coloring
            const cellStatuses = [];

            // General Packaging row
            if (gpCheck && gpCheck.found && gpCheck.checks) {
                cellStatuses.push([
                    'normal',
                    gpCheck.isValid ? 'valid' : 'invalid'
                ]);
            } else {
                cellStatuses.push([
                    'normal',
                    'invalid'
                ]);
            }

            // Other Charges row
            if (ocCheck && ocCheck.found && ocCheck.checks) {
                cellStatuses.push([
                    'normal',
                    ocCheck.isValid ? 'valid' : 'invalid'
                ]);
            } else {
                cellStatuses.push([
                    'normal',
                    'invalid'
                ]);
            }

            return {
                fileName: fileResult.fileName,
                summary: `Summary: ${validCount} out of ${totalChecks} checks passed`,
                results: fileResult.results,
                cellStatuses: cellStatuses
            };
        }),
        filenamePrefix: 'OutdoorResearchValidation_V8',
        columnWidths: [45, 105],
        headers: ['Validation Check', 'Value'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];
            const gpCheck = fileResult.results.generalPackagingCheck;
            const ocCheck = fileResult.results.otherChargesCheck;

            // General Packaging row
            if (gpCheck && gpCheck.found && gpCheck.checks) {
                const checkDetails = gpCheck.checks.map(check => {
                    return `${check.label}: ${check.actualValue || 'Empty'}`;
                }).join(' | ');

                rows.push([
                    `General Packaging (Row ${gpCheck.rowNumber})`,
                    checkDetails
                ]);
            } else if (gpCheck && !gpCheck.found) {
                rows.push([
                    'General Packaging',
                    gpCheck.message || 'Not found'
                ]);
            }

            // Other Charges row
            if (ocCheck && ocCheck.found && ocCheck.checks) {
                const checkDetails = ocCheck.checks.map(check => {
                    return `${check.label}: ${check.actualValue || 'Empty'}`;
                }).join(' | ');

                rows.push([
                    `Other Charges (Row ${ocCheck.rowNumber})`,
                    checkDetails
                ]);
            } else if (ocCheck && !ocCheck.found) {
                rows.push([
                    'Other Charges',
                    ocCheck.message || 'Not found'
                ]);
            }

            return rows;
        }
    };
}
