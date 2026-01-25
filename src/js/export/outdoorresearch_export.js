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

            // Build cell statuses for coloring (one row per check)
            const cellStatuses = [];

            // General Packaging checks - each check gets its own row
            if (gpCheck && gpCheck.found && gpCheck.checks) {
                gpCheck.checks.forEach(check => {
                    cellStatuses.push([
                        'normal',
                        check.isValid ? 'valid' : 'invalid'
                    ]);
                });
            } else {
                cellStatuses.push([
                    'normal',
                    'invalid'
                ]);
            }

            // Other Charges checks - each check gets its own row
            if (ocCheck && ocCheck.found && ocCheck.checks) {
                ocCheck.checks.forEach(check => {
                    cellStatuses.push([
                        'normal',
                        check.isValid ? 'valid' : 'invalid'
                    ]);
                });
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
        headers: ['Validation Check', 'Cost'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];
            const gpCheck = fileResult.results.generalPackagingCheck;
            const ocCheck = fileResult.results.otherChargesCheck;

            // General Packaging checks - each check gets its own row
            if (gpCheck && gpCheck.found && gpCheck.checks) {
                gpCheck.checks.forEach((check, index) => {
                    const rowLabel = index === 0
                        ? `General Packaging (Row ${gpCheck.rowNumber})`
                        : '';
                    const valueDisplay = check.isValid
                        ? `${check.label}: ${check.actualValue || 'Empty'}`
                        : `${check.label}: ${check.actualValue || 'Empty'} (Expected: ${check.expectedValue})`;
                    rows.push([
                        rowLabel,
                        valueDisplay
                    ]);
                });
            } else if (gpCheck && !gpCheck.found) {
                rows.push([
                    'General Packaging',
                    gpCheck.message || 'Not found'
                ]);
            }

            // Other Charges checks - each check gets its own row
            if (ocCheck && ocCheck.found && ocCheck.checks) {
                ocCheck.checks.forEach((check, index) => {
                    const rowLabel = index === 0
                        ? `Other Charges (Row ${ocCheck.rowNumber})`
                        : '';
                    const valueDisplay = check.isValid
                        ? `${check.label}: ${check.actualValue || 'Empty'}`
                        : `${check.label}: ${check.actualValue || 'Empty'} (Expected: ${check.expectedValue})`;
                    rows.push([
                        rowLabel,
                        valueDisplay
                    ]);
                });
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
