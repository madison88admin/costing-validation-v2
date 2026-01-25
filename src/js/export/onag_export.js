/**
 * Create On AG (V9) export configuration
 * @param {Array} fileResults - Array of file results from On AG processor
 * @returns {Object} - Configuration object for On AG export
 */
function createOnAGConfig(fileResults) {
    return {
        title: 'On AG Validation Results - V9',
        fileResults: fileResults.map(fileResult => {
            const wastageResults = fileResult.results.wastageResults || [];
            const coatsThreadCheck = fileResult.results.coatsThreadCheck;
            const processCostsCheck = fileResult.results.processCostsCheck;

            // Count valid sections
            let validSections = 0;
            let totalSections = wastageResults.length;

            // Add Coats Thread to count if found
            if (coatsThreadCheck && coatsThreadCheck.found) {
                totalSections++;
                if (coatsThreadCheck.isValid) validSections++;
            }

            // Add Process Costs to count
            if (processCostsCheck && processCostsCheck.items) {
                totalSections += processCostsCheck.items.length;
                processCostsCheck.items.forEach(item => {
                    if (item.isValid) validSections++;
                });
            }

            for (const section of wastageResults) {
                if (section.found && section.isValid) validSections++;
            }

            // Build cell statuses for coloring (one row per check)
            const cellStatuses = [];

            // Wastage sections - one row per section (combined values)
            for (const section of wastageResults) {
                if (section.found) {
                    // If section has any invalid cells, mark as invalid, otherwise valid
                    cellStatuses.push([
                        'normal',
                        section.isValid ? 'valid' : 'invalid'
                    ]);
                } else {
                    cellStatuses.push([
                        'normal',
                        'invalid'
                    ]);
                }
            }

            // Coats Thread checks - each check gets its own row
            if (coatsThreadCheck && coatsThreadCheck.found && coatsThreadCheck.checks) {
                coatsThreadCheck.checks.forEach(check => {
                    cellStatuses.push([
                        'normal',
                        check.isValid ? 'valid' : 'invalid'
                    ]);
                });
            } else if (coatsThreadCheck && !coatsThreadCheck.found) {
                cellStatuses.push([
                    'normal',
                    'invalid'
                ]);
            }

            // Process Costs - each item gets its own row
            if (processCostsCheck && processCostsCheck.items) {
                processCostsCheck.items.forEach(item => {
                    cellStatuses.push([
                        'normal',
                        item.found && item.isValid ? 'valid' : 'invalid'
                    ]);
                });
            }

            return {
                fileName: fileResult.fileName,
                summary: `Summary: ${validSections} out of ${totalSections} sections passed`,
                results: fileResult.results,
                cellStatuses: cellStatuses
            };
        }),
        filenamePrefix: 'OnAGValidation_V9',
        columnWidths: [45, 105],
        headers: ['Validation Check', 'Cost'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];
            const wastageResults = fileResult.results.wastageResults || [];
            const coatsThreadCheck = fileResult.results.coatsThreadCheck;
            const processCostsCheck = fileResult.results.processCostsCheck;

            // Wastage sections - one row per section (show all values)
            for (const section of wastageResults) {
                if (!section.found) {
                    rows.push([
                        `${section.section} Wastage`,
                        section.message || 'Not found'
                    ]);
                } else {
                    const allCells = [];

                    // Add valid cells
                    section.validCells.forEach(cell => {
                        allCells.push(cell.numericValue.toFixed(2));
                    });

                    // Add invalid cells with expected value
                    section.invalidCells.forEach(cell => {
                        allCells.push(`${cell.numericValue.toFixed(2)} (Expected: ${section.expectedWastage})`);
                    });

                    const cellsDisplay = allCells.length > 0 ? allCells.join(', ') : 'No data';

                    rows.push([
                        `${section.section} Wastage`,
                        cellsDisplay
                    ]);
                }
            }

            // Coats Thread checks - each check gets its own row
            if (coatsThreadCheck && coatsThreadCheck.found && coatsThreadCheck.checks) {
                coatsThreadCheck.checks.forEach((check, index) => {
                    const rowLabel = index === 0
                        ? `Coats Thread (Row ${coatsThreadCheck.rowNumber})`
                        : '';
                    const displayValue = !isNaN(check.numericValue) ? check.numericValue : check.actualValue;
                    const valueDisplay = check.isValid
                        ? `${check.label}: ${displayValue}`
                        : `${check.label}: ${displayValue} (Expected: ${check.expectedValue})`;
                    rows.push([
                        rowLabel,
                        valueDisplay
                    ]);
                });
            } else if (coatsThreadCheck && !coatsThreadCheck.found) {
                rows.push([
                    'Coats Thread',
                    coatsThreadCheck.message || 'Not found in Material section'
                ]);
            }

            // Process Costs - each item gets its own row
            if (processCostsCheck && processCostsCheck.items) {
                for (const item of processCostsCheck.items) {
                    if (!item.found) {
                        rows.push([
                            item.label,
                            'Not found'
                        ]);
                    } else {
                        const displayValue = !isNaN(item.numericValue) ? item.numericValue : item.actualValue;
                        const valueDisplay = item.isValid
                            ? `${displayValue}`
                            : `${displayValue} (Expected: ${item.expectedValue})`;
                        rows.push([
                            item.label,
                            valueDisplay
                        ]);
                    }
                }
            }

            return rows;
        }
    };
}
