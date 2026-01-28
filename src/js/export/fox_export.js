/**
 * Create FOX (V20) export configuration
 * @param {Array} fileResults - Array of file results from FOX processor
 * @returns {Object} - Configuration object for FOX export
 */
function createFOXConfig(fileResults) {
    return {
        title: 'FOX Validation Results - V20',
        fileResults: fileResults.map(fileResult => {
            const allResults = [
                fileResult.results.vendor,
                fileResult.results.factory,
                fileResult.results.coo,
                fileResult.results.overhead,
                fileResult.results.profitOthers,
                fileResult.results.wastagePercent,
                fileResult.results.sewingThread,
                fileResult.results.standardPackaging,
                fileResult.results.laborCost,
                fileResult.results.overheadCost,
                fileResult.results.profitCost
            ];

            const validCount = allResults.filter(r => r.isValid).length;
            const totalCount = allResults.length;

            // Build cell statuses for PDF coloring
            const cellStatuses = [];

            // Simple validations
            const simpleResults = [
                fileResult.results.vendor,
                fileResult.results.factory,
                fileResult.results.coo,
                fileResult.results.overhead,
                fileResult.results.profitOthers,
                fileResult.results.overheadCost,
                fileResult.results.profitCost
            ];

            for (const item of simpleResults) {
                cellStatuses.push([
                    'normal',
                    'normal',
                    item.isValid ? 'valid' : 'invalid'
                ]);
            }

            // Wastage Percent statuses
            const wastage = fileResult.results.wastagePercent;
            if (wastage) {
                const wastageStatuses = getWastageStatuses(wastage);
                cellStatuses.push(...wastageStatuses);
            }

            // Sewing Thread statuses
            const sewingThread = fileResult.results.sewingThread;
            if (sewingThread) {
                const sewingStatuses = getSewingThreadStatuses(sewingThread);
                cellStatuses.push(...sewingStatuses);
            }

            // Standard Packaging statuses
            const standardPackaging = fileResult.results.standardPackaging;
            if (standardPackaging) {
                const packagingStatuses = getStandardPackagingStatuses(standardPackaging);
                cellStatuses.push(...packagingStatuses);
            }

            // Labor Cost statuses
            const laborCost = fileResult.results.laborCost;
            if (laborCost) {
                const laborStatuses = getLaborCostStatuses(laborCost);
                cellStatuses.push(...laborStatuses);
            }

            return {
                fileName: fileResult.fileName,
                summary: `Validation: ${validCount}/${totalCount} passed`,
                results: fileResult.results,
                cellStatuses: cellStatuses
            };
        }),
        filenamePrefix: 'FOXValidation_V20',
        columnWidths: [50, 25, 195],
        headers: ['Field', 'Cell', 'Value'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];

            // Simple validations (single value checks)
            const simpleResults = [
                fileResult.results.vendor,
                fileResult.results.factory,
                fileResult.results.coo,
                fileResult.results.overhead,
                fileResult.results.profitOthers,
                fileResult.results.overheadCost,
                fileResult.results.profitCost
            ];

            for (const item of simpleResults) {
                let displayValue = '';
                if (item.actualValue === '' || item.actualValue === null || item.actualValue === 'Empty' || item.actualValue === 'Not found') {
                    displayValue = `${item.actualValue || 'Empty'} (Expected: ${item.expectedValue})`;
                } else if (item.isValid) {
                    displayValue = item.actualValue;
                } else {
                    displayValue = `${item.actualValue} (Expected: ${item.expectedValue})`;
                }

                rows.push([
                    item.label,
                    item.valueCell,
                    displayValue
                ]);
            }

            // Wastage Percent - each row separately
            const wastage = fileResult.results.wastagePercent;
            if (wastage) {
                const wastageRows = formatWastageForPDF(wastage);
                rows.push(...wastageRows);
            }

            // Sewing Thread - each row separately
            const sewingThread = fileResult.results.sewingThread;
            if (sewingThread) {
                const sewingRows = formatSewingThreadForPDF(sewingThread);
                rows.push(...sewingRows);
            }

            // Standard Packaging - each row separately
            const standardPackaging = fileResult.results.standardPackaging;
            if (standardPackaging) {
                const packagingRows = formatStandardPackagingForPDF(standardPackaging);
                rows.push(...packagingRows);
            }

            // Labor Cost - each item separately
            const laborCost = fileResult.results.laborCost;
            if (laborCost) {
                const laborRows = formatLaborCostForPDF(laborCost);
                rows.push(...laborRows);
            }

            return rows;
        }
    };
}

/**
 * Format Wastage results - each row separately
 */
function formatWastageForPDF(result) {
    const rows = [];

    if (result.labelCell === '-') {
        rows.push([
            'Wastage %',
            '-',
            'Section not found'
        ]);
        return rows;
    }

    // Valid rows
    if (result.validRows && result.validRows.length > 0) {
        for (const row of result.validRows) {
            rows.push([
                'Wastage %',
                row.cell,
                '5%'
            ]);
        }
    }

    // Invalid rows
    if (result.invalidRows && result.invalidRows.length > 0) {
        for (const row of result.invalidRows) {
            rows.push([
                'Wastage %',
                row.cell,
                `${row.value} (Exp: 5%)`
            ]);
        }
    }

    if (rows.length === 0) {
        rows.push([
            'Wastage %',
            '-',
            'No data rows'
        ]);
    }

    return rows;
}

/**
 * Get Wastage cell statuses
 */
function getWastageStatuses(result) {
    const statuses = [];

    if (result.labelCell === '-') {
        statuses.push(['normal', 'normal', 'invalid']);
        return statuses;
    }

    if (result.validRows && result.validRows.length > 0) {
        for (const row of result.validRows) {
            statuses.push(['normal', 'normal', 'valid']);
        }
    }

    if (result.invalidRows && result.invalidRows.length > 0) {
        for (const row of result.invalidRows) {
            statuses.push(['normal', 'normal', 'invalid']);
        }
    }

    if (statuses.length === 0) {
        statuses.push(['normal', 'normal', 'invalid']);
    }

    return statuses;
}

/**
 * Format Sewing Thread results - each row separately with compact format
 */
function formatSewingThreadForPDF(result) {
    const rows = [];

    if (result.notFound) {
        rows.push([
            'Sewing Thread',
            '-',
            'Not found'
        ]);
        return rows;
    }

    // Group fields by row number
    const allFields = [...(result.validFields || []), ...(result.invalidFields || [])];
    const rowGroups = {};

    for (const field of allFields) {
        const rowNum = field.cell.replace(/[A-Z]/g, '');
        if (!rowGroups[rowNum]) {
            rowGroups[rowNum] = { valid: [], invalid: [] };
        }
        const isValid = result.validFields?.some(f => f.cell === field.cell);
        if (isValid) {
            rowGroups[rowNum].valid.push(field);
        } else {
            rowGroups[rowNum].invalid.push(field);
        }
    }

    const sortedRows = Object.keys(rowGroups).sort((a, b) => parseInt(a) - parseInt(b));

    for (const rowNum of sortedRows) {
        const group = rowGroups[rowNum];
        const allRowValid = group.invalid.length === 0;

        // Build compact value string
        const allFieldsInRow = [...group.valid, ...group.invalid];

        // Sort by column letter
        allFieldsInRow.sort((a, b) => {
            const colA = a.cell.replace(/[0-9]/g, '');
            const colB = b.cell.replace(/[0-9]/g, '');
            return colA.localeCompare(colB);
        });

        const fieldParts = allFieldsInRow.map(field => {
            const colLetter = field.cell.replace(/[0-9]/g, '');
            const isValid = group.valid.some(f => f.cell === field.cell);
            if (isValid) {
                return `${colLetter}=${field.value}`;
            } else {
                return `${colLetter}=${field.value}(Exp:${field.expected})`;
            }
        });

        rows.push([
            `Sewing Thread R${rowNum}`,
            `Row ${rowNum}`,
            fieldParts.join(', ')
        ]);
    }

    return rows;
}

/**
 * Get Sewing Thread cell statuses
 */
function getSewingThreadStatuses(result) {
    const statuses = [];

    if (result.notFound) {
        statuses.push(['normal', 'normal', 'invalid']);
        return statuses;
    }

    const allFields = [...(result.validFields || []), ...(result.invalidFields || [])];
    const rowGroups = {};

    for (const field of allFields) {
        const rowNum = field.cell.replace(/[A-Z]/g, '');
        if (!rowGroups[rowNum]) {
            rowGroups[rowNum] = { valid: [], invalid: [] };
        }
        const isValid = result.validFields?.some(f => f.cell === field.cell);
        if (isValid) {
            rowGroups[rowNum].valid.push(field);
        } else {
            rowGroups[rowNum].invalid.push(field);
        }
    }

    const sortedRows = Object.keys(rowGroups).sort((a, b) => parseInt(a) - parseInt(b));

    for (const rowNum of sortedRows) {
        const group = rowGroups[rowNum];
        const allRowValid = group.invalid.length === 0;
        statuses.push(['normal', 'normal', allRowValid ? 'valid' : 'invalid']);
    }

    return statuses;
}

/**
 * Format Standard Packaging results - each row separately
 */
function formatStandardPackagingForPDF(result) {
    const rows = [];

    if (result.notFound) {
        rows.push([
            'Standard Packaging',
            '-',
            'Not found'
        ]);
        return rows;
    }

    const allFields = [...(result.validFields || []), ...(result.invalidFields || [])];
    const rowGroups = {};

    for (const field of allFields) {
        const rowNum = field.cell.replace(/[A-Z]/g, '');
        if (!rowGroups[rowNum]) {
            rowGroups[rowNum] = { valid: [], invalid: [] };
        }
        const isValid = result.validFields?.some(f => f.cell === field.cell);
        if (isValid) {
            rowGroups[rowNum].valid.push(field);
        } else {
            rowGroups[rowNum].invalid.push(field);
        }
    }

    const sortedRows = Object.keys(rowGroups).sort((a, b) => parseInt(a) - parseInt(b));

    for (const rowNum of sortedRows) {
        const group = rowGroups[rowNum];

        const allFieldsInRow = [...group.valid, ...group.invalid];

        // Sort by column letter
        allFieldsInRow.sort((a, b) => {
            const colA = a.cell.replace(/[0-9]/g, '');
            const colB = b.cell.replace(/[0-9]/g, '');
            return colA.localeCompare(colB);
        });

        const fieldParts = allFieldsInRow.map(field => {
            const colLetter = field.cell.replace(/[0-9]/g, '');
            const isValid = group.valid.some(f => f.cell === field.cell);
            if (isValid) {
                return `${colLetter}=${field.value}`;
            } else {
                return `${colLetter}=${field.value}(Exp:${field.expected})`;
            }
        });

        rows.push([
            `Std Packaging R${rowNum}`,
            `Row ${rowNum}`,
            fieldParts.join(', ')
        ]);
    }

    return rows;
}

/**
 * Get Standard Packaging cell statuses
 */
function getStandardPackagingStatuses(result) {
    const statuses = [];

    if (result.notFound) {
        statuses.push(['normal', 'normal', 'invalid']);
        return statuses;
    }

    const allFields = [...(result.validFields || []), ...(result.invalidFields || [])];
    const rowGroups = {};

    for (const field of allFields) {
        const rowNum = field.cell.replace(/[A-Z]/g, '');
        if (!rowGroups[rowNum]) {
            rowGroups[rowNum] = { valid: [], invalid: [] };
        }
        const isValid = result.validFields?.some(f => f.cell === field.cell);
        if (isValid) {
            rowGroups[rowNum].valid.push(field);
        } else {
            rowGroups[rowNum].invalid.push(field);
        }
    }

    const sortedRows = Object.keys(rowGroups).sort((a, b) => parseInt(a) - parseInt(b));

    for (const rowNum of sortedRows) {
        const group = rowGroups[rowNum];
        const allRowValid = group.invalid.length === 0;
        statuses.push(['normal', 'normal', allRowValid ? 'valid' : 'invalid']);
    }

    return statuses;
}

/**
 * Format Labor Cost results - each item separately
 */
function formatLaborCostForPDF(result) {
    const rows = [];

    if (result.notFound) {
        rows.push([
            'Labor Cost',
            '-',
            'LABOR COST not found'
        ]);
        return rows;
    }

    // Found items
    if (result.foundItems && result.foundItems.length > 0) {
        for (const item of result.foundItems) {
            rows.push([
                `Labor - ${item.item}`,
                item.cell,
                item.value
            ]);
        }
    }

    // Missing items
    if (result.missingItems && result.missingItems.length > 0) {
        for (const item of result.missingItems) {
            rows.push([
                `Labor - ${item}`,
                '-',
                'Missing'
            ]);
        }
    }

    return rows;
}

/**
 * Get Labor Cost cell statuses
 */
function getLaborCostStatuses(result) {
    const statuses = [];

    if (result.notFound) {
        statuses.push(['normal', 'normal', 'invalid']);
        return statuses;
    }

    if (result.foundItems && result.foundItems.length > 0) {
        for (const item of result.foundItems) {
            statuses.push(['normal', 'normal', 'valid']);
        }
    }

    if (result.missingItems && result.missingItems.length > 0) {
        for (const item of result.missingItems) {
            statuses.push(['normal', 'normal', 'invalid']);
        }
    }

    return statuses;
}
