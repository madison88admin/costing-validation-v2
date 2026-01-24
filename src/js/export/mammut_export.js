/**
 * Create Mammut (V7) export configuration
 * @param {Array} fileResults - Array of file results from Mammut processor
 * @returns {Object} - Configuration object for Mammut export
 */
function createMammutConfig(fileResults) {
    return {
        title: 'Mammut Validation Results - V7',
        fileResults: fileResults.map(fileResult => {
            const cellChecks = fileResult.results.cellChecks;
            const profitMargin = fileResult.results.profitMarginCheck;
            const wastageCost = fileResult.results.wastageCostCheck;
            const cmtCheck = fileResult.results.cmtCheck;

            // Count valid checks
            let validCount = 0;
            const wastageSectionCount = wastageCost.found && wastageCost.sections ? wastageCost.sections.length : 1;
            const cmtItemCount = cmtCheck && cmtCheck.found && cmtCheck.items ? cmtCheck.items.length : 0;
            let totalChecks = cellChecks.length + 1 + wastageSectionCount + cmtItemCount;

            cellChecks.forEach(check => {
                if (check.isValid) validCount++;
            });
            if (profitMargin.found && profitMargin.isValid) validCount++;
            if (wastageCost.found && wastageCost.sections) {
                wastageCost.sections.forEach(section => {
                    if (section.isValid) validCount++;
                });
            }
            if (cmtCheck && cmtCheck.found && cmtCheck.items) {
                cmtCheck.items.forEach(item => {
                    if (item.isValid) validCount++;
                });
            }

            // Build cell statuses for coloring
            const cellStatuses = [];

            // Cell checks rows
            cellChecks.forEach(check => {
                cellStatuses.push([
                    'normal',
                    check.isValid ? 'valid' : 'invalid'
                ]);
            });

            // Profit margin row
            cellStatuses.push([
                'normal',
                profitMargin.found && profitMargin.isValid ? 'valid' : 'invalid'
            ]);

            // Wastage rows
            if (wastageCost.found && wastageCost.sections) {
                wastageCost.sections.forEach(section => {
                    cellStatuses.push([
                        'normal',
                        section.isValid ? 'valid' : 'invalid'
                    ]);
                });
            }

            // CMT rows
            if (cmtCheck && cmtCheck.found && cmtCheck.items) {
                cmtCheck.items.forEach(item => {
                    cellStatuses.push([
                        'normal',
                        item.isValid ? 'valid' : 'invalid'
                    ]);
                });
            }

            return {
                fileName: fileResult.fileName,
                summary: `Summary: ${validCount} out of ${totalChecks} checks passed`,
                results: fileResult.results,
                cellStatuses: cellStatuses
            };
        }),
        filenamePrefix: 'MammutValidation_V7',
        columnWidths: [45, 105],
        headers: ['Validation Check', 'Value'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];
            const cellChecks = fileResult.results.cellChecks;
            const profitMargin = fileResult.results.profitMarginCheck;
            const wastageCost = fileResult.results.wastageCostCheck;
            const cmtCheck = fileResult.results.cmtCheck;

            // Helper to format number
            const formatNumber = (value) => {
                const num = parseFloat(String(value).replace(/[$,\s%]/g, ''));
                if (isNaN(num)) return value;
                return num.toFixed(2);
            };

            // Cell checks rows
            for (const check of cellChecks) {
                rows.push([
                    check.label,
                    check.found ? (check.actualValue || 'Empty') : 'Not found'
                ]);
            }

            // Profit Margin row
            if (profitMargin.found) {
                const pmActual = profitMargin.numericValue !== null ? formatNumber(profitMargin.actualValue) : profitMargin.actualValue;
                const pmRange = `${formatNumber(profitMargin.minValue)} - ${formatNumber(profitMargin.maxValue)}`;
                rows.push([
                    'PROFIT MARGIN',
                    `${pmActual} (Expected: ${pmRange})`
                ]);
            } else {
                rows.push([
                    'PROFIT MARGIN',
                    profitMargin.message || 'Not found'
                ]);
            }

            // Wastage rows for each section
            if (wastageCost.found && wastageCost.sections) {
                for (const section of wastageCost.sections) {
                    const sectionName = section.label.replace(' TOTAL', '');
                    const expectedPercent = (section.expectedValue * 100).toFixed(0);

                    // Combine valid and invalid cells
                    const allCells = [];

                    section.validCells.forEach(cell => {
                        allCells.push(cell.cellAddress);
                    });

                    section.invalidCells.forEach(cell => {
                        const roundedValue = cell.numericValue.toFixed(2);
                        allCells.push(`${cell.cellAddress} (${roundedValue})`);
                    });

                    rows.push([
                        `${sectionName} Wastage (${expectedPercent}%)`,
                        allCells.length > 0 ? allCells.join(', ') : 'No data'
                    ]);
                }
            } else {
                rows.push([
                    'Wastage Check',
                    wastageCost.message || 'No sections found'
                ]);
            }

            // CMT (Cut, Make, Trim) rows
            if (cmtCheck && cmtCheck.found && cmtCheck.items) {
                for (const item of cmtCheck.items) {
                    if (!item.found) {
                        rows.push([
                            item.label,
                            'Not found'
                        ]);
                    } else {
                        const details = [];

                        const priceDisplay = item.numericPrice !== null ? item.numericPrice.toFixed(2) : item.actualPrice;
                        details.push(`Price: ${priceDisplay}`);

                        const exRateDisplay = item.numericExRate !== null ? item.numericExRate.toFixed(2) : item.actualExRate;
                        details.push(`Ex.Rate: ${exRateDisplay}`);

                        details.push(`Currency: ${item.actualCurrency || 'N/A'}`);

                        rows.push([
                            item.label,
                            details.join(' | ')
                        ]);
                    }
                }
            } else if (cmtCheck && !cmtCheck.found) {
                rows.push([
                    'Process Check',
                    cmtCheck.message || 'Not found'
                ]);
            }

            return rows;
        }
    };
}
