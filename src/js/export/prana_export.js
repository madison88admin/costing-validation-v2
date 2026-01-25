/**
 * Create Prana (V13) export configuration
 * @param {Array} fileResults - Array of file results from Prana processor
 * @returns {Object} - Configuration object for Prana export
 */
function createPranaConfig(fileResults) {
    return {
        title: 'Prana Validation Results - V13',
        fileResults: fileResults.map(fileResult => {
            const allRows = [];
            const allCellStatuses = [];

            // Process each sheet
            for (const sheet of fileResult.sheets || []) {
                // Process sections
                for (const section of sheet.sections || []) {
                    const validItems = section.items.filter(item => item.isValid).length;
                    const invalidItems = section.items.filter(item => !item.isValid).length;
                    const totalItems = section.items.length;

                    // Determine cell status for this row
                    let cellStatus = 'normal';
                    if (totalItems > 0) {
                        cellStatus = invalidItems === 0 ? 'valid' : 'invalid';
                    }

                    // Check special items and global checks for validity
                    const hasInvalidSpecialItems = section.specialItemResults?.some(item =>
                        item.checks.some(check => !check.isValid)
                    );

                    if (hasInvalidSpecialItems) {
                        cellStatus = 'invalid';
                    }

                    allCellStatuses.push(['normal', 'normal', cellStatus]);
                    allRows.push({
                        sheetName: sheet.sheetName,
                        section: section,
                        type: 'section'
                    });
                }

                // Process global checks
                if (sheet.globalChecks && sheet.globalChecks.length > 0) {
                    const hasInvalidGlobalChecks = sheet.globalChecks.some(check =>
                        check.checks.some(c => !c.isValid)
                    );

                    allCellStatuses.push(['normal', 'normal', hasInvalidGlobalChecks ? 'invalid' : 'valid']);
                    allRows.push({
                        sheetName: sheet.sheetName,
                        globalChecks: sheet.globalChecks,
                        type: 'global'
                    });
                }
            }

            // Calculate summary
            let totalWastageItems = 0;
            let validWastageItems = 0;
            for (const sheet of fileResult.sheets || []) {
                for (const section of sheet.sections || []) {
                    totalWastageItems += section.items.length;
                    validWastageItems += section.items.filter(i => i.isValid).length;
                }
            }

            return {
                fileName: fileResult.fileName,
                summary: `Summary: ${validWastageItems} out of ${totalWastageItems} wastage values are correct`,
                results: allRows,
                cellStatuses: allCellStatuses
            };
        }),
        filenamePrefix: 'PranaValidation_V13',
        columnWidths: [28, 35, 87],
        headers: ['Sheet Name', 'Section', 'Details'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];
            const results = fileResult.results || [];

            for (const result of results) {
                if (result.type === 'section') {
                    const section = result.section;
                    const validItems = section.items.filter(item => item.isValid);
                    const invalidItems = section.items.filter(item => !item.isValid);

                    // Build details cell content
                    let details = '';

                    // Add wastage cells
                    if (validItems.length > 0) {
                        const validCells = validItems.map(item => item.cellAddress).join(', ');
                        details += validCells;
                    }

                    if (invalidItems.length > 0) {
                        if (validItems.length > 0) {
                            details += '\n';
                        }
                        const invalidCells = invalidItems.map(item =>
                            `${item.cellAddress}: ${item.actual} (Expected: ${section.expectedWastage})`
                        ).join(', ');
                        details += invalidCells;
                    }

                    if (validItems.length === 0 && invalidItems.length === 0) {
                        details = 'No items found in section';
                    }

                    // Add special items
                    if (section.specialItemResults && section.specialItemResults.length > 0) {
                        for (const specialItem of section.specialItemResults) {
                            details += '\n\n';
                            const columnADisplay = specialItem.columnA ? `${specialItem.columnA} - ` : '';
                            details += `${columnADisplay}${specialItem.name} (Row ${specialItem.rowNumber}):\n`;

                            const checkResults = specialItem.checks.map(c => {
                                if (c.isValid) {
                                    return `${c.label}: ${c.actual}`;
                                } else {
                                    return `${c.label}: ${c.actual} (Expected: ${c.expected})`;
                                }
                            }).join(', ');
                            details += checkResults;
                        }
                    }

                    rows.push([
                        result.sheetName,
                        `${section.name}\nExpected: ${section.expectedWastage}`,
                        details
                    ]);
                } else if (result.type === 'global') {
                    // Build global checks details
                    let details = '';

                    for (let i = 0; i < result.globalChecks.length; i++) {
                        const check = result.globalChecks[i];
                        if (i > 0) details += '\n\n';

                        details += `${check.name} (Row ${check.rowNumber}):\n`;

                        const checkResults = check.checks.map(c => {
                            if (c.isValid) {
                                return `${c.label}: ${c.actual}`;
                            } else {
                                return `${c.label}: ${c.actual} (Expected: ${c.expected})`;
                            }
                        }).join(', ');
                        details += checkResults;
                    }

                    rows.push([
                        result.sheetName,
                        'Global Checks',
                        details
                    ]);
                }
            }

            return rows;
        }
    };
}
