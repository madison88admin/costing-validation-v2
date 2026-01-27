/**
 * Create Foot Asylum (V18) export configuration
 * @param {Array} fileResults - Array of file results from Foot Asylum processor
 * @returns {Object} - Configuration object for Foot Asylum export
 */
function createFootAsylumConfig(fileResults) {
    const validationRules = {
        mainMaterial: { column: 'D', label: 'Main Material', expectedDisplay: 'TRUE' },
        supplierCurrency: { column: 'P', label: 'Supplier Currency', expectedDisplay: 'USD' },
        wastage: { column: 'W', label: 'Wastage %', expectedDisplay: '5%' },
        overheadCost: { column: 'AG', label: 'Overhead Cost', expectedDisplay: '0.5' },
        testingCost: { column: 'AK', label: 'Testing Cost', expectedDisplay: '0.1' },
        profitFOB: { column: 'AL', label: 'Profit %', expectedDisplay: '10%' }
    };

    const ruleKeys = Object.keys(validationRules);

    return {
        title: 'Foot Asylum Validation Results - V18',
        fileResults: fileResults.map(fileResult => {
            if (!fileResult.results.sectionFound) {
                return {
                    fileName: fileResult.fileName,
                    summary: 'Error: Fabrics section not found',
                    rows: [],
                    cellStatuses: []
                };
            }

            // Calculate totals
            let totalValid = 0;
            let totalInvalid = 0;

            for (const rowResult of fileResult.results.rows) {
                for (const cellData of Object.values(rowResult.columns)) {
                    if (!cellData.isEmpty) {
                        if (cellData.isValid) {
                            totalValid++;
                        } else {
                            totalInvalid++;
                        }
                    }
                }
            }

            // Build cell statuses for coloring
            const cellStatuses = fileResult.results.rows.map(rowResult => {
                const rowStatuses = ['normal']; // Row number column
                for (const key of ruleKeys) {
                    const cellData = rowResult.columns[key];
                    if (cellData.isEmpty) {
                        rowStatuses.push('normal');
                    } else if (cellData.isValid) {
                        rowStatuses.push('valid');
                    } else {
                        rowStatuses.push('invalid');
                    }
                }
                return rowStatuses;
            });

            return {
                fileName: fileResult.fileName,
                summary: `Fabrics Section: Rows ${fileResult.results.startRow + 1}-${fileResult.results.endRow} | Validations: ${totalValid} passed, ${totalInvalid} failed`,
                rows: fileResult.results.rows,
                cellStatuses: cellStatuses
            };
        }),
        filenamePrefix: 'FootAsylum_Validation',
        columnWidths: [20, 30, 30, 25, 30, 30, 25],
        headers: ['Row', 'Main Material (D)', 'Supplier Currency (P)', 'Wastage % (W)', 'Overhead Cost (AG)', 'Testing Cost (AK)', 'Profit % (AL)'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];

            if (!fileResult.rows || fileResult.rows.length === 0) {
                rows.push(['Error', 'Fabrics section not found or no data rows', '', '', '', '', '']);
                return rows;
            }

            // Add data rows
            for (const rowResult of fileResult.rows) {
                const rowData = [rowResult.rowNumber.toString()];

                for (const key of ruleKeys) {
                    const cellData = rowResult.columns[key];
                    if (cellData.isEmpty) {
                        rowData.push('-');
                    } else if (cellData.isValid) {
                        rowData.push(cellData.value);
                    } else {
                        rowData.push(`${cellData.value} (Expected: ${cellData.expected})`);
                    }
                }

                rows.push(rowData);
            }

            return rows;
        }
    };
}
