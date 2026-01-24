/**
 * Create LLBEAN (V6) export configuration
 * @param {Array} fileResults - Array of file results from LLBEAN processor
 * @returns {Object} - Configuration object for LLBEAN export
 */
function createLLBEANConfig(fileResults) {
    return {
        title: 'LLBEAN Validation Results - V6',
        fileResults: fileResults.map(fileResult => {
            const b5 = fileResult.results.b5Check;
            const trimsBox = fileResult.results.trimsBoxCheck;
            const totalFinancial = fileResult.results.totalFinancialCostCheck;

            // Count valid checks
            let validCount = 0;
            let totalChecks = 3;

            if (b5.isValid) validCount++;
            if (trimsBox.found && trimsBox.isValid) validCount++;
            if (totalFinancial.found && totalFinancial.isValid) validCount++;

            // Build cell statuses for coloring
            const cellStatuses = [];

            // B5 Keywords row
            cellStatuses.push([
                'normal', // Check name
                'normal', // Supplier
                b5.isValid ? 'valid' : 'invalid', // Consumption (keywords)
                'normal', // Unit Price
                'normal', // Total Cost
                'normal'  // Status
            ]);

            // Trims Box row
            if (trimsBox.found && trimsBox.boxData) {
                cellStatuses.push([
                    'normal', // Check name
                    trimsBox.validation.supplier === 'VALID' ? 'valid' : 'invalid',
                    trimsBox.validation.consumption === 'VALID' ? 'valid' : 'invalid',
                    trimsBox.validation.unitPrice === 'VALID' ? 'valid' : 'invalid',
                    trimsBox.validation.totalCost === 'VALID' ? 'valid' : 'invalid',
                    'normal'  // Status
                ]);
            } else {
                cellStatuses.push([
                    'normal', 'invalid', 'invalid', 'invalid', 'invalid', 'normal'
                ]);
            }

            // Total Financial Cost row
            if (totalFinancial.found && totalFinancial.expectedValue !== null) {
                cellStatuses.push([
                    'normal', // Check name
                    'normal', // Supplier
                    'normal', // Consumption
                    'normal', // Unit Price
                    totalFinancial.validation === 'VALID' ? 'valid' : 'invalid',
                    'normal'  // Status
                ]);
            } else {
                cellStatuses.push([
                    'normal', 'normal', 'normal', 'normal', 'invalid', 'normal'
                ]);
            }

            return {
                fileName: fileResult.fileName,
                summary: `Summary: ${validCount} out of ${totalChecks} checks passed`,
                results: fileResult.results,
                cellStatuses: cellStatuses
            };
        }),
        filenamePrefix: 'LLBEANValidation_V6',
        columnWidths: [35, 25, 30, 25, 25, 20],
        headers: ['Validation Check', 'Supplier', 'Consumption', 'Unit Price', 'Total Cost', 'Status'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];
            const b5 = fileResult.results.b5Check;
            const trimsBox = fileResult.results.trimsBoxCheck;
            const totalFinancial = fileResult.results.totalFinancialCostCheck;

            // Helper to format cell for PDF
            const formatForPDF = (expected, actual, status, isNumeric = false) => {
                if (!actual || actual === '') {
                    const displayExpected = isNumeric ? parseFloat(expected).toFixed(2) : expected;
                    return `Empty (Expected: ${displayExpected})`;
                }

                const displayActual = isNumeric ? parseFloat(actual).toFixed(2) : actual;
                const displayExpected = isNumeric ? parseFloat(expected).toFixed(2) : expected;

                if (status === 'VALID') {
                    return displayActual;
                } else {
                    return `${displayActual} (Expected: ${displayExpected})`;
                }
            };

            // B5 Keywords row
            rows.push([
                `Cell B5 Keywords\n(Value: ${b5.cellValue || 'Empty'})`,
                '-',
                b5.foundKeywords.length > 0 ? b5.foundKeywords.join(', ') : b5.requiredKeywords.join(', '),
                '-',
                '-',
                b5.isValid ? 'VALID' : 'INVALID'
            ]);

            // Trims Box row
            if (trimsBox.found && trimsBox.boxData) {
                const v = trimsBox.validation;
                const expected = trimsBox.expected;
                const actual = trimsBox.boxData;

                rows.push([
                    `Trims - Box\n(Row ${actual.rowNumber})`,
                    formatForPDF(expected.supplier, actual.supplier, v.supplier, false),
                    formatForPDF(expected.consumption, actual.consumption, v.consumption, true),
                    formatForPDF(expected.unitPrice, actual.unitPrice, v.unitPrice, true),
                    formatForPDF(expected.totalCost, actual.totalCost, v.totalCost, true),
                    trimsBox.isValid ? 'VALID' : 'INVALID'
                ]);
            } else {
                rows.push([
                    'Trims - Box',
                    '-',
                    trimsBox.message || 'Not found',
                    '-',
                    '-',
                    'NOT FOUND'
                ]);
            }

            // Total Financial Cost row
            if (totalFinancial.found && totalFinancial.expectedValue !== null) {
                rows.push([
                    `Total Financial Cost\n(Row ${totalFinancial.rowNumber}, ${totalFinancial.matchedKeyword})`,
                    '-',
                    '-',
                    '-',
                    formatForPDF(totalFinancial.expectedValue, totalFinancial.actualValue, totalFinancial.validation, true),
                    totalFinancial.isValid ? 'VALID' : 'INVALID'
                ]);
            } else {
                rows.push([
                    'Total Financial Cost',
                    '-',
                    totalFinancial.message || 'Not found',
                    '-',
                    '-',
                    'NOT FOUND'
                ]);
            }

            return rows;
        }
    };
}
