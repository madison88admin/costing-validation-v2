/**
 * Create KUHL (V19) export configuration
 * @param {Array} fileResults - Array of file results from KUHL processor
 * @returns {Object} - Configuration object for KUHL export
 */
function createKuhlConfig(fileResults) {
    // Helper to format all cells (valid in green text, invalid in red text for PDF)
    const formatAllCells = (validCells, invalidCells) => {
        const allCells = [];
        if (validCells && validCells.length > 0) {
            allCells.push(...validCells);
        }
        if (invalidCells && invalidCells.length > 0) {
            allCells.push(...invalidCells.map(c => c.cell));
        }
        return allCells.length > 0 ? allCells.join(', ') : '-';
    };

    return {
        title: 'KUHL Validation Results - V19',
        fileResults: fileResults.map(fileResult => {
            const fabricYarn = fileResult.results.fabricYarn;
            const trim = fileResult.results.trim;
            const labelling = fileResult.results.labelling;
            const profitMargin = fileResult.results.profitMargin;

            // Calculate totals
            const totalValid = fabricYarn.validCells.length +
                trim.consumption.validCells.length + trim.supplier.validCells.length + trim.cifVsFob.validCells.length +
                labelling.consumption.validCells.length + labelling.supplier.validCells.length + labelling.cifVsFob.validCells.length +
                profitMargin.validCells.length;

            const totalInvalid = fabricYarn.invalidCells.length +
                trim.consumption.invalidCells.length + trim.supplier.invalidCells.length + trim.cifVsFob.invalidCells.length +
                labelling.consumption.invalidCells.length + labelling.supplier.invalidCells.length + labelling.cifVsFob.invalidCells.length +
                profitMargin.invalidCells.length;

            // Build cell statuses for coloring (8 rows)
            const cellStatuses = [
                ['normal', 'normal', fabricYarn.invalidCells.length > 0 ? 'invalid' : (fabricYarn.validCells.length > 0 ? 'valid' : 'normal')],
                ['normal', 'normal', trim.consumption.invalidCells.length > 0 ? 'invalid' : (trim.consumption.validCells.length > 0 ? 'valid' : 'normal')],
                ['normal', 'normal', trim.supplier.invalidCells.length > 0 ? 'invalid' : (trim.supplier.validCells.length > 0 ? 'valid' : 'normal')],
                ['normal', 'normal', trim.cifVsFob.invalidCells.length > 0 ? 'invalid' : (trim.cifVsFob.validCells.length > 0 ? 'valid' : 'normal')],
                ['normal', 'normal', labelling.consumption.invalidCells.length > 0 ? 'invalid' : (labelling.consumption.validCells.length > 0 ? 'valid' : 'normal')],
                ['normal', 'normal', labelling.supplier.invalidCells.length > 0 ? 'invalid' : (labelling.supplier.validCells.length > 0 ? 'valid' : 'normal')],
                ['normal', 'normal', labelling.cifVsFob.invalidCells.length > 0 ? 'invalid' : (labelling.cifVsFob.validCells.length > 0 ? 'valid' : 'normal')],
                ['normal', 'normal', profitMargin.invalidCells.length > 0 ? 'invalid' : (profitMargin.validCells.length > 0 ? 'valid' : 'normal')]
            ];

            return {
                fileName: fileResult.fileName,
                summary: `Total Validation: ${totalValid} passed, ${totalInvalid} failed`,
                fabricYarn: fabricYarn,
                trim: trim,
                labelling: labelling,
                profitMargin: profitMargin,
                cellStatuses: cellStatuses
            };
        }),
        filenamePrefix: 'KUHL_Validation',
        columnWidths: [35, 55, 130],
        headers: ['Type', 'Column', 'Cells'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];
            const fabricYarn = fileResult.fabricYarn;
            const trim = fileResult.trim;
            const labelling = fileResult.labelling;
            const profitMargin = fileResult.profitMargin;

            // Fabric/Yarn row
            rows.push([
                'Fabric/Yarn',
                'K (Consumption = 5%)',
                formatAllCells(fabricYarn.validCells, fabricYarn.invalidCells)
            ]);

            // Trim rows
            rows.push([
                'Trim',
                'K (Consumption = 3%)',
                formatAllCells(trim.consumption.validCells, trim.consumption.invalidCells)
            ]);

            rows.push([
                'Trim',
                'E (Supplier = "Local")',
                formatAllCells(trim.supplier.validCells, trim.supplier.invalidCells)
            ]);

            rows.push([
                'Trim',
                'H (C.I.F.VS FOB = 0.012%)',
                formatAllCells(trim.cifVsFob.validCells, trim.cifVsFob.invalidCells)
            ]);

            // Labelling rows
            rows.push([
                'Labelling',
                'K (Consumption = 3%)',
                formatAllCells(labelling.consumption.validCells, labelling.consumption.invalidCells)
            ]);

            rows.push([
                'Labelling',
                'E (Supplier = "Local")',
                formatAllCells(labelling.supplier.validCells, labelling.supplier.invalidCells)
            ]);

            rows.push([
                'Labelling',
                'H (C.I.F.VS FOB = 0.012%)',
                formatAllCells(labelling.cifVsFob.validCells, labelling.cifVsFob.invalidCells)
            ]);

            // Profit Margin row
            rows.push([
                'Profit Margin',
                'M (Value = 0.60-0.95)',
                formatAllCells(profitMargin.validCells, profitMargin.invalidCells)
            ]);

            return rows;
        }
    };
}
