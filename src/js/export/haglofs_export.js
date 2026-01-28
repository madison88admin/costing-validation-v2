/**
 * Create Haglofs (V21) export configuration
 * @param {Array} fileResults - Array of file results from Haglofs processor
 * @returns {Object} - Configuration object for Haglofs export
 */
function createHaglofsConfig(fileResults) {
    return {
        title: 'Haglofs Validation Results - V21',
        fileResults: fileResults.map(fileResult => {
            const supplier = fileResult.results.supplier;
            const fabricRows = fileResult.results.fabricAllowanceRows;
            const trimsRows = fileResult.results.trimsAllowanceRows;
            const packagingRows = fileResult.results.packagingAllowanceRows;
            const genericPackaging = fileResult.results.genericPackaging;
            const overhead = fileResult.results.overhead;
            const margin = fileResult.results.margin;
            const validFabricCount = fabricRows.filter(r => r.isValid).length;
            const validTrimsCount = trimsRows.filter(r => r.isValid).length;
            const validPackagingCount = packagingRows.filter(r => r.isValid).length;

            // Build cell statuses for PDF coloring (2 columns now)
            const cellStatuses = [];

            // Supplier validation status
            cellStatuses.push([
                'normal',
                supplier.isValid ? 'valid' : 'invalid'
            ]);

            // Fabric allowance status
            const fabricAllValid = fabricRows.length > 0 && fabricRows.every(r => r.isValid);
            cellStatuses.push([
                'normal',
                fabricRows.length === 0 ? 'normal' : (fabricAllValid ? 'valid' : 'invalid')
            ]);

            // Trims allowance status
            const trimsAllValid = trimsRows.length > 0 && trimsRows.every(r => r.isValid);
            cellStatuses.push([
                'normal',
                trimsRows.length === 0 ? 'normal' : (trimsAllValid ? 'valid' : 'invalid')
            ]);

            // Packaging allowance status
            const packagingAllValid = packagingRows.length > 0 && packagingRows.every(r => r.isValid);
            cellStatuses.push([
                'normal',
                packagingRows.length === 0 ? 'normal' : (packagingAllValid ? 'valid' : 'invalid')
            ]);

            // Generic Packaging validation status
            cellStatuses.push([
                'normal',
                genericPackaging ? (genericPackaging.isValid ? 'valid' : 'invalid') : 'normal'
            ]);

            // Overhead validation status
            cellStatuses.push([
                'normal',
                overhead.isValid ? 'valid' : 'invalid'
            ]);

            // Margin validation status
            cellStatuses.push([
                'normal',
                margin.isValid ? 'valid' : 'invalid'
            ]);

            return {
                fileName: fileResult.fileName,
                summary: `Supplier: ${supplier.isValid ? 'Valid' : 'Invalid'}, Fabric (5%): ${validFabricCount}/${fabricRows.length}, Trims (3%): ${validTrimsCount}/${trimsRows.length}, Packaging (3%): ${validPackagingCount}/${packagingRows.length}${genericPackaging ? `, Generic Pkg: ${genericPackaging.isValid ? 'Valid' : 'Invalid'}` : ''}`,
                results: fileResult.results,
                cellStatuses: cellStatuses
            };
        }),
        filenamePrefix: 'HaglofsValidation_V21',
        columnWidths: [110, 280],
        headers: ['Validation Field', 'Results'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];
            const supplier = fileResult.results.supplier;
            const fabricRows = fileResult.results.fabricAllowanceRows;
            const trimsRows = fileResult.results.trimsAllowanceRows;
            const packagingRows = fileResult.results.packagingAllowanceRows;
            const genericPackaging = fileResult.results.genericPackaging;
            const overhead = fileResult.results.overhead;
            const margin = fileResult.results.margin;

            // Supplier validation row
            const supplierResult = supplier.isValid
                ? supplier.actual
                : `${supplier.actual} (Expected: ${supplier.expected})`;
            rows.push([
                `Supplier (Row ${supplier.rowIndex || 'N/A'})`,
                supplierResult
            ]);

            // Fabric allowance - consolidated
            const fabricValid = [];
            const fabricInvalid = [];
            for (const row of fabricRows) {
                if (row.isValid) {
                    fabricValid.push(`Row ${row.rowIndex}`);
                } else {
                    fabricInvalid.push(`Row ${row.rowIndex}: ${row.actual} (Expected: ${row.expected})`);
                }
            }
            let fabricResult = '';
            if (fabricRows.length === 0) {
                fabricResult = 'No section found';
            } else {
                if (fabricValid.length > 0) fabricResult += fabricValid.join(', ');
                if (fabricInvalid.length > 0) {
                    if (fabricValid.length > 0) fabricResult += ' | ';
                    fabricResult += fabricInvalid.join(' | ');
                }
            }
            rows.push(['Fabric (5%)', fabricResult]);

            // Trims allowance - consolidated
            const trimsValid = [];
            const trimsInvalid = [];
            for (const row of trimsRows) {
                if (row.isValid) {
                    trimsValid.push(`Row ${row.rowIndex}`);
                } else {
                    trimsInvalid.push(`Row ${row.rowIndex}: ${row.actual} (Expected: ${row.expected})`);
                }
            }
            let trimsResult = '';
            if (trimsRows.length === 0) {
                trimsResult = 'No section found';
            } else {
                if (trimsValid.length > 0) trimsResult += trimsValid.join(', ');
                if (trimsInvalid.length > 0) {
                    if (trimsValid.length > 0) trimsResult += ' | ';
                    trimsResult += trimsInvalid.join(' | ');
                }
            }
            rows.push(['Trims (3%)', trimsResult]);

            // Packaging allowance - consolidated
            const packagingValid = [];
            const packagingInvalid = [];
            for (const row of packagingRows) {
                if (row.isValid) {
                    packagingValid.push(`Row ${row.rowIndex}`);
                } else {
                    packagingInvalid.push(`Row ${row.rowIndex}: ${row.actual} (Expected: ${row.expected})`);
                }
            }
            let packagingResult = '';
            if (packagingRows.length === 0) {
                packagingResult = 'No section found';
            } else {
                if (packagingValid.length > 0) packagingResult += packagingValid.join(', ');
                if (packagingInvalid.length > 0) {
                    if (packagingValid.length > 0) packagingResult += ' | ';
                    packagingResult += packagingInvalid.join(' | ');
                }
            }
            rows.push(['Packaging (3%)', packagingResult]);

            // Generic Packaging - show all fields
            if (genericPackaging) {
                const parts = [];
                parts.push(`B: ${genericPackaging.colB.actual}${!genericPackaging.colB.isValid ? ` (Exp: ${genericPackaging.colB.expected})` : ''}`);
                parts.push(`F: ${genericPackaging.colF.actual}${!genericPackaging.colF.isValid ? ` (Exp: ${genericPackaging.colF.expected})` : ''}`);
                parts.push(`G: ${genericPackaging.colG.actual}${!genericPackaging.colG.isValid ? ` (Exp: ${genericPackaging.colG.expected})` : ''}`);
                parts.push(`H: ${genericPackaging.colH.actual}${!genericPackaging.colH.isValid ? ` (Exp: ${genericPackaging.colH.expected})` : ''}`);

                rows.push([
                    'Generic Packaging',
                    `Row ${genericPackaging.rowIndex}: ${parts.join(', ')}`
                ]);
            } else {
                rows.push(['Generic Packaging', 'Not found']);
            }

            // Overhead validation row
            const overheadResult = overhead.isValid
                ? overhead.actual
                : `${overhead.actual} (Expected: ${overhead.expected})`;
            rows.push([
                `Overhead (Row ${overhead.rowIndex || 'N/A'})`,
                overheadResult
            ]);

            // Margin validation row
            const marginResult = margin.isValid
                ? margin.actual
                : `${margin.actual} (Expected: ${margin.expected})`;
            rows.push([
                `Margin (Row ${margin.rowIndex || 'N/A'})`,
                marginResult
            ]);

            return rows;
        }
    };
}

// Export to global scope
if (typeof window !== 'undefined') {
    window.haglofsExportConfig = {
        createHaglofsConfig: createHaglofsConfig
    };
}
