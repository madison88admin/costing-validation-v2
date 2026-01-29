/**
 * Create Foot Asylum (V18) export configuration
 * @param {Array} fileResults - Array of file results from Foot Asylum processor
 * @returns {Object} - Configuration object for Foot Asylum export
 */
function createFootAsylumConfig(fileResults) {
    // Fabrics section has all 6 columns
    const fabricsRuleKeys = ['mainMaterial', 'supplierCurrency', 'wastage', 'overheadCost', 'testingCost', 'profitFOB'];
    // Trims and Packaging only have 3 columns
    const trimsRuleKeys = ['mainMaterial', 'supplierCurrency', 'wastage'];

    return {
        title: 'Foot Asylum Validation Results - V18',
        fileResults: fileResults.map(fileResult => {
            const hasFabrics = fileResult.results.fabrics && fileResult.results.fabrics.sectionFound;
            const hasTrims = fileResult.results.trims && fileResult.results.trims.sectionFound;
            const hasPackaging = fileResult.results.packaging && fileResult.results.packaging.sectionFound;

            if (!hasFabrics && !hasTrims && !hasPackaging) {
                return {
                    fileName: fileResult.fileName,
                    summary: 'Error: No sections found',
                    rows: [],
                    cellStatuses: []
                };
            }

            // Calculate totals across all sections
            let totalValid = 0;
            let totalInvalid = 0;

            // Count Fabrics
            if (hasFabrics && fileResult.results.fabrics.rows) {
                for (const rowResult of fileResult.results.fabrics.rows) {
                    for (const cellData of Object.values(rowResult.columns)) {
                        if (!cellData.isEmpty) {
                            if (cellData.isValid) totalValid++;
                            else totalInvalid++;
                        }
                    }
                }
            }

            // Count Trims
            if (hasTrims && fileResult.results.trims.rows) {
                for (const rowResult of fileResult.results.trims.rows) {
                    for (const cellData of Object.values(rowResult.columns)) {
                        if (!cellData.isEmpty) {
                            if (cellData.isValid) totalValid++;
                            else totalInvalid++;
                        }
                    }
                }
            }

            // Count Packaging
            if (hasPackaging && fileResult.results.packaging.rows) {
                for (const rowResult of fileResult.results.packaging.rows) {
                    for (const cellData of Object.values(rowResult.columns)) {
                        if (!cellData.isEmpty) {
                            if (cellData.isValid) totalValid++;
                            else totalInvalid++;
                        }
                    }
                }
            }

            // Build summary
            const summaryParts = [];
            if (hasFabrics) {
                summaryParts.push(`Fabrics: Rows ${fileResult.results.fabrics.startRow + 1}-${fileResult.results.fabrics.endRow}`);
            }
            if (hasTrims) {
                summaryParts.push(`Trims: Rows ${fileResult.results.trims.startRow + 1}-${fileResult.results.trims.endRow}`);
            }
            if (hasPackaging) {
                summaryParts.push(`Packaging: Rows ${fileResult.results.packaging.startRow + 1}-${fileResult.results.packaging.endRow}`);
            }

            // Get detected columns info
            const detectedCols = fileResult.results.detectedColumns;
            let detectedColsText = '';
            if (detectedCols) {
                const colInfo = [];
                if (detectedCols.wastage && detectedCols.wastage.found) colInfo.push(`Wastage: ${detectedCols.wastage.column}`);
                if (detectedCols.overheadCost && detectedCols.overheadCost.found) colInfo.push(`Overhead: ${detectedCols.overheadCost.column}`);
                if (detectedCols.testingCost && detectedCols.testingCost.found) colInfo.push(`Testing: ${detectedCols.testingCost.column}`);
                if (detectedCols.profitFOB && detectedCols.profitFOB.found) colInfo.push(`Profit: ${detectedCols.profitFOB.column}`);
                if (colInfo.length > 0) {
                    detectedColsText = ' | Detected: ' + colInfo.join(', ');
                }
            }

            // Build cell statuses for coloring - include Section column
            const cellStatuses = [];

            // Fabrics rows
            if (hasFabrics && fileResult.results.fabrics.rows) {
                for (const rowResult of fileResult.results.fabrics.rows) {
                    const rowStatuses = ['normal', 'normal']; // Section and Row number columns
                    for (const key of fabricsRuleKeys) {
                        const cellData = rowResult.columns[key];
                        if (!cellData || cellData.isEmpty) {
                            rowStatuses.push('normal');
                        } else if (cellData.isValid) {
                            rowStatuses.push('valid');
                        } else {
                            rowStatuses.push('invalid');
                        }
                    }
                    cellStatuses.push(rowStatuses);
                }
            }

            // Trims rows
            if (hasTrims && fileResult.results.trims.rows) {
                for (const rowResult of fileResult.results.trims.rows) {
                    const rowStatuses = ['normal', 'normal']; // Section and Row number columns
                    for (const key of fabricsRuleKeys) {
                        if (trimsRuleKeys.includes(key)) {
                            const cellData = rowResult.columns[key];
                            if (!cellData || cellData.isEmpty) {
                                rowStatuses.push('normal');
                            } else if (cellData.isValid) {
                                rowStatuses.push('valid');
                            } else {
                                rowStatuses.push('invalid');
                            }
                        } else {
                            rowStatuses.push('normal'); // dash columns
                        }
                    }
                    cellStatuses.push(rowStatuses);
                }
            }

            // Packaging rows
            if (hasPackaging && fileResult.results.packaging.rows) {
                for (const rowResult of fileResult.results.packaging.rows) {
                    const rowStatuses = ['normal', 'normal']; // Section and Row number columns
                    for (const key of fabricsRuleKeys) {
                        if (trimsRuleKeys.includes(key)) {
                            const cellData = rowResult.columns[key];
                            if (!cellData || cellData.isEmpty) {
                                rowStatuses.push('normal');
                            } else if (cellData.isValid) {
                                rowStatuses.push('valid');
                            } else {
                                rowStatuses.push('invalid');
                            }
                        } else {
                            rowStatuses.push('normal'); // dash columns
                        }
                    }
                    cellStatuses.push(rowStatuses);
                }
            }

            return {
                fileName: fileResult.fileName,
                summary: `${summaryParts.join(' | ')}${detectedColsText} | ${totalValid} passed, ${totalInvalid} failed`,
                fabricsRows: hasFabrics ? fileResult.results.fabrics.rows : [],
                trimsRows: hasTrims ? fileResult.results.trims.rows : [],
                packagingRows: hasPackaging ? fileResult.results.packaging.rows : [],
                activeRules: fileResult.results.fabrics.activeRules,
                cellStatuses: cellStatuses
            };
        }),
        filenamePrefix: 'FootAsylum_Validation',
        columnWidths: [22, 18, 35, 38, 28, 32, 30, 28],
        headers: ['Section', 'Row', 'Main Material (D)', 'Supplier Currency (P)', 'Wastage %', 'Overhead Cost', 'Testing Cost', 'Profit %'],
        colorRules: {},
        extractRowData: (fileResult) => {
            const rows = [];

            const hasFabrics = fileResult.fabricsRows && fileResult.fabricsRows.length > 0;
            const hasTrims = fileResult.trimsRows && fileResult.trimsRows.length > 0;
            const hasPackaging = fileResult.packagingRows && fileResult.packagingRows.length > 0;

            if (!hasFabrics && !hasTrims && !hasPackaging) {
                rows.push(['Error', '', 'No sections found or no data rows', '', '', '', '', '']);
                return rows;
            }

            // Helper to format cell
            const formatCell = (cellData) => {
                if (!cellData || cellData.isEmpty) {
                    return '-';
                } else if (cellData.isValid) {
                    return cellData.value;
                } else {
                    return `${cellData.value} (Expected: ${cellData.expected})`;
                }
            };

            // Add Fabrics rows
            if (hasFabrics) {
                for (const rowResult of fileResult.fabricsRows) {
                    const rowData = ['Fabrics', rowResult.rowNumber.toString()];
                    for (const key of fabricsRuleKeys) {
                        rowData.push(formatCell(rowResult.columns[key]));
                    }
                    rows.push(rowData);
                }
            }

            // Add Trims rows
            if (hasTrims) {
                for (const rowResult of fileResult.trimsRows) {
                    const rowData = ['Trims', rowResult.rowNumber.toString()];
                    for (const key of fabricsRuleKeys) {
                        if (trimsRuleKeys.includes(key)) {
                            rowData.push(formatCell(rowResult.columns[key]));
                        } else {
                            rowData.push('-');
                        }
                    }
                    rows.push(rowData);
                }
            }

            // Add Packaging rows
            if (hasPackaging) {
                for (const rowResult of fileResult.packagingRows) {
                    const rowData = ['Packaging', rowResult.rowNumber.toString()];
                    for (const key of fabricsRuleKeys) {
                        if (trimsRuleKeys.includes(key)) {
                            rowData.push(formatCell(rowResult.columns[key]));
                        } else {
                            rowData.push('-');
                        }
                    }
                    rows.push(rowData);
                }
            }

            return rows;
        }
    };
}
