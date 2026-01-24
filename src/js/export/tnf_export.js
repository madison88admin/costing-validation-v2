/**
 * Create TNF (V1) export configuration
 * @returns {Object} - Configuration object for TNF export
 */
function createTNFConfig() {
    return {
        title: 'Costing Validation Results - V1',
        tableId: 'v1ResultsTable',
        summarySelector: '#results-v1 div[style*="background: #f0f7ff"]',
        filenamePrefix: 'CostingValidation_V1',
        columnWidths: [35, 30, 25, 30, 25, 30, 35, 25],
        colorRules: {
            columns: {
                2: { // Match Status column
                    validator: (text) => {
                        if (text.includes('✓ FOUND')) return 'valid';
                        if (text.includes('✗ NOT FOUND')) return 'invalid';
                        return 'normal';
                    }
                },
                3: { // Standard Minute Value column
                    validator: (text) => {
                        if (text.includes('Empty') || text.includes('TNF: Empty')) return 'invalid';
                        if (text.includes('BCBD:')) {
                            const diffMatch = text.match(/\([\+\-]([\d.]+)\)/);
                            if (diffMatch) {
                                const diff = parseFloat(diffMatch[1]);
                                if (diff <= 0.01) return 'warning';
                                return 'invalid';
                            }
                            return 'warning';
                        }
                        if (text !== '-' && text !== '') return 'valid';
                        return 'normal';
                    }
                },
                4: { // Average Efficiency %
                    validator: (text) => {
                        if (text.includes('Cell Empty')) return 'invalid';
                        if (text === '-') return 'normal';
                        const match = text.match(/([\d.]+)%/);
                        if (match) {
                            const value = parseFloat(match[1]);
                            if (Math.abs(value - 50.0) < 0.1) return 'valid';
                            return 'invalid';
                        }
                        return 'normal';
                    }
                },
                5: { // Hourly Wages
                    validator: (text) => {
                        if (text.includes('Cell Empty')) return 'invalid';
                        if (text === '-') return 'normal';
                        const match = text.match(/([\d.]+)/);
                        if (match) {
                            const value = parseFloat(match[1]);
                            if (Math.abs(value - 1.750) < 0.01) return 'valid';
                            return 'invalid';
                        }
                        return 'normal';
                    }
                },
                6: { // Overhead Cost
                    validator: (text) => {
                        if (text.includes('Cell Empty')) return 'invalid';
                        if (text === '-') return 'normal';
                        const match = text.match(/([\d.]+)%/);
                        if (match) {
                            const value = parseFloat(match[1]);
                            if (Math.abs(value - 70.0) < 0.1) return 'valid';
                            return 'invalid';
                        }
                        return 'normal';
                    }
                },
                7: { // Factory Profit %
                    validator: (text) => {
                        if (text.includes('Cell Empty')) return 'invalid';
                        if (text === '-') return 'normal';
                        const match = text.match(/([\d.]+)%/);
                        if (match) {
                            const value = parseFloat(match[1]);
                            if (Math.abs(value - 10.0) < 0.1) return 'valid';
                            return 'warning';
                        }
                        return 'normal';
                    }
                }
            }
        }
    };
}
