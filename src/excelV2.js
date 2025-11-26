/**
 * Excel V2 Processing Logic
 * Automatically loads Burton_CostBreakdown.csv from public folder
 */

class ExcelV2Processor {
    constructor() {
        this.burtonCostData = null;
        this.bcbdResults = [];
    }

    /**
     * Initialize V2 - Load Burton Cost Breakdown CSV automatically
     */
    async initialize() {
        try {
            // Fetch the Burton_CostBreakdown.csv file from public folder
            const response = await fetch('../public/Burton_CostBreakdown.csv');
            if (!response.ok) {
                throw new Error('Failed to load Burton_CostBreakdown.csv');
            }
            
            const csvText = await response.text();
            this.burtonCostData = this.parseCSV(csvText);
            
            // Display the loaded data in the OB drop zone
            this.displayBurtonCostData();
            
            console.log('Burton Cost Breakdown loaded successfully:', this.burtonCostData);
        } catch (error) {
            console.error('Error loading Burton Cost Breakdown:', error);
            this.displayError('Failed to load Burton_CostBreakdown.csv from public folder');
        }
    }

    /**
     * Parse CSV text into array of objects
     */
    parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        const data = [];
        
        lines.forEach(line => {
            // Split by comma but handle quoted values
            const values = line.split(',').map(val => val.trim());
            data.push({
                description: values[0] || '',
                details: values[1] || '',
                materialName: values[2] || '',
                supplier: values[3] || '',
                quantity: values[4] || '',
                wastage: values[5] || '',
                unit: values[6] || '',
                unitPrice: values[7] || '',
                totalPrice: values[8] || ''
            });
        });
        
        return data;
    }

    /**
     * Display Burton Cost Breakdown data in the OB drop zone
     */
    displayBurtonCostData() {
        const obDropZone = document.getElementById('obDropZone-v2');
        if (!obDropZone) return;

        // Replace the drop zone content with the Burton Cost data display
        let contentHTML = `
            <div class="burton-cost-container">
                <div class="burton-cost-items">
        `;

        // Display each line from the CSV
        this.burtonCostData.forEach((item, index) => {
            contentHTML += `
                <div class="burton-cost-item">
                    <div class="burton-item-line"><strong>${item.description}</strong></div>
                    ${item.details ? `<div class="burton-item-line"><strong>Details:</strong> ${item.details}</div>` : ''}
                    <div class="burton-item-line"><strong>Material:</strong> ${item.materialName}</div>
                    <div class="burton-item-line"><strong>Supplier:</strong> ${item.supplier}</div>
                    <div class="burton-item-line"><strong>Qty:</strong> ${item.quantity}</div>
                    <div class="burton-item-line"><strong>Wastage:</strong> ${item.wastage}</div>
                    <div class="burton-item-line"><strong>Unit:</strong> ${item.unit}</div>
                    <div class="burton-item-line"><strong>Unit Price:</strong> ${item.unitPrice}</div>
                    <div class="burton-item-line"><strong>Total:</strong> ${item.totalPrice}</div>
                </div>
            `;
        });

        contentHTML += `
                </div>
            </div>
        `;

        obDropZone.innerHTML = contentHTML;
    }

    /**
     * Display error message in the OB drop zone
     */
    displayError(errorMessage) {
        const obDropZone = document.getElementById('obDropZone-v2');
        if (!obDropZone) return;

        obDropZone.innerHTML = `
            <div class="drop-zone-content">
                <div style="background: #fee; border-left: 4px solid #dc3545; padding: 1.5rem; border-radius: 8px;">
                    <p style="color: #dc3545; font-weight: 600; margin-bottom: 0.5rem;">
                        ❌ Error Loading File
                    </p>
                    <p style="color: #721c24; font-size: 0.95rem;">
                        ${errorMessage}
                    </p>
                </div>
            </div>
        `;
    }

    /**
     * Process files and generate results
     */
    async processFiles(bcbdFiles) {
        this.bcbdResults = [];

        try {
            if (!this.burtonCostData || this.burtonCostData.length === 0) {
                return this.generateErrorHTML('Burton Cost Breakdown data not loaded');
            }

            if (!bcbdFiles || bcbdFiles.length === 0) {
                return this.generateErrorHTML('Please upload Buyer CBD files');
            }

            // Process BCBD files and compare with Burton Cost data
            // This is a placeholder for future implementation
            return this.generateResultsHTML([]);

        } catch (error) {
            console.error('Error processing files:', error);
            return this.generateErrorHTML(error.message);
        }
    }

    /**
     * Generate HTML for results display
     */
    generateResultsHTML(results) {
        return `
            <div style="text-align: center; padding: 2rem; color: #2b4a6c;">
                <p style="font-size: 1.3em; margin-bottom: 10px;">✓ Burton Cost Breakdown Loaded</p>
                <p>Ready for processing. Upload Buyer CBD files to continue.</p>
                <p style="margin-top: 15px; font-size: 0.9em; color: #7a92ab;">
                    Loaded ${this.burtonCostData ? this.burtonCostData.length : 0} items from Burton_CostBreakdown.csv
                </p>
            </div>
        `;
    }

    /**
     * Generate error HTML
     */
    generateErrorHTML(errorMessage) {
        return `
            <div style="background: #fee; border-left: 4px solid #dc3545; padding: 1.5rem; border-radius: 8px;">
                <p style="color: #dc3545; font-weight: 600; margin-bottom: 0.5rem;">
                    ❌ Error Processing Files
                </p>
                <p style="color: #721c24; font-size: 0.95rem;">
                    ${errorMessage}
                </p>
            </div>
        `;
    }
}

// Initialize the processor
window.excelV2Processor = new ExcelV2Processor();

// Auto-load Burton Cost Breakdown when V2 tab is activated
document.addEventListener('DOMContentLoaded', () => {
    // Check if we're on V2 tab and initialize
    const v2Tab = document.querySelector('[data-tab="v2"]');
    if (v2Tab) {
        v2Tab.addEventListener('click', () => {
            if (!window.excelV2Processor.burtonCostData) {
                window.excelV2Processor.initialize();
            }
        });
    }

    // If V2 tab is already active on load, initialize immediately
    const v2TabContent = document.getElementById('tab-v2');
    if (v2TabContent && v2TabContent.classList.contains('active')) {
        window.excelV2Processor.initialize();
    }
});
