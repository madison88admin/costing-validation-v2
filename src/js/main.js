/**
 * Costing Validation - Main JavaScript
 * Core logic for drag-and-drop Excel file handling with multiple versions
 */

class ExcelFileHandler {
    constructor(version) {
        this.version = version;
        this.obFiles = [];
        this.bcbdFiles = [];
        this.initializeElements();
        this.attachEventListeners();
    }

    initializeElements() {
        this.obDropZone = document.getElementById(`obDropZone-${this.version}`);
        this.obFileInput = document.getElementById(`obFileInput-${this.version}`);
        this.obFileList = document.getElementById(`obFileList-${this.version}`);
        this.bcbdDropZone = document.getElementById(`bcbdDropZone-${this.version}`);
        this.bcbdFileInput = document.getElementById(`bcbdFileInput-${this.version}`);
        this.bcbdFileList = document.getElementById(`bcbdFileList-${this.version}`);
    }

    attachEventListeners() {
        // For V2-V15, don't setup OB drop zone (Cost Breakdown is auto-loaded from CSV or validation rules)
        if (this.version !== 'v2' && this.version !== 'v3' && this.version !== 'v4' && this.version !== 'v5' && this.version !== 'v6' && this.version !== 'v7' && this.version !== 'v8' && this.version !== 'v9' && this.version !== 'v10' && this.version !== 'v11' && this.version !== 'v12' && this.version !== 'v13' && this.version !== 'v14' && this.version !== 'v15') {
            this.setupDropZone(this.obDropZone, this.obFileInput, 'ob');
        }
        this.setupDropZone(this.bcbdDropZone, this.bcbdFileInput, 'bcbd');

        // Prevent default drag behavior on document
        document.addEventListener('dragover', (e) => e.preventDefault());
        document.addEventListener('drop', (e) => e.preventDefault());
    }

    setupDropZone(dropZone, fileInput, type) {
        dropZone.addEventListener('dragover', (e) => this.handleDragOver(e, dropZone));
        dropZone.addEventListener('dragleave', (e) => this.handleDragLeave(e, dropZone));
        dropZone.addEventListener('drop', (e) => this.handleDrop(e, dropZone, type));
        dropZone.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e, type));
    }

    handleDragOver(e, dropZone) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('drag-over');
    }

    handleDragLeave(e, dropZone) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('drag-over');
    }

    handleDrop(e, dropZone, type) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('drag-over');

        const files = Array.from(e.dataTransfer.files);
        if (files.length > 0) {
            this.processFiles(files, type);
        }
    }

    handleFileSelect(e, type) {
        const files = Array.from(e.target.files);
        if (files.length > 0) {
            this.processFiles(files, type);
        }
    }

    processFiles(files, type) {
        const fileArray = type === 'ob' ? this.obFiles : this.bcbdFiles;

        files.forEach(file => {
            if (!this.isValidFileType(file)) {
                alert(`Invalid file type: ${file.name}. Please select .xlsx, .xls, or .csv files.`);
                return;
            }

            // Check for duplicates
            if (!fileArray.some(f => f.name === file.name && f.size === file.size)) {
                fileArray.push(file);
            }
        });

        this.updateFileList(type);
        console.log(`${this.version.toUpperCase()} - ${type.toUpperCase()} files:`, fileArray);
    }

    updateFileList(type) {
        const fileList = type === 'ob' ? this.obFileList : this.bcbdFileList;
        const files = type === 'ob' ? this.obFiles : this.bcbdFiles;
        const dropZone = type === 'ob' ? this.obDropZone : this.bcbdDropZone;

        fileList.innerHTML = '';

        files.forEach((file, index) => {
            const fileItem = this.createFileItem(file, type, index);
            fileList.appendChild(fileItem);
        });

        dropZone.classList.toggle('has-file', files.length > 0);
    }

    createFileItem(file, type, index) {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';

        const fileContent = document.createElement('div');
        fileContent.className = 'file-item-content';

        const fileIcon = document.createElement('span');
        fileIcon.className = 'file-item-icon';
        fileIcon.innerHTML = '✓';

        const fileName = document.createElement('span');
        fileName.className = 'file-item-name';
        fileName.textContent = file.name;
        fileName.title = file.name;

        fileContent.appendChild(fileIcon);
        fileContent.appendChild(fileName);

        const removeBtn = document.createElement('button');
        removeBtn.className = 'file-item-remove';
        removeBtn.textContent = 'Remove';
        removeBtn.onclick = (e) => {
            e.stopPropagation();
            this.removeFile(type, index);
        };

        fileItem.appendChild(fileContent);
        fileItem.appendChild(removeBtn);

        return fileItem;
    }

    removeFile(type, index) {
        if (type === 'ob') {
            this.obFiles.splice(index, 1);
        } else {
            this.bcbdFiles.splice(index, 1);
        }

        this.updateFileList(type);
        console.log(`File removed from ${this.version.toUpperCase()} - ${type.toUpperCase()}`);
    }

    isValidFileType(file) {
        const validTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel',
            'text/csv'
        ];

        const validExtensions = ['.xlsx', '.xls', '.csv'];
        const fileName = file.name.toLowerCase();

        return validTypes.includes(file.type) ||
            validExtensions.some(ext => fileName.endsWith(ext));
    }

    getOBFiles() {
        return this.obFiles;
    }

    getBCBDFiles() {
        return this.bcbdFiles;
    }

    areBothFilesLoaded() {
        return this.obFiles.length > 0 && this.bcbdFiles.length > 0;
    }

    reset() {
        this.obFiles = [];
        this.bcbdFiles = [];
        this.obFileInput.value = '';
        this.bcbdFileInput.value = '';
        this.updateFileList('ob');
        this.updateFileList('bcbd');
    }
}

// Tab Management
class TabManager {
    constructor() {
        this.tabs = document.querySelectorAll('.tab-btn');
        this.tabContents = document.querySelectorAll('.tab-content');
        this.attachEventListeners();
        this.initializeActiveTab();
    }

    initializeActiveTab() {
        // Set the active tab to show full text on page load
        this.tabs.forEach(tab => {
            if (tab.classList.contains('active')) {
                const fullText = tab.getAttribute('data-full');
                if (fullText) {
                    tab.textContent = fullText;
                }
            } else {
                const shortText = tab.getAttribute('data-short');
                if (shortText) {
                    tab.textContent = shortText;
                }
            }
        });
    }

    attachEventListeners() {
        this.tabs.forEach(tab => {
            tab.addEventListener('click', () => this.switchTab(tab.dataset.tab));

            // Add hover listeners to change text
            tab.addEventListener('mouseenter', () => {
                const fullText = tab.getAttribute('data-full');
                if (fullText) {
                    tab.textContent = fullText;
                }
            });

            tab.addEventListener('mouseleave', () => {
                // Only collapse if not active
                if (!tab.classList.contains('active')) {
                    const shortText = tab.getAttribute('data-short');
                    if (shortText) {
                        tab.textContent = shortText;
                    }
                }
            });
        });
    }

    switchTab(tabId) {
        // Reset all tabs to short text and remove active class
        this.tabs.forEach(tab => {
            tab.classList.remove('active');
            const shortText = tab.getAttribute('data-short');
            if (shortText) {
                tab.textContent = shortText;
            }
        });

        this.tabContents.forEach(content => content.classList.remove('active'));

        // Use more specific selector to get the tab button, not the menu item
        const selectedTab = document.querySelector(`.tab-btn[data-tab="${tabId}"]`);
        const selectedContent = document.getElementById(`tab-${tabId}`);

        if (selectedTab && selectedContent) {
            selectedTab.classList.add('active');
            selectedContent.classList.add('active');

            // Set active tab to full text
            const fullText = selectedTab.getAttribute('data-full');
            if (fullText) {
                selectedTab.textContent = fullText;
            }
        }

        // Update menu active state
        if (window.menuManager) {
            window.menuManager.updateActiveMenuItem(tabId);
        }

        console.log(`Switched to ${tabId.toUpperCase()}`);
    }
}

// Menu Manager
class MenuManager {
    constructor() {
        this.menuToggle = document.getElementById('menuToggle');
        this.menuDropdown = document.getElementById('menuDropdown');
        this.menuItems = document.querySelectorAll('.menu-item');
        this.init();
    }

    init() {
        // Toggle menu on button click
        if (this.menuToggle) {
            this.menuToggle.addEventListener('click', (e) => {
                e.stopPropagation();
                this.toggleMenu();
            });
        }

        // Handle menu item clicks
        this.menuItems.forEach(item => {
            item.addEventListener('click', (e) => {
                const tabId = item.getAttribute('data-tab');
                if (tabId && window.tabManager) {
                    window.tabManager.switchTab(tabId);
                    this.updateActiveMenuItem(tabId);
                    this.closeMenu();
                }
            });
        });

        // Close menu when clicking outside
        document.addEventListener('click', (e) => {
            if (this.menuDropdown &&
                !this.menuDropdown.contains(e.target) &&
                !this.menuToggle.contains(e.target)) {
                this.closeMenu();
            }
        });

        // Close menu on escape key
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                this.closeMenu();
            }
        });

        // Set initial active state
        this.updateActiveMenuItem('v1');
    }

    updateActiveMenuItem(tabId) {
        // Remove active class from all menu items
        this.menuItems.forEach(item => {
            item.classList.remove('active');
        });

        // Add active class to the selected menu item
        const activeItem = Array.from(this.menuItems).find(item => item.getAttribute('data-tab') === tabId);
        if (activeItem) {
            activeItem.classList.add('active');
        }
    }

    toggleMenu() {
        if (this.menuDropdown) {
            this.menuDropdown.classList.toggle('active');
        }
    }

    closeMenu() {
        if (this.menuDropdown) {
            this.menuDropdown.classList.remove('active');
        }
    }

    openMenu() {
        if (this.menuDropdown) {
            this.menuDropdown.classList.add('active');
        }
    }
}

// Dark Mode Manager
class DarkModeManager {
    constructor() {
        this.darkModeToggle = document.getElementById('darkModeToggle');
        this.isDarkMode = this.loadDarkModePreference();
        this.init();
    }

    init() {
        // Apply saved preference
        if (this.isDarkMode) {
            document.body.classList.add('dark-mode');
            this.updateToggleIcon();
        }

        // Attach event listener
        if (this.darkModeToggle) {
            this.darkModeToggle.addEventListener('click', () => this.toggle());
        }
    }

    toggle() {
        this.isDarkMode = !this.isDarkMode;
        document.body.classList.toggle('dark-mode');
        this.saveDarkModePreference();
        this.updateToggleIcon();
        console.log(`Dark mode ${this.isDarkMode ? 'enabled' : 'disabled'}`);
    }

    updateToggleIcon() {
        if (this.darkModeToggle) {
            if (this.isDarkMode) {
                // Sun icon for light mode
                this.darkModeToggle.innerHTML = `
                    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <circle cx="12" cy="12" r="5"></circle>
                        <line x1="12" y1="1" x2="12" y2="3"></line>
                        <line x1="12" y1="21" x2="12" y2="23"></line>
                        <line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line>
                        <line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line>
                        <line x1="1" y1="12" x2="3" y2="12"></line>
                        <line x1="21" y1="12" x2="23" y2="12"></line>
                        <line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line>
                        <line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line>
                    </svg>
                `;
            } else {
                // Moon icon for dark mode
                this.darkModeToggle.innerHTML = `
                    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path>
                    </svg>
                `;
            }
        }
    }

    saveDarkModePreference() {
        try {
            localStorage.setItem('darkMode', this.isDarkMode ? 'enabled' : 'disabled');
        } catch (e) {
            console.warn('Could not save dark mode preference:', e);
        }
    }

    loadDarkModePreference() {
        try {
            const savedMode = localStorage.getItem('darkMode');
            return savedMode === 'enabled';
        } catch (e) {
            console.warn('Could not load dark mode preference:', e);
            return false;
        }
    }
}


// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    window.menuManager = new MenuManager();
    window.darkModeManager = new DarkModeManager();
    window.tabManager = new TabManager();
    window.logoEasterEgg = new LogoEasterEgg();
    window.excelHandlerV1 = new ExcelFileHandler('v1');
    window.excelHandlerV2 = new ExcelFileHandler('v2');
    window.excelHandlerV3 = new ExcelFileHandler('v3');
    window.excelHandlerV4 = new ExcelFileHandler('v4');
    window.excelHandlerV5 = new ExcelFileHandler('v5');
    window.excelHandlerV6 = new ExcelFileHandler('v6');
    window.excelHandlerV7 = new ExcelFileHandler('v7');
    window.excelHandlerV8 = new ExcelFileHandler('v8');
    window.excelHandlerV9 = new ExcelFileHandler('v9');
    window.excelHandlerV10 = new ExcelFileHandler('v10');
    window.excelHandlerV11 = new ExcelFileHandler('v11');
    window.excelHandlerV12 = new ExcelFileHandler('v12');
    window.excelHandlerV13 = new ExcelFileHandler('v13');
    window.excelHandlerV14 = new ExcelFileHandler('v14');
    window.excelHandlerV15 = new ExcelFileHandler('v15');

    document.querySelectorAll('.generate-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const version = e.target.dataset.version;
            handleGenerateResults(version);
        });
    });

    // Initialize LLBEAN processor when page loads
    if (window.llbeanProcessor) {
        window.llbeanProcessor.initialize();
    }

    // Initialize Mammut processor when page loads
    if (window.mammutProcessor) {
        window.mammutProcessor.initialize();
    }

    // Initialize Outdoor Research processor when page loads
    if (window.outdoorResearchProcessor) {
        window.outdoorResearchProcessor.initialize();
    }

    // Initialize Peak Performance processor when page loads
    if (window.peakPerformanceProcessor) {
        window.peakPerformanceProcessor.initialize();
    }

    // Initialize Skida processor when page loads
    if (window.skidaProcessor) {
        window.skidaProcessor.initialize();
    }

    // Initialize Vuori processor when page loads
    if (window.vuoriProcessor) {
        window.vuoriProcessor.initialize();
    }

    // Initialize Prana processor when page loads
    if (window.pranaProcessor) {
        window.pranaProcessor.initialize();
    }

    // Initialize Jack Wolfskin processor when page loads
    if (window.jackWolfskinProcessor) {
        window.jackWolfskinProcessor.initialize();
    }

    console.log('Costing Validation initialized with 15 versions');
});

async function handleGenerateResults(version) {
    const handler = window[`excelHandler${version.toUpperCase()}`];
    const resultsContent = document.getElementById(`results-${version}`);

    // Special handling for V2 - only needs BCBD files (Burton)
    if (version === 'v2') {
        const bcbdFiles = handler.getBCBDFiles();

        if (bcbdFiles.length === 0) {
            alert('Please upload Buyer CBD files before generating results.');
            return;
        }

        console.log(`Generating results for ${version.toUpperCase()}...`);
        console.log('BCBD Files:', bcbdFiles);

        // Show loading state with animation
        resultsContent.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <p class="loading-text">Processing ${bcbdFiles.length} BCBD file(s) with Burton Cost Breakdown...</p>
                <p class="loading-subtext">Please wait while we scan the files...</p>
            </div>
        `;

        if (window.excelV2Processor) {
            const results = await window.excelV2Processor.processFiles(bcbdFiles);
            resultsContent.innerHTML = results;
        }
        return;
    }

    // Special handling for V3 - only needs BCBD files (Columbia)
    if (version === 'v3') {
        const bcbdFiles = handler.getBCBDFiles();

        if (bcbdFiles.length === 0) {
            alert('Please upload Buyer CBD files before generating results.');
            return;
        }

        console.log(`Generating results for ${version.toUpperCase()}...`);
        console.log('BCBD Files:', bcbdFiles);

        // Show loading state with animation
        resultsContent.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <p class="loading-text">Processing ${bcbdFiles.length} BCBD file(s) with Columbia Cost Breakdown...</p>
                <p class="loading-subtext">Please wait while we scan the files...</p>
            </div>
        `;

        if (window.columbiaProcessor) {
            const results = await window.columbiaProcessor.processFiles(bcbdFiles);
            resultsContent.innerHTML = results;
        }
        return;
    }

    // Special handling for V4 - only needs BCBD files (Helly Hansen)
    if (version === 'v4') {
        const bcbdFiles = handler.getBCBDFiles();

        if (bcbdFiles.length === 0) {
            alert('Please upload Buyer CBD files before generating results.');
            return;
        }

        console.log(`Generating results for ${version.toUpperCase()}...`);
        console.log('BCBD Files:', bcbdFiles);

        // Show loading state with animation
        resultsContent.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <p class="loading-text">Processing ${bcbdFiles.length} BCBD file(s) with Helly Hansen Cost Breakdown...</p>
                <p class="loading-subtext">Please wait while we scan the files...</p>
            </div>
        `;

        if (window.hellyHansenProcessor) {
            const results = await window.hellyHansenProcessor.processFiles(bcbdFiles);
            resultsContent.innerHTML = results;
        }
        return;
    }

    // Special handling for V5 - only needs BCBD files (Fjall Raven)
    if (version === 'v5') {
        const bcbdFiles = handler.getBCBDFiles();

        if (bcbdFiles.length === 0) {
            alert('Please upload Buyer CBD files before generating results.');
            return;
        }

        console.log(`Generating results for ${version.toUpperCase()}...`);
        console.log('BCBD Files:', bcbdFiles);

        // Show loading state with animation
        resultsContent.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <p class="loading-text">Processing ${bcbdFiles.length} BCBD file(s) with Fjall Raven Cost Breakdown...</p>
                <p class="loading-subtext">Please wait while we scan the files...</p>
            </div>
        `;

        if (window.fjallRavenProcessor) {
            const results = await window.fjallRavenProcessor.processFiles(bcbdFiles);
            resultsContent.innerHTML = results;
        }
        return;
    }

    // Special handling for V6 - only needs BCBD files (LLBEAN)
    if (version === 'v6') {
        const bcbdFiles = handler.getBCBDFiles();

        if (bcbdFiles.length === 0) {
            alert('Please upload Buyer CBD files before generating results.');
            return;
        }

        console.log(`Generating results for ${version.toUpperCase()}...`);
        console.log('BCBD Files:', bcbdFiles);

        // Show loading state with animation
        resultsContent.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <p class="loading-text">Processing ${bcbdFiles.length} BCBD file(s) with LLBEAN validation...</p>
                <p class="loading-subtext">Please wait while we scan the files...</p>
            </div>
        `;

        if (window.llbeanProcessor) {
            const results = await window.llbeanProcessor.processFiles(bcbdFiles);
            resultsContent.innerHTML = results;
        }
        return;
    }

    // Special handling for V7 - only needs BCBD files (Mammut)
    if (version === 'v7') {
        const bcbdFiles = handler.getBCBDFiles();

        if (bcbdFiles.length === 0) {
            alert('Please upload Buyer CBD files before generating results.');
            return;
        }

        console.log(`Generating results for ${version.toUpperCase()}...`);
        console.log('BCBD Files:', bcbdFiles);

        // Show loading state with animation
        resultsContent.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <p class="loading-text">Processing ${bcbdFiles.length} BCBD file(s) with Mammut validation...</p>
                <p class="loading-subtext">Please wait while we scan the files...</p>
            </div>
        `;

        if (window.mammutProcessor) {
            const results = await window.mammutProcessor.processFiles(bcbdFiles);
            resultsContent.innerHTML = results;
        }
        return;
    }

    // Special handling for V8 - only needs BCBD files (Outdoor Research)
    if (version === 'v8') {
        const bcbdFiles = handler.getBCBDFiles();

        if (bcbdFiles.length === 0) {
            alert('Please upload Buyer CBD files before generating results.');
            return;
        }

        console.log(`Generating results for ${version.toUpperCase()}...`);
        console.log('BCBD Files:', bcbdFiles);

        // Show loading state with animation
        resultsContent.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <p class="loading-text">Processing ${bcbdFiles.length} BCBD file(s) with Outdoor Research validation...</p>
                <p class="loading-subtext">Please wait while we scan the files...</p>
            </div>
        `;

        if (window.outdoorResearchProcessor) {
            const results = await window.outdoorResearchProcessor.processFiles(bcbdFiles);
            resultsContent.innerHTML = results;
        }
        return;
    }

    // Special handling for V9 - only needs BCBD files (On AG)
    if (version === 'v9') {
        const bcbdFiles = handler.getBCBDFiles();

        if (bcbdFiles.length === 0) {
            alert('Please upload Buyer CBD files before generating results.');
            return;
        }

        console.log(`Generating results for ${version.toUpperCase()}...`);
        console.log('BCBD Files:', bcbdFiles);

        // Show loading state with animation
        resultsContent.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <p class="loading-text">Processing ${bcbdFiles.length} BCBD file(s) with On AG validation...</p>
                <p class="loading-subtext">Please wait while we scan the files...</p>
            </div>
        `;

        if (window.onAGProcessor) {
            const results = await window.onAGProcessor.processFiles(bcbdFiles);
            resultsContent.innerHTML = results;
        }
        return;
    }

    // Special handling for V10 - only needs BCBD files (Peak Performance)
    if (version === 'v10') {
        const bcbdFiles = handler.getBCBDFiles();

        if (bcbdFiles.length === 0) {
            alert('Please upload Buyer CBD files before generating results.');
            return;
        }

        console.log(`Generating results for ${version.toUpperCase()}...`);
        console.log('BCBD Files:', bcbdFiles);

        // Show loading state with animation
        resultsContent.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <p class="loading-text">Processing ${bcbdFiles.length} BCBD file(s) with Peak Performance validation...</p>
                <p class="loading-subtext">Please wait while we scan the files...</p>
            </div>
        `;

        if (window.peakPerformanceProcessor) {
            const results = await window.peakPerformanceProcessor.processFiles(bcbdFiles);
            resultsContent.innerHTML = results;
        }
        return;
    }

    // Special handling for V11 - only needs BCBD files (Skida)
    if (version === 'v11') {
        const bcbdFiles = handler.getBCBDFiles();

        if (bcbdFiles.length === 0) {
            alert('Please upload Buyer CBD files before generating results.');
            return;
        }

        console.log(`Generating results for ${version.toUpperCase()}...`);
        console.log('BCBD Files:', bcbdFiles);

        // Show loading state with animation
        resultsContent.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <p class="loading-text">Processing ${bcbdFiles.length} BCBD file(s) with Skida validation...</p>
                <p class="loading-subtext">Please wait while we scan the files...</p>
            </div>
        `;

        if (window.skidaProcessor) {
            const results = await window.skidaProcessor.processFiles(bcbdFiles);
            resultsContent.innerHTML = results;
        }
        return;
    }

    // Special handling for V12 - only needs BCBD files (Vuori)
    if (version === 'v12') {
        const bcbdFiles = handler.getBCBDFiles();

        if (bcbdFiles.length === 0) {
            alert('Please upload Buyer CBD files before generating results.');
            return;
        }

        console.log(`Generating results for ${version.toUpperCase()}...`);
        console.log('BCBD Files:', bcbdFiles);

        // Show loading state with animation
        resultsContent.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <p class="loading-text">Processing ${bcbdFiles.length} BCBD file(s) with Vuori validation...</p>
                <p class="loading-subtext">Please wait while we scan the files...</p>
            </div>
        `;

        if (window.vuoriProcessor) {
            const results = await window.vuoriProcessor.processFiles(bcbdFiles);
            resultsContent.innerHTML = results;
        }
        return;
    }

    // Special handling for V13 - only needs BCBD files (Prana)
    if (version === 'v13') {
        const bcbdFiles = handler.getBCBDFiles();

        if (bcbdFiles.length === 0) {
            alert('Please upload Buyer CBD files before generating results.');
            return;
        }

        console.log(`Generating results for ${version.toUpperCase()}...`);
        console.log('BCBD Files:', bcbdFiles);

        // Show loading state with animation
        resultsContent.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <p class="loading-text">Processing ${bcbdFiles.length} BCBD file(s) with Prana validation...</p>
                <p class="loading-subtext">Scanning all sheets for Fabrics section...</p>
            </div>
        `;

        if (window.pranaProcessor) {
            const results = await window.pranaProcessor.processFiles(bcbdFiles);
            resultsContent.innerHTML = results;
        }
        return;
    }

    // Special handling for V14 - only needs BCBD files (Travis Matthew)
    if (version === 'v14') {
        const bcbdFiles = handler.getBCBDFiles();

        if (bcbdFiles.length === 0) {
            alert('Please upload Buyer CBD files before generating results.');
            return;
        }

        console.log(`Generating results for ${version.toUpperCase()}...`);
        console.log('BCBD Files:', bcbdFiles);

        // Show loading state with animation
        resultsContent.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <p class="loading-text">Processing ${bcbdFiles.length} BCBD file(s) with Travis Matthew validation...</p>
                <p class="loading-subtext">Please wait while we scan the files...</p>
            </div>
        `;

        if (window.travisMatthewProcessor) {
            const results = await window.travisMatthewProcessor.processFiles(bcbdFiles);
            resultsContent.innerHTML = results;
        }
        return;
    }

    // Special handling for V15 - only needs BCBD files (Jack Wolfskin)
    if (version === 'v15') {
        const bcbdFiles = handler.getBCBDFiles();

        if (bcbdFiles.length === 0) {
            alert('Please upload Buyer CBD files before generating results.');
            return;
        }

        console.log(`Generating results for ${version.toUpperCase()}...`);
        console.log('BCBD Files:', bcbdFiles);

        // Show loading state with animation
        resultsContent.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <p class="loading-text">Processing ${bcbdFiles.length} BCBD file(s) with Jack Wolfskin validation...</p>
                <p class="loading-subtext">Please wait while we scan the files...</p>
            </div>
        `;

        if (window.jackWolfskinProcessor) {
            const results = await window.jackWolfskinProcessor.processFiles(bcbdFiles);
            resultsContent.innerHTML = results;
        }
        return;
    }

    // Standard handling for V1
    if (!handler.areBothFilesLoaded()) {
        alert('Please upload both OB and BCBD files before generating results.');
        return;
    }

    const obFiles = handler.getOBFiles();
    const bcbdFiles = handler.getBCBDFiles();

    console.log(`Generating results for ${version.toUpperCase()}...`);
    console.log('OB Files:', obFiles);
    console.log('BCBD Files:', bcbdFiles);

    // Show loading state with animation
    resultsContent.innerHTML = `
        <div class="loading-container">
            <div class="loader"></div>
            <p class="loading-text">Processing ${obFiles.length} OB file(s) and ${bcbdFiles.length} BCBD file(s)...</p>
            <p class="loading-subtext">Please wait while we scan the files...</p>
        </div>
    `;

    // Process based on version
    if (version === 'v1' && window.excelV1Processor) {
        const results = await window.excelV1Processor.processFiles(obFiles, bcbdFiles);
        resultsContent.innerHTML = results;
    } else {
        // Placeholder for V3
        resultsContent.innerHTML = `
            <div class="loading-container">
                <div class="loader"></div>
                <p class="loading-text">Processing ${obFiles.length} OB file(s) and ${bcbdFiles.length} BCBD file(s)...</p>
                <p class="loading-subtext">Template-specific processing logic for ${version.toUpperCase()} will be implemented here.</p>
            </div>
        `;
    }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { ExcelFileHandler, TabManager };
}

class LogoEasterEgg {
    constructor() {
        this.logo = document.querySelector('.logo');
        this.speechBubble = null;
        this.timeout = null;
        this.init();
    }

    init() {
        if (!this.logo) return;

        // Create speech bubble element
        this.speechBubble = document.createElement('div');
        this.speechBubble.className = 'speech-bubble';
        this.speechBubble.textContent = 'Benjamin was here!';
        this.logo.appendChild(this.speechBubble);

        // Add click listener
        this.logo.addEventListener('click', () => this.showBubble());
    }

    showBubble() {
        if (this.timeout) {
            clearTimeout(this.timeout);
        }
        this.speechBubble.classList.add('show');
        this.timeout = setTimeout(() => {
            this.speechBubble.classList.remove('show');
        }, 300);
    }
}
