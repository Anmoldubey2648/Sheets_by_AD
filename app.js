class Spreadsheet {
    constructor() {
        this.db = new SpreadsheetDB();
        this.formula = new FormulaHandler();
        this.activeCell = null;
        this.clipboard = null;
        this.selectedCells = new Set();
        this.dragStartCell = null;
        this.isDragging = false;
        this.dragHandle = null;
        this.selectionStart = null;
        this.lastSelectedCell = null;
        this.dragIndicator = null;
        this.dragMessage = null;
        this.floatingNav = null;
        this.batchCopyMessage = null;
        this.currentClockStyle = 'digital';
        this.calculatorValue = '0';
        this.lastOperator = null;
        this.lastNumber = null;
        this.newNumber = true;
        this.cellDependencies = new Map(); // Track cell dependencies
        this.dependentCells = new Map(); // Track cells that depend on each cell
        
        this.loadThemePreferences();
        this.initializeGrid();
        this.initializeSheetTabs();
        this.setupEventListeners();
        this.createDragHandle();
        this.createFloatingUI();
        this.createFunctionsMenu();
        this.createClock();
        this.createCalculator();
        
        // Add new properties for fill handle
        this.fillHandleActive = false;
        this.fillStartCell = null;
        this.fillPreview = document.createElement('div');
        this.fillPreview.className = 'fill-preview';
        document.body.appendChild(this.fillPreview);
        
        // Add header selection tracking
        this.selectedHeaders = new Set();
        this.createTestingInterface();
        
        // Check for file parameter in URL
        const urlParams = new URLSearchParams(window.location.search);
        const filePath = urlParams.get('file');
        if (filePath) {
            this.loadFileFromPath(filePath);
        }
        
        // Add undo/redo stacks
        this.undoStack = [];
        this.redoStack = [];
        this.maxStackSize = 100; // Limit stack size to prevent memory issues
    }

    loadThemePreferences() {
        const theme = localStorage.getItem('spreadsheet_theme') || 'default';
        const darkMode = localStorage.getItem('spreadsheet_dark_mode') === 'true';
        
        document.body.dataset.theme = theme;
        document.body.dataset.darkMode = darkMode;
        
        // Update UI controls
        document.getElementById('themeSelect').value = theme;
        document.getElementById('darkModeBtn').innerHTML = 
            darkMode ? '☀️ Light Mode' : '🌙 Dark Mode';
    }

    initializeGrid() {
        const table = document.getElementById('spreadsheet');
        table.innerHTML = '';
        
        // Create header row
        const headerRow = this.createHeaderRow();
        table.appendChild(headerRow);
        
        // Get row count from active sheet
        const sheet = this.db.data.sheets[this.db.getActiveSheet()];
        const rowCount = sheet ? sheet.rowCount : 50; // Default to 50 if sheet not found
        
        // Create data rows
        for (let row = 1; row <= rowCount; row++) {
            const tr = this.createDataRow(row);
            table.appendChild(tr);
        }
        
        // Create and add the column add button
        const addColBtn = document.createElement('button');
        addColBtn.innerHTML = '+';
        addColBtn.className = 'add-col-btn';
        addColBtn.title = 'Add Column';
        addColBtn.onclick = () => this.addColumn();
        document.body.appendChild(addColBtn);
        
        // Create and add the row add button
        const addRowBtn = document.createElement('button');
        addRowBtn.innerHTML = '+';
        addRowBtn.className = 'add-row-btn';
        addRowBtn.title = 'Add Row';
        addRowBtn.onclick = () => this.addRow();
        document.body.appendChild(addRowBtn);
        
        // Add this at the end of initializeGrid
        this.setupSpeechTooltips();
        
        // Add header click handlers
        const headers = document.querySelectorAll('#spreadsheet th');
        headers.forEach(header => {
            header.addEventListener('click', (e) => this.handleHeaderClick(e));
        });
    }

    createHeaderRow() {
        const tr = document.createElement('tr');
        
        // Corner cell
        const corner = document.createElement('th');
        corner.className = 'corner-header';
        tr.appendChild(corner);
        
        // Get actual column count from database
        const columnCount = this.db.getColumnCount();
        
        // Column headers A-Z, then AA, AB, etc.
        for (let i = 0; i < columnCount; i++) {
            const th = document.createElement('th');
            th.textContent = this.getColumnLabel(i);
            th.dataset.col = this.getColumnLabel(i);
            
            // Add resize handle
            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'resize-handle col-resize';
            resizeHandle.addEventListener('mousedown', (e) => {
                e.stopPropagation();
                this.startColumnResize(e, i);
            });
            th.appendChild(resizeHandle);
            
            tr.appendChild(th);
        }
        
        return tr;
    }

    createDataRow(rowNum) {
            const tr = document.createElement('tr');
            
        // Row header with resize handle
            const th = document.createElement('th');
        th.textContent = rowNum;
        th.dataset.row = rowNum;
        
        // Add row resize handle
        const resizeHandle = document.createElement('div');
        resizeHandle.className = 'resize-handle row-resize';
        resizeHandle.addEventListener('mousedown', (e) => {
            e.stopPropagation();
            this.startRowResize(e, rowNum);
        });
        th.appendChild(resizeHandle);
        
            tr.appendChild(th);
            
        // Get actual column count from database
        const columnCount = this.db.getColumnCount();
            
            // Data cells
        for (let i = 0; i < columnCount; i++) {
                const td = document.createElement('td');
            const cellId = `${this.getColumnLabel(i)}${rowNum}`;
                
            const input = document.createElement('input');
            input.type = 'text';
                input.className = 'cell';
                input.dataset.cellId = cellId;
                
                td.appendChild(input);
                tr.appendChild(td);
            
            // Apply any existing cell data
            const cellData = this.db.getCell(this.db.getActiveSheet(), cellId);
            if (cellData) {
                this.applyCellData(input, cellData);
            }
        }
        
        return tr;
    }

    applyCellData(input, cell) {
        if (!cell) return;
        
        input.value = cell.content;
        Object.assign(input.style, {
            fontWeight: cell.style.bold ? 'bold' : 'normal',
            fontStyle: cell.style.italic ? 'italic' : 'normal',
            textDecoration: cell.style.underline ? 'underline' : 'none',
            textAlign: cell.style.alignment,
            fontFamily: cell.style.fontFamily,
            fontSize: cell.style.fontSize + 'px',
            color: cell.style.color,
            backgroundColor: cell.style.backgroundColor
        });
    }

    initializeSheetTabs() {
        const sheetTabs = document.getElementById('sheetTabs');
        sheetTabs.innerHTML = '';
        
        this.db.getAllSheets().forEach(sheetName => {
            const tab = document.createElement('button');
            tab.className = 'sheet-tab';
            if (sheetName === this.db.getActiveSheet()) {
                tab.classList.add('active');
            }

            // Add sheet name span
            const nameSpan = document.createElement('span');
            nameSpan.textContent = sheetName;
            tab.appendChild(nameSpan);

            // Add close button
            const closeBtn = document.createElement('button');
            closeBtn.className = 'close-btn';
            closeBtn.innerHTML = '×';
            closeBtn.title = 'Close Sheet';
            
            // Prevent sheet switching when clicking close button
            closeBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                if (this.db.deleteSheet(sheetName)) {
                    this.initializeGrid();
                    this.initializeSheetTabs();
                }
            });

            tab.appendChild(closeBtn);
            sheetTabs.appendChild(tab);
        });
    }

    setupEventListeners() {
        // Theme controls
        document.getElementById('themeSelect').addEventListener('change', (e) => {
            const theme = e.target.value;
            document.body.dataset.theme = theme;
            localStorage.setItem('spreadsheet_theme', theme);
        });

        document.getElementById('darkModeBtn').addEventListener('click', () => {
            const isDarkMode = document.body.dataset.darkMode === 'true';
            document.body.dataset.darkMode = !isDarkMode;
            localStorage.setItem('spreadsheet_dark_mode', !isDarkMode);
            document.getElementById('darkModeBtn').innerHTML = 
                !isDarkMode ? '☀️ Light Mode' : '🌙 Dark Mode';
        });

        // Save/Load/Export handlers
        document.getElementById('saveBtn').addEventListener('click', () => {
            this.db.saveToFile();
        });

        document.getElementById('loadBtn').addEventListener('click', () => {
            document.getElementById('fileInput').click();
        });

        document.getElementById('fileInput').addEventListener('change', async (e) => {
            if (e.target.files.length > 0) {
                try {
                    await this.db.loadFromFile(e.target.files[0]);
                    this.initializeGrid();
                    this.initializeSheetTabs();
                } catch (error) {
                    alert('Error loading file: ' + error.message);
                }
                e.target.value = '';
            }
        });

        document.getElementById('exportBtn').addEventListener('click', () => {
            this.db.exportToCSV();
        });

        // Cell focus and input
        document.getElementById('spreadsheet').addEventListener('focusin', (e) => {
            if (e.target.classList.contains('cell')) {
                this.activeCell = e.target;
                document.getElementById('activeCell').textContent = e.target.dataset.cellId;
                const cell = this.db.getCell(this.db.getActiveSheet(), e.target.dataset.cellId);
                document.getElementById('formulaInput').value = cell?.formula || e.target.value;
            }
        });

        // Cell content change
        document.getElementById('spreadsheet').addEventListener('input', (e) => {
            if (e.target.classList.contains('cell')) {
                const cellId = e.target.dataset.cellId;
                const content = e.target.value;
                
                this.db.updateCell(this.db.getActiveSheet(), cellId, {
                    content: content,
                    formula: content.startsWith('=') ? content : ''
                });
                
                if (content.startsWith('=')) {
                    const result = this.formula.evaluate(content, (ref) => {
                        const cell = this.db.getCell(this.db.getActiveSheet(), ref);
                        return cell ? cell.content : '';
                    });
                    e.target.value = result;
                }
            }
        });

        // Formula input
        document.getElementById('formulaInput').addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && this.activeCell) {
                const formula = e.target.value;
                const cellId = this.activeCell.dataset.cellId;
                const sheetName = this.db.getActiveSheet();
                
                this.db.updateCell(sheetName, cellId, {
                    formula: formula,
                    content: this.formula.evaluate(formula, (ref) => {
                        const cell = this.db.getCell(sheetName, ref);
                        return cell ? cell.content : '';
                    })
                });
                
                this.applyCellData(
                    this.activeCell,
                    this.db.getCell(sheetName, cellId)
                );
            }
        });

        // Add input event listener for formula input
        document.getElementById('formulaInput').addEventListener('input', (e) => {
            if (this.activeCell) {
                const formula = e.target.value;
                const cellId = this.activeCell.dataset.cellId;
                const sheetName = this.db.getActiveSheet();
                
                // Update the cell's formula without evaluating
                this.db.updateCell(sheetName, cellId, {
                    formula: formula,
                    content: formula
                });
                
                // Show the formula in the active cell while editing
                this.activeCell.value = formula;
            }
        });

        // Add blur event listener for formula input
        document.getElementById('formulaInput').addEventListener('blur', (e) => {
            if (this.activeCell) {
                const formula = e.target.value;
                const cellId = this.activeCell.dataset.cellId;
                const sheetName = this.db.getActiveSheet();
                
                if (formula.startsWith('=')) {
                    const result = this.formula.evaluate(formula, (ref) => {
                        const cell = this.db.getCell(sheetName, ref);
                        return cell ? cell.content : '';
                    });
                    
                    this.db.updateCell(sheetName, cellId, {
                        formula: formula,
                        content: result
                    });
                    
                    this.activeCell.value = result;
                }
            }
        });

        // Style controls
        document.getElementById('bold').addEventListener('click', () => this.toggleStyle('bold'));
        document.getElementById('italic').addEventListener('click', () => this.toggleStyle('italic'));
        document.getElementById('underline').addEventListener('click', () => this.toggleStyle('underline'));
        document.getElementById('alignLeft').addEventListener('click', () => this.setAlignment('left'));
        document.getElementById('alignCenter').addEventListener('click', () => this.setAlignment('center'));
        document.getElementById('alignRight').addEventListener('click', () => this.setAlignment('right'));
        
        // Font controls
        document.getElementById('fontFamily').addEventListener('change', (e) => {
            if (!this.activeCell && this.selectedCells.size === 0) return;
            
            const value = e.target.value;
            const oldStyles = new Map();
            
            if (this.selectedCells.size > 0) {
                this.selectedCells.forEach(cell => {
                    oldStyles.set(cell, {
                        fontFamily: cell.style.fontFamily,
                        cellId: this.getCellId(cell)
                    });
                });
                
                this.selectedCells.forEach(cell => {
                    this.db.updateCellStyle(this.db.getActiveSheet(), this.getCellId(cell), { fontFamily: value });
                    cell.style.fontFamily = value;
                });

                // Add to undo stack
                this.addToHistory({
                    undo: () => {
                        oldStyles.forEach((style, cell) => {
                            this.db.updateCellStyle(this.db.getActiveSheet(), style.cellId, { fontFamily: style.fontFamily });
                            cell.style.fontFamily = style.fontFamily;
                        });
                    },
                    redo: () => {
                        oldStyles.forEach((style, cell) => {
                            this.db.updateCellStyle(this.db.getActiveSheet(), style.cellId, { fontFamily: value });
                            cell.style.fontFamily = value;
                        });
                    }
                });
            } else {
                const oldStyle = this.activeCell.style.fontFamily;
                const cellId = this.getCellId(this.activeCell);
                
                this.db.updateCellStyle(this.db.getActiveSheet(), cellId, { fontFamily: value });
                this.activeCell.style.fontFamily = value;

                // Add to undo stack
                this.addToHistory({
                    undo: () => {
                        this.db.updateCellStyle(this.db.getActiveSheet(), cellId, { fontFamily: oldStyle });
                        this.activeCell.style.fontFamily = oldStyle;
                    },
                    redo: () => {
                        this.db.updateCellStyle(this.db.getActiveSheet(), cellId, { fontFamily: value });
                        this.activeCell.style.fontFamily = value;
                    }
                });
            }
        });

        // Font Size
        document.getElementById('fontSize').addEventListener('change', (e) => {
            if (!this.activeCell && this.selectedCells.size === 0) return;
            
            const value = e.target.value;
            const oldStyles = new Map();
            
            if (this.selectedCells.size > 0) {
                this.selectedCells.forEach(cell => {
                    oldStyles.set(cell, {
                        fontSize: cell.style.fontSize,
                        cellId: this.getCellId(cell)
                    });
                });
                
                this.selectedCells.forEach(cell => {
                    this.db.updateCellStyle(this.db.getActiveSheet(), this.getCellId(cell), { fontSize: parseInt(value) });
                    cell.style.fontSize = value + 'px';
                });

                // Add to undo stack
                this.addToHistory({
                    undo: () => {
                        oldStyles.forEach((style, cell) => {
                            this.db.updateCellStyle(this.db.getActiveSheet(), style.cellId, { fontSize: parseInt(style.fontSize) });
                            cell.style.fontSize = style.fontSize;
                        });
                    },
                    redo: () => {
                        oldStyles.forEach((style, cell) => {
                            this.db.updateCellStyle(this.db.getActiveSheet(), style.cellId, { fontSize: parseInt(value) });
                            cell.style.fontSize = value + 'px';
                        });
                    }
                });
            } else {
                const oldStyle = this.activeCell.style.fontSize;
                const cellId = this.getCellId(this.activeCell);
                
                this.db.updateCellStyle(this.db.getActiveSheet(), cellId, { fontSize: parseInt(value) });
                this.activeCell.style.fontSize = value + 'px';

                // Add to undo stack
                this.addToHistory({
                    undo: () => {
                        this.db.updateCellStyle(this.db.getActiveSheet(), cellId, { fontSize: parseInt(oldStyle) });
                        this.activeCell.style.fontSize = oldStyle;
                    },
                    redo: () => {
                        this.db.updateCellStyle(this.db.getActiveSheet(), cellId, { fontSize: parseInt(value) });
                        this.activeCell.style.fontSize = value + 'px';
                    }
                });
            }
        });

        // Text Color
        document.getElementById('textColor').addEventListener('input', (e) => {
            if (!this.activeCell && this.selectedCells.size === 0) return;
            
            const value = e.target.value;
            const oldStyles = new Map();
            
            if (this.selectedCells.size > 0) {
                this.selectedCells.forEach(cell => {
                    oldStyles.set(cell, {
                        color: cell.style.color,
                        cellId: this.getCellId(cell)
                    });
                });
                
                this.selectedCells.forEach(cell => {
                    this.db.updateCellStyle(this.db.getActiveSheet(), this.getCellId(cell), { color: value });
                    cell.style.color = value;
                });

                // Add to undo stack
                this.addToHistory({
                    undo: () => {
                        oldStyles.forEach((style, cell) => {
                            this.db.updateCellStyle(this.db.getActiveSheet(), style.cellId, { color: style.color });
                            cell.style.color = style.color;
                        });
                    },
                    redo: () => {
                        oldStyles.forEach((style, cell) => {
                            this.db.updateCellStyle(this.db.getActiveSheet(), style.cellId, { color: value });
                            cell.style.color = value;
                        });
                    }
                });
            } else {
                const oldStyle = this.activeCell.style.color;
                const cellId = this.getCellId(this.activeCell);
                
                this.db.updateCellStyle(this.db.getActiveSheet(), cellId, { color: value });
                this.activeCell.style.color = value;

                // Add to undo stack
                this.addToHistory({
                    undo: () => {
                        this.db.updateCellStyle(this.db.getActiveSheet(), cellId, { color: oldStyle });
                        this.activeCell.style.color = oldStyle;
                    },
                    redo: () => {
                        this.db.updateCellStyle(this.db.getActiveSheet(), cellId, { color: value });
                        this.activeCell.style.color = value;
                    }
                });
            }
        });

        // Background Color
        document.getElementById('bgColor').addEventListener('input', (e) => {
            if (!this.activeCell && this.selectedCells.size === 0) return;
            
            const value = e.target.value;
            const oldStyles = new Map();
            
            if (this.selectedCells.size > 0) {
                this.selectedCells.forEach(cell => {
                    oldStyles.set(cell, {
                        backgroundColor: cell.style.backgroundColor,
                        cellId: this.getCellId(cell)
                    });
                });
                
                this.selectedCells.forEach(cell => {
                    this.db.updateCellStyle(this.db.getActiveSheet(), this.getCellId(cell), { backgroundColor: value });
                    cell.style.backgroundColor = value;
                });

                // Add to undo stack
                this.addToHistory({
                    undo: () => {
                        oldStyles.forEach((style, cell) => {
                            this.db.updateCellStyle(this.db.getActiveSheet(), style.cellId, { backgroundColor: style.backgroundColor });
                            cell.style.backgroundColor = style.backgroundColor;
                        });
                    },
                    redo: () => {
                        oldStyles.forEach((style, cell) => {
                            this.db.updateCellStyle(this.db.getActiveSheet(), style.cellId, { backgroundColor: value });
                            cell.style.backgroundColor = value;
                        });
                    }
                });
            } else {
                const oldStyle = this.activeCell.style.backgroundColor;
                const cellId = this.getCellId(this.activeCell);
                
                this.db.updateCellStyle(this.db.getActiveSheet(), cellId, { backgroundColor: value });
                this.activeCell.style.backgroundColor = value;

                // Add to undo stack
                this.addToHistory({
                    undo: () => {
                        this.db.updateCellStyle(this.db.getActiveSheet(), cellId, { backgroundColor: oldStyle });
                        this.activeCell.style.backgroundColor = oldStyle;
                    },
                    redo: () => {
                        this.db.updateCellStyle(this.db.getActiveSheet(), cellId, { backgroundColor: value });
                        this.activeCell.style.backgroundColor = value;
                    }
                });
            }
        });

        // Clipboard operations
        document.getElementById('copy').addEventListener('click', () => {
            if (this.activeCell) {
                const cellId = this.activeCell.dataset.cellId;
                const cell = this.db.getCell(this.db.getActiveSheet(), cellId);
                this.clipboard = {
                    content: cell?.content || '',
                    formula: cell?.formula || '',
                    style: cell?.style ? {...cell.style} : {}
                };
            }
        });

        document.getElementById('cut').addEventListener('click', () => {
            if (this.activeCell) {
                const cellId = this.activeCell.dataset.cellId;
                const cell = this.db.getCell(this.db.getActiveSheet(), cellId);
                this.clipboard = {
                    content: cell?.content || '',
                    formula: cell?.formula || '',
                    style: cell?.style ? {...cell.style} : {}
                };
                
                // Clear the cell's content and styles
                this.db.updateCell(this.db.getActiveSheet(), cellId, {
                    content: '',
                    formula: '',
                    style: {}
                });
                
                // Reset the cell's appearance
                this.activeCell.value = '';
                this.activeCell.style.fontWeight = 'normal';
                this.activeCell.style.fontStyle = 'normal';
                this.activeCell.style.textDecoration = 'none';
                this.activeCell.style.textAlign = 'left';
                this.activeCell.style.fontFamily = '';
                this.activeCell.style.fontSize = '';
                this.activeCell.style.color = '';
                this.activeCell.style.backgroundColor = '';
            }
        });

        document.getElementById('paste').addEventListener('click', () => {
            if (this.activeCell && this.clipboard) {
                const cellId = this.activeCell.dataset.cellId;
                
                // Update cell content and formula
                this.db.updateCell(this.db.getActiveSheet(), cellId, {
                    content: this.clipboard.content,
                    formula: this.clipboard.formula
                });
                this.db.updateCellStyle(this.db.getActiveSheet(), cellId, this.clipboard.style);
                this.applyCellData(this.activeCell, this.clipboard);
            }
        });

        // Sheet management
        document.getElementById('addSheet').addEventListener('click', () => {
            // Get all current sheet names
            const sheets = this.db.getAllSheets();
            
            // Generate a new sheet name
            let sheetNumber = sheets.length + 1;
            let newSheetName = `Sheet${sheetNumber}`;
            
            // Make sure we have a unique name
            while (sheets.includes(newSheetName)) {
                sheetNumber++;
                newSheetName = `Sheet${sheetNumber}`;
            }
            
            // Add the new sheet
            if (this.db.addSheet(newSheetName)) {
                // Switch to the new sheet
                this.db.switchSheet(newSheetName);
                // Update the UI
            this.initializeGrid();
            this.initializeSheetTabs();
            }
        });

        document.getElementById('sheetTabs').addEventListener('click', (e) => {
            const tab = e.target.closest('.sheet-tab');
            if (tab && !e.target.classList.contains('close-btn')) {
                const sheetName = tab.querySelector('span').textContent;
                if (this.db.switchSheet(sheetName)) {
                this.initializeGrid();
                this.initializeSheetTabs();
                }
            }
        });

        // Update cell selection and drag handling
        document.getElementById('spreadsheet').addEventListener('mousedown', (e) => {
            // Handle row header clicks (when clicking on row numbers)
            if (e.target.tagName === 'TH' && !e.target.querySelector('.grid-controls')) {
                const rowIndex = e.target.parentElement.rowIndex;
                if (rowIndex > 0) { // Skip header row
                    const cells = e.target.parentElement.querySelectorAll('.cell');
                    
                    if (!e.ctrlKey && !e.shiftKey) {
                        // Clear previous selection unless clicking on a selected row
                        const anySelected = Array.from(cells).some(cell => this.selectedCells.has(cell));
                        if (!anySelected) {
                            this.selectedCells.clear();
                            document.querySelectorAll('.cell.selected').forEach(cell => {
                                cell.classList.remove('selected');
                            });
                        }
                    }
                    
                    // Select all cells in the row
                    cells.forEach(cell => {
                        this.selectedCells.add(cell);
                        cell.classList.add('selected');
                        cell.setAttribute('readonly', 'true'); // Make cells readonly when selected
                    });
                    
                    this.lastSelectedCell = cells[cells.length - 1];
                    this.updateFloatingNav();
                }
                return;
            }
            
            // Handle regular cell clicks
            if (!e.target.classList.contains('cell')) return;
            
            const cell = e.target;
            
            // Handle drag handle click
            if (e.target === this.dragHandle) {
                this.isDragging = true;
                return;
            }
            
            // Start new selection or add to existing selection
            if (!e.ctrlKey && !e.shiftKey) {
                // Clear previous selection unless clicking on a selected cell
                if (!this.selectedCells.has(cell)) {
                    this.selectedCells.clear();
                    document.querySelectorAll('.cell.selected').forEach(cell => {
                        cell.classList.remove('selected');
                        cell.removeAttribute('readonly'); // Remove readonly when deselected
                    });
                }
            }
            
            if (e.shiftKey && this.lastSelectedCell) {
                // Extend selection
                const [startCol, startRow] = this.parseCellId(this.lastSelectedCell.dataset.cellId);
                const [endCol, endRow] = this.parseCellId(cell.dataset.cellId);
                
                const minCol = Math.min(startCol.charCodeAt(0), endCol.charCodeAt(0));
                const maxCol = Math.max(startCol.charCodeAt(0), endCol.charCodeAt(0));
                const minRow = Math.min(parseInt(startRow), parseInt(endRow));
                const maxRow = Math.max(parseInt(startRow), parseInt(endRow));
                
                // Clear previous selection
                this.selectedCells.clear();
                document.querySelectorAll('.cell.selected').forEach(cell => {
                    cell.classList.remove('selected');
                    cell.removeAttribute('readonly'); // Remove readonly when deselected
                });
                
                // Select all cells in range
                for (let col = minCol; col <= maxCol; col++) {
                    for (let row = minRow; row <= maxRow; row++) {
                        const cellId = String.fromCharCode(col) + row;
                        const cellElement = document.querySelector(`[data-cell-id="${cellId}"]`);
                        if (cellElement) {
                            this.selectedCells.add(cellElement);
                            cellElement.classList.add('selected');
                            cellElement.setAttribute('readonly', 'true'); // Make cells readonly when selected
                        }
                    }
                }
            } else {
                // Toggle selection for Ctrl+Click
                if (e.ctrlKey) {
                    if (this.selectedCells.has(cell)) {
                        this.selectedCells.delete(cell);
                        cell.classList.remove('selected');
                        cell.removeAttribute('readonly'); // Remove readonly when deselected
                    } else {
                        this.selectedCells.add(cell);
                        cell.classList.add('selected');
                        cell.setAttribute('readonly', 'true'); // Make cells readonly when selected
                    }
                } else {
                    // Single cell selection
                    if (!this.selectedCells.has(cell)) {
                        this.selectedCells.add(cell);
                        cell.classList.add('selected');
                        // Don't make readonly for single cell selection
                    }
                }
                this.lastSelectedCell = cell;
            }
            
            this.selectionStart = cell;
            this.updateFloatingNav();
            
            // Enable editing on double click
            if (e.detail === 2) {
                cell.focus();
                cell.removeAttribute('readonly'); // Remove readonly on double click
                cell.setSelectionRange(0, cell.value.length);
            }
        });

        document.addEventListener('mousemove', (e) => {
            if (!this.selectionStart && !this.isDragging) return;
            
            const target = document.elementFromPoint(e.clientX, e.clientY);
            if (!target || !target.classList.contains('cell')) return;
            
            if (this.isDragging) {
                // Handle dragging selected cells
                document.querySelectorAll('.cell.drag-over').forEach(cell => {
                    cell.classList.remove('drag-over');
                });
                
                if (!this.selectedCells.has(target)) {
                    target.classList.add('drag-over');
                    this.moveCellsToTarget(target);
                }
            } else if (!e.ctrlKey) {
                // Handle range selection
                const [startCol, startRow] = this.parseCellId(this.selectionStart.dataset.cellId);
                const [currentCol, currentRow] = this.parseCellId(target.dataset.cellId);
                
                // Clear previous selection
                this.selectedCells.clear();
                document.querySelectorAll('.cell.selected').forEach(cell => {
                    cell.classList.remove('selected');
                });
                
                // Select all cells in range
                const minCol = Math.min(startCol.charCodeAt(0), currentCol.charCodeAt(0));
                const maxCol = Math.max(startCol.charCodeAt(0), currentCol.charCodeAt(0));
                const minRow = Math.min(parseInt(startRow), parseInt(currentRow));
                const maxRow = Math.max(parseInt(startRow), parseInt(currentRow));
                
                for (let col = minCol; col <= maxCol; col++) {
                    for (let row = minRow; row <= maxRow; row++) {
                        const cellId = String.fromCharCode(col) + row;
                        const cellElement = document.querySelector(`[data-cell-id="${cellId}"]`);
                        if (cellElement) {
                            this.selectedCells.add(cellElement);
                            cellElement.classList.add('selected');
                        }
                    }
                }
                
                this.lastSelectedCell = target;
                this.updateFloatingNav();
            }
        });

        document.addEventListener('mouseup', () => {
            if (this.isDragging) {
                // Clear drag effects
                document.querySelectorAll('.cell.drag-over').forEach(cell => {
                    cell.classList.remove('drag-over');
                });
                this.isDragging = false;
            }
            this.selectionStart = null;
            this.updateFloatingNav();
        });

        // Update keyboard navigation
        document.addEventListener('keydown', (e) => {
            // If we're editing a cell, don't handle navigation keys
            if (document.activeElement.classList.contains('cell')) {
                // Only handle specific keys when editing
                switch (e.key) {
                    case 'Enter':
                        if (!e.shiftKey) {
                            const [col, row] = this.parseCellId(document.activeElement.dataset.cellId);
                            const nextCell = document.querySelector(`[data-cell-id="${col}${parseInt(row) + 1}"]`);
                            if (nextCell) {
                                document.activeElement.blur();
                                nextCell.focus();
                                nextCell.setSelectionRange(0, nextCell.value.length);
                            }
                        }
                        e.preventDefault();
                        break;
                    case 'Tab':
                        const [col, row] = this.parseCellId(document.activeElement.dataset.cellId);
                        const nextCell = document.querySelector(
                            `[data-cell-id="${String.fromCharCode(col.charCodeAt(0) + (e.shiftKey ? -1 : 1))}${row}"]`
                        );
                        if (nextCell) {
                            document.activeElement.blur();
                            nextCell.focus();
                            nextCell.setSelectionRange(0, nextCell.value.length);
                        }
                        e.preventDefault();
                        break;
                    case 'Escape':
                        document.activeElement.blur();
                        break;
                }
                return;
            }

            if (!this.lastSelectedCell) return;
            
            const [col, row] = this.parseCellId(this.lastSelectedCell.dataset.cellId);
            let nextCell = null;
            
            switch (e.key) {
                case 'ArrowUp':
                    if (parseInt(row) > 1) {
                        nextCell = document.querySelector(`[data-cell-id="${col}${parseInt(row) - 1}"]`);
                    }
                    e.preventDefault();
                    break;
                case 'ArrowDown':
                    nextCell = document.querySelector(`[data-cell-id="${col}${parseInt(row) + 1}"]`);
                    e.preventDefault();
                    break;
                case 'ArrowLeft':
                    if (col > 'A') {
                        nextCell = document.querySelector(`[data-cell-id="${String.fromCharCode(col.charCodeAt(0) - 1)}${row}"]`);
                    }
                    e.preventDefault();
                    break;
                case 'ArrowRight':
                    nextCell = document.querySelector(`[data-cell-id="${String.fromCharCode(col.charCodeAt(0) + 1)}${row}"]`);
                    e.preventDefault();
                    break;
                case 'Enter':
                    this.lastSelectedCell.focus();
                    this.lastSelectedCell.setSelectionRange(this.lastSelectedCell.value.length, this.lastSelectedCell.value.length);
                    e.preventDefault();
                    break;
                case 'F2':
                    this.lastSelectedCell.focus();
                    this.lastSelectedCell.setSelectionRange(0, this.lastSelectedCell.value.length);
                    e.preventDefault();
                    break;
            }
            
            if (nextCell) {
                if (!e.shiftKey) {
                    this.selectedCells.clear();
                    document.querySelectorAll('.cell.selected').forEach(cell => {
                        cell.classList.remove('selected');
                    });
                }
                
                this.selectedCells.add(nextCell);
                nextCell.classList.add('selected');
                this.lastSelectedCell = nextCell;
                this.updateFloatingNav();
            }
        });

        // Handle cell editing
        document.getElementById('spreadsheet').addEventListener('dblclick', (e) => {
            if (e.target.classList.contains('cell')) {
                e.target.focus();
                e.target.setSelectionRange(0, e.target.value.length);
            }
        });

        // Start editing on keypress
        document.addEventListener('keypress', (e) => {
            if (this.lastSelectedCell && !this.lastSelectedCell.matches(':focus')) {
                if (!e.ctrlKey && !e.altKey && !e.metaKey) {
                    this.lastSelectedCell.focus();
                    // Don't clear the cell value immediately
                    if (!this.lastSelectedCell.value) {
                        this.lastSelectedCell.value = e.key;
                        // Create undo command
                        this.addToHistory({
                            cellId: this.lastSelectedCell.dataset.cellId,
                            oldValue: '',
                            newValue: e.key
                        });
                    }
                }
            }
        });

        // Add input event listener for cells
        document.querySelectorAll('.cell').forEach(cell => {
            cell.addEventListener('input', (e) => {
                const cellId = e.target.dataset.cellId;
                const oldValue = this.db.getCell(this.db.getActiveSheet(), cellId).content;
                const newValue = e.target.value;
                
                // Add to undo history
                this.addToHistory({
                    cellId: cellId,
                    oldValue: oldValue,
                    newValue: newValue
                });
                
                // Update database
                this.db.updateCell(this.db.getActiveSheet(), cellId, {
                    content: newValue
                });
            });
        });

        // Add keyboard shortcuts for undo/redo
        document.addEventListener('keydown', (e) => {
            // Handle undo/redo
            if (e.ctrlKey && e.key === 'z') {
                e.preventDefault();
                if (e.shiftKey) {
                    this.redo();
                } else {
                    this.undo();
                }
                return;
            }
            
            if (e.ctrlKey && e.key === 'y') {
                e.preventDefault();
                this.redo();
                return;
            }
        });

        // Undo/Redo buttons
        document.getElementById('undoBtn').addEventListener('click', () => this.undo());
        document.getElementById('redoBtn').addEventListener('click', () => this.redo());
    }

    toggleStyle(style) {
        if (!this.activeCell) return;
        
        const cellId = this.activeCell.dataset.cellId;
        const cell = this.db.getCell(this.db.getActiveSheet(), cellId);
        const newValue = !cell.style[style];

        this.db.updateCellStyle(this.db.getActiveSheet(), cellId, {
            [style]: newValue
        });

        switch (style) {
            case 'bold':
                this.activeCell.style.fontWeight = newValue ? 'bold' : 'normal';
                break;
            case 'italic':
                this.activeCell.style.fontStyle = newValue ? 'italic' : 'normal';
                break;
            case 'underline':
                this.activeCell.style.textDecoration = newValue ? 'underline' : 'none';
                break;
        }

        document.getElementById(style).classList.toggle('active', newValue);
    }

    setAlignment(alignment) {
        if (!this.activeCell) return;
        
        const cellId = this.activeCell.dataset.cellId;
        this.db.updateCellStyle(this.db.getActiveSheet(), cellId, {
            alignment: alignment
        });
        this.activeCell.style.textAlign = alignment;

        ['alignLeft', 'alignCenter', 'alignRight'].forEach(id => {
            document.getElementById(id).classList.toggle(
                'active',
                id === 'align' + alignment.charAt(0).toUpperCase() + alignment.slice(1)
            );
        });
    }

    // Row and Column Manipulation
    addRow() {
        const sheet = this.db.data.sheets[this.db.getActiveSheet()];
        if (!sheet) return;
        
        sheet.rowCount++;
        this.initializeGrid(); // Reinitialize the grid to show the new row
        this.showAlert('Row added successfully');
    }

    deleteRow(rowIndex) {
        this.db.deleteRow(rowIndex);
        this.initializeGrid();
        this.showAlert(`Row ${rowIndex} Deleted!`);
    }

    addColumn() {
        const sheet = this.db.data.sheets[this.db.getActiveSheet()];
        if (!sheet) return;
        
        sheet.columnCount++;
        this.initializeGrid(); // Reinitialize the grid to show the new column
        this.showAlert('Column added successfully');
    }

    deleteColumn(colIndex) {
        const colLetter = String.fromCharCode(65 + colIndex);
        this.db.deleteColumn(colIndex);
        this.initializeGrid();
        this.showAlert(`Column ${colLetter} Deleted!`);
    }

    // Resize Handlers
    startColumnResize(e, colIndex) {
        e.preventDefault();
        const th = e.target.closest('th');
        const initialX = e.pageX;
        const initialWidth = th.offsetWidth;
        const table = document.getElementById('spreadsheet');
        const colCells = table.querySelectorAll(`td:nth-child(${colIndex + 2})`);
        
        const minWidth = 50; // Minimum column width
        
        const moveHandler = (moveEvent) => {
            const delta = moveEvent.pageX - initialX;
            const newWidth = Math.max(initialWidth + delta, minWidth);
            
            // Update header width
            th.style.width = `${newWidth}px`;
            th.style.minWidth = `${newWidth}px`;
            
            // Update all cells in the column
            colCells.forEach(cell => {
                cell.style.width = `${newWidth}px`;
                cell.style.minWidth = `${newWidth}px`;
                
                // Update the input element inside the cell
                const input = cell.querySelector('input');
                if (input) {
                    input.style.width = '100%';
                    input.style.minWidth = '100%';
                }
            });
        };
        
        const upHandler = () => {
            document.removeEventListener('mousemove', moveHandler);
            document.removeEventListener('mouseup', upHandler);
            document.body.style.cursor = '';
        };
        
        document.addEventListener('mousemove', moveHandler);
        document.addEventListener('mouseup', upHandler);
        document.body.style.cursor = 'col-resize';
    }

    startRowResize(e, rowIndex) {
        e.preventDefault();
        const tr = e.target.closest('tr');
        const initialY = e.pageY;
        const initialHeight = tr.offsetHeight;
        
        const minHeight = 25; // Minimum row height
        
        const moveHandler = (moveEvent) => {
            const delta = moveEvent.pageY - initialY;
            const newHeight = Math.max(initialHeight + delta, minHeight);
            tr.style.height = `${newHeight}px`;
            tr.querySelectorAll('td').forEach(cell => {
                cell.style.height = `${newHeight}px`;
                const input = cell.querySelector('input');
                if (input) {
                    input.style.height = '100%';
                }
            });
        };
        
        const upHandler = () => {
            document.removeEventListener('mousemove', moveHandler);
            document.removeEventListener('mouseup', upHandler);
            document.body.style.cursor = '';
        };
        
        document.addEventListener('mousemove', moveHandler);
        document.addEventListener('mouseup', upHandler);
        document.body.style.cursor = 'row-resize';
    }

    // Cell Drag Handler
    startCellDrag(e, cell) {
        const initialValue = cell.value;
        
        const moveHandler = (moveEvent) => {
            const target = document.elementFromPoint(moveEvent.clientX, moveEvent.clientY);
            if (target && target.classList.contains('cell')) {
                target.value = initialValue;
                this.db.updateCell(
                    this.db.getActiveSheet(),
                    target.dataset.cellId,
                    {
                        content: initialValue,
                        formula: initialValue.startsWith('=') ? initialValue : ''
                    }
                );
            }
        };
        
        const upHandler = () => {
            document.removeEventListener('mousemove', moveHandler);
            document.removeEventListener('mouseup', upHandler);
        };
        
        document.addEventListener('mousemove', moveHandler);
        document.addEventListener('mouseup', upHandler);
    }

    // Update handleCellChange method
    handleCellChange(cell) {
        const cellId = cell.dataset.cellId;
        const content = cell.value;
        
        // Remove any existing validation classes
        cell.classList.remove('invalid-number');
        
        // Validate numeric cells
        if (content.trim() !== '' && !content.startsWith('=')) {
            const isNumeric = /^-?\d*\.?\d*$/.test(content);
            const startsWithNumber = /^-?\d/.test(content);
            
            if (startsWithNumber && !isNumeric) {
                cell.classList.add('invalid-number');
            }
        }
        
        // Update cell content and formula
        this.db.updateCell(this.db.getActiveSheet(), cellId, {
            content: content,
            formula: content.startsWith('=') ? content : ''
        });
        
        // Update dependencies and evaluate formula
        if (content.startsWith('=')) {
            this.updateCellDependencies(cellId, content);
            const result = this.formula.evaluate(content, (ref) => {
                const cell = this.db.getCell(this.db.getActiveSheet(), ref);
                return cell ? cell.content : '';
            });
            cell.value = result;
            this.db.updateCell(this.db.getActiveSheet(), cellId, {
                content: result,
                formula: content
            });
        } else {
            // Clear dependencies if not a formula
            this.updateCellDependencies(cellId, '');
        }

        // Update dependent cells
        this.updateDependentCells(cellId);
    }

    // Add helper method to parse cell IDs
    parseCellId(cellId) {
        const col = cellId.match(/[A-Z]+/)[0];
        const row = cellId.match(/[0-9]+/)[0];
        return [col, row];
    }

    createDragHandle() {
        this.dragHandle = document.createElement('div');
        this.dragHandle.className = 'cell-drag-handle';
        this.dragHandle.style.display = 'none';
        document.body.appendChild(this.dragHandle);
    }

    updateDragHandle() {
        if (this.selectedCells.size === 0) {
            this.dragHandle.style.display = 'none';
            return;
        }

        // Find the last selected cell (bottom-right)
        let maxCol = 0, maxRow = 0;
        let lastCell = null;
        
        this.selectedCells.forEach(cell => {
            const [col, row] = this.parseCellId(cell.dataset.cellId);
            const colNum = col.charCodeAt(0) - 65;
            const rowNum = parseInt(row);
            
            if (colNum >= maxCol && rowNum >= maxRow) {
                maxCol = colNum;
                maxRow = rowNum;
                lastCell = cell;
            }
        });

        if (lastCell) {
            const rect = lastCell.getBoundingClientRect();
            this.dragHandle.style.display = 'block';
            this.dragHandle.style.left = (rect.right - 4) + 'px';
            this.dragHandle.style.top = (rect.bottom - 4) + 'px';
        }
    }

    // Update moveCellsToTarget method to handle drag and drop better
    moveCellsToTarget(target) {
        if (!target || !this.selectedCells.size) return;

        const targetCellId = target.dataset.cellId;
        const [targetCol, targetRow] = this.parseCellId(targetCellId);
        
        // Find the top-left selected cell
        let minCol = 'Z'.charCodeAt(0), minRow = Number.MAX_VALUE;
        let referenceCell = null;
        
        this.selectedCells.forEach(cell => {
            const [col, row] = this.parseCellId(cell.dataset.cellId);
            const colCode = col.charCodeAt(0);
            const rowNum = parseInt(row);
            
            if (colCode < minCol || (colCode === minCol && rowNum < minRow)) {
                minCol = colCode;
                minRow = rowNum;
                referenceCell = cell;
            }
        });
        
        if (!referenceCell) return;
        
        // Calculate offset from reference cell to target
        const [refCol, refRow] = this.parseCellId(referenceCell.dataset.cellId);
        const colOffset = targetCol.charCodeAt(0) - refCol.charCodeAt(0);
        const rowOffset = parseInt(targetRow) - parseInt(refRow);
        
        // Store all cell data before moving
        const cellMoves = [];
        
        // First, collect all moves to make
        this.selectedCells.forEach(sourceCell => {
            const [sourceCol, sourceRow] = this.parseCellId(sourceCell.dataset.cellId);
            const newCol = String.fromCharCode(sourceCol.charCodeAt(0) + colOffset);
            const newRow = parseInt(sourceRow) + rowOffset;
            const newCellId = newCol + newRow;
            
            const targetCell = document.querySelector(`[data-cell-id="${newCellId}"]`);
            if (targetCell) {
                cellMoves.push({
                    source: sourceCell,
                    target: targetCell,
                    sourceData: {
                        content: sourceCell.value,
                        formula: this.db.getCell(this.db.getActiveSheet(), sourceCell.dataset.cellId)?.formula || '',
                        style: {...(this.db.getCell(this.db.getActiveSheet(), sourceCell.dataset.cellId)?.style || {})}
                    },
                    targetData: {
                        content: targetCell.value,
                        formula: this.db.getCell(this.db.getActiveSheet(), targetCell.dataset.cellId)?.formula || '',
                        style: {...(this.db.getCell(this.db.getActiveSheet(), targetCell.dataset.cellId)?.style || {})}
                    }
                });
            }
        });
        
        // Then perform all moves
        cellMoves.forEach(move => {
            // Update target cell with source data
            this.db.updateCell(this.db.getActiveSheet(), move.target.dataset.cellId, {
                content: move.sourceData.content,
                formula: move.sourceData.formula
            });
            this.db.updateCellStyle(this.db.getActiveSheet(), move.target.dataset.cellId, move.sourceData.style);
            move.target.value = move.sourceData.content;
            Object.assign(move.target.style, move.sourceData.style);
            
            // Update source cell with target data
            this.db.updateCell(this.db.getActiveSheet(), move.source.dataset.cellId, {
                content: move.targetData.content,
                formula: move.targetData.formula
            });
            this.db.updateCellStyle(this.db.getActiveSheet(), move.source.dataset.cellId, move.targetData.style);
            move.source.value = move.targetData.content;
            Object.assign(move.source.style, move.targetData.style);
        });
        
        // Clear selection
        this.selectedCells.clear();
        document.querySelectorAll('.cell.selected').forEach(cell => {
            cell.classList.remove('selected');
        });
        this.updateFloatingNav();
    }

    createFloatingUI() {
        // Create floating navigation bar
        this.floatingNav = document.createElement('div');
        this.floatingNav.className = 'floating-nav';
        
        // Batch Delete button
        const deleteBtn = document.createElement('button');
        deleteBtn.innerHTML = '🗑️ Batch Delete';
        deleteBtn.onclick = () => this.deleteSelectedCells();
        
        this.floatingNav.appendChild(deleteBtn);
        document.body.appendChild(this.floatingNav);
    }

    updateFloatingNav() {
        if (this.selectedCells.size > 1) {
            const rect = this.getBoundingRectForSelectedCells();
            this.floatingNav.style.top = (rect.top - 50) + 'px';
            this.floatingNav.style.left = (rect.left + rect.width/2 - this.floatingNav.offsetWidth/2) + 'px';
            this.floatingNav.classList.add('visible');
        } else {
            this.floatingNav.classList.remove('visible');
            this.batchCopyMessage.classList.remove('visible');
        }
    }

    getBoundingRectForSelectedCells() {
        let minCol = Infinity, maxCol = -Infinity, minRow = Infinity, maxRow = -Infinity;
        
        this.selectedCells.forEach(cell => {
            const [col, row] = this.parseCellId(cell.dataset.cellId);
            const colNum = col.charCodeAt(0) - 65;
            const rowNum = parseInt(row);
            
            minCol = Math.min(minCol, colNum);
            maxCol = Math.max(maxCol, colNum);
            minRow = Math.min(minRow, rowNum);
            maxRow = Math.max(maxRow, rowNum);
        });

        const firstCell = document.querySelector(`[data-cell-id="${String.fromCharCode(minCol + 65)}${minRow}"]`);
        const lastCell = document.querySelector(`[data-cell-id="${String.fromCharCode(maxCol + 65)}${maxRow}"]`);
        
        const firstRect = firstCell.getBoundingClientRect();
        const lastRect = lastCell.getBoundingClientRect();
        
        return {
            top: firstRect.top,
            left: firstRect.left,
            bottom: lastRect.bottom,
            right: lastRect.right,
            width: lastRect.right - firstRect.left,
            height: lastRect.bottom - firstRect.top
        };
    }

    deleteSelectedCells() {
        this.selectedCells.forEach(cell => {
            const cellId = cell.dataset.cellId;
        this.db.updateCell(this.db.getActiveSheet(), cellId, {
            content: '',
                formula: ''
            });
            cell.value = '';
        });
        this.selectedCells.clear();
        this.updateFloatingNav();
    }

    // Add new method to create functions menu
    createFunctionsMenu() {
        const functionsGroup = document.createElement('div');
        functionsGroup.className = 'menu-group functions-group';

        // Create Math Functions dropdown
        const mathFunctions = document.createElement('select');
        mathFunctions.className = 'menu-btn';
        mathFunctions.innerHTML = `
            <option value="">Math Functions</option>
            <option value="SUM">SUM</option>
            <option value="AVERAGE">AVERAGE</option>
            <option value="MAX">MAX</option>
            <option value="MIN">MIN</option>
            <option value="COUNT">COUNT</option>
        `;

        // Create Data Functions dropdown
        const dataFunctions = document.createElement('select');
        dataFunctions.className = 'menu-btn';
        dataFunctions.innerHTML = `
            <option value="">Data Functions</option>
            <option value="TRIM">TRIM</option>
            <option value="UPPER">UPPER</option>
            <option value="LOWER">LOWER</option>
            <option value="REMOVE_DUPLICATES">REMOVE DUPLICATES</option>
            <option value="FIND_REPLACE">FIND & REPLACE</option>
        `;

        // Add event listeners
        mathFunctions.addEventListener('change', (e) => this.handleMathFunction(e.target.value));
        dataFunctions.addEventListener('change', (e) => this.handleDataFunction(e.target.value));

        functionsGroup.appendChild(mathFunctions);
        functionsGroup.appendChild(dataFunctions);

        // Insert after the first menu-group
        const firstMenuGroup = document.querySelector('.menu-group');
        firstMenuGroup.parentNode.insertBefore(functionsGroup, firstMenuGroup.nextSibling);
    }

    // Handle Math Functions
    handleMathFunction(functionName) {
        if (!functionName) return;

        const selectedCells = Array.from(this.selectedCells);
        if (selectedCells.length === 0) {
            alert('Please select cells to apply the function');
            return;
        }

        let result;
        const values = selectedCells.map(cell => parseFloat(cell.value)).filter(val => !isNaN(val));

        switch (functionName) {
            case 'SUM':
                result = values.reduce((sum, val) => sum + val, 0);
                break;
            case 'AVERAGE':
                result = values.length > 0 ? values.reduce((sum, val) => sum + val, 0) / values.length : 0;
                break;
            case 'MAX':
                result = values.length > 0 ? Math.max(...values) : '';
                break;
            case 'MIN':
                result = values.length > 0 ? Math.min(...values) : '';
                break;
            case 'COUNT':
                result = values.length;
                break;
        }

        // If active cell exists, put the result there
        if (this.activeCell) {
            this.activeCell.value = result;
            this.handleCellChange(this.activeCell);
        } else if (selectedCells.length > 0) {
            // Otherwise, put it in the last selected cell
            const lastCell = selectedCells[selectedCells.length - 1];
            lastCell.value = result;
            this.handleCellChange(lastCell);
        }

        // Reset the dropdown
        document.querySelector('.functions-group select').value = '';
    }

    // Handle Data Functions
    handleDataFunction(functionName) {
        if (!functionName) return;

        const selectedCells = Array.from(this.selectedCells);
        if (selectedCells.length === 0) {
            alert('Please select cells to apply the function');
            return;
        }

        switch (functionName) {
            case 'TRIM':
                selectedCells.forEach(cell => {
                    cell.value = cell.value.trim();
                    this.handleCellChange(cell);
                });
                break;

            case 'UPPER':
                selectedCells.forEach(cell => {
                    cell.value = cell.value.toUpperCase();
                    this.handleCellChange(cell);
                });
                break;

            case 'LOWER':
                selectedCells.forEach(cell => {
                    cell.value = cell.value.toLowerCase();
                    this.handleCellChange(cell);
                });
                break;

            case 'FIND_REPLACE':
                this.showFindReplaceDialog();
                break;

            case 'REMOVE_DUPLICATES':
                this.removeDuplicates(selectedCells);
                break;
        }

        // Reset the dropdown
        const dataFunctionsDropdown = document.querySelector('.functions-group select:nth-child(2)');
        if (dataFunctionsDropdown) {
            dataFunctionsDropdown.value = '';
        }
    }

    // Find and Replace Dialog
    showFindReplaceDialog() {
        const dialog = document.createElement('div');
        dialog.className = 'find-replace-dialog';
        dialog.innerHTML = `
            <div class="dialog-content">
                <h3>Find and Replace</h3>
                <div class="input-group">
                    <label>Find:</label>
                    <input type="text" id="findText">
                </div>
                <div class="input-group">
                    <label>Replace with:</label>
                    <input type="text" id="replaceText">
                </div>
                <div class="dialog-buttons">
                    <button id="replaceBtn" class="test-btn primary">Replace All</button>
                    <button id="cancelBtn" class="test-btn secondary">Cancel</button>
                </div>
            </div>
        `;

        document.body.appendChild(dialog);

        // Add event listeners
        document.getElementById('replaceBtn').onclick = () => {
            const findText = document.getElementById('findText').value;
            const replaceText = document.getElementById('replaceText').value;
            
            if (!findText) {
                alert('Please enter text to find');
                return;
            }
            
            let replacements = 0;
            Array.from(this.selectedCells).forEach(cell => {
                if (cell.value.includes(findText)) {
                    cell.value = cell.value.replaceAll(findText, replaceText);
                    this.handleCellChange(cell);
                    replacements++;
                }
            });
            
            alert(`Replaced ${replacements} occurrence(s)`);
            document.body.removeChild(dialog);
        };

        document.getElementById('cancelBtn').onclick = () => {
            document.body.removeChild(dialog);
        };
    }

    // Remove Duplicates
    removeDuplicates(cells) {
        if (cells.length < 2) {
            alert('Please select at least two cells to remove duplicates');
            return;
        }

        // Get the range of selected cells
        const cellIds = cells.map(cell => cell.dataset.cellId).sort();
        const firstCell = cellIds[0];
        const lastCell = cellIds[cellIds.length - 1];
        const range = `${firstCell}:${lastCell}`;

        // Call the database's removeDuplicates function
        this.db.removeDuplicates(range).then(count => {
            if (count > 0) {
                // Refresh the cells display after duplicates are removed
                cells.forEach(cell => {
                    const cellId = cell.dataset.cellId;
                    const cellData = this.db.getCell(this.db.getActiveSheet(), cellId);
                    if (cellData) {
                        cell.value = cellData.content || '';
                        this.handleCellChange(cell);
                    }
                });
            }

            // Reset the dropdown
            const dataFunctionsDropdown = document.querySelector('.functions-group select:nth-child(2)');
            if (dataFunctionsDropdown) {
                dataFunctionsDropdown.value = '';
            }
        }).catch(error => {
            console.error('Error removing duplicates:', error);
            alert('An error occurred while removing duplicates');
        });
    }

    // Add new method for showing alerts
    showAlert(message) {
        const alert = document.createElement('div');
        alert.className = 'quick-alert';
        alert.textContent = message;
        document.body.appendChild(alert);
        
        // Remove after 500ms
        setTimeout(() => {
            alert.classList.add('fade-out');
            setTimeout(() => document.body.removeChild(alert), 200);
        }, 500);
    }

    // Add speech synthesis for buttons
    setupSpeechTooltips() {
        const speak = (text) => {
            if ('speechSynthesis' in window) {
                const utterance = new SpeechSynthesisUtterance(text);
                speechSynthesis.speak(utterance);
            }
        };

        // Add row button tooltip
        const addRowBtn = document.querySelector('.add-row-btn');
        if (addRowBtn) {
            addRowBtn.addEventListener('mouseenter', () => speak('Add a new row'));
        }

        // Delete row buttons tooltips
        document.querySelectorAll('.delete-row-btn').forEach(btn => {
            btn.addEventListener('mouseenter', () => speak('Delete this row'));
        });

        // Delete column buttons tooltips
        document.querySelectorAll('.delete-col-btn').forEach(btn => {
            btn.addEventListener('mouseenter', () => speak('Delete this column'));
        });
    }

    createClock() {
        // Create clock container
        const clockContainer = document.createElement('div');
        clockContainer.className = 'clock-container';

        // Create clocks wrapper
        const clocksWrapper = document.createElement('div');
        clocksWrapper.className = 'clocks-wrapper';

        // Create different clock styles
        const digitalClock = document.createElement('div');
        digitalClock.className = 'digital-clock active';  // Digital clock is active by default

        const analogClock = document.createElement('div');
        analogClock.className = 'analog-clock';
        analogClock.innerHTML = `
            <div class="hour-hand clock-hand"></div>
            <div class="minute-hand clock-hand"></div>
            <div class="second-hand clock-hand"></div>
            <div class="clock-center"></div>
        `;

        const semiDigitalClock = document.createElement('div');
        semiDigitalClock.className = 'semi-digital-clock';

        const minimalClock = document.createElement('div');
        minimalClock.className = 'minimal-clock';

        // Create date display
        const dateDisplay = document.createElement('div');
        dateDisplay.className = 'date-display';

        // Create clock styles menu
        const clockStylesMenu = document.createElement('div');
        clockStylesMenu.className = 'clock-styles-menu';
        clockStylesMenu.innerHTML = `
            <div class="clock-style-option active" data-style="digital">
                <span>🟦 Digital Clock</span>
            </div>
            <div class="clock-style-option" data-style="analog">
                <span>⭕ Analog Clock</span>
                <button class="color-btn" title="Change clock color">🎨</button>
            </div>
            <div class="clock-style-option" data-style="semi-digital">
                <span>🔷 Semi-Digital Clock</span>
            </div>
            <div class="clock-style-option" data-style="minimal">
                <span>➖ Minimal Clock</span>
            </div>
        `;

        // Create color picker for analog clock
        const colorPicker = document.createElement('div');
        colorPicker.className = 'clock-color-picker';
        colorPicker.innerHTML = `
            <input type="color" value="#1a73e8">
        `;

        // Add all elements to container
        clocksWrapper.appendChild(digitalClock);
        clocksWrapper.appendChild(analogClock);
        clocksWrapper.appendChild(semiDigitalClock);
        clocksWrapper.appendChild(minimalClock);
        
        clockContainer.appendChild(clocksWrapper);
        clockContainer.appendChild(dateDisplay);
        clockContainer.appendChild(clockStylesMenu);
        clockContainer.appendChild(colorPicker);

        // Add click handler for date display to show menu
        dateDisplay.addEventListener('click', () => {
            clockStylesMenu.classList.toggle('visible');
            colorPicker.classList.remove('visible');
        });

        // Add click handlers for clock style options
        clockStylesMenu.addEventListener('click', (e) => {
            const option = e.target.closest('.clock-style-option');
            const colorBtn = e.target.closest('.color-btn');
            
            if (colorBtn) {
                e.stopPropagation();
                colorPicker.classList.toggle('visible');
                clockStylesMenu.classList.remove('visible');
            } else if (option) {
                const style = option.dataset.style;
                
                // Remove active class from all options and clocks
                clockStylesMenu.querySelectorAll('.clock-style-option').forEach(opt => opt.classList.remove('active'));
                clocksWrapper.querySelectorAll('.digital-clock, .analog-clock, .semi-digital-clock, .minimal-clock').forEach(clock => clock.classList.remove('active'));
                
                // Add active class to selected option and clock
                option.classList.add('active');
                clocksWrapper.querySelector(`.${style}-clock`).classList.add('active');
                
                // Hide the menu
                clockStylesMenu.classList.remove('visible');
            }
        });

        // Add color picker handler
        const colorInput = colorPicker.querySelector('input[type="color"]');
        colorInput.addEventListener('input', (e) => {
            const color = e.target.value;
            document.documentElement.style.setProperty('--accent-color', color);
            
            // Convert hex to RGB for shadow
            const rgb = this.hexToRgb(color);
            document.documentElement.style.setProperty('--accent-color-rgb', `${rgb.r}, ${rgb.g}, ${rgb.b}`);
        });

        // Close menus when clicking outside
        document.addEventListener('click', (e) => {
            if (!clockContainer.contains(e.target)) {
                clockStylesMenu.classList.remove('visible');
                colorPicker.classList.remove('visible');
            }
        });

        // Add clock to navigation bar
        const menuBar = document.querySelector('.menu-bar');
        menuBar.appendChild(clockContainer);

        // Start clock update
        this.updateClock();
        setInterval(() => this.updateClock(), 1000);
    }

    // Helper method to convert hex to RGB
    hexToRgb(hex) {
        const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
        return result ? {
            r: parseInt(result[1], 16),
            g: parseInt(result[2], 16),
            b: parseInt(result[3], 16)
        } : null;
    }

    updateClock() {
        const now = new Date();
        const hours = now.getHours().toString().padStart(2, '0');
        const minutes = now.getMinutes().toString().padStart(2, '0');
        const seconds = now.getSeconds().toString().padStart(2, '0');
        
        // Update date display
        const dateDisplay = document.querySelector('.date-display');
        const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
        const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
        dateDisplay.textContent = `${days[now.getDay()]}, ${months[now.getMonth()]} ${now.getDate()}, ${now.getFullYear()}`;

        // Update digital clock
        const digitalClock = document.querySelector('.digital-clock');
        digitalClock.textContent = `${hours}:${minutes}:${seconds}`;

        // Update analog clock
        const analogClock = document.querySelector('.analog-clock');
        const hourDeg = (hours % 12 + minutes / 60) * 30;
        const minuteDeg = (minutes + seconds / 60) * 6;
        const secondDeg = seconds * 6;

        const hourHand = analogClock.querySelector('.hour-hand');
        const minuteHand = analogClock.querySelector('.minute-hand');
        const secondHand = analogClock.querySelector('.second-hand');

        hourHand.style.transform = `rotate(${hourDeg}deg)`;
        minuteHand.style.transform = `rotate(${minuteDeg}deg)`;
        secondHand.style.transform = `rotate(${secondDeg}deg)`;

        // Update semi-digital clock
        const semiDigitalClock = document.querySelector('.semi-digital-clock');
        const period = hours >= 12 ? 'PM' : 'AM';
        const hour12 = (hours % 12) || 12;
        semiDigitalClock.textContent = `${hour12}:${minutes} ${period}`;

        // Update minimal clock
        const minimalClock = document.querySelector('.minimal-clock');
        minimalClock.textContent = `${hour12}:${minutes}`;
    }

    createCalculator() {
        // Create calculator button in menu bar
        const menuBar = document.querySelector('.menu-bar');
        const calcButton = document.createElement('button');
        calcButton.className = 'menu-btn';
        calcButton.innerHTML = '🧮 Calculator';
        calcButton.onclick = () => this.showCalculator();
        menuBar.querySelector('.menu-group').appendChild(calcButton);

        // Create calculator overlay
        const overlay = document.createElement('div');
        overlay.className = 'calculator-overlay';
        overlay.onclick = (e) => {
            if (e.target === overlay) {
                this.hideCalculator();
            }
        };

        // Create calculator container
        const calculator = document.createElement('div');
        calculator.className = 'calculator-container';
        calculator.innerHTML = `
            <div class="calculator-header">
                <button class="calculator-close">×</button>
            </div>
            <div class="calculator-display">0</div>
            <div class="calculator-buttons">
                <button class="calc-btn function">AC</button>
                <button class="calc-btn function">±</button>
                <button class="calc-btn function">%</button>
                <button class="calc-btn operator">÷</button>
                <button class="calc-btn">7</button>
                <button class="calc-btn">8</button>
                <button class="calc-btn">9</button>
                <button class="calc-btn operator">×</button>
                <button class="calc-btn">4</button>
                <button class="calc-btn">5</button>
                <button class="calc-btn">6</button>
                <button class="calc-btn operator">−</button>
                <button class="calc-btn">1</button>
                <button class="calc-btn">2</button>
                <button class="calc-btn">3</button>
                <button class="calc-btn operator">+</button>
                <button class="calc-btn zero">0</button>
                <button class="calc-btn">.</button>
                <button class="calc-btn equals">=</button>
            </div>
        `;

        // Add event listeners for calculator buttons
        calculator.querySelectorAll('.calc-btn').forEach(button => {
            button.addEventListener('click', () => this.handleCalculatorInput(button.textContent));
        });

        // Add close button handler
        calculator.querySelector('.calculator-close').addEventListener('click', () => this.hideCalculator());

        // Add keyboard support
        document.addEventListener('keydown', (e) => {
            if (!calculator.classList.contains('visible')) return;
            
            const key = e.key;
            const validKeys = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '.', '+', '-', '*', '/', '=', 'Enter', 'Escape', 'Backspace'];
            
            if (validKeys.includes(key)) {
                e.preventDefault();
                let input = key;
                if (key === 'Enter') input = '=';
                if (key === 'Escape') this.hideCalculator();
                if (key === '*') input = '×';
                if (key === '/') input = '÷';
                if (key === '-') input = '−';
                if (key === 'Backspace') {
                    this.calculatorValue = this.calculatorValue.slice(0, -1) || '0';
                    this.updateCalculatorDisplay();
                    return;
                }
                this.handleCalculatorInput(input);
            }
        });

        document.body.appendChild(overlay);
        document.body.appendChild(calculator);
    }

    showCalculator() {
        document.querySelector('.calculator-overlay').classList.add('visible');
        document.querySelector('.calculator-container').classList.add('visible');
    }

    hideCalculator() {
        document.querySelector('.calculator-overlay').classList.remove('visible');
        document.querySelector('.calculator-container').classList.remove('visible');
        // Reset calculator
        this.calculatorValue = '0';
        this.lastOperator = null;
        this.lastNumber = null;
        this.newNumber = true;
        this.updateCalculatorDisplay();
    }

    handleCalculatorInput(value) {
        switch (value) {
            case 'AC':
                this.calculatorValue = '0';
                this.lastOperator = null;
                this.lastNumber = null;
                this.newNumber = true;
                break;
            case '±':
                this.calculatorValue = (parseFloat(this.calculatorValue) * -1).toString();
                break;
            case '%':
                this.calculatorValue = (parseFloat(this.calculatorValue) / 100).toString();
                break;
            case '+':
            case '−':
            case '×':
            case '÷':
                this.handleOperator(value);
                break;
            case '=':
                this.calculateResult();
                break;
            case '.':
                if (!this.calculatorValue.includes('.')) {
                    this.calculatorValue += '.';
                }
                break;
            default:
                // Handle numbers
                if (this.newNumber) {
                    this.calculatorValue = value;
                    this.newNumber = false;
                } else {
                    this.calculatorValue = this.calculatorValue === '0' ? 
                        value : this.calculatorValue + value;
                }
                break;
        }
        this.updateCalculatorDisplay();
    }

    handleOperator(operator) {
        if (this.lastOperator && !this.newNumber) {
            this.calculateResult();
        }
        this.lastNumber = parseFloat(this.calculatorValue);
        this.lastOperator = operator;
        this.newNumber = true;
    }

    calculateResult() {
        if (this.lastOperator === null || this.newNumber) return;

        const current = parseFloat(this.calculatorValue);
        let result;

        switch (this.lastOperator) {
            case '+':
                result = this.lastNumber + current;
                break;
            case '−':
                result = this.lastNumber - current;
                break;
            case '×':
                result = this.lastNumber * current;
                break;
            case '÷':
                result = this.lastNumber / current;
                break;
        }

        this.calculatorValue = result.toString();
        this.lastOperator = null;
        this.newNumber = true;
    }

    updateCalculatorDisplay() {
        const display = document.querySelector('.calculator-display');
        let displayValue = this.calculatorValue;

        // Format large numbers
        if (displayValue.length > 10) {
            displayValue = parseFloat(displayValue).toExponential(6);
        }

        display.textContent = displayValue;
    }

    // Add new methods for cell dependency handling
    updateCellDependencies(cellId, formula) {
        // Clear old dependencies for this cell
        if (this.cellDependencies.has(cellId)) {
            const oldDeps = this.cellDependencies.get(cellId);
            oldDeps.forEach(dep => {
                const dependents = this.dependentCells.get(dep);
                if (dependents) {
                    dependents.delete(cellId);
                    if (dependents.size === 0) {
                        this.dependentCells.delete(dep);
                    }
                }
            });
        }

        // If no formula, clear dependencies and return
        if (!formula || !formula.startsWith('=')) {
            this.cellDependencies.delete(cellId);
            return;
        }

        // Extract cell references from formula
        const refs = this.extractCellReferences(formula);
        if (refs.length > 0) {
            this.cellDependencies.set(cellId, new Set(refs));
            
            // Update dependent cells
            refs.forEach(ref => {
                if (!this.dependentCells.has(ref)) {
                    this.dependentCells.set(ref, new Set());
                }
                this.dependentCells.get(ref).add(cellId);
            });
        }
    }

    extractCellReferences(formula) {
        const cellPattern = /[A-Z]+[0-9]+/g;
        return Array.from(new Set(formula.match(cellPattern) || []));
    }

    updateDependentCells(cellId) {
        if (!this.dependentCells.has(cellId)) return;

        const visited = new Set();
        const toUpdate = [...this.dependentCells.get(cellId)];

        while (toUpdate.length > 0) {
            const currentCell = toUpdate.shift();
            if (visited.has(currentCell)) continue;
            visited.add(currentCell);

            const cell = document.querySelector(`[data-cell-id="${currentCell}"]`);
            if (cell) {
                const formula = this.db.getCell(this.db.getActiveSheet(), currentCell)?.formula;
                if (formula) {
                    const result = this.formula.evaluate(formula, (ref) => {
                        const refCell = this.db.getCell(this.db.getActiveSheet(), ref);
                        return refCell ? refCell.content : '';
                    });

                    this.db.updateCell(this.db.getActiveSheet(), currentCell, {
                        content: result,
                        formula: formula
                    });
                    cell.value = result;

                    // Add dependent cells to update queue
                    if (this.dependentCells.has(currentCell)) {
                        toUpdate.push(...this.dependentCells.get(currentCell));
                    }
                }
            }
        }
    }

    handleHeaderClick(e) {
        const header = e.target;
        const isColumn = header.dataset.col !== undefined;
        const isRow = header.dataset.row !== undefined;
        
        if (!isColumn && !isRow) return;
        
        if (e.shiftKey && this.selectedHeaders.size > 0) {
            // Range selection
            const lastHeader = Array.from(this.selectedHeaders).pop();
            const start = isColumn ? 
                lastHeader.dataset.col.charCodeAt(0) - 65 :
                parseInt(lastHeader.dataset.row);
            const end = isColumn ?
                header.dataset.col.charCodeAt(0) - 65 :
                parseInt(header.dataset.row);
            
            this.selectHeaderRange(Math.min(start, end), Math.max(start, end), isColumn);
        } else if (e.ctrlKey || e.metaKey) {
            // Add to selection
            this.toggleHeaderSelection(header);
        } else {
            // New selection
            this.clearHeaderSelection();
            this.toggleHeaderSelection(header);
        }
        
        this.selectCellsByHeaders();
    }

    toggleHeaderSelection(header) {
        if (this.selectedHeaders.has(header)) {
            this.selectedHeaders.delete(header);
            header.classList.remove('selected');
        } else {
            this.selectedHeaders.add(header);
            header.classList.add('selected');
        }
    }

    selectHeaderRange(start, end, isColumn) {
        this.clearHeaderSelection();
        
        for (let i = start; i <= end; i++) {
            const selector = isColumn ?
                `th[data-col="${String.fromCharCode(i + 65)}"]` :
                `th[data-row="${i}"]`;
            const header = document.querySelector(selector);
            if (header) {
                this.selectedHeaders.add(header);
                header.classList.add('selected');
            }
        }
    }

    clearHeaderSelection() {
        this.selectedHeaders.forEach(header => {
            header.classList.remove('selected');
        });
        this.selectedHeaders.clear();
    }

    selectCellsByHeaders() {
        this.selectedCells.clear();
        document.querySelectorAll('.cell.selected').forEach(cell => {
            cell.classList.remove('selected');
        });
        
        this.selectedHeaders.forEach(header => {
            const isColumn = header.dataset.col !== undefined;
            const value = isColumn ? header.dataset.col : header.dataset.row;
            
            document.querySelectorAll('.cell').forEach(cell => {
                const [col, row] = this.parseCellId(cell.dataset.cellId);
                if ((isColumn && col === value) || (!isColumn && row === value)) {
                    this.selectedCells.add(cell);
                    cell.classList.add('selected');
                }
            });
        });
        
        this.updateFloatingNav();
    }

    handleCellMouseDown(e) {
        // ... existing handleCellMouseDown code ...
        
        // Check if clicking fill handle
        const rect = e.target.getBoundingClientRect();
        const isClickingFillHandle = 
            e.clientX >= rect.right - 8 &&
            e.clientX <= rect.right + 4 &&
            e.clientY >= rect.bottom - 8 &&
            e.clientY <= rect.bottom + 4;
            
        if (isClickingFillHandle) {
            this.fillHandleActive = true;
            this.fillStartCell = e.target;
            e.preventDefault();
            return;
        }
    }

    handleCellMouseMove(e) {
        if (this.fillHandleActive && this.fillStartCell) {
            this.updateFillPreview(e);
            return;
        }
        
        // ... existing handleCellMouseMove code ...
    }

    handleCellMouseUp(e) {
        if (this.fillHandleActive) {
            this.fillHandleActive = false;
            this.applyFill(e);
            this.fillPreview.style.display = 'none';
            return;
        }
        
        // ... existing handleCellMouseUp code ...
    }

    updateFillPreview(e) {
        const startRect = this.fillStartCell.getBoundingClientRect();
        const cells = document.elementsFromPoint(e.clientX, e.clientY);
        const targetCell = cells.find(el => el.classList.contains('cell'));
        
        if (targetCell) {
            const targetRect = targetCell.getBoundingClientRect();
            const previewRect = {
                left: Math.min(startRect.left, targetRect.left),
                top: Math.min(startRect.top, targetRect.top),
                right: Math.max(startRect.right, targetRect.right),
                bottom: Math.max(startRect.bottom, targetRect.bottom)
            };
            
            Object.assign(this.fillPreview.style, {
                display: 'block',
                left: previewRect.left + 'px',
                top: previewRect.top + 'px',
                width: (previewRect.right - previewRect.left) + 'px',
                height: (previewRect.bottom - previewRect.top) + 'px'
            });
        }
    }

    applyFill(e) {
        const cells = document.elementsFromPoint(e.clientX, e.clientY);
        const targetCell = cells.find(el => el.classList.contains('cell'));
        
        if (targetCell && targetCell !== this.fillStartCell) {
            const [startCol, startRow] = this.parseCellId(this.fillStartCell.dataset.cellId);
            const [endCol, endRow] = this.parseCellId(targetCell.dataset.cellId);
            
            const startColNum = startCol.charCodeAt(0) - 65;
            const endColNum = endCol.charCodeAt(0) - 65;
            const startRowNum = parseInt(startRow);
            const endRowNum = parseInt(endRow);
            
            const sourceValue = this.fillStartCell.value;
            
            // Fill direction
            const isVertical = startColNum === endColNum;
            
            if (isVertical) {
                const direction = endRowNum > startRowNum ? 1 : -1;
                for (let row = startRowNum; row !== endRowNum + direction; row += direction) {
                    const cellId = `${startCol}${row}`;
                    const cell = document.querySelector(`[data-cell-id="${cellId}"]`);
                    if (cell) {
                        this.db.updateCell(this.db.getActiveSheet(), cellId, {
                            content: sourceValue,
                            formula: ''
                        });
                        cell.value = sourceValue;
                    }
                }
            } else {
                const direction = endColNum > startColNum ? 1 : -1;
                for (let col = startColNum; col !== endColNum + direction; col += direction) {
                    const cellId = `${String.fromCharCode(col + 65)}${startRow}`;
                    const cell = document.querySelector(`[data-cell-id="${cellId}"]`);
                    if (cell) {
                        this.db.updateCell(this.db.getActiveSheet(), cellId, {
                            content: sourceValue,
                            formula: ''
                        });
                        cell.value = sourceValue;
                    }
                }
            }
        }
    }

    createTestingInterface() {
        // Create testing button in menu bar
        const menuBar = document.querySelector('.menu-bar');
        const testButton = document.createElement('button');
        testButton.className = 'menu-btn';
        testButton.innerHTML = '🧪 Test Functions';
        testButton.onclick = () => this.showTestingDialog();
        menuBar.querySelector('.menu-group').appendChild(testButton);

        // Create overlay
        const overlay = document.createElement('div');
        overlay.className = 'testing-overlay';
        document.body.appendChild(overlay);

        // Create testing dialog
        const dialog = document.createElement('div');
        dialog.className = 'testing-dialog';
        
        // Create dialog content
        dialog.innerHTML = `
            <div class="testing-header">
                <h2>Test Functions</h2>
                <button class="testing-close">&times;</button>
            </div>
            <div class="testing-content">
                <select class="function-select">
                    <optgroup label="Mathematical Functions">
                        <option value="SUM">SUM</option>
                        <option value="AVERAGE">AVERAGE</option>
                        <option value="MAX">MAX</option>
                        <option value="MIN">MIN</option>
                        <option value="COUNT">COUNT</option>
                    </optgroup>
                    <optgroup label="Data Quality Functions">
                        <option value="TRIM">TRIM</option>
                        <option value="UPPER">UPPER</option>
                        <option value="LOWER">LOWER</option>
                        <option value="REMOVE_DUPLICATES">REMOVE DUPLICATES</option>
                    </optgroup>
                </select>
                <div class="test-input-group">
                    <label>Test Input (Cell Range or Text)</label>
                    <input type="text" class="test-input" placeholder="e.g., A1:B5 or 'Sample Text'">
                </div>
                <div class="test-result">
                    <pre></pre>
                </div>
                <div class="test-buttons">
                    <button class="test-btn secondary">Clear</button>
                    <button class="test-btn primary">Test Function</button>
                </div>
            </div>
        `;

        // Add event listeners
        dialog.querySelector('.testing-close').onclick = () => this.hideTestingDialog();
        dialog.querySelector('.test-btn.secondary').onclick = () => this.clearTestInput();
        dialog.querySelector('.test-btn.primary').onclick = () => this.runFunctionTest();
        dialog.querySelector('.function-select').onchange = () => this.updateTestInputPlaceholder();

        // Add to document
        document.body.appendChild(dialog);

        // Store references
        this.testingOverlay = overlay;
        this.testingDialog = dialog;
    }

    showTestingDialog() {
        if (this.testingOverlay && this.testingDialog) {
            this.testingOverlay.style.display = 'block';
            this.testingDialog.style.display = 'block';
        }
    }

    hideTestingDialog() {
        if (this.testingOverlay && this.testingDialog) {
            this.testingOverlay.style.display = 'none';
            this.testingDialog.style.display = 'none';
            this.clearTestInput();
        }
    }

    clearTestInput() {
        const input = this.testingDialog.querySelector('.test-input');
        input.value = '';
        const resultDiv = this.testingDialog.querySelector('.test-result');
        resultDiv.style.display = 'none';
    }

    updateTestInputPlaceholder() {
        const select = this.testingDialog.querySelector('.function-select');
        const input = this.testingDialog.querySelector('.test-input');
        const function_name = select.value;
        
        switch(function_name) {
            case 'SUM':
            case 'AVERAGE':
            case 'MAX':
            case 'MIN':
            case 'COUNT':
            case 'REMOVE_DUPLICATES':
                input.placeholder = 'Enter cell range (e.g., A1:B5)';
                break;
            case 'TRIM':
            case 'UPPER':
            case 'LOWER':
                input.placeholder = 'Enter text to transform';
                break;
        }
    }

    testRemoveDuplicates(range) {
        try {
            const [start, end] = range.split(':');
            if (!start || !end) throw new Error('Invalid range format. Use format like A1:B5');
            
            const [startCol, startRow] = this.parseCellId(start);
            const [endCol, endRow] = this.parseCellId(end);
            
            // Get values from range
            const values = new Set();
            const duplicates = new Set();
            
            for (let row = startRow; row <= endRow; row++) {
                for (let colIndex = startCol.charCodeAt(0) - 65; colIndex <= endCol.charCodeAt(0) - 65; colIndex++) {
                    const col = String.fromCharCode(65 + colIndex);
                    const cellId = `${col}${row}`;
                    const cell = document.querySelector(`input[data-cell-id="${cellId}"]`);
                    if (cell && cell.value) {
                        if (values.has(cell.value)) {
                            duplicates.add(cell.value);
                        } else {
                            values.add(cell.value);
                        }
                    }
                }
            }
            
            if (duplicates.size === 0) {
                return 'No duplicates found in the range';
            }
            
            return `Found ${duplicates.size} duplicate value(s):\n${Array.from(duplicates).join(', ')}`;
        } catch (error) {
            return `Error: ${error.message}`;
        }
    }

    runFunctionTest() {
        const select = this.testingDialog.querySelector('.function-select');
        const input = this.testingDialog.querySelector('.test-input');
        const resultDiv = this.testingDialog.querySelector('.test-result');
        const resultPre = resultDiv.querySelector('pre');
        
        const function_name = select.value;
        const test_input = input.value.trim();
        
        if (!test_input) {
            resultPre.textContent = 'Error: Please enter test input';
            resultDiv.style.display = 'block';
            return;
        }

        let result;
        try {
            switch(function_name) {
                case 'SUM':
                case 'AVERAGE':
                case 'MAX':
                case 'MIN':
                case 'COUNT':
                    result = this.testMathFunction(function_name, test_input);
                    break;
                case 'TRIM':
                    result = test_input.trim();
                    break;
                case 'UPPER':
                    result = test_input.toUpperCase();
                    break;
                case 'LOWER':
                    result = test_input.toLowerCase();
                    break;
                case 'REMOVE_DUPLICATES':
                    result = this.testRemoveDuplicates(test_input);
                    break;
                default:
                    throw new Error('Unknown function');
            }
            
            resultPre.textContent = `Function: ${function_name}\nInput: ${test_input}\nResult: ${result}`;
            resultDiv.style.display = 'block';
        } catch (error) {
            resultPre.textContent = `Error: ${error.message}`;
            resultDiv.style.display = 'block';
        }
    }

    testMathFunction(func, range) {
        // Parse range (e.g., A1:B5)
        const [start, end] = range.split(':');
        if (!start || !end) throw new Error('Invalid range format. Use format like A1:B5');
        
        const [startCol, startRow] = this.parseCellId(start);
        const [endCol, endRow] = this.parseCellId(end);
        
        // Get values from range
        const values = [];
        for (let row = startRow; row <= endRow; row++) {
            for (let colIndex = startCol.charCodeAt(0) - 65; colIndex <= endCol.charCodeAt(0) - 65; colIndex++) {
                const col = String.fromCharCode(65 + colIndex);
                const cellId = `${col}${row}`;
                const cell = document.querySelector(`input[data-cell-id="${cellId}"]`);
                if (cell && cell.value) {
                    const val = parseFloat(cell.value);
                    if (!isNaN(val)) values.push(val);
                }
            }
        }
        
        // If no valid values found, throw error
        if (values.length === 0) {
            throw new Error('No valid numeric values found in the specified range');
        }
        
        // Calculate result
        switch(func) {
            case 'SUM':
                return values.reduce((a, b) => a + b, 0);
            case 'AVERAGE':
                return values.reduce((a, b) => a + b, 0) / values.length;
            case 'MAX':
                return Math.max(...values);
            case 'MIN':
                return Math.min(...values);
            case 'COUNT':
                return values.length;
            default:
                throw new Error('Unknown function');
        }
    }

    // Add new method to load file from path
    async loadFileFromPath(filePath) {
        try {
            // Extract file name from path
            const fileName = filePath.split('/').pop();
            const fileType = filePath.toLowerCase().endsWith('.csv') ? 'csv' : 'json';
            
            // Get the file content from localStorage
            const content = localStorage.getItem(`spreadsheet_file_${fileName}`);
            if (!content) {
                throw new Error('File content not found');
            }

            // For CSV files
            if (fileType === 'csv') {
                // Create a new sheet with the file name (without extension)
                const sheetName = fileName.replace('.csv', '');
                this.db.addSheet(sheetName);
                
                // Parse CSV and populate cells
                const rows = content.split('\n');
                rows.forEach((row, rowIndex) => {
                    if (row.trim()) {
                        const cells = row.split(',');
                        cells.forEach((cellContent, colIndex) => {
                            const cellId = this.db.indexToColumn(colIndex) + (rowIndex + 1);
                            // Remove quotes if present
                            const cleanContent = cellContent.replace(/^"|"$/g, '').replace(/""/g, '"');
                            this.db.updateCell(sheetName, cellId, {
                                content: cleanContent,
                                formula: ''
                            });
                        });
                    }
                });
            } 
            // For JSON files
            else {
                const data = JSON.parse(content);
                // Validate file structure
                if (!data || typeof data !== 'object') {
                    throw new Error('Invalid file format: File must contain valid JSON data');
                }
                if (!data.sheets || typeof data.sheets !== 'object') {
                    throw new Error('Invalid file format: Missing or invalid sheets data');
                }
                if (!data.activeSheet || !data.sheets[data.activeSheet]) {
                    throw new Error('Invalid file format: Missing or invalid active sheet');
                }

                // Load the data into the database
                this.db.data = data;
                this.db.saveToLocalStorage();
            }

            // Update recent files access time
            const recentFiles = JSON.parse(localStorage.getItem('spreadsheet_recent_files') || '[]');
            const updatedFiles = recentFiles.map(f => {
                if (f.path === filePath) {
                    return { ...f, lastOpened: new Date().toLocaleString() };
                }
                return f;
            });
            localStorage.setItem('spreadsheet_recent_files', JSON.stringify(updatedFiles));

            // Refresh the grid and sheet tabs
            this.initializeGrid();
            this.initializeSheetTabs();

            // Update UI elements
            const cells = document.querySelectorAll('.cell');
            cells.forEach(cell => {
                const cellId = cell.dataset.cellId;
                const cellData = this.db.getCell(this.db.getActiveSheet(), cellId);
                if (cellData && cellData.content) {
                    cell.value = cellData.content;
                    // Apply styles if present
                    if (cellData.style) {
                        cell.style.fontWeight = cellData.style.bold ? 'bold' : 'normal';
                        cell.style.fontStyle = cellData.style.italic ? 'italic' : 'normal';
                        cell.style.textDecoration = cellData.style.underline ? 'underline' : 'none';
                        cell.style.textAlign = cellData.style.alignment || 'left';
                        cell.style.fontFamily = cellData.style.fontFamily || 'arial';
                        cell.style.fontSize = (cellData.style.fontSize || 14) + 'px';
                        cell.style.color = cellData.style.color || '#000000';
                        cell.style.backgroundColor = cellData.style.backgroundColor || '#ffffff';
                    }
                }
            });
        } catch (error) {
            alert('Error loading file: ' + error.message);
        }
    }

    // Helper function to convert column index to label (0=A, 1=B, ..., 26=AA, 27=AB, etc.)
    getColumnLabel(index) {
        let label = '';
        while (index >= 0) {
            label = String.fromCharCode(65 + (index % 26)) + label;
            index = Math.floor(index / 26) - 1;
        }
        return label;
    }

    // Add command to undo stack
    addToHistory(command) {
        this.undoStack.push(command);
        if (this.undoStack.length > this.maxStackSize) {
            this.undoStack.shift();
        }
        this.redoStack = []; // Clear redo stack when new command is added
    }

    // Undo last action
    undo() {
        if (this.undoStack.length === 0) return;
        
        const command = this.undoStack.pop();
        this.redoStack.push({
            cellId: command.cellId,
            oldValue: this.db.getCell(this.db.getActiveSheet(), command.cellId).content,
            newValue: command.oldValue
        });
        
        this.db.updateCell(this.db.getActiveSheet(), command.cellId, {
            content: command.oldValue,
            formula: command.oldFormula || ''
        });
        
        // Update UI
        const cell = document.querySelector(`[data-cell-id="${command.cellId}"]`);
        if (cell) {
            cell.value = command.oldValue;
        }
    }

    // Redo last undone action
    redo() {
        if (this.redoStack.length === 0) return;
        
        const command = this.redoStack.pop();
        this.undoStack.push({
            cellId: command.cellId,
            oldValue: this.db.getCell(this.db.getActiveSheet(), command.cellId).content,
            newValue: command.newValue
        });
        
        this.db.updateCell(this.db.getActiveSheet(), command.cellId, {
            content: command.newValue,
            formula: command.formula || ''
        });
        
        // Update UI
        const cell = document.querySelector(`[data-cell-id="${command.cellId}"]`);
        if (cell) {
            cell.value = command.newValue;
        }
    }

    // Helper function to get cell ID
    getCellId(cell) {
        return cell.dataset.cellId;
    }
}

// Initialize the spreadsheet when the page loads
window.addEventListener('DOMContentLoaded', () => {
    new Spreadsheet();
}); 