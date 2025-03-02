class SpreadsheetDB {
    constructor() {
        // Clear any existing data to start fresh
        localStorage.removeItem('spreadsheet_data');
        
        this.data = {
            activeSheet: 'Sheet1',
            sheets: {
                'Sheet1': {
                    cells: {},
                    columnCount: 26,  // A to Z
                    rowCount: 50
                }
            }
        };
        
        this.saveToLocalStorage();
    }

    loadFromLocalStorage() {
        try {
            const data = localStorage.getItem('spreadsheet_data');
            if (!data) return null;
            return JSON.parse(data);
        } catch (error) {
            console.error('Error loading data:', error);
            return null;
        }
    }

    saveToLocalStorage() {
        try {
            localStorage.setItem('spreadsheet_data', JSON.stringify(this.data));
        } catch (error) {
            console.error('Error saving data:', error);
        }
    }

    getActiveSheet() {
        return this.data.activeSheet;
    }

    getAllSheets() {
        return Object.keys(this.data.sheets);
    }

    addSheet(sheetName) {
        // If no sheet name provided, generate one
        if (!sheetName) {
            const sheets = this.getAllSheets();
            let sheetNumber = sheets.length + 1;
            sheetName = `Sheet${sheetNumber}`;
            while (sheets.includes(sheetName)) {
                sheetNumber++;
                sheetName = `Sheet${sheetNumber}`;
            }
        }

        // Check if sheet already exists
        if (this.data.sheets[sheetName]) {
            return false;
        }

        // Create new sheet with default settings
        this.data.sheets[sheetName] = {
            cells: {},
            columnCount: 26,  // A to Z
            rowCount: 50
        };

        // Set as active sheet
        this.data.activeSheet = sheetName;
        this.saveToLocalStorage();
        return true;
    }

    deleteSheet(sheetName) {
        // Don't delete if it's the last sheet
        if (this.getAllSheets().length <= 1) {
            return false;
        }

        // Don't delete if sheet doesn't exist
        if (!this.data.sheets[sheetName]) {
            return false;
        }

        // If deleting active sheet, switch to another sheet
        if (this.data.activeSheet === sheetName) {
            const sheets = this.getAllSheets();
            const newActiveSheet = sheets.find(sheet => sheet !== sheetName);
            this.data.activeSheet = newActiveSheet;
        }

        // Delete the sheet
        delete this.data.sheets[sheetName];
        this.saveToLocalStorage();
        return true;
    }

    switchSheet(sheetName) {
        if (this.data.sheets[sheetName]) {
            this.data.activeSheet = sheetName;
            this.saveToLocalStorage();
            return true;
        }
        return false;
    }

    getCell(sheetName, cellId) {
        const sheet = this.data.sheets[sheetName];
        if (!sheet) return null;

        if (!sheet.cells[cellId]) {
            sheet.cells[cellId] = {
                content: '',
                formula: '',
                style: {
                    bold: false,
                    italic: false,
                    underline: false,
                    alignment: 'left',
                    fontFamily: 'arial',
                    fontSize: 14,
                    color: '#000000',
                    backgroundColor: '#ffffff'
                }
            };
        }
        return sheet.cells[cellId];
    }

    updateCell(sheetName, cellId, updates) {
        const cell = this.getCell(sheetName, cellId);
        if (cell) {
        Object.assign(cell, updates);
            this.saveToLocalStorage();
        }
    }

    updateCellStyle(sheetName, cellId, style) {
        const cell = this.getCell(sheetName, cellId);
        if (cell) {
        Object.assign(cell.style, style);
            this.saveToLocalStorage();
        }
    }

    // Row and Column Management
    getRowCount() {
        return this.data.sheets[this.data.activeSheet].rowCount;
    }

    getColumnCount() {
        return this.data.sheets[this.data.activeSheet].columnCount;
    }

    addRow() {
        const sheet = this.data.sheets[this.data.activeSheet];
        sheet.rowCount++;
        this.saveToLocalStorage();
    }

    deleteRow(rowIndex) {
        const sheet = this.data.sheets[this.data.activeSheet];
        const newCells = {};

        // Move cells up
        Object.entries(sheet.cells).forEach(([cellId, cell]) => {
            const [col, row] = this.parseCellId(cellId);
            if (row > rowIndex) {
                const newCellId = col + (row - 1);
                newCells[newCellId] = cell;
            } else if (row < rowIndex) {
                newCells[cellId] = cell;
            }
        });

        sheet.cells = newCells;
        sheet.rowCount--;
        this.saveToLocalStorage();
    }

    addColumn() {
        const sheet = this.data.sheets[this.data.activeSheet];
        sheet.columnCount++;
        this.saveToLocalStorage();
    }

    deleteColumn(colIndex) {
        const sheet = this.data.sheets[this.data.activeSheet];
        const newCells = {};

        // Move cells left
        Object.entries(sheet.cells).forEach(([cellId, cell]) => {
            const [col, row] = this.parseCellId(cellId);
            const currentColIndex = col.charCodeAt(0) - 65;
            
            if (currentColIndex > colIndex) {
                const newCol = String.fromCharCode(64 + currentColIndex);
                const newCellId = newCol + row;
                newCells[newCellId] = cell;
            } else if (currentColIndex < colIndex) {
                newCells[cellId] = cell;
            }
        });

        sheet.cells = newCells;
        sheet.columnCount--;
        this.saveToLocalStorage();
    }

    // Cell Range Operations
    removeDuplicates(range) {
        const sheet = this.data.sheets[this.data.activeSheet];
        const [startCell, endCell] = range.split(':');
        const [startCol, startRow] = this.parseCellId(startCell);
        const [endCol, endRow] = this.parseCellId(endCell);

        // Store values and their positions for each column
        const columnMaps = new Map();

        // First pass: collect all values and their positions by column
        for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
            const colLetter = String.fromCharCode(col);
            const valueMap = new Map();
            columnMaps.set(colLetter, valueMap);
            
            for (let row = parseInt(startRow); row <= parseInt(endRow); row++) {
                const cellId = colLetter + row;
                const cell = sheet.cells[cellId];
                if (!cell) continue;
                
                const content = cell.content ? cell.content.toString().trim() : '';
                if (content === '') continue;

                if (!valueMap.has(content)) {
                    valueMap.set(content, [row]);
                } else {
                    valueMap.get(content).push(row);
                }
            }
        }

        // Check if any column has duplicates
        let duplicatesFound = false;
        let duplicateOptions = [];
        
        columnMaps.forEach((valueMap, colLetter) => {
            valueMap.forEach((rows, value) => {
                if (rows.length > 1) {
                    duplicatesFound = true;
                    duplicateOptions.push({
                        column: colLetter,
                        value: value,
                        rows: rows,
                        count: rows.length - 1
                    });
                }
            });
        });

        if (!duplicatesFound) {
            alert('No duplicate values found in the selected range.');
            return Promise.resolve(0);
        }

        // Sort duplicate options by column and then by value
        duplicateOptions.sort((a, b) => {
            if (a.column !== b.column) {
                return a.column.localeCompare(b.column);
            }
            return a.value.localeCompare(b.value);
        });

        // Create dialog
        const dialog = document.createElement('div');
        dialog.style.cssText = `
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            z-index: 1000;
            max-height: 80vh;
            overflow-y: auto;
            min-width: 400px;
        `;

        // Create dialog content
        let dialogContent = '<h3 style="margin-top:0">Remove Duplicates</h3>';
        dialogContent += '<div style="margin-bottom:15px">The following duplicates were found:</div>';
        
        // Create select dropdown with improved styling
        const select = document.createElement('select');
        select.style.cssText = `
            width: 100%;
            padding: 8px;
            margin-bottom: 15px;
            border: 1px solid #ccc;
            border-radius: 4px;
            font-size: 14px;
            background-color: #fff;
        `;

        duplicateOptions.forEach((option, index) => {
            const optionElement = document.createElement('option');
            optionElement.value = index;
            optionElement.textContent = `Column ${option.column}: "${option.value}" (${option.count} duplicate${option.count > 1 ? 's' : ''})`;
            select.appendChild(optionElement);
        });

        // Create buttons with improved styling
        const buttonContainer = document.createElement('div');
        buttonContainer.style.cssText = `
            display: flex;
            justify-content: flex-end;
            gap: 10px;
            margin-top: 20px;
        `;

        const cancelButton = document.createElement('button');
        cancelButton.textContent = 'Cancel';
        cancelButton.style.cssText = `
            padding: 8px 16px;
            border: 1px solid #ccc;
            border-radius: 4px;
            cursor: pointer;
            background: #f5f5f5;
            font-size: 14px;
        `;

        const removeButton = document.createElement('button');
        removeButton.textContent = 'Remove Duplicates';
        removeButton.style.cssText = `
            padding: 8px 16px;
            background: #4CAF50;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
        `;

        // Add overlay
        const overlay = document.createElement('div');
        overlay.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0,0,0,0.5);
            z-index: 999;
        `;

        // Add elements to dialog
        dialog.innerHTML = dialogContent;
        dialog.appendChild(select);
        buttonContainer.appendChild(cancelButton);
        buttonContainer.appendChild(removeButton);
        dialog.appendChild(buttonContainer);

        // Add dialog and overlay to document
        document.body.appendChild(overlay);
        document.body.appendChild(dialog);

        // Handle button clicks
        return new Promise((resolve) => {
            cancelButton.onclick = () => {
                document.body.removeChild(dialog);
                document.body.removeChild(overlay);
                resolve(0);
            };

            removeButton.onclick = () => {
                const selected = duplicateOptions[select.value];
                
                // Create confirmation dialog
                const confirmDialog = document.createElement('div');
                confirmDialog.style.cssText = dialog.style.cssText;
                confirmDialog.innerHTML = `
                    <h3 style="margin-top:0">Confirm Removal</h3>
                    <p>Are you sure you want to remove duplicates of "${selected.value}" in column ${selected.column}?</p>
                    <p>This will keep the first occurrence and clear ${selected.count} duplicate(s).</p>
                    <p>Cells to be cleared: ${selected.rows.slice(1).map(row => selected.column + row).join(', ')}</p>
                    <div style="display:flex; justify-content:flex-end; gap:10px; margin-top:20px">
                        <button id="confirmCancel" style="padding:8px 16px; border:1px solid #ccc; border-radius:4px; cursor:pointer; background:#f5f5f5; font-size:14px">Cancel</button>
                        <button id="confirmRemove" style="padding:8px 16px; background:#4CAF50; color:white; border:none; border-radius:4px; cursor:pointer; font-size:14px">Confirm</button>
                    </div>
                `;

                document.body.appendChild(confirmDialog);

                document.getElementById('confirmCancel').onclick = () => {
                    document.body.removeChild(confirmDialog);
                };

                document.getElementById('confirmRemove').onclick = () => {
                    // Remove duplicates
                    let count = 0;
                    for (let i = 1; i < selected.rows.length; i++) {
                        const cellId = selected.column + selected.rows[i];
                        if (sheet.cells[cellId]) {
                            sheet.cells[cellId].content = '';
                            sheet.cells[cellId].formula = '';
                            count++;
                        }
                    }

                    this.saveToLocalStorage();
                    document.body.removeChild(dialog);
                    document.body.removeChild(overlay);
                    document.body.removeChild(confirmDialog);
                    alert(`Removed ${count} duplicate${count !== 1 ? 's' : ''} of value "${selected.value}" in column ${selected.column}`);
                    resolve(count);
                };
            };
        });
    }

    findAndReplace(range, find, replace) {
        const sheet = this.data.sheets[this.data.activeSheet];
        let count = 0;

        // Return early if find text is empty
        if (!find) return 0;

        if (range) {
            const [startCell, endCell] = range.split(':');
            const [startCol, startRow] = this.parseCellId(startCell);
            const [endCol, endRow] = this.parseCellId(endCell);

            for (let row = startRow; row <= endRow; row++) {
                for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
                    const cellId = String.fromCharCode(col) + row;
                    const cell = this.getCell(this.data.activeSheet, cellId);
                    if (cell && cell.content) {
                        const newContent = cell.content.replaceAll(find, replace || '');
                        if (newContent !== cell.content) {
                            cell.content = newContent;
                            count++;
                        }
                    }
                }
            }
        } else {
            // Search entire sheet
            Object.keys(sheet.cells).forEach(cellId => {
                const cell = sheet.cells[cellId];
                if (cell && cell.content) {
                    const newContent = cell.content.replaceAll(find, replace || '');
                    if (newContent !== cell.content) {
                        cell.content = newContent;
                        count++;
                    }
                }
            });
        }

        this.saveToLocalStorage();
        return count;
    }

    // Helper Methods
    parseCellId(cellId) {
        const col = cellId.match(/[A-Z]+/)[0];
        const row = parseInt(cellId.match(/[0-9]+/)[0]);
        return [col, row];
    }

    saveToFile() {
        let fileName = prompt('Enter file name:', this.data.activeSheet);
        if (!fileName) return; // User cancelled

        // Ensure .json extension
        if (!fileName.toLowerCase().endsWith('.json')) {
            fileName += '.json';
        }

        const data = JSON.stringify(this.data, null, 2);
        const filePath = `SavedSheets/${fileName}`;

        // Store the file content in localStorage
        localStorage.setItem(`spreadsheet_file_${fileName}`, data);

        // Update recent files
        const recentFiles = JSON.parse(localStorage.getItem('spreadsheet_recent_files') || '[]');
        const newFile = {
            name: fileName,
            path: filePath,
            lastOpened: new Date().toLocaleString(),
            type: 'saved'
        };
        
        // Remove if exists and add to front
        const updatedFiles = [
            newFile,
            ...recentFiles.filter(f => f.path !== filePath)
        ].slice(0, 10); // Keep only 10 most recent
        
        localStorage.setItem('spreadsheet_recent_files', JSON.stringify(updatedFiles));
        alert('File saved successfully!');
    }

    loadFromFile(file) {
        return new Promise((resolve, reject) => {
            // Check file extension
            if (!file.name.toLowerCase().endsWith('.json')) {
                reject(new Error('Please select a .json file. The file must have a .json extension.'));
                return;
            }

            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    let data;
                    try {
                        data = JSON.parse(e.target.result);
                    } catch (parseError) {
                        throw new Error('Invalid JSON format: The file is not properly formatted. Please ensure it is a valid spreadsheet file.');
                    }

                    // Validate file structure
                    if (!data || typeof data !== 'object') {
                        throw new Error('Invalid file format: File must contain valid spreadsheet data');
                    }
                    if (!data.sheets || typeof data.sheets !== 'object') {
                        throw new Error('Invalid file format: Missing or invalid sheets data');
                    }
                    if (!data.activeSheet || !data.sheets[data.activeSheet]) {
                        throw new Error('Invalid file format: Missing or invalid active sheet');
                    }

                    // Additional validation for sheet structure
                    for (const sheetName in data.sheets) {
                        const sheet = data.sheets[sheetName];
                        if (!sheet.cells || typeof sheet.cells !== 'object') {
                            throw new Error(`Invalid sheet structure in "${sheetName}": Missing or invalid cells data`);
                        }
                        if (typeof sheet.columnCount !== 'number' || typeof sheet.rowCount !== 'number') {
                            throw new Error(`Invalid sheet structure in "${sheetName}": Missing or invalid row/column count`);
                        }
                    }

                    this.data = data;
                    this.saveToLocalStorage();
                    resolve();
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = () => reject(new Error('Error reading file: The file could not be read'));
            reader.readAsText(file);
        });
    }

    exportToCSV() {
        const sheet = this.data.sheets[this.data.activeSheet];
        if (!sheet) return;

        let fileName = prompt('Enter file name:', this.data.activeSheet);
        if (!fileName) return; // User cancelled

        // Ensure .json extension
        if (!fileName.toLowerCase().endsWith('.json')) {
            fileName += '.json';
        }

        // Prepare the data for export
        const exportData = {
            activeSheet: this.data.activeSheet,
            sheets: {
                [this.data.activeSheet]: sheet
            }
        };

        const jsonData = JSON.stringify(exportData, null, 2);
        const filePath = `ExportedSheets/${fileName}`;
        
        // Store the JSON content in localStorage
        localStorage.setItem(`spreadsheet_file_${fileName}`, jsonData);

        // Create and trigger download
        const blob = new Blob([jsonData], { type: 'application/json' });
        const downloadLink = document.createElement('a');
        downloadLink.href = URL.createObjectURL(blob);
        downloadLink.download = fileName;
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
        URL.revokeObjectURL(downloadLink.href);

        // Update recent files
        const recentFiles = JSON.parse(localStorage.getItem('spreadsheet_recent_files') || '[]');
        const newFile = {
            name: fileName,
            path: filePath,
            lastOpened: new Date().toLocaleString(),
            type: 'exported'
        };
        
        // Remove if exists and add to front
        const updatedFiles = [
            newFile,
            ...recentFiles.filter(f => f.path !== filePath)
        ].slice(0, 10); // Keep only 10 most recent
        
        localStorage.setItem('spreadsheet_recent_files', JSON.stringify(updatedFiles));
        alert('File exported successfully!');
    }

    // Helper method to convert column letter to index (A=0, B=1, etc.)
    columnToIndex(column) {
        let index = 0;
        for (let i = 0; i < column.length; i++) {
            index = index * 26 + column.charCodeAt(i) - 'A'.charCodeAt(0);
        }
        return index;
    }

    // Helper method to convert index to column letter (0=A, 1=B, etc.)
    indexToColumn(index) {
        let column = '';
        while (index >= 0) {
            column = String.fromCharCode(65 + (index % 26)) + column;
            index = Math.floor(index / 26) - 1;
        }
        return column;
    }
} 