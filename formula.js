class FormulaHandler {
    constructor() {
        this.operators = {
            '+': (a, b) => a + b,
            '-': (a, b) => a - b,
            '*': (a, b) => a * b,
            '/': (a, b) => a / b
        };

        this.functions = {
            'SUM': this.sum.bind(this),
            'AVERAGE': this.average.bind(this),
            'MAX': this.max.bind(this),
            'MIN': this.min.bind(this),
            'COUNT': this.count.bind(this),
            'TRIM': this.trim.bind(this),
            'UPPER': this.upper.bind(this),
            'LOWER': this.lower.bind(this)
        };
    }

    evaluate(formula, getCellValue) {
        if (!formula.startsWith('=')) return formula;

        try {
            formula = formula.substring(1).trim(); // Remove the '=' sign and trim
            
            // Check for function calls first
            const functionMatch = formula.match(/^([A-Z]+)\((.*)\)$/);
            if (functionMatch) {
                const [_, funcName, args] = functionMatch;
                if (this.functions[funcName]) {
                    return this.functions[funcName](args.trim(), getCellValue);
                }
            }

            return this.evaluateExpression(formula, getCellValue);
        } catch (error) {
            return '#ERROR!';
        }
    }

    // Mathematical Functions
    sum(range, getCellValue) {
        const values = this.getRangeValues(range, getCellValue);
        return values.reduce((sum, val) => sum + (parseFloat(val) || 0), 0);
    }

    average(range, getCellValue) {
        const values = this.getRangeValues(range, getCellValue);
        const numbers = values.filter(val => !isNaN(parseFloat(val)));
        return numbers.length ? this.sum(range, getCellValue) / numbers.length : 0;
    }

    max(range, getCellValue) {
        const values = this.getRangeValues(range, getCellValue);
        const numbers = values.map(val => parseFloat(val)).filter(val => !isNaN(val));
        return numbers.length ? Math.max(...numbers) : 0;
    }

    min(range, getCellValue) {
        const values = this.getRangeValues(range, getCellValue);
        const numbers = values.map(val => parseFloat(val)).filter(val => !isNaN(val));
        return numbers.length ? Math.min(...numbers) : 0;
    }

    count(range, getCellValue) {
        const values = this.getRangeValues(range, getCellValue);
        return values.filter(val => !isNaN(parseFloat(val))).length;
    }

    // Data Quality Functions
    trim(text, getCellValue) {
        let value = this.getSingleValue(text, getCellValue);
        // Handle null, undefined, and non-string values
        if (value == null) return '';
        // Convert to string and trim
        return String(value).replace(/^\s+|\s+$/g, '');
    }

    upper(text, getCellValue) {
        const value = this.getSingleValue(text, getCellValue);
        return String(value).toUpperCase();
    }

    lower(text, getCellValue) {
        const value = this.getSingleValue(text, getCellValue);
        return String(value).toLowerCase();
    }

    // Helper Functions
    getRangeValues(range, getCellValue) {
        // Handle single cell
        if (range.match(/^[A-Z]+[0-9]+$/)) {
            return [getCellValue(range)];
        }

        // Handle range (e.g., A1:B3)
        const [start, end] = range.split(':').map(r => r.trim());
        if (!start || !end) return [range];

        const startCol = this.getColumnIndex(start);
        const endCol = this.getColumnIndex(end);
        const startRow = parseInt(start.match(/[0-9]+/)[0]);
        const endRow = parseInt(end.match(/[0-9]+/)[0]);

        const values = [];
        for (let row = startRow; row <= endRow; row++) {
            for (let col = startCol; col <= endCol; col++) {
                const cellId = String.fromCharCode(65 + col) + row;
                values.push(getCellValue(cellId));
            }
        }
        return values;
    }

    getColumnIndex(cellRef) {
        const col = cellRef.match(/[A-Z]+/)[0];
        let index = 0;
        for (let i = 0; i < col.length; i++) {
            index = index * 26 + col.charCodeAt(i) - 65;
        }
        return index;
    }

    getSingleValue(text, getCellValue) {
        text = text.trim();
        if (text.match(/^[A-Z]+[0-9]+$/)) {
            return getCellValue(text);
        }
        return text;
    }

    evaluateExpression(expression, getCellValue) {
        // Replace cell references with their values
        expression = expression.replace(/[A-Z]+[0-9]+/g, (match) => {
            const value = getCellValue(match.trim());
            return isNaN(value) ? 0 : parseFloat(value);
        });

        // Handle basic arithmetic
        return this.calculateExpression(expression);
    }

    calculateExpression(expression) {
        // Remove whitespace
        expression = expression.replace(/\s+/g, '');

        // Handle parentheses first
        while (expression.includes('(')) {
            expression = expression.replace(/\(([^()]+)\)/g, (_, subExpr) => 
                this.calculateSimpleExpression(subExpr));
        }

        return this.calculateSimpleExpression(expression);
    }

    calculateSimpleExpression(expression) {
        // Handle multiplication and division first
        while (expression.match(/[\d.]+[*/][\d.]+/)) {
            expression = expression.replace(/([\d.]+)([*/])([\d.]+)/, (_, a, op, b) => 
                this.operators[op](parseFloat(a), parseFloat(b)));
        }

        // Then handle addition and subtraction
        while (expression.match(/[\d.]+[+-][\d.]+/)) {
            expression = expression.replace(/([\d.]+)([+-])([\d.]+)/, (_, a, op, b) => 
                this.operators[op](parseFloat(a), parseFloat(b)));
        }

        return parseFloat(expression);
    }
} 