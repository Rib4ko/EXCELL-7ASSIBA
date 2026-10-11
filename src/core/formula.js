import { parseCoordinate, formatCoordinate } from './cell.js';

export class FormulaEngine {
    constructor(sheetGetter) {
        this.getSheet = sheetGetter;
        // DAG graph: cellKey -> Set<cellKey> (dependents: who depends on this cell)
        this.dependents = new Map();
        // Range listeners: rangeKey -> Set<cellKey>
        this.rangeListeners = new Map();
        // Precedents: cellKey -> Set<cellKey or rangeKey>
        this.precedents = new Map();

        // Built-in functions registry
        this.functions = new Map();
        this.registerBuiltins();
    }

    registerBuiltins() {
        this.functions.set('SUM', (args) => {
            const vals = this.flattenNumbers(args);
            return vals.reduce((a, b) => a + b, 0);
        });

        this.functions.set('AVERAGE', (args) => {
            const vals = this.flattenNumbers(args);
            if (vals.length === 0) return '#DIV/0!';
            return vals.reduce((a, b) => a + b, 0) / vals.length;
        });

        this.functions.set('COUNT', (args) => {
            return this.flattenNumbers(args).length;
        });

        this.functions.set('COUNTA', (args) => {
            const vals = this.flattenAll(args);
            return vals.filter(v => v !== '' && v !== null && v !== undefined).length;
        });

        this.functions.set('MIN', (args) => {
            const vals = this.flattenNumbers(args);
            if (vals.length === 0) return 0;
            return Math.min(...vals);
        });

        this.functions.set('MAX', (args) => {
            const vals = this.flattenNumbers(args);
            if (vals.length === 0) return 0;
            return Math.max(...vals);
        });

        this.functions.set('IF', (args) => {
            const condition = args[0];
            const isTrue = condition && condition !== 0 && condition !== '0' && condition !== 'FALSE' && condition !== false;
            return isTrue ? args[1] : (args.length > 2 ? args[2] : false);
        });

        this.functions.set('AND', (args) => {
            const vals = this.flattenAll(args);
            return vals.every(v => Boolean(v));
        });

        this.functions.set('OR', (args) => {
            const vals = this.flattenAll(args);
            return vals.some(v => Boolean(v));
        });

        this.functions.set('NOT', (args) => {
            return !Boolean(args[0]);
        });

        this.functions.set('CONCAT', (args) => {
            const vals = this.flattenAll(args);
            return vals.join('');
        });

        this.functions.set('CONCATENATE', (args) => {
            const vals = this.flattenAll(args);
            return vals.join('');
        });

        this.functions.set('UPPER', (args) => {
            return String(args[0] || '').toUpperCase();
        });

        this.functions.set('LOWER', (args) => {
            return String(args[0] || '').toLowerCase();
        });

        this.functions.set('TRIM', (args) => {
            return String(args[0] || '').trim();
        });

        this.functions.set('ROUND', (args) => {
            const num = parseFloat(args[0]) || 0;
            const decimals = parseInt(args[1], 10) || 0;
            const factor = Math.pow(10, decimals);
            return Math.round(num * factor) / factor;
        });

        this.functions.set('ABS', (args) => {
            return Math.abs(parseFloat(args[0]) || 0);
        });

        this.functions.set('SQRT', (args) => {
            const n = parseFloat(args[0]) || 0;
            if (n < 0) return '#NUM!';
            return Math.sqrt(n);
        });

        this.functions.set('POWER', (args) => {
            return Math.pow(parseFloat(args[0]) || 0, parseFloat(args[1]) || 0);
        });

        this.functions.set('MOD', (args) => {
            const divisor = parseFloat(args[1]) || 0;
            if (divisor === 0) return '#DIV/0!';
            return (parseFloat(args[0]) || 0) % divisor;
        });

        this.functions.set('NOW', () => new Date().toLocaleString());
        this.functions.set('TODAY', () => new Date().toLocaleDateString());

        // --- Phase 2: Advanced Lookup & Reference Functions ---

        // VLOOKUP(lookup_value, table_array, col_index, [exact_match])
        this.functions.set('VLOOKUP', (args) => {
            const lookupVal = args[0];
            const table2D = args[1]; // 2D array or range object with grid
            const colIndex = parseInt(args[2], 10) - 1; // 1-based in Excel
            const exactMatch = args.length > 3 ? (args[3] === false || args[3] === 0 || args[3] === 'FALSE') : true;

            if (!table2D || !Array.isArray(table2D) || colIndex < 0) return '#REF!';

            const rows = Array.isArray(table2D[0]) ? table2D : [table2D];

            for (let r = 0; r < rows.length; r++) {
                const cellVal = rows[r][0];
                const matches = exactMatch
                    ? String(cellVal).trim().toLowerCase() === String(lookupVal).trim().toLowerCase()
                    : cellVal == lookupVal;

                if (matches) {
                    if (colIndex >= rows[r].length) return '#REF!';
                    return rows[r][colIndex];
                }
            }
            return '#N/A';
        });

        // HLOOKUP(lookup_value, table_array, row_index, [exact_match])
        this.functions.set('HLOOKUP', (args) => {
            const lookupVal = args[0];
            const table2D = args[1];
            const rowIndex = parseInt(args[2], 10) - 1;
            const exactMatch = args.length > 3 ? (args[3] === false || args[3] === 0 || args[3] === 'FALSE') : true;

            if (!table2D || !Array.isArray(table2D) || rowIndex < 0) return '#REF!';
            const rows = Array.isArray(table2D[0]) ? table2D : [table2D];
            if (rowIndex >= rows.length) return '#REF!';

            const colsCount = rows[0].length;
            for (let c = 0; c < colsCount; c++) {
                const cellVal = rows[0][c];
                const matches = exactMatch
                    ? String(cellVal).trim().toLowerCase() === String(lookupVal).trim().toLowerCase()
                    : cellVal == lookupVal;

                if (matches) {
                    return rows[rowIndex][c];
                }
            }
            return '#N/A';
        });

        // MATCH(lookup_value, lookup_array, [match_type])
        this.functions.set('MATCH', (args) => {
            const lookupVal = args[0];
            const arr = this.flattenAll([args[1]]);
            const matchType = args.length > 2 ? parseInt(args[2], 10) : 0; // 0: exact

            for (let i = 0; i < arr.length; i++) {
                if (matchType === 0) {
                    if (String(arr[i]).trim().toLowerCase() === String(lookupVal).trim().toLowerCase()) {
                        return i + 1; // 1-based index in Excel
                    }
                } else if (arr[i] == lookupVal) {
                    return i + 1;
                }
            }
            return '#N/A';
        });

        // INDEX(reference, row_num, [col_num])
        this.functions.set('INDEX', (args) => {
            const ref = args[0];
            const rowNum = parseInt(args[1], 10) - 1;
            const colNum = args.length > 2 ? parseInt(args[2], 10) - 1 : 0;

            if (!Array.isArray(ref)) return '#REF!';
            const rows = Array.isArray(ref[0]) ? ref : [ref];

            if (rowNum < 0 || rowNum >= rows.length) return '#REF!';
            if (colNum < 0 || colNum >= rows[rowNum].length) return '#REF!';

            return rows[rowNum][colNum];
        });

        // XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])
        this.functions.set('XLOOKUP', (args) => {
            const lookupVal = args[0];
            const lookupArr = this.flattenAll([args[1]]);
            const returnArr = this.flattenAll([args[2]]);
            const ifNotFound = args.length > 3 ? args[3] : '#N/A';

            for (let i = 0; i < lookupArr.length; i++) {
                if (String(lookupArr[i]).trim().toLowerCase() === String(lookupVal).trim().toLowerCase()) {
                    return i < returnArr.length ? returnArr[i] : '#REF!';
                }
            }
            return ifNotFound;
        });
    }

    flattenAll(args) {
        const result = [];
        for (const item of args) {
            if (Array.isArray(item)) {
                result.push(...this.flattenAll(item));
            } else {
                result.push(item);
            }
        }
        return result;
    }

    flattenNumbers(args) {
        const flat = this.flattenAll(args);
        const nums = [];
        for (const item of flat) {
            if (typeof item === 'number') {
                if (!isNaN(item)) nums.push(item);
            } else if (typeof item === 'string' && item.trim() !== '') {
                const n = parseFloat(item);
                if (!isNaN(n)) nums.push(n);
            }
        }
        return nums;
    }

    tokenize(expr) {
        const tokens = [];
        let i = 0;
        const len = expr.length;

        while (i < len) {
            const char = expr[i];

            if (/\s/.test(char)) {
                i++;
                continue;
            }

            // String literal "..."
            if (char === '"' || char === "'") {
                const quote = char;
                let str = '';
                i++;
                while (i < len && expr[i] !== quote) {
                    str += expr[i];
                    i++;
                }
                i++; // skip closing quote
                tokens.push({ type: 'STRING', value: str });
                continue;
            }

            // Numbers
            if (/[0-9]/.test(char) || (char === '.' && i + 1 < len && /[0-9]/.test(expr[i + 1]))) {
                let numStr = '';
                while (i < len && (/[0-9]/.test(expr[i]) || expr[i] === '.')) {
                    numStr += expr[i];
                    i++;
                }
                tokens.push({ type: 'NUMBER', value: parseFloat(numStr) });
                continue;
            }

            // Cell Reference or Function Name or Boolean
            if (/[A-Za-z_$]/.test(char)) {
                let id = '';
                while (i < len && /[A-Za-z0-9_$:]/.test(expr[i])) {
                    id += expr[i];
                    i++;
                }
                const upper = id.toUpperCase();
                if (upper === 'TRUE') {
                    tokens.push({ type: 'BOOLEAN', value: true });
                } else if (upper === 'FALSE') {
                    tokens.push({ type: 'BOOLEAN', value: false });
                } else if (id.includes(':')) {
                    // Range like A1:B10
                    tokens.push({ type: 'RANGE', value: upper });
                } else if (this.functions.has(upper)) {
                    tokens.push({ type: 'FUNCTION', value: upper });
                } else if (parseCoordinate(upper)) {
                    tokens.push({ type: 'CELL', value: upper });
                } else {
                    tokens.push({ type: 'IDENTIFIER', value: upper });
                }
                continue;
            }

            // Comparison operators (<=, >=, <>, !=, ==, =)
            if (char === '<' || char === '>' || char === '=' || char === '!') {
                let op = char;
                if (i + 1 < len) {
                    const next = expr[i + 1];
                    if ((char === '<' && (next === '=' || next === '>')) ||
                        (char === '>' && next === '=') ||
                        (char === '!' && next === '=') ||
                        (char === '=' && next === '=')) {
                        op += next;
                        i++;
                    }
                }
                tokens.push({ type: 'OPERATOR', value: op });
                i++;
                continue;
            }

            // Arithmetic & structural
            if ('+-*/%^&(),'.includes(char)) {
                tokens.push({ type: 'SYMBOL', value: char });
                i++;
                continue;
            }

            i++;
        }

        return tokens;
    }

    parseTokens(tokens, targetCellKey, recordedPrecedents = new Set()) {
        let pos = 0;

        const peek = () => tokens[pos] || null;
        const consume = () => tokens[pos++];

        const parseExpression = () => parseComparison();

        const parseComparison = () => {
            let left = parseConcat();
            while (peek() && peek().type === 'OPERATOR') {
                const op = consume().value;
                const right = parseConcat();
                const l = left, r = right;
                left = () => {
                    const valL = l();
                    const valR = r();
                    switch (op) {
                        case '=':
                        case '==': return valL == valR;
                        case '<>':
                        case '!=': return valL != valR;
                        case '<': return valL < valR;
                        case '<=': return valL <= valR;
                        case '>': return valL > valR;
                        case '>=': return valL >= valR;
                        default: return false;
                    }
                };
            }
            return left;
        };

        const parseConcat = () => {
            let left = parseAdditive();
            while (peek() && peek().value === '&') {
                consume();
                const right = parseAdditive();
                const l = left, r = right;
                left = () => String(l()) + String(r());
            }
            return left;
        };

        const parseAdditive = () => {
            let left = parseMultiplicative();
            while (peek() && (peek().value === '+' || peek().value === '-')) {
                const op = consume().value;
                const right = parseMultiplicative();
                const l = left, r = right;
                left = () => {
                    const valL = parseFloat(l()) || 0;
                    const valR = parseFloat(r()) || 0;
                    return op === '+' ? valL + valR : valL - valR;
                };
            }
            return left;
        };

        const parseMultiplicative = () => {
            let left = parseExponent();
            while (peek() && (peek().value === '*' || peek().value === '/')) {
                const op = consume().value;
                const right = parseExponent();
                const l = left, r = right;
                left = () => {
                    const valL = parseFloat(l()) || 0;
                    const valR = parseFloat(r()) || 0;
                    if (op === '/') {
                        if (valR === 0) return '#DIV/0!';
                        return valL / valR;
                    }
                    return valL * valR;
                };
            }
            return left;
        };

        const parseExponent = () => {
            let left = parseUnary();
            while (peek() && peek().value === '^') {
                consume();
                const right = parseUnary();
                const l = left, r = right;
                left = () => Math.pow(parseFloat(l()) || 0, parseFloat(r()) || 0);
            }
            return left;
        };

        const parseUnary = () => {
            if (peek() && (peek().value === '+' || peek().value === '-')) {
                const op = consume().value;
                const expr = parseUnary();
                return () => (op === '-' ? -(parseFloat(expr()) || 0) : +(parseFloat(expr()) || 0));
            }
            return parsePrimary();
        };

        const parsePrimary = () => {
            const token = peek();
            if (!token) return () => 0;

            if (token.type === 'NUMBER' || token.type === 'STRING' || token.type === 'BOOLEAN') {
                consume();
                return () => token.value;
            }

            if (token.type === 'CELL') {
                consume();
                const cellCoord = parseCoordinate(token.value);
                const normKey = formatCoordinate(cellCoord.row, cellCoord.col);
                recordedPrecedents.add(normKey);
                return () => {
                    const sheet = this.getSheet();
                    const c = sheet.getCell(cellCoord.row, cellCoord.col);
                    if (!c) return 0;
                    if (typeof c.value === 'number') return c.value;
                    if (c.value === '' || c.value === null) return 0;
                    const num = Number(c.value);
                    return isNaN(num) ? c.value : num;
                };
            }

            if (token.type === 'RANGE') {
                consume();
                const parts = token.value.split(':');
                const startCoord = parseCoordinate(parts[0]);
                const endCoord = parseCoordinate(parts[1]);
                recordedPrecedents.add(token.value);

                return () => {
                    const sheet = this.getSheet();
                    const minRow = Math.min(startCoord.row, endCoord.row);
                    const maxRow = Math.max(startCoord.row, endCoord.row);
                    const minCol = Math.min(startCoord.col, endCoord.col);
                    const maxCol = Math.max(startCoord.col, endCoord.col);

                    const grid2D = [];
                    for (let r = minRow; r <= maxRow; r++) {
                        const rowVals = [];
                        for (let c = minCol; c <= maxCol; c++) {
                            const cell = sheet.getCell(r, c);
                            rowVals.push(cell ? cell.value : 0);
                        }
                        grid2D.push(rowVals);
                    }
                    return grid2D;
                };
            }

            if (token.type === 'FUNCTION') {
                const fnName = consume().value;
                if (!peek() || peek().value !== '(') {
                    return () => '#NAME?';
                }
                consume(); // skip '('
                const args = [];
                while (peek() && peek().value !== ')') {
                    args.push(parseExpression());
                    if (peek() && peek().value === ',') {
                        consume(); // skip ','
                    }
                }
                if (peek() && peek().value === ')') {
                    consume(); // skip ')'
                }

                const fn = this.functions.get(fnName);
                if (!fn) return () => '#NAME?';

                return () => {
                    const evaluatedArgs = args.map(a => a());
                    return fn(evaluatedArgs);
                };
            }

            if (token.value === '(') {
                consume(); // skip '('
                const expr = parseExpression();
                if (peek() && peek().value === ')') {
                    consume(); // skip ')'
                }
                return expr;
            }

            consume();
            return () => 0;
        };

        return parseExpression();
    }

    updateDependencies(cellKey, precedents) {
        const oldPrecedents = this.precedents.get(cellKey);
        if (oldPrecedents) {
            for (const prec of oldPrecedents) {
                if (prec.includes(':')) {
                    const set = this.rangeListeners.get(prec);
                    if (set) set.delete(cellKey);
                } else {
                    const set = this.dependents.get(prec);
                    if (set) set.delete(cellKey);
                }
            }
        }

        this.precedents.set(cellKey, precedents);

        for (const prec of precedents) {
            if (prec.includes(':')) {
                let set = this.rangeListeners.get(prec);
                if (!set) {
                    set = new Set();
                    this.rangeListeners.set(prec, set);
                }
                set.add(cellKey);
            } else {
                let set = this.dependents.get(prec);
                if (!set) {
                    set = new Set();
                    this.dependents.set(prec, set);
                }
                set.add(cellKey);
            }
        }
    }

    isCellInRange(row, col, rangeStr) {
        const parts = rangeStr.split(':');
        const start = parseCoordinate(parts[0]);
        const end = parseCoordinate(parts[1]);
        if (!start || !end) return false;
        const minRow = Math.min(start.row, end.row);
        const maxRow = Math.max(start.row, end.row);
        const minCol = Math.min(start.col, end.col);
        const maxCol = Math.max(start.col, end.col);
        return row >= minRow && row <= maxRow && col >= minCol && col <= maxCol;
    }

    getDependentsOf(cellKey, row, col) {
        const result = new Set();

        const direct = this.dependents.get(cellKey);
        if (direct) {
            for (const d of direct) result.add(d);
        }

        for (const [rangeKey, listeners] of this.rangeListeners.entries()) {
            if (this.isCellInRange(row, col, rangeKey)) {
                for (const listener of listeners) {
                    result.add(listener);
                }
            }
        }

        return result;
    }

    evaluateCell(cell) {
        if (!cell.formula || !cell.formula.startsWith('=')) {
            return;
        }

        const cellKey = formatCoordinate(cell.row, cell.col);
        const expr = cell.formula.substring(1);
        const tokens = this.tokenize(expr);
        const recordedPrecedents = new Set();

        try {
            const evalFn = this.parseTokens(tokens, cellKey, recordedPrecedents);
            this.updateDependencies(cellKey, recordedPrecedents);
            const val = evalFn();
            cell.value = val;
        } catch (e) {
            cell.value = '#ERROR!';
        }
    }

    recalculateFrom(row, col) {
        const startKey = formatCoordinate(row, col);
        const queue = [startKey];
        const visited = new Set();
        const evalOrder = [];

        while (queue.length > 0) {
            const current = queue.shift();
            if (visited.has(current)) continue;
            visited.add(current);

            const coord = parseCoordinate(current);
            if (!coord) continue;

            const deps = this.getDependentsOf(current, coord.row, coord.col);
            for (const dep of deps) {
                if (!visited.has(dep)) {
                    queue.push(dep);
                    evalOrder.push(dep);
                }
            }
        }

        const sheet = this.getSheet();
        for (const cellKey of evalOrder) {
            const coord = parseCoordinate(cellKey);
            if (!coord) continue;
            const cell = sheet.getCell(coord.row, coord.col);
            if (cell && cell.formula) {
                this.evaluateCell(cell);
            }
        }
    }
}
